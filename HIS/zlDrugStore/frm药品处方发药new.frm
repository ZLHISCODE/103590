VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Begin VB.Form frm药品处方发药New 
   Caption         =   "药品处方发药"
   ClientHeight    =   5310
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   7635
   Icon            =   "frm药品处方发药new.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5310
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer tmrMsgRefresh 
      Interval        =   60000
      Left            =   4200
      Top             =   3840
   End
   Begin VB.Timer tmrCall 
      Interval        =   5000
      Left            =   6240
      Top             =   2880
   End
   Begin VB.Timer TimePrintCancelBill 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   5520
      Top             =   2880
   End
   Begin VB.Timer TimeRefresh 
      Enabled         =   0   'False
      Left            =   4920
      Top             =   2880
   End
   Begin VB.Timer TimePrint 
      Enabled         =   0   'False
      Left            =   4320
      Top             =   2880
   End
   Begin VB.PictureBox picCondition 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   120
      ScaleHeight     =   4455
      ScaleWidth      =   3615
      TabIndex        =   3
      Top             =   240
      Width           =   3615
      Begin VB.PictureBox picList 
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   120
         ScaleHeight     =   1335
         ScaleWidth      =   2535
         TabIndex        =   15
         Top             =   2760
         Width           =   2535
         Begin XtremeSuiteControls.TabControl tbcList 
            Height          =   975
            Left            =   120
            TabIndex        =   16
            Top             =   120
            Width           =   1455
            _Version        =   589884
            _ExtentX        =   2566
            _ExtentY        =   1720
            _StockProps     =   64
            Enabled         =   -1  'True
         End
      End
      Begin VB.PictureBox picConMain 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Height          =   2175
         Left            =   0
         ScaleHeight     =   2175
         ScaleWidth      =   3495
         TabIndex        =   4
         Top             =   120
         Width           =   3495
         Begin VB.TextBox txtPati 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   960
            TabIndex        =   21
            Top             =   1080
            Width           =   1245
         End
         Begin VB.CheckBox chk显示已确认单据 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            Caption         =   "显示已确认单据"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   1920
            Width           =   1695
         End
         Begin VB.CheckBox Chk显示退药待发单据 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            Caption         =   "显示退药待发单据"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   1920
            Value           =   1  'Checked
            Width           =   3015
         End
         Begin VB.CommandButton cmdFind 
            Height          =   300
            Left            =   2880
            Picture         =   "frm药品处方发药new.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "处方定位(F2)"
            Top             =   1080
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.ComboBox cbo病区 
            Enabled         =   0   'False
            Height          =   276
            Left            =   960
            TabIndex        =   7
            Top             =   1440
            Width           =   2415
         End
         Begin VB.ComboBox cbo时间范围 
            Height          =   300
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   40
            Width           =   2415
         End
         Begin VB.CommandButton cmdIC 
            Caption         =   "读卡"
            Height          =   300
            Left            =   2760
            TabIndex        =   5
            Top             =   1080
            Visible         =   0   'False
            Width           =   615
         End
         Begin MSComCtl2.DTPicker Dtp结束时间 
            Height          =   315
            Left            =   960
            TabIndex        =   8
            Top             =   720
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
            Format          =   123207683
            CurrentDate     =   36985
         End
         Begin MSComCtl2.DTPicker Dtp开始时间 
            Height          =   315
            Left            =   960
            TabIndex        =   9
            Top             =   375
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
            Format          =   123207683
            CurrentDate     =   36985
         End
         Begin VB.CheckBox Chk显示过程单据 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            Caption         =   "显示所有过程单据"
            Height          =   195
            Left            =   120
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   1920
            Width           =   3000
         End
         Begin zlIDKind.IDKindNew IDKNType 
            Height          =   375
            Left            =   120
            TabIndex        =   22
            Top             =   1028
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   661
            ShowSortName    =   0   'False
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
            BackColor       =   12632319
         End
         Begin VB.Image imgFilter 
            Height          =   240
            Left            =   2400
            Picture         =   "frm药品处方发药new.frx":0454
            Top             =   1080
            Width           =   240
         End
         Begin VB.Image img病区 
            Height          =   240
            Left            =   600
            Picture         =   "frm药品处方发药new.frx":6CA6
            ToolTipText     =   "选择病区"
            Top             =   1470
            Width           =   240
         End
         Begin VB.Label lbl病区 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0FF&
            Caption         =   "病区"
            Height          =   180
            Left            =   120
            TabIndex        =   14
            Top             =   1500
            Width           =   360
         End
         Begin VB.Label lblTimeEnd 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0FF&
            Caption         =   "结束时间"
            Height          =   180
            Left            =   120
            TabIndex        =   13
            Top             =   787
            Width           =   720
         End
         Begin VB.Label lblTimeBegin 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0FF&
            Caption         =   "开始时间"
            Height          =   180
            Left            =   120
            TabIndex        =   12
            Top             =   469
            Width           =   720
         End
         Begin VB.Label lbl时间范围 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0FF&
            Caption         =   "时间范围"
            Height          =   180
            Left            =   120
            TabIndex        =   11
            Top             =   110
            Width           =   720
         End
      End
   End
   Begin VB.PictureBox picDetail 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   4200
      ScaleHeight     =   1575
      ScaleWidth      =   3015
      TabIndex        =   0
      Top             =   840
      Width           =   3015
      Begin VB.Frame fraLine 
         Height          =   2085
         Left            =   120
         TabIndex        =   1
         Top             =   -120
         Width           =   45
      End
      Begin XtremeSuiteControls.TabControl tbcDetail 
         Height          =   975
         Left            =   360
         TabIndex        =   2
         Top             =   120
         Width           =   1455
         _Version        =   589884
         _ExtentX        =   2566
         _ExtentY        =   1720
         _StockProps     =   64
         Enabled         =   -1  'True
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   19
      Top             =   4950
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm药品处方发药new.frx":6DF0
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8387
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   4920
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frm药品处方发药new.frx":7684
      Left            =   5520
      Top             =   240
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frm药品处方发药New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'''''''''''''''''''其它定义
Private mlngMode As Long
Private mstrPrivs As String                              '权限串

Private gstrProductName As String
Private mint字号 As Integer
Private mlngIC病人id As Long                           '通过IC卡获取病人id

Private Const cstLocate As Integer = 0
Private Const cstFilter As Integer = 1

Private mfrmList As New frm处方发药列表
Private mfrmDetail As New frm处方发药明细
Private mfrmRecipe As New frm处方

'消费卡
Private mstrCardType As String   '消费卡/银行卡类别，格式：短名|全名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密);…
Private mintCardCount As Integer  '卡数量
Private mint就诊卡长度 As Integer
Private mobjcard As Card

Private mobjSquareCard As Object             '一卡通接口
Private mobjPlugIn As Object             '外挂接口对象
Private mobjCISJOB As Object  '电子病案查阅对象

Private mstrStockName As String

Private mint过滤查询 As Integer                             '是否通过滤过界面进行查询

Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1

'消息相关对象变量
Private WithEvents mobjMipModule As zl9ComLib.clsMipModule
Attribute mobjMipModule.VB_VarHelpID = -1
Private mdteMsgRefresh As Date              '上次刷新时间
Private mblnExistMsg As Boolean             '在一定时间段内是否收到消息

Private mblnCard As Boolean                             '是否刷就诊卡
Private mblnScaner As Boolean                           '是否扫描器输入
Private mstrScanerLastNo As String                      '扫描器输入的上次NO
Private mblnScaned As Boolean                           '已经扫描过一次
Private mblnFinding As Boolean                          '查找定位模式时是否找到数据
Private mintOld输入模式 As Integer                      '扫描身份证之前的类别
Private mblnBrushCard As Boolean                        '是否刷卡
Private mstrLastBrushCardNo As String                   '上次刷卡NO

Private mblnStart As Boolean
Private mblnInput As Boolean

Private mdate上次校验时间 As Date
Private mstr自动配药人 As String                        '用于自动配药功能中
Private mbln未取药发药 As Boolean
Private mstr窗口 As String

Private mint发药方式 As Integer                         '区分用户是批量发药模式还是单个发药模式。用于住院单据的销帐检查。

Private mstrChargePrivs As String                        '门诊划价权限串
Private mstrStuffPrivs As String                         '卫材发放管理权限串

Public RecPart As New ADODB.Recordset                   '药房
Private mrsDrugStock As ADODB.Recordset                 '存储库房
Private mrsIsDosage As ADODB.Recordset                  '配药控制
Private mrsApplyforcredit As Recordset                  '用于记录存在销帐申请的单据

Private mblnLoadDrug As Boolean
Private mblnPackerConnect As Boolean    '发药机是否已经链接
Private mstrOpr As String               '发药窗口
Private mintAutoSendFlow As Integer     '发药流程控制：0-仅有开始发药流程，1-有开始发药，结束发药流程
Private Enum mSendOper                  '发药操作流程：0-开始发药,1-结束发药
    StartSend = 0
    EndSend = 1
End Enum
Private mblnCompatible As Boolean       '兼容性检查：true-兼容最新接口,false-不兼容

Public BlnSetParaSuccess As Boolean                     '设置成功与否
Private BlnRefresh As Boolean
Private IntTimes As Integer                             '已延迟
Private BlnInRefresh As Boolean                         '是否处于刷新状态
Private mblnIsFirst As Boolean                          '未校验

Private mstrDeptNode As String          '当前药房的站点

Private Type Type_Queue
    blnCallOver As Boolean             '当前呼叫是否已完成
    strPCName As String                '本机机器名
    strSendWin As String                '当前药房的发药窗口
    blnRemoteCall As Boolean             '本机是远程呼叫机器
    blnWin As Boolean
End Type
Private mQueue As Type_Queue  '排队叫号使用的一些变量

Private mbln允许两次刷卡 As Boolean

Private mblnStateTimeRefresh As Boolean
Private mblnStateTimePrint As Boolean
Private mblnStateTimeCall As Boolean

Private mrsList As ADODB.Recordset
Private mrsDetail As ADODB.Recordset

Private mstr操作员 As String
Private mstr配药人 As String

Private mstr毒麻类提示 As String
Private mstr价格失效提示 As String

Private mstrPrintRecipe As String                       '用于发药后打印，记录单据号、单据类型：单据号1,单据类型1,记录性质1,门诊标志1,处方类型1,|单据号2,单据类型2......

Private mbln就诊卡 As Boolean                           '是否自动定位到就诊卡

Private str单位串 As String                             '单位串

Private mstr序号 As String

Private mstrBill As String                               '记录上张已发药处方号及单据类型

Private mintUnit As Integer                             '单位系数：1-售价;2-门诊;3-住院;4-药库

Public int模式  As Integer

Private mintTab As Integer

'从参数表中取药品价格、数量、金额小数位数
'Private mintCostDigit As Integer            '成本价小数位数
'Private mintPriceDigit As Integer           '售价小数位数
'Private mintNumberDigit As Integer          '数量小数位数
Private mintMoneyDigit As Integer           '金额小数位数

Private Const mconint售价单位 As Integer = 1
Private Const mconint门诊单位 As Integer = 2
Private Const mconint住院单位 As Integer = 3
Private Const mconint药库单位 As Integer = 4

Private mstrOracleMoneyForamt As String                 'ORACLE中金额格式
Private mstrVBMoneyForamt As String                     'VB中金额格式

Private mstrRPTDefaultScheme_Recipt As String           '处方签报表的默认格式
Private mstrRPTScheme_配药单 As String
Private mstrRPTScheme_其他格式 As String

Private mblnSendIsOver As Boolean         '发药是否结束：false-未结束，true-结束

'默认的窗体大小
Private Const mcstlngWinNormalWidth As Long = 13500
Private Const mcstlngWinNormalHeight As Long = 9000

Private mstr病人类型 As String  '格式：病人类型,颜色|病人类型,颜色...
Private mclsComLib As Object
Private mobjDrugMAC As Object

'列表类型
Private Enum mListType
    配药确认 = 0
    待配药 = 1
    已配药 = 2
    待发药 = 3
    超时未发 = 4
    退药 = 5
End Enum

'时间范围
Private Enum mTimeRange
    当天 = 0
    
    两天内 = 1
    三天内 = 2
    指定时间范围 = 3
End Enum

Private Enum mFindType
    单据号 = 1
    门诊号 = 2
    姓名 = 3
    身份证 = 4
    IC卡 = 5
    医保号 = 6
    住院号 = 7
End Enum

'权限
Private Type Type_Privs
    bln所有药房 As Boolean
    bln发药 As Boolean
    bln退药 As Boolean
    bln退其它药房的处方 As Boolean
    bln发已结帐处方 As Boolean
    bln退已结帐处方 As Boolean
    bln发退出院病人处方 As Boolean
    bln校验处方 As Boolean
    bln医生查询 As Boolean
    bln合理用药监测 As Boolean
    bln过滤附加条件 As Boolean
    bln修改过滤日期 As Boolean
    bln参数设置 As Boolean
    bln发其它药房的处方 As Boolean
    bln发病区处方 As Boolean
    bln配药 As Boolean
    bln停止发药 As Boolean
    bln恢复发药 As Boolean
    bln取药确认 As Boolean
    bln药品自动化接口 As Boolean
    bln电子病案查阅 As Boolean
    bln允许查询所有时间范围单据 As Boolean
End Type
Private mPrives As Type_Privs

'使用到的参数（来自系统参数表或其它参数表或本机注册表）
Private Type Type_Params
    '参数表中的系统参数
    bln允许未审核处方发药 As Boolean
    bln允许未收费处方发药 As Boolean
    bln医嘱作废 As Boolean
    int金额保留位数 As Integer
    bln审核划价单 As Boolean
    bln刷卡验证 As Boolean
    bln报警包含划价费用 As Boolean
    int药品名称显示 As Integer          '0-编码和名称，1-仅编码，2-仅名称
    bln发药前收费或审核 As Boolean
    bln启用审方     As Boolean          '是否启用处方审查系统

    '参数表中的其它参数
    intShowBill收费 As Integer
    intShowBill记帐 As Integer
    intShowBill配药 As Integer          '0-显示所有配药单,1-只显示未打印的待配药单据,2-只显示已打印的待配药单据
    bln记帐单 As Boolean
    lngPrintBackInterval As Long
    lngPrintDelay As Long
    int显示病区处方 As Integer
    lngRefreshInterval As Long
    lngPrintInterval As Long
    int校验发药人 As Integer
    int校验配药人 As Integer
    int自动销帐 As Integer
    bln显示大小单位 As Boolean
    IntShowCol As Integer
    IntAutoPrint As Integer
    intPrint As Integer
    intPrintDrugLable As Integer
    int打印卫材发料清单 As Integer
    lng药房ID As Long
    Str窗口 As String
    str配药人 As String
    strPrintWindow As String
    bln自动配药 As Boolean
    int自动配药时限 As Integer
    strSourceDep As String
    int配药后自动打印 As Integer
    int查询未发药单据天数 As Integer                '根据参数设置的查询天使，在发药时查询当前病人[当前药房其他窗口]或[其他药房]的未发药单据
    int发药后自动打印药品标签 As Integer
    bln发药后刷卡验证 As Boolean
    bln配药扫描 As Boolean
    bln配药收费 As Boolean
    bln打印所有格式 As Boolean
    blnPreview As Boolean
    
    intOverTime As Integer
    intType As Integer              '0-显示门诊和住院处方；1-显示门诊处方；2-显示住院处方
    str两次刷卡发药 As String       '两次刷卡发药的卡类别ID
    bln发生时间过滤 As Boolean      '药品医嘱按发生时间(首次时间)过滤：0-按产生单据时间过滤；1-按发生时间过滤
    int金额显示 As Integer          '金额显示方式：0-显示应收金额,1-显示实收金额,2-显示应收和实收金额
    bln取药确认 As Boolean          '是否启用病人实际取药确认模式：0-不启用，1-启用
    bln发药后检查 As Boolean        '是否检查该病人在当前药房是否有未发的材料单据
    bln扫描后呼叫 As Boolean        '0-不自动呼叫,1-扫描后自动语音呼叫'
    int回车方式 As Integer          '通过录入或刷卡查找时系统自动添加回车处理的方式，0-系统不自动回车,1-当录入达到项目或卡号长度时自动回车
    
    '排队叫号涉及参数
    blnStartQueue As Boolean        '启用排队叫号
    intSoundType As Integer         '语音类型：0-系统语音；1-微软语音
    blnShowQueue As Boolean         '显示排队队列
    blnStartCall As Boolean         '启用语音呼叫
    intCallType As Integer          '叫号方式：0-本地叫号；1-远端叫号
    strRemoteCall As String         '远端呼叫站点
    intSoundSpeed As Integer        '语音广播语速
    intSoundTimes As Integer        '语音播放次数
    lngShowComponent As Long        '显示设备类别
    intCircleTime As Integer        '呼叫轮询时间
    blnSign As Boolean              '签到和配药一起完成
    
    
    '注册表参数
    int界面定位 As Integer
    int待发单据 As Integer
    int过程单据 As Integer
    int输入模式索引 As Integer
    int已确认单据 As Integer
    strDefaultPrinter As String   '西药处方签默认的打印机
    
    int输入模式 As Integer
    
    '药房是否需要配药
    blnMustDosageProcess As Boolean
    
    '药房是否需要配药确认
    blnMustDosageOkProcess As Boolean
    
    '库存检查
    IntCheckStock As Integer
    
    '用户定义的处方颜色，从注册表取的字符串，用;分隔
    strUserRecipeColor As String
    
    '各种颜色的西药处方签对应的打印机列表，用;分隔
    strPrinters As String
    
    '配药单和处方签指定的打印格式
    str配药格式 As String
    str处方格式 As String
    
    '启用合理用药PASS
    blnStarPass As Boolean
    
    '库房单位
    strUnit As String
    
    intShowName As Integer         '药品名称显示方式
    intFont As Integer             '字体号
    
    blnDispensing As Boolean        '呼叫是否同时通知接口准备发药
End Type
Private mParams As Type_Params

Private Type Type_Condition
    intListType As Integer
    bln发病区处方 As Boolean
    int服务对象 As Integer                  '药房的服务对象：1-门诊病人;2-住院病人;3-门诊和住院
    int离院带药 As Integer
    bln显示过程单据 As Boolean
    bln显示退药待发单据 As Boolean
    bln显示已确认单据 As Boolean
End Type
Private mcondition As Type_Condition

Private Type Type_mSQLCondition
    lng药房ID As Long
    date开始日期 As Date
    date结束日期 As Date
    str开始NO As String
    str结束NO As String
    str姓名 As String
    str就诊卡 As String
    str标识号 As String
    lng科室ID As Long
    str填制人 As String
    str审核人 As String
    lng药品id As Long
    str当前NO As String
    str门诊号 As String
    str身份证 As String
    str医保号 As String
    lng住院号 As Long
    intOverTime As Integer
    lng病人ID As Long
End Type
Private mSQLCondition As Type_mSQLCondition
Private Function CardConfirm(ByVal rsData As ADODB.Recordset) As Boolean
    '消费卡消费确认接口
    '如果是批量发药，并且包含多个病人，按病人多次调用刷卡消费接口
    '实际在之前已进行校验，如果包含多个病人需要刷卡消费，则禁止发药，所以这里应该不包含多个病人刷卡消费
    '暂时保留这种处理方式，可能以后会变动
    Dim lngCard病人ID As Long
    Dim strCardNo As String
        
    On Error GoTo ErrHand
    
    If mParams.bln发药前收费或审核 = False Then
        CardConfirm = True
        Exit Function
    End If
    
    If mobjSquareCard Is Nothing Then
        MsgBox "一卡通部件故障，不能进行刷卡消费。", vbInformation, gstrSysName
        Exit Function
    End If
    
    '注意传入的记录集是处方明细
    '收费单据
    rsData.Filter = "标志=1 And 记录性质=1 And 已收费=0 And 病人ID>0 "
    rsData.Sort = "病人ID,NO"
     
    Do While Not rsData.EOF
        If lngCard病人ID <> rsData!病人ID Then
            If strCardNo <> "" Then
                '刷卡消费
                If zlfuncCard_Confirm(mobjSquareCard, Me, mlngMode, mstrPrivs, lngCard病人ID, mobjcard.接口序号, 1, strCardNo) = False Then
                    Exit Function
                End If
            End If
             
            lngCard病人ID = rsData!病人ID
            strCardNo = rsData!NO
        Else
            If strCardNo = "" Then
                strCardNo = rsData!NO
            ElseIf InStr(1, strCardNo, rsData!NO) = 0 Then
                strCardNo = strCardNo & "," & rsData!NO
            End If
        End If
        rsData.MoveNext
     Loop
     
     If strCardNo <> "" Then
         '刷卡消费
         If zlfuncCard_Confirm(mobjSquareCard, Me, mlngMode, mstrPrivs, lngCard病人ID, mobjcard.接口序号, 1, strCardNo) = False Then
             Exit Function
         End If
     End If
    
    lngCard病人ID = 0
    strCardNo = ""
    
    '记账单据：只对门诊病人进行处理
    rsData.Filter = "标志=1 And 记录性质=2 And 已收费=0 And 病人ID>0 "
    rsData.Sort = "病人ID,NO"
    Do While Not rsData.EOF
        If rsData!门诊标志 = 1 Or rsData!门诊标志 = 4 Then
            If lngCard病人ID <> rsData!病人ID Then
                If strCardNo <> "" Then
                    '刷卡消费
                    If zlfuncCard_Confirm(mobjSquareCard, Me, mlngMode, mstrPrivs, lngCard病人ID, mobjcard.接口序号, 2, strCardNo) = False Then
                        Exit Function
                    End If
                    strCardNo = ""
                End If
                
                lngCard病人ID = rsData!病人ID
                strCardNo = rsData!NO
            Else
                If strCardNo = "" Then
                    strCardNo = rsData!NO
                ElseIf InStr(1, strCardNo, rsData!NO) = 0 Then
                    strCardNo = strCardNo & "," & rsData!NO
                End If
            End If
        End If
        rsData.MoveNext
    Loop
    If strCardNo <> "" Then
        '刷卡消费
        If zlfuncCard_Confirm(mobjSquareCard, Me, mlngMode, mstrPrivs, lngCard病人ID, mobjcard.接口序号, 2, strCardNo) = False Then
            Exit Function
        End If
    End If
    
    CardConfirm = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    CardConfirm = False
End Function



Private Function CheckPati(ByVal rsData As ADODB.Recordset) As Boolean
    '如果是刷消费卡确认模式，则不能多病人批量发药
    Dim lng收费病人ID As Long
    Dim lng记账病人ID As Long
    Dim blnSend As Boolean
    Const cstMsg As String = "不能多个病人同时进行刷卡消费确认，请确保所选单据是同一个病人！"
    
    If mParams.bln发药前收费或审核 = False Then
        CheckPati = True
        Exit Function
    End If
    
    blnSend = True
    
    '刷卡模式，检查病人是否有病人信息记录
    rsData.Filter = "已收费=0 And 病人ID=0"
    If Not rsData.EOF Then
        blnSend = False
        CheckPati = False
        MsgBox "未收费的划价单不允许发药，请先收费！", vbInformation, gstrSysName
        Exit Function
    End If
        
    '检查收费单是否存在不同的病人
    rsData.Filter = "标志=1 And 记录性质=1 And 已收费=0 And 病人ID>0"
    rsData.Sort = "病人ID,NO"
    Do While Not rsData.EOF
        If lng收费病人ID = 0 Then
            lng收费病人ID = rsData!病人ID
        ElseIf lng收费病人ID <> rsData!病人ID Then
            blnSend = False
            Exit Do
        End If
        rsData.MoveNext
    Loop
    
    If blnSend = False Then
        MsgBox cstMsg, vbInformation, gstrSysName
        Exit Function
    End If
    
    blnSend = True
    
    '检查记账单是否存在不同的病人
    rsData.Filter = "标志=1 And 记录性质=2 And 已收费=0 And 病人ID>0"
    rsData.Sort = "病人ID,NO"
    Do While Not rsData.EOF
        If rsData!门诊标志 = 1 Or rsData!门诊标志 = 4 Then
            If lng记账病人ID = 0 Then
                lng记账病人ID = rsData!病人ID
            ElseIf lng记账病人ID <> rsData!病人ID Then
                blnSend = False
                Exit Do
            End If
        End If
        rsData.MoveNext
    Loop
        
    If blnSend = False Then
        MsgBox cstMsg, vbInformation, gstrSysName
        Exit Function
    End If
    
    '检查收费单和记账单是否同一个病人
    If lng收费病人ID <> 0 And lng记账病人ID <> 0 And lng收费病人ID <> lng记账病人ID Then
        MsgBox cstMsg, vbInformation, gstrSysName
        Exit Function
    End If
    
    CheckPati = True
End Function

Private Sub CloseQueue()
    '关闭LCD窗口
    If Not gobjLEDShow Is Nothing Then
        Call gobjLEDShow.zlDrugShowClose
        Set gobjLEDShow = Nothing
    End If
End Sub

Private Sub GetPatiType(ByVal rsList As ADODB.Recordset)
    '汇总当前操作中的病人类型及对应颜色，并在状态栏显示
    
    If rsList Is Nothing Then Exit Sub
    If rsList.RecordCount = 0 Then Exit Sub
    
    With rsList
        .MoveFirst
        
        Do While Not .EOF
            If InStr(1, "|" & mstr病人类型, "|" & !病人类型 & ",") = 0 Then
                mstr病人类型 = IIf(mstr病人类型 = "", "", mstr病人类型 & "|") & !病人类型 & "," & zldatabase.GetPatiColor(IIf(IsNull(!病人类型), "", !病人类型))
            End If
            .MoveNext
        Loop
    End With
    
    
End Sub

Private Sub GetSendWindows(ByVal lng药房ID As Long)
    '取当前药房的发药窗口
    Dim rstemp As ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = "select 编码,名称,上班否 from 发药窗口 where 药房id=[1]"
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "GetSendWindows", lng药房ID)
    
    mQueue.blnWin = False
    mQueue.strSendWin = ""
    Do While Not rstemp.EOF
        mQueue.strSendWin = IIf(mQueue.strSendWin = "", "", mQueue.strSendWin & ",") & rstemp!名称
        If rstemp!上班否 = 1 Then mQueue.blnWin = True
        rstemp.MoveNext
    Loop
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function IsValidMsg(strMsgCode As String, strMsgXml As String) As Boolean
    '根据业务参数设置判断是否是有效消息
    Dim objXML As New zl9ComLib.clsXML
    Dim strCodeNode As String
    Dim str药房id As String
    Dim int单据性质 As Integer
    Dim int收费状态 As Integer
    Dim str发药窗口 As String
    Dim blnValid As Boolean
    Dim rsMsg As New ADODB.Recordset
    Dim lngParentID As Long
    
'    'ZLHIS_CHARGE_003
'    patient_info 病人信息
'    patient_id 病人id
'    patient_name 姓名
'    identity_card 身份证号
'    in_number 住院号
'    out_number 门诊号
'    charge_bill
'       bill_no 单据号码
'       bill_kind 单据性质 1-收费单;2-记帐单
'       drug_window 发药窗口
'       charge_state 收费状态 1-未收费;2-已收费
'       charge_time 收费时间
'       charge_person 收费人员
'    bill_item
'       charge_item_id 收费项目id
'       charge_item_kind 收费类别
'       execute_dept_id 执行部门id
'       execute_dept_title 执行部门
    
'    'ZLHIS_CIS_006
'    patient_info 病人信息
'    patient_id 病人id
'    patient_name 姓名
'    in_number 住院号
'    out_number 门诊号
'    patient_clinic 就诊信息
'    patient_source 病人来源
'    clinic_id 就诊id
'    clinic_dept_id 就诊科室id
'    clinic_dept_title 科室名称
'    clinic_room 就诊病房
'    clinic_bed 就诊病床
'    charge_bill
'        send_serial 发送批号
'        send_time 发送时间
'        send_person 发送人员
'        bill_no 单据号码
'        bill_kind 单据性质 1-收费单;2-记帐单
'        charge_state 收费状态 1-未收费;2-已收费
'        send_order 发送医嘱
'        order_id 医嘱id
'        order_relevant_id 相关ID
'        order_info 医嘱内容
'        order_rate 执行频率
'        order_route_id 给药途径id
'        order_route 给药途径
'        order_starttime 开始时间
'        order_single 单量
'        order_total 总量
'        order_entrust 医嘱嘱托
'        order_item_id 品种id
'        charge_item_kind 药品类别
'        charge_item_id 药品id
'        execute_dept_id 执行部门id
    
    On Error GoTo ErrHand
    
    If objXML Is Nothing Then Exit Function

    '打开XML文件
    objXML.OpenXMLDocument strMsgXml
    
    '从打开的XML文件中取指定节点的值和当前客户机参数设置比较，在消息可能包含多个NO的情况下只要有一个NO满足条件就表示有效
    '1.判断药房id
    If strMsgCode = "ZLHIS_CHARGE_003" Then
        strCodeNode = "bill_item"
    ElseIf strMsgCode = "ZLHIS_CIS_006" Then
        strCodeNode = "charge_bill"
    End If
    If objXML.GetMultiNodeRecord(strCodeNode, rsMsg) = False Then Exit Function
    If rsMsg Is Nothing Then Exit Function
    If rsMsg.RecordCount = 0 Then Exit Function
    
    blnValid = False
    Do While Not rsMsg.EOF
        If rsMsg("node_name").Value = "execute_dept_id" Then
            If Val(rsMsg("node_value").Value) = mSQLCondition.lng药房ID Then
                blnValid = True
                Exit Do
            End If
        End If
        rsMsg.MoveNext
    Loop
    If blnValid = False Then Exit Function
     
     
    '    Select Case mParams.intShowBill收费
'        Case 0  '不显示处方
'        Case 1  '显示未收费
'        Case 2  '显示已收费
'        Case 3  '显示所有处方
'    End Select
        
'        Select Case mParams.intShowBill记帐
'        Case 0  '不显示处方
'        Case 1  '显示未审核
'        Case 2  '显示已审核
'        Case 3  '显示所有处方

    '2.判断单据性质和收费/审核状态
    If mParams.intShowBill收费 = 0 And mParams.intShowBill记帐 = 0 Then Exit Function
    If mParams.intShowBill收费 <> 3 Or mParams.intShowBill记帐 <> 3 Then
        If objXML.GetMultiNodeRecord("charge_bill", rsMsg) = False Then Exit Function
        If rsMsg Is Nothing Then Exit Function
        If rsMsg.RecordCount = 0 Then Exit Function
                
        blnValid = False
        Do While Not rsMsg.EOF
            If zlStr.NVL(rsMsg("parent_id").Value) <> "" Then
                If lngParentID = 0 Then
                    lngParentID = Val(rsMsg("parent_id").Value)
                ElseIf lngParentID <> Val(rsMsg("parent_id").Value) Then
                    '只要有一个NO满足单据性质和收费/审核状态就表示有效就退出循环
                    If (mParams.intShowBill收费 = 1 And int单据性质 = 1 And int收费状态 = 1) Or _
                        (mParams.intShowBill收费 = 2 And int单据性质 = 1 And int收费状态 = 2) Or _
                        (mParams.intShowBill记帐 = 1 And int单据性质 = 2 And int收费状态 = 1) Or _
                        (mParams.intShowBill记帐 = 2 And int单据性质 = 2 And int收费状态 = 2) Then
                        blnValid = True
                        Exit Do
                    End If
                    
                    lngParentID = Val(rsMsg("parent_id").Value)
                End If
                 
                If rsMsg("node_name").Value = "bill_kind" Then
                    int单据性质 = Val(rsMsg("node_value").Value)
                ElseIf rsMsg("node_name").Value = "charge_state" Then
                    int收费状态 = Val(rsMsg("node_value").Value)
                End If
            End If
          
            rsMsg.MoveNext
            
            If rsMsg.EOF Then
                '只要有一个NO满足单据性质和收费/审核状态就表示有效就退出循环
                If (mParams.intShowBill收费 = 1 And int单据性质 = 1 And int收费状态 = 1) Or _
                    (mParams.intShowBill收费 = 2 And int单据性质 = 1 And int收费状态 = 2) Or _
                    (mParams.intShowBill记帐 = 1 And int单据性质 = 2 And int收费状态 = 1) Or _
                    (mParams.intShowBill记帐 = 2 And int单据性质 = 2 And int收费状态 = 2) Then
                    blnValid = True
                    Exit Do
                End If
            End If
        Loop
        If blnValid = False Then Exit Function
    End If
    
    '3.判断发药窗口
    If strMsgCode = "ZLHIS_CHARGE_003" Then
        If objXML.GetMultiNodeRecord("charge_bill", rsMsg) = False Then Exit Function
        If rsMsg Is Nothing Then Exit Function
        If rsMsg.RecordCount = 0 Then Exit Function
        
        blnValid = False
        Do While Not rsMsg.EOF
            If rsMsg("node_name").Value = "drug_window" Then
                If InStr(mParams.Str窗口, rsMsg("node_value").Value) > 0 Or mParams.Str窗口 = "" Or rsMsg("node_value").Value = "" Then
                    blnValid = True
                    Exit Do
                End If
            End If
            rsMsg.MoveNext
        Loop
        If blnValid = False Then Exit Function
    End If
    
    IsValidMsg = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub ShowQueue()
    On Error GoTo errHandle
    
    '显示排队队列
    If mParams.blnShowQueue = True And mParams.blnStartQueue = True Then
        If gobjLEDShow Is Nothing Then
            If Not CreateObject_LED(mParams.lngShowComponent) Then Exit Sub
        End If
        
        If Not gobjLEDShow Is Nothing Then
            '关闭LCD窗口
            Call gobjLEDShow.zlDrugShowClose
            Call gobjLEDShow.zlDrugShow(mParams.lng药房ID, mParams.Str窗口, mParams.blnMustDosageProcess, mParams.blnMustDosageOkProcess)
        End If
    Else
        If Not gobjLEDShow Is Nothing Then
            '关闭LCD窗口
            Call gobjLEDShow.zlDrugShowClose
            Set gobjLEDShow = Nothing
        End If
    End If
    
    Exit Sub
errHandle:
    Set gobjLEDShow = Nothing
'    If ErrCenter = 1 Then
'        Resume
'    End If
End Sub

Private Sub ShowMedicalRecord(ByVal rsData As ADODB.Recordset)
    '【功能】：查阅当前病人的电子病案
    
    Dim int门诊 As Integer
    Dim lng主页ID As Long
    Dim rstemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    With rsData
        If Not .EOF Then
            '判断当前单据时门诊还是住院
            If !记录性质 = 1 Or (!记录性质 = 2 And (!门诊标志 = 1 Or !门诊标志 = 4)) Then
                int门诊 = 1
            Else
                int门诊 = 2
            End If

            '修正住院医生工作站将医嘱发送到门诊收费的情况
            If int门诊 = 1 And !在院 = 1 Then
                int门诊 = 2
            End If

            '调用电子病案查阅接口
            If Not mobjCISJOB Is Nothing Then
                If int门诊 = 2 Then
                    lng主页ID = Val(!主页id)
                    
                    '修正住院病人直接通过门诊收费的方式发药(未经过医嘱流程),导致主页ID为空的情况
                    If lng主页ID = 0 Then
                        gstrSQL = "Select 主页id From 在院病人 Where 病人id = [1]"
                        
                        Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "查询住院病人主页ID", !病人ID)
                        
                        If Not rstemp.EOF Then lng主页ID = rstemp!主页id
                    End If
                    
                    On Error Resume Next
                    Call mobjCISJOB.ShowArchive(Me, !病人ID, lng主页ID)
                    err.Clear: On Error GoTo 0
                Else
                    '若为门诊病人，查询对应的挂号id
                    gstrSQL = "Select ID As 挂号id From 病人挂号记录 Where 病人id = [1]"
                    
                    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "提取门诊病人挂号ID", !病人ID)
                    
                    If Not rstemp.EOF Then
                        On Error Resume Next
                        Call mobjCISJOB.ShowArchive(Me, !病人ID, rstemp!挂号id)
                        err.Clear: On Error GoTo 0
                    End If
                End If
            End If
        End If
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub zlCallMain()
    Dim intCount As Integer
    Dim dateStart As Date
    Dim strsql As String
    Dim rstemp As ADODB.Recordset
    Dim strCall As String
    Dim strCallName As String
    Dim strCallWindows As String
    Dim blnCallTime As Boolean

    '如果没有启用呼叫功能,则退出
    If mParams.blnStartCall = False Then Exit Sub
    
    '如果是全局远程语音，则退出
    If mParams.intCallType = 1 Then Exit Sub
    
    '如果上次呼叫未完成，则退出
    If mQueue.blnCallOver = False Then Exit Sub
    
    '呼叫需要一定时间，先关闭Timer控件
    blnCallTime = tmrCall.Enabled
    If blnCallTime = True Then
        tmrCall.Enabled = False
    End If
    
    mQueue.blnCallOver = False
    
    On Error GoTo errHandle
        
    '提取当前呼叫的单据
    If mQueue.blnRemoteCall = True Then
        '远端呼叫模式时允许多个叫号窗口
        strsql = "Select /*+ Rule*/ Distinct a.单据, a.NO, a.库房id, a.发药窗口, a.呼叫内容 " & _
            " From 未发药品记录 A, Table(Cast(f_Str2list([2]) As Zltools.t_Strlist)) B " & _
            " Where a.发药窗口 = b.Column_Value And (a.单据 = 8 or a.单据=9) And a.库房id = [1] And a.排队状态 = 3 And a.呼叫内容 Is Not Null "
        strCallWindows = mQueue.strSendWin
    Else
        '本地叫号模式时只能有一个叫号窗口
        strsql = "Select /*+ Rule*/ Distinct a.单据, a.NO, a.库房id, a.发药窗口, a.呼叫内容 " & _
            " From 未发药品记录 A, Table(Cast(f_Str2list([2]) As Zltools.t_Strlist)) B " & _
            " Where a.发药窗口 = b.Column_Value And (a.单据 = 8 or a.单据=9) And a.库房id = [1] And a.排队状态 = 3 And a.呼叫内容 Is Not Null "
        strCallWindows = mParams.Str窗口
    End If
    Set rstemp = zldatabase.OpenSQLRecord(strsql, "zlCallMain", mParams.lng药房ID, strCallWindows)

    While rstemp.EOF = False
        DoEvents
        FS.ShowFlash "正在呼叫...", Me
        
        strCall = rstemp!呼叫内容
        
        '当呼叫次数大于1时循环呼叫
        intCount = 0
        While intCount < mParams.intSoundTimes
            If mParams.intSoundType = CALLSOUND_MS Then
                '微软语音
                Call zlCall_MsSoundPlay(strCall, mParams.intSoundSpeed)
            Else
                '系统语音
                Call zlCall_SystemSoundPlay(strCall, mParams.intSoundSpeed)
            End If

            intCount = intCount + 1
                                         
            If mParams.intSoundTimes > 1 Then
                DoEvents
                Call Sleep(3)
            End If
        Wend

        '呼叫完后清除呼叫内容，放在刷新显示的处理后面
        gstrSQL = "Zl_未发药品记录_呼叫("
            'NO
            gstrSQL = gstrSQL & "'" & rstemp!NO & "'"
            '单据
            gstrSQL = gstrSQL & "," & rstemp!单据
            '药房id
            gstrSQL = gstrSQL & "," & rstemp!库房id
            '发药窗口
            gstrSQL = gstrSQL & ",'" & rstemp!发药窗口 & "'"
            '呼叫内容
            gstrSQL = gstrSQL & ",Null"
            gstrSQL = gstrSQL & ")"
        Call zldatabase.ExecuteProcedure(gstrSQL, "tmrCall_Timer")
        
        rstemp.MoveNext
    Wend
    
    DoEvents
    FS.StopFlash
    DoEvents

    '重新开启时间控件
    If blnCallTime = True Then
        tmrCall.Interval = mParams.intCircleTime * 1000
        tmrCall.Enabled = True
    End If
    
    mQueue.blnCallOver = True
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
    '重新开启时间控件
    If blnCallTime = True Then
        tmrCall.Interval = mParams.intCircleTime * 1000
        tmrCall.Enabled = True
    End If
    
    mQueue.blnCallOver = True
End Sub

Private Sub cbo病区_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim str工作性质 As String
    
    str工作性质 = "F"
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cbo病区.ListCount = 0 Then Exit Sub
    
    If cbo病区.ListIndex >= 0 Then
        If Val(cbo病区.Tag) = cbo病区.ItemData(cbo病区.ListIndex) Then
            Exit Sub
        End If
    End If
    
    If Select部门选择器(Me, cbo病区, Trim(cbo病区.Text), str工作性质, , "2,3") = False Then
        Exit Sub
    End If
    If cbo病区.ListIndex >= 0 Then
        cbo病区.Tag = cbo病区.ItemData(cbo病区.ListIndex)
    End If
End Sub

Private Sub cbo病区_KeyPress(KeyAscii As Integer)
    '屏蔽输入单引号
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub cbo病区_LostFocus()
    If cbo病区.ListIndex = -1 Then
        cbo病区.Text = ""
    End If
End Sub

Private Sub cbo病区_Validate(Cancel As Boolean)
    If cbo病区.ListIndex = -1 Then
        cbo病区.Text = ""
    End If
End Sub

Private Sub chk显示已确认单据_Click()
    Call SaveSetting("ZLSOFT", "公共模块\操作\" & App.ProductName & "\" & "药品处方发药", "显示已确认单据", chk显示已确认单据.Value)
    RefreshList mcondition.intListType
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strNo As String)
    If Not txtPati.Locked And txtPati.Text = "" And Me.ActiveControl Is txtPati And strNo <> "" Then
        txtPati.Text = strNo
        
        If txtPati.Text = "" Then
            Call mobjICCard.SetEnabled(False)
        Else
            mParams.int输入模式 = mFindType.IC卡
'            Call SetInputState(mParams.int输入模式)
            
            DoEvents
            
            Call txtPati_KeyPress(vbKeyReturn)
        End If
    End If
End Sub
Private Sub BillPrint_Back()
    '打印退费单据
    
    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_8", Me, "药房=" & mParams.lng药房ID)
End Sub

Private Sub BillPrint_Custom(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '打印自定义报表
    
    '默认参数：药品=药品id，药房=药房id，NO=处方NO，单据类型=药品收发记录.单据，病人ID=病人ID
    Dim lng病人ID As Long
    Dim rsTmp As ADODB.Recordset
    Dim strName As String
    Dim str当前处方 As String
    Dim Int单据 As Integer, strNo As String
    
    strName = Split(Control.Parameter, ",")(1)
    
    If strName = "ZL" & glngSys \ 100 & "_INSIDE_1341" Then
        Call ReportOpen(gcnOracle, glngSys, strName, Me)
    Else
        str当前处方 = mfrmList.GetCurrentRecipe
    
        If str当前处方 <> "" Then
            Int单据 = Val(Split(str当前处方, "|")(0))
            strNo = Split(str当前处方, "|")(1)
            lng病人ID = Val(Split(str当前处方, "|")(3))
        End If
        
        Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), strName, Me, _
            "药品=" & IIf(mSQLCondition.lng药品id = 0, "", mSQLCondition.lng药品id), _
            "药房=" & IIf(mParams.lng药房ID = 0, "", mParams.lng药房ID), _
            "NO=" & strNo, _
            "单据类型=" & IIf(Int单据 = 0, "", Int单据), _
            "病人ID=" & IIf(lng病人ID = 0, "", lng病人ID))
    End If
End Sub
Private Sub BillPrint_Dosage()
    '打印配药单
    Dim str当前处方 As String
    Dim Int单据 As Integer, strNo As String
    Dim strUnit As String
    Dim int门诊标志 As Integer
    Dim int处方类型 As Integer
    Dim str收费类别 As String
    Dim lngRow As Long
    
    On Error GoTo errHandle
    
    str当前处方 = mfrmList.GetCurrentRecipe
    
    If str当前处方 = "" Then Exit Sub
    
    Int单据 = Val(Split(str当前处方, "|")(0))
    strNo = Split(str当前处方, "|")(1)
    int门诊标志 = Val(Split(str当前处方, "|")(5))
    int处方类型 = Val(Split(str当前处方, "|")(6))
    str收费类别 = Split(str当前处方, "|")(7)
    lngRow = Val(Split(str当前处方, "|")(12))
    
    '检查单据是否存在
    If Not CheckBillExist(Int单据, strNo) Then
        MsgBox "未找到指定单据,可能已经被其他操作员处理,操作被迫中止！", vbInformation, gstrSysName
        RefreshList mcondition.intListType
        Exit Sub
    End If
    
    strUnit = GetUnit(mParams.lng药房ID, Int单据, strNo, int门诊标志)

    If str收费类别 = "1" Then
        SetLocatePrinter int处方类型, Val(Split(mParams.str配药格式, ";")(0)) - 1
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
            "NO=" & strNo, "性质=" & IIf(Int单据 = 8, 1, 2), "药房=" & mParams.lng药房ID, "包装系数=" & IIf(strUnit = "门诊单位", "C.门诊包装", "C.住院包装"), "ReportFormat=" & Val(Split(mParams.str配药格式, ";")(0)), "PrintEmpty=0", 2)
        
        '恢复处方签的本地打印机设置
        Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTDefaultScheme_Recipt, mParams.strDefaultPrinter)
    ElseIf str收费类别 = "2" Then
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
            "NO=" & strNo, "性质=" & IIf(Int单据 = 8, 1, 2), "ReportFormat=" & Val(Split(mParams.str配药格式, ";")(1)), "PrintEmpty=0", 2)
    Else
        '同时打印中药和西药的处方签
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
            "NO=" & strNo, "性质=" & IIf(Int单据 = 8, 1, 2), "ReportFormat=" & Val(Split(mParams.str配药格式, ";")(1)), "PrintEmpty=0", 2)
            
        SetLocatePrinter int处方类型, Val(Split(mParams.str配药格式, ";")(0)) - 1
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
            "NO=" & strNo, "性质=" & IIf(Int单据 = 8, 1, 2), "药房=" & mParams.lng药房ID, "包装系数=" & IIf(strUnit = "门诊单位", "C.门诊包装", "C.住院包装"), "ReportFormat=" & Val(Split(mParams.str配药格式, ";")(0)), "PrintEmpty=0", 2)
        
        '恢复处方签的本地打印机设置
        Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTDefaultScheme_Recipt, mParams.strDefaultPrinter)
    End If
    
    gstrSQL = "Zl_未发药品记录_更新打印状态("
    '单据
    gstrSQL = gstrSQL & Int单据
    'NO
    gstrSQL = gstrSQL & ",'" & strNo & "'"
    '库房ID
    gstrSQL = gstrSQL & "," & mParams.lng药房ID
    '来源科室
    gstrSQL = gstrSQL & ",Null"
    '打印内容
    gstrSQL = gstrSQL & ",3"
    gstrSQL = gstrSQL & ")"
    Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-更新单据已打印")
    
    '更新列表打印标识
    Call mfrmList.SetPrintFlag(lngRow)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub BillPrint_Lable()
    '打印药品标签
    Dim str当前处方 As String
    Dim Int单据 As Integer, strNo As String
    Dim strUnit As String
    Dim int门诊标志 As Integer
    Dim str收费类别 As String
    
    str当前处方 = mfrmList.GetCurrentRecipe
    
    If str当前处方 = "" Then Exit Sub
    
    Int单据 = Val(Split(str当前处方, "|")(0))
    strNo = Split(str当前处方, "|")(1)
    int门诊标志 = Val(Split(str当前处方, "|")(5))
    str收费类别 = Split(str当前处方, "|")(7)
    
    strUnit = GetUnit(mParams.lng药房ID, Int单据, strNo, int门诊标志)
    
    '检查单据是否存在
    If Not CheckBillExist(Int单据, strNo) Then
        MsgBox "未找到指定单据,可能已经被其他操作员处理,操作被迫中止！", vbInformation, gstrSysName
        RefreshList mcondition.intListType
        Exit Sub
    End If
    
    If str收费类别 = "1" Then
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_6", Me, _
            "NO=" & strNo, "性质=" & IIf(Int单据 = 8, 1, 2), "药房=" & mParams.lng药房ID, "包装系数=" & IIf(strUnit = "门诊单位", "D.门诊包装", "D.住院包装"), "PrintEmpty=0", 2)
    ElseIf str收费类别 = "2" Then
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_7", Me, _
            "NO=" & strNo, "药房=" & mParams.lng药房ID, "性质=" & IIf(Int单据 = 8, 1, 2), "PrintEmpty=0", 2)
    Else
        '同时打印中药和西药的药品标签
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_6", Me, _
            "NO=" & strNo, "性质=" & IIf(Int单据 = 8, 1, 2), "药房=" & mParams.lng药房ID, "包装系数=" & IIf(strUnit = "门诊单位", "D.门诊包装", "D.住院包装"), "PrintEmpty=0", 2)
            
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_7", Me, _
            "NO=" & strNo, "药房=" & mParams.lng药房ID, "性质=" & IIf(Int单据 = 8, 1, 2), "PrintEmpty=0", 2)
    End If
    
End Sub

Private Sub BillPrint_Recipe()
    '打印处方签
    Dim str当前处方 As String
    Dim Int单据 As Integer, strNo As String
    Dim strUnit As String
    Dim int门诊标志 As Integer
    Dim int处方类型 As Integer
    Dim str收费类别 As String
    
    str当前处方 = mfrmList.GetCurrentRecipe
    
    If str当前处方 = "" Then Exit Sub
    
    Int单据 = Val(Split(str当前处方, "|")(0))
    strNo = Split(str当前处方, "|")(1)
    int门诊标志 = Val(Split(str当前处方, "|")(5))
    int处方类型 = Val(Split(str当前处方, "|")(6))
    str收费类别 = Split(str当前处方, "|")(7)
    
    strUnit = GetUnit(mParams.lng药房ID, Int单据, strNo, int门诊标志)
    
    If str收费类别 = "1" Then
        SetLocatePrinter int处方类型, Val(Split(mParams.str处方格式, ";")(0)) - 1
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
            "NO=" & strNo, _
            "性质=" & IIf(Int单据 = 8, 1, 2), _
            "药房=" & mParams.lng药房ID, "包装系数=" & IIf(strUnit = "门诊单位", "C.门诊包装", "C.住院包装"), _
            "ReportFormat=" & Val(Split(mParams.str处方格式, ";")(0)), "PrintEmpty=0", 2)
        
        '恢复处方签的本地打印机设置
        Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTDefaultScheme_Recipt, mParams.strDefaultPrinter)
    ElseIf str收费类别 = "2" Then
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
            "NO=" & strNo, _
            "性质=" & IIf(Int单据 = 8, 1, 2), _
            "ReportFormat=" & Val(Split(mParams.str处方格式, ";")(1)), "PrintEmpty=0", 2)
    Else
        '同时打印中药和西药的处方签
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
            "NO=" & strNo, _
            "性质=" & IIf(Int单据 = 8, 1, 2), _
            "ReportFormat=" & Val(Split(mParams.str处方格式, ";")(1)), "PrintEmpty=0", 2)
        
        SetLocatePrinter int处方类型, Val(Split(mParams.str处方格式, ";")(0)) - 1
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
            "NO=" & strNo, _
            "性质=" & IIf(Int单据 = 8, 1, 2), _
            "药房=" & mParams.lng药房ID, "包装系数=" & IIf(strUnit = "门诊单位", "C.门诊包装", "C.住院包装"), _
            "ReportFormat=" & Val(Split(mParams.str处方格式, ";")(0)), "PrintEmpty=0", 2)
        
        '恢复处方签的本地打印机设置
        Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTDefaultScheme_Recipt, mParams.strDefaultPrinter)
    End If
    
End Sub

Private Sub BillPrint_Report()
    '打印发药清单
    
    Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1341_2", "ZL8_BILL_1341_2"), Me, _
        "库房=" & mstrStockName & "|" & mParams.lng药房ID, _
        "包装系数=" & IIf(mintUnit = mconint门诊单位, "D.门诊包装", "D.住院包装"))
End Sub

Private Sub BillPrint_Return()
    '打印退药通知单
    Dim str当前处方 As String
    Dim Int单据 As Integer, strNo As String, Str发药时间 As String
    Dim strUnit As String
    Dim int门诊标志 As Integer
    
    str当前处方 = mfrmList.GetCurrentRecipe
    
    If str当前处方 = "" Then Exit Sub
    
    Int单据 = Val(Split(str当前处方, "|")(0))
    strNo = Split(str当前处方, "|")(1)
    Str发药时间 = Split(str当前处方, "|")(2)
    int门诊标志 = Split(str当前处方, "|")(5)
    
    strUnit = GetUnit(mParams.lng药房ID, Int单据, strNo, int门诊标志)
    
    Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1341_1", "ZL8_BILL_1341_1"), _
    Me, "No=" & strNo, "单据=" & Int单据, "包装系数=" & IIf(strUnit = "门诊单位", "D.门诊包装", "D.住院包装"), "退药时间=" & Str发药时间, 2)
End Sub

Private Sub BillPrint_Change()
    '打印医嘱更改通知单
    Dim str当前处方 As String
    Dim Int单据 As Integer, strNo As String, Str发药时间 As String
    Dim strUnit As String
    Dim int门诊标志 As Integer
    
    str当前处方 = mfrmList.GetCurrentRecipe
    
    If str当前处方 = "" Then Exit Sub
    
    Int单据 = Val(Split(str当前处方, "|")(0))
    strNo = Split(str当前处方, "|")(1)
    
    strUnit = GetUnit(mParams.lng药房ID, Int单据, strNo, 1)
    
    Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1341_9", "ZL1_BILL_1341_9"), _
    Me, "No=" & strNo, "包装系数=" & IIf(strUnit = "门诊单位", "C.门诊包装", "c.住院包装"), "发药库房=" & mParams.lng药房ID, 2)
End Sub

Private Sub ChangeDosagePeople()
    '切换配药人
    Dim strName As String
    
    SetTimerState False
    
    strName = zldatabase.UserIdentify(Me, "校验配药人", glngSys, 1341, "配药")
    
    SetTimerState True
    
    If Trim(strName) = "" Then Exit Sub
    
    mstr自动配药人 = strName
    
    mdate上次校验时间 = Sys.Currentdate
End Sub

Private Function CheckCard(ByVal rsData As ADODB.Recordset) As Boolean
    '一卡通消费刷卡验证
    Dim dblSumMoney As Double
    Dim lng病人ID As Long
    Dim blnCheck As Boolean
    Dim bytType As Byte '门诊住院标志
    
    If mParams.bln发药前收费或审核 = True Then
        CheckCard = True
        Exit Function
    End If
    
    If mParams.bln发药后刷卡验证 = True Then
        blnCheck = True
        rsData.Filter = "标志=1 And 病人ID>0 And 就诊卡号<>''"
    ElseIf mParams.bln审核划价单 = True And mParams.bln刷卡验证 = True Then
        blnCheck = True
        rsData.Filter = "门诊标志=1 And 标志=1 And 记录性质=2 And 记录状态=0 And 病人ID>0 And 就诊卡号<>''"
    End If
    
    If blnCheck = True Then
        rsData.Sort = "病人ID"

        With rsData
            Do While Not .EOF
                If Val(!记录性质) = 1 Or (Val(!记录性质) = 2 And (Val(!门诊标志)) = 1 Or (Val(!门诊标志)) = 4) Then
                    bytType = 1
                Else
                    bytType = 2
                End If
            
                If lng病人ID <> !病人ID Then
                    If lng病人ID <> 0 Then
                        If zldatabase.PatiIdentify(Me, glngSys, lng病人ID, dblSumMoney, mlngMode, bytType) = False Then Exit Function
                    End If
                    
                    dblSumMoney = !实收金额
                    lng病人ID = !病人ID
                Else
                    dblSumMoney = dblSumMoney + !实收金额
                End If

                .MoveNext

                If .EOF Then
                    If zldatabase.PatiIdentify(Me, glngSys, lng病人ID, dblSumMoney, mlngMode, bytType) = False Then Exit Function
                End If
            Loop
        End With
    End If

    CheckCard = True
End Function

Public Sub ClearForm_Detail()
    If Not mfrmDetail Is Nothing Then mfrmDetail.FormClear
End Sub

Public Sub ClearForm_Recipe()
    If Not mfrmDetail Is Nothing Then mfrmRecipe.FormClear
End Sub
Public Sub FindListRow(ByVal intFindType As Integer, ByVal strFind As String, ByVal str姓名 As String)
    If Not mfrmList Is Nothing Then
        mfrmList.FindSpecialRow "单据号", strFind, "", mobjSquareCard, str姓名
    End If
End Sub

Private Sub GetDosage(ByVal lng药房ID As Long)
    '取药房配药控制信息
    On Error GoTo errHandle
    gstrSQL = "Select 配药, Nvl(门诊, 1) As 门诊 From 药房配药控制 Where 药房id = [1]"
    Set mrsIsDosage = zldatabase.OpenSQLRecord(gstrSQL, "取药房配药控制信息", lng药房ID)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetDrugStock(ByVal lng库房ID As Long)
    On Error GoTo errHandle
    gstrSQL = "Select 收费细目id As 药品ID From 收费执行科室 Where 执行科室id = [1]"
    Set mrsDrugStock = zldatabase.OpenSQLRecord(gstrSQL, "取存储库房", lng库房ID)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function GetRecipeByNO(ByVal strNo As String, Optional ByVal int查询 As Integer) As ADODB.Recordset
'功能：获取指定处方号的药品数据
'参数：
'  strNO：处方号
'  int查询:1-[退药]标签下可以查询[待配药]、[待发药]中的单据
'返回：药品记录集对象

    Dim rsData As ADODB.Recordset
    Dim strMsg As String
    Dim strsql As String
    
    If strNo = "" Then Exit Function
    On Error GoTo errHandle
    If mcondition.intListType <> mListType.退药 Or int查询 = 1 Then
        '先判断单据是否存在
        gstrSQL = "Select Distinct A.记录状态,A.NO, A.单据, B.姓名, Decode(A.单据,8,'收费',9,'记帐') 类型, A.库房id As 药房ID, C.名称 As 药房, " & _
                  "    B.记录性质, A.填制日期, B.门诊标志, a.配药日期, a.审核日期 " & vbNewLine & _
                  "From 药品收发记录 A, 门诊费用记录 B, 部门表 C, 部门表 D " & vbNewLine & _
                  "Where A.费用id = B.ID And A.库房id = C.ID And A.对方部门id = D.ID And Nvl(B.费用状态,0)<>1 " & _
                  "    And mod(A.记录状态,3)=1 And A.NO = [1] "
            
        If mPrives.bln发其它药房的处方 = False Or mcondition.intListType = mListType.已配药 Then
            gstrSQL = gstrSQL & " And (Nvl(A.库房id, 0) = 0 Or A.库房id + 0 = [2]) "
        End If
        
        If mstrDeptNode <> "" Then
            gstrSQL = gstrSQL & " And (D.站点 = [3] Or D.站点 Is Null) "
        End If
        
        If mcondition.int服务对象 = 3 Then
            gstrSQL = gstrSQL & " And A.单据 In (8,9)" '门诊及住院所有单据
            strsql = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
            strsql = Replace(strsql, "And Nvl(B.费用状态,0)<>1", "")
            gstrSQL = gstrSQL & " Union All " & strsql
        ElseIf mcondition.int服务对象 = 1 Then
            gstrSQL = gstrSQL & " And A.单据 In (8,9) " '门诊划价及门诊记帐
        Else
            gstrSQL = gstrSQL & " And A.单据 = 9 " '住院记帐
            gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
            gstrSQL = Replace(gstrSQL, "And Nvl(B.费用状态,0)<>1", "")
        End If
        
        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "判断单据是否存在", strNo, mSQLCondition.lng药房ID, mstrDeptNode)
        
        If rsData.EOF Then
            Set GetRecipeByNO = Nothing
            Exit Function
        End If
        
        Set GetRecipeByNO = rsData
    Else
        gstrSQL = " Select Distinct A.记录状态,P.名称 As 药房,Decode(A.单据,8,'收费',9,'记帐') 类型,Decode(A.单据,8,'收费',9,'记帐') 类型,A.No,A.单据,H.姓名,A.库房id as 药房id, '' 配药人,'' 审核人, '' 审核日期,H.门诊标志,H.记录性质,A.填制日期 " & _
                 " From " & _
                 "     (SELECT A.ID,A.No,A.单据,A.药品ID,A.费用ID,A.批次,A.批号,A.效期," & _
                 "          DECODE(SIGN((A.实际数量*NVL(A.付数,1))-B.已发数量),0,A.付数,1) 付数," & _
                 "          DECODE(SIGN((A.实际数量*NVL(A.付数,1))-B.已发数量),0,A.实际数量,B.已发数量) 实际数量,A.记录状态," & _
                 "          A.零售价,A.零售金额,A.单量,A.频次,A.用法,A.摘要,A.审核人,A.审核日期,A.填制人,A.填制日期,A.配药人,A.对方部门ID,A.库房ID" & _
                 "      From" & _
                 "          (SELECT A.ID,A.No,A.单据,A.药品ID,A.序号,A.费用ID,A.批次,A.批号,A.效期,A.实际数量,A.付数,A.记录状态,A.零售价,A.零售金额,A.单量,A.频次,A.用法,A.摘要,A.审核人,A.审核日期,A.填制人,A.填制日期,A.配药人,A.对方部门ID,A.库房ID " & _
                 "          From 药品收发记录 A" & _
                 "          WHERE A.审核人 IS NOT NULL AND (A.记录状态=1 OR MOD(A.记录状态,3)=0)" & _
                 "          And A.库房ID+0=[2] And A.No =[1] ) A," & _
                 "          (SELECT A.NO,A.单据,A.药品ID,A.序号,SUM(NVL(A.付数,1)*A.实际数量) 已发数量" & _
                 "          From 药品收发记录 A" & _
                 "          Where A.审核人 Is Not Null" & _
                 "          And A.库房ID+0=[2] And A.No =[1] " & _
                 "          GROUP BY A.no,A.单据,A.药品ID,A.序号) B" & _
                 "      Where A.no = B.no And A.单据 = B.单据 And A.药品ID+0 = B.药品ID And A.序号 = B.序号" & _
                 "     ) A,门诊费用记录 H,部门表 P " & _
                 " Where A.库房ID=P.id And A.库房ID+0=[2] " & _
                 " And A.No =[1] " & _
                 " And A.费用ID=H.ID And (Mod(A.记录状态,3)=0 Or A.记录状态=1) And A.实际数量<>0 "
        
        If mcondition.int服务对象 = 3 Then
            gstrSQL = gstrSQL & " And A.单据 In (8,9)" '门诊及住院所有单据
            gstrSQL = gstrSQL & " Union All " & Replace(gstrSQL, "门诊费用记录", "住院费用记录")
        ElseIf mcondition.int服务对象 = 1 Then
            gstrSQL = gstrSQL & " And A.单据 In (8,9) " '门诊划价及门诊记帐
        Else
            gstrSQL = gstrSQL & " And A.单据 = 9 " '住院记帐
            gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
        End If
        
        '一张处方不可能同时存在于在线与后备表中，因此，如果数据移出，就直接从后备表中提取，否则原SQL不变
        '药品处方发药可同时对单据 IN (8,9)的单据，因此不排除可能8在线而9后备中的情况
        Dim blnMoved As Boolean
        
        blnMoved = Sys.IsMovedByNO("药品收发记录", strNo, " 单据 IN ", " (8,9)")
        
        '如果存在数据转出，则需要同时从后备表中提取数据（可能存在不同类型的单据分别在线与后备表中）
        If blnMoved Then
            strsql = gstrSQL
            strsql = Replace(strsql, "药品收发记录", "H药品收发记录")
            strsql = Replace(strsql, "门诊费用记录", "H门诊费用记录")
            strsql = Replace(strsql, "住院费用记录", "H住院费用记录")
            gstrSQL = gstrSQL & " UNION ALL " & strsql
        End If
        
        Set GetRecipeByNO = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, mSQLCondition.lng药房ID)
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetStockName(ByVal lng库房ID As Long)
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = "Select 名称 From 部门表 Where ID = [1]"
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "取库房名称", lng库房ID)
    
    If Not rsTmp.EOF Then
        mstrStockName = rsTmp!名称
    Else
        mstrStockName = ""
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub PrintRecipe()
    '打印处方签
    Dim blnPrint As Boolean
    Dim arrRecipe
    Dim n As Integer
    Dim intNum As Integer
    Dim strRecipeNo As String
    Dim intBillType As Integer
    Dim int处方类型 As Integer
    Dim str收费类别 As String
    
    If mstrPrintRecipe = "" Then Exit Sub
    
    mstrPrintRecipe = mstrPrintRecipe & "|"
    
    If mParams.IntAutoPrint < 2 Then
        blnPrint = IIf(mParams.IntAutoPrint = 1, True, False)
        If mParams.IntAutoPrint = 0 Then
            If MsgBox("打印该处方单据吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then blnPrint = True
        End If
        
        If blnPrint Then
            arrRecipe = Split(mstrPrintRecipe, "|")
            intNum = UBound(arrRecipe)
            
            For n = 0 To intNum
                If arrRecipe(n) <> "" Then
                    strRecipeNo = Split(arrRecipe(n), ",")(0)
                    intBillType = Val(Split(arrRecipe(n), ",")(1))
                    int处方类型 = Val(Split(arrRecipe(n), ",")(4))
                    str收费类别 = Split(arrRecipe(n), ",")(5)
    
                    If str收费类别 = "1" Then
                        SetLocatePrinter int处方类型, Val(Split(mParams.str处方格式, ";")(0)) - 1
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
                            "NO=" & strRecipeNo, _
                            "性质=" & IIf(intBillType = 8, 1, 2), _
                            "药房=" & mParams.lng药房ID, "包装系数=" & IIf(mParams.strUnit = "门诊单位", "C.门诊包装", "C.住院包装"), _
                            "ReportFormat=" & Val(Split(mParams.str处方格式, ";")(0)), "PrintEmpty=0", IIf(mParams.blnPreview And mParams.IntAutoPrint = 0, 0, 2))
                        
                        '恢复处方签的本地打印机设置
                        Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTDefaultScheme_Recipt, mParams.strDefaultPrinter)
                    ElseIf str收费类别 = "2" Then
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
                            "NO=" & strRecipeNo, _
                            "性质=" & IIf(intBillType = 8, 1, 2), _
                            "ReportFormat=" & Val(Split(mParams.str处方格式, ";")(1)), "PrintEmpty=0", IIf(mParams.blnPreview And mParams.IntAutoPrint = 0, 0, 2))
                    Else
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
                            "NO=" & strRecipeNo, _
                            "性质=" & IIf(intBillType = 8, 1, 2), _
                            "ReportFormat=" & Val(Split(mParams.str处方格式, ";")(1)), "PrintEmpty=0", IIf(mParams.blnPreview And mParams.IntAutoPrint = 0, 0, 2))
                            
                        SetLocatePrinter int处方类型, Val(Split(mParams.str处方格式, ";")(0)) - 1
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
                            "NO=" & strRecipeNo, _
                            "性质=" & IIf(intBillType = 8, 1, 2), _
                            "药房=" & mParams.lng药房ID, "包装系数=" & IIf(mParams.strUnit = "门诊单位", "C.门诊包装", "C.住院包装"), _
                            "ReportFormat=" & Val(Split(mParams.str处方格式, ";")(0)), "PrintEmpty=0", IIf(mParams.blnPreview And mParams.IntAutoPrint = 0, 0, 2))
                        
                        '恢复处方签的本地打印机设置
                        Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTDefaultScheme_Recipt, mParams.strDefaultPrinter)
                    End If
                End If
            Next
        End If
    End If
    
    blnPrint = False
    If mParams.int发药后自动打印药品标签 < 2 Then
        blnPrint = IIf(mParams.int发药后自动打印药品标签 = 1, True, False)
        If mParams.int发药后自动打印药品标签 = 0 Then
            If MsgBox("打印药品标签吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then blnPrint = True
        End If
    
        If blnPrint Then
            arrRecipe = Split(mstrPrintRecipe, "|")
            intNum = UBound(arrRecipe)
            
            For n = 0 To intNum
                If arrRecipe(n) <> "" Then
                    strRecipeNo = Split(arrRecipe(n), ",")(0)
                    intBillType = Val(Split(arrRecipe(n), ",")(1))
                    int处方类型 = Val(Split(arrRecipe(n), ",")(4))
                    str收费类别 = Split(arrRecipe(n), ",")(5)
                    
                    If str收费类别 = "1" Then
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_6", Me, _
                            "NO=" & strRecipeNo, "性质=" & IIf(intBillType = 8, 1, 2), "药房=" & mParams.lng药房ID, "包装系数=" & IIf(mParams.strUnit = "门诊单位", "D.门诊包装", "D.住院包装"), "PrintEmpty=0", 2)
                    ElseIf str收费类别 = "2" Then
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_7", Me, _
                            "NO=" & strRecipeNo, "药房=" & mParams.lng药房ID, "性质=" & IIf(intBillType = 8, 1, 2), "PrintEmpty=0", 2)
                    Else
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_6", Me, _
                            "NO=" & strRecipeNo, "性质=" & IIf(intBillType = 8, 1, 2), "药房=" & mParams.lng药房ID, "包装系数=" & IIf(mParams.strUnit = "门诊单位", "D.门诊包装", "D.住院包装"), "PrintEmpty=0", 2)
                        
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_7", Me, _
                            "NO=" & strRecipeNo, "药房=" & mParams.lng药房ID, "性质=" & IIf(intBillType = 8, 1, 2), "PrintEmpty=0", 2)
                    End If
                End If
            Next
        End If
    End If
    
    mstrPrintRecipe = ""
End Sub

Private Sub PrintDosage()
    '打印配药单
    Dim blnPrint As Boolean
    Dim arrRecipe
    Dim n As Integer
    Dim intNum As Integer
    Dim strRecipeNo As String
    Dim intBillType As Integer
    Dim int门诊标志 As Integer
    Dim int处方类型 As Integer
    Dim str收费类别 As String
    Dim strUnit As String
    Dim blnIsPrintForPrintDosage As Boolean
    
    On Error GoTo errHandle
     
    If mstrPrintRecipe = "" Then Exit Sub
    
    If mParams.int配药后自动打印 = 2 Then Exit Sub
    
    If mParams.int配药后自动打印 < 2 Then
        blnIsPrintForPrintDosage = IIf(mParams.int配药后自动打印 = 1, True, False)
    
        If mParams.int配药后自动打印 = 0 Then
            blnIsPrintForPrintDosage = IIf(MsgBox("打印该配药单据吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes, True, False)
        End If
        
        If blnIsPrintForPrintDosage Then
            arrRecipe = Split(mstrPrintRecipe, "|")
            intNum = UBound(arrRecipe)
            
            For n = 0 To intNum
                If arrRecipe(n) <> "" Then
                    strRecipeNo = Split(arrRecipe(n), ",")(0)
                    intBillType = Val(Split(arrRecipe(n), ",")(1))
                    int门诊标志 = Val(Split(arrRecipe(n), ",")(3))
                    int处方类型 = Val(Split(arrRecipe(n), ",")(4))
                    str收费类别 = Split(arrRecipe(n), ",")(5)
                    
                    strUnit = GetUnit(mParams.lng药房ID, intBillType, strRecipeNo, int门诊标志)
                    
                    If str收费类别 = "1" Then
                        SetLocatePrinter int处方类型, Val(Split(mParams.str配药格式, ";")(0)) - 1
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
                            "NO=" & strRecipeNo, "性质=" & IIf(intBillType = 8, 1, 2), "药房=" & mParams.lng药房ID, "包装系数=" & IIf(strUnit = "门诊单位", "C.门诊包装", "C.住院包装"), "ReportFormat=" & Val(Split(mParams.str配药格式, ";")(0)), "PrintEmpty=0", 2)
                        
                        '恢复处方签的本地打印机设置
                        Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTDefaultScheme_Recipt, mParams.strDefaultPrinter)
                    ElseIf str收费类别 = "2" Then
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
                            "NO=" & strRecipeNo, "性质=" & IIf(intBillType = 8, 1, 2), "ReportFormat=" & Val(Split(mParams.str配药格式, ";")(1)), "PrintEmpty=0", 2)
                    Else
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
                            "NO=" & strRecipeNo, "性质=" & IIf(intBillType = 8, 1, 2), "ReportFormat=" & Val(Split(mParams.str配药格式, ";")(1)), "PrintEmpty=0", 2)
                            
                        SetLocatePrinter int处方类型, Val(Split(mParams.str配药格式, ";")(0)) - 1
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
                            "NO=" & strRecipeNo, "性质=" & IIf(intBillType = 8, 1, 2), "药房=" & mParams.lng药房ID, "包装系数=" & IIf(strUnit = "门诊单位", "C.门诊包装", "C.住院包装"), "ReportFormat=" & Val(Split(mParams.str配药格式, ";")(0)), "PrintEmpty=0", 2)
                        
                        '恢复处方签的本地打印机设置
                        Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTDefaultScheme_Recipt, mParams.strDefaultPrinter)
                    End If
                End If
            Next
        End If
        
        '单独循环更新打印状态，不放到打印报表的循环处理中
        If blnIsPrintForPrintDosage Then
            arrRecipe = Split(mstrPrintRecipe, "|")
            intNum = UBound(arrRecipe)
                
            For n = 0 To intNum
                If arrRecipe(n) <> "" Then
                    strRecipeNo = Split(arrRecipe(n), ",")(0)
                    intBillType = Val(Split(arrRecipe(n), ",")(1))
                                           
                    gstrSQL = "Zl_未发药品记录_更新打印状态("
                    '单据
                    gstrSQL = gstrSQL & intBillType
                    'NO
                    gstrSQL = gstrSQL & ",'" & strRecipeNo & "'"
                    '库房ID
                    gstrSQL = gstrSQL & "," & mParams.lng药房ID
                    '来源科室
                    gstrSQL = gstrSQL & ",Null"
                    '打印内容
                    gstrSQL = gstrSQL & ",3"
                    gstrSQL = gstrSQL & ")"
                    Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-更新单据已打印")
                End If
            Next
        End If
    End If
    mstrPrintRecipe = ""
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function BillHaveHerial(ByVal strNo As String, ByVal Int单据 As Integer, ByVal int门诊 As Integer, Optional ByRef str收费细目id As String, Optional ByRef str收费类别 As String) As String
'--------------------------------------------
'检查是否有中药处方
'-------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    
    If int门诊 = 1 Then
        gstrSQL = "Select NO,收费类别,收费细目id From 门诊费用记录 Where NO=[1] And 记录状态 IN(0,1,3)" & _
            " And 记录性质=[3] And 执行部门ID+0=[2]"
    Else
        gstrSQL = "Select NO,收费类别,收费细目id From 住院费用记录 Where NO=[1] And 记录状态 IN(0,1,3)" & _
            " And 记录性质=[3] And 执行部门ID+0=[2]"
    End If
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, mParams.lng药房ID, IIf(Int单据 = 8, 1, 2))
    
    Do While Not rsTmp.EOF
        str收费类别 = str收费类别 & rsTmp!收费类别 & ","
        BillHaveHerial = BillHaveHerial & rsTmp!收费类别 & ";"
        If InStr(1, "," & str收费细目id, "," & rsTmp!收费细目id & ",") < 1 Then str收费细目id = str收费细目id & rsTmp!收费细目id & ","
        rsTmp.MoveNext
    Loop
'    BillHaveHerial = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckAnother() As Boolean
'------------------------------------------
'检查是否设置过药房和配药人,这里是根据参数发药药房来的
'------------------------------------------
    Dim BlnIn药房 As Boolean, bln住院 As Boolean, Bln单据 As Boolean
    Dim BlnSetPeople As Boolean
    Dim RecTestPeople As New ADODB.Recordset
    Dim LngOld药房ID As Long, StrOld配药人 As String
    
    CheckAnother = False
    On Error GoTo errHandle
    If mParams.lng药房ID <> 0 Then
        With RecPart
            .MoveFirst
            .Find "ID=" & mParams.lng药房ID
            BlnIn药房 = (RecPart.EOF <> True)
            
            If BlnIn药房 Then   '说明该部门仍属药房
                '取单位
                bln住院 = False

                gstrSQL = "Select nvl(服务对象,1) 服务对象 From 部门性质说明 Where 部门ID+0=[1]"
                Set RecTestPeople = zldatabase.OpenSQLRecord(gstrSQL, "取部门服务对象", mParams.lng药房ID)
                
                With RecTestPeople
                    Do While Not .EOF
                        If !服务对象 = 2 Or !服务对象 = 3 Then bln住院 = True: Exit Do
                        .MoveNext
                    Loop
                    Bln单据 = False
                    If bln住院 Then
                        If .RecordCount <> 0 Then .MoveFirst
                        Do While Not .EOF
                            If !服务对象 = 3 Then Bln单据 = True: Exit Do
                            .MoveNext
                        Loop
                    End If
                End With
                If bln住院 = False Then
                    mcondition.int服务对象 = 1
                Else
                    mcondition.int服务对象 = IIf(Bln单据, 3, 2)
                End If
            End If
        End With
    End If
    
    '设置对应的药房，如果启用自动化发药接口则必须要设置发药窗口
    If mParams.lng药房ID = 0 Or BlnIn药房 = False Or (Not mobjDrugMAC Is Nothing And InStr(1, mParams.Str窗口, ",") > 0) Then
        '调设置窗体
        With Frm发药参数设置
            If mParams.lng药房ID = 0 Or BlnIn药房 = False Then
                MsgBox IIf(mParams.str配药人 = "", "请设置药房及配药人！", "请设置药房！"), vbInformation, gstrSysName
                Set .RecPart = RecPart.Clone
                .strShow = IIf(mParams.str配药人 = "", "请设置药房及配药人！", "请设置药房！")
            Else
                MsgBox "门诊药房自动发药只能设置一个发药窗口！", vbInformation, gstrSysName
                Set .RecPart = RecPart.Clone
                .strShow = "门诊药房自动发药只能设置一个发药窗口!"
            End If
            .mstrPrivs = mstrPrivs
            .In_启用发药 = mblnLoadDrug
            .Show 1, Me
        End With
        Call GetParams

        '仍未设置药房，退出
        If mParams.lng药房ID = 0 Then Exit Function
        
        mfrmList.SetParams
        mfrmDetail.SetParams
        mfrmRecipe.SetParams
        
        '重新获取该药房的使用单位
        With RecPart
            .MoveFirst
            .Find "ID=" & mParams.lng药房ID
            BlnIn药房 = (RecPart.EOF <> True)
            
            If BlnIn药房 Then   '说明该部门仍属药房
                '取单位
                bln住院 = False

                gstrSQL = "Select nvl(服务对象,1) 服务对象 From 部门性质说明 Where 部门ID+0=[1]"
                Set RecTestPeople = zldatabase.OpenSQLRecord(gstrSQL, "取部门服务对象", mParams.lng药房ID)
                
                With RecTestPeople
                    Do While Not .EOF
                        If !服务对象 = 2 Or !服务对象 = 3 Then bln住院 = True: Exit Do
                        .MoveNext
                    Loop
                    Bln单据 = False
                    If bln住院 Then
                        If .RecordCount <> 0 Then .MoveFirst
                        Do While Not .EOF
                            If !服务对象 = 3 Then Bln单据 = True: Exit Do
                            .MoveNext
                        Loop
                    End If
                End With
                If bln住院 = False Then
                    mcondition.int服务对象 = 1
                Else
                    mcondition.int服务对象 = IIf(Bln单据, 3, 2)
                End If
            Else
                Exit Function    '非药房，退出
            End If
        End With
    End If
    
    If mParams.blnMustDosageProcess = True And mParams.str配药人 <> "|当前操作员|" Then
        LngOld药房ID = mParams.lng药房ID
        StrOld配药人 = mParams.str配药人
        
        '设置配药人
        BlnSetPeople = False
        If mParams.str配药人 = "" Then
            MsgBox "请设置配药人！", vbInformation, gstrSysName
            With Frm发药参数设置
                Set .RecPart = RecPart.Clone
                .strShow = "请设置配药人！"
                .mstrPrivs = mstrPrivs
                .In_启用发药 = mblnLoadDrug
                .Show 1, Me
            End With
            Call GetParams
            mfrmList.SetParams
            mfrmDetail.SetParams
            mfrmRecipe.SetParams

            If mParams.str配药人 = "" Then
                MsgBox "需重新设置配药人，请与系统管理员联系！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        '如果配药人非本部门,则必须重新设置
        gstrSQL = " Select Count(*) Records From 部门人员 Where 人员ID=" & _
                 " (Select Distinct ID From 人员表 Where 姓名=[2]) And " & _
                 " 部门ID+0 =[1]"
        Set RecTestPeople = zldatabase.OpenSQLRecord(gstrSQL, "取部门人员", mParams.lng药房ID, mParams.str配药人)
        
        With RecTestPeople
            If .EOF Then
                BlnSetPeople = True
            Else
                If IsNull(!Records) Then
                    BlnSetPeople = True
                Else
                    If !Records = 0 Then
                        BlnSetPeople = True
                    End If
                End If
            End If
        End With
        If BlnSetPeople Then
            MsgBox "请设置配药人（原配药人已不属于本药房）！", vbInformation, gstrSysName
            With Frm发药参数设置
                Set .RecPart = RecPart.Clone
                .strShow = "请设置配药人（原配药人已不属于本药房）！"
                .mstrPrivs = mstrPrivs
                .In_启用发药 = mblnLoadDrug
                .Show 1, Me
            End With
            Call GetParams
            mfrmList.SetParams
            mfrmDetail.SetParams
            mfrmRecipe.SetParams
        
            If mParams.str配药人 = "" Then
                MsgBox "需重新设置配药人（原配药人已不属于本药房），请与系统管理员联系！", vbInformation, gstrSysName
                Exit Function
            End If
            If StrOld配药人 = mParams.str配药人 And LngOld药房ID = mParams.lng药房ID Then Exit Function
        End If
    End If
    
    '如果启用排队叫号，但是又设置了多个发药窗口，需要重新设置参数
    If mParams.blnStartQueue = True And InStr(mstr窗口, ",") > 0 Then
        MsgBox "已启用排队叫号，不能设置多个发药窗口，请重新设置！", vbInformation, gstrSysName
        
        With Frm发药参数设置
            Set .RecPart = RecPart.Clone
            .mstrPrivs = mstrPrivs
            .In_启用发药 = mblnLoadDrug
            .Show 1, Me
        End With
        
        Call GetParams
        mfrmList.SetParams
        mfrmDetail.SetParams
        mfrmRecipe.SetParams
        
        '任未设置正确的，退出
        If mParams.blnStartQueue = True And InStr(mParams.Str窗口, ",") > 0 Then
            MsgBox "发药窗口设置不正确，不能进行发药操作！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    CheckAnother = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function DependOnCheck() As Boolean
    Dim strsql As String
    '依赖数据检测
    DependOnCheck = False
    On Error GoTo errHandle
    With RecPart
        gstrSQL = " Select A.简码||'-'||A.姓名 医生 From 人员表 A,人员性质说明 B" & _
                 " Where B.人员性质='医生' And A.ID=B.人员ID" & _
                 " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) " & _
                 " Order by A.简码"
        Call zldatabase.OpenRecordset(RecPart, gstrSQL, "依赖数据检测")
        
        If .EOF Then
            MsgBox "请初始化人员表（医生）", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    If zlStr.IsHavePrivs(mstrPrivs, "所有药房") Then
        strsql = "(Select Distinct 部门ID From 部门性质说明 Where 工作性质 Like '%药房')"
    Else
        strsql = "(Select distinct A.部门ID From 部门人员 A,部门性质说明 B " & _
                 " Where A.人员ID=[1] And A.部门ID=B.部门ID And B.工作性质 Like '%药房')"
    End If
    gstrSQL = " Select Distinct P.ID,P.名称 From 部门表 P " & _
             " Where (P.站点 = '" & gstrNodeNo & "' Or P.站点 is Null) And P.ID In " & strsql & _
             " And (P.撤档时间 Is Null Or P.撤档时间=To_Date('3000-01-01','yyyy-MM-dd'))"
    Set RecPart = zldatabase.OpenSQLRecord(gstrSQL, "取药房", glngUserId)
    
    With RecPart
        If .EOF Then
            If zlStr.IsHavePrivs(mstrPrivs, "所有药房") Then
                strsql = "请初始化药房！（部门管理）"
            Else
                strsql = "你不是药房人员，不能使用本模块！"
            End If
            MsgBox strsql, vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    DependOnCheck = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Sub GetCondition()
    Dim dteTime As Date
    Dim strName As String
    
    
    
    dteTime = Sys.Currentdate
    
    mSQLCondition.lng药房ID = mParams.lng药房ID
    
    '时间范围
    Select Case cbo时间范围.ListIndex
        Case mTimeRange.当天
            mSQLCondition.date开始日期 = CDate(Format(dteTime, "yyyy-mm-dd") & " 00:00:00")
            mSQLCondition.date结束日期 = CDate(Format(dteTime, "yyyy-mm-dd") & " 23:59:59")
        Case mTimeRange.两天内
            mSQLCondition.date开始日期 = CDate(Format(DateAdd("d", -1, dteTime), "yyyy-mm-dd") & " 00:00:00")
            mSQLCondition.date结束日期 = CDate(Format(dteTime, "yyyy-mm-dd") & " 23:59:59")
        Case mTimeRange.三天内
            mSQLCondition.date开始日期 = CDate(Format(DateAdd("d", -2, dteTime), "yyyy-mm-dd") & " 00:00:00")
            mSQLCondition.date结束日期 = CDate(Format(dteTime, "yyyy-mm-dd") & " 23:59:59")
        Case mTimeRange.指定时间范围
            mSQLCondition.date开始日期 = CDate(Format(Dtp开始时间.Value, "yyyy-mm-dd hh:mm:ss"))
            mSQLCondition.date结束日期 = CDate(Format(Dtp结束时间.Value, "yyyy-mm-dd hh:mm:ss"))
        Case Else
            mSQLCondition.date开始日期 = CDate(Format(dteTime, "yyyy-mm-dd") & " 00:00:00")
            mSQLCondition.date结束日期 = CDate(Format(dteTime, "yyyy-mm-dd") & " 23:59:59")
    End Select
    
    mcondition.bln显示退药待发单据 = (Chk显示退药待发单据.Value = 1)
    mcondition.bln显示过程单据 = (Chk显示过程单据.Value = 1)
    mcondition.bln显示已确认单据 = (chk显示已确认单据.Value = 1)
    
    If mint过滤查询 = 1 Then
        If mSQLCondition.str就诊卡 <> "" Then
            If Split(Split(mSQLCondition.str就诊卡, "|")(1), ",")(0) = "二代身份证" Then
                '身份证
                If UBound(Split(mSQLCondition.str就诊卡, "|")) > 1 Then
                    mSQLCondition.lng病人ID = Split(mSQLCondition.str就诊卡, "|")(2)
                Else
                    mSQLCondition.str身份证 = Split(mSQLCondition.str就诊卡, "|")(0)
                End If
            ElseIf Split(Split(mSQLCondition.str就诊卡, "|")(1), ",")(0) = "IC卡" Then
                'IC卡
                If Not mobjSquareCard Is Nothing Then Call mobjSquareCard.zlGetPatiID("IC卡", UCase(Trim(Split(mSQLCondition.str就诊卡, "|")(0))), False, mlngIC病人id)
                mSQLCondition.lng病人ID = mlngIC病人id
            Else
                '其他消费卡，取病人ID
                mSQLCondition.lng病人ID = zlfuncCard_GetPatiID(mobjSquareCard, Split(Split(mSQLCondition.str就诊卡, "|")(1), ",")(1), Split(mSQLCondition.str就诊卡, "|")(0))
            End If
            
            mSQLCondition.str就诊卡 = ""
        End If
    End If
    
    If imgFilter.BorderStyle = cstFilter Then
        '清空条件
        mSQLCondition.str就诊卡 = ""
        mSQLCondition.str当前NO = ""
        mSQLCondition.str门诊号 = ""
        mSQLCondition.str姓名 = ""
        mSQLCondition.str身份证 = ""
        mSQLCondition.lng病人ID = 0
        mSQLCondition.str医保号 = ""
        mSQLCondition.lng住院号 = 0
    
        Select Case IDKNType.GetCurCard.名称
            Case "单据号"
                mSQLCondition.str当前NO = txtPati.Text
            Case "门诊号"
                mSQLCondition.str门诊号 = txtPati.Text
            Case "姓名"
                mSQLCondition.str姓名 = txtPati.Text & "%"
            Case "身份证"
                mSQLCondition.str身份证 = txtPati.Text
            Case "IC卡"
                mSQLCondition.lng病人ID = mlngIC病人id
            Case "医保号"
                mSQLCondition.str医保号 = txtPati.Text
            Case "住院号"
                mSQLCondition.lng住院号 = Val(txtPati.Text)
            Case Else
                '其他消费卡，取病人ID
                mSQLCondition.lng病人ID = zlfuncCard_GetPatiID(mobjSquareCard, mobjcard.接口序号, txtPati.Text)
        End Select
    End If
    
    mSQLCondition.intOverTime = mParams.intOverTime
End Sub

Private Sub GetParams()
    Dim arrColumn
    Dim strTmp As String
    Dim bln是否配药确认 As Boolean
    Dim intParaType As Integer
    Dim rsData As ADODB.Recordset
    Dim i As Integer
    
    With mParams
        .bln允许未审核处方发药 = (gtype_UserSysParms.P6_未审核记帐处方发药 = 1)
        .bln允许未收费处方发药 = (gtype_UserSysParms.P148_未收费处方发药 = 1)
        .bln医嘱作废 = (gtype_UserSysParms.P68_门诊药嘱先作废后退药 = 0)
        .int金额保留位数 = GetDigit(0, 1, 4)
        .bln审核划价单 = True
        .bln刷卡验证 = (Val(Left(gtype_UserSysParms.P28_门诊病人消费时需要刷卡验证, 1)) = 1)
        .bln报警包含划价费用 = (gtype_UserSysParms.P98_记帐报警包含划价费用 <> 0)
        .bln发药前收费或审核 = (gtype_UserSysParms.P163_项目执行前必须先收费或先记帐审核 = 1)
        .bln启用审方 = ((gtype_UserSysParms.P240_药房处方审查 = 1 Or gtype_UserSysParms.P240_药房处方审查 = 3) And gtype_UserSysParms.P241_处方审查时机 = 2)
        
        '呼叫同时通知设备准备发药
        .blnDispensing = Val(zldatabase.GetPara("呼叫时通知开始发药", glngSys, 1341)) = 1
        
        '参数设置：基础
        .lng药房ID = Val(zldatabase.GetPara("发药药房", glngSys, 1341))
        .Str窗口 = Replace(zldatabase.GetPara("发药窗口", glngSys, 1341), "'", "")
        .str配药人 = zldatabase.GetPara("配药人", glngSys, 1341)
        .bln自动配药 = (Val(zldatabase.GetPara("自动配药", glngSys, 1341)) = 1)
        .int自动配药时限 = Val(zldatabase.GetPara("自动配药时限", glngSys, 1341))
        .IntAutoPrint = Val(zldatabase.GetPara("发药后自动打印", glngSys, 1341))
        .int配药后自动打印 = Val(zldatabase.GetPara("配药后自动打印", glngSys, 1341, 2))
        .int发药后自动打印药品标签 = Val(zldatabase.GetPara("发药后打印药品标签", glngSys, 1341, 2))
        .intShowBill收费 = Val(zldatabase.GetPara("收费处方显示方式", glngSys, 1341, 3))
        .intShowBill记帐 = Val(zldatabase.GetPara("记帐处方显示方式", glngSys, 1341, 3))
        .intShowBill配药 = Val(zldatabase.GetPara("待配药单据打印显示方式", glngSys, 1341, 0))
        .int查询未发药单据天数 = Val(zldatabase.GetPara("查询未发药单据天数", glngSys, 1341, 0))
        
        '参数设置：辅助
        .int校验发药人 = Val(zldatabase.GetPara("校验发药人", glngSys, 1341))
        .int校验配药人 = Val(zldatabase.GetPara("校验配药人", glngSys, 1341))
        .IntShowCol = Val(zldatabase.GetPara("显示付数", glngSys, 1341))
        .bln显示大小单位 = (Val(zldatabase.GetPara("显示大小单位", glngSys, 1341)) = 1)
        .int自动销帐 = Val(zldatabase.GetPara("自动销帐", glngSys, 1341))
        .bln发药后刷卡验证 = (Val(zldatabase.GetPara("发药后刷卡验证", glngSys, 1341)) = 1)
        .bln配药扫描 = (Val(zldatabase.GetPara("配药模式扫描器确认", glngSys, 1341)) = 1)
        .intOverTime = Val(zldatabase.GetPara("超时未发药品显示时间间隔", glngSys, 1341, 0))
        .intType = Val(zldatabase.GetPara("发门诊住院处方", glngSys, 1341, 0))
        .str两次刷卡发药 = zldatabase.GetPara("两次刷卡发药", glngSys, 1341, "")
        .bln发生时间过滤 = (Val(zldatabase.GetPara("药品医嘱按发生时间过滤", glngSys, 1341, 0)) = 1)
        .int金额显示 = Val(zldatabase.GetPara("金额显示方式", glngSys, 1341, 0))
        .bln取药确认 = (Val(zldatabase.GetPara("启用病人实际取药确认模式", glngSys, 1341, 0)) = 1)
        .bln发药后检查 = (Val(zldatabase.GetPara("发药后检查卫材发放情况", glngSys, 1341, 0)) = 1)
        .bln扫描后呼叫 = (Val(zldatabase.GetPara("待发药单据扫描后自动呼叫", glngSys, 1341, 0)) = 1)
        .bln配药收费 = (Val(zldatabase.GetPara("配药时对未收费的单据进行收费", glngSys, 1341, 0)) = 1)
        .int回车方式 = Val(zldatabase.GetPara("查找时系统自动回车方式", glngSys, 1341, 0))
         
        '参数设置：打印
        .intPrint = Val(zldatabase.GetPara("发现新单据是否打印", glngSys, 1341))
        .intPrintDrugLable = Val(zldatabase.GetPara("打印药品标签", glngSys, 1341))
        .int打印卫材发料清单 = Val(zldatabase.GetPara("打印卫材发料单", glngSys, 1341))
        .bln记帐单 = (Val(zldatabase.GetPara("打印包含记帐单", glngSys, 1341)) = 1)
        .strPrintWindow = Replace(zldatabase.GetPara("打印指定发药窗口", glngSys, 1341), "'", "")
        
        .lngPrintInterval = Val(zldatabase.GetPara("打印间隔", glngSys, 1341))
        .lngRefreshInterval = Val(zldatabase.GetPara("刷新间隔", glngSys, 1341))
        .lngPrintDelay = Val(zldatabase.GetPara("打印延迟", glngSys, 1341, 60))
        .lngPrintBackInterval = Val(zldatabase.GetPara("打印退费单据间隔", glngSys, 1341))
        .blnSign = Val(zldatabase.GetPara("签到时进行配药", glngSys, 1341))
        .bln打印所有格式 = (Val(zldatabase.GetPara("打印票据的所有格式", glngSys, 1341, 0)) = 1)
        .blnPreview = (Val(zldatabase.GetPara("打印处方签时先预览再打印", glngSys, 1341, 0)) = 1)
        
        '参数设置：来源科室
        .strSourceDep = zldatabase.GetPara("来源科室", glngSys, 1341)
        
        '参数设置：处方颜色
        .strUserRecipeColor = zldatabase.GetPara("处方颜色", glngSys, 1341)
        If .strUserRecipeColor = "" Then .strUserRecipeColor = GetDefaultRecipeColor
        
        '打印机列表
        .strPrinters = zldatabase.GetPara("处方对应的打印机", glngSys, 1341)
        
        '配药单和处方签指定的打印格式
        .str配药格式 = zldatabase.GetPara("配药单打印格式", glngSys, 1341, "2;2")
        .str处方格式 = zldatabase.GetPara("处方签打印格式", glngSys, 1341, "1;1")
        
        '其它参数
        .int显示病区处方 = Val(zldatabase.GetPara("显示病区处方", glngSys, 1341))
        strTmp = zldatabase.GetPara("列设置", glngSys, 1341, "0")
        .intFont = Val(zldatabase.GetPara("字体", glngSys, 1341))
        
        '取得药品名称的格式方式
        If strTmp = "" Then strTmp = "0"
'        .str列设置 = "0|药品名称,0|其它名,0|英文名,0|规格,0|批号,0|单位,0|单价,0|数量,0|金额,0|重量,0|用法,0|频次,0|用量,0|库存数,0|库房货位,0|已退数,0|准退数,0|退药数,0|备注"
        If InStr(1, strTmp, "|") > 0 Then
            .int药品名称显示 = Val(Mid(strTmp, 1, 1))
        Else
            .int药品名称显示 = Val(strTmp)
        End If
        
        '排队叫号相关参数
        .blnStartQueue = (Val(zldatabase.GetPara("启用排队叫号", glngSys, 1341, 0, Null, True, intParaType, .lng药房ID)) = 1)
        .intSoundType = Val(zldatabase.GetPara("语音类型", glngSys, 1341, 0, Null, True, intParaType, .lng药房ID))
        .blnShowQueue = (Val(zldatabase.GetPara("显示排队队列", glngSys, 1341, 1, Null, True, intParaType, .lng药房ID)) = 1)
        .blnStartCall = (Val(zldatabase.GetPara("启用语音呼叫", glngSys, 1341, 1, Null, True, intParaType, .lng药房ID)) = 1)
        .intCallType = Val(zldatabase.GetPara("叫号方式", glngSys, 1341, 0, Null, True, intParaType, .lng药房ID))
        .strRemoteCall = zldatabase.GetPara("远端呼叫站点", glngSys, 1341, "", Null, True, intParaType, .lng药房ID)
        .intSoundSpeed = Val(zldatabase.GetPara("语音广播语速", glngSys, 1341, 65, Null, True, intParaType, .lng药房ID))
        .intSoundTimes = Val(zldatabase.GetPara("语音播放次数", glngSys, 1341, 1, Null, True, intParaType, .lng药房ID))
        .lngShowComponent = Val(zldatabase.GetPara("显示设备类别", glngSys, 1341, 101, Null, True, intParaType, .lng药房ID))
        .intCircleTime = Val(zldatabase.GetPara("呼叫轮询时间", glngSys, 1341, 5, Null, True, intParaType, .lng药房ID))
        
        '取报表的格式名称（默认取第一个格式）
        If mstrRPTDefaultScheme_Recipt = "" Then
            Set rsData = DeptSendWork_Get发药单格式("ZL1_BILL_1341_3")
            If Not rsData.EOF Then
                mstrRPTDefaultScheme_Recipt = rsData!格式
                rsData.MoveNext
            End If
            If Not rsData.EOF Then mstrRPTScheme_配药单 = rsData!格式
            If rsData.RecordCount >= 3 Then
                rsData.MoveNext
                For i = 3 To rsData.RecordCount
                    mstrRPTScheme_其他格式 = mstrRPTScheme_其他格式 & IIf(mstrRPTScheme_其他格式 = "", "", ";") & rsData!格式
                    rsData.MoveNext
                Next
            End If
        End If
        '默认的西药处方签打印机，兼容以前的版本，依次从不同的位置取值
        If mstrRPTDefaultScheme_Recipt <> "" Then .strDefaultPrinter = GetSetting("ZLSOFT", "私有模块\zl9Report\LocalSet\ZL1_BILL_1341_3\" & mstrRPTDefaultScheme_Recipt, "Printer")
        If .strDefaultPrinter = "" Then .strDefaultPrinter = GetSetting("ZLSOFT", "私有模块\zl9Report\LocalSet\ZL1_BILL_1341_3\所有格式", "Printer")
        If .strDefaultPrinter = "" Then .strDefaultPrinter = GetSetting("ZLSOFT", "私有模块\zl9Report\LocalSet\ZL1_BILL_1341_3", "Printer")
        If .strDefaultPrinter = "" Then .strDefaultPrinter = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\zl9Report\LocalSet\ZL1_BILL_1341_3", "Printer")
        
        '库存检查规则
        .IntCheckStock = MediWork_GetCheckStockRule(.lng药房ID)
        
        '是否需要配药过程
        .blnMustDosageProcess = RecipeSendWork_DispensingMedi(.lng药房ID, bln是否配药确认)
        '是否需要配药确认过程
        .blnMustDosageOkProcess = bln是否配药确认
        
        'PASS
        If gintPass <> 0 And zlStr.IsHavePrivs(mstrPrivs, "合理用药监测") Then
            .blnStarPass = True
        End If
        
        mstr窗口 = .Str窗口
        If .blnStartQueue = True And .blnStartCall = True And .Str窗口 <> "" Then
            GetChildWin
        End If
        
        '站点
        mstrDeptNode = GetDeptStationNode(.lng药房ID)
        
        Load病区
    End With
End Sub
Private Sub GetPrivs()
    Dim strPrivs As String
    
    With mPrives
        .bln所有药房 = IsInString(mstrPrivs, "所有药房", ";")
        .bln发药 = IsInString(mstrPrivs, "发药", ";")
        .bln退药 = IsInString(mstrPrivs, "退药", ";")
        .bln退其它药房的处方 = IsInString(mstrPrivs, "退其它药房的处方", ";")
        .bln发已结帐处方 = IsInString(mstrPrivs, "发已结帐处方", ";")
        .bln退已结帐处方 = IsInString(mstrPrivs, "退已结帐处方", ";")
        .bln发退出院病人处方 = IsInString(mstrPrivs, "发退出院病人处方", ";")
        .bln校验处方 = IsInString(mstrPrivs, "校验处方", ";")
        .bln医生查询 = IsInString(mstrPrivs, "医生查询", ";")
        .bln合理用药监测 = IsInString(mstrPrivs, "合理用药监测", ";")
        .bln过滤附加条件 = IsInString(mstrPrivs, "过滤附加条件", ";")
        .bln修改过滤日期 = IsInString(mstrPrivs, "修改过滤日期", ";")
        .bln参数设置 = IsInString(mstrPrivs, "参数设置", ";")
        .bln发其它药房的处方 = IsInString(mstrPrivs, "发其它药房的处方", ";")
        .bln发病区处方 = IsInString(mstrPrivs, "发病区处方", ";")
        .bln配药 = IsInString(mstrPrivs, "配药", ";")
        .bln停止发药 = IsInString(mstrPrivs, "停止发药", ";")
        .bln恢复发药 = IsInString(mstrPrivs, "恢复发药", ";")
        .bln取药确认 = IsInString(mstrPrivs, "取药确认", ";")
        .bln电子病案查阅 = IsInString(mstrPrivs, "电子病案查阅", ";")
        .bln允许查询所有时间范围单据 = IsInString(mstrPrivs, "允许查询所有时间范围单据", ";")
        
        '药品自动化设备接口（虚拟模块）
        strPrivs = ";" & zl9ComLib.GetPrivFunc(glngSys, 9010) & ";"
        .bln药品自动化接口 = IsInString(strPrivs, "基本", ";")
    End With

End Sub


Private Sub Load时间范围()
    With cbo时间范围
        .Clear
        .AddItem "0-当天"
        .AddItem "1-两天内"
        .AddItem "2-三天内"
        .AddItem "3-指定时间范围"
        
        .ListIndex = 0
        .Tag = 0
    End With
End Sub

Private Sub InitApplyforcredit()
    '存在销帐申请的记录集
    Set mrsApplyforcredit = New ADODB.Recordset
    With mrsApplyforcredit
        If .State = 1 Then .Close
        
        .Fields.Append "费用ID", adDouble, 18, adFldIsNullable
        .Fields.Append "收发ID", adDouble, 18, adFldIsNullable              '药品收发ID
        .Fields.Append "标志", adDouble, 1, adFldIsNullable      '0-不允许该单据发药；1-允许该单据发药
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "药品名称", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "批号", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "数量", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "销帐申请数量", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "姓名", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "性别", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "年龄", adLongVarChar, 10, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub InitPanes()
    Dim lngHeight As Long
    
    '初始化分栏控件
    'DockingPane
    '-----------------------------------------------------
    Me.dkpMain.SetCommandBars Me.cbsMain
    Me.dkpMain.Options.UseSplitterTracker = False '实时拖动
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.AlphaDockingContext = True
'    Me.dkpMain.Options.DefaultPaneOptions = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption
    
    Dim objPaneCon As Pane
    Dim objPaneList As Pane
    Dim objPaneDetail As Pane
    
    lngHeight = 145
    
    If cbo时间范围.ListIndex <> 3 Then
        lngHeight = lngHeight - 55
    End If
    
    If lbl病区.Visible = False Then
        lngHeight = lngHeight - 25
    End If
    
    Set objPaneCon = Me.dkpMain.CreatePane(mconPane_Recipe_Condition, 230, lngHeight, DockLeftOf, Nothing)
    objPaneCon.Title = "过滤条件"
    objPaneCon.Options = PaneNoCloseable Or PaneNoFloatable
    If Val(zldatabase.GetPara("使用个性化风格")) = 0 Then objPaneCon.Hidden = False
End Sub
Private Sub InitComandBars()
    '初始化菜单：加载全部菜单，工具栏，弹出菜单等
    Dim cbrControlMain As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim objPopup As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim objPane As Pane
    Dim blnGroup As Boolean
    Dim intCount As Integer
    Dim strCardName As String
    
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    Me.cbsMain.VisualTheme = xtpThemeOffice2003

    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16

    End With
    
    Me.cbsMain.EnableCustomization False
    Me.cbsMain.Icons = frmPublic.imgPublic.Icons
    
    '-----------------------------------------------------
    '菜单定义
    Me.cbsMain.ActiveMenuBar.Title = "菜单"
    Me.cbsMain.ActiveMenuBar.EnableDocking (xtpFlagStretched)
    
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.Id = mconMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_PrintSet, "打印设置(&S)…")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Preview, "预览(&V)")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Print, "打印(&P)")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Excel, "输出到&Excel…")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Recipe_BillPrintDosage, "打印配药单(&B)")
        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Recipe_BillPrintRecipe, "打印处方签(&D)")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Recipe_BillPrintReport, "打印发药清单(&W)")
        If InStr(1, mstrPrivs, "打印已发药清单") > 0 Then
            cbrControlMain.Visible = True
        Else
            cbrControlMain.Visible = False
        End If
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Recipe_BillPrintReturn, "打印退药通知单(&R)")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Recipe_BillPrintLable, "打印药品标签(&L)")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Recipe_BillPrintBack, "打印退费单据(T)")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Recipe_BillPrintChange, "打印医嘱更改通知单(C)")
        
        If InStr(1, mstrPrivs, "打印已退费单据") > 0 Then
            cbrControlMain.Visible = True
        Else
            cbrControlMain.Visible = False
        End If
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Parameter, "参数设置(&T)")
        cbrControlMain.BeginGroup = True
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Exit, "退出(&X)")
        cbrControlMain.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_EditPopup, "编辑(&E)", -1, False)
    cbrMenuBar.Id = mconMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Dosage, "配药模式(&D)")
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Abolish, "取消模式(&A)")
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Send, "发药模式(&C)")
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Return, "退药模式(&H)")
        
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Batch, "批量发药(&B)")
'        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_SendOther, "发其它药房的处方(&F)")
        If InStr(1, mstrPrivs, "发其它药房的处方") > 0 Then
            cbrControlMain.Visible = True
        Else
            cbrControlMain.Visible = False
        End If
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_ReturnBatch, "退其它药房的处方(&T)")
        If InStr(1, mstrPrivs, "退其它药房的处方") > 0 Then
            cbrControlMain.Visible = True
        Else
            cbrControlMain.Visible = False
        End If
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_SendByBill, "按票据号发药(&I)")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_ReturnByBill, "按票据号退药(&R)")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Flag, "停止发药标记(&S)")
        cbrControlMain.Visible = (mPrives.bln停止发药 = True Or mPrives.bln恢复发药 = True)
        blnGroup = (mPrives.bln停止发药 = True Or mPrives.bln恢复发药 = True)
        cbrControlMain.BeginGroup = blnGroup
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Charge, "门诊划价(&M)")
        cbrControlMain.Visible = IsHavePrivs(mstrChargePrivs, "划价")
        cbrControlMain.BeginGroup = Not blnGroup
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Stuff, "卫材发料(&W)")
        cbrControlMain.Visible = IsHavePrivs(mstrStuffPrivs, "卫生材料发料")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_TakeDrug, "取药确认(&T)")
        cbrControlMain.Visible = (mParams.bln取药确认 And mPrives.bln取药确认)
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Change, "切换配药人(&E)")
        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Windows, "调整发药窗口(&N)")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Call, "呼叫(&G)")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Cancle, "取消确认(&G)")
        cbrControlMain.BeginGroup = True
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_SendHot, "发药")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Visible = False
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_VerifySign, "验证签名(&S)")
        cbrControlMain.Visible = gblnESign处方发药
                        
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_EMR, "病案查询(&L)")
'        cbrControlMain.BeginGroup = True
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Hot_IC, "读IC卡(&I)")
        cbrControlMain.Visible = False
        
        If TypeName(mobjDrugMAC) = "clsDrugPacker" Then
            Set cbrControlMain = .Add(xtpControlPopup, mconMenu_Edit_Recipe_AutoSend, "门诊自动化药房设置")
            
            cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_Edit_Recipe_AutoSend_Open, "启用处方上传").Checked = True
            cbrControlMain.CommandBar.Controls.Add xtpControlButton, mconMenu_Edit_Recipe_AutoSend_Set, "设置WebService服务的地址"
            cbrControlMain.CommandBar.Controls.Add xtpControlButton, mconMenu_Edit_Recipe_AutoSend_LoadDrug, "上传药品基础数据"
            cbrControlMain.CommandBar.Controls.Add xtpControlButton, mconMenu_Edit_Recipe_AutoSend_LoadStock, "上传药品库存数据"
            mblnLoadDrug = True
        End If
        
        '外挂部件有扩展功能
        Call zlPlugIn_SetMenu(glngSys, glngModul, mobjPlugIn, cbrMenuBar.CommandBar.Controls, mconMenu_Edit_PlugIn)
    End With
    
'    '自动化发药设置菜单
'    If Not gobjPackerMZ Is Nothing Then
'        Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_AutoSend, "药房自动化接口(&V)", -1, False)
'        cbrMenuBar.Id = mconMenu_AutoSend
'    End If
        
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.Id = mconMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlPopup, mconMenu_View_ToolBar, "工具栏(&T)")
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False)
        cbrControl.Checked = True
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_ToolBar_Text, "文本标签(&T)", -1, False)
        cbrControl.Checked = True
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_ToolBar_Size, "大图标(&B)", -1, False)
        cbrControl.Checked = True
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_View_StatusBar, "状态栏(&S)")
        cbrControlMain.Checked = True
        
        Set cbrControlMain = .Add(xtpControlPopup, mconMenu_View_FontSize, "字体(&F)")
        cbrControlMain.BeginGroup = True
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_FontSize_1, "小字体(&S)", -1, False)
        If mParams.intFont = 0 Then cbrControl.Checked = True
        cbrControl.Parameter = 0
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_FontSize_2, "中字体(&M)", -1, False)
        If mParams.intFont = 1 Then cbrControl.Checked = True
        cbrControl.Parameter = 1
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_FontSize_3, "大字体(&B)", -1, False)
        If mParams.intFont = 2 Then cbrControl.Checked = True
        cbrControl.Parameter = 2
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_View_Filter, "过滤(&F)")
        cbrControlMain.BeginGroup = True
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_View_Refresh, "刷新(&R)")
        cbrControlMain.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.Id = mconMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Help_Help, "帮助主题(&H)")
        Set cbrControlMain = .Add(xtpControlPopup, mconMenu_Help_Web, "&WEB上的" & gstrProductName)
        cbrControlMain.CommandBar.Controls.Add xtpControlButton, mconMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        cbrControlMain.CommandBar.Controls.Add xtpControlButton, mconMenu_Help_Web_Forum, "发送反馈(&M)", -1, False
        cbrControlMain.CommandBar.Controls.Add xtpControlButton, mconMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Help_About, "关于(&A)…")
        cbrControlMain.BeginGroup = True
    End With
    
    '快键绑定
    With Me.cbsMain.KeyBindings
        .Add FCONTROL, Asc("P"), mconMenu_File_Print
        .Add FCONTROL, Asc("D"), mconMenu_Edit_Recipe_Dosage
        .Add FCONTROL, Asc("A"), mconMenu_Edit_Recipe_Abolish
        .Add FCONTROL, Asc("C"), mconMenu_Edit_Recipe_Send
        .Add FCONTROL, Asc("H"), mconMenu_Edit_Recipe_Return
        .Add FCONTROL, Asc("Q"), mconMenu_Edit_Recipe_Cancel
    
        .Add FCONTROL, VK_F4, mconMenu_Edit_Recipe_Hot_IC
        
        .Add 0, VK_F12, mconMenu_File_Parameter
        .Add 0, VK_F5, mconMenu_View_Refresh
        .Add 0, VK_F1, mconMenu_Help_Help
        
        .Add 0, VK_F2, mconMenu_Edit_Recipe_SendHot
        
        .Add 0, VK_F6, mconMenu_File_Recipe_BillPrintDosage
        .Add 0, VK_F4, mconMenu_File_Recipe_BillPrintRecipe
        .Add 0, VK_F11, mconMenu_File_Recipe_BillPrintLable
        .Add 0, VK_F8, mconMenu_Edit_Recipe_Charge
        .Add 0, VK_F9, mconMenu_Edit_Recipe_Stuff
        .Add 0, VK_F12, mconMenu_File_Parameter
        .Add 0, VK_F7, mconMenu_View_Filter
        .Add 0, VK_F3, mconMenu_Edit_Recipe_Call
    End With

'    '设置不常用菜单
'    With Me.cbsMain.Options
'        .AddHiddenCommand conMenu_File_PrintSet
'        .AddHiddenCommand conMenu_File_Excel
'        .AddHiddenCommand conMenu_View_Refresh
'    End With
    
    '设置弹出菜单
    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_InputPopup, "录入(&I)", -1, False)
    cbrMenuBar.Id = mconMenu_InputPopup
    cbrMenuBar.Visible = False
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Input_Recipe_NO, "单据号(&0)")
        cbrControlMain.Parameter = "单|单据号|0||||||"
        cbrControlMain.Checked = True
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Input_Recipe_OPNO, "门诊号(&1)")
        cbrControlMain.Parameter = "门|门诊号|0|||||"
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Input_Recipe_Name, "姓名(&2)")
        cbrControlMain.Parameter = "姓|姓名|0|||||"
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Input_Recipe_IDCard, "身份证(&3)")
        cbrControlMain.Parameter = "身|身份证|0|||||"
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Input_Recipe_ICCard, "IC卡(&4)")
        cbrControlMain.Parameter = "IC|IC卡号|1|||||"
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Input_Recipe_MINo, "医保号(&5)")
        cbrControlMain.Parameter = "医|医保号|0|||||"
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Input_Recipe_HosNumber, "住院号(&6)")
        cbrControlMain.Parameter = "住|住院号|0|||||"
        
        '动态取其他医疗卡（主要是消费卡）
        If mstrCardType <> "" Then
            mintCardCount = UBound(Split(mstrCardType, ";")) + 1
            For intCount = 0 To UBound(Split(mstrCardType, ";"))
                '取消费卡名称
                strCardName = Split(Split(mstrCardType, ";")(intCount), "|")(1)
                
                Set cbrControlMain = .Add(xtpControlButton, mconMenu_Input_Recipe_HosNumber + intCount + 1, strCardName & "(&" & intCount + 7 & ")")
                
                '保存卡信息
                cbrControlMain.Parameter = Split(mstrCardType, ";")(intCount)
                
                If intCount = 0 Then
                    cbrControlMain.BeginGroup = True
                End If
                
                If Split(cbrControlMain.Parameter, "|")(gCardFormat.短名) = "就" Then
                    mint就诊卡长度 = Val(Split(cbrControlMain.Parameter, "|")(gCardFormat.卡号长度))
                End If
            Next
        End If
    End With
    
    '-----------------------------------------------------
    '工具栏定义
    Set cbrToolBar = Me.cbsMain.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Preview, "预览")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Print, "打印")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_View_Filter, "过滤")
        cbrControlMain.BeginGroup = True
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Charge, "划价")
        cbrControlMain.Visible = IsHavePrivs(mstrChargePrivs, "划价")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Stuff, "发料")
        cbrControlMain.Visible = IsHavePrivs(mstrStuffPrivs, "卫生材料发料")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_TakeDrug, "取药")
        cbrControlMain.Visible = (mParams.bln取药确认 And mPrives.bln取药确认)
                
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Call, "呼叫")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Cancle, "取消确认")
        
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_EMR, "病案查询")
'        cbrControlMain.BeginGroup = True
        
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_AddSign, "签名")
'        cbrControlMain.Visible = gbln药品使用电子签名
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_VerifySign, "验证签名")
        cbrControlMain.Visible = gblnESign处方发药
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_View_Refresh, "刷新")
        cbrControlMain.BeginGroup = True
        
        '电子病案查阅
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_MedicalRecord, "电子病案查阅")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Visible = mPrives.bln电子病案查阅
        
        '外挂部件有扩展功能
        Call zlPlugIn_SetToolbar(glngSys, glngModul, mobjPlugIn, cbrToolBar.Controls, mconMenu_Edit_PlugIn)
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Help_Help, "帮助")
        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Exit, "退出")
        
    End With
    For Each cbrControlMain In cbrToolBar.Controls
        cbrControlMain.Style = xtpButtonIconAndCaption
    Next
End Sub


Private Function RecipeWork_SendByBatch(ByVal intListType As Integer) As Boolean
    Dim rsBatchData As ADODB.Recordset
    Dim rsData As ADODB.Recordset
    Dim strRecipeString As String
    Dim intCount As Integer
    Dim n As Integer
    Dim arrRecipe
    Dim intBillType As Integer
    Dim strNo As String
    Dim int记录性质 As Integer
    Dim int门诊标志 As Integer
    
    strRecipeString = mfrmList.GetCurrentBatchRecipe
    
    If strRecipeString = "" Then Exit Function
    
    Set rsBatchData = New ADODB.Recordset
    With rsBatchData
        If .State = 1 Then .Close
        .Fields.Append "标志", adLongVarChar, 1, adFldIsNullable
        .Fields.Append "单据", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "收发ID", adDouble, 18, adFldIsNullable
        .Fields.Append "药品ID", adDouble, 18, adFldIsNullable
        .Fields.Append "批次", adDouble, 18, adFldIsNullable
        .Fields.Append "序号", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "品名", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "实收金额", adDouble, 18, adFldIsNullable
        .Fields.Append "记录性质", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "记录状态", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "病人ID", adDouble, 18, adFldIsNullable
        .Fields.Append "就诊卡号", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "门诊标志", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "药房ID", adDouble, 18, adFldIsNullable
        .Fields.Append "处方类型", adDouble, 1, adFldIsNullable
        .Fields.Append "收费类别", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "姓名", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "已收费", adDouble, 1, adFldIsNullable
        .Fields.Append "填制日期", adDate, , adFldIsNullable
        .Fields.Append "结算模式", adDouble, 1, adFldIsNullable
        
        .Fields.Append "药名", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "批号", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "数量", adDouble, 50, adFldIsNullable
        .Fields.Append "包装", adDouble, 50, adFldIsNullable
        .Fields.Append "单位", adLongVarChar, 1, adFldIsNullable
        .Fields.Append "性别", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "年龄", adLongVarChar, 10, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    arrRecipe = Split(strRecipeString, "|")
    intCount = UBound(arrRecipe)
    
    For n = 0 To intCount
        intBillType = Val(Split(arrRecipe(n), ",")(0))
        strNo = Split(arrRecipe(n), ",")(1)
        int记录性质 = Split(arrRecipe(n), ",")(5)
        int门诊标志 = Split(arrRecipe(n), ",")(6)
       
        Set rsData = GetRecipeRecord(intBillType, strNo, int门诊标志, int记录性质)
        
        If Not rsData Is Nothing Then
            With rsBatchData
                Do While Not rsData.EOF
                    .AddNew
                    !标志 = 1
                    !单据 = rsData!单据
                    !NO = rsData!NO
                    !收发ID = rsData!收发ID
                    !药品ID = rsData!药品ID
                    !批次 = rsData!批次
                    !序号 = rsData!序号
                    !品名 = rsData!品名
                    !实收金额 = rsData!实收金额
                    !记录性质 = rsData!记录性质
                    !记录状态 = rsData!记录状态
                    !病人ID = rsData!病人ID
                    !就诊卡号 = zlStr.NVL(rsData!就诊卡号, "")
                    !门诊标志 = rsData!门诊标志
                    !药房ID = rsData!药房ID
                    !处方类型 = rsData!处方类型
                    !收费类别 = rsData!收费类别
                    !姓名 = rsData!姓名
                    !已收费 = rsData!已收费
                    !填制日期 = rsData!填制日期
                    !结算模式 = rsData!结算模式
                    
                    !药名 = rsData!药品名称
                    !批号 = rsData!批号
                    !数量 = rsData!数量 * rsData!包装
                    !包装 = rsData!包装
                    !单位 = rsData!单位
                    !性别 = rsData!性别
                    !年龄 = rsData!年龄

                    .Update
                    
                    rsData.MoveNext
                Loop
            End With
        End If
    Next
    
    If intListType = mListType.待发药 Then
        If RecipeWork_Send(rsBatchData) = False Then
            RecipeWork_SendByBatch = False
        Else
            If imgFilter.BorderStyle = cstFilter Then
                txtPati.Text = ""
                txtPati.SetFocus
            End If
        End If
    ElseIf intListType = mListType.待配药 Then
        If RecipeWork_Dosage(rsBatchData) = False Then
            RecipeWork_SendByBatch = False
        Else
            If imgFilter.BorderStyle = cstFilter Then
                txtPati.Text = ""
                txtPati.SetFocus
            End If
        End If
    End If
End Function


Private Sub SetComandBars()
    '主窗体的菜单状态
    Dim cbrControl As CommandBarControl
    Dim cbrMenu As CommandBarControl
    
    On Error GoTo errHandle
    
    If mParams.bln取药确认 And mPrives.bln取药确认 Then
        Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_TakeDrug, , True)
        Set cbrControl = Me.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_TakeDrug, , True)
        
        If Not cbrMenu Is Nothing Then cbrMenu.Visible = (mcondition.intListType = mListType.退药)
        If Not cbrControl Is Nothing Then cbrControl.Visible = (mcondition.intListType = mListType.退药)
    End If
    
    If gblnESign处方发药 = True Then
'        Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_AddSign, , True)
'        Set cbrControl = Me.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_AddSign, , True)
'
'        If mcondition.intListType <> mListType.退药 Then
'            If Not cbrMenu Is Nothing Then cbrMenu.Enabled = False
'            If Not cbrControl Is Nothing Then cbrControl.Enabled = False
'        End If
        
        Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_VerifySign, , True)
        Set cbrControl = Me.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_VerifySign, , True)
        
        If Not cbrMenu Is Nothing Then cbrMenu.Visible = True
        If Not cbrControl Is Nothing Then cbrControl.Visible = True
        
        If Not cbrMenu Is Nothing Then cbrMenu.Enabled = False
        If Not cbrControl Is Nothing Then cbrControl.Enabled = False
    Else
        Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_VerifySign, , True)
        Set cbrControl = Me.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_VerifySign, , True)
        
        If Not cbrMenu Is Nothing Then cbrMenu.Visible = False
        If Not cbrControl Is Nothing Then cbrControl.Visible = False
    End If
    
    Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Windows, , True)
    If Not cbrMenu Is Nothing Then
        If InStr(1, ";" & mstrPrivs & ";", ";调整发药窗口;") < 1 Or mParams.Str窗口 = "" Then
            cbrMenu.Visible = False
        Else
            cbrMenu.Visible = True
        End If
    End If
    
    Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Call, , True)
    Set cbrControl = Me.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Call, , True)
    If Not cbrMenu Is Nothing Then
        If mParams.blnStartQueue And mParams.blnStartCall And InStr(1, ";" & mstrPrivs & ";", ";叫号;") > 0 Then
            cbrMenu.Visible = True
        Else
            cbrMenu.Visible = False
        End If
        
        If tbcList.Selected.index = mListType.待发药 Then
            cbrMenu.Enabled = True
        Else
            cbrMenu.Enabled = False
        End If
    End If
    
    If Not cbrControl Is Nothing Then
        If mParams.blnStartQueue And mParams.blnStartCall And InStr(1, ";" & mstrPrivs & ";", ";叫号;") > 0 Then
            cbrControl.Visible = True
        Else
            cbrControl.Visible = False
        End If
    
        If tbcList.Selected.index = mListType.待发药 Then
            cbrControl.Enabled = True
        Else
            cbrControl.Enabled = False
        End If
        
    End If
    
'    Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_SendOther, , True)
'    If Not cbrMenu Is Nothing Then
'        If InStr(1, mstrPrivs, "发其它药房的处方") > 0 Then
'            cbrMenu.Visible = True
'        Else
'            cbrMenu.Visible = False
'        End If
'    End If
'
'    Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_ReturnBatch, , True)
'    If Not cbrMenu Is Nothing Then
'        If InStr(1, mstrPrivs, "退其它药房的处方") > 0 Then
'            cbrMenu.Visible = True
'        Else
'            cbrMenu.Visible = False
'        End If
'    End If
    
    Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Cancle, , True)
    Set cbrControl = Me.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Cancle, , True)
    If tbcList.Selected.index = mListType.待发药 And mParams.blnMustDosageProcess Then
        cbrControl.Visible = True
        cbrMenu.Visible = True
        cbrMenu.Caption = "取消配药"
        cbrControl.Caption = "取消配药"
    ElseIf tbcList.Selected.index = mListType.待配药 And mParams.blnMustDosageOkProcess And InStr(1, ";" & mstrPrivs & ";", ";配药确认;") > 0 Then
        cbrControl.Visible = True
        cbrMenu.Visible = True
        cbrMenu.Caption = "取消确认"
        cbrControl.Caption = "取消确认"
    Else
        cbrControl.Visible = False
    End If
    
    If TypeName(mobjDrugMAC) = "clsDrugPacker" Then
        If mblnPackerConnect Then
            Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_AutoSend_Open, , True)
            cbrMenu.Checked = True
            cbrMenu.Enabled = True
            mblnLoadDrug = True
            
            Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_AutoSend_LoadDrug, , True)
            cbrMenu.Enabled = True
            
            Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_AutoSend_LoadStock, , True)
            cbrMenu.Enabled = True
        Else
            Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_AutoSend_Open, , True)
            cbrMenu.Checked = False
            cbrMenu.Enabled = False
            mblnLoadDrug = False
            
            Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_AutoSend_LoadDrug, , True)
            cbrMenu.Enabled = False
            
            Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_AutoSend_LoadStock, , True)
            cbrMenu.Enabled = False
        End If
    End If
    
    Exit Sub
errHandle:
    If ErrCenter <> 1 Then Resume
    Call SaveErrLog
End Sub

Public Function RecipeWork(ByVal intType As Integer, ByVal blnByNo As Boolean, Optional vsfDetail As VSFlexGrid, Optional bln未取药发药 As Boolean = False) As Boolean
    Select Case intType
        Case mListType.配药确认
            If RecipeWork_DosageOk = False Then RecipeWork = False
        Case mListType.待配药
            If imgFilter.BorderStyle = cstFilter Then
                '批量配药
                If RecipeWork_SendByBatch(mListType.待配药) = False Then RecipeWork = False
            Else
                If RecipeWork_Dosage(mfrmDetail.GetRecord) = False Then RecipeWork = False
            End If
        Case mListType.已配药
            If RecipeWork_Abolish = False Then RecipeWork = False
        Case mListType.待发药, mListType.超时未发
            mbln未取药发药 = bln未取药发药
            If imgFilter.BorderStyle = cstFilter And blnByNo = False Then
                '批量发药
                mint发药方式 = 0
                If RecipeWork_SendByBatch(mListType.待发药) = False Then RecipeWork = False
            Else
                mint发药方式 = 1
                If RecipeWork_Send(mfrmDetail.GetRecord) = False Then RecipeWork = False
                
                
            End If
        Case mListType.退药
            If RecipeWork_Return(vsfDetail) = False Then RecipeWork = False
    End Select
    
    RefreshList intType
    
    RecipeWork = True
    
    txtPati.SetFocus
    mstrScanerLastNo = ""
End Function

Private Function RecipeWork_TakeDrug() As Boolean
    '病人取药确认
    Dim blnInTrans As Boolean
    Dim str当前处方 As String
    Dim Int单据 As Integer, strNo As String
    Dim strUnit As String
    Dim int门诊 As Integer
    Dim lng签名id As Long
    Dim in未取药 As Integer
    Dim date发药时间 As Date
    
    On Error GoTo errHandle
    
    If mcondition.intListType <> mListType.退药 Then Exit Function
    
    str当前处方 = mfrmList.GetCurrentRecipe
    
    If str当前处方 = "" Then Exit Function
    
    Int单据 = Val(Split(str当前处方, "|")(0))
    strNo = Split(str当前处方, "|")(1)
    in未取药 = Val(Split(str当前处方, "|")(11))
    date发药时间 = Sys.Currentdate
    
    If in未取药 = 1 Then
        If MsgBox("是否将处方[" & strNo & "]标记为病人已取药？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    Else
        If MsgBox("是否将处方[" & strNo & "]标记为病人未取药？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    End If
    
    gstrSQL = "Zl_药品收发记录_确认取药("
    '库房ID
    gstrSQL = gstrSQL & mParams.lng药房ID
    '单据
    gstrSQL = gstrSQL & "," & Int单据
    'NO
    gstrSQL = gstrSQL & ",'" & strNo & "'"
    '是否未取药
    gstrSQL = gstrSQL & "," & IIf(in未取药 = 1, "Null", 1)
    '取药确认人员
    gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
    '取药时间
    gstrSQL = gstrSQL & ",to_date('" & date发药时间 & "','yyyy-MM-dd hh24:mi:ss') "
    gstrSQL = gstrSQL & ")"

    Call zldatabase.ExecuteProcedure(gstrSQL, "RecipeWork_TakeDrug")

    RefreshList mcondition.intListType
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function RecipeWork_DosageOk() As Boolean
    '配药确认
    Dim str操作员 As String
    Dim str当前处方 As String
    Dim int是否确认 As Integer
    Dim rsData As ADODB.Recordset
    Dim arrSql As Variant
    Dim i As Integer
    
    On Error GoTo ErrHand
    
    If mfrmDetail.CmdSend.Caption = "配药确认(&O)" Then
        str操作员 = gstrUserName
'        If mParams.blnMustDosageProcess Then
        int是否确认 = 1
'        Else
'            int是否确认 = 2
'        End If
        
    End If
    Set rsData = mfrmDetail.GetRecord
    
    rsData.Filter = "标志=1"
    rsData.Sort = "单据,No"
    arrSql = Array()
    
    
    Do While Not rsData.EOF
        If str当前处方 <> rsData!单据 & "|" & rsData!NO Then
            str当前处方 = rsData!单据 & "|" & rsData!NO
            
            '检查单据是否存在
            If Not CheckBillExist(rsData!单据, rsData!NO) Then
                MsgBox "未找到指定单据,可能已经被其他操作员处理,操作被迫中止！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        rsData.MoveNext
    Loop
    
    
    rsData.MoveFirst
    Do While Not rsData.EOF
        gstrSQL = "Zl_未发药品记录_配药确认("
            'NO
            gstrSQL = gstrSQL & "'" & rsData!NO & "'"
            '单据
            gstrSQL = gstrSQL & "," & rsData!单据
            '库房ID
            gstrSQL = gstrSQL & "," & mParams.lng药房ID
            '配药确认
            gstrSQL = gstrSQL & "," & int是否确认
            '操作员
            gstrSQL = gstrSQL & ",'" & str操作员 & "')"
            
        ReDim Preserve arrSql(UBound(arrSql) + 1)
        arrSql(UBound(arrSql)) = gstrSQL
        rsData.MoveNext
    Loop
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSql)
        Call zldatabase.ExecuteProcedure(CStr(arrSql(i)), "RecipeWork_DosageOk")
    Next
    gcnOracle.CommitTrans
    RecipeWork_DosageOk = True
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function


Private Function RecipeWork_Dosage(ByVal rsData As ADODB.Recordset) As Boolean
    '配药
    Dim blnInTrans As Boolean
    Dim str操作员 As String
    Dim str当前处方 As String
    Dim strDosUser As String
    Dim int门诊 As Integer
    Dim arrSql As Variant
    Dim i As Long
    Dim blnBeginTrans As Boolean
    Dim str签名记录 As String
    Dim date配药日期 As Date
    Dim strNosToPlugIn As String
    Dim strReserve As String
    
    On Error GoTo ErrHand
    
    mstrPrintRecipe = ""
    
    arrSql = Array()
    
    date配药日期 = Sys.Currentdate
    
    rsData.Filter = "标志=1"
    rsData.Sort = "单据,No"
    
    Do While Not rsData.EOF
        If str当前处方 <> rsData!单据 & "|" & rsData!NO Then
            str当前处方 = rsData!单据 & "|" & rsData!NO
            
            '检查单据是否存在
            If Not CheckBillExist(rsData!单据, rsData!NO) Then
                MsgBox "未找到指定单据,可能已经被其他操作员处理,操作被迫中止！", vbInformation, gstrSysName
                Exit Function
            End If
        
            '检测是否允许
            If CheckBill(rsData!药房ID, 1, rsData!单据, rsData!NO, rsData!记录性质, rsData!门诊标志) <> 0 Then Exit Function
        End If
        
        rsData.MoveNext
    Loop
    
    '启用电子签名时检查用户是否注册
    If gblnESign处方发药 = True Then
        If Not gobjESign.CheckCertificate(gstrDbUser) Then
            Exit Function
        End If
    End If
        
    '校验配药人，如果启用电子签名则不使用
    If gblnESign处方发药 = False Then
        If mParams.int校验配药人 = 1 Then
            str操作员 = zldatabase.UserIdentify(Me, "校验配药人", glngSys, 1341, "配药")
        Else
            str操作员 = mParams.str配药人
        End If
        If str操作员 = "" Then Exit Function
    End If
    
    If mParams.bln配药收费 And mParams.bln发药前收费或审核 Then
        '老的一卡通消费刷卡
        If CheckCard(rsData) = False Then Exit Function
        
        '新的消费卡刷卡消费接口
        If Not CardConfirm(rsData) Then Exit Function
    End If
    
    '先更新批次
    rsData.Filter = "标志=1"
    rsData.Sort = "单据,No"
    
    If mParams.IntCheckStock = 2 Then
        Do While Not rsData.EOF
            gstrSQL = "zl_药品收发记录_更新批次("
            '收发ID
            gstrSQL = gstrSQL & rsData!收发ID
            '药品ID
            gstrSQL = gstrSQL & "," & rsData!药品ID
            '批次
            gstrSQL = gstrSQL & "," & rsData!批次
            gstrSQL = gstrSQL & ")"
            
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = gstrSQL
            
            rsData.MoveNext
        Loop
    End If
    
    '再设置配药人
    rsData.Filter = "标志=1"
    rsData.Sort = "单据,No"
    str当前处方 = ""
    strDosUser = mfrmDetail.Get配药人
    If strDosUser = "" Then strDosUser = IIf(mParams.str配药人 = "|当前操作员|", gstrUserName, str操作员)
    
    Do While Not rsData.EOF
        If str当前处方 <> rsData!单据 & "|" & rsData!NO Then
            If Val(rsData!记录性质) = 1 Or (Val(rsData!记录性质) = 2 And (Val(rsData!门诊标志)) = 1 Or (Val(rsData!门诊标志)) = 4) Then
                int门诊 = 1
            Else
                int门诊 = 2
            End If
            
            If mPrives.bln发其它药房的处方 = True And mParams.lng药房ID <> Val(rsData!药房ID) Then
                gstrSQL = "Zl_药品收发记录_更改库房("
                '现库房ID
                gstrSQL = gstrSQL & mParams.lng药房ID
                '单据
                gstrSQL = gstrSQL & "," & rsData!单据
                'NO
                gstrSQL = gstrSQL & ",'" & rsData!NO & "'"
                '原库房ID
                gstrSQL = gstrSQL & "," & Val(rsData!药房ID)
                '门诊
                gstrSQL = gstrSQL & "," & int门诊
                '填制日期
                gstrSQL = gstrSQL & ",to_date('" & rsData!填制日期 & "','yyyy-MM-dd')"
                gstrSQL = gstrSQL & ")"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            End If
        
            str当前处方 = rsData!单据 & "|" & rsData!NO
            
            gstrSQL = "zl_药品收发记录_设置配药人("
            '库房ID
            gstrSQL = gstrSQL & mParams.lng药房ID
            '单据
            gstrSQL = gstrSQL & "," & rsData!单据
            'NO
            gstrSQL = gstrSQL & ",'" & rsData!NO & "'"
            '门诊
            gstrSQL = gstrSQL & "," & int门诊
            '配药人
            gstrSQL = gstrSQL & ",'" & IIf(gblnESign处方发药 = True, gstrUserName, IIf(mParams.int校验配药人 = 1, str操作员, strDosUser)) & "'"
            '配药日期
            gstrSQL = gstrSQL & ",to_date('" & date配药日期 & "','yyyy-MM-dd hh24:mi:ss') "
            gstrSQL = gstrSQL & ")"
            
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = gstrSQL
            
            '如果已启用了电子签名，则需要对配药人进行电子签名处理
            If gblnESign处方发药 = True And gblnESignUserStoped = False Then
                str签名记录 = ""
                If GetSignatureRecored(EsignTache.Dosage, rsData!单据, rsData!NO, mParams.lng药房ID, str签名记录, 0, date配药日期, gstrUserName) = False Then
                    Exit Function
                End If
                
                If str签名记录 <> "" Then
                    gstrSQL = "Zl_药品签名记录_Insert(" & str签名记录 & ")"
                    
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSQL
                End If
            End If

            mstrPrintRecipe = IIf(mstrPrintRecipe = "", "", mstrPrintRecipe & "|") & rsData!NO & "," & rsData!单据 & "," & rsData!记录性质 & "," & rsData!门诊标志 & "," & rsData!处方类型 & "," & rsData!收费类别
            
            strNosToPlugIn = strNosToPlugIn & rsData!单据 & "," & rsData!NO & "|"
        End If
        
        rsData.MoveNext
    Loop
    
    gcnOracle.BeginTrans
    blnInTrans = True
    For i = 0 To UBound(arrSql)
        Call zldatabase.ExecuteProcedure(CStr(arrSql(i)), "RecipeWork_Abolish")
    Next
    gcnOracle.CommitTrans
   
    blnInTrans = False
    
    PrintDosage
        
    '调用发药后的外挂接口
    If Not mobjPlugIn Is Nothing Then
        If Right(strNosToPlugIn, 1) = "|" Then strNosToPlugIn = Left(strNosToPlugIn, Len(strNosToPlugIn) - 1)
        On Error Resume Next
        mobjPlugIn.DrugDosageByRecipe mParams.lng药房ID, strNosToPlugIn, date配药日期, strReserve
        err.Clear: On Error GoTo 0
    End If
    
    RecipeWork_Dosage = True
    Exit Function
ErrHand:
    If blnInTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub VerifySign()
    Dim rsData As Recordset
    Dim int可操作 As Integer
    
    If gblnESign处方发药 = False Then Exit Sub
    Set rsData = mfrmDetail.GetRecord(int可操作)
    
'    rsData.Filter = "标志=1"
    rsData.Sort = "单据,No"
    
    If Not rsData.EOF Then
        '如果已启用了电子签名，则需要对配药人进行电子签名处理
        If VerifySignatureRecored_bak(IIf(Me.tbcList.Item(mListType.待发药).Selected = True, EsignTache.Dosage, IIf(int可操作 = 1, EsignTache.send, EsignTache.returnStep)), rsData!单据, rsData!NO, mParams.lng药房ID, 0, IIf(Me.tbcList.Item(mListType.待发药).Selected = True, rsData!配药日期, rsData!审核日期)) = False Then
            Exit Sub
        End If
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function RecipeWork_Abolish() As Boolean
    '取消配药
    Dim blnInTrans As Boolean
    Dim str操作员 As String
    Dim str当前处方 As String
    Dim rsData As ADODB.Recordset
    Dim int门诊 As Integer
    Dim arrSql As Variant
    Dim i As Long
    Dim blnBeginTrans As Boolean
    Dim lng签名id As Long
    
    On Error GoTo ErrHand
    
    '启用电子签名时检查用户是否注册
    If gblnESign处方发药 = True Then
        If Not gobjESign.CheckCertificate(gstrDbUser) Then
            Exit Function
        End If
    End If
    
    arrSql = Array()
    
    Set rsData = mfrmDetail.GetRecord
    
    rsData.Filter = "标志=1"
    rsData.Sort = "单据,No"
    
    Do While Not rsData.EOF
        If str当前处方 <> rsData!单据 & "|" & rsData!NO Then
            str当前处方 = rsData!单据 & "|" & rsData!NO
        
            '检测是否允许
            If CheckBill(rsData!药房ID, 2, rsData!单据, rsData!NO, rsData!记录性质, rsData!门诊标志) <> 0 Then Exit Function
        End If
        
        rsData.MoveNext
    Loop
    
    rsData.Filter = "标志=1"
    rsData.Sort = "单据,No"
    str当前处方 = ""
    
    Do While Not rsData.EOF
        If str当前处方 <> rsData!单据 & "|" & rsData!NO Then
            str当前处方 = rsData!单据 & "|" & rsData!NO
        
            '如果已启用了电子签名，则取消配药人电子签名
            If gblnESign处方发药 = True And gblnESignUserStoped = False Then
                lng签名id = 0
                If DelSignatureRecored_Check(EsignTache.Dosage, rsData!单据, rsData!NO, mParams.lng药房ID, lng签名id, 0, CDate(rsData!配药日期)) = False Then
                    Exit Function
                End If
                
                If lng签名id > 0 Then
                    gstrSQL = "zl_药品签名记录_Delete(" & lng签名id & ")"
                    
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSQL
                End If
            End If
                
            If Val(rsData!记录性质) = 1 Or (Val(rsData!记录性质) = 2 And (Val(rsData!门诊标志)) = 1 Or (Val(rsData!门诊标志)) = 4) Then
                int门诊 = 1
            Else
                int门诊 = 2
            End If
            
            gstrSQL = "zl_药品收发记录_设置配药人("
            '库房ID
            gstrSQL = gstrSQL & mParams.lng药房ID
            '单据
            gstrSQL = gstrSQL & "," & rsData!单据
            'NO
            gstrSQL = gstrSQL & ",'" & rsData!NO & "'"
            '门诊
            gstrSQL = gstrSQL & "," & int门诊
            '配药人
            gstrSQL = gstrSQL & ",Null"
            '配药日期
            gstrSQL = gstrSQL & ",Null"
            gstrSQL = gstrSQL & ")"

            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = gstrSQL
        End If
        
        rsData.MoveNext
    Loop
    
    gcnOracle.BeginTrans
    blnInTrans = True
    For i = 0 To UBound(arrSql)
        Call zldatabase.ExecuteProcedure(CStr(arrSql(i)), "RecipeWork_Abolish")
    Next
    gcnOracle.CommitTrans
    blnInTrans = False
    
    RecipeWork_Abolish = True
    Exit Function
ErrHand:
    If blnInTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function RecipeWork_Return(ByVal vsfDetail As VSFlexGrid) As Boolean
    Dim str日期 As String, dbl退药数 As Double, strSubSql As String
    Dim Int单据 As Integer
    Dim strNo As String
    Dim dblSumMoney  As Double
    Dim bln是否有退药  As Boolean
    Dim lngRow As Integer
    Dim rstemp As ADODB.Recordset
    Dim str序号串 As String
    Dim blnInTrans As Boolean
    Dim blnIsReturn As Boolean
    Dim int门诊 As Integer
    Dim arrSql As Variant
    Dim i As Integer
    Dim str签名记录 As String
    Dim Int退药 As Integer
    Dim strReturnInfo As String
    Dim strReserve As String
    
    On Error GoTo ErrHand
    
    '启用电子签名时检查用户是否注册
    If gblnESign处方发药 = True Then
        If Not gobjESign.CheckCertificate(gstrDbUser) Then
            Exit Function
        End If
    End If
    
    arrSql = Array()
    
    Int单据 = Val(vsfDetail.TextMatrix(1, vsfDetail.ColIndex("单据")))
    strNo = vsfDetail.TextMatrix(1, vsfDetail.ColIndex("NO"))
    
    '已转出的数据不允许操作
    If Sys.IsMovedByNO("药品收发记录", strNo, "单据 = ", Int单据) Then
        MsgBox "该处方已被转出，不允许进行退药操作！", vbInformation, gstrSysName
        Exit Function
    End If
    '检测是否允许
    If CheckBill(0, 4, Int单据, strNo, Val(vsfDetail.TextMatrix(1, vsfDetail.ColIndex("记录性质"))), Val(vsfDetail.TextMatrix(1, vsfDetail.ColIndex("门诊标志"))), True) <> 0 Then Exit Function
    Call GetBillSequence(vsfDetail)
    
    mrsList.Filter = "单据=" & Int单据 & " And NO='" & strNo & "' "
    If Not mrsList.EOF Then dblSumMoney = Val(mrsList!金额)
    
    If mstr序号 = "" Then Exit Function
    If Not IsReceiptBalance_Charge(1, mstrPrivs, Int单据, strNo, mstr序号, Val(vsfDetail.TextMatrix(1, vsfDetail.ColIndex("记录性质"))), Val(vsfDetail.TextMatrix(1, vsfDetail.ColIndex("门诊标志")))) Then Exit Function
    If Not IsOutPatient(mstrPrivs, Int单据, strNo, Val(vsfDetail.TextMatrix(1, vsfDetail.ColIndex("记录性质"))), Val(vsfDetail.TextMatrix(1, vsfDetail.ColIndex("门诊标志")))) Then Exit Function
    If Not CheckBillControl(mcondition.intListType + 1, Int单据, strNo, dblSumMoney) Then Exit Function

    If MsgBox("你确定单号为[" & strNo & "]" & "的处方退药吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    
    str日期 = Format(Sys.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    Select Case mParams.strUnit
        Case "售价单位"
            strSubSql = "*1"
        Case "门诊单位"
            strSubSql = "*Decode(门诊包装,Null,1,0,1,门诊包装)"
        Case "住院单位"
            strSubSql = "*Decode(住院包装,Null,1,0,1,住院包装)"
        Case "药库单位"
            strSubSql = "*Decode(药库包装,Null,1,0,1,药库包装)"
        End Select
    
    bln是否有退药 = False
    For lngRow = 1 To vsfDetail.rows - 2
        dbl退药数 = Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("退药数")))

        gstrSQL = " Select round(" & dbl退药数 & strSubSql & ",5) 数量 From 药品规格" & _
                     " Where 药品ID=[1]"
        Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("药品ID"))))
                     
        With rstemp
            dbl退药数 = !数量
        End With
        
        If mParams.bln显示大小单位 = True Then
            If (Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("退药数(大包装)"))) = Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("准退数大"))) And _
                Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("退药数(小包装)"))) = Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("准退数小")))) Or _
                (Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("退药数"))) = Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("准退数大"))) * Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("包装"))) + Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("准退数小")))) Then
                
                dbl退药数 = Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("实际数量")))
            End If
        Else
            If Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("退药数"))) = Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("准退数"))) Then
                dbl退药数 = Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("实际数量")))
            End If
        End If
        
        If dbl退药数 <> 0 Then
            blnIsReturn = False
            
            '先检查或执行预调价
            Call AutoAdjustPrice_ByID(Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("药品ID"))))
        
            '检查价格
            If CheckPrice(Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("id"))), mstr价格失效提示) = False Then
                If MsgBox("药品[" & vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("药品名称")) & "]" & mstr价格失效提示, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    blnIsReturn = True
                End If
            Else
                blnIsReturn = True
            End If
            
            If blnIsReturn = True Then
                If Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("记录性质"))) = 1 Or (Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("记录性质"))) = 2 And (Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("门诊标志")))) = 1 Or (Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("门诊标志")))) = 4) Then
                    int门诊 = 1
                Else
                    int门诊 = 2
                End If
                
                gstrSQL = "zl_药品收发记录_部门退药("
                '收发ID
                gstrSQL = gstrSQL & Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("id")))
                '审核人
                gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
                '审核日期
                gstrSQL = gstrSQL & ",to_date('" & str日期 & "','yyyy-MM-dd hh24:mi:ss') "
                '批号
                gstrSQL = gstrSQL & "," & IIf(Trim(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("新批号"))) = "", "NULL", "'" & vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("新批号")) & "'")
                '效期
                gstrSQL = gstrSQL & "," & IIf(Trim(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("新效期"))) = "", "NULL", "to_date('" & vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("新效期")) & "','yyyy-MM-dd')")
                '产地
                gstrSQL = gstrSQL & "," & IIf(Trim(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("新产地"))) = "", "NULL", "'" & Trim(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("新产地"))) & "'")
                '退药数
                gstrSQL = gstrSQL & "," & dbl退药数
                '退药库房
                gstrSQL = gstrSQL & ",NULL"
                '退药人
                gstrSQL = gstrSQL & ",NULL"
                '金额保留位数
                gstrSQL = gstrSQL & "," & mParams.int金额保留位数
                '门诊
                gstrSQL = gstrSQL & "," & int门诊
                gstrSQL = gstrSQL & ")"
                    
'                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-药品退药")
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            
                bln是否有退药 = True
                
                strReturnInfo = IIf(strReturnInfo = "", "", strReturnInfo & "|") & Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("id"))) & "," & dbl退药数
            End If
        End If
    Next
    
    '如果本地参数设置了自动销帐，并且当前退费单据是记帐单，那么执行门诊/住院销帐
    If mParams.int自动销帐 = 1 And Val(vsfDetail.TextMatrix(1, vsfDetail.ColIndex("记录性质"))) = 2 And bln是否有退药 = True Then
        For lngRow = 1 To vsfDetail.rows - 2
            If Val(vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("退药数"))) <> 0 Then
                str序号串 = str序号串 & IIf(str序号串 = "", vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("序号")), "," & vsfDetail.TextMatrix(lngRow, vsfDetail.ColIndex("序号")))
            End If
        Next
        If Val(vsfDetail.TextMatrix(1, vsfDetail.ColIndex("门诊标志"))) = 1 Or Val(vsfDetail.TextMatrix(1, vsfDetail.ColIndex("门诊标志"))) = 4 Then
            gstrSQL = "Zl_门诊记帐记录_Delete("
            'NO
            gstrSQL = gstrSQL & "'" & strNo & "'"
            '序号串
            gstrSQL = gstrSQL & ",'" & str序号串 & "'"
            '操作员编号
            gstrSQL = gstrSQL & ",'" & gstrUserCode & "'"
            '操作员姓名
            gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
            gstrSQL = gstrSQL & ")"
        Else
            gstrSQL = "Zl_住院记帐记录_Delete("
            'NO
            gstrSQL = gstrSQL & "'" & strNo & "'"
            '序号串
            gstrSQL = gstrSQL & ",'" & str序号串 & "'"
            '操作员编号
            gstrSQL = gstrSQL & ",'" & gstrUserCode & "'"
            '操作员姓名
            gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
            '记录性质
            gstrSQL = gstrSQL & "," & Val(vsfDetail.TextMatrix(1, vsfDetail.ColIndex("记录性质")))
            gstrSQL = gstrSQL & ")"
        End If
'        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-退药销帐")

        ReDim Preserve arrSql(UBound(arrSql) + 1)
        arrSql(UBound(arrSql)) = gstrSQL
    End If
    
    '提示停用药品
    Int退药 = 1
    Call CheckStopMedi(Int单据 & "|" & strNo, Int退药)
    If Int退药 = 2 Then Exit Function
    
    
    '集中处理退药事务
    gcnOracle.BeginTrans
    blnInTrans = True
    For i = 0 To UBound(arrSql)
        Call zldatabase.ExecuteProcedure(CStr(arrSql(i)), Me.Caption & "-药品退药")
    Next
    
    If gblnESign处方发药 = True And gblnESignUserStoped = False Then
        str签名记录 = ""
        If GetSignatureRecored(EsignTache.returnStep, Int单据, strNo, mParams.lng药房ID, str签名记录, 0, CDate(str日期)) = False Then
            gcnOracle.RollbackTrans
            blnInTrans = False
            Exit Function
        End If
        
        If str签名记录 = "" Then
            gcnOracle.RollbackTrans
            blnInTrans = False
            MsgBox "对退药人电子签名失败！", vbInformation, gstrSysName
            Exit Function
        End If
        
        If str签名记录 <> "" Then
            gstrSQL = "Zl_药品签名记录_Insert(" & str签名记录 & ")"
            Call zldatabase.ExecuteProcedure(gstrSQL, "签名处理")
        End If
    End If
    gcnOracle.CommitTrans
    blnInTrans = False
    
    '打印退费通知单
    Dim Str发药时间 As String, int包装系数 As Integer
    
    If bln是否有退药 Then
        Str发药时间 = str日期
        int包装系数 = IIf(Int单据 = 8, 1, 2)
        
        If MsgBox("你需要打印退药通知单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1341_1", "ZL8_BILL_1341_1"), _
            Me, "No=" & strNo, "单据=" & Int单据, "包装系数=" & IIf(int包装系数 = 1, "D.门诊包装", "D.住院包装"), "退药时间=" & Str发药时间, 2)
        End If
    Else
        MsgBox "本次没有退药。"
    End If
    
    '调用退药后的外挂接口
    If Not mobjPlugIn Is Nothing And bln是否有退药 Then
        On Error Resume Next
        mobjPlugIn.DrugReturnByID mParams.lng药房ID, strReturnInfo, CDate(str日期), strReserve
        err.Clear: On Error GoTo 0
    End If
    
    Exit Function
ErrHand:
    If blnInTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub ResetFilter()
    '重新设置过滤条件
    Dim strReturn As String, IntOper As Integer
    
    IntOper = mcondition.intListType + 1
    
    With Frm药品发药查找
        strReturn = .ShowMe(Me, mParams.lng药房ID, IntOper, mstrPrivs, mbln就诊卡, _
            mSQLCondition.date开始日期, _
            mSQLCondition.date结束日期, _
            mSQLCondition.str开始NO, _
            mSQLCondition.str结束NO, _
            mSQLCondition.str姓名, _
            mSQLCondition.str就诊卡, _
            mSQLCondition.str标识号, _
            mSQLCondition.lng科室ID, _
            mSQLCondition.str填制人, _
            mSQLCondition.str审核人, _
            mSQLCondition.lng药品id, _
            mSQLCondition.str医保号, _
            mcondition.int离院带药)
        If strReturn = "" Then Exit Sub
    End With
    
    mint过滤查询 = 1

    If mPrives.bln允许查询所有时间范围单据 Then
        cbo时间范围.ListIndex = 3
    Else
        cbo时间范围.ListIndex = 0
    End If
    
    Call picConMain_Resize
    Call picCondition_Resize
    Dtp开始时间.Value = mSQLCondition.date开始日期
    Dtp结束时间.Value = mSQLCondition.date结束日期
    
    If imgFilter.BorderStyle = cstFilter Then
        Call txtPati_KeyPress(13)
    Else
        Call RefreshList(mcondition.intListType)
    End If
    
    mint过滤查询 = 0
End Sub

Private Sub ResetParams()
    Dim strTmp As String
    Dim intCurrTab As Integer
    
    BlnSetParaSuccess = False
    BlnRefresh = False
    
    '关闭Timer
    Call SetTimerState(False)
    
    With Frm发药参数设置
        Set .RecPart = RecPart.Clone
        .mstrPrivs = mstrPrivs
'        .In_启用发药 = (Not gobjPackerMZ Is Nothing)
        .In_启用发药 = mblnLoadDrug
        If Not mobjMipModule Is Nothing Then
            If mobjMipModule.IsConnect = True Then
                .In_启用消息 = True
            Else
                .In_启用消息 = False
            End If
        Else
            .In_启用消息 = False
        End If
        .Show 1, Me
    End With
    
    If Not BlnSetParaSuccess Then
        '参数无变化时
    
        '开启Timer
        Call SetTimerState(True)
    Else
        '参数有变化时，更新各窗口的参数
        Call GetParams
        mfrmList.SetParams
        mfrmDetail.SetParams
        mfrmRecipe.SetParams
        
        '设置时间控件
        If mParams.lngRefreshInterval > 0 Then
            If mParams.lngRefreshInterval > 60 Then
                mParams.lngRefreshInterval = 60
            End If
            With TimeRefresh
                .Enabled = True
                .Interval = mParams.lngRefreshInterval * 1000
            End With
        Else
            TimeRefresh.Enabled = False
        End If
        
        If mParams.lngPrintInterval > 0 Then
            If mParams.lngPrintInterval > 60 Then
                mParams.lngPrintInterval = 60
            End If
            With TimePrint
                .Enabled = True
                .Interval = mParams.lngPrintInterval * 1000
            End With
        Else
            TimePrint.Enabled = False
        End If
        
        IntTimes = 0
        
        If mParams.lngPrintBackInterval <> 0 Then
            With TimePrintCancelBill
                .Enabled = False
                .Enabled = True
            End With
        Else
            TimePrintCancelBill.Enabled = False
        End If
        
        '设置叫号轮询时间间隔：当本机机器名等于全局叫号远端机器名时
        tmrCall.Enabled = False
        If mParams.blnStartQueue = True And mParams.blnStartCall = True And (mParams.intCallType = 0 And mQueue.strPCName = mParams.strRemoteCall And mQueue.strPCName <> "") Then
            tmrCall.Enabled = True
            tmrCall.Interval = mParams.intCircleTime * 1000
        End If
        
        GetDrugStock mParams.lng药房ID
        GetDosage mParams.lng药房ID
        GetStockName mParams.lng药房ID
        GetSendWindows mParams.lng药房ID
        
        If Not gobjESign Is Nothing Then
            gblnESign处方发药 = EsignIsOpen(mParams.lng药房ID)
        End If
        
        strTmp = Me.dkpMain.FindPane(mconPane_Recipe_Condition).Title
        strTmp = mstrStockName & Mid(strTmp, InStr(strTmp, ":"))
        Me.dkpMain.FindPane(mconPane_Recipe_Condition).Title = strTmp
        
        intCurrTab = mcondition.intListType
        
        If mParams.blnMustDosageProcess = True Then
'            tbcList.Item(mconTab_Recipe_Abolish).Visible = True
            tbcList.Item(mconTab_Recipe_Dosage).Visible = True
'            tbcList.Item(mconTab_Recipe_Abolish).Selected = True
'            tbcList.Item(mconTab_Recipe_Dosage).Selected = True
        Else
            tbcList.Item(mconTab_Recipe_Dosage).Visible = False
'            tbcList.Item(mconTab_Recipe_Abolish).Visible = False
'            tbcList.Item(mconTab_Recipe_Return).Selected = True
'            tbcList.Item(mconTab_Recipe_Send).Selected = True
        End If
        
'        If mParams.blnMustDosageOkProcess = True Then
'            tbcList.Item(mconTab_Recipe_DosageOk).Visible = True
'        Else
'            tbcList.Item(mconTab_Recipe_DosageOk).Visible = False
'        End If
        
        tbcList.Item(mconTab_Recipe_OverTime).Visible = (mParams.intOverTime > 0)
    
        
        If CheckAnother = False Then Exit Sub
        
        If tbcList.Item(mcondition.intListType).Visible = True Then
            tbcList.Item(mcondition.intListType).Selected = True
        Else
            If mParams.blnMustDosageProcess = True Then
                tbcList.Item(mconTab_Recipe_Dosage).Selected = True
            Else
                tbcList.Item(mconTab_Recipe_Send).Selected = True
            End If
        End If
        
        Call tbcList_SelectedChanged(tbcList.Item(mcondition.intListType))
        If intCurrTab = mcondition.intListType Then
            RefreshList mcondition.intListType
        End If
        
        '重新显示排队窗口
        If mParams.blnShowQueue And mParams.blnStartQueue Then
            Call ShowQueue
        Else
            CloseQueue
        End If
        
        Call GetOpr
        
        mbln允许两次刷卡 = False
        If mParams.str两次刷卡发药 <> "" Then
            mbln允许两次刷卡 = InStr(1, "," & mParams.str两次刷卡发药 & ",", "," & mobjcard.接口序号 & ",") > 0
        End If
    End If
End Sub

'Private Sub SetInputState(ByVal intType As Integer)
'    Dim cbrControl As CommandBarControl
'
'    Set cbrControl = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Input_Recipe_NO + intType, , True)
'    If Not cbrControl Is Nothing Then
'        SetInputPopupCheck cbrControl
'    End If
'End Sub
Private Sub SetLocatePrinter(ByVal intRecipeType As Integer, Optional ByVal int格式 As Integer)
    '打印西药处方签时，根据颜色来指定对应的打印机
    'int格式:   报表所对应的格式
    Dim strPrinter As String
    Dim i As Integer
    
    If mParams.strPrinters = "" Then Exit Sub
    
    If intRecipeType < 0 Or intRecipeType > 5 Then intRecipeType = 0
    
    On Error GoTo errHandle
    
    If InStr(mParams.strPrinters, "?") = 0 Then
        '兼容以前的存储规则
        strPrinter = Split(mParams.strPrinters, ";")(intRecipeType)
    Else
        strPrinter = Mid(Split(Split(mParams.strPrinters, ";")(intRecipeType), ",")(int格式), InStr(Split(Split(mParams.strPrinters, ";")(intRecipeType), ",")(int格式), "?") + 1)
    End If
    
    If strPrinter <> "" Then
        '保存处方类型指定的打印机到本地注册表
        Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTDefaultScheme_Recipt, strPrinter)
        Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTScheme_配药单, strPrinter)
        For i = 0 To UBound(Split(mstrRPTScheme_其他格式, ";"))
            Call SavePrinterSet("ZL1_BILL_1341_3", Split(mstrRPTScheme_其他格式, ";")(i), strPrinter)
        Next
    End If
    
    '同时打印所有格式
    If int格式 = -1 Then
        If InStr(mParams.strPrinters, "?") = 0 Then
            Exit Sub
        Else
            For i = 0 To UBound(Split(Split(mParams.strPrinters, ";")(intRecipeType), ","))
                strPrinter = Mid(Split(Split(mParams.strPrinters, ";")(intRecipeType), ",")(i), InStr(Split(Split(mParams.strPrinters, ";")(intRecipeType), ",")(i), "?") + 1)
                If strPrinter <> "" Then
                    If i = 0 Then
                        Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTDefaultScheme_Recipt, strPrinter)
                    ElseIf i = 1 Then
                        Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTScheme_配药单, strPrinter)
                    Else
                        Call SavePrinterSet("ZL1_BILL_1341_3", Split(mstrRPTScheme_其他格式, ";")(i - 2), strPrinter)
                    End If
                End If
            Next
        End If
    End If
    
    Exit Sub
errHandle:
    Resume Next
End Sub

Private Sub SavePrinterSet(ByVal strRPTCode As String, ByVal strRPTScheme As String, ByVal strPrinter As String)
    '保存打印机信息，这里用于报表打印时临时更改打印机名称；用法为调用两次，打印前更换本次打印需要的打印机，打印后恢复为默认的打印机
    '传参：strRPTCode-报表编码；strRPTScheme-报表格式；strPrinter-打印机名称
    SaveSetting "ZLSOFT", "私有模块\zl9Report\LocalSet\" & strRPTCode & "\" & strRPTScheme, "Printer", strPrinter
End Sub
Private Sub SetPaneTitle(ByVal intType As Integer)
    Dim strTitleCon As String
    Dim strTitleList As String
    
    Select Case intType
        Case mListType.配药确认
            strTitleCon = "配药确认"
        Case mListType.待配药
            strTitleCon = "待配药"
        Case mListType.已配药
            strTitleCon = "已配药"
        Case mListType.待发药
            strTitleCon = "待发药"
        Case mListType.超时未发
            strTitleCon = "超时未发"
        Case mListType.退药
            strTitleCon = "退药"
    End Select
    
    Me.dkpMain.FindPane(mconPane_Recipe_Condition).Title = mstrStockName & ":" & strTitleCon
End Sub

Private Sub SetTimerState(ByVal BlnSet As Boolean)
    '关闭和启用Timer控件，有弹出窗口时调用
    'blnSet：True-开启；False-关闭
    
    If BlnSet Then
        '开启时恢复原来的状态
        TimeRefresh.Enabled = mblnStateTimeRefresh
        TimePrint.Enabled = mblnStateTimePrint
        tmrCall.Enabled = mblnStateTimeCall
    Else
        '关闭时先记录原来的状态
        mblnStateTimeRefresh = TimeRefresh.Enabled
        mblnStateTimePrint = TimePrint.Enabled
        mblnStateTimeCall = tmrCall.Enabled
        
        If mblnStateTimeRefresh Then TimeRefresh.Enabled = False
        If mblnStateTimePrint Then TimePrint.Enabled = False
        If mblnStateTimeCall Then tmrCall.Enabled = False
    End If
End Sub
Private Sub GetBillSequence(ByVal vsfDetail As VSFlexGrid)
    Dim intRow As Integer, intRows As Integer
    Dim int序号 As Integer
    '获取当前待发药、待退药处方的有效序号
    mstr序号 = ""
    intRows = vsfDetail.rows - 2
    
    If mcondition.intListType = mListType.退药 Then
        '退药数不为零表示本次要退的明细，仅统计出这类明细的序号
        For intRow = 1 To intRows
            If Val(vsfDetail.TextMatrix(intRow, vsfDetail.ColIndex("退药数"))) <> 0 Then
                int序号 = Val(vsfDetail.TextMatrix(intRow, vsfDetail.ColIndex("序号")))
                If InStr(1, mstr序号 & ",", "," & int序号 & ",") = 0 Then
                    mstr序号 = mstr序号 & "," & int序号
                End If
            End If
        Next
    Else
        For intRow = 1 To intRows
            int序号 = Val(vsfDetail.TextMatrix(intRow, vsfDetail.ColIndex("序号")))
            If InStr(1, mstr序号 & ",", "," & int序号 & ",") = 0 Then
                mstr序号 = mstr序号 & "," & int序号
            End If
        Next
    End If
    If mstr序号 <> "" Then mstr序号 = Mid(mstr序号, 2)
End Sub
Private Function RecipeWork_Send(ByVal rsData As ADODB.Recordset) As Boolean
    '发药
    Dim str操作员 As String
    Dim str当前处方 As String
    
    On Error GoTo ErrHand
    
    mblnSendIsOver = False
    
    If rsData Is Nothing Then Exit Function
    
    rsData.Filter = "标志=1"
    rsData.Sort = "单据,No"
    
    '启用电子签名时检查用户是否注册
    If gblnESign处方发药 = True Then
        If Not gobjESign.CheckCertificate(gstrDbUser) Then
            Exit Function
        End If
    End If
    
    '批量检查处方
    If Not CheckBatchRecipe(rsData) Then Exit Function
    
    '老的一卡通消费刷卡
    If CheckCard(rsData) = False Then Exit Function
    
    '新的消费卡刷卡消费接口
    If Not CardConfirm(rsData) Then Exit Function
    
    '批量处方发药
    If Not SendBatchRecipe(rsData) Then
        Exit Function
    End If
    
    '启用支付宝之后，返回提示信息
    Call msg_upload(rsData)
    
    PrintRecipe
    
    '调用外挂接口功能（如发药后评价器功能，每次发药只调用一次）
    If Not mobjPlugIn Is Nothing Then
        rsData.MoveFirst
        On Error Resume Next
        Call mobjPlugIn.OutPatiMedicineAfter(rsData!病人ID, rsData!NO, rsData!单据, mParams.lng药房ID)
        err.Clear: On Error GoTo 0
    End If
    
    '检查该病人在该药房是否有未发的卫生材料
    If mParams.bln发药后检查 Then
        Call checkStuff(rsData!病人ID)
    End If
    
    RecipeWork_Send = True
    mblnSendIsOver = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    mblnSendIsOver = True
End Function

Private Sub checkStuff(ByVal lng病人ID As Long)
    Dim strsql As String
    Dim rstemp As Recordset
    
    On Error GoTo ErrHand
    strsql = " select count(A.id) 数量 from 药品收发记录 A,门诊费用记录 B where A.NO=B.NO and A.费用id=B.id and B.病人id=[2] and A.单据 in (24,25) and A.库房id=[1] and A.审核人 is null and (A.记录状态=1 or MOD(A.记录状态,3)=1)"
    Set rstemp = zldatabase.OpenSQLRecord(strsql, "checkStuff", mParams.lng药房ID, lng病人ID)
    
    If rstemp!数量 > 0 Then
        MsgBox "该病人还有未发的卫生材料，请注意发放！", vbInformation, gstrSysName
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub msg_upload(ByVal rsData As Recordset)
    '用于支付宝提示信息
    Dim strMsg As String
    Dim strsql As String
    Dim cmdTmp As New ADODB.Command
    Dim cmdPara As New ADODB.Parameter
    
    On Error GoTo ErrHand
    rsData.MoveFirst
    
    Set cmdTmp = New ADODB.Command
    Set cmdPara = cmdTmp.CreateParameter("病人ID", adVarNumeric, adParamInput, 18, rsData!病人ID)
    cmdTmp.Parameters.Append cmdPara
    Set cmdPara = cmdTmp.CreateParameter("NO", adVarChar, adParamInput, 100, rsData!NO)
    cmdTmp.Parameters.Append cmdPara
    Set cmdPara = cmdTmp.CreateParameter("说明", adLongVarChar, adParamOutput, 4000)
    cmdTmp.Parameters.Append cmdPara
    
    cmdTmp.ActiveConnection = gcnOracle
    cmdTmp.CommandType = adCmdStoredProc
    cmdTmp.CommandText = "Zl_MSG_PointOut"
    cmdTmp.Execute
    strMsg = Trim(zlStr.NVL(cmdTmp.Parameters("说明"), ""))
    
    If strMsg <> "" Then
        MsgBox strMsg, vbInformation, gstrSysName
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function RefreshDetail_Return(ByVal BillStyle As Integer, ByVal BillNo As String, ByVal str审核日期 As String, ByVal int可操作 As Integer, ByVal int门诊标志 As Integer, ByVal int记录性质 As Integer, Optional blnByNo As Boolean = False, Optional lng记录状态 As Long) As Boolean
    Dim IntStyle As Integer, intUnit As Integer
    Dim strSubSql As String
    Dim strName As String
    Dim blnMoved As Boolean
    Dim lng病人ID As Long
    Dim int主页id As Integer
    Dim strWeight As String
    
    Dim rstemp As New ADODB.Recordset
    Dim RecBill As New ADODB.Recordset
    '--读取单据内容--
    'BillStyle-单据类型;BIllNO-单据号
    '单位显示根据服务对象来（门诊：门诊单位；住院或住院门诊：住院单位；其它；售价单位）
'    On Error Resume Next
'    err = 0
    On Error GoTo errHandle
    RefreshDetail_Return = False
  
    mParams.strUnit = GetUnit(mSQLCondition.lng药房ID, BillStyle, BillNo, int门诊标志)
    Select Case mParams.strUnit
    Case "售价单位"
        strSubSql = "1"
    Case "门诊单位"
        strSubSql = "Decode(门诊包装,Null,1,0,1,门诊包装)"
    Case "住院单位"
        strSubSql = "Decode(住院包装,Null,1,0,1,住院包装)"
    Case "药库单位"
        strSubSql = "Decode(药库包装,Null,1,0,1,药库包装)"
    End Select
    Call Get单位串
    
    '得到药品名称串
    Select Case mParams.int药品名称显示
    Case 0  '药品编码与名称
        strName = "'['||C.编码||']'||" & IIf(gint药品名称显示 = 1, "NVL(E.名称,C.名称)", "C.名称") & " As 品名,"
    Case 1  '药品编码
        strName = "C.编码 As 品名,"
    Case 2  '药品名称
        strName = IIf(gint药品名称显示 = 1, "NVL(E.名称,C.名称)", "C.名称") & " As 品名,"
    End Select
    
    strName = strName & IIf(gint药品名称显示 <> 1, "NVL(E.名称,'')", "Decode(E.名称,Null,'',C.名称)") & " As 其它名, "
    
    '汇总显示单据内容
    '不可能存在一张处方同时在线与后备表中都存在
    blnMoved = Sys.IsMovedByNO("药品收发记录", BillNo, " 单据 = ", BillStyle)
    gstrSQL = " SELECT DISTINCT B.配药日期,B.审核日期,B.核查人,S.名称 As 药房,B.记录状态 状态,B.单据,B.处方类型,B.NO,H.序号,T.名称 科室,H.姓名,H.性别,H.年龄,H.标识号 住院号,H.床号,H.开单人,B.ID As 收发ID,B.药品ID,nvl(n.名称,'') 配方名称," & _
             " DECODE(B.批号,NULL,'',B.批号)||DECODE(B.批次,NULL,'',0,'','('||B.批次||')') 批号,DECODE(D.高危药品,null,0,0,0,1) 高危药品,to_char(B.效期,'yyyy-mm-dd') 效期,X.门诊号,X.病人类型,X.就诊卡号,decode(X.联系人电话,null,decode(X.手机号,null,X.家庭电话,X.手机号),X.联系人电话) 联系人电话," & _
             " NVL(B.批次,0) 批次,NVL(D.药房分批,0) 分批," & strName & _
             IIf(gint药品名称显示 = 1, "NVL(E.名称,C.名称)", "C.名称") & " As 药品名称, " & _
             " DECODE(C.规格,NULL,B.产地,DECODE(B.产地,NULL,C.规格,C.规格||'|'||B.产地)) 规格,Nvl(b.产地, Nvl(c.产地, '')) 产地, b.原产地," & str单位串 & "," & _
             " NVL(B.付数,1) 付数,NVL(H.付数,1) 原始付数," & _
             " B.已退数量/" & strSubSql & " 已退数量, B.已退数量 小单位已退数," & _
             " B.已发数量/" & strSubSql & " 准退数,B.已发数量 小单位准退数,B.已发数量 实际数量,B.实际数量 小单位数量," & _
             " B.单量,B.用法,B.频次,B.填制人,B.填制日期,H.操作员姓名,B.配药人,B.审核人 发药人,I.计算单位," & _
             " round(B.零售金额," & mintMoneyDigit & " ) 零售金额,round(Nvl(B.付数, 1) * B.实际数量 / (Nvl(H.付数, 1) * H.数次) * Nvl(H.实收金额,0)," & mintMoneyDigit & " ) 实收金额,H.费别,I.名称 As 药名 ," & _
             " P.毒理分类,Nvl(P.抗生素,0) 抗生素,Nvl(P.是否皮试,0) As 是否皮试, H.门诊标志, H.记录性质,B.库房id As 药房id,Nvl(M.相关ID,0) As 相关ID,M.开嘱医生,M.频率间隔,M.超量说明,M.间隔单位,Nvl(M.开嘱时间,H.登记时间) As 开嘱时间,M.医嘱期效 医嘱标志,M.开始执行时间 开始时间,M.执行终止时间 结束时间,M.频率次数,Nvl(Nvl(M.相关ID,M.id),0) As 医嘱id,Nvl(M.主页id,0) as 主页id," & _
             " M.皮试结果,M.禁忌药品说明,D.药名ID, f.名称 As 险类,H.结论 As 中药形态,C.规格 As 药品规格,M.医生嘱托,decode(m.用药目的,1,'预防',2,'治疗',3,'预防和治疗','') 用药目的,m.用药理由,D.剂量系数,"
             
    If int可操作 = 1 Then  '输入的情况考虑进去
        gstrSQL = gstrSQL & " B.已发数量*D.剂量系数 重量,Decode(Sign(Nvl(K.库存数量, 0) - Nvl(L.下限, 0)), -1, 0, 1) 库存下限,Z.名称 As 英文名,Nvl(H.病人ID,0) As 病人ID,Nvl(x.在院, 0) As 在院 FROM "
        gstrSQL = gstrSQL & "   (SELECT A.ID,A.NO,A.单据,A.药品ID,A.费用ID,A.批次,A.批号,A.产地,A.原产地,A.效期," & _
                 "          NVL(A.付数,1) 付数,A.实际数量,NVL(A.付数,1)*A.实际数量-B.已发数量 已退数量,B.已发数量,A.记录状态," & _
                 "          A.零售价,A.零售金额,A.单量,A.频次,A.用法,A.摘要,A.配药日期,A.填制人,A.填制日期,A.核查人,A.配药人,A.审核人,A.审核日期,A.对方部门ID,A.库房ID, A.处方类型 " & _
                 "      FROM" & _
                 "          (SELECT A.ID,A.NO,A.单据,A.药品ID,A.序号,A.费用ID,A.批次,A.批号,A.产地,A.原产地,A.效期,A.付数,A.实际数量,A.记录状态,A.零售价,A.零售金额,A.单量,A.频次,A.用法,A.摘要,A.配药日期,A.填制人,A.配药人,A.填制日期,A.核查人,A.审核人,A.审核日期,A.对方部门ID,A.库房ID, Nvl(A.注册证号, 0) As 处方类型 " & _
                 "          FROM 药品收发记录 A" & _
                 "          WHERE A.审核人 IS NOT NULL AND (A.记录状态=1 OR MOD(A.记录状态,3)=0)" & _
                 "          AND A.库房ID+0=[1] "
        If blnByNo = False Then
            gstrSQL = gstrSQL & " AND A.审核日期 Between [2] And [3] "
        Else
            gstrSQL = gstrSQL & " And A.单据=[4] And A.NO=[5] "
        End If
        
        gstrSQL = gstrSQL & "          ) A," & _
                 "          (SELECT A.NO,A.单据,A.药品ID,A.序号,SUM(NVL(A.付数,1)*A.实际数量) 已发数量" & _
                 "          FROM 药品收发记录 A" & _
                 "          WHERE A.审核人 IS NOT NULL" & _
                 "          AND A.库房ID+0=[1] "
        
        If blnByNo = False Then
            gstrSQL = gstrSQL & " AND A.审核日期 Between [2] And [3] "
        Else
            gstrSQL = gstrSQL & " And A.单据=[4] And A.NO=[5] "
        End If
                
        gstrSQL = gstrSQL & "          GROUP BY A.NO,A.单据,A.药品ID,A.序号) B" & _
                 "      WHERE A.NO = B.NO AND A.单据 = B.单据 AND A.药品ID+0 = B.药品ID AND A.序号 = B.序号 " & _
                 "      )"
    Else
        gstrSQL = gstrSQL & " B.实际数量*D.剂量系数 重量,Decode(Sign(Nvl(K.库存数量, 0) - Nvl(L.下限, 0)), -1, 0, 1) 库存下限,Z.名称 As 英文名,Nvl(H.病人ID,0) As 病人ID,Nvl(x.在院, 0) As 在院 FROM "
        gstrSQL = gstrSQL & "(Select 0 已发数量,0 已退数量,0 准退数量,Nvl(A.注册证号, 0) As 处方类型,A.* From 药品收发记录 A where FLOOR(A.记录状态/3)+1=[10])"
    End If
    gstrSQL = gstrSQL & _
            "       B,药品规格 D,药品特性 P,收费项目目录 C,收费项目别名 E,门诊费用记录 H,病人医嘱记录 M,病人医嘱记录 G,部门表 S,部门表 T,诊疗项目目录 I,诊疗项目别名 Z ,保险支付大类 F,诊疗项目目录 N,病人信息 X, " & _
            "(Select b.库房id, b.药品id, Nvl(Sum(b.实际数量), 0) 库存数量 " & _
            " From 药品收发记录 A, 药品库存 B " & _
            " Where a.药品id = b.药品id And b.性质 = 1 And b.库房id + 0 = [1] And a.单据 = [4] And a.No = [5] " & _
            " Group By b.库房id, b.药品id) K, 药品储备限额 L " & _
            " Where H.开单部门ID=T.ID(+) And B.药品ID=D.药品ID And D.药名ID=P.药名ID And C.ID=D.药品ID And H.医嘱序号=M.ID(+) And Nvl(M.相关id, M.ID) = G.ID(+) and G.配方id=N.id(+) " & _
            " And D.药品ID=E.收费细目ID(+) and E.性质(+)=3 And D.药名id = Z.诊疗项目id(+) And Z.性质(+) = 2 And h.保险大类id = f.Id(+) " & _
            " And S.ID=B.库房ID And B.费用ID=H.ID And B.NO=[5] And B.单据=[4] And B.库房ID+0=[1] and H.病人id=X.病人id(+) "
    
    If mSQLCondition.str填制人 <> "" Then gstrSQL = gstrSQL & " And B.填制人=[7] "
    If mSQLCondition.str审核人 <> "" Then gstrSQL = gstrSQL & " And B.审核人=[8] "
    If mSQLCondition.lng药品id > 0 Then gstrSQL = gstrSQL & " And B.药品ID=[9] "
    
    If IsDate(str审核日期) Then
             gstrSQL = gstrSQL & " And B.审核日期=To_Date([6],'yyyy-MM-dd hh24:mi:ss')"
    End If
    gstrSQL = gstrSQL & " And B.审核人 Is Not Null And D.药名id=I.id " & _
                        " And B.药品id = L.药品id(+) And Nvl(B.库房id, 24) = L.库房id(+) And" & _
                        " D.药名id = I.ID And Nvl(B.库房id, 24) + 0 = K.库房id(+) And B.药品id = K.药品id(+) "
    
    gstrSQL = gstrSQL & " Order by H.序号,B.药品ID,Nvl(B.批次,0)"
    
    If int记录性质 = 1 Or (int记录性质 = 2 And (int门诊标志 = 1 Or int门诊标志 = 4)) Then
        '门诊
        gstrSQL = Replace(gstrSQL, "H.床号", "'' 床号")
    Else
        '住院
        gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
    End If
    
    '如果数据转出，则直接从后备表中提取数据
    If blnMoved Then
        gstrSQL = Replace(gstrSQL, "药品收发记录", "H药品收发记录")
        gstrSQL = Replace(gstrSQL, "门诊费用记录", "H门诊费用记录")
        gstrSQL = Replace(gstrSQL, "住院费用记录", "H住院费用记录")
    End If
    
    Set RecBill = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
            mSQLCondition.lng药房ID, _
            mSQLCondition.date开始日期, _
            mSQLCondition.date结束日期, _
            BillStyle, _
            BillNo, _
            str审核日期, _
            mSQLCondition.str填制人, _
            mSQLCondition.str审核人, _
            mSQLCondition.lng药品id, _
            lng记录状态)
    
    If err <> 0 Then
        MsgBox "读取处方时，发生不可预知的错误！", vbInformation, gstrSysName
        Exit Function
    End If
    
    If Not RecBill.EOF Then
        If NVL(RecBill!病人ID) <> 0 Then
            lng病人ID = RecBill!病人ID
            int主页id = NVL(RecBill!主页id)
            If int记录性质 = 1 Or (int记录性质 = 2 And (int门诊标志 = 1 Or int门诊标志 = 4)) Then
                '门诊
                gstrSQL = "select A.id,B.记录内容 体重 from 病人护理记录 A,病人护理内容 B where A.id=B.记录id and B.项目名称='体重' and 病人id=[1] order by A.Id desc"
            Else
                '住院
                 gstrSQL = "select 体重 from 病案主页 where 病人id=[1] and 主页id=[2]"
                 
            End If
            Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng病人ID, int主页id)
            
            If Not rstemp.EOF Then
                strWeight = NVL(rstemp!体重)
            End If
        End If
    End If

    mfrmDetail.RefreshList RecBill, strWeight, int可操作
    mfrmRecipe.RefreshRecipe RecBill, strWeight, int可操作
    
    RefreshDetail_Return = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SendBatchRecipe(ByVal rsData As ADODB.Recordset) As Boolean
    Dim n As Integer
    Dim lngRow As Long, lng药品id As Long, LngID As Long, lng批次 As Long, lng分批 As Long
    Dim rsSendRecipeByNo As ADODB.Recordset
    Dim rsSendRecipeDetail As ADODB.Recordset
    Dim int门诊 As Integer
    Dim strNO串 As String
    Dim arrSql As Variant
    Dim blnInTrans As Boolean
    Dim i As Integer
    Dim str签名记录 As String
    Dim date发药时间 As Date
    Dim strNo As String  '用于传给发药机
    Dim strReturn As String, strMessage As String
    Dim strReserve As String
    
    On Error GoTo ErrHand
    
    date发药时间 = Sys.Currentdate
    
    Set rsSendRecipeByNo = New ADODB.Recordset
    With rsSendRecipeByNo
        If .State = 1 Then .Close
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "单据", adDouble, 18, adFldIsNullable
        .Fields.Append "配药人", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "填制人", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "药房ID", adDouble, 18, adFldIsNullable
        .Fields.Append "记录性质", adDouble, 1, adFldIsNullable
        .Fields.Append "门诊标志", adDouble, 1, adFldIsNullable
        .Fields.Append "处方类型", adDouble, 1, adFldIsNullable
        .Fields.Append "收费类别", adDouble, 1, adFldIsNullable
        .Fields.Append "姓名", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "核查人", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "填制日期", adDate, , adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    Set rsSendRecipeDetail = New ADODB.Recordset
    With rsSendRecipeDetail
        If .State = 1 Then .Close
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "收发ID", adDouble, 18, adFldIsNullable
        .Fields.Append "药品ID", adDouble, 18, adFldIsNullable
        .Fields.Append "批次", adDouble, 18, adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    rsData.Filter = "标志=1"
    rsData.Sort = "单据,NO"
    Do While Not rsData.EOF
        With rsSendRecipeByNo
            If strNO串 <> rsData!单据 & "|" & rsData!NO Then
                strNO串 = rsData!单据 & "|" & rsData!NO
                .AddNew
                !药房ID = rsData!药房ID
                !NO = rsData!NO
                !单据 = rsData!单据
                !配药人 = IIf(tbcDetail.Selected.index = 1, mfrmRecipe.Get配药人, mfrmDetail.Get配药人)
                !填制人 = IIf(tbcDetail.Selected.index = 1, mfrmRecipe.Get开单医生, mfrmDetail.Get开单医生)
                !记录性质 = rsData!记录性质
                !门诊标志 = rsData!门诊标志
                !处方类型 = rsData!处方类型
                !收费类别 = rsData!收费类别
                !姓名 = IIf(IsNull(rsData!姓名), "", rsData!姓名)
                !核查人 = IIf(tbcDetail.Selected.index = 1, mfrmRecipe.Get核查人, mfrmDetail.Get核查人)
                !填制日期 = rsData!填制日期
                .Update
            End If
        End With
        rsData.MoveNext
    Loop
    
    rsData.Filter = "标志=1"
    rsData.Sort = "单据,NO,药品ID"
    Do While Not rsData.EOF
        With rsSendRecipeDetail
            .AddNew
            !NO = rsData!NO
            !收发ID = rsData!收发ID
            !药品ID = rsData!药品ID
            !批次 = rsData!批次
            .Update
        End With
        
        rsData.MoveNext
    Loop
    
    arrSql = Array()
    
    mstrPrintRecipe = ""
    
    '按处方号排序后批量发药
    rsSendRecipeByNo.Sort = "NO"
    rsSendRecipeByNo.MoveFirst
    For n = 1 To rsSendRecipeByNo.RecordCount
        '先检查或执行预调价
        Call AutoAdjustPrice_ByNO(rsSendRecipeByNo!单据, rsSendRecipeByNo!NO)

        rsSendRecipeDetail.Filter = "NO='" & rsSendRecipeByNo!NO & "'"
        rsSendRecipeDetail.MoveFirst
        
        If Val(rsSendRecipeByNo!记录性质) = 1 Or (Val(rsSendRecipeByNo!记录性质) = 2 And (Val(rsSendRecipeByNo!门诊标志) = 1 Or Val(rsSendRecipeByNo!门诊标志) = 4)) Then
            int门诊 = 1
        Else
            int门诊 = 2
        End If
        
        If mPrives.bln发其它药房的处方 = True And mParams.lng药房ID <> Val(rsSendRecipeByNo!药房ID) Then
            gstrSQL = "Zl_药品收发记录_更改库房("
            '现库房ID
            gstrSQL = gstrSQL & mParams.lng药房ID
            '单据
            gstrSQL = gstrSQL & "," & rsSendRecipeByNo!单据
            'NO
            gstrSQL = gstrSQL & ",'" & rsSendRecipeByNo!NO & "'"
            '原库房ID
            gstrSQL = gstrSQL & "," & Val(rsSendRecipeByNo!药房ID)
            '门诊
            gstrSQL = gstrSQL & "," & int门诊
            '填制日期
            gstrSQL = gstrSQL & ",to_date('" & rsSendRecipeByNo!填制日期 & "','yyyy-MM-dd')"
            gstrSQL = gstrSQL & ")"
            
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = gstrSQL
            
'            Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-更改库房")
        End If
        
        If mParams.IntCheckStock = 2 Then
            For lngRow = 1 To rsSendRecipeDetail.RecordCount
                gstrSQL = "zl_药品收发记录_更新批次("
                '收发ID
                gstrSQL = gstrSQL & rsSendRecipeDetail!收发ID
                '药品ID
                gstrSQL = gstrSQL & "," & rsSendRecipeDetail!药品ID
                '批次
                gstrSQL = gstrSQL & "," & rsSendRecipeDetail!批次
                gstrSQL = gstrSQL & ")"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
'                Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-更新批次")
                
                rsSendRecipeDetail.MoveNext
            Next
        End If
        
        gstrSQL = "zl_药品收发记录_处方发药("
        '库房ID
        gstrSQL = gstrSQL & mParams.lng药房ID
        '单据
        gstrSQL = gstrSQL & "," & rsSendRecipeByNo!单据
        'NO
        gstrSQL = gstrSQL & ",'" & rsSendRecipeByNo!NO & "'"
        '发药人(审核人)
        gstrSQL = gstrSQL & ",'" & mstr操作员 & "'"
        '配药人(必须经过配药过程时，则配药人不填)
        gstrSQL = gstrSQL & "," & IIf(mParams.blnMustDosageProcess = True, "Null", IIf(rsSendRecipeByNo!配药人 = "", "NULL", "'" & rsSendRecipeByNo!配药人 & "'")) & ""
        '校验人（开单医生）
        gstrSQL = gstrSQL & "," & IIf(rsSendRecipeByNo!填制人 = "", "NULL", "'" & rsSendRecipeByNo!填制人 & "'") & ""
        '发药方式
        gstrSQL = gstrSQL & ",1"
        '发药时间
        gstrSQL = gstrSQL & ",to_date('" & date发药时间 & "','yyyy-MM-dd hh24:mi:ss') "
        '操作员编码
        gstrSQL = gstrSQL & ",'" & gstrUserCode & "'"
        '操作员名称
        gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
        '金额保留位数
        gstrSQL = gstrSQL & "," & mParams.int金额保留位数
        '自动审核记账单
        gstrSQL = gstrSQL & "," & IIf(mParams.bln审核划价单, 1, 0)
        '是否门诊
        gstrSQL = gstrSQL & "," & int门诊
        '核查人
        gstrSQL = gstrSQL & ",'" & rsSendRecipeByNo!核查人 & "'"
        '病人是否实际取药
        gstrSQL = gstrSQL & "," & IIf(mbln未取药发药, 1, "Null")
        
        gstrSQL = gstrSQL & ")"
        
        ReDim Preserve arrSql(UBound(arrSql) + 1)
        arrSql(UBound(arrSql)) = gstrSQL
            
        '记录该处方号及单据类型
        mstrBill = rsSendRecipeByNo!NO & "|" & rsSendRecipeByNo!单据
        mstrPrintRecipe = IIf(mstrPrintRecipe = "", "", mstrPrintRecipe & "|") & rsSendRecipeByNo!NO & "," & rsSendRecipeByNo!单据 & "," & rsSendRecipeByNo!记录性质 & "," & rsSendRecipeByNo!门诊标志 & "," & rsSendRecipeByNo!处方类型 & "," & rsSendRecipeByNo!收费类别
        mfrmList.mstrLastName = rsSendRecipeByNo!姓名

        strNo = strNo & rsSendRecipeByNo!单据 & "," & rsSendRecipeByNo!NO & "|"
        rsSendRecipeByNo.MoveNext
    Next
    
    mstr操作员 = ""
    mstr配药人 = ""
    
'    '先处理发药机事务，发药系统未准备好则显示接口返回信息，操作员可以选择是否发药
'    If Not gobjPackerMZ Is Nothing And strNo <> "" Then
'        If gobjPackerMZ.HisUpload(mlngMode, 2, strNo, mParams.lng药房ID) = False Then
'            If MsgBox("自动发药系统未准备好，是否继续发药？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
'                Exit Function
'            End If
'        End If
'    End If

    '先处理发药机事务，发药系统未准备好则显示接口返回信息，操作员可以选择是否发药
    If TypeName(mobjDrugMAC) = "clsDrugPacker" Then
        If mblnPackerConnect And strNo <> "" And mblnLoadDrug Then
            If mblnCompatible = False Then
                '不是最新接口部件
                If mobjDrugMAC.DYEY_MZ_TransRecipeList(mstrOpr, UserInfo.用户编码, UserInfo.用户姓名, mParams.lng药房ID, Mid(strNo, 1, Len(strNo) - 1), strReturn) = False Then
                    If MsgBox("自动发药系统未准备好，是否继续发药？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                End If
            Else
                If mobjDrugMAC.DYEY_MZ_TransRecipeList(mstrOpr, UserInfo.用户编码, UserInfo.用户姓名, mParams.lng药房ID, Mid(strNo, 1, Len(strNo) - 1), strReturn, IIf(mintAutoSendFlow = 0, mSendOper.StartSend, mSendOper.EndSend)) = False Then
                    If MsgBox("自动发药系统未准备好，是否继续发药？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                End If
            End If
        End If
    ElseIf TypeName(mobjDrugMAC) = "clsDrugMachine" Then
        If mblnPackerConnect Then
            If Right(strNo, 1) = "|" Then strNo = Left(strNo, Len(strNo) - 1)
            mobjDrugMAC.Operation gstrDbUser, Val("23-完成发药"), "1|" & Replace(strNo, "|", ";"), strMessage
'           If strMessage <> "" Then MsgBox strMessage, vbInformation, gstrSysName
        End If
    End If
    
    '调用发药前的外挂接口
    err.Clear
    If Not mobjPlugIn Is Nothing Then
        If Right(strNo, 1) = "|" Then strNo = Left(strNo, Len(strNo) - 1)
        On Error Resume Next
        If mobjPlugIn.DrugBeforeSendByRecipe(mParams.lng药房ID, strNo, strReserve) = False Then
            If err.Number <> 0 Then
                err.Clear: On Error GoTo 0
            Else
                Exit Function
            End If
        End If
        err.Clear: On Error GoTo 0
    End If
    
    On Error GoTo ErrHand
    
    '集中处理发药事务
    gcnOracle.BeginTrans
    blnInTrans = True
    
    '单独处理电子签名
    '如果已启用了电子签名，则需要对发药人进行电子签名处理
    If gblnESign处方发药 = True And gblnESignUserStoped = False Then
        rsSendRecipeByNo.MoveFirst
        For n = 1 To rsSendRecipeByNo.RecordCount
            str签名记录 = ""
            If GetSignatureRecored(EsignTache.send, rsSendRecipeByNo!单据, rsSendRecipeByNo!NO, mParams.lng药房ID, str签名记录, 0, date发药时间, IIf(mstr操作员 = "", gstrUserName, mstr操作员)) = False Then
                gcnOracle.RollbackTrans
                blnInTrans = False
                Exit Function
            End If
            
            If str签名记录 <> "" Then
                gstrSQL = "Zl_药品签名记录_Insert(" & str签名记录 & ")"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            Else
                gcnOracle.RollbackTrans
                blnInTrans = False
                MsgBox "对发药人电子签名失败！", vbInformation, gstrSysName
                Exit Function
            End If
            
            rsSendRecipeByNo.MoveNext
        Next
    End If
    
    For i = 0 To UBound(arrSql)
        Call zldatabase.ExecuteProcedure(CStr(arrSql(i)), Me.Caption & "-发药")
    Next
        
    gcnOracle.CommitTrans
    blnInTrans = False
    SendBatchRecipe = True
    
    '调用发药后的外挂接口
    If Not mobjPlugIn Is Nothing Then
        If Right(strNo, 1) = "|" Then strNo = Left(strNo, Len(strNo) - 1)
        On Error Resume Next
        mobjPlugIn.DrugSendByRecipe mParams.lng药房ID, strNo, date发药时间, strReserve
        err.Clear: On Error GoTo 0
    End If
    
    Exit Function
ErrHand:
    SendBatchRecipe = False
    If blnInTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckNotAudited(ByRef rsData As ADODB.Recordset) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim bln销帐申请 As Boolean
    Dim bln允许发送 As Boolean
    
    On Error GoTo errHandle
    
    Call InitApplyforcredit
    
    CheckNotAudited = True
    bln销帐申请 = True
    
    '检测当前药房是否为住院药房，不是则退出此项检查
    gstrSQL = "Select *" & vbNewLine & _
            "From 部门表 A, 部门性质说明 B" & vbNewLine & _
            "Where a.Id = b.部门id And a.Id = [1] And (b.工作性质 Like '%药房' Or b.工作性质 Like '%药库') And b.服务对象 In (2, 3)"

    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "检测当前药房是否为住院药房", mParams.lng药房ID)
    If rsTmp.EOF Then Exit Function
    

    gstrSQL = "Select c.数量 As 销帐申请数量, c.费用id" & vbNewLine & _
            "From 药品收发记录 A, 住院费用记录 B, 病人费用销帐 C" & vbNewLine & _
            "Where a.费用id = b.Id And b.Id = c.费用id And a.药品id = c.收费细目id And a.Id = [1] And c.状态 = 0"

    
    With rsData
        rsData.Filter = "标志=1"
        rsData.Sort = "单据,NO"
    
        Do While Not .EOF
            Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "检查是否存在销帐申请未审核的单据", rsData!收发ID)

            If rsTmp.RecordCount > 0 Then
                bln销帐申请 = False

                With mrsApplyforcredit
                    .AddNew
                    
                    !标志 = 1
                    !NO = rsData!NO
                    !药品名称 = rsData!药名
                    !批号 = rsData!批号
                    !数量 = zlStr.FormatEx(rsData!数量 / rsData!包装, 4) & rsData!单位
                    !销帐申请数量 = zlStr.FormatEx(rsTmp!销帐申请数量 / rsData!包装, 4) & rsData!单位
                    !姓名 = rsData!姓名
                    !性别 = rsData!性别
                    !年龄 = rsData!年龄
                    !费用ID = rsTmp!费用ID
                    !收发ID = rsData!收发ID
                End With

            End If

            .MoveNext
        Loop
    End With

    '对含有销帐申请的单据进行处理
    If bln销帐申请 = False Then
        Call frm部门发药销帐申请清单.ShowCard(Me, mrsApplyforcredit, bln允许发送, 1)

        '由子窗体返回用户是否继续执行操作，若【取消】则禁止继续发送
        CheckNotAudited = bln允许发送
        If CheckNotAudited = False Then Exit Function
        
        '修正取消发送的单据的执行状态
        mrsApplyforcredit.Filter = "标志 = 0"
        
        If mrsApplyforcredit.RecordCount > 0 Then
            If mint发药方式 = 1 Then
                CheckNotAudited = False
                Exit Function
            End If
            
            Do While Not mrsApplyforcredit.EOF
                rsData.Filter = "NO = '" & mrsApplyforcredit!NO & "'"
                If rsData.RecordCount > 0 Then
                    Do While Not rsData.EOF
                        rsData!标志 = 0
                        rsData.Update
                        rsData.MoveNext
                    Loop
                End If
                mrsApplyforcredit.MoveNext
            Loop
        End If

        rsData.Filter = ""
    End If
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckBatchRecipe(ByVal rsData As ADODB.Recordset) As Boolean
    Dim n As Integer
    Dim rstemp As ADODB.Recordset
    Dim blnFirst As Boolean
    Dim lngRow As Long, lng药品id As Long, LngID As Long, lng批次 As Long, lng分批 As Long
    Dim blnBatchSend As Boolean
    Dim i As Integer
    Dim dblSumMoney As Double
    Dim strRecipeString As String
    Dim rsCheck As ADODB.Recordset
    Dim arrRecipe
    Dim intCount As Integer
    Dim int当前单据 As Integer
    Dim str当前NO As String
    Dim str收费类型 As String
    Dim str收费细目id As String
    Dim strTemp As String
    Dim str核查人 As String
    
    On Error GoTo ErrHand
    If rsData!结算模式 = 1 Then
        If gobjCharge Is Nothing Then
            Set gobjCharge = CreateObject("zl9OutExse.clsOutExse")
            If gobjCharge Is Nothing Then Exit Function
        End If
        
        If Not gobjCharge Is Nothing Then
            strTemp = BillHaveHerial(rsData!NO, rsData!单据, 1, str收费细目id, str收费类型)
            If str收费类型 <> "" Then
                If Not gobjCharge.zlCheckExcuteItemValied(Me, gcnOracle, UserInfo.用户姓名, glngSys, mlngMode, rsData!病人ID, str收费类型, rsData!NO, str收费细目id) Then
                    CheckBatchRecipe = False
                    Exit Function
                End If
            End If
        End If
    End If
       
    '检查病人费用余额
    Set rsCheck = New ADODB.Recordset
    With rsCheck
        If .State = 1 Then .Close
        .Fields.Append "单据", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "病人ID", adDouble, 18, adFldIsNullable
        .Fields.Append "实收金额", adDouble, 18, adFldIsNullable
        .Fields.Append "已收费", adDouble, 1, adFldIsNullable
        .Fields.Append "记录性质", adDouble, 1, adFldIsNullable
        .Fields.Append "门诊标志", adDouble, 1, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    strRecipeString = mfrmList.GetCurrentBatchRecipe
    
    arrRecipe = Split(strRecipeString, "|")
    intCount = UBound(arrRecipe)
    
    For n = 0 To intCount
        rsCheck.AddNew
        rsCheck!单据 = Val(Split(arrRecipe(n), ",")(0))
        rsCheck!NO = Split(arrRecipe(n), ",")(1)
        rsCheck!病人ID = Val(Split(arrRecipe(n), ",")(2))
        rsCheck!实收金额 = Val(Split(arrRecipe(n), ",")(3))
        rsCheck!已收费 = Val(Split(arrRecipe(n), ",")(4))
        rsCheck!记录性质 = Val(Split(arrRecipe(n), ",")(5))
        rsCheck!门诊标志 = Val(Split(arrRecipe(n), ",")(6))
        rsCheck.Update
    Next
    If Not CheckSendBillMoney(rsCheck) Then Exit Function
    
    '检查药品存储库房
    If CheckDrugStock(rsData) = False Then Exit Function
    
    '检查[住院单据]是否存在销帐申请未审核的单据
    If CheckNotAudited(rsData) = False Then Exit Function
    
    rsData.Filter = "标志=1"
    rsData.Sort = "单据,NO"
    
    '检查当前病人其他窗口或其他药房的未发药单据
    Call CheckOtherUndeliveredDocuments(rsData!病人ID)
    
    '检查[核查人]是否为空
    str核查人 = IIf(tbcDetail.Selected.index = 1, mfrmRecipe.Get核查人, mfrmDetail.Get核查人)
    If str核查人 = "" Then
        MsgBox "核查人为空，不能执行发药操作！", vbInformation, gstrSysName
        Exit Function
    End If
    
    Do While Not rsData.EOF
        '检测是否允许
        If CheckBill(rsData!药房ID, 3, rsData!单据, rsData!NO, rsData!记录性质, rsData!门诊标志) <> 0 Then Exit Function
        
        '检查是否收费(发药处理)
        gstrSQL = " Select Decode(配药人,Null,'','部门发药','',配药人) 配药人,已收费 From 未发药品记录" & _
                 " Where No=[1] And (库房ID=[3] Or 库房ID Is NULL) And 单据=[2]"
        Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, rsData!NO, Val(rsData!单据), Val(rsData!药房ID))
        
        With rstemp
            If .EOF Then
                MsgBox "该处方已经被其他操作员处理！", vbInformation, gstrSysName
                CheckBatchRecipe = False
                Exit Function
            End If
            
            If mParams.blnMustDosageProcess = True Then
                If IsNull(!配药人) Then
                    MsgBox "该处方还未配药，不能执行发药操作！", vbInformation, gstrSysName
                    Exit Function
                End If
                If Trim(!配药人) = "" Then
                    MsgBox "该处方还未配药，不能执行发药操作！", vbInformation, gstrSysName
                    Exit Function
                End If
            End If

            mstr配药人 = zlStr.NVL(!配药人)
            
            If mParams.bln发药前收费或审核 = False Then
                '未收费的划价单
                If rsData!单据 = 8 And !已收费 = 0 And mParams.bln允许未收费处方发药 = False Then
                    MsgBox "该处方还未收费，不能执行发药操作！", vbInformation, gstrSysName
                    Exit Function
                End If
            
                '未审核的记账划价单
                If rsData!单据 = 9 And !已收费 = 0 And mParams.bln允许未审核处方发药 = False Then
                    MsgBox "该处方还未审核，不能执行发药操作！", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            
            If Not IsReceiptBalance_Charge(0, mstrPrivs, rsData!单据, rsData!NO, rsData!序号, rsData!记录性质, rsData!门诊标志) Then Exit Function
            If Not IsOutPatient(mstrPrivs, rsData!单据, rsData!NO, rsData!记录性质, rsData!门诊标志) Then Exit Function
            
            mrsList.Filter = "单据=" & rsData!单据 & " And NO='" & rsData!NO & "' "
            If Not mrsList.EOF Then dblSumMoney = Val(mrsList!金额)
            
            If Not CheckBillControl(mcondition.intListType + 1, rsData!单据, rsData!NO, dblSumMoney) Then Exit Function
            
            '校验发药人
            If mParams.int校验发药人 = 1 And Not blnFirst Then
                mstr操作员 = zldatabase.UserIdentify(Me, "校验发药人", glngSys, 1341, "发药")
                blnFirst = True
            Else
                mstr操作员 = gstrUserName
            End If
            If mstr操作员 = "" Then Exit Function
        End With
        
        rsData.MoveNext
    Loop
        
    '其他检查
    rsData.Sort = "单据,NO"
    rsData.MoveFirst
    Do While Not rsData.EOF
        If int当前单据 <> rsData!单据 And str当前NO <> rsData!NO Then
            int当前单据 = rsData!单据
            str当前NO = rsData!NO
                                    
            '零差价管理
            If CheckPriceAdjustByNO(Val(rsData!单据), Val(rsData!药房ID), rsData!NO) = False Then
                Exit Function
            End If
            
            '毒麻药品提示
            If Not CheckSpec(rsData!药房ID, rsData!NO, rsData!单据) Then Exit Function
            
            If mstr毒麻类提示 <> "" Then
                If MsgBox("单号为[" & rsData!NO & "]" & "的处方中含有以下毒麻类药品，确定发药吗？" & mstr毒麻类提示, vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
            End If
            
            '库存检查
            If Not CheckStock(rsData!药房ID, rsData!NO, rsData!单据) Then Exit Function
        End If
        
        rsData.MoveNext
    Loop
    
    '发药时病人消费卡确认只支持一个病人模式
    If CheckPati(rsData) = False Then Exit Function
    
    CheckBatchRecipe = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    CheckBatchRecipe = False
End Function

Private Function CheckStock(ByVal lngNO药房id As Long, ByVal strNo As String, ByVal IntBillStyle As Integer) As Boolean
    Dim RecCheckStock As New ADODB.Recordset, RecBillData As New ADODB.Recordset
    Dim dblStock As Double, intCheck As Integer
    Dim dblUsableStock As Double
    '--检查库存--
    '0-不检查;1-检查,不足提醒;2-检查,不足禁止
'    On Error Resume Next
'    err = 0
    On Error GoTo errHandle
    CheckStock = False
    intCheck = mParams.IntCheckStock
    
    '逐行检查
    If intCheck <> 0 Then
        gstrSQL = " SELECT A.药品ID,SUM(NVL(A.实际数量,0)*NVL(A.付数,1)) 数量," & _
                " '['||C.编码||']'||" & IIf(gint药品名称显示 = 1, "NVL(L.名称,C.名称)", "C.名称") & " 品名,NVL(A.批次,0) 批次, Nvl(A.批号,'') 批号 " & _
                " FROM 药品收发记录 A,药品规格 B,收费项目目录 C,收费项目别名 L " & _
                " WHERE A.药品ID=B.药品ID AND B.药品ID=C.ID" & _
                " AND B.药品ID=L.收费细目ID(+) AND L.性质(+)=3 AND L.码类(+)=1 " & _
                " AND A.审核人 IS NULL AND MOD(A.记录状态,3)=1 AND NVL(A.摘要,'小宝')<>'拒发'" & _
                " AND A.NO=[1] AND A.单据=[2] AND (A.库房ID+0=[3] OR A.库房ID IS NULL) " & _
                " GROUP BY A.药品ID,'['||C.编码||']'||" & IIf(gint药品名称显示 = 1, "NVL(L.名称,C.名称)", "C.名称") & ",批次 ,A.批号"
        Set RecBillData = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, IntBillStyle, lngNO药房id)
        
        With RecBillData
            Do While Not .EOF
                gstrSQL = " Select nvl(可用数量,0) AS 可用数量, nvl(实际数量,0) AS 实际数量 " & _
                         " From 药品库存 " & _
                         " Where 库房ID+0=[1] And 药品ID=[2] " & _
                         " And 性质=1 And Nvl(批次,0)=[3]"
                Set RecCheckStock = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mParams.lng药房ID, CLng(RecBillData!药品ID), CLng(RecBillData!批次))
                
                With RecCheckStock
                    If .EOF Then
                        dblStock = 0
                        dblUsableStock = 0
                    Else
                        dblStock = !实际数量
                        dblUsableStock = !可用数量
                    End If
                    
                    '如果是代发其他药房处方(处方药房和当前药房不一样时)，则既要检查实际数量，也要检查可用数量
                    If dblStock < RecBillData!数量 Or (lngNO药房id <> mParams.lng药房ID And dblUsableStock < RecBillData!数量) Then
                        If RecBillData!批次 > 0 And NVL(RecBillData!批号, "") <> "" Then
                            Select Case intCheck
                                Case 1
                                    If MsgBox(RecBillData!品名 & "批号为[" & RecBillData!批号 & "]的库存数不够，是否继续发药？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                                Case 2
                                    MsgBox RecBillData!品名 & "批号为[" & RecBillData!批号 & "]的库存数不够，不能继续发药！", vbInformation, gstrSysName: Exit Function
                            End Select
                        Else
                            Select Case intCheck
                                Case 1
                                    If MsgBox(RecBillData!品名 & "的库存数不够，是否继续发药？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                                Case 2
                                    MsgBox RecBillData!品名 & "的库存数不够，不能继续发药！", vbInformation, gstrSysName: Exit Function
                            End Select
                        End If
                    End If
                End With
                .MoveNext
            Loop
        End With
    End If
    
    If err <> 0 Then
        MsgBox "检查库存时，发生不可预知的错误！", vbInformation, gstrSysName
        Exit Function
    End If
    CheckStock = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function CheckSpec(ByVal lngNO药房id As Long, ByVal strNo As String, ByVal IntBillStyle As Integer) As Boolean
    Dim strNote As String
    Dim rstemp As New ADODB.Recordset
    
    mstr毒麻类提示 = ""
    
    '对毒麻类药品进行检查
    On Error GoTo errHandle
    gstrSQL = " SELECT Distinct " & _
        " '['||C.编码||']'||" & IIf(gint药品名称显示 = 1, "NVL(L.名称,C.名称)", "C.名称") & " 品名,X.毒理分类" & _
        " FROM 药品收发记录 A,药品规格 B,收费项目目录 C,收费项目别名 L,药品特性 X " & _
        " WHERE A.药品ID=B.药品ID AND B.药名ID=X.药名ID And B.药品ID=C.ID " & _
        " AND B.药品ID=L.收费细目ID(+) AND L.性质(+)=3 AND L.码类(+)=1 " & _
        " AND A.审核人 IS NULL AND MOD(A.记录状态,3)=1 AND NVL(A.摘要,'小宝')<>'拒发'" & _
        " AND A.NO=[1] AND A.单据=[2] AND (A.库房ID+0=[3] OR A.库房ID IS NULL) " & _
        " And X.毒理分类<>'普通药'" & _
        " Order by X.毒理分类"
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[对毒麻类药品进行检查]", strNo, IntBillStyle, lngNO药房id)
    
    If rstemp.RecordCount = 0 Then
        CheckSpec = True
        Exit Function
    End If
    
    With rstemp
        Do While Not .EOF
            strNote = strNote & vbCrLf & Space(4) & !毒理分类 & "-" & !品名
            .MoveNext
        Loop
    End With
'    If MsgBox("是否对以下毒、麻、精神类药品进行发药？" & strNote, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    mstr毒麻类提示 = strNote
    CheckSpec = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function CheckDrugStock(ByVal rsData As ADODB.Recordset) As Boolean
    Dim lng药品id As Long
    
    If mrsDrugStock Is Nothing Then
        GetDrugStock mParams.lng药房ID
        If mrsDrugStock Is Nothing Then
            MsgBox "未设置存储库房，不能发药！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    rsData.Filter = "标志=1"
    rsData.Sort = "药品ID"
    
    Do While Not rsData.EOF
        If lng药品id <> rsData!药品ID Then
            lng药品id = rsData!药品ID
            
            mrsDrugStock.Filter = "药品ID=" & lng药品id
            If mrsDrugStock.EOF Then
                MsgBox rsData!品名 & "未设置存储库房，不能发药！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        rsData.MoveNext
    Loop

    CheckDrugStock = True
End Function

Private Function CheckSendBillMoney(ByVal rsData As ADODB.Recordset) As Boolean
    '发药检查－检查病人费用余额，并根据记帐报警设置作相应处理
    'blnBatch：True-批量发药;False-单处方发药
    '主要算法：
    '1、系统参数"执行后自动审核"有效时才检查
    '2、只对记帐划价单
    '3、按病人ID计算单据汇总金额
    '4、根据记帐报警设置作相应处理
    Dim n As Integer
    Dim rsTmp As ADODB.Recordset
    Dim rs费用类别 As ADODB.Recordset
    Dim strNo As String
    Dim lng病人ID As Long
    Dim cur处方金额 As Currency
    
    Dim strFirstNo As String
    Dim str费用类别 As String
    Dim str费用类别名 As String
    
    On Error GoTo errH
    
    '系统参数"执行后自动审核"有效时才检查
    If mParams.bln审核划价单 = False Then
        CheckSendBillMoney = True
        Exit Function
    End If
    
    If rsData Is Nothing Then
        CheckSendBillMoney = True
        Exit Function
    End If
    
    With rsData
        '只对记帐划价单才检查
        .Filter = "单据=9 And 已收费=0"
        
        '按病人ID计算单据汇总金额
        .Sort = "病人ID"
        
        If .RecordCount = 0 Then
            CheckSendBillMoney = True
            Exit Function
        End If
        
        .MoveFirst
        
        '根据记帐报警设置作相应处理
        Do While Not .EOF
            If lng病人ID <> Val(!病人ID) Then
                If lng病人ID <> 0 Then
                    '判断是住院还是门诊病人
                    If !记录性质 = 1 Or (!记录性质 = 2 And (!门诊标志 = 1 Or !门诊标志 = 4)) Then
                        gstrSQL = "Select Distinct '门诊' As 来源, " & _
                            " B.病人id,0 主页id,0 病人病区id, C.姓名 " & _
                            " From 药品收发记录 A,门诊费用记录 B,病人信息 C " & _
                            " Where A.费用id=B.Id And b.病人id = c.病人id " & _
                            " And A.单据=9 And A.no=[1] "
                    Else
                        gstrSQL = " Select Distinct '住院' As 来源, " & _
                            " B.病人id,nvl(B.主页id,0) 主页id,B.病人病区id, C.姓名 " & _
                            " From 药品收发记录 A,住院费用记录 B,病人信息 C " & _
                            " Where A.费用id=B.Id And b.病人id = c.病人id " & _
                            " And A.单据=9 And A.no=[1] "
                    End If
                    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strFirstNo)
                    
                    '取费用类别
                    gstrSQL = " Select /*+ Rule*/ Distinct b.编码, b.名称 " & _
                    " From 门诊费用记录 a, 收费项目类别 b, 药品收发记录 c,Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) d " & _
                    " Where a.收费类别 = b.编码 And a.Id = c.费用id And c.单据 = 9 And c.No=d.Column_Value "
                    If !记录性质 = 1 Or (!记录性质 = 2 And (!门诊标志 = 1 Or !门诊标志 = 4)) Then
                    Else
                        gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
                    End If
                    Set rs费用类别 = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo)
                    
                    Do While Not rs费用类别.EOF
                        str费用类别 = str费用类别 & rs费用类别!编码
                        str费用类别名 = str费用类别名 & "," & rs费用类别!名称
                        rs费用类别.MoveNext
                    Loop
                                        
                    '检查费用余额
                    If Not FinishBillingWarn(rsTmp, cur处方金额, str费用类别, str费用类别名) Then
                        CheckSendBillMoney = False
                        Exit Function
                    End If
                End If
                
                strNo = !NO
                cur处方金额 = Val(Get实收金额(Val(!单据), !NO, !门诊标志))
                strFirstNo = !NO
                lng病人ID = Val(!病人ID)
            Else
                strNo = strNo & "," & !NO
                cur处方金额 = cur处方金额 + Val(Get实收金额(Val(!单据), !NO, !门诊标志))
            End If
            
            .MoveNext
            
            If .EOF Then
                .MovePrevious
                '判断是住院还是门诊病人
                If !记录性质 = 1 Or (!记录性质 = 2 And (!门诊标志 = 1 Or !门诊标志 = 4)) Then
                    gstrSQL = "Select Distinct '门诊' As 来源, " & _
                        " B.病人id,0 主页id,0 病人病区id, C.姓名 " & _
                        " From 药品收发记录 A,门诊费用记录 B,病人信息 C " & _
                        " Where A.费用id=B.Id And b.病人id = c.病人id " & _
                        " And A.单据=9 And A.no=[1] "
                Else
                    gstrSQL = "Select Distinct '住院' As 来源, " & _
                        " B.病人id,nvl(B.主页id,0) 主页id,B.病人病区id, C.姓名 " & _
                        " From 药品收发记录 A,住院费用记录 B,病人信息 C " & _
                        " Where A.费用id=B.Id And b.病人id = c.病人id " & _
                        " And A.单据=9 And A.no=[1] "
                End If
                Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strFirstNo)
                
                '取费用类别
                gstrSQL = " Select /*+ Rule*/ Distinct b.编码, b.名称 " & _
                    " From 门诊费用记录 a, 收费项目类别 b, 药品收发记录 c,Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) d " & _
                    " Where a.收费类别 = b.编码 And a.Id = c.费用id And c.单据 = 9 And c.No=d.Column_Value "
                If !记录性质 = 1 Or (!记录性质 = 2 And (!门诊标志 = 1 Or !门诊标志 = 4)) Then
                Else
                    gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
                End If
                Set rs费用类别 = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo)
                
                Do While Not rs费用类别.EOF
                    str费用类别 = str费用类别 & rs费用类别!编码
                    str费用类别名 = str费用类别名 & "," & rs费用类别!名称
                    rs费用类别.MoveNext
                Loop
                                    
                '检查费用余额
                If Not FinishBillingWarn(rsTmp, cur处方金额, str费用类别, str费用类别名) Then
                    CheckSendBillMoney = False
                    Exit Function
                End If
                
                .MoveNext
            End If
        Loop
    End With
    
    CheckSendBillMoney = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function FinishBillingWarn(ByVal rsTmp As ADODB.Recordset, ByVal cur金额 As Currency, ByVal str类别 As String, ByVal str类别名 As String) As Boolean
'功能：当执行完成有自动审核的费用时，对病人费用进行记帐报警。
'参数：objRecord=包含要完成执行的病人信息的数据行
'      str类别="CDE..."，报警金额涉及到的收费类别
'      str类别名="检查,检验,..."，对应的类别名用于提示
    Dim rsPati As ADODB.Recordset
    Dim rsWarn As ADODB.Recordset
    Dim strWarn As String, intWarn As Integer
    Dim strsql As String, intR As Integer, i As Long
    Dim cur当日 As Currency
    
    On Error GoTo errH
    
    If rsTmp!来源.Value = "住院" Then
        '住院病人报警
        strsql = _
            " Select 病人ID,预交余额,费用余额,0 as 预结费用 From 病人余额 Where 性质=1 And 类型=2 And 病人ID=[1]" & _
            " Union ALL" & _
            " Select A.病人ID,0,0,Sum(金额) From 保险模拟结算 A,病案主页 B" & _
            " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And B.险类 Is Not Null And A.病人ID=[1] And A.主页ID=[2] Group by A.病人ID"
        strsql = "Select 病人ID,Nvl(Sum(预交余额),0)-Nvl(Sum(费用余额),0)+Nvl(Sum(预结费用),0) as 剩余款 From (" & strsql & ") Group by 病人ID"
        
        strsql = "Select zl_PatiWarnScheme(A.病人ID,B.主页ID) As 适用病人,C.剩余款," & _
            " Decode(A.担保额,Null,Null,zl_PatientSurety(A.病人ID,B.主页ID)) as 担保额" & _
            " From 病人信息 A,病案主页 B,(" & strsql & ") C" & _
            " Where A.病人ID=B.病人ID And A.主页id=B.主页id And A.病人ID=C.病人ID(+)" & _
            " And A.病人ID=[1] And B.主页ID=[2]"
        Set rsPati = zldatabase.OpenSQLRecord(strsql, Me.Caption, Val(rsTmp!病人ID), Val(rsTmp!主页id))
    Else
        '其他按门诊报警
        strsql = "Select 病人ID,预交余额,费用余额 From 病人余额 Where 性质=1 And 类型=1 And 病人ID=[1]"
        strsql = "Select zl_PatiWarnScheme(A.病人ID) As 适用病人,A.担保额," & _
            " Nvl(B.预交余额,0)-Nvl(B.费用余额,0)+Nvl(E.帐户余额,0) as 剩余款" & _
            " From 病人信息 A,(" & strsql & ") B,医保病人关联表 D,医保病人档案 E" & _
            " Where A.病人ID=B.病人ID(+) " & _
            " And A.病人id = D.病人id(+) And A.险类=D.险类(+) And D.险类=E.险类(+) And D.医保号=E.医保号(+) And D.标志(+)=1" & _
            " And A.病人ID=[1]"
        Set rsPati = zldatabase.OpenSQLRecord(strsql, Me.Caption, Val(rsTmp!病人ID))
    End If
    
    intWarn = -1 '记帐报警时缺省要提示
    '执行报警:门诊病人病区ID=0
    strsql = "Select Nvl(报警方法,1) as 报警方法," & _
        " 报警值,报警标志1,报警标志2,报警标志3 From 记帐报警线" & _
        " Where Nvl(病区ID,0)=[1] And 适用病人=[2]"
    Set rsWarn = zldatabase.OpenSQLRecord(strsql, Me.Caption, Val(rsTmp!病人病区ID), CStr(zlStr.NVL(rsPati!适用病人)))
    If Not rsWarn.EOF Then
        If rsWarn!报警方法 = 2 Then cur当日 = GetPatiDayMoney(Val(rsTmp!病人ID))
        str类别名 = Mid(str类别名, 2)
        For i = 1 To Len(str类别)
            intR = BillingWarn(Me, mstrPrivs, rsWarn, rsTmp!姓名, zlStr.NVL(rsPati!剩余款, 0), cur当日, cur金额, zlStr.NVL(rsPati!担保额, 0), Mid(str类别, i, 1), Split(str类别名, ",")(i - 1), strWarn, intWarn)
            If InStr(",2,3,", intR) > 0 Then Exit Function
        Next
    End If
    
    FinishBillingWarn = True
    Exit Function
errH:
    
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillingWarn(frmParent As Object, ByVal strPrivs As String, _
    rsWarn As ADODB.Recordset, ByVal str姓名 As String, ByVal cur剩余款额 As Currency, _
    ByVal cur当日金额 As Currency, ByVal Cur记帐金额 As Currency, ByVal cur担保金额 As Currency, _
    ByVal str收费类别 As String, ByVal str类别名称 As String, str已报类别 As String, _
    intWarn As Integer) As Integer
'功能:对病人记帐进行报警提示
'参数:rsWarn=包含报警参数设置的记录集(该病人病区,并区分好了医保)
'     str收费类别=当前要检查的类别,用于分类报警
'     str类别名称=类别名称,用于提示
'     intWarn=是否显示询问性的提示,-1=要显示,0=缺省为否,1-缺省为是
'返回:str已报类别="CDE":具体在本次报警的一组类别,"-"为所有类别。该返回用于处理重复报警
'     intWarn=本次询问性提示中的选择结果,0=为否,1-为是
'     0;没有报警,继续
'     1:报警提示后用户选择继续
'     2:报警提示后用户选择中断
'     3:报警提示必须中断
'     4:强制记帐报警,继续
    Dim bln已报警 As Boolean, byt标志 As Byte
    Dim byt方式 As Byte, byt已报方式 As Byte
    Dim ArrTmp As Variant, vMsg As VbMsgBoxResult
    Dim str担保 As String, i As Long
    
    BillingWarn = 0
    
    '报警参数检查:NULL是没有设置,0是设置了的
    If rsWarn.State = 0 Then Exit Function
    If rsWarn.EOF Then Exit Function
    If IsNull(rsWarn!报警值) Then Exit Function
    
    '对应类别定位有效报警设置
    If Not IsNull(rsWarn!报警标志1) Then
        If rsWarn!报警标志1 = "-" Or InStr(rsWarn!报警标志1, str收费类别) > 0 Then byt标志 = 1
        If rsWarn!报警标志1 = "-" Then str类别名称 = "" '所有类别时,不必提示具体的类别
    End If
    If byt标志 = 0 And Not IsNull(rsWarn!报警标志2) Then
        If rsWarn!报警标志2 = "-" Or InStr(rsWarn!报警标志2, str收费类别) > 0 Then byt标志 = 2
        If rsWarn!报警标志2 = "-" Then str类别名称 = "" '所有类别时,不必提示具体的类别
    End If
    If byt标志 = 0 And Not IsNull(rsWarn!报警标志3) Then
        If rsWarn!报警标志3 = "-" Or InStr(rsWarn!报警标志3, str收费类别) > 0 Then byt标志 = 3
        If rsWarn!报警标志3 = "-" Then str类别名称 = "" '所有类别时,不必提示具体的类别
    End If
    If byt标志 = 0 Then Exit Function '无有效设置
    
    '报警标志2实际上是两种判断①②,其它只有一种判断①
    '这种处理的前提是一种类别只能属于一种报警方式(报警参数设置时)
    '示例："-" 或 ",ABC,567,DEF"
    '报警标志2示例："-①" 或 ",ABC②,567①,DEF①"
    bln已报警 = InStr(str已报类别, str收费类别) > 0 Or str已报类别 Like "-*"
    
    If bln已报警 Then '当intWarn = -1时,也可强行再报警
        If byt标志 = 2 Then
            If str已报类别 Like "-*" Then
                byt已报方式 = IIf(Right(str已报类别, 1) = "②", 2, 1)
            Else
                ArrTmp = Split(str已报类别, ",")
                For i = 0 To UBound(ArrTmp)
                    If InStr(ArrTmp(i), str收费类别) > 0 Then
                        byt已报方式 = IIf(Right(ArrTmp(i), 1) = "②", 2, 1)
                        'Exit For '取消说明见住院记帐模块
                    End If
                Next
            End If
        Else
            Exit Function
        End If
    End If
    
    If str类别名称 <> "" Then str类别名称 = """" & str类别名称 & """费用"
    str担保 = IIf(cur担保金额 = 0, "", "(含担保额:" & Format(cur担保金额, "0.00") & ")")
    cur剩余款额 = cur剩余款额 + cur担保金额 - Cur记帐金额
    cur当日金额 = cur当日金额 + Cur记帐金额
        
    '---------------------------------------------------------------------
    If rsWarn!报警方法 = 1 Then  '累计费用报警(低于)
        Select Case byt标志
            Case 1 '低于报警值(包括预交款耗尽)提示询问记帐
                If cur剩余款额 < rsWarn!报警值 Then
                    If InStr(";" & strPrivs & ";", ";强制记帐;") = 0 Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & ":" & Format(cur剩余款额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",允许该病人记帐吗？", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 1
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 1
                            End If
                        End If
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox("强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款" & str担保 & ":" & Format(cur剩余款额, "0.00") & " 低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 4
                    End If
                End If
            Case 2 '低于报警值提示询问记帐,预交款耗尽时禁止记帐
                If Not bln已报警 Then
                    If cur剩余款额 < 0 Then
                        byt方式 = 2
                        If InStr(";" & strPrivs & ";", ";强制记帐;") = 0 Then
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & "已经耗尽," & str类别名称 & "禁止记帐。", frmParent, True)
                                If vMsg = vbIgnore Then intWarn = 1
                            End If
                            BillingWarn = 3
                        Else
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str类别名称 & "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款" & str担保 & "已经耗尽。", frmParent, True)
                                If vMsg = vbIgnore Then intWarn = 1
                            End If
                            BillingWarn = 4
                        End If
                    ElseIf cur剩余款额 < rsWarn!报警值 Then
                        byt方式 = 1
                        If InStr(";" & strPrivs & ";", ";强制记帐;") = 0 Then
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & ":" & Format(cur剩余款额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",允许该病人记帐吗？", frmParent)
                                If vMsg = vbNo Or vMsg = vbCancel Then
                                    If vMsg = vbCancel Then intWarn = 0
                                    BillingWarn = 2
                                ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                    If vMsg = vbIgnore Then intWarn = 1
                                    BillingWarn = 1
                                End If
                            Else
                                If intWarn = 0 Then
                                    BillingWarn = 2
                                ElseIf intWarn = 1 Then
                                    BillingWarn = 1
                                End If
                            End If
                        Else
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox("强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款" & str担保 & ":" & Format(cur剩余款额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", frmParent, True)
                                If vMsg = vbIgnore Then intWarn = 1
                            End If
                            BillingWarn = 4
                        End If
                    End If
                Else
                    '上次已报警并选择继续或强制继续
                    If byt已报方式 = 1 Then
                        '上次低于报警值并选择继续或强制继续,不再处理低于的情况,但还需要判断预交款是否耗尽
                        If cur剩余款额 < 0 Then
                            byt方式 = 2
                            If InStr(";" & strPrivs & ";", ";强制记帐;") = 0 Then
                                If intWarn = -1 Then
                                    vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & "已经耗尽," & str类别名称 & "禁止记帐。", frmParent, True)
                                    If vMsg = vbIgnore Then intWarn = 1
                                End If
                                BillingWarn = 3
                            Else
                                If intWarn = -1 Then
                                    vMsg = frmMsgBox.ShowMsgBox(str类别名称 & "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款" & str担保 & "已经耗尽。", frmParent, True)
                                    If vMsg = vbIgnore Then intWarn = 1
                                End If
                                BillingWarn = 4
                            End If
                        End If
                    ElseIf byt已报方式 = 2 Then
                        '上次预交款已经耗尽并强制继续,不再处理
                        Exit Function
                    End If
                End If
            Case 3 '低于报警值禁止记帐
                If cur剩余款额 < rsWarn!报警值 Then
                    If InStr(";" & strPrivs & ";", ";强制记帐;") = 0 Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & ":" & Format(cur剩余款额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",禁止记帐。", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 3
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox("强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款" & str担保 & ":" & Format(cur剩余款额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 4
                    End If
                End If
        End Select
    ElseIf rsWarn!报警方法 = 2 Then  '每日费用报警(高于)
        Select Case byt标志
            Case 1 '高于报警值提示询问记帐
                If cur当日金额 > rsWarn!报警值 Then
                    If InStr(";" & strPrivs & ";", ";强制记帐;") = 0 Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当日费用:" & Format(cur当日金额, gtype_UserSysParms.P9_费用金额保留位数) & ",高于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",允许该病人记帐吗？", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 1
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 1
                            End If
                        End If
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox("强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当日费用:" & Format(cur当日金额, gtype_UserSysParms.P9_费用金额保留位数) & ",高于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 4
                    End If
                End If
            Case 3 '高于报警值禁止记帐
                If cur当日金额 > rsWarn!报警值 Then
                    If InStr(";" & strPrivs & ";", ";强制记帐;") = 0 Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当日费用:" & Format(cur当日金额, gtype_UserSysParms.P9_费用金额保留位数) & ",高于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",禁止记帐。", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 3
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox("强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当日费用:" & Format(cur当日金额, gtype_UserSysParms.P9_费用金额保留位数) & ",高于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 4
                    End If
                End If
        End Select
    End If
    
    '对于继续类的操作,返回已报警类别
    If BillingWarn = 1 Or BillingWarn = 4 Then
        If byt标志 = 1 Then
            If rsWarn!报警标志1 = "-" Then
                str已报类别 = "-"
            Else
                str已报类别 = str已报类别 & "," & rsWarn!报警标志1
            End If
        ElseIf byt标志 = 2 Then
            If rsWarn!报警标志2 = "-" Then
                str已报类别 = "-"
            Else
                str已报类别 = str已报类别 & "," & rsWarn!报警标志2
            End If
            '附加标注以判断已报警的具体方式
            str已报类别 = str已报类别 & IIf(byt方式 = 2, "②", "①")
        ElseIf byt标志 = 3 Then
            If rsWarn!报警标志3 = "-" Then
                str已报类别 = "-"
            Else
                str已报类别 = str已报类别 & "," & rsWarn!报警标志3
            End If
        End If
    End If
End Function
Public Function GetPatiDayMoney(lng病人ID As Long) As Currency
'功能：获取指定病人当天发生的费用总额
    Dim rsTmp As New ADODB.Recordset
    Dim strsql As String
    
    On Error GoTo errH
    
    strsql = "Select zl_PatiDayCharge([1]) as 金额 From Dual"
    Set rsTmp = zldatabase.OpenSQLRecord(strsql, "mdlCISKernel", lng病人ID)
    If Not rsTmp.EOF Then GetPatiDayMoney = zlStr.NVL(rsTmp!金额, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function CheckBill(ByVal lngNO药房id As Long, ByVal IntOper As Integer, ByVal IntBillStyle As Integer, ByVal strNo As String, ByVal int记录性质 As Integer, ByVal int门诊标志 As Integer, Optional ByVal bln提示 As Boolean = False) As Integer
    Dim dblCount As Double
    Dim intRow As Integer, intRows As Integer
    Dim rstemp As New ADODB.Recordset
    Dim RecCheck As New ADODB.Recordset
    Dim vsfDetail As VSFlexGrid
    
    '--根据将要执行的操作，判断是否允许--
    'IntOper:1-配药;2-取消配药;3-发药;4-退药;5-取消发药
    '返回:
    '0-允许操作
    '1-未配药
    '2-已配药
    '3-已发药
    '4-已删除
    '5-未发药
    On Error GoTo errHandle
    If lngNO药房id = 0 Then lngNO药房id = mParams.lng药房ID
    
    '单独处理取消发药时的检查
    If IntOper = 5 Then
        gstrSQL = "Select 审核人 From 药品收发记录 Where No=[1] And 单据=[2] And 库房ID+0=[3] And 记录状态=1 And 审核人 IS Not Null And Rownum=1 "
        Set RecCheck = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, IntBillStyle, lngNO药房id)
        If RecCheck.EOF Then
            CheckBill = 4
            MsgBox "未找到指定单据，或已被其他操作员处理,操作被迫中止！", vbInformation, gstrSysName
        End If
        Exit Function
    End If
     
    gstrSQL = " Select A.配药人,A.审核人 From 药品收发记录 A" & _
        " Where A.No=[1] And A.单据=[2] " & _
        " " & IIf(IntOper <> 4, " And mod(A.记录状态,3)=1", "") & " And Rownum=1 " & _
        " And Nvl(Ltrim(Rtrim(A.摘要)),'小宝')<>'拒发' And (A.库房ID+0=[3] Or A.库房ID Is NULL)"
    
    If IntOper = 4 Then
        gstrSQL = gstrSQL & " And 审核人 IS Not Null"
    Else
        gstrSQL = gstrSQL & " And 审核人 IS Null"
    End If

    Set RecCheck = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, IntBillStyle, lngNO药房id)
    
    With RecCheck
        If .EOF Then CheckBill = 4: MsgBox "未找到指定单据,可能已经被其他操作员处理,操作被迫中止！", vbInformation, gstrSysName: Exit Function
        If Not IsNull(!审核人) Then
            If InStr(1, "123", IntOper) <> 0 Then CheckBill = 3: MsgBox "该处方已被其它操作员发药，" & IIf(IntOper = 1, "配药", IIf(IntOper = 2, "取消配药", IIf(IntOper = 3, "发药", "退药"))) & "操作中止！", vbInformation, gstrSysName: Exit Function
        Else
            If InStr(1, "4", IntOper) <> 0 Then CheckBill = 5: MsgBox "该处方还未发药，操作被迫中止！", vbInformation, gstrSysName: Exit Function
            If Not IsNull(!配药人) Then
                If InStr(1, "1", IntOper) <> 0 Then CheckBill = 2: MsgBox "该处方已配药，操作被迫中止！", vbInformation, gstrSysName: Exit Function
            Else
                If InStr(1, "2", IntOper) <> 0 Then CheckBill = 1: MsgBox "该处方未配药，操作被迫中止！", vbInformation, gstrSysName: Exit Function
            End If
        End If
    End With
    
    '如果是退药，检查是否允许未作废医嘱退药
    If mParams.bln医嘱作废 = False And bln提示 Then
        Set vsfDetail = mfrmDetail.GetDetailList
        intRows = vsfDetail.rows - 2
        For intRow = 1 To intRows
            dblCount = Val(vsfDetail.TextMatrix(intRow, vsfDetail.ColIndex("退药数")))
            If dblCount <> 0 Then
                gstrSQL = "select 扣率 From 药品收发记录 Where ID=[1] "
                Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[检查是否是临嘱]", Val(vsfDetail.TextMatrix(intRow, vsfDetail.ColIndex("Id"))))

                If (rstemp!扣率 Like "1*") Then       '临嘱
                    gstrSQL = "select B.执行状态 from 病人医嘱记录 A,病人医嘱发送 B,门诊费用记录 C where A.相关id=B.医嘱ID and A.id=C.医嘱序号 and  C.ID=(Select 费用ID From 药品收发记录 Where ID=[1])"
                    If int记录性质 = 1 Or (int记录性质 = 2 And (int门诊标志 = 1 Or int门诊标志 = 4)) Then
                    Else
                        gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
                    End If
                    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[检查医嘱的给药途径是否已经执行]", Val(vsfDetail.TextMatrix(intRow, vsfDetail.ColIndex("Id"))))
                
                    If Not rstemp.EOF Then
                        If rstemp!执行状态 = 0 Then
                            gstrSQL = "Select Nvl(医嘱序号,0) 医嘱序号,Nvl(门诊标志,1) 门诊标志 From 门诊费用记录 Where ID=(Select 费用ID From 药品收发记录 Where ID=[1])"
                            If int记录性质 = 1 Or (int记录性质 = 2 And (int门诊标志 = 1 Or int门诊标志 = 4)) Then
                            Else
                                gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
                            End If
                            
                            Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[检查是否是医嘱]", Val(vsfDetail.TextMatrix(intRow, vsfDetail.ColIndex("Id"))))
        
                            If Not rstemp.EOF Then
                                If (rstemp!门诊标志 = 1 Or rstemp!门诊标志 = 4) And rstemp!医嘱序号 <> 0 Then
                                    gstrSQL = "Select Nvl(主页id, 0) As 主页id, 挂号单, decode(医嘱状态,4,1,0) 作废 From 病人医嘱记录 Where 病人来源=1  And ID=[1]"
                                    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[判断该医嘱是否作废]", CLng(rstemp!医嘱序号))
                                    
                                    If Not rstemp.EOF Then
                                        If rstemp!主页id > 0 And IsNull(rstemp!挂号单) Then
                                            '填了主页ID，但没有挂号单的不受医嘱是否作废的限制
                                        Else
                                            If rstemp!作废 = 0 Then
                                                CheckBill = 1
                                                MsgBox "第" & intRow & "行的药品记录对应的医嘱还未作废，不允许退药！", vbInformation, gstrSysName
                                                Exit Function
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        gstrSQL = "Select Nvl(医嘱序号,0) 医嘱序号,Nvl(门诊标志,1) 门诊标志 From 门诊费用记录 Where ID=(Select 费用ID From 药品收发记录 Where ID=[1])"
                        If int记录性质 = 1 Or (int记录性质 = 2 And (int门诊标志 = 1 Or int门诊标志 = 4)) Then
                        Else
                            gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
                        End If
                        
                        Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[检查是否是医嘱]", Val(vsfDetail.TextMatrix(intRow, vsfDetail.ColIndex("Id"))))
    
                        If Not rstemp.EOF Then
                            If (rstemp!门诊标志 = 1 Or rstemp!门诊标志 = 4) And rstemp!医嘱序号 <> 0 Then
                                gstrSQL = "Select Nvl(主页id, 0) As 主页id, 挂号单, decode(医嘱状态,4,1,0) 作废 From 病人医嘱记录 Where 病人来源=1  And ID=[1]"
                                Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[判断该医嘱是否作废]", CLng(rstemp!医嘱序号))
                                
                                If Not rstemp.EOF Then
                                    If rstemp!主页id > 0 And IsNull(rstemp!挂号单) Then
                                        '填了主页ID，但没有挂号单的不受医嘱是否作废的限制
                                    Else
                                        If rstemp!作废 = 0 Then
                                            CheckBill = 1
                                            MsgBox "第" & intRow & "行的药品记录对应的医嘱还未作废，不允许退药！", vbInformation, gstrSysName
                                            Exit Function
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End If
    
    CheckBill = 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function CheckBillExist(ByVal Int单据 As Integer, ByVal strNo As String) As Boolean
    Dim rstemp As ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = "Select ID From 药品收发记录 " & _
             " Where 单据=[1] And NO=[2] And Rownum<2"
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "检查单据是否存在", Int单据, strNo)
    CheckBillExist = Not rstemp.EOF
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub RefreshList(ByVal intType As Integer)
    If mblnStart = False Then Exit Sub

    Call AviShow(Me)
    
    Call GetCondition
    
    Select Case intType
        Case mListType.配药确认
            RefreshList_DosageOk
        Case mListType.待配药
            Call RefreshList_Dosage
        Case mListType.已配药
            Call RefreshList_Abolish
        Case mListType.待发药
            Call RefreshList_Send
        Case mListType.超时未发
            Call RefreshList_OverTime
        Case mListType.退药
            Call RefreshList_Return
    End Select
    
    Call AviShow(Me, False)
    
    If mblnInput = False Then
        With mfrmList.vsfList
            If .Visible And .Enabled Then .SetFocus
        End With
    Else
        If txtPati.Enabled = True Then txtPati.SetFocus
    End If
End Sub

Private Sub CheckOtherUndeliveredDocuments(ByVal lng病人ID As Long)
    '功能:根据参数检查的当前病人在[当前药房其他窗口]或[其他药房]是否存在未发药单据
    Dim rstemp As New ADODB.Recordset
    Dim date开始日期 As Date
    Dim date结束日期 As Date
    Dim dteTime As Date
    Dim strMsg As String
    Dim i As Integer
    
    On Error GoTo errHandle
    
    If mParams.int查询未发药单据天数 = 0 Then Exit Sub
        
    dteTime = Sys.Currentdate
    date开始日期 = CDate(Format(DateAdd("d", -mParams.int查询未发药单据天数 + 1, dteTime), "yyyy-mm-dd") & " 00:00:00")
    date结束日期 = CDate(DateAdd("s", -1, Format(DateAdd("d", 1, dteTime), "yyyy-mm-dd") & " 00:00:00"))
    
    gstrSQL = "Select Distinct a.No, a.发药窗口, b.名称 As 药房名称" & vbNewLine & _
        "From 未发药品记录 A, 部门表 B" & vbNewLine & _
        "Where a.库房id = b.Id And a.病人id = [1] And a.库房id = [2] And a.发药窗口 Is Not Null And a.发药窗口 Not In (Select b.Column_Value From Table(Cast(f_Str2list([3]) As Zltools.t_Strlist)) B) And" & vbNewLine & _
        "      a.填制日期 Between [4] And [5]" & vbNewLine & _
        "Union All" & vbNewLine & _
        "Select a.No, a.发药窗口, b.名称 As 药房名称" & vbNewLine & _
        "From 未发药品记录 A, 部门表 B" & vbNewLine & _
        "Where a.库房id = b.Id And a.病人id = [1] And a.库房id <> [2] And a.填制日期 Between [4] And [5]"
    
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "提起病人其他药房或其他窗口的未发药单据", _
            lng病人ID, _
            mParams.lng药房ID, _
            mParams.Str窗口, _
            date开始日期, _
            date结束日期)
    
    '若存在则弹出提示框
    If Not rstemp.EOF Then
        If rstemp.RecordCount > 3 Then
            For i = 1 To 3
                strMsg = strMsg & vbCrLf & "单据号:" & rstemp!NO & "   药房:" & rstemp!药房名称 & "   发药窗口:" & IIf(IsNull(rstemp!发药窗口), "空", rstemp!发药窗口)
                rstemp.MoveNext
            Next
            strMsg = strMsg & vbCrLf & "等一共 " & rstemp.RecordCount & " 条单据"
        Else
            Do While Not rstemp.EOF
                strMsg = strMsg & vbCrLf & "单据号:" & rstemp!NO & "   药房:" & rstemp!药房名称 & "   发药窗口:" & IIf(IsNull(rstemp!发药窗口), "空", rstemp!发药窗口)
                rstemp.MoveNext
            Loop
        End If
        
        MsgBox "该病人还有其他未发药处方单据" & strMsg
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub RefreshList_DosageOk()
    '刷新配药列表
    Dim bln医保号 As Boolean
    Dim rsData As ADODB.Recordset
    Dim strSub1 As String
    Dim strSub2 As String
    Dim str住院 As String
    Dim str门诊 As String
    Dim lng病人ID As Long
    Dim strSqlTmp As String
    Dim str发生时间 As String
    
    On Error GoTo errHandle
    If mSQLCondition.str身份证 <> "" And Not mobjSquareCard Is Nothing Then
        If mobjSquareCard.zlGetPatiID("身份证", UCase(mSQLCondition.str身份证), False, lng病人ID) = False Then lng病人ID = 0
    End If
    
    gstrSQL = "Select '' As 颜色, 处方类型 ,'' As 选择 ,'0' As 标志,类型,单据,已收费,配药人,NO,姓名,签到时间," & _
            " to_Char(Sum(Round(零售金额," & mintMoneyDigit & ")),'" & mstrOracleMoneyForamt & "') AS 金额,日期,可操作," & _
            " 说明,就诊卡号,门诊号,身份证号,IC卡号,病人ID,医保号,住院号,排队状态,发药窗口," & _
            " Sum(Round(实收金额," & mintMoneyDigit & ")) 实收金额,门诊标志,记录性质,Zl_Get收费类别(单据,NO,[1]) As 收费类别,病人类型 " & _
            " From ("
            
    strSqlTmp = " Select A.优先级,A.类型,A.单据,A.已收费,A.配药人,A.NO,A.姓名,C.零售金额,A.日期,A.可操作,A.说明,A.就诊卡号,A.门诊号,A.身份证号," & _
            " A.IC卡号,A.病人ID,A.医保号,A.住院号,A.排队状态,d.实收金额, Nvl(A.处方类型,Nvl(C.注册证号,0)) As 处方类型,D.门诊标志,D.记录性质,D.收费类别,A.发药窗口,c.签到时间,a.病人类型 " & _
            " From ("
    
    str门诊 = "Select distinct B.就诊卡号,B.门诊号,B.身份证号,B.IC卡号,B.医保号,B.住院号,A.优先级,A.发药窗口,A.填制日期,Decode(Nvl(A.已收费,0),1,'','(未)')||Decode(A.单据,8,'收费',9,'记帐') 类型," & _
            " A.单据,A.已收费,'' 配药人,A.No,A.姓名,To_Char(A.填制日期,'yyyy-MM-dd hh24:mi:ss') 日期,1 可操作,' ' 说明,B.病人ID, A.处方类型,A.排队状态,a.对方部门id, b.病人类型 " & _
            " From 未发药品记录 A,病人信息 B,门诊费用记录 C " & _
            " Where A.配药人 is null "
    
    '是否显示已确认单据
    If mcondition.bln显示已确认单据 = False Then
        str门诊 = str门诊 & " And (A.排队状态=0 or A.排队状态 is null) "
    Else
        str门诊 = str门诊 & " And (A.排队状态=0 or A.排队状态=1 or A.排队状态 is null) "
    End If
    
    '主要条件
    str门诊 = str门诊 & " And (A.库房ID=[1] Or A.库房ID Is NULL) And A.填制日期 Between [2] And [3] "
    
    If mSQLCondition.str开始NO <> "" Or mSQLCondition.str结束NO <> "" Then
        If mSQLCondition.str开始NO <> "" And mSQLCondition.str结束NO <> "" Then
            str门诊 = str门诊 & " And A.NO Between [4] And [5] "
        Else
            If mSQLCondition.str开始NO <> "" Then
                str门诊 = str门诊 & " And A.NO = [4] "
            Else
                str门诊 = str门诊 & " And A.NO = [5] "
            End If
        End If
    End If
    
    If mSQLCondition.str姓名 <> "" Then str门诊 = str门诊 & " And Upper(A.姓名) Like [6] "
    
    If mSQLCondition.str就诊卡 <> "" Then str门诊 = str门诊 & " And Upper(B.就诊卡号) = [7] "
    
    If mSQLCondition.str标识号 <> "" Then str门诊 = str门诊 & " And Upper(DECODE(A.单据,8,B.门诊号,B.住院号)) Like [8] "
    
    If mSQLCondition.lng科室ID > 0 Then str门诊 = str门诊 & " And A.对方部门ID+0=[9] "
    
    If mSQLCondition.str当前NO <> "" Then str门诊 = str门诊 & " And A.NO=[13] "
    
    If mSQLCondition.str门诊号 <> "" Then str门诊 = str门诊 & " And B.门诊号=[14] "
    
'    If mSQLCondition.str身份证 <> "" Then str门诊 = str门诊 & " And B.身份证号=[15] "

    If mSQLCondition.str身份证 <> "" Then str门诊 = str门诊 & " And B.病人ID=[15] "

    If mSQLCondition.lng病人ID <> 0 Then str门诊 = str门诊 & " And B.病人ID=[16] "
    
    If mSQLCondition.str医保号 <> "" Then str门诊 = str门诊 & " And B.医保号=[17] "
    
    If mSQLCondition.lng住院号 <> 0 Then str门诊 = str门诊 & " And B.住院号=[18] "
    
            
    bln医保号 = (mSQLCondition.str医保号 <> "")
    str门诊 = str门诊 & " And A.病人ID=B.病人ID" & IIf(bln医保号 = True, "", "(+)") & ""
    
    str门诊 = str门诊 & IIf(mParams.Str窗口 = "", "", " And (A.发药窗口 In (Select * From Table(Cast(f_Str2list([19]) As Zltools.t_Strlist))) Or A.发药窗口 Is Null) ")
    
    '设置显示及自动打印的条件:注意"未发药品记录"的别名为A
    Select Case mParams.intShowBill收费
        Case 0  '不显示处方
            strSub1 = "1=2"
        Case 1  '显示未收费
            strSub1 = "A.单据<>9 And Nvl(A.已收费,0)=0 And A.单据=8"
        Case 2  '显示已收费
            strSub1 = "A.单据<>9 And A.已收费=1 And A.单据=8"
        Case 3  '显示所有处方
            strSub1 = "A.单据<>9 And A.单据=8"
    End Select
    Select Case mParams.intShowBill记帐
        Case 0  '不显示处方
            strSub2 = "1=2"
        Case 1  '显示未审核
            strSub2 = "A.单据<>8 And Nvl(A.已收费,0)=0 And A.单据=9"
        Case 2  '显示已审核
            strSub2 = "A.单据<>8 And A.已收费=1 And A.单据=9"
        Case 3  '显示所有处方
            strSub2 = "A.单据<>8 And A.单据=9"
    End Select
    
    str门诊 = str门诊 & " And A.单据 IN(8,9) And (" & strSub1 & " Or " & strSub2 & ")"
    
    str门诊 = str门诊 & " And Mod(C.记录性质, 10) = Decode(A.单据, 8, 1, 2) And A.No = C.No And A.库房id = C.执行部门id "
            
    If mParams.bln发生时间过滤 = False Then
'        str门诊 = Replace(str门诊, ",门诊费用记录 C", "")
    Else
        str发生时间 = Replace(str门诊, "And A.填制日期 Between [2] And [3]", "")
        str门诊 = str门诊 & " And C.医嘱序号 Is Null "
        
        str发生时间 = str发生时间 & " And C.医嘱序号 Is Not Null And C.发生时间 Between [2] And [3] "
        str门诊 = str门诊 & " Union All " & str发生时间
    End If
    
    str门诊 = strSqlTmp & str门诊 & ") A,药品收发记录 C, 门诊费用记录 D, 部门表 B " & IIf(mParams.strSourceDep = "", "", ",Table(Cast(f_Num2List([20]) As zlTools.t_NumList)) E ") & _
              " Where C.费用id = D.ID And nvl(c.发药方式,-999)<>-1 and A.单据=C.单据 And A.NO=C.NO And C.审核人 Is NULL " & _
              " And Nvl(D.费用状态,0)<>1 And (C.库房id=[1] Or C.库房id Is null)  And a.对方部门id = b.Id "
    
    If mstrDeptNode <> "" Then
        str门诊 = str门诊 & " And (b.站点 = [21] Or b.站点 Is Null) "
    End If
    
    '排除已经部分停止发药的No
    str门诊 = str门诊 & " and Not Exists(select 1 from 药品收发记录 F where F.单据=C.单据 and F.库房id=C.库房id and F.no=C.no and 发药方式=-1) "
    
    '排除已在输液配置中心管理中产生的单据
    str门诊 = str门诊 & " And Not Exists (Select 1 From 输液配药内容 Y Where Y.收发id = C.ID) "
    
    '离院带药:'0-所有,1-不含离院带药,2-仅含离院带药
    If mcondition.int离院带药 = 0 Then
    ElseIf mcondition.int离院带药 = 1 Then
        str门诊 = str门诊 & " And Not Ltrim(To_Char(Nvl(C.扣率,0),'00')) Like '_3'"
    ElseIf mcondition.int离院带药 = 2 Then
        str门诊 = str门诊 & " And Ltrim(To_Char(Nvl(C.扣率,0),'00')) Like '_3'"
    End If
    
    str门诊 = str门诊 & IIf(mParams.strSourceDep = "", "", " And C.对方部门id+0=E.Column_Value ")
    
    If imgFilter.BorderStyle = cstFilter Then
        If Val(mParams.int输入模式) <= 7 Then
            If Trim(txtPati.Text) = "" Then
                str门诊 = str门诊 & " And 1 = 2 "
            End If
        Else
            If mSQLCondition.lng病人ID = 0 Then
                str门诊 = str门诊 & " And 1 = 2 "
            End If
        End If
    End If
    
    If mParams.intType = 1 Then
        str门诊 = str门诊 & " And (D.门诊标志=1 or D.门诊标志=4)"
    ElseIf mParams.intType = 2 Then
        str门诊 = str门诊 & " And (D.门诊标志<>1 and D.门诊标志<>4)"
    End If
    
    If mcondition.int服务对象 = 1 Then
        '门诊划价及门诊记帐
        gstrSQL = gstrSQL & str门诊
    Else
        If mcondition.int服务对象 = 3 Then
            '门诊及住院所有单据
            str住院 = Replace(str门诊, "门诊费用记录", "住院费用记录")
            str住院 = Replace(str住院, "And Nvl(D.费用状态,0)<>1", "")
        Else
            '住院记帐
            str住院 = Replace(str门诊, "门诊费用记录", "住院费用记录")
            str住院 = Replace(str住院, "And Nvl(D.费用状态,0)<>1", "")
            str门诊 = ""
        End If
    
        If mPrives.bln发病区处方 Then
            If img病区.BorderStyle = 0 Then
                '不显示病区处方
                str住院 = str住院 & " And (D.门诊标志 <> 2 Or (D.门诊标志 = 2 And D.病人病区id <> D.开单部门id)) "
            End If
            If img病区.BorderStyle = 1 And cbo病区.ListIndex <> -1 Then
                '要显示病区处方，并且病人病区等于当前选择的病区
                str住院 = str住院 & " And D.病人病区id = " & cbo病区.ItemData(cbo病区.ListIndex)
                str门诊 = ""
            End If
        End If
        
        If str门诊 = "" Then
            gstrSQL = gstrSQL & str住院
        Else
            gstrSQL = gstrSQL & str门诊 & " Union All " & str住院
        End If
    End If
    
    '''''Group By
    gstrSQL = gstrSQL & ") A GROUP BY A.优先级,A.类型,A.单据,A.已收费,A.配药人,A.No,A.姓名,A.日期,A.可操作,A.说明," & _
        " A.就诊卡号,A.门诊号,A.身份证号,A.IC卡号,A.病人ID,A.医保号,A.住院号,A.排队状态,A.处方类型,A.门诊标志,A.记录性质, a.发药窗口,A.签到时间,a.病人类型 "
    
    '''''Order By
    gstrSQL = gstrSQL & " Order By A.签到时间,A.类型,A.单据,A.No"
    
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
            mSQLCondition.lng药房ID, _
            mSQLCondition.date开始日期, _
            mSQLCondition.date结束日期, _
            mSQLCondition.str开始NO, _
            mSQLCondition.str结束NO, _
            UCase(mSQLCondition.str姓名), _
            mSQLCondition.str就诊卡, _
            mSQLCondition.str标识号, _
            mSQLCondition.lng科室ID, _
            mSQLCondition.str填制人, _
            mSQLCondition.str审核人, _
            mSQLCondition.lng药品id, _
            mSQLCondition.str当前NO, _
            mSQLCondition.str门诊号, _
            lng病人ID, _
            mSQLCondition.lng病人ID, _
            mSQLCondition.str医保号, _
            mSQLCondition.lng住院号, _
            mParams.Str窗口, _
            mParams.strSourceDep, _
            mstrDeptNode)
    
    stbThis.Panels(2) = ""
    If Not rsData.EOF Then
        stbThis.Panels(2) = "共有" & rsData.RecordCount & "张处方；" & GetSumMoney(rsData)
    End If
    
    Set mrsList = rsData
    
    If Not mfrmList Is Nothing Then mfrmList.RefreshList mListType.配药确认, mrsList
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub




Public Sub RefreshList_Dosage(Optional ByVal strNo As String, Optional ByVal blnNoRefreshDetail As Boolean)
    '刷新配药列表
    Dim bln医保号 As Boolean
    Dim rsData As ADODB.Recordset
    Dim strSub1 As String
    Dim strSub2 As String
    Dim str住院 As String
    Dim str门诊 As String
    Dim cbrControl As CommandBarControl
    Dim cbrMenu As CommandBarControl
    Dim lng病人ID As Long
    Dim strSqlTmp As String
    Dim str发生时间 As String
    
    On Error GoTo errHandle
    If mSQLCondition.str身份证 <> "" And Not mobjSquareCard Is Nothing Then
        If mobjSquareCard.zlGetPatiID("身份证", UCase(mSQLCondition.str身份证), False, lng病人ID) = False Then lng病人ID = 0
    End If
    
    gstrSQL = "Select '' As 颜色, 处方类型 ,'' As 选择 ,'0' As 标志,类型,单据,已收费,配药人,NO,姓名,签到时间," & _
            " to_Char(Sum(Round(零售金额," & mintMoneyDigit & ")),'" & mstrOracleMoneyForamt & "') AS 金额,日期,可操作," & _
            " 说明,就诊卡号,门诊号,身份证号,IC卡号,病人ID,医保号,住院号,发药窗口," & _
            " Sum(Round(实收金额," & mintMoneyDigit & ")) 实收金额,门诊标志,记录性质,Zl_Get收费类别(单据,NO,[1]) As 收费类别,病人类型, 打印状态 " & IIf(mParams.bln启用审方, ",审查结果,审查id", "") & _
            " From ("
            
    strSqlTmp = " Select A.优先级,A.类型,A.单据,A.已收费,A.配药人,A.NO,A.姓名,C.零售金额,A.日期,A.可操作,A.说明,A.就诊卡号,A.门诊号,A.身份证号," & _
            " A.IC卡号,A.病人ID,A.医保号,A.住院号,d.实收金额*(Nvl(c.付数,1)*c.实际数量/(Nvl(d.付数,1)*d.数次)) 实收金额, Nvl(A.处方类型,Nvl(C.注册证号,0)) As 处方类型,D.门诊标志,D.记录性质,D.收费类别,A.发药窗口,c.签到时间,a.病人类型, a.打印状态 " & IIf(mParams.bln启用审方, ",a.审查结果,a.审查id", "") & _
            " From ("
            
    str门诊 = "Select distinct B.就诊卡号,B.门诊号,B.身份证号,B.IC卡号,B.医保号,B.住院号,A.优先级,A.发药窗口,A.填制日期,Decode(Nvl(A.已收费,0),1,'','(未)')||Decode(A.单据,8,'收费',9,'记帐') 类型," & _
            " A.单据,A.已收费,'' 配药人,A.No,A.姓名,To_Char(A.填制日期,'yyyy-MM-dd hh24:mi:ss') 日期,1 可操作,' ' 说明,B.病人ID, A.处方类型,a.对方部门id, b.病人类型, Decode(a.打印状态,1,1,3,1,0) 打印状态 " & IIf(mParams.bln启用审方, ",Q.审查结果,Q.id  审查id", "") & _
            " From 未发药品记录 A,病人信息 B,门诊费用记录 C " & IIf(mParams.bln启用审方, ",处方审查记录 Q,处方审查明细 K ", "") & _
            " Where A.配药人 Is Null "
            
    str门诊 = str门诊 & IIf(mParams.bln启用审方, " and c.医嘱序号=k.医嘱id(+) and Q.id(+)=K.审方id and K.最后提交(+)=1 ", "")
    
    '是否有配药确认环节
    If mParams.blnMustDosageOkProcess = True Then
        str门诊 = str门诊 & " and A.排队状态=1"
    End If
    
    '主要条件
    str门诊 = str门诊 & " And (A.库房ID=[1] Or A.库房ID Is NULL) And A.填制日期 Between [2] And [3] "
    
    If mSQLCondition.str开始NO <> "" Or mSQLCondition.str结束NO <> "" Then
        If mSQLCondition.str开始NO <> "" And mSQLCondition.str结束NO <> "" Then
            str门诊 = str门诊 & " And A.NO Between [4] And [5] "
        Else
            If mSQLCondition.str开始NO <> "" Then
                str门诊 = str门诊 & " And A.NO = [4] "
            Else
                str门诊 = str门诊 & " And A.NO = [5] "
            End If
        End If
    End If
    
    If mSQLCondition.str姓名 <> "" Then str门诊 = str门诊 & " And Upper(A.姓名) Like [6] "
    
    If mSQLCondition.str就诊卡 <> "" Then str门诊 = str门诊 & " And Upper(B.就诊卡号) = [7] "
    
    If mSQLCondition.str标识号 <> "" Then str门诊 = str门诊 & " And Upper(DECODE(A.单据,8,B.门诊号,B.住院号)) Like [8] "
    
    If mSQLCondition.lng科室ID > 0 Then str门诊 = str门诊 & " And A.对方部门ID+0=[9] "
    
    If mSQLCondition.str当前NO <> "" Then str门诊 = str门诊 & " And A.NO=[13] "
    
    If mSQLCondition.str门诊号 <> "" Then str门诊 = str门诊 & " And B.门诊号=[14] "
    
'    If mSQLCondition.str身份证 <> "" Then str门诊 = str门诊 & " And B.身份证号=[15] "
    
    If mSQLCondition.str身份证 <> "" Then str门诊 = str门诊 & " And B.病人ID=[15] "
    
    If mSQLCondition.lng病人ID <> 0 Then str门诊 = str门诊 & " And B.病人ID=[16] "
    
    If mSQLCondition.str医保号 <> "" Then str门诊 = str门诊 & " And B.医保号=[17] "
    
    If mSQLCondition.lng住院号 <> 0 Then str门诊 = str门诊 & " And B.住院号=[18] "
    
            
    bln医保号 = (mSQLCondition.str医保号 <> "")
    str门诊 = str门诊 & " And A.病人ID=B.病人ID" & IIf(bln医保号 = True, "", "(+)") & ""
    
    str门诊 = str门诊 & IIf(mParams.Str窗口 = "", "", " And (A.发药窗口 In (Select * From Table(Cast(f_Str2list([19]) As Zltools.t_Strlist))) Or A.发药窗口 Is Null) ")
    
    '设置显示及自动打印的条件:注意"未发药品记录"的别名为A
    Select Case mParams.intShowBill收费
        Case 0  '不显示处方
            strSub1 = "1=2"
        Case 1  '显示未收费
            strSub1 = "A.单据<>9 And Nvl(A.已收费,0)=0 And A.单据=8"
        Case 2  '显示已收费
            strSub1 = "A.单据<>9 And A.已收费=1 And A.单据=8"
        Case 3  '显示所有处方
            strSub1 = "A.单据<>9 And A.单据=8"
    End Select
    Select Case mParams.intShowBill记帐
        Case 0  '不显示处方
            strSub2 = "1=2"
        Case 1  '显示未审核
            strSub2 = "A.单据<>8 And Nvl(A.已收费,0)=0 And A.单据=9"
        Case 2  '显示已审核
            strSub2 = "A.单据<>8 And A.已收费=1 And A.单据=9"
        Case 3  '显示所有处方
            strSub2 = "A.单据<>8 And A.单据=9"
    End Select
    
    str门诊 = str门诊 & " And A.单据 IN(8,9) And (" & strSub1 & " Or " & strSub2 & ")"
    
    str门诊 = str门诊 & " And Mod(C.记录性质, 10) = Decode(A.单据, 8, 1, 2) And A.No = C.No And A.库房id = C.执行部门id "
    
    '配药打印状态：0-显示所有配药单,1-只显示未打印的待配药单据,2-只显示已打印的待配药单据
    If mParams.intShowBill配药 = 1 Then
        str门诊 = str门诊 & " And Nvl(A.打印状态,0) Not In(1,3)"
    ElseIf mParams.intShowBill配药 = 2 Then
        str门诊 = str门诊 & " And Nvl(A.打印状态,0) In(1,3)"
    End If
    
    If mParams.bln发生时间过滤 = False Then
'        str门诊 = Replace(str门诊, ",门诊费用记录 C", "")
    Else
        
        
        str发生时间 = Replace(str门诊, "And A.填制日期 Between [2] And [3]", "")
        str门诊 = str门诊 & " And C.医嘱序号 Is Null "
        
        str发生时间 = str发生时间 & " And C.医嘱序号 Is Not Null And C.发生时间 Between [2] And [3] "
        str门诊 = str门诊 & " Union All " & str发生时间
    End If
    
    str门诊 = strSqlTmp & str门诊 & ") A,药品收发记录 C, 门诊费用记录 D, 部门表 B " & IIf(mParams.strSourceDep = "", "", ",Table(Cast(f_Num2List([20]) As zlTools.t_NumList)) E ") & _
              " Where C.费用id = D.ID And nvl(c.发药方式,-999)<>-1 and A.单据=C.单据 And A.NO=C.NO And C.审核人 Is NULL " & _
              " And Nvl(D.费用状态,0)<>1 And (C.库房id=[1] Or C.库房id Is null)  And a.对方部门id = b.Id "
    
    If mstrDeptNode <> "" Then
        str门诊 = str门诊 & " And (b.站点 = [21] Or b.站点 Is Null) "
    End If
    
    '排除已经部分停止发药的No
    str门诊 = str门诊 & " and Not Exists(select 1 from 药品收发记录 F where F.单据=C.单据 and F.库房id=C.库房id and F.no=C.no and 发药方式=-1) "
    
    '排除已在输液配置中心管理中产生的单据
    str门诊 = str门诊 & " And Not Exists (Select 1 From 输液配药内容 Y Where Y.收发id = C.ID) "
    
    '是否显示退药待发单据
    If mcondition.bln显示退药待发单据 = False Then
        str门诊 = str门诊 & " And C.记录状态=1 "
    Else
        str门诊 = str门诊 & " And MOD(C.记录状态,3)=1 "
    End If
    
    '离院带药:'0-所有,1-不含离院带药,2-仅含离院带药
    If mcondition.int离院带药 = 0 Then
    ElseIf mcondition.int离院带药 = 1 Then
        str门诊 = str门诊 & " And Not Ltrim(To_Char(Nvl(C.扣率,0),'00')) Like '_3'"
    ElseIf mcondition.int离院带药 = 2 Then
        str门诊 = str门诊 & " And Ltrim(To_Char(Nvl(C.扣率,0),'00')) Like '_3'"
    End If
    
    str门诊 = str门诊 & IIf(mParams.strSourceDep = "", "", " And C.对方部门id+0=E.Column_Value ")
    
    If imgFilter.BorderStyle = cstFilter Then
        If Val(mParams.int输入模式) <= 7 Then
            If Trim(txtPati.Text) = "" Then
                str门诊 = str门诊 & " And 1 = 2 "
            End If
        Else
            If mSQLCondition.lng病人ID = 0 Then
                str门诊 = str门诊 & " And 1 = 2 "
            End If
        End If
    End If
    
    If mParams.intType = 1 Then
        str门诊 = str门诊 & " And (D.门诊标志=1 or D.门诊标志=4)"
    ElseIf mParams.intType = 2 Then
        str门诊 = str门诊 & " And (D.门诊标志<>1 and D.门诊标志<>4)"
    End If
    
    If mcondition.int服务对象 = 1 Then
        '门诊划价及门诊记帐
        gstrSQL = gstrSQL & str门诊
    Else
        If mcondition.int服务对象 = 3 Then
            '门诊及住院所有单据
            str住院 = Replace(str门诊, "门诊费用记录", "住院费用记录")
            str住院 = Replace(str住院, "And Nvl(D.费用状态,0)<>1", "")
        Else
            '住院记帐
            str住院 = Replace(str门诊, "门诊费用记录", "住院费用记录")
            str住院 = Replace(str住院, "And Nvl(D.费用状态,0)<>1", "")
            str门诊 = ""
        End If
    
        If mPrives.bln发病区处方 Then
            If img病区.BorderStyle = 0 Then
                '不显示病区处方
                str住院 = str住院 & " And (D.门诊标志 <> 2 Or (D.门诊标志 = 2 And D.病人病区id <> D.开单部门id)) "
            End If
            If img病区.BorderStyle = 1 And cbo病区.ListIndex <> -1 Then
                '要显示病区处方，并且病人病区等于当前选择的病区
                str住院 = str住院 & " And D.病人病区id = " & cbo病区.ItemData(cbo病区.ListIndex)
                str门诊 = ""
            End If
        End If
        
        If str门诊 = "" Then
            gstrSQL = gstrSQL & str住院
        Else
            gstrSQL = gstrSQL & str门诊 & " Union All " & str住院
        End If
    End If
    
    '''''Group By
    gstrSQL = gstrSQL & ") A GROUP BY A.优先级,A.类型,A.单据,A.已收费,A.配药人,A.No,A.姓名,A.日期,A.可操作,A.说明," & _
        " A.就诊卡号,A.门诊号,A.身份证号,A.IC卡号,A.病人ID,A.医保号,A.住院号,A.处方类型,A.门诊标志,A.记录性质, a.发药窗口,A.签到时间,a.病人类型, a.打印状态 " & IIf(mParams.bln启用审方, ",a.审查结果,a.审查id", "")
    
    '''''Order By
    gstrSQL = gstrSQL & " Order By A.签到时间,A.类型,A.单据,A.No"
    
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
            mSQLCondition.lng药房ID, _
            mSQLCondition.date开始日期, _
            mSQLCondition.date结束日期, _
            mSQLCondition.str开始NO, _
            mSQLCondition.str结束NO, _
            UCase(mSQLCondition.str姓名), _
            mSQLCondition.str就诊卡, _
            mSQLCondition.str标识号, _
            mSQLCondition.lng科室ID, _
            mSQLCondition.str填制人, _
            mSQLCondition.str审核人, _
            mSQLCondition.lng药品id, _
            mSQLCondition.str当前NO, _
            mSQLCondition.str门诊号, _
            lng病人ID, _
            mSQLCondition.lng病人ID, _
            mSQLCondition.str医保号, _
            mSQLCondition.lng住院号, _
            mstr窗口, _
            mParams.strSourceDep, _
            mstrDeptNode)
    
    stbThis.Panels(2) = ""
    
    Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Cancle, , True)
    Set cbrControl = Me.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Cancle, , True)
        
    If Not rsData.EOF Then
        cbrMenu.Enabled = True
        cbrControl.Enabled = True
        stbThis.Panels(2) = "共有" & rsData.RecordCount & "张处方；" & GetSumMoney(rsData)
    Else
        cbrMenu.Enabled = False
        cbrControl.Enabled = False
    End If
    
    Set mrsList = rsData
    
    If Not mfrmList Is Nothing Then mfrmList.RefreshList mListType.待配药, mrsList, strNo, blnNoRefreshDetail
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Function GetSumMoney(ByVal rsRecipt As ADODB.Recordset) As String
    Dim rstemp As ADODB.Recordset
    Dim dbl应收金额 As Double
    Dim dlb实收金额 As Double
    Set rstemp = rsRecipt.Clone
    
    With rstemp
        .MoveFirst
        Do While Not .EOF
            dbl应收金额 = dbl应收金额 + Val(.Fields("金额").Value)
            dlb实收金额 = dlb实收金额 + Val(.Fields("实收金额").Value)
            .MoveNext
        Loop
    End With
    
    If mParams.int金额显示 = 1 Then
        GetSumMoney = "实收金额：" & FormatEx(dlb实收金额, mintMoneyDigit) & "元"
    ElseIf mParams.int金额显示 = 2 Then
        GetSumMoney = "应收金额：" & FormatEx(dbl应收金额, mintMoneyDigit) & "元" & "  实收金额：" & FormatEx(dlb实收金额, mintMoneyDigit) & "元"
    Else
        GetSumMoney = "应收金额：" & FormatEx(dbl应收金额, mintMoneyDigit) & "元"
    End If
End Function
Public Sub RefreshList_Send(Optional ByVal strNo As String, Optional ByVal blnNoRefreshDetail As Boolean)
    '刷新待发药列表
    Dim bln医保号 As Boolean
    Dim rsData As ADODB.Recordset
    Dim strSub1 As String
    Dim strSub2 As String
    Dim str住院 As String
    Dim str门诊 As String
    Dim strInput As String
    Dim cbrControl As CommandBarControl
    Dim cbrMenu As CommandBarControl
    Dim lng病人ID As Long
    Dim strSqlTmp As String
    Dim str发生时间 As String
    
    On Error GoTo errHandle
    If mSQLCondition.str身份证 <> "" And Not mobjSquareCard Is Nothing Then
        If mobjSquareCard.zlGetPatiID("身份证", UCase(mSQLCondition.str身份证), False, lng病人ID) = False Then lng病人ID = 0
    End If
    
    gstrSQL = "Select '' As 颜色, 处方类型,'' As 选择 ,'0' As 标志,类型,单据,已收费,配药人,NO,姓名,呼叫时间,签到时间," & _
            " to_Char(Sum(Round(零售金额," & mintMoneyDigit & ")),'" & mstrOracleMoneyForamt & "') AS 金额,日期,可操作," & _
            " 说明,就诊卡号,门诊号,身份证号,IC卡号,病人ID,医保号,住院号,发药窗口," & _
            " Sum(Round(实收金额," & mintMoneyDigit & ")) 实收金额,门诊标志,记录性质,Zl_Get收费类别(单据,NO,[1]) As 收费类别,病人类型" & IIf(mParams.bln启用审方, ",审查结果,审查id", "") & _
            " From ("
            
    strSqlTmp = " Select A.呼叫时间,A.优先级,A.类型,A.单据,A.已收费,A.配药人,A.NO,A.姓名,C.零售金额,A.日期,A.可操作,A.说明,A.就诊卡号,A.门诊号,A.身份证号," & _
            " A.IC卡号,A.病人ID,A.医保号,A.住院号,d.实收金额*(Nvl(c.付数,1)*c.实际数量/(Nvl(d.付数,1)*d.数次)) 实收金额, Nvl(A.处方类型,Nvl(C.注册证号,0)) As 处方类型,D.门诊标志,D.记录性质,D.收费类别,A.发药窗口,c.签到时间,a.病人类型" & IIf(mParams.bln启用审方, ",a.审查结果,a.审查id", "") & _
            " From ("
    str门诊 = "Select distinct A.呼叫时间,B.就诊卡号,B.门诊号,B.身份证号,B.IC卡号,B.医保号,B.住院号,A.优先级,A.发药窗口,A.填制日期,Decode(Nvl(A.已收费,0),1,'','(未)')||Decode(A.单据,8,'收费',9,'记帐') 类型, " & _
            " A.单据,A.已收费,'' 配药人,A.No,A.姓名,To_Char(A.填制日期,'yyyy-MM-dd hh24:mi:ss') 日期,1 可操作,' ' 说明,B.病人ID, A.处方类型,a.对方部门id, b.病人类型" & IIf(mParams.bln启用审方, ",Q.审查结果,Q.id  审查id", "") & _
            " From 未发药品记录 A,病人信息 B,门诊费用记录 C " & IIf(mParams.bln启用审方, ",处方审查记录 Q,处方审查明细 K ", "") & _
            " Where 1=1 "
    
    str门诊 = str门诊 & IIf(mParams.bln启用审方, " and c.医嘱序号=k.医嘱id(+) and Q.id(+)=K.审方id and K.最后提交(+)=1 ", "")
    
    '是否有配药确认环节
    If mParams.blnMustDosageOkProcess = True And mParams.blnMustDosageProcess = True Then
        str门诊 = str门诊 & " and A.排队状态 in (2,3,4)"
    ElseIf mParams.blnMustDosageOkProcess = True And mParams.blnMustDosageProcess = False Then
        str门诊 = str门诊 & " and A.排队状态 in (1,2,3,4)"
    End If
    
    '主要条件
    If mParams.blnMustDosageProcess = True Then str门诊 = str门诊 & " And A.配药人 Is Not Null "
    
    str门诊 = str门诊 & " And (A.库房ID=[1] Or A.库房ID Is NULL) And A.填制日期 Between [2] And [3] "

    If mSQLCondition.str开始NO <> "" Or mSQLCondition.str结束NO <> "" Then
        If mSQLCondition.str开始NO <> "" And mSQLCondition.str结束NO <> "" Then
            str门诊 = str门诊 & " And A.NO Between [4] And [5] "
        Else
            If mSQLCondition.str开始NO <> "" Then
                str门诊 = str门诊 & " And A.NO = [4] "
            Else
                str门诊 = str门诊 & " And A.NO = [5] "
            End If
        End If
    End If
    
    If mSQLCondition.str姓名 <> "" Then str门诊 = str门诊 & " And Upper(A.姓名) Like Upper([6]) "
    
    If mSQLCondition.str就诊卡 <> "" Then str门诊 = str门诊 & " And Upper(B.就诊卡号) = [7] "
    
    If mSQLCondition.str标识号 <> "" Then str门诊 = str门诊 & " And Upper(DECODE(A.单据,8,B.门诊号,B.住院号)) Like [8] "
    
    If mSQLCondition.lng科室ID > 0 Then str门诊 = str门诊 & " And A.对方部门ID+0=[9] "
    
    If mSQLCondition.str当前NO <> "" Then str门诊 = str门诊 & " And A.NO=[13] "
    
    If mSQLCondition.str门诊号 <> "" Then str门诊 = str门诊 & " And B.门诊号=[14] "
    
'    If mSQLCondition.str身份证 <> "" Then str门诊 = str门诊 & " And B.身份证号=[15] "

    If mSQLCondition.str身份证 <> "" Then str门诊 = str门诊 & " And B.病人ID=[15] "
    
    If mSQLCondition.lng病人ID <> 0 Then str门诊 = str门诊 & " And B.病人ID=[16] "
    
    If mSQLCondition.str医保号 <> "" Then str门诊 = str门诊 & " And B.医保号=[17] "
    
    If mSQLCondition.lng住院号 <> 0 Then str门诊 = str门诊 & " And B.住院号=[18] "
    
    bln医保号 = (mSQLCondition.str医保号 <> "")
    
    str门诊 = str门诊 & " And A.病人ID=B.病人ID" & IIf(bln医保号 = True, "", "(+)") & ""
    
    str门诊 = str门诊 & IIf(mParams.Str窗口 = "", "", " And (A.发药窗口 In (Select * From Table(Cast(f_Str2list([19]) As Zltools.t_Strlist))) Or A.发药窗口 Is Null) ")
    
    '收费单据
    Select Case mParams.intShowBill收费
        Case 0  '不显示处方
            strSub1 = "1=2"
        Case 1  '显示未收费
            strSub1 = "(Nvl(A.已收费,0)=0 And A.单据=8)"
        Case 2  '显示已收费
            strSub1 = "(A.已收费=1 And A.单据=8)"
        Case 3  '显示所有处方
            strSub1 = "A.单据=8"
    End Select
    '记帐单据
    Select Case mParams.intShowBill记帐
        Case 0  '不显示处方
            strSub2 = "1=2"
        Case 1  '显示未审核
            strSub2 = "(Nvl(A.已收费,0)=0 And A.单据=9)"
        Case 2  '显示已审核
            strSub2 = "(A.已收费=1 And A.单据=9)"
        Case 3  '显示所有处方
            strSub2 = "A.单据=9"
    End Select
    
    str门诊 = str门诊 & " And (" & strSub1 & " Or " & strSub2 & ")"
    
    str门诊 = str门诊 & " And Mod(C.记录性质, 10) = Decode(A.单据, 8, 1, 2) And A.No = C.No And A.库房id = C.执行部门id "
    
    If mParams.bln发生时间过滤 = False Then
'        str门诊 = Replace(str门诊, ",门诊费用记录 C", "")
    Else
        str发生时间 = Replace(str门诊, "And A.填制日期 Between [2] And [3]", "")
        str门诊 = str门诊 & " And C.医嘱序号 Is Null "
        
        str发生时间 = str发生时间 & " And C.医嘱序号 Is Not Null And C.发生时间 Between [2] And [3] "
        str门诊 = str门诊 & " Union All " & str发生时间
    End If
    
    str门诊 = strSqlTmp & str门诊 & ") A,药品收发记录 C, 门诊费用记录 D, 部门表 B " & IIf(mParams.strSourceDep = "", "", ",Table(Cast(f_Num2List([20]) As zlTools.t_NumList)) E ") & _
              " Where C.费用id = D.ID And nvl(c.发药方式,-999)<>-1 and A.单据=C.单据 And A.NO=C.NO And C.审核人 Is NULL " & _
              " And Nvl(D.费用状态,0)<>1 And (C.库房id=[1] Or C.库房id Is null) And a.对方部门id = b.Id "
    
    If mstrDeptNode <> "" Then
        str门诊 = str门诊 & " And (b.站点 = [21] Or b.站点 Is Null) "
    End If
    
    '排除已经部分停止发药的No
    str门诊 = str门诊 & " and Not Exists(select 1 from 药品收发记录 F where F.单据=C.单据 and F.库房id=C.库房id and F.no=C.no and 发药方式=-1) "
    
    '排除已在输液配置中心管理中产生的单据
    str门诊 = str门诊 & " And Not Exists (Select 1 From 输液配药内容 Y Where Y.收发id = C.ID) "
    
    '是否显示退药待发单据
    If mcondition.bln显示退药待发单据 = False Then
        str门诊 = str门诊 & " And C.记录状态=1 "
    Else
        str门诊 = str门诊 & " And MOD(C.记录状态,3)=1 "
    End If
    
    '离院带药:'0-所有,1-不含离院带药,2-仅含离院带药
    If mcondition.int离院带药 = 0 Then
    ElseIf mcondition.int离院带药 = 1 Then
        str门诊 = str门诊 & " And Not Ltrim(To_Char(Nvl(C.扣率,0),'00')) Like '_3'"
    ElseIf mcondition.int离院带药 = 2 Then
        str门诊 = str门诊 & " And Ltrim(To_Char(Nvl(C.扣率,0),'00')) Like '_3'"
    End If
    
    str门诊 = str门诊 & IIf(mParams.strSourceDep = "", "", " And C.对方部门id+0=E.Column_Value ")
    
    If imgFilter.BorderStyle = cstFilter Then
        If Val(mParams.int输入模式) <= 7 Then
            If Trim(txtPati.Text) = "" Then
                str门诊 = str门诊 & " And 1 = 2 "
            End If
        Else
            If mSQLCondition.lng病人ID = 0 Then
                str门诊 = str门诊 & " And 1 = 2 "
            End If
        End If
    End If
    
    If mParams.intType = 1 Then
        str门诊 = str门诊 & " And (D.门诊标志=1 or D.门诊标志=4)"
    ElseIf mParams.intType = 2 Then
        str门诊 = str门诊 & " And (D.门诊标志<>1 and D.门诊标志<>4)"
    End If
    
    
    If mcondition.int服务对象 = 1 Then
        '门诊划价及门诊记帐
        gstrSQL = gstrSQL & str门诊
    Else
        If mcondition.int服务对象 = 3 Then
            '门诊及住院所有单据
            str住院 = Replace(str门诊, "门诊费用记录", "住院费用记录")
            str住院 = Replace(str住院, "And Nvl(D.费用状态,0)<>1", "")
        Else
            '住院记帐
            str住院 = Replace(str门诊, "门诊费用记录", "住院费用记录")
            str住院 = Replace(str住院, "And Nvl(D.费用状态,0)<>1", "")
            str门诊 = ""
        End If
    
        If mPrives.bln发病区处方 Then
            If img病区.BorderStyle = 0 Then
                '不显示病区处方
                str住院 = str住院 & " And (D.门诊标志 <> 2 Or (D.门诊标志 = 2 And D.病人病区id <> D.开单部门id)) "
            End If
            If img病区.BorderStyle = 1 And cbo病区.ListIndex <> -1 Then
                '要显示病区处方，并且病人病区等于当前选择的病区
                str住院 = str住院 & " And D.病人病区id = " & cbo病区.ItemData(cbo病区.ListIndex)
                str门诊 = ""
            End If
        End If
        
        If str门诊 = "" Then
            gstrSQL = gstrSQL & str住院
        Else
            gstrSQL = gstrSQL & str门诊 & " Union All " & str住院
        End If
    End If
    
    '''''Group By
    gstrSQL = gstrSQL & ") A GROUP BY A.呼叫时间,A.优先级,A.类型,A.单据,A.已收费,A.配药人,A.No,A.姓名,A.日期,A.可操作,A.说明," & _
        " A.就诊卡号,A.门诊号,A.身份证号,A.IC卡号,A.病人ID,A.医保号,A.住院号, A.处方类型,A.门诊标志,A.记录性质,a.发药窗口,A.签到时间,a.病人类型" & IIf(mParams.bln启用审方, ",a.审查结果,a.审查id", "")
    
    '''''Order By
    gstrSQL = gstrSQL & " Order By A.签到时间,A.类型,A.单据,A.No"
    
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
            mSQLCondition.lng药房ID, _
            mSQLCondition.date开始日期, _
            mSQLCondition.date结束日期, _
            mSQLCondition.str开始NO, _
            mSQLCondition.str结束NO, _
            UCase(mSQLCondition.str姓名), _
            mSQLCondition.str就诊卡, _
            mSQLCondition.str标识号, _
            mSQLCondition.lng科室ID, _
            mSQLCondition.str填制人, _
            mSQLCondition.str审核人, _
            mSQLCondition.lng药品id, _
            mSQLCondition.str当前NO, _
            mSQLCondition.str门诊号, _
            lng病人ID, _
            mSQLCondition.lng病人ID, _
            mSQLCondition.str医保号, _
            mSQLCondition.lng住院号, _
            mParams.Str窗口, _
            mParams.strSourceDep, _
            mstrDeptNode)
    
    stbThis.Panels(2) = ""
    
    Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Cancle, , True)
    Set cbrControl = Me.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Cancle, , True)
        
    If Not rsData.EOF Then
        cbrMenu.Enabled = True
        cbrControl.Enabled = True
        stbThis.Panels(2) = "共有" & rsData.RecordCount & "张处方；" & GetSumMoney(rsData)
    Else
        cbrMenu.Enabled = False
        cbrControl.Enabled = False
    End If

    Set mrsList = rsData
    If Not mfrmList Is Nothing Then
        '过滤出有记录就标记下
        mblnFinding = True
        
        If Val(mParams.int输入模式) <= 7 Then
            strInput = txtPati.Text
        Else
            '消费卡类别时输入为卡ID+卡号
            strInput = mobjcard.接口序号 & "|" & txtPati.Text
        End If
                
        mfrmList.ShowList mListType.待发药, imgFilter.BorderStyle, (mParams.blnStartCall And mParams.blnStartQueue), mParams.blnMustDosageOkProcess, mParams.blnMustDosageProcess, mParams.bln启用审方, IDKNType.GetCurCard.名称, strInput
        mfrmList.RefreshList mListType.待发药, mrsList, strNo, blnNoRefreshDetail
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub RefreshList_OverTime(Optional ByVal strNo As String, Optional ByVal blnNoRefreshDetail As Boolean)
    '刷新超时待发药列表
    Dim bln医保号 As Boolean
    Dim rsData As ADODB.Recordset
    Dim strSub1 As String
    Dim strSub2 As String
    Dim str住院 As String
    Dim str门诊 As String
    Dim strInput As String
    Dim lng病人ID As Long
    Dim strSqlTmp As String
    Dim str发生时间 As String
    
    On Error GoTo errHandle
    If mSQLCondition.str身份证 <> "" And Not mobjSquareCard Is Nothing Then
        If mobjSquareCard.zlGetPatiID("身份证", UCase(mSQLCondition.str身份证), False, lng病人ID) = False Then lng病人ID = 0
    End If
    
    gstrSQL = "Select '' As 颜色, 处方类型,'' As 选择 ,'0' As 标志,类型,单据,已收费,配药人,NO,姓名,签到时间," & _
            " to_Char(Sum(Round(零售金额," & mintMoneyDigit & ")),'" & mstrOracleMoneyForamt & "') AS 金额,日期,可操作," & _
            " 说明,就诊卡号,门诊号,身份证号,IC卡号,病人ID,医保号,住院号,发药窗口," & _
            " Sum(Round(实收金额," & mintMoneyDigit & ")) 实收金额,门诊标志,记录性质,Zl_Get收费类别(单据,NO,[1]) As 收费类别,病人类型 " & _
            " From ("
            
    strSqlTmp = " Select A.优先级,A.类型,A.单据,A.已收费,A.配药人,A.NO,A.姓名,C.零售金额,A.日期,A.可操作,A.说明,A.就诊卡号,A.门诊号,A.身份证号," & _
            " A.IC卡号,A.病人ID,A.医保号,A.住院号,d.实收金额*(Nvl(c.付数,1)*c.实际数量/(Nvl(d.付数,1)*d.数次)) 实收金额, Nvl(A.处方类型,Nvl(C.注册证号,0)) As 处方类型,D.门诊标志,D.记录性质,D.收费类别,A.发药窗口,c.签到时间,a.病人类型 " & _
            " From ( "
            
    str门诊 = "Select distinct B.就诊卡号,B.门诊号,B.身份证号,B.IC卡号,B.医保号,B.住院号,A.优先级,A.发药窗口,A.填制日期,Decode(Nvl(A.已收费,0),1,'','(未)')||Decode(A.单据,8,'收费',9,'记帐') 类型," & _
            " A.单据,A.已收费,'' 配药人,A.No,A.姓名,To_Char(A.填制日期,'yyyy-MM-dd hh24:mi:ss') 日期,1 可操作,' ' 说明,B.病人ID, A.处方类型,a.对方部门id, b.病人类型 " & _
            " From 未发药品记录 A,病人信息 B,门诊费用记录 C " & _
            " Where 1=1 "
    
    '主要条件
    If mParams.blnMustDosageProcess = True Then str门诊 = str门诊 & " And A.配药人 Is Not Null "
    
    str门诊 = str门诊 & " And (A.库房ID=[1] Or A.库房ID Is NULL) And A.填制日期 Between [2] And [3] "
    
    str门诊 = str门诊 & " And A.填制日期 < Sysdate - (1 / 24 / 60) * [22] "

    If mSQLCondition.str开始NO <> "" Or mSQLCondition.str结束NO <> "" Then
        If mSQLCondition.str开始NO <> "" And mSQLCondition.str结束NO <> "" Then
            str门诊 = str门诊 & " And A.NO Between [4] And [5] "
        Else
            If mSQLCondition.str开始NO <> "" Then
                str门诊 = str门诊 & " And A.NO = [4] "
            Else
                str门诊 = str门诊 & " And A.NO = [5] "
            End If
        End If
    End If
    
    If mSQLCondition.str姓名 <> "" Then str门诊 = str门诊 & " And Upper(A.姓名) Like Upper([6]) "
    
    If mSQLCondition.str就诊卡 <> "" Then str门诊 = str门诊 & " And Upper(B.就诊卡号) = [7] "
    
    If mSQLCondition.str标识号 <> "" Then str门诊 = str门诊 & " And Upper(DECODE(A.单据,8,B.门诊号,B.住院号)) Like [8] "
    
    If mSQLCondition.lng科室ID > 0 Then str门诊 = str门诊 & " And A.对方部门ID+0=[9] "
    
    If mSQLCondition.str当前NO <> "" Then str门诊 = str门诊 & " And A.NO=[13] "
    
    If mSQLCondition.str门诊号 <> "" Then str门诊 = str门诊 & " And B.门诊号=[14] "
    
'    If mSQLCondition.str身份证 <> "" Then str门诊 = str门诊 & " And B.身份证号=[15] "

    If mSQLCondition.str身份证 <> "" Then str门诊 = str门诊 & " And B.病人ID=[15] "
    
    If mSQLCondition.lng病人ID <> 0 Then str门诊 = str门诊 & " And B.病人ID=[16] "
    
    If mSQLCondition.str医保号 <> "" Then str门诊 = str门诊 & " And B.医保号=[17] "
    
    If mSQLCondition.lng住院号 <> 0 Then str门诊 = str门诊 & " And B.住院号=[18] "
    
    bln医保号 = (mSQLCondition.str医保号 <> "")
    str门诊 = str门诊 & " And A.病人ID=B.病人ID" & IIf(bln医保号 = True, "", "(+)") & ""
    
    str门诊 = str门诊 & IIf(mParams.Str窗口 = "", "", " And (A.发药窗口 In (Select * From Table(Cast(f_Str2list([19]) As Zltools.t_Strlist))) Or A.发药窗口 Is Null) ")
    
    '收费单据
    Select Case mParams.intShowBill收费
        Case 0  '不显示处方
            strSub1 = "1=2"
        Case 1  '显示未收费
            strSub1 = "(Nvl(A.已收费,0)=0 And A.单据=8)"
        Case 2  '显示已收费
            strSub1 = "(A.已收费=1 And A.单据=8)"
        Case 3  '显示所有处方
            strSub1 = "A.单据=8"
    End Select
    '记帐单据
    Select Case mParams.intShowBill记帐
        Case 0  '不显示处方
            strSub2 = "1=2"
        Case 1  '显示未审核
            strSub2 = "(Nvl(A.已收费,0)=0 And A.单据=9)"
        Case 2  '显示已审核
            strSub2 = "(A.已收费=1 And A.单据=9)"
        Case 3  '显示所有处方
            strSub2 = "A.单据=9"
    End Select
    
    str门诊 = str门诊 & " And (" & strSub1 & " Or " & strSub2 & ")"
    
    str门诊 = str门诊 & " And Mod(C.记录性质, 10) = Decode(A.单据, 8, 1, 2) And A.No = C.No And A.库房id = C.执行部门id "
              
    If mParams.bln发生时间过滤 = False Then
'        str门诊 = Replace(str门诊, ",门诊费用记录 C", "")
    Else
        str发生时间 = Replace(str门诊, "And A.填制日期 Between [2] And [3]", "")
        str门诊 = str门诊 & " And C.医嘱序号 Is Null "
        
        str发生时间 = str发生时间 & " And C.医嘱序号 Is Not Null And C.发生时间 Between [2] And [3] "
        str门诊 = str门诊 & " Union All " & str发生时间
    End If
    
    str门诊 = strSqlTmp & str门诊 & ") A,药品收发记录 C, 门诊费用记录 D, 部门表 B " & IIf(mParams.strSourceDep = "", "", ",Table(Cast(f_Num2List([20]) As zlTools.t_NumList)) E ") & _
              " Where C.费用id = D.ID And nvl(c.发药方式,-999)<>-1 and A.单据=C.单据 And A.NO=C.NO And C.审核人 Is NULL " & _
              " And Nvl(D.费用状态,0)<>1 And (C.库房id=[1] Or C.库房id Is null) And a.对方部门id = b.Id "
    
    If mstrDeptNode <> "" Then
        str门诊 = str门诊 & " And (b.站点 = [21] Or b.站点 Is Null) "
    End If
    
    '排除已经部分停止发药的No
    str门诊 = str门诊 & " and Not Exists(select 1 from 药品收发记录 F where F.单据=C.单据 and F.库房id=C.库房id and F.no=C.no and 发药方式=-1) "
    
    
    '排除已在输液配置中心管理中产生的单据
    str门诊 = str门诊 & " And Not Exists (Select 1 From 输液配药内容 Y Where Y.收发id = C.ID) "
    
    '是否显示退药待发单据
    If mcondition.bln显示退药待发单据 = False Then
        str门诊 = str门诊 & " And C.记录状态=1 "
    Else
        str门诊 = str门诊 & " And MOD(C.记录状态,3)=1 "
    End If
    
    '离院带药:'0-所有,1-不含离院带药,2-仅含离院带药
    If mcondition.int离院带药 = 0 Then
    ElseIf mcondition.int离院带药 = 1 Then
        str门诊 = str门诊 & " And Not Ltrim(To_Char(Nvl(C.扣率,0),'00')) Like '_3'"
    ElseIf mcondition.int离院带药 = 2 Then
        str门诊 = str门诊 & " And Ltrim(To_Char(Nvl(C.扣率,0),'00')) Like '_3'"
    End If
    
    str门诊 = str门诊 & IIf(mParams.strSourceDep = "", "", " And C.对方部门id+0=E.Column_Value ")
    
    If imgFilter.BorderStyle = cstFilter Then
        If Val(mParams.int输入模式) <= 7 Then
            If Trim(txtPati.Text) = "" Then
                str门诊 = str门诊 & " And 1 = 2 "
            End If
        Else
            If mSQLCondition.lng病人ID = 0 Then
                str门诊 = str门诊 & " And 1 = 2 "
            End If
        End If
    End If
    
    If mParams.intType = 1 Then
        str门诊 = str门诊 & " And (D.门诊标志=1 or D.门诊标志=4)"
    ElseIf mParams.intType = 2 Then
        str门诊 = str门诊 & " And (D.门诊标志<>1 and D.门诊标志<>4)"
    End If
    
    If mcondition.int服务对象 = 1 Then
        '门诊划价及门诊记帐
        gstrSQL = gstrSQL & str门诊
    Else
        If mcondition.int服务对象 = 3 Then
            '门诊及住院所有单据
            str住院 = Replace(str门诊, "门诊费用记录", "住院费用记录")
            str住院 = Replace(str住院, "And Nvl(D.费用状态,0)<>1", "")
        Else
            '住院记帐
            str住院 = Replace(str门诊, "门诊费用记录", "住院费用记录")
            str住院 = Replace(str住院, "And Nvl(D.费用状态,0)<>1", "")
            str门诊 = ""
        End If
    
        If mPrives.bln发病区处方 Then
            If img病区.BorderStyle = 0 Then
                '不显示病区处方
                str住院 = str住院 & " And (D.门诊标志 <> 2 Or (D.门诊标志 = 2 And D.病人病区id <> D.开单部门id)) "
            End If
            If img病区.BorderStyle = 1 And cbo病区.ListIndex <> -1 Then
                '要显示病区处方，并且病人病区等于当前选择的病区
                str住院 = str住院 & " And D.病人病区id = " & cbo病区.ItemData(cbo病区.ListIndex)
                str门诊 = ""
            End If
        End If
        
        If str门诊 = "" Then
            gstrSQL = gstrSQL & str住院
        Else
            gstrSQL = gstrSQL & str门诊 & " Union All " & str住院
        End If
    End If
    
    '''''Group By
    gstrSQL = gstrSQL & ") A GROUP BY A.优先级,A.类型,A.单据,A.已收费,A.配药人,A.No,A.姓名,A.日期,A.可操作,A.说明," & _
        " A.就诊卡号,A.门诊号,A.身份证号,A.IC卡号,A.病人ID,A.医保号,A.住院号,A.处方类型,A.门诊标志,A.记录性质,a.发药窗口,A.签到时间,a.病人类型 "
    
    '''''Order By
    gstrSQL = gstrSQL & " Order By A.签到时间,A.类型,A.单据,A.No"
    
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
            mSQLCondition.lng药房ID, _
            mSQLCondition.date开始日期, _
            mSQLCondition.date结束日期, _
            mSQLCondition.str开始NO, _
            mSQLCondition.str结束NO, _
            UCase(mSQLCondition.str姓名), _
            mSQLCondition.str就诊卡, _
            mSQLCondition.str标识号, _
            mSQLCondition.lng科室ID, _
            mSQLCondition.str填制人, _
            mSQLCondition.str审核人, _
            mSQLCondition.lng药品id, _
            mSQLCondition.str当前NO, _
            mSQLCondition.str门诊号, _
            lng病人ID, _
            mSQLCondition.lng病人ID, _
            mSQLCondition.str医保号, _
            mSQLCondition.lng住院号, _
            mParams.Str窗口, _
            mParams.strSourceDep, _
            mstrDeptNode, _
            mSQLCondition.intOverTime)
    
    stbThis.Panels(2) = ""
    If Not rsData.EOF Then
        stbThis.Panels(2) = "超过开单时间" & mParams.intOverTime & "分钟未发药处方：共有" & rsData.RecordCount & "张；" & GetSumMoney(rsData)
    Else
        stbThis.Panels(2) = "超过开单时间" & mParams.intOverTime & "分钟未发药处方：共有0张"
    End If

    Set mrsList = rsData
    If Not mfrmList Is Nothing Then
        If Val(mParams.int输入模式) <= 7 Then
            strInput = txtPati.Text
        Else
            '消费卡类别时输入为卡ID+卡号
            strInput = mobjcard.接口序号 & "|" & txtPati.Text
        End If
        
        mfrmList.ShowList mListType.超时未发, imgFilter.BorderStyle, (mParams.blnStartCall And mParams.blnStartQueue), mParams.blnMustDosageOkProcess, mParams.blnMustDosageProcess, mParams.bln启用审方, IDKNType.GetCurCard.名称, strInput
        mfrmList.RefreshList mListType.超时未发, mrsList, strNo, blnNoRefreshDetail
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub




Private Sub Load病区()
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    If cbo病区.ListCount > 0 And mstrDeptNode = cbo病区.Tag Then Exit Sub
    
    '病区
    gstrSQL = " Select 编码||'-'||名称 科室,ID From 部门表 " & _
             " Where ID in (Select 部门ID From 部门性质说明 Where 工作性质='护理' And 服务对象 IN(2,3))" & _
             " And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','yyyy-MM-dd')) "
    
    If mstrDeptNode <> "" Then
        gstrSQL = gstrSQL & " And (站点 = [1] Or 站点 Is Null) "
    End If
    
    gstrSQL = gstrSQL & " Order By 编码||'-'||名称 "
    
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "取所有病区", mstrDeptNode)
    
    With cbo病区
        .Clear
        .Tag = mstrDeptNode
        Do While Not rsTmp.EOF
            .AddItem rsTmp!科室
            .ItemData(.NewIndex) = rsTmp!Id
            rsTmp.MoveNext
        Loop
        If .ListIndex <> -1 Then
            .ListIndex = 0
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Public Sub RefreshList_Return(Optional ByVal strNo As String, Optional ByVal blnNoRefreshDetail As Boolean)
    '刷新退药列表
    Dim rsData As ADODB.Recordset
    Dim strSqlSendType As String
    Dim strSqlSourceDep As String
    Dim strSql单据类型 As String
    Dim strSqlFilter As String
    Dim strSqlSub As String
    Dim strSql医保号 As String
    Dim strSub1 As String
    Dim strSub2 As String
    Dim bln医保号 As Boolean
    Dim strGroup As String
    Dim str门诊 As String
    Dim str住院 As String
    Dim strSql病区 As String
    Dim bln不显示门诊 As Boolean
    Dim lng病人ID As Long
    
    On Error GoTo errHandle
    
    If mSQLCondition.str身份证 <> "" And Not mobjSquareCard Is Nothing Then
        If mobjSquareCard.zlGetPatiID("身份证", UCase(mSQLCondition.str身份证), False, lng病人ID) = False Then lng病人ID = 0
    End If
    
    ''strCond1
    If mSQLCondition.str开始NO <> "" Or mSQLCondition.str结束NO <> "" Then
        If mSQLCondition.str开始NO <> "" And mSQLCondition.str结束NO <> "" Then
            strSqlSub = strSqlSub & " And A.NO Between [4] And [5] "
        Else
            If mSQLCondition.str开始NO <> "" Then
                strSqlSub = strSqlSub & " And A.NO = [4] "
            Else
                strSqlSub = strSqlSub & " And A.NO = [5] "
            End If
        End If
    End If

    If mcondition.int服务对象 = 2 Then
        strSql单据类型 = " And A.单据 = 9 "
    Else
        strSql单据类型 = " And A.单据 In (8,9)"
    End If
    
    If mSQLCondition.str姓名 <> "" Then strSqlSub = strSqlSub & " And Upper(H.姓名) Like [6] "
    
    If mSQLCondition.str就诊卡 <> "" Then strSqlSub = strSqlSub & " And Upper(B.就诊卡号) = [7] "
    
    If mSQLCondition.str标识号 <> "" Then strSqlSub = strSqlSub & " And Upper(DECODE(A.单据,8,B.门诊号,B.住院号)) Like [8] "
    
    If mSQLCondition.lng科室ID > 0 Then strSqlSub = strSqlSub & " And A.对方部门ID+0=[9] "
    
    If mSQLCondition.str填制人 <> "" Then strSqlSub = strSqlSub & " And A.填制人=[10] "
    
    If mSQLCondition.str审核人 <> "" Then strSqlSub = strSqlSub & " And A.审核人=[11] "
    
    If mSQLCondition.lng药品id > 0 Then strSqlSub = strSqlSub & " And A.药品ID+0=[12] "
    
    If mSQLCondition.str当前NO <> "" Then strSqlSub = strSqlSub & " And A.NO=[13] "
    
    If mSQLCondition.str门诊号 <> "" Then strSqlSub = strSqlSub & " And B.门诊号=[14] "
    
'    If mSQLCondition.str身份证 <> "" Then strSqlSub = strSqlSub & " And B.身份证号=[15] "
    
    If mSQLCondition.str身份证 <> "" Then strSqlSub = strSqlSub & " And B.病人ID=[15] "
    
    If mSQLCondition.lng病人ID <> 0 Then strSqlSub = strSqlSub & " And B.病人ID=[16] "
    
    If mSQLCondition.str医保号 <> "" Then strSqlSub = strSqlSub & " And B.医保号=[17] "
    
    If mSQLCondition.lng住院号 <> 0 Then strSqlSub = strSqlSub & " And B.住院号=[18] "
    
    ''strSql医保号
    bln医保号 = (mSQLCondition.str医保号 <> "")
    strSql医保号 = " AND H.病人ID=B.病人ID" & IIf(bln医保号 = True, "", "(+)") & ""
    
    ''strSqlSendType
    '离院带药:'0-所有,1-不含离院带药,2-仅含离院带药
    If mcondition.int离院带药 = 0 Then
    ElseIf mcondition.int离院带药 = 1 Then
        strSqlSendType = " And Not Ltrim(To_Char(Nvl(A.扣率,0),'00')) Like '_3'"
    ElseIf mcondition.int离院带药 = 2 Then
        strSqlSendType = " And Ltrim(To_Char(Nvl(A.扣率,0),'00')) Like '_3'"
    End If
    
    ''strSqlSourceDep
    strSqlSourceDep = IIf(mParams.strSourceDep = "", "", " And A.对方部门id In (Select * From Table(Cast(f_Num2list([19]) As Zltools.t_Numlist))) ")

    ''strSqlFilter
    If imgFilter.BorderStyle = cstFilter Then
        If Val(mParams.int输入模式) <= 7 Then
            If Trim(txtPati.Text) = "" Then
                strSqlFilter = " And 1 = 2 "
            End If
        Else
            If mSQLCondition.lng病人ID = 0 Then
                strSqlFilter = " And 1 = 2 "
            End If
        End If
    End If
    
    ''病区发药
    If mPrives.bln发病区处方 Then
        If img病区.BorderStyle = 0 Then
            '不显示病区处方
            strSql病区 = " And (H.门诊标志 <> 2 Or (H.门诊标志 = 2 And H.病人病区id <> H.开单部门id)) "
        End If
        If img病区.BorderStyle = 1 And cbo病区.ListIndex <> -1 Then
            '要显示病区处方，并且病人病区等于当前选择的病区
            strSql病区 = " And H.病人病区id = " & cbo病区.ItemData(cbo病区.ListIndex)
            bln不显示门诊 = True
        End If
    End If
    
    '收费单据
    Select Case mParams.intShowBill收费
        Case 0  '不显示处方
            strSub1 = "1=2"
        Case 3  '显示所有处方
            strSub1 = "A.单据=8"
    End Select
    '记帐单据
    Select Case mParams.intShowBill记帐
        Case 0  '不显示处方
            strSub2 = "1=2"
        Case 3  '显示所有处方
            strSub2 = "A.单据=9"
    End Select
    
    '收费单据
    Select Case mParams.intShowBill收费
        Case 0  '不显示处方
            strSub1 = "1=2"
        Case 1  '显示未收费
            strSub1 = "(Nvl(H.记录状态,0)=0 And A.单据=8)"
        Case 2  '显示已收费
            strSub1 = "(H.记录状态>=1 And A.单据=8)"
        Case 3  '显示所有处方
            strSub1 = "A.单据=8"
    End Select
    '记帐单据
    Select Case mParams.intShowBill记帐
        Case 0  '不显示处方
            strSub2 = "1=2"
        Case 1  '显示未审核
            strSub2 = "(Nvl(H.记录状态,0)=0 And A.单据=9)"
        Case 2  '显示已审核
            strSub2 = "(H.记录状态>=1 And A.单据=9)"
        Case 3  '显示所有处方
            strSub2 = "A.单据=9"
    End Select
    
    strSqlSub = strSqlSub & " And (" & strSub1 & " Or " & strSub2 & ")"
    
    '针对任何一张药品处方，不会存在一部分明细分别在线与后备中存在的情况，因此，可直接通过在线UNION后备的方式解决
    '由于费用记录在最外层，而且无主要条件，通过费用记录的在线与后备联接后，其效果是全表扫描，因此，只能通过整个在线SQL UNION 整个后备SQL的方式解决
    If mcondition.bln显示过程单据 = False Then
        gstrSQL = " SELECT DISTINCT '' As 颜色, A.处方类型,'' As 选择,'0' As 标志,Decode(Nvl(h.记录状态, 0),  0,'(未)','') || Decode(a.单据, 8, '收费', 9, '记帐') 类型," & _
                 "      A.单据,1 已收费,A.审核人 配药人,A.NO,H.姓名,trim(to_char(sum(A.零售金额),'" & mstrOracleMoneyForamt & "')) AS 金额,trim(to_char(Sum((Nvl(a.付数, 1) * a.实际数量)/(Nvl(H.付数,1)*H.数次) * H.实收金额),'" & mstrOracleMoneyForamt & "')) As 实收金额," & _
                 "      TO_CHAR(A.审核日期,'YYYY-MM-DD HH24:MI:SS') 日期,1 可操作,' ' 说明,B.就诊卡号,B.门诊号,B.身份证号,B.IC卡号,B.病人ID,B.医保号,B.住院号,H.门诊标志, H.记录性质,Zl_Get收费类别(A.单据,A.NO,[1]) As 收费类别,B.病人类型,A.未取药 " & _
                 " FROM " & _
                 "      (SELECT A.ID,A.NO,A.单据,A.药品ID,A.费用ID,A.批次,A.批号,A.效期," & _
                 "          NVL(A.付数,1) 付数,A.实际数量,NVL(A.付数,1)*A.实际数量-B.已发数量 已退数量,B.已发数量,A.记录状态,A.发药窗口," & _
                 "          A.零售价,round(a.零售价*Nvl(a.付数, 1)*a.实际数量," & mintMoneyDigit & ") 零售金额,A.单量,A.频次,A.用法,A.摘要,A.审核人,A.审核日期,A.对方部门ID,A.库房ID, A.填制人, A.处方类型,A.未取药 " & _
                 "      FROM" & _
                 "          (SELECT A.ID,A.NO,A.单据,A.药品ID,A.序号,A.费用ID,A.批次,A.批号,A.效期,A.付数,A.实际数量,A.记录状态,A.发药窗口,A.零售价,A.单量,A.频次,A.用法,A.摘要,A.审核人,A.审核日期,A.对方部门ID,A.库房ID, A.填制人, Nvl(A.注册证号, 0) As 处方类型,Nvl(A.是否未取药,0) As 未取药 " & _
                 "          FROM 药品收发记录 A" & _
                 "          WHERE nvl(A.发药方式,-999)<>-1 and A.审核人 IS NOT NULL AND (A.记录状态=1 OR MOD(A.记录状态,3)=0)" & _
                 "          AND A.库房ID+0=[1] And A.审核日期 Between [2] And [3]  " & strSql单据类型 & strSqlSendType & _
                 "          And Not Exists (Select 1 From 输液配药内容 Y,药品收发记录 Z Where y.收发id=Z.ID AND Z.NO= A.NO And z.单据=a.单据 And z.库房id = a.库房id) " & _
                 "          ) A," & _
                 "          (SELECT A.NO,A.单据,A.药品ID,A.序号,SUM(NVL(A.付数,1)*A.实际数量) 已发数量,SUM(A.零售金额) 零售金额" & _
                 "          FROM 药品收发记录 A" & _
                 "          WHERE nvl(A.发药方式,-999)<>-1 and A.审核人 IS NOT NULL" & strSql单据类型 & strSqlSendType & _
                 "          AND A.库房ID+0=[1] And A.审核日期 Between [2] And [3]  " & strSqlSourceDep & _
                 "          GROUP BY A.NO,A.单据,A.药品ID,A.序号) B"
        gstrSQL = gstrSQL & _
                 "      WHERE A.NO = B.NO AND A.单据 = B.单据 AND A.药品ID+0 = B.药品ID AND A.序号 = B.序号 AND B.已发数量<>0" & _
                 "     ) A,门诊费用记录 H,病人信息 B" & _
                 " WHERE A.库房ID+0=[1] " & _
                 " " & strSqlSub & strSqlFilter & strSql医保号 & _
                 " AND (A.记录状态=1 OR MOD(A.记录状态,3)=0) AND A.审核人 IS NOT NULL AND A.费用ID=H.ID AND A.实际数量<>0 "
    Else
        gstrSQL = " SELECT DISTINCT '' As 颜色, A.处方类型,'' As 选择,'0' As 标志,Decode(Nvl(h.记录状态, 0),  0,'(未)','') || Decode(a.单据, 8, '收费', 9, '记帐') 类型,A.单据,1 已收费,A.审核人 配药人," & _
                  "      A.NO,H.姓名,trim(to_char(sum(A.零售金额),'" & mstrOracleMoneyForamt & "')) AS 金额,trim(to_char(Sum((Nvl(a.付数, 1) * a.实际数量)/(Nvl(H.付数,1)*H.数次) * H.实收金额),'" & mstrOracleMoneyForamt & "')) As 实收金额,TO_CHAR(A.审核日期,'YYYY-MM-DD HH24:MI:SS') 日期,A.可操作," & _
                  "      DECODE(A.记录状态,1,'第1次发药',DECODE(MOD(A.记录状态,3),0,'第1次发药',1,'第'||(FLOOR(A.记录状态/3)+1)||'次发药',2,'第'||(FLOOR(A.记录状态/3)+1)||'次退药')) 说明,B.就诊卡号,B.门诊号,B.身份证号,B.IC卡号,B.病人ID,B.医保号,B.住院号,H.门诊标志, H.记录性质,Zl_Get收费类别(A.单据,A.NO,[1]) As 收费类别,B.病人类型,A.未取药 " & _
                  " FROM " & _
                  "      (SELECT * FROM" & _
                  "          (SELECT A.ID,A.NO,A.单据,A.药品ID,A.费用ID,A.批次,A.批号,A.效期," & _
                  "              NVL(A.付数,1) 付数,A.实际数量,NVL(A.付数,1)*A.实际数量-B.已发数量 已退数量,B.已发数量,A.记录状态,A.发药窗口," & _
                  "              A.零售价 , round(a.零售价*Nvl(a.付数, 1)*a.实际数量," & mintMoneyDigit & ") 零售金额, A.单量, A.频次, A.用法, A.摘要, A.审核人, A.审核日期, A.对方部门ID, A.库房ID,1 可操作, A.填制人, A.处方类型,A.未取药 " & _
                  "          FROM" & _
                  "              (SELECT A.ID,A.NO,A.单据,A.药品ID,A.序号,A.费用ID,A.批次,A.批号,A.效期,A.付数,A.实际数量,A.记录状态,A.发药窗口,A.零售价,A.零售金额, A.单量, A.频次, A.用法, A.摘要, A.审核人, A.审核日期, A.对方部门ID, A.库房ID, A.填制人, Nvl(A.注册证号, 0) As 处方类型,Nvl(A.是否未取药,0) As 未取药 " & _
                  "              FROM 药品收发记录 A" & _
                  "              WHERE nvl(a.发药方式,-999)<>-1 and A.审核人 IS NOT NULL AND (A.记录状态=1 OR MOD(A.记录状态,3)=0)" & _
                  "              AND A.库房ID+0=[1] And A.审核日期 Between [2] And [3]  " & strSql单据类型 & strSqlSendType & _
                  "              And Not Exists (Select 1 From 输液配药内容 Y,药品收发记录 Z Where y.收发id=Z.ID AND Z.NO=A.NO And z.单据=a.单据 And z.库房id = a.库房id)  " & _
                  "              ) A," & _
                  "              (SELECT A.NO,A.单据,A.药品ID,A.序号,SUM(NVL(A.付数,1)*A.实际数量) 已发数量" & _
                  "              FROM 药品收发记录 A" & _
                  "              WHERE nvl(a.发药方式,-999)<>-1 and A.审核人 IS NOT NULL " & strSql单据类型 & strSqlSendType & _
                  "              AND A.库房ID+0=[1] And A.审核日期 Between [2] And [3]  " & strSqlSourceDep & _
                  "              GROUP BY A.NO,A.单据,A.药品ID,A.序号) B"
         gstrSQL = gstrSQL & _
                  "          WHERE A.NO = B.NO AND A.单据 = B.单据 AND A.药品ID+0 = B.药品ID AND A.序号 = B.序号)" & _
                  "          UNION" & _
                  "          SELECT A.ID,A.NO,A.单据,A.药品ID,A.费用ID,A.批次,A.批号,A.效期," & _
                  "          NVL(A.付数,1) 付数,A.实际数量,0 已退数,0 已发数量,A.记录状态,A.发药窗口," & _
                  "          A.零售价 , round(A.零售金额," & mintMoneyDigit & ") 零售金额, A.单量, A.频次, A.用法, A.摘要, A.审核人, A.审核日期, A.对方部门ID, A.库房ID," & _
                  "          DECODE(记录状态,1,1,DECODE(MOD(记录状态,3),0,1,MOD(记录状态,3)+1)) 可操作, A.填制人, Nvl(A.注册证号, 0) As 处方类型,Nvl(A.是否未取药,0) As 未取药 " & _
                  "          FROM 药品收发记录 A" & _
                  "          WHERE nvl(a.发药方式,-999)<>-1 and Not Exists (Select 1 From 输液配药内容 Y,药品收发记录 Z Where y.收发id=Z.ID AND  Z.NO= A.NO And z.单据=a.单据 And z.库房id = a.库房id) and NOT (记录状态=1 OR MOD(记录状态,3)=0) And A.审核日期 Between [2] And [3]  " & strSql单据类型 & strSqlSendType & strSqlSourceDep
         gstrSQL = gstrSQL & _
                  "     ) A,门诊费用记录 H,病人信息 B" & _
                  " WHERE A.库房ID+0=[1] " & _
                  " " & strSqlSub & strSqlFilter & strSql医保号 & _
                  " AND A.审核人 IS NOT NULL AND A.费用ID=H.ID "
    End If
    
    'Group
    If mcondition.bln显示过程单据 = False Then
        strGroup = " GROUP BY A.处方类型,Decode(Nvl(h.记录状态, 0),  0,'(未)','') || Decode(a.单据, 8, '收费', 9, '记帐'),A.单据,1,A.审核人,A.NO,H.姓名," & _
            " TO_CHAR(A.审核日期,'YYYY-MM-DD HH24:MI:SS'),B.就诊卡号,B.门诊号,B.身份证号,B.IC卡号,B.病人ID,B.医保号,B.住院号,H.门诊标志, H.记录性质,B.病人类型,A.未取药 "
    Else
        strGroup = " GROUP BY A.处方类型,Decode(Nvl(h.记录状态, 0),  0,'(未)','') || Decode(a.单据, 8, '收费', 9, '记帐') ,A.单据,1,A.审核人," & _
            " A.NO,H.姓名,TO_CHAR(A.审核日期,'YYYY-MM-DD HH24:MI:SS'),A.可操作," & _
            " DECODE(A.记录状态,1,'第1次发药',DECODE(MOD(A.记录状态,3),0,'第1次发药',1,'第'||(FLOOR(A.记录状态/3)+1)||'次发药',2,'第'||(FLOOR(A.记录状态/3)+1)||'次退药')),B.就诊卡号,B.门诊号,B.身份证号,B.IC卡号,B.病人ID,B.医保号,B.住院号,H.门诊标志, H.记录性质,B.病人类型,A.未取药 "
    End If
    
    
    
    If mParams.intType = 1 Then
        gstrSQL = gstrSQL & " And (H.门诊标志=1 or H.门诊标志=4)"
    ElseIf mParams.intType = 2 Then
        gstrSQL = gstrSQL & " And (H.门诊标志<>1 and H.门诊标志<>4)"
    End If
    
    '区分门诊、住院
    If mcondition.int服务对象 = 1 Then
        '门诊划价及门诊记帐
        gstrSQL = gstrSQL & strGroup
    Else
        If mcondition.int服务对象 = 3 Then
            '门诊及住院所有单据
            str门诊 = gstrSQL
            str住院 = Replace(str门诊, "门诊费用记录", "住院费用记录")
            
            str门诊 = str门诊 & strGroup
            str住院 = str住院 & strSql病区 & strGroup
            
            If bln不显示门诊 = True Then
                gstrSQL = str住院
            Else
                gstrSQL = str门诊 & " Union All " & str住院
            End If
        Else
            '住院记帐
            str住院 = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
            str住院 = str住院 & strSql病区 & strGroup
            gstrSQL = str住院
        End If
    End If
     
    'order by
    gstrSQL = gstrSQL & " order by 类型,单据,NO "
     
    Dim blnMoved As Boolean
    Dim str开始日期 As String, strsql As String
    
    str开始日期 = Format(mSQLCondition.date开始日期, "yyyy-mm-dd hh:mm:ss")
    
    '判断从开始日期后，是否存在转出的处方数据
    blnMoved = Sys.IsMovedByDate(str开始日期)
    
    '如果存在数据转出，则需要同时从后备表中提取数据
    If blnMoved Then
        strsql = gstrSQL
        strsql = Replace(strsql, "药品收发记录", "H药品收发记录")
        strsql = Replace(strsql, "门诊费用记录", "H门诊费用记录")
        strsql = Replace(strsql, "住院费用记录", "H住院费用记录")
        gstrSQL = gstrSQL & " UNION ALL " & strsql
    End If
     
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
            mSQLCondition.lng药房ID, _
            mSQLCondition.date开始日期, _
            mSQLCondition.date结束日期, _
            mSQLCondition.str开始NO, _
            mSQLCondition.str结束NO, _
            UCase(mSQLCondition.str姓名), _
            mSQLCondition.str就诊卡, _
            mSQLCondition.str标识号, _
            mSQLCondition.lng科室ID, _
            mSQLCondition.str填制人, _
            mSQLCondition.str审核人, _
            mSQLCondition.lng药品id, _
            mSQLCondition.str当前NO, _
            mSQLCondition.str门诊号, _
            lng病人ID, _
            mSQLCondition.lng病人ID, _
            mSQLCondition.str医保号, _
            mSQLCondition.lng住院号, _
            mParams.strSourceDep)

    stbThis.Panels(2) = ""
    If Not rsData.EOF Then
        stbThis.Panels(2) = "共有" & rsData.RecordCount & "张处方；" & GetSumMoney(rsData)
    End If
    
    Set mrsList = rsData
    If Not mrsList Is Nothing Then mfrmList.RefreshList mListType.退药, mrsList, strNo, blnNoRefreshDetail
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub RefreshList_Abolish()
    '刷新取消配药列表
    Dim bln医保号 As Boolean
    Dim rsData As ADODB.Recordset
    Dim strSub1 As String
    Dim strSub2 As String
    Dim str住院 As String
    Dim str门诊 As String
    Dim lng病人ID As Long
    Dim strSqlTmp As String
    Dim str发生时间 As String
    
    On Error GoTo errHandle
    If mSQLCondition.str身份证 <> "" And Not mobjSquareCard Is Nothing Then
        If mobjSquareCard.zlGetPatiID("身份证", UCase(mSQLCondition.str身份证), False, lng病人ID) = False Then lng病人ID = 0
    End If
    
    gstrSQL = "Select '' As 颜色, 处方类型,'' As 选择 ,'0' As 标志,类型,单据,已收费,配药人,NO,姓名,签到时间," & _
            " to_Char(Sum(Round(零售金额," & mintMoneyDigit & ")),'" & mstrOracleMoneyForamt & "') AS 金额,日期," & _
            " 可操作,说明,就诊卡号,门诊号,身份证号,IC卡号,病人ID,医保号,住院号,发药窗口," & _
            " Sum(Round(实收金额," & mintMoneyDigit & ")) 实收金额,门诊标志,记录性质,Zl_Get收费类别(单据,NO,[1]) As 收费类别,病人类型 " & _
            " From ("
            
    strSqlTmp = " Select A.优先级,A.类型,A.单据,A.已收费,A.配药人,A.NO,A.姓名,C.零售金额,A.日期,A.可操作,A.说明,A.就诊卡号,A.门诊号,A.身份证号," & _
            " A.IC卡号,A.病人ID,A.医保号,A.住院号,d.实收金额, Nvl(A.处方类型,Nvl(C.注册证号,0)) As 处方类型,D.门诊标志,D.记录性质,D.收费类别,A.发药窗口,c.签到时间,a.病人类型 " & _
            " From ("
            
    str门诊 = "Select distinct B.就诊卡号,B.门诊号,B.身份证号,B.IC卡号,B.医保号,B.住院号,A.优先级,A.发药窗口,A.填制日期,Decode(Nvl(A.已收费,0),1,'','(未)')||Decode(A.单据,8,'收费',9,'记帐') 类型," & _
            " A.单据,A.已收费,'' 配药人,A.No,A.姓名,To_Char(A.填制日期,'yyyy-MM-dd hh24:mi:ss') 日期,1 可操作,' ' 说明,B.病人ID, A.处方类型,a.对方部门id, b.病人类型 " & _
            "  From 未发药品记录 A,病人信息 B,门诊费用记录 C " & _
            "  Where A.配药人 Is Not Null "
    
    '主要条件
    str门诊 = str门诊 & " And (A.库房ID=[1] Or A.库房ID Is NULL) And A.填制日期 Between [2] And [3] "
    
    If mSQLCondition.str开始NO <> "" Or mSQLCondition.str结束NO <> "" Then
        If mSQLCondition.str开始NO <> "" And mSQLCondition.str结束NO <> "" Then
            str门诊 = str门诊 & " And A.NO Between [4] And [5] "
        Else
            If mSQLCondition.str开始NO <> "" Then
                str门诊 = str门诊 & " And A.NO = [4] "
            Else
                str门诊 = str门诊 & " And A.NO = [5] "
            End If
        End If
    End If
    
    If mSQLCondition.str姓名 <> "" Then str门诊 = str门诊 & " And Upper(A.姓名) Like [6] "
    
    If mSQLCondition.str就诊卡 <> "" Then str门诊 = str门诊 & " And Upper(B.就诊卡号) = [7] "
    
    If mSQLCondition.str标识号 <> "" Then str门诊 = str门诊 & " And Upper(DECODE(A.单据,8,B.门诊号,B.住院号)) Like [8] "
    
    If mSQLCondition.lng科室ID > 0 Then str门诊 = str门诊 & " And A.对方部门ID+0=[9] "
    
    If mSQLCondition.str当前NO <> "" Then str门诊 = str门诊 & " And A.NO=[13] "
    
    If mSQLCondition.str门诊号 <> "" Then str门诊 = str门诊 & " And B.门诊号=[14] "
    
'    If mSQLCondition.str身份证 <> "" Then str门诊 = str门诊 & " And B.身份证号=[15] "
    
    If mSQLCondition.str身份证 <> "" Then str门诊 = str门诊 & " And B.病人ID=[15] "
    
    If mSQLCondition.lng病人ID <> 0 Then str门诊 = str门诊 & " And B.病人ID=[16] "
    
    If mSQLCondition.str医保号 <> "" Then str门诊 = str门诊 & " And B.医保号=[17] "
    
    If mSQLCondition.lng住院号 <> 0 Then str门诊 = str门诊 & " And B.住院号=[18] "
    
            
    bln医保号 = (mSQLCondition.str医保号 <> "")
    str门诊 = str门诊 & " And A.病人ID=B.病人ID" & IIf(bln医保号 = True, "", "(+)") & ""
    
    str门诊 = str门诊 & IIf(mParams.Str窗口 = "", "", " And (A.发药窗口 In (Select * From Table(Cast(f_Str2list([19]) As Zltools.t_Strlist))) Or A.发药窗口 Is Null) ")
    
    '设置显示及自动打印的条件:注意"未发药品记录"的别名为A
    Select Case mParams.intShowBill收费
        Case 0  '不显示处方
            strSub1 = "1=2"
        Case 1  '显示未收费
            strSub1 = "A.单据<>9 And Nvl(A.已收费,0)=0 And A.单据=8"
        Case 2  '显示已收费
            strSub1 = "A.单据<>9 And A.已收费=1 And A.单据=8"
        Case 3  '显示所有处方
            strSub1 = "A.单据<>9 And A.单据=8"
    End Select
    Select Case mParams.intShowBill记帐
        Case 0  '不显示处方
            strSub2 = "1=2"
        Case 1  '显示未审核
            strSub2 = "A.单据<>8 And Nvl(A.已收费,0)=0 And A.单据=9"
        Case 2  '显示已审核
            strSub2 = "A.单据<>8 And A.已收费=1 And A.单据=9"
        Case 3  '显示所有处方
            strSub2 = "A.单据<>8 And A.单据=9"
    End Select
    
    str门诊 = str门诊 & " And A.单据 IN(8,9) And (" & strSub1 & " Or " & strSub2 & ")"
    
    str门诊 = str门诊 & " And Mod(C.记录性质, 10) = Decode(A.单据, 8, 1, 2) And A.No = C.No And A.库房id = C.执行部门id "
            
    If mParams.bln发生时间过滤 = False Then
'        str门诊 = Replace(str门诊, ",门诊费用记录 C", "")
    Else
        str发生时间 = Replace(str门诊, "And A.填制日期 Between [2] And [3]", "")
        str门诊 = str门诊 & " And C.医嘱序号 Is Null "
        
        str发生时间 = str发生时间 & " And C.医嘱序号 Is Not Null And C.发生时间 Between [2] And [3] "
        str门诊 = str门诊 & " Union All " & str发生时间
    End If
    
    str门诊 = strSqlTmp & str门诊 & ") A,药品收发记录 C, 门诊费用记录 D, 部门表 B " & IIf(mParams.strSourceDep = "", "", ",Table(Cast(f_Num2List([20]) As zlTools.t_NumList)) E ") & _
              " Where C.费用id = D.ID And nvl(c.发药方式,-999)<>-1 and A.单据=C.单据 And A.NO=C.NO And C.审核人 Is NULL " & _
              " And Nvl(D.费用状态,0)<>1 And (C.库房id=[1] Or C.库房id Is null) And a.对方部门id = b.Id "
    
    If mstrDeptNode <> "" Then
        str门诊 = str门诊 & " And (b.站点 = [21] Or b.站点 Is Null) "
    End If
    
    '排除已在输液配置中心管理中产生的单据
    str门诊 = str门诊 & " And Not Exists (Select 1 From 输液配药内容 Y Where Y.收发id = C.ID) "
    
    '是否显示退药待发单据
    If mcondition.bln显示退药待发单据 = False Then
        str门诊 = str门诊 & " And C.记录状态=1 "
    Else
        str门诊 = str门诊 & " And MOD(C.记录状态,3)=1 "
    End If
    
    '离院带药:'0-所有,1-不含离院带药,2-仅含离院带药
    If mcondition.int离院带药 = 0 Then
    ElseIf mcondition.int离院带药 = 1 Then
        str门诊 = str门诊 & " And Not Ltrim(To_Char(Nvl(C.扣率,0),'00')) Like '_3'"
    ElseIf mcondition.int离院带药 = 2 Then
        str门诊 = str门诊 & " And Ltrim(To_Char(Nvl(C.扣率,0),'00')) Like '_3'"
    End If
    
    str门诊 = str门诊 & IIf(mParams.strSourceDep = "", "", " And C.对方部门id+0=E.Column_Value ")
    
    If imgFilter.BorderStyle = cstFilter Then
        If Val(mParams.int输入模式) <= 7 Then
            If Trim(txtPati.Text) = "" Then
                str门诊 = str门诊 & " And 1 = 2 "
            End If
        Else
            If mSQLCondition.lng病人ID = 0 Then
                str门诊 = str门诊 & " And 1 = 2 "
            End If
        End If
    End If
    
    If mParams.intType = 1 Then
        str门诊 = str门诊 & " And (D.门诊标志=1 or D.门诊标志=4)"
    ElseIf mParams.intType = 2 Then
        str门诊 = str门诊 & " And (D.门诊标志<>1 and D.门诊标志<>4)"
    End If
    
    
    If mcondition.int服务对象 = 1 Then
        '门诊划价及门诊记帐
        gstrSQL = gstrSQL & str门诊
    Else
        If mcondition.int服务对象 = 3 Then
            '门诊及住院所有单据
            str住院 = Replace(str门诊, "门诊费用记录", "住院费用记录")
            str住院 = Replace(str住院, "And Nvl(D.费用状态,0)<>1", "")
        Else
            '住院记帐
            str住院 = Replace(str门诊, "门诊费用记录", "住院费用记录")
            str住院 = Replace(str住院, "And Nvl(D.费用状态,0)<>1", "")
            str门诊 = ""
        End If
    
        If mPrives.bln发病区处方 Then
            If img病区.BorderStyle = 0 Then
                '不显示病区处方
                str住院 = str住院 & " And (D.门诊标志 <> 2 Or (D.门诊标志 = 2 And D.病人病区id <> D.开单部门id)) "
            End If
            If img病区.BorderStyle = 1 And cbo病区.ListIndex <> -1 Then
                '要显示病区处方，并且病人病区等于当前选择的病区
                str住院 = str住院 & " And D.病人病区id = " & cbo病区.ItemData(cbo病区.ListIndex)
                str门诊 = ""
            End If
        End If
        
        If str门诊 = "" Then
            gstrSQL = gstrSQL & str住院
        Else
            gstrSQL = gstrSQL & str门诊 & " Union All " & str住院
        End If
    End If
    
    '''''Group By
    gstrSQL = gstrSQL & ") A GROUP BY A.优先级,A.类型,A.单据,A.已收费,A.配药人,A.No,A.姓名,A.日期,A.可操作,A.说明," & _
        " A.就诊卡号,A.门诊号,A.身份证号,A.IC卡号,A.病人ID,A.医保号,A.住院号,A.处方类型,A.门诊标志,A.记录性质, a.发药窗口,A.签到时间,a.病人类型 "
    
    '''''Order By
    gstrSQL = gstrSQL & " Order By A.签到时间,A.类型,A.单据,A.No"
    
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
            mSQLCondition.lng药房ID, _
            mSQLCondition.date开始日期, _
            mSQLCondition.date结束日期, _
            mSQLCondition.str开始NO, _
            mSQLCondition.str结束NO, _
            UCase(mSQLCondition.str姓名), _
            mSQLCondition.str就诊卡, _
            mSQLCondition.str标识号, _
            mSQLCondition.lng科室ID, _
            mSQLCondition.str填制人, _
            mSQLCondition.str审核人, _
            mSQLCondition.lng药品id, _
            mSQLCondition.str当前NO, _
            mSQLCondition.str门诊号, _
            lng病人ID, _
            mSQLCondition.lng病人ID, _
            mSQLCondition.str医保号, _
            mSQLCondition.lng住院号, _
            mParams.Str窗口, _
            mParams.strSourceDep, _
            mstrDeptNode)
    
    stbThis.Panels(2) = ""
    If Not rsData.EOF Then
        stbThis.Panels(2) = "共有" & rsData.RecordCount & "张处方；" & GetSumMoney(rsData)
    End If
    
    Set mrsList = rsData
    If Not mfrmList Is Nothing Then mfrmList.RefreshList mListType.已配药, mrsList
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ShowWindow_Batch()
    '调用批量发药窗口
    
    With Frm药品批量发药
        .In_服务对象 = mcondition.int服务对象
        .In_发药窗口 = mParams.Str窗口
        .In_药房ID = mParams.lng药房ID
        .In_库存检查 = mParams.IntCheckStock
        .In_校验处方 = IIf(mPrives.bln校验处方, 1, 0)
        .In_允许未配药发药 = IIf(mParams.blnMustDosageProcess = False, 1, 0)
        .IN_允许未审核发药 = IIf(mParams.bln允许未审核处方发药, 1, 0)
        .IN_允许未收费发药 = IIf(mParams.bln允许未收费处方发药, 1, 0)
        .In_权限 = mstrPrivs
        .str配药人 = IIf(tbcDetail.Selected.index = 1, mfrmRecipe.Get配药人, mfrmDetail.Get配药人)
        .str核查人 = IIf(tbcDetail.Selected.index = 1, mfrmRecipe.Get核查人, mfrmDetail.Get核查人)
        .In_金额保留位数 = mParams.int金额保留位数
        .IN_审核划价单 = IIf(mParams.bln审核划价单, 1, 0)
        .In_发其他药房处方 = False
        .In_窗口 = mstrOpr
        .In_自动发药 = mblnPackerConnect
        .In_启用发药 = mblnLoadDrug
        Set .In_DrugMAC = mobjDrugMAC
        Set .In_PlugIn = mobjPlugIn
        .Show 1, Me
    End With
    
    RefreshList mcondition.intListType
End Sub

Private Sub ShowWindow_Charge()
    '门诊划价
    Dim strCommon As String, intAtom As Integer, blnOK As Boolean
            
    On Error Resume Next
    If gobjCharge Is Nothing Then
        Set gobjCharge = CreateObject("zl9OutExse.clsOutExse")
        If gobjCharge Is Nothing Then Exit Sub
    End If
    
    err.Clear: On Error GoTo 0
    
    '部件调用合法性设置
    strCommon = Format(Now, "yyyyMMddHHmm")
    strCommon = TranPasswd(strCommon) & "||" & AnalyseComputer
    intAtom = GlobalAddAtom(strCommon)
    Call SaveSetting("ZLSOFT", "公共全局", "公共", intAtom)
    blnOK = gobjCharge.Charge(Me, gcnOracle, glngSys, gstrDbUser, 1, 0)
    Call GlobalDeleteAtom(intAtom)
    
    '完成划价
    '刷新未发药处方
    RefreshList mcondition.intListType
End Sub

Private Sub ShowWindow_EMR()
    '病案查询
End Sub

Private Sub ShowWindow_Flag()
    '停止发药标记
    Dim frmFlag As New Frm不再发药处方标志
    
    frmFlag.In_库存检查 = mParams.IntCheckStock
    frmFlag.gstrParentName = Me.Name
    frmFlag.Show vbModal
    
    RefreshList mcondition.intListType
End Sub

Private Sub ShowWindow_ReturnBatch()
    '退其它药房的处方
    
    frm批量退药.In_权限 = mstrPrivs
    Set frm批量退药.In_PlugIn = mobjPlugIn
    If Not frm批量退药.ShowEditor(Me, mParams.lng药房ID, True, mParams.int金额保留位数) Then Exit Sub
    
    RefreshList mcondition.intListType
End Sub

Private Sub ShowWindow_ReturnByBill()
    '按票据号退药
    
    frm按票据号批量退药.In_权限 = mstrPrivs
    Set frm按票据号批量退药.In_PlugIn = mobjPlugIn
    If Not frm按票据号批量退药.ShowEditor(Me, mParams.lng药房ID, mParams.int金额保留位数) Then Exit Sub
    
    RefreshList mcondition.intListType
End Sub

Private Sub ShowWindow_SendByBill()
    '按票据号发药
    
    With Frm按票据号批量发药
        .In_服务对象 = mcondition.int服务对象
        .In_发药窗口 = mParams.Str窗口
        .In_药房ID = mParams.lng药房ID
        .In_库存检查 = mParams.IntCheckStock
        .In_校验处方 = IIf(mPrives.bln校验处方, 1, 0)
        .In_允许未配药发药 = IIf(mParams.blnMustDosageProcess = False, 1, 0)
        .IN_允许未审核发药 = IIf(mParams.bln允许未审核处方发药, 1, 0)
        .IN_允许未收费发药 = IIf(mParams.bln允许未收费处方发药, 1, 0)
        .In_权限 = mstrPrivs
        .str配药人 = IIf(mParams.str配药人 = "|当前操作员|", gstrUserName, mParams.str配药人)
        .In_金额保留位数 = mParams.int金额保留位数
        .IN_审核划价单 = IIf(mParams.bln审核划价单, 1, 0)
        .In_窗口 = mstrOpr
        Set .In_DrugMAC = mobjDrugMAC
        Set .In_PlugIn = mobjPlugIn
        .Show 1, Me
    End With
    RefreshList mcondition.intListType
End Sub

Private Sub ShowWindow_SendOther()
    '发其它药房的处方
    
    With Frm药品批量发药
        .In_服务对象 = mcondition.int服务对象
        .In_发药窗口 = mParams.Str窗口
        .In_药房ID = mParams.lng药房ID
        .In_库存检查 = mParams.IntCheckStock
        .In_校验处方 = IIf(mPrives.bln校验处方, 1, 0)
        .In_允许未配药发药 = IIf(mParams.blnMustDosageProcess = False, 1, 0)
        .IN_允许未审核发药 = IIf(mParams.bln允许未审核处方发药, 1, 0)
        .IN_允许未收费发药 = IIf(mParams.bln允许未收费处方发药, 1, 0)
        .In_权限 = mstrPrivs
        .str配药人 = IIf(tbcDetail.Selected.index = 1, mfrmRecipe.Get配药人, mfrmDetail.Get配药人)
        .str核查人 = IIf(tbcDetail.Selected.index = 1, mfrmRecipe.Get核查人, mfrmDetail.Get核查人)
        .In_金额保留位数 = mParams.int金额保留位数
        .IN_审核划价单 = IIf(mParams.bln审核划价单, 1, 0)
        .In_发其他药房处方 = True
        .In_自动发药 = mblnPackerConnect
        .In_启用发药 = mblnLoadDrug
        .Show 1, Me
    End With
    
    RefreshList mcondition.intListType
End Sub

Private Sub ShowWindow_Stuff()
    '卫材发料
    Dim strCommon As String, intAtom As Integer, blnOK As Boolean
    Dim str当前处方 As String
    Dim strNo As String
    Dim lng病人ID As Long
    Dim rsTmp As ADODB.Recordset
    
    str当前处方 = mfrmList.GetCurrentRecipe
    
    If str当前处方 <> "" Then
        strNo = Split(str当前处方, "|")(1)
        lng病人ID = Val(Split(str当前处方, "|")(3))
    End If
    
    On Error Resume Next
    If gobjStuff Is Nothing Then
        Set gobjStuff = CreateObject("zl9Stuff.clsStuff")
        If gobjStuff Is Nothing Then Exit Sub
    End If

    err.Clear: On Error GoTo 0

    '部件调用合法性设置
    strCommon = Format(Now, "yyyyMMddHHmm")
    strCommon = TranPasswd(strCommon) & "||" & AnalyseComputer
    intAtom = GlobalAddAtom(strCommon)
    Call SaveSetting("ZLSOFT", "公共全局", "公共", intAtom)
    Call gobjStuff.TransStuff(Me, gcnOracle, glngSys, gstrDbUser, lng病人ID, strNo, mParams.lng药房ID, Format(mSQLCondition.date开始日期, "yyyy-mm-dd hh:mm:ss"), Format(mSQLCondition.date结束日期, "yyyy-mm-dd hh:mm:ss"))
    Call GlobalDeleteAtom(intAtom)
End Sub

Private Sub cbo病区_Click()
    If cbo病区.ListIndex = -1 Then Exit Sub
    If cbo病区.Enabled = False Then Exit Sub
    
    If cbo病区.ItemData(cbo病区.ListIndex) <> Val(cbo病区.Tag) Then
        cbo病区.Tag = cbo病区.ItemData(cbo病区.ListIndex)
        Call RefreshList(mcondition.intListType)
    End If
End Sub


Private Sub cbo时间范围_Click()
    With cbo时间范围
        If .ListIndex <> Val(.Tag) Then
            If (Val(.Tag) = 3 And .ListIndex < 3) Or (Val(.Tag) < 3 And .ListIndex = 3) Then
                Call picConMain_Resize
                Call picCondition_Resize
            End If
            .Tag = .ListIndex
        End If
        
        If .ListIndex < mTimeRange.指定时间范围 And mblnStart = True Then
            RefreshList mcondition.intListType
        End If
    End With
End Sub


Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim cbrControl As CommandBarControl
    Dim strReturn As String
    
    Select Case Control.Id
        '''''文件
        Case mconMenu_File_PrintSet     '打印设置
            zlPrintSet
        Case mconMenu_File_Preview      '打印预览
            zlSubPrint 2
        Case mconMenu_File_Print        '打印
            zlSubPrint 1
        Case mconMenu_File_Excel        '输出到Excel
            zlSubPrint 3
        
        Case mconMenu_File_Recipe_BillPrintDosage       '打印配药单
            Call BillPrint_Dosage
        Case mconMenu_File_Recipe_BillPrintRecipe       '打印处方签
            Call BillPrint_Recipe
        Case mconMenu_File_Recipe_BillPrintReport       '打印发药清单
            Call BillPrint_Report
        Case mconMenu_File_Recipe_BillPrintReturn       '打印退药通知单
            Call BillPrint_Return
        Case mconMenu_File_Recipe_BillPrintLable        '打印药品标签
            Call BillPrint_Lable
        Case mconMenu_File_Recipe_BillPrintBack         '打印退费单据
            Call BillPrint_Back
        Case mconMenu_File_Recipe_BillPrintChange        '打印医嘱更改通知单
            Call BillPrint_Change
        Case mconMenu_File_Parameter                    '参数设置
            ResetParams
            
        Case mconMenu_File_Exit                         '退出
            Unload Me
        
        '''''编辑
        Case mconMenu_Edit_Recipe_Batch                 '批量发药(&B)
            Call ShowWindow_Batch
        Case mconMenu_Edit_Recipe_SendOther             '发其它药房的处方(&F)
            Call ShowWindow_SendOther
        Case mconMenu_Edit_Recipe_ReturnBatch           '退其它药房的处方(&T)
            Call ShowWindow_ReturnBatch
        Case mconMenu_Edit_Recipe_SendByBill            '按票据号发药(&I)
            Call ShowWindow_SendByBill
        Case mconMenu_Edit_Recipe_ReturnByBill          '按票据号退药(&R)
            Call ShowWindow_ReturnByBill
        Case mconMenu_Edit_Recipe_Flag                  '停止发药标记(&S)
            Call ShowWindow_Flag
        Case mconMenu_Edit_Recipe_Charge                '门诊划价(&M)-F8
            Call ShowWindow_Charge
        Case mconMenu_Edit_Recipe_Stuff                 '卫材发料(@W)-F9
            Call ShowWindow_Stuff
        Case mconMenu_Edit_Recipe_TakeDrug              '取药确认(&T)
            Call RecipeWork_TakeDrug
        Case mconMenu_Edit_Recipe_Call                  '呼叫
            Call RecipeWork_Call
        Case mconMenu_Edit_Recipe_Cancle                '取消确认
            If Control.Caption = "取消确认" Then
                Call RecipeWork_DosageOk
            Else
                Call RecipeWork_Abolish
            End If
            
            Call RefreshList(mcondition.intListType)
        Case mconMenu_Edit_PlugIn + 1 To mconMenu_Edit_PlugIn + 99 '外挂发药业务功能调用
            DrugSendRecipeNormal Control.Parameter
        Case mconMenu_Edit_Recipe_Change                '切换配药人(&E)
            Call ChangeDosagePeople
        Case mconMenu_Edit_Recipe_Windows               '调整窗口
            Call ChangWin
        Case mconMenu_Edit_Recipe_EMR                   '病案查询
            Call ShowWindow_EMR
        Case mconMenu_Edit_Recipe_SendHot               '发药快捷键操作-F2
            If tbcDetail.Selected.index = 0 Then
                mfrmDetail.CmdProcess
            ElseIf tbcDetail.Item(1).Visible = True Then
                mfrmRecipe.CmdProcess
            End If
        
        '''''查看
        Case mconMenu_View_ToolBar_Button               '标准按钮
            Control.Checked = Not Control.Checked
            Me.cbsMain(2).Visible = Control.Checked
            Me.cbsMain.RecalcLayout
        Case mconMenu_View_ToolBar_Text                 '文本标签
            Control.Checked = Not Control.Checked
            For Each cbrControl In Me.cbsMain(2).Controls
                cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            Me.cbsMain.RecalcLayout
        Case mconMenu_View_ToolBar_Size                 '大图标
            Control.Checked = Not Control.Checked
            Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
            Me.cbsMain.RecalcLayout
        Case mconMenu_View_StatusBar                    '状态栏
            Me.stbThis.Visible = Not Me.stbThis.Visible
            Me.cbsMain.RecalcLayout
        Case mconMenu_View_FontSize_1, mconMenu_View_FontSize_2, mconMenu_View_FontSize_3                   '字号设置
            mParams.intFont = Val(Control.Parameter)
            Call SetFontSize
            Call zldatabase.SetPara("字体", mParams.intFont, glngSys, 1341)
        
        Case mconMenu_View_Filter                       '数据过滤
            Call ResetFilter
        Case mconMenu_View_Refresh                      '刷新
            Call RefreshList(mcondition.intListType)
        Case mconMenu_Edit_Recipe_VerifySign            '验证电子签名
            VerifySign
        Case mconMenu_Edit_Recipe_AutoSend_Open
            '启用处方上传
            Control.Checked = Not Control.Checked
            mblnLoadDrug = Control.Checked
        Case mconMenu_Edit_Recipe_AutoSend_Set
            mblnPackerConnect = mobjDrugMAC.DYEY_MZ_SetServer
            SetComandBars
        Case mconMenu_Edit_Recipe_AutoSend_LoadDrug
            Call mobjDrugMAC.DYEY_MZ_TransDrug(1, UserInfo.用户编码, UserInfo.用户姓名, strReturn)
        Case mconMenu_Edit_Recipe_AutoSend_LoadStock
            Call mobjDrugMAC.DYEY_MZ_TransStock(Val(mstrOpr), UserInfo.用户编码, UserInfo.用户姓名, mParams.lng药房ID, strReturn)
            
        '''''帮助
        Case mconMenu_Help_Help                         '帮助
'            Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
            Call ShowHelp(App.ProductName, Me.hWnd, "Frm药品发药管理")
        Case mconMenu_Help_Web                          'WEB上的中联
        Case mconMenu_Help_Web_Home                     '中联主页
            Call zlHomePage(Me.hWnd)
        Case mconMenu_Help_Web_Forum                    '中联论坛
            Call zlWebForum(Me.hWnd)
        Case mconMenu_Help_Web_Mail                     '发送反馈
            Call zlMailTo(Me.hWnd)
        Case mconMenu_Help_About                        '关于
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case mconMenu_Edit_Recipe_MedicalRecord         '电子病案查阅
            Call ShowMedicalRecord(mfrmDetail.GetRecord)
     
        ''''特殊热键
'        Case mconMenu_Edit_Recipe_Hot_IC
'            If mParams.int输入模式 = mFindType.IC卡 Then
'                Call cmdIC_Click
'            End If
            
        Case Else
            If Control.Id > 401 And Control.Id < 499 Then
                '执行自定义报表
                Call BillPrint_Custom(Control)
            End If
            
            '弹出菜单
'            If Control.Id >= mconMenu_Input_Recipe_NO And Control.Id <= mconMenu_Input_Recipe_NO + 6 + mintCardCount Then
'                Call SetInputPopupCheck(Control)            '输入项目弹出菜单
'            End If
            
'            '药房自动发药接口菜单
'            If Control.Id > mconMenu_AutoSend And Control.Id < mconMenu_AutoSend + 10 Then
'                gobjPackerMZ.SetInterface Control.Id - mconMenu_AutoSend - 1, mParams.lng药房ID
'            End If
    End Select
End Sub

Private Sub DrugSendRecipeNormal(ByVal strFunName As String)
    Dim str当前处方 As String, Int单据 As Integer, strNo As String
    
    If Not mobjPlugIn Is Nothing Then
        str当前处方 = mfrmList.GetCurrentRecipe
        
        If str当前处方 <> "" Then
            Int单据 = Val(Split(str当前处方, "|")(0))
            strNo = Split(str当前处方, "|")(1)
        End If
        
        On Error Resume Next
        Call mobjPlugIn.DrugSendWorkNormal(glngModul, strFunName, mParams.lng药房ID, strNo, Int单据)
        err.Clear: On Error GoTo 0
    End If
    
End Sub

Private Function RecipeWork_Call() As Boolean
    '呼叫
    Dim str当前处方 As String
    Dim Int单据 As Integer
    Dim strNo As String
    Dim Str窗口 As String
    Dim strName As String
    Dim strCall As String
    Dim strMsg As String
    
    On Error GoTo ErrHand
    
    str当前处方 = mfrmList.GetCurrentRecipe
    
    If str当前处方 <> "" Then
        Int单据 = Val(Split(str当前处方, "|")(0))
        strNo = Split(str当前处方, "|")(1)
        strName = Split(str当前处方, "|")(8)
        Str窗口 = Split(str当前处方, "|")(9)
        strMsg = Split(str当前处方, "|")(10)
    End If
     
    Call mfrmList.SetCalling
    
    strCall = "请、" & strName & "、" & strName & "、" & "、到" & mstr窗口
        
    gstrSQL = "Zl_未发药品记录_呼叫("
            'NO
            gstrSQL = gstrSQL & "'" & strNo & "'"
            '单据
            gstrSQL = gstrSQL & "," & Int单据
            '药房id
            gstrSQL = gstrSQL & "," & mParams.lng药房ID
            '发药窗口
            gstrSQL = gstrSQL & ",'" & Str窗口 & "'"
            '呼叫内容
            gstrSQL = gstrSQL & ",'" & strCall & "'"
            gstrSQL = gstrSQL & ")"
            
    Call zldatabase.ExecuteProcedure(gstrSQL, "RecipeWork_Call")
    
    '刷新显示输出队列
    If mParams.blnShowQueue = True Then
        If Not gobjLEDShow Is Nothing Then
            Call gobjLEDShow.zlDrugShow(mParams.lng药房ID, mParams.Str窗口, mParams.blnMustDosageProcess, mParams.blnMustDosageOkProcess, strName)
        End If
    End If
    
    '如果启用了本地语音系统，立即播放
    If mParams.blnStartQueue = True Then
        If mParams.blnStartCall = True And mParams.intCallType = 0 And mQueue.blnCallOver = True Then
            Call zlCallMain
        End If
    End If
    
    RecipeWork_Call = True
    
    '呼叫同时通知设备准备发药
    If mParams.blnDispensing Then
        Call DrugDispensing("" & Int单据 & "," & strNo)
    End If
    
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub DrugDispensing(ByVal strNos As String)
'功能：通知设备准备发药

    Dim strReturn As String
    
    On Error GoTo hErr
    
    If UCase(TypeName(mobjDrugMAC)) = UCase("clsDrugPacker") Then
        If mcondition.intListType = mListType.待发药 And mblnPackerConnect And mintAutoSendFlow = 1 _
            And strNos <> "" And mblnCompatible Then
            If mobjDrugMAC.DYEY_MZ_TransRecipeList(mstrOpr, UserInfo.用户编码, UserInfo.用户姓名, _
                mParams.lng药房ID, strNos, strReturn, mSendOper.StartSend) = False Then
                Call MsgBox("药品自动化设备系统未准备好，通知发药开始失败！", vbInformation, gstrSysName)
            End If
        End If
    ElseIf UCase(TypeName(mobjDrugMAC)) = UCase("clsDrugMachine") Then
        If mcondition.intListType = mListType.待发药 And mblnPackerConnect Then
            Call mobjDrugMAC.Operation(gstrDbUser, Val("22-开始发药"), "1|" & Replace(strNos, "|", ";"), strReturn)
        End If
    End If
    
    Exit Sub
    
hErr:
    Call ErrCenter
End Sub

Private Sub SetFontSize()
    Dim intFont As Integer
    Dim stdfnt As StdFont
    
    Select Case mParams.intFont
        Case 0
            intFont = 9
        Case 1
            intFont = 11
        Case 2
            intFont = 15
        Case Else
            intFont = 9
    End Select
    
    mfrmList.SetFontSize intFont
    mfrmDetail.SetFontSize intFont
    
    If Not tbcList.PaintManager.Font Is Nothing Then
        With tbcList
            Set stdfnt = .PaintManager.Font
            stdfnt.Size = intFont
             Set .PaintManager.Font = stdfnt
              .PaintManager.Layout = xtpTabLayoutAutoSize
        End With
    End If
    
    If Not tbcDetail.PaintManager.Font Is Nothing Then
        With tbcDetail
            Set stdfnt = .PaintManager.Font
            stdfnt.Size = intFont
             Set .PaintManager.Font = stdfnt
              .PaintManager.Layout = xtpTabLayoutAutoSize
        End With
    End If
    Me.FontSize = intFont
End Sub
Private Sub zlSubPrint(ByVal bytMode As Byte)
    'bytMode：1-打印；2-预览；3-输出到Excel
    Dim ObjThis As Object
    Dim objPrint As New zlPrint1Grd
    Dim ObjAppRow As New zlTabAppRow
    Dim strTitle As String
    
    '取打印列表对象
    Set ObjThis = mfrmList.GetPrintObject(True)
    
    If ObjThis Is Nothing Then
        mfrmList.GetPrintObject False
        Exit Sub
    End If
    
    Select Case tbcList.Selected.index
        Case mListType.待配药
            strTitle = "药品待配药清单"
        Case mListType.已配药
            strTitle = "药品已配药清单"
        Case mListType.待发药, mListType.超时未发
            strTitle = "药品待发药清单"
        Case mListType.退药
            strTitle = "药品退药清单"
    End Select
    
    Set ObjAppRow = New zlTabAppRow
    ObjAppRow.Add "打印人:" & gstrUserName
    ObjAppRow.Add "打印日期:" & Format(Sys.Currentdate, "yyyy-MM-dd")
    objPrint.BelowAppRows.Add ObjAppRow
    
    Set ObjAppRow = New zlTabAppRow
    ObjAppRow.Add "开始时间:" & Format(Dtp开始时间.Value, "yyyy-MM-dd HH:mm:ss")
    ObjAppRow.Add "结束时间:" & Format(Dtp结束时间.Value, "yyyy-MM-dd HH:mm:ss")
    objPrint.UnderAppRows.Add ObjAppRow
    
    objPrint.Title.Text = strTitle
    Set objPrint.Body = ObjThis
    
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
    
    mfrmList.GetPrintObject False
End Sub
Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub


Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    On Error Resume Next
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    Me.picDetail.Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop
End Sub


Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
        Case mconMenu_View_StatusBar '状态栏
            Control.Checked = Me.stbThis.Visible
        Case mconMenu_View_FontSize_1, mconMenu_View_FontSize_2, mconMenu_View_FontSize_3       '字体
            Control.Checked = Val(Control.Parameter) = mParams.intFont
        Case mconMenu_Edit_Recipe_MedicalRecord
            Control.Enabled = mfrmDetail.CmdSend.Enabled
     End Select
End Sub
Private Sub Chk显示过程单据_Click()
    Call SaveSetting("ZLSOFT", "公共模块\操作\" & App.ProductName & "\" & "药品处方发药", "显示退药过程单据", Chk显示过程单据.Value)
    RefreshList mcondition.intListType
End Sub

Private Sub Chk显示退药待发单据_Click()
    RefreshList mcondition.intListType
End Sub

Private Sub cmdFind_Click()
    Call Form_KeyDown(vbKeyF3, 0)
End Sub

'Private Sub cmdIC_Click()
'    Dim strOutXML As String
'    Dim strText As String
'
'    If Val(lblPati.Tag) = mFindType.IC卡 Then
'        If mobjICCard Is Nothing Then
'            Set mobjICCard = CreateObject("zlICCard.clsICCard")
'            Set mobjICCard.gcnOracle = gcnOracle
'        End If
'        If Not mobjICCard Is Nothing Then
'            txtPati.Text = mobjICCard.Read_Card()
'            If txtPati.Text <> "" Then Call txtPati_KeyPress(vbKeyReturn)
'        End If
'    Else
'        If Not mobjSquareCard Is Nothing Then
'            Call mobjSquareCard.zlReadCard(Me, mlngMode, Val(Split(txtPati.Tag, "|")(gCardFormat.卡类别ID)), True, "", strText, strOutXML)
'            txtPati.Text = strText
'            If txtPati.Text <> "" Then Call txtPati_KeyPress(vbKeyReturn)
'        End If
'    End If
'End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.Id
        Case 1
            Item.Handle = picCondition.hWnd
        Case 2
            Item.Handle = picList.hWnd
    End Select
End Sub

Private Sub Form_Activate()
'    If mblnStart = False Then
'        Unload Me
'        Exit Sub
'    End If

    Call picConMain_Resize
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnFirst As Boolean
    Dim strInput As String
    Dim strNos As String
    Dim strReturn As String
    Dim strCard As String
    
    If KeyCode = vbKeyF3 Then
        If imgFilter.BorderStyle = cstLocate Then
            If txtPati.Text = "" Then
                txtPati.SetFocus
            Else
                Call txtPati_Validate(False)
                Call zlControl.TxtSelAll(txtPati)
                strCard = IDKNType.GetCurCard.名称
                If strCard = "IC卡" Then
                    If Not mobjSquareCard Is Nothing Then Call mobjSquareCard.zlGetPatiID("IC卡", UCase(Trim(txtPati.Text)), False, mlngIC病人id)
                    strInput = mlngIC病人id
                
                ElseIf strCard = "姓名" Or strCard = "单据号" Or strCard = "住院号" Or strCard = "医保号" Or strCard = "身份证" Or strCard = "门诊号" Then
                    
                    strInput = txtPati.Text
                Else
                    '消费卡类别时输入为卡ID+卡号
                    strInput = mobjcard.接口序号 & "|" & txtPati.Text
                End If
                If mfrmList.FindSpecialRow(IDKNType.GetCurCard.名称, strInput, strNos, mobjSquareCard) = True Then
                    mblnFinding = True
                    If mcondition.intListType = mListType.待配药 And mParams.int输入模式 = mFindType.单据号 And mParams.bln配药扫描 = True Then
                        '配药模式启用扫描器时处理
                        If mblnScaned = False Then
                            '第一次扫描
                            mblnScaned = True
                        Else
                            '第二次扫描，确认配药
                            mblnScaned = False
                            mstrScanerLastNo = ""
                            If tbcDetail.Selected.index = 0 Then
                                mfrmDetail.CmdProcess
                            ElseIf tbcDetail.Item(1).Visible = True Then
                                mfrmRecipe.CmdProcess
                            End If
                        End If
                        txtPati.SetFocus
                        txtPati.Text = ""
                    ElseIf mcondition.intListType = mListType.待发药 And mbln允许两次刷卡 = True Then
                        '两次刷卡发药模式
                         
                        If mblnBrushCard = False Then
                            '第一次刷卡
                            mblnBrushCard = True
                        Else
                            If txtPati.Text = mstrLastBrushCardNo Then
                                '第二次刷卡，确认发药
                                mblnBrushCard = False
                                mstrLastBrushCardNo = ""
                                If tbcDetail.Selected.index = 0 Then
                                    mfrmDetail.CmdProcess
                                ElseIf tbcDetail.Item(1).Visible = True Then
                                    mfrmRecipe.CmdProcess
                                End If
                            Else
                                '不是刷的同一张卡
                                mblnBrushCard = True
                                mstrLastBrushCardNo = txtPati.Text
                            End If
                        End If
                        txtPati.SetFocus
                        txtPati.Text = ""
                    Else
                        If tbcDetail.Selected.index = 0 Then
                            If mfrmDetail.CmdSend.Enabled Then mfrmDetail.CmdSend.SetFocus
                        ElseIf tbcDetail.Item(1).Visible = True Then
                            If mfrmRecipe.CmdSend.Enabled Then mfrmRecipe.CmdSend.SetFocus
                        End If
                    End If
                    
                    '如果自动发药有开始发药流程，则调用接口上传处方
                    '如果列表中有该行病人的多个处方，一并上传
                    '不兼容接口时没有这个功能
                    If TypeName(mobjDrugMAC) = "clsDrugPacker" Then
                        If mcondition.intListType = mListType.待发药 And mblnPackerConnect And mblnLoadDrug And mintAutoSendFlow = 1 And strNos <> "" Then
                            If mblnCompatible = True Then
                                If mobjDrugMAC.DYEY_MZ_TransRecipeList(mstrOpr, UserInfo.用户编码, UserInfo.用户姓名, mParams.lng药房ID, strNos, strReturn, mSendOper.StartSend) = False Then
                                    If MsgBox("自动发药系统未准备好，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                        Exit Sub
                                    End If
                                End If
                            End If
                        End If
                    ElseIf TypeName(mobjDrugMAC) = "clsDrugMachine" Then
                        If mcondition.intListType = mListType.待发药 And mblnPackerConnect Then
                            mobjDrugMAC.Operation gstrDbUser, Val("22-开始发药"), "1|" & Replace(strNos, "|", ";"), strReturn
'                           If strReturn <> "" Then MsgBox strReturn, vbInformation, gstrSysName
                        End If
                    End If
                 Else
                    '没有找到时
                    If mcondition.intListType = mListType.待配药 And mParams.int输入模式 = mFindType.单据号 And mParams.bln配药扫描 = True Then
                        mblnScaned = False
                        mstrScanerLastNo = ""
                        txtPati.SetFocus
                        txtPati.Text = ""
                    End If
                    mblnBrushCard = False
                    mstrLastBrushCardNo = ""
                End If
            End If
        Else
'            Call SetFilter(MnuEditHandback.Checked)
            Me.IDKNType.ActiveFastKey
            RefreshList mcondition.intListType
        End If
    End If
    
    If KeyCode = 70 Or KeyCode = 102 Then
        If Shift = vbCtrlMask Then   'Ctrl+F
            txtPati.SetFocus
        End If
    End If
    
    'Ctrl+F4  读IC卡
'    If KeyCode = vbKeyF4 Or KeyCode = 102 Then
'        If Shift = vbCtrlMask Then
'            If cmdIC.Visible = True Then
'                Call cmdIC_Click
'            End If
'        End If
'    End If
End Sub

Private Sub Form_Load()
    Dim dteTime As Date
    Dim strMessage As String, strPrivs As String
   
    mblnStart = False
    mblnSendIsOver = True
    mlngMode = glngModul
    mstrPrivs = gstrprivs
    
    mQueue.strPCName = AnalyseComputer
    
    Me.Width = mcstlngWinNormalWidth
    Me.Height = mcstlngWinNormalHeight
    
    picConMain.BackColor = &H80000005
    lbl时间范围.BackColor = picConMain.BackColor
    lblTimeBegin.BackColor = picConMain.BackColor
    lblTimeEnd.BackColor = picConMain.BackColor
    IDKNType.BackColor = picConMain.BackColor
    lbl病区.BackColor = picConMain.BackColor
    Chk显示过程单据.BackColor = picConMain.BackColor
    chk显示已确认单据.BackColor = picConMain.BackColor
    Chk显示退药待发单据.BackColor = picConMain.BackColor
    
    mdate上次校验时间 = Sys.Currentdate
    mstr自动配药人 = ""
    
    int模式 = 1
    Set mclsComLib = New zl9ComLib.clsComLib
    
    mstrChargePrivs = GetPrivFunc(glngSys, 1120)
    mstrStuffPrivs = GetPrivFunc(glngSys, 1723)
    
    If gstrUserName = "" Then
        MsgBox "请为当前用户设置对应的操作员后再使用本模块！", vbInformation, gstrSysName
        Exit Sub
    End If
     
    '取费用金额位数，用于界面显示
    mintMoneyDigit = gtype_UserSysParms.P9_费用金额保留位数
    '设置金额格式
    Call GetMoneyFormat
    
    '取权限
    Call GetPrivs
    
    '依赖数据检测
    If DependOnCheck = False Then Exit Sub
    
    '取参数
    Call GetParams
    
    With mParams
        '注册表参数
        .int界面定位 = Val(GetSetting("ZLSOFT", "公共模块\操作\" & App.ProductName & "\" & "药品处方发药", "界面定位", cstLocate))
        .int待发单据 = GetSetting("ZLSOFT", "公共模块\操作\" & App.ProductName & "\" & "药品处方发药", "显示退药待发单据", 1)
        .int过程单据 = GetSetting("ZLSOFT", "公共模块\操作\" & App.ProductName & "\" & "药品处方发药", "显示退药过程单据", 1)
        .int已确认单据 = GetSetting("ZLSOFT", "公共模块\操作\" & App.ProductName & "\" & "药品处方发药", "显示已确认单据", 1)
        .int输入模式索引 = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品处方发药", "输入模式", "1"))
        If .int输入模式索引 < 1 Then
            .int输入模式索引 = 1
        End If
    End With
    
    Call GetOpr
    
    Call SetFontSize
    
    '检查相关设置
    If CheckAnother = False Then Exit Sub
    
    Call Load时间范围
    
    If Not mPrives.bln允许查询所有时间范围单据 Then
        cbo时间范围.ListIndex = 0
        cbo时间范围.Tag = 0
        cbo时间范围.Enabled = False
    End If
    
    Call Load病区
    
    Call GetDrugStock(mParams.lng药房ID)
    Call GetDosage(mParams.lng药房ID)
    Call GetSendWindows(mParams.lng药房ID)
    
    '创建电子病案查阅对象
    If mobjCISJOB Is Nothing Then
        On Error Resume Next
        Set mobjCISJOB = CreateObject("zl9CISJob.clsCISJob")
        
        If Not mobjCISJOB Is Nothing Then
            Call mobjCISJOB.InitCISJob(gcnOracle, Me, glngSys, mstrPrivs, gobjBrower.mobjEmr)
        End If
        err.Clear: On Error GoTo 0
    End If
    
    '初始化数据
    dteTime = Sys.Currentdate
    Dtp开始时间.Value = Format(dteTime, "yyyy-MM-dd 00:00:00")
    Dtp结束时间.Value = Format(dteTime, "yyyy-MM-dd") & " 23:59:59"
    
    GetStockName mParams.lng药房ID
    
    '过滤开关
    imgFilter.BorderStyle = mParams.int界面定位
    If imgFilter.BorderStyle = 0 Then
        imgFilter.ToolTipText = "点击切换到过滤模式"
    Else
        imgFilter.ToolTipText = "点击切换到定位模式"
    End If
    
    '病区过滤开关：默认是0-不显示
    img病区.BorderStyle = mParams.int显示病区处方
    
    cbo病区.Enabled = (img病区.BorderStyle = 1)
    
    If mPrives.bln发病区处方 = False Then
        lbl病区.Visible = False
        img病区.Visible = False
        cbo病区.Visible = False
    End If
    
    '设置时间控件
    TimeRefresh.Enabled = False
    TimePrint.Enabled = False
    If mParams.lngRefreshInterval > 0 Then
        If mParams.lngRefreshInterval > 60 Then
            mParams.lngRefreshInterval = 60
        End If
        With TimeRefresh
            .Enabled = True
            .Interval = mParams.lngRefreshInterval * 1000
        End With
    End If
    
    If mParams.lngPrintInterval > 0 Then
        If mParams.lngPrintInterval > 60 Then
            mParams.lngPrintInterval = 60
        End If
        With TimePrint
            .Enabled = True
            .Interval = mParams.lngPrintInterval * 1000
        End With
    End If
    IntTimes = 0
    If mParams.lngPrintBackInterval <> 0 Then
        With TimePrintCancelBill
            .Enabled = False
            .Enabled = True
        End With
    Else
        TimePrintCancelBill.Enabled = False
    End If
    
    '判断本机是否是远程呼叫机器
    mQueue.blnRemoteCall = False
    If mParams.intCallType = 0 And mParams.strRemoteCall = mQueue.strPCName And mQueue.strPCName <> "" Then
        mQueue.blnRemoteCall = True
    End If
    
    '设置叫号轮询时间间隔：当本机机器名等于全局远端机器名时
    tmrCall.Enabled = False
    If mParams.blnStartQueue = True And mParams.blnStartCall = True And mQueue.blnRemoteCall = True Then
        tmrCall.Enabled = True
        tmrCall.Interval = mParams.intCircleTime * 1000
    End If
    
    Set mobjIDCard = New clsIDCard
    Set mobjICCard = New clsICCard
    Call mobjIDCard.SetParent(Me.hWnd)
    Call mobjICCard.SetParent(Me.hWnd)
    Set mobjICCard.gcnOracle = gcnOracle
    
    '电子签名接口控制
    gblnESign处方发药 = EsignIsOpen(mParams.lng药房ID)
    gblnESignUserStoped = False
    If gblnESign处方发药 = True Then
        On Error Resume Next
        Set gobjESign = CreateObject("zl9ESign.clsESign")
        err.Clear: On Error GoTo 0
        If Not gobjESign Is Nothing Then
            If Not gobjESign.Initialize(gcnOracle, glngSys) Then
                Set gobjESign = Nothing
                gblnESign处方发药 = False
            Else
                gblnESign处方发药 = True
                gblnESignUserStoped = gobjESign.CertificateStoped(gstrUserName)
            End If
        Else
            gblnESign处方发药 = False
        End If
    End If
    
    '一卡通接口
    mstrCardType = zlfuncCard_Ini(mobjSquareCard, Me, mlngMode)
    
    '自动发药机接口
    mblnPackerConnect = False
    On Error Resume Next
    
    '检查药品自动化接口权限和参数
    If Val(zldatabase.GetPara("启用药品自动化设备接口", glngSys, Val("9010-药品自动化设备接口"))) = 1 _
        And mPrives.bln药品自动化接口 = True Then
        
        Set mobjDrugMAC = Nothing
        '优先新接口
        Set mobjDrugMAC = CreateObject("zlDrugMachine.clsDrugMachine")
        If err.Number <> 0 Then
            '其次旧接口
            Set mobjDrugMAC = CreateObject("zlDrugPacker.clsDrugPacker")
        End If
    Else
        Set mobjDrugMAC = CreateObject("zlDrugPacker.clsDrugPacker")
    End If
    On Error GoTo 0
    
    If TypeName(mobjDrugMAC) = "clsDrugMachine" Then
        '新接口
        ''获取接口的权限
        strPrivs = ";" & zl9ComLib.GetPrivFunc(glngSys, Val("9010-药品自动化设备接口")) & ";"
        If strPrivs Like "*;基本;*" Then
            mblnPackerConnect = mobjDrugMAC.Init(1, mclsComLib, strMessage)
        Else
            mblnPackerConnect = False
        End If
    ElseIf TypeName(mobjDrugMAC) = "clsDrugPacker" Then
        '旧接口
        mblnPackerConnect = mobjDrugMAC.DYEY_MZ_IniSoap(, , gstrUnitName)
        
        On Error Resume Next
        mintAutoSendFlow = mobjDrugMAC.DYEY_MZ_GetSendType      '部分旧接口无该方法
        mblnCompatible = (err.Number = 0)
        On Error GoTo 0
    Else
        mblnPackerConnect = False
        mintAutoSendFlow = False
    End If
    
    '发药业务外挂部件
    Call zlPlugIn_Ini(glngSys, glngModul, mobjPlugIn)
    
    '设置菜单
    Call InitComandBars
    Call InitPanes
    Call InitTabControl
    Call InitIDKindNew
        
    Chk显示退药待发单据.Value = IIf(mParams.int待发单据 = 1, 1, 0)
    Chk显示过程单据.Value = IIf(mParams.int过程单据 = 1, 1, 0)
    chk显示已确认单据.Value = IIf(mParams.int已确认单据 = 1, 1, 0)
    
'    添加自定义报表
    Call zldatabase.ShowReportMenu(Me, glngSys, glngModul, gstrprivs)
    
    '恢复录入状态
'    Call SetInputState(mParams.int输入模式)
    
    '恢复窗口
    If Val(zldatabase.GetPara("使用个性化风格")) = 1 Then
        On Error Resume Next
        
        dkpMain.LoadStateFromString GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name & dkpMain.PanesCount, "")
        SetPaneTitle tbcList.Selected.index
    End If
    Call RestoreWinState(Me, App.ProductName)
    
    '打开排队显示窗口
    Call ShowQueue

    Call zlCall_SystemSoundPlay("", 65)
    
    mQueue.blnCallOver = True
    
'    '门诊药房自动发药
'    If gtype_UserSysParms.P222_药房自动化发药接口 = 1 Then
'        err = 0
'        On Error Resume Next
'
'        If gobjPackerMZ Is Nothing Then
'            Set gobjPackerMZ = CreateObject("zlDrugPacker.clsDrugPacker")
'            err.Clear
'
'            If Not gobjPackerMZ Is Nothing Then
'                gobjPackerMZ.InitCommon gcnOracle, Me, glngSys, mlngMode, mParams.lng药房ID
'            End If
'        End If
'    End If

 
    '初始化消息对象
    err = 0
    On Error Resume Next
    Set mobjMipModule = New zl9ComLib.clsMipModule
    Call mobjMipModule.InitMessage(glngSys, mlngMode, mstrPrivs)
    Call AddMipModule(mobjMipModule)
       
    mblnStart = True

'    '加载窗口时在消息机制有效的前提下需要刷新一次
'    If Not mobjMipModule Is Nothing Then
'        If mobjMipModule.IsConnect = True Then
            RefreshList IIf(mParams.blnMustDosageProcess = True, mListType.待配药, mListType.待发药)
'        End If
'    End If

    mdteMsgRefresh = Now
End Sub

Private Sub GetMoneyFormat()
    Dim n As Integer
    Dim strOracleTmp As String
    Dim strVbTmp As String
    
    strOracleTmp = "999999990."
    strVbTmp = "########0."
    For n = 1 To mintMoneyDigit
        strOracleTmp = strOracleTmp & "0"
        strVbTmp = strVbTmp & "0"
    Next
    
    mstrOracleMoneyForamt = strOracleTmp
    mstrVBMoneyForamt = strVbTmp
    
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.Width < mcstlngWinNormalWidth Then Me.Width = mcstlngWinNormalWidth
    If Me.Height < mcstlngWinNormalHeight Then Me.Height = mcstlngWinNormalHeight
End Sub



Public Function RefreshDetail_Send(ByVal lngNO库房ID As Long, ByVal BillStyle As Integer, ByVal BillNo As String, ByVal int门诊标志 As Integer, ByVal int记录性质 As Integer, Optional ByVal int排队类型 As Integer, Optional ByVal int审查结果 As Integer) As Boolean
    Dim IntStyle As Integer, intUnit As Integer
    Dim strSubSql As String
    Dim strName As String
    Dim blnMoved As Boolean
    Dim lng库房ID As Long
    Dim lng病人ID As Long
    Dim int主页id As Integer
    Dim strWeight As String
    
    Dim rstemp As New ADODB.Recordset
    Dim RecBill As New ADODB.Recordset
    '--读取单据内容--
    'BillStyle-单据类型;BIllNO-单据号
    '单位显示根据服务对象来（门诊：门诊单位；住院或住院门诊：住院单位；其它；售价单位）
'    On Error Resume Next
'    err = 0
    On Error GoTo errHandle
    RefreshDetail_Send = False
    
    If mPrives.bln发其它药房的处方 = False Then
        lng库房ID = mSQLCondition.lng药房ID
    Else
        lng库房ID = lngNO库房ID
    End If
    
    If lng库房ID = 0 Then lng库房ID = mSQLCondition.lng药房ID
    
    mParams.strUnit = GetUnit(lng库房ID, BillStyle, BillNo, int门诊标志)
    Select Case mParams.strUnit
    Case "售价单位"
        strSubSql = "1"
    Case "门诊单位"
        strSubSql = "Decode(门诊包装,Null,1,0,1,门诊包装)"
    Case "住院单位"
        strSubSql = "Decode(住院包装,Null,1,0,1,住院包装)"
    Case "药库单位"
        strSubSql = "Decode(药库包装,Null,1,0,1,药库包装)"
    End Select
    Call Get单位串
    
    '得到药品名称串
    Select Case mParams.int药品名称显示
    Case 0  '药品编码与名称
        strName = "'['||C.编码||']'||" & IIf(gint药品名称显示 = 1, "NVL(E.名称,C.名称)", "C.名称") & " As 品名,"
    Case 1  '药品编码
        strName = "C.编码 As 品名,"
    Case 2  '药品名称
        strName = IIf(gint药品名称显示 = 1, "NVL(E.名称,C.名称)", "C.名称") & " As 品名,"
    End Select
    
    strName = strName & IIf(gint药品名称显示 <> 1, "NVL(E.名称,'')", "Decode(E.名称,Null,'',C.名称)") & " As 其它名, "
    
    gstrSQL = " SELECT DISTINCT B.记录状态 状态,S.名称 As 药房,Nvl(B.库房ID,0) as 药房ID,B.单据,B.NO,Nvl(A.处方类型,Nvl(B.注册证号,0)) As 处方类型,nvl(n.名称,'') 配方名称,H.序号,T.名称 科室,H.姓名,H.性别,H.年龄,H.标识号 住院号,H.床号,H.开单人,B.ID As 收发ID," & _
        " B.药品ID,D.药名id,DECODE(B.批号,NULL,'',B.批号)||DECODE(B.批次,NULL,'',0,'','('||B.批次||')') 批号,A.已收费,DECODE(D.高危药品,null,0,0,0,1) 高危药品,to_char(B.效期,'yyyy-mm-dd') 效期,X.门诊号," & _
        " NVL(B.批次,0) 批次,NVL(D.药房分批,0) 分批,F.名称 As 险类,B.序号 As 收发序号,C.规格 As 药品规格,H.门诊标志, " & strName & _
        " DECODE(C.规格,NULL,B.产地,DECODE(B.产地,NULL,C.规格,C.规格||'|'||B.产地)) 规格,Nvl(b.产地, Nvl(c.产地, '')) 产地,b.原产地," & str单位串 & ",Nvl(K.实际数量,0)/" & strSubSql & " 库存数,Nvl(K.实际数量,0) 库存实际数量," & _
        IIf(gint药品名称显示 = 1, "NVL(E.名称,C.名称)", "C.名称") & " As 药品名称,decode(m.用药目的,1,'预防',2,'治疗',3,'预防和治疗','') 用药目的,m.用药理由, " & _
        " NVL(B.付数,1) 付数,NVL(H.付数,1) 原始付数,B.单量,B.用法,B.频次,B.填制人,B.填制日期,H.操作员姓名," & IIf(mcondition.intListType <> mListType.退药, "B.配药人", "B.审核人") & " 配药人,B.配药日期,B.审核日期, " & _
        " L.库房货位,Nvl(M.相关ID,0) As 相关ID,M.医生嘱托,M.禁忌药品说明,M.开嘱医生,M.频率间隔,M.间隔单位,Nvl(M.开嘱时间,H.登记时间) As 开嘱时间,M.医嘱期效 医嘱标志,M.开始执行时间 开始时间,M.执行终止时间 结束时间," & _
        " M.频率次数,Nvl(Nvl(M.相关ID,M.id),0) As 医嘱id,nvl(M.审查结果,-1) 审查结果,M.皮试结果,M.超量说明,D.药名ID,I.计算单位,D.剂量系数," & _
        " round(B.零售金额," & mintMoneyDigit & ") 零售金额,Nvl(B.付数, 1) * B.实际数量 / (Nvl(H.付数, 1) * H.数次) * Nvl(H.实收金额,0) As 实收金额,H.费别,P.毒理分类,Nvl(p.抗生素,0) 抗生素, " & _
        " B.实际数量*D.剂量系数* Nvl(B.付数, 1) 重量,B.实际数量,Decode(Sign(Nvl(J.库存数量, 0) - Nvl(L.下限, 0)), -1, 0, 1) 库存下限,Z.名称 As 英文名,1 As 标志,Nvl(H.病人ID,0) As 病人ID,H.记录性质, H.记录状态,Zl_Get收费类别([15],[14],[13]) 收费类别,X.就诊卡号,X.结算模式,decode(X.联系人电话,null,decode(X.手机号,null,X.家庭电话,X.手机号),X.联系人电话) 联系人电话,Nvl(m.主页id,0) As 主页id, Nvl(G.检查方法, H.结论) As 中药形态, I.名称 As 药名,X.病人类型,Nvl(P.是否皮试,0) As 是否皮试, Nvl(x.在院, 0) As 在院 "
    gstrSQL = gstrSQL & _
        " FROM 药品收发记录 B,药品规格 D,药品特性 P,收费项目目录 C,收费项目别名 E," & _
        " 门诊费用记录 H,病人医嘱记录 M,病人医嘱记录 G,病人信息 X,部门表 S,部门表 T,药品库存 K,药品储备限额 L,诊疗项目目录 I,诊疗项目别名 Z ,未发药品记录 A,保险支付大类 F,诊疗项目目录 N, " & _
        " (Select b.库房id, b.药品id, Nvl(Sum(b.实际数量), 0) 库存数量 " & _
        " From 药品收发记录 A, 药品库存 B " & _
        " Where a.药品id = b.药品id And b.性质 = 1 And b.库房id + 0 = [13] And a.单据 = [15] And a.No = [14] " & _
        " Group By b.库房id, b.药品id) J " & _
        " WHERE A.单据=B.单据 And A.NO=B.No And D.药品ID=C.ID And D.药名ID=P.药名ID And H.医嘱序号=M.ID(+) And Nvl(M.相关id, M.ID) = G.ID(+) AND C.ID=E.收费细目ID(+) AND E.性质(+)=3 " & _
        " And B.药品ID=L.药品ID(+) And Nvl(B.库房ID,[13])=L.库房ID(+) And H.保险大类ID=F.ID(+) and G.配方id=N.id(+) " & _
        " AND H.开单部门ID=T.ID(+) AND B.药品ID=D.药品ID AND MOD(B.记录状态,3)=1" & _
        " AND S.ID=NVL(B.库房ID,[13]) AND B.费用ID=H.ID AND B.NO=[14] AND B.单据=[15] AND NVL(B.库房ID,[13])+0=[13] AND LTRIM(RTRIM(NVL(B.摘要,'小宝')))<>'拒发'" & _
        " AND B.药品ID=K.药品ID(+) AND K.性质(+)=1 AND NVL(B.库房ID,[13])=K.库房ID(+) AND NVL(B.批次,0)=NVL(K.批次(+),0) AND B.审核人 IS NULL And D.药名id=I.id " & _
        " And Nvl(B.库房id, [13]) + 0 = J.库房id(+) And B.药品id = J.药品id(+) And D.药名id = Z.诊疗项目id(+) And Z.性质(+) = 2 And H.病人id = X.病人id(+) "
        
    gstrSQL = gstrSQL & " Order by H.序号,B.药品ID,Nvl(B.批次,0)"
    
    If int记录性质 = 1 Or (int记录性质 = 2 And (int门诊标志 = 1 Or int门诊标志 = 4)) Then
        '门诊
        gstrSQL = Replace(gstrSQL, "H.床号", "'' 床号")
    Else
        '住院
        gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
    End If
     
    Set RecBill = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
            mSQLCondition.date开始日期, _
            mSQLCondition.date结束日期, _
            mSQLCondition.str开始NO, _
            mSQLCondition.str结束NO, _
            mSQLCondition.str姓名, _
            mSQLCondition.str就诊卡, _
            mSQLCondition.str标识号, _
            mSQLCondition.lng科室ID, _
            mSQLCondition.str填制人, _
            mSQLCondition.str审核人, _
            mSQLCondition.lng药品id, _
            mSQLCondition.str当前NO, _
            lng库房ID, BillNo, BillStyle)
    
    If err <> 0 Then
        MsgBox "读取处方时，发生不可预知的错误！", vbInformation, gstrSysName
        Exit Function
    End If
    
'    If RecBill!药房id <> mParams.lng药房ID Then RecBill.Filter = "状态=1"
    
    If Not RecBill.EOF Then
        If NVL(RecBill!病人ID) <> 0 Then
            lng病人ID = RecBill!病人ID
            int主页id = NVL(RecBill!主页id)
            If int记录性质 = 1 Or (int记录性质 = 2 And (int门诊标志 = 1 Or int门诊标志 = 4)) Then
                '门诊
                gstrSQL = "select A.id,B.记录内容 体重 from 病人护理记录 A,病人护理内容 B where A.id=B.记录id and B.项目名称='体重' and 病人id=[1] order by A.Id desc"
            Else
                '住院
                 gstrSQL = "select 体重 from 病案主页 where 病人id=[1] and 主页id=[2]"
                 
            End If
            Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng病人ID, int主页id)
            
            If Not rstemp.EOF Then
                strWeight = NVL(rstemp!体重)
            End If
        End If
    End If
    
    mfrmDetail.RefreshList RecBill, strWeight, 0, int排队类型, int审查结果
    
    mfrmRecipe.RefreshRecipe RecBill, strWeight, 0, int排队类型, int审查结果
    
    RefreshDetail_Send = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetRecipeRecord(ByVal BillStyle As Integer, ByVal BillNo As String, ByVal int门诊标志 As Integer, ByVal int记录性质 As Integer) As ADODB.Recordset
    Dim IntStyle As Integer, intUnit As Integer
    Dim strSubSql As String
    Dim strName As String
    Dim blnMoved As Boolean
    
    Dim rstemp As New ADODB.Recordset
    Dim RecBill As New ADODB.Recordset
    '--读取单据内容--
    'BillStyle-单据类型;BIllNO-单据号
    '单位显示根据服务对象来（门诊：门诊单位；住院或住院门诊：住院单位；其它；售价单位）
'    On Error Resume Next
    On Error GoTo errHandle
    mParams.strUnit = GetUnit(mSQLCondition.lng药房ID, BillStyle, BillNo, int门诊标志)
    Select Case mParams.strUnit
    Case "售价单位"
        strSubSql = "1"
    Case "门诊单位"
        strSubSql = "Decode(门诊包装,Null,1,0,1,门诊包装)"
    Case "住院单位"
        strSubSql = "Decode(住院包装,Null,1,0,1,住院包装)"
    Case "药库单位"
        strSubSql = "Decode(药库包装,Null,1,0,1,药库包装)"
    End Select
    Call Get单位串
    
    '得到药品名称串
    Select Case mParams.int药品名称显示
    Case 0  '药品编码与名称
        strName = "'['||C.编码||']'||" & IIf(gint药品名称显示 = 1, "NVL(E.名称,C.名称)", "C.名称") & " As 品名,"
    Case 1  '药品编码
        strName = "C.编码 As 品名,"
    Case 2  '药品名称
        strName = IIf(gint药品名称显示 = 1, "NVL(E.名称,C.名称)", "C.名称") & " As 品名,"
    End Select
    
    strName = strName & IIf(gint药品名称显示 <> 1, "NVL(E.名称,'')", "Decode(E.名称,Null,'',C.名称)") & " As 其它名, "
    
    gstrSQL = " SELECT DISTINCT B.单据,B.NO,Nvl(A.处方类型,0) As 处方类型,H.序号,T.名称 科室,H.姓名,H.性别,H.年龄,H.标识号 住院号,H.开单人,B.ID As 收发ID," & _
        " B.药品ID,DECODE(B.批号,NULL,'',B.批号)||DECODE(B.批次,NULL,'',0,'','('||B.批次||')') 批号,A.已收费," & _
        " NVL(B.批次,0) 批次,NVL(D.药房分批,0) 分批,F.名称 As 险类,B.序号 As 收发序号,C.规格 As 药品规格,H.门诊标志, " & strName & _
        " DECODE(C.规格,NULL,B.产地,DECODE(B.产地,NULL,C.规格,C.规格||'|'||B.产地)) 规格," & str单位串 & ",K.实际数量/" & strSubSql & " 库存数," & _
        IIf(gint药品名称显示 = 1, "NVL(E.名称,C.名称)", "C.名称") & " As 药品名称, " & _
        " NVL(H.付数,1) 付数,B.单量,B.用法,B.频次,B.填制人,B.填制日期,H.操作员姓名," & IIf(mcondition.intListType <> mListType.退药, "B.配药人", "B.审核人") & " 配药人," & _
        " L.库房货位,Nvl(M.相关ID,0) As 相关ID,M.医生嘱托,Nvl(Nvl(M.相关ID,M.id),0) As 医嘱id,nvl(M.审查结果,-1) 审查结果,I.计算单位," & _
        " round(B.零售金额," & mintMoneyDigit & ") 零售金额,Nvl(B.付数, 1) * B.实际数量 / (Nvl(H.付数, 1) * H.数次) * Nvl(H.实收金额,0) As 实收金额,H.费别,P.毒理分类, " & _
        " B.实际数量*D.剂量系数* Nvl(B.付数, 1) 重量,Decode(Sign(Nvl(J.库存数量, 0) - Nvl(L.下限, 0)), -1, 0, 1) 库存下限,Z.名称 As 英文名,1 As 标志,Nvl(H.病人ID,0) As 病人ID,H.记录性质, H.记录状态,Zl_Get收费类别(b.单据,b.NO,[13]) As 收费类别,X.结算模式, X.就诊卡号,X.病人类型,B.库房id As 药房ID " & _
        " FROM 药品收发记录 B,药品规格 D,药品特性 P,收费项目目录 C,收费项目别名 E,"
            
    If int记录性质 = 1 Or (int记录性质 = 2 And (int门诊标志 = 1 Or int门诊标志 = 4)) Then
        gstrSQL = gstrSQL & " 门诊费用记录 H,"
    Else
        gstrSQL = gstrSQL & " 住院费用记录 H,"
    End If
            
    gstrSQL = gstrSQL & " 病人医嘱记录 M,病人信息 X,部门表 S,部门表 T,药品库存 K,药品储备限额 L,诊疗项目目录 I,诊疗项目别名 Z ,未发药品记录 A,保险支付大类 F," & _
        " (Select 库房id, 药品id, Nvl(Sum(实际数量), 0) 库存数量 From 药品库存 Where 性质 = 1 And 库房id = [13] Group By 库房id, 药品id) J " & _
        " WHERE A.单据=B.单据 And A.NO=B.No And D.药品ID=C.ID And D.药名ID=P.药名ID And H.医嘱序号=M.ID(+) AND C.ID=E.收费细目ID(+) AND E.性质(+)=3 " & _
        " And B.药品ID=L.药品ID(+) And Nvl(B.库房ID,[13])=L.库房ID(+) And H.保险大类ID=F.ID(+) " & _
        " AND H.开单部门ID=T.ID(+) AND B.药品ID=D.药品ID AND MOD(B.记录状态,3)=1" & _
        " AND S.ID=NVL(B.库房ID,[13]) AND B.费用ID=H.ID AND B.NO=[14] AND B.单据=[15] AND NVL(B.库房ID,[13])+0=[13] AND LTRIM(RTRIM(NVL(B.摘要,'小宝')))<>'拒发'" & _
        " AND B.药品ID=K.药品ID(+) AND K.性质(+)=1 AND NVL(B.库房ID,[13])=K.库房ID(+) AND NVL(B.批次,0)=NVL(K.批次(+),0) AND B.审核人 IS NULL And D.药名id=I.id " & _
        " And Nvl(B.库房id, [13]) + 0 = J.库房id(+) And B.药品id = J.药品id(+) And D.药名id = Z.诊疗项目id(+) And Z.性质(+) = 2 And H.病人id = X.病人id(+) " & _
        " Order by H.序号,B.药品ID,Nvl(B.批次,0)"
     
    Set GetRecipeRecord = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
            mSQLCondition.date开始日期, _
            mSQLCondition.date结束日期, _
            mSQLCondition.str开始NO, _
            mSQLCondition.str结束NO, _
            mSQLCondition.str姓名, _
            mSQLCondition.str就诊卡, _
            mSQLCondition.str标识号, _
            mSQLCondition.lng科室ID, _
            mSQLCondition.str填制人, _
            mSQLCondition.str审核人, _
            mSQLCondition.lng药品id, _
            mSQLCondition.str当前NO, _
            mSQLCondition.lng药房ID, BillNo, BillStyle)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Private Sub Get单位串()
    Const str售价 As String = "C.计算单位 As 售价单位,C.计算单位 As 单位,1 As 包装,ltrim(to_char(B.零售价,'999990.00000')) 单价,ltrim(to_char(B.实际数量,'999990.00000')) 数量"
    Const str门诊 As String = "C.计算单位 As 售价单位,D.门诊单位 As 单位,D.门诊包装 As 包装,ltrim(to_char(B.零售价*Decode(D.门诊包装,Null,1,0,1,D.门诊包装),'999990.00000')) 单价,ltrim(to_char(B.实际数量/Decode(D.门诊包装,Null,1,0,1,D.门诊包装),'999990.00000')) 数量"
    Const str住院 As String = "C.计算单位 As 售价单位,D.住院单位 As 单位,D.住院包装 As 包装,ltrim(to_char(B.零售价*Decode(D.住院包装,Null,1,0,1,D.住院包装),'999990.00000')) 单价,ltrim(to_char(B.实际数量/Decode(D.住院包装,Null,1,0,1,D.住院包装),'999990.00000')) 数量"
    Const str药库 As String = "C.计算单位 As 售价单位,D.药库单位 As 单位,D.药库包装 As 包装,ltrim(to_char(B.零售价*Decode(D.药库包装,Null,1,0,1,D.药库包装),'999990.00000')) 单价,ltrim(to_char(B.实际数量/Decode(D.药库包装,Null,1,0,1,D.药库包装),'999990.00000')) 数量"
    
    Select Case mParams.strUnit
    Case "售价单位"
        str单位串 = str售价
    Case "门诊单位"
        str单位串 = str门诊
    Case "住院单位"
        str单位串 = str住院
    Case "药库单位"
        str单位串 = str药库
    End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
    mblnBrushCard = False
    mstrLastBrushCardNo = ""
    mQueue.strSendWin = ""
    mstr病人类型 = ""
    
    '如果启用了语音系统，在退出窗口时关闭正在播放的语言
    If mParams.blnStartQueue = True And mParams.blnStartCall = True And mQueue.blnCallOver = False Then
        Call StopPlayStr
    End If
    tmrCall.Enabled = False
    mQueue.blnCallOver = True

    Set mobjDrugMAC = Nothing
    Set mclsComLib = Nothing
    
    TimeRefresh.Enabled = False
    TimePrint.Enabled = False
    TimePrintCancelBill.Enabled = False
    
    zldatabase.SetPara "显示病区处方", img病区.BorderStyle, glngSys, 1341, IsInString(mstrPrivs, "参数设置", ";")
    
    Call SaveSetting("ZLSOFT", "公共模块\操作\" & App.ProductName & "\" & "药品处方发药", "界面定位", imgFilter.BorderStyle)
    Call SaveSetting("ZLSOFT", "公共模块\操作\" & App.ProductName & "\" & "药品处方发药", "显示退药待发单据", Chk显示退药待发单据.Value)
    Call SaveSetting("ZLSOFT", "公共模块\操作\" & App.ProductName & "\" & "药品处方发药", "显示退药过程单据", Chk显示过程单据.Value)
    Call SaveSetting("ZLSOFT", "公共模块\操作\" & App.ProductName & "\" & "药品处方发药", "显示已确认单据", chk显示已确认单据.Value)
    
'    '保存排序串
'    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品处方发药", "未配药处方排序串", strOrder_1)
'    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品处方发药", "已配药处方排序串", strOrder_2)
'    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品处方发药", "未发药处方排序串", strOrder_3)
'    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品处方发药", "已发药处方排序串", strOrder_4)
    
    '保存输入模式
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品处方发药", "输入模式", mParams.int输入模式索引)
    
    Call SaveWinState(Me, App.ProductName)
    
    '保存窗口
    If Val(zldatabase.GetPara("使用个性化风格")) = 1 Then
        Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name & dkpMain.PanesCount, dkpMain.SaveStateToString)
    End If
    
    '卸载电子病案查阅接口
    Set mobjCISJOB = Nothing
    
    '卸载身份证刷卡接口
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
        Set mobjIDCard = Nothing
    End If
    
    '卸载IC卡刷卡接口
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
        Set mobjICCard = Nothing
    End If
    
    '卸载CARD对象
    If Not mobjcard Is Nothing Then
        Set mobjcard = Nothing
    End If
    
    '卸载电子签名接口
    Set gobjESign = Nothing
    
    '卸载一卡通接口
    mstrCardType = ""
    Call zlfuncCard_Unload(mobjSquareCard)
    
    '卸载引用的窗口
    If Not mfrmList Is Nothing Then
        Unload mfrmList
        Set mfrmList = Nothing
    End If
    
    If Not mfrmDetail Is Nothing Then
        Unload mfrmDetail
        Set mfrmDetail = Nothing
    End If
    
    If Not mfrmRecipe Is Nothing Then
        Unload mfrmRecipe
        Set mfrmRecipe = Nothing
    End If
    
    '关闭显示输出窗口
    CloseQueue
    
    '卸载过滤条件
    mSQLCondition.str就诊卡 = ""
    mSQLCondition.str当前NO = ""
    mSQLCondition.str门诊号 = ""
    mSQLCondition.str姓名 = ""
    mSQLCondition.str身份证 = ""
    mSQLCondition.lng病人ID = 0
    mSQLCondition.str医保号 = ""
    mSQLCondition.lng住院号 = 0
    
    
    '卸载消息对象
    If Not mobjMipModule Is Nothing Then
        Call mobjMipModule.CloseMessage
        Call DelMipModule(mobjMipModule)
        Set mobjMipModule = Nothing
    End If
    mblnExistMsg = False
    
    '卸载外挂接口
    Call zlPlugIn_Unload(mobjPlugIn)
End Sub

Private Sub imgFilter_Click()
    imgFilter.BorderStyle = Abs(imgFilter.BorderStyle - 1)
    
    If imgFilter.BorderStyle = 0 Then
        imgFilter.ToolTipText = "点击切换到过滤模式"
    Else
        imgFilter.ToolTipText = "点击切换到定位模式"
    End If
    
    mParams.int界面定位 = imgFilter.BorderStyle
    '保存界面定位方式
    Call SaveSetting("ZLSOFT", "公共模块\操作\" & App.ProductName & "\" & "药品处方发药", "界面定位", mParams.int界面定位)
    
    '重新刷新
    mSQLCondition.str就诊卡 = ""
    mSQLCondition.str当前NO = ""
    mSQLCondition.str门诊号 = ""
    mSQLCondition.str姓名 = ""
    mSQLCondition.str身份证 = ""
    mSQLCondition.lng病人ID = 0
    mSQLCondition.str医保号 = ""
    mSQLCondition.lng住院号 = 0
    mlngIC病人id = 0
    
    txtPati.Text = ""
    RefreshList mcondition.intListType
End Sub

Private Sub img病区_Click()
    With img病区
        .BorderStyle = Abs(.BorderStyle - 1)
        
        cbo病区.Enabled = (.BorderStyle = 1)
        
        If cbo病区.Enabled = True Then
            Load病区
        Else
            cbo病区.ListIndex = -1
        End If
        
        Call RefreshList(mcondition.intListType)
    End With
End Sub


Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    
    If Not txtPati.Locked And txtPati.Text = "" And Me.ActiveControl Is txtPati Then
        txtPati.Text = strID
        
        If txtPati.Text <> "" Then
            mParams.int输入模式 = mFindType.身份证
'            Call SetInputState(mParams.int输入模式)
            
            DoEvents
            
            Call txtPati_KeyPress(vbKeyReturn)
            
            DoEvents
            
            If mintOld输入模式 <> mParams.int输入模式 Then
                mParams.int输入模式 = mintOld输入模式
'                Call SetInputState(mParams.int输入模式)
            End If
        End If
    End If
End Sub

Private Sub mobjMipModule_ReceiveMessage(ByVal strMsgItemIdentity As String, ByVal strMsgContent As String)
    '1.接收消息：消息类型和上游业务约定
    '2.根据客户机参数设置判断是否是有效消息
    '3.默认1分钟刷新一次
    Const CST_INT_MSGREFRESHINTERVAL As Integer = 1
    Const CST_STR_MSGCODE As String = "ZLHIS_CHARGE_003,ZLHIS_CIS_006"
    
    '消息对象为空时退出
    If mobjMipModule Is Nothing Then Exit Sub
    
    '消息服务连接失败时不接收消息
    If mobjMipModule.IsConnect = False Then Exit Sub
        
    '处方发药接收的消息类型：划价/收费和门诊医生站产生的药品处方，其余不接收
    If InStr("," & CST_STR_MSGCODE & ",", "," & strMsgItemIdentity & ",") = 0 Then Exit Sub

    '根据客户机参数设置判断是否是有效消息
    If IsValidMsg(strMsgItemIdentity, strMsgContent) = False Then Exit Sub
    
    '执行到这里表示已接收到有效消息
    mblnExistMsg = True
    
    '当前如果不是待配药或待发药界面则不继续
    If (mParams.blnMustDosageProcess = True And mcondition.intListType <> mListType.待配药) Or _
        (mParams.blnMustDosageProcess = False And mcondition.intListType <> mListType.待发药) Then
        Exit Sub
    End If
    
    '如果接收到有效消息时离上次刷新超过1分钟则立即刷新
    If DateDiff("n", mdteMsgRefresh, Now) > CST_INT_MSGREFRESHINTERVAL Then
        '刷新前关闭计时器，刷新后再开启
        tmrMsgRefresh.Enabled = False
        DoEvents
        Call RefreshList(mcondition.intListType)
        DoEvents
        tmrMsgRefresh.Enabled = True
        
        '刷新后记录当前刷新时间
        mdteMsgRefresh = Now
        
        '消息变量设为False
        mblnExistMsg = False
    End If
End Sub


Private Sub picCondition_Resize()
    On Error Resume Next
    
    With picConMain
        .Top = 0
        .Left = 0
        .Width = picCondition.Width
        
    End With
    
    With picList
        .Top = picConMain.Top + picConMain.Height
        .Left = 0
        .Width = picCondition.Width
        .Height = picCondition.Height - .Top
    End With
End Sub


Private Sub picConMain_Resize()
    On Error Resume Next
    
    With cbo时间范围
        .Width = picCondition.Width - .Left - 50
    End With

    If cbo时间范围.ListIndex <> 3 Then
        lblTimeBegin.Visible = False
        Dtp开始时间.Visible = False
        lblTimeEnd.Visible = False
        Dtp结束时间.Visible = False
        
        With Me.IDKNType
            .Top = lbl时间范围.Top + lbl时间范围.Height + 180
        End With
        
        With txtPati
            .Top = cbo时间范围.Top + cbo时间范围.Height + 180
        End With
    Else
        lblTimeBegin.Visible = True
        Dtp开始时间.Visible = True
        lblTimeEnd.Visible = True
        Dtp结束时间.Visible = True
        
        With lblTimeBegin
            .Top = lbl时间范围.Top + lbl时间范围.Height + 180
        End With
        
        With Dtp开始时间
            .Top = lblTimeBegin.Top + lblTimeBegin.Height / 2 - .Height / 2
            .Width = cbo时间范围.Width
        End With
        
        With lblTimeEnd
            .Top = lblTimeBegin.Top + lblTimeBegin.Height + 180
        End With
        
        With Dtp结束时间
            .Top = Dtp开始时间.Top + Dtp开始时间.Height + 60
            .Width = cbo时间范围.Width
        End With
        
        With IDKNType
            .Top = lblTimeEnd.Top + lblTimeEnd.Height + 180
        End With
        
        With txtPati
            .Top = IDKNType.Top + IDKNType.Height / 2 - .Height / 2
        End With
    End If
    
    With cmdIC
        .Visible = (mobjcard.是否刷卡 = 1)
        .Top = txtPati.Top
        .Left = picCondition.Width - .Width - 80
    End With

    With imgFilter
        .Top = txtPati.Top
        .Left = IIf(mobjcard.是否刷卡 = 1, cmdIC.Left, picCondition.Width) - imgFilter.Width - 120
    End With

    With cmdFind
        .Top = cmdIC.Top
        .Left = imgFilter.Left + 120
    End With

    With txtPati
        .Width = imgFilter.Left - .Left - 200
    End With
    
    If lbl病区.Visible = True Then
        With lbl病区
            .Top = IDKNType.Top + IDKNType.Height + 180
        End With
        
        With img病区
            .Top = lbl病区.Top - 30
        End With
        
        With cbo病区
            .Top = img病区.Top - 30
            .Width = cbo时间范围.Width
        End With
        
        With Chk显示退药待发单据
            .Left = lbl病区.Left
            .Top = lbl病区.Top + 350
        End With
    Else
        With Chk显示退药待发单据
            .Left = IDKNType.Left
            .Top = IDKNType.Top + 350
        End With
    End If
    
    With Chk显示过程单据
        .Left = Chk显示退药待发单据.Left
        .Top = Chk显示退药待发单据.Top
    End With
    
    With chk显示已确认单据
        .Left = Chk显示退药待发单据.Left
        .Top = Chk显示退药待发单据.Top
    End With
    
    With picConMain
        .Height = Chk显示退药待发单据.Top + Chk显示退药待发单据.Height + 50
    End With
End Sub


Private Sub picDetail_Resize()
    On Error Resume Next
    
    With fraLine
'        .Top = 0
        .Left = 0
        .Height = picDetail.Height + 100
    End With
    
    With tbcDetail
        .Top = 0
        .Left = fraLine.Left + 50
        .Width = picDetail.Width - fraLine.Width
        .Height = picDetail.Height - 50
    End With
End Sub

Private Sub picList_Resize()
    On Error Resume Next
    
    With tbcList
        .Move 0, 0, picList.Width, picList.Height
    End With
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
        
        .InsertItem(0, "处方明细清单", mfrmDetail.hWnd, 0).Tag = "处方明细清单_"
        .InsertItem(1, "处方签", mfrmRecipe.hWnd, 0).Tag = "处方签_"
        
        .Item(0).Selected = True
    End With
    
    With Me.tbcList
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        
        .InsertItem(mconTab_Recipe_DosageOk, "配药确认", mfrmList.hWnd, 0).Tag = "配药确认"
        .InsertItem(mconTab_Recipe_Dosage, "待配药", mfrmList.hWnd, 0).Tag = "待配药_"
        .InsertItem(mconTab_Recipe_Abolish, "已配药", mfrmList.hWnd, 0).Tag = "已配药_"
        .InsertItem(mconTab_Recipe_Send, "待发药", mfrmList.hWnd, 0).Tag = "待发药_"
        .InsertItem(mconTab_Recipe_OverTime, "超时未发", mfrmList.hWnd, 0).Tag = "超时未发_"
        .InsertItem(mconTab_Recipe_Return, "退药", mfrmList.hWnd, 0).Tag = "退药_"
        
        .Item(mconTab_Recipe_Send).Selected = True
        .Item(mconTab_Recipe_DosageOk).Visible = False
        .Item(mconTab_Recipe_Abolish).Visible = False
        
        If mParams.blnMustDosageProcess = True Then
'            .Item(mconTab_Recipe_Abolish).Selected = True
            .Item(mconTab_Recipe_Dosage).Selected = True
        Else
            .Item(mconTab_Recipe_Dosage).Visible = False
'            .Item(mconTab_Recipe_Abolish).Visible = False
        End If
        
'        If mParams.blnMustDosageOkProcess = True And InStr(1, mstrPrivs, "配药确认") > 0 Then
'            .Item(mconTab_Recipe_DosageOk).Selected = True
'        Else
'            .Item(mconTab_Recipe_DosageOk).Visible = False
'        End If
'
        If mParams.blnMustDosageOkProcess = False And mParams.blnMustDosageProcess = False Then
            .Item(mconTab_Recipe_Send).Selected = True
        End If
        
        If Not .Item(mconTab_Recipe_OverTime) Is Nothing Then
            .Item(mconTab_Recipe_OverTime).Visible = (mParams.intOverTime > 0)
        End If
    End With
End Sub

Private Sub tbcDetail_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'    If Item.Index = 0 Then
'        If Not mfrmDetail Is Nothing Then mfrmDetail.ShowList mcondition.intListType, mcondition.bln显示过程单据
'    Else
'        If Not mfrmRecipe Is Nothing Then mfrmRecipe.ShowRecipe mcondition.intListType
'    End If
    
    mintTab = Item.index
End Sub

Private Sub tbcList_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim strTitleCon As String
    Dim strTitleList As String
    Dim objPaneCon As Pane
    '主窗体的菜单状态
    Dim cbrControl As CommandBarControl
    Dim cbrMenu As CommandBarControl
    
    If Item.Tag = "" Then Exit Sub
    
    Select Case Item.index
        Case mListType.配药确认
            strTitleCon = "配药确认"
            strTitleList = "处方列表(配药确认)"
        Case mListType.待配药
            strTitleCon = "待配药"
            strTitleList = "处方列表(待配药)"
        Case mListType.已配药
            strTitleCon = "已配药"
            strTitleList = "处方列表(已配药)"
        Case mListType.待发药
            strTitleCon = "待发药"
            strTitleList = "处方列表(待发药)"
        Case mListType.超时未发
            strTitleCon = "待发药"
            strTitleList = "处方列表(超时未发)"
        Case mListType.退药
            strTitleCon = "退药"
            strTitleList = "处方列表(退药)"
    End Select
    
    If mPrives.bln允许查询所有时间范围单据 Then
        If mPrives.bln修改过滤日期 = False Then
            '无权限时退药不能指定时间范围查询
            If Item.index = mListType.退药 Then
                With cbo时间范围
                    .Clear
                    .AddItem "0-当天"
                    .AddItem "1-两天内"
                    .AddItem "2-三天内"
                    
                    If Val(.Tag) = 3 Then
                        .ListIndex = 0
                        Call picConMain_Resize
                        Call picCondition_Resize
                    Else
                        .ListIndex = Val(.Tag)
                    End If
                     .Tag = .ListIndex
                End With
            
                Me.Dtp开始时间.Enabled = False
                Me.Dtp结束时间.Enabled = False
            Else
                With cbo时间范围
                    .Clear
                    .AddItem "0-当天"
                    .AddItem "1-两天内"
                    .AddItem "2-三天内"
                    .AddItem "3-指定时间范围"
                    
                    .ListIndex = Val(.Tag)
                 End With
        
                Me.Dtp开始时间.Enabled = True
                Me.Dtp结束时间.Enabled = True
            End If
        End If
    End If
             
    Chk显示退药待发单据.Visible = (Item.index <> mListType.退药 And Item.index <> mListType.配药确认)
    Chk显示过程单据.Visible = (Item.index = mListType.退药)
    Me.chk显示已确认单据.Visible = (Item.index = mListType.配药确认)
    
    Me.dkpMain.FindPane(mconPane_Recipe_Condition).Title = mstrStockName & ":" & strTitleCon
    
    If Not mfrmList Is Nothing Then
        mfrmList.ShowList Item.index, imgFilter.BorderStyle, (mParams.blnStartCall And mParams.blnStartQueue), mParams.blnMustDosageOkProcess, mParams.blnMustDosageProcess, mParams.bln启用审方
    End If
    
    If Not mfrmDetail Is Nothing Then
        mfrmDetail.ShowList Item.index, mcondition.bln显示过程单据
    End If
    
    If Not mfrmRecipe Is Nothing Then
        mfrmRecipe.ShowRecipe Item.index
    End If
    
    mcondition.intListType = Item.index
    
    SetComandBars
    
    DoEvents
    Call RefreshList(mcondition.intListType)
    
    If Me.dkpMain.FindPane(mconPane_Recipe_Condition).Hidden = False And Visible Then mfrmList.SetFocus
End Sub

Private Sub TimePrint_Timer()
    If Me.WindowState = vbMinimized Then Exit Sub
    
    '如果活动窗口不是当前窗口时退出
    If InStr(1, "frm药品处方发药New;frm处方发药明细;frm处方发药列表;frm处方", Screen.ActiveForm.Name) = 0 Then Exit Sub
'    If Screen.ActiveForm.hWnd <> Me.hWnd Then Exit Sub
    
    If mcondition.intListType = mListType.退药 Then
        Exit Sub
    End If
    
    '如果消息机制有效则不通过轮询方式自动打印
    If Not mobjMipModule Is Nothing Then
        If mobjMipModule.IsConnect = True Then
            Exit Sub
        End If
    End If
     
    TimePrint.Enabled = False
    DoEvents
    '调用打印程序
    Call AutoPrint
    DoEvents
    TimePrint.Enabled = True
End Sub
Private Function AutoPrint()
'功能：自动打印单据
    Dim recAutoPrint As New ADODB.Recordset, strErr As String
    Dim datCurr As Date, strRefresh As String, strCond As String
    Dim strUnit As String
    Dim str操作员 As String
    Dim blnInTrans As Boolean
    Dim blnIgnore As Boolean
    Dim strName As String
    Dim strSub1 As String
    Dim strSub2 As String
    Dim bln中药处方 As Boolean
    Dim str签名记录 As String
    Dim str收费类型 As String
    Dim lng病人ID As Long
    Dim strNo As String
    Dim strReturn As String
    Dim str住院 As String
    
    '根据打印参数组合条件
    '0-不打印未配药单据
    '1-打印本部门所有未配药单据
    '2-打印本窗口所有未配药单据
    '3-选择打印(发药窗口)
    If BlnInRefresh Then Exit Function
    
    On Error GoTo ErrHand
    
    If mblnIsFirst = False And mParams.bln自动配药 Then
        If mParams.int自动配药时限 > 0 Then
            If DateDiff("s", mdate上次校验时间, Sys.Currentdate) > mParams.int自动配药时限 * 60 Then
                If mParams.int校验配药人 = 1 Then
                    strName = zldatabase.UserIdentify(Me, "校验配药人", glngSys, 1341, "配药")
                
                    If Trim(strName) = "" Then Exit Function
                    mstr自动配药人 = strName
                End If
                
                mdate上次校验时间 = Sys.Currentdate
            End If
        End If
    End If
    
    '打印卫材发料清单，放前面，不受其他参数影响，要有发料模块单据打印的权限
    If mParams.int打印卫材发料清单 = 1 And IsHavePrivs(mstrStuffPrivs, "单据打印") Then
        gstrSQL = "Select NO, 单据, 填制日期, 1 As 门诊, Nvl(处方类型, 0) As 处方类型, Nvl(打印状态, 0) As 打印状态, a.病人id, a.优先级, a.姓名 " & vbNewLine & _
            "From 未发药品记录 A, 病人信息 B" & vbNewLine & _
            "Where a.病人id = b.病人id And a.库房id + 0 = [1] And a.填制日期 Between [2] And [3] And a.打印状态 = 0 And" & vbNewLine & _
            "      a.单据 In (24, 25)"
            
        '排除异常单据
        gstrSQL = gstrSQL & " And Not Exists (Select 1 From 门诊费用记录 C Where c.No = a.No And c.执行部门id = a.库房id And Decode(a.单据, 24, 1, 25, 2) = c.记录性质 And c.执行状态 = 9) "
        
        '关联费用记录，判断门诊或住院
        gstrSQL = gstrSQL & " And Exists (Select 1 From 门诊费用记录 C Where a.No = c.No And a.库房id = c.执行部门id And Decode(a.单据, 24, 1, 25, 2) = c.记录性质 ) "
        
        '门诊和住院
        str住院 = gstrSQL
        str住院 = Replace(str住院, "1 As 门诊", "2 As 门诊")
        str住院 = Replace(str住院, "门诊费用记录", "住院费用记录")
        gstrSQL = gstrSQL & " Union All " & str住院
    
        gstrSQL = gstrSQL & " Order by 优先级,姓名,No"
        
        Set recAutoPrint = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
            mSQLCondition.lng药房ID, _
            mSQLCondition.date开始日期, _
            mSQLCondition.date结束日期)
        
        datCurr = Sys.Currentdate()
        
        With recAutoPrint
            If .RecordCount > 0 Then
                If DateDiff("s", !填制日期, datCurr) > mParams.lngPrintDelay Then
                    Do While Not .EOF
                        '更新打印状态
                        gstrSQL = "Zl_未发药品记录_更新打印状态("
                        '单据
                        gstrSQL = gstrSQL & Val(!单据)
                        'NO
                        gstrSQL = gstrSQL & ",'" & !NO & "'"
                        '库房ID
                        gstrSQL = gstrSQL & "," & mParams.lng药房ID
                        '来源科室
                        gstrSQL = gstrSQL & ",Null"
                        '打印内容
                        gstrSQL = gstrSQL & ",1"
                        gstrSQL = gstrSQL & ")"
                        
                        Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-更新单据已打印")
                        
                        '打印单据
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1723", Me, "库房==" & mParams.lng药房ID, "NO=" & !NO, "单据=" & Val(!单据), "审核人=审核人 is null", 2)
                        
                        .MoveNext
                    Loop
                End If
            End If
        End With
    End If
    '以上打印卫材发料清单
    
    gstrSQL = " Select  NO,单据,填制日期,1 As 门诊,Nvl(处方类型,0) As 处方类型,Nvl(打印状态,0) As 打印状态,A.病人ID, a.优先级, a.姓名 " & _
               " From 未发药品记录 A, 病人信息 B "
    gstrSQL = gstrSQL & " Where 库房ID+0=[1] " & _
               " And 填制日期 Between [2] And [3] " & _
               " And 打印状态 Not In (1,2) "
    
    gstrSQL = gstrSQL & IIf(mParams.blnMustDosageOkProcess, " And A.排队状态=1 ", "")
    gstrSQL = gstrSQL & IIf(mParams.bln自动配药, " And 配药人 Is Null ", "")
    gstrSQL = gstrSQL & " And A.病人ID=B.病人ID" & IIf(mSQLCondition.str医保号 <> "", "", "(+)") & ""
    
    Select Case mParams.intPrint
        Case 0
            If mParams.intPrintDrugLable = 0 Then Exit Function
        Case 1
            If Not mParams.bln记帐单 Then gstrSQL = gstrSQL & " And 单据=8"
        Case 2
            If mParams.bln记帐单 Then
                If mParams.Str窗口 <> "" Then
                    gstrSQL = gstrSQL & " And A.单据 In (8,9) And A.发药窗口 In (Select * From Table(Cast(f_Str2list([19]) As Zltools.t_Strlist))) "
                End If
            Else
                If mParams.Str窗口 <> "" Then
                    gstrSQL = gstrSQL & " And A.单据=8 And A.发药窗口 In (Select * From Table(Cast(f_Str2list([19]) As Zltools.t_Strlist))) "
                Else
                    gstrSQL = gstrSQL & " And A.单据=8"
                End If
            End If
        Case 3
            If mParams.bln记帐单 Then
                If mParams.strPrintWindow <> "" Then
                    gstrSQL = gstrSQL & " And A.单据 In (8,9) And A.发药窗口 In (Select * From Table(Cast(f_Str2list([20]) As Zltools.t_Strlist))) "
                End If
            Else
                If mParams.strPrintWindow <> "" Then
                    gstrSQL = gstrSQL & " And A.单据=8 And A.发药窗口 In (Select * From Table(Cast(f_Str2list([20]) As Zltools.t_Strlist))) "
                Else
                    gstrSQL = gstrSQL & " And A.单据=8"
                End If
            End If
    End Select
    
    If mcondition.int服务对象 = 2 Then
        gstrSQL = gstrSQL & " And A.单据 = 9 And A.主页ID Is Not NULL " '仅住院记帐
    Else
        gstrSQL = gstrSQL & " And A.单据 In (8,9)" '门诊及住院所有单据
    End If
    
    If mSQLCondition.str开始NO <> "" Or mSQLCondition.str结束NO <> "" Then
        If mSQLCondition.str开始NO <> "" And mSQLCondition.str结束NO <> "" Then
            gstrSQL = gstrSQL & " And A.NO Between [4] And [5] "
        Else
            If mSQLCondition.str开始NO <> "" Then
                gstrSQL = gstrSQL & " And A.NO = [4] "
            Else
                gstrSQL = gstrSQL & " And A.NO = [5] "
            End If
        End If
    End If
    
    If mSQLCondition.str姓名 <> "" Then gstrSQL = gstrSQL & " And Upper(A.姓名) Like [6] "
    
    If mSQLCondition.str就诊卡 <> "" Then gstrSQL = gstrSQL & " And Upper(B.就诊卡号) = [7] "
    
    If mSQLCondition.str标识号 <> "" Then gstrSQL = gstrSQL & " And Upper(DECODE(A.单据,8,B.门诊号,B.住院号)) Like [8] "
    
    If mSQLCondition.lng科室ID > 0 Then gstrSQL = gstrSQL & " And A.对方部门ID+0=[9] "
    
    If mSQLCondition.str当前NO <> "" Then gstrSQL = gstrSQL & " And A.NO=[13] "
    
    If mSQLCondition.str门诊号 <> "" Then gstrSQL = gstrSQL & " And B.门诊号=[14] "
    
'    If mSQLCondition.str身份证 <> "" Then gstrSQL = gstrSQL & " And B.身份证号=[15] "

    If mSQLCondition.str身份证 <> "" Then gstrSQL = gstrSQL & " And B.病人ID=[15] "
    
    If mSQLCondition.lng病人ID <> 0 Or (Me.txtPati.Text <> "" And mParams.int输入模式 = mFindType.IC卡) Then gstrSQL = gstrSQL & " And B.病人id=[16] "
    
    If mSQLCondition.str医保号 <> "" Then gstrSQL = gstrSQL & " And B.医保号=[17] "
    
    If mSQLCondition.lng住院号 <> 0 Then gstrSQL = gstrSQL & " And B.住院号=[18] "
    
    Select Case mParams.intShowBill收费
        Case 0  '不显示处方
            strSub1 = "1=2"
        Case 1  '显示未收费
            strSub1 = "A.单据<>9 And Nvl(A.已收费,0)=0 And A.单据=8"
        Case 2  '显示已收费
            strSub1 = "A.单据<>9 And A.已收费=1 And A.单据=8"
        Case 3  '显示所有处方
            strSub1 = "A.单据<>9 And A.单据=8"
    End Select
    Select Case mParams.intShowBill记帐
        Case 0  '不显示处方
            strSub2 = "1=2"
        Case 1  '显示未审核
            strSub2 = "A.单据<>8 And Nvl(A.已收费,0)=0 And A.单据=9"
        Case 2  '显示已审核
            strSub2 = "A.单据<>8 And A.已收费=1 And A.单据=9"
        Case 3  '显示所有处方
            strSub2 = "A.单据<>8 And A.单据=9"
    End Select
    
    gstrSQL = gstrSQL & " And A.单据 IN(8,9) And (" & strSub1 & " Or " & strSub2 & ")"
    
    gstrSQL = gstrSQL & IIf(mParams.strSourceDep = "", "", " And A.对方部门id+0 In (Select * From Table(Cast(f_Num2list([21]) As Zltools.t_Numlist)))")
    
    '排除异常单据
    gstrSQL = gstrSQL & " And Not Exists (Select 1 From 门诊费用记录 C Where c.No = a.No And c.执行部门id = a.库房id And c.执行状态 = 9) "
    
    '关联费用记录，判断门诊或住院
    gstrSQL = gstrSQL & " And Exists (Select 1 From 门诊费用记录 C Where a.No = c.No And a.库房id = c.执行部门id And Decode(a.单据, 8, 1, 9, 2) = c.记录性质) "
    
    '联合门诊和住院费用记录查询
    If mcondition.int服务对象 = 2 Then
        '仅住院
        gstrSQL = Replace(gstrSQL, "1 As 门诊", "2 As 门诊")
        gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
    ElseIf mcondition.int服务对象 = 3 Then
        '门诊和住院
        str住院 = gstrSQL
        str住院 = Replace(str住院, "1 As 门诊", "2 As 门诊")
        str住院 = Replace(str住院, "门诊费用记录", "住院费用记录")
        gstrSQL = gstrSQL & " Union All " & str住院
    End If
    
    gstrSQL = gstrSQL & " Order by 优先级,姓名,No"
    
    On Error GoTo ErrHand
    BlnInRefresh = True
    
    recAutoPrint.Sort = "病人id"
    
    If mSQLCondition.str身份证 <> "" And Not mobjSquareCard Is Nothing Then
        If mobjSquareCard.zlGetPatiID("身份证", UCase(mSQLCondition.str身份证), False, lng病人ID) = False Then lng病人ID = 0
        mSQLCondition.lng病人ID = lng病人ID
    End If
    
    Set recAutoPrint = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
            mSQLCondition.lng药房ID, _
            mSQLCondition.date开始日期, _
            mSQLCondition.date结束日期, _
            mSQLCondition.str开始NO, _
            mSQLCondition.str结束NO, _
            mSQLCondition.str姓名, _
            mSQLCondition.str就诊卡, _
            mSQLCondition.str标识号, _
            mSQLCondition.lng科室ID, _
            mSQLCondition.str填制人, _
            mSQLCondition.str审核人, _
            mSQLCondition.lng药品id, _
            mSQLCondition.str当前NO, _
            mSQLCondition.str门诊号, _
            lng病人ID, _
            mSQLCondition.lng病人ID, _
            mSQLCondition.str医保号, _
            mSQLCondition.lng住院号, _
            mParams.Str窗口, _
            mParams.strPrintWindow, _
            mParams.strSourceDep)

    datCurr = Sys.Currentdate()
        
    With recAutoPrint
        Do While Not .EOF
            '打印单据
            If DateDiff("s", !填制日期, datCurr) > mParams.lngPrintDelay Then
                If mParams.intPrint > 0 Then
                    If mParams.bln自动配药 = True And IsDosage(Val(!单据), !NO, Val(!门诊)) Then
                        '处理自动配药，在打印前完成
                        blnIgnore = False

'                        '检查是否需要配药
'                        If Not IsDosage(Val(!单据), !NO, Val(!门诊)) Then
'                            blnIgnore = True
'                        End If

                        '检测是否允许
                        If CheckBill(mSQLCondition.lng药房ID, 1, Val(!单据), !NO, Val(!门诊), Val(!门诊)) <> 0 Then
                            blnIgnore = True
                        End If

                        If blnIgnore = False Then
                            gcnOracle.BeginTrans
                            blnInTrans = True

                            '再设置配药人
                            str操作员 = IIf(mstr自动配药人 <> "", mstr自动配药人, IIf(mParams.str配药人 = "|当前操作员|", gstrUserName, mParams.str配药人))

                            gstrSQL = "zl_药品收发记录_设置配药人("
                            '库房ID
                            gstrSQL = gstrSQL & mParams.lng药房ID
                            '单据
                            gstrSQL = gstrSQL & "," & Val(!单据)
                            'NO
                            gstrSQL = gstrSQL & ",'" & !NO & "'"
                            '门诊
                            gstrSQL = gstrSQL & "," & Val(!门诊)
                            '配药人
                            gstrSQL = gstrSQL & ",'" & str操作员 & "'"
                            '配药日期
                            gstrSQL = gstrSQL & ",to_date('" & datCurr & "','yyyy-MM-dd hh24:mi:ss') "
                            gstrSQL = gstrSQL & ")"

                            Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-设置配药人")
                            
                            If Val(!打印状态) <> 3 Then
                                gstrSQL = "Zl_未发药品记录_更新打印状态("
                                '单据
                                gstrSQL = gstrSQL & Val(!单据)
                                'NO
                                gstrSQL = gstrSQL & ",'" & !NO & "'"
                                '库房ID
                                gstrSQL = gstrSQL & "," & mParams.lng药房ID
                                '来源科室
                                gstrSQL = gstrSQL & "," & IIf(mParams.strSourceDep = "", "Null", "'" & mParams.strSourceDep & "'")
                                '打印内容
                                gstrSQL = gstrSQL & ",3"
                                gstrSQL = gstrSQL & ")"
                                
                                Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-更新单据已打印")
                                
                                strUnit = GetUnit(mParams.lng药房ID, !单据, !NO, Val(!门诊))
                                str收费类型 = BillHaveHerial(!NO, !单据, Val(!门诊))
                                
                                If mParams.bln打印所有格式 Then
                                    If InStr(1, str收费类型, "7") <> 0 And (InStr(1, str收费类型, "5") <> 0 Or InStr(1, str收费类型, "6") <> 0) Then
                                        SetLocatePrinter !处方类型
                                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
                                            "NO=" & !NO, "性质=" & IIf(!单据 = 8, 1, 2), "药房=" & mParams.lng药房ID, "包装系数=" & IIf(strUnit = "门诊单位", "C.门诊包装", "C.住院包装"), "PrintEmpty=0", 2)
                                        
                                        '恢复处方签的本地打印机设置
                                        Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTDefaultScheme_Recipt, mParams.strDefaultPrinter)
                                                                        
                                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
                                            "NO=" & !NO, "性质=" & IIf(!单据 = 8, 1, 2), "PrintEmpty=0", 2)
                                    ElseIf (InStr(1, str收费类型, "5") <> 0 Or InStr(1, str收费类型, "6") <> 0) Then
                                        SetLocatePrinter !处方类型
                                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
                                            "NO=" & !NO, "性质=" & IIf(!单据 = 8, 1, 2), "药房=" & mParams.lng药房ID, "包装系数=" & IIf(strUnit = "门诊单位", "C.门诊包装", "C.住院包装"), "PrintEmpty=0", 2)
                                        
                                        '恢复处方签的本地打印机设置
                                        Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTDefaultScheme_Recipt, mParams.strDefaultPrinter)
                                    Else
                                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
                                            "NO=" & !NO, "性质=" & IIf(!单据 = 8, 1, 2), "PrintEmpty=0", 2)
                                    End If

                                Else
                                    If InStr(1, str收费类型, "7") <> 0 And (InStr(1, str收费类型, "5") <> 0 Or InStr(1, str收费类型, "6") <> 0) Then
                                        SetLocatePrinter !处方类型, Val(Split(mParams.str配药格式, ";")(0)) - 1
                                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
                                            "NO=" & !NO, "性质=" & IIf(!单据 = 8, 1, 2), "药房=" & mParams.lng药房ID, "包装系数=" & IIf(strUnit = "门诊单位", "C.门诊包装", "C.住院包装"), "ReportFormat=" & Val(Split(mParams.str配药格式, ";")(0)), "PrintEmpty=0", 2)
                                        
                                        '恢复处方签的本地打印机设置
                                        Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTDefaultScheme_Recipt, mParams.strDefaultPrinter)
                                                                        
                                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
                                            "NO=" & !NO, "性质=" & IIf(!单据 = 8, 1, 2), "ReportFormat=" & Val(Split(mParams.str配药格式, ";")(1)), "PrintEmpty=0", 2)
                                    ElseIf (InStr(1, str收费类型, "5") <> 0 Or InStr(1, str收费类型, "6") <> 0) Then
                                        SetLocatePrinter !处方类型, Val(Split(mParams.str配药格式, ";")(0)) - 1
                                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
                                            "NO=" & !NO, "性质=" & IIf(!单据 = 8, 1, 2), "药房=" & mParams.lng药房ID, "包装系数=" & IIf(strUnit = "门诊单位", "C.门诊包装", "C.住院包装"), "ReportFormat=" & Val(Split(mParams.str配药格式, ";")(0)), "PrintEmpty=0", 2)
                                        
                                        '恢复处方签的本地打印机设置
                                        Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTDefaultScheme_Recipt, mParams.strDefaultPrinter)
                                    Else
                                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
                                            "NO=" & !NO, "性质=" & IIf(!单据 = 8, 1, 2), "ReportFormat=" & Val(Split(mParams.str配药格式, ";")(1)), "PrintEmpty=0", 2)
                                    End If
                                End If
                            End If

                            '如果已启用了电子签名，则需要对配药人进行电子签名处理
                            If gblnESign处方发药 = True And gblnESignUserStoped = False Then
                                str签名记录 = ""
                                If GetSignatureRecored(EsignTache.Dosage, Val(!单据), !NO, mParams.lng药房ID, str签名记录) = False Then
                                    If blnInTrans = True Then gcnOracle.RollbackTrans
                                    Exit Function
                                End If
                                
                                If str签名记录 <> "" Then
                                    gstrSQL = "Zl_药品签名记录_Insert(" & str签名记录 & ")"
                                    
                                    Call zldatabase.ExecuteProcedure(gstrSQL, "签名记录")
                                End If
                            End If

                            gcnOracle.CommitTrans
                            blnInTrans = False

                            mblnIsFirst = False
                        End If
                    ElseIf Val(!打印状态) <> 3 Then
                        gstrSQL = "Zl_未发药品记录_更新打印状态("
                        '单据
                        gstrSQL = gstrSQL & Val(!单据)
                        'NO
                        gstrSQL = gstrSQL & ",'" & !NO & "'"
                        '库房ID
                        gstrSQL = gstrSQL & "," & mParams.lng药房ID
                        '来源科室
                        gstrSQL = gstrSQL & "," & IIf(mParams.strSourceDep = "", "Null", "'" & mParams.strSourceDep & "'")
                        '打印内容
                        gstrSQL = gstrSQL & ",3"
                        gstrSQL = gstrSQL & ")"
                        
                        Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-更新单据已打印")
                        
                        strUnit = GetUnit(mParams.lng药房ID, !单据, !NO, Val(!门诊))
                        str收费类型 = BillHaveHerial(!NO, !单据, Val(!门诊))
                        
                        If mParams.bln打印所有格式 Then
                            If InStr(1, str收费类型, "7") <> 0 And (InStr(1, str收费类型, "5") <> 0 Or InStr(1, str收费类型, "6") <> 0) Then
                                SetLocatePrinter !处方类型
                                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
                                    "NO=" & !NO, "性质=" & IIf(!单据 = 8, 1, 2), "药房=" & mParams.lng药房ID, "包装系数=" & IIf(strUnit = "门诊单位", "C.门诊包装", "C.住院包装"), "PrintEmpty=0", 2)
                                
                                '恢复处方签的本地打印机设置
                                Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTDefaultScheme_Recipt, mParams.strDefaultPrinter)
                                                        
                                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
                                    "NO=" & !NO, "性质=" & IIf(!单据 = 8, 1, 2), "PrintEmpty=0", 2)
                            ElseIf (InStr(1, str收费类型, "5") <> 0 Or InStr(1, str收费类型, "6") <> 0) Then
                                SetLocatePrinter !处方类型, -1
                                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
                                    "NO=" & !NO, "性质=" & IIf(!单据 = 8, 1, 2), "药房=" & mParams.lng药房ID, "包装系数=" & IIf(strUnit = "门诊单位", "C.门诊包装", "C.住院包装"), "PrintEmpty=0", 2)
                                
                                '恢复处方签的本地打印机设置
                                Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTDefaultScheme_Recipt, mParams.strDefaultPrinter)
                            Else
                                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
                                    "NO=" & !NO, "性质=" & IIf(!单据 = 8, 1, 2), "PrintEmpty=0", 2)
                            End If

                        Else
                        
                            If InStr(1, str收费类型, "7") <> 0 And (InStr(1, str收费类型, "5") <> 0 Or InStr(1, str收费类型, "6") <> 0) Then
                                SetLocatePrinter !处方类型, Val(Split(mParams.str配药格式, ";")(0)) - 1
                                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
                                    "NO=" & !NO, "性质=" & IIf(!单据 = 8, 1, 2), "药房=" & mParams.lng药房ID, "包装系数=" & IIf(strUnit = "门诊单位", "C.门诊包装", "C.住院包装"), "ReportFormat=" & Val(Split(mParams.str配药格式, ";")(0)), "PrintEmpty=0", 2)
                                
                                '恢复处方签的本地打印机设置
                                Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTDefaultScheme_Recipt, mParams.strDefaultPrinter)
                                                        
                                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
                                    "NO=" & !NO, "性质=" & IIf(!单据 = 8, 1, 2), "ReportFormat=" & Val(Split(mParams.str配药格式, ";")(1)), "PrintEmpty=0", 2)
                            ElseIf (InStr(1, str收费类型, "5") <> 0 Or InStr(1, str收费类型, "6") <> 0) Then
                                SetLocatePrinter !处方类型, Val(Split(mParams.str配药格式, ";")(0)) - 1
                                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
                                    "NO=" & !NO, "性质=" & IIf(!单据 = 8, 1, 2), "药房=" & mParams.lng药房ID, "包装系数=" & IIf(strUnit = "门诊单位", "C.门诊包装", "C.住院包装"), "ReportFormat=" & Val(Split(mParams.str配药格式, ";")(0)), "PrintEmpty=0", 2)
                                
                                '恢复处方签的本地打印机设置
                                Call SavePrinterSet("ZL1_BILL_1341_3", mstrRPTDefaultScheme_Recipt, mParams.strDefaultPrinter)
                            Else
                                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
                                    "NO=" & !NO, "性质=" & IIf(!单据 = 8, 1, 2), "ReportFormat=" & Val(Split(mParams.str配药格式, ";")(1)), "PrintEmpty=0", 2)
                            End If
                        End If
                    End If
                End If
                
                '如果前面没有判断是否是中药处方，这里再处理
                If mParams.intPrint <= 0 Then
                    str收费类型 = BillHaveHerial(!NO, !单据, Val(!门诊))
                End If
                
                '打印药品标签
                If mParams.intPrintDrugLable = 1 And Val(!打印状态) <> 4 Then
                    gstrSQL = "Zl_未发药品记录_更新打印状态("
                    '单据
                    gstrSQL = gstrSQL & Val(!单据)
                    'NO
                    gstrSQL = gstrSQL & ",'" & !NO & "'"
                    '库房ID
                    gstrSQL = gstrSQL & "," & mParams.lng药房ID
                    '来源科室
                    gstrSQL = gstrSQL & "," & IIf(mParams.strSourceDep = "", "Null", "'" & mParams.strSourceDep & "'")
                    '打印内容
                    gstrSQL = gstrSQL & ",4"
                    gstrSQL = gstrSQL & ")"
                    
                    Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-更新单据已打印")
                        
                    If InStr(1, str收费类型, "7") <> 0 And (InStr(1, str收费类型, "5") <> 0 Or InStr(1, str收费类型, "6") <> 0) Then
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_6", Me, _
                            "NO=" & !NO, "性质=" & IIf(!单据 = 8, 1, 2), "药房=" & mParams.lng药房ID, "包装系数=" & IIf(strUnit = "门诊单位", "D.门诊包装", "D.住院包装"), "PrintEmpty=0", 2)
                        
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_7", Me, _
                            "NO=" & !NO, "药房=" & mParams.lng药房ID, "性质=" & IIf(!单据 = 8, 1, 2), "PrintEmpty=0", 2)
                    ElseIf (InStr(1, str收费类型, "5") <> 0 Or InStr(1, str收费类型, "6") <> 0) Then
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_6", Me, _
                            "NO=" & !NO, "性质=" & IIf(!单据 = 8, 1, 2), "药房=" & mParams.lng药房ID, "包装系数=" & IIf(strUnit = "门诊单位", "D.门诊包装", "D.住院包装"), "PrintEmpty=0", 2)
                    Else
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_7", Me, _
                            "NO=" & !NO, "药房=" & mParams.lng药房ID, "性质=" & IIf(!单据 = 8, 1, 2), "PrintEmpty=0", 2)
                    End If
                End If
            End If

            .MoveNext
        Loop
        If .RecordCount > 0 Then .MoveFirst
    End With
    BlnInRefresh = False
    Exit Function
ErrHand:
    If blnInTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetOpr()
'获取发药窗口的编号
    Dim strsql As String
    Dim rstemp As Recordset
    
    On Error GoTo errHandle
    strsql = "Select 编码 From 发药窗口 Where 药房id=[1] And 名称=[2]"
    Set rstemp = zldatabase.OpenSQLRecord(strsql, "GetOpr", mParams.lng药房ID, mParams.Str窗口)
    
    If Not rstemp.EOF Then
        mstrOpr = rstemp!编码
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function IsDosage(ByVal Int单据 As Integer, ByVal strNo As String, ByVal int门诊 As Integer) As Boolean
    '检查当前处方是否需要经过配药过程
    
    On Error GoTo ErrHand
    
    If Int单据 = 0 Then Exit Function
    If strNo = "" Then Exit Function
    
    If mrsIsDosage Is Nothing Then
        GetDosage mParams.lng药房ID
        If mrsIsDosage Is Nothing Then
            Exit Function
        End If
    End If
    
    mrsIsDosage.Filter = "门诊=" & int门诊
    If mrsIsDosage.EOF Then Exit Function

    IsDosage = (mrsIsDosage!配药 = 1)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub TimePrintCancelBill_Timer()
    Dim curDateBegin As Date
    Dim curDateEnd As Date
    
    '调用打印退费单
    IntTimes = IntTimes + 1
    '不到分钟数退出
    If IntTimes < mParams.lngPrintBackInterval Then Exit Sub
    IntTimes = 0
    
    curDateEnd = Format(Sys.Currentdate, "yyyy-MM-dd hh:mm:ss")
    curDateBegin = DateAdd("n", 0 - mParams.lngPrintBackInterval, curDateEnd)
    
    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_8", Me, "开始时间=" & Format(curDateBegin, "yyyy-MM-dd hh:mm"), "结束时间=" & Format(curDateEnd, "yyyy-MM-dd hh:mm"), "药房=" & mParams.lng药房ID, 2)
End Sub

Private Sub TimeRefresh_Timer()
    '处理自动刷新未知错误
    Dim thwnd As Long
    
    On Error Resume Next
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    '如果活动窗口不是当前窗口时退出
    If InStr(1, "frm药品处方发药New;frm处方发药明细;frm处方发药列表;frm处方", Screen.ActiveForm.Name) = 0 Then Exit Sub
'    If Screen.ActiveForm.hWnd <> Me.hWnd Then Exit Sub

    thwnd = GetForegroundWindow()
    If thwnd <> Me.hWnd Then Exit Sub
    
    If mcondition.intListType = mListType.退药 Then Exit Sub
    
    '如果发药未结束则退出
    If mblnSendIsOver = False Then Exit Sub
    
    '如果消息机制有效则不通过轮询方式自动刷新
    If Not mobjMipModule Is Nothing Then
        If mobjMipModule.IsConnect = True Then
            '当前是待配药或待发药界面时不执行轮询刷新
            If (mParams.blnMustDosageProcess = True And mcondition.intListType = mListType.待配药) Or _
                (mParams.blnMustDosageProcess = False And mcondition.intListType = mListType.待发药) Then
                Exit Sub
            End If
        End If
    End If
    
    TimeRefresh.Enabled = False
    DoEvents
        RefreshList mcondition.intListType
    DoEvents
    TimeRefresh.Enabled = True
End Sub

Private Sub tmrCall_Timer()
    '本机作为其他机器的远端呼叫机器使用时，才开启时间控件
    
    '调用呼叫主程序
    Call zlCallMain
End Sub

Private Sub tmrMsgRefresh_Timer()
    If Me.WindowState = vbMinimized Then Exit Sub
    
    If Screen.ActiveForm Is Nothing Then Exit Sub
    '如果活动窗口不是当前窗口时退出
    If InStr(1, "frm药品处方发药New;frm处方发药明细;frm处方发药列表;frm处方", Screen.ActiveForm.Name) = 0 Then Exit Sub
'    If Screen.ActiveForm.hWnd <> Me.hWnd Then Exit Sub
    
    '退药业务时不进行自动刷新或打印
    If mcondition.intListType = mListType.退药 Then Exit Sub
    
    '消息机制无效时不继续
    If mobjMipModule Is Nothing Then Exit Sub
    If mobjMipModule.IsConnect = False Then Exit Sub
     
    '有消息时刷新或打印
    If mblnExistMsg = True Then
        '刷新前先关闭计时器，刷新完后再开启
        tmrMsgRefresh.Enabled = False
        DoEvents
        
        '当前是待配药或待发药界面则刷新
        If (mParams.blnMustDosageProcess = True And mcondition.intListType = mListType.待配药) Or _
            (mParams.blnMustDosageProcess = False And mcondition.intListType = mListType.待发药) Then
            Call RefreshList(mcondition.intListType)
        End If
                
        '同时也进行自动打印
        If mParams.intPrint > 0 Then
            Call AutoPrint
        End If
        
        DoEvents
        tmrMsgRefresh.Enabled = True
                
        '刷新后记录当前刷新时间
        mdteMsgRefresh = Now
        
        '消息变量设为False
        mblnExistMsg = False
    End If
End Sub

Private Sub txtPati_Change()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtPati.Text = "" And Me.ActiveControl Is txtPati)
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (txtPati.Text = "" And Me.ActiveControl Is txtPati)
    
End Sub


Private Sub txtPati_GotFocus()
    txtPati.BackColor = &HE1FEDA
    
    If Not mobjIDCard Is Nothing And txtPati.Text = "" Then
        Call mobjIDCard.SetEnabled(True)
    End If
    
    If Not mobjICCard Is Nothing And txtPati.Text = "" Then
        Call mobjICCard.SetEnabled(True)
    End If
    
    txtPati.Text = ""
    Call zlControl.TxtSelAll(txtPati)
    
    mblnInput = True
End Sub

Private Sub txtPati_KeyPress(KeyAscii As Integer)
    Dim blnDoIt As Boolean
    Dim strInput As String
    Dim strCondition As String
    Dim i As Integer
    Dim bln刷卡 As Boolean
    Dim blnSta As Boolean
    Dim lng病人ID As Long
    Dim rsData As Recordset
    Dim str姓名 As String
    Dim str病人id As String
    Dim strCard As String
    Dim strRecipeString As String
    Dim arrRecipe
    Dim intCount As Integer, n As Integer
    Dim strNos As String
    Dim strReturn As String
    
    strCard = IDKNType.GetCurCard.名称
    If KeyAscii = 13 Then
        KeyAscii = 0
        blnDoIt = True
        
        If strCard = "IC卡" Then
            If Not mobjSquareCard Is Nothing Then Call mobjSquareCard.zlGetPatiID("IC卡", UCase(Trim(txtPati.Text)), False, mlngIC病人id)
            If txtPati.Text <> "" Then blnDoIt = True
        Else
            If Trim(txtPati.Text) <> "" Then blnDoIt = True
        End If
        
        If Not (strCard = "姓名" Or strCard = "单据号" Or strCard = "住院号" Or strCard = "医保号" Or strCard = "身份证" Or strCard = "门诊号") And KeyAscii <> 8 Then bln刷卡 = True
    ElseIf KeyAscii <> 13 Then
        mblnCard = False
        mblnScaner = False
        If strCard = "姓名" Then
            '姓名类别
            mblnCard = zlCommFun.InputIsCard(txtPati, KeyAscii, glngSys)
        ElseIf mcondition.intListType = mListType.待配药 And mParams.int输入模式 = mFindType.单据号 And mParams.bln配药扫描 = True Then
            mblnScaner = InputIsScaner(txtPati, KeyAscii)
        ElseIf mcondition.intListType = mListType.待发药 And mParams.bln扫描后呼叫 = True Then
            mblnScaner = InputIsScaner(txtPati, KeyAscii)
        Else
            mblnScaner = InputIsScaner(txtPati, KeyAscii)
        End If
        
        If mblnCard Then
            If strCard = "姓名" Then
                If Len(txtPati.Text) = mint就诊卡长度 - 1 And KeyAscii <> 8 And txtPati.SelLength <> Len(txtPati.Text) Then
                    txtPati.Text = txtPati.Text & Chr(KeyAscii)
                    txtPati.SelStart = Len(txtPati.Text)
                    KeyAscii = 0: blnDoIt = True
                End If
            End If
        Else
            Select Case strCard
'                Case mFindType.就诊卡
'                    If InStr(":：;；?？''||" & Chr(22), Chr(KeyAscii)) > 0 Then
'                        KeyAscii = 0
'                    Else
'                        KeyAscii = Asc(UCase(Chr(KeyAscii)))
'                    End If
                Case "门诊号"
                    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
                Case "单据号"
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                    If mcondition.intListType = mListType.待配药 And mblnScaner And mParams.bln配药扫描 = True Then
                        txtPati.Text = txtPati.Text & UCase(Chr(KeyAscii))
                        txtPati.SelStart = Len(txtPati.Text)
                        KeyAscii = 0
                        
                        If Len(txtPati.Text) = 8 Then
                            blnDoIt = True
                            If mstrScanerLastNo <> txtPati.Text Then
                                mblnScaned = False
                                mstrScanerLastNo = txtPati.Text
                            End If
                        End If
                    Else
                        If Not (InStr("0123456789" & Chr(8), Chr(KeyAscii)) > 0 Or (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z"))) Then
                            KeyAscii = 0
                        End If
                    End If
                Case "姓名"
                    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then
                        KeyAscii = 0
                    End If
'                Case "二代身份证"
'                Case "IC卡"
                Case "住院号"
                Case "医保号"
                Case Else
                    bln刷卡 = True
'                    Me.txtPati.MaxLength = 100
                    '其他的是消费卡
                    If InStr(":：;；?？''||" & Chr(22) & Chr(32), Chr(KeyAscii)) > 0 Then
                        KeyAscii = 0
                    Else
                        KeyAscii = Asc(UCase(Chr(KeyAscii)))
                    End If
                    
                    If Len(txtPati.Text) = IDKNType.GetCardNoLen - 1 And KeyAscii <> 8 And mParams.int回车方式 = 1 Then
                        txtPati.Text = txtPati.Text & Chr(KeyAscii)
                        txtPati.SelStart = Len(txtPati.Text)
                        KeyAscii = 0
                        blnDoIt = True
                    End If
                    mstrLastBrushCardNo = txtPati.Text & IIf(KeyAscii = 0, "", Chr(KeyAscii))
            End Select
        End If
    End If
    
    If blnDoIt Then
        If mParams.blnMustDosageOkProcess And bln刷卡 And InStr(1, mstrPrivs, "配药确认") > 0 Then
            On Error GoTo errHandle
            
            If strCard = "姓名" Or strCard = "单据号" Or strCard = "住院号" Or strCard = "医保号" Or strCard = "身份证" Or strCard = "门诊号" Then
                strInput = txtPati.Text
            Else
                '消费卡类别时输入为卡ID+卡号
                strInput = mobjcard.接口序号 & "|" & txtPati.Text
            End If
            lng病人ID = zlfuncCard_GetPatiID(mobjSquareCard, Val(Split(strInput, "|")(0)), Split(strInput, "|")(1))
            
            If lng病人ID <> 0 Then
                gstrSQL = "Select distinct A.NO,A.单据 From 未发药品记录 A,药品收发记录 B Where A.NO=B.NO and A.单据=B.单据 and A.库房id=B.库房id and A.病人id=[1] And A.库房id=[2] and A.填制日期 between [3] and [4] and nvl(A.排队状态,0)=0"
                Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "", lng病人ID, mParams.lng药房ID, mSQLCondition.date开始日期, mSQLCondition.date结束日期)
                strCondition = ""
                
                If Not rsData Is Nothing Then
                    
                    If rsData.RecordCount > 0 Then
                        rsData.MoveFirst
                    Else
                        
                        Me.stbThis.Panels(2).Text = "卡号：" & txtPati.Text & "暂无处方信息！"
                        blnSta = True
                    End If
                    gcnOracle.BeginTrans
                    Do While Not rsData.EOF
                        strCondition = IIf(strCondition = "", strCondition, strCondition & " OR ") & "NO='" & rsData!NO & "'"
                        gstrSQL = "Zl_未发药品记录_配药确认("
                            'NO
                            gstrSQL = gstrSQL & "'" & rsData!NO & "'"
                            '单据
                            gstrSQL = gstrSQL & "," & rsData!单据
                            '库房ID
                            gstrSQL = gstrSQL & "," & mParams.lng药房ID
                            '配药确认
                            gstrSQL = gstrSQL & "," & 1
                            '操作员
                            gstrSQL = gstrSQL & ",'" & gstrUserName & "')"
        
                            Call zldatabase.ExecuteProcedure(gstrSQL, "RecipeWork_DosageOk")
                        rsData.MoveNext
                    Loop
                
                End If
    
                gcnOracle.CommitTrans
    
                Call RefreshList(mcondition.intListType)
            Else
                Me.stbThis.Panels(2).Text = "卡号：" & txtPati.Text & "暂无病人信息！"
                blnSta = True
            End If
        End If
        
        If strCard = "单据号" And mParams.blnMustDosageOkProcess And mParams.blnMustDosageProcess And InStr(1, mstrPrivs, "配药确认") > 0 Then
            txtPati.Text = GetFullNO(txtPati.Text, 13)
            gstrSQL = _
                "Select distinct nvl(A.病人id,'') 病人id,nvl(A.姓名,'') 姓名,A.NO,A.单据 " & _
                "From 未发药品记录 A,药品收发记录 B " & _
                "Where A.NO=B.NO and A.单据=B.单据 and A.库房id=B.库房id and A.NO=[1] And A.库房id=[2] " & _
                "    And A.填制日期 between [3] and [4] and nvl(A.排队状态,0)=0"
            Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "", Me.txtPati.Text, mParams.lng药房ID, mSQLCondition.date开始日期, mSQLCondition.date结束日期)
            
            If rsData.EOF Then
                Me.stbThis.Panels(2).Text = "NO为[" & txtPati.Text & "]的单据不存在！"
                blnSta = True
            Else
                
                Do While Not rsData.EOF
                    If zlStr.NVL(rsData!病人ID) = "" Then
                        str姓名 = str姓名 & rsData!姓名 & ","
                    Else
                        str病人id = str病人id & rsData!病人ID & ","
                    End If
                    rsData.MoveNext
                Loop
                
                If str病人id <> "" Or str姓名 <> "" Then
                    If str病人id <> "" And str姓名 = "" Then
                        gstrSQL = "Select /*+ Rule*/ distinct A.NO,A.单据,nvl(A.姓名,'') 姓名,D.性别,D.年龄,A.填制日期 From 未发药品记录 A,药品收发记录 B,门诊费用记录 D,Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)) C " & _
                                    "Where A.单据=B.单据 And A.NO=B.NO And A.库房id=B.库房id And B.费用id=D.id And  A.病人id=C.Column_Value And A.库房id=[2] and A.填制日期 between [3] and [4] and nvl(A.排队状态,0)=0 " & _
                                    " Union All " & _
                                    "Select /*+ Rule*/ distinct A.NO,A.单据,nvl(A.姓名,'') 姓名,D.性别,D.年龄,A.填制日期 From 未发药品记录 A,药品收发记录 B,住院费用记录 D,Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)) C " & _
                                    "Where A.单据=B.单据 And A.NO=B.NO And A.库房id=B.库房id And B.费用id=D.id And  A.病人id=C.Column_Value And A.库房id=[2] and A.填制日期 between [3] and [4] and nvl(A.排队状态,0)=0 "
                        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "", str病人id, mParams.lng药房ID, mSQLCondition.date开始日期, mSQLCondition.date结束日期)
                    ElseIf str姓名 <> "" And str病人id = "" Then
                        gstrSQL = "Select /*+ Rule*/ distinct A.NO,A.单据,nvl(A.姓名,'') 姓名,D.性别,D.年龄,A.填制日期 From 未发药品记录 A,药品收发记录 B,门诊费用记录 D,Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) C " & _
                                    "Where A.单据=B.单据 And A.NO=B.NO And A.库房id=B.库房id And B.费用id=D.id And  A.姓名=C.Column_Value And A.库房id=[2] and A.填制日期 between [3] and [4] and nvl(A.排队状态,0)=0 " & _
                                    " Union All " & _
                                    "Select /*+ Rule*/ distinct A.NO,A.单据,nvl(A.姓名,'') 姓名,D.性别,D.年龄,A.填制日期 From 未发药品记录 A,药品收发记录 B,住院费用记录 D,Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) C " & _
                                    "Where A.单据=B.单据 And A.NO=B.NO And A.库房id=B.库房id And B.费用id=D.id And  A.姓名=C.Column_Value And A.库房id=[2] and A.填制日期 between [3] and [4] and nvl(A.排队状态,0)=0 "

                        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "", str姓名, mParams.lng药房ID, mSQLCondition.date开始日期, mSQLCondition.date结束日期)
                    Else
                        gstrSQL = "Select /*+ Rule*/ distinct A.NO,A.单据,nvl(A.姓名,'') 姓名,D.性别,D.年龄,A.填制日期 From 未发药品记录 A,药品收发记录 B,门诊费用记录 D,Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)) C,Table(Cast(f_Str2list([5]) As zlTools.t_Strlist)) E " & _
                                    "Where A.单据=B.单据 And A.NO=B.NO And A.库房id=B.库房id And B.费用id=D.id And (A.病人id=C.Column_Value or A.姓名=E.Column_Value) And A.库房id=[2] and A.填制日期 between [3] and [4] and nvl(A.排队状态,0)=0 " & _
                                    " Union All " & _
                                    "Select /*+ Rule*/ distinct A.NO,A.单据,nvl(A.姓名,'') 姓名,D.性别,D.年龄,A.填制日期 From 未发药品记录 A,药品收发记录 B,住院费用记录 D,Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)) C,Table(Cast(f_Str2list([5]) As zlTools.t_Strlist)) E " & _
                                    "Where A.单据=B.单据 And A.NO=B.NO And A.库房id=B.库房id And B.费用id=D.id And (A.病人id=C.Column_Value or A.姓名=E.Column_Value) And A.库房id=[2] and A.填制日期 between [3] and [4] and nvl(A.排队状态,0)=0 "
                                    
                        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "", str病人id, mParams.lng药房ID, mSQLCondition.date开始日期, mSQLCondition.date结束日期, str姓名)
                    End If
                End If
                
                If rsData.RecordCount > 1 Then
                    frmSelectNo.ShowMe rsData, Me, txtPati.Text
                End If
                
                If Not rsData Is Nothing Then
                    If rsData.RecordCount > 0 Then
                        rsData.MoveFirst
                        gcnOracle.BeginTrans
                        Do While Not rsData.EOF
                            strCondition = IIf(strCondition = "", strCondition, strCondition & " OR ") & "NO='" & rsData!NO & "'"
                            gstrSQL = "Zl_未发药品记录_配药确认("
                                'NO
                                gstrSQL = gstrSQL & "'" & rsData!NO & "'"
                                '单据
                                gstrSQL = gstrSQL & "," & rsData!单据
                                '库房ID
                                gstrSQL = gstrSQL & "," & mParams.lng药房ID
                                '配药确认
                                gstrSQL = gstrSQL & "," & 1
                                '操作员
                                gstrSQL = gstrSQL & ",'" & gstrUserName & "')"
            
                                Call zldatabase.ExecuteProcedure(gstrSQL, "RecipeWork_DosageOk")
                            rsData.MoveNext
                        Loop
                        
                        gcnOracle.CommitTrans
                        Call RefreshList(mcondition.intListType)
                    End If
                End If
            End If
        End If
        
        DoEvents
        KeyAscii = 0
        mblnFinding = False
        
        If imgFilter.BorderStyle = cstLocate Then
            Call Form_KeyDown(vbKeyF3, 0)
        Else
            If strCard = "单据号" Then
                If IsNumeric(txtPati.Text) Then
                    txtPati.Text = UCase(GetFullNO(txtPati.Text, 13))
                End If
            End If
            
            DoEvents
            RefreshList mcondition.intListType
            
            '获取过滤出的所有处方
            strRecipeString = mfrmList.GetCurrentBatchRecipe
    
            arrRecipe = Split(strRecipeString, "|")
            intCount = UBound(arrRecipe)
            
            For n = 0 To intCount
                strNos = IIf(strNos = "", "", strNos & "|") & Val(Split(arrRecipe(n), ",")(0)) & "," & Split(arrRecipe(n), ",")(1)
            Next
            
            '如果自动发药有开始发药流程，则调用接口上传处方
            '过滤模式上传过滤出的所有处方
            '不兼容接口时没有这个功能
            If TypeName(mobjDrugMAC) = "clsDrugPacker" Then
                If mcondition.intListType = mListType.待发药 And mblnPackerConnect And mblnLoadDrug And mintAutoSendFlow = 1 And strNos <> "" Then
                    If mblnCompatible = True Then
                        If mobjDrugMAC.DYEY_MZ_TransRecipeList(mstrOpr, UserInfo.用户编码, UserInfo.用户姓名, mParams.lng药房ID, strNos, strReturn, mSendOper.StartSend) = False Then
                            If MsgBox("自动发药系统未准备好，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            ElseIf TypeName(mobjDrugMAC) = "clsDrugMachine" Then
                If mcondition.intListType = mListType.待发药 And mblnPackerConnect Then
                    mobjDrugMAC.Operation gstrDbUser, Val("22-开始发药"), "1|" & Replace(strNos, "|", ";"), strReturn
'                           If strReturn <> "" Then MsgBox strReturn, vbInformation, gstrSysName
                End If
            End If
                    
            If mblnScaner Then
                txtPati.Text = ""
                txtPati.SetFocus
            End If
        End If
        
        If mParams.blnSign And mParams.blnMustDosageOkProcess And bln刷卡 And InStr(1, mstrPrivs, "配药确认") > 0 Then
            For i = 0 To rsData.RecordCount - 1
                mfrmList.SetSign (strCondition)
                If tbcDetail.Selected.index = 0 Then
                    mfrmDetail.CmdProcess
                ElseIf tbcDetail.Item(1).Visible = True Then
                    mfrmRecipe.CmdProcess
                End If
            Next
        End If
        
        If bln刷卡 And blnSta = False Then
            Me.stbThis.Panels(2).Text = "卡号：" & txtPati.Text
            txtPati.Text = ""
            txtPati.SetFocus
        End If

        Call zlControl.TxtSelAll(txtPati)
        
        '待发药状态，扫描后自动呼叫
        If mParams.bln扫描后呼叫 And mcondition.intListType = mListType.待发药 And mblnScaner And mParams.blnStartCall And mblnFinding And (mblnBrushCard Or Not mbln允许两次刷卡) Then
            txtPati.Text = ""
            txtPati.SetFocus
            Call RecipeWork_Call
        End If

    End If
    Exit Sub
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Function InputIsScaner(ByRef txtInput As Object, ByVal KeyAscii As Integer) As Boolean
'功能：判断指定文本框中当前输入是否是由条码设备读入：暂时支持对“药品收发记录.NO”读入
'参数：KeyAscii=在KeyPress事件中调用的参数
    Static sngInputBegin As Single
    Dim sngNow As Single, blnScaner As Boolean, strText As String
    
    '处理当前键入后显示的内容(还未显示出来)
    strText = txtInput.Text
    If txtInput.SelLength = Len(txtInput.Text) Then strText = ""
    If KeyAscii = 8 Or KeyAscii = 9 Or KeyAscii = 10 Then
        If strText <> "" Then strText = Mid(strText, 1, Len(strText) - 1)
    Else
        strText = strText & Chr(KeyAscii)
    End If
    
    '判断是否由条码设备读入
    sngNow = Timer
    If txtInput.Text = "" Or strText = "" Then
        sngInputBegin = sngNow
    Else
        If Format((sngNow - sngInputBegin) / Len(strText), "0.000") < 0.04 Then blnScaner = True
    End If
    
    InputIsScaner = blnScaner
End Function

Private Sub txtPati_LostFocus()
    txtPati.BackColor = &H80000005
    
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (False)
    
    mblnInput = False
End Sub


Private Sub txtPati_Validate(Cancel As Boolean)
    If Val(mParams.int输入模式) = mFindType.单据号 Then
        If IsNumeric(txtPati.Text) Then
            txtPati.Text = GetFullNO(txtPati.Text, 13)
        End If
    End If
End Sub

Private Function Get实收金额(ByVal Int单据 As Integer, ByVal strNo As String, ByVal int门诊标志 As Integer) As Double
    Dim strsql As String
    Dim rs实收金额 As ADODB.Recordset
    
    If int门诊标志 = 1 Or int门诊标志 = 4 Then
        strsql = "Select Nvl(Sum(A.实收金额), 0) 实收金额 From 门诊费用记录 A Where A.记录状态 = 0 And A.Id In (Select Distinct B.费用id From 药品收发记录 B Where B.单据 = [1] And B.No = [2]) "
    Else
        strsql = "Select Nvl(Sum(A.实收金额), 0) 实收金额 From 住院费用记录 A Where A.记录状态 = 0 And A.Id In (Select Distinct B.费用id From 药品收发记录 B Where B.单据 = [1] And B.No = [2]) "
    End If
    
    On Error GoTo errRow
    Set rs实收金额 = zldatabase.OpenSQLRecord(strsql, "Get实收金额", Int单据, strNo, int门诊标志)
    Get实收金额 = rs实收金额!实收金额
    Exit Function
errRow:
    If ErrCenter = 1 Then Resume
End Function

Private Sub ChangWin()
    Dim date开始日期 As Date
    Dim date结束日期 As Date
    Dim dteTime As Date
    
    dteTime = Sys.Currentdate
    '时间范围
    Select Case cbo时间范围.ListIndex
        Case mTimeRange.当天
            date开始日期 = CDate(Format(dteTime, "yyyy-mm-dd") & " 00:00:00")
            date结束日期 = CDate(Format(dteTime, "yyyy-mm-dd") & " 23:59:59")
        Case mTimeRange.两天内
            date开始日期 = CDate(Format(DateAdd("d", -1, dteTime), "yyyy-mm-dd") & " 00:00:00")
            date结束日期 = CDate(Format(dteTime, "yyyy-mm-dd") & " 23:59:59")
        Case mTimeRange.三天内
            date开始日期 = CDate(Format(DateAdd("d", -2, dteTime), "yyyy-mm-dd") & " 00:00:00")
            date结束日期 = CDate(Format(dteTime, "yyyy-mm-dd") & " 23:59:59")
        Case mTimeRange.指定时间范围
            date开始日期 = CDate(Format(Dtp开始时间.Value, "yyyy-mm-dd hh:mm:ss"))
            date结束日期 = CDate(Format(Dtp结束时间.Value, "yyyy-mm-dd hh:mm:ss"))
        Case Else
            date开始日期 = CDate(Format(dteTime, "yyyy-mm-dd") & " 00:00:00")
            date结束日期 = CDate(Format(dteTime, "yyyy-mm-dd") & " 23:59:59")
    End Select
    
    Call GetSendWindows(mParams.lng药房ID)
    If mQueue.blnWin = False Then
        MsgBox "该药房所有窗口已下班，不能进行调整发药窗口操作！", vbInformation, gstrSysName
    Else
        Call frm调整发药窗口.ShowMe(mParams.lng药房ID, Me, date开始日期, date结束日期, mstrDeptNode)
    End If
End Sub

Private Sub InitIDKindNew()
    Dim int输入模式 As Integer
    Dim strTemp As String
    
    int输入模式 = mParams.int输入模式索引
    strTemp = "单|单据号|0;门|门诊号|0;姓|姓名|0;身|身份证|0;IC|IC卡号|1;医|医保号|0;住|住院号|0"
    Me.IDKNType.IDKindStr = strTemp
    Call IDKNType.zlInit(Me, glngSys, mlngMode, gcnOracle, gstrDbUser, mobjSquareCard, strTemp, txtPati)
'    IDKNType.SetAutoReadCard True
    Me.IDKNType.IDKind = int输入模式
End Sub

Private Sub IDKNType_ItemClick(index As Integer, objCard As zlIDKind.Card)
    Set mobjcard = objCard
    mParams.int输入模式索引 = index
    mParams.int输入模式 = Get输入模式(IDKNType.GetCurCard.名称)
    mintOld输入模式 = mParams.int输入模式
    
'    txtPati.MaxLength = objCard.卡号长度
    If objCard.卡号密文规则 <> "" Then
        txtPati.PasswordChar = "*"
    Else
        txtPati.PasswordChar = ""
    End If
    
    mbln允许两次刷卡 = False
    If mParams.str两次刷卡发药 <> "" Then
        mbln允许两次刷卡 = InStr(1, "," & mParams.str两次刷卡发药 & ",", "," & objCard.接口序号 & ",") > 0
    End If
    
    picConMain_Resize
End Sub

Private Function Get输入模式(ByVal str类型 As String) As Integer
    '从IDKind中返回当前程序内部所定义的类型
    Dim i As Integer
    Dim str类型串 As String
    
    'str类型串与传入的IDKindStr类型名称、顺序一致
    str类型串 = "单据号,门诊号,姓名,二代身份证,IC卡,医保号,住院号"
    
    For i = 0 To UBound(Split(str类型串, ","))
        If Split(str类型串, ",")(i) = str类型 Then
            Get输入模式 = i + 1
            Exit For
        End If
    Next
    
    '当IDKindf返回的类型不是IDKindStr传入的类型，则赋值一个大于IDKindStr类型个数的数字
    If Get输入模式 = 0 Then Get输入模式 = 8
     
End Function

Private Sub IDKNType_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    
    txtPati.Text = objPatiInfor.卡号
    If txtPati.Text <> "" Then Call txtPati_KeyPress(vbKeyReturn)
End Sub


Private Sub GetChildWin()
    Dim strsql As String
    Dim rstemp As Recordset
    
    On Error GoTo errHandle
    
    strsql = "select 名称 from 发药窗口 A,Table(f_Str2list([1])) B  where A.名称=B.Column_Value or A.叫号窗口=B.Column_Value "
    
    Set rstemp = zldatabase.OpenSQLRecord(strsql, "获取子窗口", mParams.Str窗口)
    
    mParams.Str窗口 = ""
    Do While Not rstemp.EOF
        mParams.Str窗口 = mParams.Str窗口 & rstemp!名称 & ","
        rstemp.MoveNext
    Loop
    
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub









