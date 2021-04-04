VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmOtherOutputCard 
   Caption         =   "药品其他出库单"
   ClientHeight    =   8295
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   14475
   Icon            =   "frmOtherOutputCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   14475
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdAllCls 
      Caption         =   "全清(&L)"
      Height          =   350
      Left            =   7560
      TabIndex        =   32
      Top             =   5700
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdAllSel 
      Caption         =   "全冲(&A)"
      Height          =   350
      Left            =   6240
      TabIndex        =   31
      Top             =   5700
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   3720
      TabIndex        =   13
      Top             =   5370
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "查找(&F)"
      Height          =   350
      Left            =   2040
      TabIndex        =   12
      Top             =   5280
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   11
      Top             =   5280
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6240
      TabIndex        =   9
      Top             =   5280
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7560
      TabIndex        =   10
      Top             =   5280
      Width           =   1100
   End
   Begin VB.PictureBox Pic单据 
      BackColor       =   &H80000004&
      Height          =   5175
      Left            =   30
      ScaleHeight     =   5115
      ScaleWidth      =   14295
      TabIndex        =   14
      Top             =   0
      Width           =   14355
      Begin VB.ComboBox cbo外调单位 
         Height          =   300
         Left            =   8010
         TabIndex        =   5
         Text            =   "cbo外调单位"
         Top             =   900
         Visible         =   0   'False
         Width           =   3435
      End
      Begin VB.ComboBox cbo外销单位 
         Height          =   300
         Left            =   8010
         TabIndex        =   36
         Text            =   "cbo外销单位"
         Top             =   900
         Visible         =   0   'False
         Width           =   3435
      End
      Begin VB.CheckBox chkIn 
         Caption         =   "导"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   780
         Style           =   1  'Graphical
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   "导入记帐单:F3"
         Top             =   90
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtIn 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1260
         MaxLength       =   8
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   105
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.ComboBox cboType 
         Height          =   300
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   900
         Width           =   1965
      End
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   2805
         Left            =   195
         TabIndex        =   6
         Top             =   1275
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   4948
         Appearance      =   0
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Active          =   -1  'True
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483634
      End
      Begin VB.TextBox txt摘要 
         Height          =   300
         Left            =   900
         MaxLength       =   40
         TabIndex        =   8
         Top             =   4380
         Width           =   10410
      End
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   510
         Width           =   1965
      End
      Begin VB.Label Txt修改日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7680
         TabIndex        =   41
         Top             =   4740
         Width           =   1875
      End
      Begin VB.Label Txt修改人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5640
         TabIndex        =   40
         Top             =   4740
         Width           =   915
      End
      Begin VB.Label lbl修改人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "修改人"
         Height          =   180
         Left            =   5040
         TabIndex        =   39
         Top             =   4800
         Width           =   540
      End
      Begin VB.Label lbl修改日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "修改日期"
         Height          =   180
         Left            =   6900
         TabIndex        =   38
         Top             =   4800
         Width           =   720
      End
      Begin VB.Label lblOther 
         AutoSize        =   -1  'True
         Caption         =   "外调(销)合计:"
         Height          =   180
         Left            =   6360
         TabIndex        =   37
         Top             =   4140
         Width           =   1170
      End
      Begin VB.Label lbl外调单位 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "外调单位(&D)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6960
         TabIndex        =   4
         Top             =   960
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lbl外销单位 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "外销单位(&D)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6960
         TabIndex        =   35
         Top             =   960
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lblDifference 
         AutoSize        =   -1  'True
         Caption         =   "差价合计:"
         Height          =   180
         Left            =   4920
         TabIndex        =   29
         Top             =   4140
         Width           =   810
      End
      Begin VB.Label lblSalePrice 
         AutoSize        =   -1  'True
         Caption         =   "售价金额合计:"
         Height          =   180
         Left            =   2040
         TabIndex        =   28
         Top             =   4140
         Width           =   1170
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "成本金额合计:"
         Height          =   180
         Left            =   240
         TabIndex        =   27
         Top             =   4140
         Width           =   1170
      End
      Begin VB.Label Txt审核人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   10350
         TabIndex        =   25
         Top             =   4740
         Width           =   915
      End
      Begin VB.Label Txt审核日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   12450
         TabIndex        =   24
         Top             =   4740
         Width           =   1875
      End
      Begin VB.Label Txt填制日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2940
         TabIndex        =   23
         Top             =   4740
         Width           =   1875
      End
      Begin VB.Label Txt填制人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   22
         Top             =   4740
         Width           =   915
      End
      Begin VB.Label txtNo 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   9960
         TabIndex        =   21
         Top             =   158
         Width           =   1425
      End
      Begin VB.Label LblNo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NO."
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   9480
         TabIndex        =   20
         Top             =   195
         Width           =   480
      End
      Begin VB.Label lbl摘要 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "摘要(&M)"
         Height          =   180
         Left            =   240
         TabIndex        =   7
         Top             =   4455
         Width           =   645
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "药品其他出库单"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   30
         TabIndex        =   19
         Top             =   120
         Width           =   11535
      End
      Begin VB.Label LblStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "库房(&S)"
         Height          =   180
         Left            =   540
         TabIndex        =   0
         Top             =   570
         Width           =   630
      End
      Begin VB.Label Lbl填制人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "填制人"
         Height          =   180
         Left            =   300
         TabIndex        =   18
         Top             =   4800
         Width           =   540
      End
      Begin VB.Label Lbl填制日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "填制日期"
         Height          =   180
         Left            =   2160
         TabIndex        =   17
         Top             =   4800
         Width           =   720
      End
      Begin VB.Label Lbl审核人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审核人"
         Height          =   180
         Left            =   9765
         TabIndex        =   16
         Top             =   4800
         Width           =   540
      End
      Begin VB.Label Lbl审核日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审核日期"
         Height          =   180
         Left            =   11640
         TabIndex        =   15
         Top             =   4800
         Width           =   720
      End
      Begin VB.Label lblType 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入出类别(&T)"
         Height          =   180
         Left            =   210
         TabIndex        =   2
         Top             =   960
         Width           =   990
      End
   End
   Begin MSComctlLib.ImageList imghot 
      Left            =   840
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":014A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":0364
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":057E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":0798
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":09B2
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":0BCC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":0DE6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":1000
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgcold 
      Left            =   120
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":121A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":1434
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":164E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":1868
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":1A82
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":1C9C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":1EB6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":20D0
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   30
      Top             =   7935
      Width           =   14475
      _ExtentX        =   25532
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmOtherOutputCard.frx":22EA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   19182
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmOtherOutputCard.frx":2B7E
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmOtherOutputCard.frx":3080
            Key             =   "WB"
            Object.ToolTipText     =   "五笔(F7)"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin VB.Label lblCode 
      Caption         =   "编码"
      Height          =   255
      Left            =   3240
      TabIndex        =   26
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Menu mnuCol 
      Caption         =   "列名"
      Visible         =   0   'False
      Begin VB.Menu mnuColDrug 
         Caption         =   "药名(编码和名称)"
         Index           =   0
      End
      Begin VB.Menu mnuColDrug 
         Caption         =   "药名(仅编码)"
         Index           =   1
      End
      Begin VB.Menu mnuColDrug 
         Caption         =   "药名(仅名称)"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmOtherOutputCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mint编辑状态 As Integer             '1.新增；2、修改；3、验收；4、查看；5
Private mstr单据号 As String                '具体的单据号;
Private mint记录状态 As Integer             '1:正常记录;2-冲销记录;3-已经冲销的原记录
Private mblnSuccess As Boolean              '只要有一张成功，即为True，否则为False
Private mblnSave As Boolean                 '是否存盘和审核   TURE：成功。
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEnterCell As Boolean            '是否允许激法ENTERCELL()事件
Private mblnEdit As Boolean                 '是否可以修改
Private mblnChange As Boolean               '是否进行过编辑
Private mintParallelRecord As Integer       '对于新增后单据并行执行的处理： 1、代表正常情况；2、已经删除的记录；3、已经审核的记录
Private mbln下可用数量 As Boolean           '填单是否下可用数量
Private mblnViewCost As Boolean             '查看成本价 true-可以查看成本价 false-不可以查看成本价

Private mint库存检查 As Integer             '表示药品出库时是否进行库存检查：0-不检查;1-检查，不足提醒；2-检查，不足禁止
Private mcolUsedCount As Collection         '已使用的数量集合
Private mstrPrivs As String                 '权限
Private recSort As ADODB.Recordset          '按药品ID排序的专用记录集

Private Const mlng紫色 As Long = &HC000C0

Private mblnUpdate As Boolean               '表示是否已根据最新价格更新单据内容

Private mlng出库库房 As Long
Private mintUnit As Integer             '单位系数：1-售价;2-门诊;3-住院;4-药库

Private mintDrugNameShow As Integer         '药品显示：0－显示编码和名称；1－仅显示编码；2－仅显示名称

Private Const MStrCaption As String = "药品其他出库管理"

Dim mstrLike As String

Private mblnLoad As Boolean              '记录是否执行完成Form_Load事件

'从参数表中取药品价格、数量、金额小数位数（计算精度）
Private mintCostDigit As Integer        '成本价小数位数
Private mintPriceDigit As Integer       '售价小数位数
Private mintNumberDigit As Integer      '数量小数位数
Private mintMoneyDigit As Integer       '金额小数位数

Private Const mconint售价单位 As Integer = 1
Private Const mconint门诊单位 As Integer = 2
Private Const mconint住院单位 As Integer = 3
Private Const mconint药库单位 As Integer = 4

Private mstrTime_Start As String                      '进入单据编辑界面时，待编辑单据的最大修改时间
Private mstrTime_End As String                        '此刻该编辑单据的最大修改时间

'=========================================================================================

Private Const mconIntCol行号 As Integer = 1
Private Const mconIntCol药名 As Integer = 2
Private Const mconIntCol商品名 As Integer = 3
Private Const mconIntCol来源 As Integer = 4
Private Const mconIntCol基本药物 As Integer = 5
Private Const mconIntCol序号 As Integer = 6
Private Const mconIntCol规格 As Integer = 7
Private Const mconIntCol可用数量 As Integer = 8
Private Const mconIntcol加成率 As Integer = 9
Private Const mconIntCol实际金额 As Integer = 10
Private Const mconIntCol实际差价 As Integer = 11
Private Const mconIntCol比例系数 As Integer = 12
Private Const mconIntCol批次 As Integer = 13
Private Const mconIntCol产地 As Integer = 14
Private Const mconIntCol原产地 As Integer = 15
Private Const mconIntCol单位 As Integer = 16
Private Const mconIntCol批号 As Integer = 17
Private Const mconIntCol效期 As Integer = 18
Private Const mconIntCol批准文号 As Integer = 19
Private Const mconIntCol数量 As Integer = 20
Private Const mconIntCol冲销数量 As Integer = 21
Private Const mconIntCol采购价 As Integer = 22
Private Const mconIntCol采购金额 As Integer = 23
Private Const mconIntCol售价 As Integer = 24
Private Const mconIntCol售价金额 As Integer = 25
Private Const mconIntCol外调价 As Integer = 26
Private Const mconIntCol外调金额 As Integer = 27
Private Const mconIntCol增值税率 As Integer = 28
Private Const mconIntCol税金 As Integer = 29
Private Const mconintCol差价 As Integer = 30
Private Const mconIntCol药品编码和名称 = 31
Private Const mconIntCol药品编码 = 32
Private Const mconIntCol药品名称 = 33
Private Const mconintCol原始数量 As Integer = 34
Private Const mconIntColS  As Integer = 35            '总列数
'=========================================================================================

Private Sub SetDrugName(ByVal intType As Integer)
    '药品名称显示：
    'intType：0－显示编码和名称；1－仅显示编码；2－仅显示名称
    Dim lngRow As Long
    
    If intType = mintDrugNameShow Then Exit Sub
    
    mintDrugNameShow = intType
    
    With mshBill
        For lngRow = 1 To .rows - 1
            If .TextMatrix(lngRow, mconIntCol药名) <> "" Then
                If mintDrugNameShow = 1 Then
                    .TextMatrix(lngRow, mconIntCol药名) = .TextMatrix(lngRow, mconIntCol药品编码)
                ElseIf mintDrugNameShow = 2 Then
                    .TextMatrix(lngRow, mconIntCol药名) = .TextMatrix(lngRow, mconIntCol药品名称)
                Else
                    .TextMatrix(lngRow, mconIntCol药名) = .TextMatrix(lngRow, mconIntCol药品编码和名称)
                End If
            End If
        Next
    End With
End Sub

Private Sub SetSortRecord()
    Dim n As Integer
    
    If mshBill.rows < 2 Then Exit Sub
    If mshBill.TextMatrix(1, 0) = "" Then Exit Sub
    
    Set recSort = New ADODB.Recordset
    With recSort
        If .State = 1 Then .Close
        .Fields.Append "行号", adDouble, 18, adFldIsNullable
        .Fields.Append "序号", adDouble, 18, adFldIsNullable
        .Fields.Append "药品ID", adDouble, 18, adFldIsNullable
        .Fields.Append "批次", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        For n = 1 To mshBill.rows - 1
            If mshBill.TextMatrix(n, 0) <> "" Then
                .AddNew
                !行号 = n
                !序号 = IIf(Val(mshBill.TextMatrix(n, mconIntCol序号)) = 0, n, Val(mshBill.TextMatrix(n, mconIntCol序号)))
                !药品ID = Val(mshBill.TextMatrix(n, 0))
                !批次 = Val(mshBill.TextMatrix(n, mconIntCol批次))
                
                .Update
            End If
        Next
        
    End With
End Sub
Private Sub GetSysParm()
    mbln下可用数量 = (gtype_UserSysParms.P96_药品填单下可用库存 = 1)
End Sub

'检查数据依赖性
Private Function GetDepend() As Boolean
    Dim rsDepend As New Recordset
    
    On Error GoTo errHandle
    GetDepend = False
    gstrSQL = "SELECT B.Id " _
        & " FROM 药品单据性质 A, 药品入出类别 B " _
        & "Where A.类别id = B.ID " _
      & "AND A.单据 = 11 "
    Call SQLTest(App.Title, "药品其他出库单", gstrSQL)
    Set rsDepend = zlDataBase.OpenSQLRecord(gstrSQL, "GetDepend")
    Call SQLTest
    If rsDepend.EOF Then
        MsgBox "没有设置药品其他出库的入出类别，请检查药品入出分类！", vbInformation + vbOKOnly, gstrSysName
        rsDepend.Close
        Exit Function
    End If
    rsDepend.Close
    
    
    GetDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Sub ShowCard(FrmMain As Form, ByVal str单据号 As String, ByVal int编辑状态 As Integer, Optional int记录状态 As Integer = 1, Optional BlnSuccess As Boolean = False)
    mblnSave = False
    mblnSuccess = False
    mstr单据号 = str单据号
    mint编辑状态 = int编辑状态
    mint记录状态 = int记录状态
    mblnSuccess = BlnSuccess
    mblnChange = False
    mintParallelRecord = 1
    mstrPrivs = GetPrivFunc(glngSys, 1306)
    
    Set mfrmMain = FrmMain
    If Not GetDepend Then Exit Sub
    
    If mint编辑状态 = 1 Then
        mblnEdit = True
    ElseIf mint编辑状态 = 2 Then
        mblnEdit = True
    ElseIf mint编辑状态 = 3 Then
        mblnEdit = False
        CmdSave.Caption = "审核(&V)"
    ElseIf mint编辑状态 = 4 Then
        mblnEdit = False
        CmdSave.Caption = "打印(&P)"
        If Not zlStr.IsHavePrivs(mstrPrivs, "单据打印") Then
            CmdSave.Visible = False
        Else
            CmdSave.Visible = True
        End If
    ElseIf mint编辑状态 = 6 Then
        mblnEdit = False
        CmdSave.Caption = "冲销(&O)"
        cmdAllSel.Visible = True
        cmdAllCls.Visible = True
    End If
    
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    Me.Show vbModal, FrmMain
    BlnSuccess = mblnSuccess
    str单据号 = mstr单据号
    
End Sub


Private Sub cboStock_Change()
    mblnChange = True
End Sub


Private Sub cboStock_Click()
    Dim rsDetail As New Recordset
    Dim bln中药库房 As Boolean
    Dim str库房性质 As String
    
    On Error GoTo errHandle
    
    str库房性质 = ""
    mint库存检查 = MediWork_GetCheckStockRule(cboStock.ItemData(cboStock.ListIndex))
    If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
        gstrSQL = "Select 工作性质 From 部门性质说明 Where 部门id =[1]"
        Set rsDetail = zlDataBase.OpenSQLRecord(gstrSQL, "判断库房性质", cboStock.ItemData(cboStock.ListIndex))
        Do While Not rsDetail.EOF
            str库房性质 = str库房性质 & "," & rsDetail!工作性质
            rsDetail.MoveNext
        Loop
        If str库房性质 Like "*中药*" Or str库房性质 Like "*制剂室*" Then bln中药库房 = True
        mshBill.ColWidth(mconIntCol原产地) = IIf(bln中药库房, 800, 0)
        
        If mblnLoad = True Then Call SetSelectorRS(2, MStrCaption, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex))
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboStock_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboStock_Validate False
        OS.PressKey (vbKeyTab)
    End If
    
End Sub


Private Sub cboStock_Validate(Cancel As Boolean)
    Dim i As Integer
        
    With cboStock
        If .ListIndex <> mintcboIndex Then
            For i = 1 To mshBill.rows - 1
                If mshBill.TextMatrix(i, 0) <> "" Then
                    Exit For
                End If
            Next
            If i <> mshBill.rows Then
                If MsgBox("如果改变库房，有可能要改变相应药品的单位，且要清除现有单据内容，你是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    '处理药品单位改变
                    mintcboIndex = .ListIndex
                    mshBill.ClearBill
                            
                    mlng出库库房 = Me.cboStock.ItemData(Me.cboStock.ListIndex)
                    Call GetDrugDigit(mlng出库库房, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
                Else
                    .ListIndex = mintcboIndex
                End If
            Else
                mintcboIndex = .ListIndex
            End If
        End If
        
    End With
End Sub

Private Sub cboType_click()
    Dim i As Integer
    Dim j As Integer
    Dim intresult As Integer
    
    On Error Resume Next
    
    Me.lbl外调单位.Visible = False
    Me.cbo外调单位.Visible = False
    Me.lbl外销单位.Visible = False
    Me.cbo外销单位.Visible = False
    
    If cboType.Text = "药品外调" Then
        Me.lbl外调单位.Visible = True
        Me.cbo外调单位.Visible = True
        
        mshBill.TextMatrix(0, mconIntCol外调价) = "外调价"
        mshBill.TextMatrix(0, mconIntCol外调金额) = "外调金额"
        
        mshBill.ColWidth(mconIntCol外调价) = 1000
        mshBill.ColWidth(mconIntCol外调金额) = 1000
        cbo外调单位.Enabled = (mint编辑状态 = 1 Or mint编辑状态 = 2)
        mshBill.ColData(mconIntCol外调价) = IIf(cbo外调单位.Enabled, 4, 5)
        
        mshBill.ColWidth(mconIntCol增值税率) = 0
        mshBill.ColWidth(mconIntCol税金) = 0
    ElseIf cboType.Text = "药品外销" Then
        If mshBill.TextMatrix(1, 0) <> "" Then
            intresult = MsgBox("将清空列表数据，是否继续！", vbYesNo, gstrSysName)
            If intresult = vbYes Then
                Me.lbl外销单位.Visible = True
                Me.cbo外销单位.Visible = True
                
                mshBill.TextMatrix(0, mconIntCol外调价) = "外销价"
                mshBill.TextMatrix(0, mconIntCol外调金额) = "外销金额"
                mshBill.ColWidth(mconIntCol外调价) = 1000
                mshBill.ColWidth(mconIntCol外调金额) = 1000
                cbo外销单位.Enabled = (mint编辑状态 = 1 Or mint编辑状态 = 2)
                mshBill.ColData(mconIntCol外调价) = IIf(cbo外销单位.Enabled, 4, 5)
                
                mshBill.ColWidth(mconIntCol增值税率) = 1000
                mshBill.ColWidth(mconIntCol税金) = 1000
                
                For i = 1 To mshBill.rows - 1
                  For j = 0 To mshBill.Cols - 1
                      mshBill.TextMatrix(i, j) = ""
                  Next
                Next
                mshBill.rows = 2
                mshBill.SetFocus
            Else
                cboType.Text = "药品外调"
            End If
        Else
            Me.lbl外销单位.Visible = True
            Me.cbo外销单位.Visible = True
            
            mshBill.TextMatrix(0, mconIntCol外调价) = "外销价"
            mshBill.TextMatrix(0, mconIntCol外调金额) = "外销金额"
            mshBill.ColWidth(mconIntCol外调价) = 1000
            mshBill.ColWidth(mconIntCol外调金额) = 1000
            cbo外销单位.Enabled = (mint编辑状态 = 1 Or mint编辑状态 = 2)
            mshBill.ColData(mconIntCol外调价) = IIf(cbo外销单位.Enabled, 4, 5)
            
            mshBill.ColWidth(mconIntCol增值税率) = 1000
            mshBill.ColWidth(mconIntCol税金) = 1000
        End If
    Else
        mshBill.ColWidth(mconIntCol外调价) = 0
        mshBill.ColWidth(mconIntCol外调金额) = 0
        mshBill.ColData(mconIntCol外调价) = 5
        
        mshBill.ColWidth(mconIntCol增值税率) = 0
        mshBill.ColWidth(mconIntCol税金) = 0
    End If
End Sub

Private Sub cboType_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Function SeekCboIndex(objCbo As Object, varData As Variant) As Long
'功能：由ItemData或Text查找ComboBox的索引值
    Dim strType As String, i As Integer
    
    SeekCboIndex = -1
    
    strType = TypeName(varData)
    If strType = "Field" Then
        If Rec.IsType(varData.Type, adVarChar) Then strType = "String"
    End If
    
    If strType = "String" Then
        If varData <> "" Then
            '先精确查找
            For i = 0 To objCbo.ListCount - 1
                If objCbo.List(i) = varData Then
                    SeekCboIndex = i: Exit Function
                ElseIf NeedName(objCbo.List(i)) = varData And varData <> "" Then
                    SeekCboIndex = i: Exit Function
                End If
            Next
            '再模糊查找
            For i = 0 To objCbo.ListCount - 1
                If InStr(objCbo.List(i), varData) > 0 And varData <> "" Then
                    SeekCboIndex = i: Exit Function
                End If
            Next
        End If
    Else
        If varData <> 0 Then
            For i = 0 To objCbo.ListCount - 1
                If objCbo.ItemData(i) = varData Then
                    SeekCboIndex = i: Exit Function
                End If
            Next
        End If
    End If
End Function

Public Function NeedName(strList As String) As String
    If InStr(strList, Chr(13)) > 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, Chr(13)) + 1))
    ElseIf InStr(strList, "]") > 0 And InStr(strList, "-") = 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, "]") + 1))
    ElseIf InStr(strList, ")") > 0 And InStr(strList, "-") = 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, ")") + 1))
    Else
        NeedName = LTrim(Mid(strList, InStr(strList, "-") + 1))
    End If
End Function

Private Sub cbo外调单位_GotFocus()
    If cbo外调单位.Style = 0 Then
        Call zlControl.TxtSelAll(cbo外调单位)
    End If
End Sub

Private Sub cbo外调单位_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode <> vbKeyReturn Then Exit Sub
'    Call zlCommFun.PressKey(vbKeyTab)
    If KeyCode = vbKeyDelete Then
        If cbo外调单位.Style = 2 And cbo外调单位.ListIndex <> -1 Then
            cbo外调单位.ListIndex = -1
        End If
    End If
End Sub

Private Sub cbo外调单位_KeyPress(KeyAscii As Integer)
'    Dim IntMatchIdx As Integer
'
'    With cbo外调单位
'        IntMatchIdx = MatchIndex(.hWnd, KeyAscii, 1)
'        If IntMatchIdx = -2 Then Exit Sub
'        .ListIndex = IntMatchIdx
'        If .ListIndex = -1 Then .ListIndex = 0
'    End With

    Dim lngIdx As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf KeyAscii >= 32 Then
        If Not cbo外调单位.Locked And cbo外调单位.Style = 2 Then
            lngIdx = Cbo.MatchIndex(cbo外调单位.hWnd, KeyAscii)
            
            If lngIdx = -1 And cbo外调单位.ListCount > 0 Then lngIdx = 0
            cbo外调单位.ListIndex = lngIdx
        End If
    End If
End Sub

Private Sub cbo外调单位_Validate(Cancel As Boolean)
    '功能：根据输入的内容,自动匹配执行科室
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long, i As Long
    Dim strInput As String
    Dim vRect As RECT, blnCancel As Boolean
        
    If cbo外调单位.ListIndex <> -1 Then Exit Sub '已选中
    If cbo外调单位.Text = "" Then cbo外调单位.Tag = "": Exit Sub '无输入
    
    strInput = UCase(NeedName(cbo外调单位.Text))
    strSQL = "Select Rownum As id,编码,简码,名称 From 药品外调单位 Where Upper(编码) Like [1] Or Upper(名称) Like [2] Or Upper(简码) Like [2] Order By 编码"
        
    On Error GoTo errH
    vRect = zlControl.GetControlRect(cbo外调单位.hWnd)
    Set rsTmp = zlDataBase.ShowSQLSelect(Me, strSQL, 0, "外调单位", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, cbo外调单位.Height, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%")
    If Not rsTmp Is Nothing Then
        intIdx = SeekCboIndex(cbo外调单位, nvl(rsTmp!简码) & "-" & Chr(13) & rsTmp!名称)
        If intIdx <> -1 Then
            cbo外调单位.ListIndex = intIdx
        Else
            cbo外调单位.AddItem nvl(rsTmp!编码) & "-" & Chr(13) & rsTmp!名称, cbo外调单位.ListCount - 1
            cbo外调单位.ListIndex = cbo外调单位.NewIndex
        End If
    Else
        If Not blnCancel Then
            MsgBox "未找到对应的外调单位。", vbInformation, gstrSysName
        End If
        Cancel = True: Exit Sub
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub cbo外销单位_GotFocus()
    If cbo外销单位.Style = 0 Then
        Call zlControl.TxtSelAll(cbo外销单位)
    End If
End Sub

Private Sub cbo外销单位_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode <> vbKeyReturn Then Exit Sub
'    Call zlCommFun.PressKey(vbKeyTab)
    If KeyCode = vbKeyDelete Then
        If cbo外销单位.Style = 2 And cbo外销单位.ListIndex <> -1 Then
            cbo外销单位.ListIndex = -1
        End If
    End If
End Sub


Private Sub cbo外销单位_KeyPress(KeyAscii As Integer)
'    Dim IntMatchIdx As Integer
'
'    With cbo外销单位
'        IntMatchIdx = MatchIndex(.hWnd, KeyAscii, 1)
'        If IntMatchIdx = -2 Then Exit Sub
'        .ListIndex = IntMatchIdx
'        If .ListIndex = -1 Then .ListIndex = 0
'    End With
    Dim lngIdx As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf KeyAscii >= 32 Then
        If Not cbo外销单位.Locked And cbo外销单位.Style = 2 Then
            lngIdx = Cbo.MatchIndex(cbo外销单位.hWnd, KeyAscii)
            If lngIdx = -1 And cbo外销单位.ListCount > 0 Then lngIdx = 0
            cbo外销单位.ListIndex = lngIdx
        End If
    End If
End Sub


Private Sub cbo外销单位_Validate(Cancel As Boolean)
    '功能：根据输入的内容,自动匹配执行科室
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long, i As Long
    Dim strInput As String
    Dim vRect As RECT, blnCancel As Boolean
        
    If cbo外销单位.ListIndex <> -1 Then Exit Sub '已选中
    If cbo外销单位.Text = "" Then cbo外销单位.Tag = "": Exit Sub '无输入
    
    strInput = UCase(NeedName(cbo外销单位.Text))
    strSQL = "Select Rownum As id,编码,简码,名称 From 药品外销单位 Where Upper(编码) Like [1] Or Upper(名称) Like [2] Or Upper(简码) Like [2] Order By 编码"
        
    On Error GoTo errH
    vRect = zlControl.GetControlRect(cbo外销单位.hWnd)
    Set rsTmp = zlDataBase.ShowSQLSelect(Me, strSQL, 0, "外销单位", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, cbo外销单位.Height, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%")
    If Not rsTmp Is Nothing Then
        intIdx = SeekCboIndex(cbo外销单位, nvl(rsTmp!简码) & "-" & Chr(13) & rsTmp!名称)
        If intIdx <> -1 Then
            cbo外销单位.ListIndex = intIdx
        Else
            cbo外销单位.AddItem nvl(rsTmp!编码) & "-" & Chr(13) & rsTmp!名称, cbo外销单位.ListCount - 1
            cbo外销单位.ListIndex = cbo外销单位.NewIndex
        End If
    Else
        If Not blnCancel Then
            MsgBox "未找到对应的外销单位。", vbInformation, gstrSysName
        End If
        Cancel = True: Exit Sub
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub chkIn_Click()
    txtIn.Enabled = chkIn.Value
    If chkIn.Value Then
        txtIn.SetFocus
    Else
        txtIn.Text = ""
    End If
End Sub

Private Sub cmdAllCls_Click()
    Dim intRow As Integer
    
    With mshBill
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                .TextMatrix(intRow, mconIntCol冲销数量) = zlStr.FormatEx(0, mintNumberDigit, , True)
                .TextMatrix(intRow, mconIntCol采购金额) = zlStr.FormatEx(0, mintMoneyDigit, , True)
                .TextMatrix(intRow, mconIntCol售价金额) = zlStr.FormatEx(0, mintMoneyDigit, , True)
                .TextMatrix(intRow, mconintCol差价) = zlStr.FormatEx(0, mintMoneyDigit, , True)
            End If
        Next
    End With
    Call 显示合计金额
    mblnChange = False
End Sub

Private Sub cmdAllSel_Click()
    Dim intRow As Integer
    
    With mshBill
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                .TextMatrix(intRow, mconIntCol冲销数量) = .TextMatrix(intRow, mconIntCol数量)
                .TextMatrix(intRow, mconIntCol采购金额) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol数量) * .TextMatrix(intRow, mconIntCol采购价), mintMoneyDigit, , True)
                .TextMatrix(intRow, mconIntCol售价金额) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol数量) * .TextMatrix(intRow, mconIntCol售价), mintMoneyDigit, , True)
                .TextMatrix(intRow, mconintCol差价) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol售价金额) - .TextMatrix(intRow, mconIntCol采购金额), mintMoneyDigit, , True)
            End If
        Next
    End With
    Call 显示合计金额
    
    mblnChange = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'查找
Private Sub cmdFind_Click()
    
    If lblCode.Visible = False Then
        lblCode.Visible = True
        txtCode.Visible = True
        txtCode.SetFocus
    Else
        FindRow mshBill, mconIntCol药品编码和名称, txtCode.Text, True
        lblCode.Visible = False
        txtCode.Visible = False
    End If
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub Form_Activate()
'    mblnChange = False
    Select Case mintParallelRecord
        Case 1
            '正常
        Case 2
            If mint编辑状态 = 6 Then
                MsgBox "该单据已没有可以冲销的药品，请检查！", vbOKOnly, gstrSysName
            Else
                '单据已被删除
                MsgBox "该单据已被删除，请检查！", vbOKOnly, gstrSysName
            End If
            Unload Me
            Exit Sub
        Case 3
            '修改的单据已被审核
            MsgBox "该单据已被其他人审核，请检查！", vbOKOnly, gstrSysName
            Unload Me
            Exit Sub
    End Select
    
    '初始化简码方式
    If (mint编辑状态 = 1 Or mint编辑状态 = 2) And gbytSimpleCodeTrans = 1 Then
        staThis.Panels("PY").Visible = True
        staThis.Panels("WB").Visible = True
        gint简码方式 = Val(zlDataBase.GetPara("简码方式", , , 0))    '默认拼音简码
        Logogram staThis, gint简码方式
    Else
        staThis.Panels("PY").Visible = False
        staThis.Panels("WB").Visible = False
    End If
    
    'Call cboType_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 70 Or KeyCode = 102 Then
        If Shift = vbCtrlMask Then   'Ctrl+F
            cmdFind_Click
        End If
    ElseIf KeyCode = vbKeyF3 Then
        FindRow mshBill, mconIntCol药名, txtCode.Text, False
    ElseIf KeyCode = vbKeyF7 Then
        If staThis.Panels("PY").Bevel = sbrRaised Then
            Logogram staThis, 0
        Else
            Logogram staThis, 1
        End If
    End If
End Sub

Private Sub CmdSave_Click()
    Dim lng药品ID As Long
    Dim intRow As Integer
    Dim bln库房 As Boolean
    Dim bln分批 As Boolean
    Dim rsTemp As ADODB.Recordset
    Dim BlnSuccess As Boolean
    Dim blnTrans As Boolean
    Dim intLop As Integer
    Dim lng上次药品ID As Long
    
    On Error GoTo ErrHand
    
    '设置排序数据集
    Call SetSortRecord
        
    If mint编辑状态 = 4 Then    '查看
        '打印
        printbill
        '退出
        Unload Me
        Exit Sub
    End If
    
    '检查界面上药品进行预调价处理
    For intLop = 1 To Me.mshBill.rows - 1
        If mshBill.TextMatrix(intLop, 0) <> "" Then '有药品
            Call AutoAdjustPrice_ByID(Val(mshBill.TextMatrix(intLop, 0)))
        End If
    Next
    
    If mint编辑状态 = 3 Then        '审核
        mstrTime_End = GetBillInfo(11, mstr单据号)
        If mstrTime_End = "" Then
            MsgBox "该单据已经被其他操作员删除！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If mstrTime_End > mstrTime_Start Then
            MsgBox "该单据已经被其他操作员编辑，请退出后重试！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If Not 检查单价(11, txtNo, False) And Not mblnUpdate Then
            '以最新的价格更新单据体，退出的目的是让用户看一下最终的单据
            MsgBox "有记录未使用最新价格，程序将自动完成更新（售价、成本价、售价金额、成本金额、差价），更新后请检查！", vbInformation, gstrSysName
            Call RefreshBill
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
    
        If Not 药品单据审核(Txt填制人.Caption) Then Exit Sub

        '零差价管理：检查是否存在不满足零差价的药品
        For intLop = 1 To mshBill.rows - 1
            If mshBill.TextMatrix(intLop, 0) <> "" And gtype_UserSysParms.P275_零差价管理模式 = 2 Then
                If IsPriceAdjustMod(Val(mshBill.TextMatrix(intLop, 0))) = True Then
                    If CheckPriceAdjust(Val(mshBill.TextMatrix(intLop, 0)), cboStock.ItemData(cboStock.ListIndex), Val(mshBill.TextMatrix(intLop, mconIntCol批次))) = False Then
                        MsgBox "第" & intLop & "行药品已启用零差价管理，但库存记录中售价和成本价不一致，不能进行业务，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        mshBill.Row = intLop
                        mshBill.MsfObj.TopRow = intLop
                        Exit Sub
                    End If
                End If
            End If
        Next
        
        blnTrans = True
        gcnOracle.BeginTrans
        '如果审核时修改了单据，则重新生成单据保存
        If mblnChange Then
            If Not SaveCard(True) Then
                gcnOracle.RollbackTrans: Exit Sub
            End If
        End If
        
        If Not SaveCheck Then
            gcnOracle.RollbackTrans: Exit Sub
        End If
                
        gcnOracle.CommitTrans
        
        If Val(zlDataBase.GetPara("审核打印", glngSys, 模块号.其他出库)) = 1 Then
            '打印
            If zlStr.IsHavePrivs(mstrPrivs, "单据打印") Then
                printbill
                
                If Val(zlDataBase.GetPara("打印药品条码", glngSys, 模块号.其他出库)) = 1 And zlStr.IsHavePrivs(mstrPrivs, "药品条码打印") Then
                    '按药品ID顺序更新数据
                    recSort.Sort = "药品id"
                    recSort.MoveFirst
                    '打印药品条码
                    Do While Not recSort.EOF
                        If lng上次药品ID <> Val(recSort!药品ID) Then
                            ReportOpen gcnOracle, glngSys, "ZL1_INSIDE_1306_1", Me, "药品=" & Val(recSort!药品ID), 2
                            lng上次药品ID = recSort!药品ID
                        End If
                        recSort.MoveNext
                    Loop
                End If

            End If
        End If

        Unload Me
        Exit Sub
    End If
    
    If mint编辑状态 = 6 Then '冲销
        If mblnChange = False Then
            MsgBox "请录入冲销数量！", vbInformation + vbOKOnly, gstrSysName
            Exit Sub
        End If
        
        If MsgBox("你确实要冲销单据吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            If SaveStrike = True Then
                Unload Me
            End If
        End If
        Exit Sub
    End If
    
    If mint编辑状态 = 2 Then
        If Not 检查单价(11, txtNo, False) And Not mblnUpdate Then
            '以最新的价格更新单据体，退出的目的是让用户看一下最终的单据
            MsgBox "有记录未使用最新价格，程序将自动完成更新（售价、成本价、售价金额、成本金额、差价），更新后请检查！", vbInformation, gstrSysName
            Call RefreshBill
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
    End If
    
    If mint编辑状态 = 1 Then '新增保存时，判断价格是否已经更新
        If 检查价格 Then
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
    End If
    
    If ValidData = False Then Exit Sub
    
    BlnSuccess = SaveCard
        
    If BlnSuccess = True Then
            
        If Val(zlDataBase.GetPara("存盘打印", glngSys, 模块号.其他出库)) = 1 Then
            '打印
            If zlStr.IsHavePrivs(mstrPrivs, "单据打印") Then
                printbill
                
                If Val(zlDataBase.GetPara("打印药品条码", glngSys, 模块号.其他出库)) = 1 And zlStr.IsHavePrivs(mstrPrivs, "药品条码打印") Then
                    '按药品ID顺序更新数据
                    recSort.Sort = "药品id"
                    recSort.MoveFirst
                    '打印药品条码
                    Do While Not recSort.EOF
                        If lng上次药品ID <> Val(recSort!药品ID) Then
                            ReportOpen gcnOracle, glngSys, "ZL1_INSIDE_1306_1", Me, "药品=" & Val(recSort!药品ID), 2
                            lng上次药品ID = recSort!药品ID
                        End If
                        recSort.MoveNext
                    Loop
                End If

            End If
        End If
        If mint编辑状态 = 2 Then   '修改
            Unload Me
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    
    mblnSave = False
    mblnEdit = True
    mshBill.ClearBill
    Call RefreshRowNO(mshBill, mconIntCol行号, 1)
    
    txt摘要.Text = ""
    cboType.SetFocus
    mblnChange = False
    If txtNo.Tag <> "" Then Me.staThis.Panels(2).Text = "上一张单据的NO号：" & txtNo.Tag
    
    Exit Sub
ErrHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
End Sub

Private Sub RefreshBill()
    '以最新价格最新单据相关数据，用于单据审核时
    Dim lngRow As Long, lngRows As Long, lng药品ID As Long
    Dim Dbl数量 As Double, dbl成本价 As Double, dbl成本金额 As Double, dbl零售价 As Double, dbl零售金额 As Double, dbl差价 As Double
    Dim rsPrice As New ADODB.Recordset
    Dim rsStock As ADODB.Recordset
    Dim blnAdj As Boolean
    Dim intCostDigit As Integer
    Dim intPriceDigit As Integer
        
    On Error GoTo errHandle
    intPriceDigit = GetDigit(0, 1, 2, 1)
    intCostDigit = GetDigit(0, 1, 1, 1)
        
    gstrSQL = " Select '售价' As 类型, a.序号, a.药品id ,nvl(a.批次,0) as 批次, 0 原价, b.现价" & _
            " From 药品收发记录 A," & _
                 " (Select 收费细目id, Nvl(现价, 0) 现价, 执行日期" & _
                   " From 收费价目" & _
                   " Where (终止日期 Is Null Or Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'yyyy-MM-dd')))" & _
                   GetPriceClassString("") & ") B, 收费项目目录 C" & _
            " Where a.单据 = 11 And a.No = [1] And a.药品id = b.收费细目id And c.Id = b.收费细目id And Round(a.零售价," & intPriceDigit & ") <> Round(b.现价, " & intPriceDigit & ") And" & _
              "    NVL(c.是否变价, 0) = 0 " & _
            " Union All" & _
            " Select '售价' As 类型, a.序号, a.药品id ,nvl(a.批次,0) as 批次, 0 原价, decode(x.现价,null,decode(nvl(b.零售价,0),0,b.实际金额 / b.实际数量,b.零售价),x.现价) As 现价" & _
            " From 药品收发记录 A, 药品库存 B, 收费项目目录 C , " & _
            "      (Select x.药品id,x.库房id,x.批次,x.现价 from 药品价格记录 x where x.价格类型 = 1 and (x.终止日期 Is Null Or Sysdate Between x.执行日期 And Nvl(x.终止日期, To_Date('3000-01-01', 'yyyy-MM-dd')))) X" & _
            " Where a.单据 = 11 And a.No = [1] And c.Id = a.药品id And Round(a.零售价," & intPriceDigit & ") <> Round(decode(x.现价,null,decode(nvl(b.零售价,0),0,b.实际金额 / b.实际数量,b.零售价),x.现价), " & intPriceDigit & ") And Nvl(c.是否变价, 0) = 1 And" & _
                  " b.性质 = 1 And b.库房id = a.库房id And b.药品id = a.药品id And NVL(b.批次, 0) = NVL(a.批次, 0) And NVL(b.实际数量, 0) <> 0 And a.入出系数 = -1" & _
                  " AND a.药品id = x.药品id(+) And a.库房id = x.库房id(+) And Nvl(a.批次, 0) = Nvl(x.批次(+), 0) " & _
            " Union All" & _
            " Select '成本价' As 类型, a.序号, a.药品id ,nvl(a.批次,0) as 批次, 0 原价, decode(x.现价,null,b.平均成本价,x.现价) As 现价" & _
            " From 药品收发记录 A, 药品库存 B , " & _
            "      (Select x.药品id,x.库房id,x.批次,x.现价 from 药品价格记录 x where x.价格类型 = 2 and (x.终止日期 Is Null Or Sysdate Between x.执行日期 And Nvl(x.终止日期, To_Date('3000-01-01', 'yyyy-MM-dd')))) X" & _
            " Where a.单据 = 11 And a.No = [1] And a.药品id = b.药品id And Nvl(a.批次, 0) = Nvl(b.批次, 0) and round(a.成本价," & intCostDigit & ")<>round(decode(x.现价,b.平均成本价,x.现价)," & intCostDigit & ") And a.库房id = b.库房id and a.入出系数=-1 and b.性质=1" & _
            " AND a.药品id = x.药品id(+) And a.库房id = x.库房id(+) And Nvl(a.批次, 0) = Nvl(x.批次(+), 0) " & _
            " Order By 类型, 药品id, 序号"
    Set rsPrice = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[取当前价格]", CStr(Me.txtNo.Tag))
    
    If rsPrice.EOF Then Exit Sub
    
    lngRows = mshBill.rows - 1
    For lngRow = 1 To lngRows
        blnAdj = False
        lng药品ID = Val(mshBill.TextMatrix(lngRow, 0))
        Dbl数量 = Val(mshBill.TextMatrix(lngRow, mconIntCol数量))
        dbl成本价 = Val(mshBill.TextMatrix(lngRow, mconIntCol采购价))
        dbl零售价 = Val(mshBill.TextMatrix(lngRow, mconIntCol售价))
        dbl成本金额 = dbl成本价 * Dbl数量
        dbl零售金额 = dbl零售价 * Dbl数量
        dbl差价 = dbl零售金额 - dbl成本金额
                
        If lng药品ID <> 0 Then
            rsPrice.Filter = "类型='售价' And 药品ID=" & lng药品ID & " And 批次=" & Val(mshBill.TextMatrix(lngRow, mconIntCol批次))
            If rsPrice.RecordCount > 0 Then
                blnAdj = True
                dbl零售价 = Val(FormatEx(rsPrice!现价 * Val(mshBill.TextMatrix(lngRow, mconIntCol比例系数)), intPriceDigit))
                dbl零售金额 = Val(zlStr.FormatEx(Val(FormatEx(dbl零售价, intPriceDigit)) * Dbl数量, mintMoneyDigit, , True))
                dbl差价 = Val(zlStr.FormatEx(dbl零售金额 - dbl成本金额, mintMoneyDigit, , True))
            End If
            
            rsPrice.Filter = "类型='成本价' And 药品ID=" & lng药品ID & " And 批次=" & Val(mshBill.TextMatrix(lngRow, mconIntCol批次))
            If rsPrice.RecordCount > 0 Then
                blnAdj = True
                dbl零售金额 = Val(zlStr.FormatEx(Val(FormatEx(dbl零售价, intPriceDigit)) * Dbl数量, mintMoneyDigit, , True))
                dbl成本价 = Val(FormatEx(rsPrice!现价 * Val(mshBill.TextMatrix(lngRow, mconIntCol比例系数)), intCostDigit))
                dbl成本金额 = Val(zlStr.FormatEx(dbl成本价 * Dbl数量, mintMoneyDigit, , True))
                dbl差价 = Val(zlStr.FormatEx(dbl零售金额 - dbl成本金额, mintMoneyDigit, , True))
            End If
            
            If blnAdj = True Then
                '以当前最新价格最新单据相关数据（售价、成本价、零售金额、成本金额、差价）
                mshBill.TextMatrix(lngRow, mconIntCol售价) = zlStr.FormatEx(dbl零售价, intPriceDigit, , True)
                mshBill.TextMatrix(lngRow, mconIntCol售价金额) = zlStr.FormatEx(dbl零售金额, mintMoneyDigit, , True)
                mshBill.TextMatrix(lngRow, mconIntCol采购价) = zlStr.FormatEx(dbl成本价, intCostDigit, , True)
                mshBill.TextMatrix(lngRow, mconIntCol采购金额) = zlStr.FormatEx(dbl成本金额, mintMoneyDigit, , True)
                mshBill.TextMatrix(lngRow, mconintCol差价) = zlStr.FormatEx(dbl差价, mintMoneyDigit, , True)
            End If
        End If
    Next
    rsPrice.Filter = 0
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CheckQuantity(ByVal intRow As Integer, ByVal dbl填写数量 As Double, ByVal dbl比例系数 As Integer) As Boolean
    '功能：填单时，检查实际数量是否足够，批次>0说明是按照批次出库，批次=0说明是整体出库，两种方式都需要检查库存
    '返回值：true-库存足够，false-库存不足够
    Dim rsData As ADODB.Recordset
    Dim lng药品ID As Long
    Dim lng批次 As Long
    Dim lng库房ID As Long
    Dim dbl实际数量 As Double
    
    With mshBill
        lng药品ID = Val(.TextMatrix(intRow, 0))
        lng批次 = Val(.TextMatrix(intRow, mconIntCol批次))
        lng库房ID = cboStock.ItemData(cboStock.ListIndex)
        
        If lng批次 > 0 Then
            gstrSQL = "Select (a.实际数量 - [1]) As 剩余数量,a.实际数量" & vbNewLine & _
                        "From 药品库存 a" & vbNewLine & _
                        "Where a.药品id = [2] And a.库房id = [3] And Nvl(a.批次, 0) = [4] and a.性质 = 1"
        Else
            gstrSQL = "Select Sum(a.实际数量) - [1] As 剩余数量, Sum(a.实际数量) As 实际数量" & vbNewLine & _
                        "From 药品库存 A" & vbNewLine & _
                        "Where a.药品id = [2] And a.库房id = [3] And a.性质 = 1"
        End If
        Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "库存检查", dbl填写数量 * dbl比例系数, lng药品ID, lng库房ID, lng批次)
        
        If lng批次 > 0 Then
            If rsData.RecordCount > 0 Then
                dbl实际数量 = zlStr.FormatEx(nvl(rsData!实际数量, 0) / dbl比例系数, mintNumberDigit, , True)
                
                If rsData!剩余数量 >= 0 Then
                    CheckQuantity = True
                Else
                    CheckQuantity = False
                End If
            Else
                CheckQuantity = False
            End If
        Else
            If rsData.RecordCount > 0 Then
                dbl实际数量 = zlStr.FormatEx(nvl(rsData!实际数量, 0) / dbl比例系数, mintNumberDigit, , True)
                
                If IsNull(rsData!剩余数量) Then
                    CheckQuantity = False
                Else
                    If rsData!剩余数量 >= 0 Then
                        CheckQuantity = True
                    Else
                        CheckQuantity = False
                    End If
                End If
            Else
                CheckQuantity = False
            End If
        End If
        
        If CheckQuantity = False Then
            If mint库存检查 = 0 Then
                '0-不足不检查
                CheckQuantity = True
            ElseIf mint库存检查 = 1 Then
                '1-检查，不足提醒
                If MsgBox("你输入的数量大于了库存实际数量(" & dbl实际数量 & ")，是否继续？", vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    CheckQuantity = True
                End If
            ElseIf mint库存检查 = 2 Then
                '2-检查，不足禁止
                MsgBox "你输入的数量大于了库存实际数量(" & dbl实际数量 & ")", vbInformation, gstrSysName
            End If
        End If
    End With
End Function


Private Sub Form_Load()
    Dim rsTemp As New Recordset
    Dim rsDetail As New Recordset
    Dim bln中药库房 As Boolean
    Dim str库房性质 As String
    
    On Error GoTo errHandle
    mblnLoad = False
    mblnEnterCell = False
    mstrLike = IIf(Val(zlDataBase.GetPara("输入匹配")) = 0, "%", "")
    mblnViewCost = zlStr.IsHavePrivs(mstrPrivs, "查看成本价")
    txtNo = mstr单据号
    txtNo.Tag = txtNo
    mblnUpdate = False
    
    mintDrugNameShow = Int(Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品其他出库管理", "药品名称显示方式", 0)))
    If mintDrugNameShow > 2 Or mintDrugNameShow < 0 Then mintDrugNameShow = 0
    mnuColDrug.Item(mintDrugNameShow).Checked = True
    
    Call GetSysParm
    
    With cboType
        .Clear
        gstrSQL = "SELECT b.Id,b.名称 " _
            & " FROM 药品单据性质 A, 药品入出类别 B " _
            & "Where A.类别id = B.ID " _
              & "AND A.单据 = 11 "
        Call SQLTest(App.Title, "药品其他出库单", gstrSQL)
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "Form_Load")
        Call SQLTest
        
        Do While Not rsTemp.EOF
            .AddItem rsTemp.Fields(1)
            .ItemData(.NewIndex) = rsTemp.Fields(0)
            rsTemp.MoveNext
        Loop
        rsTemp.Close
        .ListIndex = 0
    End With
    
    With cbo外调单位
        .Clear
        gstrSQL = "Select Rownum As Id, 编码, 简码, 名称 From 药品外调单位 Order By 编码"
        Call zlDataBase.OpenRecordset(rsTemp, gstrSQL, "读取外调单位")
        
        .AddItem ""
        Do While Not rsTemp.EOF
            .AddItem rsTemp!编码 & "-" & rsTemp!名称
            rsTemp.MoveNext
        Loop
        rsTemp.Close
        .ListIndex = 0
    End With
    
    With cbo外销单位
        .Clear
        gstrSQL = "Select Rownum As Id, 编码, 简码, 名称 From 药品外销单位 Order By 编码"
        Call zlDataBase.OpenRecordset(rsTemp, gstrSQL, "读取外销单位")
        
        .AddItem ""
        Do While Not rsTemp.EOF
            .AddItem rsTemp!编码 & "-" & rsTemp!名称
            rsTemp.MoveNext
        Loop
        rsTemp.Close
        .ListIndex = 0
    End With
    
    mlng出库库房 = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    Call GetDrugDigit(mlng出库库房, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    
    Call initCard
    
    mstrTime_Start = GetBillInfo(11, mstr单据号)
    RestoreWinState Me, App.ProductName, MStrCaption
    
    '只有中药类库房才显示"原产地"列
    str库房性质 = ""
    gstrSQL = "Select 工作性质 From 部门性质说明 Where 部门id =[1]"
    Set rsDetail = zlDataBase.OpenSQLRecord(gstrSQL, "判断库房性质", cboStock.ItemData(cboStock.ListIndex))
    Do While Not rsDetail.EOF
        str库房性质 = str库房性质 & "," & rsDetail!工作性质
        rsDetail.MoveNext
    Loop
    If str库房性质 Like "*中药*" Or str库房性质 Like "*制剂室*" Then bln中药库房 = True
    mshBill.ColWidth(mconIntCol原产地) = IIf(bln中药库房, 800, 0)
    
    mshBill.ColWidth(mconIntCol冲销数量) = IIf(mint编辑状态 = 6, 1100, 0)
    
    '根据人员权限决定是否显示成本价
    mshBill.ColWidth(mconIntCol采购价) = IIf(mblnViewCost, 1000, 0)
    mshBill.ColWidth(mconIntCol采购金额) = IIf(mblnViewCost, 900, 0)
    mshBill.ColWidth(mconintCol差价) = IIf(mblnViewCost, 900, 0)
    
    '商品名列处理
    If gint药品名称显示 = 2 Then
        '显示商品名列
        mshBill.ColWidth(mconIntCol商品名) = IIf(mshBill.ColWidth(mconIntCol商品名) = 0, 2000, mshBill.ColWidth(mconIntCol商品名))
    Else
        '不单独显示商品名列
        mshBill.ColWidth(mconIntCol商品名) = 0
    End If
    
    mblnEnterCell = True
    
    Call cboType_click
    mblnChange = False
    mblnLoad = True
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub initCard()
    Dim i As Integer
    Dim rsInitCard As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim intRow As Integer
    Dim numUseAbleCount As Double
    Dim vardrug As Variant
    Dim strOrder As String, strCompare As String
    Dim intCostDigit As Integer        '成本价小数位数
    Dim intPriceDigit As Integer       '售价小数位数
    Dim intNumberDigit As Integer      '数量小数位数
    Dim intMoneyDigit As Integer       '金额小数位数
    Dim str药名 As String
    Dim strSqlOrder As String
    
    On Error GoTo errHandle
    '库房
    strOrder = zlDataBase.GetPara("排序", glngSys, 模块号.其他出库)
    strCompare = Mid(strOrder, 1, 1)
    
    strSqlOrder = "序号"
    
    If strCompare = "0" Then
        strSqlOrder = "序号"
    ElseIf strCompare = "1" Then
        strSqlOrder = "药品编码"
    ElseIf strCompare = "2" Then
        If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
            strSqlOrder = "通用名"
        Else
            strSqlOrder = "Nvl(商品名, 通用名)"
        End If
    End If
    
    strSqlOrder = strSqlOrder & IIf(Right(strOrder, 1) = "0", " ASC", " DESC")
    
    intCostDigit = mintCostDigit
    intPriceDigit = mintPriceDigit
    intNumberDigit = mintNumberDigit
    intMoneyDigit = mintMoneyDigit
    
    If mint编辑状态 <> 4 Then
        With mfrmMain.cboStock
            cboStock.Clear
            For i = 0 To .ListCount - 1
                cboStock.AddItem .List(i)
                cboStock.ItemData(cboStock.NewIndex) = .ItemData(i)
            Next
            mintcboIndex = .ListIndex
            cboStock.ListIndex = .ListIndex
            cboStock.Enabled = .Enabled
        End With
    End If
    
    Select Case mint编辑状态
        Case 1
            Txt填制人 = UserInfo.用户姓名
            Txt填制日期 = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
'            Txt修改人 = UserInfo.用户姓名
'            Txt修改日期 = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            initGrid
        
        Case 2, 3, 4, 6
            Call initGrid
            
            If mint编辑状态 = 4 Then
                gstrSQL = "select distinct b.id,b.名称 from 药品收发记录 a,部门表 b  " _
                    & " where a.库房id=b.id and A.单据 =11 and  a.no=[1]"
                Set rsInitCard = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, mstr单据号)
                If rsInitCard.EOF Then
                    mintParallelRecord = 2
                    Exit Sub
                End If
                    
                With cboStock
                    .AddItem rsInitCard!名称
                    .ItemData(.NewIndex) = rsInitCard!id
                    .ListIndex = 0
                End With
                rsInitCard.Close
            End If
            
            Select Case mintUnit
                Case mconint售价单位
                    strUnitQuantity = "F.计算单位 AS 单位, A.填写数量 as 数量,a.成本价,a.零售价,nvl(a.单量,0) As 外调价,'1' as 比例系数,"
                Case mconint门诊单位
                    strUnitQuantity = "B.门诊单位 AS 单位,(A.填写数量 / B.门诊包装) AS 数量,a.成本价*B.门诊包装 as 成本价,a.零售价*B.门诊包装 as 零售价,nvl(a.单量,0)*B.门诊包装 As 外调价,B.门诊包装 as 比例系数,"
                Case mconint住院单位
                    strUnitQuantity = "B.住院单位 AS 单位,(A.填写数量 / B.住院包装) AS 数量,a.成本价*B.住院包装 as 成本价,a.零售价*B.住院包装 as 零售价,nvl(a.单量,0)*B.住院包装 As 外调价,B.住院包装 as 比例系数,"
                Case mconint药库单位
                    strUnitQuantity = "B.药库单位 AS 单位,(A.填写数量 / B.药库包装) AS 数量,a.成本价*B.药库包装 as 成本价,a.零售价*B.药库包装 as 零售价,nvl(a.单量,0)*B.药库包装 As 外调价,B.药库包装 as 比例系数,"
            End Select
            
            If mint编辑状态 <> 6 Then
                gstrSQL = "SELECT W.*,Z.可用数量,Z.实际金额,Z.实际差价 " & _
                    " FROM " & _
                    " (SELECT DISTINCT A.药品ID,A.序号,'[' || F.编码 || ']' As 药品编码, F.名称 As 通用名, E.名称 As 商品名, " & _
                    " B.药品来源,B.基本药物,F.规格,F.产地 AS 原生产商,A.产地, A.原产地,A.批号,A.批次,B.加成率,A.效期," & _
                    strUnitQuantity & _
                    " A.成本金额,A.零售金额, A.差价,A.摘要,填制人,填制日期,修改人,修改日期,审核人,审核日期,A.库房ID,A.入出类别ID,F.是否变价,B.药房分批 AS 药房分批核算," & _
                    " G.名称 AS 外调单位,A.批准文号,H.名称 AS 外销单位, To_Number(Trim(To_Char(Nvl(A.频次, '0'), '999999999999.0000'))) As 增值税率 " & _
                    " FROM 药品收发记录 A, 药品规格 B,收费项目别名 E ,收费项目目录 F,药品外调单位 G,药品外销单位 H " & _
                    " WHERE A.药品ID = B.药品ID AND B.药品ID=F.ID And A.发药窗口=G.编码(+) And A.发药窗口=H.编码(+) " & _
                    " AND B.药品ID=E.收费细目ID(+) AND E.性质(+)=3 AND E.码类(+)=1 " & _
                    " AND A.记录状态 =[3] " & _
                    " AND A.单据 = 11 AND A.NO = [1]) W," & _
                    " (SELECT  药品ID,NVL(批次,0) 批次,可用数量,实际金额,实际差价 " & _
                    " FROM 药品库存 WHERE 库房ID=[2] AND 性质=1)  Z " & _
                    " WHERE W.药品ID=Z.药品ID(+) AND NVL(W.批次,0)=Nvl(Z.批次(+),0) " & _
                    " ORDER BY " & strSqlOrder
            Else
                gstrSQL = "SELECT W.*,Z.可用数量,Z.实际金额,Z.实际差价 " & _
                    " FROM " & _
                    " (SELECT DISTINCT A.药品ID,A.序号,'[' || F.编码 || ']' As 药品编码, F.名称 As 通用名, E.名称 As 商品名, " & _
                    " B.药品来源,B.基本药物,F.规格,F.产地 AS 原生产商,A.产地, A.原产地,A.批号,A.批次,B.加成率,A.效期,G.名称 AS 外调单位,H.名称 AS 外销单位,A.增值税率," & _
                    strUnitQuantity & _
                    " A.成本金额,0 零售金额, 0 差价,A.摘要,A.库房ID,A.入出类别ID,F.是否变价,B.药房分批 AS 药房分批核算,A.批准文号,A.填写数量 As 原始数量 " & _
                    " FROM " & _
                    "     (SELECT MIN(ID) AS ID, SUM(实际数量) AS 填写数量,SUM(成本金额) AS 成本金额,药品ID,序号,产地, 原产地,批号,效期,NVL(批次,0) 批次," & _
                    " 扣率,成本价,零售价,摘要,库房ID,入出类别ID,单量,发药窗口,批准文号, To_Number(Trim(To_Char(Nvl(频次, '0'), '999999999999.0000'))) As 增值税率" & _
                    "     FROM 药品收发记录 X " & _
                    "     WHERE NO=[1] AND 单据=11  " & _
                    "     GROUP BY 药品ID,序号,产地,原产地,批号,效期,NVL(批次,0),扣率,成本价,零售价,摘要,库房ID,入出类别ID,单量,发药窗口,批准文号, To_Number(Trim(To_Char(Nvl(频次, '0'), '999999999999.0000'))) " & _
                    "     HAVING SUM(实际数量)<>0 ) A," & _
                    "     药品规格 B,收费项目别名 E ,收费项目目录 F,药品外调单位 G,药品外销单位 H " & _
                    "     WHERE A.药品ID = B.药品ID AND B.药品ID=F.ID And A.发药窗口=G.编码(+) And A.发药窗口=H.编码(+) " & _
                    "     AND B.药品ID=E.收费细目ID(+) AND E.性质(+)=3 AND E.码类(+)=1) W," & _
                    "     (SELECT  药品ID,NVL(批次,0) 批次,可用数量,实际金额,实际差价 " & _
                    "     FROM 药品库存 WHERE 库房ID=[2]  AND 性质=1)  Z " & _
                    " WHERE W.药品ID=Z.药品ID(+) AND NVL(W.批次,0)=Nvl(Z.批次(+),0) " & _
                    " ORDER BY " & strSqlOrder
            End If
            
            Set rsInitCard = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[药品其他出库单]", mstr单据号, cboStock.ItemData(cboStock.ListIndex), mint记录状态)
            
            If rsInitCard.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            
            Select Case mint编辑状态
            Case 2, 6
                If mint编辑状态 = 2 Then
                    Txt填制人 = rsInitCard!填制人
                    Txt填制日期 = Format(rsInitCard!填制日期, "yyyy-mm-dd hh:mm:ss")
                    Txt修改人 = IIf(IsNull(rsInitCard!修改人), "", rsInitCard!修改人)
                    Txt修改日期 = IIf(IsNull(rsInitCard!修改日期), "", Format(rsInitCard!修改日期, "yyyy-mm-dd hh:mm:ss"))
                End If
                If mint编辑状态 = 6 Then
                    Txt填制人 = UserInfo.用户姓名
                    Txt填制日期 = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
'                    Txt修改人 = UserInfo.用户姓名
'                    Txt修改日期 = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                    Txt审核人 = UserInfo.用户姓名
                    Txt审核日期 = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                End If
            Case Else
                Txt填制人 = rsInitCard!填制人
                Txt填制日期 = Format(rsInitCard!填制日期, "yyyy-mm-dd hh:mm:ss")
                Txt修改人 = IIf(IsNull(rsInitCard!修改人), "", rsInitCard!修改人)
                Txt修改日期 = IIf(IsNull(rsInitCard!修改日期), "", Format(rsInitCard!修改日期, "yyyy-mm-dd hh:mm:ss"))
                Txt审核人 = IIf(IsNull(rsInitCard!审核人), "", rsInitCard!审核人)
                Txt审核日期 = IIf(IsNull(rsInitCard!审核日期), "", Format(rsInitCard!审核日期, "yyyy-mm-dd hh:mm:ss"))
            End Select
            
            txt摘要.Text = IIf(IsNull(rsInitCard!摘要), "", rsInitCard!摘要)
            If (mint编辑状态 = 2 Or mint编辑状态 = 3) And Txt审核人 <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            Dim intCount As Integer
            With cboType
                For intCount = 0 To .ListCount - 1
                    If .ItemData(intCount) = rsInitCard!入出类别ID Then
                        .ListIndex = intCount
                        Exit For
                    End If
                Next
                
                If .Text = "药品外调" Then
                    Me.cbo外调单位.Visible = True
                    
                    '定位外调单位
                    If Not IsNull(rsInitCard!外调单位) Then
                        For i = 1 To cbo外调单位.ListCount - 1
                            If Mid(cbo外调单位.List(i), InStr(1, cbo外调单位.List(i), "-") + 1) = rsInitCard!外调单位 Then
                                cbo外调单位.ListIndex = i
                                Exit For
                            End If
                        Next
                    End If
                End If

                If .Text = "药品外销" Then
                    Me.cbo外销单位.Visible = True
                    
                    '定位外销单位
                    If Not IsNull(rsInitCard!外销单位) Then
                        For i = 1 To cbo外销单位.ListCount - 1
                            If Mid(cbo外销单位.List(i), InStr(1, cbo外销单位.List(i), "-") + 1) = rsInitCard!外销单位 Then
                                cbo外销单位.ListIndex = i
                                Exit For
                            End If
                        Next
                    End If
                End If
            End With
            
            If mint编辑状态 = 2 Or mint编辑状态 = 3 Then
                Set mcolUsedCount = New Collection
            End If
            
            intRow = 0
            With mshBill
                Do While Not rsInitCard.EOF
                    
                    intRow = intRow + 1
                    .rows = intRow + 1
                    
                    .TextMatrix(intRow, 0) = rsInitCard.Fields(0)
                    
                    If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
                        str药名 = rsInitCard!通用名
                    Else
                        str药名 = IIf(IsNull(rsInitCard!商品名), rsInitCard!通用名, rsInitCard!商品名)
                    End If
                    
                    .TextMatrix(intRow, mconIntCol药品编码和名称) = rsInitCard!药品编码 & str药名
                    .TextMatrix(intRow, mconIntCol药品编码) = rsInitCard!药品编码
                    .TextMatrix(intRow, mconIntCol药品名称) = str药名
                    
                    If mintDrugNameShow = 1 Then
                        .TextMatrix(intRow, mconIntCol药名) = .TextMatrix(intRow, mconIntCol药品编码)
                    ElseIf mintDrugNameShow = 2 Then
                        .TextMatrix(intRow, mconIntCol药名) = .TextMatrix(intRow, mconIntCol药品名称)
                    Else
                        .TextMatrix(intRow, mconIntCol药名) = .TextMatrix(intRow, mconIntCol药品编码和名称)
                    End If
                    
                    .TextMatrix(intRow, mconIntCol商品名) = IIf(IsNull(rsInitCard!商品名), "", rsInitCard!商品名)
                    
                    .TextMatrix(intRow, mconIntCol来源) = nvl(rsInitCard!药品来源)
                    .TextMatrix(intRow, mconIntCol基本药物) = nvl(rsInitCard!基本药物)
                    .TextMatrix(intRow, mconIntCol序号) = rsInitCard!序号
                    .TextMatrix(intRow, mconIntCol规格) = IIf(IsNull(rsInitCard!规格), "", rsInitCard!规格)
                    .TextMatrix(intRow, mconIntCol产地) = IIf(IsNull(rsInitCard!产地), "", rsInitCard!产地)
                    .TextMatrix(intRow, mconIntCol原产地) = IIf(IsNull(rsInitCard!原产地), "", rsInitCard!原产地)
                    .TextMatrix(intRow, mconIntCol单位) = rsInitCard!单位
                    .TextMatrix(intRow, mconIntCol批号) = IIf(IsNull(rsInitCard!批号), "", rsInitCard!批号)
                    .TextMatrix(intRow, mconIntCol效期) = IIf(IsNull(rsInitCard!效期), "", Format(rsInitCard!效期, "yyyy-mm-dd"))
                    If gtype_UserSysParms.P149_效期显示方式 = 1 And .TextMatrix(intRow, mconIntCol效期) <> "" Then
                        '换算为有效期
                        .TextMatrix(intRow, mconIntCol效期) = Format(DateAdd("D", -1, .TextMatrix(intRow, mconIntCol效期)), "yyyy-mm-dd")
                    End If
                    
                    .TextMatrix(intRow, mconIntCol数量) = zlStr.FormatEx(rsInitCard!数量, intNumberDigit, , True)
                    .TextMatrix(intRow, mconIntCol采购价) = zlStr.FormatEx(rsInitCard!成本价, intCostDigit, , True)
                    .TextMatrix(intRow, mconIntCol采购金额) = zlStr.FormatEx(IIf(mint编辑状态 = 6, 0, rsInitCard!成本金额), intMoneyDigit, , True)
                    .TextMatrix(intRow, mconIntCol售价) = zlStr.FormatEx(rsInitCard!零售价, intPriceDigit, , True)
                    .TextMatrix(intRow, mconIntCol售价金额) = zlStr.FormatEx(rsInitCard!零售金额, intMoneyDigit, , True)
                    .TextMatrix(intRow, mconIntCol外调价) = zlStr.FormatEx(rsInitCard!外调价, intCostDigit, , True)
                    .TextMatrix(intRow, mconIntCol外调金额) = zlStr.FormatEx(rsInitCard!外调价 * rsInitCard!数量, intMoneyDigit, , True)
                    .TextMatrix(intRow, mconintCol差价) = zlStr.FormatEx(rsInitCard!差价, intMoneyDigit, , True)
                    .TextMatrix(intRow, mconIntCol批次) = IIf(IsNull(rsInitCard!批次), "0", rsInitCard!批次)
                    .TextMatrix(intRow, mconIntCol比例系数) = rsInitCard!比例系数
                    .TextMatrix(intRow, mconIntcol加成率) = rsInitCard!加成率 / 100 & "||" & rsInitCard!是否变价 & "||" & rsInitCard!药房分批核算
                    .TextMatrix(intRow, mconIntCol可用数量) = zlStr.FormatEx(IIf(IsNull(rsInitCard!可用数量), "0", rsInitCard!可用数量), intNumberDigit, , True)
                    .TextMatrix(intRow, mconIntCol实际差价) = IIf(IsNull(rsInitCard!实际差价), "0", rsInitCard!实际差价)
                    .TextMatrix(intRow, mconIntCol实际金额) = IIf(IsNull(rsInitCard!实际金额), "0", rsInitCard!实际金额)
                    .TextMatrix(intRow, mconIntCol批准文号) = IIf(IsNull(rsInitCard!批准文号), "", rsInitCard!批准文号)
                    .TextMatrix(intRow, mconIntCol增值税率) = zlStr.FormatEx(IIf(IsNull(rsInitCard!增值税率), "0", rsInitCard!增值税率), 2, , True)
                    .TextMatrix(intRow, mconIntCol税金) = zlStr.FormatEx(rsInitCard!外调价 * rsInitCard!数量 * (Val(.TextMatrix(intRow, mconIntCol增值税率)) / 100 / (1 + Val(.TextMatrix(intRow, mconIntCol增值税率)) / 100)), intMoneyDigit, , True)
                    
                    If mint编辑状态 = 6 Then
                        .TextMatrix(intRow, mconIntCol冲销数量) = zlStr.FormatEx(0, intNumberDigit, , True)
                        .TextMatrix(intRow, mconIntCol外调金额) = zlStr.FormatEx(0, intMoneyDigit, , True)
                        .TextMatrix(intRow, mconIntCol税金) = zlStr.FormatEx(0, intMoneyDigit, , True)
                        .TextMatrix(intRow, mconintCol原始数量) = Val(rsInitCard!原始数量)
                    End If
                    
                    If mint编辑状态 = 2 Or mint编辑状态 = 3 Then
                        numUseAbleCount = 0
                        For Each vardrug In mcolUsedCount
                            If vardrug(0) = CStr(rsInitCard!药品ID & IIf(IsNull(rsInitCard!批次), "0", rsInitCard!批次)) Then
                                numUseAbleCount = vardrug(1)
                                mcolUsedCount.Remove vardrug(0)
                                Exit For
                            End If
                        Next
                        mcolUsedCount.Add Array(CStr(rsInitCard!药品ID & IIf(IsNull(rsInitCard!批次), "0", rsInitCard!批次)), CStr(numUseAbleCount + IIf(IsNull(rsInitCard!数量), "0", rsInitCard!数量))), CStr(rsInitCard!药品ID) & CStr(IIf(IsNull(rsInitCard!批次), "0", rsInitCard!批次))
                    End If
                    
                    rsInitCard.MoveNext
                Loop
                .rows = intRow + 2
            End With
            rsInitCard.Close
    End Select
    
    Call RefreshRowNO(mshBill, mconIntCol行号, 1)
    Call 显示合计金额
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Sub initGrid()
    With mshBill
        .Active = True
        .Cols = mconIntColS
        
        .MsfObj.FixedCols = 1
        
        .TextMatrix(0, mconIntCol行号) = ""
        .TextMatrix(0, mconIntCol药名) = "药品名称与编码"
        .TextMatrix(0, mconIntCol商品名) = "商品名"
        .TextMatrix(0, mconIntCol来源) = "药品来源"
        .TextMatrix(0, mconIntCol基本药物) = "基本药物"
        .TextMatrix(0, mconIntCol序号) = "序号"
        .TextMatrix(0, mconIntCol规格) = "规格"
        .TextMatrix(0, mconIntCol产地) = "生产商"
        .TextMatrix(0, mconIntCol原产地) = "原产地"
        .TextMatrix(0, mconIntCol单位) = "单位"
        .TextMatrix(0, mconIntCol批号) = "批号"
        .TextMatrix(0, mconIntCol效期) = IIf(gtype_UserSysParms.P149_效期显示方式 = 1, "有效期至", "失效期")
        .TextMatrix(0, mconIntCol批准文号) = "批准文号"
        .TextMatrix(0, mconIntCol数量) = "数量"
        .TextMatrix(0, mconIntCol冲销数量) = "冲销数量"
        .TextMatrix(0, mconIntCol采购价) = "成本价"
        .TextMatrix(0, mconIntCol采购金额) = "成本金额"
        .TextMatrix(0, mconIntCol售价) = "售价"
        .TextMatrix(0, mconIntCol售价金额) = "售价金额"
        .TextMatrix(0, mconIntCol外调价) = "外调价"
        .TextMatrix(0, mconIntCol外调金额) = "外调金额"
        .TextMatrix(0, mconIntCol增值税率) = "增值税率%"
        .TextMatrix(0, mconIntCol税金) = "税金"
        .TextMatrix(0, mconintCol差价) = "差价"
        .TextMatrix(0, mconIntCol可用数量) = "可用数量"
        .TextMatrix(0, mconIntCol实际差价) = "实际差价"
        .TextMatrix(0, mconIntCol实际金额) = "实际金额"
        .TextMatrix(0, mconIntcol加成率) = "加成率"
        .TextMatrix(0, mconIntCol比例系数) = "比例系数"
        .TextMatrix(0, mconIntCol批次) = "批次"
        .TextMatrix(0, mconIntCol药品编码和名称) = "药品编码和名称"
        .TextMatrix(0, mconIntCol药品编码) = "药品编码"
        .TextMatrix(0, mconIntCol药品名称) = "药品名称"
        .TextMatrix(0, mconintCol原始数量) = "原始数量"
        
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, mconIntCol行号) = "1"
        
        .ColWidth(0) = 0
        .ColWidth(mconIntCol行号) = 300
        .ColWidth(mconIntCol药名) = 2000
        .ColWidth(mconIntCol商品名) = 2000
        .ColWidth(mconIntCol来源) = 900
        .ColWidth(mconIntCol基本药物) = 900
        .ColWidth(mconIntCol序号) = 0
        .ColWidth(mconIntCol规格) = 900
        .ColWidth(mconIntCol产地) = 800
        .ColWidth(mconIntCol原产地) = 0
        .ColWidth(mconIntCol单位) = 500
        .ColWidth(mconIntCol批号) = 800
        .ColWidth(mconIntCol效期) = 1000
        .ColWidth(mconIntCol批准文号) = 1000
        .ColWidth(mconIntCol数量) = 1100
        .ColWidth(mconIntCol冲销数量) = IIf(mint编辑状态 = 6, 1100, 0)
        .ColWidth(mconIntCol采购价) = 1000
        .ColWidth(mconIntCol采购金额) = 1000
        .ColWidth(mconIntCol售价) = 1000
        .ColWidth(mconIntCol售价金额) = 1000
        .ColWidth(mconintCol差价) = 1000
        .ColWidth(mconIntCol外调价) = 0
        .ColWidth(mconIntCol外调金额) = 0
        .ColWidth(mconIntCol增值税率) = 0
        .ColWidth(mconIntCol税金) = 0
        
        .ColWidth(mconIntCol可用数量) = 0
        
        .ColWidth(mconIntCol实际差价) = 0
        .ColWidth(mconIntCol实际金额) = 0
        .ColWidth(mconIntcol加成率) = 0
        .ColWidth(mconIntCol比例系数) = 0
        .ColWidth(mconIntCol批次) = 0
        .ColWidth(mconIntCol药品编码和名称) = 0
        .ColWidth(mconIntCol药品编码) = 0
        .ColWidth(mconIntCol药品名称) = 0
        .ColWidth(mconintCol原始数量) = 0
        
        
        '-1：表示该列可以选择，是布尔型［"√"，" "］
        ' 0：表示该列可以选择，但不能修改
        ' 1：表示该列可以输入，外部显示为按钮选择
        ' 2：表示该列是日期列，外部显示为按钮选择，弹出是日期选择框
        ' 3：表示该列是选择列，外部显示为下拉框选择
        '4:  表示该列为单纯的文本框供用户输入
        '5:  表示该列不允许选择

        .ColData(0) = 5
        .ColData(mconIntCol商品名) = 5
        .ColData(mconIntCol行号) = 5
        .ColData(mconIntCol来源) = 5
        .ColData(mconIntCol基本药物) = 5
        .ColData(mconIntCol规格) = 5
        .ColData(mconIntCol序号) = 5
        .ColData(mconIntCol产地) = 5
        .ColData(mconIntCol原产地) = 5
        .ColData(mconIntCol单位) = 5
        .ColData(mconIntCol批号) = 5
        .ColData(mconIntCol效期) = 5
        .ColData(mconIntCol批准文号) = 5
        .ColData(mconIntCol药品编码和名称) = 5
        .ColData(mconIntCol药品编码) = 5
        .ColData(mconIntCol药品名称) = 5
        .ColData(mconIntCol增值税率) = 5
        
        chkIn.Visible = (mint编辑状态 = 1)
        txtIn.Visible = (mint编辑状态 = 1)
        
        cbo外调单位.Enabled = False
        If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
            cboType.Enabled = True
            txt摘要.Enabled = True
            
            cboStock.Enabled = True
            
            .ColData(mconIntCol药名) = 1
            .ColData(mconIntCol数量) = 4
            .ColData(mconIntCol外调价) = IIf(Me.cbo外调单位.Visible Or Me.cbo外调单位.Visible, 4, 5)
        ElseIf mint编辑状态 = 3 Or mint编辑状态 = 4 Or mint编辑状态 = 6 Then
            cboStock.Enabled = False
            cboType.Enabled = False
            txt摘要.Enabled = False
            
            .ColData(mconIntCol数量) = 5
            .ColData(mconIntCol外调价) = 5
        End If
        .ColData(mconIntCol冲销数量) = 4
        .ColData(mconIntCol采购价) = 5
        .ColData(mconIntCol采购金额) = 5
        .ColData(mconIntCol售价) = 5
        .ColData(mconIntCol售价金额) = 5
        .ColData(mconintCol差价) = 5
        .ColData(mconIntCol外调金额) = 5
        .ColData(mconIntCol税金) = 5
        .ColData(mconIntCol可用数量) = 5
        .ColData(mconIntCol实际差价) = 5
        .ColData(mconIntCol实际金额) = 5
        .ColData(mconIntcol加成率) = 5
        .ColData(mconIntCol比例系数) = 5
        .ColData(mconIntCol批次) = 5
        
        .ColAlignment(mconIntCol药名) = flexAlignLeftCenter
        .ColAlignment(mconIntCol商品名) = flexAlignLeftCenter
        .ColAlignment(mconIntCol来源) = flexAlignLeftCenter
        .ColAlignment(mconIntCol基本药物) = flexAlignLeftCenter
        .ColAlignment(mconIntCol规格) = flexAlignLeftCenter
        .ColAlignment(mconIntCol产地) = flexAlignLeftCenter
        .ColAlignment(mconIntCol原产地) = flexAlignLeftCenter
        .ColAlignment(mconIntCol单位) = flexAlignCenterCenter
        .ColAlignment(mconIntCol批号) = flexAlignLeftCenter
        .ColAlignment(mconIntCol效期) = flexAlignLeftCenter
        .ColAlignment(mconIntCol批准文号) = flexAlignLeftCenter
        .ColAlignment(mconIntCol数量) = flexAlignRightCenter
        .ColAlignment(mconIntCol采购价) = flexAlignRightCenter
        .ColAlignment(mconIntCol采购金额) = flexAlignRightCenter
        .ColAlignment(mconIntCol售价) = flexAlignRightCenter
        .ColAlignment(mconIntCol售价金额) = flexAlignRightCenter
        .ColAlignment(mconintCol差价) = flexAlignRightCenter
        
        .PrimaryCol = mconIntCol药名
        .LocateCol = mconIntCol药名
        If InStr(1, "346", mint编辑状态) <> 0 Then .ColData(mconIntCol药名) = 0
    End With
    txt摘要.MaxLength = Sys.FieldsLength("药品收发记录", "摘要")
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    With Pic单据
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - IIf(staThis.Visible, staThis.Height, 0) - .Top - 100 - CmdCancel.Height - 200
    End With
    
    With LblTitle
        .Left = 0
        .Top = 150
        .Width = Pic单据.Width
    End With
    
    With mshBill
        .Left = 200
        .Width = Pic单据.Width - .Left * 2
    End With
    With txtNo
        .Left = mshBill.Left + mshBill.Width - .Width
        LblNo.Left = .Left - LblNo.Width - 100
        .Top = LblTitle.Top
        LblNo.Top = .Top
    End With
    
    cbo外调单位.Left = mshBill.Left + mshBill.Width - cbo外调单位.Width
    lbl外调单位.Left = cbo外调单位.Left - lbl外调单位.Width - 100
    
    lbl外销单位.Left = lbl外调单位.Left
    lbl外销单位.Top = lbl外调单位.Top
    cbo外销单位.Left = cbo外调单位.Left
    cbo外销单位.Top = cbo外调单位.Top
    
    With Lbl填制人
        .Top = Pic单据.Height - 200 - .Height
        .Left = mshBill.Left + 100
    End With
    
    With Txt填制人
        .Top = Lbl填制人.Top - 80
        .Left = Lbl填制人.Left + Lbl填制人.Width + 100
    End With
    
    With Lbl填制日期
        .Top = Lbl填制人.Top
        .Left = Txt填制人.Left + Txt填制人.Width + 250
    End With
    
    With Txt填制日期
        .Top = Lbl填制日期.Top - 80
        .Left = Lbl填制日期.Left + Lbl填制日期.Width + 100
    End With
    
    With lbl修改人
        .Top = Lbl填制人.Top
        .Left = Pic单据.Width / 2 - (450 + Txt修改人.Width + lbl修改人.Width + Txt修改日期.Width + lbl修改日期.Width) / 2
    End With
    
    With Txt修改人
        .Top = Lbl填制人.Top - 80
        .Left = lbl修改人.Left + lbl修改人.Width + 100
    End With
    
    With lbl修改日期
        .Top = Lbl填制人.Top
        .Left = Txt修改人.Left + Txt修改人.Width + 250
    End With
    
    With Txt修改日期
        .Top = Lbl填制人.Top - 80
        .Left = lbl修改日期.Left + lbl修改日期.Width + 100
    End With
    
    With Txt审核日期
        .Top = Lbl填制人.Top - 80
        .Left = mshBill.Left + mshBill.Width - .Width
    End With
    
    With Lbl审核日期
        .Top = Lbl填制人.Top
        .Left = Txt审核日期.Left - 100 - .Width
    End With
    
    With Txt审核人
        .Top = Lbl填制人.Top - 80
        .Left = Lbl审核日期.Left - 200 - .Width
    End With
    
    With Lbl审核人
        .Top = Lbl填制人.Top
        .Left = Txt审核人.Left - 100 - .Width
    End With
    
    With txt摘要
        .Top = Lbl填制人.Top - 140 - .Height
        .Left = Txt填制人.Left
        .Width = mshBill.Left + mshBill.Width - .Left
    End With
    
    With lbl摘要
        .Top = txt摘要.Top + 50
        .Left = txt摘要.Left - .Width - 100
    End With
    
    With lblPurchasePrice
        .Left = mshBill.Left
        .Top = txt摘要.Top - 60 - .Height
        .Width = mshBill.Width
        lblSalePrice.Top = .Top
        lblDifference.Top = .Top
        lblOther.Top = .Top
    End With
    
    With lblSalePrice
        .Left = lblPurchasePrice.Left + mshBill.Width / 4
    End With
    With lblDifference
        .Left = lblPurchasePrice.Left + mshBill.Width / 4 * 2
    End With
    If mblnViewCost = False Then
        lblPurchasePrice.Visible = False
        lblDifference.Visible = False
    End If
    With lblOther
        .Left = lblPurchasePrice.Left + mshBill.Width / 4 * 3
    End With
    
    With mshBill
        .Height = lblPurchasePrice.Top - .Top - 60
    End With
    
    With CmdCancel
        .Left = Pic单据.Left + mshBill.Left + mshBill.Width - .Width
        .Top = Pic单据.Top + Pic单据.Height + 100
    End With
    
    With CmdSave
        .Left = CmdCancel.Left - .Width - 100
        .Top = CmdCancel.Top
    End With
    
    With cmdAllCls
        .Left = CmdSave.Left - .Width - 500
        .Top = CmdCancel.Top
    End With
    
    With cmdAllSel
        .Left = cmdAllCls.Left - .Width - 100
        .Top = CmdCancel.Top
    End With
    
    With cmdHelp
        .Left = Pic单据.Left + mshBill.Left
        .Top = CmdCancel.Top
    End With
        
    With cmdFind
        .Top = CmdCancel.Top
    End With
    
    With lblCode
        .Top = CmdCancel.Top + 50
    End With
    With txtCode
        .Top = CmdCancel.Top + 30
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\药品其他出库管理", "药品名称显示方式", mintDrugNameShow)
    
    If mblnChange = False Or mint编辑状态 = 4 Or mint编辑状态 = 3 Then
        SaveWinState Me, App.ProductName, MStrCaption
        Call ReleaseSelectorRS
        Exit Sub
    End If
    If MsgBox("数据可能已改变，但未存盘，真要退出吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
        Exit Sub
    Else
        SaveWinState Me, App.ProductName, MStrCaption
    End If
    Call ReleaseSelectorRS
End Sub

Private Function SaveCheck() As Boolean
    Dim intRow As Integer
    Dim strNo As String
    Dim lng库房ID As Long
    Dim str审核人 As String
    Dim dat审核日期 As String
    
    Dim int序号 As Integer
    Dim lng药品ID As Long
    Dim lng批次 As Long
    Dim num数量 As Double
    Dim num成本价 As Double
    Dim num成本金额 As Double
    Dim num零售金额 As Double
    Dim num差价 As Double
    Dim lng入出类别id As Long
    Dim str药品 As String
    Dim intNumCol As Integer
    
    Dim arrSql As Variant
    Dim n As Integer
    
    arrSql = Array()
    
    mblnSave = False
    SaveCheck = False
    
    lng库房ID = cboStock.ItemData(cboStock.ListIndex)
    lng入出类别id = cboType.ItemData(cboType.ListIndex)
    str审核人 = UserInfo.用户姓名
    strNo = txtNo.Tag
    
    dat审核日期 = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
    
    If mint编辑状态 = 6 Then
        intNumCol = mconIntCol冲销数量
    Else
        intNumCol = mconIntCol数量
    End If
    '检查库存
    str药品 = CheckNumStock(mshBill, cboStock.ItemData(cboStock.ListIndex), 0, mconIntCol批次, intNumCol, mconIntCol比例系数, 1, 1, mintNumberDigit)
    If str药品 <> "" Then
        If mint库存检查 = 1 Then '不足提醒
            If MsgBox("药品【" & str药品 & "】库存不足，是否继续？", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        ElseIf mint库存检查 = 2 Then '不足禁止
            MsgBox "药品【" & str药品 & "】库存不足，不能审核！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    With mshBill
        On Error GoTo errHandle
        
        '按药品ID顺序更新数据
        recSort.Sort = "药品id,批次,序号"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!行号
            If .TextMatrix(intRow, 0) <> "" Then
                If Val(.TextMatrix(intRow, mconIntCol数量)) = 0 Then
                    .TextMatrix(intRow, mconIntCol采购价) = 0
                Else
                    .TextMatrix(intRow, mconIntCol采购价) = zlStr.FormatEx((.TextMatrix(intRow, mconIntCol售价金额) - .TextMatrix(intRow, mconintCol差价)) / (.TextMatrix(intRow, mconIntCol数量)), gtype_UserDrugDigits.Digit_成本价, , True)
                End If
                .TextMatrix(intRow, mconIntCol采购金额) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol采购价) * (.TextMatrix(intRow, mconIntCol数量)), mintMoneyDigit, , True)
                
                lng药品ID = .TextMatrix(intRow, 0)
                lng批次 = .TextMatrix(intRow, mconIntCol批次)
                num数量 = .TextMatrix(intRow, mconIntCol数量) * .TextMatrix(intRow, mconIntCol比例系数)
                
'                num成本价 = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol采购价) / .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_成本价)
                num成本价 = Get成本价(lng药品ID, lng库房ID, lng批次)
                
                num成本金额 = .TextMatrix(intRow, mconIntCol采购金额)
                num零售金额 = .TextMatrix(intRow, mconIntCol售价金额)
                num差价 = .TextMatrix(intRow, mconintCol差价)
                int序号 = Val(.TextMatrix(intRow, mconIntCol序号))

                gstrSQL = "zl_药品其他出库_Verify("
                '序号
                gstrSQL = gstrSQL & int序号
                'NO
                gstrSQL = gstrSQL & ",'" & strNo & "'"
                '库房ID
                gstrSQL = gstrSQL & "," & lng库房ID
                '药品ID
                gstrSQL = gstrSQL & "," & lng药品ID
                '批次
                gstrSQL = gstrSQL & "," & lng批次
                '实际数量
                gstrSQL = gstrSQL & "," & num数量
                '成本价
                gstrSQL = gstrSQL & "," & num成本价
                '成本金额
                gstrSQL = gstrSQL & "," & num成本金额
                '零售金额
                gstrSQL = gstrSQL & "," & num零售金额
                '差价
                gstrSQL = gstrSQL & "," & num差价
                '审核人
                gstrSQL = gstrSQL & ",'" & str审核人 & "'"
                '审核日期
                gstrSQL = gstrSQL & ",to_date('" & dat审核日期 & "','yyyy-mm-dd HH24:MI:SS')"
                gstrSQL = gstrSQL & ")"

                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = CStr(lng药品ID) & ";" & gstrSQL
            End If
            recSort.MoveNext
        Next
    End With
    
    If Not ExecuteSql(arrSql, MStrCaption, False, False) Then Exit Function
    
    SaveCheck = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SaveStrike() As Boolean
    '单笔冲销 Write by zyb, ##20021016##
    Dim 行次_IN As Integer
    Dim 原记录状态_IN As Integer
    Dim NO_IN As String
    Dim 序号_IN As Integer
    Dim 药品ID_IN As Long
    Dim 冲销数量_IN As Double
    Dim 填制人_IN As String
    Dim 填制日期_IN  As String
    Dim intRow As Integer
    Dim n As Integer
    Dim str药品ID As String
    Dim i As Integer
    Dim arrSql As Variant
    Dim str药品 As String
    Dim intNumCol As Integer
    
    SaveStrike = False
    arrSql = Array()
    With mshBill
        '检查冲销数量，不能小于零
        For intRow = 1 To .rows - 1
            If Val(.TextMatrix(intRow, mconIntCol冲销数量)) <> 0 Then
                If Not 相同符号(Val(.TextMatrix(intRow, mconIntCol数量)), Val(.TextMatrix(intRow, mconIntCol冲销数量))) Then
                    MsgBox "请输入合法的冲销数量（第" & intRow & "行）！", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        Next
        
        If mint编辑状态 = 6 Then
            intNumCol = mconIntCol冲销数量
        Else
            intNumCol = mconIntCol数量
        End If
        '检查库存
        str药品 = CheckNumStock(mshBill, cboStock.ItemData(cboStock.ListIndex), 0, mconIntCol批次, intNumCol, mconIntCol比例系数, 2, 1, mintNumberDigit)
        If str药品 <> "" Then
            If mint库存检查 = 1 Then '不足提醒
                If MsgBox("药品【" & str药品 & "】库存不足，是否继续？", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            ElseIf mint库存检查 = 2 Then '不足禁止
                MsgBox "药品【" & str药品 & "】库存不足，不能审核！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        NO_IN = Trim(txtNo.Tag)
        填制人_IN = UserInfo.用户姓名
        填制日期_IN = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        原记录状态_IN = mint记录状态
        
        On Error GoTo errHandle
        
        行次_IN = 0

        '按药品ID顺序更新数据
        recSort.Sort = "药品id,批次,序号"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!行号
            If .TextMatrix(intRow, 0) <> "" And Val(.TextMatrix(intRow, mconIntCol冲销数量)) <> 0 Then
                行次_IN = 行次_IN + 1
                
                药品ID_IN = .TextMatrix(intRow, 0)
                str药品ID = IIf(str药品ID = "", "", str药品ID & ",") & 药品ID_IN
                
                If Val(.TextMatrix(intRow, mconIntCol冲销数量)) = Val(.TextMatrix(intRow, mconIntCol数量)) Then
                    '如果是全冲，冲销数量等于原始数量，避免单位换算出现的误差
                    冲销数量_IN = Val(.TextMatrix(intRow, mconintCol原始数量))
                Else
                    冲销数量_IN = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol冲销数量) * .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_数量, , True)
                End If
                
                序号_IN = .TextMatrix(intRow, mconIntCol序号)
                
                gstrSQL = "ZL_药品其他出库_STRIKE("
                '行次
                gstrSQL = gstrSQL & 行次_IN
                '原记录状态
                gstrSQL = gstrSQL & "," & 原记录状态_IN
                'NO
                gstrSQL = gstrSQL & ",'" & NO_IN & "'"
                '序号
                gstrSQL = gstrSQL & "," & 序号_IN
                '药品ID
                gstrSQL = gstrSQL & "," & 药品ID_IN
                '冲销数量
                gstrSQL = gstrSQL & "," & 冲销数量_IN
                '填制人
                gstrSQL = gstrSQL & ",'" & 填制人_IN & "'"
                '填制日期
                gstrSQL = gstrSQL & ",to_date('" & Format(填制日期_IN, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS')"
                gstrSQL = gstrSQL & ")"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            End If
            recSort.MoveNext
        Next
        
        gcnOracle.BeginTrans
        For i = 0 To UBound(arrSql)
            Call zlDataBase.ExecuteProcedure(CStr(arrSql(i)), "SaveCard")
        Next
        gcnOracle.CommitTrans
        
        If 行次_IN = 0 Then
            MsgBox "没有选择一行药品来冲销，请录入冲销数量！", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        
        '提示停用药品
        If str药品ID <> "" Then
            Call CheckStopMedi(str药品ID)
        End If
        
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveStrike = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub mnuColDrug_Click(Index As Integer)
    Dim n As Integer
    
    With mnuColDrug
        For n = 0 To .count - 1
            .Item(n).Checked = False
        Next
        
        .Item(Index).Checked = True
        
        Call SetDrugName(Index)
    End With
End Sub

Private Sub mshBill_AfterAddRow(Row As Long)
    Call RefreshRowNO(mshBill, mconIntCol行号, Row)
End Sub

Private Sub mshBill_AfterDeleteRow()
    Call 显示合计金额
    Call RefreshRowNO(mshBill, mconIntCol行号, mshBill.Row)
End Sub

Private Sub mshBill_BeforeAddRow(Row As Long)
    If mshBill.ColData(mconIntCol药名) = 0 Then
        'Cancel = True    '等待加CANCEL参数
        Exit Sub
    End If
        
        
End Sub

Private Sub mshBill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    If InStr(1, "34", mint编辑状态) <> 0 Then
        Cancel = True
        Exit Sub
    End If
    With mshBill
        If .TextMatrix(.Row, 0) <> "" Then
            If MsgBox("你确实要删除该行药品？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True
            End If
        End If
    End With
End Sub

Private Sub mshbill_CommandClick()
    Dim RecReturn As Recordset
    Dim i As Integer
    Dim intCurRow As Integer
    Dim intOldRow As Integer
    
    intOldRow = mshBill.Row
    mshBill.CmdEnable = False
    mblnChange = True
'    Set RecReturn = Frm药品选择器.ShowME(Me, 2, cboStock.ItemData(cboStock.ListIndex))
    If grsMaster.State = adStateClosed Then
        Call SetSelectorRS(2, MStrCaption, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex))
    End If
    Set RecReturn = frmSelector.ShowME(Me, 0, 2, , , , cboStock.ItemData(cboStock.ListIndex), , , , , , , , , mstrPrivs)
    If RecReturn.RecordCount > 0 Then
        Set RecReturn = CheckData(RecReturn)
    End If
      
    mshBill.CmdEnable = True
    If RecReturn.RecordCount > 0 Then
        RecReturn.MoveFirst
        For i = 1 To RecReturn.RecordCount
            intCurRow = mshBill.Row
            With mshBill
                .TextMatrix(intCurRow, mconIntCol行号) = .Row
                SetColValue .Row, RecReturn!药品ID, "[" & RecReturn!药品编码 & "]", RecReturn!通用名, IIf(IsNull(RecReturn!商品名), "", RecReturn!商品名), _
                    nvl(RecReturn!药品来源), "" & RecReturn!基本药物, _
                    IIf(IsNull(RecReturn!规格), "", RecReturn!规格), IIf(IsNull(RecReturn!产地), "", RecReturn!产地), _
                    Choose(mintUnit, RecReturn!售价单位, RecReturn!门诊单位, RecReturn!住院单位, RecReturn!药库单位), _
                    IIf(IsNull(RecReturn!售价), 0, RecReturn!售价), IIf(IsNull(RecReturn!批号), "", RecReturn!批号), _
                    IIf(IsNull(RecReturn!效期), "", Format(RecReturn!效期, "yyyy-MM-dd")), _
                    IIf(IsNull(RecReturn!可用数量), "0", RecReturn!可用数量), _
                    IIf(IsNull(RecReturn!实际金额), "0", RecReturn!实际金额), _
                    IIf(IsNull(RecReturn!实际差价), "0", RecReturn!实际差价), _
                    IIf(IsNull(RecReturn!加成率), "0", RecReturn!加成率 / 100), _
                    Choose(mintUnit, 1, RecReturn!门诊包装, RecReturn!住院包装, RecReturn!药库包装), _
                    IIf(IsNull(RecReturn!批次), 0, RecReturn!批次), RecReturn!时价, RecReturn!药房分批, _
                    IIf(IsNull(RecReturn!批准文号), "", RecReturn!批准文号), nvl(RecReturn!原产地)
                .Col = mconIntCol数量
                
                If (.TextMatrix(intCurRow, 0) = "" Or intCurRow = 1 Or .Row = .rows - 1) And .TextMatrix(.rows - 1, 0) <> "" Then
                    .rows = .rows + 1
                End If
                .Row = .rows - 1
                RecReturn.MoveNext
            End With
        Next
        mshBill.Row = intOldRow
        RecReturn.Close
    End If
End Sub


Private Sub mshbill_EditChange(curText As String)
    mblnChange = True
End Sub


Private Sub mshBill_EditKeyPress(KeyAscii As Integer)
    Dim strKey As String
    Dim intDigit As Integer
    
    With mshBill
        strKey = .Text
        If strKey = "" Then
            strKey = .TextMatrix(.Row, .Col)
        End If
        
        If .Col = mconIntCol数量 Or .Col = mconIntCol冲销数量 Or .Col = mconIntCol采购价 Or .Col = mconIntCol外调价 Or .Col = mconIntCol售价 Or .Col = mconIntCol采购金额 Or .Col = mconIntCol外调金额 Then
            Select Case .Col
                Case mconIntCol数量, mconIntCol冲销数量
                    intDigit = mintNumberDigit
                Case mconIntCol采购价, mconIntCol外调价
                   intDigit = mintCostDigit
                Case mconIntCol售价
                    intDigit = mintPriceDigit
                Case mconIntCol采购金额, mconIntCol外调金额
                    intDigit = mintMoneyDigit
            End Select
            
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
                If .SelLength = Len(strKey) Then Exit Sub
                If Len(Mid(strKey, InStr(1, strKey, ".") + 1)) >= intDigit And strKey Like "*.*" Then
                    KeyAscii = 0
                    Exit Sub
                Else
                    Exit Sub
                End If
            End If
        End If
    
    End With
End Sub

Private Sub mshbill_EnterCell(Row As Long, Col As Long)
    If Not mblnEnterCell Then Exit Sub
    With mshBill
        If Row > 0 Then
            .SetRowColor CLng(Row), &HFFCECE, True
        End If
        Select Case .Col
            Case mconIntCol药名
                .txtCheck = False
                .MaxLength = 40
                '只在药名列才显示合计信息和库存数
                Call 显示合计金额
                Call 提示库存数
            Case mconIntCol数量
                .txtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
                Call 提示库存数
        End Select
    End With
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim rsTemp As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim i As Integer
    Dim intCurRow As Integer
    Dim intOldRow As Integer
    
    intOldRow = mshBill.Row
    On Error GoTo errHandle
    If KeyCode <> vbKeyReturn Then Exit Sub
    With mshBill
        .Text = UCase(Trim(.Text))
        strKey = UCase(Trim(.Text))
        
        If Mid(strKey, 1, 1) = "[" Then
            If InStr(2, strKey, "]") <> 0 Then
                strKey = Mid(strKey, 2, InStr(2, strKey, "]") - 2)
            Else
                strKey = Mid(strKey, 2)
            End If
        End If
        Select Case .Col
            Case mconIntCol药名
                If strKey <> "" Then
                    Dim RecReturn As Recordset
                    Dim sngLeft As Single
                    Dim sngTop As Single
                    
                    sngLeft = Me.Left + Pic单据.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                    sngTop = Me.Top + Me.Height - Me.ScaleHeight + Pic单据.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                    If sngTop + 3630 > Screen.Height Then
                        sngTop = sngTop - mshBill.MsfObj.CellHeight - 4530
                    End If
                    
'                    Set RecReturn = Frm药品多选选择器.ShowME(Me, 2, cboStock.ItemData(cboStock.ListIndex), , , strkey, sngLeft, sngTop)
                    If grsMaster.State = adStateClosed Then
                        Call SetSelectorRS(2, MStrCaption, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex))
                    End If
                    Set RecReturn = frmSelector.ShowME(Me, 1, 2, strKey, sngLeft, sngTop, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , , , , , , , mstrPrivs)
                    If RecReturn.RecordCount > 0 Then
                        Set RecReturn = CheckData(RecReturn)
                    End If
                    
                    If RecReturn.RecordCount > 0 Then
                        RecReturn.MoveFirst
                        For i = 1 To RecReturn.RecordCount
                            intCurRow = .Row
                            .TextMatrix(intCurRow, mconIntCol行号) = .Row
                            If SetColValue(.Row, RecReturn!药品ID, "[" & RecReturn!药品编码 & "]", RecReturn!通用名, IIf(IsNull(RecReturn!商品名), "", RecReturn!商品名), _
                                    nvl(RecReturn!药品来源), "" & RecReturn!基本药物, _
                                    IIf(IsNull(RecReturn!规格), "", RecReturn!规格), IIf(IsNull(RecReturn!产地), "", RecReturn!产地), _
                                    Choose(mintUnit, RecReturn!售价单位, RecReturn!门诊单位, RecReturn!住院单位, RecReturn!药库单位), _
                                    IIf(IsNull(RecReturn!售价), 0, RecReturn!售价), IIf(IsNull(RecReturn!批号), "", RecReturn!批号), _
                                    IIf(IsNull(RecReturn!效期), "", Format(RecReturn!效期, "yyyy-MM-dd")), _
                                    IIf(IsNull(RecReturn!可用数量), "0", RecReturn!可用数量), _
                                    IIf(IsNull(RecReturn!实际金额), "0", RecReturn!实际金额), _
                                    IIf(IsNull(RecReturn!实际差价), "0", RecReturn!实际差价), _
                                    IIf(IsNull(RecReturn!加成率), "0", RecReturn!加成率), _
                                    Choose(mintUnit, 1, RecReturn!门诊包装, RecReturn!住院包装, RecReturn!药库包装), _
                                    IIf(IsNull(RecReturn!批次), 0, RecReturn!批次), RecReturn!时价, RecReturn!药房分批, _
                                    IIf(IsNull(RecReturn!批准文号), "", RecReturn!批准文号), nvl(RecReturn!原产地)) = False Then
                                Cancel = True
                                Exit Sub
                            End If
                            .Text = .TextMatrix(.Row, .Col)
                            If (.TextMatrix(intCurRow, 0) = "" Or intCurRow = 1 Or .Row = .rows - 1) And .TextMatrix(.rows - 1, 0) <> "" Then
                                .rows = .rows + 1
                            End If
                            .Row = .rows - 1
                            RecReturn.MoveNext
                        Next
                        .Row = intOldRow
                    Else
                        If Val(.TextMatrix(.Row, 0)) = 0 Then
                            .Text = .TextMatrix(.Row, .Col)
                            Cancel = True
                        Else
                            .Text = .TextMatrix(.Row, .Col)
                        End If
                    End If
                    Call 提示库存数
                End If
            
            Case mconIntCol数量
                If .TextMatrix(.Row, 0) = "" Then
                    .Text = ""
                    Exit Sub
                End If
                
                If .TextMatrix(.Row, .Col) = "" And strKey = "" Then
                    MsgBox "对不起，数量必须输入！", vbOKOnly + vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "对不起，数量必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Val(strKey) = 0 Then
                        MsgBox "对不起，数量必须大于零,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strKey) < 0 Then
                        If Not zlStr.IsHavePrivs(mstrPrivs, "负数开单") Then
                            MsgBox "对不起，你没有负数开单的权限，请重输！", vbInformation, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                    
                    '检查库存
                    If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
                        If Not CheckUsableNum(cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol批次)), strKey, Val(mshBill.TextMatrix(.Row, mconIntCol比例系数)), txtNo.Caption, 11, mint库存检查, mintNumberDigit) Then
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                    
                    strKey = zlStr.FormatEx(strKey, mintNumberDigit, , True)
                    .Text = strKey
                    
                    If .TextMatrix(.Row, mconIntCol售价) <> "" Then
                        .TextMatrix(.Row, mconIntCol售价金额) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol售价) * strKey, mintMoneyDigit, , True)
                    End If
                    
                    If strKey <> 0 Then
                        .TextMatrix(.Row, mconIntCol采购价) = zlStr.FormatEx(Get成本价(Val(.TextMatrix(.Row, 0)), Val(cboStock.ItemData(cboStock.ListIndex)), Val(.TextMatrix(.Row, mconIntCol批次))) * Val(Val(mshBill.TextMatrix(.Row, mconIntCol比例系数))), mintCostDigit, , True)
                    End If
                    .TextMatrix(.Row, mconIntCol采购金额) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol采购价)) * strKey, mintMoneyDigit, , True)
                    .TextMatrix(.Row, mconintCol差价) = zlStr.FormatEx(Val(Val(.TextMatrix(.Row, mconIntCol售价金额))) - Val(.TextMatrix(.Row, mconIntCol采购金额)), mintMoneyDigit, , True)
                    
                    '计算外调价及外调金额:外调价=(1+管理费比例)*进价
                    If Val(.TextMatrix(.Row, 0)) <> 0 And cboType.Text = "药品外调" And .TextMatrix(.Row, mconIntCol外调价) = "" Then
                        gstrSQL = "Select Nvl(管理费比例,0) 比例 From 药品规格 Where 药品ID=[1]"
                        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[取该药品的管理费比例]", Val(.TextMatrix(.Row, 0)))
                        
                        .TextMatrix(.Row, mconIntCol外调价) = zlStr.FormatEx((1 + rsTemp!比例 / 100) * Val(.TextMatrix(.Row, mconIntCol采购价)), mintCostDigit, , True)
                    End If
                    .TextMatrix(.Row, mconIntCol外调金额) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol外调价)) * Val(strKey), mintMoneyDigit, , True)
                    
                    '税金=外销金额*增值税率
                    .TextMatrix(.Row, mconIntCol税金) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol外调价)) * Val(strKey) * (Val(.TextMatrix(.Row, mconIntCol增值税率)) / 100 / (1 + Val(.TextMatrix(.Row, mconIntCol增值税率)) / 100)), mintMoneyDigit, , True)
                End If
                显示合计金额
            
            Case mconIntCol冲销数量
                If .TextMatrix(.Row, 0) = "" Then
                    .Text = ""
                    Exit Sub
                End If
            
                If .TextMatrix(.Row, .Col) = "" And strKey = "" Then
                    MsgBox "对不起，数量必须输入！", vbOKOnly + vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "对不起，数量必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Val(strKey) < 0 Then
                        If Not zlStr.IsHavePrivs(mstrPrivs, "负数开单") Then
                            MsgBox "对不起，你没有负数开单的权限，请重输！", vbInformation, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                    
                    If Val(strKey) >= 0 Then
                        If Val(strKey) > Val(.TextMatrix(.Row, mconIntCol数量)) Then
                            MsgBox "对不起，冲销数量不能大于原有数量,请重输！", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    Else
                        If Val(strKey) < Val(.TextMatrix(.Row, mconIntCol数量)) Then
                            MsgBox "对不起，冲销数量不能大于原有数量,请重输！", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                    
                    If Val(strKey) >= 10 ^ 11 - 1 Then
                        MsgBox "冲销数量必须小于" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    strKey = zlStr.FormatEx(strKey, mintNumberDigit, , True)
                    .Text = strKey
                    
                    If .TextMatrix(.Row, mconIntCol采购价) <> "" Then
                        .TextMatrix(.Row, mconIntCol采购金额) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol采购价) * Val(strKey), mintMoneyDigit, , True)
                    End If
                    If .TextMatrix(.Row, mconIntCol售价) <> "" Then
                        .TextMatrix(.Row, mconIntCol售价金额) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol售价) * Val(strKey), mintMoneyDigit, , True)
                    End If
                    If .TextMatrix(.Row, mconIntCol外调价) <> "" Then
                        .TextMatrix(.Row, mconIntCol外调金额) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol外调价)) * Val(strKey), mintMoneyDigit, , True)
                    End If
                    .TextMatrix(.Row, mconIntCol税金) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol外调价)) * Val(strKey) * (Val(.TextMatrix(.Row, mconIntCol增值税率)) / 100 / (1 + Val(.TextMatrix(.Row, mconIntCol增值税率)) / 100)), mintMoneyDigit, , True)
                    .TextMatrix(.Row, mconintCol差价) = zlStr.FormatEx(IIf(.TextMatrix(.Row, mconIntCol售价金额) = "", 0, .TextMatrix(.Row, mconIntCol售价金额)) - IIf(.TextMatrix(.Row, mconIntCol采购金额) = "", 0, .TextMatrix(.Row, mconIntCol采购金额)), mintMoneyDigit, , True)
                End If
                显示合计金额
            Case mconIntCol外调价
                If Val(.TextMatrix(.Row, 0)) = 0 Then Exit Sub
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "外调价必须为数字型，请重输！", vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If strKey <> "" Then
                    If Val(strKey) < 0.001 Then
                        MsgBox "对不起，外调价必须大于0.001,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strKey) >= 10 ^ 11 - 1 Then
                        MsgBox "外调价必须小于" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    .Text = zlStr.FormatEx(strKey, mintCostDigit, , True)
                    .TextMatrix(.Row, .Col) = .Text
                    
                    '重算外调金额
                    .TextMatrix(.Row, mconIntCol外调金额) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol外调价)) * Val(.TextMatrix(.Row, mconIntCol数量)), mintMoneyDigit, , True)
                    
                    '重算税金
                    .TextMatrix(.Row, mconIntCol税金) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol外调价)) * Val(.TextMatrix(.Row, mconIntCol数量)) * (Val(.TextMatrix(.Row, mconIntCol增值税率)) / 100 / (1 + Val(.TextMatrix(.Row, mconIntCol增值税率)) / 100)), mintMoneyDigit, , True)
                End If
        End Select
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'从药品目录中取值并附给相应的列
Private Function SetColValue(ByVal intRow As Integer, ByVal lng药品ID As Long, _
    ByVal str药品编码 As String, ByVal str通用名 As String, ByVal str商品名 As String, _
    ByVal str药品来源 As String, ByVal str基本药物 As String, ByVal str规格 As String, ByVal str产地 As String, _
    ByVal str单位 As String, ByVal num售价 As Double, ByVal str批号 As String, _
    ByVal str效期 As String, ByVal num可用数量 As Double, ByVal num实际金额 As Double, _
    ByVal num实际差价 As Double, ByVal dbl加成率 As Double, _
    ByVal num比例系数 As Double, ByVal lng批次 As Long, _
    ByVal int是否变价 As Integer, ByVal int药房分批 As Integer, ByVal str批准文号 As String, ByVal str原产地 As String) As Boolean
    
    Dim intCount As Integer
    Dim intCol As Integer
    Dim dblPrice As Double
    Dim rsPrice As New Recordset
    
    Dim dbl外调价 As Double
    Dim dbl增值税率 As Double
    Dim dbl税金 As Double
    Dim str药名 As String
    
    SetColValue = False
    On Error GoTo errHandle
    
    With mshBill

        For intCol = 0 To .Cols - 1
            If intCol <> mconIntCol行号 Then .TextMatrix(intRow, intCol) = ""
        Next
        
        .TextMatrix(intRow, 0) = lng药品ID
        
        If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
            str药名 = str通用名
        Else
            str药名 = IIf(str商品名 <> "", str商品名, str通用名)
        End If
        
        .TextMatrix(intRow, mconIntCol药品编码和名称) = str药品编码 & str药名
        .TextMatrix(intRow, mconIntCol药品编码) = str药品编码
        .TextMatrix(intRow, mconIntCol药品名称) = str药名
        
        If mintDrugNameShow = 1 Then
            .TextMatrix(intRow, mconIntCol药名) = .TextMatrix(intRow, mconIntCol药品编码)
        ElseIf mintDrugNameShow = 2 Then
            .TextMatrix(intRow, mconIntCol药名) = .TextMatrix(intRow, mconIntCol药品名称)
        Else
            .TextMatrix(intRow, mconIntCol药名) = .TextMatrix(intRow, mconIntCol药品编码和名称)
        End If
        
        .TextMatrix(intRow, mconIntCol商品名) = str商品名
        .TextMatrix(intRow, mconIntCol来源) = str药品来源
        .TextMatrix(intRow, mconIntCol基本药物) = str基本药物
        .TextMatrix(intRow, mconIntCol规格) = str规格
        .TextMatrix(intRow, mconIntCol产地) = str产地
        .TextMatrix(intRow, mconIntCol原产地) = str原产地
        .TextMatrix(intRow, mconIntCol单位) = str单位
        .TextMatrix(intRow, mconIntCol批号) = str批号
        .TextMatrix(intRow, mconIntCol效期) = Format(str效期, "yyyy-mm-dd")
        .TextMatrix(intRow, mconIntCol批准文号) = str批准文号
        .TextMatrix(intRow, mconIntCol售价) = zlStr.FormatEx(num售价 * num比例系数, mintPriceDigit, , True)
        .TextMatrix(intRow, mconIntCol可用数量) = zlStr.FormatEx(num可用数量, mintNumberDigit, , True)
        .TextMatrix(intRow, mconIntCol实际差价) = num实际差价
        .TextMatrix(intRow, mconIntCol实际金额) = num实际金额
        .TextMatrix(intRow, mconIntcol加成率) = dbl加成率 & "||" & int是否变价 & "||" & int药房分批
        .TextMatrix(intRow, mconIntCol比例系数) = num比例系数
        .TextMatrix(intRow, mconIntCol批次) = lng批次
        .TextMatrix(intRow, mconIntCol增值税率) = "100.00"
        
        If int是否变价 = 1 Then
            dblPrice = Get零售价(lng药品ID, Val(cboStock.ItemData(cboStock.ListIndex)), lng批次, num比例系数)
            .TextMatrix(intRow, mconIntCol售价) = zlStr.FormatEx(dblPrice, mintPriceDigit, , True)
        End If
        
        If IsLowerLimit(cboStock.ItemData(cboStock.ListIndex), lng药品ID) Then Call SetForeColor_ROW(mlng紫色)
        Call CheckLapse(str效期)
                
        If cboType.Text = "药品外销" Then
            '外销价默认为采购价=结算价/扣率
            gstrSQL = "Select A.指导批发价, A.增值税率, Nvl(B.采购价,0) As 采购价 " & _
                " From 药品规格 A, " & _
                " (Select 药品id, 上次采购价 / Nvl(上次扣率, 100) * 100 As 采购价 " & _
                " From 药品库存 " & _
                " Where 性质 = 1 And 库房id + 0 = [1] And 药品id = [2] And Nvl(批次, 0) = [3]) B " & _
                " Where A.药品id = B.药品id(+) And A.药品id = [2]"
            Set rsPrice = zlDataBase.OpenSQLRecord(gstrSQL, "取药品外销信息", Val(cboStock.ItemData(cboStock.ListIndex)), lng药品ID, lng批次)
            
            If Not rsPrice.EOF Then
                .TextMatrix(intRow, mconIntCol增值税率) = zlStr.FormatEx(rsPrice!增值税率, 2)
                
                If rsPrice!采购价 > 0 Then
                    .TextMatrix(intRow, mconIntCol外调价) = zlStr.FormatEx(rsPrice!采购价 * num比例系数, mintPriceDigit, , True)
                Else
                    .TextMatrix(intRow, mconIntCol外调价) = zlStr.FormatEx(rsPrice!指导批发价 * num比例系数, mintPriceDigit, , True)
                End If
            End If
        End If
    End With
    SetColValue = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Private Sub mshBill_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        With mshBill
           If .Col = mconIntCol药名 Then
                PopupMenu mnuCol, 2
            End If
        End With
    End If
End Sub

Private Sub staThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "PY" And staThis.Tag <> "PY" Then
        Logogram staThis, 0
        staThis.Tag = Panel.Key
    ElseIf Panel.Key = "WB" And staThis.Tag <> "WB" Then
        Logogram staThis, 1
        staThis.Tag = Panel.Key
    End If
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
    If KeyAscii = 13 Then
        cmdFind_Click
    End If
End Sub

Private Function ValidData() As Boolean
    ValidData = False
    Dim intLop As Integer
    
    With mshBill
        If cboType.Text = "药品外调" Then
            If cbo外调单位.ListIndex = 0 Then
                MsgBox "请选择药品外调单位！", vbInformation, gstrSysName
                cbo外调单位.SetFocus
                Exit Function
            End If
        End If
        
        If .TextMatrix(1, 0) <> "" Then         '先判有否数据
            
            If LenB(StrConv(txt摘要.Text, vbFromUnicode)) > txt摘要.MaxLength Then
                MsgBox "摘要超长,最多能输入" & CInt(txt摘要.MaxLength / 2) & "个汉字或" & txt摘要.MaxLength & "个字符!", vbInformation + vbOKOnly, gstrSysName
                txt摘要.SetFocus
                Exit Function
            End If
        
            For intLop = 1 To .rows - 1
                If Trim(.TextMatrix(intLop, mconIntCol药名)) <> "" Then
                    If Trim(Trim(.TextMatrix(intLop, mconIntCol数量))) = "" Then
                        MsgBox "第" & intLop & "行药品的数量为空了，请检查！", vbInformation, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol数量
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mconIntCol数量)) > 9999999999# Then
                        MsgBox "第" & intLop & "行药品的数量大于了数据库能够保存的" & vbCrLf & "最大范围9999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol数量
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mconIntCol采购金额)) > 9999999999999# Then
                        MsgBox "第" & intLop & "行药品的成本金额大于了数据库能够保存的" & vbCrLf & "最大范围9999999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol数量
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mconIntCol售价金额)) > 9999999999999# Then
                        MsgBox "第" & intLop & "行药品的售价金额大于了数据库能够保存的" & vbCrLf & "最大范围9999999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol数量
                        Exit Function
                    End If
                     
                    If Not CheckUsableNum(cboStock.ItemData(cboStock.ListIndex), Val(mshBill.TextMatrix(intLop, 0)), Val(mshBill.TextMatrix(intLop, mconIntCol批次)), _
                                    Val(mshBill.TextMatrix(intLop, mconIntCol数量)), Val(.TextMatrix(intLop, mconIntCol比例系数)), _
                                    Trim(txtNo.Caption), 11, mint库存检查, mintNumberDigit) Then
                        mshBill.SetFocus
                        .Row = intLop
                        .Col = mconIntCol数量
                        Exit Function
                    End If
           
                    '零差价管理：检查是否存在不满足零差价的药品
                    If gtype_UserSysParms.P275_零差价管理模式 = 2 Then
                        If IsPriceAdjustMod(Val(.TextMatrix(intLop, 0))) = True Then
                            If CheckPriceAdjust(Val(.TextMatrix(intLop, 0)), cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(intLop, mconIntCol批次))) = False Then
                                MsgBox "第" & intLop & "行药品已启用零差价管理，但库存记录中售价和成本价不一致，不能进行业务，请检查！", vbInformation + vbOKOnly, gstrSysName
                                mshBill.SetFocus
                                .Row = intLop
                                .MsfObj.TopRow = intLop
                                Exit Function
                            End If
                        End If
                    End If
                End If
            Next
        Else
            Exit Function
        End If
    End With
    
    ValidData = True
End Function

Private Function SaveCard(Optional ByVal bln强制保存 As Boolean = False) As Boolean
    Dim lng入出类别id As Long
    Dim chrNo As Variant
    Dim lngSerial As Long
    Dim lngStockid As Long
    Dim lngTypeID As Long
    Dim lngDrugID As Long
    Dim strBatchNo As String
    Dim lngBatchID As Long
    Dim strProducingArea As String
    Dim strOldProducingArea As String
    Dim datTimeLimit As String
    Dim dblQuantity As Double
    Dim dblPurchasePrice As Double
    Dim dblPurchaseMoney As Double
    Dim dblSalePrice As Double
    Dim dblSaleMoney As Double
    Dim dblMistakePrice As Double
    Dim dblOutPrice As Double   '外调价
    Dim strOutUnit As String    '外调单位
    Dim strBrief As String
    Dim strBooker As String
    Dim datBookDate As String
    Dim strModifier As String
    Dim datModifyDate As String
    Dim strAssessor As String
    Dim datAssessDate As String
    Dim arrSql As Variant
    Dim intRow As Integer
    Dim str批准文号 As String
    Dim blnTran As Boolean
    Dim dbl增值税率 As Double
    
    Dim rsTemp As New Recordset
    Dim n As Integer
    
    SaveCard = False
    arrSql = Array()
    
    '在外面设置入出类别ID，主要是所有药品都要用他
    On Error GoTo errHandle
    With mshBill
        chrNo = Trim(txtNo)
        If chrNo = "" Then chrNo = Sys.GetNextNo(28, Me.cboStock.ItemData(Me.cboStock.ListIndex))
        If IsNull(chrNo) Then Exit Function
        Me.txtNo.Tag = chrNo
        If cboType.Text = "药品外调" Then
            If cbo外调单位.Text <> "" Then
                strOutUnit = Mid(cbo外调单位.Text, 1, InStr(1, cbo外调单位.Text, "-") - 1)
            Else
                MsgBox "请输入药品外调单位！", vbInformation, gstrSysName
                SaveCard = False
                Exit Function
            End If
        ElseIf cboType.Text = "药品外销" Then
            If cbo外销单位.Text <> "" Then
                strOutUnit = Mid(cbo外销单位.Text, 1, InStr(1, cbo外销单位.Text, "-") - 1)
            Else
                MsgBox "请输入药品外销单位！", vbInformation, gstrSysName
                SaveCard = False
                Exit Function
            End If
        Else
            strOutUnit = ""
        End If
        lngStockid = cboStock.ItemData(cboStock.ListIndex)
        lng入出类别id = cboType.ItemData(cboType.ListIndex)
        strBrief = Trim(txt摘要.Text)
        strBooker = Txt填制人
        datBookDate = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        
        strAssessor = Txt审核人
        
        If bln强制保存 Then blnTran = True
        
        If mint编辑状态 = 2 Or bln强制保存 Then        '修改
            gstrSQL = "zl_药品其他出库_Delete('" & mstr单据号 & "')"
            
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = "0;" & gstrSQL
            
            strBooker = Txt填制人
            datBookDate = Format(Txt填制日期, "yyyy-mm-dd hh:mm:ss")
            strModifier = UserInfo.用户姓名
            datModifyDate = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        End If
            
        '按药品ID顺序更新数据
        recSort.Sort = "药品id,批次,序号"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!行号
            If .TextMatrix(intRow, 0) <> "" Then
                lngDrugID = .TextMatrix(intRow, 0)
                strProducingArea = .TextMatrix(intRow, mconIntCol产地)
                strOldProducingArea = .TextMatrix(intRow, mconIntCol原产地)
                strBatchNo = .TextMatrix(intRow, mconIntCol批号)
                lngBatchID = .TextMatrix(intRow, mconIntCol批次)
                datTimeLimit = IIf(.TextMatrix(intRow, mconIntCol效期) = "", "", .TextMatrix(intRow, mconIntCol效期))
                If gtype_UserSysParms.P149_效期显示方式 = 1 And datTimeLimit <> "" Then
                    '换算为失效期来保存
                    datTimeLimit = Format(DateAdd("D", 1, datTimeLimit), "yyyy-mm-dd")
                End If
                
                dblQuantity = .TextMatrix(intRow, mconIntCol数量) * .TextMatrix(intRow, mconIntCol比例系数)
                
'                dblPurchasePrice = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol采购价)) / .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_成本价)
                dblPurchasePrice = Get成本价(lngDrugID, lngStockid, lngBatchID)
                
                dblPurchaseMoney = Val(zlStr.FormatEx(Val(FormatEx(dblPurchasePrice * Val(.TextMatrix(intRow, mconIntCol比例系数)), mintCostDigit)) * Val(.TextMatrix(intRow, mconIntCol数量)), mintMoneyDigit, , True)) ' .TextMatrix(intRow, mconIntCol采购金额)
                
                dblSalePrice = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol售价) / .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_零售价)
                dblSalePrice = Get售价(Split(.TextMatrix(intRow, mconIntcol加成率), "||")(1) = 1, lngDrugID, lngStockid, lngBatchID)
                
                dblSaleMoney = Val(zlStr.FormatEx(Val(FormatEx(dblSalePrice * Val(.TextMatrix(intRow, mconIntCol比例系数)), mintPriceDigit)) * Val(.TextMatrix(intRow, mconIntCol数量)), mintMoneyDigit, , True)) ' .TextMatrix(intRow, mconIntCol售价金额)
                str批准文号 = IIf(Trim(.TextMatrix(intRow, mconIntCol批准文号)) = "", "", .TextMatrix(intRow, mconIntCol批准文号))
                
                '如果是药品外调，且外调价等于零，则直动提取管理费比例并计算外调价
                If Val(.TextMatrix(intRow, mconIntCol外调价)) = 0 And cboType.Text = "药品外调" Then
                    gstrSQL = "Select Nvl(管理费比例,0) 比例 From 药品规格 Where 药品ID=[1]"
                    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[取该药品的管理费比例]", lngDrugID)
                    
                    .TextMatrix(intRow, mconIntCol外调价) = zlStr.FormatEx((1 + rsTemp!比例 / 100) * Val(.TextMatrix(intRow, mconIntCol采购价)), gtype_UserDrugDigits.Digit_成本价)
                    .TextMatrix(intRow, mconIntCol外调金额) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol外调价) * Val(.TextMatrix(intRow, mconIntCol数量)), mintMoneyDigit, , True)
                End If
                If cboType.Text = "药品外调" Or cboType.Text = "药品外销" Then
                    dblOutPrice = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol外调价)) / .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_成本价)
                End If
                dblMistakePrice = Val(zlStr.FormatEx(dblSaleMoney - dblPurchaseMoney, mintMoneyDigit, , True)) '.TextMatrix(intRow, mconintCol差价)
                
                dbl增值税率 = Val(.TextMatrix(intRow, mconIntCol增值税率))
                
'                If Val(.TextMatrix(intRow, mconIntCol序号)) = 0 Then
'                    lngSerial = intRow
'                Else
'                    lngSerial = Val(.TextMatrix(intRow, mconIntCol序号))
'                End If
                lngSerial = intRow
                
                gstrSQL = "zl_药品其他出库_INSERT("
                '入出类别ID
                gstrSQL = gstrSQL & lng入出类别id
                'NO
                gstrSQL = gstrSQL & ",'" & chrNo & "'"
                '序号
                gstrSQL = gstrSQL & "," & lngSerial
                '库房ID
                gstrSQL = gstrSQL & "," & lngStockid
                '药品ID
                gstrSQL = gstrSQL & "," & lngDrugID
                '批次
                gstrSQL = gstrSQL & "," & lngBatchID
                '填写数量
                gstrSQL = gstrSQL & "," & dblQuantity
                '成本价
                gstrSQL = gstrSQL & "," & dblPurchasePrice
                '成本金额
                gstrSQL = gstrSQL & "," & dblPurchaseMoney
                '零售价
                gstrSQL = gstrSQL & "," & dblSalePrice
                '零售金额
                gstrSQL = gstrSQL & "," & dblSaleMoney
                '差价
                gstrSQL = gstrSQL & "," & dblMistakePrice
                '外调价(外销价)
                gstrSQL = gstrSQL & "," & dblOutPrice
                '外调单位(外销单位)
                gstrSQL = gstrSQL & ",'" & strOutUnit & "'"
                '填制人
                gstrSQL = gstrSQL & ",'" & strBooker & "'"
                '填制日期
                gstrSQL = gstrSQL & ",to_date('" & datBookDate & "','yyyy-mm-dd HH24:MI:SS')"
                '产地
                gstrSQL = gstrSQL & ",'" & strProducingArea & "'"
                '批号
                gstrSQL = gstrSQL & ",'" & strBatchNo & "'"
                '效期
                gstrSQL = gstrSQL & "," & IIf(datTimeLimit = "", "Null", "to_date('" & Format(datTimeLimit, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                '摘要
                gstrSQL = gstrSQL & ",'" & strBrief & "'"
                '批准文号
                gstrSQL = gstrSQL & ",'" & str批准文号 & "'"
                '增值税率
                gstrSQL = gstrSQL & "," & dbl增值税率
                '原产地
                gstrSQL = gstrSQL & ",'" & strOldProducingArea & "'"
                '修改人
                gstrSQL = gstrSQL & ",'" & strModifier & "'"
                '修改日期
                gstrSQL = gstrSQL & "," & IIf(datModifyDate = "", "Null", "to_date('" & datModifyDate & "','yyyy-mm-dd HH24:MI:SS')")
                gstrSQL = gstrSQL & ")"
                    
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = CStr(lngDrugID) & ";" & gstrSQL
            End If
            recSort.MoveNext
        Next
        
        If Not ExecuteSql(arrSql, MStrCaption, False, Not bln强制保存) Then Exit Function
        If Not bln强制保存 Then gcnOracle.CommitTrans: blnTran = False
        
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveCard = True
    Exit Function
errHandle:
    If Not bln强制保存 Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub 显示合计金额()
    Dim curTotal As Double, Cur记帐金额 As Double, Cur记帐差价 As Double, Cur外调金额 As Double
    Dim intLop As Integer
    
    curTotal = 0: Cur记帐金额 = 0: Cur记帐差价 = 0:
    
    With mshBill
        For intLop = 1 To .rows - 1
            curTotal = curTotal + Val(.TextMatrix(intLop, mconIntCol采购金额))
            Cur记帐金额 = Cur记帐金额 + Val(.TextMatrix(intLop, mconIntCol售价金额))
            Cur外调金额 = Cur外调金额 + Val(.TextMatrix(intLop, mconIntCol外调金额))
        Next
    End With
    
    Cur记帐差价 = Cur记帐金额 - curTotal
    lblPurchasePrice.Caption = "成本金额合计：" & zlStr.FormatEx(curTotal, mintMoneyDigit, , True)
    lblSalePrice.Caption = "售价金额合计：" & zlStr.FormatEx(Cur记帐金额, mintMoneyDigit, , True)
    lblDifference.Caption = "差价合计：" & zlStr.FormatEx(Cur记帐差价, mintMoneyDigit, , True)
    lblOther.Caption = "外调(销)合计：" & zlStr.FormatEx(Cur外调金额, mintMoneyDigit, , True)
End Sub

Private Sub 提示库存数()
    Dim rsUseCount As New Recordset
    
    On Error GoTo errHandle
    With mshBill
        If .TextMatrix(.Row, mconIntCol药名) = "" Then
            staThis.Panels(2).Text = ""
            Exit Sub
        End If
        If .TextMatrix(mshBill.Row, 0) = "" Then Exit Sub
        
        gstrSQL = "select 可用数量/" & .TextMatrix(.Row, mconIntCol比例系数) & " as  可用数量   from 药品库存 where 库房id=[1] " _
            & " and 药品id=[2] " _
            & " and 性质=1 and " _
            & " nvl(批次,0)=[3]"
        Set rsUseCount = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[提示库存数]", cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol批次)))
            
        If rsUseCount.EOF Then
            .TextMatrix(.Row, mconIntCol可用数量) = 0
        Else
            .TextMatrix(.Row, mconIntCol可用数量) = IIf(IsNull(rsUseCount.Fields(0)), 0, rsUseCount.Fields(0))
        End If
        rsUseCount.Close
        
        staThis.Panels(2).Text = "该药品当前库存数为[" & zlStr.FormatEx(.TextMatrix(.Row, mconIntCol可用数量), mintNumberDigit, , True) & "]" & .TextMatrix(.Row, mconIntCol单位)
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtIn_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim IntCheck As Integer
    Dim intRow As Integer
    Dim blnEXIST As Boolean
    Dim intIndex As Integer, intCount As Integer
    Dim rsBill As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim lng库房ID As Long
    Dim intNO As Integer, strNo As String
    On Error GoTo ErrHand
    
    '初始准备
    intNO = 28
    lng库房ID = Me.cboStock.ItemData(Me.cboStock.ListIndex)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txtIn.Text) = "" Then Exit Sub
    
    If Len(txtIn.Text) < 8 Then
        txtIn.Text = zlCommFun.GetFullNO(txtIn.Text, intNO, lng库房ID)
    End If
    
    '设置入出类别为药品外调
    intCount = cboType.ListCount
    For intIndex = 1 To intCount
        If cboType.List(intIndex - 1) = "药品外调" Then
            cboType.ListIndex = intIndex - 1
            blnEXIST = True
            Exit For
        End If
    Next
'    If Not blnEXIST Then
'        MsgBox "导入外购入库单的功能只能应用于入出类别“药品外调”！", vbInformation, gstrSysName
'        Exit Sub
'    End If
    
    '需要要清除现有单据内容
    For IntCheck = 1 To mshBill.rows - 1
        If mshBill.TextMatrix(IntCheck, 0) <> "" Then
            Exit For
        End If
    Next
    If IntCheck <> mshBill.rows Then
        If MsgBox("需要要清除现有单据内容，你是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        '处理药品单位改变
        mshBill.ClearBill
    End If
    
    '取出库检查性质
    IntCheck = 0
    gstrSQL = "Select Nvl(检查方式,0) 库存检查 From 药品出库检查 Where 库房ID=[1] "
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[取出库检查性质]", Me.cboStock.ItemData(Me.cboStock.ListIndex))

    If rsTemp.RecordCount <> 0 Then IntCheck = rsTemp!库存检查
    
    '提取该单据并清空表格（只允许提取正常单据，且非退货单）
    gstrSQL = "SELECT A.药品ID,'['||C.编码||']' As 编码,'['||C.编码||']'|| Nvl(F.名称,C.名称) As 药品名称, C.名称 As 通用名,F.名称 As 商品名,C.规格,C.产地,A.原产地," & _
             "        C.计算单位 AS 零售单位,1 AS 零售系数,B.门诊单位,B.门诊包装,B.住院单位,B.住院包装,B.药库单位,B.药库包装, " & _
             "        NVL(A.批次,0) AS 批次,Nvl(C.是否变价,0) AS 时价,Nvl(B.药房分批,0) AS 药房分批,A.批号,A.效期," & _
             "        B.管理费比例,B.加成率,A.实际数量,D.可用数量,D.实际金额,D.实际差价,E.现价,A.批准文号,B.药品来源,B.基本药物,d.平均成本价 " & _
             " FROM 药品收发记录 A,药品规格 B,收费项目目录 C,药品库存 D,收费价目 E,收费项目别名 F " & _
             " WHERE A.药品ID=B.药品ID AND B.药品ID=C.ID AND B.药品ID=D.药品ID(+) " & _
             " AND B.药品ID=F.收费细目ID(+) AND F.性质(+)=3 AND F.码类(+)=1" & _
             " AND B.药品ID=E.收费细目ID(+) AND SYSDATE >=E.执行日期(+)  AND sysdate<=NVL(E.终止日期(+),SYSDATE)" & _
             GetPriceClassString("E") & _
             " AND D.库房ID(+)=[2] AND D.性质(+)=1 AND Nvl(A.批次,0)=Nvl(D.批次,0)" & _
             " AND A.单据=1 AND A.记录状态=1 AND NVL(A.发药方式,0)=0 AND A.审核日期 Is Not NULL" & _
             " AND A.NO=[1] And A.库房ID+0=[2] " & _
             " ORDER BY A.序号"
    Set rsBill = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[提取外购入库单]", txtIn.Text, Me.cboStock.ItemData(Me.cboStock.ListIndex))
             
    If rsBill.RecordCount = 0 Then
        MsgBox "没有找到该外购入库单据！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    With rsBill
        intRow = 1
        Do While Not .EOF
            '装入数据前，先检查库存
            If !实际数量 > !可用数量 Then
                '批次或时价药品不允许零出库
                If !批次 <> 0 Or !时价 <> 0 Then
                    MsgBox !药品名称 & "库存不足，不允许出库！（时价或分批药品）", vbInformation, gstrSysName
                    mshBill.ClearBill
                    Exit Sub
                End If
                Select Case IntCheck
                Case 1
                    If MsgBox(!药品名称 & "已经没有库存，是否继续！", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        mshBill.ClearBill
                        Exit Sub
                    End If
                Case 2
                    MsgBox !药品名称 & "已经没有库存，不能继续操作！", vbInformation, gstrSysName
                    mshBill.ClearBill
                    Exit Sub
                End Select
            End If
            
            '装入数据(SetColValue)
            If Not SetColValue(intRow, !药品ID, !编码, !通用名, IIf(IsNull(!商品名), "", !商品名), _
                nvl(!药品来源), nvl(!基本药物), nvl(!规格), nvl(!产地), _
                Choose(mintUnit, !零售单位, !门诊单位, !住院单位, !药库单位), nvl(!现价, 0), _
                nvl(!批号), nvl(!效期), nvl(!可用数量, 0), nvl(!实际金额, 0), nvl(!实际差价, 0), _
                nvl(!加成率 / 100, 0), Choose(mintUnit, 1, !门诊包装, !住院包装, !药库包装), nvl(!批次, 0), !时价, _
                !药房分批, IIf(IsNull(!批准文号), "", !批准文号), nvl(!原产地)) Then
                mshBill.ClearBill
                Exit Sub
            End If
            
            '填写数量、采购价等列
            mshBill.TextMatrix(intRow, mconIntCol行号) = intRow
            mshBill.TextMatrix(intRow, mconIntCol数量) = zlStr.FormatEx(nvl(!实际数量, 0) / Choose(mintUnit, 1, !门诊包装, !住院包装, !药库包装), mintNumberDigit, , True)
            If mshBill.TextMatrix(intRow, mconIntCol售价) <> "" Then
                mshBill.TextMatrix(intRow, mconIntCol售价金额) = zlStr.FormatEx(Val(mshBill.TextMatrix(intRow, mconIntCol售价)) * Val(mshBill.TextMatrix(intRow, mconIntCol数量)), mintMoneyDigit, , True)
            End If
            
'            mshBill.TextMatrix(intRow, mconintCol差价) =Str.FormatEx(Get出库差价(Val(cboStock.ItemData(cboStock.ListIndex)), Val(mshBill.TextMatrix(intRow, 0)), Val(mshBill.TextMatrix(intRow, mconIntCol批次)), Val(mshBill.TextMatrix(intRow, mconIntCol实际金额)), Val(mshBill.TextMatrix(intRow, mconIntCol实际差价)), Val(mshBill.TextMatrix(intRow, mconIntCol售价金额)), Val(mshBill.TextMatrix(intRow, mconIntCol数量)) * Val(mshBill.TextMatrix(intRow, mconIntCol比例系数))), mintMoneyDigit)
            
            If nvl(!实际数量, 0) <> 0 Then
                mshBill.TextMatrix(intRow, mconIntCol采购价) = zlStr.FormatEx(Get成本价(Val(mshBill.TextMatrix(intRow, 0)), Val(cboStock.ItemData(cboStock.ListIndex)), Val(mshBill.TextMatrix(intRow, mconIntCol批次))) * Val(mshBill.TextMatrix(intRow, mconIntCol比例系数)), mintCostDigit, , True)
'                mshBill.TextMatrix(intRow, mconIntCol采购价) =Str.FormatEx((mshBill.TextMatrix(intRow, mconIntCol售价金额) - mshBill.TextMatrix(intRow, mconintCol差价)) / Val(mshBill.TextMatrix(intRow, mconIntCol数量)), mintCostDigit)
            End If
            mshBill.TextMatrix(intRow, mconIntCol采购金额) = zlStr.FormatEx(Val(mshBill.TextMatrix(intRow, mconIntCol采购价)) * Val(mshBill.TextMatrix(intRow, mconIntCol数量)), mintMoneyDigit, , True)
            mshBill.TextMatrix(intRow, mconintCol差价) = zlStr.FormatEx(Val(mshBill.TextMatrix(intRow, mconIntCol售价金额)) - Val(mshBill.TextMatrix(intRow, mconIntCol采购金额)), mintMoneyDigit, , True)
            
            '计算外调价及外调金额:外调价=(1+管理费比例)*进价
            mshBill.TextMatrix(intRow, mconIntCol外调价) = zlStr.FormatEx((1 + !管理费比例 / 100) * Val(mshBill.TextMatrix(intRow, mconIntCol采购价)), mintCostDigit, , True)
            mshBill.TextMatrix(intRow, mconIntCol外调金额) = zlStr.FormatEx(Val(mshBill.TextMatrix(intRow, mconIntCol外调价)) * Val(mshBill.TextMatrix(intRow, mconIntCol数量)), mintMoneyDigit, , True)
            
            intRow = intRow + 1
            mshBill.rows = mshBill.rows + 1
            .MoveNext
        Loop
    End With
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    mshBill.ClearBill
End Sub

Private Sub txt摘要_Change()
    mblnChange = True
End Sub

Private Sub txt摘要_GotFocus()
    OS.OpenIme True
    With txt摘要
        .SelStart = 0
        .SelLength = Len(txt摘要.Text)
    End With
End Sub

Private Sub txt摘要_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey (vbKeyTab)
        KeyCode = 0
    End If
End Sub

Private Sub txt摘要_LostFocus()
    OS.OpenIme
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

'打印单据
Private Sub printbill()
    Dim strUnit As String
    Dim int单位系数 As Integer
    Dim strNo As String
    
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
    strNo = txtNo.Tag
    FrmBillPrint.ShowME Me, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1306", "zl8_bill_1306"), mint记录状态, int单位系数, 1306, "药品其它出库单", strNo
End Sub

Private Sub SetForeColor_ROW(ByVal lngColor As Long)
    Dim i As Integer, j As Integer
    Dim intCol As Integer
    '设置某行的颜色
    With mshBill
        intCol = .Col
        mblnEnterCell = False
        For i = mconIntCol药名 To .Cols - 1
            j = .ColData(i)
            If .ColData(i) = 5 Then .ColData(i) = 0
            .Col = i
            .MsfObj.CellForeColor = lngColor
            .ColData(i) = j
        Next
        .Col = intCol
        mblnEnterCell = True
    End With
End Sub

Private Function CheckData(ByVal rsTemp As ADODB.Recordset) As ADODB.Recordset
    '功能：用来检查列表中已有药品与新选择的药品是否重复和时价药品是否有库存

    Dim i As Integer
    Dim strTemp As String
    Dim str批次 As String
    Dim strInfo As String
    Dim rsPrice As ADODB.Recordset
    Dim str库存 As String
    Dim strSQL As String
    Dim strDub As String    '重复药品
    Dim strNotNum As String  '无库存药品
    Dim str重复药名 As String   '用来记录重复选择了的药品名称
    Dim strNot药名 As String    '用来记录哪些药品是时价但无库存
    
    On Error GoTo errHandle

    rsTemp.MoveFirst
    str批次 = ""
    strTemp = ""
    Do While Not rsTemp.EOF
        str批次 = IIf(IsNull(rsTemp!批次), "0", rsTemp!批次)
        If InStr(1, strTemp, rsTemp!药品ID & "," & str批次) = 0 Then
            strTemp = strTemp & rsTemp!药品ID & "," & str批次 & "," & rsTemp!通用名 & "|"
        End If
    
        rsTemp.MoveNext
    Loop
        
    With mshBill    '把重复的查询出来
        For i = 1 To .rows - 2
            If InStr(1, strTemp, .TextMatrix(i, 0) & "," & .TextMatrix(i, mconIntCol批次)) > 0 Then
                strInfo = strInfo & .TextMatrix(i, 0) & "," & .TextMatrix(i, mconIntCol药名) & "|"
            End If
        Next
        
        If strInfo <> "" Then   '为过滤数据拼接sql
            strDub = ""
            For i = 0 To UBound(Split(strInfo, "|")) - 1
                strDub = strDub & "药品id<>" & Split(Split(strInfo, "|")(i), ",")(0) & " and "
                If UBound(Split(str重复药名, ",")) <= 2 Then
                    str重复药名 = str重复药名 & Split(Split(strInfo, "|")(i), ",")(1) & ","
                End If
            Next
            If strDub <> "" Then
                strDub = Mid(strDub, 1, Len(strDub) - 4)
            End If
        End If
        '判断以什么方式拼接sql
        If str重复药名 <> "" Then
            MsgBox str重复药名 & "列表中已经含有了！" & vbCrLf & "以上药品不再添加！", vbInformation, gstrSysName
            strSQL = strDub
        End If
        If strSQL <> "" Then
            rsTemp.Filter = strSQL
        End If
        
        Set CheckData = rsTemp
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetPrice(ByVal lng药品ID As Long, ByVal lng批次 As Long, ByVal dbl比例系数 As Double) As Double
    Dim rsPrice As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "select Decode(Nvl(批次,0),0,实际金额/实际数量,Nvl(零售价,实际金额/实际数量))*" & dbl比例系数 & " as  售价 " _
        & "  from 药品库存 " _
        & " where 库房id=[1] " _
        & " and 药品id=[2] " _
        & " and 性质=1 and 实际数量>0 and " _
        & " nvl(批次,0)=[3]"
    Set rsPrice = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, cboStock.ItemData(cboStock.ListIndex), lng药品ID, lng批次)

    If rsPrice.EOF Then
        gstrSQL = "Select 现价 From 收费价目 Where 收费细目id = [1] And Sysdate Between 执行日期 And 终止日期" & _
                GetPriceClassString("")
        
        Set rsPrice = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lng药品ID)
        If rsPrice.RecordCount > 0 Then
            GetPrice = rsPrice!现价 * dbl比例系数
        End If
        Exit Function
    End If
    GetPrice = rsPrice.Fields(0).Value
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

Private Function 检查价格() As Boolean
    '功能：新增时，判断药品是否是最新价格，不是则修改后提示
    Dim strMsg As String '保存提示信息
    Dim i As Integer, intSum As Integer, intPriceDigit As Integer
    Dim rsPrice As New ADODB.Recordset
    Dim Dbl数量 As Double, dbl成本价 As Double, dbl成本金额 As Double, dbl零售价 As Double, dbl零售金额 As Double, dbl差价 As Double
    Dim bln是否时价 As Boolean
    
    On Error GoTo errHandle
    
    检查价格 = False
    
    With mshBill
        For i = 1 To .rows - 1
            If mshBill.TextMatrix(i, 0) <> "" And Trim(.TextMatrix(i, mconIntCol数量)) <> "" Then
            
                bln是否时价 = Val(Split(.TextMatrix(i, mconIntcol加成率), "||")(1)) = 1
                Dbl数量 = Val(.TextMatrix(i, mconIntCol数量))
                
                '检查成本价
                dbl成本价 = zlStr.FormatEx(Get成本价(Val(.TextMatrix(i, 0)), cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(i, mconIntCol批次))) * Val(.TextMatrix(i, mconIntCol比例系数)), mintCostDigit)
                If .TextMatrix(i, mconIntCol采购价) <> dbl成本价 Then
                    intSum = intSum + 1
                    .TextMatrix(i, mconIntCol采购价) = zlStr.FormatEx(dbl成本价, mintCostDigit, , True)
                    .TextMatrix(i, mconIntCol采购金额) = zlStr.FormatEx(.TextMatrix(i, mconIntCol采购价) * Dbl数量, mintMoneyDigit, , True)
                End If
                
                '检查售价
                dbl零售价 = zlStr.FormatEx(Get售价(bln是否时价, Val(.TextMatrix(i, 0)), cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(i, mconIntCol批次))) * Val(.TextMatrix(i, mconIntCol比例系数)), mintPriceDigit)
                If .TextMatrix(i, mconIntCol售价) <> dbl零售价 Then
                    intSum = intSum + 1
                    .TextMatrix(i, mconIntCol售价) = zlStr.FormatEx(dbl零售价, mintPriceDigit, , True)
                    .TextMatrix(i, mconIntCol售价金额) = zlStr.FormatEx(.TextMatrix(i, mconIntCol售价) * Dbl数量, mintMoneyDigit, , True)
                End If
                
                .TextMatrix(i, mconintCol差价) = zlStr.FormatEx(Val(.TextMatrix(i, mconIntCol售价金额)) - Val(.TextMatrix(i, mconIntCol采购金额)), mintMoneyDigit, , True)
                
            End If
        Next
        
        If intSum > 0 Then
            MsgBox "有记录未使用最新价格，程序已自动完成更新（成本价、成本金额、售价、售价金额、差价），更新后请检查！", vbInformation, gstrSysName
            检查价格 = True
        End If
    End With
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

