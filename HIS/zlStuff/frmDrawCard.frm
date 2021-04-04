VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmDrawCard 
   Caption         =   "卫材领用单"
   ClientHeight    =   6975
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "frmDrawCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   11400
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdExpend 
      Caption         =   "自动分解(&A)"
      Height          =   350
      Left            =   1680
      TabIndex        =   44
      Top             =   6000
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton cmdRequestDraw 
      Caption         =   "按申购单领用(&R)"
      Height          =   350
      Left            =   1800
      TabIndex        =   41
      Top             =   5520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "跟踪(&G)"
      Height          =   360
      Left            =   270
      TabIndex        =   28
      Top             =   5940
      Width           =   810
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshProvider 
      Height          =   1812
      Left            =   1092
      TabIndex        =   27
      Top             =   6840
      Visible         =   0   'False
      Width           =   4092
      _ExtentX        =   7223
      _ExtentY        =   3201
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   32768
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdAllCls 
      Caption         =   "全清(&L)"
      Height          =   350
      Left            =   9720
      TabIndex        =   25
      Top             =   5940
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdAllSel 
      Caption         =   "全冲(&A)"
      Height          =   350
      Left            =   8400
      TabIndex        =   24
      Top             =   5940
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   5520
      TabIndex        =   15
      Top             =   5610
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "查找(&F)"
      Height          =   350
      Left            =   3840
      TabIndex        =   14
      Top             =   5520
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   13
      Top             =   5520
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   8400
      TabIndex        =   11
      Top             =   5520
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   9720
      TabIndex        =   12
      Top             =   5520
      Width           =   1100
   End
   Begin VB.PictureBox Pic单据 
      BackColor       =   &H80000004&
      Height          =   5205
      Left            =   0
      ScaleHeight     =   5145
      ScaleWidth      =   11655
      TabIndex        =   16
      Top             =   0
      Width           =   11715
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
         Height          =   360
         Left            =   960
         MaxLength       =   8
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   135
         Visible         =   0   'False
         Width           =   1515
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
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   42
         TabStop         =   0   'False
         ToolTipText     =   "导入记帐单:F3"
         Top             =   150
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtDrawPerson 
         Height          =   300
         Left            =   9660
         TabIndex        =   6
         Top             =   585
         Width           =   1425
      End
      Begin VB.CommandButton cmdDrawPerson 
         Caption         =   "…"
         Height          =   300
         Left            =   11100
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   555
         Width           =   300
      End
      Begin VB.TextBox txtNO 
         Height          =   300
         IMEMode         =   2  'OFF
         Left            =   9945
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   165
         Width           =   1425
      End
      Begin VB.TextBox txtDraw 
         Height          =   300
         Left            =   5715
         TabIndex        =   3
         Top             =   600
         Width           =   2415
      End
      Begin VB.CommandButton cmdDraw 
         Caption         =   "…"
         Height          =   300
         Left            =   8100
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   600
         Width           =   300
      End
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   2790
         Left            =   195
         TabIndex        =   8
         Top             =   945
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   4921
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
         TabIndex        =   10
         Top             =   4080
         Width           =   10410
      End
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   945
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   1905
      End
      Begin VB.Label Lbl审核日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审核日期"
         Height          =   180
         Left            =   8640
         TabIndex        =   40
         Top             =   4860
         Width           =   720
      End
      Begin VB.Label lbl审核人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  审核人"
         Height          =   180
         Left            =   8640
         TabIndex        =   39
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label Lbl填制日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "填制日期"
         Height          =   180
         Left            =   120
         TabIndex        =   38
         Top             =   4860
         Width           =   720
      End
      Begin VB.Label Lbl填制人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  填制人"
         Height          =   180
         Left            =   120
         TabIndex        =   37
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label Txt填制人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   36
         Top             =   4440
         Width           =   1890
      End
      Begin VB.Label Txt填制日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   35
         Top             =   4800
         Width           =   1890
      End
      Begin VB.Label Txt审核日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   9480
         TabIndex        =   34
         Top             =   4800
         Width           =   1890
      End
      Begin VB.Label Txt审核人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   9480
         TabIndex        =   33
         Top             =   4455
         Width           =   1890
      End
      Begin VB.Label txt核查人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5085
         TabIndex        =   32
         Top             =   4455
         Width           =   1890
      End
      Begin VB.Label txt核查日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5085
         TabIndex        =   31
         Top             =   4800
         Width           =   1890
      End
      Begin VB.Label lbl核查人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  核查人"
         Height          =   180
         Left            =   4320
         TabIndex        =   30
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label lbl核查日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "核查日期"
         Height          =   180
         Left            =   4275
         TabIndex        =   29
         Top             =   4860
         Width           =   720
      End
      Begin VB.Label lbl领用人 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "领用人(&L)"
         Height          =   180
         Left            =   8790
         TabIndex        =   5
         Top             =   645
         Width           =   825
      End
      Begin VB.Label lblDifference 
         AutoSize        =   -1  'True
         Caption         =   "差价合计:"
         Height          =   180
         Left            =   4920
         TabIndex        =   22
         Top             =   3840
         Width           =   810
      End
      Begin VB.Label lblSalePrice 
         AutoSize        =   -1  'True
         Caption         =   "售价金额合计:"
         Height          =   180
         Left            =   2040
         TabIndex        =   21
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "成本金额合计:"
         Height          =   180
         Left            =   240
         TabIndex        =   20
         Top             =   3840
         Width           =   1170
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
         TabIndex        =   18
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
         TabIndex        =   9
         Top             =   4155
         Width           =   650
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "卫生材料领用单"
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
         TabIndex        =   17
         Top             =   135
         Width           =   11535
      End
      Begin VB.Label LblStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "库房(&S)"
         Height          =   180
         Left            =   225
         TabIndex        =   0
         Top             =   660
         Width           =   630
      End
      Begin VB.Label LblEnterStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "领料部门(&D)"
         Height          =   180
         Left            =   4635
         TabIndex        =   2
         Top             =   660
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
            Picture         =   "frmDrawCard.frx":014A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":0364
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":057E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":0798
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":09B2
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":0BCC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":0DE6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":1000
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
            Picture         =   "frmDrawCard.frx":121A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":1434
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":164E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":1868
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":1A82
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":1C9C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":1EB6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":20D0
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   23
      Top             =   6615
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDrawCard.frx":22EA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13758
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmDrawCard.frx":2B7E
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmDrawCard.frx":3080
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
      Caption         =   "材料"
      Height          =   255
      Left            =   5040
      TabIndex        =   19
      Top             =   5640
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmDrawCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbln单据增加    As Boolean          '进入时单据号累加1
Private mstr入库单号 As String              '入库单号

Private mintUnit  As Integer                '显示单位:0-散装单位,1-包装单位
Private mblnFirst As Boolean
Private mint编辑状态 As Integer             '1－新增；2－修改；3－审核；4－查看；5－财务审核；6－冲销；7-从入库单读取数据

Private mstr单据号 As String                '具体的单据号;
Private mint记录状态 As Integer             '1:正常记录;2-冲销记录;3-已经冲销的原记录
Private mblnSuccess As Boolean              '只要有一张成功，即为True，否则为False
Private mblnSave As Boolean                 '是否存盘和审核   TURE：成功。
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEdit As Boolean                 '是否可以修改
Private mblnChange As Boolean               '是否进行过编辑
Private mintParallelRecord As Integer       '对于新增后单据并行执行的处理： 1、代表正常情况；2、已经删除的记录；3、已经审核的记录
Private mint库存检查 As Integer             '表示卫生材料出库时是否进行库存检查：0-不检查;1-检查，不足提醒；2-检查，不足禁止
Private mcolUsedCount As Collection         '已使用的数量集合
Private mstrPrivs As String                     '权限
Private Const mstrCaption As String = "卫材领用单"
Private mint领用明确批次 As Integer         '0-领用不明确批次 1-领用明确批次

Private mlng部门ID As Long          '从入库单读取数据时有效
Private mstr领用人 As String        '从入库单读取数据时有效
 '刘兴宏:2007/06/10:问题10813
Private mstrTime_Start As String            '进入单据编辑的单据时间 ,主要判断是否单据被他人更改过,如果编辑过,则不能进行审核
Private mstrTime_End As String
Private mblnEnter As Boolean        '不移动列
Private Const mlngModule = 1717
Private mblnCostView As Boolean                 '查看成本价 true-允许查看 false-不允许查看
 
'刘兴宏:20060803加入领用部门申领
Private mbln普通科室 As Boolean
Private mblnHave领用用途 As Boolean '确定是否初始化了材料领用用途的,如果初始,则提供选择器选择,否则自由录入.
Private mblnUpdate As Boolean               '表示是否已根据最新价格更新单据内容
Private mstrRequestNO As String     '按申购单领用NO ，空代表不按照申购单方式领用，否则按照申购单领用
Private mstr重复卫材 As String '记录重复的卫材

Private recSort As ADODB.Recordset          '按药品ID排序的专用记录集

'----------------------------------------------------------------------------------------------------------
'刘兴宏:增加小数位数的格式串
'修改:2007/03/06
Private mFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------
'=========================================================================================
Private Type POINTAPI
     x As Long
     y As Long
End Type
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Private Enum mBillCol
        C_材料ID = 0
        C_行号 = 1
        C_材料 = 2
        C_序号 = 3
        c_规格 = 4
        C_可用数量 = 5
        C_指导差价率 = 6
        C_实际金额 = 7
        C_实际差价 = 8
        c_比例系数 = 9
        c_批次 = 10
        C_产地 = 11
        C_批准文号 = 12
        c_单位 = 13
        c_批号 = 14
        C_效期 = 15
        C_灭菌失效期 = 16
        C_申购数量 = 17
        C_填写数量 = 18
        C_实际数量 = 19
        c_原始数量 = 20
        C_采购价 = 21
        C_采购金额 = 22
        C_售价 = 23
        C_售价金额 = 24
        C_差价 = 25
        C_跟踪标志 = 26
        C_跟踪信息 = 27 '病人ID|使用时间|条码
        C_跟踪病人 = 28
        C_分批属性 = 29
        C_库存数量 = 30 '冲销时才用，主要针对负数入库后的冲销
End Enum
Private mstr默认材料用途 As String
Private Const mBillCols As Integer = 31              '总列数


'=========================================================================================


'检查数据依赖性
Private Function GetDepend() As Boolean
    Dim rsTemp As New Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    GetDepend = False
    strSQL = "" & _
        "   SELECT B.Id " & _
        "   FROM 药品单据性质 A, 药品入出类别 B " & _
        "   Where A.类别id = B.ID AND A.单据 = 35"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "卫材领用单")
    If rsTemp.EOF Then
        ShowMsgBox "没有设置卫材领用的出库类别，请在入出分类中设置！"
        rsTemp.Close
        Exit Function
    End If
    
    strSQL = "" & _
        "   SELECT DISTINCT a.id, a.名称 " & _
        "   FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " & _
        "   Where c.工作性质 = b.名称   AND a.id = c.部门id and (a.站点=[1] or a.站点 is null) " & _
        "       AND ( TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01' or a.撤档时间 is null)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, mstrCaption, gstrNodeNo)
    If rsTemp.EOF Then
        MsgBox "部门体系不全,请在部门管理中设置！", vbInformation, gstrSysName
        rsTemp.Close
        Exit Function
    End If
    
    strSQL = "Select 编码 From 材料领用用途 where rownum<=1 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "材料领用用途-编码")
    mblnHave领用用途 = rsTemp.EOF = False
    
    strSQL = "Select 名称 From 材料领用用途 where nvl(缺省标志,0)=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "材料领用用途-名称")
    If rsTemp.EOF = False Then
        mstr默认材料用途 = zlStr.Nvl(rsTemp!名称)
    Else
        mstr默认材料用途 = ""
    End If
    
    GetDepend = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Sub ShowCard(frmMain As Form, ByVal str单据号 As String, ByVal int编辑状态 As Integer, _
    Optional int记录状态 As Integer = 1, Optional strPrivs As String, Optional blnSuccess As Boolean = False, _
    Optional lng领用部门id As Long = 0, Optional str领用人 As String = "")
    '-----------------------------------------------------------------------------------------------------------
    '功能:单据入口
    '入参:frmMain-调用的主窗口
    '    str单据号-单据号(对于编辑类型为7<从入库单读取数据>,表示入库单号,否则代表领用单据号
    '    int编辑状态-编辑类型:1.新增；2、修改；3、验收；4、查看； 6-冲销单据,7-从入库单读取数据
    '    int记录状态-记录状态
    '    strPrivs-权限串
    '    lng部门id-传入的领用部门ID(编辑类型=7有效)
    '    str领用人-传入的领用人(编辑类型=7有效)
    '出参:blnSuccess-返回成功标块,true,表示至少有一张单据保存成功,否则表示无一张单据保存成功
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-10-27 11:45:51
    '-----------------------------------------------------------------------------------------------------------
    
    Dim strReg As String
    mlng部门ID = lng领用部门id: mstr领用人 = str领用人
    
    mblnSave = False: mblnSuccess = False
    mstr入库单号 = "": mstr单据号 = ""
    
    If int编辑状态 = 7 Then
        mstr入库单号 = str单据号
    Else
        mstr单据号 = str单据号
    End If
    
    mint编辑状态 = int编辑状态
    mint记录状态 = int记录状态
    mblnSuccess = blnSuccess
    mblnChange = False
    mintParallelRecord = 1
    mstrPrivs = strPrivs
    mblnCostView = zlStr.IsHavePrivs(mstrPrivs, "查看成本价")
    
    Set mfrmMain = frmMain
    If Not GetDepend Then Exit Sub
    
    Call GetRegInFor(g私有模块, "卫材领用管理", "单据号累加", strReg)
    mbln单据增加 = IIf(strReg = "", True, Val(strReg) = 1)
    
   
     
    If mint编辑状态 = 1 Or mint编辑状态 = 7 Then
'        If mbln单据增加 Then
'            mstr单据号 = NextNo(73)
'        End If
        mblnEdit = True

        txtNo.Locked = True
        txtNo.TabStop = True

        txtNo = mstr单据号
        txtNo.Tag = txtNo.Text
    ElseIf mint编辑状态 = 2 Then
        mblnEdit = True
    ElseIf mint编辑状态 = 3 Then
        mblnEdit = False
        CmdSave.Caption = "审核(&V)"
    ElseIf mint编辑状态 = 5 Then
        mblnEdit = False
        CmdSave.Caption = "审核(&V)"
    ElseIf mint编辑状态 = 4 Then
        mblnEdit = False
        CmdSave.Caption = "打印(&P)"
        If InStr(mstrPrivs, "单据打印") = 0 Then
            CmdSave.Visible = False
        Else
            CmdSave.Visible = True
        End If
    ElseIf mint编辑状态 = 6 Then
        CmdSave.Caption = "冲销(&O)"
        cmdAllCls.Visible = True
        cmdAllSel.Visible = True
    End If
      
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    Me.Show vbModal, frmMain
    blnSuccess = mblnSuccess
    str单据号 = mstr单据号
End Sub

Private Sub cboStock_Change()
    mblnChange = True
End Sub
Private Sub cboStock_Click()
    mint库存检查 = Get出库检查(cboStock.ItemData(cboStock.ListIndex))
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
            For i = 1 To mshBill.Rows - 1
                If mshBill.TextMatrix(i, 0) <> "" Then
                    Exit For
                End If
            Next
            If i <> mshBill.Rows Then
                If MsgBox("如果改变库房，有可能要改变相应卫材的单位，" & vbCrLf & "且要清除现有单据内容，你是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    '处理卫材单位改变
                    mintcboIndex = .ListIndex
                    mshBill.ClearBill
                            
                Else
                    .ListIndex = mintcboIndex
                End If
            Else
                mintcboIndex = .ListIndex
            End If
        End If
        
    End With
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
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                .TextMatrix(intRow, mBillCol.C_实际数量) = Format(0, mFMT.FM_数量)
                .TextMatrix(intRow, mBillCol.C_采购金额) = Format(0, mFMT.FM_金额)
                .TextMatrix(intRow, mBillCol.C_售价金额) = Format(0, mFMT.FM_金额)
                .TextMatrix(intRow, mBillCol.C_差价) = Format(0, mFMT.FM_金额)
            End If
        Next
    End With
    Call 显示合计金额
End Sub

Private Sub cmdAllSel_Click()
    Dim intRow As Integer
    
    With mshBill
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                .TextMatrix(intRow, mBillCol.C_实际数量) = .TextMatrix(intRow, mBillCol.C_填写数量)
                .TextMatrix(intRow, mBillCol.C_采购金额) = Format(.TextMatrix(intRow, mBillCol.C_填写数量) * .TextMatrix(intRow, mBillCol.C_采购价), mFMT.FM_金额)
                .TextMatrix(intRow, mBillCol.C_售价金额) = Format(.TextMatrix(intRow, mBillCol.C_填写数量) * .TextMatrix(intRow, mBillCol.C_售价), mFMT.FM_金额)
                .TextMatrix(intRow, mBillCol.C_差价) = Format(.TextMatrix(intRow, mBillCol.C_售价金额) - .TextMatrix(intRow, mBillCol.C_采购金额), mFMT.FM_金额)
            End If
        Next
    End With
    Call 显示合计金额
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDraw_Click()
    Dim rsTemp As New Recordset
    Dim blnClear As Boolean, blnCancel As Boolean
    Dim i As Long
    Dim vRect As RECT
    Dim str站点限制 As String
    
    On Error GoTo ErrHandle
    vRect = zlControl.GetControlRect(txtDraw.hwnd)
    str站点限制 = GetDeptStationNode(cboStock.ItemData(cboStock.ListIndex))
    
    If mbln普通科室 Then
        '普通科室申领，只能选择自己所属的科室
        '刘兴宏:20060803
        '问题:8468
        gstrSQL = "" & _
            " SELECT a.id, null as 上级id, 末级, a.编码,a.简码,a.名称 " & _
            " FROM 部门表 a " & _
            " Where (TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01' Or a.撤档时间 Is NULL) " & _
            IIf(str站点限制 <> "", " And (a.站点 = [2] or a.站点 is null) ", "")
        gstrSQL = gstrSQL & " And a.ID in (Select 部门ID From 部门人员 where  人员id = [1] ) "
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "所有领用部门选择", False, "", "选择相关的领用部门", _
                     False, False, True, vRect.Left - 15, vRect.Top, txtDraw.Height, blnCancel, False, False, UserInfo.Id, str站点限制)
    Else
        If gstrNodeNo = "-" Then
            '没有站点号,以树型显示
            gstrSQL = "" & _
                " SELECT  a.id, 上级id, 末级, a.编码,a.简码,a.名称 " & _
                " FROM  部门表 a " & _
                " Where (TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01' Or a.撤档时间 Is NULL) and (a.站点=[1] or a.站点 is null) "
            gstrSQL = gstrSQL & " start with 上级id is null connect by prior id=上级id "
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 1, "所有领用部门选择", False, "", "选择相关的领用部门", _
                         False, False, True, vRect.Left - 15, vRect.Top, txtDraw.Height, blnCancel, False, False, gstrNodeNo)
        Else
            '存在站点，主要是可能上级设置了站点编号，而下级未设置的情况，因此只能用列表方式进行处理
            gstrSQL = "" & _
                " SELECT  a.id, null as 上级id, 末级, a.编码,a.简码,a.名称 " & _
                " FROM  部门表 a " & _
                " Where (TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01' Or a.撤档时间 Is NULL) " & _
                IIf(str站点限制 <> "", " And (a.站点 = [1] or a.站点 is null) ", "")
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "所有领用部门选择", False, "", "选择相关的领用部门", _
                         False, False, True, vRect.Left - 15, vRect.Top, txtDraw.Height, blnCancel, False, False, str站点限制)
        End If
    
    End If
       
       '     frmParent=显示的父窗体
       '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
       '     bytStyle=选择器风格
       '       为0时:列表风格:ID,…
       '       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
       '       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
       '     strTitle=选择器功能命名,也用于个性化区分
       '     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
       '     strSeek=当bytStyle<>2时有效,缺省定位的项目。
       '             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
       '             bytStyle=1时,可以是编码或名称
       '     strNote=选择器的说明文字
       '     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
       '     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
       '     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
       '     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
       '     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
       '     blnSearch=是否显示行号,并可以输入行号定位
    If rsTemp Is Nothing Then
        If txtDraw.Enabled Then txtDraw.SetFocus
        Exit Sub
    End If
    If rsTemp.State <> 1 Then
        If txtDraw.Enabled Then txtDraw.SetFocus
        Exit Sub
    End If
    blnClear = False
    If Val(txtDraw.Tag) <> Val(zlStr.Nvl(rsTemp!Id)) And Val(txtDraw.Tag) <> 0 Then
            '需要检查是否已经有设置好的跟踪部分信息
            With mshBill
                For i = 1 To .Rows - 1
                    If Trim(.TextMatrix(.Row, mBillCol.C_跟踪信息)) <> "" And Trim(.TextMatrix(.Row, mBillCol.C_跟踪信息)) <> "||" Then
                        If MsgBox("在第" & i & "行中,已经设置了跟踪病人信息, " & vbCrLf & "是否需要清空已经设置好的跟踪病人信息?", vbQuestion + vbDefaultButton2 + vbYesNo) = vbNo Then
                            If txtDraw.Enabled Then txtDraw.SetFocus
                            Exit Sub
                        Else
                            blnClear = True
                            Exit For
                        End If
                        
                    End If
                Next
                If blnClear Then
                    For i = 1 To .Rows - 1
                        .TextMatrix(i, mBillCol.C_跟踪信息) = ""
                        .TextMatrix(i, mBillCol.C_跟踪病人) = ""
                    Next
                End If
                
            End With
    End If
    
    Me.txtDraw = zlStr.Nvl(rsTemp!编码) & "-" & zlStr.Nvl(rsTemp!名称)
    Me.txtDraw.Tag = zlStr.Nvl(rsTemp!Id)
    
    gstrSQL = "Select 工作性质, 部门id, 服务对象 From 部门性质说明 Where 部门id=[1] And 工作性质='护理'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, Val(zlStr.Nvl(rsTemp!Id)))
    If rsTemp.EOF Then
        gstrSQL = "Select 工作性质, 部门id, 服务对象 From 部门性质说明 Where 部门id=[1] And 工作性质='临床'"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, Val(Me.txtDraw.Tag))
        If rsTemp.EOF = False Then
            cmdDraw.Tag = "临床"
        Else
            cmdDraw.Tag = ""
        End If
    Else
        cmdDraw.Tag = "护理"
    End If
    If txtDrawPerson.Enabled Then txtDrawPerson.SetFocus
    Local跟踪病人信息
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdDrawPerson_Click()
    If ShowSelect("") = False Then Exit Sub
    mshBill.SetFocus
End Sub

Private Sub cmdExpend_Click()
    Call AutoExpend
End Sub

'查找
Private Sub cmdFind_Click()
    
    If lblCode.Visible = False Then
        lblCode.Visible = True
        txtCode.Visible = True
        txtCode.SetFocus
    Else
        FindRownew mshBill, mBillCol.C_材料, txtCode.Text, True
        lblCode.Visible = False
        txtCode.Visible = False
    End If
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Sub cmdRequestDraw_Click()
    Dim rsTemp As ADODB.Recordset
    Dim lngRow As Long
    Dim blnDo As Boolean
    Dim str灭菌效期 As String
    Dim bln药房 As Boolean
    Dim dblPrice As Double
    Dim str效期 As String
    Dim dbl数量 As Double
    Dim lng材料ID As Long
    Dim bln分批 As Boolean
    Dim dbl申购数量  As Double
    Dim dbl已导数量 As Double
    
    If Val(txtDraw.Tag) = 0 Then
        MsgBox "领料部门不能为空！", vbInformation, gstrSysName
        txtDraw.SetFocus
        Exit Sub
    End If
    
    mstrRequestNO = frmDrawCondition.ShowMe(Me, mintUnit, cboStock.List(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), txtDraw.Text, txtDraw.Tag)
    If mstrRequestNO <> "" Then
        blnDo = False
        mstrRequestNO = Mid(mstrRequestNO, 1, LenB(StrConv(mstrRequestNO, vbFromUnicode)) - 1)
        
        bln药房 = True
        gstrSQL = "Select Distinct 0 " & _
                                    "From 部门性质说明 " & _
                                    "Where ((工作性质 Like '发料部门') Or (工作性质 Like '制剂室')) And 部门id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex))
        If rsTemp.RecordCount = 0 Then
            bln药房 = False
        End If
        
        gstrSQL = "Select a.Id as 材料id, d.数量 as 计划数量,a.编码,a.名称 ,a.规格,c.现价 as 售价,a.计算单位 as 散装单位,a.是否变价 as 时价,b.包装单位,b.换算系数,b.指导差价率" & vbNewLine & _
                    ",e.上次产地 as 产地,e.上次批号 as 批号,nvl(e.批次,0) as 批次,e.效期,e.灭菌效期,e.可用数量,nvl(e.实际数量,0) as 实际数量,e.实际金额,e.实际差价,e.零售价,e.平均成本价,e.批准文号,b.库房分批,b.在用分批, nvl(b.跟踪病人,0) as 跟踪病人" & vbNewLine & _
                    "From 收费项目目录 A, 材料特性 B, 收费价目 C," & vbNewLine & _
                    "     (Select  b.材料id, Sum(b.计划数量) As 数量" & vbNewLine & _
                    "       From 材料采购计划 A, 材料计划内容 B" & vbNewLine & _
                    "       Where a.Id = b.计划id and a.单据=1 And a.No In (Select * From Table(Cast(f_Str2list([1]) As Zltools.t_Strlist)))" & vbNewLine & _
                    "       Group By b.材料id) D,药品库存 e" & vbNewLine & _
                    "Where a.Id = b.材料id And b.材料id = c.收费细目id And a.Id = d.材料id and b.材料id=e.药品id(+)  and e.库房id=[2] and e.实际数量>0 and e.性质=1 And Sysdate Between c.执行日期 And c.终止日期" & _
                    GetPriceClassString("C")
        
        If gSystem_Para.P156_出库算法 = 0 Then
            gstrSQL = gstrSQL & " Order by a.id,Nvl(e.批次, 0)"
        Else
            gstrSQL = gstrSQL & " Order by a.id,e.效期,Nvl(e.批次, 0)"
        End If
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "cmdRequestDraw_Click", mstrRequestNO, cboStock.ItemData(cboStock.ListIndex))
                
        Do While Not rsTemp.EOF
            With mshBill
                For lngRow = 1 To .Rows - 1
                    If Val(.TextMatrix(lngRow, 0)) <> 0 Then
                        If Val(.TextMatrix(lngRow, 0)) = rsTemp!材料ID And Val(.TextMatrix(lngRow, mBillCol.c_批次)) = rsTemp!批次 Then
                            blnDo = True
                            MsgBox "重复材料" & "[" & rsTemp!编码 & "-" & rsTemp!名称 & "]" & "不再添加！", vbInformation, gstrSysName
                            Exit For
                        End If
                    End If
                Next
            
                If Val(.TextMatrix(.Rows - 1, 0)) = 0 Then
                    lngRow = .Rows - 1
                Else
                    .Rows = .Rows + 1
                    lngRow = .Rows - 1
                End If
                
                str灭菌效期 = IIf(IsNull(rsTemp!灭菌效期), "", Format(rsTemp!灭菌效期, "yyyy-MM-dd"))
                If Format(str灭菌效期, "yyyy-mm-dd") < Format(sys.Currentdate, "yyyy-mm-dd") And Trim(str灭菌效期) <> "" Then
                   If MsgBox("[" & rsTemp!编码 & "-" & rsTemp!名称 & "]" & "卫材已经过了灭菌效期,是否还要领用！", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) <> vbYes Then
                        blnDo = True
                   End If
                End If
                
                str效期 = IIf(IsNull(rsTemp!效期), "", Format(rsTemp!效期, "yyyy-MM-dd"))
                If IsDate(str效期) Then
                    If Format(str效期, "yyyy-MM-dd") < Format(sys.Currentdate, "yyyy-MM-dd") Then
                        MsgBox "[" & rsTemp!编码 & "-" & rsTemp!名称 & "]" & "卫生材料已经失效了！", vbInformation, gstrSysName
                    End If
                End If
                
                '取售价
                If rsTemp!时价 = 1 Then
                    If rsTemp!在用分批 = 0 Then
                        If rsTemp!库房分批 = 1 And bln药房 = False Then
                            bln分批 = True
                        Else
                            bln分批 = False
                        End If
                    Else
                        bln分批 = True
                    End If
                                
                    If bln分批 = True Then
                        If IsNull(rsTemp!零售价) Then
                            If rsTemp!实际数量 = 0 Then
                                dblPrice = 0
                            Else
                                dblPrice = rsTemp!实际金额 / rsTemp!实际数量
                            End If
                        Else
                            dblPrice = rsTemp!零售价
                        End If
                    Else
                        If rsTemp!实际数量 = 0 Then
                            dblPrice = 0
                        Else
                            dblPrice = rsTemp!实际金额 / rsTemp!实际数量
                        End If
                    End If
                Else
                    dblPrice = IIf(IsNull(rsTemp!售价), 0, rsTemp!售价)
                End If
                                
                If lng材料ID = rsTemp!材料ID Then
                    If rsTemp!可用数量 + dbl已导数量 > rsTemp!计划数量 Then
                        If rsTemp!计划数量 - dbl已导数量 <> 0 Then
                            dbl数量 = rsTemp!计划数量 - dbl已导数量
                            dbl已导数量 = dbl已导数量 + dbl数量
                        Else
                            blnDo = True
                        End If
                    Else
                        dbl数量 = rsTemp!可用数量
                        dbl已导数量 = dbl已导数量 + dbl数量
                    End If
                Else
                    If rsTemp!可用数量 > rsTemp!计划数量 Then
                        dbl数量 = rsTemp!计划数量
                    Else
                        dbl数量 = rsTemp!可用数量
                    End If
                    dbl已导数量 = dbl数量
                End If
                lng材料ID = rsTemp!材料ID
                               
                If dbl数量 = 0 Then
                    blnDo = True
                End If
                
                '只有不重复的才添加到表格中去
                If blnDo = False Then
                    SetRequestColValue lngRow, rsTemp!材料ID, "[" & rsTemp!编码 & "]" & rsTemp!名称, _
                                IIf(IsNull(rsTemp!规格), "", rsTemp!规格), IIf(IsNull(rsTemp!产地), "", rsTemp!产地), _
                                IIf(mintUnit = 0, rsTemp!散装单位, rsTemp!包装单位), _
                                dblPrice, rsTemp!平均成本价, IIf(IsNull(rsTemp!批号), "", rsTemp!批号), _
                                IIf(IsNull(rsTemp!效期), "", Format(rsTemp!效期, "yyyy-MM-dd")), _
                                IIf(IsNull(rsTemp!灭菌效期), "", Format(rsTemp!灭菌效期, "yyyy-MM-dd")), _
                                rsTemp!计划数量, _
                                IIf(IsNull(rsTemp!可用数量), "0", rsTemp!可用数量), _
                                dbl数量, _
                                IIf(IsNull(rsTemp!指导差价率), "0", rsTemp!指导差价率), _
                                IIf(mintUnit = 0, 1, rsTemp!换算系数), IIf(IsNull(rsTemp!批次), 0, rsTemp!批次), rsTemp!时价, rsTemp!在用分批, IIf(IsNull(rsTemp!批准文号), "", rsTemp!批准文号), rsTemp!跟踪病人, rsTemp!库房分批
                End If
                blnDo = False
                rsTemp.MoveNext
            End With
        Loop
    End If
End Sub

Private Function SetRequestColValue(ByVal intRow As Integer, ByVal lng材料ID As Long, _
        ByVal str材料 As String, ByVal str规格 As String, ByVal str产地 As String, _
        ByVal str单位 As String, ByVal num售价 As Double, ByVal num成本价 As Double, ByVal str批号 As String, _
        ByVal str效期 As String, ByVal str灭菌失效期 As String, ByVal num申购数量 As Double, ByVal num可用数量 As Double, ByVal num实际数量 As Double, _
        ByVal num指导差价率 As Double, _
        ByVal num比例系数 As Double, ByVal lng批次 As Long, _
        ByVal int是否变价 As Integer, ByVal int在用分批 As Integer, ByVal str批准文号 As String, ByVal int跟踪病人 As Integer, ByVal int库房分批 As Integer) As Boolean
    
        Dim intCount As Integer
        Dim intCol As Integer
        Dim dblPrice As Double
        Dim rsTemp As New Recordset
        Dim bln分批 As Boolean
        Dim lngRow As Long
        
    On Error GoTo ErrHandle
    SetRequestColValue = False
    
    With mshBill
        .TextMatrix(intRow, 0) = lng材料ID
        .TextMatrix(intRow, mBillCol.C_材料) = str材料
        .TextMatrix(intRow, mBillCol.c_规格) = str规格
        .TextMatrix(intRow, mBillCol.C_产地) = str产地
        .TextMatrix(intRow, mBillCol.C_批准文号) = str批准文号
        .TextMatrix(intRow, mBillCol.c_单位) = str单位
        .TextMatrix(intRow, mBillCol.c_批号) = str批号
        .TextMatrix(intRow, mBillCol.C_效期) = Format(str效期, "yyyy-mm-dd")
        .TextMatrix(intRow, mBillCol.C_灭菌失效期) = Format(str灭菌失效期, "yyyy-mm-dd")
    
        .TextMatrix(intRow, mBillCol.C_售价) = Format(num售价 * num比例系数, mFMT.FM_零售价)
        .TextMatrix(intRow, mBillCol.C_采购价) = Format(num成本价 * num比例系数, mFMT.FM_成本价)
        .TextMatrix(intRow, mBillCol.C_申购数量) = Format(num申购数量 / num比例系数, mFMT.FM_数量)
        .TextMatrix(intRow, mBillCol.C_可用数量) = Format(num可用数量 / num比例系数, mFMT.FM_数量)
        .TextMatrix(intRow, mBillCol.C_填写数量) = Format(num实际数量 / num比例系数, mFMT.FM_数量)
        .TextMatrix(intRow, mBillCol.C_实际数量) = Format(num实际数量 / num比例系数, mFMT.FM_数量)
        .TextMatrix(intRow, mBillCol.C_售价金额) = Format(Val(.TextMatrix(intRow, mBillCol.C_售价)) * Val(.TextMatrix(intRow, mBillCol.C_填写数量)), mFMT.FM_金额)
        .TextMatrix(intRow, mBillCol.C_采购金额) = Format(Val(.TextMatrix(intRow, mBillCol.C_采购价)) * Val(.TextMatrix(intRow, mBillCol.C_填写数量)), mFMT.FM_金额)
        .TextMatrix(intRow, mBillCol.C_差价) = Format(Val(.TextMatrix(intRow, mBillCol.C_售价金额)) - Val(.TextMatrix(intRow, mBillCol.C_采购金额)), mFMT.FM_金额)
        .TextMatrix(intRow, mBillCol.C_指导差价率) = num指导差价率 & "||" & int是否变价 & "||" & int在用分批
        .TextMatrix(intRow, mBillCol.c_比例系数) = num比例系数
        .TextMatrix(intRow, mBillCol.c_批次) = lng批次
        .TextMatrix(intRow, mBillCol.C_分批属性) = Check分批属性(intRow, int在用分批, int库房分批)
        
        .TextMatrix(intRow, mBillCol.C_跟踪标志) = int跟踪病人
    End With
'    Call 提示库存数
    SetRequestColValue = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdSel_Click()
    Dim lng材料ID As Long, lng收发ID As Long, lng病人id As Long, str使用时间 As String, str条码 As String, blnEdit As Boolean
    Dim strTemp As String, arrtemp As Variant
    Dim str姓名 As String
    
    If Val(txtDraw.Tag) = 0 Then
        ShowMsgBox "领用部门未选择,请先选择领用部门后再选择病人!"
        Exit Sub
    End If
    
    lng收发ID = Get收发ID()
    With mshBill
        lng材料ID = Val(.TextMatrix(.Row, 0))
        strTemp = .TextMatrix(.Row, C_跟踪信息)
        If Trim(strTemp) <> "" Then
            arrtemp = Split(strTemp, "|")
            lng病人id = Val(arrtemp(0))
            str使用时间 = arrtemp(1)
            str条码 = arrtemp(2)
        Else
            lng病人id = 0
            str使用时间 = ""
            str条码 = ""
        End If
    End With
    blnEdit = IIf(mint编辑状态 = 1 Or mint编辑状态 = 2 Or mint编辑状态 = 7, True, False)
    If frmDrawPatiInfor.ShowEdit(Me, lng收发ID, Val(txtDraw.Tag), cmdDraw.Tag, lng材料ID, blnEdit, lng病人id, str条码, str使用时间, str姓名) = False Then
        mshBill.SetFocus
        Exit Sub
    End If
    With mshBill
        .TextMatrix(.Row, mBillCol.C_跟踪信息) = lng病人id & "|" & str使用时间 & "|" & str条码
        .TextMatrix(.Row, mBillCol.C_跟踪病人) = str姓名
        .SetFocus
    End With
End Sub

Private Function AutoExpend(Optional blnCheck As Boolean = False) As Boolean
    Dim lng库房ID As Long, lng材料ID As Long, lng材料ID_Last As Long, lng批次 As Long
    Dim bln库房 As Boolean, bln分批 As Boolean, bln时价 As Boolean, blnAddRow As Boolean
    Dim dbl填写数量 As Double, dbl申领数量 As Double, dbl数量 As Double, dbl比例系数 As Double
    Dim dbl现价 As Currency, dbl现价_时价 As Double, dbl成本价 As Double
    Dim lngCol As Long, lngCols As Long, lngRow As Long
    Dim intCount As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim rsCheck As New ADODB.Recordset
    Dim dbl实际数量 As Double
        
    '对卫材记录进行自动分解，仅处理批次卫材
    On Error GoTo ErrHand
    Screen.MousePointer = 11
    lngRow = 1: lngCols = mshBill.Cols - 1
    lng库房ID = Val(cboStock.ItemData(cboStock.ListIndex))
    bln库房 = CheckStockProperty(lng库房ID)
    
    Do While True
        lng材料ID = Val(mshBill.TextMatrix(lngRow, 0))
        dbl申领数量 = Val(mshBill.TextMatrix(lngRow, mBillCol.C_填写数量))
        dbl填写数量 = Val(mshBill.TextMatrix(lngRow, mBillCol.C_实际数量))
        dbl比例系数 = Val(mshBill.TextMatrix(lngRow, mBillCol.c_比例系数))
        lng批次 = Val(mshBill.TextMatrix(lngRow, mBillCol.c_批次))
        If lng材料ID = 0 Then Exit Do
        
        '提取该卫材对于出库库房是否分批、时价的属性
        If lng材料ID <> lng材料ID_Last Then
            lng材料ID_Last = lng材料ID
            gstrSQL = " Select Nvl(A.库房分批,0) 库房分批,Nvl(A.在用分批,0) 在用分批," & _
                      " Nvl(B.是否变价,0) 时价,Nvl(P.现价,0) 现价,Nvl(A.成本价,0) 成本价" & _
                      " From 材料特性 A,收费项目目录 B,收费价目 P" & _
                      " Where A.材料ID = B.ID And B.ID=P.收费细目ID And A.材料ID =[1] " & _
                      " And Sysdate between P.执行日期 And Nvl(P.终止日期,Sysdate)" & _
                      GetPriceClassString("P")
                      
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取该材料对于出库库房是否分批、时价的属性", lng材料ID)
                      
            bln时价 = (rsTemp!时价 = 1)
            dbl现价 = rsTemp!现价 * dbl比例系数
            dbl成本价 = rsTemp!成本价 * dbl比例系数
            bln分批 = IIf(bln库房, (rsTemp!库房分批 = 1), (rsTemp!在用分批 = 1))
        End If
        
        '如果该卫材批次为零，需要自动分解；批次不为0的已在填单时检查了对应批次的数量
        blnAddRow = False
        If bln分批 = True And lng批次 = 0 Then
            If blnCheck Then
                If dbl填写数量 > Val(mshBill.TextMatrix(lngRow, mBillCol.C_可用数量)) Then
                    MsgBox "第" & lngRow & "行的卫材是批次或时价卫材，而该卫材当前库存不足，不能继续！", vbInformation, gstrSysName
                    Screen.MousePointer = 0: Exit Function
                End If
            End If
            gstrSQL = " Select Nvl(可用数量,0)/" & dbl比例系数 & " As 可用数量,Nvl(实际数量,0)/" & dbl比例系数 & " As 实际数量," & _
                      " Nvl(实际金额,0) 实际金额,Nvl(实际差价,0) 实际差价,平均成本价,nvl(零售价,0) * " & dbl比例系数 & " as 零售价," & _
                      " Nvl(批次,0) 批次,上次批号 批号,to_char(效期,'yyyy-MM-dd') 效期,上次产地 产地,批准文号" & _
                      " From 药品库存 Where nvl(可用数量,0)<>0   and 库房ID=[1] And 药品ID=[2]  And 性质=1 "
            If gSystem_Para.P156_出库算法 = 0 Then '批次还是效期优先先出库
                gstrSQL = gstrSQL & " Order by Nvl(批次, 0)"
            Else
                gstrSQL = gstrSQL & " Order by 效期,Nvl(批次, 0)"
            End If
            
            Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "提取该卫材在指定库存的所有库存记录", lng库房ID, lng材料ID)
                      
            intCount = 0
            With rsCheck
                Do While Not .EOF
                    '重新写记录
                    mblnChange = True
                    
                    intCount = intCount + 1
                    blnAddRow = False
                    If .AbsolutePosition <> 1 Then
                        Call InsertRow(lngRow)
                        For lngCol = 0 To lngCols
                            mshBill.TextMatrix(lngRow, lngCol) = mshBill.TextMatrix(lngRow - 1, lngCol)
                        Next
                        mshBill.RowData(lngRow) = mshBill.RowData(lngRow - 1)
                    End If
   
                    If intCount = 1 Then
                        dbl实际数量 = Val(mshBill.TextMatrix(lngRow, mBillCol.C_实际数量))
                    End If
                    '填写批次相关信息
                    mshBill.TextMatrix(lngRow, mBillCol.C_行号) = lngRow
                    mshBill.TextMatrix(lngRow, mBillCol.C_序号) = lngRow
                    mshBill.TextMatrix(lngRow, mBillCol.c_批次) = rsCheck!批次
                    mshBill.TextMatrix(lngRow, mBillCol.c_批号) = IIf(IsNull(rsCheck!批号), "", rsCheck!批号)
                    mshBill.TextMatrix(lngRow, mBillCol.C_效期) = IIf(IsNull(rsCheck!效期), "", rsCheck!效期)
                    mshBill.TextMatrix(lngRow, mBillCol.C_产地) = IIf(IsNull(rsCheck!产地), "", rsCheck!产地)
                    mshBill.TextMatrix(lngRow, mBillCol.C_批准文号) = IIf(IsNull(rsCheck!批准文号), "", rsCheck!批准文号)
                    mshBill.TextMatrix(lngRow, mBillCol.C_分批属性) = IIf(bln分批 = True, 1, 0)
                    
                    '重新计算价格相关信息
                    If bln时价 = True Then
                        If bln分批 = True Then
                            dbl现价_时价 = rsCheck!零售价
                        Else
                            If rsCheck!实际数量 > 0 Then
                                dbl现价_时价 = rsCheck!实际金额 / rsCheck!实际数量
                            Else
                                dbl现价_时价 = dbl现价
                            End If
                        End If
                    End If
                    
                    If dbl填写数量 <= rsCheck!可用数量 Then
                        dbl数量 = dbl填写数量
                    Else
                        dbl数量 = rsCheck!可用数量
                    End If
                    If dbl数量 > dbl填写数量 Then dbl数量 = dbl填写数量
                    
                    If dbl实际数量 <> mshBill.TextMatrix(lngRow, mBillCol.C_实际数量) Then
                        mshBill.TextMatrix(lngRow, mBillCol.c_原始数量) = Val(mshBill.TextMatrix(lngRow, mBillCol.C_实际数量)) * Val(mshBill.TextMatrix(lngRow, mBillCol.c_比例系数))
                    End If
                    
                    mshBill.TextMatrix(lngRow, mBillCol.C_填写数量) = Format(dbl数量, mFMT.FM_数量)
                    mshBill.TextMatrix(lngRow, mBillCol.C_实际数量) = Format(dbl数量, mFMT.FM_数量)
                    
                    If Trim(mshBill.TextMatrix(lngRow, mBillCol.C_实际数量)) = "" Then mshBill.TextMatrix(lngRow, mBillCol.C_实际数量) = 0
                    mshBill.TextMatrix(lngRow, mBillCol.C_实际差价) = Format(rsCheck!实际差价, mFMT.FM_金额)
                    mshBill.TextMatrix(lngRow, mBillCol.C_实际金额) = Format(rsCheck!实际金额, mFMT.FM_金额)
                    mshBill.TextMatrix(lngRow, mBillCol.C_可用数量) = Format(rsCheck!可用数量, mFMT.FM_金额)
                    mshBill.TextMatrix(lngRow, mBillCol.C_售价) = Format(IIf(bln时价, dbl现价_时价, dbl现价), mFMT.FM_零售价)
                    mshBill.TextMatrix(lngRow, mBillCol.C_售价金额) = Format(Val(mshBill.TextMatrix(lngRow, mBillCol.C_售价)) * dbl数量, mFMT.FM_金额)
                    
                    '采用新的方式计算成本价 成本价=药品库存.平均成本价
                    mshBill.TextMatrix(lngRow, mBillCol.C_采购价) = Format(Get成本价(lng材料ID, lng库房ID, Val(mshBill.TextMatrix(lngRow, mBillCol.c_批次))) * dbl比例系数, mFMT.FM_成本价)
                    mshBill.TextMatrix(lngRow, mBillCol.C_采购金额) = Format(Val(mshBill.TextMatrix(lngRow, mBillCol.C_采购价)) * dbl数量, mFMT.FM_金额)
                    mshBill.TextMatrix(lngRow, mBillCol.C_差价) = Format(Val(mshBill.TextMatrix(lngRow, mBillCol.C_售价金额)) - Val(mshBill.TextMatrix(lngRow, mBillCol.C_采购金额)), mFMT.FM_金额)
                    
                    dbl填写数量 = dbl填写数量 - dbl数量
                    dbl申领数量 = dbl申领数量 - dbl数量
                    If dbl填写数量 = 0 Then Exit Do
                    lngRow = lngRow + 1
                    blnAddRow = True
                    .MoveNext
                Loop
                If dbl申领数量 <> 0 And rsCheck.RecordCount <> 0 Then
                    If blnAddRow Then
                        mshBill.TextMatrix(lngRow - 1, mBillCol.C_填写数量) = Format(dbl申领数量 + dbl数量, mFMT.FM_数量)
                    Else
                        mshBill.TextMatrix(lngRow, mBillCol.C_填写数量) = Format(dbl申领数量 + dbl数量, mFMT.FM_数量)
                    End If
                End If
            End With
            
            '如果库存记录为零，则说明未进行分解，需要将申领数量与实际数量清为零
            If dbl填写数量 <> 0 And rsCheck.RecordCount = 0 Then
                mshBill.TextMatrix(lngRow, mBillCol.C_行号) = lngRow
                mshBill.TextMatrix(lngRow, mBillCol.C_序号) = lngRow
                mshBill.TextMatrix(lngRow, mBillCol.C_实际数量) = ""
                mshBill.TextMatrix(lngRow, mBillCol.C_售价金额) = ""
                mshBill.TextMatrix(lngRow, mBillCol.C_采购金额) = ""
                mshBill.TextMatrix(lngRow, mBillCol.C_差价) = ""
            End If
        Else
            mshBill.TextMatrix(lngRow, mBillCol.C_行号) = lngRow
            mshBill.TextMatrix(lngRow, mBillCol.C_序号) = lngRow
        End If
        If blnAddRow = False Then lngRow = lngRow + 1
    Loop
    
    AutoExpend = True
    Screen.MousePointer = 0
    Exit Function
ErrHand:
    Screen.MousePointer = 0
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub InsertRow(ByVal lngRow As Long)
    Dim lngReserve As Long, lngRows As Long
    Dim lngCol As Long, lngCols As Long
    lngReserve = lngRow
    lngRows = mshBill.Rows - 1
    lngCols = mshBill.Cols - 1
    mshBill.Rows = mshBill.Rows + 1
    
    '将当前行及以下行全部下移
    For lngRow = lngRows To lngReserve Step -1
        For lngCol = 0 To lngCols
            mshBill.TextMatrix(lngRow + 1, lngCol) = mshBill.TextMatrix(lngRow, lngCol)
        Next
        mshBill.RowData(lngRow + 1) = mshBill.RowData(lngRow)
        '校正行号
        mshBill.TextMatrix(lngRow + 1, mBillCol.C_行号) = lngRow + 1
    Next
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
'    mblnChange = False
    Select Case mintParallelRecord
        Case 1
            '正常
        Case 2
            If mint编辑状态 = 6 Then
                ShowMsgBox "该单据已没有可以冲销的卫材，请检查！"
            Else
                '单据已被删除
                ShowMsgBox "该单据已被删除，请检查！"
            End If
            Unload Me
            Exit Sub
        Case 3
            '修改的单据已被审核
            ShowMsgBox "该单据已被其他人审核，请检查！"
            Unload Me
            Exit Sub
    End Select
    If mint编辑状态 = 7 Then
        If IsCtrlSetFocus(CmdSave) Then
            zlControl.ControlSetFocus CmdSave
        End If
    End If
    '初始化简码方式
    If (mint编辑状态 = 1 Or mint编辑状态 = 2) And gbytSimpleCodeTrans = 1 Then
        stbThis.Panels("PY").Visible = True
        stbThis.Panels("WB").Visible = True
        gSystem_Para.int简码方式 = Val(zlDatabase.GetPara("简码方式", , , 0))    '默认拼音简码
        Logogram stbThis, gSystem_Para.int简码方式
    Else
        stbThis.Panels("PY").Visible = False
        stbThis.Panels("WB").Visible = False
    End If
End Sub

Private Function CheckStockProperty(ByVal lng库房ID As Long) As Boolean
    Dim rsCheck As New ADODB.Recordset
    '检查指定库房是库房、发料部门还是制剂室(传入的库房肯定是库房、发料部门或制剂室中的一个)
    On Error GoTo ErrHandle
    gstrSQL = " Select 部门ID From 部门性质说明 " & _
              " Where (工作性质 like '发料部门' Or 工作性质 like '%制剂室') And 部门id=[1]"
    
    Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "判断是不是库房或制剂室", lng库房ID)
              
    If rsCheck.EOF Then
        CheckStockProperty = True
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckStock() As Boolean
    Dim dbl比例系数 As Double, dbl可用数量 As Double, dbl填写数量 As Double
    Dim lngRow As Long, lngRows As Long, int库存检查 As Integer
    Dim lng材料ID As Long, lng库房ID As Long, lng批次 As Long
    Dim bln库房 As Boolean, bln特殊 As Boolean
    Dim str材料ID As String, strMsg As String
    Dim rsProperty As New ADODB.Recordset           '材料规格
    Dim rsCheck As New ADODB.Recordset              '材料库存
    Dim bln下库存 As Boolean
    
    
    '检查单据中各材料的库存
    'mint库存检查:0-不检查;1-检查，不足提醒；2-检查，不足禁止
    '分批或时价材料不受此限'
    
    On Error GoTo ErrHandle
    bln下库存 = Val(zlDatabase.GetPara(95, glngSys, 0)) = 1
    
    lngRows = mshBill.Rows - 1
    lng库房ID = Val(cboStock.ItemData(cboStock.ListIndex))
    bln库房 = CheckStockProperty(lng库房ID)
    
    For lngRow = 1 To lngRows
        lng材料ID = Val(mshBill.TextMatrix(lngRow, 0))
        If lng材料ID <> 0 Then
            If InStr(1, str材料ID & ",", "," & lng材料ID & ",") = 0 Then str材料ID = str材料ID & "," & lng材料ID
        End If
    Next
    
    If str材料ID = "" Then
        CheckStock = True
        Exit Function
    Else
        str材料ID = Mid(str材料ID, 2)
    End If
    
    '提取本单据内所有材料的属性
    gstrSQL = " Select A.材料ID,'['||B.编码||']'||B.名称 通用名,A.库房分批,A.在用分批,B.是否变价" & _
              " From 材料特性 A,收费项目目录 B,Table(Cast(f_Num2List([1]) As zlTools.t_NumList)) C" & _
              " Where A.材料ID=B.ID And A.材料ID =C.Column_Value "
    Set rsProperty = zlDatabase.OpenSQLRecord(gstrSQL, "提取本单据内所有材料的属性", str材料ID)

    '提取本单据内所有材料的当前库存（没有库存的材料该记录集中也不会有记录）
    gstrSQL = " Select A.药品id 材料ID,Nvl(A.批次,0) As 批次," & _
              " SUM(NVL(可用数量,0)) As 可用数量,SUM(NVL(实际数量,0)) As 实际数量" & _
              " From 药品库存 A,收费项目目录 B,材料特性 C,Table(Cast(f_Num2List([2]) As zlTools.t_NumList)) D " & _
              " Where A.库房ID=[1] And A.药品ID=B.ID And B.ID=C.材料ID And A.性质=1 " & _
              "         And A.药品ID=D.Column_Value" & _
              " Group by A.药品ID,Nvl(A.批次,0)"
    Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "提取本单据内所有材料的当前库存", lng库房ID, str材料ID)
              
    '检查每个材料
    For lngRow = 1 To lngRows
        lng材料ID = Val(mshBill.TextMatrix(lngRow, 0))
        If lng材料ID <> 0 Then
            lng批次 = Val(mshBill.TextMatrix(lngRow, mBillCol.c_批次))
            dbl比例系数 = Val(mshBill.TextMatrix(lngRow, mBillCol.c_比例系数))
            dbl填写数量 = Val(mshBill.TextMatrix(lngRow, mBillCol.C_实际数量))
            
            dbl可用数量 = 0
            '查找该材料的库存记录
            rsCheck.Filter = "材料ID=" & lng材料ID & " And 批次=" & lng批次
            If rsCheck.RecordCount <> 0 Then
                dbl可用数量 = zlStr.Nvl(rsCheck!实际数量, 0) / dbl比例系数
            End If
            
            '如果库存的可用数量不够
            If Not (dbl可用数量 >= dbl填写数量) Then
                int库存检查 = mint库存检查
                
                rsProperty.Filter = "材料ID=" & lng材料ID
                
                If Not (Val(mshBill.TextMatrix(lngRow, mBillCol.C_分批属性)) = 0 And Split(mshBill.TextMatrix(lngRow, mBillCol.C_指导差价率), "||")(1) = 0) Then
                    '如果该材料是时价或分批，库存不足不允许出库，相当于禁止出库；定价不分批不需要判断分解，只需要根据参数控制即可
                    bln特殊 = (IIf(bln库房, (rsProperty!库房分批 = 1), (rsProperty!在用分批 = 1)) Or (rsProperty!是否变价 = 1))
                    strMsg = ""
                    If bln特殊 Then
                        int库存检查 = 2
                        '如果是批次材料，但批次小于等于零，说明未执行分解功能
                        If lng批次 <= 0 And IIf(bln库房, (rsProperty!库房分批 = 1), (rsProperty!在用分批 = 1)) Then
                            strMsg = "（请先执行分解功能明确批次材料的出库批次）"
                        End If
                    End If
                End If
                '按正常流程进行提示或禁止
                Select Case int库存检查
                Case 1  '仅提示
                    If MsgBox(rsProperty!通用名 & "的库存不足" & "(库存实际数量为" & dbl可用数量 & ")，是否继续？" & strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        mshBill.Row = lngRow
                        mshBill.MsfObj.TopRow = lngRow
                        Exit Function
                    End If
                Case 2
                    MsgBox rsProperty!通用名 & "的库存不足" & "(库存实际数量为" & dbl可用数量 & ")！" & strMsg, vbInformation, gstrSysName
                    mshBill.Row = lngRow
                    mshBill.MsfObj.TopRow = lngRow
                    Exit Function
                End Select
            End If
        End If
    Next
    rsCheck.Filter = 0
    rsCheck.Close
    rsProperty.Filter = 0
    rsProperty.Close
    CheckStock = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 70 Or KeyCode = 102 Then
        If Shift = vbCtrlMask Then   'Ctrl+F
            cmdFind_Click
        End If
    ElseIf KeyCode = vbKeyF3 Then
        FindRownew mshBill, mBillCol.C_材料, txtCode.Text, False
    ElseIf KeyCode = vbKeyF7 Then
        If stbThis.Panels("PY").Bevel = sbrRaised Then
            Logogram stbThis, 0
        Else
            Logogram stbThis, 1
        End If
    End If
End Sub

Private Sub CmdSave_Click()
    Dim blnSuccess As Boolean
    Dim strReg As String
    Dim intRow As Integer
    
    '设置排序数据集
    Call SetSortRecord
    
    If mint编辑状态 = 4 Then    '查看
        '打印
        printbill
        '退出
        Unload Me
        Exit Sub
    End If
    
    If mint编辑状态 = 5 Then        '财务审核
        
        If Not 材料单据审核(Txt填制人.Caption) Then Exit Sub
        
        mstrTime_End = GetBillInfo(20, txtNo.Tag)
        If mstrTime_End = "" Then
            MsgBox "注意:" & vbCrLf & "  该单据已经被其他操作员删除,不能继续！", vbInformation, gstrSysName
            Exit Sub
        End If
        If mstrTime_End <> mstrTime_Start Then
            If MsgBox("注意:" & vbCrLf & "  该单据已经被其他操作员编辑，不能继续!" & vbCrLf & "  是否重新刷新单据?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                Call initCard
            End If
            Exit Sub
        End If
        '检查是否分解
        If CheckStock = False Then Exit Sub
        
        If Not 检查单价(20, txtNo.Tag, False) And Not mblnUpdate Then
            MsgBox "有记录未使用最新价格，程序将自动完成更新（售价、成本价、售价金额、成本金额、差价），更新后请检查！", vbInformation, gstrSysName
            Call RefreshBill
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
        
        '如果审核时修改了单据，则重新生成单据保存
        If mblnChange Then
            If Not SaveCard() Then
                gcnOracle.RollbackTrans: Exit Sub
            End If
        End If
                
        If SaveCheck = True Then
            Unload Me
        End If
        Exit Sub
    End If
    
    If mint编辑状态 = 3 Then        '审核
        
        If Not 材料单据审核(Txt填制人.Caption) Then Exit Sub
        
        '刘兴宏:2007/06/10:问题10813
        mstrTime_End = GetBillInfo(20, txtNo.Tag)
        If mstrTime_End = "" Then
            MsgBox "注意:" & vbCrLf & "  该单据已经被其他操作员删除,不能继续！", vbInformation, gstrSysName
            Exit Sub
        End If
        If mstrTime_End <> mstrTime_Start Then
            If MsgBox("注意:" & vbCrLf & "  该单据已经被其他操作员编辑，不能继续!" & vbCrLf & "  是否重新刷新单据?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                Call initCard
            End If
            Exit Sub
        End If
        
        '检查是否分解
        If CheckStock = False Then Exit Sub
        
        For intRow = 1 To mshBill.Rows - 1
            If Val(mshBill.TextMatrix(intRow, 0)) <> 0 Then
                If Val(mshBill.TextMatrix(intRow, mBillCol.C_分批属性)) = 1 And Val(mshBill.TextMatrix(intRow, mBillCol.C_实际数量)) = 0 Then
                    MsgBox "第" & intRow & "行的卫材是批次卫材且无库存，不允许0数量领用！", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        Next
        
        If Not 检查单价(20, txtNo.Tag, False) And Not mblnUpdate Then
            MsgBox "有记录未使用最新价格，程序将自动完成更新（售价、成本价、售价金额、成本金额、差价），更新后请检查！", vbInformation, gstrSysName
            Call RefreshBill
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
        
        '如果审核时修改了单据，则重新生成单据保存
        If mblnChange Then
            If Not SaveCard(True) Then
                gcnOracle.RollbackTrans: Exit Sub
            End If
        End If
                
        If SaveCheck = True Then
            strReg = IIf(Val(zlDatabase.GetPara("审核打印", glngSys, mlngModule, "0")) = 1, 1, 0)
            If Val(strReg) = 1 Then
                '打印
                If InStr(mstrPrivs, "单据打印") <> 0 Then
                    printbill
                End If
            End If
            Unload Me
        End If
        Exit Sub
    End If
    
    If mint编辑状态 = 6 Then '冲销
        For intRow = 1 To mshBill.Rows - 1
            If Val(mshBill.TextMatrix(intRow, mBillCol.C_实际数量)) < 0 Then '负数领用冲销才检查
                If CompareUsableQuantity(intRow, Val(mshBill.TextMatrix(intRow, mBillCol.C_实际数量))) = False Then
                    mshBill.SetFocus
                    mshBill.Row = intRow
                    Exit Sub
                End If
            End If
        Next
        
        If SaveStrike Then Unload Me
        Exit Sub
    End If
            
    If ValidData = False Then Exit Sub
    blnSuccess = SaveCard
        
    If blnSuccess = True Then
        strReg = IIf(Val(zlDatabase.GetPara("存盘打印", glngSys, mlngModule, "0")) = 1, 1, 0)
        If Val(strReg) = 1 Then
            '打印
            If InStr(mstrPrivs, "单据打印") <> 0 Then
                printbill
            End If
        End If
        If mint编辑状态 = 2 Then   '修改
            Unload Me
            Exit Sub
        End If
    Else
        Exit Sub
    End If
   
    If mint编辑状态 = 7 Then    '从入库单读取
        Unload Me
        Exit Sub
    End If
    
'    If mbln单据增加 Then
'        mstr单据号 = NextNo(73)
'        txtNO = mstr单据号
'    End If
    
    mblnSave = False
    mblnEdit = True
    mshBill.ClearBill
    Call RefreshRowNO(mshBill, mBillCol.C_行号, 1)

    txtDraw.Text = ""
    txtDraw.Tag = "0"
    txt摘要.Text = ""
    If txtDraw.Enabled = True Then
        txtDraw.SetFocus
        txtDraw.SelStart = 0
        txtDraw.SelLength = Len(txtDraw.Text)
    End If
    mblnChange = False
    If txtNo.Tag <> "" Then Me.stbThis.Panels(2).Text = "上一张单据的NO号：" & txtNo.Tag
End Sub

Private Sub RefreshBill()
    '以最新价格最新单据相关数据，用于单据审核时
    Dim lngRow As Long, lngRows As Long, lng材料ID As Long
    Dim dbl数量 As Double, dbl成本价 As Double, dbl成本金额 As Double, dbl零售价 As Double, dbl零售金额 As Double, dbl差价 As Double
    Dim rsprice As New ADODB.Recordset
    Dim rsStock As ADODB.Recordset
    Dim blnAdj As Boolean
    
    On Error GoTo ErrHandle
    
    gstrSQL = " Select '售价' As 类型, a.序号, a.药品id As 材料id, Nvl(a.批次, 0) As 批次, b.现价" & _
            " From 药品收发记录 A," & _
                 " (Select 收费细目id, Nvl(现价, 0) 现价, 执行日期" & _
                   " From 收费价目" & _
                   " Where (终止日期 Is Null Or Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'yyyy-MM-dd')))" & _
                   GetPriceClassString("") & ") B, 收费项目目录 C" & _
            " Where a.单据 = 20 And a.No = [1] And a.药品id = b.收费细目id And c.Id = b.收费细目id And Round(a.零售价," & g_小数位数.obj_散装小数.零售价小数 & ") <> Round(b.现价, " & g_小数位数.obj_散装小数.零售价小数 & ") And" & _
              "    NVL(c.是否变价, 0) = 0" & _
            " Union All" & _
            " Select '售价' As 类型, a.序号, a.药品id As 材料id, Nvl(a.批次, 0) As 批次, decode(nvl(b.批次,0),0,b.实际金额 / b.实际数量,b.零售价) As 现价" & _
            " From 药品收发记录 A, 药品库存 B, 收费项目目录 C" & _
            " Where a.单据 = 20 And a.No = [1] And c.Id = a.药品id And Round(a.零售价," & g_小数位数.obj_散装小数.零售价小数 & ") <> Round(decode(nvl(b.批次,0),0,b.实际金额 / b.实际数量,b.零售价), " & g_小数位数.obj_散装小数.零售价小数 & ") And Nvl(c.是否变价, 0) = 1 And" & _
                  " b.性质 = 1 And b.库房id = a.库房id And b.药品id = a.药品id And NVL(b.批次, 0) = NVL(a.批次, 0) And NVL(b.实际数量, 0) <> 0 And a.入出系数 = -1" & _
            " Union All" & _
            " Select '成本价' As 类型, a.序号, a.药品id As 材料id, Nvl(a.批次, 0) As 批次, b.平均成本价 As 现价" & _
            " From 药品收发记录 A, 药品库存 B" & _
            " Where a.单据 = 20 And a.No = [1] And a.药品id = b.药品id And Nvl(a.批次, 0) = Nvl(b.批次, 0) and round(a.成本价," & g_小数位数.obj_散装小数.成本价小数 & ")<>round(b.平均成本价," & g_小数位数.obj_散装小数.成本价小数 & ") And a.库房id = b.库房id and a.入出系数=-1 and b.性质=1" & _
            " Order By 类型, 材料id, 序号"

    Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "[取当前价格]", CStr(Me.txtNo.Text))
    
    If rsprice.EOF Then Exit Sub
    
    lngRows = mshBill.Rows - 1
    For lngRow = 1 To lngRows
        blnAdj = False
        lng材料ID = Val(mshBill.TextMatrix(lngRow, 0))
        dbl数量 = Val(mshBill.TextMatrix(lngRow, mBillCol.C_实际数量))
        dbl成本价 = Val(mshBill.TextMatrix(lngRow, mBillCol.C_采购价))
        dbl零售价 = Val(mshBill.TextMatrix(lngRow, mBillCol.C_售价))
        dbl成本金额 = dbl成本价 * dbl数量
        dbl零售金额 = dbl零售价 * dbl数量
        dbl差价 = dbl零售金额 - dbl成本金额
'
        If lng材料ID <> 0 Then
            rsprice.Filter = "类型='售价' And 材料id=" & lng材料ID & " And 批次=" & Val(mshBill.TextMatrix(lngRow, mBillCol.c_批次))
            If rsprice.RecordCount > 0 Then
                blnAdj = True
                dbl零售价 = Val(Format(rsprice!现价 * Val(mshBill.TextMatrix(lngRow, mBillCol.c_比例系数)), mFMT.FM_零售价))
                dbl零售金额 = Val(Format(dbl零售价 * dbl数量, mFMT.FM_金额))
                dbl差价 = Val(Format(dbl零售金额 - dbl成本金额, mFMT.FM_金额))
            End If

            rsprice.Filter = "类型='成本价' And 材料id=" & lng材料ID & " And 批次=" & Val(mshBill.TextMatrix(lngRow, mBillCol.c_批次))
            If rsprice.RecordCount > 0 Then
                blnAdj = True
                dbl零售金额 = Val(Format(dbl零售价 * dbl数量, mFMT.FM_金额))
                dbl成本价 = Val(Format(rsprice!现价 * Val(mshBill.TextMatrix(lngRow, mBillCol.c_比例系数)), mFMT.FM_金额))
                dbl成本金额 = Val(Format(dbl成本价 * dbl数量, mFMT.FM_金额))
                dbl差价 = Val(Format(dbl零售金额 - dbl成本金额, mFMT.FM_金额))
            End If

            If blnAdj = True Then
                '以当前最新价格最新单据相关数据（售价、成本价、零售金额、成本金额、差价）
                mshBill.TextMatrix(lngRow, mBillCol.C_售价) = Format(dbl零售价, mFMT.FM_零售价)
                mshBill.TextMatrix(lngRow, mBillCol.C_售价金额) = Format(dbl零售金额, mFMT.FM_金额)
                mshBill.TextMatrix(lngRow, mBillCol.C_采购价) = Format(dbl成本价, mFMT.FM_成本价)
                mshBill.TextMatrix(lngRow, mBillCol.C_采购金额) = Format(dbl成本金额, mFMT.FM_金额)
                mshBill.TextMatrix(lngRow, mBillCol.C_差价) = Format(dbl差价, mFMT.FM_金额)
            End If
        End If
    Next
    rsprice.Filter = 0
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim strReg As String
    
    mblnUpdate = False
    strReg = Val(zlDatabase.GetPara("卫材单位", glngSys, mlngModule, "0"))
    mintUnit = Val(strReg)
    mint领用明确批次 = Val(zlDatabase.GetPara(258, glngSys, 0))
    
    mblnFirst = True
  
    '刘兴宏:增加小数格式化串
    With mFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价)
        .FM_金额 = GetFmtString(mintUnit, g_金额)
        .FM_零售价 = GetFmtString(mintUnit, g_售价)
        .FM_数量 = GetFmtString(mintUnit, g_数量)
    End With
    
       
    txtNo = mstr单据号
    txtNo.Tag = txtNo.Text
    
    '------------------------------------------------------------------------------------------------------------------
    '刘兴宏:20060803:部门申领
    '问题:8468
    mbln普通科室 = Check普通科室
    '------------------------------------------------------------------------------------------------------------------
    Call initCard
    
    '恢复个性化参数设置
    RestoreWinState Me, App.ProductName, mstrCaption
    '恢复个性化参数设置后，还需要对权限控制的列进一步设置
    With mshBill
        .ColWidth(mBillCol.C_采购价) = IIf(mblnCostView = True, 900, 0)
        .ColWidth(mBillCol.C_采购金额) = IIf(mblnCostView = True, 900, 0)
        .ColWidth(mBillCol.C_差价) = IIf(mblnCostView = True, 900, 0)
    End With
    mblnChange = False
End Sub

Private Sub initCard()
    Dim i As Integer
    Dim rsTemp As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim intRow As Integer
    Dim numUseAbleCount As Double
    Dim vardrug As Variant
    Dim strOrder As String, strCompare As String
    Dim str批次 As String, strArray As String
    
    '新增、修改按申领单领用按钮可见
    If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
        cmdRequestDraw.Visible = True
    Else
        cmdRequestDraw.Visible = False
    End If
    
    '库房
    On Error GoTo ErrHandle
    strOrder = zlDatabase.GetPara("单据排序", glngSys, mlngModule, "00")
    strOrder = IIf(strOrder = "", "00", strOrder)
    
    strCompare = Mid(strOrder, 1, 1)
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
            Txt填制人 = UserInfo.用户名
            Txt填制日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            initGrid
            '如果是普通科室领用,则看是否只具备一个科室,如果当前人员只具体一个科室，则填充在领用部门中
            If mbln普通科室 Then
                gstrSQL = "" & _
                   "   SELECT DISTINCT a.id, a.编码,a.简码,a.名称 " & _
                   "   FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " & _
                   "   Where c.工作性质 = b.名称 " & _
                   "           AND a.id = c.部门id " & _
                   "           AND (TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01' Or a.撤档时间 Is NULL)" & _
                   "           And a.ID in (Select 部门ID From 部门人员 where 缺省=1 and 人员id =[1])"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, UserInfo.Id)
                
                If Not rsTemp.EOF Then
                    Me.txtDraw = zlStr.Nvl(rsTemp!编码) & "-" & zlStr.Nvl(rsTemp!名称)
                    Me.txtDraw.Tag = zlStr.Nvl(rsTemp!Id)
                End If
                txtDrawPerson.Text = gstrUserName
                txtDrawPerson.Tag = gstrUserName
            End If
            txt摘要 = mstr默认材料用途
        Case 2, 3, 4, 5, 6, 7
            initGrid
            
            If mint编辑状态 = 4 Then
                gstrSQL = "" & _
                    "   Select b.id,b.名称 " & _
                    "   From 药品收发记录 a,部门表 b " & _
                    "   Where a.库房id=b.id and A.单据 = 20 and a.no=[1] "
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr单据号)
                
                If rsTemp.EOF Then
                    mintParallelRecord = 2
                    Exit Sub
                End If

                With cboStock
                    .AddItem rsTemp!名称
                    .ItemData(.NewIndex) = rsTemp!Id
                    .ListIndex = 0
                End With
                rsTemp.Close
            End If
            Select Case mintUnit
                Case 0
                    strUnitQuantity = "d.计算单位 AS 单位,a.单量 as 申购数量, A.填写数量,a.实际数量,a.成本价,a.零售价,'1' as 比例系数,"
                Case Else
                    strUnitQuantity = "B.包装单位 AS 单位,(A.单量 / B.换算系数) AS 申购数量,(A.填写数量 / B.换算系数) AS 填写数量,(A.实际数量 / B.换算系数) AS 实际数量,a.成本价*B.换算系数 as 成本价,a.零售价*B.换算系数 as 零售价,B.换算系数 as 比例系数,"
            End Select
            
            Select Case mint编辑状态
            Case 7
                    gstrSQL = "" & _
                    "   Select w.材料ID,w.序号,w.卫材信息,W.名称,w.规格,w.原产地,w.产地,w.批准文号 ,w.批次,w.批号,w.指导差价率, " & _
                    "           w.库房分批,w.最大效期,w.效期,w.灭菌日期,w.灭菌失效期, " & _
                    "           w.一次性材料,w.灭菌效期,w.单位,w.原始数量 原始数量,w.填写数量,w.实际数量,w.申购数量,w.零售价,w.零售金额,w.比例系数, " & _
                    "           (w.零售金额 - Decode(Sign(nvl(z.实际金额,0)),1,w.零售金额 * (nvl(z.实际差价,0) / z.实际金额),w.零售金额 * w.指导差价率 / 100)) / decode(w.实际数量,0,1,w.实际数量)  成本价, " & _
                    "           (w.零售金额 - Decode(Sign(z.实际金额),1,w.零售金额 * (z.实际差价 / z.实际金额),w.零售金额 * w.指导差价率 / 100)) 成本金额, " & _
                    "           Decode(Sign(z.实际金额),1,w.零售金额 * (z.实际差价 / z.实际金额),w.零售金额 * w.指导差价率 / 100) 差价, " & _
                    "            w.摘要,w.填制人,w.填制日期,w.配药人, w.审核人, w.审核日期,w.库房id,w.对方部门id,W.领用部门,W.领用人,w.是否变价,w.在用分批,z.可用数量/w.比例系数 as  可用数量,z.实际金额,z.实际差价,W.跟踪病人   " & _
                    "    From (  SELECT distinct a.药品id 材料id,A.序号,('[' || D.编码 || ']' || D.名称) AS 卫材信息,  " & _
                    "                    zlSpellCode(D.名称) 名称,D.规格,D.产地 as 原产地,A.产地,A.批准文号, A.批号,a.批次,b.指导差价率,b.库房分批,  " & _
                    "                    b.最大效期,A.效期,A.灭菌日期,A.灭菌效期 as 灭菌失效期,B.一次性材料,b.灭菌效期,A.填写数量 原始数量, " & strUnitQuantity & _
                    "                    A.成本金额,A.零售金额, A.差价," & _
                    "                    a.摘要,a.填制人,A.填制日期,A.配药人, A.审核人, A.审核日期,a.库房id ,D.是否变价,b.在用分批," & _
                    "                    M.ID as 对方部门ID,M.名称 as 领用部门,[5] as 领用人,b.跟踪病人 " & _
                    "            FROM 药品收发记录 A, 材料特性 B,收费项目目录 D,(Select ID,名称 From 部门表 where id=[4] ) M  " & _
                    "            Where A.药品id = B.材料id and a.药品id=D.id   " & _
                    "                    AND A.记录状态 =[3]  " & _
                    "                    AND A.单据 = 15 AND A.No = [1]  " & _
                    "           ) w  , (  Select 药品id 材料id,Nvl(批次,0) 批次,可用数量,实际金额,实际差价   " & _
                    "                    From 药品库存 where 库房id=[2]  and 性质=1)  z " & _
                    "    Where w.材料id=z.材料id(+)  and nvl(w.批次,0)=nvl(z.批次(+),0)   " & _
                    "   ORDER BY " & IIf(strCompare = "0", "序号", IIf(strCompare = "1", "卫材信息", "名称")) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
                    
            Case 6
                    gstrSQL = "" & _
                    "   Select w.*,z.可用数量/w.比例系数 可用数量,nvl(z.实际数量,0) / w.比例系数 As 库存数量,z.实际金额,z.实际差价 " & _
                    "   From (  SELECT distinct a.材料id,A.序号,('[' || d.编码 || ']' || d.名称) AS 卫材信息," & _
                    "                   zlSpellCode(d.名称) 名称,d.规格,d.产地 as 原产地,A.产地,A.批准文号, A.批号,a.批次,b.指导差价率,a.效期," & _
                                        strUnitQuantity & _
                    "                   a.填写数量 原始数量,A.成本金额,0 零售金额,0 差价, " & _
                    "                   a.摘要,a.库房id,a.对方部门id,c.名称 as 领用部门,a.领用人,d.是否变价,b.库房分批,b.在用分批," & _
                    "                   b.跟踪病人,a.病人ID,a.主页ID,a.姓名,a.性别,a.年龄,a.床号,a.医疗付款方式,a.当前科室ID,a.当前病区ID,a.使用时间,a.条码,a.灭菌效期 " & _
                    "           FROM (  Select min(x.id) as id, sum(x.实际数量) as 填写数量,0 实际数量,sum(x.成本金额) as 成本金额,x.领用人,x.药品id 材料ID," & _
                    "                           x.序号,x.产地,x.批准文号, x.批号,x.效期,x.灭菌效期,0 as 单量,Nvl(x.批次,0) 批次,x.扣率,x.成本价,x.零售价,x.摘要,x.库房ID,x.对方部门ID,x.入出类别ID," & _
                    "                           max(M.病人ID) as 病人ID,max(M.主页ID) as 主页ID,max(M.姓名) as 姓名,max(M.性别) 性别,max(M.年龄) 年龄,max(M.床号) as 床号,max(M.医疗付款方式) 医疗付款方式,max(M.当前科室ID) 当前科室ID,max(M.当前病区ID) 当前病区ID,max(M.使用时间) 使用时间,max(M.条码 ) 条码 " & _
                    "                   From 药品收发记录 x,材料领用信息 M  " & _
                    "                   WHERE x.NO=[1] AND x.单据=20 and x.id=M.收发ID(+) " & _
                    "                   Group by x.药品ID,x.序号,x.产地,x.批准文号,x.批号,x.效期,x.灭菌效期,Nvl(x.批次,0),x.扣率,x.成本价,x.零售价,x.摘要,x.库房ID,x.对方部门ID,x.入出类别ID,x.领用人" & _
                    "                   having sum(x.填写数量)<>0 " & _
                    "               ) A, 材料特性 B,收费项目目录 D,部门表 C " & _
                    "           Where A.材料id = B.材料id  and a.材料id=d.id AND a.对方部门id=c.id " & _
                    "       ) w,(Select  药品id 材料id,Nvl(批次,0) 批次,可用数量,实际数量,实际金额,实际差价 " & _
                    "            From 药品库存 " & _
                    "            Where 库房id=[2] and 性质=1)  z " & _
                    "   Where w.材料id=z.材料id(+) and nvl(w.批次,0)=nvl(z.批次(+),0) " & _
                    "   ORDER BY " & IIf(strCompare = "0", "序号", IIf(strCompare = "1", "卫材信息", "名称")) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
            Case Else
                    gstrSQL = "" & _
                    "   Select w.*,z.可用数量/w.比例系数 可用数量,z.实际金额,z.实际差价 " & _
                    "   From (  SELECT distinct a.药品id 材料id,A.序号,('[' || d.编码 || ']' ||d.名称) AS 卫材信息," & _
                    "                   zlSpellCode(d.名称) 名称,d.规格,d.产地 as 原产地,A.产地,A.批准文号, A.批号,a.批次,b.指导差价率,a.效期," & _
                                        strUnitQuantity & _
                    "                   a.填写数量 原始数量,A.成本金额,A.零售金额, A.差价, " & _
                    "                   a.摘要,a.领用人,a.填制人,a.填制日期,a.配药人 as 核查人,a.配药日期 as 核查日期,a.审核人,a.审核日期,a.灭菌效期,a.库房id,a.对方部门id,c.名称 as 领用部门,d.是否变价,b.库房分批,b.在用分批 ," & _
                    "                   b.跟踪病人,M.病人ID,M.主页ID,M.姓名,M.性别,M.年龄,M.床号,M.医疗付款方式,M.当前科室ID,M.当前病区ID,M.使用时间,M.条码 " & _
                    "           FROM 药品收发记录 A, 材料特性 B,收费项目目录 D,部门表 C,材料领用信息 M" & _
                    "           Where A.药品id = B.材料id and a.药品id=d.id and A.id=M.收发ID(+)  " & _
                    "                   AND a.对方部门id=c.id and A.记录状态 =[3]" & _
                    "                   AND A.单据 = 20 AND A.No = [1] " & _
                    "           ) w,(   Select  药品id 材料id,Nvl(批次,0) 批次,可用数量,实际金额,实际差价 " & _
                    "                   From 药品库存 where 库房id=[2]   and 性质=1)  z " & _
                    "   Where w.材料id=z.材料id(+) and nvl(w.批次,0)=nvl(z.批次(+),0)" & _
                    " ORDER BY " & IIf(strCompare = "0", "序号", IIf(strCompare = "1", "卫材信息", "名称")) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
            End Select
            
            
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "卫材领用单", IIf(mint编辑状态 = 7, mstr入库单号, mstr单据号), cboStock.ItemData(cboStock.ListIndex), mint记录状态, mlng部门ID, mstr领用人)
            '刘兴宏:2007/06/10:问题10813
            mstrTime_Start = GetBillInfo(20, mstr单据号)
             
            If rsTemp.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            
            Select Case mint编辑状态
            Case 2, 6, 7
                Txt填制人 = UserInfo.用户名
                Txt填制日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                If mint编辑状态 = 6 Then
                    Txt审核人 = UserInfo.用户名
                    Txt审核日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                End If
            Case Else
                Txt填制人 = rsTemp!填制人
                Txt填制日期 = Format(rsTemp!填制日期, "yyyy-mm-dd hh:mm:ss")
                Txt审核人 = IIf(IsNull(rsTemp!审核人), "", rsTemp!审核人)
                Txt审核日期 = IIf(IsNull(rsTemp!审核日期), "", Format(rsTemp!审核日期, "yyyy-mm-dd hh:mm:ss"))
                txt核查人 = IIf(IsNull(rsTemp!核查人), "", rsTemp!核查人)
                txt核查日期 = IIf(IsNull(rsTemp!核查日期), "", Format(rsTemp!核查日期, "yyyy-mm-dd hh:mm:ss"))
            End Select
            txt摘要.Text = IIf(IsNull(rsTemp!摘要), "", rsTemp!摘要)
            
            If (mint编辑状态 = 2 Or mint编辑状态 = 3 Or mint编辑状态 = 5) And Txt审核人 <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            txtDraw.Text = rsTemp!领用部门
            txtDraw.Tag = rsTemp!对方部门id
            
            txtDrawPerson.Text = zlStr.Nvl(rsTemp!领用人)
            txtDrawPerson.Tag = zlStr.Nvl(rsTemp!领用人)
            
            If mint编辑状态 = 2 Or mint编辑状态 = 3 Or mint编辑状态 = 5 Then
                Set mcolUsedCount = New Collection
            End If
            
            intRow = 0
            With mshBill
                Do While Not rsTemp.EOF
                    
                    intRow = intRow + 1
                    .Rows = intRow + 1
                    .TextMatrix(intRow, 0) = rsTemp.Fields(0)
                    .TextMatrix(intRow, mBillCol.C_材料) = rsTemp!卫材信息
                    .TextMatrix(intRow, mBillCol.C_序号) = rsTemp!序号
                    .TextMatrix(intRow, mBillCol.c_规格) = IIf(IsNull(rsTemp!规格), "", rsTemp!规格)
                    .TextMatrix(intRow, mBillCol.C_产地) = IIf(IsNull(rsTemp!产地), "", rsTemp!产地)
                    .TextMatrix(intRow, mBillCol.C_批准文号) = IIf(IsNull(rsTemp!批准文号), "", rsTemp!批准文号)
                    .TextMatrix(intRow, mBillCol.c_单位) = rsTemp!单位
                    .TextMatrix(intRow, mBillCol.c_批号) = IIf(IsNull(rsTemp!批号), "", rsTemp!批号)
                    .TextMatrix(intRow, mBillCol.C_效期) = IIf(IsNull(rsTemp!效期), "", Format(rsTemp!效期, "yyyy-mm-dd"))
                    
                    .TextMatrix(intRow, mBillCol.C_申购数量) = Format(rsTemp!申购数量, mFMT.FM_数量)
                    .TextMatrix(intRow, mBillCol.C_填写数量) = Format(rsTemp!填写数量, mFMT.FM_数量)
                    .TextMatrix(intRow, mBillCol.C_实际数量) = Format(rsTemp!实际数量, mFMT.FM_数量)
                    
                    .TextMatrix(intRow, mBillCol.c_原始数量) = Val(zlStr.Nvl(rsTemp!原始数量))
                    
                    If mint编辑状态 = 6 Then
                        .TextMatrix(intRow, mBillCol.C_库存数量) = Format(rsTemp!库存数量, mFMT.FM_数量) '只有冲销时才加载
                    End If
                    
                    .TextMatrix(intRow, mBillCol.C_采购价) = Format(rsTemp!成本价, mFMT.FM_成本价)
                    .TextMatrix(intRow, mBillCol.C_采购金额) = Format(IIf(mint编辑状态 = 6, 0, rsTemp!成本金额), mFMT.FM_金额)
                    .TextMatrix(intRow, mBillCol.C_售价) = Format(rsTemp!零售价, mFMT.FM_零售价)
                    .TextMatrix(intRow, mBillCol.C_售价金额) = Format(rsTemp!零售金额, mFMT.FM_金额)
                    .TextMatrix(intRow, mBillCol.C_差价) = Format(rsTemp!差价, mFMT.FM_金额)
                    .TextMatrix(intRow, mBillCol.c_批次) = IIf(IsNull(rsTemp!批次), "0", rsTemp!批次)
                    .TextMatrix(intRow, mBillCol.c_比例系数) = rsTemp!比例系数
                    .TextMatrix(intRow, mBillCol.C_指导差价率) = rsTemp!指导差价率 & "||" & rsTemp!是否变价 & "||" & rsTemp!在用分批
                    .TextMatrix(intRow, mBillCol.C_可用数量) = IIf(IsNull(rsTemp!可用数量), "0", rsTemp!可用数量)
                    .TextMatrix(intRow, mBillCol.C_实际差价) = IIf(IsNull(rsTemp!实际差价), "0", rsTemp!实际差价)
                    .TextMatrix(intRow, mBillCol.C_实际金额) = IIf(IsNull(rsTemp!实际金额), "0", rsTemp!实际金额)
                    .TextMatrix(intRow, mBillCol.C_灭菌失效期) = IIf(IsNull(rsTemp!灭菌效期), "", Format(rsTemp!灭菌效期, "yyyy-mm-dd"))
                    
                    .TextMatrix(intRow, mBillCol.C_跟踪标志) = zlStr.Nvl(rsTemp!跟踪病人)
                    .TextMatrix(intRow, mBillCol.C_分批属性) = Check分批属性(intRow, rsTemp!在用分批, rsTemp!库房分批)
                    If mint编辑状态 <> 7 Then
                        '病人ID|使用时间|条码
                        .TextMatrix(intRow, mBillCol.C_跟踪信息) = zlStr.Nvl(rsTemp!病人ID) & "|" & IIf(IsNull(rsTemp!使用时间), "", Format(rsTemp!使用时间, "yyyy-mm-dd")) & "|" & zlStr.Nvl(rsTemp!条码)
                        .TextMatrix(intRow, mBillCol.C_跟踪病人) = zlStr.Nvl(rsTemp!姓名)
                    End If
                    
                    If mint编辑状态 = 2 Or mint编辑状态 = 3 Or mint编辑状态 = 5 Then
                        numUseAbleCount = 0
                        For Each vardrug In mcolUsedCount
                            If vardrug(0) = CStr(rsTemp!材料ID & IIf(IsNull(rsTemp!批次), "0", rsTemp!批次)) Then
                                numUseAbleCount = vardrug(1)
                                mcolUsedCount.Remove vardrug(0)
                                Exit For
                            End If
                        Next
                        str批次 = rsTemp!材料ID & IIf(IsNull(rsTemp!批次), "0", rsTemp!批次)
                        If mint编辑状态 = 2 Or mint编辑状态 = 5 Then
                            strArray = numUseAbleCount + IIf(IsNull(rsTemp!填写数量), "0", rsTemp!填写数量)
                        Else
                            strArray = numUseAbleCount + IIf(IsNull(rsTemp!实际数量), "0", rsTemp!实际数量)
                        End If
                        mcolUsedCount.Add Array(str批次, strArray), str批次
                    End If
                    rsTemp.MoveNext
                Loop
                .Rows = intRow + 2
            End With
            rsTemp.Close
    End Select
    
    gstrSQL = "Select 工作性质, 部门id, 服务对象 From 部门性质说明 Where 部门id=[1] And 工作性质='护理'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, Val(Me.txtDraw.Tag))
    If rsTemp.EOF Then
        gstrSQL = "Select 工作性质, 部门id, 服务对象 From 部门性质说明 Where 部门id=[1] And 工作性质='临床'"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, Val(Me.txtDraw.Tag))
        If rsTemp.EOF = False Then
            cmdDraw.Tag = "临床"
        Else
            cmdDraw.Tag = ""
        End If
    Else
        cmdDraw.Tag = "护理"
    End If
    


    rsTemp.Close
    Call RefreshRowNO(mshBill, mBillCol.C_行号, 1)
    Call 显示合计金额
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub initGrid()
    With mshBill
        .Active = True
        .Cols = mBillCols
        
        .MsfObj.FixedCols = 1
        .TextMatrix(0, mBillCol.C_行号) = ""
        .TextMatrix(0, mBillCol.C_材料) = "名称与编码"
        .TextMatrix(0, mBillCol.C_序号) = "序号"
        .TextMatrix(0, mBillCol.c_规格) = "规格"
        .TextMatrix(0, mBillCol.C_产地) = "产地"
        .TextMatrix(0, mBillCol.C_批准文号) = "批准文号"
        .TextMatrix(0, mBillCol.c_单位) = "单位"
        .TextMatrix(0, mBillCol.c_批号) = "批号"
        .TextMatrix(0, mBillCol.C_效期) = "失效期"
        .TextMatrix(0, mBillCol.C_灭菌失效期) = "灭菌失效期"
                
        .TextMatrix(0, mBillCol.C_申购数量) = "申购数量"
        .TextMatrix(0, mBillCol.C_填写数量) = IIf(mint编辑状态 = 6, "数量", "填写数量")
        .TextMatrix(0, mBillCol.C_实际数量) = IIf(mint编辑状态 = 6, "冲销数量", "实际数量")
        .TextMatrix(0, mBillCol.c_原始数量) = "原始数量"
        
        .TextMatrix(0, mBillCol.C_采购价) = "成本价"
        .TextMatrix(0, mBillCol.C_采购金额) = "成本金额"
        .TextMatrix(0, mBillCol.C_售价) = "售价"
        .TextMatrix(0, mBillCol.C_售价金额) = "售价金额"
        .TextMatrix(0, mBillCol.C_差价) = "差价"
        .TextMatrix(0, mBillCol.C_可用数量) = "可用数量"
        .TextMatrix(0, mBillCol.C_实际差价) = "实际差价"
        .TextMatrix(0, mBillCol.C_实际金额) = "实际金额"
        .TextMatrix(0, mBillCol.C_指导差价率) = "指导差价率"
        .TextMatrix(0, mBillCol.c_比例系数) = "比例系数"
        .TextMatrix(0, mBillCol.c_批次) = "批次"
         
        .TextMatrix(0, mBillCol.C_跟踪标志) = "跟踪标志"
        .TextMatrix(0, mBillCol.C_跟踪信息) = "跟踪信息"
        .TextMatrix(0, mBillCol.C_跟踪病人) = "跟踪病人"
        .TextMatrix(0, mBillCol.C_分批属性) = "分批属性"
        .TextMatrix(0, mBillCol.C_库存数量) = "库存数量" '领用负数冲销时有用
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, mBillCol.C_行号) = "1"
        
        .ColWidth(0) = 0
        .ColWidth(mBillCol.C_行号) = 300
        .ColWidth(mBillCol.C_材料) = 2000
        .ColWidth(mBillCol.C_序号) = 0
        .ColWidth(mBillCol.c_规格) = 900
        .ColWidth(mBillCol.C_产地) = 800
        .ColWidth(mBillCol.C_批准文号) = 1000
        .ColWidth(mBillCol.c_单位) = 500
        .ColWidth(mBillCol.c_批号) = 800
        .ColWidth(mBillCol.C_效期) = 1000
        .ColWidth(mBillCol.C_灭菌失效期) = 1000
        .ColWidth(mBillCol.C_申购数量) = IIf(mint编辑状态 = 6, 0, 800)
        .ColWidth(mBillCol.C_填写数量) = 800
        .ColWidth(mBillCol.C_实际数量) = 800
        .ColWidth(mBillCol.c_原始数量) = 0
        .ColWidth(mBillCol.C_库存数量) = 0
        
        .ColWidth(mBillCol.C_采购价) = IIf(mblnCostView = False, 0, 800)
        .ColWidth(mBillCol.C_采购金额) = IIf(mblnCostView = False, 0, 900)
        .ColWidth(mBillCol.C_售价) = 800
        .ColWidth(mBillCol.C_售价金额) = 900
        .ColWidth(mBillCol.C_差价) = IIf(mblnCostView = False, 0, 800)
        
        .ColWidth(mBillCol.C_可用数量) = 0
        
        .ColWidth(mBillCol.C_实际差价) = 0
        .ColWidth(mBillCol.C_实际金额) = 0
        .ColWidth(mBillCol.C_指导差价率) = 0
        .ColWidth(mBillCol.c_比例系数) = 0
        .ColWidth(mBillCol.c_批次) = 0
        .ColWidth(mBillCol.C_跟踪信息) = 0
        .ColWidth(mBillCol.C_跟踪标志) = 0
        .ColWidth(mBillCol.C_跟踪病人) = 1000
        .ColWidth(mBillCol.C_分批属性) = 0

        '-1：表示该列可以选择，是布尔型［"√"，" "］
        ' 0：表示该列可以选择，但不能修改
        ' 1：表示该列可以输入，外部显示为按钮选择
        ' 2：表示该列是日期列，外部显示为按钮选择，弹出是日期选择框
        ' 3：表示该列是选择列，外部显示为下拉框选择
        '4:  表示该列为单纯的文本框供用户输入
        '5:  表示该列不允许选择

        .ColData(0) = 5
        .ColData(mBillCol.C_行号) = 5
        .ColData(mBillCol.c_规格) = 5
        .ColData(mBillCol.C_序号) = 5
        .ColData(mBillCol.C_产地) = 5
        .ColData(mBillCol.C_批准文号) = 5
        .ColData(mBillCol.c_单位) = 5
        .ColData(mBillCol.c_批号) = 5
        .ColData(mBillCol.C_效期) = 5
        .ColData(mBillCol.C_灭菌失效期) = 5
        .ColData(mBillCol.C_申购数量) = 5
        .ColData(mBillCol.c_原始数量) = 5

        .ColData(mBillCol.C_跟踪信息) = 0
        .ColData(mBillCol.C_跟踪病人) = 5
        .ColData(mBillCol.C_跟踪标志) = 5
        .ColData(mBillCol.C_分批属性) = 5
        .ColData(mBillCol.C_库存数量) = 5

        If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
            txtDraw.Enabled = True
            cmdDraw.Enabled = True
            txt摘要.Enabled = True
            txtDrawPerson.Enabled = True
            cmdDrawPerson.Enabled = True
            cboStock.Enabled = True

            .ColData(mBillCol.C_材料) = 1
            .ColData(mBillCol.C_填写数量) = 4
            .ColData(mBillCol.C_实际数量) = 5
        ElseIf mint编辑状态 = 3 Or mint编辑状态 = 5 Or mint编辑状态 = 6 Then
            cboStock.Enabled = False
            
            txtDraw.Enabled = False
            cmdDraw.Enabled = False
            txtDrawPerson.Enabled = False
            cmdDrawPerson.Enabled = False
            txt摘要.Enabled = False
            
            .ColData(mBillCol.C_填写数量) = 5
            .ColData(mBillCol.C_实际数量) = 4
        ElseIf mint编辑状态 = 4 Then
            cboStock.Enabled = False
            
            txtDraw.Enabled = False
            cmdDraw.Enabled = False
            txtDrawPerson.Enabled = False
            cmdDrawPerson.Enabled = False
            txt摘要.Enabled = False
            
            .ColData(mBillCol.C_填写数量) = 5
            .ColData(mBillCol.C_实际数量) = 5
            
        End If
        
        .ColData(mBillCol.C_采购价) = 5
        .ColData(mBillCol.C_采购金额) = 5
        .ColData(mBillCol.C_售价) = 5
        .ColData(mBillCol.C_售价金额) = 5
        .ColData(mBillCol.C_差价) = 5
        
        
        .ColData(mBillCol.C_可用数量) = 5
        
        .ColData(mBillCol.C_实际差价) = 5
        .ColData(mBillCol.C_实际金额) = 5
        .ColData(mBillCol.C_指导差价率) = 5
        .ColData(mBillCol.c_比例系数) = 5
        .ColData(mBillCol.c_批次) = 5
        
        .ColAlignment(mBillCol.C_材料) = flexAlignLeftCenter
        .ColAlignment(mBillCol.c_规格) = flexAlignLeftCenter
        .ColAlignment(mBillCol.C_产地) = flexAlignLeftCenter
        .ColAlignment(mBillCol.C_批准文号) = flexAlignLeftCenter
        .ColAlignment(mBillCol.c_单位) = flexAlignCenterCenter
        .ColAlignment(mBillCol.c_批号) = flexAlignLeftCenter
        .ColAlignment(mBillCol.C_效期) = flexAlignLeftCenter
        .ColAlignment(mBillCol.C_申购数量) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_填写数量) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_实际数量) = flexAlignRightCenter
        
        .ColAlignment(mBillCol.C_采购价) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_采购金额) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_售价) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_售价金额) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_差价) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_灭菌失效期) = flexAlignCenterCenter
        .ColAlignment(mBillCol.C_跟踪病人) = flexAlignLeftCenter
        .PrimaryCol = mBillCol.C_材料
        .LocateCol = mBillCol.C_材料
        If InStr(1, "345", mint编辑状态) <> 0 Then .ColData(mBillCol.C_材料) = 0
    End With
    txt摘要.MaxLength = sys.FieldsLength("药品收发记录", "摘要")
    chkIn.Visible = (mint编辑状态 = 1)
    txtIn.Visible = (mint编辑状态 = 1)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    With Pic单据
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - .Top - 100 - CmdCancel.Height - 200
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
        LblNO.Left = .Left - LblNO.Width - 100
        .Top = LblTitle.Top
        LblNO.Top = .Top
    End With
    
    
    LblStock.Left = mshBill.Left
    cboStock.Left = LblStock.Left + LblStock.Width + 100
    
    cmdDrawPerson.Left = mshBill.Left + mshBill.Width - cmdDraw.Width
    txtDrawPerson.Left = cmdDrawPerson.Left - txtDrawPerson.Width
    lbl领用人.Left = txtDrawPerson.Left - lbl领用人.Width '
    
    cmdDraw.Left = lbl领用人.Left - cmdDraw.Width * 2
    txtDraw.Left = cmdDraw.Left - txtDraw.Width
    LblEnterStock.Left = txtDraw.Left - LblEnterStock.Width - 100
    
    
    With Lbl填制日期
        .Top = Pic单据.Height - 200 - .Height
        .Left = mshBill.Left + 100
    End With
    
    With Txt填制日期
        .Top = Lbl填制日期.Top - 80
        .Left = Lbl填制日期.Left + Lbl填制日期.Width + 100
    End With
    
    
    With Lbl填制人
        .Top = Lbl填制日期.Top - .Height - 140
        .Left = mshBill.Left + 100
    End With
    
    With Txt填制人
        .Top = Lbl填制人.Top - 80
        .Left = Lbl填制人.Left + Lbl填制人.Width + 100
    End With
    
    With lbl核查人
        .Top = Lbl填制人.Top
        .Left = Abs(mshBill.Width - .Width - txt核查人.Width - 100) / 2
    End With
    With txt核查人
        .Top = lbl核查人.Top - 80
        .Left = lbl核查人.Left + lbl核查人.Width + 100
    End With
    
    With lbl核查日期
        .Top = Lbl填制日期.Top
        .Left = lbl核查人.Left
    End With
    With txt核查日期
        .Top = Txt填制日期.Top
        .Left = txt核查人.Left
    End With
    
    
    With Txt审核日期
        .Top = Lbl填制日期.Top - 80
        .Left = mshBill.Left + mshBill.Width - .Width
    End With
    
    With Lbl审核日期
        .Top = Lbl填制日期.Top
        .Left = Txt审核日期.Left - 100 - .Width
    End With
    
    With Txt审核人
        .Top = Lbl填制人.Top - 80
        .Left = mshBill.Left + mshBill.Width - .Width
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
    End With
    If mblnCostView = False Then
        lblPurchasePrice.Visible = False
    End If
    
    With lblSalePrice
        .Left = lblPurchasePrice.Left + mshBill.Width / 3
    End With
        
    With lblDifference
        .Left = lblPurchasePrice.Left + mshBill.Width / 3 * 2
    End With
    If mblnCostView = False Then
        lblDifference.Visible = False
    End If
    
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
    
    With cmdRequestDraw
        .Top = cmdHelp.Top
        .Left = cmdHelp.Left + cmdHelp.Width + 100
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
    
    If mint编辑状态 = 5 Or mint编辑状态 = 3 Then
        cmdExpend.Visible = True
        cmdExpend.Move CmdSave.Left - cmdExpend.Width - 100, CmdSave.Top
    End If
    
    Call Local跟踪病人信息
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mshProvider.Visible = True Then
        mshProvider.Visible = False
        txtDraw.SetFocus
        txtDraw.SelLength = Len(txtDraw.Text)
        txtDraw.SelStart = 0
        Cancel = True
        Exit Sub
    End If
    
    If mblnChange = False Or mint编辑状态 = 4 Or mint编辑状态 = 3 Or mint编辑状态 = 5 Then
        SaveWinState Me, App.ProductName, mstrCaption
        Exit Sub
    End If
    If MsgBox("数据可能已改变，但未存盘，真要退出吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
        Exit Sub
    Else
        SaveWinState Me, App.ProductName, mstrCaption
    End If
End Sub

Private Function SaveCheck() As Boolean
    Dim rsTemp As New Recordset
    Dim intRow As Integer
    
    Dim strNo As String
    Dim lng库房ID As Long
    Dim lng对方部门id As Long
    Dim str审核人 As String
    Dim dat审核日期 As String
    
    Dim int序号 As Integer
    Dim lng材料ID As Long
    Dim str产地 As String
    Dim lng批次 As Long
    Dim dbl填写数量 As Double
    Dim dbl实际数量 As Double
    Dim dbl成本价 As Double
    Dim dbl成本金额 As Double
    Dim dbl零售金额 As Double
    Dim dbl差价 As Double
    Dim lng入出类别ID As Long
    Dim str批号 As String
    Dim str效期 As String
    Dim arrSQL As Variant
    Dim n As Long
    
    mblnSave = False
    SaveCheck = False
    arrSQL = Array()
    
    On Error GoTo ErrHandle
    lng库房ID = cboStock.ItemData(cboStock.ListIndex)
    lng对方部门id = txtDraw.Tag
    str审核人 = UserInfo.用户名
    strNo = txtNo.Tag
    gstrSQL = "" & _
        "   SELECT b.id " & _
        "   FROM 药品单据性质 a, 药品入出类别 b " & _
        "   Where a.类别id = b.ID  AND a.单据 = 35 "
    zlDatabase.OpenRecordset rsTemp, gstrSQL, mstrCaption
   
    If rsTemp.EOF Then
        MsgBox "没有设置卫材领用的入出类别，请在入出分类中设置!", vbExclamation + vbOKOnly, gstrSysName
        Exit Function
    End If
    lng入出类别ID = rsTemp!Id
    rsTemp.Close
    
    dat审核日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
    With mshBill
        '按药品ID顺序更新数据
        recSort.Sort = "药品id,批次,序号"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!行号
'        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                
                lng材料ID = Val(.TextMatrix(intRow, 0))
                str产地 = .TextMatrix(intRow, mBillCol.C_产地)
                lng批次 = Val(.TextMatrix(intRow, mBillCol.c_批次))
                dbl填写数量 = Round(Val(.TextMatrix(intRow, mBillCol.C_填写数量)) * Val(.TextMatrix(intRow, mBillCol.c_比例系数)), g_小数位数.obj_最大小数.数量小数)
                dbl实际数量 = Round(Val(.TextMatrix(intRow, mBillCol.C_实际数量)) * Val(.TextMatrix(intRow, mBillCol.c_比例系数)), g_小数位数.obj_最大小数.数量小数)
                
'                If dbl填写数量 = dbl实际数量 Then
'                    dbl填写数量 = Round(Val(.TextMatrix(intRow, mBillCol.c_原始数量)), g_小数位数.obj_最大小数.数量小数)
'                    dbl实际数量 = dbl填写数量
'                End If
                
                dbl成本价 = Round(Val(.TextMatrix(intRow, mBillCol.C_采购价)) / Val(.TextMatrix(intRow, mBillCol.c_比例系数)), g_小数位数.obj_最大小数.成本价小数)
                dbl成本金额 = Round(Val(.TextMatrix(intRow, mBillCol.C_采购金额)), g_小数位数.obj_最大小数.金额小数)
                dbl零售金额 = Round(Val(.TextMatrix(intRow, mBillCol.C_售价金额)), g_小数位数.obj_最大小数.金额小数)
                
                dbl差价 = Round(Val(.TextMatrix(intRow, mBillCol.C_差价)), g_小数位数.obj_最大小数.金额小数)
                str批号 = .TextMatrix(intRow, mBillCol.c_批号)
                str效期 = IIf(.TextMatrix(intRow, mBillCol.C_效期) = "", "Null", "to_date('" & .TextMatrix(intRow, mBillCol.C_效期) & "','yyyy-mm-dd')")
                int序号 = Val(.TextMatrix(intRow, mBillCol.C_序号))
                         
                'zl_材料领用_VERIFY( /*NO_IN*/, /*库房ID_IN*/, /*对方部门ID_IN*/,
                    '/*材料ID_IN*/, /*产地_IN*/, /*批次_IN*/, /*填写数量_IN*/,
                    '/*实际数量_IN*/, /*成本价_IN*/, /*成本金额_IN*/, /*零售金额_IN*/,
                    '/*差价_IN*/, /*入出类别ID_IN*/, /*审核人_IN*/, /*审核日期_IN*/,
                    '/*批号_IN*/, /*效期_IN*/, /*审核方式_In*/ );
                    
                gstrSQL = "zl_材料领用_Verify(" & int序号 & ",'" & strNo & "'," & lng库房ID & "," & lng对方部门id & "," & _
                     lng材料ID & ",'" & str产地 & "'," & lng批次 & "," & dbl填写数量 & "," & _
                     dbl实际数量 & "," & dbl成本价 & "," & dbl成本金额 & "," & dbl零售金额 & "," & _
                     dbl差价 & "," & lng入出类别ID & ",'" & str审核人 & "',to_date('" & dat审核日期 & "','yyyy-mm-dd HH24:MI:SS'),'" & _
                     str批号 & "'," & str效期 & "," & IIf(mint编辑状态 = 3, 0, 1) & ")"
                     
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = CStr(lng材料ID) & ";" & vbCrLf & gstrSQL
            End If
            
            recSort.MoveNext
        Next
    End With
    
    If Not ExecuteSql(arrSQL, mstrCaption, False) Then Exit Function
'    If Not 检查单价(20, txtNO.Tag) Then
'        gcnOracle.RollbackTrans
'        Exit Function
'    End If
    gcnOracle.CommitTrans
    
    SaveCheck = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SaveStrike() As Boolean
    Dim 行次_IN As Integer
    Dim 原记录状态_IN As Integer
    Dim NO_IN As String
    Dim 序号_IN As Integer
    Dim 材料ID_IN As Long
    Dim 冲销数量_IN As Double
    Dim 填制人_IN As String
    Dim 填制日期_IN  As String
    Dim intRow As Integer
    Dim n As Long
    
    SaveStrike = False
    With mshBill
        '检查冲销数量，不能小于零
        For intRow = 1 To .Rows - 1
            If Val(.TextMatrix(intRow, mBillCol.C_实际数量)) <> 0 Then
                If Not 相同符号(Val(.TextMatrix(intRow, mBillCol.C_填写数量)), Val(.TextMatrix(intRow, mBillCol.C_实际数量))) Then
                    MsgBox "请输入合法的冲销数量（第" & intRow & "行）！", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        Next
    
        NO_IN = Trim(txtNo.Tag)
        填制人_IN = UserInfo.用户名
        填制日期_IN = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        原记录状态_IN = mint记录状态
        
        On Error GoTo ErrHandle
        gcnOracle.BeginTrans
        
        行次_IN = 0
        
        '按药品ID顺序更新数据
        recSort.Sort = "药品id,批次,序号"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!行号
'        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" And Val(.TextMatrix(intRow, mBillCol.C_实际数量)) <> 0 Then
                行次_IN = 行次_IN + 1
                
                材料ID_IN = Val(.TextMatrix(intRow, 0))
                冲销数量_IN = Round(Val(.TextMatrix(intRow, mBillCol.C_实际数量)) * Val(.TextMatrix(intRow, mBillCol.c_比例系数)), g_小数位数.obj_散装小数.数量小数)
                If Val(.TextMatrix(intRow, mBillCol.C_实际数量)) = Val(.TextMatrix(intRow, mBillCol.C_填写数量)) Then
                    冲销数量_IN = Val(.TextMatrix(intRow, mBillCol.c_原始数量))
                End If
                
                
                序号_IN = Val(.TextMatrix(intRow, mBillCol.C_序号))
                
                'ZL_材料领用_STRIKE(/*行次_IN*/,/*原记录状态_IN*/,/*NO_IN*/,/*序号_IN*/, /*材料ID_IN*/,
                '/*冲销数量_IN*/,/*填制人_IN*/, /*填制日期_IN*/);
                gstrSQL = "ZL_材料领用_STRIKE(" & _
                    行次_IN & "," & _
                    原记录状态_IN & ",'" & _
                    NO_IN & "'," & _
                    序号_IN & "," & _
                    材料ID_IN & "," & _
                    冲销数量_IN & ",'" & _
                     填制人_IN & "',to_date('" & _
                     Format(填制日期_IN, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS'))"
                
                Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
            End If
            
            recSort.MoveNext
        Next
        gcnOracle.CommitTrans
        
        If 行次_IN = 0 Then
            MsgBox "没有选择一行材料来冲销，不能冲销，请检查！", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveStrike = True
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Function

Private Sub mshBill_AfterAddRow(Row As Long)
    Call RefreshRowNO(mshBill, mBillCol.C_行号, Row)
End Sub

Private Sub mshBill_AfterDeleteRow()
    Call RefreshRowNO(mshBill, mBillCol.C_行号, mshBill.Row)
End Sub

Private Sub mshBill_BeforeAddRow(Row As Long)
    If mshBill.ColData(mBillCol.C_材料) = 0 Then
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
            If MsgBox("你确实要删除该行卫材？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True
            End If
        End If
    End With
End Sub
Private Function Get收发ID() As Long
    '------------------------------------------------------------------------------------------
    '功能:获取当前行的收发ID
    '参数:
    '返回:收发ID
    '------------------------------------------------------------------------------------------
    Dim lng材料ID As Long, lng序号 As Long
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    If mint编辑状态 = 1 Then Exit Function
    With mshBill
        lng材料ID = Val(.TextMatrix(.Row, 0))
        lng序号 = Val(.TextMatrix(.Row, C_序号))
    End With
    gstrSQL = "Select ID From 药品收发记录 where 单据=20 and  NO=[1] AND (记录状态=1 or mod(记录状态,3)=0) and 药品id=[2] and 序号=[3]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr单据号, lng材料ID, lng序号)
    If rsTemp.EOF = False Then
        Get收发ID = Val(zlStr.Nvl(rsTemp!Id))
        Exit Function
    End If
    Get收发ID = 0
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

 
Private Sub mshbill_CommandClick()
    Dim RecReturn As Recordset
    Dim lng材料ID As Long, lng收发ID As Long, lng病人id As Long, str使用时间 As String, str条码 As String, blnEdit As Boolean
    Dim strTemp As String, arrtemp As Variant
    Dim i As Integer
    Dim int点击行 As Integer
    
    int点击行 = mshBill.Row
    
    Select Case mshBill.Col
    Case C_跟踪病人
    Case Else
            Set RecReturn = Frm材料选择器.ShowMe(Me, 2, cboStock.ItemData(cboStock.ListIndex), _
                                cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), _
                                IIf(mint领用明确批次 = 1, True, False), , , , , , , , , , mlngModule, , mstrPrivs, IIf(mint领用明确批次 = 1, True, False), False)
            If RecReturn.RecordCount > 0 Then
                mblnChange = True
                With mshBill
                    RecReturn.MoveFirst
                    For i = 1 To RecReturn.RecordCount
                        If SetColValue(.Row, RecReturn!材料ID, "[" & RecReturn!编码 & "]" & RecReturn!名称, _
                            IIf(IsNull(RecReturn!规格), "", RecReturn!规格), IIf(IsNull(RecReturn!产地), "", RecReturn!产地), _
                            IIf(mintUnit = 0, RecReturn!散装单位, RecReturn!包装单位), _
                            IIf(IsNull(RecReturn!售价), 0, RecReturn!售价), IIf(IsNull(RecReturn!批号), "", RecReturn!批号), _
                            IIf(IsNull(RecReturn!效期), "", Format(RecReturn!效期, "yyyy-MM-dd")), _
                            IIf(IsNull(RecReturn!灭菌失效期), "", Format(RecReturn!灭菌失效期, "yyyy-MM-dd")), _
                            IIf(IsNull(RecReturn!可用数量), "0", RecReturn!可用数量), _
                            IIf(IsNull(RecReturn!实际金额), "0", RecReturn!实际金额), _
                            IIf(IsNull(RecReturn!实际差价), "0", RecReturn!实际差价), _
                            IIf(IsNull(RecReturn!指导差价率), "0", RecReturn!指导差价率), _
                            IIf(mintUnit = 0, 1, RecReturn!换算系数), IIf(IsNull(RecReturn!批次), 0, RecReturn!批次), RecReturn!时价, RecReturn!库房分批, RecReturn!在用分批, IIf(IsNull(RecReturn!批准文号), "", RecReturn!批准文号)) Then

                            If .Row = .Rows - 1 Then .Rows = .Rows + 1 '只有当前行是最后一行时才新增行
                            .Row = .Row + 1
                        End If
                        
                        .Col = mBillCol.C_填写数量
                        RecReturn.MoveNext
                    Next
                
                    mshBill.Row = int点击行
                    
                    If mstr重复卫材 <> "" Then
                        MsgBox mstr重复卫材 & "列表中已经含有了！" & vbCrLf & "以上卫材不再添加！", vbInformation + vbOKOnly, gstrSysName
                        mstr重复卫材 = ""
                    End If
                
'                    If RecReturn.RecordCount = 1 Then
'
'                        SetColValue .Row, RecReturn!材料ID, "[" & RecReturn!编码 & "]" & RecReturn!名称, _
'                            IIf(IsNull(RecReturn!规格), "", RecReturn!规格), IIf(IsNull(RecReturn!产地), "", RecReturn!产地), _
'                            IIf(mintUnit = 0, RecReturn!散装单位, RecReturn!包装单位), _
'                            IIf(IsNull(RecReturn!售价), 0, RecReturn!售价), IIf(IsNull(RecReturn!批号), "", RecReturn!批号), _
'                            IIf(IsNull(RecReturn!效期), "", Format(RecReturn!效期, "yyyy-MM-dd")), _
'                            IIf(IsNull(RecReturn!灭菌失效期), "", Format(RecReturn!灭菌失效期, "yyyy-MM-dd")), _
'                            IIf(IsNull(RecReturn!可用数量), "0", RecReturn!可用数量), _
'                            IIf(IsNull(RecReturn!实际金额), "0", RecReturn!实际金额), _
'                            IIf(IsNull(RecReturn!实际差价), "0", RecReturn!实际差价), _
'                            IIf(IsNull(RecReturn!指导差价率), "0", RecReturn!指导差价率), _
'                            IIf(mintUnit = 0, 1, RecReturn!换算系数), IIf(IsNull(RecReturn!批次), 0, RecReturn!批次), RecReturn!时价, RecReturn!库房分批, RecReturn!在用分批, IIf(IsNull(RecReturn!批准文号), "", RecReturn!批准文号)
'                        .Col = mBillCol.C_填写数量
'                    End If
                End With
                RecReturn.Close
            End If
    End Select
End Sub

Private Sub mshbill_EditChange(curText As String)
    mblnChange = True
End Sub

Private Sub mshBill_EditKeyPress(KeyAscii As Integer)
    Dim strKey As String
    Dim intDigit As Integer
    
    With mshBill
        If .Col = mBillCol.C_填写数量 Or .Col = mBillCol.C_实际数量 Then
            strKey = .Text
            If strKey = "" Then
                strKey = .TextMatrix(.Row, .Col)
            End If
            Select Case .Col
                Case mBillCol.C_填写数量, mBillCol.C_实际数量
                    intDigit = IIf(mintUnit = 1, g_小数位数.obj_包装小数.数量小数, g_小数位数.obj_散装小数.数量小数)
            End Select
            
            If InStr(strKey, ".") <> 0 And Chr(KeyAscii) = "." Then   '只能存在一个小数点
                KeyAscii = 0
                Exit Sub
            End If
            
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
    If Row > 0 Then
        mshBill.SetRowColor CLng(Row), &HFFCECE, True
    End If
    If mblnEnter Then Exit Sub
    
    Call Local跟踪病人信息
    
    With mshBill
        Call SetInputFormat(.Row)
        If .Row <> .LastRow Then
        End If
        
        Select Case .Col
            Case mBillCol.C_材料
                .TxtCheck = False
                .MaxLength = 80
                '只在药名列才显示合计信息和库存数
                Call 显示合计金额
                Call 提示库存数
            Case mBillCol.C_填写数量, mBillCol.C_实际数量
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = "-.1234567890"
        End Select
    End With
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim rsDrug As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim i As Integer
    Dim int点击行 As Integer
    
    int点击行 = mshBill.Row
    
    
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
            
            Case mBillCol.C_材料
                If strKey <> "" Then
                    Dim RecReturn As Recordset
                    Dim sngLeft As Single
                    Dim sngTop As Single
                    
                    sngLeft = Me.Left + Pic单据.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                    sngTop = Me.Top + Me.Height - Me.ScaleHeight + Pic单据.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                    If sngTop + 3630 > Screen.Height Then
                        sngTop = sngTop - mshBill.MsfObj.CellHeight - 4530
                    End If
                    
                    Set RecReturn = FrmMulitSel.ShowSelect(Me, 2, cboStock.ItemData(cboStock.ListIndex), _
                                        cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), _
                                        strKey, sngLeft, sngTop, mshBill.MsfObj.CellWidth, mshBill.MsfObj.CellHeight, _
                                        IIf(mint领用明确批次 = 1, True, False), , , , , , , , , mlngModule, , mstrPrivs, IIf(mint领用明确批次 = 1, True, False), False)
                    
                    If RecReturn.RecordCount <= 0 Then
                        Cancel = True
                        Exit Sub
                    End If
                    
                    RecReturn.MoveFirst
                    For i = 1 To RecReturn.RecordCount
                        If SetColValue(.Row, RecReturn!材料ID, "[" & RecReturn!编码 & "]" & RecReturn!名称, _
                                IIf(IsNull(RecReturn!规格), "", RecReturn!规格), IIf(IsNull(RecReturn!产地), "", RecReturn!产地), _
                                IIf(mintUnit = 0, RecReturn!散装单位, RecReturn!包装单位), _
                                IIf(IsNull(RecReturn!售价), 0, RecReturn!售价), IIf(IsNull(RecReturn!批号), "", RecReturn!批号), _
                                IIf(IsNull(RecReturn!效期), "", Format(RecReturn!效期, "yyyy-MM-dd")), _
                                IIf(IsNull(RecReturn!灭菌失效期), "", Format(RecReturn!灭菌失效期, "yyyy-MM-dd")), _
                                IIf(IsNull(RecReturn!可用数量), "0", RecReturn!可用数量), _
                                IIf(IsNull(RecReturn!实际金额), "0", RecReturn!实际金额), _
                                IIf(IsNull(RecReturn!实际差价), "0", RecReturn!实际差价), _
                                IIf(IsNull(RecReturn!指导差价率), "0", RecReturn!指导差价率), _
                                IIf(mintUnit = 0, 1, RecReturn!换算系数), IIf(IsNull(RecReturn!批次), 0, RecReturn!批次), RecReturn!时价, RecReturn!库房分批, RecReturn!在用分批, IIf(IsNull(RecReturn!批准文号), "", RecReturn!批准文号)) Then
                            
                            If .Row = .Rows - 1 Then .Rows = .Rows + 1 '只有当前行是最后一行时才新增行
                            .Row = .Row + 1
                            
                            .Text = .TextMatrix(.Row, .Col)
                        Else
                            Cancel = True
                        End If
                        
                        RecReturn.MoveNext
                    Next
                    
                    mshBill.Row = int点击行
                    
                    If mstr重复卫材 <> "" Then
                        MsgBox mstr重复卫材 & "列表中已经含有了！" & vbCrLf & "以上卫材不再添加！", vbInformation + vbOKOnly, gstrSysName
                        mstr重复卫材 = ""
                    End If
                    
'                    If RecReturn.RecordCount = 1 Then
'                        If SetColValue(.Row, RecReturn!材料ID, "[" & RecReturn!编码 & "]" & RecReturn!名称, _
'                                IIf(IsNull(RecReturn!规格), "", RecReturn!规格), IIf(IsNull(RecReturn!产地), "", RecReturn!产地), _
'                                IIf(mintUnit = 0, RecReturn!散装单位, RecReturn!包装单位), _
'                                IIf(IsNull(RecReturn!售价), 0, RecReturn!售价), IIf(IsNull(RecReturn!批号), "", RecReturn!批号), _
'                                IIf(IsNull(RecReturn!效期), "", Format(RecReturn!效期, "yyyy-MM-dd")), _
'                                IIf(IsNull(RecReturn!灭菌失效期), "", Format(RecReturn!灭菌失效期, "yyyy-MM-dd")), _
'                                IIf(IsNull(RecReturn!可用数量), "0", RecReturn!可用数量), _
'                                IIf(IsNull(RecReturn!实际金额), "0", RecReturn!实际金额), _
'                                IIf(IsNull(RecReturn!实际差价), "0", RecReturn!实际差价), _
'                                IIf(IsNull(RecReturn!指导差价率), "0", RecReturn!指导差价率), _
'                                IIf(mintUnit = 0, 1, RecReturn!换算系数), IIf(IsNull(RecReturn!批次), 0, RecReturn!批次), RecReturn!时价, RecReturn!库房分批, RecReturn!在用分批, IIf(IsNull(RecReturn!批准文号), "", RecReturn!批准文号)) = False Then
'                            Cancel = True
'                            Exit Sub
'                        End If
'                        .Text = .TextMatrix(.Row, .Col)
'                    Else
'                        Cancel = True
'                    End If
                    Call 提示库存数
                End If
            Case mBillCol.C_跟踪病人
                
 
            Case mBillCol.C_填写数量, mBillCol.C_实际数量
                If .TextMatrix(.Row, .Col) = "" And strKey = "" Then
                    MsgBox "数量必须输入！", vbOKOnly + vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "数量必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Val(strKey) = 0 And mint编辑状态 <> 3 Then
                        MsgBox "数量不能为零,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If mint编辑状态 = 6 Then
                        If Abs(Val(strKey)) > Abs(Val(.TextMatrix(.Row, mBillCol.C_填写数量))) Then
                            MsgBox "冲销数量不能大于领用数量！", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                    
                    If .TextMatrix(.Row, 0) = "" Then Exit Sub
                    If Not CompareUsableQuantity(.Row, strKey) Then
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    '成本价的公式：     出库金额=数量*售价
                    '                  出库差价=出库金额*（实际差价/实际金额）
                    '                  if 实际金额=0 then  出库差价=出库金额*指导差价率
                    '                  购价（成本价）=（出库金额-出库差价）/数量
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    
                    strKey = Format(strKey, mFMT.FM_数量)
                    .Text = strKey
                    
                    If .TextMatrix(.Row, mBillCol.C_售价) <> "" Then
                        .TextMatrix(.Row, mBillCol.C_售价金额) = Format(.TextMatrix(.Row, mBillCol.C_售价) * strKey, mFMT.FM_金额)
                    End If
                    
                    If mint编辑状态 <> 6 Then
'                        Dim dbl差价 As Double, dbl购价 As Double, dbl成本金额 As Double
'                        Call 验证出库差价计算(cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.row, 0)), Val(.TextMatrix(.row, mBillCol.c_批次)), Val(.TextMatrix(.row, mBillCol.C_比例系数)), Val(.TextMatrix(.row, mBillCol.C_实际差价)), Val(.TextMatrix(.row, mBillCol.C_实际金额)), Val(Split(.TextMatrix(.row, mBillCol.C_指导差价率), "||")(0)) / 100, Val(strKey), Val(.TextMatrix(.row, mBillCol.C_售价金额)), dbl差价, dbl购价, dbl成本金额)
'                        .TextMatrix(.row, mBillCol.C_差价) = Format(dbl差价, mFMT.FM_金额)
                        .TextMatrix(.Row, mBillCol.C_采购价) = Format(Get成本价(Val(.TextMatrix(.Row, 0)), cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, mBillCol.c_批次))) * Val(.TextMatrix(.Row, mBillCol.c_比例系数)), mFMT.FM_成本价)
'                        .TextMatrix(.row, mBillCol.C_采购金额) = Format(dbl成本金额, mFMT.FM_金额)
'                    Else
'                        .TextMatrix(.row, mBillCol.C_采购金额) = Format(Val(.TextMatrix(.row, mBillCol.C_采购价)) * strKey, mFMT.FM_金额)
'                        .TextMatrix(.row, mBillCol.C_差价) = Format(Val(.TextMatrix(.row, mBillCol.C_售价金额)) - Val(.TextMatrix(.row, mBillCol.C_采购金额)), mFMT.FM_金额)
                    End If
                    .TextMatrix(.Row, mBillCol.C_采购金额) = Format(Val(.TextMatrix(.Row, mBillCol.C_采购价)) * strKey, mFMT.FM_金额)
                    .TextMatrix(.Row, mBillCol.C_差价) = Format(Val(.TextMatrix(.Row, mBillCol.C_售价金额)) - Val(.TextMatrix(.Row, mBillCol.C_采购金额)), mFMT.FM_金额)
                    
                    If .Col = mBillCol.C_填写数量 Then
                        .TextMatrix(.Row, mBillCol.C_实际数量) = strKey
                    End If
                End If
                显示合计金额
        End Select
    End With
End Sub

Private Function Check分批属性(ByVal intRow As Integer, ByVal int在用分批 As Integer, ByVal int库房分批) As Integer
    '功能：用来检查材料在当前库房是否分批
    '返回值：1-分批，0-不分批
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    gstrSQL = "Select Distinct 0 " & _
            "From 部门性质说明 " & _
            "Where ((工作性质 Like '发料部门') Or (工作性质 Like '制剂室')) And 部门id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex))
    If rsTemp.RecordCount = 0 Then
        Check分批属性 = IIf(int库房分批 = 1, 1, 0)
    Else
        Check分批属性 = IIf(int在用分批 = 1, 1, 0)
    End If

    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'从材料特性中取值并附给相应的列
Private Function SetColValue(ByVal intRow As Integer, ByVal lng材料ID As Long, _
        ByVal str材料 As String, ByVal str规格 As String, ByVal str产地 As String, _
        ByVal str单位 As String, ByVal num售价 As Double, ByVal str批号 As String, _
        ByVal str效期 As String, ByVal str灭菌失效期 As String, ByVal num可用数量 As Double, ByVal num实际金额 As Double, _
        ByVal num实际差价 As Double, ByVal num指导差价率 As Double, _
        ByVal num比例系数 As Double, ByVal lng批次 As Long, _
        ByVal int是否变价 As Integer, ByVal int库房分批 As Integer, ByVal int在用分批 As Integer, ByVal str批准文号 As String) As Boolean
    
        Dim intCount As Integer
        Dim intCol As Integer
        Dim dblPrice As Double
        Dim rsTemp As New Recordset
        Dim bln分批 As Boolean
        
    On Error GoTo ErrHandle
    SetColValue = False
    If Format(str灭菌失效期, "yyyy-mm-dd") < Format(sys.Currentdate, "yyyy-mm-dd") And Trim(str灭菌失效期) <> "" Then
       If MsgBox("卫材【" & str材料 & "(" & lng批次 & ")】已经过了灭菌失效期,是否还要领用！", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) <> vbYes Then
            Exit Function
       End If
    End If
    
    With mshBill
        .TextMatrix(intRow, mBillCol.C_分批属性) = Check分批属性(intRow, int在用分批, int库房分批)
        
        If int是否变价 = 1 Then
            gstrSQL = "" & _
                "   Select nvl(零售价,0)*" & num比例系数 & " as  分批售价,实际金额/实际数量* " & num比例系数 & " as 平均零售价" & _
                "   From 药品库存 " & _
                "   Where 库房id=[1]" & _
                "           and 药品id=[2]" & _
                "           and 性质=1 and 实际数量>0 and " & _
                "           nvl(批次,0)=[3]"
            
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex), lng材料ID, lng批次)
                        
            If rsTemp.EOF Then
                If mint领用明确批次 = 1 Then
                    MsgBox "时价卫材没有库存，不能出库，请检查！", vbOKOnly, gstrSysName
                    Exit Function
                Else
                    dblPrice = num售价 * num比例系数
                End If
            Else
                If Val(.TextMatrix(intRow, mBillCol.C_分批属性)) = 1 Then
                    dblPrice = rsTemp!分批售价
                Else
                    dblPrice = rsTemp!平均零售价
                End If
            End If
        End If
        
        Dim lngRow As Long
        For lngRow = 1 To .Rows - 1
            If lngRow <> intRow And .TextMatrix(lngRow, 0) <> "" Then
                If .TextMatrix(lngRow, 0) = lng材料ID And Val(.TextMatrix(lngRow, mBillCol.c_批次)) = lng批次 Then
                    If UBound(Split(mstr重复卫材, "，")) < 3 Then mstr重复卫材 = mstr重复卫材 & .TextMatrix(lngRow, mBillCol.C_材料) & "，"  '最多记录三个重复的卫材
                    'Call MsgBox("卫生材料【" & .TextMatrix(lngRow, mBillCol.C_材料) & "( " & lng批次 & ")】已经存在，不再添加！", vbOKOnly + vbInformation + vbDefaultButton2, gstrSysName)
                    Exit Function
                End If
            End If
        Next
        
        .TextMatrix(intRow, mBillCol.C_行号) = intRow
        .TextMatrix(intRow, 0) = lng材料ID
        .TextMatrix(intRow, mBillCol.C_材料) = str材料
        .TextMatrix(intRow, mBillCol.c_规格) = str规格
        .TextMatrix(intRow, mBillCol.C_产地) = str产地
        .TextMatrix(intRow, mBillCol.C_批准文号) = str批准文号
        .TextMatrix(intRow, mBillCol.c_单位) = str单位
        .TextMatrix(intRow, mBillCol.c_批号) = str批号
        .TextMatrix(intRow, mBillCol.C_效期) = Format(str效期, "yyyy-mm-dd")
        .TextMatrix(intRow, mBillCol.C_灭菌失效期) = Format(str灭菌失效期, "yyyy-mm-dd")
    
        .TextMatrix(intRow, mBillCol.C_售价) = Format(num售价 * num比例系数, mFMT.FM_零售价)
        .TextMatrix(intRow, mBillCol.C_可用数量) = Format(num可用数量, mFMT.FM_数量)
        .TextMatrix(intRow, mBillCol.C_实际差价) = num实际差价
        .TextMatrix(intRow, mBillCol.C_实际金额) = num实际金额
        .TextMatrix(intRow, mBillCol.C_指导差价率) = num指导差价率 & "||" & int是否变价 & "||" & int在用分批
        .TextMatrix(intRow, mBillCol.c_比例系数) = num比例系数
        .TextMatrix(intRow, mBillCol.c_批次) = lng批次
        If int是否变价 = 1 Then .TextMatrix(intRow, mBillCol.C_售价) = Format(dblPrice, mFMT.FM_零售价)
        
        gstrSQL = "Select 跟踪病人 From 材料特性 where 材料id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng材料ID)
        If rsTemp.EOF = False Then
            .TextMatrix(intRow, mBillCol.C_跟踪标志) = zlStr.Nvl(rsTemp!跟踪病人)
        End If
        Call CheckLapse(str效期)
    End With
    Call 提示库存数
    SetColValue = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub mshBill_KeyPress(KeyAscii As Integer)
    If mshBill.Col = C_跟踪病人 Then
         If KeyAscii = vbKeySpace Or KeyAscii = vbKeyBack Then
             With mshBill
                .TextMatrix(.Row, C_跟踪信息) = ""
                .TextMatrix(.Row, C_跟踪病人) = ""
             End With
         End If
    End If
End Sub
 

Private Sub mshBill_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        'If button = 1 Then
'            If mshBill.MouseRow = 0 Then
'              Call Local跟踪病人信息
'            End If
       ' End If
End Sub

Private Sub mshProvider_DblClick()
    mshProvider_KeyDown vbKeyReturn, 0
End Sub

Private Sub mshProvider_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    If KeyCode = vbKeyEscape Then
        mshProvider.Visible = False
        txtDraw.SetFocus
        txtDraw.SelStart = 0
        txtDraw.SelLength = Len(txtDraw.Text)
    End If
    
    If KeyCode = vbKeyReturn Then
        txtDraw.Text = mshProvider.TextMatrix(mshProvider.Row, 3)
        txtDraw.Tag = mshProvider.TextMatrix(mshProvider.Row, 0)
        mshProvider.Visible = False
        
        gstrSQL = "Select 工作性质, 部门id, 服务对象 From 部门性质说明 Where 部门id=[1] And 工作性质='护理'"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, Val(Me.txtDraw.Tag))
        If rsTemp.EOF Then
            gstrSQL = "Select 工作性质, 部门id, 服务对象 From 部门性质说明 Where 部门id=[1] And 工作性质='临床'"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, Val(Me.txtDraw.Tag))
            If rsTemp.EOF = False Then
                cmdDraw.Tag = "临床"
            Else
                cmdDraw.Tag = ""
            End If
        Else
            cmdDraw.Tag = "护理"
        End If
        
        If txtDrawPerson.Enabled Then txtDrawPerson.SetFocus
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mshProvider_LostFocus()
    If mshProvider.Visible Then
        mshProvider.Visible = False
    End If
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "PY" And stbThis.Tag <> "PY" Then
        Logogram stbThis, 0
        stbThis.Tag = Panel.Key
    ElseIf Panel.Key = "WB" And stbThis.Tag <> "WB" Then
        Logogram stbThis, 1
        stbThis.Tag = Panel.Key
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
    
    If txtNo.Locked = False Then
        If Trim(txtNo.Text) = "" Then
            ShowMsgBox "单据号不能为空"
            Exit Function
        End If
        
        If InStr(1, txtNo.Text, "'") <> 0 Then
            ShowMsgBox "单据号中不能含有非法字符"
            Exit Function
        End If
        If InStr(1, txtNo.Text, ";") <> 0 Then
            ShowMsgBox "单据号中不能含有非法字符"
            Exit Function
        End If
    End If
    
    If LenB(StrConv(txtNo.Text, vbFromUnicode)) > txtNo.MaxLength Then
        ShowMsgBox "单据号超长,最多能输入" & CInt(txtNo.MaxLength / 2) & "个汉字（最好不要汉字）或" & txtNo.MaxLength & "个字符!"
        txtNo.SetFocus
        Exit Function
    End If
    If InStr(1, txt摘要.Text, ";") <> 0 Then
        ShowMsgBox "摘要中不能输入分号"
        If txt摘要.Enabled Then txt摘要.SetFocus
        Exit Function
    End If
    
    With mshBill
        If .TextMatrix(1, 0) <> "" Then         '先判有否数据
            If Val(txtDraw.Tag) = 0 Then
                If Trim(txtDraw.Text) = "" Then
                    ShowMsgBox "领用部门不能为空！"
                    txtDraw.SetFocus
                    Exit Function
                Else
                    ShowMsgBox "没有你输入的领用部门！"
                    txtDraw.SetFocus
                    Exit Function
                End If
            End If
            
            If Trim(txtDrawPerson.Tag) = "" Then
                If MsgBox("你未选择相关的领用人,是否继续?", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo Then
                    If txtDrawPerson.Enabled Then txtDrawPerson.SetFocus
                    Exit Function
                End If
            End If
            
            If LenB(StrConv(txt摘要.Text, vbFromUnicode)) > txt摘要.MaxLength Then
                ShowMsgBox "摘要超长,最多能输入" & CInt(txt摘要.MaxLength / 2) & "个汉字或" & txt摘要.MaxLength & "个字符!"
                txt摘要.SetFocus
                Exit Function
            End If
        
            For intLop = 1 To .Rows - 1
                If Val(.TextMatrix(intLop, 0)) > 0 And Trim(.TextMatrix(intLop, mBillCol.C_材料)) = "" Then
                    MsgBox "第" & intLop & "行卫材的名称为空了，请检查！", vbInformation, gstrSysName
                    mshBill.SetFocus
                    .Row = intLop
                    .MsfObj.TopRow = intLop
                    .Col = mBillCol.C_材料
                    Exit Function
                End If
                
                If Trim(.TextMatrix(intLop, mBillCol.C_材料)) <> "" Then
                    If Trim(Trim(.TextMatrix(intLop, mBillCol.C_填写数量))) = "" Then
                        ShowMsgBox "第" & intLop & "行卫材的数量为空了，请检查！"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mBillCol.C_填写数量
                        Exit Function
                    End If
                    
                    If Trim(Trim(.TextMatrix(intLop, mBillCol.C_实际数量))) = "" Then
                        ShowMsgBox "第" & intLop & "行卫材的数量为空了，请检查！"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mBillCol.C_实际数量
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mBillCol.C_填写数量)) > 9999999999# Then
                        ShowMsgBox "第" & intLop & "行卫材的填写数量大于了数据库能够保存的" & vbCrLf & "最大范围9999999999，请检查！"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mBillCol.C_填写数量
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mBillCol.C_实际数量)) > 9999999999# Then
                        ShowMsgBox "第" & intLop & "行卫材的实际数量大于了数据库能够保存的" & vbCrLf & "最大范围9999999999，请检查！"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mBillCol.C_实际数量
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mBillCol.C_采购金额)) > 9999999999999# Then
                        ShowMsgBox "第" & intLop & "行卫材的成本金额大于了数据库能够保存的" & vbCrLf & "最大范围9999999999999，请检查！"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = IIf(.ColData(mBillCol.C_填写数量) = 4, mBillCol.C_填写数量, mBillCol.C_实际数量)
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mBillCol.C_售价金额)) > 9999999999999# Then
                        ShowMsgBox "第" & intLop & "行卫材的售价金额大于了数据库能够保存的" & vbCrLf & "最大范围9999999999999，请检查！"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = IIf(.ColData(mBillCol.C_填写数量) = 4, mBillCol.C_填写数量, mBillCol.C_实际数量)
                        Exit Function
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
    Dim lng入出类别ID As Long
    Dim chrNo As Variant
    Dim lng序号 As Long
    Dim lng库房ID As Long
    Dim lng领用部门id As Long
    Dim lng材料ID As Long
    Dim str批号 As String
    Dim lng批次 As Long
    Dim str产地 As String
    Dim str效期 As String
    Dim dbl填写数量 As Double
    Dim dbl成本价  As Double
    Dim dbl成本金额  As Double
    Dim dbl零价 As Double
    Dim dbl零售金额 As Double
    Dim dbl差价  As Double
    Dim str摘要 As String
    Dim str填制人 As String
    Dim str填制日期 As String
    Dim str审核人 As String
    Dim datAssessDate As String
    Dim str灭菌效期 As String
    Dim rsTemp As New Recordset
    Dim intRow As Integer
    Dim arrtemp As Variant
    Dim dbl申购数量 As Double
    Dim cllProc As Collection
    Dim n As Long
    
    SaveCard = False
    Set cllProc = New Collection
    
    On Error GoTo ErrHandle
     
    '在外面设置入出类别ID，主要是所有卫材都要用他
    gstrSQL = "" & _
        "   SELECT b.id " & _
        "   FROM 药品单据性质 a, 药品入出类别 b " & _
        "   Where a.类别id = b.ID " & _
        "           AND a.单据 = 35 " & _
        "           AND b.系数 = -1 " & _
        "           AND ROWNUM < 2"
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "取入出类别"
    
    'Call OpenRecordset(rsTemp, "取入出类别")
    If rsTemp.EOF Then
        MsgBox "没有设置卫材领用的出库类别，请在入出分类中设置！", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    lng入出类别ID = rsTemp.Fields(0)
    rsTemp.Close
    
    With mshBill
        chrNo = Trim(txtNo)
        lng库房ID = cboStock.ItemData(cboStock.ListIndex)
        
        If mint编辑状态 = 1 Or mint编辑状态 = 7 Then 'mbln单据增加 Or
            If chrNo <> "" Then
                If CheckNOExists(73, chrNo) Then Exit Function
            End If
        
            If chrNo = "" Then chrNo = sys.GetNextNo(73, lng库房ID)
            If IsNull(chrNo) Then Exit Function
        End If
        txtNo.Tag = chrNo
        
        lng领用部门id = txtDraw.Tag
        str摘要 = Trim(txt摘要.Text)
        str填制人 = Txt填制人
        str填制日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        str审核人 = Txt审核人
        If mint编辑状态 = 2 Or mint编辑状态 = 5 Or bln强制保存 = True Then        '修改
            gstrSQL = "zl_材料领用_Delete('" & mstr单据号 & "')"
            AddArray cllProc, gstrSQL
        End If
            
        '按药品ID顺序更新数据
        recSort.Sort = "药品id,批次,序号"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!行号
'        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                lng材料ID = .TextMatrix(intRow, 0)
                str产地 = .TextMatrix(intRow, mBillCol.C_产地)
                str批号 = .TextMatrix(intRow, mBillCol.c_批号)
                lng批次 = .TextMatrix(intRow, mBillCol.c_批次)
                str效期 = IIf(.TextMatrix(intRow, mBillCol.C_效期) = "", "", .TextMatrix(intRow, mBillCol.C_效期))
                str灭菌效期 = IIf(.TextMatrix(intRow, mBillCol.C_灭菌失效期) = "", "", .TextMatrix(intRow, mBillCol.C_灭菌失效期))
                dbl填写数量 = Round(Val(.TextMatrix(intRow, mBillCol.C_填写数量)) * Val(.TextMatrix(intRow, mBillCol.c_比例系数)), g_小数位数.obj_最大小数.数量小数)
                
                If mint编辑状态 = 3 Or mint编辑状态 = 5 Then '财务审核和审核时，如果启用自动分解后需要删除原始的再插入新的，而插入新的时分解可能出现填写数量大于实际数量情况，这个时候应该用实际数量判断库存是否足够
                    If Val(.TextMatrix(intRow, mBillCol.C_填写数量)) <> Val(.TextMatrix(intRow, mBillCol.C_实际数量)) Then
                        dbl填写数量 = Round(Val(.TextMatrix(intRow, mBillCol.C_实际数量)) * Val(.TextMatrix(intRow, mBillCol.c_比例系数)), g_小数位数.obj_最大小数.数量小数)
                    End If
                End If
                
                dbl申购数量 = Round(Val(.TextMatrix(intRow, mBillCol.C_申购数量)) * Val(.TextMatrix(intRow, mBillCol.c_比例系数)), g_小数位数.obj_最大小数.数量小数)
                
                dbl成本价 = Round(Val(.TextMatrix(intRow, mBillCol.C_采购价)) / Val(.TextMatrix(intRow, mBillCol.c_比例系数)), g_小数位数.obj_最大小数.成本价小数)
                dbl成本金额 = Round(Val(.TextMatrix(intRow, mBillCol.C_采购金额)), g_小数位数.obj_最大小数.金额小数)
                dbl零价 = Round(Val(.TextMatrix(intRow, mBillCol.C_售价)) / Val(.TextMatrix(intRow, mBillCol.c_比例系数)), g_小数位数.obj_最大小数.零售价小数)
                dbl零售金额 = Round(Val(.TextMatrix(intRow, mBillCol.C_售价金额)), g_小数位数.obj_最大小数.金额小数)
                dbl差价 = Round(Val(.TextMatrix(intRow, mBillCol.C_差价)), g_小数位数.obj_最大小数.金额小数)
                arrtemp = Split(.TextMatrix(intRow, mBillCol.C_跟踪信息) & "||", "|")
                
                lng序号 = intRow
                
                'Zl_材料领用_Insert
                gstrSQL = "zl_材料领用_INSERT("
                '  入出类别id_In In 药品收发记录.入出类别id%Type,
                gstrSQL = gstrSQL & "" & lng入出类别ID & ","
                '  No_In         In 药品收发记录.NO%Type,
                gstrSQL = gstrSQL & "'" & chrNo & "',"
                '  序号_In       In 药品收发记录.序号%Type,
                gstrSQL = gstrSQL & "" & lng序号 & ","
                '  库房id_In     In 药品收发记录.库房id%Type,
                gstrSQL = gstrSQL & "" & lng库房ID & ","
                '  对方部门id_In In 药品收发记录.对方部门id%Type,
                gstrSQL = gstrSQL & "" & lng领用部门id & ","
                '  材料id_In     In 药品收发记录.药品id%Type,
                gstrSQL = gstrSQL & "" & lng材料ID & ","
                '  批次_In       In 药品收发记录.批次%Type,
                gstrSQL = gstrSQL & "" & lng批次 & ","
                '  填写数量_In   In 药品收发记录.填写数量%Type,
                gstrSQL = gstrSQL & "" & dbl填写数量 & ","
                '  成本价_In     In 药品收发记录.成本价%Type,
                gstrSQL = gstrSQL & "" & dbl成本价 & ","
                '  成本金额_In   In 药品收发记录.成本金额%Type,
                gstrSQL = gstrSQL & "" & dbl成本金额 & ","
                '  零售价_In     In 药品收发记录.零售价%Type,
                gstrSQL = gstrSQL & "" & dbl零价 & ","
                '  零售金额_In   In 药品收发记录.零售金额%Type,
                gstrSQL = gstrSQL & "" & dbl零售金额 & ","
                '  差价_In       In 药品收发记录.差价%Type,
                gstrSQL = gstrSQL & "" & dbl差价 & ","
                '  领用人_In     In 药品收发记录.领用人%Type,
                gstrSQL = gstrSQL & "" & IIf(txtDrawPerson.Text = "", "NULL", "'" & txtDrawPerson.Text & "'") & ","
                '  填制人_In     In 药品收发记录.填制人%Type,
                gstrSQL = gstrSQL & "'" & str填制人 & "',"
                '  填制日期_In   In 药品收发记录.填制日期%Type,
                gstrSQL = gstrSQL & "to_date('" & str填制日期 & "','yyyy-mm-dd HH24:MI:SS'),"
                '  产地_In       In 药品收发记录.产地%Type := Null,
                gstrSQL = gstrSQL & "'" & str产地 & "',"
                '  批号_In       In 药品收发记录.批号%Type := Null,
                gstrSQL = gstrSQL & "'" & str批号 & "',"
                '  效期_In       In 药品收发记录.效期%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str效期 = "", "Null", "to_date('" & Format(str效期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '  灭菌效期_In   In 药品收发记录.灭菌效期%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str灭菌效期 = "", "Null", "to_date('" & Format(str灭菌效期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '  摘要_In       In 药品收发记录.摘要%Type := Null
                gstrSQL = gstrSQL & "'" & str摘要 & "',"
                If Val(arrtemp(0)) = 0 Then
                    '  病人id_In     In 材料领用信息.病人id%Type := Null,
                    gstrSQL = gstrSQL & "NULL,"
                    '  使用时间_In   In 材料领用信息.使用时间%Type := Null,
                    gstrSQL = gstrSQL & "NULL,"
                    '  条码_In       In 材料领用信息.条码%Type := Null
                    gstrSQL = gstrSQL & "NULL,"
                    '   申购数量_in
                    gstrSQL = gstrSQL & dbl申购数量 & ")"
                Else
                    '  病人id_In     In 材料领用信息.病人id%Type := Null,
                    gstrSQL = gstrSQL & "" & Val(arrtemp(0)) & ","
                    '  使用时间_In   In 材料领用信息.使用时间%Type := Null,
                    gstrSQL = gstrSQL & "" & IIf(Trim(arrtemp(1)) = "", "NULL", "to_date('" & Trim(arrtemp(1)) & "','yyyy-mm-dd')") & ","
                    '  条码_In       In 材料领用信息.条码%Type := Null
                    gstrSQL = gstrSQL & "'" & Trim(arrtemp(2)) & "',"
                    '   申购数量_in
                    gstrSQL = gstrSQL & dbl申购数量 & ")"
                End If
                AddArray cllProc, gstrSQL
            End If
            
            recSort.MoveNext
        Next
    End With
        
                
    Call ExecuteProcedureArrAy(cllProc, mstrCaption, True)
'    If Not 检查单价(20, txtNO.Tag) Then
'        gcnOracle.RollbackTrans
'        Exit Function
'    End If
    gcnOracle.CommitTrans
    
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    SaveCard = True
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub 显示合计金额()
    Dim curTotal As Double, Cur记帐金额 As Double, Cur记帐差价 As Double
    Dim intLop As Integer
    
    curTotal = 0: Cur记帐金额 = 0: Cur记帐差价 = 0:
    
    With mshBill
        For intLop = 1 To .Rows - 1
            curTotal = curTotal + Val(.TextMatrix(intLop, mBillCol.C_采购金额))
            Cur记帐金额 = Cur记帐金额 + Val(.TextMatrix(intLop, mBillCol.C_售价金额))
        Next
    End With
    
    Cur记帐差价 = Cur记帐金额 - curTotal
    lblPurchasePrice.Caption = "成本金额合计：" & Format(curTotal, mFMT.FM_金额)
    lblSalePrice.Caption = "售价金额合计：" & Format(Cur记帐金额, mFMT.FM_金额)
    lblDifference.Caption = "差价合计：" & Format(Cur记帐差价, mFMT.FM_金额)
End Sub

Private Sub 提示库存数()
    Dim rsTemp As New Recordset
    
    On Error GoTo ErrHandle
    With mshBill
        If .TextMatrix(.Row, mBillCol.C_材料) = "" Then
            stbThis.Panels(2).Text = ""
            Exit Sub
        End If
        If .TextMatrix(mshBill.Row, 0) = "" Then Exit Sub
        
        If Val(.TextMatrix(.Row, mBillCol.c_批次)) > 0 Then
            gstrSQL = "" & _
                "   Select 可用数量/" & .TextMatrix(.Row, mBillCol.c_比例系数) & " as  可用数量 " & _
                "   From 药品库存 " & _
                "   Where 库房id=[1]" & _
                "           and 药品id=[2]" & _
                "           and 性质=1 and " & _
                "           nvl(批次,0)=[3]"
        Else
            gstrSQL = "Select Sum(Nvl(可用数量, 0)) / " & .TextMatrix(.Row, mBillCol.c_比例系数) & " As 可用数量 " & _
                " From 药品库存 Where 库房id = [1] And 药品id = [2] And 性质 = 1 "
        End If
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mBillCol.c_批次)))
        
        If rsTemp.EOF Then
            .TextMatrix(.Row, mBillCol.C_可用数量) = 0
        Else
            .TextMatrix(.Row, mBillCol.C_可用数量) = IIf(IsNull(rsTemp.Fields(0)), 0, rsTemp.Fields(0))
        End If
        rsTemp.Close
        
        stbThis.Panels(2).Text = "该卫材当前库存数为[" & Format(.TextMatrix(.Row, mBillCol.C_可用数量), mFMT.FM_数量) & "]" & .TextMatrix(.Row, mBillCol.c_单位)
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtDraw_LostFocus()
    If txtDraw.Text = "" Then
        txtDraw.Tag = "0"
        Exit Sub
    End If
End Sub

Private Sub txtDraw_Validate(Cancel As Boolean)
    If txtDraw.Text = "" Then
        txtDraw.Tag = "0"
        Exit Sub
    End If
End Sub

Private Sub txtDrawPerson_Change()
    txtDrawPerson.Tag = ""
End Sub

Private Sub txtDrawPerson_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txtDrawPerson.Tag) <> "" Then
        OS.PressKey vbKeyTab
        Exit Sub
    End If
    If Trim(txtDrawPerson.Text) = "" Then
        OS.PressKey vbKeyTab
        Exit Sub
    End If
    If ShowSelect(Trim(txtDrawPerson.Text)) = False Then
        Exit Sub
    End If
    OS.PressKey vbKeyTab
End Sub

Private Sub txtDrawPerson_LostFocus()
    If txtDrawPerson.Tag = "" Then
        If Trim(txtDrawPerson.Text) <> "" Then
            If ShowSelect(Trim(txtDrawPerson.Text)) = False Then
                txtDrawPerson.Text = ""
                Exit Sub
            End If
        End If
    End If
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
    Dim int包装系数 As Integer
    Dim lng药品ID As Long
    Dim blnInput As Boolean
    
    '初始准备
    intNO = 68
    lng库房ID = Me.cboStock.ItemData(Me.cboStock.ListIndex)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txtIn.Text) = "" Then Exit Sub
    
    If Len(txtIn.Text) < 8 Then
        txtIn.Text = zlCommFun.GetFullNO(txtIn.Text, intNO, lng库房ID)
    End If
    
    '需要要清除现有单据内容
    For IntCheck = 1 To mshBill.Rows - 1
        If mshBill.TextMatrix(IntCheck, 0) <> "" Then
            Exit For
        End If
    Next
    If IntCheck <> mshBill.Rows Then
        If MsgBox("需要要清除现有单据内容，你是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        '处理药品单位改变
        mshBill.ClearBill
    End If
    
    gstrSQL = "select 收费细目id,执行科室id from 收费执行科室"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "查询存储库房")
    
    '提取该单据并清空表格（只允许提取正常单据，且非退货单）
    gstrSQL = "Select a.药品id As 材料id, '[' || c.编码 || ']' As 编码, '[' || c.编码 || ']' || Nvl(f.名称, c.名称) As 药品名称, c.名称 As 通用名, f.名称 As 商品名," & vbNewLine & _
                "       c.规格, a.产地, c.计算单位 As 零售单位, 1 As 零售系数, b.包装单位, b.换算系数, Nvl(a.批次, 0) As 批次, Nvl(c.是否变价, 0) As 时价," & vbNewLine & _
                "       Nvl(b.库房分批, 0) As 库房分批, Nvl(b.在用分批, 0) As 在用分批, b.最大效期, a.批号, a.效期, a.灭菌效期, b.最大效期, b.指导差价率, a.实际数量, d.可用数量," & vbNewLine & _
                "       d.实际金额, d.实际差价, e.现价, a.批准文号, Nvl(d.平均成本价, 0) As 平均成本价, a.供药单位id" & vbNewLine & _
                "From 药品收发记录 A, 材料特性 B, 收费项目目录 C, 药品库存 D, 收费价目 E, 收费项目别名 F" & vbNewLine & _
                "Where a.药品id = b.材料id And b.材料id = c.Id And b.材料id = d.药品id(+) And b.材料id = f.收费细目id(+) And f.性质(+) = 3 And f.码类(+) = 1 And" & vbNewLine & _
                "      b.材料id = e.收费细目id(+) And Sysdate >= e.执行日期(+) And Sysdate <= Nvl(e.终止日期(+), Sysdate) And d.库房id(+) = [2] And" & vbNewLine & _
                "      d.性质(+) = 1 And Nvl(a.批次, 0) = Nvl(d.批次, 0) And a.单据 = 15 And a.记录状态 = 1 And Nvl(a.发药方式, 0) = 0 And" & vbNewLine & _
                "      a.审核日期 Is Not Null And a.No = [1] And a.库房id + 0 = [2]" & GetPriceClassString("E") & vbNewLine & _
                "Order By a.序号"

    Set rsBill = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "[提取外购入库单]", txtIn.Text, Me.cboStock.ItemData(Me.cboStock.ListIndex))
             
    If rsBill.RecordCount = 0 Then
        MsgBox "没有找到该外购入库单据！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    With rsBill
        intRow = 1
        Do While Not .EOF
            lng药品ID = !材料ID
            rsTemp.Filter = " 收费细目id=" & lng药品ID & " and 执行科室id=" & lng库房ID
            If rsTemp.RecordCount = 0 Then
                MsgBox "材料[" & !药品名称 & "]未在" & cboStock.Text & "中设置存储属性，将不能领用！"
                blnInput = True
            End If
            
            If blnInput = False Then
                '导入计划单相当于都是按批次移库，需要在装入数据前，先检查库存
                If !实际数量 > !可用数量 Then
                    Select Case mint库存检查
                    Case 1
                        If MsgBox(!药品名称 & "库存不足，是否继续！", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            blnInput = True
                        End If
                    Case 2
                        MsgBox !药品名称 & "库存不足，将不能领用！", vbInformation, gstrSysName
                        blnInput = True
                    End Select
                End If
            End If
            
            '装入数据(SetColValue)
            If blnInput = False Then
                int包装系数 = IIf(mintUnit = 0, 1, !换算系数)
                If Not SetColValue(intRow, !材料ID, "[" & !编码 & "]" & !通用名, _
                   Nvl(!规格), Nvl(!产地), _
                   IIf(mintUnit = 0, !零售单位, !包装单位), _
                    Nvl(!现价, 0), Nvl(!批号), _
                    Nvl(!效期), IIf(IsNull(!灭菌效期), "", Format(!灭菌效期, "yyyy-MM-dd")), _
                    Nvl(!可用数量, 0), Nvl(!实际金额, 0), Nvl(!实际差价, 0), _
                    IIf(IsNull(!指导差价率), "0", !指导差价率), int包装系数, Nvl(!批次, 0), !时价, _
                    !库房分批, !在用分批, IIf(IsNull(!批准文号), "", !批准文号)) Then
                    mshBill.ClearBill
                    Exit Sub
                End If

                '填写数量、采购价、售价等列
                mshBill.TextMatrix(intRow, mBillCol.C_行号) = intRow
                mshBill.TextMatrix(intRow, mBillCol.C_填写数量) = Format(!实际数量 / int包装系数, mFMT.FM_数量)
                mshBill.TextMatrix(intRow, mBillCol.C_实际数量) = Format(!实际数量 / int包装系数, mFMT.FM_数量)
                mshBill.TextMatrix(intRow, mBillCol.C_采购价) = Format(!平均成本价 * int包装系数, mFMT.FM_成本价)
                mshBill.TextMatrix(intRow, mBillCol.C_采购金额) = Format(Val(mshBill.TextMatrix(intRow, mBillCol.C_采购价)) * Val(mshBill.TextMatrix(intRow, mBillCol.C_实际数量)), mFMT.FM_金额)
                mshBill.TextMatrix(intRow, mBillCol.C_售价金额) = Format(Val(mshBill.TextMatrix(intRow, mBillCol.C_售价)) * Val(mshBill.TextMatrix(intRow, mBillCol.C_实际数量)), mFMT.FM_金额)
                mshBill.TextMatrix(intRow, mBillCol.C_差价) = Format(Val(mshBill.TextMatrix(intRow, mBillCol.C_售价金额)) - mshBill.TextMatrix(intRow, mBillCol.C_采购金额), mFMT.FM_金额)

                intRow = intRow + 1
                mshBill.Rows = mshBill.Rows + 1
            End If
            blnInput = False
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
    txt摘要.Tag = ""
End Sub

Private Sub txt摘要_GotFocus()
    ImeLanguage True
    
    With txt摘要
        .SelStart = 0
        .SelLength = Len(txt摘要.Text)
    End With
End Sub

Private Sub txt摘要_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strKey As String
    If KeyCode = vbKeyReturn Then
        If mblnHave领用用途 = False Then
            OS.PressKey vbKeyTab: Exit Sub
        End If
        strKey = Trim(txt摘要)
        If txt摘要.Tag <> "" Then OS.PressKey vbKeyTab: Exit Sub
        If SelectItem(Me, txt摘要, strKey, "材料领用用途", "材料领用用途选择", True) = False Then
            OS.PressKey vbKeyTab
            Exit Sub
        End If
        KeyCode = 0
    End If
End Sub

Private Sub txt摘要_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt摘要, KeyAscii, m文本式
    If KeyAscii = Asc(";") Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt摘要_LostFocus()
    ImeLanguage False
End Sub

'与可用数量进行比较
Private Function CompareUsableQuantity(ByVal intRow As Integer, ByVal dbl填写数量 As Double) As Boolean
    Dim dblUsableQuantity As Double      '实际数量对应的组成数量
    Dim dbltotal As Double
    Dim vardrug As Variant, intLop As Integer
    
    'mint库存检查: 0-不检查;1-检查，不足提醒；2-检查，不足禁止
    
    CompareUsableQuantity = False
    
    If Not (mint编辑状态 = 5 Or mint编辑状态 = 3 Or mint编辑状态 = 6) Then '如果是核查，审核或者冲销必须检查数量
        If dbl填写数量 < 0 And mint领用明确批次 = 0 And Val(mshBill.TextMatrix(intRow, mBillCol.C_分批属性)) = 1 Then '负数领用必须明确批次
            MsgBox "分批材料负数领用必须明确批次，请到卫材系统参数->按批次领用卫生材料设置！", vbInformation, gstrSysName
            Exit Function
        End If
        If mint领用明确批次 = 0 Then
            CompareUsableQuantity = True
            Exit Function
        End If
    ElseIf mint编辑状态 = 6 And dbl填写数量 > 0 Then    '正数冲销相当于是入库 不需要检查库存
        CompareUsableQuantity = True
        Exit Function
    End If
    
    With mshBill
        If .TextMatrix(intRow, 0) = "" Then Exit Function
        
        If mint编辑状态 = 6 Then '冲销时直接检查实际库存，而不用可用存储
            dblUsableQuantity = Format(.TextMatrix(intRow, mBillCol.C_库存数量), mFMT.FM_数量)
        ElseIf mint编辑状态 = 2 Then
            If gSystem_Para.para_卫材填单下可用库存 = False Then
                '如果没有预减可用数量，则不算界面的原始数量
                dblUsableQuantity = Val(.TextMatrix(intRow, mBillCol.C_可用数量))
            Else
                dblUsableQuantity = Val(.TextMatrix(intRow, mBillCol.C_可用数量)) + Val(.TextMatrix(intRow, mBillCol.c_原始数量)) / Val(.TextMatrix(intRow, mBillCol.c_比例系数))
            End If
        ElseIf mint编辑状态 = 3 Then
            dblUsableQuantity = Val(.TextMatrix(intRow, mBillCol.C_可用数量)) + Val(.TextMatrix(intRow, mBillCol.c_原始数量)) / Val(.TextMatrix(intRow, mBillCol.c_比例系数))
        Else
            dblUsableQuantity = Val(Format(.TextMatrix(intRow, mBillCol.C_可用数量), mFMT.FM_数量))
        End If

        '加ABS是考虑可以负数冲销情况
        If mint库存检查 = 0 Then
            '0-不检查
        ElseIf mint库存检查 = 1 Then
            '1-检查，不足提醒
            If IIf(mint编辑状态 = 6, Abs(dbl填写数量), dbl填写数量) > dblUsableQuantity Then
                If MsgBox("你输入的数量“" & IIf(mint编辑状态 = 6, Abs(dbl填写数量), dbl填写数量) & "”大于了该卫材的" & IIf(mint编辑状态 = 6, "实际", "可用") & "库存数量“" & dblUsableQuantity & "”，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            End If
        ElseIf mint库存检查 = 2 Then
            '2-检查，不足禁止
            If IIf(mint编辑状态 = 6, Abs(dbl填写数量), dbl填写数量) > dblUsableQuantity Then
                MsgBox "你输入的数量“" & IIf(mint编辑状态 = 6, Abs(dbl填写数量), dbl填写数量) & "”大于了该卫材的" & IIf(mint编辑状态 = 6, "实际", "可用") & "库存数量“" & dblUsableQuantity & "”，请重输！", vbExclamation + vbOKOnly, gstrSysName
                Exit Function
            End If
        End If
    End With
    CompareUsableQuantity = True
End Function

'打印单据
Private Sub printbill()
    Dim strNo As String
    strNo = txtNo.Tag
    FrmBillPrint.ShowMe Me, glngSys, "zl1_bill_1717", mint记录状态, mintUnit, 1717, "卫材领用单", strNo
End Sub
Private Sub txtDraw_Change()
    With txtDraw
        .Text = UCase(.Text)
        .SelStart = Len(.Text)
    End With
    txtDraw.Tag = ""
    txtDrawPerson.Text = ""
    mblnChange = True
End Sub

Private Sub txtDraw_GotFocus()
    txtDraw.SelStart = 0
    txtDraw.SelLength = Len(txtDraw.Text)
End Sub

Private Sub txtDraw_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strProviderText As String, str站点限制 As String
    Dim rsTemp As New Recordset
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If mint编辑状态 = 3 Or mint编辑状态 = 5 Or mint编辑状态 = 4 Then Exit Sub
    If txtDraw.Tag <> "" Then OS.PressKey vbKeyTab: Exit Sub
    
    str站点限制 = GetDeptStationNode(cboStock.ItemData(cboStock.ListIndex))
    On Error GoTo ErrHandle
    With txtDraw
        If Trim(.Text) = "" Then Exit Sub
        strProviderText = GetMatchingSting(UCase(.Text))
        
        gstrSQL = "" & _
            " SELECT a.id,a.编码,a.简码,a.名称 " & _
            " FROM 部门表 a " & _
            " Where ( TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01' or a.撤档时间 is null ) " & _
            IIf(str站点限制 <> "", " And a.站点 = [3] ", "") & _
            "   And (a.简码 like [1] Or a.编码 like [1] or a.名称 like [1])"
        If mbln普通科室 Then
            '普通科室申领，只能选择自己所属的科室
            '刘兴宏:20060803
            '问题:8468
            gstrSQL = gstrSQL & " And a.ID in (Select 部门ID From 部门人员 where 人员id =[2]) "
        End If
            
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, strProviderText, UserInfo.Id, str站点限制)
     
        If rsTemp.EOF Then
            MsgBox "没有你输入的领用部门，请重输！", vbOKOnly + vbInformation, gstrSysName
            KeyCode = 0
            .SelStart = 0
            .SelLength = Len(.Text)
            .Tag = 0
            Exit Sub
        End If
        If rsTemp.RecordCount > 1 Then
            Set mshProvider.Recordset = rsTemp
            Dim intCol As Integer
            Dim intRow As Integer
            
            With mshProvider
                If .Visible = False Then .Visible = True
                .Redraw = False
                .SetFocus
                
                For intRow = 0 To .Rows - 1
                    .Row = intRow
                    For intCol = 0 To .Cols - 1
                        .Col = intCol
                        If .Row = 0 Then
                            .CellFontBold = True
                        Else
                            .CellFontBold = False
                        End If
                    Next
                Next
                .Font.Bold = False
                .FontFixed.Bold = True
                .ColWidth(0) = 0
                .ColWidth(1) = 800
                .ColWidth(2) = 1000
                .ColWidth(3) = 2500
                .Row = 1
                .TopRow = 1
                .Col = 0
                .ColSel = .Cols - 1
                
                .Top = txtDraw.Top + txtDraw.Height + 25
                .Left = cmdDraw.Left + cmdDraw.Width - .Width
                .Redraw = True
            End With
            SetObjMuchSelectHeigth Me, mshProvider, txtDraw
            mshProvider.TopRow = 1
            mshProvider.Row = 1
            mshProvider.ColSel = mshProvider.Cols - 1
            mshProvider.SelectionMode = flexSelectionByRow
            Exit Sub
        Else
            .Text = rsTemp!编码 & "-" & rsTemp!名称
            .Tag = rsTemp!Id
        End If
        
        gstrSQL = "Select 工作性质, 部门id, 服务对象 From 部门性质说明 Where 部门id=[1] And 工作性质='护理'"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, Val(Me.txtDraw.Tag))
        If rsTemp.EOF Then
            gstrSQL = "Select 工作性质, 部门id, 服务对象 From 部门性质说明 Where 部门id=[1] And 工作性质='临床'"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, Val(Me.txtDraw.Tag))
            If rsTemp.EOF = False Then
                cmdDraw.Tag = "临床"
            Else
                cmdDraw.Tag = ""
            End If
        Else
            cmdDraw.Tag = "护理"
        End If
        
        If txtDrawPerson.Enabled Then txtDrawPerson.SetFocus
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub SetObjMuchSelectHeigth(ByVal frmMain As Object, _
    ByVal mshSel As MSHFlexGrid, _
    ByVal objCtl As Object)
    
    
    '设置多选的高度和顶部
    Dim sngHeight As Single
    Dim sngminHeight As Single
    Dim intRow As Long
    Dim intMinRow As Long
    Dim sngTop As Single
    Dim sngFrmMinHeight As Single
    
   
    sngTop = objCtl.Top + objCtl.Height + 25
    intRow = mshSel.Row
    
    mshSel.Row = mshSel.Rows - 1
    sngHeight = ((mshSel.RowHeight(1) + 5) * (mshSel.Rows + 1))
    mshSel.Row = IIf(mshSel.Rows - 1 < 6, mshSel.Row, 6)
    sngminHeight = mshSel.CellTop + mshSel.CellHeight
    sngFrmMinHeight = IIf(frmMain.ScaleHeight - (sngTop) > 0, frmMain.ScaleHeight - sngTop, 0)
       
    If sngHeight > sngFrmMinHeight Then
        If sngFrmMinHeight - sngminHeight < 0 Then
            sngHeight = IIf(sngFrmMinHeight < 2000, 2000, sngFrmMinHeight)
        Else
            sngHeight = sngFrmMinHeight
        End If
        
    ElseIf sngHeight < sngminHeight Then
            sngHeight = sngminHeight
    End If
    mshSel.Height = sngHeight
End Sub
Private Function ShowSelect(ByVal strSeach As String) As Boolean
    '功能:提供各种输入选择
    '参数:intSelect:0-领用人
    
    Dim blnCancel As Boolean, strKey As String, strTittle As String, lngH As Long
    Dim objCtl As Object: Dim vRect As RECT
    Dim rsTemp  As ADODB.Recordset
    'zlDatabase.ShowSelect
    '功能：多功能选择器
    '参数：
    '     frmParent=显示的父窗体
    '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
    '     bytStyle=选择器风格
    '       为0时:列表风格:ID,…
    '       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
    '       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
    '     strTitle=选择器功能命名,也用于个性化区分
    '     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
    '     strSeek=当bytStyle<>2时有效,缺省定位的项目。
    '             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
    '             bytStyle=1时,可以是编码或名称
    '     strNote=选择器的说明文字
    '     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
    '     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
    '     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
    '     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
    '     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
    '     blnSearch=是否显示行号,并可以输入行号定位
    '返回：取消=Nothing,选择=SQL源的单行记录集
    '说明：
    '     1.ID和上级ID可以为字符型数据
    '     2.末级等字段不要带空值
    '应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    
    
    Set objCtl = txtDrawPerson
      
    strTittle = "人员选择"
    If strSeach = "" Then
        gstrSQL = "" & _
                "   Select ID, 编号,简码,姓名 From 人员表 a " & _
                "   Where   exists(select 1 from 部门人员 where 人员id=a.id and 部门id=[1]) " & _
                "           and (a.撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) and (a.站点=[3] or a.站点 is null) " & _
                "   order by 编号"
    Else
        gstrSQL = "" & _
                "   Select ID, 编号,简码,姓名 From 人员表 a " & _
                "   Where ((姓名) like [2] or  编号  like [2] or  简码  like  [2]) and (a.站点=[3] or a.站点 is null) " & _
                "           and exists(select 1 from 部门人员 where 人员id=a.id and 部门id=[1]) " & _
                "       and (a.撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null)" & _
                "   order by 编号"
    End If
    
    vRect = zlControl.GetControlRect(objCtl.hwnd)
    lngH = objCtl.Height
    strKey = GetMatchingSting(strSeach)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, Val(txtDraw.Tag), strKey, gstrNodeNo)
        
    If blnCancel = True Then
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    If rsTemp Is Nothing Then
        ShowMsgBox "没有满足条件的领用人,请检查!"
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    objCtl.Text = zlStr.Nvl(rsTemp!姓名)
    objCtl.Tag = zlStr.Nvl(rsTemp!姓名)
    
    ShowSelect = True
End Function

Private Function Local跟踪病人信息()
    '--------------------------------------------------------------------------------------
    '功能:定位跟踪病人信息
    '--------------------------------------------------------------------------------------
    Dim lngTemp As Long, lngPreCol As Long
    Dim i As Long
    
    With mshBill
        If Val(.TextMatrix(.Row, mBillCol.C_跟踪标志)) = 0 Then
            cmdSel.Visible = False
            Exit Function
        End If
                
        If cmdDraw.Tag <> "临床" And cmdDraw.Tag <> "护理" Then
            If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
                cmdSel.Visible = False
                Exit Function
            End If
        End If
        
        If mint编辑状态 > 2 And mint编辑状态 <> 7 Then
            If Trim(.TextMatrix(.Row, mBillCol.C_跟踪信息)) = "" Or Trim(.TextMatrix(.Row, mBillCol.C_跟踪信息)) = "||" Then
                cmdSel.Visible = False
                Exit Function
            End If
        End If
        lngPreCol = .Col
        mblnEnter = True
        
        .Redraw = False
        .ColData(mBillCol.C_跟踪病人) = 0
        .Col = mBillCol.C_跟踪病人
        
        lngTemp = .Left
        cmdSel.Left = .Left + .MsfObj.CellLeft + .MsfObj.CellWidth - cmdSel.Width + 30   ' lngTemp - cmdSel.Width + 30 '
        cmdSel.Top = .CellTop + .Top + 15
        cmdSel.Height = .RowHeight(.Row) ' .MsfObj.CellHeight
        .Col = lngPreCol
        If .MsfObj.ColIsVisible(mBillCol.C_跟踪病人) = True And .MsfObj.RowIsVisible(.Row) = True Then
            cmdSel.Visible = True
        Else
            cmdSel.Visible = False
        End If
        
        .Redraw = True
        mblnEnter = False
    End With
End Function

Private Sub SetInputFormat(ByVal intRow As Integer)
    '--------------------------------------------------------------------------------------------------------
    '功能:设置当前行的编辑格式
    '参数:introw-当前行
    '返回:
    '编制:刘兴宏
    '日期:2007/08/21
    '--------------------------------------------------------------------------------------------------------
    
    With mshBill
    
        '1.新增；2、修改；3、验收；4、查看；5、修改发票；6、冲销；
        If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
            If Val(.TextMatrix(intRow, mBillCol.C_跟踪标志)) = 1 And cmdDraw.Tag = "临床" Then
                .ColData(mBillCol.C_跟踪病人) = 0
            Else
                .ColData(mBillCol.C_跟踪病人) = 0
            End If
        Else
             .ColData(mBillCol.C_跟踪病人) = 0
        End If
    End With
End Sub

Private Sub SetSortRecord()
    Dim n As Integer
    
    If mshBill.Rows < 2 Then Exit Sub
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
        
        For n = 1 To mshBill.Rows - 1
            If mshBill.TextMatrix(n, 0) <> "" Then
                .AddNew
                !行号 = n
                !序号 = IIf(Val(mshBill.TextMatrix(n, mBillCol.C_序号)) = 0, n, Val(mshBill.TextMatrix(n, mBillCol.C_序号)))
                !药品id = Val(mshBill.TextMatrix(n, 0))
                !批次 = Val(mshBill.TextMatrix(n, mBillCol.c_批次))
                
                .Update
            End If
        Next
        
    End With
End Sub
