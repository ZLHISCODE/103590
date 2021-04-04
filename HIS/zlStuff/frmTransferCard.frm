VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmTransferCard 
   Caption         =   "卫材移库单"
   ClientHeight    =   6855
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11130
   Icon            =   "frmTransferCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   11130
   StartUpPosition =   2  '屏幕中心
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh产地 
      Height          =   2175
      Left            =   2520
      TabIndex        =   35
      Top             =   1200
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3836
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
   Begin VB.CommandButton cmdRequestTransfer 
      Caption         =   "按申购单移库(&T)"
      Height          =   350
      Left            =   3840
      TabIndex        =   34
      Top             =   6000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdExpend 
      Caption         =   "自动分解(&A)"
      Height          =   350
      Left            =   4950
      TabIndex        =   7
      Top             =   5490
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton cmdAllSel 
      Caption         =   "全冲(&A)"
      Height          =   350
      Left            =   6180
      TabIndex        =   30
      Top             =   5490
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdAllCls 
      Caption         =   "全清(&L)"
      Height          =   350
      Left            =   7500
      TabIndex        =   29
      Top             =   5490
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   3720
      TabIndex        =   12
      Top             =   5137
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "查找(&F)"
      Height          =   350
      Left            =   2040
      TabIndex        =   11
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   10
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   7560
      TabIndex        =   9
      Top             =   5040
      Width           =   1100
   End
   Begin VB.PictureBox Pic单据 
      BackColor       =   &H80000004&
      Height          =   4965
      Left            =   0
      ScaleHeight     =   4905
      ScaleWidth      =   11655
      TabIndex        =   13
      Top             =   0
      Width           =   11715
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
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "导入记帐单:F3"
         Top             =   150
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
         Height          =   360
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   135
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.TextBox txtNO 
         Height          =   300
         IMEMode         =   2  'OFF
         Left            =   9930
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   165
         Width           =   1425
      End
      Begin VB.ComboBox cboEnterStock 
         Height          =   300
         Left            =   9240
         TabIndex        =   3
         Text            =   "cboEnterStock"
         Top             =   600
         Width           =   2115
      End
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   2805
         Left            =   195
         TabIndex        =   4
         Top             =   950
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
         TabIndex        =   6
         Top             =   4080
         Width           =   10410
      End
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   1320
         TabIndex        =   1
         Text            =   "cboStock"
         Top             =   585
         Width           =   2745
      End
      Begin VB.Label lblDifference 
         AutoSize        =   -1  'True
         Caption         =   "差价合计:"
         Height          =   180
         Left            =   4920
         TabIndex        =   27
         Top             =   3840
         Width           =   810
      End
      Begin VB.Label lblSalePrice 
         AutoSize        =   -1  'True
         Caption         =   "售价金额合计:"
         Height          =   180
         Left            =   2040
         TabIndex        =   26
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "成本金额合计:"
         Height          =   180
         Left            =   240
         TabIndex        =   25
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label Txt审核人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7950
         TabIndex        =   23
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label Txt审核日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   10050
         TabIndex        =   22
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt填制日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2940
         TabIndex        =   21
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt填制人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   20
         Top             =   4440
         Width           =   915
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
         TabIndex        =   19
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
         TabIndex        =   5
         Top             =   4155
         Width           =   650
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "卫生材料移库单"
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
         TabIndex        =   18
         Top             =   120
         Width           =   11535
      End
      Begin VB.Label LblStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "移出库房(&S)"
         Height          =   180
         Left            =   240
         TabIndex        =   0
         Top             =   660
         Width           =   990
      End
      Begin VB.Label Lbl填制人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "填制人"
         Height          =   180
         Left            =   300
         TabIndex        =   17
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label Lbl填制日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "填制日期"
         Height          =   180
         Left            =   2160
         TabIndex        =   16
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label Lbl审核人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审核人"
         Height          =   180
         Left            =   7365
         TabIndex        =   15
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label Lbl审核日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审核日期"
         Height          =   180
         Left            =   9240
         TabIndex        =   14
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label LblEnterStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "移入库房(&I)"
         Height          =   180
         Left            =   8040
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
            Picture         =   "frmTransferCard.frx":014A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":0364
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":057E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":0798
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":09B2
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":0BCC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":0DE6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":1000
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
            Picture         =   "frmTransferCard.frx":121A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":1434
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":164E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":1868
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":1A82
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":1C9C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":1EB6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":20D0
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   28
      Top             =   6492
      Width           =   11124
      _ExtentX        =   19632
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmTransferCard.frx":22EA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13282
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmTransferCard.frx":2B7E
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmTransferCard.frx":3080
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
   Begin VB.CommandButton CmdSave 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6000
      TabIndex        =   8
      Top             =   5025
      Width           =   1335
   End
   Begin VB.Label lblCode 
      Caption         =   "材料"
      Height          =   255
      Left            =   3240
      TabIndex        =   24
      Top             =   5160
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmTransferCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mint编辑状态 As Integer             '1.新增；2、修改；3、验收；4、查看；5,6-冲销,10-发送,11-从入库单读取数据

Private mstr单据号 As String                '具体的单据号;
Private mint记录状态 As Integer             '1:正常记录;2-冲销记录;3-已经冲销的原记录
Private mblnSuccess As Boolean              '只要有一张成功，即为True，否则为False
Private mblnSave As Boolean                 '是否存盘和审核   TURE：成功。
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEdit As Boolean                 '是否可以修改
Private mblnChange As Boolean               '是否进行过编辑
Private mbln申领单 As Boolean               '是否是申领单，如果是则允许执行自动分解的功能
Private mbln明确批次 As Boolean             '是否明确批次，仅对申领单有效
Private mbln移库明确批次 As Boolean         '是否明确批次，仅对移库单有效

Private mint库存检查 As Integer             '表示卫材出库时是否进行库存检查：0-不检查;1-检查，不足提醒；2-检查，不足禁止
Private mcolUsedCount As Collection         '已使用的数量集合
Private mstrEnterSQL As String
Private mblnNoClick As Long
Private mblnCostView As Boolean                 '查看成本价 true-允许查看 false-不允许查看
Private mblnUpdate As Boolean               '表示是否已根据最新价格更新单据内容
Private Const mstrCaption As String = "卫材移库单"
Private mstr核查人 As String                '申领单据使用，记录申领核查人
Private mstr核查日期 As String              '申领单据使用，记录申领核查日期
Private mbln申领核查 As Boolean             '申领单据使用，记录申领是否需要核查流程 true-需要 false-不需要

Private mbln分批卫材批号产地控制 As Boolean  '是否检查分批卫材批号产地是否录入

Private mstrRequestNO As String     '按申购单移库NO ，空代表不按照申购单方式移库，否则按照申购单移库
'----------------------------------------------------------------------------------------------------------
'刘兴宏:增加小数位数的格式串
'修改:2007/03/06
Private mFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------
Private mbln仅显示有库存物资  As Boolean

Dim mstrPrivs As String                     '权限
Private mintParallelRecord As Integer       '对于新增后单据并行执行的处理： 1、代表正常情况；2、已经删除的记录；3、已经审核的记录
Private mintBatchNoLen As Integer           '数据库中批号定义长度

Private mstrTime_Start As String                        '进入单据编辑界面时，待编辑单据的最大修改时间
Private mstrTime_End As String                        '此刻该编辑单据的最大修改时间
Private mblnFirst As Boolean
Private mint移库处理流程 As Integer                    '1-需要备药、发送、接收这一过程  0-不需要这一过程
Private mstr入库单号 As String
Private mstr重复卫材 As String '记录重复的卫材

Private recSort As ADODB.Recordset          '按药品ID排序的专用记录集

'=========================================================================================
Private Const mlngModule = 1716

Private mbln单据增加    As Boolean          '进入时单据号累加1
Private mintUnit  As Integer                '显示单位:0-散装单位,1-包装单位
Private mint冲销方式 As Integer             '0－正常冲销方式；1－产生冲销申请单据；2－审核已产生的冲销申请单据

Private Enum mBillCol
     C_行号 = 1
     C_材料 = 2
     c_序号 = 3
     c_规格 = 4
     C_库房分批 = 5
     C_最大效期 = 6
     C_可用数量 = 7
     C_指导差价率 = 8
     C_实际金额 = 9
     C_实际差价 = 10
     C_比例系数 = 11
     c_批次 = 12
     C_产地 = 13
     C_批准文号 = 14
     c_单位 = 15
     c_批号 = 16
     C_效期 = 17
     C_一次性材料 = 18
     C_灭菌效期 = 19
     C_灭菌日期 = 20
     C_灭菌失效期 = 21
     C_填写数量 = 22
     C_实际数量 = 23
     c_原始数量 = 24
     C_采购价 = 25
     C_采购金额 = 26
     C_售价 = 27
'     C_售价金额 = 28
'
'     C_差价 = 29
End Enum
Private mconintcol售价金额 As Integer
Private mconintcol差价 As Integer

Private Const mBillCols  As Integer = 30              '总列数
Private mlng出库库房 As Long
Private mlngPreEnterId As Long      '上次移入库房
Private mlngPreStockId As Long  '上次移出库房

Private Function Auto处理移库流程() As Boolean
    '自动处理移库流程 1－备料 2－发送 3－接收
    
    On Error GoTo ErrHandle
    
    If Not 检查单价(19, txtNO.Tag, False) And Not mblnUpdate Then
        MsgBox "有记录未使用最新价格，程序将自动完成更新（售价、成本价、售价金额、成本金额、差价），更新后请检查！", vbInformation, gstrSysName
        Call RefreshBill
        mblnUpdate = True
        Exit Function
    End If
        
    If Not 材料单据审核(Txt填制人.Caption) Then Exit Function
    
    '2-
    If Not ValidData Then Exit Function
    If Not CheckStock Then Exit Function
    
    '先删除申领单，再依据当前数据产生移库单
    If Not SaveCard(True) Then
        gcnOracle.RollbackTrans: Exit Function
    End If
    
    '备料
    gstrSQL = "zl_材料移库_Prepare('" & txtNO.Tag & "','" & UserInfo.用户名 & "')"
    zlDatabase.ExecuteProcedure gstrSQL, mstrCaption
                    
    '发送（下出库库房的材料可用库存）
    gstrSQL = "zl_材料移库_Prepare('" & txtNO.Tag & "')"
    zlDatabase.ExecuteProcedure gstrSQL, mstrCaption
       
    '3-
    If SaveCheck() = True Then
        If IIf(Val(zlDatabase.GetPara("审核打印", glngSys, mlngModule, "0")) = 1, 1, 0) = 1 Then
            '打印
            If zlStr.IsHavePrivs(mstrPrivs, "单据打印") Then
                printbill
            End If
        End If
        Unload Me
    Else
        GoTo ErrHandle
    End If
    
    Auto处理移库流程 = True
    Exit Function
ErrHandle:
    Auto处理移库流程 = False
End Function

'=========================================================================================


'检查数据依赖性
Private Function GetDepend() As Boolean
    Dim strMsg As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHand
    GetDepend = False
    With rsTemp
        '检查卫材入出类别是否完整
        strMsg = "没有设置卫材移库的入库及出库类别，请在入出分类中设置！"
        
        gstrSQL = "" & _
            "   SELECT B.Id,B.系数 " & _
            "   FROM 药品单据性质 A, 药品入出类别 B " & _
            "   Where A.类别id = B.ID  AND A.单据 = 34"
            
        zlDatabase.OpenRecordset rsTemp, gstrSQL, "卫材移库管理"
        
        If .RecordCount = 0 Then GoTo ErrHand
        .Filter = "系数=1"
        If .RecordCount = 0 Then
            .Filter = 0
            strMsg = "没有设置卫材移库的入库类别，请在入出分类中设置！"
            GoTo ErrHand
        End If
        .Filter = "系数=-1"
        If .RecordCount = 0 Then
            .Filter = 0
            strMsg = "没有设置卫材移库的出库类别，请在入出分类中设置！"
            GoTo ErrHand
        End If
        .Filter = 0
        .Close
    End With
    GetDepend = True
    Exit Function
ErrHand:
    MsgBox strMsg, vbInformation, gstrSysName
    rsTemp.Close
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Sub ShowCard(frmMain As Form, ByVal str单据号 As String, ByVal int编辑状态 As Integer, _
    Optional int记录状态 As Integer = 1, Optional strPrivs As String, Optional blnSuccess As Boolean = False, Optional int冲销方式 As Integer = 0)
    Dim strReg As String
    
    mblnSave = False
    mblnSuccess = False
    
    mstr入库单号 = ""
    mstr单据号 = ""
    If int编辑状态 = 11 Then
        mstr入库单号 = str单据号
    Else
        mstr单据号 = str单据号
    End If
    
    mint编辑状态 = int编辑状态
    mint记录状态 = int记录状态
    mint冲销方式 = int冲销方式
    
    mblnSuccess = blnSuccess
    mblnChange = False
    mintParallelRecord = 1
    mstrPrivs = strPrivs
    
    Set mfrmMain = frmMain
    If Not GetDepend Then Exit Sub
    mblnCostView = zlStr.IsHavePrivs(mstrPrivs, "查看成本价")
    
    '没有成本价权限将售价金额和差价顺序换一下，防止不可见列在最后被拖出来
    mconintcol售价金额 = IIf(mblnCostView = False, 29, 28)
    mconintcol差价 = IIf(mblnCostView = False, 28, 29)
    
    Call GetRegInFor(g私有模块, "卫材移库管理", "单据号累加", strReg)
    mbln单据增加 = IIf(strReg = "", True, Val(strReg) = 1)
    
    If mint编辑状态 = 1 Or mint编辑状态 = 11 Then
        
        mblnEdit = True

        txtNO.Locked = True
        txtNO.TabStop = True

        txtNO = mstr单据号
        txtNO.Tag = txtNO.Text
    ElseIf mint编辑状态 = 2 Then
        mblnEdit = True
    ElseIf mint编辑状态 = 3 Then
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
        mblnEdit = False
        If mint冲销方式 = 0 Then '正常冲销
            CmdSave.Caption = "冲销(&O)"
        ElseIf mint冲销方式 = 1 Then    '申请冲销
            CmdSave.Caption = "申请冲销(&O)"
        ElseIf mint冲销方式 = 2 Then    '审核申请冲销单据
            CmdSave.Caption = "审核冲销(&O)"
        End If
        If mint冲销方式 = 2 Then
            cmdAllSel.Visible = False
            cmdAllCls.Visible = False
        Else
            cmdAllSel.Visible = True
            cmdAllCls.Visible = True
        End If
        
    ElseIf mint编辑状态 = 10 Then
        mblnEdit = False
        CmdSave.Caption = "发送(&S)"
        CmdSave.Visible = True
    End If
    
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    
    Me.Show vbModal, frmMain
    blnSuccess = mblnSuccess
    str单据号 = mstr单据号
    
End Sub

Private Sub cboEnterStock_Click()
    If mblnNoClick Then Exit Sub
    If cboEnterStock.ListIndex >= 0 Then mlngPreEnterId = cboEnterStock.ItemData(cboEnterStock.ListIndex)
End Sub

Private Sub cboEnterStock_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If cboEnterStock.ListCount = 0 Then Call zlControl.ControlSetFocus(mshBill): Exit Sub
    
    If cboEnterStock.ListIndex >= 0 Then
        If mlngPreEnterId = cboEnterStock.ItemData(cboEnterStock.ListIndex) Then
            OS.PressKey vbKeyTab
            'Call zlControl.ControlSetFocus(mshBill, True)
            Exit Sub
        End If
    End If
    If Select部门选择器(Me, cboEnterStock, Trim(cboEnterStock.Text), "", False, mstrEnterSQL) = False Then
        Exit Sub
    End If
    If cboEnterStock.ListIndex >= 0 Then
        mlngPreEnterId = cboEnterStock.ItemData(cboEnterStock.ListIndex)
    End If
End Sub

Private Sub cboEnterStock_LostFocus()
    Dim i As Long
    If cboEnterStock.ListCount = 0 Then Exit Sub
    If cboEnterStock.ListIndex < 0 Then
        For i = 0 To cboEnterStock.ListCount - 1
            If mlngPreEnterId = cboEnterStock.ItemData(i) Then
                mblnNoClick = True
                cboEnterStock.ListIndex = i: Exit For
            End If
        Next
    End If
    mblnNoClick = False
End Sub

Private Sub cboEnterStock_Validate(Cancel As Boolean)
    Dim i As Integer
    
    With cboEnterStock
        If .ListCount = 0 Then Exit Sub
        If .ListIndex < 0 Then Exit Sub
        If .ListIndex <> Val(.Tag) Then
            For i = 1 To mshBill.Rows - 1
                If mshBill.TextMatrix(i, 0) <> "" Then
                    Exit For
                End If
            Next
            If i <> mshBill.Rows Then
                If MsgBox("如果改变移入库房，有可能要改变相应卫材的单位和数量，" & vbCrLf & "且要清除现有单据内容，你是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    '处理卫材单位改变
                    cboEnterStock.Tag = .ListIndex
                    mshBill.ClearBill
                Else
                    .ListIndex = Val(.Tag)
                End If
            Else
                .Tag = .ListIndex
            End If
        End If
        
    End With
End Sub

Private Sub cboStock_Change()
    mblnChange = True
End Sub

Private Sub cboStock_Click()
    Dim rsTemp As New ADODB.Recordset
    If mblnNoClick Then Exit Sub
    If cboEnterStock.ListIndex >= 0 Then mlngPreEnterId = cboEnterStock.ItemData(cboEnterStock.ListIndex)
    
    '检查并装入移入库房
    err = 0: On Error Resume Next
    Set rsTemp = ReturnSQL(cboStock.ItemData(cboStock.ListIndex), mstrCaption, True, mstrEnterSQL, 1716)
    With rsTemp
        cboEnterStock.Clear
        Do While Not .EOF
            cboEnterStock.AddItem !名称
            cboEnterStock.ItemData(cboEnterStock.NewIndex) = !Id
            If mint编辑状态 = 11 Then
                If Val(zlStr.NVL(!Id)) = mfrmMain.cboEnterStock.ItemData(mfrmMain.cboEnterStock.ListIndex) Then
                    cboEnterStock.ListIndex = cboEnterStock.NewIndex
                End If
            End If
            .MoveNext
        Loop
        If cboEnterStock.ListIndex < 0 Then cboEnterStock.ListIndex = 0
        If mint编辑状态 <> 11 Then
            If cboEnterStock.ListCount <> 0 Then cboEnterStock.ListIndex = Val(cboEnterStock.Tag)
        End If
    End With
    
    mint移库处理流程 = IIf(Val(zlDatabase.GetPara("移库流程", glngSys, mlngModule, "0", , , , cboStock.ItemData(cboStock.ListIndex))) = 1, 1, 0)
    
    mint库存检查 = Get出库检查(cboStock.ItemData(cboStock.ListIndex))
End Sub

Private Sub cboStock_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cboStock.ListCount = 0 Then OS.PressKey vbKeyTab: Exit Sub
    
    If cboStock.ListIndex >= 0 Then
        If mlngPreStockId = cboStock.ItemData(cboStock.ListIndex) Then
           OS.PressKey vbKeyTab
            Exit Sub
        End If
    End If
    
    If Select部门选择器(Me, cboStock, Trim(cboStock.Text), "V,K,W", Not zlStr.IsHavePrivs(mstrPrivs, "所有库房")) = False Then
        Exit Sub
    End If
    If cboStock.ListIndex >= 0 Then
        mlngPreStockId = cboStock.ItemData(cboStock.ListIndex)
    End If
End Sub

Private Sub cboStock_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim i As Integer
        Dim blnreturn As Boolean
        blnreturn = False
        cboStock_Validate blnreturn
        If blnreturn = True Then Exit Sub
        
        OS.PressKey (vbKeyTab)
    End If
    
End Sub

Private Sub cboEnterStock_KeyPress(KeyAscii As Integer)
    Dim blnreturn As Boolean

    If KeyAscii <> 13 Then Exit Sub
    blnreturn = False
    cboEnterStock_Validate blnreturn
    If blnreturn = True Then Exit Sub

    With mshBill
        .Row = 1
        .Col = mBillCol.C_材料
    End With
    zlControl.ControlSetFocus mshBill, True
End Sub

Private Sub cboStock_LostFocus()
    Dim i As Long
    If cboStock.ListCount = 0 Then Exit Sub
    If cboStock.ListIndex < 0 Then
        For i = 0 To cboStock.ListCount - 1
            If mlngPreStockId = cboStock.ItemData(i) Then
                mblnNoClick = True
                cboStock.ListIndex = i: Exit For
            End If
        Next
    End If
    mblnNoClick = False
End Sub

Private Sub cboStock_Validate(Cancel As Boolean)
    Dim i As Integer
    
    With cboStock
        If .ListIndex < 0 Then Exit Sub
        If .ListIndex <> mintcboIndex Then
            For i = 1 To mshBill.Rows - 1
                If mshBill.TextMatrix(i, 0) <> "" Then
                    Exit For
                End If
            Next
            If i <> mshBill.Rows Then
                If MsgBox("如果改变移出库房，有可能要改变相应卫材的单位，" & vbCrLf & "且要清除现有单据内容，你是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
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
                .TextMatrix(intRow, mconintcol售价金额) = Format(0, mFMT.FM_金额)
                .TextMatrix(intRow, mconintcol差价) = Format(0, mFMT.FM_金额)
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
                .TextMatrix(intRow, mconintcol售价金额) = Format(.TextMatrix(intRow, mBillCol.C_填写数量) * .TextMatrix(intRow, mBillCol.C_售价), mFMT.FM_金额)
                .TextMatrix(intRow, mconintcol差价) = Format(.TextMatrix(intRow, mconintcol售价金额) - .TextMatrix(intRow, mBillCol.C_采购金额), mFMT.FM_金额)
            End If
        Next
    End With
    Call 显示合计金额
End Sub

Private Sub cmdCancel_Click()
    Unload Me
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
        
        cmdRequestTransfer.Left = txtCode.Left + txtCode.Width + (cmdFind.Left - cmdHelp.Left - cmdHelp.Width)
    Else
        FindRownew mshBill, mBillCol.C_材料, txtCode.Text, True
        lblCode.Visible = False
        txtCode.Visible = False
        
        cmdRequestTransfer.Left = cmdFind.Left + cmdFind.Width + (cmdFind.Left - cmdHelp.Left - cmdHelp.Width)
    End If
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Sub cmdRequestTransfer_Click()
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
    
    If cboEnterStock.ListCount = 0 Then  '无移入库房
        MsgBox "移入库房不能为空！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If Val(cboEnterStock.ItemData(cboEnterStock.ListIndex)) = 0 Then
        MsgBox "移入库房不能为空！", vbInformation, gstrSysName
        cboEnterStock.SetFocus
        Exit Sub
    End If
    
    mstrRequestNO = frmDrawCondition.ShowMe(Me, mintUnit, cboStock.Text, Val(cboStock.ItemData(cboStock.ListIndex)), cboEnterStock.Text, Val(cboEnterStock.ItemData(cboEnterStock.ListIndex)))
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

        gstrSQL = "Select a.Id as 材料id, d.数量 as 计划数量,a.编码,a.名称 ,a.规格,c.现价 as 售价,a.计算单位 as 散装单位,a.是否变价 as 时价,b.包装单位,b.换算系数,b.指导差价率,b.最大效期" & vbNewLine & _
                    ",e.上次产地 as 产地,e.上次批号 as 批号,nvl(e.批次,0) as 批次,e.效期,e.灭菌效期,e.可用数量,nvl(e.实际数量,0) as 实际数量,e.实际金额,e.实际差价,e.零售价,e.平均成本价,e.批准文号,b.库房分批,b.在用分批, nvl(b.跟踪病人,0) as 跟踪病人" & vbNewLine & _
                    "From 收费项目目录 A, 材料特性 B, 收费价目 C," & vbNewLine & _
                    "     (Select  b.材料id, Sum(b.计划数量) As 数量" & vbNewLine & _
                    "       From 材料采购计划 A, 材料计划内容 B" & vbNewLine & _
                    "       Where a.Id = b.计划id and a.单据=1 And a.No In (Select * From Table(Cast(f_Str2list([1]) As Zltools.t_Strlist)))" & vbNewLine & _
                    "       Group By b.材料id) D,药品库存 e" & vbNewLine & _
                    "Where a.Id = b.材料id And b.材料id = c.收费细目id And a.Id = d.材料id and b.材料id=e.药品id(+)  and e.库房id=[2] and e.实际数量>0 and e.性质=1 And Sysdate Between c.执行日期 And c.终止日期"

        If gSystem_Para.P156_出库算法 = 0 Then '批次还是效期优先先出库
            gstrSQL = gstrSQL & " Order by a.id,Nvl(e.批次, 0)"
        Else
            gstrSQL = gstrSQL & " Order by a.id,e.效期,Nvl(e.批次, 0)"
        End If

        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "cmdRequestTransfer_Click", mstrRequestNO, cboStock.ItemData(cboStock.ListIndex))

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
                If Format(str灭菌效期, "yyyy-mm-dd") < Format(zlDatabase.Currentdate, "yyyy-mm-dd") And Trim(str灭菌效期) <> "" Then
                   If MsgBox("[" & rsTemp!编码 & "-" & rsTemp!名称 & "]" & "卫材已经过了灭菌效期,是否还要领用！", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) <> vbYes Then
                        blnDo = True
                   End If
                End If

'                str效期 = IIf(IsNull(rsTemp!效期), "", Format(rsTemp!效期, "yyyy-MM-dd"))
'                If IsDate(str效期) Then
'                    If Format(str效期, "yyyy-MM-dd") < Format(zldatabase.Currentdate, "yyyy-MM-dd") Then
'                        MsgBox "[" & rsTemp!编码 & "-" & rsTemp!名称 & "]" & "卫生材料已经失效了！", vbInformation, gstrSysName
'                    End If
'                End If

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

'                '只有不重复的才添加到表格中去
'                If blnDo = False Then
'                    SetRequestColValue lngRow, rsTemp!材料ID, "[" & rsTemp!编码 & "]" & rsTemp!名称, _
'                                IIf(IsNull(rsTemp!规格), "", rsTemp!规格), IIf(IsNull(rsTemp!产地), "", rsTemp!产地), _
'                                IIf(mintUnit = 0, rsTemp!散装单位, rsTemp!包装单位), _
'                                dblPrice, rsTemp!平均成本价, IIf(IsNull(rsTemp!批号), "", rsTemp!批号), _
'                                IIf(IsNull(rsTemp!效期), "", Format(rsTemp!效期, "yyyy-MM-dd")), _
'                                IIf(IsNull(rsTemp!灭菌效期), "", Format(rsTemp!灭菌效期, "yyyy-MM-dd")), _
'                                rsTemp!计划数量, _
'                                IIf(IsNull(rsTemp!可用数量), "0", rsTemp!可用数量), _
'                                dbl数量, _
'                                IIf(IsNull(rsTemp!指导差价率), "0", rsTemp!指导差价率), _
'                                IIf(mintUnit = 0, 1, rsTemp!换算系数), IIf(IsNull(rsTemp!批次), 0, rsTemp!批次), rsTemp!时价, rsTemp!在用分批, IIf(IsNull(rsTemp!批准文号), "", rsTemp!批准文号), rsTemp!跟踪病人, rsTemp!库房分批
'                End If

                '只有不重复的才添加到表格中去
                If blnDo = False Then
                    SetRequestColValue lngRow, rsTemp!材料ID, "[" & rsTemp!编码 & "]" & rsTemp!名称, _
                    IIf(IsNull(rsTemp!规格), "", rsTemp!规格), IIf(IsNull(rsTemp!产地), "", rsTemp!产地), _
                    IIf(mintUnit = 0, rsTemp!散装单位, rsTemp!包装单位), _
                    rsTemp!售价, IIf(IsNull(rsTemp!批号), "", rsTemp!批号), _
                    IIf(IsNull(rsTemp!效期), "", Format(rsTemp!效期, "yyyy-MM-dd")), _
                    IIf(IsNull(rsTemp!灭菌效期), "", Format(rsTemp!灭菌效期, "yyyy-MM-dd")), _
                    IIf(IsNull(rsTemp!最大效期), "0", rsTemp!最大效期), _
                    rsTemp!库房分批, _
                    IIf(IsNull(rsTemp!可用数量), "0", rsTemp!可用数量), _
                    IIf(IsNull(rsTemp!实际金额), "0", rsTemp!实际金额), _
                    IIf(IsNull(rsTemp!实际差价), "0", rsTemp!实际差价), _
                    IIf(IsNull(rsTemp!指导差价率), "0", rsTemp!指导差价率), _
                    IIf(mintUnit = 0, 1, rsTemp!换算系数), IIf(IsNull(rsTemp!批次), 0, rsTemp!批次), rsTemp!时价, rsTemp!在用分批, IIf(IsNull(rsTemp!批准文号), "", rsTemp!批准文号)
                    
                    With mshBill
                        .Row = lngRow
                        .TextMatrix(lngRow, mBillCol.C_行号) = lngRow
                    
                        .TextMatrix(lngRow, mBillCol.C_填写数量) = Format(dbl数量 / IIf(mintUnit = 0, 1, rsTemp!换算系数), mFMT.FM_数量)
                        
                        If .TextMatrix(lngRow, mBillCol.C_售价) <> "" Then
                            .TextMatrix(lngRow, mconintcol售价金额) = Format(.TextMatrix(lngRow, mBillCol.C_售价) * .TextMatrix(lngRow, mBillCol.C_填写数量), mFMT.FM_金额)
                        End If
                        
                        .TextMatrix(lngRow, mBillCol.C_采购价) = Format(Get成本价(Val(.TextMatrix(lngRow, 0)), cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(lngRow, mBillCol.c_批次))) * Val(.TextMatrix(lngRow, mBillCol.C_比例系数)), mFMT.FM_成本价)
                        .TextMatrix(lngRow, mBillCol.C_采购金额) = Format(Val(.TextMatrix(lngRow, mBillCol.C_采购价)) * .TextMatrix(lngRow, mBillCol.C_填写数量), mFMT.FM_金额)
                        .TextMatrix(lngRow, mconintcol差价) = Format(Val(.TextMatrix(lngRow, mconintcol售价金额)) - Val(.TextMatrix(lngRow, mBillCol.C_采购金额)), mFMT.FM_金额)
    
    
                        .TextMatrix(lngRow, mBillCol.C_实际数量) = Format(dbl数量 / IIf(mintUnit = 0, 1, rsTemp!换算系数), mFMT.FM_数量)
                    End With
                    
                End If
                
                blnDo = False
                rsTemp.MoveNext
            End With
        Loop
    End If
End Sub





Private Sub CmdSave_Click()
    Dim blnSuccess As Boolean
    Dim blnTrans As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    Dim strReg As String
    
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
    
   
    If mint编辑状态 = 10 Then        '发送
        '考虑如果不分解，则库存检查过不了，因此此处不检查，强制用户手工点击分解功能
        If Not ValidData Then Exit Sub
        If Not CheckStock Then Exit Sub
        
        '检查是否已备药
        gstrSQL = "Select 1 From 药品收发记录 Where 单据=19 And NO=[1] And 配药人 Is Not NULL"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查是否备料", txtNO.Tag)
        
        If rsTemp.RecordCount = 0 Then
            MsgBox "该单据已被其它操作员取消备料，当前操作中止！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '检查是否已发送
        gstrSQL = "Select 1 From 药品收发记录 Where 单据=19 And NO=[1] And 配药日期 Is Not NULL"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查是否发送", txtNO.Tag)
        If rsTemp.RecordCount <> 0 Then
            MsgBox "该单据已被其它操作员发送，当前操作中止！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        blnTrans = True
        gcnOracle.BeginTrans
        
        '先删除申领单，再依据当前数据产生移库单
        If Not SaveCard(True) Then
            gcnOracle.RollbackTrans: Exit Sub
        End If
        
        '备料
        gstrSQL = "zl_材料移库_Prepare('" & txtNO.Tag & "','" & Txt审核人.Caption & "')"
        zlDatabase.ExecuteProcedure gstrSQL, mstrCaption
                        
        
        '发送（下出库库房的材料可用库存）
        gstrSQL = "zl_材料移库_Prepare('" & txtNO.Tag & "')"
        zlDatabase.ExecuteProcedure gstrSQL, mstrCaption
        
        gcnOracle.CommitTrans
        blnTrans = True
        Unload Me
        Exit Sub
    End If
    
    If mint编辑状态 = 3 Then        '审核
        '判断是否自动执行移库流程，如果是就自动完成备料、发送、接收过程
        If mint移库处理流程 = 0 Then
            blnSuccess = Auto处理移库流程
            Exit Sub
        End If
    
        If Not CheckSend Then Exit Sub
        If Not 材料单据审核(Txt填制人.Caption) Then Exit Sub
        
        If Not 检查单价(19, txtNO.Tag, False) And Not mblnUpdate Then
            MsgBox "有记录未使用最新价格，程序将自动完成更新（售价、成本价、售价金额、成本金额、差价），更新后请检查！", vbInformation, gstrSysName
            Call RefreshBill
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
        
        gcnOracle.BeginTrans
        '如果审核时修改了单据，则重新生成单据保存
        If mblnChange Then
            If Not SaveCard(True) Then
                gcnOracle.RollbackTrans: Exit Sub
            End If
            
            '备料
            gstrSQL = "zl_材料移库_Prepare('" & txtNO.Tag & "','" & UserInfo.用户名 & "')"
            zlDatabase.ExecuteProcedure gstrSQL, mstrCaption
            '发送（下出库库房的材料可用库存）
            gstrSQL = "zl_材料移库_Prepare('" & txtNO.Tag & "')"
            zlDatabase.ExecuteProcedure gstrSQL, mstrCaption
        End If

        If SaveCheck(True) = True Then
            strReg = IIf(Val(zlDatabase.GetPara("审核打印", glngSys, mlngModule, "0")) = 1, 1, 0)
            If Val(strReg) = 1 Then
                '打印
                If InStr(mstrPrivs, "单据打印") <> 0 Then
                    printbill
                End If
            End If
            blnTrans = False
            gcnOracle.CommitTrans
            Unload Me
        Else
            gcnOracle.RollbackTrans: Exit Sub
        End If
        Exit Sub
    End If
    
    If mint编辑状态 = 6 Then '冲销
        If SaveStrike Then
            If mint冲销方式 = 2 Then
                strReg = IIf(Val(zlDatabase.GetPara("审核打印", glngSys, mlngModule, "0")) = 1, 1, 0)
                If Val(strReg) = 1 Then
                    '打印
                    If InStr(mstrPrivs, "单据打印") <> 0 Then
                        printbill
                    End If
                End If
            End If
            Unload Me
        End If
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
    
    If mint编辑状态 = 11 Then
        Unload Me
        Exit Sub
    End If
    
    
    mblnSave = False
    mblnEdit = True
    mshBill.ClearBill
    Call RefreshRowNO(mshBill, mBillCol.C_行号, 1)
    
    txt摘要.Text = ""
    If cboEnterStock.Enabled Then cboEnterStock.SetFocus
    mblnChange = False
    Exit Sub
ErrHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
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
            " Where a.单据 = 19 And a.No = [1] And a.药品id = b.收费细目id And c.Id = b.收费细目id And Round(a.零售价," & g_小数位数.obj_散装小数.零售价小数 & ") <> Round(b.现价, " & g_小数位数.obj_散装小数.零售价小数 & ") And" & _
              "    NVL(c.是否变价, 0) = 0" & _
            " Union All" & _
            " Select '售价' As 类型, a.序号, a.药品id As 材料id, Nvl(a.批次, 0) As 批次, decode(nvl(b.批次,0),0,b.实际金额 / b.实际数量,b.零售价) As 现价" & _
            " From 药品收发记录 A, 药品库存 B, 收费项目目录 C" & _
            " Where a.单据 = 19 And a.No = [1] And c.Id = a.药品id And Round(a.零售价," & g_小数位数.obj_散装小数.零售价小数 & ") <> Round(decode(nvl(b.批次,0),0,b.实际金额 / b.实际数量,b.零售价), " & g_小数位数.obj_散装小数.零售价小数 & ") And Nvl(c.是否变价, 0) = 1 And" & _
                  " b.性质 = 1 And b.库房id = a.库房id And b.药品id = a.药品id And NVL(b.批次, 0) = NVL(a.批次, 0) And NVL(b.实际数量, 0) <> 0 And a.入出系数 = -1" & _
            " Union All" & _
            " Select '成本价' As 类型, a.序号, a.药品id As 材料id, Nvl(a.批次, 0) As 批次, b.平均成本价 As 现价" & _
            " From 药品收发记录 A, 药品库存 B" & _
            " Where a.单据 = 19 And a.No = [1] And a.药品id = b.药品id And Nvl(a.批次, 0) = Nvl(b.批次, 0) and round(a.成本价," & g_小数位数.obj_散装小数.成本价小数 & ")<>round(b.平均成本价," & g_小数位数.obj_散装小数.成本价小数 & ") And a.库房id = b.库房id and a.入出系数=-1 and b.性质=1" & _
            " Order By 类型, 材料id, 序号"

    Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "[取当前价格]", CStr(Me.txtNO.Text))
    
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
                dbl零售价 = Val(Format(rsprice!现价 * Val(mshBill.TextMatrix(lngRow, mBillCol.C_比例系数)), mFMT.FM_零售价))
                dbl零售金额 = Val(Format(dbl零售价 * dbl数量, mFMT.FM_金额))
                dbl差价 = Val(Format(dbl零售金额 - dbl成本金额, mFMT.FM_金额))
            End If

            rsprice.Filter = "类型='成本价' And 材料id=" & lng材料ID & " And 批次=" & Val(mshBill.TextMatrix(lngRow, mBillCol.c_批次))
            If rsprice.RecordCount > 0 Then
                blnAdj = True
                dbl零售金额 = Val(Format(dbl零售价 * dbl数量, mFMT.FM_金额))
                dbl成本价 = Val(Format(rsprice!现价 * Val(mshBill.TextMatrix(lngRow, mBillCol.C_比例系数)), mFMT.FM_金额))
                dbl成本金额 = Val(Format(dbl成本价 * dbl数量, mFMT.FM_金额))
                dbl差价 = Val(Format(dbl零售金额 - dbl成本金额, mFMT.FM_金额))
            End If

            If blnAdj = True Then
                '以当前最新价格最新单据相关数据（售价、成本价、零售金额、成本金额、差价）
                mshBill.TextMatrix(lngRow, mBillCol.C_售价) = Format(dbl零售价, mFMT.FM_零售价)
                mshBill.TextMatrix(lngRow, mconintcol售价金额) = Format(dbl零售金额, mFMT.FM_金额)
                mshBill.TextMatrix(lngRow, mBillCol.C_采购价) = Format(dbl成本价, mFMT.FM_成本价)
                mshBill.TextMatrix(lngRow, mBillCol.C_采购金额) = Format(dbl成本金额, mFMT.FM_金额)
                mshBill.TextMatrix(lngRow, mconintcol差价) = Format(dbl差价, mFMT.FM_金额)
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

Private Sub Form_Activate()
    
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
'    mblnChange = False
    Select Case mintParallelRecord
        Case 1
            '正常
        Case 2
            If mint编辑状态 = 6 Then
                MsgBox "该单据已没有可以冲销的卫材，请检查！", vbOKOnly, gstrSysName
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
        stbThis.Panels("PY").Visible = True
        stbThis.Panels("WB").Visible = True
        gSystem_Para.int简码方式 = Val(zlDatabase.GetPara("简码方式", , , 0))    '默认拼音简码
        Logogram stbThis, gSystem_Para.int简码方式
    Else
        stbThis.Panels("PY").Visible = False
        stbThis.Panels("WB").Visible = False
    End If
End Sub

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

Private Sub Form_Load()
    Dim strStock As String
    Dim rsEnterStock As New Recordset
    Dim rsPara As New ADODB.Recordset
    Dim strReg As String
    
    On Error GoTo ErrHandle
    mblnFirst = True
    mblnUpdate = False
    
    mbln申领核查 = IIf((zlDatabase.GetPara("申领需要核查后才能移库", glngSys, 1722, "0")) = 0, False, True)
    strReg = Val(zlDatabase.GetPara("卫材单位", glngSys, mlngModule, "0"))
    mintUnit = Val(strReg)
    
    '刘兴宏:增加小数格式化串
    With mFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价)
        .FM_金额 = GetFmtString(mintUnit, g_金额)
        .FM_零售价 = GetFmtString(mintUnit, g_售价)
        .FM_数量 = GetFmtString(mintUnit, g_数量)
    End With
     
    mintBatchNoLen = GetBatchNoLen()

    txtNO = mstr单据号
    txtNO.Tag = txtNO.Text
    
    strStock = "And b.编码 In('V','K','W','12') "
    gstrSQL = "" & _
        "   SELECT DISTINCT a.id, a.名称 " & _
        "   FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " & _
        "   Where c.工作性质 = b.名称 And (a.站点=[1] or a.站点 is null) " & _
        "        " & strStock & _
        "       AND a.id = c.部门id " & _
        "       AND a.撤档时间 = to_date('3000-01-01','yyyy-MM-dd')"
    Set rsEnterStock = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, gstrNodeNo)
    
    With cboEnterStock
        .Clear
        Do While Not rsEnterStock.EOF
            .AddItem rsEnterStock.Fields(1)
            .ItemData(.NewIndex) = rsEnterStock.Fields(0)
            rsEnterStock.MoveNext
        Loop
        If .ListCount > 0 Then
            .ListIndex = 0
        End If
        .Tag = 0
    End With
    
    '取系统参数“明确申领卫材批次”
    mbln明确批次 = IS批次申领
    
    '取系统参数“明确移库卫材批次”
    mbln移库明确批次 = IS批次移库
    
    mbln分批卫材批号产地控制 = Val(zlDatabase.GetPara(305, glngSys, 0)) = 1
    
    '出库库房缺省为主界面当前选择的库房，对于新增有效
    On Error Resume Next
    mlng出库库房 = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    
    Call initCard
    mstrTime_Start = GetBillInfo(19, mstr单据号)
    '恢复个性化参数设置
    RestoreWinState Me, App.ProductName, mstrCaption
    '恢复个性化参数设置后，还需要对权限控制的列进一步设置
    With mshBill
        .ColWidth(mBillCol.C_采购价) = IIf(mblnCostView = True, 800, 0)
        .ColWidth(mBillCol.C_采购金额) = IIf(mblnCostView = True, 900, 0)
        .ColWidth(mconintcol差价) = IIf(mblnCostView = True, 800, 0)
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub initCard()
    Dim i As Integer
    Dim rsInitCard As New Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim str批次 As String
    Dim strArray As String
    Dim intRow As Integer
    Dim numUseAbleCount As Double
    Dim vardrug As Variant
    Dim strOrder As String, strCompare As String
    Dim lng出库库房 As Long, lng入库库房 As Long
    
    '库房
    On Error GoTo ErrHandle
    mbln申领单 = False
    mstr核查人 = ""
    mstr核查日期 = ""
    
    strOrder = zlDatabase.GetPara("单据排序", glngSys, mlngModule, "00")
    strOrder = IIf(strOrder = "", "00", strOrder)
    strCompare = Mid(strOrder, 1, 1)
    
    '取指定单据的出库库房与入库库房
    gstrSQL = " Select 库房ID,对方部门ID From 药品收发记录" & _
              " Where NO=[1] And 单据=19 And 入出系数=-1 And Rownum<2"
    
    Set rsInitCard = zlDatabase.OpenSQLRecord(gstrSQL, "取指定单据的出库库房与入库库房", mstr单据号)
              
    If rsInitCard.RecordCount <> 0 Then
        lng出库库房 = rsInitCard!库房ID
        lng入库库房 = rsInitCard!对方部门id
    End If
    If lng出库库房 = 0 Then lng出库库房 = mlng出库库房
    
    mint移库处理流程 = IIf(Val(zlDatabase.GetPara("移库流程", glngSys, mlngModule, "0", , , , lng出库库房)) = 1, 1, 0)
        
    If mint编辑状态 <> 4 Then
        With mfrmMain.cboStock
            cboStock.Clear
            For i = 0 To .ListCount - 1
                cboStock.AddItem .List(i)
                cboStock.ItemData(cboStock.NewIndex) = .ItemData(i)
                If .ItemData(i) = lng出库库房 Then cboStock.ListIndex = cboStock.ListCount - 1
            Next
            mintcboIndex = cboStock.ListIndex
            '如果没有指定的部门，将其加入
            If mintcboIndex = -1 Then
                gstrSQL = "Select ID,名称 From 部门表 Where ID=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "如果没有指定的出库部门，将其加入", lng出库库房)
                
                cboStock.AddItem rsTemp!名称
                cboStock.ItemData(cboStock.NewIndex) = rsTemp!Id
                cboStock.ListIndex = cboStock.ListCount - 1
            End If
            mintcboIndex = cboStock.ListIndex
            cboStock.Enabled = .Enabled
        End With
        
    End If
    
    Select Case mint编辑状态
        Case 1
            Txt填制人 = UserInfo.用户名
            Txt填制日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            initGrid
            
            If cboEnterStock.ListCount <> 0 Then
                If cboStock.ItemData(cboStock.ListIndex) = cboEnterStock.ItemData(cboEnterStock.ListIndex) Then
                    If cboEnterStock.ListCount > 1 Then
                        cboEnterStock.ListIndex = cboEnterStock.ListIndex + 1
                    End If
                End If
            End If
        
        Case 2, 3, 4, 6, 10, 11
            initGrid
            '检查该单据是否是申领单据
            gstrSQL = "" & _
                "   Select Nvl(发药方式,0) 申领,核查人,核查日期 From 药品收发记录 " & _
                "   Where 单据=19 And NO=[1] And 序号=1"
                
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr单据号)
            
            If Not rsTemp.EOF Then
                mbln申领单 = (rsTemp!申领 = 1)
                mstr核查人 = IIf(IsNull(rsTemp!核查人), "", rsTemp!核查人)
                mstr核查日期 = IIf(IsNull(rsTemp!核查日期), "", rsTemp!核查日期)
                If mbln申领单 Then LblTitle.Caption = GetUnitName & "卫材申领单"
            End If
            
            If mint编辑状态 = 4 Then
                gstrSQL = "" & _
                    "   Select distinct b.id,b.名称 " & _
                    "   From 药品收发记录 a,部门表 b " & _
                    "   Where a.库房id=b.id and A.单据 = 19 and a.no=[1] and a.入出系数=-1"
                
                Set rsInitCard = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr单据号)
                    
                If rsInitCard.EOF Then
                    mintParallelRecord = 2
                    Exit Sub
                End If
                
                With cboStock
                    .AddItem rsInitCard!名称
                    .ItemData(.NewIndex) = rsInitCard!Id
                    .ListIndex = 0
                End With
                rsInitCard.Close
            End If
            
            Select Case mintUnit
            
                Case 0
                    strUnitQuantity = "c.计算单位 AS 单位, A.填写数量,a.实际数量,a.成本价,a.零售价,'1' as 比例系数,"
                Case Else
                    strUnitQuantity = "B.包装单位 AS 单位,(A.填写数量 / B.换算系数) AS 填写数量,(A.实际数量 / B.换算系数) AS 实际数量,a.成本价*B.换算系数 as 成本价,a.零售价*B.换算系数 as 零售价,B.换算系数 as 比例系数,"
            End Select
            
            
            Select Case mint编辑状态
            Case 6
                If mint冲销方式 <> 2 Then
                    gstrSQL = "" & _
                        "   select w.*,z.可用数量/w.比例系数 as  可用数量,z.实际金额,z.实际差价 " & _
                        "   From (  SELECT distinct a.材料id,A.序号,('[' || c.编码 || ']' || c.名称) AS 卫材信息," & _
                        "                       zlSpellCode(c.名称) 名称,c.规格,c.产地 as 原产地,A.产地,A.批准文号, A.批号,a.批次,b.指导差价率,b.库房分批," & _
                        "                       b.最大效期,A.效期,A.灭菌日期,A.灭菌效期 as 灭菌失效期,B.一次性材料,b.灭菌效期,A.填写数量 原始数量," & strUnitQuantity & _
                        "                       A.成本金额,0 零售金额, 0 差价,a.摘要,a.库房id,a.对方部门id,c.是否变价,b.在用分批  " & _
                        "           FROM (  Select min(id) as id, sum(实际数量) as 填写数量,0 实际数量,sum(成本金额) as 成本金额,药品id 材料ID,序号,产地,批准文号, 批号,效期,灭菌日期,灭菌效期 ," & _
                        "                           Nvl(批次,0) 批次,扣率,成本价,零售价,摘要,库房ID,对方部门ID,入出类别ID" & _
                        "                   From 药品收发记录 x " & _
                        "                   WHERE NO=[1] AND 单据=19 And 入出系数=-1 " & _
                        "                   group by 药品ID,序号,产地,批准文号,批号,效期,灭菌日期,灭菌效期,Nvl(批次,0),扣率,成本价,零售价,摘要,库房ID,对方部门ID,入出类别ID" & _
                        "                   having sum(实际数量)<>0 ) A, 材料特性 B,收费项目目录 C " & _
                        "           Where A.材料id = B.材料id  and A.材料id=c.id " & _
                        "       ) w,(   Select  药品id 材料id,Nvl(批次,0) 批次,可用数量,实际金额,实际差价 " & _
                        "               From 药品库存 " & _
                        "               where 库房id=[2]  and 性质=1)  z " & _
                        "   Where w.材料id=z.材料id(+) and nvl(w.批次,0)=nvl(z.批次(+),0) " & _
                        "   ORDER BY " & IIf(strCompare = "0", "序号", IIf(strCompare = "1", "卫材信息", "名称")) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
                Else
                    '用于审核冲销时，显示未审核的申请冲销单据
                    gstrSQL = "SELECT W.*,Z.可用数量/W.比例系数 AS  可用数量,Z.实际金额,Z.实际差价 " & _
                        " FROM " & _
                        "     (SELECT DISTINCT A.药品ID as 材料id,A.序号,('[' || c.编码 || ']' || c.名称) AS 卫材信息,zlSpellCode(c.名称) as 名称," & _
                        "     C.规格,C.产地 AS 原产地,A.产地,A.批准文号, A.批号,A.批次,B.指导差价率,B.库房分批," & _
                        "     B.最大效期,A.效期,A.灭菌日期,A.灭菌效期 as 灭菌失效期,B.一次性材料,b.灭菌效期,A.填写数量 原始数量," & strUnitQuantity & "A.成本金额,A.零售金额, A.差价,A.配药人, " & _
                        "     A.摘要,填制人,填制日期,审核人,审核日期,A.库房ID,A.对方部门ID,C.是否变价,B.在用分批,NVL(A.供药单位ID,0) 上次供应商ID" & _
                        "     FROM 药品收发记录 A, 材料特性 B,收费项目目录 C,收费项目别名 E " & _
                        "     WHERE A.药品ID = B.材料ID AND B.材料ID=C.ID AND B.材料ID=E.收费细目ID(+) AND E.性质(+)=3 " & _
                        "     AND A.记录状态 =[3] " & _
                        "     AND A.单据 = 19 AND A.入出系数=-1 AND A.NO =[1] ) W," & _
                        "     (SELECT  药品ID,NVL(批次,0) 批次,可用数量,实际金额,实际差价 " & _
                        "     FROM 药品库存 WHERE 库房ID=[2] AND 性质=1) Z " & _
                        " WHERE W.材料id=Z.药品ID(+) AND NVL(W.批次,0)=Nvl(Z.批次(+),0) " & _
                        "   ORDER BY " & IIf(strCompare = "0", "序号", IIf(strCompare = "1", "卫材信息", "名称")) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
                End If
            Case 11
                Dim bln具备跟踪材料 As Boolean
                bln具备跟踪材料 = 判断只具备发料部门(cboEnterStock.ItemData(cboEnterStock.ListIndex))

                gstrSQL = "" & _
                    "   Select w.材料ID,w.序号,w.卫材信息,W.名称,w.规格,w.原产地,w.产地 ,w.批准文号,w.批次,w.批号,w.指导差价率, " & _
                    "           w.库房分批,w.最大效期,w.效期,w.灭菌日期,w.灭菌失效期, " & _
                    "           w.一次性材料,w.灭菌效期,w.单位,w.原始数量 原始数量,w.填写数量,w.实际数量,w.零售价,w.零售金额,w.比例系数, " & _
                    "           (w.零售金额 - Decode(Sign(nvl(z.实际金额,0)),1,w.零售金额 * (nvl(z.实际差价,0) / z.实际金额),w.零售金额 * w.指导差价率 / 100)) / decode(w.实际数量,0,1,w.实际数量)  成本价, " & _
                    "           (w.零售金额 - Decode(Sign(z.实际金额),1,w.零售金额 * (z.实际差价 / z.实际金额),w.零售金额 * w.指导差价率 / 100)) 成本金额, " & _
                    "           Decode(Sign(z.实际金额),1,w.零售金额 * (z.实际差价 / z.实际金额),w.零售金额 * w.指导差价率 / 100) 差价, " & _
                    "            w.摘要,w.填制人,w.填制日期,w.配药人, w.审核人, w.审核日期,w.库房id,w.对方部门id,w.是否变价,w.在用分批,z.可用数量/w.比例系数 as  可用数量,z.实际金额,z.实际差价   " & _
                    "    From (  SELECT distinct a.药品id 材料id,A.序号,('[' || c.编码 || ']' || c.名称) AS 卫材信息,  " & _
                    "                    zlSpellCode(C.名称) 名称,c.规格,C.产地 as 原产地,A.产地,A.批准文号, A.批号,a.批次,b.指导差价率,b.库房分批,  " & _
                    "                    b.最大效期,A.效期,A.灭菌日期,A.灭菌效期 as 灭菌失效期,B.一次性材料,b.灭菌效期,A.填写数量 原始数量, " & strUnitQuantity & _
                    "                    A.成本金额,A.零售金额, A.差价,   " & _
                    "                    a.摘要,a.填制人,A.填制日期,A.配药人, A.审核人, A.审核日期,a.库房id,a.对方部门id,c.是否变价,b.在用分批   " & _
                    "            FROM 药品收发记录 A, 材料特性 B,收费项目目录 C  " & _
                    "            Where A.药品id = B.材料id and a.药品id=c.id   " & IIf(bln具备跟踪材料, " and nvl(B.跟踪在用,0)=1 ", "") & _
                    "                    AND A.记录状态 =[3]  " & _
                    "                    AND A.单据 = 15 AND A.No = [1]  " & _
                    "           ) w, (  Select 药品id 材料id,Nvl(批次,0) 批次,可用数量,实际金额,实际差价   " & _
                    "                    From 药品库存 where 库房id=[2]  and 性质=1)  z, " & _
                    "  (Select Distinct 收费细目id From 收费执行科室 f Where 执行科室id = [4]) y " & _
                    "    Where w.材料id=z.材料id(+)  AND W.材料id=Y.收费细目id  and nvl(w.批次,0)=nvl(z.批次(+),0)   " & _
                    "   ORDER BY " & IIf(strCompare = "0", "序号", IIf(strCompare = "1", "卫材信息", "名称")) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
                    
            Case Else
                gstrSQL = "" & _
                    "   Select w.*,z.可用数量/w.比例系数 as  可用数量,z.实际金额,z.实际差价 " & _
                    "   From (  SELECT distinct a.药品id 材料id,A.序号,('[' || c.编码 || ']' || c.名称) AS 卫材信息," & _
                    "                   zlSpellCode(C.名称) 名称,c.规格,C.产地 as 原产地,A.产地,A.批准文号, A.批号,a.批次,b.指导差价率,b.库房分批," & _
                    "                   b.最大效期,A.效期,A.灭菌日期,A.灭菌效期 as 灭菌失效期,B.一次性材料,b.灭菌效期,A.填写数量 原始数量," & strUnitQuantity & _
                    "                   A.成本金额,A.零售金额, A.差价, " & _
                    "                   a.摘要,填制人,填制日期,A.配药人,审核人,审核日期,a.库房id,a.对方部门id,c.是否变价,b.在用分批 " & _
                    "           FROM 药品收发记录 A, 材料特性 B,收费项目目录 C " & _
                    "           Where A.药品id = B.材料id and a.药品id=c.id " & _
                    "                   AND A.记录状态 =[3]" & _
                    "                   AND A.单据 = 19 and a.入出系数=-1 AND A.No = [1]" & _
                    "          ) w, (  Select 药品id 材料id,Nvl(批次,0) 批次,可用数量,实际金额,实际差价 " & _
                    "                   From 药品库存 where 库房id=[2]  and 性质=1)  z " & _
                    "   Where w.材料id=z.材料id(+) and nvl(w.批次,0)=nvl(z.批次(+),0) " & _
                    "   ORDER BY " & IIf(strCompare = "0", "序号", IIf(strCompare = "1", "卫材信息", "名称")) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
            
            End Select
            
            Set rsInitCard = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, IIf(mint编辑状态 = 11, mstr入库单号, mstr单据号), cboStock.ItemData(cboStock.ListIndex), mint记录状态, cboEnterStock.ItemData(cboEnterStock.ListIndex))
            
            If rsInitCard.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            
            Select Case mint编辑状态
            Case 2, 6, 10, 11
                Txt填制人 = UserInfo.用户名
                Txt填制日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                If mint编辑状态 = 6 Or mint编辑状态 = 10 Then
                    Txt审核人 = UserInfo.用户名
                    Txt审核日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                End If
                If mint编辑状态 = 10 Then
                    Txt审核人 = zlStr.NVL(rsInitCard!配药人)
                    Txt填制人 = rsInitCard!填制人
                    Txt填制日期 = Format(rsInitCard!填制日期, "yyyy-mm-dd hh:mm:ss")
                    Lbl审核人.Caption = "备料人"
                    Lbl审核日期.Caption = "发送日期"
                End If
            Case Else
                Txt填制人 = rsInitCard!填制人
                Txt填制日期 = Format(rsInitCard!填制日期, "yyyy-mm-dd hh:mm:ss")
                Txt审核人 = IIf(IsNull(rsInitCard!审核人), "", rsInitCard!审核人)
                Txt审核日期 = IIf(IsNull(rsInitCard!审核日期), "", Format(rsInitCard!审核日期, "yyyy-mm-dd hh:mm:ss"))
            End Select
            
            txt摘要.Text = IIf(IsNull(rsInitCard!摘要), "", rsInitCard!摘要)
            
            If (mint编辑状态 = 2 Or mint编辑状态 = 3) And Txt审核人 <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            Dim intCount As Integer
            With cboEnterStock
                For intCount = 0 To .ListCount - 1
                    If .ItemData(intCount) = rsInitCard!对方部门id Then
                        .ListIndex = intCount
                        .Tag = intCount
                        Exit For
                    End If
                Next
            End With
            
            If mint编辑状态 = 2 Or mint编辑状态 = 3 Then
                Set mcolUsedCount = New Collection
            End If
            
            With mshBill
                Do While Not rsInitCard.EOF
                    intRow = rsInitCard.AbsolutePosition
                    .Rows = intRow + 1
                    .TextMatrix(intRow, 0) = rsInitCard.Fields(0)
                    .TextMatrix(intRow, mBillCol.C_材料) = rsInitCard!卫材信息
                    .TextMatrix(intRow, mBillCol.c_序号) = rsInitCard!序号
                    .TextMatrix(intRow, mBillCol.c_规格) = IIf(IsNull(rsInitCard!规格), "", rsInitCard!规格)
                    .TextMatrix(intRow, mBillCol.C_产地) = IIf(IsNull(rsInitCard!产地), "", rsInitCard!产地)
                    .TextMatrix(intRow, mBillCol.C_批准文号) = IIf(IsNull(rsInitCard!批准文号), "", rsInitCard!批准文号)
                    .TextMatrix(intRow, mBillCol.c_单位) = rsInitCard!单位
                    .TextMatrix(intRow, mBillCol.c_批号) = IIf(IsNull(rsInitCard!批号), "", rsInitCard!批号)
                    .TextMatrix(intRow, mBillCol.C_效期) = IIf(IsNull(rsInitCard!效期), "", Format(rsInitCard!效期, "yyyy-mm-dd"))
                    
                    .TextMatrix(intRow, mBillCol.C_一次性材料) = zlStr.NVL(rsInitCard!一次性材料)
                    .TextMatrix(intRow, mBillCol.C_灭菌效期) = zlStr.NVL(rsInitCard!灭菌效期)
                    .TextMatrix(intRow, mBillCol.C_灭菌日期) = IIf(IsNull(rsInitCard!灭菌日期), "", Format(rsInitCard!灭菌日期, "yyyy-mm-dd"))
                    .TextMatrix(intRow, mBillCol.C_灭菌失效期) = IIf(IsNull(rsInitCard!灭菌失效期), "", Format(rsInitCard!灭菌失效期, "yyyy-mm-dd"))
        
                    .TextMatrix(intRow, mBillCol.C_填写数量) = Format(IIf(mint编辑状态 = 6 And mint冲销方式 = 2, -1, 1) * rsInitCard!填写数量, mFMT.FM_数量)
                    .TextMatrix(intRow, mBillCol.C_实际数量) = Format(IIf(mint编辑状态 = 6 And mint冲销方式 = 2, -1, 1) * rsInitCard!实际数量, mFMT.FM_数量)
                    
                    
                    If mint编辑状态 = 2 Or mint编辑状态 = 6 Or mint编辑状态 = 3 Or mint编辑状态 = 10 Or mint编辑状态 = 11 Then
                        .TextMatrix(intRow, mBillCol.c_原始数量) = Format(IIf(mint编辑状态 = 6 And mint冲销方式 = 2, -1, 1) * rsInitCard!原始数量, mFMT.FM_数量)
                    End If

                    .TextMatrix(intRow, mBillCol.C_采购价) = Format(rsInitCard!成本价, mFMT.FM_成本价)
                    .TextMatrix(intRow, mBillCol.C_采购金额) = Format(IIf(mint编辑状态 = 6 And mint冲销方式 <> 2, 0, IIf(mint编辑状态 = 6 And mint冲销方式 = 2, -1, 1)) * rsInitCard!成本金额, mFMT.FM_金额)
                    .TextMatrix(intRow, mBillCol.C_售价) = Format(rsInitCard!零售价, mFMT.FM_零售价)
                    .TextMatrix(intRow, mconintcol售价金额) = Format(IIf(mint编辑状态 = 6 And mint冲销方式 = 2, -1, 1) * rsInitCard!零售金额, mFMT.FM_金额)
                    .TextMatrix(intRow, mconintcol差价) = Format(IIf(mint编辑状态 = 6 And mint冲销方式 = 2, -1, 1) * rsInitCard!差价, mFMT.FM_金额)
                    .TextMatrix(intRow, mBillCol.C_最大效期) = IIf(IsNull(rsInitCard!最大效期), "0", rsInitCard!最大效期) & "||" & rsInitCard!是否变价 & "||" & rsInitCard!在用分批
                    .TextMatrix(intRow, mBillCol.c_批次) = IIf(IsNull(rsInitCard!批次), "0", rsInitCard!批次)
                    .TextMatrix(intRow, mBillCol.C_比例系数) = rsInitCard!比例系数
                    
                    .TextMatrix(intRow, mBillCol.C_指导差价率) = rsInitCard!指导差价率
                    .TextMatrix(intRow, mBillCol.C_库房分批) = IIf(IsNull(rsInitCard!库房分批), "0", rsInitCard!库房分批)
                    .TextMatrix(intRow, mBillCol.C_可用数量) = Format(IIf(IsNull(rsInitCard!可用数量), "0", rsInitCard!可用数量), mFMT.FM_数量)
                    .TextMatrix(intRow, mBillCol.C_实际差价) = IIf(IsNull(rsInitCard!实际差价), "0", rsInitCard!实际差价)
                    .TextMatrix(intRow, mBillCol.C_实际金额) = IIf(IsNull(rsInitCard!实际金额), "0", rsInitCard!实际金额)
                    
                    If mint编辑状态 = 2 Or mint编辑状态 = 3 Then
                        numUseAbleCount = 0
                        For Each vardrug In mcolUsedCount
                            If vardrug(0) = CStr(rsInitCard!材料ID & IIf(IsNull(rsInitCard!批次), "0", rsInitCard!批次)) Then
                                numUseAbleCount = vardrug(1)
                                mcolUsedCount.Remove vardrug(0)
                                Exit For
                            End If
                        Next
                        str批次 = rsInitCard!材料ID & IIf(IsNull(rsInitCard!批次), "0", rsInitCard!批次)
                        If mint编辑状态 = 2 Then
                            strArray = numUseAbleCount + IIf(IsNull(rsInitCard!填写数量), "0", rsInitCard!填写数量)
                        Else
                            strArray = numUseAbleCount + IIf(IsNull(rsInitCard!实际数量), "0", rsInitCard!实际数量)
                        End If
                        mcolUsedCount.Add Array(str批次, strArray), str批次
                    End If
                    rsInitCard.MoveNext
                Loop
                .Rows = intRow + 2
            End With
            rsInitCard.Close
    End Select
    Call RefreshRowNO(mshBill, mBillCol.C_行号, 1)
    
    SetEdit         '设置编辑属性
    '查阅、修改或审核时，根据库存与申领数量显示单据
    If (mint编辑状态 = 2 Or mint编辑状态 = 3 Or mint编辑状态 = 4 Or mint编辑状态 = 10) Then
        If mbln申领单 Then Call ShowColor
        Select Case mint编辑状态
        Case 2, 10
            cmdExpend.Visible = True
        End Select
    End If
    If mint移库处理流程 = 0 And mint编辑状态 = 3 Then
        cmdExpend.Visible = True
    End If
    
    Call 显示合计金额
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetEdit()
    Dim intCol As Integer
    
    With mshBill
        If mblnEdit = False Then
            For intCol = 0 To .Cols - 1
                .ColData(intCol) = 0
            Next
            cboStock.Enabled = False
            cboEnterStock.Enabled = False
            txt摘要.Enabled = False
            If mint编辑状态 = 6 Then
                If mint冲销方式 <> 2 Then
                    .ColData(mBillCol.C_实际数量) = 4
                End If
            End If
        Else
            .ColData(0) = 5
            .ColData(mBillCol.C_材料) = 1
            .ColData(mBillCol.c_序号) = 5
            .ColData(mBillCol.c_规格) = 5
            .ColData(mBillCol.C_产地) = 5
            .ColData(mBillCol.c_单位) = 5
            .ColData(mBillCol.c_批号) = 5
            .ColData(mBillCol.C_效期) = 5
            .ColData(mBillCol.C_一次性材料) = 5
            .ColData(mBillCol.C_灭菌效期) = 5
            .ColData(mBillCol.C_灭菌失效期) = 5
            .ColData(mBillCol.C_灭菌日期) = 5
            .ColData(mBillCol.c_原始数量) = 5
            If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
                .ColData(mBillCol.C_填写数量) = 4
                .ColData(mBillCol.C_灭菌日期) = 5
                .ColData(mBillCol.C_实际数量) = 5
            ElseIf mint编辑状态 = 3 Then
                .ColData(mBillCol.C_填写数量) = 5
                .ColData(mBillCol.C_实际数量) = 4
            ElseIf mint编辑状态 = 11 Then
                .ColData(mBillCol.C_填写数量) = 5
                .ColData(mBillCol.C_实际数量) = 5
            End If
            
            .ColData(mBillCol.C_采购价) = 5
            .ColData(mBillCol.C_采购金额) = 5
            .ColData(mBillCol.C_售价) = 5
            .ColData(mconintcol售价金额) = 5
            .ColData(mconintcol差价) = 5
            
            .ColData(mBillCol.C_库房分批) = 5
            .ColData(mBillCol.C_可用数量) = 5
            .ColData(mBillCol.C_最大效期) = 5
            
            .ColData(mBillCol.C_指导差价率) = 5
            .ColData(mBillCol.C_实际金额) = 5
            .ColData(mBillCol.C_实际差价) = 5
            .ColData(mBillCol.C_比例系数) = 5
            .ColData(mBillCol.c_批次) = 5
        
            .ColAlignment(mBillCol.C_材料) = flexAlignLeftCenter
            .ColAlignment(mBillCol.c_规格) = flexAlignLeftCenter
            .ColAlignment(mBillCol.C_产地) = flexAlignLeftCenter
            .ColAlignment(mBillCol.c_单位) = flexAlignCenterCenter
            .ColAlignment(mBillCol.c_批号) = flexAlignLeftCenter
            .ColAlignment(mBillCol.C_效期) = flexAlignLeftCenter
            .ColAlignment(mBillCol.C_填写数量) = flexAlignRightCenter
            .ColAlignment(mBillCol.C_实际数量) = flexAlignRightCenter
            
            .ColAlignment(mBillCol.C_采购价) = flexAlignRightCenter
            .ColAlignment(mBillCol.C_采购金额) = flexAlignRightCenter
            .ColAlignment(mBillCol.C_售价) = flexAlignRightCenter
            .ColAlignment(mconintcol售价金额) = flexAlignRightCenter
            .ColAlignment(mconintcol差价) = flexAlignRightCenter
            
            If mint编辑状态 = 11 Then
                '入库转入也不能进行编辑
                cboStock.Enabled = False
            Else
                cboStock.Enabled = True
            End If
            If mint编辑状态 = 11 Then
                cboEnterStock.Enabled = False
            Else
                cboEnterStock.Enabled = True
            End If
            txt摘要.Enabled = True
        End If
    End With
End Sub


Private Sub initGrid()
    With mshBill
        .Active = (mint编辑状态 <> 11)
        .Cols = mBillCols
        
        .MsfObj.FixedCols = 1
        
        .TextMatrix(0, mBillCol.C_行号) = ""
        .TextMatrix(0, mBillCol.C_材料) = "名称与编码"
        .TextMatrix(0, mBillCol.c_序号) = "序号"
        .TextMatrix(0, mBillCol.c_规格) = "规格"
        .TextMatrix(0, mBillCol.C_产地) = "产地"
        .TextMatrix(0, mBillCol.C_批准文号) = "批准文号"
        .TextMatrix(0, mBillCol.c_单位) = "单位"
        .TextMatrix(0, mBillCol.c_批号) = "批号"
        .TextMatrix(0, mBillCol.C_效期) = "失效期"
        
        .TextMatrix(0, mBillCol.C_一次性材料) = "一次性材料"
        .TextMatrix(0, mBillCol.C_灭菌效期) = "灭菌效期"
        .TextMatrix(0, mBillCol.C_灭菌失效期) = "灭菌失效期"
        .TextMatrix(0, mBillCol.C_灭菌日期) = "灭菌日期"
        
        .TextMatrix(0, mBillCol.C_填写数量) = IIf(mint编辑状态 = 6, "数量", "填写数量")
        .TextMatrix(0, mBillCol.C_实际数量) = IIf(mint编辑状态 = 6, "冲销数量", "实际数量")
        
        .TextMatrix(0, mBillCol.C_采购价) = "成本价"
        .TextMatrix(0, mBillCol.C_采购金额) = "成本金额"
        .TextMatrix(0, mBillCol.C_售价) = "售价"
        .TextMatrix(0, mconintcol售价金额) = "售价金额"
        .TextMatrix(0, mconintcol差价) = "差价"
        
        .TextMatrix(0, mBillCol.C_可用数量) = "可用数量"
        .TextMatrix(0, mBillCol.C_库房分批) = "库房分批"
        .TextMatrix(0, mBillCol.C_最大效期) = "最大效期"
        .TextMatrix(0, mBillCol.C_实际差价) = "实际差价"
        .TextMatrix(0, mBillCol.C_实际金额) = "实际金额"
        .TextMatrix(0, mBillCol.C_指导差价率) = "指导差价率"
        .TextMatrix(0, mBillCol.C_比例系数) = "比例系数"
        .TextMatrix(0, mBillCol.c_批次) = "批次"
        
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, mBillCol.C_行号) = "1"
        
        .ColWidth(0) = 0
        .ColWidth(mBillCol.C_行号) = 300
        .ColWidth(mBillCol.C_材料) = 2000
        .ColWidth(mBillCol.c_序号) = 0
        .ColWidth(mBillCol.c_规格) = 900
        .ColWidth(mBillCol.C_产地) = 800
        .ColWidth(mBillCol.C_批准文号) = 800
        .ColWidth(mBillCol.c_单位) = 500
        .ColWidth(mBillCol.c_批号) = 800
        .ColWidth(mBillCol.C_效期) = 1000
     
        .ColWidth(mBillCol.C_一次性材料) = 0
        .ColWidth(mBillCol.C_灭菌效期) = 0
        .ColWidth(mBillCol.C_灭菌失效期) = 1000
        .ColWidth(mBillCol.C_灭菌日期) = 0
          
        .ColWidth(mBillCol.C_填写数量) = 800
        .ColWidth(mBillCol.C_实际数量) = 800
        .ColWidth(mBillCol.C_采购价) = IIf(mblnCostView = False, 0, 800)
        .ColWidth(mBillCol.C_采购金额) = IIf(mblnCostView = False, 0, 900)
        .ColWidth(mBillCol.C_售价) = 800
        .ColWidth(mconintcol售价金额) = 900
        .ColWidth(mconintcol差价) = IIf(mblnCostView = False, 0, 800)
        
        .ColWidth(mBillCol.C_库房分批) = 0
        .ColWidth(mBillCol.C_可用数量) = 0
        .ColWidth(mBillCol.C_最大效期) = 0
        .ColWidth(mBillCol.C_实际差价) = 0
        .ColWidth(mBillCol.C_实际金额) = 0
        .ColWidth(mBillCol.C_指导差价率) = 0
        .ColWidth(mBillCol.C_比例系数) = 0
        .ColWidth(mBillCol.c_批次) = 0
        .ColWidth(mBillCol.c_原始数量) = 0
        
        
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
        .ColData(mBillCol.c_序号) = 5
        .ColData(mBillCol.C_产地) = 5
        .ColData(mBillCol.C_批准文号) = 5
        .ColData(mBillCol.c_单位) = 5
        .ColData(mBillCol.c_批号) = 5
        .ColData(mBillCol.C_效期) = 5
   
        .ColData(mBillCol.C_一次性材料) = 5
        .ColData(mBillCol.C_灭菌效期) = 5
        .ColData(mBillCol.C_灭菌失效期) = 5
        .ColData(mBillCol.C_灭菌日期) = 5
        .ColData(mBillCol.c_原始数量) = 5
        If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
            cboEnterStock.Enabled = True
            txt摘要.Enabled = True
            
            cboStock.Enabled = True
   
            .ColData(mBillCol.C_材料) = 1
            .ColData(mBillCol.C_填写数量) = 4
            .ColData(mBillCol.C_实际数量) = 5
            .ColData(mBillCol.C_灭菌日期) = 5
            
        ElseIf mint编辑状态 = 3 Or mint编辑状态 = 6 Then
            cboStock.Enabled = False
            cboEnterStock.Enabled = False
            txt摘要.Enabled = False
            
            .ColData(mBillCol.C_材料) = 5
            .ColData(mBillCol.C_填写数量) = 5
            .ColData(mBillCol.C_实际数量) = 4

        ElseIf mint编辑状态 = 4 Then
            cboStock.Enabled = False
            cboEnterStock.Enabled = False
            txt摘要.Enabled = False
            
            .ColData(mBillCol.C_填写数量) = 5
            .ColData(mBillCol.C_实际数量) = 5
            .ColData(mBillCol.C_材料) = 5
        End If
        
        .ColData(mBillCol.C_采购价) = 5
        .ColData(mBillCol.C_采购金额) = 5
        .ColData(mBillCol.C_售价) = 5
        .ColData(mconintcol售价金额) = 5
        .ColData(mconintcol差价) = 5
        
        .ColData(mBillCol.C_库房分批) = 5
        .ColData(mBillCol.C_可用数量) = 5
        .ColData(mBillCol.C_最大效期) = 5
        .ColData(mBillCol.C_实际差价) = 5
        .ColData(mBillCol.C_实际金额) = 5
        .ColData(mBillCol.C_指导差价率) = 5
        .ColData(mBillCol.C_比例系数) = 5
        .ColData(mBillCol.c_批次) = 5
        
        .ColAlignment(mBillCol.C_材料) = flexAlignLeftCenter
        .ColAlignment(mBillCol.c_规格) = flexAlignLeftCenter
        .ColAlignment(mBillCol.C_产地) = flexAlignLeftCenter
        .ColAlignment(mBillCol.C_批准文号) = flexAlignLeftCenter
        .ColAlignment(mBillCol.c_单位) = flexAlignCenterCenter
        .ColAlignment(mBillCol.c_批号) = flexAlignLeftCenter
        .ColAlignment(mBillCol.C_效期) = flexAlignLeftCenter
        .ColAlignment(mBillCol.C_填写数量) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_实际数量) = flexAlignRightCenter
        
        .ColAlignment(mBillCol.C_采购价) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_采购金额) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_售价) = flexAlignRightCenter
        .ColAlignment(mconintcol售价金额) = flexAlignRightCenter
        .ColAlignment(mconintcol差价) = flexAlignRightCenter
        
        .ColAlignment(mBillCol.C_一次性材料) = flexAlignLeftCenter
        .ColAlignment(mBillCol.C_灭菌效期) = flexAlignLeftCenter
        .ColAlignment(mBillCol.C_灭菌失效期) = flexAlignCenterCenter
        .ColAlignment(mBillCol.C_灭菌日期) = flexAlignCenterCenter
        
        .PrimaryCol = mBillCol.C_材料
        .LocateCol = mBillCol.C_材料
        If InStr(1, "346", mint编辑状态) <> 0 Then .ColData(mBillCol.C_材料) = 0
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
    With txtNO
        .Left = mshBill.Left + mshBill.Width - .Width
        LblNo.Left = .Left - LblNo.Width - 100
        .Top = LblTitle.Top
        LblNo.Top = .Top
    End With
    
    
    LblStock.Left = mshBill.Left
    cboStock.Left = LblStock.Left + LblStock.Width + 100
    
    cboEnterStock.Left = mshBill.Left + mshBill.Width - cboEnterStock.Width
    
    LblEnterStock.Left = cboEnterStock.Left - LblEnterStock.Width - 100
    
    
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
        
    With cmdFind
        .Top = CmdCancel.Top
    End With
    
    With cmdRequestTransfer
        .Top = cmdFind.Top
        
        .Visible = (mint编辑状态 = 1 Or mint编辑状态 = 2) '新增和修改才可见
        
    End With
    
    With cmdExpend
        .Top = CmdSave.Top
        .Left = CmdSave.Left - 150 - .Width
    End With
    
    With lblCode
        .Top = CmdCancel.Top + 50
    End With
    With txtCode
        .Top = CmdCancel.Top + 30
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If msh产地.Visible = True Then '产地的列表打开则先关闭产地的列表
        msh产地.Visible = False
        mshBill.SetFocus
        mshBill.Col = mBillCol.C_产地
        Cancel = True
        Exit Sub
    End If
    
    If mblnChange = False Or mint编辑状态 = 4 Or mint编辑状态 = 3 Then
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

Private Function SaveCheck(Optional ByVal bln强制保存 As Boolean = False) As Boolean
    Dim rs类别 As New Recordset
    Dim intRow As Integer
    
    Dim strNo As String
    Dim lng库房ID As Long
    Dim lng对方部门id As Long
    Dim str审核人 As String
    
    Dim lng材料ID As Long
    Dim str产地 As String
    Dim lng出批次 As Long
    Dim dbl填写数量 As Double
    Dim dbl实际数量 As Double
    Dim dbl成本价 As Double
    Dim dbl成本金额 As Double
    Dim dbl售价 As Double
    Dim dbl零售金额 As Double
    Dim dbl差价 As Double
    Dim lng出类别id As Long
    Dim lng入类别id As Long
    Dim str批号 As String
    Dim str效期 As String
    Dim str审核日期 As String
    Dim str灭菌日期 As String
    Dim int序列号 As Integer
    Dim n As Long
    
    Dim arrSQL As Variant
    
    On Error GoTo ErrHandle
    arrSQL = Array()
    mblnSave = False
    SaveCheck = False
    
    '检查该单据是否在进入编辑界面后，被其他操作员修改
    mstrTime_End = GetBillInfo(19, mstr单据号)
    If mstrTime_End = "" Then
        MsgBox "该单据已经被其他操作员删除！", vbInformation, gstrSysName
        Exit Function
    End If
    If Not bln强制保存 And mint移库处理流程 <> 0 Then
        If mstrTime_End > mstrTime_Start Then
            MsgBox "该单据已经被其他操作员编辑，请退出后重试！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    lng库房ID = cboStock.ItemData(cboStock.ListIndex)
    lng对方部门id = cboEnterStock.ItemData(cboEnterStock.ListIndex)
    str审核人 = UserInfo.用户名
    strNo = txtNO.Tag
    
    gstrSQL = "" & _
        "   SELECT b.系数,b.id AS 类别id " & _
        "   FROM 药品单据性质 a, 药品入出类别 b " & _
        "   Where a.类别id = b.ID " & _
        "           AND a.单据 = 34 "
    
    zlDatabase.OpenRecordset rs类别, gstrSQL, mstrCaption
    
    If rs类别.EOF Then
        MsgBox "卫材入出分类不全，请检查!", vbExclamation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If rs类别.RecordCount < 2 Then
        MsgBox "卫材入出分类不全，请检查!", vbExclamation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    rs类别.MoveFirst
    Do While Not rs类别.EOF
        If rs类别!系数 = 1 Then
            lng入类别id = rs类别!类别ID
        Else
            lng出类别id = rs类别!类别ID
        End If
        rs类别.MoveNext
    Loop
    rs类别.Close
    
    str审核日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
    With mshBill
        '按药品ID顺序更新数据
        recSort.Sort = "药品id,批次,序号"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!行号
'        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                lng材料ID = .TextMatrix(intRow, 0)
                str产地 = .TextMatrix(intRow, mBillCol.C_产地)
                lng出批次 = .TextMatrix(intRow, mBillCol.c_批次)
                dbl填写数量 = Round(Val(.TextMatrix(intRow, mBillCol.C_填写数量)) * .TextMatrix(intRow, mBillCol.C_比例系数), g_小数位数.obj_最大小数.数量小数)
                dbl实际数量 = Round(Val(.TextMatrix(intRow, mBillCol.C_实际数量)) * .TextMatrix(intRow, mBillCol.C_比例系数), g_小数位数.obj_最大小数.数量小数)
                If Val(Format(Val(.TextMatrix(intRow, mBillCol.c_原始数量)) / Val(.TextMatrix(intRow, mBillCol.C_比例系数)), mFMT.FM_数量)) = Val(.TextMatrix(intRow, mBillCol.C_填写数量)) Then
                    If dbl填写数量 = dbl实际数量 Then
                        dbl实际数量 = Val(.TextMatrix(intRow, mBillCol.c_原始数量))
                        dbl填写数量 = dbl实际数量
                    End If
                End If
                
                dbl成本价 = Round(Val(.TextMatrix(intRow, mBillCol.C_采购价)) / .TextMatrix(intRow, mBillCol.C_比例系数), g_小数位数.obj_最大小数.成本价小数)
                dbl成本金额 = Round(Val(.TextMatrix(intRow, mBillCol.C_采购金额)), g_小数位数.obj_最大小数.金额小数)
                dbl售价 = Round(Val(.TextMatrix(intRow, mBillCol.C_售价)) / .TextMatrix(intRow, mBillCol.C_比例系数), g_小数位数.obj_最大小数.零售价小数)
                dbl零售金额 = Round(Val(.TextMatrix(intRow, mconintcol售价金额)), g_小数位数.obj_最大小数.金额小数)
                dbl差价 = Round(Val(.TextMatrix(intRow, mconintcol差价)), g_小数位数.obj_最大小数.金额小数)
                str批号 = .TextMatrix(intRow, mBillCol.c_批号)
                str效期 = IIf(.TextMatrix(intRow, mBillCol.C_效期) = "", "Null", "to_date('" & .TextMatrix(intRow, mBillCol.C_效期) & "','yyyy-mm-dd')")
                str灭菌日期 = IIf(.TextMatrix(intRow, mBillCol.C_灭菌失效期) = "", "Null", "to_date('" & .TextMatrix(intRow, mBillCol.C_灭菌失效期) & "','yyyy-mm-dd')")
                int序列号 = Val(.TextMatrix(intRow, mBillCol.c_序号))
                'zl_材料移库_VERIFY( /*库房ID_IN*/, /*对方部门ID_IN*/, /*材料ID_IN*/,
                    '产地_IN*/, /*出批次_IN*/, /*填写数量_IN*/, /*实际数量_IN*/, /*成本价_IN*/,
                    '/*成本金额_IN*/, /*零售金额_IN*/, /*差价_IN*/, /*出类别ID_IN*/, /*入类别ID_IN*/,
                    '/*NO_IN*/, /*审核人_IN*/, /*批号_IN*/, /*效期_IN*/灭菌失效期/审核日期 ,移库单标志);
                        
                gstrSQL = "" & _
                    "zl_材料移库_Verify(" & int序列号 & "," & lng库房ID & "," & lng对方部门id & "," & _
                     lng材料ID & ",'" & str产地 & "'," & lng出批次 & "," & dbl填写数量 & "," & _
                     dbl实际数量 & "," & dbl成本价 & "," & dbl成本金额 & "," & dbl零售金额 & "," & _
                     dbl差价 & "," & lng出类别id & "," & lng入类别id & ",'" & _
                     strNo & "','" & str审核人 & "','" & str批号 & "'," & str效期 & "," & str灭菌日期 & ",to_date('" & str审核日期 & "','yyyy-mm-dd HH24:MI:SS')," & IIf(mbln申领单 = True, 0, 1) & "," & dbl售价 & ")"
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = CStr(lng材料ID) & ";" & vbCrLf & gstrSQL
            End If
            recSort.MoveNext
        Next
    End With
    
    If Not ExecuteSql(arrSQL, mstrCaption, False, Not bln强制保存) Then Exit Function
'    If Not 检查单价(19, txtNO.Tag) Then
'
'        If Not bln强制保存 Then gcnOracle.RollbackTrans
'        Exit Function
'    End If
    If Not bln强制保存 Then gcnOracle.CommitTrans
    
    SaveCheck = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

Private Sub mshBill_AfterAddRow(Row As Long)
    Call RefreshRowNO(mshBill, mBillCol.C_行号, Row)
End Sub

Private Sub mshBill_AfterDeleteRow()
    Call 显示合计金额
    Call RefreshRowNO(mshBill, mBillCol.C_行号, mshBill.Row)
    If mbln申领单 Then Call ShowColor
End Sub



Private Sub mshBill_BeforeAddRow(Row As Long)
    If mshBill.ColData(mBillCol.C_材料) = 0 Then
        Exit Sub
    End If
        
        
End Sub

Private Sub mshBill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    If mint编辑状态 = 10 Then
        Cancel = True
        Exit Sub
    End If
    If InStr(1, "34", mint编辑状态) <> 0 Then
        If mint编辑状态 = 3 And mbln申领单 Then Exit Sub
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

Private Sub mshbill_CommandClick()
    Dim RecReturn As Recordset
    Dim i As Integer
    Dim int点击行 As Integer
    Dim rsTemp As Recordset
    
    On Error GoTo ErrHandle
    
    int点击行 = mshBill.Row
    
    If cboEnterStock.ListCount = 0 Then Exit Sub
    
    If mshBill.Col = mBillCol.C_材料 Then
        mbln仅显示有库存物资 = gSystem_Para.para_卫材填单下可用库存 And mint库存检查 = 2
        If Not mbln申领单 Then
            Set RecReturn = Frm材料选择器.ShowMe(Me, 2, cboStock.ItemData(cboStock.ListIndex), cboEnterStock.ItemData(cboEnterStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), _
                mbln移库明确批次, True, False, False, True, , , , , mbln仅显示有库存物资, , , mstrPrivs, mbln移库明确批次, False)
        Else
            Set RecReturn = Frm材料选择器.ShowMe(Me, 2, cboStock.ItemData(cboStock.ListIndex), cboEnterStock.ItemData(cboEnterStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), _
                mbln明确批次, mbln明确批次, False, False, True, , , , , mbln仅显示有库存物资, , , mstrPrivs, IIf(mbln申领单 = True, mbln明确批次, True), False)
        End If
        If RecReturn.RecordCount > 0 Then
            mblnChange = True
            With mshBill
                Dim intUnit As Integer
                
                RecReturn.MoveFirst
                For i = 1 To RecReturn.RecordCount
                    If SetColValue(.Row, RecReturn!材料ID, "[" & RecReturn!编码 & "]" & RecReturn!名称, _
                        IIf(IsNull(RecReturn!规格), "", RecReturn!规格), IIf(IsNull(RecReturn!产地), "", RecReturn!产地), _
                        IIf(mintUnit = 0, RecReturn!散装单位, RecReturn!包装单位), _
                        RecReturn!售价, IIf(IsNull(RecReturn!批号), "", RecReturn!批号), _
                        IIf(IsNull(RecReturn!效期), "", Format(RecReturn!效期, "yyyy-MM-dd")), _
                        IIf(IsNull(RecReturn!灭菌失效期), "", Format(RecReturn!灭菌失效期, "yyyy-MM-dd")), _
                        IIf(IsNull(RecReturn!最大效期), "0", RecReturn!最大效期), _
                        RecReturn!库房分批, _
                        IIf(IsNull(RecReturn!可用数量), "0", RecReturn!可用数量), _
                        IIf(IsNull(RecReturn!实际金额), "0", RecReturn!实际金额), _
                        IIf(IsNull(RecReturn!实际差价), "0", RecReturn!实际差价), _
                        IIf(IsNull(RecReturn!指导差价率), "0", RecReturn!指导差价率), _
                        IIf(mintUnit = 0, 1, RecReturn!换算系数), IIf(IsNull(RecReturn!批次), 0, RecReturn!批次), RecReturn!时价, RecReturn!在用分批, IIf(IsNull(RecReturn!批准文号), "", RecReturn!批准文号)) Then
                        
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
                
    '            If RecReturn.RecordCount = 1 Then
    '                SetColValue .Row, RecReturn!材料ID, "[" & RecReturn!编码 & "]" & RecReturn!名称, _
    '                    IIf(IsNull(RecReturn!规格), "", RecReturn!规格), IIf(IsNull(RecReturn!产地), "", RecReturn!产地), _
    '                    IIf(mintUnit = 0, RecReturn!散装单位, RecReturn!包装单位), _
    '                    RecReturn!售价, IIf(IsNull(RecReturn!批号), "", RecReturn!批号), _
    '                    IIf(IsNull(RecReturn!效期), "", Format(RecReturn!效期, "yyyy-MM-dd")), _
    '                    IIf(IsNull(RecReturn!灭菌失效期), "", Format(RecReturn!灭菌失效期, "yyyy-MM-dd")), _
    '                    IIf(IsNull(RecReturn!最大效期), "0", RecReturn!最大效期), _
    '                    RecReturn!库房分批, _
    '                    IIf(IsNull(RecReturn!可用数量), "0", RecReturn!可用数量), _
    '                    IIf(IsNull(RecReturn!实际金额), "0", RecReturn!实际金额), _
    '                    IIf(IsNull(RecReturn!实际差价), "0", RecReturn!实际差价), _
    '                    IIf(IsNull(RecReturn!指导差价率), "0", RecReturn!指导差价率), _
    '                    IIf(mintUnit = 0, 1, RecReturn!换算系数), IIf(IsNull(RecReturn!批次), 0, RecReturn!批次), RecReturn!时价, RecReturn!在用分批, IIf(IsNull(RecReturn!批准文号), "", RecReturn!批准文号)
    '
    '                .Col = mBillCol.C_填写数量
    '            End If
            End With
            RecReturn.Close
        End If
    Else
        gstrSQL = "Select rownum as id,null as 上级id,编码,名称,简码,1 as 末级 From 材料生产商 "
        Set rsTemp = zlDatabase.ShowSelect(Me, gstrSQL, 1, "材料生产商选择", True, , "选择卫生材料生产商或厂牌")
        
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
        If rsTemp Is Nothing Then Exit Sub
        If rsTemp.State <> 1 Then Exit Sub
        
        With rsTemp
            If CheckQualifications(mlngModule, 1, CStr(NVL(!名称))) = False Then Exit Sub
            mshBill.TextMatrix(mshBill.Row, mBillCol.C_产地) = NVL(!名称)
        End With
        
        gstrSQL = "select 批准文号 from 药品生产商对照 where 厂家名称=[1] and 药品id=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "mshbill_CommandClick", mshBill.TextMatrix(mshBill.Row, mBillCol.C_产地), mshBill.TextMatrix(mshBill.Row, 0))
        If rsTemp.RecordCount > 0 Then
            mshBill.TextMatrix(mshBill.Row, mBillCol.C_批准文号) = IIf(IsNull(rsTemp!批准文号), "", rsTemp!批准文号)
        Else
            mshBill.TextMatrix(mshBill.Row, mBillCol.C_批准文号) = ""
        End If

    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mshbill_EditChange(curText As String)
    mblnChange = True
End Sub


Private Sub mshBill_EditKeyPress(KeyAscii As Integer)
    Dim strKey As String
    Dim intDigit As Integer
    
    With mshBill
        If .Col = mBillCol.C_填写数量 Or mBillCol.C_实际数量 Then
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
    With mshBill
        If Row > 0 Then
            .SetRowColor CLng(Row), &HFFCECE, True
        End If
        If .Row <> .LastRow Or .LastRow = 1 Then 'Or .LastRow = 1加这个是因为第一次进来.Row 、 .LastRow 都 = 1
            SetInputFormat .Row
        End If
        
        Select Case .Col
            Case mBillCol.C_材料
                .TxtCheck = False
                .MaxLength = 80
                '只在药名列才显示合计信息和库存数
                Call 显示合计金额
                Call 提示库存数
                
            Case mBillCol.c_批号
                .TxtCheck = False
                .MaxLength = mintBatchNoLen
            
            Case mBillCol.C_效期
                .TxtCheck = True
                .TextMask = "1234567890-"
                .MaxLength = 10
                If .TextMatrix(.Row, mBillCol.c_批号) <> "" And .ColData(.Col) = 2 Then
                    Dim strxq As String
                    
                    If IsNumeric(.TextMatrix(.Row, mBillCol.c_批号)) And .TextMatrix(.Row, mBillCol.C_最大效期) <> "" Then
                        If Split(.TextMatrix(.Row, mBillCol.C_最大效期), "||")(0) <> 0 Then
                            strxq = UCase(.TextMatrix(.Row, mBillCol.c_批号))
                            If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                                strxq = TranNumToDate(strxq)
                                If strxq = "" Then Exit Sub
                                
                                .TextMatrix(.Row, mBillCol.C_效期) = Format(DateAdd("M", Split(.TextMatrix(.Row, mBillCol.C_最大效期), "||")(0), strxq), "yyyy-mm-dd")
                            End If
                        End If
                    End If
                End If
            Case mBillCol.C_填写数量, mBillCol.C_实际数量
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
            Case mBillCol.C_产地
                ImeLanguage True
                .TxtCheck = False
                .MaxLength = 30
                .TxtSetFocus
        End Select
        
    End With
End Sub

Private Sub mshBill_GotFocus()
    If mintParallelRecord <> 1 Then Exit Sub
    If cboEnterStock.ListCount = 0 Then Exit Sub
    If cboStock.ListIndex < 0 Then Exit Sub
    If cboStock.ItemData(cboStock.ListIndex) = cboEnterStock.ItemData(cboEnterStock.ListIndex) Then
        MsgBox "移入库房和移出库房相同了，请检查后重新选择！", vbOKOnly + vbExclamation, gstrSysName
       If cboEnterStock.Enabled Then cboEnterStock.SetFocus
        Exit Sub
    End If
    
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim rsDrug As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim i As Integer
    Dim int点击行 As Integer
    Dim rsTemp As New Recordset
    
    int点击行 = mshBill.Row
    
    On Error GoTo ErrHandle
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cboEnterStock.ListCount = 0 Then Exit Sub
    With mshBill
'        .Text = UCase(Trim(.Text))
        strKey = Trim(.Text)
        
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
                    mbln仅显示有库存物资 = gSystem_Para.para_卫材填单下可用库存 And mint库存检查 = 2

                    If Not mbln申领单 Then
                        Set RecReturn = FrmMulitSel.ShowSelect(Me, 2, cboStock.ItemData(cboStock.ListIndex), cboEnterStock.ItemData(cboEnterStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), _
                            strKey, sngLeft, sngTop, mshBill.MsfObj.CellWidth, mshBill.MsfObj.CellHeight, mbln移库明确批次, True, False, False, True, , , , mbln仅显示有库存物资, , , mstrPrivs, mbln移库明确批次, False)
                    Else
                        Set RecReturn = FrmMulitSel.ShowSelect(Me, 2, cboStock.ItemData(cboStock.ListIndex), cboEnterStock.ItemData(cboEnterStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), _
                            strKey, sngLeft, sngTop, mshBill.MsfObj.CellWidth, mshBill.MsfObj.CellHeight, mbln明确批次, mbln明确批次, False, False, True, , , , mbln仅显示有库存物资, , , mstrPrivs, IIf(mbln申领单 = True, mbln明确批次, True), False)
                    End If
                    
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
                                IIf(IsNull(RecReturn!最大效期), "0", RecReturn!最大效期), _
                                RecReturn!库房分批, _
                                IIf(IsNull(RecReturn!可用数量), "0", RecReturn!可用数量), _
                                IIf(IsNull(RecReturn!实际金额), "0", RecReturn!实际金额), _
                                IIf(IsNull(RecReturn!实际差价), "0", RecReturn!实际差价), _
                                IIf(IsNull(RecReturn!指导差价率), "0", RecReturn!指导差价率), _
                                IIf(mintUnit = 0, 1, RecReturn!换算系数), IIf(IsNull(RecReturn!批次), 0, RecReturn!批次), RecReturn!时价, RecReturn!在用分批, IIf(IsNull(RecReturn!批准文号), "", RecReturn!批准文号)) Then
                            
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
                    
                    If mbln移库明确批次 = False Then
                        .Col = mBillCol.C_填写数量
                    End If
'                    If RecReturn.RecordCount = 1 Then
'                        If SetColValue(.Row, RecReturn!材料ID, "[" & RecReturn!编码 & "]" & RecReturn!名称, _
'                                IIf(IsNull(RecReturn!规格), "", RecReturn!规格), IIf(IsNull(RecReturn!产地), "", RecReturn!产地), _
'                                IIf(mintUnit = 0, RecReturn!散装单位, RecReturn!包装单位), _
'                                IIf(IsNull(RecReturn!售价), 0, RecReturn!售价), IIf(IsNull(RecReturn!批号), "", RecReturn!批号), _
'                                IIf(IsNull(RecReturn!效期), "", Format(RecReturn!效期, "yyyy-MM-dd")), _
'                                IIf(IsNull(RecReturn!灭菌失效期), "", Format(RecReturn!灭菌失效期, "yyyy-MM-dd")), _
'                                IIf(IsNull(RecReturn!最大效期), "0", RecReturn!最大效期), _
'                                RecReturn!库房分批, _
'                                IIf(IsNull(RecReturn!可用数量), "0", RecReturn!可用数量), _
'                                IIf(IsNull(RecReturn!实际金额), "0", RecReturn!实际金额), _
'                                IIf(IsNull(RecReturn!实际差价), "0", RecReturn!实际差价), _
'                                IIf(IsNull(RecReturn!指导差价率), "0", RecReturn!指导差价率), _
'                                IIf(mintUnit = 0, 1, RecReturn!换算系数), IIf(IsNull(RecReturn!批次), 0, RecReturn!批次), RecReturn!时价, RecReturn!在用分批, IIf(IsNull(RecReturn!批准文号), "", RecReturn!批准文号)) = False Then
'                            Cancel = True
'                            Exit Sub
'                        End If
'                        .Text = .TextMatrix(.Row, .Col)
'                    Else
'                        Cancel = True
'                    End If
                    Call 提示库存数
                End If
            Case mBillCol.c_批号
                '无处理
                If strKey = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mBillCol.c_批号) = ""
                    End If
                    If .ColData(mBillCol.C_效期) = 2 Then
                        .Col = mBillCol.C_效期
                    Else
                        .Col = mBillCol.C_填写数量
                    End If
                    
                    
                    Cancel = True
                    Exit Sub
                End If
                
            Case mBillCol.C_效期
                '有处理
                If strKey <> "" Then
                    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
                        strKey = TranNumToDate(strKey)
                        If strKey = "" Then
                            MsgBox "失效期必须为日期型！", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            Exit Sub
                        End If
                        .Text = strKey
                        Exit Sub
                    End If
                    If Not IsDate(strKey) Then
                        MsgBox "失效期必须为日期型如(2000-10-10) 或（20001010）,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                ElseIf strKey = "" And strKey <> .TextMatrix(.Row, mBillCol.C_效期) Then
                
                    If .TxtVisible = True Then
                        .Text = " "
                        Exit Sub
                    End If
                    
                    Exit Sub
                End If
            Case mBillCol.C_灭菌日期
                '有处理
                If strKey <> "" Then
                    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
                        strKey = TranNumToDate(strKey)
                        If strKey = "" Then
                            MsgBox "灭菌日期必须为日期型！", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            Exit Sub
                        End If
                        .Text = strKey
                        'Exit Sub
                    End If
                    If Not IsDate(strKey) Then
                        MsgBox "灭菌日期必须为日期型如(2000-10-10) 或（20001010）,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                    If Format(sys.Currentdate, "yyyy-mm-dd") >= Format(DateAdd("m", Val(.TextMatrix(.Row, mBillCol.C_灭菌效期)), CDate(strKey)), "yyyy-mm-dd") Then
                        If MsgBox("该卫材已经过了灭菌失效期(" & Format(DateAdd("m", Val(.TextMatrix(.Row, mBillCol.C_灭菌效期)), CDate(strKey)), "yyyy-mm-dd") & "),是否还要进行入库!", vbQuestion + vbDefaultButton2 + vbYesNo) = vbNo Then
                            Cancel = True
                            Exit Sub
                        End If
                    End If
                    
                    .Text = strKey
                    '计算失效期
                    .TextMatrix(.Row, mBillCol.C_灭菌失效期) = Format(DateAdd("m", Val(.TextMatrix(.Row, mBillCol.C_灭菌效期)), CDate(strKey)), "yyyy-mm-dd")
                ElseIf strKey = "" And strKey <> .TextMatrix(.Row, mBillCol.C_灭菌日期) Then
                    If .TxtVisible = True Then
                        .Text = " "
                        Exit Sub
                    End If
                    Exit Sub
                End If
            Case mBillCol.C_填写数量, mBillCol.C_实际数量
                If .TextMatrix(.Row, 0) = "" Then .Text = "": .TextMatrix(.Row, mBillCol.C_填写数量) = "": Exit Sub
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
                        MsgBox "数量必须大于零,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Not CompareUsableQuantity(.Row, strKey) Then
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    '成本价的公式：     出库金额=数量*售价
                    '                  出库差价=出库金额*（实际差价/实际金额）
                    '                  if 实际金额<=0 then  出库差价=出库金额*指导差价率
                    '                  购价（成本价）=（出库金额-出库差价）/数量
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    
                    strKey = Format(strKey, mFMT.FM_数量)
                    .Text = strKey
                    .TextMatrix(.Row, .Col) = .Text
                    
                    If .TextMatrix(.Row, mBillCol.C_售价) <> "" Then
                        .TextMatrix(.Row, mconintcol售价金额) = Format(.TextMatrix(.Row, mBillCol.C_售价) * strKey, mFMT.FM_金额)
                    End If
                    
                    If mint编辑状态 <> 6 Then
                        Dim dbl差价 As Double, dbl购价 As Double, dbl成本金额 As Double
                        'cboStock.ItemData(cboStock.ListIndex), lng材料ID, lng批次
'                        Call 验证出库差价计算(cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mBillCol.c_批次)), Val(.TextMatrix(.Row, mBillCol.C_比例系数)), Val(.TextMatrix(.Row, mBillCol.C_实际差价)), Val(.TextMatrix(.Row, mBillCol.C_实际金额)), Val(Split(.TextMatrix(.Row, mBillCol.C_指导差价率), "||")(0)) / 100, Val(strKey), Val(.TextMatrix(.Row, mBillCol.C_售价金额)), dbl差价, dbl购价, dbl成本金额)
'                        .TextMatrix(.Row, mBillCol.C_差价) = Format(dbl差价, mFMT.FM_金额)
                        .TextMatrix(.Row, mBillCol.C_采购价) = Format(Get成本价(Val(.TextMatrix(.Row, 0)), cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, mBillCol.c_批次))) * Val(.TextMatrix(.Row, mBillCol.C_比例系数)), mFMT.FM_成本价)
'                        .TextMatrix(.Row, mBillCol.C_采购金额) = Format(dbl成本金额, mFMT.FM_金额)
'                    Else
'                        .TextMatrix(.Row, mBillCol.C_采购金额) = Format(Val(.TextMatrix(.Row, mBillCol.C_采购价)) * strKey, mFMT.FM_金额)
'                        .TextMatrix(.Row, mBillCol.C_差价) = Format(Val(.TextMatrix(.Row, mBillCol.C_售价金额)) - Val(.TextMatrix(.Row, mBillCol.C_采购金额)), mFMT.FM_金额)
                    End If
                    .TextMatrix(.Row, mBillCol.C_采购金额) = Format(Val(.TextMatrix(.Row, mBillCol.C_采购价)) * strKey, mFMT.FM_金额)
                    .TextMatrix(.Row, mconintcol差价) = Format(Val(.TextMatrix(.Row, mconintcol售价金额)) - Val(.TextMatrix(.Row, mBillCol.C_采购金额)), mFMT.FM_金额)
                 
                    If .Col = mBillCol.C_填写数量 Then
                        .TextMatrix(.Row, mBillCol.C_实际数量) = strKey
                    End If
                End If
                显示合计金额
                If mbln申领单 Then Call ShowColor(.Row)
            Case mBillCol.C_产地
                '无处理
                If .Text = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mBillCol.C_产地) = ""
                    End If
                    .Col = mBillCol.c_批号
                    Cancel = True
                    Exit Sub
                Else
                    Dim rs产地 As New Recordset
                    
                    gstrSQL = "" & _
                        "   Select 编码,简码,名称 From 材料生产商 " & _
                        "   Where upper(名称) like [1] or Upper(编码) like [1] or Upper(简码) like [1]"
                    
                    Set rs产地 = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, IIf(gstrMatchMethod = "0", "%", "") & UCase(Trim(strKey)) & "%")
                    
                    
                    If rs产地.EOF Then
                        If MsgBox("卫生材料生产商没有找到你输入的产地，你要把它加入卫生材料生产商中吗？", vbYesNo + vbQuestion, mstrCaption) = vbNo Then
                            Cancel = True
                            Exit Sub
                        Else
                            Dim rsMax As New Recordset
                            Dim int编码 As Integer, strCode As String, strSpecify As String
                            
                            If rsMax.State = 1 Then rsMax.Close
                            gstrSQL = "SELECT Nvl(MAX(LENGTH(编码)),2) As Length FROM 材料生产商"
                            zlDatabase.OpenRecordset rsMax, gstrSQL, mstrCaption
                            int编码 = rsMax!Length
                            
                            gstrSQL = "SELECT Nvl(MAX(LPAD(编码," & int编码 & ",'0')),'00') As Code FROM 材料生产商"
                            rsMax.Close
                            zlDatabase.OpenRecordset rsMax, gstrSQL, mstrCaption
                            strCode = rsMax!Code
                            
                            int编码 = Len(strCode)
                            strCode = strCode + 1
                            
                            If int编码 >= Len(strCode) Then
                                strCode = String(int编码 - Len(strCode), "0") & strCode
                            End If
                            strSpecify = zlCommFun.SpellCode(strKey)
                            
                            
                            gstrSQL = "ZL_材料生产商_INSERT('" & strCode & "','" & strKey & "','" & strSpecify & "')"
                            Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
                        End If
                    Else
                        If rs产地.RecordCount = 1 Then
                            If CheckQualifications(mlngModule, 1, rs产地.Fields("名称")) = False Then
                                Exit Sub
                            End If
                            
                            .TextMatrix(.Row, mBillCol.C_产地) = rs产地.Fields("名称")
                            .Text = rs产地.Fields("名称")
                            
                            gstrSQL = "select 批准文号 from 药品生产商对照 where 厂家名称=[1] and 药品id=[2]"
                            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "mshbill_CommandClick", .TextMatrix(.Row, mBillCol.C_产地), Val(.TextMatrix(.Row, 0)))
                            If rsTemp.RecordCount > 0 Then
                                .TextMatrix(.Row, mBillCol.C_批准文号) = IIf(IsNull(rsTemp!批准文号), "", rsTemp!批准文号)
                            Else
                                .TextMatrix(.Row, mBillCol.C_批准文号) = ""
                            End If
                        Else
                            Set msh产地.Recordset = rs产地
                            With msh产地
                                .Redraw = False
                                .Left = Pic单据.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                                .Top = Pic单据.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight
                                .Visible = True
                                .SetFocus
                                .ColWidth(0) = 800
                                .ColWidth(1) = 800
                                .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(0)
                                .Row = 1
                                .Col = 0
                                .TopRow = 1
                                .ColSel = .Cols - 1
                                .Redraw = True
                                Cancel = True
                                Exit Sub
                            End With
                        End If
                    End If
                End If
                zlCommFun.OpenIme False
        End Select
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'从卫材目录中取值并附给相应的列
Private Function SetColValue(ByVal intRow As Integer, ByVal lng材料ID As Long, _
    ByVal str材料 As String, ByVal str规格 As String, ByVal str产地 As String, _
    ByVal str单位 As String, ByVal num售价 As Double, ByVal str批号 As String, _
    ByVal str效期 As String, ByVal str灭菌失效期 As String, ByVal int最大效期 As Integer, ByVal int库房分批 As Integer, _
    ByVal num可用数量 As Double, ByVal num实际金额 As Double, ByVal num实际差价 As Double, _
    ByVal num指导差价率 As Double, ByVal num比例系数 As Double, ByVal lng批次 As Long, _
    ByVal int是否变价 As Integer, ByVal int在用分批 As Integer, ByVal str批准文号 As String) As Boolean
    
    Dim intCount As Integer
    Dim intCol As Integer
    Dim dbltotal As Double
    Dim dblPrice As Double
    Dim intLop As Integer
    Dim rsprice As New Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim rs效期 As ADODB.Recordset
    Dim bln分批 As Boolean
    
    On Error GoTo ErrHandle
    If str灭菌失效期 <> "" Then
        If Format(str灭菌失效期, "yyyy-mm-dd") <= Format(sys.Currentdate, "yyyy-mm-dd") Then
            If MsgBox("卫材【" & str材料 & "(" & lng批次 & ")】的灭菌效期已经过期,是否还要进行移库?", vbYesNo + vbDefaultButton1 + vbQuestion) = vbNo Then
                Exit Function
            End If
        End If
    End If
    
    gstrSQL = "Select 一次性材料,灭菌效期 from 材料特性 where 材料id=[1]"
    Set rs效期 = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng材料ID)
    
    SetColValue = False
    With mshBill
        Dim lngRow As Long
        For lngRow = 1 To .Rows - 1
            If lngRow <> intRow And .TextMatrix(lngRow, 0) <> "" Then
                If .TextMatrix(lngRow, 0) = lng材料ID And Val(.TextMatrix(lngRow, mBillCol.c_批次)) = lng批次 Then
                    If UBound(Split(mstr重复卫材, "，")) < 3 Then mstr重复卫材 = mstr重复卫材 & str材料 & "，"  '最多记录三个重复的卫材
                    'Call MsgBox("卫生材料【" & str材料 & "(" & lng批次 & ")】已经存在，请合并后再增加！", vbOKOnly + vbInformation + vbDefaultButton2, gstrSysName)
                    Exit Function
                End If
            End If
        Next
        
        If lng批次 > 0 Then   '对移出库房是库房且卫材是库房分批的卫材的判断
            If mint编辑状态 = 1 Then
                dbltotal = 0
                For intLop = 1 To .Rows - 1
                    If .TextMatrix(intLop, 0) <> "" Then
                        If intLop <> intRow And .TextMatrix(intLop, 0) = .TextMatrix(intRow, 0) And lng批次 = .TextMatrix(intLop, mBillCol.c_批次) Then
                            dbltotal = dbltotal + .TextMatrix(intLop, mBillCol.C_填写数量)
                        End If
                    End If
                Next
                
                If dbltotal >= num可用数量 And dbltotal <> 0 Then
                    MsgBox "该卫材的可用库存数量已没有了，请重输！", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
                
            End If
        End If
        
        If int是否变价 = 1 Then
            If int在用分批 = 0 Then
                If int库房分批 = 1 Then
                    gstrSQL = "Select Distinct 0 " & _
                            "From 部门性质说明 " & _
                            "Where ((工作性质 Like '发料部门') Or (工作性质 Like '制剂室')) And 部门id = [1]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex))
                    If rsTemp.RecordCount = 0 Then
                        bln分批 = True
                    End If
                End If
            Else
                bln分批 = True
            End If
        
            gstrSQL = "" & _
                "   Select nvl(零售价,0)*" & num比例系数 & " as  分批售价,实际金额/实际数量* " & num比例系数 & " as 平均零售价" & _
                "   From 药品库存 " & _
                "   Where 库房id=[1]" & _
                "       and 药品id=[2]" & _
                "       and 性质=1 and 实际数量>0 and " & _
                "       nvl(批次,0)=[3]"
            
            Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex), lng材料ID, lng批次)
            If rsprice.EOF Then
                If (mbln明确批次 = True And mbln申领单 = True) Or (mbln移库明确批次 = True And mbln申领单 = False) Then 'Or (mbln申领单 = False And bln分批 = False)
                    MsgBox "时价卫材没有库存，不能出库，请检查！", vbOKOnly, gstrSysName
                    Exit Function
                ElseIf mbln明确批次 = False And mbln申领单 = True Then
                    dblPrice = num售价 * num比例系数
                ElseIf mbln移库明确批次 = False Then
                    dblPrice = Get零售价(lng材料ID, cboStock.ItemData(cboStock.ListIndex), lng批次, num比例系数)
                End If
            Else
                If bln分批 = True Then
                    dblPrice = rsprice!分批售价
                Else
                    dblPrice = rsprice!平均零售价
                End If
            End If
        End If

        For intCol = 0 To .Cols - 1
            If intCol <> mBillCol.C_行号 Then .TextMatrix(intRow, intCol) = ""
        Next
        
        .TextMatrix(intRow, mBillCol.C_行号) = intRow
        .TextMatrix(intRow, 0) = lng材料ID
        .TextMatrix(intRow, mBillCol.C_材料) = str材料
        .TextMatrix(intRow, mBillCol.c_规格) = str规格
        .TextMatrix(intRow, mBillCol.C_产地) = str产地
        .TextMatrix(intRow, mBillCol.c_单位) = str单位
        .TextMatrix(intRow, mBillCol.c_批号) = str批号
        .TextMatrix(intRow, mBillCol.C_效期) = Format(str效期, "yyyy-mm-dd")
        .TextMatrix(intRow, mBillCol.C_灭菌失效期) = Format(str灭菌失效期, "yyyy-mm-dd")
        .TextMatrix(intRow, mBillCol.C_一次性材料) = zlStr.NVL(rs效期!一次性材料)
        .TextMatrix(intRow, mBillCol.C_灭菌效期) = zlStr.NVL(rs效期!灭菌效期)
        
        .TextMatrix(intRow, mBillCol.C_售价) = Format(num售价 * num比例系数, mFMT.FM_零售价)
        .TextMatrix(intRow, mBillCol.C_库房分批) = int库房分批
        .TextMatrix(intRow, mBillCol.C_可用数量) = Format(num可用数量 / num比例系数, mFMT.FM_数量)
        .TextMatrix(intRow, mBillCol.C_最大效期) = int最大效期 & "||" & int是否变价 & "||" & int在用分批
        .TextMatrix(intRow, mBillCol.C_实际差价) = num实际差价
        .TextMatrix(intRow, mBillCol.C_实际金额) = num实际金额
        .TextMatrix(intRow, mBillCol.C_指导差价率) = num指导差价率
        .TextMatrix(intRow, mBillCol.C_比例系数) = num比例系数
        .TextMatrix(intRow, mBillCol.C_批准文号) = str批准文号
        
        If (mbln明确批次 = True And mbln申领单 = True) Or mbln申领单 = False Then
            .TextMatrix(intRow, mBillCol.c_批次) = lng批次
        Else
            .TextMatrix(intRow, mBillCol.c_批次) = 0
        End If
        If int是否变价 = 1 Then
            .TextMatrix(intRow, mBillCol.C_售价) = Format(dblPrice, mFMT.FM_零售价)
        End If
        Call CheckLapse(str效期)
        SetInputFormat intRow
        
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

'从卫材目录中取值并附给相应的列
Private Function SetRequestColValue(ByVal intRow As Integer, ByVal lng材料ID As Long, _
    ByVal str材料 As String, ByVal str规格 As String, ByVal str产地 As String, _
    ByVal str单位 As String, ByVal num售价 As Double, ByVal str批号 As String, _
    ByVal str效期 As String, ByVal str灭菌失效期 As String, ByVal int最大效期 As Integer, ByVal int库房分批 As Integer, _
    ByVal num可用数量 As Double, ByVal num实际金额 As Double, ByVal num实际差价 As Double, _
    ByVal num指导差价率 As Double, ByVal num比例系数 As Double, ByVal lng批次 As Long, _
    ByVal int是否变价 As Integer, ByVal int在用分批 As Integer, ByVal str批准文号 As String) As Boolean
    
    Dim intCount As Integer
    Dim intCol As Integer
    Dim dbltotal As Double
    Dim dblPrice As Double
    Dim intLop As Integer
    Dim rsprice As New Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim rs效期 As ADODB.Recordset
    Dim bln分批 As Boolean
    
    On Error GoTo ErrHandle
    If str灭菌失效期 <> "" Then
        If Format(str灭菌失效期, "yyyy-mm-dd") <= Format(sys.Currentdate, "yyyy-mm-dd") Then
            If MsgBox("卫材【" & str材料 & "(" & lng批次 & ")】的灭菌效期已经过期,是否还要进行移库?", vbYesNo + vbDefaultButton1 + vbQuestion) = vbNo Then
                Exit Function
            End If
        End If
    End If
    
    gstrSQL = "Select 一次性材料,灭菌效期 from 材料特性 where 材料id=[1]"
    Set rs效期 = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng材料ID)
    
    SetRequestColValue = False
    With mshBill
        Dim lngRow As Long
        For lngRow = 1 To .Rows - 1
            If lngRow <> intRow And .TextMatrix(lngRow, 0) <> "" Then
                If .TextMatrix(lngRow, 0) = lng材料ID And Val(.TextMatrix(lngRow, mBillCol.c_批次)) = lng批次 Then
                    If UBound(Split(mstr重复卫材, "，")) < 3 Then mstr重复卫材 = mstr重复卫材 & str材料 & "，"  '最多记录三个重复的卫材
                    'Call MsgBox("卫生材料【" & str材料 & "(" & lng批次 & ")】已经存在，请合并后再增加！", vbOKOnly + vbInformation + vbDefaultButton2, gstrSysName)
                    Exit Function
                End If
            End If
        Next
        
        If lng批次 > 0 Then   '对移出库房是库房且卫材是库房分批的卫材的判断
            If mint编辑状态 = 1 Then
                dbltotal = 0
                For intLop = 1 To .Rows - 1
                    If .TextMatrix(intLop, 0) <> "" Then
                        If intLop <> intRow And .TextMatrix(intLop, 0) = .TextMatrix(intRow, 0) And lng批次 = .TextMatrix(intLop, mBillCol.c_批次) Then
                            dbltotal = dbltotal + .TextMatrix(intLop, mBillCol.C_填写数量)
                        End If
                    End If
                Next
                
                If dbltotal >= num可用数量 And dbltotal <> 0 Then
                    MsgBox "该卫材的可用库存数量已没有了，请重输！", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
                
            End If
        End If
        
        If int是否变价 = 1 Then
            If int在用分批 = 0 Then
                If int库房分批 = 1 Then
                    gstrSQL = "Select Distinct 0 " & _
                            "From 部门性质说明 " & _
                            "Where ((工作性质 Like '发料部门') Or (工作性质 Like '制剂室')) And 部门id = [1]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex))
                    If rsTemp.RecordCount = 0 Then
                        bln分批 = True
                    End If
                End If
            Else
                bln分批 = True
            End If
        
            gstrSQL = "" & _
                "   Select nvl(零售价,0)*" & num比例系数 & " as  分批售价,实际金额/实际数量* " & num比例系数 & " as 平均零售价" & _
                "   From 药品库存 " & _
                "   Where 库房id=[1]" & _
                "       and 药品id=[2]" & _
                "       and 性质=1 and 实际数量>0 and " & _
                "       nvl(批次,0)=[3]"
            
            Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex), lng材料ID, lng批次)
            If rsprice.EOF Then
                If (mbln明确批次 = True And mbln申领单 = True) Or (mbln申领单 = False And bln分批 = False) Then
                    MsgBox "时价卫材没有库存，不能出库，请检查！", vbOKOnly, gstrSysName
                    Exit Function
                ElseIf mbln明确批次 = False And mbln申领单 = True Then
                    dblPrice = num售价 * num比例系数
                ElseIf mbln移库明确批次 = False Then
                    dblPrice = Get零售价(lng材料ID, cboStock.ItemData(cboStock.ListIndex), lng批次, num比例系数)
                End If
            Else
                If bln分批 = True Then
                    dblPrice = rsprice!分批售价
                Else
                    dblPrice = rsprice!平均零售价
                End If
            End If
        End If

        For intCol = 0 To .Cols - 1
            If intCol <> mBillCol.C_行号 Then .TextMatrix(intRow, intCol) = ""
        Next
        
        .TextMatrix(intRow, mBillCol.C_行号) = intRow
        .TextMatrix(intRow, 0) = lng材料ID
        .TextMatrix(intRow, mBillCol.C_材料) = str材料
        .TextMatrix(intRow, mBillCol.c_规格) = str规格
        .TextMatrix(intRow, mBillCol.C_产地) = str产地
        .TextMatrix(intRow, mBillCol.c_单位) = str单位
        .TextMatrix(intRow, mBillCol.c_批号) = str批号
        .TextMatrix(intRow, mBillCol.C_效期) = Format(str效期, "yyyy-mm-dd")
        .TextMatrix(intRow, mBillCol.C_灭菌失效期) = Format(str灭菌失效期, "yyyy-mm-dd")
        .TextMatrix(intRow, mBillCol.C_一次性材料) = zlStr.NVL(rs效期!一次性材料)
        .TextMatrix(intRow, mBillCol.C_灭菌效期) = zlStr.NVL(rs效期!灭菌效期)
        
        .TextMatrix(intRow, mBillCol.C_售价) = Format(num售价 * num比例系数, mFMT.FM_零售价)
        .TextMatrix(intRow, mBillCol.C_库房分批) = int库房分批
        .TextMatrix(intRow, mBillCol.C_可用数量) = Format(num可用数量, mFMT.FM_数量)
        .TextMatrix(intRow, mBillCol.C_最大效期) = int最大效期 & "||" & int是否变价 & "||" & int在用分批
        .TextMatrix(intRow, mBillCol.C_实际差价) = num实际差价
        .TextMatrix(intRow, mBillCol.C_实际金额) = num实际金额
        .TextMatrix(intRow, mBillCol.C_指导差价率) = num指导差价率
        .TextMatrix(intRow, mBillCol.C_比例系数) = num比例系数
        .TextMatrix(intRow, mBillCol.C_批准文号) = str批准文号
        '按申购单移库是明确批次的，不用判断是否按批次移库
        .TextMatrix(intRow, mBillCol.c_批次) = lng批次
        
        If int是否变价 = 1 Then
            .TextMatrix(intRow, mBillCol.C_售价) = Format(dblPrice, mFMT.FM_零售价)
        End If
        Call CheckLapse(str效期)
        SetInputFormat intRow
        
    End With
    Call 提示库存数
    SetRequestColValue = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetInputFormat(ByVal intRow As Integer)
    Dim rsData As ADODB.Recordset
    Dim bln入库库房 As Boolean, bln出库库房 As Boolean
    Dim bln库存分批 As Boolean, bln在用分批 As Boolean
    '说明：1、移出库房为库房，且为库房分批卫材，
    '         A、如果有卫材，不管它，但在控制可用数量时，填写数量大于了可用数量，不允许出库，且该点不受库存参数的影响；
    '         B、如果无，卫材选择器出不来
    '      2、移出库房不为库房，且为库房分批卫材，
    '         A、如果移入库房为库房，从卫材选择器中出来一定没有批次和效期，
    '                这时，允许输入批次和效期
    '         B、如果移入库房不为库房，则
    '                这时，不允许输入批次和效期
    '      3、卫材不为库房分批卫材
    '         这时，不允许输入批次和效期
    
'    If mblnEdit = False Then Exit Sub
    If cboEnterStock.ListCount = 0 Then Exit Sub
    If Val(mshBill.TextMatrix(intRow, 0)) = 0 Then Exit Sub
    
    With mshBill
        If .TextMatrix(intRow, mBillCol.C_库房分批) = "0" Then  '不是库房分批卫材，不允许输入，其他不管了
            .ColData(mBillCol.c_批号) = 5                    '禁止
            .ColData(mBillCol.C_效期) = 5
        Else
            If .TextMatrix(intRow, mBillCol.c_批号) = "" Then        'And GetDrugUnit(cboEnterStock.ItemData(cboEnterStock.ListIndex), mfrmMain.Caption) = "卫材库单位"
                .ColData(mBillCol.c_批号) = 4              '纯文本输入
                If .TextMatrix(intRow, mBillCol.C_最大效期) <> "" Then
                    If Split(.TextMatrix(intRow, mBillCol.C_最大效期), "||")(0) <> 0 Then
                        .ColData(mBillCol.C_效期) = 2          '日期输入框
                    Else
                        .ColData(mBillCol.C_效期) = 5
                    End If
                Else
                    .ColData(mBillCol.C_效期) = 5
                End If
            Else
                .ColData(mBillCol.c_批号) = 5              '禁止
                .ColData(mBillCol.C_效期) = 5
            End If
        End If
        If .TextMatrix(intRow, mBillCol.C_一次性材料) = "1" Then
            .ColData(mBillCol.C_灭菌日期) = 5
            .ColData(mBillCol.C_灭菌失效期) = 5
        Else
            .ColData(mBillCol.C_灭菌日期) = 5              '禁止
            .ColData(mBillCol.C_灭菌失效期) = 5
        End If
        
        '出库房批号或产地为空，入库房分批的可以对批号或产地进行编辑
        
        If mbln分批卫材批号产地控制 = True Then
            '1、查询药品库存信息
            gstrSQL = "Select 上次批号,上次产地 From 药品库存 Where 库房id=[1] And 药品id=[2] and nvl(批次,0) = [3] "
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "判断效期", Val(cboStock.ItemData(cboStock.ListIndex)), Val(.TextMatrix(intRow, 0)), Val(.TextMatrix(intRow, mBillCol.c_批次)))
            
            '2、入库房是否分批
            bln入库库房 = CheckStockProperty(cboEnterStock.ItemData(cboEnterStock.ListIndex))
            bln库存分批 = (Val(.TextMatrix(intRow, mBillCol.C_库房分批)) = 1)
            bln在用分批 = (Split(.TextMatrix(intRow, mBillCol.C_最大效期), "||")(2) = 1)
            If ((bln入库库房 And bln库存分批) Or (Not bln入库库房 And bln在用分批)) Then '入库房分批
                If (IsNull(rsData!上次批号) Or rsData.EOF) Then '出库房无库存或批号为空
                    .ColData(mBillCol.c_批号) = 4
                Else
                    .ColData(mBillCol.c_批号) = 5
                End If
                If (IsNull(rsData!上次产地) Or rsData.EOF) Then
                    .ColData(mBillCol.C_产地) = 1
                Else
                    .ColData(mBillCol.C_产地) = 5
                End If
            End If
        End If
        
    End With
End Sub

Private Sub msh产地_DblClick()
    msh产地_KeyDown vbKeyReturn, 0
End Sub

Private Sub msh产地_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsProvider As ADODB.Recordset
    
    With mshBill
    
        If KeyCode = vbKeyEscape Then
            msh产地.Visible = False
            .SetFocus
        End If
        
        If .Col = mBillCol.C_产地 Then
            If CheckQualifications(mlngModule, 1, msh产地.TextMatrix(msh产地.Row, 2)) = False Then
                Exit Sub
            End If
            
            If KeyCode = vbKeyReturn Then
                .TextMatrix(.Row, .Col) = msh产地.TextMatrix(msh产地.Row, 2)
                msh产地.Visible = False
                
                gstrSQL = "select 批准文号 from 药品生产商对照 where 厂家名称=[1] and 药品id=[2]"
                Set rsProvider = zlDatabase.OpenSQLRecord(gstrSQL, "msh产地_KeyDown", .TextMatrix(.Row, .Col), .TextMatrix(.Row, 0))
                If rsProvider.RecordCount > 0 Then
                    .TextMatrix(.Row, mBillCol.C_批准文号) = IIf(IsNull(rsProvider!批准文号), "", rsProvider!批准文号)
                Else
                    .TextMatrix(.Row, mBillCol.C_批准文号) = ""
                End If
                
                .Col = mBillCol.c_批号
                .SetFocus
            End If
        End If
    End With
End Sub

Private Sub msh产地_LostFocus()
    If msh产地.Visible Then
        msh产地.Visible = False
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
    Dim bln入库库房 As Boolean, bln出库库房 As Boolean
    Dim bln库存分批 As Boolean, bln在用分批 As Boolean
    ValidData = False
    If cboEnterStock.ListCount = 0 Then
        cboEnterStock.SetFocus
        Exit Function
    End If
    If cboStock.ListCount = 0 Then
        cboStock.SetFocus
        Exit Function
    End If
    
    bln入库库房 = CheckStockProperty(cboEnterStock.ItemData(cboEnterStock.ListIndex))
    bln出库库房 = CheckStockProperty(cboStock.ItemData(cboStock.ListIndex))

    
    ValidData = False
    
    Dim intLop As Integer
    
    If txtNO.Locked = False Then
        If Trim(txtNO.Text) = "" Then
            ShowMsgBox "单据号不能为空"
            Exit Function
        End If
        
        If InStr(1, txtNO.Text, "'") <> 0 Then
            ShowMsgBox "单据号中不能含有非法字符"
            Exit Function
        End If
        
        If LenB(StrConv(txtNO.Text, vbFromUnicode)) > txtNO.MaxLength Then
            ShowMsgBox "单据号超长,最多能输入" & CInt(txtNO.MaxLength / 2) & "个汉字（最好不要汉字）或" & txtNO.MaxLength & "个字符!"
            txtNO.SetFocus
            Exit Function
        End If
    End If
    
    With mshBill
        If .TextMatrix(1, 0) <> "" Then         '先判有否数据
            If cboEnterStock.ListCount = 0 Then
                MsgBox "请设置允许调拨的部门，[基础参数设置]中的卫材流向！", vbInformation, gstrSysName
                Exit Function
            End If
            If cboStock.ItemData(cboStock.ListIndex) = cboEnterStock.ItemData(cboEnterStock.ListIndex) Then
                MsgBox "移入库房和移出库房相同了，请重新选择！", vbInformation, gstrSysName
                If cboEnterStock.Enabled Then cboEnterStock.SetFocus
                Exit Function
            End If
            
            
            If LenB(StrConv(txt摘要.Text, vbFromUnicode)) > txt摘要.MaxLength Then
                MsgBox "摘要超长,最多能输入" & CInt(txt摘要.MaxLength / 2) & "个汉字或" & txt摘要.MaxLength & "个字符!", vbInformation + vbOKOnly, gstrSysName
                txt摘要.SetFocus
                Exit Function
            End If
        
            For intLop = 1 To .Rows - 1
                If Trim(.TextMatrix(intLop, mBillCol.C_材料)) <> "" Then
                    If Val(Trim(.TextMatrix(intLop, mBillCol.C_填写数量))) = 0 Then
                        MsgBox "第" & intLop & "行卫材的数量为空了，请检查！", vbInformation, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mBillCol.C_填写数量
                        Exit Function
                    End If
                    

                    If LenB(StrConv(Trim(Trim(.TextMatrix(intLop, mBillCol.c_批号))), vbFromUnicode)) > mintBatchNoLen Then
                        MsgBox "第" & intLop & "行卫材的批号超长,最多能输入" & Int(mintBatchNoLen / 2) & "个汉字或" & mintBatchNoLen & "个字符!", vbInformation + vbOKOnly, gstrSysName
                        .SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mBillCol.c_批号
                        Exit Function
                    End If
                    
                    '说明：只根据入库库房进行判断
                    '   1、入库库房且药库分批，则允许输入批次信息
                    '   2、入库药房且药房分批，则允许输入批次信息
                    bln库存分批 = (Val(mshBill.TextMatrix(intLop, mBillCol.C_库房分批)) = 1)
                    bln在用分批 = (Split(mshBill.TextMatrix(intLop, mBillCol.C_最大效期), "||")(2) = 1)
                    
                    If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
                        If mbln移库明确批次 = True Then '不按批次出库，不判断批号和效期是否填写
                            If ((bln入库库房 And bln库存分批) Or (Not bln入库库房 And bln在用分批)) Then
                                If Split(.TextMatrix(intLop, mBillCol.C_最大效期), "||")(0) <> 0 Then
'                            If .TextMatrix(intLop, mBillCol.C_库房分批) <> "0" And Split(.TextMatrix(intLop, mBillCol.C_最大效期), "||")(0) <> 0 Then
                                    If .TextMatrix(intLop, mBillCol.c_批号) = "" Or .TextMatrix(intLop, mBillCol.C_效期) = "" Then
                                        MsgBox "第" & intLop & "行的卫材是效期卫材,请把它的批号及失效期完整输入单据中！", vbInformation, gstrSysName
                                        mshBill.SetFocus
                                        .Row = intLop
                                        .MsfObj.TopRow = intLop
                                        If .TextMatrix(intLop, mBillCol.c_批号) = "" Then
                                            .Col = mBillCol.c_批号
                                        Else
                                            .Col = mBillCol.C_效期
                                        End If
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                            
                            '新增/修改时检查可用数量，防止并发
                        If Not CompareUsableQuantity(intLop, Val(Trim(.TextMatrix(intLop, mBillCol.C_填写数量))), True) Then
                            .SetFocus
                            .Row = intLop
                            .MsfObj.TopRow = intLop
                            .Col = mBillCol.C_填写数量
                            Exit Function
                        End If
                    End If
                    
                    If Val(.TextMatrix(intLop, mBillCol.C_填写数量)) > 9999999999# Then
                        MsgBox "第" & intLop & "行卫材的填写数量大于了数据库能够保存的" & vbCrLf & "最大范围9999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mBillCol.C_填写数量
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mBillCol.C_实际数量)) > 9999999999# Then
                        MsgBox "第" & intLop & "行卫材的实际数量大于了数据库能够保存的" & vbCrLf & "最大范围9999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mBillCol.C_实际数量
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mBillCol.C_采购金额)) > 9999999999999# Then
                        MsgBox "第" & intLop & "行卫材的成本金额大于了数据库能够保存的" & vbCrLf & "最大范围9999999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = IIf(.ColData(mBillCol.C_填写数量) = 4, mBillCol.C_填写数量, mBillCol.C_实际数量)
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mconintcol售价金额)) > 9999999999999# Then
                        MsgBox "第" & intLop & "行卫材的售价金额大于了数据库能够保存的" & vbCrLf & "最大范围9999999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = IIf(.ColData(mBillCol.C_填写数量) = 4, mBillCol.C_填写数量, mBillCol.C_实际数量)
                        Exit Function
                    End If
                    
                    If mbln分批卫材批号产地控制 = True Then
                        If ((bln入库库房 And bln库存分批) Or (Not bln入库库房 And bln在用分批)) And (.TextMatrix(intLop, mBillCol.c_批号) = "" Or .TextMatrix(intLop, mBillCol.C_产地) = "") And .TextMatrix(intLop, 0) <> "" Then
                            MsgBox "第" & intLop & "行，入库库房是分批管理，必须录入批号和产地！", vbInformation, gstrSysName
                            .SetFocus
                            .Row = intLop
                            .MsfObj.TopRow = intLop
                            If .TextMatrix(intLop, mBillCol.c_批号) = "" Then
                                .Col = mBillCol.c_批号
                            Else
                                .Col = mBillCol.C_产地
                            End If
                            Exit Function
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

'Private Function ReValidData() As Boolean
'    Dim intLop As Integer
'    Dim rsStock As New Recordset
'
'    With mshBill
'        If .TextMatrix(1, 0) <> "" Then         '先判有否数据
'            For intLop = 1 To .Rows - 1
'                If Trim(.TextMatrix(intLop, mBillCol.C_材料)) <> "" Then
'                    gstrSQL = "select * from 药品库存 where 药品id=" & .TextMatrix(intLop, 0)
'                End If
'            Next
'        Else
'            Exit Function
'        End If
'    End With
'End Function

Private Function SaveCard(Optional ByVal bln强制保存 As Boolean = False) As Boolean
    Dim chrNo As Variant
    Dim lng序号 As Long
    
    Dim lng库房ID As Long
    Dim lng入库房ID As Long
    Dim lng材料ID As Long
    Dim str批号 As String
    Dim lng批次 As Long
    Dim str产地 As String
    Dim str效期 As String
    Dim str填写数量 As Double
    Dim dbl采购价 As Double
    Dim dbl实际数量 As Double
    Dim dbl成本金额 As Double
    Dim dbl零售价 As Double
    Dim dbl零售金额 As Double
    Dim dbl差价 As Double
    Dim str摘要 As String
    Dim str填制人 As String
    Dim str填制日期 As String
    Dim str审核人 As String
    Dim datAssessDate As String
    Dim str灭菌效期 As String
    Dim n As Long
    
    Dim arrSQL As Variant
    Dim intRow As Integer
    
    arrSQL = Array()
    SaveCard = False
    
    '检查该单据是否在进入编辑界面后，被其他操作员修改
    If mint编辑状态 = 2 Or bln强制保存 Then          '修改
        mstrTime_End = GetBillInfo(19, mstr单据号)
        If mstrTime_End = "" Then
            MsgBox "该单据已经被其他操作员删除！", vbInformation, gstrSysName
            Exit Function
        End If
        If mstrTime_End > mstrTime_Start Then
            MsgBox "该单据已经被其他操作员编辑，请退出后重试！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    With mshBill
        chrNo = Trim(txtNO)
        lng库房ID = cboStock.ItemData(cboStock.ListIndex)
        
        If mint编辑状态 = 1 Or mint编辑状态 = 11 Then  ' Or mbln单据增加
            If chrNo <> "" Then
                If CheckNOExists(72, chrNo) Then Exit Function
            End If
        
            If chrNo = "" Then chrNo = sys.GetNextNo(72, lng库房ID)
            If IsNull(chrNo) Then Exit Function
        End If
        txtNO.Tag = chrNo
        
        lng入库房ID = cboEnterStock.ItemData(cboEnterStock.ListIndex)
        str摘要 = Trim(txt摘要.Text)
        If Txt填制人 <> "" Then
            str填制人 = Txt填制人
        Else
            str填制人 = UserInfo.用户名
        End If
        If Txt填制日期 <> "" Then
            str填制日期 = Txt填制日期.Caption
        Else
            str填制日期 = Format(sys.Currentdate, "yyyy-mm-dd HH:MM:SS")
        End If
        str审核人 = Txt审核人
        On Error GoTo ErrHandle
        
        If mint编辑状态 = 2 Or bln强制保存 Then        '修改
            If Not mbln申领单 Then
                gstrSQL = "zl_材料移库_Delete('" & mstr单据号 & "')"
            Else
                gstrSQL = "zl_材料申领_Delete('" & mstr单据号 & "')"
            End If
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "0;" & vbCrLf & gstrSQL
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
                lng批次 = Val(.TextMatrix(intRow, mBillCol.c_批次))
                str效期 = IIf(.TextMatrix(intRow, mBillCol.C_效期) = "", "", .TextMatrix(intRow, mBillCol.C_效期))
                str填写数量 = Round(Val(.TextMatrix(intRow, mBillCol.C_填写数量)) * Val(.TextMatrix(intRow, mBillCol.C_比例系数)), g_小数位数.obj_最大小数.数量小数)
                dbl实际数量 = Round(Val(.TextMatrix(intRow, mBillCol.C_实际数量)) * Val(.TextMatrix(intRow, mBillCol.C_比例系数)), g_小数位数.obj_最大小数.数量小数)
                
                If Val(.TextMatrix(intRow, mBillCol.c_原始数量)) <> 0 Then
                    
                    If Val(Format(Val(.TextMatrix(intRow, mBillCol.c_原始数量)) / Val(.TextMatrix(intRow, mBillCol.C_比例系数)), mFMT.FM_数量)) = Val(.TextMatrix(intRow, mBillCol.C_填写数量)) Then
                        If str填写数量 = dbl实际数量 Then
                            str填写数量 = Val(.TextMatrix(intRow, mBillCol.c_原始数量))
                            dbl实际数量 = str填写数量
                        ElseIf str填写数量 < dbl实际数量 Then
                            str填写数量 = Val(.TextMatrix(intRow, mBillCol.c_原始数量))
                        End If
                    End If
                End If
                                
                dbl采购价 = Round(Val(.TextMatrix(intRow, mBillCol.C_采购价)) / .TextMatrix(intRow, mBillCol.C_比例系数), g_小数位数.obj_最大小数.成本价小数)
                dbl成本金额 = Round(Val(.TextMatrix(intRow, mBillCol.C_采购金额)), g_小数位数.obj_最大小数.金额小数)
                dbl零售价 = Round(Val(.TextMatrix(intRow, mBillCol.C_售价)) / Val(.TextMatrix(intRow, mBillCol.C_比例系数)), g_小数位数.obj_最大小数.零售价小数)
                dbl零售金额 = Round(Val(.TextMatrix(intRow, mconintcol售价金额)), g_小数位数.obj_最大小数.金额小数)
                dbl差价 = Round(Val(.TextMatrix(intRow, mconintcol差价)), g_小数位数.obj_最大小数.金额小数)
                str灭菌效期 = IIf(.TextMatrix(intRow, mBillCol.C_灭菌失效期) = "", "", .TextMatrix(intRow, mBillCol.C_灭菌失效期))
                lng序号 = 2 * intRow - 1
                
                'zl_材料移库_INSERT( /*NO_IN*/, /*序号_IN*/, /*库房ID_IN*/,
                '/*对方部门ID_IN*/, /*材料ID_IN*/, /*批次_IN*/, /*填写数量_IN*/,实际数量/,
                '/*成本价_IN*/, /*成本金额_IN*/, /*零售价_IN*/, /*零售金额_IN*/,
                '/*差价_IN*/, /*填制人_IN*/, /*产地_IN*/, /*批号_IN*/, /*效期_IN*/,/灭菌效期_IN/
                '/*摘要_IN*/填制日期_IN );
                
                If Not mbln申领单 Or bln强制保存 Then
                    gstrSQL = "zl_材料移库_INSERT('" & chrNo & "'," & lng序号 & "," & lng库房ID & "," & _
                         lng入库房ID & "," & lng材料ID & "," & lng批次 & "," & str填写数量 & "," & dbl实际数量 & "," & _
                         dbl采购价 & "," & dbl成本金额 & "," & dbl零售价 & "," & dbl零售金额 & "," & _
                         dbl差价 & ",'" & str填制人 & "','" & str产地 & "','" & _
                         str批号 & "'," & _
                        IIf(str效期 = "", "Null", "to_date('" & Format(str效期, "yyyy-MM-dd") & "','yyyy-mm-dd')") & "," & _
                        IIf(str灭菌效期 = "", "Null", "to_date('" & Format(str灭菌效期, "yyyy-MM-dd") & "','yyyy-mm-dd')") & ",'" & _
                        str摘要 & "',to_date('" & str填制日期 & "','yyyy-mm-dd HH24:MI:SS')," & _
                        IIf(mstr核查人 = "", "null", "'" & mstr核查人 & "'") & "," & _
                        IIf(mstr核查日期 = "", "Null", "to_date('" & Format(mstr核查日期, "yyyy-MM-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS')") & ")"
                Else
                    gstrSQL = "zl_材料申领_INSERT('" & _
                        chrNo & "'," & _
                        lng序号 & "," & _
                        lng库房ID & "," & _
                        lng入库房ID & "," & _
                        lng材料ID & "," & _
                        lng批次 & "," & _
                        str填写数量 & "," & _
                        dbl实际数量 & "," & _
                        dbl采购价 & "," & _
                        dbl成本金额 & "," & _
                        dbl零售价 & "," & _
                        dbl零售金额 & "," & _
                        dbl差价 & ",'" & _
                        str填制人 & "','" & _
                        str产地 & "','" & _
                        str批号 & "'," & _
                        IIf(str效期 = "", "Null", "to_date('" & Format(str效期, "yyyy-MM-dd") & "','yyyy-mm-dd')") & "," & _
                        IIf(str灭菌效期 = "", "Null", "to_date('" & Format(str灭菌效期, "yyyy-MM-dd") & "','yyyy-mm-dd')") & ",'" & _
                        str摘要 & "',to_date('" & str填制日期 & "','yyyy-mm-dd HH24:MI:SS')" & "," & _
                        IIf(mstr核查人 = "", "null,", "'" & mstr核查人 & "',") & _
                        IIf(mstr核查日期 = "", "Null", "to_date('" & Format(mstr核查日期, "yyyy-MM-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS')") & ")"
                End If
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = CStr(lng材料ID) & ";" & vbCrLf & gstrSQL
            End If
            recSort.MoveNext
        Next
        
        If Not ExecuteSql(arrSQL, mstrCaption, False, Not bln强制保存) Then Exit Function
'        If Not 检查单价(19, txtNO.Tag) Then
'            If Not bln强制保存 Then gcnOracle.RollbackTrans
'            Exit Function
'        End If
        If Not bln强制保存 Then gcnOracle.CommitTrans
        
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveCard = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveStrike() As Boolean
    Dim int行次 As Integer
    Dim int原记录状态 As Integer
    Dim strNo As String
    Dim str序号 As Integer
    Dim lng材料ID As Long
    Dim dbl冲销数量 As Double
    Dim str填制人 As String
    Dim str填制日期  As String
    Dim intRow As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim int库存检查 As Integer, lng库房ID As Long, lng批次 As Long
    Dim n As Long
    
    SaveStrike = False
    
    With mshBill
        strNo = Trim(txtNO.Tag)
        lng库房ID = cboEnterStock.ItemData(cboEnterStock.ListIndex)
        int库存检查 = Get出库检查(lng库房ID)
    
        '检查冲销数量，不能小于零
        For intRow = 1 To .Rows - 1
            If Val(.TextMatrix(intRow, mBillCol.C_实际数量)) <> 0 Then
                If Not 相同符号(Val(.TextMatrix(intRow, mBillCol.C_填写数量)), Val(.TextMatrix(intRow, mBillCol.C_实际数量))) Then
                    MsgBox "请输入合法的冲销数量（第" & intRow & "行）！", vbInformation, gstrSysName
                    Exit Function
                End If
                If int库存检查 <> 0 Then
                    dbl冲销数量 = Round(Val(.TextMatrix(intRow, mBillCol.C_实际数量)) * .TextMatrix(intRow, mBillCol.C_比例系数), g_小数位数.obj_散装小数.数量小数)
                    If Val(.TextMatrix(intRow, mBillCol.C_实际数量)) = Val(.TextMatrix(intRow, mBillCol.C_填写数量)) Then
                        dbl冲销数量 = Val(.TextMatrix(intRow, mBillCol.c_原始数量))
                    End If
                    lng批次 = 取单据批次(19, strNo, Val(.TextMatrix(intRow, 0)), Val(.TextMatrix(intRow, mBillCol.c_序号)) + 1)
                    If Check可用数量(lng库房ID, Val(.TextMatrix(intRow, 0)), lng批次, dbl冲销数量, int库存检查, IIf(mint冲销方式 = 2, 1, 0)) = False Then Exit Function
                End If
                
            End If
        Next
        
        str填制人 = UserInfo.用户名
        str填制日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        int原记录状态 = mint记录状态
        
        On Error GoTo ErrHandle
        gcnOracle.BeginTrans
        
        int行次 = 0
        '按药品ID顺序更新数据
        recSort.Sort = "药品id,批次,序号"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!行号
'        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" And Val(.TextMatrix(intRow, mBillCol.C_实际数量)) <> 0 Then
                int行次 = int行次 + 1
                
                lng材料ID = .TextMatrix(intRow, 0)
                dbl冲销数量 = Round(Val(.TextMatrix(intRow, mBillCol.C_实际数量)) * .TextMatrix(intRow, mBillCol.C_比例系数), g_小数位数.obj_散装小数.数量小数)
  
                If Val(.TextMatrix(intRow, mBillCol.C_实际数量)) = Val(.TextMatrix(intRow, mBillCol.C_填写数量)) Then
                    dbl冲销数量 = Val(.TextMatrix(intRow, mBillCol.c_原始数量))
                End If
                dbl冲销数量 = IIf(mint编辑状态 = 6 And mint冲销方式 = 2, -1, 1) * dbl冲销数量
                           
                str序号 = .TextMatrix(intRow, mBillCol.c_序号)
                
                'ZL_材料移库_STRIKE(/*int行次*/,/*int原记录状态*/,/*strNO*/,/*str序号*/, /*lng材料ID*/,
                '/*dbl冲销数量*/,/*str填制人*/, /*str填制日期*/);
                gstrSQL = "ZL_材料移库_STRIKE(" & int行次 & "," & int原记录状态 & ",'" & strNo & "'," & str序号 & "," & lng材料ID & "," & dbl冲销数量 & ",'" _
                    & str填制人 & "',to_date('" & Format(str填制日期, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS') ," & mint冲销方式 & ")"
                zlDatabase.ExecuteProcedure gstrSQL, mstrCaption
            End If
            recSort.MoveNext
        Next
        gcnOracle.CommitTrans
        
        If int行次 = 0 Then
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
    If ErrCenter() = 1 Then Resume
End Function

Private Sub 显示合计金额()
    Dim curTotal As Double, Cur记帐金额 As Double, Cur记帐差价 As Double
    Dim intLop As Integer
    
    curTotal = 0: Cur记帐金额 = 0: Cur记帐差价 = 0:
    
    With mshBill
        For intLop = 1 To .Rows - 1
            curTotal = curTotal + Val(.TextMatrix(intLop, mBillCol.C_采购金额))
            Cur记帐金额 = Cur记帐金额 + Val(.TextMatrix(intLop, mconintcol售价金额))
        Next
    End With
    
    Cur记帐差价 = Cur记帐金额 - curTotal
    lblPurchasePrice.Caption = "成本金额合计：" & Format(curTotal, mFMT.FM_金额)
    lblSalePrice.Caption = "售价金额合计：" & Format(Cur记帐金额, mFMT.FM_金额)
    lblDifference.Caption = "差价合计：" & Format(Cur记帐差价, mFMT.FM_金额)
End Sub

Private Sub 提示库存数()
    Dim rsUseCount As New Recordset
    Dim strNote As String
    
    On Error GoTo ErrHandle
    With mshBill
        If .TextMatrix(.Row, mBillCol.C_材料) = "" Then
            stbThis.Panels(2).Text = ""
            Exit Sub
        End If
        If .TextMatrix(mshBill.Row, 0) = "" Then Exit Sub
        
        If mint编辑状态 <> 10 Then
            If mbln申领单 Then '申领单
                If mbln明确批次 Then
                    gstrSQL = "" & _
                        "   Select 可用数量/" & .TextMatrix(.Row, mBillCol.C_比例系数) & " as  可用数量 " & _
                        "   From 药品库存 " & _
                        "   Where 库房id=[1]" & _
                        "          and 药品id=[2]" & _
                        "           and 性质=1 and " & _
                        "          nvl(批次,0)=[3]"
                Else
                    gstrSQL = "" & _
                        "   Select Sum(可用数量)/" & .TextMatrix(.Row, mBillCol.C_比例系数) & " as  可用数量 " & _
                        "   From 药品库存 " & _
                        "   Where 库房id=[1]" & _
                        "          and 药品id=[2]" & _
                        "           and 性质=1  "
                End If
            Else '移库单
                If mbln移库明确批次 Then
                    gstrSQL = "" & _
                        "   Select 可用数量/" & .TextMatrix(.Row, mBillCol.C_比例系数) & " as  可用数量 " & _
                        "   From 药品库存 " & _
                        "   Where 库房id=[1]" & _
                        "          and 药品id=[2]" & _
                        "           and 性质=1 and " & _
                        "          nvl(批次,0)=[3]"
                Else
                    gstrSQL = "" & _
                        "   Select Sum(可用数量)/" & .TextMatrix(.Row, mBillCol.C_比例系数) & " as  可用数量 " & _
                        "   From 药品库存 " & _
                        "   Where 库房id=[1]" & _
                        "          and 药品id=[2]" & _
                        "           and 性质=1  "
                End If
                    
            End If
        
                
                Set rsUseCount = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mBillCol.c_批次)))
                
                If rsUseCount.EOF Then
                    .TextMatrix(.Row, mBillCol.C_可用数量) = 0
                Else
                    .TextMatrix(.Row, mBillCol.C_可用数量) = IIf(IsNull(rsUseCount.Fields(0)), 0, rsUseCount.Fields(0))
                End If
                rsUseCount.Close
                
                stbThis.Panels(2).Text = "该卫材当前库存数为[" & Format(.TextMatrix(.Row, mBillCol.C_可用数量), mFMT.FM_数量) & "]" & .TextMatrix(.Row, mBillCol.c_单位)
        Else
            '仅在发送时，显示该药品在所有库房的库存，以便于库房人员决定实际的发送数量
            gstrSQL = "" & _
            "   Select B.名称 AS 库房,Nvl(A.可用数量,0)/" & .TextMatrix(.Row, mBillCol.C_比例系数) & " as 可用数量 " & _
            "   From 药品库存 A,部门表 B" & _
            "    Where A.库房ID=B.ID And A.药品id=[1]" & _
            "           And A.性质=1 "
            
            Set rsUseCount = zlDatabase.OpenSQLRecord(gstrSQL, "提示库存数", Val(.TextMatrix(.Row, 0)))
            With rsUseCount
                Do While Not .EOF
                    strNote = strNote & "," & !库房 & ":" & Format(zlStr.NVL(!可用数量, 0), mFMT.FM_数量) & mshBill.TextMatrix(mshBill.Row, mBillCol.c_单位)
                    .MoveNext
                Loop
            End With
            stbThis.Panels(2).Text = Mid(strNote, 2)
        End If
    End With
    Exit Sub
ErrHandle:
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
                "      a.审核日期 Is Not Null And a.No = [1] And a.库房id + 0 = [2]" & vbNewLine & _
                GetPriceClassString("E") & "Order By a.序号"

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
                MsgBox "材料[" & !药品名称 & "]未在" & cboStock.Text & "中设置存储属性，将不能移库！"
                blnInput = True
            End If
            rsTemp.Filter = ""
            rsTemp.Filter = " 收费细目id=" & lng药品ID & " and 执行科室id=" & cboEnterStock.ItemData(cboEnterStock.ListIndex)
            If rsTemp.RecordCount = 0 Then
                MsgBox "材料[" & !药品名称 & "]未在" & cboEnterStock.Text & "中设置存储属性，将不能移库！"
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
                        MsgBox !药品名称 & "库存不足，将不能移库！", vbInformation, gstrSysName
                        blnInput = True
                    End Select
                End If
            End If
            
            '装入数据(SetColValue)
            If blnInput = False Then
                int包装系数 = IIf(mintUnit = 0, 1, !换算系数)
                If Not SetColValue(intRow, !材料ID, "[" & !编码 & "]" & !通用名, _
                   NVL(!规格), NVL(!产地), IIf(mintUnit = 0, !零售单位, !包装单位), _
                    NVL(!现价, 0), NVL(!批号), NVL(!效期), IIf(IsNull(!灭菌效期), "", Format(!灭菌效期, "yyyy-MM-dd")), _
                    NVL(!最大效期, 0), !库房分批, NVL(!可用数量, 0), NVL(!实际金额, 0), NVL(!实际差价, 0), _
                    IIf(IsNull(!指导差价率), "0", !指导差价率), int包装系数, NVL(!批次, 0), !时价, _
                    !在用分批, IIf(IsNull(!批准文号), "", !批准文号)) Then
                    mshBill.ClearBill
                    Exit Sub
                End If

                '填写数量、采购价、售价等列
                mshBill.TextMatrix(intRow, mBillCol.C_行号) = intRow
                mshBill.TextMatrix(intRow, mBillCol.C_填写数量) = Format(!实际数量 / int包装系数, mFMT.FM_数量)
                mshBill.TextMatrix(intRow, mBillCol.C_实际数量) = Format(!实际数量 / int包装系数, mFMT.FM_数量)
                mshBill.TextMatrix(intRow, mBillCol.C_采购价) = Format(!平均成本价 * int包装系数, mFMT.FM_成本价)
                mshBill.TextMatrix(intRow, mBillCol.C_采购金额) = Format(Val(mshBill.TextMatrix(intRow, mBillCol.C_采购价)) * Val(mshBill.TextMatrix(intRow, mBillCol.C_实际数量)), mFMT.FM_金额)
                mshBill.TextMatrix(intRow, mconintcol售价金额) = Format(Val(mshBill.TextMatrix(intRow, mBillCol.C_售价)) * Val(mshBill.TextMatrix(intRow, mBillCol.C_实际数量)), mFMT.FM_金额)
                mshBill.TextMatrix(intRow, mconintcol差价) = Format(Val(mshBill.TextMatrix(intRow, mconintcol售价金额)) - mshBill.TextMatrix(intRow, mBillCol.C_采购金额), mFMT.FM_金额)

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
End Sub

Private Sub txt摘要_GotFocus()
    
    OS.OpenIme (True)
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
    OS.OpenIme False
End Sub

'与可用数量进行比较
Private Function CompareUsableQuantity(ByVal intRow As Integer, ByVal dbl填写数量 As Double, Optional ByVal blnSave As Boolean = False) As Boolean
    Dim dblUsableQuantity As Double      '实际数量对应的组成数量
    Dim numUsedCount As Double
    Dim vardrug As Variant
    Dim dbltotal As Double              '某种卫材输入的所有数量
    Dim intLop As Integer
    Dim rsCheck As ADODB.Recordset
    Dim strSaveCheck As String
    
    'mint库存检查: 0-不检查;1-检查，不足提醒；2-检查，不足禁止
    
    CompareUsableQuantity = False
    If Not mbln移库明确批次 Then CompareUsableQuantity = True: Exit Function
    
    With mshBill
        If .TextMatrix(intRow, 0) = "" Then Exit Function
        
        If Not blnSave Then
            dblUsableQuantity = Format(.TextMatrix(intRow, mBillCol.C_可用数量), mFMT.FM_数量)
        Else
            '新增，修改保存时重新取数据库中的可用数量，主要防止并发或多人同时填单对可用数量取值的影响
            gstrSQL = "Select Nvl(可用数量, 0) 可用数量 From 药品库存 Where 性质 = 1 And 库房id = [1] And 药品id = [2] And Nvl(批次, 0) = [3] "
            Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "CompareUsableQuantity", Val(cboStock.ItemData(cboStock.ListIndex)), Val(.TextMatrix(intRow, 0)), Val(.TextMatrix(intRow, mBillCol.c_批次)))
            
            If rsCheck.EOF Then
                dblUsableQuantity = 0
            Else
                dblUsableQuantity = Val(Format(rsCheck!可用数量 / Val(.TextMatrix(intRow, mBillCol.C_比例系数)), mFMT.FM_数量))
                
                If dblUsableQuantity <> Val(Format(.TextMatrix(intRow, mBillCol.C_可用数量), mFMT.FM_数量)) Then
                    .TextMatrix(intRow, mBillCol.C_可用数量) = dblUsableQuantity
                End If
                
                strSaveCheck = "，可能是其他操作员填单占用了"
            End If
        End If
        
        If mint库存检查 = 0 Then
            '0-不检查
        ElseIf mint库存检查 = 1 Then
            '1-检查，不足提醒
            If mint编辑状态 = 1 Then
                If dbl填写数量 > dblUsableQuantity Then
                    If MsgBox("第" & intRow & "行[" & .TextMatrix(intRow, C_材料) & "]的数量“" & dbl填写数量 & "”大于了该卫材的可用库存数量“" & dblUsableQuantity & "”" & strSaveCheck & "，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                End If
            ElseIf mint编辑状态 = 2 Or mint编辑状态 = 3 Then
                numUsedCount = 0
                For Each vardrug In mcolUsedCount
                    If vardrug(0) = .TextMatrix(intRow, 0) & .TextMatrix(intRow, mBillCol.c_批次) Then
                        numUsedCount = vardrug(1)
                        Exit For
                    End If
                Next
                
                If gSystem_Para.para_卫材填单下可用库存 = False Then
                    '如果没有预减可用数量，则不算界面的原始数量
                    numUsedCount = 0
                End If
                
                If dbl填写数量 > dblUsableQuantity + numUsedCount Then
                    If MsgBox("第" & intRow & "行[" & .TextMatrix(intRow, C_材料) & "]的数量“" & dbl填写数量 & "”大于了该卫材的可用库存数量“" & dblUsableQuantity + numUsedCount & "”" & strSaveCheck & "，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                End If
            End If
            
        ElseIf mint库存检查 = 2 Then
            '2-检查，不足禁止
            If mint编辑状态 = 1 Then
                dbltotal = 0
                For intLop = 1 To .Rows - 1
                    If .TextMatrix(intLop, 0) <> "" Then
                        If intLop <> intRow And .TextMatrix(intLop, 0) = .TextMatrix(intRow, 0) And IIf(.TextMatrix(intLop, mBillCol.c_批次) = "", "0", .TextMatrix(intLop, mBillCol.c_批次)) = "0" Then
                            dbltotal = dbltotal + .TextMatrix(intLop, mBillCol.C_填写数量)
                        End If
                    End If
                Next
                
                
                If dbl填写数量 + dbltotal > dblUsableQuantity Then
                    MsgBox "第" & intRow & "行[" & .TextMatrix(intRow, C_材料) & "]的数量“" & dbl填写数量 & "”大于了该卫材的可用库存数量“" & dblUsableQuantity - dbltotal & "”" & strSaveCheck & "，请重输！", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            ElseIf mint编辑状态 = 2 Or mint编辑状态 = 3 Then
                numUsedCount = 0
                For Each vardrug In mcolUsedCount
                    If vardrug(0) = .TextMatrix(intRow, 0) & .TextMatrix(intRow, mBillCol.c_批次) Then
                        numUsedCount = vardrug(1)
                        Exit For
                    End If
                Next
                
                dbltotal = 0
                For intLop = 1 To .Rows - 1
                    If .TextMatrix(intLop, 0) <> "" Then
                        If intLop <> intRow And .TextMatrix(intLop, 0) = .TextMatrix(intRow, 0) And IIf(.TextMatrix(intLop, mBillCol.c_批次) = "", "0", .TextMatrix(intLop, mBillCol.c_批次)) = "0" Then
                            dbltotal = dbltotal + Val(.TextMatrix(intLop, mBillCol.C_实际数量))
                        End If
                    End If
                Next
                
                If gSystem_Para.para_卫材填单下可用库存 = False Then
                    '如果没有预减可用数量，则不算界面的原始数量
                    numUsedCount = 0
                End If
                
                If dbl填写数量 + dbltotal > dblUsableQuantity + numUsedCount Then
                    MsgBox "第" & intRow & "行[" & .TextMatrix(intRow, C_材料) & "]的数量“" & dbl填写数量 & "”大于了该卫材的可用库存数量“" & dblUsableQuantity + numUsedCount - dbltotal & "”" & strSaveCheck & "，请重输！", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            End If
        End If
            
    End With
    
    CompareUsableQuantity = True
    
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

'打印单据
Private Sub printbill()
    Dim strNo As String
    strNo = txtNO.Tag
    FrmBillPrint.ShowMe Me, glngSys, "zl1_bill_1716", mint记录状态, mintUnit, 1716, "卫材调拨单", strNo
End Sub

'取数据库中批号的长度，这样，程序中的批号长度与数据库中保持一致了
Private Function GetBatchNoLen() As Integer
    Dim rsBatchNolen As New Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "select 批号 from 药品收发记录 where rownum<1 "
    
    zlDatabase.OpenRecordset rsBatchNolen, gstrSQL, "取字段长度"
    GetBatchNoLen = rsBatchNolen.Fields(0).DefinedSize
    rsBatchNolen.Close
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

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
        dbl比例系数 = Val(mshBill.TextMatrix(lngRow, mBillCol.C_比例系数))
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
        
        '如果该卫材是分批卫材，但批次为零，则说明需要自动分解
        blnAddRow = False
        If bln分批 And lng批次 = 0 Then
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
                    mshBill.TextMatrix(lngRow, mBillCol.c_序号) = (lngRow - 1) * 2 + 1
                    mshBill.TextMatrix(lngRow, mBillCol.c_批次) = rsCheck!批次
                    mshBill.TextMatrix(lngRow, mBillCol.c_批号) = IIf(IsNull(rsCheck!批号), "", rsCheck!批号)
                    mshBill.TextMatrix(lngRow, mBillCol.C_效期) = IIf(IsNull(rsCheck!效期), "", rsCheck!效期)
                    mshBill.TextMatrix(lngRow, mBillCol.C_产地) = IIf(IsNull(rsCheck!产地), "", rsCheck!产地)
                    mshBill.TextMatrix(lngRow, mBillCol.C_批准文号) = IIf(IsNull(rsCheck!批准文号), "", rsCheck!批准文号)
                    
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
                        mshBill.TextMatrix(lngRow, mBillCol.c_原始数量) = Val(mshBill.TextMatrix(lngRow, mBillCol.C_实际数量)) * Val(mshBill.TextMatrix(lngRow, mBillCol.C_比例系数))
                    End If
                    
                    mshBill.TextMatrix(lngRow, mBillCol.C_填写数量) = Format(dbl数量, mFMT.FM_数量)
                    mshBill.TextMatrix(lngRow, mBillCol.C_实际数量) = Format(dbl数量, mFMT.FM_数量)
                    
                    If Trim(mshBill.TextMatrix(lngRow, mBillCol.C_实际数量)) = "" Then mshBill.TextMatrix(lngRow, mBillCol.C_实际数量) = 0
                    mshBill.TextMatrix(lngRow, mBillCol.C_实际差价) = Format(rsCheck!实际差价, mFMT.FM_金额)
                    mshBill.TextMatrix(lngRow, mBillCol.C_实际金额) = Format(rsCheck!实际金额, mFMT.FM_金额)
                    mshBill.TextMatrix(lngRow, mBillCol.C_可用数量) = Format(rsCheck!可用数量, mFMT.FM_金额)
                    mshBill.TextMatrix(lngRow, mBillCol.C_售价) = Format(IIf(bln时价, dbl现价_时价, dbl现价), mFMT.FM_零售价)
                    mshBill.TextMatrix(lngRow, mconintcol售价金额) = Format(Val(mshBill.TextMatrix(lngRow, mBillCol.C_售价)) * dbl数量, mFMT.FM_金额)
                    
'                    If rsCheck!实际金额 > 0 Then
'                        mshBill.TextMatrix(lngRow, mBillCol.C_差价) = Format(Val(mshBill.TextMatrix(lngRow, mBillCol.C_售价金额)) * rsCheck!实际差价 / rsCheck!实际金额, mFMT.FM_金额)
'                    Else
'                        mshBill.TextMatrix(lngRow, mBillCol.C_差价) = Format(Val(mshBill.TextMatrix(lngRow, mBillCol.C_售价金额)) * Val(mshBill.TextMatrix(lngRow, mBillCol.C_指导差价率)) / 100, mFMT.FM_金额)
'                    End If
'                    mshBill.TextMatrix(lngRow, mBillCol.C_采购金额) = Format(Val(mshBill.TextMatrix(lngRow, mBillCol.C_售价金额)) - Val(mshBill.TextMatrix(lngRow, mBillCol.C_差价)), mFMT.FM_金额)
                    
'                    If dbl数量 <> 0 Then
'                        mshBill.TextMatrix(lngRow, mBillCol.C_采购价) = Format(Val(mshBill.TextMatrix(lngRow, mBillCol.C_采购金额)) / dbl数量, mFMT.FM_成本价)
'                    Else
'                        mshBill.TextMatrix(lngRow, mBillCol.C_采购价) = Format(dbl成本价, mFMT.FM_成本价)
'                    End If
                    '采用新的方式计算成本价 成本价=药品库存.平均成本价
                    mshBill.TextMatrix(lngRow, mBillCol.C_采购价) = Format(Get成本价(lng材料ID, lng库房ID, Val(mshBill.TextMatrix(lngRow, mBillCol.c_批次))) * dbl比例系数, mFMT.FM_成本价)
                    mshBill.TextMatrix(lngRow, mBillCol.C_采购金额) = Format(Val(mshBill.TextMatrix(lngRow, mBillCol.C_采购价)) * dbl数量, mFMT.FM_金额)
                    mshBill.TextMatrix(lngRow, mconintcol差价) = Format(Val(mshBill.TextMatrix(lngRow, mconintcol售价金额)) - Val(mshBill.TextMatrix(lngRow, mBillCol.C_采购金额)), mFMT.FM_金额)
                    
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
                mshBill.TextMatrix(lngRow, mBillCol.c_序号) = (lngRow - 1) * 2 + 1
                mshBill.TextMatrix(lngRow, mBillCol.C_实际数量) = ""
                mshBill.TextMatrix(lngRow, mconintcol售价金额) = ""
                mshBill.TextMatrix(lngRow, mBillCol.C_采购金额) = ""
                mshBill.TextMatrix(lngRow, mconintcol差价) = ""
            End If
        Else
            mshBill.TextMatrix(lngRow, mBillCol.C_行号) = lngRow
            mshBill.TextMatrix(lngRow, mBillCol.c_序号) = (lngRow - 1) * 2 + 1
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

Private Sub ShowColor(Optional ByVal lngCurRow As Long = 0)
    '在查阅或审核时，将库存不足的记录以暗红色显示出来
    Dim lngSelect_Row  As Long, lngSelect_Col As Long
    Dim lng库房ID As Long, lng材料ID As Long, lng材料ID_Last As Long, lng批次 As Long
    Dim bln库房 As Boolean, bln分批 As Boolean, bln时价 As Boolean
    Dim dbl填写数量 As Double, dbl数量 As Double, dbl比例系数 As Double
    Dim dbl现价 As Currency, dbl现价_时价 As Currency
    Dim rsTemp As New ADODB.Recordset
    Dim rsCheck As New ADODB.Recordset
    
    Dim lngRow As Long
    On Error GoTo ErrHand
    
    mshBill.Redraw = False
    lngSelect_Row = mshBill.Row: lngSelect_Col = mshBill.Col
    lngRow = IIf(lngCurRow > 0, lngCurRow, 1)
    lng库房ID = Val(cboStock.ItemData(cboStock.ListIndex))
    bln库房 = CheckStockProperty(lng库房ID)
    
    Do While True
        If lngRow > mshBill.Rows - 1 Then Exit Do
        mshBill.Row = lngRow: mshBill.Col = mBillCol.C_材料
        mshBill.MsfObj.CellForeColor = &H0&
    
        lng材料ID = Val(mshBill.TextMatrix(lngRow, 0))
        dbl填写数量 = Val(mshBill.TextMatrix(lngRow, mBillCol.C_填写数量))
        dbl比例系数 = Val(mshBill.TextMatrix(lngRow, mBillCol.C_比例系数))
        lng批次 = Val(mshBill.TextMatrix(lngRow, mBillCol.c_批次))
        If lng材料ID = 0 Then Exit Do
        
        '提取该材料对于出库库房是否分批、时价的属性
        If lng材料ID <> lng材料ID_Last Then
            lng材料ID_Last = lng材料ID
            gstrSQL = "" & _
                "   Select Nvl(A.库房分批,0) 库房分批,Nvl(A.在用分批,0) 在用分批,Nvl(B.是否变价,0) 时价,Nvl(P.现价,0) 现价 " & _
                "   From 材料特性 A,收费项目目录 B,收费价目 P" & _
                "   Where A.材料ID = B.ID And B.ID=P.收费细目ID And A.材料ID = [1]" & _
                "           And Sysdate between P.执行日期 And Nvl(P.终止日期,Sysdate)" & _
                GetPriceClassString("P")
                
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取该材料对于出库库房是否分批、时价的属性", lng材料ID)
            
            dbl现价 = rsTemp!现价
            bln时价 = (rsTemp!时价 = 1)
            bln分批 = IIf(bln库房, (rsTemp!库房分批 = 1), (rsTemp!在用分批 = 1))
        End If
        
        '根据卫材申领数量与库存可用数量，给单元格上色
        If bln分批 And lng批次 <> 0 Then
            '如果该卫材是分批卫材，且指定批次
            gstrSQL = "" & _
                "   Select Nvl(可用数量,0)/" & dbl比例系数 & " As 可用数量,Nvl(实际数量,0)/" & dbl比例系数 & " As 实际数量," & _
                "           Nvl(实际金额,0) 实际金额,Nvl(实际差价,0) 实际差价" & _
                "   From 药品库存 " & _
                "   Where 库房ID=[1] And 药品ID=[2] And 性质=1 And Nvl(批次,0)=[3]"
        Else
            '未指定批次或不分批的卫材，直接将出库库房该卫材所有库存记录累加
            gstrSQL = "" & _
                "   Select 药品id 材料ID,Sum(Nvl(可用数量,0))/" & dbl比例系数 & " As 可用数量,Sum(Nvl(实际数量,0))/" & dbl比例系数 & " As 实际数量," & _
                "           Sum(Nvl(实际金额,0)) 实际金额,Sum(Nvl(实际差价,0)) 实际差价" & _
                "   From 药品库存 Where 库房ID=[1] And 药品ID=[2] And 性质=1 " & _
                "   Group by 药品ID"
        End If
        
        Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "提取该卫材在指定库存的所有库存记录", lng库房ID, lng材料ID, lng批次)
        
        If rsCheck.EOF Then
            mshBill.MsfObj.CellForeColor = &H400040
        Else
            If rsCheck!可用数量 < dbl填写数量 Then
                mshBill.MsfObj.CellForeColor = &H400040
            End If
        End If
        If lngCurRow > 0 Then Exit Do
        lngRow = lngRow + 1
    Loop
    
    mshBill.Row = lngSelect_Row: mshBill.Col = lngSelect_Col
    mshBill.Redraw = True
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    mshBill.Redraw = True
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
            dbl比例系数 = Val(mshBill.TextMatrix(lngRow, mBillCol.C_比例系数))
            dbl填写数量 = Val(mshBill.TextMatrix(lngRow, mBillCol.C_实际数量))
            
            dbl可用数量 = 0
            '查找该材料的库存记录
            rsCheck.Filter = "材料ID=" & lng材料ID & " And 批次=" & lng批次
            If rsCheck.RecordCount <> 0 Then
                If mint编辑状态 = 10 Then '发送时应该用可用数量判断
                    dbl可用数量 = zlStr.NVL(rsCheck!可用数量, 0) / dbl比例系数
                ElseIf mint编辑状态 = 3 Then  '审核时应该用实际数量判断
                    dbl可用数量 = zlStr.NVL(rsCheck!实际数量, 0) / dbl比例系数
                End If
            End If
            
            '如果库存的可用数量不够
            If Not (dbl可用数量 >= dbl填写数量) Then
                int库存检查 = mint库存检查
                '如果该材料是时价或分批，库存不足不允许出库，相当于禁止出库
                rsProperty.Filter = "材料ID=" & lng材料ID
                bln特殊 = (IIf(bln库房, (rsProperty!库房分批 = 1), (rsProperty!在用分批 = 1)) Or (rsProperty!是否变价 = 1))
                strMsg = ""
                If bln特殊 Then
                    int库存检查 = 2
                    '如果是批次材料，但批次小于等于零，说明未执行分解功能
                    If lng批次 <= 0 And IIf(bln库房, (rsProperty!库房分批 = 1), (rsProperty!在用分批 = 1)) Then
                        strMsg = "（请先执行分解功能明确批次材料的出库批次）"
                    End If
                End If
                
                If bln下库存 = True And (mint编辑状态 = 10 Or (mint编辑状态 = 3 And mint移库处理流程 = 0)) Then
                Else
                    '按正常流程进行提示或禁止
                    Select Case int库存检查
                    Case 1  '仅提示
                        If MsgBox(rsProperty!通用名 & "的可用库存不足，是否继续？" & strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Exit Function
                        End If
                    Case 2
                        MsgBox rsProperty!通用名 & "的可用库存不足！" & strMsg, vbInformation, gstrSysName
                        Exit Function
                    End Select
                End If
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
Private Function CheckSend() As Boolean
    Dim rsTemp As New ADODB.Recordset
    '检查当前单据是否已发送
    On Error GoTo ErrHand
    
    gstrSQL = "Select 配药日期 From 药品收发记录 " & _
              "Where 单据=19 And NO=[1] And Rownum<2"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查当前单据是否已发送", Me.txtNO.Tag)
              
    If (zlStr.NVL(rsTemp!配药日期) = "") Then
        MsgBox "该单据已被其他操作员取消发送，不允许接收！", vbInformation, gstrSysName
        Exit Function
    End If
    CheckSend = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function


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
                !序号 = IIf(Val(mshBill.TextMatrix(n, mBillCol.c_序号)) = 0, n, Val(mshBill.TextMatrix(n, mBillCol.c_序号)))
                !药品id = Val(mshBill.TextMatrix(n, 0))
                !批次 = Val(mshBill.TextMatrix(n, mBillCol.c_批次))
                
                .Update
            End If
        Next
        
    End With
End Sub
