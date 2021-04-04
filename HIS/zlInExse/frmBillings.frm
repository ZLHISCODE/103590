VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBillings 
   AutoRedraw      =   -1  'True
   Caption         =   "住院记帐表"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11760
   Icon            =   "frmBillings.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   11760
   Begin VB.Timer tmrStatuPati 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picStatuPancl 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   6690
      ScaleHeight     =   300
      ScaleWidth      =   2340
      TabIndex        =   34
      Top             =   7065
      Width           =   2340
      Begin VB.Label lblStatuPati 
         Caption         =   "病人欠费"
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
         Height          =   345
         Left            =   0
         TabIndex        =   35
         Top             =   45
         Width           =   855
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   6750
      Top             =   540
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillings.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillings.frx":11A4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   23
      Top             =   7005
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmBillings.frx":1A7E
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13705
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   88
            Key             =   "病人余额"
            Object.ToolTipText     =   "病人余额"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   71
            Key             =   "MedicareType"
            Object.ToolTipText     =   "医保大类"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmBillings.frx":2312
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmBillings.frx":294C
            Key             =   "WB"
            Object.ToolTipText     =   "五笔(F7)"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
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
   Begin MSComctlLib.ListView lvwPati 
      Height          =   2550
      Left            =   60
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1875
      Visible         =   0   'False
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   4498
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "病人ID"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "住院号"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "姓名"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "床号"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "性别"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "年龄"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "在院"
         Object.Width           =   970
      EndProperty
   End
   Begin ZL9BillEdit.BillEdit Bill 
      Height          =   4290
      Left            =   15
      TabIndex        =   2
      Top             =   1080
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   7567
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      TxtCheck        =   -1  'True
      TxtCheck        =   -1  'True
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Active          =   -1  'True
      Cols            =   2
      RowHeight0      =   315
      RowHeightMin    =   315
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
   Begin VB.Frame fraAppend 
      Height          =   570
      Left            =   0
      TabIndex        =   18
      ToolTipText     =   "清除:F6"
      Top             =   5280
      Width           =   11895
      Begin VB.ComboBox cboBaby 
         Height          =   300
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   165
         Width           =   1800
      End
      Begin VB.ComboBox cbo开单人 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   7110
         TabIndex        =   6
         Top             =   180
         Width           =   1890
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   300
         Left            =   9840
         TabIndex        =   7
         Top             =   180
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   529
         _Version        =   393216
         AutoTab         =   -1  'True
         HideSelection   =   0   'False
         MaxLength       =   19
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "YYYY-MM-DD HH:mm:ss"
         Mask            =   "####-##-## ##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.CheckBox chk加班 
         Caption         =   "加班执行(&A)"
         Height          =   270
         Left            =   240
         TabIndex        =   3
         Top             =   195
         Width           =   1395
      End
      Begin VB.Label lblBaby 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "婴儿费(&B)"
         Height          =   240
         Left            =   1560
         TabIndex        =   4
         Top             =   225
         Width           =   1080
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         Caption         =   "费用时间"
         Height          =   180
         Left            =   9075
         TabIndex        =   24
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lbl开单人 
         AutoSize        =   -1  'True
         Caption         =   "开单人"
         Height          =   180
         Left            =   6480
         TabIndex        =   21
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.PictureBox picAppend 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   11760
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   6360
      Width           =   11760
      Begin VB.TextBox txt病人 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   450
         Left            =   5685
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Tag             =   "5685"
         Text            =   "0.00"
         Top             =   90
         Width           =   1815
      End
      Begin VB.TextBox txt单据 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   450
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Tag             =   "2760"
         Text            =   "0.00"
         Top             =   90
         Width           =   1815
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   10035
         TabIndex        =   13
         ToolTipText     =   "热键:Esc"
         Top             =   165
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   8745
         TabIndex        =   12
         ToolTipText     =   "热键：F2"
         Top             =   165
         Width           =   1100
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
         Height          =   420
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "导入访帐表:F3"
         Top             =   105
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtIn 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   690
         MaxLength       =   8
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "全清(&R)"
         Height          =   350
         Left            =   7260
         TabIndex        =   15
         Top             =   165
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.CommandButton cmdSelALL 
         Caption         =   "全选(&A)"
         Height          =   350
         Left            =   6150
         TabIndex        =   14
         Top             =   165
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.Label lbl单据 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "合计"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   2205
         TabIndex        =   30
         Top             =   195
         Width           =   510
      End
      Begin VB.Label lbl病人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   5130
         TabIndex        =   29
         Top             =   195
         Width           =   510
      End
   End
   Begin VB.Frame fraTitle 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   0
      TabIndex        =   22
      Top             =   420
      Width           =   11910
      Begin VB.PictureBox picUnit 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   180
         ScaleHeight     =   375
         ScaleWidth      =   2730
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   150
         Width           =   2730
         Begin VB.ComboBox cbo开单科室 
            Height          =   300
            Left            =   840
            TabIndex        =   0
            Text            =   "cbo开单科室"
            Top             =   60
            Width           =   1905
         End
         Begin VB.Label lbl开单科室 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "开单科室"
            Height          =   180
            Left            =   60
            TabIndex        =   32
            Top             =   120
            Width           =   720
         End
      End
      Begin VB.ComboBox cboNO 
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   210
         Width           =   1305
      End
      Begin VB.CheckBox chkCancel 
         Caption         =   "销"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   11310
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "热键:F8"
         Top             =   165
         Width           =   525
      End
      Begin VB.Label lblFlag 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "销"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   11400
         TabIndex        =   27
         Top             =   195
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label lblNO 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "单据号"
         ForeColor       =   &H80000007&
         Height          =   180
         Left            =   9360
         TabIndex        =   26
         Top             =   270
         Width           =   540
      End
   End
   Begin VB.Frame fraDrawDept 
      Height          =   645
      Left            =   0
      TabIndex        =   33
      Top             =   5730
      Visible         =   0   'False
      Width           =   13575
      Begin VB.ComboBox cboDrawDept 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   225
         Width           =   3315
      End
      Begin VB.Label lblDrawDrugDept 
         AutoSize        =   -1  'True
         Caption         =   "领药部门"
         Height          =   180
         Left            =   495
         TabIndex        =   8
         Top             =   270
         Width           =   720
      End
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "住院记帐表"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   195
      TabIndex        =   25
      ToolTipText     =   "清除:F6"
      Top             =   60
      Width           =   1875
   End
End
Attribute VB_Name = "frmBillings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
'————————————————————————————————————————————————————————————————————
'入口参数：
'表单初始状态参数:
Public mbytInState As Byte '0-执行,1-浏览,2-调整,3-销帐
Public mstrInNO As String '当mbytInState=1,2,3时有效,等于单据号
Public mblnNOMoved As Boolean '操作的单据是否在后备数据表中

Public mstr病人IDs As String '销帐时,允许的病人ID串,空为所有
Public mstrTime As String '操作单据内容的登记时间
Public mblnDelete As Boolean '是否查阅退费单据
Public mlngDelRow As Long '从外部调用销帐时，缺省销帐的费用记录

Public mlngUnitID As Long '当前记帐病区,为0时表示所有病区
Public mlngDeptID As Long '当前记帐科室,为0时表示所有科室
Public mbytUseType As Byte '记帐单用途,0-普通记帐,1-按科室分散记帐
Public mlng病人ID As Long '科室分散记帐用
Public mstrPrivs As String
Public mlngModule As Long

Private mstrPrivsOpt As String '记帐操作1150模块的授权功能

'————————————————————————————————————————————————————————————————————
'数据对象
Private mrsWork As New ADODB.Recordset '当天上班的药房
Private mblnWork As Boolean '当前是否有正在上班的药房
Private mlng药品类别ID As Long '当前单据操作的药品入出类别ID
Private mlng卫材类别ID As Long '当前单据操作的卫材入出类别ID
''''''''''''''''
Private mrsClass As ADODB.Recordset '根据参数读取的当前可用的收费类别
Private mrsMedPayMode As ADODB.Recordset '所有可用的医疗付款方式
Private mrsLevel As ADODB.Recordset '可选病人费别
Private mrsUnit As ADODB.Recordset '可选择的执行科室
Private mrsInfo As New ADODB.Recordset '病人信息(含病人标识、就诊卡号、主页ID、住院号、床号、姓名、性别、年龄、费别)
Private mrs开单科室 As ADODB.Recordset  '可选的开单科室
Private mrs开单人 As ADODB.Recordset    '可选医生和护士
Private mrs领药部门 As ADODB.Recordset

Private mstrUseMoney As String  '当前病人剩余款
Private mstrUnitIDs As String   '当前操作员的所有病区ID
'程序对象
Private mobjBill As ExpenseBill '★★★费用单据对象★★★
Private mobjBillDetail As BillDetail '单据的收费细目对象
Private mobjBillIncome As BillInCome '收费细目的收入项目对象
Private mobjDetail As Detail '单独的收费细目对象
Private mcolDetails As Details '单独的收费细目集合
Private mcolMoneys As BillInComes  '★★收入项目汇总集合(显示及打印时使用)★★

Private Enum BillColType       '单据控件的列类型
    CheckBox = -1
    Text_UnModify = 0
    CommandButton = 1
    Date = 2
    ComboBox = 3
    Text = 4
    UnFocus = 5
End Enum

Private Enum BillCol
    姓名 = 0
    性别 = 1
    年龄 = 2
    床号 = 3
    费别 = 4
    类别 = 5
    项目 = 6
    商品名 = 7
    规格 = 8
    单位 = 9
    付数 = 10
    数次 = 11
    单价 = 12
    应收金额 = 13
    实收金额 = 14
    执行科室 = 15
    标志 = 16
    类型 = 17
End Enum

Private Enum Pan
    C2提示信息 = 2
End Enum


'程序变量
Private mblnSendMateria As Boolean  '记帐后自动发药
Private mstrWarn As String '病人报警后选择继续的情况(eg:";张三:DEF5;李四:已报警;王五:567DF;张三:G")
Private mrsWarn As ADODB.Recordset  '所有病区报警数据
Private mcolStock1 As Collection '存放各个药品库房的出库检查方式
Private mcolStock2 As Collection '存放各个卫材库的出库检查方式

Private mlngPreRow As Long '当前行号,用于列改变时判断
Private mblnEnterCell As Boolean '控制是否执行EnterCell事件
Private mcurModiMoney As Currency '修改的当前单据当前病人的金额,在bill_entercell中取值


Private mblnDrop As Boolean '在KeyDown中判断cbo开单人当前是否弹出
Private mblnPrint As Boolean '读取审核单时是否包含要打印的收费类别
Private marrColData() As Integer '当前单据编辑属性映象
Private marrSerial() As Integer '记录销帐单据费用行的序号

Private mcolPatiInfo As Collection '记录销帐单据的病人ID,主页ID,婴儿费

Private mblnOne As Boolean      '是否只有一个可用收费类别
Private mblnSelect As Boolean '用于控制收费细目对象是否来自于列表选择或选择器
Private mlngPreUnit As Long '作用简单
Private Const STR_HEAD = "病人,750,1;病人,300,1;病人,300,1;床号,450,1;费别,500,1;类别,650,1;项目,1700,1;商品名,1800,1;规格,950,1;" & "单位,550,4;付,300,1;数次,450,1;单价,850,7;应收金额,850,7;实收金额,850,7;执行科室,1200,1;标志,450,5;类型,450,1"
Private mstr药品价格等级 As String, mstr卫材价格等级 As String, mstr普通价格等级 As String

Private Sub Bill_cboKeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    '刘兴洪 问题:27378 日期:2010-01-27 13:35:37
    If Bill.cboStyle = DropOlnyDown Then Exit Sub
    
    Select Case Bill.TextMatrix(0, Bill.Col)
        Case "执行科室"
            If Bill.ListIndex <> -1 Then Exit Sub
        Case "发药药店"
            If Bill.ListIndex <> -1 Then Exit Sub
        Case Else
        Exit Sub
    End Select
    lngRow = Bill.Row
    If mobjBill.Details.Count < lngRow Then Exit Sub
    
    With mobjBill.Details(lngRow)
        If InStr(",4,5,6,7,", .收费类别) > 0 Then
            If mrsWork Is Nothing Then Exit Sub
            If mrsWork.State <> 1 Then Exit Sub
            If zlSelectDept(Me, mlngModule, Bill.cboObj, mrsWork, Bill.CboText, True, , False) = False Then Exit Sub
        Else
            If mrsUnit Is Nothing Then Exit Sub
            If mrsUnit.State <> 1 Then Exit Sub
            If zlSelectDept(Me, mlngModule, Bill.cboObj, mrsUnit, Bill.CboText, True, , False) = False Then Exit Sub
        End If
    End With
    Exit Sub
End Sub

Private Sub Bill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Dim i As Long, j As Long, bytsubs As Integer
    Dim bln从项汇总折扣 As Boolean
    Dim lngMainRow As Long
    
    If mbytInState <> 0 Or chkCancel.Value = 1 Then Cancel = True: Exit Sub
    
    If mobjBill.Details.Count >= Row Then
        '带从属项目的项删除确认
        For i = Row + 1 To mobjBill.Details.Count
            If mobjBill.Details(i).从属父号 = Row Then bytsubs = bytsubs + 1
        Next
        If bytsubs > 0 Then
            If MsgBox("该项目带有 " & bytsubs & " 个从属项目,删除该项目也将删除它的从属项目,继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True: Exit Sub
            End If
        ElseIf mobjBill.Details(Row).从属父号 <> 0 Then '从属项目删除确认
            If MsgBox("该项目是[" & mobjBill.Details(mobjBill.Details(Row).从属父号).Detail.名称 & "]的从属项目,确定要删除它吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True: Exit Sub
            Else
                bln从项汇总折扣 = gbln从项汇总折扣
            End If
        ElseIf MsgBox("确实要删除该收费项目吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = True: Exit Sub
        End If
        
        If bln从项汇总折扣 Then lngMainRow = mobjBill.Details(Bill.Row).从属父号 '如果是从项,删除之前记下从项的从属父号,如果是主项,则级联删除,不用重算
        
        '删除处理
        For i = mobjBill.Details.Count To Row + 1 Step -1
            If mobjBill.Details(i).从属父号 = Row Then
                Call DeleteDetail(i) '反顺序删除其从属行
            End If
        Next
        Call DeleteDetail(Row) '删除该行
        
        '删除只输入了病人的上面行时,该病人信息必须清除
        For i = 1 To Bill.Rows - 1
            If mobjBill.Details.Count < i Then
                For j = 0 To Bill.Cols - 1
                    Bill.TextMatrix(i, j) = ""
                Next
            End If
        Next
        
        '重新计算并刷新
        If bln从项汇总折扣 Then
            If CheckItemHaveSub(lngMainRow) Then
                Call Calc重算主项实收(lngMainRow)
            Else
                Call CalcMoney(lngMainRow, False) '只有一个主项了,从项全部被删除时,当成普通独立项计算
            End If
        End If
        
        Call ShowDetails
        
        Bill.TxtVisible = False
        Bill.CmdVisible = False
        Bill.CboVisible = False
        
        Cancel = True '不用控件来处理删除
        
        mlngPreRow = 0  '表示行改变了
        Call Bill_EnterCell(Bill.Row, Bill.Col)
        Call SetDrawDrugDeptEnabled
    End If
End Sub

Private Sub Bill_CellCheck(Row As Long, Col As Long)
'说明：可以全部为主要手术,但不能全部为附加手术
    Dim i As Long, strCheck As String, bytTime As Integer
    Dim blnReSet As Boolean
    
    If Bill.TextMatrix(Row, BillCol.性别) = "" Then Bill.TextMatrix(Row, Col) = "": Exit Sub
    
    If mbytInState = 3 Or (chkCancel.Visible And chkCancel.Value = 1) Then Exit Sub
    
    If mobjBill.Details.Count < Row Then
        Bill.TextMatrix(Row, Col) = "" '新增的未处理行无效
        Exit Sub
    End If
    
    strCheck = Bill.TextMatrix(Row, Col)
    '问题:  If mobjBill.Details(i).病人ID = mobjBill.Details(Row).病人ID Then '刘兴洪:多病人时,没有检查相关的病人
    
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).病人ID = mobjBill.Details(Row).病人ID Then
            If mobjBill.Details(i).收费类别 = "F" And mobjBill.Details(i).附加标志 = 0 And i <> Row Then bytTime = bytTime + 1
        End If
    Next
    
    blnReSet = bytTime > 0
    If blnReSet = False Then     '可能只存在附加手术后又改成了主手术,需要重新计处理:25495
        blnReSet = (strCheck = "" And mobjBill.Details(Row).收费类别 = "F" And mobjBill.Details(Row).附加标志 = 1)
    End If
    
    If blnReSet Then
        mobjBill.Details(Row).附加标志 = IIf(strCheck = "", 0, 1)
        Call CalcMoneys(Row)
        Call ShowDetails(Row)
        CalcOneTotal (Bill.Row)
    ElseIf strCheck <> "" Then
        Bill.TextMatrix(Row, Col) = ""
        MsgBox "单据中必然有一个手术不是附加手术！", vbInformation, gstrSysName
    End If
End Sub

Private Sub Bill_CommandClick()
    Dim lng项目id As Long, blnCancel As Boolean, bln护士 As Boolean, int险类 As Integer
    Dim str类别 As String, str特准项目 As String
    Dim int病人性质 As Integer, int病人来源 As Integer
    Dim lng病人ID As Long
    
    Call GetOperatorInfo(mrs开单人, mobjBill.开单人, bln护士)
    If gbln收费类别 Then
        If Bill.RowData(Bill.Row) <> 0 Then
            str类别 = "'" & Chr(Bill.RowData(Bill.Row)) & "'"
        Else
            str类别 = IIf(bln护士, "'E','M','4'", gstr收费类别)
        End If
    Else
        str类别 = IIf(bln护士, "'E','M','4'", gstr收费类别)
    End If
    
    '医保病人特准项目
    If mobjBill.Details.Count >= Bill.Row Then
        If Val(mobjBill.Details(Bill.Row).发药窗口) > 0 Then '险类
            int险类 = Val(mobjBill.Details(Bill.Row).发药窗口)
            '刘兴洪:24862
            If zl_Check特准项目(gclsInsure, int险类, mobjBill.Details(Bill.Row).病人ID, False) Then str特准项目 = Get保险特准项目(mobjBill.Details(Bill.Row).病人ID, "A.ID")
        End If
        lng病人ID = mobjBill.Details(Bill.Row).病人ID
    ElseIf mrsInfo.State = 1 Then
        If Not IsNull(mrsInfo!险类) Then
            int险类 = mrsInfo!险类
            '刘兴洪:24862
            If zl_Check特准项目(gclsInsure, int险类, Val(Nvl(mrsInfo!病人ID)), False) Then str特准项目 = Get保险特准项目(Val(Nvl(mrsInfo!病人ID)), "A.ID")
        End If
         lng病人ID = Val(Nvl(mrsInfo!病人ID))
    End If
    
    int病人性质 = -2
    If mobjBill.Details.Count >= Bill.Row Then
        int病人性质 = mobjBill.Details(Bill.Row).病人性质
    ElseIf mrsInfo.State = 1 Then
        int病人性质 = mrsInfo!病人性质
    End If
    If int病人性质 <> -2 Then
        If int病人性质 = 0 Or int病人性质 = 2 Then
            int病人来源 = 2
        ElseIf int病人性质 = 1 Or int病人性质 = -1 Then
            int病人来源 = 1
        End If
    Else
        int病人来源 = 2
    End If
    
    lng项目id = frmItemSelect.ShowSelect(Me, mstrPrivs, int病人来源, int险类, gbln住院单位, str类别, , , str特准项目, _
        zl获取中药形态(lng病人ID, Bill.Row), , , , , mstr药品价格等级, mstr卫材价格等级, mstr普通价格等级)
    If lng项目id <> 0 Then
        Bill.Text = lng项目id
        mblnSelect = True
        Call Bill_KeyDown(13, 0, blnCancel)
        Bill.SetFocus
        If Not blnCancel Then
            Bill.Text = "": Bill.TxtVisible = False
            Call zlCommFun.PressKey(13)
        End If
    Else
        mblnSelect = False
    End If
End Sub

Private Sub Bill_EditKeyPress(KeyAscii As Integer)
    '第一位可以输入字母,其它位不行,汉字的ascii是负数
     If Bill.TextMatrix(0, Bill.Col) = "病人" Then
        If KeyAscii <> 13 Then
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            If Not (Bill.Text = "" Or Bill.SelLength = Len(Bill.Text)) And _
                InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 And KeyAscii > 0 Then
                '如果用床号,则不进行相关的检查
                '53113
                If Left(Bill.Text, 1) = "/" Then Exit Sub
                KeyAscii = 0: Beep: Exit Sub
            End If
        Else
            If Bill.Active And Bill.ColData(Bill.Col) <> BillColType.Text_UnModify Then
                
                 If cbo开单科室.ListIndex <> -1 And Bill.Text = "" Then
                    KeyAscii = 0
                    Call FillPatient(cbo开单科室.ItemData(cbo开单科室.ListIndex))
                    If Bill.Top + Bill.CellTop + lvwPati.Height > sta.Top Then
                        lvwPati.Top = Bill.Top + Bill.CellTop - lvwPati.Height - 30
                    Else
                        lvwPati.Top = Bill.Top + Bill.CellTop + Bill.RowHeight(1) - 15
                    End If
                        
                    lvwPati.Visible = True
                    lvwPati.SetFocus
                End If
            End If
        End If
    End If
End Sub

Private Sub ShowStock(str药品 As String, dbl库存 As Double)
'功能：显示药品或卫材的库存
    If InStr(1, mstrPrivsOpt, ";显示库存;") > 0 Then
        sta.Panels(Pan.C2提示信息).Text = "[" & str药品 & "]可用库存:" & dbl库存
    Else
        sta.Panels(Pan.C2提示信息).Text = "[" & str药品 & "]" & IIf(dbl库存 > 0, "有", "无") & "库存."
    End If
End Sub

Private Sub Bill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
'功能：处理单据输入
    Dim lng病人ID As Long, lng主页ID As Long, dblStock As Double, i As Long, lngCur病人ID As Long
    Dim rsTmp As ADODB.Recordset, strInfo As String
    Dim curTotal As Currency, blnCopy As Boolean, bln护士 As Boolean
    Dim dblPreTime As Double, dblPreMoney As Double, dblNum As Double, dblTemp As Double
    Dim lngDoUnit As Long, strScope As String, curModi As Currency
    Dim str医疗付款 As String, blnSkip As Boolean, blnInput As Boolean
    Dim cur余额 As Currency, cur当日额 As Currency, curItemMoney As Currency
    Dim rs药品信息 As ADODB.Recordset, int险类 As Integer
    Dim lng病人科室ID As Long, lng病人病区ID As Long, int病人来源 As Integer
    Dim str特准项目 As String, int病人性质 As Integer, lngOld付数 As Long
    Dim lng项目id As Long, str类别 As String, str药房IDs As String, str摘要 As String
    Dim bln负数记帐 As Boolean, strPriceGrade As String
    Dim colStock As Collection
    
    On Error GoTo errH
    
    If KeyCode = 13 Then
        If mbytInState = 2 Then
            If Bill.Col = Bill.Cols - 1 And Bill.Row = Bill.Rows - 1 Then
                Cancel = True: Exit Sub
            ElseIf Bill.TextMatrix(0, Bill.Col) <> "执行科室" Then
                Exit Sub
            End If
        End If
        If Bill.ColData(Bill.Col) = BillColType.Text_UnModify Then Exit Sub
        
        Select Case Bill.TextMatrix(0, Bill.Col)
            Case "病人" '读取病人信息
                If Bill.Text <> "" Then
                                        
                    If mobjBill.Details.Count >= Bill.Row Then
                        If Bill.Text = mobjBill.Details(Bill.Row).姓名 Then Exit Sub '双击已读取的病人,未改变时不重读
                    End If
                    Dim blnMsgbox As Boolean
                    If Not GetPatient(Bill.Text, IsNumeric(Bill.Text) And IsNumeric(Left(Bill.Text, 1)), blnMsgbox) Then
                        If Not blnMsgbox Then
                            MsgBox "输入的标识不能读取病人信息，请检查输入是否正确！", vbExclamation, gstrSysName
                        End If
                        If Bill.TxtVisible Then Bill.SelStart = 0: Bill.SelLength = Len(Bill.Text)
                        Call Beep: Cancel = True: Exit Sub
                    Else
                        
                        '病人剩余款(利用就诊卡号域)
                        curModi = 0
                        If mstrInNO <> "" And gbytBilling = 0 Then
                            curModi = GetBillMoney(2, mstrInNO, mrsInfo!病人ID)
                        End If
                        '病人当日费用额
                        cur当日额 = mrsInfo!当日额 - curModi
                        
                        Set rsTmp = GetMoneyInfo(mrsInfo!病人ID, CDbl(curModi), True, 2)
                    
                        '--------------------------------------------------------------------------------------------------------------------------------------------------------------
                        '刘兴洪:26952
                        cur余额 = 0
                        If Not rsTmp Is Nothing Then
                            If rsTmp.State = 1 Then
                                If rsTmp.EOF = False Then
                                    cur余额 = Val(Nvl(rsTmp!预交余额)) - Val(Nvl(rsTmp!费用余额))
                                End If
                            End If
                        End If
                        If gbln报警包含划价费用 Then cur余额 = cur余额 - GetPriceMoneyTotal(1, Val(Nvl(mrsInfo!病人ID))) + IIf(gbytBilling = 1, mcurModiMoney, 0)
                        gbytWarn = BillingWarn(mstrPrivsOpt, Trim(Nvl(mrsInfo!姓名)) & IIf(Nvl(mrsInfo!住院号) = "", "", "(住院号:" & mrsInfo!住院号 & " 床号:" & mrsInfo!床号 & ")"), Val("" & mrsInfo!病区ID), GetPatiWarnRange(Val(Nvl(mrsInfo!病人ID)), IIf(IsNull(mrsInfo!主页ID), 0, mrsInfo!主页ID)), _
                             mrsWarn, cur余额, cur当日额, 0, IIf(IsNull(mrsInfo!担保额), 0, mrsInfo!担保额), "", "", mstrWarn, True, , 0, True)
                        '返回:0;没有报警,继续
                        '     1:报警提示后用户选择继续
                        '     2:报警提示后用户选择中断
                        '     3:报警提示必须中断
                        '     4:强制记帐报警,继续
                        '     5.报警提示后用户选择继续,但只允许保存存为划价单
                        If gbytWarn = 2 Or gbytWarn = 3 Then
                            If Bill.TxtVisible Then Bill.SelStart = 0: Bill.SelLength = Len(Bill.Text)
                            Call Beep: Cancel = True: Exit Sub
                        End If
                        '--------------------------------------------------------------------------------------------------------------------------------------------------------------
                        
                        
                        '第一行定位住院医师
                        If Bill.Row = 1 And cbo开单科室.ListIndex <> -1 Then Call cbo开单科室_Click
                        
                        Bill.Text = IIf(IsNull(mrsInfo!姓名), "", mrsInfo!姓名)
                        Bill.TextMatrix(Bill.Row, BillCol.姓名) = IIf(IsNull(mrsInfo!姓名), "", mrsInfo!姓名)
                        Bill.TextMatrix(Bill.Row, BillCol.性别) = IIf(IsNull(mrsInfo!性别), "", mrsInfo!性别)
                        Bill.TextMatrix(Bill.Row, BillCol.年龄) = IIf(IsNull(mrsInfo!年龄), "", mrsInfo!年龄)
                        Bill.TextMatrix(Bill.Row, BillCol.床号) = IIf(IsNull(mrsInfo!床号), "", mrsInfo!床号)
                        '缺省费别
                        Bill.TextMatrix(Bill.Row, BillCol.费别) = IIf(IsNull(mrsInfo!费别), "", mrsInfo!费别)
                        Bill.MsfObj.CellForeColor = zlDatabase.GetPatiColor(Nvl(mrsInfo!病人类型))
                        Dim dbl预交余额 As Double, dblFee As Double, dbl剩余 As Double
                        If Not rsTmp Is Nothing Then
                            curTotal = CalcOneTotal(Bill.Row)
                            'sta.Panels(3).Text = mrsInfo!姓名 & "预交:" & Format(rsTmp!预交余额, "0.00")
                            'sta.Panels(3).Text = sta.Panels(3) & "/费用:" & Format(rsTmp!费用余额, gstrDec)
                            'sta.Panels(3).Text = sta.Panels(3) & "/余额:" & Format(rsTmp!预交余额 - rsTmp!费用余额, "0.00")
                            dbl预交余额 = Val(Nvl(rsTmp!预交余额)): dblFee = Val(Nvl(rsTmp!费用余额))
                            dbl剩余 = Val(Nvl(rsTmp!预交余额)) - Val(Nvl(rsTmp!费用余额))
                            mstrUseMoney = dbl预交余额 & "," & dblFee & "," & dbl剩余
                        Else
                            'sta.Panels(3).Text = mrsInfo!姓名 & "预交:0.00/费用:" & gstrDec & "/余额:0.00"
                            mstrUseMoney = "0,0,0": dbl预交余额 = 0: dblFee = 0: dbl剩余 = 0
                        End If
                        
                        strInfo = GetPatientDue(Val(mrsInfo!病人ID))
                        'If Val(strInfo) <> 0 Then sta.Panels(3).Text = sta.Panels(3).Text & "/应收款:" & Format(strInfo, "0.00")
                        Call SetStatuPatiInfor(Nvl(mrsInfo!姓名), dbl预交余额, dblFee, dbl剩余, Val(strInfo))
                                                
                        Call LoadPatientBaby(cboBaby, mrsInfo!病人ID, mrsInfo!主页ID)
                        
                        mstrUseMoney = mstrUseMoney & "," & cur当日额
                        
                        
                        If mobjBill.Details.Count >= Bill.Row Then
                            mlngPreRow = 0  '修改已有列时,恢复此值,以便显示库存
                            '修改病人信息
                            With mobjBill.Details(Bill.Row)
                                .病人性质 = IIf(IsNull(mrsInfo!病人性质), 0, mrsInfo!病人性质)
                                .病人ID = IIf(IsNull(mrsInfo!病人ID), 0, mrsInfo!病人ID)
                                .主页ID = IIf(IsNull(mrsInfo!主页ID), 0, mrsInfo!主页ID)
                                .婴儿费 = cboBaby.ItemData(cboBaby.ListIndex)
                                
                                .病区ID = IIf(IsNull(mrsInfo!病区ID), 0, mrsInfo!病区ID)
                                .科室ID = IIf(IsNull(mrsInfo!科室ID), 0, mrsInfo!科室ID)
                                
                                .担保额 = IIf(IsNull(mrsInfo!担保额), 0, mrsInfo!担保额)
                                
                                .床号 = IIf(IsNull(mrsInfo!床号), "", mrsInfo!床号)
                                .住院号 = IIf(IsNull(mrsInfo!住院号), 0, mrsInfo!住院号)
                                .姓名 = Bill.TextMatrix(Bill.Row, BillCol.姓名)
                                .性别 = IIf(IsNull(mrsInfo!性别), "", mrsInfo!性别)
                                .年龄 = Bill.TextMatrix(Bill.Row, BillCol.年龄)
                                
                                '仅记帐表中,记帐时,发药窗口用于记录该病人的险类
                                .发药窗口 = IIf(IsNull(mrsInfo!险类), "", mrsInfo!险类)
                                
                                .费别 = zlStr.NeedName(Bill.TextMatrix(Bill.Row, BillCol.费别))
                                .医疗付款 = IIf(IsNull(mrsInfo!医疗付款方式), "", mrsInfo!医疗付款方式)
                                
                                .就诊卡号 = mstrUseMoney
                                
                                '如果有对应的从属项目,则从属项目的病人信息也更改
                                For i = Bill.Row + 1 To mobjBill.Details.Count
                                    If mobjBill.Details(i).从属父号 = Bill.Row Then
                                        mobjBill.Details(i).病人性质 = .病人性质
                                        mobjBill.Details(i).病人ID = .病人ID
                                        mobjBill.Details(i).主页ID = .主页ID
                                        mobjBill.Details(i).病区ID = .病区ID
                                        mobjBill.Details(i).科室ID = .科室ID
                                        mobjBill.Details(i).担保额 = .担保额
                                        mobjBill.Details(i).床号 = .床号
                                        mobjBill.Details(i).住院号 = .住院号
                                        mobjBill.Details(i).姓名 = .姓名
                                        mobjBill.Details(i).性别 = .性别
                                        mobjBill.Details(i).年龄 = .年龄
                                        mobjBill.Details(i).发药窗口 = .发药窗口
                                        mobjBill.Details(i).费别 = .费别
                                        mobjBill.Details(i).医疗付款 = .医疗付款
                                        mobjBill.Details(i).就诊卡号 = .就诊卡号
                                        mobjBill.Details(i).婴儿费 = .婴儿费
                                    End If
                                Next
                            End With
                        End If
                        
                        If Not IsNull(mrsInfo!出院日期) Then
                            MsgBox "提醒您：" & vbCrLf & vbCrLf & "该病人已于 " & Format(mrsInfo!出院日期, "yyyy-MM-dd") & " 出院，现在对该病人强制进行记帐！", vbInformation, gstrSysName
                            If mrsInfo!出院日期 < CDate(txtDate.Text) Then
                                txtDate.Text = Format(mrsInfo!出院日期, "yyyy-MM-dd HH:mm:ss")
                            End If
                        Else
                            txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
                        End If
                        If Not IsNull(mrsInfo!入院日期) Then
                            sta.Panels(2).Text = "入院日期:" & Format(mrsInfo!入院日期, "yyyy-MM-dd")
                            strInfo = GetInsureInfo(mrsInfo!病人ID)
                            If strInfo <> "" Then sta.Panels(2).Text = sta.Panels(2).Text & "/帐号:" & Split(strInfo, ";")(1)
                        End If
                        
                        If mobjBill.Details.Count = 0 Then txt单据.Text = gstrDec
                        
                        '如果最末行有未输入费用的病人则删除该行(未确认行)
                        If mobjBill.Details.Count = Bill.Rows - 2 And Bill.Row = Bill.Rows - 2 Then
                            Bill.RemoveMSFItem Bill.Rows - 1
                        End If
                        
                        '病人变了，重新按费别算
                        If mobjBill.Details.Count >= Bill.Row Then
                            Call CalcMoneys(Bill.Row)
                            Call ShowDetails(Bill.Row)
                            '从属项目也需要重新计算和刷新显示
                            For i = Bill.Row + 1 To mobjBill.Details.Count
                                If mobjBill.Details(i).从属父号 = Bill.Row Then
                                    Call CalcMoneys(i)
                                    Call ShowDetails(i)
                                End If
                            Next
                        End If
                    End If
                End If
                
                '自动复制上一行的费用项目(新行时病人不同)
                If Bill.Row > 1 And mobjBill.Details.Count < Bill.Row Then
                    If mrsInfo.State = 1 Then '医保病人不复制
                        If IsNull(mrsInfo!险类) Then
                            '如果是时价或分批药品,则不自动复制(禁止重复)
                            If mobjBill.Details(Bill.Row - 1).姓名 <> mrsInfo!姓名 And mobjBill.Details(Bill.Row - 1).从属父号 = 0 _
                                And Not (mobjBill.Details(Bill.Row - 1).Detail.变价 _
                                    Or mobjBill.Details(Bill.Row - 1).Detail.分批) And Not (mobjBill.Details(Bill.Row - 1).收费类别 = "F" And mobjBill.Details(Bill.Row - 1).附加标志 = 1) Then
                                
                                '刘兴洪:加入附加手术不复制: Not (mobjBill.Details(Bill.Row - 1).收费类别 = "F" And mobjBill.Details(Bill.Row - 1).附加标志 = 1)
                                '问题:
                                blnCopy = True '标志要复制费用了
                                
                                With mobjBill.Details(Bill.Row - 1)
                                    mobjBill.Details.Add .Detail, .收费细目ID, .序号 + 1, .从属父号, IIf(IsNull(mrsInfo!病人ID), 0, mrsInfo!病人ID), _
                                    IIf(IsNull(mrsInfo!主页ID), 0, mrsInfo!主页ID), IIf(IsNull(mrsInfo!病区ID), 0, mrsInfo!病区ID), _
                                    IIf(IsNull(mrsInfo!科室ID), 0, mrsInfo!科室ID), IIf(IsNull(mrsInfo!姓名), "", mrsInfo!姓名), _
                                    IIf(IsNull(mrsInfo!性别), "", mrsInfo!性别), IIf(IsNull(mrsInfo!年龄), "", mrsInfo!年龄), _
                                    IIf(IsNull(mrsInfo!住院号), 0, mrsInfo!住院号), IIf(IsNull(mrsInfo!床号), "", mrsInfo!床号), _
                                    IIf(IsNull(mrsInfo!费别), "", mrsInfo!费别), IIf(IsNull(mrsInfo!病人性质), "", mrsInfo!病人性质), _
                                    .收费类别, .计算单位, "", .付数, .数次, .附加标志, .执行部门ID, .InComes, mstrUseMoney, , _
                                    IIf(IsNull(mrsInfo!担保额), 0, mrsInfo!担保额), IIf(IsNull(mrsInfo!医疗付款方式), "", mrsInfo!医疗付款方式)
                                    
                                End With
                                Call CalcMoneys(Bill.Row)
                                
                                Call ShowDetails(Bill.Row)
                                Set mrsInfo = New ADODB.Recordset: mstrUseMoney = "" '当前病人信息无效(以免被其它行误用)
                                Bill.Col = IIf(gbln收费类别, BillCol.类别, BillCol.项目)   '定位到项目输入栏
                                Bill.Text = "": Bill.SetFocus
                                Cancel = True
                            End If
                        End If
                    End If
                    
                    If Bill.ColData(BillCol.类别) <> BillColType.UnFocus And Me.Visible And _
                        Bill.TextMatrix(Bill.Row - 1, BillCol.姓名) = Bill.TextMatrix(Bill.Row, BillCol.姓名) Then _
                        Call zlCommFun.PressKey(13)
                End If
            Case "类别"
                If Bill.ListIndex <> -1 Then '不输入类别时不会定位到类别列
                    If Bill.RowData(Bill.Row) <> Bill.ItemData(Bill.ListIndex) Then
                        '一旦改更收费类别,则清除(如有)原有该项目内容
                        For i = 5 To Bill.Cols - 1
                            Bill.TextMatrix(Bill.Row, i) = ""
                        Next
                        If mobjBill.Details.Count >= Bill.Row Then
                            Set mobjBill.Details(Bill.Row).Detail = New Detail
                            Set mobjBill.Details(Bill.Row).InComes = New BillInComes
                            With mobjBill.Details(Bill.Row)
                                .收费细目ID = 0: .收费类别 = ""
                            End With
                            Call CalcMoneys
                        End If
                    End If
                    Bill.RowData(Bill.Row) = Bill.ItemData(Bill.ListIndex) '暂时用RowData记录所选择的收费类别
                End If
            Case "项目"
                If blnCopy Then Bill.Text = mobjBill.Details(mobjBill.Details.Count).Detail.ID
                
                '此项目确定,该收费细目对应的程序对象才生成,同时这里处理收费从属项目
                If Bill.Text <> "" Then
                    '如果在已输入的项目上按回车,或选择器选择
                    If mobjBill.Details.Count >= Bill.Row Then
                        '通过按钮选择是返回的ID,而输入则是文本,如果是一样的,则不改变
                        If Bill.TextMatrix(Bill.Row, BillCol.项目) = Bill.Text Then
                            Bill.TxtVisible = False
                            Bill.CmdVisible = False
                            Exit Sub
                        End If
                    End If
                        
                    '医保病人特准项目
                    If mobjBill.Details.Count >= Bill.Row Then  '修改已加入的费用
                        If Val(mobjBill.Details(Bill.Row).发药窗口) > 0 Then '险类
                            int险类 = Val(mobjBill.Details(Bill.Row).发药窗口)
                            '刘兴洪:24862
                            If zl_Check特准项目(gclsInsure, int险类, mobjBill.Details(Bill.Row).病人ID, False) Then str特准项目 = Get保险特准项目(mobjBill.Details(Bill.Row).病人ID, "A.ID")
                            
                        End If
                        lngCur病人ID = mobjBill.Details(Bill.Row).病人ID
                    ElseIf mrsInfo.State = 1 Then   '该病人第一条费用
                        If Not IsNull(mrsInfo!险类) Then
                            int险类 = mrsInfo!险类
                            '刘兴洪:24862
                            If zl_Check特准项目(gclsInsure, int险类, Val(Nvl(mrsInfo!病人ID)), False) Then str特准项目 = Get保险特准项目(Val(Nvl(mrsInfo!病人ID)), "A.ID")
                        End If
                        lngCur病人ID = Val(Nvl(mrsInfo!病人ID))
                    ElseIf Bill.TextMatrix(Bill.Row, BillCol.姓名) <> "" And mobjBill.Details.Count < Bill.Row And Bill.Row > 1 Then  '同一病人多条费用
                        If Val(mobjBill.Details(Bill.Row - 1).发药窗口) > 0 Then '险类
                            int险类 = Val(mobjBill.Details(Bill.Row - 1).发药窗口)
                            '刘兴洪:24862
                            If zl_Check特准项目(gclsInsure, int险类, mobjBill.Details(Bill.Row - 1).病人ID, False) Then str特准项目 = Get保险特准项目(mobjBill.Details(Bill.Row - 1).病人ID, "A.ID")
                        End If
                        lngCur病人ID = mobjBill.Details(Bill.Row - 1).病人ID
                    End If
                
                    sta.Panels(2).Text = ""
                    sta.Panels("MedicareType").Text = ""
                    blnInput = True
                    If mblnSelect Or blnCopy Then
                        mblnSelect = False: blnCopy = False '立即清除标志
                        Set mobjDetail = GetInputDetail(Val(Bill.Text), int险类)
                    Else
                        If gbln收费类别 Then
                            If Bill.RowData(Bill.Row) = 0 Then
                                sta.Panels(2) = "没有确定费用类别,请先输入类别！"
                                Bill.TxtSetFocus: Cancel = True: Exit Sub
                            End If
                            str类别 = "'" & Chr(Bill.RowData(Bill.Row)) & "'"
                        Else
                            Call GetOperatorInfo(mrs开单人, mobjBill.开单人, bln护士)
                            str类别 = IIf(bln护士, "'E','M','4'", gstr收费类别)
                        End If
                        
                        int病人性质 = -2
                        If mobjBill.Details.Count >= Bill.Row Then
                            int病人性质 = mobjBill.Details(Bill.Row).病人性质
                        ElseIf mrsInfo.State = 1 Then
                            int病人性质 = mrsInfo!病人性质
                        End If
                        If int病人性质 <> -2 Then
                            If int病人性质 = 0 Or int病人性质 = 2 Then
                                int病人来源 = 2
                            ElseIf int病人性质 = 1 Or int病人性质 = -1 Then
                                int病人来源 = 1
                            End If
                        Else
                            int病人来源 = 2
                        End If
                        
                        lng项目id = frmItemSelect.ShowSelect(Me, mstrPrivs, int病人来源, int险类, gbln住院单位, str类别, _
                            Bill.Text, Bill.TxtHwnd, str特准项目, zl获取中药形态(lngCur病人ID, Bill.Row), _
                            , , , , mstr药品价格等级, mstr卫材价格等级, mstr普通价格等级)
                        If lng项目id <> 0 Then
                            Set mobjDetail = GetInputDetail(lng项目id, int险类)
                            
                            If int险类 <> 0 Then sta.Panels("MedicareType").Text = Get医保大类(lng项目id, int险类)
                        Else
                            Bill.Text = "": Bill.TxtVisible = False
                            Bill.SetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    If mrsInfo.State = 0 And Bill.TextMatrix(Bill.Row, BillCol.姓名) = "" Then
                        sta.Panels(2) = "没有确定病人信息,该行不能继续输入！"
                        Bill.TxtSetFocus: Cancel = True: Exit Sub
                    End If
                    
                    Bill.TxtVisible = False '(不加不行)
                    
                    '收费与发药分离时不允许输入时价及分批药品
                    If InStr(",5,6,7,", mobjDetail.类别) > 0 And gbln分离发药 Then
                        If mobjDetail.变价 Or mobjDetail.分批 Then
                            MsgBox "发药分离处理时不能输入时价或分批药品！", vbInformation, gstrSysName
                            Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    '检查毒理分类和价值分类权限
                    If InStr(",5,6,7,", mobjDetail.类别) > 0 Then
                        Set rs药品信息 = Read药品信息(mobjDetail.ID)
                        If Not rs药品信息 Is Nothing Then
                            If IIf(IsNull(rs药品信息!毒理分类), "", rs药品信息!毒理分类) = "麻醉药" _
                                And InStr(mstrPrivsOpt, ";麻醉药品记帐;") = 0 Then
                                MsgBox """" & mobjDetail.名称 & """为麻醉药品，你没有权限对该类药品记帐！", vbInformation, gstrSysName
                                Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                            ElseIf IIf(IsNull(rs药品信息!毒理分类), "", rs药品信息!毒理分类) = "毒性药" _
                                And InStr(mstrPrivsOpt, ";毒性药品记帐;") = 0 Then
                                MsgBox """" & mobjDetail.名称 & """为毒性药品，你没有权限对该类药品记帐！", vbInformation, gstrSysName
                                Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                            ElseIf (IIf(IsNull(rs药品信息!价值分类), "", rs药品信息!价值分类) = "贵重" _
                                Or IIf(IsNull(rs药品信息!价值分类), "", rs药品信息!价值分类) = "昂贵") _
                                And InStr(mstrPrivsOpt, ";贵重药品记帐;") = 0 Then
                                MsgBox """" & mobjDetail.名称 & """为贵重或昂贵药品，你没有权限对该类药品记帐！", vbInformation, gstrSysName
                                Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                    '求病人ID,险类
                    lng病人ID = 0: int险类 = 0
                    If mrsInfo.State = 1 Then
                        lng病人ID = Nvl(mrsInfo!病人ID, 0)
                        lng主页ID = Nvl(mrsInfo!主页ID, 0)
                        int险类 = Nvl(mrsInfo!险类, 0)
                        lng病人科室ID = Nvl(mrsInfo!科室ID, 0)
                        lng病人病区ID = Nvl(mrsInfo!病区ID, 0)
                        str医疗付款 = "" & mrsInfo!医疗付款方式
                    ElseIf Bill.TextMatrix(Bill.Row, 0) <> "" And mobjBill.Details.Count < Bill.Row And Bill.Row > 1 Then
                        lng病人ID = mobjBill.Details(Bill.Row - 1).病人ID
                        lng主页ID = mobjBill.Details(Bill.Row - 1).主页ID
                        int险类 = Val(mobjBill.Details(Bill.Row - 1).发药窗口)
                        lng病人科室ID = mobjBill.Details(Bill.Row - 1).科室ID
                        lng病人病区ID = mobjBill.Details(Bill.Row - 1).病区ID
                        str医疗付款 = mobjBill.Details(Bill.Row - 1).医疗付款
                    Else
                        lng病人ID = mobjBill.Details(Bill.Row).病人ID
                        lng主页ID = mobjBill.Details(Bill.Row).主页ID
                        int险类 = Val(mobjBill.Details(Bill.Row).发药窗口)
                        lng病人科室ID = mobjBill.Details(Bill.Row).科室ID
                        lng病人病区ID = mobjBill.Details(Bill.Row).病区ID
                        str医疗付款 = mobjBill.Details(Bill.Row).医疗付款
                    End If
                                        
                    '主项适用病人病区科室
                    If InStr(",5,6,7,", mobjDetail.类别) = 0 Then
                        If Not CheckFeeItemLimitDept(mobjDetail.ID, lng病人病区ID, lng病人科室ID) Then
                            MsgBox "该收费项目对当前病人病区和科室不适用！", vbInformation, gstrSysName
                            Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    '检查药品输入是否重复:分批及时价同一药房不允许重复(这里只提醒)
                    If InStr(",5,6,7,", mobjDetail.类别) > 0 _
                        Or (mobjDetail.类别 = "4" And mobjDetail.跟踪在用) Then
                        If PhysicExist(mobjDetail, Bill.Row, lng病人ID) Then
                            Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    
                    '检查处方职务
                    If InStr(",5,6,7,", mobjDetail.类别) > 0 Then
                        mobjDetail.处方职务 = Get处方职务(mobjDetail.ID)
                        '医保或公费病人检查
                        If str医疗付款 <> "" Then
                            If CheckDuty(mobjDetail, False, str医疗付款) > 0 Then
                                Bill.TxtSetFocus: Cancel = True: Exit Sub
                            End If
                        End If
                        '所有病人项目检查
                        If CheckDuty(mobjDetail, True) > 0 Then
                            Bill.TxtSetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    '医保病人费用审批,每行病人可能不同，并且可能修改任意已有行，所以每次读取病人审批项目
                    If int险类 > 0 And mobjDetail.要求审批 Then
                        Set rsTmp = GetAuditRecord(lng病人ID, lng主页ID, mobjDetail.ID)
                        If rsTmp.RecordCount = 0 Then
                            MsgBox "当前病人未被批准使用[" & mobjDetail.名称 & "]！", vbInformation, gstrSysName
                            Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                        ElseIf Not IsNull(rsTmp!可用数量) Then
                            If rsTmp!可用数量 <= 0 Then
                                MsgBox "当前病人使用[" & mobjDetail.名称 & "]已达到批准的使用限量" & FormatEx(rsTmp!使用限量 / IIf(gbln住院单位, mobjDetail.住院包装, 1), 5) & "。", vbInformation, gstrSysName
                                Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                                        
                    
                    '读取药品相关信息
                     '卫材执行科室缺省为病人病区,如果本地指定了,则为指定科室
                    If mobjDetail.类别 = "4" Then
                        lngDoUnit = IIf(glng发料部门 > 0, glng发料部门, lng病人病区ID)
                        If lngDoUnit = 0 Then lngDoUnit = Get开单科室ID
                    End If
                    
                    '病人科室ID
                    If lng病人科室ID = 0 And cbo开单科室.ListIndex <> -1 Then lng病人科室ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
                    
                    int病人来源 = Get病人来源(Bill.Row)
                    lngDoUnit = Get收费执行科室ID(mobjDetail.类别, mobjDetail.ID, _
                        mobjDetail.执行科室, lng病人科室ID, Get开单科室ID, int病人来源, lngDoUnit, lng病人病区ID)
                        
                    '读取药品库存
                    If ReadDrugAndStuffStock(lngDoUnit, mobjDetail) = False Then
                        Bill.TxtSetFocus: Cancel = True: Exit Sub
                    End If
             
                    If InStr(",5,6,7,", mobjDetail.类别) > 0 Then
                        '处方限量
                        mobjDetail.处方限量 = Get处方限量(mobjDetail.ID)
                    End If
                    
                   '保险项目对应检查
                    If int险类 > 0 Then
                        If InStr(",5,6,7,", mobjDetail.类别) > 0 Then
                            strPriceGrade = mstr药品价格等级
                        ElseIf mobjDetail.类别 = "4" Then
                            strPriceGrade = mstr卫材价格等级
                        Else
                            strPriceGrade = mstr普通价格等级
                        End If
                        If Not CheckMediCareItem(mobjDetail.ID, int险类, mobjDetail.名称, mobjDetail.变价 = False, , strPriceGrade) Then
                            Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    '输入摘要(取已有的行以便修改)
                    If mobjBill.Details.Count >= Bill.Row Then
                        If mobjBill.Details(Bill.Row).Detail.ID = mobjDetail.ID Then
                            str摘要 = mobjBill.Details(Bill.Row).摘要
                        End If
                    End If
                    
                    If mrsInfo.State = 1 Then
                        '加入或修改该收费细目行
                        Call SetDetail(mobjDetail, Bill.Row, lngDoUnit)
                        '设置病人信息
                        With mobjBill.Details(Bill.Row)
                            .病人性质 = Nvl(mrsInfo!病人性质, 0)
                            .病人ID = Nvl(mrsInfo!病人ID, 0)
                            .主页ID = Nvl(mrsInfo!主页ID, 0)
                            
                            .病区ID = Nvl(mrsInfo!病区ID, 0)
                            .科室ID = Nvl(mrsInfo!科室ID, 0)
                            
                            .担保额 = Nvl(mrsInfo!担保额, 0)
                            
                            .床号 = "" & mrsInfo!床号
                            .住院号 = Nvl(mrsInfo!住院号, 0)
                            .姓名 = Bill.TextMatrix(Bill.Row, BillCol.姓名)
                            .性别 = Nvl(mrsInfo!性别)
                            .年龄 = Bill.TextMatrix(Bill.Row, BillCol.年龄)
                            .费别 = zlStr.NeedName(Bill.TextMatrix(Bill.Row, BillCol.费别))
                            .医疗付款 = Nvl(mrsInfo!医疗付款方式)
                            
                            '仅记帐表中,记帐时,发药窗口用于记录该病人的险类
                            .发药窗口 = Nvl(mrsInfo!险类)
                            
                            .就诊卡号 = mstrUseMoney
                        End With
                    ElseIf Bill.TextMatrix(Bill.Row, BillCol.姓名) <> "" And mobjBill.Details.Count < Bill.Row And Bill.Row > 1 Then
                        '加入或修改该收费细目行
                        Call SetDetail(mobjDetail, Bill.Row, lngDoUnit)
                        '设置病人信息
                        With mobjBill.Details(Bill.Row)
                            .病人性质 = mobjBill.Details(Bill.Row - 1).病人性质
                            .病人ID = mobjBill.Details(Bill.Row - 1).病人ID
                            .主页ID = mobjBill.Details(Bill.Row - 1).主页ID
                            
                            .病区ID = mobjBill.Details(Bill.Row - 1).病区ID
                            .科室ID = mobjBill.Details(Bill.Row - 1).科室ID
                            
                            .担保额 = mobjBill.Details(Bill.Row - 1).担保额
                            
                            .床号 = mobjBill.Details(Bill.Row - 1).床号
                            .住院号 = mobjBill.Details(Bill.Row - 1).住院号
                            .姓名 = mobjBill.Details(Bill.Row - 1).姓名
                            .性别 = mobjBill.Details(Bill.Row - 1).性别
                            .年龄 = mobjBill.Details(Bill.Row - 1).年龄
                            .费别 = Mid(Bill.TextMatrix(Bill.Row, BillCol.费别), InStr(Bill.TextMatrix(Bill.Row, BillCol.费别), "-") + 1)
                            .医疗付款 = mobjBill.Details(Bill.Row - 1).医疗付款
                            
                            '仅记帐表中,记帐时,发药窗口用于记录该病人的险类
                            .发药窗口 = mobjBill.Details(Bill.Row - 1).发药窗口
                            
                            .就诊卡号 = mobjBill.Details(Bill.Row - 1).就诊卡号
                        End With
                    Else
                        '加入或修改该收费细目行,仅修改费用项目
                        Call SetDetail(mobjDetail, Bill.Row, lngDoUnit)
                    End If
                    '59051
                    '输入摘要(根据新输入的行更改摘要)
                    If mobjBill.Details(Bill.Row).Detail.补充摘要 Then
                        If frmInputBox.InputBox(Me, "摘要", "请输入""" & mobjBill.Details(Bill.Row).Detail.名称 & """的摘要信息:", 200, 3, True, False, str摘要) Then
                            mobjBill.Details(Bill.Row).摘要 = str摘要
                        End If
                    Else
                        str摘要 = gclsInsure.GetItemInfo(int险类, mobjBill.Details(Bill.Row).病人ID, mobjBill.Details(Bill.Row).收费细目ID, str摘要, 2)
                        mobjBill.Details(Bill.Row).摘要 = str摘要
                    End If
                    
                    Call CalcMoneys(Bill.Row)
                    
                    '如果是医保Calcmoney中可能返回摘要
                    If mobjBill.Details(Bill.Row).摘要 <> "" Then str摘要 = mobjBill.Details(Bill.Row).摘要
                    
                    '记帐分类报警(在已经算出该行费用但未显示前)
                    mrsWarn.Filter = ""
                    If mrsWarn.RecordCount > 0 And mobjBill.Details.Count = Bill.Row Then
                        curTotal = CalcOneTotal(Bill.Row)
                        If curTotal > 0 Then
                            If mobjBill.Details(Bill.Row).就诊卡号 = "" Then
                                cur余额 = 0
                                cur当日额 = 0
                            Else
                                cur余额 = Val(Split(mobjBill.Details(Bill.Row).就诊卡号, ",")(2))
                                cur当日额 = Val(Split(mobjBill.Details(Bill.Row).就诊卡号, ",")(3))
                            End If
                            
                            If gbln报警包含划价费用 Then cur余额 = cur余额 - GetPriceMoneyTotal(1, mobjBill.Details(Bill.Row).病人ID) + IIf(gbytBilling = 1, mcurModiMoney, 0)
                            '刘兴洪:24491
                            curItemMoney = GetBillRowTotal(mobjBill.Details(Bill.Row).InComes)
                            
                            gbytWarn = BillingWarn(mstrPrivsOpt, mobjBill.Details(Bill.Row).姓名 & IIf(mobjBill.Details(Bill.Row).住院号 = "", "", "(住院号:" & mobjBill.Details(Bill.Row).住院号 & " 床号:" & mobjBill.Details(Bill.Row).床号 & ")"), lng病人病区ID, GetPatiWarnRange(mobjBill.Details(Bill.Row).病人ID, mobjBill.Details(Bill.Row).主页ID), _
                                mrsWarn, cur余额, cur当日额, curTotal, mobjBill.Details(Bill.Row).担保额, mobjBill.Details(Bill.Row).收费类别, mobjBill.Details(Bill.Row).Detail.类别名称, mstrWarn, True, , curItemMoney)
                            
                            If gbytWarn = 2 Or gbytWarn = 3 Then
                                mobjBill.Details.Remove Bill.Row
                                For i = 0 To Bill.Cols - 1
                                    Bill.TextMatrix(Bill.Row, i) = ""
                                Next
                                Bill.Text = "": Cancel = True
                                Bill.Col = BillCol.姓名: Bill.SetFocus
                                Exit Sub
                            End If
                        End If
                    End If
                    
                    If int险类 <> 0 And mobjBill.Details(Bill.Row).数次 <> 0 Then
                        If gclsInsure.GetCapability(support实时监控, mobjBill.Details(Bill.Row).病人ID, int险类) Then
                            If gclsInsure.CheckItem(int险类, 1, 0, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 2, gbytBilling, Bill.Row)) = False Then
                                mobjBill.Details.Remove Bill.Row '删除刚刚想要加入的费用行
                                Bill.Text = "": Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                    If mobjBill.Details(Bill.Row).数次 <> 0 Then
                        If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModule, 1, 0, _
                            MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 2, gbytBilling, Bill.Row)) = False Then
                            mobjBill.Details.Remove Bill.Row '删除刚刚想要加入的费用行
                            Bill.Text = "": Cancel = True: Exit Sub
                        End If
                    End If
                    
                    Call ShowDetails(Bill.Row)
                    
                    '费用类型检查
                    Call Check费用类型(Bill.Row)

                    
                    Set mrsInfo = New ADODB.Recordset: mstrUseMoney = ""
                    '当前病人信息无效(以免被其它行误用)
                    
                    Bill.Text = "": Bill.SetFocus
                ElseIf mobjBill.Details.Count < Bill.Row Then
                    Call zlCommFun.PressKey(vbKeyTab): Exit Sub
                End If
                
                If mobjBill.Details.Count >= Bill.Row Then
                    With mobjBill.Details(Bill.Row)
                        '下一列的性质确定
                        If .收费类别 = "7" And gblnPay Then Bill.ColData(BillCol.付数) = BillColType.Text   '付数
                        If .收费类别 = "F" Then Bill.ColData(BillCol.标志) = BillColType.CheckBox '附加标志
                        
                        '变价允许输入数次
                        If .Detail.变价 And InStr(",5,6,7,", .收费类别) = 0 _
                            And Not (.收费类别 = "4" And .Detail.跟踪在用) Then
                            Bill.ColData(BillCol.数次) = IIf(gblnTime, BillColType.Text, BillColType.UnFocus) '数次
                            Bill.ColData(BillCol.单价) = BillColType.Text '单价
                        Else
                            Bill.ColData(BillCol.数次) = BillColType.Text '数次
                            Bill.ColData(BillCol.单价) = BillColType.UnFocus '单价
                        End If
                        
                        '执行科室
                        If InStr(",5,6,7,", .收费类别) > 0 And gbln分离发药 Then
                            Bill.ColData(BillCol.执行科室) = BillColType.UnFocus: .Key = 1
                        Else
                            '在FillBillComboBox中设置ListIndex时调用CboClick事件
                            mblnEnterCell = False: Bill.Col = BillCol.执行科室: mblnEnterCell = True
                            Call FillBillComboBox(Bill.Row, BillCol.执行科室, lng病人科室ID, int病人来源, Not blnInput) '直接回车时保持执行科室
                            mblnEnterCell = False: Bill.Col = BillCol.项目: mblnEnterCell = True
                            
                            blnSkip = Bill.ListCount = 1
                            If Not blnSkip And InStr(",4,5,6,7,", .收费类别) > 0 Then
                                '指定了固定药房时,不允许再选择
                                Select Case .收费类别
                                    Case "4"
                                        blnSkip = glng发料部门 > 0 And .执行部门ID = glng发料部门
                                    Case "5"
                                        blnSkip = glng西药房 > 0 And .执行部门ID = glng西药房
                                    Case "6"
                                        blnSkip = glng成药房 > 0 And .执行部门ID = glng成药房
                                    Case "7"
                                        blnSkip = glng中药房 > 0 And .执行部门ID = glng中药房
                                End Select
                            End If
                            If blnSkip Then
                                Bill.ColData(BillCol.执行科室) = BillColType.UnFocus: .Key = 1
                            Else
                                Bill.ColData(BillCol.执行科室) = BillColType.ComboBox: .Key = Bill.ListCount
                            End If
                            If lngDoUnit <> .执行部门ID Then
                                '读取药品库存
                                If ReadDrugAndStuffStock(.执行部门ID, mobjBill.Details(Bill.Row).Detail) = False Then
                                    Bill.TxtSetFocus: Cancel = True: Exit Sub
                                End If
                            End If
                        End If
                        
                        '检查卫生材料的灭菌效期,在确定执行科室之后
                        If .收费类别 = "4" And .Detail.跟踪在用 Then
                            Call CheckValidity(.收费细目ID, .执行部门ID, .数次, False) '已确认输入,仅能提醒
                        End If
                        
                        '从属项目处理
                        If Bill.TextMatrix(0, Bill.Col) = "项目" And InStr(",5,6,7,", .收费类别) = 0 Then
                            If (gbln从项汇总折扣 And mobjBill.Details(Bill.Row).从属父号 = 0) Or Not gbln从项汇总折扣 Then  '(如果有级联,只取一级)
                                If ShouldDO(Bill.Row) Then
                                   Call SetSubItem(lng病人科室ID, int病人来源)
                                   mlngPreRow = 0 '通过行变化标志来重新确定列性质
                                End If
                            End If
                        End If
                    End With
                End If
                
                '只输入一次付数
                If mobjBill.Details.Count >= Bill.Row And Bill.Row >= 2 And Bill.Active And Visible Then
                    If mobjBill.Details(Bill.Row).收费类别 = "7" Then
                        For i = 1 To Bill.Row - 1
                            If mobjBill.Details(i).收费类别 = "7" Then
                                '正常执行该过程：本身会定位下一个单元,先定位到付数,则下一个单元是数次
                                '选择调用该过程：调用后会送个回车，这里不能再回车，否则是三个回车的效果(控件原因)。
                                Bill.Col = BillCol.付数: Exit For
                            End If
                        Next
                    End If
                End If
                Call SetDrawDrugDeptEnabled
            Case "付"
                If mobjBill.Details.Count >= Bill.Row And Bill.Text <> "" Then
                    '数值合法性
                    If Not IsNumeric(Bill.Text) Then
                        MsgBox "非法数值！", vbInformation, gstrSysName
                        Bill.Text = mobjBill.Details(Bill.Row).付数: Cancel = True: Exit Sub
                    End If
                    If Val(Bill.Text) <= 0 Or Val(Bill.Text) <> Int(Val(Bill.Text)) Then
                        MsgBox "付数应该为正的整数！", vbInformation, gstrSysName
                        Bill.Text = mobjBill.Details(Bill.Row).付数: Cancel = True: Exit Sub
                    End If
                    '最大金额检查
                    If gcurMaxMoney > 0 Then
                        If CSng(Bill.Text) * mobjBill.Details(Bill.Row).数次 * Bill.TextMatrix(Bill.Row, BillCol.单价) > gcurMaxMoney Then
                            If MsgBox("当前金额超过了" & gcurMaxMoney & ",你确定要继续吗?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                                Bill.Text = mobjBill.Details(Bill.Row).付数: Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                
                    '仅中药及非从属项目才可更改付数(主项付数改变,从属也变)
                    If mobjBill.Details(Bill.Row).收费类别 = "7" Then 'And mobjBill.Details(Bill.Row).从属父号 = 0 Then
                        '分批或时价药品不足禁止输入(没有分批的时价药品可以修改付数、数次)
                        If mobjBill.Details(Bill.Row).Detail.分批 Or mobjBill.Details(Bill.Row).Detail.变价 Then
                            If CSng(Bill.Text) * mobjBill.Details(Bill.Row).数次 > mobjBill.Details(Bill.Row).Detail.库存 Then
                                MsgBox """" & mobjBill.Details(Bill.Row).Detail.名称 & """为分批或时价药品,当前可用库存不足输入数量！", vbInformation, gstrSysName
                                Bill.Text = mobjBill.Details(Bill.Row).付数: Cancel = True: Exit Sub
                            End If
                        End If
                              
                        '检查其它时价或分批中药更改付数后库存是否足够
                        For i = 1 To mobjBill.Details.Count
                            If i <> Bill.Row And mobjBill.Details(i).收费类别 = "7" _
                                And (mobjBill.Details(i).Detail.变价 Or mobjBill.Details(i).Detail.分批) Then
                                If Val(Bill.Text) * mobjBill.Details(i).数次 > mobjBill.Details(i).Detail.库存 Then
                                    MsgBox """" & mobjBill.Details(i).Detail.名称 & """为分批或时价药品,当前可用库存不足输入数量！", vbInformation, gstrSysName
                                    Bill.Text = mobjBill.Details(Bill.Row).付数: Cancel = True: Exit Sub
                                End If
                            End If
                        Next
                        
                        '计算并刷新该行
                        lngOld付数 = mobjBill.Details(Bill.Row).付数
                        mobjBill.Details(Bill.Row).付数 = Bill.Text
                        Call CalcMoneys(Bill.Row)
                        
                        
                        int险类 = Val(mobjBill.Details(Bill.Row).发药窗口)
                        If int险类 <> 0 And mobjBill.Details(Bill.Row).数次 <> 0 Then
                            If gclsInsure.GetCapability(support实时监控, mobjBill.Details(Bill.Row).病人ID, int险类) Then
                                If gclsInsure.CheckItem(int险类, 1, 0, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 2, gbytBilling, Bill.Row)) = False Then
                                    mobjBill.Details(Bill.Row).付数 = lngOld付数
                                    Call CalcMoneys(Bill.Row)
                                    Bill.Text = "": Bill.TxtVisible = False
                                    Cancel = True: Exit Sub
                                End If
                            End If
                        End If
                        
                        If mobjBill.Details(Bill.Row).数次 <> 0 Then
                            If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModule, 1, 0, _
                                MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 2, gbytBilling, Bill.Row)) = False Then
                                mobjBill.Details(Bill.Row).付数 = lngOld付数
                                Call CalcMoneys(Bill.Row)
                                Bill.Text = "": Bill.TxtVisible = False
                                Cancel = True: Exit Sub
                            End If
                        End If
                        
                        Call ShowDetails(Bill.Row)
                        CalcOneTotal (Bill.Row)
                        
                        '处理其它中药付数,如果是独立项,则修改其它非从项的,如果是从项,则修改同一主项的从项的.因为限定为中草药,不可能有主项
                        For i = 1 To mobjBill.Details.Count
                            If i <> Bill.Row And mobjBill.Details(i).收费类别 = "7" And mobjBill.Details(i).从属父号 = mobjBill.Details(Bill.Row).从属父号 _
                                And mobjBill.Details(i).病人ID = mobjBill.Details(Bill.Row).病人ID Then
                                If mobjBill.Details(i).从属父号 = 0 Or (mobjBill.Details(i).从属父号 <> 0 And mobjBill.Details(i).Detail.固有从属 = 0) Then     '1和2固定和按比例的不改
                                    mobjBill.Details(i).付数 = Bill.Text
                                    Call CalcMoneys(i)
                                    Call ShowDetails(i)
                                End If
                            End If
                        Next
                    Else
                        sta.Panels(2) = "从属项目的付数不能更改！"
                        Bill.Text = mobjBill.Details(Bill.Row).付数: Beep '恢复原有付数值
                    End If
                End If
            Case "数次"
                If mobjBill.Details.Count >= Bill.Row And Bill.Text <> "" Then
                    With mobjBill.Details(Bill.Row)
                        '数值合法性
                        If Not IsNumeric(Bill.Text) Then
                          MsgBox "非法数值！", vbInformation, gstrSysName
                          Bill.Text = .数次: Cancel = True: Exit Sub
                        End If
                        If Val(Bill.Text) = 0 Then
                          If MsgBox("数量输入为零，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                              Bill.Text = .数次: Cancel = True: Exit Sub
                          End If
                        End If
                        '药品输入小数
                        If InStr(",5,6,7,", .收费类别) > 0 Then
                          If Val(Bill.Text) - Int(Val(Bill.Text)) <> 0 And InStr(mstrPrivsOpt, "药品输入小数") = 0 Then
                              MsgBox "你没有权限输入小数！", vbInformation, gstrSysName
                              Bill.Text = .数次: Cancel = True: Exit Sub
                          End If
                        End If
                        '最大金额检查
                        If gcurMaxMoney > 0 Then
                          If CSng(Bill.Text) * .付数 * Bill.TextMatrix(Bill.Row, BillCol.单价) > gcurMaxMoney Then
                              If MsgBox("当前金额超过了" & gcurMaxMoney & ",你确定要继续吗?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                                  Bill.Text = .数次: Cancel = True: Exit Sub
                              End If
                          End If
                        End If
                        
                        Bill.Text = FormatEx(Bill.Text, 5)
                        
                        int险类 = Val(.发药窗口)
                        If InStr(",5,6,7,", .收费类别) > 0 And gbln住院单位 Then
                            dblNum = CSng(Bill.Text) * .付数 * .Detail.住院包装
                        Else
                            dblNum = CSng(Bill.Text) * .付数
                        End If
                        
                        '负数合法性检查
                        If CSng(Bill.Text) * .付数 < 0 Then
                            '权限
                            bln负数记帐 = True
                            If InStr(",5,6,", .收费类别) > 0 Then
                                bln负数记帐 = (InStr(mstrPrivsOpt, ";成药负数记帐;") > 0)
                            ElseIf InStr(",7,", .收费类别) > 0 Then
                                bln负数记帐 = (InStr(mstrPrivsOpt, ";草药负数记帐;") > 0)
                            Else
                                bln负数记帐 = (InStr(mstrPrivsOpt, ";诊疗负数记帐;") > 0)
                            End If
                            
                            If Not bln负数记帐 Then
                                MsgBox "你没有权限输入负数！", vbInformation, gstrSysName
                                Bill.Text = .数次: Cancel = True: Exit Sub
                            Else
                                If .Detail.分批 Then
                                    MsgBox "分批药品不允许输入负数。", vbInformation, gstrSysName
                                    Bill.Text = .数次: Cancel = True: Exit Sub
                                End If
                                If int险类 <> 0 Then
                                    If Not gclsInsure.GetCapability(support负数记帐, .病人ID, int险类) Then
                                        MsgBox "本地医保不支持对医保病人进行负数记帐！", vbInformation, gstrSysName
                                        Bill.Text = .数次: Cancel = True: Exit Sub
                                    End If
                                End If
                            End If
                            
                            '负数冲销数量检查
                            If Not (InStr(",5,6,7,", .收费类别) > 0 And gbln分离发药) Then
                                'dblTemp = zlGetBillOtherRowNumToTal(.病人ID, .主页ID, .收费细目ID, True, Bill.Row)
                                If Not CheckNegative(.病人ID, .主页ID, .收费细目ID, .执行部门ID, dblNum, .Detail.住院包装, mstrPrivsOpt) Then
                                    Bill.Text = .数次: Cancel = True: Exit Sub
                                End If
                            End If
                        End If
                        
                        '药品库存检查
                        If (.收费类别 = "4" And .Detail.跟踪在用) Or (InStr(",5,6,7,", .收费类别) > 0 And Not gbln分离发药) Then
                            If .Detail.分批 Or .Detail.变价 Then
                                '分批或时价药品不足禁止输入
                                If .付数 * CSng(Bill.Text) > .Detail.库存 Then
                                    If .收费类别 = "4" Then
                                        MsgBox """" & .Detail.名称 & """为分批或时价卫生材料,当前可用库存不足输入数量！", vbInformation, gstrSysName
                                    Else
                                        MsgBox """" & .Detail.名称 & """为分批或时价药品,当前可用库存不足输入数量！", vbInformation, gstrSysName
                                    End If
                                    Bill.Text = .数次: Cancel = True: Exit Sub
                                End If
                            Else
                                Set colStock = IIf(.收费类别 = "4", mcolStock2, mcolStock1)
                                If colStock("_" & .执行部门ID) <> 0 And Bill.ColData(BillCol.执行科室) = BillColType.UnFocus Then
                                    '其它药品正常检查
                                    If .付数 * CSng(Bill.Text) > .Detail.库存 Then
                                        If colStock("_" & .执行部门ID) = 1 Then
                                            If MsgBox("""" & .Detail.名称 & """的当前可用库存不足输入数量,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                                Bill.Text = .数次: Cancel = True: Exit Sub
                                            End If
                                        ElseIf colStock("_" & .执行部门ID) = 2 Then
                                            MsgBox """" & .Detail.名称 & """的当前可用库存不足输入数量！", vbInformation, gstrSysName
                                            Bill.Text = .数次: Cancel = True: Exit Sub
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf InStr(",5,6,7,", .收费类别) > 0 And gbln分离发药 Then
                            '没有权限时，固定以提示方式检查
                            str药房IDs = Decode(.收费类别, "5", gstr西药房, "6", gstr成药房, "7", gstr中药房)
                            If str药房IDs <> "" And .付数 * CSng(Bill.Text) > .Detail.库存 Then
                                If gblnStock Then
                                    MsgBox "[" & .Detail.名称 & "]的当前可用库存不足输入数量!", vbInformation, gstrSysName
                                    Bill.Text = .数次: Cancel = True: Exit Sub
                                Else
                                    If MsgBox("""" & .Detail.名称 & """的当前可用库存不足输入数量,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                        Bill.Text = .数次: Cancel = True: Exit Sub
                                    End If
                                End If
                            End If
                        End If
                        
                        dblPreTime = .数次
                        .数次 = Bill.Text
                                            
                        '处方限量检查
                        If Not gbln处方限量 Then
                            If Not CheckLimit(mobjBill, Bill.Row, gbln住院单位) Then
                                .数次 = dblPreTime: Bill.Text = dblPreTime
                                Cancel = True: Exit Sub
                            End If
                        End If
                        If .Detail.录入限量 > 0 And dblNum > .Detail.录入限量 Then
                            If MsgBox("输入的数次超过了录入限量" & FormatEx(.Detail.录入限量 / IIf(gbln住院单位, .Detail.住院包装, 1), 5) & ",是否继续?", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo Then
                                .数次 = dblPreTime: Bill.Text = dblPreTime
                                Cancel = True: Exit Sub
                            End If
                        End If
                        '审批使用限量
                        If int险类 > 0 And .Detail.要求审批 Then
                            Set rsTmp = GetAuditRecord(.病人ID, .主页ID, .收费细目ID)
                            If rsTmp.RecordCount > 0 Then
                                If Not IsNull(rsTmp!可用数量) Then
                                    If dblNum > rsTmp!可用数量 Then
                                        MsgBox "输入的数次超过了批准的使用限量" & FormatEx(rsTmp!可用数量 / IIf(gbln住院单位, .Detail.住院包装, 1), 5) & "。", vbInformation, gstrSysName
                                        .数次 = dblPreTime: Bill.Text = dblPreTime
                                        Cancel = True: Exit Sub
                                    End If
                                End If
                            End If
                        End If
                        
                        
                        '固有从属不能更改数次(主项目数次改变,固有从属的数次也变)
                        If .从属父号 <> 0 And .Detail.固有从属 <> 0 Then
                            sta.Panels(2) = "该项目是固有从属项目,其数次不能够更改。"
                            .数次 = dblPreTime: Bill.Text = dblPreTime
                            Exit Sub
                        End If
                        
                        
                        Call CalcMoneys(Bill.Row)
                        
                        '数据溢出检查(在已经算出该行费用但未显示前)
                        If MoneyOverFlow(mobjBill) Then
                            MsgBox "输入数量导致单据金额过大，请作适当调整。", vbInformation, gstrSysName
                            .数次 = dblPreTime
                            Bill.Text = ""
                            Call CalcMoneys(Bill.Row)
                            Cancel = True: Bill.TxtVisible = False: Exit Sub
                        End If
                        
                        '记帐分类报警(在已经算出该行费用但未显示前)
                        mrsWarn.Filter = ""
                        If mrsWarn.RecordCount > 0 Then
                            curTotal = CalcOneTotal(Bill.Row)
                            If curTotal > 0 Then
                                If .就诊卡号 = "" Then
                                    cur余额 = 0
                                    cur当日额 = 0
                                Else
                                    cur余额 = Val(Split(.就诊卡号, ",")(2))
                                    cur当日额 = Val(Split(.就诊卡号, ",")(3))
                                End If
                                If gbln报警包含划价费用 Then cur余额 = cur余额 - GetPriceMoneyTotal(1, .病人ID) + IIf(gbytBilling = 1, mcurModiMoney, 0)
                                
                                '刘兴洪:24491
                                curItemMoney = GetBillRowTotal(mobjBill.Details(Bill.Row).InComes)
                                
                                gbytWarn = BillingWarn(mstrPrivsOpt, .姓名 & IIf(.住院号 = "", "", "(住院号:" & .住院号 & " 床号:" & .床号 & ")"), .病区ID, GetPatiWarnRange(.病人ID, .主页ID), mrsWarn, _
                                        cur余额, cur当日额, curTotal, .担保额, .收费类别, .Detail.类别名称, mstrWarn, True, , curItemMoney)
                                If gbytWarn = 2 Or gbytWarn = 3 Then
                                    .数次 = dblPreTime
                                    Bill.Text = ""
                                    Call CalcMoneys(Bill.Row)
                                    Cancel = True: Bill.TxtVisible = False: Exit Sub
                                End If
                            End If
                        End If
                        
                        If int险类 <> 0 And mobjBill.Details(Bill.Row).数次 <> 0 Then
                            If gclsInsure.GetCapability(support实时监控, mobjBill.Details(Bill.Row).病人ID, int险类) Then
                                If gclsInsure.CheckItem(int险类, 1, 0, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 2, gbytBilling, Bill.Row)) = False Then
                                    .数次 = dblPreTime
                                    Call CalcMoneys(Bill.Row)
                                    Bill.Text = "": Bill.TxtVisible = False
                                    Cancel = True: Exit Sub
                                End If
                            End If
                        End If
                        
                        If .数次 <> 0 Then
                            If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModule, 1, 0, _
                                MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 2, gbytBilling, Bill.Row)) = False Then
                                .数次 = dblPreTime
                                Call CalcMoneys(Bill.Row)
                                Bill.Text = "": Bill.TxtVisible = False
                                Cancel = True: Exit Sub
                            End If
                        End If
                    End With
                        
                    Call ShowDetails(Bill.Row)
                    '更改其固有从属的数次
                    For i = Bill.Row + 1 To mobjBill.Details.Count
                        If mobjBill.Details(i).从属父号 = Bill.Row Then
                            '28136
                            '如果是输入的负数,需要将下级中的负数集中更新成负数
                            With mobjBill.Details(i)
                                If .Detail.固有从属 = 0 Then  '非固有从属
                                    If Abs(.数次) <> Abs(.Detail.从项数次) Then GoTo NotCalc:
                                    .数次 = IIf(Val(Bill.Text) < 0, -1, 1) * .Detail.从项数次
                                ElseIf .Detail.固有从属 = 1 Then '固定的固有从属
                                    .数次 = IIf(Val(Bill.Text) < 0, -1, 1) * IIf(.Detail.从项数次 = 0, 1, .Detail.从项数次)
                                ElseIf .Detail.固有从属 = 2 Then   '按比例的固有从属
                                    .数次 = Val(Bill.Text) * .Detail.从项数次
                                Else
                                     GoTo NotCalc:
                                End If
                            End With
                            
                            Call CalcMoneys(i)
                            Call ShowDetails(i)
NotCalc:
                        End If
                    Next
                    
                 
                ElseIf mobjBill.Details.Count >= Bill.Row Then
                    If Val(Bill.TextMatrix(Bill.Row, Bill.Col)) = 0 Then
                        If MsgBox("数量输入为零，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Cancel = True: Exit Sub
                        End If
                    End If
                End If
                If Bill.ColData(BillCol.执行科室) = BillColType.UnFocus Then
                    If CheckItemHaveSub(Bill.Row) Then
                        KeyCode = 0
                        Call LocateMainItemNextRow(Bill.Row)
                    End If
                End If
            Case "单价"
                If mobjBill.Details.Count >= Bill.Row And Bill.Text <> "" Then
                    '合法性检查
                    If Not IsNumeric(Bill.Text) Then
                        MsgBox "非法数值！", vbInformation, gstrSysName
                        Bill.Text = "": Cancel = True: Bill.TxtVisible = False: Exit Sub
                    End If
                    If Val(Bill.Text) < 0 Then
                        MsgBox "项目价格不应该为负数，要冲销费用，请输入负的数量来实现！", vbInformation, gstrSysName
                        Bill.Text = "": Cancel = True: Bill.TxtVisible = False: Exit Sub
                    End If
                    '最大金额检查
                    If gcurMaxMoney > 0 Then
                        If CSng(Bill.Text) * mobjBill.Details(Bill.Row).数次 * mobjBill.Details(Bill.Row).付数 > gcurMaxMoney Then
                            If MsgBox("当前金额超过了" & gcurMaxMoney & ",你确定要继续吗?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                                Bill.Text = "": Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                    Bill.Text = FormatEx(Bill.Text, 5)
                    
                    '如果没有对应的收入项目,则无法计算
                    If mobjBill.Details(Bill.Row).Detail.变价 And mobjBill.Details(Bill.Row).InComes.Count > 0 Then
                        If Not (mobjBill.Details(Bill.Row).InComes(1).现价 = 0 And mobjBill.Details(Bill.Row).InComes(1).原价 = 0) Then
                            strScope = CheckScope(mobjBill.Details(Bill.Row).InComes(1).原价, mobjBill.Details(Bill.Row).InComes(1).现价, CCur(Bill.Text))
                            If strScope <> "" Then
                                sta.Panels(2) = strScope
                                If Bill.TxtVisible And Len(Bill.Text) > 9 Then Bill.Text = mobjBill.Details(Bill.Row).InComes(1).标准单价
                                If Bill.TxtVisible Then Bill.SelStart = 0: Bill.SelLength = Len(Bill.Text)
                                Cancel = True: Beep: Exit Sub
                            End If
                        End If
                        
                        dblPreMoney = mobjBill.Details(Bill.Row).InComes(1).标准单价
                        
                        mobjBill.Details(Bill.Row).InComes(1).标准单价 = Bill.Text '这种收费细目只能对应一个收入项目
                        Call CalcMoneys(Bill.Row)
                        
                        '记帐分类报警(在已经算出该行费用但未显示前)
                        mrsWarn.Filter = ""
                        If mrsWarn.RecordCount > 0 Then
                            curTotal = CalcOneTotal(Bill.Row)
                            If curTotal > 0 Then
                                If mobjBill.Details(Bill.Row).就诊卡号 = "" Then
                                    cur余额 = 0
                                    cur当日额 = 0
                                Else
                                    cur余额 = Split(mobjBill.Details(Bill.Row).就诊卡号, ",")(2)
                                    cur当日额 = Split(mobjBill.Details(Bill.Row).就诊卡号, ",")(3)
                                End If
                                
                                If gbln报警包含划价费用 Then cur余额 = cur余额 - GetPriceMoneyTotal(1, mobjBill.Details(Bill.Row).病人ID) + IIf(gbytBilling = 1, mcurModiMoney, 0)
                                gbytWarn = BillingWarn(mstrPrivsOpt, mobjBill.Details(Bill.Row).姓名 & IIf(mobjBill.Details(Bill.Row).住院号 = "", "", "(住院号:" & mobjBill.Details(Bill.Row).住院号 & " 床号:" & mobjBill.Details(Bill.Row).床号 & ")"), mobjBill.Details(Bill.Row).病区ID, GetPatiWarnRange(mobjBill.Details(Bill.Row).病人ID, mobjBill.Details(Bill.Row).主页ID), mrsWarn, _
                                        cur余额, cur当日额, curTotal, mobjBill.Details(Bill.Row).担保额, mobjBill.Details(Bill.Row).收费类别, mobjBill.Details(Bill.Row).Detail.类别名称, mstrWarn, True)
                                If gbytWarn = 2 Or gbytWarn = 3 Then
                                    mobjBill.Details(Bill.Row).InComes(1).标准单价 = dblPreMoney
                                    Bill.Text = ""
                                    Call CalcMoneys(Bill.Row)
                                    Cancel = True: Bill.TxtVisible = False: Exit Sub
                                End If
                            End If
                        End If
                        Call ShowDetails(Bill.Row)
                    Else
                        Bill.Text = "0"
                        sta.Panels(2) = "该项目设有设置对应的费目，所以无法计算费用！"
                        Beep
                    End If
                End If
            Case "执行科室"
                If mobjBill.Details.Count >= Bill.Row And Bill.ListIndex <> -1 Then
                    With mobjBill.Details(Bill.Row)
                        If .执行部门ID <> Bill.ItemData(Bill.ListIndex) Then
                            .执行部门ID = Bill.ItemData(Bill.ListIndex)
                             If CheckItemHaveSub(Bill.Row) Then Call SetSubItemDept(Bill.Row) '如果存在从项,则改变非药品行的执行科室
                        End If
                
                        '药品库存检查:动态药房,分批或时价药品也要检查了
                        If (.收费类别 = "4" And .Detail.跟踪在用) Or (InStr(",5,6,7,", .收费类别) > 0 And Not gbln分离发药) Then
                            If .Detail.分批 Or .Detail.变价 Then '分批或时价药品库存不足禁止输入
                                If .付数 * .数次 > .Detail.库存 Then
                                    If .收费类别 = "4" Then
                                        MsgBox "[" & .Detail.名称 & "]为分批或时价卫生材料,当前可用库存不足输入数量！", vbInformation, gstrSysName
                                    Else
                                        MsgBox "[" & .Detail.名称 & "]为分批或时价药品,当前可用库存不足输入数量！", vbInformation, gstrSysName
                                    End If
                                    Cancel = True
                                End If
                            Else
                                Set colStock = IIf(.收费类别 = "4", mcolStock2, mcolStock1)
                                If colStock("_" & .执行部门ID) <> 0 Then
                                    If .付数 * .数次 > .Detail.库存 Then
                                        If colStock("_" & .执行部门ID) = 1 Then
                                            If MsgBox("[" & .Detail.名称 & "]的当前可用库存不足输入数量,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                                Cancel = True
                                            End If
                                        ElseIf colStock("_" & .执行部门ID) = 2 Then
                                            MsgBox "[" & .Detail.名称 & "]的当前可用库存不足输入数量！", vbInformation, gstrSysName
                                            Cancel = True
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        
                        '检查卫生材料的灭菌效期,在确定执行科室之后
                        If .收费类别 = "4" And .Detail.跟踪在用 Then
                            Call CheckValidity(.收费细目ID, .执行部门ID, .数次, False) '已确认输入,仅能提醒
                        End If
                        
                        If CheckItemHaveSub(Bill.Row) Then
                            KeyCode = 0
                            Call LocateMainItemNextRow(Bill.Row)
                        End If
                        If int险类 <> 0 And mobjBill.Details(Bill.Row).数次 <> 0 Then
                            If gclsInsure.GetCapability(support实时监控, mobjBill.Details(Bill.Row).病人ID, int险类) Then
                                If gclsInsure.CheckItem(int险类, 1, 0, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 2, gbytBilling, Bill.Row)) = False Then
                                    Bill.Text = "": Bill.TxtVisible = False
                                    Cancel = True: Exit Sub
                                End If
                            End If
                        End If
                        
                        If mobjBill.Details(Bill.Row).数次 <> 0 Then
                            If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModule, 1, 0, _
                                MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 2, gbytBilling, Bill.Row)) = False Then
                                Bill.Text = "": Bill.TxtVisible = False
                                Cancel = True: Exit Sub
                            End If
                        End If
                        
                    End With
                End If
        End Select
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Cancel = True
End Sub


Private Sub LocateMainItemNextRow(ByVal lngRow As Long)
    Dim i As Long
    
    For i = lngRow + 1 To mobjBill.Details.Count
        If mobjBill.Details(i).从属父号 = lngRow Then
            If mobjBill.Details(i).Detail.固有从属 = 0 Then Exit For
        End If
    Next
    
    If i <= mobjBill.Details.Count Then
        Bill.Col = BillCol.数次
        Bill.Row = i: Bill.MsfObj.TopRow = i
    Else
        Call LocateNewRow
    End If
End Sub

Private Sub LocateNewRow()
    If mobjBill.Details.Count >= Bill.Rows - 1 Then
        Bill.Rows = Bill.Rows + 1
        Call bill_AfterAddRow(Bill.Rows - 1)
        Bill.Row = Bill.Rows - 1
        Bill.MsfObj.TopRow = Bill.Row
        Bill.Col = BillCol.性别
    Else
        Bill.Row = Bill.Rows - 1
        Bill.MsfObj.TopRow = Bill.Row
        Bill.Col = BillCol.性别
    End If
    '问题:27792
    If Not Me.ActiveControl Is Bill Then
        If Bill.Active And Bill.Visible Then Bill.SetFocus
    End If
End Sub

Private Sub SetSubItem(lng病人科室ID As Long, int病人来源 As Integer)
'功能:输入收费项目后,加载当前收费项目的从属项目到费用集对象,并显示在单据控件中
'参数:
'调用者:Bill_KeyDown中输入项目后
Dim i As Integer, j As Integer, lngMainRow As Long
Dim lngDoUnit As Long
Dim bln从项汇总折扣 As Boolean
Dim str摘要 As String, strPriceGrade As String

lngMainRow = Bill.Row               '主项的行
If gbln从项汇总折扣 Then            '如果主项屏蔽费别,则汇总计算折扣参数无效,不汇总计算
    bln从项汇总折扣 = Not mobjBill.Details(lngMainRow).Detail.屏蔽费别
End If

With mobjBill.Details(lngMainRow)
    Set mcolDetails = New Details
    Set mcolDetails = GetSubDetails(.收费细目ID)
    For i = 1 To mcolDetails.Count
        If mobjBill.Details.Count >= Bill.Rows - 1 Then
            Bill.Rows = Bill.Rows + 1
            Call bill_AfterAddRow(Bill.Rows - 1)
        End If
        Bill.TextMatrix(Bill.Rows - 1, BillCol.费别) = "" '有必要加上
        
        'a.从属项目为非药品项目的执行科室
        lngDoUnit = 0
        If InStr(",4,5,6,7,", mcolDetails(i).类别) = 0 Then
             If mcolDetails(i).类别 = .收费类别 Or mcolDetails(i).执行科室 = 0 Then
                '1.从项收费类别与主项相同的,缺省与主项执行科室相同。
                '2.从项设置为无明确科室的,缺省与主项执行科室相同。
                lngDoUnit = .执行部门ID
             Else
                '3.其它非药项目的执行科室
                lngDoUnit = Get收费执行科室ID(mcolDetails(i).类别, mcolDetails(i).ID, _
                    mcolDetails(i).执行科室, lng病人科室ID, Get开单科室ID, int病人来源, , .病区ID)
             End If
        'b.从属项目为药品,卫材的执行科室
        Else
            lngDoUnit = Get收费执行科室ID(mcolDetails(i).类别, mcolDetails(i).ID, _
                mcolDetails(i).执行科室, lng病人科室ID, Get开单科室ID, int病人来源, .执行部门ID, .病区ID) '卫材从项缺省与主项执行科室相同
        End If
        
        
        '保险支付项目对应检查
        If Val(mobjBill.Details(lngMainRow).发药窗口) > 0 Then
            If InStr(",5,6,7,", mcolDetails(i).类别) > 0 Then
                strPriceGrade = mstr药品价格等级
            ElseIf mcolDetails(i).类别 = "4" Then
                strPriceGrade = mstr卫材价格等级
            Else
                strPriceGrade = mstr普通价格等级
            End If
            If Not CheckMediCareItem(mcolDetails(i).ID, Val(mobjBill.Details(lngMainRow).发药窗口), mcolDetails(i).名称, _
                mcolDetails(i).变价 = False, , strPriceGrade) Then
                Exit Sub
            End If
        End If
        
        Call SetDetailtStock(lngDoUnit, mcolDetails(i))
        Call SetDetail(mcolDetails(i), Bill.Rows - 1, lngDoUnit, Bill.Row)
        
        Call CalcMoney(Bill.Rows - 1, bln从项汇总折扣)
        Call ShowDetails(Bill.Rows - 1)
        
        
        'CalcMoney中先调用GetuItemInsure可能返回摘要
         str摘要 = mobjBill.Details(Bill.Rows - 1).摘要
         str摘要 = gclsInsure.GetItemInfo(Val(mobjBill.Details(lngMainRow).发药窗口), mobjBill.Details(lngMainRow).病人ID, mcolDetails(i).ID, str摘要, 2)
         mobjBill.Details(Bill.Rows - 1).摘要 = str摘要
        
    Next
    
    If bln从项汇总折扣 Then
        Call CalcMoney(lngMainRow, bln从项汇总折扣) '先重算主项的应收与实收,因为在没有加入从项前可能是按单独打折算的.
        
        Call Calc重算主项实收(lngMainRow)
    End If
End With

End Sub

Private Sub Calc重算主项实收(ByVal lngMainRow As Long)
'功能:当从项汇总折扣时,根据指定的主项的行ID的第一个收入项目重算主项的实收金额
'参数:  lngMainRow-主项行ID

Dim i As Long, j As Long
Dim cur打折前应收合计 As Currency     '记录所有主从项的应收合计
Dim cur打折后实收 As Currency


With mobjBill
    For i = lngMainRow To .Details.Count
        'If i <> lngMainRow And .Details(i).从属父号 <> lngMainRow Then Exit For    '虽然目前限制了不允许在从项中间插入别的主从项,但因一张单据行数不多,为了将来可能的需求,还是全部扫描
        If i = lngMainRow Or .Details(i).从属父号 = lngMainRow Then
            For j = 1 To .Details(i).InComes.Count
                cur打折前应收合计 = cur打折前应收合计 + .Details(i).InComes(j).应收金额
            Next
        End If
    Next
       
    '仅收费项目涉及主从项汇总折扣,所以药品成本加收的参数无需传入
    cur打折后实收 = CCur(Format(ActualMoney(.Details(lngMainRow).费别, .Details(lngMainRow).InComes(1).收入项目ID, cur打折前应收合计, 0, 0, 0, 0), gstrDec))
    cur打折后实收 = cur打折后实收 - cur打折前应收合计 + .Details(lngMainRow).InComes(1).应收金额
    .Details(lngMainRow).InComes(1).实收金额 = Format(cur打折后实收, gstrDec)
    
    Call ShowDetails(lngMainRow)
End With
End Sub

Private Sub SetSubItemDept(ByVal lngRow As Long)
'功能:根据主项执行科室的变化,刷新非药从项的执行科室

    Dim i As Long, j As Long, lng病人科室ID As Long
    
    With mobjBill
        '获取所有从项及其执行科室类型,必须现取(因为界面上的从项信息可能是修改过的)
        Set mcolDetails = GetSubDetails(.Details(lngRow).收费细目ID)
        
        lng病人科室ID = .Details(lngRow).科室ID
        If lng病人科室ID = 0 And cbo开单科室.ListIndex <> -1 Then lng病人科室ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
        
        For i = lngRow + 1 To .Details.Count
            If .Details(i).从属父号 = lngRow Then
                '从属项为药品和卫材的项目的执行科室不随主项变动
                If InStr(",4,5,6,7,", .Details(i).收费类别) = 0 Then
                    If .Details(i).收费类别 = .Details(lngRow).收费类别 Then
                        '1.从项收费类别与主项相同的,缺省与主项执行科室相同。
                        .Details(i).执行部门ID = .Details(lngRow).执行部门ID
                    Else
                        For j = 1 To mcolDetails.Count
                            If mcolDetails.Item(j).ID = .Details(i).Detail.ID Then
                                Exit For
                            End If
                        Next
                        If j <= mcolDetails.Count Then
                            If mcolDetails.Item(j).执行科室 = 0 Then
                                '2.从项设置为无明确科室的,缺省与主项执行科室相同。
                                 .Details(i).执行部门ID = .Details(lngRow).执行部门ID
                            Else
                                '3.其它非药项目的执行科室
                                .Details(i).执行部门ID = Get收费执行科室ID(mcolDetails(j).类别, mcolDetails(j).ID, _
                                    mcolDetails(j).执行科室, lng病人科室ID, Get开单科室ID, Get病人来源(lngRow), , .Details(i).病区ID)
                            End If
                        End If
                    End If
                    
                    '刷新显示从项执行科室
                    If .Details(i).执行部门ID <> 0 Then
                        mrsUnit.Filter = "ID=" & .Details(i).执行部门ID
                        If mrsUnit.RecordCount <> 0 Then
                            If mbytInState = 0 Then
                                Bill.TextMatrix(i, BillCol.执行科室) = mrsUnit!编码 & "-" & mrsUnit!名称
                            Else
                                '浏览单据只(能)显示名称
                                Bill.TextMatrix(i, BillCol.执行科室) = mrsUnit!名称
                            End If
                        Else
                            Bill.TextMatrix(i, BillCol.执行科室) = GET部门名称(.Details(i).执行部门ID, mrsUnit)
                        End If
                    Else
                        Bill.TextMatrix(i, BillCol.执行科室) = ""
                    End If
                    
                End If
            End If
        Next
    End With
End Sub

Private Function CheckItemHaveSub(ByVal lngRow As Long) As Boolean
'功能：判断当前行的项目是否具有从属项目
    Dim i As Long
    
    If mobjBill.Details.Count >= lngRow Then
        For i = lngRow + 1 To mobjBill.Details.Count
            If mobjBill.Details(i).从属父号 = lngRow Then
                CheckItemHaveSub = True: Exit Function
            End If
        Next
    End If
End Function

Private Sub Bill_EnterCell(Row As Long, Col As Long)
    Dim i As Long, curTotal As Currency
    Dim arrMoney As Variant, strMoney As String, arrPatiInfo As Variant
    Dim rsTmp As ADODB.Recordset, strStock As String
    Dim lng病人ID As Long, str药房IDs As String
        
    If Row = 0 Then Exit Sub
    If Bill.ColData(Col) = BillColType.UnFocus Then Exit Sub
    
    If Not mblnEnterCell Then Exit Sub
    
    
    If mlngPreRow <> Row Then
        '计算当前病人费用
         curTotal = CalcOneTotal(Row)
    End If
    
    If mobjBill.Details.Count = 0 And mcolPatiInfo.Count > 0 Then
        '显示婴儿费
        arrPatiInfo = Split(mcolPatiInfo("R" & Row), ",")
        Call LoadPatientBaby(cboBaby, Val(arrPatiInfo(0)), Val(arrPatiInfo(1)))
        Call zlControl.CboLocate(cboBaby, arrPatiInfo(2), True)
    End If
    
    If Not Bill.Active Or mstrInNO <> "" Then
        If mobjBill.Details.Count = 0 And mcolPatiInfo.Count > 0 Then  '读入时(如审核)显示病人费用
            If Val(mcolPatiInfo("R" & Row)) <> 0 And mlngPreRow <> Row Then
                Set rsTmp = GetMoneyInfo(Val(mcolPatiInfo("R" & Row)), , True, 2)
                If Not rsTmp Is Nothing Then
                    'sta.Panels(3).Text = Bill.TextMatrix(Row, BillCol.姓名) & "预交:" & Format(rsTmp!预交余额, "0.00")
                    'sta.Panels(3).Text = sta.Panels(3) & "/费用:" & Format(rsTmp!费用余额, gstrDec)
                    'sta.Panels(3).Text = sta.Panels(3) & "/余额:" & Format(rsTmp!预交余额 - rsTmp!费用余额, "0.00")
                    '30604
                    Call SetStatuPatiInfor(Bill.TextMatrix(Row, BillCol.姓名), Val(Nvl(rsTmp!预交余额)), _
                        Val(Nvl(rsTmp!费用余额)), Val(Nvl(rsTmp!预交余额)) - Val(Nvl(rsTmp!费用余额)))
                End If
            End If
        End If
        If Not Bill.Active Then mlngPreRow = Row: Exit Sub
    End If
    
    If mbytInState = 3 Or (chkCancel.Visible And chkCancel.Value = 1) Then
        '针对列编辑性质设置颜色
        mlngPreRow = Row
        Exit Sub
    End If
    
    '--------------------------------------------------------------------------
    '1.行改变的相关数据处理和设置     mlngPreRow    当前行是否改变
    If mobjBill.Details.Count >= Bill.Row And mlngPreRow <> Row Then
        With mobjBill.Details(Bill.Row)
            '显示库存
            If InStr(",5,6,7,", .收费类别) > 0 And .收费细目ID <> 0 Then
                If Not gbln分离发药 Then
                    If gbln其它药房 Or gbln其它药库 Then
                        strStock = GetStockInfo(.收费细目ID, gbln其它药房, gbln其它药库, gbln住院单位)
                        If strStock <> "" Then
                            If InStr(1, mstrPrivsOpt, ";显示库存;") > 0 Then
                                sta.Panels(Pan.C2提示信息) = "第" & Bill.Row & "行库存:" & strStock
                            Else
                                sta.Panels(Pan.C2提示信息) = "第" & Bill.Row & "行有库存."
                            End If
                        End If
                    End If
                    If strStock = "" Then
                        '随时更新库存显示
                        .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                        If gbln住院单位 Then .Detail.库存 = .Detail.库存 / .Detail.住院包装
                        
                        Call ShowStock(.Detail.名称, .Detail.库存)
                    End If
                Else
                    str药房IDs = Decode(.收费类别, "5", gstr西药房, "6", gstr成药房, "7", gstr中药房)
                    If str药房IDs <> "" Then
                        .Detail.库存 = GetMultiStock(.收费细目ID, str药房IDs)
                        If gbln住院单位 Then .Detail.库存 = .Detail.库存 / .Detail.住院包装
                        
                        Call ShowStock(.Detail.名称, .Detail.库存)
                    Else
                        sta.Panels(2) = ""
                    End If
                End If
            ElseIf .收费类别 = "4" And .Detail.跟踪在用 And .收费细目ID <> 0 Then
                .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                Call ShowStock(.Detail.名称, .Detail.库存)
            Else
                sta.Panels(2) = ""
            End If
            
            If mobjBill.Details.Count >= Row Then
                Call LoadPatientBaby(cboBaby, .病人ID, .主页ID)
                Call zlControl.CboLocate(cboBaby, .婴儿费, True)
            End If
            
            Bill.ColData(BillCol.姓名) = BillColType.Text
            Bill.ColData(BillCol.类别) = IIf(gbln收费类别 And Not mblnOne, BillColType.ComboBox, BillColType.UnFocus)
            Bill.ColData(BillCol.项目) = BillColType.CommandButton
            
            '如果是从属项目的主项目或从项,则不允许更改类别和项目
            If CheckItemHaveSub(Row) Or .从属父号 > 0 Then
                Bill.ColData(BillCol.姓名) = BillColType.Text_UnModify
                Bill.ColData(BillCol.类别) = BillColType.Text_UnModify
                Bill.ColData(BillCol.项目) = BillColType.Text_UnModify
            End If
            
            '如果是非调整状态
            If mbytInState <> 2 Then
                If .收费类别 = "7" And gblnPay Then
                    Bill.ColData(BillCol.付数) = BillColType.Text
                Else
                    Bill.ColData(BillCol.付数) = BillColType.UnFocus
                End If
                
                '变价允许输入数次
                If .Detail.变价 And InStr(",5,6,7,", .收费类别) = 0 _
                    And Not (.收费类别 = "4" And .Detail.跟踪在用) Then
                    Bill.ColData(BillCol.数次) = IIf(gblnTime, BillColType.Text, BillColType.UnFocus) '数次
                    Bill.ColData(BillCol.单价) = BillColType.Text '单价
                Else
                    Bill.ColData(BillCol.数次) = BillColType.Text '数次
                    Bill.ColData(BillCol.单价) = BillColType.UnFocus '单价
                End If
                
                If .收费类别 = "F" Then
                    Bill.ColData(BillCol.标志) = BillColType.CheckBox
                Else
                    Bill.ColData(BillCol.标志) = BillColType.UnFocus
                End If
                
                If .Key = "1" Then    '指定了固定药房时,不允许再选择执行科室
                    Bill.ColData(BillCol.执行科室) = BillColType.UnFocus
                Else
                    Bill.ColData(BillCol.执行科室) = BillColType.ComboBox
                End If
                
                If mblnOne Then Bill.ColData(BillCol.类别) = BillColType.UnFocus '只允许一个类别
            End If
            
            '修改时动态读取病人费用信息,以加快速度
            If mstrInNO <> "" And .就诊卡号 = "" Then
                '修改前的当前单据的病人费用金额
                mcurModiMoney = GetBillMoney(2, mstrInNO, .病人ID)
                
                '病人剩余款(利用就诊卡号域)
                Set rsTmp = Nothing
                Set rsTmp = GetMoneyInfo(.病人ID, IIf(gbytBilling = 0, mcurModiMoney, 0), True, 2)
                If Not rsTmp Is Nothing Then
                   'sta.Panels(3).Text = .姓名 & "预交:" & Format(rsTmp!预交余额, "0.00")
                   ' sta.Panels(3).Text = sta.Panels(3) & "/费用:" & Format(rsTmp!费用余额, gstrDec)
                   ' sta.Panels(3).Text = sta.Panels(3) & "/余额:" & Format(rsTmp!预交余额 - rsTmp!费用余额, "0.00")
                    
                    Call SetStatuPatiInfor(.姓名, Val(Nvl(rsTmp!预交余额)), _
                        Val(Nvl(rsTmp!费用余额)), Val(Nvl(rsTmp!预交余额)) - Val(Nvl(rsTmp!费用余额)))
                    
                    
                    .就诊卡号 = rsTmp!预交余额 & "," & rsTmp!费用余额 & "," & rsTmp!预交余额 - rsTmp!费用余额
                Else
                    .就诊卡号 = "0,0,0"
                End If
                '病人当日费用额
                .就诊卡号 = .就诊卡号 & "," & GetPatiDayMoney(.病人ID) - mcurModiMoney
            End If
        End With
    End If
    
        
    '计算病人实际剩余款
    If mlngPreRow <> Row Then
        If Bill.TextMatrix(Row, BillCol.姓名) = "" Then
            sta.Panels(3).Text = "": picStatuPancl.Visible = False: lblStatuPati.Caption = ""
        Else
            If mobjBill.Details.Count >= Row Then
                If InStr(mobjBill.Details(Row).就诊卡号, ",") > 0 Then
                    arrMoney = Split(mobjBill.Details(Row).就诊卡号, ",")
                    'sta.Panels(3).Text = mobjBill.Details(Row).姓名 & "预交:" & Format(Val(arrMoney(0)), "0.00")
                    'sta.Panels(3).Text = sta.Panels(3) & "/费用:" & Format(Val(arrMoney(1)) + IIf(gbytBilling = 0, curTotal, 0), gstrDec)
                    'sta.Panels(3).Text = sta.Panels(3) & "/余额:" & Format(Val(arrMoney(2)) - IIf(gbytBilling = 0, curTotal, 0), "0.00")
                    Call SetStatuPatiInfor(mobjBill.Details(Row).姓名, Val(arrMoney(0)), _
                                     Val(arrMoney(1)) + IIf(gbytBilling = 0, curTotal, 0), Val(arrMoney(2)) - IIf(gbytBilling = 0, curTotal, 0))
                                          
                End If
            Else
                '新的未保存的行
                If mrsInfo.State = 1 Then
                    lng病人ID = mrsInfo!病人ID
                ElseIf mobjBill.Details.Count < Row And mobjBill.Details.Count >= Row - 1 And Row > 1 Then
                    lng病人ID = mobjBill.Details(Row - 1).病人ID
                End If
                If lng病人ID > 0 Then
                    strMoney = GetMoneyStr(lng病人ID)
                    If InStr(strMoney, ",") > 0 Then
                        arrMoney = Split(strMoney, ",")
                        'sta.Panels(3).Text = Bill.TextMatrix(Row, BillCol.姓名) & "预交:" & Format(Val(arrMoney(0)), "0.00")
                        'sta.Panels(3).Text = sta.Panels(3) & "/费用:" & Format(Val(arrMoney(1)) + IIf(gbytBilling = 0, curTotal, 0), gstrDec)
                        'sta.Panels(3).Text = sta.Panels(3) & "/余额:" & Format(Val(arrMoney(2)) - IIf(gbytBilling = 0, curTotal, 0), "0.00")
                        Call SetStatuPatiInfor(Bill.TextMatrix(Row, BillCol.姓名), Val(arrMoney(0)), _
                                        Val(arrMoney(1)) + IIf(gbytBilling = 0, curTotal, 0), Val(arrMoney(2)) - IIf(gbytBilling = 0, curTotal, 0))
                    End If
                End If
            End If
        End If
        
        '如果点击未保存的行,则恢复列的性质
        If mobjBill.Details.Count < Bill.Row Then
            Bill.ColData(BillCol.姓名) = BillColType.Text
            Bill.ColData(BillCol.类别) = IIf(gbln收费类别 And Not mblnOne, BillColType.ComboBox, BillColType.UnFocus)
            Bill.ColData(BillCol.项目) = BillColType.CommandButton
        End If
    End If
     
    
    
    '-----------------------------------------------------------------
    '2.列改变的相关数据处理和显示设置
    If Bill.ColData(Bill.Col) = BillColType.ComboBox Then
        Call FillBillComboBox(Bill.Row, Bill.Col, True) '进入该列
    End If
    
    If gbln收费类别 And Bill.TextMatrix(Row, BillCol.类别) = "" And mblnOne Then
        mrsClass.Filter = "编码=" & gstr收费类别
        Bill.TextMatrix(Row, BillCol.类别) = mrsClass!类别
        Bill.RowData(Row) = Asc(mrsClass!编码)
    End If
    
    Bill.TextLen = 0: Bill.TextMask = ""
    Select Case Bill.TextMatrix(0, Col)
        Case "病人"
            Bill.TextLen = 19
        Case "费别"
            SetWidth Bill.cboHwnd, 70
            For i = 0 To Bill.ListCount - 1
                If InStr(Bill.List(i), Bill.TextMatrix(Row, Col)) > 0 Then
                    Bill.ListIndex = i: Exit For
                End If
            Next
        Case "类别" '不输入类别时不会定位到类别列
            SetWidth Bill.cboHwnd, 65
            '类别如果为空,则自动默认为上一收费细目的类别
            If Bill.TextMatrix(Row, Col) = "" Then
                If mblnOne Then
                    mrsClass.Filter = "编码=" & gstr收费类别
                    Bill.TextMatrix(Row, Col) = mrsClass!类别
                    Bill.RowData(Row) = Asc(mrsClass!编码)
                ElseIf Row > 1 Then
                    Bill.ListIndex = GetBillIndex(Bill.TextMatrix(Row - 1, Col))
                End If
            ElseIf Row >= 1 And Bill.TextMatrix(Row, Col) <> "" Then
                For i = 0 To Bill.ListCount - 1
                    If InStr(Bill.List(i), Bill.TextMatrix(Row, Col)) > 0 Then
                        Bill.ListIndex = i: Exit For
                    End If
                Next
                If Bill.ListIndex = -1 Then
                    Bill.ListIndex = GetBillIndex(Bill.TextMatrix(Row - 1, Col))
                End If
            End If
        Case "执行科室"
            SetWidth Bill.cboHwnd, 110
        Case "付"
            Bill.TextLen = 3: Bill.TextMask = "0123456789" & Chr(8)
        Case "数次"
            Bill.TextLen = 8: Bill.TextMask = "0123456789." & Chr(8)
            If mobjBill.Details.Count >= Bill.Row Then
                If InStr(",5,6,7,", mobjBill.Details(Bill.Row).收费类别) > 0 Then
                    If InStr(mstrPrivsOpt, ";药品输入小数;") = 0 Then
                        Bill.TextMask = Replace(Bill.TextMask, ".", "")
                    End If
                End If
                '可否输入负数
                If Not mobjBill.Details(Bill.Row).Detail.分批 Then
                    If InStr(",5,6,", mobjBill.Details(Bill.Row).收费类别) > 0 Then
                        If InStr(mstrPrivsOpt, ";成药负数记帐;") > 0 Then Bill.TextMask = "-" & Bill.TextMask
                    ElseIf InStr(",7,", mobjBill.Details(Bill.Row).收费类别) > 0 Then
                        If InStr(mstrPrivsOpt, ";草药负数记帐;") > 0 Then Bill.TextMask = "-" & Bill.TextMask
                    Else
                        If InStr(mstrPrivsOpt, ";诊疗负数记帐;") > 0 Then Bill.TextMask = "-" & Bill.TextMask
                    End If
                    
                    If InStr(Bill.TextMask, "-") > 0 And mrsInfo.State = 1 Then
                        If Not IsNull(mrsInfo!险类) Then
                            If Not gclsInsure.GetCapability(support负数记帐, mrsInfo!病人ID, mrsInfo!险类) Then
                                Bill.TextMask = Replace(Bill.TextMask, "-", "")
                            End If
                        End If
                    End If
                End If
            End If
        Case "单价"
            Bill.TextLen = 10: Bill.TextMask = "0123456789." & Chr(8)
    End Select
    If Bill.MsfObj.ColIsVisible(Bill.Col) = False Then
        Bill.MsfObj.LeftCol = Bill.Col
    End If
    '新行,或更改已有行的类别时,看作换行还没有开始
    If Bill.TextMatrix(Row, BillCol.项目) = "" Then
        mlngPreRow = 0
    ElseIf mobjBill.Details.Count >= Row Then
        mlngPreRow = Row
    End If
End Sub

Private Sub Bill_KeyPress(KeyAscii As Integer)
    If Bill.TextMatrix(0, Bill.Col) = "病人" And Bill.Active And Bill.ColData(Bill.Col) <> BillColType.Text_UnModify Then
         If cbo开单科室.ListIndex <> -1 And KeyAscii = Asc("*") Then
            KeyAscii = 0
            Call FillPatient(cbo开单科室.ItemData(cbo开单科室.ListIndex))
            If Bill.Top + Bill.CellTop + lvwPati.Height > sta.Top Then
                lvwPati.Top = Bill.Top + Bill.CellTop - lvwPati.Height - 30
            Else
                lvwPati.Top = Bill.Top + Bill.CellTop + Bill.RowHeight(1) - 15
            End If
            lvwPati.Visible = True
            lvwPati.SetFocus
        End If
    End If
End Sub

Private Sub Bill_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Bill.ToolTipText = Bill.TextMatrix(Bill.MouseRow, Bill.MouseCol)
End Sub

Private Sub cboBaby_Click()
    Dim i As Long, lngParent As Long
    
    If Bill.Row <= mobjBill.Details.Count Then
        mobjBill.Details(Bill.Row).婴儿费 = cboBaby.ItemData(cboBaby.ListIndex)
        For i = Bill.Row + 1 To mobjBill.Details.Count
            If mobjBill.Details(i).从属父号 = Bill.Row Then
                mobjBill.Details(i).婴儿费 = mobjBill.Details(Bill.Row).婴儿费
            End If
        Next
    End If
End Sub

Private Sub cboBaby_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub SetDefaultDoctor()
'功能:设置缺省开单人
    If cbo开单人.ListCount = 0 Then Exit Sub
    
    If cbo开单人.ListCount = 1 Then
        cbo开单人.ListIndex = 0
    Else
        If mrsInfo.State = 1 Then
            If Not IsNull(mrsInfo!住院医师) Then
                Call zlControl.CboSetIndex(cbo开单人.hWnd, cbo.FindIndex(cbo开单人, mrsInfo!住院医师, True))
            End If
        End If
    End If
End Sub

Private Sub cboDrawDept_Click()
    Dim lng领药部门ID As Long
    If Not (mbytInState = 0 And chkCancel.Value = 0) Then Exit Sub
    If cboDrawDept.ListIndex <> -1 Then lng领药部门ID = cboDrawDept.ItemData(cboDrawDept.ListIndex)
    If mobjBill.领药部门ID = lng领药部门ID Then Exit Sub
    mobjBill.领药部门ID = lng领药部门ID
End Sub

Private Sub cboDrawDept_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii >= 32 And Not cboDrawDept.Locked Then
        lngIdx = zlControl.CboMatchIndex(cboDrawDept.hWnd, KeyAscii)
        If lngIdx = -1 And cboDrawDept.ListCount > 0 Then lngIdx = 0
        cboDrawDept.ListIndex = lngIdx
    ElseIf KeyAscii = 13 Then
        If cboDrawDept.ListIndex = -1 Then
            Beep
        Else
            mobjBill.领药部门ID = cboDrawDept.ItemData(cboDrawDept.ListIndex)
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub cbo开单科室_Click()
    Dim i As Long, lng开单部门ID As Long
    
    If Not (mbytInState = 0 And chkCancel.Value = 0) Then Exit Sub
    If cbo开单科室.ListIndex <> -1 Then lng开单部门ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
    If mobjBill.开单部门ID = lng开单部门ID Then Exit Sub
    mobjBill.开单部门ID = lng开单部门ID
    
    '问题:
    If mrs领药部门.RecordCount <> 0 Then
        For i = 0 To cboDrawDept.ListCount - 1
             If cboDrawDept.ItemData(i) = lng开单部门ID Then
                mobjBill.领药部门ID = lng开单部门ID
                cboDrawDept.ListIndex = i: Exit For
             End If
        Next
    End If
    
    '开单科室确定医生
    If Not gblnFromDr Then
        If cbo开单科室.ListIndex <> -1 Then
            If gbln它科人 Then
                Call FillDoctor(cbo开单人, mrs开单人)
            Else
                Call FillDoctor(cbo开单人, mrs开单人, lng开单部门ID)
            End If
            Call SetDefaultDoctor
        Else
            cbo开单人.Clear
        End If
        Call cbo开单人_Click
    End If
        
    '重新设置相关项目的执行科室(修改和查看时保持原样)
    If cbo开单科室.ListIndex <> -1 And cbo开单科室.Visible Then
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
               '仅处理收费项目
                If InStr(",4,5,6,7,", .Detail.类别) = 0 And .Detail.执行科室 = 6 Then '6-开单人科室
                    .执行部门ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
                    '刷新显示从项执行科室
                    If i <= Bill.Rows - 1 And .执行部门ID <> 0 Then
                        If mbytInState = 0 Then
                            mrsUnit.Filter = "ID=" & .执行部门ID
                            If mrsUnit.RecordCount <> 0 Then
                                Bill.TextMatrix(i, BillCol.执行科室) = mrsUnit!编码 & "-" & mrsUnit!名称
                            Else
                                Bill.TextMatrix(i, BillCol.执行科室) = GET部门名称(.执行部门ID, mrsUnit)
                            End If
                        Else
                            '浏览单据只(能)显示名称
                            Bill.TextMatrix(i, BillCol.执行科室) = GET部门名称(.执行部门ID, mrsUnit)
                        End If
                    Else
                        Bill.TextMatrix(i, BillCol.执行科室) = ""
                    End If
                End If
            End With
        Next
    End If
    
    If Not gblnFromDr Then '当不是由开单人定科室时
        If mobjBill.Details.Count > 0 And cbo开单科室.ListIndex <> mlngPreUnit And Visible And Not txtIn.Enabled Then
            MsgBox "提醒：你已经更改开单科室为""" & zlStr.NeedName(cbo开单科室.Text) & """,注意检查单据中的病人是否属于该科室！", vbInformation, gstrSysName
        End If
    End If
    
    mlngPreUnit = cbo开单科室.ListIndex
End Sub

Private Sub cbo开单科室_Validate(Cancel As Boolean)
    If cbo开单科室.Text <> "" And cbo开单科室.ListIndex < 0 Then cbo开单科室.Text = ""
End Sub

Private Sub cbo开单人_Click()
    Dim lng开单人ID As Long, lng开单科室ID As Long
    
    If Not (mbytInState = 0 And chkCancel.Value = 0) Then Exit Sub
    If mobjBill.开单人 = IIf(cbo开单人.ListIndex = -1, "", zlStr.NeedName(cbo开单人.Text)) Then Exit Sub
    
    mobjBill.开单人 = IIf(cbo开单人.ListIndex = -1, "", zlStr.NeedName(cbo开单人.Text))
    If gblnFromDr Then
        If cbo开单人.ListIndex <> -1 Then
            lng开单人ID = cbo开单人.ItemData(cbo开单人.ListIndex)
            If cbo开单科室.ListIndex <> -1 Then lng开单科室ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
            
            Call FillDept(cbo开单科室, mrs开单科室, mrs开单人, mstrPrivs, mbytUseType, mlngDeptID, lng开单人ID)
            If lng开单科室ID > 0 Then Call zlControl.CboSetIndex(cbo开单科室.hWnd, cbo.FindIndex(cbo开单科室, lng开单科室ID))
            If cbo开单科室.ListIndex = -1 Then Call SetDefaultDept(cbo开单科室, mrs开单科室, mrs开单人, lng开单人ID)
        Else
            cbo开单科室.Clear
        End If
        Call cbo开单科室_Click
    End If
    
    '护士类别
    If Bill.Active Then
        If mobjBill.Details.Count < Bill.Rows - 1 And Bill.Row = Bill.Rows - 1 _
            And Bill.RowData(Bill.Rows - 1) <> 0 Then
            '清除无效输入
            Bill.TextMatrix(Bill.Rows - 1, BillCol.类别) = ""
            Bill.RowData(Bill.Rows - 1) = 0
        ElseIf Bill.Col = BillCol.类别 Then
            Call Bill_EnterCell(Bill.Row, Bill.Col) '刷新
        End If
    End If
    
    '护士类别:判断非法输入
    If CheckInhibitiveByNurse(mobjBill, mrs开单人) Then
        MsgBox "护士只能输入治疗及材料项目,而单据中存在其它类型的项目。", vbInformation, gstrSysName
    End If
End Sub

Private Sub cbo开单人_KeyDown(KeyCode As Integer, Shift As Integer)
    If cbo开单人.Locked Then Exit Sub
    mblnDrop = False
    If KeyCode = 13 Then mblnDrop = SendMessage(cbo开单人.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 1
End Sub


Private Sub cbo开单人_LostFocus()
    If gstrIme <> "不自动开启" Then Call OpenIme
End Sub

Private Sub cbo开单人_Validate(Cancel As Boolean)
    If cbo开单人.Text <> "" Then
        If cbo.FindIndex(cbo开单人, zlStr.NeedName(cbo开单人.Text), True) = -1 Then cbo开单人.ListIndex = -1: cbo开单人.Text = ""
    End If
    If cbo开单人.Text = "" Then Call cbo开单人_KeyPress(vbKeyReturn)
    '当开单科室确定开单人时,可能此时不选开单人,先去调整开单科室后再来选
    If gblnFromDr And gbln开单人 And cbo开单人.ListIndex = -1 And mobjBill.Details.Count > 0 Then Cancel = True
End Sub

Private Sub chkCancel_Click()
    Dim i As Long
    
    mstrInNO = ""
    Call NewBill
    Call ClearRows
    Call Bill.ClearBill
    
    Bill.AllowAddRow = (chkCancel.Value = 0)
    
    Call SetDrawDrugDeptVisible
        
    If chkCancel.Value = 1 Then
        chkCancel.ForeColor = &HFF&
        
        picUnit.Enabled = False
        fraAppend.Enabled = False
        chkIn.Enabled = False
        
        For i = 0 To UBound(marrColData)
            Bill.ColData(i) = BillColType.Text_UnModify
        Next
        Call ShowDeleteCol(True)
        Bill.Active = True
        
        If InStr(mstrPrivsOpt, ";医生查询;") = 0 Then cbo开单人.Visible = False: lbl开单人.Visible = False
        Call SetDisible
        cboNO.Locked = False
        cboNO.SetFocus
    Else
        chkCancel.ForeColor = 0
        
        For i = 0 To UBound(marrColData)
            Bill.ColData(i) = marrColData(i)
        Next
        Call ShowDeleteCol(False)
        
        If gbytBilling = 2 Then
            Call SetDisible
            Bill.Active = False
            cboNO.Locked = False
            cboNO.SetFocus
        Else
            Call SetDisible(True)
            picUnit.Enabled = True
            fraAppend.Enabled = True
            chkIn.Enabled = True
            Bill.Active = True
            cboNO.Locked = True
            Bill.SetFocus
        End If
        
        If InStr(mstrPrivsOpt, ";医生查询;") = 0 Then cbo开单人.Visible = True: lbl开单人.Visible = True
        Call cbo开单科室_Click
    End If
End Sub

Private Sub chkIn_Click()
    sta.Panels(2) = ""
    If chkIn.Value = Checked Then
        txtIn.Enabled = True
        txtIn.BackColor = &H80000005
        sta.Panels(2) = "请输入要导入的记帐单单据号码"
        txtIn.SetFocus
    Else
        txtIn.Text = ""
        txtIn.Enabled = False
        txtIn.BackColor = &HE0E0E0
        Bill.SetFocus
    End If
End Sub

Private Sub chk加班_Click()
    Dim blnAdd As Boolean
    
    If mbytInState = 1 Or chkCancel.Value = 1 Or gbytBilling = 2 Then Exit Sub
    If mbytInState = 2 Then Exit Sub
    If Not chk加班.Visible Then Exit Sub
    
    blnAdd = OverTime(zlDatabase.Currentdate)
    If chk加班.Value = Unchecked And blnAdd Then
        If MsgBox("当前处于加班时间范围内,要取消加班加价吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            chk加班.Value = Checked
        End If
    End If
    If chk加班.Value = Checked And Not blnAdd Then
        If MsgBox("当前不处于加班时间范围内,要执行加班加价吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            chk加班.Value = Unchecked
        End If
    End If
    mobjBill.加班标志 = IIf(chk加班.Value = Checked, 1, 0)
    
    '重新计算价格
    Call CalcMoneys
    Call ShowDetails
End Sub

Private Sub chk加班_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    If mobjBill.Details.Count > 0 And Bill.Active And mbytInState = 0 And mstrInNO = "" Then
        Call Form_KeyDown(vbKeyF6, 0): Exit Sub
    End If
    
    Unload Me
End Sub

Private Sub cmdClear_Click()
    Dim i As Long
    For i = 1 To Bill.Rows - 1
        If Bill.TextMatrix(i, BillCol.项目) <> "" Then Bill.TextMatrix(i, Bill.Cols - 1) = ""
    Next
End Sub
Private Function CheckMainOperation() As Boolean
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是手术输入情况(如果不存在主要手术,但存在附加手术,则禁止
    '入参:
    '出参:lngRow-返回附加手术的行
    '返回:存在主手术或没有输入附加手术,返回true,否则返回False
    '编制:
    '修改:刘兴洪(退号时,增加定位功能),增加参数;strBackNo
    '日期:2009/7/10
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCount As Long, lngRow As Long   '指定行
    Dim i As Long, j As Long
    
    
    For i = 1 To mobjBill.Details.Count
        lngCount = 0: lngRow = 0
        For j = 1 To mobjBill.Details.Count
            If mobjBill.Details(i).病人ID = mobjBill.Details(j).病人ID Then
                If mobjBill.Details(j).收费类别 = "F" Then
                   If mobjBill.Details(j).附加标志 = 0 Then lngCount = 0: Exit For     '存在主要手术,则不检查,直接返回true
                   lngCount = lngCount + 1  '表示附加手术
                   If lngRow <= 0 Then lngRow = j
                End If
            End If
        Next
        If lngCount <> 0 Then Exit For
    Next
    
    If lngCount <> 0 Then
          MsgBox "单据中不存主要手术,但存在附加手术,请检查！", vbInformation, gstrSysName
          If Bill.Rows > lngRow Then Bill.Row = lngRow
          If Bill.Visible Then Bill.SetFocus
          Exit Function
    End If
    CheckMainOperation = True
End Function

Private Sub cmdOK_Click()
    Dim rsTmp As ADODB.Recordset, rsExamine As ADODB.Recordset, rsFeeItem As ADODB.Recordset
    Dim strSQL As String, strInfo As String, strTmp As String, strAddDate As String, strRows As String
    Dim arrPati() As Variant, strMoney As String, str汇总号 As String, str开单科室 As String
    Dim strPatis As String, i As Long, j As Long, lng病人ID As Long, lng主页ID As Long, lng险类 As Long
    Dim curModiMoney As Currency, Curdate As Date, cur余额 As Currency, dbl数次 As Double
    Dim strInsure As String, arrInsure As Variant
    Dim dblTotal As Double, blnTrans As Boolean
    Dim colStock As Collection
    Dim arrSMSQL As Variant, str销帐申请IDs As String, str申请人s As String
    Dim cllPro As Collection
    Dim rsItems As ADODB.Recordset
    
    If mbytInState = 3 Or (mbytInState = 0 And chkCancel.Visible And chkCancel.Value = 1) Then
        If mbytInState = 0 And mstrInNO = "" Then
            MsgBox "没有读取单据内容,不能销帐！", vbInformation, gstrSysName
            cboNO.SetFocus: Exit Sub
        End If
        
        For i = 1 To UBound(marrSerial)
            If Bill.TextMatrix(i, Bill.Cols - 1) = "√" Then
                strRows = strRows & "," & marrSerial(i)
            End If
        Next
        If strRows = "" Then
            MsgBox "请至少选择一行要销帐的费用！", vbInformation, gstrSysName
            Bill.SetFocus: Exit Sub
        End If
        
        If zlCheckIsExistsApplied(mstrInNO, strRows, str销帐申请IDs, str申请人s) Then
            '问题:47416
            If MsgBox("注意:" & vbCrLf & "    单据" & mstrInNO & "中存在申请销帐的项目,销帐后,将会自动取消" & vbCrLf & "申请人的申请项目,是否继续销帐?" & vbCrLf & "申请人如下: " & str申请人s, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        
        '所有行选择处理
        strRows = Mid(strRows, 2)
        i = GetBillRows(mstrInNO, 2)
        If UBound(Split(strRows, ",")) + 1 = i Then strRows = ""
        
        If strRows <> "" And InStr(1, mstrPrivsOpt, ";部分销帐;") = 0 Then
            MsgBox "你没有部分销帐的权限，只能对该单据全部销帐！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '判断是否包含医保病人费用,以及是否允许部份销帐(只要有一个不允许,则不允许)
        '去掉了医保连接匹配检查
        If gbytBilling = 0 Then
            Call GetBillInsures(strInsure, mstrInNO, mstrTime)
            If strInsure <> "" Then
                arrInsure = Split(strInsure, ",")
                If strRows <> "" Then
                    For i = 0 To UBound(arrInsure)
                        If gclsInsure.GetCapability(support记帐作废上传, , arrInsure(i)) Then
                            If Not gclsInsure.GetCapability(support允许部份冲销单据, , arrInsure(i)) Then
                                MsgBox "因为医保处理需要,该单据中的项目必须全部销帐！", vbInformation, gstrSysName
                                Exit Sub
                            End If
                        End If
                    Next
                End If
            End If
        End If
        
        
        Set rsFeeItem = GetNOFeeItem(mstrInNO, 2, strRows)
        If rsFeeItem.RecordCount > 0 Then
            For i = 1 To Bill.Rows - 1
                rsFeeItem.Filter = "序号=" & marrSerial(i)
                If rsFeeItem.RecordCount > 0 Then
                    If Not (InStr(",5,6,7,", rsFeeItem!收费类别) > 0 And gbln分离发药) Then
                        strTmp = mcolPatiInfo("R" & i)
                        Set rsTmp = GetPatientFeeItemTotal(Split(strTmp, ",")(0), Split(strTmp, ",")(1), mstrInNO)
                        rsTmp.Filter = "收费细目id=" & rsFeeItem!收费细目ID & " And 执行部门id=" & rsFeeItem!执行部门ID
                        If rsTmp.RecordCount > 0 Then
                            If Bill.TextMatrix(i, BillCol.数次) * Bill.TextMatrix(i, BillCol.付数) > rsTmp!数量 Then
                                MsgBox "第" & i & "行销帐数量大于可销帐数量" & rsTmp!数量 & ".", vbInformation, gstrSysName
                                Exit Sub
                            End If
                        Else
                            MsgBox "第" & i & "行可销帐数量为零.", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                End If
            Next
        End If
         '问题:47416
        Set cllPro = New Collection
        If str销帐申请IDs <> "" Then
            strSQL = "zl_病人费用销帐_Delete('" & str销帐申请IDs & "')"
            zlAddArray cllPro, strSQL
        End If
        strSQL = "zl_住院记帐记录_DELETE('" & mstrInNO & "','" & strRows & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
        zlAddArray cllPro, strSQL
        
        cmdOK.Enabled = False
        On Error GoTo errH
            blnTrans = True
            zlExecuteProcedureArrAy cllPro, Me.Caption, True
            'Call zldatabase.ExecuteProcedure(strSQL, Me.Caption)
            '医保记帐作废上传(只要第一个成功则提交)
            If gbytBilling = 0 And strInsure <> "" Then
                For i = 0 To UBound(arrInsure)
                    If gclsInsure.GetCapability(support记帐作废上传, , arrInsure(i)) And Not gclsInsure.GetCapability(support记帐完成后上传, , arrInsure(i)) Then
                        If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , arrInsure(i)) Then
                            If i = 0 Then gcnOracle.RollbackTrans: cmdOK.Enabled = True: Exit Sub
                        End If
                    End If
                Next
            End If
        gcnOracle.CommitTrans: blnTrans = False
        
        '医保记帐作废上传
        If gbytBilling = 0 And strInsure <> "" Then
            For i = 0 To UBound(arrInsure)
                If gclsInsure.GetCapability(support记帐作废上传, , arrInsure(i)) And gclsInsure.GetCapability(support记帐完成后上传, , arrInsure(i)) Then
                    If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , arrInsure(i)) Then
                        MsgBox "单据中 " & GetInsureName(Val(arrInsure(i))) & " 的费用向医保传送失败，这些费用已销帐。", vbInformation, gstrSysName
                    End If
                End If
            Next
        End If
        
        cmdOK.Enabled = True
        If mbytInState = 0 Then
            mstrInNO = "": mstr病人IDs = ""
            cboNO.Text = ""
            Call ClearRows
            Call Bill.ClearBill
            Call NewBill
            
            chkCancel.Value = 0
            
            If gbytBilling = 2 Then
                cboNO.SetFocus
            Else
                Bill.SetFocus
            End If
        Else
           gblnOK = True: Unload Me: Exit Sub
        End If
    ElseIf mbytInState = 2 Then
        If Not IsDate(txtDate.Text) Then
            MsgBox "请输入合法的费用时间！", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        
        strInfo = Check发生时间(CDate(txtDate.Text), cboNO.Text)
        If strInfo <> "" Then
            MsgBox strInfo, vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        
        If Not SaveModi() Then Exit Sub
        gblnOK = True: Unload Me: Exit Sub
    ElseIf Bill.Active And chkCancel.Value = 0 Then '正常输入单据状态
        If mobjBill.Details.Count = 0 Then
            MsgBox "单据中没有任何内容,请正确输入单据内容！", vbExclamation, gstrSysName
            Bill.SetFocus: Exit Sub
        End If
        
        i = Check执行科室
        If i <> 0 Then
            MsgBox "单据中第 " & i & " 行项目没有指定执行科室！", vbInformation, gstrSysName
            Bill.SetFocus: Exit Sub
        End If
        
        If mobjBill.开单部门ID = 0 Then
            MsgBox "请选择开单科室！", vbExclamation, gstrSysName
            cbo开单科室.SetFocus: Exit Sub
        End If
        
        '非法行
        dbl数次 = 0
        For i = 1 To mobjBill.Details.Count
            '27467,52828
            If mobjBill.Details(i).数次 <> 0 And dbl数次 = 0 Then
                dbl数次 = mobjBill.Details(i).数次
            End If
            If mobjBill.Details(i).收费细目ID = 0 Then
                MsgBox "单据中第 " & i & " 行没有正确输入数据,请修正或删除该行！", vbExclamation, gstrSysName
                Bill.SetFocus: Exit Sub
            ElseIf InStr(1, ",5,6,7,", mobjBill.Details(i).收费类别) > 0 Then
                '收集药品的发药药房
                strTmp = strTmp & "," & mobjBill.Details(i).收费细目ID
            End If
        Next
        '27467,52828
        If mbytInState = 0 And FormatEx(dbl数次, 7) = 0 Then
            MsgBox "单据中至少要有一条不为零的数次,请检查！", vbInformation, gstrSysName
            Bill.SetFocus: Exit Sub
        End If
        
         '检查药品的发药药房对应的服务科室
        If strTmp <> "" And Not gbln分离发药 Then
            strTmp = Mid(strTmp, 2)
            Set rsTmp = GetServiceDept(strTmp)
            If Not rsTmp Is Nothing Then
                strTmp = ""
                For i = 1 To mobjBill.Details.Count
                    If InStr(1, ",5,6,7,", mobjBill.Details(i).收费类别) > 0 Then
                        strInfo = mobjBill.Details(i).收费细目ID
                        '先检查是否是允许的存储库房
                        rsTmp.Filter = "收费细目ID=" & strInfo & " And 执行科室id=" & mobjBill.Details(i).执行部门ID
                        If rsTmp.RecordCount = 0 Then
                            strTmp = strTmp & "," & i
                        Else
                            '再检查是否是允许的服务科室(没有设置服务科室的,开单科室ID为零)
                            rsTmp.Filter = "(" & rsTmp.Filter & " And 开单科室ID=" & mobjBill.Details(i).科室ID & ") Or (" & rsTmp.Filter & " And 开单科室ID=0)"
                            If rsTmp.RecordCount = 0 Then
                                strTmp = strTmp & "," & i
                            End If
                        End If
                    End If
                Next
                If strTmp <> "" Then
                    strTmp = Mid(strTmp, 2)
                    MsgBox "请检查,第" & strTmp & "行药品是否违反以下规则:" & vbCrLf & vbCrLf & _
                        "A.选择的执行科室不是药品的存储库房" & vbCrLf & _
                        "B.病人科室不属于药品在此存储库房的服务科室.", _
                        vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        End If
        
        '费用发生时间检查
        If Not IsDate(txtDate.Text) Then
            MsgBox "请输入正确的费用日期！", vbExclamation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        strSQL = ""
        For i = 1 To mobjBill.Details.Count
            If InStr(strSQL & ",", "," & mobjBill.Details(i).病人ID & ",") = 0 Then
                strInfo = Check发生时间(CDate(txtDate.Text), mobjBill.Details(i).病人ID)
                If strInfo <> "" Then
                    MsgBox strInfo, vbInformation, gstrSysName
                    txtDate.SetFocus: Exit Sub
                End If
                strSQL = strSQL & "," & mobjBill.Details(i).病人ID
            End If
        Next
        
        If mobjBill.开单人 = "" And gbln开单人 Then
            MsgBox "请输入开单人！", vbInformation, gstrSysName
            cbo开单人.SetFocus: Exit Sub
        End If
        
        '护士类别:判断非法输入
        If CheckInhibitiveByNurse(mobjBill, mrs开单人) Then
            MsgBox "护士只能输入治疗及材料项目,而单据中存在其它类型的项目。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        
        '医保负数记帐检查    因为操作员可能先输非医保病人的负数单据,再换为医保病人,所以要再检查一次
        If InStr(mstrPrivsOpt, ";负数记帐;") > 0 Then     '至少有其中一种负数记帐权限,才可能是负数
            For i = 1 To mobjBill.Details.Count
                If Val(mobjBill.Details(i).发药窗口) <> 0 Then
                    If mobjBill.Details(i).数次 * mobjBill.Details(i).付数 < 0 Then
                        If Not gclsInsure.GetCapability(support负数记帐, mobjBill.Details(i).病人ID, Val(mobjBill.Details(i).发药窗口)) Then
                            MsgBox "单据中第 " & i & " 行是负数,本地医保不支持负数记帐！", vbInformation, gstrSysName
                            Bill.SetFocus: Exit Sub
                        End If
                    End If
                End If
            Next
        End If
        
        '出院强制记帐权限检查
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
                If InStr(strPatis & ",", "," & .病人ID & ",") = 0 Then
                    strPatis = strPatis & "," & .病人ID
                    If Not PatiCanBilling(.病人ID, .主页ID, mstrPrivsOpt) Then Exit Sub
                    If zlPatiIS病案已编目(.病人ID, .主页ID) = True Then     '问题:28725
                        Exit Sub
                    End If
                    If zlIsAllowFeeChange(.病人ID, .主页ID) = False Then
                        Exit Sub
                    End If
                End If
            End With
        Next
       
              
        '处方职务检查
        '公费或医保病人
        i = CheckDuty(, False)
        If i > 0 Then
            Bill.Row = i: Bill.MsfObj.TopRow = i
            Bill.Col = BillCol.项目: Bill.SetFocus
            Exit Sub
        End If
        
        '所有病人项目
        i = CheckDuty(, True)
        If i > 0 Then
            Bill.Row = i: Bill.MsfObj.TopRow = i
            Bill.Col = BillCol.项目: Bill.SetFocus
            Exit Sub
        End If
                
        '医保费用项目是否审批检查，输入时已检查，保存时再检查是因为：
        '1。主从项批量添加时只检查了主项；2.导入单据时未检查
        '注意:可能存在医保和非医保病人混合的情况,以及只有一行费用的情况
        strInfo = "": strTmp = "": lng病人ID = 0: lng主页ID = 0: lng险类 = 0
        str开单科室 = zlStr.NeedName(cbo开单科室.Text)
        For i = 1 To mobjBill.Details.Count
            lng主页ID = mobjBill.Details(i).主页ID
            lng险类 = Val(mobjBill.Details(i).发药窗口)
            
            If lng病人ID <> mobjBill.Details(i).病人ID And lng险类 > 0 Then
                Set rsTmp = GetAuditRecord(lng病人ID, lng主页ID)
                Set rsExamine = GetExamineItem(strTmp, lng险类)
                For j = 1 To rsExamine.RecordCount
                    rsTmp.Filter = "项目ID=" & rsExamine!收费细目ID
                    If rsTmp.RecordCount = 0 Then
                        strInfo = strInfo & "," & GetRowByFeeItemID(mobjBill.Details, rsExamine!收费细目ID, lng病人ID)
                    ElseIf Not IsNull(rsTmp!可用数量) Then
                        If mobjBill.Details(i).付数 * mobjBill.Details(i).数次 * IIf(gbln住院单位, mobjBill.Details(i).Detail.住院包装, 1) > rsTmp!可用数量 Then
                            MsgBox "第" & i & "行收费项目的数次超过了批准的使用限量" & FormatEx(rsTmp!可用数量 / IIf(gbln住院单位, mobjBill.Details(i).Detail.住院包装, 1), 5) & "。", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                    rsExamine.MoveNext
                Next
                strTmp = ""
                
                
                If gclsInsure.GetCapability(support实时监控, lng病人ID, lng险类) Then
                    If gclsInsure.CheckItem(lng险类, 1, 2, MakeDetailRecord(mobjBill, mobjBill.开单人, str开单科室, 2, gbytBilling)) = False Then
                       Exit Sub
                    End If
                End If
            End If
            
            lng病人ID = mobjBill.Details(i).病人ID
            If lng险类 > 0 Then
                strTmp = IIf(strTmp = "", "", strTmp & ",") & mobjBill.Details(i).收费细目ID
            End If
        Next
        
        If strTmp <> "" Then
            Set rsTmp = GetAuditRecord(lng病人ID, lng主页ID)
            Set rsExamine = GetExamineItem(strTmp, lng险类)
            For j = 1 To rsExamine.RecordCount
                rsTmp.Filter = "项目ID=" & rsExamine!收费细目ID
                If rsTmp.RecordCount = 0 Then
                    strInfo = strInfo & "," & GetRowByFeeItemID(mobjBill.Details, rsExamine!收费细目ID, lng病人ID)
                ElseIf Not IsNull(rsTmp!可用数量) Then
                    i = GetRowByFeeItemID(mobjBill.Details, rsExamine!收费细目ID, lng病人ID)
                    If mobjBill.Details(i).付数 * mobjBill.Details(i).数次 * IIf(gbln住院单位, mobjBill.Details(i).Detail.住院包装, 1) > rsTmp!可用数量 Then
                        MsgBox "第" & i & "行收费项目的数次超过了批准的使用限量" & FormatEx(rsTmp!可用数量 / IIf(gbln住院单位, mobjBill.Details(i).Detail.住院包装, 1), 5) & "。", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
                rsExamine.MoveNext
            Next
        End If
        
        If strInfo <> "" Then
            MsgBox "第" & Mid(strInfo, 2) & "行收费项目要求审批,当前病人未被批准使用!", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '主项适用病人病区科室
        For i = 1 To mobjBill.Details.Count
            If InStr(",5,6,7,", mobjBill.Details(i).收费类别) = 0 Then
                If CheckItemHaveSub(i) Then
                    If Not CheckFeeItemLimitDept(mobjBill.Details(i).收费细目ID, mobjBill.Details(i).病区ID, mobjBill.Details(i).科室ID) Then
                        MsgBox "第" & i & "行的收费项目对该病人病区和科室不适用！", vbInformation, gstrSysName
                        Bill.Row = i: Bill.MsfObj.TopRow = i
                        Bill.Col = BillCol.项目: Bill.SetFocus
                        Exit Sub
                    End If
                End If
            End If
        Next
        
        '费用类型检查
        If Not Check费用类型 Then Exit Sub
                
        '记帐分类报警:多个病人记帐报警
        mrsWarn.Filter = ""
        If mrsWarn.RecordCount > 0 Then
            ReDim arrPati(7) '病人信息数组
            '循环单据行对每个病人处理
            For i = 1 To mobjBill.Details.Count
                '收集该行病人信息
                If mobjBill.Details(i).病人ID <> arrPati(0) Then
                    arrPati(0) = mobjBill.Details(i).病人ID  '病人ID
                    arrPati(1) = CStr(mobjBill.Details(i).姓名 & IIf(mobjBill.Details(i).住院号 = "", "", "(住院号:" & mobjBill.Details(i).住院号 & " 床号:" & mobjBill.Details(i).床号 & ")")) '姓名
                    arrPati(2) = CCur(CalcOneTotal(CLng(i), False)) '单据金额
                    arrPati(3) = CCur(mobjBill.Details(i).担保额) '担保额
                    arrPati(4) = GetMedPayMode(mobjBill.Details(i).医疗付款, mrsMedPayMode)
                    
                    '刷新读取:预交余额,费用余额;剩余款额;当日费用
                    curModiMoney = 0
                    If mstrInNO <> "" Then
                        curModiMoney = GetBillMoney(2, mstrInNO, mobjBill.Details(i).病人ID)
                    End If
                    
                    strMoney = "0,0,0"
                    Set rsTmp = GetMoneyInfo(mobjBill.Details(i).病人ID, IIf(gbytBilling = 0, curModiMoney, 0), True, 2)
                    If Not rsTmp Is Nothing Then
                        strMoney = rsTmp!预交余额 & "," & rsTmp!费用余额 & "," & rsTmp!预交余额 - rsTmp!费用余额
                    End If
                    strMoney = strMoney & "," & GetPatiDayMoney(mobjBill.Details(i).病人ID) - mcurModiMoney '病人当日费用额
                    
                    For j = 1 To mobjBill.Details.Count
                        If mobjBill.Details(j).病人ID = mobjBill.Details(i).病人ID Then
                            mobjBill.Details(j).就诊卡号 = strMoney
                        End If
                    Next
                                    
                    '剩余款额,当日费用额
                    arrPati(5) = Val(Split(strMoney, ",")(2))
                    arrPati(6) = Val(Split(strMoney, ",")(3))
                    
                    '病人险类
                    arrPati(7) = Val(mobjBill.Details(i).发药窗口)
                                    
                    cur余额 = CCur(arrPati(5))
                    If gbln报警包含划价费用 Then cur余额 = CCur(arrPati(5)) - GetPriceMoneyTotal(1, CLng(arrPati(0))) + IIf(gbytBilling = 1, mcurModiMoney, 0)
                End If
                
                '该行进行报警
                If CCur(arrPati(2)) > 0 Then
                    gbytWarn = BillingWarn(mstrPrivsOpt, CStr(arrPati(1)), mobjBill.Details(i).病区ID, GetPatiWarnRange(mobjBill.Details(i).病人ID, mobjBill.Details(i).主页ID), mrsWarn, _
                            cur余额, CCur(arrPati(6)), CCur(arrPati(2)), CCur(arrPati(3)), mobjBill.Details(i).收费类别, mobjBill.Details(i).Detail.类别名称, mstrWarn, True)
                    If gbytWarn = 2 Or gbytWarn = 3 Then
                        Bill.Row = i: Exit Sub
                    End If
                End If
            Next
        End If
        
        '药品禁忌检查
        strInfo = CheckDisable(mobjBill)
        If strInfo <> "" Then
            If strInfo Like "*(互相禁用)*" Then
                MsgBox strInfo, vbInformation, gstrSysName
                Exit Sub
            Else
                If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
        End If
        
        '处方限量检查
        If Not gbln处方限量 Then
            If Not CheckLimit(mobjBill, , gbln住院单位) Then Exit Sub
        End If
        
        '检查分批或时价药品同一药房是否有重复输入
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
                If (.Detail.分批 Or .Detail.变价) _
                    And (InStr(",5,6,7,", .收费类别) > 0 Or .收费类别 = "4" And .Detail.跟踪在用) Then
                    For j = 1 To mobjBill.Details.Count
                        If i <> j And .收费细目ID = mobjBill.Details(j).收费细目ID And .执行部门ID = mobjBill.Details(j).执行部门ID And .病人ID = mobjBill.Details(j).病人ID Then
                            If .收费类别 = "4" Then
                                MsgBox "第 " & j & " 行的分批或时价卫生材料""" & .Detail.名称 & """在同一个发料部门被重复输入，请合并！", vbInformation, gstrSysName
                            Else
                                MsgBox "第 " & j & " 行的分批或时价药品""" & .Detail.名称 & """在同一个药房被重复输入，请合并！", vbInformation, gstrSysName
                            End If
                            Exit Sub
                        End If
                    Next
                End If
            End With
        Next
        
        '药品库存检查,71188:刘尔旋,2014-04-03,对不足提醒的也要进行检查
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
                Set colStock = IIf(.收费类别 = "4", mcolStock2, mcolStock1)
                If InStr(",5,6,7,", .收费类别) > 0 And Not gbln分离发药 Then
                    If .Detail.分批 Or .Detail.变价 Then
                        dblTotal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
                        .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                        If gbln住院单位 Then .Detail.库存 = .Detail.库存 / .Detail.住院包装
                        
                        If mbytInState = 0 And mstrInNO <> "" Then .Detail.库存 = .Detail.库存 + GetOriginalTotal(mobjBill, .收费细目ID, .执行部门ID)
                        If dblTotal > .Detail.库存 Then
                            MsgBox "第 " & i & " 行时价或分批药品""" & .Detail.名称 & _
                                """的当前库存" & IIf(InStr(1, mstrPrivsOpt, ";显示库存;") > 0, .Detail.库存, "") & "不足输入数量""" & dblTotal & """。", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    ElseIf colStock("_" & .执行部门ID) = 2 Then
                        dblTotal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
                        .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                        If gbln住院单位 Then .Detail.库存 = .Detail.库存 / .Detail.住院包装
                        
                        If mbytInState = 0 And mstrInNO <> "" Then .Detail.库存 = .Detail.库存 + GetOriginalTotal(mobjBill, .收费细目ID, .执行部门ID)
                        If dblTotal > .Detail.库存 Then
                            MsgBox "第 " & i & " 行药品""" & .Detail.名称 & _
                                """的当前库存" & IIf(InStr(1, mstrPrivsOpt, ";显示库存;") > 0, .Detail.库存, "") & "不足输入数量""" & dblTotal & """,请修改或检查是否有多行输入。", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    ElseIf colStock("_" & .执行部门ID) = 1 Then
                        dblTotal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
                        .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                        If gbln住院单位 Then .Detail.库存 = .Detail.库存 / .Detail.住院包装
                        
                        If mbytInState = 0 And mstrInNO <> "" Then .Detail.库存 = .Detail.库存 + GetOriginalTotal(mobjBill, .收费细目ID, .执行部门ID)
                        If dblTotal > .Detail.库存 Then
                            If MsgBox("第 " & i & " 行药品""" & .Detail.名称 & _
                                """的当前库存" & IIf(InStr(1, mstrPrivsOpt, ";显示库存;") > 0, .Detail.库存, "") & "不足输入数量""" & dblTotal & """,要继续吗?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
                        End If
                    End If
                ElseIf InStr(",5,6,7,", .收费类别) > 0 And gbln分离发药 And gblnStock Then
                    '单据对象的库存是本地参数指定的药房的库存之和
                    strInfo = Decode(.Detail.类别, "5", gstr西药房, "6", gstr成药房, "7", gstr中药房)
                    If strInfo <> "" Then
                        dblTotal = GetDrugTotal(mobjBill, .收费细目ID, 0)
                        If mbytInState = 0 And mstrInNO <> "" Then .Detail.库存 = .Detail.库存 + GetOriginalTotal(mobjBill, .收费细目ID, 0)
                        
                        If dblTotal > .Detail.库存 Then
                            MsgBox "第 " & i & " 行药品""" & .Detail.名称 & "]的当前库存" & IIf(InStr(1, mstrPrivsOpt, ";显示库存;") > 0, .Detail.库存, "") & _
                                "不足输入数量""" & dblTotal & """,请修改或检查是否有多行输入。", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                ElseIf .收费类别 = "4" And .Detail.跟踪在用 Then
                    If .Detail.分批 Or .Detail.变价 Then
                        dblTotal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
                        .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                        
                        If mbytInState = 0 And mstrInNO <> "" Then .Detail.库存 = .Detail.库存 + GetOriginalTotal(mobjBill, .收费细目ID, .执行部门ID)
                        If dblTotal > .Detail.库存 Then
                            MsgBox "第 " & i & " 行时价或分批卫生材料""" & .Detail.名称 & _
                                """的当前库存" & IIf(InStr(1, mstrPrivsOpt, ";显示库存;") > 0, .Detail.库存, "") & "不足输入数量""" & dblTotal & """。", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    ElseIf colStock("_" & .执行部门ID) = 2 Then
                        dblTotal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
                        .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                        
                        If mbytInState = 0 And mstrInNO <> "" Then .Detail.库存 = .Detail.库存 + GetOriginalTotal(mobjBill, .收费细目ID, .执行部门ID)
                        If dblTotal > .Detail.库存 Then
                            MsgBox "第 " & i & " 行卫生材料""" & .Detail.名称 & _
                                """的当前库存" & IIf(InStr(1, mstrPrivsOpt, ";显示库存;") > 0, .Detail.库存, "") & "不足输入数量""" & dblTotal & """,请修改或检查是否有多行输入。", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    ElseIf colStock("_" & .执行部门ID) = 1 Then
                        dblTotal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
                        .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                        
                        If mbytInState = 0 And mstrInNO <> "" Then .Detail.库存 = .Detail.库存 + GetOriginalTotal(mobjBill, .收费细目ID, .执行部门ID)
                        If dblTotal > .Detail.库存 Then
                            If MsgBox("第 " & i & " 行卫生材料""" & .Detail.名称 & _
                                """的当前库存" & IIf(InStr(1, mstrPrivsOpt, ";显示库存;") > 0, .Detail.库存, "") & "不足输入数量""" & dblTotal & """,要继续吗?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
                        End If
                    End If
                End If
            End With
        Next
    
        '零差价检查,105875
        If Not gobjPublicDrug Is Nothing Then
            'Private Function zlCheckPriceAdjustBySell(ByVal lng药品id As Long, ByVal lng药房id As Long) As Boolean
            '零差价管理模式时，判断价格是否满足零差价管理要（成本价和售价一致）
            '定价药品：售价是固定的，比较所有药房的成本价，如果存在不一致的就不能销售出库
            '时价药品：比较药房库存记录的零售价和成本价，如果存在不一致的就不能销售出库
            '销售出库时只判断药房
            '返回：True-正常进行销售出库；false-不能进行销售出库
            For i = 1 To mobjBill.Details.Count
                With mobjBill.Details(i)
                    If InStr(",5,6,7,", .收费类别) > 0 Then
                        If gobjPublicDrug.zlCheckPriceAdjustBySell(.收费细目ID, .执行部门ID) = False Then
                            Exit Sub
                        End If
                    End If
                End With
            Next
        End If
        
        '刘兴洪:22441,检查主手术和附加手术情况
        If CheckMainOperation = False Then Exit Sub
        
        '项目服务对象检查(主要因为多了门诊留观病人)
        If Check服务对象 > 0 Then Exit Sub
        
        If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModule, 1, 1, _
            MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 2, gbytBilling)) = False Then
            Exit Sub
        End If
        
        '负数退费检查
        If Not CheckBillNegative Then Exit Sub
        
        '检查卫生材料的灭菌效期
        '药品自动发药
        mblnSendMateria = False
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
                If .收费类别 = "4" And .Detail.跟踪在用 Then
                    dblTotal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
                    If Not CheckValidity(.收费细目ID, .执行部门ID, dblTotal) Then Exit Sub
             
                ElseIf InStr(1, ",5,6,7,", .收费类别) > 0 Then
                    '打印发药单,仅普通记帐,且划价单除外
                    If gbytSendMateria <> 0 And mbytUseType = 0 And gbytBilling = 0 Then
                        '全部药品都确定了药房的才自动发药(分离发药时,没有确定药房)
                        mblnSendMateria = .执行部门ID <> 0
                    End If
                End If
            End With
        Next
        If InStr(mstrPrivsOpt, ";药品发药;") = 0 Then mblnSendMateria = False
        
        If mstrInNO <> "" Then
            If HaveExecute(2, mstrInNO, 2) Then
                MsgBox "该单据包含完全执行或部分执行的项目,不允许修改。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        If mblnSendMateria And gbytSendMateria = 2 Then
            If MsgBox("记帐完成后自动执行发药吗?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                mblnSendMateria = False
            End If
        End If
        
        If IsDate(txtDate.Text) Then mobjBill.发生时间 = CDate(txtDate.Text)
        mobjBill.登记时间 = zlDatabase.Currentdate     '注意:打印发药单时要用到这个时间
        If zlGetSaveDataItems_Plugin(mobjBill, rsItems, True) = False Then Exit Sub
        If zlChargeSaveValied_Plugin(mlngModule, 2, False, gbytBilling = 1, "", rsItems) = False Then Exit Sub
        
        cmdOK.Enabled = False
        If Not SaveBill Then
            cmdOK.Enabled = True
            Exit Sub
        Else
            Call zlChargeSaveAfter_Plugin(mlngModule, 0, 0, False, 2, mobjBill.NO)
            If gbytBilling = 0 And gbln记帐打印 Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_113" & 3 + mbytUseType, Me, "NO=" & mobjBill.NO, "登记时间=" & Format(mobjBill.登记时间, "yyyy-MM-dd HH:mm:ss"), "药品单位=" & IIf(gbln住院单位, 1, 0), "PrintEmpty=0", "重打=0", 2)
            ElseIf gbytBilling = 1 And gbln划价打印 Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1133", Me, "NO=" & mobjBill.NO, "登记时间=" & Format(mobjBill.登记时间, "yyyy-MM-dd HH:mm:ss"), "药品单位=" & IIf(gbln住院单位, 1, 0), "PrintEmpty=0", "重打=0", 2)
            End If
                        
            '打印发药单
            If mblnSendMateria Then
                If MsgBox("单据""" & mobjBill.NO & """发药完成，要打印发药清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1133_2", Me, "单据号=" & mobjBill.NO, "登记时间=" & Format(mobjBill.登记时间, "yyyy-MM-dd HH:mm:ss"), 1)
                End If
            End If
        
            cmdOK.Enabled = True
            If mstrInNO = "" Then
                sta.Panels(2) = "上一张单据:" & mobjBill.NO
                Call ClearRows: Call Bill.ClearBill
                Call NewBill: mstrInNO = ""
                Bill.SetFocus
            Else '修改
                '刘兴洪 问题:27083 日期:2009-12-25 10:09:21
                gblnOK = True: Unload Me: Exit Sub
            End If
        End If
    ElseIf Not Bill.Active Then '审核住院划价状态
        If mstrInNO = "" Then
            MsgBox "没有住院划价单据,请先输入！", vbInformation, gstrSysName
            cboNO.SetFocus: Exit Sub
        End If
        
        '取本次审核的行序号
        strSQL = ""
        For i = 1 To UBound(marrSerial)
            strSQL = strSQL & "," & marrSerial(i)
        Next
        strSQL = Mid(strSQL, 2)
        i = GetBillRows(mstrInNO, 2)
        If UBound(Split(strSQL, ",")) + 1 = i Then strSQL = ""
        
        '费用报警
        mrsWarn.Filter = ""
        If mrsWarn.RecordCount > 0 Then
            If Not AuditingWarn(mstrPrivsOpt, mrsWarn, mstrInNO, strSQL) Then Exit Sub
        End If
        
        '去掉了医保连接匹配检查
        Call GetBillInsures(strInsure, mstrInNO, , True)
        If strInsure <> "" Then arrInsure = Split(strInsure, ",")
        
        '药品自动发药
        mblnSendMateria = False
        If gbytSendMateria <> 0 And mbytUseType = 0 And InStr(mstrPrivsOpt, ";药品发药;") > 0 Then
            For i = 1 To Bill.Rows - 1
                If InStr(",西成药,中成药,中草药,", "," & Bill.TextMatrix(i, BillCol.类别) & ",") > 0 Then '因读取单据时没有存储类别编码,简化为根据名称判断
                    '全部药品都确定了药房的才自动发药(分离发药时,没有确定药房)
                    mblnSendMateria = Trim(Bill.TextMatrix(i, BillCol.执行科室)) <> ""
                End If
            Next
        End If
        If mblnSendMateria And gbytSendMateria = 2 Then
            If MsgBox("记帐审核后自动执行发药吗?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                mblnSendMateria = False
            End If
        End If
        
        cmdOK.Enabled = False
        arrSMSQL = Array()
        Curdate = zlDatabase.Currentdate
        strAddDate = "To_Date('" & Format(Curdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        strSQL = "zl_住院记帐记录_Verify('" & mstrInNO & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "','" & strSQL & "',NULL," & strAddDate & ")"
        str汇总号 = zlDatabase.GetNextNo(20)
        
        On Error GoTo errH
        gcnOracle.BeginTrans: blnTrans = True
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            
            '准备自动发药(仅普通记帐),必须在事务中才能读到数据
            If mblnSendMateria Then
                Set rsTmp = Get待发药清单(mstrInNO, Format(Curdate, "yyyy-MM-dd HH:mm:ss"), True)
                If rsTmp.RecordCount > 0 Then
                    ReDim arrSMSQL(rsTmp.RecordCount - 1)
                    For i = 0 To rsTmp.RecordCount - 1
                        arrSMSQL(i) = "ZL_药品收发记录_部门发药(" & rsTmp!库房ID & "," & rsTmp!ID & ",'" & UserInfo.姓名 & "'," & strAddDate & ",Null,Null,Null," & str汇总号 & ")"
                        rsTmp.MoveNext
                    Next
                End If
                rsTmp.Close
            End If
            '执行自动发药
            For i = 0 To UBound(arrSMSQL)
                Call zlDatabase.ExecuteProcedure(CStr(arrSMSQL(i)), Me.Caption)
            Next
            
            '医保上传(只要第一个成功则提交)
            If strInsure <> "" Then
                For i = 0 To UBound(arrInsure)
                    If gclsInsure.GetCapability(support记帐上传, , arrInsure(i)) And Not gclsInsure.GetCapability(support记帐完成后上传, , arrInsure(i)) Then
                        strInfo = ""
                        If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 1, strInfo, , arrInsure(i)) Then
                            If i = 0 Then gcnOracle.RollbackTrans
                            If strInfo <> "" Then MsgBox strInfo, vbInformation, gstrSysName
                            If i = 0 Then cmdOK.Enabled = True: Exit Sub
                        End If
                    End If
                Next
            End If
        gcnOracle.CommitTrans: blnTrans = False
        
        '医保上传
        If strInsure <> "" Then
            For i = 0 To UBound(arrInsure)
                If gclsInsure.GetCapability(support记帐上传, , arrInsure(i)) And gclsInsure.GetCapability(support记帐完成后上传, , arrInsure(i)) Then
                    strInfo = ""
                    If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 1, strInfo, , arrInsure(i)) Then
                        If strInfo <> "" Then
                            MsgBox strInfo, vbInformation, gstrSysName
                        Else
                            MsgBox "单据中的 " & GetInsureName(Val(arrInsure(i))) & " 费用向医保传送失败,这些费用已审核！", vbInformation, gstrSysName
                        End If
                    End If
                End If
            Next
        End If
        
        On Error GoTo 0
        
        If gbytBilling = 2 And gbln审核打印 And mblnPrint Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1133", Me, "NO=" & mstrInNO, "登记时间=" & Format(Curdate, "yyyy-MM-dd HH:mm:ss"), "药品单位=" & IIf(gbln住院单位, 1, 0), "PrintEmpty=0", "重打=0", 2)
        End If
        
        '打印发药单
        If mblnSendMateria Then
            If MsgBox("单据""" & mstrInNO & """发药完成，要打印发药清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1133_2", Me, "单据号=" & mstrInNO, "登记时间=" & Format(Curdate, "yyyy-MM-dd HH:mm:ss"), 1)
            End If
        End If
        
        cmdOK.Enabled = True
        mstrInNO = "": cboNO.Text = ""
        Call ClearRows: Call Bill.ClearBill
        Call NewBill: cboNO.Locked = False
        cboNO.SetFocus
    End If
    gblnOK = True
    Call SetDrawDrugDeptEnabled
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    cmdOK.Enabled = True
End Sub

Private Sub cmdOK_GotFocus()
    If mbytInState = 3 Or (chkCancel.Visible And chkCancel.Value = 1) Then
        Bill.Row = 1: Bill.Col = Bill.Cols - 1
    End If
End Sub

Private Sub cmdSelAll_Click()
    Dim i As Long
    For i = 1 To Bill.Rows - 1
        If Bill.TextMatrix(i, BillCol.项目) <> "" Then Bill.TextMatrix(i, Bill.Cols - 1) = "√"
    Next
End Sub

Private Sub Form_Activate()
    On Error Resume Next
        
    If mbytInState = 0 And gblnFromDr And mobjBill.Details.Count = 0 Then
        cbo开单人.SetFocus
    ElseIf mbytUseType = 1 And mbytInState = 0 Then
        Bill.SetFocus
    ElseIf gbytBilling = 2 Then
        cboNO.SetFocus
    ElseIf mbytInState = 1 Then
        cmdCancel.SetFocus
    ElseIf mbytInState = 2 Then
        txtDate.SetFocus
    ElseIf mbytInState = 3 Then
        cmdOK.SetFocus
    Else
        Bill.SetFocus
    End If
    Call SetDrawDrugDeptVisible
    Call SetDrawDrugDeptEnabled
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("',:：;；?？|~" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub



Private Sub Form_Load()
    Dim i As Long, tmpBill As ExpenseBill
    
    mstrPrivsOpt = GetInsidePrivs(Enum_Inside_Program.p记帐操作)
    glngFormW = 12000: glngFormH = 7290
    If Not OS.IsDesinMode Then
        glngOld = GetWindowLong(Me.hWnd, GWL_WNDPROC)
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf Custom_WndMessage)
    End If
    RestoreWinState Me, App.ProductName, mbytInState
    sta.Visible = True
    
    gblnOK = False: mblnEnterCell = True
    mlngPreUnit = -1
    
    
    
    '初始化单据数据
    Set mobjBill = New ExpenseBill
    
    Call zlLoadDrawDeptData(mbytUseType, mlngDeptID)
    
    If mbytInState = 0 Or mbytInState = 2 Or mbytInState = 3 Then
        If Not InitData Then Unload Me: Exit Sub
    Else
        If Init开单人开单科室(cbo开单人, cbo开单科室, mrs开单人, mrs开单科室, mstrPrivs, mbytUseType, mlngDeptID) = False Then
            Exit Sub
        End If
    End If
    mstrUnitIDs = GetUserUnits
    Set mcolPatiInfo = New Collection
    
    
    Call InitFace
    Call NewBill
    
    
    If mbytInState <> 0 Then '显示、调整、销帐单据(1,2,3)
        If Not ReadBill(mstrInNO, (mbytInState = 3)) Then Unload Me: Exit Sub
        cboNO.Text = mstrInNO
        If InStr(mstrPrivsOpt, ";医生查询;") = 0 Then cbo开单人.Visible = False: lbl开单人.Visible = False
    Else '新增
        mstr药品价格等级 = gstr药品价格等级
        mstr卫材价格等级 = gstr卫材价格等级
        mstr普通价格等级 = gstr普通价格等级
        '读取该单据的内容
        If mstrInNO <> "" Then '修改单据
            Set mobjBill = ImportBill(mstrInNO, True, Me, True, gbln住院单位, True, , , mstr药品价格等级, mstr卫材价格等级, mstr普通价格等级)
            If mobjBill.NO = "" Then
                MsgBox "读取单据失败。", vbInformation, gstrSysName
                Unload Me: Exit Sub
            Else
                Call ReCalcInsure '重新计算统筹金额
                
                cboNO.Text = mobjBill.NO '显示原单号
                
                Bill.ClearBill
                Bill.Rows = mobjBill.Details.Count + 1
                Call InitBillColumnColor
                
                txtDate.Text = Format(mobjBill.发生时间, "yyyy-MM-dd HH:mm:ss")
                chk加班.Value = mobjBill.加班标志

                Call Set开单人开单科室(cbo开单人, cbo开单科室, mrs开单人, mrs开单科室, mobjBill.开单人, mobjBill.开单部门ID)
                                
                mobjBill.操作员编号 = UserInfo.编号
                mobjBill.操作员姓名 = UserInfo.姓名
                
                If gintPriceGradeStartType < 2 Then
                    If gbln从项汇总折扣 Then Call CalcMoneys
                Else
                    '每一行根据价格等级重算价格
                    For i = 1 To mobjBill.Details.Count
                        Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, _
                            mobjBill.Details(i).病人ID, mobjBill.Details(i).主页ID, mobjBill.Details(i).医疗付款, _
                            mstr药品价格等级, mstr卫材价格等级, mstr普通价格等级)
                        Call CalcMoneys(i)
                    Next
                End If
                Call ShowDetails
                Call SetIntureColor
            End If
        Else
            If mbytUseType = 1 And mlng病人ID <> 0 Then
                Bill.Row = 1: Bill.Col = BillCol.姓名
                Bill.Text = "-" & mlng病人ID
                Call Bill_KeyDown(13, 0, False)
                Bill.Text = ""
                Bill.TxtVisible = False
            End If
        End If
    End If
End Sub

Private Sub Form_Resize()
    Dim lngCancelW As Long
        
    On Error Resume Next
    
    Bill.Height = Me.ScaleHeight - Bill.Top - sta.Height - fraAppend.Height - picAppend.Height - IIf(fraDrawDept.Visible, fraDrawDept.Height, 0) + 75
    Bill.Left = 0: Bill.Width = Me.ScaleWidth
    
    If chkCancel.Visible Or lblFlag.Visible Then lngCancelW = chkCancel.Width
    fraTitle.Width = Me.ScaleWidth - fraTitle.Left
    chkCancel.Left = fraTitle.Width - chkCancel.Width - 60
    lblFlag.Left = chkCancel.Left + (chkCancel.Width - lblFlag.Width) / 2
    
    cboNO.Left = fraTitle.Width - lngCancelW - 60 - cboNO.Width - 30
    lblNO.Left = cboNO.Left - lblNO.Width - 45
        
    fraAppend.Top = Bill.Top + Bill.Height - 75
    fraAppend.Width = Me.ScaleWidth - fraAppend.Left
    
    fraDrawDept.Top = fraAppend.Top + fraAppend.Height - 150
    fraDrawDept.Width = Me.ScaleWidth - fraDrawDept.Left
    
    
    txtDate.Left = fraAppend.Width - txtDate.Width - 90
    lblDate.Left = txtDate.Left - lblDate.Width - 45
    
    If cbo开单人.Container Is picUnit Then
        cbo开单科室.Left = lblDate.Left - cbo开单科室.Width - 300
        lbl开单科室.Left = cbo开单科室.Left - lbl开单科室.Width - 45
    Else
        cbo开单人.Left = lblDate.Left - cbo开单人.Width - 300
        lbl开单人.Left = cbo开单人.Left - lbl开单人.Width - 45
    End If
    
    cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - 500
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 200
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, mbytInState
    
    mbytInState = Empty
    mstrInNO = Empty
    mblnNOMoved = False '查阅退出后清空,避免影响后续操作
    
    mlngDelRow = 0
    mlngUnitID = Empty
    mstrTime = ""
    mblnDelete = False
    gbytBilling = 0
    mbytUseType = 0
    mlngDeptID = 0
    mlng病人ID = 0
    mstr病人IDs = ""
    
    mlng药品类别ID = 0
    mlng卫材类别ID = 0
    
    Set mrs开单科室 = Nothing
    Set mrs开单人 = Nothing
    Set mrsWarn = Nothing
    Set mrsMedPayMode = Nothing
    
    If Not OS.IsDesinMode Then
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, glngOld)
    End If
End Sub

Private Sub lvwPati_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnDesc As Boolean
    Static intIdx As Integer
    
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvwPati.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvwPati.SortOrder = lvwDescending
    Else
        lvwPati.SortOrder = lvwAscending
    End If
    lvwPati.Sorted = True
    intIdx = ColumnHeader.Index
        
    If Not lvwPati.SelectedItem Is Nothing Then lvwPati.SelectedItem.EnsureVisible
End Sub

Private Sub lvwPati_DblClick()
    If lvwPati.SelectedItem Is Nothing Then Exit Sub
    Bill.Text = "-" & Mid(lvwPati.SelectedItem.Key, 2)
    lvwPati.Visible = False
    Bill.SetFocus
    Call zlCommFun.PressKey(13)
End Sub

Private Sub lvwPati_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then lvwPati_DblClick
End Sub

Private Sub lvwPati_LostFocus()
    lvwPati.Visible = False
End Sub

 

Private Sub sta_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If gbln简码切换 Then    '35242
        If Panel.Bevel = sbrRaised And (Panel.Key = "PY" Or Panel.Key = "WB") Then
            '切换并保存简码匹配方式
            Panel.Bevel = IIf(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
            If Panel.Key = "PY" Then
                sta.Panels("WB").Bevel = IIf(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
            Else
                sta.Panels("PY").Bevel = IIf(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
            End If
            zlDatabase.SetPara "简码方式", IIf(sta.Panels("PY").Bevel = sbrInset And sta.Panels("WB").Bevel = sbrInset, 2, IIf(sta.Panels("WB").Bevel = sbrInset, 1, 0))
            gbytCode = Val(zlDatabase.GetPara("简码方式", , , 0))
        End If
    End If
End Sub

Private Sub tmrStatuPati_Timer()
  If picStatuPancl.Visible Then Call MoveStatuPatiInfor
 
End Sub

Private Sub txtDate_GotFocus()
    zlControl.TxtSelAll txtDate
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And IsDate(txtDate.Text) Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtDate_LostFocus()
    txtDate.SelLength = 0
    If IsDate(txtDate.Text) Then mobjBill.发生时间 = CDate(txtDate.Text)
End Sub

Private Sub txtIn_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim tmpBill As New ExpenseBill
    Dim i As Long, strSQL As String
    Dim strInfo As String
    Dim Curdate As Date     '服务器当前时间
    
    On Error GoTo errH
    
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))

    '第一位可以输入字母,其它位不行
    If KeyAscii <> 13 Then
        Call SetNOInputLimit(txtIn, KeyAscii)
    Else
        txtIn.Text = GetFullNO(txtIn.Text, 14)
        
        '导入留观病人权限
        strInfo = Check留观病人(txtIn.Text, mstrPrivsOpt)
        If strInfo <> "" Then
            MsgBox "单据中包含" & strInfo & ",你没有权限对该单据进行操作！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        Set tmpBill = ImportBill(txtIn.Text, True, Me, False, gbln住院单位, False, mlngUnitID, , mstr药品价格等级, mstr卫材价格等级, mstr普通价格等级)
        If tmpBill.NO = "" Then
            MsgBox "读取单据失败。", vbExclamation, gstrSysName
            zlControl.TxtSelAll txtIn: txtIn.SetFocus
        Else
            '单据修改及显示
            Screen.MousePointer = 11
            Set mobjBill = New ExpenseBill
            Set mobjBill = tmpBill
            
            Call ReCalcInsure '重新计算统筹金额
            
            Curdate = zlDatabase.Currentdate
            mobjBill.NO = cboNO.Text
            mobjBill.登记时间 = Curdate
            mobjBill.操作员编号 = UserInfo.编号
            mobjBill.操作员姓名 = UserInfo.姓名
            mobjBill.加班标志 = chk加班.Value
            If InStr(mstrPrivsOpt, ";医生查询;") = 0 Then mobjBill.开单人 = ""
            
            '取当前时间
            txtDate.Text = Format(Curdate, "yyyy-MM-dd HH:mm:ss")
            
            Call Set开单人开单科室(cbo开单人, cbo开单科室, mrs开单人, mrs开单科室, mobjBill.开单人, mobjBill.开单部门ID)
            
            Bill.Redraw = False
            Bill.ClearBill
            Bill.Rows = mobjBill.Details.Count + 1
            
            Call InitBillColumnColor
            
            If gintPriceGradeStartType < 2 Then
                Call CalcMoneys
            Else
                '每一行根据价格等级重算价格
                For i = 1 To mobjBill.Details.Count
                    Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, _
                        mobjBill.Details(i).病人ID, mobjBill.Details(i).主页ID, mobjBill.Details(i).医疗付款, _
                        mstr药品价格等级, mstr卫材价格等级, mstr普通价格等级)
                    Call CalcMoneys(i)
                Next
            End If
            Call ShowDetails
            Call SetIntureColor
            
            Bill.Redraw = True
            chkIn.Value = 0
            Screen.MousePointer = 0
            
            '记帐分类报警
            mstrWarn = ""
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboNO_GotFocus()
    zlControl.TxtSelAll cboNO
    
    If gbytBilling = 2 Or chkCancel.Value = Checked Then
        cboNO.Locked = False
    Else
        cboNO.Locked = True
    End If
End Sub

Private Sub cboNO_KeyPress(KeyAscii As Integer)
    Dim blnRead As Boolean, strOper As String
    Dim vDate As Date, intTmp As Integer, str病人IDs As String
    Dim strInfo As String, intInsure As Integer
    Dim strInsure As String, arrInsure As Variant
    Dim i As Long, blnFlagPrint As Boolean
    
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))

    '第一位可以输入字母,其它位不行
    If KeyAscii <> 13 Then
        Call SetNOInputLimit(cboNO, KeyAscii)
    End If
    
    If KeyAscii = 13 And cboNO.Text <> "" And Not cboNO.Locked Then
        cboNO.Text = GetFullNO(cboNO.Text, 14)
        
        If chkCancel.Value = 1 Then
            '销帐
            
            '是否已转入后备数据表中
            If gbytBilling = 0 Then
                If zlDatabase.NOMoved("住院费用记录", cboNO.Text, , 2, Me.Caption) Then
                    If Not ReturnMovedExes(cboNO.Text, 2, Me.Caption) Then Exit Sub
                    mblnNOMoved = False
                End If
            End If
            
            '多次审核或不完全审核的不允许销帐
            If Not BillIdentical(cboNO.Text) Then
                MsgBox "单据中包含部份不全完审核或分多次审核的内容，不允许在这里销帐。" & _
                    vbCrLf & "请退回管理界面过滤出相应的单据内容，然后再销帐。", vbInformation, gstrSysName
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
        
            '单据权限
            If Not ReadBillInfo(2, cboNO.Text, 2, strOper, vDate) Then
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
            If mbytUseType = 0 And InStr(mstrPrivs, ";所有操作员;") <= 0 Then
                If UserInfo.姓名 <> strOper Then
                    MsgBox "你没有""所有操作员""权限,不能对" & strOper & "的单据进行销帐!", vbInformation, gstrSysName
                    cboNO.Text = "": cboNO.SetFocus: Exit Sub
                End If
            End If
            If Not BillOperCheck(5, strOper, vDate, "销帐", cboNO.Text) Then
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
        
            '项目冲销权限
            If mbytUseType = 0 Or mbytUseType = 1 Then
                If Not CheckDelPriv(cboNO.Text, mstrPrivsOpt) Then
                    cboNO.Text = "": cboNO.SetFocus: Exit Sub
                End If
            End If
        
            '留观病人权限
            strInfo = Check留观病人(cboNO.Text, mstrPrivsOpt)
            If strInfo <> "" Then
                MsgBox "单据中包含" & strInfo & ",你没有权限对该单据进行操作！", vbInformation, gstrSysName
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
        
            '是否已执行
            intTmp = BillCanDelete(cboNO.Text, 2, True, , mstrPrivsOpt, blnFlagPrint)
            If intTmp <> 0 Then
                Select Case intTmp
                    Case 1 '该单据不存在
                        MsgBox "指定单据中的内容不存在,或者你没有相关收费项目的销帐权限！", vbInformation, gstrSysName
                    Case 2 '已经全部完全执行
                        MsgBox "指定单据中的内容已经全部完全执行！", vbInformation, gstrSysName
                    Case 3 '未完全执行部分剩余数量为0
                        MsgBox "指定单据中的内容未完全执行部分项目剩余数量为零,没有可以销帐的费用！", vbInformation, gstrSysName
                End Select
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
            If blnFlagPrint Then
                If MsgBox("注意:检验医嘱的条码已打印，是否继续？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                    cboNO.Text = "": cboNO.SetFocus: Exit Sub
                End If
            End If
                        
            '出院病人操作权限判断
            If Not BillCanBeOperate(cboNO.Text, mstrPrivsOpt, "销帐", , str病人IDs) Then
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
                        
            '是否已结帐判断
            intTmp = HaveBilling(2, cboNO.Text, False)
            If intTmp <> 0 Then
                Call GetBillInsures(strInsure, cboNO.Text, , , True)
                If strInsure <> "" Then
                    arrInsure = Split(strInsure, ",")
                    For i = 0 To UBound(arrInsure)
                        If arrInsure(i) <> 0 Then
                            If Not gclsInsure.GetCapability(support允许冲销已结帐的记帐单据, , arrInsure(i)) Then
                                '医保病人的单据,固定为已结帐的禁止销帐
                                If intTmp = 1 Then
                                    MsgBox "该医保记帐单据未销帐部分已经结帐,不能销帐！", vbInformation, gstrSysName
                                    cboNO.Text = "": cboNO.SetFocus: Exit Sub
                                Else
                                    MsgBox "该医保记帐单据包含已经结帐的内容,只能对未结帐部分进行销帐！", vbInformation, gstrSysName
                                End If
                            End If
                        Else
                            Select Case gbytBillOpt
                                Case 0
                                Case 1
                                    If MsgBox("该记帐单据中存在已经结帐的内容,要销帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                        cboNO.Text = "": cboNO.SetFocus: Exit Sub
                                    End If
                                Case 2
                                    If intTmp = 1 Then
                                        MsgBox "该记帐单据未销帐部分已经结帐,不能销帐！", vbInformation, gstrSysName
                                        cboNO.Text = "": cboNO.SetFocus: Exit Sub
                                    Else
                                        MsgBox "该记帐单据包含已经结帐的内容,只能对未结帐部分进行销帐！", vbInformation, gstrSysName
                                    End If
                            End Select
                        End If
                    Next
                End If
            End If
            
            intInsure = BillExistInsure(cboNO.Text) '判断是否含有医保病人记的帐,记帐表检查其中只要有医保病人
            '医保销帐不允许对负数记录进行销帐
            If intInsure <> 0 Then
                If CheckNONegative(cboNO.Text) Then
                    MsgBox "该单据存在负数记帐记录,不允许进行医保销帐操作！", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
                        
            '是否存在重算冲减记录
            If CheckRecalcRecord(cboNO.Text) Then
                MsgBox "发现该记帐单据存在按费别重算的打折冲减记录!" & vbCrLf & _
                    "结帐前请按费别重算费用，否则病人将享受已销帐单据的打折优惠金额！", vbInformation, Me.Caption
            End If
        ElseIf mobjBill.Details.Count = 0 Then
            '记帐划价单(记帐审核)
            
            If Not BillExistMoney(cboNO.Text, 2) Then
                MsgBox "该单据费用已经全部销帐或单据不存在！", vbInformation, gstrSysName
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
            
            '出院病人操作权限判断
            If Not BillCanBeOperate(cboNO.Text, mstrPrivsOpt, "审核", , str病人IDs) Then
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
        End If
        
        mstr病人IDs = str病人IDs
        
        If chkCancel.Value = 1 Then '读取销帐单
            blnRead = ReadBill(cboNO.Text, True)
        ElseIf mobjBill.Details.Count = 0 Then '读取住院划价单
            blnRead = ReadBill(cboNO.Text, False)
        End If
        
        If blnRead Then
            
            mstrInNO = cboNO.Text '确定时以mstrInNO为准
            If chkCancel.Value = 0 Then '划价单
                Bill.Active = False
            Else '销帐
                'Call SetDisible 'cboNO在获取焦点后unLock
                Bill.Active = True
            End If
            cmdOK.SetFocus
        Else
            mstrInNO = "": cboNO.Text = "": cboNO.SetFocus
        End If
    End If
End Sub

Private Sub bill_AfterAddRow(Row As Long)
    Dim lngColor As Long, i As Long
    Dim lngRow As Long, lngCol As Long
    
    If mbytInState = 3 Or (chkCancel.Visible And chkCancel.Value = 1) Then
        'Bill.RemoveMSFItem Row'用属性AllowAddRow代替
        Bill.Row = 1: Call zlCommFun.PressKey(vbKeyTab)
        Exit Sub
    End If
    
     '自动复制上一行的病人信息
    If Row > 1 Then
        Bill.Redraw = False
        lngRow = Bill.Row: lngCol = Bill.Col
        
        Bill.Col = BillCol.姓名: Bill.Row = Row - 1
        lngColor = Bill.MsfObj.CellForeColor
        
        Bill.Col = BillCol.姓名: Bill.Row = Row
        Bill.MsfObj.CellForeColor = lngColor
        
        Bill.Row = lngRow: Bill.Col = lngCol
        Bill.Redraw = True
        
        Bill.TextMatrix(Row, BillCol.姓名) = Bill.TextMatrix(Row - 1, BillCol.姓名)
        Bill.TextMatrix(Row, BillCol.性别) = Bill.TextMatrix(Row - 1, BillCol.性别)
        Bill.TextMatrix(Row, BillCol.年龄) = Bill.TextMatrix(Row - 1, BillCol.年龄)
        Bill.TextMatrix(Row, BillCol.床号) = Bill.TextMatrix(Row - 1, BillCol.床号)
        '缺省费别
        Bill.TextMatrix(Row, BillCol.费别) = Bill.TextMatrix(Row - 1, BillCol.费别)
    End If
    
    With Bill
        '新增行时,重新设置可能已经被更改的可变性质列的列值
        If mbytInState <> 2 Then
            .ColData(BillCol.姓名) = BillColType.Text      '如果定位到主从项,会改变
            .ColData(BillCol.类别) = IIf(gbln收费类别 And Not mblnOne, BillColType.ComboBox, BillColType.UnFocus)
            .ColData(BillCol.项目) = BillColType.CommandButton
            .ColData(BillCol.付数) = BillColType.UnFocus '付数缺省跳过(=1),当类别为中药时,设为输入(4)(有值,一改全改)
            .ColData(BillCol.单价) = BillColType.UnFocus '单价缺省跳过,当项目变价时,设为输入(4)
            .ColData(BillCol.标志) = BillColType.UnFocus '标志缺省跳过,当为手术时,设为复选(-1)
        End If
        
        '针对列编辑性质设置颜色
        
        .SetColColor BillCol.类别, &HE7CFBA
        .SetColColor BillCol.项目, &HE7CFBA
        .SetColColor BillCol.数次, &HE7CFBA
        .SetColColor BillCol.执行科室, &HE7CFBA
        
        .SetColColor BillCol.付数, &HE0E0E0
        .SetColColor BillCol.单价, &HE0E0E0
        .SetColColor BillCol.标志, &HE0E0E0
    End With
    
   
    
    On Error Resume Next
    Bill.Text = "": Bill.SetFocus
    
    Set mrsInfo = New ADODB.Recordset: mstrUseMoney = "" '当前病人信息无效(以免被其它行误用)
End Sub

Private Sub cbo开单科室_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long, lng医生ID As Long
    
    If KeyAscii <> 13 Then Exit Sub
    If cbo开单科室.ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    
    If cbo开单人.ListIndex >= 0 Then lng医生ID = cbo开单人.ItemData(cbo开单人.ListIndex)
    If mrs开单科室 Is Nothing Then Call FillDept(cbo开单科室, mrs开单科室, mrs开单人, mstrPrivs, mbytUseType, mlngDeptID, lng医生ID)
    
    If zlSelectDept(Me, mlngModule, cbo开单科室, mrs开单科室, cbo开单科室.Text) = False Then
        Call Beep: mobjBill.开单部门ID = 0
        KeyAscii = 0: Exit Sub
    End If
    mobjBill.开单部门ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
    Exit Sub




'
'
'
'
'
'    Dim lngIdx As Long
'
'    If KeyAscii >= 32 And Not cbo开单科室.Locked Then
'        lngIdx = zlControl.CboMatchIndex(cbo开单科室.hwnd, KeyAscii)
'        If lngIdx = -1 And cbo开单科室.ListCount > 0 Then lngIdx = 0
'        cbo开单科室.ListIndex = lngIdx
'
'    ElseIf KeyAscii = 13 Then
'        If cbo开单科室.ListIndex = -1 Then
'            Beep
'        Else
'            mobjBill.开单部门ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
'            Call zlcommfun.PressKey(vbKeyTab)
'        End If
'    End If
End Sub
Private Function isCheck开单人Exists(ByVal str姓名 As String, Optional blnLocateItem As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查姓名是否在开单人下拉列表中.
    '入参:str姓名-姓名
    '     blnLocateItem:是否直接定位
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-07-20 17:53:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To cbo开单人.ListCount - 1
        If zlStr.NeedName(cbo开单人.List(i)) = str姓名 Then
            If blnLocateItem Then cbo开单人.ListIndex = i
            isCheck开单人Exists = True
            Exit Function
        End If
    Next
End Function
Private Sub cbo开单人_KeyPress(KeyAscii As Integer)
    Dim i As Long, intIdx As Integer, rsTemp As ADODB.Recordset, iCount As Integer
    Dim strText As String, strResult As String, strFilter As String
    
    If KeyAscii = 13 Then
        If cbo开单人.Locked Then
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        strText = UCase(cbo开单人.Text)
        If cbo开单人.ListIndex <> -1 Then
            '弹出列表时,又在文本框输入了内容
            If strText <> cbo开单人.List(cbo开单人.ListIndex) Then Call zlControl.CboSetIndex(cbo开单人.hWnd, -1)
        End If
        If strText = "" Then
            cbo开单人.ListIndex = -1
        ElseIf cbo开单人.ListIndex = -1 Then
            intIdx = -1
            strFilter = IIf(gbln护士, "人员性质<>''", "人员性质<>'护士'")
            '刘兴洪:22383
            '先复制记录集
            Set rsTemp = zlDatabase.zlCopyDataStructure(mrs开单人)
            Dim intInputType As Integer '0-输入的是全数字,1-输入的是全字母,2-其他
            Dim strCompents As String '匹配串
            
            strCompents = Replace(gstrLike, "%", "*") & strText & "*"
            
            If IsNumeric(strText) Then
                intInputType = 0
            ElseIf zlCommFun.IsCharAlpha(strText) Then
                intInputType = 1
            Else
                intInputType = 2
            End If
            
            mrs开单人.Filter = strFilter: iCount = 0
            With mrs开单人
                If .RecordCount <> 0 Then .MoveFirst
                Do While Not mrs开单人.EOF
                    Select Case intInputType
                    Case 0  '输入的是全数字
                        '1.编号输入值相等,主要输入如:12 匹配000012这种况,但如果输入的是01与编号01相等,则直接定位到01,则不定位在1上.
                        '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                        '主要是检查输入的内容与编号完全相同,则直接就定位到该姓名
                        If Nvl(!编号) = strText Then strResult = Nvl(!姓名): iCount = 0: Exit Do
                        
                        '1.编号输入值相等,主要输入如:12 匹配000012这种情况,因为这种情况有很多:如0012,012,000012等.因此如果存在此种情况,需要弹出选择器供选择
                        If Val(Nvl(!编号)) = Val(strText) Then
                            If iCount = 0 Then strResult = Nvl(!姓名)
                            iCount = iCount + 1
                        End If
                        
                        '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                         If Val(Nvl(!编号)) Like strText & "*" Then
                            If isCheck开单人Exists(Nvl(!姓名)) Then Call zlDatabase.zlInsertCurrRowData(mrs开单人, rsTemp)
                         End If
                    Case 1  '输入的是全字母
                        '规则:
                        ' 1.输入的简码相等,则直接定位
                        ' 2.根据参数来匹配相同数据
                        
                        '1.输入的简码相等,则直接定位
                        If Trim(Nvl(!简码)) = strText Then
                            If iCount = 0 Then strResult = Nvl(!姓名)   '可能存在多个相同的多个
                            iCount = iCount + 1
                        End If
                        
                        '2.根据参数来匹配相同数据
                        If Trim(Nvl(!简码)) Like strCompents Then
                            If isCheck开单人Exists(Nvl(!姓名)) Then Call zlDatabase.zlInsertCurrRowData(mrs开单人, rsTemp)
                        End If
                    Case Else  ' 2-其他
                        '规则:可能存在汉字等情况,或编号类似于N001简码可能有ZYK01这种情况
                        '1.编码\简码相等,直接定位
                        '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                        
                        '1.编码\简码相等,直接定位
                        If Trim(!编号) = strText Or Trim(!简码) = strText Or Trim(!姓名) = strText Then
                            If iCount = 0 Then strResult = Nvl(!姓名)   '可能存在多个相同的多个
                            iCount = iCount + 1
                        End If
                        
                        '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                        If Trim(!编号) Like strText & "*" Or Trim(Nvl(!简码)) Like strCompents Or Trim(Nvl(!姓名)) Like strCompents Then
                            If isCheck开单人Exists(Nvl(!姓名)) Then Call zlDatabase.zlInsertCurrRowData(mrs开单人, rsTemp)
                        End If
                    End Select
                    mrs开单人.MoveNext
                Loop
            End With
            If iCount > 1 Then strResult = ""
            If strResult = "" And rsTemp.RecordCount = 1 Then strResult = Nvl(rsTemp!姓名)
            '刘兴洪:直接定位
            If strResult <> "" Then
                rsTemp.Close: Set rsTemp = Nothing
                If isCheck开单人Exists(strResult, True) Then zlCommFun.PressKey vbKeyTab
                Exit Sub
            End If
            
            '需要检查是否有多条满足条件的记录
            If rsTemp.RecordCount <> 0 Then
                '先按某种方式进行排序
                Select Case intInputType
                Case 0 '输入全数字
                    rsTemp.Sort = "编号"
                Case 1 '输入全拼音
                    rsTemp.Sort = "简码"
                Case Else
                    '根据选择来定
                    If gbyt开单人显示 = 1 Then '简码
                        rsTemp.Sort = "简码"
                    Else
                        rsTemp.Sort = "编号"
                    End If
                End Select
                '弹出选择器
                Dim rsReturn As ADODB.Recordset
                If zlDatabase.zlShowListSelect(Me, glngSys, 1133, cbo开单人, rsTemp, True, "", "缺省,职务,优先级别", rsReturn) Then
                    If Not rsReturn Is Nothing Then
                        If rsReturn.RecordCount <> 0 Then
                            '进行定位
                            If isCheck开单人Exists(Nvl(rsReturn!姓名), True) Then
                                'zlCommFun.PressKey vbKeyTab
                            End If
                        End If
                    End If
                End If
            Else
                '未找到
                rsTemp.Close: Set rsTemp = Nothing
                KeyAscii = 0: zlControl.TxtSelAll cbo开单人: Exit Sub
            End If
            rsTemp.Close: Set rsTemp = Nothing
            
        ElseIf Not mblnDrop Then
            '回车光标经过
            Call cbo开单人_Click
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        If cbo开单人.ListIndex = -1 Then
            cbo开单人.Text = ""
            mobjBill.开单人 = ""
            If gblnFromDr Then Exit Sub
        Else
            mobjBill.开单人 = zlStr.NeedName(cbo开单人.Text)
            If intIdx <> -1 And mblnDrop Then
                '弹出回车-强行激活Click
                Call cbo开单人_Click
            ElseIf intIdx <> cbo开单人.ListIndex And intIdx <> -1 Then
                '弹出让选择-自动激活Click
                cbo开单人.SetFocus
                Call zlCommFun.PressKey(vbKeyF4)
                Exit Sub
            ElseIf intIdx <> -1 Then
                '一次性输中-强行激活Click
                Call cbo开单人_Click
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            ShowHelp App.ProductName, Me.hWnd, Me.Name
        Case vbKeyF2
            If ActiveControl Is cbo开单人 Then Call cbo开单人_KeyPress(vbKeyReturn)
            If cmdOK.Enabled And cmdOK.Visible Then
                Call cmdOK.SetFocus
                Call cmdOK_Click
            End If
        Case vbKeyF3
            If chkIn.Visible And chkIn.Enabled Then chkIn.Value = IIf(chkIn.Value = 1, 0, 1)
        Case vbKeyF6 '清除当前单据内容,进入新单状态
            If mbytInState = 0 Then
                If MsgBox("确实要清除当前单据中的内容吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If

                If chkCancel.Value = Checked Then '退据单状态
                    Call ClearRows: Call Bill.ClearBill
                    chkCancel.Value = Unchecked
                    Call NewBill
                    Call SetDisible(True)
                    If Bill.Enabled Then Bill.SetFocus
                ElseIf Bill.Active Then '正常输入单据状态
                    Call ClearRows: Call Bill.ClearBill
                    Call NewBill   '保持原单据号
                    If Bill.Enabled Then Bill.SetFocus
                End If
            End If
        Case vbKeyF7 '切换输入法
            If gbln简码切换 Then
                If sta.Panels("WB").Visible And sta.Panels("PY").Visible Then
                    If sta.Panels("WB").Bevel = sbrRaised Then
                        Call sta_PanelClick(sta.Panels("WB"))
                    Else
                        Call sta_PanelClick(sta.Panels("PY"))
                    End If
                End If
            End If
        Case vbKeyF8 '退(自动激活事件)
            If chkCancel.Visible And chkCancel.Enabled Then chkCancel.Value = IIf(chkCancel.Value = Checked, Unchecked, Checked)
        Case vbKeyA, vbKeyR
            '全选，全清
            If Shift = vbCtrlMask Then
                If KeyCode = vbKeyA And cmdSelALL.Visible And cmdSelALL.Enabled Then
                    Call cmdSelAll_Click
                ElseIf KeyCode = vbKeyR And cmdClear.Visible And cmdClear.Enabled Then
                    Call cmdClear_Click
                End If
            End If
        Case vbKeyQ
            If Shift = vbCtrlMask Then Call LocateNewRow
        Case vbKeyEscape
            If Bill.TxtVisible Then
                Bill.Text = "": Bill.TxtVisible = False: Bill.SetFocus
            ElseIf lvwPati.Visible Then
                lvwPati.Visible = False
            Else
                Call cmdCancel_Click
            End If
    End Select
End Sub

Private Function InitData() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim Curdate As Date     '服务器当前时间
    On Error GoTo errH
    
    Curdate = zlDatabase.Currentdate
    '自动识别加班
    If mbytInState <> 2 And mstrInNO = "" Then
        If OverTime(Curdate) Then chk加班.Value = Checked
    End If
    
    '不同药房药品出库检查方式
    Set mcolStock1 = GetStockCheck(0)
    Set mcolStock2 = GetStockCheck(1)
    
    '可选费别
    strSQL = "Select 编码,名称,简码,Nvl(缺省标志,0) as 缺省 From 费别 Order by 编码"
    Set mrsLevel = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(mrsLevel, strSQL, Me.Caption)
    If mrsLevel.EOF Then
        MsgBox "没有初始化费别，请先到费别管理中进行设置！", vbInformation, gstrSysName
        Exit Function
    End If
        
    If Init开单人开单科室(cbo开单人, cbo开单科室, mrs开单人, mrs开单科室, mstrPrivs, mbytUseType, mlngDeptID) = False Then
        Exit Function
    End If
        
    If gstr收费类别 = "" Then
        strSQL = "Select 编码,名称 as 类别 from 收费项目类别 Where 编码<>'1' Order by 序号"
    Else
        strSQL = "" & _
        "   Select /*+ RULE */   A.编码,A.名称 as 类别 " & _
        "   From 收费项目类别 A," & _
        "          (Select Column_Value From Table(Cast(f_str2list([1]) As Zltools.t_strlist))) J " & _
        "   Where A.编码=J. Column_Value " & _
        "   Order by 序号"
    End If
    Set mrsClass = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Replace(gstr收费类别, "'", ""))
    
    If mrsClass.EOF Then
        MsgBox "没有设置可用的收费类别,请先在本地参数中设置！", vbInformation, gstrSysName
        Exit Function
    End If
    '当只有一种可选收费类别时,不用用户选择
    mblnOne = (mrsClass.RecordCount = 1)
    If InStr(gstr收费类别, "'5'") > 0 Or InStr(gstr收费类别, "'6'") > 0 Or InStr(gstr收费类别, "'7'") > 0 Or gstr收费类别 = "" Then
        mlng药品类别ID = ExistIOClass(10)
        If mlng药品类别ID = 0 Then
            MsgBox "不能确定处方单据的入出类别,请先到药品入出分类管理中设置！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If InStr(gstr收费类别, "'4'") > 0 Or gstr收费类别 = "" Then
        mlng卫材类别ID = ExistIOClass(42)
        If mlng卫材类别ID = 0 Then
            MsgBox "不能确定卫材单据的入出类别,请先到入出分类管理中设置！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '执行部门
    strSQL = _
        "Select Distinct A.ID,A.编码,A.简码,A.名称,B.工作性质,B.服务对象 " & _
        "From 部门表 A,部门性质说明 B " & _
        "Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        "and B.部门ID=A.ID and B.服务对象 IN(2,3) " & _
        "Order by B.服务对象,A.编码"
    Set mrsUnit = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(mrsUnit, strSQL, Me.Caption)
    If mrsUnit.EOF Then
        MsgBox "没有初始化部门信息,单据无法处理执行部门。请先到部门管理中设置！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '开单日期
    txtDate.Text = Format(Curdate, "yyyy-MM-dd HH:mm:ss")
    
    If mbytInState = 0 Then Set mrsWarn = GetUnitWarn
    Set mrsInfo = New ADODB.Recordset
    
    InitData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetLastDeptID(ByVal str类别 As String, ByVal lngRow As Long, ByVal strDeptIDs As String) As Long
'功能：获取最近输入的相同类别项目的执行科室ID
    Dim i As Long
    
    For i = lngRow - 1 To 1 Step -1
        If mobjBill.Details(i).收费类别 = str类别 _
            And mobjBill.Details(i).执行部门ID <> 0 Then
            If InStr("," & strDeptIDs & ",", "," & mobjBill.Details(i).执行部门ID & ",") > 0 Then
                GetLastDeptID = mobjBill.Details(i).执行部门ID
                Exit Function
            End If
        End If
    Next
    
    '如果是卫生材料,再取与最近其它类别相匹配的执行科室
    If str类别 = "4" Then
        For i = lngRow - 1 To 1 Step -1
            If mobjBill.Details(i).执行部门ID <> 0 Then
                If InStr("," & strDeptIDs & ",", "," & mobjBill.Details(i).执行部门ID & ",") > 0 Then
                    GetLastDeptID = mobjBill.Details(i).执行部门ID
                    Exit Function
                End If
            End If
        Next
    End If
End Function

Private Sub FillBillComboBox(ByVal lngRow As Long, ByVal lngCol As Long, _
    Optional ByVal lng科室ID As Long, Optional ByVal int来源 As Integer, Optional blnEnter As Boolean)
'功能：根据单据列设置下拉列表框内容
'参数：blnEnter=是否按进入该列处理,比如执行科室保持不变
    Dim rsTmp As ADODB.Recordset, i As Long
    Dim strSQL As String, strIDs As String
    Dim bln护士 As Boolean, strTmp As String
    Dim lng病区ID As Long, lngListIndex As Long
    
    
    On Error GoTo errHandle
    
    Bill.Clear '清除列表框内容
    Select Case Bill.TextMatrix(0, lngCol)
        Case "费别"
            mrsLevel.Filter = adFilterNone
            If mrsLevel.RecordCount <> 0 Then
                For i = 1 To mrsLevel.RecordCount
                    Bill.AddItem mrsLevel!编码 & "-" & mrsLevel!名称
                    mrsLevel.MoveNext
                Next
            End If
             Bill.cboStyle = DropOlnyDown
        Case "类别"
            Call GetOperatorInfo(mrs开单人, mobjBill.开单人, bln护士)
        
            mrsClass.Filter = adFilterNone
            If mrsClass.RecordCount <> 0 Then
                mrsClass.MoveFirst
                For i = 1 To mrsClass.RecordCount
                    If Not (bln护士 And InStr(",E,M,4,", mrsClass!编码) = 0) Then
                        Bill.AddItem Bill.ListCount + 1 & "-" & mrsClass!类别
                        Bill.ItemData(Bill.NewIndex) = Asc(mrsClass!编码)  '存放类别编码的ASCII码
                    End If
                    mrsClass.MoveNext
                Next
            End If
             Bill.cboStyle = DropOlnyDown
        Case "执行科室"
             Bill.cboStyle = DropDownAndEdit
            '根据当前项目执行科室性质,动态设置可选科室
            If mobjBill.Details.Count >= lngRow Then
                With mobjBill.Details(lngRow)
                    If InStr(",4,5,6,7,", .收费类别) > 0 And .收费类别 <> "" Then
                        Call GetWorkUnit(.收费细目ID, .收费类别)
                        If mrsWork.RecordCount > 0 Then
                            '取上一个药的药房
                            mrsWork.MoveFirst
                            For i = 1 To mrsWork.RecordCount
                                strIDs = strIDs & "," & mrsWork!ID
                                mrsWork.MoveNext
                            Next
                            If Not blnEnter Then '进入该列时保持已确定值不变
                                lng科室ID = GetLastDeptID(.收费类别, lngRow, Mid(strIDs, 2))
                            End If
                            If lng科室ID = 0 Then lng科室ID = .执行部门ID
                                                    
                            '确定当前行的药房
                            mrsWork.MoveFirst
                            For i = 1 To mrsWork.RecordCount
                                Bill.AddItem IIf(zlIsShowDeptCode, mrsWork!编码 & "-", "") & mrsWork!名称
                                Bill.ItemData(Bill.NewIndex) = mrsWork!ID
                                If mrsWork!ID = lng科室ID Then Bill.ListIndex = Bill.NewIndex
                                mrsWork.MoveNext
                            Next
                        End If
                    Else
                        Bill.TextMatrix(lngRow, lngCol) = ""
                        
                        If int来源 = 0 Then int来源 = Get病人来源(lngRow)
                        
                        If lng科室ID = 0 Then
                            lng科室ID = .科室ID
                            If lng科室ID = 0 Then lng科室ID = Get开单科室ID
                        End If
                        
                        lng病区ID = .病区ID
                        If lng病区ID = 0 Then lng病区ID = Get病区ID(lng科室ID)
                        If lng病区ID = 0 Then lng病区ID = lng科室ID
                        
                        '0-不明确,1-病人科室,2-病人病区,3-操作员科室,4-指定科室,5-院外执行(预留,程序暂未用),6-开单人科室
                        Select Case .Detail.执行科室
                            Case 0 '不明确
                                mrsUnit.Filter = 0
                            Case 1 '病人科室
                                mrsUnit.Filter = "ID=" & lng科室ID & " Or ID=" & .执行部门ID
                            Case 2 '病人病区
                                mrsUnit.Filter = "ID=" & lng病区ID & " Or ID=" & .执行部门ID
                            Case 3 '操作员科室
                                mrsUnit.Filter = "ID=" & UserInfo.部门ID & " Or ID=" & .执行部门ID
                            Case 4 '指定科室
                                strSQL = "" & _
                                "   Select Nvl(A.开单科室ID,0) as 开单科室ID,A.执行科室ID" & _
                                "   From 收费执行科室 A,部门表 C" & _
                                "   Where A.收费细目ID=[1]　And A.执行科室ID+0=C.ID " & _
                                "       And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                                "       And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null) " & vbNewLine & _
                                "       And (A.病人来源 is NULL Or A.病人来源=[2])" & _
                                "       And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
                                " Order by Decode(A.病人来源,Null,2,1)" '默认科室优先
                                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .收费细目ID, int来源, lng科室ID)
                                If Not rsTmp.EOF Then
                                    For i = 1 To rsTmp.RecordCount
                                        strTmp = strTmp & "ID=" & rsTmp!执行科室ID & " OR "
                                        rsTmp.MoveNext
                                    Next
                                    strTmp = strTmp & "ID=" & .执行部门ID & " OR "
                                    strTmp = Left(strTmp, Len(strTmp) - 4)
                                    mrsUnit.Filter = strTmp
                                Else
                                    mrsUnit.Filter = "ID=" & UserInfo.部门ID & " Or ID=" & .执行部门ID
                                End If
                            Case 5 '院外执行(预留,程序暂未用)
                            Case 6 '开单人科室
                               mrsUnit.Filter = "ID=" & Get开单科室ID & " Or ID=" & .执行部门ID
                        End Select
                        If mrsUnit.EOF Then mrsUnit.Filter = "ID=" & UserInfo.部门ID & " Or ID=" & .执行部门ID
                        If Not mrsUnit.EOF Then
                            mrsUnit.MoveFirst: lngListIndex = -1
                            For i = 1 To mrsUnit.RecordCount
                                strTmp = IIf(zlIsShowDeptCode, mrsUnit!编码 & "-", "") & mrsUnit!名称
                                '刘兴洪:28947
                                If zlCboFindItem(Bill.cboObj, Val(Nvl(mrsUnit!ID))) = False Then
                                'If Not (SendMessage(Bill.cboHwnd, CB_FINDSTRING, -1, ByVal strTmp) >= 0) Then
                                    Bill.AddItem strTmp
                                    Bill.ItemData(Bill.ListCount - 1) = mrsUnit!ID
                                    
                                    '设置缺省执行科室
                                    If Not blnEnter Then '进入该列时保持已确定值不变
                                        If lngRow = 1 Then
                                            If mrsUnit!ID = lng科室ID Then lngListIndex = Bill.NewIndex
                                        ElseIf lngRow > 1 Then
                                            '与上一行非药品相同
                                            If mrsUnit!ID = mobjBill.Details(lngRow - 1).执行部门ID And mobjBill.Details(lngRow - 1).Detail.执行科室 = .Detail.执行科室 _
                                                And InStr(",5,6,7,", mobjBill.Details(lngRow - 1).收费类别) = 0 Then
                                                lngListIndex = Bill.NewIndex
                                            ElseIf mrsUnit!ID = lng科室ID And Bill.ListIndex = -1 Then
                                               lngListIndex = Bill.NewIndex
                                            End If
                                        End If
                                    End If
                                End If
                                mrsUnit.MoveNext
                            Next
                            '28378 由于要触发Bill_CboClick事件,因此,不能将Bill.Listindex放在循环(因为事件中包含了对mrsUnit的过滤处理，会造成记录集出错)
                            If lngListIndex >= 0 Then Bill.ListIndex = lngListIndex
                        End If
                            
                        If Not blnEnter And .Detail.执行科室 = 4 Then    '执行科室为指定科室的,缺省为操作员所在科室
                            For i = 0 To Bill.ListCount - 1
                                If Bill.ItemData(i) = UserInfo.部门ID Then Bill.ListIndex = i: Exit For
                            Next
                        End If
                        If Bill.ListIndex = -1 Then '如果没有则取现有的执行科室
                            For i = 0 To Bill.ListCount - 1
                                If Bill.ItemData(i) = .执行部门ID Then Bill.ListIndex = i: Exit For
                            Next
                        End If
                        
                        If Bill.ListIndex = -1 And Bill.ListCount > 0 Then Bill.ListIndex = 0
                    End If
                    
                    If Bill.ListIndex <> -1 Then
                        .执行部门ID = Bill.ItemData(Bill.ListIndex)
                        Bill.TextMatrix(lngRow, lngCol) = Bill.List(Bill.ListIndex)
                    Else
                        .执行部门ID = 0
                    End If
                End With
            End If
    End Select
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub InitFace()
'功能：根据表单要完成的功能设置界面布局
    Dim arrHead() As String, i As Long
    
    '公用单据表格式
    With Bill
        .LocateCol = BillCol.姓名 '缺省定位到病人栏
        .PrimaryCol = BillCol.姓名 '病人栏为主栏
        
        arrHead = Split(STR_HEAD, ";")
        .Cols = UBound(arrHead) + 1
        For i = 0 To UBound(arrHead)
            .TextMatrix(0, i) = Split(arrHead(i), ",")(0)
            .ColWidth(i) = Split(arrHead(i), ",")(1)
            .ColAlignment(i) = Split(arrHead(i), ",")(2)
        Next
                
        .MsfObj.MergeCells = flexMergeRestrictRows
        .MsfObj.MergeRow(0) = True
                
        If mbytInState = 0 And gbytBilling <> 2 Then
            .ColData(BillCol.姓名) = BillColType.Text  '病人输入！！！
            .ColData(BillCol.性别) = BillColType.UnFocus   '性别跳过
            .ColData(BillCol.年龄) = BillColType.UnFocus  '年龄跳过
            .ColData(BillCol.床号) = BillColType.UnFocus  '床号跳过
            .ColData(BillCol.费别) = BillColType.UnFocus  '费别跳过
            
            .ColData(BillCol.类别) = IIf(gbln收费类别 And Not mblnOne, BillColType.ComboBox, BillColType.UnFocus)
            
            .ColData(BillCol.项目) = 1  '项目输入,按扭可选
            '刘兴洪:27990 2010-02-22 17:00:04
            .ColData(BillCol.商品名) = BillColType.UnFocus  '商品名跳过
            .ColData(BillCol.规格) = BillColType.UnFocus  '规格跳过
            .ColData(BillCol.单位) = BillColType.UnFocus  '单位跳过
            .ColData(BillCol.付数) = BillColType.UnFocus '付数缺省跳过(=1),当类别为中药时,设为输入(4)(有值,一改全改)
            .ColData(BillCol.数次) = BillColType.Text   '数/次输入
            .ColData(BillCol.单价) = BillColType.UnFocus  '单价缺省跳过,当项目变价时,设为输入(4)
            .ColData(BillCol.应收金额) = BillColType.UnFocus  '应收金额跳过
            .ColData(BillCol.实收金额) = BillColType.UnFocus  '实收金额跳过
            .ColData(BillCol.执行科室) = BillColType.ComboBox   '默认取开单科室或上一科室
            .ColData(BillCol.标志) = BillColType.UnFocus  '标志缺省跳过,当为手术时,设为复选(-1)
            .ColData(BillCol.类型) = BillColType.UnFocus
        End If
        
                
        .SetColColor BillCol.类别, &HE7CFBA
        .SetColColor BillCol.项目, &HE7CFBA
        .SetColColor BillCol.数次, &HE7CFBA
        .SetColColor BillCol.执行科室, &HE7CFBA
        
        .SetColColor BillCol.付数, &HE0E0E0
        .SetColColor BillCol.单价, &HE0E0E0
        .SetColColor BillCol.标志, &HE0E0E0
    
        ReDim marrColData(.Cols - 1)
        For i = 0 To .Cols - 1
            marrColData(i) = .ColData(i)
        Next
        
        If mbytInState = 3 Then .AllowAddRow = False
    End With
    Call RestoreFlexState(Bill, App.ProductName & "\" & Me.Name & mbytInState)
    If gTy_System_Para.byt药品名称显示 <> 2 Then
        '0-显示通用名，1-显示商品名，2-同时显示通用名和商品名
        Bill.ColWidth(BillCol.商品名) = 0
    Else
        If Bill.ColWidth(BillCol.商品名) = 0 Then
             Bill.ColWidth(BillCol.商品名) = GetOrigColWidth(BillCol.商品名)
        End If
    End If
    '读取简码匹配方式
    sta.Panels("MedicareType").Visible = mbytInState = 0
    sta.Panels("PY").Visible = mbytInState = 0 And gbln简码切换 '35242
    sta.Panels("WB").Visible = mbytInState = 0 And gbln简码切换
    If mbytInState = 0 Then
        '简码匹配方式：0-拼音,1-五笔,2-两者
        If gbytCode = 0 Then
            sta.Panels("PY").Bevel = sbrInset
            sta.Panels("WB").Bevel = sbrRaised
        ElseIf gbytCode = 1 Then
            sta.Panels("PY").Bevel = sbrRaised
            sta.Panels("WB").Bevel = sbrInset
        Else
            sta.Panels("PY").Bevel = sbrInset
            sta.Panels("WB").Bevel = sbrInset
        End If
    End If
    
    '标题
    Select Case gbytBilling
        Case 0
            lblTitle.Caption = gstrUnitName & "住院记帐表"
        Case 1
            lblTitle.Caption = gstrUnitName & "住院记帐表(划价)"
        Case 2
            lblTitle.Caption = gstrUnitName & "住院记帐表(审核)"
    End Select
    
    txt单据.Text = gstrDec: txt病人.Text = gstrDec
    
    If mbytInState = 0 And (gbytBilling = 0 Or gbytBilling = 1) Then
        chkIn.Visible = True
        txtIn.Visible = True
    Else
        txt单据.Left = Val(txt单据.Tag) - chkIn.Width - txtIn.Width
        lbl单据.Left = txt单据.Left - lbl单据.Width - 45
        txt病人.Left = Val(txt病人.Tag) - chkIn.Width - txtIn.Width
        lbl病人.Left = txt病人.Left - lbl病人.Width - 45
    End If
    
    Select Case mbytInState
        Case 0 '执行
            '55380
            If mstrInNO <> "" Or _
                (InStr(mstrPrivsOpt, ";药品销帐;") = 0 _
                And InStr(mstrPrivsOpt, ";卫材销帐;") = 0 _
                And InStr(mstrPrivsOpt, ";诊疗销帐;") = 0) Then
                chkCancel.Visible = False
            End If
            Select Case gbytBilling
                Case 0, 1 '执行记帐、划价
                    Call SetShowCol
                Case 2 '执行审核
                    Call SetDisible
                    cboNO.Locked = False
                    picUnit.Enabled = False
                    fraAppend.Enabled = False
            End Select
        Case 1 '查阅
            Call SetDisible
            chkCancel.Visible = False
            If mblnDelete Then lblFlag.Visible = True
            fraTitle.Enabled = False
            picUnit.Enabled = False
            fraAppend.Enabled = False
            
            cmdOK.Visible = False
            cmdCancel.Caption = "退出(&X)"
        Case 2 '调整
            Call SetDisible
            txtDate.Enabled = True
            chkCancel.Visible = False
            fraTitle.Enabled = False
            picUnit.Enabled = False
            
        Case 3 '销帐
            Call SetDisible
            chkCancel.Visible = False
            fraTitle.Enabled = False
            picUnit.Enabled = False
            fraAppend.Enabled = False
            
            Call ShowDeleteCol(True)
            Bill.Active = True
    End Select
    
    '交换开单科室与开单人位置
    If gblnFromDr Then
        Call ExChangeLocate(cbo开单科室, cbo开单人)
        Call ExChangeLocate(lbl开单科室, lbl开单人)
        cbo开单科室.TabStop = False
    End If
End Sub

Private Sub SetDisible(Optional bln As Boolean = False)
'界面设置为不可修改状态
    cboNO.Locked = Not bln
    cbo开单科室.Locked = Not bln
    chk加班.Enabled = bln
    cboBaby.Enabled = bln
    cbo开单人.Locked = Not bln
    txtDate.Enabled = bln
    Bill.Active = bln
End Sub

Private Function GetPatient(ByVal strInput As String, ByVal blnCard As Boolean, Optional blnOutMsg As Boolean = False) As Boolean
    '功能：获取病人信息
    '参数：blnCard=是否就诊卡刷卡
    '出参:blnOutMsg-已经提示,过程外不用再调Msgbox
    Dim strSQL As String, lng卡类别ID As Long, strPassWord As String, strErrMsg As String, bln所有病区 As Boolean
    Dim strPati As String, strIF As String, strWhere As String
    Dim lng病人ID As Long, blnHavePassWord As Boolean
    Dim rsOutSel As ADODB.Recordset
    mstrUseMoney = ""
    On Error GoTo errH
        
    'a.是否具有强制记帐权限
    If InStr(mstrPrivsOpt, ";出院未结强制记帐;") > 0 And InStr(mstrPrivsOpt, ";出院结清强制记帐;") > 0 Then
        strIF = ""
    ElseIf InStr(mstrPrivsOpt, ";出院未结强制记帐;") > 0 Then
        strIF = " And ((B.出院日期 is NULL And Nvl(B.状态,0)<>3) Or Nvl(X.费用余额,0)<>0)"
    ElseIf InStr(mstrPrivsOpt, ";出院结清强制记帐;") > 0 Then
        strIF = " And ((B.出院日期 is NULL And Nvl(B.状态,0)<>3) Or Nvl(X.费用余额,0)=0)"
    Else
        strIF = " And B.出院日期 is NULL And Nvl(B.状态,0)<>3"
    End If
    
    'b.是否可以记所有病区病人
    bln所有病区 = True
    If (mbytUseType = 0 Or mbytUseType = 1) And InStr(mstrPrivs, ";所有病区;") <= 0 Then
        bln所有病区 = False
        If InStr(1, mstrUnitIDs, ",") = 0 Then
            strIF = strIF & " And B.当前病区ID+0=[3]"
        Else
            strIF = strIF & " And B.当前病区ID+0 IN (Select Column_Value From Table(Cast(f_num2list([4]) As zlTools.t_numlist)))"
        End If
    End If
       
    'c.是否留观病人记帐权限
    If (InStr(mstrPrivsOpt, ";门诊留观记帐;") > 0 And gbln门诊留观) And (InStr(mstrPrivsOpt, ";住院留观记帐;") > 0 And gbln住院留观) Then
        strIF = strIF & " And Nvl(B.病人性质,0) IN(0,1,2)"
    ElseIf InStr(mstrPrivsOpt, ";门诊留观记帐;") > 0 And gbln门诊留观 Then
        strIF = strIF & " And Nvl(B.病人性质,0) IN(0,1)"
    ElseIf InStr(mstrPrivsOpt, ";住院留观记帐;") > 0 And gbln住院留观 Then
        strIF = strIF & " And Nvl(B.病人性质,0) IN(0,2)"
    Else
        strIF = strIF & " And Nvl(B.病人性质,0)=0"
    End If
    
    strSQL = _
            "Select A.病人ID,B.主页ID,B.当前病区ID as 病区ID,B.出院科室ID as 科室ID,B.入院日期,B.出院日期," & _
            "   A.就诊卡号,A.卡验证码,A.住院号,B.出院病床 as 床号,X.费用余额,B.状态," & _
            "   nvl(B.姓名,A.姓名) as 姓名,nvl(b.性别,A.性别) as 性别,A.年龄,B.费别,B.住院医师,B.医疗付款方式," & _
            "   A.担保人,Decode(A.担保额,null,A.担保额,Zl_Patientsurety(A.病人ID,B.主页ID)) 担保额,zl_PatiDayCharge(A.病人ID) as 当日额," & _
            "   B.险类,Nvl(B.病人性质,0) as 病人性质,B.病人类型,B.审核标志" & _
            " From 病人信息 A,病案主页 B,病人余额 X" & _
            " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID" & _
            "       And Nvl(B.主页ID,0)<>0 And A.病人ID=X.病人ID(+) And X.性质(+)=1 And X.类型(+)=2 And A.停用时间 is NULL " & strIF
            
    If blnCard Then '就诊卡号
        strInput = UCase(strInput)
        strWhere = strWhere & " And A.就诊卡号=[2]"
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '病人ID
        strWhere = strWhere & " And A.病人ID=[1]"
    ElseIf Left(strInput, 1) = "/" Then   '床位号
        '41654 And IsNumeric(Mid(strInput, 2))
        strInput = Mid(strInput, 2)
        If mlngUnitID = 0 Then '病区不确定、则不能通过床号确定病人
            Set mrsInfo = New ADODB.Recordset: Exit Function
        End If
        strSQL = _
            "Select A.病人ID,B.主页ID,B.当前病区ID as 病区ID,B.出院科室ID as 科室ID,B.入院日期,B.出院日期," & _
            "   A.就诊卡号,A.卡验证码,A.住院号,B.出院病床 as 床号,X.费用余额,B.状态," & _
            "   nvl(B.姓名,A.姓名) as 姓名,nvl(b.性别,A.性别) as 性别,A.年龄,B.费别,B.住院医师,B.医疗付款方式," & _
            "   A.担保人,Decode(A.担保额,null,A.担保额,Zl_Patientsurety(A.病人ID,B.主页ID)) 担保额,zl_PatiDayCharge(A.病人ID) as 当日额," & _
            "   B.险类,Nvl(B.病人性质,0) as 病人性质,B.病人类型,B.审核标志" & _
            "   From 病人信息 A,病案主页 B,床位状况记录 C,病人余额 X" & _
            " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID" & _
            " And Nvl(B.主页ID,0)<>0 And A.病人ID=C.病人ID And A.病人ID=X.病人ID(+) And X.性质(+)=1 And X.类型(+)=2 And A.停用时间 is NULL" & _
            " And C.病区ID=[3] And C.床号=[2] " & strIF
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then '住院号(病人在院)
        strWhere = strWhere & " And A.病人ID = (Select Max(病人id) From 病案主页 Where 住院号 = [1])"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '门诊号(医技记帐)
        strWhere = strWhere & " And A.门诊号=[1]"
    Else '当作姓名
        If zlSelectChargePatiFromInputName(Me, mstrPrivsOpt, strInput, bln所有病区, mstrUnitIDs, gintOutDay, lng病人ID, strErrMsg, Bill.TxtHwnd, Bill.RowHeight(Bill.Row)) = False Then
            If strErrMsg = "" Then blnOutMsg = True: Set mrsInfo = New Recordset: Exit Function
            MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
            blnOutMsg = True: Set mrsInfo = New Recordset: Exit Function
        End If
        strInput = "-" & lng病人ID
        strWhere = strWhere & " And A.病人ID=[1]"
    End If
    
    strSQL = strSQL & vbCrLf & strWhere
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Mid(strInput, 2)), strInput, mlngUnitID, mstrUnitIDs)
    
    If Not mrsInfo.EOF Then
        '去掉了医保连接匹配检查
        If zlPatiIS病案已编目(Val(Nvl(mrsInfo!病人ID)), Val(Nvl(mrsInfo!主页ID))) = True Then      '问题:28725
            Set mrsInfo = New ADODB.Recordset
            blnOutMsg = True
            Exit Function
        End If
        If zlIsAllowFeeChange(Val(Nvl(mrsInfo!病人ID)), Val(Nvl(mrsInfo!主页ID)), Val(Nvl(mrsInfo!审核标志))) = False Then
            Set mrsInfo = New ADODB.Recordset
            blnOutMsg = True
            Exit Function
        End If
        sta.Panels(2) = ""
        If cbo开单科室.ListIndex <> -1 Then
            If mrsInfo!科室ID <> cbo开单科室.ItemData(cbo开单科室.ListIndex) And mrsInfo!病区ID <> cbo开单科室.ItemData(cbo开单科室.ListIndex) Then
                MsgBox "提醒：该住院病人不属于""" & zlStr.NeedName(cbo开单科室.Text) & """！", vbInformation, gstrSysName
            End If
        End If
        
        '读取价格等级
        If gintPriceGradeStartType >= 2 Then
            Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, _
                Val(Nvl(mrsInfo!病人ID)), Val(Nvl(mrsInfo!主页ID)), Nvl(mrsInfo!医疗付款方式), _
                mstr药品价格等级, mstr卫材价格等级, mstr普通价格等级)
        End If
        
        '审批信息在输入收费项目每次读取,因为每行的病人可能不同，并且可能修改任意已有行
        GetPatient = True
        Exit Function
    End If
    
    Set mrsInfo = New ADODB.Recordset
    
    If strWhere = "" Then Exit Function '无其他条件，直接退出
    
    '未找到病人，需要对该病人的具体错误信息进行提示
    strSQL = _
    " Select A.病人ID,B.主页ID,B.当前病区ID as 病区ID,B.出院科室ID as 科室ID,a.在院,B.入院日期,B.出院日期,X.费用余额,B.状态, " & _
    "       nvl(B.姓名,A.姓名) as 姓名,nvl(b.性别,A.性别) as 性别,nvl(b.年龄,A.年龄) as 年龄,B.费别,Nvl(B.病人性质,0) as 病人性质,B.病人类型" & _
    " From 病人信息 A,病案主页 B,病人余额 X" & _
    " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID" & _
    "   And Nvl(B.主页ID,0)<>0 And A.病人ID=X.病人ID(+) and X.性质(+)=1 and X.类型(+)=2 And A.停用时间 is NULL " & strWhere
    
    Set rsOutSel = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Mid(strInput, 2)), strInput)
    If rsOutSel.EOF Then Exit Function
    
    '1.病区检查
    If (mbytUseType = 0 Or mbytUseType = 1) And InStr(mstrPrivs, ";所有病区;") <= 0 Then
        If InStr(1, "," & mstrUnitIDs & ",", "," & Val(rsOutSel!病区ID) & ",") = 0 Then
            MsgBox "病人:『" & Nvl(rsOutSel!姓名) & "』不在你负责的病区,不能对该病人进行记账操作!", vbInformation + vbOKOnly, gstrSysName
            blnOutMsg = True
            Exit Function
        End If
    End If
    
    '2.留观病人检查(是否留观病人记帐权限)
    If (InStr(mstrPrivsOpt, ";门诊留观记帐;") > 0 And gbln门诊留观) And (InStr(mstrPrivsOpt, ";住院留观记帐;") > 0 And gbln住院留观) Then
        '0-普通住院病人,1-门诊留观病人,2-住院留观病人
    ElseIf InStr(mstrPrivsOpt, ";门诊留观记帐;") > 0 And gbln门诊留观 Then
        If Val(Nvl(rsOutSel!病人性质)) = 2 Then
            MsgBox "病人:『" & Nvl(rsOutSel!姓名) & "』为住院留观病人,你不具备『住院留观记帐』权限,不能对该病人进行记账操作!", vbInformation + vbOKOnly, gstrSysName
            blnOutMsg = True
            Exit Function
        End If
    ElseIf InStr(mstrPrivsOpt, ";住院留观记帐;") > 0 And gbln住院留观 Then
        If Val(Nvl(rsOutSel!病人性质)) = 1 Then
            MsgBox "病人:『" & Nvl(rsOutSel!姓名) & "』为门诊留观病人,你不具备『门诊留观记帐』权限,不能对该病人进行记账操作!", vbInformation + vbOKOnly, gstrSysName
            blnOutMsg = True
            Exit Function
        End If
    Else
        If Val(Nvl(rsOutSel!病人性质)) <> 0 Then
            MsgBox "病人:『" & Nvl(rsOutSel!姓名) & "』为" & IIf(Val(Nvl(rsOutSel!病人性质)) = 1, "门诊", "住院") & "留观病人,你不具备『门诊或住院 留观记帐』权限,不能对该病人进行记账操作!", vbInformation + vbOKOnly, gstrSysName
            blnOutMsg = True
            Exit Function
        End If
    End If
    
    '124007
    If InStr(mstrPrivsOpt, ";出院未结强制记帐;") > 0 And InStr(mstrPrivsOpt, ";出院结清强制记帐;") > 0 Then
        strErrMsg = ""
    ElseIf InStr(mstrPrivsOpt, ";出院未结强制记帐;") > 0 Then
        If Not (Val(Nvl(rsOutSel!状态)) <> 3 And IsNull(rsOutSel!出院日期) Or Val(Nvl(rsOutSel!费用余额)) <> 0) Then
              
                If Val(Nvl(rsOutSel!状态)) = 3 And IsNull(rsOutSel!出院日期) Then
                    strErrMsg = "病人已经预出院，不能对病人进行记账操作!"
                Else
                    strErrMsg = "病人于" & Format(rsOutSel!出院日期, "yyyy年mm月DD日") & " 出院，不能对病人进行记账操作!"
                End If
        End If
    ElseIf InStr(mstrPrivsOpt, ";出院结清强制记帐;") > 0 Then
        If Not (Val(Nvl(rsOutSel!状态)) <> 3 And IsNull(rsOutSel!出院日期) Or Val(Nvl(rsOutSel!费用余额)) = 0) Then
                If Val(Nvl(rsOutSel!状态)) = 3 And IsNull(rsOutSel!出院日期) Then
                strErrMsg = "病人已经预出院，不能对病人进行记账操作!"
                Else
                strErrMsg = "病人于" & Format(rsOutSel!出院日期, "yyyy年mm月DD日") & " 出院，不能对病人进行记账操作!"
                End If
        End If
    Else
        If Not (Val(Nvl(rsOutSel!状态)) <> 3 And IsNull(rsOutSel!出院日期)) Then
            If Val(Nvl(rsOutSel!状态)) = 3 And IsNull(rsOutSel!出院日期) Then
                strErrMsg = "病人已经预出院，不能对病人进行记账操作!"
            Else
                strErrMsg = "病人于" & Format(rsOutSel!出院日期, "yyyy年mm月DD日") & " 出院，不能对病人进行记账操作!"
            End If
        End If
    End If
    
    If strErrMsg <> "" Then
        MsgBox strErrMsg, vbInformation, gstrSysName
        blnOutMsg = True
        Exit Function
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Set mrsInfo = New ADODB.Recordset
End Function

Private Sub CalcMoneys(Optional lngRow As Long = 0)
'功能：计算或重新计算指定行或所有行的金额
'参数：lngRow=指定行,为0表示计算所有行
'说明：ExpenseBill集合的索引对应单据的行号
    Dim i As Long, j As Long, k As Long
    Dim blnExist As Boolean
    
    Dim strMainRows As String
    Dim bln从项汇总折扣 As Boolean
    
    If mobjBill.Details.Count = 0 Then Exit Sub
    
    For i = IIf(lngRow = 0, 1, lngRow) To IIf(lngRow = 0, mobjBill.Details.Count, lngRow)
        
        bln从项汇总折扣 = False
        If gbln从项汇总折扣 Then                    '如果主项屏蔽费别,则汇总计算折扣参数无效,不汇总计算
            If mobjBill.Details(i).从属父号 > 0 Then    '从项
                bln从项汇总折扣 = Not mobjBill.Details(mobjBill.Details(i).从属父号).Detail.屏蔽费别
                If bln从项汇总折扣 And lngRow <> 0 Then strMainRows = "," & mobjBill.Details(i).从属父号      '单独计算一行的时候
            Else
                If CheckItemHaveSub(i) Then                          '主项或独立项
                     bln从项汇总折扣 = Not mobjBill.Details(i).Detail.屏蔽费别
                     If bln从项汇总折扣 Then strMainRows = strMainRows & "," & i  '一页可能有多个主从项,先记录主项行号,后面再重算主项折扣
                End If
            End If
        End If
                    
        Call CalcMoney(i, bln从项汇总折扣)
    Next
    
    '重算所有主项,不能用bln从项汇总折扣变量,因为可能在遇到不是从项的行时已改变
    If gbln从项汇总折扣 Then
        For i = 1 To UBound(Split(strMainRows, ","))
            Call Calc重算主项实收(Split(strMainRows, ",")(i))
        Next
    End If
    
    Set mcolMoneys = New BillInComes
    '产生汇总费目
    For i = 1 To mobjBill.Details.Count
        For j = 1 To mobjBill.Details(i).InComes.Count
            '查找是否已经加入此类收入项目,如是则合计,否则新入
            blnExist = False
            For k = 1 To mcolMoneys.Count
                If mcolMoneys(k).收入项目ID = mobjBill.Details(i).InComes(j).收入项目ID Then
                    blnExist = True: Exit For
                End If
            Next
            
            If blnExist Then
                mcolMoneys(k).实收金额 = mcolMoneys(k).实收金额 + mobjBill.Details(i).InComes(j).实收金额
            Else
                With mobjBill.Details(i).InComes(j)
                    mcolMoneys.Add .收入项目ID, .收入项目, .收据费目, .标准单价, .应收金额, .实收金额
                End With
            End If
        Next
    Next
End Sub

Private Sub CalcMoney(lngRow As Long, Optional bln从项汇总折扣 As Boolean)
'功能：计算或重新计算指定行的金额
'参数：lngRow=指定行
'说明：1.ExpenseBill集合的索引对应单据的行号
'      2.变价只能对应一个收入项目:mobjBill.Details(lngRow).InComes(1)
'      3.如果变价细目未计算出收入项目(第一次计算),则使用默认现价
'      4.如果变价细目已经计算出收入项目(按第2步),并手动更改(也可能未改)了单价,则按该单价计算。
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strInfo As String, i As Long
    Dim intInsure As Integer, dblMoney As Double '用户输入的变价金额
    
    Dim dblAllTime As Double, dbl加班加价率 As Double
    Dim rsPrice As ADODB.Recordset, strPrice As String, varPrice As Variant, dbl剩余数量 As Double
    Dim strPriceGrade As String, strWherePriceGrade As String
    
    On Error GoTo errH
    If InStr(",5,6,7,", mobjBill.Details(lngRow).收费类别) > 0 Then
        strPriceGrade = mstr药品价格等级
    ElseIf mobjBill.Details(lngRow).收费类别 = "4" Then
        strPriceGrade = mstr卫材价格等级
    Else
        strPriceGrade = mstr普通价格等级
    End If
    
    If InStr(",5,6,7,", mobjBill.Details(lngRow).收费类别) > 0 Then
        Call AdjustCpt(mobjBill.Details(lngRow).收费细目ID)
    End If
    
    If strPriceGrade <> "" Then
        strWherePriceGrade = _
            "       And (b.价格等级 = [2]" & vbNewLine & _
            "            Or (b.价格等级 Is Null" & vbNewLine & _
            "                And Not Exists(Select 1" & vbNewLine & _
            "                               From 收费价目" & vbNewLine & _
            "                               Where b.收费细目Id = 收费细目id And 价格等级 = [2]" & vbNewLine & _
            "                                     And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')))))"
    Else
        strWherePriceGrade = " And b.价格等级 Is Null"
    End If
    strSQL = _
        " Select B.收入项目ID,C.名称,C.收据费目,B.现价,B.原价,B.加班加价率,B.附术收费率,B.缺省价格 " & _
        " From 收费项目目录 A,收费价目 B,收入项目 C " & _
        " Where B.收费细目ID = A.ID And C.ID = B.收入项目ID " & _
        " And Sysdate Between B.执行日期 And Nvl(B.终止日期,To_Date('3000-1-1', 'YYYY-MM-DD')) " & _
        " And A.ID=[1]" & vbNewLine & _
        strWherePriceGrade
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjBill.Details(lngRow).收费细目ID, strPriceGrade)
    If rsTmp.EOF Then
        '如果没有收入项目,则清除对应的程序对象
        Set mobjBill.Details(lngRow).InComes = New BillInComes
        Exit Sub
    End If

    With mobjBill.Details(lngRow)
        '先获取操作员以前输入的变价金额
        If InStr(",5,6,7,", .收费类别) > 0 Or (.收费类别 = "4" And .Detail.跟踪在用) Then
            '计算药品时价(分批或不分批)
            '必然有记录(输入该项目时已判断)
            dblAllTime = .付数 * .数次
            If gbln住院单位 And InStr(",5,6,7,", .收费类别) > 0 Then
                dblAllTime = dblAllTime * .Detail.住院包装 '库存时价按售价数量进行计算
            End If
            If dblAllTime <> 0 Or Not .Detail.变价 Then
                Set rsPrice = zlDatabase.OpenSQLRecord("Select Zl_Fun_Getprice([1],[2],[3]) As Price From Dual", _
                            Me.Caption, .收费细目ID, .执行部门ID, dblAllTime)
                If rsPrice.EOF Then
                    '获取价格失败
                    If InStr(",5,6,7,", .收费类别) > 0 Then
                        MsgBox "第 " & lngRow & " 行药品""" & .Detail.名称 & """获取价格失败！", vbInformation, gstrSysName
                    Else
                        MsgBox "第 " & lngRow & " 行卫生材料""" & .Detail.名称 & """获取价格失败！", vbInformation, gstrSysName
                    End If
                Else
                    strPrice = Nvl(rsPrice!Price) & "|||"
                    varPrice = Split(strPrice, "|")
                    dblMoney = Val(varPrice(0))
                    dbl剩余数量 = Val(varPrice(2))
                    
                    If dbl剩余数量 <> 0 And .Detail.变价 Then
                        '数量未分解完毕
                        If InStr(",5,6,7,", .收费类别) > 0 Then
                            MsgBox "第 " & lngRow & " 行时价药品""" & .Detail.名称 & """库存不足,无法计算价格！", vbInformation, gstrSysName
                        Else
                            MsgBox "第 " & lngRow & " 行时价卫生材料""" & .Detail.名称 & """库存不足,无法计算价格！", vbInformation, gstrSysName
                        End If
                        dblMoney = 0
                    End If
                End If
            Else
                dblMoney = 0
            End If
        Else
            If .Detail.变价 Then
                If .InComes.Count = 0 Then '第一次计算金额取缺省值
                    dblMoney = Val(Nvl(rsTmp!缺省价格))
                Else                        '获取操作员以前输入的变价金额
                    dblMoney = .InComes(1).标准单价
                    '如果用户输入的变价不满足变价范围，则取缺省值
                    If CheckScope(Val(Nvl(rsTmp!原价)), Val(Nvl(rsTmp!现价)), dblMoney) <> "" Then
                        dblMoney = Val(Nvl(rsTmp!缺省价格))
                    End If
                End If
            End If
        End If
    End With
    
    '再清除原有记录
    Set mobjBill.Details(lngRow).InComes = New BillInComes
    
    '填写现有费用记录
    For i = 1 To rsTmp.RecordCount
        Set mobjBillIncome = New BillInCome
        With mobjBillIncome
            .收入项目ID = rsTmp!收入项目ID
            .收入项目 = rsTmp!名称
            .收据费目 = Nvl(rsTmp!收据费目)
            .原价 = Val(Nvl(rsTmp!原价))
            .现价 = Val(Nvl(rsTmp!现价))
            
            If InStr(",5,6,7,", mobjBill.Details(lngRow).收费类别) > 0 Then
                If gbln住院单位 Then
                    .标准单价 = Format(dblMoney * mobjBill.Details(lngRow).Detail.住院包装, gstrFeePrecisionFmt)
                Else
                    .标准单价 = Format(dblMoney, gstrFeePrecisionFmt)
                End If
            Else
                If mobjBill.Details(lngRow).Detail.变价 Then
                    .标准单价 = Format(dblMoney, gstrFeePrecisionFmt)
                Else
                    .标准单价 = Format(Nvl(rsTmp!现价, 0), gstrFeePrecisionFmt)
                End If
            End If
            
            '应收金额=单价 * 付数 * 数次
            .应收金额 = .标准单价 * IIf(mobjBill.Details(lngRow).付数 = 0, 1, mobjBill.Details(lngRow).付数) * mobjBill.Details(lngRow).数次
            
            '附加手术费率用计算(所有收入项目)
            If mobjBill.Details(lngRow).附加标志 = 1 And mobjBill.Details(lngRow).收费类别 = "F" Then
                .应收金额 = .应收金额 * IIf(IsNull(rsTmp!附术收费率), 1, rsTmp!附术收费率 / 100)
            End If
            
            '加班费用率计算
            dbl加班加价率 = 0
            If mobjBill.加班标志 = 1 And mobjBill.Details(lngRow).Detail.加班加价 Then
                dbl加班加价率 = IIf(IsNull(rsTmp!加班加价率), 0, rsTmp!加班加价率 / 100)
                .应收金额 = .应收金额 + .应收金额 * dbl加班加价率
            End If
            
            .应收金额 = CCur(Format(.应收金额, gstrDec))
            dblAllTime = mobjBill.Details(lngRow).付数 * mobjBill.Details(lngRow).数次
            If InStr(",5,6,7,", mobjBill.Details(lngRow).收费类别) > 0 Then
                If gbln住院单位 Then dblAllTime = dblAllTime * mobjBill.Details(lngRow).Detail.住院包装
            End If
            
            If mobjBill.Details(lngRow).Detail.屏蔽费别 Or bln从项汇总折扣 Or .应收金额 = 0 Then
                .实收金额 = .应收金额
            Else
                If .应收金额 = 0 Then
                    .实收金额 = 0
                    mobjBill.Details(lngRow).费别 = mobjBill.费别
                Else
                    '药品按成本价加收,传入数量
                    .实收金额 = CCur(Format(ActualMoney(mobjBill.Details(lngRow).费别, .收入项目ID, .应收金额, _
                        mobjBill.Details(lngRow).收费细目ID, mobjBill.Details(lngRow).执行部门ID, dblAllTime, dbl加班加价率), gstrDec))   '当前病人的费别
                End If
            End If
            
            '获取项目保险信息,医保病人才处理,不需要连接医保
            intInsure = GetPatiInsure(mobjBill.Details(lngRow).病人ID)
            If intInsure > 0 Then
                strInfo = gclsInsure.GetItemInsure(mobjBill.Details(lngRow).病人ID, mobjBill.Details(lngRow).收费细目ID, .实收金额, False, intInsure, _
                    mobjBill.Details(lngRow).摘要 & "||" & dblAllTime)
                If strInfo <> "" Then
                    mobjBill.Details(lngRow).保险项目否 = Val(Split(strInfo, ";")(0)) <> 0
                    mobjBill.Details(lngRow).保险大类ID = Val(Split(strInfo, ";")(1))
                    .统筹金额 = Format(Val(Split(strInfo, ";")(2)), gstrDec)
                    mobjBill.Details(lngRow).保险编码 = CStr(Split(strInfo, ";")(3))
                    
                    If UBound(Split(strInfo, ";")) >= 4 Then
                        If CStr(Split(strInfo, ";")(4)) <> "" Then mobjBill.Details(lngRow).摘要 = CStr(Split(strInfo, ";")(4))
                        If UBound(Split(strInfo, ";")) >= 5 Then
                            If Split(strInfo, ";")(5) <> "" Then mobjBill.Details(lngRow).Detail.类型 = Split(strInfo, ";")(5)
                        End If
                    End If
                End If
            End If
            
            mobjBill.Details(lngRow).InComes.Add .收入项目ID, .收入项目, .收据费目, .标准单价, .应收金额, .实收金额, .原价, .现价, , .统筹金额
        End With
        rsTmp.MoveNext
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetPatiInsure(Optional ByVal lng病人ID As Long) As Integer
'功能:得到病人险类
'参数：lng病人ID=不传时，取第一个医保病人的险类
    Dim i As Long
    For i = 1 To mobjBill.Details.Count
        If lng病人ID <> 0 Then
            If mobjBill.Details(i).病人ID = lng病人ID Then
                GetPatiInsure = Val(mobjBill.Details(i).发药窗口)
                Exit Function
            End If
        Else
            If Val(mobjBill.Details(i).发药窗口) > 0 Then
                GetPatiInsure = Val(mobjBill.Details(i).发药窗口)
                Exit Function
            End If
        End If
    Next
End Function

Private Function GetMultiInsures() As String
'功能:得到单据中包含的多个病人险类
    Dim strInsure As String, i As Long
    
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).病人ID <> 0 And Val(mobjBill.Details(i).发药窗口) <> 0 Then
            If InStr(strInsure & ",", "," & Val(mobjBill.Details(i).发药窗口) & ",") = 0 Then
                strInsure = strInsure & "," & Val(mobjBill.Details(i).发药窗口)
            End If
        End If
    Next
    GetMultiInsures = Mid(strInsure, 2)
End Function

Private Sub ShowDetails(Optional lngRow As Long = 0)
'功能：刷新显示指定行或所有行的内容
'参数：lngRow=指定行,为0表示显示所有行
'说明：ExpenseBill集合的索引对应单据的行号
    Dim i As Long

    Bill.Redraw = False
    If lngRow = 0 Then
        For i = 1 To mobjBill.Details.Count
            ShowDetail i
        Next
    Else
        ShowDetail lngRow
    End If
    Bill.Redraw = True
    
    txt单据.Text = Format(GetBillTotal(mobjBill), gstrDec)
End Sub

Private Sub ShowDetail(lngRow As Long)
'功能：刷新显示指定行的内容
'参数：lngRow=指定行
'说明：ExpenseBill集合的索引对应单据的行号
    Dim dbl单价 As Double, cur金额 As Currency
    Dim i As Long, j As Long
    
    If lngRow > Bill.Rows - 1 Then Exit Sub
    
    '清除单据行
    For i = 0 To Bill.Cols - 1
        '输入时收费类别不清除
        If i > 5 Then Bill.TextMatrix(lngRow, i) = ""
    Next
    
    Bill.RowData(lngRow) = Asc(mobjBill.Details(lngRow).收费类别)
    
    '刷新单据行
    For i = 0 To Bill.Cols - 1
        If i = 0 Then Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).姓名
        If i = 1 Then Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).性别
        If i = 2 Then Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).年龄
        Select Case Bill.TextMatrix(0, i)
            Case "床号"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).床号
            Case "费别"
                '浏览单据或从属项目只(能)显示名称
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).费别
            Case "类别"
                '浏览单据或从属项目只(能)显示名称
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.类别名称
            Case "项目"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.名称
            Case "规格"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.规格
            Case "商品名"   '刘兴洪:27990 2010-02-22 17:00:49
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.商品名
            Case "单位"
                If InStr(",5,6,7,", mobjBill.Details(lngRow).收费类别) > 0 And gbln住院单位 Then
                    Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.住院单位
                Else
                    Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.计算单位
                End If
            Case "付"
                Bill.TextMatrix(lngRow, i) = IIf(mobjBill.Details(lngRow).付数 = 0, 1, mobjBill.Details(lngRow).付数)
            Case "数次"
                '数次在第一次显示时已默认设置为1
                Bill.TextMatrix(lngRow, i) = FormatEx(mobjBill.Details(lngRow).数次, 5)
            Case "单价"
                '单价是该收费细目所有收入项目的合计
                '第一次计算时是在默认数次为1的基础上计算出来的
                dbl单价 = 0
                If mobjBill.Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Details(lngRow).InComes.Count
                        dbl单价 = dbl单价 + mobjBill.Details(lngRow).InComes(j).标准单价
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(dbl单价, gstrFeePrecisionFmt)
            Case "应收金额"
                '应收金额是该收费细目所有收入项目的合计
                cur金额 = 0
                If mobjBill.Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Details(lngRow).InComes.Count
                        cur金额 = cur金额 + mobjBill.Details(lngRow).InComes(j).应收金额
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(cur金额, gstrDec)
            Case "实收金额"
                '实收金额是该收费细目所有收入项目的合计
                cur金额 = 0
                If mobjBill.Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Details(lngRow).InComes.Count
                        cur金额 = cur金额 + mobjBill.Details(lngRow).InComes(j).实收金额
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(cur金额, gstrDec)
            Case "执行科室"
                If mobjBill.Details(lngRow).执行部门ID <> 0 Then
                    mrsUnit.Filter = "ID=" & mobjBill.Details(lngRow).执行部门ID
                    If mrsUnit.RecordCount <> 0 Then
                        If mbytInState = 0 Then
                            Bill.TextMatrix(lngRow, i) = mrsUnit!编码 & "-" & mrsUnit!名称
                        Else
                            '浏览单据只(能)显示名称
                            Bill.TextMatrix(lngRow, i) = mrsUnit!名称
                        End If
                    Else
                        Bill.TextMatrix(lngRow, i) = GET部门名称(mobjBill.Details(lngRow).执行部门ID, mrsUnit)
                    End If
                Else
                    Bill.TextMatrix(lngRow, i) = ""
                End If
            Case "标志"
                If mobjBill.Details(lngRow).收费类别 = "F" And mobjBill.Details(lngRow).附加标志 = 1 Then
                    Bill.TextMatrix(lngRow, i) = "√"
                End If
            Case "类型"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.类型
        End Select
    Next
    Bill.Text = Bill.MsfObj.Text
End Sub

Private Sub SetDetail(Detail As Detail, lngRow As Long, lngDoUnit As Long, Optional bytParent As Integer = 0)
'功能：根据指定的收费细目对象设定单据指点定行的收费细目(新增的或修改)
'说明：
'      1.用于新输入或更改收费细目行！！！
'      2.当bytParent<>0时,则为设置从属项目,从属项目一定是新增行,且主项目一定存在

    Dim tmpIncomes As New BillInComes
    Dim intPay As Integer, i As Long, dblTime As Double
    
    '取其它中药的付数
    intPay = 1
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).收费类别 = "7" And i <> lngRow Then
            intPay = mobjBill.Details(i).付数
            Exit For
        End If
    Next
    If Detail.类别 <> "7" Then intPay = 1
    
    If mobjBill.Details.Count < lngRow Then
        '如果该行对应的程序对象尚未初始,则加入
        With Detail
            '序号=行号,父号=0
            '付数=1
            '次数=1,从属项目的次数由主项计算确定
            '执行部门ID:根据细目执行科室标志取
            '附加标志:以第一行为假,其它为真优先权
            '收入集=空
            If bytParent <> 0 Then
                '设置该行RowData
                Bill.RowData(lngRow) = Asc(Detail.类别)
                '初始数次
                If Detail.固有从属 = 0 Then '非固有从属
                    dblTime = Detail.从项数次
                ElseIf Detail.固有从属 = 1 Then '固定的固有从属
                    dblTime = IIf(Detail.从项数次 = 0, 1, Detail.从项数次)
                ElseIf Detail.固有从属 = 2 Then '按比例的固有从属
                    dblTime = Detail.从项数次 * mobjBill.Details(bytParent).数次
                End If
            Else
                
                If InStr(",5,6,7,", Detail.类别) > 0 Then
                    dblTime = 0
                Else
                    dblTime = 1
                End If
            End If
            If bytParent <> 0 Then
                '刘兴洪:mobjBill.Details(bytParent).发药窗口 加入:
                '问题:
                mobjBill.Details.Add Detail, .ID, CInt(lngRow), bytParent, mobjBill.Details(bytParent).病人ID, _
                mobjBill.Details(bytParent).主页ID, mobjBill.Details(bytParent).病区ID, _
                mobjBill.Details(bytParent).科室ID, mobjBill.Details(bytParent).姓名, _
                mobjBill.Details(bytParent).性别, mobjBill.Details(bytParent).年龄, mobjBill.Details(bytParent).住院号, _
                mobjBill.Details(bytParent).床号, mobjBill.Details(bytParent).费别, mobjBill.Details(bytParent).病人性质, _
                .类别, .计算单位, mobjBill.Details(bytParent).发药窗口, intPay, dblTime, 0, lngDoUnit, tmpIncomes, mobjBill.Details(bytParent).就诊卡号, , mobjBill.Details(bytParent).担保额, _
                mobjBill.Details(bytParent).医疗付款
            Else
                mobjBill.Details.Add Detail, .ID, CInt(lngRow), bytParent, 0, 0, 0, 0, "", "", "", 0, 0, "", 0, _
                .类别, .计算单位, "", intPay, dblTime, 0, lngDoUnit, tmpIncomes
            End If
        End With
    Else '如果该行已经存在,则修改
        
        If InStr(",5,6,7,", Detail.类别) > 0 Then
            dblTime = 0
        Else
            dblTime = 1
        End If
        
        With mobjBill.Details(lngRow)
            Set .Detail = Detail
            Set .InComes = tmpIncomes
            .付数 = intPay
            .附加标志 = 0
            .计算单位 = Detail.计算单位
            .收费类别 = Detail.类别
            .收费细目ID = Detail.ID
            .数次 = dblTime
            .序号 = lngRow
            .从属父号 = 0
            .执行部门ID = lngDoUnit
        End With
    End If
End Sub

Private Function ShouldDO(lngRow As Long) As Boolean
'功能：判断该行是否应该取从属项目
'说明：仅该行收费项目有从属项目及尚未取才取。
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, blnExist As Boolean
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    strSQL = "Select count(从项ID) as NUM From 收费从属项目 Where 主项ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjBill.Details(lngRow).收费细目ID)
    
    If Not rsTmp.EOF Then
        If IsNull(rsTmp!Num) Then
            ShouldDO = False
        ElseIf rsTmp!Num = 0 Then
            ShouldDO = False
        Else
            blnExist = False
            For i = lngRow + 1 To mobjBill.Details.Count
                If mobjBill.Details(i).从属父号 = lngRow Then
                    blnExist = True: Exit For
                End If
            Next
            If Not blnExist Then
                ShouldDO = True
            Else
                ShouldDO = False
            End If
        End If
    Else
        ShouldDO = False
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetSubDetails(ByVal lng项目id As Long) As Details
'功能：返回一个收费细目的从属项目集
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long, lngMediCareNO As Long
    Dim objDetail As New Detail
    
    Set GetSubDetails = New Details
    
    If mrsInfo.State = 1 Then lngMediCareNO = Val("" & mrsInfo!险类)
    If lngMediCareNO > 0 Then
        strSQL = _
        "Select A.ID,Decode(A.类别,'4',E.诊疗ID,D.药名ID) as 药名ID,A.类别,B.名称 as 类别名称," & _
        "       A.费用类型,A.编码,Nvl(F.名称,A.名称) as 名称,E1.名称 as 商品名,A.规格,A.计算单位,A.屏蔽费别,G.要求审批," & _
        "       Decode(A.类别,'4',E.在用分批,D.药房分批) as 分批,A.是否变价," & _
        "       Decode(A.类别,'4',1,D.住院包装) as 住院包装,A.服务对象," & _
        "       Decode(A.类别,'4',A.计算单位,D.住院单位) as 住院单位," & _
        "       A.加班加价,A.执行科室,C.固有从属,C.从项数次,E.跟踪在用,D.中药形态" & _
        " From 收费项目目录 A,收费项目类别 B,收费从属项目 C,药品规格 D,材料特性 E,收费项目别名 F,收费项目别名 E1,保险支付项目 G" & _
        " Where B.编码=A.类别 And C.从项ID=A.ID And A.ID=D.药品ID(+) And A.ID=E.材料ID(+)" & _
        "       And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        "       And A.ID=F.收费细目ID(+) And F.码类(+)=1 And F.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
        "       And A.ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3" & _
        "       And C.主项ID=[1] And A.ID=G.收费细目ID(+) And G.险类(+)=[2] " & _
        " Order by 编码"
    Else
        strSQL = _
        " Select A.ID,Decode(A.类别,'4',E.诊疗ID,D.药名ID) as 药名ID,A.类别,B.名称 as 类别名称," & _
        "       A.费用类型,A.编码,Nvl(F.名称,A.名称) as 名称,E1.名称 as 商品名,A.规格,A.计算单位,A.屏蔽费别,0 as 要求审批," & _
        "       Decode(A.类别,'4',E.在用分批,D.药房分批) as 分批,A.是否变价," & _
        "       Decode(A.类别,'4',1,D.住院包装) as 住院包装,A.服务对象," & _
        "       Decode(A.类别,'4',A.计算单位,D.住院单位) as 住院单位," & _
        "       A.加班加价,A.执行科室,C.固有从属,C.从项数次,E.跟踪在用,D.中药形态" & _
        " From 收费项目目录 A,收费项目类别 B,收费从属项目 C,药品规格 D,材料特性 E,收费项目别名 F,收费项目别名 E1" & _
        " Where B.编码=A.类别 And C.从项ID=A.ID And A.ID=D.药品ID(+) And A.ID=E.材料ID(+)" & _
        "       And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        "       And A.ID=F.收费细目ID(+) And F.码类(+)=1 And F.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
        "       And A.ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3" & _
        "       And C.主项ID=[1]  " & _
        " Order by 编码"
    End If
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng项目id, lngMediCareNO)
    For i = 1 To rsTmp.RecordCount
        Set objDetail = New Detail
        With objDetail
            .ID = rsTmp!ID
            .药名ID = Nvl(rsTmp!药名ID, 0)
            .编码 = rsTmp!编码
            .变价 = Nvl(rsTmp!是否变价, 0) = 1
            .规格 = Nvl(rsTmp!规格)
            .住院包装 = Nvl(rsTmp!住院包装, 1)
            .住院单位 = Nvl(rsTmp!住院单位)
            .计算单位 = Nvl(rsTmp!计算单位)
            .分批 = Nvl(rsTmp!分批, 0) = 1
            .加班加价 = Nvl(rsTmp!加班加价, 0) = 1
            .类别 = rsTmp!类别
            .类别名称 = rsTmp!类别名称
            .名称 = rsTmp!名称
            .屏蔽费别 = Nvl(rsTmp!屏蔽费别, 0) = 1
            .执行科室 = Nvl(rsTmp!执行科室, 0)
            .服务对象 = Nvl(rsTmp!服务对象, 0)
            .固有从属 = Nvl(rsTmp!固有从属, 0)
            .从项数次 = Nvl(rsTmp!从项数次, 1)
            .类型 = Nvl(rsTmp!费用类型)
            .跟踪在用 = Nvl(rsTmp!跟踪在用, 0) = 1
            .要求审批 = Nvl(rsTmp!要求审批, 0) = 1
            .中药形态 = Val(Nvl(rsTmp!中药形态))
            .商品名 = Nvl(rsTmp!商品名)
            GetSubDetails.Add .ID, .药名ID, .类别, .类别名称, .名称, .编码, .简码, .别名, .规格, .计算单位, .说明, .屏蔽费别, _
                .住院包装, .住院单位, .分批, .变价, .加班加价, .执行科室, .服务对象, .类型, .补充摘要, .固有从属, .从项数次, .跟踪在用, , , , , , .要求审批, , .中药形态, .商品名
        End With
        rsTmp.MoveNext
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub DeleteDetail(lngRow As Long)
'功能：删除指定收费项目行
'说明：这时不处理从属行的删除,但要对其它单据行从属关系作相应的调整
    Dim i As Long
    For i = lngRow + 1 To mobjBill.Details.Count
        If mobjBill.Details(i).从属父号 <> 0 And mobjBill.Details(i).从属父号 > lngRow Then
            mobjBill.Details(i).从属父号 = mobjBill.Details(i).从属父号 - 1
        End If
        mobjBill.Details(i).序号 = mobjBill.Details(i).序号 - 1 '序号与行号对应
    Next
    mobjBill.Details.Remove lngRow
    If lngRow = 1 And mobjBill.Details.Count = 0 And Bill.Rows = 2 Then
        For i = 0 To Bill.Cols - 1
            Bill.TextMatrix(lngRow, i) = ""
        Next
    Else
        Bill.RemoveMSFItem lngRow
    End If
End Sub

Private Sub NewBill()
'功能：初始化一张新的单据(程序对象)
    Dim Curdate As Date     '服务器当前时间
    
    '记帐分类报警
    mstrWarn = ""
    Set mrsInfo = New ADODB.Recordset
    Set mobjBill = New ExpenseBill
    Set mcolPatiInfo = New Collection

    mstrUseMoney = "": sta.Panels(3).Text = "": picStatuPancl.Visible = False: lblStatuPati.Caption = ""
    mcurModiMoney = 0
    mlngPreRow = 0
    cboNO.Text = ""
    
    Call LoadPatientBaby(cboBaby, 0, 0)
    Call cbo开单科室_Click
    
    Curdate = zlDatabase.Currentdate
    chk加班.Value = IIf(OverTime(Curdate), 1, 0)
    txtDate.Text = Format(Curdate, "yyyy-MM-dd HH:mm:ss")
    txt单据.Text = gstrDec: txt病人.Text = gstrDec: lbl病人.Caption = "病人"
    
    With mobjBill
        .门诊标志 = 2
        .多病人单 = True
        .划价人 = UserInfo.姓名
        .开单人 = zlStr.NeedName(cbo开单人.Text)
        .操作员编号 = UserInfo.编号
        .操作员姓名 = UserInfo.姓名
        .发生时间 = CDate(txtDate.Text)
        .加班标志 = chk加班.Value
        .婴儿费 = 0
        If cbo开单科室.ListIndex = -1 Then
            .开单部门ID = 0
        Else
            .开单部门ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
        End If
        If cboDrawDept.ListIndex = -1 Then
            .领药部门ID = 0
        Else
            .领药部门ID = cboDrawDept.ItemData(cboDrawDept.ListIndex)
        End If
        
    End With
End Sub

Private Function SaveBill() As Boolean
'功能:保存当前输入的记帐单据(适用住院记帐、划价、或对两者的修改)
'入口:mobjBill=单据对象
'出口:保存是否成功
    Dim i As Long, j As Long, arrSQL As Variant, arrSMSQL As Variant
    Dim int序号 As Integer, int行号 As Integer, strNO As String, strTmp As String
    Dim intParent As Integer, intParentNO As Integer
    Dim dbl数次 As Double, dbl单价 As Double, str消息 As String, str汇总号 As String
    Dim strDelInsure As String, arrDelInsure As Variant
    Dim strInsure As String, arrInsure As Variant
    Dim blnModiBill As Boolean
    Dim strSQL As String, strStuffDept As String '记录卫料发料部门
    
    Dim strAddDate As String '记帐发生,自动发药,发料的时间
    Dim blnTrans As Boolean
    Dim rsTmp As ADODB.Recordset
    
    mobjBill.NO = zlDatabase.GetNextNo(14)
    strAddDate = "To_Date('" & Format(mobjBill.登记时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    
    
    Call zlReSetDrawDrugDept
    
    gstrModiNO = mobjBill.NO
    arrSQL = Array()
    arrSMSQL = Array()
    
    For Each mobjBillDetail In mobjBill.Details
        If mobjBillDetail.数次 <> 0 Then
            intParent = 0: intParentNO = int序号
            For Each mobjBillIncome In mobjBillDetail.InComes
                int序号 = int序号 + 1 '当前记录序号
                
                '单据主体及病人明细
                With mobjBill
                    gstrSQL = "zl_住院记帐记录_INSERT('" & .NO & "'," & int序号 & "," & mobjBillDetail.病人ID & "," & _
                        mobjBillDetail.主页ID & "," & mobjBillDetail.住院号 & "," & "'" & mobjBillDetail.姓名 & "','" & _
                        mobjBillDetail.性别 & "','" & mobjBillDetail.年龄 & "','" & mobjBillDetail.床号 & "','" & mobjBillDetail.费别 & "'," & _
                        IIf(mobjBillDetail.病区ID = 0, .开单部门ID, mobjBillDetail.病区ID) & "," & mobjBillDetail.科室ID & "," & .加班标志 & "," & _
                        mobjBillDetail.婴儿费 & "," & .开单部门ID & ",'" & .开单人 & "',"
                End With
                
                '收费细目部份
                With mobjBillDetail
                    '处理从属父号
                    If .序号 <> int行号 Then
                        int行号 = .序号
                        
                        '重新处理从属父号
                        If mobjBill.Details(.序号).从属父号 = 0 Then
                            For i = .序号 + 1 To mobjBill.Details.Count
                                If mobjBill.Details(i).从属父号 = .序号 Then
                                    mobjBill.Details(i).从属父号 = int序号 '当父项目有多个收入项目(多个序号)时,取第一个序号
                                End If
                            Next
                        End If
                    End If
                    gstrSQL = gstrSQL & .从属父号 & "," & .收费细目ID & ",'" & .收费类别 & "','" & .计算单位 & "',"
    
                    gstrSQL = gstrSQL & IIf(.保险项目否, 1, 0) & "," & IIf(.保险大类ID = 0, "NULL", .保险大类ID) & ",'" & .保险编码 & "',"
                    
                    dbl数次 = .数次
                    If InStr(",5,6,7,", .收费类别) > 0 And gbln住院单位 Then
                        dbl数次 = Format(.数次 * .Detail.住院包装, gstrFeePrecisionFmt)
                    End If
                    gstrSQL = gstrSQL & IIf(.付数 = 0, 1, .付数) & "," & dbl数次 & "," & .附加标志 & "," & IIf(.执行部门ID = 0, "NULL", .执行部门ID) & ","
                    
                    '收集卫料发料部门,以便自动发料
                    If gbytBilling = 0 And gint卫材发料控制 <> 0 Then
                        'gint卫材发料控制:0-不自动发料，1-自动发料，2-本科室开单时自动发料
                        If .执行部门ID <> 0 And .收费类别 = "4" And .Detail.跟踪在用 _
                            And ((gint卫材发料控制 = 2 And .执行部门ID = mobjBill.开单部门ID) Or gint卫材发料控制 = 1) Then
                            If InStr("," & strStuffDept, "," & .执行部门ID & ",") = 0 Then
                                strStuffDept = strStuffDept & "," & .执行部门ID
                            End If
                        End If
                    End If
                End With
                
                '收入项目部份
                With mobjBillIncome
                    intParent = intParent + 1
                    dbl单价 = .标准单价
                    If InStr(",5,6,7,", mobjBillDetail.收费类别) > 0 And gbln住院单位 Then
                        dbl单价 = Format(.标准单价 / mobjBillDetail.Detail.住院包装, gstrFeePrecisionFmt)
                    End If
                    gstrSQL = gstrSQL & IIf(intParent = 1, "Null", intParentNO + 1) & "," & .收入项目ID & "," & _
                        "'" & .收据费目 & "'," & dbl单价 & "," & .应收金额 & "," & .实收金额 & "," & _
                        IIf(.统筹金额 = 0, "NULL", .统筹金额) & ","
                End With
                                                
                '其它部分
                gstrSQL = gstrSQL & _
                    "To_Date('" & Format(mobjBill.发生时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & strAddDate & "," & _
                    "'" & mstrInNO & "'," & IIf(gbytBilling = 1, 1, 0) & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & _
                    "1," & IIf(mobjBillDetail.收费类别 = "4", mlng卫材类别ID, mlng药品类别ID) & ",Null,'" & mobjBillDetail.摘要 & "'," & _
                    "Null,Null,Null,Null,Null,Null,Null,Null,'" & mobjBillDetail.Detail.类型 & "',0," & mobjBill.领药部门ID & ")"
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = mobjBillDetail.收费细目ID & ";" & gstrSQL
            Next
        End If
    Next
    
    '修改前退除原单据
    If mstrInNO <> "" Then
        '先判断是否医保病人记的帐,并作合法性检查(进入修改时已作了一次相关判断)
        If gbytBilling = 0 Then
            '去掉了医保连接匹配检查
            Call GetBillInsures(strDelInsure, mstrInNO)
            If strDelInsure <> "" Then arrDelInsure = Split(strDelInsure, ",")
        End If
        
        gstrSQL = "zl_住院记帐记录_DELETE('" & mstrInNO & "',NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
        If gstrSQL <> "" Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "0;" & gstrSQL
        End If
    End If

    If UBound(arrSQL) >= 0 Then
        '对SQL序列按收费细目ID排序
        For i = 0 To UBound(arrSQL) - 1
            For j = i + 1 To UBound(arrSQL)
                If CLng(Split(arrSQL(j), ";")(0)) < CLng(Split(arrSQL(i), ";")(0)) Then
                    strTmp = CStr(arrSQL(j))
                    arrSQL(j) = arrSQL(i)
                    arrSQL(i) = strTmp
                End If
            Next
        Next
        
        '判断记帐表之中是否有医保病人(发药窗口记录险类)
        strInsure = GetMultiInsures
        If strInsure <> "" Then arrInsure = Split(strInsure, ",")

        '执行SQL语句
        On Error GoTo errH
        gcnOracle.BeginTrans
            blnTrans = True
            For i = 0 To UBound(arrSQL)
                Call zlDatabase.ExecuteProcedure(Mid(arrSQL(i), InStr(arrSQL(i), ";") + 1), Me.Caption)
            Next
            
            '执行自动发料
            If strStuffDept <> "" Then
                strStuffDept = Mid(strStuffDept, 2)
                For i = 0 To UBound(Split(strStuffDept, ","))
                    strSQL = "zl_材料收发记录_处方发料(" & Split(strStuffDept, ",")(i) & ",26,'" & mobjBill.NO & _
                        "','" & UserInfo.姓名 & "','" & UserInfo.姓名 & "','" & UserInfo.姓名 & "',1," & strAddDate & ")"
                    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                Next
            End If
            
            '准备自动发药(仅普通记帐),必须在事务中才能读到数据
            If mblnSendMateria Then
                Set rsTmp = Get待发药清单(mobjBill.NO, Format(mobjBill.登记时间, "yyyy-MM-dd HH:mm:ss"), True)
                If rsTmp.RecordCount > 0 Then
                    str汇总号 = zlDatabase.GetNextNo(20)
                    ReDim arrSMSQL(rsTmp.RecordCount - 1)
                    For i = 0 To rsTmp.RecordCount - 1
                        arrSMSQL(i) = "ZL_药品收发记录_部门发药(" & rsTmp!库房ID & "," & rsTmp!ID & ",'" & UserInfo.姓名 & "'," & strAddDate & ",Null,Null,Null," & str汇总号 & ")"
                        rsTmp.MoveNext
                    Next
                End If
                rsTmp.Close
            End If
            '执行自动发药
            For i = 0 To UBound(arrSMSQL)
                Call zlDatabase.ExecuteProcedure(CStr(arrSMSQL(i)), Me.Caption)
            Next
            
            '医保接口
            '1.医保记帐作废上传(只要第一个成功则提交)
            blnModiBill = False
            If mstrInNO <> "" And gbytBilling = 0 And strDelInsure <> "" Then
                For i = 0 To UBound(arrDelInsure)
                    If gclsInsure.GetCapability(support记帐作废上传, , arrDelInsure(i)) And Not gclsInsure.GetCapability(support记帐完成后上传, , arrDelInsure(i)) Then
                        If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , arrDelInsure(i)) Then
                            If i = 0 Then gcnOracle.RollbackTrans: Exit Function
                        Else
                            blnModiBill = True '表明销帐时成功上传过明细
                        End If
                    End If
                Next
            End If
            
            '2.记帐实时上传(只要第一个成功则提交)
            If gbytBilling = 0 And strInsure <> "" Then
                For i = 0 To UBound(arrInsure)
                    If gclsInsure.GetCapability(support记帐上传, , arrInsure(i)) And Not gclsInsure.GetCapability(support记帐完成后上传, , arrInsure(i)) Then
                        str消息 = ""
                        If Not gclsInsure.TranChargeDetail(2, mobjBill.NO, 2, 1, str消息, , arrInsure(i)) Then
                            '如果是修改,只要销帐时有成功上传则提交
                            If i = 0 And Not blnModiBill Then gcnOracle.RollbackTrans
                            If str消息 <> "" Then MsgBox str消息, vbInformation, gstrSysName
                            If i = 0 And Not blnModiBill Then Exit Function
                        End If
                    End If
                Next
            End If
        gcnOracle.CommitTrans
        blnTrans = False
        
        '1.医保记帐作废上传
        If mstrInNO <> "" And gbytBilling = 0 And strDelInsure <> "" Then
            For i = 0 To UBound(arrDelInsure)
                If gclsInsure.GetCapability(support记帐作废上传, , arrDelInsure(i)) And gclsInsure.GetCapability(support记帐完成后上传, , arrDelInsure(i)) Then
                    If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , arrDelInsure(i)) Then
                        MsgBox "单据中的 " & GetInsureName(Val(arrDelInsure(i))) & " 销帐费用向医保传送失败,这些费用已销帐！", vbInformation, gstrSysName
                    End If
                End If
            Next
        End If
        
        '2.记帐实时上传
        If gbytBilling = 0 And strInsure <> "" Then
            For i = 0 To UBound(arrInsure)
                If gclsInsure.GetCapability(support记帐上传, , arrInsure(i)) And gclsInsure.GetCapability(support记帐完成后上传, , arrInsure(i)) Then
                    str消息 = ""
                    If Not gclsInsure.TranChargeDetail(2, mobjBill.NO, 2, 1, str消息, , arrInsure(i)) Then
                        If str消息 <> "" Then
                            MsgBox str消息, vbInformation, gstrSysName
                        Else
                            MsgBox "单据中 " & GetInsureName(Val(arrInsure(i))) & " 的费用向医保传送失败,这些费用已保存！", vbInformation, gstrSysName
                        End If
                    End If
                End If
            Next
        End If
        
        '加入单据历史记录(所有类型单据)
        For i = 0 To cboNO.ListCount - 1
            strNO = strNO & "," & cboNO.List(i)
        Next
        strNO = mobjBill.NO & strNO
        cboNO.Clear
        For i = 0 To UBound(Split(strNO, ","))
            cboNO.AddItem Split(strNO, ",")(i)
            If i = 9 Then Exit For '只显示10个
        Next
        
        If str消息 <> "" Then MsgBox str消息, vbInformation, gstrSysName
    End If
    
    SaveBill = True
    Exit Function
errH:
    If Err.Description Like "*当前计算单价不一致*" Then
       If blnTrans Then gcnOracle.RollbackTrans
       
       If MsgBox("某些分批药品价格已发生变化，要自动重算价格吗？", vbYesNo + vbQuestion + vbDefaultButton1, App.ProductName) = vbYes Then
           Call CalcMoneys
           Call ShowDetails
           Exit Function
       End If
    Else
        If blnTrans Then gcnOracle.RollbackTrans
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
End Function

Private Function ReadBill(ByVal strNO As String, Optional blnDelete As Boolean) As Boolean
'功能：根据单据号读取一张单据并将其填入表格
'参数：strNO=单据号
'      blnDelete=True:销帐单据时调用,False:查阅单据时调用
    Dim rsTmp As ADODB.Recordset
    Dim curTotal As Currency, blnDo As Boolean, arrInsure As Variant
    Dim i As Long, lng病区ID As Long, intSign As Integer
    Dim strSQL As String, strSQL1 As String, strSQL2 As String, strInsure As String, strFeeKind As String, strUserUnitIDs As String
        
    On Error GoTo errH
    
    mblnPrint = False
    
     '读单据之前已检查,至少有一种销帐权限
    If blnDelete Then
        '55380
        Dim blnYP As Boolean, blnZL As Boolean, blnWC As Boolean
        blnYP = zlStr.IsHavePrivs(mstrPrivsOpt, "药品销帐")
        blnZL = zlStr.IsHavePrivs(mstrPrivsOpt, "诊疗销帐")
        blnWC = zlStr.IsHavePrivs(mstrPrivsOpt, "卫材销帐")
        If blnYP And blnWC And blnZL Then
            '所有,不限制
        ElseIf blnYP And blnWC And Not blnZL Then
            strFeeKind = " And 收费类别   In('4','5','6','7')"
        ElseIf blnYP And Not blnWC And blnZL Then
            strFeeKind = " And 收费类别   <>'4'"
        ElseIf blnYP And Not blnWC And Not blnZL Then
            strFeeKind = " And 收费类别 In('5','6','7')"
        ElseIf Not blnYP And blnWC And blnZL Then
            strFeeKind = " And 收费类别 Not In('5','6','7')"
        ElseIf Not blnYP And Not blnWC And blnZL Then
            strFeeKind = " And 收费类别 Not In('4','5','6','7')"
        ElseIf Not blnYP And blnWC And Not blnZL Then
            strFeeKind = " And 收费类别 ='4'"
        End If
    End If
    
    Call ClearRows: Call Bill.ClearBill: mlngPreRow = 0 '重新读取单据时，初始化行号标志
    
    '读取单据主体
    strNO = GetFullNO(strNO, 14)
    
    strSQL = _
    " Select A.开单部门ID,Nvl(A.加班标志,0) as 加班标志," & _
    "       A.开单人,A.划价人,A.操作员姓名,A.发生时间,A.病人病区ID " & _
    " From " & IIf(mblnNOMoved, zlGetFullFieldsTable("住院费用记录"), "住院费用记录 A") & " ,人员表 C " & _
    " Where NO=[1] And A.记录性质=2 And A.门诊标志=2 And Nvl(A.多病人单,0)=1 And Nvl(A.操作员姓名,A.划价人)=C.姓名" & _
    "       And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & vbNewLine & _
    "       And Rownum=1 And A.记录状态" & IIf(mblnDelete, "=2", " IN(0,1,3)") & _
            IIf(mstrTime <> "", " And A.登记时间=[2]", "") & _
            IIf(mbytInState = 0 And gbytBilling = 0, " And A.操作员姓名 is Not Null", "") & _
            IIf(mbytInState = 0 And gbytBilling = 1, " And A.操作员姓名 is Null And A.划价人 is Not NULL", "") & _
            IIf(mbytInState = 0 And gbytBilling = 2, " And A.操作员姓名 is Null And A.划价人 is Not NULL", "")
    
    If mstrTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, CDate(mstrTime))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    End If
    
    If rsTmp.EOF Then
        MsgBox "没有找到该单据！请检查该单据是否属于住院记帐表.", vbInformation, gstrSysName
        Exit Function
    Else
        If blnDelete Then
            If InStr(mstrPrivsOpt, ";全院销帐;") = 0 Then
                strUserUnitIDs = GetUserUnits(True)
                If InStr("," & strUserUnitIDs & ",", "," & rsTmp!开单部门ID & ",") = 0 Then
                    MsgBox "你没有权限对其它科室的单据销帐！", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        Else
            If mbytUseType = 0 Or mbytUseType = 1 Then
                If InStr(mstrPrivs, ";所有病区;") = 0 And mlngUnitID > 0 Then
                    If InStr(1, "," & mstrUnitIDs & ",", "," & IIf(IsNull(rsTmp!病人病区ID), 0, rsTmp!病人病区ID) & ",") = 0 Then
                        MsgBox "你没有权限读取其它病区的单据！", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    
    cboNO.Text = strNO
    
    txtDate.Text = Format(rsTmp!发生时间, "yyyy-MM-dd HH:mm:ss")
    chk加班.Value = IIf(IsNull(rsTmp!加班标志), 0, rsTmp!加班标志)
                
    Call Set开单人开单科室(cbo开单人, cbo开单科室, mrs开单人, mrs开单科室, Nvl(rsTmp!开单人), Nvl(rsTmp!开单部门ID, 0))
    
    '-----------------------------------------------------------------------------------
    '读取单据收费细目
    If blnDelete Then
         '退费单无需考虑后备表,前面的操作已禁止
        '读取准退数,并计算应收金额,实收金额(金额=剩余金额*(准退数/剩余数))
        
        '读取单据中原始记录的费用ID
        strSQL1 = _
            " Select A.ID,A.序号,A.收费细目ID," & _
            " Nvl(A.付数,1)*A.数次" & IIf(gbln住院单位, "/Nvl(B.住院包装,1)", "") & " as 原始数量" & _
            " From 住院费用记录 A,药品规格 B" & _
            " Where A.NO=[1] And A.记录状态 IN(0,1,3) And A.价格父号 is NULL" & _
            " And A.收费细目ID=B.药品ID(+) And A.记录性质=2 And A.门诊标志=2 And Nvl(A.多病人单,0)=1" & _
            IIf(mstrTime <> "", " And A.登记时间=[3]", "") & _
            IIf(mstr病人IDs <> "", " And Instr([2],','||A.病人ID||',')>0", "")
        
        '读取药品收发记录中的准退数
        strSQL2 = _
            " Select A.费用ID,Sum(Nvl(A.付数,1)*A.实际数量" & IIf(gbln住院单位, "/Nvl(B.住院包装,1)", "") & ") as 准退数量" & _
            " From 药品收发记录 A,药品规格 B" & _
            " Where A.NO=[1] And MOD(A.记录状态,3)=1" & _
            " And A.药品ID=B.药品ID(+) And A.单据 IN(10,26) And A.审核人 is NULL" & _
            " Group by A.费用ID"
        
        '整张单据汇总结果(明细到收费细目)
        '执行状态应该在原始记录上判断(部分退药且部分退费的记录)
        '当退两次以上时"记录状态,序号"重复,AVG有问题,所以要用"执行状态"
        strSQL = "Select Nvl(价格父号,序号) From 住院费用记录 " & _
            " Where 记录性质=2 And 门诊标志=2 And Nvl(多病人单,0)=1" & _
            " And 记录状态 IN(0,1,3) And NO=[1] And Nvl(执行状态,0)<>1" & _
            IIf(mstrTime <> "", " And 登记时间=[3]", "") & strFeeKind
            
        '如果已结帐单据禁止销帐,或是医保记帐的单据。则在原始单据行中只取未结帐部分
        Call GetBillInsures(strInsure, strNO)
        If strInsure <> "" Then
            arrInsure = Split(strInsure, ",")
            For i = 0 To UBound(arrInsure)
                If Not gclsInsure.GetCapability(support允许冲销已结帐的记帐单据, , arrInsure(i)) Then
                    blnDo = True: Exit For '只要有一个不允许，则不允许
                End If
            Next
        Else
            blnDo = gbytBillOpt = 2
        End If
        If blnDo Then
            strSQL = strSQL & " And Nvl(价格父号,序号) IN" & _
                " (" & _
                " Select Nvl(价格父号,序号) as 序号" & _
                " From 住院费用记录 " & _
                " Where NO=[1] And 记录性质 IN(2,12)" & _
                " Group by Nvl(价格父号,序号)" & _
                " Having Sum(Nvl(结帐金额,0))=0" & _
                " )"
        End If
                    
        '因为是将要汇总求有剩余数量的，所以不能用直接用时间限制，用序号限制
        strSQL = _
            " Select A.记录状态,A.执行状态,Nvl(A.价格父号,A.序号) as 序号," & _
            " A.病人ID,A.主页ID,A.婴儿费,A.姓名,A.性别,A.年龄,A.床号,A.费别," & _
            " A.收费细目ID,C.编码,C.名称 as 类别,B.名称,B.规格,Nvl(A.费用类型,B.费用类型) 费用类型," & _
            IIf(gbln住院单位, "Decode(X.药品ID,NULL,A.计算单位,X.住院单位)", "A.计算单位") & " as 计算单位," & _
            " Avg(Nvl(A.付数,1)) as 付数," & _
            " Avg(A.数次" & IIf(gbln住院单位, "/Nvl(X.住院包装,1)", "") & ") as 数次," & _
            " Sum(A.标准单价" & IIf(gbln住院单位, "*Nvl(X.住院包装,1)", "") & ") as 单价," & _
            " Sum(A.应收金额) as 应收金额,Sum(A.实收金额) as 实收金额, " & _
            " D.名称 as 执行部门,A.附加标志" & _
            " From 住院费用记录 A,收费项目目录 B,收费项目类别 C,部门表 D,药品规格 X" & _
            " Where A.收费细目ID=B.ID And C.编码=A.收费类别 And A.执行部门ID=D.ID(+)" & _
            " And A.收费细目ID=X.药品ID(+) And A.记录性质=2 And A.门诊标志=2 And Nvl(A.多病人单,0)=1" & _
            " And A.NO=[1] And Nvl(A.价格父号,A.序号) IN(" & strSQL & ")" & _
            IIf(mstr病人IDs <> "", " And Instr([2],','||A.病人ID||',')>0", "") & _
            " Group by A.记录状态,A.执行状态,Nvl(A.价格父号,A.序号),A.病人ID,A.主页ID,A.婴儿费,A.姓名,A.性别,A.年龄,A.床号,A.费别," & _
            " A.收费细目ID,C.编码,C.名称,B.名称,B.规格,Nvl(A.费用类型,B.费用类型),A.计算单位,D.名称,A.附加标志,X.药品ID,X.住院单位"
            
        '最后计算结果
        '当"准退数量=原始数量"时,付数才保留
        '排开已经全部退费的行(执行状态=0的一种可能)
        '有剩余数量无准退数量的有两种情况：
            '1.无对应的收发记录(如普通费用或不跟踪在用的卫材),这时应用剩余数量
            '2.收发记录中已全部发放,即已全部执行,SQL已排除这种记录
        strSQL = _
            " Select A.序号,A.病人ID,A.主页ID,A.婴儿费,A.姓名,A.性别,A.年龄,A.床号,A.费别," & _
            " A.收费细目ID,A.编码,A.类别,A.名称,A.规格,A.费用类型,A.计算单位," & _
            " Decode(Sign(Nvl(C.准退数量,Sum(A.付数*A.数次))-B.原始数量),0,Avg(A.付数),1) as 准退付数," & _
            " Decode(Sign(Nvl(C.准退数量,Sum(A.付数*A.数次))-B.原始数量),0,Sum(A.数次),Nvl(C.准退数量,Sum(A.付数*A.数次))) as 准退数次," & _
            " Nvl(C.准退数量,Sum(A.付数*A.数次)) as 准退数量,Sum(A.付数*A.数次) as 剩余数量," & _
            " A.单价,Sum(A.应收金额) as 剩余应收,Sum(A.实收金额) as 剩余实收,A.执行部门,A.附加标志" & _
            " From (" & strSQL & ") A,(" & strSQL1 & ") B,(" & strSQL2 & ") C" & _
            " Where A.序号=B.序号 And B.ID=C.费用ID(+)" & _
            " Group by A.序号,A.病人ID,A.主页ID,A.婴儿费,A.姓名,A.性别,A.年龄,A.床号,A.费别,A.收费细目ID,A.编码,A.类别," & _
            " A.名称,A.规格,A.费用类型,A.计算单位,A.单价,B.原始数量,C.准退数量,A.执行部门,A.附加标志" & _
            " Having Sum(A.付数*A.数次)<>0"
        If strInsure <> "" Then
            '医保病人费用可能部份销帐,但必须整笔销帐(准退数量=原始数量)
            arrInsure = Split(strInsure, ",")
            For i = 0 To UBound(arrInsure)
                If Not gclsInsure.GetCapability(support允许部分冲销明细, , arrInsure(i)) Then
                    strSQL = strSQL & " And (Nvl(C.准退数量,Sum(A.付数*A.数次))=B.原始数量" & _
                        " Or A.病人ID+0 IN(Select 病人ID From 病人信息 Where 险类 is NULL And 病人ID=A.病人ID))"
                    Exit For  '只要有一个不允许,则整个都不允许
                End If
            Next
        End If
            
        strSQL = _
        " Select A.序号,A.病人ID,A.主页ID,A.婴儿费,A.姓名,A.性别,A.年龄,A.床号,A.费别,A.编码,A.类别,Nvl(B.名称,A.名称) as 名称,E1.名称 as 商品名," & _
        "       A.规格,A.费用类型,A.计算单位,A.准退付数 as 付数,A.准退数次 as 数次,A.单价," & _
        "       A.剩余应收*(A.准退数量/A.剩余数量) as 应收金额," & _
        "       A.剩余实收*(A.准退数量/A.剩余数量) as 实收金额," & _
        "       A.执行部门,A.附加标志" & _
        " From (" & strSQL & ") A,收费项目别名 B,收费项目别名 E1" & _
        " Where A.收费细目ID=B.收费细目ID(+) And B.码类(+)=1 And B.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
        "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3" & _
        " Order by LPAD(A.床号,10,' '),A.病人ID,A.序号"
    ElseIf mbytInState = 0 And gbytBilling = 2 Then
        '读取记帐划价单(记帐审核),只读取未审核部份
        '不用考虑在后备表中
        strSQL = _
            " Select Nvl(A.价格父号,A.序号) as 序号,A.病人ID,A.主页ID,A.婴儿费,A.姓名,A.性别,A.年龄," & _
            " A.床号,A.费别,A.收费细目ID,C.编码,C.名称 as 类别,B.名称,B.规格,Nvl(A.费用类型,B.费用类型) 费用类型," & _
            IIf(gbln住院单位, "Decode(X.药品ID,NULL,A.计算单位,X.住院单位)", "A.计算单位") & " as 计算单位," & _
            " Avg(Nvl(A.付数,1)) as 付数," & _
            " Avg(A.数次" & IIf(gbln住院单位, "/Nvl(X.住院包装,1)", "") & ") as 数次," & _
            " Sum(A.标准单价" & IIf(gbln住院单位, "*Nvl(X.住院包装,1)", "") & ") as 单价," & _
            " Sum(A.应收金额) as 应收金额,Sum(A.实收金额) as 实收金额, " & _
            " D.名称 as 执行部门,A.附加标志" & _
            " From 住院费用记录 A,收费项目目录 B,收费项目类别 C,部门表 D,药品规格 X" & _
            " Where A.收费细目ID=B.ID And C.编码=A.收费类别 And A.执行部门ID=D.ID(+)" & _
            " And A.收费细目ID=X.药品ID(+) And A.记录性质=2 And Nvl(A.多病人单,0)=1" & _
            " And A.记录状态=0 And 门诊标志=2 And A.NO=[1]" & _
            " Group by Nvl(A.价格父号,A.序号),A.记录状态,A.病人ID,A.主页ID,A.婴儿费,A.姓名,A.性别,A.年龄,A.床号,A.费别," & _
            " A.收费细目ID,C.编码,C.名称,B.名称,B.规格,Nvl(A.费用类型,B.费用类型),A.计算单位,D.名称,A.附加标志,X.药品ID,X.住院单位"
            
        strSQL = "" & _
        " Select A.序号,A.病人ID,A.主页ID,A.婴儿费,A.姓名,A.性别,A.年龄,A.床号,A.费别," & _
        "       A.编码,A.类别,Nvl(B.名称,A.名称) as 名称,E1.名称 as 商品名,A.规格,A.费用类型,A.计算单位," & _
        "       A.付数,A.数次,A.单价,A.应收金额,A.实收金额,A.执行部门,A.附加标志" & _
        " From (" & strSQL & ") A,收费项目别名 B,收费项目别名 E1" & _
        " Where A.收费细目ID=B.收费细目ID(+) And B.码类(+)=1 And B.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
        "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3" & _
        " Order by LPAD(A.床号,10,' '),A.病人ID,A.序号"
    Else
        '读取单据原始内容
        intSign = IIf(mblnDelete, -1, 1) '数量,金额正负符号
        strSQL = _
            " Select Nvl(A.价格父号,A.序号) as 序号,A.病人ID,A.主页ID,A.婴儿费,A.姓名,A.性别,A.年龄,A.床号,A.费别," & _
            " A.收费细目ID,C.编码,C.名称 as 类别,B.名称,B.规格,Nvl(A.费用类型,B.费用类型) 费用类型," & _
            IIf(gbln住院单位, "Decode(X.药品ID,NULL,A.计算单位,X.住院单位)", "A.计算单位") & " as 计算单位," & _
            " Avg(Nvl(A.付数,1)) as 付数," & _
            " Avg(" & intSign & "*A.数次" & IIf(gbln住院单位, "/Nvl(X.住院包装,1)", "") & ") as 数次," & _
            " Sum(A.标准单价" & IIf(gbln住院单位, "*Nvl(X.住院包装,1)", "") & ") as 单价," & _
            " Sum(" & intSign & "*A.应收金额) as 应收金额,Sum(" & intSign & "*A.实收金额) as 实收金额, " & _
            " D.名称 as 执行部门,A.附加标志" & _
            " From " & IIf(mblnNOMoved, zlGetFullFieldsTable("住院费用记录"), "住院费用记录 A") & ",收费项目目录 B,收费项目类别 C,部门表 D,药品规格 X" & _
            " Where A.收费细目ID=B.ID And C.编码=A.收费类别 And A.执行部门ID=D.ID(+)" & _
            " And A.收费细目ID=X.药品ID(+) And A.记录性质=2 And Nvl(A.多病人单,0)=1 And 门诊标志=2" & _
            " And A.记录状态" & IIf(mblnDelete, "=2", " IN(0,1,3)") & " And A.NO=[1]" & _
            IIf(mstrTime <> "", " And A.登记时间=[3]", "") & _
            " Group by Nvl(A.价格父号,A.序号),A.病人ID,A.主页ID,A.婴儿费,A.姓名,A.性别,A.年龄,A.床号,A.费别," & _
            " A.收费细目ID,C.编码,C.名称,B.名称,B.规格,Nvl(A.费用类型,B.费用类型),A.计算单位,D.名称,A.附加标志,X.药品ID,X.住院单位"
            
        strSQL = "" & _
        " Select A.序号,A.病人ID,A.主页ID,A.婴儿费,A.姓名,A.性别,A.年龄,A.床号,A.费别," & _
        "       A.编码,A.类别,Nvl(B.名称,A.名称) as 名称,E1.名称 as 商品名,A.规格,A.费用类型,A.计算单位," & _
        "       A.付数,A.数次,A.单价,A.应收金额,A.实收金额,A.执行部门,A.附加标志" & _
        " From (" & strSQL & ") A,收费项目别名 B,收费项目别名 E1" & _
        " Where A.收费细目ID=B.收费细目ID(+) And B.码类(+)=1 And B.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
        "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3" & _
        " Order by LPAD(A.床号,10,' '),A.病人ID,A.序号"
        
    End If
    If mstrTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, "," & mstr病人IDs & ",", CDate(mstrTime))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, "," & mstr病人IDs & ",")
    End If
    
    If rsTmp.EOF Then Exit Function
    
    Bill.Redraw = False
    Bill.Rows = rsTmp.RecordCount + 1
    ReDim marrSerial(1 To rsTmp.RecordCount)
    Set mcolPatiInfo = New Collection
    
    For i = 1 To rsTmp.RecordCount
        If gbytBilling = 2 And Not mblnPrint Then mblnPrint = True

        marrSerial(i) = rsTmp!序号 '用于记帐销帐或划价审核
        mcolPatiInfo.Add rsTmp!病人ID & "," & Val("" & rsTmp!主页ID) & "," & Val("" & rsTmp!婴儿费), "R" & i
        
        Bill.TextMatrix(i, BillCol.姓名) = IIf(IsNull(rsTmp!姓名), "", rsTmp!姓名)
        Bill.TextMatrix(i, BillCol.性别) = IIf(IsNull(rsTmp!性别), "", rsTmp!性别)
        Bill.TextMatrix(i, BillCol.年龄) = IIf(IsNull(rsTmp!年龄), "", rsTmp!年龄)
        Bill.TextMatrix(i, BillCol.床号) = IIf(IsNull(rsTmp!床号), "", rsTmp!床号)
        Bill.TextMatrix(i, BillCol.费别) = IIf(IsNull(rsTmp!费别), "", rsTmp!费别)
        Bill.TextMatrix(i, BillCol.类别) = rsTmp!类别
        Bill.TextMatrix(i, BillCol.项目) = rsTmp!名称
        '刘兴洪:27990 2010-02-22 17:01:30
        Bill.TextMatrix(i, BillCol.商品名) = Nvl(rsTmp!商品名)
        Bill.TextMatrix(i, BillCol.规格) = IIf(IsNull(rsTmp!规格), "", rsTmp!规格)
        Bill.TextMatrix(i, BillCol.单位) = IIf(IsNull(rsTmp!计算单位), "", rsTmp!计算单位)
        Bill.TextMatrix(i, BillCol.付数) = IIf(IsNull(rsTmp!付数), "", rsTmp!付数)
        Bill.TextMatrix(i, BillCol.数次) = FormatEx(rsTmp!数次, 5)
        Bill.TextMatrix(i, BillCol.单价) = Format(rsTmp!单价, gstrFeePrecisionFmt)
        Bill.TextMatrix(i, BillCol.应收金额) = Format(rsTmp!应收金额, gstrDec)
        Bill.TextMatrix(i, BillCol.实收金额) = Format(rsTmp!实收金额, gstrDec)
        Bill.TextMatrix(i, BillCol.执行科室) = Nvl(rsTmp!执行部门)
        Bill.TextMatrix(i, BillCol.标志) = IIf(rsTmp!附加标志 = 1, "√", "")
        Bill.TextMatrix(i, BillCol.类型) = IIf(IsNull(rsTmp!费用类型), "", rsTmp!费用类型)
        
        '设置销帐标志
        If Bill.TextMatrix(0, Bill.Cols - 1) = "销帐" Then
            If mlngDelRow = 0 Or mlngDelRow <> 0 And mlngDelRow = rsTmp!序号 Then
                Bill.TextMatrix(i, Bill.Cols - 1) = "√"
            End If
        End If
        
        curTotal = curTotal + rsTmp!实收金额
        rsTmp.MoveNext
    Next
    '针对列编辑性质设置颜色
    Call InitBillColumnColor
    
    Bill.Redraw = True
    
    ReadBill = True
    txt单据.Text = Format(curTotal, gstrDec)
    Call Bill_EnterCell(Bill.Row, Bill.Col)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetShowCol()
'功能：付数列的控制(浏览时展开)
    mrsClass.Filter = "编码='7'"
    If mrsClass.RecordCount = 0 Then
        Bill.ColWidth(BillCol.付数) = 0
    ElseIf Bill.ColWidth(BillCol.付数) = 0 Then
        Bill.ColWidth(BillCol.付数) = 300
    End If
End Sub
Private Sub InitBillColumnColor()
    
    Bill.SetColColor BillCol.类别, &HE7CFBA
    Bill.SetColColor BillCol.项目, &HE7CFBA
    Bill.SetColColor BillCol.数次, &HE7CFBA
    Bill.SetColColor BillCol.执行科室, &HE7CFBA
    Bill.SetColColor BillCol.付数, &HE0E0E0
    Bill.SetColColor BillCol.单价, &HE0E0E0
    Bill.SetColColor BillCol.标志, &HE0E0E0
End Sub
Private Sub ClearRows()
    Dim i As Long
    For i = 1 To Bill.Rows - 1
        Bill.RowData(i) = 0
    Next
End Sub

Private Sub Bill_cboClick(ListIndex As Long)
    Dim dblStock As Double
    Dim int险类 As Integer
    Dim lng执行科室 As Long, str执行科室 As String
    '药品库存检查
    If ListIndex <> -1 And Bill.TextMatrix(0, Bill.Col) = "执行科室" Then
        If mobjBill.Details.Count >= Bill.Row Then
            With mobjBill.Details(Bill.Row)
                If .执行部门ID <> Bill.ItemData(Bill.ListIndex) Then
                    lng执行科室 = .执行部门ID: str执行科室 = Bill.TextMatrix(Bill.Row, Bill.Col)
                    .执行部门ID = Bill.ItemData(Bill.ListIndex)
                    Bill.TextMatrix(Bill.Row, Bill.Col) = Bill.CboText
                
                    If InStr(",5,6,7,", .收费类别) > 0 And Not gbln分离发药 Then
                        '取库存
                        dblStock = GetStock(.收费细目ID, .执行部门ID)
                        If gbln住院单位 Then
                            dblStock = dblStock / .Detail.住院包装
                        End If
                        .Detail.库存 = dblStock  '记录当前行药品库存
                        Call ShowStock(.Detail.名称, .Detail.库存)
                        
                        '药房改变,实价药品重新计算价格
                        'If .Detail.变价 Then    '如果费别的计算方式是成本价加收法,则需要重算价格,这里简化不作判断
                            Call CalcMoneys(Bill.Row)
                            Call ShowDetails(Bill.Row)
                            Call CalcOneTotal(Bill.Row)
                        'End If
                    ElseIf .收费类别 = "4" And .Detail.跟踪在用 Then
                        '取库存
                        dblStock = GetStock(.收费细目ID, .执行部门ID)
                        .Detail.库存 = dblStock
                        Call ShowStock(.Detail.名称, .Detail.库存)
                        
                        '发料部门改变,时价卫材重新计算价格
                        If .Detail.变价 Then
                            Call CalcMoneys(Bill.Row)
                            Call ShowDetails(Bill.Row)
                            Call CalcOneTotal(Bill.Row)
                        End If
                    ElseIf InStr(",4,5,6,7,", .收费类别) = 0 Then
                        If CheckItemHaveSub(Bill.Row) Then Call SetSubItemDept(Bill.Row) '如果存在从项,则改变非药品行的执行科室
                    End If
                    int险类 = Val(mobjBill.Details(Bill.Row).发药窗口)
                    If int险类 <> 0 And mobjBill.Details(Bill.Row).数次 <> 0 Then
                        If gclsInsure.GetCapability(support实时监控, mobjBill.Details(Bill.Row).病人ID, int险类) Then
                            If gclsInsure.CheckItem(int险类, 1, 0, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 2, gbytBilling, Bill.Row)) = False Then
                                Bill.Text = "": Bill.TxtVisible = False
                                Bill.cboObj.Text = str执行科室: .执行部门ID = lng执行科室
                                Exit Sub
                            End If
                        End If
                    End If
                    
                    If mobjBill.Details(Bill.Row).数次 <> 0 Then
                        If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModule, 1, 0, _
                            MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 2, gbytBilling, Bill.Row)) = False Then
                            Bill.Text = "": Bill.TxtVisible = False
                            Bill.cboObj.Text = str执行科室: .执行部门ID = lng执行科室
                            Exit Sub
                        End If
                    End If
                End If
            End With
        End If
    End If
End Sub

Public Function GetBillIndex(strFind As String) As Long
'功能：由字任串查找ComboBox的索引值
'参数：cbo=ComboBox,strFind=查找字符串
    Dim i As Long
    If strFind = "" Then GetBillIndex = -1: Exit Function
    For i = 0 To Bill.ListCount - 1
        If InStr(Bill.List(i), strFind) > 0 Then
            GetBillIndex = i
            Exit Function
        End If
    Next
    GetBillIndex = -1
End Function

Private Function CalcOneTotal(lngRow As Long, Optional blnShow As Boolean = True) As Currency
'功能：计算记帐单中当前行病人的在当前单子中的费用合计
'说明：以姓名为准
    Dim i As Long, strName As String, curTotal As Currency
    Dim tmpBillInCome As New BillInCome
    
    If Bill.TextMatrix(lngRow, BillCol.姓名) = "" Then
        If blnShow Then txt病人.Text = gstrDec
    Else
        If mobjBill.Details.Count = 0 Then
            strName = Bill.TextMatrix(lngRow, BillCol.姓名) '病人姓名
            If blnShow Then lbl病人.Caption = strName
            For i = 1 To Bill.Rows - 1
                If IsNumeric(Bill.TextMatrix(i, BillCol.实收金额)) Then
                    If Bill.TextMatrix(i, BillCol.姓名) = strName Then
                        curTotal = curTotal + CCur(Bill.TextMatrix(i, BillCol.实收金额))
                    End If
                End If
            Next
        Else
            If mobjBill.Details.Count >= lngRow Then
                strName = mobjBill.Details(lngRow).病人ID & mobjBill.Details(lngRow).姓名   '病人姓名
                If blnShow Then lbl病人.Caption = mobjBill.Details(lngRow).姓名
            ElseIf mrsInfo.State = 1 Then
                strName = mrsInfo!病人ID & mrsInfo!姓名
                If blnShow Then lbl病人.Caption = mrsInfo!姓名
            ElseIf Bill.TextMatrix(lngRow, BillCol.姓名) <> "" And mobjBill.Details.Count < lngRow And mobjBill.Details.Count >= lngRow - 1 And lngRow > 1 Then
                strName = mobjBill.Details(lngRow - 1).病人ID & mobjBill.Details(lngRow - 1).姓名
                If blnShow Then lbl病人.Caption = mobjBill.Details(lngRow - 1).姓名
            End If
            For i = 1 To Bill.Rows - 1
                If mobjBill.Details.Count >= i Then
                    If mobjBill.Details(i).病人ID & mobjBill.Details(i).姓名 = strName Then
                        For Each tmpBillInCome In mobjBill.Details(i).InComes
                            curTotal = curTotal + tmpBillInCome.实收金额
                        Next
                    End If
                End If
            Next
        End If
        If blnShow Then txt病人.Text = Format(curTotal, gstrDec)
    End If
    CalcOneTotal = curTotal
End Function

Private Function GetDetailNum(lngRow As Long) As Double
'功能：获取病人指定细目的总记帐数据(含本单据中)
'参数：lngRow=当前单据行
    Dim rsTmp As ADODB.Recordset
    Dim lngNum As Long, i As Long
    Dim strSQL As String
    Dim lng病人ID As Long, lng主页ID As Long
        
    If lngRow <= mobjBill.Details.Count Then
        lng病人ID = mobjBill.Details(lngRow).病人ID
        lng主页ID = mobjBill.Details(lngRow).主页ID
        
        '当前单据中的数量
        For i = 1 To mobjBill.Details.Count
            If i <> lngRow And mobjBill.Details(i).收费细目ID = mobjBill.Details(lngRow).收费细目ID And mobjBill.Details(i).病人ID = lng病人ID Then
                lngNum = lngNum + mobjBill.Details(i).数次 * IIf(mobjBill.Details(i).付数 = 0, 1, mobjBill.Details(i).付数)
            End If
        Next
        '数据库中的数量
        strSQL = _
            "Select Sum(A.数次*Nvl(A.付数,1)" & IIf(gbln住院单位, "/Nvl(B.住院包装,1)", "") & ") as NUM" & _
            " From 住院费用记录 A,药品规格 B" & _
            " Where A.价格父号 is Null And A.记帐费用=1" & _
            IIf(gbytBilling = 0, " And A.记录状态<>0", "") & _
            " And A.病人ID=[1] And Nvl(A.主页ID,0)=[2]" & _
            " And A.收费细目ID=B.药品ID(+) And A.收费细目ID+0=[3]"
        
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID, mobjBill.Details(lngRow).收费细目ID)
        If Not rsTmp.EOF Then
            lngNum = lngNum + Nvl(rsTmp!Num, 0)
        End If
        GetDetailNum = lngNum
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetWorkUnit(ByVal lng药品ID As Long, ByVal str类别 As String) As Boolean
'功能：取所有可供选择的药房
    Dim strSQL As String, str药房 As String, bytDay As Byte
    Dim int服务对象 As Integer, str服务对象 As String
    Dim int病人来源 As Integer, lng开单科室ID As Long
    
    '根据项目及权限确定药房的服务对象
    int服务对象 = Get服务对象(lng药品ID)
    'int服务对象 = mobjDetail.服务对象  '修改,导入时没有赋值
    If int服务对象 = 1 Then
        str服务对象 = "1,3"
    ElseIf int服务对象 = 2 Then
        str服务对象 = "2,3"
    ElseIf int服务对象 = 3 Then
        If InStr(mstrPrivsOpt, ";门诊留观记帐;") > 0 And gbln门诊留观 Then
            str服务对象 = "1,2,3"
        Else
            str服务对象 = "2,3"
        End If
    Else
        str服务对象 = "2,3"
    End If
        
    '确定病人来源
    int病人来源 = Get病人来源(Bill.Row)
    
    '确定病人科室
    If mrsInfo.State = 1 Then
        lng开单科室ID = Nvl(mrsInfo!科室ID, 0)
    ElseIf Bill.TextMatrix(Bill.Row, BillCol.姓名) <> "" And mobjBill.Details.Count < Bill.Row And Bill.Row > 1 Then
        lng开单科室ID = mobjBill.Details(Bill.Row - 1).科室ID
    Else
        lng开单科室ID = mobjBill.Details(Bill.Row).科室ID
    End If
    
    '开单科室
    If lng开单科室ID = 0 And cbo开单科室.ListIndex <> -1 Then lng开单科室ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
    
    If str类别 = "4" Then
        strSQL = _
        "Select Distinct c.Id, c.编码, c.简码, c.名称, b.工作性质, b.服务对象" & vbNewLine & _
        "From 收费执行科室 A, 部门性质说明 B, 部门表 C" & vbNewLine & _
        "Where a.执行科室id + 0 = b.部门id And b.工作性质 = '发料部门' And b.服务对象 IN(" & str服务对象 & ") And b.部门id = c.Id And" & vbNewLine & _
        "      (c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.撤档时间 Is Null) And (c.站点 = '" & gstrNodeNo & "' Or c.站点 Is Null) And" & vbNewLine & _
        "      (a.病人来源 Is Null Or a.病人来源 = [1]) And" & vbNewLine & _
        "      (a.开单科室id Is Null Or a.开单科室id = [2] Or Exists (Select 1 From 病区科室对应 Where 科室id = [2] And a.开单科室id = 病区id)) And a.收费细目id = [3]" & vbNewLine & _
        "Order By b.服务对象, c.编码"
    Else
        '由药品材质确定药房性质
        Select Case str类别
            Case "5"
                str药房 = "西药房"
            Case "6"
                str药房 = "成药房"
            Case "7"
                str药房 = "中药房"
        End Select
    
        '药品从系统指定的储备药房中找
        If Not gbln药房上班安排 Then
            strSQL = _
            " Select Distinct C.ID,C.编码,C.简码,C.名称,B.工作性质,B.服务对象 " & _
            " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
            " Where A.执行科室ID+0=B.部门ID And B.工作性质=[4]" & _
            "       And B.服务对象 IN(" & str服务对象 & ") And B.部门ID=C.ID" & _
            "       And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
            "       And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null) " & vbNewLine & _
            "       And (A.病人来源 is NULL Or A.病人来源=[1])" & _
            "       And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
            "       And A.收费细目ID=[3]" & _
            " Order by B.服务对象,C.编码"
        Else
            bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=周日,1=周一
            strSQL = _
            " Select Distinct C.ID,C.编码,C.简码,C.名称,B.工作性质,B.服务对象 " & _
            " From 收费执行科室 A,部门性质说明 B,部门表 C,部门安排 D" & _
            " Where A.执行科室ID+0=B.部门ID And B.工作性质=[4]" & _
            "       And B.服务对象 IN(" & str服务对象 & ") And B.部门ID=C.ID" & _
            "       And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
            "       And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null) " & vbNewLine & _
            "       And D.部门ID=C.ID And D.星期=[5]" & _
            "       And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.开始时间,'HH24:MI:SS') and To_Char(D.终止时间,'HH24:MI:SS') " & _
            "       And (A.病人来源 is NULL Or A.病人来源=[1])" & _
            "       And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
            "       And A.收费细目ID=[3]" & _
            " Order by B.服务对象,C.编码"
        End If
    End If
    
    On Error GoTo errH
    Set mrsWork = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, int病人来源, lng开单科室ID, lng药品ID, str药房, bytDay)
    GetWorkUnit = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FillPatient(lng科室ID As Long)
    Dim i As Long, j As Long, strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim objItem As ListItem
    Dim str在院 As String
    On Error GoTo errH
    
    str在院 = "    Exists(Select 1 From 在院病人 ZY Where ZY.病人ID=B.病人ID)"
    '是否具有强制记帐权限
    If InStr(mstrPrivsOpt, ";出院未结强制记帐;") > 0 And InStr(mstrPrivsOpt, ";出院结清强制记帐;") > 0 Then
        If gintOutDay = 0 Then
            strSQL = "   And " & str在院
        Else
            strSQL = " And (" & str在院 & " Or B.出院日期>Trunc(Sysdate)-" & gintOutDay & ")"
        End If
    ElseIf InStr(mstrPrivsOpt, ";出院未结强制记帐;") > 0 Then
        If gintOutDay = 0 Then
            strSQL = " And (" & str在院 & " And B.状态<>3 Or X.费用余额<>0 And " & str在院 & "  And B.状态=3)"
        Else
            strSQL = " And (" & str在院 & " And B.状态<>3 Or X.费用余额<>0 And (" & str在院 & " And B.状态=3 Or B.出院日期>Trunc(Sysdate)-" & gintOutDay & "))"
        End If
    ElseIf InStr(mstrPrivsOpt, ";出院结清强制记帐;") > 0 Then
        If gintOutDay = 0 Then
            strSQL = " And (" & str在院 & "  And B.状态<>3 Or X.费用余额=0 And " & str在院 & "  And B.状态=3)"
        Else
            strSQL = " And (" & str在院 & "  And B.状态<>3 Or X.费用余额=0 And (" & str在院 & "  And B.状态=3 Or B.出院日期>Trunc(Sysdate)-" & gintOutDay & "))"
        End If
    Else
        '没有权限对出院和预出院病人结帐
        strSQL = " And " & str在院 & "  And B.状态<>3"
    End If
    
    '留观病人记帐权限
    If (InStr(mstrPrivsOpt, ";门诊留观记帐;") > 0 And gbln门诊留观) And (InStr(mstrPrivsOpt, ";住院留观记帐;") > 0 And gbln住院留观) Then
        strSQL = strSQL & " And Nvl(B.病人性质,0) IN(0,1,2)"
    ElseIf InStr(mstrPrivsOpt, ";门诊留观记帐;") > 0 And gbln门诊留观 Then
        strSQL = strSQL & " And Nvl(B.病人性质,0) IN(0,1)"
    ElseIf InStr(mstrPrivsOpt, ";住院留观记帐;") > 0 And gbln住院留观 Then
        strSQL = strSQL & " And Nvl(B.病人性质,0) IN(0,2)"
    Else
        strSQL = strSQL & " And Nvl(B.病人性质,0)=0"
    End If
    
    lvwPati.ListItems.Clear
    
    strSQL = "Select A.病人ID,A.住院号,nvl(B.姓名,A.姓名) as 姓名,nvl(b.性别,A.性别) as 性别,A.年龄," & _
            " B.出院病床 as 床位,B.出院日期,B.险类,B.病人性质,B.病人类型" & _
            " From 病人信息 A,病案主页 B,病人余额 X" & _
            " Where A.停用时间 is NULL And A.病人ID=B.病人ID" & _
            " And Nvl(B.主页ID,0)<>0 And A.主页ID=B.主页ID" & strSQL & _
            " And A.病人ID=X.病人ID(+)  And X.性质(+)=1 And X.类型(+)=2 And B.出院科室ID = [1]" & _
            " Order by A.住院号 Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng科室ID)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            If IIf(IsNull(rsTmp!病人性质), 0, rsTmp!病人性质) = 0 Then
                Set objItem = lvwPati.ListItems.Add(, "_" & rsTmp!病人ID, rsTmp!病人ID, , 1)
            Else
                Set objItem = lvwPati.ListItems.Add(, "_" & rsTmp!病人ID, rsTmp!病人ID, , 2)
            End If
            objItem.SubItems(1) = IIf(IsNull(rsTmp!住院号), "", rsTmp!住院号)
            objItem.SubItems(2) = rsTmp!姓名
            objItem.SubItems(3) = IIf(IsNull(rsTmp!床位), "", rsTmp!床位)
            objItem.SubItems(4) = IIf(IsNull(rsTmp!性别), "", rsTmp!性别)
            objItem.SubItems(5) = IIf(IsNull(rsTmp!年龄), "", rsTmp!年龄)
            objItem.SubItems(6) = IIf(IsNull(rsTmp!出院日期), "√", "")
                        
            objItem.ForeColor = zlDatabase.GetPatiColor(Nvl(rsTmp!病人类型))
            For j = 1 To objItem.ListSubItems.Count
                objItem.ListSubItems(j).ForeColor = zlDatabase.GetPatiColor(Nvl(rsTmp!病人类型))
            Next
            
            rsTmp.MoveNext
        Next
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetMoneyStr(lng病人ID As Long) As String
'功能：根据姓人ID获取病人费用信息
    Dim i As Long
    
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).病人ID = lng病人ID Then
            GetMoneyStr = mobjBill.Details(i).就诊卡号
            Exit For
        End If
    Next
End Function

Private Sub ShowDeleteCol(blnShow As Boolean)
'功能：显示\隐藏销帐标志列
    Dim i As Long, blnACT As Boolean
    If blnShow Then
        If Bill.TextMatrix(0, Bill.Cols - 1) <> "销帐" Then
            Bill.Redraw = False
            Bill.Cols = Bill.Cols + 1
            Bill.TextMatrix(0, Bill.Cols - 1) = "销帐"
            Bill.ColAlignment(Bill.Cols - 1) = 4
            Bill.ColWidth(Bill.Cols - 1) = 450
            Bill.ColData(Bill.Cols - 1) = BillColType.CheckBox
            
            blnACT = Bill.Active: Bill.Active = False
            Bill.Row = 0: Bill.Col = Bill.Cols - 1: Bill.MsfObj.CellForeColor = vbRed
            Bill.Row = 1: Bill.Col = Bill.Cols - 1
            Bill.Active = blnACT
            
            Bill.ColWidth(BillCol.类别) = GetOrigColWidth(BillCol.类别) - 100
            Bill.ColWidth(BillCol.项目) = GetOrigColWidth(BillCol.项目) - 200
            Bill.ColWidth(BillCol.执行科室) = GetOrigColWidth(BillCol.执行科室) - 150
            Bill.Redraw = True
        End If
    Else
        If Bill.TextMatrix(0, Bill.Cols - 1) = "销帐" Then
            Bill.Redraw = False
            Bill.Cols = Bill.Cols - 1
            Bill.ColWidth(BillCol.类别) = GetOrigColWidth(BillCol.类别)
            Bill.ColWidth(BillCol.项目) = GetOrigColWidth(BillCol.项目)
            Bill.ColWidth(BillCol.执行科室) = GetOrigColWidth(BillCol.执行科室)
            Bill.Redraw = True
        End If
    End If
    
    cmdSelALL.Visible = blnShow
    cmdClear.Visible = blnShow
    
    If blnShow Then
        chkIn.Visible = False
        txtIn.Visible = False
        txt单据.Left = Val(txt单据.Tag) - chkIn.Width - txtIn.Width
        lbl单据.Left = txt单据.Left - lbl单据.Width - 45
        txt病人.Left = Val(txt病人.Tag) - chkIn.Width - txtIn.Width
        lbl病人.Left = txt病人.Left - lbl病人.Width - 45
    Else
        If mbytInState = 0 And (gbytBilling = 0 Or gbytBilling = 1) Then
            chkIn.Visible = True
            txtIn.Visible = True
            txt单据.Left = Val(txt单据.Tag)
            lbl单据.Left = txt单据.Left - lbl单据.Width - 45
            txt病人.Left = Val(txt病人.Tag)
            lbl病人.Left = txt病人.Left - lbl病人.Width - 45
        Else
            chkIn.Visible = False
            txtIn.Visible = False
            txt单据.Left = Val(txt单据.Tag) - chkIn.Width - txtIn.Width
            lbl单据.Left = txt单据.Left - lbl单据.Width - 45
            txt病人.Left = Val(txt病人.Tag) - chkIn.Width - txtIn.Width
            lbl病人.Left = txt病人.Left - lbl病人.Width - 45
        End If
    End If
End Sub

Private Function GetOrigColWidth(ByVal intIdx As Integer) As Long
'功能：获取指定列的原始列宽
    GetOrigColWidth = Val(Split(Split(STR_HEAD, ";")(intIdx), ",")(1))
End Function

Private Function SaveModi() As Boolean
'功能：保存当前修改的费用单据
    Dim strSQL As String
    
    strSQL = "zl_病人费用记录_Update('" & cboNO.Text & "',2,'" & zlStr.NeedName(cbo开单人.Text) & "'," & _
        "To_Date('" & txtDate.Text & "','YYYY-MM-DD HH24:MI:SS'),NULL,2)"
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    SaveModi = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetColNum(strHead As String) As Integer
    Dim i As Long
    For i = 0 To Bill.Cols - 1
        If Bill.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
    Next
End Function

Private Function GetInputDetail(ByVal lng项目id As Long, ByVal int险类 As Integer) As Detail
'功能：读取收费项目信息
    Dim objDetail As New Detail
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
       
    If int险类 > 0 Then
        strSQL = _
        " Select A.ID,A.类别,B.名称 as 类别名称,A.编码,Nvl(E.名称,A.名称) as 名称,E1.名称 as 商品名,A.规格,A.计算单位," & _
        "       A.屏蔽费别,A.是否变价,A.加班加价,A.执行科室,A.费用类型,A.补充摘要,A.服务对象,M.要求审批," & _
        "       Decode(A.类别,'4',D.诊疗ID,C.药名ID) as 药名ID," & _
        "       Decode(A.类别,'4',D.在用分批,C.药房分批) as 分批," & _
        "       Decode(A.类别,'4',1,C.住院包装) as 住院包装," & _
        "       Decode(A.类别,'4',A.计算单位,C.住院单位) as 住院单位,D.跟踪在用,A.录入限量,C.中药形态" & _
        " From 收费项目目录 A,收费项目类别 B,药品规格 C,材料特性 D,收费项目别名 E,收费项目别名 E1,保险支付项目 M" & _
        " Where A.类别=B.编码 And A.ID=C.药品ID(+) And A.ID=D.材料ID(+)" & _
        "       And A.ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
        "       And A.ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3" & _
        "       And A.ID=M.收费细目ID(+) And M.险类(+)=[2]" & vbNewLine & _
        "       And A.ID=[1]"
    Else
        strSQL = _
        " Select A.ID,A.类别,B.名称 as 类别名称,A.编码,Nvl(E.名称,A.名称) as 名称,E1.名称 as 商品名,A.规格,A.计算单位," & _
        "       A.屏蔽费别,A.是否变价,A.加班加价,A.执行科室,A.费用类型,A.补充摘要,A.服务对象,0 as 要求审批," & _
        "       Decode(A.类别,'4',D.诊疗ID,C.药名ID) as 药名ID," & _
        "       Decode(A.类别,'4',D.在用分批,C.药房分批) as 分批," & _
        "       Decode(A.类别,'4',1,C.住院包装) as 住院包装," & _
        "       Decode(A.类别,'4',A.计算单位,C.住院单位) as 住院单位,D.跟踪在用,A.录入限量,C.中药形态" & _
        " From 收费项目目录 A,收费项目类别 B,药品规格 C,材料特性 D,收费项目别名 E,收费项目别名 E1" & _
        " Where A.类别=B.编码 And A.ID=C.药品ID(+) And A.ID=D.材料ID(+)" & _
        "       And A.ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
        "       And A.ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3" & _
        "       And A.ID=[1]"
    End If
        
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng项目id, int险类)
    With objDetail
        .ID = rsTmp!ID
        .药名ID = Nvl(rsTmp!药名ID, 0) '用于判断输入重复
        .类别 = rsTmp!类别
        .类别名称 = rsTmp!类别名称
        .编码 = rsTmp!编码
        .名称 = rsTmp!名称
        .规格 = Nvl(rsTmp!规格)
        .计算单位 = Nvl(rsTmp!计算单位)
        .住院单位 = Nvl(rsTmp!住院单位)
        .住院包装 = Nvl(rsTmp!住院包装, 1)
        .分批 = Nvl(rsTmp!分批, 0) = 1 '是否药房分批
        .变价 = Nvl(rsTmp!是否变价, 0) = 1 '对药品表明是否时价
        .类型 = Nvl(rsTmp!费用类型)
        .加班加价 = Nvl(rsTmp!加班加价, 0) = 1
        .屏蔽费别 = Nvl(rsTmp!屏蔽费别, 0) = 1
        .执行科室 = Nvl(rsTmp!执行科室, 0)
        .服务对象 = Nvl(rsTmp!服务对象, 0)
        .补充摘要 = Nvl(rsTmp!补充摘要, 0) = 1
        .跟踪在用 = Nvl(rsTmp!跟踪在用, 0) = 1
        .要求审批 = Nvl(rsTmp!要求审批, 0) = 1
        .录入限量 = Val("" & rsTmp!录入限量)
        .中药形态 = Val(Nvl(rsTmp!中药形态))
        .商品名 = Nvl(rsTmp!商品名)
    End With
    Set GetInputDetail = objDetail
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckDuty(Optional tmpDetail As Detail, Optional blnCommon As Boolean = True, Optional str医疗付款 As String) As Integer
'功能：检查指定药品行的职务是否与当前医生的职务相匹配
'参数：tmpDetail=输入的项目,不传为所有行;
'      blnCommon=是否正常的检查或只针对医保及公费病人的检查
'      str医疗付款=单据一行公费或医保病人检查时,要传入
'返回：不匹配的行,0为正确
'说明：职务：1=正高,2=副高,3=中级,4=助理/师级,5=员/士,9=待聘
    Dim i As Long, int职务A As Integer, int职务B As Integer
    Dim strTmp As String, byt医疗付款码 As Byte, strAllDuty As String
    
    
    If cbo开单人.ListIndex = -1 Then Exit Function
    strAllDuty = "正高,副高,中级,助理/师级,员/士,,,,待聘"
    Call GetOperatorInfo(mrs开单人, mobjBill.开单人, , int职务A)
        
    If tmpDetail Is Nothing Then
        For i = 1 To mobjBill.Details.Count
            If InStr(",5,6,7,", mobjBill.Details(i).收费类别) > 0 Then
                If Not blnCommon Then
                    If mobjBill.Details(i).医疗付款 <> "" Then
                        '医保或公费病人
                        '问题:45605
                         If zlIsCheckMedicinePayMode(mobjBill.Details(i).医疗付款) Then

                            int职务B = Val(Right(mobjBill.Details(i).Detail.处方职务, 1))
                            If int职务B > 0 Then
                                If int职务A = 0 Then
                                    strTmp = "第 " & i & " 行病人:" & mobjBill.Details(i).姓名 & ",医疗付款方式为:" & mobjBill.Details(i).医疗付款 & "," & _
                                        vbCrLf & "药品""" & mobjBill.Details(i).Detail.名称 & """要求医生职务至少为""" & Split(strAllDuty, ",")(int职务B - 1) & """,而当前医生未设置职务！"
                                    CheckDuty = 1
                                ElseIf int职务B < int职务A Then
                                    strTmp = "第 " & i & " 行病人:" & mobjBill.Details(i).姓名 & ",医疗付款方式为:" & mobjBill.Details(i).医疗付款 & "," & _
                                        vbCrLf & "药品""" & mobjBill.Details(i).Detail.名称 & """要求医生职务为""" & Split(strAllDuty, ",")(int职务B - 1) & """以上,而当前医生职务为""" & Split(strAllDuty, ",")(int职务A - 1) & """！"
                                    CheckDuty = i: Exit For
                                End If
                            End If
                        End If
                    End If
                Else
                    int职务B = Val(Left(mobjBill.Details(i).Detail.处方职务, 1))
                    If int职务B > 0 Then
                        If int职务A = 0 Then
                            strTmp = "第 " & i & " 行药品""" & mobjBill.Details(i).Detail.名称 & """要求医生职务至少为""" & Split(strAllDuty, ",")(int职务B - 1) & """,而当前医生未设置职务！"
                            CheckDuty = 1
                        ElseIf int职务B < int职务A Then
                            strTmp = "第 " & i & " 行药品""" & mobjBill.Details(i).Detail.名称 & """要求医生职务为""" & Split(strAllDuty, ",")(int职务B - 1) & """以上,而当前医生职务为""" & Split(strAllDuty, ",")(int职务A - 1) & """！"
                            CheckDuty = i: Exit For
                        End If
                    End If
                End If
            End If
        Next
    Else
        If InStr(",5,6,7,", tmpDetail.类别) = 0 Then Exit Function
        If Not blnCommon Then
            If str医疗付款 = "" Then Exit Function
            '医保或公费病人
            '问题:45605
             If zlIsCheckMedicinePayMode(str医疗付款) = False Then Exit Function
            int职务B = Val(Right(tmpDetail.处方职务, 1))
            If int职务B > 0 Then
                If int职务A = 0 Then
                    strTmp = "当前病人医疗付款方式为:" & str医疗付款 & "," & _
                        vbCrLf & "药品""" & tmpDetail.名称 & """要求医生职务至少为""" & Split(strAllDuty, ",")(int职务B - 1) & """,而当前医生未设置职务！"
                    CheckDuty = 1
                ElseIf int职务B < int职务A Then
                    strTmp = "当前病人医疗付款方式为:" & str医疗付款 & "," & _
                        vbCrLf & "药品""" & tmpDetail.名称 & """要求医生职务为""" & Split(strAllDuty, ",")(int职务B - 1) & """以上,而当前医生职务为""" & Split(strAllDuty, ",")(int职务A - 1) & """！"
                    CheckDuty = 1
                End If
            End If
        Else
            int职务B = Val(Left(tmpDetail.处方职务, 1))
            If int职务B > 0 Then
                If int职务A = 0 Then
                    strTmp = "药品""" & tmpDetail.名称 & """要求医生职务至少为""" & Split(strAllDuty, ",")(int职务B - 1) & """,而当前医生未设置职务！"
                    CheckDuty = 1
                ElseIf int职务B < int职务A Then
                    strTmp = "药品""" & tmpDetail.名称 & """要求医生职务为""" & Split(strAllDuty, ",")(int职务B - 1) & """以上,而当前医生职务为""" & Split(strAllDuty, ",")(int职务A - 1) & """！"
                    CheckDuty = 1
                End If
            End If
        End If
    End If
    
    If CheckDuty > 0 Then MsgBox strTmp, vbInformation, gstrSysName
End Function

Private Function Check费用类型(Optional intRow As Integer) As Boolean
'功能：根据当前病人的类型判断指定行的项目是否可以输入,适用于所有类别的项目
    Dim strSQL As String
    Dim i As Long, bytType As Byte
    Dim rsTmp As New ADODB.Recordset
    Dim rs医保 As New ADODB.Recordset
    Dim rs公费 As New ADODB.Recordset
    Dim bln医保 As Boolean, bln公费 As Boolean
    
    Check费用类型 = True
        
    On Error GoTo errH
    '无法检查
    If intRow > 0 Then
        If mobjBill.Details(intRow).医疗付款 = "" Then Exit Function
        '医保或公费病人
        '问题:45605
        If zlIsCheckMedicinePayMode(mobjBill.Details(intRow).医疗付款, bln医保, bln公费) = False Then Exit Function
        bytType = IIf(bln医保, 1, 2)
        
        '读取检查数据
        If bytType = 1 Then
            strSQL = "Select 编码,名称,简码,性质,缺省标志 From 费用类型 Where 编码 In(" & gstr医保费用类型 & ") Order by 编码"
        Else
            strSQL = "Select 编码,名称,简码,性质,缺省标志 From 费用类型 Where 编码 In(" & gstr公费费用类型 & ") Order by 编码"
        End If
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        If rsTmp.EOF Then Exit Function
    
        If mobjBill.Details(intRow).Detail.类型 = "" Then
            MsgBox """" & mobjBill.Details(intRow).Detail.名称 & """的费用类型未设置！", vbInformation, gstrSysName
            Check费用类型 = False
        Else
            rsTmp.Filter = "名称='" & mobjBill.Details(intRow).Detail.类型 & "'"
            If rsTmp.EOF Then
                MsgBox """" & mobjBill.Details(intRow).Detail.名称 & """的费用类型为""" & _
                    mobjBill.Details(intRow).Detail.类型 & """,不是" & _
                    IIf(bytType = 1, "医保", "公费") & "费用类型！", vbInformation, gstrSysName
                Check费用类型 = False
            End If
        End If
    Else
        strSQL = "Select 编码,名称,简码,性质,缺省标志 From 费用类型 Where 编码 In(" & gstr医保费用类型 & ") Order by 编码"
        Call zlDatabase.OpenRecordset(rs医保, strSQL, Me.Caption)
        
        strSQL = "Select 编码,名称,简码,性质,缺省标志 From 费用类型 Where 编码 In(" & gstr公费费用类型 & ") Order by 编码"
        Call zlDatabase.OpenRecordset(rs公费, strSQL, Me.Caption)
        
        For i = 1 To mobjBill.Details.Count
        
            Call zlIsCheckMedicinePayMode(mobjBill.Details(i).医疗付款, bln医保, bln公费)
            bytType = IIf(bln医保, 1, IIf(bln公费, 2, 0))
            
            If InStr(",1,2,", bytType) > 0 Then
                Set rsTmp = Nothing
                If bytType = 1 Then
                    rs医保.Filter = 0
                    Set rsTmp = rs医保
                Else
                    rs公费.Filter = 0
                    Set rsTmp = rs公费
                End If
                If Not rsTmp.EOF Then
                    If mobjBill.Details(i).Detail.类型 = "" Then
                        If MsgBox("单据中第 " & i & " 行项目""" & mobjBill.Details(i).Detail.名称 & """的费用类型未设置！" & vbCrLf & "确实要保存单据吗？", _
                            vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Check费用类型 = False: Exit For
                        End If
                    Else
                        rsTmp.Filter = "名称='" & mobjBill.Details(i).Detail.类型 & "'"
                        If rsTmp.EOF Then
                            If MsgBox("单据中第 " & i & " 行项目""" & mobjBill.Details(i).Detail.名称 & """的费用类型为""" & _
                                mobjBill.Details(i).Detail.类型 & """,不是" & _
                                IIf(bytType = 1, "医保", "公费") & "费用类型！" & vbCrLf & "确实要保存单据吗？", _
                                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                Check费用类型 = False: Exit For
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ReCalcInsure()
'功能：修改单据时,重新计算统筹金额及更新相关信息
    Dim i As Long, j As Long, dblAllTime As Double
    Dim strInfo As String
    Dim intInsure As Integer
    
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).病人ID <> 0 And Val(mobjBill.Details(i).发药窗口) > 0 Then
            For j = 1 To mobjBill.Details(i).InComes.Count
                intInsure = Val(mobjBill.Details(i).发药窗口)
                If intInsure <> 0 Then
                    dblAllTime = mobjBill.Details(i).付数 * mobjBill.Details(i).数次
                    If InStr(",5,6,7,", mobjBill.Details(i).收费类别) > 0 Then
                        If gbln住院单位 Then dblAllTime = dblAllTime * mobjBill.Details(i).Detail.住院包装
                    End If
                
                    strInfo = gclsInsure.GetItemInsure(mobjBill.Details(i).病人ID, mobjBill.Details(i).收费细目ID, mobjBill.Details(i).InComes(j).实收金额, False, intInsure, _
                        mobjBill.Details(i).摘要 & "||" & dblAllTime)
                    If strInfo <> "" Then
                        mobjBill.Details(i).保险项目否 = Val(Split(strInfo, ";")(0)) <> 0
                        mobjBill.Details(i).保险大类ID = Val(Split(strInfo, ";")(1))
                        mobjBill.Details(i).InComes(j).统筹金额 = Val(Split(strInfo, ";")(2))
                        mobjBill.Details(i).保险编码 = CStr(Split(strInfo, ";")(3))
                        
                        If UBound(Split(strInfo, ";")) >= 4 Then
                            If CStr(Split(strInfo, ";")(4)) <> "" Then mobjBill.Details(i).摘要 = CStr(Split(strInfo, ";")(4))
                            If UBound(Split(strInfo, ";")) >= 5 Then
                                If Split(strInfo, ";")(5) <> "" Then mobjBill.Details(i).Detail.类型 = Split(strInfo, ";")(5)
                            End If
                        End If
                    End If
                End If
            Next
        End If
    Next
End Sub

Private Function PhysicExist(objDetail As Detail, intRow As Integer, lng病人ID As Long) As Boolean
'功能：判断指定药品在单据中是否已经存在
'参数：objDetail=项目,intRow=要判断的行
'说明：时价或分批药品在同一药房禁止重复输入(这里仅提示,保存时禁止)
    Dim i As Long
    For i = 1 To mobjBill.Details.Count
        If i <> intRow And InStr(",4,5,6,7,", mobjBill.Details(i).收费类别) > 0 Then
            If mobjBill.Details(i).Detail.ID = objDetail.ID Then
                If (mobjBill.Details(i).Detail.分批 Or mobjBill.Details(i).Detail.变价) _
                    And (objDetail.分批 Or objDetail.变价) Then
                    If objDetail.类别 = "4" Then
                        If MsgBox("卫生材料""" & objDetail.名称 & """在单据中已经输入,要继续吗？" & _
                            vbCrLf & vbCrLf & "注意：该卫生材料为分批或时价材料,重复输入时必须保证它们的执行药房不同。", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            PhysicExist = True
                        End If
                    Else
                        If MsgBox("药品""" & objDetail.名称 & """在单据中已经输入,要继续吗？" & _
                            vbCrLf & vbCrLf & "注意：该药品为分批或时价药品,重复输入时必须保证它们的执行药房不同。", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            PhysicExist = True
                        End If
                    End If
                    Exit Function
                ElseIf mobjBill.Details(i).病人ID = lng病人ID Then
                    If objDetail.类别 = "4" Then
                        If MsgBox("该病人已经输入卫生材料""" & objDetail.名称 & """,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            PhysicExist = True
                        End If
                    Else
                        If MsgBox("该病人已经输入药品""" & objDetail.名称 & """,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            PhysicExist = True
                        End If
                    End If
                    Exit Function
                End If
            End If
        End If
    Next
End Function


Private Function Check执行科室() As Integer
    Dim i As Long
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).执行部门ID = 0 Or Bill.TextMatrix(i, BillCol.执行科室) = "" Then
            If Not (InStr(",5,6,7,", mobjBill.Details(i).收费类别) > 0 And gbln分离发药) Then
                Check执行科室 = i: Exit Function
            End If
        End If
    Next
End Function

Private Function Check服务对象() As Integer
'功能：检查记帐表中病人的记帐费用项目的服务对象是否一致
'说明：因为加入了门诊留观病人,所以有此检查
'返回：不一致的费用行,为0时正常
    Dim i As Integer
    
    If mrsInfo.State = 0 Then Exit Function
    With mobjBill
        For i = 1 To .Details.Count
            If InStr(",0,2,", .Details(i).病人性质) > 0 Then
                '住院病人或住院留观病人,不能用只服务于门诊的项目
                If .Details(i).Detail.服务对象 = 1 Then
                    MsgBox "第 " & i & " 行项目""" & .Details(i).Detail.名称 & """仅服务于门诊,病人""" & .Details(i).姓名 & """不能使用.", vbInformation, gstrSysName
                    Check服务对象 = i: Exit Function
                End If
            ElseIf InStr(",1,-1,", .Details(i).病人性质) > 0 Then
                '门诊或出院病人(医技记帐)或门诊留观病人,不能用只服务于住院的项目
                If .Details(i).Detail.服务对象 = 2 Then
                    MsgBox "第 " & i & " 行项目""" & .Details(i).Detail.名称 & """仅服务于住院,病人""" & .Details(i).姓名 & """不能使用.", vbInformation, gstrSysName
                    Check服务对象 = i: Exit Function
                End If
            End If
            If .Details(i).Detail.服务对象 = 0 Then
                MsgBox "第 " & i & " 行项目""" & .Details(i).Detail.名称 & """不服务于病人,病人""" & .Details(i).姓名 & """不能使用.", vbInformation, gstrSysName
                Check服务对象 = i: Exit Function
            End If
        Next
    End With
End Function

Private Sub SetIntureColor()
'功能：导入单据后对医保病人设置为红色
    Dim intRow As Integer, intCol As Integer, i As Integer
    
    intRow = Bill.Row: intCol = Bill.Col
    Bill.Col = BillCol.姓名
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).发药窗口 <> "" Then
            Bill.Row = i
            Bill.MsfObj.CellForeColor = vbRed
        End If
    Next
    Bill.Row = intRow: Bill.Col = intCol
End Sub
Private Function Get开单科室ID() As Long
    If cbo开单科室.ListIndex <> -1 Then
        Get开单科室ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
    Else
        Get开单科室ID = UserInfo.部门ID
    End If
End Function
Private Function Get病人来源(ByVal lngRow As Long) As Integer
'功能：获取当前病人的来源(因为可以对门诊留观病人记帐)
    Dim int病人性质 As Integer
    
    int病人性质 = -2
    If mobjBill.Details.Count >= lngRow Then
        int病人性质 = mobjBill.Details(lngRow).病人性质
    ElseIf mrsInfo.State = 1 Then
        int病人性质 = mrsInfo!病人性质
    ElseIf Bill.TextMatrix(lngRow, BillCol.姓名) <> "" And lngRow > 1 Then
        int病人性质 = mobjBill.Details(lngRow - 1).病人性质
    End If
    If int病人性质 <> -2 Then
        If int病人性质 = 0 Or int病人性质 = 2 Then
            Get病人来源 = 2
        ElseIf int病人性质 = 1 Or int病人性质 = -1 Then
            Get病人来源 = 1 '门诊病人(医技记帐)或门诊留观病人
        End If
    Else
        Get病人来源 = 2 '缺省为2
    End If
End Function

Private Sub zlReSetDrawDrugDept()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据相应的规则,重新获取领药部门
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-07-29 18:23:12
    '问题:24729
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '3)  医技科室记帐时，对应的领药部门固定确定为主界面所选定的医技科室。(单据中应只提供主界面科室和病人科室可选)
    '4)  住院记帐、科室分散记帐，可能由病区使用，也可能由医技科室使用。
    '    a)  判断当前操作员所属科室，如果不属于医技性质的科室，则领药部门固定为病人病区。(检查、检验、手术、治疗、营养)
    '    b)  如果操作员属于医技性质的科室，则在单据界面上增加"领药部门"选择框，可选择范围为操作员所属的医技性质的科室(可能多个)，缺省与开单科室相同。
    If mbytUseType = 2 Then
        '医技科室记帐时，对应的领药部门固定确定为主界面所选定的医技科室。(单据中应只提供主界面科室和病人科室可选)
        mobjBill.领药部门ID = mlngDeptID: Exit Sub
    End If
    If mrs领药部门.RecordCount = 0 Then
        '判断当前操作员所属科室，如果不属于医技性质的科室，则领药部门固定为病人病区。(检查、检验、手术、治疗、营养)
        mobjBill.领药部门ID = mobjBill.病区ID: Exit Sub
    End If
    '如果操作员属于医技性质的科室，则在单据界面上增加"领药部门"选择框，可选择范围为操作员所属的医技性质的科室(可能多个)，缺省与开单科室相同。
    If mrs领药部门.RecordCount = 1 Then
        '只有一个部分,肯定是他
        If mrs领药部门.EOF Then mrs领药部门.MoveFirst
         mobjBill.领药部门ID = Val(Nvl(mrs领药部门!ID)): Exit Sub
    End If
    '选择的科室是哪个就是哪个
    With cboDrawDept
        If .ListIndex < 0 Then Exit Sub
        If mobjBill.领药部门ID <> .ItemData(.ListIndex) Then mobjBill.领药部门ID = .ItemData(.ListIndex): Exit Sub
    End With
End Sub
Private Sub zlLoadDrawDeptData(ByVal bytUseType As Byte, Optional ByVal lngDeptID As Long = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:
    '入参:bytUseType:记帐单用途,0-普通记帐,1-按科室分散记帐,2-医技科室记帐
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-07-29 15:05:18
    '问题:24729
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, strSQL As String
    
    '3)  医技科室记帐时，对应的领药部门固定确定为主界面所选定的医技科室。(单据中应只提供主界面科室和病人科室可选)
    '4)  住院记帐、科室分散记帐，可能由病区使用，也可能由医技科室使用。
    '    a)  判断当前操作员所属科室，如果不属于医技性质的科室，则领药部门固定为病人病区。(检查、检验、手术、治疗、营养)
    '    b)  如果操作员属于医技性质的科室，则在单据界面上增加"领药部门"选择框，可选择范围为操作员所属的医技性质的科室(可能多个)，缺省与开单科室相同。
    
    On Error GoTo errHandle
    
    '医技科室
    If bytUseType = 2 Then
        '3)  医技科室记帐时，对应的领药部门固定确定为主界面所选定的医技科室。(单据中应只提供主界面科室和病人科室可选)
        strSQL = "Select ID,编码,名称 From 部门表 where id=[2]"
    Else
        strSQL = _
            " Select distinct  A.ID, A.编码,A.名称   " & vbNewLine & _
            " From 部门表 A, 部门性质说明 B,部门人员 C" & vbNewLine & _
            " Where (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)  " & _
            "       And A.ID = B.部门id and a.id=C.部门ID and C.人员id=[1] " & vbNewLine & _
            "       And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
            "       AND B.工作性质 IN('检查','检验','手术','治疗','营养') " & _
            " Order by 编码"
    End If
    Set mrs领药部门 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID, lngDeptID)
    With mrs领药部门
        cboDrawDept.Clear
        Do While Not .EOF
            cboDrawDept.AddItem IIf(zlIsShowDeptCode, Nvl(!编码) & "-", "") & Nvl(!名称)
            cboDrawDept.ItemData(cboDrawDept.NewIndex) = Val(Nvl(!ID))
            If Val(Nvl(!ID)) = UserInfo.部门ID Then cboDrawDept.ListIndex = cboDrawDept.NewIndex
            .MoveNext
        Loop
        If .RecordCount <> 0 And cboDrawDept.ListIndex < 0 Then cboDrawDept.ListIndex = 0
        If .RecordCount <> 0 Then .MoveFirst
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub SetDrawDrugDeptVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置领药部门的visibled属性
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-07-29 19:07:38
    '问题:24729
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    ' mbytUseType As Byte '记帐单用途,0-普通记帐,1-按科室分散记帐,2-医技科室记帐
    
    '3)  医技科室记帐时，对应的领药部门固定确定为主界面所选定的医技科室。(单据中应只提供主界面科室和病人科室可选)
    If mbytUseType = 2 Then
        fraDrawDept.Visible = False
    ElseIf chkCancel.Value = 1 Then
        '销帐也不能看见
        fraDrawDept.Visible = False
    Else
        'mbytInState As Byte '0-执行,1-查阅,2-调整,3-销帐
        ' gbytBilling:0-记帐,1-划价,2-审核
        fraDrawDept.Visible = mrs领药部门.RecordCount > 1 And (mbytInState = 0 And gbytBilling <> 2)
    End If
    Call Form_Resize
End Sub
Private Sub SetDrawDrugDeptEnabled()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置领药部门的Enabled属性
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-07-31 11:55:07
    '问题:24729
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, blnHaveDrug As Boolean '存在药品
    
    '如果没有领用部门的选择,则直接退出
    If fraDrawDept.Visible = False Then cboDrawDept.Enabled = False: lblDrawDrugDept.Enabled = False: Exit Sub
    blnHaveDrug = False
    For i = 1 To mobjBill.Details.Count
        If InStr(1, ",5,6,7,", "," & mobjBill.Details(i).收费类别 & ",") > 0 Then
            blnHaveDrug = True
            Exit For
        End If
    Next
    cboDrawDept.Enabled = blnHaveDrug: lblDrawDrugDept.Enabled = blnHaveDrug
End Sub
Public Function zl获取中药形态(ByVal lng病人ID As Long, Optional ByVal lngRow As Long = -1, Optional blnOnly中成药 As Boolean = False) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取单据是否录入了中草药的
    '入参:intPage-当前第几页
    '     blnOnly中成药-仅判断是否有中成药(对配方时判断有效):原因是中划药在配方中已经存在,就不需要检查
    '     lngRow-当前操作的行
    '出参:
    '返回:录入了中草药的,则返回中药形态属性(0-散装,1-饮片,2-免煎剂),否则返回-1 表示还没有录入中药形态项目
    '编制:刘兴洪
    '日期:2010-02-02 11:44:17
    '问题:27816
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strTemp As String
    
    zl获取中药形态 = -1
    '如果未指定页,则用当前页
    If mobjBill Is Nothing Then Exit Function
    strTemp = IIf(blnOnly中成药, ",6,", ",6,7,")
    Err = 0: On Error GoTo ErrHand:
    
    With mobjBill.Details
        For i = 1 To .Count
            If InStr(1, strTemp, "," & .Item(i).收费类别 & ",") > 0 And .Item(i).收费细目ID <> 0 And i <> lngRow And .Item(i).病人ID = lng病人ID Then
                zl获取中药形态 = .Item(i).Detail.中药形态
                Exit Function
            End If
        Next
    End With
ErrHand:
End Function
Private Function zlGetBillOtherRowNumToTal(lng病人ID As Long, lng主页ID As Long, lng细目ID As Long, _
    Optional blnOnly负数数次 As Boolean, Optional ByVal lngCurRow As Long = 0) As Double
    '------------------------------------------------------------------------------------------------------------------------
    '功能：获取指定行以外的其他行的合计金额
    '入参：lng病人ID-病人ID
    '         lngCurRow-当前行(为零时,为所有行)
    '出参：
    '返回：
    '编制：刘兴洪
    '日期：2010-05-05 16:09:12
    '说明：29412
    '------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, dbl数次 As Double
    
    dbl数次 = 0
    For i = 1 To mobjBill.Details.Count
        With mobjBill.Details(i)
            If .病人ID = lng病人ID And .主页ID = lng主页ID And .收费细目ID = lng细目ID And i <> lngCurRow Then
                If blnOnly负数数次 Then
                    If .数次 < 0 And .执行部门ID <> 0 Then
                        dbl数次 = dbl数次 + .数次 * .付数
                    End If
                Else
                    dbl数次 = dbl数次 + .数次 * .付数       '* IIf(InStr(",5,6,7,", .收费类别) > 0 And gbln住院单位, .Detail.住院包装, 1)
                End If
            End If
        End With
    Next
     zlGetBillOtherRowNumToTal = dbl数次
End Function

Private Function CheckBillNegative() As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：检查单据中输入的负数数量及退回科室是否正确
    '返回：合法,返回true,否则返回False
    '编制：刘兴洪
    '日期：2010-05-05 17:02:57
    '说明：29412
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long, intR As Long
    Dim strItems As String, str部门 As String, strValues(0 To 10) As String
    Dim str单位 As String, dbl数量 As Double, dbl已结数量 As Double, dbl数次合计 As Double
    Dim strSubTable As String
     
    '问题:26951
    If InStr(1, mstrPrivsOpt, ";负数记帐不检查发生项目;") > 0 Then
        '对于负数冲销时不检查本次住院发生的项目数量,有此权限,允许录入病人未曾发生的费用项目进行冲销,否则检查本次住院发生的项目数量才能冲销
        CheckBillNegative = True: Exit Function
    End If
    
    strItems = ""
    strSubTable = ""
    intR = 0
    
    For i = 1 To mobjBill.Details.Count
        With mobjBill.Details(i)
            If .数次 < 0 And .执行部门ID <> 0 Then
                If Len(strItems) > 2000 Then
                    If intR <= 10 Then
                        strValues(intR) = Mid(strItems, 2)
                        strSubTable = strSubTable & " Union ALL " & _
                        "  Select To_Number(Substr(Column_Value, 1, Instr(Column_Value, ';') - 1)) As 病人id, " & _
                        "          To_Number(Substr(Column_Value, Instr(Column_Value, ';') + 1, Instr(Column_Value, ':') - 1- Instr(Column_Value, ';'))) As 主页id, " & _
                        "          To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1, Instr(Column_Value, '_') - 1- Instr(Column_Value, ':') )) As 收费细目id, " & _
                        "          To_Number(Substr(Column_Value, Instr(Column_Value, '_') + 1)) As 执行部门id, 0 As 数量,0 as 结帐数量" & _
                        " From Table(Cast(f_str2list([" & intR + 2 & "]) As ZLTOOLS.t_strlist))"
                    Else
                        strSubTable = strSubTable & " Union ALL " & _
                        "  Select To_Number(Substr(Column_Value, 1, Instr(Column_Value, ';') - 1)) As 病人id, " & _
                        "          To_Number(Substr(Column_Value, Instr(Column_Value, ';') + 1, Instr(Column_Value, ':') - 1- Instr(Column_Value, ';'))) As 主页id, " & _
                        "          To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1, Instr(Column_Value, '_') - 1- Instr(Column_Value, ':') )) As 收费细目id, " & _
                        "          To_Number(Substr(Column_Value, Instr(Column_Value, '_') + 1)) As 执行部门id, 0 As 数量,0 as 结帐数量" & _
                        " From Table(Cast(f_str2list('" & Mid(strItems, 2) & "') As ZLTOOLS.t_strlist))"
                    End If
                    strItems = "": intR = intR + 1
                End If
                strItems = strItems & "," & .病人ID & ";" & .主页ID & ":" & .收费细目ID & "_" & .执行部门ID & ""
            End If
        End With
    Next
    
    If strItems <> "" Then
        If intR <= 10 Then
            strValues(intR) = Mid(strItems, 2)
            strSubTable = strSubTable & " Union ALL " & _
            "  Select To_Number(Substr(Column_Value, 1, Instr(Column_Value, ';') - 1)) As 病人id, " & _
            "          To_Number(Substr(Column_Value, Instr(Column_Value, ';') + 1, Instr(Column_Value, ':') - 1- Instr(Column_Value, ';'))) As 主页id, " & _
            "          To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1, Instr(Column_Value, '_') - 1- Instr(Column_Value, ':') )) As 收费细目id, " & _
            "          To_Number(Substr(Column_Value, Instr(Column_Value, '_') + 1)) As 执行部门id, 0 As 数量,0 as 结帐数量" & _
            " From Table(Cast(f_str2list([" & intR + 2 & "]) As ZLTOOLS.t_strlist))"
        Else
            strSubTable = strSubTable & " Union ALL " & _
            "  Select To_Number(Substr(Column_Value, 1, Instr(Column_Value, ';') - 1)) As 病人id, " & _
            "          To_Number(Substr(Column_Value, Instr(Column_Value, ';') + 1, Instr(Column_Value, ':') - 1- Instr(Column_Value, ';'))) As 主页id, " & _
            "          To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1, Instr(Column_Value, '_') - 1- Instr(Column_Value, ':') )) As 收费细目id, " & _
            "          To_Number(Substr(Column_Value, Instr(Column_Value, '_') + 1)) As 执行部门id, 0 As 数量,0 as 结帐数量" & _
            " From Table(Cast(f_str2list('" & Mid(strItems, 2) & "') As ZLTOOLS.t_strlist))"
        End If
    End If
    CheckBillNegative = True
    If strSubTable = "" Then Exit Function
    strSubTable = Mid(strSubTable, 11)
    
    strSQL = " " & _
    " with C1 as (" & strSubTable & ") " & vbCrLf & _
    " Select  /*+ RULE */ A.病人ID,A.主页ID, A.收费细目ID,A.执行部门ID,  " & _
    "             Nvl(Sum(Decode(A.记录性质, 2, 1, 3, 1, 0) * Nvl(A.付数, 1) * A.数次), 0) As 数量, " & _
     "            Sum(Decode(nvL(Mod(M.记录状态 , 3),1),  0, 1, 1, 1, -1) * Decode(A.结帐id, Null, 0, 1) * Nvl(付数, 1) * 数次) As 结帐数量 " & _
     "     From 住院费用记录 A, 病人结帐记录 M " & _
     "     Where  A.结帐id = M.ID(+)  And A.记帐费用=1 And A.价格父号 Is Null   " & IIf(gbytBilling = 0, " And A.记录状态<>0", "") & _
     "             And (A.病人ID,A.主页ID,A.收费细目ID,执行部门ID,0,0) in (select * From C1) " & _
                    IIf(mstrInNO <> "", " And NO<>[1]", "") & _
     "     Group By A.病人ID,A.主页ID,A.收费细目ID,A.执行部门ID" & _
     "     Union ALL Select * From C1 "
    'strSQL = _
    " with C1 as (" & strSubTable & ") " & vbCrLf & _
    " Select  /*+ RULE */ A.病人ID,A.主页ID, A.收费细目ID,A.执行部门ID,Sum(Nvl(A.付数,1)*A.数次) as 数量, " & _
    "           Sum(decode(结帐ID,NULL,0,1)* Nvl(A.付数,1)*A.数次) as 结帐数量  " & _
    " From  住院费用记录 A " & _
    " Where 记帐费用=1 And 价格父号 is NULL   " & _
                IIf(gbytBilling = 0, " And 记录状态<>0", "") & _
                IIf(mstrInNO <> "", " And NO<>[1]", "") & _
    "           And (A.病人ID,A.主页ID,A.收费细目ID,执行部门ID,0,0) in (select * From C1) " & _
    " Group by A.病人ID,A.主页ID,A.收费细目ID,A.执行部门ID" & _
    " Union ALL Select * From C1"
    strSQL = "" & _
    "   Select 病人ID,主页ID,收费细目ID,执行部门ID,Sum(数量) as 数量,sum(nvl(结帐数量,0)) as 结帐数量 " & _
    "   From (" & strSQL & ") " & _
    "   Group by 病人ID,主页ID,收费细目ID,执行部门ID"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrInNO, strValues(0), strValues(1), strValues(2), strValues(3), strValues(4), strValues(5), strValues(6), strValues(7), strValues(8), strValues(9), strValues(10))
    For i = 1 To mobjBill.Details.Count
        With mobjBill.Details(i)
            If .数次 < 0 And .执行部门ID <> 0 Then
                rsTmp.Filter = " 病人ID=" & .病人ID & " And 主页ID = " & .主页ID & " And 收费细目ID = " & .收费细目ID & " And 执行部门ID = " & .执行部门ID
                
                If Not rsTmp.EOF Then
                    If InStr(",5,6,7,", .收费类别) > 0 And gbln住院单位 Then
                        str单位 = .Detail.住院单位
                        dbl数量 = Nvl(rsTmp!数量, 0) / .Detail.住院包装
                        dbl已结数量 = Val(Nvl(rsTmp!结帐数量)) / .Detail.住院包装
                    Else
                        str单位 = .Detail.计算单位
                        dbl数量 = Nvl(rsTmp!数量, 0)
                        dbl已结数量 = Val(Nvl(rsTmp!结帐数量))
                    End If
                    '可能存在两条相同的记录
                    '问题:29412
                    dbl数次合计 = Abs(.数次) * .付数
                    For j = i + 1 To mobjBill.Details.Count
                         If .收费细目ID = mobjBill.Details(j).收费细目ID And mobjBill.Details(j).病人ID = .病人ID And mobjBill.Details(j).主页ID = .主页ID _
                            And mobjBill.Details(j).数次 < 0 And mobjBill.Details(j).执行部门ID = .执行部门ID Then
                                dbl数次合计 = dbl数次合计 + Abs(.数次) * .付数
                         End If
                    Next
                    '问题:32106
                    If dbl数次合计 > dbl数量 - dbl已结数量 Then
                        Select Case gbytBillOpt '对已结帐的记帐单据的操作权限:0-允许,1-提醒,2-禁止。
                        Case 0  '允许
                            If dbl数次合计 > dbl数量 Then
                                    str部门 = GET部门名称(.执行部门ID, mrsUnit)
                                    MsgBox "第 " & i & " 行[" & .Detail.名称 & "]退回" & str部门 & "的数量 " & FormatEx(dbl数次合计, 5) & str单位 & _
                                        " 大于可销帐数量 " & FormatEx(dbl数量, 5) & str单位 & "。", vbInformation, gstrSysName
                                    CheckBillNegative = False: Exit Function
                            End If
                        Case 1   '提醒
                            str部门 = GET部门名称(.执行部门ID, mrsUnit)
                            If dbl数次合计 > dbl数量 Then
                                    MsgBox "第 " & i & " 行[" & .Detail.名称 & "]退回" & str部门 & "的数量 " & FormatEx(dbl数次合计, 5) & str单位 & _
                                        " 大于可销帐数量 " & FormatEx(dbl数量, 5) & str单位 & "。", vbInformation, gstrSysName
                                    CheckBillNegative = False: Exit Function
                            End If
                            
                            If MsgBox("第 " & i & " 行[" & .Detail.名称 & "]退回" & str部门 & "的数量 " & FormatEx(dbl数次合计, 5) & str单位 & _
                                " 中包含了已结部分(未结:" & FormatEx(dbl数量 - dbl已结数量, 5) & str单位 & "; 已结:" & FormatEx(dbl已结数量, 5) & str单位 & ") 。" & vbCrLf & _
                                " 是否继续?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                                CheckBillNegative = False: Exit Function
                            End If
                        Case 2   '禁止
                            str部门 = GET部门名称(.执行部门ID, mrsUnit)
                            MsgBox "第 " & i & " 行[" & .Detail.名称 & "]退回" & str部门 & "的数量 " & FormatEx(dbl数次合计, 5) & str单位 & _
                                " 大于可销帐数量 " & FormatEx(dbl数量 - dbl已结数量, 5) & str单位 & "。", vbInformation, gstrSysName
                                CheckBillNegative = False: Exit Function
                        End Select
                    End If
                Else
                    MsgBox "第 " & i & " 行[" & .Detail.名称 & "]可销帐数量为零，不允许冲销。", vbInformation, gstrSysName
                    CheckBillNegative = False: Exit Function
                End If
            End If
        End With
    Next
    CheckBillNegative = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Sub SetStatuPatiInfor(ByVal str姓名 As String, ByVal dbl预交 As Double, dblFee As Double, dbl剩余 As Double, Optional dbl应收 As Double = 0)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：设置状态栏信息
    '编制：刘兴洪
    '日期：2010-06-23 11:28:31
    '说明：30604
    '------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String
    picStatuPancl.Visible = False
    '78082:李南春,2014/10/10,预交金额显示
    strTemp = str姓名 & "预交:" & Format(Val(dbl预交), "0.00")
    strTemp = strTemp & "/费用:" & Format(dblFee, gstrDec)
    strTemp = strTemp & "/剩余:" & Format(dbl剩余, "0.00")
    If dbl应收 <> 0 Then
        strTemp = strTemp & "/应收款:" & Format(dbl应收, "0.00")
    End If
    
    sta.Panels(3).Text = strTemp
    Call MoveStatuPatiInfor
    If dbl剩余 <= 0 Then
        lblStatuPati.Caption = strTemp
        lblStatuPati.AutoSize = True
        picStatuPancl.Visible = True
    End If
    Err = 0
End Sub
Private Sub MoveStatuPatiInfor()
    '------------------------------------------------------------------------------------------------------------------------
    '功能：移动状态栏的病人欠费信息
    '入参：
    '出参：
    '返回：
    '编制：刘兴洪
    '日期：2010-06-23 13:51:45
    '说明：30604
    '------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    With picStatuPancl
        .Left = sta.Panels(3).Left + 50
        .Width = sta.Panels(3).Width - 10
        .Top = Me.ScaleHeight - .Height - 10
    End With
End Sub

Private Sub SetDetailtStock(ByVal lng执行科室ID As Long, ByRef objDetail As Detail)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：设置明细的库存数据
    '编制：刘兴洪
    '日期：2010-07-12 14:27:51
    '说明：
    '      bug:31374
    '------------------------------------------------------------------------------------------------------------------------
    Dim str药房IDs As String, dblStock As Double
    
    '获取库存
    '不处理非药品或卫材
    If InStr(1, "5,6,7,4", objDetail.类别) = 0 Then Exit Sub
    If objDetail.类别 = "4" And objDetail.跟踪在用 = False Then Exit Sub
    If objDetail.类别 = "4" Then
        '卫材
        dblStock = GetStock(objDetail.ID, lng执行科室ID)
        objDetail.库存 = dblStock
        Exit Sub
    End If
    
    If Not gbln分离发药 Then
        dblStock = GetStock(objDetail.ID, lng执行科室ID)
        If gbln住院单位 Then
            dblStock = dblStock / objDetail.住院包装
        End If
        objDetail.库存 = dblStock  '记录当前行药品库存
        Exit Sub
    End If
    str药房IDs = Decode(mobjDetail.类别, "5", gstr西药房, "6", gstr成药房, "7", gstr中药房)
    If str药房IDs <> "" Then
        dblStock = GetMultiStock(mobjDetail.ID, str药房IDs)
        If gbln住院单位 Then
            dblStock = dblStock / mobjDetail.住院包装
        End If
        mobjDetail.库存 = dblStock
    End If
End Sub
Private Function ReadDrugAndStuffStock(ByVal lng库房ID As Long, ByRef objDetail As Detail) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取药品和卫材料的库存信息
    '入参:lng库房ID-库房ID
    '出参:objDetail-Detail对象
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-10 09:34:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblStock As Double, str药房IDs As String
    
    On Error GoTo errHandle
    If objDetail Is Nothing Then Exit Function
    '非药品和卫生材料的，直接返回True
    If InStr(",5,6,7,4,", objDetail.类别) = 0 Then ReadDrugAndStuffStock = True: Exit Function
    If objDetail.类别 = "4" And objDetail.跟踪在用 = False Then ReadDrugAndStuffStock = True: Exit Function
   
    If objDetail.类别 = "4" And objDetail.跟踪在用 Then
        dblStock = GetStock(objDetail.ID, lng库房ID)
        objDetail.库存 = dblStock
        Call ShowStock(objDetail.名称, objDetail.库存)
        ReadDrugAndStuffStock = True: Exit Function
    End If
    If InStr(",5,6,7,", objDetail.类别) > 0 Then
        '当前行药品库存
        If Not gbln分离发药 Then
            dblStock = GetStock(objDetail.ID, lng库房ID)
            If gbln住院单位 Then
                dblStock = dblStock / objDetail.住院包装
            End If
            objDetail.库存 = dblStock
            Call ShowStock(objDetail.名称, objDetail.库存)
        Else
            str药房IDs = Decode(objDetail.类别, "5", gstr西药房, "6", gstr成药房, "7", gstr中药房)
            If str药房IDs <> "" Then
                dblStock = GetMultiStock(objDetail.ID, str药房IDs)
                
                If dblStock = 0 And gblnStock Then
                    MsgBox "[" & objDetail.名称 & "]的可用库存为零!", vbInformation, gstrSysName
                    Exit Function
                End If
                If gbln住院单位 Then
                    dblStock = dblStock / objDetail.住院包装
                End If
                objDetail.库存 = dblStock
                Call ShowStock(objDetail.名称, objDetail.库存)
            End If
        End If
    End If
    ReadDrugAndStuffStock = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
