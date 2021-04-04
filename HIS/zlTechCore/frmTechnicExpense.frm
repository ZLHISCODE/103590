VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.3#0"; "ZL9BillEdit.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Begin VB.Form frmTechnicExpense 
   AutoRedraw      =   -1  'True
   Caption         =   "病人计费处理"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTechnicExpense.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   25
      Top             =   7275
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2619
            MinWidth        =   882
            Picture         =   "frmTechnicExpense.frx":058A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13838
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   370
            MinWidth        =   88
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmTechnicExpense.frx":0E1E
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmTechnicExpense.frx":1458
            Key             =   "WB"
            Object.ToolTipText     =   "五笔(F7)"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picAppend 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2160
      Left            =   0
      ScaleHeight     =   2160
      ScaleWidth      =   11880
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   5115
      Width           =   11880
      Begin MSComctlLib.ImageList imgList 
         Left            =   7335
         Top             =   570
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   1
         ImageHeight     =   18
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTechnicExpense.frx":1A92
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "取消(&C)"
         Height          =   420
         Left            =   9780
         TabIndex        =   20
         ToolTipText     =   "热键:Esc"
         Top             =   1125
         Width           =   1680
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00C0C0C0&
         Caption         =   "确定(&O)"
         Height          =   420
         Left            =   7965
         TabIndex        =   19
         ToolTipText     =   "热键：F2"
         Top             =   1125
         Width           =   1680
      End
      Begin VB.Frame fraAppend 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   0
         TabIndex        =   35
         ToolTipText     =   "清除:F6"
         Top             =   -90
         Width           =   11880
         Begin VB.ComboBox cboBaby 
            Height          =   360
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   180
            Width           =   1800
         End
         Begin VB.CheckBox chk急诊 
            Caption         =   "急诊费用"
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   4440
            TabIndex        =   14
            Top             =   240
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.CheckBox chk加班 
            Caption         =   "加班(&A)"
            Height          =   270
            Left            =   120
            TabIndex        =   11
            Top             =   225
            Width           =   1170
         End
         Begin VB.ComboBox cbo开单人 
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   6555
            TabIndex        =   15
            Top             =   180
            Width           =   2085
         End
         Begin MSMask.MaskEdBox txtDate 
            Height          =   360
            Left            =   9360
            TabIndex        =   16
            Top             =   180
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   635
            _Version        =   393216
            AutoTab         =   -1  'True
            HideSelection   =   0   'False
            MaxLength       =   19
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "yyyy-MM-dd hh:mm:ss"
            Mask            =   "####-##-## ##:##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label lblBaby 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "婴儿费(&B)"
            Height          =   240
            Left            =   1440
            TabIndex        =   12
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label lbl开单人 
            AutoSize        =   -1  'True
            Caption         =   "开单人"
            Height          =   240
            Left            =   5790
            TabIndex        =   37
            Top             =   240
            Width           =   720
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            Caption         =   "时间"
            Height          =   240
            Left            =   8820
            TabIndex        =   36
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.Frame fraStat 
         Height          =   1770
         Left            =   3510
         TabIndex        =   38
         Top             =   390
         Width           =   3675
         Begin VB.TextBox txt实收 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Left            =   1065
            Locked          =   -1  'True
            TabIndex        =   18
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   1020
            Width           =   1845
         End
         Begin VB.TextBox txt应收 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   450
            Left            =   1065
            Locked          =   -1  'True
            TabIndex        =   17
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   405
            Width           =   1845
         End
         Begin VB.Label lbl实收 
            AutoSize        =   -1  'True
            Caption         =   "实收"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   315
            Left            =   270
            TabIndex        =   40
            Top             =   1095
            Width           =   690
         End
         Begin VB.Label lbl应收 
            AutoSize        =   -1  'True
            Caption         =   "应收"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   270
            TabIndex        =   39
            Top             =   480
            Width           =   690
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshMoney 
         Height          =   1635
         Left            =   0
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   525
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   2884
         _Version        =   393216
         Rows            =   5
         FixedCols       =   0
         RowHeightMin    =   320
         BackColorBkg    =   15466495
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         FormatString    =   "^         项目|^          金额"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame fraTitle 
      Height          =   1095
      Left            =   30
      TabIndex        =   23
      ToolTipText     =   "清除:F6"
      Top             =   -120
      Width           =   11865
      Begin VB.ComboBox cboNO 
         ForeColor       =   &H00C00000&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   660
         Width           =   1425
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   0
         X2              =   18000
         Y1              =   585
         Y2              =   585
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   30
         X2              =   18000
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "病人计费单"
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
         Left            =   225
         TabIndex        =   27
         ToolTipText     =   "清除:F6"
         Top             =   180
         Width           =   1875
      End
      Begin VB.Label lblNO 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "单据号"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   9540
         TabIndex        =   24
         Top             =   720
         Width           =   720
      End
   End
   Begin VB.Frame fraUnit 
      Height          =   1065
      Left            =   8520
      TabIndex        =   22
      Top             =   855
      Width           =   3375
      Begin VB.ComboBox cbo开单科室 
         Height          =   360
         Left            =   1095
         Style           =   2  'Dropdown List
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   405
         Width           =   2175
      End
      Begin VB.Label lbl开单科室 
         AutoSize        =   -1  'True
         Caption         =   "开单科室"
         Height          =   240
         Left            =   120
         TabIndex        =   26
         Top             =   465
         Width           =   960
      End
   End
   Begin VB.Frame fraInfo 
      Height          =   1065
      Left            =   30
      TabIndex        =   21
      Top             =   855
      Width           =   8490
      Begin VB.TextBox txt费别 
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   675
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "热键：F11"
         Top             =   615
         Width           =   1590
      End
      Begin VB.TextBox txt付款方式 
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   2970
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "热键：F11"
         Top             =   615
         Width           =   1590
      End
      Begin VB.TextBox txt性别 
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   2970
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "热键：F11"
         Top             =   210
         Width           =   1590
      End
      Begin VB.TextBox txt担保额 
         Height          =   360
         Left            =   7275
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   615
         Width           =   1095
      End
      Begin VB.TextBox txt担保人 
         Height          =   360
         Left            =   5490
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   615
         Width           =   870
      End
      Begin VB.TextBox txt床号 
         Height          =   360
         Left            =   7275
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   210
         Width           =   1095
      End
      Begin VB.TextBox txt姓名 
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   675
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   210
         Width           =   1590
      End
      Begin VB.TextBox txt年龄 
         Height          =   360
         Left            =   5490
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   210
         Width           =   870
      End
      Begin VB.Label lbl担保额 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "担保额"
         Height          =   240
         Left            =   6510
         TabIndex        =   44
         Top             =   675
         Width           =   720
      End
      Begin VB.Label lbl担保人 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "担保人"
         Height          =   240
         Left            =   4740
         TabIndex        =   43
         Top             =   675
         Width           =   720
      End
      Begin VB.Label lbl付款方式 
         Caption         =   "付款 方式"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2445
         TabIndex        =   42
         Top             =   585
         Width           =   420
      End
      Begin VB.Label lbl床号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床号"
         Height          =   240
         Left            =   6750
         TabIndex        =   33
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lbl姓名 
         AutoSize        =   -1  'True
         Caption         =   "病人"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   165
         TabIndex        =   31
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lbl性别 
         AutoSize        =   -1  'True
         Caption         =   "性别"
         Height          =   240
         Left            =   2415
         TabIndex        =   30
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lbl年龄 
         AutoSize        =   -1  'True
         Caption         =   "年龄"
         Height          =   240
         Left            =   4980
         TabIndex        =   29
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lbl费别 
         AutoSize        =   -1  'True
         Caption         =   "费别"
         Height          =   240
         Left            =   150
         TabIndex        =   28
         Top             =   675
         Width           =   480
      End
   End
   Begin ZL9BillEdit.BillEdit Bill 
      Height          =   3195
      Left            =   15
      TabIndex        =   10
      Top             =   1920
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   5636
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
      RowHeight0      =   360
      RowHeightMin    =   360
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "合计:"
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
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   945
   End
End
Attribute VB_Name = "frmTechnicExpense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
'入口参数
'————————————————————————————————————————————————————————————————————
Public mlng医嘱ID As Long '新增费用时用
Public mlng发送号 As Long '新增费用时用
Public mlng病人ID As Long '确定要计费的病人ID
Public mlng主页ID As Long '确定要计费的主页ID

Public mint病人来源 As Integer '1-门诊病人,2-住院病人
Public mint记录性质 As Integer '1-收费(划价),2-记帐(门/住)

Public mbln费用登记 As Boolean '仅登记,不计实收金额
Public mlng开单科室ID As Long '为当前主界面医技科室
Public mlng病人科室id As Long '主要是用于确定门诊病人的科室ID

Public mlng开嘱科室ID As Long
Public mstr开嘱医生 As String

Public mbytInState As Byte '0-执行,1-查阅,2-调整(不支持),3-删费
Public mstrInNO As String '所操作的单据号(执行时为修改)

Public mstrTime As String '操作单据内容的登记时间
Public mblnDelete As Boolean '是否处理退费单据(查阅)

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
    行 = 0
    类别 = 1
    项目 = 2
    规格 = 3
    单位 = 4
    付数 = 5
    数次 = 6
    单价 = 7
    应收金额 = 8
    实收金额 = 9
    执行科室 = 10
    标志 = 11
    类型 = 12
End Enum

Public mstrPrivs As String
'医技工作站本地费用参数
'————————————————————————————————————————————————————————————————————
Private mstrLike As String '输入匹配方式
Private mblnPay As Boolean '中药是否输入付数
Private mblnTime As Boolean '变价是否输入付数
Private mbln其它药房 As Boolean '是否显示其它药房库存
Private mbln其它药库 As Boolean '是否显示其它药库库存
Private mstr收费类别 As String '可输入的收费类别
Private mbln药房单位 As Boolean '是否按照门诊单位或住院单位显示药品
Private mstr药房单位 As String '根据病人来源决定如"门诊单位"或"住院单位"
Private mstr药房包装 As String '根据病人来源决定如"门诊包装"或"住院包装"
Private mlngPreRow As Long '记录当前行,当仅改变列时
Private mlng西药房 As Long, mlng成药房 As Long, mlng中药房 As Long
'————————————————————————————————————————————————————————————————————
'数据对象
Private mrsInfo As New ADODB.Recordset '病人信息
Private mrsMedAudit As ADODB.Recordset  '病人已审批的费用项目
Private mrsUnit As ADODB.Recordset '可选择的执行科室
Private mrsClass As ADODB.Recordset '根据参数读取的当前可用的收费类别
Private mrsWork As New ADODB.Recordset '当天上班的药房
Private mblnWork As Boolean '当前是否有正在上班的药房
Private mlng药品类别ID As Long '当前单据操作的药品入出类别ID
Private mlng卫材类别ID As Long '当前单据操作的卫材入出类别ID
'程序对象
Private mobjBill As ExpenseBill '费用单据对象
Private mobjBillDetail As BillDetail '单据的收费细目对象
Private mobjBillIncome As BillInCome '收费细目的收入项目对象
Private mobjDetail As Detail '单独的收费细目对象
Private mcolDetails As Details '单独的收费细目集合
Private mcolMoneys As BillInComes '收入项目汇总集合

'程序变量
Private mbytWarn As Byte '记帐报警的返回值
Private mintWarn As Integer '记帐报警提示的继续选项
Private mstrWarn As String '已经报过警并选择继续的类别
Private mrsWarn As New ADODB.Recordset '病区报警线
Private mcolStock1 As Collection '存放各个药品库房的出库检查方式
Private mcolStock2 As Collection '存放各个卫材库的出库检查方式

Private mcurModiMoney As Currency '修改单据时原单据的金额
Private mblnDrop As Boolean '在KeyDown中判断cbo开单人当前是否弹出
Private mblnNewRow As Boolean
Private mblnOne As Boolean '是否只有一个可用收费类别
Private marrColData() As Integer '当前单据编辑属性映象
Private mdblItemNum As Double '数据库中当前输入费目的数次
Private mblnSelect As Boolean '用于控制收费细目对象是否来自于列表选择或选择器
Private marrDr() As String '记录医生的"ID|科室ID|编号|姓名|简码"
Private mblnEnterCell As Boolean '控制是否执行Entercell事件

Private Const STR_HEAD = "行,450,4;类别,750,1;项目,2175,1;规格,1105,1;单位,520,4;付数,520,1;数次,570,1;单价,1055,7;" & "应收金额,1030,7;实收金额,1080,7;执行科室,1255,1;标志,520,4;类型,520,1"

Private Sub Bill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Dim i As Long, bytsubs As Byte
    Dim bln从项汇总折扣 As Boolean
    Dim lngMainRow As Long
    
    If mbytInState <> 0 Then Cancel = True: Exit Sub
    
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
        
        '重新计算并刷新
        If bln从项汇总折扣 Then
            If ItemHaveSub(lngMainRow) Then
                Call Calc重算主项实收(lngMainRow)
            Else
                Call CalcMoney(lngMainRow, False) '只有一个主项了,从项全部被删除时,当成普通独立项计算
            End If
        End If
        
        Call ShowDetails
        Call ShowMoney
                
        Bill.TxtVisible = False
        Bill.CmdVisible = False
        Bill.CboVisible = False
        
        Cancel = True '不用控件来处理删除
        
        mlngPreRow = 0  '表示行改变了
        Call Bill_EnterCell(Bill.Row, Bill.Col)
        
    ElseIf Row = 1 Then
        For i = 1 To Bill.Cols - 1
            Bill.TextMatrix(Row, i) = ""
        Next
        Cancel = True
    End If
    Call SetColNum(Row)
End Sub

Private Sub bill_cboClick(ListIndex As Long)
    Dim dblStock As Double, i As Long
    
    '药品库存检查
    If ListIndex <> -1 And Bill.TextMatrix(0, Bill.Col) = "执行科室" Then
        If mobjBill.Details.Count >= Bill.Row Then
            With mobjBill.Details(Bill.Row)
                If .执行部门ID <> Bill.ItemData(Bill.ListIndex) Then
                    .执行部门ID = Bill.ItemData(Bill.ListIndex)
                    Bill.TextMatrix(Bill.Row, Bill.Col) = Bill.CboText
                    
                    If InStr(",5,6,7,", .收费类别) > 0 Then
                        '取库存
                        dblStock = GetStock(.收费细目ID, .执行部门ID)
                        If mbln药房单位 Then
                            dblStock = dblStock / .Detail.药房包装
                        End If
                        .Detail.库存 = dblStock  '记录当前行药品库存
                        sta.Panels(2) = "[" & .Detail.名称 & "]可用库存量:" & dblStock
                        
                        '药房改变,实价药品重新计算价格
                        If .Detail.变价 Then
                            Call CalcMoneys(Bill.Row)
                            Call ShowDetails(Bill.Row)
                            Call ShowMoney
                        End If
                    ElseIf .收费类别 = "4" And .Detail.跟踪在用 Then
                        '取库存
                        dblStock = GetStock(.收费细目ID, .执行部门ID)
                        .Detail.库存 = dblStock
                        sta.Panels(2) = "[" & .Detail.名称 & "]可用库存量:" & dblStock
                        
                        '发料部门改变,时价卫材重新计算价格
                        If .Detail.变价 Then
                            Call CalcMoneys(Bill.Row)
                            Call ShowDetails(Bill.Row)
                            Call ShowMoney
                        End If
                    ElseIf InStr(",4,5,6,7,", .收费类别) = 0 Then
                        If ItemHaveSub(Bill.Row) Then Call SetSubDept(Bill.Row) '如果存在从项,则改变非药品行的执行科室
                    End If
                End If
            End With
        End If
    End If
End Sub

Private Sub bill_CellCheck(Row As Long, Col As Long)
'说明：可以全部为主要手术,但不能全部为附加手术
    Dim i As Long, strCheck As String, bytTime As Byte
    
    If Bill.TextMatrix(Row, 2) = "" Then Bill.TextMatrix(Row, Col) = "": Exit Sub
    If mbytInState = 3 Then Exit Sub
    
    '新增的未处理行无效
    If mobjBill.Details.Count < Row Then
        Bill.TextMatrix(Row, Col) = "": Exit Sub
    End If
    
    strCheck = Bill.TextMatrix(Row, Col)
    
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).收费类别 = "F" And mobjBill.Details(i).附加标志 = 0 And i <> Row Then bytTime = bytTime + 1
    Next
    If bytTime > 0 Then
        mobjBill.Details(Row).附加标志 = IIF(strCheck = "", 0, 1)
        Call CalcMoneys(Row)
        Call ShowDetails(Row)
        Call ShowMoney
    ElseIf strCheck <> "" Then
        Bill.TextMatrix(Row, Col) = ""
        MsgBox "单据中必然有一个手术不是附加手术！", vbInformation, gstrSysName
    End If
End Sub

Private Function SelectIsNurse() As Boolean
'功能：判断当前开单人是否护士
    Dim str性质 As String
    
    If cbo开单人.ListIndex <> -1 Then
        If cbo开单人.ItemData(cbo开单人.ListIndex) = 0 Then Exit Function
        
        If cbo开单人.ListIndex <= UBound(marrDr) Then
            If UBound(Split(marrDr(cbo开单人.ListIndex), "|")) >= 6 Then
                str性质 = Split(marrDr(cbo开单人.ListIndex), "|")(6)
                SelectIsNurse = str性质 = "护士"
            End If
        End If
    End If
End Function

Private Sub bill_CommandClick()
    Dim lng项目ID As Long, blnCancel As Boolean
    Dim str类别 As String, str特准项目 As String
    
    If gbln收费类别 Then
        If Bill.RowData(Bill.Row) <> 0 Then
            str类别 = "'" & Chr(Bill.RowData(Bill.Row)) & "'"
        Else
            str类别 = IIF(SelectIsNurse, "'E','M','4'", mstr收费类别)
        End If
    Else
        str类别 = IIF(SelectIsNurse, "'E','M','4'", mstr收费类别)
    End If
    If Not IsNull(mrsInfo!险类) Then
        str特准项目 = Get保险特准项目(mrsInfo!病人ID, "A.ID")
    End If
    
    lng项目ID = frmItemSelect.ShowSelect(Me, mstrPrivs, mint病人来源, True, str类别, , , str特准项目)
    If lng项目ID <> 0 Then
        Bill.Text = lng项目ID
        mblnSelect = True
        Call bill_KeyDown(13, 0, blnCancel)
        Bill.SetFocus
        If Not blnCancel Then
            Bill.Text = "": Bill.TxtVisible = False
            Call zlCommFun.PressKey(13)
        End If
    Else
        mblnSelect = False
    End If
End Sub

Private Sub bill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
'功能：处理单据输入
    Dim dblStock As Double, strScope As String, i As Long
    Dim dblPreTime As Double, dblPreMoney As Double
    Dim blnSkip As Boolean, curTotal As Currency, bln医保 As Boolean
    Dim blnStock As Boolean, lngDoUnit As Long, str摘要 As String
    Dim lng项目ID As Long, str特准项目 As String, str类别 As String
    Dim blnInput As Boolean, cur余额 As Currency, lng病人科室ID As Long
    Dim colStock As Collection
    On Error GoTo errH
    
    If KeyCode = 13 And Bill.Active Then
        If mbytInState = 2 Then
            If Bill.Col = Bill.Cols - 1 And Bill.Row = Bill.Rows - 1 Then
                Cancel = True: Exit Sub
            ElseIf Bill.TextMatrix(0, Bill.Col) <> "执行科室" Then
                Exit Sub
            End If
        End If
        If Bill.ColData(Bill.Col) = 0 Then Exit Sub
        
        '是否医保病人
        bln医保 = Val(txt付款方式.Tag) = 1 Or Not IsNull(mrsInfo!险类)
        
        Select Case Bill.TextMatrix(0, Bill.Col)
            Case "类别"
                If Bill.ListIndex <> -1 Then '如果不输入类别则不会进入类别列
                    If Bill.RowData(Bill.Row) <> Bill.ItemData(Bill.ListIndex) Then
                        '一旦改更收费类别,则清除(如有)原有该项目内容
                        For i = 2 To Bill.Cols - 1
                            Bill.TextMatrix(Bill.Row, i) = ""
                        Next
                        If mobjBill.Details.Count >= Bill.Row Then
                            Set mobjBill.Details(Bill.Row).Detail = New Detail
                            Set mobjBill.Details(Bill.Row).InComes = New BillInComes
                            With mobjBill.Details(Bill.Row)
                                .收费细目ID = 0: .收费类别 = ""
                            End With
                            Call CalcMoneys
                            Call ShowMoney
                        End If
                    End If
                    Bill.RowData(Bill.Row) = Bill.ItemData(Bill.ListIndex) '暂时用RowData记录所选择的收费类别
                End If
            Case "项目"
                '此项目确定,该收费细目对应的程序对象才生成,同时这里处理收费从属项目
                If Bill.Text <> "" Then
                    '如果在已输入的项目上按回车,或选择器选择
                    If mobjBill.Details.Count >= Bill.Row Then
                        '通过按钮选择是返回的ID,而输入则是文本,如果是一样的,则不改变
                        If Bill.TextMatrix(Bill.Row, 2) = Bill.Text Then
                            Bill.TxtVisible = False
                            Bill.CmdVisible = False
                            Exit Sub
                        End If
                    End If
                
                    blnInput = True
                    If mblnSelect Then
                        mblnSelect = False '立即清除该标志
                        Set mobjDetail = GetInputDetail(Val(Bill.Text))
                    Else
                        If gbln收费类别 Then
                            If Bill.RowData(Bill.Row) = 0 Then
                                sta.Panels(2) = "没有确定费用类别,请先输入类别！"
                                Bill.TxtSetFocus: Cancel = True: Exit Sub
                            End If
                            str类别 = "'" & Chr(Bill.RowData(Bill.Row)) & "'"
                        Else
                            str类别 = IIF(SelectIsNurse, "'E','M','4'", mstr收费类别)
                        End If
                        If Not IsNull(mrsInfo!险类) Then
                            str特准项目 = Get保险特准项目(mrsInfo!病人ID, "A.ID")
                        End If
                        lng项目ID = frmItemSelect.ShowSelect(Me, mstrPrivs, mint病人来源, True, str类别, Bill.Text, Bill.TxtHwnd, str特准项目)
                        If lng项目ID <> 0 Then
                            Set mobjDetail = GetInputDetail(lng项目ID)
                        Else
                            Bill.Text = "": Bill.TxtVisible = False
                            Bill.SetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    sta.Panels(2) = ""
                    Bill.TxtVisible = False '(不加不行)
                    
                    '医保费用项目是否审批检查
                    If mint病人来源 = 2 And mint记录性质 = 2 And Not IsNull(mrsInfo!险类) Then
                        If mobjDetail.要求审批 And Not mrsMedAudit Is Nothing Then
                            mrsMedAudit.Filter = "项目ID=" & mobjDetail.ID
                            If mrsMedAudit.RecordCount = 0 Then
                                MsgBox "当前病人未被批准使用该项目！", vbInformation, gstrSysName
                                Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                    
                    '检查药品输入是否重复:分批及时价同一药房不允许重复(这里只提醒)
                    If InStr(",5,6,7,", mobjDetail.类别) > 0 _
                        Or (mobjDetail.类别 = "4" And mobjDetail.跟踪在用) Then
                        If PhysicExist(mobjDetail, Bill.Row) Then
                            Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    '检查处方职务
                    If InStr(",5,6,7,", mobjDetail.类别) > 0 Then
                        mobjDetail.处方职务 = Get处方职务(mobjDetail.ID)
                        '医保或公费病人
                        If InStr(",1,2,", txt付款方式.Tag) > 0 Then
                            If CheckDuty(mobjDetail, False) > 0 Then
                                Bill.TxtSetFocus: Cancel = True: Exit Sub
                            End If
                        End If
                        '所有病人
                        If CheckDuty(mobjDetail, True) > 0 Then
                            Bill.TxtSetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    
                    '病人科室ID
                    lng病人科室ID = mobjBill.科室ID
                    If lng病人科室ID = 0 And cbo开单科室.ListIndex <> -1 Then lng病人科室ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
                    
                    sta.Panels(2) = ""
                    lngDoUnit = Get收费执行科室ID(mlng病人ID, mlng主页ID, mobjDetail.类别, mobjDetail.ID, _
                        mobjDetail.执行科室, lng病人科室ID, Get开单科室ID, mint病人来源, Nvl(mrsInfo!病区ID, 0)) '卫材缺省与病人病区(开单科室)相同
                    
                    
                    '读取药品相关信息
                    If InStr(",5,6,7,", mobjDetail.类别) > 0 Then
                        '当前行药品库存
                        dblStock = GetStock(mobjDetail.ID, lngDoUnit, blnStock)
                        If mbln药房单位 Then
                            dblStock = dblStock / mobjDetail.药房包装
                        End If
                        mobjDetail.库存 = dblStock
                        sta.Panels(2) = "[" & mobjDetail.名称 & "]可用库存量:" & mobjDetail.库存

                        '处方限量
                        mobjDetail.处方限量 = Get处方限量(mobjDetail.ID)
                    ElseIf mobjDetail.类别 = "4" And mobjDetail.跟踪在用 Then
                        dblStock = GetStock(mobjDetail.ID, lngDoUnit)
                        mobjDetail.库存 = dblStock
                        sta.Panels(2).Text = "[" & mobjDetail.名称 & "]可用库存量:" & mobjDetail.库存
                    End If
                    
                    '保险支付项目对应检查
                    If Not IsNull(mrsInfo!险类) Then
                        If Not ItemExistInsure(mobjDetail.ID, mrsInfo!险类) Then
                            If gint医保对码 = 1 Then
                                If MsgBox("项目""" & mobjDetail.名称 & """没有设置对应的保险项目，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                    Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                                End If
                            ElseIf gint医保对码 = 2 Then
                                MsgBox "项目""" & mobjDetail.名称 & """没有设置对应的保险项目。", vbInformation, gstrSysName
                                Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                    '输入摘要(取已有的行以便修改)
                    If mobjBill.Details.Count >= Bill.Row Then
                        If mobjBill.Details(Bill.Row).Detail.ID = mobjDetail.ID Then
                            str摘要 = mobjBill.Details(Bill.Row).摘要
                        End If
                    End If
                    
                    '加入或修改该收费细目行
                    Call SetDetail(mobjDetail, Bill.Row, lngDoUnit)
                    Call CalcMoneys(Bill.Row)
                    
                    'Calcmoney中医保可能返回摘要
                    If mobjBill.Details(Bill.Row).摘要 <> "" Then str摘要 = mobjBill.Details(Bill.Row).摘要
                    
                    '记帐分类报警(在已经算出该行费用但未显示前)
                    If mint记录性质 = 2 And mrsWarn.State = 1 And mobjBill.Details.Count = Bill.Row Then
                        curTotal = GetBillTotal(mobjBill)
                        If curTotal > 0 Then
                            cur余额 = Val(txt实收.Tag)
                            If gbln报警包含划价费用 Then cur余额 = cur余额 - GetPriceMoneyTotal(mrsInfo!病人ID)
                            mbytWarn = BillingWarn(Me, mstrPrivs, mrsWarn, mrsInfo!姓名, cur余额, mrsInfo!当日额 - mcurModiMoney, curTotal, _
                                Nvl(mrsInfo!担保额, 0), mobjBill.Details(Bill.Row).收费类别, mobjBill.Details(Bill.Row).Detail.类别名称, mstrWarn, mintWarn, bln医保)
                            If mbytWarn = 2 Or mbytWarn = 3 Then
                                mobjBill.Details.Remove Bill.Row '删除刚刚想要加入的费用行
                                Bill.Text = "": Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                    Call ShowDetails(Bill.Row)
                    Call ShowMoney
                    
                    '费用类型检查
                    Call Check费用类型(Bill.Row)
                    
                    '输入摘要(根据新输入的行更改摘要)
                    If mobjBill.Details(Bill.Row).Detail.补充摘要 Then
                        If frmInputBox.InputBox(Me, "摘要", "请输入""" & mobjBill.Details(Bill.Row).Detail.名称 & """的摘要信息:", 200, 3, True, False, str摘要) Then
                            mobjBill.Details(Bill.Row).摘要 = str摘要
                        End If
                    ElseIf mint病人来源 = 2 And Not IsNull(mrsInfo!险类) Then
                        str摘要 = gclsInsure.GetItemInfo(mrsInfo!险类, mrsInfo!病人ID, mobjBill.Details(Bill.Row).收费细目ID, str摘要, 2)
                        mobjBill.Details(Bill.Row).摘要 = str摘要
                    End If
                    
                    Bill.Text = "": Bill.SetFocus
                End If
                
                If mobjBill.Details.Count >= Bill.Row Then
                    With mobjBill.Details(Bill.Row)
                        '下一列的性质确定
                        If .收费类别 = "7" And mblnPay Then Bill.ColData(5) = 4 '付数
                        If .收费类别 = "F" Then Bill.ColData(11) = -1 '附加标志
                        
                        '变价允许输入数次
                        If .Detail.变价 And InStr(",5,6,7,", .收费类别) = 0 _
                            And Not (.收费类别 = "4" And .Detail.跟踪在用) Then
                            Bill.ColData(6) = IIF(mblnTime, 4, 5) '数次
                            Bill.ColData(7) = 4 '单价
                        Else
                            Bill.ColData(6) = 4 '数次
                            Bill.ColData(7) = 5 '单价
                        End If
                        
                        '执行科室
                        mblnEnterCell = False: Bill.Col = BillCol.执行科室: mblnEnterCell = True
                        Call FillBillComboBox(Bill.Row, 10, Not blnInput) '直接回车时保持执行科室
                        mblnEnterCell = False: Bill.Col = BillCol.项目: mblnEnterCell = True
                        
                        blnSkip = Bill.ListCount = 1
                        If Not blnSkip And InStr(",5,6,7,", .收费类别) > 0 Then
                            '指定了固定药房时,不允许再选择
                            Select Case .收费类别
                                Case "5"
                                    blnSkip = mlng西药房 > 0 And .执行部门ID = mlng西药房
                                Case "6"
                                    blnSkip = mlng成药房 > 0 And .执行部门ID = mlng成药房
                                Case "7"
                                    blnSkip = mlng中药房 > 0 And .执行部门ID = mlng中药房
                            End Select
                        End If
                        If blnSkip Then
                            Bill.ColData(10) = 5: .Key = 1
                        Else
                            Bill.ColData(10) = 3: .Key = Bill.ListCount
                        End If
                        
                        '检查卫生材料的灭菌效期,在确定执行科室之后
                        If .收费类别 = "4" And .Detail.跟踪在用 Then
                            Call CheckValidity(.收费细目ID, .执行部门ID, .数次, False) '已确认输入,仅能提醒
                        End If
                                                
                         '从属项目处理,仅该行收费项目有从属项目及尚未取才取,药品无需判断,药品不能设置主从项
                        If Bill.TextMatrix(0, Bill.Col) = "项目" And InStr(",5,6,7,", .收费类别) = 0 Then
                            If (gbln从项汇总折扣 And mobjBill.Details(Bill.Row).从属父号 = 0) Or Not gbln从项汇总折扣 Then  '(如果有级联,只取一级)
                                If ShouldDO(Bill.Row) Then
                                   Call SetSubItem
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
                                Bill.Col = 5: Exit For
                            End If
                        Next
                    End If
                End If
            Case "付数"
                If mobjBill.Details.Count >= Bill.Row And Bill.Text <> "" Then
                    '数字合法性
                    If Not IsNumeric(Bill.Text) Then
                        MsgBox "非法数值！", vbInformation, gstrSysName
                        Bill.Text = mobjBill.Details(Bill.Row).付数: Cancel = True: Exit Sub
                    End If
                    If Val(Bill.Text) <= 0 Or Val(Bill.Text) <> Int(Val(Bill.Text)) Then
                        MsgBox "付数应该为正的整数！", vbInformation, gstrSysName
                        Bill.Text = mobjBill.Details(Bill.Row).付数: Cancel = True: Exit Sub
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
                                    MsgBox "第 " & i & " 行药品""" & mobjBill.Details(Bill.Row).Detail.名称 & """为分批或时价药品,修改付数后可用库存不足！", vbInformation, gstrSysName
                                    Bill.Text = mobjBill.Details(Bill.Row).付数: Cancel = True: Exit Sub
                                End If
                            End If
                        Next
                        
                        '计算并刷新该行
                        mobjBill.Details(Bill.Row).付数 = Bill.Text
                        Call CalcMoneys(Bill.Row)
                        Call ShowDetails(Bill.Row)
                                               
                         '处理其它中药付数,如果是独立项,则修改其它非从项的,如果是从项,则修改同一主项的从项的.因为限定为中草药,不可能有主项
                        For i = 1 To mobjBill.Details.Count
                            If i <> Bill.Row And mobjBill.Details(i).收费类别 = "7" And mobjBill.Details(i).从属父号 = mobjBill.Details(Bill.Row).从属父号 Then
                                If mobjBill.Details(i).从属父号 = 0 Or (mobjBill.Details(i).从属父号 <> 0 And mobjBill.Details(i).Detail.固有从属 = 0) Then     '1和2固定和按比例的不改
                                    mobjBill.Details(i).付数 = Bill.Text
                                    Call CalcMoneys(i)
                                    Call ShowDetails(i)
                                End If
                            End If
                        Next
                                                
                        Call ShowMoney
                    Else
                        sta.Panels(2) = "从属项目的付数不能更改！"
                        Bill.Text = mobjBill.Details(Bill.Row).付数: Beep '恢复原有付数值
                    End If
                End If
            Case "数次"
                If mobjBill.Details.Count >= Bill.Row And Bill.Text <> "" Then
                    '数字合法性
                    If Not IsNumeric(Bill.Text) Then
                        MsgBox "非法数值！", vbInformation, gstrSysName
                        Bill.Text = mobjBill.Details(Bill.Row).数次: Cancel = True: Exit Sub
                    End If
                    If Val(Bill.Text) = 0 Then
                        If MsgBox("数量输入为零，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Bill.Text = mobjBill.Details(Bill.Row).数次: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    Bill.Text = FormatEx(Bill.Text, 5)
                    
                    '负数合法性检查
                    If CSng(Bill.Text) * mobjBill.Details(Bill.Row).付数 < 0 Then
                        '权限
                        If Not ((InStr(",5,6,7,", mobjBill.Details(Bill.Row).收费类别) > 0 And InStr(mstrPrivs, "药品负数费用") > 0) _
                             Or (InStr(",5,6,7,", mobjBill.Details(Bill.Row).收费类别) = 0 And InStr(mstrPrivs, "诊疗负数费用") > 0)) Then
                            MsgBox "你没有权限输入负数！", vbInformation, gstrSysName
                            Bill.Text = mobjBill.Details(Bill.Row).数次: Cancel = True: Exit Sub
                        Else
                            If mobjBill.Details(Bill.Row).Detail.分批 Then
                                MsgBox "分批药品不允许输入负数。", vbInformation, gstrSysName
                                Bill.Text = mobjBill.Details(Bill.Row).数次: Cancel = True: Exit Sub
                            End If
                            If mrsInfo.State = 1 And mint记录性质 = 2 Then
                                If Not IsNull(mrsInfo!险类) Then
                                    If Not gclsInsure.GetCapability(support负数记帐, , mrsInfo!险类) Then
                                        MsgBox "本地医保不支持对医保病人进行负数记帐！", vbInformation, gstrSysName
                                        Bill.Text = mobjBill.Details(Bill.Row).数次: Cancel = True: Exit Sub
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                    '药品库存检查
                    With mobjBill.Details(Bill.Row)
                        If (.收费类别 = "4" And .Detail.跟踪在用) Or InStr(",5,6,7,", .收费类别) > 0 Then
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
                                Set colStock = IIF(.收费类别 = "4", mcolStock2, mcolStock1)
                                If colStock("_" & .执行部门ID) <> 0 And Bill.ColData(10) = 5 Then
                                    '其它药品正常检查
                                    If .付数 * CSng(Bill.Text) > .Detail.库存 Then
                                        If colStock("_" & .执行部门ID) = 1 Then
                                            If MsgBox("""" & .Detail.名称 & """的当前库存不足当前需求量,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                                Bill.Text = .数次: Cancel = True: Exit Sub
                                            End If
                                        ElseIf colStock("_" & .执行部门ID) = 2 Then
                                            MsgBox """" & .Detail.名称 & """的当前库存不足当前输入付数数量！", vbInformation, gstrSysName
                                            Bill.Text = .数次: Cancel = True: Exit Sub
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End With
                    
                    dblPreTime = mobjBill.Details(Bill.Row).数次
                    mobjBill.Details(Bill.Row).数次 = Bill.Text
                    
                    '处方限量检查
                    If Not CheckLimit(mobjBill, Bill.Row, mbln药房单位) Then
                        mobjBill.Details(Bill.Row).数次 = dblPreTime: Bill.Text = dblPreTime
                        Cancel = True: Exit Sub
                    End If
                    If mobjBill.Details(Bill.Row).Detail.录入限量 > 0 And mobjBill.Details(Bill.Row).数次 > mobjBill.Details(Bill.Row).Detail.录入限量 Then
                        If MsgBox("输入的数次超过了录入限量" & mobjBill.Details(Bill.Row).Detail.录入限量 & ",是否继续?", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo Then
                            mobjBill.Details(Bill.Row).数次 = dblPreTime: Bill.Text = dblPreTime
                            Cancel = True: Exit Sub
                        End If
                    End If
                    
                    '固有从属不能更改数次(主项目数次改变,固有从属的数次也变)
                    If mobjBill.Details(Bill.Row).从属父号 <> 0 And mobjBill.Details(Bill.Row).Detail.固有从属 <> 0 Then
                        sta.Panels(2) = "该项目是固有从属项目,其数次不能够更改。"
                        mobjBill.Details(Bill.Row).数次 = dblPreTime: Bill.Text = dblPreTime
                        Exit Sub
                    End If
                
                    Call CalcMoneys(Bill.Row)
                    
                    '数据溢出检查(在已经算出该行费用但未显示前)
                    If MoneyOverFlow(mobjBill) Then
                        MsgBox "输入数量导致单据金额过大，请作适当调整。", vbInformation, gstrSysName
                        mobjBill.Details(Bill.Row).数次 = dblPreTime
                        Bill.Text = ""
                        Call CalcMoneys(Bill.Row)
                        Cancel = True: Bill.TxtVisible = False: Exit Sub
                    End If
                    
                    '记帐分类报警(在已经算出该行费用但未显示前)
                    If mint记录性质 = 2 And mrsWarn.State = 1 Then
                        curTotal = GetBillTotal(mobjBill)
                        If curTotal > 0 Then
                            cur余额 = Val(txt实收.Tag)
                            If gbln报警包含划价费用 Then cur余额 = cur余额 - GetPriceMoneyTotal(mrsInfo!病人ID)
                            mbytWarn = BillingWarn(Me, mstrPrivs, mrsWarn, mrsInfo!姓名, cur余额, mrsInfo!当日额 - mcurModiMoney, curTotal, _
                                Nvl(mrsInfo!担保额, 0), mobjBill.Details(Bill.Row).收费类别, mobjBill.Details(Bill.Row).Detail.类别名称, mstrWarn, mintWarn, bln医保)
                            If mbytWarn = 2 Or mbytWarn = 3 Then
                                mobjBill.Details(Bill.Row).数次 = dblPreTime
                                Bill.Text = ""
                                Call CalcMoneys(Bill.Row)
                                Cancel = True: Bill.TxtVisible = False: Exit Sub
                            End If
                        End If
                    End If
                    
                    Call ShowDetails(Bill.Row)
                    
                    '更改其固有从属的数次
                    For i = Bill.Row + 1 To mobjBill.Details.Count
                        If mobjBill.Details(i).从属父号 = Bill.Row And mobjBill.Details(i).Detail.固有从属 = 2 Then
                            mobjBill.Details(i).数次 = Bill.Text * mobjBill.Details(i).Detail.从项数次
                            Call CalcMoneys(i)
                            Call ShowDetails(i)
                        End If
                    Next
                    Call ShowMoney

                 ElseIf mobjBill.Details.Count >= Bill.Row Then
                    If Val(Bill.TextMatrix(Bill.Row, Bill.Col)) = 0 Then
                        If MsgBox("数量输入为零，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Cancel = True: Exit Sub
                        End If
                    End If
                    If Bill.ColData(BillCol.执行科室) = BillColType.UnFocus Then
                        If ItemHaveSub(Bill.Row) Then
                            KeyCode = 0
                            Call LocateMainItemNextRow(Bill.Row)
                        End If
                    End If
               End If
            Case "单价"
                If mobjBill.Details.Count >= Bill.Row And Bill.Text <> "" Then
                    '数字合法性
                    If Not IsNumeric(Bill.Text) Then
                        MsgBox "非法数值！", vbInformation, gstrSysName
                        Bill.Text = "": Cancel = True: Bill.TxtVisible = False: Exit Sub
                    End If
                    If Val(Bill.Text) < 0 Then
                        MsgBox "项目价格不应该为负数，要删除费用，请输入负的数量来实现！", vbInformation, gstrSysName
                        Bill.Text = "": Cancel = True: Bill.TxtVisible = False: Exit Sub
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
                        If mint记录性质 = 2 And mrsWarn.State = 1 Then
                            curTotal = GetBillTotal(mobjBill)
                            If curTotal > 0 Then
                                cur余额 = Val(txt实收.Tag)
                                If gbln报警包含划价费用 Then cur余额 = cur余额 - GetPriceMoneyTotal(mrsInfo!病人ID)
                                mbytWarn = BillingWarn(Me, mstrPrivs, mrsWarn, mrsInfo!姓名, cur余额, mrsInfo!当日额 - mcurModiMoney, curTotal, _
                                    Nvl(mrsInfo!担保额, 0), mobjBill.Details(Bill.Row).收费类别, mobjBill.Details(Bill.Row).Detail.类别名称, mstrWarn, mintWarn, bln医保)
                                If mbytWarn = 2 Or mbytWarn = 3 Then
                                    mobjBill.Details(Bill.Row).InComes(1).标准单价 = dblPreMoney
                                    Bill.Text = ""
                                    Call CalcMoneys(Bill.Row)
                                    Cancel = True: Bill.TxtVisible = False: Exit Sub
                                End If
                            End If
                        End If
                        
                        Call ShowDetails(Bill.Row)
                        Call ShowMoney
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
                                If ItemHaveSub(Bill.Row) Then Call SetSubDept(Bill.Row) '如果存在从项,则改变非药品行的执行科室
                            End If
                    
                            '药品库存检查:动态药房,分批或时价药品也要检查了
                            If (.收费类别 = "4" And .Detail.跟踪在用) Or InStr(",5,6,7,", .收费类别) > 0 Then
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
                                    Set colStock = IIF(.收费类别 = "4", mcolStock2, mcolStock1)
                                    If colStock("_" & .执行部门ID) <> 0 Then
                                        If .付数 * .数次 > .Detail.库存 Then
                                            If colStock("_" & .执行部门ID) = 1 Then
                                                If MsgBox("[" & .Detail.名称 & "]的当前库存不足当前需求量,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                                    Cancel = True
                                                End If
                                            ElseIf colStock("_" & .执行部门ID) = 2 Then
                                                MsgBox "[" & .Detail.名称 & "]的当前库存不足当前输入付数数量！", vbInformation, gstrSysName
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
                        
                            If ItemHaveSub(Bill.Row) Then
                                KeyCode = 0
                                Call LocateMainItemNextRow(Bill.Row)
                            End If
                    End With
                End If
        End Select
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
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
        mblnNewRow = True
        Call bill_AfterAddRow(Bill.Rows - 1)
        mblnNewRow = False
        Bill.Row = Bill.Rows - 1
        Bill.MsfObj.TopRow = Bill.Row
        Bill.Col = 1
    Else
        Bill.Row = Bill.Rows - 1
        Bill.MsfObj.TopRow = Bill.Row
        Bill.Col = 1
    End If
End Sub

Private Sub SetSubItem()
'功能:输入收费项目后,加载当前收费项目的从属项目到费用集对象,并显示在单据控件中
'参数:
'调用者:Bill_KeyDown中输入项目后
Dim i As Integer, j As Integer, lngMainRow As Long
Dim lngDoUnit As Long               '执行科室ID
Dim bln从项汇总折扣 As Boolean
Dim str摘要 As String

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
            mblnNewRow = True
            Call bill_AfterAddRow(Bill.Rows - 1)
            mblnNewRow = False
        End If
        Bill.TextMatrix(Bill.Rows - 1, 1) = "" '有必要加上
        
        'a.从属项目为非药品项目的执行科室
        lngDoUnit = 0
        If InStr(",4,5,6,7,", mcolDetails(i).类别) = 0 Then
             If mcolDetails(i).类别 = .收费类别 Then
                '1.从项收费类别与主项相同的,缺省与主项执行科室相同。
                lngDoUnit = .执行部门ID
             Else
                If mcolDetails(i).执行科室 = 0 Then
                    '2.从项设置为无明确科室的,缺省与主项执行科室相同。
                    lngDoUnit = .执行部门ID
                Else
                    '其它非药项目的执行科室
                    lngDoUnit = Get收费执行科室ID(mlng病人ID, mlng主页ID, mcolDetails(i).类别, _
                        mcolDetails(i).ID, mcolDetails(i).执行科室, lngDoUnit, Get开单科室ID, mint病人来源)
                End If
             End If
        End If
        
        'b.从属项目为药品,卫材的执行科室(如果主项的执行科室为空,也会执行到这里)
        If lngDoUnit = 0 Then
            lngDoUnit = mobjBill.科室ID
            If lngDoUnit = 0 And cbo开单科室.ListIndex <> -1 Then
                lngDoUnit = cbo开单科室.ItemData(cbo开单科室.ListIndex)
            End If
            lngDoUnit = Get收费执行科室ID(mlng病人ID, mlng主页ID, mcolDetails(i).类别, mcolDetails(i).ID, _
                mcolDetails(i).执行科室, lngDoUnit, Get开单科室ID, mint病人来源, .执行部门ID) '卫材从项缺省与主项执行科室相同
        End If
            
        '保险支付项目对应检查
        If mrsInfo.State = 1 Then
            If Not IsNull(mrsInfo!险类) Then
                If Not ItemExistInsure(mcolDetails(i).ID, mrsInfo!险类) Then
                    If gint医保对码 = 1 Then
                        If MsgBox("项目""" & mcolDetails(i).名称 & """没有设置对应的保险项目，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Exit Sub
                        End If
                    ElseIf gint医保对码 = 2 Then
                        MsgBox "项目""" & mcolDetails(i).名称 & """没有设置对应的保险项目。", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
            End If
        End If
        
        Call SetDetail(mcolDetails(i), Bill.Rows - 1, lngDoUnit, Bill.Row)
        
        Call CalcMoney(Bill.Rows - 1, bln从项汇总折扣)
        Call ShowDetails(Bill.Rows - 1)
        
        If mrsInfo.State = 1 Then
             If Not IsNull(mrsInfo!险类) Then
                'CalcMoney中先调用GetuItemInsure可能返回摘要
                str摘要 = mobjBill.Details(Bill.Rows - 1).摘要
                
                str摘要 = gclsInsure.GetItemInfo(mrsInfo!险类, mrsInfo!病人ID, mcolDetails(i).ID, str摘要, 1)
                mobjBill.Details(Bill.Rows - 1).摘要 = str摘要
             End If
        End If
    Next
    
    If bln从项汇总折扣 And Not mbln费用登记 Then
        Call CalcMoney(lngMainRow, bln从项汇总折扣) '先重算主项的应收与实收,因为在没有加入从项前可能是按单独打折算的.
        
        Call Calc重算主项实收(lngMainRow)
    End If
    
    Call ShowMoney
End With

End Sub

Private Function Get开单科室ID() As Long
    If cbo开单科室.ListIndex <> -1 Then
        Get开单科室ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
    Else
        Get开单科室ID = UserInfo.部门ID
    End If
End Function

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
   
    cur打折后实收 = CCur(Format(ActualMoney(.费别, .Details(lngMainRow).InComes(1).收入项目ID, cur打折前应收合计), gstrDec))
    
    cur打折后实收 = cur打折后实收 - cur打折前应收合计 + .Details(lngMainRow).InComes(1).应收金额
    
    .Details(lngMainRow).InComes(1).实收金额 = Format(cur打折后实收, gstrDec)
    .Details(lngMainRow).InComes(1).Key = "_" & Format(cur打折后实收, gstrDec)
    
    
    Call ShowDetails(lngMainRow)
End With
End Sub

Private Sub SetSubDept(ByVal lngRow As Long)
Dim i As Long, j As Long
    With mobjBill
        Set mcolDetails = GetSubDetails(.Details(lngRow).收费细目ID) '必须现取
        
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
                                If cbo开单科室.ListIndex <> -1 Then
                                    .Details(i).执行部门ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
                                End If
                                .Details(i).执行部门ID = Get收费执行科室ID(mlng病人ID, mlng主页ID, mcolDetails(j).类别, _
                                    mcolDetails(j).ID, mcolDetails(j).执行科室, .Details(i).执行部门ID, Get开单科室ID, mint病人来源)
                            End If
                        End If
                    End If
                    
                    '显示从项执行科室
                    If .Details(i).执行部门ID <> 0 Then
                        If mbytInState = 0 Then
                            mrsUnit.Filter = "ID=" & .Details(i).执行部门ID
                            If mrsUnit.RecordCount <> 0 Then
                                Bill.TextMatrix(i, BillCol.执行科室) = mrsUnit!编码 & "-" & mrsUnit!名称
                            Else
                                Bill.TextMatrix(i, BillCol.执行科室) = Get部门名称(.Details(i).执行部门ID)
                            End If
                        Else
                            '浏览单据只(能)显示名称
                            Bill.TextMatrix(i, BillCol.执行科室) = Get部门名称(.Details(i).执行部门ID)
                        End If
                    Else
                        Bill.TextMatrix(i, BillCol.执行科室) = ""
                    End If
                End If
            End If
        Next
    End With
End Sub

Private Function ItemHaveSub(ByVal lngRow As Long) As Boolean
'功能：判断当前行的项目是否具有从属项目
    Dim i As Long
    
    If mobjBill.Details.Count >= lngRow Then
        For i = lngRow + 1 To mobjBill.Details.Count
            If mobjBill.Details(i).从属父号 = lngRow Then
                ItemHaveSub = True: Exit Function
            End If
        Next
    End If
End Function

Private Sub Bill_EnterCell(Row As Long, Col As Long)
    Dim strStock As String, i As Long
    
    If Not Bill.Active Then Exit Sub
    If Bill.ColData(Col) = BillColType.UnFocus Then Exit Sub
    If Not mblnEnterCell Then Exit Sub
    
    If mbytInState = 3 Then
        '针对列编辑性质设置颜色
        Bill.SetColColor 1, &HE7CFBA '不然要成白色
        Exit Sub
    End If
    
     '--------------------------------------------------------------------------
    '1.行改变的相关数据处理和设置
    If mobjBill.Details.Count >= Bill.Row And mlngPreRow <> Row Then
        With mobjBill.Details(Bill.Row)
            '显示库存
            If InStr(",5,6,7,", .收费类别) > 0 And .收费细目ID <> 0 Then
                If mbln其它药房 Or mbln其它药库 Then
                    strStock = GetStockInfo(.收费细目ID, mbln其它药房, mbln其它药库, mbln药房单位, mstr药房包装)
                    If strStock <> "" Then sta.Panels(2) = "第" & Bill.Row & "行库存:" & strStock
                End If
                If strStock = "" Then
                    '随时更新库存显示
                    .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                    If mbln药房单位 Then
                        .Detail.库存 = .Detail.库存 / .Detail.药房包装
                    End If
                    sta.Panels(2) = "[" & .Detail.名称 & "]可用库存:" & .Detail.库存
                End If
            ElseIf .收费类别 = "4" And .Detail.跟踪在用 And .收费细目ID <> 0 Then
                .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                sta.Panels(2) = "[" & .Detail.名称 & "]可用库存:" & .Detail.库存
            ElseIf .Detail.变价 And .InComes.Count > 0 And Bill.TextMatrix(0, Bill.Col) = "单价" Then
                sta.Panels(2) = "价格范围:" & FormatEx(.InComes(1).原价, 5) & "-" & FormatEx(.InComes(1).现价, 5)
            Else
                sta.Panels(2) = ""
            End If
            
            Bill.ColData(1) = IIF(gbln收费类别, BillColType.ComboBox, BillColType.UnFocus)
            Bill.ColData(2) = BillColType.CommandButton
            
             '如果是从属项目的主项目或从项,则不允许更改类别和项目
            If ItemHaveSub(Row) Or .从属父号 > 0 Then
                Bill.ColData(1) = BillColType.Text_UnModify
                Bill.ColData(2) = BillColType.Text_UnModify
            End If
            
            '如果是非调整状态
            If mbytInState <> 2 Then
                If .收费类别 = "7" And mblnPay Then
                    Bill.ColData(5) = 4
                Else
                    Bill.ColData(5) = 5
                End If
                
                '变价允许输入数次
                If .Detail.变价 And InStr(",5,6,7,", .收费类别) = 0 _
                    And Not (.收费类别 = "4" And .Detail.跟踪在用) Then
                    Bill.ColData(6) = IIF(mblnTime, 4, 5) '数次
                    Bill.ColData(7) = 4 '金额
                Else
                    Bill.ColData(6) = 4
                    Bill.ColData(7) = 5
                End If
                
                If .Key = "1" Then    '指定了固定药房时,不允许再选择执行科室
                    Bill.ColData(10) = BillColType.UnFocus
                Else
                    Bill.ColData(10) = BillColType.ComboBox
                End If
                
                If .收费类别 = "F" Then
                    Bill.ColData(11) = -1
                Else
                    Bill.ColData(11) = 5
                End If
                
                 '只允许一个类别
                If mblnOne Then Bill.ColData(1) = 5
            End If
        End With
    End If
   
    '如果点击未保存的行,则恢复列的性质
    If mobjBill.Details.Count < Bill.Row Then
        Bill.ColData(1) = IIF(gbln收费类别, BillColType.ComboBox, BillColType.UnFocus) '类别列,当主从项时会被改变
        Bill.ColData(2) = BillColType.CommandButton  '项目列,当主从项时会被改变
    End If
    
    
    '-----------------------------------------------------------------
    '2.列改变的相关数据处理和显示设置
    If Bill.ColData(Bill.Col) = BillColType.ComboBox Then
        Call FillBillComboBox(Bill.Row, Bill.Col, True) '进入该列
    End If
    
    If gbln收费类别 And Bill.TextMatrix(Row, 1) = "" And mblnOne Then
        mrsClass.Filter = "编码=" & mstr收费类别
        Bill.TextMatrix(Row, 1) = mrsClass!类别
        Bill.RowData(Row) = Asc(mrsClass!编码)
    End If
    
    Bill.TextLen = 0: Bill.TextMask = ""
    Select Case Bill.TextMatrix(0, Col)
        Case "类别" '不输入收费类别时不会进入类别列
            Call zlControl.CboSetWidth(Bill.CboHwnd, 1000)
            '类别如果为空,则自动默认为上一收费细目的类别
            If Bill.TextMatrix(Row, Col) = "" Then
                If mblnOne Then
                    mrsClass.Filter = "编码=" & mstr收费类别
                    Bill.TextMatrix(Row, Col) = mrsClass!类别
                    Bill.RowData(Row) = Asc(mrsClass!编码)
                ElseIf Row > 1 Then
                    Bill.ListIndex = -1
                    For i = 0 To Bill.ListCount - 1
                        If InStr(Bill.List(i), Bill.TextMatrix(Row - 1, Col)) > 0 Then Bill.ListIndex = i: Exit For
                    Next
                End If
            ElseIf Row >= 1 And Bill.TextMatrix(Row, Col) <> "" Then
                For i = 0 To Bill.ListCount - 1
                    If InStr(Bill.List(i), Bill.TextMatrix(Row, Col)) > 0 Then
                        Bill.ListIndex = i: Exit For
                    End If
                Next
                If Bill.ListIndex = -1 Then
                    Bill.ListIndex = SendMessage(Bill.CboHwnd, CB_FINDSTRING, -1, ByVal Bill.TextMatrix(Row - 1, Col))
                End If
            End If
        Case "执行科室"
            Call zlControl.CboSetWidth(Bill.CboHwnd, 2000)
        Case "付数"
            Bill.TextLen = 3
            Bill.TextMask = "0123456789" & Chr(8)
        Case "数次"
            Bill.TextLen = 8
            Bill.TextMask = "0123456789" & Chr(8)
            
            If mobjBill.Details.Count >= Bill.Row Then
                '可否输入小数
                If InStr(",5,6,7,", mobjBill.Details(Bill.Row).收费类别) > 0 Then
                    If InStr(mstrPrivs, "药品小数输入") > 0 Then
                        Bill.TextMask = "." & Bill.TextMask
                    End If
                Else
                    Bill.TextMask = "." & Bill.TextMask
                End If
                
                '可否输入负数
                If Not mobjBill.Details(Bill.Row).Detail.分批 Then
                    If InStr(",5,6,7,", mobjBill.Details(Bill.Row).收费类别) > 0 Then
                        If InStr(mstrPrivs, "药品负数费用") > 0 Then Bill.TextMask = "-" & Bill.TextMask
                    Else
                        If InStr(mstrPrivs, "诊疗负数费用") > 0 Then Bill.TextMask = "-" & Bill.TextMask
                    End If
                                    
                    If InStr(Bill.TextMask, "-") > 0 Then
                        If mrsInfo.State = 1 And mint记录性质 = 2 Then
                            If Not IsNull(mrsInfo!险类) Then
                                If Not gclsInsure.GetCapability(support负数记帐, , mrsInfo!险类) Then
                                    Bill.TextMask = Replace(Bill.TextMask, "-", "")
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Case "单价"
            Bill.TextLen = 10
            Bill.TextMask = "0123456789." & Chr(8)
    End Select
   
    
    '显示摘要
    If mobjBill.Details.Count >= Bill.Row Then
        If mobjBill.Details(Bill.Row).摘要 <> "" Then
            sta.Panels(2) = sta.Panels(2) & "  摘要:" & mobjBill.Details(Bill.Row).摘要
        End If
    End If
    
    '新行,或更改已有行的类别时,看作换行还没有开始
    If Bill.TextMatrix(Row, BillCol.项目) = "" Then
        mlngPreRow = 0
    ElseIf mobjBill.Details.Count >= Row Then
        mlngPreRow = Row
    End If
End Sub

Private Sub Bill_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'bill.ToolTipText = bill.TextMatrix(bill.MouseRow, bill.MouseCol)
End Sub

Private Sub cboBaby_Click()
    mobjBill.婴儿费 = cboBaby.ListIndex
End Sub

Private Sub cbo开单科室_Click()
    Dim i As Long, strDoctor As String
    
    If mbytInState <> 0 Then Exit Sub
    
    '定位医生
    cbo开单人.Clear
    If cbo开单科室.ListIndex <> -1 Then
        Call Load开单人(cbo开单科室.ItemData(cbo开单科室.ListIndex))
    End If
    
    '数据对象
    If mbytInState = 0 Then
        If cbo开单科室.ListIndex = -1 Then
            mobjBill.开单部门ID = 0
        Else
            mobjBill.开单部门ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
        End If
    End If
    
    If cboBaby.ListCount > 0 Then cboBaby.ListIndex = 0 '触发click事件
    cboBaby.Enabled = DeptIsWoman(mobjBill.开单部门ID)
    
    '重新设置相关项目的执行科室
    If mbytInState = 0 And cbo开单科室.ListIndex <> -1 And cbo开单科室.Visible Then
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
                                Bill.TextMatrix(i, BillCol.执行科室) = Get部门名称(.执行部门ID, mrsUnit)
                            End If
                        Else
                            '浏览单据只(能)显示名称
                            Bill.TextMatrix(i, BillCol.执行科室) = Get部门名称(.执行部门ID, mrsUnit)
                        End If
                    Else
                        Bill.TextMatrix(i, BillCol.执行科室) = ""
                    End If
                    '撤销8113的修改
'                ElseIf InStr(",4,5,6,7,", .Detail.类别) > 0 Then
'                '重设可用药房为存储库房中设置的服务于病人科室(开单科室)的执行科室
'                    If Bill.ColData(BillCol.执行科室) = BillColType.UnFocus Then
'                        Bill.ColData(BillCol.执行科室) = BillColType.ComboBox
'                    End If
'                    If .Key = "1" Then .Key = "0"        '1表示执行科室不可选择
'                    mlngPreRow = 0      '用于在Entercell事件中再重视执行科室的可选性
                End If
            End With
        Next
    End If
End Sub

Private Sub cbo开单人_Click()
    Dim arrDepts As Variant, i As Long, k As Long
    
    If mbytInState = 0 Then
        mobjBill.开单人 = IIF(cbo开单人.ListIndex = -1, "", NeedName(cbo开单人.Text))
                        
        '护士类别
        If Bill.Active Then
            If mobjBill.Details.Count < Bill.Rows - 1 And Bill.Row = Bill.Rows - 1 _
                And Bill.RowData(Bill.Rows - 1) <> 0 Then
                '清除无效输入
                Bill.TextMatrix(Bill.Rows - 1, 1) = ""
                Bill.RowData(Bill.Rows - 1) = 0
            ElseIf Bill.Col = 1 Then
                Call Bill_EnterCell(Bill.Row, Bill.Col) '刷新
            End If
        End If
        
        '护士类别:判断非法输入
'        If HaveStopClass > 0 Then
'            MsgBox "护士只能输入治疗及材料项目,而单据中存在其它类型的项目。", vbInformation, gstrSysName
'        End If
    End If
End Sub

Private Sub cbo开单人_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnDrop = False
    If KeyCode = 13 Then mblnDrop = SendMessage(cbo开单人.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 1
End Sub


Private Sub cboBaby_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk急诊_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk加班_Click()
    If mbytInState = 1 Then Exit Sub
    If mbytInState = 2 Then Exit Sub
    If Not chk加班.Visible Then Exit Sub
    
    Dim blnAdd As Boolean
    
    blnAdd = OverTime
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
    mobjBill.加班标志 = IIF(chk加班.Value = Checked, 1, 0)
    
    '重新计算价格
    If Not mobjBill.Details.Count = 0 Then
        Call CalcMoneys
        Call ShowDetails
        Call ShowMoney
    End If
End Sub

Private Sub chk加班_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    If mobjBill.Details.Count > 0 Or gblnOK Then
        If MsgBox("确实要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    Unload Me
End Sub

Private Function CheckNegative() As Boolean
'功能：检查单据中输入的负数数量及退回科室是否正确
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strItems As String, str部门 As String
    Dim str单位 As String, dbl数量 As Double
    
    CheckNegative = True
    If mobjBill.病人ID = 0 Then Exit Function
    
    For i = 1 To mobjBill.Details.Count
        With mobjBill.Details(i)
            If .数次 < 0 And .执行部门ID <> 0 Then
                strItems = strItems & ",(" & .收费细目ID & "," & .执行部门ID & ")"
                strSQL = strSQL & " Union ALL Select " & .收费细目ID & "," & .执行部门ID & ",0 From Dual"
            End If
        End With
    Next
    strItems = Mid(strItems, 2)
    If strItems = "" Then Exit Function
    
    strSQL = _
        " Select 收费细目ID,执行部门ID,Sum(Nvl(付数,1)*数次) as 数量" & _
        " From 病人费用记录" & _
        " Where (收费细目ID+0,执行部门ID+0) IN(" & strItems & ")" & _
        " And 记录状态<>0 And 记帐费用=1 And 价格父号 is NULL" & _
        " And 病人ID=" & mobjBill.病人ID & " And Nvl(主页ID,0)=" & mobjBill.主页ID & _
        " Group by 收费细目ID,执行部门ID" & strSQL
    strSQL = "Select 收费细目ID,执行部门ID,Sum(数量) as 数量 From (" & strSQL & ") Group by 收费细目ID,执行部门ID"
    
    On Error GoTo errH
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption) 'Union:个数不定
    For i = 1 To mobjBill.Details.Count
        With mobjBill.Details(i)
            If .数次 < 0 And .执行部门ID <> 0 Then
                rsTmp.Filter = "收费细目ID=" & .收费细目ID & " And 执行部门ID=" & .执行部门ID
                If Not rsTmp.EOF Then
                    If InStr(",5,6,7,", .收费类别) > 0 Then
                        str单位 = .Detail.药房单位
                        dbl数量 = Nvl(rsTmp!数量, 0) / .Detail.药房包装
                    Else
                        str单位 = .Detail.计算单位
                        dbl数量 = Nvl(rsTmp!数量, 0)
                    End If
                    str部门 = Get部门名称(.执行部门ID)
                    If Abs(.数次) * .付数 > dbl数量 Then
                        MsgBox "第 " & i & " 行[" & .Detail.名称 & "]退回" & str部门 & "的数量 " & FormatEx(Abs(.数次) * .付数, 5) & str单位 & _
                            " 多于已计费数量 " & FormatEx(dbl数量, 5) & str单位 & "。", vbInformation, gstrSysName
                        CheckNegative = False: Exit Function
                    End If
                End If
            End If
        End With
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdOK_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strInfo As String, strSQL As String, strTmp As String
    Dim i As Long, j As Long, lng结帐ID As Long
    Dim bln医保 As Boolean, cur当日额 As Currency
    Dim curTotal As Currency, intInsure As Integer
    Dim dblTotal As Double, cur余额 As Currency
    Dim colStock As Collection
    
    If mbytInState = 3 Then
        If mint记录性质 <> 1 And False Then '划价是全部删除
            For i = 1 To Bill.Rows - 1
                If Bill.TextMatrix(i, Bill.Cols - 1) = "√" And Bill.RowData(i) > 0 Then
                    strSQL = strSQL & "," & Bill.RowData(i)
                End If
            Next
            If strSQL = "" Then
                MsgBox "请至少选择一行要删除的费用！", vbInformation, gstrSysName
                Bill.SetFocus: Exit Sub
            End If
            
            '所有行选择处理
            strSQL = Mid(strSQL, 2)
            i = GetBillRows(mstrInNO, mint记录性质)
            If UBound(Split(strSQL, ",")) + 1 = i Then strSQL = ""
        Else
            '因为要处理为全退，如果结帐后不允许销帐，部份结帐后就要检查
            j = 0
            For i = 1 To Bill.Rows - 1
                If Bill.RowData(i) > 0 Then j = j + 1
            Next
            i = GetBillRows(mstrInNO, mint记录性质)
            If j < i Then
                MsgBox "单据中的部份项目当前已不允许销帐(比如已结帐的项目)。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        '医保记帐作废上传(注意判断顺序)
        If mint病人来源 = 2 Then
            intInsure = BillExistInsure(mstrInNO) '判断是否医保病人记的帐
            If intInsure > 0 Then
                If gclsInsure.GetCapability(support记帐作废上传, , intInsure) Then
                    '去掉了医保连接匹配检查
                    If strSQL <> "" Then '不能部分销帐
                        MsgBox "因为医保处理需要,该单据中的项目必须全部删除！", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
            End If
        End If
        
        If mint病人来源 = 2 Then
            strSQL = "zl_住院记帐记录_DELETE('" & mstrInNO & "','" & strSQL & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
        Else
            If mint记录性质 = 2 Then
                strSQL = "zl_门诊记帐记录_DELETE('" & mstrInNO & "','" & strSQL & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
            Else
                strSQL = "zl_门诊划价记录_DELETE('" & mstrInNO & "')"
            End If
        End If
        
        On Error GoTo errH
        gcnOracle.BeginTrans
        
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        
        '医保记帐作废上传
        If mint病人来源 = 2 And intInsure > 0 Then
            If gclsInsure.GetCapability(support记帐作废上传, , intInsure) And Not gclsInsure.GetCapability(support记帐完成后上传, , intInsure) Then
                If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                    gcnOracle.RollbackTrans: Exit Sub
                End If
            End If
        End If
        
        gcnOracle.CommitTrans
        
        '医保记帐作废上传
        If mint病人来源 = 2 And intInsure > 0 Then
            If gclsInsure.GetCapability(support记帐作废上传, , intInsure) And gclsInsure.GetCapability(support记帐完成后上传, , intInsure) Then
                If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                    MsgBox "单据""" & mstrInNO & """的删费数据向医保传送失败，该单据已删除。", vbInformation, gstrSysName
                End If
            End If
        End If
        
        On Error GoTo 0
        
        gblnOK = True: Unload Me: Exit Sub
    Else '正常输入单据状态
        If mobjBill.病人ID = 0 Or mrsInfo.State = 0 Then
            MsgBox "没有发现病人信息，单据不能保存。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If mobjBill.Details.Count = 0 Then
            MsgBox "单据中没有任何内容,请正确输入单据内容！", vbInformation, gstrSysName
            Bill.SetFocus: Exit Sub
        End If
        
        i = Check执行科室
        If i <> 0 Then
            MsgBox "单据中第 " & i & " 行项目没有指定执行科室！", vbInformation, gstrSysName
            Bill.SetFocus: Exit Sub
        End If
        
        If cbo开单科室.ListIndex = -1 Then
            MsgBox "请确定开单科室！", vbInformation, gstrSysName
            cbo开单科室.SetFocus: Exit Sub
        End If
        
        If cbo开单人.ListIndex = -1 Then
            MsgBox "请输入开单人！", vbInformation, gstrSysName
            cbo开单人.SetFocus: Exit Sub
        End If
        
        '护士类别:判断非法输入
'        If HaveStopClass > 0 Then
'            MsgBox "护士只能输入治疗及材料项目,而单据中存在其它类型的项目。", vbInformation, gstrSysName
'            Exit Sub
'        End If
                
        '发生时间检查
        If Not IsDate(txtDate.Text) Then
            MsgBox "请输入正确的费用日期！", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        
        '出院强制记帐权限检查
        If mint病人来源 = 2 Then
            If Not PatiCanBilling(mrsInfo!病人ID, Nvl(mrsInfo!主页ID, 0), mstrPrivs) Then Exit Sub
        End If
                
        '发生时间检查
        If Not IsNull(mrsInfo!出院日期) Then
            If Format(txtDate.Text, txtDate.Format) > Format(mrsInfo!出院日期, txtDate.Format) Then
                MsgBox "强制对出院病人记帐时，费用时间不能大于病人出院时间:" & Format(mrsInfo!出院日期, txtDate.Format), vbInformation, gstrSysName
                txtDate.SetFocus: Exit Sub
            End If
        End If
        If Not IsNull(mrsInfo!入院日期) Then
            If Format(txtDate.Text, txtDate.Format) < Format(mrsInfo!入院日期, txtDate.Format) Then
                MsgBox "费用的发生时间不能小于病人的入院时间:" & Format(mrsInfo!入院日期, txtDate.Format), vbInformation, gstrSysName
                txtDate.SetFocus: Exit Sub
            End If
        End If
        
        '非法行
        For i = 1 To mobjBill.Details.Count
            If mobjBill.Details(i).收费细目ID = 0 Then
                MsgBox "单据中第 " & i & " 行没有正确输入数据,请修正或删除该行！", vbInformation, gstrSysName
                Bill.SetFocus: Exit Sub
             '8407
'            ElseIf InStr(1, ",5,6,7,", mobjBill.Details(i).收费类别) > 0 Then
'                '收集药品的发药药房
'                strTmp = strTmp & "," & mobjBill.Details(i).收费细目ID
            End If
        Next
        
'        '检查药品的发药药房对应的服务科室(存储库房)
'        If strTmp <> "" Then
'            strTmp = Mid(strTmp, 2)
'            Set rsTmp = GetServiceDept(strTmp)
'            If Not rsTmp Is Nothing Then
'                strTmp = ""
'                For i = 1 To mobjBill.Details.Count
'                    If InStr(1, ",5,6,7,", mobjBill.Details(i).收费类别) > 0 Then
'                        strInfo = mobjBill.Details(i).收费细目ID
'                        '先检查是否是允许的存储库房
'                        rsTmp.Filter = "收费细目ID=" & strInfo & " And 执行科室id=" & mobjBill.Details(i).执行部门ID
'                        If rsTmp.RecordCount = 0 Then
'                            strTmp = strTmp & "," & i
'                        Else
'                            '再检查是否是允许的服务科室(没有设置服务科室的,开单科室ID为零)
'                            rsTmp.Filter = "(" & rsTmp.Filter & " And 开单科室ID=" & mobjBill.开单部门ID & ") Or (" & rsTmp.Filter & " And 开单科室ID=0)"
'                            If rsTmp.RecordCount = 0 Then
'                                strTmp = strTmp & "," & i
'                            End If
'                        End If
'                    End If
'                Next
'                If strTmp <> "" Then
'                    strTmp = Mid(strTmp, 2)
'                    MsgBox "请检查,第" & strTmp & "行药品是否违反以下规则:" & vbCrLf & vbCrLf & _
'                        "A.选择的执行科室不是药品的存储库房" & vbCrLf & _
'                        "B.开单科室[" & NeedName(cbo开单科室.Text) & "]不属于药品在此存储库房的服务科室.", _
'                        vbInformation, gstrSysName
'                    Exit Sub
'                End If
'            End If
'        End If
        
        
        '*********此模块不存在先输单据再确定病人的情况**********
        '医保负数记帐检查    因为操作员可能先输单据,再确定病人,所以要再检查一次(此处不用判断权限,因为有权限才可能是负数)
'        If InStr(mstrPrivs, "负数费用") > 0 And mint记录性质 = 2 Then    '至少有其中一种负数记帐权限,才可能有负数
'            If Not IsNull(mrsInfo!险类) Then
'                If Not gclsInsure.GetCapability(support负数记帐, , mrsInfo!险类) Then
'                    For i = 1 To mobjBill.Details.Count
'                        If mobjBill.Details(i).数次 * mobjBill.Details(i).付数 < 0 Then
'                            MsgBox "单据中第 " & i & " 行是负数,本地医保不支持负数记帐！", vbInformation, gstrSysName
'                            Bill.SetFocus: Exit Sub
'                        End If
'                    Next
'                End If
'            End If
'        End If
                
        '处方职务检查
        If InStr(",1,2,", txt付款方式.Tag) > 0 Then '公费或医保病人
            i = CheckDuty(, False)
            If i > 0 Then
                Bill.Row = i: Bill.MsfObj.TopRow = i
                Bill.Col = 2: Bill.SetFocus
                Exit Sub
            End If
        End If

        '所有病人项目
        i = CheckDuty(, True)
        If i > 0 Then
            Bill.Row = i: Bill.MsfObj.TopRow = i
            Bill.Col = 2: Bill.SetFocus
            Exit Sub
        End If
        
        '费用类型检查
        If Not Check费用类型 Then Exit Sub
        
        '要求审批,医保费用项目审批检查
        If mint病人来源 = 2 And mint记录性质 = 2 Then
            If Not IsNull(mrsInfo!险类) And Not mrsMedAudit Is Nothing Then
                If Not CheckExamine(mobjBill.Details, mrsMedAudit, mrsInfo!险类) Then Exit Sub
            End If
        End If
        
        '记帐分类报警
        If mint记录性质 = 2 And mrsWarn.State = 1 And mstrWarn <> "-" Then
            '单据费用
            curTotal = CalcGridToTal
            If curTotal > 0 Then
                '刷新病人费用状况
                Set rsTmp = GetMoneyInfo(mrsInfo!病人ID, mcurModiMoney, True)
                If Not rsTmp Is Nothing Then
                    cmdOK.Tag = rsTmp!预交余额
                    cmdCancel.Tag = rsTmp!费用余额
                    txt实收.Tag = rsTmp!预交余额 - rsTmp!费用余额
                Else
                    cmdOK.Tag = 0: cmdCancel.Tag = 0: txt实收.Tag = 0
                End If
                sta.Panels(3).Text = "预交:" & Format(Val(cmdOK.Tag), "0.00")
                sta.Panels(3).Text = sta.Panels(3).Text & "/费用:" & Format(Val(cmdCancel.Tag) + curTotal, gstrDec)
                sta.Panels(3).Text = sta.Panels(3).Text & "/剩余:" & Format(Val(txt实收.Tag) - curTotal, "0.00")
                
                '重新读取当日额
                cur当日额 = GetPatiDayMoney(mrsInfo!病人ID)
                
                '是否医保病人
                bln医保 = txt付款方式.Tag = "1" Or Not IsNull(mrsInfo!险类)
                cur余额 = Val(txt实收.Tag)
                If gbln报警包含划价费用 Then cur余额 = cur余额 - GetPriceMoneyTotal(mrsInfo!病人ID)
                        
                For i = 1 To mobjBill.Details.Count
                    mbytWarn = BillingWarn(Me, mstrPrivs, mrsWarn, mrsInfo!姓名, cur余额, cur当日额 - mcurModiMoney, curTotal, IIF(IsNull(mrsInfo!担保额), 0, mrsInfo!担保额), mobjBill.Details(i).收费类别, mobjBill.Details(i).Detail.类别名称, mstrWarn, mintWarn, bln医保)
                    If mbytWarn = 2 Or mbytWarn = 3 Then Exit Sub
                Next
            End If
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
        If Not CheckLimit(mobjBill, , mbln药房单位) Then Exit Sub
        
        '检查分批或时价药品同一药房是否有重复输入
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
                If (.Detail.分批 Or .Detail.变价) _
                    And (InStr(",5,6,7,", .收费类别) > 0 Or .收费类别 = "4" And .Detail.跟踪在用) Then
                    For j = 1 To mobjBill.Details.Count
                        If i <> j And .收费细目ID = mobjBill.Details(j).收费细目ID And .执行部门ID = mobjBill.Details(j).执行部门ID Then
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
        
        '药品库存检查(仅不足禁止时或分批时价药品)
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
                Set colStock = IIF(.收费类别 = "4", mcolStock2, mcolStock1)
                If InStr(",5,6,7,", .收费类别) > 0 Then
                    If .Detail.分批 Or .Detail.变价 Then
                        dblTotal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
                        .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                        If mbln药房单位 Then
                            .Detail.库存 = .Detail.库存 / .Detail.药房包装
                        End If
                        If dblTotal > .Detail.库存 Then
                            MsgBox "第 " & i & " 行时价或分批药品""" & .Detail.名称 & _
                                """的当前库存""" & .Detail.库存 & """不足输入数量""" & dblTotal & """。", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    ElseIf colStock("_" & .执行部门ID) = 2 Then
                        dblTotal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
                        .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                        If mbln药房单位 Then
                            .Detail.库存 = .Detail.库存 / .Detail.药房包装
                        End If
                        If dblTotal > .Detail.库存 Then
                            MsgBox "第 " & i & " 行药品""" & .Detail.名称 & _
                                """的当前库存""" & .Detail.库存 & """不足输入数量""" & dblTotal & """,请修改或检查是否有多行输入。", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                ElseIf .收费类别 = "4" And .Detail.跟踪在用 Then
                    If .Detail.分批 Or .Detail.变价 Then
                        dblTotal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
                        .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                        If dblTotal > .Detail.库存 Then
                            MsgBox "第 " & i & " 行时价或分批卫生材料""" & .Detail.名称 & _
                                """的当前库存""" & .Detail.库存 & """不足输入数量""" & dblTotal & """。", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    ElseIf colStock("_" & .执行部门ID) = 2 Then
                        dblTotal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
                        .Detail.库存 = GetStock(.收费细目ID, .执行部门ID)
                        If dblTotal > .Detail.库存 Then
                            MsgBox "第 " & i & " 行卫生材料""" & .Detail.名称 & _
                                """的当前库存""" & .Detail.库存 & """不足输入数量""" & dblTotal & """,请修改或检查是否有多行输入。", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                End If
            End With
        Next
        
        '检查卫生材料的灭菌效期
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
                If .收费类别 = "4" And .Detail.跟踪在用 Then
                    dblTotal = GetDrugTotal(mobjBill, .收费细目ID, .执行部门ID)
                    If Not CheckValidity(.收费细目ID, .执行部门ID, dblTotal) Then Exit Sub
                End If
            End With
        Next
        
        '负数退费检查
        If mint记录性质 = 2 Then
            If Not CheckNegative Then Exit Sub
        End If
        
        If Not SaveBill Then Exit Sub
        
        If mstrInNO <> "" Then
            gblnOK = True: Unload Me: Exit Sub
        Else
            sta.Panels(2) = "上一张单据:" & mobjBill.NO
            Call ClearRows: Call Bill.ClearBill
            Call SetColNum: Call ClearMoney
            Call SetMoneyList
            Call NewBill
            
            '重新读取病人信息
            Call GetPatient(mlng病人ID, mlng主页ID)
            
            Bill.SetFocus
        End If
    End If
    gblnOK = True
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdOK_GotFocus()
    If mbytInState = 3 Then
        Bill.Row = 1: Bill.Col = Bill.Cols - 1
    End If
End Sub

Private Sub Form_Activate()
    If mbytInState <> 0 Then
        If cmdOK.Visible And cmdOK.Enabled Then
            cmdOK.SetFocus
        ElseIf cmdCancel.Visible And cmdCancel.Enabled Then
            cmdCancel.SetFocus
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("',;|~" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim tmpBill As ExpenseBill, i As Long
    
    glngFormW = 12000: glngFormH = 7710
    If Not InDesign Then
        glngOld = GetWindowLong(Me.Hwnd, GWL_WNDPROC)
        Call SetWindowLong(Me.Hwnd, GWL_WNDPROC, AddressOf Custom_WndMessage)
    End If
    Call RestoreWinState(Me, App.ProductName, mbytInState)
    
    gblnOK = False
    mblnEnterCell = True
    mintWarn = -1: mstrWarn = ""
    Call InitLocPar
    
    '初始化单据数据
    Set mobjBill = New ExpenseBill
    If mbytInState = 0 Then
        If Not InitData Then
            Unload Me: Exit Sub
        End If
    End If
    Call InitFace
    Call NewBill
    
    If mbytInState <> 0 Then
        If Not ReadBill(mstrInNO, mbytInState = 3) Then
            Unload Me: Exit Sub
        End If
    Else
        '读取该单据的内容
        If mstrInNO <> "" Then '修改单据
            Set mobjBill = ImportBill(mint病人来源, mstrInNO, mint记录性质, mbln费用登记)
            If mobjBill.NO = "" Then
                MsgBox "不能正确读取计费单据的内容！", vbInformation, gstrSysName
                Unload Me: Exit Sub
            Else
                Bill.ClearBill: Call SetColNum
                Bill.Rows = mobjBill.Details.Count + 1
                
                '针对列编辑性质设置颜色
                Bill.SetColColor 1, &HE7CFBA
                Bill.SetColColor 2, &HE7CFBA
                Bill.SetColColor 6, &HE7CFBA
                Bill.SetColColor 10, &HE7CFBA
                Bill.SetColColor 5, &HE0E0E0
                Bill.SetColColor 7, &HE0E0E0
                Bill.SetColColor 11, &HE0E0E0
                
                cboNO.Text = mobjBill.NO
                
                mobjBill.开单部门ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
                Call GetCboIndex(cbo开单人, mobjBill.开单人, True)
                mobjBill.开单人 = NeedName(cbo开单人.Text)
                
                Call zlControl.CboSetIndex(cboBaby.Hwnd, mobjBill.婴儿费)
                If cbo开单科室.ListIndex <> -1 Then cboBaby.Enabled = DeptIsWoman(mobjBill.开单部门ID)
                
                mobjBill.操作员编号 = UserInfo.编号
                mobjBill.操作员姓名 = UserInfo.姓名
                
                If mint记录性质 = 2 Then
                    mcurModiMoney = GetBillMoney(mobjBill.NO) '在读取病人前取
                End If
                
                '新单时读取病人,看单据时根据单据显示病人信息
                Call GetPatient(mlng病人ID, mlng主页ID)
                If mrsInfo.State = 0 Then
                    MsgBox "不能读取病人信息，可能是你不具有对该病人计费的权限。", vbInformation, gstrSysName
                    Unload Me: Exit Sub
                End If
                
                If gbln从项汇总折扣 Then CalcMoneys
                Call ShowDetails
                Call ShowMoney
                
                '调整库存:修改时加上将要退回的库存
                For i = 1 To mobjBill.Details.Count
                    With mobjBill.Details(i)
                        Bill.RowData(i) = Asc(.收费类别) '特殊处理
                        If InStr(",5,6,7,", .收费类别) > 0 Then
                            .Detail.库存 = .Detail.库存 + .付数 * .数次
                        ElseIf .收费类别 = "4" And .Detail.跟踪在用 Then
                            .Detail.库存 = .Detail.库存 + .付数 * .数次
                        End If
                    End With
                Next
                
                Call SetColNum
            End If
        Else
            '新单时读取病人,看单据时根据单据显示病人信息
            Call GetPatient(mlng病人ID, mlng主页ID)
            If mrsInfo.State = 0 Then
                MsgBox "不能读取病人信息，可能是你不具有对该病人计费的权限。", vbInformation, gstrSysName
                Unload Me: Exit Sub
            End If
        End If
        
        If mstrInNO <> "" And mint记录性质 = 2 And mint病人来源 = 2 Then
            Call ReCalcInsure '重新计算统筹金额
        End If
        
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Bill.Height = Me.ScaleHeight - picAppend.Height - sta.Height - fraTitle.Height - fraInfo.Height + 230
    
    fraTitle.Width = Me.ScaleWidth - fraTitle.Left
    
    cboNO.Left = fraTitle.Width - cboNO.Width - 90
    lblNO.Left = cboNO.Left - lblNO.Width - 45
        
    fraUnit.Left = Me.ScaleWidth - fraUnit.Width
    fraInfo.Width = Me.ScaleWidth - fraUnit.Width - fraInfo.Left
    
    Bill.Width = Me.ScaleWidth - Bill.Left
    
    fraAppend.Width = Me.ScaleWidth - fraAppend.Left
    
    txtDate.Left = fraAppend.Width - txtDate.Width - 90
    lblDate.Left = txtDate.Left - lblDate.Width - 45
        
    cbo开单人.Left = lblDate.Left - cbo开单人.Width - 200
    lbl开单人.Left = cbo开单人.Left - lbl开单人.Width - 45
    
    cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - 500
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 200

    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName, mbytInState)
    
    mlng医嘱ID = 0
    mlng发送号 = 0
    mlng病人ID = 0
    mlng主页ID = 0
    mint病人来源 = 0
    mint记录性质 = 0
    mbln费用登记 = False
    mlng开单科室ID = 0
    mlng病人科室id = 0
    
    mlng开嘱科室ID = 0
    mstr开嘱医生 = ""
    
    mlng药品类别ID = 0
    mlng卫材类别ID = 0
    
    mbytInState = 0
    mstrInNO = ""
    mstrTime = ""
    mblnDelete = False
    mstrPrivs = ""
    
    Set mrsInfo = Nothing
    Set mrsUnit = Nothing
    Set mrsClass = Nothing
    Set mrsWork = Nothing
    Set mrsMedAudit = Nothing
    
    If Not InDesign Then
        Call SetWindowLong(Me.Hwnd, GWL_WNDPROC, glngOld)
    End If
End Sub

Private Sub sta_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Bevel = sbrRaised And (Panel.Key = "PY" Or Panel.Key = "WB") Then
        '切换并保存简码匹配方式
        Panel.Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        If Panel.Key = "PY" Then
            sta.Panels("WB").Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        Else
            sta.Panels("PY").Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        End If
        SaveSetting "ZLSOFT", "私有全局\" & gstrDBUser, "简码生成", _
            IIF(sta.Panels("PY").Bevel = sbrInset And sta.Panels("WB").Bevel = sbrInset, 2, IIF(sta.Panels("WB").Bevel = sbrInset, 1, 0))
    End If
End Sub

Private Sub txtDate_GotFocus()
    txtDate.SelStart = 0
    txtDate.SelLength = Len(txtDate.Text)
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And IsDate(txtDate.Text) Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtDate_LostFocus()
    txtDate.SelLength = 0
End Sub

Private Sub bill_AfterAddRow(Row As Long)
    Dim i As Long
    If mbytInState = 3 Then
        Bill.Row = 1: Call zlCommFun.PressKey(vbKeyTab)
        Exit Sub
    End If
    
    
    With Bill
        '新增行时,重新设置可能已经被更改的可变性质列的列值
        If mbytInState <> 2 Then
            .ColData(1) = IIF(gbln收费类别, BillColType.ComboBox, BillColType.UnFocus) '类别列,当主从项时会被改变
            .ColData(2) = BillColType.CommandButton  '项目列,当主从项时会被改变
            .ColData(5) = 5 '付数缺省跳过(=1),当类别为中药时,设为输入(4)(有值,一改全改)
            .ColData(7) = 5 '单价缺省跳过,当项目变价时,设为输入(4)
            .ColData(11) = 5 '标志缺省跳过,当为手术时,设为复选(-1)
        End If
        
        '针对列编辑性质设置颜色
        .SetColColor 1, &HE7CFBA
        .SetColColor 2, &HE7CFBA
        .SetColColor 6, &HE7CFBA
        .SetColColor 10, &HE7CFBA
        .SetColColor 5, &HE0E0E0
        .SetColColor 7, &HE0E0E0
        .SetColColor 11, &HE0E0E0
        
        .TextMatrix(Row, 0) = Row
        
        '特殊地方手动调用不执行
        If Row > 0 And .ColData(1) <> 5 And Me.Visible And Not mblnNewRow Then
            Call zlCommFun.PressKey(13)
        End If
    End With
End Sub

Private Sub cbo开单科室_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii = 13 And cbo开单科室.ListIndex <> -1 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf KeyAscii >= 32 And Not cbo开单科室.Locked Then
        lngIdx = zlControl.CboMatchIndex(cbo开单科室.Hwnd, KeyAscii)
        If lngIdx = -1 And cbo开单科室.ListCount > 0 Then lngIdx = 0
        cbo开单科室.ListIndex = lngIdx
    End If
End Sub

Private Sub cbo开单人_KeyPress(KeyAscii As Integer)
    Dim i As Long, intIdx As Integer
    Dim strText As String
    
    If KeyAscii = 13 Then
        strText = cbo开单人.Text
        If strText = "" Then
            cbo开单人.ListIndex = -1
        ElseIf cbo开单人.ListIndex = -1 Then
            intIdx = -1
            If IsNumeric(strText) Then
                For i = 0 To cbo开单人.ListCount - 1
                    If i > UBound(marrDr) Then Exit For
                    If CStr(Split(marrDr(i), "|")(2)) = strText Then
                        If intIdx = -1 Then cbo开单人.ListIndex = i
                        intIdx = i
                    End If
                Next
                If intIdx = -1 Then
                    For i = 0 To cbo开单人.ListCount - 1
                        If i > UBound(marrDr) Then Exit For
                        If Val(Split(marrDr(i), "|")(2)) = Val(strText) Then
                            If intIdx = -1 Then cbo开单人.ListIndex = i
                            intIdx = i
                        End If
                    Next
                End If
            Else
                For i = 0 To cbo开单人.ListCount - 1
                    If UCase(cbo开单人.List(i)) Like UCase(strText) & "*" Then
                        If intIdx = -1 Then cbo开单人.ListIndex = i
                        intIdx = i
                    End If
                Next
            End If
        ElseIf Not mblnDrop Then
            '回车光标经过
            Call cbo开单人_Click
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        If cbo开单人.ListIndex = -1 Then
            cbo开单人.Text = ""
            mobjBill.开单人 = ""
        Else
            mobjBill.开单人 = NeedName(cbo开单人.Text)
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
            ShowHelp "zl9InExse", Me.Hwnd, "frmCharge"
        Case vbKeyF2
            If cmdOK.Enabled And cmdOK.Visible Then Call cmdOK_Click
        Case vbKeyF6 '清除当前单据内容,进入新单状态
            If mbytInState = 0 Then
                txt实收.Text = gstrDec: txt应收.Text = gstrDec
                Call ClearRows: Call Bill.ClearBill
                Call SetColNum: Call ClearMoney
                Call NewBill
                Bill.SetFocus
            End If
        Case vbKeyF7 '切换输入法
            If sta.Panels("WB").Visible And sta.Panels("PY").Visible Then
                If sta.Panels("WB").Bevel = sbrRaised Then
                    Call sta_PanelClick(sta.Panels("WB"))
                Else
                    Call sta_PanelClick(sta.Panels("PY"))
                End If
            End If
        Case vbKeyQ
            If Shift = vbCtrlMask Then Call LocateNewRow
        Case vbKeyEscape, vbKeyX
            If KeyCode = vbKeyX And Shift <> 4 Then Exit Sub
            If Bill.TxtVisible Then
                Bill.Text = "": Bill.TxtVisible = False: Bill.SetFocus
            Else
                Call cmdCancel_Click
            End If
    End Select
End Sub

Private Sub SetMoneyList()
'功能:根据当前收入项目行数调整各列宽
    Dim lngW As Long
    lngW = mshMoney.Width - 60
    If mshMoney.Rows > mshMoney.Height / mshMoney.RowHeight(0) Then
        lngW = lngW - 250
    End If
    mshMoney.ColWidth(0) = lngW * 0.5
    mshMoney.ColWidth(1) = lngW * 0.5
    
    mshMoney.ColAlignment(0) = 1
    mshMoney.ColAlignment(1) = 7
    
    mshMoney.TextMatrix(0, 0) = "项目"
    mshMoney.TextMatrix(0, 1) = "金额"
    mshMoney.Row = 0
    mshMoney.ColAlignmentFixed(0) = 4
    mshMoney.ColAlignmentFixed(1) = 4
End Sub

Private Function InitData() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    On Error GoTo errH
    
    '不同药房药品出库检查方式
    Set mcolStock1 = GetStockCheck(0)
    Set mcolStock2 = GetStockCheck(1)
    
    '开单科室
    strSQL = "Select 开嘱科室ID,开嘱医生 From 病人医嘱记录 Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng医嘱ID)
    If Not rsTmp.EOF Then
        mlng开嘱科室ID = Nvl(rsTmp!开嘱科室ID, 0)
        mstr开嘱医生 = Nvl(rsTmp!开嘱医生)
    End If
    If mlng开单科室ID = 0 Or mstr开嘱医生 = "" Then
        MsgBox "没有发现源医嘱信息。", vbInformation, gstrSysName
        Exit Function
    End If
    If mbln费用登记 Then
        '就为当前选择的医技科室
        strSQL = "Select ID,编码,名称,简码 From 部门表 Where ID=[1]"
    Else
        '就为当前选择的医技科室或开嘱科室
        strSQL = "Select ID,编码,名称,简码 From 部门表 Where ID IN([1],[2]) Order by 编码"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng开单科室ID, mlng开嘱科室ID)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo开单科室.AddItem rsTmp!编码 & "-" & rsTmp!名称
            cbo开单科室.ItemData(cbo开单科室.ListCount - 1) = rsTmp!ID
            If rsTmp!ID = mlng开单科室ID Then
                cbo开单科室.ListIndex = cbo开单科室.NewIndex
            End If
            rsTmp.MoveNext
        Next
        If cbo开单科室.ListIndex = -1 Then cbo开单科室.ListIndex = 0
    Else
        MsgBox "不能确定开单科室，请先到部门管理中设置。", vbInformation, gstrSysName
        Exit Function
    End If
    
    '可用收费类别:"'5','E','Z'"
    If mstr收费类别 = "" Then
        strSQL = "Select 编码,名称 as 类别 From 收费项目类别 Where 编码<>'1' Order by 序号"
    Else
        strSQL = "Select 编码,名称 as 类别 From 收费项目类别 Where Instr([1],编码)>0 Order by 序号"
    End If
    'Set mrsClass = New ADODB.Recordset
    Set mrsClass = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr收费类别)
    If mrsClass.EOF Then
        MsgBox "没有设置可用的收费类别,请先在本地参数中设置！", vbInformation, gstrSysName
        Exit Function
    End If
    '当只有一种可选收费类别时,不用用户选择
    mblnOne = (mrsClass.RecordCount = 1)
    If InStr(mstr收费类别, "'5'") > 0 Or InStr(mstr收费类别, "'6'") > 0 Or InStr(mstr收费类别, "'7'") > 0 Or mstr收费类别 = "" Then
        mlng药品类别ID = ExistIOClass(IIF(mint记录性质 = 1, 8, 9))
        If mlng药品类别ID = 0 Then
            MsgBox "不能确定药品单据的入出类别,请先到入出分类管理中设置！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If InStr(mstr收费类别, "'4'") > 0 Or mstr收费类别 = "" Then
        mlng卫材类别ID = ExistIOClass(IIF(mint记录性质 = 1, 40, 41))
        If mlng卫材类别ID = 0 Then
            MsgBox "不能确定卫材单据的入出类别,请先到入出分类管理中设置！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '执行部门
    strSQL = _
        "Select Distinct A.ID,A.编码,A.简码,A.名称,B.工作性质,B.服务对象 " & _
        " From 部门表 A,部门性质说明 B " & _
        " Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " And B.部门ID=A.ID and B.服务对象 IN([1],3) " & _
        " Order by B.服务对象,A.编码"
    'Set mrsUnit = New ADODB.Recordset
    Set mrsUnit = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mint病人来源)
    If mrsUnit.EOF Then
        MsgBox "没有初始化部门信息,单据无法处理执行部门。请先到部门管理中设置！", vbInformation, gstrSysName
        Exit Function
    End If
    InitData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
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

Private Sub FillBillComboBox(ByVal lngRow As Long, ByVal lngCol As Long, Optional blnEnter As Boolean)
'功能：根据单据列设置下拉列表框内容
'参数：blnEnter=是否按进入该列处理,比如执行科室保持不变
    Dim rsTmp As New ADODB.Recordset
    Dim str人员性质 As String, strTmp As String
    Dim lng科室ID As Long, strIDs As String
    Dim strSQL As String, i As Long
    
    Bill.Clear
    
    Select Case Bill.TextMatrix(0, lngCol)
        Case "类别"
            If cbo开单人.ListIndex <> -1 Then
                If cbo开单人.ListIndex <= UBound(marrDr) Then
                    If UBound(Split(marrDr(cbo开单人.ListIndex), "|")) >= 6 Then
                        str人员性质 = Split(marrDr(cbo开单人.ListIndex), "|")(6)
                    End If
                End If
            End If
        
            mrsClass.Filter = 0
            If mrsClass.RecordCount <> 0 Then
                mrsClass.MoveFirst
                For i = 1 To mrsClass.RecordCount
                    '护士类别:限制
'                    If Not (str人员性质 = "护士" And InStr(",E,M,4,", mrsClass!编码) = 0) Then
                        Bill.AddItem Bill.ListCount + 1 & "-" & mrsClass!类别
                        Bill.ItemData(Bill.NewIndex) = Asc(mrsClass!编码)  '存放类别编码的ASCII码
'                    End If
                    mrsClass.MoveNext
                Next
            End If
        Case "执行科室"
            '根据当前项目执行科室性质,动态设置可选科室
            If mobjBill.Details.Count >= lngRow Then
                With mobjBill.Details(lngRow)
                    If InStr(",4,5,6,7,", .收费类别) > 0 Then
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
                                Bill.AddItem mrsWork!编码 & "-" & mrsWork!名称
                                Bill.ItemData(Bill.NewIndex) = mrsWork!ID
                                If mrsWork!ID = lng科室ID Then Bill.ListIndex = Bill.NewIndex
                                mrsWork.MoveNext
                            Next
                        End If
                    Else
                        lng科室ID = Get开单科室ID
                        Bill.TextMatrix(lngRow, lngCol) = ""
                        '0-不明确,1-病人科室,2-病人病区,3-开单人科室,4-指定科室
                        Select Case .Detail.执行科室
                            Case 0 '不明确
                                mrsUnit.Filter = 0
                            Case 1 '病人科室
                                mrsUnit.Filter = "ID=" & Nvl(mrsInfo!科室ID, 0) & " Or ID=" & .执行部门ID
                            Case 2 '病人病区
                                mrsUnit.Filter = "ID=" & Nvl(mrsInfo!病区ID, 0) & " Or ID=" & .执行部门ID
                            Case 3 '操作员科室
                                mrsUnit.Filter = "ID=" & UserInfo.部门ID & " Or ID=" & .执行部门ID
                            Case 4 '指定科室
                                strSQL = "Select Nvl(开单科室ID,0) as 开单科室ID,执行科室ID" & _
                                    " From 收费执行科室" & _
                                    " Where 收费细目ID=[1]" & _
                                    " And (病人来源 is NULL Or 病人来源=[2])" & _
                                    " And (开单科室ID is NULL Or 开单科室ID=[3])" & _
                                    " Order by Decode(病人来源,Null,2,1)" '默认科室优先
                                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .收费细目ID, mint病人来源, Val(Nvl(mrsInfo!科室ID, 0)))
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
                            Case 6 '开单人科室
                                mrsUnit.Filter = "ID=" & lng科室ID & " Or ID=" & .执行部门ID
                        End Select
                        If mrsUnit.EOF Then mrsUnit.Filter = "ID=" & UserInfo.部门ID & " Or ID=" & .执行部门ID
                        If Not mrsUnit.EOF Then
                            For i = 1 To mrsUnit.RecordCount
                                strTmp = mrsUnit!编码 & "-" & mrsUnit!名称
                                If Not (SendMessage(Bill.CboHwnd, CB_FINDSTRING, -1, ByVal strTmp) >= 0) Then
                                    Bill.AddItem strTmp
                                    Bill.ItemData(Bill.NewIndex) = mrsUnit!ID
                                    
                                    '设置缺省执行科室
                                    If Not blnEnter Then '进入该列时保持已确定值不变
                                        If lngRow = 1 Then
                                            If mrsUnit!ID = lng科室ID Then Bill.ListIndex = Bill.NewIndex
                                        ElseIf lngRow > 1 Then
                                            '与上一行非药品相同
                                            If mrsUnit!ID = mobjBill.Details(lngRow - 1).执行部门ID And mobjBill.Details(lngRow - 1).Detail.执行科室 = .Detail.执行科室 _
                                                And InStr(",5,6,7,", mobjBill.Details(lngRow - 1).收费类别) = 0 Then
                                                Bill.ListIndex = Bill.NewIndex
                                            ElseIf mrsUnit!ID = lng科室ID And Bill.ListIndex = -1 Then
                                                Bill.ListIndex = Bill.NewIndex
                                            End If
                                        End If
                                    End If
                                End If
                                mrsUnit.MoveNext
                            Next
                        End If
                            
                        If Not blnEnter And .Detail.执行科室 = 4 Then '执行科室为指定科室的,缺省为操作员所在科室
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
End Sub

Private Sub InitFace()
'功能：根据表单要完成的功能设置界面布局
    Dim arrHead() As String, i As Long, arrBaby As Variant
    
    '公用单据表格式
    With Bill
        .Font.Size = 10.5
        .CboFont.Size = 11
        .TxtEditFont.Size = 11
        
        arrHead = Split(STR_HEAD, ";")
        .Cols = UBound(arrHead) + 1
        
        .MsfObj.FixedCols = 1
        .MsfObj.ScrollBars = flexScrollBarVertical
        .LocateCol = 2
        .PrimaryCol = 2
        .MsfObj.ColAlignmentFixed(0) = 4
        .TextMatrix(1, 0) = 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(0, i) = Split(arrHead(i), ",")(0)
            .ColWidth(i) = Split(arrHead(i), ",")(1)
            .ColAlignment(i) = Split(arrHead(i), ",")(2)
        Next
                
        If mbytInState = 0 Then
            .ColData(0) = 5
            
            .ColData(1) = IIF(gbln收费类别, 3, 5)
            If mblnOne Then .ColData(1) = 5
            
            .ColData(2) = 1 '项目输入,按扭可选
            .ColData(6) = 4 '数/次输入
            .ColData(3) = 5 '规格跳过
            .ColData(4) = 5 '单位跳过
            .ColData(5) = 5 '付数缺省跳过(=1),当类别为中药时,设为输入(4)(有值,一改全改)
            .ColData(7) = 5 '单价缺省跳过,当项目变价时,设为输入(4)
            .ColData(8) = 5 '应收金额跳过
            .ColData(9) = 5 '实收金额跳过
            .ColData(10) = 3 '默认取开单科室或上一科室
            .ColData(11) = 5 '标志缺省跳过,当为手术时,设为复选(-1)
            .ColData(12) = 5 '类型缺省跳过
        End If
        .SetColColor 1, &HE7CFBA
        .SetColColor 2, &HE7CFBA
        .SetColColor 6, &HE7CFBA
        .SetColColor 10, &HE7CFBA
        .SetColColor 5, &HE0E0E0
        .SetColColor 7, &HE0E0E0
        .SetColColor 11, &HE0E0E0
        
        ReDim marrColData(.Cols - 1)
        For i = 0 To .Cols - 1
            marrColData(i) = .ColData(i)
        Next
        If mbytInState = 3 Then .AllowAddRow = False
    End With
    Call RestoreFlexState(Bill, App.ProductName & "\" & Me.Name & mbytInState)
    Call SetMoneyList

    '读取简码匹配方式
    sta.Panels("PY").Visible = mbytInState = 0
    sta.Panels("WB").Visible = mbytInState = 0
    If mbytInState = 0 Then
        '简码匹配方式：0-拼音,1-五笔
        i = Val(GetSetting("ZLSOFT", "私有全局\" & gstrDBUser, "简码生成", 0))
        If i = 0 Then
            sta.Panels("PY").Bevel = sbrInset
            sta.Panels("WB").Bevel = sbrRaised
        ElseIf i = 1 Then
            sta.Panels("PY").Bevel = sbrRaised
            sta.Panels("WB").Bevel = sbrInset
        Else
            sta.Panels("PY").Bevel = sbrInset
            sta.Panels("WB").Bevel = sbrInset
        End If
    End If

    '标题
    If mbln费用登记 Then
        lblTitle.Caption = gstrUnitName & "零费耗用登记"
    Else
        If mint记录性质 = 1 Then
            lblTitle.Caption = gstrUnitName & "病人收费单"
        ElseIf mint记录性质 = 2 Then
            lblTitle.Caption = gstrUnitName & "病人记帐单"
        End If
    End If
    txt应收.Text = gstrDec: txt实收.Text = gstrDec
    
    
    arrBaby = Array("0-病人本人", "1-第1个婴儿", "2-第2个婴儿", "3-第3个婴儿", "4-第4个婴儿", "5-第5个婴儿")
    For i = 0 To UBound(arrBaby)
        cboBaby.AddItem arrBaby(i)
    Next
    cboBaby.ListIndex = 0
    
    Select Case mbytInState
        Case 0 '执行
            Call SetShowCol
        Case 1 '查阅
            fraTitle.Enabled = False
            fraInfo.Enabled = False
            fraUnit.Enabled = False
            fraAppend.Enabled = False
            Bill.Active = False
            cmdOK.Visible = False
            cmdCancel.Caption = "退出(&X)"
        Case 3 '销帐
            fraTitle.Enabled = False
            fraInfo.Enabled = False
            fraUnit.Enabled = False
            fraAppend.Enabled = False
            
            '暂时不支持部份删除
            If mint记录性质 <> 1 And False Then
                Call ShowDeleteCol(True)
                Bill.Active = True
            Else
                Bill.Active = False
            End If
    End Select
End Sub

Private Sub SetDisible(Optional bln As Boolean = False)
'界面设置为不可修改状态
    cboNO.Locked = Not bln
    txt姓名.Locked = Not bln
    cbo开单科室.Locked = Not bln
    cbo开单人.Locked = Not bln
    
    chk加班.Enabled = bln
    cboBaby.Enabled = bln
    txtDate.Enabled = bln
    Bill.Active = bln
End Sub

Private Function GetPatient(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
'功能：获取病人信息
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    mintWarn = -1: mstrWarn = ""
    Set mrsWarn = New ADODB.Recordset
    
    txt姓名.ForeColor = Me.ForeColor
    Set mrsInfo = New ADODB.Recordset
    
    If mint病人来源 = 2 Then '对住院病人是否具有强制记帐权限
        If InStr(mstrPrivs, "出院未结强制记帐") > 0 And InStr(mstrPrivs, "出院结清强制记帐") > 0 Then
            strSQL = ""
        ElseIf InStr(mstrPrivs, "出院未结强制记帐") > 0 Then
            strSQL = " And ((B.出院日期 is NULL And Nvl(B.状态,0)<>3) Or Nvl(X.费用余额,0)<>0)"
        ElseIf InStr(mstrPrivs, "出院结清强制记帐") > 0 Then
            strSQL = " And ((B.出院日期 is NULL And Nvl(B.状态,0)<>3) Or Nvl(X.费用余额,0)=0)"
        Else
            strSQL = " And B.出院日期 is NULL And Nvl(B.状态,0)<>3"
        End If
    End If
    
    '字段中使用参数时，如果不明确类型(如Null值),则结果为adVarChar类型
    strSQL = "Select" & _
        " A.病人ID,B.主页ID,To_Number(Nvl(B.当前病区ID,[3])) as 病区ID," & _
        " Nvl(B.出院科室ID,[3]) as 科室ID,B.入院日期,B.出院日期," & _
        " A.门诊号,A.住院号,B.出院病床 as 床号,A.姓名,A.性别,A.年龄,Nvl(B.费别,A.费别) as 费别," & _
        " A.担保人," & IIF(mint病人来源 = 2 And mint记录性质 = 2, "Decode(A.担保额,null,A.担保额,Zl_Patientsurety(A.病人ID,B.主页ID)) 担保额,", "A.担保额,") & _
        " Nvl(B.医疗付款方式,A.医疗付款方式) as 医疗付款方式,Y.编码 as 付款码," & _
        " zl_PatiDayCharge(A.病人ID) as 当日额,Nvl(B.险类,A.险类) as 险类,Nvl(B.病人性质,0) as 病人性质" & _
        " From 病人信息 A,病案主页 B,病人余额 X,医疗付款方式 Y" & _
        " Where A.病人ID=B.病人ID(+) And A.病人ID=X.病人ID(+)" & strSQL & _
        " And A.病人ID=[1] And B.主页ID(+)=[2] And A.医疗付款方式=Y.名称(+)"
        
    On Error GoTo errH
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID, mlng病人科室id)
    If Not mrsInfo.EOF Then
        If Not IsNull(mrsInfo!险类) Then
            txt姓名.ForeColor = vbRed
        End If
        cboBaby.ListIndex = 0
        cboBaby.Enabled = DeptIsWoman(Val("" & mrsInfo!科室ID))
        
        '除了门诊划价以外要处理的内容
        If mint记录性质 = 2 Then
            '刷新病人费用状况
            Set rsTmp = GetMoneyInfo(mrsInfo!病人ID, mcurModiMoney, True)
            If Not rsTmp Is Nothing Then
                cmdOK.Tag = rsTmp!预交余额
                cmdCancel.Tag = rsTmp!费用余额
                txt实收.Tag = rsTmp!预交余额 - rsTmp!费用余额
            Else
                cmdOK.Tag = 0: cmdCancel.Tag = 0: txt实收.Tag = 0
            End If
            sta.Panels(3).Text = "预交:" & Format(Val(cmdOK.Tag), "0.00")
            sta.Panels(3).Text = sta.Panels(3).Text & "/费用:" & Format(Val(cmdCancel.Tag), gstrDec)
            sta.Panels(3).Text = sta.Panels(3).Text & "/剩余:" & Format(Val(txt实收.Tag), "0.00")
            
            '刷新报警信息
            strSQL = "Select Nvl(适用病人,1) as 适用病人,Nvl(报警方法,1) as 报警方法," & _
                " 报警值,报警标志1,报警标志2,报警标志3 From 记帐报警线" & _
                " Where " & IIF(mint病人来源 = 1, "病区ID is NULL", "病区ID=[1]")
            Set mrsWarn = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Nvl(mrsInfo!病区ID, 0)))
        End If
                            
        '住院记帐才处理的内容
        If mint病人来源 = 2 Then
            '急诊费用
            If Not IsNull(mrsInfo!险类) Then
                chk急诊.Value = 0: chk急诊.Visible = True
            Else
                chk急诊.Value = 0: chk急诊.Visible = False
            End If
            
            '发生时间
            If Not IsNull(mrsInfo!出院日期) Then
                txtDate.Text = Format(mrsInfo!出院日期, "yyyy-MM-dd HH:mm:ss")
            Else
                txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
            End If
        End If
        
        '显示病人信息
        txt姓名.Text = Nvl(mrsInfo!姓名)
        txt性别.Text = Nvl(mrsInfo!性别)
        txt年龄.Text = Nvl(mrsInfo!年龄)
        txt费别.Text = Nvl(mrsInfo!费别)
        txt付款方式.Text = Nvl(mrsInfo!医疗付款方式)
        txt付款方式.Tag = Nvl(mrsInfo!付款码, 0) '不要填写为空
        txt床号.Text = Nvl(mrsInfo!床号)
        txt担保人.Text = Nvl(mrsInfo!担保人)
        txt担保额.Text = Format(Nvl(mrsInfo!担保额), "0.00")
        
        With mobjBill
            .病人ID = Nvl(mrsInfo!病人ID, 0)
            .主页ID = Nvl(mrsInfo!主页ID, 0)
            .病区ID = Nvl(mrsInfo!病区ID, 0)
            .科室ID = Nvl(mrsInfo!科室ID, 0)
            .床号 = Nvl(mrsInfo!床号)
            .标识号 = IIF(mint病人来源 = 1, Nvl(mrsInfo!门诊号, 0), Nvl(mrsInfo!住院号, 0))
            .姓名 = Nvl(mrsInfo!姓名)
            .性别 = Nvl(mrsInfo!性别)
            .年龄 = Nvl(mrsInfo!年龄)
            .费别 = Nvl(mrsInfo!费别)
        End With
        
        '在第一次进入时读取病人审批费用项目信息
        If Not Visible And mint病人来源 = 2 And mint记录性质 = 2 Then Set mrsMedAudit = GetAuditRecord(mrsInfo!病人ID, mrsInfo!主页ID)
        
        GetPatient = True
    Else
        Set mrsMedAudit = Nothing
        Set mrsInfo = New ADODB.Recordset
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub CalcMoneys(Optional lngRow As Long = 0)
'功能：计算或重新计算指定行或所有行的金额
'参数：lngRow=指定行,为0表示计算所有行
'说明：ExpenseBill集合的索引对应单据的行号
    Dim i As Long
    Dim strMainRows As String
    Dim bln从项汇总折扣 As Boolean
    
    If mobjBill.Details.Count = 0 Then Exit Sub
    
    For i = IIF(lngRow = 0, 1, lngRow) To IIF(lngRow = 0, mobjBill.Details.Count, lngRow)
        
        bln从项汇总折扣 = False
        If gbln从项汇总折扣 And Not mbln费用登记 Then                    '如果主项屏蔽费别,则汇总计算折扣参数无效,不汇总计算
            If mobjBill.Details(i).从属父号 > 0 Then    '从项
                bln从项汇总折扣 = Not mobjBill.Details(mobjBill.Details(i).从属父号).Detail.屏蔽费别
                If bln从项汇总折扣 And lngRow <> 0 Then strMainRows = "," & mobjBill.Details(i).从属父号      '单独计算一行的时候
            Else
                If ItemHaveSub(i) Then                          '主项或独立项
                     bln从项汇总折扣 = Not mobjBill.Details(i).Detail.屏蔽费别
                     If bln从项汇总折扣 Then strMainRows = strMainRows & "," & i  '一页可能有多个主从项,先记录主项行号,后面再重算主项折扣
                End If
            End If
        End If
                    
        Call CalcMoney(i, bln从项汇总折扣)
    Next
    
    '重算所有主项,不能用bln从项汇总折扣变量,因为可能在遇到不是从项的行时已改变
    If gbln从项汇总折扣 And Not mbln费用登记 Then
        For i = 1 To UBound(Split(strMainRows, ","))
            Call Calc重算主项实收(Split(strMainRows, ",")(i))
        Next
    End If
End Sub

Private Sub CalcMoney(lngRow As Long, Optional bln从项汇总折扣 As Boolean)
'功能：计算或重新计算指定行的金额
'参数：lngRow=指定行
'说明：1.ExpenseBill集合的索引对应单据的行号
'      2.变价只能对应一个收入项目:mobjBill.Details(lngRow).InComes(1)
'      3.如果变价细目未计算出收入项目(第一次计算),则使用默认现价
'      4.如果变价细目已经计算出收入项目(按第2步),并手动更改(也可能未改)了单价,则按该单价计算。
    Dim rsTmp As New ADODB.Recordset
    Dim strInfo As String, i As Long
    Dim dblMoney As Double '用户输入的变价金额
    Dim dbl加班加价率 As Double
    
    Dim rsPrice As New ADODB.Recordset '用于计算时价
    Dim dblAllTime As Double, dblCurTime As Double
    Dim dblPrice As Double
    
    On Error GoTo errH
    
    If InStr(",5,6,7,", mobjBill.Details(lngRow).收费类别) > 0 Then
        Call AdjustCpt(mobjBill.Details(lngRow).收费细目ID)
    End If
    
    gstrSQL = _
        " Select B.收入项目ID,C.名称,C.收据费目,B.现价,B.原价,B.加班加价率,B.附术收费率 " & _
        " From 收费项目目录 A,收费价目 B,收入项目 C " & _
        " Where B.收费细目ID = A.ID And C.ID = B.收入项目ID " & _
        " And ((Sysdate Between B.执行日期 and B.终止日期) Or (Sysdate>=B.执行日期 And B.终止日期 is NULL)) " & _
        " And A.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mobjBill.Details(lngRow).收费细目ID)
    If Not rsTmp.EOF Then
        With mobjBill.Details(lngRow)
            '先获取操作员以前输入的变价金额
            If .Detail.变价 Then
                If InStr(",5,6,7,", .收费类别) > 0 Or (.收费类别 = "4" And .Detail.跟踪在用) Then
                    '计算药品时价(分批或不分批)
                    '必然有记录(输入该项目时已判断)
                    dblAllTime = .付数 * .数次
                    If mbln药房单位 And InStr(",5,6,7,", .收费类别) > 0 Then
                        dblAllTime = dblAllTime * .Detail.药房包装 '库存时价按售价数量进行计算
                    End If
                    If dblAllTime <> 0 Then
                        '药房不分批药品不管效期(这里的库房一定是药房)
                        gstrSQL = "Select Nvl(批次,0) as 批次,Nvl(可用数量,0) as 库存," & _
                            " Nvl(Decode(Nvl(实际数量,0),0,0,实际金额/实际数量),0) as 时价 From 药品库存" & _
                            " Where 库房ID=[1] And 药品ID=[2] And 性质=1 And Nvl(可用数量,0)>0" & _
                            " And (Nvl(批次,0)=0 Or 效期 is NULL Or 效期>Trunc(Sysdate))" & _
                            " Order by Nvl(批次,0)"
                        Set rsPrice = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, .执行部门ID, .收费细目ID)
                        
                        '时价=总金额/总数量
                        dblPrice = 0 '本笔总应收金额
                        For i = 1 To rsPrice.RecordCount
                            If dblAllTime = 0 Then Exit For
                            '取小者
                            If dblAllTime <= rsPrice!库存 Then
                                dblCurTime = dblAllTime
                            Else
                                dblCurTime = rsPrice!库存
                            End If
                            dblPrice = dblPrice + Format(dblCurTime * Format(rsPrice!时价, "0.00000"), gstrDec)
                            dblAllTime = Val(dblAllTime) - Val(dblCurTime)
                            rsPrice.MoveNext
                        Next
                        If dblAllTime <> 0 Then
                            '数量未分解完毕
                            If InStr(",5,6,7,", .收费类别) > 0 Then
                                MsgBox "第 " & lngRow & " 行时价药品""" & .Detail.名称 & """库存不足,无法计算价格！", vbInformation, gstrSysName
                            Else
                                MsgBox "第 " & lngRow & " 行时价卫生材料""" & .Detail.名称 & """库存不足,无法计算价格！", vbInformation, gstrSysName
                            End If
                            dblMoney = 0
                        Else
                            '注意：货币型最多只能保留4位小数,且不四舍五入,所以需要手工舍入;而用其它型在计算精度上又有问题
                            dblAllTime = .付数 * .数次
                            If mbln药房单位 And InStr(",5,6,7,", .收费类别) > 0 Then
                                dblAllTime = dblAllTime * .Detail.药房包装 '按售价数量计算实价
                            End If
                            dblMoney = Format(dblPrice / dblAllTime, "0.00000") '这里结果是按售价单位
                        End If
                    Else
                        dblMoney = 0
                    End If
                Else
                    If .InComes.Count = 0 Then
                        '如果第一次计算金额,变价默认取原价
                        dblMoney = 0 'IIf(IsNull(rsTmp!原价), 0, rsTmp!原价)
                    Else
                        dblMoney = .InComes(1).标准单价
                        '如果用户输入的变价不满足变价范围，则取默认值
                        If Abs(dblMoney) > Abs(IIF(IsNull(rsTmp!现价), 0, rsTmp!现价)) Then
                            dblMoney = IIF(IsNull(rsTmp!原价), 0, rsTmp!原价)
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
                .收据费目 = IIF(IsNull(rsTmp!收据费目), "", rsTmp!收据费目)
                .原价 = IIF(IsNull(rsTmp!原价), 0, rsTmp!原价)
                .现价 = IIF(IsNull(rsTmp!现价), 0, rsTmp!现价)
                If mobjBill.Details(lngRow).Detail.变价 Then
                    If InStr(",5,6,7,", mobjBill.Details(lngRow).收费类别) > 0 And mbln药房单位 Then
                        .标准单价 = Format(dblMoney * mobjBill.Details(lngRow).Detail.药房包装, "0.00000")
                    Else
                        .标准单价 = Format(dblMoney, "0.00000")
                    End If
                Else
                    If InStr(",5,6,7,", mobjBill.Details(lngRow).收费类别) > 0 And mbln药房单位 Then
                        .标准单价 = Format(Nvl(rsTmp!现价, 0) * mobjBill.Details(lngRow).Detail.药房包装, "0.00000")
                    Else
                        .标准单价 = Format(Nvl(rsTmp!现价, 0), "0.00000")
                    End If
                End If
                '应收金额=单价 * 付数 * 数次
                If mobjBill.Details(lngRow).Detail.变价 And (InStr(",5,6,7,", mobjBill.Details(lngRow).收费类别) > 0 _
                        Or mobjBill.Details(lngRow).收费类别 = "4" And mobjBill.Details(lngRow).Detail.跟踪在用) Then
                    .应收金额 = dblPrice '保证应收金额与零售金额没有误差
                Else
                    .应收金额 = .标准单价 * mobjBill.Details(lngRow).付数 * mobjBill.Details(lngRow).数次
                End If
                
                '附加手术费率用计算(所有收入项目)
                If mobjBill.Details(lngRow).附加标志 = 1 And mobjBill.Details(lngRow).收费类别 = "F" Then
                    .应收金额 = .应收金额 * IIF(IsNull(rsTmp!附术收费率), 1, rsTmp!附术收费率 / 100)
                End If
                '加班费用率计算
                dbl加班加价率 = 0
                If mobjBill.加班标志 = 1 And mobjBill.Details(lngRow).Detail.加班加价 Then
                    dbl加班加价率 = Nvl(rsTmp!加班加价率, 0) / 100
                    .应收金额 = .应收金额 * (1 + dbl加班加价率)
                End If
                
                .应收金额 = CCur(Format(.应收金额, gstrDec))
                
                If mbln费用登记 Then
                    .实收金额 = 0
                Else
                    If mobjBill.Details(lngRow).Detail.屏蔽费别 Or bln从项汇总折扣 Then
                        .实收金额 = .应收金额
                    Else
                        .实收金额 = CCur(Format(ActualMoney(mobjBill.费别, .收入项目ID, .应收金额, _
                            mobjBill.Details(lngRow).收费细目ID, mobjBill.Details(lngRow).执行部门ID, _
                            mobjBill.Details(lngRow).付数 * mobjBill.Details(lngRow).数次, dbl加班加价率), gstrDec))
                    End If
                End If
                
                '获取项目保险信息,医保病人才处理,不需要连接医保
                If Not IsNull(mrsInfo!险类) Then
                    strInfo = gclsInsure.GetItemInsure(mobjBill.病人ID, mobjBill.Details(lngRow).收费细目ID, .实收金额, False, mrsInfo!险类)
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
                
                '实收金额存入Key中,以处理分币问题(即Key中存放原始实收金额,不变)
                mobjBill.Details(lngRow).InComes.Add .收入项目ID, .收入项目, .收据费目, .标准单价, .应收金额, .实收金额, .原价, .现价, "_" & .实收金额, .统筹金额
            End With
            rsTmp.MoveNext
        Next
    Else
        '如果没有收入项目,则清除对应的程序对象
        Set mobjBill.Details(lngRow).InComes = New BillInComes
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ShowDetails(Optional lngRow As Long = 0)
'功能：刷新显示指定行或所有行的内容
'参数：lngRow=指定行,为0表示显示所有行
'说明：ExpenseBill集合的索引对应单据的行号
    Dim i As Long, curTotal As Currency
    
    Bill.Redraw = False
    If lngRow = 0 Then
        For i = 1 To mobjBill.Details.Count
            ShowDetail i
        Next
    Else
        ShowDetail lngRow
    End If
    Bill.Redraw = True
    
    curTotal = GetBillTotal(mobjBill)
    
    If IsNumeric(cmdOK.Tag) Then
        sta.Panels(3).Text = "预交:" & Format(Val(cmdOK.Tag), "0.00")
        sta.Panels(3).Text = sta.Panels(3).Text & "/费用:" & Format(Val(cmdCancel.Tag) + curTotal, gstrDec)
        sta.Panels(3).Text = sta.Panels(3).Text & "/剩余:" & Format(Val(txt实收.Tag) - curTotal, "0.00")
    End If
End Sub

Private Sub ShowDetail(lngRow As Long)
'功能：刷新显示指定行的内容
'参数：lngRow=指定行
'说明：ExpenseBill集合的索引对应单据的行号
    Dim dbl单价 As Double, cur金额 As Currency
    Dim i As Long, j As Long
    
    If lngRow > Bill.Rows - 1 Then Exit Sub
    
    '清除单据行
    For i = 1 To Bill.Cols - 1
        '输入时收费类别不清除
        If Not (i = 1 And Bill.TextMatrix(lngRow, i) <> "") Then Bill.TextMatrix(lngRow, i) = ""
    Next
    
    If mobjBill.Details(lngRow).收费类别 <> "" Then
        Bill.RowData(lngRow) = Asc(mobjBill.Details(lngRow).收费类别)
    End If
    
    '刷新单据行
    For i = 1 To Bill.Cols - 1
        Select Case Bill.TextMatrix(0, i)
            Case "类别"
                '浏览单据或从属项目只(能)显示名称
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.类别名称
            Case "项目"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.名称
            Case "规格"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.规格
            Case "单位"
                If InStr(",5,6,7,", mobjBill.Details(lngRow).收费类别) > 0 And mbln药房单位 Then
                    Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.药房单位
                Else
                    Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.计算单位
                End If
            Case "付数"
                Bill.TextMatrix(lngRow, i) = IIF(mobjBill.Details(lngRow).付数 = 0, 1, mobjBill.Details(lngRow).付数)
            Case "数次"
                '数次在第一次显示时已默认设置为1
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).数次
            Case "单价"
                '单价是该收费细目所有收入项目的合计
                '第一次计算时是在默认数次为1的基础上计算出来的
                dbl单价 = 0
                If mobjBill.Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Details(lngRow).InComes.Count
                        dbl单价 = dbl单价 + mobjBill.Details(lngRow).InComes(j).标准单价
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(dbl单价, "0.00000")
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
                    If mbytInState = 0 Then
                        mrsUnit.Filter = "ID=" & mobjBill.Details(lngRow).执行部门ID
                        If mrsUnit.RecordCount <> 0 Then
                            Bill.TextMatrix(lngRow, i) = mrsUnit!编码 & "-" & mrsUnit!名称
                        Else
                            Bill.TextMatrix(lngRow, i) = Get部门名称(mobjBill.Details(lngRow).执行部门ID, mrsUnit)
                        End If
                    Else
                        '浏览单据只(能)显示名称
                        Bill.TextMatrix(lngRow, i) = Get部门名称(mobjBill.Details(lngRow).执行部门ID, mrsUnit)
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

Public Sub ShowMoney()
'功能：刷新显示收入项目费用区
    Dim blnExist As Boolean
    Dim cur实收 As Currency, cur应收 As Currency
    Dim i As Long, j As Long, k As Long
    
    mshMoney.Redraw = False
    
    '产生汇总费目
    Set mcolMoneys = New BillInComes
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
                mcolMoneys(k).应收金额 = mcolMoneys(k).应收金额 + mobjBill.Details(i).InComes(j).应收金额
            Else
                With mobjBill.Details(i).InComes(j)
                    mcolMoneys.Add .收入项目ID, .收入项目, .收据费目, .标准单价, .应收金额, .实收金额
                End With
            End If
        Next
    Next
    
    '刷新显示
    If mcolMoneys.Count > 0 Then
        mshMoney.Rows = mcolMoneys.Count + 1
    End If
    If mshMoney.Rows < 5 Then mshMoney.Rows = 5

    Call SetMoneyList
    
    For i = 1 To mcolMoneys.Count
        mshMoney.TextMatrix(i, 0) = mcolMoneys(i).收入项目
        mshMoney.TextMatrix(i, 1) = Format(mcolMoneys(i).实收金额, gstrDec)
        cur实收 = cur实收 + mcolMoneys(i).实收金额
        cur应收 = cur应收 + mcolMoneys(i).应收金额
    Next
    
    txt应收.Text = Format(cur应收, gstrDec)
    txt实收.Text = Format(cur实收, gstrDec)
    
    mshMoney.TopRow = mshMoney.Rows - 1
    mshMoney.Redraw = True
End Sub

Private Function GetCur应收() As Currency
'功能：获取病人当前单据合计金额(收费病人累加单据时用)
    Dim i As Long
    For i = 1 To mcolMoneys.Count
        GetCur应收 = GetCur应收 + mcolMoneys(i).应收金额
    Next
End Function

Private Function GetInputDetail(ByVal lng项目ID As Long) As Detail
'功能：读取收费项目信息
    Dim objDetail As New Detail
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long, lngMediCareNO As Long
        
    If mrsInfo.State = 1 Then lngMediCareNO = Val("" & mrsInfo!险类)
    
    If lngMediCareNO > 0 Then
        strSQL = _
            " Select" & _
            " A.ID,A.类别,B.名称 as 类别名称,A.编码,Nvl(E.名称,A.名称) as 名称,A.规格,A.计算单位," & _
            " A.屏蔽费别,A.是否变价,A.加班加价,A.执行科室,A.服务对象,A.费用类型,A.补充摘要,F.要求审批," & _
            " Decode(A.类别,'4',D.诊疗ID,C.药名ID) as 药名ID," & _
            " Decode(A.类别,'4',D.在用分批,C.药房分批) as 分批," & _
            " Decode(A.类别,'4',1,C." & mstr药房包装 & ") as 药房包装," & _
            " Decode(A.类别,'4',A.计算单位,C." & mstr药房单位 & ") as 药房单位,D.跟踪在用,A.录入限量" & _
            " From 收费项目目录 A,收费项目类别 B,药品规格 C,材料特性 D,收费项目别名 E,保险支付项目 F" & _
            " Where A.ID=C.药品ID(+) And A.ID=D.材料ID(+) And B.编码=A.类别" & _
            " And A.ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=[2] And A.ID=[1] And A.ID=F.收费细目ID(+) And F.险类(+)=[3]"

    Else
        strSQL = _
            " Select" & _
            " A.ID,A.类别,B.名称 as 类别名称,A.编码,Nvl(E.名称,A.名称) as 名称,A.规格,A.计算单位," & _
            " A.屏蔽费别,A.是否变价,A.加班加价,A.执行科室,A.服务对象,A.费用类型,A.补充摘要,0 as 要求审批," & _
            " Decode(A.类别,'4',D.诊疗ID,C.药名ID) as 药名ID," & _
            " Decode(A.类别,'4',D.在用分批,C.药房分批) as 分批," & _
            " Decode(A.类别,'4',1,C." & mstr药房包装 & ") as 药房包装," & _
            " Decode(A.类别,'4',A.计算单位,C." & mstr药房单位 & ") as 药房单位,D.跟踪在用,A.录入限量" & _
            " From 收费项目目录 A,收费项目类别 B,药品规格 C,材料特性 D,收费项目别名 E" & _
            " Where A.ID=C.药品ID(+) And A.ID=D.材料ID(+) And B.编码=A.类别" & _
            " And A.ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=[2] And A.ID=[1]"
    End If
        
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng项目ID, IIF(gbln商品名, 3, 1), lngMediCareNO)
    With objDetail
        .ID = rsTmp!ID
        .药名ID = Nvl(rsTmp!药名ID, 0)
        .编码 = rsTmp!编码
        .规格 = Nvl(rsTmp!规格)
        .药房包装 = Nvl(rsTmp!药房包装, 1)
        .药房单位 = Nvl(rsTmp!药房单位)
        .分批 = Nvl(rsTmp!分批, 0) = 1
        .变价 = Nvl(rsTmp!是否变价, 0) = 1
        .计算单位 = Nvl(rsTmp!计算单位)
        .加班加价 = Nvl(rsTmp!加班加价, 0) = 1
        .类别 = rsTmp!类别
        .类别名称 = rsTmp!类别名称
        .名称 = rsTmp!名称
        .屏蔽费别 = Nvl(rsTmp!屏蔽费别, 0) = 1
        .执行科室 = Nvl(rsTmp!执行科室, 0)
        .服务对象 = Nvl(rsTmp!服务对象, 0)
        .类型 = Nvl(rsTmp!费用类型)
        .补充摘要 = Nvl(rsTmp!补充摘要, 0) = 1
        .跟踪在用 = Nvl(rsTmp!跟踪在用, 0) = 1
        .要求审批 = Nvl(rsTmp!要求审批, 0) = 1
        .录入限量 = Val("" & rsTmp!录入限量)
    End With
    Set GetInputDetail = objDetail
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetDetail(Detail As Detail, lngRow As Long, lngDoUnit As Long, Optional bytParent As Byte = 0)
'功能：根据指定的收费细目对象设定单据指点定行的收费细目(新增的或修改)
'说明：
'      1.用于新输入或更改收费细目行！！！
'      2.当bytParent<>0时,则为设置从属项目,从属项目一定是新增行,且主项目一定存在

    Dim tmpIncomes As New BillInComes
    Dim intPay As Integer, i As Long, dblTime As Double
    
    '取其它中药的付数
    intPay = GetPay(lngRow)
    If Detail.类别 <> "7" Then intPay = 1
    
    If mobjBill.Details.Count < lngRow Then
        '如果该行对应的程序对象尚未初始,则加入
        With Detail
            '序号=行号,父号=0
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
                    dblTime = IIF(Detail.从项数次 = 0, 1, Detail.从项数次)
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
            mobjBill.Details.Add tmpIncomes, Detail, .ID, CByte(lngRow), CInt(bytParent), .类别, .计算单位, intPay, dblTime, 0, lngDoUnit, ""
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
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long, blnExist As Boolean
    Dim strSQL As String
    
    strSQL = "Select count(从项ID) as NUM from 收费从属项目 where 主项ID=" & mobjBill.Details(lngRow).收费细目ID
    rsTmp.CursorLocation = adUseClient
    Call SQLTest(App.ProductName, Me.Caption, strSQL)
    rsTmp.Open strSQL, gcnOracle, adOpenKeyset
    Call SQLTest
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
End Function

Private Function GetSubDetails(ByVal lng项目ID As Long) As Details
'功能：返回一个收费细目的从属项目集
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objDetail As New Detail, lngMediCareNO As Long
            
    Set GetSubDetails = New Details
    
    If mrsInfo.State = 1 Then lngMediCareNO = Val("" & mrsInfo!险类)
    If lngMediCareNO > 0 Then
        strSQL = _
            "Select" & _
            " A.ID,Decode(A.类别,'4',E.诊疗ID,D.药名ID) as 药名ID,A.类别,B.名称 as 类别名称," & _
            " A.费用类型,A.编码,Nvl(F.名称,A.名称) as 名称,A.规格,A.计算单位,A.屏蔽费别," & _
            " Decode(A.类别,'4',E.在用分批,D.药房分批) as 分批,A.是否变价," & _
            " Decode(A.类别,'4',1,D." & mstr药房包装 & ") as 药房包装,A.服务对象," & _
            " Decode(A.类别,'4',A.计算单位,D." & mstr药房单位 & ") as 药房单位," & _
            " A.加班加价,A.执行科室,C.固有从属,C.从项数次,E.跟踪在用,G.要求审批" & _
            " From 收费项目目录 A,收费项目类别 B,收费从属项目 C,药品规格 D,材料特性 E,收费项目别名 F,保险支付项目 G" & _
            " Where B.编码=A.类别 And C.从项ID=A.ID And A.ID=D.药品ID(+) And A.ID=E.材料ID(+)" & _
            " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
            " And A.ID=F.收费细目ID(+) And F.码类(+)=1 And F.性质(+)=[2] And C.主项ID=[1] And A.ID=G.收费细目ID(+) And G.险类(+)=[3]"
    Else
        strSQL = _
            "Select" & _
            " A.ID,Decode(A.类别,'4',E.诊疗ID,D.药名ID) as 药名ID,A.类别,B.名称 as 类别名称," & _
            " A.费用类型,A.编码,Nvl(F.名称,A.名称) as 名称,A.规格,A.计算单位,A.屏蔽费别," & _
            " Decode(A.类别,'4',E.在用分批,D.药房分批) as 分批,A.是否变价," & _
            " Decode(A.类别,'4',1,D." & mstr药房包装 & ") as 药房包装,A.服务对象," & _
            " Decode(A.类别,'4',A.计算单位,D." & mstr药房单位 & ") as 药房单位," & _
            " A.加班加价,A.执行科室,C.固有从属,C.从项数次,E.跟踪在用,0 as 要求审批" & _
            " From 收费项目目录 A,收费项目类别 B,收费从属项目 C,药品规格 D,材料特性 E,收费项目别名 F" & _
            " Where B.编码=A.类别 And C.从项ID=A.ID And A.ID=D.药品ID(+) And A.ID=E.材料ID(+)" & _
            " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
            " And A.ID=F.收费细目ID(+) And F.码类(+)=1 And F.性质(+)=[2] And C.主项ID=[1]"
    End If

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng项目ID, IIF(gbln商品名, 3, 1), lngMediCareNO)
    For i = 1 To rsTmp.RecordCount
        Set objDetail = New Detail
        With objDetail
                .ID = rsTmp!ID
                .药名ID = Nvl(rsTmp!药名ID, 0)
                .编码 = rsTmp!编码
                .变价 = Nvl(rsTmp!是否变价, 0) = 1
                .规格 = Nvl(rsTmp!规格)
                .药房包装 = Nvl(rsTmp!药房包装, 1)
                .药房单位 = Nvl(rsTmp!药房单位)
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
                GetSubDetails.Add .ID, .药名ID, .类别, .类别名称, .名称, .编码, .简码, .别名, .规格, .计算单位, .说明, .屏蔽费别, _
                    .药房包装, .药房单位, .分批, .变价, .加班加价, .执行科室, .服务对象, .类型, .补充摘要, .固有从属, .从项数次, .跟踪在用, , , , , , .要求审批
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
        For i = 1 To Bill.Cols - 1
            Bill.TextMatrix(lngRow, i) = ""
            Bill.RowData(lngRow) = 0
        Next
    Else
        Bill.RemoveMSFItem lngRow
    End If
End Sub

Private Sub NewBill()
    Set mobjBill = New ExpenseBill
    
    mcurModiMoney = 0
    mlngPreRow = 0
    cboNO.Text = ""
    chk加班.Value = IIF(OverTime, 1, 0)
    txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
            
    '婴儿费处理
    Call cbo开单科室_Click
    With mobjBill
        .门诊标志 = mint病人来源
        .开单人 = NeedName(cbo开单人.Text)
        If cbo开单科室.ListIndex = -1 Then
            .开单部门ID = 0
        Else
            .开单部门ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
        End If
        .发生时间 = CDate(txtDate.Text)
        .加班标志 = chk加班.Value
        .划价人 = UserInfo.姓名
        .操作员编号 = UserInfo.编号
        .操作员姓名 = UserInfo.姓名
    End With
End Sub

Private Sub ClearMoney()
'功能：清除费用显示区
    Dim i As Long, j As Long
    mshMoney.Redraw = False
    For i = 1 To mshMoney.Rows - 1
        For j = 0 To mshMoney.Cols - 1
            mshMoney.TextMatrix(i, j) = ""
        Next
    Next
    mshMoney.Rows = 5
    mshMoney.Redraw = True
End Sub

Private Function SaveBill() As Boolean
'功能:保存当前输入的记帐单据(适用住院记帐、划价、或对两者的修改)
'入口:mobjBill=单据对象
'出口:保存是否成功
    Dim int行号 As Integer, int序号 As Integer, int价格父号 As Integer
    Dim dbl数次 As Double, dbl单价 As Double
    Dim intInsure As Integer, strNO As String, strTmp As String
    Dim arrSQL As Variant, i As Long, j As Long
    Dim int划价 As Integer, bln上传 As Boolean
    Dim strSQL As String, strStuffDept As String '记录卫料发料部门
    
    If mint记录性质 = 1 Then
        mobjBill.NO = zlDatabase.GetNextNo(13)
    Else
        mobjBill.NO = zlDatabase.GetNextNo(14)
    End If
    mobjBill.发生时间 = CDate(txtDate.Text)
    mobjBill.登记时间 = zlDatabase.Currentdate
    
    int序号 = 0
    arrSQL = Array()

    For Each mobjBillDetail In mobjBill.Details
        If mobjBillDetail.数次 <> 0 Then
            For Each mobjBillIncome In mobjBillDetail.InComes
                int序号 = int序号 + 1 '当前记录序号
                
                '单据主体
                With mobjBill
                    If mint病人来源 = 2 Then
                        gstrSQL = "zl_住院记帐记录_INSERT('" & .NO & "'," & int序号 & "," & .病人ID & "," & ZVal(.主页ID) & "," & _
                            ZVal(.标识号) & "," & "'" & .姓名 & "','" & .性别 & "','" & .年龄 & "','" & .床号 & "','" & .费别 & "'," & _
                            ZVal(.病区ID) & "," & ZVal(.科室ID) & "," & .加班标志 & "," & .婴儿费 & "," & .开单部门ID & ",'" & .开单人 & "',"
                    Else
                        If mint记录性质 = 2 Then
                            gstrSQL = "zl_门诊记帐记录_INSERT('" & .NO & "'," & int序号 & "," & .病人ID & "," & _
                                ZVal(.标识号) & "," & "'" & .姓名 & "','" & .性别 & "','" & .年龄 & "'," & _
                                "'" & .费别 & "'," & .加班标志 & "," & .婴儿费 & "," & _
                                ZVal(.病区ID) & "," & ZVal(.科室ID) & "," & .开单部门ID & ",'" & .开单人 & "',"
                        Else
                            gstrSQL = "zl_门诊划价记录_Insert('" & .NO & "'," & int序号 & "," & .病人ID & "," & ZVal(.主页ID) & "," & _
                                ZVal(.标识号) & ",'" & IIF(Val(txt付款方式.Tag) = 0, "", txt付款方式.Tag) & "','" & .姓名 & "'," & _
                                "'" & .性别 & "','" & .年龄 & "','" & .费别 & "'," & .加班标志 & "," & _
                                ZVal(.病区ID) & "," & ZVal(.科室ID) & "," & .开单部门ID & ",'" & .开单人 & "',"
                        End If
                    End If
                End With
                
                '收费细目部份
                With mobjBillDetail
                    '处理从属父号
                    If .序号 <> int行号 Then
                        int行号 = .序号
                        int价格父号 = int序号
                        
                        '重新处理从属父号
                        For i = .序号 + 1 To mobjBill.Details.Count
                            If mobjBill.Details(i).从属父号 = .序号 Then
                                mobjBill.Details(i).从属父号 = int序号
                            End If
                        Next
                    End If
                    gstrSQL = gstrSQL & .从属父号 & "," & .收费细目ID & ",'" & .收费类别 & "','" & .计算单位 & "',"
                    
                    If mint病人来源 = 2 Then
                        gstrSQL = gstrSQL & IIF(.保险项目否, 1, 0) & "," & ZVal(.保险大类ID) & ",'" & .保险编码 & "',"
                    ElseIf mint记录性质 = 1 Then
                        gstrSQL = gstrSQL & "NULL,"
                    End If
                    
                    dbl数次 = .数次
                    If InStr(",5,6,7,", .收费类别) > 0 And mbln药房单位 Then
                        dbl数次 = Format(.数次 * .Detail.药房包装, "0.00000")
                    End If
                    gstrSQL = gstrSQL & IIF(.付数 = 0, 1, .付数) & "," & dbl数次 & "," & .附加标志 & "," & ZVal(.执行部门ID) & ","
                End With
                
                '收入项目部份
                With mobjBillIncome
                    dbl单价 = .标准单价
                    If InStr(",5,6,7,", mobjBillDetail.收费类别) > 0 And mbln药房单位 Then
                        dbl单价 = Format(.标准单价 / mobjBillDetail.Detail.药房包装, "0.00000")
                    End If
                    gstrSQL = gstrSQL & IIF(int价格父号 = int序号, "NULL", int价格父号) & "," & .收入项目ID & "," & _
                        "'" & .收据费目 & "'," & dbl单价 & "," & .应收金额 & "," & .实收金额 & ","
                    If mint病人来源 = 2 Then
                        gstrSQL = gstrSQL & IIF(.统筹金额 = 0, "NULL", .统筹金额) & ","
                    End If
                End With
                                                
                '其它部分
                gstrSQL = gstrSQL & _
                    "To_Date('" & Format(mobjBill.发生时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                    "To_Date('" & Format(mobjBill.登记时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),'" & mstrInNO & "',"
                
                '是否只生成划价单
    '                If mbln费用登记 Then
    '                    int划价 = 0 '零耗费用登记不必生成划价单
    '                Else
    '                    '本来应是诊疗类别,应按执行医嘱的类别来判断
    '                    If InStr(",5,6,7,", mobjBillDetail.收费类别) > 0 Then
    '                        int划价 = IIF(InStr(gstr发送划价单, "5") > 0, 1, 0)
    '                    Else
    '                        int划价 = IIF(InStr(gstr发送划价单, mobjBillDetail.收费类别) > 0, 1, 0)
    '                    End If
    '                End If
                If int划价 = 0 Then bln上传 = True '只要存在不是划价单就要上传
                
                '收集卫料发料部门,以便自动发料,门诊病人仅记帐时(发送为划价时不管),住院病人只有记帐
                With mobjBillDetail
                    If (mint病人来源 = 1 And mint记录性质 = 2 And gbln门诊自动发料 Or mint病人来源 = 2 And gbln住院自动发料) And int划价 = 0 Then
                        If .执行部门ID <> 0 And .收费类别 = "4" And .Detail.跟踪在用 Then
                            If InStr("," & strStuffDept, "," & .执行部门ID & ",") = 0 Then
                                strStuffDept = strStuffDept & "," & .执行部门ID
                            End If
                        End If
                    End If
                End With
                
                If mint病人来源 = 2 Then
                    gstrSQL = gstrSQL & int划价 & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & _
                        "0," & IIF(mobjBillDetail.收费类别 = "4", mlng卫材类别ID, mlng药品类别ID) & "," & _
                        "NULL,'" & mobjBillDetail.摘要 & "'," & chk急诊.Value & "," & ZVal(mlng医嘱ID) & "," & _
                        "Null,Null,Null,Null,Null,Null,'" & mobjBillDetail.Detail.类型 & "')"
                Else
                    If mint记录性质 = 2 Then
                        gstrSQL = gstrSQL & int划价 & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & _
                            IIF(mobjBillDetail.收费类别 = "4", mlng卫材类别ID, mlng药品类别ID) & "," & _
                            "NULL,'" & mobjBillDetail.摘要 & "'," & ZVal(mlng医嘱ID) & ")"
                    Else
                        gstrSQL = gstrSQL & "'" & UserInfo.姓名 & "'," & _
                            IIF(mobjBillDetail.收费类别 = "4", mlng卫材类别ID, mlng药品类别ID) & "," & _
                            "'" & mobjBillDetail.摘要 & "'," & ZVal(mlng医嘱ID) & ")"
                    End If
                End If
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = mobjBillDetail.收费细目ID & ";" & gstrSQL
            Next
        End If
    Next
    
    '-----------------------------------------------------------------------------------------------------------------
    '插入医嘱院加费用
    gstrSQL = "ZL_病人医嘱附费_Insert(" & mlng医嘱ID & "," & mlng发送号 & "," & mint记录性质 & ",'" & mobjBill.NO & "')"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "0;" & gstrSQL
    
    '修改前退除原单据
    If mstrInNO <> "" Then
        '先判断是否医保病人记的帐,并作合法性检查(进入修改时已作了一次相关判断)
        If mint病人来源 = 2 Then
            '去掉了医保连接匹配检查
            intInsure = BillExistInsure(mstrInNO)
        End If
        If mint病人来源 = 2 Then
            gstrSQL = "zl_住院记帐记录_DELETE('" & mstrInNO & "',NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
        Else
            If mint记录性质 = 2 Then
                gstrSQL = "zl_门诊记帐记录_DELETE('" & mstrInNO & "',NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
            Else
                gstrSQL = "zl_门诊划价记录_DELETE('" & mstrInNO & "')"
            End If
        End If
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
        
        '执行SQL语句
        strTmp = ""
        On Error GoTo errH
        gcnOracle.BeginTrans
        
            For i = 0 To UBound(arrSQL)
                Call zlDatabase.ExecuteProcedure(Mid(arrSQL(i), InStr(arrSQL(i), ";") + 1), Me.Caption)
            Next
            
            '执行自动发料
            If strStuffDept <> "" Then
                strStuffDept = Mid(strStuffDept, 2)
                For i = 0 To UBound(Split(strStuffDept, ","))
                    strSQL = "zl_材料收发记录_处方发料(" & Split(strStuffDept, ",")(i) & ",25,'" & mobjBill.NO & "','" & UserInfo.姓名 & "','" & UserInfo.姓名 & "','" & UserInfo.姓名 & "',1,Sysdate)"
                    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                Next
            End If
            
            '医保接口
            '1.医保记帐作废上传
            If mint病人来源 = 2 And mstrInNO <> "" And intInsure <> 0 Then
                If gclsInsure.GetCapability(support记帐作废上传, , intInsure) And Not gclsInsure.GetCapability(support记帐完成后上传, , intInsure) Then
                    If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                        gcnOracle.RollbackTrans: Exit Function
                    End If
                End If
            End If
            
            '2.记帐实时上传
            If mint病人来源 = 2 And bln上传 And Not IsNull(mrsInfo!险类) Then
                '医保传输费用明细
                If gclsInsure.GetCapability(support记帐上传, , mrsInfo!险类) And Not gclsInsure.GetCapability(support记帐完成后上传, , mrsInfo!险类) Then
                    strTmp = ""
                    If Not gclsInsure.TranChargeDetail(2, mobjBill.NO, 2, 1, strTmp, , mrsInfo!险类) Then
                        gcnOracle.RollbackTrans
                        If strTmp <> "" Then MsgBox strTmp, vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            End If
        
        gcnOracle.CommitTrans
        
        '医保接口
        '1.医保记帐作废上传
        If mint病人来源 = 2 And mstrInNO <> "" And intInsure > 0 Then
            If gclsInsure.GetCapability(support记帐作废上传, , intInsure) And gclsInsure.GetCapability(support记帐完成后上传, , intInsure) Then
                If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                    MsgBox "单据""" & mstrInNO & """向医保传送失败,该单据的费用已删除！", vbInformation, gstrSysName
                End If
            End If
        End If
        
        '2.记帐实时上传
        If mint病人来源 = 2 And bln上传 And Not IsNull(mrsInfo!险类) Then
            '医保传输费用明细
            If gclsInsure.GetCapability(support记帐上传, , mrsInfo!险类) And gclsInsure.GetCapability(support记帐完成后上传, , mrsInfo!险类) Then
                strTmp = ""
                If Not gclsInsure.TranChargeDetail(2, mobjBill.NO, 2, 1, strTmp, , mrsInfo!险类) Then
                    If strTmp <> "" Then
                        MsgBox strTmp, vbInformation, gstrSysName
                    Else
                        MsgBox "单据""" & mobjBill.NO & """的数据向医保传送失败,该单据已保存！", vbInformation, gstrSysName
                    End If
                End If
            End If
        End If
        
        '加入单据历史记录(所有类型单据)
        cboNO.AddItem mobjBill.NO, 0
        For i = cboNO.ListCount - 1 To 10 Step -1
            cboNO.RemoveItem i '只显示10个
        Next
        
        '医保接口
        If strTmp <> "" Then MsgBox strTmp, vbInformation, gstrSysName
    End If
    SaveBill = True
    Exit Function
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ReadBill(ByVal strNO As String, Optional blnDelete As Boolean) As Integer
'功能：根据单据号读取一张单据并将其填入表格
'参数：strNO=单据号
'      blnDelete=是否读取要退费的单据
    Dim rsTmp As New ADODB.Recordset
    Dim rsPatiMoney As ADODB.Recordset
    Dim strSQL1 As String, strSQL2 As String, intSign As Integer
    Dim curTotal As Currency, cur应收Total As Currency
    Dim strSQL As String, i As Long
    Dim intInsure As Integer, blnDo As Boolean
    Dim blnNOMoved As Boolean
    
    If mbytInState = 1 Then
        blnNOMoved = MovedByNO(strNO, "病人费用记录", "记录性质=" & mint记录性质 & IIF(mstrTime <> "", " And 登记时间=To_Date('" & mstrTime & "','YYYY-MM-DD HH24:MI:SS')", ""))
    End If
    
    On Error GoTo errH
    
    Call ClearRows: Call Bill.ClearBill
    Call SetColNum: Call ClearMoney
    
    strSQL = _
        " Select A.病人ID,A.主页ID,A.姓名,A.性别,A.年龄,A.费别,A.床号," & _
        " A.病人病区ID,A.开单部门ID,A.加班标志,A.婴儿费,A.开单人,A.划价人,A.操作员姓名," & _
        " A.开单部门ID,C.编码||'-'||C.名称 as 开单部门,A.发生时间," & _
        " B.医疗付款方式,B.担保人,B.担保额,A.是否急诊" & _
        " From 病人费用记录 A,病人信息 B,部门表 C" & _
        " Where Rownum=1 And NO=[1] And A.记录性质=[2]" & _
        " And A.病人ID=B.病人ID And Instr([3],A.记录状态)>0" & _
        IIF(mstrTime <> "", " And A.登记时间=[4]", "") & _
        " And A.开单部门ID=C.ID"
    If blnNOMoved Then
        strSQL = Replace(strSQL, "病人费用记录", "H病人费用记录")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, mint记录性质, _
        IIF(mblnDelete, "2", "0,1,3"), CDate(IIF(mstrTime = "", "1990-01-01", mstrTime)))
    If rsTmp.EOF Then
        MsgBox "没有发现该单据。", vbInformation, gstrSysName
        Exit Function
    End If

    cboNO.Text = strNO
    txt姓名.Text = Nvl(rsTmp!姓名)
    txt性别.Text = Nvl(rsTmp!性别)
    txt年龄.Text = Nvl(rsTmp!年龄)
    If Nvl(rsTmp!主页ID, 0) <> 0 Then
        txt床号.Text = Nvl(rsTmp!床号)
    End If
    txt费别.Text = Nvl(rsTmp!费别)
    txt担保人.Text = Nvl(rsTmp!担保人)
    txt担保额.Text = Format(Nvl(rsTmp!担保额), "0.00")
    txt付款方式.Text = Nvl(rsTmp!医疗付款方式)
    
    cbo开单科室.AddItem Nvl(rsTmp!开单部门)
    cbo开单科室.ItemData(cbo开单科室.NewIndex) = Nvl(rsTmp!开单部门ID, 0)
    cbo开单科室.ListIndex = cbo开单科室.NewIndex
    
    If Nvl(rsTmp!是否急诊, 0) = 1 Then
        chk急诊.Value = 1: chk急诊.Visible = True
    End If
    
    chk加班.Value = Nvl(rsTmp!加班标志, 0)
    cboBaby.ListIndex = IIF(Val("" & rsTmp!婴儿费) > cboBaby.ListCount - 1, 0, Val("" & rsTmp!婴儿费))
    
    '开单人
    Call GetCboIndex(cbo开单人, Nvl(rsTmp!开单人))
    If cbo开单人.ListIndex = -1 And Not IsNull(rsTmp!开单人) Then
        cbo开单人.AddItem rsTmp!开单人
        cbo开单人.ListIndex = cbo开单人.NewIndex
    End If
    
    txtDate.Text = Format(rsTmp!发生时间, "yyyy-MM-dd HH:mm:ss")
    
    If mint记录性质 = 2 Then
        Set rsPatiMoney = GetMoneyInfo(rsTmp!病人ID, , True)
        If Not rsPatiMoney Is Nothing Then
            sta.Panels(3).Text = "预交:" & Format(rsPatiMoney!预交余额, "0.00") & _
                "/费用:" & Format(rsPatiMoney!费用余额, gstrDec) & _
                "/剩余:" & Format(rsPatiMoney!预交余额 - rsPatiMoney!费用余额, "0.00")
        End If
    End If
    
    '------------------------------------------------------------------------------------
    If blnDelete Then
        '读取准退数,并计算应收金额,实收金额(金额=剩余金额*(准退数/剩余数))
        
        '读取单据中原始记录的费用ID
        strSQL1 = _
            " Select A.ID,A.序号,A.收费细目ID," & _
            " Nvl(A.付数,1)*A.数次" & IIF(mbln药房单位, "/Nvl(B." & mstr药房包装 & ",1)", "") & " as 原始数量" & _
            " From 病人费用记录 A,药品规格 B" & _
            " Where A.NO=[1] And A.记录状态 IN(0,1,3) And A.价格父号 is NULL" & _
            " And A.收费细目ID=B.药品ID(+) And A.记录性质=[2]"
        
        '读取药品收发记录中的准退数
        strSQL2 = _
            " Select A.费用ID,Sum(Nvl(A.付数,1)*A.实际数量" & IIF(mbln药房单位, "/Nvl(B." & mstr药房包装 & ",1)", "") & ") as 准退数量" & _
            " From 药品收发记录 A,药品规格 B" & _
            " Where A.NO=[1] And MOD(A.记录状态,3)=1" & _
            " And A.药品ID=B.药品ID(+) And A.审核人 is NULL" & _
            " And Instr([3],','||A.单据||',')>0" & _
            " Group by A.费用ID"
        
        '整张单据汇总结果(明细到收费细目)
        '执行状态应该在原始记录上判断(部分退药且部分退费的记录)
        '当退两次以上时"记录状态,序号"重复,AVG有问题,所以要用"执行状态"
        strSQL = "Select Nvl(价格父号,序号) From 病人费用记录" & _
            " Where 记录性质=[2] And 记录状态 IN(0,1,3) And NO=[1] And Nvl(执行状态,0)<>1"
        
        '如果已结帐单据禁止销帐,或是医保记帐的单据。则在原始单据行中只取未结帐部分
        If mint记录性质 = 2 Then
            If mint病人来源 = 2 Then intInsure = BillExistInsure(strNO)
            If intInsure <> 0 Then
                blnDo = Not gclsInsure.GetCapability(support允许冲销已结帐的记帐单据, , intInsure)
            Else
                blnDo = gbytBillOpt = 2
            End If
            If blnDo Then
                strSQL = strSQL & " And Nvl(价格父号,序号) IN" & _
                    " (" & _
                    " Select Nvl(价格父号,序号) as 序号" & _
                    " From 病人费用记录" & _
                    " Where NO=[1] And 记录性质 IN(2,12)" & _
                    " Group by Nvl(价格父号,序号)" & _
                    " Having Sum(Nvl(结帐金额,0))=0" & _
                    " )"
            End If
        End If
        
        '因为是将要汇总求有剩余数量的，所以不能用直接用时间限制，用序号限制
        strSQL = _
            " Select A.记录状态,A.执行状态,Nvl(A.价格父号,A.序号) as 序号," & _
            " C.编码,C.名称 as 类别,A.收费细目ID,B.名称,B.规格,Nvl(A.费用类型,B.费用类型) 费用类型," & _
            IIF(mbln药房单位, "Decode(X.药品ID,NULL,A.计算单位,X." & mstr药房单位 & ")", "A.计算单位") & " as 计算单位," & _
            " Avg(Nvl(A.付数,1)) as 付数," & _
            " Avg(A.数次" & IIF(mbln药房单位, "/Nvl(X." & mstr药房包装 & ",1)", "") & ") as 数次," & _
            " Sum(A.标准单价" & IIF(mbln药房单位, "*Nvl(X." & mstr药房包装 & ",1)", "") & ") as 单价," & _
            " Sum(A.应收金额) as 应收金额,Sum(A.实收金额) as 实收金额, " & _
            " D.名称 as 执行部门,A.附加标志" & _
            " From 病人费用记录 A,收费项目目录 B,收费项目类别 C,部门表 D,药品规格 X" & _
            " Where A.收费细目ID=B.ID And C.编码=A.收费类别 And A.执行部门ID=D.ID(+)" & _
            " And A.收费细目ID=X.药品ID(+) And A.记录性质=[2]" & _
            " And A.NO=[1] And Nvl(A.价格父号,A.序号) IN(" & strSQL & ")" & _
            " Group by A.记录状态,A.执行状态,Nvl(A.价格父号,A.序号),C.编码,C.名称,A.收费细目ID,B.名称," & _
            " B.规格,Nvl(A.费用类型,B.费用类型),A.计算单位,D.名称,A.附加标志,X.药品ID,X." & mstr药房单位 & ",X." & mstr药房包装
            
        '最后计算结果
        '当"准退数量=原始数量"时,付数才保留
        '排开已经全部退费的行(执行状态=0的一种可能)
        '有剩余数量无准退数量的有两种情况：
            '1.无对应的收发记录(如普通费用或不跟踪在用的卫材),这时应用剩余数量
            '2.收发记录中已全部发放,即已全部执行,SQL已排除这种记录
        strSQL = _
            " Select A.序号,A.编码,A.类别,A.收费细目ID,A.名称,A.规格,A.费用类型,A.计算单位," & _
            " Decode(Sign(Nvl(C.准退数量,Sum(A.付数*A.数次))-B.原始数量),0,Avg(A.付数),1) as 准退付数," & _
            " Decode(Sign(Nvl(C.准退数量,Sum(A.付数*A.数次))-B.原始数量),0,Sum(A.数次),Nvl(C.准退数量,Sum(A.付数*A.数次))) as 准退数次," & _
            " Nvl(C.准退数量,Sum(A.付数*A.数次)) as 准退数量,Sum(A.付数*A.数次) as 剩余数量," & _
            " A.单价,Sum(A.应收金额) as 剩余应收,Sum(A.实收金额) as 剩余实收,A.执行部门,A.附加标志" & _
            " From (" & strSQL & ") A,(" & strSQL1 & ") B,(" & strSQL2 & ") C" & _
            " Where A.序号=B.序号 And B.ID=C.费用ID(+)" & _
            " Group by A.序号,A.编码,A.类别,A.收费细目ID,A.名称,A.规格,A.费用类型," & _
            " A.计算单位,A.单价,B.原始数量,C.准退数量,A.执行部门,A.附加标志" & _
            " Having Sum(A.付数*A.数次)<>0"
            
        strSQL = _
            " Select A.序号,A.编码,A.类别,Nvl(B.名称,A.名称) as 名称,A.规格," & _
            " A.费用类型,A.计算单位,A.准退付数 as 付数,A.准退数次 as 数次,A.单价," & _
            " A.剩余应收*(A.准退数量/A.剩余数量) as 应收金额," & _
            " A.剩余实收*(A.准退数量/A.剩余数量) as 实收金额," & _
            " A.执行部门,A.附加标志" & _
            " From (" & strSQL & ") A,收费项目别名 B" & _
            " Where A.收费细目ID=B.收费细目ID(+) And B.码类(+)=1 And B.性质(+)=[6]" & _
            " Order by A.序号"
    Else
        '读取单据原始内容
        intSign = IIF(mblnDelete, -1, 1) '数量,金额正负符号
        
        strSQL = _
            "Select A.收费细目ID,A.收费类别,A.执行部门ID,Nvl(A.价格父号,A.序号) as 序号," & _
            " A.计算单位,A.付数,A.数次,A.标准单价,A.应收金额,A.实收金额,A.附加标志,A.费用类型" & _
            " From 病人费用记录 A Where A.记录性质=[2]" & _
            " And Instr([4],A.记录状态)>0 And A.NO=[1]" & _
            IIF(mstrTime <> "", " And A.登记时间=[5]", "")
        If blnNOMoved Then
            strSQL = strSQL & " Union ALL " & Replace(strSQL, "病人费用记录", "H病人费用记录")
        End If
        
        strSQL = _
            " Select A.序号,C.编码,C.名称 as 类别,A.收费细目ID,B.名称,B.规格,Nvl(A.费用类型,B.费用类型) 费用类型," & _
            IIF(mbln药房单位, "Decode(X.药品ID,NULL,A.计算单位,X." & mstr药房单位 & ")", "A.计算单位") & " as 计算单位," & _
            " Avg(Nvl(A.付数,1)) as 付数," & _
            " Avg([7]*A.数次" & IIF(mbln药房单位, "/Nvl(X." & mstr药房包装 & ",1)", "") & ") as 数次," & _
            " Sum(A.标准单价" & IIF(mbln药房单位, "*Nvl(X." & mstr药房包装 & ",1)", "") & ") as 单价," & _
            " Sum([7]*A.应收金额) as 应收金额,Sum([7]*A.实收金额) as 实收金额, " & _
            " D.名称 as 执行部门,A.附加标志" & _
            " From (" & strSQL & ") A,收费项目目录 B,收费项目类别 C,部门表 D,药品规格 X" & _
            " Where A.收费细目ID=B.ID And C.编码=A.收费类别" & _
            " And A.执行部门ID=D.ID(+) And A.收费细目ID=X.药品ID(+)" & _
            " Group by A.序号,C.编码,C.名称,A.收费细目ID,B.名称,B.规格," & _
            " Nvl(A.费用类型,B.费用类型),A.计算单位,D.名称,A.附加标志,X.药品ID,X." & mstr药房单位
            
        strSQL = _
            " Select A.序号,A.编码,A.类别,Nvl(B.名称,A.名称) as 名称,A.规格,A.费用类型," & _
            " A.计算单位,A.付数,A.数次,A.单价,A.应收金额,A.实收金额,A.执行部门,A.附加标志" & _
            " From (" & strSQL & ") A,收费项目别名 B" & _
            " Where A.收费细目ID=B.收费细目ID(+) And B.码类(+)=1 And B.性质(+)=[6]" & _
            " Order by 序号"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, mint记录性质, IIF(mint记录性质 = 2, ",9,25,", ",8,24,"), _
        IIF(mblnDelete, "2", "0,1,3"), CDate(IIF(mstrTime = "", "1990-01-01", mstrTime)), IIF(gbln商品名, 3, 1), intSign)
    If rsTmp.EOF Then
        If blnDelete Then
            MsgBox "单据中当前无可以操作的记录，可能单据中的项目已经全部执行。", vbInformation, gstrSysName
        Else
            MsgBox "单据中当前无可以操作的记录。", vbInformation, gstrSysName
        End If
        Exit Function
    End If
    
    Bill.Redraw = False
    Bill.Rows = rsTmp.RecordCount + 1
    For i = 1 To rsTmp.RecordCount
        Bill.RowData(i) = rsTmp!序号 '用于记帐销帐
        
        Bill.TextMatrix(i, 1) = rsTmp!类别
        Bill.TextMatrix(i, 2) = rsTmp!名称
        Bill.TextMatrix(i, 3) = Nvl(rsTmp!规格)
        Bill.TextMatrix(i, 4) = Nvl(rsTmp!计算单位)
        Bill.TextMatrix(i, 5) = Nvl(rsTmp!付数)
        Bill.TextMatrix(i, 6) = FormatEx(rsTmp!数次, 5)
        Bill.TextMatrix(i, 7) = Format(rsTmp!单价, "0.00000")
        Bill.TextMatrix(i, 8) = Format(rsTmp!应收金额, gstrDec)
        Bill.TextMatrix(i, 9) = Format(rsTmp!实收金额, gstrDec)
        Bill.TextMatrix(i, 10) = Nvl(rsTmp!执行部门)
        Bill.TextMatrix(i, 11) = IIF(rsTmp!附加标志 = 1, "√", "")
        Bill.TextMatrix(i, 12) = Nvl(rsTmp!费用类型)
        
        '设置销帐标志
        If Bill.TextMatrix(0, Bill.Cols - 1) = "删除" Then
            Bill.TextMatrix(i, Bill.Cols - 1) = "√"
        End If
        
        rsTmp.MoveNext
    Next
    '针对列编辑性质设置颜色
    Bill.SetColColor 1, &HE7CFBA
    Bill.SetColColor 2, &HE7CFBA
    Bill.SetColColor 6, &HE7CFBA
    Bill.SetColColor 10, &HE7CFBA
    Bill.SetColColor 5, &HE0E0E0
    Bill.SetColColor 7, &HE0E0E0
    Bill.SetColColor 11, &HE0E0E0
    Call SetColNum
    Bill.Redraw = True
    
    '----------------------------------------------------------------------------
    If blnDelete Then
        '读取准退数,并计算应收金额,实收金额(金额=剩余金额*(准退数/剩余数))

        '读取药品收发记录中的准退数
        strSQL1 = _
            " Select A.费用ID,Sum(Nvl(A.付数,1)*A.实际数量" & IIF(mbln药房单位, "/Nvl(B." & mstr药房包装 & ",1)", "") & ") as 准退数量" & _
            " From 药品收发记录 A,药品规格 B" & _
            " Where A.NO=[1] And MOD(A.记录状态,3)=1" & _
            " And A.药品ID=B.药品ID(+) And A.审核人 is NULL" & _
            " And Instr([3],','||A.单据||',')>0" & _
            " Group by A.费用ID"
        
        '整张费用单据(明细到收入项目)
        '执行状态应该在原始记录上判断(部分退药且部分退费的记录)
        strSQL = "Select Nvl(价格父号,序号) From 病人费用记录" & _
            " Where 记录性质=[2] And 记录状态 IN(0,1,3) And NO=[1] And Nvl(执行状态,0)<>1"
        If blnDo Then
            strSQL = strSQL & " And Nvl(价格父号,序号) IN" & _
                " (" & _
                " Select Nvl(价格父号,序号) as 序号" & _
                " From 病人费用记录" & _
                " Where NO=[1] And 记录性质 IN(2,12)" & _
                " Group by Nvl(价格父号,序号)" & _
                " Having Sum(Nvl(结帐金额,0))=0" & _
                " )"
        End If
        
        strSQL = _
            " Select Sum(A.ID) as ID,A.序号,A.名称,A.收费类别," & _
            " Sum(A.数量) as 剩余数量,Sum(A.应收金额) as 剩余应收," & _
            " Sum(A.实收金额) as 剩余实收 From (" & _
            " Select Decode(A.记录状态,2,0,A.ID) as ID,A.序号,B.名称,A.收费类别," & _
            " Nvl(A.付数,1)*A.数次" & IIF(mbln药房单位, "/Nvl(X." & mstr药房包装 & ",1)", "") & " as 数量," & _
            " A.应收金额,A.实收金额" & _
            " From 病人费用记录 A,收入项目 B,药品规格 X" & _
            " Where A.记录性质=[2] And A.NO=[1]" & _
            " And A.收入项目ID=B.ID And Nvl(A.价格父号,A.序号) IN(" & strSQL & ")" & _
            " And A.收费细目ID=X.药品ID(+)) A" & _
            " Group by A.序号,A.名称,A.收费类别" & _
            " Having Sum(A.数量)<>0"
                    
        '最后计算结果
        strSQL = _
            " Select A.名称,Sum(A.剩余应收*(A.准退数量/A.剩余数量)) as 应收金额," & _
            " Sum(剩余实收*(A.准退数量/A.剩余数量)) as 实收金额 From (" & _
            " Select A.名称,A.剩余数量,A.剩余应收,A.剩余实收," & _
            " Decode(Instr(',4,5,6,7,',A.收费类别),0,A.剩余数量,Nvl(B.准退数量,A.剩余数量)) as 准退数量" & _
            " From (" & strSQL & ") A,(" & strSQL1 & ") B" & _
            " Where A.ID=B.费用ID(+)" & _
            " ) A Group by A.名称"
    Else
        '读取单据原始内容
        intSign = IIF(mblnDelete, -1, 1) '数量,金额正负符号
        
        strSQL = "Select A.收入项目ID,A.应收金额,A.实收金额 From 病人费用记录 A" & _
            " Where Instr([4],A.记录状态)>0 And A.记录性质=[2] And A.NO=[1]" & _
            IIF(mstrTime <> "", " And A.登记时间=[5]", "")
        If blnNOMoved Then
            strSQL = strSQL & " Union ALL " & Replace(strSQL, "病人费用记录", "H病人费用记录")
        End If
        
        strSQL = _
            " Select B.名称,Sum([6]*A.应收金额) as 应收金额,Sum([6]*A.实收金额) as 实收金额 " & _
            " From (" & strSQL & ") A,收入项目 B Where A.收入项目ID=B.ID Group By B.名称"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, mint记录性质, IIF(mint记录性质 = 2, ",9,25,", ",8,24,"), _
        IIF(mblnDelete, "2", "0,1,3"), CDate(IIF(mstrTime = "", "1990-01-01", mstrTime)), intSign)
    If rsTmp.EOF Then Exit Function
    
    '刷新显示(收费要叠加)
    mshMoney.Rows = rsTmp.RecordCount + 1
    If mshMoney.Rows < 5 Then mshMoney.Rows = 5
    Call SetMoneyList
    
    For i = 1 To rsTmp.RecordCount
        mshMoney.TextMatrix(i, 0) = rsTmp!名称
        mshMoney.TextMatrix(i, 1) = Format(rsTmp!实收金额, gstrDec)
        curTotal = curTotal + rsTmp!实收金额
        cur应收Total = cur应收Total + rsTmp!应收金额
        rsTmp.MoveNext
    Next
    
    txt实收.Text = Format(curTotal, gstrDec)
    txt应收.Text = Format(cur应收Total, gstrDec)
    
    ReadBill = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetShowCol()
'功能：付数列的控制(浏览时展开)
    mrsClass.Filter = "编码='7'"
    If mrsClass.RecordCount = 0 Then
        Bill.ColWidth(5) = 0
    ElseIf Bill.ColWidth(5) = 0 Then
        Bill.ColWidth(5) = 520
    End If
End Sub

Private Sub ClearRows()
    Dim i As Long
    For i = 1 To Bill.Rows - 1
        Bill.RowData(i) = 0
    Next
End Sub

Private Function GetPay(lngRow As Long) As Integer
    Dim i As Long
    '取其它中药的付数
    GetPay = 1
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).收费类别 = "7" And i <> lngRow Then
            GetPay = mobjBill.Details(i).付数
            Exit For
        End If
    Next
End Function

Private Function GetDetailNum(lngRow As Long) As Double
'功能：获取病人指定细目的总记帐数据(含本单据中)
'参数：lngRow=当前单据行
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim lngNum As Long, i As Long
    
    If lngRow <= mobjBill.Details.Count Then
        '当前单据中的数量
        For i = 1 To mobjBill.Details.Count
            If i <> lngRow And mobjBill.Details(i).收费细目ID = mobjBill.Details(lngRow).收费细目ID Then
                lngNum = lngNum + mobjBill.Details(i).数次 * IIF(mobjBill.Details(i).付数 = 0, 1, mobjBill.Details(i).付数)
            End If
        Next
        '数据库中的数量
        strSQL = _
            "Select Sum(A.数次*Nvl(A.付数,1)" & IIF(mbln药房单位, "/Nvl(B." & mstr药房包装 & ",1)", "") & ") as NUM" & _
            " From 病人费用记录 A,药品规格 B" & _
            " Where A.价格父号 is Null And A.记录状态<>0 And A.记帐费用=1" & _
            " And A.病人ID=[1] And Nvl(A.主页ID,0)=[2]" & _
            " And A.收费细目ID=B.药品ID(+) And A.收费细目ID+0=[3]"
            
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mrsInfo!病人ID), Val(Nvl(mrsInfo!主页ID, 0)), mobjBill.Details(lngRow).收费细目ID)
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
    Dim strSQL As String, bytDay As Byte
    Dim str药房 As String, lng开单科室ID As Long
    
    lng开单科室ID = mrsInfo!科室ID    '开单科室优先
    If lng开单科室ID = 0 And cbo开单科室.ListIndex <> -1 Then lng开单科室ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
    
    If str类别 = "4" Then
        strSQL = _
            " Select Distinct C.ID,C.编码,C.简码,C.名称,B.工作性质,B.服务对象 " & _
            " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
            " Where A.执行科室ID+0=B.部门ID And B.工作性质='发料部门'" & _
            " And B.服务对象 IN([1],3) And B.部门ID=C.ID" & _
            " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
            " And (A.病人来源 is NULL Or A.病人来源=[1])" & _
            " And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
            " And A.收费细目ID=[3]" & _
            " Order by B.服务对象,C.编码"
            
        '以及SQL在卫材不支持存储库房设置之前用
'        strSQL = "Select A.ID,A.编码,A.简码,A.名称,B.工作性质,B.服务对象" & _
'            " From 部门表 A,部门性质说明 B" & _
'            " Where A.ID=B.部门ID And B.工作性质='发料部门' And B.服务对象 IN([1],3)" & _
'            " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
'            " Order by B.服务对象,A.编码"
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
        If Not Check上班安排(True) Then
            strSQL = _
                " Select Distinct C.ID,C.编码,C.简码,C.名称,B.工作性质,B.服务对象 " & _
                " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
                " Where A.执行科室ID+0=B.部门ID And B.工作性质=[4]" & _
                " And B.服务对象 IN([1],3) And B.部门ID=C.ID" & _
                " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                " And (A.病人来源 is NULL Or A.病人来源=[1])" & _
                " And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
                " And A.收费细目ID=[3]" & _
                " Order by B.服务对象,C.编码"
        Else
            bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=周日,1=周一
            strSQL = _
                " Select Distinct C.ID,C.编码,C.简码,C.名称,B.工作性质,B.服务对象 " & _
                " From 收费执行科室 A,部门性质说明 B,部门表 C,部门安排 D" & _
                " Where A.执行科室ID+0=B.部门ID And B.工作性质=[4]" & _
                " And B.服务对象 IN([1],3) And B.部门ID=C.ID" & _
                " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                " And D.部门ID=C.ID And D.星期=[5]" & _
                " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.开始时间,'HH24:MI:SS') and To_Char(D.终止时间,'HH24:MI:SS') " & _
                " And (A.病人来源 is NULL Or A.病人来源=[1])" & _
                " And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
                " And A.收费细目ID=[3]" & _
                " Order by B.服务对象,C.编码"
        End If
    End If
    
    On Error GoTo errH
    'Set mrsWork = New ADODB.Recordset
    Set mrsWork = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mint病人来源, lng开单科室ID, lng药品ID, str药房, bytDay)
    GetWorkUnit = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Load开单人(ByVal lng科室ID As Long)
'功能：读取当前开单科室中包含的所有人员
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim lngOldID As Long
    
    cbo开单人.Clear
    
    '科室医生或护士
    strSQL = _
        "Select Distinct A.ID,B.部门ID,A.编号,A.姓名,Upper(A.简码) as 简码," & _
        " C.人员性质,Nvl(A.聘任技术职务,0) as 职务" & _
        " From 人员表 A,部门人员 B,人员性质说明 C" & _
        " Where A.ID=B.人员ID And A.ID=C.人员ID" & _
        " And C.人员性质 IN('医生','护士') And B.部门ID=[1]"
    '仅为开嘱医生
    If lng科室ID = mlng开嘱科室ID And mlng开嘱科室ID <> mlng开单科室ID Then
        strSQL = strSQL & " And A.姓名=[2]"
    End If
    strSQL = strSQL & " Order by 简码,人员性质 Desc"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng科室ID, mstr开嘱医生)
    
    i = IIF(rsTmp.RecordCount = 0, 0, rsTmp.RecordCount - 1)
    ReDim marrDr(i)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            If lngOldID <> rsTmp!ID Then
                cbo开单人.AddItem IIF(IsNull(rsTmp!简码), "", rsTmp!简码 & "-") & rsTmp!姓名
                cbo开单人.ItemData(cbo开单人.ListCount - 1) = rsTmp!部门ID
                marrDr(i - 1) = rsTmp!ID & "|" & rsTmp!部门ID & "|" & Nvl(rsTmp!编号) & "|" & rsTmp!姓名 & "|" & Nvl(rsTmp!简码) & "|" & rsTmp!职务 & "|" & Nvl(rsTmp!人员性质)
                
                If rsTmp!ID = UserInfo.ID And cbo开单人.ListIndex = -1 Then cbo开单人.ListIndex = cbo开单人.NewIndex
                lngOldID = rsTmp!ID
            End If
            rsTmp.MoveNext
        Next
        
        If cbo开单人.ListCount > 0 Then ReDim Preserve marrDr(cbo开单人.ListCount - 1)
        
        If cbo开单人.ListCount = 1 And cbo开单人.ListIndex = -1 Then cbo开单人.ListIndex = 0
    End If
End Sub

Private Function CalcGridToTal(Optional bln应收 As Boolean) As Currency
    Dim objTmpDetail As New BillDetail
    Dim objTmpIncome As New BillInCome
    Dim i As Long, intCol As Integer

    If mobjBill.Details.Count > 0 Then
        For Each objTmpDetail In mobjBill.Details
            For Each objTmpIncome In objTmpDetail.InComes
                If bln应收 Then
                    CalcGridToTal = CalcGridToTal + objTmpIncome.应收金额
                Else
                    CalcGridToTal = CalcGridToTal + objTmpIncome.实收金额
                End If
            Next
        Next
    Else
        For i = 1 To Bill.Cols - 1
            If bln应收 Then
                If Bill.TextMatrix(0, i) = "应收金额" Then intCol = i: Exit For
            Else
                If Bill.TextMatrix(0, i) = "实收金额" Then intCol = i: Exit For
            End If
        Next
    
        For i = 1 To Bill.Rows - 1
            CalcGridToTal = CalcGridToTal + Val(Bill.TextMatrix(i, intCol))
        Next
    End If
End Function

Private Sub ShowDeleteCol(blnShow As Boolean)
'功能：显示\隐藏销帐标志列
    Dim i As Long, blnACT As Boolean
    If blnShow Then
        If Bill.TextMatrix(0, Bill.Cols - 1) <> "删除" Then
            Bill.Redraw = False
            Bill.Cols = Bill.Cols + 1
            Bill.TextMatrix(0, Bill.Cols - 1) = "删除"
            Bill.ColAlignment(Bill.Cols - 1) = 4
            Bill.ColWidth(Bill.Cols - 1) = 550
            Bill.ColData(Bill.Cols - 1) = -1
            
            blnACT = Bill.Active: Bill.Active = False
            Bill.Row = 0: Bill.Col = Bill.Cols - 1: Bill.MsfObj.CellForeColor = vbRed
            Bill.Row = 1: Bill.Col = Bill.Cols - 1
            Bill.Active = blnACT
            
            Bill.ColWidth(1) = GetOrigColWidth(1) - 120
            Bill.ColWidth(2) = GetOrigColWidth(2) - 100
            Bill.ColWidth(10) = GetOrigColWidth(10) - 200
            
            Bill.ColWidth(7) = GetOrigColWidth(7) - 50
            Bill.ColWidth(8) = GetOrigColWidth(8) - 50
            Bill.ColWidth(9) = GetOrigColWidth(9) - 50
            Bill.Redraw = True
        End If
    Else
        If Bill.TextMatrix(0, Bill.Cols - 1) = "删除" Then
            Bill.Redraw = False
            Bill.Cols = Bill.Cols - 1
            Bill.ColWidth(1) = GetOrigColWidth(1)
            Bill.ColWidth(2) = GetOrigColWidth(2)
            Bill.ColWidth(10) = GetOrigColWidth(10)
            
            Bill.ColWidth(7) = GetOrigColWidth(7)
            Bill.ColWidth(8) = GetOrigColWidth(8)
            Bill.ColWidth(9) = GetOrigColWidth(9)
            Bill.Redraw = True
        End If
    End If
End Sub

Private Function GetOrigColWidth(ByVal intIdx As Integer) As Long
'功能：获取指定列的原始列宽
    GetOrigColWidth = Val(Split(Split(STR_HEAD, ";")(intIdx), ",")(1))
End Function

Private Sub SetColNum(Optional intRow As Long = 1)
'功能：重新显示各行的行号
'参数：intRow=从该行开始
    Dim bln As Boolean, i As Long
    
    Bill.Redraw = False
    For i = intRow To Bill.Rows - 1
        Bill.TextMatrix(i, 0) = i
    Next
    Bill.Redraw = True
End Sub

Private Function CheckDuty(Optional tmpDetail As Detail, Optional blnCommon As Boolean = True) As Integer
'功能：检查指定药品行的职务是否与当前医生的职务相匹配
'参数：tmpDetail=输入的项目,不传为所有行,blnCommon=是否正常的判断,否则为医保或公费病人的判断
'返回：不匹配的行,0为正确
'说明：职务：1=正高,2=副高,3=中级,4=助理/师级,5=员/士,9=待聘
    Dim i As Long, int职务A As Integer, int职务B As Integer
    Dim strTmp As String
    
    strTmp = "正高,副高,中级,助理/师级,员/士,,,,待聘"
    
    If cbo开单人.ListIndex = -1 Then Exit Function
    If cbo开单人.ListIndex <= UBound(marrDr) Then
        If UBound(Split(marrDr(cbo开单人.ListIndex), "|")) >= 5 Then
            int职务A = Val(Split(marrDr(cbo开单人.ListIndex), "|")(5))
        End If
    End If
        
    If tmpDetail Is Nothing Then
        For i = 1 To mobjBill.Details.Count
            If InStr(",5,6,7,", mobjBill.Details(i).收费类别) > 0 Then
                If Not blnCommon Then
                    int职务B = Val(Right(mobjBill.Details(i).Detail.处方职务, 1))
                    If int职务B > 0 Then
                        If int职务A = 0 Then
                            strTmp = "对医保或公费病人,第 " & i & " 行药品""" & mobjBill.Details(i).Detail.名称 & """要求医生职务至少为""" & Split(strTmp, ",")(int职务B - 1) & """,而当前医生未设置职务！"
                            CheckDuty = 1
                        ElseIf int职务B < int职务A Then
                            strTmp = "对医保或公费病人,第 " & i & " 行药品""" & mobjBill.Details(i).Detail.名称 & """要求医生职务为""" & Split(strTmp, ",")(int职务B - 1) & """以上,而当前医生职务为""" & Split(strTmp, ",")(int职务A - 1) & """！"
                            CheckDuty = i: Exit For
                        End If
                    End If
                Else
                    int职务B = Val(Left(mobjBill.Details(i).Detail.处方职务, 1))
                    If int职务B > 0 Then
                        If int职务A = 0 Then
                            strTmp = "第 " & i & " 行药品""" & mobjBill.Details(i).Detail.名称 & """要求医生职务至少为""" & Split(strTmp, ",")(int职务B - 1) & """,而当前医生未设置职务！"
                            CheckDuty = 1
                        ElseIf int职务B < int职务A Then
                            strTmp = "第 " & i & " 行药品""" & mobjBill.Details(i).Detail.名称 & """要求医生职务为""" & Split(strTmp, ",")(int职务B - 1) & """以上,而当前医生职务为""" & Split(strTmp, ",")(int职务A - 1) & """！"
                            CheckDuty = i: Exit For
                        End If
                    End If
                End If
            End If
        Next
    Else
        If InStr(",5,6,7,", tmpDetail.类别) = 0 Then Exit Function
        If Not blnCommon Then
            int职务B = Val(Right(tmpDetail.处方职务, 1))
            If int职务B > 0 Then
                If int职务A = 0 Then
                    strTmp = "对医保或公费病人,药品""" & tmpDetail.名称 & """要求医生职务至少为""" & Split(strTmp, ",")(int职务B - 1) & """,而当前医生未设置职务！"
                    CheckDuty = 1
                ElseIf int职务B < int职务A Then
                    strTmp = "对医保或公费病人,药品""" & tmpDetail.名称 & """要求医生职务为""" & Split(strTmp, ",")(int职务B - 1) & """以上,而当前医生职务为""" & Split(strTmp, ",")(int职务A - 1) & """！"
                    CheckDuty = 1
                End If
            End If
        Else
            int职务B = Val(Left(tmpDetail.处方职务, 1))
            If int职务B > 0 Then
                If int职务A = 0 Then
                    strTmp = "药品""" & tmpDetail.名称 & """要求医生职务至少为""" & Split(strTmp, ",")(int职务B - 1) & """,而当前医生未设置职务！"
                    CheckDuty = 1
                ElseIf int职务B < int职务A Then
                    strTmp = "药品""" & tmpDetail.名称 & """要求医生职务为""" & Split(strTmp, ",")(int职务B - 1) & """以上,而当前医生职务为""" & Split(strTmp, ",")(int职务A - 1) & """！"
                    CheckDuty = 1
                End If
            End If
        End If
    End If
    
    If CheckDuty > 0 Then MsgBox strTmp, vbInformation, gstrSysName
End Function

Private Function PhysicExist(objDetail As Detail, intRow As Integer) As Boolean
'功能：判断指定药品在单据中是否已经存在
'参数：objDetail=项目,intRow=要判断的行
'说明：时价或分批药品在同一药房禁止重复输入(这里仅提示,保存时禁止)
    Dim i As Integer
    
    For i = 1 To mobjBill.Details.Count
        If i <> intRow And InStr(",4,5,6,7,", mobjBill.Details(i).收费类别) > 0 Then
            If mobjBill.Details(i).Detail.ID = objDetail.ID Then
                If (mobjBill.Details(i).Detail.分批 Or mobjBill.Details(i).Detail.变价) _
                    And (objDetail.分批 Or objDetail.变价) Then
                    If objDetail.类别 = "4" Then
                        If MsgBox("卫生材料""" & objDetail.名称 & """在第 " & i & " 行已经输入,要继续吗？" & _
                            vbCrLf & vbCrLf & "注意：该卫生材料为分批或时价药品,重复输入时必须保证它们的发料部门不同。", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            PhysicExist = True
                        End If
                    Else
                        If MsgBox("药品""" & objDetail.名称 & """在第 " & i & " 行已经输入,要继续吗？" & _
                            vbCrLf & vbCrLf & "注意：该药品为分批或时价药品,重复输入时必须保证它们的执行药房不同。", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            PhysicExist = True
                        End If
                    End If
                    Exit Function
                Else
                    If objDetail.类别 = "4" Then
                        If MsgBox("卫生材料""" & objDetail.名称 & """在第 " & i & " 行已经输入,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            PhysicExist = True
                        End If
                    Else
                        If MsgBox("药品""" & objDetail.名称 & """在第 " & i & " 行已经输入,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            PhysicExist = True
                        End If
                    End If
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Private Function Check费用类型(Optional intRow As Integer) As Boolean
'功能：根据当前病人的类型判断指定行的项目是否可以输入,适用于所有类别的项目
    Dim strSQL As String
    Dim i As Long, bytType As Byte
    Dim rsTmp As New ADODB.Recordset
    
    Check费用类型 = True
    
    '无法检查
    If txt付款方式.Tag = "" Then Exit Function
    
    '确定病人类型
    bytType = Val(txt付款方式.Tag)
    
    '只检查医保病人和公费病人
    If bytType <> 1 And bytType <> 2 Then Exit Function
    
    '读取检查数据
    If bytType = 1 Then
        strSQL = "Select * From 费用类型 Where 编码 In(" & gstr医保费用类型 & ") Order by 编码"
    Else
        strSQL = "Select * From 费用类型 Where 编码 In(" & gstr公费费用类型 & ") Order by 编码"
    End If
    rsTmp.CursorLocation = adUseClient
    Call SQLTest(App.ProductName, Me.Caption, strSQL)
    rsTmp.Open strSQL, gcnOracle, adOpenKeyset
    Call SQLTest
    If rsTmp.EOF Then Exit Function
    
    If intRow > 0 Then
        If mobjBill.Details(intRow).Detail.类型 = "" Then
            MsgBox """" & mobjBill.Details(intRow).Detail.名称 & """的费用类型未设置！", vbInformation, gstrSysName
            Check费用类型 = False
        Else
            rsTmp.Filter = "名称='" & mobjBill.Details(intRow).Detail.类型 & "'"
            If rsTmp.EOF Then
                MsgBox """" & mobjBill.Details(intRow).Detail.名称 & """的费用类型为""" & _
                    mobjBill.Details(intRow).Detail.类型 & """,不是" & _
                    IIF(bytType = 1, "医保", "公费") & "费用类型！", vbInformation, gstrSysName
                Check费用类型 = False
            End If
        End If
    Else
        For i = 1 To mobjBill.Details.Count
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
                        IIF(bytType = 1, "医保", "公费") & "费用类型！" & vbCrLf & "确实要保存单据吗？", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Check费用类型 = False: Exit For
                    End If
                End If
            End If
        Next
    End If
End Function

Private Sub ReCalcInsure()
'功能：修改单据时,重新计算统筹金额及更新相关信息
    Dim i As Long, j As Long
    Dim strInfo As String
    
    If Not IsNull(mrsInfo!险类) Then
        For i = 1 To mobjBill.Details.Count
            For j = 1 To mobjBill.Details(i).InComes.Count
                strInfo = gclsInsure.GetItemInsure(mobjBill.病人ID, mobjBill.Details(i).收费细目ID, mobjBill.Details(i).InComes(j).实收金额, False, mrsInfo!险类)
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
            Next
        Next
    End If
End Sub

Private Function HaveStopClass() As Integer
'功能：判断当前单据中是否有护士禁止输入的内容
    Dim i As Long, str性质 As String
    
    If cbo开单人.ListIndex <> -1 Then
        If cbo开单人.ListIndex <= UBound(marrDr) Then
            If UBound(Split(marrDr(cbo开单人.ListIndex), "|")) >= 6 Then
                str性质 = Split(marrDr(cbo开单人.ListIndex), "|")(6)
            End If
        End If
    End If
    
    For i = 1 To mobjBill.Details.Count
        If str性质 = "护士" And InStr(",E,M,4,", mobjBill.Details(i).收费类别) = 0 Then
            HaveStopClass = i: Exit Function
        End If
    Next
End Function

Private Function Check执行科室() As Integer
    Dim i As Long
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).执行部门ID = 0 Or Bill.TextMatrix(i, 10) = "" Then
            Check执行科室 = i: Exit Function
        End If
    Next
End Function

Public Sub InitLocPar()
'功能：初始化费用本机参数
    mstrLike = IIF(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = 0, "%", "")
    mblnPay = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "中药付数", 1)) <> 0
    mblnTime = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "变价数次", 0)) <> 0
    mbln其它药房 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "显示其它药房库存", 0)) = 1
    mbln其它药库 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "显示其它药库库存", 0)) = 1
    mstr收费类别 = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "收费类别", "")
    
    '药品单位
    mbln药房单位 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "药品单位", 0)) <> 0
    If mint病人来源 = 1 Then
        mstr药房单位 = "门诊单位": mstr药房包装 = "门诊包装"
    Else
        mstr药房单位 = "住院单位": mstr药房包装 = "住院包装"
    End If
    
    '缺省药房
    mlng西药房 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, IIF(mint病人来源 = 2, "住院", "门诊") & "缺省西药房", 0))
    mlng成药房 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, IIF(mint病人来源 = 2, "住院", "门诊") & "缺省成药房", 0))
    mlng中药房 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, IIF(mint病人来源 = 2, "住院", "门诊") & "缺省中药房", 0))
End Sub
