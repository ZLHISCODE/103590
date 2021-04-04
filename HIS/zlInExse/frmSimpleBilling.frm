VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSimpleBilling 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "住院简单记帐"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10365
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSimpleBilling.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   10365
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   25
      Top             =   6285
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmSimpleBilling.frx":08CA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12065
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   88
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmSimpleBilling.frx":115E
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmSimpleBilling.frx":1798
            Key             =   "WB"
            Object.ToolTipText     =   "五笔(F7)"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
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
   Begin VB.PictureBox picAppend 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1905
      Left            =   0
      ScaleHeight     =   1905
      ScaleWidth      =   10365
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   4380
      Width           =   10365
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
         Height          =   585
         Left            =   0
         TabIndex        =   34
         ToolTipText     =   "清除:F6"
         Top             =   -105
         Width           =   10290
         Begin VB.ComboBox cboBaby 
            Height          =   360
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   165
            Width           =   1680
         End
         Begin VB.CheckBox chk加班 
            Caption         =   "加班(&A)"
            Height          =   270
            Left            =   120
            TabIndex        =   11
            Top             =   210
            Width           =   1155
         End
         Begin VB.ComboBox cbo开单人 
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   4815
            TabIndex        =   14
            Top             =   165
            Width           =   1785
         End
         Begin MSMask.MaskEdBox txtDate 
            Height          =   360
            Left            =   7770
            TabIndex        =   15
            Top             =   165
            Width           =   2430
            _ExtentX        =   4286
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
            Caption         =   "婴儿费"
            Height          =   240
            Left            =   1320
            TabIndex        =   12
            Top             =   225
            Width           =   720
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            Caption         =   "时间"
            Height          =   240
            Left            =   7215
            TabIndex        =   36
            Top             =   225
            Width           =   480
         End
         Begin VB.Label lbl开单人 
            AutoSize        =   -1  'True
            Caption         =   "开单人"
            Height          =   240
            Left            =   4020
            TabIndex        =   35
            Top             =   225
            Width           =   720
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   420
         Left            =   7230
         TabIndex        =   19
         ToolTipText     =   "热键:Esc"
         Top             =   1200
         Width           =   1500
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   420
         Left            =   7230
         TabIndex        =   18
         ToolTipText     =   "热键：F2"
         Top             =   675
         Width           =   1500
      End
      Begin VB.Frame fraMoney 
         Height          =   1545
         Left            =   0
         TabIndex        =   37
         Top             =   360
         Width           =   3195
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshMoney 
            Height          =   1335
            Left            =   60
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   165
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   2355
            _Version        =   393216
            Rows            =   4
            FixedCols       =   0
            RowHeightMin    =   300
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            FocusRect       =   0
            HighLight       =   0
            ScrollBars      =   2
            FormatString    =   "^项目       |^金额        "
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
      Begin VB.Frame fraStat 
         Height          =   1545
         Left            =   3195
         TabIndex        =   38
         Top             =   360
         Width           =   3405
         Begin VB.TextBox txt实收 
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
            Left            =   1125
            Locked          =   -1  'True
            TabIndex        =   17
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   870
            Width           =   1845
         End
         Begin VB.TextBox txt应收 
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
            ForeColor       =   &H00000000&
            Height          =   450
            Left            =   1125
            Locked          =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   345
            Width           =   1845
         End
         Begin VB.Label lbl实收 
            AutoSize        =   -1  'True
            Caption         =   "实收"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   300
            Left            =   450
            TabIndex        =   40
            Top             =   945
            Width           =   630
         End
         Begin VB.Label lbl应收 
            AutoSize        =   -1  'True
            Caption         =   "应收"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   450
            TabIndex        =   39
            Top             =   420
            Width           =   630
         End
      End
   End
   Begin VB.Frame fraTitle 
      Height          =   645
      Left            =   30
      TabIndex        =   23
      ToolTipText     =   "清除:F6"
      Top             =   -120
      Width           =   10275
      Begin VB.ComboBox cboNO 
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   8220
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   210
         Width           =   1380
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
         Height          =   420
         Left            =   9690
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "热键:F8"
         Top             =   180
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
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   9720
         TabIndex        =   31
         Top             =   195
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "住院记帐单"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   75
         TabIndex        =   26
         ToolTipText     =   "清除:F6"
         Top             =   225
         Width           =   1725
      End
      Begin VB.Label lblNO 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "单据号"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   7365
         TabIndex        =   24
         Top             =   270
         Width           =   720
      End
   End
   Begin VB.Frame fraInfo 
      Height          =   1050
      Left            =   30
      TabIndex        =   22
      Top             =   405
      Width           =   10275
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   360
         Left            =   615
         TabIndex        =   45
         Top             =   195
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   635
         Appearance      =   2
         IDKindStr       =   $"frmSimpleBilling.frx":1DD2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   12
         FontName        =   "宋体"
         IDKind          =   -1
         ShowPropertySet =   -1  'True
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         MustSelectItems =   "姓名"
         BackColor       =   -2147483633
      End
      Begin VB.TextBox txt担保额 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   6195
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   615
         Width           =   1095
      End
      Begin VB.TextBox txt担保人 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   4050
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   615
         Width           =   840
      End
      Begin VB.ComboBox cbo开单科室 
         Height          =   360
         Left            =   8415
         TabIndex        =   9
         Text            =   "cbo开单科室"
         Top             =   615
         Width           =   1755
      End
      Begin VB.TextBox txt医疗付款 
         Height          =   360
         Left            =   1305
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   615
         Width           =   1680
      End
      Begin VB.TextBox txt床号 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   6780
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   195
         Width           =   525
      End
      Begin VB.TextBox txt费别 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   8430
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   195
         Width           =   1740
      End
      Begin VB.TextBox txtSex 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   4050
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   195
         Width           =   600
      End
      Begin VB.TextBox txtPatient 
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   1290
         MaxLength       =   64
         TabIndex        =   1
         Top             =   195
         Width           =   1680
      End
      Begin VB.TextBox txtOld 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   5295
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   195
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "担保额"
         Height          =   240
         Left            =   5340
         TabIndex        =   44
         Top             =   675
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "担保人"
         Height          =   240
         Left            =   3285
         TabIndex        =   43
         Top             =   675
         Width           =   720
      End
      Begin VB.Label lbl开单科室 
         AutoSize        =   -1  'True
         Caption         =   "开单科室"
         Height          =   240
         Left            =   7395
         TabIndex        =   42
         Top             =   675
         Width           =   960
      End
      Begin VB.Label lbl医疗付款 
         AutoSize        =   -1  'True
         Caption         =   "付款方式"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   390
         TabIndex        =   41
         Top             =   690
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床号"
         Height          =   240
         Left            =   6195
         TabIndex        =   32
         Top             =   255
         Width           =   480
      End
      Begin VB.Label lblPatient 
         AutoSize        =   -1  'True
         Caption         =   "病人"
         ForeColor       =   &H80000007&
         Height          =   240
         Index           =   7
         Left            =   135
         TabIndex        =   30
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "性别"
         Height          =   240
         Index           =   8
         Left            =   3495
         TabIndex        =   29
         Top             =   255
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "年龄"
         Height          =   240
         Index           =   9
         Left            =   4770
         TabIndex        =   28
         Top             =   255
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "病人费别"
         Height          =   240
         Index           =   12
         Left            =   7395
         TabIndex        =   27
         Top             =   255
         Width           =   960
      End
   End
   Begin ZL9BillEdit.BillEdit Bill 
      Height          =   2895
      Left            =   30
      TabIndex        =   10
      Top             =   1455
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   5106
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   8445
      Top             =   5205
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
            Picture         =   "frmSimpleBilling.frx":1DDE
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSimpleBilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
'————————————————————————————————————————————————————————————————————
'入口参数：
'2.表单初始状态参数：
Public mbytInState As Byte '0-执行,1-查阅,2-调整,3-销帐
Public mstrInNO As String '当mbytInState=1时有效,等于单据号
Public mblnNOMoved As Boolean '操作的单据是否在后备数据表中,查阅时传入
 
Public mstrTime As String '操作单据内容的登记时间
Public mblnDelete As Boolean '是否查阅退费单据

Public mlngUnitID As Long '当前记帐病区,为0时表示所有病区
Public mlngDeptID As Long '当前记帐科室,为0时表示所有科室
Public mbytUseType As Byte '记帐单用途,0-普通记帐,1-按科室分散记帐,2-医技科室记帐
Public mlng病人ID As Long '科室分散记帐用
Public mstrPrivs As String
Public mlngModule As Long
Private mobjICCard As Object
Private mstrPrivsOpt As String '记帐操作1150模块的授权功能
Private mcurModiMoney As Currency '修改单据时原单据的金额
Private mstrUnitIDs As String   '当前操作员的所有病区ID
'————————————————————————————————————————————————————————————————————
Private mblnNotCick As Boolean
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
    项目 = 0
    应收金额 = 1
    实收金额 = 2
    执行科室 = 3
    类型 = 4
End Enum

'数据对象
Private mrsUnit As ADODB.Recordset '可选择的执行科室
Private mrsInfo As New ADODB.Recordset '病人信息
Private mrsMedAudit As ADODB.Recordset  '病人已审批的费用项目
Private mrsMedPayMode As ADODB.Recordset '所有可用的医疗付款方式

'程序对象
Private mobjBill As ExpenseBill '★★★费用单据对象★★★
Private mobjBillDetail As BillDetail '单据的收费细目对象
Private mobjBillIncome As BillInCome '收费细目的收入项目对象
Private mobjDetail As Detail '单独的收费细目对象
Private mcolDetails As Details '单独的收费细目集合
Private mcolMoneys As BillInComes

'程序变量
Private mstrWarn As String '已经报过警并选择继续的类别
Private mrsWarn As ADODB.Recordset '病区报警线
Private mrs开单科室 As ADODB.Recordset  '可选的开单科室
Private mrs开单人 As ADODB.Recordset    '可选医生和护士

Private mblnDrop As Boolean '在KeyDown中判断cbo开单人当前是否弹出
Private mblnValid As Boolean
Private mblnPrint As Boolean '读取审核单时是否包含要打印的收费类别
Private marrColData() As Integer '当前单据编辑属性映象
Private mblnSelect As Boolean '用于控制收费细目对象是否来自于列表选择或选择器

Private Const STR_HEAD = "项目,3000,1;应收金额,1500,7;实收金额,1500,7;执行科室,1900,1;类型,870,1"
'-----------------------------------------------------------------------------------
'结算卡相关
Private mstrPassWord As String
'-----------------------------------------------------------------------------------
Private mstr药品价格等级 As String, mstr卫材价格等级 As String, mstr普通价格等级 As String

Private Sub Bill_cboClick(ListIndex As Long)
    Dim lng执行科室 As Long, str执行科室 As String
    If ListIndex <> -1 And Bill.TextMatrix(0, Bill.Col) = "执行科室" Then
        If mobjBill.Details.Count >= Bill.Row Then
            With mobjBill.Details(Bill.Row)
                If .执行部门ID <> Bill.ItemData(Bill.ListIndex) Then
                    lng执行科室 = .执行部门ID: str执行科室 = Bill.TextMatrix(Bill.Row, Bill.Col)
                    .执行部门ID = Bill.ItemData(Bill.ListIndex)
                    Bill.TextMatrix(Bill.Row, Bill.Col) = Bill.CboText
                    If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModule, 1, 0, _
                        MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 2, gbytBilling, Bill.Row)) = False Then
                        Bill.Text = "": Bill.TxtVisible = False
                        Bill.cboObj.Text = str执行科室
                        .执行部门ID = lng执行科室: Exit Sub
                    End If
                End If
            End With
        End If
    End If
End Sub

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
        If mrsUnit Is Nothing Then Exit Sub
        If mrsUnit.State <> 1 Then Exit Sub
        If zlSelectDept(Me, mlngModule, Bill.cboObj, mrsUnit, Bill.CboText, True, , False) = False Then Exit Sub
    End With
    Exit Sub
End Sub

Private Sub Bill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Dim i As Long, bytsubs As Byte
    
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
            End If
        ElseIf MsgBox("确实要删除该收费项目吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = True: Exit Sub
        End If
        
        '删除处理
        For i = mobjBill.Details.Count To Row + 1 Step -1
            If mobjBill.Details(i).从属父号 = Row Then
                Call DeleteDetail(i) '反顺序删除其从属行
            End If
        Next
        Call DeleteDetail(Row) '删除该行
        
        '重新计算并刷新
        'Call CalcMoneys
        Call ShowDetails
        Call ShowMoney
        
        Bill.TxtVisible = False
        Bill.CmdVisible = False
        Bill.CboVisible = False
        
        Cancel = True '不用控件来处理了
    End If
End Sub

Private Sub Bill_CommandClick()
    Dim lng项目id As Long, blnCancel As Boolean
    Dim str特准项目 As String, int病人来源 As Integer, int险类 As Integer
    
    If mrsInfo.State = 1 Then
        If Not IsNull(mrsInfo!险类) Then
            int险类 = mrsInfo!险类
            '刘兴洪:24862
            If zl_Check特准项目(gclsInsure, int险类, Val(Nvl(mrsInfo!病人ID)), False) Then str特准项目 = Get保险特准项目(Val(Nvl(mrsInfo!病人ID)), "A.ID")
        End If
        If mrsInfo!病人性质 = 0 Or mrsInfo!病人性质 = 2 Then
            int病人来源 = 2
        ElseIf mrsInfo!病人性质 = 1 Or mrsInfo!病人性质 = -1 Then
            int病人来源 = 1
        End If
    Else
        int病人来源 = 2
    End If
    lng项目id = frmItemSelect.ShowSelect(Me, mstrPrivs, int病人来源, int险类, False, "'Z'", , , str特准项目, _
        , , , , , mstr药品价格等级, mstr卫材价格等级, mstr普通价格等级)
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

Private Sub Bill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
'功能：处理单据输入
    Dim objDetail As Detail, lng项目id As Long
    Dim str特准项目 As String, int病人来源 As Integer, int险类 As Integer
    Dim cur合计 As Currency, strScope As String
    Dim dblPreMoney As Double, i As Long, lngDoUnit As Long
    Dim cur余额 As Currency, curItemMoney As Currency
    
    On Error GoTo errH
    
    If Bill.ColData(Bill.Col) = 0 Then Exit Sub
    
    If KeyCode = 13 Then
        Select Case Bill.TextMatrix(0, Bill.Col)
            Case "项目"
                '此项目确定,该收费细目对应的程序对象才生成,同时这里处理收费从属项目
                If Bill.Text <> "" Then
                    If mblnSelect Then
                        mblnSelect = False '立即清除该标志
                        Set objDetail = GetInputDetail(Val(Bill.Text))
                    Else
                        If mrsInfo.State = 1 Then
                            If Not IsNull(mrsInfo!险类) Then
                                int险类 = mrsInfo!险类
                                '刘兴洪:24862
                                If zl_Check特准项目(gclsInsure, int险类, Val(Nvl(mrsInfo!病人ID)), False) Then str特准项目 = Get保险特准项目(Val(Nvl(mrsInfo!病人ID)), "A.ID")
                                
                            End If
                            If mrsInfo!病人性质 = 0 Or mrsInfo!病人性质 = 2 Then
                                int病人来源 = 2
                            ElseIf mrsInfo!病人性质 = 1 Or mrsInfo!病人性质 = -1 Then
                                int病人来源 = 1
                            End If
                        Else
                            int病人来源 = 2
                        End If
                        lng项目id = frmItemSelect.ShowSelect(Me, mstrPrivs, int病人来源, int险类, False, "'Z'", Bill.Text, _
                            Bill.TxtHwnd, , , , , , , mstr药品价格等级, mstr卫材价格等级, mstr普通价格等级)
                        If lng项目id <> 0 Then
                            Set objDetail = GetInputDetail(lng项目id)
                        Else
                            Bill.Text = "": Bill.TxtVisible = False
                            Bill.SetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    sta.Panels(2) = ""
                    Bill.TxtVisible = False '(不加不行)
                    
                    '医保病人费用审批
                    If mrsInfo.State = 1 Then
                        If Not IsNull(mrsInfo!险类) Then
                            If objDetail.要求审批 And Not mrsMedAudit Is Nothing Then
                                mrsMedAudit.Filter = "项目ID=" & objDetail.ID
                                If mrsMedAudit.RecordCount = 0 Then
                                    MsgBox "当前病人未被批准使用[" & objDetail.名称 & "]！", vbInformation, gstrSysName
                                    Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                                End If
                            End If
                            
                            '医保对码
                            If Not CheckMediCareItem(objDetail.ID, mrsInfo!险类, objDetail.名称, _
                                objDetail.变价 = False, , mstr普通价格等级) Then
                                Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                    '加入或修改该收费细目行
                    Call SetDetail(objDetail, Bill.Row)
                    Call CalcMoneys(Bill.Row)
                    
                    '输入摘要(根据新输入的行更改摘要)
                    Dim str摘要 As String '90304
                    If mobjBill.Details(Bill.Row).Detail.补充摘要 Then
                        If frmInputBox.InputBox(Me, "摘要", "请输入""" & mobjBill.Details(Bill.Row).Detail.名称 & """的摘要信息:", 200, 3, True, False, str摘要) Then
                            mobjBill.Details(Bill.Row).摘要 = str摘要
                        End If
                    Else
                         str摘要 = gclsInsure.GetItemInfo(0, mobjBill.病人ID, mobjBill.Details(Bill.Row).收费细目ID, str摘要, 1)
                         mobjBill.Details(Bill.Row).摘要 = str摘要
                    End If
                    
                    '记帐分类报警(在已经算出该行费用但未显示前)
                    mrsWarn.Filter = ""
                    If mrsWarn.RecordCount > 0 And mrsInfo.State = 1 And mobjBill.Details.Count = Bill.Row Then
                        cur合计 = GetBillTotal(mobjBill)
                        If cur合计 > 0 Then
                            cur余额 = Val(txt实收.Tag)
                            '刘兴洪:24491
                            curItemMoney = GetBillRowTotal(mobjBill.Details(Bill.Row).InComes)
                            
                            If gbln报警包含划价费用 Then cur余额 = Val(txt实收.Tag) - GetPriceMoneyTotal(1, mrsInfo!病人ID) + IIf(gbytBilling = 1, mcurModiMoney, 0)
                            gbytWarn = BillingWarn(mstrPrivsOpt, mrsInfo!姓名 & IIf(Nvl(mrsInfo!住院号) = "", "", "(住院号:" & mrsInfo!住院号 & " 床号:" & mrsInfo!床号 & ")"), Val("" & mrsInfo!病区ID), mrsInfo!适用病人, mrsWarn, cur余额, mrsInfo!当日额 - mcurModiMoney, cur合计, IIf(IsNull(mrsInfo!担保额), 0, mrsInfo!担保额), mobjBill.Details(Bill.Row).收费类别, mobjBill.Details(Bill.Row).Detail.类别名称, mstrWarn, , , curItemMoney)
                            If gbytWarn = 2 Or gbytWarn = 3 Then
                                mobjBill.Details.Remove Bill.Row '删除刚刚想要加入的费用行
                                Bill.Text = "": Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                    If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModule, 1, 0, _
                        MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 2, gbytBilling, Bill.Row)) = False Then
                        mobjBill.Details.Remove Bill.Row '删除刚刚想要加入的费用行
                        Bill.Text = "": Bill.TxtVisible = False
                        Cancel = True: Exit Sub
                    End If
                    
                    Call ShowDetails(Bill.Row)
                    Call ShowMoney
                    
                    '费用类型检查
                    Call Check费用类型(Bill.Row)
                    
                    Bill.Text = "": Bill.SetFocus
                End If
                
                If mobjBill.Details.Count >= Bill.Row Then
                    '下一列的性质确定
                    If mobjBill.Details(Bill.Row).Detail.变价 Then Bill.ColData(1) = 4 '应收金额
                    
                    '执行科室!!!
                    Call FillBillComboBox(Bill.Row, 3)
                    If Bill.ListCount = 1 Then
                        Bill.ColData(3) = 5
                        mobjBill.Details(Bill.Row).Key = 1
                    Else
                        Bill.ColData(3) = 3
                        mobjBill.Details(Bill.Row).Key = Bill.ListCount
                    End If
                    
                    '从属项目处理(在这里可以处理多级从属-从属的从属...)
                    If Bill.TextMatrix(0, Bill.Col) = "项目" Then
                        If ShouldDO(Bill.Row) Then
                            Set mcolDetails = New Details
                            Set mcolDetails = GetSubDetails(mobjBill.Details(Bill.Row).收费细目ID)
                            For i = 1 To mcolDetails.Count
                                If mobjBill.Details.Count >= Bill.Rows - 1 Then
                                    Bill.Rows = Bill.Rows + 1
                                    Call bill_AfterAddRow(Bill.Rows - 1)
                                End If
                                Bill.TextMatrix(Bill.Rows - 1, 0) = "" '有必要加上
                                
                                If mcolDetails(i).类别 = mobjBill.Details(Bill.Row).收费类别 Then
                                    '1.从项收费类别与主项相同的,缺省与主项执行科室相同。
                                    lngDoUnit = mobjBill.Details(Bill.Row).执行部门ID
                                Else
                                    If mcolDetails(i).执行科室 = 0 Then
                                        '2.从项设置为无明确科室的,缺省与主项执行科室相同。
                                        lngDoUnit = mobjBill.Details(Bill.Row).执行部门ID
                                    End If
                                        '其余情况,取本身设置的执行科室
                                End If
                                
                                Call SetDetail(mcolDetails(i), Bill.Rows - 1, Bill.Row, lngDoUnit)
                                Call CalcMoneys(Bill.Rows - 1)
                                Call ShowDetails(Bill.Rows - 1)
                                Call ShowMoney
                            Next
                        End If
                    End If
                End If
            Case "应收金额" '实际上是单价(因为数据次缺省为1,且不能更改)
                If mobjBill.Details.Count >= Bill.Row And Bill.Text <> "" Then
                    '数值合法性
                    If Not IsNumeric(Bill.Text) Then
                        MsgBox "非法数值！", vbInformation, gstrSysName
                        Bill.Text = "": Cancel = True: Bill.TxtVisible = False: Exit Sub
                    End If
                    '负数权限
                    If CDbl(Bill.Text) < 0 Then
                        If InStr(mstrPrivsOpt, ";诊疗负数记帐;") = 0 Then
                            MsgBox "你没有权限输入负数！", vbInformation, gstrSysName
                            Bill.Text = "": Cancel = True: Bill.TxtVisible = False: Exit Sub
                        Else
                            If mrsInfo.State = 1 Then
                                If Not IsNull(mrsInfo!险类) Then
                                    If Not gclsInsure.GetCapability(support负数记帐, mrsInfo!病人ID, mrsInfo!险类) Then
                                        MsgBox "本地医保不支持对医保病人进行负数记帐！", vbInformation, gstrSysName
                                        Bill.Text = "": Cancel = True: Bill.TxtVisible = False: Exit Sub
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                    Bill.Text = Format(Bill.Text, gstrDec)
                    
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
                        
                        '变价项目只对应一个收入项目
                        dblPreMoney = mobjBill.Details(Bill.Row).数次 * mobjBill.Details(Bill.Row).InComes(1).标准单价
                        mobjBill.Details(Bill.Row).数次 = Sgn(Val(Bill.Text))
                        mobjBill.Details(Bill.Row).InComes(1).标准单价 = Abs(Val(Bill.Text))
                        Call CalcMoneys(Bill.Row)
                        
                        '记帐分类报警(在已经算出该行费用但未显示前)
                        mrsWarn.Filter = ""
                        If mrsWarn.RecordCount > 0 And mrsInfo.State = 1 Then
                            cur合计 = GetBillTotal(mobjBill)
                            If cur合计 > 0 Then
                                cur余额 = Val(txt实收.Tag)
                                '刘兴洪:24491
                                curItemMoney = GetBillRowTotal(mobjBill.Details(Bill.Row).InComes)
                                If gbln报警包含划价费用 Then cur余额 = Val(txt实收.Tag) - GetPriceMoneyTotal(1, mrsInfo!病人ID) + IIf(gbytBilling = 1, mcurModiMoney, 0)
                                gbytWarn = BillingWarn(mstrPrivsOpt, mrsInfo!姓名 & IIf(Nvl(mrsInfo!住院号) = "", "", "(住院号:" & mrsInfo!住院号 & " 床号:" & mrsInfo!床号 & ")"), Val("" & mrsInfo!病区ID), mrsInfo!适用病人, mrsWarn, cur余额, mrsInfo!当日额 - mcurModiMoney, cur合计, IIf(IsNull(mrsInfo!担保额), 0, mrsInfo!担保额), mobjBill.Details(Bill.Row).收费类别, mobjBill.Details(Bill.Row).Detail.类别名称, mstrWarn, , , curItemMoney)
                                If gbytWarn = 2 Or gbytWarn = 3 Then
                                    mobjBill.Details(Bill.Row).数次 = Sgn(dblPreMoney)
                                    mobjBill.Details(Bill.Row).InComes(1).标准单价 = Abs(dblPreMoney)
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
                If mobjBill.Details.Count >= Bill.Row Then
                    If Bill.ListIndex <> -1 Then
                        'If mobjBill.Details(Bill.Row).执行部门ID <> Bill.ItemData(Bill.ListIndex) Then
                            mobjBill.Details(Bill.Row).执行部门ID = Bill.ItemData(Bill.ListIndex)
                            If ItemHaveSub(Bill.Row) Then Call SetSubDept(Bill.Row)
                        'End If
                        If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModule, 1, 0, _
                            MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 2, gbytBilling, Bill.Row)) = False Then
                            Bill.Text = "": Bill.TxtVisible = False
                            Cancel = True: Exit Sub
                        End If
                    End If
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

Private Sub SetSubDept(ByVal lngRow As Long)
    Dim i As Long, j As Long
    
    For i = lngRow + 1 To mobjBill.Details.Count
        If mobjBill.Details(i).从属父号 = lngRow Then
            '从属项为药品和卫材的项目的执行科室不随主项变动
            If InStr(",4,5,6,7,", mobjBill.Details(i).收费类别) = 0 Then
                With mobjBill
                    If .Details(i).收费类别 = .Details(lngRow).收费类别 Then
                        '1.从项收费类别与主项相同的,缺省与主项执行科室相同。
                        .Details(i).执行部门ID = .Details(lngRow).执行部门ID
                    Else
                        Set mcolDetails = GetSubDetails(.Details(lngRow).收费细目ID) '必须现取
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
                                .Details(i).执行部门ID = Get收费执行科室ID(mcolDetails(j).类别, _
                                    mcolDetails(j).ID, mcolDetails(j).执行科室, .Details(i).执行部门ID, Get开单科室ID, Get病人来源, , mobjBill.病区ID)
                            End If
                        End If
                    End If
                    
                    If .Details(i).执行部门ID > 0 Then
                        If mbytInState = 0 Then
                            mrsUnit.Filter = "ID=" & .Details(i).执行部门ID
                            If mrsUnit.RecordCount <> 0 Then
                                Bill.TextMatrix(i, 3) = IIf(zlIsShowDeptCode, mrsUnit!编码 & "-", "") & mrsUnit!名称
                            Else
                                Bill.TextMatrix(i, 3) = GET部门名称(.Details(i).执行部门ID, mrsUnit)
                            End If
                        Else
                            '浏览单据只(能)显示名称
                            Bill.TextMatrix(i, 3) = GET部门名称(.Details(i).执行部门ID, mrsUnit)
                        End If
                    End If
                End With
            End If
        End If
    Next
    
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
    Dim i As Long
    
    If Not Bill.Active Then Exit Sub
    
    If mbytInState = 3 Or (chkCancel.Visible And chkCancel.Value = 1) Then
        '针对列编辑性质设置颜色
        Bill.SetColColor 0, &HE7CFBA '不然要成白色
        Exit Sub
    End If
    
    If mbytInState = 0 Then
        For i = 0 To UBound(marrColData)
            Bill.ColData(i) = marrColData(i)
        Next
    End If

    '如果是从属项目的主项目或从项,则不允许更改类别和项目
    If mobjBill.Details.Count >= Row Then
    If ItemHaveSub(Row) Or mobjBill.Details(Row).从属父号 > 0 Then
        Bill.ColData(0) = BillColType.Text_UnModify
    End If
    End If
    
    If mobjBill.Details.Count >= Bill.Row And mbytInState <> 2 Then
        If mobjBill.Details(Bill.Row).Key = "1" Then
            Bill.ColData(3) = 5
        Else
            Bill.ColData(3) = 3
        End If
    End If
    If Bill.ColData(Bill.Col) = 3 Then Call FillBillComboBox(Bill.Row, Bill.Col)
    
    Bill.TextLen = 0: Bill.TextMask = ""
    Select Case Bill.TextMatrix(0, Col)
        Case "执行科室"
            SetWidth Bill.cboHwnd, 130
        Case "应收金额"
            Bill.TextLen = 10
            If InStr(mstrPrivsOpt, ";诊疗负数记帐;") = 0 Then
                Bill.TextMask = "0123456789." & Chr(8)
            Else
                Bill.TextMask = "-0123456789." & Chr(8)
            End If
            
            If InStr(Bill.TextMask, "-") > 0 And mrsInfo.State = 1 Then
                If Not IsNull(mrsInfo!险类) Then
                    If Not gclsInsure.GetCapability(support负数记帐, mrsInfo!病人ID, mrsInfo!险类) Then
                        Bill.TextMask = Replace(Bill.TextMask, "-", "")
                    End If
                End If
            End If
    End Select

    '进入行时,重新设置该行的编辑性质
    If mobjBill.Details.Count >= Bill.Row Then
        If mobjBill.Details(Bill.Row).Detail.变价 Then
            Bill.ColData(1) = 4
        Else
            Bill.ColData(1) = 5
        End If
    End If
End Sub

Private Sub Bill_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Bill.ToolTipText = Bill.TextMatrix(Bill.MouseRow, Bill.MouseCol)
End Sub



Private Sub cboBaby_Click()
    mobjBill.婴儿费 = cboBaby.ItemData(cboBaby.ListIndex)
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

Private Sub cbo开单科室_Click()
    Dim i As Long, lng开单部门ID As Long
    
    If Not (mbytInState = 0 And chkCancel.Value = 0) Then Exit Sub
    If cbo开单科室.ListIndex <> -1 Then lng开单部门ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
    If mobjBill.开单部门ID = lng开单部门ID Then Exit Sub
    mobjBill.开单部门ID = lng开单部门ID
        
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
        
    '重新设置相关项目的执行科室
    If cbo开单科室.ListIndex <> -1 And cbo开单科室.Visible Then
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
                '仅处理收费项目
                If .Detail.执行科室 = 6 Then '6-开单人科室
                    .执行部门ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
                    '刷新显示从项执行科室
                    If i <= Bill.Rows - 1 And .执行部门ID <> 0 Then
                        If mbytInState = 0 Then
                            mrsUnit.Filter = "ID=" & .执行部门ID
                            If mrsUnit.RecordCount <> 0 Then
                                Bill.TextMatrix(i, BillCol.执行科室) = IIf(zlIsShowDeptCode, mrsUnit!编码 & "-", "") & mrsUnit!名称
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
    
End Sub

Private Sub cbo开单科室_Validate(Cancel As Boolean)
    If cbo开单科室.Text <> "" And cbo开单科室.ListIndex < 0 Then cbo开单科室.Text = ""
End Sub

Private Sub cbo开单人_Click()
    Dim lng开单人ID As Long
    
    If Not (mbytInState = 0 And chkCancel.Value = 0) Then Exit Sub
    If mobjBill.开单人 = IIf(cbo开单人.ListIndex = -1, "", zlStr.NeedName(cbo开单人.Text)) Then Exit Sub
    
    mobjBill.开单人 = IIf(cbo开单人.ListIndex = -1, "", zlStr.NeedName(cbo开单人.Text))
    If gblnFromDr Then
        If cbo开单人.ListIndex <> -1 Then
            lng开单人ID = cbo开单人.ItemData(cbo开单人.ListIndex)
            
            Call FillDept(cbo开单科室, mrs开单科室, mrs开单人, mstrPrivs, mbytUseType, mlngDeptID, lng开单人ID)
            Call SetDefaultDept(cbo开单科室, mrs开单科室, mrs开单人, lng开单人ID)
        Else
            cbo开单科室.Clear
        End If
        Call cbo开单科室_Click
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
    If gblnFromDr And gbln开单人 And cbo开单人.ListIndex = -1 And txtPatient.Text <> "" Then Cancel = True
End Sub

Private Sub chkCancel_Click()
    Dim i As Long
    
    mstrInNO = ""
    Call NewBill
    Call ClearRows
    Call Bill.ClearBill
    Call ClearMoney
    
    Bill.AllowAddRow = (chkCancel.Value = 0)
    
    If chkCancel.Value = 1 Then
        chkCancel.ForeColor = &HFF&
        
        fraInfo.Enabled = False
        fraAppend.Enabled = False
        
        For i = 0 To UBound(marrColData)
            Bill.ColData(i) = 0
        Next
        Call ShowDeleteCol(True)
        Bill.SetColColor 0, &HE7CFBA '不然要成白色
        Bill.Active = True
        
        Call SetDisible
        cboNO.Locked = False
        cboNO.SetFocus
    Else
        chkCancel.ForeColor = 0
        Call cbo开单科室_Click
        
        For i = 0 To UBound(marrColData)
            Bill.ColData(i) = marrColData(i)
        Next
        Call ShowDeleteCol(False)
        Bill.SetColColor 0, &HE7CFBA '不然要成白色
        
        If gbytBilling = 2 Then
            Call SetDisible
            Bill.Active = False
            cboNO.Locked = False
            cboNO.SetFocus
        Else
            Call SetDisible(True)
            fraInfo.Enabled = True
            fraAppend.Enabled = True
            Bill.Active = True
            cboNO.Locked = True
            If mbytUseType = 1 And mlng病人ID > 0 Then
                txtPatient.Text = "-" & mlng病人ID
                Call txtPatient_KeyPress(13)
                Bill.SetFocus
            Else
                txtPatient.SetFocus
            End If
        End If
    End If
End Sub

Private Sub chk加班_Click()
    If mbytInState = 1 Or chkCancel.Value = 1 Or gbytBilling = 2 Then Exit Sub
    If Not chk加班.Visible Then Exit Sub
    
    Dim blnAdd As Boolean
    
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
    Call ShowMoney
End Sub

Private Sub chk加班_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    If (mobjBill.Details.Count > 0 Or txtPatient.Text <> "") And Bill.Active And mbytInState = 0 And mstrInNO = "" Then
        
        If MsgBox("确实要清除当前单据中的内容吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
        
        If chkCancel.Value = Checked Then '退据单状态
            Call ClearRows: Call Bill.ClearBill
            
            chkCancel.Value = Unchecked
            Call NewBill
            Call SetDisible(True)
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        ElseIf Bill.Active Then '正常输入单据状态'(清除后当作是新病人单据)
            Call ClearRows: Call Bill.ClearBill
            
            Call NewBill   '保持原单据号
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        End If
        
        Exit Sub
    End If
    
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strInfo As String, strSQL As String
    Dim curTotal As Currency, i As Long
    Dim lng结帐ID As Long, cur当日额 As Currency, cur余额 As Currency
    Dim intInsure As Integer, Curdate As Date, blnTrans As Boolean
    Dim str销帐申请IDs As String, str申请人s As String, cllPro As Collection
    Dim rsItems As ADODB.Recordset
    If mbytInState = 3 Or (mbytInState = 0 And chkCancel.Visible And chkCancel.Value = 1) Then
        If mbytInState = 0 And mstrInNO = "" Then
            MsgBox "没有读取单据内容,不能销帐！", vbInformation, gstrSysName
            cboNO.SetFocus: Exit Sub
        End If
        
        For i = 1 To Bill.Rows - 1
            If Bill.TextMatrix(i, Bill.Cols - 1) = "√" And Bill.RowData(i) > 0 Then
                strSQL = strSQL & "," & Bill.RowData(i)
            End If
        Next
        If strSQL = "" Then
            MsgBox "请至少选择一行要销帐的费用！", vbInformation, gstrSysName
            Bill.SetFocus: Exit Sub
        End If
        
        '所有行选择处理
        strSQL = Mid(strSQL, 2)
        i = GetBillRows(mstrInNO, 2)
        If UBound(Split(strSQL, ",")) + 1 = i Then strSQL = ""
        
        
        If strSQL <> "" And InStr(1, mstrPrivsOpt, ";部分销帐;") = 0 Then
            MsgBox "你没有部分销帐的权限，只能对该单据全部销帐！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If zlCheckIsExistsApplied(mstrInNO, strSQL, str销帐申请IDs, str申请人s) Then
            '问题:47416
            If MsgBox("注意:" & vbCrLf & "    单据" & mstrInNO & "中存在申请销帐的项目,销帐后,将会自动取消" & vbCrLf & "申请人的申请项目,是否继续销帐?" & vbCrLf & "申请人如下: " & str申请人s, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        
        '医保记帐作废上传(注意判断顺序)
        If gbytBilling = 0 Then
            intInsure = BillExistInsure(mstrInNO, mstrTime) '判断是否医保病人记的帐
            If intInsure > 0 Then
                If gclsInsure.GetCapability(support记帐作废上传, , intInsure) Then
                    '去掉了医保连接匹配检查
                    If Not gclsInsure.GetCapability(support允许部份冲销单据, , intInsure) And strSQL <> "" Then '不能部分作废
                        MsgBox "因为医保处理需要,该单据中的项目必须全部销帐！", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
            End If
        End If
         '问题:47416
        Set cllPro = New Collection
        If str销帐申请IDs <> "" Then
            strSQL = "zl_病人费用销帐_Delete('" & str销帐申请IDs & "')"
            zlAddArray cllPro, strSQL
        End If
        strSQL = "zl_住院记帐记录_DELETE('" & mstrInNO & "','" & strSQL & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
        zlAddArray cllPro, strSQL
        
        On Error GoTo errH
         blnTrans = True
         zlExecuteProcedureArrAy cllPro, Me.Caption, True
            '医保记帐作废上传
            If gbytBilling = 0 And intInsure <> 0 Then
                If gclsInsure.GetCapability(support记帐作废上传, , intInsure) And Not gclsInsure.GetCapability(support记帐完成后上传, , intInsure) Then
                    If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                        gcnOracle.RollbackTrans: Exit Sub
                    End If
                End If
            End If
        gcnOracle.CommitTrans: blnTrans = False
        
        '医保记帐作废上传
        If gbytBilling = 0 And intInsure <> 0 Then
            If gclsInsure.GetCapability(support记帐作废上传, , intInsure) And gclsInsure.GetCapability(support记帐完成后上传, , intInsure) Then
                If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                    MsgBox "单据""" & mstrInNO & """的销帐数据向医保传送失败，该单据已销帐。", vbInformation, gstrSysName
                End If
            End If
        End If
        
        If mbytInState = 0 Then
            mstrInNO = "": cboNO.Text = ""
            txtPatient.Text = "": txtOld.Text = ""
            txt实收.Text = gstrDec: txt应收.Text = gstrDec
            Call ClearRows: Call Bill.ClearBill
            Call ClearMoney: Call NewBill
            Call SetMoneyList
            chkCancel.Value = 0
            If gbytBilling = 2 Then
                cboNO.SetFocus
            Else
                txtPatient.SetFocus
            End If
        Else
            Unload Me
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
        Unload Me
    ElseIf Bill.Active And chkCancel.Value = 0 Then '正常输入单据状态
        If mrsInfo.State = adStateClosed Then
            MsgBox "没有发现病人信息,请确定病人信息！", vbInformation, gstrSysName
            txtPatient.SetFocus: Exit Sub
        End If
        If txt费别.Text = "" Or mobjBill.费别 = "" Then
            MsgBox "请选择病人费别！", vbInformation, gstrSysName
            txt费别.SetFocus: Exit Sub
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
        
        If mobjBill.开单部门ID = 0 Then
            MsgBox "请确定开单科室！", vbInformation, gstrSysName
            cbo开单科室.SetFocus
            Exit Sub
        End If
        
        If Not IsDate(txtDate.Text) Then
            MsgBox "请输入正确的费用日期！", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        strInfo = Check发生时间(CDate(txtDate.Text), mrsInfo!病人ID)
        If strInfo <> "" Then
            MsgBox strInfo, vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        
        If mobjBill.开单人 = "" And gbln开单人 Then
            MsgBox "请输入开单人！", vbInformation, gstrSysName
            cbo开单人.SetFocus: Exit Sub
        End If
        
        '出院强制记帐权限检查
        If Not PatiCanBilling(mrsInfo!病人ID, Nvl(mrsInfo!主页ID, 0), mstrPrivsOpt) Then Exit Sub
        
        If zlPatiIS病案已编目(mrsInfo!病人ID, Nvl(mrsInfo!主页ID, 0)) = True Then     '问题:28725
            Exit Sub
        End If
        
        '49501
        If zlIsAllowFeeChange(Val(Nvl(mrsInfo!病人ID)), Val(Nvl(mrsInfo!主页ID))) = False Then
            Exit Sub
        End If
        
        '发生时间检查
        If Not IsNull(mrsInfo!出院日期) Then
            If Format(txtDate.Text, txtDate.Format) > Format(mrsInfo!出院日期, txtDate.Format) Then
                MsgBox "强制对出院病人记帐时，费用时间不能大于病人出院时间:" & Format(mrsInfo!出院日期, txtDate.Format), vbInformation, gstrSysName
                txtDate.SetFocus: Exit Sub
            End If
        End If
        If Not IsNull(mrsInfo!险类) And Not IsNull(mrsInfo!入院日期) Then
            If Format(txtDate.Text, txtDate.Format) < Format(mrsInfo!入院日期, txtDate.Format) Then
                MsgBox "费用的发生时间不能小于医保病人的入院时间:" & Format(mrsInfo!入院日期, txtDate.Format), vbInformation, gstrSysName
                txtDate.SetFocus: Exit Sub
            End If
        End If
        
        '非法行
        For i = 1 To mobjBill.Details.Count
            If mobjBill.Details(i).收费细目ID = 0 Then
                MsgBox "单据中第 " & i & " 行没有正确输入数据,请修正或删除该行！", vbInformation, gstrSysName
                Bill.SetFocus: Exit Sub
            End If
        Next
        
        '医保负数记帐检查    因为操作员可能先输单据,再确定病人,所以要再检查一次
        If InStr(mstrPrivsOpt, ";诊疗负数记帐;") > 0 Then
            If Not IsNull(mrsInfo!险类) Then
                If Not gclsInsure.GetCapability(support负数记帐, mrsInfo!病人ID, mrsInfo!险类) Then
                    For i = 1 To mobjBill.Details.Count
                        If mobjBill.Details(i).数次 * mobjBill.Details(i).付数 < 0 Then
                                MsgBox "单据中第 " & i & " 行是负数,本地医保不支持负数记帐！", vbInformation, gstrSysName
                                Bill.SetFocus: Exit Sub
                        End If
                    Next
                End If
            End If
        End If
        
        '要求审批,检查医保审批
        If Not IsNull(mrsInfo!险类) And Not mrsMedAudit Is Nothing Then
           If Not CheckExamine(mobjBill.Details, mrsMedAudit, mrsInfo!险类) Then Exit Sub
        End If
        
        
        '费用类型检查
        If Not Check费用类型 Then Exit Sub
        
        '记帐分类报警(只有一种类别)
        mrsWarn.Filter = ""
        If mrsWarn.RecordCount > 0 Then
            curTotal = GetBillTotal(mobjBill)
            If curTotal > 0 Then
                '刷新病人预交款信息
                Set rsTmp = GetMoneyInfo(mrsInfo!病人ID, IIf(gbytBilling = 0, mcurModiMoney, 0), True, 2)
                If Not rsTmp Is Nothing Then
                    cmdOK.Tag = rsTmp!预交余额
                    cmdCancel.Tag = rsTmp!费用余额
                    txt实收.Tag = rsTmp!预交余额 - rsTmp!费用余额
                Else
                    cmdOK.Tag = 0: cmdCancel.Tag = 0: txt实收.Tag = 0
                End If
                '划价时显示不算当前单据费用,但划价报警要算
                sta.Panels(3).Text = "预交:" & Format(Val(cmdOK.Tag), "0.00")
                sta.Panels(3).Text = sta.Panels(3).Text & "/费用:" & Format(Val(cmdCancel.Tag) + IIf(gbytBilling = 0, curTotal, 0), gstrDec)
                sta.Panels(3).Text = sta.Panels(3).Text & "/剩余:" & Format(Val(txt实收.Tag) - IIf(gbytBilling = 0, curTotal, 0), "0.00")
                
                
                '重新读取当日额
                cur当日额 = GetPatiDayMoney(mrsInfo!病人ID)
                
                cur余额 = Val(txt实收.Tag)
                If gbln报警包含划价费用 Then cur余额 = Val(txt实收.Tag) - GetPriceMoneyTotal(1, mrsInfo!病人ID) + IIf(gbytBilling = 1, mcurModiMoney, 0)
                
                gbytWarn = BillingWarn(mstrPrivsOpt, mrsInfo!姓名 & IIf(Nvl(mrsInfo!住院号) = "", "", "(住院号:" & mrsInfo!住院号 & " 床号:" & mrsInfo!床号 & ")"), Val("" & mrsInfo!病区ID), Val(txt医疗付款.Tag) = 1 Or Not IsNull(mrsInfo!险类), _
                    mrsWarn, cur余额, cur当日额 - mcurModiMoney, curTotal, Nvl(mrsInfo!担保额, 0), "Z", "其它", mstrWarn)
                If gbytWarn = 2 Or gbytWarn = 3 Then Exit Sub
            End If
        End If
        
        '项目服务对象检查(主要因为多了门诊留观病人)
        If Check服务对象 > 0 Then Exit Sub
        
        If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModule, 1, 1, _
            MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), 2, gbytBilling)) = False Then
            Exit Sub
        End If
        
        If IsDate(txtDate.Text) Then mobjBill.发生时间 = CDate(txtDate.Text)
        mobjBill.登记时间 = zlDatabase.Currentdate
        If zlGetSaveDataItems_Plugin(mobjBill, rsItems) = False Then Exit Sub
        If zlChargeSaveValied_Plugin(mlngModule, 2, False, gbytBilling = 1, "", rsItems) = False Then Exit Sub
        '保存
        If Not SaveBill Then
            Exit Sub
        Else
            Call zlChargeSaveAfter_Plugin(mlngModule, mobjBill.病人ID, mobjBill.主页ID, False, 2, mobjBill.NO)
            If gbytBilling = 0 And gbln记帐打印 Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_113" & 3 + mbytUseType, Me, "NO=" & mobjBill.NO, "登记时间=" & Format(mobjBill.登记时间, "yyyy-MM-dd HH:mm:ss"), "药品单位=0", "PrintEmpty=0", "重打=0", 2)
            ElseIf gbytBilling = 1 And gbln划价打印 Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1133", Me, "NO=" & mobjBill.NO, "登记时间=" & Format(mobjBill.登记时间, "yyyy-MM-dd HH:mm:ss"), "药品单位=0", "PrintEmpty=0", "重打=0", 2)
            End If
            
            If mstrInNO = "" Then
                sta.Panels(2) = "上一张单据:" & mobjBill.NO
                Call ClearRows: Call Bill.ClearBill
                Call ClearMoney
                Call SetMoneyList
                mstrInNO = ""
                
                If mrsInfo.State = 1 Then
                    Call NewBill(False)
                    txtPatient.Tag = "-" & mrsInfo!病人ID
                    
                    With mobjBill
                        .病人ID = IIf(IsNull(mrsInfo!病人ID), 0, mrsInfo!病人ID)
                        .主页ID = IIf(IsNull(mrsInfo!主页ID), 0, mrsInfo!主页ID)
                        
                        .病区ID = IIf(IsNull(mrsInfo!病区ID), 0, mrsInfo!病区ID)
                        .科室ID = IIf(IsNull(mrsInfo!科室ID), 0, mrsInfo!科室ID)
                        
                        .床号 = "" & mrsInfo!床号
                        .标识号 = IIf(IsNull(mrsInfo!住院号), 0, mrsInfo!住院号)
                        .姓名 = IIf(IsNull(mrsInfo!姓名), "", mrsInfo!姓名)
                        .性别 = IIf(IsNull(mrsInfo!性别), "", mrsInfo!性别)
                        .年龄 = IIf(IsNull(mrsInfo!年龄), "", mrsInfo!年龄)
                        .费别 = IIf(IsNull(mrsInfo!费别), "", mrsInfo!费别)
                        
                        .婴儿费 = cboBaby.ItemData(cboBaby.ListIndex)
                        .开单人 = IIf(cbo开单人.ListIndex = -1, "", zlStr.NeedName(cbo开单人.Text))
                    End With
                    
                    If mbytUseType = 1 Then
                        Call txtPatient_KeyPress(13) '刷新一些费用信息
                        Bill.SetFocus
                    Else
                        txtPatient.SetFocus
                    End If
                Else
                    Call NewBill
                    txtPatient.SetFocus
                End If
            Else '修改
                Unload Me
            End If
        End If
    ElseIf Not Bill.Active Then '审核住院划价状态
        If mstrInNO = "" Then
            MsgBox "没有住院划价单据,请先输入！", vbInformation, gstrSysName
            cboNO.SetFocus: Exit Sub
        End If
        
        '取本次审核的行序号
        strSQL = ""
        For i = 1 To Bill.Rows - 1
            If Bill.RowData(i) > 0 Then
                strSQL = strSQL & "," & Bill.RowData(i)
            End If
        Next
        strSQL = Mid(strSQL, 2)
        i = GetBillRows(mstrInNO, 2)
        If UBound(Split(strSQL, ",")) + 1 = i Then strSQL = ""
        
        '医保检查
        intInsure = BillExistInsure(mstrInNO, , True)
        If intInsure > 0 Then
            '去掉了医保连接匹配检查
        End If
        
        '费用报警
        mrsWarn.Filter = ""
        If mrsWarn.RecordCount > 0 Then
            If Not AuditingWarn(mstrPrivsOpt, mrsWarn, mstrInNO, strSQL) Then Exit Sub
        End If
        
        Curdate = zlDatabase.Currentdate
        strSQL = "zl_住院记帐记录_Verify('" & mstrInNO & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "','" & strSQL & "',NULL,To_Date('" & Format(Curdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
        
        On Error GoTo errH
        gcnOracle.BeginTrans: blnTrans = True
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            
            '医保上传
            If intInsure <> 0 Then
                '医保传输费用明细
                If gclsInsure.GetCapability(support记帐上传, , intInsure) And Not gclsInsure.GetCapability(support记帐完成后上传, , intInsure) Then
                    strInfo = ""
                    If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 1, strInfo, , intInsure) Then
                        gcnOracle.RollbackTrans
                        If strInfo <> "" Then MsgBox strInfo, vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
            End If
        gcnOracle.CommitTrans: blnTrans = False
        
        '医保上传
        If intInsure <> 0 Then
            '医保传输费用明细
            If gclsInsure.GetCapability(support记帐上传, , intInsure) And gclsInsure.GetCapability(support记帐完成后上传, , intInsure) Then
                strInfo = ""
                If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 1, strInfo, , intInsure) Then
                    If strInfo <> "" Then
                        MsgBox strInfo, vbInformation, gstrSysName
                    Else
                        MsgBox "单据""" & mstrInNO & """的数据向医保传送失败,该单据已审核！", vbInformation, gstrSysName
                    End If
                    Exit Sub
                End If
            End If
        End If
        
        On Error GoTo 0
        
        If gbytBilling = 2 And gbln审核打印 And mblnPrint Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1133", Me, "NO=" & mstrInNO, "登记时间=" & Format(Curdate, "yyyy-MM-dd HH:mm:ss"), "药品单位=0", "PrintEmpty=0", "重打=0", 2)
        End If
        
        mstrInNO = "": cboNO.Text = ""
        txtPatient.Text = "": txtOld.Text = ""
        txt实收.Text = gstrDec: txt应收.Text = gstrDec
        Call ClearRows: Call Bill.ClearBill
        Call ClearMoney: Call NewBill
        Call SetMoneyList
        cboNO.Locked = False: cboNO.SetFocus
    End If
    gblnOK = True
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdOK_GotFocus()
    If mbytInState = 3 Or (chkCancel.Visible And chkCancel.Value = 1) Then
        Bill.Row = 1: Bill.Col = Bill.Cols - 1
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If mbytUseType = 1 And mlng病人ID <> 0 And mbytInState = 0 Then
        If gblnFromDr Then
            cbo开单人.SetFocus
        Else
            Bill.SetFocus
        End If
    ElseIf gbytBilling = 2 Then
        cboNO.SetFocus
    ElseIf mbytInState = 1 Then
        cmdCancel.SetFocus
    ElseIf mbytInState = 2 Then
        txtDate.SetFocus
    ElseIf mbytInState = 3 Then
        cmdOK.SetFocus
    ElseIf mstrInNO <> "" Then
        Bill.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("',|~" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim i As Long, strPre As String, lngPre As Long, strTmp As String
    
    RestoreWinState Me, App.ProductName, mbytInState
    
    gblnOK = False: mblnValid = False
    mstrPrivsOpt = GetInsidePrivs(Enum_Inside_Program.p记帐操作)
    Call initCardSquareData
    '初始化单据数据
    Set mobjBill = New ExpenseBill
    
    If mbytInState = 0 Or mbytInState = 2 Or mbytInState = 3 Then
        If Not InitData Then Unload Me: Exit Sub
    Else
        If Init开单人开单科室(cbo开单人, cbo开单科室, mrs开单人, mrs开单科室, mstrPrivs, mbytUseType, mlngDeptID) = False Then
            Exit Sub
        End If
    End If
    mstrUnitIDs = GetUserUnits
    
    Call InitFace
    Call NewBill
    
    If mbytInState <> 0 Then '显示、调整、销帐单据(1,2,3)
        If Not ReadBill(mstrInNO, (mbytInState = 3)) Then Unload Me: Exit Sub
        cboNO.Text = mstrInNO
    Else '新增
        mstr药品价格等级 = gstr药品价格等级
        mstr卫材价格等级 = gstr卫材价格等级
        mstr普通价格等级 = gstr普通价格等级
        '读取该单据的内容
        If mstrInNO <> "" Then '修改单据
            Set mobjBill = ImportBill(mstrInNO, False, Me, True, True, , , , mstr药品价格等级, mstr卫材价格等级, mstr普通价格等级)
            If mobjBill.NO = "" Then
                MsgBox "读取单据失败。", vbInformation, gstrSysName
                Unload Me: Exit Sub
            Else
                mcurModiMoney = GetBillMoney(2, mobjBill.NO) '要在读取病人信息前先读
                
                lngPre = mobjBill.开单部门ID
                strPre = mobjBill.开单人
                
                txtPatient.Text = "-" & mobjBill.病人ID
                Call txtPatient_KeyPress(13)
                                
                Call ReCalcInsure '重新计算统筹金额
                
                '显示的是原单据号,保存的是新单据号
                cboNO.Text = mobjBill.NO
                Bill.ClearBill
                Bill.Rows = mobjBill.Details.Count + 1
                '针对列编辑性质设置颜色
                Bill.SetColColor 0, &HE7CFBA
                Bill.SetColColor 1, &HE7CFBA
                Bill.SetColColor 3, &HE7CFBA
                
                txtDate.Text = Format(mobjBill.发生时间, "yyyy-MM-dd HH:mm:ss")
                chk加班.Value = mobjBill.加班标志
                
                mobjBill.开单部门ID = lngPre
                mobjBill.开单人 = strPre
                Call Set开单人开单科室(cbo开单人, cbo开单科室, mrs开单人, mrs开单科室, mobjBill.开单人, mobjBill.开单部门ID)
                Call zlControl.CboLocate(cboBaby, mobjBill.婴儿费, True)
                                
                '修改时应保存当前操作员的名字
                mobjBill.操作员编号 = UserInfo.编号
                mobjBill.操作员姓名 = UserInfo.姓名
                
                Call ShowDetails
                Call ShowMoney
                
                For i = 1 To mobjBill.Details.Count
                    '特殊处理
                    Bill.RowData(i) = Asc(mobjBill.Details(i).收费类别)
                Next
            End If
        Else
            If mbytUseType = 1 And mlng病人ID <> 0 Then
                txtPatient.Text = "-" & mlng病人ID
                Call txtPatient_KeyPress(13)
            End If
        End If
    End If
    '问题:47798
    If mbytInState = 0 Then
        Call GetRegisterItem(g私有模块, Me.Name, "idkind", strTmp)
        Err = 0: On Error Resume Next
        mblnNotCick = True
        IDKIND.IDKIND = Val(strTmp)
        mblnNotCick = False
        Err = 0: On Error GoTo 0
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, mbytInState
    mbytInState = Empty
    mstrInNO = Empty
    mblnNOMoved = False  '查阅退出后清空,避免影响后续操作
    mstrTime = ""
    mblnDelete = False
    gbytBilling = 0
    mlngDeptID = 0
    mbytUseType = 0
    mlng病人ID = 0
    Set mrs开单科室 = Nothing
    Set mrs开单人 = Nothing
    Set mrsWarn = Nothing
    Set mrsMedAudit = Nothing
    Set mrsMedPayMode = Nothing
    '问题:47798
    If mbytInState = 0 Then
        Call SaveRegisterItem(g私有模块, Me.Name, "idkind", IDKIND.IDKIND)
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Bill.Height = Me.ScaleHeight - picAppend.Height - sta.Height - fraTitle.Height - fraInfo.Height + 230
    Me.Refresh
End Sub
 
Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng卡类别ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    If objCard.名称 Like "IC卡*" And objCard.系统 Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If mobjICCard Is Nothing Then Exit Sub
        txtPatient.Text = mobjICCard.Read_Card()
        If txtPatient.Text <> "" Then
             Call FindPati(objCard, True, txtPatient.Text)
        End If
        Exit Sub
    End If
    lng卡类别ID = objCard.接口序号
    If lng卡类别ID <= 0 Then Exit Sub
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '功能:读卡接口
    '    '入参:frmMain-调用的父窗口
    '    '       lngModule-调用的模块号
    '    '       strExpand-扩展参数,暂无用
    '    '       blnOlnyCardNO-仅仅读取卡号
    '    '出参:strOutCardNO-返回的卡号
    '    '       strOutPatiInforXML-(病人信息返回.XML串)
    '    '返回:函数返回    True:调用成功,False:调用失败\
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModule, lng卡类别ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
End Sub
Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0
    Set gobjSquare.objCurCard = objCard
    If txtPatient.Text <> "" And Not mblnNotCick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub
 

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    Dim lngPreIDKind As Long, intIndex As Integer
    If txtPatient.Locked Then Exit Sub
    If objPatiInfor.卡号 = "" Then Exit Sub
    txtPatient.Text = objPatiInfor.卡号
    Call FindPati(objCard, True, txtPatient.Text)
     
End Sub

Private Sub sta_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If gbln简码切换 = False Then Exit Sub
    
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
End Sub

Private Sub txtDate_GotFocus()
    zlControl.TxtSelAll txtDate
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And IsDate(txtDate.Text) Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
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
    Dim intInsure As Integer, vDate As Date, i As Long
    Dim strInfo As String, intTmp As Integer
    
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))

    '第一位可以输入字母,其它位不行
    If KeyAscii <> 13 Then
        Call SetNOInputLimit(cboNO, KeyAscii)
    End If
    
    If KeyAscii = 13 And cboNO.Text <> "" And Not cboNO.Locked Then
        cboNO.Text = GetFullNO(cboNO.Text, 14)
        
        If chkCancel.Value = 1 Then
            '销帐
            
            If gbytBilling = 0 Then
                '是否已转入后备数据表中
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
            i = BillCanDelete(cboNO.Text, 2)
            If i <> 0 Then
                Select Case i
                    Case 1 '该单据不存在
                        MsgBox "指定单据中的内容不存在！", vbInformation, gstrSysName
                    Case 2 '已经全部完全执行
                        MsgBox "指定单据中的内容已经全部完全执行！", vbInformation, gstrSysName
                    Case 3 '未完全执行部分剩余数量为0
                        MsgBox "指定单据中的内容未完全执行部分项目剩余数量为零,没有可以销帐的费用！", vbInformation, gstrSysName
                End Select
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If

            '出院病人操作权限判断
            If Not BillCanBeOperate(cboNO.Text, mstrPrivsOpt, "销帐") Then
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If

            '是否已经结帐
            intTmp = HaveBilling(2, cboNO.Text, False)
            If intTmp <> 0 Then
                intInsure = BillExistInsure(cboNO.Text)
                If intInsure <> 0 Then
                    If Not gclsInsure.GetCapability(support允许冲销已结帐的记帐单据, , intInsure) Then
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
                            If MsgBox("该记帐单据包含已经结帐的内容,要销帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
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
            End If
            
            '是否存在重算冲减记录
            If CheckRecalcRecord(cboNO.Text) Then
                MsgBox "发现该记帐单据存在按费别重算的打折冲减记录!" & vbCrLf & _
                    "结帐前请按费别重算费用，否则病人将享受已销帐单据的打折优惠金额！", vbInformation, Me.Caption
            End If
        ElseIf mobjBill.Details.Count = 0 Then
            '记帐划价单(记帐审核)
            
            '出院病人操作权限判断
            If Not BillCanBeOperate(cboNO.Text, mstrPrivsOpt, "审核") Then
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
            
            If Not BillExistMoney(cboNO.Text, 2) Then
                MsgBox "该单据费用已经全部销帐或单据不存在！", vbInformation, gstrSysName
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
        End If
        
        '销帐或审核时,单据必须为简单收费的单据
        If Not BillisSimple(cboNO.Text) Then
            MsgBox "该单据不存在或不是简单记帐单据！", vbInformation, gstrSysName
            cboNO.Text = "": cboNO.SetFocus: Exit Sub
        End If
        
        If chkCancel.Value = 1 Then '读取退费单
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

Private Sub txtDate_LostFocus()
    txtDate.SelLength = 0
    If IsDate(txtDate.Text) Then mobjBill.发生时间 = CDate(txtDate.Text)
End Sub

Private Sub txtOld_Gotfocus()
    zlControl.TxtSelAll txtOld
End Sub

Private Sub txtOld_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        mobjBill.年龄 = txtOld.Text
        Call zlCommFun.PressKey(vbKeyTab)
    End If
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep
End Sub

Private Sub txtPatient_Change()
    txtPatient.Tag = ""
    If txtPatient.Locked Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    'If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
    IDKIND.SetAutoReadCard (txtPatient.Text = "")
    
End Sub

Private Sub txtPatient_GotFocus()
    txtPatient.SelStart = 0
    txtPatient.SelLength = Len(txtPatient.Text)
    If txtPatient.Locked Then Exit Sub
    Call IDKIND.SetAutoReadCard(txtPatient.Text = "")
    
End Sub

Private Sub bill_AfterAddRow(Row As Long)
    If mbytInState = 3 Or (chkCancel.Visible And chkCancel.Value = 1) Then
        'Bill.RemoveMSFItem Row'用属性AllowAddRow代替
        Bill.Row = 1: Call zlCommFun.PressKey(vbKeyTab)
        Exit Sub
    End If
    
    With Bill
        '新增行时,重新设置可能已经被更改的可变性质列的列值
        .ColData(1) = 5 '应收缺省跳过,当项目变价时,设为输入(4)
        '针对列编辑性质设置颜色
        .SetColColor 0, &HE7CFBA
        .SetColColor 1, &HE7CFBA
        .SetColColor 3, &HE7CFBA
    End With
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





''    Dim lngIdx As Long
''
''    If KeyAscii >= 32 And Not cbo开单科室.Locked Then
''        lngIdx = zlControl.CboMatchIndex(cbo开单科室.hwnd, KeyAscii)
''        If lngIdx = -1 And cbo开单科室.ListCount > 0 Then lngIdx = 0
''        cbo开单科室.ListIndex = lngIdx
''    ElseIf KeyAscii = 13 Then
''        If cbo开单科室.ListIndex = -1 Then
''            Beep
''        Else
''            mobjBill.开单部门ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
''            Call zlCommFun.PressKey(vbKeyTab)
''        End If
''    End If
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
    Dim i As Long, intIdx As Integer, iCount As Integer
    Dim strText As String, strResult As String, strFilter As String
    Dim rsTemp As ADODB.Recordset
    
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
                        '如果输入的数字,需要检查:
                        '1.编号输入值相等,主要输入如:12 匹配000012这种情况
                        '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                        
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
                         If Val(mrs开单人!编号) Like strText & "*" Then
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
            If ActiveControl Is txtPatient Then Call txtPatient_Validate(False)
            If ActiveControl Is cbo开单人 Then Call cbo开单人_KeyPress(vbKeyReturn)
            If cmdOK.Enabled And cmdOK.Visible Then
                Call cmdOK.SetFocus
                Call cmdOK_Click
            End If
        Case vbKeyF6 '清除当前单据内容,进入新单状态
            txtPatient.SetFocus
            Call zlControl.TxtSelAll(txtPatient)
        Case vbKeyF7 '切换输入法
            If gbln简码切换 = False Then Exit Sub   '34242
            If sta.Panels("WB").Visible And sta.Panels("PY").Visible Then
                If sta.Panels("WB").Bevel = sbrRaised Then
                    Call sta_PanelClick(sta.Panels("WB"))
                Else
                    Call sta_PanelClick(sta.Panels("PY"))
                End If
            End If
        Case vbKeyF8 '退(自动激活事件)
            If chkCancel.Visible And chkCancel.Enabled Then chkCancel.Value = IIf(chkCancel.Value = Checked, Unchecked, Checked)
        Case vbKeyEscape
            If Bill.TxtVisible Then
                Bill.Text = "": Bill.TxtVisible = False: Bill.SetFocus
            Else
                Call cmdCancel_Click
            End If
    End Select
End Sub

Private Function InitData() As Boolean
    Dim i As Long, strSQL As String
    Dim Curdate As Date     '服务器当前时间
    Err = 0: On Error GoTo errH:
    Curdate = zlDatabase.Currentdate
    
    '自动识别加班
    If mbytInState <> 2 And mstrInNO = "" Then
        If OverTime(Curdate) Then chk加班.Value = Checked
    End If
            
    If Init开单人开单科室(cbo开单人, cbo开单科室, mrs开单人, mrs开单科室, mstrPrivs, mbytUseType, mlngDeptID) = False Then
        Exit Function
    End If
    
    '执行部门
    strSQL = _
        "Select Distinct A.ID,A.编码,A.简码,A.名称,B.工作性质,B.服务对象 " & _
        " From 部门表 A,部门性质说明 B " & _
        " Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " And B.部门ID=A.ID and B.服务对象 IN(2,3) " & _
        " Order by B.服务对象,A.编码"
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
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub InitFace()
'功能：根据表单要完成的功能设置界面布局
    Dim arrHead() As String, i As Long, arrBaby As Variant
    
    '公用单据表格式
    With Bill
        .LocateCol = 0
        .PrimaryCol = 0
        .Font.Size = 11
        .CboFont.Size = 11
        .TxtEditFont.Size = 11
        
        arrHead = Split(STR_HEAD, ";")
        .Cols = UBound(arrHead) + 1
        For i = 0 To UBound(arrHead)
            .TextMatrix(0, i) = Split(arrHead(i), ",")(0)
            .ColWidth(i) = Split(arrHead(i), ",")(1)
            .ColAlignment(i) = Split(arrHead(i), ",")(2)
        Next
        
        If mbytInState = 0 And gbytBilling <> 2 Then
            .ColData(0) = 1 '项目输入,按扭可选
            .ColData(1) = 5 '应收金额缺省跳过,当项目变价时,设为输入(4)
            .ColData(2) = 5 '实收金额跳过
            .ColData(3) = 3 '默认取开单科室或上一科室
            .ColData(4) = 5
        End If
        
        .SetColColor 0, &HE7CFBA
        .SetColColor 1, &HE7CFBA
        .SetColColor 3, &HE7CFBA
        
        ReDim marrColData(.Cols - 1)
        For i = 0 To .Cols - 1
            marrColData(i) = .ColData(i)
        Next
        
        If mbytInState = 3 Then .AllowAddRow = False
    End With
    Call RestoreFlexState(Bill, App.ProductName & "\" & Me.Name & mbytInState)
    Call SetMoneyList
    
    '读取简码匹配方式
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
            lblTitle.Caption = gstrUnitName & "住院记帐单"
        Case 1
            lblTitle.Caption = gstrUnitName & "住院记帐单(划价)"
        Case 2
            lblTitle.Caption = gstrUnitName & "住院记帐单(审核)"
    End Select
    
    txt应收.Text = gstrDec: txt实收.Text = gstrDec
    
    Select Case mbytInState
        Case 0 '执行
            If mstrInNO <> "" Or _
                (InStr(mstrPrivsOpt, ";药品销帐;") = 0 _
                    And InStr(mstrPrivsOpt, ";卫材销帐;") = 0 _
                    And InStr(mstrPrivsOpt, ";诊疗销帐;") = 0) Then
                chkCancel.Visible = False
                lblNO.Left = lblNO.Left + chkCancel.Width
                cboNO.Left = cboNO.Left + chkCancel.Width
            End If
            Select Case gbytBilling
                Case 0, 1 '执行记帐、划价
                    txtPatient.Enabled = (mstrInNO = "")
                Case 2 '执行审核
                    Call SetDisible
                    cboNO.Locked = False
                    fraInfo.Enabled = False
                    fraAppend.Enabled = False
                    Bill.Active = False
            End Select
        Case 1 '查阅
            Call SetDisible
            
            chkCancel.Visible = False
            If mblnDelete Then
                lblFlag.Visible = True
            Else
                lblNO.Left = lblNO.Left + chkCancel.Width
                cboNO.Left = cboNO.Left + chkCancel.Width
            End If
            fraTitle.Enabled = False
            fraInfo.Enabled = False
            fraAppend.Enabled = False
            Bill.Active = False
            cmdOK.Visible = False
            cmdCancel.Caption = "退出(&X)"
        Case 2 '调整
            Call SetDisible
            txtDate.Enabled = True
            chkCancel.Visible = False
            lblNO.Left = lblNO.Left + chkCancel.Width
            cboNO.Left = cboNO.Left + chkCancel.Width
            fraTitle.Enabled = False
            fraInfo.Enabled = False
            Bill.Active = False
        Case 3 '销帐
            Call SetDisible
            
            chkCancel.Visible = False
            lblNO.Left = lblNO.Left + chkCancel.Width
            cboNO.Left = cboNO.Left + chkCancel.Width
            fraTitle.Enabled = False
            fraInfo.Enabled = False
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
    txtPatient.Locked = Not bln
    cbo开单科室.Locked = Not bln
    chk加班.Enabled = bln
    
    cbo开单人.Locked = Not bln
    txtDate.Enabled = bln
    Bill.Active = bln
    
    If Not bln Then
        txtPatient.BackColor = &HE0E0E0
        txt医疗付款.BackColor = &HE0E0E0
    Else
        txtPatient.BackColor = &HFFFFFF
        txt医疗付款.BackColor = &HFFFFFF
    End If
End Sub

Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtPatient.Locked Or txtPatient.Enabled = False Or txtPatient.Text <> "" Then Exit Sub
    If IDKIND.ActiveFastKey = True Then Exit Sub
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim lngID As Long, lngUnit As Long, i As Long
    Dim rsTmp As ADODB.Recordset, strInfo As String
    Dim strSQL As String, curTotal As Currency
    Dim blnCard As Boolean
    
    On Error GoTo errH
    
    If txtPatient.Locked Then Exit Sub
    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    
    If Len(Trim(Me.txtPatient.Text)) = 0 And KeyAscii = 13 Then
        With frmPatiSelect
            If (mbytUseType = 0 Or mbytUseType = 1) Then
                .mlngUnitID = mlngUnitID
            Else
                .mlngUnitID = mlngDeptID
            End If
            .mbytUseType = mbytUseType
            .mstrPrivs = mstrPrivs
            Set .mfrmParent = Me
            .Show 1, Me
        End With
    Else
        If IDKIND.GetCurCard.名称 Like "姓名*" Then
            blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKIND.ShowPassText)
        ElseIf IDKIND.GetCurCard.名称 = "门诊号" Or IDKIND.GetCurCard.名称 = "住院号" Then
            If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
                If InStr("0123456789-*+", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
            End If
        Else
            txtPatient.PasswordChar = IIf(IDKIND.ShowPassText, "*", "")
            '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
            txtPatient.IMEMode = 0
        End If
    End If
    If blnCard And Len(txtPatient.Text) = IDKIND.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtPatient.Text) <> "" Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        ElseIf IsNumeric(txtPatient.Tag) And mrsInfo.State = 1 Then
            KeyAscii = 0
            
            '刷新病人信息:"-病人ID"
            Call GetPatient(IDKIND.GetCurCard, txtPatient.Tag, False)
            If mrsInfo.State = 0 Then   '连续记帐时，可能此时病人因产生了费用，而操作员没有"出院未结强制记帐"权限，读不出病人
                txtPatient.Text = "": txtOld.Text = ""
                txt床号.Text = ""
                Exit Sub
            End If
            
            '刷新病人预交款信息
            curTotal = GetBillTotal(mobjBill)
            Set rsTmp = GetMoneyInfo(mrsInfo!病人ID, CDbl(mcurModiMoney), True, 2)
            If Not rsTmp Is Nothing Then
                cmdOK.Tag = rsTmp!预交余额
                cmdCancel.Tag = rsTmp!费用余额
                txt实收.Tag = rsTmp!预交余额 - rsTmp!费用余额
            Else
                cmdOK.Tag = 0: cmdCancel.Tag = 0: txt实收.Tag = 0
            End If
            '划价时显示不算当前单据费用,但划价报警要算
            sta.Panels(3).Text = "预交:" & Format(Val(cmdOK.Tag), "0.00")
            sta.Panels(3).Text = sta.Panels(3).Text & "/费用:" & Format(Val(cmdCancel.Tag) + IIf(gbytBilling = 0, curTotal, 0), gstrDec)
            sta.Panels(3).Text = sta.Panels(3).Text & "/剩余:" & Format(Val(txt实收.Tag) - IIf(gbytBilling = 0, curTotal, 0), "0.00")
            strInfo = GetPatientDue(Val(mrsInfo!病人ID))
            If Val(strInfo) <> 0 Then sta.Panels(3).Text = sta.Panels(3).Text & "/应收款:" & Format(strInfo, "0.00")
            Call LoadPatientBaby(cboBaby, mrsInfo!病人ID, mrsInfo!主页ID)
            If Not mblnValid Then Bill.SetFocus
            Exit Sub
        End If
        KeyAscii = 0
        '69282,刘尔旋,2014-01-03,通过姓名+住院号方式找病人出错的问题
        Call FindPati(IDKIND.GetCurCard, blnCard, txtPatient.Text)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FindPati(ByVal objCard As Card, ByVal blnCard As Boolean, ByVal strInput As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:查找病人
    '编制:刘兴洪
    '日期:2012-08-29 17:53:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnMsg As Boolean
    Dim lngID As Long, lngUnit As Long, i As Long
    Dim rsTmp As ADODB.Recordset, strInfo As String
    Dim strSQL As String, curTotal As Currency
    
    '20030617:保存后未清除金额
    If mobjBill.Details.Count = 0 Then
        Call ClearMoney
        txt实收.Text = gstrDec: txt应收.Text = gstrDec
    End If
    
    '读取病人信息
    If Not (mbytInState = 0 And mbytUseType = 1 And sta.Panels(2) Like "上一张*") Then
        sta.Panels(2) = ""
    End If
    If Not GetPatient(objCard, strInput, blnCard, blnMsg) Then
        If blnCard Then
            If Not blnMsg Then MsgBox "不能确定病人信息，请检查是否正确刷卡！", vbInformation, gstrSysName
            txtPatient.Text = "": txtOld.Text = "": txt床号.Text = "": Exit Sub
        Else
            If Not blnMsg Then MsgBox "不能读取病人信息！", vbInformation, gstrSysName
            zlControl.TxtSelAll txtPatient
            If mstrInNO = "" Then txtOld.Text = "": txt床号.Text = ""
            Exit Sub
        End If
        Exit Sub
    End If
    
    '就诊卡密码检查
    If (objCard.名称 Like "*IC卡*" Or objCard.名称 Like "*身份证*") And objCard.系统 And mstrPassWord = "" Then blnCard = False
    If Mid(gstrCardPass, 6, 1) = "1" And blnCard Then
        If Not zlCommFun.VerifyPassWord(Me, mstrPassWord, mrsInfo!姓名, mrsInfo!性别, "" & mrsInfo!年龄) Then
            Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": txtPatient.SetFocus: Exit Sub
        End If
    End If
    
    If mbytUseType = 1 And mrsInfo!病人ID <> mlng病人ID Then mlng病人ID = 0
     
      '自动设置开单科室(同时设置记帐报警信息),医技记帐病人科室不一定是开单科室
     If mbytUseType = 2 Then lngUnit = cbo开单科室.ListIndex
    
    If gblnFromDr Then
        If Not IsNull(mrsInfo!住院医师) Then
            cbo开单人.ListIndex = -1
            cbo开单人.ListIndex = cbo.FindIndex(cbo开单人, mrsInfo!住院医师, True)
        End If
    Else
        cbo开单科室.ListIndex = -1
        cbo开单科室.ListIndex = cbo.FindIndex(cbo开单科室, IIf(IsNull(mrsInfo!科室ID), 0, mrsInfo!科室ID))
        If cbo开单科室.ListIndex <> -1 Then
            mobjBill.开单部门ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
        ElseIf mbytUseType = 2 And lngUnit <> -1 Then
            cbo开单科室.ListIndex = lngUnit
        End If
    End If
    
    '病人预交款信息
    curTotal = GetBillTotal(mobjBill)
    Set rsTmp = GetMoneyInfo(mrsInfo!病人ID, CDbl(mcurModiMoney), True, 2)
    If Not rsTmp Is Nothing Then
        cmdOK.Tag = rsTmp!预交余额
        cmdCancel.Tag = rsTmp!费用余额
        txt实收.Tag = rsTmp!预交余额 - rsTmp!费用余额
    Else
        cmdOK.Tag = 0: cmdCancel.Tag = 0: txt实收.Tag = 0
    End If
    '--------------------------------------------------------------------------------------------------------------------------------------------------------------
    '刘兴洪:26952
    Dim cur余额 As Currency, curItemMoney As Currency
    
    cur余额 = Val(txt实收.Tag)
    
    '刘兴洪:24491
    curItemMoney = 0 ' GetBillRowTotal(mobjBill.Details(Bill.Row).InComes)
     
    If gbln报警包含划价费用 Then cur余额 = Val(txt实收.Tag) - GetPriceMoneyTotal(1, mrsInfo!病人ID) + IIf(gbytBilling = 1, mcurModiMoney, 0)
    
    gbytWarn = BillingWarn(mstrPrivsOpt, mrsInfo!姓名 & IIf(Nvl(mrsInfo!住院号) = "", "", "(住院号:" & mrsInfo!住院号 & " 床号:" & mrsInfo!床号 & ")"), Val("" & mrsInfo!病区ID), mrsInfo!适用病人, mrsWarn, cur余额, mrsInfo!当日额 - mcurModiMoney, curTotal, _
                IIf(IsNull(mrsInfo!担保额), 0, mrsInfo!担保额), "", "", _
                 mstrWarn, , gblnPrice And (gbytBilling = 0 And mstrInNO = "" Or gbytBilling = 1), curItemMoney, True)
    '返回:0;没有报警,继续
    '     1:报警提示后用户选择继续
    '     2:报警提示后用户选择中断
    '     3:报警提示必须中断
    '     4:强制记帐报警,继续
    '     5.报警提示后用户选择继续,但只允许保存存为划价单
    If gbytWarn = 2 Or gbytWarn = 3 Then
        Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "":
         mlng病人ID = 0
        If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
        Exit Sub
    End If
    '--------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '划价时显示不算当前单据费用,但划价报警要算
    sta.Panels(3).Text = "预交:" & Format(Val(cmdOK.Tag), "0.00")
    sta.Panels(3).Text = sta.Panels(3).Text & "/费用:" & Format(Val(cmdCancel.Tag) + IIf(gbytBilling = 0, curTotal, 0), gstrDec)
    sta.Panels(3).Text = sta.Panels(3).Text & "/剩余:" & Format(Val(txt实收.Tag) - IIf(gbytBilling = 0, curTotal, 0), "0.00")
    strInfo = GetPatientDue(Val(mrsInfo!病人ID))
    If Val(strInfo) <> 0 Then sta.Panels(3).Text = sta.Panels(3).Text & "/应收款:" & Format(strInfo, "0.00")
    
    Call LoadPatientBaby(cboBaby, mrsInfo!病人ID, mrsInfo!主页ID)
                
    '病人信息
    txtPatient.Text = IIf(IsNull(mrsInfo!姓名), "", mrsInfo!姓名)
    txtSex.Text = IIf(IsNull(mrsInfo!性别), "", mrsInfo!性别)
    txtOld.Text = IIf(IsNull(mrsInfo!年龄), "", mrsInfo!年龄)
    txt费别.Text = IIf(IsNull(mrsInfo!费别), "", mrsInfo!费别)
    txt医疗付款.Text = IIf(IsNull(mrsInfo!医疗付款方式), "", mrsInfo!医疗付款方式)
    txt医疗付款.Tag = GetMedPayMode(txt医疗付款.Text, mrsMedPayMode)
    If gintPriceGradeStartType >= 2 Then
        Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, Val(Nvl(mrsInfo!病人ID)), Val(Nvl(mrsInfo!主页ID)), txt医疗付款.Text, mstr药品价格等级, mstr卫材价格等级, mstr普通价格等级)
        If mobjBill.Details.Count > 0 Then
            '重新计算并刷新
            Call CalcMoneys
            Call ShowDetails
            Call ShowMoney
        End If
    End If
    txt床号.Text = "" & mrsInfo!床号
    txt担保人.Text = IIf(IsNull(mrsInfo!担保人), "", mrsInfo!担保人)
    txt担保额.Text = Format(IIf(IsNull(mrsInfo!担保额), "", mrsInfo!担保额), "0.00")
    txtPatient.PasswordChar = ""
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0
     
     With mobjBill
         .病人ID = IIf(IsNull(mrsInfo!病人ID), 0, mrsInfo!病人ID)
         .主页ID = IIf(IsNull(mrsInfo!主页ID), 0, mrsInfo!主页ID)
         
         .病区ID = IIf(IsNull(mrsInfo!病区ID), 0, mrsInfo!病区ID)
         .科室ID = IIf(IsNull(mrsInfo!科室ID), 0, mrsInfo!科室ID)
         
         .床号 = "" & mrsInfo!床号
         .标识号 = IIf(IsNull(mrsInfo!住院号), 0, mrsInfo!住院号)
         .姓名 = txtPatient.Text
         .性别 = IIf(IsNull(mrsInfo!性别), "", mrsInfo!性别)
         .年龄 = txtOld.Text
         .费别 = IIf(IsNull(mrsInfo!费别), "", mrsInfo!费别)
     End With
     If Not IsNull(mrsInfo!出院日期) Then
         MsgBox "提醒您：" & vbCrLf & vbCrLf & "该病人已于 " & Format(mrsInfo!出院日期, "yyyy-MM-dd") & " 出院，现在对该病人强制进行记帐！", vbInformation, gstrSysName
         txtDate.Text = Format(mrsInfo!出院日期, "yyyy-MM-dd HH:mm:ss")
     Else
         txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
     End If
     If Not (mbytInState = 0 And mbytUseType = 1 And sta.Panels(2) Like "上一张*") Then
         If Not IsNull(mrsInfo!入院日期) Then
             sta.Panels(2).Text = "入院:" & Format(mrsInfo!入院日期, "yyyy-MM-dd")
             strInfo = GetInsureInfo(mrsInfo!病人ID)
             If strInfo <> "" Then sta.Panels(2).Text = sta.Panels(2).Text & "/帐号:" & Split(strInfo, ";")(1)
         End If
     End If
     If Visible Then
        If gblnFromDr Then
            cbo开单人.SetFocus
        Else
            cbo开单科室.SetFocus
        End If
     End If
 End Sub


Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, ByVal blnCard As Boolean, Optional blnOutMsg As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息
    '入参:blnCard=是否就诊卡刷卡
    '出参:
    '返回:查找到病人,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-03 17:54:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, lng卡类别ID As Long, strPassWord As String, strErrMsg As String
    Dim lng病人ID As Long, blnHavePassWord As Boolean
    Dim strIF As String, strWhere As String
    Dim rsOutSel As ADODB.Recordset, bln所有病区 As Boolean
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
            strIF = strIF & " And B.当前病区ID+0 IN(Select Column_Value From Table(Cast(f_num2list([4]) As zlTools.t_numlist)))"
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
    " Select A.病人ID,B.主页ID,B.当前病区ID as 病区ID,B.出院科室ID as 科室ID,B.入院日期,B.出院日期," & _
    "       A.就诊卡号,A.卡验证码,A.住院号,B.出院病床 as 床号,X.费用余额,B.状态," & _
    "       Nvl(b.姓名, a.姓名) As 姓名, Nvl(b.性别, a.性别) As 性别,A.年龄,B.费别,B.住院医师,B.医疗付款方式," & _
    "       A.担保人,Decode(A.担保额,null,A.担保额,Zl_Patientsurety(A.病人ID,B.主页ID)) 担保额,zl_PatiDayCharge(A.病人ID) as 当日额," & _
    "       Zl_Patiwarnscheme(B.病人id, B.主页id) As 适用病人,B.险类,Nvl(B.病人性质,0) as 病人性质,B.审核标志,B.病人类型" & _
    " From 病人信息 A,病案主页 B,病人余额 X " & _
    " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID" & _
    "        And Nvl(B.主页ID,0)<>0 And A.病人ID=X.病人ID(+) And X.性质(+)=1 And X.类型(+)=2  And A.停用时间 is NULL " & strIF
    If blnCard = True And objCard.名称 Like "姓名*" Then  '刷卡
    
        If IDKIND.Cards.按缺省卡查找 And Not IDKIND.GetfaultCard Is Nothing Then
            lng卡类别ID = IDKIND.GetfaultCard.接口序号
        Else
            lng卡类别ID = "-1"
        End If
        If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg, lng卡类别ID) = False Then GoTo NotFoundPati:
        If lng病人ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng病人ID
        blnHavePassWord = True
        strWhere = strWhere & " And A.病人ID=[1] "
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '病人ID
        strWhere = strWhere & " And A.病人ID=[1]"
    ElseIf Left(strInput, 1) = "/" Then   '床位号
        '41654 And IsNumeric(Mid(strInput, 2))
        strInput = Mid(strInput, 2)
        If mlngUnitID = 0 Then '病区不确定、则不能通过床号确定病人
            Set mrsInfo = New ADODB.Recordset: Exit Function
        End If
        strSQL = _
            " Select A.病人ID,B.主页ID,B.当前病区ID as 病区ID,B.出院科室ID as 科室ID,B.入院日期,B.出院日期," & _
            "       A.就诊卡号,A.卡验证码,A.住院号,B.出院病床 as 床号,X.费用余额,B.状态," & _
            "       Nvl(b.姓名, a.姓名) As 姓名, Nvl(b.性别, a.性别) As 性别,A.年龄,B.费别,B.住院医师,B.医疗付款方式," & _
            "       A.担保人,Decode(A.担保额,null,A.担保额,Zl_Patientsurety(A.病人ID,B.主页ID)) 担保额,zl_PatiDayCharge(A.病人ID) as 当日额," & _
            "       Zl_Patiwarnscheme(B.病人id, B.主页id) As 适用病人,B.险类,Nvl(B.病人性质,0) as 病人性质,B.审核标志,B.病人类型" & _
            " From 病人信息 A,病案主页 B,床位状况记录 C,病人余额 X" & _
            " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID" & _
            "       And Nvl(B.主页ID,0)<>0 And A.病人ID=C.病人ID And A.病人ID=X.病人ID(+) And X.性质(+)=1 And X.类型(+)=2 And A.停用时间 is NULL" & _
            "       And C.病区ID=[3] And C.床号=[2] " & strIF
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then '住院号(病人在院)
        strWhere = strWhere & " And A.病人ID = (Select Max(病人id) From 病案主页 Where 住院号 = [1])"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '门诊号(医技记帐)
        strWhere = strWhere & " And A.门诊号=[1]"
    Else '当作姓名
        Select Case objCard.名称
            Case "姓名", "姓名或就诊卡"
                If mrsInfo.State = 1 Then
                    If Not mrsInfo.EOF Then
                        If mrsInfo!姓名 = Trim(txtPatient.Text) Then GetPatient = True: Exit Function
                    End If
                End If
                If zlSelectChargePatiFromInputName(Me, mstrPrivsOpt, strInput, bln所有病区, mstrUnitIDs, gintOutDay, lng病人ID, strErrMsg, txtPatient.hWnd, txtPatient.Height) = False Then
                    If strErrMsg = "" Then blnOutMsg = True: Set mrsInfo = New Recordset: Exit Function
                    If mbytUseType = 2 And InStr(mstrPrivsOpt, ";出院未结强制记帐;") > 0 And InStr(mstrPrivsOpt, ";出院结清强制记帐;") > 0 Then GoTo GoYJReadPati:
                    MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
                    blnOutMsg = True: Set mrsInfo = New Recordset: Exit Function
                End If
                strInput = "-" & lng病人ID
                strWhere = strWhere & " And A.病人ID=[1]"
            Case "医保号"
                strInput = UCase(strInput)
                strWhere = strWhere & " And A.医保号=[2]"
            Case "门诊号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strWhere = strWhere & " And A.门诊号=[2]"
            Case "住院号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strWhere = strWhere & " And A.病人ID = (Select Max(病人id) From 病案主页 Where 住院号 = [2])"
            Case Else
                '其他类别的,获取相关的病人ID
                If objCard.接口序号 > 0 Then
                    lng卡类别ID = objCard.接口序号
                    If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    If lng病人ID = 0 Then GoTo NotFoundPati:
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.名称, strInput, False, lng病人ID, _
                        strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                End If
                If lng病人ID <= 0 Then GoTo NotFoundPati:
                strWhere = strWhere & " And A.病人ID=[1]"
                strInput = "-" & lng病人ID
                blnHavePassWord = True
         End Select
    End If
    
    strSQL = strSQL & vbCrLf & strWhere
    txtPatient.ForeColor = Me.ForeColor
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput, mlngUnitID, mstrUnitIDs)
    
    If Not mrsInfo.EOF Then
        txtPatient.ForeColor = zlDatabase.GetPatiColor(Nvl(mrsInfo!病人类型))
        If zlPatiIS病案已编目(Val(Nvl(mrsInfo!病人ID)), Val(Nvl(mrsInfo!主页ID))) = True Then    '问题:28725
            Set mrsInfo = New ADODB.Recordset
            Set mrsMedAudit = Nothing
            blnOutMsg = True
            Exit Function
        End If
        If zlIsAllowFeeChange(Val(Nvl(mrsInfo!病人ID)), Val(Nvl(mrsInfo!主页ID)), Val(Nvl(mrsInfo!审核标志))) = False Then
            Set mrsInfo = New ADODB.Recordset
            Set mrsMedAudit = Nothing
            blnOutMsg = True
            Exit Function
        End If
        
        If mrsInfo!病人ID <> mobjBill.病人ID Or mbytInState = 0 And mstrInNO <> "" Then    '同一病人不用重复读取
            If GetMedPayMode("" & mrsInfo!医疗付款方式, mrsMedPayMode) = 1 Then
                Set mrsMedAudit = GetAuditRecord(mrsInfo!病人ID, mrsInfo!主页ID)
            Else
                Set mrsMedAudit = Nothing
            End If
        End If
         mstrPassWord = strPassWord
        If Not blnHavePassWord Then
            mstrPassWord = Nvl(mrsInfo!卡验证码)
        End If
        GetPatient = True
        Exit Function
    Else
        Set mrsMedAudit = Nothing   '医保病人必须在院才检查费用审批
    End If
    
        
    '医技科室记帐：没有发现住院(在院或出院)病人,以门诊病人读
    If mbytUseType = 2 And InStr(mstrPrivsOpt, ";出院未结强制记帐;") > 0 And InStr(mstrPrivsOpt, ";出院结清强制记帐;") > 0 Then
GoYJReadPati:
        '76451,冉俊明,2014-8-19
        strSQL = _
        " Select A.病人ID,Nvl(A.主页ID,0) as 主页ID,A.当前病区ID as 病区ID,A.当前科室ID as 科室ID," & _
        "       A.出院时间 as 出院日期,A.就诊卡号,A.卡验证码,A.住院号,A.当前床号 as 床号,A.姓名,A.性别,A.年龄," & _
        "       A.入院时间 as 入院日期,A.费别,A.担保人,Decode(A.担保额,null,A.担保额,Zl_Patientsurety(A.病人ID,null)) 担保额,Zl_Patiwarnscheme(A.病人id) As 适用病人,NULL as 住院医师,A.医疗付款方式," & _
        "       zl_PatiDayCharge(A.病人ID) as 当日额,A.险类,-1 as 病人性质" & _
        " From 病人信息 A Where A.停用时间 is NULL "
        If blnCard = True And objCard.名称 Like "姓名*" Then   '刷卡
            If IDKIND.Cards.按缺省卡查找 And Not IDKIND.GetfaultCard Is Nothing Then
                lng卡类别ID = IDKIND.GetfaultCard.接口序号
            Else
                lng卡类别ID = "-1"
            End If
            '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户);…
            If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
            If lng病人ID <= 0 Then GoTo NotFoundPati:
            strInput = "-" & lng病人ID
            blnHavePassWord = True
            strSQL = strSQL & " And A.病人ID=[1] "
        ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '病人ID
            strSQL = strSQL & " And A.病人ID=[1]"
        ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '门诊号(医技记帐)
            strSQL = strSQL & " And A.门诊号=[1]"
        Else '当作姓名
            Select Case objCard.名称
                  Case "姓名", "姓名或就诊卡"
                      If mrsInfo.State = 1 Then
                          If mrsInfo!姓名 = Trim(txtPatient.Text) Then GetPatient = True: Exit Function
                      End If
                      strSQL = strSQL & " And A.姓名=[2]"
                  Case "医保号"
                      strInput = UCase(strInput)
                      strSQL = strSQL & " And A.医保号=[2]"
                  Case "门诊号"
                      If Not IsNumeric(strInput) Then strInput = "0"
                      strSQL = strSQL & " And A.门诊号=[2]"
                  Case "住院号"
                      If Not IsNumeric(strInput) Then strInput = "0"
                      strSQL = strSQL & " And A.住院号=[2]"
                  Case Else
                      '其他类别的,获取相关的病人ID
                      If objCard.接口序号 > 0 Then
                          lng卡类别ID = objCard.接口序号
                          If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                          If lng病人ID = 0 Then GoTo NotFoundPati:
                      Else
                          If gobjSquare.objSquareCard.zlGetPatiID(objCard.名称, strInput, False, lng病人ID, _
                              strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                      End If
                      If lng病人ID <= 0 Then GoTo NotFoundPati:
                      strSQL = strSQL & " And A.病人ID=[1]"
                      strInput = "-" & lng病人ID
                      blnHavePassWord = True
               End Select
        End If
        Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
        If Not mrsInfo.EOF Then
            If zlPatiIS病案已编目(Val(Nvl(mrsInfo!病人ID)), Val(Nvl(mrsInfo!主页ID))) = True Then    '问题:28725
                Set mrsInfo = New ADODB.Recordset
                blnOutMsg = True
                Exit Function
            End If
            mstrPassWord = strPassWord
            If Not blnHavePassWord Then
               mstrPassWord = Nvl(mrsInfo!卡验证码)
            End If
            GetPatient = True
            Exit Function
        End If
        Set mrsInfo = New ADODB.Recordset
        Exit Function
    End If
    Set mrsMedAudit = Nothing   '医保病人必须在院才检查费用审批'
    Set mrsInfo = New ADODB.Recordset
    If strWhere = "" Then Exit Function '无其他条件，直接退出
    
    '未找到病人，需要对该病人的具体错误信息进行提示
    strSQL = _
    " Select A.病人ID,B.主页ID,B.当前病区ID as 病区ID,B.出院科室ID as 科室ID,a.在院,B.入院日期,B.出院日期,X.费用余额,B.状态, " & _
    "       nvl(B.姓名,A.姓名) as 姓名,nvl(b.性别,A.性别) as 性别,nvl(b.年龄,A.年龄) as 年龄,B.费别,Nvl(B.病人性质,0) as 病人性质,B.病人类型" & _
    " From 病人信息 A,病案主页 B,病人余额 X" & _
    " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID" & _
    "   And Nvl(B.主页ID,0)<>0 And A.病人ID=X.病人ID(+) and X.性质(+)=1 and X.类型(+)=2 And A.停用时间 is NULL " & strWhere
    
    Set rsOutSel = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
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
        Set mrsMedAudit = Nothing   '医保病人必须在院才检查费用审批'
        blnOutMsg = True
        Exit Function
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
NotFoundPati:
    Set mrsInfo = New ADODB.Recordset
End Function

Private Sub CalcMoneys(Optional lngRow As Long = 0)
'功能：计算或重新计算指定行或所有行的金额
'参数：lngRow=指定行,为0表示计算所有行
'说明：ExpenseBill集合的索引对应单据的行号
    Dim i As Long
    If mobjBill.Details.Count = 0 Then Exit Sub
    If lngRow = 0 Then
        For i = 1 To mobjBill.Details.Count
            CalcMoney i
        Next
    Else
        CalcMoney lngRow
    End If
End Sub

Private Sub CalcMoney(lngRow As Long)
'功能：计算或重新计算指定行的金额
'参数：lngRow=指定行
'说明：1.ExpenseBill集合的索引对应单据的行号
'      2.变价只能对应一个收入项目:mobjBill.Details(lngRow).InComes(1)
'      3.如果变价细目未计算出收入项目(第一次计算),则使用默认现价
'      4.如果变价细目已经计算出收入项目(按第2步),并手动更改(也可能未改)了单价,则按该单价计算。
    Dim i As Long, strInfo As String
    Dim rsTmp As ADODB.Recordset
    Dim dblMoney As Double '用户输入的变价金额
    Dim dbl加班加价率 As Double
    Dim strWherePriceGrade As String
    
    On Error GoTo errH
    If mstr普通价格等级 <> "" Then
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
    
    gstrSQL = _
        " Select B.收入项目ID,C.名称,C.收据费目,B.现价,B.原价,B.加班加价率,B.附术收费率,b.缺省价格 " & _
        " From 收费项目目录 A,收费价目 B,收入项目 C " & _
        " Where B.收费细目ID = A.ID And C.ID = B.收入项目ID " & _
        " And Sysdate Between B.执行日期 And Nvl(B.终止日期,To_Date('3000-1-1', 'YYYY-MM-DD'))" & _
        " And A.ID=[1]" & vbNewLine & _
        strWherePriceGrade
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mobjBill.Details(lngRow).收费细目ID, mstr普通价格等级)
    
    If rsTmp.EOF Then
        '如果没有收入项目,则清除对应的程序对象
        Set mobjBill.Details(lngRow).InComes = New BillInComes
        Exit Sub
    End If
    
    '先获取操作员以前输入的变价金额
    With mobjBill.Details(lngRow)
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
            .原价 = Nvl(rsTmp!原价, 0)
            .现价 = Nvl(rsTmp!现价, 0)
            If mobjBill.Details(lngRow).Detail.变价 Then
                .标准单价 = Format(dblMoney, gstrFeePrecisionFmt)
            Else
                .标准单价 = Format(Nvl(rsTmp!现价), gstrFeePrecisionFmt)
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
            
            If mobjBill.Details(lngRow).Detail.屏蔽费别 Then
                .实收金额 = .应收金额
            Else
                If .应收金额 = 0 Then
                    .实收金额 = 0
                    mobjBill.Details(lngRow).费别 = mobjBill.费别
                Else
                    .实收金额 = CCur(Format(ActualMoney(mobjBill.费别, .收入项目ID, .应收金额, 0, 0, 0, dbl加班加价率), gstrDec))
                End If
            End If
            
            '获取项目保险信息,医保病人才处理,不需要连接医保
            If mrsInfo.State = 1 Then
                If Not IsNull(mrsInfo!险类) Then
                    strInfo = gclsInsure.GetItemInsure(mobjBill.病人ID, mobjBill.Details(lngRow).收费细目ID, .实收金额, False, mrsInfo!险类, _
                        mobjBill.Details(lngRow).摘要 & "||" & mobjBill.Details(lngRow).数次)
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
            End If
            
            mobjBill.Details(lngRow).InComes.Add .收入项目ID, .收入项目, .收据费目, .标准单价, .应收金额, .实收金额, .原价, .现价, , .统筹金额
        End With
        rsTmp.MoveNext
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ShowDetails(Optional lngRow As Long = 0)
'功能：刷新显示指定行或所有行的内容
'参数：lngRow=指定行,为0表示显示所有行
'说明：ExpenseBill集合的索引对应单据的行号
    Dim i As Long
    Dim curTotal As Currency
    
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
        '划价时显示不算当前单据费用,但划价报警要算
        sta.Panels(3).Text = "预交:" & Format(Val(cmdOK.Tag), "0.00")
        sta.Panels(3).Text = sta.Panels(3).Text & "/费用:" & Format(Val(cmdCancel.Tag) + IIf(gbytBilling = 0, curTotal, 0), gstrDec)
        sta.Panels(3).Text = sta.Panels(3).Text & "/剩余:" & Format(Val(txt实收.Tag) - IIf(gbytBilling = 0, curTotal, 0), "0.00")
    End If
End Sub

Private Sub ShowDetail(lngRow As Long)
'功能：刷新显示指定行的内容
'参数：lngRow=指定行
'说明：ExpenseBill集合的索引对应单据的行号
    Dim curMoney As Currency
    Dim i As Long, j As Long
    '清除单据行
    For i = 0 To Bill.Cols - 1
        '输入时收费类别不清除
        If Not (i = 0 And Bill.TextMatrix(lngRow, i) <> "") Then Bill.TextMatrix(lngRow, i) = ""
    Next
    '刷新单据行
    For i = 0 To Bill.Cols - 1
        Select Case Bill.TextMatrix(0, i)
            Case "项目"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.名称
            Case "应收金额" '实际上是单价
                '单价是该收费细目所有收入项目的合计
                '第一次计算时是在默认数次为1的基础上计算出来的
                curMoney = 0
                If mobjBill.Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Details(lngRow).InComes.Count
                        curMoney = curMoney + mobjBill.Details(lngRow).InComes(j).应收金额
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(curMoney, gstrDec)
            Case "实收金额"
                '实收金额是该收费细目所有收入项目的合计
                curMoney = 0
                If mobjBill.Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Details(lngRow).InComes.Count
                        curMoney = curMoney + mobjBill.Details(lngRow).InComes(j).实收金额
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(curMoney, gstrDec)
            Case "执行科室"
                If mbytInState = 0 Then
                    mrsUnit.Filter = "ID=" & mobjBill.Details(lngRow).执行部门ID
                    If mrsUnit.RecordCount <> 0 Then
                        Bill.TextMatrix(lngRow, i) = IIf(zlIsShowDeptCode, mrsUnit!编码 & "-", "") & mrsUnit!名称
                    Else
                        Bill.TextMatrix(lngRow, i) = GET部门名称(mobjBill.Details(lngRow).执行部门ID, mrsUnit)
                    End If
                Else
                    '浏览单据只(能)显示名称
                    Bill.TextMatrix(lngRow, i) = GET部门名称(mobjBill.Details(lngRow).执行部门ID, mrsUnit)
                End If
            Case "类型"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.类型
        End Select
    Next
    Bill.Text = Bill.MsfObj.Text
End Sub

Private Function GetInputDetail(ByVal lng项目id As Long) As Detail
'功能：读取收费项目信息
    Dim objDetail As New Detail
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long, lngMediCareNO As Long
        
    If mrsInfo.State = 1 Then lngMediCareNO = Val("" & mrsInfo!险类)
    If lngMediCareNO > 0 Then
        strSQL = _
            " Select A.ID,A.类别,B.名称 as 类别名称,A.编码,A.名称,A.规格,A.计算单位," & _
            " A.屏蔽费别,A.是否变价,A.加班加价,A.执行科室,A.费用类型,A.补充摘要,A.服务对象,F.要求审批" & _
            " From 收费项目目录 A,收费项目类别 B,保险支付项目 F" & _
            " Where A.类别=B.编码 And A.ID=[1] And A.ID=F.收费细目ID(+) And F.险类(+)=[2]"
    Else
        strSQL = _
            " Select A.ID,A.类别,B.名称 as 类别名称,A.编码,A.名称,A.规格,A.计算单位," & _
            " A.屏蔽费别,A.是否变价,A.加班加价,A.执行科室,A.费用类型,A.补充摘要,A.服务对象,0 as 要求审批" & _
            " From 收费项目目录 A,收费项目类别 B" & _
            " Where A.类别=B.编码 And A.ID=[1]"
    End If
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng项目id, lngMediCareNO)
    With objDetail
        .ID = rsTmp!ID
        .类别 = rsTmp!类别
        .类别名称 = rsTmp!类别名称
        .编码 = rsTmp!编码
        .名称 = rsTmp!名称
        .规格 = Nvl(rsTmp!规格)
        .计算单位 = Nvl(rsTmp!计算单位)
        .变价 = Nvl(rsTmp!是否变价, 0) = 1 '对药品表明是否时价
        .类型 = Nvl(rsTmp!费用类型)
        .加班加价 = Nvl(rsTmp!加班加价, 0) = 1
        .屏蔽费别 = Nvl(rsTmp!屏蔽费别, 0) = 1
        .执行科室 = Nvl(rsTmp!执行科室, 0)
        .服务对象 = Nvl(rsTmp!服务对象, 0)
        .补充摘要 = Nvl(rsTmp!补充摘要, 0) = 1
        .要求审批 = Nvl(rsTmp!要求审批, 0) = 1
    End With
    Set GetInputDetail = objDetail
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetDetail(Detail As Detail, lngRow As Long, Optional bytParent As Byte = 0, Optional ByVal lngDoUnit As Long)
'功能：根据指定的收费细目对象设定单据指点定行的收费细目(新增的或修改)
'说明：
'      1.用于新输入或更改收费细目行！！！
'      2.当bytParent<>0时,则为设置从属项目,从属项目一定是新增行,且主项目一定存在
    Dim tmpIncomes As New BillInComes
    Dim dblTime As Double, i As Long
        
     '执行科室
    If bytParent <> 0 Then
        '从属项目的执行科室,如果类别与主项相同,或设为无明确执行科室,则取主项执行科室,否则取本身的
        If lngDoUnit <> 0 Then
            lngDoUnit = mobjBill.Details(bytParent).执行部门ID
        Else
            If cbo开单科室.ListIndex <> -1 Then lngDoUnit = cbo开单科室.ItemData(cbo开单科室.ListIndex)
            
            lngDoUnit = Get收费执行科室ID("Z", Detail.ID, Detail.执行科室, lngDoUnit, Get开单科室ID, Get病人来源, , mobjBill.病区ID)
        End If
    Else
        lngDoUnit = mobjBill.科室ID
        If lngDoUnit = 0 And cbo开单科室.ListIndex <> -1 Then
            lngDoUnit = cbo开单科室.ItemData(cbo开单科室.ListIndex)
        End If
        lngDoUnit = Get收费执行科室ID("Z", Detail.ID, Detail.执行科室, lngDoUnit, Get开单科室ID, Get病人来源, , mobjBill.病区ID)
    End If
    
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
                '初始数次
                If Detail.固有从属 = 0 Then '非固有从属
                    dblTime = mobjBill.Details(bytParent).数次
                ElseIf Detail.固有从属 = 1 Then '固定的固有从属
                    dblTime = Detail.从项数次
                ElseIf Detail.固有从属 = 2 Then '按比例的固有从属
                    dblTime = Detail.从项数次 * mobjBill.Details(bytParent).数次
                End If
            Else
                dblTime = 1
            End If
            mobjBill.Details.Add Detail, .ID, CByte(lngRow), CInt(bytParent), 0, 0, 0, 0, "", "", "", _
            0, 0, mobjBill.费别, 0, .类别, .计算单位, "", 1, dblTime, 0, lngDoUnit, tmpIncomes
        End With
    Else
        '如果该行已经存在,则修改
        With mobjBill.Details(lngRow)
            Set .Detail = Detail
            Set .InComes = tmpIncomes
            .费别 = mobjBill.费别
            .付数 = 1
            .附加标志 = 0
            .计算单位 = Detail.计算单位
            .收费类别 = Detail.类别
            .收费细目ID = Detail.ID
            .数次 = 1
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
    Dim strSQL As String, i As Long
    Dim objDetail As New Detail
    
    Set GetSubDetails = New Details
    
    strSQL = _
        "Select A.ID,A.类别,B.名称 as 类别名称,A.费用类型,A.编码,A.名称,A.规格,A.要求审批," & _
        " A.计算单位,A.屏蔽费别,A.是否变价,A.加班加价,A.执行科室,A.服务对象,C.固有从属,C.从项数次" & _
        " From 收费项目目录 A,收费项目类别 B,收费从属项目 C" & _
        " Where B.编码=A.类别 And C.从项ID=A.ID And A.类别='Z' And C.主项ID=[1]" & _
        " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng项目id)
    For i = 1 To rsTmp.RecordCount
        Set objDetail = New Detail
        With objDetail
            .ID = rsTmp!ID
            .编码 = rsTmp!编码
            .变价 = Nvl(rsTmp!是否变价, 0) = 1
            .规格 = Nvl(rsTmp!规格)
            .计算单位 = Nvl(rsTmp!计算单位)
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
            .要求审批 = Nvl(rsTmp!要求审批, 0) = 1
            GetSubDetails.Add .ID, .药名ID, .类别, .类别名称, .名称, .编码, .简码, .别名, .规格, .计算单位, .说明, .屏蔽费别, _
                1, .计算单位, .分批, .变价, .加班加价, .执行科室, .服务对象, .类型, .补充摘要, .固有从属, .从项数次, , , , , , , .要求审批
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
            Bill.RowData(lngRow) = 0
        Next
    Else
        Bill.RemoveMSFItem lngRow
    End If
End Sub

Private Sub NewBill(Optional blnPati As Boolean = True)
'功能：初始化一张新的单据(程序对象)
'参数：blnPati=是否初始化病人信息
    Dim blnKeepDate As Boolean
    Dim Curdate As Date     '服务器当前时间
    mcurModiMoney = 0
            
    If mrsInfo.State = 0 Then txtPatient.ForeColor = Me.ForeColor
    
    If blnPati Then
        cmdOK.Tag = "": cmdCancel.Tag = "": txt实收.Tag = ""
        txtPatient.Text = "": txtOld.Text = "": txtSex.Text = ""
        txt费别.Text = "": txt床号.Text = "": txt医疗付款.Text = ""
        txt担保人.Text = "": txt担保额.Text = ""
                
        Set mrsMedAudit = Nothing
        Set mrsInfo = New ADODB.Recordset
        sta.Panels(3).Text = ""
    End If
            
    mstrWarn = ""
    cboNO.Text = ""
    Set mobjBill = New ExpenseBill
        
    Curdate = zlDatabase.Currentdate
    chk加班.Value = IIf(OverTime(Curdate), 1, 0)
    
    If Not blnPati And mrsInfo.State = 1 Then
        If mrsInfo!出院日期 < Curdate Then blnKeepDate = True
    End If
    If Not blnKeepDate Then txtDate.Text = Format(Curdate, "yyyy-MM-dd HH:mm:ss")
    
    Call LoadPatientBaby(cboBaby, 0, 0)
    
    Call cbo开单科室_Click
    
    With mobjBill
        .门诊标志 = 2
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
    End With
End Sub

Private Function SaveBill() As Boolean
'功能:保存当前输入的记帐单据(适用住院记帐、划价、或对两者的修改)
'入口:mobjBill=单据对象
'出口:保存是否成功
    Dim i As Long, j As Long, arrSQL As Variant
    Dim int序号 As Integer, int行号 As Integer, strNO As String, strTmp As String
    Dim intParent As Integer, intParentNO As Integer
    Dim str消息 As String, intInsure As Integer, blnTrans As Boolean
    
    mobjBill.NO = zlDatabase.GetNextNo(14)
    gstrModiNO = mobjBill.NO
    arrSQL = Array()
    
    For Each mobjBillDetail In mobjBill.Details
        intParent = 0: intParentNO = int序号
        For Each mobjBillIncome In mobjBillDetail.InComes
            int序号 = int序号 + 1 '当前记录序号
            
            '单据主体
            With mobjBill
                gstrSQL = "zl_住院记帐记录_INSERT('" & .NO & "'," & int序号 & "," & .病人ID & "," & IIf(.主页ID = 0, "NULL", .主页ID) & "," & _
                    IIf(Val(.标识号) = 0, "NULL", .标识号) & "," & "'" & .姓名 & "','" & .性别 & "','" & .年龄 & "','" & .床号 & "','" & .费别 & "'," & _
                    IIf(.病区ID = 0, .开单部门ID, .病区ID) & "," & IIf(.科室ID = 0, .开单部门ID, .科室ID) & "," & .加班标志 & "," & .婴儿费 & "," & .开单部门ID & ",'" & .开单人 & "',"
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
                
                gstrSQL = gstrSQL & IIf(.付数 = 0, 1, .付数) & "," & .数次 & "," & .附加标志 & "," & .执行部门ID & ","
            End With
            
            '收入项目部份
            With mobjBillIncome
                intParent = intParent + 1
                gstrSQL = gstrSQL & IIf(intParent = 1, "Null", intParentNO + 1) & "," & .收入项目ID & "," & _
                    "'" & .收据费目 & "'," & .标准单价 & "," & .应收金额 & "," & .实收金额 & "," & _
                    IIf(.统筹金额 = 0, "NULL", .统筹金额) & ","
            End With
                                            
            '其它部分:最后标记为是简单记帐(发药窗口)
            gstrSQL = gstrSQL & _
                "To_Date('" & Format(mobjBill.发生时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                "To_Date('" & Format(mobjBill.登记时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                "'" & mstrInNO & "'," & IIf(gbytBilling = 1, 1, 0) & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & _
                "NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,1,'" & mobjBillDetail.Detail.类型 & "')"
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = mobjBillDetail.收费细目ID & ";" & gstrSQL
        Next
    Next
    
    '修改前退除原单据
    If mstrInNO <> "" Then
        '先判断是否医保病人记的帐,并作合法性检查(进入修改时已作了一次相关判断)
        If gbytBilling = 0 Then
            intInsure = BillExistInsure(mstrInNO)
            If intInsure > 0 Then
                '去掉了医保连接匹配检查
            End If
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
        
        '执行SQL语句
        On Error GoTo errH
        gcnOracle.BeginTrans: blnTrans = True
            For i = 0 To UBound(arrSQL)
                Call zlDatabase.ExecuteProcedure(Mid(arrSQL(i), InStr(arrSQL(i), ";") + 1), Me.Caption)
            Next
            
            '医保接口
            '1.医保记帐作废上传
            If mstrInNO <> "" And gbytBilling = 0 And intInsure <> 0 Then
                If gclsInsure.GetCapability(support记帐作废上传, , intInsure) And Not gclsInsure.GetCapability(support记帐完成后上传, , intInsure) Then
                    If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                        gcnOracle.RollbackTrans: Exit Function
                    End If
                End If
            End If
                    
            '2.记帐实时上传
            If gbytBilling = 0 And Not IsNull(mrsInfo!险类) Then
                '医保传输费用明细
                If gclsInsure.GetCapability(support记帐上传, , mrsInfo!险类) And Not gclsInsure.GetCapability(support记帐完成后上传, , mrsInfo!险类) Then
                    str消息 = ""
                    If Not gclsInsure.TranChargeDetail(2, mobjBill.NO, 2, 1, str消息, , mrsInfo!险类) Then
                        gcnOracle.RollbackTrans
                        If str消息 <> "" Then MsgBox str消息, vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            End If
        gcnOracle.CommitTrans: blnTrans = False
        
        '1.医保记帐作废上传
        If mstrInNO <> "" And gbytBilling = 0 And intInsure <> 0 Then
            If gclsInsure.GetCapability(support记帐作废上传, , intInsure) And gclsInsure.GetCapability(support记帐完成后上传, , intInsure) Then
                If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                    MsgBox "单据""" & mstrInNO & """的销帐数据向医保传送失败,该单据已销帐！", vbInformation, gstrSysName
                End If
            End If
        End If
                
        '2.记帐实时上传
        If gbytBilling = 0 And Not IsNull(mrsInfo!险类) Then
            '医保传输费用明细
            If gclsInsure.GetCapability(support记帐上传, , mrsInfo!险类) And gclsInsure.GetCapability(support记帐完成后上传, , mrsInfo!险类) Then
                str消息 = ""
                If Not gclsInsure.TranChargeDetail(2, mobjBill.NO, 2, 1, str消息, , mrsInfo!险类) Then
                    If str消息 <> "" Then
                        MsgBox str消息, vbInformation, gstrSysName
                    Else
                        MsgBox "单据""" & mobjBill.NO & """的数据向医保传送失败,该单据已保存！", vbInformation, gstrSysName
                    End If
                End If
            End If
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
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ReadBill(ByVal strNO As String, Optional blnDelete As Boolean) As Boolean
'功能：根据单据号读取一张单据并将其填入表格
'参数：strNO=单据号
'      blnDelete=True:销帐单据时调用,False:查阅单据时调用
    Dim rsTmp As ADODB.Recordset
    Dim rsPatiMoney As ADODB.Recordset
    Dim curTotal As Currency, cur应收Total As Currency
    Dim intInsure As Integer, blnDo As Boolean
    Dim strSQL1 As String, intSign As Integer
    Dim strSQL As String, i As Long
    
    On Error GoTo errH
    
    mblnPrint = False
        
    Call ClearRows: Call Bill.ClearBill: Call ClearMoney
    
    '读取单据主体
    strNO = GetFullNO(strNO, 14)
   
    strSQL = _
    " Select A.病人ID,Nvl(A.主页ID,0) as 主页ID,A.姓名,A.性别,A.年龄,A.费别,A.床号," & _
    "       A.病人病区ID,A.开单部门ID,Nvl(A.加班标志,0) as 加班标志,Nvl(A.婴儿费,0) as 婴儿费," & _
    "       A.开单人,A.划价人,A.操作员姓名,A.发生时间,A.结帐ID,B.担保人,B.担保额" & _
    " From " & IIf(mblnNOMoved And gbytBilling = 0, zlGetFullFieldsTable("住院费用记录"), "住院费用记录 A") & " ,病人信息 B,人员表 C " & _
    " Where NO=[1] And A.记录性质=2 And A.门诊标志=2 And Nvl(A.多病人单,0)=0 And Nvl(A.操作员姓名,A.划价人)=C.姓名" & _
    "       And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & vbNewLine & _
    "       And A.病人ID=B.病人ID And Rownum=1 And A.记录状态" & IIf(mblnDelete, "=2", " IN(0,1,3)") & _
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
        MsgBox "没有发现该单据！", vbInformation, gstrSysName
        Exit Function
    Else
        If mbytUseType = 0 Or mbytUseType = 1 Then
            If InStr(mstrPrivs, ";所有病区;") = 0 And mlngUnitID > 0 Then
                If InStr(1, "," & mstrUnitIDs & ",", "," & IIf(IsNull(rsTmp!病人病区ID), 0, rsTmp!病人病区ID) & ",") = 0 Then
                    MsgBox "你没有权限读取其它病区的单据！", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        ElseIf mbytUseType = 2 Then
            If InStr(mstrPrivs, ";所有科室;") = 0 And mlngDeptID > 0 Then
                If IIf(IsNull(rsTmp!开单部门ID), 0, rsTmp!开单部门ID) <> mlngDeptID Then
                    MsgBox "你没有权限读取其它科室开单的单据！", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    End If

    '单据号
    cboNO.Text = strNO

    '姓名
    txtPatient.Text = IIf(IsNull(rsTmp!姓名), "", rsTmp!姓名)
    
    '性别
    txtSex.Text = IIf(IsNull(rsTmp!性别), "", rsTmp!性别)
    '年龄
    txtOld.Text = IIf(IsNull(rsTmp!年龄), "", rsTmp!年龄)
    txt床号.Text = IIf(IsNull(rsTmp!床号), "", rsTmp!床号)
    
    txt担保人.Text = IIf(IsNull(rsTmp!担保人), "", rsTmp!担保人)
    txt担保额.Text = Format(IIf(IsNull(rsTmp!担保额), "", rsTmp!担保额), "0.00")
    
    '费别
    txt费别.Text = IIf(IsNull(rsTmp!费别), "", rsTmp!费别)
    txt医疗付款.Text = Get病人医疗付款方式(rsTmp!病人ID, rsTmp!主页ID)
    If gintPriceGradeStartType >= 2 Then
        Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, Val(Nvl(rsTmp!病人ID)), Val(Nvl(rsTmp!主页ID)), txt医疗付款.Text, mstr药品价格等级, mstr卫材价格等级, mstr普通价格等级)
    End If
    
    txtDate.Text = Format(rsTmp!发生时间, "yyyy-MM-dd HH:mm:ss")
    chk加班.Value = IIf(IsNull(rsTmp!加班标志), 0, rsTmp!加班标志)
    Call LoadPatientBaby(cboBaby, rsTmp!病人ID, rsTmp!主页ID)
    Call zlControl.CboLocate(cboBaby, rsTmp!婴儿费, True)
        
    Call Set开单人开单科室(cbo开单人, cbo开单科室, mrs开单人, mrs开单科室, Nvl(rsTmp!开单人), Nvl(rsTmp!开单部门ID, 0))
    
    '病人费用信息
    If Not IsNull(rsTmp!病人ID) Then
        Set rsPatiMoney = GetMoneyInfo(rsTmp!病人ID, , True, 2)
        If Not rsPatiMoney Is Nothing Then
            sta.Panels(3).Text = "预交:" & Format(rsPatiMoney!预交余额, "0.00") & _
            "/费用:" & Format(rsPatiMoney!费用余额, gstrDec) & _
            "/剩余:" & Format(rsPatiMoney!预交余额 - rsPatiMoney!费用余额, "0.00")
        End If
    End If
    
    '-----------------------------------------------------------------
    If blnDelete Then
         '销帐单无需考虑后备表,前面的操作已禁止
        '读取准退数,并计算应收金额,实收金额(金额=剩余金额*(准退数/剩余数))
        '整张单据汇总结果(明细到收费细目)
        '执行状态应该在原始记录上判断(部分退药且部分退费的记录)
        strSQL = "Select Nvl(价格父号,序号) From 住院费用记录 " & _
            " Where 记录性质=2 And 门诊标志=2 And Nvl(多病人单,0)=0" & _
            " And 记录状态 IN(0,1,3) And NO=[1] And Nvl(执行状态,0)<>1" & _
            IIf(mstrTime <> "", " And 登记时间=[2]", "")
            
        '如果已结帐单据禁止销帐,或是医保记帐的单据。则在原始单据行中只取未结帐部分
        intInsure = BillExistInsure(strNO)
        If intInsure <> 0 Then
            blnDo = Not gclsInsure.GetCapability(support允许冲销已结帐的记帐单据, , intInsure)
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
            " Select A.记录状态,Nvl(A.价格父号,A.序号) as 序号," & _
            " C.编码,C.名称 as 类别,B.名称,B.规格,Nvl(A.费用类型,B.费用类型) 费用类型,A.计算单位," & _
            " Avg(Nvl(A.付数,1)*A.数次) as 数量,Sum(A.标准单价) as 单价," & _
            " Sum(A.应收金额) as 应收金额,Sum(A.实收金额) as 实收金额, " & _
            " D.名称 as 执行部门,A.附加标志" & _
            " From 住院费用记录 A,收费项目目录 B,收费项目类别 C,部门表 D " & _
            " Where A.收费细目ID=B.ID And C.编码=A.收费类别 And A.执行部门ID=D.ID " & _
            " And A.记录性质=2 And A.门诊标志=2 And Nvl(A.多病人单,0)=0" & _
            " And A.NO=[1] And Nvl(A.价格父号,A.序号) IN(" & strSQL & ")" & _
            " Group by A.记录状态,Nvl(A.价格父号,A.序号),C.编码,C.名称,B.名称," & _
            " B.规格,Nvl(A.费用类型,B.费用类型),A.计算单位,D.名称,A.附加标志"
            
        '最后计算结果(剩余数量即为准退数量,不必计算)
        '排开已经全部退费的行(执行状态=0的一种可能)
        strSQL = _
            " Select A.序号,A.编码,A.类别,A.名称,A.规格," & _
            " A.费用类型,A.计算单位,A.执行部门,A.附加标志," & _
            " Sum(A.数量) as 数量,A.单价,Sum(A.应收金额) as 应收金额,Sum(A.实收金额) as 实收金额" & _
            " From (" & strSQL & ") A" & _
            " Group by A.序号,A.编码,A.类别,A.名称,A.规格,A.费用类型," & _
            " A.计算单位,A.单价,A.执行部门,A.附加标志" & _
            " Having Sum(A.数量)<>0" & _
            " Order by A.序号"
    ElseIf mbytInState = 0 And gbytBilling = 2 Then
        '读取记帐划价单(记帐审核),只读取剩余数量,金额
        '划价单不涉及后备表
        strSQL = _
            " Select Nvl(A.价格父号,A.序号) as 序号,C.编码,C.名称 as 类别," & _
            " B.名称,B.规格,Nvl(A.费用类型,B.费用类型) 费用类型,A.计算单位,Avg(Nvl(A.付数,1)*A.数次) as 数量," & _
            " Sum(A.标准单价) as 单价,Sum(A.应收金额) as 应收金额,Sum(A.实收金额) as 实收金额, " & _
            " D.名称 as 执行部门,A.附加标志" & _
            " From 住院费用记录 A,收费项目目录 B,收费项目类别 C,部门表 D " & _
            " Where A.记录状态=0 And A.收费细目ID=B.ID And C.编码=A.收费类别 And A.执行部门ID=D.ID " & _
            " And A.记录性质=2 And Nvl(A.多病人单,0)=0 And 门诊标志=2 And A.NO=[1]" & _
            " Group by Nvl(A.价格父号,A.序号),A.记录状态,C.编码,C.名称,B.名称,B.规格," & _
            " Nvl(A.费用类型,B.费用类型),A.计算单位,D.名称,A.附加标志"
    Else
        '读取单据原始内容
        intSign = IIf(mblnDelete, -1, 1) '数量,金额正负符号
        strSQL = _
            " Select Nvl(A.价格父号,A.序号) as 序号," & _
            " C.编码,C.名称 as 类别,B.名称,B.规格,Nvl(A.费用类型,B.费用类型) 费用类型,A.计算单位," & _
            " Avg(" & intSign & "*Nvl(A.付数,1)*A.数次) as 数量," & _
            " Sum(A.标准单价) as 单价,Sum(" & intSign & "*A.应收金额) as 应收金额, " & _
            " Sum(" & intSign & "*A.实收金额) as 实收金额, " & _
            " D.名称 as 执行部门,A.附加标志" & _
            " From " & IIf(mblnNOMoved, zlGetFullFieldsTable("住院费用记录"), "住院费用记录 A") & " ,收费项目目录 B,收费项目类别 C,部门表 D " & _
            " Where A.收费细目ID=B.ID And C.编码=A.收费类别 And A.执行部门ID=D.ID " & _
            " And A.记录性质=2 And A.门诊标志=2 And Nvl(A.多病人单,0)=0 And A.NO=[1]" & _
            " And A.记录状态" & IIf(mblnDelete, "=2", " IN(0,1,3)") & _
            IIf(mstrTime <> "", " And A.登记时间=[2]", "") & _
            " Group by Nvl(A.价格父号,A.序号),C.编码,C.名称,B.名称," & _
            " B.规格,Nvl(A.费用类型,B.费用类型),A.计算单位,D.名称,A.附加标志" & _
            " Order by 序号"
    End If
    
    If mstrTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, CDate(mstrTime))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    End If
    
    If rsTmp.EOF Then Exit Function
    
    Bill.Redraw = False
    Bill.Rows = rsTmp.RecordCount + 1
    For i = 1 To rsTmp.RecordCount
        If gbytBilling = 2 And Not mblnPrint Then mblnPrint = True
    
        Bill.RowData(i) = rsTmp!序号 '用于记帐销帐及划价审核
        Bill.TextMatrix(i, 0) = rsTmp!名称
        Bill.TextMatrix(i, 1) = Format(rsTmp!应收金额, gstrDec)
        Bill.TextMatrix(i, 2) = Format(rsTmp!实收金额, gstrDec)
        Bill.TextMatrix(i, 3) = rsTmp!执行部门
        Bill.TextMatrix(i, 4) = IIf(IsNull(rsTmp!费用类型), "", rsTmp!费用类型)
        '设置销帐标志
        If Bill.TextMatrix(0, Bill.Cols - 1) = "销帐" Then
            Bill.TextMatrix(i, Bill.Cols - 1) = "√"
        End If
        rsTmp.MoveNext
    Next
    '针对列编辑性质设置颜色
    Bill.SetColColor 0, &HE7CFBA
    Bill.SetColColor 1, &HE7CFBA
    Bill.SetColColor 3, &HE7CFBA
    Bill.Redraw = True
    
    '-------------------------------------------------------------------------------
    '读取单据收入项目
    If blnDelete Then
         '退费单无需考虑后备表,前面的操作已禁止
        '读取准退数,并计算应收金额,实收金额(金额=剩余金额*(准退数/剩余数))
        '整张费用单据(明细到收入项目)
        '执行状态应该在原始记录上判断(部分退药且部分退费的记录)
        strSQL = "Select Nvl(价格父号,序号) From 住院费用记录 " & _
            " Where 记录性质=2 And 门诊标志=2 And Nvl(多病人单,0)=0" & _
            " And 记录状态 IN(0,1,3) And NO=[1] And Nvl(执行状态,0)<>1" & _
            IIf(mstrTime <> "", " And 登记时间=[2]", "")
            
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
        strSQL = _
            " Select A.序号,A.名称," & _
                " Sum(A.数量) as 剩余数量,Sum(A.应收金额) as 剩余应收," & _
                " Sum(A.实收金额) as 剩余实收" & _
            " From (" & _
                " Select A.记录状态,A.序号,B.名称," & _
                " Nvl(A.付数,1)*A.数次 as 数量,A.应收金额,A.实收金额" & _
                " From 住院费用记录 A,收入项目 B" & _
                " Where A.记录性质=2 And A.门诊标志=2 And Nvl(A.附加标志,0)<>9 And Nvl(A.多病人单,0)=0" & _
                    " And A.NO=[1] And Nvl(A.价格父号,A.序号) IN(" & strSQL & ")" & _
                    " And A.收入项目ID=B.ID" & _
                " ) A" & _
            " Group by A.序号,A.名称 Having Sum(数量)<>0"
                    
        '最后计算结果(准退数量即剩余数量,不必真正计算)
        strSQL = _
            " Select A.名称,Sum(A.剩余应收) as 应收金额," & _
            " Sum(A.剩余实收) as 实收金额" & _
            " From (" & strSQL & ") A" & _
            " Group by A.名称"
    ElseIf mbytInState = 0 And gbytBilling = 2 Then
        '读取记帐划价单(记帐审核),只读取未审核部份
        '划价单不涉及后备表
        strSQL = _
            "Select B.名称,Sum(A.应收金额) as 应收金额," & _
            " Sum(A.实收金额) as 实收金额 " & _
            " From 住院费用记录 A,收入项目 B" & _
            " Where A.记录状态=0 And A.记录性质=2 And A.门诊标志=2" & _
            " And Nvl(A.多病人单,0)=0 And A.NO=[1] And A.收入项目ID=B.ID" & _
            " Group By B.名称"
    Else
        '读取单据原始内容
        intSign = IIf(mblnDelete, -1, 1) '数量,金额正负符号
        strSQL = _
            "Select B.名称," & _
            " Sum(" & intSign & "*A.应收金额) as 应收金额," & _
            " Sum(" & intSign & "*A.实收金额) as 实收金额 " & _
            " From " & IIf(mblnNOMoved, zlGetFullFieldsTable("住院费用记录"), "住院费用记录 A") & " ,收入项目 B" & _
            " Where A.收入项目ID=B.ID And A.记录状态" & IIf(mblnDelete, "=2", " IN(0,1,3)") & _
            " And A.记录性质=2 And A.门诊标志=2 And Nvl(A.多病人单,0)=0 And A.NO=[1]" & _
            IIf(mstrTime <> "", " And A.登记时间=[2]", "") & _
            " Group By B.名称"
    End If
    
    If mstrTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, CDate(mstrTime))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    End If
    
    If rsTmp.EOF Then Exit Function
    
    '刷新显示(收费要叠加)
    mshMoney.Rows = rsTmp.RecordCount + 1
    If mshMoney.Rows < 4 Then mshMoney.Rows = 4
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

Private Sub ClearRows()
    Dim i As Long
    For i = 1 To Bill.Rows - 1
        Bill.RowData(i) = 0
    Next
End Sub


Private Sub FillBillComboBox(lngRow As Long, lngCol As Long)
'功能：根据单据列设置下拉列表框内容
    Dim rsTmp As New ADODB.Recordset
    Dim str人员性质 As String, strTmp As String
    Dim strSQL As String, i As Long
    Dim lng病区ID As Long, lng科室ID As Long
    
    Bill.Clear
    
    On Error GoTo errHandle
    

    Select Case Bill.TextMatrix(0, lngCol)
        Case "执行科室"
            Bill.cboStyle = DropDownAndEdit
            
            '根据当前项目执行科室性质,动态设置可选科室
            If mobjBill.Details.Count >= lngRow Then
                With mobjBill.Details(lngRow)
                    Bill.TextMatrix(lngRow, lngCol) = ""
                    
                    lng科室ID = mobjBill.科室ID
                    If lng科室ID = 0 Then lng科室ID = Get开单科室ID
                    
                    lng病区ID = mobjBill.病区ID
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
                            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .收费细目ID, Get病人来源, lng科室ID)
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
                        For i = 1 To mrsUnit.RecordCount
                            strTmp = IIf(zlIsShowDeptCode, mrsUnit!编码 & "-", "") & mrsUnit!名称
                            '刘兴洪:28947
                            If zlCboFindItem(Bill.cboObj, Val(Nvl(mrsUnit!ID))) = False Then
                            
                            'If Not (SendMessage(Bill.cboHwnd, CB_FINDSTRING, -1, ByVal strTmp) >= 0) Then
                                Bill.AddItem strTmp
                                Bill.ItemData(Bill.ListCount - 1) = mrsUnit!ID
                                
                                '设置缺省执行科室
                                If lngRow = 1 Then
                                    If mrsUnit!ID = lng科室ID Then Bill.ListIndex = Bill.NewIndex
                                ElseIf lngRow > 1 Then
                                    If mrsUnit!ID = mobjBill.Details(lngRow - 1).执行部门ID And mobjBill.Details(lngRow - 1).Detail.执行科室 = .Detail.执行科室 Then
                                        Bill.ListIndex = Bill.NewIndex
                                    ElseIf mrsUnit!ID = lng科室ID And Bill.ListIndex = -1 Then
                                        Bill.ListIndex = Bill.NewIndex
                                    End If
                                End If
                            End If
                            mrsUnit.MoveNext
                        Next
                        
                        If .Detail.执行科室 = 4 Then    '执行科室为指定科室的,缺省为操作员所在科室
                            For i = 0 To Bill.ListCount - 1
                                If Bill.ItemData(i) = UserInfo.部门ID Then Bill.ListIndex = i: Exit For
                            Next
                        End If
                        
                        If Bill.ListIndex = -1 Then '如果没有则取现有的执行科室
                            For i = 0 To Bill.ListCount - 1
                                If Bill.ItemData(i) = .执行部门ID Then Bill.ListIndex = i: Exit For
                            Next
                        End If
                    End If
                    
                    If Bill.ListIndex = -1 And Bill.ListCount > 0 Then Bill.ListIndex = 0
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

Private Sub SetMoneyList()
'功能:根据当前收入项目行数调整各列宽
    Dim lngW As Long
    lngW = mshMoney.Width - 75
    If mshMoney.Rows > mshMoney.Height / mshMoney.RowHeight(0) Then
        lngW = lngW - 250
    End If
    
    mshMoney.ColWidth(0) = lngW * 0.45
    mshMoney.ColWidth(1) = lngW * 0.55
    
    mshMoney.ColAlignment(0) = 1
    mshMoney.ColAlignment(1) = 7
    mshMoney.ColAlignmentFixed(0) = 4
    mshMoney.ColAlignmentFixed(1) = 4
    
    mshMoney.TextMatrix(0, 0) = "项目"
    mshMoney.TextMatrix(0, 1) = "金额"
    mshMoney.Row = 0
End Sub

Public Sub ShowMoney()
'功能：刷新显示收入项目费用区
    Dim i As Long, j As Long, k As Long
    Dim blnExist As Boolean, curTotal As Currency, cur应收Total As Currency
    mshMoney.Redraw = False
    
    '清除显示
    For i = 1 To mshMoney.Rows - 1
        For j = 0 To mshMoney.Cols - 1
            mshMoney.TextMatrix(i, j) = ""
        Next
    Next
    
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
    mshMoney.Rows = IIf(mcolMoneys.Count = 0, 2, mcolMoneys.Count + 1)
    If mshMoney.Rows < 4 Then mshMoney.Rows = 4
    Call SetMoneyList
    
    For i = 1 To mcolMoneys.Count
        mshMoney.TextMatrix(i, 0) = mcolMoneys(i).收入项目
        mshMoney.TextMatrix(i, 1) = Format(mcolMoneys(i).实收金额, gstrDec)
        curTotal = curTotal + mcolMoneys(i).实收金额
        cur应收Total = cur应收Total + mcolMoneys(i).应收金额
    Next
    txt实收.Text = Format(curTotal, gstrDec)
    txt应收.Text = Format(cur应收Total, gstrDec)
    
    mshMoney.TopRow = mshMoney.Rows - 1
    
    mshMoney.Redraw = True
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
    mshMoney.Rows = 4
    mshMoney.Redraw = True
    
    '20030617:不清除金额
    'txt实收.Text = gstrdec
    'txt应收.Text = gstrdec
End Sub

Private Sub ShowDeleteCol(blnShow As Boolean)
'功能：显示\隐藏销帐标志列
    Dim i As Long, blnACT As Boolean
    If blnShow Then
        If Bill.TextMatrix(0, Bill.Cols - 1) <> "销帐" Then
            Bill.Redraw = False
            Bill.Cols = Bill.Cols + 1
            Bill.TextMatrix(0, Bill.Cols - 1) = "销帐"
            Bill.ColAlignment(Bill.Cols - 1) = 4
            Bill.ColWidth(Bill.Cols - 1) = 550
            Bill.ColData(Bill.Cols - 1) = -1
            
            blnACT = Bill.Active: Bill.Active = False
            Bill.Row = 0: Bill.Col = Bill.Cols - 1: Bill.MsfObj.CellForeColor = vbRed
            Bill.Row = 1: Bill.Col = Bill.Cols - 1
            Bill.Active = blnACT
            
            Bill.ColWidth(0) = GetOrigColWidth(0) - 300
            Bill.ColWidth(3) = GetOrigColWidth(3) - 250
            Bill.Redraw = True
        End If
    Else
        If Bill.TextMatrix(0, Bill.Cols - 1) = "销帐" Then
            Bill.Redraw = False
            Bill.Cols = Bill.Cols - 1
            Bill.ColWidth(0) = GetOrigColWidth(0)
            Bill.ColWidth(3) = GetOrigColWidth(3)
            Bill.Redraw = True
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
    '  No_In       门诊费用记录.NO%Type,
    '  记录性质_In 门诊费用记录.记录性质%Type,
    '  开单人_In   门诊费用记录.开单人%Type,
    '  发生时间_In 门诊费用记录.发生时间%Type,
    '  姓名_In     门诊费用记录.姓名%Type := Null,
    '  来源_In Integer:=1
  
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

Private Function Check费用类型(Optional intRow As Integer) As Boolean
'功能：根据当前病人的类型判断指定行的项目是否可以输入,适用于所有类别的项目
    Dim strSQL As String
    Dim i As Long, bytType As Byte
    Dim rsTmp As New ADODB.Recordset
    Dim bln医保 As Boolean, bln公费 As Boolean
    
    Check费用类型 = True
    
    On Error GoTo errHandle
    

    '无法检查
    If txt医疗付款.Text = "" Then Exit Function
    
    '医保或公费病人
    '问题:45605
    If zlIsCheckMedicinePayMode(txt医疗付款.Text, bln医保, bln公费) = False Then Exit Function
    '确定病人类型
    bytType = IIf(bln医保, 1, 2)
    
    '读取检查数据
    If bytType = 1 Then
        strSQL = "Select 编码,名称,简码,性质,缺省标志 From 费用类型 Where 编码 In(" & gstr医保费用类型 & ") Order by 编码"
    Else
        strSQL = "Select 编码,名称,简码,性质,缺省标志 From 费用类型 Where 编码 In(" & gstr公费费用类型 & ") Order by 编码"
    End If
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
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
                    IIf(bytType = 1, "医保", "公费") & "费用类型！", vbInformation, gstrSysName
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
                        IIf(bytType = 1, "医保", "公费") & "费用类型！" & vbCrLf & "确实要保存单据吗？", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Check费用类型 = False: Exit For
                    End If
                End If
            End If
        Next
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub ReCalcInsure()
'功能：修改单据时,重新计算统筹金额及更新相关信息
    Dim i As Long, j As Long
    Dim strInfo As String
    
    If mrsInfo.State = 1 Then
        If Not IsNull(mrsInfo!险类) Then
            For i = 1 To mobjBill.Details.Count
                For j = 1 To mobjBill.Details(i).InComes.Count
                    strInfo = gclsInsure.GetItemInsure(mobjBill.病人ID, mobjBill.Details(i).收费细目ID, mobjBill.Details(i).InComes(j).实收金额, False, mrsInfo!险类, _
                     mobjBill.Details(i).摘要 & "||" & mobjBill.Details(i).数次)
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
    End If
End Sub

Private Function Check执行科室() As Integer
    Dim i As Long
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).执行部门ID = 0 Or Bill.TextMatrix(i, 3) = "" Then
            Check执行科室 = i: Exit Function
        End If
    Next
End Function

Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        lngTXTProc = GetWindowLong(txtPatient.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, lngTXTProc)
    End If
End Sub

Private Function Check服务对象() As Integer
'功能：检查当前病人的记帐费用项目的服务对象是否一致
'说明：因为加入了门诊留观病人,所以有此检查
'返回：不一致的费用行,为0时正常
    Dim i As Integer
    
    If mrsInfo.State = 0 Then Exit Function
    For i = 1 To mobjBill.Details.Count
        If mrsInfo!病人性质 = 0 Or mrsInfo!病人性质 = 2 Then
            '住院病人或住院留观病人,不能用只服务于门诊的项目
            If mobjBill.Details(i).Detail.服务对象 = 1 Then
                MsgBox "第 " & i & " 行项目""" & mobjBill.Details(i).Detail.名称 & """仅服务于门诊,该病人不能使用.", vbInformation, gstrSysName
                Check服务对象 = i: Exit Function
            End If
        ElseIf mrsInfo!病人性质 = 1 Or mrsInfo!病人性质 = -1 Then
            '门诊或出院病人(医技记帐)或门诊留观病人,不能用只服务于住院的项目
            If mobjBill.Details(i).Detail.服务对象 = 2 Then
                MsgBox "第 " & i & " 行项目""" & mobjBill.Details(i).Detail.名称 & """仅服务于住院,该病人不能使用.", vbInformation, gstrSysName
                Check服务对象 = i: Exit Function
            End If
        End If
    Next
End Function

Private Sub txtPatient_Validate(Cancel As Boolean)
    If IsNumeric(txtPatient.Tag) And mrsInfo.State = 1 Then
        mblnValid = True
        Call txtPatient_KeyPress(13)
        mblnValid = False
    End If
End Sub
Private Function Get开单科室ID() As Long
    If cbo开单科室.ListIndex <> -1 Then
        Get开单科室ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
    Else
        Get开单科室ID = UserInfo.部门ID
    End If
End Function
Private Function Get病人来源() As Integer
'功能：获取当前病人的来源(因为可以对门诊留观病人记帐)
    If mrsInfo.State = 1 Then
        If mrsInfo!病人性质 = 0 Or mrsInfo!病人性质 = 2 Then
            Get病人来源 = 2
        ElseIf mrsInfo!病人性质 = 1 Or mrsInfo!病人性质 = -1 Then
            Get病人来源 = 1 '门诊病人(医技记帐)或门诊留观病人
        End If
    Else
        Get病人来源 = 2 '缺省为2
    End If
End Function
Private Sub initCardSquareData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取结算卡对象的相关信息
    '入参:blnClosed:关闭对象
    '编制:刘兴洪
    '日期:2010-01-05 14:51:23
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    If gobjSquare.objSquareCard Is Nothing Then Exit Sub
    Call IDKIND.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtPatient)
    Set objCard = IDKIND.GetfaultCard
    Set gobjSquare.objDefaultCard = objCard
    If IDKIND.Cards.按缺省卡查找 And Not objCard Is Nothing Then
        gobjSquare.bln缺省卡号密文 = objCard.卡号密文规则 <> ""
        gobjSquare.int缺省卡号长度 = objCard.卡号长度
    Else
        gobjSquare.bln缺省卡号密文 = IDKIND.Cards.加密显示
        gobjSquare.int缺省卡号长度 = 100
    End If
    gobjSquare.bln按缺省卡查找 = IDKIND.Cards.按缺省卡查找
End Sub

