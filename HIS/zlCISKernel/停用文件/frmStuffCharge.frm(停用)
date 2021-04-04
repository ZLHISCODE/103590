VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Begin VB.Form frmStuffCharge 
   AutoRedraw      =   -1  'True
   Caption         =   "备货材料记帐"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11805
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStuffCharge.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   11805
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   26
      Top             =   7875
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2619
            MinWidth        =   882
            Picture         =   "frmStuffCharge.frx":058A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15584
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   370
            MinWidth        =   88
            Key             =   "病人余额"
            Object.ToolTipText     =   "病人余额"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   370
            MinWidth        =   71
            Key             =   "MedicareType"
            Object.ToolTipText     =   "医保大类"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmStuffCharge.frx":0E1E
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmStuffCharge.frx":1458
            Key             =   "WB"
            Object.ToolTipText     =   "五笔(F7)"
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
      Height          =   2865
      Left            =   0
      ScaleHeight     =   2865
      ScaleWidth      =   11805
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   5010
      Width           =   11805
      Begin MSComctlLib.ImageList imgList 
         Left            =   11070
         Top             =   1980
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
               Picture         =   "frmStuffCharge.frx":1A92
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "取消(&C)"
         Height          =   420
         Left            =   9780
         TabIndex        =   21
         ToolTipText     =   "热键:Esc"
         Top             =   1785
         Width           =   1680
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00C0C0C0&
         Caption         =   "确定(&O)"
         Height          =   420
         Left            =   7965
         TabIndex        =   20
         ToolTipText     =   "热键：F2"
         Top             =   1785
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
         TabIndex        =   36
         ToolTipText     =   "清除:F6"
         Top             =   -90
         Width           =   11880
         Begin VB.ComboBox cboBaby 
            Height          =   360
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   180
            Width           =   1800
         End
         Begin VB.CheckBox chk急诊 
            Caption         =   "急诊费用"
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   4440
            TabIndex        =   15
            Top             =   240
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.CheckBox chk加班 
            Caption         =   "加班(&A)"
            Height          =   270
            Left            =   120
            TabIndex        =   12
            Top             =   225
            Width           =   1170
         End
         Begin VB.ComboBox cbo开单人 
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   6555
            TabIndex        =   16
            Top             =   180
            Width           =   2085
         End
         Begin MSMask.MaskEdBox txtDate 
            Height          =   360
            Left            =   9360
            TabIndex        =   17
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
            TabIndex        =   13
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label lbl开单人 
            AutoSize        =   -1  'True
            Caption         =   "开单人"
            Height          =   240
            Left            =   5790
            TabIndex        =   38
            Top             =   240
            Width           =   720
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            Caption         =   "时间"
            Height          =   240
            Left            =   8820
            TabIndex        =   37
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.Frame fraDrawDept 
         Height          =   720
         Left            =   0
         TabIndex        =   47
         Top             =   360
         Width           =   13575
         Begin VB.ComboBox cbo执行部门 
            Height          =   360
            Left            =   1155
            Style           =   2  'Dropdown List
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   255
            Width           =   2265
         End
         Begin VB.TextBox txt病人备注 
            BackColor       =   &H00E0E0E0&
            Height          =   360
            Left            =   5145
            Locked          =   -1  'True
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   240
            Width           =   2700
         End
         Begin VB.Label lbl执行部门 
            Caption         =   "执行部门"
            Height          =   315
            Left            =   105
            TabIndex        =   53
            Top             =   285
            Width           =   1050
         End
         Begin VB.Label lbl病人备注 
            Caption         =   "病人备注"
            Height          =   225
            Left            =   4155
            TabIndex        =   49
            Top             =   308
            Width           =   1005
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshMoney 
         Height          =   1635
         Left            =   0
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   1200
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
      Begin VB.Frame fraStat 
         Height          =   1770
         Left            =   3510
         TabIndex        =   39
         Top             =   1065
         Width           =   3675
         Begin VB.TextBox txtPreNO 
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
            Left            =   1065
            Locked          =   -1  'True
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   1230
            Width           =   1845
         End
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
            TabIndex        =   19
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   750
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
            TabIndex        =   18
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   250
            Width           =   1845
         End
         Begin VB.Label lblPreNO 
            AutoSize        =   -1  'True
            Caption         =   "上张"
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
            TabIndex        =   51
            Top             =   1298
            Width           =   690
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
            TabIndex        =   41
            Top             =   818
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
            TabIndex        =   40
            Top             =   318
            Width           =   690
         End
      End
   End
   Begin VB.Frame fraTitle 
      Height          =   1095
      Left            =   45
      TabIndex        =   24
      ToolTipText     =   "清除:F6"
      Top             =   -120
      Width           =   11865
      Begin VB.CommandButton cmdSel 
         Caption         =   "导入"
         Height          =   375
         Left            =   75
         TabIndex        =   54
         ToolTipText     =   "快键:F11"
         Top             =   645
         Width           =   855
      End
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
         Caption         =   "备货记帐单"
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
         TabIndex        =   28
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
         TabIndex        =   25
         Top             =   720
         Width           =   720
      End
   End
   Begin VB.Frame fraUnit 
      Height          =   1065
      Left            =   9375
      TabIndex        =   23
      Top             =   855
      Width           =   2505
      Begin VB.ComboBox cbo开单科室 
         Height          =   360
         Left            =   135
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "cbo开单科室"
         Top             =   615
         Width           =   2265
      End
      Begin VB.Label lbl开单科室 
         AutoSize        =   -1  'True
         Caption         =   "开单科室"
         Height          =   240
         Left            =   150
         TabIndex        =   27
         Top             =   270
         Width           =   960
      End
   End
   Begin VB.Frame fraInfo 
      Height          =   1065
      Left            =   30
      TabIndex        =   22
      Top             =   855
      Width           =   9345
      Begin VB.TextBox txt住院号 
         Height          =   360
         Left            =   7905
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   210
         Width           =   1290
      End
      Begin VB.TextBox txt费别 
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   705
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "热键：F11"
         Top             =   615
         Width           =   1545
      End
      Begin VB.TextBox txt付款方式 
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "热键：F11"
         Top             =   615
         Width           =   2085
      End
      Begin VB.TextBox txt性别 
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "热键：F11"
         Top             =   210
         Width           =   795
      End
      Begin VB.TextBox txt担保额 
         Height          =   360
         Left            =   7905
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   615
         Width           =   1290
      End
      Begin VB.TextBox txt担保人 
         Height          =   360
         Left            =   5895
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   615
         Width           =   1110
      End
      Begin VB.TextBox txt床号 
         Height          =   360
         Left            =   5895
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   210
         Width           =   1110
      End
      Begin VB.TextBox txt姓名 
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   705
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   210
         Width           =   1545
      End
      Begin VB.TextBox txt年龄 
         Height          =   360
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   210
         Width           =   765
      End
      Begin VB.Label lbl住院号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号"
         Height          =   240
         Left            =   7140
         TabIndex        =   46
         Top             =   270
         Width           =   720
      End
      Begin VB.Label lbl担保额 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "担保额"
         Height          =   240
         Left            =   7140
         TabIndex        =   45
         Top             =   675
         Width           =   720
      End
      Begin VB.Label lbl担保人 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "担保人"
         Height          =   240
         Left            =   5145
         TabIndex        =   44
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
         Left            =   2400
         TabIndex        =   43
         Top             =   585
         Width           =   420
      End
      Begin VB.Label lbl床号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床号"
         Height          =   240
         Left            =   5385
         TabIndex        =   34
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lbl姓名 
         AutoSize        =   -1  'True
         Caption         =   "病人"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   165
         TabIndex        =   32
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lbl性别 
         AutoSize        =   -1  'True
         Caption         =   "性别"
         Height          =   240
         Left            =   2370
         TabIndex        =   31
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lbl年龄 
         AutoSize        =   -1  'True
         Caption         =   "年龄"
         Height          =   240
         Left            =   3705
         TabIndex        =   30
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lbl费别 
         AutoSize        =   -1  'True
         Caption         =   "费别"
         Height          =   240
         Left            =   150
         TabIndex        =   29
         Top             =   675
         Width           =   480
      End
   End
   Begin ZL9BillEdit.BillEdit Bill 
      Height          =   3105
      Left            =   15
      TabIndex        =   11
      Top             =   1920
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   5477
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
      TabIndex        =   33
      Top             =   0
      Width           =   945
   End
End
Attribute VB_Name = "frmStuffCharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
'入口参数
'————————————————————————————————————————————————————————————————————
Private mlng医嘱ID As Long  '新增费用时用
Private mlng发送号 As Long  '新增费用时用
Private mlng病人ID As Long  '确定要计费的病人ID
Private mlng主页ID As Long  '确定要计费的主页ID
Private mstrPrivs As String
Private mstrPrivsOpt As String '住院记帐操作的相关权限
Private mint病人来源 As Integer  '1-门诊病人,2-住院病人
Private mint记录性质 As Integer  '1-收费(划价),2-记帐(门/住)
Private mstrFeeTab As String
Private mlng虚拟库房ID As Long
Private mbln费用登记 As Boolean  '仅登记,不计实收金额
Private mlng开单科室ID As Long  '为当前主界面医技科室
Private mlng病人科室id As Long  '主要是用于确定门诊病人的科室ID
Private mblnCboNotClick As Boolean
Private mlng开嘱科室ID As Long
Private mstr开嘱医生 As String
Private mblnUnload As Boolean
Private mrsAll开单科室 As ADODB.Recordset

Private mbytInState As Byte  '0-执行,1-查阅,2-调整(不支持),3-删费
Private mstrInNO As String  '所操作的单据号(执行时为修改)
Private mstrOriginalNO As String  '补充主费用时,医嘱发送中的单据号

Private mstrTime As String  '操作单据内容的登记时间
Private mblnDelete As Boolean  '是否处理退费单据(查阅)
Private mblnWarnCloseed As Boolean  '刘兴洪:因费用报敬发生的关闭
Private mblnSendMateria  As Boolean
Private mbytSendMateria As Byte '0-记帐后不发药,1-自动发药,2-提示发药
Private mlng执行库房ID As Long

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
    项目 = 1
    商品名 = 2
    规格 = 3
    单位 = 4
    付数 = 5
    数次 = 6
    单价 = 7
    应收金额 = 8
    实收金额 = 9
    商品条码 = 10
    内部条码 = 11
    类型 = 12
End Enum

Private Enum Pan
    C2提示信息 = 2
End Enum

'当前病人险类的医保支持参数
Private Type TYPE_MedicarePAR
    负数记帐 As Boolean
    记帐上传 As Boolean
    记帐完成后上传 As Boolean
    实时监控 As Boolean
    医生确定处方类型 As Boolean '目前只有北京医保专用
End Type
Private MCPAR As TYPE_MedicarePAR
Private mrsDept As ADODB.Recordset
'医技工作站本地费用参数
'————————————————————————————————————————————————————————————————————
Private mstrLike As String '输入匹配方式
Private mblnPay As Boolean '中药是否输入付数
Private mblnTime As Boolean '变价是否输入付数
Private mlngPreRow As Long '记录当前行,当仅改变列时
'————————————————————————————————————————————————————————————————————
'数据对象
Private mrsInfo As New ADODB.Recordset '病人信息
Private mrsMedAudit As ADODB.Recordset  '病人已审批的费用项目
Private mrsUnit As ADODB.Recordset '可选择的执行科室
Private mrsClass As ADODB.Recordset '根据参数读取的当前可用的收费类别
Private mrsWork As New ADODB.Recordset '当天上班的药房
Private mblnWork As Boolean '当前是否有正在上班的药房
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
Private mintSuccess As Integer
Private Const STR_HEAD = "行,450,4;项目,2175,1;商品名,930,1;规格,900,1;单位,520,4;付数,520,1;数次,570,1;单价,795,7;应收金额,945,7;实收金额,945,7;商品条码,1450,4;内部条码,1450,4;类型,520,1"
Public Function zlBillEdit(ByVal frmMain As Form, _
    ByVal bytInState As Byte, ByVal lngModule As Long, ByVal strPrivs As String, _
    Optional int记录性质 As Integer = 2, Optional ByVal strInNO As String, _
    Optional int病人来源 As Integer = 2, Optional lng病人ID As Long, Optional lng主页ID As Long, _
    Optional ByVal lng开单科室ID As Long, Optional ByVal lng病人科室ID As Long, _
    Optional ByVal lng开嘱科室ID As Long, Optional ByVal str开嘱医生 As String, _
    Optional ByVal bln费用登记 As Boolean, Optional ByVal str单据登记时间 As String, _
    Optional ByVal lng医嘱ID As Long, Optional ByVal lng发送号 As Long, _
    Optional strOriginalNO As String, _
    Optional strFeeTab As Long, Optional blnDelete As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:单据查阅或编辑入库
    '入参:bytInState:0-执行,1-查阅,2-调整(不支持),3-删费
    '       strInNO:所操作的单据号(执行时为修改)
    '       int记录性质 :1-收费(划价),2-记帐(门/住)
    '       int病人来源:1-门诊病人,2-住院病人
    '       lng医嘱ID -新增费用时用
    '       lng发送号-新增费用时用
    '       lng病人ID-病人ID
    '       strFeeTab:
    '       bln费用登记:仅登记,不计实收金额
    '       strOriginalNO -补充主费用时,医嘱发送中的单据号
    '       str单据登记时间:操作单据内容的登记时间
    '       blnDelete-是否处理退费单据(查阅)
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2010-12-13 17:09:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytInState = bytInState: mstrInNO = strInNO: mint病人来源 = int病人来源: mlng医嘱ID = lng医嘱ID
    mlng发送号 = lng发送号: mlng病人ID = lng病人ID: mlng主页ID = lng主页ID: mintSuccess = 0
    mstrFeeTab = strFeeTab: mbln费用登记 = bln费用登记: mlng开嘱科室ID = lng开嘱科室ID
    mlng病人科室id = lng病人科室ID
    mlng开单科室ID = lng开单科室ID:    mstrPrivsOpt = GetInsidePrivs(Enum_Inside_Program.p住院记帐操作)
    mstrPrivs = strPrivs: mint记录性质 = int记录性质
    mstr开嘱医生 = str开嘱医生: mstrOriginalNO = strOriginalNO: mblnDelete = blnDelete
    mblnUnload = False
    Me.Show 1, frmMain
    zlBillEdit = mintSuccess > 0
 End Function
Private Sub Bill_cboKeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Bill.cboStyle = DropOlnyDown Then Exit Sub
    Exit Sub
End Sub

Private Sub cbo开单科室_Validate(Cancel As Boolean)
    If cbo开单科室.Text <> "" And cbo开单科室.ListIndex < 0 Then
        mobjBill.开单部门ID = 0
        cbo开单科室.Text = ""
    End If
End Sub
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
            If CheckItemHaveSub(lngMainRow) Then
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

Private Sub ShowStock(str材料 As String, dbl库存 As Double)
'功能：显示药品或卫材的库存
    If InStr(1, mstrPrivs, "显示库存") > 0 Then
        sta.Panels(Pan.C2提示信息).Text = "[" & str材料 & "]可用库存:" & dbl库存
    Else
        sta.Panels(Pan.C2提示信息).Text = "[" & str材料 & "]" & IIF(dbl库存 > 0, "有", "无") & "库存."
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
    Dim blnCancel As Boolean
    If SelectItem(False) = False Then
         mblnSelect = False: Exit Sub
    End If
    mblnSelect = True
    Bill.Text = mobjDetail.ID
    Call bill_KeyDown(13, 0, blnCancel)
    Bill.SetFocus
    mblnSelect = False
    If Not blnCancel Then
        Bill.Text = "": Bill.TxtVisible = False
        Call zlCommFun.PressKey(13)
    End If
End Sub

Private Sub bill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    '功能：处理单据输入
    Dim dblStock As Double, strScope As String, i As Long
    Dim dblPreTime As Double, dblPreMoney As Double
    Dim blnSkip As Boolean, curTotal As Currency
    Dim blnStock As Boolean, lngDoUnit As Long, str摘要 As String
    Dim lng项目id As Long, str特准项目 As String, str类别 As String
    Dim blnInput As Boolean, cur余额 As Currency, lng病人科室ID As Long, int险类 As Integer, lngOld付数 As Long
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
        
        Select Case Bill.TextMatrix(0, Bill.Col)
            Case "项目"
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
                
                    sta.Panels(2).Text = ""
                    sta.Panels(4).Text = ""
                    blnInput = True
                    If Not mblnSelect Then
                        If SelectItem(True) = False Then
                              Bill.Text = "": Bill.TxtVisible = False
                              Bill.SetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    mblnSelect = False '立即清除该标志
                    Bill.TxtVisible = False '(不加不行)
                    
                    '医保费用项目是否审批检查
                    If mint病人来源 = 2 And mint记录性质 = 2 And Not IsNull(mrsInfo!险类) Then
                        If mobjDetail.要求审批 And Not mrsMedAudit Is Nothing Then
                            mrsMedAudit.Filter = "项目ID=" & mobjDetail.ID
                            If mrsMedAudit.RecordCount = 0 Then
                                MsgBox "当前病人未被批准使用该项目！", vbInformation, gstrSysName
                                Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                            ElseIf Not IsNull(mrsMedAudit!可用数量) Then
                                If mrsMedAudit!可用数量 <= 0 Then
                                    MsgBox "当前病人使用[" & mobjDetail.名称 & "]已达到批准的使用限量" & FormatEx(mrsMedAudit!使用限量, 5) & "。", vbInformation, gstrSysName
                                    Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                                End If
                            End If
                        End If
                    End If
                    
                    '主项适用病人病区科室
                    If mint病人来源 = 2 And mint记录性质 = 2 Then
                        If Not CheckFeeItemLimitDept(mobjDetail.ID) Then
                            MsgBox "该卫生材料对当前病人病区和科室不适用！", vbInformation, gstrSysName
                            Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                        End If
                    End If
           
                    
                    '病人科室ID
                    lng病人科室ID = mobjBill.科室ID
                    If lng病人科室ID = 0 And cbo开单科室.ListIndex <> -1 Then lng病人科室ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
                    
                    Call ShowStock(mobjDetail.名称, mobjDetail.库存)
                    
                    '保险支付项目对应检查
                    If Not IsNull(mrsInfo!险类) Then
                        
                        If zlCheck定价零价格对码(mobjDetail.ID, Not mobjDetail.变价) Then
                            '问题:27286
                        Else
                            If Not ItemExistInsure(mrsInfo!病人ID, mobjDetail.ID, mrsInfo!险类) Then
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
                    End If
                    
                    '输入摘要(取已有的行以便修改)
                    If mobjBill.Details.Count >= Bill.Row Then
                        If mobjBill.Details(Bill.Row).Detail.ID = mobjDetail.ID Then
                            str摘要 = mobjBill.Details(Bill.Row).摘要
                        End If
                    End If
                    
                    '加入或修改该收费细目行
                    Call SetDetail(mobjDetail, Bill.Row, lngDoUnit)
                    '59051
                    '输入摘要(根据新输入的行更改摘要)
                    If mobjBill.Details(Bill.Row).Detail.补充摘要 Then
                        If frmInputBox.InputBox(Me, "摘要", "请输入""" & mobjBill.Details(Bill.Row).Detail.名称 & """的摘要信息:", 200, 3, True, False, str摘要) Then
                            mobjBill.Details(Bill.Row).摘要 = str摘要
                        End If
                    ElseIf mint病人来源 = 2 And Not IsNull(mrsInfo!险类) Then
                        str摘要 = gclsInsure.GetItemInfo(mrsInfo!险类, mrsInfo!病人ID, mobjBill.Details(Bill.Row).收费细目ID, str摘要, 2)
                        mobjBill.Details(Bill.Row).摘要 = str摘要
                    End If
                    Call CalcMoneys(Bill.Row)
                    
                    'Calcmoney中医保可能返回摘要
                    If mobjBill.Details(Bill.Row).摘要 <> "" Then str摘要 = mobjBill.Details(Bill.Row).摘要
                    
                    '记帐分类报警(在已经算出该行费用但未显示前)
                    If mint记录性质 = 2 And mrsWarn.State = 1 And mobjBill.Details.Count = Bill.Row Then
                        curTotal = GetBillTotal(mobjBill)
                        '刘兴洪:30504: and mbln费用登记=False
                        If curTotal > 0 And mbln费用登记 = False Then
                            cur余额 = Val(txt实收.Tag)
                            If gbln报警包含划价费用 Then cur余额 = cur余额 - GetPriceMoneyTotal(mrsInfo!病人ID, mint病人来源)
                            mbytWarn = BillingWarn(Me, mstrPrivs, mrsWarn, mrsInfo!姓名, cur余额, mrsInfo!当日额 - mcurModiMoney, curTotal, _
                                Nvl(mrsInfo!担保额, 0), mobjBill.Details(Bill.Row).收费类别, mobjBill.Details(Bill.Row).Detail.类别名称, mstrWarn, mintWarn)
                            If mbytWarn = 2 Or mbytWarn = 3 Then
                                mobjBill.Details.Remove Bill.Row '删除刚刚想要加入的费用行
                                Bill.Text = "": Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                    If mint记录性质 = 2 Then
                        If Not IsNull(mrsInfo!险类) And mobjBill.Details(Bill.Row).数次 <> 0 And MCPAR.实时监控 Then
                            If gclsInsure.CheckItem(mrsInfo!险类, 1, 0, MakeDetailRecord(mobjBill, NeedName(cbo开单人.Text), NeedName(cbo开单科室.Text), Bill.Row)) = False Then
                                mobjBill.Details.Remove Bill.Row '删除刚刚想要加入的费用行
                                Bill.Text = "": Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                    Call ShowDetails(Bill.Row)
                    Call ShowMoney
                    
                    '费用类型检查
                    Call Check费用类型(Bill.Row)
                    
                    Bill.Text = "": Bill.SetFocus
                End If
                
                If mobjBill.Details.Count >= Bill.Row Then
                    With mobjBill.Details(Bill.Row)
                        Bill.ColData(BillCol.数次) = 4 '数次
                        Bill.ColData(BillCol.单价) = 5 '单价
                        '检查卫生材料的灭菌效期,在确定执行科室之后
                        Call CheckValidity(.收费细目ID, mlng执行库房ID, .数次, False, .Detail.批次)     '已确认输入,仅能提醒
                        '备用的卫生材料,不能设置成重属项目
                    End With
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
                    '最大金额检查
                    If gcurMaxMoney > 0 Then
                        If CSng(Bill.Text) * mobjBill.Details(Bill.Row).数次 * Bill.TextMatrix(Bill.Row, BillCol.单价) > gcurMaxMoney Then
                            If MsgBox("当前金额超过了" & gcurMaxMoney & ",你确定要继续吗?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                                Bill.Text = mobjBill.Details(Bill.Row).付数: Cancel = True: Exit Sub
                            End If
                        End If
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
                    '最大金额检查
                    If gcurMaxMoney > 0 Then
                        If CSng(Bill.Text) * mobjBill.Details(Bill.Row).付数 * Bill.TextMatrix(Bill.Row, BillCol.单价) > gcurMaxMoney Then
                            If MsgBox("当前金额超过了" & gcurMaxMoney & ",你确定要继续吗?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                                Bill.Text = mobjBill.Details(Bill.Row).数次: Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                    Bill.Text = FormatEx(Bill.Text, 5)
                    
                    '负数合法性检查
                    If CSng(Bill.Text) * mobjBill.Details(Bill.Row).付数 < 0 Then
                        '权限
                        If InStr(mstrPrivs, "诊疗负数费用") = 0 Then
                            MsgBox "你没有权限输入负数！", vbInformation, gstrSysName
                            Bill.Text = mobjBill.Details(Bill.Row).数次: Cancel = True: Exit Sub
                        Else
                            If mobjBill.Details(Bill.Row).Detail.分批 Then
                                MsgBox "分批卫生材料不允许输入负数。", vbInformation, gstrSysName
                                Bill.Text = mobjBill.Details(Bill.Row).数次: Cancel = True: Exit Sub
                            End If
                            If mrsInfo.State = 1 And mint记录性质 = 2 Then
                                If Not IsNull(mrsInfo!险类) Then
                                    If Not MCPAR.负数记帐 Then
                                        MsgBox "本地医保不支持对医保病人进行负数记帐！", vbInformation, gstrSysName
                                        Bill.Text = mobjBill.Details(Bill.Row).数次: Cancel = True: Exit Sub
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                    '药品库存检查
                    With mobjBill.Details(Bill.Row)
                        If .Detail.分批 Or .Detail.变价 Then
                            '分批或时价药品不足禁止输入
                            If .付数 * CSng(Bill.Text) > .Detail.库存 Then
                                MsgBox """" & .Detail.名称 & """为分批或时价卫生材料,当前可用库存不足输入数量！", vbInformation, gstrSysName
                                Bill.Text = .数次: Cancel = True: Exit Sub
                            End If
                        Else
                            Set colStock = mcolStock2
                            If .付数 * CSng(Bill.Text) > .Detail.库存 Then
                                Call MsgBox("""" & .Detail.名称 & """的当前可用库存不足输入数量,不能继续!", vbInformation, gstrSysName)
                                    Bill.Text = .数次: Cancel = True: Exit Sub
                             End If
                        End If
                    
                        dblPreTime = .数次
                        .数次 = Bill.Text
                        '固有从属不能更改数次(主项目数次改变,固有从属的数次也变)
                        If .从属父号 <> 0 And .Detail.固有从属 <> 0 Then
                            sta.Panels(2) = "该项目是固有从属项目,其数次不能够更改。"
                            .数次 = dblPreTime: Bill.Text = dblPreTime
                            Exit Sub
                        End If
                    End With
                
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
                        
                        '刘兴洪:2010-07-01 10:23:11:30504:and mbln费用登记=False
                        If curTotal > 0 And mbln费用登记 = False Then
                            cur余额 = Val(txt实收.Tag)
                            If gbln报警包含划价费用 Then cur余额 = cur余额 - GetPriceMoneyTotal(mrsInfo!病人ID, mint病人来源)
                            mbytWarn = BillingWarn(Me, mstrPrivs, mrsWarn, mrsInfo!姓名, cur余额, mrsInfo!当日额 - mcurModiMoney, curTotal, _
                                Nvl(mrsInfo!担保额, 0), mobjBill.Details(Bill.Row).收费类别, mobjBill.Details(Bill.Row).Detail.类别名称, mstrWarn, mintWarn)
                            If mbytWarn = 2 Or mbytWarn = 3 Then
                                mobjBill.Details(Bill.Row).数次 = dblPreTime
                                Bill.Text = ""
                                Call CalcMoneys(Bill.Row)
                                Cancel = True: Bill.TxtVisible = False: Exit Sub
                            End If
                        End If
                    End If
                    
                                     
                    If mint记录性质 = 2 Then
                        If Not IsNull(mrsInfo!险类) And mobjBill.Details(Bill.Row).数次 <> 0 And MCPAR.实时监控 Then
                            If gclsInsure.CheckItem(mrsInfo!险类, 1, 0, MakeDetailRecord(mobjBill, NeedName(cbo开单人.Text), NeedName(cbo开单科室.Text), Bill.Row)) = False Then
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
                        If mobjBill.Details(i).从属父号 = Bill.Row Then
                            '28136
                            '如果是输入的负数,需要将下级中的负数集中更新成负数
                            With mobjBill.Details(i)
                                If .Detail.固有从属 = 0 Then  '非固有从属
                                    If Abs(.数次) <> Abs(.Detail.从项数次) Then GoTo NotCalc:
                                    .数次 = IIF(Val(Bill.Text) < 0, -1, 1) * .Detail.从项数次
                                ElseIf .Detail.固有从属 = 1 Then '固定的固有从属
                                    .数次 = IIF(Val(Bill.Text) < 0, -1, 1) * IIF(.Detail.从项数次 = 0, 1, .Detail.从项数次)
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
                    
                    Call ShowMoney

                 ElseIf mobjBill.Details.Count >= Bill.Row Then
                    If Val(Bill.TextMatrix(Bill.Row, Bill.Col)) = 0 Then
                        If MsgBox("数量输入为零，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Cancel = True: Exit Sub
                        End If
                    End If
               End If
                If CheckItemHaveSub(Bill.Row) Then
                    KeyCode = 0
                    Call LocateMainItemNextRow(Bill.Row)
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
                    '最大金额检查
                    If gcurMaxMoney > 0 Then
                        If CSng(Bill.Text) * mobjBill.Details(Bill.Row).付数 * mobjBill.Details(Bill.Row).数次 > gcurMaxMoney Then
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
                        If mint记录性质 = 2 And mrsWarn.State = 1 Then
                            curTotal = GetBillTotal(mobjBill)
                            '30504:and mbln费用登记=False
                            If curTotal > 0 And mbln费用登记 = False Then
                                cur余额 = Val(txt实收.Tag)
                                If gbln报警包含划价费用 Then cur余额 = cur余额 - GetPriceMoneyTotal(mrsInfo!病人ID, mint病人来源)
                                mbytWarn = BillingWarn(Me, mstrPrivs, mrsWarn, mrsInfo!姓名, cur余额, mrsInfo!当日额 - mcurModiMoney, curTotal, _
                                    Nvl(mrsInfo!担保额, 0), mobjBill.Details(Bill.Row).收费类别, mobjBill.Details(Bill.Row).Detail.类别名称, mstrWarn, mintWarn)
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
                            '材料库存检查:动态药房,分批或时价材料也要检查了
                            If .Detail.分批 Or .Detail.变价 Then '分批或时价药品库存不足禁止输入
                                If .付数 * .数次 > .Detail.库存 Then
                                    MsgBox "[" & .Detail.名称 & "]为分批或时价卫生材料,当前可用库存不足输入数量！", vbInformation, gstrSysName
                                    Cancel = True
                                End If
                            Else
                                Set colStock = mcolStock2
                                If .付数 * .数次 > .Detail.库存 Then
                                    MsgBox "[" & .Detail.名称 & "]的当前可用库存不足输入数量！", vbInformation, gstrSysName
                                    Cancel = True
                                End If
                            End If
                        
                            '检查卫生材料的灭菌效期,在确定执行科室之后
                            Call CheckValidity(.收费细目ID, mlng执行库房ID, .数次, False, .Detail.批次)     '已确认输入,仅能提醒
                            If CheckItemHaveSub(Bill.Row) Then
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
        Bill.Col = BillCol.项目
    Else
        Bill.Row = Bill.Rows - 1
        Bill.MsfObj.TopRow = Bill.Row
        Bill.Col = BillCol.项目
    End If
    '问题:27792
    If Not Me.ActiveControl Is Bill Then
        If Bill.Active And Bill.Visible Then Bill.SetFocus
    End If
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
    Dim strStock As String, i As Long
    
    If Not Bill.Active Then Exit Sub
    If Bill.ColData(Col) = BillColType.UnFocus Then Exit Sub
    If Not mblnEnterCell Then Exit Sub
    
    If mbytInState = 3 Then
        '针对列编辑性质设置颜色
        Exit Sub
    End If
    
     '--------------------------------------------------------------------------
    '1.行改变的相关数据处理和设置
    If mobjBill.Details.Count >= Bill.Row And mlngPreRow <> Row Then
        With mobjBill.Details(Bill.Row)
            '显示库存
            Call ShowStock(.Detail.名称, .Detail.库存)
            Bill.ColData(BillCol.项目) = BillColType.CommandButton
             '如果是从属项目的主项目或从项,则不允许更改类别和项目
            If CheckItemHaveSub(Row) Or .从属父号 > 0 Then
                Bill.ColData(BillCol.项目) = BillColType.Text_UnModify
            End If
            
            '如果是非调整状态
            If mbytInState <> 2 Then
                Bill.ColData(BillCol.付数) = 5
                '变价允许输入数次
                Bill.ColData(BillCol.数次) = 4
                Bill.ColData(BillCol.单价) = 5
            End If
        End With
    End If
   
    '如果点击未保存的行,则恢复列的性质
    If mobjBill.Details.Count < Bill.Row Then
        Bill.ColData(BillCol.项目) = BillColType.CommandButton  '项目列,当主从项时会被改变
    End If
    
    
    '-----------------------------------------------------------------
    '2.列改变的相关数据处理和显示设置
    Bill.RowData(Row) = Asc("4")
    Bill.TextLen = 0: Bill.TextMask = ""
    Select Case Bill.TextMatrix(0, Col)
        Case "执行科室"
            Call zlControl.CboSetWidth(Bill.CboHwnd, 2000)
        Case "付数"
            Bill.TextLen = 3
            Bill.TextMask = "0123456789" & Chr(8)
        Case "数次"
            Bill.TextLen = 8
            Bill.TextMask = "0123456789" & Chr(8)
            
            If mobjBill.Details.Count >= Bill.Row Then
                Bill.TextMask = "." & Bill.TextMask
                
                '可否输入负数
                If Not mobjBill.Details(Bill.Row).Detail.分批 Then
                    If InStr(mstrPrivs, "诊疗负数费用") > 0 Then Bill.TextMask = "-" & Bill.TextMask
                                    
                    If InStr(Bill.TextMask, "-") > 0 Then
                        If mrsInfo.State = 1 And mint记录性质 = 2 Then
                            If Not IsNull(mrsInfo!险类) Then
                                If Not MCPAR.负数记帐 Then
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
 

Private Sub cboBaby_Click()
    mobjBill.婴儿费 = cboBaby.ItemData(cboBaby.ListIndex)
End Sub

Private Sub cbo开单科室_Click()
    Dim i As Long, strDoctor As String
    Dim rsReturn As ADODB.Recordset
    Dim intIndex As Integer
    If mbytInState <> 0 Then Exit Sub
    mrsAll开单科室.Filter = ""
    If cbo开单科室.ItemData(cbo开单科室.ListIndex) = 0 And cbo开单科室.Text Like "其他科室*" Then
        If zlDatabase.zlShowListSelect(Me, glngSys, 1150, cbo开单科室, mrsAll开单科室, True, "", "缺省,优先级", rsReturn) = False Then
            mobjBill.开单部门ID = 0
            Exit Sub
        End If
        If rsReturn Is Nothing Then Exit Sub
        If rsReturn.State <> 1 Then Exit Sub
        If rsReturn.RecordCount = 0 Then Exit Sub
        rsReturn.MoveFirst
        If zlControl.CboLocate(cbo开单科室, Val(rsReturn!ID), True) = False Then
            cbo开单科室.RemoveItem cbo开单科室.ListCount - 1
            cbo开单科室.AddItem IIF(zlIsShowDeptCode, rsReturn!编码 & "-", "") & rsReturn!名称
            cbo开单科室.ItemData(cbo开单科室.ListCount - 1) = Val(Nvl(rsReturn!ID))
            intIndex = cbo开单科室.NewIndex
            cbo开单科室.AddItem "其他科室…"
            cbo开单科室.ItemData(cbo开单科室.ListCount - 1) = 0
            cbo开单科室.ListIndex = intIndex
        End If
        Exit Sub
    End If
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
                Bill.RowData(Bill.Rows - 1) = 0
            End If
        End If
    End If
End Sub

Private Sub cbo开单人_KeyDown(KeyCode As Integer, Shift As Integer)
    If cbo开单人.Locked Then Exit Sub
    mblnDrop = False
    If KeyCode = 13 Then mblnDrop = SendMessage(cbo开单人.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 1
End Sub


Private Sub cboBaby_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo开单人_Validate(Cancel As Boolean)
    If cbo开单人.Text <> "" Then
        Call GetCboIndex(cbo开单人, NeedName(cbo开单人.Text))
        If cbo开单人.ListIndex = -1 Then cbo开单人.Text = ""
    End If
    If cbo开单人.Text = "" Then Call cbo开单人_KeyPress(vbKeyReturn)
End Sub

Private Sub cbo执行部门_Click()
    If mblnCboNotClick = True Then Exit Sub
    mlng执行库房ID = cbo执行部门.ItemData(cbo执行部门.ListIndex)
    mlng虚拟库房ID = Set虚拟库房ID(mlng执行库房ID)
    If mlng虚拟库房ID = 0 Then
        MsgBox "注意:" & vbCrLf & "    执行库房与虚拟库房未设置对应关系,请与管理员联系!", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
    End If
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
    If chk加班.value = Unchecked And blnAdd Then
        If MsgBox("当前处于加班时间范围内,要取消加班加价吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            chk加班.value = Checked
        End If
    End If
    If chk加班.value = Checked And Not blnAdd Then
        If MsgBox("当前不处于加班时间范围内,要执行加班加价吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            chk加班.value = Unchecked
        End If
    End If
    mobjBill.加班标志 = IIF(chk加班.value = Checked, 1, 0)
    
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
    Dim strSQL As String, i As Long, j As Long
    Dim strItems As String, str部门 As String
    Dim str单位 As String, dbl数量 As Double
    Dim strValues(0 To 10) As String, intR As Long
    Dim strSubTable As String, dbl数次合计 As Double, dbl已结数量 As Double
    
    '问题:26951
    If InStr(1, mstrPrivs, ";负数记帐不检查发生项目;") > 0 Then
        '对于负数冲销时不检查本次住院发生的项目数量,有此权限,允许录入病人未曾发生的费用项目进行冲销,否则检查本次住院发生的项目数量才能冲销
        CheckNegative = True: Exit Function
    End If
    
    CheckNegative = True
    If mobjBill.病人ID = 0 Then Exit Function
    
    strItems = ""
    strSubTable = ""
    intR = 0
    For i = 1 To mobjBill.Details.Count
        With mobjBill.Details(i)
            If .数次 < 0 And mlng执行库房ID <> 0 Then
                If Len(strItems) > 2000 Then
                    If intR <= 10 Then
                        strValues(intR) = Mid(strItems, 2)
                        '"           to_number(substr(Column_Value, instr(Column_Value,'_')+1)) As 批次, "
                        strSubTable = strSubTable & " Union ALL " & _
                        " Select  Column_Value As 收费细目ID" & _
                        " From Table(Cast(f_num2list([" & intR + 4 & "]) As ZLTOOLS.t_numlist))"
                    Else
                        strSubTable = strSubTable & " Union ALL " & _
                        " Select  Column_Value As 收费细目ID " & _
                        " From Table(Cast(f_num2list('" & Mid(strItems, 2) & "') As ZLTOOLS.t_numlist))"
                    End If
                    strItems = "": intR = intR + 1
                End If
                'strItems = strItems & "," & .收费细目ID & "_" & .Detail.批次 & ""
                strItems = strItems & "," & .收费细目ID
            End If
        End With
    Next
    If strItems <> "" Then
        If intR <= 10 Then
            strValues(intR) = Mid(strItems, 2)
            strSubTable = strSubTable & " Union ALL " & _
            " Select  Column_Value As 收费细目ID" & _
            " From Table(Cast(f_num2list([" & intR + 4 & "]) As ZLTOOLS.t_numlist))"
        Else
            strSubTable = strSubTable & " Union ALL " & _
            " Select  Column_Value As 收费细目ID" & _
            " From Table(Cast(f_num2list('" & Mid(strItems, 2) & "') As ZLTOOLS.t_numlist))"
        End If
    End If
    
    If strSubTable = "" Then Exit Function
    strSubTable = Mid(strSubTable, 11)
    
  
    strSQL = " " & _
    " with C1 as (" & strSubTable & ") " & vbCrLf & _
    " Select  /*+ RULE */  A.收费细目ID,A.执行部门ID, " & _
    "             Nvl(Sum(Decode(A.记录性质, 2, 1, 3, 1, 0) * Nvl(A.付数, 1) * A.数次), 0) As 数量, " & _
     "            Sum(Decode(nvL(Mod(M.记录状态 , 3),1),  0, 1, 1, 1, -1) * Decode(A.结帐id, Null, 0, 1) * Nvl(付数, 1) * 数次) As 结帐数量 " & _
     "     From " & mstrFeeTab & " A ,   病人结帐记录 M,C1   " & _
     "     Where  A.结帐id = M.ID(+)     And A.记帐费用=1 And A.价格父号 Is Null  And A.记录状态<>0 " & _
     "             And A.病人ID=[1] " & IIF(mint病人来源 = 2, "  And Nvl(A.主页ID,0)=[2]", "") & _
     "             and A.执行部门ID=[3]  " & _
    "               And A.收费细目ID=c1.收费细目ID  " & _
     "     Group By  A.收费细目ID,A.执行部门ID" & _
     "     Union ALL Select 收费细目ID,[3]+0 as 执行部门ID,0 as 数量,0 as 结帐数量 From C1 "
    
    'strSQL = _
    " with C1 as (" & strSubTable & ") " & vbCrLf & _
    " Select  /*+ RULE */  A.收费细目ID,A.执行部门ID,nvl(A.批次,0) as 批次,Sum(Nvl(A.付数,1)*A.数次) as 数量," & _
    "           Sum(decode(A.结帐ID,NULL,0,1)* Nvl(A.付数,1)*A.数次) as 结帐数量 " & _
    " From  " & mstrFeeTab & " A " & _
    " Where A.记录状态<>0 And A.记帐费用=1 and A.执行部门ID=[3] And A.价格父号 is NULL" & _
    "           And A.病人ID=[1] " & IIF(mint病人来源 = 2, "  And Nvl(A.主页ID,0)=[2]", "") & _
    "           And (A.收费细目ID+0,A.批次,0,0) in (select * From C1) " & _
    " Group by A.收费细目ID,A.执行部门ID,A.批次" & _
    " Union ALL Select 收费细目ID,[3] as 执行部门ID,批次,数量,结帐数量 From C1"
    
    strSQL = "" & _
    "   Select 收费细目ID,执行部门ID,0 as 批次,Sum(数量) as 数量,sum(结帐数量) as 结帐数量 " & _
    "   From (" & strSQL & ") " & _
    "   Group by 收费细目ID,执行部门ID"
    
    On Error GoTo errH
     Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjBill.病人ID, mobjBill.主页ID, mlng执行库房ID, strValues(0), strValues(1), strValues(2), strValues(3), strValues(4), strValues(5), strValues(6), strValues(7), strValues(8), strValues(9), strValues(10))
    
    For i = 1 To mobjBill.Details.Count
        With mobjBill.Details(i)
            If .数次 < 0 And mlng执行库房ID <> 0 Then
                rsTmp.Filter = "收费细目ID=" & .收费细目ID & " And 执行部门ID=" & mlng执行库房ID & " And 批次=" & .Detail.批次
                If Not rsTmp.EOF Then
                    str单位 = .Detail.计算单位
                    dbl数量 = Nvl(rsTmp!数量, 0)
                    dbl数次合计 = Abs(.数次) * .付数
                    dbl已结数量 = Val(Nvl(rsTmp!结帐数量))
                    '可能存在两条相同的记录
                    '问题:29412
                    For j = i + 1 To mobjBill.Details.Count
                         If .收费细目ID = mobjBill.Details(j).收费细目ID And .Detail.批次 And mobjBill.Details(j).Detail.批次 _
                            And mobjBill.Details(j).数次 < 0 Then
                            dbl数次合计 = dbl数次合计 + Abs(.数次) * .付数
                         End If
                    Next
                    '问题:32106
                    If dbl数次合计 > dbl数量 - dbl已结数量 Then
                        Select Case gbytBillOpt '对已结帐的记帐单据的操作权限:0-允许,1-提醒,2-禁止。
                        Case 0  '允许
                            If dbl数次合计 > dbl数量 Then
                                str部门 = Trim(cbo执行部门.Text)
                                MsgBox "第 " & i & " 行[" & .Detail.名称 & "]退回" & str部门 & "的数量 " & FormatEx(dbl数次合计, 5) & str单位 & _
                                    " 多于已计费数量 " & FormatEx(dbl数量, 5) & str单位 & "。", vbInformation, gstrSysName
                                CheckNegative = False: Exit Function
                            End If
                        Case 1   '提醒
                            str部门 = Trim(cbo执行部门.Text)
                            If dbl数次合计 > dbl数量 Then
                                    MsgBox "第 " & i & " 行[" & .Detail.名称 & "]退回" & str部门 & "的数量 " & FormatEx(dbl数次合计, 5) & str单位 & _
                                        " 多于已计费数量 " & FormatEx(dbl数量, 5) & str单位 & "。", vbInformation, gstrSysName
                                    CheckNegative = False: Exit Function
                            End If
                            
                            If MsgBox("第 " & i & " 行[" & .Detail.名称 & "]退回" & str部门 & "的数量 " & FormatEx(dbl数次合计, 5) & str单位 & _
                                " 中包含了已结部分(未结:" & FormatEx(dbl数量 - dbl已结数量, 5) & str单位 & "; 已结:" & FormatEx(dbl已结数量, 5) & str单位 & ") 。" & vbCrLf & _
                                " 是否继续?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                                CheckNegative = False: Exit Function
                            End If
                        Case 2   '禁止
                            str部门 = Trim(cbo执行部门.Text)
                            MsgBox "第 " & i & " 行[" & .Detail.名称 & "]退回" & str部门 & "的数量 " & FormatEx(dbl数次合计, 5) & str单位 & _
                                " 多于已计费数量 " & FormatEx(dbl数量 - dbl已结数量, 5) & str单位 & "。", vbInformation, gstrSysName
                                CheckNegative = False: Exit Function
                        End Select
                    End If
                Else
                    MsgBox "第 " & i & " 行[" & .Detail.名称 & "]可销帐数量为零，不允许冲销。", vbInformation, gstrSysName
                End If
            End If
        End With
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    CheckNegative = False
    Call SaveErrLog
End Function

Private Sub cmdOK_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strInfo As String, strSQL As String, strTmp As String
    Dim i As Long, j As Long, lng结帐ID As Long
    Dim curTotal As Currency, intInsure As Integer
    Dim dblTotal As Double, cur余额 As Currency, dbl数次 As Double
    Dim cur当日额 As Currency, colStock As Collection
    Dim blnTrans As Boolean, strNos As String
    
    If mbytInState = 3 Then
        If mint记录性质 <> 1 And (False Or mlng医嘱ID <> 0) Then '划价是全部删除
            For i = 1 To Bill.Rows - 1
                If Bill.RowData(i) > 0 Then
                    strSQL = strSQL & "," & Bill.RowData(i)
                End If
            Next
            If strSQL = "" Then
                MsgBox "请至少选择一行要删除的费用！", vbInformation, gstrSysName
                Bill.SetFocus: Exit Sub
            End If
            
            '所有行选择处理
            strSQL = Mid(strSQL, 2)
            i = GetBillRows(mstrInNO, mint记录性质, mint病人来源)
            If UBound(Split(strSQL, ",")) + 1 = i Then strSQL = ""
        Else
            '因为要处理为全退，如果结帐后不允许销帐，部份结帐后就要检查
            j = 0
            For i = 1 To Bill.Rows - 1
                If Bill.RowData(i) > 0 Then j = j + 1
            Next
            i = GetBillRows(mstrInNO, mint记录性质, mint病人来源)
            If j < i Then
                MsgBox "单据中的部份项目当前已不允许销帐(比如已执行或已结帐的项目)。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        '医保记帐作废上传(注意判断顺序)
        If mint病人来源 = 2 Then
            intInsure = BillExistInsure(mstrInNO) '判断是否医保病人记的帐
            If intInsure > 0 Then
                If gclsInsure.GetCapability(support记帐作废上传, mlng病人ID, intInsure) Then
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
        gcnOracle.BeginTrans: blnTrans = True
        
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                        
            '医保记帐作废上传
            If mint病人来源 = 2 And intInsure > 0 Then
                If gclsInsure.GetCapability(support记帐作废上传, mlng病人ID, intInsure) And Not gclsInsure.GetCapability(support记帐完成后上传, mlng病人ID, intInsure) Then
                    If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                        gcnOracle.RollbackTrans: Exit Sub
                    End If
                End If
            End If
        
        gcnOracle.CommitTrans: blnTrans = False
        
        '医保记帐作废上传
        If mint病人来源 = 2 And intInsure > 0 Then
            If gclsInsure.GetCapability(support记帐作废上传, mlng病人ID, intInsure) And gclsInsure.GetCapability(support记帐完成后上传, mlng病人ID, intInsure) Then
                If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                    MsgBox "单据""" & mstrInNO & """的删费数据向医保传送失败，该单据已删除。", vbInformation, gstrSysName
                End If
            End If
        End If
        
        On Error GoTo 0
        mintSuccess = 1
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
        If cbo执行部门.ListIndex < 0 Then
            MsgBox "单据没有指定执行科室！", vbInformation, gstrSysName
            Bill.SetFocus: Exit Sub
        End If
        
        If mobjBill.开单部门ID = 0 Then
            MsgBox "请确定开单科室！", vbInformation, gstrSysName
            cbo开单科室.SetFocus: Exit Sub
        End If
        
        If mobjBill.开单人 = "" Then
            MsgBox "请输入开单人！", vbInformation, gstrSysName
            cbo开单人.SetFocus: Exit Sub
        End If
        
        '护士类别:判断非法输入
        '发生时间检查
        If Not IsDate(txtDate.Text) Then
            MsgBox "请输入正确的费用日期！", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        
        '出院强制记帐权限检查
        If mint病人来源 = 2 Then
            If Not PatiCanBilling(mrsInfo!病人ID, Nvl(mrsInfo!主页ID, 0), mstrPrivs, p医嘱附费管理) Then Exit Sub
            If zlPatiIS病案已编目(mrsInfo!病人ID, Nvl(mrsInfo!主页ID, 0)) Then Exit Sub
            '49501:住院
            If zlIsAllowFeeChange(mrsInfo!病人ID, Val(Nvl(mrsInfo!主页ID))) = False Then Exit Sub
        End If
        
        '刘兴洪 问题:?? 日期:2010-01-07 10:37:09
        If zlCheck北京医保(Val(Nvl(mrsInfo!险类))) = False Then Exit Sub
        
        
        
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
         dbl数次 = 0
        For i = 1 To mobjBill.Details.Count
           '27467,52828
            If mobjBill.Details(i).数次 <> 0 And dbl数次 = 0 Then
                dbl数次 = mobjBill.Details(i).数次
            End If
            If mobjBill.Details(i).收费细目ID = 0 Then
                MsgBox "单据中第 " & i & " 行没有正确输入数据,请修正或删除该行！", vbInformation, gstrSysName
                Bill.SetFocus: Exit Sub
             '8407
            End If
        Next
        '27467,52828
        If mbytInState = 0 And Round(dbl数次, 7) = 0 Then
            MsgBox "单据中至少要有一条不为零的数次,请检查！", vbInformation, gstrSysName
            Bill.SetFocus: Exit Sub
        End If
        
        '费用类型检查
        If Not Check费用类型 Then Exit Sub
                
        '记帐分类报警
        If mint记录性质 = 2 And mrsWarn.State = 1 And mstrWarn <> "-" Then
            '单据费用
            curTotal = CalcGridToTal
            If curTotal > 0 Then
                '刷新病人费用状况
                Set rsTmp = GetMoneyInfo(mrsInfo!病人ID, IIF(mint病人来源 = 1, 0, mlng主页ID), mcurModiMoney)
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
                cur余额 = Val(txt实收.Tag)
                If gbln报警包含划价费用 Then cur余额 = cur余额 - GetPriceMoneyTotal(mrsInfo!病人ID, mint病人来源)
                
                If mbln费用登记 = False Then    '30504
                    For i = 1 To mobjBill.Details.Count
                        mbytWarn = BillingWarn(Me, mstrPrivs, mrsWarn, mrsInfo!姓名, cur余额, cur当日额 - mcurModiMoney, curTotal, IIF(IsNull(mrsInfo!担保额), 0, mrsInfo!担保额), mobjBill.Details(i).收费类别, mobjBill.Details(i).Detail.类别名称, mstrWarn, mintWarn)
                        If mbytWarn = 2 Or mbytWarn = 3 Then Exit Sub
                    Next
                End If
            End If
        End If
        
        If mint记录性质 = 2 And Not IsNull(mrsInfo!险类) And MCPAR.实时监控 Then
            If gclsInsure.CheckItem(mrsInfo!险类, 1, 2, MakeDetailRecord(mobjBill, NeedName(cbo开单人.Text), NeedName(cbo开单科室.Text))) = False Then
                Exit Sub
            End If
        End If
        
              
        '检查分批或时价药品同一药房是否有重复输入
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
                If (.Detail.分批 Or .Detail.变价) Then
                    For j = 1 To mobjBill.Details.Count
                        If i <> j And .收费细目ID = mobjBill.Details(j).收费细目ID And .Detail.批次 = mobjBill.Details(j).Detail.批次 Then
                            MsgBox "第 " & j & " 行的分批或时价卫生材料""" & .Detail.名称 & """在同一个发料部门被重复输入，请合并！", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    Next
                End If
            End With
        Next
        
        '药品库存检查(仅不足禁止时或分批时价药品)
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
                Set colStock = mcolStock2
                If .Detail.分批 Or .Detail.变价 Then
                    dblTotal = GetDrugTotal(mobjBill, .收费细目ID, mlng执行库房ID, .Detail.批次)
                    If dblTotal > .Detail.库存 Then
                        MsgBox "第 " & i & " 行时价或分批卫生材料""" & .Detail.名称 & _
                            """的当前库存" & IIF(InStr(1, mstrPrivs, "显示库存") > 0, .Detail.库存, "") & "不足输入数量""" & dblTotal & """。", vbInformation, gstrSysName
                        Exit Sub
                    End If
                ElseIf colStock("_" & mlng执行库房ID) = 2 Then
                    dblTotal = GetDrugTotal(mobjBill, .收费细目ID, mlng执行库房ID, .Detail.批次)
                    If dblTotal > .Detail.库存 Then
                        MsgBox "第 " & i & " 行卫生材料""" & .Detail.名称 & _
                            """的当前库存" & IIF(InStr(1, mstrPrivs, "显示库存") > 0, .Detail.库存, "") & "不足输入数量""" & dblTotal & """,请修改或检查是否有多行输入。", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
            End With
        Next
        
        '检查卫生材料的灭菌效期
 
        mblnSendMateria = False
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
                dblTotal = GetDrugTotal(mobjBill, .收费细目ID, mlng执行库房ID, .Detail.批次)
                If Not CheckValidity(.收费细目ID, mlng执行库房ID, dblTotal, True, .Detail.批次) Then Exit Sub
            End With
        Next
        If InStr(mstrPrivs, ";药品发药;") = 0 Then mblnSendMateria = False
        
        If mblnSendMateria And mbytSendMateria = 2 Then
            If MsgBox("记帐完成后自动执行发药吗?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                mblnSendMateria = False
            End If
        End If
        
        '负数退费检查
        If mint记录性质 = 2 Then
            If Not CheckNegative Then Exit Sub
        End If
        
        '刷卡消费验卡
        If mint病人来源 = 1 And mint记录性质 = 2 And gbln消费验证 Then
            curTotal = CalcGridToTal
            If curTotal > 0 Then
                If Not zlDatabase.PatiIdentify(Me, glngSys, mobjBill.病人ID, curTotal) Then Exit Sub
            End If
        End If
        '74231,冉俊明,2014-7-21,项目开单后立即收费或记帐审核
        If gobjSquareCard Is Nothing Then
            If mint病人来源 = 1 And gbln开单后立即结算 Then
                If MsgBox("注意：" & vbCrLf & "      医疗卡部件（zl9CardSquare）未创建，在您开单后将不能进行收费或记帐审核，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
        End If
        
        If Not SaveBill(strNos) Then Exit Sub
        
        '74231,冉俊明,2014-7-21,项目开单后立即收费或记帐审核
        If mint病人来源 = 1 And gbln开单后立即结算 And strNos <> "" Then
            If Not gobjSquareCard Is Nothing Then
                Call gobjSquareCard.zlSquareAffirm(Me, p医嘱附费管理, mstrPrivs, mlng病人ID, , , mint记录性质, strNos)
            End If
        End If
        
        mintSuccess = mintSuccess + 1
        '刘兴洪:打印发药单:25490
        If mblnSendMateria Then
            If InStr(1, mstrPrivs, ";发药清单打印;") > 0 Then
                If MsgBox("单据""" & mobjBill.NO & """发药完成，要打印发药清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1133_2", Me, "单据号=" & mobjBill.NO, "登记时间=" & Format(mobjBill.登记时间, "yyyy-MM-dd HH:mm:ss"), 1)
                End If
            End If
        End If
        
        If mstrInNO <> "" Or mstrOriginalNO <> "" Then
            gblnOK = True: Unload Me: Exit Sub
        Else
            txtPreNO.Text = mobjBill.NO
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
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function MakeDetailRecord(ByRef objBill As ExpenseBill, ByVal str开单人 As String, ByVal str开单科室 As String, _
    Optional ByVal lngRow As Long) As ADODB.Recordset
    '功能：根据单据对象内容创建一个明细记录集信息(以售价单位)
    '字段：病人ID，主页ID，收费类别，收费细目ID，数量，单价，实收金额，开单人，开单科室
    '参数：intPage=指定的单据,lngRow=指定的行，不指定时包含所有单据的所有行
    Dim i As Integer, j As Integer
    Dim intB As Integer, intE As Integer, blnNew As Boolean
    Dim dbl单价 As Double, cur实收 As Currency
    Dim rsTmp As New ADODB.Recordset
    
    rsTmp.Fields.Append "病人ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "主页ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "收费类别", adVarChar, 2, adFldIsNullable
    rsTmp.Fields.Append "收费细目ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "数量", adDouble, , adFldIsNullable
    rsTmp.Fields.Append "单价", adDouble, , adFldIsNullable
    rsTmp.Fields.Append "实收金额", adCurrency, , adFldIsNullable
    '69788:李南春,2014-6-5,调整开单人字段大小，由20改为100
    rsTmp.Fields.Append "开单人", adVarChar, 100, adFldIsNullable
    rsTmp.Fields.Append "开单科室", adVarChar, 50, adFldIsNullable
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    
    If lngRow = 0 Then
        intB = 1
        intE = objBill.Details.Count
    Else
        intB = lngRow
        intE = lngRow
    End If
    
    For i = intB To intE
        dbl单价 = 0: cur实收 = 0
        With objBill.Details(i)
            If lngRow = 0 Then
                rsTmp.Filter = "收费细目ID=" & .收费细目ID
                blnNew = rsTmp.RecordCount = 0
            Else
                blnNew = True
            End If
                            
            If blnNew Then
                rsTmp.AddNew
                
                rsTmp!病人ID = objBill.病人ID
                rsTmp!主页ID = objBill.主页ID
                
                rsTmp!收费类别 = .收费类别
                rsTmp!收费细目ID = .收费细目ID
                
                
                For j = 1 To .InComes.Count
                    dbl单价 = dbl单价 + .InComes(j).标准单价
                    cur实收 = cur实收 + .InComes(j).实收金额
                Next
                rsTmp!数量 = IIF(.付数 = 0, 1, .付数) * .数次
                rsTmp!单价 = Format(dbl单价, gstrDecPrice)
                rsTmp!实收金额 = Format(cur实收, gstrDec)
                
                rsTmp!开单人 = str开单人
                rsTmp!开单科室 = str开单科室
            Else
                For j = 1 To .InComes.Count
                    dbl单价 = dbl单价 + .InComes(j).标准单价
                    cur实收 = cur实收 + .InComes(j).实收金额
                Next
                rsTmp!数量 = rsTmp!数量 + IIF(.付数 = 0, 1, .付数) * .数次
                rsTmp!单价 = Format((rsTmp!单价 + Format(dbl单价, gstrDecPrice)) / 2, gstrDecPrice)
                rsTmp!实收金额 = rsTmp!实收金额 + Format(cur实收, gstrDec)
            End If
            
            rsTmp.Update
        End With
    Next
    
    rsTmp.Filter = ""
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    Set MakeDetailRecord = rsTmp
End Function

Private Sub cmdOK_GotFocus()
    If mbytInState = 3 Then
        Bill.Row = 1: Bill.Col = Bill.Cols - 1
    End If
End Sub

Private Sub cmdSel_Click()
    Dim rsSel As ADODB.Recordset
    If frmStuffInSel.zlSelect(Me, 1250, mstrPrivs, mlng虚拟库房ID, rsSel) = False Then
        Bill.SetFocus
        Exit Sub
    End If
    Call LoadSelBillData(rsSel)
    Bill.SetFocus
End Sub

Private Sub LoadSelBillData(ByVal rsSel As ADODB.Recordset)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:批量导入
    '编制:刘兴洪
    '日期:2010-12-16 17:04:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnALLIgnore As Boolean '忽略所有重复项
    Dim bln重复 As Boolean, i As Long
    Dim objDetail As Detail, IntMsg As VbMsgBoxResult
    '病人科室或开单科室ID
    With rsSel
        blnALLIgnore = False
        If .RecordCount > 0 Then .MoveFirst
        
        Do While Not .EOF
            bln重复 = False
            For i = 1 To mobjBill.Details.Count
                If mobjBill.Details(i).Detail.ID = Val(Nvl(!收费项目ID)) _
                    And mobjBill.Details(i).Detail.批次 = Val(Nvl(!批次)) Then
                    If blnALLIgnore = False Then
                        IntMsg = MsgBox("注意:" & vbCrLf & "  在第" & i & " & 行中已经存在了卫生材料“" & Nvl(!材料名称) & "”,是否忽略此项?" & _
                                            "『是』表示仅忽略当前批次卫材!" & vbCrLf & _
                                            "『否』表示忽略所有已经在单据中存在的批次！" & vbCrLf & _
                                            "『取消』表示退出本次选择的卫材!", vbYesNoCancel + vbQuestion + vbDefaultButton3, gstrSysName)
                        If IntMsg = vbCancel Then Exit Sub
                        If IntMsg = vbNo Then
                            blnALLIgnore = True
                        End If
                    End If
                    bln重复 = True
                End If
            Next
            If bln重复 = False Then
                Set objDetail = GetInputDetail(Val(!收费项目ID))
                objDetail.批次 = Val(Nvl(!批次))
                objDetail.商品条码 = Trim(Nvl(!商品条码))
                objDetail.内部条码 = Trim(Nvl(!内部条码))
                objDetail.库存 = Val(Nvl(!可用库存))
                
                '增加数据
                Call SetDetail(objDetail, mobjBill.Details.Count + 1, mlng虚拟库房ID)
                mobjBill.Details(mobjBill.Details.Count).数次 = Val(Nvl(!数量))
                Call CalcMoneys(mobjBill.Details.Count)
            End If
            .MoveNext
        Loop
    End With
    Bill.ClearBill: Call SetColNum
    Bill.Rows = mobjBill.Details.Count + 1
    '针对列编辑性质设置颜色
    Bill.SetColColor BillCol.项目, &HE7CFBA
    Bill.SetColColor BillCol.数次, &HE7CFBA
    Bill.SetColColor BillCol.付数, &HE0E0E0
    Bill.SetColColor BillCol.单价, &HE0E0E0
    mobjBill.操作员编号 = UserInfo.编号
    mobjBill.操作员姓名 = UserInfo.姓名
    
    Call ShowDetails
    Call ShowMoney
    For i = 1 To mobjBill.Details.Count
        With mobjBill.Details(i)
            Bill.RowData(i) = Asc("4") '特殊处理
            
        End With
    Next
    Call SetColNum
    If Bill.Enabled Then Bill.SetFocus
    
End Sub

Private Sub Form_Activate()
    If mblnUnload Then
        Unload Me: Exit Sub
    End If
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
    
    mblnWarnCloseed = False
    glngFormW = 12000: glngFormH = 7710
    If Not InDesign Then
        glngOld = GetWindowLong(Me.Hwnd, GWL_WNDPROC)
        Call SetWindowLong(Me.Hwnd, GWL_WNDPROC, AddressOf Custom_WndMessage)
    End If
    Call RestoreWinState(Me, App.ProductName, mbytInState)
    
    gblnOK = False
    mblnEnterCell = True
    mintWarn = -1: mstrWarn = ""
    mstrFeeTab = IIF(mint病人来源 = 2, "住院费用记录", "门诊费用记录")
    Call InitLocPar
    
    '初始化单据数据
    Set mobjBill = New ExpenseBill
    If mbytInState = 0 Then
        If Not InitData Then
            mblnUnload = True: Exit Sub
        End If
    End If
    Call InitFace
    Call NewBill
    
    If mbytInState <> 0 Then
        If Not ReadBill(mstrInNO, mbytInState = 3) Then
            mblnUnload = True: Exit Sub
        End If
    Else
        '读取该单据的内容
        If mstrInNO <> "" Then '修改单据
            Set mobjBill = ImportStuffBill(mint病人来源, mstrInNO, mint记录性质, mlng虚拟库房ID)
            If mobjBill.NO = "" Then
                MsgBox "不能正确读取计费单据的内容！", vbInformation, gstrSysName
                mblnUnload = True: Exit Sub
            Else
                Bill.ClearBill: Call SetColNum
                Bill.Rows = mobjBill.Details.Count + 1
                
                '针对列编辑性质设置颜色
                Bill.SetColColor BillCol.项目, &HE7CFBA
                Bill.SetColColor BillCol.数次, &HE7CFBA
                Bill.SetColColor BillCol.付数, &HE0E0E0
                Bill.SetColColor BillCol.单价, &HE0E0E0
                
                cboNO.Text = mobjBill.NO
                               
                
                mobjBill.操作员编号 = UserInfo.编号
                mobjBill.操作员姓名 = UserInfo.姓名
                
                If mint记录性质 = 2 Then
                    mcurModiMoney = GetBillMoney(mobjBill.NO) '在读取病人前取
                End If
                
                '新单时读取病人,看单据时根据单据显示病人信息
                Call GetPatient(mlng病人ID, mlng主页ID)
                If mrsInfo.State = 0 Then
                    If Not mblnWarnCloseed Then
                        MsgBox "不能读取病人信息，可能是你不具有对该病人计费的权限。", vbInformation, gstrSysName
                    End If
                    Unload Me: Exit Sub
                End If
                
                Call FindCboIndex(cbo开单科室, mobjBill.开单部门ID, False)
                Call GetCboIndex(cbo开单人, mobjBill.开单人)
                Call zlControl.CboLocate(cboBaby, mobjBill.婴儿费, True)
                
                If gbln从项汇总折扣 Then CalcMoneys
                Call ShowDetails
                Call ShowMoney
                
                '调整库存:修改时加上将要退回的库存
                For i = 1 To mobjBill.Details.Count
                    With mobjBill.Details(i)
                        Bill.RowData(i) = Asc("4") '特殊处理
                        .Detail.库存 = .Detail.库存 + .付数 * .数次
                    End With
                Next
                Call SetColNum
            End If
        Else
            '新单时读取病人,看单据时根据单据显示病人信息
            Call GetPatient(mlng病人ID, mlng主页ID)
            If mrsInfo.State = 0 Then
                If Not mblnWarnCloseed Then
                    MsgBox "不能读取病人信息，可能是你不具有对该病人计费的权限。", vbInformation, gstrSysName
                End If
                mblnUnload = True: Exit Sub
            End If
            If Not IsNull(mrsInfo!险类) Then
                MCPAR.负数记帐 = gclsInsure.GetCapability(support负数记帐, mrsInfo!病人ID, mrsInfo!险类)
                MCPAR.记帐上传 = gclsInsure.GetCapability(support记帐上传, mrsInfo!病人ID, mrsInfo!险类)
                MCPAR.记帐完成后上传 = gclsInsure.GetCapability(support记帐完成后上传, mrsInfo!病人ID, mrsInfo!险类)
                MCPAR.实时监控 = gclsInsure.GetCapability(support实时监控, mrsInfo!病人ID, mrsInfo!险类)
                MCPAR.医生确定处方类型 = gclsInsure.GetCapability(support医生确定处方类型, mrsInfo!病人ID, mrsInfo!险类)
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
    mstrOriginalNO = ""
    
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

 
Private Sub picAppend_Resize()
    err = 0: On Error Resume Next
    With picAppend
        txt病人备注.Width = .ScaleWidth - txt病人备注.Left - 100
    End With
End Sub

Private Sub sta_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Not gbln简码匹配方式切换 Then Exit Sub
    If Panel.Bevel = sbrRaised And (Panel.Key = "PY" Or Panel.Key = "WB") Then
        '切换并保存简码匹配方式
        Panel.Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        If Panel.Key = "PY" Then
            sta.Panels("WB").Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        Else
            sta.Panels("PY").Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        End If
        Call zlDatabase.SetPara("简码方式", IIF(sta.Panels("PY").Bevel = sbrInset And sta.Panels("WB").Bevel = sbrInset, 2, IIF(sta.Panels("WB").Bevel = sbrInset, 1, 0)))
        
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
            .ColData(BillCol.项目) = BillColType.CommandButton  '项目列,当主从项时会被改变
            .ColData(BillCol.付数) = 5 '付数缺省跳过(=1),当类别为中药时,设为输入(4)(有值,一改全改)
            .ColData(BillCol.单价) = 5 '单价缺省跳过,当项目变价时,设为输入(4)
        End If
        '针对列编辑性质设置颜色
        .SetColColor BillCol.项目, &HE7CFBA
        .SetColColor BillCol.数次, &HE7CFBA
        .SetColColor BillCol.付数, &HE0E0E0
        .SetColColor BillCol.单价, &HE0E0E0
        
        .TextMatrix(Row, BillCol.行) = Row
        
        '特殊地方手动调用不执行
        If Row > 0 And .ColData(BillCol.项目) <> 5 And Me.Visible And Not mblnNewRow Then
            'Call zlCommFun.PressKey(13)
            
        End If
    End With
End Sub

Private Sub cbo开单科室_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long, lng医生ID As Long
     '刘兴洪 问题:27378 日期:2010-01-27 16:20:02
    
    If KeyAscii <> 13 Then Exit Sub
    
    If cbo开单科室.ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    If cbo开单科室.Locked Then Exit Sub
    If mrsAll开单科室 Is Nothing Then Exit Sub
    
    If zlSelectDept(Me, 1150, cbo开单科室, mrsAll开单科室, cbo开单科室.Text, True, , , True) = False Then
        mobjBill.开单部门ID = 0
        KeyAscii = 0: Exit Sub
    End If
    Exit Sub
End Sub

Private Sub cbo开单人_KeyPress(KeyAscii As Integer)
    Dim i As Long, intIdx As Integer
    Dim strText As String
    
    If KeyAscii = 13 Then
        If cbo开单人.Locked Then
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        strText = cbo开单人.Text
        If cbo开单人.ListIndex <> -1 Then
            '弹出列表时,又在文本框输入了内容
            If strText <> cbo开单人.List(cbo开单人.ListIndex) Then
                Call zlControl.CboSetIndex(cbo开单人.Hwnd, -1)
            Else
                zlCommFun.PressKey vbKeyTab: Exit Sub
            End If
        End If
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
            If ActiveControl Is cbo开单人 Then Call cbo开单人_KeyPress(vbKeyReturn)
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
            If Not gbln简码匹配方式切换 Then Exit Sub
            If sta.Panels("WB").Visible And sta.Panels("PY").Visible Then
                If sta.Panels("WB").Bevel = sbrRaised Then
                    Call sta_PanelClick(sta.Panels("WB"))
                Else
                    Call sta_PanelClick(sta.Panels("PY"))
                End If
            End If
        Case vbKeyF11
            If cmdSel.Visible And cmdSel.Enabled Then cmdSel_Click
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
    Dim strOperDoc As String
    
    On Error GoTo errH
    
    Set mcolStock2 = GetStockCheck(1)
    
    '开单科室
    strSQL = "Select 开嘱科室ID,开嘱医生 From 病人医嘱记录 Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng医嘱ID)
    If Not rsTmp.EOF Then
        mlng开嘱科室ID = Nvl(rsTmp!开嘱科室id, 0)
        mstr开嘱医生 = Nvl(rsTmp!开嘱医生)
    End If
    If mlng开单科室ID = 0 Or mstr开嘱医生 = "" Then
        MsgBox "没有发现源医嘱信息。", vbInformation, gstrSysName
        Exit Function
    End If
    
    strSQL = _
    "   Select A.ID, A.编码, A.名称, A.简码, 0 As 缺省, B.工作性质, D.优先级" & vbNewLine & _
    "   From 部门表 A, 部门性质说明 B," & vbNewLine & _
    "       (Select 部门id, Max(Decode(服务对象, 2, 1, 2)) As 优先级 From 部门性质说明 Where 服务对象 <> 0 Group By 部门id) D" & vbNewLine & _
    "   Where (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) And A.ID = B.部门id" & vbNewLine & _
    "       And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
    "       And B.部门id = D.部门id And (B.服务对象 IN(1,2,3) AND B.工作性质 IN('临床','手术') Or b.工作性质='产科')" & vbNewLine & _
    "Order By 优先级,编码"
    Set mrsAll开单科室 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    '70434:刘尔旋,2014-02-12,开单科室下拉列表新增主刀医生科室
    strOperDoc = Get医嘱附项内容(mlng医嘱ID, "主刀医生科室")
    
    If mbln费用登记 Then
        '就为当前选择的医技科室
        strSQL = "(Select ID,编码,名称,简码 From 部门表 Where ID=[1]"
    Else
        '就为当前选择的医技科室或开嘱科室
        strSQL = "(Select ID,编码,名称,简码 From 部门表 Where ID IN([1],[2])"
    End If
    
    If strOperDoc <> "" Then
        strSQL = strSQL & " Union " & _
                "Select ID,编码,名称,简码 From 部门表 Where 名称=[3]"
    End If
    strSQL = strSQL & ") Order By 编码"
    
    Set mrsDept = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng开单科室ID, mlng开嘱科室ID, strOperDoc)
    cbo执行部门.Clear
    mblnCboNotClick = True
    If Not mrsDept.EOF Then
        For i = 1 To mrsDept.RecordCount
            cbo开单科室.AddItem IIF(zlIsShowDeptCode, mrsDept!编码 & "-", "") & mrsDept!名称
            cbo开单科室.ItemData(cbo开单科室.ListCount - 1) = mrsDept!ID
            If mrsDept!ID = mlng开单科室ID Then
                cbo开单科室.ListIndex = cbo开单科室.NewIndex
                cbo执行部门.AddItem IIF(zlIsShowDeptCode, mrsDept!编码 & "-", "") & mrsDept!名称
                cbo执行部门.ItemData(cbo执行部门.NewIndex) = Val(Nvl(mrsDept!ID))
                cbo执行部门.ListIndex = cbo执行部门.NewIndex
                mlng执行库房ID = cbo执行部门.ItemData(cbo执行部门.NewIndex)
            End If
            mrsDept.MoveNext
        Next
        cbo开单科室.AddItem "其他科室…"
        cbo开单科室.ItemData(cbo开单科室.ListCount - 1) = 0
        If cbo开单科室.ListIndex = -1 Then cbo开单科室.ListIndex = 0
    Else
        MsgBox "不能确定开单科室，请先到部门管理中设置。", vbInformation, gstrSysName
        mblnCboNotClick = False
        Exit Function
    End If
    mblnCboNotClick = False
    mlng虚拟库房ID = Set虚拟库房ID(mlng执行库房ID)
    If mlng虚拟库房ID = 0 Then
        MsgBox "注意:" & vbCrLf & "    执行库房与虚拟库房未设置对应关系,请与管理员联系!", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    
    If cbo执行部门.ListCount = 0 Then
        MsgBox "不能确定执行部门，请先到部门管理中设置。", vbInformation, gstrSysName
        Exit Function
    End If
    mlng卫材类别ID = ExistIOClass(IIF(mint记录性质 = 1, 40, 41))
    If mlng卫材类别ID = 0 Then
        MsgBox "不能确定卫材单据的入出类别,请先到入出分类管理中设置！", vbInformation, gstrSysName
        Exit Function
    End If
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
        .Font.Size = 10.5
        .CboFont.Size = 11
        .TxtEditFont.Size = 11
        
        arrHead = Split(STR_HEAD, ";")
        .Cols = UBound(arrHead) + 1
        
        .MsfObj.FixedCols = 1
        .MsfObj.ScrollBars = flexScrollBarVertical
        .LocateCol = BillCol.项目
        .PrimaryCol = BillCol.项目
        .MsfObj.ColAlignmentFixed(0) = 4
        .TextMatrix(1, BillCol.行) = 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(0, i) = Split(arrHead(i), ",")(0)
            .ColWidth(i) = Split(arrHead(i), ",")(1)
            .ColAlignment(i) = Split(arrHead(i), ",")(2)
        Next
                
        If mbytInState = 0 Then
            .ColData(BillCol.行) = 5
            .ColData(BillCol.项目) = 1 '项目输入,按扭可选
            .ColData(BillCol.数次) = 4 '数/次输入
            '刘兴洪:27990 2010-02-23 12:04:37
            .ColData(BillCol.商品名) = 5 '规格跳过
            .ColData(BillCol.规格) = 5 '规格跳过
            .ColData(BillCol.单位) = 5 '单位跳过
            .ColData(BillCol.付数) = 5 '付数缺省跳过(=1),当类别为中药时,设为输入(4)(有值,一改全改)
            .ColData(BillCol.单价) = 5 '单价缺省跳过,当项目变价时,设为输入(4)
            .ColData(BillCol.应收金额) = 5 '应收金额跳过
            .ColData(BillCol.实收金额) = 5 '实收金额跳过
            .ColData(BillCol.类型) = 5 '类型缺省跳过
            .ColData(BillCol.商品条码) = 5
            .ColData(BillCol.内部条码) = 5
        End If
        .SetColColor BillCol.项目, &HE7CFBA
        .SetColColor BillCol.数次, &HE7CFBA
        .SetColColor BillCol.付数, &HE0E0E0
        .SetColColor BillCol.单价, &HE0E0E0
        
        ReDim marrColData(.Cols - 1)
        For i = 0 To .Cols - 1
            marrColData(i) = .ColData(i)
        Next
        If mbytInState = 3 Then .AllowAddRow = False
    End With
    Call RestoreFlexState(Bill, App.ProductName & "\" & Me.Name & mbytInState)
    If gbyt药品名称显示 <> 2 Then
        '0-显示通用名，1-显示商品名，2-同时显示通用名和商品名
        Bill.ColWidth(BillCol.商品名) = 0
    Else
        If Bill.ColWidth(BillCol.商品名) = 0 Then
             Bill.ColWidth(BillCol.商品名) = GetOrigColWidth(BillCol.商品名)
        End If
    End If
    
    Call SetMoneyList

    '读取简码匹配方式
    sta.Panels("MedicareType").Visible = mbytInState = 0
    sta.Panels("PY").Visible = mbytInState = 0 And gbln简码匹配方式切换 '35242
    sta.Panels("WB").Visible = mbytInState = 0 And gbln简码匹配方式切换
    If mbytInState = 0 Then
        '简码匹配方式：0-拼音,1-五笔
        i = Val(zlDatabase.GetPara("简码方式"))
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
    If mint记录性质 = 1 Then
        lblTitle.Caption = gstrUnitName & "病人收费单"
    ElseIf mint记录性质 = 2 Then
        lblTitle.Caption = gstrUnitName & "病人记帐单"
    End If
    txt应收.Text = gstrDec: txt实收.Text = gstrDec
    
    Select Case mbytInState
        Case 0 '执行
            Call SetShowCol
        Case 1 '查阅
            fraTitle.Enabled = False
            fraInfo.Enabled = False
            fraUnit.Enabled = False
            fraDrawDept.Enabled = False
            fraAppend.Enabled = False
            Bill.Active = False
            cmdSel.Visible = False
            cmdOK.Visible = False
            cmdCancel.Caption = "退出(&X)"
        Case 3 '销帐
            fraTitle.Enabled = False
            fraInfo.Enabled = False
            fraUnit.Enabled = False
            fraDrawDept.Enabled = False
            fraAppend.Enabled = False
            cmdSel.Visible = False
            '暂时不支持部份删除
            If mint记录性质 <> 1 And False Then
                Call ShowDeleteCol(True)
                Bill.Active = True
            Else
                Bill.Active = False
            End If
    End Select
    
    If mbytInState <> 0 Then
        lblPreNO.Visible = False: txtPreNO.Visible = False
        lbl应收.Top = lbl应收.Top + txtPreNO.Height / 2
        txt应收.Top = txt应收.Top + txtPreNO.Height / 2
        lbl实收.Top = lbl实收.Top + txtPreNO.Height * 0.75
        txt实收.Top = txt实收.Top + txtPreNO.Height * 0.75
    End If
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
    
    mblnWarnCloseed = False
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
        " A.病人ID,Nvl(B.主页ID,0) 主页ID,To_Number(Nvl(B.当前病区ID,[3])) as 病区ID," & _
        " Nvl(B.出院科室ID,[3]) as 科室ID,B.入院日期,B.出院日期," & _
        " A.门诊号,B.住院号,B.出院病床 as 床号,NVL(B.姓名,A.姓名) 姓名,NVL(B.性别,A.性别) 性别 ,NVL(B.年龄,A.年龄) 年龄,Nvl(B.费别,A.费别) as 费别," & _
        " A.担保人," & IIF(mint病人来源 = 2 And mint记录性质 = 2, "Decode(A.担保额,null,A.担保额,Zl_Patientsurety(A.病人ID,B.主页ID)) 担保额,", "A.担保额,") & _
        " Nvl(B.医疗付款方式,A.医疗付款方式) as 医疗付款方式,Y.编码 as 付款码,zl_PatiWarnScheme(A.病人ID,B.主页ID) as 适用病人," & _
        " zl_PatiDayCharge(A.病人ID) as 当日额,Nvl(B.险类,A.险类) as 险类,Nvl(B.病人性质,0) as 病人性质,b.审核标志,B.备注 as 病人备注" & _
        " From 病人信息 A,病案主页 B,病人余额 X,医疗付款方式 Y" & _
        " Where A.病人ID=B.病人ID(+) And A.病人ID=X.病人ID(+) And X.类型(+) = " & IIF(mint病人来源 = 1, 1, 2) & strSQL & _
        " And A.病人ID=[1] And B.主页ID(+)=[2]   And A.医疗付款方式=Y.名称(+)"
        
    On Error GoTo errH
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID, mlng病人科室id)
    If Not mrsInfo.EOF Then
        If Not IsNull(mrsInfo!险类) Then
            txt姓名.ForeColor = vbRed
        End If
        
        '除了门诊划价以外要处理的内容
        If mint记录性质 = 2 Then
            If mint病人来源 = 2 Then
                '49501:住院
                If zlIsAllowFeeChange(mrsInfo!病人ID, Val(Nvl(mrsInfo!主页ID)), Val(Nvl(mrsInfo!审核标志))) = False Then
                    Set mrsMedAudit = Nothing
                    Set mrsInfo = New ADODB.Recordset: txt姓名.Text = "":
                    mlng病人ID = 0
                    If txt姓名.Enabled And txt姓名.Visible Then txt姓名.SetFocus
                    mblnWarnCloseed = True
                    Exit Function
                End If
            End If
            '刷新病人费用状况
            Set rsTmp = GetMoneyInfo(mrsInfo!病人ID, IIF(mint病人来源 = 1, 0, mlng主页ID), mcurModiMoney)
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
            strSQL = "Select Nvl(报警方法,1) as 报警方法," & _
                " 报警值,报警标志1,报警标志2,报警标志3 From 记帐报警线" & _
                " Where 适用病人=[2] And " & IIF(mint病人来源 = 1, "Nvl(病区ID,0)=0", "病区ID=[1]")
            Set mrsWarn = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Nvl(mrsInfo!病区ID, 0)), CStr(Nvl(mrsInfo!适用病人)))
            
            '--------------------------------------------------------------------------------------------------------------------------------------------------------------
            '刘兴洪:26952
            Dim cur余额 As Currency, curItemMoney As Currency, curTotal As Double
            curItemMoney = 0
            curTotal = GetBillTotal(mobjBill)
            cur余额 = Val(txt实收.Tag)
            If gbln报警包含划价费用 Then cur余额 = cur余额 - GetPriceMoneyTotal(mrsInfo!病人ID, mint病人来源)
            
            If mbln费用登记 = False Then    '30504
            
                mbytWarn = BillingWarn(Me, mstrPrivs, mrsWarn, mrsInfo!姓名, cur余额, Val(Nvl(mrsInfo!当日额)) - mcurModiMoney, curTotal, _
                     Nvl(mrsInfo!担保额, 0), "", "", mstrWarn, mintWarn, , True)
                '返回:0;没有报警,继续
                '     1:报警提示后用户选择继续
                '     2:报警提示后用户选择中断
                '     3:报警提示必须中断
                '     4:强制记帐报警,继续
                '     5.报警提示后用户选择继续,但只允许保存存为划价单
                If mbytWarn = 2 Or mbytWarn = 3 Then
                    Set mrsMedAudit = Nothing
                    Set mrsInfo = New ADODB.Recordset: txt姓名.Text = "":
                    mlng病人ID = 0
                    If txt姓名.Enabled And txt姓名.Visible Then txt姓名.SetFocus
                    mblnWarnCloseed = True
                    Exit Function
                End If
                '--------------------------------------------------------------------------------------------------------------------------------------------------------------
                If mrsWarn.EOF Then mrsWarn.Close '用于后面状态判断
            End If
        End If
                            
        '住院记帐才处理的内容
        If mint病人来源 = 2 Then
            '急诊费用
            If Not IsNull(mrsInfo!险类) Then
                chk急诊.value = 0: chk急诊.Visible = True
            Else
                chk急诊.value = 0: chk急诊.Visible = False
            End If
            
            '发生时间
            If Not IsNull(mrsInfo!出院日期) Then
                txtDate.Text = Format(mrsInfo!出院日期, "yyyy-MM-dd HH:mm:ss")
            Else
                txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
            End If
        End If
        
        Call LoadPatientBaby(cboBaby, mrsInfo!病人ID, mrsInfo!主页ID)
        
        '显示病人信息
        txt姓名.Text = Nvl(mrsInfo!姓名)
        txt性别.Text = Nvl(mrsInfo!性别)
        txt年龄.Text = Nvl(mrsInfo!年龄)
        txt费别.Text = Nvl(mrsInfo!费别)
        txt付款方式.Text = Nvl(mrsInfo!医疗付款方式)
        txt付款方式.Tag = Nvl(mrsInfo!付款码, 0) '不要填写为空
        txt床号.Text = Nvl(mrsInfo!床号)
        
        '刘兴洪 问题:26953 日期:2009-12-25 15:21:47
        txt病人备注 = Nvl(mrsInfo!病人备注)
        If mint病人来源 = 1 Then
            lbl住院号.Caption = "门诊号"
            txt住院号.Text = Nvl(mrsInfo!门诊号)
        Else
            lbl住院号.Caption = "住院号"
            txt住院号.Text = Nvl(mrsInfo!住院号)
        End If
        
        txt担保人.Text = Nvl(mrsInfo!担保人)
        txt担保额.Text = Format(Nvl(mrsInfo!担保额), "0.00")
        
        With mobjBill
            .病人ID = Nvl(mrsInfo!病人ID, 0)
            .主页ID = Nvl(mrsInfo!主页ID, 0)
            .病区ID = Nvl(mrsInfo!病区ID, 0)
            .科室ID = Nvl(mrsInfo!科室ID, 0)
            .床号 = Nvl(mrsInfo!床号)
            .标识号 = IIF(mint病人来源 = 1, Nvl(mrsInfo!门诊号), Nvl(mrsInfo!住院号))
            .姓名 = Nvl(mrsInfo!姓名)
            .性别 = Nvl(mrsInfo!性别)
            .年龄 = Nvl(mrsInfo!年龄)
            .费别 = Nvl(mrsInfo!费别)
        End With
        
        '在第一次进入时读取病人审批费用项目信息
        If Not Visible And mint病人来源 = 2 And mint记录性质 = 2 And mbytInState = 0 Then Set mrsMedAudit = GetAuditRecord(mrsInfo!病人ID, mrsInfo!主页ID)
        
        GetPatient = True
    Else
        Set mrsMedAudit = Nothing
        Set mrsInfo = New ADODB.Recordset
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
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
                If CheckItemHaveSub(i) Then                          '主项或独立项
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
        
    Dim dblAllTime As Double, dblPrice As Double, dblPriceSingle As Double
    
    On Error GoTo errH
    
    gstrSQL = _
        " Select B.收入项目ID,C.名称,C.收据费目,B.现价,B.原价,B.加班加价率,B.附术收费率,B.缺省价格 " & _
        " From 收费项目目录 A,收费价目 B,收入项目 C " & _
        " Where B.收费细目ID = A.ID And C.ID = B.收入项目ID " & _
        "       And ((Sysdate Between B.执行日期 and B.终止日期) Or (Sysdate>=B.执行日期 And B.终止日期 is NULL)) " & _
        "       And A.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mobjBill.Details(lngRow).收费细目ID)
    If Not rsTmp.EOF Then
        With mobjBill.Details(lngRow)
            '先获取操作员以前输入的变价金额
            If .Detail.变价 Then
                '计算药品时价(分批或不分批)
                '必然有记录(输入该项目时已判断)
                dblAllTime = .付数 * .数次
                If dblAllTime <> 0 Then
                    dblPrice = Get时价材料应收金额(mlng虚拟库房ID, .收费细目ID, .Detail.批次, dblAllTime, gstrDec, dblPriceSingle)
                    If dblAllTime <> 0 Then
                        '数量未分解完毕
                        MsgBox "第 " & lngRow & " 行时价卫生材料""" & .Detail.名称 & """库存不足,无法计算价格！", vbInformation, gstrSysName
                        dblMoney = 0
                    Else
                        '注意：货币型最多只能保留4位小数,且不四舍五入,所以需要手工舍入;而用其它型在计算精度上又有问题
                        dblAllTime = .付数 * .数次
                        dblMoney = IIF(dblPriceSingle = 0, Format(dblPrice / dblAllTime, gstrDecPrice), dblPriceSingle) '这里结果是按售价单位
                    End If
                Else
                    dblMoney = 0
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
                    .标准单价 = Format(dblMoney, gstrDecPrice)
                Else
                    .标准单价 = Format(Nvl(rsTmp!现价, 0), gstrDecPrice)
                End If
                '应收金额=单价 * 付数 * 数次
                If mobjBill.Details(lngRow).Detail.变价 Then
                    .应收金额 = dblPrice '保证应收金额与零售金额没有误差
                Else
                    .应收金额 = .标准单价 * mobjBill.Details(lngRow).付数 * mobjBill.Details(lngRow).数次
                End If
                
                '加班费用率计算
                dbl加班加价率 = 0
                If mobjBill.加班标志 = 1 And mobjBill.Details(lngRow).Detail.加班加价 Then
                    dbl加班加价率 = Nvl(rsTmp!加班加价率, 0) / 100
                    .应收金额 = .应收金额 * (1 + dbl加班加价率)
                End If
                
                .应收金额 = CCur(Format(.应收金额, gstrDec))
                
                dblAllTime = mobjBill.Details(lngRow).付数 * mobjBill.Details(lngRow).数次
                If mbln费用登记 Or .应收金额 = 0 Then
                    .实收金额 = 0
                Else
                    If mobjBill.Details(lngRow).Detail.屏蔽费别 Or bln从项汇总折扣 Then
                        .实收金额 = .应收金额
                    Else
                        .实收金额 = CCur(Format(ActualMoney(mobjBill.费别, .收入项目ID, .应收金额, _
                            mobjBill.Details(lngRow).收费细目ID, mobjBill.Details(lngRow).执行部门ID, _
                            dblAllTime, dbl加班加价率), gstrDec))
                    End If
                End If
                
                '获取项目保险信息,医保病人才处理,不需要连接医保
                If Not IsNull(mrsInfo!险类) Then
                    strInfo = gclsInsure.GetItemInsure(mobjBill.病人ID, mobjBill.Details(lngRow).收费细目ID, .实收金额, False, mrsInfo!险类, _
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
            Case "项目"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.名称
            Case "商品名"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.商品名
            Case "规格"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.规格
            Case "单位"
                    Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.计算单位
            Case "付数"
                Bill.TextMatrix(lngRow, i) = IIF(mobjBill.Details(lngRow).付数 = 0, 1, mobjBill.Details(lngRow).付数)
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
                Bill.TextMatrix(lngRow, i) = Format(dbl单价, gstrDecPrice)
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
            Case "类型"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.类型
            Case "商品条码"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.商品条码
            Case "内部条码"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.内部条码
                
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
 
Private Sub SetDetail(Detail As Detail, lngRow As Long, lngDoUnit As Long, Optional bytParent As Byte = 0)
'功能：根据指定的收费细目对象设定单据指点定行的收费细目(新增的或修改)
'说明：
'      1.用于新输入或更改收费细目行！！！
'      2.当bytParent<>0时,则为设置从属项目,从属项目一定是新增行,且主项目一定存在

    Dim tmpIncomes As New BillInComes
    Dim intPay As Integer, i As Long, dblTime As Double
    
    '取其它中药的付数
    intPay = 1
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
                Bill.RowData(lngRow) = Asc("4")
                '初始数次
                If Detail.固有从属 = 0 Then '非固有从属
                    dblTime = Detail.从项数次
                ElseIf Detail.固有从属 = 1 Then '固定的固有从属
                    dblTime = IIF(Detail.从项数次 = 0, 1, Detail.从项数次)
                ElseIf Detail.固有从属 = 2 Then '按比例的固有从属
                    dblTime = Detail.从项数次 * mobjBill.Details(bytParent).数次
                End If
            Else
                dblTime = 1
            End If
            mobjBill.Details.Add tmpIncomes, Detail, .ID, CByte(lngRow), CInt(bytParent), .类别, .计算单位, intPay, dblTime, 0, mlng执行库房ID, ""
        End With
    Else '如果该行已经存在,则修改
        dblTime = 1
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
            .执行部门ID = mlng执行库房ID
        End With
    End If
End Sub

Private Function ShouldDO(lngRow As Long) As Boolean
'功能：判断该行是否应该取从属项目
'说明：仅该行收费项目有从属项目及尚未取才取。
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long, blnExist As Boolean
    Dim strSQL As String
    
    strSQL = "Select count(从项ID) as NUM from 收费从属项目 where 主项ID=[1]"
    On Error GoTo errH
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
    chk加班.value = IIF(OverTime, 1, 0)
    txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
                
    Call LoadPatientBaby(cboBaby, 0, 0)
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
        .加班标志 = chk加班.value
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

Private Function SaveBill(Optional ByRef strNos As String) As Boolean
'功能:保存当前输入的记帐单据(适用住院记帐、划价、或对两者的修改)
'入口:mobjBill=单据对象
'出口:保存是否成功
    Dim int行号 As Integer, int序号 As Integer, int价格父号 As Integer
    Dim dbl数次 As Double, dbl单价 As Double
    Dim intInsure As Integer, strNO As String, strTmp As String
    Dim arrSQL As Variant, i As Long, j As Long
    Dim int划价 As Integer, bln上传 As Boolean
    Dim strSQL As String, strStuffDept As String '记录卫料发料部门
    Dim strDeptIDs As String, str汇总号 As String
    Dim cllProExeute As New Collection, varTemp As Variant
    Dim rsTmp As ADODB.Recordset
    Dim blnTrans As Boolean
    
    strNos = ""
    If mstrOriginalNO = "" Then
        If mint记录性质 = 1 Then
            mobjBill.NO = zlDatabase.GetNextNo(13)
        Else
            mobjBill.NO = zlDatabase.GetNextNo(14)
        End If
    Else
        mobjBill.NO = mstrOriginalNO
    End If
    mobjBill.发生时间 = CDate(txtDate.Text)
    mobjBill.登记时间 = zlDatabase.Currentdate
    
    int序号 = 0
    arrSQL = Array()
    Set cllProExeute = New Collection
    For Each mobjBillDetail In mobjBill.Details
        If mobjBillDetail.数次 <> 0 Then
            For Each mobjBillIncome In mobjBillDetail.InComes
                int序号 = int序号 + 1 '当前记录序号
                
                '单据主体
                With mobjBill
                    If mint病人来源 = 2 Then
                        gstrSQL = "zl_住院记帐记录_INSERT('" & .NO & "'," & int序号 & "," & .病人ID & "," & ZVal(.主页ID) & "," & _
                            IIF(.标识号 = "", "NULL", "'" & .标识号 & "'") & "," & "'" & .姓名 & "','" & .性别 & "','" & .年龄 & "','" & .床号 & "','" & .费别 & "'," & _
                            ZVal(.病区ID) & "," & ZVal(.科室ID) & "," & .加班标志 & "," & .婴儿费 & "," & .开单部门ID & ",'" & .开单人 & "',"
                    Else
                        If mint记录性质 = 2 Then
                            gstrSQL = "zl_门诊记帐记录_INSERT('" & .NO & "'," & int序号 & "," & .病人ID & "," & _
                                IIF(.标识号 = "", "NULL", "'" & .标识号 & "'") & "," & "'" & .姓名 & "','" & .性别 & "','" & .年龄 & "'," & _
                                "'" & .费别 & "'," & .加班标志 & "," & .婴儿费 & "," & _
                                ZVal(.科室ID) & "," & .开单部门ID & ",'" & .开单人 & "',"
                        Else
                            gstrSQL = "zl_门诊划价记录_Insert('" & .NO & "'," & int序号 & "," & .病人ID & "," & ZVal(.主页ID) & "," & _
                                IIF(.标识号 = "", "NULL", "'" & .标识号 & "'") & ",'" & IIF(Val(txt付款方式.Tag) = 0, "", txt付款方式.Tag) & "','" & .姓名 & "'," & _
                                "'" & .性别 & "','" & .年龄 & "','" & .费别 & "'," & .加班标志 & "," & _
                                  ZVal(.科室ID) & "," & .开单部门ID & ",'" & .开单人 & "',"
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
                        If mobjBill.Details(.序号).从属父号 = 0 Then
                            For i = .序号 + 1 To mobjBill.Details.Count
                                If mobjBill.Details(i).从属父号 = .序号 Then
                                    mobjBill.Details(i).从属父号 = int序号
                                End If
                            Next
                        End If
                    End If
                    gstrSQL = gstrSQL & .从属父号 & "," & .收费细目ID & ",'" & .收费类别 & "','" & .计算单位 & "',"
                    
                    If mint病人来源 = 2 Then
                        gstrSQL = gstrSQL & IIF(.保险项目否, 1, 0) & "," & ZVal(.保险大类ID) & ",'" & .保险编码 & "',"
                    ElseIf mint记录性质 = 1 Then
                        gstrSQL = gstrSQL & "NULL,"
                    End If
                    
                    dbl数次 = .数次
                    gstrSQL = gstrSQL & IIF(.付数 = 0, 1, .付数) & "," & dbl数次 & "," & .附加标志 & "," & mlng执行库房ID & ","
                End With
                
                '收入项目部份
                With mobjBillIncome
                    dbl单价 = .标准单价
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
                If int划价 = 0 Then bln上传 = True '只要存在不是划价单就要上传
                
                '收集卫料发料部门,以便自动发料,门诊病人仅记帐时(发送为划价时不管),住院病人只有记帐
                'mint病人来源 :1-门诊病人,2-住院病人
                'mint记录性质 :1-收费(划价),2-记帐(门/住)
                
                With mobjBillDetail
                    If (mint病人来源 = 1 And mint记录性质 = 2 And gbln门诊自动发料 Or mint病人来源 = 2 And gbln住院自动发料) And int划价 = 0 Then
                        strStuffDept = "," & mlng执行库房ID
                    End If
                End With
                
                If mint病人来源 = 2 Then
                    gstrSQL = gstrSQL & int划价 & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & _
                        "0," & mlng卫材类别ID & "," & _
                        "NULL,'" & mobjBillDetail.摘要 & "'," & chk急诊.value & "," & ZVal(mlng医嘱ID) & "," & _
                        "Null,Null,Null,Null,Null,Null,'" & mobjBillDetail.Detail.类型 & "'," & _
                        IIF(mobjBill.开单部门ID = mlng开嘱科室ID, "1", "0") & "," & mlng开单科室ID & ",NULL,-1,1," & mobjBillDetail.Detail.批次 & ")"
                        '    医技补临床费用_In Number := 0,
                        '    领药部门id_In     药品收发记录.对方部门id%Type := Null,
                        '    中药形态_In       住院费用记录.结论%Type := Null,
                        '    医疗小组id_In     住院费用记录.医疗小组id%Type := -1,
                        '    备货材料_In       Number := 0,
                        '    批次_In           药品收发记录.批次%Type := Null
                Else
                    If mint记录性质 = 2 Then
                        gstrSQL = gstrSQL & int划价 & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & _
                              mlng卫材类别ID & "," & _
                            "NULL,'" & mobjBillDetail.摘要 & "'," & ZVal(mlng医嘱ID) & ",NULL,NULL,NULL,NULL,NULL,1,NULL,1," & mobjBillDetail.Detail.批次 & ")"
                            '    频次_In       药品收发记录.频次%Type := Null,
                            '    单量_In       药品收发记录.单量%Type := Null,
                            '    用法_In       药品收发记录.用法%Type := Null, --用法[|煎法]
                            '    期效_In       药品收发记录.扣率%Type := Null,
                            '    计价特性_In   药品收发记录.扣率%Type := Null,
                            '    门诊标志_In   门诊费用记录.门诊标志%Type := 1,
                            '    中药形态_In   门诊费用记录.结论%Type := Null,
                            '    备货材料_In   Number := 0,
                            '    批次_In       药品收发记录.批次%Type := Null
                    Else
                        gstrSQL = gstrSQL & "'" & UserInfo.姓名 & "'," & _
                             mlng卫材类别ID & "," & _
                            "'" & mobjBillDetail.摘要 & "'," & ZVal(mlng医嘱ID) & ",NULL,NULL,NULL,NULL,NULL,1,NULL,NULL,NULL,NULL,NULL,1," & mobjBillDetail.Detail.批次 & "  )"
                            '频次_In       药品收发记录.频次%Type := Null,
                            '单量_In       药品收发记录.单量%Type := Null,
                            '用法_In       药品收发记录.用法%Type := Null, --用法[|煎法]
                            '期效_In       药品收发记录.扣率%Type := Null,
                            '计价特性_In   药品收发记录.扣率%Type := Null,
                            '病人来源_In   Number := 1,
                            '保险编码_In   门诊费用记录.保险编码%Type := Null,
                            '费用类型_In   门诊费用记录.费用类型%Type := Null,
                            '保险项目否_In 门诊费用记录.保险项目否%Type := Null,
                            '保险大类id_In 门诊费用记录.保险大类id%Type := Null,
                            '中药形态_In   门诊费用记录.结论%Type := Null,
                            '备货材料_In   Number := 0,
                            '批次_In       药品收发记录.批次%Type := Null
                    End If
                End If
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = mobjBillDetail.收费细目ID & ";" & gstrSQL
            Next
        End If
    Next
    
    '-----------------------------------------------------------------------------------------------------------------
    If mstrOriginalNO = "" Then
        '插入医嘱院加费用
        gstrSQL = "ZL_病人医嘱附费_Insert(" & mlng医嘱ID & "," & mlng发送号 & "," & mint记录性质 & ",'" & mobjBill.NO & "')"
    Else
        '补主费用
        gstrSQL = "ZL_病人医嘱发送_计费(" & mlng医嘱ID & "," & mlng发送号 & ")"
    End If
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
        gcnOracle.BeginTrans: blnTrans = True
        
            For i = 0 To UBound(arrSQL)
                Call zlDatabase.ExecuteProcedure(Mid(arrSQL(i), InStr(arrSQL(i), ";") + 1), Me.Caption)
            Next
            
            '-----------------------------------------------------------------------
            '执行自动发料
            If strStuffDept <> "" Then
                strStuffDept = Mid(strStuffDept, 2)
                varTemp = Split(strStuffDept, ",")
                For i = 0 To UBound(varTemp)
                    '69902:刘尔旋,2014-02-09,只对同开单科室一致的执行科室项目进行自动发料
                    If Val(varTemp(i)) = Val(cbo开单科室.ItemData(cbo开单科室.ListIndex)) Then
                        strSQL = "zl_材料收发记录_处方发料(" & Val(varTemp(i)) & ",25,'" & mobjBill.NO & "','" & UserInfo.姓名 & "','" & UserInfo.姓名 & "','" & UserInfo.姓名 & "',1,Sysdate)"
                        zlAddArray cllProExeute, strSQL
                    End If
                Next
            End If
            '执行发药和发料
            zlExecuteProcedureArrAy cllProExeute, Me.Caption, False, False
            '-----------------------------------------------------------------------
            
            
            '医保接口
            '1.医保记帐作废上传
            If mint病人来源 = 2 And mstrInNO <> "" And intInsure <> 0 Then
                If gclsInsure.GetCapability(support记帐作废上传, mlng病人ID, intInsure) And Not gclsInsure.GetCapability(support记帐完成后上传, mlng病人ID, intInsure) Then
                    If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                        gcnOracle.RollbackTrans: Exit Function
                    End If
                End If
            End If
            
            '2.记帐实时上传
            If mint病人来源 = 2 And bln上传 And Not IsNull(mrsInfo!险类) Then
                '医保传输费用明细
                If gclsInsure.GetCapability(support记帐上传, mlng病人ID, mrsInfo!险类) And Not gclsInsure.GetCapability(support记帐完成后上传, mlng病人ID, mrsInfo!险类) Then
                    strTmp = ""
                    If Not gclsInsure.TranChargeDetail(2, mobjBill.NO, 2, 1, strTmp, , mrsInfo!险类) Then
                        gcnOracle.RollbackTrans
                        If strTmp <> "" Then MsgBox strTmp, vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            End If
        
        gcnOracle.CommitTrans: blnTrans = False
        
        '医保接口
        '1.医保记帐作废上传
        If mint病人来源 = 2 And mstrInNO <> "" And intInsure > 0 Then
            If gclsInsure.GetCapability(support记帐作废上传, mlng病人ID, intInsure) And gclsInsure.GetCapability(support记帐完成后上传, mlng病人ID, intInsure) Then
                If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                    MsgBox "单据""" & mstrInNO & """向医保传送失败,该单据的费用已删除！", vbInformation, gstrSysName
                End If
            End If
        End If
        
        '2.记帐实时上传
        If mint病人来源 = 2 And bln上传 And Not IsNull(mrsInfo!险类) Then
            '医保传输费用明细
            If MCPAR.记帐上传 And MCPAR.记帐完成后上传 Then
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
    '74231,冉俊明,2014-7-21,项目开单后立即收费或记帐审核
    strNos = mobjBill.NO
    
    SaveBill = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
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
        If mint记录性质 = 1 Or (mint记录性质 = 2 And mint病人来源 = 1) Then
            blnNOMoved = zlDatabase.NOMoved("门诊费用记录", strNO, "记录性质=", mint记录性质)
        Else
            blnNOMoved = zlDatabase.NOMoved("住院费用记录", strNO, "记录性质=", mint记录性质)
        End If
    End If
    
    On Error GoTo errH
    
    Call ClearRows: Call Bill.ClearBill
    Call SetColNum: Call ClearMoney
    
    If mstrFeeTab = "住院费用记录" Then
        strSQL = _
        " Select A.病人ID,Nvl(A.主页ID,0) 主页ID,A.姓名,A.性别,A.年龄,A.费别,A.床号,A.标识号," & _
        "           A.病人病区ID,A.开单部门ID,A.加班标志,A.婴儿费,A.开单人,A.划价人,A.操作员姓名," & _
        "           A.开单部门ID,A.执行部门ID," & IIF(zlIsShowDeptCode, "C.编码||'-'||", "") & "C.名称 as 开单部门," & IIF(zlIsShowDeptCode, "C.编码||'-'||", "") & "C.名称  as 执行部门,A.发生时间," & _
        "            B.医疗付款方式,B.担保人,B.担保额,A.是否急诊,B1.备注 as 病人备注" & _
        " From 住院费用记录 A,病人信息 B,部门表 C,部门表 C1,病案主页 B1 " & _
        " Where Rownum=1  And A.病人id=B1.病人id(+) and A.主页id=B1.主页ID(+) And NO=[1] And A.记录性质=[2]" & _
        "       And A.病人ID=B.病人ID And Instr([3],A.记录状态)>0" & _
                IIF(mstrTime <> "", " And A.登记时间=[4]", "") & _
        "     And A.开单部门ID=C.ID and A.执行部门ID=C1.ID(+)"
    Else
        strSQL = _
        " Select A.病人ID,0 as 主页ID,A.姓名,A.性别,A.年龄,A.费别,A.付款方式 as 床号,A.标识号," & _
        "           0 as 病人病区ID,A.开单部门ID,A.加班标志,A.婴儿费,A.开单人,A.划价人,A.操作员姓名," & _
        "           A.开单部门ID,A.执行部门ID," & IIF(zlIsShowDeptCode, "C.编码||'-'||", "") & "C.名称 as 开单部门 ," & IIF(zlIsShowDeptCode, "C.编码||'-'||", "") & "C.名称  as 执行部门,A.发生时间," & _
        "           B.医疗付款方式,B.担保人,B.担保额,A.是否急诊,Null as 病人备注" & _
        " From 门诊费用记录 A,病人信息 B,部门表 C ,部门表 C1" & _
        " Where Rownum=1  And NO=[1] And A.记录性质=[2]" & _
        "           And A.病人ID=B.病人ID And Instr([3],A.记录状态)>0" & _
                    IIF(mstrTime <> "", " And A.登记时间=[4]", "") & _
        "           And A.开单部门ID=C.ID and A.执行部门ID=C1.ID(+)"
    End If
    If blnNOMoved Then
        strSQL = Replace(strSQL, mstrFeeTab, "H" & mstrFeeTab)
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, mint记录性质, _
        IIF(mblnDelete, "2", "0,1,3"), CDate(IIF(mstrTime = "", "1990-01-01", mstrTime)))
    If rsTmp.EOF Then
        MsgBox "没有发现该单据。", vbInformation, gstrSysName
        Exit Function
    End If
    
    If mlng病人ID = 0 Then mlng病人ID = Nvl(rsTmp!病人ID, 0)
    
    cboNO.Text = strNO
    txt姓名.Text = Nvl(rsTmp!姓名)
    txt性别.Text = Nvl(rsTmp!性别)
    txt年龄.Text = Nvl(rsTmp!年龄)
    If Nvl(rsTmp!主页ID, 0) <> 0 Then
        txt床号.Text = Nvl(rsTmp!床号)
    End If
    
    '刘兴洪 问题:26953 日期:2009-12-25 15:23:48
    txt病人备注.Text = Nvl(rsTmp!病人备注)
    If mint病人来源 = 1 Then
        lbl住院号.Caption = "门诊号"
    Else
        lbl住院号.Caption = "住院号"
    End If
    txt住院号.Text = Nvl(rsTmp!标识号)
    
    txt费别.Text = Nvl(rsTmp!费别)
    txt担保人.Text = Nvl(rsTmp!担保人)
    txt担保额.Text = Format(Nvl(rsTmp!担保额), "0.00")
    txt付款方式.Text = Nvl(rsTmp!医疗付款方式)
    
    mblnCboNotClick = True
    cbo开单科室.AddItem Nvl(rsTmp!开单部门)
    cbo开单科室.ItemData(cbo开单科室.NewIndex) = Nvl(rsTmp!开单部门ID, 0)
    cbo开单科室.ListIndex = cbo开单科室.NewIndex
    
    
    mlng执行库房ID = Nvl(rsTmp!执行部门ID, 0)
    cbo执行部门.AddItem Nvl(rsTmp!执行部门)
    cbo执行部门.ItemData(cbo执行部门.NewIndex) = mlng执行库房ID
    cbo执行部门.ListIndex = cbo执行部门.NewIndex

    mlng虚拟库房ID = Set虚拟库房ID(mlng执行库房ID)
    mblnCboNotClick = False
    
    If Nvl(rsTmp!是否急诊, 0) = 1 Then
        chk急诊.value = 1: chk急诊.Visible = True
    End If
    
    chk加班.value = Nvl(rsTmp!加班标志, 0)
    Call LoadPatientBaby(cboBaby, rsTmp!病人ID, rsTmp!主页ID)
    Call zlControl.CboLocate(cboBaby, Nvl(rsTmp!婴儿费, 0), True)
    
    '开单人
    Call GetCboIndex(cbo开单人, Nvl(rsTmp!开单人))
    If cbo开单人.ListIndex = -1 And Not IsNull(rsTmp!开单人) Then
        cbo开单人.AddItem rsTmp!开单人
        cbo开单人.ListIndex = cbo开单人.NewIndex
    End If
    
    txtDate.Text = Format(rsTmp!发生时间, "yyyy-MM-dd HH:mm:ss")
    
    If mint记录性质 = 2 Then
        Set rsPatiMoney = GetMoneyInfo(rsTmp!病人ID, IIF(mint病人来源 = 1, 0, rsTmp!主页ID))
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
            "       Nvl(A.付数,1)*A.数次 as 原始数量" & _
            " From " & mstrFeeTab & " A " & _
            " Where A.NO=[1] And A.记录状态 IN(0,1,3) And A.价格父号 is NULL" & _
            "            And A.记录性质=[2]"
        
        '读取药品收发记录中的准退数
        strSQL2 = _
            " Select A.费用ID,Max(A.批次) as 批次,Max(A.商品条码) as 商品条码 ,Max(内部条码) as 内部条码, " & _
            "       Sum(Nvl(A.付数,1)*A.实际数量)  as 准退数量" & _
            " From 药品收发记录 A " & _
            " Where A.NO=[1] And MOD(A.记录状态,3)=1" & _
            "       And A.审核人 is NULL And Instr([3],','||A.单据||',')>0" & _
            " Group by A.费用ID"
        
        '整张单据汇总结果(明细到收费细目)
        '执行状态应该在原始记录上判断(部分退药且部分退费的记录)
        '当退两次以上时"记录状态,序号"重复,AVG有问题,所以要用"执行状态"
        strSQL = "Select Nvl(价格父号,序号) From " & mstrFeeTab & _
            " Where 记录性质=[2] And 记录状态 IN(0,1,3) And NO=[1]" & _
            " And Nvl(执行状态,0)<>1" & IIF(mlng医嘱ID <> 0, " And 医嘱序号+0=[8]", "")
        
        '如果已结帐单据禁止销帐,或是医保记帐的单据。则在原始单据行中只取未结帐部分
        If mint记录性质 = 2 Then
            If mint病人来源 = 2 Then intInsure = BillExistInsure(strNO)
            If intInsure <> 0 Then
                blnDo = Not gclsInsure.GetCapability(support允许冲销已结帐的记帐单据, rsTmp!病人ID, intInsure)
            Else
                blnDo = gbytBillOpt = 2
            End If
            If blnDo Then
                strSQL = strSQL & " And Nvl(价格父号,序号) IN" & _
                    " (" & _
                    " Select Nvl(价格父号,序号) as 序号" & _
                    " From " & mstrFeeTab & _
                    " Where NO=[1] And 记录性质 IN(2,12)" & _
                    " Group by Nvl(价格父号,序号)" & _
                    " Having Sum(Nvl(结帐金额,0))=0" & _
                    " )"
            End If
        End If
        
        '因为是将要汇总求有剩余数量的，所以不能用直接用时间限制，用序号限制
        strSQL = _
            " Select A.记录状态,A.执行状态,Nvl(A.价格父号,A.序号) as 序号," & _
            "       C.编码,C.名称 as 类别,A.收费细目ID,B.名称,B.规格,Nvl(A.费用类型,B.费用类型) 费用类型, A.计算单位," & _
            "       Avg(Nvl(A.付数,1)) as 付数, Avg(A.数次) as 数次," & _
            "       Sum(A.标准单价) as 单价, Sum(A.应收金额) as 应收金额,Sum(A.实收金额) as 实收金额, " & _
            "       D.名称 as 执行部门,A.附加标志" & _
            " From " & mstrFeeTab & " A,收费项目目录 B,收费项目类别 C,部门表 D " & _
            " Where A.收费细目ID=B.ID And C.编码=A.收费类别 And A.执行部门ID=D.ID(+)" & _
            "           And A.记录性质=[2]" & _
            "           And A.NO=[1] And Nvl(A.价格父号,A.序号) IN(" & strSQL & ")" & _
            " Group by A.记录状态,A.执行状态,Nvl(A.价格父号,A.序号),C.编码,C.名称,A.收费细目ID,B.名称," & _
            "           B.规格,Nvl(A.费用类型,B.费用类型),A.计算单位,D.名称,A.附加标志"
            
        '最后计算结果
        '当"准退数量=原始数量"时,付数才保留
        '排开已经全部退费的行(执行状态=0的一种可能)
        '有剩余数量无准退数量的有两种情况：
            '1.无对应的收发记录(如普通费用或不跟踪在用的卫材),这时应用剩余数量
            '2.收发记录中已全部发放,即已全部执行,SQL已排除这种记录
        strSQL = _
            " Select A.序号,A.编码,A.类别,A.收费细目ID,A.名称,A.规格,A.费用类型,A.计算单位, " & _
            "           max(C.批次) as 批次,Max(C.商品条码) as 商品条码,Max(C.内部条码) as 内部条码," & _
            "           Decode(Sign(Nvl(C.准退数量,Sum(A.付数*A.数次))-B.原始数量),0,Avg(A.付数),1) as 准退付数," & _
            "           Decode(Sign(Nvl(C.准退数量,Sum(A.付数*A.数次))-B.原始数量),0,Sum(A.数次),Nvl(C.准退数量,Sum(A.付数*A.数次))) as 准退数次," & _
            "           Nvl(C.准退数量,Sum(A.付数*A.数次)) as 准退数量,Sum(A.付数*A.数次) as 剩余数量," & _
            "           A.单价,Sum(A.应收金额) as 剩余应收,Sum(A.实收金额) as 剩余实收,A.执行部门,A.附加标志" & _
            " From (" & strSQL & ") A,(" & strSQL1 & ") B,(" & strSQL2 & ") C" & _
            " Where A.序号=B.序号 And B.ID=C.费用ID(+)" & _
            " Group by A.序号,A.编码,A.类别,A.收费细目ID,A.名称,A.规格,A.费用类型," & _
            "           A.计算单位,A.单价,B.原始数量,C.准退数量,A.执行部门,A.附加标志" & _
            " Having Sum(A.付数*A.数次)<>0"
            
        strSQL = _
            " Select A.序号,A.编码,A.类别,Nvl(B.名称,A.名称) as 名称,E1.名称 as 商品名,A.规格," & _
            "       A.费用类型,A.计算单位,A.批次,A.商品条码,A.内部条码,A.准退付数 as 付数,A.准退数次 as 数次,A.单价," & _
            "       A.剩余应收*(A.准退数量/A.剩余数量) as 应收金额," & _
            "       A.剩余实收*(A.准退数量/A.剩余数量) as 实收金额," & _
            "       A.执行部门,A.附加标志" & _
            " From (" & strSQL & ") A,收费项目别名 B,收费项目别名 E1" & _
            " Where A.收费细目ID=B.收费细目ID(+) And B.码类(+)=1 And B.性质(+)=[6]" & _
            "       And  A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3" & _
            " Order by A.序号"
    Else
        '读取单据原始内容
        intSign = IIF(mblnDelete, -1, 1) '数量,金额正负符号
        
        strSQL2 = _
            " Select A.费用ID,Max(A.批次) as 批次,Max(A.商品条码) as 商品条码 ,Max(内部条码) as 内部条码 " & _
            " From 药品收发记录 A " & _
            " Where A.NO=[1] And MOD(A.记录状态,3)=1 And Instr([4],A.记录状态)>0 " & _
            "       And Instr([3],','||A.单据||',')>0" & _
            " Group by A.费用ID"
            
        strSQL = _
            "   Select A.收费细目ID,A.收费类别,A.执行部门ID,Nvl(A.价格父号,A.序号) as 序号,B.批次,B.商品条码,B.内部条码," & _
            "           A.计算单位,A.付数,A.数次,A.标准单价,A.应收金额,A.实收金额,A.附加标志,A.费用类型" & _
            "   From " & mstrFeeTab & " A,( " & strSQL2 & ") B " & _
            "   Where A.记录性质=[2] And Instr([4],A.记录状态)>0 And A.NO=[1]" & _
                        IIF(mstrTime <> "", " And A.登记时间=[5]", "") & _
            "           A.ID=B.费用ID(+) "
        If blnNOMoved Then
            strSQL = strSQL & " Union ALL " & Replace("药品收发记录", Replace(strSQL, mstrFeeTab, "H" & mstrFeeTab), "H药品收发记录")
        End If
        
        strSQL = _
            " Select A.序号,C.编码,C.名称 as 类别,A.收费细目ID,B.名称,B.规格,Nvl(A.费用类型,B.费用类型) 费用类型, A.计算单位, " & _
            "           max(A.批次) as 批次,Max(A.商品条码) as 商品条码,Max(A.内部条码) as 内部条码," & _
            "           Avg(Nvl(A.付数,1)) as 付数, Avg([7]*A.数次) as 数次," & _
            "           Sum(A.标准单价) as 单价," & _
            "           Sum([7]*A.应收金额) as 应收金额,Sum([7]*A.实收金额) as 实收金额, " & _
            "           D.名称 as 执行部门,A.附加标志" & _
            " From (" & strSQL & ") A,收费项目目录 B,收费项目类别 C,部门表 D" & _
            " Where A.收费细目ID=B.ID And C.编码=A.收费类别" & _
            "           And A.执行部门ID=D.ID(+) " & _
            " Group by A.序号,C.编码,C.名称,A.收费细目ID,B.名称,B.规格," & _
            "           Nvl(A.费用类型,B.费用类型),A.计算单位,D.名称,A.附加标志 "
            
        strSQL = _
            " Select A.序号,A.编码,A.类别,Nvl(B.名称,A.名称) as 名称,E1.名称 as 商品名,A.规格,A.费用类型," & _
            "       A.计算单位,A.批次,A.商品条码,A.内部条码,A.付数,A.数次,A.单价,A.应收金额,A.实收金额,A.执行部门,A.附加标志" & _
            " From (" & strSQL & ") A,收费项目别名 B,收费项目别名 E1" & _
            " Where A.收费细目ID=B.收费细目ID(+) And B.码类(+)=1 And B.性质(+)=[6]" & _
            "       And  A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3" & _
            " Order by 序号"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, mint记录性质, IIF(mint记录性质 = 2, ",9,25,", ",8,24,"), _
        IIF(mblnDelete, "2", "0,1,3"), CDate(IIF(mstrTime = "", "1990-01-01", mstrTime)), IIF(gbyt药品名称显示 = 1, 3, 1), intSign, mlng医嘱ID)
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
        Bill.TextMatrix(i, BillCol.项目) = rsTmp!名称
        Bill.TextMatrix(i, BillCol.商品名) = Nvl(rsTmp!商品名)
        Bill.TextMatrix(i, BillCol.规格) = Nvl(rsTmp!规格)
        Bill.TextMatrix(i, BillCol.单位) = Nvl(rsTmp!计算单位)
        Bill.TextMatrix(i, BillCol.付数) = Nvl(rsTmp!付数)
        Bill.TextMatrix(i, BillCol.数次) = FormatEx(rsTmp!数次, 5)
        Bill.TextMatrix(i, BillCol.单价) = Format(rsTmp!单价, gstrDecPrice)
        Bill.TextMatrix(i, BillCol.应收金额) = Format(rsTmp!应收金额, gstrDec)
        Bill.TextMatrix(i, BillCol.实收金额) = Format(rsTmp!实收金额, gstrDec)
       Bill.TextMatrix(i, BillCol.内部条码) = Nvl(rsTmp!内部条码)
       Bill.TextMatrix(i, BillCol.商品条码) = Nvl(rsTmp!商品条码)
       Bill.TextMatrix(i, BillCol.类型) = Nvl(rsTmp!费用类型)
        
        '设置销帐标志
        If Bill.TextMatrix(0, Bill.Cols - 1) = "删除" Then
            Bill.TextMatrix(i, Bill.Cols - 1) = "√"
        End If
        
        rsTmp.MoveNext
    Next
    '针对列编辑性质设置颜色
    Bill.SetColColor BillCol.项目, &HE7CFBA
    Bill.SetColColor BillCol.数次, &HE7CFBA
    Bill.SetColColor BillCol.付数, &HE0E0E0
    Bill.SetColColor BillCol.单价, &HE0E0E0
    Call SetColNum
    Bill.Redraw = True
    
    '----------------------------------------------------------------------------
    If blnDelete Then
        '读取准退数,并计算应收金额,实收金额(金额=剩余金额*(准退数/剩余数))

        '读取药品收发记录中的准退数
        strSQL1 = _
            " Select A.费用ID,Sum(Nvl(A.付数,1)*A.实际数量) as 准退数量" & _
            " From 药品收发记录 A " & _
            " Where    A.NO=[1] And MOD(A.记录状态,3)=1" & _
            "           And A.审核人 is NULL And Instr([3],','||A.单据||',')>0" & _
            " Group by A.费用ID"
        
        '整张费用单据(明细到收入项目)
        '执行状态应该在原始记录上判断(部分退药且部分退费的记录)
        strSQL = "" & _
            "   Select Nvl(价格父号,序号) From " & mstrFeeTab & _
            "   Where 记录性质=[2] And 记录状态 IN(0,1,3) And NO=[1]" & _
            "       And Nvl(执行状态,0)<>1" & IIF(mlng医嘱ID <> 0, " And 医嘱序号+0=[7]", "")
        If blnDo Then
            strSQL = strSQL & " And Nvl(价格父号,序号) IN" & _
                " (" & _
                " Select Nvl(价格父号,序号) as 序号" & _
                " From " & mstrFeeTab & _
                " Where NO=[1] And 记录性质 IN(2,12)" & _
                " Group by Nvl(价格父号,序号)" & _
                " Having Sum(Nvl(结帐金额,0))=0" & _
                " )"
        End If
        
        strSQL = _
            "   Select Sum(A.ID) as ID,A.序号,A.名称,A.收费类别," & _
            "       Sum(A.数量) as 剩余数量,Sum(A.应收金额) as 剩余应收," & _
            "       Sum(A.实收金额) as 剩余实收 " & _
            "   From (  Select Decode(A.记录状态,2,0,A.ID) as ID,A.序号,B.名称,A.收费类别," & _
            "                       Nvl(A.付数,1)*A.数次  as 数量, A.应收金额,A.实收金额" & _
            "               From " & mstrFeeTab & " A,收入项目 B " & _
            "               Where A.记录性质=[2] And A.NO=[1]" & _
            "                           And A.收入项目ID=B.ID And Nvl(A.价格父号,A.序号) IN(" & strSQL & ")" & _
            "              ) A" & _
            " Group by A.序号,A.名称,A.收费类别" & _
            " Having Sum(A.数量)<>0"
                    
        '最后计算结果
        strSQL = _
            "   Select A.名称,Sum(A.剩余应收*(A.准退数量/A.剩余数量)) as 应收金额," & _
            "       Sum(剩余实收*(A.准退数量/A.剩余数量)) as 实收金额  " & _
            "   From ( Select A.名称,A.剩余数量,A.剩余应收,A.剩余实收," & _
            "                   Decode(Instr(',4,5,6,7,',A.收费类别),0,A.剩余数量,Nvl(B.准退数量,A.剩余数量)) as 准退数量" & _
            "               From (" & strSQL & ") A,(" & strSQL1 & ") B" & _
            "               Where A.ID=B.费用ID(+)" & _
            "              ) A  " & _
            "   Group by A.名称"
    Else
        '读取单据原始内容
        intSign = IIF(mblnDelete, -1, 1) '数量,金额正负符号
        
        strSQL = "Select A.收入项目ID,A.应收金额,A.实收金额 From " & mstrFeeTab & " A" & _
            " Where Instr([4],A.记录状态)>0 And A.记录性质=[2] And A.NO=[1]" & _
            IIF(mstrTime <> "", " And A.登记时间=[5]", "")
        If blnNOMoved Then
            strSQL = strSQL & " Union ALL " & Replace(strSQL, mstrFeeTab, "H" & mstrFeeTab)
        End If
        
        strSQL = _
            " Select B.名称,Sum([6]*A.应收金额) as 应收金额,Sum([6]*A.实收金额) as 实收金额 " & _
            " From (" & strSQL & ") A,收入项目 B Where A.收入项目ID=B.ID Group By B.名称"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, mint记录性质, IIF(mint记录性质 = 2, ",9,25,", ",8,24,"), _
        IIF(mblnDelete, "2", "0,1,3"), CDate(IIF(mstrTime = "", "1990-01-01", mstrTime)), intSign, mlng医嘱ID)
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
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetShowCol()
    '功能：付数列的控制(浏览时展开)
    Bill.ColWidth(BillCol.付数) = 0
End Sub

Private Sub ClearRows()
    Dim i As Long
    For i = 1 To Bill.Rows - 1
        Bill.RowData(i) = 0
    Next
End Sub
 
Private Function GetWorkUnit(ByVal lng药品ID As Long, ByVal str类别 As String) As Boolean
'功能：取所有可供选择的药房
    Dim strSQL As String, bytDay As Byte
    Dim str药房 As String, lng开单科室ID As Long
    
    lng开单科室ID = mrsInfo!科室ID    '开单科室优先
    If lng开单科室ID = 0 And cbo开单科室.ListIndex <> -1 Then lng开单科室ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)

    strSQL = _
    " Select Distinct C.ID,C.编码,C.简码,C.名称,B.工作性质,B.服务对象 " & _
    " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
    " Where A.执行科室ID+0=B.部门ID And B.工作性质='发料部门'" & _
    "       And B.服务对象 IN([1],3) And B.部门ID=C.ID" & _
    "       And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
    "       And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
    "       And (A.病人来源 is NULL Or A.病人来源=[1])" & _
    "       And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
    "       And A.收费细目ID=[3]" & _
    " Order by B.服务对象,C.编码"
    On Error GoTo errH
    Set mrsWork = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mint病人来源, lng开单科室ID, lng药品ID, str药房, bytDay)
    GetWorkUnit = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Load开单人(ByVal lng科室id As Long)
'功能：读取当前开单科室中包含的所有人员
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim lngOldID As Long
    
    cbo开单人.Clear
    
    '科室医生或护士
    strSQL = _
    "Select Distinct A.ID,B.部门ID,A.编号,A.姓名,Upper(A.简码) as 简码," & _
    "           C.人员性质,Nvl(A.聘任技术职务,0) as 职务" & _
    " From 人员表 A,部门人员 B,人员性质说明 C" & _
    " Where A.ID=B.人员ID And A.ID=C.人员ID" & _
    "       And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
    "       And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
    "       And C.人员性质 IN('医生','护士') And B.部门ID=[1]  " & _
    "  Order by 简码,人员性质 Desc"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng科室id)
    
    i = IIF(rsTmp.RecordCount = 0, 0, rsTmp.RecordCount - 1)
    ReDim marrDr(i)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            If lngOldID <> rsTmp!ID Then
                cbo开单人.AddItem IIF(IsNull(rsTmp!简码), "", rsTmp!简码 & "-") & rsTmp!姓名
                cbo开单人.ItemData(cbo开单人.ListCount - 1) = rsTmp!部门ID
                marrDr(cbo开单人.ListCount - 1) = rsTmp!ID & "|" & rsTmp!部门ID & "|" & Nvl(rsTmp!编号) & "|" & rsTmp!姓名 & "|" & Nvl(rsTmp!简码) & "|" & rsTmp!职务 & "|" & Nvl(rsTmp!人员性质)
                
                If rsTmp!姓名 = mstr开嘱医生 Then cbo开单人.ListIndex = cbo开单人.NewIndex
                If rsTmp!ID = UserInfo.ID And cbo开单人.ListIndex = -1 Then cbo开单人.ListIndex = cbo开单人.NewIndex
                lngOldID = rsTmp!ID
            End If
            rsTmp.MoveNext
        Next
        
        If cbo开单人.ListCount > 0 Then ReDim Preserve marrDr(cbo开单人.ListCount - 1)
        
        If cbo开单人.ListCount = 1 And cbo开单人.ListIndex = -1 Then cbo开单人.ListIndex = 0
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
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
            
            Bill.ColWidth(BillCol.项目) = GetOrigColWidth(BillCol.项目) - 100
            Bill.ColWidth(BillCol.单价) = GetOrigColWidth(BillCol.单价) - 50
            Bill.ColWidth(BillCol.应收金额) = GetOrigColWidth(BillCol.应收金额) - 50
            Bill.ColWidth(BillCol.实收金额) = GetOrigColWidth(BillCol.实收金额) - 50
            Bill.Redraw = True
        End If
    Else
        If Bill.TextMatrix(0, Bill.Cols - 1) = "删除" Then
            Bill.Redraw = False
            Bill.Cols = Bill.Cols - 1
            Bill.ColWidth(BillCol.项目) = GetOrigColWidth(BillCol.项目)
           Bill.ColWidth(BillCol.单价) = GetOrigColWidth(BillCol.单价)
            Bill.ColWidth(BillCol.应收金额) = GetOrigColWidth(BillCol.应收金额)
            Bill.ColWidth(BillCol.实收金额) = GetOrigColWidth(BillCol.实收金额)
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
        Bill.TextMatrix(i, BillCol.行) = i
    Next
    Bill.Redraw = True
End Sub
 

Private Function PhysicExist(objDetail As Detail, intRow As Integer) As Boolean
'功能：判断指定材料在单据中是否已经存在
'参数：objDetail=项目,intRow=要判断的行
'说明：时价或分批药品在同一药房禁止重复输入(这里仅提示,保存时禁止)
    Dim i As Integer
    
    For i = 1 To mobjBill.Details.Count
        If i <> intRow Then
            If mobjBill.Details(i).Detail.ID = objDetail.ID Then
                If (mobjBill.Details(i).Detail.分批 Or mobjBill.Details(i).Detail.变价) _
                    And (objDetail.分批 Or objDetail.变价) Then
                    If MsgBox("卫生材料""" & objDetail.名称 & """在第 " & i & " 行已经输入,要继续吗？" & _
                        vbCrLf & vbCrLf & "注意：该卫生材料为分批或时价药品,重复输入时必须保证它们的发料部门不同。", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        PhysicExist = True
                    End If
                    Exit Function
                Else
                    If MsgBox("卫生材料""" & objDetail.名称 & """在第 " & i & " 行已经输入,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        PhysicExist = True
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
    Dim bln医保 As Boolean, bln公费 As Boolean
    
    Check费用类型 = True
    
    '无法检查
    If txt付款方式.Tag = "" Then Exit Function
    
    '45605
    '只检查医保病人和公费病人
    If zlIsCheckMedicinePayMode(txt付款方式.Text, bln医保, bln公费) = False Then Exit Function
    '确定病人类型
    bytType = IIF(bln医保, 1, 2) ' Val(txt付款方式.Tag)
    
    '读取检查数据
    If bytType = 1 Then
        strSQL = "Select 编码,名称,简码,性质,缺省标志 From 费用类型 Where 编码 In(" & gstr医保费用类型 & ") Order by 编码"
    Else
        strSQL = "Select 编码,名称,简码,性质,缺省标志 From 费用类型 Where 编码 In(" & gstr公费费用类型 & ") Order by 编码"
    End If

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
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
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ReCalcInsure()
'功能：修改单据时,重新计算统筹金额及更新相关信息
    Dim i As Long, j As Long, dblAllTime As Double
    Dim strInfo As String
    
    If Not IsNull(mrsInfo!险类) Then
        For i = 1 To mobjBill.Details.Count
            For j = 1 To mobjBill.Details(i).InComes.Count
                dblAllTime = mobjBill.Details(i).付数 * mobjBill.Details(i).数次
                strInfo = gclsInsure.GetItemInsure(mobjBill.病人ID, mobjBill.Details(i).收费细目ID, mobjBill.Details(i).InComes(j).实收金额, False, mrsInfo!险类, _
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


Public Sub InitLocPar()
'功能：初始化费用本机参数
    mstrLike = IIF(Val(zlDatabase.GetPara("输入匹配")) = 0, "%", "")
    mblnTime = Val(zlDatabase.GetPara("变价输入数次", glngSys, p医嘱附费管理)) <> 0
    mbytSendMateria = Val(zlDatabase.GetPara("记帐后发药", glngSys, p医嘱附费管理))
    'mlng发料部门 = Val(zldatabase.GetPara(IIF(mint病人来源 = 2, "住院", "门诊") & "缺省发料部门", glngSys, p医嘱附费管理))
End Sub

Public Function zlCheck北京医保(ByVal intInsure As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:对北京医保的一些检查
    '入参:intInsuer-险类
    '出参:
    '返回:检查成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-07 10:25:04
    '问题:27278
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String
    If intInsure = 0 Then zlCheck北京医保 = True: Exit Function
    
    err = 0: On Error GoTo Errhand:
    '刘兴洪:???
    'mint病人来源:1-门诊病人,2-住院病人
    'mint记录性质:1-收费(划价),2-记帐(门/住)
    'mbytInState :0-执行,1-查阅,2-调整(不支持),3-删费
    
    '只有划价才支持检查
    If (mint病人来源 = 2 Or mint记录性质 = 2) And mbytInState <> 0 Or MCPAR.医生确定处方类型 = False Then
        zlCheck北京医保 = True: Exit Function
    End If
    
    'showmsgbox
    '参数：strCaption=消息窗体标题
    '      strInfo=具体提示内容,可用"^"表示换行,">"表示缩进。
    '      strCmds=按钮描述,如"重试(&R),!忽略(&A),?取消(&C)"
    '              至少要有两个按钮,"!"表示缺省按钮,"?"表示取消按钮
    '              每个按钮文字最多支持4个汉字
    '      vStyle=vbInformation,vbQuestion,vbExclamation,vbCritical
    '返回：按钮文字,如"按钮2"(不包含()和&),如果按关闭或取消则返回""
    strTemp = zlCommFun.ShowMsgBox("处方类型", "请确定当前医保病人本次要发送的药品处方的类型。", "!医保内(&A),医保外(&B),?取消(&C)", Me)
    If strTemp = "" Then Exit Function
    '如果是补门诊收费划价单，且是医保病人，则当医保参数”support医生确定处方类型”有效时，保存时提示该单据是”医保内，医保外”，如果是医保内费用记录摘要中存放1，医保外存放2。
    strTemp = IIF(strTemp = "医保内", 1, 2)
    For Each mobjBillDetail In mobjBill.Details
        mobjBillDetail.摘要 = strTemp
    Next
    zlCheck北京医保 = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function

Public Function zlCheck定价零价格对码(ByVal lng收费细目ID As Long, bln定价 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查医保对码(定价为零的)
    '入参:
    '出参:
    '返回:如果定价项目为零,则返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-12 11:22:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
   '刘兴洪 问题:27286 定价的价格为零的不进行检查对码 日期:2010-01-07 15:13:45
   Dim strSQL As String, rs价格 As ADODB.Recordset, dbl价格 As Double
    err = 0: On Error GoTo Errhand:
   zlCheck定价零价格对码 = False
    If bln定价 Then
        strSQL = _
        " Select  B.现价 " & _
        " From 收费价目 B " & _
        " Where   ((Sysdate Between B.执行日期 and B.终止日期) Or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
        "       And B.收费细目ID=[1]"
        Set rs价格 = zlDatabase.OpenSQLRecord(strSQL, "获取当前价格", lng收费细目ID)
        If rs价格.EOF = False Then
            dbl价格 = Val(Nvl(rs价格!现价))
        Else
            dbl价格 = 0
        End If
        If dbl价格 = 0 Then zlCheck定价零价格对码 = True: Exit Function
    End If
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
 
Public Function Get待发药清单(strNO As String, strTime As String) As ADODB.Recordset
    '------------------------------------------------------------------------------------------------------------------------
    '功能：根据费用单据号,登记时间,获取待发药品清单
    '入参：strNO-单据号
    '          strTime-登记时间
    '出参：
    '返回：待发药清单
    '编制：刘兴洪
    '日期：2010-03-19 18:59:27
    '说明：普通发药时为病人科室，急诊、医技则为开单科室。
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSQL = _
        " Select A.ID,A.库房ID,A.对方部门ID" & _
        " From 药品收发记录 A,住院费用记录 B" & _
        " Where A.NO=[1] And A.单据=[2] And Mod(A.记录状态,3)=1 And A.审核人 is NULL" & _
        " And A.NO=B.NO And A.费用ID=B.ID And B.记录状态<>0 And B.登记时间+0=[3]" & _
        " Order by A.药品ID"
    If strTime <> "" Then
        Set Get待发药清单 = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, 9, CDate(strTime))
    Else
        Set Get待发药清单 = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, 9)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
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
    '卫材
    dblStock = GetStock(objDetail.ID, lng执行科室ID)
    objDetail.库存 = dblStock
End Sub

Private Function GetInputDetail(ByVal lng项目id As Long) As Detail
    '功能：读取收费项目信息
    Dim objDetail As New Detail
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long, lngMediCareNO As Long
        
    If mrsInfo.State = 1 Then lngMediCareNO = Val("" & mrsInfo!险类)
    
    If lngMediCareNO > 0 Then
        strSQL = _
        " Select A.ID,A.类别,B.名称 as 类别名称,A.编码,Nvl(E.名称,A.名称) as 名称,E1.名称 as 商品名,A.规格,A.计算单位," & _
        "       A.屏蔽费别,A.是否变价,A.加班加价,A.执行科室,A.服务对象,A.费用类型,A.补充摘要,M.要求审批," & _
        "       D.诊疗ID as 药名ID, D.在用分批  as 分批, 1 as 药房包装, A.计算单位  as 药房单位,D.跟踪在用,A.录入限量" & _
        " From 收费项目目录 A,收费项目类别 B,材料特性 D,收费项目别名 E,收费项目别名 E1,保险支付项目 M" & _
        " Where   A.ID=D.材料ID(+) And B.编码=A.类别" & _
        "       And A.ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=[2] " & _
        "       And A.ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3 " & _
        "       And A.ID=[1] And A.ID=M.收费细目ID(+) And M.险类(+)=[3]"
    Else
        strSQL = _
        " Select A.ID,A.类别,B.名称 as 类别名称,A.编码,Nvl(E.名称,A.名称) as 名称,E1.名称 as 商品名,A.规格,A.计算单位," & _
        "       A.屏蔽费别,A.是否变价,A.加班加价,A.执行科室,A.服务对象,A.费用类型,A.补充摘要,0 as 要求审批, D.诊疗ID as 药名ID," & _
        "       D.在用分批 as 分批, 1 as 药房包装, A.计算单位 as 药房单位,D.跟踪在用,A.录入限量" & _
        " From 收费项目目录 A,收费项目类别 B,材料特性 D,收费项目别名 E,收费项目别名 E1" & _
        " Where  A.ID=D.材料ID(+) And B.编码=A.类别" & _
        "       And A.ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=[2] " & _
        "       And A.ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3 " & _
        "       And A.ID=[1]"
    End If
        
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng项目id, IIF(gbyt药品名称显示 = 1, 3, 1), lngMediCareNO)
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
        .中药形态 = 0
        .商品名 = Nvl(rsTmp!商品名)
        
    End With
    Set GetInputDetail = objDetail
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SelectItem(ByVal blnInput As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:选择指定的卫生材料项目
    '入参:strKey-输入选项
    '       blnInput-输入
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2010-12-14 14:32:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngTXTHwnd As Long, strInput As String
    Dim str特准项目 As String, int险类 As Integer, lng项目id As Long
    Dim rsItem As ADODB.Recordset
    
    On Error GoTo errHandle
    If blnInput Then
        lngTXTHwnd = Bill.TxtHwnd
        strInput = Bill.Text
    End If
    If Not IsNull(mrsInfo!险类) Then
        int险类 = mrsInfo!险类
        '刘兴洪:24862
        'mint病人来源 As Integer '1-门诊病人,2-住院病人
        'mint记录性质 As Integer '1-收费(划价),2-记帐(门/住)
        If zl_Check特准项目(gclsInsure, int险类, Val(Nvl(mrsInfo!病人ID)), (mint记录性质 = 1 Or mint病人来源 = 1)) Then str特准项目 = Get保险特准项目(Val(Nvl(mrsInfo!病人ID)), "A.ID")
    End If
    
    If frmStuffSelect.ShowSelect(Me, mstrPrivs, mint病人来源, int险类, strInput, lngTXTHwnd, str特准项目, mlng虚拟库房ID, False, rsItem) = False Then GoTo GoNotSel
    If rsItem Is Nothing Then GoTo GoNotSel:
    If rsItem.State <> 1 Then GoTo GoNotSel:
    If rsItem.RecordCount = 0 Then GoTo GoNotSel:
    lng项目id = Val(Nvl(rsItem!收费项目ID))
    Set mobjDetail = GetInputDetail(lng项目id)
    If int险类 <> 0 Then sta.Panels(4).Text = Get医保大类(lng项目id, int险类)
    mobjDetail.批次 = Val(Nvl(rsItem!批次))
    mobjDetail.商品条码 = Trim(Nvl(rsItem!商品条码))
    mobjDetail.内部条码 = Trim(Nvl(rsItem!内部条码))
    mobjDetail.库存 = Val(Nvl(rsItem!可用库存))
    SelectItem = True
GoNotSel:
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function Set虚拟库房ID(ByVal lng执行科室 As Long) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据执行科室,确定虚拟库房ID
    '返回:虚拟库房ID
    '编制:刘兴洪
    '日期:2010-12-15 10:06:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    strSQL = "Select 虚拟库房id  From 虚拟库房对照 Where 科室id = [1] And Rownum <= 1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng执行科室)
    If Not rsTemp.EOF Then
        Set虚拟库房ID = Val(Nvl(rsTemp!虚拟库房id))
    Else
        Set虚拟库房ID = 0
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function
