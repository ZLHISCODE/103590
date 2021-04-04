VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReplenishTheBalanceDelWin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "医保补结算退费信息"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9405
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReplenishTheBalanceDelWin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame3 
      Height          =   60
      Left            =   -30
      TabIndex        =   29
      Top             =   1230
      Width           =   10260
   End
   Begin VB.Frame Frame2 
      Height          =   60
      Left            =   -690
      TabIndex        =   25
      Top             =   5085
      Width           =   10635
   End
   Begin VB.TextBox txtAge 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   7905
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   150
      Width           =   1185
   End
   Begin VB.TextBox txtSex 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   5655
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   150
      Width           =   1185
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7755
      TabIndex        =   20
      Top             =   5265
      Width           =   1470
   End
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   -30
      TabIndex        =   24
      Top             =   660
      Width           =   10260
   End
   Begin VB.TextBox txtPatiName 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   1365
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   150
      Width           =   2895
   End
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   3675
      Left            =   75
      ScaleHeight     =   3675
      ScaleWidth      =   9300
      TabIndex        =   22
      Top             =   1290
      Width           =   9300
      Begin VB.PictureBox PicBalanceBack 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         ForeColor       =   &H80000008&
         Height          =   3150
         Left            =   45
         ScaleHeight     =   3120
         ScaleWidth      =   4020
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   480
         Width           =   4050
         Begin VSFlex8Ctl.VSFlexGrid vsBalance 
            Height          =   2430
            Left            =   90
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   90
            Width           =   3810
            _cx             =   6720
            _cy             =   4286
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   14.25
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
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   2
            HighLight       =   2
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   4
            Cols            =   2
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   460
            RowHeightMax    =   500
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmReplenishTheBalanceDelWin.frx":6852
            ScrollTrack     =   0   'False
            ScrollBars      =   2
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
         Begin VB.Label lblYBMoney 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1350
            TabIndex        =   30
            Top             =   2640
            Width           =   2565
         End
         Begin VB.Label lbl医保合计 
            AutoSize        =   -1  'True
            Caption         =   "医保合计"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   150
            TabIndex        =   8
            Top             =   2685
            Width           =   1140
         End
      End
      Begin VB.PictureBox picPay 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         ForeColor       =   &H80000008&
         Height          =   3150
         Left            =   4290
         ScaleHeight     =   3120
         ScaleWidth      =   4875
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   480
         Width           =   4905
         Begin VB.TextBox txt预存款 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            IMEMode         =   3  'DISABLE
            Left            =   1350
            MaxLength       =   12
            TabIndex        =   11
            Top             =   2430
            Visible         =   0   'False
            Width           =   3210
         End
         Begin VB.TextBox txt摘要 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1140
            Left            =   1350
            MaxLength       =   50
            MultiLine       =   -1  'True
            TabIndex        =   19
            Top             =   1845
            Width           =   3210
         End
         Begin VB.TextBox txt缴款 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            IMEMode         =   3  'DISABLE
            Left            =   1350
            MaxLength       =   12
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   675
            Width           =   3210
         End
         Begin VB.TextBox txt结算号码 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            IMEMode         =   3  'DISABLE
            Left            =   1350
            MaxLength       =   30
            TabIndex        =   17
            Top             =   1260
            Width           =   3210
         End
         Begin VB.ComboBox cbo支付方式 
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   405
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   120
            Width           =   3210
         End
         Begin VB.Label lbl预存款 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "退预存款"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   90
            TabIndex        =   10
            Top             =   2513
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.Label lbl摘要 
            AutoSize        =   -1  'True
            Caption         =   "摘  要"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   420
            TabIndex        =   18
            Top             =   1815
            Width           =   870
         End
         Begin VB.Label lbl退款金额 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "退款金额"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   90
            TabIndex        =   14
            Top             =   750
            Width           =   1200
         End
         Begin VB.Label lbl结算号码 
            AutoSize        =   -1  'True
            Caption         =   "结算号码"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   150
            TabIndex        =   16
            Top             =   1335
            Width           =   1140
         End
         Begin VB.Label lblPayType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "退款方式"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   90
            TabIndex        =   12
            Top             =   195
            Width           =   1200
         End
      End
      Begin XtremeSuiteControls.ShortcutCaption stcBalanceTittle 
         Height          =   405
         Left            =   4275
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   30
         Width           =   4920
         _Version        =   589884
         _ExtentX        =   8678
         _ExtentY        =   714
         _StockProps     =   6
         Caption         =   "当前收退信息"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         GradientColorLight=   16711680
         GradientColorDark=   16711680
      End
      Begin XtremeSuiteControls.ShortcutCaption stcYbTittle 
         Height          =   405
         Left            =   45
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   45
         Width           =   4065
         _Version        =   589884
         _ExtentX        =   7170
         _ExtentY        =   714
         _StockProps     =   6
         Caption         =   "当前退费信息"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   510
      Left            =   0
      TabIndex        =   28
      Top             =   5850
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   900
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2619
            MinWidth        =   882
            Picture         =   "frmReplenishTheBalanceDelWin.frx":68A0
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8361
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1138
            MinWidth        =   1146
            Object.Tag             =   "用于收费预交余额显示"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1164
            MinWidth        =   1162
            Object.Tag             =   "用于收费三方卡余额的显示"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   882
            MinWidth        =   882
            Picture         =   "frmReplenishTheBalanceDelWin.frx":7134
            Key             =   "Calc"
            Object.ToolTipText     =   "计算器:ALT+?"
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
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出(&E)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   6195
      TabIndex        =   21
      Top             =   5265
      Width           =   1470
   End
   Begin VB.Label lbl退费合计 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1365
      TabIndex        =   32
      Top             =   810
      Width           =   2655
   End
   Begin VB.Label lbl合计 
      Caption         =   "退费合计"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   180
      TabIndex        =   31
      Top             =   855
      Width           =   1230
   End
   Begin VB.Label lbl误差 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "本次误差:0.00"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5265
      TabIndex        =   27
      Top             =   5415
      Width           =   2025
   End
   Begin VB.Label lblAge 
      AutoSize        =   -1  'True
      Caption         =   "年龄"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   7260
      TabIndex        =   4
      Top             =   210
      Width           =   570
   End
   Begin VB.Label lblSex 
      AutoSize        =   -1  'True
      Caption         =   "性别"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5010
      TabIndex        =   2
      Top             =   210
      Width           =   570
   End
   Begin VB.Label lblPati 
      AutoSize        =   -1  'True
      Caption         =   "病人姓名"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   1140
   End
End
Attribute VB_Name = "frmReplenishTheBalanceDelWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------------
'程序入口相关变量
Public Enum gEM_BalanceDel
    EM_BalanceDel = 0   '结算退费
    EM_BalanceReDel = 1  '重新退费
End Enum
Private mbytFunc As gEM_BalanceDel
Private mobjDelBalance As clsCliniDelBalance
Private mfrmMain As Object
Private mlngModule As Long, mstrPrivs As String
Private mbln医保分币 As Boolean
Private mcllForceDelToCash As Collection '强制退现信息：Array(操作员,卡类别名称,结算方式)
Private mstrDefaultBalance As String
Private mstr排除结算方式 As String '不能使用的结算方式,多个用逗号分隔
Private mblnRegister As Boolean
'------------------------------------------------------------------------------------------
'局部变量
Private mobjPayCards As Cards
Private mstrTittle As String
Private mblnFirst As Boolean
Private mblnUnLoad As Boolean '是否Unload窗体
Private mdblDelMoney As Double '本次退款金额
Private mdbl当前未退 As Double
Private mdbl可退预交 As Double
Private mblnOK As Boolean, mblnNotClick As Boolean
Private mbln已报价 As Boolean
Private mlngR As Long
Private Type TY_BrushCard    '刷卡类型
    str卡号 As String
    str密码 As String
    str交易流水号 As String    '交易流水号
    str交易说明  As String     '交易信息
    str扩展信息 As String    '交易的扩展信息
End Type
Private mCurBrushCard As TY_BrushCard   '当前的刷卡信息
Private Enum Pan
    C2提示信息 = 2
    C3个人帐户 = 3
    C4三方帐户信息 = 4
End Enum

'------------------------------------------------------------------------------------------
'API声明
Private Const VK_RETURN = &HD
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private mblnCacheKeyReturn As Boolean   '41025:是否缓存了回车键,可能存在在收费界面刷卡中本身包含了回车,因此需要判断
Private mlngPre支付方式 As Long
Private mrsOldBalance As ADODB.Recordset
Private mblnThreeSwapSingle As Boolean '是否单独调用退费接口

Public Function zlChargeWin(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal strPrivs As String, ByVal bytFunc As gEM_BalanceDel, _
    ByVal objPayCards As Cards, ByVal objDelBalance As clsCliniDelBalance, _
    ByVal bln医保分币 As Boolean, _
    Optional ByVal cllForceDelToCash As Collection, Optional ByVal str排除结算方式 As String, _
    Optional ByVal blnRegister As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入口:表示进入支付结算窗口
    '入参:frmMain-调用的主窗体
    '       lngModule -模块号
    '       strPrivs-权限串
    '       str结算序号:本次结算序号
    '       bln医保分币-医保是否分币处理
    '       dtDate-当前收费时间
    '      objPayCards-当前有效的支付类别
    '       cllForceDelToCash - 强制退现信息：Array(操作员,卡类别名称,结算方式)
    '       str排除结算方式 - 不能使用的结算方式,多个用逗号分隔
    '       blnRegister - 是否是挂号结算单据
    '返回:完成结算,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-09-18 14:33:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set mfrmMain = frmMain: mbytFunc = bytFunc
    mstrPrivs = strPrivs: mlngModule = lngModule
    mbln医保分币 = bln医保分币
    Set mobjDelBalance = objDelBalance
    Set mobjPayCards = objPayCards
    Call InitVar '初始化相关本地模块变量
    If cllForceDelToCash Is Nothing Then Set cllForceDelToCash = New Collection
    Set mcllForceDelToCash = cllForceDelToCash
    mstr排除结算方式 = str排除结算方式
    mblnRegister = blnRegister
    
    Me.Show 1, frmMain
    zlChargeWin = mblnOK
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub InitVar()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关本地模块变量
    '编制:刘兴洪
    '日期:2014-09-18 17:16:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mblnNotClick = False:  mblnUnLoad = False
    mblnOK = False
    mblnFirst = True
    mstrDefaultBalance = ""
End Sub

Private Sub InitBalanceGrid(Optional ByVal blnInitColHead As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化保险结算表格
    '入参:blnInitColHead-案卷始化列头
    '编制:刘兴洪
    '日期:2014-09-18 14:06:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    With vsBalance
        .Clear 1
        .Rows = 4
        If blnInitColHead Then
            .COLS = 2
            .TextMatrix(0, 0) = "结算方式"
            .TextMatrix(0, 1) = "支付金额"
            For i = 0 To .COLS - 1
                .ColKey(i) = .TextMatrix(0, i)
                .FixedAlignment(i) = flexAlignCenterCenter
                If .ColKey(i) Like "*金额" Then
                    .ColAlignment(i) = flexAlignRightCenter
                Else
                    .ColAlignment(i) = flexAlignLeftCenter
                End If
            Next
            .ColWidth(.ColIndex("结算方式")) = (vsBalance.Width - 300) * 0.6
            .ColWidth(.ColIndex("支付金额")) = (vsBalance.Width - 300) * 0.4
            .Row = 0: .Col = 1
        End If
        .TabStop = False
        .Cell(flexcpFontBold, 0, 0, .Rows - 1, .COLS - 1) = False
        .Cell(flexcpForeColor, 0, 0, .Rows - 1, .COLS - 1) = Me.ForeColor
    End With
End Sub

Private Function LoadData(ByVal str结算序号 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载结算数据
    '入参:str结算序号-结算序号
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-09-18 14:22:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim i As Integer, dblYbMoney As Double
    Dim dblDelMoney As Double '费用退款金额
    Dim dblDelAllMoney As Double '退费合计
    
    mdblDelMoney = 0
    mdbl当前未退 = 0
    
    On Error GoTo errHandle
    
    Call InitBalanceGrid
    strSQL = "" & _
    "   Select decode(B.名称,null ,0,1) as 医保,A.结算方式,sum(A.冲预交) as 冲预交  " & _
    "   From 病人预交记录 A,(select 名称 From 结算方式 where 性质 in (3,4)) B" & _
    "   Where  A.结算序号 = [1] and a.结算方式=b.名称(+)" & _
    "   Group by decode(B.名称,null ,0,1),A.结算方式" & _
    "   Order by 医保 Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str结算序号)
    vsBalance.Appearance = flexFlat
    With rsTemp
        i = 1
        Do While Not .EOF
            If Nvl(rsTemp!结算方式) <> "" Then
                With vsBalance
                    If .TextMatrix(i, .ColIndex("结算方式")) <> "" Then
                        .Rows = .Rows + 1
                        i = i + 1
                    End If
                    .RowData(i) = 0
                    .TextMatrix(i, .ColIndex("结算方式")) = Nvl(rsTemp!结算方式)
                    .TextMatrix(i, .ColIndex("支付金额")) = Format(-1 * Val(Nvl(rsTemp!冲预交)), "0.00")
                End With
                
                If Val(Nvl(rsTemp!医保)) = 1 Then
                    dblYbMoney = dblYbMoney + Val(Nvl(rsTemp!冲预交))
                End If
            Else
                dblDelMoney = dblDelMoney + Val(Nvl(rsTemp!冲预交))
            End If
            dblDelAllMoney = dblDelAllMoney + Val(Nvl(rsTemp!冲预交))
            .MoveNext
        Loop
    End With
    lblYBMoney.Caption = Format(-1 * dblYbMoney, "0.00")
    lbl退费合计.Caption = Format(-1 * dblDelAllMoney, "0.00")
    '计算本次实际退款
    mdblDelMoney = RoundEx(dblDelMoney, 6)
    txt缴款.Text = Format(-1 * mdblDelMoney, "0.00")
    LoadData = True
   Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub InitFace()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化控件
    '编制:刘兴洪
    '日期:2011-06-13 14:09:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim CurBrushCard As TY_BrushCard
    
    zlControl.PicShowFlat PicBalanceBack, -1, , taCenterAlign
    zlControl.PicShowFlat picPay, -1, , taCenterAlign
    
    Call InitBalanceGrid(True)
    If mblnUnLoad = False Then
        mblnUnLoad = Not LoadData(mobjDelBalance.结算序号)
        Set mrsOldBalance = zlFromIDGetChargeBalance(2, mobjDelBalance.AllNos, , , , IIf(mblnRegister, 4, 1))
        '获取收费时使用的结算方式作为缺省结算方式
        mrsOldBalance.Filter = "退费=0"
        If mrsOldBalance.EOF = False Then
            If Val(mrsOldBalance!类型) = 1 Then
                mstrDefaultBalance = "预存款"
            Else
                mstrDefaultBalance = Nvl(mrsOldBalance!结算方式)
            End If
        End If
        mrsOldBalance.Filter = ""
    End If
    mdbl当前未退 = mdblDelMoney
    Call Load支付方式
    txt缴款.Text = Format(-1 * mdbl当前未退, "0.00")
    If mdbl当前未退 <= 0 Then
        lblPayType.Caption = "退款方式"
        lbl退款金额.Caption = "退款金额"
        txt缴款.ForeColor = lblPati.ForeColor
    Else
        lblPayType.Caption = "收款方式"
        lbl退款金额.Caption = "收款金额"
        txt缴款.ForeColor = vbRed
    End If
    
    txt缴款.Locked = True
    txt缴款.BackColor = &HE0E0E0

    mCurBrushCard = CurBrushCard
    stbThis.Panels(C4三方帐户信息).Text = "": stbThis.Panels(C4三方帐户信息).ToolTipText = ""
    stbThis.Panels(C4三方帐户信息).Visible = False
    vsBalance.BackColor = Me.BackColor
    vsBalance.BackColorBkg = Me.BackColor
    txtPatiName.Text = mobjDelBalance.姓名
    txtPatiName.ForeColor = vbRed
    If mobjDelBalance.病人类型 <> "" Then
         Call SetPatiColor(txtPatiName, mobjDelBalance.病人类型, vbRed)
    End If
    txtAge.Text = mobjDelBalance.年龄
    txtAge.ForeColor = txtPatiName.ForeColor
    txtSex.Text = mobjDelBalance.性别
    txtSex.ForeColor = txtPatiName.ForeColor
End Sub
Private Sub Load支付方式()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载有效的支付方式
    '编制:刘兴洪
    '日期:2011-07-08 11:41:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, blnFind As Boolean
    Dim objCard As Card, objCards As Cards, objPayCards As Cards
    Dim lngKey As Long, i As Long, j As Long
    Dim varData As Variant
    
    On Error GoTo errHandle
    
'    If mobjPayCards Is Nothing Then
        Set objCards = New Cards: Set mobjPayCards = New Cards
        Set rsTemp = Get结算方式("补结算")
        '83533:李南春,2015/3/25,没有有效的补结算
        If rsTemp.RecordCount = 0 Then
            MsgBox "补结算没有可用的结算方式，请先到『结算方式管理』中设置补结算的应用场合。", vbInformation, gstrSysName
            mblnUnLoad = True: Exit Sub
        End If
        If Not gobjSquare Is Nothing Then
            ' zlGetCards(ByVal BytType As Byte)
            '入参:bytType-0-所有医疗卡;
            '             1-启用的医疗卡,
            '             2-所有存在三方账户的三方卡
            '             3-启用的三方账户的医疗卡
           Set objCards = gobjSquare.objSquareCard.zlGetCards(3)
        End If
        
        With rsTemp
            .Filter = 0
            If .RecordCount <> 0 Then .MoveFirst
            lngKey = 1
            Do While Not .EOF
                For i = 1 To objCards.Count
                    If objCards(i).结算方式 = Nvl(rsTemp!名称) Then
                        blnFind = True
                        Exit For
                    End If
                Next
                If Not blnFind Then
                    '83266:李南春,2015/3/18,医疗卡还需判断是否启用
                    If InStr(",1,2,", "," & Val(Nvl(rsTemp!性质)) & ",") > 0 _
                        And Val(Nvl(rsTemp!应付款)) <> 1 Then
                        '不加入医保的结算方式或退支票的
                         Set objCard = New Card
                         objCard.短名 = Mid(Nvl(!名称), 1, 1)
                         objCard.接口编码 = Nvl(!编码)
                         objCard.接口程序名 = ""
                         objCard.接口序号 = -1 * lngKey
                         objCard.结算方式 = Nvl(!名称)
                         objCard.名称 = Nvl(!名称)
                         objCard.启用 = True
                         objCard.缺省标志 = Val(Nvl(rsTemp!缺省)) = 1
                         objCard.支付启用 = True
                         objCard.结算性质 = Val(!性质)
                        mobjPayCards.Add objCard, "K" & lngKey
                        lngKey = lngKey + 1
                    End If
                End If
                .MoveNext
            Loop
        End With
        '加三方卡
        For Each objCard In objCards
            If objCard.消费卡 = False Then
                rsTemp.Filter = "名称='" & objCard.结算方式 & "'"
                If Not rsTemp.EOF Then
                    mobjPayCards.Add objCard, "K" & lngKey
                    lngKey = lngKey + 1
                End If
            End If
        Next
        
        If Exist预交款() Then
            '强制加入预交金额
            Set objCard = New Card
            objCard.短名 = "预"
            objCard.接口编码 = ""
            objCard.接口程序名 = ""
            objCard.接口序号 = -1 * lngKey
            objCard.结算方式 = "预存款"
            objCard.名称 = "预存款"
            objCard.启用 = True
            objCard.缺省标志 = False
            objCard.支付启用 = True
            objCard.结算性质 = "-99"
            mobjPayCards.Add objCard, "K" & lngKey
        End If
        
        If mobjPayCards.Count = 0 Then
            MsgBox "结算卡设置有误,原因可能如下:" & vbCrLf & _
                "1)未正常启用结算卡,请到『医疗卡类别』和『设备配置』中启用" & vbCrLf & _
                "2)未设置结算卡的[轧帐及代扣]属性,请在『医疗卡类别』中设置", vbInformation, gstrSysName
            mblnUnLoad = True: Exit Sub
        End If
'    End If
    
    mblnNotClick = True
    mlngPre支付方式 = -1
    With cbo支付方式
        .Clear
        For i = 1 To mobjPayCards.Count
            Set objCard = mobjPayCards(i)
            blnFind = False
            If mstr排除结算方式 <> "" Then
                varData = Split(mstr排除结算方式, ",")
                For j = 0 To UBound(varData)
                    If objCard.结算方式 = varData(j) Then
                        blnFind = True: Exit For
                    End If
                Next
            End If
            If blnFind = False Then '排除的不加入
                If objCard.接口序号 <= 0 _
                    Or objCard.接口序号 > 0 And (mstrDefaultBalance = objCard.结算方式 _
                                            Or mcllForceDelToCash.Count = 0 And objCard.是否转帐及代扣) Then
                    .AddItem objCard.名称
                    .ItemData(.NewIndex) = i
                    
                    If objCard.缺省标志 And .ListIndex < 0 Then .ListIndex = .NewIndex: mlngPre支付方式 = i
                    If objCard.结算性质 = 1 And .ListIndex < 0 Then .ListIndex = .NewIndex: mlngPre支付方式 = i
                    If mstrDefaultBalance = objCard.结算方式 Then .ListIndex = .NewIndex: mlngPre支付方式 = i
                End If
            End If
        Next
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0: mlngPre支付方式 = i
        If .ListCount = 0 Then
            MsgBox "没有可用的结算方式，不能继续。请先到结算方式管理中设置一个性质为1或2的结算方式。" & vbNewLine & _
                vbNewLine & _
                "原因：收款时使用的结算方式可能已无可退回金额", vbExclamation, gstrSysName
            mblnUnLoad = True: Exit Sub
        End If
    End With
    mblnNotClick = False
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function Exist预交款() As Boolean
    '功能：是否还存在可退预交款
    Dim lngTop As Long, lngTop1 As Long
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandler
    mdbl可退预交 = 0
    txt预存款.Text = ""
    If RoundEx(mdblDelMoney, 6) >= 0 Then Exit Function '收款及退款金额为零时，不能使用预交款
    If mrsOldBalance Is Nothing Then Exit Function
    
    '2.收费预交金额
    '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
    mrsOldBalance.Filter = "类型=1"
    Do While Not mrsOldBalance.EOF
        mdbl可退预交 = mdbl可退预交 + Val(mrsOldBalance!冲预交)
        mrsOldBalance.MoveNext
    Loop
    If RoundEx(mdbl可退预交, 6) = 0 Then Exit Function
    
    '2.补结算已退预交金额
    '补充结算结帐ID
    strSQL = _
        "Select 结帐id" & vbNewLine & _
        "From 病人预交记录" & vbNewLine & _
        "Where 结算序号 In (Select a.结算序号" & vbNewLine & _
        "               From 费用补充记录 A, 费用补充记录 B" & vbNewLine & _
        "               Where a.No = b.No And a.记录性质 = b.记录性质 And a.附加标志 = b.附加标志" & vbNewLine & _
        "                     And b.记录性质 = 1 And b.结算序号 = [1])"
    '费用结帐ID
    strSQL = strSQL & vbNewLine & _
        "Minus" & vbNewLine & _
        "Select Distinct a.结帐id As 原结帐id" & vbNewLine & _
        "From 门诊费用记录 A, 门诊费用记录 B" & vbNewLine & _
        "Where a.记录性质 = b.记录性质 And a.No = b.No And a.序号 = b.序号 And b.记录状态 <> 2" & vbNewLine & _
        "      And b.结帐id In (Select 收费结帐id From 费用补充记录 Where 记录性质 = 1 And 结算序号 = [1])"
    strSQL = _
        "Select Nvl(冲预交, 0) As 冲预交" & vbNewLine & _
        "From 病人预交记录" & vbNewLine & _
        "Where Mod(记录性质, 10) = 1 And Nvl(预交类别, 0) = 1 And" & vbNewLine & _
        "      结帐id In (" & strSQL & ")"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjDelBalance.结算序号)
    Do While Not rsTemp.EOF
        mdbl可退预交 = mdbl可退预交 + Val(rsTemp!冲预交)
        rsTemp.MoveNext
    Loop
    If RoundEx(mdbl可退预交, 6) = 0 Then Exit Function
    
    mdbl可退预交 = mdbl可退预交
    '医保报销大于了费用总金额
    If RoundEx(mdbl可退预交, 6) < -1 * mdblDelMoney Then
        lbl预存款.Visible = True
        txt预存款.Visible = True
        txt预存款.Text = Format(mdbl可退预交, "0.00")
        
        lngTop = lbl预存款.Top: lngTop1 = txt预存款.Top
        lbl预存款.Top = lblPayType.Top: txt预存款.Top = cbo支付方式.Top
        lblPayType.Top = lbl退款金额.Top: cbo支付方式.Top = txt缴款.Top
        lbl退款金额.Top = lbl结算号码.Top: txt缴款.Top = txt结算号码.Top
        lbl结算号码.Top = lbl摘要.Top: txt结算号码.Top = txt摘要.Top
        lbl摘要.Top = lngTop: txt摘要.Top = lngTop1: txt摘要.Height = 600
        
        mdbl当前未退 = mdblDelMoney - (-1 * mdbl可退预交)
        Exit Function
    End If
    
    Exist预交款 = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查结算数据时的有效性
    '返回:数据有效,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-09-18 15:01:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim objCard As Card
    
    On Error GoTo errHandle
    '83222,冉俊明,2015-3-17,可用方式可能只有一卡通
'    If Val(txt缴款.Text) = 0 And cbo支付方式.ItemData(cbo支付方式.ListIndex) <> 1 Then
'        MsgBox "当前" & lbl退款金额.Caption & "为零，不能使用非现金结算方式！", vbOKOnly + vbInformation, gstrSysName
'        If cbo支付方式.Enabled And cbo支付方式.Visible Then cbo支付方式.ListIndex = 0
'        Exit Function
'    End If
    '并发检查
    If mbytFunc = EM_BalanceReDel Then
        If zlIsCheckExistErrBill(mobjDelBalance.结算序号, True) = False Then
            MsgBox "当前异常单据已被处理，你不能继续！", vbInformation, gstrSysName
            Exit Function
        End If
        If zlCheckOtherSessionDoing(mobjDelBalance.结算序号) Then
            MsgBox "当前单据正在其它补结算窗口中进行处理，你不能继续！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If Not CheckTextLength("结算号码", txt结算号码) Then Exit Function
    If Not CheckTextLength("摘要", txt摘要) Then Exit Function
    If IsValid预交款() = False Then Exit Function
    
    If Not mbln已报价 Then Call LedVoiceSpeak
    
    If GetCurCard(objCard) = False Then
        MsgBox "当前" & lblPayType.Caption & "未选择,请选择!", vbOKOnly + vbInformation, gstrSysName
        If cbo支付方式.Enabled And cbo支付方式.Visible Then cbo支付方式.SetFocus
        Exit Function
    End If
    If CheckThreeSwapIsValied(objCard, mdbl当前未退) = False Then
        If cbo支付方式.Enabled And cbo支付方式.Visible Then cbo支付方式.SetFocus
        Exit Function
    End If
    '检查当前单据是否被其他人执行完成,主要是并发原因进行检查
    '防止其他操作员操作:
    gstrSQL = "" & _
    "   Select  1  From 病人预交记录 A " & _
    "   Where   A.结算序号=[1] and nvl(A.校对标志,0)<>0 and Rownum =1 and A.记录状态=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mobjDelBalance.结算序号)
    
    If rsTemp.EOF Then
        '估计是被他人执行,现在需要检查是否被他人执行
        gstrSQL = "Select 记录状态, 操作员姓名,费用状态 From 费用补充记录 Where 结算ID=[1] And rownum=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mobjDelBalance.冲销ID)
        If Not rsTemp.EOF Then
            If Val(Nvl(rsTemp!费用状态)) <> 1 Then
                MsgBox "已经被他人退费结算,不能再进行退费结算!", vbOKOnly + vbInformation, gstrSysName
                '执行收费
                Unload Me
                Exit Function
            End If
            If Nvl(rsTemp!操作员姓名) <> UserInfo.姓名 Then
                MsgBox "该单据不是本人退费结算单,不能处理其他操作员的单据!", vbOKOnly + vbInformation, gstrSysName
                '执行收费
                Unload Me
                Exit Function
            End If
        End If
    End If
    
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
   
Private Sub cbo支付方式_Click()
    Dim objCard As Card, intSelectIndex As Integer
    Dim i As Integer
    
    If mblnNotClick Then Exit Sub
    If mlngPre支付方式 = cbo支付方式.ItemData(cbo支付方式.ListIndex) Then Exit Sub
    
    '105432
    If mlngPre支付方式 > 0 And mdbl当前未退 < 0 Then '只有退款才检查
        If Not mrsOldBalance Is Nothing And Val(txt缴款.Text) <> 0 Then
            '如果不在收费结算方式中就不用检查，主要针对支持“转帐及代扣”的
            Set objCard = mobjPayCards(mlngPre支付方式)
            mrsOldBalance.Filter = "结算方式='" & objCard.结算方式 & "' And 退费=0"
            
            mblnNotClick = True
            intSelectIndex = cbo支付方式.ListIndex
            cbo支付方式.ListIndex = cbo.FindIndex(cbo支付方式, mlngPre支付方式)
            
            If Not mrsOldBalance.EOF Then
                If ThreeBalanceCheck(Me, mlngModule, mobjPayCards(mlngPre支付方式), _
                      mcllForceDelToCash, cbo支付方式.Text) = False Then mblnNotClick = False: Exit Sub
            End If
            
            Set objCard = mobjPayCards(cbo支付方式.ItemData(intSelectIndex))
            If objCard.接口序号 > 0 And objCard.是否转帐及代扣 _
                And mcllForceDelToCash.Count > 0 And mstrDefaultBalance <> objCard.结算方式 Then
                MsgBox "强制退现时，不能选择其它转帐及代扣的结算方式！", vbInformation, gstrSysName
                mblnNotClick = False: Exit Sub
            End If
            
            cbo支付方式.ListIndex = intSelectIndex
            mblnNotClick = False
        End If
    End If
    
    mlngPre支付方式 = cbo支付方式.ItemData(cbo支付方式.ListIndex)
    
    '切换回来后要清除
    Set objCard = mobjPayCards(mlngPre支付方式)
    If objCard.接口序号 > 0 And objCard.消费卡 = False Then
        For i = 1 To mcllForceDelToCash.Count
            If mcllForceDelToCash(i)(1) = objCard.名称 Then Exit For
        Next
        If i <= mcllForceDelToCash.Count Then mcllForceDelToCash.Remove i
    End If
    Call SetControlEnabled
    Call Show误差金额(-1 * mdbl当前未退)
End Sub

Private Sub cbo支付方式_GotFocus()
    If Not mbln已报价 Then Call LedVoiceSpeak
End Sub

Private Sub cbo支付方式_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub cmdExit_Click()
    If gfrmMain Is Nothing Then
       Call ExcuteMainReshData
    End If
    Unload Me: Exit Sub
End Sub

Private Sub ExcuteMainReshData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行主界面的刷新数据
    '编制:刘兴洪
    '日期:2014-06-17 15:09:44
    '说明:主要是应用医保刷新
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not gfrmMain Is Nothing Then Exit Sub
    Call mfrmMain.zlExeBalanceWinRefrshData(mblnOK, mobjDelBalance)
End Sub

Private Sub cmdOK_Click()
    '单据界面按了回车符
    If mblnCacheKeyReturn Then mblnCacheKeyReturn = False: Exit Sub
    
    mblnThreeSwapSingle = False
    If isValied = False Then Exit Sub
    If txt缴款.Text <> "0.00" Then Call ShowLedInfor
    If SaveCharge = False Then Exit Sub
    Unload Me
    Call ExcuteMainReshData
End Sub

Private Sub SetCtrlVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件的显示状态
    '编制:刘兴洪
    '日期:2014-09-18 17:10:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngLeft As Long
    
    cmdOk.Visible = True
    If mbytFunc = EM_BalanceReDel Then
        cmdExit.Visible = True
        lngLeft = cmdOk.Left
        cmdOk.Left = cmdExit.Left
        cmdExit.Left = lngLeft
    Else
        cmdExit.Visible = False
    End If
End Sub
 

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If mblnUnLoad Then Unload Me: Exit Sub
    
    Call SetControlEnabled
    Call SetCtrlVisible
    If cbo支付方式.Enabled And cbo支付方式.Visible Then cbo支付方式.SetFocus
    Call Show误差金额(-1 * mdbl当前未退)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Select Case KeyCode
    Case vbKeyAdd, vbKeyF4
        If gTy_Module_Para.bln使用加减切换 = False And KeyCode = vbKeyAdd Then Exit Sub
        If Me.ActiveControl Is cbo支付方式 Then
            i = cbo支付方式.ListIndex
            If i >= cbo支付方式.ListCount - 1 Then
                i = 0
            Else
                i = i + 1
            End If
            cbo支付方式.ListIndex = i
        End If
    Case vbKeySubtract
        If gTy_Module_Para.bln使用加减切换 = False And KeyCode = vbKeySubtract Then Exit Sub
        If Me.ActiveControl Is cbo支付方式 Then
            i = cbo支付方式.ListIndex
            If i <= 0 Then
                i = cbo支付方式.ListCount - 1
            Else
                i = i - 1
            End If
            cbo支付方式.ListIndex = i
        End If
     Case vbKeyF12
            If Shift = vbCtrlMask Then
                '强制性LED报价,(合计)
                 Call LedVoiceSpeak
            End If
    Case vbKeyF2
        cmdOK_Click '43169
    Case vbKeyReturn
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr(":'|", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Form_Load()
    '选检查主界面中是否发送了回车键的
    mblnCacheKeyReturn = (GetAsyncKeyState(VK_RETURN) And &H1) <> 0
    mstrTittle = "医保补结算退费信息"
    Me.Caption = mstrTittle
    Call InitFace
    zlControl.CboSetWidth cbo支付方式.hWnd, cbo支付方式.Width * 1.3
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    SaveWinState Me, App.ProductName, mstrTittle
    If Not mrsOldBalance Is Nothing Then Set mrsOldBalance = Nothing
End Sub
Private Sub LedVoiceSpeak()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:语音报价
    '编制:刘兴洪
    '日期:2014-09-18 17:20:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gblnLED = False Then Exit Sub
    zl9LedVoice.Speak "#21 " & Format(mdbl当前未退, "0.00")
    mbln已报价 = True
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
   If Panel.Key = "Calc" Then
        mlngR = FindWindow("SciCalc", "计算器")
        If mlngR <> 0 Then
            BringWindowToTop mlngR
        Else
            On Error Resume Next
            Shell "calc.exe", vbNormalFocus
        End If
  End If
End Sub
 
Private Sub txt缴款_GotFocus()
    If Not mbln已报价 Then Call LedVoiceSpeak
End Sub
Private Sub ShowLedInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示Led信息
    '编制:刘兴洪
    '日期:2014-09-18 17:24:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    If gblnLED = False Then Exit Sub
    If Not GetCurCard(objCard) Then Exit Sub
    
    '只有缴现才显示
    If objCard.结算性质 = 1 Then
        zl9LedVoice.DispCharge mdbl当前未退, 0, 0
    Else '部分支付现金时的处理
        Call zl9LedVoice.DisplayBank( _
            "合计:" & mdbl当前未退 & "元,应收:" & -1 * mdbl当前未退 & "元")
    End If
    zl9LedVoice.Speak "#22 " & -1 * Val(txt缴款.Text)
    zl9LedVoice.Speak "#3"
End Sub

Private Sub LedDisplayBank()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示医保信息
    '编制:刘兴洪
    '日期:2014-09-18 17:28:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl个帐合计 As Double, i As Long
    Dim str医保 As String, str三方交易 As String, str老一卡通 As String, str普通结算 As String
    Dim varPara  As Variant, str结算方式 As String
    If Not gblnLED Then Exit Sub
    
    With vsBalance
        For i = 1 To .Rows - 1
            '医保交易
            If .TextMatrix(i, .ColIndex("支付方式")) <> "" Then
                str医保 = str医保 & "||" & .TextMatrix(i, .ColIndex("支付方式")) & ":" & Format(Val(.TextMatrix(i, .ColIndex("支付金额"))), "0.00")
            End If
        Next
    End With
    str结算方式 = Mid(str结算方式, 3)
    varPara = Split(str结算方式, "||")
    
    '目前最多只能显示10个参数值
    Select Case UBound(varPara)
    Case 0
          zl9LedVoice.DisplayBank varPara(0)
    Case 1
          zl9LedVoice.DisplayBank varPara(0), varPara(1)
    Case 2
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2)
    Case 3
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3)
    Case 4
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4)
    Case 5
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5)
    Case 6
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5), varPara(6)
    Case 7
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5), varPara(6), varPara(7)
    Case 8
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5), varPara(6), varPara(7), varPara(8)
    Case 9
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5), varPara(6), varPara(7), varPara(8), varPara(9)
    Case Else
        str结算方式 = ""
         For i = 10 To UBound(varPara)
            str结算方式 = str结算方式 & ";" & varPara(i)
        Next
        If str结算方式 > "" Then str结算方式 = Mid(str结算方式, 2)
        zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5), varPara(6), varPara(7), varPara(8), varPara(9), str结算方式
    End Select
    zl9LedVoice.Speak "#21 " & Format(-1 * mdbl当前未退, "0.00")
End Sub
 
Private Sub txt缴款_KeyPress(KeyAscii As Integer)
    Dim objCard As Card
    zlControl.TxtCheckKeyPress txt缴款, KeyAscii, m金额式
    If KeyAscii <> 13 Then Exit Sub
    If mblnCacheKeyReturn = True Then mblnCacheKeyReturn = False
    If GetCurCard(objCard) = False Then Exit Sub
    KeyAscii = 0
    If objCard.结算性质 = 1 Then
        If cmdOk.Enabled And cmdOk.Visible Then cmdOk.SetFocus
        Exit Sub
    End If
    zlCommFun.PressKey vbKeyTab
End Sub
 
Private Sub txt缴款_LostFocus()
    txt缴款.Text = Format(Val(txt缴款.Text), "0.00")
End Sub

Private Sub txt结算号码_GotFocus()
   zlControl.TxtSelAll txt结算号码
End Sub
Private Sub txt结算号码_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt结算号码_KeyPress(KeyAscii As Integer)
    If InStr(":'|", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    zlControl.TxtCheckKeyPress txt结算号码, KeyAscii, m文本式
End Sub

Private Sub txt摘要_GotFocus()
    zlControl.TxtSelAll txt摘要
End Sub

Private Sub txt摘要_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cmdOk.Visible And cmdOk.Enabled Then cmdOk.SetFocus
    End If
End Sub
 
    
Private Function SaveCharge() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存结算数据
    '返回:保存成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-09-18 15:58:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnTrans  As Boolean, strSQL As String, dblErrMoney As Double '误差费
    Dim objCard As Card, dblMoney As Double, dblTemp As Double
    Dim str结算方式  As String, str结算ID As String
    Dim cllPro As Collection, rsTemp As ADODB.Recordset
    Dim str结帐IDs As String, dbl冲预交 As Double
    
    Err = 0: On Error GoTo errHandle
    If GetCurCard(objCard) = False Then
        MsgBox lblPayType.Caption & "方式未选择!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If txt预存款.Visible Then
        dbl冲预交 = -1 * Val(txt预存款.Text)
    End If
    dblMoney = -1 * mdbl当前未退
    
    Call Show误差金额(dblMoney, dblErrMoney)
    If objCard.结算性质 = 1 Then
        '误差不能大于10块钱
        If Abs(dblErrMoney) > 1.5 Then
            Call MsgBox("误差过大,请检查是否正确!", vbInformation + vbOKOnly, gstrSysName)
            Exit Function
        End If
    ElseIf objCard.结算性质 = -99 Then
        dbl冲预交 = -1 * dblMoney
    End If
    strSQL = "Select distinct 结帐ID From 病人预交记录 where 结算序号=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjDelBalance.结算序号)
    With rsTemp
        Do While Not .EOF
            str结帐IDs = str结帐IDs & "," & Val(Nvl(!结帐ID))
            .MoveNext
        Loop
    End With
    If str结帐IDs = "" Then str结帐IDs = "," & mobjDelBalance.冲销ID
    str结帐IDs = Mid(str结帐IDs, 2)
    
    Set cllPro = New Collection
    If mbytFunc = EM_BalanceReDel Then
        strSQL = "Zl_门诊收费异常_Update("
        strSQL = strSQL & "Null,"
        strSQL = strSQL & "To_Date('" & Format(mobjDelBalance.退费时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
        strSQL = strSQL & mobjDelBalance.冲销ID & ")"
        zlAddArray cllPro, strSQL
        If mobjDelBalance.结帐ID <> 0 Then
            strSQL = "Zl_门诊收费异常_Update("
            strSQL = strSQL & "Null,"
            strSQL = strSQL & "To_Date('" & Format(mobjDelBalance.退费时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
            strSQL = strSQL & mobjDelBalance.结帐ID & ")"
            zlAddArray cllPro, strSQL
        End If
    End If
    
    If mblnThreeSwapSingle = False Then
        If objCard.结算性质 = -99 Then '预交款
            str结算方式 = ""
        Else
            str结算方式 = objCard.结算方式
            str结算方式 = str结算方式 & "|" & dblMoney
            str结算方式 = str结算方式 & "|" & IIf(Trim(txt结算号码.Text) <> "", txt结算号码.Text, " ")
            str结算方式 = str结算方式 & "|" & IIf(Trim(txt摘要.Text) <> "", txt摘要.Text, " ")
        End If
        
        'Zl_费用补充结算_完成退费
        strSQL = "Zl_费用补充结算_完成退费("
        '  结算id_In     In 费用补充记录.结算id%Type,
        strSQL = strSQL & "" & IIf(mobjDelBalance.结帐ID = 0, mobjDelBalance.冲销ID, mobjDelBalance.结帐ID) & ","
        '  结算方式_In   Varchar2,格式:结算方式|结算金额|结算号码|结算摘要
        strSQL = strSQL & "'" & str结算方式 & "',"
        '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
        strSQL = strSQL & "" & IIf(objCard.接口序号 > 0, objCard.接口序号, "NULL") & ","
        '  卡号_In       病人预交记录.卡号%Type := Null,
        strSQL = strSQL & "'" & IIf(objCard.接口序号 > 0, mCurBrushCard.str卡号, "") & "',"
        '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
        strSQL = strSQL & "'" & IIf(objCard.接口序号 > 0, mCurBrushCard.str交易流水号, "") & "',"
        '  交易说明_In   病人预交记录.交易说明%Type := Null,
        strSQL = strSQL & "'" & IIf(objCard.接口序号 > 0, mCurBrushCard.str交易说明, GetForceDelToCashNote(mcllForceDelToCash)) & "',"
        '  误差金额_In   门诊费用记录.实收金额%Type := Null
        strSQL = strSQL & "" & dblErrMoney & ","
        '  完成结算_In       Number := 1,
        strSQL = strSQL & "" & 1 & ","
        '  三方卡按次结算_In Number := 0,
        strSQL = strSQL & "" & 0 & ","
        '  冲预交_In         病人预交记录.冲预交%Type := Null：退款时为负，收款时为正
        strSQL = strSQL & "" & dbl冲预交 & ")"
        zlAddArray cllPro, strSQL
        
        Err = 0: On Error GoTo ErrRoll:
        blnTrans = True
        zlExecuteProcedureArrAy cllPro, Me.Caption, True
        '83222,冉俊明,2015-3-17,结算金额为零时不调用接口直接通过
        If objCard.接口序号 > 0 And RoundEx(dblMoney, 6) <> 0 Then
            If ExecuteThreeSwapPayInterface(objCard, mobjDelBalance.结算序号, str结帐IDs, dblMoney) = False Then Exit Function
        Else
            gcnOracle.CommitTrans
        End If
        blnTrans = False
        mblnOK = True: SaveCharge = True
        Exit Function
    End If

    '三方卡按每一笔单独调用退费接口
    Err = 0: On Error GoTo ErrRoll
    gcnOracle.BeginTrans
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption, True, True
    blnTrans = False
    
    '83222,冉俊明,2015-3-17,结算金额为零时不调用接口直接通过
    If RoundEx(dblMoney, 6) <> 0 Then
       If ExecuteThreeSwapPayInterface(objCard, mobjDelBalance.结算序号, str结帐IDs, dblMoney) = False Then Exit Function
    Else
        gcnOracle.CommitTrans
    End If
    mblnOK = True: SaveCharge = True
    Exit Function
ErrRoll:
    If blnTrans Then gcnOracle.RollbackTrans
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Exit Function
End Function

Private Sub SetControlEnabled()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件的属性
    '编制:刘兴洪
    '日期:2012-02-03 15:08:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnEdit As Boolean, objCard As Card
    
    blnEdit = GetCurCard(objCard)
    txt结算号码.Enabled = blnEdit And objCard.结算性质 <> 1 And objCard.结算性质 <> -99
    txt摘要.Enabled = blnEdit And objCard.结算性质 <> 1 And objCard.结算性质 <> -99
    txt结算号码.BackColor = IIf(txt结算号码.Enabled, &H80000005, Me.BackColor)
    txt摘要.BackColor = IIf(txt摘要.Enabled, &H80000005, Me.BackColor)
    cbo支付方式.Enabled = mbytFunc <> EM_Balance_Err_Cancel
    cbo支付方式.BackColor = IIf(cbo支付方式.Enabled, &H80000005, Me.BackColor)
End Sub

Private Function GetCurCard(ByRef objCard As Card) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取当前卡
    '出参:objCard-返回当前退款或缴款的卡对象
    '返回:成功,返回卡对象
    '编制:刘兴洪
    '日期:2014-07-09 11:03:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intIndex As Integer
    On Error GoTo errHandle
    intIndex = cbo支付方式.ItemData(cbo支付方式.ListIndex)
    If intIndex <= 0 Then Exit Function
    Set objCard = mobjPayCards(intIndex)
    GetCurCard = True
    Exit Function
errHandle:
    Set objCard = New Card
End Function

Private Sub Show误差金额(ByRef dblMoney As Double, Optional ByRef dblErrMoney As Double)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示误差金额
    '入参:dblMoney-本次退的金额
    '出参:dblMoney-本次实际退的金额
    '     dblErrMoney-产生的误差费
    '编制:刘兴洪
    '日期:2014-07-09 18:44:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, dblTemp As Double
    
    On Error GoTo errHandle
    
    If GetCurCard(objCard) = False Then Exit Sub
    dblErrMoney = 0

    If objCard.结算性质 = 1 Then
        '现金
        dblTemp = dblMoney
        If mobjDelBalance.intInsure > 0 Then
            If mbln医保分币 Then
                dblMoney = CentMoney(CCur(dblTemp))
            Else
                dblMoney = Format(dblTemp, "0.00")
            End If
        Else
             dblMoney = CentMoney(CCur(dblTemp))
        End If
       dblErrMoney = RoundEx(dblMoney - dblTemp, 6)
    End If
    
    lbl误差.Visible = dblErrMoney <> 0
    lbl误差.Caption = "误差费:" & zlFormatNum(dblErrMoney)
    lbl误差.Left = cmdOk.Left - lbl误差.Width - 100
    txt缴款.Text = Format(Abs(dblMoney), "0.00")
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function ExecuteThreeSwapPayInterface(objCard As Card, ByVal str结算序号 As String, ByVal str结帐IDs As String, _
    ByVal dblMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:一卡通支付(三方接口)
    '入参:str结算序号-按结算序号进行处理
    '     str结帐Ids-本次更新的结帐IDs
    '     dblMoney-本次支付金额
    '返回:执行成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-09 18:14:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String, strXMLExpend As String
    Dim i As Long, strSQL As String, strTemp As String
    Dim cllUpdate As Collection, cllThreeSwap As Collection
    Dim rsBalance As ADODB.Recordset
    Dim strInXML As String, strOutXML As String
    Dim objXml As clsXML
    Dim dbl冲预交 As Double, cllThreeSwapDel As Collection
    Dim rsTemp As ADODB.Recordset, dblTemp As Double
    Dim lngRow As Long, strValue As String
    Dim lng原结帐ID As Long, str卡号 As String
    Dim str结算方式 As String, lng原交易ID As Long
    
    On Error GoTo errHandle
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    '非一卡通支付,直接返回
    If objCard.接口序号 <= 0 Then ExecuteThreeSwapPayInterface = True: Exit Function
    
    If objCard.是否转帐及代扣 Then
        'zlTransferAccountsMoney
        '参数名  参数类型    入/出   备注
        'frmMain Object  In  调用的主窗体
        'lngModule   Long    In  HIS调用模块号
        'lngCardTypeID   Long    In  卡类别ID
        'strCardNo   String  In  卡号
        'strBalanceID    String  In  结算ID
        'dblMoney    Double  In  转帐金额
        'strSwapGlideNO  String  Out 交易流水号
        'strSwapMemo String  Out 交易说明
        'strSwapExtendInfor  String  In 退费业务时，传入本次退费的冲销ID:
        '                               格式:收费类型1|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
        '                               收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
        '                           Out 交易扩展信息: 格式为:项目名称1|项目内容2||…||项目名称n|项目内容n
        'strXMLExpend String In   XML串:
        '                            <IN>
        '                                <CZLX>操作类型</CZLX> //0或NULL:补结算业务;1-退费业务
        '                            </IN>
        '                    Out  XML串:
        '                            <OUT>
        '                               <ERRMSG>错误信息</ERRMSG >
        '                            </OUT>
        '    Boolean 函数返回    True:调用成功,False:调用失败
        '说明:
        '１. 在医保补充结算时进行的三方转帐时调用。
        '２. 一般来说，成功转帐后，都应该打印相关的结算票据，可以放在此接口进行处理.
        '３. 在转帐成功后，返回交易流水号和相关交易说明；如果存在其他交易信息，可以放在扩展信息中返回.
        '构造XML串
        strXMLExpend = "<IN><CZLX>1</CZLX></IN>"
        '81489,冉俊明,2015-1-22,退费传入冲销ID
        strSwapExtendInfor = "3|" & str结帐IDs: strTemp = strSwapExtendInfor
        If gobjSquare.objSquareCard.zlTransferAccountsMoney(Me, mlngModule, objCard.接口序号, mCurBrushCard.str卡号, _
            str结算序号, dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor, strXMLExpend) = False Then
            gcnOracle.RollbackTrans: Call ShowErrMsg(1, strXMLExpend)
            Exit Function
        End If
        mCurBrushCard.str交易流水号 = strSwapGlideNO
        mCurBrushCard.str交易说明 = strSwapMemo
        
        Call zlAddUpdateSwapSQL(False, str结帐IDs, objCard.接口序号, objCard.消费卡, mCurBrushCard.str卡号, strSwapGlideNO, strSwapMemo, cllUpdate, 0)
        
        If strTemp <> strSwapExtendInfor Then
            Call zlAddThreeSwapSQLToCollection(False, str结帐IDs, objCard.接口序号, objCard.消费卡, mCurBrushCard.str卡号, strSwapExtendInfor, cllThreeSwap)
        End If
        zlExecuteProcedureArrAy cllUpdate, Me.Caption, True, True
        gcnOracle.CommitTrans
    Else
        If mblnThreeSwapSingle Then
            Set rsBalance = zlGetCanDelBalanceRecords(str结算序号, objCard.接口序号)
            dblTemp = dblMoney
            Do While Not rsBalance.EOF And RoundEx(dblTemp, 6) > 0
                If Val(Nvl(rsBalance!金额)) > RoundEx(dblTemp, 6) Then
                    dbl冲预交 = RoundEx(dblTemp, 6)
                    dblTemp = 0
                Else
                    dbl冲预交 = Val(Nvl(rsBalance!金额))
                    dblTemp = dblTemp - Val(Nvl(rsBalance!金额))
                End If
                
                lng原交易ID = Val(Nvl(rsBalance!原交易ID))
                lng原结帐ID = Val(Nvl(rsBalance!结帐ID))
                str卡号 = Nvl(rsBalance!卡号)
                strSwapGlideNO = Nvl(rsBalance!交易流水号)
                strSwapMemo = Nvl(rsBalance!交易说明)
                strSwapExtendInfor = "3|" & str结帐IDs: strTemp = strSwapExtendInfor
                'zlReturnMoney(frmMain As Object, ByVal lngModule As Long, _
                    ByVal lngCardTypeID As Long, ByVal strCardNo As String, _
                    ByVal strBalanceIDs As String, ByVal dblMoney As Double, _
                    ByRef strSwapGlideNO As String, ByRef strSwapMemo As String, _
                    ByRef strSwapExtendInfor As String) As Boolean
                '---------------------------------------------------------------------------------------------------------------------------------------------
                '功能:帐户扣款回退交易
                '入参:frmMain-调用的主窗体
                '       lngModule-调用的模块号
                '       lngCardTypeID-卡类别ID:医疗卡类别.ID
                '       strCardNo-卡号
                '       strBalanceIDs-本次支付所涉及的结算ID(这是原结帐ID):
                '                           格式:收费类型(|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
                '                           收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
                '       dblMoney-退款金额
                '       strSwapNo-交易流水号(扣款时的交易流水号)
                '       strSwapMemo-交易说明(扣款时的交易说明)
                '       strSwapExtendInfor-出入，本次退费的冲销ID：
                '                           格式:收费类型1|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
                '                           收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
                '       strSwapExtendInfor-传出，交易的扩展信息
                '           格式为:项目名称1|项目内容2||…||项目名称n|项目内容n 每个项目中不能包含|字符
                If gobjSquare.objSquareCard.zlReturnMoney(Me, mlngModule, objCard.接口序号, _
                    objCard.消费卡, str卡号, "3|" & lng原结帐ID, dbl冲预交, _
                    strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then
                    gcnOracle.RollbackTrans
                    
                    Call LoadData(str结算序号)
                    mdbl当前未退 = mdblDelMoney
                    Exit Function
                End If
                
                'Zl_费用补充结算_完成退费
                strSQL = "Zl_费用补充结算_完成退费("
                '  结算id_In     In 费用补充记录.结算id%Type,
                strSQL = strSQL & "" & IIf(mobjDelBalance.结帐ID = 0, mobjDelBalance.冲销ID, mobjDelBalance.结帐ID) & ","
                str结算方式 = objCard.结算方式
                str结算方式 = str结算方式 & "|" & -1 * dbl冲预交
                str结算方式 = str结算方式 & "|" & IIf(Trim(txt结算号码.Text) <> "", txt结算号码.Text, " ")
                str结算方式 = str结算方式 & "|" & IIf(Trim(txt摘要.Text) <> "", txt摘要.Text, " ")
                '  结算方式_In   Varchar2,格式:结算方式|结算金额|结算号码|结算摘要
                strSQL = strSQL & "'" & str结算方式 & "',"
                '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
                strSQL = strSQL & "" & IIf(objCard.接口序号 > 0, objCard.接口序号, "NULL") & ","
                '  卡号_In       病人预交记录.卡号%Type := Null,
                strSQL = strSQL & "" & "NULL" & ","
                '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
                strSQL = strSQL & "" & "NULL" & ","
                '  交易说明_In   病人预交记录.交易说明%Type := Null,
                strSQL = strSQL & "" & "NULL" & ","
                '  误差金额_In   门诊费用记录.实收金额%Type := Null
                strSQL = strSQL & "" & 0 & ","
                '  完成结算_In Number:=0:1-完成补充结算;0-未完成补充结算
                strSQL = strSQL & "" & IIf(RoundEx(dblTemp, 6) > 0, 0, 1) & ","
                '  三方卡按次结算_In Number:=0
                strSQL = strSQL & "" & 1 & ")"
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
                
                strSQL = "Zl_三方退款信息_Insert("
                strSQL = strSQL & "" & Val(str结算序号) & ","
                strSQL = strSQL & "" & lng原结帐ID & ","
                strSQL = strSQL & "" & dbl冲预交 & ","
                strSQL = strSQL & "'" & str卡号 & "',"
                strSQL = strSQL & "'" & strSwapGlideNO & "',"
                strSQL = strSQL & "'" & strSwapMemo & "')"
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
                
                gcnOracle.CommitTrans
                
                Set cllThreeSwap = New Collection
                If strTemp <> strSwapExtendInfor Then
                    Call zlAddThreeSwapSQLToCollection(False, Abs(Val(str结算序号)), _
                        objCard.接口序号, objCard.消费卡, str卡号, strSwapExtendInfor, cllThreeSwap, lng原交易ID)
                End If
                zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
                gcnOracle.BeginTrans
                
                rsBalance.MoveNext
            Loop
            gcnOracle.CommitTrans
            
            If RoundEx(dblTemp, 6) > 0 Then
                MsgBox "退款金额(" & Format(dblMoney, "0.00") & ")大于可退金额(" & Format(dblMoney - dblTemp, "0.00") & ")！", vbOKOnly + vbInformation, gstrSysName
                
                Call LoadData(str结算序号)
                mdbl当前未退 = mdblDelMoney
                Exit Function
            End If
            
            ExecuteThreeSwapPayInterface = True
            Exit Function
        Else
            'Public Function zlReturnMultiMoney(frmMain As Object, ByVal lngModule As Long, _
                ByVal lngCardTypeID As Long, ByVal bln消费卡 As Boolean, ByVal strInXML As String, _
                ByVal lng冲销ID As Long, ByRef strOutXml As String, ByRef strExpend As String) As Boolean
            '---------------------------------------------------------------------------------
            '功能:帐户扣款回退交易(多笔回退)
            '入参:frmMain-调用的主窗体
            '       lngModule-调用的模块号
            '       lngCardTypeID-卡类别ID:医疗卡类别.ID
            '       strInXML-XML串:
            '       <JSLIST>
            '           <JS>
            '               <KH>卡号</KH>
            '               <JYLSH>交易流水号</JYLSH>
            '               <JYSM>交易说明</JYSM>
            '               <ZFJE>作废金额</ZFJE>
            '               <JSLX>类型</JSLX>  //1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款,6-补充结算
            '               <ID></ID>    //类型=1时,预交ID;类型=2,6时，为原结帐ID
            '           </JS>
            '       </JSLIST>
            '       lng冲销ID-作废时的冲销ID(作废时或退费时有效，否则为0）;类型=6，传入结算序号
            '       strExpend-无（暂留，待以后扩展)
            '出参:
            '     strOutXML-返回XML串
            '       <JSLIST>
            '           <JS>
            '               <KH>卡号</KH>
            '               <TKLSH>退款交易流水号</TKLSH>
            '               <TKSM>退款交易说明</TKSM>
            '               <ID></ID>
            '           </JS>
            '       </JSLIST>
            '      strExpend-交易的扩展信息
            '       <EXPENDS>
            '           <EXPEND>
            '               <XMMC>项目名称1</XMMC>
            '               <XMNR>项目内容2</XMNR>
            '           </EXPEND>
            '       </EXPENDS>
            '返回:函数返回    True:调用成功,False:调用失败
            '日期:2015-11-10
            '说明:
            '   目前只有结帐程序时有效（结帐退款),用于一次性处理同一卡类别的多笔三方交易作废
            '--------------------------------------------------------------------------------
            Set cllThreeSwap = New Collection: Set cllThreeSwapDel = New Collection
            Set objXml = New clsXML
            objXml.ClearXmlText
            
            Set rsBalance = zlGetCanDelBalanceRecords(str结算序号, objCard.接口序号)
            dblTemp = dblMoney
            
            objXml.AppendNode "JSLIST"
            Do While Not rsBalance.EOF And RoundEx(dblTemp, 6) > 0
                If Val(Nvl(rsBalance!金额)) > RoundEx(dblTemp, 6) Then
                    dbl冲预交 = RoundEx(dblTemp, 6)
                    dblTemp = 0
                Else
                    dbl冲预交 = Val(Nvl(rsBalance!金额))
                    dblTemp = dblTemp - Val(Nvl(rsBalance!金额))
                End If
                
                objXml.AppendNode "JS"
                    objXml.appendData "KH", Nvl(rsBalance!卡号), xsString
                    objXml.appendData "JYLSH", Nvl(rsBalance!交易流水号), xsString
                    objXml.appendData "JYSM", Nvl(rsBalance!交易说明), xsString
                    objXml.appendData "ZFJE", dbl冲预交, xsNumber
                    objXml.appendData "JSLX", 6, xsNumber
                    objXml.appendData "ID", Val(Nvl(rsBalance!结帐ID)), xsNumber
                objXml.AppendNode "JS", True
                
                strSQL = "Zl_三方退款信息_Insert("
                strSQL = strSQL & "" & Val(str结算序号) & ","
                strSQL = strSQL & "" & Val(Nvl(rsBalance!结帐ID)) & ","
                strSQL = strSQL & "" & dbl冲预交 & ","
                strSQL = strSQL & "'" & Nvl(rsBalance!卡号) & "',"
                strSQL = strSQL & "'" & Nvl(rsBalance!交易流水号) & "',"
                strSQL = strSQL & "'" & Nvl(rsBalance!交易说明) & "')"
                zlAddArray cllThreeSwapDel, strSQL
                rsBalance.MoveNext
            Loop
            objXml.AppendNode "JSLIST", True
            
            strInXML = objXml.XmlText
            strOutXML = "": strXMLExpend = ""
            If gobjSquare.objSquareCard.zlReturnMultiMoney(Me, mlngModule, objCard.接口序号, objCard.消费卡, strInXML, _
                 Val(str结算序号), strOutXML, strXMLExpend) = False Then
                gcnOracle.RollbackTrans
                Call ShowErrMsg(1, strXMLExpend)
                Exit Function
            End If
                 
            If strOutXML <> "" Then
                If zlXML_Init = False Then Exit Function
                If zlXML_LoadXMLToDOMDocument(strOutXML, False) = False Then Exit Function
                Call zlXML_GetChildRows("JSLIST", "JS", lngRow)
                For i = 0 To lngRow - 1
                    strSQL = "Zl_三方退款信息_Insert("
                    strSQL = strSQL & "" & Val(str结算序号) & ","
                    Call zlXML_GetNodeValue("ID", i, strValue)
                    strSQL = strSQL & "" & Val(strValue) & ","
                    strSQL = strSQL & "" & 0 & ","
                    Call zlXML_GetNodeValue("KH", i, strValue)
                    strSQL = strSQL & "'" & strValue & "',"
                    Call zlXML_GetNodeValue("TKLSH", i, strValue)
                    strSQL = strSQL & "'" & strValue & "',"
                    Call zlXML_GetNodeValue("TKSM", i, strValue)
                    strSQL = strSQL & "'" & strValue & "',"
                    strSQL = strSQL & "" & 1 & ")"
                    zlAddArray cllThreeSwapDel, strSQL
                Next
            End If
            
            If strXMLExpend <> "" Then
                strSwapExtendInfor = ""
                If zlXML_LoadXMLToDOMDocument(strXMLExpend, False) = False Then Exit Function
                Call zlXML_GetChildRows("EXPENDS", "EXPEND", lngRow)
                For i = 0 To lngRow - 1
                    Call zlXML_GetNodeValue("XMMC", i, strValue)
                    strSwapExtendInfor = strSwapExtendInfor & "||" & strValue
                    Call zlXML_GetNodeValue("XMNR", i, strValue)
                    strSwapExtendInfor = strSwapExtendInfor & "|" & strValue
                Next i
            End If
            If strSwapExtendInfor <> "" Then strSwapExtendInfor = Mid(strSwapExtendInfor, 3)
            
            Call zlAddUpdateSwapSQL(False, Abs(Val(str结算序号)), objCard.接口序号, objCard.消费卡, mCurBrushCard.str卡号, "", "", cllUpdate, 0)
            Call zlAddThreeSwapSQLToCollection(False, Abs(Val(str结算序号)), objCard.接口序号, objCard.消费卡, mCurBrushCard.str卡号, strSwapExtendInfor, cllThreeSwap)
            zlExecuteProcedureArrAy cllThreeSwapDel, Me.Caption, True, True
            zlExecuteProcedureArrAy cllUpdate, Me.Caption, True, True
            gcnOracle.CommitTrans
        End If
    End If
    
    Err = 0: On Error GoTo ErrOtherHand:
    '更新其他结算信息
    zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
    ExecuteThreeSwapPayInterface = True
    Exit Function
ErrOtherHand:
    ExecuteThreeSwapPayInterface = True
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    If mblnThreeSwapSingle Then
        Call LoadData(str结算序号)
        mdbl当前未退 = mdblDelMoney
    End If
End Function

Private Sub ShowErrMsg(ByVal bytType As Byte, ByVal strXMLErrMsg As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:三方转账检查与代扣业务出错提示
    '编制:冉俊明
    '时间:2014-12-2
    '参数:
    '   bytType:0-转账检查,1-转账交易
    '   strXMLErrMsg:格式如下
    '            <OUT>
    '               <ERRMSG>错误信息</ERRMSG >
    '            </OUT>
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    
    On Error GoTo errHandle
    '解析错误信息
    If strXMLErrMsg <> "" Then
        If zlXML.OpenXMLDocument(strXMLErrMsg) = False Then strValue = ""
        If zlXML.GetSingleNodeValue("OUT/ERRMSG", strValue) = False Then strValue = ""
        Call zlXML.CloseXMLDocument
    End If
    '提示错误信息
    If Trim(strValue) = "" Then
        If bytType = 0 Then
            strValue = vbCrLf & "交易检查失败！"
        Else
            strValue = vbCrLf & "交易失败！"
        End If
    End If
    MsgBox strValue, vbExclamation + vbOKOnly, gstrSysName
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Function CheckThreeSwapIsValied(ByVal objCard As Card, dblMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:刷卡验证
    '入参:objCard-当前卡
    '返回:刷卡成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-09-18 15:03:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strXMLExpend As String, strExpand As String
    Dim cllSquareBalance As New Collection
    Dim dblTemp As Double, dbl帐户余额 As Double
    Dim blnTransfer As Boolean, strBalanceIDs As String
    Dim rsBalance As ADODB.Recordset
    Dim strCardNo As String, strPassWord As String
    Dim dbl冲预交 As Double
    
    On Error GoTo errHandle
    
    If objCard.接口序号 <= 0 Then CheckThreeSwapIsValied = True: Exit Function
    If dblMoney = 0 Then CheckThreeSwapIsValied = True: Exit Function
    
    '医保报销金额大于总费用金额的只能进行转帐及代扣
    blnTransfer = zlCheckOnlyUseTrans(mobjDelBalance.结算序号)
    If blnTransfer And objCard.是否转帐及代扣 = False Then
        MsgBox "医保报销金额大于了总费用金额， " & objCard.名称 & " 不支持转帐及代扣，请选择其它支付方式！", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    If dblMoney > 0 And objCard.是否转帐及代扣 = False Then
        MsgBox "当前为收款， " & objCard.名称 & " 不支持转帐及代扣，请选择其它支付方式！", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    
    If objCard.是否转帐及代扣 Then
        '   zlBrushCard(frmMain As Object, _
        ByVal lngModule As Long, _
        ByVal rsClassMoney As ADODB.Recordset, _
        ByVal lngCardTypeID As Long, _
        ByVal bln消费卡 As Boolean, _
        ByVal strPatiName As String, ByVal strSex As String, _
        ByVal strOld As String, ByRef dbl金额 As Double, _
        Optional ByRef strCardNo As String, _
        Optional ByRef strPassWord As String, _
        Optional ByRef bln退费 As Boolean = False, _
        Optional ByRef blnShowPatiInfor As Boolean = False, _
        Optional ByRef bln退现 As Boolean = False, _
        Optional ByVal bln余额不足禁止 As Boolean = True, _
        Optional ByRef varSquareBalance As Variant, _
        Optional ByVal bln转预交 As Boolean = False, _
        Optional ByVal blnAllPay As Boolean = False, _
        Optional ByVal strXmlIn As String = "") As Boolean
        '       strXmlIn-三方卡调用XML入参,目前格式如下:
        '       <IN>
        '           <CZLX>0</CZLX>    //操作类型,0-正常调用刷卡,1-转账调用刷卡,2-退款调用刷卡
        '       </IN>
        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModule, Nothing, objCard.接口序号, False, _
            mobjDelBalance.姓名, mobjDelBalance.性别, mobjDelBalance.年龄, dblMoney, strCardNo, strPassWord, _
            False, True, False, False, cllSquareBalance, False, False, "<IN><CZLX>1</CZLX></IN>") = False Then Exit Function
        mCurBrushCard.str卡号 = strCardNo
        mCurBrushCard.str密码 = strPassWord
    
        '调用转帐接口
        'zlTransferAccountsCheck 转帐检查接口
        '参数名  参数类型    入/出   备注
        'frmMain Object  In  调用的主窗体
        'lngModule   Long    In  HIS调用模块号
        'lngCardTypeID   Long    In  卡类别ID
        'strCardNo   String  In  卡号
        'dblMoney    Double  In  转帐金额(代扣时为负数)
        'strBalanceIDs   String  In  结帐IDs，多个用逗号分离，表示本次对哪此收费项目进行重新医保补结算
        'strXMLExpend String In   XML串:
        '                            <IN>
        '                                <CZLX>操作类型</CZLX> //0或NULL:补结算业务;1-退费业务
        '                            </IN>
        '                    Out  XML串:
        '                            <OUT>
        '                               <ERRMSG>错误信息</ERRMSG >
        '                            </OUT>
        '    Boolean 函数返回    检查的数据合法,返回True:否则返回False
        '说明:
        '１. 在医保补充结算时进行的三方转帐时的一些合法性检查，避免在转帐时弹出对话框之类的等待造成死锁或其它现象的发生。
        '２. 不存在检测的需要返回为True，否则不能完成转帐功能的调用。
        '构造XML串
        strXMLExpend = "<IN><CZLX>1</CZLX></IN>"
        If gobjSquare.objSquareCard.zltransferAccountsCheck(Me, mlngModule, objCard.接口序号, _
            mCurBrushCard.str卡号, -1 * dblMoney, mobjDelBalance.结算序号, strXMLExpend) = False Then
            Call ShowErrMsg(0, strXMLExpend)
            Exit Function
        End If
    Else
        'ZlGetParaConfig(ByVal frmMain As Object, _
            ByVal lng卡类别ID As Long, ByVal bln消费卡 As Boolean, ByVal intPara As Integer, _
            Optional strErrMsg As String, Optional strExpend As String) As Boolean
            '---------------------------------------------------------------------------------------------------------------------------------------------
            '功能:获取接口参数
            '入参: frmMain-调用的主窗体
            '       intPara: 包含如下值
            '                1-刷卡和支付在同一页面:true-新模式；False-旧模式
            '                2-作废时是否单独调用退费接口
            '       strExpend-扩展参数，暂留，现传为空
            '出参:strErrMsg-返回的错误信息
            '       strExpend-扩展参数，暂留，现传为空
            '返回:函数返回True:调用成功,False:调用失败
        mblnThreeSwapSingle = gobjSquare.objSquareCard.ZlGetParaConfig(Me, objCard.接口序号, objCard.消费卡, 2)
        If mblnThreeSwapSingle Then
            Set rsBalance = zlGetCanDelBalanceRecords(mobjDelBalance.结算序号, objCard.接口序号)
            dblTemp = -1 * dblMoney
            Do While Not rsBalance.EOF And RoundEx(dblTemp, 6) > 0
                If Val(Nvl(rsBalance!金额)) > RoundEx(dblTemp, 6) Then
                    dbl冲预交 = RoundEx(dblTemp, 6)
                    dblTemp = 0
                Else
                    dbl冲预交 = Val(Nvl(rsBalance!金额))
                    dblTemp = dblTemp - Val(Nvl(rsBalance!金额))
                End If
                
                strBalanceIDs = "6|" & Nvl(rsBalance!结帐ID)
                mCurBrushCard.str卡号 = Nvl(rsBalance!卡号)
                'zlReturnCheck(frmMain As Object, ByVal lngModule As Long, _
                    ByVal lngCardTypeID As Long, bln消费卡 As Boolean, ByVal strCardNo As String, _
                    ByVal strBalanceIDs As String, _
                    ByVal dblMoney As Double, ByVal strSwapNo As String, _
                    ByVal strSwapMemo As String, ByRef strXMLExpend As String) As Boolean
                    '---------------------------------------------------------------------------------------------------------------------------------------------
                    '功能:帐户回退交易前的检查
                    '入参:frmMain-调用的主窗体
                    '       lngModule-调用的模块号
                    '       lngCardTypeID-卡类别ID
                    '       strCardNo-卡号
                    '       strBalanceIDs   String  In  本次支付所涉及的结算ID 格式:收费类型|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
                    '                                   收费类型: 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
                    '       dblMoney-退款金额
                    '       strSwapNo-交易流水号(退款时检查)
                    '       strSwapMemo-交易说明(退款时传入)
                    '       strXMLExpend    XML IN  可选参数:异常单据重新退费(1)
                    '返回:退款合法,返回true,否则返回Flase
                    '说明:
                    '    在调用扣款前，由于存在Oracle事务问题，因此，再调用回退交易前，先进行数据的合法性检查,
                    '    以便控制死锁情况。
                If gobjSquare.objSquareCard.zlReturnCheck(Me, mlngModule, objCard.接口序号, _
                    objCard.消费卡, mCurBrushCard.str卡号, strBalanceIDs, dbl冲预交, _
                    Nvl(rsBalance!交易流水号), Nvl(rsBalance!交易说明), _
                    strXMLExpend) = False Then Exit Function
                
                If objCard.是否退款验卡 Then
                    'zlBrushCard(frmMain As Object, _
                        ByVal lngModule As Long, _
                        ByVal rsClassMoney As ADODB.Recordset, _
                        ByVal lngCardTypeID As Long, _
                        ByVal bln消费卡 As Boolean, _
                        ByVal strPatiName As String, ByVal strSex As String, _
                        ByVal strOld As String, ByRef dbl金额 As Double, _
                        Optional ByRef strCardNo As String, _
                        Optional ByRef strPassWord As String, _
                        Optional ByRef bln退费 As Boolean = False, _
                        Optional ByRef blnShowPatiInfor As Boolean = False, _
                        Optional ByRef bln退现 As Boolean = False, _
                        Optional ByVal bln余额不足禁止 As Boolean = True, _
                        Optional ByRef varSquareBalance As Variant, _
                        Optional ByVal bln转预交 As Boolean = False, _
                        Optional ByVal blnAllPay As Boolean = False, _
                        Optional ByVal strXmlIn As String = "") As Boolean
                        '---------------------------------------------------------------------------------------------------------------------------------------------
                        '功能:根据指定支付类别,弹出刷卡窗口
                        '入参:rsClassMoney:收费类别,金额
                        '        lngCardTypeID-为零时,为老一卡通刷卡
                        '       bln余额不足禁止-目前只针对消费卡,表示余额不足时,禁止继续操作,否则用余额进行支付
                        '       dblBrushTotaled-消费有效,表示已经刷消费卡总额(主要用于多次刷卡)
                        '       str上次限制类别-上次刷消费时的限制类别(同次多次刷消费卡时,需要检查本次刷卡类别与上次类别是否一致,不一致不允许刷卡消费)
                        '       varSquareBalance- Collection类型,当前已经刷卡的信息(array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文 ))
                        '       bln预交-是否转预交
                        '       blnAllPay-是否费用全支付，true-费用未支付完不能完成结算，false-可以只支付部分并返回
                        '       strXmlIn-XML入参,目前格式如下:
                        '       <IN>
                        '           <CZLX>0</CZLX>    //操作类型,0-正常调用刷卡,1-转账调用刷卡,2-退款调用刷卡
                        '       </IN>
                        '出参:str限制类别-限制类别(消费卡返回)
                        '        lng消费卡ID-消费卡目录.ID(消费卡返回)
                        '       strCardNO-返回刷卡的卡号
                        '       strPassWord-返回刷卡所对应的密码
                        '       varSquareBalance- Collection类型,返回当前刷卡数据(array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文))
                        '返回:成功,返回true,否则返回False
                    strCardNo = mCurBrushCard.str卡号
                    If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModule, Nothing, _
                        objCard.接口序号, False, mobjDelBalance.姓名, mobjDelBalance.性别, mobjDelBalance.年龄, _
                        dbl冲预交, strCardNo, strPassWord, _
                        False, True, False, False, cllSquareBalance, False, False, _
                        "<IN><CZLX>2</CZLX></IN>") = False Then Exit Function
                    mCurBrushCard.str卡号 = strCardNo
                    mCurBrushCard.str密码 = strPassWord
                End If
                
                rsBalance.MoveNext
            Loop
            
            If RoundEx(dblTemp, 6) > 0 Then
                MsgBox "退款金额(" & Format(-1 * dblMoney, "0.00") & ")大于可退金额(" & Format(-1 * dblMoney - dblTemp, "0.00") & ")！", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
        Else
            Set rsBalance = zlGetCanDelBalanceRecords(mobjDelBalance.结算序号, objCard.接口序号)
            dblTemp = -1 * dblMoney: strBalanceIDs = ""
            Do While Not rsBalance.EOF And RoundEx(dblTemp, 6) > 0
                strBalanceIDs = strBalanceIDs & "," & Nvl(rsBalance!结帐ID)
                dblTemp = dblTemp - Val(Nvl(rsBalance!金额))
                rsBalance.MoveNext
            Loop
            If strBalanceIDs <> "" Then strBalanceIDs = Mid(strBalanceIDs, 2)
            strBalanceIDs = "6|" & strBalanceIDs
            If RoundEx(dblTemp, 6) > 0 Then
                MsgBox "退款金额(" & Format(-1 * dblMoney, "0.00") & ")大于可退金额(" & Format(-1 * dblMoney - dblTemp, "0.00") & ")！", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
                
            strXMLExpend = mfrmMain.GetDelXMLExpend()
            'zlReturnCheck(frmMain As Object, ByVal lngModule As Long, _
            ByVal lngCardTypeID As Long, bln消费卡 As Boolean, ByVal strCardNo As String, _
            ByVal strBalanceIDs As String, _
            ByVal dblMoney As Double, ByVal strSwapNo As String, _
            ByVal strSwapMemo As String, ByRef strXMLExpend As String) As Boolean
            '---------------------------------------------------------------------------------------------------------------------------------------------
            '功能:帐户回退交易前的检查
            '入参:frmMain-调用的主窗体
            '       lngModule-调用的模块号
            '       lngCardTypeID-卡类别ID
            '       strCardNo-卡号
            '       strBalanceIDs   String  In  本次支付所涉及的结算ID 格式:收费类型|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
            '                                   收费类型: 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
            '       dblMoney-退款金额
            '       strSwapNo-交易流水号(退款时检查)
            '       strSwapMemo-交易说明(退款时传入)
            '       strXMLExpend    XML IN  可选参数(扩展用):
            '        <TFDATA> //退费数据
            '          <YCTF>1</YCTF> //是否异常重退:1-异常重退;0-退费 此节点可能没有
            '          <TFLIST> //退费列表
            '            <NO></NO> // 退费单据
            '            <TFITEM> //退费项
            '              <SerialNum></SerialNum> //序号
            '              …
            '            </TFITEM>
            '          </TFLIST>
            '          ....
            '        </TFDATA >
            '返回:退款合法,返回true,否则返回Flase
            If gobjSquare.objSquareCard.zlReturnCheck(Me, mlngModule, objCard.接口序号, objCard.消费卡, mCurBrushCard.str卡号, _
                strBalanceIDs, -1 * dblMoney, "", "", strXMLExpend) = False Then Exit Function
                
            If objCard.是否退款验卡 Then
                '   zlBrushCard(frmMain As Object, _
                ByVal lngModule As Long, _
                ByVal rsClassMoney As ADODB.Recordset, _
                ByVal lngCardTypeID As Long, _
                ByVal bln消费卡 As Boolean, _
                ByVal strPatiName As String, ByVal strSex As String, _
                ByVal strOld As String, ByRef dbl金额 As Double, _
                Optional ByRef strCardNo As String, _
                Optional ByRef strPassWord As String, _
                Optional ByRef bln退费 As Boolean = False, _
                Optional ByRef blnShowPatiInfor As Boolean = False, _
                Optional ByRef bln退现 As Boolean = False, _
                Optional ByVal bln余额不足禁止 As Boolean = True, _
                Optional ByRef varSquareBalance As Variant, _
                Optional ByVal bln转预交 As Boolean = False, _
                Optional ByVal blnAllPay As Boolean = False, _
                Optional ByVal strXmlIn As String = "") As Boolean
                '       strXmlIn-三方卡调用XML入参,目前格式如下:
                '       <IN>
                '           <CZLX>0</CZLX>    //操作类型,0-正常调用刷卡,1-转账调用刷卡,2-退款调用刷卡
                '       </IN>
                If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModule, Nothing, objCard.接口序号, False, _
                    mobjDelBalance.姓名, mobjDelBalance.性别, mobjDelBalance.年龄, -1 * dblMoney, strCardNo, strPassWord, _
                    False, True, False, False, cllSquareBalance, False, False, "<IN><CZLX>2</CZLX></IN>") = False Then Exit Function
                mCurBrushCard.str卡号 = strCardNo
                mCurBrushCard.str密码 = strPassWord
            End If
        End If
    End If
    
    'zlGetAccountMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
    'ByVal strCardTypeID As Long, _
    'ByVal strCardNo As String, strExpand As String, dblMoney As Double
    '入参:frmMain-调用的主窗体
    '        lngModule-模块号
    '        strCardNo-卡号
    '        strExpand-预留，为空,以后扩展
    '出参:dblMoney-返回帐户余额
    Call gobjSquare.objSquareCard.zlGetAccountMoney(Me, mlngModule, objCard.接口序号, _
          mCurBrushCard.str卡号, strExpand, dbl帐户余额, objCard.消费卡)
    If dbl帐户余额 <> 0 Then
        stbThis.Panels(C4三方帐户信息).Text = objCard.结算方式 & "帐户余额:" & Format(dbl帐户余额, "0.00")
        stbThis.Panels(C4三方帐户信息).ToolTipText = objCard.结算方式 & "的帐户余额:" & Format(dbl帐户余额, "0.00")
        stbThis.Panels(C4三方帐户信息).Visible = True
    End If
    CheckThreeSwapIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub txt预存款_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt预存款_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt预存款, KeyAscii, m金额式
End Sub

Private Sub txt预存款_LostFocus()
    If IsValid预交款() = False Then Exit Sub
    
    mdbl当前未退 = mdblDelMoney - (-1 * Val(txt预存款.Text))
    Call Show误差金额(-1 * mdbl当前未退)
End Sub

Private Function IsValid预交款() As Boolean
    '退预交款检查
    On Error GoTo errHandle
    If txt预存款.Visible = False Then IsValid预交款 = True: Exit Function
    
    If txt预存款.Text = "" Then
        txt预存款.Text = "0.00"
    ElseIf Not IsNumeric(txt预存款.Text) And txt预存款.Text <> "" Then
        ShowMsgbox "无效数值！"
        txt预存款.Text = Format(mdbl可退预交, "0.00")
        zlControl.ControlSetFocus txt预存款
        Exit Function
    ElseIf Val(txt预存款.Text) < 0 Then
        ShowMsgbox "预存款退款金额不能为负！"
        zlControl.ControlSetFocus txt预存款
        txt预存款.Text = Format(mdbl可退预交, "0.00")
        Exit Function
    ElseIf Val(txt预存款.Text) > mdbl可退预交 Then
        ShowMsgbox "预存款退款金额不能超过可退金额:" & Format(mdbl可退预交, "0.00") & " ！"
        txt预存款.Text = Format(mdbl可退预交, "0.00")
        zlControl.ControlSetFocus txt预存款
        Exit Function
    Else
        txt预存款.Text = Format(Val(txt预存款.Text), "0.00")
    End If
    IsValid预交款 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub txt预存款_GotFocus()
    zlControl.TxtSelAll txt预存款
End Sub
