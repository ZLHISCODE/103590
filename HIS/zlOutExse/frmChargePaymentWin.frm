VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChargePayMentWin 
   Caption         =   "病人收费结算"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10365
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   15.75
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChargePaymentWin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   10365
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdExit 
      Caption         =   "返回录入(&X)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   8250
      TabIndex        =   37
      Top             =   2265
      Width           =   2055
   End
   Begin VB.CommandButton cmdYBBalance 
      Caption         =   "医保结算(&Y)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   8235
      TabIndex        =   36
      Top             =   255
      Width           =   2055
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "继续收费(&J)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   8235
      TabIndex        =   32
      Top             =   915
      Width           =   2055
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   3090
      Left            =   45
      ScaleHeight     =   3090
      ScaleWidth      =   7995
      TabIndex        =   24
      Top             =   990
      Width           =   7995
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         ForeColor       =   &H80000008&
         Height          =   1350
         Left            =   45
         ScaleHeight     =   1320
         ScaleWidth      =   3060
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1650
         Width           =   3090
         Begin VB.Label lbl自付合计 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   615
            Left            =   2040
            TabIndex        =   30
            Top             =   615
            Width           =   1005
         End
         Begin XtremeSuiteControls.ShortcutCaption ShortcutCaption2 
            Height          =   420
            Left            =   15
            TabIndex        =   29
            Top             =   30
            Width           =   3045
            _Version        =   589884
            _ExtentX        =   5371
            _ExtentY        =   741
            _StockProps     =   6
            Caption         =   "自付合计"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   15.76
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
         End
      End
      Begin VB.PictureBox picPay 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         ForeColor       =   &H80000008&
         Height          =   2910
         Left            =   3195
         ScaleHeight     =   2880
         ScaleWidth      =   4710
         TabIndex        =   26
         Top             =   90
         Width           =   4740
         Begin VB.TextBox txt冲预交 
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
            Left            =   1380
            MaxLength       =   10
            TabIndex        =   3
            Top             =   165
            Width           =   3240
         End
         Begin VB.ComboBox cbo支付方式 
            BackColor       =   &H8000000F&
            ForeColor       =   &H8000000D&
            Height          =   435
            Left            =   1380
            Style           =   2  'Dropdown List
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   720
            Width           =   1245
         End
         Begin VB.TextBox txt结算号码 
            Height          =   480
            IMEMode         =   3  'DISABLE
            Left            =   1380
            TabIndex        =   10
            Top             =   1815
            Width           =   3225
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
            Left            =   2700
            MaxLength       =   12
            TabIndex        =   6
            Top             =   735
            Width           =   1920
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
            Height          =   465
            Left            =   1380
            TabIndex        =   12
            Top             =   2385
            Width           =   3210
         End
         Begin VB.TextBox txt找补 
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
            Left            =   1380
            Locked          =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   1275
            Width           =   3225
         End
         Begin VB.Label lbl冲预交 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " 预存款"
            Height          =   315
            Left            =   180
            TabIndex        =   2
            Top             =   225
            Width           =   1170
         End
         Begin VB.Label lblPayType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "缴　款"
            Height          =   315
            Left            =   360
            TabIndex        =   4
            Top             =   765
            Width           =   990
         End
         Begin VB.Label lbl结算号码 
            AutoSize        =   -1  'True
            Caption         =   "结算号码"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   90
            TabIndex        =   9
            Top             =   1950
            Width           =   1260
         End
         Begin VB.Label lbl找补 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "找　补"
            Height          =   315
            Left            =   360
            TabIndex        =   7
            Top             =   1350
            Width           =   990
         End
         Begin VB.Label lbl摘要 
            AutoSize        =   -1  'True
            Caption         =   "摘  要"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   390
            TabIndex        =   11
            Top             =   2460
            Width           =   960
         End
      End
      Begin VB.PictureBox picTotal 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         ForeColor       =   &H80000008&
         Height          =   1395
         Left            =   45
         ScaleHeight     =   1365
         ScaleWidth      =   3060
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   90
         Width           =   3090
         Begin XtremeSuiteControls.ShortcutCaption ShortcutCaption1 
            Height          =   450
            Left            =   15
            TabIndex        =   27
            Top             =   30
            Width           =   3045
            _Version        =   589884
            _ExtentX        =   5371
            _ExtentY        =   794
            _StockProps     =   6
            Caption         =   "当前未付"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   15.76
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
         End
         Begin VB.Label lbl剩余自付 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   615
            Left            =   2055
            TabIndex        =   15
            Top             =   585
            Width           =   1005
         End
      End
   End
   Begin VB.TextBox txt合计 
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
      Left            =   6435
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   285
      Width           =   1575
   End
   Begin VB.TextBox txt医保 
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
      Left            =   1500
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   315
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   0
      TabIndex        =   23
      Top             =   900
      Width           =   8100
   End
   Begin VB.Frame fraSplitLeft 
      Caption         =   "Frame3"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4275
      Left            =   8100
      TabIndex        =   20
      Top             =   -180
      Width           =   30
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   510
      Left            =   0
      TabIndex        =   21
      Top             =   6030
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   900
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   3572
            MinWidth        =   882
            Picture         =   "frmChargePaymentWin.frx":08CA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7461
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
            Picture         =   "frmChargePaymentWin.frx":115E
            Key             =   "Calc"
            Object.ToolTipText     =   "计算器:ALT+?"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1693
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
            Object.Width           =   1693
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picBlance 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   2025
      Left            =   90
      ScaleHeight     =   1995
      ScaleWidth      =   11325
      TabIndex        =   22
      Top             =   4095
      Width           =   11355
      Begin VB.CommandButton cmdDel 
         Caption         =   "删除"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   8745
         TabIndex        =   31
         Top             =   60
         Width           =   1080
      End
      Begin VSFlex8Ctl.VSFlexGrid vsBlance 
         Height          =   1815
         Left            =   15
         TabIndex        =   18
         Top             =   495
         Width           =   9915
         _cx             =   17489
         _cy             =   3201
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
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
         BackColorSel    =   16761024
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   7
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmChargePaymentWin.frx":1838
         ScrollTrack     =   0   'False
         ScrollBars      =   3
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
      Begin VB.Label lbl已结 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "已付合计:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   4305
         TabIndex        =   17
         Top             =   120
         Width           =   1365
      End
      Begin VB.Label Label4 
         Caption         =   "本次支付情况"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         TabIndex        =   16
         Top             =   98
         Width           =   2145
      End
   End
   Begin VB.PictureBox pic误差 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   8145
      ScaleHeight     =   1140
      ScaleWidth      =   2040
      TabIndex        =   33
      Top             =   2865
      Width           =   2040
      Begin VB.Label lbl误差额 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0111"
         Height          =   285
         Left            =   135
         TabIndex        =   35
         Top             =   600
         Width           =   1890
      End
      Begin VB.Label lbl误差 
         Caption         =   "本次误差"
         Height          =   315
         Left            =   105
         TabIndex        =   34
         Top             =   90
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "完成收费(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   8220
      TabIndex        =   19
      Top             =   255
      Width           =   2055
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "收费合计"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   5085
      TabIndex        =   14
      Top             =   360
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "医保支付"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Top             =   375
      Width           =   1260
   End
End
Attribute VB_Name = "frmChargePayMentWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum PayChargeType
    EM_正常收费 = 0
    EM_异常作废 = 1
    EM_重新收费 = 2
End Enum
Public Enum ExitMode
    EM_收费完成 = 0
    EM_暂停收费 = 1
    EM_本次作废 = 2
    EM_继续输入 = 3
    EM_退出收费 = 4
End Enum
Private mbytFunc As PayChargeType  '0-收费;1-作废
Private mfrmMain As Object
Private mbytReturnMode As ExitMode
Private mbln异常作废 As Boolean
Private mblnYB退款 As Boolean '医保结算金额大于了单据结算金额
'------------------------------------------------------------------------------------------
'程序入口相关变量
Private mlngModule As Long, mstrPrivs As String
Private mintInsure As Integer, mlng病人ID As Long
Private mlng结算ID As Long, mstr结帐IDs As String
Private mstr冲销IDs  As String  '目前只对异常作废有效
Private mstrNOs As String
Private mstrYBPati As String
Private mstrPatiInfo As String '病人信息
Private mlngShareUseID As Long
Private mstrUseType As String '使用类别
Private mblnOK As Boolean
Private mstr姓名 As String, mstr性别 As String, mstr年龄 As String, mstr费别 As String
Private mbln连续输入 As Boolean
Private mblnCur连续 As Boolean
Private mlngR As Long
Private mlngBrushCardTypeID As Long '在主界面中刷卡的卡类别ID,以便缺省定位在该支付类别上
Private mblnUnloaded  As Boolean
Private mblnLoad As Boolean
'问题:42791
Private mstrBalances As String   '当前的结算额:结算方式:金额:缴款标志(1-缴款;2-找补)|结算方式1:金额1:缴款标志(1-缴款;2-找补)|...
Private mstr退支票 As String
Private mCurCardPay As gTY_PayMoney '本次卡支付
Private mdbl本次应缴 As Double  '本次应缴金额(不包含扣除预交的钱)
Private mcolCardPayMode As Collection
Private Type TY_ChargeMoney
    dbl本次实收 As Double
    dbl本次应收 As Double
    dbl本次医保支付 As Double
    dbl本次已付合计 As Double
    dbl本次冲预交  As Double
    dbl当前未付 As Double
    dbl预交余额 As Double
    dbl费用余额 As Double
    dbl可用预交 As Double
    dbl应缴累计 As Double
    dbl本次误差费 As Double
End Type
Private mCurCarge As TY_ChargeMoney
'------------------------------------------------------------------------------------------
'局部变量
Private mblnFirst As Boolean
Private mblnUnLoad As Boolean '是否Unload窗体
Private mbln已报价 As Boolean
Private mstr医保结算 As String
Private mblnYbBalanced As Boolean '医保已经结算
Private mblnThreeInterface As Boolean '已经调用三方接口
Private mcur个帐余额 As Currency
'----------------------------------------------------------------------------------------------
'医保相关
'当前病人险类的医保支持参数
Private Type TYPE_MedicarePAR
    允许不设置医保项目 As Boolean
    门诊收费存为划价单 As Boolean
    不提醒缴款金额不足 As Boolean    '27536
    门诊必须传递明细 As Boolean
    医保接口打印票据 As Boolean
    医生确定处方类型 As Boolean
    多单据一次结算 As Boolean
    门诊结算作废 As Boolean
    门诊连续收费 As Boolean
    门诊预结算 As Boolean
    多单据收费 As Boolean
    分币处理 As Boolean
    实时监控 As Boolean
    先自付 As Boolean
    全自付 As Boolean
    blnOnlyBjYb As Boolean '本地仅支持北京医保:刘兴洪
    退费后打印回单 As Boolean '
    多单据调一次交易 As Boolean
End Type
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mInsurePara As TYPE_MedicarePAR
Private mrsOneCard As ADODB.Recordset
Private mrsBlance As ADODB.Recordset
Private mdbl缴款金额 As Double, mdbl找补 As Double
'---------------------------------------------------------------------------------
Private mbln连续收费 As Boolean
'---------------------------------------------------------------------------------
Private mdbl现金 As Double, mdbl原未付 As Double
Private Const VK_RETURN = &HD
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private mblnCacheKeyReturn As Boolean   '41025:是否缓存了回车键,可能存在在收费界面刷卡中本身包含了回车,因此需要判断
Public Event zlSaveData(ByRef lng结算序号 As Long, ByRef str结帐IDs As String, ByRef strSaveNos As String, ByRef blnNotCommit As Boolean, ByRef blnCancel As Boolean)
Private mrsClassMoney As ADODB.Recordset
Private mcllSquareBalance As Collection '消费卡结算信息
Private mcllCurSquareBalance As Collection '当前消费卡刷卡信息
Private mblnNotChange As Boolean
Private mstrTitle As String '用于窗体个性化保存的窗体名

Private Sub zlInitTotalStru()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化汇总金额
    '编制:刘兴洪
    '日期:2011-12-26 13:19:04
    '问题:44944
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mbln连续输入 And Not grsTotal Is Nothing Then Exit Sub
    Set grsTotal = New ADODB.Recordset
    grsTotal.Fields.Append "性质", adBigInt, , adFldIsNullable
    grsTotal.Fields.Append "结算方式", adVarChar, 60, adFldIsNullable
    grsTotal.Fields.Append "结算金额", adDouble, , adFldIsNullable
    grsTotal.CursorLocation = adUseClient
    grsTotal.LockType = adLockOptimistic
    grsTotal.CursorType = adOpenStatic
    grsTotal.Open
End Sub

Private Sub WhriteTotalDataToReCord(ByVal dbl预交 As Double, _
    ByVal dblMoney As Double, ByVal dbl退支票 As Double)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:数据保存,以便累计汇总数据
    '编制:刘兴洪
    '日期:2011-12-26 22:25:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, str结算方式 As String, dbl缴款 As Double, dbl找补 As Double
    Dim int性质 As Integer
    If grsTotal Is Nothing Then Call zlInitTotalStru
    If grsTotal.State <> 1 Then Call zlInitTotalStru
    
    If (mCurCardPay.int性质 = 1 Or mCurCardPay.int性质 = 2) And mblnCur连续 = False Then
        dbl缴款 = Val(txt缴款.Text)
        dbl找补 = Val(txt找补.Text)
    End If
    If dbl缴款 = 0 Then
        dbl缴款 = 0: dbl找补 = 0
    End If
    On Error GoTo errHandle
    
    With vsBlance
        If grsTotal.RecordCount <> 0 Then grsTotal.MoveFirst
        If dbl缴款 <> 0 Then
            grsTotal.Find "结算方式='本次缴款'", , adSearchForward, 1
            If grsTotal.EOF Then grsTotal.AddNew
            grsTotal!性质 = 0
            grsTotal!结算方式 = "缴款"
            grsTotal!结算金额 = dbl缴款
        End If
        
        If dbl找补 <> 0 Then
            grsTotal.Find "结算方式='" & IIf(mCurCardPay.bln支票, "退支票", "找补") & "'", , adSearchForward, 1
            If grsTotal.EOF Then grsTotal.AddNew
            grsTotal!性质 = 1
            grsTotal!结算方式 = IIf(mCurCardPay.bln支票, "退支付", "找补")
            grsTotal!结算金额 = dbl找补
        End If
        
        For i = 1 To .Rows - 1
            str结算方式 = Trim(.TextMatrix(i, .ColIndex("支付方式")))
            int性质 = Val(.RowData(i))
            If str结算方式 <> "" Then
                '.rowdata:0-普通的结算方式-1-医保结算;2-三方接口交易;3-一卡通结算;4-预存款
                '性质:0-缴款;1-找补,2-冲预交;其他(mod 10:0-普通结算;1-医保结算;2-三方接品;3-一卡通)
                grsTotal.Find "结算方式='" & str结算方式 & "'", , adSearchForward, 1
                If grsTotal.EOF Then grsTotal.AddNew
                grsTotal!性质 = IIf(int性质 + 10 = 14, 2, int性质 + 10)
                grsTotal!结算方式 = str结算方式
                grsTotal!结算金额 = Val(Nvl(grsTotal!结算金额)) + Val(.TextMatrix(i, .ColIndex("支付金额")))
                grsTotal.Update
            End If
        Next
        
        If dbl预交 <> 0 Then
            grsTotal.Find "结算方式='预存款'", , adSearchForward, 1
            If grsTotal.EOF Then grsTotal.AddNew
            grsTotal!性质 = 2
            grsTotal!结算方式 = "预存款"
            grsTotal!结算金额 = Val(Nvl(grsTotal!结算金额)) + dbl预交
            grsTotal.Update
        End If
        If mCurCardPay.bln消费卡 Then
            For i = 1 To mcllCurSquareBalance.Count
                '当前刷卡数据(array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文))
                grsTotal.Find "结算方式='" & mCurCardPay.str结算方式 & "'", , adSearchForward, 1
                If grsTotal.EOF Then grsTotal.AddNew
                grsTotal!性质 = IIf(mCurCardPay.blnOneCard, 13, 12)
                grsTotal!结算方式 = mCurCardPay.str结算方式
                grsTotal!结算金额 = Val(Nvl(grsTotal!结算金额)) + Val(mcllCurSquareBalance(i)(2))
                grsTotal.Update
            Next
        Else
            grsTotal.Find "结算方式='" & mCurCardPay.str结算方式 & "'", , adSearchForward, 1
            If grsTotal.EOF Then grsTotal.AddNew
            ''1-现金结算方式,2-其他非医保结算,3-医保个人帐户,4-医保各类统筹,5-代收款项,6-费用折扣,7-一卡通结算,8-结算卡结算;<0 表示第三方支付
            '性质:0-缴款;1-找补,2-冲预交;其他(mod 10:0-普通结算;1-医保结算;2-三方接品;3-一卡通)
            Select Case mCurCardPay.int性质
            Case 1, 2
                grsTotal!性质 = 10
            Case 3, 4
                grsTotal!性质 = 11
            Case 7, 8
                grsTotal!性质 = IIf(mCurCardPay.blnOneCard, 13, 12)
            Case Else
                grsTotal!性质 = 10
            End Select
            grsTotal!结算方式 = mCurCardPay.str结算方式
            grsTotal!结算金额 = Val(Nvl(grsTotal!结算金额)) + dblMoney
            grsTotal.Update
            If dbl退支票 <> 0 Then
                grsTotal.Find "结算方式='" & mstr退支票 & "'", , adSearchForward, 1
                If grsTotal.EOF Then grsTotal.AddNew
                grsTotal!性质 = 10
                grsTotal!结算方式 = mstr退支票
                grsTotal!结算金额 = Val(Nvl(grsTotal!结算金额)) + dbl退支票
                grsTotal.Update
            End If
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub initInsure()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化医保参数
    '编制:刘兴洪
    '日期:2011-08-21 18:55:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mintInsure = 0 Then Exit Sub
'    mInsurePara.允许不设置医保项目 = gclsInsure.GetCapability(support允许不设置医保项目, mlng病人ID, mintInsure)
'    mInsurePara.门诊收费存为划价单 = gclsInsure.GetCapability(support门诊收费存为划价单, mlng病人ID, mintInsure)
'    mInsurePara.门诊必须传递明细 = gclsInsure.GetCapability(support门诊必须传递明细, mlng病人ID, mintInsure)
'    mInsurePara.医生确定处方类型 = gclsInsure.GetCapability(support医生确定处方类型, mlng病人ID, mintInsure)
     mInsurePara.医保接口打印票据 = gclsInsure.GetCapability(support医保接口打印票据, mlng病人ID, mintInsure)
    mInsurePara.多单据一次结算 = gclsInsure.GetCapability(support多单据一次结算, mlng病人ID, mintInsure)
    mInsurePara.门诊连续收费 = gclsInsure.GetCapability(support门诊连续收费, mlng病人ID, mintInsure)
    '刘兴洪:27536 20100119
    mInsurePara.不提醒缴款金额不足 = gclsInsure.GetCapability(support不提醒缴款金额不足, mlng病人ID, mintInsure)
    mInsurePara.门诊结算作废 = gclsInsure.GetCapability(support门诊结算作废, , mintInsure)
    mInsurePara.多单据收费 = gclsInsure.GetCapability(support多单据收费, mlng病人ID, mintInsure)
    mInsurePara.门诊预结算 = gclsInsure.GetCapability(support门诊预算, mlng病人ID, mintInsure)
    mInsurePara.分币处理 = gclsInsure.GetCapability(support分币处理, mlng病人ID, mintInsure)
'    mInsurePara.先自付 = gclsInsure.GetCapability(support收费帐户首先自付, mlng病人ID, mintInsure)
'    mInsurePara.全自付 = gclsInsure.GetCapability(support收费帐户全自费, mlng病人ID, mintInsure)
'    mInsurePara.实时监控 = gclsInsure.GetCapability(support实时监控, mlng病人ID, mintInsure)
    'mInsurePara.退费后打印回单 = gclsInsure.GetCapability(support退费后打印回单, mlng病人ID, mintInsure)
     mInsurePara.多单据调一次交易 = gclsInsure.GetCapability(support门诊_不分单据结算, mlng病人ID, mintInsure)
End Sub
Private Sub InitBalanceData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化结算数据
    '编制:刘兴洪
    '日期:2012-02-05 16:02:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call ClearBanalce
    With mCurCarge
          .dbl本次实收 = mfrmMain.zlGetToTatal
          .dbl本次医保支付 = mfrmMain.GetMedicareSum
          .dbl本次已付合计 = 0
          .dbl本次应收 = mfrmMain.GetBillSum(True)
          .dbl当前未付 = .dbl本次实收 - .dbl本次医保支付
          .dbl本次冲预交 = 0
          .dbl本次误差费 = 0
      End With
      '保存预结算未付金额，用于与结算结果进行比较，确定是否重复报价
      mdbl原未付 = mCurCarge.dbl当前未付
End Sub
Private Sub ClearBanalce()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除结算数据
    '编制:刘兴洪
    '日期:2012-02-05 16:02:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With mCurCarge
        .dbl本次实收 = 0
        .dbl本次医保支付 = 0
        .dbl本次已付合计 = 0
        .dbl本次应收 = 0
        .dbl当前未付 = 0
        .dbl本次冲预交 = 0
        .dbl本次误差费 = 0
    End With
    With vsBlance
        .Clear 1: .Rows = 2
    End With
    txt医保.Text = "0.00"
    txt合计.Text = "0.00"
End Sub

Private Sub LoadData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载数据
    '编制:刘兴洪
    '日期:2011-08-20 19:49:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim i As Long, bln消费卡 As Boolean, lng卡类别ID As Long
    Dim strCardNo As String
    Dim blnYb As Boolean
    
    On Error GoTo errHandle
    
    Call ClearBanalce
 
    gstrSQL = "" & _
    "   Select  A.ID,Mod(A.记录性质,10) as 记录性质,A.结算方式,A.冲预交,A.摘要,A.卡类别ID,A.结算卡序号, " & _
    "               A.结算号码,A.卡号,A.交易流水号,nvl(C.是否自制,0) as 自制卡, " & _
    "               nvl(C.是否退现,0) as 是否退现, " & _
    "               nvl(C.是否全退,0) as 是否全退, " & _
    "               decode(C.卡号密文,NULL,0,1) as  是否密文," & _
    "               C.名称 as 卡类别名称,A.交易说明,A.结算序号,A.校对标志,decode(B.名称,Null,0,1) as 医保,0 as 消费卡id" & _
    "   From 病人预交记录 A ,医疗卡类别 C" & IIf(mbln异常作废, ",Table( f_Num2list( [3])) Q ", "") & _
    "           ,(Select 名称 From 结算方式 where 性质 in (3,4)) B" & _
    "   Where  " & IIf(mbln异常作废, "A.结帐ID=Q.Column_Value", "A.结算序号 = [1] ") & _
    "                And A.卡类别ID=C.ID(+) And A.结算方式=B.名称(+) And nvl(A.结算卡序号,0)=0"
    
 gstrSQL = gstrSQL & " Union ALL " & _
    "   Select   A.ID,Mod(A.记录性质,10) as 记录性质,A.结算方式,-1*nvl(b.应收金额,0) as 冲预交,A.摘要,A.卡类别ID,A.结算卡序号, " & _
    "           A.结算号码,B.卡号,B.交易流水号,nvl( M.自制卡,0) as 自制卡, " & _
    "           nvl( M.是否退现,0) as 是否退现, " & _
    "           nvl(M.是否全退,0) as 是否全退, " & _
    "           nvl(M.是否密文,0) as  是否密文," & _
    "           M.名称 as 卡类别名称,A.交易说明,A.结算序号,A.校对标志,0 as 医保,B.消费卡id" & _
    "   From 病人预交记录 A ,病人卡结算记录 B, " & _
    "              消费卡类别目录 M" & IIf(mbln异常作废, ",Table( f_Num2list( [3])) Q ", "") & _
    "   Where  a.Id = b.结算id And a.结算卡序号 = m.编号  " & _
                  IIf(mbln异常作废, "And A.结帐ID=Q.Column_Value", " And A.结算序号 = [1] ")
   gstrSQL = "" & _
   "    Select   /*+ rule */    记录性质,结算方式,摘要,卡类别ID,卡类别名称,自制卡,结算卡序号,结算号码,卡号,交易流水号, 交易说明,结算序号,校对标志,医保,消费卡id," & _
   "               max(是否密文) as 是否密文,max(是否全退) as 是否全退,max(是否退现) as 是否退现 , nvl(sum(冲预交),0) as 冲预交" & _
   "    From (" & gstrSQL & ") " & _
   "   Group by 记录性质,结算方式,摘要,卡类别ID,卡类别名称,自制卡,结算卡序号,结算号码,卡号,交易流水号, 交易说明,结算序号,校对标志,医保,消费卡id"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng结算ID, IIf(mbln异常作废, 2, 1), mstr结帐IDs)
    With rsTemp
        i = 1
        blnYb = False
        Do While Not .EOF
            If Nvl(rsTemp!摘要) = "保险结算" Or Nvl(rsTemp!医保) = "1" Then
                    mCurCarge.dbl本次医保支付 = RoundEx(mCurCarge.dbl本次医保支付 + Nvl(rsTemp!冲预交, 0), 6)
                    blnYb = True
            End If
            If Val(Nvl(rsTemp!校对标志, 0)) = 2 Then
                With vsBlance
                    If .TextMatrix(i, .ColIndex("支付方式")) <> "" Then
                        .Rows = .Rows + 1
                        i = i + 1
                    End If
                    .RowData(i) = 0
                    strCardNo = Nvl(rsTemp!卡号)
                    lng卡类别ID = Val(Nvl(rsTemp!结算卡序号))
                    bln消费卡 = lng卡类别ID <> 0
                    If bln消费卡 Then
                        If mcllSquareBalance Is Nothing Then Set mcllSquareBalance = New Collection
                        'array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文
                        mcllSquareBalance.Add Array(lng卡类别ID, Val(Nvl(rsTemp!消费卡ID)), _
                        Format(Val(Nvl(rsTemp!冲预交)), "0.00"), strCardNo, "", "", Val(Nvl(rsTemp!是否密文)))
                    End If
                    
                    If Not bln消费卡 Then lng卡类别ID = Val(Nvl(rsTemp!卡类别ID))
                    
                    If lng卡类别ID <> 0 Then .RowData(i) = 2
                    If lng卡类别ID <> 0 Then
                        strCardNo = gobjSquare.objSquareCard.zlGetCardNODencode(strCardNo, lng卡类别ID, bln消费卡)
                    End If
                    If Nvl(rsTemp!摘要) = "保险结算" Or Val(Nvl(rsTemp!医保)) = 1 Then
                        .RowData(i) = 1 '医保交易
                        If InStr(1, mstr医保结算, "," & Nvl(rsTemp!结算方式)) = 0 Then
                            mstr医保结算 = mstr医保结算 & "," & Nvl(rsTemp!结算方式)
                        End If
                    ElseIf lng卡类别ID <> 0 Then
                        '三方接口交易
                        .RowData(i) = 2 '三方接口交易
                    Else
                        '是否一卡通交易
                        mrsOneCard.Filter = "结算方式='" & Nvl(rsTemp!结算方式) & "'"
                        If Not mrsOneCard.EOF Then
                            .RowData(i) = 3 '一卡通交易
                        End If
                        mrsOneCard.Filter = 0
                    End If
                    
                    .TextMatrix(i, .ColIndex("支付方式")) = Nvl(rsTemp!结算方式)
                    ' 医疗卡类别ID|消费卡(1, 0) |自制卡|是否全退|是否退现|接口名称
                    .Cell(flexcpData, i, .ColIndex("支付方式")) = lng卡类别ID & "|" & IIf(bln消费卡, 1, 0) & "|" & Val(Nvl(rsTemp!自制卡)) & "|" & Val(Nvl(rsTemp!是否全退)) & "|" & Val(Nvl(rsTemp!是否退现)) & "|" & Nvl(rsTemp!卡类别名称)
                    
                    .TextMatrix(i, .ColIndex("支付金额")) = Format(Val(Nvl(rsTemp!冲预交)), "0.00")
                    .TextMatrix(i, .ColIndex("结算号码")) = Nvl(rsTemp!结算号码)
                    .TextMatrix(i, .ColIndex("备注")) = Nvl(rsTemp!摘要)
                    .TextMatrix(i, .ColIndex("交易流水号")) = Nvl(rsTemp!交易流水号)
                    .TextMatrix(i, .ColIndex("交易说明")) = Nvl(rsTemp!交易说明)
                    .TextMatrix(i, .ColIndex("卡号")) = IIf(Val(Nvl(rsTemp!是否密文)) = 1, String(Len(strCardNo), "*"), strCardNo)
                    .Cell(flexcpData, i, .ColIndex("卡号")) = Nvl(rsTemp!卡号)
      
                    mCurCarge.dbl本次已付合计 = RoundEx(mCurCarge.dbl本次已付合计 + Val(Nvl(rsTemp!冲预交)), 6)
                End With
            ElseIf Val(Nvl(rsTemp!记录性质)) = 1 Or Val(Nvl(rsTemp!记录性质)) = 11 Then
                mCurCarge.dbl本次冲预交 = RoundEx(mCurCarge.dbl本次冲预交 + Val(Nvl(rsTemp!冲预交)), 6)
                mCurCarge.dbl本次已付合计 = RoundEx(mCurCarge.dbl本次已付合计 + Val(Nvl(rsTemp!冲预交)), 6)
            End If
            .MoveNext
        Loop
    End With

                    
    If mbln异常作废 Then
         gstrSQL = "" & _
         "   Select /*+ rule */ B.NO,B.结帐ID, Nvl(Sum(Nvl(B.应收金额, 0)), 0)  As 本次应收合计, " & _
         "       Nvl(Sum(Nvl(B.实收金额, 0)), 0)  As 本次实收合计 " & _
         "   From 门诊费用记录 B , Table( f_Num2list( [2])) Q  " & _
        "    Where B.结帐ID=Q.Column_Value " & _
        "    Group by B.NO,B.结帐ID"
    Else
         gstrSQL = "" & _
         "   Select  /*+ rule */ B.NO,B.结帐ID, Nvl(Sum(Nvl(B.应收金额, 0)), 0)  As 本次应收合计, " & _
         "       Nvl(Sum(Nvl(B.实收金额, 0)), 0)  As 本次实收合计 " & _
         "   From 门诊费用记录 B  " & _
        "    Where B.结帐id in (Select 结帐ID From 病人预交记录 where 结算序号=[1] )  " & _
        "    Group by B.NO,B.结帐ID"
    End If
   Set mrsBlance = Nothing
   Set mrsBlance = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng结算ID, mstr结帐IDs)
   With mCurCarge
         .dbl本次实收 = 0:
         .dbl本次应收 = 0
        Do While Not mrsBlance.EOF
            .dbl本次实收 = RoundEx(.dbl本次实收 + Val(Nvl(mrsBlance!本次实收合计)), 6)
            .dbl本次应收 = RoundEx(.dbl本次应收 + Val(Nvl(mrsBlance!本次应收合计)), 6)
            mrsBlance.MoveNext
        Loop
        .dbl当前未付 = RoundEx(.dbl本次实收 - .dbl本次已付合计, 6)
        If .dbl本次冲预交 <> 0 Then
            With vsBlance
                If .Rows = 2 Then .Row = 1
                If .Row < 0 Then .Row = 1
                i = .Row
                If Trim(.TextMatrix(.Row, .ColIndex("支付方式"))) <> "" Then
                    .Rows = .Rows + 1
                    i = .Rows - 1
                End If
                .TextMatrix(i, .ColIndex("支付方式")) = "预存款"
                .RowData(i) = 4
                .TextMatrix(i, .ColIndex("支付金额")) = Format(mCurCarge.dbl本次冲预交, "0.00")
            End With
        End If
        mblnYB退款 = mCurCarge.dbl当前未付 < 0 And blnYb
   End With
   
   vsBlance_AfterRowColChange 0, 0, vsBlance.Row, vsBlance.Col
   Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Public Function zlChargeWin(ByVal frmMain As Object, ByVal bytFunc As PayChargeType, _
    ByVal lngModule As Long, ByVal strPrivs As String, _
    ByVal lngShareUseID As Long, ByVal strUseType As String, _
    ByVal lng结算ID As Long, ByVal str结帐IDs As String, _
    ByVal strNos As String, _
    ByVal lng病人ID As Long, ByVal intInsure As Integer, _
    Optional ByVal str姓名 As String = "", Optional ByVal str性别 As String, _
    Optional str年龄 As String, Optional str费别 As String = "", _
    Optional dbl缴款金额 As Double, Optional dbl找补 As Double, _
    Optional bytReturnMode As ExitMode = EM_收费完成, _
    Optional dbl应缴累计 As Double, _
    Optional bln继续输入 As Boolean, _
    Optional lngBrushCardTypeID As Long = 0, _
    Optional dbl本次应缴 As Double = 0, _
    Optional strBalance As String = "", Optional bln异常作废 As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入口:表示进入支付结算窗口
    '入参:frmMain-调用的主窗体
    '       bytFunc-0-收费;1-作废
    '       lngModule -模块号
    '       strPrivs-权限串
    '       mlng结算ID:多单据结算时,以关联的结算ID为准.否则为结帐Id
    '       strNos-单据号:以逗号分离,如"AAAA,BBBBB"
    '       dblPayMoney-本次消费总额
    '       dblYbMoney-医保支付金额
    '       lngBrushCardTypeID-缺省的刷卡类别ID
    '       bln异常作废-异常单据作废处理(异常作废时传入):如果为true,表示针对作废的异常单据进行作废
    '出参:dbl缴款金额-输入的缴款金额和找补金额(缴现金时,传出)
    '        bln继续输入-是否继续录入的票据
    '        bytReturnMode-返回操作模式(0-正常收费完成,1-暂停收费;2-本次作废收费;3-继续输入)
    '        dbl本次应缴-医保病人,在连续收费情况下,需要重新计算本次的应缴额
    '       strBalance-返回本次收费的结算方式,格式如下:
    '                       金额:缴款标志(1-缴款;2-找补)|结算方式1:金额1:缴款标志(1-缴款;2-找补)|...
    '返回:完成收费,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-12 09:59:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set mrsClassMoney = Nothing
    mblnYbBalanced = False: mblnThreeInterface = False: mblnOK = False
    mlngBrushCardTypeID = lngBrushCardTypeID: Set mfrmMain = frmMain
    mintInsure = intInsure: mlngShareUseID = lngShareUseID: mstrUseType = strUseType
    mlng结算ID = lng结算ID: mlng病人ID = lng病人ID: mstrPrivs = strPrivs
    mstr冲销IDs = "": mstr结帐IDs = str结帐IDs: mlngModule = lngModule
    
    mstr姓名 = str姓名: mstr性别 = str性别: mstr年龄 = str年龄: mstr费别 = str费别
    mstrNOs = strNos: mdbl本次应缴 = 0: mbln异常作废 = bln异常作废
    mstrPatiInfo = str姓名
   ' mstrPatiInfo = mstrPatiInfo & "性别:" & str性别 & Space(4)
    'mstrPatiInfo = mstrPatiInfo & "年龄:" & str年龄 & Space(4)
    'mstrPatiInfo = mstrPatiInfo & "费别:" & str费别 & Space(4)
    mdbl缴款金额 = 0: mdbl找补 = 0: mblnUnLoad = False: mblnUnloaded = False
    mCurCarge.dbl应缴累计 = dbl应缴累计
    mbln连续输入 = dbl应缴累计 <> 0
    mstrBalances = ""
    mbytFunc = bytFunc: mbytReturnMode = EM_收费完成
    
    If bln异常作废 Then
        mstr冲销IDs = mstr结帐IDs
        mstr结帐IDs = zlGet原结帐IDs(mstr冲销IDs)
    End If
    
    mblnOK = False
    Me.Show 1, frmMain
    bln继续输入 = mbln连续输入: dbl本次应缴 = mdbl本次应缴
    dbl缴款金额 = mdbl缴款金额: dbl找补 = mdbl找补
    strBalance = mstrBalances
    bytReturnMode = mbytReturnMode
    zlChargeWin = mblnOK
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub InitFace()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化控件
    '编制:刘兴洪
    '日期:2011-06-13 14:09:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, dbl金额 As Double, rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    

    With vsBlance
        .Cell(flexcpFontBold, 1, 0, 1, .COLS - 1) = True
        .Clear 1: .Rows = 2
    End With
    With mCurCarge
        .dbl本次冲预交 = 0
        .dbl本次实收 = 0
        .dbl本次医保支付 = 0
        .dbl本次已付合计 = 0
        .dbl本次应收 = 0
        .dbl当前未付 = 0
        .dbl费用余额 = 0
        .dbl可用预交 = 0
        .dbl预交余额 = 0
    End With
   With mCurCardPay
        .lng消费卡ID = 0
        .str限制类别 = ""
        .dbl已刷金额 = 0
        .str刷卡卡号 = ""
        .str刷卡密码 = ""
    End With
    mstr退支票 = ""
    strSQL = " " & _
    "         Select B.名称 " & _
    "         From 结算方式应用 A, 结算方式 B " & _
    "         Where A.应用场合 = '收费' And B.名称 = A.结算方式 And Nvl(B.应付款, 0) = 1 And a.付款方式 Is Null And Rownum <= 1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTemp.EOF Then
        mstr退支票 = Nvl(rsTemp!名称)
    End If
    Call initInsure
    If mbytFunc = EM_正常收费 Then
        Call InitBalanceData
    Else
        Call LoadData
    End If
    Call Load支付方式: Call LoadPatiInfor
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub SetControlProperty(Optional bln预交 As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件属性
    '参数:bln预交-是否正在输入预交
    '编制:刘兴洪
    '日期:2011-08-12 10:43:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngTop As Long, sngSplitHeight As Single, dbl现金 As Double
    Dim bln分币 As Boolean, dblMoney As Double
    Dim bln退款 As Boolean '主要是医保相关结算大于了单据收费
    
    sngSplitHeight = 80
    
    '51670
    If mlng病人ID = 0 Or mbln连续输入 Then
        lbl冲预交.Visible = False
        txt冲预交.Visible = False
        txt冲预交.Text = "0"
    End If
    
    cmdNext.Visible = Val(txt冲预交.Text) = 0 And Val(txt缴款.Text) = 0 And mbytFunc = EM_正常收费 And _
        (mCurCarge.dbl本次已付合计 - mCurCarge.dbl本次医保支付) = 0 _
        And (gTy_Module_Para.byt缴款控制 = 1 Or gTy_Module_Para.byt缴款控制 = 3) And _
        mCurCardPay.lng医疗卡类别ID = 0 And mCurCardPay.blnOneCard = False _
        And (mintInsure = 0 Or mintInsure <> 0 And mblnYbBalanced)
        
    lbl已结.Caption = "已付合计:" & Format(mCurCarge.dbl本次已付合计, "###0.00;-###0.00;0.00;0.00;")
    
    If mCurCardPay.int性质 = 1 And bln预交 = False Then
        dblMoney = mCurCarge.dbl当前未付 + mCurCarge.dbl应缴累计
        If mintInsure > 0 Then  '问题:43855,44069
            If mInsurePara.分币处理 Then
                bln分币 = True
                dbl现金 = CentMoney(CCur(dblMoney))
            Else
                dbl现金 = Format(dblMoney, "0.00")
            End If
        Else
             bln分币 = True
            dbl现金 = RoundEx(CentMoney(CCur(dblMoney)), 6)
        End If
        lbl剩余自付.Caption = Format(dbl现金, "0.00")
    Else
        lbl剩余自付.Caption = Format(mCurCarge.dbl当前未付 + mCurCarge.dbl应缴累计, "0.00")
    End If
    
    '问题:58344
    '   检查是否当前支付金额为负数,是负数时,需要提醒操作员(主要是医保结算时可能大于本身单据的费用)
    If Not mblnYB退款 Then
        lblPayType.Caption = "缴　款"
        lblPayType.ForeColor = &H80000008
        cbo支付方式.ForeColor = &H80000008
        txt缴款.ForeColor = &H80000008
    Else
        lblPayType.Caption = "退　款"
        lblPayType.ForeColor = vbRed
        cbo支付方式.ForeColor = vbRed
        txt缴款.ForeColor = vbRed
        '退款时，不处理预交
        txt冲预交.Visible = False: lbl冲预交.Visible = False
        txt冲预交.Text = 0
    End If
    
    If bln预交 Then
        '预交的处理
        lbl找补.Visible = False: txt找补.Visible = False
        txt找补.Text = 0
    ElseIf mCurCardPay.int性质 = 1 Then
        lbl找补.Visible = True: txt找补.Visible = True
        lbl找补.Caption = "找　补"
        If IIf(mblnYB退款 < 0, -1, 1) * Val(txt缴款.Text) >= dbl现金 Then
            lbl找补.ForeColor = &H80000008
            txt找补.ForeColor = &H80000008
        Else
            lbl找补.ForeColor = vbRed
            txt找补.ForeColor = vbRed
        End If
        
        If bln分币 Then
            dblMoney = CentMoney(CCur(mCurCarge.dbl当前未付))
        Else
            dblMoney = mCurCarge.dbl当前未付
        End If
        '61611
        'IIf(mblnYB退款, -1, 1) * (IIf(mblnYB退款, -1, 1) * Val(txt缴款.Text) - dblMoney - mCurCarge.dbl应缴累计), "0.00")
        txt找补.Text = Format(Val(txt缴款.Text) - dblMoney - mCurCarge.dbl应缴累计, "0.00")
        txt结算号码.Visible = False: lbl结算号码.Visible = False
        
    ElseIf mCurCardPay.bln支票 Then
        If mblnYB退款 Then
            '58344
            lbl找补.Visible = False
            txt找补.Visible = False
            txt找补.Text = 0
        Else
            If RoundEx(Val(txt缴款.Text), 6) > RoundEx(mCurCarge.dbl当前未付 + mCurCarge.dbl应缴累计, 6) Then
                  lbl找补.Visible = True: txt找补.Visible = True
                  lbl找补.Caption = "  退 支 票"
                  txt找补.Text = Format(Val(txt缴款.Text) - RoundEx(mCurCarge.dbl当前未付, 2) - mCurCarge.dbl应缴累计, "0.00")
                  txt找补.ForeColor = vbRed
                  lbl找补.ForeColor = vbRed
            Else
                  lbl找补.Visible = False
                  txt找补.Visible = False
                  txt找补.Text = 0
            End If
        End If
         txt结算号码.Visible = True
         lbl结算号码.Visible = True
    ElseIf cbo支付方式.Text Like "*卡*" And mCurCardPay.lng医疗卡类别ID = 0 Then
         txt结算号码.Visible = True
         lbl结算号码.Visible = True
        lbl找补.Visible = False
        txt找补.Visible = False
        txt找补.Text = 0
    Else
        lbl找补.Visible = False
        txt找补.Visible = False
        txt结算号码.Visible = False: lbl结算号码.Visible = False
    End If
    sngTop = txt冲预交.Top
    If txt冲预交.Visible Then
        sngTop = txt冲预交.Top + txt冲预交.Height + sngSplitHeight
    End If
    cbo支付方式.Top = sngTop
    txt缴款.Top = sngTop
    lblPayType.Top = sngTop + (cbo支付方式.Height - lblPayType.Height) \ 2
    sngTop = sngTop + cbo支付方式.Height + sngSplitHeight
    If lbl找补.Visible Then
        txt找补.Top = sngTop
        lbl找补.Top = sngTop + (txt找补.Height - lbl找补.Height) \ 2
        sngTop = sngTop + txt找补.Height + sngSplitHeight
    End If
    If txt结算号码.Visible Then
        txt结算号码.Top = sngTop
        lbl结算号码.Top = sngTop + (txt结算号码.Height - lbl结算号码.Height) \ 2
        sngTop = sngTop + txt结算号码.Height + sngSplitHeight
    End If
     
    txt摘要.Top = sngTop
    lbl摘要.Top = sngTop + 25
    txt摘要.Height = picPay.Height - sngTop - 100
    If mbytFunc = 1 Then
        txt冲预交.BackColor = Me.BackColor
        txt缴款.BackColor = Me.BackColor
        txt结算号码.BackColor = Me.BackColor
        txt摘要.BackColor = Me.BackColor
        cbo支付方式.BackColor = Me.BackColor
        txt找补.BackColor = Me.BackColor
        txt找补.Text = ""
    End If
 
End Sub
Private Sub cbo支付方式_Click()
    Dim i As Long, varData As Variant, varTemp As Variant
    Dim lngIndex As Long
    If mblnFirst Then Exit Sub
    txt缴款.Text = ""
    With mCurCardPay
        .lng医疗卡类别ID = 0
        .bln消费卡 = False
        .str结算方式 = ""
        .lng消费卡ID = 0
        .str名称 = ""
        .str刷卡卡号 = ""
        .str刷卡密码 = ""
        .lngID = 0
        .strNo = ""
        .str名称 = ""
        .bln卡号密文 = False
        .int医疗卡长度 = 0
        .bln读卡 = False
        .bln支票 = False
        .blnOneCard = False
        .int性质 = 0
        .bln自制卡 = False
     End With
    With cbo支付方式
        If .ListIndex = -1 Then GoTo SetProperty:
        lngIndex = .ListIndex + 1
        mCurCardPay.int性质 = .ItemData(.ListIndex)
        mCurCardPay.blnOneCard = .ItemData(.ListIndex) = 7
        mCurCardPay.bln支票 = False
        If .ItemData(.ListIndex) = 2 And cbo支付方式.Text Like "*支票*" Then
             mCurCardPay.bln支票 = True
        End If
    End With
    '短|全名|刷卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式|密文规则|是否自制卡;…
    If Not mcolCardPayMode Is Nothing Then
        With mCurCardPay
            .lng医疗卡类别ID = Val(mcolCardPayMode(lngIndex)(3))
            .bln消费卡 = Val(mcolCardPayMode(lngIndex)(5)) = 1
            .str结算方式 = Trim(mcolCardPayMode(lngIndex)(6))
            .str名称 = Trim(mcolCardPayMode(lngIndex)(1))
            .bln读卡 = Val(mcolCardPayMode(lngIndex)(2)) = 0
            If .lng医疗卡类别ID <> 0 Then .bln支票 = False: .blnOneCard = False
            .bln自制卡 = Val(mcolCardPayMode(lngIndex)(8)) = 1
            .bln卡号密文 = Trim(mcolCardPayMode(lngIndex)(7)) <> "" And Trim(mcolCardPayMode(lngIndex)(7)) <> "0"
            If .bln消费卡 Or (.int性质 <> 1 And mblnYB退款) Then
                '57682:缺省为所有支付金额
                txt缴款.Text = Format(IIf(mblnYB退款, -1, 1) * Val(lbl剩余自付.Caption), "0.00")
            End If
         End With
     Else
         mCurCardPay.str结算方式 = zlstr.NeedName(cbo支付方式.Text)
     End If
     If mCurCardPay.blnOneCard Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = New clsICCard
            Call mobjICCard.SetParent(Me.hWnd)
            Set mobjICCard.gcnOracle = gcnOracle
        End If
     End If
SetProperty:
     Call SetControlProperty
     If txt缴款.Enabled Then txt缴款.SetFocus
End Sub
Private Function CheckOneCard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查一卡通是否正确
    '返回:一卡通验证正确或非一卡通,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-23 17:07:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim CurOneCard As Currency, dblMoney As Double
    
    If mCurCardPay.blnOneCard = False Then CheckOneCard = True: Exit Function
    If mobjICCard Is Nothing Then
        Set mobjICCard = New clsICCard
        Call mobjICCard.SetParent(Me.hWnd)
        Set mobjICCard.gcnOracle = gcnOracle
    End If
    If mobjICCard Is Nothing Then
        MsgBox "一卡通接口创建失败!", vbOKOnly, gstrSysName
        Exit Function
    End If
    '弹出刷卡界面
    'zlBrushCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal rsClassMoney As ADODB.Recordset, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal bln消费卡 As Boolean, _
    '    ByVal strPatiName As String, ByVal strSex As String, _
    '    ByVal strOld As String, ByVal dbl金额 As Double, _
    '    Optional ByRef strCardNo As String, _
    '    Optional ByRef strPassWord As String, _
    '    Optional ByRef bln退费 As Boolean = False, _
    '    Optional ByRef blnShowPatiInfor As Boolean = False, _
    '    Optional ByRef bln退现 As Boolean = False, _
    '    Optional ByVal bln余额不足禁止 As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据指定支付类别,弹出刷卡窗口
    '入参:rsClassMoney:收费类别,金额
    '        lngCardTypeID-为零时,为老一卡通刷卡
    '       bln余额不足禁止-目前只针对消费卡,表示余额不足时,禁止继续操作,否则用余额进行支付
    dblMoney = Val(txt缴款.Text)
    If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModule, Nothing, 0, False, _
    mstr姓名, mstr性别, mstr年龄, dblMoney, mCurCardPay.str刷卡卡号, mCurCardPay.str刷卡密码, _
    False, True, False, False) = False Then Exit Function
 
    CurOneCard = mobjICCard.GetSpare
    If CurOneCard < Val(txt缴款.Text) Then
        MsgBox "卡余额不够支付,请检查!" & vbCrLf & vbCrLf & _
        "   卡 余  额" & Format(CurOneCard, "0.00") & vbCrLf & _
        "   本次支付" & Format(Val(txt缴款.Text), "0.00"), vbInformation, gstrSysName
        Exit Function
    End If
    
    stbThis.Panels(4).Text = Format(CurOneCard, "0.00")
    stbThis.Panels(4).ToolTipText = mCurCardPay.str结算方式 & "的帐户余额:" & Format(CurOneCard, "0.00")
    '已经更改了支付金额
    If dblMoney <> Val(txt缴款.Text) Then
        txt缴款.Text = Format(dblMoney, "0.00")
    End If
    CheckOneCard = True
End Function
Private Function CheckPrepayMoneyIsValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查预交数据输入是否合法
    '返回:合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-24 10:36:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    '单独的应缴
    If Not mbln已报价 Then Call LedVoiceSpeak
    
    If BrushcardStrikePrepay = False Then
        If txt冲预交.Enabled And txt冲预交.Visible Then txt冲预交.SetFocus
        zlControl.TxtSelAll txt冲预交
        Exit Function
    End If
    CheckPrepayMoneyIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function isValied(Optional bln连续 As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查收费数据时的有效性,数据有效,返回true,否则返回False
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-13 16:30:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    '单独的应缴
    If Not mbln已报价 Then Call LedVoiceSpeak
    If BrushCardThreeSwapCheck = False Then
        If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
        zlControl.TxtSelAll txt缴款
        Exit Function
    End If
    '一卡通刷卡
    If CheckOneCard = False Then
        If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
        zlControl.TxtSelAll txt缴款
        Exit Function
    End If
    
    If CheckInterfaceNumIsValied = False Then
        If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
        zlControl.TxtSelAll txt缴款
        Exit Function
    End If
    If mCurCardPay.int性质 = 1 Then
        '现金支付,需要大于本次录入金额
        '只有现金才处理,找补
        '门诊收费:
        '刘兴洪:22343,缴款金额控制
        Select Case gTy_Module_Para.byt缴款控制
        Case 1, 3 '1-多病缴款;3单病人缴款累计
            If mblnCur连续 = False Then
                If RoundEx(mCurCarge.dbl当前未付, 2) > 0 And RoundEx(Val(txt缴款.Text), 2) = 0 Then
                   If MsgBox("注意:" & vbCrLf & "    该病人未输入缴款金额,是否继续收费? ", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                       If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
                       zlControl.TxtSelAll txt缴款
                       Exit Function
                   End If
                End If
            End If
        Case 2  '2-收费时必须要输入缴款金额
            If RoundEx(mCurCarge.dbl当前未付, 2) > 0 And RoundEx(Val(txt缴款.Text), 2) = 0 Then
                MsgBox "注意:" & vbCrLf & _
                "    该病人未输入缴款金额,不能进行收费!", vbInformation + vbDefaultButton1, gstrSysName
                If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
                zlControl.TxtSelAll txt缴款
                Exit Function
            End If
        Case Else   ',0-代表不进行缴款输入和累计控制
            '医保结算缴款检查:要缴而未缴时,以缴款作为结束量不处理,因为是强行输入0跳过缴款的
            If mstrYBPati <> "" And Not mInsurePara.门诊连续收费 And RoundEx(mCurCarge.dbl当前未付, 6) > 0 And Val(txt缴款.Text) = 0 Then
                '刘兴洪:27536 20100119
                If mInsurePara.不提醒缴款金额不足 = False Then MsgBox "提醒你:" & vbCrLf & vbTab & "该医保病人的费用未全部结算，请注意收取病人缴款！", vbInformation, gstrSysName
            End If
        End Select
        
        If Val(txt缴款.Text) <> 0 Then
            If CSng(txt找补.Text) < 0 Then
                MsgBox "缴款金额不足,请补足应缴金额！", vbInformation, gstrSysName
                txt缴款.SetFocus: zlControl.TxtSelAll txt缴款
                Exit Function
            End If
        End If
    ElseIf Not mCurCardPay.bln支票 Then
            '问题:42793
            '其他结算方式,输入的金额不能大于未付部分
            If RoundEx(Val(txt缴款.Text), 2) > RoundEx(mCurCarge.dbl当前未付, 2) Then
                MsgBox "注意:" & vbCrLf & "    输入的缴款金额大于了未支付的金额,不能继续!", vbOKOnly + vbInformation, gstrSysName
                txt缴款.SetFocus: zlControl.TxtSelAll txt缴款
                Exit Function
            End If
    End If

    '检查当前单据是否被其他人执行完成,主要是并发原因进行检查
    '防止其他操作员操作:
    '45186
    gstrSQL = "" & _
    "   Select  1  From 病人预交记录 A " & _
    "   Where   A.结算序号=[1] and nvl(A.校对标志,0)<>0 and Rownum =1 and A.记录状态=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng结算ID)
    If rsTemp.EOF Then
        '估计是被他人执行,现在需要检查是否被他人执行
        gstrSQL = "Select 记录状态, 操作员姓名,执行状态 From 门诊费用记录 Where 结帐ID=[1] And rownum=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng结算ID)
        
        If Not rsTemp.EOF Then
            If Val(Nvl(rsTemp!记录状态)) <> 1 Then
                MsgBox "该单据已经被其他操作员作废,不能再进行收费!", vbOKOnly + vbInformation, gstrSysName
                '执行收费
                Unload Me
                Exit Function
            End If
            
            If Val(Nvl(rsTemp!执行状态)) <> 9 Then
                MsgBox "该次收费已经被他人收费,不能再进行收费!", vbOKOnly + vbInformation, gstrSysName
                '执行收费
                Unload Me
                Exit Function
            End If
            
            If Nvl(rsTemp!操作员姓名) <> UserInfo.姓名 Then
                MsgBox "该单据不是本人收费单,不能收取其他操作员的单据!", vbOKOnly + vbInformation, gstrSysName
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
Private Function CheckInterfaceNumIsValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查接口数量是否超过2个以上
    '返回:未超过2个数量,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-02-27 15:23:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCount As Long, varData As Variant
    Dim strNames As String, i As Long
    
    On Error GoTo errHandle
    
    lngCount = IIf(mintInsure <> 0, 1, 0)   '医保算一个数量
    If mCurCardPay.lng医疗卡类别ID = 0 Then CheckInterfaceNumIsValied = True: Exit Function
    With vsBlance
        strNames = IIf(mintInsure <> 0, vbCrLf & "医保结算", "")
        For i = 1 To .Rows - 1
            If Val(.RowData(i)) = 2 Or Val(.RowData(i)) = 3 Then
                '三方接口或一卡通(老版)
                ' 医疗卡类别ID|消费卡(1, 0) |自制卡|是否全退|是否退现|接口名称
                 varData = Split(.Cell(flexcpData, i, .ColIndex("支付方式")) & "|||||", "|")
                 If Val(varData(0)) <> 0 Then
                    If Val(varData(1)) <> 1 Then
                        lngCount = lngCount + 1
                        strNames = strNames & vbCrLf & varData(5)
                    ElseIf Val(varData(2)) = 0 Then
                        '消费卡也是接口的,才算作第三方接口
                        lngCount = lngCount + 1
                        strNames = strNames & vbCrLf & varData(5)
                    End If
                End If
            End If
        Next
    End With
    If lngCount >= 2 Then
        MsgBox "  系统暂只支持两种以内的接口,不能再刷卡消费," & vbCrLf & "  以下为当前已经刷的接口!" & vbCrLf & strNames, vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    CheckInterfaceNumIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckDelValied(ByRef blnExistThreeSwap As Boolean, ByRef bln全退 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查退费
    '出参:blnExistThreeSwap-是否存在三方接口
    '        bln全退-存在三方接口是否必须全退
    '返回:合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-25 16:14:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCardNo As String, i As Long
    Dim strSwapNO As String, strSwapMemo As String, varData As Variant
    Dim lng卡类别ID As Long, bln消费卡 As Boolean, strTemp As String
    Dim st卡类别名称 As String, dblMoney As Double
    bln全退 = False: blnExistThreeSwap = False
    With vsBlance
        For i = 1 To .Rows - 1
            dblMoney = Val(.TextMatrix(i, .ColIndex("支付金额")))
            Select Case Val(.RowData(i))
            Case 2 '三方交易
                ' 医疗卡类别ID|消费卡(1, 0) |自制卡|是否全退|是否退现|接口名称
                strTemp = .Cell(flexcpData, i, .ColIndex("支付方式"))
                If strTemp <> "" Then
                    varData = Split(strTemp & "||||", "|")
                    lng卡类别ID = Val(varData(0))
                    bln消费卡 = Val(varData(1)) = 1
                    st卡类别名称 = varData(5)
                    strSwapNO = Trim(.TextMatrix(i, .ColIndex("交易流水号")))
                    strSwapMemo = Trim(.TextMatrix(i, .ColIndex("交易说明")))
                    strCardNo = .Cell(flexcpData, i, .ColIndex("卡号"))
                    If bln消费卡 And Val(varData(2)) <> 1 Then
                        blnExistThreeSwap = True
                        bln全退 = Val(varData(3)) = 1
                    ElseIf Not bln消费卡 Then
                        blnExistThreeSwap = True
                        bln全退 = Val(varData(3)) = 1
                    End If
                    
                    If zlCheckDelValied(lng卡类别ID, st卡类别名称, bln消费卡, strCardNo, strSwapNO, strSwapMemo, mstr结帐IDs, Val(.TextMatrix(i, .ColIndex("支付金额")))) = False Then
                        Exit Function
                    End If
                End If
             Case 3 '一卡通结算
                strCardNo = .Cell(flexcpData, i, .ColIndex("卡号"))
                 If CheckDelOneCardValied(strCardNo, dblMoney) = False Then Exit Function
                blnExistThreeSwap = True
                bln全退 = True
             Case Else
             End Select
        Next
    End With
    CheckDelValied = True
End Function

Private Function CheckDelOneCardValied(ByVal str原卡号 As String, ByVal dblMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查一卡通退费的有效性
    '返回:合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-25 16:48:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCardNo As String
    On Error GoTo errHandle
    If mobjICCard Is Nothing Then
        On Error Resume Next
        Set mobjICCard = CreateObject("zlICCard.clsICCard")
        On Error GoTo 0
    End If
    If mobjICCard Is Nothing Then
        MsgBox "一卡通接口创建失败,不能进行退费!请检查接口文件.", vbInformation, gstrSysName
        Exit Function
    End If
    strCardNo = mobjICCard.Read_Card(Me)
    If strCardNo = "" Then
        MsgBox "一卡通读卡失败,请将IC卡放在读卡器中", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    If strCardNo <> str原卡号 Then
        MsgBox "当前卡号与扣款卡号不一致,不能进行退费.", vbInformation, gstrSysName
        Exit Function
    End If
    CheckDelOneCardValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function

Private Sub cbo支付方式_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Function Get医保结算ID() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取正常的医保结算(结帐ID)
    '返回:结帐IDs
    '编制:刘兴洪
    '日期:2012-01-05 19:02:52
    '问题:45217
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, str结帐ID As String
    On Error GoTo errHandle
    If mbln异常作废 Then
        gstrSQL = "" & _
        "   Select Distinct A.结帐id " & _
        "   From 门诊费用记录 A, 病人预交记录 B, 门诊费用记录 D, 病人预交记录 C, " & _
        "           (Select 名称 From 结算方式 Where 性质 In (3, 4)) U " & _
        "   Where A.结帐id = B.结帐id And B.结算方式 = U.名称 And Nvl(B.校对标志, 0) = 2 And A.NO = D.NO And " & _
        "         A.记录性质 = D.记录性质 And A.记录状态 In (1, 3) And D.结帐id = C.结帐id And Nvl(C.校对标志, 0) = 1 And " & _
        "         C.结算序号 = [1]"
    Else
        gstrSQL = "" & _
        "   Select  distinct A.结帐ID" & _
        "   From 病人预交记录 A,(Select 名称 From 结算方式 where 性质 in (3,4)) B" & _
        "   Where A.结算序号 = [1]  And A.结算方式=B.名称(+) and nvl(A.校对标志,0)=2"
    End If
   Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng结算ID)
    With rsTemp
        Do While Not .EOF
            str结帐ID = str结帐ID & "," & Val(Nvl(rsTemp!结帐ID))
            .MoveNext
        Loop
    End With
    rsTemp.Close
    Set rsTemp = Nothing
    Get医保结算ID = str结帐ID
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function DelInsureSingle(ByVal blnExistThreeBalance As Boolean, ByRef strSaveCussNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:分单张单据进行医保退费
    '入参:blnExistThreeBalance-是否存在第三方交易
    '出参:strSaveCussNo-销帐成功的单据
    '返回:医保交易成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-02-29 15:51:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, varData As Variant
    Dim cllPro As Collection, i As Long, dbl误差金额 As Double
    Dim DateDel As Date, lng领用ID As Long, strInvoice As String, strNo As String
    Dim blnCommit As Boolean, blnAffaired As Boolean
    Dim strNos As String, strYB退费ID As String
    Dim cllProNO As Collection, lng结算序号 As Long, lng冲销ID As String, str冲销IDs As String
    Dim blnCallInsure As Boolean  '是否要调医保
    Dim blnTrans As Boolean
    Dim varBalance  As Variant, j As Long, intPage As Integer
    Dim intPages As Integer, strAdvance As String, lng结帐ID As Long
    Dim strYB退费IDs As String, strSuccesNo As String
    
    
    DateDel = zlDatabase.Currentdate
    If mintInsure = 0 _
        Or (mintInsure <> 0 And (mInsurePara.多单据调一次交易 Or mInsurePara.多单据一次结算)) _
        Then DelInsureSingle = True: Exit Function
    strYB退费ID = Get医保结算ID
    varBalance = Split(mstr结帐IDs, ",")
    
    '单据作废处理
    varData = Split(Replace(mstrNOs, "'", ""), ",")
    intPages = UBound(varData) + 1
    For i = UBound(varData) To 0 Step -1
        Set cllPro = New Collection
        strNo = varData(i)
        lng冲销ID = zlDatabase.GetNextId("病人结帐记录")
        str冲销IDs = str冲销IDs & "," & lng冲销ID
        If lng结算序号 = 0 Then lng结算序号 = lng冲销ID
        
        'Zl_门诊收费记录_Delete
        strSQL = "zl_门诊收费记录_DELETE("
        '  No_In           门诊费用记录.NO%Type,
        strSQL = strSQL & "'" & varData(i) & "',"
        '  操作员编号_In   门诊费用记录.操作员编号%Type,
        strSQL = strSQL & "'" & UserInfo.编号 & "',"
        '  操作员姓名_In   门诊费用记录.操作员姓名%Type,
        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
        '  医保结算方式_In Varchar2 := Null,
        strSQL = strSQL & "NULL,"
        '  序号_In         Varchar2 := Null,
        strSQL = strSQL & "NULL,"
        '  结算方式_In     病人预交记录.结算方式%Type := Null,
        strSQL = strSQL & "NULL,"
        '  误差_In         门诊费用记录.实收金额%Type := 0,
        strSQL = strSQL & "" & dbl误差金额 & ","
        '  退费时间_In     门诊费用记录.登记时间%Type := Null,
        strSQL = strSQL & "To_Date('" & Format(DateDel, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
        '  回收票据_In     Number := 0,
        strSQL = strSQL & "1,"
        '  退费摘要_In     门诊费用记录.摘要%Type := Null
        strSQL = strSQL & "'结算作废',"
        '     校对标志_In: 0-不需要较对;1-需较对(不处理人员缴款余额,不回收票据)
        strSQL = strSQL & "1,"
        '  结帐id_In       病人预交记录.结帐id%Type := Null,
        strSQL = strSQL & lng冲销ID & ","
        '  结算序号_In     病人预交记录.结算序号%Type := Null
        strSQL = strSQL & lng结算序号 & ","
          '一卡通结算_In   Varchar2 := Null
        strSQL = strSQL & "NULL,"
        '  退款操作_In     Number := 0,
        strSQL = strSQL & "0,"
        '  多单据全退_In   Number := 0,
        strSQL = strSQL & "0)"
        zlAddArray cllPro, strSQL
        If dbl误差金额 <> 0 Then
            strSQL = "zl_门诊收费误差_Insert('" & varData(i) & "'," & dbl误差金额 & ",1,0)"
            zlAddArray cllPro, strSQL
        End If
        '先产生票据，医保接口才能取到
        If mInsurePara.医保接口打印票据 Then
            strSQL = "zl_门诊收费记录_RePrint('" & strNo & "','" & strInvoice & "'," & ZVal(lng领用ID) & ",'" & UserInfo.姓名 & "'," & _
                "To_Date('" & Format(DateDel, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),1,1)"
            zlAddArray cllPro, strSQL
        End If
        '处理其他数据
        'strAdvance = 页数 & "|当前页数"
        For j = 0 To UBound(varData)
            If varData(j) = varData(i) Then intPage = j + 1: Exit For
        Next
        '刘兴洪:医保的strAdvancey计算:本次退费总张数|当前退费第几张:27231
        strAdvance = intPages & "|" & intPage
        lng结帐ID = Val(varBalance(i))
        blnCallInsure = False
        If InStr(1, "," & strYB退费ID & ",", "," & lng结帐ID & ",") > 0 Then
            ' Zl_门诊结算_较对标志_Update
            strSQL = "Zl_门诊结算_较对标志_Update("
            '  结帐id_In     门诊费用记录.结帐id%Type,
            strSQL = strSQL & "" & lng结帐ID & ","
            '  结算序号id_In 病人预交记录.结算序号%Type,
            strSQL = strSQL & "NULL,"
            '  收费结算_In   Varchar2,
            strSQL = strSQL & "'" & mstr医保结算 & "',"
            '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
            strSQL = strSQL & "NULL,"
            '  消费卡_In     Integer := 0,
            strSQL = strSQL & "0,"
            '  卡号_In       病人预交记录.卡号%Type := Null,
            strSQL = strSQL & "NULL,"
            '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
            strSQL = strSQL & "NULL,"
            '  交易说明_In   病人预交记录.交易说明%Type := Null,
            strSQL = strSQL & "NULL,"
            '  校对标志_In   病人预交记录.校对标志%Type := 0
            strSQL = strSQL & "2)"
            zlAddArray cllPro, strSQL
            blnCallInsure = True
         End If
        
        '数据处理
        Err = 0: On Error GoTo Errhand:
        blnCommit = False
        gcnOracle.BeginTrans: blnTrans = True
        zlExecuteProcedureArrAy cllPro, Me.Caption, True, True
        If blnCallInsure Then
            '调用医保接口
            If Not gclsInsure.ClinicDelSwap(lng结帐ID, , mintInsure, strAdvance) Then
                gcnOracle.RollbackTrans
                If strSuccesNo <> "" Then strSuccesNo = Mid(strSuccesNo, 2)
                If blnExistThreeBalance Then
                    '存在第三方接口未退费完成,需要特殊处理
                    Call MsgBox("注意:" & vbCrLf & "    单据为" & varData(i) & "的收费单据作废失败,请注意在异常单据中重新作废!" & vbCrLf & _
                                      IIf(strSuccesNo <> "", vbCrLf & "但有以下单据医保作废成功,但三方交易退费还未进行:" & strSuccesNo, "") & vbCrLf, vbInformation + vbOKCancel + vbDefaultButton1, gstrSysName)
                Else
                    Call MsgBox("注意:" & vbCrLf & "    单据为" & varData(i) & "的收费单据作废失败,请注意在异常单据中重新作废!" & vbCrLf & _
                                      IIf(strSuccesNo <> "", vbCrLf & "但有以下单据作废成功:" & strSuccesNo, "") & vbCrLf, vbInformation + vbOKCancel + vbDefaultButton1, gstrSysName)
                End If
                Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, False, mintInsure)
                Exit Function
            End If
            strSuccesNo = strSuccesNo & "," & strNo
            Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, mintInsure)
        End If
        If Not blnCommit Then gcnOracle.CommitTrans
        blnTrans = False
        If Not blnExistThreeBalance Then
            If OverFeeDel(lng冲销ID, mlng病人ID, True) = False Then
                Call MsgBox("注意:" & vbCrLf & "    单据为" & varData(i) & "的收费单据医保作废成功,但HIS作废失败,请注意在异常单据中重新作废!" & vbCrLf & _
                                  IIf(strSuccesNo <> "", vbCrLf & "但有以下单据作废成功:" & strSuccesNo, "") & vbCrLf, vbInformation + vbOKCancel + vbDefaultButton1, gstrSysName)
                Exit Function
            End If
        End If
    Next
    
    If blnExistThreeBalance Then
        '调用三方回退交易
        blnCommit = True
        If DelSwapThree(str冲销IDs, lng结算序号, blnCommit) = False Then
            If Not blnCommit Then gcnOracle.RollbackTrans
            If MsgBox("注意:" & vbCrLf & "不能正常的进行第三方交易退费,是否暂停交易?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Function
             Exit Function
        End If
        If Left(str冲销IDs, 1) = "," Then str冲销IDs = Mid(str冲销IDs, 2)
        If OverFeeDel(str冲销IDs, mlng病人ID, blnCommit) = False Then
            If Not blnCommit Then gcnOracle.RollbackTrans
            Exit Function
        End If
    End If
    mbytReturnMode = 2
    DelInsureSingle = True
    Exit Function
Errhand:
    
    gcnOracle.RollbackTrans
ErrInterface:
    Call ErrCenter
    Call SaveErrLog
End Function
Public Function zlGet原结帐IDs(ByVal str退费IDs As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结算序号获取冲销ID
    '返回:以逗号分隔的退费的结帐ID,如:123,23,...
    '编制:刘兴洪
    '日期:2012-03-02 10:06:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim str结帐ID As String, i As Long
    Dim strSQL As String, varData As Variant
    
    On Error GoTo errHandle
    strSQL = "" & _
    "   Select /*+ rule */   Distinct A.结帐id,A.NO " & _
    "   From   门诊费用记录 A,门诊费用记录  B,Table( f_Num2list( [1])) C " & _
    "   Where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=3" & _
    "               And B.结帐ID=C.Column_Value" & _
    "   Order by 结帐ID desc "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取原结帐ID数据", str退费IDs)
    varData = Split(Replace(mstrNOs, "'", ""), ",")
    For i = 0 To UBound(varData)
        '保持与mstrNo的位置一致,不然在取对应单据的结帐ID时,将会出错
        rsTemp.Find " NO='" & varData(i) & "'", , , 1
        If Not rsTemp.EOF Then
            str结帐ID = str结帐ID & "," & Val(Nvl(rsTemp!结帐ID))
        Else
            str结帐ID = str结帐ID & "," & "0"
        End If
    Next
    If str结帐ID <> "" Then str结帐ID = Mid(str结帐ID, 2)
    rsTemp.Close
    Set rsTemp = Nothing
    zlGet原结帐IDs = str结帐ID
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdCancelClick()
    Dim strSQL As String, varData As Variant
    Dim cllPro As Collection, i As Long, dbl误差金额 As Double
    Dim DateDel As Date, lng领用ID As Long, strInvoice As String, strNo As String
    Dim blnCommit As Boolean, blnAffaired As Boolean
    Dim strNos As String, strYB退费ID As String
    Dim cllProNO As Collection, lng结算序号 As Long, lng冲销ID As String, str冲销IDs As String
    Dim blnIsExiseThreeSwap As Boolean, bln全退 As Boolean
    
    DateDel = zlDatabase.Currentdate
    '一卡通;第三方交易的检查
    If CheckDelValied(blnIsExiseThreeSwap, bln全退) = False Then
        If MsgBox("注意:" & vbCrLf & "不能正常的进行第三方交易退费,是否暂停交易?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
        Unload Me: Exit Sub
    End If
    If mintInsure <> 0 And mInsurePara.医保接口打印票据 Then
        If zlGetInvoiceGroupUseID(lng领用ID) = False Then
            If MsgBox("注意:" & vbCrLf & "    无有效票据,是否暂停交易?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Unload Me: Exit Sub
        End If
        strInvoice = GetNextBill(lng领用ID)
    End If
    mstrBalances = "": mbln连续输入 = False
    If mintInsure <> 0 And Not mbln异常作废 And Not (mInsurePara.多单据一次结算 Or mInsurePara.多单据调一次交易) And (Not blnIsExiseThreeSwap Or blnIsExiseThreeSwap And bln全退 = False) Then
        If DelInsureSingle(blnIsExiseThreeSwap, "") = False Then Unload Me: Exit Sub
        Unload Me
        Exit Sub
    End If
    
    If mintInsure <> 0 Then strYB退费ID = Get医保结算ID
    '单据作废处理
    Set cllPro = New Collection
    varData = Split(Replace(mstrNOs, "'", ""), ",")
    If Not mbln异常作废 Then
        For i = UBound(varData) To 0 Step -1
            strNo = varData(i)
            lng冲销ID = zlDatabase.GetNextId("病人结帐记录")
            str冲销IDs = str冲销IDs & "," & lng冲销ID
            If lng结算序号 = 0 Then lng结算序号 = lng冲销ID
            'Zl_门诊收费记录_Delete
            strSQL = "zl_门诊收费记录_DELETE("
            '  No_In           门诊费用记录.NO%Type,
            strSQL = strSQL & "'" & varData(i) & "',"
            '  操作员编号_In   门诊费用记录.操作员编号%Type,
            strSQL = strSQL & "'" & UserInfo.编号 & "',"
            '  操作员姓名_In   门诊费用记录.操作员姓名%Type,
            strSQL = strSQL & "'" & UserInfo.姓名 & "',"
            '  医保结算方式_In Varchar2 := Null,
            strSQL = strSQL & "NULL,"
            '  序号_In         Varchar2 := Null,
            strSQL = strSQL & "NULL,"
            '  结算方式_In     病人预交记录.结算方式%Type := Null,
            strSQL = strSQL & "NULL,"
            '  误差_In         门诊费用记录.实收金额%Type := 0,
            strSQL = strSQL & "" & dbl误差金额 & ","
            '  退费时间_In     门诊费用记录.登记时间%Type := Null,
            strSQL = strSQL & "To_Date('" & Format(DateDel, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
            '  回收票据_In     Number := 0,
            strSQL = strSQL & "1,"
            '  退费摘要_In     门诊费用记录.摘要%Type := Null
            strSQL = strSQL & "'结算作废',"
            '     校对标志_In: 0-不需要较对;1-需较对(不处理人员缴款余额,不回收票据)
            strSQL = strSQL & "1,"
            '  结帐id_In       病人预交记录.结帐id%Type := Null,
            strSQL = strSQL & lng冲销ID & ","
            '  结算序号_In     病人预交记录.结算序号%Type := Null
            strSQL = strSQL & lng结算序号 & ","
            '一卡通结算_In   Varchar2 := Null
            strSQL = strSQL & "NULL,"
            '  退款操作_In     Number := 0,
            strSQL = strSQL & "0,"
            '  多单据全退_In   Number := 0,
            strSQL = strSQL & "0)"
            zlAddArray cllPro, strSQL
            If dbl误差金额 <> 0 Then
                strSQL = "zl_门诊收费误差_Insert('" & varData(i) & "'," & dbl误差金额 & ",1,0)"
                zlAddArray cllPro, strSQL
            End If
            '先产生票据，医保接口才能取到
            If mInsurePara.医保接口打印票据 Then
                strSQL = "zl_门诊收费记录_RePrint('" & strNo & "','" & strInvoice & "'," & ZVal(lng领用ID) & ",'" & UserInfo.姓名 & "'," & _
                    "To_Date('" & Format(DateDel, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),1,1)"
                zlAddArray cllPro, strSQL
            End If
           
        Next
        If mstr医保结算 <> "" Then
            ' Zl_门诊结算_较对标志_Update
            strSQL = "Zl_门诊结算_较对标志_Update("
            '  结帐id_In     门诊费用记录.结帐id%Type,
            strSQL = strSQL & "NULL,"
            '  结算序号id_In 病人预交记录.结算序号%Type,
            strSQL = strSQL & "" & mlng结算ID & ","
            '  收费结算_In   Varchar2,
            strSQL = strSQL & "'" & mstr医保结算 & "',"
            '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
            strSQL = strSQL & "NULL,"
            '  消费卡_In     Integer := 0,
            strSQL = strSQL & "0,"
            '  卡号_In       病人预交记录.卡号%Type := Null,
            strSQL = strSQL & "NULL,"
            '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
            strSQL = strSQL & "NULL,"
            '  交易说明_In   病人预交记录.交易说明%Type := Null,
            strSQL = strSQL & "NULL,"
            '  校对标志_In   病人预交记录.校对标志%Type := 0
            strSQL = strSQL & "2)"
            zlAddArray cllPro, strSQL
        End If
        '全退
        Err = 0: On Error GoTo Errhand:
        zlExecuteProcedureArrAy cllPro, Me.Caption, True
    Else
        str冲销IDs = mstr冲销IDs: lng结算序号 = mlng结算ID
        gcnOracle.BeginTrans
    End If
    
    On Error GoTo ErrInterface:
    blnCommit = False
    If mintInsure <> 0 And mstr医保结算 <> "" Then
        If mInsurePara.多单据一次结算 Then
            If DelInsureMulitOneSwap(varData, DateDel, blnCommit) = False Then
                If blnCommit = False Then
                    gcnOracle.RollbackTrans: Exit Sub
                End If
            End If
        ElseIf mInsurePara.多单据调一次交易 Then
              If DelInsureMulitCallOneInterfrace(varData, blnCommit) = False Then
                    If blnCommit = False Then gcnOracle.RollbackTrans
                    Exit Sub
              End If
        Else
            '循环调用接口
            If InsureCallInterface(varData, strYB退费ID, blnCommit) = False Then
                If blnCommit = False Then gcnOracle.RollbackTrans
                Unload Me
                Exit Sub
            End If
        End If
    End If
    'blnAffaired = mstr医保结算 <> ""    '已经进行了事务处理
    '调用三方回退交易
    If DelSwapThree(str冲销IDs, lng结算序号, blnCommit) = False Then
        If Not blnCommit Then gcnOracle.RollbackTrans
        If MsgBox("注意:" & vbCrLf & "不能正常的进行第三方交易退费,是否暂停交易?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
        Unload Me: Exit Sub
    End If
    
    If OverFeeDel(str冲销IDs, mlng病人ID, blnCommit) = False Then
        If Not blnCommit Then gcnOracle.RollbackTrans
        Exit Sub
    End If
    mbytReturnMode = 2: mblnOK = True
    Unload Me
    Exit Sub
Errhand:
    gcnOracle.RollbackTrans
ErrInterface:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function OverFeeDel(ByVal str冲销IDs As String, ByVal lng病人ID As Long, ByRef blnCommited As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:完成退费收费
    '入参:strNos-完成收费的单据(可以为多张,但目前只有一张单据)
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-29 14:50:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errHandle
    If Left(str冲销IDs, 1) = "," Then str冲销IDs = Mid(str冲销IDs, 2)

    ' Zl_门诊收费结算_完成退费
    strSQL = "Zl_门诊收费结算_完成退费("
    '  病人id_In       门诊费用记录.病人id%Type,
    strSQL = strSQL & "" & lng病人ID & ","
    '  退费结算序号_In 病人预交记录.结算序号%Type,
    strSQL = strSQL & "NULL,"
    '  冲销ids_In      Varchar2,
    strSQL = strSQL & "'" & str冲销IDs & "',"
    '  操作员姓名_In   病人预交记录.操作员姓名%Type := Null
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '操作标志_In     Integer := 0:
    '0-更新缴款余额和预交余额;1-不更新缴款余额和预交余额,2-不更新人员缴款余额,只更新预交余额
    strSQL = strSQL & "2,"
    '异常作废_In     Number := 0
    strSQL = strSQL & "1)"
    '异常单据,冲销也应该为异常单据
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    If Not blnCommited Then
        gcnOracle.CommitTrans: blnCommited = True
    End If
    OverFeeDel = True
    Exit Function
errHandle:
    If Not blnCommited Then gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
    blnCommited = True
End Function
Private Function DelSwapThree(ByVal str冲销IDs As String, ByVal lng退费结算序号 As String, blnCommit As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:退费交易(一卡通或三方结算交易)
    '入参:blnCommit -已经进行了事务处理的
    '返回:交易成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-25 17:29:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim strCardNo As String, i As Long, strSQL As String, strErrMsg As String
    Dim strSwapNO As String, strSwapMemo As String, varData As Variant
    Dim lng卡类别ID As Long, bln消费卡 As Boolean, strTemp As String
    Dim st卡类别名称 As String, blnTrans As Boolean, dblMoney As Double
    Dim cllUpdate As Collection, cllThreeSwap As Collection
    Dim str医院编码 As String, rsTemp As ADODB.Recordset
    
    gstrSQL = "" & _
    "   Select A.结算方式,A.摘要, " & _
    "             nvl(A.卡类别ID,nvl(A.结算卡序号,0)) as 卡类别ID,Decode(nvl(A.结算卡序号,0),0,0,1) as 消费卡," & _
    "             A.结算号码,A.卡号,A.交易流水号, " & _
    "             nvl(C.是否自制,M.自制卡) as 自制卡, " & _
    "             nvl(C.名称,M.名称) as 名称,A.交易说明,A.结算序号," & _
    "             Sum(A.冲预交) as 冲预交" & _
    "   From 病人预交记录 A ,医疗卡类别 C,消费卡类别目录 M" & _
    "   Where A.结算序号=[1] And nvl(A.校对标志,0)=1  " & _
    "                And A.卡类别ID=C.ID(+) and A.结算卡序号=M.编号(+)   " & _
    "                And nvl(A.卡类别ID,nvl(A.结算卡序号,0))<>0 " & _
    "   Group by A.结算方式,A.摘要,nvl(A.卡类别ID,nvl(A.结算卡序号,0)),Decode(nvl(A.结算卡序号,0),0,0,1) ," & _
    "             A.结算号码,A.卡号,A.交易流水号, " & _
    "             nvl(C.是否自制,M.自制卡) , " & _
    "             nvl(C.名称,M.名称),A.交易说明,A.结算序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng退费结算序号)
    Set cllUpdate = New Collection: Set cllThreeSwap = New Collection
    With rsTemp
        Do While Not .EOF
                lng卡类别ID = Val(Nvl(!卡类别ID))
                bln消费卡 = Val(Nvl(!消费卡)) = 1
                st卡类别名称 = Nvl(!名称)
                strSwapNO = Nvl(!交易流水号)
                strSwapMemo = Nvl(!交易说明)
                strCardNo = Nvl(!卡号)
                dblMoney = Nvl(!冲预交)
                
               ' Zl_门诊结算_较对标志_Update
                strSQL = "Zl_门诊结算_较对标志_Update("
                '  结帐id_In     门诊费用记录.结帐id%Type,
                strSQL = strSQL & "" & "NULL" & ","
                '  结算序号id_In 病人预交记录.结算序号%Type,
                strSQL = strSQL & "" & lng退费结算序号 & ","
                '  收费结算_In   Varchar2,
                strSQL = strSQL & "'" & Nvl(!结算方式) & "',"
                '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
                strSQL = strSQL & "" & lng卡类别ID & ","
                '  消费卡_In     Integer := 0,
                strSQL = strSQL & "" & IIf(bln消费卡, 1, 0) & ","
                '  卡号_In       病人预交记录.卡号%Type := Null,
                strSQL = strSQL & "'" & strCardNo & "',"
                '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
                strSQL = strSQL & "'" & strSwapNO & "',"
                '  交易说明_In   病人预交记录.交易说明%Type := Null,
                strSQL = strSQL & "'" & strSwapMemo & "',"
                '  校对标志_In   病人预交记录.校对标志%Type := 0
                strSQL = strSQL & "2)"
                
                '61688
                If blnCommit Then
                    gcnOracle.BeginTrans
                End If
                 blnTrans = True
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
                If CallBackBalanceInterface(str冲销IDs, lng卡类别ID, bln消费卡, dblMoney, strCardNo, strSwapNO, strSwapMemo, cllUpdate, cllThreeSwap, strErrMsg) = False Then
                    blnCommit = True
                    gcnOracle.RollbackTrans: Exit Function
                End If
                gcnOracle.CommitTrans: blnTrans = False: blnCommit = True
                zlExecuteProcedureArrAy cllUpdate, Me.Caption
                zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
            .MoveNext
        Loop
    End With
    DelSwapThree = True
    Exit Function
errHandle:
    If blnCommit = False Then gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog:     blnCommit = True
End Function
Private Function zlDelOneCard(ByVal strCardNo As String, ByVal str医院编码 As String, _
    ByVal str交易流水号 As String, ByVal dblMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:退一卡通数据
    '出参:strErrMsg-返回的错误信息
    '返回:
    '编制:刘兴洪
    '日期:2011-08-25 17:38:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If Not mobjICCard.ReturnSwap(strCardNo, str医院编码, str交易流水号, dblMoney) Then
            MsgBox "一卡通退费交易调用失败,退费操作失败！", vbExclamation, gstrSysName
            Exit Function
    End If
    zlDelOneCard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function InsureCallInterface(ByVal varNos As Variant, ByVal strYB退费IDs As String, Optional blnCommited As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:按单据号调用接口
    '参数:strYB退费IDs-需退医保的费用ID,用逗号分离
    '返回:调用成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-25 12:21:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHaveInterface As Boolean, strAdvance As String
    Dim intPages As Integer, j As Integer
    Dim i As Long, intPage As Integer, varBalance As Variant
    Dim lng结帐ID As Long, blnTrans As Boolean, strSQL As String
    Dim strsuccesNOs As String
    
    '医保要求从最后一张开始退
    varBalance = Split(mstr结帐IDs, ",")
    intPages = UBound(varNos) + 1
    blnTrans = False: strsuccesNOs = ""
    For i = UBound(varNos) To 0 Step -1
        blnTrans = False
        'strAdvance = 页数 & "|当前页数"
        For j = 0 To UBound(varNos)
            If varNos(j) = varNos(i) Then intPage = j + 1: Exit For
        Next
        
        '刘兴洪:医保的strAdvancey计算:本次退费总张数|当前退费第几张:27231
        strAdvance = intPages & "|" & intPage
TORe:
        lng结帐ID = Val(varBalance(i))
        ' Zl_门诊结算_较对标志_Update
        strSQL = "Zl_门诊结算_较对标志_Update("
        '  结帐id_In     门诊费用记录.结帐id%Type,
        strSQL = strSQL & "" & lng结帐ID & ","
        '  结算序号id_In 病人预交记录.结算序号%Type,
        strSQL = strSQL & "NULL,"
        '  收费结算_In   Varchar2,
        strSQL = strSQL & "'" & mstr医保结算 & "',"
        '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
        strSQL = strSQL & "NULL,"
        '  消费卡_In     Integer := 0,
        strSQL = strSQL & "0,"
        '  卡号_In       病人预交记录.卡号%Type := Null,
        strSQL = strSQL & "NULL,"
        '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
        strSQL = strSQL & "NULL,"
        '  交易说明_In   病人预交记录.交易说明%Type := Null,
        strSQL = strSQL & "NULL,"
        '  校对标志_In   病人预交记录.校对标志%Type := 0
        strSQL = strSQL & "2)"
        If InStr(1, "," & strYB退费IDs & ",", "," & lng结帐ID & ",") > 0 Then
            If blnCommited Then gcnOracle.BeginTrans
            blnTrans = True
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
            If Not gclsInsure.ClinicDelSwap(lng结帐ID, , mintInsure, strAdvance) Then
                gcnOracle.RollbackTrans: blnCommited = True
                Call MsgBox("注意:" & vbCrLf & "    单据为" & varNos(i) & " 的收费单据进行医保退费时失败,你必须在『异常单据』中重新进行作废处理或与系统管理员联系!" & vbCrLf & _
                                   IIf(strsuccesNOs <> "", "以下为医保已经退费成功的单据:" & vbCrLf & strsuccesNOs, "") & vbCrLf & _
                                   "" & vbCrLf, vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName)
                      '  GoTo TORe:
                 Exit Function
            End If
            gcnOracle.CommitTrans: blnCommited = True
            strsuccesNOs = strsuccesNOs & "," & varNos(i)
            Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, mintInsure)
            blnHaveInterface = True
        End If
    Next
    InsureCallInterface = True
End Function

Private Function DelInsureMulitCallOneInterfrace(ByVal varNos As Variant, ByRef blnCommit As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:多单据调用一次接口
    '返回:多单据调用一次接口成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-25 12:17:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strAdvance As String, varBalance As Variant, lng结帐ID As Long
    Dim strSQL As String, blnTransMedicare As Boolean
    Dim dbl结算金额 As Double, dbl可分配额 As Double, dbl退款合计 As Double, dbl余额 As Double
    Dim dbl误差金额 As Double
    Dim str结算方式 As String, strBalance As String
    Dim arrData As Variant, blnTrans As Boolean
    Dim cllPro As Collection, rsTmp As ADODB.Recordset
    Dim k As Long, j As Long, i As Long
    
    blnCommit = False
    If mInsurePara.多单据调一次交易 = False Then DelInsureMulitCallOneInterfrace = True: Exit Function
    On Error GoTo errHandle
    varBalance = Split(mstr结帐IDs, ",")
    strAdvance = mstr结帐IDs
    lng结帐ID = Val(varBalance(UBound(varBalance)))
    blnTransMedicare = False
    If Not gclsInsure.ClinicDelSwap(lng结帐ID, , mintInsure, strAdvance) Then
         blnCommit = True
         gcnOracle.RollbackTrans
         Exit Function
    End If
    blnTransMedicare = True
    If strAdvance = mstr结帐IDs Or strAdvance = "" Then
        Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, mintInsure)
        If Not blnCommit Then gcnOracle.CommitTrans: blnCommit = True
        DelInsureMulitCallOneInterfrace = True
        Exit Function
    End If
    '根据返回的结算信息，修正预交记录，strAdvance返回格式:结算方式1|金额||结算方式2:金额...
    '先分摊到每张单据上
    '1.分摊的医保
    Set mrsBlance = Nothing
    Set rsTmp = GetBalanceSet
    varBalance = Split(strAdvance, "||")
    For i = 0 To UBound(varBalance)
        str结算方式 = Split(varBalance(i), "|")(0)
        dbl结算金额 = -1 * Val(Split(varBalance(i), "|")(1))
        For k = 0 To UBound(varNos)
            dbl可分配额 = Get实收金额(varNos(k))
            rsTmp.Filter = "单据序号=" & k
            For j = 1 To rsTmp.RecordCount
                dbl可分配额 = dbl可分配额 - rsTmp!结算金额
                rsTmp.MoveNext
            Next
            If dbl可分配额 > 0 Then
                If dbl可分配额 <= dbl结算金额 Then
                    dbl结算金额 = dbl结算金额 - dbl可分配额
                Else
                    dbl可分配额 = dbl结算金额
                    dbl结算金额 = 0
                End If
                rsTmp.AddNew
                rsTmp!单据序号 = k
                rsTmp!结算方式 = str结算方式
                If k = UBound(varNos) Then  '未完摊完的,放在最后一张单据上
                    rsTmp!结算金额 = dbl可分配额 + dbl结算金额
                End If
                rsTmp.Update
                If dbl结算金额 = 0 Then Exit For
            ElseIf k = UBound(varNos) Then  '未完摊完的,放在最后一张单据上
                rsTmp.AddNew
                rsTmp!单据序号 = k
                rsTmp!结算方式 = str结算方式
                rsTmp!结算金额 = dbl结算金额
                rsTmp.Update
            End If
        Next
    Next
    For k = 0 To UBound(varNos)
        strBalance = ""
        dbl误差金额 = 0
        dbl余额 = Get实收金额(varNos(k))
        rsTmp.Filter = "单据序号=" & k
        For i = 1 To rsTmp.RecordCount
            strBalance = IIf(strBalance = "", "", strBalance & "||") & rsTmp!结算方式 & "|" & -1 * rsTmp!结算金额
            dbl余额 = dbl余额 - rsTmp!结算金额
            rsTmp.MoveNext
        Next
        dbl退款合计 = dbl退款合计 + dbl余额
        lng结帐ID = GetDelBalanceID(varNos(k))
        'Zl_医保结算校对_Update
        strSQL = "Zl_医保结算校对_Update("
        '  结帐id_In   门诊费用记录.结帐id%Type,
        strSQL = strSQL & lng结帐ID & ","
        '  保险结算_In Varchar2
        strSQL = strSQL & strBalance & ")"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    Next
    gcnOracle.CommitTrans: blnTrans = False
    blnCommit = True
    Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, mintInsure)

    MsgBox "应退金额" & vbCrLf & zlstr.NeedName(cbo支付方式.Text) & "：" & Format(dbl退款合计, "0.00") & "元", vbInformation + vbOKOnly, gstrSysName
    
    DelInsureMulitCallOneInterfrace = True
    Exit Function
errHandle:
    '问题:50134
    If blnTrans Then gcnOracle.RollbackTrans
    If blnTransMedicare Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, False, mintInsure)
    Call ErrCenter
End Function

Private Function DelInsureMulitOneSwap(ByVal varNos As Variant, _
    ByVal dtDate As Date, Optional blnCommit As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:多单据一次结算
    '出参:blnCommit-是否已经提交
    '返回:多单据一次结算或非多单据一次结算,成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-25 10:45:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strAdvance As String, varBalance As Variant, lng结帐ID As Long
    Dim strSQL As String, blnTransMedicare As Boolean
    Dim dbl结算金额 As Double, dbl可分配额 As Double, dbl退款合计 As Double, dbl余额 As Double
    Dim dbl误差金额 As Double
    Dim str结算方式 As String, strBalance As String
    Dim arrData As Variant, blnTrans As Boolean
    Dim cllPro As Collection, rsTmp As ADODB.Recordset
    Dim k As Long, j As Long, i As Long
    
    blnTrans = True
    blnCommit = False
    On Error GoTo errHandle
    If mintInsure = 0 Then DelInsureMulitOneSwap = True: Exit Function
    If Not mInsurePara.多单据一次结算 Then DelInsureMulitOneSwap = True: Exit Function
    
    varBalance = Split(mstr结帐IDs, ",")
    strAdvance = mstr结帐IDs
    lng结帐ID = Val(varBalance(UBound(varBalance)))
    blnTransMedicare = False
    If Not gclsInsure.ClinicDelSwap(lng结帐ID, , mintInsure, strAdvance) Then
         blnCommit = True
         gcnOracle.RollbackTrans
         Exit Function
    End If
    blnTransMedicare = True
    If strAdvance = mstr结帐IDs Or strAdvance = "" Then
        Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, mintInsure)
        gcnOracle.CommitTrans: blnCommit = True
        DelInsureMulitOneSwap = True
        Exit Function
    End If
    '根据返回的结算信息，修正预交记录，strAdvance返回格式:结算方式1|金额||结算方式2:金额...
    '先分摊到每张单据上
    '1.分摊的医保
    Set mrsBlance = Nothing
    Set rsTmp = GetBalanceSet
    varBalance = Split(strAdvance, "||")
    For i = 0 To UBound(varBalance)
        str结算方式 = Split(varBalance(i), "|")(0)
        dbl结算金额 = -1 * Val(Split(varBalance(i), "|")(1))
        For k = 0 To UBound(varNos)
            dbl可分配额 = Get实收金额(varNos(k))
            rsTmp.Filter = "单据序号=" & k
            For j = 1 To rsTmp.RecordCount
                dbl可分配额 = dbl可分配额 - rsTmp!结算金额
                rsTmp.MoveNext
            Next
            If dbl可分配额 > 0 Then
                If dbl可分配额 <= dbl结算金额 Then
                    dbl结算金额 = dbl结算金额 - dbl可分配额
                Else
                    dbl可分配额 = dbl结算金额
                    dbl结算金额 = 0
                End If
                rsTmp.AddNew
                rsTmp!单据序号 = k
                rsTmp!结算方式 = str结算方式
                If k = UBound(varNos) Then  '未完摊完的,放在最后一张单据上
                    rsTmp!结算金额 = dbl可分配额 + dbl结算金额
                End If
                rsTmp.Update
                If dbl结算金额 = 0 Then Exit For
            ElseIf k = UBound(varNos) Then  '未完摊完的,放在最后一张单据上
                rsTmp.AddNew
                rsTmp!单据序号 = k
                rsTmp!结算方式 = str结算方式
                rsTmp!结算金额 = dbl结算金额
                rsTmp.Update
            End If
        Next
    Next
    For k = 0 To UBound(varNos)
        strBalance = ""
        dbl误差金额 = 0
        dbl余额 = Get实收金额(varNos(k))
        rsTmp.Filter = "单据序号=" & k
        For i = 1 To rsTmp.RecordCount
            strBalance = IIf(strBalance = "", "", strBalance & "||") & rsTmp!结算方式 & "|" & -1 * rsTmp!结算金额
            dbl余额 = dbl余额 - rsTmp!结算金额
            rsTmp.MoveNext
        Next
        dbl退款合计 = dbl退款合计 + dbl余额
        lng结帐ID = GetDelBalanceID(varNos(k))
        'Zl_医保结算校对_Update
        strSQL = "Zl_医保结算校对_Update("
        '  结帐id_In   门诊费用记录.结帐id%Type,
        strSQL = strSQL & lng结帐ID & ","
        '  保险结算_In Varchar2
        strSQL = strSQL & strBalance & ")"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    Next
    gcnOracle.CommitTrans: blnTrans = False
    blnCommit = True
    Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, mintInsure)

    MsgBox "应退金额" & vbCrLf & zlstr.NeedName(cbo支付方式.Text) & "：" & Format(dbl退款合计, "0.00") & "元", vbInformation + vbOKOnly, gstrSysName
    
    DelInsureMulitOneSwap = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    If blnTrans Then gcnOracle.RollbackTrans
    If blnTransMedicare Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, False, mintInsure)

End Function
 

Private Sub cmdDel_Click()
    Dim dblMoney As Double, strSQL As String
    Dim byt操作类型 As Byte
    Dim str结算方式 As String
    If mbytFunc = EM_异常作废 Then Exit Sub
    '删除相关的费用
    With vsBlance
        If .Row < 0 Then Exit Sub
        '.rowdata:0-普通的结算方式-1-医保结算;2-三方接口交易;3-一卡通结算;4-预存款
        Select Case Val(.RowData(.Row))
        Case 1, 2, 3    '1-医保结算;2-三方接口交易;3-一卡通结算
            '不能直接删除
            Exit Sub
        Case 4  '预存款
            byt操作类型 = 1
            str结算方式 = ""
        Case 0  '普通的结算方式
            byt操作类型 = 0
            str结算方式 = .TextMatrix(.Row, .ColIndex("支付方式"))
        Case Else
            Exit Sub
        End Select
        dblMoney = Val(.TextMatrix(.Row, .ColIndex("支付金额")))
        If Not (byt操作类型 = 0 Or byt操作类型 = 1) Then
            '不进行处理
            'Zl_病人收费结算_预交_Del
            strSQL = " Zl_病人收费结算_预交_Del("
            '  操作类型_In   Number,0-正常收费;1-冲预交
            strSQL = strSQL & "" & byt操作类型 & ","
            '  结算序号id_In 病人预交记录.结算序号%Type,
            strSQL = strSQL & "" & mlng结算ID & ","
            '  结算方式_In   Varchar2,
            strSQL = strSQL & "" & IIf(str结算方式 = "", "NULL", "'" & str结算方式 & "'") & ","
            '  结算金额_In   病人预交记录.冲预交%Type
            strSQL = strSQL & dblMoney & ")"
            Err = 0: On Error GoTo Errhand:
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
        End If
        mCurCarge.dbl当前未付 = RoundEx(mCurCarge.dbl当前未付 + dblMoney, 6)
        mCurCarge.dbl本次已付合计 = RoundEx(mCurCarge.dbl本次已付合计 - dblMoney, 6)
        Call SetControlProperty
        If Val(.RowData(.Row)) = 4 Then
            txt冲预交.Enabled = True: txt冲预交.BackColor = txt缴款.BackColor
            If txt冲预交.Enabled And txt冲预交.Visible Then txt冲预交.SetFocus
            zlControl.TxtSelAll txt冲预交
            txt冲预交.Tag = "": lbl冲预交.Tag = ""
        End If
        If .Rows <= 2 Then
            .Clear 1
            .RowData(1) = ""
            .Cell(flexcpData, 1, 0, 1, .COLS - 1) = ""
        Else
            vsBlance.RemoveItem .Row
        End If
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub cmdExit_Click()
    mblnOK = False: mbytReturnMode = EM_退出收费
    Unload Me
End Sub

Private Sub cmdNext_Click()
    '继续下一张单据的录入
    '保存上次支付方式
    If mCurCarge.dbl本次冲预交 <> 0 Then
        MsgBox "使用了预交款后,不能连续收费!", vbOKOnly + vbInformation, gstrSysName
        Exit Sub
    End If
    gtyPrePatiPay = mCurCardPay: mblnCur连续 = True
    If Not Check缴款(2) Then GoTo GoOver
    '先处理预交
    If BrushcardStrikePrepay = False Then GoTo GoOver
    '再处理其他
    If isValied(True) = False Then GoTo GoOver
    If SaveCharge = False Then GoTo GoOver
    mbln连续输入 = True
    mbytReturnMode = 3
GoOver:
    mstrBalances = ""
    mblnCur连续 = False
End Sub

Private Sub cmdOK_Click()
    '并发检查
    If mbytFunc = EM_异常作废 Or mbytFunc = EM_重新收费 Then
        If zlIsCheckExistErrBill(mlng结算ID) = False Then
            MsgBox "当前异常单据已被处理，你不能继续！", vbInformation, gstrSysName
            Exit Sub
        End If
        If zlCheckOtherSessionDoing(mlng结算ID) Then
            MsgBox "当前单据正在其它收费窗口中进行处理，你不能继续！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
   If mbytFunc = EM_异常作废 Then
     Call cmdCancelClick
     Exit Sub
   End If
   '单据界面按了回车符
   If mblnCacheKeyReturn Then mblnCacheKeyReturn = False: Exit Sub
    '先处理预交
    mbln连续输入 = False
    If BrushcardStrikePrepay = False Then Exit Sub
    '再处理其他
    If isValied = False Then Exit Sub
    If txt缴款.Text <> "0.00" Then
        'LED显示
        Call ShowLedInfor
    End If
    If SaveCharge = False Then Exit Sub
End Sub
Private Sub SetCtrlVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件的显示状态
    '编制:刘兴洪
    '日期:2012-02-03 13:58:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnTemp As Boolean
    If mbytFunc = EM_正常收费 Then
        '医保且医保未进行结算时,才显示
        cmdYBBalance.Visible = mintInsure <> 0 And mblnYbBalanced = False
        '医保进行结算了的,或非医保的,显示完成收费
        cmdOK.Visible = (mintInsure = 0 Or mintInsure <> 0 And mblnYbBalanced)
        '医保进行了结算后,不能退出
        cmdExit.Visible = mintInsure = 0 And Not mblnThreeInterface Or mintInsure <> 0 And mblnYbBalanced = False
        '连续收费
        blnTemp = gTy_Module_Para.byt缴款控制 = 1 Or gTy_Module_Para.byt缴款控制 = 3 '是否具体连续收费
        '普通收费或医保已经结算
        blnTemp = blnTemp And (mintInsure = 0 Or mintInsure <> 0 And mblnYbBalanced)
        blnTemp = blnTemp And Val(txt冲预交.Text) = 0 '未用预交款的
        cmdNext.Visible = blnTemp
        If (gTy_Module_Para.byt缴款控制 = 1 Or gTy_Module_Para.byt缴款控制 = 3) And mbln连续输入 Then
            cbo支付方式.Locked = True
        End If
        Exit Sub
     End If
     
     If mbytFunc = EM_重新收费 Then
        cmdExit.Caption = "退出(&E)"
        cmdOK.Visible = True: cmdYBBalance.Visible = False
        mblnYbBalanced = mintInsure <> 0    '医保结算时,异常单据一般都是结算了的.
        cmdExit.Visible = True: cmdNext.Visible = False
     End If
     If mbytFunc = EM_异常作废 Then
        cmdOK.Caption = "作废结算(&O)"
        cmdExit.Caption = "退出(&E)"
        cmdOK.Visible = True: cmdYBBalance.Visible = False
        cmdExit.Visible = True: cmdNext.Visible = False
     End If
End Sub
Private Function SaveBill(Optional blnNotCommit As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存单据处理
    '入参:不进行事务提交(主要是处理普通病人收费，让其放在一个事务中进行处理，减少异常单据的出现)
    '返回:保存成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-02-05 16:50:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
   Dim blnCancel As Boolean, strNos As String
    '数据保存
    RaiseEvent zlSaveData(mlng结算ID, mstr结帐IDs, strNos, blnNotCommit, blnCancel)
    mstrNOs = strNos
    If blnCancel Then Exit Function
    SaveBill = True
End Function
Private Function 医保结算较对() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:是否存在医保校对(主要是医保交易调用成功后,出现医保数据的较对)
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-02-07 16:21:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim str医保结算 As String, i As Long, strShowMsg As String
    Dim strTemp As String, dblMoney As Double
    Dim lng结帐ID As Long
    On Error GoTo errHandle
    If mbytFunc = EM_正常收费 Then 医保结算较对 = True: Exit Function
    If mintInsure = 0 Then 医保结算较对 = True: Exit Function
    If mstr结帐IDs = "" Then Exit Function
    
    '0-正常;1-待校对;2-完成校对;3-附加，指返回的其它非医保支付的各种结算方式
    gstrSQL = "" & _
    "   Select /*+ rule */ A.记录ID,A.校正  " & _
    "   From 保险结算记录 A,Table( f_Num2list([1]))  B " & _
    "   Where A.记录ID=B.Column_Value And nvl(A.校正,0)=1 "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstr结帐IDs)
    If rsTemp.EOF Then 医保结算较对 = True: Exit Function
    '检查医保核对表，无记录则退出
    'Select 结帐ID,结算方式,金额 From 保险结算明细 Where 标志=1
    gstrSQL = "" & _
    "   Select /*+ rule */  A.结帐ID,a.结算方式,a.金额" & _
    "    From 保险结算明细 A,Table( f_Num2list([1])) B ,结算方式 C" & _
    "   Where A.结帐id =B.Column_Value and A.标志=1 and A.结算方式=C.名称 And C.性质 in (3,4) " & _
    "   Order by A.结算方式"
    '医保管控的过程固定写入了一条"现金",所以排开非医保类的结算方式
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "保险结算管理", mstr结帐IDs)
        '未有核对数据,直接返回
    If rsTemp.RecordCount = 0 Then 医保结算较对 = True: Exit Function
    
    str医保结算 = ""   '结算方式|结算金额||
    strShowMsg = ""
    strTemp = "": dblMoney = 0
    For i = 1 To rsTemp.RecordCount
        If strTemp <> Nvl(rsTemp!结算方式, " ") Then
            If strTemp <> "" And dblMoney <> 0 Then
                 str医保结算 = str医保结算 & "||" & strTemp & "|" & dblMoney
                 strShowMsg = strShowMsg & vbCrLf & strTemp & ":" & dblMoney
            End If
            strTemp = Nvl(rsTemp!结算方式, " ")
            dblMoney = 0
        End If
        dblMoney = dblMoney + Val(Nvl(rsTemp!金额))
        rsTemp.MoveNext
    Next
    If strTemp <> "" And dblMoney <> 0 Then
         str医保结算 = str医保结算 & "||" & strTemp & "|" & dblMoney
        strShowMsg = strShowMsg & vbCrLf & strTemp & ":" & dblMoney
    End If
    If str医保结算 <> "" Then str医保结算 = Mid(str医保结算, 3)
    MsgBox "注意:" & vbCrLf & "  在进行医保结算时,医保预结结果与正式结算不一致,请校对保险结算数据,以下为正确的结算数据:" & vbCrLf & strShowMsg, vbInformation + vbOKOnly, gstrSysName
    If mInsurePara.多单据调一次交易 Or mInsurePara.多单据一次结算 Then
        'Zl_病人门诊收费_医保更新
        gstrSQL = "Zl_病人门诊收费_医保更新("
        '结帐id_In   门诊费用记录.结帐id%Type,
        gstrSQL = gstrSQL & "NULL,"
        '结算序号_In 病人预交记录.结算序号%Type,
        gstrSQL = gstrSQL & "" & mlng结算ID & ","
        '保险结算_In Varchar2
        gstrSQL = gstrSQL & "'" & str医保结算 & "')"
        Err = 0: On Error GoTo ErrCommit:
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
        '重新加载数据
        Call LoadData
        Call LoadPatiInfor
        Call SetControlProperty
        医保结算较对 = True
        Exit Function
    End If
    lng结帐ID = 0: strTemp = ""
    '进行数据效对
    rsTemp.Sort = "结帐ID"
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    With rsTemp
        str医保结算 = ""
        Do While Not .EOF
            If lng结帐ID <> Val(Nvl(!结帐ID)) Then
                If lng结帐ID <> 0 Then
                    str医保结算 = Mid(str医保结算, 3)
                    '较对数据
                    'Zl_病人门诊收费_医保更新
                    gstrSQL = "Zl_病人门诊收费_医保更新("
                    '结帐id_In   门诊费用记录.结帐id%Type,
                    gstrSQL = gstrSQL & "" & lng结帐ID & ","
                    '结算序号_In 病人预交记录.结算序号%Type,
                    gstrSQL = gstrSQL & "NULL,"
                    '保险结算_In Varchar2
                    gstrSQL = gstrSQL & "'" & str医保结算 & "')"
                    Err = 0: On Error GoTo ErrCommit:
                    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
                End If
                lng结帐ID = Val(Nvl(!结帐ID))
                str医保结算 = ""
            End If
            strTemp = Trim(Nvl(rsTemp!结算方式, " "))
            If strTemp <> "" Then
                str医保结算 = str医保结算 & "||" & strTemp & "|" & Val(Nvl(rsTemp!金额))
            End If
            .MoveNext
        Loop
        If str医保结算 <> "" And lng结帐ID <> 0 Then
            str医保结算 = Mid(str医保结算, 3)
            '较对数据
            'Zl_病人门诊收费_医保更新
            gstrSQL = "Zl_病人门诊收费_医保更新("
            '结帐id_In   门诊费用记录.结帐id%Type,
            gstrSQL = gstrSQL & "" & lng结帐ID & ","
            '结算序号_In 病人预交记录.结算序号%Type,
            gstrSQL = gstrSQL & "NULL,"
            '保险结算_In Varchar2
            gstrSQL = gstrSQL & "'" & str医保结算 & "')"
            Err = 0: On Error GoTo ErrCommit:
            zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
        End If
    End With
    '重新加载数据
    Call LoadData
    Call LoadPatiInfor
    Call SetControlProperty
    医保结算较对 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
ErrCommit:
    Call ErrCenter
    Resume  '必须执行完才能正常执行
End Function


Private Sub cmdYBBalance_Click()
    Dim blnCancel As Boolean, strNos As String
    
    '并发检查
    If mbytFunc = EM_异常作废 Or mbytFunc = EM_重新收费 Then
        If zlIsCheckExistErrBill(mlng结算ID) = False Then
            MsgBox "当前异常单据已被处理，你不能继续！", vbInformation, gstrSysName
            Exit Sub
        End If
        If zlCheckOtherSessionDoing(mlng结算ID) Then
            MsgBox "当前单据正在其它收费窗口中进行处理，你不能继续！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    '数据保存
    If SaveBill = False Then Exit Sub
    mblnYbBalanced = True   '医保已经结算
    Call LoadData
    '医保:58344
    mblnYB退款 = mCurCarge.dbl当前未付 < 0
    
    Call LoadPatiInfor
    Call SetControlProperty
    '完成医保结算,需要重新设置按钮
    Call SetCtrlVisible
    Call SetControlEnabled
    '光标定位
    '优先使用预交
    If txt冲预交.Visible And txt冲预交.Enabled And gblnPrePayPriority Then
        txt冲预交.SetFocus
        Call SetControlProperty(True): mbln已报价 = True
        Call Show误差金额(True)
    Else
        mblnNotChange = True
        txt冲预交.Text = ""
        mblnNotChange = False
        '70430,冉俊明,2014-4-24,在进行预结算时提示缴款金额，进行医保结算时再次提示相同缴款金额，造成重复提示。
        If txt缴款.Enabled And txt缴款.Visible Then
            mbln已报价 = True '先设置已报价为true,屏蔽txt缴款获得焦点而报价
            txt缴款.SetFocus
        End If
        Call Show误差金额(False)
    End If
    Call LedDisplayBank
    
    If mCurCarge.dbl当前未付 = 0 And cmdOK.Visible And cmdOK.Enabled Then
        '医保全部结算,直接确定完成:63773
        Call cmdOK_Click
    End If
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    Call cbo支付方式_Click
    Call SetControlProperty
    Call SetCtrlVisible
    Call SetControlEnabled
    If 医保结算较对 = False Then Unload Me: Exit Sub
    If txt冲预交.Visible Then txt冲预交.Enabled = True
    '光标定位
    If Val(txt冲预交.Text) <> 0 And txt冲预交.Enabled Then
        If txt冲预交.Enabled And txt冲预交.Visible Then txt冲预交.SetFocus
        Call Show误差金额(True)
    Else
        If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
        Call Show误差金额(False)
    End If
    mblnLoad = False
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Select Case KeyCode
    Case vbKeyAdd, vbKeyF4
        If (gTy_Module_Para.byt缴款控制 = 1 Or gTy_Module_Para.byt缴款控制 = 3) And mbln连续输入 Then Exit Sub
        '47457
        If gTy_Module_Para.bln使用加减切换 = False And KeyCode = vbKeyAdd Then Exit Sub
        If Me.ActiveControl Is txt缴款 Then
            i = cbo支付方式.ListIndex
            If i >= cbo支付方式.ListCount - 1 Then
                i = 0
            Else
                i = i + 1
            End If
            cbo支付方式.ListIndex = i
        End If
    Case vbKeySubtract
        If (gTy_Module_Para.byt缴款控制 = 1 Or gTy_Module_Para.byt缴款控制 = 3) And mbln连续输入 Then Exit Sub
        '47457
        If gTy_Module_Para.bln使用加减切换 = False And KeyCode = vbKeySubtract Then Exit Sub
        If Me.ActiveControl Is txt缴款 Then
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
        '强制完成
        If mintInsure <> 0 And mblnYbBalanced = False Then
            Call cmdYBBalance_Click
        Else
            cmdOK_Click '43169
        End If
    Case vbKeyReturn
      '      zlCommFun.PressKey vbKeyTab
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr(":'|", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Form_Load()
    '选检查主界面中是否发送了回车键的
    mblnCacheKeyReturn = False
    mblnCacheKeyReturn = (GetAsyncKeyState(VK_RETURN) And &H1) <> 0
    mstrTitle = "病人收费结算"
    RestoreWinState Me, App.ProductName, mstrTitle
    Call zlInitTotalStru
    Call SetWindowsSize
    Set mrsOneCard = GetOneCard
    zlControl.CboSetWidth cbo支付方式.hWnd, cbo支付方式.Width * 2
    txt冲预交.Enabled = False
    mblnFirst = True: mblnLoad = True
    mblnUnLoad = False
    zlControl.PicShowFlat picTotal, -1, , taCenterAlign
    zlControl.PicShowFlat Picture1, -1, , taCenterAlign
    zlControl.PicShowFlat picPay, -1, , taCenterAlign
    Call InitFace
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    'If Me.Width < 10530 Then Me.Width = 10530
    'If Me.Height < 7035 Then Me.Height = 7035
    With picBlance
        .Width = ScaleWidth - .Left * 2
        .Height = ScaleHeight - stbThis.Height - .Top
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    mlng结算ID = 0
    With mCurCarge
           .dbl本次冲预交 = 0
           .dbl本次实收 = 0
           .dbl本次医保支付 = 0
           .dbl本次已付合计 = 0
           .dbl本次应收 = 0
           .dbl当前未付 = 0
           .dbl费用余额 = 0
           .dbl可用预交 = 0
           .dbl预交余额 = 0
    End With
    mblnYB退款 = False
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
        Set mobjICCard = Nothing
    End If
    Set mrsClassMoney = Nothing
    With mCurCardPay
        .lng消费卡ID = 0
        .str限制类别 = ""
        .dbl已刷金额 = 0
    End With
    SaveWinState Me, App.ProductName, mstrTitle
End Sub

 

Private Sub picBlance_Resize()
    Err = 0: On Error Resume Next
    With vsBlance
'        fraSplitBottom.Left = 0
'        fraSplitBottom.Width = picBlance.ScaleWidth + 50
        .Left = picBlance.ScaleLeft
        .Width = picBlance.ScaleWidth
        .Height = picBlance.ScaleHeight - .Top
    End With
End Sub
Private Sub setDefaultPrepayMoney()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置缺省预交金额
    '编制:刘兴洪
    '日期:2011-08-13 17:21:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With mCurCarge
         txt冲预交.Text = "0.00"
         If Not mblnLoad Or (mblnLoad And gblnPrePayPriority) Then
            If .dbl可用预交 <> 0 Then
                txt冲预交.Text = Format(IIf(.dbl可用预交 > .dbl当前未付, .dbl当前未付, .dbl可用预交), "###0.00;###0.00;0.00;0.00")
            End If
        End If
    End With
End Sub
Private Sub LoadPatiInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载病人信息
    '编制:刘兴洪
    '日期:2011-08-13 10:52:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    stbThis.Panels(2).Text = mstrPatiInfo
    Set rsTemp = GetMoneyInfo(mlng病人ID, 0, False, 1, False)
    With mCurCarge
        .dbl预交余额 = 0
        .dbl费用余额 = 0
        If Not rsTemp.EOF Then
            .dbl预交余额 = RoundEx(Val(Nvl(rsTemp!预交余额)), 6)
            .dbl费用余额 = RoundEx(Val(Nvl(rsTemp!费用余额)), 6)
        End If
        .dbl可用预交 = RoundEx(.dbl预交余额 - .dbl费用余额, 6)
        If .dbl可用预交 < 0 Then .dbl可用预交 = 0
    End With
    txt医保.Text = Format(mCurCarge.dbl本次医保支付, "###0.00;-###0.00;0.00;0.00;")
    txt合计.Text = Format(mCurCarge.dbl本次实收, "###0.00;-###0.00;0.00;0.00;")
    stbThis.Panels(3).Text = Format(mCurCarge.dbl可用预交, "####0.00;-####0.00;0.00;0.00")
    
    lbl自付合计.Caption = Format(mCurCarge.dbl本次实收 - mCurCarge.dbl本次医保支付, "###0.00;-###0.00;0.00;0.00")
    Call setDefaultPrepayMoney
    If mCurCarge.dbl本次冲预交 <> 0 Then
        txt冲预交.Text = Format(mCurCarge.dbl本次冲预交, "0.00")
        txt冲预交.Tag = "1"
        txt冲预交.BackColor = Me.BackColor
        lbl冲预交.Tag = "1"
        txt冲预交.Enabled = False
    End If
End Sub
Private Sub LedVoiceSpeak()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:语音报价
    '编制:刘兴洪
    '日期:2011-08-13 16:38:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    'If mCurCardPay.int性质 <> 1 Then Exit Sub
    If gblnLED = False Then Exit Sub
    If mintInsure <> 0 And mblnYbBalanced = False Then Exit Sub
    
'    If mCurCarge.dbl本次实收 = 0 Then Exit Sub
'    If mCurCarge.dbl当前未付 = 0 Then Exit Sub
    zl9LedVoice.Speak "#21 " & Format(lbl剩余自付.Caption, "0.00")
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

Private Sub txt冲预交_Change()
    lbl冲预交.Tag = "": txt冲预交.Tag = ""
    txt冲预交.BackColor = txt缴款.BackColor
    If mblnNotChange Then Exit Sub
    Call SetControlProperty(True)
    Call Show误差金额(True)
End Sub
Private Sub txt冲预交_GotFocus()
    If Val(txt冲预交.Text) = 0 And mblnLoad = False Then
           Call setDefaultPrepayMoney
    End If
    zlControl.TxtSelAll txt冲预交
    Call SetControlProperty(True)
     
    'If Val(txt冲预交.Tag) = Val(txt冲预交.Text) Then Exit Sub
    
    '自动报价或手工报价时由热键激活
    'Call LedVoiceSpeak
   
End Sub

Private Sub txt冲预交_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
     zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt冲预交_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt冲预交, KeyAscii, m金额式
End Sub
Private Sub txt冲预交_LostFocus()
      If mblnLoad Then Exit Sub
      
      If Val(txt冲预交.Tag) = Val(txt冲预交.Text) Then Exit Sub
      If Val(txt冲预交.Text) = 0 Then Exit Sub
      If CheckPrepayMoneyIsValied = False Then Exit Sub
      
End Sub

Private Sub txt冲预交_Validate(Cancel As Boolean)
    If lbl冲预交.Tag = "1" Then Exit Sub
    If mlng病人ID = 0 Then Exit Sub
    
    If Val(txt冲预交.Tag) = Val(txt冲预交.Text) Then Exit Sub
    
    If txt冲预交.Text = "" Then
        txt冲预交.Text = "0.00"
    ElseIf Not IsNumeric(txt冲预交.Text) And txt冲预交.Text <> "" Then
        MsgBox "无效数值！", vbInformation, gstrSysName
        If txt冲预交.Enabled And txt冲预交.Visible Then txt冲预交.SetFocus
        zlControl.TxtSelAll txt冲预交: Cancel = True: Exit Sub
    ElseIf Val(txt冲预交.Text) < 0 Then
        MsgBox "预存款冲款金额不能为负！", vbInformation, gstrSysName
        Call setDefaultPrepayMoney
        If txt冲预交.Enabled And txt冲预交.Visible Then txt冲预交.SetFocus
        zlControl.TxtSelAll txt冲预交: Cancel = True: Exit Sub
    ElseIf Val(txt冲预交.Text) > 0 And mCurCarge.dbl本次实收 < 0 Then
        MsgBox "单据应付金额为负时不能使用预存款！", vbInformation, gstrSysName
        txt冲预交.Text = "0.00"
        If txt冲预交.Enabled And txt冲预交.Visible Then txt冲预交.SetFocus
        zlControl.TxtSelAll txt冲预交:   Exit Sub
    ElseIf Val(txt冲预交.Text) > mCurCarge.dbl可用预交 Then
        MsgBox "预存款冲款金额不能超过病人的预存余额:" & Format(mCurCarge.dbl可用预交, "0.00") & " ！", vbInformation, gstrSysName
        Call setDefaultPrepayMoney
        If txt冲预交.Enabled And txt冲预交.Visible Then txt冲预交.SetFocus
        zlControl.TxtSelAll txt冲预交: Cancel = True: Exit Sub
    ElseIf Val(txt冲预交.Text) > Format(mCurCarge.dbl当前未付, "0.00") And Val(txt冲预交.Text) <> 0 Then
        MsgBox "预交款冲款金额不能大于应付金额:" & Format(mCurCarge.dbl当前未付, "0.00") & " ！", vbInformation, gstrSysName
        Call setDefaultPrepayMoney
        If txt冲预交.Enabled And txt冲预交.Visible Then txt冲预交.SetFocus
        zlControl.TxtSelAll txt冲预交: Cancel = True: Exit Sub
    Else
        txt冲预交.Text = Format(Val(txt冲预交.Text), "0.00")
    End If
   ' If CheckPrepayMoneyIsValied = False Then Cancel = True: Exit Sub
End Sub

Private Function BrushcardStrikePrepay() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:验证刷卡冲预交
    '返回:冲销成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-14 14:35:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If Val(lbl冲预交.Tag) = 1 Then BrushcardStrikePrepay = True: Exit Function
    If Val(txt冲预交) = 0 Then BrushcardStrikePrepay = True: Exit Function
    If Not IsNumeric(txt冲预交.Text) And txt冲预交.Text <> "" Then
        MsgBox "无效数值！", vbInformation, gstrSysName
        If txt冲预交.Enabled And txt冲预交.Visible Then txt冲预交.SetFocus
        zlControl.TxtSelAll txt冲预交: Exit Function
    ElseIf Val(txt冲预交.Text) < 0 Then
        MsgBox "预存款冲款金额不能为负！", vbInformation, gstrSysName
        Call setDefaultPrepayMoney
        If txt冲预交.Enabled And txt冲预交.Visible Then txt冲预交.SetFocus
        zlControl.TxtSelAll txt冲预交: Exit Function
    ElseIf Val(txt冲预交.Text) > 0 And mCurCarge.dbl本次实收 < 0 Then
        MsgBox "单据应付金额为负时不能使用预存款！", vbInformation, gstrSysName
        txt冲预交.Text = "0.00"
        If txt冲预交.Enabled And txt冲预交.Visible Then txt冲预交.SetFocus
        zlControl.TxtSelAll txt冲预交:   Exit Function
    ElseIf Val(txt冲预交.Text) > mCurCarge.dbl可用预交 Then
        MsgBox "预存款冲款金额不能超过病人的预存余额:" & Format(mCurCarge.dbl可用预交, "0.00") & " ！", vbInformation, gstrSysName
        Call setDefaultPrepayMoney
        If txt冲预交.Enabled And txt冲预交.Visible Then txt冲预交.SetFocus
        zlControl.TxtSelAll txt冲预交: Exit Function
    ElseIf Val(txt冲预交.Text) > Format(mCurCarge.dbl当前未付, "0.00") And Val(txt冲预交.Text) <> 0 Then
        MsgBox "预交款冲款金额不能大于应付金额:" & Format(mCurCarge.dbl当前未付, "0.00") & " ！", vbInformation, gstrSysName
        Call setDefaultPrepayMoney
        If txt冲预交.Enabled And txt冲预交.Visible Then txt冲预交.SetFocus
        zlControl.TxtSelAll txt冲预交: Exit Function
    End If
    '刷卡确认
    'frmParent As Object, ByVal lngSys As Long, _
    ByVal lng病人ID As Long, ByVal cur金额 As Currency, _
    Optional lngModul As Long = 0, _
    Optional bytOperationType As Byte = 0
    If zlDatabase.PatiIdentify(Me, glngSys, mlng病人ID, Val(txt冲预交), mlngModule, 1, mlngBrushCardTypeID, _
            IIf(-1 * gdbl预存款消费验卡 >= Val(txt冲预交), False, True), , , (gdbl预存款消费验卡 <> 0), (gdbl预存款消费验卡 = 2)) Then
        lbl冲预交.Tag = "1"
       ' txt冲预交.ForeColor = d
       txt冲预交.BackColor = Me.BackColor
       txt冲预交.Tag = Val(txt冲预交)
       txt冲预交.Enabled = False
        If SaveCharge(True) = False Then
            txt冲预交.Enabled = True
            txt冲预交.BackColor = txt缴款.BackColor
             lbl冲预交.Tag = ""
            Exit Function
        End If
         BrushcardStrikePrepay = True
        If mblnUnloaded Then Exit Function
    Else
        lbl冲预交.Tag = ""
        txt冲预交.Enabled = True
        Call SetControlProperty
       Exit Function
    End If
    Call SetControlProperty
    BrushcardStrikePrepay = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function BrushCardThreeSwapCheck() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:刷卡验证
    '返回:返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-14 14:35:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsMoney As ADODB.Recordset, strXMLExpend As String
    Dim dblMoney  As Double, dblBrushCardMoneyed As Double '已刷消费卡金额
    Dim cllSquareBalance As Collection
    On Error GoTo errHandle
    If mCurCardPay.lng医疗卡类别ID = 0 Then BrushCardThreeSwapCheck = True: Exit Function
    If Val(txt缴款) = 0 Then
        MsgBox "未输入交易金额,请检查!", vbInformation + vbOKOnly
         Exit Function
    End If
    If Not IsNumeric(txt缴款.Text) And txt缴款.Text <> "" Then
        MsgBox "无效数值！", vbInformation, gstrSysName
        If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
        zlControl.TxtSelAll txt缴款: Exit Function
    ElseIf Val(txt缴款.Text) < 0 Then
        MsgBox "交易金额不能为负！", vbInformation, gstrSysName
        If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
        zlControl.TxtSelAll txt缴款: Exit Function
    ElseIf Abs(Val(txt缴款.Text)) > Format(Abs(mCurCarge.dbl当前未付), "0.00") And Val(txt缴款.Text) <> 0 Then
        MsgBox "交易金额不能大于本次未付金额:" & Format(mCurCarge.dbl当前未付, "0.00") & " ！", vbInformation, gstrSysName
        If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
        zlControl.TxtSelAll txt缴款: Exit Function
    End If
    If mCurCardPay.bln消费卡 And mblnYB退款 Then
        MsgBox "当前为退款模式,目前系统暂不支持将退款额退给" & mCurCardPay.str结算方式, vbInformation + vbOKOnly, gstrSysName
        If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
        zlControl.TxtSelAll txt缴款: Exit Function
        Exit Function
    End If
    If zlGetClassMoney(mlng结算ID, rsMoney) = False Then Exit Function
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
    Optional ByRef varSquareBalance As Variant) As Boolean
    '       varSquareBalance- Collection类型,返回当前刷卡数据(array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文))
    Set cllSquareBalance = Nothing
    Set mcllCurSquareBalance = Nothing
    If mCurCardPay.bln消费卡 Then
        '构建消费卡的刷卡信息
       Set cllSquareBalance = mcllSquareBalance
     End If
     
    dblMoney = Val(txt缴款.Text)
    If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModule, rsMoney, _
        mCurCardPay.lng医疗卡类别ID, mCurCardPay.bln消费卡, _
    mstr姓名, mstr性别, mstr年龄, dblMoney, mCurCardPay.str刷卡卡号, mCurCardPay.str刷卡密码, _
    False, True, False, False, cllSquareBalance) = False Then Exit Function
    '消费卡附值
    If mCurCardPay.bln消费卡 Then
        Set mcllCurSquareBalance = cllSquareBalance
    End If
    
    '保存前,一些数据检查
    'zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
    ByVal strCardTypeID As Long, ByVal strCardNo As String, _
    ByVal dblMoney As Double, ByVal strNOs As String, _
    Optional ByVal strXMLExpend As String
    'mstrNOs:单独保存时,没有相关时,可能为空.
    If gobjSquare.objSquareCard.zlPaymentCheck(Me, mlngModule, mCurCardPay.lng医疗卡类别ID, _
        mCurCardPay.bln消费卡, mCurCardPay.str刷卡卡号, dblMoney, mstrNOs, strXMLExpend) = False Then Exit Function
'    zlGetAccountMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
'    ByVal strCardTypeID As Long, _
'    ByVal strCardNo As String, strExpand As String, dblMoney As Double
    '入参:frmMain-调用的主窗体
    '        lngModule-模块号
    '        strCardNo-卡号
    '        strExpand-预留，为空,以后扩展
    '出参:dblMoney-返回帐户余额
    Dim strExpand As String, dbl帐户余额 As Double
    If gobjSquare.objSquareCard.zlGetAccountMoney(Me, mlngModule, mCurCardPay.lng医疗卡类别ID, _
          mCurCardPay.str刷卡卡号, strExpand, dbl帐户余额, mCurCardPay.bln消费卡) = False Then Exit Function
    stbThis.Panels(4).Text = Format(dbl帐户余额, "0.00")
    stbThis.Panels(4).ToolTipText = mCurCardPay.str结算方式 & "的帐户余额:" & Format(dbl帐户余额, "0.00")
    mCurCardPay.dbl帐户余额 = RoundEx(dbl帐户余额, 2)
    '已经更改了支付金额
    If dblMoney <> Val(txt缴款.Text) Then
        txt缴款.Text = Format(dblMoney, "0.00")
    End If
    BrushCardThreeSwapCheck = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function zlGetClassMoney(ByRef lng结帐序号 As Long, ByRef rsMoney As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存时,初始化支付类别(收费类别,实收金额)
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-10 17:52:18
    '问题:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    If Not mrsClassMoney Is Nothing Then
        Set rsMoney = mrsClassMoney: zlGetClassMoney = True: Exit Function
    End If
    
    '初始化数据结构
    Set mrsClassMoney = New ADODB.Recordset
    mrsClassMoney.Fields.Append "收费类别", adVarChar, 10, adFldIsNullable
    mrsClassMoney.Fields.Append "金额", adDouble, , adFldIsNullable
    mrsClassMoney.CursorLocation = adUseClient
    mrsClassMoney.LockType = adLockOptimistic
    mrsClassMoney.CursorType = adOpenStatic
    mrsClassMoney.Open
    If lng结帐序号 = 0 And mbytFunc = EM_正常收费 Then
        Call mfrmMain.zlGetClassMoney(rsTemp)
    Else
        strSQL = "" & _
        "   Select  A.收费类别,nvl(sum(实收金额) ,0) as 金额   " & _
        "   From 门诊费用记录 A,(Select 结帐ID From 病人预交记录 where 结算序号=[1] ) B " & _
        "   Where A.结帐ID=B.结帐ID " & _
        "   Group by 收费类别"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng结帐序号)
    End If
    With rsTemp
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            mrsClassMoney.Find "收费类别='" & Nvl(!收费类别, "无") & "'", , adSearchForward, 1
            If mrsClassMoney.EOF Then mrsClassMoney.AddNew
            mrsClassMoney!收费类别 = Nvl(!收费类别, "无")
            mrsClassMoney!金额 = Val(Nvl(mrsClassMoney!金额)) + Val(Nvl(!金额))
            mrsClassMoney.Update
            .MoveNext
        Loop
    End With
    Set rsMoney = mrsClassMoney
    zlGetClassMoney = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub txt缴款_Change()
    Call SetControlProperty
    Call Show误差金额(False)
End Sub
Private Sub txt缴款_GotFocus()
    '只以缴款作为收费结束条件时,必须输入缴款或0
    '刘兴洪:22343
    If gTy_Module_Para.byt缴款控制 = 1 _
        Or gTy_Module_Para.byt缴款控制 = 3 _
        Or gTy_Module_Para.byt缴款控制 = 2 Then
        If Val(txt缴款.Text) = 0 And Me.ActiveControl Is txt缴款 Then
            txt缴款.Text = ""
        End If
    End If
    Call SetControlProperty
  '  Call zlControl.TxtSelAll(txt缴款)
    '自动报价或手工报价时由热键激活
    If Not mbln已报价 Then Call LedVoiceSpeak
    zlControl.TxtSelAll txt缴款
End Sub
Private Sub ShowLedInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示Led信息
    '编制:刘兴洪
    '日期:2011-08-13 15:25:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gblnLED = False Then Exit Sub
'    If mCurCarge.dbl本次实收 = 0 Then Exit Sub
    
    '只有缴现才显示
    If Val(txt冲预交.Text) = 0 And mCurCardPay.int性质 = 1 Then
        zl9LedVoice.DispCharge mCurCarge.dbl当前未付 + mCurCarge.dbl应缴累计, Val(txt缴款.Text), Val(txt找补.Text)
    Else '部分支付现金时的处理
        Call zl9LedVoice.DisplayBank( _
            "合计:" & txt合计.Text & "元,应付:" & lbl剩余自付.Caption & "元", _
            "收您:" & txt缴款.Text & "元" & IIf(Val(txt找补.Text) = 0, "", ",找您:" & Val(txt找补.Text) & "元"))
    End If
    zl9LedVoice.Speak "#22 " & Val(txt缴款.Text)
    zl9LedVoice.Speak "#23 " & Val(txt找补.Text)
    zl9LedVoice.Speak "#3"
End Sub

Private Sub LedDisplayBank()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示保价信息
    '编制:刘兴洪
    '日期:2011-12-15 13:40:46
    '问题:52117
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl个帐合计 As Double, i As Long
    Dim str医保 As String, str三方交易 As String, str老一卡通 As String, str普通结算 As String
    Dim varPara  As Variant, str结算方式 As String
    If Not gblnLED Then Exit Sub
    
    With vsBlance
        For i = 1 To .Rows - 1
            '医保交易
            If .TextMatrix(i, .ColIndex("支付方式")) <> "" Then
                Select Case .RowData(i)
                Case 1 '医保
                    str医保 = str医保 & "||" & .TextMatrix(i, .ColIndex("支付方式")) & ":" & Format(Val(.TextMatrix(i, .ColIndex("支付金额"))), "0.00")
                Case 2 '三方接口交易
                    str三方交易 = str三方交易 & "||" & .TextMatrix(i, .ColIndex("支付方式")) & ":" & Format(Val(.TextMatrix(i, .ColIndex("支付金额"))), "0.00")
                Case 3   ' 一卡通交易
                    str老一卡通 = str老一卡通 & "||" & .TextMatrix(i, .ColIndex("支付方式")) & ":" & Format(Val(.TextMatrix(i, .ColIndex("支付金额"))), "0.00")
                Case Else
                    str普通结算 = str普通结算 & "||" & .TextMatrix(i, .ColIndex("支付方式")) & ":" & Format(Val(.TextMatrix(i, .ColIndex("支付金额"))), "0.00")
                End Select
            End If
        Next
    End With
     
    str结算方式 = ""
    If str医保 <> "" Then str结算方式 = str结算方式 & "||医保结算:||帐户余额:" & Format(mcur个帐余额, "0.00") & str医保
    If str三方交易 <> "" Then str结算方式 = str结算方式 & "||一卡通结算:" & str三方交易
    If str老一卡通 <> "" Then str结算方式 = str结算方式 & "||一卡通结算(老):" & str老一卡通
    If str普通结算 <> "" Then str结算方式 = str结算方式 & "" & str普通结算
    If str结算方式 = "" Then Exit Sub
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

    '70430,冉俊明,2014-4-24,在进行预结算时提示缴款金额，进行医保结算时再次提示相同缴款金额，造成重复提示。
    If Format(mdbl原未付, gstrDec) <> Format(Val(lbl剩余自付.Caption), gstrDec) Then
        zl9LedVoice.Speak "#21 " & Format(Val(lbl剩余自付.Caption), "0.00")
    End If
End Sub
Private Function Check缴款(ByVal bytMode As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查缴款金额
    '入参:bytMode-0-在缴款处回车输入;1-按的是完成;2-按的是继续下张输入
    '出参:
    '返回:输入会法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-09-06 10:47:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If cbo支付方式.ListIndex < 0 Then Exit Function
    If txt缴款.Text <> "" Then
        If Abs(Val(txt缴款.Text)) > 99999999 Then
            MsgBox "输入的缴款金额过大,最大不能超过99999999!", vbOKOnly, gstrSysName
            Exit Function
        End If
        If Val(txt缴款.Text) = 0 Then
            If cbo支付方式.ItemData(cbo支付方式.ListIndex) = -1 Then
                '需要排除三方接口交易
                MsgBox "未输入缴款金额,不能用" & cbo支付方式.Text & "支付,请检查!", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
            If mCurCardPay.blnOneCard Then
                '需要排除三方接口交易
                MsgBox "未输入缴款金额,不能用一卡通进行支付,请检查!", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
        End If
        Check缴款 = True
        Exit Function
    End If
    
    '未输入缴款金额检查
    '缴款控制:0-代表不进行缴款输入和累计控制,1-代表输入缴款后才结束病人累计
    '       2-收费时必须要输入缴款金额
    Select Case gTy_Module_Para.byt缴款控制
    Case 1, 3 '1-多病人累计; 3-单病人累计划
        If cbo支付方式.ItemData(cbo支付方式.ListIndex) = -1 Then
            '需要排除三方接口交易
            MsgBox "未输入缴款金额,不能用" & cbo支付方式.Text & "支付,请检查!", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
        If mCurCardPay.blnOneCard Then
            '需要排除三方接口交易
            MsgBox "未输入缴款金额,不能用一卡通进行支付,请检查!", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
    Case 2  '收费时必须要输入缴款金额
            MsgBox "未输入缴款金额,请检查", vbOKOnly + vbInformation, gstrSysName
            txt缴款.SetFocus: Exit Function
    End Select
    
    Check缴款 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub txt缴款_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt缴款, KeyAscii, m金额式
    If KeyAscii <> 13 Then Exit Sub
    If mblnCacheKeyReturn = True Then mblnCacheKeyReturn = False
    KeyAscii = 0
    If Check缴款(0) = False Then Exit Sub
    
    If (gTy_Module_Para.byt缴款控制 = 1 Or gTy_Module_Para.byt缴款控制 = 3) And txt缴款.Text = "" Then
        If cmdNext.Visible And cmdNext.Enabled Then cmdNext.SetFocus
          Exit Sub
    End If
    
    '只以缴款作为收费结束条件时,必须输入缴款或0
    If gTy_Module_Para.byt缴款控制 = 1 _
        Or gTy_Module_Para.byt缴款控制 = 3 _
        Or gTy_Module_Para.byt缴款控制 = 2 Then
        If txt缴款.Text = "" Then Exit Sub
    End If
    
    If mCurCardPay.int性质 <> 1 Then
        If mCurCardPay.bln支票 Or (cbo支付方式.Text Like "*卡*" And mCurCardPay.lng医疗卡类别ID = 0) Then
            zlCommFun.PressKey vbKeyTab
        Else
            Call cmdOK_Click
            Call txt缴款_GotFocus   '47147
        End If
        Exit Sub
    End If
    
    If Val(txt缴款.Text) = 0 Then txt缴款.Text = "0.00"
    If txt缴款.Text <> "0.00" Then
        If CSng(txt找补.Text) >= 0 Then
            'LED显示
            'Call ShowLedInfor
            '确定
             Call cmdOK_Click
        Else
            MsgBox "缴款金额不足,请补足应缴金额！", vbInformation, gstrSysName
            txt缴款.SetFocus: zlControl.TxtSelAll txt缴款
        End If
        Exit Sub
    End If
    Call cmdOK_Click
   ' Call zlCommFun.PressKey(vbKeyTab) '病人累加缴款
End Sub
Private Sub Load支付方式()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载有效的支付方式
    '编制:刘兴洪
    '日期:2011-07-08 11:41:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim j As Long, strPayType As String, varData As Variant, varTemp As Variant, i As Long
    Dim rsTemp As ADODB.Recordset, blnFind As Boolean
    Dim strSQL As String
    
    Set rsTemp = Get结算方式("收费")
    Set mcolCardPayMode = New Collection
    '短|全名|刷卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式|是否密文|是否自制卡;…
    If Not gobjSquare Is Nothing Then
        strPayType = gobjSquare.objSquareCard.zlGetAvailabilityCardType
    End If
    varData = Split(strPayType, ";")
    With cbo支付方式
        .Clear: j = 0
        Do While Not rsTemp.EOF
            blnFind = False
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "|||||", "|")
                If varTemp(6) = Nvl(rsTemp!名称) Then
                    blnFind = True
                    Exit For
                End If
            Next
            If Not blnFind Then
                If Not (Val(Nvl(rsTemp!性质)) = 3 Or Val(Nvl(rsTemp!性质)) = 4 Or Val(Nvl(rsTemp!应付款)) = 1) Then
                    '不加入医保的结算方式
                    .AddItem Nvl(rsTemp!名称)
                    mcolCardPayMode.Add Array("", Nvl(rsTemp!名称), 0, 0, 0, 0, Nvl(rsTemp!名称), 0, 0), "K" & j
                    If rsTemp!缺省 = 1 Then .ListIndex = .NewIndex
                    If Val(Nvl(rsTemp!性质)) = 1 And .ListIndex < 0 Then .ListIndex = .NewIndex
                    .ItemData(.NewIndex) = Val(Nvl(rsTemp!性质))
                    If mbln连续输入 Then
                        If gtyPrePatiPay.str结算方式 = Nvl(rsTemp!名称) Then
                             .ListIndex = .NewIndex
                        End If
                    End If
                    j = j + 1
              End If
            End If
            rsTemp.MoveNext
        Loop
        
        For i = 0 To UBound(varData)
            If InStr(1, varData(i), "|") <> 0 Then
                varTemp = Split(varData(i), "|")
                mcolCardPayMode.Add varTemp, "K" & j
                .AddItem varTemp(1): .ItemData(.NewIndex) = -1
                If mbln连续输入 Then
                    '   '短|全名|刷卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式|是否密文|是否自制卡;…
                    If gtyPrePatiPay.lng医疗卡类别ID = Val(varTemp(3)) _
                        And gtyPrePatiPay.bln消费卡 And Val(varTemp(5)) = 1 Then
                         .ListIndex = .NewIndex
                    ElseIf gtyPrePatiPay.lng医疗卡类别ID = Val(varTemp(3)) _
                        And gtyPrePatiPay.bln消费卡 = False And Val(varTemp(5)) = 0 Then
                         .ListIndex = .NewIndex
                    End If
                Else
                    '缺省为主界面中的刷卡类别
                    If mlngBrushCardTypeID = Val(varTemp(3)) And Val(varTemp(5)) <> 1 Then .ListIndex = .NewIndex
                End If
                j = j + 1
            End If
        Next
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0
        If Not mbln连续输入 And gstr结算方式 <> "" Then
            '60574
            '根据参数设置缺省的支付类别
            For j = 0 To .ListCount - 1
                If .List(j) = gstr结算方式 Then
                    .ListIndex = j: Exit For
                End If
            Next
        End If
        If .ListCount = 0 Then
            MsgBox "预交场合没有可用的结算方式,请先到结算方式管理中设置。", vbExclamation, gstrSysName
            mblnUnLoad = True: Exit Sub
        End If
    End With
End Sub
Private Sub txt缴款_LostFocus()
    txt缴款.Text = Format(Val(txt缴款.Text), "0.00")
End Sub

Private Sub txt缴款_Validate(Cancel As Boolean)
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
        If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus
    End If
End Sub
Private Sub txt找补_GotFocus()
    zlControl.TxtSelAll txt找补
End Sub

Private Function zlOneCardPrayMoney(ByVal dblMoney As Double, ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:一卡通支付
    '返回:一卡通支付成功,返回true,否则返回false
    '编制:刘兴洪
    '日期:2011-08-23 17:57:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl余额 As Double, str医院编码 As String
    If mCurCardPay.blnOneCard = False Then zlOneCardPrayMoney = True: Exit Function
    mrsOneCard.Filter = "结算方式='" & mCurCardPay.str结算方式 & "'"
    If mrsOneCard.EOF Then
         strErrMsg = "未找到结算方式为" & mCurCardPay.str结算方式 & "的一卡通!"
         Exit Function
    End If
    '一卡通结算（修改单据时因为没有读卡，无法确定使用了哪种一卡通，所以不支持修改功能)
    Dim intCardType As Integer, strSwapNO As String
    If Not mobjICCard.PaymentSwap(dblMoney, dbl余额, intCardType, Val("" & mrsOneCard!医院编码), mCurCardPay.str刷卡卡号, strSwapNO, mlng结算ID, mlng病人ID) Then
         strErrMsg = "一卡通结算失败"
        Exit Function
    End If
    gstrSQL = "zl_一卡通结算_Update(" & 0 & ",'" & mCurCardPay.str结算方式 & "','" & mCurCardPay.str刷卡卡号 & "','" & intCardType & "','" & strSwapNO & "'," & dbl余额 & "," & mlng结算ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    zlOneCardPrayMoney = True
 End Function
Private Function zlInterfacePrayMoney(ByRef cllUpdate As Collection, _
    ByRef cllThreeSwap As Collection, dblMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:接口支付金额
    '出参:cllPro-修改三方交易数据
    '        cll三方交易-增加三交方易数据
    '返回:支付成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-17 13:34:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng结帐ID As Long, strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    
    If mCurCardPay.lng医疗卡类别ID = 0 And mCurCardPay.lng医疗卡类别ID = 0 Then zlInterfacePrayMoney = True: Exit Function
    If cbo支付方式.ItemData(cbo支付方式.ListIndex) <> -1 Then zlInterfacePrayMoney = True: Exit Function
    'zlPaymentMoney(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    ByVal bln消费卡 As Boolean, _
    ByVal strCardNo As String, ByVal strBalanceIDs As String, _
    byval  strPrepayNos as string , _
    ByVal dblMoney As Double, _
    ByRef strSwapGlideNO As String, _
    ByRef strSwapMemo As String, _
    Optional ByRef strSwapExtendInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:帐户扣款交易
    '入参:frmMain-调用的主窗体
    '        lngModule-调用模块号
    '        strBalanceIDs-结帐ID,多个用逗号分离
    '        strPrepayNos-缴预交时有效. 预交单据号,多个用逗号分离
    '       strCardNo-卡号
    '       dblMoney-支付金额
    '出参:strSwapGlideNO-交易流水号
    '       strSwapMemo-交易说明
    '       strSwapExtendInfor-交易扩展信息: 格式为:项目名称1|项目内容2||…||项目名称n|项目内容n
    '返回:扣款成功,返回true,否则返回Flase
    '说明:
    '   在所有需要扣款的地方调用该接口,目前规划在:收费室；挂号室;自助查询机;医技工作站；药房等。
    '   一般来说，成功扣款后，都应该打印相关的结算票据，可以放在此接口进行处理.
    '   在扣款成功后，返回交易流水号和相关备注说明；如果存在其他交易信息，可以放在交易说明中以便退费.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gobjSquare.objSquareCard.zlPaymentMoney(Me, mlngModule, mCurCardPay.lng医疗卡类别ID, mCurCardPay.bln消费卡, mCurCardPay.str刷卡卡号, mstr结帐IDs, mCurCardPay.strNo, dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then Exit Function
    '更新三交交易数据
     If mCurCardPay.lng医疗卡类别ID <> 0 And mlng结算ID <> 0 And cbo支付方式.Visible Then
        mCurCardPay.str交易流水号 = strSwapGlideNO
        mCurCardPay.str交易说明 = strSwapMemo
        If mCurCardPay.bln消费卡 = False Then
            Call zlAddUpdateSwapSQL(False, mstr结帐IDs, mCurCardPay.lng医疗卡类别ID, mCurCardPay.bln消费卡, mCurCardPay.str刷卡卡号, strSwapGlideNO, strSwapMemo, cllUpdate, 2)
        End If
        Call zlAddThreeSwapSQLToCollection(False, mstr结帐IDs, mCurCardPay.lng医疗卡类别ID, mCurCardPay.bln消费卡, mCurCardPay.str刷卡卡号, strSwapExtendInfor, cllThreeSwap)
    End If
    zlInterfacePrayMoney = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function ChargeOver(ByVal blnNotCommit As Boolean, ByVal dbl退支票额 As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:收费完成
    '入参:blnNotCommit-是否没有进行事务提交，完成时再提交事务(原因是对普通病人进行一次提交)
    '编制:刘兴洪
    '日期:2011-08-15 15:50:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim dbl缴款 As Double, dbl找补 As Double
    Dim str收费结算 As String, dbl预存款 As Double
    dbl预存款 = 0
    str收费结算 = Get收费结算(dbl预存款)
    On Error GoTo errHandle
    If mCurCardPay.int性质 = 1 And mblnCur连续 = False Then
        dbl缴款 = Val(txt缴款.Text)
        dbl找补 = Val(txt找补.Text)
    End If
    If dbl缴款 = 0 Then
        dbl缴款 = 0: dbl找补 = 0
    End If
    'Zl_门诊收费结算_完成收费
    strSQL = "Zl_门诊收费结算_完成收费("
    '  病人id_In     门诊费用记录.病人id%Type,
    strSQL = strSQL & "" & mlng病人ID & ","
    '  结算序号id_In 病人预交记录.结算序号%Type,
    strSQL = strSQL & "" & mlng结算ID & ","
    '  缴款_In       病人预交记录.缴款%Type := Null,
    strSQL = strSQL & "" & dbl缴款 & ","
    '  找补_In       病人预交记录.找补%Type := Null,
    strSQL = strSQL & "" & dbl找补 & ","
    '  误差金额_In   门诊费用记录.实收金额%Type := Null,
    strSQL = strSQL & "" & mCurCarge.dbl本次误差费 & ","
    '  结算方式_In   病人预交记录.结算方式%Type := Null,
    strSQL = strSQL & "'" & mCurCardPay.str结算方式 & "',"
    '  预存款_In     病人预交记录.冲预交%Type := Null,
    strSQL = strSQL & "" & dbl预存款 & ","
    '  退支票额_In   病人预交记录.冲预交%Type := Null,
    strSQL = strSQL & "" & dbl退支票额 & ","
    '  收费结算_In Varchar2:=Null
    strSQL = strSQL & "'" & str收费结算 & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    mdbl缴款金额 = dbl缴款: mdbl找补 = dbl找补
    If blnNotCommit Then gcnOracle.CommitTrans
    ChargeOver = True
    Exit Function
errHandle:
    If blnNotCommit Then gcnOracle.RollbackTrans
    Call ErrCenter
End Function
Private Sub Show误差金额(ByVal bln预交 As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示误差金额
    '入参:bln预交-预交额
    '编制:刘兴洪
    '日期:2011-09-30 15:40:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, dbl退支票额 As Double
    Dim dbl剩余金额 As Double, dblTemp As Double
    mCurCarge.dbl本次误差费 = 0
    dblMoney = IIf(bln预交, Val(txt冲预交.Text), IIf(mblnYB退款, -1, 1) * Val(txt缴款.Text))
    dbl退支票额 = 0
    dbl剩余金额 = RoundEx(mCurCarge.dbl当前未付 - dblMoney, 6)
    If bln预交 Then
        dblMoney = Val(txt冲预交.Text)
        mCurCarge.dbl本次误差费 = -1 * (mCurCarge.dbl本次实收 - mCurCarge.dbl本次已付合计 - RoundEx(mCurCarge.dbl当前未付, 2))
    ElseIf mCurCardPay.int性质 = 1 Then
        dblTemp = IIf(dblMoney = 0, dbl剩余金额, mCurCarge.dbl当前未付): dbl剩余金额 = 0
        If mintInsure > 0 Then  '问题:43855
            If mInsurePara.分币处理 Then
                dblMoney = CentMoney(CCur(dblTemp))
            Else
                dblMoney = Format(dblTemp, "0.00")
            End If
        Else
                dblMoney = CentMoney(CCur(dblTemp))
        End If
        mCurCarge.dbl本次误差费 = -1 * (mCurCarge.dbl本次实收 - mCurCarge.dbl本次已付合计 - dblMoney)
    ElseIf mCurCardPay.bln支票 Then
        '只有现金才有误差费
'        If dbl剩余金额 < 0 Then
'            dbl退支票额 = -1 * Val(txt找补.Text)
'            mCurCarge.dbl本次误差费 = Format(-1 * (mCurCarge.dbl本次实收 - mCurCarge.dbl本次已付合计 - dblMoney - dbl退支票额), gstrDec)
'        End If
    Else
        '只有现金才有误差费
        'mCurCarge.dbl本次误差费 = Format(-1 * (mCurCarge.dbl本次实收 - mCurCarge.dbl本次已付合计 - RoundEx(mCurCarge.dbl当前未付, 2)), gstrDec)
    End If
    If mblnCur连续 And Val(txt缴款.Text) = 0 Then
'        dblMoney = mCurCarge.dbl当前未付
'        mCurCarge.dbl本次误差费 = Format(-1 * (mCurCarge.dbl本次实收 - mCurCarge.dbl本次已付合计 - RoundEx(mCurCarge.dbl当前未付, 2)), gstrDec)
'        dbl剩余金额 = 0
    End If
    '问题:47637
    '未进行医保结算前,不显示误差
    If mintInsure <> 0 And mblnYbBalanced = False Then mCurCarge.dbl本次误差费 = 0
    mCurCarge.dbl本次误差费 = Format(mCurCarge.dbl本次误差费, gstrDec)
    pic误差.Visible = mCurCarge.dbl本次误差费 <> 0
    lbl误差额.Caption = Format(mCurCarge.dbl本次误差费, gstrDec)
End Sub
Private Function zlCheckMulitInterfaceNumValied(Optional bln预交 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是正同时存在两种以上接口(不含两种)
    '返回:不含两种以上接口的,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-02-07 15:07:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCount As Integer, i As Long, int性质 As Integer, str结算方式 As String
    Dim varData As Variant, strErrMsg As String
    On Error GoTo errHandle
    strErrMsg = ""
    If bln预交 Or mCurCardPay.lng医疗卡类别ID = 0 Or cbo支付方式.ItemData(cbo支付方式.ListIndex) <> -1 Then
        zlCheckMulitInterfaceNumValied = True
        Exit Function
    End If
    
    '医保算一个接口
    If mintInsure <> 0 And mblnYbBalanced Then intCount = intCount + 1: strErrMsg = strErrMsg & "医保结算:" & txt医保.Text
   With vsBlance
        For i = 1 To .Rows - 1
            str结算方式 = Trim(.TextMatrix(i, .ColIndex("支付方式")))
            int性质 = Val(.RowData(i))
            '.rowdata:0-普通的结算方式-1-医保结算;2-三方接口交易;3-一卡通结算;4-预存款
            If InStr("23", int性质) > 0 Then
                If int性质 = 3 Then intCount = intCount + 1:
                If int性质 = 2 Then '三方接口
                    ' ' 医疗卡类别ID|消费卡(1, 0) |自制卡| 接口名称
                    varData = Split(.Cell(flexcpData, i, .ColIndex("支付方式")) & "|||||", "|")
                    If Val(varData(1)) = 1 Then '消费卡
                        '自制卡,不作限制
                        If Val(varData(2)) = 0 Then intCount = intCount + 1:  strErrMsg = strErrMsg & vbCrLf & varData(3) & ":" & .TextMatrix(i, .ColIndex("支付金额"))
                    Else
                         intCount = intCount + 1: strErrMsg = strErrMsg & vbCrLf & varData(3) & ":" & .TextMatrix(i, .ColIndex("支付金额"))
                    End If
                End If
            End If
        Next
    End With
    If intCount > 2 Then
        Call MsgBox("注意:" & vbCrLf & "   本系统目前只支持两种以下接口,现在已经存在如下接口交易:" & vbCrLf & strErrMsg, vbInformation + vbOKOnly)
        Exit Function
    End If
    zlCheckMulitInterfaceNumValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function
Private Function SaveCharge(Optional bln预交 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存结算数据
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-14 17:38:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHaveMoney As Boolean, dbl剩余金额 As Double, strSQL As String
    Dim dblMoney As Double, strErrMsg As String, dbl退支票额 As Double
    Dim i As Integer, blnFind As Boolean, cllPro As Collection
    Dim str消费卡结算 As String, j As Long
    Dim strCardNo As String, dblTemp As Double, blnNotCommit As Boolean '不进行事务提交
    Dim cllUpdate As Collection, cllThreeSwap As Collection
    Dim blnSaveBilling As Boolean   '当前事务保存单据
    
    On Error GoTo errHandle
    blnSaveBilling = False
    If zlCheckMulitInterfaceNumValied = False Then Exit Function
    mstrBalances = "" '问题:42791
    mdbl现金 = 0
    dblMoney = IIf(bln预交, Val(txt冲预交.Text), IIf(mblnYB退款, -1, 1) * Val(txt缴款.Text))
    dbl退支票额 = 0
    dbl剩余金额 = mCurCarge.dbl当前未付 - dblMoney
    If bln预交 Then
        dblMoney = Val(txt冲预交.Text)
        mstrBalances = mstrBalances & "|冲预交:" & dblMoney
        '问题:58344
        '检查是否当前支付金额为负数,是负数时,需要提醒操作员(主要是医保结算时可能大于本身单据的费用)
        If mblnYB退款 Then
              Call MsgBox("注意:" & vbCrLf & "    当前处于退款方式,不允许使用预交款!", vbExclamation + vbOKOnly + vbDefaultButton2, gstrSysName)
              Exit Function
        End If
    ElseIf mCurCardPay.int性质 = 1 Then
        dblTemp = IIf(dblMoney = 0, dbl剩余金额, mCurCarge.dbl当前未付): dbl剩余金额 = 0
        If mintInsure > 0 Then  '问题:43855
            If gclsInsure.GetCapability(support分币处理, , mintInsure) Then
                dblMoney = CentMoney(CCur(dblTemp))
            Else
                dblMoney = Format(dblTemp, "0.00")
            End If
        Else
                dblMoney = CentMoney(CCur(dblTemp))
        End If
        '问题:58344
        '检查是否当前支付金额为负数,是负数时,需要提醒操作员(主要是医保结算时可能大于本身单据的费用)
        If mblnYB退款 Then
             If MsgBox("注意:" & vbCrLf & "    未付部分为退款,你是否真的要退『" & mCurCardPay.str结算方式 & ":" & Abs(dblMoney) & "』给病人?" & vbCrLf & IIf(Val(txt缴款.Text) <> 0, "  当前退给病人总额:" & txt缴款.Text & vbCrLf & "  当前应收回总额:" & txt找补.Text, ""), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
             If Abs(dblMoney) < Abs(lbl剩余自付.Caption) Then
                Call MsgBox("注意:" & vbCrLf & "    未付部分为退款,你不能进行多次退款操作," & vbCrLf & "当前退金额(" & Format(dblMoney, "0.00") & ")必须大于剩余金额(" & lbl剩余自付.Caption & ")!", vbInactiveBorder + vbOKOnly + vbDefaultButton2, gstrSysName)
                Exit Function
             End If
        End If
        mdbl现金 = dblMoney
        If Val(txt缴款.Text) <> 0 Then
            mstrBalances = mstrBalances & "|缴款:" & IIf(mblnYB退款, -1, 1) * Val(txt缴款.Text) & ":1"
            mstrBalances = mstrBalances & "|找补:" & IIf(mblnYB退款, -1, 1) * Val(txt找补.Text) & ":2"
        End If
        mstrBalances = mstrBalances & "|" & mCurCardPay.str结算方式 & ":" & dblMoney
    ElseIf mCurCardPay.bln支票 Then
        mstrBalances = mstrBalances & "|" & mCurCardPay.str结算方式 & ":" & dblMoney
        '问题:58344
        '检查是否当前支付金额为负数,是负数时,需要提醒操作员(主要是医保结算时可能大于本身单据的费用)
        If mblnYB退款 Then
             If MsgBox("注意:" & vbCrLf & "    未付部分为退款,你是否真的要退『" & mCurCardPay.str结算方式 & ":" & Abs(dblMoney) & "』给病人?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
             If Abs(dblMoney) <> Abs(lbl剩余自付.Caption) Then
                Call MsgBox("注意:" & vbCrLf & "    未付部分为退款,当前退金额(" & Format(Abs(dblMoney), "0.00") & ")必须等于剩余金额(" & Abs(Val(lbl剩余自付.Caption)) & ")!", vbInactiveBorder + vbOKOnly + vbDefaultButton2, gstrSysName)
                Exit Function
             End If
        Else
            
            If dbl剩余金额 < 0 Then
                If mstr退支票 = "" Then
                    MsgBox "在结算方式中没有设置应付款的结算方式,不能进行退支票处理", vbOKOnly + vbInformation, gstrSysName
                    Exit Function
                End If
                dbl退支票额 = -1 * Val(txt找补.Text)
                mstrBalances = mstrBalances & "|" & mstr退支票 & ":" & -1 * dbl退支票额 & ":2"
            End If
        End If
    Else
        '问题:58344
        '检查是否当前支付金额为负数,是负数时,需要提醒操作员(主要是医保结算时可能大于本身单据的费用)
        If mblnYB退款 Then
             If MsgBox("注意:" & vbCrLf & "    未付部分为退款,你是否真的要退『" & mCurCardPay.str结算方式 & ":" & Abs(dblMoney) & "』给病人?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
             If Abs(dblMoney) <> Abs(lbl剩余自付.Caption) Then
                Call MsgBox("注意:" & vbCrLf & "    未付部分为退款,当前退金额(" & Format(Abs(dblMoney), "0.00") & ")必须等于剩余金额(" & Abs(Val(lbl剩余自付.Caption)) & ")!", vbInactiveBorder + vbOKOnly + vbDefaultButton2, gstrSysName)
                Exit Function
             End If
        End If
        mstrBalances = mstrBalances & "|" & mCurCardPay.str结算方式 & ":" & dblMoney
    End If
    
    If mblnCur连续 And Val(txt缴款.Text) = 0 Then
        If mCurCardPay.int性质 <> 1 Or dblMoney = 0 Then
            dblMoney = mCurCarge.dbl当前未付
        End If
        dbl剩余金额 = 0
    End If
    
    Call Show误差金额(bln预交)
    If mCurCardPay.int性质 = 1 Then
        If Abs(mCurCarge.dbl本次误差费) > 1.5 Then
            Call MsgBox("误差过大,请检查是否正确!", vbInformation + vbOKOnly, gstrSysName)
            Exit Function
        End If
    End If
    'mCurCarge.dbl本次误差费 = Format(-1 * (mCurCarge.dbl本次实收 - mCurCarge.dbl本次已付合计 - dblMoney - dbl退支票额), gstrDec)
    '误差不能大于10块钱
    If dbl剩余金额 > 0 Then blnHaveMoney = True
    With vsBlance
        blnFind = False
        For i = 1 To .Rows - 1
            If bln预交 Then
                If Val(.RowData(i)) = 4 Then blnFind = True
            ElseIf mCurCardPay.bln消费卡 And mCurCardPay.bln自制卡 Then
                '消费卡,已经检查,不用再处理
            Else
                If .TextMatrix(i, .ColIndex("支付方式")) = mCurCardPay.str结算方式 Then
                    blnFind = True
                End If
            End If
            mstrBalances = mstrBalances & "|" & .TextMatrix(i, .ColIndex("支付方式")) & ":" & .TextMatrix(i, .ColIndex("支付金额"))
        Next
        
        If blnFind Then
            If bln预交 Then
                MsgBox "已经用预存款支付,只有删除预存款后才能支付!", vbOKOnly + gstrSysName
            Else
                MsgBox mCurCardPay.str结算方式 & " 已经支付了,不能再用" & mCurCardPay.str结算方式 & "进行支付", vbOKOnly + vbDefaultButton1, gstrSysName
            End If
            Exit Function
        End If
    End With
    
    If blnHaveMoney = False And dblMoney = 0 Then
        GoTo GoOver:
    End If
    
    Set cllPro = New Collection
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    str消费卡结算 = ""  '卡类别ID|卡号|消费卡ID|消费金额||....
    If mCurCardPay.bln消费卡 Then
        If mcllCurSquareBalance Is Nothing Then Exit Function
        If mcllCurSquareBalance.Count = 0 Then Exit Function
        
        For j = 1 To mcllCurSquareBalance.Count
            ' array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文)
            str消费卡结算 = str消费卡结算 & "||" & Val(mcllCurSquareBalance(j)(0))
            str消费卡结算 = str消费卡结算 & "|" & mcllCurSquareBalance(j)(3)
            str消费卡结算 = str消费卡结算 & "|" & Val(mcllCurSquareBalance(j)(1))
            str消费卡结算 = str消费卡结算 & "|" & Val(mcllCurSquareBalance(j)(2))
        Next
        If str消费卡结算 <> "" Then str消费卡结算 = Mid(str消费卡结算, 3)
    End If
    Err = 0: On Error GoTo ErrCommit:
    If Not (bln预交 Or mCurCardPay.lng医疗卡类别ID = 0 Or cbo支付方式.ItemData(cbo支付方式.ListIndex) <> -1) Then
        '第三方接口的相关结算,需要先处理接口数据
        blnNotCommit = False
        If Not mblnYbBalanced And mlng结算ID = 0 Then
            blnNotCommit = True
            If SaveBill(blnNotCommit) = False Then
                blnNotCommit = False: mlng结算ID = 0: Exit Function
            End If
            blnSaveBilling = True
        End If
        
        'Zl_门诊收费结算_Modify
        strSQL = "Zl_门诊收费结算_Modify("
        '  操作类型_In   Number,
        '--操作类型_In:0-正常收费;
        '--            1-冲预交(结算方式为NULL,结算金额<>0);
        '--            2-医保结算:如果是医保结算,结算方式_IN可以为多个;
        '--            3-消费卡批量结算(结算方式_IN格式为:卡类别ID|卡号|消费卡ID|消费金额||....)
        strSQL = strSQL & IIf(bln预交, "1", IIf(mCurCardPay.bln消费卡, "3", "0")) & ","
        '  病人id_In     门诊费用记录.病人id%Type,
        strSQL = strSQL & mlng病人ID & ","
        '  结算序号id_In 病人预交记录.结算序号%Type,
        strSQL = strSQL & mlng结算ID & ","
        '  结算方式_In   Varchar2,
        If bln预交 Then
            strSQL = strSQL & "NULL" & ","
        ElseIf mCurCardPay.bln消费卡 Then
            strSQL = strSQL & "'" & str消费卡结算 & "'" & ","
        Else
            strSQL = strSQL & "'" & mCurCardPay.str结算方式 & "'" & ","
        End If
        '  结算金额_In   病人预交记录.冲预交%Type,
        strSQL = strSQL & dblMoney & ","
        ' 退支票额_In   病人预交记录.冲预交%Type,
        strSQL = strSQL & dbl退支票额 & ","
        '  摘要_In       病人预交记录.摘要%Type := Null,
        strSQL = strSQL & "'" & Trim(txt摘要.Text) & "',"
        '  结算号码_In   病人预交记录.结算号码%Type := Null,
        strSQL = strSQL & "" & IIf(mCurCardPay.bln支票 Or txt结算号码.Visible, "'" & Trim(txt结算号码.Text) & "'", "NULL") & ","
        '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
        strSQL = strSQL & "" & IIf(mCurCardPay.lng医疗卡类别ID = 0, "NULL", mCurCardPay.lng医疗卡类别ID) & ","
        '  消费卡_In     Integer := 0,
        strSQL = strSQL & "" & IIf(mCurCardPay.bln消费卡, 1, 0) & ","
        '  卡号_In       病人预交记录.卡号%Type := Null,
        strSQL = strSQL & "" & IIf(mCurCardPay.str刷卡卡号 <> "", "'" & mCurCardPay.str刷卡卡号 & "'", "NULL") & ","
        '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
        strSQL = strSQL & "NULL" & ","
        '  交易说明_In   病人预交记录.交易说明%Type := Null
        strSQL = strSQL & "NULL" & ")"
        zlAddArray cllPro, strSQL
        Call zlExecuteProcedureArrAy(cllPro, Me.Caption, True, blnNotCommit)
        If Not mCurCardPay.bln消费卡 Then
            '消费卡不再调用接口
             If zlInterfacePrayMoney(cllUpdate, cllThreeSwap, dblMoney) = False Then
                 '问题:47637
                  If Not (mblnYbBalanced Or mblnThreeInterface) Or blnSaveBilling Then mlng结算ID = 0
                    gcnOracle.RollbackTrans: Exit Function
            End If
        End If
        
        Err = 0: On Error GoTo ErrUpdate:
        '一卡通交易
        If zlOneCardPrayMoney(dblMoney, strErrMsg) = False Then
            gcnOracle.RollbackTrans
            MsgBox strErrMsg, vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
        gcnOracle.CommitTrans
        Call zlExecuteProcedureArrAy(cllUpdate, Me.Caption)
        blnNotCommit = False: mblnThreeInterface = True
        Call SetCtrlVisible
        On Error GoTo ErrOthers:
        Call zlExecuteProcedureArrAy(cllThreeSwap, Me.Caption)
    End If
GoOver:
    If mintInsure <> 0 Then
        If Not (bln预交 Or mCurCardPay.lng医疗卡类别ID <> 0 _
            Or mCurCardPay.blnOneCard) Then
            '只有医保病人才会出现重新较对的情况,因此才会重新计算本次应缴的情况
            '主要是更改连续收费的问题
            mdbl本次应缴 = mdbl本次应缴 + dblMoney
        End If
    End If
    
    If Not blnHaveMoney Then
        If mlng结算ID = 0 Then
            blnNotCommit = True
            If SaveBill(blnNotCommit) = False Then
                blnNotCommit = False: mlng结算ID = 0: Exit Function
            End If
            blnSaveBilling = True
        End If
        If ChargeOver(blnNotCommit, dbl退支票额) = False Then
            If blnNotCommit Or blnSaveBilling Then mlng结算ID = 0
            Exit Function
        End If
        Call WhriteTotalDataToReCord(IIf(bln预交, dblMoney, 0), IIf(Not bln预交, dblMoney, 0), dbl退支票额)
        mblnOK = True
        SaveCharge = True: mblnUnloaded = True
        
        Unload Me:
        Exit Function
    End If
    mstrBalances = ""
    If Not bln预交 And mCurCardPay.int性质 = 1 Then
       '现金
        SaveCharge = True: Exit Function
    End If
    Err = 0: On Error GoTo errHandle:
    With vsBlance
        If mCurCardPay.bln消费卡 Then
            If mcllSquareBalance Is Nothing Then Set mcllSquareBalance = New Collection
            For j = 1 To mcllCurSquareBalance.Count
                '当前刷卡数据(array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文))
                mcllSquareBalance.Add mcllCurSquareBalance(j)
                If Not (.Rows = 2 And Trim(.TextMatrix(1, .ColIndex("支付方式"))) = "") Then
                    .Rows = .Rows + 1
                    .RowPosition(.Rows - 1) = 1
                End If
                .RowData(1) = 0
                .TextMatrix(1, .ColIndex("支付方式")) = mCurCardPay.str结算方式
                 ' 医疗卡类别ID|消费卡(1, 0) |自制卡| 接口名称
                .Cell(flexcpData, 1, .ColIndex("支付方式")) = Val(mcllCurSquareBalance(j)(0)) & "|" & 1 & "|" & IIf(mCurCardPay.bln自制卡, 1, 0) & "|" & mCurCardPay.str名称
                .RowData(1) = 2
                strCardNo = Trim(mcllCurSquareBalance(j)(3))
                .TextMatrix(1, .ColIndex("卡号")) = IIf(mCurCardPay.bln卡号密文, String(Len(strCardNo), "*"), strCardNo)
                .Cell(flexcpData, 1, .ColIndex("卡号")) = strCardNo
                .TextMatrix(1, .ColIndex("支付金额")) = Format(Val(mcllCurSquareBalance(j)(2)), "0.00")
                .TextMatrix(1, .ColIndex("结算号码")) = ""
                .TextMatrix(1, .ColIndex("备注")) = ""
                mCurCarge.dbl本次已付合计 = RoundEx(mCurCarge.dbl本次已付合计 + Val(mcllCurSquareBalance(j)(2)), 6)
                mCurCarge.dbl当前未付 = RoundEx(mCurCarge.dbl当前未付 - Val(mcllCurSquareBalance(j)(2)), 6)
            Next
        Else
            If Not (.Rows = 2 And Trim(.TextMatrix(1, .ColIndex("支付方式"))) = "") Then
                .Rows = .Rows + 1
                .RowPosition(.Rows - 1) = 1
            End If
            .RowData(1) = 0
            strCardNo = mCurCardPay.str刷卡卡号
            If bln预交 Then
                .TextMatrix(1, .ColIndex("支付方式")) = "预存款"
                .RowData(1) = 4
            ElseIf mCurCardPay.lng医疗卡类别ID <> 0 Then
                .TextMatrix(1, .ColIndex("支付方式")) = mCurCardPay.str结算方式
                 ' 医疗卡类别ID|消费卡(1, 0) |自制卡| 接口名称
                .Cell(flexcpData, 1, .ColIndex("支付方式")) = mCurCardPay.lng医疗卡类别ID & "|" & IIf(mCurCardPay.bln消费卡, 1, 0) & "|" & IIf(mCurCardPay.bln自制卡, 1, 0) & "|" & mCurCardPay.str名称
                .RowData(1) = 2
                strCardNo = gobjSquare.objSquareCard.zlGetCardNODencode(mCurCardPay.str刷卡卡号, mCurCardPay.lng医疗卡类别ID, mCurCardPay.bln消费卡)
            ElseIf mCurCardPay.blnOneCard Then
                .TextMatrix(1, .ColIndex("支付方式")) = mCurCardPay.str结算方式
                .RowData(1) = 3
            Else
                .TextMatrix(1, .ColIndex("支付方式")) = mCurCardPay.str结算方式
            End If
            .TextMatrix(1, .ColIndex("支付金额")) = Format(dblMoney, "0.00")
            .TextMatrix(1, .ColIndex("结算号码")) = IIf(txt结算号码.Visible, Trim(txt结算号码.Text), "")
            .TextMatrix(1, .ColIndex("备注")) = Trim(txt摘要.Text)
            
            .TextMatrix(1, .ColIndex("卡号")) = IIf(mCurCardPay.bln卡号密文, String(Len(strCardNo), "*"), strCardNo)
            .Cell(flexcpData, 1, .ColIndex("卡号")) = mCurCardPay.str刷卡卡号
            .TextMatrix(1, .ColIndex("交易流水号")) = mCurCardPay.str交易流水号
            .TextMatrix(1, .ColIndex("交易说明")) = mCurCardPay.str交易说明
            mCurCarge.dbl本次已付合计 = RoundEx(mCurCarge.dbl本次已付合计 + dblMoney, 6)
            mCurCarge.dbl当前未付 = RoundEx(mCurCarge.dbl当前未付 - dblMoney, 6)
        End If
        For i = 0 To cbo支付方式.ListCount
            '缺省定位在现金上
            If cbo支付方式.ItemData(i) = 1 Then cbo支付方式.ListIndex = i: Exit For
        Next
        Call SetControlProperty
        txt缴款.Text = ""
        If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
        Call LedDisplayBank
    End With
    Call SetDeleteVisible
    SaveCharge = True
    Exit Function
ErrCommit:
    gcnOracle.RollbackTrans
    If blnSaveBilling Then mlng结算ID = 0
    
ErrUpdate:
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    
    Exit Function
ErrOthers:
    '三方交易信息,能保存多少算多少,不存在性能问题.
    If ErrCenter = 1 Then
        gcnOracle.RollbackTrans
        Resume
    End If
    gcnOracle.CommitTrans
End Function

Private Sub txt找补_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab

End Sub

Private Sub txt找补_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lbl找补.Caption <> "找补" Then
      ''  zlCommFun.ShowTipInfo txt找补.hWnd, mstr应付款结算方式, False
    Else
        zlCommFun.ShowTipInfo txt找补.hWnd, "", False
    End If
End Sub

Private Sub vsBlance_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
      
    If OldRow = NewRow Then Exit Sub
    If NewRow < 0 Then Exit Sub
    Call SetDeleteVisible
End Sub
Private Sub SetDeleteVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置删除控件的visible属性
    '编制:刘兴洪
    '日期:2011-09-20 10:42:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
     Dim int性质 As Integer
     If vsBlance.Row < 0 Then
        int性质 = -1
     Else
        int性质 = Val(vsBlance.RowData(vsBlance.Row))
    End If
     '.rowdata:0-普通的结算方式-1-医保结算;2-三方接口交易;3-一卡通结算;4-预存款
    cmdDel.Visible = (int性质 = 0 Or int性质 = 4) And mbytFunc <> EM_异常作废
End Sub
Private Sub SetWindowsSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置窗体大小
    '编制:刘兴洪
    '日期:2011-09-15 11:26:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If OS.IsDesinMode Then Exit Sub
    '最小窗体尺寸
    With gWinRect
        .MaxW = Me.Width
        .MaxH = Screen.Height * Screen.TwipsPerPixelY
        .MinH = Me.Height
        .MinW = Me.Width
    End With
    glngOld = GetWindowLong(hWnd, GWL_WNDPROC)
    Call SetWindowLong(hWnd, GWL_WNDPROC, AddressOf SetWindowResizeWndMessage)
End Sub

Private Function zlCheckDelValied(ByVal lng卡类别ID As Long, _
     ByVal strName As String, _
     ByVal bln消费卡 As Boolean, ByVal strCardNo As String, _
     ByVal strSwapNO As String, strSwapMemo As String, _
     ByRef str结帐ID As String, _
    ByVal dbl退款金额 As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查退费交易接口
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-02-08 16:40:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strXMLExend As String
    If lng卡类别ID = 0 Then zlCheckDelValied = True: Exit Function
    
    If gobjSquare.objSquareCard Is Nothing Then
        MsgBox "注意:" & vbCrLf & _
                     "      当前收费是按" & strName & " 收费的,但不存在操作的相关部件,不能退款,请与系统管理员联系!", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    
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
    '       strXMLExpend    XML IN  可选参数(扩展用).暂未传入
    '返回:退款合法,返回true,否则返回Flase
      If gobjSquare.objSquareCard.zlReturnCheck(Me, mlngModule, lng卡类别ID, bln消费卡, strCardNo, _
        "3|" & str结帐ID, dbl退款金额, strSwapNO, strSwapMemo, strXMLExend) = False Then
          zlCheckDelValied = False
          Exit Function
     End If
goEnd:
    zlCheckDelValied = True
    Exit Function
End Function

Private Function CallBackBalanceInterface(ByVal str冲销IDs As String, _
    ByVal lng卡类别ID As Long, _
    ByVal bln消费卡 As Boolean, _
    ByVal dblMoney As Double, _
    ByVal strCardNo As String, _
    ByVal strSwapNO As String, _
    ByVal strSwapMemo As String, _
    ByRef cllUpdate As Collection, _
    ByRef cllThreeSwap As Collection, ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用回退接口
    '入参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-13 10:33:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim str原结帐IDs As String, strSwapExtendInfor As String, strTemp As String
    Err = 0: On Error GoTo Errhand:
    
    '卡类别ID,卡号,是否消费卡(1-是;0-否),交易流水号,交易说明,strNO
    If lng卡类别ID = 0 Then CallBackBalanceInterface = True: Exit Function
    If Left(str冲销IDs, 1) = "," Then str冲销IDs = Mid(str冲销IDs, 2)
    strSQL = "" & _
    "    Select A.NO From 门诊费用记录 A,病人预交记录 M,Table( f_Num2list([1]))  B   " & _
    "    Where  A.记录性质 = 1 And A.记录状态=2  " & _
    "               And A.结帐ID=M.结帐ID  " & IIf(bln消费卡, " And nvl(M.结算卡序号,0)=[2]", " And nvl(M.卡类别ID,0)=[2] ") & _
    "           And A.结帐ID=B.Column_Value " & _
    "      "
   strSQL = "" & _
   "    Select /*+ RULE */ distinct  结帐ID From 门诊费用记录 Q, (" & strSQL & ") M  " & _
   "    Where Q.NO=M.NO and Q.记录性质=1 and Q.记录状态=3  "
   '61688
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str冲销IDs, lng卡类别ID)
    If rsTemp.EOF Then
        strErrMsg = "未找第三方结算交易信息，退费失败": Exit Function
    End If
    With rsTemp
        str原结帐IDs = ""
        Do While Not .EOF
            str原结帐IDs = str原结帐IDs & "," & Nvl(!结帐ID)
            .MoveNext
        Loop
    End With
    If str原结帐IDs <> "" Then str原结帐IDs = Mid(str原结帐IDs, 2)
    '81489,冉俊明,2015-1-22,退费传入冲销ID
    strSwapExtendInfor = "3|" & str冲销IDs: strTemp = strSwapExtendInfor
    
    'zlReturnMoney(frmMain As Object, ByVal lngModule As Long, _
        ByVal lngCardTypeID As Long, ByVal strCardNo As String, ByVal strBalanceIDs As String, _
        ByVal dblMoney As Double, _
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
    If gobjSquare.objSquareCard.zlReturnMoney(Me, mlngModule, lng卡类别ID, bln消费卡, strCardNo, "3|" & str原结帐IDs, dblMoney, strSwapNO, strSwapMemo, strSwapExtendInfor) = False Then Exit Function
    Set cllUpdate = New Collection: Set cllThreeSwap = New Collection
    Call zlAddUpdateSwapSQL(False, str冲销IDs, lng卡类别ID, bln消费卡, strCardNo, strSwapNO, strSwapMemo, cllUpdate)
    If strTemp <> strSwapExtendInfor Then
        Call zlAddThreeSwapSQLToCollection(False, str冲销IDs, lng卡类别ID, bln消费卡, strCardNo, strSwapExtendInfor, cllThreeSwap)
    End If
    CallBackBalanceInterface = True
Errhand:
End Function
Private Function zlGetInvoiceGroupUseID(ByRef lng领用ID As Long, _
    Optional intNum As Integer = 1, Optional strInvoiceNO As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取票据的领用ID
    '入参:lng领用ID-领用id
    '       intNum-页数
    '       strInvoiceNO-输入的发票号
    '出参:lng领用ID-领用ID
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-04-29 15:36:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    lng领用ID = GetInvoiceGroupID(1, intNum, lng领用ID, mlngShareUseID, strInvoiceNO, mstrUseType)
    If lng领用ID <= 0 Then
        Select Case lng领用ID
            Case 0 '操作失败
            Case -1
                If Trim(mstrUseType) = "" Then
                    MsgBox "你没有自用和共用的收费票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                Else
                    MsgBox "你没有自用和共用的『" & mstrUseType & "』收费票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                End If
                Exit Function
            Case -2
                If Trim(mstrUseType) = "" Then
                    MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                Else
                    MsgBox "本地的共用票据的『" & mstrUseType & "』收费票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                End If
                Exit Function
        End Select
    End If
    zlGetInvoiceGroupUseID = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function Get实收金额(ByVal strNo As String) As Double
    Dim i As Long
    On Error GoTo errHandle
    If Not mrsBlance Is Nothing Then
        gstrSQL = "" & _
        "   Select NO,nvl(sum(A.冲预交),0) as 冲预交" & _
        "   From 病人预交记录 A,结算方式 B " & _
        "   Where a.记录性质=3 And A.结算序号=[1]   " & _
        "               And ( 结算方式=b.名称 and b.性质 in (3,4) OR 结算方式 is null ) "
        Set mrsBlance = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng结算ID)
    End If
    mrsBlance.Filter = "NO='" & strNo & "'"
    If Not mrsBlance.EOF Then
        Get实收金额 = Val(Nvl(mrsBlance!冲预交))
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub SetControlEnabled()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件的属性
    '编制:刘兴洪
    '日期:2012-02-03 15:08:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnEdit As Boolean
    blnEdit = (mintInsure = 0 Or mintInsure <> 0 And mblnYbBalanced = True) And mbytFunc <> EM_异常作废
    blnEdit = blnEdit Or mbytFunc = EM_重新收费
    picPay.Enabled = blnEdit
    txt冲预交.Enabled = blnEdit
    txt缴款.Enabled = blnEdit
    txt找补.Enabled = blnEdit
    txt结算号码.Enabled = blnEdit
    txt摘要.Enabled = blnEdit
    
    '控制显示颜色
    txt冲预交.BackColor = IIf(txt冲预交.Enabled, &H80000005, Me.BackColor)
    txt缴款.BackColor = IIf(txt缴款.Enabled, &H80000005, Me.BackColor)
    txt找补.BackColor = IIf(txt找补.Enabled, &H80000005, Me.BackColor)
    txt结算号码.BackColor = IIf(txt结算号码.Enabled, &H80000005, Me.BackColor)
    txt摘要.BackColor = IIf(txt摘要.Enabled, &H80000005, Me.BackColor)
End Sub
Public Function Get收费结算(ByRef dbl预存款 As Double) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取收费结算数据
    '出参:dbl预存款-返回本次支付的预款
    '返回:收费用结算方式,格式如下:
    '       结算方式|结算金额|结算号码|结算摘要||.....",注意无结算号码和摘要时要用空格填充
    '编制:刘兴洪
    '日期:2012-02-06 10:58:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str结算方式 As String, i As Integer, int性质 As Integer
    Dim str收费结算 As String
    Dim dblMoney As Double
    '结算方式|结算金额|结算号码|结算摘要||.....",注意无结算号码和摘要时要用空格填充
    '收费完成
    str收费结算 = ""
    With vsBlance
        dbl预存款 = Val(txt冲预交.Text)
        For i = .Rows - 1 To 1 Step -1
            str结算方式 = Trim(.TextMatrix(i, .ColIndex("支付方式")))
            int性质 = Val(.RowData(i))
            If str结算方式 <> "" And int性质 = 0 Then
                '.rowdata:0-普通的结算方式-1-医保结算;2-三方接口交易;3-一卡通结算;4-预存款
                str收费结算 = str收费结算 & "||" & str结算方式
                str收费结算 = str收费结算 & "|" & Val(.TextMatrix(i, .ColIndex("支付金额")))
                str收费结算 = str收费结算 & "|" & IIf(Trim(.TextMatrix(i, .ColIndex("结算号码"))) = "", " ", Trim(.TextMatrix(i, .ColIndex("结算号码"))))
                str收费结算 = str收费结算 & "|" & IIf(Trim(.TextMatrix(i, .ColIndex("备注"))) = "", " ", Trim(.TextMatrix(i, .ColIndex("备注"))))
            End If
        Next
        If (mCurCardPay.lng医疗卡类别ID = 0 Or cbo支付方式.ItemData(cbo支付方式.ListIndex) <> -1) Then
            dblMoney = IIf(mblnYB退款, -1, 1) * Val(txt缴款.Text)
            If mCurCardPay.int性质 = 1 Then
                dblMoney = mdbl现金
            ElseIf mblnCur连续 And dblMoney = 0 Then
                dblMoney = mCurCarge.dbl当前未付
            End If
            If dblMoney <> 0 Then
                str收费结算 = str收费结算 & "||" & mCurCardPay.str结算方式
                If mCurCardPay.int性质 = 1 Then
                    '现金
                    str收费结算 = str收费结算 & "|" & dblMoney
                    str收费结算 = str收费结算 & "| "
                    str收费结算 = str收费结算 & "|" & IIf(Trim(txt摘要) = "", " ", Trim(txt摘要))
                Else
                    str收费结算 = str收费结算 & "|" & dblMoney
                    str收费结算 = str收费结算 & "|" & IIf(Trim(txt结算号码) = "", " ", Trim(txt结算号码))
                    str收费结算 = str收费结算 & "|" & IIf(Trim(txt摘要) = "", " ", Trim(txt摘要))
                End If
            End If
        End If
    End With
    If str收费结算 <> "" Then str收费结算 = Mid(str收费结算, 3)
    Get收费结算 = str收费结算
End Function
 
