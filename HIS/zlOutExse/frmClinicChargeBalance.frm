VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClinicChargeBalance 
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
   Icon            =   "frmClinicChargeBalance.frx":0000
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
      TabStop         =   0   'False
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
         Begin XtremeSuiteControls.ShortcutCaption stcTittleTotal 
            Height          =   420
            Left            =   15
            TabIndex        =   29
            TabStop         =   0   'False
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
         TabStop         =   0   'False
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
            MaxLength       =   30
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
            MaxLength       =   50
            MultiLine       =   -1  'True
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
         Begin XtremeSuiteControls.ShortcutCaption stcTittile 
            Height          =   450
            Left            =   15
            TabIndex        =   27
            TabStop         =   0   'False
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
      Height          =   420
      Left            =   0
      TabIndex        =   21
      Top             =   6120
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   3572
            MinWidth        =   882
            Picture         =   "frmClinicChargeBalance.frx":08CA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8599
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   582
            MinWidth        =   2
            Object.Tag             =   "用于收费预交余额显示"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   582
            MinWidth        =   2
            Object.Tag             =   "用于收费三方卡余额的显示"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   882
            MinWidth        =   882
            Picture         =   "frmClinicChargeBalance.frx":115E
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
      TabStop         =   0   'False
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
         ForeColorSel    =   -2147483640
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
         FormatString    =   $"frmClinicChargeBalance.frx":1838
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
      TabStop         =   0   'False
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
      Left            =   8208
      TabIndex        =   19
      Top             =   228
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
      Left            =   8220
      TabIndex        =   36
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
Attribute VB_Name = "frmClinicChargeBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum gChargePayType
    EM_FUN_收费 = 0
    EM_FUN_作废 = 1
    EM_FUN_重收 = 2
End Enum
Public Enum gExitMode
    EM_EX_完成 = 0
    EM_EX_暂停 = 1
    EM_EX_作废 = 2
    EM_EX_继续 = 3
    EM_EX_退出 = 4
End Enum
Private mbytFunc As gChargePayType  '0-收费;1-作废
Private mfrmMain As frmClinicCharge
Private mbytReturnMode As gExitMode
Private mbln异常作废 As Boolean
Private mblnYB退款 As Boolean '医保结算金额大于了单据结算金额
Private mbln分单据结算必须全结 As Boolean '当前异常单据是否为分单据结算必须全结产生的
Private mblnElsePersonErrBill As Boolean '是否是他人的异常单据
'------------------------------------------------------------------------------------------
'程序入口相关变量
Private mobjChargeInfor As clsClinicChargeInfor
Private mlngModule As Long, mstrPrivs As String
Private mstrYBPati As String
Private mblnOK As Boolean
Private mbln连续输入 As Boolean
Private mblnCur连续 As Boolean
Private mlngR As Long
Private mlngBrushCardTypeID As Long '在主界面中刷卡的卡类别ID,以便缺省定位在该支付类别上
Private mblnUnloaded  As Boolean
Private mblnLoad As Boolean
Private mstr退支票 As String
Private mCurCardPay As gTY_PayMoney '本次卡支付
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
Private mbln医保已报价 As Boolean
Private mstr医保结算 As String
Private mblnYbBalanced As Boolean '医保已经结算
Private mblnThreeInterface As Boolean '已经调用三方接口
Private mcur个帐余额 As Currency
Private mblnSaveBill As Boolean '单据保存成功
Private mblnCommitBill As Boolean '单据是否已经提交过
Private mblnPriceBillCommit As Boolean '划价单是否已经提交
Private mcllPriceSQL As Collection '直接收费时，先保存为划价单

Private mcllOverPro As Collection
Private mblnSavePrice As Boolean '联合医保保存为划价单
Private mrsBalance As ADODB.Recordset   '结算信息
Private mstrTittle As String '窗体标题
'----------------------------------------------------------------------------------------------
'医保相关
'当前病人险类的医保支持参数
Private Type TYPE_MedicarePAR
    不提醒缴款金额不足 As Boolean    '27536
    医保接口打印票据 As Boolean
    门诊连续收费 As Boolean
    分币处理 As Boolean
    多单据分单据结算 As Boolean '86321
    一次结算分单据退费 As Boolean '91602
End Type
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mInsurePara As TYPE_MedicarePAR
Private mrsOneCard As ADODB.Recordset
Private mrsBlance As ADODB.Recordset
'---------------------------------------------------------------------------------
Private mbln连续收费 As Boolean
'---------------------------------------------------------------------------------
Private mdbl现金 As Double, mdbl原未付 As Double
Private Const VK_RETURN = &HD
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private mblnCacheKeyReturn As Boolean   '41025:是否缓存了回车键,可能存在在收费界面刷卡中本身包含了回车,因此需要判断
Private mrsClassMoney As ADODB.Recordset
Private mcllSquareBalance As Collection '消费卡结算信息
Private mcllCurSquareBalance As Collection '当前消费卡刷卡信息
Private mblnNotChange As Boolean
Private mblnCurBrushPrepay   As Boolean '当前是否刷的预交款
    
Public Function zlChargeWin(ByVal frmMain As Object, ByVal bytFunc As gChargePayType, _
    ByVal lngModule As Long, ByVal strPrivs As String, _
    ByRef objChargeInfor As clsClinicChargeInfor, _
    Optional bytReturnMode As gExitMode = EM_EX_完成, _
    Optional bln继续输入 As Boolean, _
    Optional lngBrushCardTypeID As Long = 0, _
    Optional bln异常作废 As Boolean = False, _
    Optional blnElsePersonErrBill As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入口:表示进入支付结算窗口
    '入参:frmMain-调用的主窗体
    '       bytFunc-0-收费;1-作废
    '       lngModule -模块号
    '       strPrivs-权限串
    '       objChargeInfor-结算信息
    '       lngBrushCardTypeID-缺省的刷卡类别ID
    '       bln异常作废-异常单据作废处理(异常作废时传入):如果为true,表示针对作废的异常单据进行作废
    '       blnElsePersonErrBill-是否是他人的异常单据
    '出参:objChargeInfor.缴款金额-输入的缴款金额和找补金额(缴现金时,传出)
    '     objChargeInfor.本次应缴-医保病人,在连续收费情况下,需要重新计算本次的应缴额
    '     objChargeInfor.收费结算:返回本次收费的结算方式,格式如下:
    '                       金额:缴款标志(1-缴款;2-找补)|结算方式1:金额1:缴款标志(1-缴款;2-找补)|...
    '        bln继续输入-是否继续录入的票据
    '        bytReturnMode-返回操作模式(0-正常收费完成,1-暂停收费;2-本次作废收费;3-继续输入)
    '返回:完成收费,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-12 09:59:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mobjChargeInfor = objChargeInfor: Set mcllOverPro = Nothing
    Set mrsClassMoney = Nothing: Set mrsBalance = Nothing
    mblnYbBalanced = False: mblnThreeInterface = False: mblnOK = False
    mblnUnLoad = False: mblnUnloaded = False: mblnSaveBill = False
    mblnCommitBill = bln异常作废: mblnElsePersonErrBill = blnElsePersonErrBill
    mbln分单据结算必须全结 = False
    mblnPriceBillCommit = False
    
    mstrPrivs = strPrivs: mlngModule = lngModule
    mlngBrushCardTypeID = lngBrushCardTypeID: Set mfrmMain = frmMain
    mbln异常作废 = bln异常作废
    mbytFunc = bytFunc: mbytReturnMode = EM_EX_完成
    
    mCurCarge.dbl应缴累计 = mobjChargeInfor.应缴累计
    mbln连续输入 = mobjChargeInfor.应缴累计 <> 0
    mblnOK = False
    On Error Resume Next
    Me.Show IIf(gfrmMain Is Nothing, 0, 1), frmMain
    bln继续输入 = mbln连续输入: bytReturnMode = mbytReturnMode
    'Set objChargeInfor = mobjChargeInfor
    zlChargeWin = mblnOK
End Function

Private Function SaveFeeBilL() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存费用单据数据
    '入参:lng结帐ID-费用数据
    '     cllPro-执行的相关过程
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-06-12 14:36:17
    '说明:
    '   调用此过程时,不需要开始事务,异常时,数据回退,保存成功时,未提交数据
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPro As Collection, lng结帐ID As Long, strNos As String
    Dim cllItem As Collection, blnTransMedicare As Boolean
    Dim cllPriceSQL As Collection
    On Error GoTo errHandle
    
    If (mblnSaveBill And mblnCommitBill) Or mbytFunc = EM_FUN_重收 Then
        gcnOracle.BeginTrans
        SaveFeeBilL = True: Exit Function
    End If
    
    If mfrmMain.zlGetSaveBillSQL(lng结帐ID, cllPriceSQL, strNos, cllPro, mcllOverPro) = False Then Exit Function
    mobjChargeInfor.结帐ID = lng结帐ID
    mobjChargeInfor.结算序号 = -1 * lng结帐ID
    mobjChargeInfor.Nos = strNos
    Set mcllPriceSQL = cllPriceSQL
    
    '先提交划价单，以便不锁表（药品库存）
    blnTransMedicare = True
    If mblnPriceBillCommit = False Then
        gcnOracle.BeginTrans
        For Each cllItem In cllPriceSQL
            ExecuteProcedureArrAy cllItem, Me.Caption, True, True
        Next
        gcnOracle.CommitTrans
        mblnPriceBillCommit = True
    End If
    
    gcnOracle.BeginTrans
    For Each cllItem In cllPro '91665
        ExecuteProcedureArrAy cllItem, Me.Caption, True, True
    Next
    
    mblnSaveBill = True: SaveFeeBilL = True
    Exit Function
errHandle:
    If blnTransMedicare Then gcnOracle.RollbackTrans
    If Err.Description Like "*当前计算单价不一致*" Then
        If MsgBox("某些分批药品价格已发生变化，要自动重算价格吗？", vbYesNo + vbQuestion + vbDefaultButton1, gstrSysName) = vbYes Then
            Call mfrmMain.zlReCalcMoney(mobjChargeInfor)
            Call SetControlProperty
            Call SetCtrlVisible
            Call SetControlEnabled
            Exit Function
        End If
        Call SaveErrLog
        Exit Function
    End If
    If ErrCenter() = 1 Then
'        Resume
    End If
End Function

Private Sub DelMedicareTempNOs()
    '功能:直接收费时,删除前一个事务提交的划价单
    Dim i As Integer, varNos As Variant
    Dim strSQL As String
    
    On Error GoTo errHandle
    If mcllPriceSQL Is Nothing Then Exit Sub
    varNos = Split(mobjChargeInfor.Nos, ",")
    For i = 0 To UBound(varNos)
        If CollectionExitsValue(mcllPriceSQL, varNos(i)) Then
            strSQL = "zl_门诊划价记录_DELETE('" & varNos(i) & "')"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        End If
    Next
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

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
            grsTotal!性质 = -99
            grsTotal!结算方式 = "缴款"
            grsTotal!结算金额 = dbl缴款
        End If
        
        If dbl找补 <> 0 Then
            grsTotal.Find "结算方式='" & IIf(mCurCardPay.bln支票, "退支票", "找补") & "'", , adSearchForward, 1
            If grsTotal.EOF Then grsTotal.AddNew
            grsTotal!性质 = -98
            grsTotal!结算方式 = IIf(mCurCardPay.bln支票, "退支票", "找补")
            grsTotal!结算金额 = dbl找补
        End If
        
        For i = 1 To .Rows - 1
            str结算方式 = Trim(.TextMatrix(i, .ColIndex("支付方式")))
            int性质 = Val(.RowData(i))
            If str结算方式 <> "" Then
                '.rowdata:  0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                '性质:-99-缴款;-98-找补,0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                grsTotal.Find "结算方式='" & str结算方式 & "'", , adSearchForward, 1
                
                If grsTotal.EOF Then grsTotal.AddNew
                grsTotal!性质 = int性质
                grsTotal!结算方式 = str结算方式
                grsTotal!结算金额 = Val(Nvl(grsTotal!结算金额)) + Val(.TextMatrix(i, .ColIndex("支付金额")))
                grsTotal.Update
            End If
        Next
        
        If dbl预交 <> 0 Then
            grsTotal.Find "结算方式='预存款'", , adSearchForward, 1
            If grsTotal.EOF Then grsTotal.AddNew
            grsTotal!性质 = 1
            grsTotal!结算方式 = "预存款"
            grsTotal!结算金额 = Val(Nvl(grsTotal!结算金额)) + dbl预交
            grsTotal.Update
        End If
        If mCurCardPay.bln消费卡 And Not mcllCurSquareBalance Is Nothing Then
            For i = 1 To mcllCurSquareBalance.Count
                '当前刷卡数据(array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文))
                grsTotal.Find "结算方式='" & mCurCardPay.str结算方式 & "'", , adSearchForward, 1
                If grsTotal.EOF Then grsTotal.AddNew
                grsTotal!性质 = IIf(mCurCardPay.blnOneCard, 4, 5)
                grsTotal!结算方式 = mCurCardPay.str结算方式
                grsTotal!结算金额 = Val(Nvl(grsTotal!结算金额)) + Val(mcllCurSquareBalance(i)(2))
                grsTotal.Update
            Next
        Else
            grsTotal.Find "结算方式='" & mCurCardPay.str结算方式 & "'", , adSearchForward, 1
            If grsTotal.EOF Then grsTotal.AddNew
            ''1-现金结算方式,2-其他非医保结算,3-医保个人帐户,4-医保各类统筹,5-代收款项,6-费用折扣,7-一卡通结算,8-结算卡结算;<0 表示第三方支付
            '.rowdata:  0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
             '性质:99-缴款;98-找补,0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
        
            Select Case mCurCardPay.int性质
            Case 1, 2
                grsTotal!性质 = 0
            Case 3, 4
                grsTotal!性质 = 2
            Case 7, 8
                grsTotal!性质 = IIf(mCurCardPay.blnOneCard, 4, 3)
            Case Else
                grsTotal!性质 = 0
            End Select
            
            grsTotal!结算方式 = mCurCardPay.str结算方式
            grsTotal!结算金额 = Val(Nvl(grsTotal!结算金额)) + dblMoney
            grsTotal.Update
            If dbl退支票 <> 0 Then
                grsTotal.Find "结算方式='" & mstr退支票 & "'", , adSearchForward, 1
                If grsTotal.EOF Then grsTotal.AddNew
                grsTotal!性质 = 0
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
    If mobjChargeInfor.intInsure = 0 Then Exit Sub
    mInsurePara.医保接口打印票据 = gclsInsure.GetCapability(support医保接口打印票据, mobjChargeInfor.病人ID, mobjChargeInfor.intInsure)
    mInsurePara.门诊连续收费 = gclsInsure.GetCapability(support门诊连续收费, mobjChargeInfor.病人ID, mobjChargeInfor.intInsure)
    '刘兴洪:27536 20100119
    mInsurePara.不提醒缴款金额不足 = gclsInsure.GetCapability(support不提醒缴款金额不足, mobjChargeInfor.病人ID, mobjChargeInfor.intInsure)
    mInsurePara.分币处理 = gclsInsure.GetCapability(support分币处理, mobjChargeInfor.病人ID, mobjChargeInfor.intInsure)
    mInsurePara.多单据分单据结算 = gclsInsure.GetCapability(support多单据分单据结算, mobjChargeInfor.病人ID, mobjChargeInfor.intInsure)
    mInsurePara.一次结算分单据退费 = gclsInsure.GetCapability(support一次结算分单据退费, mobjChargeInfor.病人ID, mobjChargeInfor.intInsure)
End Sub
Private Sub InitBalanceData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化结算数据
    '编制:刘兴洪
    '日期:2012-02-05 16:02:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call ClearBanalce
    With mCurCarge
          .dbl本次实收 = mobjChargeInfor.实收金额
          .dbl本次医保支付 = mobjChargeInfor.医保结算金额
          .dbl本次已付合计 = 0
          .dbl本次应收 = mobjChargeInfor.应收金额
          .dbl当前未付 = .dbl本次实收 - .dbl本次医保支付
          .dbl本次冲预交 = 0
          .dbl本次误差费 = 0
      End With
      '保存预结算未付金额，用于与结算结果进行比较，确定是否重复报价
      mdbl原未付 = mCurCarge.dbl当前未付
      mblnYB退款 = mCurCarge.dbl当前未付 + mCurCarge.dbl应缴累计 < 0
      
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
    
    '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
    gstrSQL = "" & _
    "   Select  A.ID, " & _
    "        Case when Mod(A.记录性质,10)=1 then 1  " & _
    "             when B.名称 is not null then  2 " & _
    "             when nvl(A.卡类别ID,0)<>0  then  3 " & _
    "             when J.结算方式 is not null   then  4 " & _
    "             else 0 end as 类型, " & _
    "        Mod(A.记录性质,10) as 记录性质,A.结算方式,A.冲预交,A.摘要, " & _
    "        A.卡类别ID,A.结算卡序号, " & _
    "        A.结算号码,A.卡号,A.交易流水号,nvl(C.是否自制,0) as 自制卡, " & _
    "        nvl(C.是否退现,0) as 是否退现,nvl(C.是否全退,0) as 是否全退, " & _
    "        Decode(C.卡号密文,NULL,0,1) as  是否密文,nvl(C.是否退款验卡,0) as 是否退款验卡," & _
    "        C.名称 as 卡类别名称,A.交易说明,A.结算序号,A.校对标志, " & _
    "        decode(B.名称,Null,0,1) as 医保,0 as 消费卡id" & _
    "   From 病人预交记录 A ,医疗卡类别 C,一卡通目录 J, " & _
    "        (Select 名称 From 结算方式 where 性质 in (3,4)) B" & _
    "   Where A.结帐ID= [1] " & _
    "         And A.结算方式=J.结算方式(+) And A.卡类别ID=C.ID(+) " & _
    "         And A.结算方式=B.名称(+)  " & _
    "         And (a.记录性质 In (1, 11) Or Nvl(a.结算卡序号, 0) = 0)"
       
    gstrSQL = gstrSQL & " Union ALL " & _
    "   Select A.ID,5 as  类型,Mod(A.记录性质,10) as 记录性质,A.结算方式,-1*nvl(b.应收金额,0) as 冲预交,A.摘要,A.卡类别ID,A.结算卡序号," & _
    "        A.结算号码,B.卡号,B.交易流水号,nvl( M.自制卡,0) as 自制卡, " & _
    "        nvl( M.是否退现,0) as 是否退现,nvl(M.是否全退,0) as 是否全退, " & _
    "        nvl(M.是否密文,0) as  是否密文,0 as 是否退款验卡," & _
    "        M.名称 as 卡类别名称,A.交易说明,A.结算序号,A.校对标志,0 as 医保,B.消费卡id" & _
    "   From 病人预交记录 A ,病人卡结算记录 B, 消费卡类别目录 M " & _
    "   Where  a.Id = b.结算id And a.结算卡序号 = m.编号  " & _
    "        And A.结帐ID = [1] and Mod(A.记录性质,10)<>1 "
    
   gstrSQL = "" & _
   "   Select  类型,记录性质,结算方式,摘要,卡类别ID,卡类别名称,自制卡,结算卡序号,结算号码,卡号,交易流水号, 交易说明,结算序号,校对标志,医保,消费卡id," & _
   "         max(是否密文) as 是否密文,max(是否退款验卡) as 是否退款验卡," & _
   "         max(是否全退) as 是否全退,max(是否退现) as 是否退现 , nvl(sum(冲预交),0) as 冲预交" & _
   "   From (" & gstrSQL & ") " & _
   "   Group by 类型,记录性质,结算方式,摘要,卡类别ID,卡类别名称,自制卡,结算卡序号,结算号码,卡号,交易流水号, 交易说明,结算序号,校对标志,医保,消费卡id"
    
    Set mrsBalance = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mobjChargeInfor.结帐ID)
    With mrsBalance
        i = 1: blnYb = False
        Do While Not .EOF
            Select Case Nvl(!类型)
            Case 1 '预交款
                mCurCarge.dbl本次冲预交 = RoundEx(mCurCarge.dbl本次冲预交 + Val(Nvl(!冲预交)), 6)
                mCurCarge.dbl本次已付合计 = RoundEx(mCurCarge.dbl本次已付合计 + Val(Nvl(!冲预交)), 6)
            Case 2, 3, 5 '医保,一卡通,消费卡
                If Nvl(!类型) = 2 Then
                    mCurCarge.dbl本次医保支付 = RoundEx(mCurCarge.dbl本次医保支付 + Nvl(!冲预交, 0), 6)
                    blnYb = True
                End If
                If Val(Nvl(mrsBalance!校对标志, 0)) = 2 Then
                    With vsBlance
                        If .TextMatrix(i, .ColIndex("支付方式")) <> "" Then
                            .Rows = .Rows + 1
                            i = i + 1
                        End If
                        .RowData(i) = Nvl(mrsBalance!类型)
                        strCardNo = Nvl(mrsBalance!卡号)
                        lng卡类别ID = Val(Nvl(mrsBalance!结算卡序号))
                        If Nvl(mrsBalance!类型) = 5 Then
                            If mcllSquareBalance Is Nothing Then Set mcllSquareBalance = New Collection
                            'array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文)
                            mcllSquareBalance.Add Array(lng卡类别ID, Val(Nvl(mrsBalance!消费卡ID)), _
                            Format(Val(Nvl(mrsBalance!冲预交)), "0.00"), strCardNo, "", "", Val(Nvl(mrsBalance!是否密文)))
                        End If
                        .TextMatrix(i, .ColIndex("支付方式")) = Nvl(mrsBalance!结算方式)
                        ' 医疗卡类别ID|消费卡(1, 0) |自制卡|是否全退|是否退现|接口名称
                        .Cell(flexcpData, i, .ColIndex("支付方式")) = lng卡类别ID & "|" & IIf(Val(Nvl(mrsBalance!类型)) = 5, 1, 0) & "|" & Val(Nvl(mrsBalance!自制卡)) & "|" & Val(Nvl(mrsBalance!是否全退)) & "|" & Val(Nvl(mrsBalance!是否退现)) & "|" & Nvl(mrsBalance!卡类别名称)
                        .TextMatrix(i, .ColIndex("支付金额")) = Format(Val(Nvl(mrsBalance!冲预交)), "0.00")
                        .TextMatrix(i, .ColIndex("结算号码")) = Nvl(mrsBalance!结算号码)
                        .TextMatrix(i, .ColIndex("备注")) = Nvl(mrsBalance!摘要)
                        .TextMatrix(i, .ColIndex("交易流水号")) = Nvl(mrsBalance!交易流水号)
                        .TextMatrix(i, .ColIndex("交易说明")) = Nvl(mrsBalance!交易说明)
                        .TextMatrix(i, .ColIndex("卡号")) = IIf(Val(Nvl(mrsBalance!是否密文)) = 1, String(Len(strCardNo), "*"), strCardNo)
                        .Cell(flexcpData, i, .ColIndex("卡号")) = Nvl(mrsBalance!卡号)
                        mCurCarge.dbl本次已付合计 = RoundEx(mCurCarge.dbl本次已付合计 + Val(Nvl(mrsBalance!冲预交)), 6)
                    End With
                End If
            Case Else '0-普通结算
                With vsBlance
                   If .TextMatrix(i, .ColIndex("支付方式")) <> "" And Nvl(mrsBalance!结算方式) <> "" Then
                       .Rows = .Rows + 1
                       i = i + 1
                   End If
                   If Nvl(mrsBalance!结算方式) <> "" Then
                        .RowData(i) = Nvl(mrsBalance!类型)
                        .TextMatrix(i, .ColIndex("支付方式")) = Nvl(mrsBalance!结算方式)
                        .TextMatrix(i, .ColIndex("支付金额")) = Format(Val(Nvl(mrsBalance!冲预交)), "0.00")
                        .TextMatrix(i, .ColIndex("结算号码")) = Nvl(mrsBalance!结算号码)
                        .TextMatrix(i, .ColIndex("备注")) = Nvl(mrsBalance!摘要)
                        .TextMatrix(i, .ColIndex("交易流水号")) = Nvl(mrsBalance!交易流水号)
                        .TextMatrix(i, .ColIndex("交易说明")) = Nvl(mrsBalance!交易说明)
                        .TextMatrix(i, .ColIndex("卡号")) = IIf(Val(Nvl(mrsBalance!是否密文)) = 1, String(Len(strCardNo), "*"), strCardNo)
                        .Cell(flexcpData, i, .ColIndex("卡号")) = Nvl(mrsBalance!卡号)
                        mCurCarge.dbl本次已付合计 = RoundEx(mCurCarge.dbl本次已付合计 + Val(Nvl(mrsBalance!冲预交)), 6)
                    End If
                End With
            End Select
            .MoveNext
        Loop
    End With
                   
    gstrSQL = "" & _
    "   Select  B.NO,B.结帐ID, Nvl(Sum(Nvl(B.应收金额, 0)), 0)  As 本次应收合计, " & _
    "       Nvl(Sum(Nvl(B.实收金额, 0)), 0)  As 本次实收合计 " & _
    "   From 门诊费用记录 B  " & _
    "   Where B.结帐id =[1]  " & _
    "   Group by B.NO,B.结帐ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mobjChargeInfor.结帐ID)
    With mCurCarge
         .dbl本次实收 = 0:
         .dbl本次应收 = 0
        Do While Not rsTemp.EOF
            .dbl本次实收 = RoundEx(.dbl本次实收 + Val(Nvl(rsTemp!本次实收合计)), 6)
            .dbl本次应收 = RoundEx(.dbl本次应收 + Val(Nvl(rsTemp!本次应收合计)), 6)
            rsTemp.MoveNext
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
                ' 0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                .RowData(i) = 1
                .TextMatrix(i, .ColIndex("支付方式")) = "预存款"
                .TextMatrix(i, .ColIndex("支付金额")) = Format(mCurCarge.dbl本次冲预交, "0.00")
            End With
        End If
        mblnYB退款 = mCurCarge.dbl当前未付 + mCurCarge.dbl应缴累计 < 0 And blnYb
    End With
    
    vsBlance_AfterRowColChange 0, 0, vsBlance.Row, vsBlance.Col
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

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
    " Select B.名称 " & _
    " From 结算方式应用 A, 结算方式 B " & _
    " Where A.应用场合 = '收费' And B.名称 = A.结算方式 And a.付款方式 Is Null" & _
    "       And Nvl(B.应付款, 0) = 1 And Rownum <= 1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTemp.EOF Then
        mstr退支票 = Nvl(rsTemp!名称)
    End If
    
    Call initInsure
    If mbytFunc = EM_FUN_收费 Then
        Call InitBalanceData
    Else
        If mobjChargeInfor.intInsure <> 0 Then   '医保结算时,异常单据一般都是结算了的.
            strSQL = "Select 1" & _
                    " From 病人预交记录 A, 结算方式 B" & _
                    " Where a.结算方式 = b.名称 And b.性质 In (3, 4) And 结帐id = [1] " & _
                    "       And Nvl(校对标志, 0) = 1 And Rownum < 2"
            strSQL = strSQL & "Union All" & _
                    " Select 1" & _
                    " From 保险结算记录" & _
                    " Where 记录id = [1] " & _
                    "       And Not Exists(Select 1 From 病人预交记录 A, 结算方式 B" & _
                    "                       Where a.结算方式 = b.名称 And b.性质 In (3, 4) And a.结帐id = 记录id)" & _
                    "       And Rownum < 2"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjChargeInfor.结帐ID)
            '校对标志等于2则已成功结算
            '91914,多单据分单据结算不支持预结算时病人预交记录中有可能没有医保结算信息
            mblnYbBalanced = rsTemp.EOF
        End If
        Call LoadData
    End If
    Call Load支付方式: Call LoadPatiInfor
    Call SetDeleteVisible '进入结算界面时删除按钮应该根据情况显示
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
    Dim blnVisible As Boolean
    
    sngSplitHeight = 80
    
    '51670
    If mobjChargeInfor.病人ID = 0 Or mbln连续输入 Then
        lbl冲预交.Visible = False
        txt冲预交.Visible = False
        txt冲预交.Text = "0"
    End If
    
    blnVisible = mbytFunc = EM_FUN_收费  '功能为正常收费
    ' 0-代表不进行缴款输入和累计控制,1-代表输入缴款后才结束病人累计(改变病人除外)，2-收费时必须要输入缴款金额
    ' 3-收费时,按单病人进行累计(除非按了收费的完成收费功能或改变病人时)
    blnVisible = blnVisible And (gTy_Module_Para.byt缴款控制 = 1 Or gTy_Module_Para.byt缴款控制 = 3)
    blnVisible = blnVisible And Val(txt冲预交.Text) = 0 '未使用预交款
    blnVisible = blnVisible And Val(txt缴款.Text) = 0 '未输入缴款金额
    blnVisible = blnVisible And (mCurCarge.dbl本次已付合计 - mCurCarge.dbl本次医保支付) = 0 '全部为医保支付时
    
    '未使用三方卡结算
    blnVisible = blnVisible And mCurCardPay.lng医疗卡类别ID = 0 And mCurCardPay.blnOneCard = False
    '普通病人或仅只使用了医保结算
    blnVisible = blnVisible And (mobjChargeInfor.intInsure = 0 Or mobjChargeInfor.intInsure <> 0 And mblnYbBalanced)
    cmdNext.Visible = blnVisible
        
    lbl已结.Caption = "已付合计:" & Format(mCurCarge.dbl本次已付合计, "###0.00;-###0.00;0.00;0.00;")
    
    If mCurCardPay.int性质 = 1 And bln预交 = False Then
        dblMoney = mCurCarge.dbl当前未付 + mCurCarge.dbl应缴累计
        If mobjChargeInfor.intInsure > 0 Then  '问题:43855,44069
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
        lbl找补.Caption = "找  补"
    Else
        lblPayType.Caption = "退　款"
        lblPayType.ForeColor = vbRed
        cbo支付方式.ForeColor = vbRed
        txt缴款.ForeColor = vbRed
        lbl找补.Caption = "收  零"
        '退款时，不处理预交
        txt冲预交.Visible = False: lbl冲预交.Visible = False
        mblnNotChange = True
        txt冲预交.Text = "0"
        mblnNotChange = False
    End If
    
    If bln预交 Then
        '预交的处理
        lbl找补.Visible = False: txt找补.Visible = False
        txt找补.Text = 0
    ElseIf mCurCardPay.int性质 = 1 Then
        lbl找补.Visible = True: txt找补.Visible = True
'        lbl找补.Caption = "找　补"
        If IIf(mblnYB退款, -1, 1) * Val(txt缴款.Text) >= dbl现金 Then
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
        txt找补.Text = Format(IIf(mblnYB退款, -1, 1) * Val(txt缴款.Text) - dblMoney - mCurCarge.dbl应缴累计, "0.00")
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
    txt缴款.Text = "": txt缴款.Locked = False
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
       .str交易流水号 = ""
       .str交易说明 = ""
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
     Call Show误差金额(False)
     If Not mcolCardPayMode Is Nothing Then
        If mCurCardPay.bln消费卡 Or (mCurCardPay.int性质 <> 1 And mblnYB退款) Then
            '57682:缺省为所有支付金额
            txt缴款.Text = Format(IIf(mblnYB退款, -1, 1) * Val(lbl剩余自付.Caption), "0.00")
        ElseIf mCurCardPay.lng医疗卡类别ID > 0 And Not mCurCardPay.bln消费卡 Then
            If gTy_Module_Para.byt刷卡缺省金额操作 <> 0 Then
                txt缴款.Text = Format(IIf(mblnYB退款, -1, 1) * Val(lbl剩余自付.Caption), "0.00")
                '金额不允许修改
                If gTy_Module_Para.byt刷卡缺省金额操作 = 2 Then txt缴款.Locked = True
            End If
        End If
     End If
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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据指定支付类别,弹出刷卡窗口
    '入参:rsClassMoney:收费类别,金额
    '        lngCardTypeID-为零时,为老一卡通刷卡
    '       bln余额不足禁止-目前只针对消费卡,表示余额不足时,禁止继续操作,否则用余额进行支付
    dblMoney = Val(txt缴款.Text)
    If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModule, Nothing, 0, False, _
    mobjChargeInfor.姓名, mobjChargeInfor.性别, mobjChargeInfor.年龄, dblMoney, mCurCardPay.str刷卡卡号, mCurCardPay.str刷卡密码, _
    False, True, False, False, Nothing, False, False, "<IN><CZLX>2</CZLX></IN>") = False Then Exit Function
 
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
        Call SetControlProperty(True)
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
    
    If Not CheckTextLength("结算号码", txt结算号码) Then Exit Function
    If Not CheckTextLength("摘要", txt摘要) Then Exit Function
    
    '单独的应缴
    If Not mbln已报价 Then Call LedVoiceSpeak
    txt缴款.Text = Format(Val(txt缴款.Text), "0.00")
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
            If mblnYB退款 Then
                If CSng(txt找补.Text) > 0 Then
                    MsgBox "退款金额不足，请补足退款金额！", vbInformation, gstrSysName
                    txt缴款.SetFocus: zlControl.TxtSelAll txt缴款
                    Exit Function
                End If
            Else
                If CSng(txt找补.Text) < 0 Then
                    MsgBox "缴款金额不足，请补足缴款金额！", vbInformation, gstrSysName
                    txt缴款.SetFocus: zlControl.TxtSelAll txt缴款
                    Exit Function
                End If
            End If
        End If
    Else
        If mblnCur连续 = False Then
            If Val(txt缴款) = 0 Then
                MsgBox "未输入交易金额,请检查!", vbInformation + vbOKOnly, gstrSysName
                If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
                zlControl.TxtSelAll txt缴款: Exit Function
            End If
            If Not IsNumeric(txt缴款.Text) And txt缴款.Text <> "" Then
                MsgBox "无效数值！", vbInformation, gstrSysName
                If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
                zlControl.TxtSelAll txt缴款: Exit Function
            ElseIf Val(txt缴款.Text) < 0 Then
                MsgBox "交易金额不能为负！", vbInformation, gstrSysName
                If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
                zlControl.TxtSelAll txt缴款: Exit Function
            End If
        End If
        If Not mCurCardPay.bln支票 Then
            '问题:42793
            '其他结算方式,输入的金额不能大于未付部分
            If RoundEx(Abs(Val(txt缴款.Text)), 2) > RoundEx(Abs(mCurCarge.dbl当前未付 + mCurCarge.dbl应缴累计), 2) Then
                MsgBox "注意:" & vbCrLf & "    输入的" & IIf(mblnYB退款, "退款", "缴款") & "金额大于了未" & IIf(mblnYB退款, "退", "支付") & "的金额，不能继续！", vbOKOnly + vbInformation, gstrSysName
                txt缴款.SetFocus: zlControl.TxtSelAll txt缴款
                Exit Function
            End If
        End If
        If Val(txt缴款.Text) <> 0 And mCurCarge.dbl应缴累计 <> 0 Then '完成连续收费时，输入金额不能小于未付金额
            If RoundEx(Abs(Val(txt缴款.Text)), 2) < RoundEx(Abs(mCurCarge.dbl当前未付 + mCurCarge.dbl应缴累计), 2) Then
                MsgBox IIf(mblnYB退款, "退款", "缴款") & "金额不足，请补足" & IIf(mblnYB退款, "退款", "缴款") & "金额！", vbInformation, gstrSysName
                txt缴款.SetFocus: zlControl.TxtSelAll txt缴款
                Exit Function
            End If
        End If
    End If

    '检查当前单据是否被其他人执行完成,主要是并发原因进行检查
    '防止其他操作员操作:
    '45186
    gstrSQL = "" & _
    "   Select  1  From 病人预交记录 A " & _
    "   Where   A.结帐ID=[1] and nvl(A.校对标志,0)<>0 and Rownum =1 and A.记录状态=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mobjChargeInfor.结帐ID)
    If rsTemp.EOF Then
        '估计是被他人执行,现在需要检查是否被他人执行
        gstrSQL = "Select 记录状态, 操作员姓名,费用状态 From 门诊费用记录 Where 结帐ID=[1] And rownum=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mobjChargeInfor.结帐ID)
        
        If Not rsTemp.EOF Then
            If Val(Nvl(rsTemp!记录状态)) <> 1 Then
                MsgBox "该单据已经被其他操作员作废,不能再进行收费!", vbOKOnly + vbInformation, gstrSysName
                '执行收费
                Unload Me
                Exit Function
            End If
            
            If Val(Nvl(rsTemp!费用状态)) <> 1 Then
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
    
    lngCount = IIf(mobjChargeInfor.intInsure <> 0, 1, 0)   '医保算一个数量
    If mCurCardPay.lng医疗卡类别ID = 0 Or (mCurCardPay.bln消费卡 And mCurCardPay.bln自制卡) Then CheckInterfaceNumIsValied = True: Exit Function
    With vsBlance
        strNames = IIf(mobjChargeInfor.intInsure <> 0, vbCrLf & "医保结算", "")
        For i = 1 To .Rows - 1
            '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
            If Val(.RowData(i)) = 3 Or Val(.RowData(i)) = 4 Or Val(.RowData(i)) = 5 Then
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

Private Function CheckDelValied(ByRef blnExistThreeSwap As Boolean, _
    ByRef bln全退 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查退费
    '出参:blnExistThreeSwap-是否存在三方接口
    '        bln全退-存在三方接口是否必须全退
    '返回:合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-25 16:14:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double
    
    bln全退 = False: blnExistThreeSwap = False
    On Error GoTo errHandle
    
    mrsBalance.Filter = "  类型=3   OR 类型=4    "
    If mrsBalance.EOF Then mrsBalance.Filter = 0: CheckDelValied = True: Exit Function
    With mrsBalance
        Do While Not .EOF
            dblMoney = RoundEx(Val(Nvl(!冲预交)), 6)
            '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
            Select Case Nvl(!类型)
            Case 3  '一卡通(新)
                If Val(Nvl(!校对标志)) = 2 Then
                    '回退
                    If zlCheckDelValied(Val(Nvl(!卡类别ID)), CStr(Nvl(!卡类别名称)), False, _
                        Nvl(!卡号), Nvl(!交易流水号), Nvl(!交易说明), mobjChargeInfor.结帐ID, dblMoney, _
                        Val(Nvl(!是否退款验卡)) = 1) = False Then Exit Function
                    If Not bln全退 Then bln全退 = Val(Nvl(!是否全退)) = 1
                    
                End If
            Case 4  '一卡通(老)
                If Val(Nvl(!校对标志)) = 2 Then
                    If CheckDelOneCardValied(Nvl(!卡号), dblMoney) = False Then Exit Function
                    If Not bln全退 Then bln全退 = True
                End If
            End Select
            .MoveNext
        Loop
    End With
    blnExistThreeSwap = True

    CheckDelValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
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

Private Sub cbo支付方式_GotFocus()
    Clear预交款
End Sub

Private Sub cbo支付方式_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Function CancelBalance(ByRef blnUnload As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:异常作废
    '出参:blnUnload-是否执行unload me
    '编制:刘兴洪
    '日期:2014-06-19 14:42:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dtDelDate As Date, lng领用ID As Long
    Dim strInvoice As String, strNo As String, strSQL As String
    Dim cllPro As Collection, varData As Variant, i As Long
    Dim blnIsExiseThreeSwap As Boolean, bln全退 As Boolean
    Dim blnCommit As Boolean
    
    dtDelDate = zlDatabase.Currentdate
    blnUnload = False
    '一卡通;第三方交易的检查
    If CheckDelValied(blnIsExiseThreeSwap, bln全退) = False Then
        If MsgBox("注意:" & vbCrLf & "不能正常的进行第三方交易退费,是否暂停交易?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        blnUnload = True: Exit Function
    End If
    
    If mobjChargeInfor.intInsure <> 0 And mInsurePara.医保接口打印票据 Then
        If zlCheckInvoiceValied(lng领用ID, 1, , mobjChargeInfor.ShareUserID, mobjChargeInfor.PatiUseType) = False Then
            If MsgBox("注意:" & vbCrLf & "    无有效票据,是否暂停交易?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            blnUnload = True: Exit Function
        End If
        strInvoice = GetNextBill(lng领用ID)
    End If
    
    mobjChargeInfor.收费结算 = "": mbln连续输入 = False
    '单据作废处理
    Set cllPro = New Collection
    If mobjChargeInfor.Nos = "" And mobjChargeInfor.结帐ID <> 0 Then
        mobjChargeInfor.Nos = zlGetBalanceNos(1, mobjChargeInfor.结帐ID, False)
     End If
    
    varData = Split(Replace(mobjChargeInfor.Nos, "'", ""), ",")
    If Not mbln异常作废 And Not mblnCommitBill Then
        mobjChargeInfor.冲销ID = zlDatabase.GetNextId("病人结帐记录")
        
        For i = UBound(varData) To 0 Step -1
            strNo = varData(i)
            'Zl_门诊收费记录_销帐
            strSQL = "Zl_门诊收费记录_销帐("
            '  No_In         门诊费用记录.No%Type,
            strSQL = strSQL & "'" & varData(i) & "',"
            '  操作员编号_In 门诊费用记录.操作员编号%Type,
            strSQL = strSQL & "'" & UserInfo.编号 & "',"
            '  操作员姓名_In 门诊费用记录.操作员姓名%Type,
            strSQL = strSQL & "'" & UserInfo.姓名 & "',"
            '  序号_In       Varchar2 := Null,
            strSQL = strSQL & "NULL,"
            '  退费时间_In   门诊费用记录.登记时间%Type := Null,
            strSQL = strSQL & "To_Date('" & Format(dtDelDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
            '  退费摘要_In   门诊费用记录.摘要%Type := Null,
            strSQL = strSQL & "'结算作废',"
            '  结帐id_In     病人预交记录.结帐id%Type := Null,
            strSQL = strSQL & mobjChargeInfor.冲销ID & ","
            '  回收票据_In Number:=0
            strSQL = strSQL & 1 & ")"
            zlAddArray cllPro, strSQL
        Next
        '先产生票据，医保接口才能取到
        If mInsurePara.医保接口打印票据 And mobjChargeInfor.intInsure <> 0 Then
            strSQL = "zl_门诊收费记录_RePrint('" & strNo & "','" & strInvoice & "'," & ZVal(lng领用ID) & ",'" & UserInfo.姓名 & "'," & _
                "To_Date('" & Format(dtDelDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),1,1)"
            zlAddArray cllPro, strSQL
        End If
        
        '原样退
        'Zl_门诊退费结算_Modify
        strSQL = "Zl_门诊退费结算_Modify("
        '  操作类型_In   Number,
        strSQL = strSQL & "" & 0 & ","
        '  病人id_In     门诊费用记录.病人id%Type,
        strSQL = strSQL & "" & mobjChargeInfor.病人ID & ","
        '  冲销id_In     病人预交记录.结帐id%Type,
        strSQL = strSQL & "" & mobjChargeInfor.冲销ID & ","
        '  结算方式_In   Varchar2,
        strSQL = strSQL & "NULL)"
        zlAddArray cllPro, strSQL
 
    End If
    
    Err = 0: On Error GoTo Errhand:
    
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
     
    If mobjChargeInfor.intInsure <> 0 Then
        If ExcuteInsureDel(blnCommit) = False Then
            If blnCommit Then
                mblnCommitBill = True
                cmdExit.Visible = False
            End If
            Exit Function
        End If
        '修改校对标志
        ' Zl_病人门诊收费_医保更新
        strSQL = "Zl_病人门诊收费_医保更新("
        '  结帐id_In   门诊费用记录.结帐id%Type,
        strSQL = strSQL & mobjChargeInfor.冲销ID & ","
        '  结算序号_In 病人预交记录.结算序号%Type,
        strSQL = strSQL & "Null,"
        '  保险结算_In Varchar2
        strSQL = strSQL & "Null)"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        mblnCommitBill = True
        gcnOracle.CommitTrans: gcnOracle.BeginTrans
    End If
    
    On Error GoTo ErrInterface:
    
    '调用三方回退交易

    If ExcuteThreeSwapDel(mobjChargeInfor.冲销ID, mobjChargeInfor.结帐ID) = False Then
        If MsgBox("注意:" & vbCrLf & "不能正常的进行第三方交易退费,是否暂停交易?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        blnUnload = True: Exit Function
    End If
    mblnCommitBill = True
    gcnOracle.CommitTrans: gcnOracle.BeginTrans
      
    If ExcuteOverFeeDel(mobjChargeInfor.冲销ID, mobjChargeInfor.病人ID) = False Then Exit Function
    gcnOracle.CommitTrans
    
    mbytReturnMode = EM_EX_作废
    CancelBalance = True
    blnUnload = True: Exit Function
Errhand:
    gcnOracle.RollbackTrans
ErrInterface:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ExcuteOverFeeDel(ByVal lng冲销ID As Long, ByVal lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:完成退费收费
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-29 14:50:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errHandle
      
    'Zl_门诊退费结算_Modify
    strSQL = "Zl_门诊退费结算_Modify("
    '  操作类型_In   Number,
    strSQL = strSQL & "" & 1 & ","
    '  病人id_In     门诊费用记录.病人id%Type,
    strSQL = strSQL & "" & lng病人ID & ","
    '  冲销id_In     病人预交记录.结帐id%Type,
    strSQL = strSQL & "" & lng冲销ID & ","
    '  结算方式_In   Varchar2,
    strSQL = strSQL & "NULL,"
    '  退预交_In     病人预交记录.冲预交%Type := Null,
    strSQL = strSQL & "NULL,"
    '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
    strSQL = strSQL & "NULL,"
    '  卡号_In       病人预交记录.卡号%Type := Null,
    strSQL = strSQL & "NULL,"
    '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
    strSQL = strSQL & "NULL,"
    '  交易说明_In   病人预交记录.交易说明%Type := Null,
    strSQL = strSQL & "NULL,"
    '  缴款_In       病人预交记录.缴款%Type := Null,
    strSQL = strSQL & "NULL,"
    '  找补_In       病人预交记录.找补%Type := Null,
    strSQL = strSQL & "NULL,"
    '  误差金额_In   门诊费用记录.实收金额%Type := Null,
    strSQL = strSQL & "NULL,"
    '  完成退费_In   Number := 0,
    '  -- 完成退费_In:0-未完成退费;1-异常完成退费;2-完成退费
    strSQL = strSQL & "1)"
    '  原结帐id_In   病人预交记录.结帐id%Type := Null
    '异常单据,冲销也应该为异常单据
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    ExcuteOverFeeDel = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Function

Private Function ExcuteThreeSwapDel(ByVal lng冲销ID As Long, ByVal lng原结帐ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:退费交易(一卡通或三方结算交易)
    '返回:交易成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-25 17:29:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim cllUpdate As Collection, cllThreeSwap As Collection
    Dim strCardNo As String, i As Long, strSQL As String, strErrMsg As String
    Dim strSwapNO As String, strSwapMemo As String, varData As Variant
    Dim lng卡类别ID As Long, bln消费卡 As Boolean, strTemp As String
    Dim st卡类别名称 As String, blnTrans As Boolean, dblMoney As Double
    Dim str医院编码 As String, rsTemp As ADODB.Recordset
    
    gstrSQL = "" & _
    "   Select A.结算方式,A.摘要, " & _
    "             A.卡类别ID,A.结算号码,A.卡号,A.交易流水号, " & _
    "             nvl(C.是否自制,0) as 自制卡, " & _
    "             C.名称 as 名称,A.交易说明," & _
    "             Sum(A.冲预交) as 冲预交" & _
    "   From 病人预交记录 A ,医疗卡类别 C" & _
    "   Where A.结帐ID=[1] And nvl(A.校对标志,0)=1  " & _
    "         And A.卡类别ID=C.ID" & _
    "   Group by A.结算方式,A.摘要,A.卡类别ID , A.结算号码,A.卡号,A.交易流水号, " & _
    "           nvl(C.是否自制,0),C.名称,A.交易说明,A.结算序号"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng冲销ID)
    Set cllUpdate = New Collection: Set cllThreeSwap = New Collection
    With rsTemp
        Do While Not .EOF
                lng卡类别ID = Val(Nvl(!卡类别ID))
                bln消费卡 = False
                st卡类别名称 = Nvl(!名称)
                strSwapNO = Nvl(!交易流水号)
                strSwapMemo = Nvl(!交易说明)
                strCardNo = Nvl(!卡号)
                dblMoney = Nvl(!冲预交)
                
                'Zl_病人预交记录_更新校对标志
                strSQL = "Zl_病人预交记录_更新校对标志("
                '  结帐id_In     门诊费用记录.结帐id%Type,
                strSQL = strSQL & "" & lng冲销ID & ","
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
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
                If CallBackBalanceInterface(lng冲销ID, lng原结帐ID, lng卡类别ID, bln消费卡, dblMoney, strCardNo, strSwapNO, strSwapMemo, cllUpdate, cllThreeSwap, strErrMsg) = False Then
                    gcnOracle.RollbackTrans: Exit Function
                End If
                gcnOracle.CommitTrans
                zlExecuteProcedureArrAy cllUpdate, Me.Caption
                zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
                gcnOracle.BeginTrans
            .MoveNext
        Loop
    End With
    ExcuteThreeSwapDel = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog:
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

Private Function ExcuteInsureDel(ByRef blnCommit As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用医保退费接口
    '返回:调用成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-25 12:17:56
    '说明:需要在外层启用事务,正常退费后,该过程不提交,需要调用者提交;
    '     如果失败,则事务将回退(主要是避免弹出界面造成死锁)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strAdvance As String, blnTransMedicare As Boolean
    Dim i As Long, p As Integer, strAdvanceOld As String
    Dim colBalance As Collection '记录各张单据保险结算
    Dim strSQL As String
    Dim rsCharge As ADODB.Recordset, strNo As String
    Dim str结算方式 As String, strDel结算方式 As String
    On Error GoTo errHandle
    
    strAdvance = mobjChargeInfor.冲销ID & "|" & "0"
    
    blnTransMedicare = False
    If Not (mInsurePara.多单据分单据结算 Or mInsurePara.一次结算分单据退费) Then
        If mblnCommitBill Then ExcuteInsureDel = True: Exit Function
        If Not gclsInsure.ClinicDelSwap(mobjChargeInfor.结帐ID, , mobjChargeInfor.intInsure, strAdvance) Then
             gcnOracle.RollbackTrans: Exit Function
        End If
        
        blnTransMedicare = True
        If strAdvance = mobjChargeInfor.冲销ID & "|" & "0" Or strAdvance = "" Then
            Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, mobjChargeInfor.intInsure)
            ExcuteInsureDel = True
            Exit Function
        End If
    Else
        Set colBalance = New Collection
        strAdvanceOld = strAdvance
        
        '93337,退费时按单据号倒序进行接口调用
        strSQL = "Select Distinct NO From 门诊费用记录 Where 结帐id = [1] Order By No Desc"
        Set rsCharge = zlDatabase.OpenSQLRecord(strSQL, "获取原始费用单据号", mobjChargeInfor.冲销ID)
        
        p = 1
        Do While Not rsCharge.EOF
            colBalance.Add Array()
            strDel结算方式 = "": str结算方式 = ""
            strNo = Nvl(rsCharge!NO)
            '1.检查该张单据是否需要作废医保结算
            str结算方式 = zlGetYBBalanceNo(mobjChargeInfor.结帐ID, strNo, mobjChargeInfor.病人ID, _
                                    mobjChargeInfor.intInsure, True)
            
            '2.检查该单据是否已医保作废（可能存在多次医保结算作废）
            '如果调用成功过接口，但没有任何医保作废，则会再一次调用医保接口，因为无法确定是否调用成功过
            strDel结算方式 = zlGetYBBalanceNo(mobjChargeInfor.冲销ID, strNo)
            Call SetBalanceVal(colBalance, p, strDel结算方式)
                
            '3.调用医保退费接口，提交数据
            If str结算方式 <> "" And strDel结算方式 = "" Then
                '    Zl_医保结算明细_Insert(
                strSQL = "Zl_医保结算明细_Insert("
                '      结帐id_In   医保结算明细.结帐id%Type,
                strSQL = strSQL & "" & mobjChargeInfor.冲销ID & ","
                '      No_In       医保结算明细.No%Type,
                strSQL = strSQL & "'" & strNo & "',"
                '      结算方式_In Varchar2,
                strSQL = strSQL & "'" & str结算方式 & "')"
                '      备注_In     医保结算明细.备注%Type := Null
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
                
                strAdvance = strAdvanceOld & "|" & strNo
                '因为参数固定为医保基金,所以名称固定为医保基金(多种统筹不好确定),以后应去掉该参数
                If Not gclsInsure.ClinicDelSwap(mobjChargeInfor.结帐ID, True, mobjChargeInfor.intInsure, _
                                                strAdvance) Then gcnOracle.RollbackTrans: Exit Function
                If strAdvance = strAdvanceOld & "|" & strNo Then strAdvance = ""
                
                If zlInsureCheck(str结算方式, strAdvance) Then
                    str结算方式 = strAdvance
                    '    Zl_医保结算明细_Insert(
                    strSQL = "Zl_医保结算明细_Insert("
                    '      结帐id_In   医保结算明细.结帐id%Type,
                    strSQL = strSQL & "" & mobjChargeInfor.冲销ID & ","
                    '      No_In       医保结算明细.No%Type,
                    strSQL = strSQL & "'" & strNo & "',"
                    '      结算方式_In Varchar2,
                    strSQL = strSQL & "'" & strAdvance & "')"
                    '      备注_In     医保结算明细.备注%Type := Null
                    zlDatabase.ExecuteProcedure strSQL, Me.Caption
                End If
                gcnOracle.CommitTrans: blnCommit = True
                
                Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, mobjChargeInfor.intInsure)
                Call SetBalanceVal(colBalance, p, str结算方式)
                
                gcnOracle.BeginTrans
            End If
            
            p = p + 1
            rsCharge.MoveNext
        Loop
        
        '全部成功，返回总的结算方式
        strAdvance = GetMedicareStr(colBalance)
    End If
    
    '根据返回的结算信息，修正预交记录，strAdvance返回格式:结算方式1|金额||结算方式2|金额...
    If InStr(strAdvance, "|") > 0 Then
        Call 医保数据更正(mobjChargeInfor.病人ID, mobjChargeInfor.冲销ID, strAdvance, True, Nothing)
    End If
    If Not (mInsurePara.多单据分单据结算 Or mInsurePara.一次结算分单据退费) Then
        Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, mobjChargeInfor.intInsure)
    End If
    ExcuteInsureDel = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If blnTransMedicare Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, False, mobjChargeInfor.intInsure)
    Call ErrCenter
End Function

Private Sub cmdDel_Click()
    Dim dblMoney As Double, strSQL As String
    Dim byt操作类型 As Byte
    Dim str结算方式 As String
    
    Clear预交款
    If mbytFunc = EM_FUN_作废 Then Exit Sub
    '删除相关的费用
    With vsBlance
        If .Row < 0 Then Exit Sub
        '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
        Select Case Val(.RowData(.Row))
        Case 1  '预存款
            byt操作类型 = 1
            str结算方式 = ""
        Case 0  '普通的结算方式
            byt操作类型 = 0
            str结算方式 = .TextMatrix(.Row, .ColIndex("支付方式"))
        Case Else
            ' 2-医保,3-一卡通;4-一卡通(老);5-消费卡
             '不能直接删除
            Exit Sub
        End Select
        dblMoney = Val(.TextMatrix(.Row, .ColIndex("支付金额")))
        
        mCurCarge.dbl当前未付 = RoundEx(mCurCarge.dbl当前未付 + dblMoney, 6)
        mCurCarge.dbl本次已付合计 = RoundEx(mCurCarge.dbl本次已付合计 - dblMoney, 6)
        Call SetControlProperty
        If Val(.RowData(.Row)) = 1 Then
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
    Call SetDeleteVisible
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdExit_Click()
    Clear预交款
    mblnOK = False: mbytReturnMode = EM_EX_退出
    Call ExcuteMainReshData(EM_EX_退出)
    Unload Me
End Sub

Private Sub cmdNext_Click()
    Dim blnUnload As Boolean
    '继续下一张单据的录入
    '保存上次支付方式
    Clear预交款
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
    If SaveCharge(, blnUnload) = False Then
        If mblnPriceBillCommit And mblnCommitBill = False Then
            '直接收费时,删除前一个事务提交的划价单
            Call DelMedicareTempNOs
            mblnPriceBillCommit = False
        End If
        GoTo GoOver
    End If
    
    mbln连续输入 = True
    '刷新主界面
    ExcuteMainReshData EM_EX_继续
    mbytReturnMode = EM_EX_继续
    Unload Me
GoOver:
    mobjChargeInfor.收费结算 = ""
    mblnCur连续 = False
End Sub

Private Sub cmdOK_Click()
    Dim blnUnload As Boolean
    
    If mbytFunc = EM_FUN_重收 Or mbytFunc = EM_FUN_作废 Then
        '并发检查
        If zlIsCheckExistErrBill(mobjChargeInfor.结算序号) = False Then
            MsgBox "当前异常单据已被处理，你不能继续！", vbInformation, gstrSysName
            Exit Sub
        End If
        If zlCheckOtherSessionDoing(mobjChargeInfor.结算序号) Then
            MsgBox "当前单据正在其它收费窗口中进行处理，你不能继续！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If zlIsCheckExiseSingularity(mobjChargeInfor.结算序号) Then
            If mbytFunc = EM_FUN_重收 Then
                MsgBox "该异常单据已经被作废，因此，不能再" & IIf(mbytFunc = EM_FUN_重收, "重新收费", "进行作废") & "，请刷新费用列表！", vbInformation, gstrSysName
                Call cmdExit_Click: Exit Sub
            End If
        End If
        If Not zlIsCheckExistErrBill(mobjChargeInfor.结算序号) Then
            MsgBox "该异常单据已经被重新收费，因此，不能再" & IIf(mbytFunc = EM_FUN_重收, "重新收费", "进行作废") & "，请刷新费用列表！", vbInformation, gstrSysName
            Call cmdExit_Click: Exit Sub
        End If
    End If
    If mbytFunc = EM_FUN_作废 Then
        mblnOK = CancelBalance(blnUnload)
        If blnUnload = True Then
            ExcuteMainReshData (EM_EX_作废)
            Unload Me
        End If
        Exit Sub
    End If
   '单据界面按了回车符
   If mblnCacheKeyReturn Then mblnCacheKeyReturn = False: Exit Sub
    '先处理预交
    mbln连续输入 = False
    mblnCurBrushPrepay = False
    If BrushcardStrikePrepay = False Then
       If txt冲预交.Enabled And txt冲预交.Visible Then txt冲预交.SetFocus
        Exit Sub
    End If
    If mblnCurBrushPrepay Then
        If mblnUnloaded Then
            '刷新主界面信息
            ExcuteMainReshData EM_EX_完成
            Unload Me
        End If
        Exit Sub
    End If
    
    '再处理其他
    If isValied = False Then Exit Sub
    If txt缴款.Text <> "0.00" Then
        'LED显示
        Call ShowLedInfor
    End If
    If SaveCharge(, blnUnload) = False Then
        If mblnPriceBillCommit And mblnCommitBill = False Then
            '直接收费时,删除前一个事务提交的划价单
            Call DelMedicareTempNOs
            mblnPriceBillCommit = False
        End If
        Exit Sub
    End If
    If blnUnload Then
        '刷新主界面信息
        ExcuteMainReshData EM_EX_完成
        Unload Me
    End If
End Sub

Private Sub ExcuteMainReshData(ByVal bytExitMode As gExitMode)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行主界面的刷新数据
    '编制:刘兴洪
    '日期:2014-06-17 15:09:44
    '说明:主要是应用医保刷新
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mblnOK Then zlAutoPayDrugAndStuff mcllOverPro  '自动发料
    If Not gfrmMain Is Nothing Then Exit Sub
    Call mfrmMain.zlExeBalanceWinRefrshData(mblnOK, bytExitMode, mbln连续收费, mobjChargeInfor)
End Sub

Private Function zlAutoPayDrugAndStuff(ByRef cllDrugAndStuff As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:进行自动发料
    '返回:发料成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-02-06 14:55:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllItem As Collection, blnTrans As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strWhere As String
    Dim strDrugSql As String, strStuffSql As String
    
    On Error GoTo errHandle
    If mbytFunc = EM_FUN_重收 And (gbln收费后自动发药 Or gbln门诊自动发料) Then '104017
        '异常重收时重新自动发药/发料
        Set cllDrugAndStuff = New Collection
        If gbln收费后自动发药 Then
            strDrugSql = _
                " Select Distinct 1 As 类型, a.No, a.执行部门id, a.开单人" & vbNewLine & _
                " From 门诊费用记录 A" & vbNewLine & _
                " Where a.记录性质 = 1 And Nvl(a.执行部门id, 0) <> 0" & vbNewLine & _
                "       And a.收费类别 In ('5', '6', '7') And a.结帐ID = [1]"
        End If
        
        If gbln门诊自动发料 Then
            strStuffSql = _
                " Select Distinct 2 As 类型, a.No, a.执行部门id, a.开单人" & vbNewLine & _
                " From 门诊费用记录 A, 材料特性 B" & vbNewLine & _
                " Where a.记录性质 = 1 And a.收费细目id = b.材料id(+) And Nvl(a.执行部门id, 0) <> 0" & vbNewLine & _
                "       And a.收费类别 = '4' And b.跟踪在用 = 1 And a.结帐ID = [1]"
        End If
        strSQL = strDrugSql & vbNewLine & _
            IIf(gbln门诊自动发料, " Union All" & vbNewLine & strStuffSql, "")
        
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询药品和卫材费用", mobjChargeInfor.结帐ID)
        Do While Not rsTemp.EOF
            If Val(Nvl(rsTemp!类型)) = 1 Then '药品
                strSQL = "ZL_药品收发记录_处方发药(" & Val(Nvl(rsTemp!执行部门ID)) & ",8,'" & Nvl(rsTemp!NO) & "','" & UserInfo.姓名 & "','" & UserInfo.姓名 & "','" & Nvl(rsTemp!开单人) & "')"
            Else '卫材
                '24-收费处方发料；25-记帐单处方发料
                strSQL = "zl_材料收发记录_处方发料(" & Val(Nvl(rsTemp!执行部门ID)) & ",24,'" & Nvl(rsTemp!NO) & "','" & UserInfo.姓名 & "','" & UserInfo.姓名 & "','" & UserInfo.姓名 & "',1,Sysdate)"
            End If
            zlAddArray cllDrugAndStuff, strSQL
            rsTemp.MoveNext
        Loop
        
        If cllDrugAndStuff.Count = 0 Then zlAutoPayDrugAndStuff = True: Exit Function
        blnTrans = True
        ExecuteProcedureArrAy cllDrugAndStuff, Me.Caption
        blnTrans = False
        zlAutoPayDrugAndStuff = True
        Exit Function
    End If
    
    If cllDrugAndStuff Is Nothing Then zlAutoPayDrugAndStuff = True: Exit Function
    
    blnTrans = True
    gcnOracle.BeginTrans
    For Each cllItem In cllDrugAndStuff '91665
        ExecuteProcedureArrAy cllItem, Me.Caption, True, True
    Next
    gcnOracle.CommitTrans
    blnTrans = False
    zlAutoPayDrugAndStuff = True
    Exit Function
errHandle:
    If blnTrans Then
        gcnOracle.RollbackTrans
        Call ErrCenter '不能重试
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
    End If
End Function

Private Sub SetCtrlVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件的显示状态
    '编制:刘兴洪
    '日期:2012-02-03 13:58:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnTemp As Boolean
    If mbytFunc = EM_FUN_收费 Then
        '医保且医保未进行结算时,才显示
        cmdYBBalance.Visible = mobjChargeInfor.intInsure <> 0 And Not mblnYbBalanced
        '医保进行结算了的,或非医保的,显示完成收费
        cmdOK.Visible = (mobjChargeInfor.intInsure = 0 Or mobjChargeInfor.intInsure <> 0 And mblnYbBalanced)
        '医保进行了结算后,不能退出
        cmdExit.Visible = mobjChargeInfor.intInsure = 0 And Not (mblnThreeInterface Or mblnCommitBill) _
                          Or mobjChargeInfor.intInsure <> 0 And Not mblnYbBalanced
        '连续收费
        blnTemp = gTy_Module_Para.byt缴款控制 = 1 Or gTy_Module_Para.byt缴款控制 = 3 '是否具体连续收费
        '普通收费或医保已经结算
        blnTemp = blnTemp And (mobjChargeInfor.intInsure = 0 Or mobjChargeInfor.intInsure <> 0 And mblnYbBalanced)
        blnTemp = blnTemp And Val(txt冲预交.Text) = 0 '未用预交款的
        cmdNext.Visible = blnTemp And (mCurCarge.dbl本次实收 = mCurCarge.dbl当前未付)
        If (gTy_Module_Para.byt缴款控制 = 1 Or gTy_Module_Para.byt缴款控制 = 3) And mbln连续输入 Then
            cbo支付方式.Locked = True
        End If
        Exit Sub
     End If
     
     If mbytFunc = EM_FUN_重收 Then
        cmdExit.Caption = "退出(&E)"
        cmdOK.Visible = mblnYbBalanced Or mobjChargeInfor.intInsure = 0
        cmdYBBalance.Visible = Not mblnYbBalanced And mobjChargeInfor.intInsure <> 0
        cmdExit.Visible = True
        cmdNext.Visible = False
     End If
     If mbytFunc = EM_FUN_作废 Then
        cmdOK.Caption = "作废结算(&O)"
        cmdExit.Caption = "退出(&E)"
        cmdOK.Visible = True
        cmdYBBalance.Visible = False
        cmdExit.Visible = True
        cmdNext.Visible = False
     End If
End Sub

Private Function 医保数据更正(ByVal lng病人ID As Long, ByVal lng结帐ID As Long, _
    ByVal str医保结算 As String, ByVal bln作废 As Boolean, _
    ByRef cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医保数据校对更正
    '返回:校对成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-06-12 17:45:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    If bln作废 Then
        'Zl_门诊退费结算_Modify
        strSQL = "Zl_门诊退费结算_Modify("
        '  操作类型_In   Number,
        strSQL = strSQL & "" & 3 & ","
        '  病人id_In     门诊费用记录.病人id%Type,
        strSQL = strSQL & "" & lng病人ID & ","
        '  冲销id_In     病人预交记录.结帐id%Type,
        strSQL = strSQL & "" & lng结帐ID & ","
        '  结算方式_In   Varchar2,
        strSQL = strSQL & "'" & str医保结算 & "')"
        '  退预交_In     病人预交记录.冲预交%Type := Null,
        '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
        '  卡号_In       病人预交记录.卡号%Type := Null,
        '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
        '  交易说明_In   病人预交记录.交易说明%Type := Null,
        '  缴款_In       病人预交记录.缴款%Type := Null,
        '  找补_In       病人预交记录.找补%Type := Null,
        '  误差金额_In   门诊费用记录.实收金额%Type := Null,
        '  完成退费_In Number:=0
        ') As
        '  ------------------------------------------------------------------------------------------------------------------------------
        '  --功能:收费结算时,修改结算的相关信息
        '  --操作类型_In:
        '  --   1-普通退费方式:
        '  --     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
        '  --     ②冲预交_In:如果涉及预交款,则传入本次的退预交,非正常收费时,传入零(<0 表示退预交款;>0 表示将剩余款生成预交记录
        '  --   2.三方卡退费结算:
        '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
        '  --     ②退预交_In: 传入零
        '  --     ③卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
        '  --   3-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
        '  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
        '  --     ②退预交_In: 传入零
        '  --     ③退支票额_In:传入零
        '  --   4-消费卡结算:
        '  --     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."
        '  --     ②退预交_In: 传入零
        '  --     ③退支票额_In:传入零
        '  -- 误差金额_In:存在误差费时,传入
        '  -- 完成退费_In:0-未完成退费;1-异常完成退费;2-完成退费
        '  ------------------------------------------------------------------------------------------------------------------------------
     Else
        '需要修正结算数据
        'Zl_门诊收费结算_Modify
        strSQL = "Zl_门诊收费结算_Modify("
        '  操作类型_In   Number,
        strSQL = strSQL & "" & 2 & ","
        '  病人id_In     门诊费用记录.病人id%Type,
        strSQL = strSQL & "" & lng病人ID & ","
        '  结帐id_In     病人预交记录.结帐id%Type,
        strSQL = strSQL & "" & lng结帐ID & ","
        '  结算方式_In   Varchar2,
        strSQL = strSQL & "'" & str医保结算 & "')"
        '  冲预交_In     病人预交记录.冲预交%Type,
        '  退支票额_In   病人预交记录.冲预交%Type,
        '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
        '  卡号_In       病人预交记录.卡号%Type := Null,
        '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
        '  交易说明_In   病人预交记录.交易说明%Type := Null,
        '  缴款_In       病人预交记录.缴款%Type := Null,
        '  找补_In       病人预交记录.找补%Type := Null,
        '  误差金额_In   门诊费用记录.实收金额%Type := Null,
        '  完成结算_In Number:=0
        ') As
        '  ------------------------------------------------------------------------------------------------------------------------------
        '  --功能:收费结算时,修改结算的相关信息
        '  --操作类型_In:
        '  --   0-普通收费方式:
        '  --     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
        '  --     ②冲预交_In:如果涉及预交款,则传入本次的冲预交,非正常收费时,传入零
        '  --     ③退支票额_In:如果涉及退支票,则传入本次的退支票额,非正常收费时,传入零
        '  --   1.三方卡结算:
        '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
        '  --     ②冲预交_In: 传入零
        '  --     ③退支票额_In:传入零
        '  --     ④卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
        '  --   2-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
        '  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
        '  --     ②冲预交_In: 传入零
        '  --     ③退支票额_In:传入零
        '  --   3-消费卡结算:
        '  --     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."
        '  --     ②冲预交_In: 传入零
        '  --     ③退支票额_In:传入零
        '  -- 误差金额_In:存在误差费时,传入
        '  -- 完成结算_In:1-完成收费;0-未完成收费
        '  ------------------------------------------------------------------------------------------------------------------------------
    End If
    If cllPro Is Nothing Then
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Else
        zlAddArray cllPro, strSQL
    End If
    
    医保数据更正 = True
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
     Dim strSQL As String, lng结帐ID As Long
    Dim cllBalance As Collection
    
    On Error GoTo errHandle
    lng结帐ID = IIf(mbln异常作废, mobjChargeInfor.冲销ID, mobjChargeInfor.结帐ID)
    If lng结帐ID = 0 Then 医保结算较对 = True: Exit Function
    If mobjChargeInfor.intInsure = 0 Then 医保结算较对 = True: Exit Function
    
    '108630,不再根据"保险结算记录.校正"来判断，只要是异常单据都要校对
'    '0-正常;1-待校对;2-完成校对;3-附加，指返回的其它非医保支付的各种结算方式
'    gstrSQL = "" & _
'    "   Select /*+ rule */ A.记录ID,A.校正  " & _
'    "   From 保险结算记录 A" & _
'    "   Where A.记录ID=[1] And nvl(A.校正,0)=1 "
'    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng结帐ID)
'    If rsTemp.EOF Then 医保结算较对 = True: Exit Function
    
    '先通过“医保结算明细”进行校对
    strSQL = "Select 1" & _
            " From 病人预交记录 A, 结算方式 B" & _
            " Where a.结算方式 = b.名称 And b.性质 In (3, 4) And 结帐id = [1] " & _
            "       And Nvl(校对标志, 0) = 1 And Rownum < 2"
    strSQL = strSQL & "Union All" & _
            " Select 1" & _
            " From 保险结算记录" & _
            " Where 记录id = [1] " & _
            "       And Not Exists(Select 1 From 病人预交记录 A, 结算方式 B" & _
            "                       Where a.结算方式 = b.名称 And b.性质 In (3, 4) And a.结帐id = 记录id)" & _
            "       And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng结帐ID)
    If rsTemp.EOF Then 医保结算较对 = True: Exit Function
    
    str医保结算 = zlGetYBBalanceNo(lng结帐ID)
    
    '检查医保核对表，无记录则退出
    'Select 结帐ID,结算方式,金额 From 保险结算明细 Where 标志=1
    strSQL = "Select A.结帐ID,a.结算方式,a.金额" & _
            " From 保险结算明细 A ,结算方式 C" & _
            " Where A.结帐id =[1] and A.标志=1 and A.结算方式=C.名称 And C.性质 in (3,4) " & _
            " Order by A.结算方式"
    '医保管控的过程固定写入了一条"现金",所以排开非医保类的结算方式
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "保险结算管理", lng结帐ID)
    '未有核对数据,直接返回
    If rsTemp.RecordCount = 0 And str医保结算 = "" Then 医保结算较对 = True: Exit Function
    
    If rsTemp.RecordCount > 0 Then
        str医保结算 = "" '结算方式|结算金额||
        Set cllBalance = New Collection
        For i = 1 To rsTemp.RecordCount
            str医保结算 = str医保结算 & "||" & Nvl(rsTemp!结算方式) & "|" & Val(Nvl(rsTemp!金额))
            rsTemp.MoveNext
        Next
        If str医保结算 <> "" Then str医保结算 = Mid(str医保结算, 3)
    End If
    If str医保结算 = "" Then 医保结算较对 = True: Exit Function
    
    strShowMsg = Replace(Replace(str医保结算, "||", vbCrLf), "|", "：")
    MsgBox "注意：" & vbCrLf & "    医保" & IIf(mbln异常作废, "退费", "") & _
        "结算结果与已保存结算数据可能不一致，将校对结算数据。" & vbCrLf & _
        "以下为正确的" & IIf(mbln异常作废, "退费", "") & "结算数据：" & vbCrLf & _
        strShowMsg, vbInformation + vbOKOnly, gstrSysName
    Call 医保数据更正(mobjChargeInfor.病人ID, lng结帐ID, str医保结算, mbln异常作废, Nothing)
    
    '修改校对标志,医保肯定结算成功
    If mbln异常作废 = False And mInsurePara.多单据分单据结算 And gTy_Module_Para.bln只对医保结算成功单据收费 Then
        '通过"医保结算明细"检查是否是“只对医保结算成功单据收费”的异常单据
        strSQL = "Select 1" & vbNewLine & _
                " From 门诊费用记录 A, 医保结算明细 B" & vbNewLine & _
                " Where a.结帐id = b.结帐id(+) And a.No = b.No(+) And a.结帐id = [1] And b.No Is Null And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng结帐ID)
        mbln分单据结算必须全结 = Not rsTemp.EOF
        If mbln分单据结算必须全结 = False Then
            '修改校对标志
            ' Zl_病人门诊收费_医保更新
            strSQL = "Zl_病人门诊收费_医保更新("
            '  结帐id_In   门诊费用记录.结帐id%Type,
            strSQL = strSQL & lng结帐ID & ","
            '  结算序号_In 病人预交记录.结算序号%Type,
            strSQL = strSQL & "Null,"
            '  保险结算_In Varchar2
            strSQL = strSQL & "Null)"
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
        End If
    End If
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
    Dim lng结帐ID As Long, blnCommit As Boolean
    
    If mbytFunc = EM_FUN_重收 Or mbytFunc = EM_FUN_作废 Then
        '并发检查
        If zlIsCheckExistErrBill(mobjChargeInfor.结算序号) = False Then
            MsgBox "当前异常单据已被处理，你不能继续！", vbInformation, gstrSysName
            Exit Sub
        End If
        If zlCheckOtherSessionDoing(mobjChargeInfor.结算序号) Then
            MsgBox "当前单据正在其它收费窗口中进行处理，你不能继续！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If zlIsCheckExiseSingularity(mobjChargeInfor.结算序号) Then
            If mbytFunc = EM_FUN_重收 Then
                MsgBox "该异常单据已经被作废，因此，不能" & IIf(mbytFunc = EM_FUN_重收, "重新收费", "进行作废") & "，请刷新费用列表！", vbInformation, gstrSysName
                Call cmdExit_Click: Exit Sub
            End If
        End If
        If Not zlIsCheckExistErrBill(mobjChargeInfor.结算序号) Then
            MsgBox "该异常单据已经被重新收费，因此，不能" & IIf(mbytFunc = EM_FUN_重收, "重新收费", "进行作废") & "，请刷新费用列表！", vbInformation, gstrSysName
            Call cmdExit_Click: Exit Sub
        End If
    End If
    If mInsurePara.多单据分单据结算 And gTy_Module_Para.bln只对医保结算成功单据收费 And mbln分单据结算必须全结 = False Then
        If mfrmMain.zlSaveBillAndClinicSwapByNo(lng结帐ID, strNos, mcllOverPro, mobjChargeInfor, blnCommit) = False Then
            If blnCommit = False Then Exit Sub
            mobjChargeInfor.结帐ID = lng结帐ID
            mobjChargeInfor.结算序号 = CStr(-1 * lng结帐ID)
            mobjChargeInfor.Nos = strNos
            
            '有提交的单据则直接校对结算信息即可
            Call 医保结算较对
            '重新加载数据
            Call LoadData
            Call LoadPatiInfor
            Call SetControlProperty
        Else
            mobjChargeInfor.结帐ID = lng结帐ID
            mobjChargeInfor.结算序号 = CStr(-1 * lng结帐ID)
            mobjChargeInfor.Nos = strNos
        End If

        mblnSaveBill = True
        mblnYbBalanced = True: mblnCommitBill = True
        cmdExit.Visible = False
    Else
        '数据保存
        If SaveFeeBilL = False Then
            If mblnPriceBillCommit And mblnCommitBill = False Then
                '直接收费时,删除前一个事务提交的划价单
                Call DelMedicareTempNOs
                mblnPriceBillCommit = False
            End If
            Exit Sub
        End If
        '处理医保数据
        If zlInsureClinicSwap = False Then
            If mblnPriceBillCommit And mblnCommitBill = False Then
                '直接收费时,删除前一个事务提交的划价单
                Call DelMedicareTempNOs
                mblnPriceBillCommit = False
            End If
            Exit Sub
        End If
        If mblnElsePersonErrBill Then
            If UpdateElsePersonErrBill() = False Then Exit Sub
        End If
        mblnElsePersonErrBill = False '已更新
    End If

    Call LoadData
    '医保:58344
    mblnYB退款 = mCurCarge.dbl当前未付 + mCurCarge.dbl应缴累计 < 0
    
    Call LoadPatiInfor
    Call SetControlProperty
    '完成医保结算,需要重新设置按钮
    Call SetCtrlVisible
    Call SetControlEnabled
    '光标定位
    '优先使用预交
    If txt冲预交.Visible And txt冲预交.Enabled And gblnPrePayPriority Then
        txt冲预交.SetFocus
        Call SetControlProperty(True)
        Call Show误差金额(True)
    Else
        mblnNotChange = True
        txt冲预交.Text = ""
        mblnNotChange = False
        '70430,冉俊明,2014-4-24,在进行预结算时提示缴款金额，进行医保结算时再次提示相同缴款金额，造成重复提示。
        If txt缴款.Enabled And txt缴款.Visible Then
            mbln医保已报价 = True '先设置已报价为true,屏蔽txt缴款获得焦点而报价
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
    Dim blnSetFocus As Boolean
    
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    Call cbo支付方式_Click
    Call SetControlProperty
    Call SetCtrlVisible
    Call SetControlEnabled
    
    If txt冲预交.Visible Then txt冲预交.Enabled = True
    '光标定位
    If gTy_Module_Para.bln医保结算光标缺省定位 Then
        If cmdYBBalance.Visible And cmdYBBalance.Enabled Then
            cmdYBBalance.SetFocus: blnSetFocus = True
        End If
    End If
    If blnSetFocus = False Then
        If Val(txt冲预交.Text) <> 0 And txt冲预交.Enabled Then
            If txt冲预交.Enabled And txt冲预交.Visible Then txt冲预交.SetFocus
            Call Show误差金额(True)
        Else
            If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
            Call Show误差金额(False)
        End If
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
        If mobjChargeInfor.intInsure <> 0 And mblnYbBalanced = False Then
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
    mstrTittle = "病人收费结算"
    RestoreWinState Me, App.ProductName, mstrTittle
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
    If 医保结算较对 = False Then Unload Me: Exit Sub
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
    '直接收费时,如果划价单已提交但最终没有收费，则要删除前一个事务提交的划价单
    If mblnPriceBillCommit And mblnCommitBill = False Then
        Call DelMedicareTempNOs
        mblnPriceBillCommit = False
    End If
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
    SaveWinState Me, App.ProductName, mstrTittle
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
            If .dbl可用预交 > 0 Then
                txt冲预交.Text = Format(IIf(.dbl可用预交 > .dbl当前未付, .dbl当前未付, .dbl可用预交), "###0.00;###0.00;0.00;0.00")
                txt缴款.Text = Format(.dbl当前未付 - IIf(.dbl可用预交 > .dbl当前未付, .dbl当前未付, .dbl可用预交), "###0.00;###0.00;0.00;0.00")
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
    stbThis.Panels(2).Text = mobjChargeInfor.姓名
    Set rsTemp = GetMoneyInfo(mobjChargeInfor.病人ID, 0, False, 1, False, 0, True)
    Dim dbl家属余额 As Double
    With mCurCarge
        .dbl预交余额 = 0
        .dbl费用余额 = 0
        Do While Not rsTemp.EOF
            .dbl预交余额 = RoundEx(.dbl预交余额 + Val(Nvl(rsTemp!预交余额)), 6)
            .dbl费用余额 = RoundEx(.dbl费用余额 + Val(Nvl(rsTemp!费用余额)), 6)
            If Nvl(rsTemp!家属, 0) = 1 Then
                dbl家属余额 = RoundEx(Val(Nvl(rsTemp!预交余额)) - Val(Nvl(rsTemp!费用余额)), 6)
            End If
            rsTemp.MoveNext
        Loop
        .dbl可用预交 = RoundEx(.dbl预交余额 - .dbl费用余额, 6)
    End With
    If RoundEx(mCurCarge.dbl可用预交, 6) = 0 And RoundEx(dbl家属余额, 6) = 0 Then
        stbThis.Panels(3).Visible = False
    Else
        stbThis.Panels(3).Visible = True
        stbThis.Panels(3).Text = "预交:" & Format(mCurCarge.dbl可用预交, "0.00") & _
            IIf(dbl家属余额 > 0, "(含家属:" & Format(dbl家属余额, "0.00") & ")", "")
    End If
    
    txt医保.Text = Format(mCurCarge.dbl本次医保支付, "###0.00;-###0.00;0.00;0.00;")
    txt合计.Text = Format(mCurCarge.dbl本次实收, "###0.00;-###0.00;0.00;0.00;")
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
    If mobjChargeInfor.intInsure <> 0 And mblnYbBalanced = False Then Exit Sub
    
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

Private Sub stcTittile_GotFocus()
  Clear预交款
End Sub

Private Sub stcTittleTotal_GotFocus()
  Clear预交款
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
    Call Show误差金额(True)
    'If Val(txt冲预交.Tag) = Val(txt冲预交.Text) Then Exit Sub
    
    '自动报价或手工报价时由热键激活
    'Call LedVoiceSpeak
End Sub

Private Sub txt冲预交_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If Val(txt冲预交.Tag) = Val(txt冲预交.Text) Then GoTo SendKeyTab:
    If Val(txt冲预交.Text) = 0 Then GoTo SendKeyTab:
    If CheckPrepayMoneyIsValied = False Then Exit Sub
    If mblnUnloaded Then
        '刷新主界面信息
        ExcuteMainReshData EM_EX_完成
        Unload Me: Exit Sub
    End If
    Exit Sub
SendKeyTab:
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt冲预交_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt冲预交, KeyAscii, m金额式
End Sub

Private Sub txt冲预交_LostFocus()
    If mblnLoad Then Exit Sub
    If Val(txt冲预交.Text) = 0 Then txt缴款.Text = ""
End Sub

Private Sub txt冲预交_Validate(Cancel As Boolean)
    If lbl冲预交.Tag = "1" Then Exit Sub
    If mobjChargeInfor.病人ID = 0 Then Exit Sub
    
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
    Else
        txt冲预交.Text = Format(Val(txt冲预交.Text), "0.00")
    End If
    If Val(txt冲预交.Text) > mCurCarge.dbl可用预交 Then
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
    mblnCurBrushPrepay = True

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
    Else
        txt冲预交.Text = Format(Val(txt冲预交.Text), "0.00")
    End If
    If Val(txt冲预交.Text) > mCurCarge.dbl可用预交 Then
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
    Dim str家属IDs As String
    If zlDatabase.PatiIdentify(Me, glngSys, mobjChargeInfor.病人ID, Val(txt冲预交), mlngModule, 1, mlngBrushCardTypeID, _
            IIf(-1 * gdbl预存款消费验卡 >= Val(txt冲预交), False, True), True, str家属IDs, (gdbl预存款消费验卡 <> 0), (gdbl预存款消费验卡 = 2)) Then
        mobjChargeInfor.家属IDs = str家属IDs
        lbl冲预交.Tag = "1"
       ' txt冲预交.ForeColor = d
       txt冲预交.BackColor = Me.BackColor
       txt冲预交.Tag = Val(txt冲预交)
       txt冲预交.Enabled = False
        If SaveCharge(True) = False Then
            If mblnPriceBillCommit And mblnCommitBill = False Then
                '直接收费时,删除前一个事务提交的划价单
                Call DelMedicareTempNOs
                mblnPriceBillCommit = False
            End If
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
    If mblnYB退款 Then
        MsgBox "当前为退款模式，目前系统暂不支持将退款额退给 " & mCurCardPay.str结算方式 & "！", vbInformation + vbOKOnly, gstrSysName
        If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
        zlControl.TxtSelAll txt缴款: Exit Function
    End If
    If Val(txt缴款) = 0 Then
        MsgBox "未输入交易金额，请检查！", vbInformation + vbOKOnly, gstrSysName
        If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
        zlControl.TxtSelAll txt缴款: Exit Function
    End If
    If Not IsNumeric(txt缴款.Text) Then
        MsgBox "无效数值！", vbInformation + vbOKOnly, gstrSysName
        If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
        zlControl.TxtSelAll txt缴款: Exit Function
    ElseIf Val(txt缴款.Text) > Format(Abs(mCurCarge.dbl当前未付), "0.00") Then
        MsgBox "交易金额不能大于本次未付金额:" & Format(mCurCarge.dbl当前未付, "0.00") & " ！", vbInformation, gstrSysName
        If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
        zlControl.TxtSelAll txt缴款: Exit Function
    End If
    If mCurCardPay.lng医疗卡类别ID > 0 And Not mCurCardPay.bln消费卡 Then
        If Val(txt缴款.Text) <> Format(Abs(mCurCarge.dbl当前未付), "0.00") Then
            If gTy_Module_Para.byt刷卡缺省金额操作 = 1 Then
                If MsgBox("交易金额(" & Format(Val(txt缴款.Text), "0.00") & ")与本次未付金额(" & Format(mCurCarge.dbl当前未付, "0.00") & _
                    ")不同，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            ElseIf gTy_Module_Para.byt刷卡缺省金额操作 = 2 Then
                MsgBox "交易金额(" & Format(Val(txt缴款.Text), "0.00") & ")与本次未付金额(" & Format(mCurCarge.dbl当前未付, "0.00") & _
                    ")不同，不能继续！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    If zlGetClassMoney(mobjChargeInfor.结帐ID, rsMoney) = False Then Exit Function
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
    Optional ByVal strXmlIn As String = "", _
    Optional ByVal str费用来源 As String, _
    Optional ByVal lng病人ID As Long) As Boolean
    '       strXmlIn-三方卡调用XML入参,目前格式如下:
    '       <IN>
    '           <CZLX>0</CZLX>    //操作类型,0-正常调用刷卡,1-转账调用刷卡,2-退款调用刷卡
    '       </IN>
    '       varSquareBalance- Collection类型,返回当前刷卡数据(array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文))
    '       str费用来源 - 当前支付费用的费用来源，多种用逗号分隔(使用消费卡支付时传入)
    '       lng病人ID - 病人ID(使用消费卡支付时传入)
    Set cllSquareBalance = Nothing
    Set mcllCurSquareBalance = Nothing
    If mCurCardPay.bln消费卡 Then
        '构建消费卡的刷卡信息
       Set cllSquareBalance = mcllSquareBalance
     End If
     
    dblMoney = Val(txt缴款.Text)
    If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModule, rsMoney, _
        mCurCardPay.lng医疗卡类别ID, mCurCardPay.bln消费卡, _
    mobjChargeInfor.姓名, mobjChargeInfor.性别, mobjChargeInfor.年龄, dblMoney, mCurCardPay.str刷卡卡号, mCurCardPay.str刷卡密码, _
    False, True, False, False, cllSquareBalance, False, False, "<IN><CZLX>0</CZLX></IN>", mobjChargeInfor.费用来源, mobjChargeInfor.病人ID) = False Then Exit Function
    '消费卡附值
    If mCurCardPay.bln消费卡 Then
        Set mcllCurSquareBalance = cllSquareBalance
    End If
    
    '保存前,一些数据检查
    'zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
    ByVal strCardTypeID As Long, ByVal strCardNo As String, _
    ByVal dblMoney As Double, ByVal strNOs As String, _
    Optional ByVal strXMLExpend As String
    'mobjChargeInfor.strNOs:单独保存时,没有相关时,可能为空.
    If gobjSquare.objSquareCard.zlPaymentCheck(Me, mlngModule, mCurCardPay.lng医疗卡类别ID, _
        mCurCardPay.bln消费卡, mCurCardPay.str刷卡卡号, dblMoney, mobjChargeInfor.Nos, strXMLExpend) = False Then Exit Function
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

Private Function zlGetClassMoney(ByRef lng结帐ID As Long, ByRef rsMoney As ADODB.Recordset) As Boolean
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
    If lng结帐ID = 0 And mbytFunc = EM_FUN_收费 Then
        Call mfrmMain.zlGetClassMoney(rsTemp)
    Else
        strSQL = "" & _
        "   Select  A.收费类别,nvl(sum(实收金额) ,0) as 金额   " & _
        "   From 门诊费用记录 A,(Select 结帐ID From 病人预交记录 where 结帐ID=[1] ) B " & _
        "   Where A.结帐ID=B.结帐ID " & _
        "   Group by 收费类别"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng结帐ID)
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

Private Sub txt合计_GotFocus()
  Clear预交款
End Sub

Private Sub txt缴款_Change()
    Call SetControlProperty
    Call Show误差金额(False)
End Sub

Private Sub txt缴款_GotFocus()
    '只以缴款作为收费结束条件时,必须输入缴款或0
    '刘兴洪:22343
    Call Clear预交款
    If gTy_Module_Para.byt缴款控制 = 1 _
        Or gTy_Module_Para.byt缴款控制 = 3 _
        Or gTy_Module_Para.byt缴款控制 = 2 Then
        If Val(txt缴款.Text) = 0 And Me.ActiveControl Is txt缴款 Then
            txt缴款.Text = ""
        End If
    End If
    With mCurCardPay
       If .bln消费卡 Or (.int性质 <> 1 And mblnYB退款) Then
           '57682:缺省为所有支付金额
           txt缴款.Text = Format(IIf(mblnYB退款, -1, 1) * Val(lbl剩余自付.Caption), "0.00")
        ElseIf mCurCardPay.lng医疗卡类别ID > 0 And Not mCurCardPay.bln消费卡 Then
            If gTy_Module_Para.byt刷卡缺省金额操作 <> 0 Then
                txt缴款.Text = Format(IIf(mblnYB退款, -1, 1) * Val(lbl剩余自付.Caption), "0.00")
            End If
        End If
    End With
    Call SetControlProperty
    Call Show误差金额(False)
  '  Call zlControl.TxtSelAll(txt缴款)
    '自动报价或手工报价时由热键激活
    If mbln医保已报价 Then
        mbln医保已报价 = False
    Else
        Call LedVoiceSpeak
    End If
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
        zl9LedVoice.DispCharge mCurCarge.dbl当前未付 - mCurCarge.dbl本次误差费 + mCurCarge.dbl应缴累计, Val(txt缴款.Text), Val(txt找补.Text)
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
        mbln已报价 = True
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
        txt缴款.Text = Format(Val(txt缴款.Text), "0.00")
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
        If mblnYB退款 Then
            If CSng(txt找补.Text) <= 0 Then
                'LED显示
                'Call ShowLedInfor
                '确定
                 Call cmdOK_Click
            Else
                MsgBox "退款金额不足,请补足退款金额！", vbInformation, gstrSysName
                txt缴款.SetFocus: zlControl.TxtSelAll txt缴款
            End If
        Else
            If CSng(txt找补.Text) >= 0 Then
                'LED显示
                'Call ShowLedInfor
                '确定
                 Call cmdOK_Click
            Else
                MsgBox "缴款金额不足,请补足应缴金额！", vbInformation, gstrSysName
                txt缴款.SetFocus: zlControl.TxtSelAll txt缴款
            End If
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
    '说明：
    '   缺省结算方式的规则，优先顺序如下：
    '   1.如果是连续收费，上次选择的结算方式优先
    '   2.医疗付款方式设置的缺省结算方式
    '   3.模块参数"缺省结算方式"设置的结算方式
    '   4.结算方式缺省为主界面中的刷卡类别
    '   5.结算方式应用中设置的缺省项
    '   6.性质为"1-现金结算方式"的结算方式
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
                varTemp = Split(varData(i) & "||||||", "|")
                If varTemp(6) = Nvl(rsTemp!名称) Then blnFind = True: Exit For
            Next
            If Not blnFind Then
                If Not (Val(Nvl(rsTemp!性质)) = 3 Or Val(Nvl(rsTemp!性质)) = 4 _
                    Or Val(Nvl(rsTemp!性质)) = 7 Or Val(Nvl(rsTemp!性质)) = 8 Or Val(Nvl(rsTemp!应付款)) = 1) Then
                    '不加入医保的结算方式
                    .AddItem Nvl(rsTemp!名称)
                    .ItemData(.NewIndex) = Val(Nvl(rsTemp!性质))
                    mcolCardPayMode.Add Array("", Nvl(rsTemp!名称), 0, 0, 0, 0, Nvl(rsTemp!名称), 0, 0), "K" & j
                    
                    If rsTemp!缺省 = 1 Then .ListIndex = .NewIndex
                    If Val(Nvl(rsTemp!性质)) = 1 And .ListIndex < 0 Then .ListIndex = .NewIndex
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
            varTemp = Split(varData(i) & "||||||", "|")
            rsTemp.Filter = "名称='" & varTemp(6) & "'" '结算方式要设置了"费用"应用场合才能使用
            If Not rsTemp.EOF Then
                .AddItem varTemp(1)
                .ItemData(.NewIndex) = -1
                mcolCardPayMode.Add varTemp, "K" & j
                
                If mbln连续输入 Then
                    '短|全名|刷卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式|是否密文|是否自制卡;…
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
        
        '设置缺省的支付类别
        '注意，结算方式医疗卡显示的是卡类别名称，而不是结算方式
        If Not mbln连续输入 Then
            If gstr结算方式 <> "" Then
                '60574,根据参数设置缺省的支付类别，对于医疗卡参数保存的是名称，而不是结算方式
                For j = 0 To .ListCount - 1
                    If .List(j) = gstr结算方式 Then .ListIndex = j: Exit For
                Next
            End If
            
            If mobjChargeInfor.缺省结算方式 <> "" Then
                '根据医疗付款方式的缺省结算方式设置缺省的支付类别
                '短|全名|刷卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式|是否密文|是否自制卡;…
                For j = 1 To mcolCardPayMode.Count
                    If mcolCardPayMode(j)(6) = mobjChargeInfor.缺省结算方式 Then .ListIndex = j - 1: Exit For
                Next
            End If
        End If
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0
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
   Clear预交款
End Sub

Private Sub txt结算号码_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt结算号码_KeyPress(KeyAscii As Integer)
    If InStr(":'|", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    zlControl.TxtCheckKeyPress txt结算号码, KeyAscii, m文本式
End Sub

Private Sub txt医保_GotFocus()
    Clear预交款
End Sub

Private Sub txt摘要_GotFocus()
    zlCommFun.OpenIme True
    zlControl.TxtSelAll txt摘要
    Clear预交款
End Sub

Private Sub txt摘要_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus
    End If
End Sub

Private Sub txt摘要_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub txt找补_GotFocus()
    zlControl.TxtSelAll txt找补
    Clear预交款
End Sub

Private Function zlOneCardPrayMoney(ByVal dblMoney As Double, ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:一卡通支付
    '返回:一卡通支付成功,返回true,否则返回false
    '编制:刘兴洪
    '日期:2011-08-23 17:57:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl余额 As Double, str医院编码 As String
    On Error GoTo errHandle

    If mCurCardPay.blnOneCard = False Then zlOneCardPrayMoney = True: Exit Function
    
    mrsOneCard.Filter = "结算方式='" & mCurCardPay.str结算方式 & "'"
    If mrsOneCard.EOF Then zlOneCardPrayMoney = True: Exit Function
    
    '一卡通结算（修改单据时因为没有读卡，无法确定使用了哪种一卡通，所以不支持修改功能)
    Dim intCardType As Integer, strSwapNO As String
    If Not mobjICCard.PaymentSwap(dblMoney, dbl余额, intCardType, Val("" & mrsOneCard!医院编码), mCurCardPay.str刷卡卡号, strSwapNO, mobjChargeInfor.结帐ID, mobjChargeInfor.病人ID) Then
        gcnOracle.RollbackTrans
        MsgBox mCurCardPay.str结算方式 & "结算失败!", vbOKOnly, gstrSysName
        Exit Function
    End If
    mblnThreeInterface = True
    gstrSQL = "zl_一卡通结算_Update(" & 0 & ",'" & mCurCardPay.str结算方式 & "','" & mCurCardPay.str刷卡卡号 & "','" & intCardType & "','" & strSwapNO & "'," & dbl余额 & "," & mobjChargeInfor.结帐ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    zlOneCardPrayMoney = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
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
    
    On Error GoTo errHandle
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
    If gobjSquare.objSquareCard.zlPaymentMoney(Me, mlngModule, mCurCardPay.lng医疗卡类别ID, mCurCardPay.bln消费卡, _
                mCurCardPay.str刷卡卡号, mobjChargeInfor.结帐ID, mCurCardPay.strNo, dblMoney, strSwapGlideNO, _
                strSwapMemo, strSwapExtendInfor) = False Then gcnOracle.RollbackTrans: Exit Function
    '更新三交交易数据
    mblnThreeInterface = True
    
    If mCurCardPay.lng医疗卡类别ID <> 0 And mobjChargeInfor.结帐ID <> 0 And cbo支付方式.Visible Then
        mCurCardPay.str交易流水号 = strSwapGlideNO
        mCurCardPay.str交易说明 = strSwapMemo
        If mCurCardPay.bln消费卡 = False Then
            Call zlAddUpdateSwapSQL(False, mobjChargeInfor.结帐ID, mCurCardPay.lng医疗卡类别ID, mCurCardPay.bln消费卡, mCurCardPay.str刷卡卡号, strSwapGlideNO, strSwapMemo, cllUpdate, 2)
        End If
        Call zlAddThreeSwapSQLToCollection(False, mobjChargeInfor.结帐ID, mCurCardPay.lng医疗卡类别ID, mCurCardPay.bln消费卡, mCurCardPay.str刷卡卡号, strSwapExtendInfor, cllThreeSwap)
    End If
    zlInterfacePrayMoney = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ChargeOver(ByVal dbl退支票额 As Double) As Boolean
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
    
    ' Zl_门诊收费结算_Modify
    strSQL = "Zl_门诊收费结算_Modify("
    '    操作类型_In   Number,
    '    --操作类型_In:
    '    --   0-普通收费方式:
    '    --     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
    '    --     ②冲预交_In:如果涉及预交款,则传入本次的冲预交,非正常收费时,传入零
    '    --     ③退支票额_In:如果涉及退支票,则传入本次的退支票额,非正常收费时,传入零
    '    --   1.三方卡结算:
    '    --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
    '    --     ②冲预交_In: 传入零
    '    --     ③退支票额_In:传入零
    '    --     ④卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
    '    --   2-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
    '    --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
    '    --     ②冲预交_In: 传入零
    '    --     ③退支票额_In:传入零
    '    --   3-消费卡结算:
    '    --     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."
    '    --     ②冲预交_In: 传入零
    '    --     ③退支票额_In:传入零
    strSQL = strSQL & 0 & ","
    '    病人id_In     门诊费用记录.病人id%Type,
    strSQL = strSQL & mobjChargeInfor.病人ID & ","
    '    结帐id_In     病人预交记录.结帐id%Type,
    strSQL = strSQL & mobjChargeInfor.结帐ID & ","
    '    结算方式_In   Varchar2,
    strSQL = strSQL & "'" & str收费结算 & "'" & ","
    '    冲预交_In     病人预交记录.冲预交%Type := Null,
    strSQL = strSQL & dbl预存款 & ","
    '    退支票额_In   病人预交记录.冲预交%Type := Null,
    strSQL = strSQL & dbl退支票额 & ","
    '    卡类别id_In   病人预交记录.卡类别id%Type := Null,
    strSQL = strSQL & "NULL,"
    '    卡号_In       病人预交记录.卡号%Type := Null,
    strSQL = strSQL & "NULL,"
    '    交易流水号_In 病人预交记录.交易流水号%Type := Null,
    strSQL = strSQL & "NULL,"
    '    交易说明_In   病人预交记录.交易说明%Type := Null,
    strSQL = strSQL & "NULL,"
    '    缴款_In       病人预交记录.缴款%Type := Null,
    strSQL = strSQL & "" & dbl缴款 & ","
    '    找补_In       病人预交记录.找补%Type := Null,
    strSQL = strSQL & "" & dbl找补 & ","
    '    误差金额_In   门诊费用记录.实收金额%Type := Null,
    '    -- 误差金额_In:存在误差费时,传入
    strSQL = strSQL & "" & mCurCarge.dbl本次误差费 & ","
    '    完成结算_In Number:=0
    '    -- 完成结算_In:1-完成收费;0-未完成收费
    strSQL = strSQL & "1,"
    '77141,冉俊明,2014-8-26,给零费用病人收费/退费后,没有结算信息
    '缺省结算方式_In 结算方式.名称%Type := Null
    strSQL = strSQL & "'" & Trim(cbo支付方式.Text) & "',"
    '79868,冉俊明,2015-06-10,使用病人家属预交
    '冲预交病人ids_In Varchar2:=Null
    strSQL = strSQL & "'" & mobjChargeInfor.病人ID & "," & mobjChargeInfor.家属IDs & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    mobjChargeInfor.缴款 = dbl缴款: mobjChargeInfor.找补 = dbl找补
    ChargeOver = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
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
        mCurCarge.dbl本次误差费 = RoundEx(mCurCarge.dbl本次实收 - mCurCarge.dbl本次已付合计 - RoundEx(mCurCarge.dbl当前未付, 2), 6)
    ElseIf mCurCardPay.int性质 = 1 Then
        dblTemp = IIf(dblMoney = 0, dbl剩余金额, mCurCarge.dbl当前未付): dbl剩余金额 = 0
        If mobjChargeInfor.intInsure > 0 Then  '问题:43855
            If mInsurePara.分币处理 Then
                dblMoney = CentMoney(CCur(dblTemp))
            Else
                dblMoney = Format(dblTemp, "0.00")
            End If
        Else
            dblMoney = CentMoney(CCur(dblTemp))
        End If
        mCurCarge.dbl本次误差费 = mCurCarge.dbl本次实收 - mCurCarge.dbl本次已付合计 - dblMoney
    Else
        mCurCarge.dbl本次误差费 = mCurCarge.dbl本次实收 - mCurCarge.dbl本次已付合计 - RoundEx(mCurCarge.dbl当前未付, 2)
    End If
    
    '问题:47637
    '未进行医保结算前,不显示误差
    If mobjChargeInfor.intInsure <> 0 And mblnYbBalanced = False Then mCurCarge.dbl本次误差费 = 0
    mCurCarge.dbl本次误差费 = RoundEx(mCurCarge.dbl本次误差费, 6)
    pic误差.Visible = mCurCarge.dbl本次误差费 <> 0
    lbl误差额.Caption = FormatEx(mCurCarge.dbl本次误差费, 6)
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
    If mobjChargeInfor.intInsure <> 0 And mblnYbBalanced Then intCount = intCount + 1: strErrMsg = strErrMsg & "医保结算:" & txt医保.Text
   With vsBlance
        For i = 1 To .Rows - 1
            str结算方式 = Trim(.TextMatrix(i, .ColIndex("支付方式")))
            int性质 = Val(.RowData(i))
            'rowdata:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
            If InStr("34", int性质) > 0 Then
                If int性质 = 4 Then intCount = intCount + 1
                If int性质 = 3 Then '三方接口
                    intCount = intCount + 1: strErrMsg = strErrMsg & vbCrLf & str结算方式 & ":" & .TextMatrix(i, .ColIndex("支付金额"))
                End If
            End If
        Next
    End With
    If intCount > 2 Then
        Call MsgBox("注意:" & vbCrLf & "   本系统目前只支持两种以下接口,现在已经存在如下接口交易:" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName)
        Exit Function
    End If
    zlCheckMulitInterfaceNumValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function SaveCharge(Optional bln预交 As Boolean, _
    Optional ByRef blnUnload As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存结算数据
    '出参:blnUnload-是否收费完成，退出后，将Unload界面
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-14 17:38:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str消费卡结算 As String, str收费结算 As String, strSQL As String
    Dim strCardNo As String, strErrMsg As String
    Dim blnHaveMoney As Boolean, blnFind As Boolean, blnTrans As Boolean
    Dim dbl剩余金额 As Double, dblTemp As Double
    Dim dblMoney As Double, dbl退支票额 As Double
    Dim i As Integer, j As Long
    Dim cllUpdate As Collection, cllThreeSwap As Collection, cllPro As Collection
    Dim objCard As Card, dblCheckMoney As Double
    
    On Error GoTo errHandle
    
    blnUnload = False
    If zlCheckMulitInterfaceNumValied = False Then Exit Function
    
    mobjChargeInfor.收费结算 = "" '问题:42791
    
    mdbl现金 = 0
    dblMoney = IIf(bln预交, Val(txt冲预交.Text), IIf(mblnYB退款, -1, 1) * Val(txt缴款.Text)) - IIf(mblnCur连续, 0, mCurCarge.dbl应缴累计)
    dbl退支票额 = 0
    dbl剩余金额 = mCurCarge.dbl当前未付 - dblMoney
    
    '将界面检查的金额和实际保存的金额分开（主要因为继续收费时，应缴金额和实际数据保存金额不一致）
    If Val(txt缴款.Text) = 0 Then
        dblCheckMoney = mCurCarge.dbl当前未付 + IIf(mblnCur连续, 0, mCurCarge.dbl应缴累计)
    Else
        dblCheckMoney = IIf(mblnYB退款, -1, 1) * Val(txt缴款.Text)
    End If
    If mblnCur连续 = False And mCurCarge.dbl应缴累计 <> 0 Then
        dblMoney = IIf(dblMoney = 0, dbl剩余金额, mCurCarge.dbl当前未付)
    End If
    
    If bln预交 Then
        dblMoney = Val(txt冲预交.Text)
        mobjChargeInfor.收费结算 = mobjChargeInfor.收费结算 & "|冲预交:" & dblMoney
        '问题:58344
        '检查是否当前支付金额为负数,是负数时,需要提醒操作员(主要是医保结算时可能大于本身单据的费用)
        If mblnYB退款 And mblnCur连续 = False Then
              Call MsgBox("注意:" & vbCrLf & "    当前处于退款方式,不允许使用预交款!", vbExclamation + vbOKOnly + vbDefaultButton2, gstrSysName)
              Exit Function
        End If
        
    ElseIf mCurCardPay.int性质 = 1 Then
        dblTemp = IIf(dblMoney = 0, dbl剩余金额, mCurCarge.dbl当前未付): dbl剩余金额 = 0
        If mobjChargeInfor.intInsure > 0 Then  '问题:43855
            If gclsInsure.GetCapability(support分币处理, , mobjChargeInfor.intInsure) Then
                dblMoney = CentMoney(CCur(dblTemp))
                dblCheckMoney = CentMoney(CCur(dblCheckMoney))
            Else
                dblMoney = Format(dblTemp, "0.00")
                dblCheckMoney = Format(dblCheckMoney, "0.00")
            End If
        Else
            dblMoney = CentMoney(CCur(dblTemp))
            dblCheckMoney = CentMoney(CCur(dblCheckMoney))
        End If
        '问题:58344
        '检查是否当前支付金额为负数,是负数时,需要提醒操作员(主要是医保结算时可能大于本身单据的费用)
        If mblnYB退款 And mblnCur连续 = False Then
             If MsgBox("注意:" & vbCrLf & "    未付部分为退款,你是否真的要退『" & mCurCardPay.str结算方式 & ":" & Abs(dblCheckMoney) & "』给病人?" & vbCrLf & IIf(Val(txt缴款.Text) <> 0, "  当前退给病人总额:" & txt缴款.Text & vbCrLf & "  当前应收回总额:" & Abs(txt找补.Text), ""), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
             If Abs(dblCheckMoney) < Abs(lbl剩余自付.Caption) Then
                Call MsgBox("注意:" & vbCrLf & "    未付部分为退款,你不能进行多次退款操作," & vbCrLf & "当前退金额(" & Format(dblCheckMoney, "0.00") & ")必须大于剩余金额(" & lbl剩余自付.Caption & ")!", vbInactiveBorder + vbOKOnly + vbDefaultButton2, gstrSysName)
                Exit Function
             End If
        End If
        mdbl现金 = dblMoney
        If Val(txt缴款.Text) <> 0 Then
            mobjChargeInfor.收费结算 = mobjChargeInfor.收费结算 & "|缴款:" & IIf(mblnYB退款, -1, 1) * Val(txt缴款.Text) & ":1"
            mobjChargeInfor.收费结算 = mobjChargeInfor.收费结算 & "|找补:" & IIf(mblnYB退款, -1, 1) * Val(txt找补.Text) & ":2"
        End If
        mobjChargeInfor.收费结算 = mobjChargeInfor.收费结算 & "|" & mCurCardPay.str结算方式 & ":" & dblMoney
    ElseIf mCurCardPay.bln支票 Then
        mobjChargeInfor.收费结算 = mobjChargeInfor.收费结算 & "|" & mCurCardPay.str结算方式 & ":" & dblMoney
        '问题:58344
        '检查是否当前支付金额为负数,是负数时,需要提醒操作员(主要是医保结算时可能大于本身单据的费用)
        If mblnYB退款 And mblnCur连续 = False Then
             If MsgBox("注意:" & vbCrLf & "    未付部分为退款,你是否真的要退『" & mCurCardPay.str结算方式 & ":" & Abs(dblCheckMoney) & "』给病人?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
             If Abs(dblCheckMoney) <> Abs(lbl剩余自付.Caption) Then
                Call MsgBox("注意:" & vbCrLf & "    未付部分为退款,当前退金额(" & Format(Abs(dblCheckMoney), "0.00") & ")必须等于剩余金额(" & Abs(Val(lbl剩余自付.Caption)) & ")!", vbInactiveBorder + vbOKOnly + vbDefaultButton2, gstrSysName)
                Exit Function
             End If
        Else
            If dbl剩余金额 < 0 Then
                If mstr退支票 = "" Then
                    MsgBox "在结算方式中没有设置应付款的结算方式,不能进行退支票处理", vbOKOnly + vbInformation, gstrSysName
                    Exit Function
                End If
                dbl退支票额 = -1 * Val(txt找补.Text)
                mobjChargeInfor.收费结算 = mobjChargeInfor.收费结算 & "|" & mstr退支票 & ":" & -1 * dbl退支票额 & ":2"
            End If
        End If
    Else
        '问题:58344
        '检查是否当前支付金额为负数,是负数时,需要提醒操作员(主要是医保结算时可能大于本身单据的费用)
        If mblnYB退款 And mblnCur连续 = False Then
             If MsgBox("注意:" & vbCrLf & "    未付部分为退款,你是否真的要退『" & mCurCardPay.str结算方式 & ":" & Abs(dblCheckMoney) & "』给病人?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
             If Abs(dblCheckMoney) <> Abs(lbl剩余自付.Caption) Then
                Call MsgBox("注意:" & vbCrLf & "    未付部分为退款,当前退金额(" & Format(Abs(dblCheckMoney), "0.00") & ")必须等于剩余金额(" & Abs(Val(lbl剩余自付.Caption)) & ")!", vbInactiveBorder + vbOKOnly + vbDefaultButton2, gstrSysName)
                Exit Function
             End If
        End If
        mobjChargeInfor.收费结算 = mobjChargeInfor.收费结算 & "|" & mCurCardPay.str结算方式 & ":" & dblMoney
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
    If RoundEx(dbl剩余金额, 2) > 0 Then blnHaveMoney = True
    With vsBlance
        blnFind = False
        For i = 1 To .Rows - 1
            ' '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
            If bln预交 Then
                If Val(.RowData(i)) = 1 Then blnFind = True
            ElseIf mCurCardPay.bln消费卡 And mCurCardPay.bln自制卡 Then
                '消费卡,已经检查,不用再处理
            Else
                If .TextMatrix(i, .ColIndex("支付方式")) = mCurCardPay.str结算方式 Then
                    blnFind = True
                End If
            End If
            mobjChargeInfor.收费结算 = mobjChargeInfor.收费结算 & "|" & .TextMatrix(i, .ColIndex("支付方式")) & ":" & .TextMatrix(i, .ColIndex("支付金额"))
        Next
        
        If blnFind Then
            If bln预交 Then
                MsgBox "已经用预存款支付,只有删除预存款后才能支付!", vbOKOnly, gstrSysName
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
    If mCurCardPay.bln消费卡 And Not bln预交 Then
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
    
    If Not (bln预交 Or mCurCardPay.lng医疗卡类别ID = 0 Or cbo支付方式.ItemData(cbo支付方式.ListIndex) <> -1) Then
        '第三方接口的相关结算,需要先处理接口数据
        
        '先保存单据
        blnTrans = True
        If SaveFeeBilL = False Then blnTrans = False: Exit Function
        If mblnElsePersonErrBill Then
            If UpdateElsePersonErrBill() = False Then blnTrans = False: Exit Function
        End If
        
        ' Zl_门诊收费结算_Modify
        strSQL = "Zl_门诊收费结算_Modify("
        '    操作类型_In   Number,
        '    --操作类型_In:
        '    --   0-普通收费方式:
        '    --     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
        '    --     ②冲预交_In:如果涉及预交款,则传入本次的冲预交,非正常收费时,传入零
        '    --     ③退支票额_In:如果涉及退支票,则传入本次的退支票额,非正常收费时,传入零
        '    --   1.三方卡结算:
        '    --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
        '    --     ②冲预交_In: 传入零
        '    --     ③退支票额_In:传入零
        '    --     ④卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
        '    --   2-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
        '    --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
        '    --     ②冲预交_In: 传入零
        '    --     ③退支票额_In:传入零
        '    --   3-消费卡结算:
        '    --     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."
        '    --     ②冲预交_In: 传入零
        '    --     ③退支票额_In:传入零
        If mCurCardPay.bln消费卡 Then
            strSQL = strSQL & "3" & ","
        Else
            strSQL = strSQL & "1" & ","
        End If
        '    病人id_In     门诊费用记录.病人id%Type,
        strSQL = strSQL & mobjChargeInfor.病人ID & ","
        '    结帐id_In     病人预交记录.结帐id%Type,
        strSQL = strSQL & mobjChargeInfor.结帐ID & ","
        '    结算方式_In   Varchar2,
        If mCurCardPay.bln消费卡 Then
            strSQL = strSQL & "'" & str消费卡结算 & "'" & ","
        Else
            '"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
            str收费结算 = mCurCardPay.str结算方式
            str收费结算 = str收费结算 & "|" & dblMoney
            str收费结算 = str收费结算 & "|" & IIf(txt结算号码.Text = "", " ", txt结算号码.Text)
            str收费结算 = str收费结算 & "|" & IIf(txt摘要.Text = "", " ", txt摘要.Text)
            strSQL = strSQL & "'" & str收费结算 & "'" & ","
        End If
        '    冲预交_In     病人预交记录.冲预交%Type := Null,
        strSQL = strSQL & "NULL,"
        '    退支票额_In   病人预交记录.冲预交%Type := Null,
        strSQL = strSQL & "NULL,"
        '    卡类别id_In   病人预交记录.卡类别id%Type := Null,
        strSQL = strSQL & "" & IIf(mCurCardPay.lng医疗卡类别ID = 0, "NULL", mCurCardPay.lng医疗卡类别ID) & ","
        '    卡号_In       病人预交记录.卡号%Type := Null,
        strSQL = strSQL & "" & IIf(mCurCardPay.str刷卡卡号 <> "", "'" & mCurCardPay.str刷卡卡号 & "'", "NULL") & ","
        '    交易流水号_In 病人预交记录.交易流水号%Type := Null,
        strSQL = strSQL & "NULL,"
        '    交易说明_In   病人预交记录.交易说明%Type := Null,
        strSQL = strSQL & "NULL,"
        '    缴款_In       病人预交记录.缴款%Type := Null,
        strSQL = strSQL & "NULL,"
        '    找补_In       病人预交记录.找补%Type := Null,
        strSQL = strSQL & "NULL,"
        '    误差金额_In   门诊费用记录.实收金额%Type := Null,
        '    -- 误差金额_In:存在误差费时,传入
        strSQL = strSQL & "NULL,"
        '    完成结算_In Number:=0
        '    -- 完成结算_In:1-完成收费;0-未完成收费
        strSQL = strSQL & "0)"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        If Not mCurCardPay.bln消费卡 Then
             If zlInterfacePrayMoney(cllUpdate, cllThreeSwap, dblMoney) = False Then blnTrans = False: Exit Function
        End If
        
        '一卡通交易(老版)
        If zlOneCardPrayMoney(dblMoney, strErrMsg) = False Then
            gcnOracle.RollbackTrans
            MsgBox strErrMsg, vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
        
        gcnOracle.CommitTrans:  mblnCommitBill = True
        mblnElsePersonErrBill = False '已更新
        Call zlExecuteProcedureArrAy(cllUpdate, Me.Caption)
        
        blnTrans = False
        Call SetCtrlVisible
         blnTrans = True
        Call zlExecuteProcedureArrAy(cllThreeSwap, Me.Caption)
         blnTrans = False
    End If
GoOver:
    If mobjChargeInfor.intInsure <> 0 Then
        If Not (bln预交 Or mCurCardPay.lng医疗卡类别ID <> 0 _
            Or mCurCardPay.blnOneCard) Then
            '只有医保病人才会出现重新较对的情况,因此才会重新计算本次应缴的情况
            '主要是更改连续收费的问题
            mobjChargeInfor.本次应缴 = mobjChargeInfor.本次应缴 + dblMoney
        End If
    End If
    
    If Not blnHaveMoney Then
        blnTrans = True
        If SaveFeeBilL = False Then blnTrans = False: Exit Function
        If mblnElsePersonErrBill Then
            If UpdateElsePersonErrBill() = False Then blnTrans = False: Exit Function
        End If
        If ChargeOver(dbl退支票额) = False Then blnTrans = False:    Exit Function
        gcnOracle.CommitTrans:  mblnCommitBill = True
        mblnElsePersonErrBill = False '已更新
        blnTrans = False
        Call WhriteTotalDataToReCord(IIf(bln预交, dblMoney, 0), IIf(Not bln预交, dblMoney, 0), dbl退支票额)
        mblnOK = True: SaveCharge = True: mblnUnloaded = True
        blnUnload = True
         Exit Function
    End If
    
    mobjChargeInfor.收费结算 = ""
    If Not bln预交 And mCurCardPay.int性质 = 1 Then
       '现金
        SaveCharge = True: Exit Function
    End If
    
    Err = 0: On Error GoTo errHandle:
    With vsBlance
        If mCurCardPay.bln消费卡 And Not bln预交 Then
            If mcllSquareBalance Is Nothing Then Set mcllSquareBalance = New Collection
            
            For j = 1 To mcllCurSquareBalance.Count
                '当前刷卡数据(array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文))
                mcllSquareBalance.Add mcllCurSquareBalance(j)
                If Not (.Rows = 2 And Trim(.TextMatrix(1, .ColIndex("支付方式"))) = "") Then
                    .Rows = .Rows + 1
                    .RowPosition(.Rows - 1) = 1
                End If
                '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                .RowData(1) = 5
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
            '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
            .RowData(1) = 0
            strCardNo = mCurCardPay.str刷卡卡号
            If bln预交 Then
                .TextMatrix(1, .ColIndex("支付方式")) = "预存款"
                .RowData(1) = 1
            ElseIf mCurCardPay.lng医疗卡类别ID <> 0 Then
                Set objCard = GetPayCard(mCurCardPay.lng医疗卡类别ID, mCurCardPay.bln消费卡, False)
                .TextMatrix(1, .ColIndex("支付方式")) = objCard.结算方式
                If Not objCard Is Nothing Then
                    '医疗卡类别ID|消费卡(1, 0) |自制卡|是否全退|是否退现|接口名称
                    .Cell(flexcpData, 1, .ColIndex("支付方式")) = mCurCardPay.lng医疗卡类别ID & "|" & IIf(mCurCardPay.bln消费卡, 1, 0) & "|" & IIf(mCurCardPay.bln自制卡, 1, 0) & "|" & IIf(objCard.是否全退, 1, 0) & "|" & IIf(objCard.是否退现, 1, 0) & "|" & mCurCardPay.str名称
                Else
                    .Cell(flexcpData, 1, .ColIndex("支付方式")) = mCurCardPay.lng医疗卡类别ID & "|" & IIf(mCurCardPay.bln消费卡, 1, 0) & "|" & IIf(mCurCardPay.bln自制卡, 1, 0) & "|" & 0 & "|" & 0 & "|" & mCurCardPay.str名称
                End If
                .RowData(1) = 3
                strCardNo = gobjSquare.objSquareCard.zlGetCardNODencode(mCurCardPay.str刷卡卡号, mCurCardPay.lng医疗卡类别ID, mCurCardPay.bln消费卡)
            ElseIf mCurCardPay.blnOneCard Then
                
                .TextMatrix(1, .ColIndex("支付方式")) = mCurCardPay.str结算方式
                .RowData(1) = 4
            Else
                .TextMatrix(1, .ColIndex("支付方式")) = mCurCardPay.str结算方式
            End If
            .TextMatrix(1, .ColIndex("支付金额")) = Format(dblMoney, "0.00")
            If Not bln预交 Then
                .TextMatrix(1, .ColIndex("结算号码")) = IIf(txt结算号码.Visible, Trim(txt结算号码.Text), "")
                .TextMatrix(1, .ColIndex("备注")) = Trim(txt摘要.Text)
                
                .TextMatrix(1, .ColIndex("卡号")) = IIf(mCurCardPay.bln卡号密文, String(Len(strCardNo), "*"), strCardNo)
                .Cell(flexcpData, 1, .ColIndex("卡号")) = mCurCardPay.str刷卡卡号
                .TextMatrix(1, .ColIndex("交易流水号")) = mCurCardPay.str交易流水号
                .TextMatrix(1, .ColIndex("交易说明")) = mCurCardPay.str交易说明
            End If
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
errHandle:
    If blnTrans Then
        gcnOracle.RollbackTrans
        Call ErrCenter
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
    End If
End Function

Private Function UpdateElsePersonErrBill()
    '他人异常单据，将这部分单据更新为当前操作员
    Dim strSQL As String
    
    On Error GoTo errHandler
    'Zl_门诊异常收费_更新操作员
    strSQL = "Zl_门诊异常收费_更新操作员("
    '病人id_In     门诊费用记录.病人id%Type,
    strSQL = strSQL & "" & mobjChargeInfor.病人ID & ","
    '操作员编号_In 门诊费用记录.操作员编号%Type,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '操作员姓名_In 门诊费用记录.操作员姓名%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '结算序号_In   病人预交记录.结算序号%Type
    strSQL = strSQL & mobjChargeInfor.结算序号 & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    UpdateElsePersonErrBill = True
    Exit Function
errHandler:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
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
    If vsBlance.Row < 1 Then
        int性质 = -1
    ElseIf vsBlance.TextMatrix(vsBlance.Row, vsBlance.ColIndex("支付方式")) = "" Then
        int性质 = -1
    Else
        int性质 = Val(vsBlance.RowData(vsBlance.Row))
    End If
     '.rowdata: '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
    cmdDel.Visible = (int性质 = 0 Or int性质 = 1) And mbytFunc <> EM_FUN_作废
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
     ByRef lng结帐ID As Long, ByVal dbl退款金额 As Double, _
     ByVal bln是否退款验卡 As Boolean) As Boolean
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
        "3|" & lng结帐ID, dbl退款金额, strSwapNO, strSwapMemo, strXMLExend) = False Then
          zlCheckDelValied = False
          Exit Function
     End If
     
     If bln是否退款验卡 Then
       '弹出刷卡界面
        'zlBrushCard(frmMain As Object, _
        ByVal lngModule As Long, _
        ByVal rsClassMoney As ADODB.Recordset, _
        ByVal lngCardTypeID As Long, _
        ByVal bln消费卡 As Boolean, _
        ByVal strPatiName As String, ByVal strSex As String, _
        ByVal strOld As String, ByVal dbl金额 As Double, _
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
        Dim strPassWord As String
        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModule, Nothing, lng卡类别ID, _
            bln消费卡, mobjChargeInfor.姓名, mobjChargeInfor.性别, mobjChargeInfor.年龄, dbl退款金额, strCardNo, strPassWord, _
            True, True, False, True, Nothing, False, False, "<IN><CZLX>2</CZLX></IN>") = False Then Exit Function
    End If
goEnd:
    zlCheckDelValied = True
    Exit Function
End Function

Private Function CallBackBalanceInterface(ByVal lng冲销ID As Long, ByVal lng原结帐ID As Long, _
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
    Dim strSwapExtendInfor As String, strTemp As String
    
    Err = 0: On Error GoTo Errhand:
    
    '卡类别ID,卡号,是否消费卡(1-是;0-否),交易流水号,交易说明,strNO
    If lng卡类别ID = 0 Then CallBackBalanceInterface = True: Exit Function
    '81489,冉俊明,2015-1-22,退费传入冲销ID
    strSwapExtendInfor = "3|" & lng冲销ID: strTemp = strSwapExtendInfor
    
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
    '       strSwapExtendInfor-传入，本次退费的冲销ID：
    '                           格式:收费类型1|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '                           收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
    '       strSwapExtendInfor-传出，交易的扩展信息
    '           格式为:项目名称1|项目内容2||…||项目名称n|项目内容n 每个项目中不能包含|字符
    If gobjSquare.objSquareCard.zlReturnMoney(Me, mlngModule, lng卡类别ID, bln消费卡, strCardNo, "3|" & lng原结帐ID, dblMoney, strSwapNO, strSwapMemo, strSwapExtendInfor) = False Then Exit Function
    
    Set cllUpdate = New Collection: Set cllThreeSwap = New Collection
    Call zlAddUpdateSwapSQL(False, lng冲销ID, lng卡类别ID, bln消费卡, strCardNo, strSwapNO, strSwapMemo, cllUpdate)
    If strTemp <> strSwapExtendInfor Then
        Call zlAddThreeSwapSQLToCollection(False, lng冲销ID, lng卡类别ID, bln消费卡, strCardNo, strSwapExtendInfor, cllThreeSwap)
    End If
    CallBackBalanceInterface = True
Errhand:
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
        Set mrsBlance = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mobjChargeInfor.结帐ID)
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
    blnEdit = (mobjChargeInfor.intInsure = 0 Or mobjChargeInfor.intInsure <> 0 And mblnYbBalanced) And mbytFunc <> EM_FUN_作废
    blnEdit = blnEdit Or (mbytFunc = EM_FUN_重收 And (mblnYbBalanced Or mobjChargeInfor.intInsure = 0))
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
            ' 0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
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
            dblMoney = IIf(mblnYB退款, -1, 1) * Val(txt缴款.Text) - IIf(mblnCur连续, 0, mCurCarge.dbl应缴累计)
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
                    '77183:刘尔旋,2014-08-27,现金结算时没有填摘要的问题
                    str收费结算 = str收费结算 & "| " & IIf(Trim(txt摘要) = "", " ", Trim(txt摘要))
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

Private Function zlInsureClinicSwap() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医保调用
    '入参:blnModifyBill-是否修改单据
    '       strBalanceIDs:本次结帐的ID,分别用逗号分离
    '       strSaveNos-保存的单据号
    '出参:strSaveNos-返回已经结算成功的单据号
    '       blnAffair-是否事务处理
    '       strSaveSucessNos-保存成功的票据(对划价有效)
    '返回:医保调用成功或非医保,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-20 17:15:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varNos As Variant, strSQL As String
    Dim strBillNO As String, blnTrans As Boolean, blnTransMedicare As Boolean
    Dim p As Integer, strAdvance As String
    Dim strTmp As String, i As Long
 
    On Error GoTo errHandle
    If mobjChargeInfor.intInsure = 0 Then zlInsureClinicSwap = True: Exit Function
    blnTrans = True
'    '1. 保存为划价单
'    If mblnSavePrice Then
'        '保存为划价单
'        '如果是联合医保,收费确定时实际却保存为划价单:传划价单明细,不在Oracle事务中执行
'        varNos = Split(mobjChargeInfor.Nos, ",")
'        For p = 1 To UBound(varNos)
'            strBillNO = mobjChargeInfor(p)
'            If Not gclsInsure.TranChargeDetail(1, strBillNO, 1, 0, "", , mobjChargeInfor.intInsure) Then
'                '删除划价单(继续处理)
'                Call DelMedicareTempNO(True, strBillNO)
'                gcnOracle.RollbackTrans: Exit Function
'            End If
'        Next
'        mblnYbBalanced = True   '医保已经结算
'        zlInsureClinicSwap = True
'        Exit Function
'    End If
      
    If mInsurePara.医保接口打印票据 And mobjChargeInfor.医保不走票号 = False Then
        '不严格控制票据时保存当前票号
        If Not gblnStrictCtrl Then
            zlDatabase.SetPara "当前收费票据号", mobjChargeInfor.当前发票号, glngSys, 1121, zlstr.IsHavePrivs(mstrPrivs, "参数设置")
        End If
    End If
    
    strAdvance = CStr(-1 * mobjChargeInfor.结帐ID)
    Dim blnCommit As Boolean '部分执行成功或全部执行成功都会提交数据
    If Not mfrmMain.zlInsureClinicSwap(mobjChargeInfor.结帐ID, mobjChargeInfor.intInsure, strAdvance, blnCommit) Then
        '异常重收时无论有没有数据提交都允许返回，正常收费结算只有未提交数据前才允许返回
        cmdExit.Visible = (cmdExit.Visible And Not blnCommit) Or mbytFunc = EM_FUN_重收
        mblnCommitBill = mblnCommitBill Or blnCommit
        gcnOracle.RollbackTrans: Exit Function
    End If
    
    mblnYbBalanced = True   '医保已经结算
    blnTransMedicare = True
    
    If strAdvance = CStr(-1 * mobjChargeInfor.结帐ID) Then strAdvance = ""
    
    If Not zlInsureCheck(mobjChargeInfor.预结结算, strAdvance) Or strAdvance = "" Then
        '修改校对标志
        ' Zl_病人门诊收费_医保更新
        strSQL = "Zl_病人门诊收费_医保更新("
        '  结帐id_In   门诊费用记录.结帐id%Type,
        strSQL = strSQL & mobjChargeInfor.结帐ID & ","
        '  结算序号_In 病人预交记录.结算序号%Type,
        strSQL = strSQL & "Null,"
        '  保险结算_In Varchar2
        strSQL = strSQL & "Null)"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        gcnOracle.CommitTrans: blnTrans = False: mblnCommitBill = True
        If Not mInsurePara.多单据分单据结算 Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicSwap, True, mobjChargeInfor.intInsure)
        zlInsureClinicSwap = True: Exit Function
    End If
    
    Call 医保数据更正(mobjChargeInfor.病人ID, mobjChargeInfor.结帐ID, strAdvance, False, Nothing)
    '修改校对标志
    ' Zl_病人门诊收费_医保更新
    strSQL = "Zl_病人门诊收费_医保更新("
    '  结帐id_In   门诊费用记录.结帐id%Type,
    strSQL = strSQL & mobjChargeInfor.结帐ID & ","
    '  结算序号_In 病人预交记录.结算序号%Type,
    strSQL = strSQL & "Null,"
    '  保险结算_In Varchar2
    strSQL = strSQL & "Null)"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    gcnOracle.CommitTrans: blnTrans = False: mblnCommitBill = True
    If Not mInsurePara.多单据分单据结算 Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicSwap, True, mobjChargeInfor.intInsure)
    zlInsureClinicSwap = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    If blnTrans Then
        '医保和HIS不是同一个事务,HIS事务失败,但医保可能已上传,所以需要调"取消交易"接口
        If blnTransMedicare Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicSwap, False, mobjChargeInfor.intInsure)
    End If
'    If blnTransMedicare = False Then    '如果医保成功了，不删除划价单，费用失败可以重收
'        Call DelMedicareTempNO(False, strBillNO)
'    End If
    Call SaveErrLog
End Function

Private Sub DelMedicareTempNO(ByVal blnPriceSaved As Boolean, ByVal strBillNO As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医保直接收费时,删除前一个事务提交的划价单
    '编制:刘兴洪
    '日期:2014-06-06 18:20:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If Not blnPriceSaved Then Exit Sub
    
    gstrSQL = "zl_门诊划价记录_DELETE('" & strBillNO & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsBlance_GotFocus()
    Clear预交款
End Sub

Private Sub Clear预交款()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除预交款
    '编制:刘兴洪
    '日期:2014-08-07 15:22:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not txt冲预交.Enabled Then Exit Sub
    If Not txt冲预交.Visible Then Exit Sub
    If Val(lbl冲预交.Tag) = 1 Then Exit Sub
    If Val(txt冲预交) = 0 Then Exit Sub
    txt冲预交.Text = ""
    txt缴款.Text = ""
End Sub

Private Function GetPayCard(ByVal lngCardTypeID As Long, ByVal bln消费卡 As Boolean, Optional bln仅启用 As Boolean = True) As Card
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取卡类别ID
    '入参:lngCardTypeID-卡类别ID
    '返回:返回Card对象
    '编制:刘兴洪
    '日期:2014-07-31 15:11:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    On Error GoTo errHandle
    If Not gobjSquare.objSquareCard Is Nothing Then
        'zlGetCard:(ByVal lngCardTypeID As Long, ByVal bln消费卡 As Boolean, ByRef objCard As Card)
        If gobjSquare.objSquareCard.zlGetCard(lngCardTypeID, bln消费卡, objCard) = False Then Exit Function
        Set GetPayCard = objCard
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
