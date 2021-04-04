VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Begin VB.Form frmFeeRefundment 
   Caption         =   "门诊费用转住院费用-->退费"
   ClientHeight    =   6285
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   10860
   Icon            =   "frmFeeRefundment.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   10860
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picFilter 
      BorderStyle     =   0  'None
      Height          =   492
      Left            =   216
      ScaleHeight     =   495
      ScaleWidth      =   10860
      TabIndex        =   5
      Top             =   648
      Width           =   10860
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   360
         Left            =   480
         TabIndex        =   17
         Top             =   75
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   635
         Appearance      =   2
         IDKindStr       =   $"frmFeeRefundment.frx":058A
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
         NotContainFastKey=   "F1;CTRL+F1;F5;CTRL+A;CTRL+C"
         BackColor       =   -2147483633
      End
      Begin VB.TextBox txtPatient 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   1170
         MaxLength       =   100
         TabIndex        =   9
         Top             =   72
         Width           =   2040
      End
      Begin VB.TextBox txtSex 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
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
         Left            =   3780
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   72
         Width           =   600
      End
      Begin VB.TextBox txtOld 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
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
         Left            =   5100
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   72
         Width           =   585
      End
      Begin VB.TextBox txt住院号 
         BackColor       =   &H00E0E0E0&
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
         Left            =   6555
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   72
         Width           =   1815
      End
      Begin VB.Label lblPatient 
         AutoSize        =   -1  'True
         Caption         =   "病人"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   240
         Index           =   7
         Left            =   15
         TabIndex        =   13
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblSex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3252
         TabIndex        =   12
         Top             =   132
         Width           =   480
      End
      Begin VB.Label lblOld 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4608
         TabIndex        =   11
         Top             =   132
         Width           =   480
      End
      Begin VB.Label lbl住院号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5832
         TabIndex        =   10
         Top             =   132
         Width           =   720
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   5928
      Width           =   10860
      _ExtentX        =   19156
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmFeeRefundment.frx":0653
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14076
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin VB.PictureBox picMzToZy 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3705
      Left            =   1095
      ScaleHeight     =   3705
      ScaleWidth      =   7785
      TabIndex        =   1
      Top             =   2025
      Width           =   7788
      Begin VB.PictureBox picBack 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3465
         ScaleHeight     =   375
         ScaleWidth      =   3840
         TabIndex        =   18
         Top             =   1950
         Visible         =   0   'False
         Width           =   3840
         Begin VB.ComboBox cboStyle 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   360
            Left            =   615
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   0
            Width           =   1710
         End
         Begin VB.TextBox txtSum 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2295
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   0
            Width           =   1530
         End
         Begin VB.Label lblBack 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "退款"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   30
            TabIndex        =   21
            Top             =   60
            Width           =   480
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsFee 
         Height          =   2505
         Left            =   120
         TabIndex        =   2
         Top             =   30
         Width           =   5625
         _cx             =   9922
         _cy             =   4419
         Appearance      =   1
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483628
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
         Rows            =   10
         Cols            =   0
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmFeeRefundment.frx":0EE7
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
         ExplorerBar     =   2
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
      Begin VSFlex8Ctl.VSFlexGrid vsBalance 
         Height          =   735
         Left            =   120
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   2880
         Width           =   11160
         _cx             =   19685
         _cy             =   1296
         Appearance      =   0
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
         BackColor       =   -2147483633
         ForeColor       =   -2147483640
         BackColorFixed  =   8421504
         ForeColorFixed  =   16777215
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483633
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   3
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   8
         FixedRows       =   0
         FixedCols       =   1
         RowHeightMin    =   360
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmFeeRefundment.frx":0EFD
         ScrollTrack     =   -1  'True
         ScrollBars      =   0
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
         ExplorerBar     =   3
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
      Begin VB.Label lblSum 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "当前转出合计"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   2610
         Width           =   1350
      End
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1932
      Left            =   660
      ScaleHeight     =   1935
      ScaleWidth      =   3750
      TabIndex        =   3
      Top             =   1692
      Width           =   3756
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   4992
         Left            =   492
         TabIndex        =   4
         Top             =   36
         Width           =   9516
         _Version        =   589884
         _ExtentX        =   16775
         _ExtentY        =   8811
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picHistory 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2904
      Left            =   3240
      ScaleHeight     =   2910
      ScaleWidth      =   5910
      TabIndex        =   14
      Top             =   1212
      Width           =   5904
      Begin VSFlex8Ctl.VSFlexGrid vsHistory 
         Height          =   2208
         Left            =   108
         TabIndex        =   15
         Top             =   348
         Width           =   5628
         _cx             =   9927
         _cy             =   3895
         Appearance      =   1
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483628
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
         Rows            =   10
         Cols            =   0
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmFeeRefundment.frx":0FC8
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
         ExplorerBar     =   2
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
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   -15
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmFeeRefundment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
'---------------------------------------------------------------------------------------------------------------------------------------------
'功能:对门诊记帐费用和收费费用进行销帐或退费处理
'编制:刘兴洪
'日期:2011-03-01 14:29:10
'调用者:门诊收费(转住院费用退费);病人结帐管理(转住院费用销帐)
'问题:36076
'---------------------------------------------------------------------------------------------------------------------------------------------
Private mlngModule As Long, mstrPrivs As String
Private mcbrControl As CommandBarControl, mcbrMenuBar As CommandBarPopup, mcbrToolBar As CommandBar
Private mblnChange As Boolean
Private mblnFirst As Boolean
Private mrsFeeList As ADODB.Recordset
Private mrsHistoryList As ADODB.Recordset
Private mrsBalance As ADODB.Recordset, mrsBalanceBak As ADODB.Recordset
Private mblnNotClick As Boolean
Private mstr标志 As String   '退费;销帐
Private mintSucces As Integer  '退费成功次数
Private mlng病人ID As Long, mint性质 As Integer '1-收费;2-记帐
Private mblnSel As Boolean  '是否已经选择了相关的单据
Private mlngShareUseID As Long
Private mcur误差 As Currency
Private WithEvents mobjIDCard As zlIDCard.clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mrsInfo As ADODB.Recordset
Private mblnValid As Boolean
Private mrsBalanceDup As ADODB.Recordset
Private mstrStyle As String
Private mblnMultiBalance As Boolean
Private mcur合计 As Currency
Private mstrUsedBills As String
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private mobjSquare As Object
 
Private Enum 医院业务
    support门诊结算作废 = 33        '医保是否支持门诊结算作废，不支持只有个人帐帐户原样退,其余的医保结算方式退为现金,支持的再判断每一种结算方式是否允许退回
End Enum
Private Enum mPgIndex
    pg_销帐 = 1
    pg_历史销帐 = 2
End Enum
Private mbln立即销帐 As Boolean
Private mbln门诊转住院先审核 As Boolean
Private mstrFindNO As String '查找单据号
Private mstrFindFpNo As String '查找的发票号
Private mint收费清单 As Integer      '0-不打印,1-要打印,2-选择是否打印
Private mbln药房单位 As Boolean '划价,记帐,收费时是否按照门诊单位进行显示；划价,收费也可能按住院单位
'-----------------------------------------------------------------------------------
'结算卡相关
Private mstrPassWord As String
'-----------------------------------------------------------------------------------

Public Function zlShowEdit(ByVal frmMain As Object, ByVal int性质 As Integer, _
    ByVal lngModuel As Long, ByVal strPrivs As String, _
    Optional ByVal lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入口
    '入参:frmMain-调用的主窗体
    '       int性质-1-收费单;2-记帐单
    '       lng病人ID-对指定病人进行退费
    '出参:
    '返回:只要一次以上退费成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-02-22 16:31:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mintSucces = 0: mstrPrivs = strPrivs: mlngModule = lngModuel
    mlng病人ID = lng病人ID: mint性质 = int性质
    mstr标志 = "退费": If mint性质 = 2 Then mstr标志 = "销帐"
    Me.Caption = IIf(mint性质 = 1, "门诊收费转住院费用-退费管理", "门诊记帐转住院费用-销帐管理")
    
    If frmMain Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
    zlShowEdit = mintSucces > 0
End Function

Private Sub cboStyle_Change()
    Call SetBlanceShow
End Sub

Private Function IsYBSingle(ByVal strNO As String, ByVal intInsure As Integer) As Boolean
    Dim strSQL As String, rsTmp As ADODB.Recordset, blnInsureSingle As Boolean
    
    blnInsureSingle = gclsInsure.GetCapability(83, , intInsure)
    If blnInsureSingle = False Then
        IsYBSingle = False
        Exit Function
    Else
        strSQL = "Select 1 From 医保结算明细 Where NO = [1] And Rownum < 2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
        If rsTmp.EOF Then
            IsYBSingle = False
        Else
            If CheckAllTurn(strNO) Then
                IsYBSingle = False
            Else
                IsYBSingle = True
            End If
        End If
    End If
    
End Function

Private Sub ClsAllNO()
   Dim i As Long
    With vsFee
        If .ColIndex("单据号") >= 0 Then
            For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, .ColIndex("单据号")) <> "" Then
                    .TextMatrix(i, .ColIndex(mstr标志)) = 0
                End If
            Next
            Call CalcSUMMony
            Call SetBlanceShow
            mblnSel = False
        End If
    End With
End Sub
Private Sub SelAllNO()
    Dim i As Long
    With vsFee
        If .ColIndex("单据号") >= 0 Then
            For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, .ColIndex("单据号")) <> "" Then
                    .TextMatrix(i, .ColIndex(mstr标志)) = -1
                End If
            Next
            Call CalcSUMMony
            Call SetBlanceShow
            mblnSel = True
        End If
    End With
End Sub

Private Sub zlSaveData()
    Dim i As Integer
    If SaveData = False Then
        stbThis.Panels(2).Text = IIf(mint性质 = 1, "退费失败!", "销帐失败!")
        Exit Sub
    End If
    mstrFindNO = "": mstrFindFpNo = ""
    Call ReadListData
    Call ReadHistoryListData
    If vsFee.TextMatrix(1, vsFee.ColIndex("单据")) = "" Then
        picBack.Visible = False
        For i = 1 To vsBalance.Cols - 1
            vsBalance.TextMatrix(0, i) = ""
        Next i
    End If
    mblnChange = False
    stbThis.Panels(2).Text = IIf(mint性质 = 1, "退费成功!", "销帐成功!")
End Sub

 
Private Sub cboStyle_Click()
    Call SetBlanceShow
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
            Call txtPatient_KeyPress(vbKeyReturn)
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
    If txtPatient.Text <> "" Then Call txtPatient_KeyPress(vbKeyReturn)
End Sub
Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0
    Set gobjSquare.objCurCard = objCard
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub

 

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If txtPatient.Locked Then Exit Sub
    txtPatient.Text = objPatiInfor.卡号
    If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
End Sub


Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNo As String)
    If txtPatient.Text <> "" Or txtPatient.Locked Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objCard As Card
    '54896
    Set objCard = IDKind.GetIDKindCard("IC卡", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = strCardNo
    If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthday As Date, ByVal strAddress As String)
    If txtPatient.Text <> "" Or txtPatient.Locked Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("身份证", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = strID
    If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '------------------------------------
    Select Case Control.ID
    'bytMode=1 打印;2 预览;3 输出到EXCEL
    Case conMenu_File_Preview: Call zlRptPrint(2)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_File_Exit: Unload Me
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_Edit_ReBillingButton  '销帐
            Call zlSaveData
    Case conMenu_View_StatusBar
        stbThis.Visible = Not stbThis.Visible
        cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Button
        cbsThis(2).Visible = Not cbsThis(2).Visible
        cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each mcbrControl In cbsThis(2).Controls
            mcbrControl.Style = IIf(mcbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        cbsThis.Options.LargeIcons = Not cbsThis.Options.LargeIcons
        cbsThis.RecalcLayout
    Case conMenu_Edit_SelAll    '全选
            Call SelAllNO
    Case conMenu_Edit_ClsAll    '全清
            Call ClsAllNO
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case Else
        If (Control.ID >= conMenu_ReportPopup * 100# + 1 And Control.ID <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
                Call zlCallCustomReprot(Val(Split(Control.Parameter, ",")(0)), Trim(Split(Control.Parameter, ",")(1)))
        End If
    End Select
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbsThis_Resize()
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    
    Err = 0: On Error Resume Next
    Call cbsThis.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    With picFilter
        .Top = lngTop
        .Left = lngLeft
        .Width = lngRight - 100
    End With
    With picList
        .Left = lngLeft + 50: .Top = picFilter.Top + picFilter.Height
        .Width = lngRight - 100
        .Height = lngBottom - .Top
    End With
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnHaveData As Boolean, lngID As Long, blnEnabled As Boolean
    Dim i As Integer
    If Me.Visible = False Then Exit Sub
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        If Val(Me.tbPage.Selected.Tag) = mPgIndex.pg_销帐 Then
            Control.Enabled = Trim(vsFee.TextMatrix(1, vsFee.ColIndex("单据号"))) <> ""
        Else
            Control.Enabled = Trim(vsHistory.TextMatrix(1, vsHistory.ColIndex("单据号"))) <> ""
        End If
    Case conMenu_Edit_ReBillingButton ' 退费
        With vsFee
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("选择")) <> "" Then
                    mblnSel = True
                    Exit For
                Else
                    mblnSel = False
                End If
            Next i
        End With
        Control.Enabled = mblnSel And Val(Me.tbPage.Selected.Tag) = mPgIndex.pg_销帐
    Case conMenu_Edit_SelAll, conMenu_Edit_ClsAll    '全选
            Control.Enabled = Val(Me.tbPage.Selected.Tag) = mPgIndex.pg_销帐
    Case conMenu_View_Refresh
    End Select
End Sub

Private Sub InitPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化页面控件
    '编制:刘兴洪
    '日期:2011-01-25 15:22:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As TabControlItem, objForm As Object
    Err = 0: On Error GoTo ErrHand:
    Set objItem = tbPage.InsertItem(mPgIndex.pg_销帐, IIf(mint性质 = 1, "退费处理", "销帐处理"), picMzToZy.hWnd, 0)
    objItem.Tag = mPgIndex.pg_销帐
    Set objItem = tbPage.InsertItem(mPgIndex.pg_历史销帐, IIf(mint性质 = 1, "历史退费记录", "历史销帐记录"), picHistory.hWnd, 0)
    objItem.Tag = mPgIndex.pg_历史销帐
     With tbPage
        tbPage.Item(0).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = True
        .PaintManager.ClientFrame = xtpTabFrameBorder
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Err = 0: On Error Resume Next
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
    stbThis.Top = Me.ScaleHeight - Me.stbThis.Height
End Sub
Private Sub Form_Activate()
    Dim strKey As String
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    mblnChange = False
End Sub

Private Sub Form_Load()
    Dim i As Long
    mbln立即销帐 = Val(zlDatabase.GetPara("费用转出立即退费", glngSys, 1131)) = 1
    mbln门诊转住院先审核 = IIf(Val(zlDatabase.GetPara("门诊转住院先审核", glngSys, 1143, 0)) = 1, True, False)
    mint收费清单 = 0: mbln药房单位 = False
    If mint性质 = 1 Then
        mint收费清单 = Val(zlDatabase.GetPara("收费清单打印方式", glngSys, 1121))   '门诊收费
        mbln药房单位 = zlDatabase.GetPara("药品单位", glngSys, 1121) = "1"
    End If
    mblnSel = False
    RestoreWinState Me, App.ProductName
    mstrPrivs = gstrPrivs: mlngModule = glngModul
    If mint性质 = 1 Then
        mlngShareUseID = Val(zlDatabase.GetPara("共用收费票据批次", glngSys, mlngModule, "0"))
        IDKind.IDKindStr = "姓|姓名|0;医|医保号|0;身|身份证号|0;IC|IC卡号|1;门|门诊号|0;住|住院号|0;就|就诊卡|0;单|单据号|0;发|发票号|0"
    Else
        IDKind.IDKindStr = "姓|姓名|0;医|医保号|0;身|身份证号|0;IC|IC卡号|1;门|门诊号|0;住|住院号|0;就|就诊卡|0;单|单据号|0"
        mlngShareUseID = 0
    End If
    Call initCardSquareData
    IDKind.IDKind = Val(zlDatabase.GetPara("门诊转住院IDKIND", glngSys, mlngModule, "0"))
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, False)
    Call zlDefCommandBars
    Call InitPage
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    Set mrsInfo = New ADODB.Recordset
    vsFee.OwnerDraw = flexODContent
    '多张单据使用多种结算方式模式
    mblnMultiBalance = zlDatabase.GetPara(79, glngSys) = "1"
    Call zlCreateObject
    Call LoadStyle
 End Sub

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
'    If mblnChange Then
'        If MsgBox("注意:" & vbCrLf & "    你修改了数据,但你还未保存,是否真的要退出?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
'            Cancel = 1: Exit Sub
'        End If
'    End If
    zlDatabase.SetPara "门诊转住院IDKIND", IDKind.IDKind, glngSys, mlngModule, InStr(1, mstrPrivs, ";参数设置;") > 0
    zl_vsGrid_Para_Save mlngModule, vsFee, Me.Caption, IIf(mint性质 = 1, "退费列表", "销帐列表"), True
    zl_vsGrid_Para_Save mlngModule, vsHistory, Me.Caption, IIf(mint性质 = 1, "历史退费列表", "历史销帐列表"), True
    
    SaveWinState Me, App.ProductName
    Call zlCloseObject
     
    Set mrsFeeList = Nothing
    Set mrsInfo = Nothing
    Set mrsHistoryList = Nothing
    Set mrsBalance = Nothing
    
End Sub
Private Sub picList_Resize()
    Err = 0: On Error Resume Next
    With picList
        tbPage.Left = .ScaleLeft
        tbPage.Top = .ScaleTop
        tbPage.Width = .ScaleWidth
        tbPage.Height = .ScaleHeight
    End With
End Sub

Private Sub picMzToZy_Resize()
    Err = 0: On Error Resume Next
    With picMzToZy
        vsFee.Top = .ScaleTop + 100
        vsFee.Width = .ScaleWidth - vsFee.Left * 2
        'cmdOk.Top = .ScaleHeight - cmdOk.Height - 50
        'cmdOk.Left = .ScaleWidth - cmdOk.Width - vsFee.Left * 2
        vsBalance.Left = vsFee.Left
        vsBalance.Width = IIf(picBack.Visible, vsFee.Width - 3000, vsFee.Width)
        picBack.Left = vsFee.Width - 4000
        vsBalance.Top = .ScaleHeight - vsBalance.Height - 100 - picBack.Height
        picBack.Top = vsBalance.Top + vsBalance.Height + 45
        lblSum.Top = IIf(vsBalance.Visible, vsBalance.Top, .ScaleHeight - stbThis.Height) - lblSum.Height - 20
        
        vsFee.Height = lblSum.Top - vsFee.Top
        'cmdAllCls.Top = .ScaleHeight - cmdAllCls.Height - 50
        'cmdAllSel.Top = cmdAllCls.Top
        'cmdOk.Top = cmdAllCls.Top
    End With
End Sub

Private Function LoadStyle() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    
    cboStyle.Clear
    On Error GoTo errH
    Set rsTmp = Get结算方式("收费", "1,2")
    For i = 1 To rsTmp.RecordCount
        If InStr(",1,2,", "," & rsTmp!性质 & ",") > 0 And Val(Nvl(rsTmp!应付款)) = 0 Then
            cboStyle.AddItem rsTmp!名称
            cboStyle.ItemData(cboStyle.NewIndex) = rsTmp!性质
            If rsTmp!缺省 = 1 And cboStyle.ListIndex = -1 Then
                Call zlControl.CboSetIndex(cboStyle.hWnd, cboStyle.NewIndex)
            End If
        End If
        rsTmp.MoveNext
    Next
    If cboStyle.ListIndex = -1 And cboStyle.ListCount > 0 Then Call zlControl.CboSetIndex(cboStyle.hWnd, 0)
    txtSum.ForeColor = vbRed
    strSQL = "" & _
            " Select B.编码,B.名称,Nvl(B.缺省标志,0) as 缺省,Nvl(B.性质,1) as 性质,Nvl(B.应付款,0) as 应付款" & _
            " From 结算方式应用 A,结算方式 B" & _
            " Where A.应用场合=[1] And B.名称=A.结算方式 " & _
            " And B.性质<>8 " & _
            " Order by 性质,lpad(编码,3,' ')"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "收费")
    For i = 1 To rsTmp.RecordCount
        If InStr(",1,2,7,", "," & rsTmp!性质 & ",") > 0 Then
            mstrStyle = mstrStyle & rsTmp!名称 & ":"
        End If
        rsTmp.MoveNext
    Next
    LoadStyle = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SetPicBack(ByVal strNos As String) As Boolean
    vsBalance.Width = vsFee.Width
    'picBack.Left = vsBalance.Width + vsBalance.Left + 30
    picBack.Visible = True
    SetPicBack = True
End Function

Private Sub picHistory_Resize()
    Err = 0: On Error Resume Next
    With picHistory
        vsHistory.Top = 100: vsHistory.Left = .ScaleLeft + 50
        vsHistory.Width = .ScaleWidth - vsHistory.Left * 2
        vsHistory.Height = .ScaleHeight - vsHistory.Top - 100
    End With
End Sub

Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
   If Val(tbPage.Selected.Tag) = mPgIndex.pg_销帐 Then
        If vsFee.Enabled And vsFee.Visible Then vsFee.SetFocus
    Else
        Exit Sub
    End If
End Sub
Private Function InitBlanceData(ByVal strNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化结算数据
    '入参:strNos-指定的单据号,以逗号分离:'A0001,A0002
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2011-02-23 14:45:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Err = 0: On Error GoTo errHandle
    If mint性质 = 2 Then
        InitBlanceData = True
        Exit Function
    End If
    If strNos = "" Then InitBlanceData = True: Exit Function
    strSQL = _
    "Select Distinct 结帐id" & vbNewLine & _
    "From 门诊费用记录" & vbNewLine & _
    "Where NO In (Select Distinct NO" & vbNewLine & _
    "             From 门诊费用记录" & vbNewLine & _
    "             Where 结帐id In (Select 结帐id" & vbNewLine & _
    "                            From 病人预交记录" & vbNewLine & _
    "                            Where 结算序号 In (Select b.结算序号" & vbNewLine & _
    "                                           From 门诊费用记录 A, 病人预交记录 B" & vbNewLine & _
    "                                           Where a.No In (Select Column_Value From Table(f_Str2list([1]))) And" & vbNewLine & _
    "                                                 Mod(a.记录性质, 10) = 1 And a.记录状态 <> 0 And a.结帐id = b.结帐id))) And" & vbNewLine & _
    "      Mod(记录性质, 10) = 1 And 记录状态 <> 0"

    strSQL = _
    " Select /*+ rule */ A.结算方式,Nvl(B.性质,1) as 性质,B.应付款,A.金额,A.结算号码" & _
    " From (  Select Decode(A.记录性质,3,A.结算方式,NULL) as 结算方式,A.结算号码," & _
    "               Sum(A.冲预交) as 金额" & _
    "         From 病人预交记录 A,(" & strSQL & ") B" & _
    "         Where A.结帐ID=B.结帐ID And A.记录性质 IN(1,11,3) And Nvl(A.冲预交,0)<>0" & _
    "         Group by Decode(A.记录性质,3,A.结算方式,NULL),A.结算号码" & _
    "       ) A,结算方式 B " & _
    " Where A.结算方式=B.名称(+) " & _
    " "
    Set mrsBalance = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Replace(strNos, "'", ""))
    Set mrsBalanceBak = mrsBalance
    InitBlanceData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function InitPatialBalance(ByVal strNos As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化部分退费的结算数据
    '入参:strNos-指定的单据号,以逗号分离:'A0001,A0002
    '出参:
    '返回:
    '编制:刘尔旋
    '日期:2014-06-03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, dblSum As Double, i As Integer, strTable As String
    Dim cur金额 As Currency, rsTmp As ADODB.Recordset, rsTx As ADODB.Recordset
    Dim bln退现 As Boolean, dbl医保个帐 As Double, dbl医保基金 As Double
    Dim arrNO() As String
    Dim j As Integer
    Dim lngRow As Long
    Dim curOld As Currency
    Dim curOldTotal As Currency
    Dim strOldNOs As String, strNewNos As String
    Err = 0: On Error GoTo errHandle
    If mint性质 = 2 Then
        InitPatialBalance = 0
        Exit Function
    End If
    If strNos = "" Then InitPatialBalance = 0: Exit Function
    
    Call InitBlanceData(strNos)
    dblSum = 0
    curOld = 0
    curOldTotal = 0
    
    Set mrsBalance = New ADODB.Recordset
    mrsBalance.Fields.Append "结算方式", adVarChar, 20
    mrsBalance.Fields.Append "性质", adBigInt, 2
    mrsBalance.Fields.Append "应付款", adBigInt, 1
    mrsBalance.Fields.Append "金额", adDouble, 30
    mrsBalance.Fields.Append "摘要", adVarChar, 50
    mrsBalance.Fields.Append "结算号码", adVarChar, 30
    mrsBalance.CursorLocation = adUseClient
    mrsBalance.LockType = adLockOptimistic
    mrsBalance.CursorType = adOpenStatic
    mrsBalance.Open
    
    strNos = Replace(strNos, "'", "")
    arrNO = Split(strNos, ",")
    For i = 0 To UBound(arrNO)
        For j = 1 To vsFee.Rows - 1
            If vsFee.TextMatrix(j, vsFee.ColIndex("单据号")) = arrNO(i) Then lngRow = j: Exit For
        Next j
        If CheckAllTurn(arrNO(i)) = True Then
            strOldNOs = strOldNOs & "," & arrNO(i)
        Else
            If Val(vsFee.TextMatrix(lngRow, vsFee.ColIndex("险类"))) <> 0 Then
                If IsYBSingle(vsFee.TextMatrix(lngRow, vsFee.ColIndex("单据号")), Val(vsFee.TextMatrix(lngRow, vsFee.ColIndex("险类")))) = False Then
                    strOldNOs = strOldNOs & "," & arrNO(i)
                Else
                    strNewNos = strNewNos & "," & arrNO(i)
                End If
            Else
                strNewNos = strNewNos & "," & arrNO(i)
            End If
        End If
    Next i
    If strOldNOs <> "" Then
        strOldNOs = Mid(strOldNOs, 2)
        
        strTable = _
        "Select Distinct 结帐id" & vbNewLine & _
        "From 门诊费用记录" & vbNewLine & _
        "Where NO In" & vbNewLine & _
        "      (Select Distinct NO" & vbNewLine & _
        "       From 门诊费用记录" & vbNewLine & _
        "       Where 结帐id In (Select 结帐id" & vbNewLine & _
        "                      From 病人预交记录" & vbNewLine & _
        "                      Where 结算序号 In (Select b.结算序号" & vbNewLine & _
        "                                     From 门诊费用记录 A, 病人预交记录 B" & vbNewLine & _
        "                                     Where a.No In (Select Column_Value From Table(f_Str2list([1]))) And b.结算序号 < 0 And" & vbNewLine & _
        "                                           Mod(a.记录性质, 10) = 1 And a.记录状态 <> 0 And a.结帐id = b.结帐id))) And" & vbNewLine & _
        "      Mod(记录性质, 10) = 1 And 记录状态 <> 0" & vbNewLine & _
        "Union" & vbNewLine & _
        "Select Distinct 结帐id" & vbNewLine & _
        "From 门诊费用记录" & vbNewLine & _
        "Where NO In (Select Distinct NO" & vbNewLine & _
        "             From 门诊费用记录" & vbNewLine & _
        "             Where 结帐id In (Select a.结帐id" & vbNewLine & _
        "                            From 门诊费用记录 A, 病人预交记录 B" & vbNewLine & _
        "                            Where a.No In (Select Column_Value From Table(f_Str2list([1]))) And b.结算序号 > 0 And" & vbNewLine & _
        "                                  Mod(a.记录性质, 10) = 1 And a.记录状态 <> 0 And a.结帐id = b.结帐id))"

        
        strSQL = _
        " Select /*+ rule */ A.结算方式,0 as 性质,Null As 应付款,A.金额,Null As 摘要,Null As 结算号码" & _
        " From (  Select '冲预交' as 结算方式," & _
        "               Sum(A.冲预交) as 金额" & _
        "         From 病人预交记录 A,(" & strTable & ") B" & _
        "         Where A.结帐ID=B.结帐ID And Mod(A.记录性质,10) = 1 " & _
        "       ) A "

        strSQL = strSQL & _
        " Union " & _
        " Select /*+ rule */ A.结算方式,Nvl(B.性质,1) as 性质,B.应付款,A.金额,Null As 摘要,Null As 结算号码" & _
        " From (  Select Decode(A.记录性质,3,A.结算方式,NULL) as 结算方式," & _
        "               Sum(A.冲预交) as 金额" & _
        "         From 病人预交记录 A,(" & strTable & ") B" & _
        "         Where A.结帐ID=B.结帐ID And A.记录性质=3 And Nvl(A.冲预交,0)<>0" & _
        "         Group by Decode(A.记录性质,3,A.结算方式,NULL)" & _
        "       ) A,结算方式 B " & _
        " Where A.结算方式=B.名称 And B.性质 In (3,4)"
        
        strSQL = strSQL & _
        " Union " & _
        " Select /*+ rule */ A.结算方式,Nvl(B.性质,1) as 性质,B.应付款,A.金额,Null As 摘要,Null As 结算号码" & _
        " From (  Select Decode(A.记录性质,3,A.结算方式,NULL) as 结算方式," & _
        "               Sum(A.冲预交) as 金额" & _
        "         From 病人预交记录 A,(" & strTable & ") B" & _
        "         Where A.结帐ID=B.结帐ID And A.记录性质=3 And Nvl(A.冲预交,0)<>0" & _
        "         And Exists (Select 1 From 医疗卡类别 Where ID=A.卡类别ID And 是否退现 = 0)" & _
        "         Group by Decode(A.记录性质,3,A.结算方式,NULL)" & _
        "       ) A,结算方式 B " & _
        " Where A.结算方式=B.名称 And B.性质 In (7,8)"
        
        strSQL = strSQL & _
        " Union " & _
        " Select /*+ rule */ A.结算方式,Nvl(B.性质,1) as 性质,B.应付款,A.金额,Null As 摘要,Null As 结算号码" & _
        " From (  Select Decode(A.记录性质,3,A.结算方式,NULL) as 结算方式," & _
        "               Sum(A.冲预交) as 金额" & _
        "         From 病人预交记录 A,(" & strTable & ") B" & _
        "         Where A.结帐ID=B.结帐ID And A.记录性质=3 And Nvl(A.冲预交,0)<>0" & _
        "         And A.结算卡序号 Is Not Null" & _
        "         Group by Decode(A.记录性质,3,A.结算方式,NULL)" & _
        "       ) A,结算方式 B " & _
        " Where A.结算方式=B.名称 And B.性质 = 8"
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strOldNOs)
        Do While Not rsTmp.EOF
            If Val(Nvl(rsTmp!金额)) <> 0 Then
                With mrsBalance
                    .AddNew
                    !结算方式 = Nvl(rsTmp!结算方式)
                    !性质 = Nvl(rsTmp!性质)
                    !应付款 = "0"
                    !金额 = Val(Nvl(rsTmp!金额))
                    !摘要 = ""
                    !结算号码 = ""
                    .Update
                End With
                curOldTotal = curOldTotal + Val(Nvl(rsTmp!金额))
            End If
            rsTmp.MoveNext
        Loop
    End If
    
    strSQL = "Select Sum(实收金额) As 金额" & vbNewLine & _
            "From 门诊费用记录" & vbNewLine & _
            "Where NO In (Select Column_Value From Table(f_Str2list([1]))) And Mod(记录性质, 10) = 1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strOldNOs)
    If rsTmp.RecordCount <> 0 Then
        curOld = Val(Nvl(rsTmp!金额))
    End If
    
    dblSum = dblSum + curOld - curOldTotal
    
    If strNewNos <> "" Then strNewNos = Mid(strNewNos, 2)
    
    cur金额 = mcur合计 - curOld
    
    strSQL = "Select A.结算方式, Sum(A.金额) As 金额,B.性质" & vbNewLine & _
            "From 医保结算明细 A,结算方式 B" & vbNewLine & _
            "Where A.NO In (Select Column_Value From Table(f_Str2list([1]))) And A.结算方式=B.名称(+)" & vbNewLine & _
            "Group By A.结算方式,B.性质" & vbNewLine & _
            "Having Sum(A.金额) <> 0"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNewNos)
    
    Do While Not rsTmp.EOF
        If Val(Nvl(rsTmp!金额)) <> 0 Then
            If Val(cur金额) > Val(Nvl(rsTmp!金额)) Then
                With mrsBalance
                    .AddNew
                    !结算方式 = Nvl(rsTmp!结算方式)
                    !性质 = Nvl(rsTmp!性质)
                    !应付款 = "0"
                    !金额 = Val(Nvl(rsTmp!金额))
                    !摘要 = ""
                    !结算号码 = ""
                    .Update
                End With
                cur金额 = cur金额 - Val(Nvl(rsTmp!金额))
                If Val(Nvl(rsTmp!性质)) = 3 Then dbl医保个帐 = dbl医保个帐 + Val(Nvl(rsTmp!金额))
                If Val(Nvl(rsTmp!性质)) = 4 Then dbl医保基金 = dbl医保基金 + Val(Nvl(rsTmp!金额))
            Else
                With mrsBalance
                    .AddNew
                    !结算方式 = Nvl(rsTmp!结算方式)
                    !性质 = Nvl(rsTmp!性质)
                    !应付款 = "0"
                    !金额 = cur金额
                    !摘要 = ""
                    !结算号码 = ""
                    .Update
                End With
                InitPatialBalance = Format(dblSum, "0.00")
                Exit Function
            End If
        End If
        rsTmp.MoveNext
    Loop
    

    strTable = _
        "Select Distinct 结帐id" & vbNewLine & _
        "From 门诊费用记录" & vbNewLine & _
        "Where NO In" & vbNewLine & _
        "      (Select Distinct NO" & vbNewLine & _
        "       From 门诊费用记录" & vbNewLine & _
        "       Where 结帐id In (Select 结帐id" & vbNewLine & _
        "                      From 病人预交记录" & vbNewLine & _
        "                      Where 结算序号 In (Select b.结算序号" & vbNewLine & _
        "                                     From 门诊费用记录 A, 病人预交记录 B" & vbNewLine & _
        "                                     Where a.No In (Select Column_Value From Table(f_Str2list([1]))) And b.结算序号 < 0 And" & vbNewLine & _
        "                                           Mod(a.记录性质, 10) = 1 And a.记录状态 <> 0 And a.结帐id = b.结帐id))) And" & vbNewLine & _
        "      Mod(记录性质, 10) = 1 And 记录状态 <> 0" & vbNewLine & _
        "Union" & vbNewLine & _
        "Select Distinct 结帐id" & vbNewLine & _
        "From 门诊费用记录" & vbNewLine & _
        "Where NO In (Select Distinct NO" & vbNewLine & _
        "             From 门诊费用记录" & vbNewLine & _
        "             Where 结帐id In (Select a.结帐id" & vbNewLine & _
        "                            From 门诊费用记录 A, 病人预交记录 B" & vbNewLine & _
        "                            Where a.No In (Select Column_Value From Table(f_Str2list([1]))) And b.结算序号 > 0 And" & vbNewLine & _
        "                                  Mod(a.记录性质, 10) = 1 And a.记录状态 <> 0 And a.结帐id = b.结帐id))"
    
    strSQL = _
    " Select /*+ rule */ '冲预交' as 结算方式," & _
    "               Sum(A.冲预交) as 金额" & _
    "         From 病人预交记录 A,(" & strTable & ") B" & _
    "         Where A.结帐ID=B.结帐ID And A.记录性质 IN(1,11) And Nvl(A.冲预交,0)<>0"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNewNos)
    
    If rsTmp.RecordCount <> 0 Then
        If Val(Nvl(rsTmp!金额)) <> 0 Then
            If Val(cur金额) > Val(Nvl(rsTmp!金额)) Then
                With mrsBalance
                    .AddNew
                    !结算方式 = Nvl(rsTmp!结算方式)
                    !性质 = 0
                    !应付款 = "0"
                    !金额 = Val(Nvl(rsTmp!金额))
                    !摘要 = ""
                    !结算号码 = ""
                    .Update
                End With
                cur金额 = cur金额 - Val(Nvl(rsTmp!金额))
            Else
                With mrsBalance
                    .AddNew
                    !结算方式 = Nvl(rsTmp!结算方式)
                    !性质 = 0
                    !应付款 = "0"
                    !金额 = cur金额
                    !摘要 = ""
                    !结算号码 = ""
                    .Update
                End With
                InitPatialBalance = Format(dblSum, "0.00")
                Exit Function
            End If
        End If
    End If
    
    strSQL = "Select a.结算方式, Sum(a.冲预交) As 金额, a.卡类别id, a.结算卡序号, a.卡号, Min(a.交易流水号) As 交易流水号," & vbNewLine & _
            "                        Min(a.交易说明) As 交易说明, Min(a.合作单位) As 合作单位, b.性质" & vbNewLine & _
            "                 From 病人预交记录 A, 结算方式 B" & vbNewLine & _
            "                 Where a.记录性质 = 3 And a.结帐id In (" & strTable & ") And a.结算方式 = b.名称 And" & vbNewLine & _
            "                       b.性质 In (1, 2, 7, 8)" & vbNewLine & _
            "                 Group By a.结算方式, a.卡类别id, a.结算卡序号, a.卡号, b.性质" & vbNewLine & _
            "                 Having Sum(a.冲预交) <> 0" & vbNewLine & _
            "                 Order By 卡类别id,性质 Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNewNos)
    
    Do While Not rsTmp.EOF
        If Val(Nvl(rsTmp!性质)) = 7 Or (Val(Nvl(rsTmp!性质)) = 8 And Not IsNull(rsTmp!卡类别ID)) Then
            strSQL = "Select 1 from 医疗卡类别 Where id = [1] And 是否退现 = 1"
            Set rsTx = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Nvl(rsTmp!卡类别ID)))
            bln退现 = Not rsTx.EOF
            If bln退现 Then
                If Val(cur金额) > Val(Nvl(rsTmp!金额)) Then
                    dblSum = dblSum + Val(Nvl(rsTmp!金额))
                    cur金额 = cur金额 - Val(Nvl(rsTmp!金额))
                Else
                    dblSum = dblSum + cur金额
                    InitPatialBalance = Format(dblSum, "0.00")
                    Exit Function
                End If
            Else
                If Val(Nvl(rsTmp!金额)) <> 0 Then
                    If Val(cur金额) > Val(Nvl(rsTmp!金额)) Then
                        With mrsBalance
                            .AddNew
                            !结算方式 = Nvl(rsTmp!结算方式)
                            !性质 = Nvl(rsTmp!性质)
                            !应付款 = "0"
                            !金额 = Val(Nvl(rsTmp!金额))
                            !摘要 = ""
                            !结算号码 = ""
                            .Update
                        End With
                        cur金额 = cur金额 - Val(Nvl(rsTmp!金额))
                    Else
                        With mrsBalance
                            .AddNew
                            !结算方式 = Nvl(rsTmp!结算方式)
                            !性质 = Nvl(rsTmp!性质)
                            !应付款 = "0"
                            !金额 = cur金额
                            !摘要 = ""
                            !结算号码 = ""
                            .Update
                        End With
                        InitPatialBalance = Format(dblSum, "0.00")
                        Exit Function
                    End If
                End If
            End If
        Else
            If Val(Nvl(rsTmp!性质)) = 8 Then
                If Val(Nvl(rsTmp!金额)) <> 0 Then
                    If Val(cur金额) > Val(Nvl(rsTmp!金额)) Then
                        With mrsBalance
                            .AddNew
                            !结算方式 = Nvl(rsTmp!结算方式)
                            !性质 = Nvl(rsTmp!性质)
                            !应付款 = "0"
                            !金额 = Val(Nvl(rsTmp!金额))
                            !摘要 = ""
                            !结算号码 = ""
                            .Update
                        End With
                        cur金额 = cur金额 - Val(Nvl(rsTmp!金额))
                    Else
                        With mrsBalance
                            .AddNew
                            !结算方式 = Nvl(rsTmp!结算方式)
                            !性质 = Nvl(rsTmp!性质)
                            !应付款 = "0"
                            !金额 = cur金额
                            !摘要 = ""
                            !结算号码 = ""
                            .Update
                        End With
                        InitPatialBalance = Format(dblSum, "0.00")
                        Exit Function
                    End If
                End If
            Else
                If Val(cur金额) > Val(Nvl(rsTmp!金额)) Then
                    dblSum = dblSum + Val(Nvl(rsTmp!金额))
                    cur金额 = cur金额 - Val(Nvl(rsTmp!金额))
                Else
                    dblSum = dblSum + cur金额
                    InitPatialBalance = Format(dblSum, "0.00")
                    Exit Function
                End If
            End If
        End If
        rsTmp.MoveNext
    Loop
    

    If mrsBalance.RecordCount <> 0 Then mrsBalance.MoveFirst
    
    InitPatialBalance = Format(dblSum, "0.00")
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlDefCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化菜单及工具栏
    '返回:设置成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-01-25 15:29:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto

    cbsThis.VisualTheme = xtpThemeOffice2003
    With cbsThis.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
    End With
    cbsThis.EnableCustomization False
    '-----------------------------------------------------
    '菜单定义
    cbsThis.ActiveMenuBar.Title = "菜单"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop Or xtpFlagHideWrap Or xtpFlagStretched)
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    mcbrMenuBar.ID = conMenu_FilePopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): mcbrControl.BeginGroup = True
    End With
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    mcbrMenuBar.ID = conMenu_EditPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_SelAll, "全选(&A)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ClsAll, "全清(&C)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ReBillingButton, IIf(mint性质 = 1, "退费(&X)", "销帐(&X)")): mcbrControl.BeginGroup = True
    End With
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    mcbrMenuBar.ID = conMenu_ViewPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): mcbrControl.BeginGroup = True
    End With
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    mcbrMenuBar.ID = conMenu_HelpPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set mcbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): mcbrControl.BeginGroup = True
    End With
    
    '快键绑定
    With cbsThis.KeyBindings
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        .Add FCONTROL, vbKeyA, conMenu_Edit_SelAll
        .Add FCONTROL, vbKeyC, conMenu_Edit_ClsAll
      End With
    
    '设置不常用菜单
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    
    '-----------------------------------------------------
    '工具栏定义
    Set mcbrToolBar = cbsThis.Add("工具栏", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.ContextMenuPresent = False
    mcbrToolBar.EnableDocking xtpFlagStretched
    With mcbrToolBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_SelAll, "全选"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ClsAll, "全清")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ReBillingButton, IIf(mint性质 = 1, "退费", "销帐"))
        mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each mcbrControl In mcbrToolBar.Controls
        mcbrControl.Style = xtpButtonIconAndCaption
    Next
     zlDefCommandBars = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Public Sub zlRptPrint(ByVal bytFunc As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:进行打印,预览和输出到EXCEL
    '入参:bytMode=1 打印;2 预览;3 输出到EXCEL
    '编制:刘兴洪
    '日期:2011-01-25 15:14:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCol As Long, objPrint As Object, objRow As New zlTabAppRow, bytPrn As Byte
    Dim vsGrid As VSFlexGrid, rsTemp As New ADODB.Recordset, strSQL As String
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    If mint性质 = 1 Then
        objPrint.Title.Text = gstrUnitName & "门诊转住院退费清册"
    Else
        objPrint.Title.Text = gstrUnitName & "门诊转住院记帐清册"
    End If
    
    objRow.Add "病人：" & txtPatient.Text
    objRow.Add "性别：" & txtSex.Text
    objRow.Add "年龄：" & txtOld.Text
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "打印人:" & UserInfo.姓名
    objRow.Add "打印日期:" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    If Val(Me.tbPage.Selected.Tag) = mPgIndex.pg_销帐 Then
        Set vsGrid = vsFee
    Else
        Set vsGrid = vsHistory
    End If
    
    Err = 0: On Error GoTo ErrHand:
    With vsGrid
        .Redraw = flexRDNone
        For intCol = 0 To .Cols - 1
            .Cell(flexcpData, 0, intCol) = .ColWidth(intCol)
            If .ColHidden(intCol) Or intCol = .ColIndex(mstr标志) Then .ColWidth(intCol) = 0
        Next
    End With
    
    Set objPrint.Body = vsGrid
    If bytFunc = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytPrn
    End If
    With vsGrid
        .Redraw = flexRDNone
        For intCol = 0 To .Cols - 1
            .ColWidth(intCol) = Val(.Cell(flexcpData, 0, intCol))
        Next
        .Redraw = flexRDBuffered
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    With vsGrid
        .Redraw = flexRDNone
        For intCol = 0 To .Cols - 1
            .ColWidth(intCol) = Val(.Cell(flexcpData, 0, intCol))
        Next
        .Redraw = flexRDBuffered
    End With
End Sub

Private Sub zlCallCustomReprot(ByVal lngSys As Long, strReprotName As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用相关的自定义报表
    '编制:刘兴洪
    '日期:2011-01-25 15:16:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNO As Variant, lng结帐ID As Long
    Dim vsGrid As VSFlexGrid
    If Val(Me.tbPage.Selected.Tag) = mPgIndex.pg_销帐 Then
        Set vsGrid = vsFee
    Else
        Set vsGrid = vsHistory
    End If
    With vsGrid
        If .Row > 0 Then
            strNO = Trim(.TextMatrix(.Row, .ColIndex("单据号")))
        End If
        If strNO <> "" Then
            Call ReportOpen(gcnOracle, lngSys, strReprotName, Me, "NO=" & strNO)
        Else
            Call ReportOpen(gcnOracle, lngSys, strReprotName, Me)
        End If
    End With
End Sub

Private Sub txtPatient_Change()
    txtPatient.Tag = ""
    If txtPatient.Locked Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
    IDKind.SetAutoReadCard (txtPatient.Text = "")
    stbThis.Panels(2).Text = ""
End Sub

Private Sub txtPatient_GotFocus()
    zlControl.TxtSelAll txtPatient
    If txtPatient.Locked Then Exit Sub
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtPatient.Text = "")
    Call IDKind.SetAutoReadCard(txtPatient.Text = "")

End Sub

Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtPatient.Locked Or txtPatient.Enabled = False Or txtPatient.Text <> "" Then Exit Sub
    If IDKind.ActiveFastKey = True Then Exit Sub
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim lngID As Long, lngUnit As Long, i As Long
    Dim rsTmp As ADODB.Recordset, strInfo As String
    Dim strSQL As String, curTotal As Currency
    Dim blnCard As Boolean, blnICCard As Boolean
    
    On Error GoTo errH
    
    If txtPatient.Locked Then Exit Sub
    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    
    If Len(Trim(Me.txtPatient.Text)) = 0 And KeyAscii = 13 Then
    Else
        If IDKind.GetCurCard.名称 Like "姓名*" Then
            blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
        ElseIf IDKind.GetCurCard.名称 = "门诊号" Or IDKind.GetCurCard.名称 = "住院号" Then
            If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
                If InStr("0123456789-*+", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
            End If
         Else
            txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
            '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
            txtPatient.IMEMode = 0
        End If
    End If
    
    If blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtPatient.Text) <> "" Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        ElseIf IsNumeric(txtPatient.Tag) And mrsInfo.State = 1 Then
            KeyAscii = 0
            '刷新病人信息:"-病人ID"
            Call GetPatient(IDKind.GetCurCard, txtPatient.Tag, False)
            If mrsInfo.State = 0 Then   '
                txtPatient.Text = "": txtOld.Text = ""
                txtSex.Text = "": txt住院号.Text = ""
                Exit Sub
            End If
            Call ReadListData
            Call ReadHistoryListData
            Exit Sub
        End If
        KeyAscii = 0
        Call FindPati(IDKind.GetCurCard, blnCard, txtPatient.Text)
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
    Dim lngID As Long, lngUnit As Long, i As Long
    Dim rsTmp As ADODB.Recordset, strInfo As String
    Dim strSQL As String, curTotal As Currency
    Dim blnMsg As Boolean, blnICCard As Boolean, blnIDCard As Boolean
 
    '54899
    If objCard.名称 Like "IC卡*" And objCard.系统 = True And mstrPassWord <> "" Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    If objCard.名称 Like "*身份证*" And objCard.系统 = True And mstrPassWord <> "" Then blnIDCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
        
    If Not GetPatient(objCard, strInput, blnCard, blnMsg) Then
        If blnCard Then
            If Not blnMsg Then MsgBox "不能确定病人信息，请检查是否正确刷卡！", vbInformation, gstrSysName
            txtPatient.Text = "": txtOld.Text = ""
            txt住院号.Text = ""
            vsFee.Clear 1: vsFee.Rows = 2
            vsHistory.Clear 1: vsHistory.Rows = 2
            Exit Sub
        End If
        If Not blnMsg Then MsgBox "不能读取病人信息！", vbInformation, gstrSysName
        zlControl.TxtSelAll txtPatient
        txtOld.Text = "": txtSex.Text = "": txt住院号.Text = ""
        vsFee.Clear 1: vsFee.Rows = 2
        vsHistory.Clear 1: vsHistory.Rows = 2
        Exit Sub
    End If
    
    '读取成功
    '就诊卡密码检查
     If (objCard.名称 Like "IC卡*" Or objCard.名称 Like "*身份证*") And objCard.系统 = True And blnCard Then blnCard = False
    If Mid(gstrCardPass, 6, 1) = "1" And (blnCard Or blnICCard Or blnIDCard) Then
        If Not zlCommFun.VerifyPassWord(Me, mstrPassWord, mrsInfo!姓名, mrsInfo!性别, "" & mrsInfo!年龄) Then
            vsFee.Clear 1: vsFee.Rows = 2
            vsHistory.Clear 1: vsHistory.Rows = 2
            Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": txtPatient.SetFocus: Exit Sub
        End If
    End If
    Call ReadListData
    Call ReadHistoryListData
 
End Sub

Private Sub txtPatient_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
End Sub

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
Private Sub txtPatient_Validate(Cancel As Boolean)
    If IsNumeric(txtPatient.Tag) And mrsInfo.State = 1 Then
        mblnValid = True
        Call txtPatient_KeyPress(13)
        mblnValid = False
    End If
End Sub
Private Sub zlClearPatiInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检除病人信息
    '编制:刘兴洪
    '日期:2011-02-23 09:39:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    txtPatient.Text = "": txtOld.Text = "": txtSex.Text = ""
    txt住院号.Text = "": Set mrsInfo = New ADODB.Recordset
    vsFee.Clear 1: vsFee.Rows = 2
    vsHistory.Clear 1: vsHistory.Rows = 2
End Sub
Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, ByVal blnCard As Boolean, Optional blnOutMsg As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息
    '入参:blnCard=是否就诊卡刷卡
    '出参: blnOutMsg-已经提示,不用再外部再提示
    '返回:
    '编制:刘兴洪
    '日期:2011-01-25 16:57:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, lng卡类别ID As Long, strPassWord As String, strErrMsg As String
    Dim lng病人ID As Long, blnHavePassWord As Boolean
    Dim strIF As String, strTemp As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errH
    mstrFindNO = "": mstrFindFpNo = ""
    
    strSQL = _
        "   Select A.病人ID,Nvl(B.主页ID,0) as 主页ID,A.门诊号 as 门诊号,A.当前床号,B.出院病床," & _
        "      Nvl(b.姓名, a.姓名) As 姓名, Nvl(b.性别, a.性别) As 性别,Nvl(B.年龄,A.年龄)   as 年龄,A.IC卡号,A.就诊卡号,A.卡验证码," & _
        "       Nvl(B.费别,A.费别) as 费别,C.名称 as 当前科室,A.当前科室ID,D.名称 as 出院科室,B.出院科室ID, A.险类 as 险类,E.卡号,E.医保号,E.密码," & _
        "       A.登记时间,Nvl(B.状态,0) as 状态,Nvl(B.医疗付款方式,A.医疗付款方式) as 医疗付款方式,Nvl(B.审核标志,0) as 审核标志,B.入院日期,B.出院日期,B.病人性质,B.病人类型" & _
        " From 病人信息 A,病案主页 B,部门表 C,部门表 D,医保病人档案 E,医保病人关联表 F" & _
        " Where A.停用时间 is NULL And A.病人ID=B.病人ID(+) And A.主页ID=B.主页ID(+) " & _
        "           And A.病人ID=F.病人ID(+) And F.标志(+)=1 And F.医保号=E.医保号(+) And F.险类=E.险类(+) And F.中心 = E.中心(+)" & _
        "           And A.当前科室ID=C.ID(+) And B.出院科室ID=D.ID(+)" & _
        "           And A.停用时间 is NULL "
    
    If blnCard = True And objCard.名称 Like "姓名*" Then  '刷卡
        If IDKind.Cards.按缺省卡查找 And Not IDKind.GetfaultCard Is Nothing Then
            lng卡类别ID = IDKind.GetfaultCard.接口序号
        Else
            lng卡类别ID = "-1"
        End If
        
        If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
        If lng病人ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng病人ID
        blnHavePassWord = True
        strSQL = strSQL & " And A.病人ID=[1] "
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '病人ID
        strSQL = strSQL & " And A.病人ID=[1]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then '住院号(病人在院)
        strSQL = strSQL & " And A.病人ID = (Select Max(病人id) From 病案主页 Where 住院号 = [1])"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '门诊号
        strSQL = strSQL & " And A.门诊号=[1]"
    ElseIf Left(strInput, 1) = "." Or IDKind.IDKind = IDKind.GetKindIndex("单据号") Then
        '单据号查找
        If Left(strInput, 1) = "." Then
            strTemp = UCase(GetFullNO(Mid(strInput, 2), IIf(mint性质 = 1, 13, 14)))
        Else
            strTemp = UCase(GetFullNO(strInput, IIf(mint性质 = 1, 13, 14)))
        End If
        txtPatient.Text = strTemp
        gstrSQL = "" & _
        "   Select  distinct A.病人ID " & _
        "   From 门诊费用记录 A " & _
        "   Where A.NO=[1] and Mod(A.记录性质,10)=[2] " & _
        "              And Rownum=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strTemp, mint性质)
        If rsTemp.EOF Then
            MsgBox "注意:" & vbCrLf & "  单据号为『" & strTemp & "』不存在,请检查输入的单据是否正确!", vbInformation + vbOKOnly, gstrSysName
            Call zlClearPatiInfor
            Exit Function
        End If
        If Not GetPatient("-" & rsTemp!病人ID, False, True) Then
            Call zlClearPatiInfor
            Exit Function
        End If
        mstrFindNO = strTemp
        GetPatient = True
        Exit Function
    
    Else '当作姓名
        Select Case objCard.名称
            Case "姓名"
                If mrsInfo.State = 1 Then
                    If Not mrsInfo.EOF Then
                        If mrsInfo!姓名 = Trim(txtPatient.Text) Then GetPatient = True: Exit Function
                    End If
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
                strSQL = strSQL & " And A.病人ID = (Select Max(病人id) From 病案主页 Where 住院号 = [2])"
            Case "发票号"
                strSQL = "" & _
                "   Select distinct A.病人ID " & _
                "   From 门诊费用记录 A,票据打印内容 B,票据使用明细 C" & _
                "   Where A.NO=B.NO and Mod(A.记录性质,10)=1 and A.记录状态=1  " & _
                "               and  B.数据性质=1 And B.ID=C.打印ID and C.票种=1 And C.性质=1 And C.号码=[1] And Rownum=1 " & _
                "   "
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput, mint性质)
                If rsTemp.EOF Then
                    MsgBox "注意:" & vbCrLf & "  发票号为『" & strInput & "』不存在,请检查输入的发票号是否正确!", vbInformation + vbOKOnly, gstrSysName
                    Call zlClearPatiInfor
                    Exit Function
                End If
                If Not GetPatient(objCard, "-" & rsTemp!病人ID, False, True) Then
                    Call zlClearPatiInfor
                    Exit Function
                End If
                mstrFindFpNo = strInput
                GetPatient = True
                Exit Function
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
    txtPatient.ForeColor = Me.ForeColor
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
    
    If Not mrsInfo.EOF Then
        txtPatient.ForeColor = zlDatabase.GetPatiColor(Nvl(mrsInfo!病人类型))
        txtPatient.Text = Nvl(mrsInfo!姓名): txtOld.Text = Nvl(mrsInfo!年龄): txtSex.Text = Nvl(mrsInfo!性别)
        txt住院号.Text = Nvl(mrsInfo!门诊号)
        If Val(Nvl(mrsInfo!主页ID)) <> 0 Then
            If zlIsAllowFeeChange(Val(Nvl(mrsInfo!病人ID)), Val(Nvl(mrsInfo!主页ID))) = False Then
                txtPatient.Text = "": txtOld.Text = "": txtSex.Text = ""
                txt住院号.Text = ""
                If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
                Exit Function
            End If
        End If
        
        txtPatient.PasswordChar = ""
        '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
        txtPatient.IMEMode = 0
        mstrPassWord = strPassWord
        If Not blnHavePassWord Then mstrPassWord = Nvl(mrsInfo!卡验证码)
        GetPatient = True
        Exit Function
    Else
        txtPatient.Text = "": txtOld.Text = "": txtSex.Text = ""
        txt住院号.Text = ""
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
NotFoundPati:
    txtPatient.Text = "": txtOld.Text = "": txtSex.Text = ""
    txt住院号.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    Set mrsInfo = New ADODB.Recordset
End Function
Private Function zlGetFpToBIllNOs(ByVal strFpNo As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据指定的发票号,找出对应的单据号
    '返回:返回对应的单据号,用逗号分隔
    '编制:刘兴洪
    '日期:2011-02-25 10:50:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, strNos As String
    
    On Error GoTo errHandle
    
    strSQL = "" & _
    "   Select distinct NO From 票据打印内容 A,票据使用明细 B " & _
    "   Where A.数据性质=1 and A.ID=B.打印ID and B.票种=1 And B.号码=[1]  " & _
    "   Order by NO"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strFpNo)
    strNos = ""
    With rsTemp
        Do While Not .EOF
            strNos = strNos & "," & Nvl(rsTemp!NO)
            .MoveNext
        Loop
    End With
    If strNos <> "" Then strNos = Mid(strNos, 2)
    zlGetFpToBIllNOs = strNos
 
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function ReadListData(Optional blnFilter As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取需要销帐的明细数据
    '返回:读取成功,返回true,否则返回Flase
    '编制:刘兴洪
    '日期:2011-01-25 17:10:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng病人ID As Long, strTable As String, rsTemp As ADODB.Recordset
    Dim lngCol As Long, strSQL As String, lngRow As Long
    Dim strFilter As String, strNos As String
    Dim strWhere As String, strTable1 As String
    Dim strALLNOs As String
    If mrsInfo Is Nothing Then
        lng病人ID = 0
    ElseIf mrsInfo.State <> 1 Then
        lng病人ID = 0
    Else
        lng病人ID = Val(Nvl(mrsInfo!病人ID))
    End If
    If mstrFindNO <> "" Then
        If mint性质 = 1 Then
            strNos = Replace(GetMultiNOs(mstrFindNO), "'", "")
        Else
            strNos = mstrFindNO
        End If
        strTable1 = ",Table( f_Str2list([2])) J "
        strWhere = "  And A.NO=J.Column_Value"
    ElseIf mstrFindFpNo <> "" And mint性质 = 1 Then
        strNos = zlGetFpToBIllNOs(mstrFindFpNo)
        If strNos = "" Then
            MsgBox "未找到对应发票号的单据,请检查!"
            Exit Function
        End If
        strTable1 = ",Table( f_Str2list([2])) J "
        strWhere = "  And A.NO=J.Column_Value"
    Else
        strTable1 = ""
        strWhere = "  And A.病人ID=[1]"
    End If
    mblnSel = False
    On Error GoTo errHandle
    If blnFilter = False Then zlCommFun.ShowFlash "正在读取单据数据,请稍候 ..."
    Screen.MousePointer = 11
    DoEvents
    Me.Refresh
    If mint性质 = 1 Then
        strTable = " " & _
        "Select a.险类, a.医保, b.Id, a.单据, a.No, a.实际票号, a.序号, a.收费类别, a.从属父号, a.收费细目id, a.执行部门id, a.付数, a.数次, a.单价, a.应收金额, a.实收金额," & vbNewLine & _
        "       a.开单人, a.发生时间, a.审核人, a.审核日期, a.转出人, a.转出时间, b.结帐id " & _
        "From (Select Max(险类) as 险类, Decode(Max(险类), 0, '', '√') As 医保, '收费单' As 单据, " & vbNewLine & _
        "           NO, 实际票号, 序号, 收费类别, 从属父号, 收费细目id, 执行部门id, " & vbNewLine & _
        "           Avg(Nvl(付数, 1)) As 付数, Sum(数次) 数次, 标准单价 As 单价, Sum(应收金额) As 应收金额, " & vbNewLine & _
        "           Sum(实收金额) As 实收金额, 开单人, To_Char(Max(发生时间), 'YYYY-MM-DD HH24:MI:SS') As 发生时间, " & vbNewLine & _
        "           Max(审核人) As 审核人, Max(审核日期) As 审核日期, Max(转出人) As 转出人,Max(转出时间) As 转出时间 " & vbNewLine & _
        "      From(Select Row_Number() Over(Partition By a.ID Order By m.序号) As Rn, a.ID,Nvl(M.险类,0) as 险类, A.价格父号, " & vbNewLine & _
        "               A.NO, A.实际票号, A.序号 As 序号, A.收费类别, A.从属父号, A.收费细目id, A.执行部门id, " & vbNewLine & _
        "               A.付数, A.数次, A.标准单价, A.应收金额, " & vbNewLine & _
        "               A.实收金额, A.开单人, A.发生时间, " & vbNewLine & _
        "               Q.审核人, Q.审核日期,Q.转出人, Q.转出时间,A.记录状态" & vbNewLine & _
        "           From 门诊费用记录 A, 保险结算记录 M, 费用审核记录 Q " & strTable1 & vbNewLine & _
        "           Where Mod(A.记录性质,10) = 1  " & strWhere & _
        "               And A.记录状态 <> 0 And A.结帐id = M.记录id(+)" & vbNewLine & _
        "               And  M.性质(+) = 1 And A.ID = Q.费用id(+) And Nvl(a.附加标志,0) <> 9 " & vbNewLine & _
        "               And a.Id In (Select b.Id " & vbNewLine & _
        "                        From 门诊费用记录 B, 门诊费用记录 C, 费用审核记录 D" & vbNewLine & _
        "                        Where c.Id = d.费用id And d.记录状态 = 1 And b.No = c.No))" & vbNewLine & _
        "      Where Rn < 2" & _
        "      Group By NO, 实际票号, 序号, 收费类别, 标准单价,收费细目id, 从属父号, 执行部门id,开单人, 发生时间" & _
        "      Having Sum(数次) <> 0) A, 门诊费用记录 B Where a.No = b.No And Mod(b.记录性质,10) = 1 And a.序号 = b.序号 And b.记录状态 In (1,3) " & _
        "      And b.登记时间 = (Select Max(登记时间) From 门诊费用记录 Where NO = a.No And Mod(记录性质, 10) = 1 And 序号 = a.序号 And 记录状态 In (1,3))"
    Else
        '记帐单
        strTable = " " & _
        "    Select 0 as 险类, Decode(NULL, Null, '', '√') As 医保, Max(Decode(A.价格父号, Null, ID, 0)) As ID, '记帐单' As 单据, " & vbNewLine & _
        "           A.NO, A.实际票号, A.序号 As 序号, A.收费类别, A.从属父号, A.收费细目id, A.执行部门id, " & vbNewLine & _
        "           Avg(Nvl(A.付数, 1)) As 付数, Sum(A.数次) 数次, A.标准单价 As 单价, Sum(A.应收金额) As 应收金额, " & vbNewLine & _
        "           Sum(A.实收金额) As 实收金额, A.开单人, To_Char(A.发生时间, 'YYYY-MM-DD HH24:MI:SS') As 发生时间, " & vbNewLine & _
        "           Max(Q.审核人) As 审核人, Max(Q.审核日期) As 审核日期, Max(Q.转出人) As 转出人, " & vbNewLine & _
        "           Max(Q.转出时间) As 转出时间,0 AS 结帐ID " & vbNewLine & _
        "    From 门诊费用记录 A,  费用审核记录 Q " & strTable1 & vbNewLine & _
        "    Where  A.记录性质 = 2 " & strWhere & vbNewLine & _
        "               And A.记录状态 <> 0 And A.ID = Q.费用id(+) " & vbNewLine & _
        "           And a.Id In (Select b.Id" & vbNewLine & _
        "                        From 门诊费用记录 B, 门诊费用记录 C, 费用审核记录 D" & vbNewLine & _
        "                        Where c.Id = d.费用id And d.记录状态 = 1 And b.No = c.No)" & vbNewLine & _
        "    Group By A.NO, A.实际票号, A.序号, A.标准单价, A.收费类别, A.收费细目id, A.从属父号, A.执行部门id, " & vbNewLine & _
        "              A.开单人, A.发生时间, 结帐ID Having Sum(A.数次) <> 0"
    End If
    strSQL = "" & _
    " Select  A.ID,'' as " & mstr标志 & ",A.单据,A.No as 单据号,A.实际票号 As 票据号, " & vbNewLine & _
    "       A.序号,A.从属父号,A.收费细目ID,A.执行部门ID,A.收费类别,P.类别, " & vbNewLine & _
    "       C.编码 as 编码,Nvl(B.名称,C.名称) as 名称,E1.名称 as 商品名,C.规格," & vbNewLine & _
    "       A.付数, A.数次,C.计算单位," & vbNewLine & _
    "       ltrim(to_char(A.单价,'9999990.00000')) as 单价," & vbNewLine & _
    "       ltrim(to_char(A.应收金额,'9999990.00')) as 应收金额," & vbNewLine & _
    "       ltrim(to_char(A.实收金额,'9999990.00')) as 实收金额," & vbNewLine & _
    "       A.开单人,A.发生时间,A.医保, A.险类,A.审核人, " & vbNewLine & _
    "       A.审核日期,A.转出人,A.转出时间,A.结帐ID" & vbNewLine & _
    "From (" & strTable & ") A,收费项目目录 C,收费项目别名 B,收费项目别名 E1,收费类别 P" & _
    " Where A.收费细目ID=C.ID And A.收费细目ID=B.收费细目ID(+) And B.码类(+)=1 And B.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
    "       and A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3 " & _
    "       And A.收费类别=P.编码(+)" & _
    " Order by A.单据,A.NO,A.序号"
    If mrsFeeList Is Nothing Or blnFilter = False Then
        Set mrsFeeList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, strNos)
    Else
        mrsFeeList.Filter = 0
    End If
    vsFee.Redraw = flexRDNone
    vsFee.Clear: vsFee.Cols = 0
    Set vsFee.DataSource = mrsFeeList
    If vsFee.Rows <= 1 Then vsFee.Rows = 2
    With vsFee
        For lngCol = 0 To .Cols - 1
             .ColAlignment(lngCol) = flexAlignLeftCenter
             .FixedAlignment(lngCol) = flexAlignCenterCenter
              .ColKey(lngCol) = Trim(.TextMatrix(0, lngCol))
              If .ColKey(lngCol) Like "*ID" Or InStr(1, ",险类,编码,序号,从属父号,转出标志,收费类别,", "," & .ColKey(lngCol) & ",") > 0 Then
                    .ColHidden(lngCol) = True
              ElseIf .ColKey(lngCol) Like "*数*" Or .ColKey(lngCol) Like "*价*" Or .ColKey(lngCol) Like "*额" Then
                    .ColAlignment(lngCol) = flexAlignRightCenter
              End If
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        .ColDataType(.ColIndex(mstr标志)) = flexDTBoolean
        Call .AutoSize(0, .Cols - 1)
        zl_vsGrid_Para_Restore mlngModule, vsFee, Me.Caption, IIf(mint性质 = 1, "退费列表", "销帐列表"), True
        '画线
        Dim strNO As String, str单据 As String
        strALLNOs = ""
        For lngRow = 1 To .Rows - 1
            If strNO <> Trim(.TextMatrix(lngRow, .ColIndex("单据号"))) _
                 And strNO <> "" Then
                '画出分隔线
                .Select lngRow, .FixedCols, lngRow, .Cols - 1
                .CellBorder vbBlue, 0, 1, 0, 0, 0, 0
            End If
            .Cell(flexcpData, lngRow, .ColIndex("单据号")) = .TextMatrix(lngRow, .ColIndex("单据号"))
            .Cell(flexcpData, lngRow, .ColIndex(mstr标志)) = Val(.TextMatrix(lngRow, .ColIndex(mstr标志)))
            strNO = Trim(.TextMatrix(lngRow, .ColIndex("单据号")))
            str单据 = Trim(.TextMatrix(lngRow, .ColIndex("单据")))
            strALLNOs = strALLNOs & "," & strNO
        Next
        .Editable = flexEDKbdMouse
    End With
    If strALLNOs <> "" Then strALLNOs = Mid(strALLNOs, 2)
    If blnFilter = False Then zlCommFun.StopFlash
    vsFee.Redraw = flexRDBuffered
    '加载结算方式
    Call InitBlanceData(strALLNOs)
    Call CalcSUMMony
    Call SetBlanceShow
    Call StatusShowBillSum
    
    Screen.MousePointer = 0
    ReadListData = True
    Exit Function
errHandle:
    vsFee.Redraw = flexRDBuffered
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
   If blnFilter = False Then zlCommFun.StopFlash
End Function
Private Function ReadHistoryListData(Optional blnFilter As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取需要销帐的明细数据
    '返回:读取成功,返回true,否则返回Flase
    '编制:刘兴洪
    '日期:2011-01-25 17:10:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng病人ID As Long, strTable As String, rsTemp As ADODB.Recordset
    Dim lngCol As Long, strSQL As String, lngRow As Long
    Dim strFilter As String, strNos As String
    Dim strWhere As String
    If mrsInfo Is Nothing Then
        lng病人ID = 0
    ElseIf mrsInfo.State <> 1 Then
        lng病人ID = 0
    Else
        lng病人ID = Val(Nvl(mrsInfo!病人ID))
    End If
    On Error GoTo errHandle
    If blnFilter = False Then zlCommFun.ShowFlash "正在读取历史转出单据数据,请稍候 ..."
    Screen.MousePointer = 11
    DoEvents
    Me.Refresh
    If mint性质 = 1 Then
        strTable = "" & _
        " Select Max(Nvl(险类, 0)) As 险类, Decode(Max(Nvl(险类, 0)), 0, '', '√') As 医保, Max(Decode(价格父号, Null, ID, 0)) As ID," & vbNewLine & _
        "       '收费单' As 单据, NO, 实际票号, Nvl(价格父号, 序号) As 序号, 收费类别, 从属父号, 收费细目id, 执行部门id, Avg(Nvl(付数, 1)) As 付数, -1 * Avg(数次) 数次," & vbNewLine & _
        "       Sum(标准单价) As 单价, -1 * Sum(应收金额) As 应收金额, -1 * Sum(实收金额) As 实收金额, 开单人," & vbNewLine & _
        "       To_Char(发生时间, 'YYYY-MM-DD HH24:MI:SS') As 发生时间, 结帐id, Max(审核人) As 审核人, Max(审核日期) As 审核日期, Max(转出人) As 转出人," & vbNewLine & _
        "       Max(转出时间) As 转出时间" & vbNewLine & _
        " From (Select Row_Number() Over(Partition By a.Id Order By m.序号) As Rn, a.Id, m.险类, a.No, a.实际票号, a.价格父号, a.序号, a.收费类别," & vbNewLine & _
        "              a.从属父号, a.收费细目id, a.执行部门id, a.付数, a.数次, a.标准单价, a.应收金额, a.实收金额, a.开单人, a.发生时间, a.结帐id, q.审核人, q.审核日期, q.转出人," & vbNewLine & _
        "              q.转出时间" & vbNewLine & _
        "       From 门诊费用记录 A, 保险结算记录 M, 费用审核记录 Q, 门诊费用记录 K" & vbNewLine & _
        "       Where Mod(a.记录性质, 10) = 1 and A.病人ID=[1] " & strWhere & _
        "             And a.结帐id = m.记录id(+) And m.性质(+) = 1 " & vbNewLine & _
        "             And q.费用id(+) = k.Id" & vbNewLine & _
        "             And k.Id In (Select Min(ID) From 门诊费用记录 Where NO = a.No And Mod(记录性质,10) = 1 And 序号 = a.序号)" & vbNewLine & _
        "             And a.Id In (Select Max(d.Id)" & vbNewLine & _
        "                      From 门诊费用记录 D, 门诊费用记录 B, 费用审核记录 C" & vbNewLine & _
        "                      Where b.病人id + 0 = a.病人id And b.序号 = a.序号 And b.Id = c.费用id And a.No = b.No And d.记录状态 = 2 And" & vbNewLine & _
        "                            d.No = b.No And c.记录状态 = 2" & vbNewLine & _
        "                      Group By d.序号))" & vbNewLine & _
        " Where Rn < 2" & vbNewLine & _
        " Group By NO, 实际票号, Nvl(价格父号, 序号), 收费类别, 收费细目id, 从属父号, 执行部门id, 开单人, 发生时间, 结帐id"
    Else
        '记帐单
        strTable = " " & _
        "    Select 0 as 险类, Decode(NULL, Null, '', '√') As 医保, Max(Decode(A.价格父号, Null, a.ID, 0)) As ID, '记帐单' As 单据, " & vbNewLine & _
        "           A.NO, A.实际票号, Nvl(A.价格父号, A.序号) As 序号, A.收费类别, A.从属父号, A.收费细目id, A.执行部门id, " & vbNewLine & _
        "           Avg(Nvl(A.付数, 1)) As 付数, -1 * Avg(A.数次) 数次, Sum(A.标准单价) As 单价, -1 * Sum(A.应收金额) As 应收金额, " & vbNewLine & _
        "           -1 * Sum(A.实收金额) As 实收金额, A.开单人, To_Char(A.发生时间, 'YYYY-MM-DD HH24:MI:SS') As 发生时间, 0 as 结帐id, " & vbNewLine & _
        "           Max(Q.审核人) As 审核人, Max(Q.审核日期) As 审核日期, Max(Q.转出人) As 转出人, " & vbNewLine & _
        "           Max(Q.转出时间) As 转出时间 " & vbNewLine & _
        "    From 门诊费用记录 A,  费用审核记录 Q, 门诊费用记录 K " & vbNewLine & _
        "    Where  A.记录性质 = 2  and A.病人ID=[1] " & strWhere & vbNewLine & _
        "               And q.费用id(+) = k.Id " & vbNewLine & _
        "           And k.Id In (Select Min(ID) From 门诊费用记录 Where NO = a.No And 记录性质 = 2 And 序号 = a.序号) " & _
        "           And a.Id In (Select Max(d.Id)" & vbNewLine & _
        "         From 门诊费用记录 D, 门诊费用记录 B, 费用审核记录 C" & vbNewLine & _
        "         Where b.病人id + 0 = a.病人id And b.序号=a.序号 And b.Id = c.费用id And a.No = b.No And d.记录状态 = 2 And d.No = b.No And c.记录状态 = 2" & vbNewLine & _
        "         Group By d.序号) " & _
        "    Group By A.NO, A.实际票号, Nvl(A.价格父号, A.序号), A.收费类别, A.收费细目id, A.从属父号, A.执行部门id, " & vbNewLine & _
        "              A.开单人, A.发生时间"
    End If
    strSQL = "" & _
    " Select  A.ID,A.单据,A.No as 单据号,A.实际票号 As 票据号, " & vbNewLine & _
    "       A.序号,A.从属父号,A.收费细目ID,A.执行部门ID,A.收费类别,P.类别, " & vbNewLine & _
    "       C.编码 as 编码,Nvl(B.名称,C.名称) as 名称,E1.名称 as 商品名,C.规格," & vbNewLine & _
    "       A.付数, A.数次,C.计算单位," & vbNewLine & _
    "       ltrim(to_char(A.单价,'9999990.00000')) as 单价," & vbNewLine & _
    "       ltrim(to_char(A.应收金额,'9999990.00')) as 应收金额," & vbNewLine & _
    "       ltrim(to_char(A.实收金额,'9999990.00')) as 实收金额," & vbNewLine & _
    "       A.开单人,A.发生时间, A.结帐ID,A.医保, A.险类,A.审核人, " & vbNewLine & _
    "       A.审核日期,A.转出人,A.转出时间" & vbNewLine & _
    "From (" & strTable & ") A,收费项目目录 C,收费项目别名 B,收费项目别名 E1,收费类别 P" & _
    " Where A.收费细目ID=C.ID And A.收费细目ID=B.收费细目ID(+) And B.码类(+)=1 And B.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
    "       and A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3 " & _
    "       And A.收费类别=P.编码(+)" & _
    " Order by A.单据,A.实际票号,A.NO,A.序号"
    
    If mrsHistoryList Is Nothing Or blnFilter = False Then
        Set mrsHistoryList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, strNos)
    Else
        mrsHistoryList.Filter = 0
    End If
    vsHistory.Redraw = flexRDNone
    vsHistory.Clear: vsHistory.Cols = 0
    Set vsHistory.DataSource = mrsHistoryList
    If vsHistory.Rows <= 1 Then vsHistory.Rows = 2
    
    With vsHistory
        For lngCol = 0 To .Cols - 1
             .ColAlignment(lngCol) = flexAlignLeftCenter
             .FixedAlignment(lngCol) = flexAlignCenterCenter
              .ColKey(lngCol) = Trim(.TextMatrix(0, lngCol))
              If .ColKey(lngCol) Like "*ID" Or .ColKey(lngCol) = "从属父号" Or .ColKey(lngCol) = "转出标志" Or .ColKey(lngCol) = "收费类别" Then
                    .ColHidden(lngCol) = True
              ElseIf .ColKey(lngCol) Like "*数*" Or .ColKey(lngCol) Like "*价*" Or .ColKey(lngCol) Like "*额" Then
                    .ColAlignment(lngCol) = flexAlignRightCenter
              End If
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .Cols - 1)
        zl_vsGrid_Para_Restore mlngModule, vsHistory, Me.Caption, IIf(mint性质 = 1, "历史退费列表", "历史销帐列表"), True
        '画线
        Dim strNO As String, str单据 As String
        For lngRow = 1 To .Rows - 1
            If strNO <> Trim(.TextMatrix(lngRow, .ColIndex("单据号"))) _
                 And strNO <> "" Then
                '画出分隔线
                .Select lngRow, .FixedCols, lngRow, .Cols - 1
                .CellBorder vbBlue, 0, 1, 0, 0, 0, 0
            End If
            .Cell(flexcpData, lngRow, .ColIndex("单据号")) = .TextMatrix(lngRow, .ColIndex("单据号"))
            strNO = Trim(.TextMatrix(lngRow, .ColIndex("单据号")))
            str单据 = Trim(.TextMatrix(lngRow, .ColIndex("单据")))
        Next
        .Editable = flexEDNone
    End With
    If blnFilter = False Then zlCommFun.StopFlash
    vsHistory.Redraw = flexRDBuffered
    
    Screen.MousePointer = 0
    ReadHistoryListData = True
    Exit Function
errHandle:
    vsHistory.Redraw = flexRDBuffered
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
   If blnFilter = False Then zlCommFun.StopFlash
End Function

Private Sub vsBalance_DblClick()
    Dim i As Integer
    With vsBalance
        'If .TextMatrix(.Row, 0) = "收款结算" Then Exit Sub
        If .Cell(flexcpFontUnderline, .Row, .Col, .Row, .Col) = True Then
            If Val(.TextMatrix(.Row, .Col + 1)) = 0 Then Exit Sub
            For i = 0 To .Cols - 1
                If IsNumeric(.TextMatrix(0, i)) = False Then
                If InStr(mstrStyle, .TextMatrix(0, i)) > 0 Then
                    txtSum.Text = Val(txtSum.Text) + Val(.TextMatrix(0, i + 1))
                    .TextMatrix(0, i + 1) = "0"
                End If
                End If
            Next i
        End If
    End With
End Sub

Private Sub vsFee_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsFee
        Select Case Col
        Case .ColIndex(mstr标志)
            txtSum.Text = 0
            SetNOBill .TextMatrix(Row, .ColIndex("单据")), .TextMatrix(Row, .ColIndex("单据号")), Val(.TextMatrix(Row, .Col)) <> 0
            mblnSel = Val(.TextMatrix(Row, .Col)) <> 0
            Call SetRowSelected(Row)
            mblnChange = True
            Call CalcSUMMony
            Call SetBlanceShow
            If mblnSel = False Then mblnSel = IsCheckSelNo
        Case Else
        End Select
    End With
End Sub

Private Sub vsFee_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsFee, Me.Caption, IIf(mint性质 = 1, "退费列表", "销帐列表"), True
End Sub

Private Sub vsFee_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
   Dim cur合计 As Currency, i As Long
    If NewRow <> OldRow Then
'        With vsFee
'            If .TextMatrix(NewRow, .ColIndex("单据号")) <> "" Then
'                For i = NewRow - 1 To .FixedRows Step -1
'                    If .TextMatrix(i, .ColIndex("单据号")) = .TextMatrix(NewRow, .ColIndex("单据号")) Then
'                        cur合计 = cur合计 + Val(.TextMatrix(i, .ColIndex("实收金额")))
'                    Else
'                        Exit For
'                    End If
'                Next
'                For i = NewRow To .Rows - 1
'                    If .TextMatrix(i, .ColIndex("单据号")) = .TextMatrix(NewRow, .ColIndex("单据号")) Then
'                        cur合计 = cur合计 + Val(.TextMatrix(i, .ColIndex("实收金额")))
'                    Else
'                        Exit For
'                    End If
'                Next
'            End If
            Call StatusShowBillSum
            'Me.stbThis.Panels(2).Text = "当前单据合计:" & Format(cur合计, gstrDec)
'        End With
    End If
End Sub

Private Sub vsFee_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsFee, Me.Caption, IIf(mint性质 = 1, "退费列表", "销帐列表"), True
End Sub

Private Sub vsFee_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsFee
        Select Case Col
        Case .ColIndex(mstr标志)
            If CheckIsInput(Row) = False Then Cancel = True: Exit Sub
        Case Else
            Cancel = True
        End Select
    End With
End Sub
 

Private Sub vsFee_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:按单据画线和清除线
    '编制:刘兴洪
    '日期:2011-01-26 09:57:32
    '说明:
    '       1.OwnerDraw要设置为Over(画出单元所有内容)
    '       2.Cell的GridLine从上下左右向内都是从第1根线开始
    '       3.Cell的Border从左上是从第2根线开始,右下是从第1根线开始
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    Dim strText As String
    strText = " "
    With vsFee
        '擦除相关行列的边线及内容
        lngLeft = .ColIndex(mstr标志): lngRight = .ColIndex(mstr标志)
        
        If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        Call GetBillNOStartAndEndRow(Row, lngBegin, lngEnd)
        If lngBegin = lngEnd Then Exit Sub
        
        vRect.Left = Left '擦除左边表格线
        vRect.Right = Right - 1 '保留右边表格线
        If Row = lngBegin Then
            vRect.Top = Bottom - 1 '首行保留文字内容
            vRect.Bottom = Bottom
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 1 '底行保留下边线
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, SysColor2RGB(.BackColorSel)
        Else
            SetBkColor hDC, SysColor2RGB(.BackColor)
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, strText, 1, 0
        Done = True
    End With
End Sub

Private Function SysColor2RGB(ByVal lngColor As Long) As Long
'功能：将VB的系统颜色转换为RGB色
    If lngColor < 0 Then
        Call OleTranslateColor(lngColor, 0, lngColor)
    End If
    SysColor2RGB = lngColor
End Function

Private Sub GetBillNOStartAndEndRow(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取单据行
    '编制:刘兴洪
    '日期:2011-01-26 10:01:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    lngBegin = lngRow: lngEnd = lngRow
    With vsFee
        For i = lngRow - 1 To .FixedRows Step -1
            If .TextMatrix(i, .ColIndex("单据号")) = .TextMatrix(lngRow, .ColIndex("单据号")) Then
                lngBegin = i
            Else
                Exit For
            End If
        Next
        For i = lngRow To .Rows - 1
            If .TextMatrix(i, .ColIndex("单据号")) = .TextMatrix(lngRow, .ColIndex("单据号")) Then
                lngEnd = i
            Else
                Exit For
            End If
        Next
    End With
End Sub
Private Function SetNOBill(ByVal str单据 As String, ByVal strNO As String, ByVal blnSel As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:按单据全选或全清单据
    '入参:str单据-单据类型(收费单,记帐单)
    '       strNO-指定的NO
    '        blnSel:true表示全选,否则全清
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2011-01-24 10:47:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo errHandle
    With vsFee
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("单据号")) <> "" _
                And .TextMatrix(i, .ColIndex("单据号")) = strNO Then
                .TextMatrix(i, .ColIndex(mstr标志)) = IIf(blnSel, -1, 0)
            End If
        Next
    End With
    SetNOBill = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function CheckMulitBillValied(ByVal strNO As String, ByVal lngInsure As Long, _
    ByRef strOutNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查多单据收费单是否合法
    '入参:strNO-单据号
    '出参:strOutNos-返回的多单据,单据为:A0001,A002...
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-02-24 14:10:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNos As String, strTemp As String, strNo1 As String
    Dim i As Long, m As Long, varTemp As Variant
    Dim strNOsTemp As String
    On Error GoTo errHandle
    With vsFee
        If mint性质 <> 1 Then
            '记帐单,按单据直接返回
            strOutNos = strNO: CheckMulitBillValied = True: Exit Function
        End If
        strNos = Replace(GetMultiNOs(strNO), "'", "") '一起收费的其他单据
        If InStr(1, strNos, ",") = 0 Then
            '非多单据收费,直接返回
            strOutNos = strNO: CheckMulitBillValied = True: Exit Function
        End If
        strTemp = "": strNOsTemp = ""
        For i = 1 To .Rows - 1
             strNo1 = Trim(.TextMatrix(i, .ColIndex("单据号")))
             If strNo1 <> strNO Then
                '1. 检查是否存在未勾选的退费单
                If InStr(1, strTemp & ",", "," & strNo1 & ",") = 0 Then
                    If InStr(1, "," & strNos & ",", "," & strNo1 & ",") > 0 Then
                         If GetVsGridBoolColVal(vsFee, i, .ColIndex(mstr标志)) = False Then
                            MsgBox "注意:" & vbCrLf & "    单据号为" & strNo1 & "与单据为" & strNO & "的收费单 " & vbCrLf & "    是多单据收费,所以必须一起退!", vbInformation + vbOKOnly, gstrSysName
                             .Row = i: Exit Function
                        End If
                        strNOsTemp = strNOsTemp & "," & strNo1
                     End If
                    strTemp = strTemp & "," & strNo1
                End If
             Else
                strNOsTemp = strNOsTemp & "," & strNo1
             End If
        Next
        '2.检查未提取出来的单据
        varTemp = Split(strNos, ",")
        strTemp = ""
    
        For m = 0 To UBound(varTemp)
            If InStr(1, "," & strNOsTemp & ",", "," & varTemp(m) & ",") = 0 Then
                strTemp = strTemp & "," & varTemp(m)
            End If
        Next
            
            If strTemp <> "" Then
                strTemp = Mid(strTemp, 2)
                If MsgBox("注意:" & vbCrLf & "单据为" & strNO & "是多单据收费,其中有以下单据:" & vbCrLf & strTemp & vbCrLf & "    异常,可能是因为上次退费时异常,是否继续操作?", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName) = vbNo Then
                    .Row = 1: Exit Function
                End If
                If strNOsTemp <> "" Then
                    strNos = Mid(strNOsTemp, 2)
                Else
                    MsgBox "数据异常,不能退费!", vbOKOnly + vbInformation, gstrSysName
                    .Row = 1: Exit Function
                    Exit Function
                End If
            End If
         
        '3.合法删除,返回
         strOutNos = strNos
    End With
    CheckMulitBillValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ExecuteDelBill(ByVal strDelDate As String, ByVal strNos As String, intInsure As Integer, ByVal lng结帐ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行相关退费操作
    '入参:strNos-单据号:可以是多单据
    '       lngInsure-险类
    '返回:执行成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-02-24 15:35:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, k As Long, varTemp  As Variant, strAllBalance      As String, strBalance As String, bln产生误差 As Boolean
    Dim bln医保接口打印票据 As Boolean, bln多单据一次结算 As Boolean, blnYB结算作废 As Boolean, bln退费后打印回单 As Boolean
    Dim lng领用ID As Long, cllPro As Collection, blnTrans As Boolean, lng冲销ID As Long, str交易流水号 As String, str交易说明 As String
    Dim lng结帐ID1 As Long, varBalance As Variant, strAdvance As String, strInvoice As String
    Dim strSQL As String, j As Long, blnTransMedicare As Boolean, rsTmp As ADODB.Recordset, bln医保单张退 As Boolean, blnTurnAll As Boolean
    Dim str结算方式 As String, cur结算金额 As Currency, cur可分配额 As Currency, cur误差金额 As Currency, cur余额 As Currency, cur退款合计 As Currency
    Dim strDelNOs As String, lng病人ID As Long, blnExecuteThreeSwap As Boolean
    
    If intInsure <> 0 Then
        bln医保接口打印票据 = gclsInsure.GetCapability(support医保接口打印票据, , intInsure, CStr(lng结帐ID))
        bln多单据一次结算 = Not (gclsInsure.GetCapability(83, , intInsure) Or gclsInsure.GetCapability(85, , intInsure))
        blnYB结算作废 = gclsInsure.GetCapability(support门诊结算作废, , intInsure)
        If blnYB结算作废 = False Then
            MsgBox "注意:" & vbCrLf & "   单据号为" & strNos & "的单据,不支持医保结算作废,请检查"
            Exit Function
        End If
        bln退费后打印回单 = gclsInsure.GetCapability(support退费后打印回单, , intInsure)
    End If
    
    If intInsure <> 0 And bln医保接口打印票据 Then
        Dim strUserType As String
        Dim lngShareUseID As Long
        If mrsInfo Is Nothing Then
            lng病人ID = mlng病人ID
        ElseIf mrsInfo.State <> 1 Then
            lng病人ID = mlng病人ID
        Else
            lng病人ID = Val(Nvl(mrsInfo!病人ID))
        End If
        strUserType = zl_GetInvoiceUserType(lng病人ID, 0, intInsure)
        lngShareUseID = zl_GetInvoiceShareID(1121, strUserType)
         
        lng领用ID = GetInvoiceGroupID(1, 1, lng领用ID, lngShareUseID)
        Select Case lng领用ID
            Case -1
                MsgBox "你没有自用和共用的收费票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                Exit Function
            Case -2
                MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                Exit Function
        End Select
        strInvoice = GetNextBill(lng领用ID)
    End If
    
    '获取结帐ID
    Err = 0: On Error GoTo errHandle
    Set cllPro = New Collection
    varTemp = Split(strNos, ",")
    For i = 0 To UBound(varTemp)
            'Zl_门诊转住院_收费转出
            strSQL = "Zl_门诊转住院_收费转出("
            '     No_In         住院费用记录.NO%Type,
            strSQL = strSQL & "'" & varTemp(i) & "',"
            '     操作员编号_In 住院费用记录.操作员编号%Type,
            strSQL = strSQL & "'" & UserInfo.编号 & "',"
            '     操作员姓名_In 住院费用记录.操作员姓名%Type,
            strSQL = strSQL & "'" & UserInfo.姓名 & "',"
            '     退费时间_In   住院费用记录.发生时间%Type,
            strSQL = strSQL & "to_date('" & strDelDate & "','yyyy-mm-dd hh24:mi:ss'),"
            '     门诊退费_In   Number := 0(门诊退费_In:0-门诊转住院立即销帐;1-门诊退费模式:为1时:入院科室id_In和主页ID_IN可以不传)
            strSQL = strSQL & "1,"
            '     入院科室id_In 住院费用记录.开单部门id%Type := Null,
            strSQL = strSQL & "Null,"
            '     主页id_In     住院费用记录.主页id%Type := Null
            strSQL = strSQL & "Null,"
            '     结算方式_In   病人预交记录.结算方式%Type := Null
            strSQL = strSQL & IIf(picBack.Visible, "'" & cboStyle.Text & "'", "Null") & ","
           With vsFee
                lng结帐ID1 = 0
                For j = 1 To .Rows - 1
                    If .TextMatrix(j, .ColIndex("单据号")) = varTemp(i) Then
                        lng结帐ID1 = Val(.TextMatrix(j, .ColIndex("结帐ID")))
                        Exit For
                    End If
                Next
           End With
           strAllBalance = strAllBalance & "," & lng结帐ID1
          cllPro.Add Array(strSQL, lng结帐ID1, varTemp(0), CStr(varTemp(0)), varTemp(i))
    Next
    
    
     If intInsure <> 0 And bln多单据一次结算 Then
        On Error GoTo errH: blnTrans = True
        gcnOracle.BeginTrans
            '从最后一张开始退
        For i = cllPro.Count To 1 Step -1
            If InStr("," & mstrUsedBills & ",", "," & Val(cllPro(i)(1)) & ",") = 0 Then
                blnExecuteThreeSwap = False
                lng冲销ID = zlDatabase.GetNextId("病人结帐记录")
                If mcur误差 <> 0 Then
                    Call zlDatabase.ExecuteProcedure(CStr(cllPro(i)(0)) & lng冲销ID & "," & Val(cllPro(i)(1)) & "," & mcur误差 & ")", Me.Caption)
                    mcur误差 = 0
                Else
                    Call zlDatabase.ExecuteProcedure(CStr(cllPro(i)(0)) & lng冲销ID & "," & Val(cllPro(i)(1)) & ")", Me.Caption)
                End If
                
                If ExecuteThreeSwap(Val(cllPro(i)(1)), lng冲销ID, str交易流水号, str交易说明) = True Then
                    blnExecuteThreeSwap = True
                End If
                
                'Zl_门诊转住院_三方卡结算
                strSQL = "Zl_门诊转住院_三方卡结算("
                '  No_In         住院费用记录.NO%Type,
                strSQL = strSQL & "'" & varTemp(i - 1) & "',"
                '  操作员编号_In 住院费用记录.操作员编号%Type,
                strSQL = strSQL & "'" & UserInfo.编号 & "',"
                '  操作员姓名_In 住院费用记录.操作员姓名%Type,
                strSQL = strSQL & "'" & UserInfo.姓名 & "',"
                '  退费时间_In   住院费用记录.登记时间%Type, --多张单据退费时,每张单据的退费时间相同,都是系统当前时间
                strSQL = strSQL & "to_date('" & strDelDate & "','yyyy-mm-dd hh24:mi:ss'),"
                '  门诊退费_In   Number := 0,
                strSQL = strSQL & "" & 1 & ","
                '  入院科室id_In 病人预交记录.科室id%Type,
                strSQL = strSQL & "Null,"
                '  主页id_In     住院费用记录.主页id%Type, --医保入院补充登记时才传入
                strSQL = strSQL & "Null,"
                '  三方退费_In   Number := 0,
                strSQL = strSQL & "" & IIf(blnExecuteThreeSwap = True, 1, 0) & ","
                '  结帐ID_In     住院费用记录.结帐id%Type)
                strSQL = strSQL & "" & lng冲销ID & ")"
                Call zlDatabase.ExecuteProcedure(strSQL, "三方卡结算")
                mstrUsedBills = mstrUsedBills & "," & Val(cllPro(i)(1))
            End If
        Next
        
        '先产生票据，医保接口才能取到
        If bln医保接口打印票据 Then
            strSQL = "zl_门诊收费记录_RePrint('" & CStr(cllPro(1)(3)) & "','" & strInvoice & "'," & ZVal(lng领用ID) & ",'" & UserInfo.姓名 & "'," & _
                "To_Date('" & Format(strDelDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),1,1)"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        End If
        strAdvance = strAllBalance
        If Not gclsInsure.ClinicDelSwap(Val(cllPro(cllPro.Count)(1)), , intInsure, strAdvance) Then
            gcnOracle.RollbackTrans
            MsgBox "医保结算失败，无法进行门诊费用转出院操作。", vbInformation, gstrSysName
            Exit Function
        Else
            blnTransMedicare = True
        End If

        If Not (strAdvance = strAllBalance Or strAdvance = "") Then
            '根据返回的结算信息，修正预交记录，strAdvance返回格式:结算方式1|金额||结算方式2:金额...
            '先分摊到每张单据上
            Set rsTmp = GetBalanceSet
            varBalance = Split(strAdvance, "||")
            For i = 0 To UBound(varBalance)
                str结算方式 = Split(varBalance(i), "|")(0)
                cur结算金额 = -1 * Val(Split(varBalance(i), "|")(1))
                For k = 0 To UBound(varTemp)
                    cur可分配额 = Get实收金额(varTemp(k))
                    rsTmp.Filter = "单据序号=" & k
                    For j = 1 To rsTmp.RecordCount
                        cur可分配额 = cur可分配额 - rsTmp!结算金额
                        rsTmp.MoveNext
                    Next
                    If cur可分配额 > 0 Then
                        If cur可分配额 <= cur结算金额 Then
                            cur结算金额 = cur结算金额 - cur可分配额
                        Else
                            cur可分配额 = cur结算金额
                            cur结算金额 = 0
                        End If
                        rsTmp.AddNew
                        rsTmp!单据序号 = k
                        rsTmp!结算方式 = str结算方式
                        rsTmp!结算金额 = cur可分配额
                        rsTmp.Update
                        
                        If cur结算金额 = 0 Then Exit For
                    End If
                Next
            Next
            
            For k = 0 To UBound(varTemp)
                strBalance = ""
                cur误差金额 = 0
                cur余额 = Get实收金额(varTemp(k))
                
                rsTmp.Filter = "单据序号=" & k
                For i = 1 To rsTmp.RecordCount
                    strBalance = IIf(strBalance = "", "", strBalance & "||") & rsTmp!结算方式 & "|" & -1 * rsTmp!结算金额
                    cur余额 = cur余额 - rsTmp!结算金额
                    rsTmp.MoveNext
                Next

                '退为指定的结算方式，如果是现金，可能产生新的误差金额
                'If cbo退款方式.ItemData(cbo退款方式.ListIndex) = 1 Then
                    cur结算金额 = Format(CentMoney(cur余额), "0.00")
                    cur误差金额 = cur结算金额 - cur余额
'                Else
'                    cur结算金额 = cur余额
'                End If
                cur退款合计 = cur退款合计 + cur结算金额
                lng结帐ID = GetDelBalanceID(varTemp(k))
                strSQL = "zl_门诊收费结算_Update(" & lng结帐ID & ",'" & "现金" & "|" & -1 * cur结算金额 & "| ',0,'" & strBalance & "'," & -1 * cur误差金额 & ")"
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            Next
        End If
        gcnOracle.CommitTrans: blnTrans = False
        If blnTransMedicare Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, intInsure)
     Else
         '从最后一张开始退
        For i = cllPro.Count To 1 Step -1
            On Error GoTo errH
            blnExecuteThreeSwap = False
            bln医保单张退 = False: blnTurnAll = False
            lng冲销ID = zlDatabase.GetNextId("病人结帐记录")
            If intInsure <> 0 Then
                bln医保单张退 = IsYBSingle(CStr(cllPro(i)(4)), intInsure)
            Else
                blnTurnAll = CheckAllTurn(CStr(cllPro(i)(4)))
                If InStr("," & mstrUsedBills & ",", "," & Val(cllPro(i)(1)) & ",") > 0 Then blnTurnAll = True
            End If
            If bln医保单张退 Or (intInsure = 0 And Not blnTurnAll) Then
                If InStr("," & mstrUsedBills & ",", "," & Val(cllPro(i)(1)) & ",") = 0 Then
                    gcnOracle.BeginTrans: blnTrans = True
                    If mcur误差 <> 0 Then
                        Call zlDatabase.ExecuteProcedure(CStr(cllPro(i)(0)) & lng冲销ID & ",Null," & mcur误差 & ")", Me.Caption)
                        mcur误差 = 0
                    Else
                        Call zlDatabase.ExecuteProcedure(CStr(cllPro(i)(0)) & lng冲销ID & ")", Me.Caption)
                    End If
                    
                    blnTransMedicare = False
                    If intInsure <> 0 Then                    '处理医保接口
                          If blnYB结算作废 Then
                                strAdvance = lng冲销ID & "|" & "0" & "|" & CStr(cllPro(i)(4))
                                If Not gclsInsure.ClinicDelSwap(CStr(cllPro(i)(1)), True, intInsure, strAdvance) Then
                                    gcnOracle.RollbackTrans
                                    MsgBox "医保结算失败，无法进行门诊费用转出院操作。", vbInformation, gstrSysName
                                    Exit Function
                                Else
                                    blnTransMedicare = True
                                End If
                            End If
                    End If
                    gcnOracle.CommitTrans: blnTrans = False
                    If blnTransMedicare Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, intInsure)
                    
                    If ExecuteThreeSwap(Val(cllPro(i)(1)), lng冲销ID, str交易流水号, str交易说明) = True Then
                        blnExecuteThreeSwap = True
                    End If
                    
                    'Zl_门诊转住院_三方卡结算
                    strSQL = "Zl_门诊转住院_三方卡结算("
                    '  No_In         住院费用记录.NO%Type,
                    strSQL = strSQL & "'" & varTemp(i - 1) & "',"
                    '  操作员编号_In 住院费用记录.操作员编号%Type,
                    strSQL = strSQL & "'" & UserInfo.编号 & "',"
                    '  操作员姓名_In 住院费用记录.操作员姓名%Type,
                    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
                    '  退费时间_In   住院费用记录.登记时间%Type, --多张单据退费时,每张单据的退费时间相同,都是系统当前时间
                    strSQL = strSQL & "to_date('" & strDelDate & "','yyyy-mm-dd hh24:mi:ss'),"
                    '  门诊退费_In   Number := 0,
                    strSQL = strSQL & "" & 1 & ","
                    '  入院科室id_In 病人预交记录.科室id%Type,
                    strSQL = strSQL & "Null,"
                    '  主页id_In     住院费用记录.主页id%Type, --医保入院补充登记时才传入
                    strSQL = strSQL & "Null,"
                    '  三方退费_In   Number := 0,
                    strSQL = strSQL & "" & IIf(blnExecuteThreeSwap = True, 1, 0) & ","
                    '  结帐ID_In     住院费用记录.结帐id%Type)
                    strSQL = strSQL & "" & lng冲销ID & ")"
                    Call zlDatabase.ExecuteProcedure(strSQL, "三方卡结算")
                    
                    strDelNOs = strDelNOs & IIf(strDelNOs = "", "", ",") & cllPro(i)(0)
                End If
            Else
                If InStr("," & mstrUsedBills & ",", "," & Val(cllPro(i)(1)) & ",") = 0 Then
                    gcnOracle.BeginTrans: blnTrans = True
                    If mcur误差 <> 0 Then
                        Call zlDatabase.ExecuteProcedure(CStr(cllPro(i)(0)) & lng冲销ID & "," & Val(cllPro(i)(1)) & "," & mcur误差 & ")", Me.Caption)
                        mcur误差 = 0
                    Else
                        Call zlDatabase.ExecuteProcedure(CStr(cllPro(i)(0)) & lng冲销ID & "," & Val(cllPro(i)(1)) & ")", Me.Caption)
                    End If
                    
                    blnTransMedicare = False
                    If intInsure <> 0 Then                    '处理医保接口
                          If blnYB结算作废 Then
                                strAdvance = lng冲销ID & "|" & "0"
                                If Not gclsInsure.ClinicDelSwap(CStr(cllPro(i)(1)), True, intInsure, strAdvance) Then
                                    gcnOracle.RollbackTrans
                                    MsgBox "医保结算失败，无法进行门诊费用转出院操作。", vbInformation, gstrSysName
                                    Exit Function
                                Else
                                    blnTransMedicare = True
                                End If
                            End If
                    End If
                    gcnOracle.CommitTrans: blnTrans = False
                    If blnTransMedicare Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, intInsure)
                    
                    If ExecuteThreeSwap(Val(cllPro(i)(1)), lng冲销ID, str交易流水号, str交易说明) = True Then
                        blnExecuteThreeSwap = True
                    End If
                    
                    'Zl_门诊转住院_三方卡结算
                    strSQL = "Zl_门诊转住院_三方卡结算("
                    '  No_In         住院费用记录.NO%Type,
                    strSQL = strSQL & "'" & varTemp(i - 1) & "',"
                    '  操作员编号_In 住院费用记录.操作员编号%Type,
                    strSQL = strSQL & "'" & UserInfo.编号 & "',"
                    '  操作员姓名_In 住院费用记录.操作员姓名%Type,
                    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
                    '  退费时间_In   住院费用记录.登记时间%Type, --多张单据退费时,每张单据的退费时间相同,都是系统当前时间
                    strSQL = strSQL & "to_date('" & strDelDate & "','yyyy-mm-dd hh24:mi:ss'),"
                    '  门诊退费_In   Number := 0,
                    strSQL = strSQL & "" & 1 & ","
                    '  入院科室id_In 病人预交记录.科室id%Type,
                    strSQL = strSQL & "Null,"
                    '  主页id_In     住院费用记录.主页id%Type, --医保入院补充登记时才传入
                    strSQL = strSQL & "Null,"
                    '  三方退费_In   Number := 0,
                    strSQL = strSQL & "" & IIf(blnExecuteThreeSwap = True, 1, 0) & ","
                    '  结帐ID_In     住院费用记录.结帐id%Type)
                    strSQL = strSQL & "" & lng冲销ID & ")"
                    Call zlDatabase.ExecuteProcedure(strSQL, "三方卡结算")
                    
                    strDelNOs = strDelNOs & IIf(strDelNOs = "", "", ",") & cllPro(i)(0)
                    mstrUsedBills = mstrUsedBills & "," & Val(cllPro(i)(1))
                End If
            End If
        Next
     End If
     
    If intInsure <> 0 And bln退费后打印回单 And InStr(1, mstrPrivs, ";医保退费回单;") > 0 Then
        '问题:35248
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_4", Me, "NO=" & strNos, 2)
    End If
    ExecuteDelBill = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    
    If blnTrans Then
        '医保和HIS不是同一个事务,HIS事务失败,但医保可能已上传,所以需要调"取消交易"接口
        If blnTransMedicare Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, False, intInsure)
    End If
    
    If Err.Number <> 0 Then
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
    
    '中断提示,不打印，重新退费后再打印或自己选择重打
    If strDelNOs <> "" Then
        MsgBox "单据[" & strNos & "]退费失败。但是，单据[" & strDelNOs & "]已成功退费。" & vbCrLf & _
            "单据未打印，请对执行失败的单据重新退费！", vbInformation, gstrSysName
    End If
    Exit Function
End Function

Private Function ExecuteThreeSwap(lngBalance As Long, lng冲销ID As Long, Optional ByRef str交易流水号 As String, Optional ByRef str交易说明 As String) As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset, strBalanceIDs As String, rsTotal As ADODB.Recordset
    Dim dblMoney As Double, strAll As String, strDetail() As String, strItem() As String, strCardNo As String
    Dim i As Integer, lngCardID As Long
    If mobjSquare Is Nothing Then Set mobjSquare = gobjSquare.objSquareCard
    If mobjSquare Is Nothing Then Exit Function
    strSQL = _
        "Select 摘要" & vbNewLine & _
        "    From 病人预交记录" & vbNewLine & _
        "    Where 结算方式 Is Null And 记录性质 = 3 And 记录状态 = 2 And 结帐id = [1]"
   
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng冲销ID)
    
    If rsTemp.RecordCount = 0 Then Exit Function
    strAll = Nvl(rsTemp!摘要)
    If strAll = "" Then Exit Function
    
    strDetail = Split(strAll, "|")
    For i = 0 To UBound(strDetail)
        If strDetail(i) <> "" Then
            strItem = Split(strDetail(i), ",")
            If Val(strItem(0)) = 1 Then
                lngCardID = Val(strItem(1))
                dblMoney = -1 * Val(strItem(2))
                strSQL = "Select Distinct a.结帐id" & vbNewLine & _
                            "From 门诊费用记录 A" & vbNewLine & _
                            "Where a.No In (Select Distinct a.No From 门诊费用记录 A Where Mod(a.记录性质, 10) = 1 And a.结帐id = [1]) And Mod(a.记录性质, 10) = 1 And" & vbNewLine & _
                            "      a.记录状态 <> 0"
                strSQL = "Select Min(结帐ID) As 结帐ID,Min(卡号) As 卡号 From 病人预交记录 Where 结帐ID IN (" & strSQL & ") And 卡类别ID = [2]"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng冲销ID, lngCardID)
                strBalanceIDs = "3|" & Val(Nvl(rsTemp!结帐ID))
                If mobjSquare.zlReturnCheck(Me, mlngModule, lngCardID, False, Nvl(rsTemp!卡号), _
                    strBalanceIDs, dblMoney, str交易流水号, str交易说明, "3|" & lng冲销ID) = False Then Exit Function
                If mobjSquare.zlReturnMoney(Me, mlngModule, lngCardID, False, Nvl(rsTemp!卡号), _
                    strBalanceIDs, dblMoney, str交易流水号, str交易说明, "3|" & lng冲销ID) = False Then Exit Function
            End If
        End If
    Next i
    
    ExecuteThreeSwap = True
End Function

Public Function GetBalanceSet() As ADODB.Recordset
'功能：返回一个结算记录集对象
    Dim rsTmp As New ADODB.Recordset
       
    rsTmp.Fields.Append "单据序号", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "结算方式", adVarChar, 20, adFldIsNullable
    rsTmp.Fields.Append "结算金额", adCurrency, , adFldIsNullable
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    Set GetBalanceSet = rsTmp
End Function

Public Function Get实收金额(ByVal strNO As String) As Currency
    Dim i As Long, cur金额 As Currency
    With vsFee
        cur金额 = 0
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("单据号")) = strNO Then
                cur金额 = cur金额 + Val(.TextMatrix(i, .ColIndex("实收金额")))
            End If
        Next
        Get实收金额 = cur金额
    End With
End Function
Private Function ExecuteWirteOff(strDELDae As String, ByVal cllDel As Variant) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行门诊记帐销帐
    '编制:刘兴洪
    '日期:2011-02-25 10:22:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strSQL As String
    Dim cllPro As Collection
    Set cllPro = New Collection
    For i = 1 To cllDel.Count
        'Zl_门诊转住院_记帐转出
        strSQL = "Zl_门诊转住院_记帐转出("
        '  No_In         住院费用记录.NO%Type,
        strSQL = strSQL & "'" & cllDel(i)(0) & "',"
        '  操作员编号_In 住院费用记录.操作员编号%Type,
        strSQL = strSQL & "'" & UserInfo.编号 & "',"
        '  操作员姓名_In 住院费用记录.操作员姓名%Type,
        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
        '  退费时间_In   住院费用记录.发生时间%Type
        strSQL = strSQL & "To_Date('" & strDELDae & "','yyyy-mm-dd hh24:mi:ss'),"
        '   门诊销帐_In   Number := 0
        '   --门诊销帐_In:0-门诊转住院立即销帐;1-门诊记帐退费模式
        strSQL = strSQL & "1)"
        zlAddArray cllPro, strSQL
    Next
    On Error GoTo errHandle
    zlExecuteProcedureArrAy cllPro, Me.Caption
    ExecuteWirteOff = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function SaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:销帐或退费
    '返回:退费或销帐成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-02-23 11:21:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, lng险类 As Long, lng结帐ID As Long
    Dim strOutNos As String, strTemp As String, strDelDate As String
    Dim m As Long, i As Long, blnHaveData As Boolean, blnPrintList As Boolean '是否打印清单
    Dim cllDelNO As Collection, strDelNOs As String, lngRow As Long, strNO As String
    Dim lng病人ID As Long
    
    strDelDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    blnPrintList = False
    If InStr(mstrPrivs, ";打印清单;") > 0 And mint性质 = 1 Then
        Select Case mint收费清单    '0-不打印,1-要打印,2-选择是否打印
        Case 2
             If MsgBox("要打印收费清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                blnPrintList = True
             End If
        Case 1
            blnPrintList = True
        End Select
    End If
    mstrUsedBills = ""
    With vsFee
        If .Rows <= 1 Then Exit Function
        If .Cols <= 1 Then Exit Function
        Set cllDelNO = New Collection
        strTemp = ""
        For lngRow = 1 To .Rows - 1
            '销帐单据
            strNO = Trim(.TextMatrix(lngRow, .ColIndex("单据号")))
            If CheckBillExistReplenishData(1, , strNO) And mint性质 = 1 Then
                MsgBox "选择的单据存在补充结算记录，无法进行退费！", vbInformation, gstrSysName
                Exit Function
            End If
            If GetVsGridBoolColVal(vsFee, lngRow, .ColIndex(mstr标志)) _
                And strNO <> "" And InStr(1, "," & strTemp & ",", "," & strNO & ",") = 0 Then
                lng险类 = Val(.TextMatrix(lngRow, .ColIndex("险类")))
                lng结帐ID = Val(.TextMatrix(lngRow, .ColIndex("结帐ID")))
                strOutNos = ""
'                If CheckMulitBillValied(strNo, lng险类, strOutNos) = False Then
'                    Exit Function
'                End If
                If lng险类 <> 0 And IsYBSingle(strNO, lng险类) = False Then
                    If CheckInsureAll(lng结帐ID) = False Then
                        MsgBox "选择的单据存在其他未退费单据，无法进行退费！", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
                
                blnHaveData = False
                For i = 1 To cllDelNO.Count
                    If cllDelNO(i)(0) = strNO Then
                        blnHaveData = True: Exit For
                    End If
                    If InStr(1, "," & cllDelNO(i)(1) & ",", "," & strNO & ",") > 0 Then
                        blnHaveData = True: Exit For
                    End If
                    If lng险类 <> 0 Then
                        If IsYBSingle(strNO, lng险类) = False Then
                            If Val(cllDelNO(i)(3)) = lng结帐ID Then
                                blnHaveData = True: Exit For
                            End If
                        End If
                    End If
                Next
                If blnHaveData = False Then
                    '加入销帐单据
                    cllDelNO.Add Array(strNO, strOutNos, lng险类, lng结帐ID)
                End If
                strTemp = strTemp & "," & strNO & "," & strOutNos

            End If
        Next
    End With
    '执行具体销帐或退费操作
    If cllDelNO.Count = 0 Then
        MsgBox "注意:" & vbCrLf & "    没有选择一张需要进行退费或销帐的单据,请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    '退费
    strDelNOs = ""
    If mint性质 = 2 Then
        If ExecuteWirteOff(strDelDate, cllDelNO) = False Then Exit Function
    Else
        For i = 1 To cllDelNO.Count
            If ExecuteDelBill(strDelDate, IIf(cllDelNO(i)(1) <> "", cllDelNO(i)(1), cllDelNO(i)(0)), Val(cllDelNO(i)(2)), Val(cllDelNO(i)(2))) = False Then
                    Exit Function
            End If
            strDelNOs = strDelNOs & "," & IIf(cllDelNO(i)(1) <> "", cllDelNO(i)(1), cllDelNO(i)(0))
        Next
    End If
    If strDelNOs <> "" Then strDelNOs = Mid(strDelNOs, 2)
    '打印费用清单
    If blnPrintList And mint性质 = 1 Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me, "NO=" & "'" & Replace(strDelNOs, ",", "','") & "'", "药品单位=" & IIf(mbln药房单位, 1, 0), 2)
    End If
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub SetLocaleNO(ByVal str单据 As String, ByVal strNO As String, ByVal blnSelect As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置指定的NO
    '编制:刘兴洪
    '日期:2011-02-09 14:56:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    With vsFee
        For lngRow = 1 To .Rows - 1
            If Trim(.TextMatrix(lngRow, .ColIndex("单据号"))) = strNO Then
                    .TextMatrix(lngRow, .ColIndex("选择")) = IIf(blnSelect, -1, 0)
            End If
        Next
    End With
End Sub
Private Function CheckIsInput(ByVal lngRow As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否允许输入更改
    '入参:lngRow-指定的行
    '出参:
    '返回:有效,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-02-09 15:04:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intInsure As Integer, strNO As String, i As Long, strTmp As String
    Dim strBalanceType As String, arrBalanceType As Variant
    Dim lng病人ID As Long, str单据 As String
    
    If mrsInfo Is Nothing Then Exit Function
    If mrsInfo.State <> 1 Then Exit Function
    lng病人ID = Val(Nvl(mrsInfo!病人ID))
    With vsFee
            intInsure = Val(.TextMatrix(lngRow, .ColIndex("险类")))
            strNO = .TextMatrix(lngRow, .ColIndex("单据号"))
            str单据 = .TextMatrix(lngRow, .ColIndex("单据"))
            If intInsure > 0 And str单据 = "收费单" Then
                If Not gclsInsure.GetCapability(support门诊结算作废, lng病人ID, intInsure) Then
                    stbThis.Panels(2).Text = "单据[" & strNO & "]的病人险类不支持门诊结算作废,此行不允许选择转入!"
                    Exit Function
                Else
                    '再判断该单据的每种结算方式是否支持,正常退费时,可以退为指定结算方式,此处简化规则为不允许退费
                    strTmp = GetBalanceType(strNO)
                    If strTmp <> "" Then
                        arrBalanceType = Split(strTmp, ",")
                        For i = 0 To UBound(arrBalanceType)
                            strBalanceType = arrBalanceType(i)
                            If Not gclsInsure.GetCapability(support门诊结算作废, lng病人ID, intInsure, strBalanceType) Then
                                stbThis.Panels(2).Text = "单据[" & strNO & "]的病人险类不支持" & strBalanceType & "作废,此行不允许选择转入!"
                                Exit Function
                            End If
                        Next
                    End If
                End If
            End If
    End With
    CheckIsInput = True
End Function
Private Function SetRowSelected(ByVal lngRow As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置一行的选择状态
    '       如果是多张单据中的一张,则还需同时设置多张中的其它单据
    '编制:刘兴洪
    '日期:2011-02-09 14:50:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intInsure As Integer, strNO As String, i As Long, strTmp As String
    Dim blnSelect As Boolean, lng病人ID As Long, str单据 As String
    lng病人ID = 0
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then
            lng病人ID = Val(Nvl(mrsInfo!病人ID))
        End If
    End If
    With vsFee
        intInsure = Val(.TextMatrix(lngRow, .ColIndex("险类")))
        blnSelect = GetVsGridBoolColVal(vsFee, lngRow, .ColIndex(mstr标志))
        str单据 = Trim(.TextMatrix(lngRow, .ColIndex("单据")))
        If intInsure > 0 And str单据 = "收费单" Then '全部选择或取消
            If Not IsYBSingle(.TextMatrix(lngRow, .ColIndex("单据号")), intInsure) Then
                If Not SetMultiOther(lngRow, blnSelect, intInsure) Then Exit Function
            End If
        Else '现金病人需要处理多单据收费情况
            If Not SetMultiOther(lngRow, blnSelect, intInsure) Then Exit Function
        End If
    End With
    SetRowSelected = True
End Function

Private Function CheckInsureAll(lngBalance As Long) As Boolean
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim i As Long, blnFound As Boolean
    strSQL = "Select Distinct a.No" & vbNewLine & _
            "From 门诊费用记录 A, 门诊费用记录 B" & vbNewLine & _
            "Where b.结帐id = [1] And a.No = b.No And Mod(a.记录性质,10) = Mod(b.记录性质,10)" & vbNewLine & _
            "Group By a.No" & vbNewLine & _
            "Having Sum(a.实收金额) <> 0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngBalance)
    Do While Not rsTmp.EOF
        blnFound = False
        With vsFee
            For i = 1 To .Rows - 1
                If GetVsGridBoolColVal(vsFee, i, .ColIndex(mstr标志)) Then
                    If Trim(.TextMatrix(i, .ColIndex("单据号"))) = Trim(rsTmp!NO) Then blnFound = True: Exit For
                End If
            Next i
            If blnFound = False Then
                CheckInsureAll = False
                Exit Function
            End If
        End With
        rsTmp.MoveNext
    Loop
    CheckInsureAll = True
End Function

Private Function GetBalanceType(ByVal strNO As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取一张单据中的医保结算方式串
    '返回:医保结算方式串
    '编制:刘兴洪
    '日期:2011-02-09 15:01:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
     Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim i As Long
    On Error GoTo errH
    strSQL = "Select A.结算方式 From 病人预交记录 A, 结算方式 B" & vbNewLine & _
            "Where A.结算方式 = B.名称 And B.性质 In (3, 4) And A.NO = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    
    For i = 1 To rsTmp.RecordCount
        GetBalanceType = GetBalanceType & "," & rsTmp!结算方式
        rsTmp.MoveNext
    Next
    GetBalanceType = Mid(GetBalanceType, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckAllTurn(ByVal strNO As String) As Boolean
    Dim strSQL As String, rsData As ADODB.Recordset
    strSQL = "Select 1" & vbNewLine & _
            " From 病人预交记录 A," & vbNewLine & _
            "     (Select Distinct 结帐id" & vbNewLine & _
            "       From 门诊费用记录" & vbNewLine & _
            "       Where NO In (Select Distinct NO" & vbNewLine & _
            "                    From 门诊费用记录" & vbNewLine & _
            "                    Where 结帐id In" & vbNewLine & _
            "                          (Select 结帐id" & vbNewLine & _
            "                           From 病人预交记录" & vbNewLine & _
            "                           Where 结算序号 In (Select b.结算序号" & vbNewLine & _
            "                                          From 门诊费用记录 A, 病人预交记录 B" & vbNewLine & _
            "                                          Where a.No = [1] And a.记录性质 = 1 And a.记录状态 <> 0 And a.结帐id = b.结帐id))) And" & vbNewLine & _
            "             记录性质 = 1 And 记录状态 <> 0) B" & vbNewLine & _
            " Where a.结帐id = b.结帐id And a.记录性质 = 3 And (Exists (Select 1 From 医疗卡类别 Where ID = a.卡类别id And 是否全退 = 1) Or Exists" & vbNewLine & _
            "       (Select 1 From 消费卡类别目录 Where 编号 = a.结算卡序号 And 是否全退 = 1))" & vbNewLine & _
            " Group By 结算方式" & vbNewLine & _
            " Having Sum(冲预交) <> 0"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsData.EOF Then
        CheckAllTurn = False
    Else
        CheckAllTurn = True
    End If
End Function

Private Function SetMultiOther(ByVal lngRow As Long, blnSelect As Boolean, intInsure As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:多张单据整体选择或取消
    '       如果医保多张单据要求整体退费,选择其中一张时,全选多张,取消时全取消
    '入参:lngRow-当前行
    '        blnSelect-是否选中
    '        intInsure-险类
    '返回:
    '编制:刘兴洪
    '日期:2011-02-09 15:41:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, k As Long, strNO As String, strTmp As String
    Dim strBalanceType As String, arrBalanceType As Variant
    Dim lng病人ID As Long, str单据 As String, blnAllTurn As Boolean
    lng病人ID = 0
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then
            lng病人ID = Val(Nvl(mrsInfo!病人ID))
        End If
    End If
    With vsFee
        str单据 = Trim(.TextMatrix(lngRow, .ColIndex("单据")))
        If intInsure = 0 Then
            If CheckAllTurn(.TextMatrix(lngRow, .ColIndex("单据号"))) = True Then
                blnAllTurn = True
            Else
                blnAllTurn = False
            End If
            If mblnMultiBalance Or blnAllTurn Then     '   多单据,多种结算方式
                '33635:原因是多单据且多种结算方式,不能部分退
                strNO = ""
                For k = 1 To .Rows - 1
                      If .TextMatrix(k, .ColIndex("结帐ID")) = .TextMatrix(lngRow, .ColIndex("结帐ID")) _
                        And Trim(.TextMatrix(lngRow, .ColIndex("结帐ID"))) <> "" _
                        And .TextMatrix(k, .ColIndex("单据")) = str单据 Then
                          If InStr(1, "," & strNO & ",", "," & .TextMatrix(k, .ColIndex("单据号")) & ",") = 0 Then
                                strNO = strNO & "," & .TextMatrix(k, .ColIndex("单据号"))
                          End If
                      End If
                Next
                If strNO <> "" Then strNO = Mid(strNO, 2)
                If InStr(1, strNO, ",") > 0 Then    '证明为多单据
                    '一院要求,只要是多单据结算的,在转时,都必须全转
                    'If CheckSingleBalance(strNo) = False Then    '是多种结算方式,则不允许退费,'全选
                        For k = 1 To .Rows - 1
                              If .TextMatrix(k, .ColIndex("结帐ID")) = .TextMatrix(lngRow, .ColIndex("结帐ID")) _
                                  And Trim(.TextMatrix(lngRow, .ColIndex("结帐ID"))) <> "" _
                                   And .TextMatrix(k, .ColIndex("单据")) = str单据 Then
                                    .TextMatrix(k, .ColIndex(mstr标志)) = IIf(blnSelect, -1, 0)
                              End If
                        Next
                    'End If
                End If
            End If
            '检查是否存在消费卡的结算,如果存在,现不支持这部分数据的处理
            If strNO = "" Then strNO = .TextMatrix(lngRow, .ColIndex("单据号"))
'            If str单据 = "收费单" Then
'                If zlIsExistsSquareCard(strNO) Then
'                    stbThis.Panels(2).Text = "暂不支持对消费卡数据的门诊费用转住院费用!"
'                    For k = 1 To .Rows - 1
'                          If .TextMatrix(k, .ColIndex("结帐ID")) = .TextMatrix(lngRow, .ColIndex("结帐ID")) And Trim(.TextMatrix(lngRow, .ColIndex("结帐ID"))) <> "" Then
'                                .TextMatrix(k, .ColIndex(mstr标志)) = 0
'                          End If
'                    Next
'                End If
'            End If
            '检查是否存在消费卡,如果多单据中存在消费卡,也必须全选
            SetMultiOther = True
            Exit Function
        End If
        If IsYBSingle(vsFee.TextMatrix(lngRow, .ColIndex("单据号")), intInsure) Then SetMultiOther = True: Exit Function
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("结帐ID")) = .TextMatrix(lngRow, .ColIndex("结帐ID")) _
                And i <> lngRow And .TextMatrix(i, .ColIndex("单据")) = str单据 Then
                If GetVsGridBoolColVal(vsFee, i, .ColIndex(mstr标志)) <> GetVsGridBoolColVal(vsFee, lngRow, .ColIndex(mstr标志)) Then
                   If intInsure <> 0 And str单据 = "收费单" And blnSelect Then
                        strNO = .TextMatrix(i, .ColIndex("单据号"))
                        '判断该单据的每种结算方式是否支持,正常退费时,可以退为指定结算方式,此处简化规则为不允许退费
                         strTmp = GetBalanceType(strNO)
                         If strTmp <> "" Then
                             arrBalanceType = Split(strTmp, ",")
                             For j = 0 To UBound(arrBalanceType)
                                 strBalanceType = arrBalanceType(j)
                                 If Not gclsInsure.GetCapability(support门诊结算作废, lng病人ID, intInsure, strBalanceType) Then
                                     stbThis.Panels(2).Text = "单据[" & strNO & "]的病人险类不支持" & strBalanceType & "作废,此行不允许选择转入!"
                                     For k = 1 To .Rows - 1
                                        If .TextMatrix(k, .ColIndex("结帐ID")) = .TextMatrix(i, .ColIndex("结帐ID")) _
                                            And .TextMatrix(k, .ColIndex("单据")) = .TextMatrix(i, .ColIndex("单据")) Then
                                            .TextMatrix(k, .ColIndex(mstr标志)) = 0
                                        End If
                                     Next
                                     Exit Function
                                 End If
                             Next
                         End If
                    End If
                    .TextMatrix(i, .ColIndex(mstr标志)) = IIf(blnSelect, -1, 0)
                End If
            End If
        Next
    End With
    SetMultiOther = True
End Function
Private Function IsCheckSelNo() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查单据是否存在选择
    '返回:选中,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-02-23 15:41:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With vsFee
        For i = 1 To .Rows - 1
            If GetVsGridBoolColVal(vsFee, i, .ColIndex(mstr标志)) Then
                IsCheckSelNo = True: Exit Function
            End If
        Next
    End With
    IsCheckSelNo = False
End Function
Private Function CheckSingleBalance(ByVal strNO As String) As Boolean
'功能：判断指定单据中是否只有一种非医保结算方式(冲预交除外)
'       :strNO(格式为"E01,E02"):问题:34035
'参数：
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strNO = Replace(strNO, "'", "")
    CheckSingleBalance = True
    
    strSQL = "" & _
    " Select /*+ rule */ Count(Distinct A.结算方式) num" & vbNewLine & _
    " From 病人预交记录 A, 结算方式 B,Table( f_Str2list([1])) J" & vbNewLine & _
    " Where   A.记录性质 = 3 And A.记录状态 In (1, 3) " & _
    "           And A.结算方式 = B.名称 And B.性质 In (1, 2)  And A.NO = J.Column_Value"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNO)
    If rsTmp!Num > 1 Then CheckSingleBalance = False
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function zlIsExistsSquareCard(ByVal strNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查该单据是否为卡结算单据
    '入参:strNos-单据号(可以为多张,用逗号分离)
    '出参:
    '返回:存在,则返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-11 12:04:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, strNoIns As String
    
    On Error GoTo errHandle
    
    strNoIns = Replace(strNos, "'", "")
    strSQL = "" & _
    "   Select /*+ rule */ A.ID As 卡结算id " & _
    "   From 病人卡结算记录 A, 病人预交记录 B,Table( f_Str2list([1])) J " & _
    "   Where A.结算id = B.ID and B.记录性质=3 And B.NO = J.Column_Value And Rownum = 1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查收费单是否存在刷卡记录", strNoIns)
    zlIsExistsSquareCard = rsTemp.EOF = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub vsHistory_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsHistory, Me.Caption, IIf(mint性质 = 1, "历史退费列表", "历史销帐列表"), True
End Sub
Private Sub vsHistory_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsHistory, Me.Caption, IIf(mint性质 = 1, "历史退费列表", "历史销帐列表"), True
End Sub
Private Sub SetBlanceShow()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示结算方式
    '入参:blnAllSel-选择所有的单据
    '编制:刘兴洪
    '日期:2011-02-23 14:54:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strBalance As String, lngRow As Long, i As Long, str结算 As String
    Dim bln全选 As Boolean, bln未选 As Boolean, rsTmp As ADODB.Recordset
    Dim strFilter As String, bln退款 As Boolean, strSQL As String
    Dim strSelNos As String, strNO As String, intCol As Integer
    If mint性质 = 2 Then Exit Sub
    With vsFee
        bln全选 = True: bln未选 = True
        For lngRow = 1 To .Rows - 1
            strNO = Trim(.TextMatrix(lngRow, .ColIndex("单据号")))
            If GetVsGridBoolColVal(vsFee, lngRow, .ColIndex(mstr标志)) Then
                If InStr(1, strSelNos & ",", "," & strNO & ",") = 0 Then
                    strSelNos = strSelNos & "," & strNO
                    bln未选 = False
                End If
            End If
             If InStr(1, strSelNos & ",", "," & strNO & ",") = 0 Then bln全选 = False
        Next
    End With
    If strSelNos <> "" Then strSelNos = Mid(strSelNos, 2)
    bln退款 = False
    '显示所有选择的单据的结算方式之和
    If Not mrsBalance Is Nothing Then
        If bln全选 Or bln未选 Then
            mrsBalance.Filter = 0
            If bln全选 Then bln退款 = True
        Else
'            strFilter = Replace(strSelNos, ",", "' Or NO='")
'            strFilter = " NO='" & strFilter & "'"
'            mrsBalance.Filter = strFilter
            bln退款 = True
        End If
        If SetPicBack(strSelNos) = True Then
            txtSum.Text = InitPatialBalance(strSelNos)
        Else
            Call InitBlanceData(strSelNos)
        End If
        mcur误差 = 0
        If Val(cboStyle.ItemData(cboStyle.ListIndex)) = 1 Then
            mcur误差 = Val(txtSum.Text) - CentMoney(Val(txtSum.Text))
            If mcur误差 <> 0 Then
            With mrsBalance
                .AddNew
                !结算方式 = "误差费"
                !性质 = 1
                !应付款 = "0"
                !金额 = Format(mcur误差, "0.00")
                !摘要 = ""
                !结算号码 = ""
                .Update
            End With
            End If
            txtSum.Text = Format(txtSum.Text - mcur误差, "0.00")
        Else
            mcur误差 = 0
        End If
        
        mrsBalanceBak.Filter = "金额 <> 0"
        mrsBalanceBak.Sort = "性质,应付款,结算方式"
        mrsBalance.Sort = "性质,应付款,结算方式"
        vsBalance.Redraw = flexRDNone
        vsBalance.Clear 1
        vsBalance.Cols = 1
        
        If Not mrsBalanceBak.EOF Then
            For i = 1 To mrsBalanceBak.RecordCount
                If Nvl(mrsBalanceBak!结算方式, "冲预交") <> strBalance Then
                    strBalance = Nvl(mrsBalanceBak!结算方式, "冲预交")
                    vsBalance.Cols = vsBalance.Cols + 2
                    vsBalance.ColAlignment(vsBalance.Cols - 2) = 7
                    vsBalance.ColAlignment(vsBalance.Cols - 1) = 1
                End If
                If mrsBalanceBak!性质 <> 1 Then
                    vsBalance.Cell(flexcpFontBold, 0, vsBalance.Cols - 1, 0, vsBalance.Cols - 2) = True  '粗体
                    vsBalance.Cell(flexcpForeColor, 0, vsBalance.Cols - 1, 0, vsBalance.Cols - 2) = vbBlue
                ElseIf bln退款 Then
                    vsBalance.Cell(flexcpFontBold, 0, vsBalance.Cols - 1, 0, vsBalance.Cols - 2) = True  '粗体
                    vsBalance.Cell(flexcpForeColor, 0, vsBalance.Cols - 1, 0, vsBalance.Cols - 2) = vbBlue  '红色:退款
                End If
                vsBalance.TextMatrix(0, vsBalance.Cols - 2) = strBalance & ":"
                vsBalance.TextMatrix(0, vsBalance.Cols - 1) = _
                    Val(vsBalance.TextMatrix(0, vsBalance.Cols - 1)) + Nvl(mrsBalanceBak!金额, 0)
                    '多单据使用多种结算时,单笔结算金额看没有进行分币处理,所以不能用format取两位数
                'vsBalance.ColData(vsBalance.Cols - 2) = "摘要:" & mrsBalanceBak!摘要
                vsBalance.ColData(vsBalance.Cols - 1) = "结算号码:" & mrsBalanceBak!结算号码
                mrsBalanceBak.MoveNext
            Next
        End If
        
        intCol = 0
        strBalance = ""
        If Not mrsBalance.EOF Then
            For i = 1 To mrsBalance.RecordCount
                If Nvl(mrsBalance!结算方式, "冲预交") <> strBalance Then
                    strBalance = Nvl(mrsBalance!结算方式, "冲预交")
                    intCol = intCol + 2
                    vsBalance.ColAlignment(intCol - 1) = 7
                    vsBalance.ColAlignment(intCol) = 1
                End If
                If mrsBalance!性质 <> 1 Then
                    vsBalance.Cell(flexcpFontBold, 1, intCol, 1, intCol - 1) = True '粗体
                    vsBalance.Cell(flexcpForeColor, 1, intCol, 1, intCol - 1) = IIf(bln退款, vbRed, vbBlue) '红色
                ElseIf bln退款 Then
                    vsBalance.Cell(flexcpFontBold, 1, intCol, 1, intCol - 1) = True '粗体
                    vsBalance.Cell(flexcpForeColor, 1, intCol, 1, intCol - 1) = vbRed '红色:退款
                End If
                vsBalance.TextMatrix(1, intCol - 1) = strBalance & ":"
                If Nvl(mrsBalance!结算方式) = "误差费" Then
                    vsBalance.TextMatrix(1, intCol) = _
                        Format(Val(vsBalance.TextMatrix(1, intCol)) + Nvl(mrsBalance!金额, 0), "0.00")
                Else
                    vsBalance.TextMatrix(1, intCol) = _
                        Val(vsBalance.TextMatrix(1, intCol)) + Nvl(mrsBalance!金额, 0)
                End If
                vsBalance.ColData(intCol) = "结算号码:" & mrsBalance!结算号码
                mrsBalance.MoveNext
            Next
        End If
        If strSelNos = "" Then
            For i = 1 To vsBalance.Cols - 1
                vsBalance.TextMatrix(1, i) = ""
            Next i
        End If
        Call vsBalance.AutoSize(0, vsBalance.Cols - 1)
        vsBalance.Row = vsBalance.FixedRows
        If vsBalance.Cols <> 1 Then vsBalance.Col = vsBalance.FixedCols
        'vsBalance.TextMatrix(0, 0) = IIf(bln退款, "退款结算", "收款结算")
        vsBalance.Redraw = flexRDDirect
    End If
End Sub
Public Sub CalcSUMMony()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查单据中输入的负数数量及退回科室是否正确
    '编制:刘兴洪
    '日期:2011-03-11 18:09:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, cur金额 As Currency
    With vsFee
        cur金额 = 0
        For i = .FixedRows To .Rows - 1
            If GetVsGridBoolColVal(vsFee, i, .ColIndex(mstr标志)) Then
                cur金额 = cur金额 + Val(.TextMatrix(i, .ColIndex("实收金额")))
            End If
        Next
        lblSum.Caption = "当前转出合计:" & Format(cur金额, "###0.00;-###0.00;0.00;0.00")
        mcur合计 = cur金额
    End With
End Sub
Public Sub StatusShowBillSum()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查单据中输入的负数数量及退回科室是否正确
    '编制:刘兴洪
    '日期:2011-03-11 18:09:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, cur金额 As Currency, dbl发票金额 As Double, strNO As String, str发票号 As String
    Dim strTemp As String
    
    With vsFee
        strTemp = "": dbl发票金额 = 0: cur金额 = 0
        If Not (.Row > .Rows - 1 Or .Row < 1) Then
            strNO = .TextMatrix(.Row, .ColIndex("单据号"))
            str发票号 = .TextMatrix(.Row, .ColIndex("票据号"))
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("单据号")) = strNO Then
                        cur金额 = cur金额 + Val(.TextMatrix(i, .ColIndex("实收金额")))
                End If
                If .TextMatrix(i, .ColIndex("票据号")) = str发票号 Then
                        dbl发票金额 = dbl发票金额 + Val(.TextMatrix(i, .ColIndex("实收金额")))
                End If
            Next
            strTemp = "单据(" & strNO & ")合计:" & Format(cur金额, "###0.00;-###0.00;0.00;0.00")
            strTemp = strTemp & "  发票(" & str发票号 & ")合计:" & Format(dbl发票金额, "###0.00;-###0.00;0.00;0.00")
        End If
        stbThis.Panels(2).Text = strTemp
    End With
End Sub

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
    Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtPatient)
    Set objCard = IDKind.GetfaultCard
    Set gobjSquare.objDefaultCard = objCard
    If IDKind.Cards.按缺省卡查找 And Not objCard Is Nothing Then
        gobjSquare.bln缺省卡号密文 = objCard.卡号密文规则 <> ""
        gobjSquare.int缺省卡号长度 = objCard.卡号长度
    Else
        gobjSquare.bln缺省卡号密文 = IDKind.Cards.加密显示
        gobjSquare.int缺省卡号长度 = 100
    End If
    gobjSquare.bln按缺省卡查找 = IDKind.Cards.按缺省卡查找
End Sub
Private Sub zlCreateObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建公共事件对象
    '返回: 创建成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-08-28 16:16:00
    '说明:
    '问题:54896
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '创建公共对象
    Err = 0: On Error Resume Next
    If mobjICCard Is Nothing Then
        Set mobjICCard = New clsICCard
        Call mobjICCard.SetParent(Me.hWnd)
         Set mobjICCard.gcnOracle = gcnOracle
    End If
    
    If mobjIDCard Is Nothing Then
        Set mobjIDCard = New clsIDCard
        Call mobjIDCard.SetParent(Me.hWnd)
    End If
    
End Sub
Private Sub zlCloseObject()
    '关闭相关对象
    Err = 0: On Error Resume Next
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
    End If
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
    End If
    Set mobjIDCard = Nothing
    Set mobjICCard = Nothing
End Sub


