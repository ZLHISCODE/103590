VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "CODEJO~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDockOutAdvice 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7695
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   Begin VB.Timer timHide 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   7005
      Top             =   690
   End
   Begin VB.PictureBox PicAdviceDetail 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFEFEF&
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   135
      ScaleHeight     =   2745
      ScaleWidth      =   2775
      TabIndex        =   10
      Top             =   5760
      Visible         =   0   'False
      Width           =   2800
      Begin VSFlex8Ctl.VSFlexGrid vsfAdivceDetail 
         Height          =   2475
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   2745
         _cx             =   4851
         _cy             =   4366
         Appearance      =   2
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   16773103
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16773103
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16773103
         BackColorAlternate=   16773103
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDockOutAdvice.frx":0000
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
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   1
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   -1  'True
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
         WallPaper       =   "frmDockOutAdvice.frx":003E
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4890
      Left            =   120
      ScaleHeight     =   4890
      ScaleWidth      =   6570
      TabIndex        =   0
      Top             =   210
      Width           =   6570
      Begin VB.Frame fraMore 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   5250
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   225
         Begin VB.Image imgMore 
            Height          =   225
            Left            =   0
            Picture         =   "frmDockOutAdvice.frx":12970
            Top             =   0
            Width           =   225
         End
      End
      Begin VB.Frame fraColSel 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5460
         TabIndex        =   3
         Top             =   255
         Width           =   195
         Begin VB.Image imgColSel 
            Height          =   195
            Left            =   0
            Picture         =   "frmDockOutAdvice.frx":12D71
            ToolTipText     =   "选择需要显示的列(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
      Begin VB.Frame fraAdviceUD 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   45
         Left            =   0
         MousePointer    =   7  'Size N S
         TabIndex        =   4
         Top             =   3720
         Width           =   6975
      End
      Begin VB.Frame fraHide 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   75
         Left            =   6150
         TabIndex        =   2
         ToolTipText     =   "鼠标停留时,过滤条件栏会自动显示"
         Top             =   135
         Visible         =   0   'False
         Width           =   285
      End
      Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
         Height          =   3630
         Left            =   0
         TabIndex        =   5
         Top             =   45
         Width           =   5265
         _cx             =   9287
         _cy             =   6403
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         MouseIcon       =   "frmDockOutAdvice.frx":132BF
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16444122
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   2
         RowHeightMin    =   250
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDockOutAdvice.frx":14C51
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   -1  'True
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
         Begin VB.PictureBox pictmp 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   480
            ScaleHeight     =   240
            ScaleWidth      =   480
            TabIndex        =   6
            Top             =   1320
            Visible         =   0   'False
            Width           =   480
         End
      End
      Begin XtremeSuiteControls.TabControl tbcAppend 
         Height          =   1500
         Left            =   120
         TabIndex        =   7
         Top             =   3840
         Width           =   3285
         _Version        =   589884
         _ExtentX        =   5794
         _ExtentY        =   1482
         _StockProps     =   64
      End
      Begin VSFlex8Ctl.VSFlexGrid vsColumn 
         Height          =   2940
         Left            =   5400
         TabIndex        =   8
         Top             =   675
         Visible         =   0   'False
         Width           =   1470
         _cx             =   2593
         _cy             =   5186
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   8421504
         ForeColorFixed  =   16777215
         BackColorSel    =   14737632
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmDockOutAdvice.frx":14CEC
         ScrollTrack     =   -1  'True
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
         Editable        =   2
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
      Begin XtremeCommandBars.CommandBars cbsSub 
         Left            =   5640
         Top             =   0
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin XtremeSuiteControls.TabControl tbcMain 
      Bindings        =   "frmDockOutAdvice.frx":14D3A
      Height          =   435
      Left            =   7050
      TabIndex        =   9
      Top             =   60
      Width           =   390
      _Version        =   589884
      _ExtentX        =   688
      _ExtentY        =   767
      _StockProps     =   64
   End
   Begin RichTextLib.RichTextBox rtfInfo 
      Height          =   900
      Left            =   3495
      TabIndex        =   12
      Top             =   5580
      Width           =   200
      _ExtentX        =   344
      _ExtentY        =   1588
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmDockOutAdvice.frx":14D4E
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
   Begin RichTextLib.RichTextBox rtfAppend 
      Height          =   900
      Left            =   3150
      TabIndex        =   13
      Top             =   5580
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   1588
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmDockOutAdvice.frx":14DEB
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
   Begin RichTextLib.RichTextBox rtfSche 
      Height          =   900
      Left            =   3840
      TabIndex        =   14
      Top             =   5580
      Width           =   200
      _ExtentX        =   344
      _ExtentY        =   1588
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmDockOutAdvice.frx":14E88
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
   Begin RichTextLib.RichTextBox rtfOther 
      Height          =   900
      Left            =   4185
      TabIndex        =   15
      Top             =   5580
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   1588
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmDockOutAdvice.frx":14F25
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
   Begin VSFlex8Ctl.VSFlexGrid vsAppend 
      Height          =   1155
      Left            =   5130
      TabIndex        =   16
      Top             =   5865
      Width           =   1350
      _cx             =   2381
      _cy             =   2037
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
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
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   2
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
   Begin MSComctlLib.ImageList img16 
      Left            =   7050
      Top             =   1260
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockOutAdvice.frx":14FC2
            Key             =   ""
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockOutAdvice.frx":1555C
            Key             =   ""
            Object.Tag             =   "99"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockOutAdvice.frx":15AF6
            Key             =   ""
            Object.Tag             =   "90001"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockOutAdvice.frx":16090
            Key             =   ""
            Object.Tag             =   "90002"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockOutAdvice.frx":1662A
            Key             =   ""
            Object.Tag             =   "90003"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockOutAdvice.frx":169C4
            Key             =   ""
            Object.Tag             =   "90004"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmDockOutAdvice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event Activate() '自已激活时
Public Event RequestRefresh() '要求主窗体刷新
Public Event StatusTextUpdate(ByVal Text As String) '要求更新主窗体状态栏文字
Public Event ViewEPRReport(ByVal 报告ID As Long, ByVal CanPrint As Boolean) '要求查看报告
Public Event PrintEPRReport(ByVal 报告ID As Long, ByVal Preview As Boolean) '要求打印报告
Public Event ViewPACSImage(ByVal 医嘱ID As Long) '要求进行观片
Public Event EditDiagnose(ParentForm As Object, ByVal 挂号单 As String, Succeed As Boolean) '编辑门诊诊断
Public Event CheckInfectDisease(ByVal blnOnChek As Boolean, ByVal str疾病ID As String, ByVal str诊断Id As String, ByRef blnYes As Boolean) '根据诊断检查是否书写传染病报告卡
Public Event VSKeyPress(KeyAscii As Integer)
Private WithEvents mobjReport As clsReport
Attribute mobjReport.VB_VarHelpID = -1
Private mstrBillPrint As String '当前打印的诊疗单据：报表编号、NO、记录性质
Private mobjPublicPACS As Object             'PACS业务封装公共部件

'上次刷新数据时的病人信息
Private mblnEditable As Boolean
Private mblnCanRevoke As Boolean '是否可以作废医嘱
Private mlng病人ID As Long
Private mstr挂号单 As String
Private mstr姓名 As String
Private mstr门诊号 As String
Private mlng挂号ID As Long
Private mlng前提ID As Long
Private mlng界面科室ID As Long
Private mlng挂号科室ID As Long
Private mstr前提IDs As String
Private mstr接诊医生 As String
Private mstr药品价格等级 As String '病人的药品价格等级
Private mstr卫材价格等级 As String '病人的卫材价格等级
Private mstr普通项目价格等级 As String '病人的普通项目价格等级

Private mvRegDate As Date '挂号时间,3000-01-01表示未挂号的病人
Private mblnMoved As Boolean
Private mbln产科 As Boolean
Private mbln天数 As Boolean
Private mbln指引单打印 As Boolean
Private mint险类 As Integer
Private mblnModalNew As Boolean '新开界面是否模态

Private mint场合 As Integer '调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
Private mlng路径状态 As Long    '-1-未导入，0-不符合导入条件，1-执行中，2-正常结束，3-变异结束
Private mint就诊类型 As Integer 'pt候诊 = 0；pt就诊 = 1；pt已诊 = 2；pt转诊 = 3；pt预约 = 4；pt回诊 = 5
Private mblnNotEvaluete As Boolean  '未评估时允许添加医嘱到昨天

Private WithEvents mfrmSend As frmOutAdviceSend
Attribute mfrmSend.VB_VarHelpID = -1
Private WithEvents mfrmEdit As frmOutAdviceEdit
Attribute mfrmEdit.VB_VarHelpID = -1
Private WithEvents mfrmParent As Form
Attribute mfrmParent.VB_VarHelpID = -1
Private mcbsMain As Object
Private mMainPrivs As String
Private mblnAppend As Boolean
Private mbln皮试限制 As Boolean
Private mblnAutoRead As Boolean
Private mblnAutoReadEnabled As Boolean
Private mrsDefine As ADODB.Recordset    '医嘱内容定义
Private mobjVBA As Object
Private mobjScript As clsScript

Private mlngFontSize As Long  '字体大小

Private mblnFirst As Boolean '是否首次调用
Private mlngPlugInID As Long '自动执行的插件功能ID
Private mrsPlugInBar As ADODB.Recordset '菜单样式
Private mlngPromptRow As Long    '上一次，在鼠标移动图标列显示了提示信息的行
Private mSendControl As CommandBarControl     '发送按钮
Private mblnSignVisible As Boolean  '签名功能按钮可见性
Private mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mstr自定义申请单IDs As String 'ID1,名称1|ID2,名称2···
Private mrs危急值 As ADODB.Recordset
Private mbln危急值 As Boolean '是否有处理危急值的权限
Private mlng危急值ID As Long '当前处理的危急值记录ID
'Pass
Private mobjPassMap As Object  'PASS 窗体对象映射
Private mblnPass As Boolean  'PASS权限
'查看报告
Private mblnTag As Boolean  '是否已点击查看判断
Private mbln处方预览 As Boolean  '是否已点击为处方预览链接
Private mobjFrmBloodList As Object '血液明细窗体
'本地医嘱过滤条件
Private Enum CMD_FILTER
    ID_婴儿 = 1
    ID_废止 = 2
    ID_科内 = 5
    ID_简洁 = 7
    ID_完整 = 8
    ID_全部 = 9
    ID_检查 = 10
    ID_检验 = 11
    ID_其他 = 12
    ID_医嘱全部 = 13
    ID_医嘱处方 = 14
    ID_医嘱其他 = 15
    ID_未出报告 = 16
    ID_已出报告 = 17
End Enum

Private Type FilterCond
    婴儿 As Integer
    废止 As Boolean     'true 显示作废医嘱，false 不显示作废医嘱
    科内 As Boolean
    报告 As Integer     '0-全部，1－检查，2－检验，3－其他
    显示模式 As Integer '0-简洁，1-完整
    过滤模式 As Integer '0-医嘱，3－报告
    医嘱 As Integer '0-全部、1-处方、2-其他
    未出报告 As Boolean
    已出报告 As Boolean
End Type

Private mvarCond As FilterCond
Private mblnHideFilter As Boolean

Private Enum COL医嘱清单
    '固定列
    COL_F标志 = 0
    COL_F报告 = 1
    '隐藏列
    COL_ID = 2
    COL_相关ID = COL_ID + 1
    COL_婴儿ID = COL_ID + 2
    COL_医嘱状态 = COL_ID + 3
    COL_诊疗类别 = COL_ID + 4
    COL_操作类型 = COL_ID + 5
    COL_毒理分类 = COL_ID + 6
    COL_标志 = COL_ID + 7
    
    '可见列
    COL_警示 = COL_ID + 8 'Pass
    COL_处方号 = COL_ID + 9
    COL_处方打印 = COL_ID + 10
    COL_处方预览 = COL_ID + 11
    COL_开始时间 = COL_ID + 12
    COL_并 = COL_ID + 13
    col_医嘱内容 = COL_ID + 14
    col_内容 = COL_ID + 15
    COL_皮试 = COL_ID + 16
    COL_总量 = COL_ID + 17
    COL_单量 = COL_ID + 18
    COL_天数 = COL_ID + 19
    COL_频率 = COL_ID + 20
    COL_用法 = COL_ID + 21
    COL_医生嘱托 = COL_ID + 22
    COL_执行时间 = COL_ID + 23
    COL_执行科室 = COL_ID + 24
    COL_执行性质 = COL_ID + 25
    COL_开嘱医生 = COL_ID + 26
    COL_开嘱时间 = COL_ID + 27
    COL_发送人 = COL_ID + 28
    col_发送时间 = COL_ID + 29
    COL_超量说明 = COL_ID + 30
    COL_基本药物 = COL_ID + 31
    COL_查阅状态 = COL_ID + 32
    COL_标本状态 = COL_ID + 33
    
    '隐藏列
    COL_诊疗项目ID = COL_ID + 34
    COL_试管编码 = COL_诊疗项目ID + 1
    COL_前提ID = COL_诊疗项目ID + 2
    COL_签名否 = COL_诊疗项目ID + 3
    COL_文件ID = COL_诊疗项目ID + 4
    COL_报告项 = COL_诊疗项目ID + 5 '0-无报告，1-有报告并按编辑格式打印，2-有报告并按报表格式打印。
    COL_报告ID = COL_诊疗项目ID + 6
    COL_审核状态 = COL_诊疗项目ID + 7
    COL_申请序号 = COL_诊疗项目ID + 8
    COL_高危药品 = COL_诊疗项目ID + 9
    COL_标本部位 = COL_诊疗项目ID + 10
    COL_收费细目ID = COL_诊疗项目ID + 11   'Pass
    COL_开嘱科室ID = COL_诊疗项目ID + 12
    COL_用药目的 = COL_诊疗项目ID + 13
    COL_检查报告ID = COL_诊疗项目ID + 14
    COL_处方审查状态 = COL_诊疗项目ID + 15
    COL_处方审查结果 = COL_诊疗项目ID + 16
    COL_RIS预约ID = COL_诊疗项目ID + 17
    COL_RIS报告ID = COL_诊疗项目ID + 18
    COL_LIS报告ID = COL_诊疗项目ID + 19
    COL_RIS预约状态 = COL_诊疗项目ID + 20
    col_诊疗项目名称 = COL_诊疗项目ID + 21
    COL_检查方法 = COL_诊疗项目ID + 22  '输血医嘱区分是备血还是用血
    COL_危急值ID = COL_诊疗项目ID + 23 '医嘱关和危急值关联
    COL_易跌倒 = COL_诊疗项目ID + 24 '药品至易跌倒
End Enum

Private COLPrice As New Collection
Private COLSend As New Collection
Private COLSign As New Collection

Public Sub zlDefCommandBars(ByVal frmParent As Object, ByVal cbsMain As Object, ByVal int场合 As Integer, _
                            ByRef objPlugIn As Object, ByRef objSquareCard As Object, Optional ByVal blnModalNew As Boolean)
    
    mint场合 = int场合
    Set mfrmParent = frmParent
        mblnModalNew = blnModalNew
    If Not cbsMain Is Nothing Then

        '外挂程序对象初始化
        If Not mblnFirst Then
            mblnFirst = True
            Set mcbsMain = cbsMain
            Set cbsMain.Icons = zlCommFun.GetPubIcons
            Set gobjSquareCard = objSquareCard

            If gobjPlugIn Is Nothing Then
                If Not objPlugIn Is Nothing Then
                    '由医生站传入时，在外部初始化，此处不用再调初始化接口
                    Set gobjPlugIn = objPlugIn
                Else
                    Call CreatePlugInOK(p门诊医嘱下达, mint场合)
                End If
            End If
            Call GetPlugInBar(p门诊医嘱下达, mint场合, mrsPlugInBar)

            'PASS接口初始化
            '因为几个模块可能同时使用,有可能gobjPass对象已经创建
            If gobjPass Is Nothing Then
                Set gobjPass = DynamicCreate("zlPassInterface.clsPass", "合理用药监测", True)
                If Not gobjPass Is Nothing Then
                    Call gobjPass.zlPassInit(gcnOracle, glngSys, PM_门诊医嘱清单)
                    If gobjPass.PassType = 0 Then
                        Set gobjPass = Nothing
                    Else
                        mblnPass = True
                    End If
                End If
            End If
        End If
        
        Call zlPASSMap
        If mblnPass Then
            Call gobjPass.zlPassAdviceColHidden(mobjPassMap)
        End If
        
        If mint场合 = 0 Then    '医生站调用
            Call DefCommandsOutDoctor(cbsMain)
        ElseIf mint场合 = 2 Then    '医技站调用
            Call DefCommandsTechnic(cbsMain)
        End If

        Call DefCommandPlugIn(cbsMain, mrsPlugInBar)
    End If
End Sub

Private Sub DefCommandPlugIn(ByRef cbsMain As Object, ByRef rsBar As ADODB.Recordset)
'功能：外挂部件菜单接入。
'说明：判断关键字  Auto  InTool 决定菜单样式
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim i As Long
    Dim lngTmp As Long
    
    If rsBar Is Nothing Then Exit Sub
    rsBar.Filter = 0
    If rsBar.RecordCount = 0 Then Exit Sub
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
    '独立按钮
    rsBar.Filter = "IsInTool=1 and BarType=1"
    If Not rsBar.EOF Then
        rsBar.Sort = "序号"
        If Not objMenu Is Nothing Then
            With objMenu.CommandBar.Controls
                For i = 1 To rsBar.RecordCount
                    Set objControl = .Add(xtpControlButton, rsBar!功能ID, rsBar!菜单名)
                        objControl.IconId = rsBar!图标ID
                        objControl.Parameter = rsBar!功能名
                        objControl.Style = xtpButtonIconAndCaption
                    If Val(rsBar!IsGroup) = 1 Then
                        objControl.BeginGroup = True
                    End If
                    rsBar.MoveNext
                Next
            End With
        End If
    End If
    
    '下拉按钮，如果只有一个按钮，也当作独立按钮
    rsBar.Filter = "IsInTool=0 and BarType=1"
    If Not rsBar.EOF Then
        rsBar.Sort = "序号"
        If Not objMenu Is Nothing Then
            Set objPopup = objMenu.CommandBar.Controls.Add(xtpControlButtonPopup, conMenu_Tool_PlugIn, "扩展功能", , False)
                objPopup.BeginGroup = True
            With objPopup.CommandBar.Controls
                For i = 1 To rsBar.RecordCount
                    Set objControl = .Add(xtpControlButton, rsBar!功能ID, rsBar!菜单名)
                    objControl.IconId = rsBar!图标ID
                    objControl.Parameter = rsBar!功能名
                    objControl.Style = xtpButtonIconAndCaption
                    If Val(rsBar!IsGroup) = 1 Then
                        objControl.BeginGroup = True
                    End If
                    rsBar.MoveNext
                Next
            End With
        End If
    End If
    
    '工具栏按钮
    Set objBar = cbsMain(2)
    Set objControl = objBar.FindControl(, conMenu_Help_Help)
    If Not objControl Is Nothing Then
        objControl.BeginGroup = True
        lngTmp = objControl.Index - 1
    Else
        lngTmp = -1
    End If
    rsBar.Filter = "IsInTool=1 and BarType=2"
    If Not rsBar.EOF Then
        With objBar.Controls
            For i = 1 To rsBar.RecordCount
                Set objControl = .Add(xtpControlButton, rsBar!功能ID, rsBar!功能名, lngTmp + 1)
                    objControl.IconId = rsBar!图标ID
                    objControl.Parameter = rsBar!功能名
                    objControl.Style = xtpButtonIconAndCaption
                lngTmp = objControl.Index
                If Val(rsBar!IsGroup) = 1 Then objControl.BeginGroup = True
                rsBar.MoveNext
            Next
            objControl.BeginGroup = True
        End With
    End If
    rsBar.Filter = "IsInTool=0 and BarType=2"
    If Not rsBar.EOF Then
        rsBar.Sort = "序号"
        Set objPopup = objBar.Controls.Add(xtpControlPopup, conMenu_Tool_PlugIn, "扩展功能", lngTmp + 1, False)
            objPopup.ID = conMenu_Tool_PlugIn
            objPopup.IconId = conMenu_Tool_PlugIn
            objPopup.BeginGroup = True
            objPopup.Style = xtpButtonIconAndCaption
        lngTmp = objPopup.Index
        With objPopup.CommandBar.Controls
            For i = 1 To rsBar.RecordCount
                Set objControl = .Add(xtpControlButton, rsBar!功能ID, rsBar!菜单名, lngTmp + 1)
                objControl.IconId = rsBar!图标ID
                objControl.Parameter = rsBar!功能名
                If Val(rsBar!IsGroup) = 1 Then objControl.BeginGroup = True
                lngTmp = objPopup.Index
                rsBar.MoveNext
            Next
        End With
    End If
    '自动执行的功能
    rsBar.Filter = "IsAuto=1"
    If Not rsBar.EOF Then mlngPlugInID = rsBar!功能ID
End Sub

Private Sub DefCommandsTechnic(ByVal cbsMain As Object)
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim intTmp As Integer
    Dim strTmp As String
    Dim strName As String
    Dim lngID As Long
    Dim varArr As Variant
    Dim i As Long
    
    '医嘱菜单:放在管理菜单(主窗体可能没有)、文件菜单后面
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    Else
        Call DefCommandPlugInPopup(objMenu.CommandBar.Controls, mrsPlugInBar)
    End If
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "医嘱(&A)", objMenu.Index + 1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_NewItem, "医嘱编辑(&E)", 1)
            objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Edit_NewItem, "新开医嘱(&A)"
            .Add xtpControlButton, conMenu_Edit_Modify, "修改医嘱(&M)"
            .Add xtpControlButton, conMenu_Edit_Delete, "删除医嘱(&D)"
        End With
        
        intTmp = Val(Mid(gstrOutUseApp, 1, 1))
        If intTmp = 1 Then strTmp = strTmp & ",检查申请:" & conMenu_Edit_PacsApply
        intTmp = Val(Mid(gstrOutUseApp, 2, 1))
        If intTmp = 1 And Not gobjLIS Is Nothing Then strTmp = strTmp & ",检验申请:" & conMenu_Edit_LISApply
        intTmp = Val(Mid(gstrOutUseApp, 3, 1))
        If intTmp = 1 Then strTmp = strTmp & ",输血申请:" & conMenu_Edit_BloodApply
        intTmp = Val(Mid(gstrOutUseApp, 4, 1))
        If intTmp = 1 Then strTmp = strTmp & ",手术申请:" & conMenu_Edit_OperationApply
                Get自定义申请单 1, mstr自定义申请单IDs
        If mstr自定义申请单IDs <> "" Then
            For i = 0 To UBound(Split(mstr自定义申请单IDs, "|"))
                strTmp = strTmp & "," & Split(Split(mstr自定义申请单IDs, "|")(i), ",")(1) & ":" & (conMenu_Edit_ApplyCustom * 100# + i) & ":" & Split(Split(mstr自定义申请单IDs, "|")(i), ",")(0)
            Next
        End If
        strTmp = Mid(strTmp, 2)
        
        If strTmp <> "" Then
            If InStr(strTmp, ",") = 0 Then
                strName = Split(strTmp, ":")(0)
                lngID = Val(Split(strTmp, ":")(1))
                Set objControl = .Add(xtpControlButton, lngID, strName)
                    objControl.IconId = conMenu_Manage_Request
                    objControl.ToolTipText = strName
                    objControl.BeginGroup = True
                                If UBound(Split(strTmp, ":")) = 2 Then objControl.Parameter = Val(Split(strTmp, ":")(2))
            Else
                varArr = Split(strTmp, ",")
                For i = 0 To UBound(varArr)
                    strTmp = varArr(i)
                    strName = Split(strTmp, ":")(0)
                    lngID = Val(Split(strTmp, ":")(1))
                    
                    If i = 0 Then
                        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Apply, "下达申请"): objPopup.BeginGroup = True
                        objPopup.IconId = conMenu_Manage_Request
                        Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, lngID, strName)
                    Else
                        Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, lngID, strName)
                    End If
                    If UBound(Split(strTmp, ":")) = 2 Then objControl.Parameter = Val(Split(strTmp, ":")(2))
                Next
            End If
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyModi, "修改申请")
            objControl.IconId = 3002
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyView, "查看申请")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyDel, "取消申请")
        End If
        If HaveRIS Then
            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_NewRis, "检查预约")
                objPopup.IconId = conMenu_Manage_Request
                objPopup.BeginGroup = True
            With objPopup.CommandBar.Controls
                Set objControl = .Add(xtpControlButton, conMenu_Edit_NewRisSch, "预约(&A)")
                    objControl.IconId = conMenu_Edit_NewItem
                Set objControl = .Add(xtpControlButton, conMenu_Edit_NewRisModi, "调整预约(&M)")
                    objControl.IconId = conMenu_Edit_Modify
                Set objControl = .Add(xtpControlButton, conMenu_Edit_NewRisDel, "取消预约(&D)")
                    objControl.IconId = conMenu_Edit_Delete
            End With
        End If
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "医嘱发送(&G)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Blankoff, "医嘱作废(&B)")
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Compend, "报告(&R)")
        objPopup.BeginGroup = True
        objPopup.IconId = conMenu_Manage_Report
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "观片处理(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_MarkKeyMap, "关键图像")
       '2009-01-15
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReportLisView, "浏览检验结果(&R)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReportPacsView, "浏览检查结果(&P)")
            objControl.IconId = conMenu_Manage_ReportLisView

        If CreateObjectPacs(mobjPublicPACS) Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ViewPacs, "浏览检查图像和报告(&Y)")
                objControl.IconId = 237
        End If
        '2017-11-10 刘鹏飞
        If gbln血库系统 Then
            Set objControl = .Add(xtpControlButton, conMenu_Report_BloodInstant, "输血执行单")
            objControl.BeginGroup = True
        End If
    End With
    If Not objMenu Is Nothing Then
        With objMenu.CommandBar.Controls
            If mbln指引单打印 Then
                Set objControl = .Add(xtpControlButton, conMenu_Report_ClinicIndexBill, "打印指引单")
                    objControl.IconId = 103
            End If
            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Report_ClinicBill, "打印单据")
            objPopup.BeginGroup = True
            Call DefCommandPlugInPopup(objMenu.CommandBar.Controls, mrsPlugInBar)
        End With
    End If
    
    '报表菜单:主窗体可能没有,放在查看菜单前面
    '-----------------------------------------------------
    '工作站报表菜单自动显示报表是针对工作站的模块号统一发布
    '而这几张报表是医嘱虚拟模块中的，需要在该模块中单独处理
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ReportPopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ReportPopup, "报表(&R)", objMenu.Index, False)
        objMenu.ID = conMenu_ReportPopup '对xtpControlPopup类型的命令ID需重新赋值
    End If
    
    '查看菜单
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With objMenu.CommandBar.Controls
        Set objControl = .Find(, conMenu_View_StatusBar) '状态栏项后
        Set objControl = .Add(xtpControlButton, conMenu_View_Append, "附加信息(&A)", objControl.Index + 1)
        Set objControl = .Add(xtpControlButton, conMenu_View_Hide, "自动隐藏过滤条件栏(&H)", objControl.Index + 1)
    End With
    
    '工具菜单:主窗体可能没有,放在帮助菜单前面
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_HelpPopup)
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "工具(&T)", objMenu.Index, False)
        objMenu.ID = conMenu_ToolPopup
    End If
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Sign, "电子签名(&S)", -1, False): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Tool_SignNew, "医嘱签名(&I)")
            objControl.IconId = conMenu_Tool_Sign
            Set objControl = .Add(xtpControlButton, conMenu_Tool_SignVerify, "验证签名(&V)")
            Set objControl = .Add(xtpControlButton, conMenu_Tool_SignEarse, "取消签名(&E)"): objControl.BeginGroup = True
        End With
        If HaveRIS And gbln启用影像信息系统预约 Then
            Set objControl = .Add(xtpControlButton, conMenu_Tool_RisPrint, "打印预约单")
                objControl.IconId = 103
        End If
        If gbln科室药房对照按本机参数设置 Then
            Set objControl = .Add(xtpControlButton, conMenu_Tool_Option, "医嘱选项(&O)"): objControl.BeginGroup = True
                objControl.IconId = conMenu_File_Parameter
        End If
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Define, "成套方案定义(&S)"): objControl.BeginGroup = True
    End With

    '工具栏定义:从文件及管理菜单的命令按钮之后开始加入
    '-----------------------------------------------------
    Call AddToolBarInDoctor
    
    '命令的快键绑定
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem '新开医嘱
        .Add FCONTROL, vbKeyM, conMenu_Edit_Modify '修改医嘱
        .Add FSHIFT, vbKeyDelete, conMenu_Edit_Delete '删除医嘱
        .Add FCONTROL, vbKeyG, conMenu_Edit_Send '医嘱发送
        
        .Add FCONTROL, vbKeyR, conMenu_Edit_Compend * 10# + 1 '查阅报告
        .Add FCONTROL, vbKeyW, conMenu_Edit_MarkMap '观片处理
        .Add FCONTROL, vbKeyY, conMenu_Edit_ViewPacs '浏览检查图像和报告
        
        .Add FCONTROL, vbKeyH, conMenu_View_Hide '自动隐藏过滤条件栏
        .Add FCONTROL, vbKeyL, conMenu_Manage_ReportLisView  '浏览检验结果
        .Add FCONTROL, vbKeyP, conMenu_Manage_ReportPacsView '浏览检查结果
        .Add 0, vbKeyF11, conMenu_Tool_Option '医嘱选项
    End With

    '设置不常用命令
    '-----------------------------------------------------
    With cbsMain.Options
    End With
End Sub

Private Sub DefCommandsOutDoctor(ByVal cbsMain As Object)
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl, lngIdx As Long
    
    Dim varArr As Variant
    Dim strTmp As String
    Dim intTmp As Integer
    Dim strName As String
    Dim lngID As Long
    Dim i As Long
    

    '医嘱菜单:放在管理菜单(主窗体可能没有)、文件菜单后面
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    Else
        Call DefCommandPlugInPopup(objMenu.CommandBar.Controls, mrsPlugInBar)
    End If
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "医嘱(&A)", objMenu.Index + 1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_NewItem, "医嘱编辑(&E)", 1)
            objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Edit_NewItem, "新开医嘱(&A)"
            .Add xtpControlButton, conMenu_Edit_Modify, "修改医嘱(&M)"
            .Add xtpControlButton, conMenu_Edit_Delete, "删除医嘱(&D)"
        End With
     
        intTmp = Val(Mid(gstrOutUseApp, 1, 1))
        If intTmp = 1 Then strTmp = strTmp & ",检查申请:" & conMenu_Edit_PacsApply
        intTmp = Val(Mid(gstrOutUseApp, 2, 1))
        If intTmp = 1 And Not gobjLIS Is Nothing Then strTmp = strTmp & ",检验申请:" & conMenu_Edit_LISApply
        intTmp = Val(Mid(gstrOutUseApp, 3, 1))
        If intTmp = 1 Then strTmp = strTmp & ",输血申请:" & conMenu_Edit_BloodApply
        intTmp = Val(Mid(gstrOutUseApp, 4, 1))
        If intTmp = 1 Then strTmp = strTmp & ",手术申请:" & conMenu_Edit_OperationApply
        Get自定义申请单 1, mstr自定义申请单IDs
        If mstr自定义申请单IDs <> "" Then
            For i = 0 To UBound(Split(mstr自定义申请单IDs, "|"))
                strTmp = strTmp & "," & Split(Split(mstr自定义申请单IDs, "|")(i), ",")(1) & ":" & (conMenu_Edit_ApplyCustom * 100# + i) & ":" & Split(Split(mstr自定义申请单IDs, "|")(i), ",")(0)
            Next
        End If
        strTmp = Mid(strTmp, 2)
        
        If strTmp <> "" Then
            If InStr(strTmp, ",") = 0 Then
                strName = Split(strTmp, ":")(0)
                lngID = Val(Split(strTmp, ":")(1))
                Set objControl = .Add(xtpControlButton, lngID, strName)
                    objControl.IconId = conMenu_Manage_Request
                    objControl.ToolTipText = strName
                    objControl.BeginGroup = True
                                        If UBound(Split(strTmp, ":")) = 2 Then objControl.Parameter = Val(Split(strTmp, ":")(2))
            Else
                varArr = Split(strTmp, ",")
                For i = 0 To UBound(varArr)
                    strTmp = varArr(i)
                    strName = Split(strTmp, ":")(0)
                    lngID = Val(Split(strTmp, ":")(1))
                    
                    If i = 0 Then
                        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Apply, "下达申请"): objPopup.BeginGroup = True
                        objPopup.IconId = conMenu_Manage_Request
                        Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, lngID, strName)
                    Else
                        Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, lngID, strName)
                    End If
                    If UBound(Split(strTmp, ":")) = 2 Then objControl.Parameter = Val(Split(strTmp, ":")(2))
                Next
            End If
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyModi, "修改申请")
            objControl.IconId = 3002
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyView, "查看申请")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyDel, "取消申请")
        End If
        If HaveRIS Then
            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_NewRis, "检查预约")
                objPopup.IconId = conMenu_Manage_Request
                objPopup.BeginGroup = True
            With objPopup.CommandBar.Controls
                Set objControl = .Add(xtpControlButton, conMenu_Edit_NewRisSch, "预约(&A)")
                    objControl.IconId = conMenu_Edit_NewItem
                Set objControl = .Add(xtpControlButton, conMenu_Edit_NewRisModi, "调整预约(&M)")
                    objControl.IconId = conMenu_Edit_Modify
                Set objControl = .Add(xtpControlButton, conMenu_Edit_NewRisDel, "取消预约(&D)")
                    objControl.IconId = conMenu_Edit_Delete
            End With
        End If
                
        If gbln血库系统 Then Set objControl = .Add(xtpControlButton, conMenu_Edit_TraReaction, "输血反应"): objControl.IconId = 4113
        
        If mbln危急值 Then
            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_CriticalAdvice, "危急值医嘱")
        End If

        Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "医嘱发送(&G)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Blankoff, "医嘱作废(&B)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Test, "皮试结果(&T)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_AdvicePay, "诊间支付")
            objControl.IconId = conMenu_Edit_Pay
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Compend, "报告(&R)")
        objPopup.BeginGroup = True
        objPopup.IconId = conMenu_Manage_Report
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "观片处理(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_MarkKeyMap, "关键图像")
       '2009-01-15
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReportLisView, "浏览检验结果(&R)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReportPacsView, "浏览检查结果(&P)")
            objControl.IconId = conMenu_Manage_ReportLisView

        If CreateObjectPacs(mobjPublicPACS) Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ViewPacs, "浏览检查图像和报告(&Y)")
                objControl.IconId = 237
        End If
            
        Set objControl = .Add(xtpControlButton, conMenu_Manage_RecipeAuditView, "查看处方审查结果")
        objControl.IconId = 3205
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ViewDrugExplain, "查看药品说明书")
        objControl.IconId = 3205
        If gbln审方系统 Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Refcom, "拒绝审查理由")
                objControl.IconId = 3205
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ViewRefcom, "查阅审核未通过信息")
                objControl.IconId = 3205
        End If
        If mblnPass Then
            Call gobjPass.zlPassCommandBarAdd(mobjPassMap, objMenu.CommandBar.Controls, conMenu_Edit_MediAudit, conMenu_Edit_FeeAudit)
        End If
    End With
    With objMenu.CommandBar.Controls
        If mbln指引单打印 Then
            Set objControl = .Add(xtpControlButton, conMenu_Report_ClinicIndexBill, "打印指引单")
                objControl.IconId = 103
        End If
        '子项放在最前面,反序加入
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Report_ClinicBill, "打印单据")
        objPopup.BeginGroup = True
        Call DefCommandPlugInPopup(objMenu.CommandBar.Controls, mrsPlugInBar)
    End With
    
    
    '报表菜单:主窗体可能没有,放在查看菜单前面
    '-----------------------------------------------------
    '工作站报表菜单自动显示报表是针对工作站的模块号统一发布
    '而这几张报表是医嘱虚拟模块中的，需要在该模块中单独处理
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ReportPopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ReportPopup, "报表(&R)", objMenu.Index, False)
        objMenu.ID = conMenu_ReportPopup '对xtpControlPopup类型的命令ID需重新赋值
    End If

    '查看菜单
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With objMenu.CommandBar.Controls
        Set objControl = .Find(, conMenu_View_StatusBar) '状态栏项后
        Set objControl = .Add(xtpControlButton, conMenu_View_Append, "附加信息(&A)", objControl.Index + 1)
        Set objControl = .Add(xtpControlButton, conMenu_View_Hide, "自动隐藏过滤条件栏(&H)", objControl.Index + 1)
    End With
    
    '工具菜单:主窗体可能没有,放在帮助菜单前面
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_HelpPopup)
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "工具(&T)", objMenu.Index, False)
        objMenu.ID = conMenu_ToolPopup
    End If
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Sign, "电子签名(&S)", -1, False): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Tool_SignNew, "医嘱签名(&I)")
            objControl.IconId = conMenu_Tool_Sign
            Set objControl = .Add(xtpControlButton, conMenu_Tool_SignVerify, "验证签名(&V)")
            Set objControl = .Add(xtpControlButton, conMenu_Tool_SignEarse, "取消签名(&E)"): objControl.BeginGroup = True
        End With
        If HaveRIS And gbln启用影像信息系统预约 Then
            Set objControl = .Add(xtpControlButton, conMenu_Tool_RisPrint, "打印预约单")
                objControl.IconId = 103
        End If
        If gbln科室药房对照按本机参数设置 Then
            Set objControl = .Add(xtpControlButton, conMenu_Tool_Option, "医嘱选项(&O)"): objControl.BeginGroup = True
                objControl.IconId = conMenu_File_Parameter
        End If
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Define, "成套方案定义(&S)"): objControl.BeginGroup = True
    End With

    '工具栏定义:从文件及管理菜单的命令按钮之后开始加入
    '-----------------------------------------------------
    Call AddToolBarInDoctor
    
    '命令的快键绑定
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem '新开医嘱
        .Add FCONTROL, vbKeyM, conMenu_Edit_Modify '修改医嘱
        .Add FSHIFT, vbKeyDelete, conMenu_Edit_Delete '删除医嘱
        .Add FCONTROL, vbKeyG, conMenu_Edit_Send '医嘱发送
        .Add FCONTROL, vbKeyT, conMenu_Edit_Test '皮试结果
        
        .Add FCONTROL, vbKeyR, conMenu_Edit_Compend * 10# + 1 '查阅报告
        .Add FCONTROL, vbKeyW, conMenu_Edit_MarkMap '观片处理
        .Add FCONTROL, vbKeyY, conMenu_Edit_ViewPacs '浏览检查图像和报告
        
        .Add FCONTROL, vbKeyH, conMenu_View_Hide '自动隐藏过滤条件栏
        .Add FCONTROL, vbKeyL, conMenu_Manage_ReportLisView  '浏览检验结果
        .Add FCONTROL, vbKeyP, conMenu_Manage_ReportPacsView '浏览检查结果
        .Add 0, vbKeyF11, conMenu_Tool_Option '医嘱选项
    End With

    '设置不常用命令
    '-----------------------------------------------------
    With cbsMain.Options
    End With
End Sub

Public Sub zlPopupCommandBars(ByVal CommandBar As CommandBar)
    Dim objControl As CommandBarControl
    Dim strTmp As String
    Dim arrTmp As Variant
    Dim i As Integer
    Dim lng医嘱ID As Long
    Dim rsTmp As ADODB.Recordset
        
    If CommandBar.Parent Is Nothing Then Exit Sub
    
    Select Case CommandBar.Parent.ID
    Case conMenu_Edit_CriticalAdvice
        If mbln危急值 And Not mrs危急值 Is Nothing Then
            mrs危急值.Filter = 0
            If Not mrs危急值.EOF Then
                Set rsTmp = GetCriticalAdvice(lng医嘱ID)
                With CommandBar.Controls
                    .DeleteAll
                    mrs危急值.MoveFirst
                    For i = 1 To mrs危急值.RecordCount
                        Set objControl = .Add(xtpControlButton, conMenu_Edit_CriticalAdvice * 100# + i, mrs危急值!危急值描述 & "")
                            objControl.Parameter = mrs危急值!ID & "," & lng医嘱ID
                        rsTmp.Filter = "危急值ID=" & mrs危急值!ID
                        If Not rsTmp.EOF Then
                            objControl.Checked = True
                        End If
                        mrs危急值.MoveNext
                    Next
                    mrs危急值.MoveFirst
                End With
            End If
            mrs危急值.Filter = 0
        End If
    Case conMenu_Edit_Compend '报告
        With CommandBar.Controls
            If .Count = 0 Then
                .Add xtpControlButton, conMenu_Edit_Compend * 10# + 1, "查阅报告(病历格式)"
                .Add xtpControlButton, conMenu_Edit_Compend * 10# + 6, "查阅报告(报表格式)"
                If gobjExchange Is Nothing Then
                    If mint场合 = 1 Then    '护士站
                        .Add xtpControlButton, conMenu_Edit_Compend * 10# + 3, "预览报告(&V)"
                    Else
                        .Add(xtpControlButton, conMenu_Edit_Compend * 10# + 2, "打印报告(&P)").BeginGroup = True
                        .Add xtpControlButton, conMenu_Edit_Compend * 10# + 3, "预览报告(&V)"

                        .Add(xtpControlButton, conMenu_Edit_Compend * 10# + 4, "我已查阅(&R)").BeginGroup = True
                        .Add xtpControlButton, conMenu_Edit_Compend * 10# + 5, "自动标记(&A)"
                    End If
                End If
            End If
        End With
    Case conMenu_Edit_MediAudit, conMenu_Edit_MediAudit * 10# To conMenu_Edit_MediAudit * 10# + 99
        'PASS药嘱审查
        If mblnPass Then
            Call gobjPass.zlPASSPopupCommandBars(mobjPassMap, CommandBar, conMenu_Edit_MediAudit)
        End If
    End Select
End Sub

Private Sub AddToolBarInDoctor()
'功能：设置工具栏按钮，对应于医嘱菜单下面的工具栏的按钮，先将其删掉再添加
    Dim objControl As CommandBarControl
    Dim objMenuBar As CommandBarPopup
    Dim objPopup As CommandBarPopup
    Dim varArr As Variant
    Dim strTmp As String
    Dim lngTmp As Long
    Dim objCbs As Object
    Dim lngIdx As Long
    Dim i As Long
    
    Dim intTmp As Integer
    Dim strName As String
    Dim lngID As Long
    
    Dim blnTwo As Boolean, strInsidePrivs As String
    
    On Error GoTo errH
    
    If mcbsMain Is Nothing Then Exit Sub

    strInsidePrivs = GetInsidePrivs(p门诊医生站)
    blnTwo = Val(zlDatabase.GetPara("发送单据类型", glngSys, p门诊医嘱下达)) <> 2
    
    strTmp = "," & conMenu_Edit_NewItem & "," & conMenu_Edit_Apply & "," & conMenu_Edit_ApplyModi & "," & conMenu_Edit_ApplyView & "," & conMenu_Edit_ApplyDel & "," & _
        conMenu_Edit_Blankoff & "," & conMenu_Edit_TraReaction & "," & conMenu_Edit_SendBilling & "," & conMenu_Edit_Send & "," & IIF(blnTwo, conMenu_Edit_Send * 100# + 1 & ",", "") & conMenu_Edit_Untread & "," & _
        conMenu_Edit_Compend & "," & (conMenu_Edit_Compend * 10# + 2) & "," & (conMenu_Edit_Compend * 10# + 3) & "," & conMenu_Edit_MarkMap & "," & conMenu_Edit_MarkKeyMap & "," & conMenu_Edit_MarkKeyMap & "," & conMenu_Manage_ReportLisView & "," & _
        conMenu_Edit_MediAudit & "," & conMenu_Tool_SignNew & "," & conMenu_Edit_Audit & "," & conMenu_Edit_Price & "," & conMenu_Report_ClinicBill & "," & conMenu_Edit_PacsApply & "," & conMenu_Edit_BloodApply & ","
    strTmp = strTmp & "," & conMenu_Edit_PacsApply & "," & (conMenu_Edit_PacsApply * 10# + 1) & "," & conMenu_Edit_LISApply & "," & (conMenu_Edit_LISApply * 10# + 1) & "," & conMenu_Edit_BloodApply & "," & (conMenu_Edit_BloodApply * 10# + 1)
    strTmp = strTmp & "," & conMenu_Edit_OperationApply & "," & (conMenu_Edit_OperationApply * 10# + 1) & "," & conMenu_Edit_ConsultationApply & "," & (conMenu_Edit_ConsultationApply * 10 + 1) & "," & conMenu_Edit_AdvicePay & ","
    
    '工具栏添加
    Set objCbs = mcbsMain
    '找到要添加的位置
    lngIdx = 0
    For Each objControl In objCbs(2).Controls '先求出前面的最后一个Control
        If Val(Left(objControl.ID, 1)) <> conMenu_FilePopup And Val(Left(objControl.ID, 1)) <> conMenu_ManagePopup Then
            Set objControl = objCbs(2).Controls(objControl.Index - 1)
            lngIdx = objControl.Index
            Exit For
        End If
    Next
    
    '删除工具栏按钮
    For i = objCbs(2).Controls.Count To 1 Step -1
        If InStr(strTmp, "," & objCbs(2).Controls(i).ID & ",") > 0 Then
            objCbs(2).Controls(i).Delete
        End If
    Next i

    With objCbs(2).Controls
        If mvarCond.过滤模式 <> 3 Then
            Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Edit_NewItem, "新开", lngIdx + 1): objPopup.BeginGroup = True
            With objPopup.CommandBar.Controls
                Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem * 10# + 1, "新开")
                    objControl.IconId = conMenu_Edit_NewItem
                .Add xtpControlButton, conMenu_Edit_Modify, "修改"
                .Add xtpControlButton, conMenu_Edit_Delete, "删除"
            End With
            objPopup.Style = xtpButtonIconAndCaption
            lngIdx = objPopup.Index
        End If
        
        If mint场合 = 0 Then '只有门诊医生工作站调用时才有这几个按钮
            strTmp = ""
            intTmp = Val(Mid(gstrOutUseApp, 1, 1))
            If intTmp = 1 Then strTmp = strTmp & ",检查申请:" & conMenu_Edit_PacsApply
            intTmp = Val(Mid(gstrOutUseApp, 2, 1))
            If intTmp = 1 And Not gobjLIS Is Nothing Then strTmp = strTmp & ",检验申请:" & conMenu_Edit_LISApply
            intTmp = Val(Mid(gstrOutUseApp, 3, 1))
            If intTmp = 1 Then strTmp = strTmp & ",输血申请:" & conMenu_Edit_BloodApply
            intTmp = Val(Mid(gstrOutUseApp, 4, 1))
            If intTmp = 1 Then strTmp = strTmp & ",手术申请:" & conMenu_Edit_OperationApply
            Get自定义申请单 1, mstr自定义申请单IDs
            If mstr自定义申请单IDs <> "" Then
                For i = 0 To UBound(Split(mstr自定义申请单IDs, "|"))
                    strTmp = strTmp & "," & Split(Split(mstr自定义申请单IDs, "|")(i), ",")(1) & ":" & (conMenu_Edit_ApplyCustom * 100# + i) & ":" & Split(Split(mstr自定义申请单IDs, "|")(i), ",")(0)
                Next
            End If
            strTmp = Mid(strTmp, 2)
            
            If strTmp <> "" Then
                If InStr(strTmp, ",") = 0 Then
                    strName = Split(strTmp, ":")(0)
                    lngID = Val(Split(strTmp, ":")(1))
                    Set objControl = .Add(xtpControlButton, lngID, strName, lngIdx + 1)
                        objControl.IconId = conMenu_Manage_Request
                        objControl.ToolTipText = strName
                        objControl.Style = xtpButtonIconAndCaption
                        objControl.BeginGroup = True
                                                If UBound(Split(strTmp, ":")) = 2 Then objControl.Parameter = Val(Split(strTmp, ":")(2))
                    lngIdx = objControl.Index
                Else
                    varArr = Split(strTmp, ",")
                    For i = 0 To UBound(varArr)
                        strTmp = varArr(i)
                        strName = Split(strTmp, ":")(0)
                        lngID = Val(Split(strTmp, ":")(1))
                        
                        If i = 0 Then
                            Set objPopup = .Add(xtpControlSplitButtonPopup, lngID, strName, lngIdx + 1)
                                objPopup.IconId = conMenu_Manage_Request
                                objPopup.BeginGroup = True
                                objPopup.ToolTipText = strName
                                objPopup.Style = xtpButtonIconAndCaption
                                With objPopup.CommandBar.Controls
                                    Set objControl = .Add(xtpControlButton, lngID * 10# + 1, strName)
                                End With
                        Else
                            Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, lngID, strName)
                        End If
                        If UBound(Split(strTmp, ":")) = 2 Then objControl.Parameter = Val(Split(strTmp, ":")(2))
                    Next
                    lngIdx = objPopup.Index
                End If
            End If
        End If
        
        If mvarCond.过滤模式 = 3 And mint场合 = 0 Then '只有住院医生工作站调用时才有这几个按钮
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyModi, "修改", lngIdx + 1)
                objControl.IconId = 3002
                objControl.ToolTipText = "修改申请"
                objControl.Style = xtpButtonIconAndCaption
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyView, "查看", objControl.Index + 1)
                objControl.IconId = 102
                objControl.ToolTipText = "查看申请"
                objControl.Style = xtpButtonIconAndCaption
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyDel, "取消", objControl.Index + 1)
                objControl.IconId = 3004
                objControl.ToolTipText = "取消申请"
                objControl.Style = xtpButtonIconAndCaption
            lngIdx = objControl.Index
        End If
        
        If blnTwo Then
            Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Edit_Send, "发送", lngIdx + 1)
            objPopup.BeginGroup = True
            objPopup.ToolTipText = "医嘱发送处理"
            objPopup.Style = xtpButtonIconAndCaption
            With objPopup.CommandBar.Controls
                .Add xtpControlButton, conMenu_Edit_Send * 100# + 1, "自动完成发送"
            End With
            lngIdx = objPopup.Index
        Else
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "发送", lngIdx + 1)
            objControl.Style = xtpButtonIconAndCaption
            objControl.BeginGroup = True
            objControl.ToolTipText = "医嘱发送处理"
            lngIdx = objControl.Index
        End If
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Blankoff, "作废", lngIdx + 1)
            objControl.BeginGroup = True
            objControl.Style = xtpButtonIconAndCaption
        lngIdx = objControl.Index
        Set objControl = .Add(xtpControlButton, conMenu_Edit_AdvicePay, "诊间支付", lngIdx + 1)
            objControl.IconId = conMenu_Edit_Pay
            objControl.BeginGroup = True
            objControl.Style = xtpButtonIconAndCaption
        lngIdx = objControl.Index
        If mvarCond.过滤模式 = 3 Then
            Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Edit_Compend, "查阅", lngIdx + 1): objPopup.BeginGroup = True
                objPopup.IconId = conMenu_Manage_Report
                objPopup.ToolTipText = "查阅报告"
                
            With objPopup.CommandBar.Controls
                Set objControl = .Add(xtpControlButton, conMenu_Edit_Compend * 10# + 1, "病历格式(&B)"): objControl.IconId = 102
                Set objControl = .Add(xtpControlButton, conMenu_Edit_Compend * 10# + 6, "报表格式(&P)"): objControl.IconId = 102
                If gobjExchange Is Nothing And mint场合 <> 1 Then
                    Set objControl = .Add(xtpControlButton, conMenu_Edit_Compend * 10# + 4, "我已查阅(&R)")
                        objControl.BeginGroup = True
                    .Add xtpControlButton, conMenu_Edit_Compend * 10# + 5, "自动标记(&A)"
                End If
            End With
            objPopup.Style = xtpButtonIconAndCaption
            lngIdx = objPopup.Index
            If gobjExchange Is Nothing Then
                If mint场合 <> 1 Then
                    Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Edit_Compend * 10# + 2, "打印报告", lngIdx + 1)
                        objPopup.IconId = 103
                        objPopup.Style = xtpButtonIconAndCaption
                        With objPopup.CommandBar.Controls
                            Set objControl = .Add(xtpControlButton, conMenu_Edit_Compend * 10# + 3, "预览报告"): objControl.IconId = 102
                            objControl.Style = xtpButtonIconAndCaption
                        End With
                    lngIdx = objPopup.Index
                Else
                    Set objControl = .Add(xtpControlButton, conMenu_Edit_Compend * 10# + 3, "预览报告(&V)", lngIdx + 1)
                    objControl.IconId = 102
                    objControl.Style = xtpButtonIconAndCaption
                    lngIdx = objControl.Index
                End If
            End If
    
            Set objControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "观片", lngIdx + 1)
                objControl.BeginGroup = True
                objControl.IconId = conMenu_Edit_MarkMap
                objControl.ToolTipText = "观片处理"
                objControl.Style = xtpButtonIconAndCaption
                lngIdx = objControl.Index
            Set objControl = .Add(xtpControlButton, conMenu_Edit_MarkKeyMap, "关键图像", lngIdx + 1)
                objControl.BeginGroup = True
                objControl.IconId = conMenu_Edit_MarkMap
                objControl.ToolTipText = "关键图像"
                objControl.Style = xtpButtonIconAndCaption
            Set objControl = .Add(xtpControlButton, conMenu_Manage_ReportLisView, "结果", objControl.Index + 1): objControl.IconId = conMenu_Manage_ReportLisView
                objControl.ToolTipText = "浏览检验结果"
                objControl.Style = xtpButtonIconAndCaption
            lngIdx = objControl.Index
        Else
            If mint场合 = 0 Then
                If mblnPass Then
                    lngIdx = lngIdx + 1
                    Call gobjPass.zlPassCommandBarAdd(mobjPassMap, objCbs(2).Controls, conMenu_Edit_MediAudit, conMenu_Edit_FeeAudit, lngIdx)
                End If
            End If
            Set objControl = .Add(xtpControlButton, conMenu_Tool_SignNew, "签名", objControl.Index + 1): objControl.BeginGroup = True
                objControl.IconId = conMenu_Tool_Sign
                objControl.Style = xtpButtonIconAndCaption
            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Report_ClinicBill, "打印单据", objControl.Index + 1)
                objPopup.BeginGroup = True
                objPopup.IconId = conMenu_File_Print
                objPopup.Visible = False
        End If
    End With
    mcbsMain.RecalcLayout
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    Dim strErr As String

    Select Case Control.ID
    Case conMenu_File_PrintSet '打印设置
        SwitchPrintSet glngSys & "\" & p门诊医嘱下达
        Call zlPrintSet
        SwitchPrintSet glngSys & "\" & p门诊医嘱下达, True
    Case conMenu_File_Preview '预览医嘱清单
        Call OutputList(2)
    Case conMenu_File_Print '打印医嘱清单
        Call OutputList(1)
    Case conMenu_File_Excel '输出医嘱清单
        Call OutputList(3)
    Case conMenu_Help_Help
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_View_Append '附加数据
        mblnAppend = Not mblnAppend
        tbcAppend.Visible = Not tbcAppend.Visible
        fraAdviceUD.Visible = Not fraAdviceUD.Visible
        Call Form_Resize
        Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
        If vsAppend.Visible And vsAppend.Enabled Then
            vsAppend.SetFocus
        Else
            If vsAdvice.Visible And vsAdvice.Enabled Then vsAdvice.SetFocus
        End If
        Call cbsSub_Resize
    Case conMenu_View_Hide '自动隐藏过滤工具栏
        mblnHideFilter = Not mblnHideFilter
        cbsSub(2).Visible = Not mblnHideFilter
        fraHide.Visible = mblnHideFilter
        cbsSub.RecalcLayout
    Case conMenu_Edit_NewItem, conMenu_Edit_NewItem * 10# + 1 '新开医嘱
        If Control.Parameter <> "" Then
            mlng危急值ID = Val(Control.Parameter)
            Call GetCriticalData
        Else
            mlng危急值ID = 0
        End If
        Call FuncAdviceAdd
    Case conMenu_Edit_Modify '修改医嘱
        Call FuncAdviceModi
    Case conMenu_Edit_Delete, conMenu_Edit_ApplyDel '删除医嘱'取消检验申请
        Call FuncAdviceDel
    Case conMenu_Edit_LISApply, conMenu_Edit_LISApply * 10 + 1   '检验申请
        Call FuncLISApply(0)
    Case conMenu_Edit_PacsApply, conMenu_Edit_PacsApply * 10# + 1 '检查申请
        Call FuncPacsApply(0, 0)
    Case conMenu_Edit_BloodApply, conMenu_Edit_BloodApply * 10 + 1  '输血申请
        Call FuncBloodApply(0)
    Case conMenu_Edit_OperationApply, conMenu_Edit_OperationApply * 10 + 1 '手术申请
        Call FuncOperationApply(0)
    Case conMenu_Edit_ApplyCustom * 100# To conMenu_Edit_ApplyCustom * 101#
        FuncApplyCustom 0, Control.Parameter
    Case conMenu_Edit_ApplyView '查看申请
        Call FuncApplyView
    Case conMenu_Edit_ApplyModi '修改申请
        Call FuncApplyModi
    Case conMenu_Edit_NewRisSch 'RIS预约
        Call FuncAdviceRISSch
    Case conMenu_Edit_NewRisDel '取消预约
        Call FuncAdviceRISDel
    Case conMenu_Edit_NewRisModi
        Call FuncAdviceRISModi
    Case conMenu_Tool_RisPrint
        Call FuncAdviceRISPrintSch
    Case conMenu_Edit_TraReaction '输血反应
        Call FuncTraReaction(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)), p门诊医嘱下达, mblnMoved)
    Case conMenu_Edit_CriticalAdvice * 100# + 1 To conMenu_Edit_CriticalAdvice * 100# + 99
        Call FuncCriticalAdvice(Control.Parameter, Control.Checked)
    Case conMenu_Edit_MediAudit, conMenu_Edit_MediAudit * 10# To conMenu_Edit_MediAudit * 10# + 99  '合理用药审查
        If mblnPass Then
            Call zlPASSPati
            Call gobjPass.zlPassCommandBarExe(mobjPassMap, Control.ID - conMenu_Edit_MediAudit * 10#)
        End If
    Case conMenu_Edit_Send '发送医嘱
        Call FuncAdviceSend(False)
    Case conMenu_Edit_Send * 100# + 1 '自动发送医嘱
        Call FuncAdviceSend(True)
    Case conMenu_Edit_Blankoff '医嘱作废
        Call FuncAdviceRevoke
    Case conMenu_Edit_ViewDrugExplain '查看药品说明书
        Call FuncViewDrugExplain(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_收费细目ID)), mfrmParent)
    Case conMenu_Edit_Refcom '拒绝审查理由
        Call FuncDrugRefcom '药品审查拒绝理由
    Case conMenu_Edit_ViewRefcom '查阅审核未通过信息
        If Not gobjPass Is Nothing And mlng病人ID <> 0 And mlng挂号ID <> 0 Then Call gobjPass.ZLPharmReviewResultShow(Me, mlng病人ID, mlng挂号ID)
    Case conMenu_Edit_Compend, conMenu_Edit_Compend * 10# + 1 To conMenu_Edit_Compend * 10# + 3, conMenu_Edit_Compend * 10# + 6  '查阅、打印报告
        Call FuncEPRReport(Control.ID)
    Case conMenu_Edit_AdvicePay
        Call FuncClinicPay(mfrmParent, mlng病人ID, mstr挂号单)
    Case conMenu_Edit_Compend * 10# + 4 '我是否已经查阅该报告
        Call FuncExecReportRead(Not Control.Checked)
    Case conMenu_Edit_Compend * 10# + 5 '自动标记查阅状态
        mblnAutoRead = Not mblnAutoRead
        Call zlDatabase.SetPara("自动标记报告查阅状态", IIF(mblnAutoRead, 1, 0), glngSys, p门诊医嘱下达)
    Case conMenu_Edit_MarkMap '观片处理
        RaiseEvent ViewPACSImage(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)))
    Case conMenu_Edit_MarkKeyMap '关键图像
        If CreateObjectPacs(mobjPublicPACS) Then
            Call mobjPublicPACS.ShowStaticImage(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)), mfrmParent)
        End If
    Case conMenu_Edit_ViewPacs '浏览检查图像和报告
        If CreateObjectPacs(mobjPublicPACS) Then
            Call mobjPublicPACS.ShowPatientImage(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)), mfrmParent)
        End If
    Case conMenu_Edit_Test '皮试结果
        Call FuncAdviceTest
    Case conMenu_Tool_SignNew '医嘱签名
        Call FuncAdviceSign
    Case conMenu_Tool_SignEarse '取消签名
        Call FuncAdviceSignErase
    Case conMenu_Tool_SignVerify '验证签名
        Call FuncAdviceSignVerify
    Case conMenu_Report_ClinicBill * 100# + 1 To conMenu_Report_ClinicBill * 100# + 99 '打印诊疗单据
        Call FuncBillPrint(Control)
    Case conMenu_Tool_Reference_2 '诊疗措拖参考
        Call zlItemRef
    Case conMenu_Tool_Option '医嘱选项
        frmOutAdviceSetup.Show 1, mfrmParent
    Case conMenu_Tool_Define '成套方案定义
        Call FuncToolScheme
    Case conMenu_Manage_ReportLisView
        Call FuncViewLisRpt
    Case conMenu_Manage_ReportPacsView  '检查报告浏览
        Call FuncViewPacsRpt
    Case conMenu_Report_ClinicIndexBill
        Call FuncAdviceIndexBill
    Case conMenu_Manage_RecipeAuditView '查看处方审查结果
        If InitObjRecipeAudit(p门诊医嘱下达) Then
            Call gobjRecipeAudit.ShowResult(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_相关ID)), mfrmParent)
        End If
    Case conMenu_Report_BloodInstant
        Call PrintBloodReport(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)), mfrmParent)
    Case conMenu_Tool_PlugIn_Item + 1 To conMenu_Tool_PlugIn_Item + 99 '外挂功能执行
        If CreatePlugInOK(p门诊医嘱下达, mint场合) Then
            On Error Resume Next
            If PlugExeNew(Control.Parameter) = False Then
                Call gobjPlugIn.ExecuteFunc(glngSys, p门诊医嘱下达, Control.Parameter, _
                    mlng病人ID, mstr挂号单, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)), mlng前提ID, mint场合)
                Call zlPlugInErrH(err, "ExecuteFunc")
                err.Clear: On Error GoTo 0
            End If
        End If
    End Select
End Sub

Private Function PlugExeNew(ByVal strName As String) As Boolean
'功能：向下兼容外挂部件的ExecuteFunc过程
    Dim lngID As Long
    Dim strXML As String
On Error GoTo errH
    If CreatePlugInOK(p住院医嘱下达, mint场合) Then
        With vsAdvice
            lngID = .RowData(.Row)
            strXML = "<ROOT><诊疗项目名称>" & .TextMatrix(.Row, col_诊疗项目名称) & "</诊疗项目名称></ROOT>"
            Call gobjPlugIn.ExecuteFunc(glngSys, p门诊医嘱下达, strName, mlng病人ID, mstr挂号单, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)), mlng前提ID, mint场合, strXML)
        End With
    End If
   PlugExeNew = True
   Exit Function
errH:
    If err.Number = 450 Then
        PlugExeNew = False
        err.Clear
    Else
        PlugExeNew = True
        Call zlPlugInErrH(err, "ExecuteFunc")
        err.Clear: On Error GoTo 0
    End If
End Function


Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    Dim blnAdvice As Boolean, blnEnabled As Boolean
    Dim i As Integer

    tbcMain.Enabled = mlng病人ID <> 0
    For i = 0 To tbcMain.ItemCount - 1
        tbcMain.Item(i).Enabled = mlng病人ID <> 0
    Next
    
    If vsAdvice.Redraw = flexRDNone Then Exit Sub
    'Pass
    '如果此处不控制，当 control.Id 满足于[conMenu_Edit_MediAudit * 10#, conMenu_Edit_MediAudit * 10# + 99]这个区间 时,下面医嘱操作部分和按钮可见状态中会改变的Pass
    'Enabled属性值。这样在独立部件中设置的enabled的值将会被覆盖。
    If Between(Control.ID, conMenu_Edit_MediAudit * 10#, conMenu_Edit_MediAudit * 10# + 99) Then
        Control.Visible = IIF(Control.Category <> "", InStr(Control.Category, ";可见;") > 0, True)
        Control.Enabled = IIF(Control.Category <> "", InStr(Control.Category, ";可用;") > 0, True)
        Exit Sub
    End If
    
    '根据权限设置按钮可见状态
    Call SetControlVisible(Control)
    If Not Control.Visible Then Exit Sub
        
    '医嘱操作部份
    '------------------------------------------------------------------------------
    '总的判断:无病人或已诊病人不允许任何操作
    If Between(Control.ID, conMenu_Edit_NewItem, conMenu_Edit_NewItem + 998) _
        Or Between(Control.ID, conMenu_Edit_NewItem * 10#, (conMenu_Edit_NewItem + 998) * 10# + 9) Then '包含二级子菜单
        Control.Enabled = mlng病人ID <> 0 _
            And (Control.ID <> conMenu_Edit_Blankoff And mblnEditable Or Control.ID = conMenu_Edit_Blankoff And mblnCanRevoke _
            Or Control.ID = conMenu_Edit_MarkMap Or Control.ID = conMenu_Edit_ViewPacs Or Control.ID = conMenu_Edit_MarkKeyMap Or Control.ID = conMenu_Edit_Compend _
            Or Between(Control.ID, conMenu_Edit_Compend * 10# + 1, conMenu_Edit_Compend * 10# + 5))
        If Not Control.Enabled Then Exit Sub
    End If

    blnAdvice = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) <> 0
    Select Case Control.ID
    Case conMenu_Edit_NewItem, conMenu_Edit_Apply, conMenu_Edit_LISApply, conMenu_Edit_PacsApply, conMenu_Edit_BloodApply, conMenu_Edit_OperationApply, conMenu_Edit_ApplyCustom * 100# To conMenu_Edit_ApplyCustom * 101#  '新开医嘱
        Control.Enabled = (mvRegDate <> CDate("3000-01-01"))
        
        blnEnabled = Control.Enabled
        If blnEnabled And mint场合 = 0 Then
            If mstr接诊医生 <> UserInfo.姓名 Then
                blnEnabled = False
                If InStr(GetInsidePrivs(p门诊医生站), ";操作其他医生的病人;") > 0 Then
                    blnEnabled = True
                End If
            End If
        End If
        Control.Enabled = blnEnabled
    Case conMenu_Edit_Modify, conMenu_Edit_Delete '修改医嘱,删除医嘱
        blnEnabled = blnAdvice _
            And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱状态)) = 1 _
            And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_签名否)) = 0
        
        If blnEnabled And mint场合 = 2 Then
            blnEnabled = InStr("," & mstr前提IDs & ",", "," & Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_前提ID)) & ",") > 0
        ElseIf blnEnabled And mint场合 <> 2 Then
            blnEnabled = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_前提ID)) = 0
        End If
      
        If blnEnabled And mint场合 = 0 Then
            If mstr接诊医生 <> UserInfo.姓名 Then
                blnEnabled = False
                If InStr(GetInsidePrivs(p门诊医生站), ";操作其他医生的病人;") > 0 Then
                    blnEnabled = True
                End If
            End If
        End If
        Control.Enabled = blnEnabled
    Case conMenu_Manage_RecipeAuditView
        blnEnabled = blnAdvice And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_处方审查状态)) <> 0
        Control.Enabled = blnEnabled
    '申请单取消
    Case conMenu_Edit_ApplyDel
        blnEnabled = blnAdvice
        If blnEnabled Then
            With vsAdvice
                If Not (.TextMatrix(.Row, COL_医嘱状态) = "1" And _
                    (.TextMatrix(.Row, COL_诊疗类别) = "D" Or .TextMatrix(.Row, COL_诊疗类别) = "F" Or _
                        Val(.TextMatrix(.Row, COL_操作类型)) = 6 And .TextMatrix(.Row, COL_诊疗类别) = "E" Or _
                        .TextMatrix(.Row, COL_诊疗类别) = "K")) Then
                    blnEnabled = Val(.TextMatrix(.Row, COL_申请序号)) <> 0
                End If
                '用血医嘱待审核不允许取消（新血库流程数据）
                If blnEnabled = True And .TextMatrix(.Row, COL_诊疗类别) = "K" And .TextMatrix(.Row, COL_医嘱状态) = "1" Then
                    If Val(.TextMatrix(.Row, COL_检查方法)) = 1 And Val(.TextMatrix(.Row, COL_审核状态)) = 1 Then blnEnabled = False
                End If
            End With
        End If
        If blnEnabled And mint场合 = 0 Then
            If mstr接诊医生 <> UserInfo.姓名 Then
                blnEnabled = False
                If InStr(GetInsidePrivs(p门诊医生站), ";操作其他医生的病人;") > 0 Then
                    blnEnabled = True
                End If
            End If
        End If
        Control.Enabled = blnEnabled
    '检查申请修改
    Case conMenu_Edit_ApplyModi
        blnEnabled = blnAdvice
        If blnEnabled Then
            With vsAdvice
                If Val(.TextMatrix(.Row, COL_申请序号)) <> 0 Then
                    If Val(.TextMatrix(.Row, COL_操作类型)) = 6 And .TextMatrix(.Row, COL_诊疗类别) = "E" Then
                        If Not (.TextMatrix(.Row, COL_医嘱状态) = "1" And Val(.TextMatrix(.Row, COL_操作类型)) = 6 And .TextMatrix(.Row, COL_诊疗类别) = "E") Then blnEnabled = False
                    ElseIf .TextMatrix(.Row, COL_诊疗类别) = "D" And .TextMatrix(.Row, COL_操作类型) <> "病理" Then
                        If Not (.TextMatrix(.Row, COL_医嘱状态) = "1" And .TextMatrix(.Row, COL_诊疗类别) = "D") Then blnEnabled = False
                    ElseIf .TextMatrix(.Row, COL_诊疗类别) = "K" Then
                        If Not (.TextMatrix(.Row, COL_医嘱状态) = "1" And .TextMatrix(.Row, COL_诊疗类别) = "K") Then blnEnabled = False
                    ElseIf .TextMatrix(.Row, COL_诊疗类别) = "F" Then
                        If Not (.TextMatrix(.Row, COL_医嘱状态) = "1" And .TextMatrix(.Row, COL_诊疗类别) = "F") Then blnEnabled = False
                    Else
                        blnEnabled = Val(.TextMatrix(.Row, COL_申请序号)) <> 0
                    End If
                Else
                    blnEnabled = False
                End If
            End With
        End If
        If blnEnabled And mint场合 = 0 Then
            If mstr接诊医生 <> UserInfo.姓名 Then
                blnEnabled = False
                If InStr(GetInsidePrivs(p门诊医生站), ";操作其他医生的病人;") > 0 Then
                    blnEnabled = True
                End If
            End If
        End If
        Control.Enabled = blnEnabled
    Case conMenu_Edit_NewRis
        blnEnabled = False
        With vsAdvice
            If InStr(",D,F,", .TextMatrix(.Row, COL_诊疗类别)) > 0 Or InStr(",0,5,", Val(.TextMatrix(.Row, COL_操作类型))) > 0 And .TextMatrix(.Row, COL_诊疗类别) = "E" Then
                blnEnabled = True
            End If
        End With
        Control.Enabled = blnEnabled
    Case conMenu_Edit_NewRisSch
        blnEnabled = False
        If gbln启用影像信息系统预约 Then
            With vsAdvice
                If (InStr(",D,F,", .TextMatrix(.Row, COL_诊疗类别)) > 0 Or InStr(",0,5,", Val(.TextMatrix(.Row, COL_操作类型))) > 0 And .TextMatrix(.Row, COL_诊疗类别) = "E") And Val(.TextMatrix(.Row, COL_RIS预约ID)) = 0 Then
                    blnEnabled = True
                End If
            End With
        End If
        Control.Enabled = blnEnabled
    Case conMenu_Edit_NewRisDel, conMenu_Tool_RisPrint
        Control.Enabled = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_RIS预约ID)) <> 0
    Case conMenu_Edit_NewRisModi
        Control.Enabled = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_RIS预约ID)) <> 0 And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱状态)) = 8
    '查看申请
    Case conMenu_Edit_ApplyView
        blnEnabled = blnAdvice
        If blnEnabled Then
            With vsAdvice
                If Not (InStr(",K,F,D,", .TextMatrix(.Row, COL_诊疗类别)) > 0) Then blnEnabled = Val(.TextMatrix(.Row, COL_申请序号)) <> 0
            End With
        End If
        Control.Enabled = blnEnabled
    Case conMenu_Edit_TraReaction
        With vsAdvice
            blnEnabled = (.TextMatrix(.Row, COL_诊疗类别) = "K") And Val(.TextMatrix(.Row, COL_医嘱状态)) = 8 And gbln血库系统
        End With
        Control.Enabled = blnEnabled
    Case conMenu_Edit_Blankoff '医嘱作废
        blnEnabled = blnAdvice _
            And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱状态)) = 8 _
            And (Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_签名否)) = 0 Or Not gobjESign Is Nothing)
            
        '有权限则都能作废；
        '无权限，医技站的关联的医嘱只是在本科室才能作废，且医技站和医生站互不操作；
        If blnEnabled And mint场合 = 2 Then
            blnEnabled = InStr("," & mstr前提IDs & ",", "," & Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_前提ID)) & ",") > 0 _
            And vsAdvice.TextMatrix(vsAdvice.Row, COL_开嘱医生) = UserInfo.姓名 Or InStr(GetInsidePrivs(p门诊医嘱下达), "作废他人医嘱") > 0
        ElseIf blnEnabled And mint场合 <> 2 Then
            blnEnabled = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_前提ID)) = 0 _
            And vsAdvice.TextMatrix(vsAdvice.Row, COL_开嘱医生) = UserInfo.姓名 Or InStr(GetInsidePrivs(p门诊医嘱下达), "作废他人医嘱") > 0
        End If
        If blnEnabled And mint场合 = 0 Then
            If mstr接诊医生 <> UserInfo.姓名 Then
                blnEnabled = False
                If InStr(GetInsidePrivs(p门诊医生站), ";操作其他医生的病人;") > 0 Then
                    blnEnabled = True
                End If
            End If
        End If
        Control.Enabled = blnEnabled
    Case conMenu_Edit_Send, conMenu_Edit_Send * 100# + 1 '医嘱发送
        blnEnabled = True
        If mint场合 = 0 Then
            If mstr接诊医生 <> UserInfo.姓名 Then
                blnEnabled = False
                If InStr(GetInsidePrivs(p门诊医生站), ";操作其他医生的病人;") > 0 Then
                    blnEnabled = True
                End If
            End If
        End If
        Control.Enabled = blnEnabled
    Case conMenu_Edit_CriticalAdvice
        blnEnabled = False
        If Not mrs危急值 Is Nothing Then
            If Not mrs危急值.EOF Then
                blnEnabled = True
            End If
        End If
        If blnEnabled Then
            blnEnabled = (Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱状态)) <> 4 And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) <> 0)
        End If
        Control.Enabled = blnEnabled
    Case conMenu_Edit_ViewDrugExplain '查看药品说明书
        Control.Enabled = blnAdvice And InStr(",5,6,7,", vsAdvice.TextMatrix(vsAdvice.Row, COL_诊疗类别)) > 0
    Case conMenu_Edit_Test '皮试结果
        With vsAdvice
            Control.Enabled = blnAdvice _
                And Val(.TextMatrix(.Row, COL_前提ID)) = 0 _
                And Val(.TextMatrix(.Row, COL_医嘱状态)) <> 4 _
                And .TextMatrix(.Row, COL_诊疗类别) = "E" And .TextMatrix(.Row, COL_操作类型) = "1"
        End With
    Case conMenu_Report_ClinicIndexBill
        blnEnabled = blnAdvice And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱状态)) = 8
        Control.Enabled = blnEnabled
    Case conMenu_Edit_MediAudit '药嘱审查(中药不显示)
        If mblnPass Then
            Call gobjPass.zlPassCommandBarUpdate(mobjPassMap, Control, blnAdvice)
        End If
    Case conMenu_Edit_Compend, conMenu_Edit_Compend * 10# + 1 To conMenu_Edit_Compend * 10# + 3, conMenu_Edit_Compend * 10# + 6 '查阅、打印报告
        If Not gobjExchange Is Nothing Then
            Control.Enabled = blnAdvice And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_报告项)) <> 0
        Else
            Control.Enabled = blnAdvice And (Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_报告ID)) <> 0 Or vsAdvice.TextMatrix(vsAdvice.Row, COL_检查报告ID) <> "" Or Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_LIS报告ID)) <> 0)
        End If
        
        If Control.ID = conMenu_Edit_Compend * 10# + 1 Then
            If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_报告项)) = 1 Then
                Control.Enabled = True
            Else
                Control.Enabled = False
            End If
        ElseIf Control.ID = conMenu_Edit_Compend * 10# + 6 Then
            If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_报告项)) = 2 Then
                Control.Enabled = True
            Else
                Control.Enabled = False
            End If
        End If
    Case conMenu_Edit_Compend * 10# + 4 '我已经查阅该报告
        Control.Checked = Val(vsAdvice.Cell(flexcpData, vsAdvice.Row, COL_查阅状态)) = 1
        Control.Enabled = blnAdvice And (Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_报告ID)) <> 0 Or vsAdvice.TextMatrix(vsAdvice.Row, COL_检查报告ID) <> "")
    Case conMenu_Edit_Compend * 10# + 5 '自动标记查阅状态
        Control.Checked = mblnAutoRead
        Control.Enabled = mblnAutoReadEnabled
    Case conMenu_Edit_MarkMap, conMenu_Edit_MarkKeyMap, conMenu_Edit_ViewPacs '观片处理
        blnEnabled = blnAdvice And InStr(",4,5,6,7,8,9,H,M,Z,", vsAdvice.TextMatrix(vsAdvice.Row, COL_诊疗类别)) = 0 ' And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_报告ID)) <> 0
        If blnEnabled Then
            If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱状态)) <> 8 Then
                blnEnabled = False
            End If
        End If
        Control.Enabled = blnEnabled
    End Select

    Select Case Control.ID
    Case conMenu_Report_ClinicBill '打印诊疗单据
        Control.Enabled = Control.CommandBar.Controls.Count > 0
    End Select
    
    '电子签名部份
    '------------------------------------------------------------------------------
    Select Case Control.ID
    Case conMenu_Tool_SignNew '医嘱签名
        Control.Enabled = mlng病人ID <> 0
    Case conMenu_Tool_SignVerify, conMenu_Tool_SignEarse '验证签名,取消签名
        blnEnabled = mlng病人ID <> 0 And blnAdvice And tbcAppend.Selected.Tag = "签名" And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) <> 0
        If blnEnabled Then blnEnabled = vsAppend.RowData(vsAppend.Row) <> 0
        Control.Enabled = blnEnabled
    End Select
    
    '其它部份
    '------------------------------------------------------------------------------
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = blnAdvice
    Case conMenu_View_Append '附加信息
        Control.Checked = tbcAppend.Visible
    Case conMenu_View_Hide '自动隐藏过滤工具栏
        Control.Checked = mblnHideFilter
    Case conMenu_Manage_ReportLisView, conMenu_Manage_ReportPacsView '检查，检验结查
        Control.Enabled = mlng病人ID <> 0
    End Select
    
    '输血执行单打印
    If Control.ID = conMenu_Report_BloodInstant Then '执行单打印
        Control.Visible = InStr(GetInsidePrivs(9005, , 2200), ";输血执行打印;") <> 0
        Control.Enabled = vsAdvice.TextMatrix(vsAdvice.Row, COL_诊疗类别) = "K" And Control.Visible
    End If
    
End Sub

Private Sub SetControlVisible(ByVal Control As CommandBarControl)
'功能：根据权限设置菜单和工具栏的可见状态
    Dim blnVisible As Boolean, strItem As String

    '权限只需判断一次,已经判断过的命令不用再判断
    If Control.Category = "已判断" Then Exit Sub

    blnVisible = True
    
    '身份权限判断
    '------------------------------------------------------------------------------
    If InStr(UserInfo.性质, "医生") = 0 Then
        If Control.ID = conMenu_EditPopup Then blnVisible = False
        If Control.ID = conMenu_ReportPopup Then blnVisible = False
        If Between(Control.ID, conMenu_Edit_NewItem, conMenu_Edit_NewItem + 999) Then blnVisible = False
    End If

    '医嘱操作部份
    '------------------------------------------------------------------------------
    Select Case Control.ID
    Case conMenu_Edit_NewItem, conMenu_Edit_Modify, conMenu_Edit_Delete, conMenu_Edit_LISApply, conMenu_Edit_ApplyModi, conMenu_Edit_ApplyDel
        '新开医嘱,修改医嘱,删除医嘱,检验申请、修改、删除
        If InStr(GetInsidePrivs(p门诊医嘱下达), "医嘱下达") = 0 Then blnVisible = False
        
    Case conMenu_Edit_Blankoff
        '医嘱作废
        If InStr(GetInsidePrivs(p门诊医嘱下达), "医嘱作废") = 0 Then blnVisible = False
    Case conMenu_Edit_Send, conMenu_Edit_Send * 100# + 1
        '医嘱发送
        If mSendControl Is Nothing Then Set mSendControl = Control
        If InStr(GetInsidePrivs(p门诊医嘱下达), "医嘱发送") = 0 Then
            blnVisible = False
        ElseIf InStr(GetInsidePrivs(p门诊医嘱下达), "发送为收费单") = 0 And InStr(GetInsidePrivs(p门诊医嘱下达), "发送为记帐单") = 0 Then
            blnVisible = False
        End If
        If Val(zlDatabase.GetPara("发送单据类型", glngSys, p门诊医嘱下达)) = 0 And InStr(GetInsidePrivs(p门诊医嘱下达), "发送为收费单") = 0 Or _
           Val(zlDatabase.GetPara("发送单据类型", glngSys, p门诊医嘱下达)) = 1 And InStr(GetInsidePrivs(p门诊医嘱下达), "发送为记帐单") = 0 Then
            blnVisible = False
        End If
    Case conMenu_Edit_Test
        '皮试医嘱结果
        If InStr(GetInsidePrivs(p门诊医嘱下达), "皮试医嘱结果") = 0 Then blnVisible = False
    Case conMenu_Edit_TraReaction
        '输血反应登记
        If Not (InStr(GetInsidePrivs(9005, , 2200), "输血反应登记") <> 0 And gbln血库系统) Then blnVisible = False
    
    Case conMenu_Edit_Compend, conMenu_Edit_Compend * 10# + 1
        '报告弹出(含打印),查阅报告
        If InStr(GetInsidePrivs(p门诊医嘱下达), ";报告查阅;") = 0 Then blnVisible = False
    Case conMenu_Edit_Compend * 10# + 2, conMenu_Edit_Compend * 10# + 3
        '打印报告
        If InStr(GetInsidePrivs(p门诊医嘱下达), ";报告打印;") = 0 Then blnVisible = False
    Case conMenu_Edit_ViewDrugExplain '查看药品说明书
        If gobjDrugExplain Is Nothing Or InStr(GetInsidePrivs(p门诊医嘱下达), ";药品说明书;") = 0 Then blnVisible = False
    Case conMenu_Edit_MarkMap, conMenu_Edit_MarkKeyMap, conMenu_Edit_ViewPacs
        '观片处理
        If GetInsidePrivs(pXWPACS观片) <> "" And InStr(GetInsidePrivs(p门诊医嘱下达), ";观片处理;") <> 0 Then
            blnVisible = True
        Else
            If Control.ID = conMenu_Edit_MarkMap Or Control.ID = conMenu_Edit_ViewPacs Then
                If InStr(GetInsidePrivs(p门诊医嘱下达), ";观片处理;") = 0 Or GetInsidePrivs(p观片工具管理) = "" Then
                    blnVisible = False
                End If
            Else
                blnVisible = False
            End If
        End If
    Case conMenu_Edit_MediAudit, conMenu_Edit_MediAudit * 10# To conMenu_Edit_MediAudit * 10# + 99
        '合理用药审查
        strItem = GetInsidePrivs(p门诊医嘱下达)
        If InStr(strItem, "合理用药监测") = 0 Then blnVisible = False
    Case conMenu_Edit_AdvicePay
        blnVisible = InStr(GetInsidePrivs(p门诊医嘱下达), ";诊间无卡支付;") > 0
    End Select
        
    '电子签名部份
    Control.Category = "已判断"
    '------------------------------------------------------------------------------
    Select Case Control.ID
    Case conMenu_Tool_Sign, conMenu_Tool_SignNew '电子签名,医嘱签名
        If InStr(UserInfo.性质, "医生") = 0 Or gobjESign Is Nothing _
            Or InStr(GetInsidePrivs(p门诊医嘱下达), ";医嘱下达;") = 0 Then
            blnVisible = False
        ElseIf mblnSignVisible = False Then
            blnVisible = False '不同场合没有设置要使用签名
        End If
        Control.Category = ""  '签名按钮动态判断可见性
    End Select

    Control.Enabled = blnVisible
    Control.Visible = blnVisible
End Sub

Public Sub zlRefresh(ByVal lng病人ID As Long, ByVal str挂号单 As String, ByVal blnEditable As Boolean, _
        Optional ByVal blnMoved As Boolean, Optional ByVal lng前提ID As Long, Optional ByVal lng界面科室ID As Long, _
    Optional ByRef objMip As Object, Optional ByVal lng前提科室ID As Long, Optional ByVal lng路径状态 As Long = -1, _
    Optional ByVal int就诊类型 As Integer)
'功能：刷新门诊医嘱数据
'参数：lng前提ID=当由医技站调用时传入
'      blnMoved=该病人的数据是否已转出
'      blnEditable=可否对病人医嘱进行编辑
'      objMip 消息对象
'      lng前提科室ID= lng前提ID这条医嘱对应的执行科室ID；当医技站调用且满这个条件时：lng界面科室ID<>lng前提科室ID  lng前提科室ID参数必须传入
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strMsg  As String
    Dim objControl As CommandBarControl
    Dim lng病人IDOld As Long, str挂号单Old As String
    Dim lng病人科室ID As Long
    Dim lngOld界面科室ID As Long

    
    lng病人IDOld = mlng病人ID
    str挂号单Old = mstr挂号单
    lngOld界面科室ID = mlng界面科室ID
    
    mlng病人ID = lng病人ID
    mstr挂号单 = str挂号单
    mlng前提ID = lng前提ID
    mlng界面科室ID = lng界面科室ID
    mblnEditable = blnEditable
    mblnCanRevoke = blnEditable
    mblnMoved = blnMoved
    mlng路径状态 = lng路径状态
    mint就诊类型 = int就诊类型
    If InitObjPublicExpense Then Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, mlng病人ID, 0, "", mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
    If Not (objMip Is Nothing) Then Set mclsMipModule = objMip
        
    '读取额外的信息
    If lng病人ID <> 0 And str挂号单Old <> mstr挂号单 Then
        strSQL = "Select A.ID,A.执行部门ID,A.执行状态,A.登记时间,C.险类,Nvl(Nvl(A.续诊科室ID,Decode(A.转诊状态,1,A.转诊科室ID,NULL)),A.执行部门ID) as 病人科室ID,a.门诊号,c.姓名,a.执行人" & _
            " From 病人挂号记录 A,病人信息 C Where C.病人id=A.病人id And A.NO=[1] And a.记录性质=1 And a.记录状态=1"
        If mblnMoved Then strSQL = Replace(strSQL, "病人挂号记录", "H病人挂号记录")
        
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlRefresh", mstr挂号单, mlng病人ID)
        If Not rsTmp.EOF Then
            mvRegDate = rsTmp!登记时间
            
            '作废医嘱的允许：为已诊(不限制为最后一次挂号，要限制为已诊，避免同一次多科室挂号就诊时，作废其他人的医嘱)
            If Not mblnCanRevoke Then
                If NVL(rsTmp!执行状态, 0) = 1 Then
                    mblnCanRevoke = True
                End If
            End If
            mlng挂号ID = Val(rsTmp!ID & "")
            mstr姓名 = rsTmp!姓名 & ""
            mstr门诊号 = rsTmp!门诊号 & ""
            
            '婴儿条件
            mlng挂号科室ID = Val("" & rsTmp!执行部门ID)
            lng病人科室ID = Val("" & rsTmp!病人科室id)
            mbln产科 = DeptIsWoman(rsTmp!执行部门ID)
            mint险类 = Val("" & rsTmp!险类)
            If mbln产科 Then
                '读取最近缺省值：-1=所有,0=病人,1-婴儿1
                mvarCond.婴儿 = Val(zlDatabase.GetPara("病人婴儿过滤", glngSys, p门诊医嘱下达, "0"))
            End If
            mstr接诊医生 = "" & rsTmp!执行人
        Else
            mlng挂号科室ID = 0
            mvRegDate = CDate("3000-01-01") '体检病人,医技站临时登记的病人（未挂号）
            mbln产科 = False
            mvarCond.婴儿 = -1
            mlng挂号ID = 0
            mstr接诊医生 = ""
        End If
        Call GetCriticalData
        On Error GoTo 0
    End If
    
    If (lngOld界面科室ID <> mlng界面科室ID Or lng病人IDOld <> mlng病人ID) And mlng前提ID <> 0 Then
        mstr前提IDs = Get医技科室医嘱IDs(mlng病人ID, mstr挂号单, IIF(0 = lng前提科室ID, mlng界面科室ID, lng前提科室ID), False, mlng前提ID)
    ElseIf mlng前提ID = 0 Then
        mstr前提IDs = ""
    End If
    
    If Visible And lng病人科室ID <> 0 Then
        mblnSignVisible = True
        If mint场合 = 0 Then
            If CheckSign(0, 0, mlng界面科室ID, lng病人科室ID, 1, False, gobjESign) = False Then
                mblnSignVisible = False '不同场合没有设置要使用签名
            End If
        ElseIf mint场合 = 2 Then
            If CheckSign(3, 0, mlng界面科室ID, lng病人科室ID, 1, False, gobjESign) = False Then
                mblnSignVisible = False
            End If
        End If
    End If
    
    If Not grsTube Is Nothing Then
        If grsTube.State = 1 Then grsTube.Close
        Set grsTube = Nothing
    End If
    
    If lng病人IDOld <> mlng病人ID Then
        If mblnPass Then
            Call zlPASSPati
            On Error Resume Next
            Call gobjPass.zlPassClearLight(mobjPassMap)    '初始化状态灯
            On Error GoTo 0
        End If
    End If
    
    '刷新数据
    Call RefreshData
    
    '执行自动插件功能：病人ID=0也调用，以实现如关闭界面
    If mlngPlugInID <> 0 And lng病人ID <> 0 And str挂号单Old <> mstr挂号单 Then
        Set objControl = mcbsMain.FindControl(, mlngPlugInID, , True)
        If Not objControl Is Nothing Then
            objControl.Execute
        End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub RefreshData()
'功能：刷新数据
    If mlng病人ID = 0 Then
        '清除医嘱清单
        Call ClearAdviceData
        Call ClearAppendData
        RaiseEvent StatusTextUpdate("")
    Else
        '显示医嘱清单
        Call LoadAdvice
        '显示医嘱金额
        Call ShowTotalMoney
    End If
End Sub

Private Sub Refresh报告()
'功能：在报告页面不同报告之间切换时界面的刷新，不重新读数据库设置表格的隐藏和显示
    Dim i As Long
    Dim blnTmp As Boolean
    Dim lng医嘱ID As Long
    If mvarCond.过滤模式 = 0 Then Exit Sub
    With vsAdvice
    
        lng医嘱ID = Val(.TextMatrix(.Row, COL_ID))      '记录当前行如果是在当前界面刷新医嘱行应该不变
        
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, COL_报告项)) <> 0 Then
                If mvarCond.报告 = 0 Then ' 全部
                    blnTmp = True
                ElseIf mvarCond.报告 = 1 Then ' 检查
                    blnTmp = .TextMatrix(i, COL_诊疗类别) = "D"
                ElseIf mvarCond.报告 = 2 Then '检验
                    blnTmp = (.TextMatrix(i, COL_诊疗类别) = "E" And .TextMatrix(i, COL_操作类型) = "6" Or .TextMatrix(i, COL_诊疗类别) = "C")
                ElseIf mvarCond.报告 = 3 Then ' 其它
                    blnTmp = Not (.TextMatrix(i, COL_诊疗类别) = "E" And .TextMatrix(i, COL_操作类型) = "6" Or .TextMatrix(i, COL_诊疗类别) = "D" Or .TextMatrix(i, COL_诊疗类别) = "C")
                End If
                
                If blnTmp And .RowHeight(i) < .RowHeightMin Then .RowHeight(i) = .RowHeightMin
                
                .RowHidden(i) = Not blnTmp
            Else
                .RowHidden(i) = True: .RowHeight(i) = 0
            End If
            '增加过滤未出的报告和已出的报告
            If .RowHidden(i) = False Then
                blnTmp = IIF(.TextMatrix(i, COL_查阅状态) = "未出", mvarCond.未出报告, mvarCond.已出报告)
                If blnTmp And .RowHeight(i) < .RowHeightMin Then .RowHeight(i) = .RowHeightMin
                .RowHidden(i) = Not blnTmp
            End If
        Next
    End With
    Call LocatedDefaultAdviceRow(lng医嘱ID)
    Call SetAdviceColVisible
End Sub

Private Sub Refresh处方()
'功能：在报告页面不同报告之间切换时界面的刷新，不重新读数据库设置表格的隐藏和显示
    Dim i As Long
    Dim blnTmp As Boolean
    Dim lng医嘱ID As Long

    If mvarCond.过滤模式 = 3 Then Exit Sub
    With vsAdvice
        .Redraw = flexRDNone
        For i = .Rows - 1 To .FixedRows Step -1
            If .RowHidden(i) Then
                .RowHidden(i) = False
            End If
            .TextMatrix(i, COL_处方打印) = .Cell(flexcpData, i, COL_处方打印)
            .TextMatrix(i, COL_处方号) = .Cell(flexcpData, i, COL_处方号)
            .TextMatrix(i, COL_处方预览) = .Cell(flexcpData, i, COL_处方预览)
            If Val(.TextMatrix(i, COL_ID)) = 0 Then
                .RemoveItem i
            End If
        Next
        lng医嘱ID = Val(.TextMatrix(.Row, COL_ID))      '记录当前行如果是在当前界面刷新医嘱行应该不变
        For i = .FixedRows To .Rows - 1
            Select Case mvarCond.医嘱
            Case 0
            Case 1
            Case Else
                If .TextMatrix(i, COL_诊疗类别) = "5" Or .TextMatrix(i, COL_诊疗类别) = "6" Or .TextMatrix(i, COL_诊疗类别) = "E" And Val(.TextMatrix(i, COL_操作类型)) = 4 Then
                    If mvarCond.医嘱 = 2 Then .RowHidden(i) = True
                Else
                    If mvarCond.医嘱 = 1 Then .RowHidden(i) = True
                End If
            
            End Select
            
            If .TextMatrix(i, COL_诊疗类别) = "5" Or .TextMatrix(i, COL_诊疗类别) = "6" Or .TextMatrix(i, COL_诊疗类别) = "E" And Val(.TextMatrix(i, COL_操作类型)) = 4 Then
                If mvarCond.医嘱 = 2 Then
                    .RowHidden(i) = True
                End If
            Else
                If mvarCond.医嘱 = 1 Then
                    .RowHidden(i) = True
                End If
            End If
        Next
        .Redraw = flexRDDirect
    End With
    Call LocatedDefaultAdviceRow(lng医嘱ID)
    Call SetAdviceColVisible
End Sub

Private Sub LocatedDefaultAdviceRow(Optional ByVal lng医嘱ID As Long)
'功能：医嘱清单的缺省定位，如果有医嘱id跟据医嘱id定位
    '缺省定位，当前选择的医嘱为显示行则定位，否则定位到最后一行。
    Dim i As Long
    
    With vsAdvice
        .Redraw = flexRDNone
        .Row = .Rows - 1
        If lng医嘱ID <> 0 Then
            lng医嘱ID = .FindRow(CStr(lng医嘱ID), , COL_ID)
            If lng医嘱ID <> -1 Then
                If Not .RowHidden(lng医嘱ID) Then .Row = lng医嘱ID
            End If
        End If
        If .RowHidden(.Row) Then    '定位到了隐藏行的处理
            For i = .Row + 1 To .Rows - 1
                If Not .RowHidden(i) Then .Row = i: Exit For
            Next
        End If
        If .RowHidden(.Row) Then
            For i = .Row - 1 To .FixedRows Step -1
                If Not .RowHidden(i) Then .Row = i: Exit For
            Next
        End If
        If .RowHidden(.Row) Then
            .AddItem "": .Row = .Rows - 1
        End If
        If .Row < .FixedRows Then
            .AddItem "": .Row = .Rows - 1
        End If
        .Col = .FixedCols
        Call .ShowCell(.Row, .Col)
        .Redraw = flexRDDirect
        Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col)
        .Refresh
    End With
End Sub

Public Sub zlItemRef()
'功能：调用诊疗参考
    Dim lng诊疗项目ID As Long, i As Long

    With vsAdvice
        If Val(.TextMatrix(.Row, COL_ID)) <> 0 Then
            If .TextMatrix(.Row, COL_诊疗类别) = "E" And (RowIs配方行(.Row) Or RowIs检验行(.Row)) Then
                lng诊疗项目ID = Get诊疗项目ID(Val(.TextMatrix(.Row, COL_ID)), True)
            Else
                lng诊疗项目ID = Get诊疗项目ID(Val(.TextMatrix(.Row, COL_ID)), False)
            End If
        End If
    End With
    
    'ToDo:诊疗措施参考
End Sub

Private Sub cbsSub_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
'''''''
    Dim objControl As CommandBarControl
    Dim arrBaby As Variant, i As Long
    Dim strTmp As String
    
    If CommandBar.Parent Is Nothing Then Exit Sub
    
    Select Case CommandBar.Parent.ID
    Case ID_婴儿
        strTmp = IIF(mvarCond.过滤模式 = 3, "报告", "医嘱")
        With CommandBar.Controls
            .DeleteAll
            Set objControl = .Add(xtpControlButton, ID_婴儿 * 100#, "所有" & strTmp)
            Set objControl = .Add(xtpControlButton, ID_婴儿 * 100# + 1, "病人" & strTmp): objControl.BeginGroup = True
            For i = 0 To 4
                Set objControl = .Add(xtpControlButton, ID_婴儿 * 100# + i + 2, "婴儿 " & (i + 1) & " " & strTmp)
                If i = 0 Then objControl.BeginGroup = True
            Next
        End With
    Case Else
        Call zlPopupCommandBars(CommandBar)
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call ActiveHotKey(KeyCode, Shift)
End Sub

Private Sub fraHide_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    timHide.Enabled = True
End Sub

Private Sub mfrmEdit_CheckInfectDisease(ByVal blnOnChek As Boolean, ByVal str疾病ID As String, ByVal str诊断Id As String, ByRef blnNo As Boolean)
'    If InStr(";" & GetPrivFunc(glngSys, p门诊病历管理) & ";", ";病历书写;") > 0 Then
        RaiseEvent CheckInfectDisease(blnOnChek, str疾病ID, str诊断Id, blnNo)
'    End If
End Sub

Private Sub mfrmEdit_EditDiagnose(ParentForm As Object, ByVal 挂号单 As String, Succeed As Boolean)
    RaiseEvent EditDiagnose(ParentForm, 挂号单, Succeed)
End Sub

Private Sub mfrmEdit_FormUnload(Cancel As Integer)
    If mlng危急值ID <> 0 Then
        Call GetCriticalData
    End If
    mlng危急值ID = 0
    If Not Cancel Then
        If mfrmEdit.mblnOK Then
            RaiseEvent RequestRefresh
            'Call LoadAdvice
            'Call ShowTotalMoney
        End If
        Set mfrmEdit = Nothing
        
        If Me.Visible Then
            Call BringWindowToTop(Me.hwnd)
        End If
    End If
    RaiseEvent Activate
End Sub

Private Sub mfrmSend_EditDiagnose(ParentForm As Object, ByVal 挂号单 As String, Succeed As Boolean)
    RaiseEvent EditDiagnose(ParentForm, 挂号单, Succeed)
End Sub

Private Sub mobjReport_AfterPrint(ByVal ReportNum As String)
    Dim strSQL As String
    
    '申请单据打印之后的处理
    If mstrBillPrint <> "" Then
        If Split(mstrBillPrint, ",")(0) = ReportNum Then
            strSQL = "Zl_诊疗单据打印_Insert('" & Split(mstrBillPrint, ",")(1) & "'," & Val(Split(mstrBillPrint, ",")(2)) & ",1,'" & UserInfo.姓名 & "')"
        End If
    End If
    
    On Error GoTo errH
    If strSQL <> "" Then
        zlDatabase.ExecuteProcedure strSQL, Me.Name
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub timHide_Timer()
'功能：处理过滤工具栏的自动显示和隐藏
    Dim vPos As PointAPI, vRect As RECT
    Static sngBegin As Single
    
    If Not mblnHideFilter Then
        timHide.Enabled = False
        sngBegin = 0: Exit Sub
    End If
    
    If sngBegin = 0 Then sngBegin = Timer
    GetCursorPos vPos
    
    If fraHide.Visible Then
        ScreenToClient Me.hwnd, vPos
        If vPos.X * Screen.TwipsPerPixelX >= fraHide.Left And vPos.X * Screen.TwipsPerPixelX <= fraHide.Left + fraHide.Width _
            And vPos.Y * Screen.TwipsPerPixelY >= fraHide.Top And vPos.Y * Screen.TwipsPerPixelY <= picMain.Top + fraHide.Top + fraHide.Height Then
            fraHide.BackColor = cbsSub.GetSpecialColor(XPCOLOR_SEPARATOR)
            If Timer - sngBegin >= 0.35 Then
                fraHide.BackColor = cbsSub.GetSpecialColor(STDCOLOR_BTNFACE)
                fraHide.Visible = False: cbsSub(2).Visible = True
                sngBegin = 0: cbsSub.RecalcLayout
            End If
        Else
            fraHide.BackColor = cbsSub.GetSpecialColor(STDCOLOR_BTNFACE)
            sngBegin = 0: timHide.Enabled = False
        End If
    ElseIf cbsSub(2).Visible Then
        cbsSub(2).GetWindowRect vRect.Left, vRect.Top, vRect.Right, vRect.Bottom
        If Not (vPos.X >= vRect.Left / Screen.TwipsPerPixelX And vPos.X <= vRect.Right / Screen.TwipsPerPixelX _
            And vPos.Y >= vRect.Top / Screen.TwipsPerPixelY And vPos.Y <= vRect.Bottom / Screen.TwipsPerPixelY) Then
            If Timer - sngBegin >= 1 Then
                sngBegin = 0: timHide.Enabled = False
                fraHide.Visible = True: cbsSub(2).Visible = False
                cbsSub.RecalcLayout
            End If
        Else
            sngBegin = 0
        End If
    End If
End Sub

Private Sub cbsSub_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim bln报告 As Boolean
    Dim bln处方 As Boolean
    If Control.ID <> 0 Then
        If cbsSub.FindControl(, Control.ID, True, True) Is Nothing Then Exit Sub
        vsColumn.Visible = False
    End If

    Select Case Control.ID
        Case ID_婴儿 * 100# '所有医嘱
            If mvarCond.婴儿 = -1 Then Exit Sub
            mvarCond.婴儿 = -1
            Call zlDatabase.SetPara("病人婴儿过滤", mvarCond.婴儿, glngSys, p门诊医嘱下达)
        Case ID_婴儿 * 100# + 1 To ID_婴儿 * 100# + 6 '病人、婴儿医嘱
            If mvarCond.婴儿 = Control.ID - ID_婴儿 * 100# - 1 Then Exit Sub
            mvarCond.婴儿 = Control.ID - ID_婴儿 * 100# - 1
            Call zlDatabase.SetPara("病人婴儿过滤", mvarCond.婴儿, glngSys, p门诊医嘱下达)
        Case ID_全部
            mvarCond.报告 = 0
        Case ID_检查
            mvarCond.报告 = 1
        Case ID_检验
            mvarCond.报告 = 2
        Case ID_未出报告
            If mvarCond.未出报告 Then
                If mvarCond.已出报告 Then
                    mvarCond.未出报告 = Not mvarCond.未出报告
                End If
            Else
                mvarCond.未出报告 = Not mvarCond.未出报告
            End If
        Case ID_已出报告
            If mvarCond.已出报告 Then
                If mvarCond.未出报告 Then
                    mvarCond.已出报告 = Not mvarCond.已出报告
                End If
            Else
                mvarCond.已出报告 = Not mvarCond.已出报告
            End If
        Case ID_其他
            mvarCond.报告 = 3
        Case ID_医嘱全部
            mvarCond.医嘱 = 0
        Case ID_医嘱处方
            mvarCond.医嘱 = 1
        Case ID_医嘱其他
            mvarCond.医嘱 = 2
        Case ID_废止
            mvarCond.废止 = Not mvarCond.废止
        Case ID_科内
            mvarCond.科内 = Not mvarCond.科内
        Case ID_简洁
            mvarCond.显示模式 = 0
        Case ID_完整
            mvarCond.显示模式 = 1
    End Select
    
    bln报告 = InStr("," & ID_未出报告 & "," & "," & ID_已出报告 & "," & "," & ID_全部 & "," & ID_检查 & "," & ID_检验 & "," & ID_其他 & ",", "," & Control.ID & ",") > 0
    
    bln处方 = InStr("," & ID_医嘱全部 & "," & ID_医嘱处方 & "," & ID_医嘱其他 & ",", "," & Control.ID & ",") > 0
    
    cbsSub.RecalcLayout
    If bln报告 Then
        Call Refresh报告
    ElseIf bln处方 Then
        Call Refresh处方
    Else
        Call RefreshData
    End If
    
End Sub

Private Sub cbsSub_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Control.Enabled = mlng病人ID <> 0
    If Not Control.Enabled Then Exit Sub
    
    Select Case Control.ID
        Case ID_全部
            Control.Checked = mvarCond.报告 = 0
            Control.IconId = IIF(Control.Checked, 90004, 90003)
            Control.Visible = mvarCond.过滤模式 = 3
        Case ID_检查
            Control.Checked = mvarCond.报告 = 1
            Control.IconId = IIF(Control.Checked, 90004, 90003)
            Control.Visible = mvarCond.过滤模式 = 3
        Case ID_检验
            Control.Checked = mvarCond.报告 = 2
            Control.IconId = IIF(Control.Checked, 90004, 90003)
            Control.Visible = mvarCond.过滤模式 = 3
        Case ID_其他
            Control.Checked = mvarCond.报告 = 3
            Control.IconId = IIF(Control.Checked, 90004, 90003)
            Control.Visible = mvarCond.过滤模式 = 3
        
        Case ID_医嘱全部
            Control.Checked = mvarCond.医嘱 = 0
            Control.IconId = IIF(Control.Checked, 90004, 90003)
            Control.Visible = mvarCond.过滤模式 = 0
        Case ID_医嘱处方
            Control.Checked = mvarCond.医嘱 = 1
            Control.IconId = IIF(Control.Checked, 90004, 90003)
            Control.Visible = mvarCond.过滤模式 = 0
        Case ID_医嘱其他
            Control.Checked = mvarCond.医嘱 = 2
            Control.IconId = IIF(Control.Checked, 90004, 90003)
            Control.Visible = mvarCond.过滤模式 = 0
        
        Case ID_婴儿 '婴儿医嘱条件
            If mbln产科 Then
                Control.Visible = True
                
                If mvarCond.婴儿 = -1 Then
                    Control.Caption = IIF(mvarCond.过滤模式 = 3, "所有报告", "所有医嘱")
                ElseIf mvarCond.婴儿 = 0 Then
                    Control.Caption = IIF(mvarCond.过滤模式 = 3, "病人报告", "病人医嘱")
                Else
                    Control.Caption = "婴儿 " & mvarCond.婴儿
                End If
            Else
                If mvarCond.婴儿 <> -1 Or Control.Visible Then
                    mvarCond.婴儿 = -1
                    Control.Visible = False
                    Call zlDatabase.SetPara("病人婴儿过滤", mvarCond.婴儿, glngSys, p门诊医嘱下达)
                End If
            End If
        Case ID_婴儿 * 100# '所有医嘱
            Control.Checked = mvarCond.婴儿 = -1
        Case ID_婴儿 * 100# + 1 To ID_婴儿 * 100# + 6 '病人、婴儿医嘱
            Control.Checked = mvarCond.婴儿 = Control.ID - ID_婴儿 * 100# - 1
        Case ID_废止
            Control.Checked = mvarCond.废止
            Control.IconId = IIF(Not Control.Checked, 90001, 90002)
        Case ID_科内
            If mint场合 <> 2 Then
                Control.Visible = False
            Else
                Control.Visible = True
                Control.Checked = mvarCond.科内
                Control.IconId = IIF(Not Control.Checked, 90001, 90002)
            End If
        Case ID_简洁
            Control.Checked = mvarCond.显示模式 = 0
            Control.IconId = IIF(Control.Checked, 90004, 90003)
            Control.Visible = mvarCond.过滤模式 <> 3
        Case ID_完整
            Control.Checked = mvarCond.显示模式 = 1
            Control.IconId = IIF(Control.Checked, 90004, 90003)
            Control.Visible = mvarCond.过滤模式 <> 3
            
        Case ID_未出报告
            Control.Checked = mvarCond.未出报告
            Control.IconId = IIF(Not Control.Checked, 90001, 90002)
            Control.Visible = mvarCond.过滤模式 = 3
        Case ID_已出报告
            Control.Checked = mvarCond.已出报告
            Control.IconId = IIF(Not Control.Checked, 90001, 90002)
            Control.Visible = mvarCond.过滤模式 = 3
    End Select
End Sub

Private Sub cbsSub_Resize()
    Dim BarHideH As Long, PriceH As Long
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    Call Me.cbsSub.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    If cbsSub.Count >= 2 Then
        If Not cbsSub(2).Visible Then BarHideH = fraHide.Height
    End If
    
    On Error Resume Next
    If fraMore.Visible Then
        fraMore.Tag = ""
        fraMore.Visible = False
    End If
    
    PriceH = IIF(tbcAppend.Visible, fraAdviceUD.Height + tbcAppend.Height, 0)
    
    fraHide.Left = lngLeft
    fraHide.Top = lngTop
    fraHide.Width = lngRight - lngLeft
    
    vsAdvice.Left = lngLeft
    vsAdvice.Top = lngTop + BarHideH
    vsAdvice.Width = lngRight - lngLeft
    vsAdvice.Height = lngBottom - lngTop - PriceH - BarHideH
    
    '列选择器
    With vsAdvice
        fraColSel.Left = .Left + (.ColWidth(COL_F标志) + .ColWidth(COL_F报告) - fraColSel.Width) / 2 + 30
        fraColSel.Top = .Top + (225 - fraColSel.Height) / 2 + 30
    End With
    
    fraAdviceUD.Left = lngLeft
    fraAdviceUD.Top = vsAdvice.Top + vsAdvice.Height
    fraAdviceUD.Width = vsAdvice.Width
    
    tbcAppend.Left = lngLeft
    tbcAppend.Top = fraAdviceUD.Top + fraAdviceUD.Height
    tbcAppend.Width = vsAdvice.Width
End Sub

Private Sub tbcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim blnTmp As Boolean
    If Not Me.Visible Then Exit Sub
    Select Case Item.Tag
    Case "医嘱"
        mvarCond.过滤模式 = 0
        mvarCond.医嘱 = 0
        mvarCond.报告 = 0
    Case "报告"
        mvarCond.过滤模式 = 3
        mvarCond.医嘱 = 0
        mvarCond.报告 = 0
    End Select

    If Item.Tag <> "" And mlng病人ID <> 0 Then
        Call AddToolBarInDoctor
        Call RefreshData
    End If
End Sub

Private Sub Form_Activate()
    If Me.Visible And vsAdvice.Enabled Then vsAdvice.SetFocus
    RaiseEvent Activate
End Sub

Private Sub vsAdvice_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
    If fraMore.Visible = True Then
        fraMore.Tag = ""
        fraMore.Visible = False
        PicAdviceDetail.Visible = False
    End If
End Sub

Private Sub vsAdvice_DblClick()
    Dim lng医嘱ID As Long
    Dim lngNo As Long
    Dim bln用血 As Boolean
    'PASS
    If mblnPass Then
        Call gobjPass.zlPassAdviceMainPoint(mobjPassMap)
    End If
    '双击的医嘱如果是申请单方式下达的弹出查看界面 输血，手术，会诊，检查，检验
    With vsAdvice
        lng医嘱ID = Val(.TextMatrix(.Row, COL_ID))
        lngNo = Val(.TextMatrix(.Row, COL_申请序号))
        
        If lng医嘱ID <> 0 And lngNo <> 0 Then
            If .TextMatrix(.Row, COL_诊疗类别) = "K" Then
                bln用血 = Val(.TextMatrix(.Row, COL_检查方法)) = 1
                '输血
                If Val(Mid(gstrOutUseApp, 3, 1)) = 1 Then
                    If gbln血库系统 = True Then
                        Call frmApplyBloodNew.ShowMe(Me, mlng病人ID, 0, 1, 2, lng医嘱ID, mlng挂号科室ID, , mlng挂号科室ID, , , mrsDefine, mclsMipModule, 1, mstr挂号单, , , , , mlng前提ID, IIF(bln用血 = True, 1, 0))
                    Else
                        Call frmApplyBlood.ShowMe(Me, mlng病人ID, 0, 1, 2, lng医嘱ID, mlng挂号科室ID, , mlng挂号科室ID, , , mrsDefine, mclsMipModule, 1, mstr挂号单, , , , , mlng前提ID)
                    End If
                End If
                                
            ElseIf .TextMatrix(.Row, COL_诊疗类别) = "F" Then
                '手术
                If Val(Mid(gstrOutUseApp, 4, 1)) = 1 Then Call frmApplyOperation.ShowMe(Me, 1, 2, mlng病人ID, mstr挂号单, 1, lng医嘱ID)
               
            ElseIf .TextMatrix(.Row, COL_诊疗类别) = "D" Then
                '检查
                If Val(Mid(gstrOutUseApp, 1, 1)) = 1 Then
                    Call ShowApply检查(Me, lngNo)
                End If
            ElseIf .TextMatrix(.Row, COL_诊疗类别) = "E" And .TextMatrix(.Row, COL_操作类型) = "6" Then
                '检验
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long, strPrompt As String
    
    With vsAdvice
        lngRow = .MouseRow
        If Button = 0 And lngRow > 0 Then  '简洁模式才显该列
            If .MouseCol = col_内容 Then
                If Val(fraMore.Tag) <> lngRow Then
                    fraMore.Visible = False
                    fraMore.Tag = lngRow
                    If lngRow = .Row Then
                        fraMore.BackColor = .BackColorSel
                    Else
                        fraMore.BackColor = .BackColor
                    End If
                    fraMore.Height = .RowHeight(lngRow) - 10
                    If fraMore.Height > 250 Then fraMore.Height = 250
                    
                    fraMore.Top = .Top + .RowPos(lngRow) + .RowHeight(lngRow) - fraMore.Height
                    If fraMore.Top + fraMore.Height > .Top + .Height - IIF(Grid.HScrollVisible(vsAdvice), 230, 0) Then Exit Sub
                    
                    fraMore.Left = .Left + .ColPos(col_内容) + IIF(.ColWidth(col_内容) > .ColWidthMax, .ColWidthMax, .ColWidth(col_内容)) - fraMore.Width
                    fraMore.Visible = True
                ElseIf PicAdviceDetail.Visible = True Then
                    fraMore.Tag = ""
                    fraMore.Visible = False
                    PicAdviceDetail.Visible = False
                End If
            Else
                If fraMore.Visible = True Then
                    fraMore.Tag = ""
                    fraMore.Visible = False
                    PicAdviceDetail.Visible = False
                End If
                
                strPrompt = ""
                If .MouseCol = COL_F标志 Then
                    If Val(.TextMatrix(lngRow, COL_标志)) = 1 Then
                        strPrompt = "紧急医嘱"
                    ElseIf Val(.TextMatrix(lngRow, COL_标志)) = 2 Then
                        strPrompt = "补录医嘱"
                    End If
                    '如果有抗菌用药审核信息，优先显示
                    If Val(.TextMatrix(lngRow, COL_医嘱状态)) = 1 Then
                        Select Case Val(.TextMatrix(lngRow, COL_审核状态))
                        Case 1
                            If .TextMatrix(lngRow, COL_诊疗类别) = "K" And Val(.TextMatrix(lngRow, COL_检查方法)) = 1 Then '用血医嘱审核
                                strPrompt = "用血医嘱待核对"
                            Else
                                strPrompt = Decode(.TextMatrix(lngRow, COL_诊疗类别), "F", "手术", "K", "输血", "抗菌用药") & "待审核"
                            End If
                        Case 2
                            If Not (.TextMatrix(lngRow, COL_诊疗类别) = "K" And Val(.TextMatrix(lngRow, COL_检查方法)) = 1) Then
                                strPrompt = Decode(.TextMatrix(lngRow, COL_诊疗类别), "K", "输血", "抗菌用药") & "审核通过"
                            End If
                        Case 3
                            strPrompt = Decode(.TextMatrix(lngRow, COL_诊疗类别), "K", "输血", "抗菌用药") & "审核未通过:" & GetKSSAuditQuestion(Val(.TextMatrix(lngRow, COL_ID)))
                        Case 4
                            If gbln血库系统 = False Then strPrompt = "输血待血库审核"
                        Case 5
                            If gbln血库系统 = False Then strPrompt = "输血血库正在配血"
                        Case 7
                            strPrompt = Decode(.TextMatrix(lngRow, COL_诊疗类别), "F", "手术", "K", "输血", "抗菌用药") & "待签发"
                        End Select
                    End If
                ElseIf .MouseCol = COL_查阅状态 Then
                    If Val(.TextMatrix(lngRow, COL_ID)) <> 0 Then strPrompt = "报告未出"
                    If Val(.TextMatrix(lngRow, COL_报告ID)) <> 0 Or .TextMatrix(lngRow, COL_检查报告ID) <> "" Or _
                        Val(.TextMatrix(lngRow, COL_RIS报告ID)) <> 0 Or Val(.TextMatrix(lngRow, COL_LIS报告ID)) <> 0 Then
                        
                        If Val(.Cell(flexcpData, lngRow, COL_查阅状态)) = 0 Then
                            strPrompt = "报告未阅，点击查看"
                        ElseIf Val(.Cell(flexcpData, lngRow, COL_查阅状态)) = 1 Then
                            strPrompt = "报告已阅，点击查看"
                        ElseIf Val(.Cell(flexcpData, lngRow, COL_查阅状态)) = 2 Then
                            strPrompt = "报告部分已阅，点击查看"
                        End If
                    End If
                ElseIf .MouseCol = COL_F报告 Then
                    strPrompt = GetAdviceReportTip(lngRow)
                End If
            End If
     
            If .MouseRow > -1 And .MouseCol > -1 And (mvarCond.过滤模式 = 3 And .MouseCol = COL_查阅状态 Or .MouseCol = COL_处方打印 Or .MouseCol = COL_处方预览) Then
                If .Cell(flexcpFontUnderline, .MouseRow, .MouseCol) = True And .TextMatrix(.MouseRow, .MouseCol) <> "" Then
                    .MousePointer = 99
                Else
                    .MousePointer = 0
                End If
            Else
                .MousePointer = 0
            End If
                        
            If strPrompt <> "" Then
                Call zlCommFun.ShowTipInfo(.hwnd, strPrompt)
                mlngPromptRow = lngRow
            ElseIf mlngPromptRow <> 0 And lngRow <> mlngPromptRow Then
            '隐藏之前的提示内容
                Call zlCommFun.ShowTipInfo(.hwnd, "")
                mlngPromptRow = 0
            End If
        End If
    End With
End Sub


Private Sub vsfAdivceDetail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fraMore.Tag = ""
    fraMore.Visible = False
    PicAdviceDetail.Visible = False
End Sub

Private Sub imgMore_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If PicAdviceDetail.Visible = False And vsAdvice.MouseRow > 0 Then
        Call LoadAdviceDetail(vsAdvice.MouseRow)
    End If
End Sub

Private Sub LoadAdviceDetail(lngRow As Long)
'功能：显示某行医嘱的详细内容
    Dim i As Long, j As Long
    
    vsfAdivceDetail.Redraw = flexRDNone
    vsfAdivceDetail.Clear
    vsfAdivceDetail.Rows = vsfAdivceDetail.FixedRows
    vsfAdivceDetail.Cols = 2
    j = 0
    With vsAdvice
        For i = 0 To .Cols - 1
             If .Cell(flexcpData, 0, i) = "Detail" Then
                j = j + 1
                vsfAdivceDetail.Rows = vsfAdivceDetail.FixedRows + j
                vsfAdivceDetail.TextMatrix(j - 1, 0) = .TextMatrix(0, i) & "："
                vsfAdivceDetail.TextMatrix(j - 1, 1) = .TextMatrix(lngRow, i)
                
                vsfAdivceDetail.Col = 0: vsfAdivceDetail.Row = j - 1
                vsfAdivceDetail.CellForeColor = &H8000000C
             End If
        Next
    End With
    With vsfAdivceDetail
        If .Rows > 0 Then
            .AutoSize 0, 1
            .Height = IIF(.RowHeight(0) < .RowHeightMin, .RowHeightMin, .RowHeight(0)) * .Rows + 100
            .Width = .ColWidth(0) + .ColWidth(1)
            .Row = -1
            
            PicAdviceDetail.Height = .Height
            PicAdviceDetail.Width = .Width
            PicAdviceDetail.Left = fraMore.Left + fraMore.Width
            
            If PicAdviceDetail.Height + fraMore.Top + fraMore.Height > Me.Top + Me.Height Then
                PicAdviceDetail.Top = fraMore.Top + fraMore.Height - PicAdviceDetail.Height - 10
            Else
                PicAdviceDetail.Top = fraMore.Top - 10  '避免顶端和表格线重合
            End If
            
            Call SetPicAdviceDetailEffect
            If PicAdviceDetail.Visible = False Then PicAdviceDetail.Visible = True
        End If
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub SetPicAdviceDetailEffect()
    Dim lngR As Long
    
    '边框：API=RoundRect
    PicAdviceDetail.Line (Screen.TwipsPerPixelX, 0)-(PicAdviceDetail.Width - Screen.TwipsPerPixelX, 0), RGB(118, 118, 118)
    PicAdviceDetail.Line (Screen.TwipsPerPixelX, PicAdviceDetail.Height - Screen.TwipsPerPixelY)-(PicAdviceDetail.Width - Screen.TwipsPerPixelX, PicAdviceDetail.Height - Screen.TwipsPerPixelY), RGB(118, 118, 118)
    PicAdviceDetail.Line (0, Screen.TwipsPerPixelY)-(0, PicAdviceDetail.Height - Screen.TwipsPerPixelY), RGB(118, 118, 118)
    PicAdviceDetail.Line (PicAdviceDetail.Width - Screen.TwipsPerPixelX, Screen.TwipsPerPixelY)-(PicAdviceDetail.Width - Screen.TwipsPerPixelX, PicAdviceDetail.Height - Screen.TwipsPerPixelY), RGB(118, 118, 118)
    PicAdviceDetail.PSet (Screen.TwipsPerPixelX, Screen.TwipsPerPixelY), RGB(186, 186, 186)
    PicAdviceDetail.PSet (PicAdviceDetail.Width - Screen.TwipsPerPixelX * 2, Screen.TwipsPerPixelY), RGB(186, 186, 186)
    PicAdviceDetail.PSet (Screen.TwipsPerPixelX, PicAdviceDetail.Height - Screen.TwipsPerPixelY * 2), RGB(186, 186, 186)
    PicAdviceDetail.PSet (PicAdviceDetail.Width - Screen.TwipsPerPixelX * 2, PicAdviceDetail.Height - Screen.TwipsPerPixelY * 2), RGB(186, 186, 186)
           
    '形状
    lngR = CreateRoundRectRgn(0, 0, PicAdviceDetail.ScaleX(PicAdviceDetail.Width, PicAdviceDetail.ScaleMode, vbPixels) + 1, PicAdviceDetail.ScaleY(PicAdviceDetail.Height, PicAdviceDetail.ScaleMode, vbPixels) + 1, 2, 2)
    Call SetWindowRgn(PicAdviceDetail.hwnd, lngR, False)
    
End Sub

Private Sub vsfAdivceDetail_LostFocus()
    PicAdviceDetail.Visible = False
End Sub

Private Sub fraAdviceUD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If vsAdvice.Height + Y < 1000 Or tbcAppend.Height - Y < 500 Then Exit Sub
        fraAdviceUD.Top = fraAdviceUD.Top + Y
        vsAdvice.Height = vsAdvice.Height + Y
        tbcAppend.Top = tbcAppend.Top + Y
        tbcAppend.Height = tbcAppend.Height - Y
        Me.Refresh
    End If
End Sub

Private Sub imgColSel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    
    If Button = 1 Then '列选择器
        '根据当前状态直接确定勾选状态
        With vsColumn
            If .Visible Then
                .Visible = False
                vsAdvice.SetFocus
            Else
                For i = .FixedRows To .Rows - 1
                    If vsAdvice.ColHidden(.RowData(i)) Or vsAdvice.ColWidth(.RowData(i)) = 0 Then
                        .TextMatrix(i, 0) = 0
                    Else
                        .TextMatrix(i, 0) = 1
                    End If
                    '设置处方号的列的显示方式
                    If .TextMatrix(i, 1) = "处方号" Or .TextMatrix(i, 1) = "打印" Or .TextMatrix(i, 1) = "预览" Then
                        If mvarCond.医嘱 = 1 Then
                            .TextMatrix(i, 0) = 1
                        Else
                            .TextMatrix(i, 0) = 0
                        End If
                        .Cell(flexcpForeColor, i, 0, i, 1) = .BackColorFixed
                    End If
                Next
                
                vsColumn.Height = vsColumn.RowHeightMin * vsColumn.Rows + 150
                If vsColumn.Top + vsColumn.Height > Me.ScaleHeight Then
                    vsColumn.Height = Me.ScaleHeight - vsColumn.Top
                    vsColumn.Width = 1750
                Else
                    vsColumn.Width = 1470
                End If
                
                .Left = fraColSel.Left
                .Top = fraColSel.Top + fraColSel.Height
                .ZOrder
                .Visible = True
                .SetFocus
            End If
        End With
    End If
End Sub

Private Function CheckWindow() As Boolean
'功能：检查医嘱编辑窗口是否已经打开
    If Not mfrmEdit Is Nothing Then
        '当前窗口打开了
        MsgBox "医嘱编辑窗口已经打开，请先完成当前操作后再执行。", vbInformation, gstrSysName
        '定位到当前的窗口
        If mfrmEdit.WindowState = vbMinimized Then mfrmEdit.WindowState = vbNormal
        mfrmEdit.SetFocus
        Exit Function
    Else
        '其它窗口打开了
        If Not CheckAdviceWindow("门诊医嘱编辑") Then Exit Function
    End If
    CheckWindow = True
End Function

Private Sub FuncBillPrint(Optional objControl As CommandBarControl, Optional ByVal strPar As String, Optional strName As String)
'功能：打印诊疗单据
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strNO As String, lng记录性质 As Long

    Dim lng相关ID As Long
    Dim strParameter As String
    Dim strErr As String
    Dim blnDo As Boolean
    Dim strBillName As String '诊疗单据的名称  病历文件列表.名称
    
    If Not objControl Is Nothing Then strPar = objControl.Parameter: strName = objControl.Caption
    If strPar = "" Then Exit Sub

    If InStr(strPar, "|") > 0 Then strParameter = Split(strPar, "|")(0): strNO = Split(strPar, "|")(1)
    
    strBillName = strName
    strBillName = Replace("<Tab>" & strBillName, "<Tab>打印:", "")
    If InStr(strBillName, "(&") > 0 Then
        strBillName = Mid(strBillName, 1, InStr(strBillName, "(&") - 1)
    End If
    
    With vsAdvice
        '打印次数提示
        On Error GoTo errH
        lng相关ID = Decode(Val(.TextMatrix(.Row, COL_相关ID)), 0, Val(.TextMatrix(.Row, COL_ID)), Val(.TextMatrix(.Row, COL_相关ID)))
        If .TextMatrix(.Row, COL_诊疗类别) = "E" And Val(.TextMatrix(.Row, COL_操作类型)) = 6 Then
            If Not gobjLIS Is Nothing Then '打印检验申请单据
                blnDo = gobjLIS.CheckAcceptance(CStr(lng相关ID), strErr)
                If Not blnDo Then
                   MsgBox "该标本已经被检验科核收，不能打印:" & strBillName & "。", vbInformation, gstrSysName
                   Exit Sub
                End If
            End If
        End If
        If strNO <> "" Then
            strSQL = "Select A.NO,A.记录性质 from 病人医嘱发送 A,病人医嘱记录 B Where a.医嘱ID=b.id And a.NO=[2] And (b.ID=[3] Or b.相关ID=[3])"
        Else
            strSQL = "Select NO,记录性质 from 病人医嘱发送 Where 医嘱ID=[1]"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(.TextMatrix(.Row, COL_ID)), strNO, lng相关ID)
        If rsTmp.RecordCount > 0 Then
            strNO = rsTmp!NO & ""
            lng记录性质 = Val(rsTmp!记录性质 & "")
            strSQL = "Select 打印人,打印时间 From 诊疗单据打印 Where NO=[1] And 记录性质=[2] And 打印内容=1 Order by 打印时间 Desc"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, strNO, lng记录性质)
            If Not mbln处方预览 Then
                If Not rsTmp.EOF Then
                    If MsgBox("该[" & strBillName & "]已经打印了 " & rsTmp.RecordCount & " 次，最近一次由""" & _
                        rsTmp!打印人 & """在""" & Format(rsTmp!打印时间, "yyyy-MM-dd HH:mm") & """打印。" & vbCrLf & vbCrLf & "要继续打印吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Sub
                    End If
                End If
            End If
            On Error GoTo 0
            SwitchPrintSet glngSys & "\" & p门诊医嘱下达
            '调用打印
            If mobjReport.ReportPrintSet(gcnOracle, glngSys, strParameter, mfrmParent) Then
                mstrBillPrint = strParameter & "," & strNO & "," & lng记录性质
                Call mobjReport.ReportOpen(gcnOracle, glngSys, strParameter, mfrmParent, "NO=" & strNO, "性质=" & lng记录性质, IIF(mbln处方预览, 1, 2))
                mstrBillPrint = ""
            End If
            SwitchPrintSet glngSys & "\" & p门诊医嘱下达, True
        End If
    End With
    mbln处方预览 = False
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncAdviceSign()
'功能：对医嘱进行电子签名
    Dim strSQL As String, strIDs As String, i As Long
    Dim strSource As String, strSign As String
    Dim lng签名id As Long, lng证书ID As Long
    Dim intRule As Integer, strTimeStamp As String, strTimeStampCode As String
    Dim ColIDs As Collection, ColSource As Collection
    
    If Not mblnEditable Then Exit Sub
    If gobjESign Is Nothing Then Exit Sub
    If gobjESign.CertificateStoped(UserInfo.姓名) Then
        MsgBox "您的签名证书已被停用，请联系信息科。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '获取签名医嘱源文
    intRule = ReadAdviceSignSource(1, mlng病人ID, mstr挂号单, strIDs, 0, mblnMoved, strSource, mstr前提IDs, , ColIDs, ColSource)
    If intRule = 0 Then Exit Sub
    If strSource = "" Then
        MsgBox "该病人目前没有可以签名的医嘱。", vbInformation, gstrSysName
        Exit Sub
    End If
    For i = 1 To ColIDs.Count
        strSign = gobjESign.Signature(ColSource(i), gstrDBUser, lng证书ID, strTimeStamp, Nothing, strTimeStampCode)
        If strSign <> "" Then
            If strTimeStamp <> "" Then
                strTimeStamp = "To_Date('" & strTimeStamp & "','YYYY-MM-DD HH24:MI:SS')"
            Else
                strTimeStamp = "NULL"
            End If
            lng签名id = zlDatabase.GetNextID("医嘱签名记录")
            strSQL = "zl_医嘱签名记录_Insert(" & lng签名id & ",1," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng证书ID & ",'" & ColIDs(i) & "'," & strTimeStamp & ",'" & UserInfo.姓名 & "','" & strTimeStampCode & "')"
            On Error GoTo errH
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            On Error GoTo 0
        End If
    Next
    If strSign <> "" Then
        Call LoadAdvice '刷新界面
        MsgBox "已完成电子签名。", vbInformation, gstrSysName
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncAdviceSignErase()
'功能：取消医嘱的电子签名
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        
    If Not mblnEditable Then Exit Sub
    If gobjESign Is Nothing Then Exit Sub
    If gobjESign.CertificateStoped(UserInfo.姓名) Then
        MsgBox "您的签名证书已被停用，请联系信息科。", vbInformation, gstrSysName
        Exit Sub
    End If
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) = 0 Or tbcAppend.Selected.Tag <> "签名" Then Exit Sub
    
    With vsAppend
        If .RowData(.Row) = 0 Then
            MsgBox "当前选择的医嘱没有签名。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '作废签名不能取消
        If .Cell(flexcpData, .Row, 0) = 4 Then
            MsgBox "作废医嘱的签名不能取消。", vbInformation, gstrSysName
            Exit Sub
        End If
        '新开签名必须是在新开状态
        If .Cell(flexcpData, .Row, 0) = 1 Then
            If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱状态)) <> 1 Then
                MsgBox "由于医嘱已经发送或作废，该签名不能取消。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        '只能取消自已签的名
        If .TextMatrix(.Row, 2) <> UserInfo.姓名 Then
            MsgBox "该签名人不是你本人，不能取消签名。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If MsgBox("确实要取消这次签名吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        If Not gobjESign.CheckCertificate(gstrDBUser) Then Exit Sub
        
        strSQL = "zl_医嘱签名记录_Delete(" & .RowData(.Row) & ")"
        On Error GoTo errH
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        On Error GoTo 0
    End With
    
    Call LoadAdvice '刷新界面
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncAdviceSignVerify()
'功能：校验医嘱的电子签名(可对已转移的数据)
    Dim strSource As String
    
    If gobjESign Is Nothing Then Exit Sub
    If gobjESign.CertificateStoped(UserInfo.姓名) Then
        MsgBox "您的签名证书已被停用，请联系信息科。", vbInformation, gstrSysName
        Exit Sub
    End If
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) = 0 Or tbcAppend.Selected.Tag <> "签名" Then Exit Sub
    
    With vsAppend
        If .RowData(.Row) = 0 Then
            MsgBox "当前选择的医嘱没有签名。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '获取签名医嘱源文
        If ReadAdviceSignSource(.Cell(flexcpData, .Row, 0), 0, 0, "", .RowData(.Row), mblnMoved, strSource) = 0 Then Exit Sub
        
        '验证签名
        Call gobjESign.VerifySignature(strSource, .RowData(.Row), 1)
    End With
End Sub

Private Sub FuncAdviceAdd()
'功能：新增医嘱
    If Not CheckWindow Then Exit Sub
        '检查挂号病人是否超期
    If Not FuncTimeLimitCheck Then Exit Sub
    
    If Not FuncPathAdd() Then Exit Sub
    
    Set mfrmEdit = frmOutAdviceEdit
    Call frmOutAdviceEdit.ShowMe(mfrmParent, mint场合, mMainPrivs, mlng病人ID, mstr挂号单, mlng前提ID, , , mblnModalNew, mlng界面科室ID, mstr前提IDs, mclsMipModule, mlng挂号科室ID, mblnMoved, , mlng危急值ID)
End Sub

Private Sub FuncAdviceDel()
'删除：删除当前医嘱
'说明：在主界面删除,对检查组合,手术组合,中药配方,是整个删除,一并给药只删除当前药品
    Dim strSQL As String, lng医嘱ID As Long
    Dim blnGroup As Boolean, i As Long, blnBat As Boolean, blnTrans As Boolean, arrSQL As Variant
    Dim lngRow As Long, strXML As String, lng申请序号 As Long
    Dim strDelIDs As String, arrDelID() As String
    Dim strDelDrugIDs As String         '记录删除的药品医嘱,用于传入合理用药监测
    Dim lng组ID As Long
    Dim blnRIS预约 As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim bln处方已审查 As Boolean
    Dim bln输血 As Boolean, strErr As String
    
    If Not mblnEditable Then Exit Sub
    
    '不检查挂号单有效天数，因为超过挂号单有效天数的病人，必须删除未发送的医嘱后才能完成接诊。
    
    With vsAdvice
        '检查是否可以删除
        lng医嘱ID = Val(.TextMatrix(.Row, COL_ID))
        If lng医嘱ID = 0 Then
            MsgBox "该病人没有医嘱可以删除。", vbInformation, gstrSysName
            Exit Sub
        End If
        If InStr(",5,6,", "," & .TextMatrix(.Row, COL_诊疗类别) & ",") > 0 Then
            strDelDrugIDs = "【西药】" & lng医嘱ID
        ElseIf .TextMatrix(.Row, COL_诊疗类别) = "E" And .TextMatrix(.Row, COL_操作类型) = "4" Then
            strDelDrugIDs = "【中药】" & .Cell(flexcpData, .Row, COL_相关ID)
        End If
        '医技下达的医嘱
        If mint场合 = 2 Then
            If InStr("," & mstr前提IDs & ",", "," & .TextMatrix(.Row, COL_前提ID) & ",") = 0 Then
                MsgBox "该医嘱不为当前医技科室下达，不能删除该医嘱。", vbInformation, gstrSysName
                Exit Sub
            ElseIf Val(.TextMatrix(.Row, COL_前提ID)) = 0 Then
                MsgBox "该医嘱不是医技科室下达，不能删除该医嘱。", vbInformation, gstrSysName
                Exit Sub
            End If
        ElseIf Val(.TextMatrix(.Row, COL_前提ID)) <> 0 Then
            MsgBox "该医嘱为医技科室下达，不能删除该医嘱。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If Val(.TextMatrix(.Row, COL_医嘱状态)) <> 1 Then
            MsgBox "当前选择的医嘱已经发送或作废，不能删除。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '已签名的医嘱不能删除
        If Val(.TextMatrix(.Row, COL_签名否)) = 1 Then
            MsgBox "当前选择的医嘱已经签名，不能删除。请先取消签名。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If mlng路径状态 = 1 Then
            If CheckPathAdviceIsExeOut(lng医嘱ID) Then
                MsgBox "该医嘱对应的项目已经执行。" & vbCrLf & "请取消执行登记后再进行删除操作。", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
                Exit Sub
            End If
        End If
        '启用血库系统输血医嘱删除限制，进入血库审核阶段的新开医嘱不能删
        bln输血 = gbln血库系统 And .TextMatrix(.Row, COL_诊疗类别) = "K"
        If gbln血库系统 And .TextMatrix(.Row, COL_诊疗类别) = "K" And InStr("5,2", Val(.TextMatrix(.Row, COL_审核状态))) > 0 Then
            MsgBox "该输血医嘱已被血库接收" & IIF(Val(.TextMatrix(.Row, COL_审核状态)) = 5, "正在配血", "并且已完成配血") & "，不能删除，若需删除请与输血科联系。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        arrSQL = Array()
        
        If InStr(",5,6,", .TextMatrix(.Row, COL_诊疗类别)) > 0 Then
            If .Row - 1 >= .FixedRows Then
                If Val(.TextMatrix(.Row - 1, COL_相关ID)) = Val(.TextMatrix(.Row, COL_相关ID)) Then blnGroup = True
            End If
            If Not blnGroup And .Row + 1 <= .Rows - 1 Then
                If Val(.TextMatrix(.Row + 1, COL_相关ID)) = Val(.TextMatrix(.Row, COL_相关ID)) Then blnGroup = True
            End If
            If blnGroup Then
                lng组ID = Val(.TextMatrix(.Row, COL_相关ID))
                If MsgBox("医嘱""" & .TextMatrix(.Row, col_医嘱内容) & """与其它药品一并给药,确实要删除吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Else
                If MsgBox("确实要删除医嘱""" & .TextMatrix(.Row, col_医嘱内容) & """吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
            If mblnPass Then
                Call gobjPass.zlPassAdviceDel(mobjPassMap, lng医嘱ID, zlDatabase.Currentdate)
            End If
        
        ElseIf .TextMatrix(.Row, COL_申请序号) <> "" Then
            If .TextMatrix(.Row, COL_诊疗类别) = "K" Then
                If MsgBox("确实要取消输血申请""" & .TextMatrix(.Row, col_医嘱内容) & """吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            ElseIf .TextMatrix(.Row, COL_诊疗类别) = "F" Then
                If MsgBox("确实要取消手术申请""" & .TextMatrix(.Row, col_医嘱内容) & """吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Else
                If MsgBox("要将""" & .TextMatrix(.Row, col_医嘱内容) & """同时申请的其他项目一起取消吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    blnBat = True
                End If
            End If
        Else
            If MsgBox("确实要删除医嘱""" & .TextMatrix(.Row, col_医嘱内容) & """吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        
        If .TextMatrix(.Row, COL_诊疗类别) = "D" Then
            If HaveRIS And gbln启用影像信息系统预约 Then
                blnRIS预约 = True
            End If
        End If
        
        Call CreatePlugInOK(p门诊医嘱下达, mint场合)
        If blnBat Then
            For i = 1 To .Rows - 1
                lng申请序号 = Val(.TextMatrix(.Row, COL_申请序号))
                If .TextMatrix(i, COL_医嘱状态) = "1" And Val(.TextMatrix(i, COL_申请序号)) = lng申请序号 Then
                    '调用删除前外挂接口
                    On Error Resume Next
                    If Not gobjPlugIn Is Nothing Then
                        If gobjPlugIn.AdviceDeletBefor(glngSys, p门诊医嘱下达, mlng病人ID, mlng挂号ID, Val(.TextMatrix(i, COL_ID)), mint场合) = False Then
                            If err.Number = 0 Then Exit Sub
                        End If
                        Call zlPlugInErrH(err, "AdviceDeletBefor")
                    End If
                                        If Not CheckDelAdivceOfPathItem(Val(.TextMatrix(.Row, COL_ID))) Then Exit Sub
                    If err.Number <> 0 Then err.Clear
                    On Error GoTo 0
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_Delete(" & .TextMatrix(i, COL_ID) & ",1)"
                    strDelIDs = strDelIDs & "," & .TextMatrix(i, COL_ID)
                End If
            Next
        Else
            '调用删除前外挂接口
            On Error Resume Next
            If Not gobjPlugIn Is Nothing Then
                If gobjPlugIn.AdviceDeletBefor(glngSys, p门诊医嘱下达, mlng病人ID, mlng挂号ID, lng医嘱ID, mint场合) = False Then
                    If err.Number = 0 Then Exit Sub
                End If
                Call zlPlugInErrH(err, "AdviceDeletBefor")
            End If
                        If Not CheckDelAdivceOfPathItem(Val(.TextMatrix(.Row, COL_ID))) Then Exit Sub
            If err.Number <> 0 Then err.Clear
            On Error GoTo 0
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            If InStr(",5,6,E,", .TextMatrix(.Row, COL_诊疗类别)) > 0 Then
                If Val(.TextMatrix(.Row, COL_处方审查状态)) = 1 Or Val(.TextMatrix(.Row, COL_处方审查状态)) = 2 Then
                    bln处方已审查 = True
                Else
                    '中药配方判断
                    If .TextMatrix(.Row, COL_操作类型) = "4" Then
                        For i = .Row To 1 Step -1
                            If .TextMatrix(i, COL_诊疗类别) = "7" And Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(.Row, COL_ID)) Then
                                If Val(.TextMatrix(i, COL_处方审查状态)) = 1 Or Val(.TextMatrix(i, COL_处方审查状态)) = 2 Then
                                    bln处方已审查 = True
                                End If
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If
            If bln处方已审查 Then
                arrSQL(UBound(arrSQL)) = "Zl_病人医嘱记录_处方审查删除(" & IIF(Val(.TextMatrix(.Row, COL_相关ID)) = 0, lng医嘱ID, Val(.TextMatrix(.Row, COL_相关ID))) & ")"
            Else
                arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_Delete(" & lng医嘱ID & ",1)"
            End If
            strDelIDs = strDelIDs & "," & lng医嘱ID
        End If
        strDelIDs = Mid(strDelIDs, 2)
    End With
    
    If blnRIS预约 Then
        Set rsTmp = GetDataRIS预约(strDelIDs)
        On Error Resume Next
        For i = 1 To rsTmp.RecordCount
            If 0 <> gobjRis.HISSchedulingEx(Val(rsTmp!ID & ""), Val(rsTmp!预约id & "")) Then '删除医嘱
                MsgBox "当前启用了影像信息系统接口，但由于影像信息系统接口(HISSchedulingEx)取消息预约未调用成功，请与系统管理员联系！", vbInformation, gstrSysName
            End If
            rsTmp.MoveNext
        Next
        err.Clear: On Error GoTo 0
    End If
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    If bln输血 = True Then
        If InitObjBlood(True) = True Then
            If gobjPublicBlood.AdviceOperation(p住院医嘱下达, lng医嘱ID, 2, False, strErr) = False Then
                gcnOracle.RollbackTrans: blnTrans = False
                MsgBox "血库公共部件调用失败，详细信息：" & strErr, vbInformation, gstrSysName
                Exit Sub
            End If
        Else
            gcnOracle.RollbackTrans: blnTrans = False
            MsgBox "血库公共部件创建失败，请检查！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    For i = 0 To UBound(arrSQL)
        zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
    Next
    gcnOracle.CommitTrans: blnTrans = False
    On Error GoTo 0
    
    With vsAdvice
        '界面上直接删除
        .Redraw = False
        
        '删除一并给药第一行时的显示处理
        If blnGroup And .Row + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(.Row, COL_相关ID)) = Val(.TextMatrix(.Row + 1, COL_相关ID)) Then
                If .TextMatrix(.Row, COL_开始时间) <> "" And .TextMatrix(.Row + 1, COL_开始时间) = "" Then
                    .TextMatrix(.Row + 1, COL_开始时间) = .TextMatrix(.Row, COL_开始时间)
                    .TextMatrix(.Row + 1, COL_频率) = .TextMatrix(.Row, COL_频率)
                    .TextMatrix(.Row + 1, COL_用法) = .TextMatrix(.Row, COL_用法)
                End If
            End If
        End If
                
        lngRow = .Row
        If blnBat Then
            For i = .Rows - 1 To 1 Step -1
                If .TextMatrix(i, COL_医嘱状态) = "1" And Val(.TextMatrix(i, COL_申请序号)) = lng申请序号 Then
                    .RemoveItem i
                End If
            Next
        Else
            .RemoveItem .Row
        End If
        
        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        If lngRow <= .Rows - 1 Then
            .Row = lngRow
        Else
            .Row = .Rows - 1
        End If
        If lng组ID <> 0 Then
            i = .FindRow(CStr(lng组ID), , COL_相关ID)
            If i <> -1 Then
                .TextMatrix(i, COL_并) = ""
                Call SetTag一并给药(i)
            End If
        End If
        Call .ShowCell(.Row, .Col)
        .Redraw = True
        
        Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col) '颜色及附表更新
        
        '调用删除后外挂接口
        On Error Resume Next
        arrDelID = Split(strDelIDs, ",")
        For i = 0 To UBound(arrDelID)
            If Val(arrDelID(i)) <> 0 Then
                If Not gobjPlugIn Is Nothing Then
                    Call gobjPlugIn.AdviceDeleted(glngSys, p门诊医嘱下达, mlng病人ID, mlng挂号ID, Val(arrDelID(i)), mint场合)
                    Call zlPlugInErrH(err, "AdviceDeleted")
                End If
            End If
        Next
        If err.Number <> 0 Then err.Clear
        On Error GoTo 0
        'PASS医嘱删除后自动调用审查功能
        If mblnPass Then
            Call gobjPass.zlPassAdviceSave(mobjPassMap, , , 4, strDelDrugIDs)
        End If
    End With
    Call ShowTotalMoney
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function CheckDelAdivceOfPathItem(ByVal lng医嘱ID As Long) As Boolean
'功能：检查医嘱对应的路径项目是否允许删除，如果是必须执行的项目所对应的医嘱，则需要弹出原因选择并更新变异原因，
'       添加过变异原因的不再添加
'返回：True-可以删除该医嘱，false-不可删除
'参数:lng医嘱ID
    Dim blnCancel As Boolean, blnMust As Boolean
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset, rsAdvice As ADODB.Recordset
    Dim strReason As String
    Dim vPoint As PointAPI
    Dim strTemp As String
    Dim arrTmp As Variant
    Dim arrSQL As Variant
    Dim i As Long

    '1.检查路径项目
    strSQL = "Select  c.Id as 执行Id, c.分类,c.变异原因,d.执行方式,c.天数,c.阶段ID,c.路径记录ID,c.项目ID " & _
             " From 病人门诊路径医嘱 B, 病人门诊路径执行 C, 门诊路径项目 D" & vbNewLine & _
             "Where b.病人医嘱Id=[1] And b.路径执行id = c.Id And d.Id = c.项目id And d.执行方式 in (1,2,4)"

    On Error GoTo errH

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查路径医嘱", lng医嘱ID)

    If rsTmp.RecordCount < 1 Then
        CheckDelAdivceOfPathItem = True
        Exit Function    '非 必须生成的路径医嘱
    End If
    '2.检查医嘱能否删除
    '该路径项目存在已校对但未作废的其他医嘱，提示并禁止删除    医嘱状态 ：3-已校对
    strSQL = "Select a.病人医嘱ID,b.医嘱状态 " & vbNewLine & _
             "From 病人门诊路径医嘱 A, 病人医嘱记录 B" & vbNewLine & _
             "Where a.路径执行id = [1] And a.病人医嘱id = b.Id  And b.医嘱状态>1 and b.医嘱状态<>4"

    Set rsAdvice = zlDatabase.OpenSQLRecord(strSQL, "检查路径医嘱", rsTmp!执行Id)

    If rsAdvice.RecordCount > 0 Then
        MsgBox "删除医嘱所在的路径项目中存在已发送但未作废的医嘱，请先作废该医嘱后再执行此操作。", vbInformation, gstrSysName
        CheckDelAdivceOfPathItem = False
        Exit Function
    End If
    

    
    '根据执行方式 决定是否有必要添加变异原因
    blnMust = CheckPathItemIsMust(Val(rsTmp!执行方式 & ""), Val("" & rsTmp!天数), Val("" & rsTmp!路径记录id), Val("" & rsTmp!阶段id), Val("" & rsTmp!项目ID), 1)
    If Not blnMust Then CheckDelAdivceOfPathItem = True: Exit Function
    
    '----------------------------
    '3.必须生成的项目填写变异原因
    For i = 1 To rsTmp.RecordCount
        If rsTmp!变异原因 & "" = "" Then
            strTemp = strTemp & rsTmp!执行Id & "," & rsTmp!分类 & ";"
        End If
        rsTmp.MoveNext
    Next
    
    If strTemp = "" Then
        CheckDelAdivceOfPathItem = True
        Exit Function
    Else
        strTemp = Mid(strTemp, 1, Len(strTemp) - 1)
    End If

    strSQL = "Select b.名称 as 分类,a.编码 as ID,a.编码,a.名称,a.简码 From 门诊变异常见原因 a,门诊变异常见原因 b" & _
             " Where a.性质=1 And a.末级=1 And a.上级=b.编码 And b.末级=0 " & _
             " Order by 分类,a.编码"
    vPoint = zlControl.GetCoordPos(vsAdvice.hwnd, vsAdvice.CellLeft, vsAdvice.CellTop)

    Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "门诊变异常见原因", True, , , True, True, True, _
                                      vPoint.X, vPoint.Y, vsAdvice.RowHeight(vsAdvice.Row), blnCancel, False, True)
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "系统没有初始门诊变异常见原因，请与系统管理员联系。", vbInformation, gstrSysName
        End If
        Exit Function
    Else
        strReason = rsTmp!ID
    End If

    If strReason <> "" Then
        arrSQL = Array()
        For i = 0 To UBound(Split(strTemp, ";"))
            arrTmp = Split(Split(strTemp, ";")(i), ",")
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_病人门诊路径生成_Update(" & arrTmp(0) & ",'" & arrTmp(1) & "',Null ,Null,Null,Null,'" & strReason & "')"
        Next
        '不添加事务处理，若变异原因添加失败，医嘱不会删除，再次删除时，会重新添加变异原因后才可删除。
        For i = LBound(arrSQL) To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
        Next
        CheckDelAdivceOfPathItem = True
    End If

    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub FuncAdviceRevoke()
'删除：当前医嘱作废(一组医嘱作废)
    Dim strSQL As String, lng医嘱ID As Long
       
    
    Dim strNO As String
    Dim lngType As Long
    Dim rsTmp As ADODB.Recordset
    
    If Not mblnCanRevoke Then Exit Sub
    
    With vsAdvice
        '检查是否可以作废
        If Val(.TextMatrix(.Row, COL_相关ID)) <> 0 Then
            lng医嘱ID = Val(.TextMatrix(.Row, COL_相关ID))
        Else
            lng医嘱ID = Val(.TextMatrix(.Row, COL_ID))
        End If
        
        If RowIn一并给药(.Row, 0, 0) Then
            lngType = 1
        End If
        
        If .TextMatrix(.Row, COL_诊疗类别) = "E" Then
            If RowIs检验行(.Row) Then
                lngType = 2
            End If
        End If
        
        If .TextMatrix(.Row, COL_诊疗类别) = "5" Or .TextMatrix(.Row, COL_诊疗类别) = "6" Then
            strNO = .Cell(flexcpData, .Row, COL_处方号)
        End If
        
        If RevokeOutAdvice(mlng病人ID, mlng挂号ID, mstr挂号单, mstr姓名, mstr门诊号, mlng挂号科室ID, lng医嘱ID, Val(.TextMatrix(.Row, COL_医嘱状态)), .TextMatrix(.Row, COL_诊疗类别), .TextMatrix(.Row, COL_操作类型), Val(.TextMatrix(.Row, COL_审核状态)), _
            .Cell(flexcpData, .Row, col_发送时间), Val(.TextMatrix(.Row, COL_签名否)), lngType, .TextMatrix(.Row, col_医嘱内容), mblnMoved, mclsMipModule, mint场合) = False Then Exit Sub
        
    End With
    
    Call LoadAdvice '刷新界面
    Call ShowTotalMoney
    
    'PASS医嘱作废后自动调用审查功能
    If mblnPass Then
        Call gobjPass.zlPassAdviceSave(mobjPassMap, , , 3)
    End If
    
    '药品医嘱用废后判断是不是要重打
    If strNO <> "" Then
        strSQL = "Select Distinct D.编号,D.名称,D.说明,B.NO,B.记录性质" & _
            " From 病人医嘱记录 A,病人医嘱发送 B,病历单据应用 C,病历文件列表 D" & _
            " Where C.诊疗项目ID = A.诊疗项目ID And a.ID=b.医嘱ID " & _
            " And C.应用场合=1 And C.病历文件ID=D.ID And D.种类=7 And b.NO=[1] and a.诊疗类别 in ('5','6') and a.挂号单||''=[2]" & _
            " Order by D.编号"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, strNO, mstr挂号单)
        If Not rsTmp.EOF Then
            If MsgBox("您作废的药品处方签已经打印，是否重打？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                '调用打印
                SwitchPrintSet glngSys & "\" & p门诊医嘱下达
                If mobjReport.ReportPrintSet(gcnOracle, glngSys, "ZLCISBILL" & Format(rsTmp!编号, "00000") & "-1", mfrmParent) Then
                    mstrBillPrint = "ZLCISBILL" & Format(rsTmp!编号, "00000") & "-1" & "," & rsTmp!NO & "," & rsTmp!记录性质
                    Call mobjReport.ReportOpen(gcnOracle, glngSys, "ZLCISBILL" & Format(rsTmp!编号, "00000") & "-1", mfrmParent, "NO=" & rsTmp!NO, "性质=" & rsTmp!记录性质, 2)
                    mstrBillPrint = ""
                End If
                SwitchPrintSet glngSys & "\" & p门诊医嘱下达, True
            End If
        End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncAdviceRevoke停用()
'删除：当前医嘱作废(一组医嘱作废)
    Dim strSQL As String, lng医嘱ID As Long
    Dim lng证书ID As Long, lng签名id As Long
    Dim strSign As String, intRule As Integer
    Dim strSource As String, strIDs As String, blnDo As Boolean
    Dim strTimeStamp As String, blnTran As Boolean, strErr As String, strTimeStampCode As String
    
    If Not mblnCanRevoke Then Exit Sub
    
    With vsAdvice
        '检查是否可以作废
        If Val(.TextMatrix(.Row, COL_相关ID)) <> 0 Then
            lng医嘱ID = Val(.TextMatrix(.Row, COL_相关ID))
        Else
            lng医嘱ID = Val(.TextMatrix(.Row, COL_ID))
        End If
        
        If lng医嘱ID = 0 Then
            MsgBox "该病人没有医嘱可以作废。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If Val(.TextMatrix(.Row, COL_医嘱状态)) <> 8 Then
            MsgBox "当前选择的医嘱尚未发送或已经作废。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '92129:医嘱已被输血科接收则不能进行作废
        If .TextMatrix(.Row, COL_诊疗类别) = "K" And gbln血库系统 And InStr(1, ",2,5,6,", "," & Val(.TextMatrix(.Row, COL_审核状态)) & ",") <> 0 Then
            MsgBox "本次作废的输血医嘱" & IIF(Val(.TextMatrix(.Row, COL_审核状态)) = 2, "已经完成配血", "处于正在配血阶段") & "，不能直接作废医嘱，若要作废请与输血科联系。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '已有费用转出不允许作废
        If zlDatabase.DateMoved(.Cell(flexcpData, .Row, col_发送时间)) Then
            If MovedBySend(lng医嘱ID, 0, 1) Then
                MsgBox "该医嘱的费用已经全部或部份转出到后备数据库，不允许操作。" & vbCrLf & _
                    "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        '电子签名检查和提示
        If Val(.TextMatrix(.Row, COL_签名否)) = "1" Then
            If gobjESign Is Nothing Then
                If gintCA = 0 Then
                    MsgBox "作废已签名医嘱时需要再次签名，但系统没有设置签名认证中心，不能作废。", vbInformation, gstrSysName
                Else
                    MsgBox "作废已签名医嘱时需要再次签名，但电子签名部件未能正确安装，不能作废。", vbInformation, gstrSysName
                End If
                Exit Sub
            End If
            If gobjESign.CertificateStoped(UserInfo.姓名) = False Then strSign = vbCrLf & vbCrLf & "提示：该医嘱已经签名，作废时你需要再次签名。"
        End If
        
        '检查作废医嘱对应的费用结帐情况
        If Not CheckAdviceBalanceRevoke(lng医嘱ID) Then Exit Sub
        
        '已审核记帐费用检查
        If InStr(GetInsidePrivs(p门诊医嘱下达), "作废已审核记帐医嘱") = 0 Then
            If Not CheckAdviceBillingRevoke(lng医嘱ID) Then
                MsgBox "要作废医嘱的对应记帐划价费用已经审核，不能作废。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        If RowIn一并给药(.Row, 0, 0) Then
            If MsgBox("该组一并给药的医嘱将会一起作废，确实要作废吗？" & strSign, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Else
            If MsgBox("确实要作废医嘱""" & .TextMatrix(.Row, col_医嘱内容) & """吗？" & strSign, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        
        strSQL = "ZL_病人医嘱记录_作废(" & lng医嘱ID & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
        
        '作废时进行电子签名
        If strSign <> "" Then
            If gobjESign.CertificateStoped(UserInfo.姓名) = False Then
                '获取签名医嘱源文
                strIDs = lng医嘱ID
                intRule = ReadAdviceSignSource(4, mlng病人ID, mstr挂号单, strIDs, 0, mblnMoved, strSource)
                If intRule = 0 Then Exit Sub
                If strSource = "" Then
                    MsgBox "不能读取需要作废的已签名医嘱源文内容。", vbInformation, gstrSysName
                    Exit Sub
                End If
                
                strSign = gobjESign.Signature(strSource, gstrDBUser, lng证书ID, strTimeStamp, Nothing, strTimeStampCode)
                If strSign <> "" Then
                    If strTimeStamp <> "" Then
                        strTimeStamp = "To_Date('" & strTimeStamp & "','YYYY-MM-DD HH24:MI:SS')"
                    Else
                        strTimeStamp = "NULL"
                    End If
                    lng签名id = zlDatabase.GetNextID("医嘱签名记录")
                    strSign = "zl_医嘱签名记录_Insert(" & lng签名id & ",4," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng证书ID & ",'" & strIDs & "'," & strTimeStamp & ",'" & UserInfo.姓名 & "','" & strTimeStampCode & "')"
                Else
                    Exit Sub
                End If
            End If
        End If
    End With
    
    'RIS检查
    If HaveRIS Then
        If 1 <> gobjRis.ReqInteractive(5, "AppNO", lng医嘱ID) Then
            Exit Sub
        End If
    End If
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTran = True
    zlDatabase.ExecuteProcedure strSQL, Me.Name
    If strSign <> "" Then
        zlDatabase.ExecuteProcedure strSign, Me.Name
    End If
    gcnOracle.CommitTrans: blnTran = False
    On Error GoTo 0
    If Not (mclsMipModule Is Nothing) Then
        If mclsMipModule.IsConnect Then
            Call ZLHIS_CIS_024(mclsMipModule, mlng病人ID, mstr姓名, , mstr门诊号, 1, mlng挂号ID, mlng挂号科室ID, "", lng医嘱ID, _
                vsAdvice.TextMatrix(vsAdvice.Row, COL_诊疗类别), vsAdvice.TextMatrix(vsAdvice.Row, COL_操作类型))
        End If
    End If
    '调用作废后外挂接口
    Call CreatePlugInOK(p门诊医嘱下达, mint场合)
    On Error Resume Next
    If Not gobjPlugIn Is Nothing Then
        Call gobjPlugIn.AdviceRevoked(glngSys, p门诊医嘱下达, mlng病人ID, mlng挂号ID, lng医嘱ID, mint场合)
        Call zlPlugInErrH(err, "AdviceRevoked")
    End If
    If err.Number <> 0 Then err.Clear
    On Error GoTo 0
    Call InitObjLis(p门诊医生站)
    '调用LIS作废申请单
    If Not gobjLIS Is Nothing Then
        If gobjLIS.DelLisApplicationForm(CStr(lng医嘱ID), strErr) = False Then
            MsgBox "删除检验申请失败：" & strErr, vbInformation, gstrSysName
        End If
    End If
    '调用数据交换平台，向LIS,PACS取消申请单
    If gobjExchange Is Nothing Then
        On Error Resume Next
        Set gobjExchange = CreateObject("zlExchange.clsExchange")
        If Not gobjExchange Is Nothing Then Call gobjExchange.Init(gcnOracle)
        err.Clear: On Error GoTo 0
    End If
    
    If Not gobjExchange Is Nothing Then
        With vsAdvice
            If .TextMatrix(.Row, COL_诊疗类别) = "D" Then
                blnDo = True
            ElseIf .TextMatrix(.Row, COL_诊疗类别) = "E" Then
                blnDo = RowIs检验行(.Row)
            End If
            If blnDo Then
                Call gobjExchange.SendMsg(IIF(.TextMatrix(.Row, COL_诊疗类别) = "D", 2, 1), "病人ID::" & mlng病人ID & "||主页ID::0||医嘱ID::" & lng医嘱ID & "||操作类型::0")
            End If
        End With
    End If
    
    Call LoadAdvice '刷新界面
    Call ShowTotalMoney
    
    'PASS医嘱作废后自动调用审查功能
    If mblnPass Then
        Call gobjPass.zlPassAdviceSave(mobjPassMap, , , 3)
    End If
    Exit Sub
errH:
    If blnTran Then gcnOracle.RollbackTrans: blnTran = False
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncAdviceModi()
'功能：修改当前医嘱
    Dim lng医嘱ID As Long
    
    If Not CheckWindow Then Exit Sub
        '检查挂号病人是否超期
    If Not FuncTimeLimitCheck Then Exit Sub
    
    With vsAdvice
        lng医嘱ID = Val(.TextMatrix(.Row, COL_ID))
        If lng医嘱ID = 0 Then Exit Sub
        
        '医技下达的医嘱
        If Val(.TextMatrix(.Row, COL_前提ID)) <> mlng前提ID Then
            MsgBox "不能修改该医嘱,该医嘱是根据其他主医嘱产生的。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '已校对或已废止
        If Val(.TextMatrix(.Row, COL_医嘱状态)) <> 1 Then
            MsgBox "当前选择的医嘱已经发送或作废，不能修改。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '已签名的医嘱不能修改
        If Val(.TextMatrix(.Row, COL_签名否)) = 1 Then
            MsgBox "当前选择的医嘱已经签名，不能修改。请先取消签名。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        Set mfrmEdit = frmOutAdviceEdit
        Call frmOutAdviceEdit.ShowMe(mfrmParent, mint场合, mMainPrivs, mlng病人ID, mstr挂号单, mlng前提ID, _
            Val(.TextMatrix(.Row, COL_婴儿ID)), lng医嘱ID, , mlng界面科室ID, mstr前提IDs, mclsMipModule, mlng挂号科室ID, mblnMoved, mint就诊类型)
    End With
End Sub

Private Sub FuncAdviceTest()
'功能：填写皮试结果
    Dim strSQL As String, str结果 As String
    Dim int结果 As Integer, strLabel As String
    Dim rsTmp As New ADODB.Recordset
    Dim blnTran As Boolean
    Dim dateInput As Date
    Dim strSelect As String, i As Long
    Dim strSelectInput As String
    Dim strTextInput As String
    
    If mlng病人ID = 0 Then Exit Sub
    If Not mblnEditable Then Exit Sub
    
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) = 0 Then Exit Sub
    If Not (vsAdvice.TextMatrix(vsAdvice.Row, COL_诊疗类别) = "E" And vsAdvice.TextMatrix(vsAdvice.Row, COL_操作类型) = "1") Then
        MsgBox "当前医嘱内容不是过敏试验项目。", vbInformation, gstrSysName
        Exit Sub
    End If
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_前提ID)) <> 0 Then
        MsgBox "你不能给该过敏试验填写结果。", vbInformation, gstrSysName
        Exit Sub
    End If
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱状态)) = 4 Then
        MsgBox "该过敏试验医嘱已经作废，不能填写结果。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    On Error GoTo errH
    
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱状态)) = 1 Then
        If vsAdvice.TextMatrix(vsAdvice.Row, COL_皮试) = "免试" Then
            If MsgBox("该过敏试验医嘱已经标记为免试，要清除免试标记吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            strLabel = ""
        Else
            If MsgBox("该过敏试验医嘱尚未发送，不允许填写过敏试验结果。" & vbCrLf & vbCrLf & _
                "但可以标记为免试，同时该医嘱将不会发送。要标记为免试吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            strLabel = "免试"
        End If
        int结果 = -1 '特殊区分出来
        strSQL = "ZL_病人医嘱记录_皮试(" & Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) & ",'" & strLabel & "',NULL)"
    Else
        If vsAdvice.TextMatrix(vsAdvice.Row, COL_皮试) <> "" Then
            If MsgBox("该过敏试验医嘱已经填写了结果，要重新填写吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            '检查对应的医嘱是否已经发送
            If mbln皮试限制 Then
                If AdviceSended(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)), CDate(vsAdvice.TextMatrix(vsAdvice.Row, COL_开嘱时间))) Then
                    MsgBox "该皮试对应的药品已经发送，不能再更改皮试结果。", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        End If
        
        strSQL = "Select Nvl(B.标本部位,'阳性(+);阴性(-)') as 标本部位 From 病人医嘱记录 A,诊疗项目目录 B Where A.诊疗项目ID=B.ID And A.ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)))
        '阳性
        For i = 0 To UBound(Split(Split(rsTmp!标本部位 & "", ";")(0), ","))
            strSelect = strSelect & "," & Split(Split(rsTmp!标本部位 & "", ";")(0), ",")(i) & "|0"
        Next
        '阴性
        For i = 0 To UBound(Split(Split(rsTmp!标本部位 & "", ";")(1), ","))
            strSelect = strSelect & "," & Split(Split(rsTmp!标本部位 & "", ";")(1), ",")(i) & "|0|2"
        Next
        strSelect = Mid(strSelect, 2)
        str结果 = zlCommFun.ShowMsgBox("皮试结果", vsAdvice.TextMatrix(vsAdvice.Row, col_医嘱内容) & "：^^请根据过敏试验结果选择相应的按钮操作。", _
            "确定(&O),?取消(&C)", Me, vbQuestion, "皮试时间", dateInput, "yyyy-MM-dd HH:mm", "皮试结果(&P):" & strSelect, strSelectInput, _
            "过敏反应(&F)", 50, strTextInput, , True)
        If str结果 = "" Then Exit Sub
        If strSelectInput = "" Then Exit Sub
        If Format(vsAdvice.TextMatrix(vsAdvice.Row, COL_开始时间), "yyyy-MM-dd HH:mm") > dateInput Then
            MsgBox "皮试时间不能在医嘱生效时间以前，请重新录入。", vbInformation, gstrSysName
            Exit Sub
        End If
        Call GetTestLabel(rsTmp!标本部位, strSelectInput, strLabel, int结果)
        strSQL = "ZL_病人医嘱记录_皮试(" & Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) & ",'" & strLabel & "'," & int结果 & _
                ",'',to_date('" & dateInput & "','YYYY-MM-DD HH24:MI:SS')" & ",'" & strTextInput & "')"
    End If
        
    zlDatabase.ExecuteProcedure strSQL, Me.Name
    
    vsAdvice.TextMatrix(vsAdvice.Row, COL_皮试) = strLabel
    If mvarCond.显示模式 = 0 Then
        '如果是简洁模式，则设置药品后的皮试结果。
        If InStr(vsAdvice.TextMatrix(vsAdvice.Row, col_内容), "(+)") > 0 Or InStr(vsAdvice.TextMatrix(vsAdvice.Row, col_内容), "(-)") > 0 Then
            vsAdvice.TextMatrix(vsAdvice.Row, col_内容) = Replace(vsAdvice.TextMatrix(vsAdvice.Row, col_内容), "(+)", strLabel)
            vsAdvice.TextMatrix(vsAdvice.Row, col_内容) = Replace(vsAdvice.TextMatrix(vsAdvice.Row, col_内容), "(-)", strLabel)
        Else
            vsAdvice.TextMatrix(vsAdvice.Row, col_内容) = vsAdvice.TextMatrix(vsAdvice.Row, col_内容) & strLabel
        End If
    End If
    
    If int结果 = 1 Then
        vsAdvice.Cell(flexcpForeColor, vsAdvice.Row, COL_皮试) = vbRed
    ElseIf int结果 = 0 Then
        vsAdvice.Cell(flexcpForeColor, vsAdvice.Row, COL_皮试) = vbBlue
    Else
        vsAdvice.Cell(flexcpForeColor, vsAdvice.Row, COL_皮试) = vsAdvice.Cell(flexcpForeColor, vsAdvice.Row, col_医嘱内容)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function AdviceSended(ByVal lng医嘱ID As Long, Optional dat开嘱时间 As Date) As Boolean
'功能：判断皮试对应的医嘱是否已经发送(只判断皮试医嘱开始时间之后的医嘱77377)
'参数：lng医嘱ID=皮试医嘱的ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '已作废的不管
    strSQL = "Select 诊疗项目ID From 病人医嘱记录 Where ID=[3]"
    strSQL = "Select A.ID From 病人医嘱记录 A,诊疗用法用量 B" & _
        " Where Rownum<2 And A.诊疗类别 IN('5','6') And A.医嘱状态=8" & _
        " And A.诊疗项目ID=B.项目ID And B.性质=0 And B.用法ID=(" & strSQL & ")" & _
        " And A.病人ID+0=[1] And A.挂号单=[2] And A.开嘱时间>=[4]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng病人ID, mstr挂号单, lng医嘱ID, dat开嘱时间)
    AdviceSended = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FuncAdviceSend(blnAuto As Boolean)
'功能：发送病人医嘱(可以设置计价项目)

    If mlng病人ID = 0 Then Exit Sub
    If Not mblnEditable Then Exit Sub

    If mfrmSend Is Nothing Then Set mfrmSend = New frmOutAdviceSend
    If mfrmSend.ShowMe(mfrmParent, mMainPrivs, mlng病人ID, mstr挂号单, mstr前提IDs, blnAuto, mlng界面科室ID, mint场合, mclsMipModule) Then
        Call LoadAdvice
        Call ShowTotalMoney
    End If
End Sub


Private Sub FuncToolScheme()
'功能：调用成套方案维护
    On Error Resume Next
    If gobjCISBase Is Nothing Then
        Set gobjCISBase = CreateObject("zl9CISBase.clsCISBase")
        If gobjCISBase Is Nothing Then
            MsgBox "诊疗基础部件没有正确安装，功能无法执行。", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    err.Clear: On Error GoTo 0
    
    Call gobjCISBase.CallClinicScheme(mfrmParent, gcnOracle, glngSys, gstrDBUser, IIF(mint场合 = 2, 3, 1))
End Sub

Private Sub FuncEPRReport(ByVal lngMenu As Long)
'功能：查阅、打印、预览报告
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strBill As String, strTmp As String
    Dim strNO As String, int性质 As Long, i As Long
    Dim lng医嘱ID As Long, lngReportID As Long, blnPrint As Boolean, bln打印 As Boolean
    Dim bln检验行 As Boolean, bln配方行 As Boolean, arrRPTPar(19) As String, strFlagString As String
    Dim str检查报告ID As String
    Dim lngViewMode As Long ' 1-病历格式，6-报表格式
    Dim blnLis接口 As Boolean
    
    On Error GoTo errH
    If mblnMoved Then
        MsgBox "当前病人报告数据已转出，请统一到电子病案查阅模块中进行查看。", vbInformation, gstrSysName
        Exit Sub
    End If
    '调用数据交换平台，向LIS,PACS查阅报告
    If lngMenu = conMenu_Edit_Compend * 10# + 1 Or lngMenu = conMenu_Edit_Compend * 10# + 6 Or lngMenu = conMenu_Edit_Compend Then
        If lngMenu = conMenu_Edit_Compend * 10# + 1 Then
            lngViewMode = 1
        ElseIf lngMenu = conMenu_Edit_Compend * 10# + 6 Then
            lngViewMode = 6
        Else
            If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_报告项)) = 1 Then
                lngViewMode = 1
            Else
                lngViewMode = 6
            End If
        End If
        
        If gobjExchange Is Nothing Then
            On Error Resume Next
            Set gobjExchange = CreateObject("zlExchange.clsExchange")
            If Not gobjExchange Is Nothing Then Call gobjExchange.Init(gcnOracle)
            err.Clear: On Error GoTo 0
        End If
        If Not gobjExchange Is Nothing Then
            With vsAdvice
                '检验行存的是采集方法（诊疗类别为E），所以只判断检查行
                Call gobjExchange.SendMsg(IIF(.TextMatrix(.Row, COL_诊疗类别) = "D", 4, 3), "医嘱ID::" & .TextMatrix(.Row, COL_ID) & "||操作员姓名::" & UserInfo.姓名 & "||操作员缺省部门::" & UserInfo.部门名)
            End With
            Exit Sub
        End If
    End If
    
    lngReportID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_报告ID))
    lng医嘱ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
    str检查报告ID = vsAdvice.TextMatrix(vsAdvice.Row, COL_检查报告ID)
    
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_LIS报告ID)) <> 0 Then
        Call FuncLisRptFileView(mfrmParent, lng医嘱ID)   '三方的LIS文件报告
        If lngReportID = 0 And str检查报告ID = "" Then Exit Sub
    End If
    
    '先判断是否可以继续操作
    Select Case CheckEPRReport(lng医嘱ID, lngReportID, , , mblnMoved)
    Case 0
        MsgBox "该医嘱的报告没有书写！", vbInformation, gstrSysName
        Exit Sub
    Case 2
        strTmp = ""
        '紧急医嘱或者标记绿色通的项目可以查看未完成的报告
        If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_标志)) = 1 Then
            strTmp = "允许查看未完成报告"
        Else
            If vsAdvice.TextMatrix(vsAdvice.Row, COL_诊疗类别) = "D" Then
                strSQL = "select 1 from 影像检查记录 a where a.绿色通道=1 and a.医嘱id=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
                If Not rsTmp.EOF Then
                    strTmp = "允许查看未完成报告"
                End If
            End If
        End If
        If InStr(GetInsidePrivs(p门诊医嘱下达), "查阅未完成报告") > 0 Or strTmp <> "" Then
            MsgBox "注意：该医嘱的报告还没有正式签名！", vbInformation, gstrSysName
        Else
            MsgBox "该医嘱的报告还没有完成(没有正式签名或完成执行)，你没有权限操作！", vbInformation, gstrSysName
            Exit Sub
        End If
    End Select
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_RIS报告ID)) <> 0 Then
        If HaveRIS Then 'RIS报告兼容
            i = gobjRis.ShowViewReport(mfrmParent.hwnd, lng医嘱ID, InStr(GetInsidePrivs(p门诊医嘱下达), ";报告打印;") > 0)
            If i = 0 Then Exit Sub
        End If
    End If
    
    '执行操作
    '新版PACS报告，直接强制使用新版PACS报告编辑器
    If str检查报告ID <> "" Then
        Call CreateObjectPacs(mobjPublicPACS)
        Call mobjPublicPACS.zlDocShowReport(lng医嘱ID, , mblnAutoRead, mfrmParent)
    Else
        bln打印 = InStr(GetInsidePrivs(p门诊医嘱下达), ";报告打印;") > 0 And mblnEditable
        
        '检验项目应该调用LIS接口
        If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_操作类型)) = 6 And vsAdvice.TextMatrix(vsAdvice.Row, COL_诊疗类别) = "E" Then
            Call InitObjLis(p门诊医生站)
            If Not gobjLIS Is Nothing Then
                blnLis接口 = True
            End If
        End If
        
        If lngMenu = conMenu_Edit_Compend * 10# + 1 Or (lngMenu = conMenu_Edit_Compend And lngViewMode = 1) Then
            '查阅报告
            If blnLis接口 Then
                strTmp = ""
                Call gobjLIS.LisSingleReprotBrowse(mfrmParent, lng医嘱ID, 0, strTmp)
                If strTmp <> "" Then
                    MsgBox strTmp, vbInformation, gstrSysName
                    Exit Sub
                End If
            Else
                RaiseEvent ViewEPRReport(lngReportID, bln打印)
            End If
        Else
            If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_报告项)) = 1 And lngMenu <> conMenu_Edit_Compend * 10# + 6 And Not (lngMenu = conMenu_Edit_Compend And lngViewMode = 6) Then
                '按编辑格式打印、预览报告
                If blnLis接口 Then
                    strTmp = ""
                    Call gobjLIS.LisSingleReprotBrowse(mfrmParent, lng医嘱ID, 0, strTmp)
                    If strTmp <> "" Then
                        MsgBox strTmp, vbInformation, gstrSysName
                        Exit Sub
                    End If
                Else
                    RaiseEvent PrintEPRReport(lngReportID, lngMenu = conMenu_Edit_Compend * 10# + 3)
                End If
            Else
                bln检验行 = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_操作类型)) = 6 And vsAdvice.TextMatrix(vsAdvice.Row, COL_诊疗类别) = "E"
                If Not bln检验行 Then bln配方行 = RowIs配方行(vsAdvice.Row)
                    
                If bln检验行 Then
                    If blnLis接口 Then
                        strTmp = ""
                        Call gobjLIS.LisSingleReprotBrowse(mfrmParent, lng医嘱ID, 1, strTmp)
                        If strTmp <> "" Then
                            MsgBox strTmp, vbInformation, gstrSysName
                            Exit Sub
                        End If
                    Else
                        '调用LisWork打印检验报告
                        blnPrint = IIF(lngMenu = conMenu_Edit_Compend * 10# + 2, True, False)
                        If Not Open_LIS_Report(Me, lng医嘱ID, mlng病人ID, mblnMoved, blnPrint, Not bln打印) Then
                            MsgBox "该医嘱的报告为新版LIS产生，请使用(浏览检验结果)功能！", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                Else
                    '读取最近一次发送的NO,性质
                    If bln检验行 Or bln配方行 Then
                        '检验医嘱应以检验项目的NO为准
                        strSQL = "Select ID From 病人医嘱记录 Where 相关ID=[1] And Rownum=1"
                        strSQL = "Select 医嘱ID,NO,记录性质 From 病人医嘱发送 Where 医嘱ID=(" & strSQL & ") Order by 发送号 Desc"
                    Else
                        strSQL = "Select 医嘱ID,NO,记录性质 From 病人医嘱发送 Where 医嘱ID=[1] Order by 发送号 Desc"
                    End If
                                        If mblnMoved Then
                        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
                        strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
                    End If
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
                    If Not rsTmp.EOF Then
                        strNO = NVL(rsTmp!NO): int性质 = NVL(rsTmp!记录性质, 0)
                    End If
                    
                    '按报表格式打印、预览报告
                    strSQL = "Select 编号 From 病历文件列表 Where ID=[1]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_文件ID)))
                    If Not rsTmp.EOF Then
                        strBill = "ZLCISBILL" & Format(rsTmp!编号, "00000") & "-2"
                    End If
                    
                    If lngMenu = conMenu_Edit_Compend * 10# + 2 Then
                        If Not ReportPrintSet(gcnOracle, glngSys, strBill, Me) Then Exit Sub
                    End If
                    
                    
                    If Not bln检验行 And Not bln配方行 Then
                        strFlagString = GetRPTPicture(mblnMoved, lngReportID, strBill, arrRPTPar)
                    End If
                    
                    If lngMenu <> conMenu_Edit_Compend * 10# + 2 And Not bln打印 Then
                        strTmp = "DisabledPrint=1"
                    Else
                        strTmp = "DisabledPrint=0"
                    End If
                    
                    '医嘱ID为采集方式的ID，即检验的相关ID
                    Call ReportOpen(gcnOracle, glngSys, strBill, Me, "NO=" & strNO, "性质=" & int性质, _
                        "医嘱ID=" & lng医嘱ID, _
                        strFlagString, _
                        arrRPTPar(0), arrRPTPar(1), arrRPTPar(2), arrRPTPar(3), arrRPTPar(4), arrRPTPar(5), _
                        arrRPTPar(6), arrRPTPar(7), arrRPTPar(8), arrRPTPar(9), arrRPTPar(10), arrRPTPar(11), _
                        arrRPTPar(12), arrRPTPar(13), arrRPTPar(14), arrRPTPar(15), arrRPTPar(16), arrRPTPar(17), _
                        arrRPTPar(18), arrRPTPar(19), strTmp, _
                        IIF(lngMenu = conMenu_Edit_Compend * 10# + 2, 2, 1))
                End If
            End If
        End If
        
        '自动标记为已查阅：护士查阅不算
        If mblnAutoRead And mint场合 <> 1 Then Call FuncExecReportRead(True, True)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncExecReportRead(ByVal blnRead As Boolean, Optional ByVal blnAuto As Boolean)
'功能：设置当前报告为已查阅，或者取消当前报告的查阅状态
'参数：blnRead=已阅或者取消阅读状态
'      blnAuto=设置为已阅时，是否自动设置在调用
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strAdvice As String
    Dim strTmp As String
    Dim strErr As String
    
    With vsAdvice
        If Val(.TextMatrix(.Row, COL_ID)) = 0 Then Exit Sub
        '新版PACS编辑器报告，直接调用接口标记已阅
        If .TextMatrix(.Row, COL_检查报告ID) = "" Then
            If Val(.TextMatrix(.Row, COL_报告ID)) = 0 Then Exit Sub
            
            If blnRead Then
                If Not blnAuto Then
                    If Val(.Cell(flexcpData, .Row, COL_查阅状态)) = 1 Then Exit Sub '自动标记时不计次数
                    If MsgBox("请确认该项目报告您已经仔细阅读了吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                End If
                
                strSQL = "Zl_报告查阅记录_Insert(" & Val(.TextMatrix(.Row, COL_ID)) & "," & Val(.TextMatrix(.Row, COL_报告ID)) & ")"
            Else
                If MsgBox("你确实要取消该报告的查阅状态吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                strSQL = "Zl_报告查阅记录_Cancel(" & Val(.TextMatrix(.Row, COL_ID)) & "," & Val(.TextMatrix(.Row, COL_报告ID)) & ",'" & UserInfo.姓名 & "')"
            End If
            Call InitObjLis(p门诊医生站)
            On Error GoTo errH
            Call zlDatabase.ExecuteProcedure(strSQL, "FuncExecReportRead")
            If Not gobjLIS Is Nothing Then
                '检验调用标记接口
                strTmp = "Select ID From 病人医嘱记录 Where ID=[1] Or 相关ID=[1] order by 序号"
                Set rsTmp = zlDatabase.OpenSQLRecord(strTmp, "FuncExecReportRead", Val(.TextMatrix(.Row, COL_ID)))
                Do While Not rsTmp.EOF
                    strAdvice = strAdvice & "," & rsTmp!ID
                    rsTmp.MoveNext
                Loop
                If .TextMatrix(.Row, COL_诊疗类别) = "E" And .TextMatrix(.Row, COL_操作类型) = "6" Then
                    gobjLIS.WriteAdvicesLookState Mid(strAdvice, 2), IIF(blnRead, 1, 0)
                End If
            End If
            On Error GoTo 0
        Else
            Call CreateObjectPacs(mobjPublicPACS)
            Call mobjPublicPACS.zlDocViewStateUpdate(blnRead, Val(.TextMatrix(.Row, COL_ID)))
        End If
        '设置界面状态
        If blnRead Then
            .Cell(flexcpData, .Row, COL_查阅状态) = 1 '我已查阅
        Else
            On Error GoTo errH
            strSQL = "Select Count(*) as 次数 From 报告查阅记录 Where 医嘱ID=[1] And 取消时间 Is Null"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "FuncExecReportRead", Val(.TextMatrix(.Row, COL_ID)))
            If NVL(rsTmp!次数, 0) = 0 Then
                .Cell(flexcpData, .Row, COL_查阅状态) = 0 '我未查阅
            End If
        End If
        Call SetAdviceReportIcon(.Row)
        .TextMatrix(.Row, COL_查阅状态) = "查阅"
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub tbcAppend_GotFocus()
    If vsAppend.Visible And vsAppend.Enabled Then
        vsAppend.SetFocus
    ElseIf rtfAppend.Visible And rtfAppend.Enabled Then
        rtfAppend.SetFocus
    End If
End Sub

Private Sub tbcAppend_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim blnDo As Boolean
    
    If Item.Tag = "" Then Exit Sub
    
    If Visible Then
        If Decode(vsAppend.Tag, "计价", True, "发送", True, "签名", True, False) Then
            Call SaveFlexState(vsAppend, App.ProductName & "\" & Me.Name)
        End If
    End If
    vsAppend.Tag = Item.Tag '用于公共函数区分个性化
    
    If Item.Tag = "计价" Then
        Call InitPriceTable
    ElseIf Item.Tag = "发送" Then
        Call InitSendTable
    ElseIf Item.Tag = "签名" Then
        Call InitSignTable
    ElseIf Item.Tag = "附项" Then
        'NoneCode
    ElseIf Item.Tag = "安排" Then
        'NoneCode
    End If
    
    If Visible Then
        If Decode(Item.Tag, "计价", True, "发送", True, "签名", True, False) Then
            Call RestoreFlexState(vsAppend, App.ProductName & "\" & Me.Name)
        End If
    End If
    
    Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    
    If Visible Then
        If vsAppend.Visible And vsAppend.Enabled Then
            vsAppend.SetFocus
        ElseIf rtfAppend.Visible And rtfAppend.Enabled Then
            rtfAppend.SetFocus
        End If
    End If
End Sub

Private Sub vsAdvice_Click()
'功能：查阅报告
    Dim lngMouseRow As Long, lngMouseCol As Long
    
    If mblnTag Then Exit Sub '如果已点击过查看报告，在显示报告前不允许在点击查看
    'PASS
    If mblnPass And Me.Visible Then
        Call gobjPass.zlPassAdviceMainPoint(mobjPassMap, 1)
    End If
    
    On Error GoTo errH
    
    If mvarCond.过滤模式 <> 3 Then Exit Sub
    With vsAdvice
        lngMouseRow = .MouseRow
        lngMouseCol = .MouseCol
        
        If lngMouseRow > -1 And lngMouseCol > -1 Then
            If .Cell(flexcpFontUnderline, lngMouseRow, lngMouseCol) = True Then
                .Redraw = False
                mblnTag = True
                Call FuncEPRReport(conMenu_Edit_Compend)
                .Cell(flexcpForeColor, lngMouseRow, COL_查阅状态) = &H80& '暗红
                mblnTag = False
                .Redraw = True
            End If
        End If
    End With
    Exit Sub
errH:
    mblnTag = False
End Sub

Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim blnExist As Boolean, blnSel As Boolean, bln输血 As Boolean
    Dim varDraw As RedrawSettings, intIdx As Integer
    
    If NewRow = OldRow Then Exit Sub
    If fraMore.Visible = True Then fraMore.BackColor = vsAdvice.BackColorSel
    
    If vsAdvice.Col >= vsAdvice.FixedCols Then
        vsAdvice.ForeColorSel = vsAdvice.Cell(flexcpForeColor, NewRow, COL_开始时间)
    End If
    
     'PASS
    If mblnPass And Me.Visible Then
        If NewRow <> OldRow Then
            Call gobjPass.zlPassSetDrug(mobjPassMap)
        End If
    End If
    
    Call LoadBillList '显示可打印的诊疗单据
    If vsAdvice.Redraw <> flexRDNone Then
        If Val(vsAdvice.TextMatrix(NewRow, COL_ID)) <> 0 Then
            '显示报告是否我已阅读
            If Val(vsAdvice.TextMatrix(NewRow, COL_报告ID)) <> 0 Or vsAdvice.TextMatrix(NewRow, COL_检查报告ID) <> "" Then
                On Error GoTo errH
                strSQL = "Select 1 From 报告查阅记录 Where 医嘱ID=[1] And 查阅人=[2] And 取消时间 Is NULL"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "vsAdvice_AfterRowColChange", _
                    Val(vsAdvice.TextMatrix(NewRow, COL_ID)), UserInfo.姓名)
                If Not rsTmp.EOF Then
                    If vsAdvice.TextMatrix(NewRow, COL_检查报告ID) = "" Then
                        vsAdvice.Cell(flexcpData, NewRow, COL_查阅状态) = 1
                    Else
                        '部分查阅的
                        strSQL = "Select 1 From 病人医嘱报告 A Where not exists(select 1 from 报告查阅记录 B where B.医嘱ID=A.医嘱ID And A.检查报告ID=B.检查报告ID And B.查阅人=[2] And B.取消时间 Is NULL) and A.医嘱ID=[1] "
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "vsAdvice_AfterRowColChange", Val(vsAdvice.TextMatrix(NewRow, COL_ID)), UserInfo.姓名)
                        vsAdvice.Cell(flexcpData, NewRow, COL_查阅状态) = IIF(Not rsTmp.EOF, 2, 1)
                    End If
                Else
                    vsAdvice.Cell(flexcpData, NewRow, COL_查阅状态) = 0
                End If
                On Error GoTo 0
            End If
        
            '显示医嘱附加表格的内容
            If mblnAppend Then
                '判断单据附项是否有内容
                blnSel = False: blnExist = False
                Call ShowBillAppend(NewRow, blnExist)
                For intIdx = 0 To tbcAppend.ItemCount - 1
                    If tbcAppend(intIdx).Tag = "附项" Then
                        If tbcAppend(intIdx).Selected Then blnSel = True
                        tbcAppend(intIdx).Visible = blnExist
                        Exit For
                    End If
                Next
                If blnSel And Not blnExist Then
                    varDraw = vsAdvice.Redraw '根据条件屏蔽重复调用
                    vsAdvice.Redraw = flexRDNone
                    tbcAppend.Item(0).Selected = True
                    vsAdvice.Redraw = varDraw
                End If
                
                '判断附加信息的显示
                blnSel = False: blnExist = False
                Call ShowAdvicePlan(NewRow, blnExist)
                For intIdx = 0 To tbcAppend.ItemCount - 1
                    If tbcAppend(intIdx).Tag = "安排" Then
                        If tbcAppend(intIdx).Selected Then blnSel = True
                        tbcAppend(intIdx).Visible = blnExist
                        Exit For
                    End If
                Next
                If blnSel And Not blnExist Then
                    varDraw = vsAdvice.Redraw '根据条件屏蔽重复调用
                    vsAdvice.Redraw = flexRDNone
                    tbcAppend.Item(0).Selected = True
                    vsAdvice.Redraw = varDraw
                End If
                
                '判断医嘱是否审核(作废的医嘱不显示血液页面)
                blnSel = False: blnExist = False: bln输血 = False
                If gbln血库系统 And vsAdvice.TextMatrix(NewRow, COL_诊疗类别) = "K" Then
                    bln输血 = True
                    With vsAdvice
                        '用血医嘱审核状态=1表明是输血科发血产生的待核对医嘱，对于输血医嘱，审核状态=4，紧急医嘱和未用输血分级管理时，显示为等待配血
                        If Val(.TextMatrix(NewRow, COL_审核状态)) = 1 And Val(.TextMatrix(NewRow, COL_检查方法)) = 1 Then
                            blnExist = True
                        Else
                            blnExist = InStr(",,2,3,4,5,6,", "," & .TextMatrix(NewRow, COL_审核状态) & ",") > 0 And Not (.TextMatrix(NewRow, COL_医嘱状态) = "4")
                        End If
                    End With
                End If
                For intIdx = 0 To tbcAppend.ItemCount - 1
                    If tbcAppend(intIdx).Tag = "血液" Then
                        If tbcAppend(intIdx).Selected Then blnSel = True
                        tbcAppend(intIdx).Visible = blnExist
                        Exit For
                    End If
                Next
                If blnSel And Not blnExist Then
                    varDraw = vsAdvice.Redraw '根据条件屏蔽重复调用
                    vsAdvice.Redraw = flexRDNone
                    tbcAppend.Item(0).Selected = True
                    vsAdvice.Redraw = varDraw
                End If
                
                blnSel = False: blnExist = False
                If bln输血 = False Then
                    With vsAdvice
                        blnExist = InStr(",2,3,4,5,", "," & .TextMatrix(NewRow, COL_审核状态) & ",") > 0
                        '是输血医嘱时，用血库系统后才有为4的审核状态。紧急医嘱，未用输血分级管理时。 审核状态为4时没有相应的操作记录<病人医嘱状态>
                        If Val(.TextMatrix(NewRow, COL_审核状态)) = 4 And .TextMatrix(NewRow, COL_诊疗类别) = "K" Then
                            If Val(.TextMatrix(NewRow, COL_标志)) = 1 Or Not gbln输血分级管理 Then blnExist = False
                        End If
                    End With
                End If
                
                For intIdx = 0 To tbcAppend.ItemCount - 1
                    If tbcAppend(intIdx).Tag = "其他" Then
                        If tbcAppend(intIdx).Selected Then blnSel = True
                        tbcAppend(intIdx).Visible = blnExist
                        Exit For
                    End If
                Next
                If blnSel And Not blnExist Then
                    varDraw = vsAdvice.Redraw '根据条件屏蔽重复调用
                    vsAdvice.Redraw = flexRDNone
                    tbcAppend.Item(0).Selected = True
                    vsAdvice.Redraw = varDraw
                End If
                
                '判预约信息的显示
                blnSel = False: blnExist = False
                Call ShowAdviceRISSch(NewRow, blnExist)
                For intIdx = 0 To tbcAppend.ItemCount - 1
                    If tbcAppend(intIdx).Tag = "预约" Then
                        If tbcAppend(intIdx).Selected Then blnSel = True
                        tbcAppend(intIdx).Visible = blnExist
                        Exit For
                    End If
                Next
                If blnSel And Not blnExist Then
                    varDraw = vsAdvice.Redraw '根据条件屏蔽重复调用
                    vsAdvice.Redraw = flexRDNone
                    tbcAppend.Item(0).Selected = True
                    vsAdvice.Redraw = varDraw
                End If
                
                If tbcAppend.Selected.Tag = "计价" Then
                    Call ShowPrice(NewRow)
                ElseIf tbcAppend.Selected.Tag = "发送" Then
                    Call ShowSendList(NewRow)
                ElseIf tbcAppend.Selected.Tag = "签名" Then
                    Call ShowSignList(NewRow)
                ElseIf tbcAppend.Selected.Tag = "附项" Then
                    '前面已固定读取
                ElseIf tbcAppend.Selected.Tag = "预约" Then
                    '前面已固定读取
                ElseIf tbcAppend.Selected.Tag = "安排" Then
                    '前面已固定读取
                ElseIf tbcAppend.Selected.Tag = "其他" Then
                    Call ShowOtherAppend(NewRow)
                ElseIf tbcAppend.Selected.Tag = "血液" Then
                    If Not mobjFrmBloodList Is Nothing Then
                        Call mobjFrmBloodList.zlRefresh(Val(vsAdvice.TextMatrix(NewRow, COL_ID)), mlngFontSize, mblnMoved)
                    End If
                End If
            End If
        ElseIf mblnAppend Then
            Call ClearAppendData
            vsAppend.Row = vsAppend.FixedRows
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    If Col = col_医嘱内容 Or Col = col_内容 Then
        vsAdvice.AutoSize Col, COL_用法
    ElseIf Col = COL_皮试 Then
        If vsAdvice.ColWidth(Col) > 1200 Then vsAdvice.ColWidth(Col) = 1200
    ElseIf Row = -1 Then
        lngW = Me.TextWidth(vsAdvice.TextMatrix(vsAdvice.FixedRows - 1, Col) & "A")
        If vsAdvice.ColWidth(Col) < lngW Then
            vsAdvice.ColWidth(Col) = lngW
        ElseIf vsAdvice.ColWidth(Col) > vsAdvice.Width * 0.5 Then
            vsAdvice.ColWidth(Col) = vsAdvice.Width * 0.5
        End If
    End If
End Sub

Private Sub vsAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Row = -1 Then
        If Col <= vsAdvice.FixedCols - 1 Then
            Cancel = True
        ElseIf Col = COL_警示 Then 'Pass
            Cancel = True
        End If
    End If
End Sub

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'说明：1.OwnerDraw要设置为Over(画出单元所有内容)
'      2.Cell的GridLine从上下左右向内都是从第1根线开始
'      3.Cell的Border从左上是从第2根线开始,右下是从第1根线开始
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsAdvice
        If Col <= .FixedCols - 1 Then
            '擦除固定列中的表格线
            SetBkColor hDC, OS.SysColor2RGB(.BackColorFixed)

            '仅左边表格线
            vRect.Left = Left
            vRect.Top = Top
            vRect.Right = Left + 1
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '仅上边表格线
            vRect.Left = Left
            vRect.Top = Top
            vRect.Right = Right
            vRect.Bottom = Top + 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '仅下边表格线
            vRect.Left = Left
            vRect.Top = Bottom - 1
            vRect.Right = Right
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '仅右边表格线
            vRect.Left = Right - 1
            vRect.Top = Top
            vRect.Right = Right
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        ElseIf Col = COL_处方号 Or Col = COL_处方打印 Or Col = COL_处方预览 Then
            lngLeft = COL_处方号: lngRight = COL_处方预览
            If Not RowInSameNo(Row, lngBegin, lngEnd) Then Exit Sub
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
                '为了支持预览输出
                If .TextMatrix(Row, Col) <> "" Then .TextMatrix(Row, Col) = ""
            End If
            
            If Between(Row, .Row, .RowSel) Then
                SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
            Else
                SetBkColor hDC, OS.SysColor2RGB(.BackColor)
            End If
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Else
            '擦除一并给药相关行列的边线及内容
            lngLeft = COL_开始时间: lngRight = COL_开始时间
            If Not Between(Col, lngLeft, lngRight) Then
                lngLeft = COL_天数: lngRight = COL_用法
            End If
            If Not Between(Col, lngLeft, lngRight) Then
                lngLeft = COL_皮试: lngRight = COL_皮试
            End If
            If Not Between(Col, lngLeft, lngRight) Then Exit Sub
            
            If Not RowIn一并给药(Row, lngBegin, lngEnd) Then Exit Sub
            
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
                '为了支持预览输出
                If .TextMatrix(Row, Col) <> "" Then .TextMatrix(Row, Col) = ""
            End If
            
            If Between(Row, .Row, .RowSel) Then
                SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
            Else
                SetBkColor hDC, OS.SysColor2RGB(.BackColor)
            End If
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        End If
        Done = True
    End With
End Sub

Private Sub vsAdvice_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
        Dim rsTmp As Recordset
    
    If Button = 2 Then
        If mcbsMain Is Nothing Then Exit Sub
        Set objPopup = mcbsMain.ActiveMenuBar.FindControl(, conMenu_EditPopup)
        If Not objPopup Is Nothing Then
            objPopup.CommandBar.ShowPopup
        End If
    ElseIf Button = 1 Then
        If mvarCond.过滤模式 = 0 And mvarCond.医嘱 = 1 Then
            With vsAdvice
                If .MouseRow >= .FixedRows And (.MouseCol = COL_处方打印 Or .MouseCol = COL_处方预览) Then
                    If .TextMatrix(.MouseRow, .MouseCol) = "" Then
                        Exit Sub
                    End If
                Else
                    Exit Sub
                End If
                mbln处方预览 = (.MouseCol = COL_处方预览)
            End With
            vsAdvice.Redraw = flexRDNone
            If mcbsMain Is Nothing Then
                Set rsTmp = GetBillList
                If rsTmp.RecordCount > 0 Then
                    FuncBillPrint , "ZLCISBILL" & Format(rsTmp!编号, "00000") & "-1" & "|" & rsTmp!NO, rsTmp!名称 '对应的自定义报表编号
                End If
                Exit Sub
            End If
            Set objControl = mcbsMain.FindControl(, conMenu_Report_ClinicBill * 100# + 1, , True)
            If Not objControl Is Nothing Then
                objControl.Execute
            Else
                MsgBox "该药品没有对应诊疗单据，没有可以打印的处方签。", vbInformation, gstrSysName
            End If
            vsAdvice.Redraw = flexRDDirect
        End If
    End If
End Sub

Private Function GetBillList() As Recordset
    Dim strSQL As String
    With vsAdvice
        strSQL = "Select Distinct D.编号,D.名称,D.说明,B.NO" & _
            " From 病人医嘱记录 A,病人医嘱发送 B,病历单据应用 C,病历文件列表 D" & _
            " Where C.诊疗项目ID = A.诊疗项目ID And a.ID=b.医嘱ID " & _
            " And C.应用场合=1 And C.病历文件ID=D.ID And D.种类=7 And (a.ID=[1] or A.相关ID=[1])" & _
            " Order by D.编号"
       
        If mblnMoved Then
            strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
            strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
        End If
         Set GetBillList = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Decode(Val(.TextMatrix(.Row, COL_相关ID)), 0, Val(.TextMatrix(.Row, COL_ID)), Val(.TextMatrix(.Row, COL_相关ID))))
    End With
End Function

Private Function GetPatiInfo() As ADODB.Recordset
'功能：读取病人信息
    Dim strSQL As String
    
    On Error GoTo errH
    
    '执行部门(号别科室)即病人科室
    strSQL = "Select A.姓名,A.性别,A.年龄,B.门诊号,B.住院号,B.健康号,a.ID as 挂号ID," & _
        " B.险类,B.就诊诊室,C.名称 as 执行部门,A.登记时间,B.费别" & _
        " From 病人挂号记录 A,病人信息 B,部门表 C" & _
        " Where A.NO(+)=[2] And a.记录性质(+)=1 And a.记录状态(+)=1 And B.病人ID=[1]" & _
        " And A.病人ID(+)=B.病人ID And A.执行部门ID=C.ID(+)"
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人挂号记录", "H病人挂号记录")
    End If
    Set GetPatiInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng病人ID, mstr挂号单)
        
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub OutputList(bytStyle As Byte)
'功能：输入出列表
'参数：bytStyle=1-打印,2-预览,3-输出到Excel
    Dim objOut As New zlPrint1Grd
    Dim objRow As zlTabAppRow
    Dim bytR As Byte, i As Long
    Dim lngRow As Long, lngCol As Long
    Dim strWidth As String, strInfo As String, rsTmp As ADODB.Recordset
    
    If mlng病人ID = 0 Then Exit Sub
    
    '表头
    objOut.Title.Text = "病人医嘱清单"
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '表上
    Set rsTmp = GetPatiInfo
    strInfo = _
        "姓名：" & rsTmp!姓名 & " 性别：" & NVL(rsTmp!性别) & _
        " 年龄：" & NVL(rsTmp!年龄) & " 门诊号：" & NVL(rsTmp!门诊号) & _
        " 挂号：" & IIF(IsNull(rsTmp!登记时间), "", Format(rsTmp!登记时间, "yyyy-MM-dd HH:mm")) & _
        " 科室：" & NVL(rsTmp!执行部门) & " 诊室：" & NVL(rsTmp!就诊诊室)
    Set objRow = New zlTabAppRow
    objRow.Add strInfo
    objOut.UnderAppRows.Add objRow
    
    '表下
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(zlDatabase.Currentdate(), "yyyy年MM月dd日")
    objOut.BelowAppRows.Add objRow
    
    '表体
    Set objOut.Body = vsAdvice
    
    '输出
    vsAdvice.Redraw = False
    lngRow = vsAdvice.Row: lngCol = vsAdvice.Col
    
    strWidth = ""
    For i = 0 To vsAdvice.FixedCols - 1
        strWidth = strWidth & "," & vsAdvice.ColWidth(i)
        vsAdvice.ColWidth(i) = 0
    Next
        
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    strWidth = Mid(strWidth, 2)
    For i = 0 To vsAdvice.FixedCols - 1
        vsAdvice.ColWidth(i) = Split(strWidth, ",")(i)
    Next
    
    vsAdvice.Row = lngRow: vsAdvice.Col = lngCol
    vsAdvice.Redraw = True
End Sub

Private Sub Form_Load()
    Dim objPane As Pane
    Dim strTab As String, i As Integer
    Dim intType As Integer
    
    mblnFirst = False
    Set mrsPlugInBar = Nothing
    mlngPromptRow = 0
    mbln处方预览 = False
    
    If Not grsSkinTest Is Nothing Then
        grsSkinTest.Close
        Set grsSkinTest = Nothing
    End If
    
    '医嘱清单
    '-----------------------------------------------------
    mlngFontSize = 9
    Call InitAdviceTable
    Call InitColumnSelect '初始化列选择器
    
    'CommandBars
    '-----------------------------------------------------
    Call GetFilterSetting '本地过滤参数
    Call InitFilterBar
    
    'TabControl
    '-----------------------------------------------------
    With tbcMain
        With .PaintManager
            .Appearance = xtpTabAppearanceVisualStudio
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
        End With
        .InsertItem(0, " 医  嘱 ", picMain.hwnd, 0).Tag = "医嘱"
        .InsertItem(1, " 报  告 ", picMain.hwnd, 0).Tag = "报告"
    End With
    tbcMain.Item(tbcMain.ItemCount - 1).Selected = True
    i = IIF(mvarCond.过滤模式 = 0, 0, 1)
    tbcMain.Item(i).Selected = True
    
    With tbcAppend
        With .PaintManager
            .Appearance = xtpTabAppearanceVisualStudio
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
        End With
        .InsertItem(0, "医嘱计价内容", vsAppend.hwnd, 0).Tag = "计价"
        .InsertItem(1, "医嘱发送记录", vsAppend.hwnd, 0).Tag = "发送"
        If Not gobjESign Is Nothing Then '电子签名记录
            .InsertItem(2, "医嘱签名记录", vsAppend.hwnd, 0).Tag = "签名"
        End If
        .InsertItem(3, "申请附项", rtfAppend.hwnd, 0).Tag = "附项"
        .InsertItem(4, "安排情况", rtfInfo.hwnd, 0).Tag = "安排"
        .InsertItem(5, "预约信息", rtfSche.hwnd, 0).Tag = "预约" 'RIS预约信息
        .InsertItem(6, "其他信息", rtfOther.hwnd, 0).Tag = "其他" '抗菌药物审核信息
        If gbln血库系统 = True Then
            If InitObjBlood = True Then
                Set mobjFrmBloodList = gobjPublicBlood.zlGetBloodListInfo
                .InsertItem(7, "血液信息", mobjFrmBloodList.hwnd, 0).Tag = "血液"  '血液配血信息
            End If
        End If
        '因为绑定相同,最后要切换回第1个;无数据不影响速度
        .Item(.ItemCount - 1).Selected = True
        .Item(0).Selected = True
    End With
    mblnAppend = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "AppendData", 1)) <> 0
    tbcAppend.Visible = mblnAppend: fraAdviceUD.Visible = mblnAppend
    If mblnAppend Then
        strTab = zlDatabase.GetPara("医嘱子列表", glngSys, p门诊医嘱下达, "")
        If strTab <> "" Then
            For i = 0 To tbcAppend.ItemCount - 1
                If tbcAppend(i).Visible And tbcAppend(i).Tag = strTab Then
                    tbcAppend.Item(i).Selected = True
                    Exit For
                End If
            Next
        End If
    End If
        
    '恢复个性化设置
    '-----------------------------------------------------
    Call RestoreWinState(Me, App.ProductName)
    vsAdvice.ColWidth(COL_F标志) = 11 * Screen.TwipsPerPixelX
    vsAdvice.ColWidth(COL_F报告) = 11 * Screen.TwipsPerPixelX
    
    '变量初始化
    '-----------------------------------------------------
    mMainPrivs = gMainPrivs '主界面模块权限
    Set mfrmEdit = Nothing
    Set mobjReport = New clsReport
    Set mrsDefine = InitAdviceDefine
    
    
    '本地注册表设置
    Call GetLocalSetting
    mblnAutoRead = Val(zlDatabase.GetPara("自动标记报告查阅状态", glngSys, p门诊医嘱下达, "1", , , intType)) = 1
    mblnAutoReadEnabled = Not ((intType = 3 Or intType = 15))
        
    If gblnKSSStrict Then Call CheckKSSPrivilege(2)
    If mint场合 = 0 Then Call InitObjLis(p门诊医生站)
    On Error Resume Next
    Set gobjExchange = CreateObject("zlExchange.clsExchange")
    If Not gobjExchange Is Nothing Then Call gobjExchange.Init(gcnOracle)
    err.Clear: On Error GoTo 0
End Sub

Private Sub InitFilterBar()
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsSub.VisualTheme = xtpThemeOffice2003
    With Me.cbsSub.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize False, 16, 16
        .UseSharedImageList = False 'ImageList方式时,因同一App中共享,在AddImageList之前设置为False
    End With
    cbsSub.AddImageList img16 '以VB.ImageList的Tag与ID进行关联
    cbsSub.EnableCustomization False
    cbsSub.ActiveMenuBar.Visible = False
    
    Set objBar = cbsSub.Add("工具栏", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
    With objBar.Controls
        Set objPopup = .Add(xtpControlPopup, ID_婴儿, "病人医嘱")
            objPopup.ID = ID_婴儿: objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, ID_婴儿 * 100#, "所有医嘱")
            Set objControl = .Add(xtpControlButton, ID_婴儿 * 100# + 1, "病人医嘱"): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, ID_婴儿 * 100# + 2, "婴儿 1 医嘱"): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, ID_婴儿 * 100# + 3, "婴儿 2 医嘱")
            Set objControl = .Add(xtpControlButton, ID_婴儿 * 100# + 4, "婴儿 3 医嘱")
            Set objControl = .Add(xtpControlButton, ID_婴儿 * 100# + 5, "婴儿 4 医嘱")
            Set objControl = .Add(xtpControlButton, ID_婴儿 * 100# + 6, "婴儿 5 医嘱")
        End With
        
        Set objControl = .Add(xtpControlButton, ID_废止, "已作废")
            objControl.BeginGroup = True
            objControl.ToolTipText = "显示已经作废的医嘱"
            
        '----------------报告页面
        Set objControl = .Add(xtpControlButton, ID_全部, "全部")
            objControl.BeginGroup = True
            objControl.Checked = True
        Set objControl = .Add(xtpControlButton, ID_检查, "检查")
            objControl.IconId = 1 '初始时不置图标
        Set objControl = .Add(xtpControlButton, ID_检验, "检验")
        Set objControl = .Add(xtpControlButton, ID_其他, "其他")
            objControl.IconId = 1
        
        Set objControl = .Add(xtpControlButton, ID_未出报告, "未出报告")
            objControl.ToolTipText = "显示未出报告"
            objControl.BeginGroup = True
            mvarCond.未出报告 = True
            
        Set objControl = .Add(xtpControlButton, ID_已出报告, "已出报告")
            objControl.ToolTipText = "显示已出报告"
            mvarCond.已出报告 = True
        
        Set objControl = .Add(xtpControlButton, ID_医嘱全部, "全部")
            objControl.BeginGroup = True
            objControl.Checked = True
        Set objControl = .Add(xtpControlButton, ID_医嘱处方, "处方")
            objControl.IconId = 1
        Set objControl = .Add(xtpControlButton, ID_医嘱其他, "其他")
            objControl.IconId = 1
        '-----------------医嘱
        
        
        Set objControl = .Add(xtpControlButton, ID_科内, "本科下达")
            objControl.BeginGroup = True
            objControl.ToolTipText = "只显示医技本科下达的医嘱"
        
        
        Set objControl = .Add(xtpControlButton, ID_简洁, "简洁")
            objControl.BeginGroup = True
            objControl.Flags = xtpFlagRightAlign
        Set objControl = .Add(xtpControlButton, ID_完整, "详细")
            objControl.BeginGroup = True
            objControl.Flags = xtpFlagRightAlign
            
    End With
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    With cbsSub.KeyBindings
        .Add FCONTROL, vbKeyB, ID_婴儿 * 100#
        .Add FCONTROL, vbKey0, ID_婴儿 * 100# + 1
        .Add FCONTROL, vbKey1, ID_婴儿 * 100# + 2
        .Add FCONTROL, vbKey2, ID_婴儿 * 100# + 3
        .Add FCONTROL, vbKey3, ID_婴儿 * 100# + 4
        .Add FCONTROL, vbKey4, ID_婴儿 * 100# + 5
        .Add FCONTROL, vbKey5, ID_婴儿 * 100# + 6
        .Add FCONTROL, vbKey8, ID_废止
        .Add FCONTROL, vbKeyK, ID_科内
    End With
    objBar.Visible = Not mblnHideFilter
    fraHide.Visible = mblnHideFilter
    fraHide.BackColor = cbsSub.GetSpecialColor(STDCOLOR_BTNFACE)
End Sub

Private Sub mfrmParent_KeyDown(KeyCode As Integer, Shift As Integer)
'功能：捕获主窗体的按键,用于处理医嘱过滤热键
'说明：
'1.当医嘱子窗体未激活时,子窗体CommandBar的热键无效
'2.主窗体CommandBar或KeyDown事件处理了的键不会再激活该事件
    
    If Not Me.Visible Then Exit Sub '在其他子窗体时仍会激活
    If mlng病人ID = 0 Then Exit Sub
    
    Call ActiveHotKey(KeyCode, Shift)

End Sub

Private Sub ActiveHotKey(KeyCode As Integer, Shift As Integer)
    Dim objControl As CommandBarControl
    Dim lngID As Long
    Dim intTab As Integer
    
    If Not Me.Visible Then Exit Sub
    If mlng病人ID = 0 Then Exit Sub
    intTab = -1
    
    If Shift = vbCtrlMask And KeyCode >= vbKey0 And KeyCode <= vbKey5 Then
        lngID = ID_婴儿 * 100# + KeyCode - vbKey0 + 1
    ElseIf Shift = vbCtrlMask Then
        Select Case KeyCode
            Case vbKey6
                intTab = 0
            Case vbKey8
                lngID = ID_废止
            Case vbKey9
                intTab = 1
            Case vbKeyB
                lngID = ID_婴儿 * 100#
            Case vbKeyK
                lngID = ID_科内
            Case vbKeyX
                lngID = ID_检查
            Case vbKeyY
                lngID = ID_检验
            Case vbKeyQ
                lngID = ID_其他
        End Select
    End If
    If lngID <> 0 Then
        Set objControl = cbsSub.FindControl(, lngID, , True)
        If Not objControl Is Nothing Then objControl.Execute
    End If
    If intTab <> -1 Then tbcMain.Item(intTab).Selected = True
End Sub

Private Sub GetLocalSetting()
'功能：读取本地注册表设置
    mbln皮试限制 = Val(zlDatabase.GetPara("医嘱发送皮试限制", glngSys, p门诊医嘱下达)) <> 0
    '执行天数
    mbln天数 = Val(zlDatabase.GetPara("医嘱执行天数", glngSys, p门诊医嘱下达)) <> 0
    
    mbln指引单打印 = Val(zlDatabase.GetPara("指引单打印方式", glngSys, p门诊医嘱下达)) <> 0
    
    mbln危急值 = InStr(GetInsidePrivs(p门诊医生站), ";危急值处理;") > 0
End Sub

Private Sub GetFilterSetting()
'功能：读取医嘱过滤设置条件
    Dim strPar As String
    
    mvarCond.婴儿 = 0
    mvarCond.废止 = Val(zlDatabase.GetPara("医嘱显示作废", glngSys, p门诊医嘱下达, "1")) = 1
    mvarCond.科内 = Val(zlDatabase.GetPara("科内医嘱过滤", glngSys, p门诊医嘱下达, "1")) <> 0
    mblnHideFilter = Val(zlDatabase.GetPara("过滤条件自动隐藏", glngSys, p门诊医嘱下达, "0")) <> 0
    
    strPar = Val(zlDatabase.GetPara("报告查看类型", glngSys, p门诊医嘱下达, "0"))
    If InStr(",0,1,2,3,", "," & strPar & ",") > 0 Then
        mvarCond.报告 = Val(strPar)
    Else
        mvarCond.报告 = 0
    End If
    
    strPar = Val(zlDatabase.GetPara("显示模式", glngSys, p门诊医嘱下达, "0"))
    mvarCond.显示模式 = IIF(Val(strPar) = 0, 0, 1)
    
    strPar = Val(zlDatabase.GetPara("医嘱显示处方", glngSys, p门诊医嘱下达, "0"))
    If InStr(",0,1,2,", "," & strPar & ",") > 0 Then
        mvarCond.医嘱 = Val(strPar)
    Else
        mvarCond.医嘱 = 0
    End If
    
    strPar = Val(zlDatabase.GetPara("医嘱过滤方式", glngSys, p门诊医嘱下达, "0"))
    mvarCond.过滤模式 = IIF(Val(strPar) = 0, 0, 3)
End Sub

Private Sub SaveFilterSetting()
'功能：保存医嘱过滤设置条件
    Call zlDatabase.SetPara("科内医嘱过滤", IIF(mvarCond.科内, 1, 0), glngSys, p门诊医嘱下达)
    Call zlDatabase.SetPara("显示模式", mvarCond.显示模式, glngSys, p门诊医嘱下达)
    Call zlDatabase.SetPara("报告查看类型", mvarCond.报告, glngSys, p门诊医嘱下达)
    Call zlDatabase.SetPara("过滤条件自动隐藏", IIF(mblnHideFilter, 1, 0), glngSys, p门诊医嘱下达)
    Call zlDatabase.SetPara("医嘱显示处方", mvarCond.医嘱, glngSys, p门诊医嘱下达)
    Call zlDatabase.SetPara("医嘱显示作废", IIF(mvarCond.废止, 1, 0), glngSys, p门诊医嘱下达)
    Call zlDatabase.SetPara("医嘱过滤方式", mvarCond.过滤模式, glngSys, p门诊医嘱下达)
End Sub

Private Sub Form_Resize()
    If WindowState = 1 Then Exit Sub
    With Me.tbcMain
        .Left = 0
        .Top = 0
        .Height = Me.Height
        .Width = Me.Width
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mfrmSend Is Nothing Then Unload mfrmSend: Set mfrmSend = Nothing
    If Not mfrmEdit Is Nothing Then Unload mfrmEdit: Set mfrmEdit = Nothing
    Set mobjReport = Nothing
    Set gobjExchange = Nothing
    Set gobjLIS = Nothing
    Set mobjPublicPACS = Nothing
    Set gobjRecipeAudit = Nothing
    
    If Not mobjFrmBloodList Is Nothing Then
        Unload mobjFrmBloodList
        Set mobjFrmBloodList = Nothing
    End If
    Set mrsDefine = Nothing
    Set mobjVBA = Nothing
    Set mobjScript = Nothing
    Set mSendControl = Nothing
    If Not gobjEmrInterface Is Nothing Then
        Set gobjEmrInterface = Nothing
    End If
        
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "AppendData", IIF(mblnAppend, 1, 0)
    If mblnAppend And Not tbcAppend.Selected Is Nothing Then
        Call zlDatabase.SetPara("医嘱子列表", tbcAppend.Selected.Tag, glngSys, p门诊医嘱下达)
    End If
    Call SaveFilterSetting
    Call SaveWinState(Me, App.ProductName)
    'PASS
    If mblnPass Then
        Call gobjPass.zlPassClearLight(mobjPassMap, 1)
    End If
    mblnPass = False
    Set mobjPassMap = Nothing
 
    '外挂程序对象终止
    If CreatePlugInOK(p门诊医嘱下达, mint场合) Then
        On Error Resume Next
        Call gobjPlugIn.Terminate(glngSys, p门诊医嘱下达, mint场合)
        Call zlPlugInErrH(err, "Terminate")
        err.Clear: On Error GoTo 0
    End If
    Set mclsMipModule = Nothing
    mbln处方预览 = False
    Set mrs危急值 = Nothing
    mbln危急值 = False
    mlng危急值ID = 0
End Sub

Private Sub ClearAppendData()
'功能：清除附加表格和申请附项的数据
    Dim blnSel As Boolean, intIdx As Integer
    Dim varDraw As RedrawSettings
    
    vsAppend.Rows = vsAppend.FixedRows
    vsAppend.Rows = vsAppend.FixedRows + 1
    vsAppend.Row = vsAppend.FixedRows
        
    If rtfAppend.Visible Then rtfAppend.Text = ""
    If rtfInfo.Visible Then rtfInfo.Text = ""
    For intIdx = 0 To tbcAppend.ItemCount - 1
        If InStr("附项,安排,预约,其他,血液", tbcAppend(intIdx).Tag) > 0 Then
            If tbcAppend(intIdx).Selected Then blnSel = True
            tbcAppend(intIdx).Visible = False
        End If
    Next
    If blnSel Then
        varDraw = vsAdvice.Redraw '根据条件屏蔽重复调用
        vsAdvice.Redraw = flexRDNone
        tbcAppend.Item(0).Selected = True
        vsAdvice.Redraw = varDraw
    End If
End Sub

Private Sub InitPriceTable()
'功能：初始化计价清单格式
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "计价医嘱,2000,1;类别,650,1;收费项目,2500,1;单位,500,4;计价数量,850,1;单价,900,7;执行科室,1000,1;费用类型,800,1;从项,450,4;收费方式,1500,1"
    arrHead = Split(strHead, ";")
    With vsAppend
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
            If COLPrice.Count <> UBound(arrHead) + 1 Then COLPrice.Add i, Split(arrHead(i), ",")(0)
            .MergeCol(i) = False
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .MergeCells = flexMergeRestrictAll
        .MergeCompare = flexMCIncludeNulls
    End With
End Sub

Private Sub InitSendTable()
'功能：初始化发送清单格式
    Dim arrHead As Variant, strHead As String, i As Long, lngWidth As Long
    
    strHead = "发送号;发送时间,1530,1;单据号,850,1;发送医嘱,1800,1;收费项目,1800,1;发送数次,850,1;计费状态,850,1;执行状态,850,1;状态说明,1800,1;执行科室,1000,1;执行人,800,1;执行时间,1530,1;执行说明,1800,1;发送人,800,1;记录性质"
    arrHead = Split(strHead, ";")
    With vsAppend
        .Redraw = flexRDNone
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                lngWidth = Val(Split(arrHead(i), ",")(1))
                If mlngFontSize <> 9 Then
                    lngWidth = lngWidth * mlngFontSize / 9
                    If lngWidth > .ColWidthMax And .ColWidthMax <> 0 Then lngWidth = .ColWidthMax
                End If
                .ColWidth(.FixedCols + i) = lngWidth
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
            If COLSend.Count <> UBound(arrHead) + 1 Then COLSend.Add i, Split(arrHead(i), ",")(0)
            .MergeCol(i) = False
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .Redraw = flexRDDirect
        
        .MergeCells = flexMergeRestrictAll
        .MergeCompare = flexMCIncludeNulls
    End With
End Sub

Private Sub InitSignTable()
'功能：初始化签名清单格式
    Dim arrHead As Variant, strHead As String, i As Long, lngWidth As Long
    
    strHead = "签名类型,1150,1;签名时间,1900,1;签名人,800,1;时间戳,1900,1"
    arrHead = Split(strHead, ";")
    With vsAppend
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                lngWidth = Val(Split(arrHead(i), ",")(1))
                If mlngFontSize <> 9 Then
                    lngWidth = lngWidth * mlngFontSize / 9
                    If lngWidth > .ColWidthMax And .ColWidthMax <> 0 Then lngWidth = .ColWidthMax
                End If
                .ColWidth(.FixedCols + i) = lngWidth
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
            If COLSign.Count <> UBound(arrHead) + 1 Then COLSign.Add i, Split(arrHead(i), ",")(0)
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .MergeCells = flexMergeNever
    End With
End Sub

Private Sub ClearAdviceData()
'功能：清除医嘱清单数据
    vsAdvice.Rows = vsAdvice.FixedRows
    vsAdvice.Rows = vsAdvice.FixedRows + 1
    vsAdvice.Editable = flexEDNone
End Sub

Private Sub InitColumnSelect()
'功能：根据医嘱清单原始列显示状态初始化列选择器
    Dim lngRow As Long, i As Long
    
    vsColumn.Rows = vsColumn.FixedRows
    With vsAdvice
        For i = .FixedCols To .Cols - 1
            If Not (.ColHidden(i) Or .ColWidth(i) = 0) Then
                If .TextMatrix(0, i) <> "" And Not (i = COL_查阅状态 Or i = COL_标本状态) Then '审查结果,皮试
                    vsColumn.Rows = vsColumn.Rows + 1
                    lngRow = vsColumn.Rows - 1
                    vsColumn.TextMatrix(lngRow, 1) = .TextMatrix(0, i)
                    vsColumn.RowData(lngRow) = i
                    
                    '固定显示列
                    If InStr(",开始时间,医嘱内容,开嘱医生,", "," & .TextMatrix(0, i) & ",") > 0 Then
                        vsColumn.TextMatrix(lngRow, 0) = 1
                        vsColumn.Cell(flexcpForeColor, lngRow, 0, lngRow, 1) = vsColumn.BackColorFixed
                    End If
                End If
            End If
        Next
    End With
    If vsColumn.Rows > 1 Then vsColumn.Row = 1
End Sub

Private Sub InitAdviceTable()
'功能：初始化医嘱清单格式
    Dim arrHead As Variant, strHead As String, i As Long, lngWidth As Long

    strHead = "ID;相关ID;婴儿ID;医嘱状态;诊疗类别;操作类型;毒理分类;标志;" & _
        ",240,4;处方号,1000,4;打印,800,4;预览,800,4;生效时间,1530,1;,200,7;医嘱内容,3000,1;内容,4000,1;,375,1;总量,850,1;单量,850,1;天数,450,1;频率,1000,1;" & _
        "用法,1000,1;医生嘱托,1000,1;执行时间,1000,1;执行科室,1000,1;执行性质,850,1;开嘱医生,850,1;开嘱时间,1530,1;发送人,850,1;发送时间,1530,1;超量说明,1000,1;基本药物,850,1;查阅状态,700,4;标本状态,850,1;诊疗项目ID;试管编码;" & _
        "前提ID;签名否;文件ID;报告项;报告ID;审核状态;申请序号;高危药品;标本部位;收费细目ID;开嘱科室ID;用药目的;检查报告ID;处方审查状态;处方审查结果;RIS预约ID;RIS报告ID;LIS报告ID;RIS预约状态;诊疗项目名称;检查方法;危急值ID;易跌倒"
    arrHead = Split(strHead, ";")
    With vsAdvice
        .Clear
        .FixedRows = 1: .FixedCols = 2
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1

        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                lngWidth = Val(Split(arrHead(i), ",")(1))
                If mlngFontSize <> 9 Then
                    lngWidth = lngWidth * mlngFontSize / 9
                    If lngWidth > .ColWidthMax And .ColWidthMax <> 0 Then lngWidth = .ColWidthMax
                End If
                .ColWidth(.FixedCols + i) = lngWidth
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                '为了支持zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0    '为了支持zl9PrintMode
            End If
            .ColData(.FixedCols + i) = .ColWidth(.FixedCols + i)    '记录原始列宽用于列选择器
        Next
        '未启用合理用药时，该列不可见，启用美康，太元通时，即当gbytPass=1 or 3 时 可见
        vsAdvice.ColHidden(COL_警示) = True
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .ColWidth(COL_F标志) = 11 * Screen.TwipsPerPixelX
        .ColWidth(COL_F报告) = 11 * Screen.TwipsPerPixelX
        .MergeCells = flexMergeFree
        .MergeCol(COL_处方号) = True
        .MergeCol(COL_处方打印) = True
        .MergeCol(COL_处方预览) = True
    End With
End Sub

Private Sub SetRTFFont(bytKind As Byte)
    If bytKind = 0 Or bytKind = 1 Then
        With rtfAppend
            .SelStart = 0
            .SelLength = Len(.Text)
            .SelFontSize = mlngFontSize
            .SelLength = 0
        End With
    End If
    If bytKind = 0 Or bytKind = 2 Then
        With rtfInfo
            .SelStart = 0
            .SelLength = Len(.Text)
            .SelFontSize = mlngFontSize
            .SelLength = 0
        End With
    End If
    If bytKind = 0 Or bytKind = 3 Then
        With rtfOther
            .SelStart = 0
            .SelLength = Len(.Text)
            .SelFontSize = mlngFontSize
            .SelLength = 0
        End With
    End If
    If bytKind = 0 Or bytKind = 4 Then
        With rtfSche
            .SelStart = 0
            .SelLength = Len(.Text)
            .SelFontSize = mlngFontSize
            .SelLength = 0
        End With
    End If
End Sub

Private Function LoadAdvice() As Boolean
'功能：根据当前界面设置读取并显示医嘱清单
    Dim rsTmp As ADODB.Recordset
    Dim rs血型 As ADODB.Recordset
    Dim strSQL As String
    Dim strFormat As String, strTmp As String, blnDo As Boolean
    Dim bln给药途径 As Boolean, bln中药用法 As Boolean
    Dim bln采集方法 As Boolean, bln输血途径 As Boolean
    Dim blnFirst As Boolean, lng医嘱ID As Long
    Dim strBill As String, i As Long, j As Long
    Dim strWhere As String, str医嘱状态 As String
    Dim strSameDay As String '同一天
    Dim datCur As Date, strGroupBy As String
 
    If mlng病人ID = 0 Then Exit Function

    Screen.MousePointer = 11

    On Error GoTo errH
    
    lng医嘱ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))    '记录当前行如果是在当前界面刷新医嘱行应该不变
    
    If mvarCond.婴儿 <> -1 Then
        strWhere = strWhere & " And Nvl(A.婴儿,0)=[4]"
    End If
    
    If Not mvarCond.废止 Then
        strWhere = strWhere & " And Instr(',1,8,',','||Nvl(A.医嘱状态,0)||',')>0"
    End If
    
    '医技站  本科下达
    If mlng前提ID <> 0 And mvarCond.科内 Then
        strWhere = strWhere & " And Nvl(A.前提ID,0)<>0 and (A.前提ID in (Select /*+cardinality(x,10)*/ x.Column_Value From Table(Cast(f_Num2list([3]) As zlTools.t_Numlist)) X) or a.开嘱科室ID=[5])"
    End If
    
    '医嘱记录：不含附加手术,手术麻醉,检查部位,中药煎法
    strSQL = _
    "Select /*+ RULE */ A.ID,A.相关ID," & _
             " Nvl(A.婴儿,0) as 婴儿ID,A.医嘱状态,A.诊疗类别,B.操作类型,C.毒理分类,A.紧急标志 as 标志," & _
             " A.审查结果,k.No as 处方号,Decode(k.no,null,null,'打印') as 打印,Decode(k.no,null,null,'预览') as 预览,To_Char(A.开始执行时间,'YYYY-MM-DD HH24:MI') as 开始时间,Null as 并,A.医嘱内容,Null as 内容,A.皮试结果 as 皮试," & _
             " Decode(A.总给予量,NULL,NULL,Decode(A.诊疗类别,'E',Decode(B.操作类型,'4',A.总给予量||'付',A.总给予量||B.计算单位),'4',A.总给予量||G.计算单位,'5',Round(A.总给予量/D.门诊包装,5)||D.门诊单位,'6',Round(A.总给予量/D.门诊包装,5)||D.门诊单位,A.总给予量||B.计算单位)) as 总量," & _
             " Decode(A.单次用量,NULL,NULL,A.单次用量||Decode(A.诊疗类别,'4',G.计算单位,B.计算单位)) as 单量,A.天数," & _
             " A.执行频次 as 频率,Decode(A.诊疗类别,'E',Decode(Instr('2468',Nvl(B.操作类型,'0')),0,NULL,B.名称),NULL) as 用法," & _
             " A.医生嘱托,A.执行时间方案 as 执行时间,Nvl(E.名称,Decode(Nvl(A.执行性质,0),0,'<叮嘱>',5,'-')) as 执行科室," & _
             " Decode(Instr('567E',A.诊疗类别),0,NULL,A.执行性质) as 执行性质,A.开嘱医生,To_Char(A.开嘱时间,'YYYY-MM-DD HH24:MI') as 开嘱时间," & _
             " A.停嘱医生 as 发送人,A.停嘱时间 as 发送时间,a.超量说明,D.基本药物,Decode(Max(NVL(y.查阅状态,0)),MiN(NVL(y.查阅状态,0)),Max(NVL(y.查阅状态,0)),2) As 查阅状态,null as 标本状态,A.诊疗项目ID,B.试管编码,A.前提ID,Decode(A.新开签名ID,NULL,0,1) as 签名否," & _
             " M.病历文件ID as 文件ID,Nvl(N.通用,0) as 报告项,Max(y.病历id) As 报告id,A.审核状态,A.申请序号,d.高危药品,A.标本部位,A.收费细目ID,a.开嘱科室ID,a.用药目的," & _
             " Max(y.检查报告id)||'' As 检查报告id,J.状态 as 处方审查状态,J.审查结果 as 处方审查结果,f.预约id As RIS预约ID,Max(y.RISID) As RIS报告ID,Max(y.报告ID) as LIS报告ID,f.是否调整 as RIS预约状态,b.名称 as 诊疗项目名称,Max(a.检查方法) as 检查方法,max(h.危急值id) as 危急值ID,D.是否易至跌倒"
    strSQL = strSQL & _
             " From 病人医嘱记录 A,部门表 E,药品特性 C,药品规格 D,诊疗项目目录 B,收费项目目录 G,病人医嘱报告 Y,病历单据应用 M,病历文件列表 N, 处方审查明细 I, 处方审查记录 J,病人医嘱发送 K,RIS检查预约 f,病人危急值医嘱 H" & _
             " Where A.诊疗项目ID=B.ID And A.执行科室ID=E.ID(+) And A.诊疗项目ID=C.药名ID(+) And a.ID = i.医嘱ID(+) And I.审方ID = J.ID(+) and (I.最后提交 =1 Or I.审方ID is NULL) and Nvl(A.执行标记,0)<>-1 " & _
             " And a.id=k.医嘱id(+) And Nvl(A.医嘱期效,0)=1 And A.收费细目ID=D.药品ID(+) And A.收费细目ID=G.ID(+) And a.Id=f.医嘱id(+) and a.id=h.医嘱ID(+)" & _
             " And A.ID=Y.医嘱ID(+) And (Not(A.诊疗类别 IN ('F','G','D','E') And A.相关ID is Not NULL) Or A.诊疗类别='E' And B.操作类型='8')" & _
             " And A.开始执行时间 is Not NULL" & IIF(mint场合 = 2, "", " And A.病人来源<>3") & _
             " And A.诊疗项目ID=M.诊疗项目ID(+) And M.应用场合(+)=1 And M.病历文件ID=N.ID(+) And N.种类(+)=7" & _
             " And A.病人ID+0=[1] And A.挂号单=[2]" & strWhere
    strGroupBy = " Group By a.Id, a.相关id, a.序号, a.婴儿, a.医嘱状态, a.诊疗类别, b.操作类型, c.毒理分类, a.紧急标志, a.审查结果, a.医嘱期效, a.开始执行时间, a.医嘱内容, a.皮试结果," & vbNewLine & _
            "         a.总给予量, a.首次用量, g.计算单位, d.门诊包装, d.门诊单位, a.单次用量, a.天数, a.执行频次, a.医生嘱托, b.名称, a.执行性质, a.执行时间方案, a.执行终止时间, e.名称," & vbNewLine & _
            "         a.上次执行时间, a.开嘱时间, a.开嘱医生, a.校对护士, a.校对时间, a.停嘱医生, a.停嘱时间, a.确认停嘱护士, a.确认停嘱时间, a.诊疗项目id, b.试管编码, a.执行标记, a.屏蔽打印," & vbNewLine & _
            "         a.前提id, a.新开签名id, m.病历文件id, n.通用, a.收费细目id, b.计算单位, a.开嘱科室id, a.审核状态, a.申请序号, a.审核标记, d.基本药物, d.高危药品, a.标本部位,J.状态,J.审查结果," & vbNewLine & _
            "         a.用药目的,a.超量说明,k.No,f.预约id,f.是否调整,b.名称,D.是否易至跌倒"
     
    strSQL = strSQL & strGroupBy & " Order by Nvl(A.婴儿,0),A.序号"

    If mblnMoved Then    '挂号单与医嘱同个数据库
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        strSQL = Replace(strSQL, "病人医嘱报告", "H病人医嘱报告")
    End If
    datCur = zlDatabase.Currentdate
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng病人ID, mstr挂号单, IIF(mstr前提IDs = "", "0", mstr前提IDs), mvarCond.婴儿, mlng界面科室ID)
    If Not rsTmp.EOF Then
        strSQL = "Select a.医嘱id,decode(a.输血血型,1,'A',2,'B',3,'AB',4,'O','') As 血型 From 输血申请记录 A, 病人医嘱记录 B Where 医嘱id = b.Id And b.挂号单 =[1] And a.输血血型>0 and b.诊疗类别='K'"
        Set rs血型 = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mstr挂号单)
    
        With vsAdvice
            .Redraw = False
                
            '绑定时按设计时的FormatString恢复一些缺省值(固定行列数，固定行列文字及行列对齐,尺寸,可见)
            'FormatString在运行时赋值无效
            '如果AutoResize=True,则所有列宽或行高被自动调整(根据AutoSizeMode)
            '如果WordWrap=True,则行高会被自动调整
            .WordWrap = False
            strFormat = GetColFormat(vsAdvice)
            Call ClearAdviceData
            .ScrollBars = flexScrollBarNone
            Set .DataSource = rsTmp
            .ScrollBars = flexScrollBarBoth
            If err.Number = 0 And gcnOracle.Errors.Count > 0 Then
                gcnOracle.Errors.Clear    '怪,绑定时固定有此错误
            End If
            Call SetColFormat(vsAdvice, strFormat)
            .TextMatrix(0, COL_皮试) = ""
            .TextMatrix(0, COL_警示) = ""    'Pass
            .TextMatrix(0, COL_开始时间) = "生效时间"
            .TextMatrix(0, COL_并) = ""
            '自动调整行高
            .WordWrap = True

            '处理每行医嘱
            i = .FixedRows
            Do While i <= .Rows - 1
                .Cell(flexcpData, i, COL_处方号) = CStr(.TextMatrix(i, COL_处方号))
                .Cell(flexcpData, i, COL_处方打印) = CStr(.TextMatrix(i, COL_处方打印))
                .Cell(flexcpFontUnderline, i, COL_处方打印, i, COL_处方打印) = True
                .Cell(flexcpData, i, COL_处方预览) = CStr(.TextMatrix(i, COL_处方预览))
                .Cell(flexcpFontUnderline, i, COL_处方预览, i, COL_处方预览) = True
                .Cell(flexcpData, i, COL_查阅状态) = Val(.TextMatrix(i, COL_查阅状态)) '报告查阅状态值
                
                '处理发送时间
                If .TextMatrix(i, col_发送时间) <> "" Then
                    .Cell(flexcpData, i, col_发送时间) = .TextMatrix(i, col_发送时间)
                    .TextMatrix(i, col_发送时间) = Format(.TextMatrix(i, col_发送时间), "yyyy-MM-dd HH:mm")
                End If
                .Cell(flexcpData, i, COL_开始时间) = CStr(.TextMatrix(i, COL_开始时间))
                
                If .TextMatrix(i, COL_诊疗类别) = "K" And gbln血库系统 Then
                    strSQL = "select zl_Get_输血执行血型([1]) as 血型 from dual"
                    Set rs血型 = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(.TextMatrix(i, COL_ID)))
                    If Not rs血型.EOF Then
                        If rs血型!血型 & "" <> "" Then .TextMatrix(i, COL_皮试) = "(" & rs血型!血型 & ")"
                    End If
                End If

                '成药及中药的一些处理
                bln给药途径 = False: bln中药用法 = False: bln采集方法 = False: bln输血途径 = False
                If .TextMatrix(i, COL_诊疗类别) = "E" Then
                    If Val(.TextMatrix(i - 1, COL_相关ID)) = Val(.TextMatrix(i, COL_ID)) Then
                        If InStr(",5,6,", .TextMatrix(i - 1, COL_诊疗类别)) > 0 Then
                            bln给药途径 = True
                            For j = i - 1 To .FixedRows Step -1
                                If Val(.TextMatrix(j, COL_相关ID)) = Val(.TextMatrix(i, COL_ID)) Then
                                    '显示成药的给药途径+滴速
                                    .TextMatrix(j, COL_用法) = .TextMatrix(i, COL_用法) & .TextMatrix(i, COL_医生嘱托)

                                    If mvarCond.显示模式 = 0 Then    '合并用法列:用法 频率 天数
                                        strFormat = .TextMatrix(j, COL_用法)
                                        strTmp = .TextMatrix(j, COL_频率)
                                        If strTmp <> "" Then strFormat = strFormat & IIF(strFormat <> "", ",", "") & strTmp

                                        strTmp = .TextMatrix(j, COL_天数)
                                        If strTmp <> "" Then
                                            strFormat = strFormat & IIF(strFormat <> "", ",", "") & "共" & strTmp & "天"
                                        End If
                                        .TextMatrix(j, COL_用法) = strFormat
                                    End If

                                    '显示成药的执行性质
                                    If Val(.TextMatrix(j, COL_执行性质)) = 5 And Val(.TextMatrix(i, COL_执行性质)) <> 5 Then
                                        .TextMatrix(j, COL_执行性质) = "自备药"
                                    ElseIf Val(.TextMatrix(j, COL_执行性质)) <> 5 And Val(.TextMatrix(i, COL_执行性质)) = 5 Then
                                        .TextMatrix(j, COL_执行性质) = "离院带药"
                                    Else
                                        .TextMatrix(j, COL_执行性质) = ""
                                    End If
                                    
                                    '危急值ID是只关联在主医嘱主的，复制到药品行上
                                    .TextMatrix(j, COL_危急值ID) = .TextMatrix(i, COL_危急值ID)

                                    If mvarCond.显示模式 = 0 Then
                                        If .TextMatrix(j, COL_皮试) <> "" Then
                                            If Not (.TextMatrix(i, COL_诊疗类别) = "E" And .TextMatrix(i, COL_操作类型) = "1") Then
                                                .TextMatrix(j, col_内容) = .TextMatrix(j, col_内容) & "," & .TextMatrix(j, COL_皮试)
                                            End If
                                        End If
                                    End If
                                Else
                                    Exit For
                                End If
                            Next
                        ElseIf InStr(",7,C,", .TextMatrix(i - 1, COL_诊疗类别)) > 0 Then
                            bln中药用法 = .TextMatrix(i - 1, COL_诊疗类别) = "7"    '中药用法行
                            bln采集方法 = .TextMatrix(i - 1, COL_诊疗类别) = "C"    '采集方法行
                            
                            If bln中药用法 Then
                                .TextMatrix(i, COL_处方号) = .TextMatrix(i - 1, COL_处方号)
                                .TextMatrix(i, COL_处方打印) = .TextMatrix(i - 1, COL_处方打印)
                                .Cell(flexcpData, i, COL_处方号) = CStr(.TextMatrix(i - 1, COL_处方号))
                                .Cell(flexcpData, i, COL_处方打印) = CStr(.TextMatrix(i - 1, COL_处方打印))
                                .Cell(flexcpFontUnderline, i, COL_处方打印, i, COL_处方打印) = True
                                .Cell(flexcpData, i, COL_处方预览) = CStr(.TextMatrix(i - 1, COL_处方预览))
                                .Cell(flexcpFontUnderline, i, COL_处方预览, i, COL_处方预览) = True
                            End If

                            '采集方式的管码与一并的第一个检验相同
                            If bln采集方法 Then
                                j = .FindRow(.TextMatrix(i, COL_ID), .FixedRows, COL_相关ID)
                                If j <> -1 Then
                                    .TextMatrix(i, COL_试管编码) = .TextMatrix(j, COL_试管编码)
                                End If
                                .Cell(flexcpData, i, COL_皮试) = .TextMatrix(i, COL_皮试)
                                .TextMatrix(i, COL_皮试) = "" '耐受试验的时间ID界面上不显示
                            End If

                            '显示中药配方或检验组合的执行科室
                            .TextMatrix(i, COL_执行科室) = .TextMatrix(i - 1, COL_执行科室)

                            If bln中药用法 Then
                                '显示中药配方执行性质
                                If Val(.TextMatrix(i - 1, COL_执行性质)) = 5 And Val(.TextMatrix(i, COL_执行性质)) <> 5 Then
                                    .TextMatrix(i, COL_执行性质) = "自备药"
                                ElseIf Val(.TextMatrix(i - 1, COL_执行性质)) <> 5 And Val(.TextMatrix(i, COL_执行性质)) = 5 Then
                                    .TextMatrix(i, COL_执行性质) = "离院带药"
                                Else
                                    .TextMatrix(i, COL_执行性质) = ""
                                End If
                            Else
                                .TextMatrix(i, COL_执行性质) = ""
                            End If

                            '删除单味中药行,以及检验组合中的检验项目;同时判断检验申请
                            strTmp = ""
                            For j = i - 1 To .FixedRows Step -1
                                If Val(.TextMatrix(j, COL_相关ID)) = Val(.TextMatrix(i, COL_ID)) Then
                                    .TextMatrix(i, COL_报告项) = .TextMatrix(j, COL_报告项)    '检验、配方以首行医嘱为准
                                    .TextMatrix(i, COL_文件ID) = .TextMatrix(j, COL_文件ID)
                                    If bln中药用法 Then  '单味中药行ID记录下来，合理用药删除使用
                                        strTmp = strTmp & IIF(strTmp = "", .TextMatrix(j, COL_ID), "," & .TextMatrix(j, COL_ID))
                                    End If
                                    .RemoveItem j: i = i - 1
                                Else
                                    If bln中药用法 Then
                                        .Cell(flexcpData, i, COL_相关ID) = strTmp
                                    End If
                                    Exit For
                                End If
                            Next
                        End If
                    ElseIf .TextMatrix(i - 1, COL_诊疗类别) = "K" And Val(.TextMatrix(i - 1, COL_ID)) = Val(.TextMatrix(i, COL_相关ID)) Then
                        bln输血途径 = True
                        '显示输血途径
                        .TextMatrix(i - 1, COL_用法) = .TextMatrix(i, COL_用法) & .TextMatrix(i, COL_医生嘱托)
                    Else
                        .TextMatrix(i, COL_执行性质) = ""
                    End If
                End If

                '处理可见行的的一些标识:排开不可见但暂时未删除的行
                If Not (bln给药途径 Or bln输血途径) And .TextMatrix(i, COL_诊疗类别) <> "7" Then
                    
                    '行高：为了支持zl9PrintMode:Resize之后,取RowHeight可能小于RowHeightMin
                    If .RowHeight(i) < .RowHeightMin Then .RowHeight(i) = .RowHeightMin
                    
                    '只显示需的报告的医嘱
                    If mvarCond.过滤模式 = 3 And Val(.TextMatrix(i, COL_报告项)) = 0 Then
                        .RowHidden(i) = True: .RowHeight(i) = 0
                    End If
                    
                    '显示各种报告的医嘱
                    If mvarCond.过滤模式 = 3 Then
                        If mvarCond.报告 = 1 Then ' 检查
                            If Not .TextMatrix(i, COL_诊疗类别) = "D" Then
                                .RowHidden(i) = True: .RowHeight(i) = 0
                            End If
                        ElseIf mvarCond.报告 = 2 Then '检验
                            If Not (.TextMatrix(i, COL_诊疗类别) = "E" And .TextMatrix(i, COL_操作类型) = "6" Or .TextMatrix(i, COL_诊疗类别) = "C") Then
                                .RowHidden(i) = True: .RowHeight(i) = 0
                            End If
                        ElseIf mvarCond.报告 = 3 Then ' 其它
                            If .TextMatrix(i, COL_诊疗类别) = "E" And .TextMatrix(i, COL_操作类型) = "6" Or .TextMatrix(i, COL_诊疗类别) = "D" Or .TextMatrix(i, COL_诊疗类别) = "C" Then
                                .RowHidden(i) = True: .RowHeight(i) = 0
                            End If
                        End If
                    End If

                    '处理小数点问题,暂未想到办法
                    If Left(.TextMatrix(i, COL_总量), 1) = "." Then
                        .TextMatrix(i, COL_总量) = "0" & .TextMatrix(i, COL_总量)
                    End If
                    If Left(.TextMatrix(i, COL_单量), 1) = "." Then
                        .TextMatrix(i, COL_单量) = "0" & .TextMatrix(i, COL_单量)
                    End If
                    
                    '报告列及打印状态标识
                    Call SetAdviceReportIcon(i)

                    '医嘱颜色
                    If Val(.TextMatrix(i, COL_医嘱状态)) = 4 Then
                        '已作废(发送后作废)
                        .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H808080    '灰色
                    ElseIf Val(.TextMatrix(i, COL_医嘱状态)) = 8 Then
                        '已发送(发送后自动停止)
                        .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &HC00000    '深蓝
                    End If

                    '毒麻精药品标识:中药配方及组成味中药不处理
                    If .TextMatrix(i, COL_毒理分类) <> "" Then
                        If InStr(",麻醉药,毒性药,精神药,精神I类,精神II类,", .TextMatrix(i, COL_毒理分类)) > 0 Then
                            .Cell(flexcpFontBold, i, col_医嘱内容) = True
                            .Cell(flexcpFontBold, i, col_内容) = True
                        End If
                    End If

                    '皮试结果标识
                    If .TextMatrix(i, COL_诊疗类别) = "E" And .TextMatrix(i, COL_操作类型) = "1" And .TextMatrix(i, COL_皮试) <> "" Then
                        j = GetSkinTestResult(Val(.TextMatrix(i, COL_诊疗项目ID)), .TextMatrix(i, COL_皮试))
                        .Cell(flexcpForeColor, i, COL_皮试) = Decode(j, 1, vbRed, -1, vbBlue, .Cell(flexcpForeColor, i, COL_皮试))
                    End If

                    '紧急标志:一并给药只显示在第一行
                    blnFirst = True
                    If InStr(",5,6,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                        If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(i - 1, COL_相关ID)) Then
                            blnFirst = False
                        End If
                    End If
                    If blnFirst Then
                        If Val(.TextMatrix(i, COL_标志)) = 1 Then
                            Set .Cell(flexcpPicture, i, COL_F标志) = frmIcons.imgFlag.ListImages("紧急").Picture
                        ElseIf Val(.TextMatrix(i, COL_标志)) = 2 Then
                            Set .Cell(flexcpPicture, i, COL_F标志) = frmIcons.imgFlag.ListImages("补录").Picture
                        End If

                        '一并给药的，每行的审核状态单独设置
                        If Val(.TextMatrix(i, COL_医嘱状态)) < 2 Then   '新开或暂存的医嘱
                            Select Case Val(.TextMatrix(i, COL_审核状态))
                                '0-无需审核，1-待审核，2-审核通过，3-审核未通过
                            Case 1
                                If .TextMatrix(i, COL_诊疗类别) = "K" And Val(.TextMatrix(i, COL_检查方法)) = 1 Then
                                    '用血医嘱审核图标单独显示(表明是有医生核对)
                                    Set .Cell(flexcpPicture, i, COL_F标志) = frmIcons.imgFlag.ListImages("核对").Picture
                                Else
                                    Set .Cell(flexcpPicture, i, COL_F标志) = frmIcons.imgFlag.ListImages("待审核").Picture
                                End If
                            Case 2
                                If Not (.TextMatrix(i, COL_诊疗类别) = "K" And Val(.TextMatrix(i, COL_检查方法)) = 1) Then
                                    Set .Cell(flexcpPicture, i, COL_F标志) = frmIcons.imgFlag.ListImages("审核通过").Picture
                                End If
                            Case 3
                                Set .Cell(flexcpPicture, i, COL_F标志) = frmIcons.imgFlag.ListImages("审核未通过").Picture
                            Case 4, 5
                                If gbln血库系统 = False Then Set .Cell(flexcpPicture, i, COL_F标志) = frmIcons.imgFlag.ListImages("待审核").Picture
                            Case 7
                                Set .Cell(flexcpPicture, i, COL_F标志) = frmIcons.imgFlag.ListImages("待签发").Picture
                            Case Else
                            End Select
                            .Cell(flexcpPictureAlignment, i, COL_F标志) = 4
                        End If
                        '处方审查系统
                        If .TextMatrix(i, COL_处方审查状态) = "0" Then
                            Set .Cell(flexcpPicture, i, COL_F标志) = frmIcons.imgFlag.ListImages("待审核").Picture
                        ElseIf .TextMatrix(i, COL_处方审查状态) = "2" Or .TextMatrix(i, COL_处方审查结果) = "1" Then
                            '超时免审当作合格处理
                            Set .Cell(flexcpPicture, i, COL_F标志) = frmIcons.imgFlag.ListImages("审核通过").Picture
                        ElseIf .TextMatrix(i, COL_处方审查结果) = "2" Then
                            ' 不合格
                            Set .Cell(flexcpPicture, i, COL_F标志) = frmIcons.imgFlag.ListImages("审核未通过").Picture
                        End If
                    End If


                    'Pass:根据审查结果显示警示灯
                    '
                    If mblnPass Then
                        If .TextMatrix(i, COL_警示) <> "" Then
                            Call gobjPass.zlPassSetWarnLight(mobjPassMap, i, Val(.TextMatrix(i, COL_警示)))
                        End If
                    End If
                    .TextMatrix(i, COL_警示) = ""  '清空显示值
                End If


                If bln给药途径 Or bln输血途径 Then
                    .RemoveItem i
                Else
                    '简洁模式，组合医嘱内容
                    If mvarCond.显示模式 = 0 Then
                        strFormat = .TextMatrix(i, col_医嘱内容)
                        If .TextMatrix(i, COL_诊疗类别) <> "Z" And Val(.TextMatrix(i, COL_诊疗项目ID)) <> 0 Then
                            '医嘱内容定义中包含了相关项时，不再重复组合
                            mrsDefine.Filter = "诊疗类别='" & .TextMatrix(i, COL_诊疗类别) & "'"
                            If Not (.TextMatrix(i, COL_诊疗类别) = "E" And .TextMatrix(i, COL_操作类型) = "1") Then
                                strFormat = strFormat & .TextMatrix(i, COL_皮试)
                            End If

                            If Not (InStr("5,6,7", .TextMatrix(i, COL_诊疗类别)) = 0 And .TextMatrix(i, COL_频率) = "一次性") Then
                                blnDo = True
                                If mrsDefine.RecordCount > 0 Then blnDo = InStr(mrsDefine!医嘱内容, "[总量]") = 0
                                If blnDo Then
                                    strTmp = .TextMatrix(i, COL_总量)
                                    If strTmp <> "" Then strFormat = strFormat & ",共" & strTmp
                                End If

                                blnDo = True
                                If mrsDefine.RecordCount > 0 Then blnDo = InStr(mrsDefine!医嘱内容, "[单量]") = 0
                                If blnDo Then
                                    strTmp = .TextMatrix(i, COL_单量)
                                    If strTmp <> "" Then strFormat = strFormat & ",每次" & strTmp
                                End If
                            End If
                        End If
                        .TextMatrix(i, col_内容) = strFormat

                        '合并用法列:用法 频率 天数(一并给药的在前面已处理)
                        
                        '简洁模式下除药品、手术项目外其他的医嘱不显示用法
                        If .TextMatrix(i, COL_诊疗类别) = "F" And Val(.TextMatrix(i, COL_相关ID)) = 0 Or _
                            InStr(",5,6,7,", "," & .TextMatrix(i, COL_诊疗类别) & ",") > 0 And Val(.TextMatrix(i, COL_相关ID)) <> 0 Then
                            strFormat = .TextMatrix(i, COL_用法)
                        Else
                            strFormat = ""
                        End If
                        
                        '检验 '检查 '输血 '手术   简洁模式下不显示频率
                        If .TextMatrix(i, COL_诊疗类别) = "E" And Val(.TextMatrix(i, COL_操作类型)) = 6 Or _
                            .TextMatrix(i, COL_诊疗类别) = "D" And Val(.TextMatrix(i, COL_相关ID)) = 0 Or _
                            .TextMatrix(i, COL_诊疗类别) = "K" And Val(.TextMatrix(i, COL_相关ID)) = 0 Or _
                            .TextMatrix(i, COL_诊疗类别) = "F" And Val(.TextMatrix(i, COL_相关ID)) = 0 Then
                            strTmp = ""
                        Else
                            strTmp = .TextMatrix(i, COL_频率)
                        End If
                         
                        If strTmp <> "" Then strFormat = strFormat & IIF(strFormat <> "", ",", "") & strTmp

                        strTmp = .TextMatrix(i, COL_天数)
                        If strTmp <> "" Then
                            strFormat = strFormat & IIF(strFormat <> "", ",", "") & "共" & strTmp & "天"
                        End If

                        .TextMatrix(i, COL_用法) = strFormat
                    End If
                    
                    If mvarCond.过滤模式 = 3 Then
                        '如果是报告页签下，内容 列 可能为空，重新赋值
                        .TextMatrix(i, col_内容) = .TextMatrix(i, col_医嘱内容)
                        
                        If Val(.TextMatrix(i, COL_报告ID)) = 0 And .TextMatrix(i, COL_检查报告ID) = "" And Val(.TextMatrix(i, COL_RIS报告ID)) = 0 And Val(.TextMatrix(i, COL_LIS报告ID)) = 0 Then
                            .TextMatrix(i, COL_查阅状态) = "未出"
                        Else
                            .TextMatrix(i, COL_查阅状态) = "查阅"
                            If Val(.Cell(flexcpData, i, COL_查阅状态)) = 0 Then  '未读
                                .Cell(flexcpForeColor, i, COL_查阅状态, i, COL_查阅状态) = &HFF0000     '蓝色
                            ElseIf Val(.Cell(flexcpData, i, COL_查阅状态)) = 2 Then  '部分已读
                                .Cell(flexcpForeColor, i, COL_查阅状态, i, COL_查阅状态) = &HFF00FF     '紫色
                            Else
                                .Cell(flexcpForeColor, i, COL_查阅状态, i, COL_查阅状态) = &H80&     '暗红
                            End If
                            .Cell(flexcpFontUnderline, i, COL_查阅状态, i, COL_查阅状态) = True
                        End If
                        '增加过滤未出的报告和已出的报告
                        If .RowHidden(i) = False Then
                            If Not IIF(.TextMatrix(i, COL_查阅状态) = "未出", mvarCond.未出报告, mvarCond.已出报告) Then
                                .RowHidden(i) = True: .RowHeight(i) = 0
                            End If
                        End If
                    End If
                    i = i + 1
                End If
            Loop
            
            '设置医嘱内容单元格的图标
            For i = 1 To .Rows - 1
                Call SetAdviceIcon(i)
            Next
            
            '自动调整行高
            If mvarCond.显示模式 = 0 And mvarCond.过滤模式 <> 3 Then
                If InStr("2505,3345,1005,1335", .ColWidth(COL_用法)) > 0 Then .ColWidth(COL_用法) = IIF(mlngFontSize = 9, 2505, 3345)   '用户未改该列宽时才设置
                .AutoSize col_内容, COL_用法
                .ColWidth(COL_开始时间) = IIF(mlngFontSize = 9, 1130, 1510)
            Else
                If InStr("2505,3345,1005,1335", .ColWidth(COL_用法)) > 0 Then .ColWidth(COL_用法) = IIF(mlngFontSize = 9, 1005, 1335)
                .AutoSize col_医嘱内容, COL_用法
                .ColWidth(COL_开始时间) = IIF(mlngFontSize = 9, 1530, 2040)
            End If

            '固定列图标对齐:设置为中对齐,不然擦边框时可能有问题
            .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, .FixedCols - 1) = 4
            '电子签名图标对齐
            .Cell(flexcpPictureAlignment, .FixedRows, col_医嘱内容, .Rows - 1, col_医嘱内容) = 0
            Call SetTag一并给药
            Call Set标本状态
            .Redraw = True
        End With
    Else
        Call ClearAdviceData
        Call ClearAppendData
    End If
    imgColSel.Visible = (mvarCond.显示模式 = 1 And mvarCond.过滤模式 = 0)
    
    If mvarCond.过滤模式 <> 0 Then
        Call Refresh报告
    Else
        Call Refresh处方
    End If
 
    Screen.MousePointer = 0
    LoadAdvice = True
    Exit Function
errH:
    vsAdvice.Redraw = True
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function RowIs配方行(ByVal lngRow As Long) As Boolean
'功能：判断指定行是否中药配方行
'说明：指定行为显示行,且类别="E"
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
        
    On Error GoTo errH
    
    strSQL = "Select ID From 病人医嘱记录 Where Rownum=1 And 诊疗类别='7' And 相关ID=[1]"
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)))
    If Not rsTmp.EOF Then RowIs配方行 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function RowIs检验行(ByVal lngRow As Long) As Boolean
'功能：判断指定行是否检验组合行
'说明：指定行为显示行,且类别="E"
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
        
    On Error GoTo errH
    
    strSQL = "Select ID From 病人医嘱记录 Where Rownum=1 And 诊疗类别='C' And 相关ID=[1]"
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)))
    If Not rsTmp.EOF Then RowIs检验行 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ShowPrice(ByVal lngRow As Long) As Boolean
'功能：读取指定医嘱的计价,并根据当前的诊疗收费 关系进行更新
    Dim rs诊疗项目 As New ADODB.Recordset
    Dim rs收费细目 As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim str医嘱IDs As String, str收费细目IDs As String, str诊疗收费 As String
    Dim strSQL As String, i As Long, j As Long
    Dim bln配方行 As Boolean, bln检验行 As Boolean, blnLoad As Boolean
    Dim lng病人科室ID As Long, lng执行科室ID As Long
    Dim dblPrice As Double, lng材料ID As Long
    Dim lng医嘱ID As Long, lng相关ID As Long
    Dim strPriceType As String
    
    On Error GoTo errH
    
    With vsAppend
        .Redraw = False
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
    
        If Val(vsAdvice.TextMatrix(lngRow, COL_ID)) = 0 Then
            .Redraw = True: ShowPrice = True: Exit Function
        End If
        If vsAdvice.TextMatrix(lngRow, COL_诊疗类别) = "E" Then
            bln配方行 = RowIs配方行(lngRow)
            bln检验行 = RowIs检验行(lngRow)
        End If
        
        lng医嘱ID = Val(vsAdvice.TextMatrix(lngRow, COL_ID))
        lng相关ID = Val(vsAdvice.TextMatrix(lngRow, COL_相关ID))
                                    
        blnLoad = True
        
        '药品、卫材的计价
        If vsAdvice.TextMatrix(lngRow, COL_诊疗类别) = "4" Then
            '卫材的计价
            strSQL = "Select A.ID,A.相关ID,A.序号,A.诊疗类别,A.诊疗项目ID,NULL as 标本部位,NULL as 检查方法,0 as 执行标记,0 as 费用性质,0 as 收费方式," & _
                " A.收费细目ID,1 as 门诊包装,C.计算单位,1 as 数量,Decode(Nvl(C.是否变价,0),1,Nvl(B.单价,D.缺省价格),D.现价) as 单价,A.执行科室ID,0 as 从项" & _
                " From 病人医嘱记录 A,病人医嘱计价 B,收费项目目录 C,收费价目 D" & _
                " Where Rownum=1 And A.ID=[1] And A.ID=B.医嘱ID(+) And A.收费细目ID=C.ID And Nvl(A.执行性质,0) Not IN(0,5)" & _
                " And (C.撤档时间 is NULL Or C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "C", "D", "4", "5", "6") & _
                " And C.服务对象 IN(1,3) And D.收费细目ID=C.ID" & _
                " And ((Sysdate Between D.执行日期 and D.终止日期) or (Sysdate>=D.执行日期 And D.终止日期 is NULL))"
                
                blnLoad = False
        ElseIf InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
            '中,西成药:可能按规格下医嘱,计算门诊包装的单价
            strSQL = "Select A.ID,A.相关ID,A.序号,A.诊疗类别,A.诊疗项目ID,NULL as 标本部位,NULL as 检查方法,0 as 执行标记,0 as 费用性质,0 as 收费方式," & _
                " C.ID as 收费细目ID,B.门诊包装,B.门诊单位 as 计算单位,A.总给予量 as 数量,Decode(Nvl(C.是否变价,0),1,-NULL,D.现价)*B.门诊包装 as 单价," & _
                " A.执行科室ID,0 as 从项" & _
                " From 病人医嘱记录 A,药品规格 B,收费项目目录 C,收费价目 D" & _
                " Where Rownum=1 And A.ID=[1]" & _
                " And A.诊疗项目ID=B.药名ID And B.药品ID=C.ID And Nvl(A.执行性质,0) Not IN(0,5)" & _
                GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "C", "D", "4", "5", "6") & _
                " And (A.收费细目ID is NULL Or A.收费细目ID=B.药品ID)" & _
                " And (C.撤档时间 is NULL Or C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And C.服务对象 IN(1,3) And D.收费细目ID=C.ID" & _
                " And ((Sysdate Between D.执行日期 and D.终止日期) or (Sysdate>=D.执行日期 And D.终止日期 is NULL))"
                
                '仅一并给药(如果是)的第一成药行才显示给药途径的计价
                blnLoad = Val(vsAdvice.TextMatrix(lngRow - 1, COL_相关ID)) <> Val(vsAdvice.TextMatrix(lngRow, COL_相关ID))
        ElseIf bln配方行 Then
            '中草药:一定对应有规格记录且填写了收费细目ID
            strSQL = "Select A.ID,A.相关ID,A.序号,A.诊疗类别,A.诊疗项目ID,NULL as 标本部位,NULL as 检查方法,0 as 执行标记,0 as 费用性质,0 as 收费方式," & _
                " C.ID as 收费细目ID,B.门诊包装,B.门诊单位 as 计算单位,1 as 数量,Decode(Nvl(C.是否变价,0),1,-NULL,D.现价)*B.门诊包装 as 单价," & _
                " A.执行科室ID,0 as 从项" & _
                " From 病人医嘱记录 A,药品规格 B,收费项目目录 C,收费价目 D" & _
                " Where A.诊疗类别='7' And A.相关ID=[1]" & _
                " And A.收费细目ID=B.药品ID And A.收费细目ID=C.ID And C.服务对象 IN(1,3)" & _
                GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "C", "D", "4", "5", "6") & _
                " And D.收费细目ID=C.ID And Nvl(A.执行性质,0) Not IN(0,5)" & _
                " And (C.撤档时间 is NULL Or C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And ((Sysdate Between D.执行日期 and D.终止日期) or (Sysdate>=D.执行日期 And D.终止日期 is NULL))"
        End If
        
        '读取现有计价(取最新价格)：除药品、卫材外的计价,包含相关医嘱计价
        '不计价,手工计价的医嘱不读取
        '用Union方式可以利用索引
        If blnLoad Then
            '不是新开的医嘱，根据病人医嘱计价提取
            If InStr(",1,2,-1,", vsAdvice.TextMatrix(lngRow, COL_医嘱状态)) = 0 Then
                strSQL = strSQL & IIF(strSQL = "", "", " Union ALL") & _
                    " Select A.ID,A.相关ID,A.序号,A.诊疗类别,A.诊疗项目ID,A.标本部位,A.检查方法,A.执行标记,Nvl(B.费用性质,0) as 费用性质,Nvl(B.收费方式,0) as 收费方式," & _
                    " B.收费细目ID,1 as 门诊包装,C.计算单位,B.数量,Decode(C.是否变价,1,B.单价,Sum(D.现价)) as 单价," & _
                    " Nvl(B.执行科室ID,A.执行科室ID) as 执行科室ID,Nvl(B.从项,0) as 从项" & _
                    " From 病人医嘱记录 A,病人医嘱计价 B,收费项目目录 C,收费价目 D" & _
                    " Where A.诊疗类别 Not IN('4','5','6','7') And A.ID=B.医嘱ID" & _
                    GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "C", "D", "4", "5", "6") & _
                    " And Nvl(A.计价特性,0)=0 And Nvl(A.执行性质,0) Not IN(0,5) And B.收费细目ID=C.ID And B.收费细目ID=D.收费细目ID" & _
                    " And ((Sysdate Between D.执行日期 and D.终止日期) or (Sysdate>=D.执行日期 And D.终止日期 is NULL))" & _
                    " And (A.ID=[1]" & IIF(lng相关ID <> 0, " Or A.ID=[2]", "") & " Or A.相关ID=[1])" & _
                    " Group by A.ID,A.相关ID,A.序号,A.诊疗类别,A.诊疗项目ID,A.标本部位,A.检查方法,A.执行标记,Nvl(B.费用性质,0),Nvl(B.收费方式,0)," & _
                    " B.收费细目ID,C.计算单位,B.数量,C.是否变价,B.单价,Nvl(B.执行科室ID,A.执行科室ID),Nvl(B.从项,0)"
            Else
                '新开的医嘱，根据诊疗收费 关系提取(非药变价显示为0)
                '几种对应的计价：
                '   1.加收的费用，只在主项目上面加收，目前只有床旁或术中这种情况
                '   2.基本的费用，但是具体的检查部位和检查方法的
                '   3.基本的费用，非检查部位和方法的(注意检验标本填写在标本部位中)
                lng材料ID = 0 '检验试管费用,只收取试管对应的卫材费用
                If vsAdvice.TextMatrix(lngRow, COL_试管编码) <> "" Then
                    lng材料ID = GetTubeMaterial(vsAdvice.TextMatrix(lngRow, COL_试管编码))
                End If
                
                str诊疗收费 = "Select * From (" & _
                    "Select C.诊疗项目ID,C.收费项目ID,C.检查部位,C.检查方法,C.费用性质,C.收费数量,C.固有对照,C.从属项目,C.收费方式,C.适用科室id" & _
                    " ,Max(Nvl(c.适用科室id, 0)) Over(Partition By c.诊疗项目id, c.检查部位, c.检查方法, c.费用性质) As Top" & _
                    " From 诊疗收费关系 C,病人医嘱记录 A Where (A.ID=[1]" & IIF(lng相关ID <> 0, " Or A.ID=[2]", "") & " Or A.相关ID=[1]) And A.诊疗项目ID+0=C.诊疗项目ID" & _
                    "   And (a.相关id Is Null And a.执行标记 In (1, 2) And c.费用性质 = 1 Or" & vbNewLine & _
                    "   a.标本部位 = c.检查部位 And a.检查方法 = c.检查方法 And Nvl(c.费用性质, 0) = 0 Or" & vbNewLine & _
                    "   (a.检查方法 Is Null or a.诊疗类别 = 'E' And Exists(Select 1 From 诊疗项目目录 Z Where Z.id=a.诊疗项目ID And Z.操作类型='4')) And Nvl(c.费用性质, 0) = 0 And c.检查部位 Is Null And c.检查方法 Is Null)" & _
                    "      And (C.适用科室ID is Null or C.适用科室ID = A.执行科室ID And C.病人来源 = 1)" & _
                    " ) Where Nvl(适用科室id, 0) = Top"
                
                strSQL = strSQL & IIF(strSQL = "", "", " Union ALL") & _
                    " Select A.ID,A.相关ID,A.序号,A.诊疗类别,A.诊疗项目ID,A.标本部位,A.检查方法,A.执行标记,Nvl(B.费用性质,0) as 费用性质,Nvl(B.收费方式,0) as 收费方式," & _
                    " B.收费项目ID as 收费细目ID,1 as 门诊包装,C.计算单位,B.收费数量 as 数量,Decode(C.是否变价,1,Sum(D.缺省价格),Sum(D.现价)) as 单价," & _
                    " A.执行科室ID,Nvl(B.从属项目,0) as 从项" & _
                    " From 病人医嘱记录 A,(" & str诊疗收费 & ") B,收费项目目录 C,收费价目 D" & _
                    " Where A.诊疗类别 Not IN('4','5','6','7') And A.医嘱状态 IN(-1,1,2) And A.诊疗项目ID+0=B.诊疗项目ID" & _
                    GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "C", "D", "4", "5", "6") & _
                    " And (A.相关ID is Null And A.执行标记 IN(1,2) And B.费用性质=1" & _
                    "       Or A.标本部位=B.检查部位 And A.检查方法=B.检查方法 And Nvl(B.费用性质,0)=0" & _
                    "       Or (A.检查方法 is Null or a.诊疗类别 = 'E' And Exists(Select 1 From 诊疗项目目录 Z Where Z.id=a.诊疗项目ID And Z.操作类型='4')) And Nvl(B.费用性质,0)=0 And B.检查部位 is Null And B.检查方法 is Null)" & _
                    " And Nvl(A.计价特性,0)=0 And Nvl(A.执行性质,0) Not IN(0,5) And B.收费项目ID=C.ID And B.收费项目ID=D.收费细目ID" & _
                    " And ((Sysdate Between D.执行日期 and D.终止日期) or (Sysdate>=D.执行日期 And D.终止日期 is NULL))" & _
                    " And (C.撤档时间 is NULL Or C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD')) And C.服务对象 IN(1,3)" & _
                    " And (Nvl(B.收费方式,0)=1 And C.类别='4' And B.收费项目ID=[3] Or Not(Nvl(B.收费方式,0)=1 And C.类别='4' And [3]<>0))" & _
                    " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null) And (A.ID=[1]" & IIF(lng相关ID <> 0, " Or A.ID=[2]", "") & " Or A.相关ID=[1])" & _
                    " Group by A.ID,A.相关ID,A.序号,A.诊疗类别,A.诊疗项目ID,A.标本部位,A.检查方法,A.执行标记,Nvl(B.费用性质,0),Nvl(B.收费方式,0)," & _
                    " B.收费项目ID,C.计算单位,B.收费数量,C.是否变价,A.执行科室ID,Nvl(B.从属项目,0)"
            End If
        End If
        strSQL = strSQL & " Order by 序号,费用性质,从项"
        
        If mblnMoved Then '挂号单与医嘱在同个数据库
            strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
            strSQL = Replace(strSQL, "病人医嘱计价", "H病人医嘱计价")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, lng医嘱ID, lng相关ID, lng材料ID, mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
        
        '显示计价内容
        If Not rsTmp.EOF Then
            '确定显示行数
            .Rows = .FixedRows + rsTmp.RecordCount
            
            '获取诊疗项目,收费细目信息
            For i = 1 To rsTmp.RecordCount
                If InStr("," & str医嘱IDs & ",", "," & rsTmp!ID & ",") = 0 Then str医嘱IDs = str医嘱IDs & "," & rsTmp!ID
                If InStr("," & str收费细目IDs & ",", "," & rsTmp!收费细目ID & ",") = 0 Then str收费细目IDs = str收费细目IDs & "," & rsTmp!收费细目ID
                rsTmp.MoveNext
            Next
            str医嘱IDs = Mid(str医嘱IDs, 2)
            str收费细目IDs = Mid(str收费细目IDs, 2)
                        
            strSQL = "Select/*+ Rule*/ B.ID,B.类别,C.名称 as 类别名称,B.名称,B.标本部位" & _
                " From 病人医嘱记录 A,诊疗项目目录 B,诊疗项目类别 C,Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)) D" & _
                " Where A.ID = D.Column_Value And A.诊疗项目ID=B.ID And B.类别=C.编码"
                
            If mblnMoved Then '挂号单与医嘱在同个数据库
                strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
            End If
            Set rs诊疗项目 = zlDatabase.OpenSQLRecord(strSQL, Me.Name, str医嘱IDs) 'In
            
            strSQL = "Select A.ID,A.类别,B.名称 as 类别名称,A.编码," & _
                " A.名称,A.规格,A.产地,A.费用类型,A.是否变价" & _
                " From 收费项目目录 A,收费项目类别 B,Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)) D" & _
                " Where A.类别=B.编码 And A.ID = D.Column_Value"
            strSQL = "Select/*+ Rule*/ A.ID,A.类别,A.类别名称,A.编码,Nvl(B.名称,A.名称) as 名称," & _
                " A.规格,A.产地,A.费用类型,A.是否变价,C.跟踪在用" & _
                " From (" & strSQL & ") A,收费项目别名 B,材料特性 C" & _
                " Where A.ID=C.材料ID(+) And A.ID=B.收费细目ID(+) And B.码类(+)=1 And B.性质(+)=[2]"
            Set rs收费细目 = zlDatabase.OpenSQLRecord(strSQL, Me.Name, str收费细目IDs, IIF(gbyt药品名称显示 = 0, 1, 3))
            
            '显示每行内容
            rsTmp.MoveFirst
            For i = 1 To rsTmp.RecordCount
                rs诊疗项目.Filter = "ID=" & rsTmp!诊疗项目ID
                rs收费细目.Filter = "ID=" & rsTmp!收费细目ID
                
                '计价医嘱
                If rsTmp!诊疗类别 = "4" Then
                    .TextMatrix(i, COLPrice("计价医嘱")) = "卫生材料-" & rs诊疗项目!名称
                ElseIf InStr(",5,6,7,", rsTmp!诊疗类别) > 0 Then
                    .TextMatrix(i, COLPrice("计价医嘱")) = "药品医嘱-" & rs诊疗项目!名称
                ElseIf rsTmp!诊疗类别 = "E" And InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
                    .TextMatrix(i, COLPrice("计价医嘱")) = "给药途径-" & rs诊疗项目!名称
                ElseIf rsTmp!诊疗类别 = "E" And vsAdvice.TextMatrix(lngRow, COL_诊疗类别) = "K" Then
                    .TextMatrix(i, COLPrice("计价医嘱")) = "输血途径-" & rs诊疗项目!名称
                ElseIf rsTmp!诊疗类别 = "E" And (bln配方行 Or bln检验行) Then
                    If bln检验行 Then
                        .TextMatrix(i, COLPrice("计价医嘱")) = "采集方法-" & rs诊疗项目!名称
                    ElseIf Not IsNull(rsTmp!相关ID) Then
                        .TextMatrix(i, COLPrice("计价医嘱")) = "中药煎法-" & rs诊疗项目!名称
                    Else
                        .TextMatrix(i, COLPrice("计价医嘱")) = "中药用法-" & rs诊疗项目!名称
                    End If
                ElseIf Not IsNull(rsTmp!相关ID) Then
                    If rsTmp!诊疗类别 = "C" Then
                        .TextMatrix(i, COLPrice("计价医嘱")) = "检验项目-" & rs诊疗项目!名称
                    ElseIf rsTmp!诊疗类别 = "D" Then
                        '部位及方法
                        .TextMatrix(i, COLPrice("计价医嘱")) = "检查部位-" & NVL(rsTmp!标本部位) & "(" & NVL(rsTmp!检查方法) & ")"
                    ElseIf rsTmp!诊疗类别 = "F" Then
                        .TextMatrix(i, COLPrice("计价医嘱")) = "附加手术-" & rs诊疗项目!名称
                    ElseIf rsTmp!诊疗类别 = "G" Then
                        .TextMatrix(i, COLPrice("计价医嘱")) = "麻醉项目-" & rs诊疗项目!名称
                    End If
                Else
                    If NVL(rsTmp!费用性质, 0) = 1 Then
                        '床旁或术中加收费用
                        .TextMatrix(i, COLPrice("计价医嘱")) = rs诊疗项目!类别名称 & "医嘱-" & rs诊疗项目!名称 & "(" & Decode(NVL(rsTmp!执行标记, 0), 1, "床旁", 2, "术中", "") & "加收)"
                    Else
                        .TextMatrix(i, COLPrice("计价医嘱")) = rs诊疗项目!类别名称 & "医嘱-" & rs诊疗项目!名称
                    End If
                End If
                
                '类别
                .TextMatrix(i, COLPrice("类别")) = rs收费细目!类别名称
                '收费项目:规格/产地
                .TextMatrix(i, COLPrice("收费项目")) = rs收费细目!名称
                If Not IsNull(rs收费细目!产地) Then
                    .TextMatrix(i, COLPrice("收费项目")) = .TextMatrix(i, COLPrice("收费项目")) & "(" & rs收费细目!产地 & ")"
                End If
                If Not IsNull(rs收费细目!规格) Then
                    .TextMatrix(i, COLPrice("收费项目")) = .TextMatrix(i, COLPrice("收费项目")) & " " & rs收费细目!规格
                End If
                
                '计算单位:药嘱药品为门诊单位,非药嘱药品为售价单位
                .TextMatrix(i, COLPrice("单位")) = NVL(rsTmp!计算单位)
                '计价数量:药嘱药品为1,非药嘱药品为对应售价数
                If InStr(",5,6,7,", rs诊疗项目!类别) > 0 Then
                    .TextMatrix(i, COLPrice("计价数量")) = 1
                Else
                    .TextMatrix(i, COLPrice("计价数量")) = FormatEx(rsTmp!数量, 5)
                End If
                '执行科室
                lng执行科室ID = NVL(rsTmp!执行科室ID, 0)
                If rs收费细目!类别 = "4" And NVL(rs收费细目!跟踪在用, 0) = 1 _
                    Or InStr(",5,6,7,", rs收费细目!类别) > 0 And InStr(",5,6,7,", rs诊疗项目!类别) = 0 Then
                    lng病人科室ID = mlng挂号科室ID
                    lng执行科室ID = Get收费执行科室ID(mlng病人ID, 0, rs收费细目!类别, rs收费细目!ID, 4, lng病人科室ID, 0, 1, lng执行科室ID)
                End If
                
                '单价处理
                If InStr(",5,6,7,", rs收费细目!类别) > 0 Then
                    If NVL(rs收费细目!是否变价, 0) = 1 Then
                        '求药品时价
                        If InStr(",5,6,7,", rs诊疗项目!类别) > 0 Then
                            '药嘱药品计算门诊包装的门诊时价
                            .TextMatrix(i, COLPrice("单价")) = CalcDrugPrice(rs收费细目!ID, lng执行科室ID, NVL(rsTmp!数量, 1), , , 1, mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
                            .TextMatrix(i, COLPrice("单价")) = Format(Val(.TextMatrix(i, COLPrice("单价"))) * NVL(rsTmp!门诊包装, 0), gstrDecPrice)
                        Else
                            '非药嘱药品计算相对售价数量的售价实价
                            .TextMatrix(i, COLPrice("单价")) = Format(CalcDrugPrice(rs收费细目!ID, lng执行科室ID, NVL(rsTmp!数量, 0), , , 1, mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级), gstrDecPrice)
                        End If
                    Else
                        '药嘱药品为门诊单价,非药药品为售价
                        .TextMatrix(i, COLPrice("单价")) = Format(NVL(rsTmp!单价), gstrDecPrice)
                    End If
                ElseIf rs收费细目!类别 = "4" And NVL(rs收费细目!跟踪在用, 0) = 1 And NVL(rs收费细目!是否变价, 0) = 1 Then
                    '时价卫材的单价和药品一样计算
                    .TextMatrix(i, COLPrice("单价")) = Format(CalcDrugPrice(rs收费细目!ID, lng执行科室ID, NVL(rsTmp!数量, 0), , , 1, mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级), gstrDecPrice)
                Else
                    .TextMatrix(i, COLPrice("单价")) = Format(NVL(rsTmp!单价), gstrDecPrice)
                End If
                
                '执行科室
                If lng执行科室ID <> 0 Then
                    .TextMatrix(i, COLPrice("执行科室")) = Sys.RowValue("部门表", lng执行科室ID, "名称")
                End If
                
                '显示医保费用类型
                If Val(rsTmp!收费细目ID & "") <> 0 Then
                    strPriceType = GetPriceType(mlng病人ID, Val(rsTmp!收费细目ID & ""), mint险类, True)
                End If
                '费用类型
                If strPriceType = "" Then
                    .TextMatrix(i, COLPrice("费用类型")) = NVL(rs收费细目!费用类型)
                Else
                    .TextMatrix(i, COLPrice("费用类型")) = strPriceType
                End If
                
                '从属项目
                .TextMatrix(i, COLPrice("从项")) = IIF(NVL(rsTmp!从项, 0) = 0, "", "√")
                
                '收费方式
                .TextMatrix(i, COLPrice("收费方式")) = getChargeMode(Val(NVL(rsTmp!收费方式, 0)))
                
                dblPrice = dblPrice + Format(Val(.TextMatrix(i, COLPrice("计价数量"))) * Val(.TextMatrix(i, COLPrice("单价"))), "0.00000")
                
                rsTmp.MoveNext
            Next
        End If
        
        '合计行
        If .Rows > 2 Then
            .MergeCol(COLPrice("计价医嘱")) = True
            .MergeCol(COLPrice("类别")) = True
            
            .Rows = .Rows + 1
            .Cell(flexcpText, .Rows - 1, COLPrice("计价医嘱"), .Rows - 1, COLPrice("单位")) = "合计"
            .Cell(flexcpAlignment, .Rows - 1, COLPrice("计价医嘱"), .Rows - 1, COLPrice("单位")) = 4
            .Cell(flexcpText, .Rows - 1, COLPrice("计价数量"), .Rows - 1, COLPrice("单价")) = Format(dblPrice, gstrDecPrice)
            .Cell(flexcpAlignment, .Rows - 1, COLPrice("计价数量"), .Rows - 1, COLPrice("单价")) = 7
            .MergeCells = flexMergeRestrictRows
            .MergeRow(.Rows - 1) = True
            
        End If
        
        .Row = 1: .Col = 0
        .Redraw = True
        Call vsAppend_AfterRowColChange(-1, -1, .Row, .Col)
    End With
    
    ShowPrice = True
    Exit Function
errH:
    vsAppend.Redraw = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ShowSendList(ByVal lngRow As Long) As Boolean
'功能：显示指定行医嘱的发送记录
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strSub As String, i As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim strExe1 As String, strExe2 As String, strState As String
    Dim bln配方行 As Boolean, bln检验行 As Boolean
    Dim bln状态说明 As Boolean
    Dim lng输血 As Long
    Dim j As Long
    
    On Error GoTo errH
    lng输血 = -1
    With vsAppend
        .Redraw = False
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
    
        If Val(vsAdvice.TextMatrix(lngRow, COL_ID)) = 0 Then
            .Redraw = True: ShowSendList = True: Exit Function
        End If
        
        If vsAdvice.TextMatrix(lngRow, COL_诊疗类别) = "E" Then
            bln配方行 = RowIs配方行(lngRow)
            bln检验行 = RowIs检验行(lngRow)
        End If
                
        strExe1 = "Decode(Nvl(A.执行状态,0),0,'未执行',1,'完全执行',2,'部份执行')"
        strExe2 = "Decode(Nvl(B.执行状态,0),0,'未执行',1,'执行完成',2,'拒绝执行',3,'正在执行')"
        strState = "Decode(A.执行状态,9,'收费异常',Decode(A.记录性质,1,Decode(A.记录状态,0,'收费划价',1,'已收费',3,'已退费'),2,Decode(A.记录状态,0,'记帐划价',1,'已记帐',3,'已销帐'),'未计费'))"
        
        '药嘱对应的药品计价按门诊包装显示,非药嘱对应的药品计价按零售单位显示
        If InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
            If Not RowIn一并给药(lngRow, lngBegin, lngEnd) Then lngBegin = lngRow
            '成药部份:填写了发送记录,但可能无对应费用(如自备药,但医嘱有规格)
            strSub = "Select a.医嘱序号,MIN(a.记录性质) AS 记录性质 ,A.NO,A.执行状态,Min(A.记录状态) as 记录状态,A.序号,A.执行部门ID,a.收费类别,A.数次,A.付数, a.收费细目id,B.门诊包装,B.门诊单位" & _
                " From 门诊费用记录 A,药品规格 B" & _
                " Where A.记录状态 IN(0,1,3) And A.价格父号 is NULL And A.收费类别 IN('5','6','7')" & _
                " And A.收费细目ID=B.药品ID And A.医嘱序号=[1] Group By a.医嘱序号,A.NO,A.执行状态,A.序号,A.执行部门ID,a.收费类别,A.数次,A.付数, a.收费细目id,b.门诊包装, b.门诊单位"
            If mblnMoved Then
                strSub = Replace(strSub, "门诊费用记录", "H门诊费用记录")
            ElseIf zlDatabase.DateMoved(mvRegDate) Then
                strSub = strSub & " Union ALL " & Replace(strSub, "门诊费用记录", "H门诊费用记录")
            End If
            
            strSQL = _
                " Select C.相关ID,C.标本部位,C.检查方法,B.发送时间,B.NO,B.记录性质,A.收费细目ID," & _
                " Nvl(A.门诊单位,D.门诊单位) as 单位," & _
                " Nvl(A.数次/Nvl(A.门诊包装,1),B.发送数次/Nvl(D.剂量系数,1)/Nvl(D.门诊包装,1)) as 发送数次," & _
                " Nvl(A.执行部门ID,B.执行部门ID) as 执行部门ID," & _
                " Decode(Nvl(Instr(',4,5,6,7,',A.收费类别),0),0," & strExe2 & "," & strExe1 & ") as 执行状态,B.首次时间,B.末次时间," & _
                " Decode(Nvl(B.计费状态,0),-1,'无需计费',0,'未计费'," & strState & ") as 计费状态," & _
                " B.发送人,B.状态说明,B.发送号,B.记录序号 as 发送序号,A.序号 as 费用序号,C.诊疗项目ID,C.诊疗类别,B.完成时间,B.完成人,B.执行说明" & _
                " From (" & strSub & ") A,病人医嘱发送 B,病人医嘱记录 C,药品规格 D" & _
                " Where B.医嘱ID=C.ID And C.收费细目ID=D.药品ID And C.ID=[1]" & _
                " And A.NO(+)=B.NO And A.记录性质(+)=B.记录性质 And A.医嘱序号(+)=B.医嘱ID"
            
            '在一并给药的首行才显示给药途径的发送
            If lngRow = lngBegin Then
                '给药途径部份:填写了发送记录(叮嘱无),但不一定有费用
                strSub = "Select a.医嘱序号,MIN(a.记录性质) AS 记录性质 ,A.NO,A.执行状态,Min(A.记录状态) as 记录状态,A.序号,A.执行部门ID,a.收费类别,A.数次,A.付数, a.收费细目id,B.门诊包装,B.门诊单位" & _
                    " From 门诊费用记录 A,药品规格 B" & _
                    " Where A.记录状态 IN(0,1,3) And A.价格父号 is NULL" & _
                    " And A.收费细目ID=B.药品ID(+) And A.医嘱序号=[2] Group By a.医嘱序号,A.NO,A.执行状态,A.序号,A.执行部门ID,a.收费类别,A.数次,A.付数, a.收费细目id,b.门诊包装, b.门诊单位"
                If mblnMoved Then
                    strSub = Replace(strSub, "门诊费用记录", "H门诊费用记录")
                ElseIf zlDatabase.DateMoved(mvRegDate) Then
                    strSub = strSub & " Union ALL " & Replace(strSub, "门诊费用记录", "H门诊费用记录")
                End If
                    
                strSQL = strSQL & " Union ALL " & _
                    " Select C.相关ID,C.标本部位,C.检查方法,B.发送时间,B.NO,B.记录性质,A.收费细目ID," & _
                    " Decode(Nvl(Instr('567',A.收费类别),0),0,Decode(A.收费类别,'4',F.计算单位,D.计算单位),Nvl(A.门诊单位,E.门诊单位)) as 单位," & _
                    " Nvl(A.数次/Nvl(A.门诊包装,1),B.发送数次/Nvl(E.剂量系数,1)/Nvl(E.门诊包装,1)) as 发送数次," & _
                    " Nvl(A.执行部门ID,B.执行部门ID) as 执行部门ID," & _
                    " Decode(Nvl(Instr(',4,5,6,7,',A.收费类别),0),0," & strExe2 & "," & strExe1 & ") as 执行状态,B.首次时间," & _
                    " B.末次时间,Decode(Nvl(B.计费状态,0),-1,'无需计费',0,'未计费'," & strState & ") as 计费状态," & _
                    " B.发送人,B.状态说明,B.发送号,B.记录序号 as 发送序号,A.序号 as 费用序号,C.诊疗项目ID,C.诊疗类别,B.完成时间,B.完成人,B.执行说明" & _
                    " From (" & strSub & ") A,病人医嘱发送 B,病人医嘱记录 C,诊疗项目目录 D,药品规格 E,收费项目目录 F" & _
                    " Where B.医嘱ID=C.ID And C.诊疗项目ID=D.ID And C.收费细目ID=E.药品ID(+) And C.收费细目ID=F.ID(+)" & _
                    " And A.NO(+)=B.NO And A.记录性质(+)=B.记录性质 And 0+A.医嘱序号(+)=B.医嘱ID And C.ID=[2]"
            End If
            
            If mblnMoved Then
                strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
                strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
            End If
        Else
            '其它医嘱(包括卫材、配方及检查，手术一组医嘱):填写了发送记录(叮嘱无),但不一定有费用
            '中药自备药也是无对应费用(但医嘱有规格)
            strSub = _
                " Select a.医嘱序号,MIN(a.记录性质) AS 记录性质 ,A.NO,A.执行状态,Min(A.记录状态) as 记录状态,A.序号,A.执行部门ID,a.收费类别,A.数次,A.付数, a.收费细目id,B.门诊包装,B.门诊单位" & _
                " From 门诊费用记录 A,药品规格 B" & _
                " Where A.记录状态 IN(0,1,3) And A.价格父号 is NULL" & _
                " And A.收费细目ID=B.药品ID(+) And A.医嘱序号=[1]  Group By a.医嘱序号,A.NO,A.执行状态,A.序号,A.执行部门ID,a.收费类别,A.数次,A.付数, a.收费细目id,b.门诊包装, b.门诊单位"
            strSub = strSub & " Union ALL " & _
                " Select a.医嘱序号,MIN(a.记录性质) AS 记录性质 ,A.NO,A.执行状态,Min(A.记录状态) as 记录状态,A.序号,A.执行部门ID,a.收费类别,A.数次,A.付数, a.收费细目id,B.门诊包装,B.门诊单位" & _
                " From 门诊费用记录 A,药品规格 B,病人医嘱记录 C" & _
                " Where A.记录状态 IN(0,1,3) And A.价格父号 is NULL" & _
                " And A.收费细目ID=B.药品ID(+) And A.医嘱序号=C.ID" & _
                " And C.相关ID=[1]  Group By a.医嘱序号,A.NO,A.执行状态,A.序号,A.执行部门ID,a.收费类别,A.数次,A.付数, a.收费细目id,b.门诊包装, b.门诊单位"
            If mblnMoved Then
                strSub = Replace(strSub, "门诊费用记录", "H门诊费用记录")
            ElseIf zlDatabase.DateMoved(mvRegDate) Then
                strSub = strSub & " Union ALL " & Replace(strSub, "门诊费用记录", "H门诊费用记录")
            End If
            
            strSQL = _
                " Select * From 病人医嘱记录 Where ID=[1]" & _
                " Union ALL " & _
                " Select * From 病人医嘱记录 Where 相关ID=[1]"
            strSQL = _
                " Select C.相关ID,C.标本部位,C.检查方法,B.发送时间,B.NO,B.记录性质,A.收费细目ID," & _
                " Decode(Nvl(Instr('567',A.收费类别),0),0,Decode(A.收费类别,'4',F.计算单位,D.计算单位),Nvl(A.门诊单位,E.门诊单位)) as 单位," & _
                " Nvl(Nvl(A.付数,1)*A.数次/Nvl(A.门诊包装,1),B.发送数次/Nvl(E.剂量系数,1)/Nvl(E.门诊包装,1)) as 发送数次," & _
                " Nvl(A.执行部门ID,B.执行部门ID) as 执行部门ID," & _
                " Decode(Nvl(Instr(',4,5,6,7,',A.收费类别),0),0," & strExe2 & "," & strExe1 & ") as 执行状态,B.首次时间,B.末次时间," & _
                " Decode(Nvl(B.计费状态,0),-1,'无需计费',0,'未计费'," & strState & ") as 计费状态," & _
                " B.发送人,B.状态说明,B.发送号,B.记录序号 as 发送序号,A.序号 as 费用序号,C.诊疗项目ID,C.诊疗类别,B.完成时间,B.完成人,B.执行说明" & _
                " From (" & strSub & ") A,病人医嘱发送 B,(" & strSQL & ") C,诊疗项目目录 D,药品规格 E,收费项目目录 F" & _
                " Where B.医嘱ID=C.ID And C.诊疗项目ID=D.ID And C.收费细目ID=E.药品ID(+) And C.收费细目ID=F.ID(+)" & _
                " And A.NO(+)=B.NO And A.记录性质(+)=B.记录性质 And 0+A.医嘱序号(+)=B.医嘱ID"
            If mblnMoved Then
                strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
                strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
            End If
        End If
        
        strSQL = "Select /*+ RULE */ A.发送序号,A.费用序号," & _
            " A.相关ID,A.诊疗类别,F.名称 as 类别名称,D.名称 as 诊疗项目,A.标本部位,A.检查方法,A.发送时间,A.NO,A.记录性质," & _
            " Nvl(G.名称,B.名称)||Decode(B.产地,NULL,NULL,'('||B.产地||')')||Decode(B.规格,NULL,NULL,' '||B.规格) as 收费项目," & _
            " A.单位,A.发送数次 as 数量,C.名称 as 执行科室,A.执行状态,A.首次时间,A.末次时间,A.计费状态,A.发送人,A.状态说明,A.发送号,A.完成时间,A.完成人,A.执行说明" & _
            " From (" & strSQL & ") A,收费项目目录 B,部门表 C,诊疗项目目录 D,诊疗项目类别 F,收费项目别名 G" & _
            " Where A.收费细目ID=B.ID(+) And A.执行部门ID=C.ID(+)" & _
            " And A.诊疗项目ID=D.ID And A.诊疗类别=F.编码" & _
            " And A.收费细目ID=G.收费细目ID(+) And G.码类(+)=1 And G.性质(+)=" & IIF(gbyt药品名称显示 = 0, 1, 3) & _
            " Order by A.发送号 Desc,A.诊疗类别,A.发送序号,A.费用序号"
            
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(lngRow, COL_ID)), Val(vsAdvice.TextMatrix(lngRow, COL_相关ID)))
        If Not rsTmp.EOF Then
            .Rows = rsTmp.RecordCount + 1
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, COLSend("发送号")) = NVL(rsTmp!发送号, 0)
                .TextMatrix(i, COLSend("发送时间")) = Format(NVL(rsTmp!发送时间), "yyyy-MM-dd HH:mm")
                
                '发送医嘱
                If rsTmp!诊疗类别 = "4" Then
                    .TextMatrix(i, COLSend("发送医嘱")) = "卫生材料-" & rsTmp!诊疗项目
                ElseIf InStr(",5,6,7,", rsTmp!诊疗类别) > 0 Then
                    .TextMatrix(i, COLSend("发送医嘱")) = "药品医嘱-" & rsTmp!诊疗项目
                ElseIf rsTmp!诊疗类别 = "E" And InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
                    .TextMatrix(i, COLSend("发送医嘱")) = "给药途径-" & rsTmp!诊疗项目
                ElseIf rsTmp!诊疗类别 = "E" And vsAdvice.TextMatrix(lngRow, COL_诊疗类别) = "K" Then
                    .TextMatrix(i, COLSend("发送医嘱")) = "输血途径-" & rsTmp!诊疗项目
                ElseIf rsTmp!诊疗类别 = "E" And (bln配方行 Or bln检验行) Then
                    If bln检验行 Then
                        .TextMatrix(i, COLSend("发送医嘱")) = "采集方法-" & rsTmp!诊疗项目
                    ElseIf Not IsNull(rsTmp!相关ID) Then
                        .TextMatrix(i, COLSend("发送医嘱")) = "中药煎法-" & rsTmp!诊疗项目
                    Else
                        .TextMatrix(i, COLSend("发送医嘱")) = "中药用法-" & rsTmp!诊疗项目
                    End If
                ElseIf Not IsNull(rsTmp!相关ID) Then
                    If rsTmp!诊疗类别 = "C" Then
                        .TextMatrix(i, COLSend("发送医嘱")) = "检验项目-" & rsTmp!诊疗项目
                    ElseIf rsTmp!诊疗类别 = "D" Then
                        .TextMatrix(i, COLSend("发送医嘱")) = "检查部位-" & NVL(rsTmp!标本部位) & "(" & NVL(rsTmp!检查方法) & ")"
                    ElseIf rsTmp!诊疗类别 = "F" Then
                        .TextMatrix(i, COLSend("发送医嘱")) = "附加手术-" & rsTmp!诊疗项目
                    ElseIf rsTmp!诊疗类别 = "G" Then
                        .TextMatrix(i, COLSend("发送医嘱")) = "麻醉项目-" & rsTmp!诊疗项目
                    End If
                Else
                    .TextMatrix(i, COLSend("发送医嘱")) = rsTmp!类别名称 & "医嘱-" & rsTmp!诊疗项目
                End If
               
                .TextMatrix(i, COLSend("单据号")) = NVL(rsTmp!NO)
                .TextMatrix(i, COLSend("收费项目")) = NVL(rsTmp!收费项目)
                .TextMatrix(i, COLSend("发送数次")) = FormatEx(NVL(rsTmp!数量), 5) & NVL(rsTmp!单位)
                .TextMatrix(i, COLSend("计费状态")) = NVL(rsTmp!计费状态)
                If rsTmp!状态说明 & "" <> "" Then
                    bln状态说明 = True
                End If
                .TextMatrix(i, COLSend("执行状态")) = NVL(rsTmp!执行状态)
                .TextMatrix(i, COLSend("执行科室")) = NVL(rsTmp!执行科室)
                .TextMatrix(i, COLSend("发送人")) = NVL(rsTmp!发送人)
                .TextMatrix(i, COLSend("状态说明")) = NVL(rsTmp!状态说明)
                .TextMatrix(i, COLSend("记录性质")) = NVL(rsTmp!记录性质)
                .TextMatrix(i, COLSend("执行时间")) = Format(NVL(rsTmp!完成时间), "yyyy-MM-dd HH:mm")
                .TextMatrix(i, COLSend("执行人")) = NVL(rsTmp!完成人)
                .TextMatrix(i, COLSend("执行说明")) = NVL(rsTmp!执行说明)
                
                '已收费的划价单突出显示
                If .TextMatrix(i, COLSend("计费状态")) = "已缴费" Then
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &HC00000 '深蓝
                ElseIf .TextMatrix(i, COLSend("计费状态")) = "已退费" Then
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &H808080 '灰色
                End If
                If vsAdvice.TextMatrix(lngRow, COL_诊疗类别) = "K" And rsTmp!诊疗类别 & "" = "K" Then
                    If gbln血库系统 Then
                        lng输血 = i
                    End If
                End If
                rsTmp.MoveNext
            Next
        End If
        
        If lng输血 <> -1 Then
            '输血医嘱其它诊疗项目的信息
            strSQL = "select b.名称 as 诊疗项目,a.申请量 as 数量,b.计算单位 as 单位,a.诊疗项目id from 输血申请项目 a,诊疗项目目录 b where a.诊疗项目id=b.id and a.医嘱id=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(lngRow, COL_ID)))
            For i = 1 To rsTmp.RecordCount
                If Val(rsTmp!诊疗项目ID & "") <> Val(vsAdvice.TextMatrix(lngRow, COL_诊疗项目ID)) Then
                    .AddItem ""
                    For j = .FixedCols To .Cols - 1
                        .TextMatrix(.Rows - 1, j) = .TextMatrix(lng输血, j)
                    Next
                    .TextMatrix(.Rows - 1, COLSend("发送医嘱")) = "输血医嘱-" & rsTmp!诊疗项目
                    .TextMatrix(.Rows - 1, COLSend("发送数次")) = FormatEx(NVL(rsTmp!数量), 5) & NVL(rsTmp!单位)
                    .Cell(flexcpBackColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = .Cell(flexcpBackColor, lng输血, .FixedCols)
                Else
                    .TextMatrix(lng输血, COLSend("发送数次")) = FormatEx(NVL(rsTmp!数量), 5) & NVL(rsTmp!单位)
                End If
                rsTmp.MoveNext
            Next
        End If
        
        .MergeCells = flexMergeFree
        .MergeCol(COLSend("发送号")) = True
        .MergeCol(COLSend("发送时间")) = True
        .MergeCol(COLSend("单据号")) = True
        .MergeCol(COLSend("发送医嘱")) = True
        .MergeCol(COLSend("收费项目")) = True
        .MergeCol(COLSend("执行时间")) = True
        .MergeCol(COLSend("执行说明")) = True
        .MergeCol(COLSend("发送人")) = True
        .MergeCol(COLSend("状态说明")) = True
        
        .ColHidden(COLSend("状态说明")) = Not bln状态说明
        .Row = 1: .Col = COLSend("发送医嘱")
        .Redraw = True
        Call vsAppend_AfterRowColChange(-1, -1, .Row, .Col)
    End With
    ShowSendList = True
    Exit Function
errH:
    vsAppend.Redraw = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ShowSignList(ByVal lngRow As Long) As Boolean
'功能：显示指定行医嘱的签名记录
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strSub As String, i As Long
    Dim lngBegin As Long, lngEnd As Long
    
    On Error GoTo errH
    
    With vsAppend
        .Redraw = False
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
    
        If Val(vsAdvice.TextMatrix(lngRow, COL_ID)) = 0 Then
            .Redraw = True: ShowSignList = True: Exit Function
        End If
        
        strSQL = "Select A.签名ID,A.操作类型,B.签名时间,B.签名人,B.时间戳," & _
            " Decode(A.操作类型,1,'新开医嘱',4,'作废医嘱','其它操作') as 签名类型" & _
            " From 病人医嘱状态 A,医嘱签名记录 B Where A.医嘱ID=[1] And A.签名ID=B.ID Order by B.签名时间"
        If mblnMoved Then
            strSQL = Replace(strSQL, "病人医嘱状态", "H病人医嘱状态")
            strSQL = Replace(strSQL, "医嘱签名记录", "H医嘱签名记录")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(lngRow, COL_ID)))
        If Not rsTmp.EOF Then
            .Rows = rsTmp.RecordCount + 1
            For i = 1 To rsTmp.RecordCount
                .RowData(i) = Val(rsTmp!签名ID)
                .TextMatrix(i, 0) = rsTmp!签名类型
                .Cell(flexcpData, i, 0) = Val(rsTmp!操作类型)
                .TextMatrix(i, 1) = Format(rsTmp!签名时间, "yyyy-MM-dd HH:mm:ss")
                .TextMatrix(i, 2) = rsTmp!签名人
                .TextMatrix(i, 3) = Format(rsTmp!时间戳, "yyyy-MM-dd HH:mm:ss")
                Set .Cell(flexcpPicture, i, 0) = frmIcons.imgSign.ListImages("签名").Picture
                rsTmp.MoveNext
            Next
        End If
        .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, 0) = 0
        .Row = 1
        .Redraw = True
        Call vsAppend_AfterRowColChange(-1, -1, .Row, .Col)
    End With
    ShowSignList = True
    Exit Function
errH:
    vsAppend.Redraw = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ShowBillAppend(ByVal lngRow As Long, Optional blnExist As Boolean) As Boolean
'功能：显示指定行医嘱的单据附项内容
'返回：blnExist=医嘱是否存在单据附项内容
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lngIdx As Long
    
    blnExist = False
    rtfAppend.Text = "": rtfAppend.SelStart = 0
    
    On Error GoTo errH
    
    strSQL = "Select 项目,内容 From 病人医嘱附件 Where 医嘱ID=[1] Order by 排列"
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人医嘱附件", "H病人医嘱附件")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsAdvice.TextMatrix(lngRow, COL_ID)))
    If Not rsTmp.EOF Then
        With rtfAppend
            Do While Not rsTmp.EOF
                .SelBold = False
                .SelText = IIF(.Text = "", "", vbCrLf) & rsTmp!项目 & "：" & NVL(rsTmp!内容)
                lngIdx = .Find(rsTmp!项目 & "：", , , rtfNoHighlight Or rtfMatchCase)
                If lngIdx <> -1 Then
                    .SelStart = lngIdx
                    .SelLength = Len(rsTmp!项目 & "：")
                    .SelBold = True
                    .SelIndent = 100
                End If
                .SelStart = Len(.Text)
                
                rsTmp.MoveNext
            Loop
            
            '光标定位在第一个输入附项
            rsTmp.MoveFirst
            lngIdx = .Find(rsTmp!项目 & "：", 0, , rtfNoHighlight Or rtfMatchCase)
            If lngIdx <> -1 Then .SelStart = lngIdx + Len(rsTmp!项目 & "：")
            
            Call SetRTFFont(1)
        End With
        blnExist = True
    End If
    
    ShowBillAppend = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ShowAdvicePlan(ByVal lngRow As Long, Optional blnExist As Boolean) As Boolean
'功能：显示指定行医嘱的执行安排信息
'返回：blnExist=医嘱是否存在执行安排信息
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    blnExist = False
    rtfInfo.Text = "": rtfInfo.SelStart = 0
    
    On Error GoTo errH
    
    With vsAdvice
        If InStr("D,F,G,", .TextMatrix(lngRow, COL_诊疗类别)) > 0 Or _
            .TextMatrix(lngRow, COL_诊疗类别) = "E" And InStr(",0,6,", "," & .TextMatrix(lngRow, COL_操作类型) & ",") > 0 Then
            
            If .TextMatrix(lngRow, COL_诊疗类别) = "E" And Val(.TextMatrix(lngRow, COL_操作类型)) = 6 Then
                strSQL = "Select a.安排时间,a.执行间,a.执行说明 From 病人医嘱发送 a,病人医嘱记录 b " & _
                        "Where a.医嘱ID = b.ID And b.相关ID=[1] And (a.执行说明 is Not Null Or a.安排时间 is Not Null) And Rownum=1"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COL_ID)))
            Else
                strSQL = "Select 安排时间,执行间,执行说明 From 病人医嘱发送 Where 医嘱ID=[1] And (执行说明 is Not Null Or 安排时间 is Not Null)"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COL_ID)))
            End If
            
            If Not rsTmp.EOF Then
                strSQL = ""
                
                If Not IsNull(rsTmp!安排时间) Then
                    strSQL = strSQL & vbCrLf & "安排时间：" & Format(rsTmp!安排时间, "yyyy-MM-dd HH:mm")
                End If
                If Not IsNull(rsTmp!执行间) Then
                    strSQL = strSQL & vbCrLf & "执行间：" & rsTmp!执行间
                End If
                strSQL = strSQL & vbCrLf & NVL(rsTmp!执行说明)
                
                rtfInfo.Text = Mid(strSQL, 3)
                
                Call SetRTFFont(2)
                blnExist = True
            End If
        End If
    End With
    ShowAdvicePlan = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ShowOtherAppend(ByVal lngRow As Long) As Boolean
'功能：显示指定行医嘱的审核信息
'说明：只检查审核状态通过和未通过的医嘱
'返回：是否存在审核信息
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim int类型 As Integer
    Dim str操作员 As String
    Dim str时间 As String
    
    strSQL = "Select 操作人员,操作时间 From 病人医嘱状态 Where 医嘱id = [1] And 操作类型 = [2]"
    
    str操作员 = "审核人：": str时间 = "审核时间："
    Select Case vsAdvice.TextMatrix(lngRow, COL_审核状态)
        Case 2
            If gbln血库系统 And vsAdvice.TextMatrix(lngRow, COL_诊疗类别) = "K" Then '输血医嘱处理流程变动 70823
                int类型 = 15 '血库审核通过
                str操作员 = "血库审核人："
                str时间 = "血库审核时间："
            Else
                int类型 = 11
            End If
        Case 3
            int类型 = 12
        Case 4
            int类型 = 11
        Case 5
            int类型 = 14
            str操作员 = "血库接收人："
            str时间 = "血库接收时间："
    End Select
    rtfOther.Text = ""
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsAdvice.TextMatrix(lngRow, COL_ID)), int类型)
    If Not rsTmp.EOF Then
        strSQL = ""
        Do While Not rsTmp.EOF
            strSQL = IIF(strSQL = "", "", strSQL & vbCrLf) & str操作员 & rsTmp!操作人员 & vbCrLf & _
                str时间 & Format(rsTmp!操作时间 & "", "YYYY-MM-DD HH:MM:SS")
            rsTmp.MoveNext
        Loop
        rtfOther.Text = strSQL
        Call SetRTFFont(3)
        ShowOtherAppend = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function LoadBillList() As Boolean
'功能：显示指定行的医嘱发送可以打印的诊疗单据在菜单上
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objControl As CommandBarControl
    Dim objPopup As CommandBarPopup
    Dim objBar As CommandBar
    Dim objMenu As CommandBarPopup
    Dim objpopup1 As CommandBarPopup
    
    If mcbsMain Is Nothing Then LoadBillList = True: Exit Function
    Set objPopup = mcbsMain.FindControl(, conMenu_Report_ClinicBill, False, True)
    If objPopup Is Nothing Then LoadBillList = True: Exit Function
    Set objBar = mcbsMain(2)
    If objBar Is Nothing Then LoadBillList = True: Exit Function
    objPopup.Visible = True
    
    objPopup.CommandBar.Controls.DeleteAll
    
    If mcbsMain Is Nothing Then LoadBillList = True: Exit Function
    Set objMenu = mcbsMain.FindControl(, conMenu_EditPopup, False, True)
    If objMenu Is Nothing Then LoadBillList = True: Exit Function
    Set objpopup1 = objMenu.CommandBar.FindControl(, conMenu_Report_ClinicBill)
    objpopup1.Visible = True
    For i = objMenu.CommandBar.Controls.Count To 1 Step -1
        If objMenu.CommandBar.Controls(i).ID > conMenu_Report_ClinicBill * 100# And objMenu.CommandBar.Controls(i).ID < conMenu_Report_ClinicBill * 100# + 100 Then
            objMenu.CommandBar.Controls(i).Delete
        End If
    Next
    For i = objBar.Controls.Count To 1 Step -1
        If objBar.Controls(i).ID > conMenu_Report_ClinicBill * 100# And objBar.Controls(i).ID < conMenu_Report_ClinicBill * 100# + 100 Then
            objBar.Controls(i).Delete
        End If
    Next
    For i = objpopup1.CommandBar.Controls.Count To 1 Step -1
        If objpopup1.CommandBar.Controls(i).ID > conMenu_Report_ClinicBill * 100# And objpopup1.CommandBar.Controls(i).ID < conMenu_Report_ClinicBill * 100# + 100 Then
            objpopup1.CommandBar.Controls(i).Delete
        End If
    Next
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) = 0 _
         Or Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱状态)) <> 8 Then
        LoadBillList = True: Exit Function
    End If
        
    On Error GoTo errH
    
    Set rsTmp = GetBillList

    '如果只有一个诊疗单据可用，则直接加入到医嘱菜单里
    If rsTmp.RecordCount = 1 Then
        objPopup.Visible = False
        objPopup.Category = "已判断"
        objpopup1.Visible = False
        objpopup1.Category = "已判断"
        Set objPopup = objMenu
    End If
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            With objPopup.CommandBar.Controls
                Set objControl = .Add(xtpControlButton, conMenu_Report_ClinicBill * 100# + i, IIF(rsTmp.RecordCount = 1, "打印:", "") & rsTmp!名称)
                If i <= 10 Then
                    objControl.Caption = objControl.Caption & "(&" & i - 1 & ")"
                ElseIf i <= 36 Then
                    objControl.Caption = objControl.Caption & "(&" & Chr(i - 11 + Asc("A")) & ")"
                End If
                '中药的煎法用法单据号和中药不一样，界面上显示的中药用法，所以把单据的NO拼进去
                objControl.Parameter = "ZLCISBILL" & Format(rsTmp!编号, "00000") & "-1" & "|" & rsTmp!NO  '对应的自定义报表编号
                'If i > 1 Then objControl.Enabled = False '一个项目只能设置一个诊疗单据
            End With
            '菜单和工具栏要分开加
            If rsTmp.RecordCount > 1 Then
                With objpopup1.CommandBar.Controls
                    Set objControl = .Add(xtpControlButton, conMenu_Report_ClinicBill * 100# + i, rsTmp!名称)
                    If i <= 10 Then
                        objControl.Caption = objControl.Caption & "(&" & i - 1 & ")"
                    ElseIf i <= 36 Then
                        objControl.Caption = objControl.Caption & "(&" & Chr(i - 11 + Asc("A")) & ")"
                    End If
                    objControl.Parameter = "ZLCISBILL" & Format(rsTmp!编号, "00000") & "-1" & "|" & rsTmp!NO  '对应的自定义报表编号
                End With
            End If
            If rsTmp.RecordCount = 1 Then
                With objBar.Controls
                    Set objControl = .Find(, conMenu_Report_ClinicBill)
                    If Not objControl Is Nothing Then
                        Set objControl = .Add(xtpControlButton, conMenu_Report_ClinicBill * 100# + i, "打印单据", objControl.Index + 1)
                        If i <= 10 Then
                            objControl.Caption = objControl.Caption
                        ElseIf i <= 36 Then
                            objControl.Caption = objControl.Caption
                        End If
                        objControl.Parameter = "ZLCISBILL" & Format(rsTmp!编号, "00000") & "-1" & "|" & rsTmp!NO  '对应的自定义报表编号
                        objControl.IconId = conMenu_File_Print
                        objControl.Style = xtpButtonIconAndCaption
                        'If i > 1 Then objControl.Enabled = False '一个项目只能设置一个诊疗单据
                    End If
                End With
            End If
            rsTmp.MoveNext
        Next
    End If
    
    LoadBillList = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub vsAppend_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow = OldRow Then Exit Sub
    
End Sub

Private Sub vsAppend_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    
    With vsAppend
        If Button = 2 And tbcAppend.Selected.Tag = "签名" Then
            If mcbsMain Is Nothing Then Exit Sub
            If Between(.MouseRow, .FixedRows, .Rows - 1) Then
                If Between(.MouseCol, .FixedCols, .Cols - 1) Then
                    Set objPopup = mcbsMain.FindControl(, conMenu_Tool_Sign, False, True) '可能故意隐藏了
                    If Not objPopup Is Nothing Then
                        If objPopup.CommandBar.Controls.Count > 0 Then
                            'ShowPopup不会触发InitCommandsPopup事件
                            objPopup.CommandBar.ShowPopup
                        End If
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub vsAppend_GotFocus()
    vsAppend.BackColorSel = &HFFCC99
    
    '因为绑定相同,获取焦点时会丢失绑定,Resize会恢复
    tbcAppend.Height = tbcAppend.Height + 30
    tbcAppend.Height = tbcAppend.Height - 30
End Sub

Private Sub vsAppend_LostFocus()
    vsAppend.BackColorSel = &HFFEBD7
End Sub

Private Sub vsAdvice_GotFocus()
    vsAdvice.BackColorSel = &HFFCC99
End Sub

Private Sub vsAdvice_LostFocus()
    vsAdvice.BackColorSel = &HFFEBD7
End Sub

Private Function RowIn一并给药(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'功能：判断指定行是否在一并给药的范围中,如果是,同时返回行号范围
    Dim i As Long, blnTmp As Boolean
    With vsAdvice
        If .TextMatrix(lngRow, COL_诊疗类别) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, COL_诊疗类别)) = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowIn一并给药 = blnTmp
    End With
End Function

Private Function RowInSameNo(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'功能：判断指定行是否是在同一个单据号范围中，如果是返回行号范围
    Dim i As Long
 
    With vsAdvice
        lngBegin = lngRow
        For i = lngRow - 1 To .FixedRows Step -1
            If .TextMatrix(i, COL_诊疗类别) = "5" Or .TextMatrix(i, COL_诊疗类别) = "6" Or .TextMatrix(i, COL_诊疗类别) = "E" And Val(.TextMatrix(i, COL_操作类型)) = 4 Then
                If .Cell(flexcpData, i, COL_处方号) = .Cell(flexcpData, lngRow, COL_处方号) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            End If
        Next
        lngEnd = lngRow
        For i = lngRow + 1 To .Rows - 1
            If .TextMatrix(i, COL_诊疗类别) = "5" Or .TextMatrix(i, COL_诊疗类别) = "6" Or .TextMatrix(i, COL_诊疗类别) = "E" And Val(.TextMatrix(i, COL_操作类型)) = 4 Then
                If .Cell(flexcpData, i, COL_处方号) = .Cell(flexcpData, lngRow, COL_处方号) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            End If
        Next
    End With
    If lngEnd <> lngRow Or lngBegin <> lngRow Then
        RowInSameNo = True
    End If
End Function

Private Sub ShowTotalMoney()
'功能：医嘱总金额的提示
'说明：由于药品时价，和给药途径，中药煎法用法等，新开医嘱不一定准确
    Dim rsMoney As New ADODB.Recordset, strSQL As String, str诊疗收费 As String
    Dim cur应收 As Currency, cur实收 As Currency
    Dim cur药品应收 As Currency, cur药品实收 As Currency
    Dim cur新开 As Currency, cur药品新开 As Currency
    Dim cur预交 As Currency, curTmp As Currency
    Dim strSQLTmp As String
    Dim strTmp As String
    
    '计算变价药品的时价，拼接到查询语句中
    strSQLTmp = "Zl_Calcdrugprice(a.执行科室id, s.药品id, a.总给予量," & gbytMediOutMode & "," & Len(gstrDecPrice) - 2 & "," & Len(gstrDecPrice) - 2 & ")"
 
    On Error GoTo errH
    
    strSQL = _
        " Select /*+ RULE */ Sum(A.应收金额) as 应收金额,Sum(A.实收金额) as 实收金额," & _
        " Sum(Decode(Instr('567',A.收费类别),0,0,A.应收金额)) as 药品应收," & _
        " Sum(Decode(Instr('567',A.收费类别),0,0,A.实收金额)) as 药品实收" & _
        " From 门诊费用记录 A,病人医嘱发送 B,病人医嘱记录 C" & _
        " Where A.医嘱序号=B.医嘱ID And B.医嘱ID=C.ID" & _
        " And C.病人ID+0=[1] And C.挂号单=[2]"
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
        strSQL = Replace(strSQL, "门诊费用记录", "H门诊费用记录")
    ElseIf zlDatabase.DateMoved(mvRegDate) Then
        strTmp = strSQL
        strTmp = Replace(strTmp, "病人医嘱记录", "H病人医嘱记录")
        strTmp = Replace(strTmp, "病人医嘱发送", "H病人医嘱发送")
        strTmp = Replace(strTmp, "门诊费用记录", "H门诊费用记录")
        strSQL = strSQL & " Union ALL " & strTmp

        strSQL = "Select Sum(应收金额) as 应收金额,Sum(实收金额) as 实收金额," & _
            " Sum(药品应收) as 药品应收,Sum(药品实收) as 药品实收 From (" & strSQL & ")"
    End If
    Set rsMoney = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng病人ID, mstr挂号单)
    If Not rsMoney.EOF Then
        cur应收 = NVL(rsMoney!应收金额, 0)
        cur实收 = NVL(rsMoney!实收金额, 0)
        cur药品应收 = NVL(rsMoney!药品应收, 0)
        cur药品实收 = NVL(rsMoney!药品实收, 0)
    End If
    
    str诊疗收费 = "Select * From (" & _
        "Select Distinct C.诊疗项目ID,C.收费项目ID,C.检查部位,C.检查方法,C.费用性质,C.收费数量,C.固有对照,C.从属项目,C.收费方式,C.适用科室id" & _
        " ,Max(Nvl(c.适用科室id, 0)) Over(Partition By c.诊疗项目id, c.检查部位, c.检查方法, c.费用性质) As Top" & _
        " From 诊疗收费关系 C,病人医嘱记录 A Where A.病人ID+0=[1] And A.挂号单=[2] And A.诊疗项目ID+0=C.诊疗项目ID" & _
        "   And (a.相关id Is Null And a.执行标记 In (1, 2) And c.费用性质 = 1 Or" & vbNewLine & _
        "   a.标本部位 = c.检查部位 And a.检查方法 = c.检查方法 And Nvl(c.费用性质, 0) = 0 Or" & vbNewLine & _
        "   (a.检查方法 Is Null or a.诊疗类别 = 'E' And Exists(Select 1 From 诊疗项目目录 Z Where Z.id=a.诊疗项目ID And Z.操作类型='4')) And Nvl(c.费用性质, 0) = 0 And c.检查部位 Is Null And c.检查方法 Is Null)" & _
        "      And (C.适用科室ID is Null or C.适用科室ID = A.执行科室ID And C.病人来源 = 1)" & _
        " ) Where Nvl(适用科室id, 0) = Top"
        
    '时价药品取"指导零售价"
    strSQL = _
        "Select Sum(Round(金额," & gbytDec & ")) As 金额,Sum(Round(药品金额," & gbytDec & ")) As 药品金额" & _
        " From (Select A.总给予量*Decode(I.是否变价,1," & strSQLTmp & ",P.现价) As 金额," & _
        "              A.总给予量*Decode(I.是否变价,1," & strSQLTmp & ",P.现价) As 药品金额" & _
        "       From 病人医嘱记录 A,收费项目目录 I,收费价目 P,药品规格 S" & _
        "       Where A.收费细目ID=I.ID And I.ID=P.收费细目ID And I.ID=S.药品ID" & _
        GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "I", "P", "3", "4", "5") & _
        "             And (Sysdate Between P.执行日期 And P.终止日期 Or Sysdate>=P.执行日期 And P.终止日期 is Null)" & _
        "             And A.医嘱状态=1 And A.诊疗类别 In ('5','6')" & _
        "             And A.病人ID+0=[1] And A.挂号单=[2]" & _
        "       Union All" & _
        "       Select A.总给予量*A.单次用量/S.剂量系数*Decode(I.是否变价,1," & strSQLTmp & ",P.现价) As 金额," & _
        "              A.总给予量*A.单次用量/S.剂量系数*Decode(I.是否变价,1," & strSQLTmp & ",P.现价) As 药品金额" & _
        "       From 病人医嘱记录 A,收费项目目录 I,收费价目 P,药品规格 S" & _
        "       Where A.收费细目ID=I.ID And I.ID=P.收费细目ID And I.ID=S.药品ID" & _
        GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "I", "P", "3", "4", "5") & _
        "             And (Sysdate Between P.执行日期 And P.终止日期 Or Sysdate>=P.执行日期 And P.终止日期 is Null)" & _
        "             And A.医嘱状态=1 And A.诊疗类别='7'" & _
        "             And A.病人ID+0=[1] And A.挂号单=[2]"
        strSQL = strSQL & "  Union All" & _
        "       Select Nvl(A.总给予量,A.频率次数)*R.收费数量*Decode(I.是否变价,1,P.缺省价格,P.现价) As 金额,0 as 药品金额" & _
        "       From 病人医嘱记录 A,(" & str诊疗收费 & ") R,收费项目目录 I,收费价目 P" & _
        "       Where A.诊疗项目ID+0=R.诊疗项目ID And I.ID=R.收费项目ID And I.ID=P.收费细目ID" & _
        GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "I", "P", "3", "4", "5") & _
        "             And (I.站点='" & gstrNodeNo & "' Or I.站点 is Null)" & _
        "             And (Sysdate Between P.执行日期 And P.终止日期 Or Sysdate>=P.执行日期 And P.终止日期 is Null)" & _
        "             And Nvl(A.计价特性,0)=0 And A.医嘱状态=1 And A.诊疗类别 Not In ('5','6','7')" & _
        "             And A.病人ID+0=[1] And A.挂号单=[2]" & _
        "             And (a.相关id Is Null And a.执行标记 In (1, 2) And r.费用性质 = 1 Or" & _
        "                 a.标本部位 = r.检查部位 And a.检查方法 = r.检查方法 And Nvl(r.费用性质, 0) = 0 Or" & _
        "                 a.检查方法 Is Null And Nvl(r.费用性质, 0) = 0 And r.检查部位 Is Null And r.检查方法 Is Null)) A"
    Set rsMoney = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng病人ID, mstr挂号单, mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
    If Not rsMoney.EOF Then
        cur新开 = NVL(rsMoney!金额, 0)
        cur药品新开 = NVL(rsMoney!药品金额, 0)
    End If
    
    strSQL = "Select Nvl(预交余额,0)-Nvl(费用余额,0) as 金额 From 病人余额 Where 性质=1 And 类型 = 1 And 病人ID=[1]"
    Set rsMoney = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng病人ID)
    If Not rsMoney.EOF Then cur预交 = NVL(rsMoney!金额, 0)
    
    strSQL = _
        "医嘱已发送应收:" & FormatEx(cur应收, gbytDec) & "(药" & FormatEx(cur药品应收, gbytDec) & ")," & _
        "实收:" & FormatEx(cur实收, gbytDec) & "(药" & FormatEx(cur药品实收, gbytDec) & ")" & _
        "  新开约:" & FormatEx(cur新开, gbytDec) & "(药" & FormatEx(cur药品新开, gbytDec) & ")" & _
        IIF(cur预交 = 0, "", "  预交:" & FormatEx(cur预交, 2))
    If cur预交 <> 0 And cur新开 > cur预交 Then
        curTmp = cur新开 - cur预交
        strSQL = strSQL & "  需补交:" & FormatEx(curTmp, gbytDec)
    End If
    RaiseEvent StatusTextUpdate(strSQL)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsColumn_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngCol As Long
    
    If Col = 0 Then
        lngCol = vsColumn.RowData(Row)
        If Val(vsColumn.TextMatrix(Row, 0)) <> 0 Then
            vsAdvice.ColWidth(lngCol) = vsAdvice.ColData(lngCol)
            vsAdvice.ColHidden(lngCol) = False
        Else
            vsAdvice.ColWidth(lngCol) = 0
            vsAdvice.ColHidden(lngCol) = True
        End If
    End If
End Sub

Private Sub vsColumn_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsColumn
        If NewRow >= .FixedRows - 1 And NewCol >= .FixedCols - 1 Then
            .ForeColorSel = .Cell(flexcpForeColor, NewRow, 1)
            .Col = 0
        End If
    End With
End Sub

Private Sub vsColumn_LostFocus()
    vsColumn.Visible = False
End Sub

Private Sub vsColumn_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Or vsColumn.Cell(flexcpForeColor, Row, 1) = vsColumn.BackColorFixed Then Cancel = True
End Sub

Private Sub FuncBloodApply(ByVal intType As Long)
'功能：输血申请单
'参数：intType=0 新增，=1修改，=2查看，=4核对
    Dim datTurn As Date
    Dim lngUpdateAdvice As Long
    Dim lngRow As Long
    Dim bln用血 As Boolean
    Dim blnApply As Boolean
    
    If intType <> 2 Then
        '检查挂号病人是否超期
        If Not FuncTimeLimitCheck Then Exit Sub
        '检查是否满足中级以上专业技术职务
        If gbln输血申请中级以上 Then
            If UserInfo.专业技术职务 <> "主治医师" And UserInfo.专业技术职务 <> "主任医师" And UserInfo.专业技术职务 <> "副主任医师" Then
                MsgBox "启用了输血分级管理后，输血医嘱只有中级及以上专业技术职务医师才能下达。", vbInformation, "输血申请单"
                Exit Sub
            End If
        End If
        '修改时检查是否审核
        If intType = 1 Then
            If Not CanEditBloodAdvice(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_审核状态)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_标志)) = 1, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_检查方法)) = 1) Then Exit Sub
        End If
    End If
    
    If intType <> 0 Then
         lngUpdateAdvice = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
         bln用血 = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_检查方法)) = 1
         lngRow = vsAdvice.Row
    End If
    
    If gbln血库系统 = True Then
        blnApply = frmApplyBloodNew.ShowMe(Me, mlng病人ID, 0, 1, intType, lngUpdateAdvice, mlng挂号科室ID, , mlng挂号科室ID, , , mrsDefine, mclsMipModule, 1, mstr挂号单, , , , , mlng前提ID, IIF(bln用血 = True, 1, 0))
    Else
        blnApply = frmApplyBlood.ShowMe(Me, mlng病人ID, 0, 1, intType, lngUpdateAdvice, mlng挂号科室ID, , mlng挂号科室ID, , , mrsDefine, mclsMipModule, 1, mstr挂号单, , , , , mlng前提ID)
    End If
    
    If blnApply = True Then
        '刷新医嘱
        Call RefreshData
        '选择最后一行医嘱
        vsAdvice.Row = vsAdvice.Rows - 1
        vsAdvice.ShowCell vsAdvice.Rows - 1, col_医嘱内容
    End If
End Sub

Private Sub FuncOperationApply(ByVal intType As Long)
'功能：手术申请单
'参数：intType=0 新增，=1修改，=2查看
    Dim datTurn As Date
    Dim lngUpdateAdvice As Long
    Dim lngRow As Long, strDefine As String
    
    If intType <> 2 Then
        '检查挂号病人是否超期
        If Not FuncTimeLimitCheck Then Exit Sub
        '修改时检查是否审核
        If intType = 1 Then
            If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_审核状态)) = 2 Then
                MsgBox "申请单已经审核，不允许再修改。", vbInformation, "手术申请单"
                intType = 2
            End If
        End If
    End If
    
    If intType <> 0 Then
         lngUpdateAdvice = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
         lngRow = vsAdvice.Row
    End If
     
    If Not mrsDefine Is Nothing Then
        mrsDefine.Filter = "诊疗类别='F'"
        If Not mrsDefine.EOF Then strDefine = Trim(NVL(mrsDefine!医嘱内容))
    End If

    If frmApplyOperation.ShowMe(Me, 1, intType, mlng病人ID, mstr挂号单, 1, lngUpdateAdvice, mlng挂号科室ID, mlng挂号科室ID, strDefine, , , , 0, mclsMipModule, , , mlng前提ID) Then
        '刷新医嘱
        Call RefreshData
        '选择最后一行医嘱
        vsAdvice.Row = vsAdvice.Rows - 1
        vsAdvice.ShowCell vsAdvice.Rows - 1, col_医嘱内容
    End If
End Sub

Private Sub FuncLISApply(ByVal lng申请序号 As Long)
'功能：调用检验申请产生申请单和检验医嘱
'参数：lng申请序号=修改申请单时的申请序号
    Dim arrTmp As Variant, arrSQL As Variant, i As Long, blnTrans As Boolean, strSQL As String
    Dim strResult As String, strDiag As String, strDept As String, strErr As String
    Dim rsPati As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset '注意此变量不要乱用,在LisInfoTrans方法被赋值
    Dim rsTemp As ADODB.Recordset
    Dim lng医嘱ID As Long, lng相关ID As Long, lng序号 As Long
    Dim lng执行科室ID As Long, lng采集科室ID As Long, lng检验项目ID As Long, lng采集项目ID As Long
    Dim str检验计价性质 As String, str采集计价性质 As String, str检验执行性质 As String, str采集执行性质 As String
    Dim str检验项目 As String, str采集方法 As String, str标本 As String, str紧急 As String
    Dim strCurDate As String, str开始执行时间 As String, str医嘱内容 As String, str医嘱IDs As String, blnCancel As Boolean
    Dim strDelIDs As String, arrDelID() As String
    Dim Y As Long, j As Long, str检验项目组合 As String
    Dim str嘱托 As String, str附项 As String
    Dim arrAppend As Variant
    Dim lng附项序号 As Long
    Dim str诊断 As String
    Dim lng假医嘱ID As Long '避免医嘱ID序列值的浪费，在最后提交事务时产生真的医嘱ID
    Dim str医嘱ID As String, str相关ID As String
    Dim varID As Variant
    Dim strTmp As String
    Dim bln提醒对码 As Boolean
    Dim vMsg As VbMsgBoxResult
    Dim strItems As String
    Dim strTabAdvice As String
    Dim blnCheckItem As Boolean '医保管控监测
    Dim rsPrice As ADODB.Recordset
    Dim str摘要 As String, strMsg As String
    Dim rsLISInfo As ADODB.Recordset
    Dim lng申请组号 As Long
    Dim dat开始执行时间 As Date
    Dim dat当前时间 As Date
    
    '检查挂号病人是否超期
    If Not FuncTimeLimitCheck Then Exit Sub
    
    Set rsPati = GetPatiInfo()
    If rsPati.RecordCount = 0 Then
        MsgBox "未能正确读取病人信息！", vbInformation, gstrSysName
        Exit Sub
    End If
    If lng申请序号 <> 0 Then
        strDiag = GetAdviceDiag(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)))
    End If
    
    strDept = Sys.RowValue("部门表", mlng挂号科室ID, "名称")
    Call InitObjLis(p门诊医生站)
    If gobjLIS Is Nothing Then Exit Sub
    Call CreatePlugInOK(p门诊医嘱下达, mint场合)
    
    On Error GoTo errH
 
    '返回已选择的检验项目格式如下: 采诊科室ID1,执行科室ID1,申请时间1,诊疗项目编码1,标本1,紧急医嘱1,采集方式诊疗项目ID 1;采诊科室ID2,执行科室ID2,申请时间2,诊疗项目编码2,标本2,紧急医嘱2,采集方式诊疗项目ID 2;.....
    strResult = gobjLIS.ShowLisApplicationForm(mfrmParent, lng申请序号, mlng病人ID, 0, Val("" & rsPati!挂号ID), rsPati!姓名, "" & rsPati!性别, "" & rsPati!年龄, 1, _
        Val("" & rsPati!门诊号), Val("" & rsPati!住院号), Val("" & rsPati!健康号), strDiag, UserInfo.姓名, UserInfo.部门ID, UserInfo.部门名, mlng挂号科室ID, strDept, blnCancel, strErr)
 
    If strErr <> "" Then
        MsgBox "检验接口内部错误：" & strErr, vbInformation, gstrSysName
    ElseIf blnCancel Then
        Exit Sub    '取消，退出
    Else
        arrSQL = Array()
        '修改申请单时，先删除旧的医嘱
        If lng申请序号 <> 0 Then
            str医嘱IDs = GetAdivceBy申请序号(lng申请序号)
            For i = 0 To UBound(Split(str医嘱IDs, ","))
                '调用删除前外挂接口
                On Error Resume Next
                If Not gobjPlugIn Is Nothing Then
                    If gobjPlugIn.AdviceDeletBefor(glngSys, p门诊医嘱下达, mlng病人ID, mlng挂号ID, Val(Split(str医嘱IDs, ",")(i)), mint场合) = False Then
                        If err.Number = 0 Then Exit Sub
                    End If
                    Call zlPlugInErrH(err, "AdviceDeletBefor")
                End If
                If err.Number <> 0 Then err.Clear
                On Error GoTo errH
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_Delete(" & Split(str医嘱IDs, ",")(i) & ",1)"
                strDelIDs = strDelIDs & "," & Split(str医嘱IDs, ",")(i)
            Next
            strDelIDs = Mid(strDelIDs, 2)
        End If
        
        If strResult <> "" Then
            If strDiag = "" Then
                '取一条诊断来进行默认关联
                strSQL = "Select a.ID,a.诊断描述 From 病人诊断记录 A Where a.病人id=[1] and a.主页id =[2] and a.记录来源 = 3 order by a.诊断类型,a.诊断次序"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng挂号ID)
                If Not rsTemp.EOF Then
                    strDiag = rsTemp!ID & ""
                    str诊断 = "申请单诊断<Split2>0<Split2><Split2>" & rsTemp!诊断描述
                End If
            Else
                str诊断 = GetDiag诊断描述(strDiag)
                If str诊断 <> "" Then
                    str诊断 = "申请单诊断<Split2>0<Split2><Split2>" & str诊断
                End If
            End If
              
            bln提醒对码 = True
            
            If mint险类 <> 0 Then
                If gclsInsure.GetCapability(support实时监控, mlng病人ID, mint险类) Then
                    blnCheckItem = True
                End If
            End If
            dat当前时间 = zlDatabase.Currentdate()
            strCurDate = "To_Date('" & Format(dat当前时间, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')"
            lng序号 = GetMaxAdviceNO(mlng病人ID, 0, 0)
            lng申请组号 = -1
            '在该方法中对rsLISInfo, rsTmp赋值
            Call LisInfoTrans(strResult, rsLISInfo, rsTmp)
                        
            '只产生临嘱
            For i = 1 To rsLISInfo.RecordCount
                If lng申请组号 <> Val(rsLISInfo!组号 & "") Then
                    lng申请组号 = Val(rsLISInfo!组号 & "")
                    lng申请序号 = Get申请序号
                End If
                lng假医嘱ID = lng假医嘱ID + 1
                str相关ID = "<FAKEID>" & lng假医嘱ID & "</FAKEID>"
                lng相关ID = lng假医嘱ID
                lng采集科室ID = Val(rsLISInfo!采集科室ID & "")
                lng执行科室ID = Val(rsLISInfo!执行科室ID & "")
                str开始执行时间 = rsLISInfo!开始执行时间 & ""
                str标本 = rsLISInfo!标本 & ""
                str附项 = rsLISInfo!附项 & ""
                str嘱托 = rsLISInfo!嘱托 & ""
                str紧急 = rsLISInfo!紧急 & ""
                lng采集项目ID = Val(rsLISInfo!采集项目ID & "")
                lng检验项目ID = Val(rsLISInfo!检验项目ID & "")
                
                dat开始执行时间 = CDate(str开始执行时间)
                str开始执行时间 = "To_Date('" & Format(dat开始执行时间, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')"
                
                'a.先产生检验医嘱 申请单开出来的的检验医嘱只有一个检验项目ID
                rsTmp.Filter = "ID=" & lng检验项目ID
                str检验项目 = rsTmp!名称 & ""
                str检验计价性质 = Val("" & rsTmp!计价性质)
                str检验执行性质 = IIF("" & rsTmp!执行科室 = "", "NULL", "" & rsTmp!执行科室)
                str医嘱内容 = str检验项目 & IIF("" = rsLISInfo!时间内容 & "", "", "(" & rsLISInfo!时间内容 & ")")
                lng序号 = lng序号 + 1
                str摘要 = gclsInsure.GetItemInfo(mint险类, mlng病人ID, 0, "", 0, "", CStr(lng检验项目ID) & "||1")
                blnCancel = CheckLISAppAdvice(1, mlng病人ID, mlng挂号ID, mint险类, "C", lng检验项目ID, mlng挂号科室ID, UserInfo.姓名, lng执行科室ID, Val(rsTmp!执行科室 & ""), str摘要 & "||0||0|| ||0")
                
                If Not blnCancel Then Exit Sub
                
                lng假医嘱ID = lng假医嘱ID + 1
                str医嘱ID = "<FAKEID>" & lng假医嘱ID & "</FAKEID>"
                lng医嘱ID = lng假医嘱ID
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_Insert(" & _
                    str医嘱ID & "," & str相关ID & "," & lng序号 & ",1," & mlng病人ID & "," & _
                    "Null,0,1,1,'C'," & _
                    lng检验项目ID & ",Null,Null,Null,1," & _
                    "'" & str医嘱内容 & "',Null," & "'" & str标本 & "','一次性',Null," & _
                    "Null,Null,Null," & str检验计价性质 & "," & lng执行科室ID & _
                    "," & str检验执行性质 & "," & str紧急 & "," & str开始执行时间 & ",Null," & mlng挂号科室ID & "," & _
                    mlng挂号科室ID & ",'" & UserInfo.姓名 & "'," & strCurDate & ",'" & mstr挂号单 & "'," & ZVal(mlng前提ID) & "," & _
                    "NULL,0,Null," & IIF(str摘要 = "", "Null", "'" & str摘要 & "'") & ",'" & UserInfo.姓名 & "'" & _
                    ",Null,Null,Null,Null," & lng申请序号 & ",null,null,null,null,null,'" & rsLISInfo!时间ID & "')"
                    
                strItems = strItems & "," & lng检验项目ID & ":" & lng执行科室ID
                
                If blnCheckItem Then
                    strTabAdvice = _
                        "select " & lng医嘱ID & " as ID," & lng序号 & " as 序号," & lng相关ID & " as 相关ID,'C' as 诊疗类别," & lng检验项目ID & " as 管码项目ID," & _
                        lng检验项目ID & " as 诊疗项目ID,-null as 收费细目ID, 1 As 总量, 0 As 单量,'" & str标本 & "' as 标本部位,'' As 检查方法," & _
                        "0 as 执行标记," & Val("" & rsTmp!计价性质) & " as 计价特性, 0 As 附加手术," & Val("" & rsTmp!执行科室) & " As 执行性质," & lng执行科室ID & " as 执行科室id from dual"
                End If
            

                'b.再产生采集方法医嘱
                rsTmp.Filter = "ID=" & lng采集项目ID
                str采集方法 = rsTmp!名称 & ""
                str采集计价性质 = Val("" & rsTmp!计价性质)
                str采集执行性质 = "" & rsTmp!执行科室
                str医嘱内容 = AdviceMakeText(str检验项目, str采集方法, str标本)
                If "" <> rsLISInfo!时间内容 & "" Then str医嘱内容 = str医嘱内容 & "(" & rsLISInfo!时间内容 & ")"
                lng序号 = lng序号 + 1
                str摘要 = gclsInsure.GetItemInfo(mint险类, mlng病人ID, 0, "", 0, "", CStr(lng采集项目ID) & "||1")
                blnCancel = CheckLISAppAdvice(1, mlng病人ID, mlng挂号ID, mint险类, "E", lng采集项目ID, mlng挂号科室ID, UserInfo.姓名, lng采集科室ID, Val(rsTmp!执行科室 & ""), str摘要 & "||0||0|| ||0")
                If Not blnCancel Then Exit Sub
                    
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_Insert(" & _
                    str相关ID & ",Null," & lng序号 & ",1," & mlng病人ID & "," & _
                    "Null,0,1,1,'E'," & _
                    lng采集项目ID & ",Null,Null,Null,1," & _
                    "'" & str医嘱内容 & "','" & str嘱托 & "'," & "'" & str标本 & "','一次性',Null," & _
                    "Null,Null,Null," & str采集计价性质 & "," & lng采集科室ID & _
                    "," & str采集执行性质 & "," & str紧急 & "," & str开始执行时间 & ",Null," & mlng挂号科室ID & "," & _
                    mlng挂号科室ID & ",'" & UserInfo.姓名 & "'," & strCurDate & ",'" & mstr挂号单 & "'," & ZVal(mlng前提ID) & "," & _
                    "NULL,0,Null," & IIF(str摘要 = "", "Null", "'" & str摘要 & "'") & ",'" & UserInfo.姓名 & "'" & _
                    ",Null,Null,Null,Null," & lng申请序号 & ",null,null,null,null,null,'" & rsLISInfo!时间ID & "')"
                    
                strItems = strItems & "," & lng采集项目ID & ":" & lng采集科室ID
                
                If blnCheckItem Then
                    strTabAdvice = strTabAdvice & " Union ALL " & _
                        "select " & lng相关ID & " as ID," & lng序号 & " as 序号,-null as 相关ID,'E' as 诊疗类别," & lng检验项目ID & " as 管码项目ID," & _
                        lng采集项目ID & " as 诊疗项目ID,-null as 收费细目ID, 1 As 总量, 0 As 单量,'" & str标本 & "' as 标本部位,'' As 检查方法," & _
                        "0 as 执行标记," & Val("" & rsTmp!计价性质) & " as 计价特性, 0 As 附加手术," & Val("" & rsTmp!执行科室) & " As 执行性质," & lng采集科室ID & " as 执行科室id from dual"
                End If
                
                '医保对码检查
                If gint医保对码 = 2 Then bln提醒对码 = True
                strMsg = CheckAdviceInsure(mint险类, bln提醒对码, mlng病人ID, 1, "", Mid(strItems, 2), Left(str医嘱内容, 50))
                If strMsg <> "" Then
                    If gint医保对码 = 1 Then
                        vMsg = frmMsgBox.ShowMsgBox(strMsg & vbCrLf & vbCrLf & "要继续保存医嘱吗？", Me)
                        If vMsg = vbNo Or vMsg = vbCancel Then Exit Sub
                        If vMsg = vbIgnore Then bln提醒对码 = False
                    ElseIf gint医保对码 = 2 Then
                        MsgBox strMsg & vbCrLf & vbCrLf & "请先和相关人员联系处理，否则医嘱将不允许保存。", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    strMsg = ""
                End If
                
                '医保管控实时监测：首次输入(经过)或者更改时检查
                If blnCheckItem Then
                    If MakePriceRecord申请单("11", mlng病人ID, mlng挂号ID, strTabAdvice, strItems, rsPati!费别 & "", mlng挂号科室ID, rsPrice) Then
                        If Not gclsInsure.CheckItem(mint险类, 0, 0, rsPrice) Then
                            MsgBox "医保监测检查未通(执行Insure.CheckItem接口)，本次下达的LIS申请单不能保存。", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                End If
                strItems = ""
                
                If str附项 <> "" And str诊断 <> "" Then
                    str附项 = str诊断 & "<Split1>" & str附项
                ElseIf str附项 = "" And str诊断 <> "" Then
                    str附项 = str诊断
                End If
                
                '单据申请附项，有外键，所以先产生医嘱
                If str附项 <> "" Then
                    arrAppend = Split(str附项, "<Split1>")
                    For j = 0 To UBound(arrAppend)
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "Zl_病人医嘱附件_Insert(" & str相关ID & "," & _
                            "'" & Split(arrAppend(j), "<Split2>")(0) & "'," & Val(Split(arrAppend(j), "<Split2>")(1)) & "," & _
                            j + 1 & "," & ZVal(Split(arrAppend(j), "<Split2>")(2)) & ",'" & Replace(Split(arrAppend(j), "<Split2>")(3), "'", "''") & "'" & _
                            IIF(j = 0, ",1", "") & ")"
                        lng附项序号 = j + 1
                    Next
                End If
                
                If strDiag <> "" Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_病人诊断医嘱_Insert(" & str相关ID & ",'" & strDiag & "')"
                End If
                rsLISInfo.MoveNext
            Next
        End If
        
        '用序列产生真实的医嘱ID
        If lng假医嘱ID > 0 Then
            For j = 1 To lng假医嘱ID
                Y = zlDatabase.GetNextID("病人医嘱记录")
                If j = 1 Then
                    str医嘱IDs = ""
                    str医嘱IDs = Y
                Else
                    str医嘱IDs = str医嘱IDs & "," & Y
                End If
            Next
            varID = Split(str医嘱IDs, ",")
            
            For i = 0 To UBound(arrSQL)
                strTmp = arrSQL(i)
                If InStr(strTmp, "<FAKEID>") > 0 Then
                    j = Mid(strTmp, (InStr(strTmp, "<FAKEID>") + 8), (InStr(strTmp, "</FAKEID>") - InStr(strTmp, "<FAKEID>") - 8))
                    strTmp = Replace(strTmp, "<FAKEID>" & j & "</FAKEID>", varID(j - 1))
                    
                    If InStr(strTmp, "<FAKEID>") > 0 Then '最多替换两次
                        j = Mid(strTmp, (InStr(strTmp, "<FAKEID>") + 8), (InStr(strTmp, "</FAKEID>") - InStr(strTmp, "<FAKEID>") - 8))
                        strTmp = Replace(strTmp, "<FAKEID>" & j & "</FAKEID>", varID(j - 1))
                    End If
                    arrSQL(i) = strTmp
                End If
            Next
        End If
        
        gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
        Next
        gcnOracle.CommitTrans: blnTrans = False
        
        '刷新医嘱
        Call RefreshData
        
    End If
    '调用删除后外挂接口
    On Error Resume Next
    arrDelID = Split(strDelIDs, ",")
    For i = 0 To UBound(arrDelID)
        If Val(arrDelID(i)) <> 0 Then
            If Not gobjPlugIn Is Nothing Then
                Call gobjPlugIn.AdviceDeleted(glngSys, p门诊医嘱下达, mlng病人ID, mlng挂号ID, Val(arrDelID(i)), mint场合)
            End If
            Call zlPlugInErrH(err, "AdviceDeleted")
        End If
    Next
    If err.Number <> 0 Then err.Clear
    
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetAdivceBy申请序号(ByVal lng申请序号 As Long) As String
'功能：根据申请序号获取所有检查医嘱ID串（采集医嘱ID）
    Dim i As Long, strTmp As String
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, COL_申请序号)) = lng申请序号 Then
                If Val(.TextMatrix(i, COL_操作类型)) = 6 And .TextMatrix(i, COL_诊疗类别) = "E" Then
                    strTmp = strTmp & "," & .TextMatrix(i, COL_ID)
                End If
            End If
        Next
        GetAdivceBy申请序号 = Mid(strTmp, 2)
    End With
End Function

Private Function AdviceMakeText(ByVal str检验 As String, ByVal str采集 As String, ByVal str标本 As String) As String
'功能：产生检验医嘱的医嘱内容
    Dim i As Long, strText As String, strField As String, blnDefine As Boolean
    
    If mobjVBA Is Nothing Then
        On Error Resume Next
        Set mobjVBA = CreateObject("ScriptControl")
        err.Clear: On Error GoTo 0
        
        If Not mobjVBA Is Nothing Then
            mobjVBA.Language = "VBScript"
            Set mobjScript = New clsScript
            mobjVBA.AddObject "clsScript", mobjScript, True
        End If
    End If
               
    '确定是否定义
    blnDefine = Not mrsDefine Is Nothing And Not mobjVBA Is Nothing
    If blnDefine Then
        mrsDefine.Filter = "诊疗类别='C'"
        If mrsDefine.EOF Then
            blnDefine = False
        ElseIf Trim(NVL(mrsDefine!医嘱内容)) = "" Then
            blnDefine = False
        End If
    End If
    
    If Not blnDefine Then
        strText = str检验 & IIF(str标本 <> "", "(" & str标本 & ")", "")
    Else
        strText = mrsDefine!医嘱内容
        If InStr(strText, "[检验项目]") > 0 Then
            strField = str检验
            strText = Replace(strText, "[检验项目]", """" & strField & """")
        End If
        If InStr(strText, "[检验标本]") > 0 Then
            strField = str标本
            strText = Replace(strText, "[检验标本]", """" & strField & """")
        End If
        If InStr(strText, "[采集方法]") > 0 Then
            strField = str采集
            strText = Replace(strText, "[采集方法]", """" & strField & """")
        End If
        
        '计算医嘱内容
        On Error Resume Next
        strText = mobjVBA.Eval(strText)
        If mobjVBA.Error.Number <> 0 Then
            strText = str检验 & IIF(str标本 <> "", "(" & str标本 & ")", "")
        End If
        err.Clear: On Error GoTo 0
    End If
        
    AdviceMakeText = strText
End Function


Public Sub SetFontSize(ByVal bytSize As Byte)
'功能:设置医嘱清单的字体大小
'入参:bytSize：0-小(缺省)，1-大
    mlngFontSize = IIF(bytSize = 0, 9, 12)
    '对于vsFlexGrid控件在使用个性化设置时会加大列宽，因此在窗体初次加载是不设置字体,主要是getForm方法引起
    If Not Me.Visible Then
        vsAdvice.FontSize = mlngFontSize
        vsAppend.FontSize = mlngFontSize
        vsfAdivceDetail.FontSize = mlngFontSize
    End If
    If mvarCond.显示模式 = 0 Then
        Call Grid.SetFontSize(vsAdvice, mlngFontSize, col_内容)
    Else
        Call Grid.SetFontSize(vsAdvice, mlngFontSize, col_医嘱内容)
    End If
    
    Call Grid.SetFontSize(vsAppend, mlngFontSize)
    
    Call Grid.SetFontSize(vsfAdivceDetail, mlngFontSize)
    
    '血液执行和血液明细窗体
    If Not mobjFrmBloodList Is Nothing Then
        If mobjFrmBloodList.Visible = True Then Call mobjFrmBloodList.SetFontSize(mlngFontSize)
    End If
    
    Call SetRTFFont(0)
End Sub

Private Sub FuncPacsApply(ByVal lng医嘱ID As Long, ByRef lng申请序号 As Long)
'功能：调用检查申请单
'参数：lng医嘱ID=修改申请单时当前行的医嘱ID,lng申请序号 =当前修改行的申请序号
    Dim lngNo As Long
    
    '检查挂号病人是否超期
    If Not FuncTimeLimitCheck Then Exit Sub
    
    lngNo = ApplyOutPacs(Me, lng申请序号, mlng病人ID, mstr挂号单, lng医嘱ID, mlng挂号科室ID, mobjVBA, mobjScript, mrsDefine, mblnMoved, , mlng前提ID)
    
    If lngNo <> 0 Then Call LoadAdvice

End Sub

Private Sub FuncApplyModi()
'功能：修改申请单
    Dim strSQL As String, rsTmp As Recordset
    With vsAdvice
        '先判断是否是自定义申请单
        strSQL = "Select 文件ID From 医嘱申请单文件 Where 医嘱ID=[1] And RowNum<2"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIF(Val(.TextMatrix(.Row, COL_相关ID)) = 0, Val(.TextMatrix(.Row, COL_ID)), Val(.TextMatrix(.Row, COL_相关ID))))
        If rsTmp.RecordCount > 0 Then
            FuncApplyCustom 1, Val(rsTmp!文件ID)
        Else
                        If Val(.TextMatrix(.Row, COL_医嘱状态)) <> 1 Then
                MsgBox "不允许修改已发送的申请。", vbInformation, gstrSysName
                Exit Sub
            End If
            If Val(.TextMatrix(.Row, COL_操作类型)) = 6 And .TextMatrix(.Row, COL_诊疗类别) = "E" Then
                Call FuncLISApply(Val(.TextMatrix(.Row, COL_申请序号)))
            ElseIf .TextMatrix(.Row, COL_诊疗类别) = "D" Then
                Call FuncPacsApply(Val(.TextMatrix(.Row, COL_ID)), Val(.TextMatrix(.Row, COL_申请序号)))
            ElseIf .TextMatrix(.Row, COL_诊疗类别) = "K" Then
                If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_审核状态)) = 1 And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_检查方法)) = 1 Then
                    Call FuncBloodApply(4)
                Else
                    Call FuncBloodApply(1)
                End If
            ElseIf .TextMatrix(.Row, COL_诊疗类别) = "F" Then
                Call FuncOperationApply(1)
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncApplyView()
'功能：查看申请单
    Dim lng医嘱ID As Long
    Dim lngNo As Long
    Dim strSQL As String, rsTmp As Recordset
    With vsAdvice
        lng医嘱ID = Val(.TextMatrix(.Row, COL_ID))
        lngNo = Val(.TextMatrix(.Row, COL_申请序号))
        
        If lng医嘱ID <> 0 And lngNo <> 0 Then
            strSQL = "Select 文件ID From 医嘱申请单文件 Where 医嘱ID=[1] And RowNum<2"
            On Error GoTo errH
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIF(Val(.TextMatrix(.Row, COL_相关ID)) = 0, lng医嘱ID, Val(.TextMatrix(.Row, COL_相关ID))))
            If rsTmp.RecordCount > 0 Then
                FuncApplyCustom 2, Val(rsTmp!文件ID)
            Else
                If .TextMatrix(.Row, COL_诊疗类别) = "K" Then
                    Call FuncBloodApply(2)
                ElseIf .TextMatrix(.Row, COL_诊疗类别) = "F" Then
                    Call FuncOperationApply(2)
                ElseIf .TextMatrix(.Row, COL_诊疗类别) = "D" Then
                    '检查
                    If Val(Mid(gstrOutUseApp, 1, 1)) = 1 Then
                        Call ShowApply检查(Me, lngNo)
                    End If
                End If
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub zlPASSMap()
'功能:设置Pass VsAdvie及列映射
'注意:删除或修改下面列中数据时，请检查合理用药部件中的关联处理。
    Dim blnTmp As Boolean
    
    If mobjPassMap Is Nothing Then
        Set mobjPassMap = DynamicCreate("zlPassInterface.clsPassMap", "合理用药监测", True)
    End If
    
    If gobjPass Is Nothing Then  '83970
        blnTmp = False
    Else
        blnTmp = gobjPass.PassType <> UNPASS
    End If
    
    mblnPass = Not mobjPassMap Is Nothing And blnTmp
    
    If mblnPass Then
        With mobjPassMap
            .lngModel = PM_门诊医嘱清单
            Set .frmMain = Me
            Set .vsAdvice = vsAdvice
            Set .VSCOL = .GetVSCOL(COL_ID, COL_相关ID, COL_诊疗类别, COL_诊疗项目ID, COL_收费细目ID, col_医嘱内容, , COL_单量, , COL_用法, COL_天数, , COL_开嘱时间, COL_开嘱医生, _
                        COL_开始时间, COL_开嘱科室ID, , COL_频率, , , , COL_警示, , COL_医嘱状态, , , , , COL_执行性质, COL_标本部位, , , , , , COL_总量, , COL_医生嘱托, COL_用药目的, COL_操作类型)
            Set .PassPati = .GetPatient(mlng病人ID, mlng挂号ID)
            mblnPass = gobjPass.zlPassCheck(mobjPassMap)
        End With
    End If
End Sub

Private Sub zlPASSPati()
'功能:设置病人信息
    If Not mobjPassMap Is Nothing Then
        With mobjPassMap.PassPati
            .int婴儿 = -1 '缺省新增病人为0
            .dbl标识号 = -1
            .Dat出生日期 = -1
            .lng病人ID = mlng病人ID
            .lng主页ID = -1
            .str挂号单 = mstr挂号单
            .str床号 = ""
            .str性别 = ""
            .str姓名 = ""
        End With
    End If
End Sub

Private Sub SetAdviceColVisible()
'功能：设置医嘱表格列的可见性和表头列名
    Dim i As Long
    
    '根据显示模式调整显示列
    With vsAdvice
        .ColHidden(col_医嘱内容) = mvarCond.显示模式 = 0
        .ColHidden(col_内容) = mvarCond.显示模式 = 1
        .ColHidden(COL_皮试) = False
        .ColHidden(COL_总量) = mvarCond.显示模式 = 0
        .ColHidden(COL_单量) = mvarCond.显示模式 = 0
        .ColHidden(COL_天数) = mvarCond.显示模式 = 0
        .ColHidden(COL_频率) = mvarCond.显示模式 = 0
        .ColHidden(COL_执行时间) = mvarCond.显示模式 = 0: .Cell(flexcpData, 0, COL_执行时间) = "Detail"
        .ColHidden(COL_执行性质) = mvarCond.显示模式 = 0: .Cell(flexcpData, 0, COL_执行性质) = "Detail"
        .ColHidden(COL_开嘱时间) = mvarCond.显示模式 = 0: .Cell(flexcpData, 0, COL_开嘱时间) = "Detail"
        .ColHidden(COL_基本药物) = mvarCond.显示模式 = 0: .Cell(flexcpData, 0, COL_基本药物) = "Detail"
        .ColHidden(COL_高危药品) = True
        .ColHidden(COL_标本部位) = True
        .ColHidden(COL_收费细目ID) = True
        .ColHidden(COL_检查报告ID) = True
        .ColHidden(COL_处方审查状态) = True
        .ColHidden(COL_处方审查结果) = True
        .ColHidden(COL_处方号) = True
        .ColHidden(COL_处方打印) = True
        .ColHidden(COL_处方预览) = True
        .ColHidden(COL_并) = True
        .ColHidden(COL_标本状态) = True
        If mvarCond.医嘱 = 1 Then
            .ColHidden(COL_处方号) = False
            .ColHidden(COL_处方打印) = False
            .ColHidden(COL_处方预览) = False
        End If
        If mvarCond.过滤模式 = 0 And (mvarCond.医嘱 = 0 Or mvarCond.医嘱 = 1) Then
            .ColHidden(COL_并) = False
        End If
        If mvarCond.过滤模式 = 3 Then '如是报告卡片先藏再显示
            For i = COL_开始时间 + 1 To COL_标本部位
                .ColHidden(i) = True
            Next
            .ColHidden(COL_开始时间) = False
            .ColHidden(col_内容) = False
            .ColHidden(COL_执行科室) = False
            .ColHidden(COL_开嘱医生) = False
            .TextMatrix(0, COL_开嘱医生) = "申请医生"
            .ColHidden(COL_查阅状态) = mfrmParent Is Nothing     '电子病案查阅未传入主窗体,禁止显示查阅状态
            .ColWidth(COL_查阅状态) = 700
            .TextMatrix(0, COL_查阅状态) = "报告"
            .ColHidden(COL_标本状态) = False
            .ColWidth(COL_标本状态) = 850
        Else
            .TextMatrix(0, COL_开嘱医生) = "开嘱医生"
            .ColHidden(COL_用法) = False
            .ColHidden(COL_医生嘱托) = False
            .ColHidden(COL_查阅状态) = True
            .TextMatrix(0, COL_查阅状态) = "查阅状态"
        End If
        If mvarCond.显示模式 = 1 Then .ColHidden(COL_天数) = Not mbln天数
    End With
End Sub

Private Sub vsAdvice_KeyPress(KeyAscii As Integer)
    RaiseEvent VSKeyPress(KeyAscii)
End Sub

Private Sub SetTag一并给药(Optional ByVal lngRow As Long)
'功能：在一并给药的医嘱前加标志
    Dim i As Long
    Dim lngBg As Long, lngEd As Long
    Dim j As Long
    Dim lngStart As Long, lngEnd As Long

    If mvarCond.过滤模式 = 3 Then Exit Sub

    With vsAdvice
        If lngRow = 0 Then
            lngStart = .FixedRows
            lngEnd = .Rows - 1
        Else
            lngStart = lngRow
            lngEnd = lngRow
        End If
        For i = lngStart To lngEnd
             lngBg = -1: lngEd = -1
             If RowIn一并给药(i, lngBg, lngEd) Then
                For j = lngBg To lngEd
                    If j = lngBg Then
                        .TextMatrix(j, COL_并) = "┏"
                    ElseIf j = lngEd Then
                        .TextMatrix(j, COL_并) = "┗"
                    Else
                        .TextMatrix(j, COL_并) = "┃"
                    End If
                Next
                If lngEd <> -1 Then
                   i = lngEd + 1
                End If
            Else
                .TextMatrix(i, COL_并) = ""
            End If
        Next
    End With
End Sub

Private Function FuncTimeLimitCheck() As Boolean
'挂号病人超过期限检查，true：未超期，false：超期/mlng病人ID = 0/mblnEditable=false
    If mlng病人ID = 0 Then FuncTimeLimitCheck = False: Exit Function
    If Not mblnEditable Then FuncTimeLimitCheck = False: Exit Function
    
    If mint场合 = 0 Then
        '发送选项:0-发送为收费单,1-发送为记帐单,2-手工选择
        If Val(zlDatabase.GetPara("发送单据类型", glngSys, p门诊医嘱下达)) = 0 Then
            If BillExpend(mstr挂号单) Then
                MsgBox "该病人挂号已超过有效天数，不允许下达医嘱。", vbInformation, gstrSysName
                FuncTimeLimitCheck = False
                Exit Function
            End If
        End If
    End If
    FuncTimeLimitCheck = True
End Function

Private Function ShowAdviceRISSch(ByVal lngRow As Long, ByRef blnExist As Boolean) As Boolean
'功能：显示指定行的预约信息
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lngIdx As Long
    Dim i As Long
    
    blnExist = False
    rtfSche.Text = "": rtfSche.SelStart = 0
    
    On Error GoTo errH
    
    If Val(vsAdvice.TextMatrix(lngRow, COL_RIS预约ID)) = 0 Then Exit Function
    
    strSQL = "select 检查设备名称,To_Char(预约日期,'YYYY-MM-DD') as 预约日期," & vbNewLine & _
        "To_Char(预约开始时间,'YYYY-MM-DD HH24:MI:SS') as 预约开始时间," & vbNewLine & _
        "To_Char(预约结束时间,'YYYY-MM-DD HH24:MI:SS') as 预约结束时间," & vbNewLine & _
        "To_Char(预约开始时间段,'YYYY-MM-DD HH24:MI:SS') as 预约开始时间段," & vbNewLine & _
        "To_Char(预约结束时间段,'YYYY-MM-DD HH24:MI:SS') as 预约结束时间段,DECODE(是否调整,1,'已经预约调整','已经预约') as 预约状态" & vbNewLine & _
        "from RIS检查预约 Where 医嘱ID=[1]"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsAdvice.TextMatrix(lngRow, COL_ID)))
    If Not rsTmp.EOF Then
        With rtfSche
            For i = 0 To rsTmp.Fields.Count - 1
                .SelBold = False
                .SelText = IIF(.Text = "", "", vbCrLf) & rsTmp.Fields(i).Name & "：" & NVL(rsTmp.Fields(i).value)
                lngIdx = .Find(rsTmp.Fields(i).Name & "：", , , rtfNoHighlight Or rtfMatchCase)
                If lngIdx <> -1 Then
                    .SelStart = lngIdx
                    .SelLength = Len(rsTmp.Fields(i).Name & "：")
                    .SelBold = True
                    .SelIndent = 100
                End If
                .SelStart = Len(.Text)
            Next
            '光标定位在第一个
            lngIdx = .Find(rsTmp.Fields(0).Name & "：", 0, , rtfNoHighlight Or rtfMatchCase)
            If lngIdx <> -1 Then .SelStart = lngIdx + Len(rsTmp.Fields(0).Name & "：")
            Call SetRTFFont(4)
        End With
        blnExist = True
    End If
    ShowAdviceRISSch = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FuncAdviceRISSch()
'功能：RIS医嘱预约
    Dim lngResult As Long
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    lngResult = -1
    If HaveRIS Then
        With vsAdvice
            If InStr(",1,8,", "," & .TextMatrix(.Row, COL_医嘱状态) & ",") >= 0 Then
                lngResult = gobjRis.HISScheduling(1, Val(.TextMatrix(.Row, COL_ID)), Val(.TextMatrix(.Row, COL_诊疗项目ID)))
                If lngResult = 0 Then
                    '成功预约后更新状态
                    strSQL = "select min(预约ID) as ID from RIS检查预约 where 医嘱id=[1]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(.Row, COL_ID)))
                    .TextMatrix(.Row, COL_RIS预约ID) = rsTmp!ID & ""
                    Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col)
                End If
            Else
                MsgBox "医嘱状态为新开、已发送时，才能预约！", vbInformation, gstrSysName
            End If
        End With
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncAdviceRISDel()
'功能：RIS医嘱取消预约
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lngResult As Long
    Dim blnDo As Boolean
    
    On Error GoTo errH
    
    With vsAdvice
        If Val(.TextMatrix(.Row, COL_RIS预约ID)) <> 0 Then
            If Val(.TextMatrix(.Row, COL_医嘱状态)) = 8 Then
                strSQL = "Select Max(b.执行状态) As 结果 From 病人医嘱记录 A, 病人医嘱发送 B Where a.Id = b.医嘱id And (a.Id =[1] Or a.相关id=[1])"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(.Row, COL_ID)))
                If Not rsTmp.EOF Then
                    If Val(rsTmp!结果 & "") = 0 Then
                        blnDo = True
                    Else
                        MsgBox "该医嘱已经被执行或者部分执行不能取消预约！", vbInformation, gstrSysName
                    End If
                End If
            Else
                blnDo = True
            End If
        End If
        If blnDo Then
            If HaveRIS Then
                lngResult = gobjRis.HISSchedulingEx(Val(.TextMatrix(.Row, COL_ID)), Val(.TextMatrix(.Row, COL_RIS预约ID)))
                If lngResult = 0 Then
                    '成功能取消更改状态
                    .TextMatrix(.Row, COL_RIS预约ID) = ""
                    Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col)
                End If
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetAdviceReportIcon(ByVal lngRow As Long)
'功能：根据当前行的内容设置医嘱报告列的图标标识
'说明：注意是单行设置，不是一组设置

    With vsAdvice
        If Val(.TextMatrix(lngRow, COL_报告ID)) <> 0 Or _
            .TextMatrix(lngRow, COL_检查报告ID) <> "" Or _
            Val(.TextMatrix(lngRow, COL_RIS报告ID)) <> 0 Or _
            Val(.TextMatrix(lngRow, COL_LIS报告ID)) <> 0 Then
            
            
            If Val(.Cell(flexcpData, lngRow, COL_查阅状态)) = 0 Then
                Set .Cell(flexcpPicture, lngRow, COL_F报告) = frmIcons.imgFlag.ListImages("报告").Picture
            ElseIf Val(.Cell(flexcpData, lngRow, COL_查阅状态)) = 1 Then
                Set .Cell(flexcpPicture, lngRow, COL_F报告) = frmIcons.imgFlag.ListImages("报告已阅").Picture
            ElseIf Val(.Cell(flexcpData, lngRow, COL_查阅状态)) = 2 Then
                Set .Cell(flexcpPicture, lngRow, COL_F报告) = frmIcons.imgFlag.ListImages("报告部分阅").Picture
            End If
        Else
            If Val(.TextMatrix(lngRow, COL_RIS预约ID)) <> 0 Then
                Set .Cell(flexcpPicture, lngRow, COL_F报告) = frmIcons.imgFlag.ListImages("预约").Picture
            End If
        End If
    End With
End Sub

Private Sub FuncAdviceRISPrintSch()
'功能：RIS医嘱预约单打印
    Dim lngResult As Long
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    lngResult = -1
    If HaveRIS Then
        With vsAdvice
            If Not .TextMatrix(.Row, COL_诊疗类别) = "D" Then
                MsgBox "当前医嘱不是影像检查项目。", vbInformation, gstrSysName
                Exit Sub
            End If
            If .TextMatrix(.Row, COL_RIS预约ID) = 0 Then
                MsgBox "当前影像检查医嘱没有被预约，不能打印。", vbInformation, gstrSysName
                Exit Sub
            End If
            lngResult = gobjRis.HISPrintOneRisScheduleRpt(Val(.TextMatrix(.Row, COL_ID)))
        End With
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetAdviceReportTip(ByVal lngRow As Long) As String
'功能：获取鼠标悬浮提示字
    Dim strTmp As String
    With vsAdvice
        If Val(.TextMatrix(lngRow, COL_RIS报告ID)) <> 0 Then
            strTmp = "(RIS报告)"
        ElseIf Val(.TextMatrix(lngRow, COL_报告ID)) <> 0 Then
            strTmp = "(HIS报告)"
        ElseIf .TextMatrix(lngRow, COL_检查报告ID) <> "" Then
            strTmp = "(专业版PACS报告)"
        ElseIf Val(.TextMatrix(lngRow, COL_LIS报告ID)) <> 0 Then
            strTmp = "(三方LIS报告)"
        Else
            If Val(.TextMatrix(lngRow, COL_RIS预约ID)) <> 0 Then
                If Val(.TextMatrix(lngRow, COL_RIS预约状态)) = 0 Then
                    strTmp = "已经预约"
                Else
                    strTmp = "已经预约调整"
                End If
            End If
        End If
        If strTmp <> "" And Val(.TextMatrix(lngRow, COL_RIS预约ID)) = 0 Then
            If Val(.Cell(flexcpData, lngRow, COL_查阅状态)) = 0 Then
                strTmp = "报告未阅" & strTmp
            ElseIf Val(.Cell(flexcpData, lngRow, COL_查阅状态)) = 1 Then
                strTmp = "报告已阅" & strTmp
            ElseIf Val(.Cell(flexcpData, lngRow, COL_查阅状态)) = 2 Then
                strTmp = "报告部分已阅" & strTmp
            End If
        End If
    End With
    GetAdviceReportTip = strTmp
End Function

Private Sub FuncApplyCustom(ByVal intType As Long, ByVal lng文件ID As Long)
'功能：自定义申请单
'参数：intType=0 新增，=1修改，=2查看
    Dim lng申请序号 As Long
    Dim datTurn As Date
    Dim lngRow As Long
    Dim lng开嘱科室ID As Long
    Dim lngNo As Long
    Dim objApplyCustom As New frmApplyCustom
    
    If intType <> 2 Then
        '检查挂号病人是否超期
        If Not FuncTimeLimitCheck Then Exit Sub
        '修改时检查是否审核
        If intType = 1 Then
            If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_审核状态)) = 2 Then
                MsgBox "申请单已经审核，不允许再修改。", vbInformation, "申请单"
                intType = 2
            End If
        End If
    End If
    
    If intType <> 0 Then
         lng申请序号 = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_申请序号))
         lngRow = vsAdvice.Row
    End If
    
    If objApplyCustom.ShowMe(mfrmParent, 1, intType, mlng病人ID, mstr挂号单, 1, lng文件ID, lng申请序号, mlng挂号科室ID, IIF(mlng界面科室ID = 0, mlng挂号科室ID, mlng界面科室ID), , mrsDefine, , , 0, mclsMipModule, mlng前提ID, , mint险类) Then
        '刷新医嘱
        Call RefreshData
        '选择最后一行医嘱
        vsAdvice.Row = vsAdvice.Rows - 1
        vsAdvice.ShowCell vsAdvice.Rows - 1, col_医嘱内容
    End If
End Sub

Private Sub FuncAdviceRISModi()
'功能：调整RIS预约
    Dim lng医嘱ID As Long
    Dim lng预约ID As Long
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    With vsAdvice
        lng医嘱ID = Val(.TextMatrix(.Row, COL_ID))
        lng预约ID = Val(.TextMatrix(.Row, COL_RIS预约ID))
    End With
    
    strSQL = "select 1 from 病人医嘱发送 a where a.医嘱id=[1] and nvl(a.执行状态,0) in (0,3) and nvl(a.执行过程,0)<=2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
    If Not rsTmp.EOF Then
        If HaveRIS(False) Then
            Call gobjRis.HISReSchedule(lng医嘱ID, lng预约ID)
        End If
    Else
        MsgBox "该项目已经执行，不允许再做调整。", vbInformation, gstrSysName
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncAdviceIndexBill()
'功能：打印指引单
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lng发送号 As Long
    
    On Error GoTo errH

    With vsAdvice
        If Val(.TextMatrix(.Row, COL_医嘱状态)) = 8 Then
            strSQL = "select a.发送号 from 病人医嘱发送 a where a.医嘱id=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(.Row, COL_ID)))
            If Not rsTmp.EOF Then
                lng发送号 = Val(rsTmp!发送号 & "")
            End If
        End If
    End With
    '打印指引单
    Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1260_2", Me, "发送号=" & lng发送号, "病人ID=" & mlng病人ID, "挂号单=" & mstr挂号单, "PrintEmpty=0", 2)
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub PrintBloodReport(ByVal lngAdviceID As Long, objFrm As Object)
    '输血执行单打印
    If InitObjBlood(True) = True Then
        Call gobjPublicBlood.ShowBloodInstantRptPrint(objFrm, lngAdviceID)
    End If
End Sub

Private Sub SetAdviceIcon(ByVal lngRow As Long)
'功能：根据当前行的内容设置医嘱内容的图标标识
'说明：注意是单行设置，不是一组设置
    Dim int图标数 As Integer '医嘱内容上面的图标个数
    
    With vsAdvice
        '电子签名标识
        If Val(vsAdvice.TextMatrix(lngRow, COL_签名否)) = 1 Then
            Set vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容) = frmIcons.imgSign.ListImages("签名").Picture
            Set vsAdvice.Cell(flexcpPicture, lngRow, col_内容) = frmIcons.imgSign.ListImages("签名").Picture
            int图标数 = 1
        End If
        
        If Val(vsAdvice.TextMatrix(lngRow, COL_高危药品)) > 0 Then
            If vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容) Is Nothing Then
                Set vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容) = frmIcons.imgQuestion.ListImages("高危药品").Picture
                Set vsAdvice.Cell(flexcpPicture, lngRow, col_内容) = frmIcons.imgQuestion.ListImages("高危药品").Picture
                int图标数 = 1
            Else
                If vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容) <> frmIcons.imgQuestion.ListImages("高危药品").Picture Then
                    pictmp.Cls
                    pictmp.PaintPicture vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容), 0, 0, pictmp.Width / 2, pictmp.Height
                    pictmp.PaintPicture frmIcons.imgQuestion.ListImages("高危药品").Picture, pictmp.Width / 2, 0, pictmp.Width / 2, pictmp.Height
                    Set vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容) = pictmp.Image
                    Set vsAdvice.Cell(flexcpPicture, lngRow, col_内容) = pictmp.Image
                    int图标数 = 2
                End If
            End If
        End If
        
        '危急值图标
        If Val(vsAdvice.TextMatrix(lngRow, COL_危急值ID)) > 0 Then
            If int图标数 = 0 Then
                Set vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容) = frmIcons.imgQuestion.ListImages("危急值").Picture
                Set vsAdvice.Cell(flexcpPicture, lngRow, col_内容) = frmIcons.imgQuestion.ListImages("危急值").Picture
            ElseIf int图标数 = 1 Then
                pictmp.Cls
                pictmp.PaintPicture vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容), 0, 0, pictmp.Width / 2, pictmp.Height
                pictmp.PaintPicture frmIcons.imgQuestion.ListImages("危急值").Picture, pictmp.Width / 2, 0, pictmp.Width / 2, pictmp.Height
                Set vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容) = pictmp.Image
                Set vsAdvice.Cell(flexcpPicture, lngRow, col_内容) = pictmp.Image
                int图标数 = 2
            ElseIf int图标数 = 2 Then
                pictmp.Cls
                pictmp.Width = 720
                pictmp.PaintPicture vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容), 0, 0, 480, pictmp.Height
                pictmp.PaintPicture frmIcons.imgQuestion.ListImages("危急值").Picture, 480, 0, 240, pictmp.Height
                Set vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容) = pictmp.Image
                Set vsAdvice.Cell(flexcpPicture, lngRow, col_内容) = pictmp.Image
                pictmp.Width = 480
                int图标数 = 3
            End If
        End If
        
        '易跌倒图标
        If Val(vsAdvice.TextMatrix(lngRow, COL_易跌倒)) > 0 Then
            If int图标数 = 0 Then
                Set vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容) = frmIcons.imgQuestion.ListImages("易跌倒").Picture
                Set vsAdvice.Cell(flexcpPicture, lngRow, col_内容) = frmIcons.imgQuestion.ListImages("易跌倒").Picture
            ElseIf int图标数 = 1 Then
                pictmp.Cls
                pictmp.PaintPicture vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容), 0, 0, pictmp.Width / 2, pictmp.Height
                pictmp.PaintPicture frmIcons.imgQuestion.ListImages("易跌倒").Picture, pictmp.Width / 2, 0, pictmp.Width / 2, pictmp.Height
                Set vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容) = pictmp.Image
                Set vsAdvice.Cell(flexcpPicture, lngRow, col_内容) = pictmp.Image
                int图标数 = 2
            ElseIf int图标数 = 2 Then
                pictmp.Cls
                pictmp.Width = 720
                pictmp.PaintPicture vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容), 0, 0, 480, pictmp.Height
                pictmp.PaintPicture frmIcons.imgQuestion.ListImages("易跌倒").Picture, 480, 0, 240, pictmp.Height
                Set vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容) = pictmp.Image
                Set vsAdvice.Cell(flexcpPicture, lngRow, col_内容) = pictmp.Image
                pictmp.Width = 480
                int图标数 = 3
            ElseIf int图标数 = 3 Then
                pictmp.Cls
                pictmp.Width = 960
                pictmp.PaintPicture vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容), 0, 0, 720, pictmp.Height
                pictmp.PaintPicture frmIcons.imgQuestion.ListImages("易跌倒").Picture, 720, 0, 240, pictmp.Height
                Set vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容) = pictmp.Image
                Set vsAdvice.Cell(flexcpPicture, lngRow, col_内容) = pictmp.Image
                pictmp.Width = 480
                int图标数 = 4
            End If
        End If
    End With
End Sub

Private Sub FuncCriticalAdvice(ByVal strPar As String, ByVal blnCheck As Boolean)
'功能：设置（关联/取消）危值医嘱关联
'参数：strPar-格式：危急值ID,医嘱ID(主医嘱ID)
'      blnCheck-true 取消关系，false 设置关系
    Dim lng危急值ID As Long
    Dim lng医嘱ID As Long
    Dim lng功能 As Long
    Dim strSQL As String
    Dim lngBegin As Long
    Dim lngEnd As Long
    Dim i As Long
    Dim lngOther危急值ID As Long
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    lng功能 = IIF(blnCheck, 2, 1)
    lng危急值ID = Split(strPar, ",")(0)
    lng医嘱ID = Split(strPar, ",")(1)
    strSQL = "Zl_病人危急值医嘱_Update(" & lng功能 & "," & lng危急值ID & "," & lng医嘱ID & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    If blnCheck Then
        '同一条医嘱可关联多个危急值，取消时要进一步判断是否还有关联
        strSQL = "select a.危急值ID,a.医嘱ID from 病人危急值医嘱 a where a.医嘱ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
        If Not rsTmp.EOF Then
            lngOther危急值ID = rsTmp!危急值ID & ""
        End If
    End If
    
    
    If RowIn一并给药(vsAdvice.Row, lngBegin, lngEnd) Then
        For i = lngBegin To lngEnd
            Set vsAdvice.Cell(flexcpPicture, i, col_医嘱内容) = Nothing
            Set vsAdvice.Cell(flexcpPicture, i, col_内容) = Nothing
            If blnCheck Then
                vsAdvice.TextMatrix(i, COL_危急值ID) = lngOther危急值ID
            Else
                vsAdvice.TextMatrix(i, COL_危急值ID) = lng危急值ID
            End If
            Call SetAdviceIcon(i)
        Next
    Else
        '更新界面表格图标
        Set vsAdvice.Cell(flexcpPicture, vsAdvice.Row, col_医嘱内容) = Nothing
        Set vsAdvice.Cell(flexcpPicture, vsAdvice.Row, col_内容) = Nothing
        If blnCheck Then
            vsAdvice.TextMatrix(vsAdvice.Row, COL_危急值ID) = lngOther危急值ID
        Else
            vsAdvice.TextMatrix(vsAdvice.Row, COL_危急值ID) = lng危急值ID
        End If
        Call SetAdviceIcon(vsAdvice.Row)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetCriticalAdvice(ByRef lng医嘱ID As Long) As ADODB.Recordset
'功能：根据当前选中行的医嘱查询出与之关联的危急值记录
'参数：出参 lng医嘱ID 即当前界面上选中医嘱的主医嘱ID
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    With vsAdvice
        If Val(.TextMatrix(.Row, COL_相关ID)) <> 0 Then
            lng医嘱ID = Val(.TextMatrix(.Row, COL_相关ID))
        Else
            lng医嘱ID = Val(.TextMatrix(.Row, COL_ID))
        End If
    End With
    
    strSQL = "select a.危急值ID,a.医嘱ID from 病人危急值医嘱 a where a.医嘱ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
    
    Set GetCriticalAdvice = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetCriticalData()
'功能：获取危急值记录
    Dim strSQL As String
    On Error GoTo errH
    If mbln危急值 Then
        strSQL = "select a.id,a.危急值描述 from 病人危急值记录 a where a.挂号单=[1] order by a.报告时间 desc"
        Set mrs危急值 = zlDatabase.OpenSQLRecord(strSQL, "zlRefresh", mstr挂号单)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function FuncPathAdd() As Boolean
    Dim strSQL As String
    Dim str当前日期 As String
    Dim i As Long
    Dim lng疾病ID As Long, lng诊断ID As Long
    Dim bln中医 As Boolean
    Dim blnDo As Boolean, blnIsCancel As Boolean
    Dim blnIsSend As Boolean, blnYes As Boolean
    Dim rsTmp As ADODB.Recordset, rsPath As ADODB.Recordset
    Dim objDiagEdit As zlMedRecPage.clsDiagEdit
    
    
    '路径中的病人，当天没有生成路径项目，则先调用生成
    If mlng路径状态 = 1 And mvarCond.婴儿 <= 0 Then
        blnDo = True
        If mint场合 = 2 Then
            blnDo = zlDatabase.GetPara("医技医嘱在路径表外", glngSys, P门诊路径应用, 0) = 0
        End If
        '未评估时允许添加医嘱到昨天
        mblnNotEvaluete = Val(zlDatabase.GetPara("未评估时允许添加医嘱到昨天", glngSys, P门诊路径应用, 1)) = 1
        If blnDo Then
            If CheckPathNotEvalueteOut(mlng挂号ID, blnIsSend, str当前日期) = False Then
                If gobjPathOut Is Nothing Then
                    MsgBox "该病人当天当前阶段的路径项目未生成，不能新开医嘱。", vbInformation, gstrSysName
                ElseIf InStr(GetInsidePrivs(P门诊路径应用), ";生成路径;") = 0 Then
                    MsgBox "该病人当天当前阶段的路径项目未生成，你没有生成路径的权限，不能新开医嘱。", vbInformation, gstrSysName
                Else
                    '之前可能没有进过路径页面，需要先通过刷新接口读取初始数据
                    Call gobjPathOut.zlRefresh(mlng病人ID, mlng挂号ID, mstr挂号单, mlng挂号科室ID, mint就诊类型, mblnMoved, True)
                    Call gobjPathOut.zlExecPathSend(blnIsCancel)
                    Call LoadAdvice
                End If
                If Not blnIsCancel Then Exit Function
             Else
                If Not blnIsSend Then
                    If gobjPathOut Is Nothing Then
                        MsgBox "该病人当天当前阶段的路径项目未生成，不能新开医嘱。", vbInformation, gstrSysName
                        Exit Function
                    ElseIf InStr(GetInsidePrivs(P门诊路径应用), ";生成路径;") = 0 Then
                        MsgBox "该病人当天当前阶段的路径项目未生成，你没有生成路径的权限，不能新开医嘱。", vbInformation, gstrSysName
                        Exit Function
                    Else
                        '如果启用了参数：未评估时允许添加医嘱到昨天，则提示，否则直接进行评估生成操作
                        If mblnNotEvaluete Then
                            blnYes = MsgBox("你要添加路径外项目到''" & str当前日期 & "'?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
                        End If
                        '如果选择否，则进行评估生成操作，选择是则允许新开路径外项目到 当前日期
                        If blnYes = False Then
                            '之前可能没有进过路径页面，需要先通过刷新接口读取初始数据
                            Call gobjPathOut.zlRefresh(mlng病人ID, mlng挂号ID, mstr挂号单, mlng挂号科室ID, mint就诊类型, mblnMoved, True)
                            '没有生成，则返回false禁止新开操作
                            If Not gobjPathOut.zlExecPathSend Then
                                Call LoadAdvice
                                Exit Function
                            End If
                            Call LoadAdvice
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    FuncPathAdd = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FuncViewLisRpt()
'功能：浏览检验报告
'说明：分两种模式，先判断本次就诊是否有PDF报告
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    If mblnMoved Then
        strSQL = "select 1 from H病人医嘱记录 a,H病人医嘱报告 b,H医嘱报告内容 c where a.id=b.医嘱id and b.报告id=c.id and c.类型  in (0,2) and a.挂号单=[1]"
    Else
        strSQL = "select 1 from 病人医嘱记录 a,病人医嘱报告 b,医嘱报告内容 c where a.id=b.医嘱id and b.报告id=c.id and c.类型  in (0,2) and a.挂号单=[1]"
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr挂号单)
    
    If Not rsTmp.EOF Then
        '两个页签显示
        Call frmLisALL.ShowMe(mfrmParent, mlng病人ID, mlng挂号ID, mlng挂号科室ID, 0, p门诊医嘱下达, mMainPrivs)
    Else
        '以前的老模式
        Call InitObjLis(p门诊医生站)
        If Not gobjLIS Is Nothing Then
            gobjLIS.PatientSampleBrowse mfrmParent, mlng病人ID, mMainPrivs, mlng挂号科室ID, 0, 1
        Else
            frmLisView.ShowMe mlng病人ID, p门诊医嘱下达, mfrmParent
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncDrugRefcom()
'功能：弹出填写拒绝审核理由窗口调用合理用药部件接口
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strAdviceIDs As String
    Dim strErr As String
    
    On Error GoTo errH
    
    strSQL = "select 1 from 病人医嘱记录 a where a.挂号单=[1] and a.医嘱状态=1 and a.诊疗类别 in ('5','6') and rownum<2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr挂号单)
    If Not rsTmp.EOF Then
        '有新开的药品医嘱
        Call gobjPass.ZLPharmReviewResultOut(mfrmParent, mlng病人ID, mlng挂号ID, mstr挂号单, "", rsTmp, strErr)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Set标本状态()
'功能：对检验医嘱设置标本状态列，结果多LIS部件中返回
    Dim i As Long, str医嘱IDs As String, strMsg As String
    Dim rsAdvice As ADODB.Recordset
    Dim strIDAndRow As String, strTmp As String
    Dim lngRow As Long
    
    On Error GoTo errH
    
    If mvarCond.过滤模式 <> 3 Then Exit Sub
    Call InitObjLis(p门诊医生站)
    If gobjLIS Is Nothing Then Exit Sub
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, COL_诊疗类别) = "E" And .TextMatrix(i, COL_操作类型) = "6" And Val(.TextMatrix(i, COL_相关ID)) = 0 And Val(.TextMatrix(i, COL_医嘱状态)) = 8 Then
                str医嘱IDs = str医嘱IDs & "," & Val(.TextMatrix(i, COL_ID))
                strIDAndRow = strIDAndRow & "," & Val(.TextMatrix(i, COL_ID)) & ";" & i & "<Tab>"
            End If
        Next
        If str医嘱IDs <> "" Then
            Set rsAdvice = gobjLIS.GetSampleType(Mid(str医嘱IDs, 2), strMsg)
            If strMsg <> "" Then
                MsgBox strMsg, vbInformation, gstrSysName
            End If
            If Not rsAdvice Is Nothing Then
                rsAdvice.Filter = 0
                For i = 1 To rsAdvice.RecordCount
                    If InStr(strIDAndRow, "," & rsAdvice!医嘱ID & ";") > 0 Then
                        strTmp = Split(strIDAndRow, "," & rsAdvice!医嘱ID & ";")(1)
                        lngRow = Val(Split(strTmp, "<Tab>")(0))
                        .TextMatrix(lngRow, COL_标本状态) = rsAdvice!医嘱状态 & ""
                    End If
                    rsAdvice.MoveNext
                Next
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncViewPacsRpt()
'功能：浏览检检查报告
'说明：未处理阅读标记
    Dim blnAutoRead As Boolean
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lng医嘱ID As Long
    
    On Error GoTo errH
    Call CreateObjectPacs(mobjPublicPACS)
    If Not mobjPublicPACS Is Nothing Then
        strSQL = "select max(b.id) as 医嘱ID  from 病人医嘱报告 a,病人医嘱记录 b " & _
                " Where a.检查报告ID Is Not Null And a.医嘱ID = b.ID And b.挂号单 = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr挂号单)
        lng医嘱ID = Val(rsTmp!医嘱ID & "")
        Call mobjPublicPACS.zlDocShowReport(lng医嘱ID, , blnAutoRead, mfrmParent)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
