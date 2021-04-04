VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.12#0"; "zlIDKind.ocx"
Begin VB.Form frmTechnicStation 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "医技工作站"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11580
   Icon            =   "frmTechnicStation.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleMode       =   0  'User
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   615
      ScaleHeight     =   720
      ScaleWidth      =   1350
      TabIndex        =   32
      Top             =   6765
      Visible         =   0   'False
      Width           =   1350
      Begin XtremeReportControl.ReportControl rptNotify 
         Height          =   540
         Left            =   0
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   0
         Width           =   675
         _Version        =   589884
         _ExtentX        =   1191
         _ExtentY        =   952
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
   End
   Begin VB.PictureBox picTmp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   6120
      ScaleHeight     =   735
      ScaleWidth      =   855
      TabIndex        =   27
      Top             =   4320
      Width           =   855
   End
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   4500
      Left            =   3705
      TabIndex        =   1
      Top             =   2865
      Width           =   7770
      _Version        =   589884
      _ExtentX        =   13705
      _ExtentY        =   7937
      _StockProps     =   64
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   7485
      Width           =   11580
      _ExtentX        =   20426
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmTechnicStation.frx":058A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15901
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "病人颜色"
            TextSave        =   "病人颜色"
            Key             =   "病人颜色"
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
   Begin VB.Frame fraUD_S 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   3720
      MousePointer    =   7  'Size N S
      TabIndex        =   13
      Top             =   2730
      Width           =   3255
   End
   Begin VB.PictureBox picExec 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2040
      Left            =   3720
      ScaleHeight     =   2040
      ScaleWidth      =   7755
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   675
      Width           =   7755
      Begin VB.PictureBox picBlood 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   75
         ScaleHeight     =   390
         ScaleWidth      =   1980
         TabIndex        =   34
         Top             =   1155
         Width           =   1980
         Begin VB.Timer timBRefresh 
            Enabled         =   0   'False
            Interval        =   10
            Left            =   270
            Top             =   0
         End
         Begin XtremeDockingPane.DockingPane DkpBlood 
            Left            =   0
            Top             =   0
            _Version        =   589884
            _ExtentX        =   450
            _ExtentY        =   423
            _StockProps     =   0
         End
      End
      Begin VB.PictureBox picApplyUD_S 
         Height          =   855
         Left            =   6650
         MousePointer    =   9  'Size W E
         ScaleHeight     =   855
         ScaleWidth      =   45
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.PictureBox picApplyInfo 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   6700
         ScaleHeight     =   855
         ScaleWidth      =   1005
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1125
         Visible         =   0   'False
         Width           =   1000
         Begin RichTextLib.RichTextBox rtfAppend 
            Height          =   1395
            Left            =   0
            TabIndex        =   24
            Top             =   240
            Width           =   7200
            _ExtentX        =   12700
            _ExtentY        =   2461
            _Version        =   393217
            BorderStyle     =   0
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            TextRTF         =   $"frmTechnicStation.frx":0E1C
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
         Begin VB.Label lblApply 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "申请附项"
            Height          =   180
            Left            =   45
            TabIndex        =   25
            Top             =   30
            Width           =   720
         End
      End
      Begin VB.Frame fraDiag 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   60
         TabIndex        =   19
         Top             =   15
         Width           =   7605
         Begin VB.Label lblDiag 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "病人诊断："
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   21
            Top             =   45
            Width           =   900
         End
         Begin VB.Label lblDiag 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   1
            Left            =   1005
            TabIndex        =   20
            Top             =   45
            Width           =   90
         End
      End
      Begin VB.Frame fraExec 
         Height          =   795
         Left            =   120
         TabIndex        =   9
         Top             =   300
         Width           =   7380
         Begin VB.Label lblRec 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "记"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   465
            Left            =   6600
            TabIndex        =   28
            Top             =   330
            Width           =   495
         End
         Begin VB.Label lblAdvice 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00C00000&
            Height          =   540
            Left            =   120
            TabIndex        =   11
            Top             =   195
            Width           =   6480
         End
         Begin VB.Label lblCash 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "收"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   465
            Left            =   6825
            TabIndex        =   10
            Top             =   120
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsExec 
         Height          =   885
         Left            =   60
         TabIndex        =   12
         Top             =   1125
         Width           =   6645
         _cx             =   11721
         _cy             =   1561
         Appearance      =   1
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
   End
   Begin VB.PictureBox picPati 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6690
      Left            =   120
      ScaleHeight     =   6690
      ScaleWidth      =   3615
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   675
      Width           =   3615
      Begin XtremeReportControl.ReportControl rptPati 
         Height          =   4515
         Left            =   45
         TabIndex        =   0
         Top             =   1230
         Width           =   3375
         _Version        =   589884
         _ExtentX        =   5953
         _ExtentY        =   7964
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.CheckBox chkFilter 
         Height          =   255
         Left            =   3120
         Picture         =   "frmTechnicStation.frx":0EB9
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   "按照查找条件对病人进行过滤显示"
         Top             =   840
         Width           =   270
      End
      Begin VB.Timer timRefresh 
         Interval        =   1000
         Left            =   270
         Top             =   1155
      End
      Begin VB.Frame fraFilter 
         Caption         =   "执行状态"
         Height          =   1125
         Left            =   60
         TabIndex        =   14
         Top             =   5505
         Width           =   3480
         Begin VB.CheckBox chk执行状态 
            Caption         =   "正在执行中包含已经核对(&5)"
            Height          =   195
            Index           =   4
            Left            =   360
            TabIndex        =   26
            Top             =   846
            Width           =   2565
         End
         Begin VB.CheckBox chk执行状态 
            Caption         =   "已经执行(&4)"
            Height          =   195
            Index           =   3
            Left            =   1980
            TabIndex        =   18
            Top             =   555
            Width           =   1290
         End
         Begin VB.CheckBox chk执行状态 
            Caption         =   "尚未执行(&2)"
            Height          =   195
            Index           =   1
            Left            =   1980
            TabIndex        =   16
            Top             =   300
            Value           =   1  'Checked
            Width           =   1290
         End
         Begin VB.CheckBox chk执行状态 
            Caption         =   "拒绝执行(&1)"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   15
            Top             =   300
            Width           =   1290
         End
         Begin VB.CheckBox chk执行状态 
            Caption         =   "正在执行(&3)"
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   17
            Top             =   555
            Value           =   1  'Checked
            Width           =   1290
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   4
            Left            =   105
            Picture         =   "frmTechnicStation.frx":770B
            Top             =   846
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   3
            Left            =   1755
            Picture         =   "frmTechnicStation.frx":7C95
            Top             =   555
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   2
            Left            =   105
            Picture         =   "frmTechnicStation.frx":821F
            Top             =   555
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   1
            Left            =   1755
            Picture         =   "frmTechnicStation.frx":87A9
            Top             =   300
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   0
            Left            =   105
            Picture         =   "frmTechnicStation.frx":8D33
            Top             =   300
            Width           =   240
         End
      End
      Begin VB.ComboBox cboUnit 
         Height          =   300
         Left            =   1170
         TabIndex        =   5
         Top             =   495
         Width           =   2265
      End
      Begin MSComctlLib.ImageList img16 
         Left            =   345
         Top             =   1320
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   8
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTechnicStation.frx":92BD
               Key             =   "未执行"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTechnicStation.frx":9857
               Key             =   "已执行"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTechnicStation.frx":9DF1
               Key             =   "拒绝执行"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTechnicStation.frx":A38B
               Key             =   "正在执行"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTechnicStation.frx":A925
               Key             =   "已报到"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTechnicStation.frx":AEBF
               Key             =   "CheckCol"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTechnicStation.frx":B459
               Key             =   "Path"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTechnicStation.frx":B9F3
               Key             =   "已核对"
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   1170
         TabIndex        =   3
         Text            =   "cboDept"
         Top             =   120
         Width           =   2265
      End
      Begin zlIDKind.PatiIdentify PatiIdentify 
         Height          =   270
         Left            =   825
         TabIndex        =   29
         Top             =   840
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindStr       =   $"frmTechnicStation.frx":BF8D
         BeginProperty IDKindFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoSize        =   -1  'True
         IDKindAppearance=   0
         ShowPropertySet =   -1  'True
         DefaultCardType =   "就诊卡"
         IDKindWidth     =   555
         FindPatiShowName=   0   'False
         HiddenMoseRightKey=   0   'False
         BeginProperty CardNoShowFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblFind 
         Caption         =   "查找(F3)"
         Height          =   255
         Left            =   75
         TabIndex        =   30
         Top             =   870
         Width           =   975
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人病区(&U)"
         Height          =   180
         Left            =   135
         TabIndex        =   4
         Top             =   555
         Width           =   990
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医技科室(&D)"
         Height          =   180
         Left            =   135
         TabIndex        =   2
         Top             =   180
         Width           =   990
      End
   End
   Begin MSComctlLib.ImageList imgPati 
      Left            =   1185
      Top             =   75
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C054
            Key             =   "Pati"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C0B2
            Key             =   "Meet"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C110
            Key             =   "MeetFinish"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C16E
            Key             =   "Notify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C1CC
            Key             =   "等待审查"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C22A
            Key             =   "拒绝审查"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C288
            Key             =   "正在审查"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C2E6
            Key             =   "正在抽查"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C344
            Key             =   "审查反馈"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C3A2
            Key             =   "抽查反馈"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C400
            Key             =   "审查整改"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C45E
            Key             =   "抽查整改"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C4BC
            Key             =   "未导入"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C51A
            Key             =   "变异结束"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C578
            Key             =   "正常结束"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C5D6
            Key             =   "不符合"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C634
            Key             =   "执行中"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C692
            Key             =   "Child"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C6F0
            Key             =   "单病种"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnicStation.frx":C74E
            Key             =   "Out"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   270
      Top             =   180
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmTechnicStation.frx":C7AC
      Left            =   705
      Top             =   210
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmTechnicStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const COLOR_LOST = &HFFEBD7
Private Const COLOR_FOCUS = &HFFCC99
Private Enum PATIREPORT_COLUMN
    col_综合状态 = 0
    col_选择 = 1
    col_路径 = 2
    col_执行状态 = 3
    col_图标 = 4
    col_来源 = 5    '门诊留观病人，value='住院',caption='门诊'
    col_单据号 = 6
    col_紧急 = 7
    col_姓名 = 8
    col_内容 = 9
    col_险类 = 10
    col_科室 = 11
    col_标识号 = 12
    col_床号 = 13
    col_费别 = 14
    col_要求时间 = 15
    col_发送时间 = 16
    col_执行间 = 17
    col_性别 = 18
    col_年龄 = 19
    col_完成人 = 20
    col_完成时间 = 21
    '隐藏列
    col_执行科室 = 22
    col_病人Id = 23
    col_主页ID = 24
    col_挂号单 = 25
    col_挂号ID = 26 '
    col_婴儿 = 27 '
    col_就诊卡号 = 28
    col_身份证号 = 29
    col_IC卡号 = 30
    col_医保号 = 31
    col_病区id = 32
    col_出院日期 = 33
    COL_状态 = 34
    col_医嘱ID = 35
    col_相关ID = 36
    col_发送号 = 37
    COL_诊疗类别 = 38
    col_执行过程 = 39
    col_执行安排 = 40
    col_记录性质 = 41
    COL_数据转出 = 42
    col_文件ID = 43
    col_报告项 = 44
    col_报告ID = 45
    col_病人类型 = 46
    col_门诊记帐 = 47
    col_开单人 = 48
    col_审查 = 49
    col_类型 = 50
    COL_操作类型 = 51
    COL_核对人 = 52
    col_审核标志 = 53
    COL_接收时间 = 54
    col_结算模式 = 55
    COL_诊疗项目ID = 56
    col_期效 = 57
    COL_执行分类 = 58
    COL_主页挂号ID = 59  '病案主页.挂号ID
    COL_附加标志 = 60    '病人医嘱记录.附加标志
End Enum

Private Enum NOTIFYREPORT_COLUMN
    c_图标 = 0
    C_病人ID = 1
    C_主页ID = 2
    
    c_姓名 = 3
    C_状态 = 4
    
    C_消息 = 5
    C_序号 = 6
    C_日期 = 7
    C_业务 = 8
    
End Enum

Private mblnShowBed As Boolean

Private Enum PATI_TYPE
    pt我的 = 1
    pt在院 = 2
    pt预出 = 3
    pt出院 = 4
    pt死亡 = 5
    pt会诊 = 6
    pt最近转出 = 7
End Enum

Private Enum Msg_Type '消息提醒类别
    m销帐申请 = 1
    m待安排 = 2
    m血袋回收 = 3
End Enum

'子窗体对象定义
Private mclsEMR As Object  '新版病历zlRichEMR.clsDockEMR
Private WithEvents mclsInAdvices As zlPublicAdvice.clsDockInAdvices
Attribute mclsInAdvices.VB_VarHelpID = -1
Private WithEvents mclsOutAdvices As zlPublicAdvice.clsDockOutAdvices
Attribute mclsOutAdvices.VB_VarHelpID = -1
Private WithEvents mclsExpenses As zlPublicExpense.clsDockExpense
Attribute mclsExpenses.VB_VarHelpID = -1
Private WithEvents mclsInEPRs As zlRichEPR.cDockInEPRs
Attribute mclsInEPRs.VB_VarHelpID = -1
Private WithEvents mclsOutEPRs As zlRichEPR.cDockOutEPRs
Attribute mclsOutEPRs.VB_VarHelpID = -1
Private mclsTendEPRs As zlRichEPR.cDockInTendEPRs
Attribute mclsTendEPRs.VB_VarHelpID = -1
Private WithEvents mclsTends As zlRichEPR.cDockInTends
Attribute mclsTends.VB_VarHelpID = -1
Private WithEvents mclsTendsNew As zl9TendFile.clsTendFile    '新版护理记录
Attribute mclsTendsNew.VB_VarHelpID = -1
Private mcolSubForm As Collection
Private mfrmActive As Form
Private mobjFrmBloodExe As Object
Private WithEvents mobjIDCard As clsIDCard '身份证对象
Attribute mobjIDCard.VB_VarHelpID = -1
Private mobjICCard As Object 'IC卡对象
Private mobjAppendBill As Object '附费相关的对象
Private mbln附费按钮 As Boolean '添加新附费窗口弹出功能按钮
Private mstr密码 As String

'病历报告对象
Private WithEvents mclsEPRReport As zlRichEPR.cEPRDocument
Attribute mclsEPRReport.VB_VarHelpID = -1

'医疗卡
Private mobjSquareCard As Object      '卡结算对象
Private mstrCardKind As String        '卡结算对象返回的可用的医疗卡
Private Enum CardProperty
    CP短名 = 0
    CP全名 = 1
    CP可读卡 = 2
    CP卡类别ID = 3
    CP卡号长度 = 4
    CP缺省类别 = 5
    CP存在帐户 = 6
    CP卡号密文显示 = 7
End Enum

Private mintFindType As Integer '1-就诊卡,2-标识号（门诊号）,3-单据号,4-姓名,5-二代身份证身份证,6-IC卡,7-医保号
Private mstrFindType As String '存储当前查找类型名称
Private mblnFindTypeEnabled As Boolean
Private mblnFilter As Boolean '查找是否以过滤模式进行显示
Private mblnFilterEnabled As Boolean

'过滤条件变量
Private Type FilterCond
    Begin As Date
    End As Date
    NO As String
    科室ID As Long
    来源 As String
    本次 As Boolean
    期效 As Integer
    标识号 As String
    就诊卡 As String
    姓名 As String
    身份证 As String
    IC卡号 As String
    医保号 As String
    开单人 As String
    病人ID As Long
End Type
Private mvarCond As FilterCond
Private mbln只显已收费 As Boolean
Private mbln只显已收费Enabled As Boolean
Private mbln他科执行 As Boolean
Private mstr状态 As String  '1-5位分别表示：拒绝执行,尚未执行,正在执行,已经执行,正在执行且已经核对
Private mstr开单人 As String
Private mstr过滤条件 As String

'本地参数变量
Private mblnExeLog As Boolean
Private mbln皮试验证 As Boolean
Private mintRefresh As Integer
Private mstrRoom As String
Private mstr诊疗类别 As String
Private mstr治疗类别 As String

'其它窗体变量
Private mstrPrivs As String
Private mlngModul As Long
Private mlngDept As Long
Private mstrDeptNode As String '当前医技科室所属的站点
Private mstrPrePati As String
Private mlng病人ID As Long, mlng主页ID As Long, mstr挂号单 As String

Private mblnFirstLoad As Boolean '判断是否是第一次加载
Private mblnMoved As Boolean
Private mlngFontSize As Long  '字体大小
Private mblnReturn As Boolean      'cboDept/cboUnit的回车按键
Private mbln血透室 As Boolean
Private mbln产科 As Boolean

Private mbyt病人审核方式 As Byte '49501:病人审核方式:0-未审核不允许结帐，缺省为0;1-审核时不许调整费用和医嘱（包含医嘱调整和费用调整）
Private mstrNotify As String '提醒的类型
Private mintDay As Integer '提醒多少天内的消息
Private mintMin As Integer '提醒自动刷新间隔(分钟)
Private mclsMsg As clsCISMsg
Private mrsMsg As ADODB.Recordset
Private mbln消息语音 As Boolean
Private mbln未收费完成 As Boolean '参数：未收费完成
Private mstrBloodControlIDs  As String '用血执行菜单ID串

'病历
Private mlngType As Long
Private mlngState As TYPE_PATI_State
Private mlngInIndex As Long
Private mlngOutIndex As Long
Private mlngNurIndex As Long
Private mlngNewNurIndex As Long
Private mlngNurEMRIndex As Long
Private mlngNewIndex As Long '新版病历选项卡，如果未添加为 －1
Private mblnNewNurRecord As Boolean     '血透室是否使用新版护理记录

Private mblnIsInit As Boolean

Private COLExec As New Collection

Private mbytSize As Byte '字体大小 0-小字体（9号），1-大字体（12号）
Private mblnTabTmp As Boolean
Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mobjKernel As zlPublicAdvice.clsPublicAdvice         '临床核心部件
Private mclsPExp As zlPublicExpense.clsPublicExpense

Private Function ExchangeAdvice(ByVal blnClinic As Boolean) As Boolean
'功能：切换显示门诊/住院医嘱页卡
'参数：blnClinic=是否显示门诊医嘱页
'返回：是否进行了切换选择
    Dim blnSel As Boolean
    Dim blnOld As Boolean, intIdx As Integer
    
    If Not tbcSub.Selected Is Nothing Then
        blnSel = tbcSub.Selected.Tag Like "*医嘱"
    End If
    
    For intIdx = 0 To tbcSub.ItemCount - 1
        If tbcSub(intIdx).Tag = "门诊医嘱" Then
            If tbcSub(intIdx).Visible <> blnClinic Then
                tbcSub(intIdx).Visible = blnClinic
                If blnSel And blnClinic Then
                    tbcSub(intIdx).Selected = True
                    ExchangeAdvice = True
                End If
            End If
        ElseIf tbcSub(intIdx).Tag = "住院医嘱" Then
            If tbcSub(intIdx).Visible <> Not blnClinic Then
                tbcSub(intIdx).Visible = Not blnClinic
                If blnSel And Not blnClinic Then
                    tbcSub(intIdx).Selected = True
                    ExchangeAdvice = True
                End If
            End If
        End If
    Next
End Function

Private Sub cboDept_GotFocus()
    Call zlControl.TxtSelAll(cboDept)
End Sub

Private Sub cboDept_Validate(Cancel As Boolean)
    If mblnReturn Then
        mblnReturn = False
    Else
        Call Cbo.SetIndex(cboDept.hwnd, Val(cboDept.Tag))
    End If
End Sub

Private Sub cboDept_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    If cboDept.ListIndex <> -1 Then cboDept.Tag = cboDept.ListIndex
    mblnReturn = False
    If KeyAscii = 13 Then
        mblnReturn = True
        KeyAscii = 0
        '来源部门不限制站点
        If cboDept.Text <> "" Then
            strSQL = "Select A.ID,A.编码,A.名称" & _
                " From 部门表 A,部门性质说明 B" & _
                " Where A.ID=B.部门ID And B.服务对象 In(1,2,3) And B.工作性质 IN('检查','检验','手术','治疗','营养')" & _
                " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is Null)" & _
                IIf(mstrDeptNode <> "", " And (A.站点 = [3] Or A.站点 is Null)", "") & _
                " And (A.编码 Like [1] Or A.简码 Like [2] Or A.名称 Like [2])" & _
                " Order by A.编码"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UCase(cboDept.Text) & "%", gstrLike & UCase(cboDept.Text) & "%", mstrDeptNode)
            If Not rsTmp.EOF Then
                Call Cbo.SeekIndex(cboDept, rsTmp!ID)
            Else
                cboDept.ListIndex = Val(cboDept.Tag)
            End If
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            cboDept.ListIndex = Val(cboDept.Tag)
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboUnit_Click()
'功能：重新读取病人
    If cboUnit.ListIndex = -1 Then Exit Sub
    mblnReturn = True
    
    If Val(cboUnit.Tag) = cboUnit.ListIndex Then Exit Sub
    cboUnit.Tag = cboUnit.ListIndex
 
    Call LoadPatients
End Sub

Private Sub cboUnit_GotFocus()
    Call zlControl.TxtSelAll(cboUnit)
End Sub

Private Sub cboUnit_Validate(Cancel As Boolean)
    If mblnReturn Then
        mblnReturn = False
    Else
        Call Cbo.SetIndex(cboUnit.hwnd, Val(cboUnit.Tag))
    End If
End Sub

Private Sub cboUnit_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    mblnReturn = False
    If KeyAscii = 13 Then
        mblnReturn = True
        KeyAscii = 0
        '来源部门不限制站点
        If cboUnit.Text <> "" Then
            strSQL = "Select A.ID,A.编码,A.名称" & _
                " From 部门表 A,部门性质说明 B" & _
                " Where A.ID=B.部门ID And B.服务对象 In(1,2,3) And B.工作性质='护理'" & _
                " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is Null)" & _
                IIf(mstrDeptNode <> "", " And (A.站点 = [3] Or A.站点 is Null)", "") & _
                " And (A.编码 Like [1] Or A.简码 Like [2] Or A.名称 Like [2])" & _
                " Order by A.编码"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UCase(cboUnit.Text) & "%", gstrLike & UCase(cboUnit.Text) & "%", mstrDeptNode)
            If Not rsTmp.EOF Then
                Call Cbo.SeekIndex(cboUnit, rsTmp!ID)
            Else
                cboUnit.ListIndex = Val(cboUnit.Tag)
            End If
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub chkFilter_Click()
    mblnFilter = chkFilter.Value = 1
    PatiIdentify.Text = ""
    If PatiIdentify.Visible And PatiIdentify.Enabled Then PatiIdentify.SetFocus
    
    '切换时清除条件重新刷新清单
    Call ClearPatiCond
    Call LoadPatients
End Sub

Private Sub chk执行状态_Click(Index As Integer)
    If Visible Then
        mstr状态 = zlStr.SetBit(mstr状态, Index + 1, chk执行状态(Index).Value)
        chk执行状态(4).Enabled = chk执行状态(2).Value = 1
        Call LoadPatients
    End If
End Sub

Private Sub DkpBlood_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        If Not mobjFrmBloodExe Is Nothing Then
            Item.Handle = mobjFrmBloodExe.hwnd
        End If
    End If
End Sub

Private Sub Form_Activate()
    mblnFirstLoad = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '读卡
    PatiIdentify.ActiveFastKey
End Sub

Private Sub Form_Load()
    Dim objPane As Pane, strTab As String, intIdx As Integer, i As Long
    Dim strKey As String, intType As Integer
    Dim objControl As CommandBarControl
    Dim arrTmp As Variant, strTmp As String
    
    Set mclsMipModule = New zl9ComLib.clsMipModule
    Call mclsMipModule.InitMessage(glngSys, p医技工作站, GetInsidePrivs(p医技工作站))
    Call AddMipModule(mclsMipModule)
    Set mobjKernel = New zlPublicAdvice.clsPublicAdvice
    On Error Resume Next
    Set mobjAppendBill = CreateObject("ZlSoft.HIS.Charge.AppendCharge")
    err.Clear: On Error GoTo 0
    mbln附费按钮 = False
    mstr密码 = ""
    Set mclsPExp = New zlPublicExpense.clsPublicExpense
    mstrPrivs = gstrPrivs
    mlngModul = glngModul
    
    mblnFirstLoad = True
    picApplyUD_S.Left = 9200
    
    mstrBloodControlIDs = Join(Array(conMenu_Manage_Complete, conMenu_Manage_Undone, conMenu_Manage_ThingAdd, conMenu_Manage_ThingModi, conMenu_Manage_ThingDel, _
        conMenu_Manage_ThingAudit, conMenu_Manage_ThingDelAudit, conMenu_Manage_ThingAudit * 100# + 1, conMenu_Manage_ThingDelAudit * 100# + 1), ",")

    Call zlCommFun.SetWindowsInTaskBar(Me.hwnd, gblnShowInTaskBar)
    Call InitExecTable
        
    mstrRoom = zlDatabase.GetPara("执行间范围", glngSys, p医技工作站)
    '字体设置
    mbytSize = zlDatabase.GetPara("字体", glngSys, p医技工作站, "0")
    
    '提醒刷新设置
    mstrNotify = zlDatabase.GetPara("自动刷新医嘱类型", glngSys, p医技工作站, "000")
    mintDay = Val(zlDatabase.GetPara("自动刷新医嘱天数", glngSys, p医技工作站, 1))
    mintMin = Val(zlDatabase.GetPara("自动刷新医嘱间隔", glngSys, p医技工作站))
    mbln未收费完成 = (Val(zlDatabase.GetPara("未收费完成", glngSys, p医技工作站)) = 1)
    mbln消息语音 = Val(zlDatabase.GetPara("启用语音提示", glngSys, p医技工作站)) = 1
    '血透室是否使用新版护理记录
    mblnNewNurRecord = (Val(zlDatabase.GetPara("血透室书写新版护理记录", glngSys, p医技工作站)) = 1)

    '过滤条件初始
    '-----------------------------------------------------
    mvarCond.本次 = Val(zlDatabase.GetPara("只显示本次住院项目", glngSys, p医技工作站, "1")) = 1
    mstr过滤条件 = IIf(Val(zlDatabase.GetPara("病人过滤方式", glngSys, p医技工作站)) = 1, "发送时间", "首次时间")
    
    '病人来源
    strKey = zlDatabase.GetPara("病人来源", glngSys, p医技工作站, "111")
    mvarCond.来源 = ""
    If Not (Val(Mid(strKey, 1, 1)) = 1 And Val(Mid(strKey, 2, 1)) = 1 And Val(Mid(strKey, 3, 1)) = 1) Then
        If Val(Mid(strKey, 1, 1)) = 1 Then mvarCond.来源 = mvarCond.来源 & ",1"
        If Val(Mid(strKey, 2, 1)) = 1 Then mvarCond.来源 = mvarCond.来源 & ",2"
        If Val(Mid(strKey, 3, 1)) = 1 Then mvarCond.来源 = mvarCond.来源 & ",4"
        mvarCond.来源 = Mid(mvarCond.来源 & ",3", 2)
    End If
    Call SetUnitVisible

    '医嘱期效
    strKey = zlDatabase.GetPara("医嘱期效", glngSys, p医技工作站, "11")
    mvarCond.期效 = 0
    If Not (Val(Mid(strKey, 1, 1)) = 1 And Val(Mid(strKey, 2, 1)) = 1) Then
        If Val(Mid(strKey, 1, 1)) = 1 Then
            mvarCond.期效 = 1
        ElseIf Val(Mid(strKey, 2, 1)) = 1 Then
            mvarCond.期效 = 2
        End If
    End If
    
    '其它条件初始
    mvarCond.Begin = CDate(0)
    mvarCond.End = CDate(0)
    mvarCond.科室ID = 0
    
    Call ClearPatiCond
    
    '开单人
    mstr开单人 = zlDatabase.GetPara("开单人", glngSys, p医技工作站, "")
    If mstr开单人 <> "" Then mvarCond.开单人 = mstr开单人
    
    '一卡通部件初始，须在tbcSub_SelectedChanged之前，以便传递给医嘱部件
    'zlGetIDKindStr中会自动补齐为至少8位属性
    mstrCardKind = "就|就诊卡|0|0|8|0|0|0;门|标识号|0|0|0|0|0|0;单|单据号|0|0|0|0|0|0;姓|姓名|0|0|0|0|0|0;身|二代身份证|0|0|0|0|0|0;ＩＣ|ＩＣ卡|1|0|0|0|0|0;医|医保号|0|0|0|0|0|0"
    On Error Resume Next
    Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    err.Clear: On Error GoTo 0
    
    
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    'DockingPane
    '-----------------------------------------------------
    Me.DkpMain.SetCommandBars Me.cbsMain
    Me.DkpMain.Options.UseSplitterTracker = False '实时拖动
    Me.DkpMain.Options.ThemedFloatingFrames = True
    Me.DkpMain.Options.AlphaDockingContext = True
    Set objPane = Me.DkpMain.CreatePane(1, IIf(mbytSize = 0, 280, 300), 400, DockLeftOf, Nothing)
    objPane.Title = "执行病人列表"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    Set objPane = Me.DkpMain.CreatePane(2, 310, 100, DockBottomOf, objPane)
    objPane.Title = "消息提醒"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    
    '待发血液布局
    If gbln血库系统 = True Then
        With DkpBlood
            .Options.UseSplitterTracker = False '实时拖动
            .Options.ThemedFloatingFrames = True
            .Options.AlphaDockingContext = True
            .Options.HideClient = True
            
            Set objPane = .CreatePane(1, 100, 100, DockLeftOf, Nothing)
            objPane.Title = "输血执行登记"
            objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable
        End With
        If InitObjBlood = True Then
            Set mobjFrmBloodExe = gobjPublicBlood.zlGetBloodExec
            mobjFrmBloodExe.IsShowExec = True
        End If
    End If
    
    'TabControl
    '-----------------------------------------------------
    Call zlCommFun.SetWindowsInTaskBar(Me.hwnd, True)
    If GetInsidePrivs(p新版门诊病历, True) <> "" Or GetInsidePrivs(p新版住院病历, True) <> "" Then
        Set mclsEMR = DynamicCreate("zlRichEMR.clsDockEMR", "电子病历")
        If Not mclsEMR Is Nothing Then
            If Not mclsEMR.Init(gobjEmr, gcnOracle, glngSys) Then
                Set mclsEMR = Nothing
            End If
        End If
    End If
    Set mclsOutAdvices = New zlPublicAdvice.clsDockOutAdvices
    Set mclsInAdvices = New zlPublicAdvice.clsDockInAdvices
    Set mclsExpenses = New zlPublicExpense.clsDockExpense
    Set mclsInEPRs = New zlRichEPR.cDockInEPRs
    Set mclsOutEPRs = New zlRichEPR.cDockOutEPRs
    Set mclsTends = New zlRichEPR.cDockInTends
    Set mclsTendsNew = New zl9TendFile.clsTendFile
    Set mclsTendEPRs = New zlRichEPR.cDockInTendEPRs
    Call mclsTendsNew.InitTendFile(gcnOracle, glngSys)
    
    If Not mclsExpenses Is Nothing Then
        Call mclsExpenses.zlInitCommon(glngSys, gcnOracle, gstrDBUser)
    End If
    
    Set mcolSubForm = New Collection
    If Not mclsEMR Is Nothing Then
        mcolSubForm.Add mclsEMR.zlGetForm, "_新病历"
    End If
    mcolSubForm.Add mclsExpenses.zlGetForm, "_医嘱附费"
    mcolSubForm.Add mclsOutAdvices.zlGetForm, "_门诊医嘱"
    mcolSubForm.Add mclsInAdvices.zlGetForm, "_住院医嘱"
    mcolSubForm.Add mclsInEPRs.zlGetForm, "_住院病历"
    mcolSubForm.Add mclsOutEPRs.zlGetForm, "_门诊病历"
    mcolSubForm.Add mclsTends.zlGetForm, "_护理"
    mcolSubForm.Add mclsTendsNew.zlGetForm, "_新版护理"
    mcolSubForm.Add mclsTendEPRs.zlGetForm, "_护理病历"
    
    With Me.tbcSub
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        '绑定子窗体时会Form_Load，且自动选中第一个加入的卡片
        '如果设置当前卡片隐藏,则不会自动切换选择,但显示内容未变
        '任意指定索引号无效，最终变为0-N，只是可能改变加入顺序。
        If GetInsidePrivs(p医嘱附费管理, True) <> "" Then
            If mobjAppendBill Is Nothing Then
                .InsertItem(intIdx, "医嘱附加费用", picTmp.hwnd, 0).Tag = "医嘱附费": intIdx = intIdx + 1
            Else
                mbln附费按钮 = True
            End If
        End If
        If GetInsidePrivs(p门诊医嘱下达, True) <> "" Then
            .InsertItem(intIdx, "门诊医嘱", picTmp.hwnd, 0).Tag = "门诊医嘱": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(p住院医嘱下达, True) <> "" Then
            .InsertItem(intIdx, "住院医嘱", picTmp.hwnd, 0).Tag = "住院医嘱": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(p住院病历管理, True) <> "" Then
            .InsertItem(intIdx, "住院病历", picTmp.hwnd, 0).Tag = "住院病历": intIdx = intIdx + 1
            .Item(intIdx - 1).Visible = False
            mlngInIndex = intIdx - 1
        End If
        mlngNewIndex = -1
        If (GetInsidePrivs(p新版门诊病历, True) <> "" Or GetInsidePrivs(p新版住院病历, True) <> "") And Not mclsEMR Is Nothing Then
            .InsertItem(intIdx, "电子病历", picTmp.hwnd, 0).Tag = "新病历": intIdx = intIdx + 1
            .Item(intIdx - 1).Visible = False
            mlngNewIndex = intIdx - 1
        End If
        If GetInsidePrivs(p门诊病历管理, True) <> "" Then
            .InsertItem(intIdx, "门诊病历", picTmp.hwnd, 0).Tag = "门诊病历": intIdx = intIdx + 1
            .Item(intIdx - 1).Visible = False
            mlngOutIndex = intIdx - 1
        End If
        
        If GetInsidePrivs(p护理记录管理, True) <> "" Then
            .InsertItem(intIdx, "护理记录", picTmp.hwnd, 0).Tag = "护理": intIdx = intIdx + 1
            .Item(intIdx - 1).Visible = False
            mlngNurIndex = intIdx - 1
        End If
        
        If GetInsidePrivs(p护理记录管理, True) <> "" Then
            .InsertItem(intIdx, "护理记录", picTmp.hwnd, 0).Tag = "新版护理": intIdx = intIdx + 1
            .Item(intIdx - 1).Visible = False
            mlngNewNurIndex = intIdx - 1
            .InsertItem(intIdx, "护理病历", picTmp.hwnd, 0).Tag = "护理病历": intIdx = intIdx + 1
            .Item(intIdx - 1).Visible = False
            mlngNurEMRIndex = intIdx - 1
        End If
        
        '外挂部件中卡片
        Call CreatePlugInOK(p医技工作站)
        If Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            strTmp = gobjPlugIn.GetFormCaption(glngSys, p医技工作站)
            Call zlPlugInErrH(err, "GetFormCaption")
            If strTmp <> "" Then
                arrTmp = Split(strTmp, ",")
                For i = 0 To UBound(arrTmp)
                    strTmp = arrTmp(i)
                    
                    mcolSubForm.Add gobjPlugIn.GetForm(glngSys, p医技工作站, strTmp), "_" & strTmp
                    .InsertItem(intIdx, strTmp, mcolSubForm("_" & strTmp).hwnd, 0).Tag = strTmp: intIdx = intIdx + 1
                    Call zlPlugInErrH(err, "GetForm")
                Next
            End If
            err.Clear: On Error GoTo 0
        End If
        
        If .ItemCount = 0 Then
            MsgBox "你没有使用医技工作站的权限(请检查是否具备:医嘱附加费用,门诊医嘱下达,住院医嘱下达的权限之一)。", vbInformation, gstrSysName
            Unload Me: Exit Sub
        End If
        
        '恢复上次选择的卡片
        strTab = zlDatabase.GetPara("医护功能", glngSys, p医技工作站)
        If mvarCond.来源 = "2,3" Then
            Call ExchangeAdvice(False) '缺省显示住院的
        Else
            Call ExchangeAdvice(True) '缺省显示门诊的
        End If
        For intIdx = 0 To tbcSub.ItemCount - 1
            If tbcSub(intIdx).Visible And tbcSub(intIdx).Tag = strTab Then Exit For
        Next
        If intIdx <= tbcSub.ItemCount - 1 Then
            strTab = .Item(intIdx).Tag
            .Item(intIdx).Tag = "" '避免激活事件
            .Item(intIdx).Selected = True
            .Item(intIdx).Tag = strTab
        Else
            .Item(0).Selected = True '新建时就自动选中了这个,不会再激活事件
        End If
        '只加载选择的子窗体
        Call tbcSub_SelectedChanged(.Selected)
    End With
    
     '其它界面设置
    Call InitReportColumn
    picPati.BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
    picExec.BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
    
    '读取界面数据
    '-----------------------------------------------------
    mlngDept = -1
    mbln血透室 = False
    mbln产科 = False
    mstrDeptNode = ""
    mstrPrePati = ""
    mstr状态 = zlDatabase.GetPara("执行状态", glngSys, p医技工作站, "01101", _
        Array(chk执行状态(0), chk执行状态(1), chk执行状态(2), chk执行状态(3), chk执行状态(4)), InStr(mstrPrivs, "参数设置") > 0) '缺省显示未执行、正在执行
    chk执行状态(0).Value = Val(Mid(mstr状态, 1, 1))
    chk执行状态(1).Value = Val(Mid(mstr状态, 2, 1))
    chk执行状态(2).Value = Val(Mid(mstr状态, 3, 1))
    chk执行状态(3).Value = Val(Mid(mstr状态, 4, 1))
    chk执行状态(4).Value = Val(Mid(mstr状态, 5, 1))
    chk执行状态(4).Enabled = chk执行状态(2).Value = 1
    If Val(gstr医嘱核对) = 0 Then
        chk执行状态(4).Visible = False
        Image1(4).Visible = False
    End If
    
    mintFindType = Val(zlDatabase.GetPara("病人查找方式", glngSys, p医技工作站, "1", , , intType))
    mblnFindTypeEnabled = Not ((intType = 3 Or intType = 15) And InStr(mstrPrivs, "参数设置") = 0)
    mblnFilter = Val(zlDatabase.GetPara("过滤显示模式", glngSys, p医技工作站, , , , intType)) <> 0
    mblnFilterEnabled = Not ((intType = 3 Or intType = 15) And InStr(mstrPrivs, "参数设置") = 0)
    mbln只显已收费 = Val(zlDatabase.GetPara("只显示已收费的病人", glngSys, p医技工作站, , , , intType)) <> 0
    mbln只显已收费Enabled = Not ((intType = 3 Or intType = 15) And InStr(mstrPrivs, "参数设置") = 0)
    mbln他科执行 = False
    mblnExeLog = Val(zlDatabase.GetPara("记录执行情况", glngSys, p医技工作站, "0")) <> 0
    mbln皮试验证 = Val(zlDatabase.GetPara("皮试验证身份", glngSys, p医技工作站)) <> 0
    
    mstr诊疗类别 = zlDatabase.GetPara("诊疗类别", glngSys, p医技工作站)
    mstr治疗类别 = zlDatabase.GetPara("治疗类别", glngSys, p医技工作站)
    mbyt病人审核方式 = Val(zlDatabase.GetPara(185, glngSys))
        
    Set mobjIDCard = New clsIDCard
    Call mobjIDCard.SetParent(Me.hwnd)
    On Error Resume Next
    Set mobjICCard = CreateObject("zlICCard.clsICCard")
    err.Clear: On Error GoTo 0
    
    
    Call SetTimer '设置自动刷新
    
    
    '医技科室初始化
    '-----------------------------------------------------
    If Not InitDepts Then Unload Me: Exit Sub
    If cboDept.ListIndex = -1 Then
        If InStr(mstrPrivs, "所有科室") > 0 Then
            MsgBox "没有发现医技科室信息,请先到部门管理中设置。", vbInformation, gstrSysName
        Else
            MsgBox "没有发现你所属科室,不能使用医技工作站。", vbInformation, gstrSysName
        End If
        Unload Me: Exit Sub
    End If
    
    '界面恢复:放在最后执行
    '-----------------------------------------------------
    If Val(zlDatabase.GetPara("使用个性化风格")) = 1 Then
                strTab = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(DkpMain), DkpMain.Name, "")
        If InStr(strTab, "消息提醒") <> 0 Then DkpMain.LoadStateFromString strTab
    End If
    Call RestoreWinState(Me, App.ProductName, , True)
    Me.WindowState = vbMaximized
    Call SetFixedCommandBar(cbsMain(2).Controls)
    rptPati.Columns.Find(col_选择).Visible = True '= mblnFilter
    rptPati.Columns.Find(col_床号).Visible = mblnShowBed
End Sub

Private Sub SetTimer()
    mintRefresh = Val(zlDatabase.GetPara("医技刷新间隔", glngSys, p医技工作站))
    If mintRefresh <> 0 And mintRefresh < 30 Then mintRefresh = 30
End Sub

Private Sub InitExecTable()
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "要求时间,1520,1;执行时间,1520,1;本次数次,815,1;执行摘要,1550,1;执行人,700,1;登记时间,1520,1;登记人,700,1;执行结果,815,1;核对人,750,1;核对时间,1530,1;说明,500,1;来源,600,1"
    arrHead = Split(strHead, ";")
    With vsExec
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
            If COLExec.Count <> UBound(arrHead) + 1 Then COLExec.Add i, Split(arrHead(i), ",")(0)
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        
    End With
End Sub

Private Function FuncExecAuditBatch() As Boolean
'功能：批量核对
    Dim bln输血皮试 As Boolean
    Dim strSQL As String
    Dim str核对人 As String
    Dim i As Long
    Dim arrSQL As Variant
    Dim rsTmp As Recordset
    Dim blnTrans As Boolean
    Dim strMsgNameSame As String
    Dim strMsgNoRecord As String
    Dim strMsgHave As String
    Dim strMsgKZ As String
    Dim strMsg As String
    
    bln输血皮试 = False
    
    On err GoTo errH
    arrSQL = Array()
    For i = 0 To rptPati.Rows.Count - 1
        With rptPati.Rows(i)
            If Not .GroupRow Then
                If rptPati.Rows(i).Record(col_选择).Checked Then
                    If (Mid(gstr医嘱核对, 2, 1) = "1" And .Record(COL_诊疗类别).Value = "E" And .Record(COL_操作类型).Value = "1" Or _
                        Mid(gstr医嘱核对, 1, 1) = "1" And .Record(COL_诊疗类别).Value = "E" And .Record(COL_操作类型).Value = "8" Or _
                        Mid(gstr医嘱核对, 1, 1) = "1" And .Record(COL_诊疗类别).Value = "K") Then
                        
                        bln输血皮试 = True
                        If .Record(COL_核对人).Value & "" = "" Then
                            If Val(.Record(col_执行状态).Value & "") = 3 Then
                                If str核对人 = "" Then str核对人 = zlDatabase.UserIdentifyByUser(Me, "在核对执行情况前，请您先输入用户名和密码进行身份验证。", glngSys, p医技工作站, "执行情况登记", , True)
                                If str核对人 = "" Then Exit Function
                                
                                '读取执行人
                                strSQL = "Select 执行人 From 病人医嘱执行 Where 医嘱ID=[1] and 发送号=[2]"
                                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.Record(col_医嘱ID).Value & ""), Val(.Record(col_发送号).Value & ""))
                                If rsTmp.RecordCount > 0 Then
                                    If str核对人 = rsTmp!执行人 & "" Then
                                        strMsgNameSame = strMsgNameSame & "," & .Record(col_单据号).Value
                                    Else
                                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                        arrSQL(UBound(arrSQL)) = "Zl_病人医嘱核对_Insert(" & Val(.Record(col_医嘱ID).Value) & "," & Val(.Record(col_发送号).Value) & ",'" & str核对人 & "')"
                                    End If
                                End If
                            Else
                                strMsgNoRecord = strMsgNoRecord & "," & .Record(col_单据号).Value
                            End If
                        Else
                            strMsgHave = strMsgHave & "," & .Record(col_单据号).Value
                        End If
                    Else
                        strMsgKZ = strMsgKZ & "," & .Record(col_单据号).Value
                    End If
                    rptPati.Rows(i).Record(col_选择).Checked = False
                End If
            End If
        End With
    Next
    
    If bln输血皮试 = False Then
        If Val(gstr医嘱核对) = 1 Then
            strSQL = "你勾选的项目中没有皮试项目，无需核对。"
        ElseIf Val(gstr医嘱核对) = 10 Then
            strSQL = "你勾选的项目中没有输血项目，无需核对。"
        Else
            strSQL = "你勾选的项目中没有输血或皮试项目，无需核对。"
        End If
        MsgBox strSQL, vbInformation, "医嘱核对"
        '显示执行情况
        Call LoadPatients
        Exit Function
    End If

    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    If strMsgNameSame <> "" Then
        strMsg = strMsg & "以下单据的审核人和执行人为同一个人：" & vbCrLf & Mid(strMsgNameSame, 2) & "。" & vbCrLf
    End If
    If strMsgNoRecord <> "" Then
        strMsg = strMsg & "以下单据还未进行执行情况登记：" & vbCrLf & Mid(strMsgNoRecord, 2) & "。" & vbCrLf
    End If
    If strMsgHave <> "" Then
        strMsg = strMsg & "以下单据已经进行了核对：" & vbCrLf & Mid(strMsgHave, 2) & "。" & vbCrLf
    End If
    If strMsgKZ <> "" Then
        strMsg = strMsg & "以下单据不是输血或是皮试类项目：" & vbCrLf & Mid(strMsgKZ, 2) & "。" & vbCrLf
    End If
    
    If UBound(arrSQL) < 0 Then
        MsgBox "您勾选的项目未核对成功，其中：" & vbCrLf & strMsg, vbInformation, "医嘱核对"
    Else
        If strMsg <> "" Then
            MsgBox "共核对了" & UBound(arrSQL) + 1 & "个项目，其中：" & vbCrLf & strMsg, vbInformation, "医嘱核对"
        End If
    End If

    '显示执行情况
    Call LoadPatients
    FuncExecAuditBatch = True

    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function FuncThingAudit() As Boolean
'功能：核对
    Dim bln输血皮试 As Boolean
    Dim strSQL As String
    Dim str核对人 As String
    Dim i As Long
    
    '判断是否批量执行模式，是则在单独的流程中调用
    If rptPati.Columns(col_选择).Visible Then
        For i = 0 To rptPati.Rows.Count - 1
            If Not rptPati.Rows(i).GroupRow Then
                If rptPati.Rows(i).Record(col_选择).Checked Then Exit For
            End If
        Next
        If i <= rptPati.Rows.Count - 1 Then
            If MsgBox("要对当前选择的一个或多个项目进行核对吗？", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                Call FuncExecAuditBatch
            End If
            Exit Function
        End If
    End If
    
    With rptPati.SelectedRows(0)
        bln输血皮试 = (Mid(gstr医嘱核对, 2, 1) = "1" And .Record(COL_诊疗类别).Value = "E" And .Record(COL_操作类型).Value = "1" Or _
            Mid(gstr医嘱核对, 1, 1) = "1" And .Record(COL_诊疗类别).Value = "E" And .Record(COL_操作类型).Value = "8" Or _
            Mid(gstr医嘱核对, 1, 1) = "1" And .Record(COL_诊疗类别).Value = "K")
            
        If Not bln输血皮试 Then
            If Val(gstr医嘱核对) = 1 Then
                MsgBox "只能核对皮试医嘱。", vbInformation, gstrSysName
            ElseIf Val(gstr医嘱核对) = 10 Then
                MsgBox "只能核对输血医嘱。", vbInformation, gstrSysName
            Else
                MsgBox "只能核对输血或是皮试医嘱。", vbInformation, gstrSysName
            End If
            Exit Function
        End If
        
        If vsExec.TextMatrix(vsExec.FixedRows, COLExec("核对人")) <> "" Then
            MsgBox "该医嘱还已经核对，不能再次核对。", vbInformation, gstrSysName
            Exit Function
        End If
        If vsExec.TextMatrix(vsExec.FixedRows, vsExec.FixedCols) = "" Then
            MsgBox "该医嘱还未进行执行情况登记，不能核对。", vbInformation, gstrSysName
            Exit Function
        End If
        str核对人 = zlDatabase.UserIdentifyByUser(Me, "在核对执行情况前，请您先输入用户名和密码进行身份验证。", glngSys, p医技工作站, "执行情况登记", , True)
        If str核对人 = "" Then Exit Function
        
        If str核对人 = vsExec.TextMatrix(vsExec.FixedRows, COLExec("执行人")) Then
            MsgBox "执行人不能和审核人相同，不能核对。", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    With vsExec
        On Error GoTo errH
        strSQL = "Zl_病人医嘱核对_Insert(" & Val(rptPati.SelectedRows(0).Record(col_医嘱ID).Value) & "," & Val(rptPati.SelectedRows(0).Record(col_发送号).Value) & ",'" & str核对人 & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, "医嘱核对")
        '显示执行情况
        Call LoadPatients
        FuncThingAudit = True
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function FuncExecDelAuditBatch() As Boolean
'功能：批量取消核对
    Dim bln输血皮试 As Boolean
    Dim strSQL As String
    Dim str核对人 As String
    Dim i As Long
    Dim arrSQL As Variant
    Dim strMsg As String
    Dim rsTmp As Recordset
    Dim blnTrans As Boolean
    Dim blnIsTwo As Boolean   '判断是否存在有两个以上的核对人
    Dim strTmp As String
    Dim bln核对人 As Boolean
    Dim strMsgNoRecord As String
    Dim strMsgHave As String
    Dim strMsgKZ As String
    Dim strMsgNoExec As String
    Dim datCur As Date
    
    bln输血皮试 = False
    
    On err GoTo errH
    arrSQL = Array()
    datCur = zlDatabase.Currentdate
    For i = 0 To rptPati.Rows.Count - 1
        With rptPati.Rows(i)
            If Not .GroupRow Then
                If rptPati.Rows(i).Record(col_选择).Checked And Not blnIsTwo And Not bln核对人 Then
                    If (Mid(gstr医嘱核对, 2, 1) = "1" And .Record(COL_诊疗类别).Value = "E" And .Record(COL_操作类型).Value = "1" Or _
                        Mid(gstr医嘱核对, 1, 1) = "1" And .Record(COL_诊疗类别).Value = "E" And .Record(COL_操作类型).Value = "8" Or _
                        Mid(gstr医嘱核对, 1, 1) = "1" And .Record(COL_诊疗类别).Value = "K") Then
                        
                        bln输血皮试 = True
                        If .Record(COL_核对人).Value & "" <> "" Then
                            If Val(.Record(col_执行状态).Value & "") = 3 Then
                                If strTmp <> "" And strTmp <> .Record(COL_核对人).Value & "" Then
                                    blnIsTwo = True
                                Else
                                    strTmp = .Record(COL_核对人).Value & ""
                                End If
                                If CanUnExec(CDate(.Record(COL_接收时间).Value & ""), datCur) Then
                                    If .Record(COL_核对人).Value & "" <> UserInfo.姓名 Then
                                        If str核对人 = "" Then str核对人 = zlDatabase.UserIdentifyByUser(Me, "在取消核对前，请您先输入用户名和密码进行身份验证。", glngSys, p医技工作站, "执行情况登记", , True)
                                        If str核对人 = "" Then Exit Function
                                        
                                        If str核对人 = .Record(COL_核对人).Value & "" Then
                                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                            arrSQL(UBound(arrSQL)) = "Zl_病人医嘱核对_Delete(" & Val(.Record(col_医嘱ID).Value) & "," & Val(.Record(col_发送号).Value) & ")"
                                        Else
                                            bln核对人 = True
                                            str核对人 = .Record(COL_核对人).Value & ""
                                        End If
                                    Else
                                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                        arrSQL(UBound(arrSQL)) = "Zl_病人医嘱核对_Delete(" & Val(.Record(col_医嘱ID).Value) & "," & Val(.Record(col_发送号).Value) & ")"
                                    End If
                                Else
                                    strMsgNoExec = strMsgNoExec & "," & .Record(col_单据号).Value
                                End If
                            Else
                                strMsgNoRecord = strMsgNoRecord & "," & .Record(col_单据号).Value
                            End If
                        Else
                            strMsgHave = strMsgHave & "," & .Record(col_单据号).Value
                        End If
                    Else
                        strMsgKZ = strMsgKZ & "," & .Record(col_单据号).Value
                    End If
                    rptPati.Rows(i).Record(col_选择).Checked = False
                End If
            End If
        End With
    Next
    
    If bln输血皮试 = False Then
        If Val(gstr医嘱核对) = 1 Then
            strSQL = "你勾选的项目中没有皮试项目，无需取消核对。"
        ElseIf Val(gstr医嘱核对) = 10 Then
            strSQL = "你勾选的项目中没有输血项目，无需取消核对。"
        Else
            strSQL = "你勾选的项目中没有输血或皮试项目，无需取消核对。"
        End If
        MsgBox strSQL, vbInformation, "取消核对"
        '显示执行情况
        Call LoadPatients
        Exit Function
    End If
    
    If blnIsTwo Then
        MsgBox "不能同时取消多个人核对的项目，请选择同一个人所核对的项目。", vbInformation, "取消核对"
        '显示执行情况
        Call LoadPatients
        Exit Function
    End If
    
    If bln核对人 Then
        MsgBox "只能取消自己核对的医嘱，当前选择的医嘱核对人是""" & str核对人 & """。", vbInformation, "取消核对"
        '显示执行情况
        Call LoadPatients
        Exit Function
    End If
    

    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "取消核对")
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    If strMsgNoRecord <> "" Then
        strMsg = strMsg & "以下单据还未进行执行情况登记：" & vbCrLf & Mid(strMsgNoRecord, 2) & "。" & vbCrLf
    End If
    If strMsgHave <> "" Then
        strMsg = strMsg & "以下单据还未进行核对：" & vbCrLf & Mid(strMsgHave, 2) & "。" & vbCrLf
    End If
    If strMsgKZ <> "" Then
        strMsg = strMsg & "以下单据不是输血或是皮试类项目：" & vbCrLf & Mid(strMsgKZ, 2) & "。" & vbCrLf
    End If
    If strMsgNoExec <> "" Then
        strMsg = strMsg & "以下单据的核对时间超过了医嘱执行有效天数：" & vbCrLf & Mid(strMsgNoExec, 2) & "。" & vbCrLf
    End If
    
    If UBound(arrSQL) < 0 Then
        MsgBox "您勾选的项目未取消成功，其中：" & vbCrLf & strMsg, vbInformation, "取消核对"
    Else
        If strMsg <> "" Then
            MsgBox "共取消核对了" & UBound(arrSQL) + 1 & "个项目,其中：" & vbCrLf & strMsg, vbInformation, "取消核对"
        End If
    End If
    '显示执行情况
    Call LoadPatients
    FuncExecDelAuditBatch = True

    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function FuncThingDelAudit() As Boolean
'功能：取消核对
    Dim bln输血皮试 As Boolean
    Dim strSQL As String
    Dim str核对人 As String
    Dim i As Long
    
    '判断是否批量执行模式，是则在单独的流程中调用
    If rptPati.Columns(col_选择).Visible Then
        For i = 0 To rptPati.Rows.Count - 1
            If Not rptPati.Rows(i).GroupRow Then
                If rptPati.Rows(i).Record(col_选择).Checked Then Exit For
            End If
        Next
        If i <= rptPati.Rows.Count - 1 Then
            If MsgBox("要对当前选择的一个或多个项目进行取消核对吗？", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                Call FuncExecDelAuditBatch
            End If
            Exit Function
        End If
    End If
    
    With rptPati.SelectedRows(0)
        bln输血皮试 = (Mid(gstr医嘱核对, 2, 1) = "1" And .Record(COL_诊疗类别).Value = "E" And .Record(COL_操作类型).Value = "1" Or _
            Mid(gstr医嘱核对, 1, 1) = "1" And .Record(COL_诊疗类别).Value = "E" And .Record(COL_操作类型).Value = "8" Or _
            Mid(gstr医嘱核对, 1, 1) = "1" And .Record(COL_诊疗类别).Value = "K")
            
        If Not bln输血皮试 Then
            If Val(gstr医嘱核对) = 1 Then
                MsgBox "只能取消核对皮试医嘱。", vbInformation, gstrSysName
            ElseIf Val(gstr医嘱核对) = 10 Then
                MsgBox "只能取消核对输血医嘱。", vbInformation, gstrSysName
            Else
                MsgBox "只能取消核对输血或是皮试医嘱。", vbInformation, gstrSysName
            End If
            Exit Function
        End If
        
        If vsExec.TextMatrix(vsExec.FixedRows, COLExec("核对人")) = "" Then
            MsgBox "该医嘱还未进行核对，不能取消。", vbInformation, gstrSysName
            Exit Function
        End If

    End With
    With vsExec
        If vsExec.TextMatrix(vsExec.FixedRows, COLExec("核对人")) <> UserInfo.姓名 Then
            str核对人 = zlDatabase.UserIdentifyByUser(Me, "在取消核对前，请您先输入用户名和密码进行身份验证。", glngSys, p医技工作站, "执行情况登记", , True)
            If str核对人 = "" Then Exit Function
            If str核对人 <> vsExec.TextMatrix(vsExec.FixedRows, COLExec("核对人")) Then
                MsgBox "只能取消自己核对的医嘱，当前医嘱核对人是""" & vsExec.TextMatrix(vsExec.FixedRows, COLExec("核对人")) & """", vbInformation, gstrSysName
                Exit Function
            End If
        Else
            If MsgBox("你确定要取消核对吗？", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then Exit Function
        End If
        On Error GoTo errH
        
        strSQL = "Zl_病人医嘱核对_Delete(" & Val(rptPati.SelectedRows(0).Record(col_医嘱ID).Value) & "," & Val(rptPati.SelectedRows(0).Record(col_发送号).Value) & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, "取消医嘱核对")
        Call LoadPatients
        FuncThingDelAudit = True
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim objRow As ReportRow, i As Long, strCardNO As String
    Dim lng医嘱ID As Long
    
    If Control.ID <> 0 Then
        If cbsMain.FindControl(, Control.ID, True, True) Is Nothing Then Exit Sub
    End If
    
    '用血医嘱的特出处理
    If picBlood.Visible = True And InStr(1, "," & mstrBloodControlIDs & ",", "," & Control.ID & ",") <> 0 Then
        If Not mobjFrmBloodExe Is Nothing Then
            Call mobjFrmBloodExe.zlExecuteCommandBars(Control)
        End If
        Exit Sub
    End If
    
    Select Case Control.ID
    Case conMenu_View_ToolBar_Button '工具栏
        For i = 2 To cbsMain.Count
            Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Text '按钮文字
        For i = 2 To cbsMain.Count
            For Each objControl In Me.cbsMain(i).Controls
                If objControl.ID = conMenu_Help_Help Or objControl.ID = conMenu_File_Exit Or objControl.ID = conMenu_File_Print Or objControl.ID = conMenu_File_Preview Then
                    objControl.Style = xtpButtonIcon
                Else
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                End If
            Next
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Size '大图标
        Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
        Me.cbsMain.RecalcLayout
    Case conMenu_View_StatusBar '状态栏
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsMain.RecalcLayout
    Case conMenu_View_FontSize_S '小字体
        If mbytSize <> 0 Then
            mbytSize = 0
            Call SetFontSize(True)
            Me.cbsMain.RecalcLayout
        End If
    Case conMenu_View_FontSize_L '大字体
        If mbytSize <> 1 Then
            mbytSize = 1
            Call SetFontSize(True)
            Me.cbsMain.RecalcLayout
        End If
    Case conMenu_View_Jump '跳转
        If Me.tbcSub.Selected.Index + 1 <= Me.tbcSub.ItemCount - 1 Then
            Me.tbcSub.Item(Me.tbcSub.Selected.Index + 1).Selected = True
        Else
            Me.tbcSub.Item(0).Selected = True
        End If
    Case conMenu_Tool_Archive '电子病案查阅
        If rptPati.SelectedRows.Count > 0 Then
            If rptPati.SelectedRows(0).Record(col_来源).Value = "住院" Then
                Call frmArchiveView.ShowArchive(Me, mlng病人ID, mlng主页ID)
            Else
                Call frmArchiveView.ShowArchive(Me, mlng病人ID, Get挂号ID(mstr挂号单))
            End If
        End If
    Case conMenu_Tool_Reference_1 '疾病诊断参考
        Call gobjKernel.ShowDiagHelp(vbModeless, Me)
    Case conMenu_Tool_Reference_2 '诊疗措施参考
        Call gobjKernel.ShowClincHelp(vbModeless, Me)
    Case conMenu_Manage_FeeItemSet  '诊疗项目费用设置
        Call Set诊疗项目费用设置
         
    Case conMenu_View_Find '查找
        If Me.ActiveControl Is PatiIdentify Then
            PatiIdentify.SetFocus '有时需要定位一下
            If PatiIdentify.Text <> "" Then
                Call ExecuteFindPati
            End If
        Else
            PatiIdentify.SetFocus
        End If
    Case conMenu_View_FindNext '查找下一个
        If PatiIdentify.Text = "" And mvarCond.身份证 = "" Then
            PatiIdentify.SetFocus
        Else
            Call ExecuteFindPati(True, IIf(PatiIdentify.Text = "", mvarCond.身份证, ""))
        End If
    Case conMenu_View_Expend_CurCollapse '折叠当前组
        If rptPati.SelectedRows.Count > 0 Then
            If rptPati.SelectedRows(0).GroupRow Then
                rptPati.SelectedRows(0).Expanded = False
            ElseIf Not rptPati.SelectedRows(0).ParentRow Is Nothing Then
                If rptPati.SelectedRows(0).ParentRow.GroupRow Then
                    rptPati.SelectedRows(0).ParentRow.Expanded = False
                End If
            End If
        End If
        '因折叠定位到分组上,不会自动激活该事件
        Call rptPati_SelectionChanged
    Case conMenu_View_Expend_CurExpend '展开当前组
        If rptPati.SelectedRows.Count > 0 Then
            rptPati.SelectedRows(0).Expanded = True
        End If
    Case conMenu_View_Expend_AllCollapse '折叠所有组
        For Each objRow In rptPati.Rows
            If objRow.GroupRow Then objRow.Expanded = False
        Next
        '因折叠定位到分组上,不会自动激活该事件
        Call rptPati_SelectionChanged
    Case conMenu_View_Expend_AllExpend '展开所有组
        For Each objRow In rptPati.Rows
            If objRow.GroupRow Then objRow.Expanded = True
        Next
    Case conMenu_View_PatInfor '病人显示:只显示已收费病人
        mbln只显已收费 = Not mbln只显已收费
        cbsMain.RecalcLayout: Call LoadPatients
    Case conMenu_View_ShowAll '病人显示:显示他科执行的病人
        mbln他科执行 = Not mbln他科执行
        cbsMain.RecalcLayout: Call LoadPatients
    Case conMenu_View_Show '病人过滤
        Call PatientFilter
    Case conMenu_View_Refresh '刷新
        Call LoadPatients
        Call LoadNotify
    
    Case conMenu_File_RoomSet '执行间设置
        With frmTechnicRoom
            .lblDept.Tag = Me.cboDept.ItemData(cboDept.ListIndex)
            .lblDept.Caption = Me.cboDept.Text & "执行间"
            .Show 1, Me
        End With
    Case conMenu_File_Parameter '参数设置
        Call ParameterSetup
    Case conMenu_Manage_Bespeak '时间安排
        Call FuncExecPlanTime
    Case conMenu_Manage_Plan '执行报到
        Call FuncExecPlan
    Case conMenu_Manage_Logout '取消报到
        Call FuncExecErase
    Case conMenu_Manage_Refuse '拒绝执行
        Call FuncExecRefuse
    Case conMenu_Manage_ReGet '取消拒绝
        Call FuncExecRestore
    Case conMenu_Manage_ThingAdd '记录执行情况
        Call FuncThingNew
    Case conMenu_Manage_ThingModi '调整执行情况
        Call FuncThingModi
    Case conMenu_Manage_ThingDel '删除执行情况
        Call FuncThingDel
    Case conMenu_Manage_ThingAudit '核对
        Call FuncThingAudit
    Case conMenu_Manage_ThingDelAudit '取消核对
        Call FuncThingDelAudit
    Case conMenu_Manage_Complete '执行完成
        Call FuncExecFinish
    Case conMenu_Manage_Undone '取消完成
        Call FuncExecCancel
    Case conMenu_Manage_RequestPrint * 10# + 1 To conMenu_Manage_RequestPrint * 10# + 9 '打印诊疗单据
        Call FuncBillPrint(Control)
    Case conMenu_Manage_RequestBatPrint '批量打印条码
        Call FuncBatchPrint
    Case conMenu_Manage_ReportEdit '报告填写
        Call FuncShowReport(0)
    Case conMenu_Manage_ReportView '报告查阅
        Call FuncShowReport(1)
    Case conMenu_Manage_ReportPrint '报告打印
        Call FuncShowReport(2)
    Case conMenu_Manage_ReportPreview '报告预览
        Call FuncShowReport(3)
    Case conMenu_Manage_AppendBill '补费按钮
        Call FuncAppendBill
    Case conMenu_Help_Web_Home 'Web上的中联
        Call zlHomePage(Me.hwnd)
    Case conMenu_Help_Web_Forum '中联论坛
        Call zlWebForum(Me.hwnd)
    Case conMenu_Help_Web_Mail '发送反馈
        Call zlMailTo(Me.hwnd)
    Case conMenu_Help_Help '帮助
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_About '关于
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_File_Exit '退出
        Unload Me
    Case Else
        timRefresh.Enabled = False
        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            '执行发布到当前模块的报表
            If Split(Control.Parameter, ",")(1) = "ZL1_INSIDE_1263_1" Then '医技工作报表
                Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                    "执行科室=" & zlCommFun.GetNeedName(cboDept.Text) & "|" & cboDept.ItemData(cboDept.ListIndex))
            Else
                If rptPati.SelectedRows.Count = 0 Then
                    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, "执行科室=" & cboDept.ItemData(cboDept.ListIndex))
                Else
                    With rptPati.SelectedRows(0)
                        If .Record(col_来源).Value = "住院" Then
                            Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                                "执行科室=" & cboDept.ItemData(cboDept.ListIndex), "医嘱ID=" & .Record(col_医嘱ID).Value, "发送号=" & .Record(col_发送号).Value, _
                                    "NO=" & .Record(col_单据号).Value, "病人ID=" & mlng病人ID, "主页ID=" & mlng主页ID, "住院号=" & .Record(col_标识号).Value)
                        Else
                            Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                                "执行科室=" & cboDept.ItemData(cboDept.ListIndex), "医嘱ID=" & .Record(col_医嘱ID).Value, "发送号=" & .Record(col_发送号).Value, _
                                    "NO=" & .Record(col_单据号).Value, "病人ID=" & mlng病人ID, "挂号单=" & mstr挂号单, "门诊号=" & .Record(col_标识号).Value)
                        End If
                    End With
                End If
            End If
        Else
            Select Case Me.tbcSub.Selected.Tag
            Case "医嘱附费"
                Call mclsExpenses.zlExecuteCommandBars(Control)
            Case "门诊医嘱"
                Call mclsOutAdvices.zlExecuteCommandBars(Control)
            Case "住院医嘱"
                Call mclsInAdvices.zlExecuteCommandBars(Control)
            Case "住院病历"
                Call mclsInEPRs.zlExecuteCommandBars(Control)
            Case "门诊病历"
                Call mclsOutEPRs.zlExecuteCommandBars(Control)
            Case "新病历"
                Call mclsEMR.zlExecuteCommandBars(Control)
            Case "护理"
                Call mclsTends.zlExecuteCommandBars(Control)
            Case "新版护理"
                Call mclsTendsNew.zlExecuteCommandBars(Control)
            Case "护理病历"
                Call mclsTendEPRs.zlExecuteCommandBars(Control)
            Case Else
                If Not gobjPlugIn Is Nothing Then
                    On Error Resume Next
                    If rptPati.SelectedRows.Count <> 0 Then lng医嘱ID = rptPati.SelectedRows(0).Record(col_医嘱ID).Value
                    Call gobjPlugIn.ExeButtomClick(glngSys, p医技工作站, mcolSubForm("_" & tbcSub.Selected.Tag), tbcSub.Selected.Tag, Control.Caption, mlng病人ID, mlng主页ID, mstr挂号单, lng医嘱ID)
                    Call zlPlugInErrH(err, "ExeButtomClick")
                    err.Clear: On Error GoTo 0
                End If
            End Select
        End If
        timRefresh.Enabled = True
    End Select
End Sub

Private Sub cbsMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Dim i As Long, strKinds As String
    Dim arrKind() As String
    Dim objControl As CommandBarControl
    
    If CommandBar.Parent Is Nothing Then Exit Sub
    
    
    Select Case CommandBar.Parent.ID
    Case conMenu_Manage_Report '报告
        With CommandBar.Controls
            If .Count = 0 Then
                .Add xtpControlButton, conMenu_Manage_ReportEdit, "填写报告(&E)"
                .Add xtpControlButton, conMenu_Manage_ReportView, "查阅报告(&W)"
                .Add(xtpControlButton, conMenu_Manage_ReportPrint, "打印报告(&P)").BeginGroup = True
                .Add xtpControlButton, conMenu_Manage_ReportPreview, "预览报告(&V)"
            End If
        End With
    Case Else
        Select Case tbcSub.Selected.Tag
        Case "医嘱附费"
            Call mclsExpenses.zlPopupCommandBars(CommandBar)
        Case "门诊医嘱"
            Call mclsOutAdvices.zlPopupCommandBars(CommandBar)
        Case "住院医嘱"
            Call mclsInAdvices.zlPopupCommandBars(CommandBar)
        Case "住院病历"
            
        Case "门诊病历"
            
        End Select
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean, blnSelect As Boolean
    Dim int执行状态 As Integer, int执行过程 As Integer
    Dim int综合状态 As Integer, int执行安排 As Integer
    Dim objControl As CommandBarControl
    Dim arrType() As String, i As Long
    
    '初始化一卡通部件,由于activate事件在加载时不激活
    If Not mblnIsInit Then
        mblnIsInit = True
        If Not mobjSquareCard Is Nothing Then
            If mobjSquareCard.zlInitComponents(Me, glngModul, glngSys, gstrDBUser, gcnOracle, False) = False Then
                Set mobjSquareCard = Nothing
                MsgBox "医疗卡部件（zl9CardSquare）初始化失败!", vbInformation, gstrSysName
            Else
                mstrCardKind = mobjSquareCard.zlGetIDKindStr(mstrCardKind)
            End If
            Call PatiIdentify.zlInit(Me, glngSys, p医技工作站, gcnOracle, gstrDBUser, mobjSquareCard, mstrCardKind, "zl9CISJob")
            PatiIdentify.objIDKind.AllowAutoICCard = True
            PatiIdentify.objIDKind.AllowAutoIDCard = True
        
            arrType = Split(mstrCardKind, ";")
            For i = 1 To UBound(arrType) + 1
                If i = mintFindType Then
                    PatiIdentify.objIDKind.IDKind = i
                    mstrFindType = PatiIdentify.objIDKind.Cards(i).名称
                    Exit For
                End If
            Next
            chkFilter.Value = IIf(mblnFilter, 1, 0)
            chkFilter.Enabled = mblnFilterEnabled
        End If
    End If
    '根据权限设置按钮可见状态
    Call SetControlVisible(Control)
    If Not Control.Visible Then Exit Sub
    
    '用血医嘱的特出处理
    If picBlood.Visible = True And InStr(1, "," & mstrBloodControlIDs & ",", "," & Control.ID & ",") <> 0 Then
        If Not mobjFrmBloodExe Is Nothing Then
            Call mobjFrmBloodExe.zlUpdateCommandBars(Control)
        End If
        Exit Sub
    End If
    
    '是否选中了病人
    If rptPati.SelectedRows.Count > 0 Then
        If Not rptPati.SelectedRows(0).GroupRow Then
            blnSelect = True
            '0-未执行,1-已执行,2-拒绝执行,3-正在执行
            int执行状态 = rptPati.SelectedRows(0).Record(col_执行状态).Value
            int综合状态 = rptPati.SelectedRows(0).Record(col_综合状态).Value '子项和独项同执行状态
            '0-无意义;1-已报到;2-检查中;3-检查完成;4-填写报告;5-审核驳回;6-报告完成
            int执行过程 = rptPati.SelectedRows(0).Record(col_执行过程).Value
            '0-不需安排,1-需要安排
            int执行安排 = rptPati.SelectedRows(0).Record(col_执行安排).Value
        End If
    End If
        
    Select Case Control.ID
    Case conMenu_View_ToolBar_Button '工具栏
        If cbsMain.Count >= 2 Then
            Control.Checked = Me.cbsMain(2).Visible
        End If
    Case conMenu_View_ToolBar_Text '图标文字
        If cbsMain.Count >= 2 Then
            Control.Checked = Not (Me.cbsMain(2).Controls(GetFirstCommandBar(cbsMain(2).Controls)).Style = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size '大图标
        Control.Checked = Me.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar '状态栏
        Control.Checked = Me.stbThis.Visible
    Case conMenu_View_FontSize_S '小字体
        Control.Checked = Not (mbytSize = 1)
    Case conMenu_View_FontSize_L '大字体
        Control.Checked = (mbytSize = 1)
    Case conMenu_View_Expend_CurExpend '展开当前组
        blnEnabled = False
        If rptPati.SelectedRows.Count > 0 Then
            If rptPati.SelectedRows(0).GroupRow Then
                blnEnabled = Not rptPati.SelectedRows(0).Expanded
            End If
        End If
        Control.Enabled = blnEnabled
    Case conMenu_View_Expend_CurCollapse '折叠当前组
        blnEnabled = False
        If rptPati.SelectedRows.Count > 0 Then
            If rptPati.SelectedRows(0).GroupRow Then
                blnEnabled = rptPati.SelectedRows(0).Expanded
            ElseIf Not rptPati.SelectedRows(0).ParentRow Is Nothing Then
                If rptPati.SelectedRows(0).ParentRow.GroupRow Then
                    blnEnabled = rptPati.SelectedRows(0).ParentRow.Expanded
                End If
            End If
        End If
        Control.Enabled = blnEnabled
    Case conMenu_View_Expend '折叠/展开组
        Control.Enabled = rptPati.GroupsOrder.Count > 0 And rptPati.Rows.Count > 0
       
    Case conMenu_View_FindNext '查找下一个
        Control.Enabled = Not mblnFilter
    Case conMenu_View_PatInfor '病人显示:只显示已收费病人
        Control.Checked = mbln只显已收费
        Control.Enabled = mbln只显已收费Enabled
    Case conMenu_View_ShowAll '病人显示:显示他科执行的病人
        Control.Checked = mbln他科执行
    Case conMenu_Tool_Archive '电子病案查阅
        Control.Enabled = mlng病人ID <> 0
    Case conMenu_Manage_Bespeak '时间安排:不管是否已报到
        blnEnabled = blnSelect And int综合状态 = 0 And int执行状态 = 0 And int执行安排 = 1 '已有分散状态了不允许
        If blnEnabled Then blnEnabled = rptPati.SelectedRows(0).ParentRow.GroupRow '子项不允许单独进行
        Control.Enabled = blnEnabled
    Case conMenu_Manage_Plan '执行报到
        blnEnabled = blnSelect And int综合状态 = 0 And int执行状态 = 0 '已有分散状态了不允许
        If blnEnabled Then blnEnabled = rptPati.SelectedRows(0).ParentRow.GroupRow '子项不允许单独进行
        Control.Enabled = blnEnabled
    Case conMenu_Manage_Logout '取消报到
        blnEnabled = blnSelect And int综合状态 = 0 And int执行状态 = 0 And int执行过程 = 1 '已有分散状态了不允许
        If blnEnabled Then blnEnabled = rptPati.SelectedRows(0).ParentRow.GroupRow '子项不允许单独进行
        Control.Enabled = blnEnabled
    Case conMenu_Manage_Refuse '拒绝执行
        blnEnabled = blnSelect And int综合状态 = 0 And int执行状态 = 0 And int执行过程 = 0 '已有分散状态了不允许
        If blnEnabled Then blnEnabled = rptPati.SelectedRows(0).ParentRow.GroupRow '子项不允许单独进行
        Control.Enabled = blnEnabled
    Case conMenu_Manage_ReGet '取消拒绝
        blnEnabled = blnSelect And int执行状态 = 2
        If blnEnabled Then blnEnabled = rptPati.SelectedRows(0).ParentRow.GroupRow '子项不允许单独进行
        Control.Enabled = blnEnabled
    Case conMenu_Manage_ThingAdd '记录执行情况
        Control.Enabled = blnSelect And (int综合状态 = 0 Or int综合状态 = 3)
    Case conMenu_Manage_ThingModi, conMenu_Manage_ThingDel '调整执行情况,删除执行情况
        Control.Enabled = blnSelect And (int综合状态 = 0 Or int综合状态 = 3) _
            And vsExec.TextMatrix(vsExec.Row, vsExec.FixedCols) <> "" And vsExec.Row = vsExec.FixedRows
    Case conMenu_Manage_ThingAudit, conMenu_Manage_ThingDelAudit
        Control.Enabled = blnSelect And (int综合状态 = 0 Or int综合状态 = 3)
    Case conMenu_Manage_Complete '执行完成
        If Me.Visible Then
            If rptPati.Columns(col_选择).Visible Then
                Control.Enabled = CanBatchFinish(Control.ID)
            Else
                Control.Enabled = blnSelect And (int综合状态 = 0 Or int综合状态 = 3)
            End If
        End If
    Case conMenu_Manage_Undone '取消完成
        Control.Enabled = blnSelect And int综合状态 = 1
    Case conMenu_Manage_Request, conMenu_Manage_Report '申请、报告菜单
        If blnSelect Then blnEnabled = rptPati.SelectedRows(0).ParentRow.GroupRow '子项不允许单独进行
        Control.Enabled = blnEnabled
    Case conMenu_Manage_RequestPrint '打印诊疗单据
        Control.Enabled = Control.CommandBar.Controls.Count > 0
    Case conMenu_Manage_RequestBatPrint '批量打印条码
        Control.Enabled = blnSelect
    Case conMenu_Manage_ReportEdit '报告填写：对应了病历单据，并且有报告的情况才需要填写
        If blnSelect Then blnEnabled = (int执行状态 = 0 Or int执行状态 = 3) _
            And rptPati.SelectedRows(0).Record(col_文件ID).Value <> 0 And rptPati.SelectedRows(0).Record(col_报告项).Value <> 0
        If blnEnabled Then blnEnabled = rptPati.SelectedRows(0).ParentRow.GroupRow '子项不允许单独进行
        Control.Enabled = blnEnabled
    Case conMenu_Manage_ReportView, conMenu_Manage_ReportPrint, conMenu_Manage_ReportPreview '报告查阅/打印/预览
        If blnSelect Then blnEnabled = rptPati.SelectedRows(0).Record(col_报告ID).Value <> 0
        If blnEnabled Then blnEnabled = rptPati.SelectedRows(0).ParentRow.GroupRow '子项不允许单独进行
        Control.Enabled = blnEnabled
    Case conMenu_Manage_AppendBill
        Control.Enabled = blnSelect And (int综合状态 = 0 Or int综合状态 = 3)
    Case Else
        Select Case tbcSub.Selected.Tag
        Case "医嘱附费"
            Call mclsExpenses.zlUpdateCommandBars(Control)
        Case "门诊医嘱"
            Call mclsOutAdvices.zlUpdateCommandBars(Control)
        Case "住院医嘱"
            Call mclsInAdvices.zlUpdateCommandBars(Control)
        Case "住院病历"
            Call mclsInEPRs.zlUpdateCommandBars(Control)
        Case "门诊病历"
            Call mclsOutEPRs.zlUpdateCommandBars(Control)
        Case "新病历"
            Call mclsEMR.zlUpdateCommandBars(Control)
        Case "护理"
            Call mclsTends.zlUpdateCommandBars(Control)
        Case "新版护理"
            Call mclsTendsNew.zlUpdateCommandBars(Control)
        Case "护理病历"
            Call mclsTendEPRs.zlUpdateCommandBars(Control)
        End Select
    End Select
End Sub

Private Function CanBatchFinish(ByVal lngCmdID As Long) As Boolean
'功能：判断指定的命令在当前选择状态下能否批量执行
    Dim lngCount As Long, i As Long
    Dim blnEnabled As Boolean
    Dim str病人ID As String
    
    With rptPati
        '有选择的情况下，以选择的行为准
        For i = 0 To .Rows.Count - 1
            If Not .Rows(i).GroupRow Then
                If .Rows(i).Record(col_选择).Checked Then
                    lngCount = lngCount + 1
                    If InStr(str病人ID & ",", "," & .Rows(i).Record(col_病人Id).Value & ",") = 0 Then
                        str病人ID = str病人ID & "," & .Rows(i).Record(col_病人Id).Value
                    End If
                    If lngCmdID = conMenu_Manage_Complete Then
                        If Not (.Rows(i).Record(col_综合状态).Value = 0 Or .Rows(i).Record(col_综合状态).Value = 3) Then
                            Exit Function
                        End If
                        'If UBound(Split(Mid(str病人ID, 2), ",")) > 0 Then Exit Function '只可以对一个病人批量完成
                    End If
                End If
            End If
        Next
    
        '一个都没有选择的情况下，以当前项为准
        If lngCount = 0 Then
            blnEnabled = False
            If .SelectedRows.Count > 0 Then
                If Not .SelectedRows(0).GroupRow Then
                    If lngCmdID = conMenu_Manage_Complete Then
                        If .SelectedRows(0).Record(col_综合状态).Value = 0 Or .SelectedRows(0).Record(col_综合状态).Value = 3 Then
                            blnEnabled = True
                        End If
                    End If
                End If
            End If
            If Not blnEnabled Then Exit Function
        End If
    End With
    
    CanBatchFinish = True
End Function

Private Sub SetControlVisible(ByVal Control As CommandBarControl)
'功能：根据权限设置菜单和工具栏的可见状态
    Dim str完成人 As String
    '1.不能通过"已判断"标志跳过,因为子窗体还要判断其他的
    '2.只对要判断的命令进行处理，不然把不该处理的也处理了(如子窗体中的)
    Control.Visible = True
    Select Case Control.ID
    
    Case conMenu_Manage_FeeItemSet
        If InStr(mstrPrivs, "门诊病人") = 0 Then    '没有"诊疗项目费用设置"权限时可查看
            Control.Visible = False
        End If
            
    Case conMenu_Tool_Archive '电子病案查阅
        If GetInsidePrivs(p电子病案查阅) = "" Then Control.Visible = False
    Case conMenu_Tool_Reference_1 '疾病诊断参考
        If GetInsidePrivs(p疾病诊断参考) = "" Then Control.Visible = False
    Case conMenu_Tool_Reference_2 '药品及诊疗参考
        If GetInsidePrivs(p药品诊疗参考) = "" Then Control.Visible = False
    Case conMenu_File_Parameter '参数设置
        'If InStr(mstrPrivs, "参数设置") = 0 Then Control.Visible = False
    Case conMenu_View_ShowAll '显示他科执行的病人
        If InStr(mstrPrivs, "执行他科项目") = 0 Then Control.Visible = False
    Case conMenu_File_RoomSet '执行间设置
        If InStr(mstrPrivs, "执行间设置") = 0 Then Control.Visible = False
    Case conMenu_Manage_Bespeak, conMenu_Manage_Plan, conMenu_Manage_Logout '时间安排,执行安排,取消报到
        If InStr(mstrPrivs, "执行安排") = 0 Then Control.Visible = False
    Case conMenu_Manage_Refuse, conMenu_Manage_ReGet '拒绝执行,取消拒绝
        If InStr(mstrPrivs, "拒绝执行") = 0 Then Control.Visible = False
    Case conMenu_Manage_ThingAdd, conMenu_Manage_ThingModi, conMenu_Manage_ThingDel '记录,调整,删除执行情况
        If InStr(mstrPrivs, "执行情况登记") = 0 Then Control.Visible = False
     Case conMenu_Manage_ThingAudit, conMenu_Manage_ThingDelAudit
        If InStr(GetInsidePrivs(p医技工作站), "执行情况登记") = 0 Or Val(gstr医嘱核对) = 0 Then Control.Visible = False
    Case conMenu_Manage_ThingAudit * 100# + 1, conMenu_Manage_ThingDelAudit * 100# + 1
        If InStr(GetInsidePrivs(p医技工作站), "执行情况登记") = 0 Or picBlood.Visible = False Then Control.Visible = False
    Case conMenu_Manage_Complete '执行完成
        If InStr(mstrPrivs, "确认执行完成") = 0 Then Control.Visible = False
    Case conMenu_Manage_Undone '取消完成
        If rptPati.SelectedRows.Count > 0 Then
            If Not rptPati.SelectedRows(0).GroupRow Then
                str完成人 = Trim(rptPati.SelectedRows(0).Record(col_完成人).Value & "")
            End If
        End If
        If str完成人 = "" Then
            Control.Visible = False
        Else
            If InStr(mstrPrivs, "取消执行完成") > 0 And str完成人 = UserInfo.姓名 Or InStr(mstrPrivs, "取消他人执行完成") > 0 And str完成人 <> UserInfo.姓名 Then
                Control.Visible = True
            Else
                Control.Visible = False
            End If
        End If
    Case conMenu_Manage_ReportEdit '报告填定
        If InStr(mstrPrivs, "报告填写") = 0 Then Control.Visible = False
    Case conMenu_Manage_RequestBatPrint '批量打印条码
        If InStr(mstrPrivs, "打印检验条码") = 0 Then Control.Visible = False
    End Select
End Sub

Private Sub SubWinDefCommandBar(ByVal objItem As TabControlItem)
'功能：刷新子窗体菜单及工具条
    Dim objControl As CommandBarControl
    Dim bytStyle As XTPButtonStyle
    Dim blnShowBar As Boolean
    Dim lngCount As Long, idx As Long
    Dim strName As String
    
    '记录现有菜单样式
    blnShowBar = True
    bytStyle = xtpButtonIconAndCaption
    If cbsMain.Count >= 2 Then
        idx = GetFirstCommandBar(cbsMain(2).Controls)
        If idx > 0 Then
            blnShowBar = cbsMain(2).Visible
            bytStyle = cbsMain(2).Controls(idx).Style
        End If
    End If
    
    '刷新子窗口菜单
    Call LockWindowUpdate(Me.hwnd)
        
    Me.Caption = "医技工作站 - " & objItem.Caption & "(当前用户：" & UserInfo.姓名 & ")"
    
    If InStr(mstrNotify, "1") > 0 Then
        DkpMain.Panes(2).Closed = False '这句先
        DkpMain.Panes(2).Hidden = Val(DkpMain.Panes(2).Tag) = 1
        DkpMain.Panes(2).Title = "消息提醒"
    Else
        DkpMain.Panes(2).Tag = IIf(DkpMain.Panes(2).Hidden, 1, 0)
        DkpMain.Panes(2).Close
    End If
    
    '删除现在的工具栏及顶级菜单项
    For lngCount = cbsMain.ActiveMenuBar.Controls.Count To 1 Step -1
        cbsMain.ActiveMenuBar.Controls(lngCount).Delete
    Next
    For lngCount = cbsMain.Count To 2 Step -1
        cbsMain(lngCount).Delete
    Next
    
    '主窗口重新加入
    Call MainDefCommandBar
    
    '子窗口重新加入
    Select Case objItem.Tag
    Case "医嘱附费"
        Call mclsExpenses.zlDefCommandBars(Me, Me.cbsMain, mobjSquareCard)
    Case "门诊医嘱"
        Call mclsOutAdvices.zlDefCommandBars(Me, Me.cbsMain, 2, Nothing, mobjSquareCard)
    Case "住院医嘱"
        Call mclsInAdvices.zlDefCommandBars(Me, Me.cbsMain, 2, False, mobjSquareCard)
    Case "住院病历"
        Call mclsInEPRs.zlDefCommandBars(Me.cbsMain)
    Case "门诊病历"
        Call mclsOutEPRs.zlDefCommandBars(Me.cbsMain)
    Case "新病历"
        Call mclsEMR.zlDefCommandBars(Me.cbsMain)
    Case "护理"
        Call mclsTends.zlDefCommandBars(Me.cbsMain)
    Case "新版护理"
        Call mclsTendsNew.zlDefCommandBars(Me.cbsMain, True)
    Case "护理病历"
        Call mclsTendEPRs.zlDefCommandBars(Me.cbsMain, True)
    Case Else
        If Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            strName = gobjPlugIn.GetButtomName(glngSys, p医技工作站, mcolSubForm("_" & objItem.Tag), objItem.Tag)
            Call zlPlugInErrH(err, "GetButtomName")
            '构建菜单
            If strName <> "" Then Call PlugInInSideBar(cbsMain, strName)
            err.Clear: On Error GoTo 0
        End If
    End Select
    
    '恢复及固定的一些菜单设置
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    For lngCount = 2 To cbsMain.Count
        cbsMain(lngCount).ContextMenuPresent = False
        cbsMain(lngCount).ShowTextBelowIcons = False
        cbsMain(lngCount).EnableDocking xtpFlagHideWrap
        For Each objControl In cbsMain(lngCount).Controls
            If objControl.ID = conMenu_Help_Help Or objControl.ID = conMenu_File_Exit Or objControl.ID = conMenu_File_Print Or objControl.ID = conMenu_File_Preview Then
                objControl.Style = xtpButtonIcon
            Else
                objControl.Style = bytStyle
            End If
        Next
        cbsMain(lngCount).Visible = blnShowBar
    Next
    
    '如果用了RecalcLayout反而不正常
    Call LockWindowUpdate(0)
    
    Set mfrmActive = mcolSubForm("_" & objItem.Tag)
End Sub

Private Sub SubWinRefreshData(ByVal objItem As TabControlItem)
'功能：刷新子窗体数据及状态
    Dim int类型 As Integer
    Dim blnEdit As Boolean
    
    If mlng病人ID = 0 Or (Me.Visible And Not objItem.Visible) Then '后面条件为异常不明原因
        '要求子窗体按无数据处理界面
        Select Case objItem.Tag
        Case "医嘱附费"
            Call mclsExpenses.zlRefresh(0, "")
        Case "门诊医嘱"
            Call mclsOutAdvices.zlRefresh(0, "", False)
        Case "住院医嘱"
            Call mclsInAdvices.zlRefresh(0, 0, 0, 0, CDate(0), 0)
        Case "住院病历"
            Call mclsInEPRs.zlRefresh(0, 0, 0, False, False)
        Case "门诊病历"
            Call mclsOutEPRs.zlRefresh(0, 0, 0, False, False)
        Case "新病历"
            Call mclsEMR.zlRefresh(0, 0, 0, 0, 2)
        Case "护理"
            Call mclsTends.zlRefresh(0, 0, 0, False, False)
        Case "新版护理"
            Call mclsTendsNew.zlRefresh(0, 0, 0, False, False)
        Case "护理病历"
            Call mclsTendEPRs.zlRefresh(0, 0, 0, False, False, False)
        Case Else
            If Not gobjPlugIn Is Nothing Then
                On Error Resume Next
                Call gobjPlugIn.RefreshForm(glngSys, p医技工作站, mcolSubForm("_" & objItem.Tag), objItem.Tag, 0, "", 0, False)
                Call zlPlugInErrH(err, "RefreshForm")
                err.Clear: On Error GoTo 0
            End If
        End Select
    Else
        Select Case objItem.Tag
        Case "医嘱附费"
            With rptPati.SelectedRows(0)
                Call mclsExpenses.zlRefresh(cboDept.ItemData(cboDept.ListIndex), .Record(col_医嘱ID).Value & ":" & _
                    .Record(col_发送号).Value & ":" & IIf(Not .ParentRow.GroupRow, 1, 0), .Record(COL_数据转出).Value = 1)
            End With
        Case "门诊医嘱"
            With rptPati.SelectedRows(0)
                Call mclsOutAdvices.zlRefresh(mlng病人ID, mstr挂号单, _
                    InStr(",0,3,", .Record(col_执行状态).Value) > 0 And .Record(col_来源).Value <> "体检", _
                    .Record(COL_数据转出).Value = 1, _
                    .Record(col_医嘱ID).Value, cboDept.ItemData(cboDept.ListIndex), mclsMipModule)
            End With
        Case "住院医嘱"
            With rptPati.SelectedRows(0)
                If .Record(col_出院日期).Value = "" Then
                    If .Record(COL_状态).Value = 3 Then
                        int类型 = 1 '预出院
                    ElseIf .Record(COL_状态).Value = 2 Then
                        int类型 = 6 '转科或转病区待入住病人
                    Else
                        int类型 = 0 '在院
                    End If
                Else
                    int类型 = 2 '出院
                End If
                Call mclsInAdvices.zlRefresh(mlng病人ID, mlng主页ID, .Record(col_病区id).Value, _
                    .Record(col_科室).Value, int类型, .Record(COL_数据转出).Value = 1, _
                    .Record(col_医嘱ID).Value, .Record(col_执行状态).Value, cboDept.ItemData(cboDept.ListIndex), , , mclsMipModule)
            End With
         Case "住院病历"
            blnEdit = True
            With rptPati.SelectedRows(0)
                If mlngType = pt出院 Or mlngType = pt死亡 Then
                    '1-等待审查;2-拒绝审查;3-正在审查;4-审查反馈;5-审查归档
                    If Not (.Record(col_审查).Value = 0 Or .Record(col_审查).Value = 2) Then
                        '可能是在院抽查反馈状态，出院后并未提交审查
                        If .Record(col_审查).Value = 1 Then
                            blnEdit = False
                        Else
                            If PatiMedRecHaveSubmit(mlng病人ID, mlng主页ID) Then blnEdit = False
                        End If
                    End If
                End If
    
                Call mclsInEPRs.zlRefresh(mlng病人ID, mlng主页ID, _
                    mlngDept, blnEdit, Val(.Record(COL_数据转出).Value), 0, False, .Record(col_病区id).Value + 0, mlngState)
            End With
        Case "门诊病历"
            With rptPati.SelectedRows(0)
                Call mclsOutEPRs.zlRefresh(mlng病人ID, Val(.Record(col_挂号ID).Value), mlngDept, mlng病人ID <> 0, Val(.Record(COL_数据转出).Value))
            End With
        Case "新病历"
            With rptPati.SelectedRows(0)
                If Val(.Record(col_挂号ID).Value) = 0 Then
                    If .Record(col_出院日期).Value = "" Then
                        If .Record(COL_状态).Value = 3 Then
                            int类型 = 1 '预出院
                        ElseIf .Record(COL_状态).Value = 2 Then
                            int类型 = 6 '转科或转病区待入住病人
                        Else
                            int类型 = 0 '在院
                        End If
                    Else
                        int类型 = 2 '出院
                    End If
                    Call mclsEMR.zlRefresh(mlng病人ID, mlng主页ID, mlngDept, int类型, 2)
                Else
                    Call mclsEMR.zlRefresh(mlng病人ID, Val(.Record(col_挂号ID).Value), mlngDept, 1, 1)
                End If
            End With
        Case "护理"
            blnEdit = True
            With rptPati.SelectedRows(0)
                If mlngType = pt出院 Or mlngType = pt死亡 Then
                    If Not (.Record(col_审查).Value = 0 Or .Record(col_审查).Value = 2 Or .Record(col_审查).Value = 999) Then
                        '可能是在院抽查反馈状态，出院后并未提交审查
                        If .Record(col_图标).Value = 1 Then blnEdit = False
                    End If
                End If
                Call mclsTends.zlRefresh(mlng病人ID, mlng主页ID, .Record(col_病区id).Value + 0, blnEdit, False, .Record(col_病区id).Value + 0, mlngState)
            End With
        Case "护理病历"
                Call mclsTendEPRs.zlRefresh(mlng病人ID, mlng主页ID, rptPati.SelectedRows(0).Record(col_病区id).Value + 0, True, False, Val(rptPati.SelectedRows(0).Record(COL_数据转出).Value))
        Case "新版护理"
            blnEdit = True
            With rptPati.SelectedRows(0)
                If mlngType = pt出院 Or mlngType = pt死亡 Then
                    If Not (.Record(col_审查).Value = 0 Or .Record(col_审查).Value = 2 Or .Record(col_审查).Value = 999) Then
                        '可能是在院抽查反馈状态，出院后并未提交审查
                        If .Record(col_图标).Value = 1 Then blnEdit = False
                    End If
                End If
                Call mclsTendsNew.zlRefresh(mlng病人ID, mlng主页ID, .Record(col_病区id).Value + 0, blnEdit, False, .Record(col_病区id).Value + 0, mlngState)
            End With
        Case Else
            If Not gobjPlugIn Is Nothing Then
                On Error Resume Next
                Call gobjPlugIn.RefreshForm(glngSys, p医技工作站, mcolSubForm("_" & objItem.Tag), objItem.Tag, mlng病人ID, mstr挂号单, mlng主页ID, Val(rptPati.SelectedRows(0).Record(COL_数据转出).Value), 0, cboDept.ItemData(cboDept.ListIndex))
                Call zlPlugInErrH(err, "RefreshForm")
                err.Clear: On Error GoTo 0
            End If
        End Select
    End If
    Call SetFontSize(Not Me.Visible)
End Sub

Private Sub MainDefCommandBar()
'功能：主窗口菜单定义部份
'说明：
'1.其中固有的菜单和按钮必须有，作为子窗体处理菜单的基准
'2.其他命令根据主窗体业务的不同，可能不同
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom

    '菜单定义
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False) '固有
    objMenu.ID = conMenu_FilePopup '对xtpControlPopup类型的命令ID需重新赋值
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…") '固有
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)") '固有
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)") '固有
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…") '固有
        
        Set objControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置(&M)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_RoomSet, "执行间设置(&R)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): objControl.BeginGroup = True '固有
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ManagePopup, "执行(&C)", -1, False)
    objMenu.ID = conMenu_ManagePopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Manage_Request, "申请(&R)")
        With objPopup.CommandBar.Controls
            .Add(xtpControlButtonPopup, conMenu_Manage_RequestPrint, "打印申请单据(&J)").BeginGroup = True
            .Add xtpControlButton, conMenu_Manage_RequestBatPrint, "批量打印条码(&B)"
        End With
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Manage_Report, "报告(&O)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Bespeak, "时间安排(&B)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Plan, "执行报到(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Logout, "取消报到(&L)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Refuse, "拒绝执行(&R)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReGet, "取消拒绝(&G)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Complete, "执行完成(&C)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Undone, "取消完成(&U)")
        '用血医嘱执行前核对
        If gbln血库系统 = True Then
            Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingAudit * 100# + 1, "核查(&V)")
            objControl.ToolTipText = "执行前核查"
            objControl.BeginGroup = True
            objControl.IconId = conMenu_Manage_ThingAudit
            Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingDelAudit * 100# + 1, "取消核查(&Z)")
            objControl.IconId = conMenu_Manage_ThingDelAudit
        End If
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingAdd, "记录执行情况(&A)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingModi, "调整执行情况(&M)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingDel, "删除执行情况(&D)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingAudit, "核对"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingDelAudit, "取消核对")
        If mbln附费按钮 Then
            Set objControl = .Add(xtpControlButton, conMenu_Manage_AppendBill, "补费")
                objControl.IconId = conMenu_Edit_Append
        End If
        
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False) '固有
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)") '固有
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False '固有
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False '固有
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False '固有
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)") '固有
        objControl.BeginGroup = True
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_FontSize, "字体大小(&N)") '固有
        objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_FontSize_S, "小字体(&S)", -1, False '固有
            .Add xtpControlButton, conMenu_View_FontSize_L, "大字体(&L)", -1, False '固有
        End With

        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_Expend, "展开/折叠组(&X)"): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_AllCollapse, "折叠所有组(&L)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_AllExpend, "展开所有组(&X)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_CurCollapse, "折叠当前组(&C)", -1, False): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_CurExpend, "展开当前组(&E)", -1, False)
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_FindNext, "查找下一个(&N)")
        
        Set objControl = .Add(xtpControlButton, conMenu_View_PatInfor, "只显示已收费的病人(&P)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_ShowAll, "显示他科执行的病人(&O)")
        Set objControl = .Add(xtpControlButton, conMenu_View_Show, "病人过滤(&O)")
            objControl.IconId = conMenu_View_Filter
            
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): objControl.BeginGroup = True '固有
        Set objControl = .Add(xtpControlButton, conMenu_View_Jump, "窗格跳转(&J)")
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "工具(&T)", -1, False)
    objMenu.ID = conMenu_ToolPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Archive, "电子病案查阅(&I)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Reference, "资料参考(&R)"): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Tool_Reference_1, "疾病诊断参考(&D)", -1, False
            .Add xtpControlButton, conMenu_Tool_Reference_2, "诊疗措施参考(&C)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Manage_FeeItemSet, "诊疗项目费用设置(&C)")
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False) '固有
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)") '固有
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName) '固有
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False '固有
            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "论坛(&F)", -1, False '固有
            .Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False '固有
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): objControl.BeginGroup = True '固有
    End With


    '工具栏定义:包括公共部份
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印") '固有
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览") '固有
        
        If gbln血库系统 = True Then
            Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingAudit * 100# + 1, "核查")
            objControl.IconId = conMenu_Manage_ThingAudit
            objControl.ToolTipText = "执行前核查"
            objControl.BeginGroup = True
        End If
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingAdd, "记录")
        If gbln血库系统 = False Then objControl.BeginGroup = True
        objControl.ToolTipText = "记录执行情况"
        
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingAudit, "核对")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Complete, "完成")
        Set objPopup = .Add(xtpControlPopup, conMenu_Manage_Report, "报告")
        objPopup.ID = conMenu_Manage_Report: objPopup.IconId = conMenu_Manage_Report
        If mbln附费按钮 Then
            Set objControl = .Add(xtpControlButton, conMenu_Manage_AppendBill, "补费")
                objControl.IconId = conMenu_Edit_Append
        End If
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Show, "过滤")
            objControl.BeginGroup = True: objControl.IconId = conMenu_View_Filter: objControl.ToolTipText = "病人过滤"
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助") '固有
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出") '固有
    End With
    
    '设置一些公共的热键绑定
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyW, conMenu_Manage_Bespeak '时间安排
        .Add FCONTROL, vbKeyL, conMenu_Manage_Plan '执行报到
        .Add FCONTROL, vbKeyV, conMenu_Manage_ThingAdd '记录执行情况
        .Add FCONTROL, vbKeyI, conMenu_Manage_Complete '执行完成
        .Add FCONTROL, vbKeyAdd, conMenu_View_Expend_AllExpend '展开所有组
        .Add FCONTROL, vbKeySubtract, conMenu_View_Expend_AllCollapse '折叠所有组
        .Add FCONTROL, vbKeyF, conMenu_View_Find '查找病人
        .Add 0, vbKeyF3, conMenu_View_FindNext '查找下一个
        .Add 0, vbKeyF12, conMenu_File_Parameter '参数设置
        .Add FCONTROL, vbKeyP, conMenu_File_Print '打印
        .Add FCONTROL, vbKeyT, conMenu_View_Show '过滤
        .Add 0, vbKeyF5, conMenu_View_Refresh '刷新
        .Add 0, vbKeyF6, conMenu_View_Jump '跳转
        .Add 0, vbKeyF1, conMenu_Help_Help '帮助
    End With

    '设置一些公共的不常用命令
    '-----------------------------------------------------
    With cbsMain.Options
'        .AddHiddenCommand conMenu_File_PrintSet '打印设置
'        .AddHiddenCommand conMenu_File_Excel '输出到Excel
'        .AddHiddenCommand conMenu_View_Jump '跳转
    End With
            
    '读取发布到该模块的报表(不含虚拟模块的)
    '-----------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs)
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize
End Sub

Private Sub mclsMipModule_ReceiveMessage(ByVal strMsgItemIdentity As String, ByVal strMsgContent As String)
'功能：消息接收
    Dim blnRecToLis As Boolean '是否加载到提醒列表中
    Dim rsMsg As ADODB.Recordset
    
    If cboDept.ListIndex = -1 Then Exit Sub
    
    If Mid(mstrNotify, 1, 1) = "1" And strMsgItemIdentity = "ZLHIS_CHARGE_001" Then
        blnRecToLis = True
    ElseIf Mid(mstrNotify, 2, 1) = "1" And strMsgItemIdentity = "ZLHIS_CIS_004" Then
        blnRecToLis = True
    End If
    
    If blnRecToLis Then
        Set rsMsg = zlDatabase.ParseXMLToRecord(strMsgItemIdentity, strMsgContent)
        If rsMsg Is Nothing Then Exit Sub
        Call AddMsgToLis(rsMsg)
    End If
End Sub

Private Sub PatiIdentify_FindPatiArfter(ByVal objCard As zlIDKind.Card, ByVal blnCard As Boolean, ShowName As String, objHisPati As zlIDKind.PatiInfor, objCardData As zlIDKind.PatiInfor, strErrMsg As String, blnCancel As Boolean)
    Dim lngPatiID As Long
    
    If objHisPati Is Nothing Then
        lngPatiID = 0
    Else
        lngPatiID = objHisPati.病人ID
    End If
    
    Call ExecuteFindPati(False, , lngPatiID)
End Sub

Private Sub PatiIdentify_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If mblnIsInit = True Then mintFindType = Index: mstrFindType = objCard.名称
End Sub

Private Sub picApplyInfo_Resize()
    rtfAppend.Width = picApplyInfo.ScaleWidth
    rtfAppend.Height = picApplyInfo.ScaleHeight - lblApply.Height
End Sub

Private Sub picApplyUD_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If vsExec.Width + X < 800 Or picApplyInfo.Width - X < 800 Then Exit Sub
        picApplyUD_S.Left = picApplyUD_S.Left + X
        vsExec.Width = vsExec.Width + X
        picApplyInfo.Left = picApplyInfo.Left + X
        picApplyInfo.Width = picApplyInfo.Width - X
        picBlood.Width = picBlood.Width + X
        Me.Refresh
    End If
End Sub

Private Sub fraUD_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If picExec.Height + Y < 1500 Or tbcSub.Height - Y < 2000 Then Exit Sub
        fraUD_S.Top = fraUD_S.Top + Y
        picExec.Height = picExec.Height + Y
        tbcSub.Top = fraUD_S.Top + fraUD_S.Height
        tbcSub.Height = tbcSub.Height - Y
        Me.Refresh
    End If
End Sub

Private Sub mclsEPRReport_AfterSaved(lngRecordId As Long)
    '填写报告之后，刷新相关数据
    Dim rsTemp As New ADODB.Recordset, lng医嘱ID As Long, strSQL As String
       
    On Error GoTo ErrHand
    strSQL = "Select 医嘱id,查阅状态 From 病人医嘱报告 Where 病历Id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngRecordId)
    If rsTemp.RecordCount > 0 Then
        lng医嘱ID = Val("" & rsTemp!医嘱ID)
        If Val("" & rsTemp!查阅状态) = 1 Then
            strSQL = "Zl_报告查阅记录_Cancel(" & lng医嘱ID & "," & lngRecordId & ",Null)"
            Call zlDatabase.ExecuteProcedure(strSQL, "更新查阅状态")
        End If
    End If
    
    mstrPrePati = ""
    Call rptPati_SelectionChanged
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mclsInAdvices_RequestRefresh(ByVal RefreshNotify As Boolean)
'功能：医嘱子窗体要求刷新
    If Not RefreshNotify Then Call LoadPatients '注意要判断
End Sub

Private Sub mclsInAdvices_StatusTextUpdate(ByVal Text As String)
'功能：医嘱子窗体要求更新状态栏
    Me.stbThis.Panels(2).Text = Text
End Sub

Private Sub mclsInAdvices_ViewEPRReport(ByVal 报告ID As Long, ByVal CanPrint As Boolean)
'功能：查看电子病历报告
    Call gobjRichEPR.ViewDocument(Me, 报告ID, CanPrint)
End Sub

Private Sub mclsInAdvices_PrintEPRReport(ByVal 报告ID As Long, ByVal Preview As Boolean)
'功能：按编辑格式打印报告
    Call gobjRichEPR.PrintOrPreviewDoc(Me, cpr诊疗报告, 报告ID, Not Preview)
End Sub

Private Sub mclsInAdvices_ViewPACSImage(ByVal 医嘱ID As Long)
'功能：PACS观片处理
    With rptPati.SelectedRows(0)
        If CreateObjectPacs(gobjPublicPacs) Then
            Call gobjPublicPacs.ShowImage(医嘱ID, Me, .Record(COL_数据转出).Value = 1)
        End If
    End With
End Sub

Private Sub mclsOutAdvices_RequestRefresh()
'功能：医嘱子窗体要求刷新
    Call LoadPatients
End Sub

Private Sub mclsOutAdvices_StatusTextUpdate(ByVal Text As String)
'功能：医嘱子窗体要求更新状态栏
    Me.stbThis.Panels(2).Text = Text
End Sub

Private Sub mclsExpenses_RequestRefresh()
'功能：医嘱子窗体要求刷新
    Call LoadPatients
End Sub

Private Sub mclsExpenses_StatusTextUpdate(ByVal bytType As Byte, ByVal Text As String)
    '功能：医嘱子窗体要求更新状态栏
    Me.stbThis.Panels(2).Text = Text
End Sub

Private Sub mclsOutAdvices_ViewEPRReport(ByVal 报告ID As Long, ByVal CanPrint As Boolean)
'功能：查看电子病历报告
    Call gobjRichEPR.ViewDocument(Me, 报告ID, CanPrint)
End Sub

Private Sub mclsOutAdvices_PrintEPRReport(ByVal 报告ID As Long, ByVal Preview As Boolean)
'功能：按编辑格式打印报告
    Call gobjRichEPR.PrintOrPreviewDoc(Me, cpr诊疗报告, 报告ID, Not Preview)
End Sub

Private Sub mclsOutAdvices_ViewPACSImage(ByVal 医嘱ID As Long)
'功能：PACS观片处理
    With rptPati.SelectedRows(0)
        If CreateObjectPacs(gobjPublicPacs) Then
            Call gobjPublicPacs.ShowImage(医嘱ID, Me, .Record(COL_数据转出).Value = 1)
        End If
    End With
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
'功能：身份证识别成功后激活
    mvarCond.身份证 = strID
    If mstrFindType = "二代身份证" Then
        PatiIdentify.Text = mvarCond.身份证
    Else
        PatiIdentify.Text = "" '否则清除(目前是在已清除情况下才能激活)。
    End If
    Call ExecuteFindPati(False, mvarCond.身份证)
End Sub

Private Sub rptNotify_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Call rptNotify_KeyUp(vbKeyReturn, 0)
End Sub

Private Sub rptNotify_KeyUp(KeyCode As Integer, Shift As Integer)
'功能：阅读消息后删除消息，（双击消息或者选中消息后再按回车键）
    Dim objControl As CommandBarControl
    Dim lngIndex As Long, lng病人ID As Long, lng主页ID As Long
    Dim str业务 As String, strSQLRead As String
    Dim blnTmp As Boolean
    Dim strNO As String
    Dim blnEnabled As Boolean
    Dim objRow As ReportRow
    Dim strTmp As String
    Dim i As Long
 
    If KeyCode = vbKeyReturn Then
        If rptNotify.SelectedRows.Count > 0 Then
            With rptNotify.SelectedRows(0).Record
                strNO = .Item(C_消息).Value
                str业务 = .Item(C_业务).Value
                lng病人ID = Val(.Item(C_病人ID).Value)
                lng主页ID = Val(.Item(C_主页ID).Value)
                lngIndex = .Index
            End With
            strSQLRead = "Zl_业务消息清单_Read(" & lng病人ID & "," & lng主页ID & ",'" & strNO & "',4,'" & UserInfo.姓名 & "'," & cboDept.ItemData(cboDept.ListIndex) & ")"
            
            If strNO = "ZLHIS_CHARGE_001" Then
                If Val(str业务) = 2 Then '病人来源必须是住院
                    '调用销帐审核接口，如果在费用界面则要求刷新界面
                    blnTmp = mobjKernel.ChargeDelAudit(Me, mlngDept, lng病人ID)
                    If tbcSub.Selected.Tag = "医嘱附费" Then
                        Call mclsInAdvices_RequestRefresh(blnTmp)
                    End If
                End If
            End If
            If strNO = "ZLHIS_BLOOD_007" And gbln血库系统 Then     '未回收前不允许设为已读
                If gobjPublicBlood Is Nothing And gbln血库系统 Then InitObjBlood
                If gobjPublicBlood.zlIsBloodMessageDone(1, lng病人ID, lng主页ID, 4, cboDept.ItemData(cboDept.ListIndex)) Then
                    Call rptNotify.Records.RemoveAt(lngIndex)
                    Call rptNotify.Populate
                End If
                Exit Sub
            End If
            Call zlDatabase.ExecuteProcedure(strSQLRead, Me.Caption)
            Call rptNotify.Records.RemoveAt(lngIndex)
            Call rptNotify.Populate
        End If
    End If
End Sub

Private Sub rptNotify_SelectionChanged()
    Dim strCurPati As String
    Dim lngIndex As Long
    Dim str业务  As String
     
    If rptNotify.SelectedRows.Count = 0 Then Exit Sub  '非正常情况
    
    With rptNotify.SelectedRows(0)
        lngIndex = rptNotify.FocusedRow.Record.Index
        If rptNotify.Rows(lngIndex).Record(C_消息).Value = "ZLHIS_CIS_004" Or rptNotify.Rows(lngIndex).Record(C_消息).Value = "ZLHIS_BLOOD_007" Then
            str业务 = Val(rptNotify.Rows(lngIndex).Record(C_业务).Value)
            If rptNotify.Rows(lngIndex).Record(C_消息).Value = "ZLHIS_BLOOD_007" Then str业务 = Split(rptNotify.Rows(lngIndex).Record(C_业务).Value, ":")(1)
            If rptPati.SelectedRows.Count > 0 Then
                If Not rptPati.SelectedRows(0).GroupRow Then
                    strCurPati = rptPati.SelectedRows(0).Record.Tag & "_"
                End If
            End If
            
            If InStr(strCurPati, str业务 & "_") = 0 Then
                If Not LocatePati(str业务) Then
                    Call LoadPatients
                    Call LocatePati(str业务)
                End If
            End If
        End If
    End With
End Sub

Private Sub rptPati_ItemCheck(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim objRecord As ReportRecord
    
    '主项目和子项目不能同时勾选，以区分两种模式
    If Row.Record.Childs.Count > 0 And Item.Checked Then
        For Each objRecord In Row.Record.Childs
            objRecord(col_选择).Checked = False
        Next
    ElseIf Not Row.ParentRow.GroupRow And Item.Checked Then
        Row.ParentRow.Record(col_选择).Checked = False
    End If
End Sub

Private Sub rptPati_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Row.Childs.Count > 0 And Not Row.GroupRow Then Row.Expanded = Not Row.Expanded
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "病人颜色" Then
        Call zlDatabase.ShowPatiColorTip(Me)
    End If
End Sub

Private Sub tbcSub_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'功能：刷新子窗体界面及数据
'说明：仅在人为切换界面卡片激活
    Dim Index As Long, objItem As TabControlItem
    
    If mblnTabTmp Then Exit Sub
    If Item.Tag = "" Then Exit Sub '初始添卡时,还没赋值
    
    If Item.Handle = picTmp.hwnd Then
        Index = Item.Index
        mblnTabTmp = True
        Screen.MousePointer = 11
        On Error GoTo errH
        Select Case Item.Tag
            Case "医嘱附费"
                Set objItem = tbcSub.InsertItem(Index, "医嘱附加费用", mcolSubForm("_医嘱附费").hwnd, 0)
                objItem.Tag = "医嘱附费"
            Case "门诊医嘱"
                Set objItem = tbcSub.InsertItem(Index, "门诊医嘱", mcolSubForm("_门诊医嘱").hwnd, 0)
                objItem.Tag = "门诊医嘱"
            Case "住院医嘱"
                Set objItem = tbcSub.InsertItem(Index, "住院医嘱", mcolSubForm("_住院医嘱").hwnd, 0)
                objItem.Tag = "住院医嘱"
            Case "住院病历"
                Set objItem = tbcSub.InsertItem(Index, "住院病历", mcolSubForm("_住院病历").hwnd, 0)
                objItem.Tag = "住院病历"
            Case "新病历"
                Set objItem = tbcSub.InsertItem(Index, "电子病历", mcolSubForm("_新病历").hwnd, 0)
                objItem.Tag = "新病历"
            Case "门诊病历"
                Set objItem = tbcSub.InsertItem(Index, "门诊病历", mcolSubForm("_门诊病历").hwnd, 0)
                objItem.Tag = "门诊病历"
            Case "护理"
                Set objItem = tbcSub.InsertItem(Index, "护理信息", mcolSubForm("_护理").hwnd, 0)
                objItem.Tag = "护理"
            Case "护理病历"
                Set objItem = tbcSub.InsertItem(Index, "护理病历", mcolSubForm("_护理病历").hwnd, 0)
                objItem.Tag = "护理病历"
            Case "新版护理"
                Set objItem = tbcSub.InsertItem(Index, "护理信息", mcolSubForm("_新版护理").hwnd, 0)
                objItem.Tag = "新版护理"
        End Select
        Call tbcSub.RemoveItem(Index + 1)
        objItem.Selected = True
        Screen.MousePointer = 0
        mblnTabTmp = False
    End If
    '刷新子窗体对应的CommandBar
    Call SubWinDefCommandBar(Item)
    
    '刷新子窗体数据
    If Visible Then Call SubWinRefreshData(Item)
    
    If Visible Then mfrmActive.SetFocus
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboDept_Click()
'功能：刷新界面数据
'说明：从该事件开始,会不重复引发相关的数据读取
    Dim strDeptNode As String
    
    If cboDept.ListIndex = -1 Then Exit Sub
    cboDept.Tag = cboDept.ListIndex
    mblnReturn = True
    If Val(cboDept.ItemData(cboDept.ListIndex)) = mlngDept Then Exit Sub
    
    mlngDept = Val(cboDept.ItemData(cboDept.ListIndex))
    mbln血透室 = Sys.DeptHaveProperty(mlngDept, "血透室")
    mbln产科 = Sys.DeptHaveProperty(mlngDept, "产科")
    
    '如果站点变化，则改变病人病区列表
    strDeptNode = GetDeptNode(mlngDept)
    If strDeptNode <> mstrDeptNode Then
        mstrDeptNode = strDeptNode
        
        '医技科室是服务于特定站点的，清空之前的条件中所选服务于其他站点的病人科室
        If mvarCond.科室ID <> 0 And mstrDeptNode <> "" Then
            strDeptNode = GetDeptNode(mvarCond.科室ID)
            If strDeptNode <> mstrDeptNode Then mvarCond.科室ID = 0
        End If
        
        Call LoadPatiUnit(mstrDeptNode)
        
    ElseIf Me.Visible = False Then
        '启动时调用(mstrDeptNode为空)
        Call LoadPatiUnit(mstrDeptNode)
    End If
    
    '重新读取病人
    Call LoadPatients
End Sub

Private Sub InitReportColumn()
    Dim objCol As ReportColumn

    With rptPati
        '当列顺序或数量(代码或人为隐藏)改变后,要用Find(列号)查找列,但仍可用Record(列号)访问数据行
        Set objCol = .Columns.Add(col_综合状态, "分组", 0, False): objCol.Sortable = False: objCol.Visible = False
        Set objCol = .Columns.Add(col_选择, "", 18, False)
        objCol.Sortable = False
        objCol.AllowDrag = False
        objCol.Alignment = xtpAlignmentLeft
        objCol.Editable = True
        objCol.Icon = 5
        
        Set objCol = .Columns.Add(col_路径, "路径", 30, True): objCol.Sortable = False: objCol.Visible = False
        Set objCol = .Columns.Add(col_执行状态, "状态", 0, False): objCol.Sortable = False: objCol.Visible = False
        Set objCol = .Columns.Add(col_图标, "", 18, False): objCol.Sortable = False: objCol.Alignment = xtpAlignmentCenter: objCol.AllowDrag = False
        Set objCol = .Columns.Add(col_来源, "来源", 30, False)
        Set objCol = .Columns.Add(col_单据号, "单据号", 65, True)
        Set objCol = .Columns.Add(col_紧急, "急", 20, True)
        Set objCol = .Columns.Add(col_姓名, "姓名", 55, True)
        Set objCol = .Columns.Add(col_内容, "内容", 150, True)
        Set objCol = .Columns.Add(col_险类, "险类", 60, True)
        Set objCol = .Columns.Add(col_科室, "科室", 65, True)
        Set objCol = .Columns.Add(col_标识号, "标识号", 62, True)
        Set objCol = .Columns.Add(col_床号, "床号", 35, True)
        Set objCol = .Columns.Add(col_费别, "费别", 55, True)
        Set objCol = .Columns.Add(col_要求时间, "要求时间", 106, True)
        Set objCol = .Columns.Add(col_发送时间, "发送时间", 106, True)
        Set objCol = .Columns.Add(col_执行间, "执行间", 65, True)
        Set objCol = .Columns.Add(col_性别, "性别", 30, True)
        Set objCol = .Columns.Add(col_年龄, "年龄", 30, True)
        Set objCol = .Columns.Add(col_完成人, "完成人", 55, True)
        Set objCol = .Columns.Add(col_完成时间, "完成时间", 106, True)
        
        '隐藏数据列
        Set objCol = .Columns.Add(col_执行科室, "执行科室", 0, False)
        Set objCol = .Columns.Add(col_病人Id, "病人ID", 0, False)
        Set objCol = .Columns.Add(col_主页ID, "主页ID", 0, False)
        Set objCol = .Columns.Add(col_挂号单, "挂号单", 0, False)
        Set objCol = .Columns.Add(col_挂号ID, "挂号ID", 0, False)
        Set objCol = .Columns.Add(col_婴儿, "婴儿", 0, False)
        If ISPassShowCard Then
            Set objCol = .Columns.Add(col_就诊卡号, "就诊卡号", 0, False)
        Else
            Set objCol = .Columns.Add(col_就诊卡号, "就诊卡号", 70, True)
        End If
        Set objCol = .Columns.Add(col_身份证号, "身份证号", 0, False)
        Set objCol = .Columns.Add(col_IC卡号, "IC卡号", 0, False)
        Set objCol = .Columns.Add(col_医保号, "医保号", 0, False)
        Set objCol = .Columns.Add(col_病区id, "病区ID", 0, False)
        Set objCol = .Columns.Add(col_出院日期, "出院日期", 0, False)
        Set objCol = .Columns.Add(COL_状态, "状态", 0, False)
        Set objCol = .Columns.Add(col_医嘱ID, "医嘱ID", 0, False)
        Set objCol = .Columns.Add(col_相关ID, "相关ID", 0, False)
        Set objCol = .Columns.Add(col_发送号, "发送号", 0, False)
        Set objCol = .Columns.Add(COL_诊疗类别, "诊疗类别", 0, False)
        Set objCol = .Columns.Add(col_执行过程, "执行过程", 0, False)
        Set objCol = .Columns.Add(col_执行安排, "执行安排", 0, False)
        Set objCol = .Columns.Add(col_记录性质, "记录性质", 0, False)
        Set objCol = .Columns.Add(COL_数据转出, "数据转出", 0, False)
        Set objCol = .Columns.Add(col_文件ID, "文件ID", 0, False)
        Set objCol = .Columns.Add(col_报告项, "报告项", 0, False)
        Set objCol = .Columns.Add(col_报告ID, "报告ID", 0, False)
        Set objCol = .Columns.Add(col_病人类型, "病人类型", 80, True)
        Set objCol = .Columns.Add(col_门诊记帐, "门诊记帐", 0, False)
        
        Set objCol = .Columns.Add(col_开单人, "开单人", 65, True)
        Set objCol = .Columns.Add(col_审查, "审查", 16, False)
        objCol.TreeColumn = True: objCol.Visible = False
        Set objCol = .Columns.Add(col_类型, "类型", 16, False)
        objCol.TreeColumn = True: objCol.Visible = False
        Set objCol = .Columns.Add(COL_操作类型, "操作类型", 0, False)
        Set objCol = .Columns.Add(COL_核对人, "核对人", 0, False)
        Set objCol = .Columns.Add(col_审核标志, "审核标志", 0, False)
        Set objCol = .Columns.Add(COL_接收时间, "接收时间", 0, False)
        Set objCol = .Columns.Add(col_结算模式, "结算模式", 0, False)
        Set objCol = .Columns.Add(COL_诊疗项目ID, "诊疗项目ID", 0, False)
        Set objCol = .Columns.Add(col_期效, "期效", 0, False)
        Set objCol = .Columns.Add(COL_执行分类, "执行分类", 0, False)
        Set objCol = .Columns.Add(COL_主页挂号ID, "主页挂号ID", 0, False)
        Set objCol = .Columns.Add(COL_附加标志, "附加标志", 0, False)
        For Each objCol In .Columns
            If objCol.Index <> col_选择 Then objCol.Editable = False
            objCol.Groupable = objCol.Index = col_综合状态
            If objCol.Width = 0 Then objCol.Visible = False
        Next
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .TreeIndent = 0 '有分组列时，树形线边上会再有一根边线
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的病人..."
        End With
        .PreviewMode = True
        .AllowColumnRemove = False
        .MultipleSelection = False '会引发SelectionChanged事件
        .ShowItemsInGroups = False
        .SetImageList Me.img16
        
        .GroupsOrder.Add .Columns(col_综合状态)
        .GroupsOrder(0).SortAscending = True '分组之后,如果分组列不显示,分组列的排序是不变的
        
        '分组之后可能失去记录集中的顺序,因此强行加入排序列
        .SortOrder.Add .Columns(col_发送时间)
        .SortOrder(0).SortAscending = False
    End With
    
    
    With rptNotify
        Set objCol = .Columns.Add(c_图标, "", 18, False): objCol.Alignment = xtpAlignmentCenter: objCol.AllowDrag = False
        Set objCol = .Columns.Add(C_病人ID, "病人ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_主页ID, "主页ID", 0, False): objCol.Visible = False
        
        Set objCol = .Columns.Add(c_姓名, "姓名", 60, True)
        Set objCol = .Columns.Add(C_状态, "状态", 150, True)
         
        Set objCol = .Columns.Add(C_消息, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_序号, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_日期, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_业务, "", 0, False): objCol.Visible = False
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = False
            If objCol.Index <> C_序号 Or objCol.Index <> C_日期 Then objCol.Sortable = False
        Next
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .HideSelection = True
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有提醒内容..."
        End With
        .PreviewMode = False
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        .SetImageList Me.imgPati
        
        
        '排序 降序
        .SortOrder.Add .Columns(C_序号)
        .SortOrder(0).SortAscending = False
        .SortOrder.Add .Columns(C_日期)
        .SortOrder(1).SortAscending = False
    End With
    
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = picPati.hwnd
    ElseIf Item.ID = 2 Then
        Item.Handle = picMsg.hwnd
    End If
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    If Not Me.Visible Then Exit Sub
    With Me.picExec
        .Left = lngLeft: .Top = lngTop
        .Width = lngRight - lngLeft
        .Height = fraUD_S.Top - lngTop
    End With
    With Me.fraUD_S
        .Left = lngLeft
         If Not mblnFirstLoad Then .Top = lngTop + picExec.Height
        .Width = lngRight - lngLeft
    End With
    With Me.tbcSub
        .Left = lngLeft: .Width = lngRight - lngLeft
        .Top = fraUD_S.Top + fraUD_S.Height: .Height = lngBottom - .Top
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    Dim blnSetup As Boolean
    
    blnSetup = InStr(";" & mstrPrivs & ";", ";参数设置;") > 0
    If Not tbcSub.Selected Is Nothing Then
        Call zlDatabase.SetPara("医护功能", tbcSub.Selected.Tag, glngSys, p医技工作站, blnSetup)
    End If
    Call zlDatabase.SetPara("执行状态", mstr状态, glngSys, p医技工作站, blnSetup)
    Call zlDatabase.SetPara("只显示已收费的病人", IIf(mbln只显已收费, 1, 0), glngSys, p医技工作站, blnSetup)
    Call zlDatabase.SetPara("病人查找方式", mintFindType, glngSys, p医技工作站, blnSetup)
    Call zlDatabase.SetPara("过滤显示模式", IIf(mblnFilter, 1, 0), glngSys, p医技工作站, blnSetup)
    If Val(zlDatabase.GetPara("使用个性化风格")) = 1 Then
        Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(DkpMain), DkpMain.Name, DkpMain.SaveStateToString)
    End If
    Call zlDatabase.SetPara("字体", mbytSize, glngSys, p医技工作站, blnSetup)
    If Me.Visible Then
        '公共部件固定按第一个控件的样式保存，工作站部件如果第一个是打印，则固定是图标样式,所以需恢复为其它按钮的样式
        cbsMain(2).Controls(1).Style = cbsMain(2).Controls(GetFirstCommandBar(cbsMain(2).Controls)).Style
        Call SaveWinState(Me, App.ProductName)
    End If
    mblnIsInit = False
    If Not mobjIDCard Is Nothing Then
        mobjIDCard.SetEnabled False
        Set mobjIDCard = Nothing
    End If
    Set mobjICCard = Nothing
    Set mobjSquareCard = Nothing
    
    Unload frmTechnicFilter
    
    '强行Unload,不然不会激活子窗体的事件
    For i = 1 To mcolSubForm.Count
        Unload mcolSubForm(i)
    Next
    Set mclsOutAdvices = Nothing
    Set mclsInAdvices = Nothing
    Set mclsExpenses = Nothing
    Set mclsEMR = Nothing
    Set mfrmActive = Nothing
    Set mclsInEPRs = Nothing
    Set mclsOutEPRs = Nothing
    Set mclsEPRReport = Nothing
    Set gobjPublicPacs = Nothing
    
    If Not (mclsMipModule Is Nothing) Then
        mclsMipModule.CloseMessage
        Call DelMipModule(mclsMipModule)
        Set mclsMipModule = Nothing
    End If
    Set mobjKernel = Nothing
    Set mclsTends = Nothing
    Set mclsTendsNew = Nothing
    Set mclsMsg = Nothing
    Set mrsMsg = Nothing
    Set mclsPExp = Nothing
    Set mobjAppendBill = Nothing
    If Not mobjFrmBloodExe Is Nothing Then
        Unload mobjFrmBloodExe
        Set mobjFrmBloodExe = Nothing
    End If
End Sub

Private Sub picExec_Resize()
    On Error Resume Next
    
    fraDiag.Left = 0
    fraDiag.Top = 0
    fraDiag.Width = picExec.ScaleWidth

    fraExec.Left = 0
    fraExec.Width = picExec.ScaleWidth
    fraExec.Top = fraDiag.Top + fraDiag.Height - Screen.TwipsPerPixelY * 6
    
    lblAdvice.Width = fraExec.Width - lblCash.Width - lblRec.Width - lblAdvice.Left - Screen.TwipsPerPixelX * 4
    lblCash.Left = fraExec.Width - lblCash.Width - Screen.TwipsPerPixelX * 4
    
    lblRec.Top = lblCash.Top
    lblRec.Left = lblCash.Left - lblRec.Width - 10
    
    With Me.picApplyUD_S
        .Top = fraExec.Top + fraExec.Height
        .Height = picExec.ScaleHeight - vsExec.Top
    End With
    
    vsExec.Left = 0
    vsExec.Top = fraExec.Top + fraExec.Height
    vsExec.Width = IIf(Me.picApplyInfo.Visible, picApplyUD_S.Left, picExec.Width)
    vsExec.Height = picExec.ScaleHeight - vsExec.Top
    
    picBlood.Left = 0
    picBlood.Top = vsExec.Top
    picBlood.Width = vsExec.Width
    picBlood.Height = vsExec.Height
    If picBlood.Tag = "可见" Then
        vsExec.Visible = False
        picBlood.Visible = True
    Else
        vsExec.Visible = True
        picBlood.Visible = False
    End If
    
    picApplyInfo.Move vsExec.Width + picApplyUD_S.Width, vsExec.Top, _
        picExec.ScaleWidth - vsExec.Width - picApplyUD_S.Width, vsExec.Height
    picApplyUD_S.Visible = picApplyInfo.Visible
    
    fraUD_S.Top = picExec.Top + picExec.Height
    
    tbcSub.Top = fraUD_S.Top + fraUD_S.Height
    
End Sub

Private Sub picPati_GotFocus()
    If rptPati.Visible Then rptPati.SetFocus
End Sub

Private Sub picPati_Resize()
    On Error Resume Next
    
    cboDept.Top = 30
    lblDept.Top = (cboDept.Height - lblDept.Height) / 2 + cboDept.Top
    lblDept.Left = lblDept.Top
    cboDept.Left = lblDept.Left + lblDept.Width + 30
    cboDept.Width = picPati.ScaleWidth - cboDept.Left - lblDept.Left
    
    If cboUnit.Visible Then
        cboUnit.Top = cboDept.Top + cboDept.Height + 45
        lblUnit.Top = (cboUnit.Height - lblUnit.Height) / 2 + cboUnit.Top
        lblUnit.Left = lblDept.Left
        cboUnit.Left = lblUnit.Left + lblUnit.Width + 30
        cboUnit.Width = cboDept.Width
    End If
    
    If Val(gstr医嘱核对) = 0 Then
        fraFilter.Height = chk执行状态(3).Top + chk执行状态(3).Height + IIf(mbytSize = 0, 90, 250)
    Else
        fraFilter.Height = chk执行状态(4).Top + chk执行状态(4).Height + IIf(mbytSize = 0, 90, 250)
    End If
    
    lblFind.Top = IIf(cboUnit.Visible, cboUnit.Top + cboUnit.Height, cboDept.Top + cboDept.Height) + 100
    PatiIdentify.Top = lblFind.Top - 50
    PatiIdentify.Width = cboDept.Left + cboDept.Width - PatiIdentify.Left - chkFilter.Width - 60
    chkFilter.Left = PatiIdentify.Left + PatiIdentify.Width + 30
    chkFilter.Top = PatiIdentify.Top
    
    rptPati.Left = 0
    rptPati.Top = PatiIdentify.Top + PatiIdentify.Height + 30
    rptPati.Width = picPati.ScaleWidth
    rptPati.Height = picPati.ScaleHeight - rptPati.Top - fraFilter.Height
    
    fraFilter.Left = 30
    fraFilter.Top = rptPati.Top + rptPati.Height
    fraFilter.Width = rptPati.Width - 45
End Sub

Private Sub rptPati_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim objColumn As ReportColumn
    
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        '全选病人
        Set objColumn = rptPati.Columns.Find(col_选择)
        If objColumn.Visible Then
            objColumn.Caption = "1"
            Call SelectALLPati(True)
        End If
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        '全清病人
        Set objColumn = rptPati.Columns.Find(col_选择)
        If objColumn.Visible Then
            objColumn.Caption = ""
            Call SelectALLPati(False)
        End If
    ElseIf KeyCode = vbKeyTab Then
        'Panne中的Report控件需要强行处理光标顺序
        '无数据时不能捕获到vbKeyTab
        If Shift = vbShiftMask Then
            If cboDept.Enabled Then cboDept.SetFocus
        Else
            If vsExec.Enabled Then vsExec.SetFocus
        End If
    ElseIf KeyCode = vbKeySpace Then
        '处理复选
        If rptPati.Columns(col_选择).Visible Then
            If rptPati.SelectedRows.Count > 0 Then
                If Not rptPati.SelectedRows(0).GroupRow Then
                    rptPati.SelectedRows(0).Record(col_选择).Checked = Not rptPati.SelectedRows(0).Record(col_选择).Checked
                    Call rptPati_ItemCheck(rptPati.SelectedRows(0), rptPati.SelectedRows(0).Record(col_选择))
                    rptPati.Redraw
                End If
            End If
        End If
    End If
End Sub

Private Sub rptPati_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objHitTest As ReportHitTestInfo
    Dim objPopup As CommandBar
    Dim objControl As CommandBarControl
    Dim objColumn As ReportColumn
        
    If Button = 2 Then
        Set objHitTest = rptPati.HitTest(X, Y)
        If objHitTest.ht = xtpHitTestReportArea And Not objHitTest.Row Is Nothing Then
            If objHitTest.Row.GroupRow Then
                Set objPopup = cbsMain.Add("Popup", xtpBarPopup)
                With objPopup.Controls
                    .Add(xtpControlSplitButtonPopup, conMenu_View_Show, "病人过滤(&O)").IconId = conMenu_View_Filter
                    .Add(xtpControlButton, conMenu_View_Expend_AllCollapse, "折叠所有组(&L)").BeginGroup = True
                    .Add xtpControlButton, conMenu_View_Expend_AllExpend, "展开所有组(&X)"
                    .Add(xtpControlButton, conMenu_View_Expend_CurCollapse, "折叠当前组(&C)").BeginGroup = True
                    .Add xtpControlButton, conMenu_View_Expend_CurExpend, "展开当前组(&E)"
                End With
            Else
                Set objPopup = cbsMain.ActiveMenuBar.Controls(2).CommandBar
            End If
        End If
        
        rptPati.SetFocus
        If Not objPopup Is Nothing Then objPopup.ShowPopup
    ElseIf Button = 1 Then
        If rptPati.HitTest(X, Y).ht = xtpHitTestHeader Then
            Set objColumn = rptPati.HitTest(X, Y).Column
            If Not objColumn Is Nothing Then
                If objColumn.Index = col_选择 Then
                    If objColumn.Caption = "" Then
                        objColumn.Caption = "1"
                        Call SelectALLPati(True)
                    Else
                        objColumn.Caption = ""
                        Call SelectALLPati(False)
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub SelectALLPati(ByVal blnSelect As Boolean)
    Dim objParent As ReportRow, i As Long
    
    If rptPati.Columns(col_选择).Visible And rptPati.SelectedRows.Count > 0 Then
        '先清除所有记录的选择状态
        For i = 0 To rptPati.Records.Count - 1
            rptPati.Records(i)(col_选择).Checked = False
        Next
        
        '当前分组
        Set objParent = rptPati.SelectedRows(0)
        If Not objParent.GroupRow Then
            If objParent.ParentRow.GroupRow Then
                Set objParent = objParent.ParentRow
            Else
                Set objParent = objParent.ParentRow.ParentRow
            End If
        End If
        
        '只针对可见有效行进行处理
        For i = 0 To rptPati.Rows.Count - 1
            If Not rptPati.Rows(i).GroupRow Then
                If Not rptPati.Rows(i).ParentRow.GroupRow Then
                    '检验子项在全选/全清时保持不选状态
                    rptPati.Rows(i).Record(col_选择).Checked = False
                ElseIf rptPati.Rows(i).ParentRow Is objParent Then
                    '同一分组的才处理
                    rptPati.Rows(i).Record(col_选择).Checked = blnSelect
                End If
            End If
        Next
        rptPati.Redraw
    End If
End Sub

Private Sub rptPati_SelectionChanged()
    Dim rsTmp As New ADODB.Recordset, bln收 As Boolean
    Dim strCurPati As String, blnChange As Boolean
    Dim blnUseBlood As Boolean
    If rptPati.SelectedRows.Count = 0 Then Exit Sub  '非正常情况
    
    With rptPati.SelectedRows(0)
        picBlood.Tag = ""
        If Not .GroupRow Then strCurPati = .Record.Tag
        If strCurPati = mstrPrePati Then Exit Sub
        Me.stbThis.Panels(2).Text = ""
        mstrPrePati = strCurPati
        
        If Not .GroupRow Then
            mlng病人ID = .Record(col_病人Id).Value
            mlng主页ID = .Record(col_主页ID).Value
            mstr挂号单 = .Record(col_挂号单).Value
            '病历
            mlngType = rptPati.SelectedRows(0).Record(col_类型).Value
            If (.Record(col_来源).Value & "" = "住院") Then
                '3-住院病历
                If GetInsidePrivs(p住院病历管理, True) <> "" Then
                    Me.tbcSub.Item(mlngInIndex).Visible = True
                End If
                If GetInsidePrivs(p门诊病历管理, True) <> "" Then
                    Me.tbcSub.Item(mlngOutIndex).Visible = False
                End If
                
                If mlngNewIndex <> -1 Then
                    If GetInsidePrivs(p新版住院病历, True) <> "" Then
                        Me.tbcSub.Item(mlngNewIndex).Visible = True
                    Else
                        Me.tbcSub.Item(mlngNewIndex).Visible = False
                    End If
                End If
                
                If GetInsidePrivs(p护理记录管理, True) <> "" Then
                    If mbln血透室 Or mbln产科 Then
                        If mblnNewNurRecord Then
                            Me.tbcSub.Item(mlngNewNurIndex).Visible = True
                            Me.tbcSub.Item(mlngNurEMRIndex).Visible = True
                        Else
                            Me.tbcSub.Item(mlngNurIndex).Visible = True
                        End If
                    Else
                        Me.tbcSub.Item(mlngNurIndex).Visible = False
                        Me.tbcSub.Item(mlngNewNurIndex).Visible = False
                        Me.tbcSub.Item(mlngNurEMRIndex).Visible = False
                        '隐藏后要再定位
                        If tbcSub.Selected.Tag = "护理" Or tbcSub.Selected.Tag = "新版护理" Then
                            tbcSub.Item(0).Selected = True
                        End If
                    End If
                End If
                
            Else
                '4-门诊病历
                If GetInsidePrivs(p门诊病历管理, True) <> "" Then
                    Me.tbcSub.Item(mlngOutIndex).Visible = True
                End If
                If GetInsidePrivs(p住院病历管理, True) <> "" Then
                    Me.tbcSub.Item(mlngInIndex).Visible = False
                End If
                If mlngNewIndex <> -1 Then
                    If GetInsidePrivs(p新版门诊病历, True) <> "" Then
                        Me.tbcSub.Item(mlngNewIndex).Visible = True
                    Else
                        Me.tbcSub.Item(mlngNewIndex).Visible = False
                    End If
                End If
                If GetInsidePrivs(p护理记录管理, True) <> "" Then
                    Me.tbcSub.Item(mlngNurIndex).Visible = False
                     Me.tbcSub.Item(mlngNewNurIndex).Visible = False
                    '隐藏后要再定位
                    If tbcSub.Selected.Tag = "护理" Or tbcSub.Selected.Tag = "新版护理" Then
                        tbcSub.Item(0).Selected = True
                    End If
                End If
            End If
            mlngState = IIf(IsNull(.Record(col_出院日期).Value), IIf(.Record(COL_状态).Value + 0 = 3, ps预出, ps在院), mlngType)
            
            '读取报告ID
            If .Record(col_报告ID).Value = 0 Then
                Call ReadMoreInfo
            End If
            
            '显示医嘱内容
            lblAdvice.Caption = Get执行内容(.Record(col_发送号).Value, .Record(col_医嘱ID).Value, .Record(col_相关ID).Value, .Record(COL_诊疗类别).Value, rptPati.SelectedRows(0))
            
            '显示病人诊断
            If .Record(col_来源).Value <> "住院" Then
                lblDiag(1).Caption = GetPatiDiagnose(.Record(col_病人Id).Value, .Record(col_挂号ID).Value, 1)
            Else
                lblDiag(1).Caption = GetPatiDiagnose(.Record(col_病人Id).Value, .Record(col_主页ID).Value, 2)
            End If
            
            '是否已收费
            If .Record(col_险类).Value <> "" And Val(.Record(col_记录性质).Value) = 1 Then
                '医保病人
                bln收 = Set收费标记(IIf(.Record(col_来源).Value = "住院", 2, 1), Not .ParentRow.GroupRow, _
                    .Record(col_医嘱ID).Value, .Record(col_相关ID).Value, .Record(col_发送号).Value, .Record(COL_诊疗类别).Value, _
                    .Record(col_单据号).Value, .Record(col_记录性质).Value, .Record(col_门诊记帐).Value, .Record(COL_数据转出).Value = 1, .Record(col_发送时间).Value)
            Else
                bln收 = ItemHaveCash(IIf(.Record(col_来源).Value = "住院", 2, 1), Not .ParentRow.GroupRow, _
                    .Record(col_医嘱ID).Value, .Record(col_相关ID).Value, .Record(col_发送号).Value, .Record(COL_诊疗类别).Value, _
                    .Record(col_单据号).Value, .Record(col_记录性质).Value, .Record(col_门诊记帐).Value, 1, .Record(COL_数据转出).Value = 1, .Record(col_发送时间).Value)
            End If
            lblCash.Visible = bln收
            lblRec.Visible = Val(.Record(col_结算模式).Value) = 1
            
             '此处判断是否是用血医嘱
            blnUseBlood = False
            If gbln血库系统 = True And .Record(COL_诊疗类别).Value = "E" And .Record(COL_操作类型).Value = "8" And .Record(COL_执行分类).Value = "1" Then
                blnUseBlood = True
            End If
            
            If blnUseBlood = False Then
                picBlood.Tag = ""
                '显示执行情况
                Call LoadExecList(.Record(col_医嘱ID).Value, .Record(col_发送号).Value)
            Else
                picBlood.Tag = "可见"
                If Not mobjFrmBloodExe Is Nothing Then
                    Call mobjFrmBloodExe.zlRefresh(Me, glngSys, p医技工作站, Val(.Record(col_医嘱ID).Value), mlngDept, GetInsidePrivs(p医技工作站), 1, mlngDept, .Record(COL_数据转出).Value = 1, IIf(mbytSize = 0, 9, 12))
                End If
            End If
            
            '切换显示不同的医嘱子窗体
            blnChange = ExchangeAdvice(.Record(col_来源).Value <> "住院")
        Else
            Call ClearPatiInfo
        End If
        
        '显示可打印的诊疗单据:之所以即时加载,是为了使用F2热键
        Call ShowBillList(cbsMain.FindControl(, conMenu_Manage_RequestPrint, , True))
        
        '刷新子窗体数据
        If Not blnChange Then
            Call SubWinRefreshData(tbcSub.Selected)
        End If
        If Not rptPati.SelectedRows(0).GroupRow Then
            Call ShowBillAppend(1, True)
            picApplyInfo.Visible = Not (rtfAppend.Text = "")
            picApplyUD_S.Visible = picApplyInfo.Visible
            Call picExec_Resize
        Else
            picApplyInfo.Visible = False
            rtfAppend.Text = ""
            picApplyUD_S.Visible = False
            Call picExec_Resize
        End If
    End With
End Sub

Private Sub ReadMoreInfo()
'功能：读取病人相关的更多信息，为效率不在主SQL中读取
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lng医嘱ID As Long
    
    On Error GoTo errH
    
    With rptPati.SelectedRows(0)
        '读取挂号ID
        .Record(col_挂号ID).Value = 0
        If .Record(col_挂号单).Value <> "" Then
            strSQL = "Select ID,附加标志 From 病人挂号记录 Where NO=[1] And 记录性质=1 And 记录状态=1"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CStr(.Record(col_挂号单).Value))
            If Not rsTmp.EOF Then
                .Record(col_挂号ID).Value = rsTmp!ID
                .Record(COL_附加标志).Value = rsTmp!附加标志 & ""
            End If
        ElseIf .Record(COL_主页挂号ID).Value <> "" Then
            strSQL = "Select ID,附加标志 From 病人挂号记录 Where ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.Record(COL_主页挂号ID).Value))
            If Not rsTmp.EOF Then
                .Record(COL_附加标志).Value = rsTmp!附加标志 & ""
            End If
        End If
        
        '读取报告ID
        .Record(col_报告ID).Value = 0
        If (.Record(COL_诊疗类别).Value = "C" Or .Record(COL_诊疗类别).Value = "D") And .Record(col_相关ID).Value <> 0 Then
            lng医嘱ID = .Record(col_相关ID).Value '检验组合取相关ID
        Else
            lng医嘱ID = .Record(col_医嘱ID).Value
        End If
    
        strSQL = "Select 病历ID From 病人医嘱报告 Where 医嘱ID=[1]"
        If .Record(COL_数据转出).Value = 1 Then
            strSQL = Replace(strSQL, "病人医嘱报告", "H病人医嘱报告")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
        If Not rsTmp.EOF Then '可能有多个，目前只有一个
            .Record(col_报告ID).Value = Val(rsTmp!病历id & "")
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub tbcSub_GotFocus()
    On Error Resume Next
    If Not mfrmActive Is Nothing Then mfrmActive.SetFocus
End Sub

Private Function InitDepts() As Boolean
'功能：初始化住院临床科室
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str科室IDs As String, str来源 As String
    
    On Error GoTo errH
    
    '包含门诊/住院医技科室
    str来源 = "3"
    If InStr(mstrPrivs, "门诊病人") > 0 And InStr(mstrPrivs, "住院病人") > 0 Then
        str来源 = "1,2,3"
    ElseIf InStr(mstrPrivs, "门诊病人") > 0 Then
        str来源 = "1,3"
    ElseIf InStr(mstrPrivs, "住院病人") > 0 Then
        str来源 = "2,3"
    End If
    If InStr(mstrPrivs, "所有科室") > 0 Then
        strSQL = _
            " Select Distinct A.ID,A.编码,A.名称" & _
            " From 部门表 A,部门性质说明 B " & _
            " Where B.部门ID = A.ID And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
            " And B.服务对象 IN(" & str来源 & ") And B.工作性质 IN('检查','检验','手术','治疗','营养')" & _
            " Order by A.编码"
    Else
        strSQL = _
            " Select Distinct A.ID,A.编码,A.名称" & _
            " From 部门表 A,部门性质说明 B,部门人员 C " & _
            " Where B.部门ID = A.ID And A.ID=C.部门ID And C.人员ID=[1]" & _
            " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
            " And B.服务对象 IN(" & str来源 & ") And B.工作性质 IN('检查','检验','手术','治疗','营养')" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " Order by A.编码"
    End If
    
    cboDept.Clear
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    str科室IDs = GetUser科室IDs
    For i = 1 To rsTmp.RecordCount
        cboDept.AddItem rsTmp!编码 & "-" & rsTmp!名称
        cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID
        
        If rsTmp!ID = UserInfo.部门ID Then
            Call Cbo.SetIndex(cboDept.hwnd, cboDept.NewIndex) '直接所属优先
        End If
        If InStr("," & str科室IDs & ",", "," & rsTmp!ID & ",") > 0 And cboDept.ListIndex = -1 Then
            Call Cbo.SetIndex(cboDept.hwnd, cboDept.NewIndex)
        End If
        
        rsTmp.MoveNext
    Next
    If cboDept.ListIndex = -1 And cboDept.ListCount > 0 Then
        Call Cbo.SetIndex(cboDept.hwnd, 0)
    End If
        
    If cboDept.ListIndex <> -1 Then
        Call cboDept_Click  '同时对mstrDeptNode赋值
    End If
    InitDepts = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub LoadPatiUnit(ByVal strDeptNode As String)
'功能：读取并加载病人病区
'   strDeptNode=当前医技科室所属的站点
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lngUnit As Long

    On Error GoTo errH
    If cboUnit.ListIndex > 0 Then lngUnit = Val(cboUnit.ItemData(cboUnit.ListIndex))
    
    cboUnit.Clear
    cboUnit.AddItem "所有病区"
    Call Cbo.SetIndex(cboUnit.hwnd, 0)
    
    '来源部门根据当前医技科室的站点来限制站点
    strSQL = "Select A.ID,A.编码,A.名称 From 部门表 A,部门性质说明 B" & _
        " Where A.ID=B.部门ID And B.服务对象 IN(1,2,3) And B.工作性质='护理'" & _
        " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
        IIf(strDeptNode <> "", " And (A.站点 = [1] Or A.站点 is Null)", "") & _
        " Order by A.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strDeptNode)
    Do While Not rsTmp.EOF
        cboUnit.AddItem rsTmp!编码 & "-" & rsTmp!名称
        cboUnit.ItemData(cboUnit.NewIndex) = rsTmp!ID
        If rsTmp!ID = lngUnit Then Call Cbo.SetIndex(cboUnit.hwnd, cboUnit.NewIndex)
        rsTmp.MoveNext
    Loop
        
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Function GetDeptNode(ByVal lngDept As Long) As String
'功能：读取指定部门所属的站点编号
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo errH
    strSQL = "Select 站点 From 部门表 Where ID = [1] And 站点 is not Null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取站点", lngDept)
    If rsTmp.RecordCount > 0 Then GetDeptNode = rsTmp!站点
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function LoadPatients() As Boolean
'功能：读取病人列表
    Dim rsPati As New ADODB.Recordset
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    
    Dim objRow As ReportRow, objPreRecord As ReportRecord
    Dim strPatiRow As String, lngPatiRow As Long, strExpend As String
    
    Dim strSQL As String, strSQL1 As String
    Dim str检验检查 As String, str他科 As String
    Dim blnDateMoved As Boolean, str来源 As String, str病人来源 As String
    Dim datBegin As Date, datEnd As Date
    Dim curDate As Date, lng病区ID As Long
    Dim blnDo As Boolean, blnSub As Boolean, blnPath As Boolean
    Dim lngColor As Long, i As Long, j As Long
    Dim blnNoFilter As Boolean
    Dim strWhere核对 As String, strBloodWhere As String
    Dim str收费判断 As String
    
    '过滤模式时，无病人查找条件则清除界面
    If mblnFilter Then
        If mvarCond.IC卡号 = "" And mvarCond.NO = "" And mvarCond.标识号 = "" _
            And mvarCond.就诊卡 = "" And mvarCond.身份证 = "" And mvarCond.姓名 = "" And mvarCond.医保号 = "" And mvarCond.病人ID = 0 Then
            rptPati.Records.DeleteAll
            rptPati.Populate
            Call ClearPatiInfo
            Call SubWinRefreshData(tbcSub.Selected)
            LoadPatients = True: Exit Function
        End If
    End If
    
    If Not mblnFilter And Val(PatiIdentify.Text) = 0 And mvarCond.开单人 = "" And mstr开单人 <> "" Then
        blnNoFilter = True
        mvarCond.开单人 = mstr开单人
    End If
    
    '当页面下拉框清空，F5刷新，应该恢复上一个的值
    If cboUnit.ListIndex = -1 Then Call Cbo.SetIndex(cboUnit.hwnd, Val(cboUnit.Tag))
    If cboDept.ListIndex = -1 Then Call Cbo.SetIndex(cboDept.hwnd, Val(cboDept.Tag))
            
    mblnShowBed = False
            
    '查询时间段
    curDate = zlDatabase.Currentdate
    If mvarCond.Begin = CDate(0) Then
        datBegin = Int(curDate - 1)
    Else
        datBegin = mvarCond.Begin
    End If
    If mvarCond.End = CDate(0) Then
        datEnd = Format(curDate, "yyyy-MM-dd 23:59")
    Else
        datEnd = mvarCond.End
    End If
    blnDateMoved = zlDatabase.DateMoved(datBegin) '按时间看是否可能已转出
    
    '病人来源权限:(1-门诊,2-住院,3-外来,4-体检),体检病人为门诊病人
    '当门诊留观发送到住院时，医嘱记录中的病人来源仍为2
    If InStr(mstrPrivs, "门诊病人") > 0 And InStr(mstrPrivs, "住院病人") > 0 Then
        str来源 = "1,2,3,4"
        str病人来源 = ""
    ElseIf InStr(mstrPrivs, "门诊病人") > 0 Then
        str来源 = "1,4"
        str病人来源 = " And (Instr([2],','||A.病人来源||',')>0 Or f.病人性质 = 1 And a.病人来源 = 2)"
    ElseIf InStr(mstrPrivs, "住院病人") > 0 Then
        str来源 = "2"
        str病人来源 = " And Instr([2],','||A.病人来源||',')>0 And f.病人性质 <> 1"
    Else
        str来源 = "3"
        str病人来源 = " And Instr([2],','||A.病人来源||',')>0"
    End If
    
    If mvarCond.来源 <> "" Then
        If InStr(mvarCond.来源, "1") > 0 And InStr(mvarCond.来源, "2") = 0 Then
            str病人来源 = str病人来源 & " And (Instr([12],','||a.病人来源||',')>0 Or f.病人性质 = 1 And a.病人来源 = 2)"
        ElseIf InStr(mvarCond.来源, "1") = 0 And InStr(mvarCond.来源, "2") > 0 Then
            str病人来源 = str病人来源 & " And Instr([12],','||a.病人来源||',')>0 And f.病人性质 <> 1"
        Else
            str病人来源 = str病人来源 & " And Instr([12],','||a.病人来源||',')>0"
        End If
    End If
    
    If cboUnit.ListIndex <> -1 Then
        lng病区ID = cboUnit.ItemData(cboUnit.ListIndex)
    End If
    
    '以下规则与发送时的NO分号对应：
    '检验组合中:检验项目只显示一条,采集方法显示一条
    '中药煎法,用法各自独立显示一条
    '附加手术,检查部位执行科室及时间与主项目相同,不显示
    '手术麻醉执行科室为单独，需要显示
    '输血项目，输血途径分别执行
    '特殊医嘱不显示(虽然执行科室一般不会为医技科室)
    If mbln他科执行 Then
        If gbln血库系统 = True Then
            '输血医嘱类医嘱时间默认3天内的都可以显示
            strBloodWhere = " And ((Nvl(a.执行状态, 0) = 0 And Trunc(a.发送时间) = Trunc(Sysdate)) Or (Exists" & vbNewLine & _
                                       "              (Select 1 From 诊疗项目目录 Where b.诊疗项目id = Id And 类别 || 操作类型 || 执行分类 = 'E81') And" & vbNewLine & _
                                       "              (Nvl(a.执行状态, 0) = 0 Or Exists" & vbNewLine & _
                                       "               (Select 1" & vbNewLine & _
                                       "                 From 血液发送记录 a, 血液执行记录 b, 血液配血记录 c" & vbNewLine & _
                                       "                 Where a.收发id = b.收发id(+) And a.配发id = c.Id　and c.申请id = b.相关id　and(Nvl(a.执行状态, 0) = 0 Or b.执行科室id = [1])))))"
        Else
            strBloodWhere = " And Nvl(A.执行状态,0)=0 And Trunc(A.发送时间)=Trunc(Sysdate) "
        End If
        
        str他科 = "" & _
            " And (A.执行部门ID+0=[1] Or A.执行部门ID+0<>[1]  " & strBloodWhere & _
            " And Exists(Select 1 From 诊疗项目目录 C,诊疗执行科室 D Where C.ID=B.诊疗项目ID And C.执行科室=4 And C.ID=D.诊疗项目ID And D.执行科室ID=[1])" & _
            " And Exists(Select 1 From 部门表 C,部门性质说明 D Where C.ID=A.执行部门ID And C.ID=D.部门ID" & _
            " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null) And D.工作性质 IN('检查','检验','手术','治疗','营养'))" & _
            " And Exists(Select 1 From 部门性质说明 C Where C.部门ID=[1] And C.工作性质 IN('检查','检验','手术','治疗','营养')" & _
            " And (C.服务对象=Decode(B.病人来源,4,1,3,1,B.病人来源) Or C.服务对象=3))  )"
    Else
        str他科 = " And A.执行部门ID+0=[1]"
    End If
    If mstr诊疗类别 <> "" Then
        str他科 = str他科 & " And Instr('" & mstr诊疗类别 & "',B.诊疗类别)>0"
    End If
    If mstr治疗类别 <> "" Then
        str他科 = str他科 & " And (Not B.诊疗类别='E' Or B.诊疗类别='E' And Exists(Select 1 From 诊疗项目目录 C Where C.ID=B.诊疗项目ID And Instr('" & mstr治疗类别 & "',C.操作类型)>0))"
    End If
    
'    str检验检查 = "Select 1 From 病人医嘱记录 C,病人医嘱发送 D" & _
'        " Where ((C.诊疗类别='C' And C.相关ID=B.相关ID) or (C.诊疗类别='D' And (C.相关ID=B.相关ID or C.ID=B.相关ID or C.相关ID=B.ID)))" & _
'        " And C.ID=D.医嘱ID And D.发送号=A.发送号  And D.医嘱ID=A.医嘱ID And D.执行状态 IN([3],[4],[5],[6])"
    
    If mbln只显已收费 Then
        str收费判断 = " And (A.记录性质 <> 1 Or A.记录性质 = 1 And a.计费状态 = 3 Or " & _
        " A.记录性质 = 1 And a.计费状态 in (-1,0) And Exists ( Select 1 From 病人医嘱记录 C,病人医嘱发送 D" & _
        " Where ((C.诊疗类别='C' And C.相关ID=B.相关ID) or (C.诊疗类别='D' And (C.相关ID=B.相关ID or C.ID=B.相关ID or C.相关ID=B.ID)))" & _
        " And C.ID=D.医嘱ID And D.发送号=A.发送号 And D.执行状态 IN([3],[4],[5],[6]) And D.计费状态 = 3) )"
    End If

    strSQL = _
        " Select B.姓名,B.年龄,B.性别, A.医嘱ID,A.发送号,B.相关ID,B.序号,B.诊疗类别,b.医嘱期效,B.诊疗项目ID,A.发送时间,A.NO,a.接收人," & _
        "       A.安排时间,Nvl(A.安排时间,Decode(Nvl(B.医嘱期效,0),1,B.开始执行时间,A.首次时间)) as 要求时间," & _
        "       A.记录性质,A.门诊记帐,A.执行状态,A.执行过程,A.执行部门ID,A.完成人,A.完成时间,A.接收时间," & _
        "       B.病人ID,B.主页ID,B.挂号单,B.婴儿,B.病人科室ID,B.病人来源,A.执行间,0 as 数据转出,b.医嘱内容,b.标本部位,b.检查方法,b.开嘱医生, B.紧急标志" & _
        " From 病人医嘱发送 A,  病人医嘱记录 B,材料特性 C" & _
        " Where A.医嘱ID=B.ID And B.收费细目ID=C.材料ID(+) And B.诊疗类别 Not IN('5','6','7')" & _
                str他科 & _
        IIf(Mid(mstr状态, 1, 4) = "1111", "", " And A.执行状态 IN([3],[4],[5],[6]) ") & str收费判断 & _
        IIf(mstrRoom <> "", " And Instr([7],'|'||A.执行间||'|')>0", "") & _
        "       And A." & mstr过滤条件 & " Between [8] And [9]" & _
                IIf(mvarCond.NO <> "", " And A.NO=[10]", "") & _
                IIf(mvarCond.科室ID <> 0, " And B.病人科室ID+0=[11]", "") & _
                IIf(mvarCond.期效 <> 0, " And Nvl(B.医嘱期效,0)=[19]", "") & _
                IIf(mvarCond.开单人 <> "", " And B.开嘱医生=[21]", "") & _
                IIf(mvarCond.病人ID <> 0, " And B.病人ID=[22]", "")
            
    If blnDateMoved Then
        strSQL1 = strSQL
        strSQL1 = Replace(strSQL1, "0 as 数据转出", "1 as 数据转出")
        strSQL1 = Replace(strSQL1, "病人医嘱记录", "H病人医嘱记录")
        strSQL1 = Replace(strSQL1, "病人医嘱发送", "H病人医嘱发送")
        strSQL = strSQL & " Union ALL " & strSQL1
    End If
    
    If Not (Mid(mstr状态, 3, 1) = "0" Or Mid(mstr状态, 5, 1) = "1" Or Val(gstr医嘱核对) = 0 Or Mid(mstr状态, 5, 1) = "") Then
        If Val(gstr医嘱核对) = 11 Then
            strWhere核对 = " And (A.执行状态 <> 3 or A.执行状态 = 3 And (Not (C.操作类型 in ('1','8') And a.诊疗类别='E' Or a.诊疗类别='K') Or (C.操作类型 in ('1','8') And a.诊疗类别='E' Or a.诊疗类别='K') and a.接收人 is null))"
        ElseIf Mid(gstr医嘱核对, 2, 1) = "1" Then
            strWhere核对 = " And (A.执行状态 <> 3 or A.执行状态 = 3 And (Not C.操作类型='1' And a.诊疗类别 = 'E' Or C.操作类型='1' And a.诊疗类别 = 'E' and a.接收人 is null))"
        ElseIf Mid(gstr医嘱核对, 1, 1) = "1" Then
            strWhere核对 = " And (A.执行状态 <> 3 or A.执行状态 = 3 And (Not (C.操作类型='8' And a.诊疗类别='E' Or a.诊疗类别='K') Or (C.操作类型='8' And a.诊疗类别='E' Or a.诊疗类别='K') and a.接收人 is null))"
        End If
    End If
    
    strSQL = _
    " Select /*+ RULE */ DISTINCT" & vbNewLine & _
    "   a.医嘱id, a.发送号, a.相关id, a.序号, a.诊疗类别,c.操作类型,c.执行分类,a.医嘱期效 as 期效, a.诊疗项目id, a.安排时间, a.要求时间, a.发送时间, a.No, a.记录性质, a.执行状态, a.执行过程, a.执行部门id, a.病人id,a.接收人," & vbNewLine & _
    "   a.主页id, a.挂号单, a.婴儿, a.病人科室id, a.紧急标志 ,e.名称 As 科室, g.名称 As 执行科室,NVl(NVl(decode(A.病人来源,4,D.姓名, A.姓名), F.姓名),D.姓名) 姓名 ,NVl(NVl(decode(A.病人来源,4,D.性别, A.性别), F.性别),D.性别) 性别,NVl(NVl(decode(A.病人来源,4,D.年龄, A.年龄), F.年龄),D.年龄) 年龄 , d.就诊卡号, d.身份证号, d.Ic卡号, d.医保号," & vbNewLine & _
    "   Nvl(f.费别, d.费别) As 费别, Decode(a.病人来源, 1, d.门诊号, 2, Decode(f.病人性质, 1, d.门诊号, f.住院号), 4, d.门诊号, Null) As 标识号," & vbNewLine & _
    "   f.出院病床 As 床号,d.结算模式, Decode(a.病人来源, 1, '门诊', 2, '住院', 3, '外来', 4, '体检') As 来源, a.门诊记帐, a.数据转出, c.名称 As 内容, a.完成人, a.完成时间,A.接收时间," & vbNewLine & _
    "   a.执行间, f.当前病区id As 病区id, f.出院日期, f.状态, c.执行安排, Nvl(z.病历文件id, 0) As 文件id, Nvl(y.通用, 0) As 报告项, NVL(f.病人类型,D.病人类型) As 病人类型,f.审核标志, h.名称 As 险类," & vbNewLine & _
    "   Decode(f.路径状态, Null, 0, 1) As 路径, a.医嘱内容, a.标本部位, a.检查方法, f.病人性质, a.开嘱医生,F.病案状态, " & vbNewLine & _
    "   Decode(f.出院方式, Null, Decode(f.状态, 1, 0, 3, 3, 2), Decode(f.出院方式, '死亡', 5, 4)) As 类型,F.挂号ID as 主页挂号ID" & _
    " From (" & strSQL & ") A,诊疗项目目录 C,病人信息 D,病案主页 F,部门表 E,部门表 G,病历单据应用 Z,病历文件列表 Y,保险类别 H" & _
    " Where A.诊疗项目ID=C.ID And A.病人ID=D.病人ID And A.执行部门ID=G.ID " & _
            IIf(mvarCond.标识号 <> "", " And Decode(A.病人来源,1,D.门诊号,2,Decode(F.病人性质,1,D.门诊号,F.住院号),3,D.门诊号,4,D.门诊号,NULL)=[13]", "") & _
            IIf(mvarCond.就诊卡 <> "", " And D.就诊卡号||''=[14]", "") & IIf(mvarCond.姓名 <> "", " And D.姓名||''=[15]", "") & _
            IIf(mvarCond.身份证 <> "", " And D.身份证号||''=[16]", "") & IIf(mvarCond.IC卡号 <> "", " And D.IC卡号||''=[17]", "") & _
            IIf(mvarCond.医保号 <> "", " And D.医保号||''=[18]", "") & IIf(mvarCond.病人ID <> 0, " And A.病人ID=[22]", "") & _
            IIf(mvarCond.本次, " And (A.病人来源=2 And A.主页ID=D.主页ID Or Nvl(A.病人来源,0)<>2)", "") & str病人来源 & strWhere核对 & _
    "       And A.病人科室ID=E.ID And A.病人ID=F.病人ID(+) And A.主页ID=F.主页ID(+)" & IIf(lng病区ID <> 0, " And F.当前病区ID+0=[20]", "") & _
    "       And A.诊疗项目ID=Z.诊疗项目ID(+) And A.病人来源=Z.应用场合(+) And Z.病历文件ID=Y.ID(+) And D.险类=H.序号(+)" & _
    "       And Not(A.诊疗类别='Z' And Nvl(C.操作类型,'0')<>'0') " & _
    " Order by 发送时间 Desc,病人ID,序号"
        
    Screen.MousePointer = 11
    On Error GoTo errH
    
    Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboDept.ItemData(cboDept.ListIndex), "," & str来源 & ",", _
        IIf(Mid(mstr状态, 1, 1) = "1", 2, -1), IIf(Mid(mstr状态, 2, 1) = "1", 0, -1), IIf(Mid(mstr状态, 3, 1) = "1", 3, -1), IIf(Mid(mstr状态, 4, 1) = "1", 1, -1), _
        "|" & mstrRoom & "|", datBegin, datEnd, mvarCond.NO, mvarCond.科室ID, "," & mvarCond.来源 & ",", mvarCond.标识号, mvarCond.就诊卡, _
        mvarCond.姓名, mvarCond.身份证, mvarCond.IC卡号, mvarCond.医保号, mvarCond.期效 - 1, lng病区ID, mvarCond.开单人, mvarCond.病人ID)
    
    If blnNoFilter Then mvarCond.开单人 = ""
    
    '记录现在选中的病人
    If rptPati.SelectedRows.Count > 0 Then
        If Not rptPati.SelectedRows(0).GroupRow Then
            lngPatiRow = rptPati.SelectedRows(0).Index '用于快速重新定位
            strPatiRow = rptPati.SelectedRows(0).Record.Tag
            If Not rptPati.SelectedRows(0).ParentRow.GroupRow Then
                '记录当前子项所展开的父项，展开与否影响Rows
                strExpend = rptPati.SelectedRows(0).ParentRow.Record.Tag
            End If
        End If
    End If
    rptPati.Records.DeleteAll
    rptPati.Columns.Find(col_内容).TreeColumn = False
    
    '刷新后分组自动展开
    For i = 1 To rsPati.RecordCount
        '是否只显示已收费的病人
        '1.只管主费用,不判断附加费用
        '2.只管收费划价单据，不记帐划价单据(因为需要这里执行后审核)
        '3.不管直接把划价单删了的情况(费用性质=NULL)
        '4.无需计费或尚未生成主费用的也显示
        blnDo = True: blnSub = False
        '组合医嘱只显示一行(为加快速度不用SQL处理)
        If Not objPreRecord Is Nothing Then
            If rsPati!诊疗类别 = "C" And Not IsNull(rsPati!相关ID) _
                And objPreRecord(COL_诊疗类别).Value = "C" And objPreRecord(col_相关ID).Value = Nvl(rsPati!相关ID, 0) Then
                objPreRecord(col_内容).Value = objPreRecord(col_内容).Value & "," & rsPati!内容
                blnDo = True: blnSub = True '一并采集的检验树形显示
                If Not rptPati.Columns.Find(col_内容).TreeColumn Then rptPati.Columns.Find(col_内容).TreeColumn = True
            ElseIf rsPati!诊疗类别 = "D" Then
                If Not IsNull(rsPati!相关ID) And objPreRecord(col_医嘱ID).Value = Nvl(rsPati!相关ID, 0) Then
                    blnDo = True: blnSub = True '检查和方法部位树形显示
                    If Not rptPati.Columns.Find(col_内容).TreeColumn Then rptPati.Columns.Find(col_内容).TreeColumn = True
                Else
                    blnDo = True
                End If
            ElseIf rsPati!诊疗类别 = "F" And Not IsNull(rsPati!相关ID) Then
                blnDo = False '手术和附加手术
            End If
        End If
                
        If blnDo Then
            If blnSub Then
                '第一检验组合和多部位检查项目复制为子项行
                If rsPati!诊疗类别 = "C" Or rsPati!诊疗类别 = "D" Then
                    If objPreRecord.Childs.Count = 0 Then
                        Set objRecord = objPreRecord.Childs.Add()
                        objRecord.Tag = "Sub_" & objPreRecord.Tag
                        For j = 0 To rptPati.Columns.Count - 1
                            Set objItem = objRecord.AddItem(objPreRecord(j).Value)
                            objItem.Caption = objPreRecord(j).Caption
                            objItem.Icon = objPreRecord(j).Icon
                            objItem.HasCheckbox = objPreRecord(j).HasCheckbox
                            objItem.ForeColor = objPreRecord(j).ForeColor
                            If j = col_内容 Then
                                If rsPati!诊疗类别 = "D" Then
                                    objItem.Value = "" & rsPati!内容
                                Else
                                    objItem.Value = Replace(objItem.Value, "," & Nvl(rsPati!内容), "")
                                End If
                            End If
                        Next
                    End If
                End If
                Set objRecord = objPreRecord.Childs.Add()
                objRecord.Tag = CStr("Sub_" & rsPati!医嘱ID & "_" & rsPati!发送号) '用于病人定位
            Else
                Set objRecord = Me.rptPati.Records.Add()
                objRecord.Tag = CStr("_" & rsPati!医嘱ID & "_" & rsPati!发送号) '用于病人定位
                If objRecord.Tag = strExpend Then objRecord.Expanded = True
            End If
            
            '分组
            Set objItem = objRecord.AddItem(Val(Nvl(rsPati!执行状态, 0))) '分组以Value进行排序
            objItem.Caption = Decode(objItem.Value, 0, "未执行", 1, "已执行", 2, "拒绝执行", 3, "正在执行")
            
            '选择
            Set objItem = objRecord.AddItem("")
            objItem.HasCheckbox = True
            If mblnFilter And Not blnSub And Val(Nvl(rsPati!执行状态, 0)) = 0 Then
                objItem.Checked = True '非子项缺省选中(未执行)
                rptPati.Columns(col_选择).Caption = "1"
            End If
            
            '路径
            Set objItem = objRecord.AddItem("")
            objItem.Value = Val("" & rsPati!路径)
            objItem.Caption = " "
            If rsPati!路径 = 1 Then
                objItem.Icon = img16.ListImages("Path").Index - 1
                If blnPath = False Then blnPath = True
            End If
            
            '执行状态
            Set objItem = objRecord.AddItem(Val(Nvl(rsPati!执行状态, 0)))
            
            '图标
            Set objItem = objRecord.AddItem("")
            objItem.Icon = Nvl(rsPati!执行状态, 0) 'ImageList是从1开始,用于ReportControl时是从0开始
            If Nvl(rsPati!执行状态, 0) = 0 And Nvl(rsPati!执行过程, 0) = 1 Then
                objItem.Icon = 5 '已报到
            End If
            
            If objItem.Icon = 3 Then
                '已核对的图标只对正在执行的生效
                If Val(gstr医嘱核对) > 0 Then
                    If rsPati!接收人 & "" <> "" Then
                        If rsPati!操作类型 & "" = "1" And rsPati!诊疗类别 & "" = "E" And Mid(gstr医嘱核对, 2, 1) = "1" Or _
                            rsPati!操作类型 & "" = "8" And rsPati!诊疗类别 & "" = "E" And Mid(gstr医嘱核对, 1, 1) = "1" Or _
                            rsPati!诊疗类别 & "" = "K" And Mid(gstr医嘱核对, 1, 1) = "1" Then
                            
                            objRecord(col_图标).Icon = 7
                        End If
                    End If
                End If
            End If
            
            If Nvl(rsPati!来源) = "住院" And Val("" & rsPati!病人性质) <> 1 Then mblnShowBed = True
            Set objItem = objRecord.AddItem(CStr(Nvl(rsPati!来源)))
            If Nvl(rsPati!来源) = "住院" And Val("" & rsPati!病人性质) = 1 Then objItem.Caption = "门诊"    'value仍为住院
            
            objRecord.AddItem CStr(Nvl(rsPati!NO))
            objRecord.AddItem IIf(Val(rsPati!紧急标志 & "") = 1, "急", "")
            objRecord.AddItem CStr(Nvl(rsPati!姓名))
'            objRecord.AddItem "急"
            If rsPati!诊疗类别 = "D" Then
                If IsNull(rsPati!相关ID) Then
                    objRecord.AddItem "" & rsPati!医嘱内容  '文字中包含了所有的部位，方法
                Else
                    If IsNull(rsPati!检查方法) Then
                        objRecord.AddItem "" & rsPati!标本部位
                    Else
                        objRecord.AddItem rsPati!标本部位 & "(" & rsPati!检查方法 & ")"
                    End If
                End If
            Else
                objRecord.AddItem "" & rsPati!内容
            End If
            
            objRecord.AddItem CStr(Nvl(rsPati!险类))
            
            Set objItem = objRecord.AddItem(Val(Nvl(rsPati!病人科室ID, 0)))
            objItem.Caption = Nvl(rsPati!科室, " ")
            
            Set objItem = objRecord.AddItem(CStr(Nvl(rsPati!标识号)))
            objItem.Caption = Nvl(rsPati!标识号, " ")
            
            objRecord.AddItem CStr(Nvl(rsPati!床号))
            objRecord.AddItem CStr(Nvl(rsPati!费别))
            
            Set objItem = objRecord.AddItem(Format(rsPati!要求时间, "yyyy-MM-dd HH:mm:ss"))
            objItem.Caption = Format(rsPati!要求时间, "yyyy-MM-dd HH:mm")
            If Nvl(rsPati!执行状态, 0) = 0 And Not IsNull(rsPati!安排时间) Then
                '未执行的项目，重新安排了要求执行时间的突出显示
                objItem.Bold = True
            End If
            
            Set objItem = objRecord.AddItem(Format(rsPati!发送时间, "yyyy-MM-dd HH:mm:ss"))
            objItem.Caption = Format(rsPati!发送时间, "yyyy-MM-dd HH:mm")
            
            objRecord.AddItem CStr(Nvl(rsPati!执行间))
            objRecord.AddItem CStr(Nvl(rsPati!性别))
            objRecord.AddItem CStr(Nvl(rsPati!年龄))
            objRecord.AddItem CStr(Nvl(rsPati!完成人))
            Set objItem = objRecord.AddItem(Format(rsPati!完成时间, "yyyy-MM-dd HH:mm:ss"))
            objItem.Caption = Format(rsPati!完成时间, "yyyy-MM-dd HH:mm")
            
            '隐藏数据列
            Set objItem = objRecord.AddItem(Val(rsPati!执行部门ID)): objItem.Caption = rsPati!执行科室
            objRecord.AddItem Val(rsPati!病人ID)
            objRecord.AddItem Val(Nvl(rsPati!主页ID, 0))
            objRecord.AddItem CStr(Nvl(rsPati!挂号单))
            objRecord.AddItem 0 '挂号ID，在进入行时读取
            objRecord.AddItem Val(Nvl(rsPati!婴儿, 0))
            objRecord.AddItem CStr(Nvl(rsPati!就诊卡号))
            objRecord.AddItem CStr(Nvl(rsPati!身份证号))
            objRecord.AddItem CStr(Nvl(rsPati!IC卡号))
            objRecord.AddItem CStr(Nvl(rsPati!医保号))
            objRecord.AddItem Val(Nvl(rsPati!病区ID, 0))
            objRecord.AddItem Format(Nvl(rsPati!出院日期, ""), "yyyy-MM-dd HH:mm:ss")
            objRecord.AddItem Val(Nvl(rsPati!状态, 0))
            objRecord.AddItem Val(rsPati!医嘱ID)
            objRecord.AddItem Val(Nvl(rsPati!相关ID, 0))
            objRecord.AddItem Val(rsPati!发送号)
            objRecord.AddItem CStr(rsPati!诊疗类别)
            objRecord.AddItem Val(Nvl(rsPati!执行过程, 0))
            objRecord.AddItem Val(Nvl(rsPati!执行安排, 0))
            objRecord.AddItem Val(Nvl(rsPati!记录性质, 1))
            objRecord.AddItem Val(Nvl(rsPati!数据转出, 0))
            
            objRecord.AddItem Val(Nvl(rsPati!文件ID, 0))
            objRecord.AddItem Val(Nvl(rsPati!报告项, 0))
            objRecord.AddItem 0 '报告ID，在进入行时读取
            objRecord.AddItem CStr(Nvl(rsPati!病人类型))
            objRecord.AddItem Val("" & rsPati!门诊记帐)
            
            objRecord.AddItem CStr("" & rsPati!开嘱医生)
            Set objItem = objRecord.AddItem(Val(rsPati!病案状态 & ""))
            objItem.Caption = " "
            objRecord.AddItem (Val(rsPati!类型))
            objRecord.AddItem rsPati!操作类型 & ""
            If InStr(",1,8,", rsPati!操作类型 & "") > 0 And rsPati!诊疗类别 & "" = "E" Or rsPati!诊疗类别 & "" = "K" Then
                objRecord.AddItem rsPati!接收人 & ""
            Else
                objRecord.AddItem ""
            End If
            objRecord.AddItem rsPati!审核标志 & ""
            objRecord.AddItem Format(rsPati!接收时间 & "", "yyyy-MM-dd HH:mm")
            objRecord.AddItem rsPati!结算模式 & ""
            objRecord.AddItem rsPati!诊疗项目ID & ""
            objRecord.AddItem rsPati!期效 & ""
            objRecord.AddItem rsPati!执行分类 & ""
	    objRecord.AddItem rsPati!主页挂号ID & ""
            objRecord.AddItem ""
            '病人颜色:黑色,灰色,棕色,兰色
'            objRecord.Item(0).ForeColor = Decode(Nvl(rsPati!执行状态, 0), 0, 0, 1, &H808080, 2, &H40C0&, 3, &HC00000)
'            For j = 0 To rptPati.Columns.Count - 1
'                objRecord.Item(j).ForeColor = objRecord.Item(0).ForeColor
'            Next
            If Not IsNull(rsPati!病人类型) Then
                '保险病人用指定色显示
                lngColor = zlDatabase.GetPatiColor(rsPati!病人类型)
                objRecord.Item(col_单据号).ForeColor = lngColor
                objRecord.Item(col_病人类型).ForeColor = lngColor
            ElseIf Not IsNull(rsPati!险类) Then
                '未指定病人类型的保险病人用红色显示
                For j = 0 To rptPati.Columns.Count - 1
                    objRecord.Item(j).ForeColor = vbRed
                Next
            End If
            
            '他科执行的表示
            If Val(rsPati!执行部门ID) <> cboDept.ItemData(cboDept.ListIndex) Then
                objRecord.Item(col_内容).Value = objRecord.Item(col_执行科室).Caption & "：" & objRecord.Item(col_内容).Value
            End If
            If Val(rsPati!紧急标志 & "") = 1 Then
                objRecord.Item(col_紧急).ForeColor = vbRed
            End If
            If Not blnSub Then Set objPreRecord = objRecord
        End If
        rsPati.MoveNext
    Next
    
    '如果没有临床路径病人，则隐藏列
    rptPati.Columns(col_路径).Visible = blnPath
    
    
    '一并检验和多部位检查的综合分组状态
    If rptPati.Columns.Find(col_内容).TreeColumn Then
        For Each objRecord In rptPati.Records
            If objRecord.Childs.Count > 0 Then
                strSQL = ""
                For Each objPreRecord In objRecord.Childs
                    If InStr(strSQL, objPreRecord(col_执行状态).Value) = 0 Then
                        strSQL = strSQL & objPreRecord(col_执行状态).Value
                    End If
                Next
                '拒绝执行只能整体进行，其他混合状态显示为正在执行
                '子项是跟随父项所在的分组出现
                objRecord(col_综合状态).Value = IIf(Len(strSQL) = 1, Val(strSQL), 3)
                objRecord(col_综合状态).Caption = Decode(objRecord(col_综合状态).Value, 0, "未执行", 1, "已执行", 2, "拒绝执行", 3, "正在执行")
                objRecord(col_图标).Icon = objRecord(col_综合状态).Value
'                objRecord.Item(0).ForeColor = Decode(objRecord(col_综合状态).Value, 0, 0, 1, &H808080, 2, &H40C0&, 3, &HC00000)
'                For j = 0 To rptPati.Columns.Count - 1
'                    objRecord.Item(j).ForeColor = objRecord.Item(0).ForeColor
'                Next
                If objRecord(col_病人类型).Value <> "" Then
                    '保险病人用采色显示
                    lngColor = zlDatabase.GetPatiColor(objRecord(col_病人类型).Value)
                    objRecord.Item(col_单据号).ForeColor = lngColor
                    objRecord.Item(col_病人类型).ForeColor = lngColor
                ElseIf objRecord(col_险类).Value <> "" Then
                    '未指定病人类型的保险病人用红色显示
                    For j = 0 To rptPati.Columns.Count - 1
                        objRecord.Item(j).ForeColor = vbRed
                    Next
                End If
            End If
        Next
    End If
    
    rptPati.Columns.Find(col_床号).Visible = mblnShowBed
    rptPati.Populate
    
    '定位病人行:在Populate之后
    mstrPrePati = ""
    If rptPati.Rows.Count = 0 Then
        '按无数据刷新子窗体
        Call ClearPatiInfo
        Call SubWinRefreshData(tbcSub.Selected)
    Else
        '取指定病人行
        If strPatiRow <> "" Then
            '先快速定位
            If lngPatiRow <= rptPati.Rows.Count - 1 Then
                If Not rptPati.Rows(lngPatiRow).GroupRow Then
                    If rptPati.Rows(lngPatiRow).Record.Tag = strPatiRow Then
                        Set objRow = rptPati.Rows(lngPatiRow)
                    End If
                End If
            End If
            '再进行查找
            If objRow Is Nothing Then
                For i = 0 To rptPati.Rows.Count - 1
                    If Not rptPati.Rows(i).GroupRow Then
                        If rptPati.Rows(i).Record.Tag = strPatiRow Then
                            Set objRow = rptPati.Rows(i): Exit For
                        End If
                    End If
                Next
            End If
        End If
        '取第一个非分组行
        If objRow Is Nothing Then
            For i = 0 To rptPati.Rows.Count - 1
                If Not rptPati.Rows(i).GroupRow Then Set objRow = rptPati.Rows(i): Exit For
            Next
        End If
        If Not objRow.ParentRow.GroupRow Then objRow.ParentRow.Expanded = True
        Set rptPati.FocusedRow = objRow '该行选中且显示在可见区域,并引发SelectionChanged事件
    End If
    
    stbThis.Panels(2).Text = " 共 " & rptPati.Records.Count & " 个病人项目"
    Screen.MousePointer = 0
    LoadPatients = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadNotify() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim strSQL As String, i As Long, j As Long
    Dim strTmp As String
    Dim strMsgType As String
    
    On Error GoTo errH
    rptNotify.Records.DeleteAll
    If cboDept.ListIndex = -1 Or cboUnit.ListIndex = -1 Then LoadNotify = True: Exit Function
    If Mid(mstrNotify, m销帐申请, 1) = "1" Then strTmp = strTmp & ",ZLHIS_CHARGE_001"
    If Mid(mstrNotify, m待安排, 1) = "1" Then strTmp = strTmp & ",ZLHIS_CIS_004"
    If Mid(mstrNotify, m血袋回收, 1) = "1" Then strTmp = strTmp & ",ZLHIS_BLOOD_007"
    strTmp = Mid(strTmp, 2)
    If strTmp = "" Then LoadNotify = True: Exit Function
    
    strSQL = "Select b.病人id, b.就诊id as 主页id,a.住院号,a.姓名, a.性别, a.年龄, a.当前床号 As 床号, Nvl(b.就诊科室id, a.当前科室id) As 就诊科室id," & _
        " Nvl(b.就诊病区id, a.当前病区id) As 就诊病区id, b.病人来源, b.消息内容, b.类型编码, b.业务标识, b.优先程度, b.登记时间,a.险类" & _
        " From 病人信息 A, 业务消息清单 B, 业务消息提醒部门 C, 业务消息提醒人员 D" & _
        " Where a.病人id = b.病人id And b.Id = c.消息id And b.Id = d.消息id(+) And b.登记时间 >=Trunc(Sysdate-" & (mintDay - 1) & ") and substr(b.提醒场合,[4],1)='1'" & _
        " And Nvl(b.是否已阅, 0) = 0  And instr(','||[5]||',',','||b.类型编码||',')>0 and (c.部门id = [1] Or d.提醒人员 = [3])" & _
        " Order By b.优先程度, b.登记时间 Desc"
    
    Screen.MousePointer = 11
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, cboDept.ItemData(cboDept.ListIndex), , UserInfo.姓名, 4, strTmp)
    
    If cboDept.ListIndex <> -1 Then
        strTmp = ","
        For i = 1 To rsTmp.RecordCount
            If InStr(strTmp, "," & rsTmp!病人ID & "," & rsTmp!主页ID & "," & rsTmp!类型编码 & ",") = 0 Then
                strTmp = strTmp & rsTmp!病人ID & "," & rsTmp!主页ID & "," & rsTmp!类型编码 & ","
                Call AddReportRow(rsTmp!病人ID & "," & rsTmp!主页ID, rsTmp!病人ID, rsTmp!主页ID, Nvl(rsTmp!姓名), Nvl(rsTmp!消息内容), rsTmp!类型编码, _
                        rsTmp!优先程度 & "", Format(rsTmp!登记时间 & "", "yyyy-MM-dd HH:mm:ss"), rsTmp!业务标识 & "")
            End If
            rsTmp.MoveNext
        Next
    End If
    rptNotify.Populate '缺省不选中任何行
    rptNotify.TabStop = rptNotify.Rows.Count > 0
    Screen.MousePointer = 0
    LoadNotify = True
    If mbln消息语音 Then
        If mclsMsg Is Nothing Then
            Set mclsMsg = New clsCISMsg
            Call mclsMsg.InitCISMsg(3)
        End If
        If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            Set mrsMsg = rsTmp
        End If
    End If
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ClearPatiInfo()
'功能：清除单个病人相关的显示信息
    mlng病人ID = 0
    mlng主页ID = 0
    mstr挂号单 = ""
    
    lblAdvice.Caption = ""
    lblDiag(1).Caption = ""
    lblCash.Visible = False
    lblRec.Visible = False
        
    vsExec.Rows = vsExec.FixedRows
    vsExec.Rows = vsExec.FixedRows + 1
    picBlood.Visible = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("[']", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    If (InStr("0123456789", Chr(KeyAscii)) > 0 Or UCase(Chr(KeyAscii)) >= "A" And UCase(Chr(KeyAscii)) <= "Z") _
        And Not (Me.ActiveControl Is PatiIdentify Or Me.ActiveControl Is rtfAppend Or Me.ActiveControl Is cboDept Or Me.ActiveControl Is cboUnit) And mstrFindType = "就诊卡" Then
        PatiIdentify.Text = UCase(Chr(KeyAscii))
        PatiIdentify.SetFocus
        Call zlCommFun.PressKey(vbKeyRight)
    End If
End Sub

Private Sub timBRefresh_Timer()
    '供血库输血执行窗体填写完执行内容后，医嘱对应内容的刷新
    Dim intState As Integer
    timBRefresh.Enabled = False
    If Not mobjFrmBloodExe Is Nothing Then
        On Error Resume Next
        intState = mobjFrmBloodExe.AdviceExecState
        If err <> 0 Then
            err.Clear
        Else
            mobjFrmBloodExe.ExecFresh = True
            Select Case intState
                Case 1, 2 '记录执行或调整执行，删除执行
                    Call LoadPatients '要更新执行状态
                Case 3, 4 '执行完成,取消完成
                    Call LoadPatients '要更新执行状态
            End Select
            mobjFrmBloodExe.ExecFresh = False
            mobjFrmBloodExe.AdviceExecState = 0
        End If
    End If
End Sub


Private Sub timRefresh_Timer()
    Static lngSec病人列表 As Long
    Static strTim消息列表 As String
    
    Dim curTime As Date
    
    If mintRefresh <> 0 Then
        lngSec病人列表 = lngSec病人列表 + 1 '秒数
        If lngSec病人列表 Mod mintRefresh = 0 Then
            lngSec病人列表 = 0
            Call LoadPatients
        End If
    End If
    
    If mbln消息语音 Then
        If Not mrsMsg Is Nothing Then
            If mrsMsg.RecordCount > 0 Then
                timRefresh.Enabled = False
                Call mclsMsg.PlayMsgSound(mrsMsg)
                Set mrsMsg = Nothing
                timRefresh.Enabled = True
            End If
        End If
    End If
    
    If Not mclsMipModule Is Nothing Then
        If mclsMipModule.IsConnect Then '使用了消息平台不用自动刷新消息列表
            Exit Sub
        End If
    End If
    
    If mintMin > 0 And rptNotify.Visible Then
        curTime = Now
        
        If strTim消息列表 = "" Then
            strTim消息列表 = Format(curTime, "yyyy-MM-dd HH:mm:ss")
        End If
        If DateDiff("s", CDate(strTim消息列表), curTime) > mintMin * CLng(60) Then
            strTim消息列表 = Format(curTime, "yyyy-MM-dd HH:mm:ss")
            Call LoadNotify
        End If
    End If
    
End Sub

Private Sub ExecuteFindPati(Optional ByVal blnNext As Boolean, Optional ByVal strIDCard As String, Optional ByVal lngPatiID As Long)
'功能：查找(下一个)病人
'参数：blnNext=是否查找下一个
'      strIDCard=当有值时，表示固定按身份证号查找
    Static blnReStart As Boolean
    Dim blnHave As Boolean, i As Long
    
    '按其他方式查找后，自动刷身份证的继续查找则取消
    If strIDCard = "" And PatiIdentify.Text <> "" Then mvarCond.身份证 = ""
    
    If Not blnNext And mstrFindType = "单据号" Then
        PatiIdentify.Text = GetFullNO(PatiIdentify.Text, 12)
    End If
    PatiIdentify.SetFocus
            
    '过滤模式时，以指定查找条件读取病人清单
    If mblnFilter Then
        Call ClearPatiCond '其他在过滤中设置的身份识别条件清除
        If strIDCard <> "" Then '身份证自动识别强制优先
            mvarCond.身份证 = strIDCard
        Else
            Select Case mstrFindType
                Case "就诊卡"
                    mvarCond.就诊卡 = PatiIdentify.Text '就诊卡
                Case "标识号"
                    mvarCond.标识号 = PatiIdentify.Text '标识号
                Case "单据号"
                    mvarCond.NO = PatiIdentify.Text '单据号
                Case "姓名"
                    mvarCond.姓名 = PatiIdentify.Text '姓名
                Case "二代身份证"
                    mvarCond.身份证 = PatiIdentify.Text '身份证
                Case "IC卡"
                    If Not mobjSquareCard Is Nothing Then 'IC卡
                        Call mobjSquareCard.zlGetPatiID("IC卡", PatiIdentify.Text, , mvarCond.病人ID)
                    Else
                        mvarCond.IC卡号 = PatiIdentify.Text
                    End If
                Case "医保号"
                    mvarCond.医保号 = PatiIdentify.Text '医保号
                Case Else
                    If Not mobjSquareCard Is Nothing Then
                        Call mobjSquareCard.zlGetPatiID(Val(PatiIdentify.objIDKind.GetCurCard.接口序号), PatiIdentify.Text, , mvarCond.病人ID)
                    End If
            End Select
        End If
        Call LoadPatients
        Exit Sub
    End If
    
    '开始查找行
    If rptPati.SelectedRows.Count > 0 Then
        If Not rptPati.SelectedRows(0).GroupRow Then blnHave = True
    End If
    If Not blnNext Or blnReStart Or Not blnHave Then
        i = 0 'ReportControl的索引从是0开始
    Else
        i = rptPati.SelectedRows(0).Index + 1
    End If
    
    '查找病人
    If lngPatiID = 0 And Not mobjSquareCard Is Nothing And mstrFindType <> "就诊卡" And mstrFindType <> "标识号" And mstrFindType <> "单据号" And mstrFindType <> "姓名" And mstrFindType <> "二代身份证" And mstrFindType <> "医保号" Then
        If mstrFindType = "IC卡" Then
            Call mobjSquareCard.zlGetPatiID("IC卡", PatiIdentify.Text, , lngPatiID)
        Else
            Call mobjSquareCard.zlGetPatiID(Val(PatiIdentify.objIDKind.GetCurCard.接口序号), PatiIdentify.Text, , lngPatiID)
        End If
    End If
    
    With rptPati
        For i = i To .Rows.Count - 1
            If Not .Rows(i).GroupRow Then
                If .Rows(i).ParentRow.GroupRow Then
                    If strIDCard <> "" Then '身份证自动识别强制优先
                        If UCase(.Rows(i).Record(col_身份证号).Value) = UCase(strIDCard) Then Exit For
                    Else
                        If Val(.Rows(i).Record(col_病人Id).Value) = lngPatiID And lngPatiID <> 0 Then Exit For
                        Select Case mstrFindType
                            Case "就诊卡"
                                If .Rows(i).Record(col_就诊卡号).Value = PatiIdentify.Text Then Exit For
                            Case "标识号"
                                If .Rows(i).Record(col_标识号).Value = PatiIdentify.Text Then Exit For
                            Case "单据号"
                                If UCase(.Rows(i).Record(col_单据号).Value) = UCase(PatiIdentify.Text) Then Exit For
                            Case "姓名"
                                If .Rows(i).Record(col_姓名).Value Like "*" & PatiIdentify.Text & "*" Then Exit For
                            Case "二代身份证"
                                If UCase(.Rows(i).Record(col_身份证号).Value) = UCase(PatiIdentify.Text) Then Exit For
                            Case "医保号"
                                If UCase(.Rows(i).Record(col_医保号).Value) = UCase(PatiIdentify.Text) Then Exit For
                            Case Else
                                If Val(.Rows(i).Record(col_病人Id).Value) = lngPatiID Then Exit For
                        End Select
                    End If
                End If
            End If
        Next
    End With

    If i <= rptPati.Rows.Count - 1 Then
        blnReStart = False
        '该行选中且显示在可见区域,并引发SelectionChanged事件
        Set rptPati.FocusedRow = rptPati.Rows(i)
        If rptPati.Visible Then rptPati.SetFocus
    Else
        blnReStart = True
        MsgBox IIf(blnNext, "后面已", "") & "找不到符合条件的病人。", vbInformation, gstrSysName
    End If
End Sub

Private Sub ClearPatiCond()
    mvarCond.IC卡号 = "": mvarCond.NO = "": mvarCond.标识号 = ""
    mvarCond.就诊卡 = "": mvarCond.身份证 = "": mvarCond.姓名 = ""
    mvarCond.医保号 = "": mvarCond.开单人 = "": mvarCond.病人ID = 0
End Sub

Private Sub PatientFilter()
    timRefresh.Enabled = False
    frmTechnicFilter.mstrDeptNode = mstrDeptNode
    frmTechnicFilter.mstrPrivs = mstrPrivs
    frmTechnicFilter.mstrCardKind = mstrCardKind
    Set frmTechnicFilter.mobjSquareCard = mobjSquareCard
    frmTechnicFilter.Show 1, Me '怪,不激活过滤窗口的Form_Activate事件
    If frmTechnicFilter.mblnOK Then
        '重置过滤变量
        With frmTechnicFilter
            '发送时间
            mvarCond.Begin = Format(.dtpBegin.Value, "yyyy-MM-dd HH:mm:00")
            If Format(.dtpEnd.Value, "yyyy-MM-dd HH:mm") = Format(.dtpEnd.Tag, "yyyy-MM-dd HH:mm") Then
                mvarCond.End = CDate(0) '表示取当前时间
            Else
                mvarCond.End = Format(.dtpEnd.Value, "yyyy-MM-dd HH:mm:59")
            End If
            
            Call ClearPatiCond
            If .mlngPatiID <> 0 Then
                '当缺省类与刷卡类不一致时:选择姓名实际是刷就诊卡
                mvarCond.病人ID = .mlngPatiID
            Else
                Select Case .mstrFindType
                    Case "就诊卡"
                        mvarCond.就诊卡 = .PatiIdentify.Text '就诊卡
                    Case "标识号"
                        mvarCond.标识号 = .PatiIdentify.Text '标识号
                    Case "单据号"
                        mvarCond.NO = .PatiIdentify.Text '单据号
                    Case "姓名"
                        mvarCond.姓名 = .PatiIdentify.Text '姓名
                    Case "二代身份证"
                        mvarCond.身份证 = .PatiIdentify.Text '身份证
                    Case "IC卡"
                        mvarCond.IC卡号 = .PatiIdentify.Text
                    Case "医保号"
                        mvarCond.医保号 = .PatiIdentify.Text '医保号
                    Case Else
                        If Not mobjSquareCard Is Nothing Then
                            Call mobjSquareCard.zlGetPatiID(Val(.PatiIdentify.objIDKind.GetCurCard.接口序号), .PatiIdentify.Text, , mvarCond.病人ID)
                        End If
                End Select
            End If
            '病人科室
            If .cboDept.ListIndex <> 0 Then
                mvarCond.科室ID = .cboDept.ItemData(.cboDept.ListIndex)
            Else
                mvarCond.科室ID = 0
            End If
            
            '病人来源
            mvarCond.来源 = ""
            If Not (.chk来源(0).Value = 1 And .chk来源(1).Value = 1 And .chk来源(2).Value = 1) Then
                If .chk来源(0).Value = 1 Then mvarCond.来源 = mvarCond.来源 & ",1"
                If .chk来源(1).Value = 1 Then mvarCond.来源 = mvarCond.来源 & ",2"
                If .chk来源(2).Value = 1 Then mvarCond.来源 = mvarCond.来源 & ",4"
                mvarCond.来源 = Mid(mvarCond.来源 & ",3", 2)
            End If
            
            '本次住院
            mvarCond.本次 = .chk本次住院.Value = 1
            
            '医嘱期效
            mvarCond.期效 = 0
            If Not (.chk期效(0).Value = 1 And .chk期效(1).Value = 1) Then
                If .chk期效(0).Value = 1 Then
                    mvarCond.期效 = 1
                ElseIf .chk来源(1).Value = 1 Then
                    mvarCond.期效 = 2
                End If
            End If
            
            '开单人
            mvarCond.开单人 = ""
            If .cboDoctor.Text <> "" And .cboDoctor.ListIndex <> 0 Then
                mvarCond.开单人 = Split(.cboDoctor.Text, "-")(1)
            End If
        End With
         '外面的即时查找条件清除
        Me.PatiIdentify.Text = ""
        mstr开单人 = mvarCond.开单人
                        
        Call SetUnitVisible
        Call LoadPatients '刷新
                
        '没有病人时缺省显示的医嘱页面
        If rptPati.Rows.Count = 0 Then
            If mvarCond.来源 = "2,3" Then
                Call ExchangeAdvice(False) '缺省显示住院的
            Else
                Call ExchangeAdvice(True) '缺省显示门诊的
            End If
        End If
    End If
    timRefresh.Enabled = True
End Sub

Private Sub SetUnitVisible()
    If Not (mvarCond.来源 = "2,3") Then
        If cboUnit.ListCount > 0 Then
            Call Cbo.SetIndex(cboUnit.hwnd, 0)
        End If
    End If
    
    lblUnit.Visible = mvarCond.来源 = "2,3"
    cboUnit.Visible = mvarCond.来源 = "2,3"
    Call picPati_Resize
    Me.Refresh
End Sub

Private Sub ParameterSetup()
    Dim strRoom As String, str诊疗类别 As String, str治疗类别 As String
    
    timRefresh.Enabled = False
    frmTechnicSetup.mstrPrivs = mstrPrivs
    frmTechnicSetup.mlng科室ID = cboDept.ItemData(cboDept.ListIndex)
    frmTechnicSetup.Show 1, Me
    If frmTechnicSetup.mblnOK Then
        '严格打求记录执行的情况
        mblnExeLog = Val(zlDatabase.GetPara("记录执行情况", glngSys, p医技工作站, "0")) <> 0
        
        '皮试验证身份
        mbln皮试验证 = Val(zlDatabase.GetPara("皮试验证身份", glngSys, p医技工作站)) <> 0
        
        str诊疗类别 = zlDatabase.GetPara("诊疗类别", glngSys, p医技工作站)
        str治疗类别 = zlDatabase.GetPara("治疗类别", glngSys, p医技工作站)
    
        '执行间范围改变
        strRoom = zlDatabase.GetPara("执行间范围", glngSys, p医技工作站)
        If strRoom <> mstrRoom Or str诊疗类别 <> mstr诊疗类别 Or str治疗类别 <> mstr治疗类别 Then
            mstrRoom = strRoom
            mstr诊疗类别 = str诊疗类别
            mstr治疗类别 = str治疗类别
            Call LoadPatients
        End If
        
        '设置自动刷新
        Call SetTimer
    End If
    timRefresh.Enabled = True
End Sub

Private Function Get执行内容(ByVal lng发送号 As Long, ByVal lng医嘱ID As Long, ByVal lng相关ID As Long, ByVal str类别 As String, ByVal objRow As ReportRow) As String
'功能：根据指定的医嘱ID,返回医嘱内容供显示
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strTmp As String
    Dim bln给药途径 As Boolean, i As Integer
    Dim str皮试结果 As String

    On Error GoTo errH
    
    '读取医嘱内容
    If (str类别 = "C" And lng相关ID <> 0) Or str类别 = "D" Then
        strTmp = rptPati.SelectedRows(0).Record(col_内容).Value
        
    ElseIf str类别 <> "E" Or lng相关ID <> 0 Then
        '配方煎法,手术麻醉,输血途径,或其它医嘱,直接显示医嘱内容
        strSQL = "Select 医嘱内容 From 病人医嘱记录 Where ID=[1]"
        If rptPati.SelectedRows(0).Record(COL_数据转出).Value = 1 Then
            strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIf(str类别 = "E", lng相关ID, lng医嘱ID))
        If Not rsTmp.EOF Then strTmp = Nvl(rsTmp!医嘱内容)
    Else
        '类别为E,且相关ID=0
        strSQL = "Select A.ID,A.相关ID,A.诊疗类别,A.医嘱内容,A.皮试结果,A.单次用量,B.计算单位,B.操作类型,A.执行频次,A.执行时间方案,B.名称" & _
            " From 病人医嘱记录 A,诊疗项目目录 B" & _
            " Where Not (A.诊疗类别='E' And 相关ID is Not NULL) And A.诊疗项目ID=B.ID" & _
            " And (A.相关ID=[1] Or A.ID=[1])" & _
            " Order by A.序号"
        If rptPati.SelectedRows(0).Record(COL_数据转出).Value = 1 Then
            strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
        rsTmp.Filter = "相关ID=" & lng医嘱ID
        If Not rsTmp.EOF Then bln给药途径 = InStr(",5,6,", rsTmp!诊疗类别) > 0
        
        If Not bln给药途径 Then
            '一般治疗项目或中药用法，或采集方法
            rsTmp.Filter = 0
            If Not rsTmp.EOF Then
                If rsTmp!诊疗类别 = "E" And rsTmp!操作类型 = "1" Then
                    str皮试结果 = "，皮试结果：" & Nvl(rsTmp!皮试结果)
                    
                    strSQL = "Select b.过敏反应, b.过敏时间 From 病人医嘱记录 A, 病人过敏记录 B, 诊疗项目目录 C, 诊疗用法用量 D" & _
                        " Where a.病人id = b.病人id And a.诊疗项目id = d.用法id And d.项目id = c.Id And c.类别 In ('5', '6') And d.项目id = b.药物id And" & _
                        " Nvl(d.性质, 0) = 0 And b.记录时间 = (Select Max(操作时间) From 病人医嘱状态 Where 医嘱id = a.id And 操作类型 = 10) And a.Id = [1] And RowNum<2"

                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
                    
                    If Not rsTmp.EOF Then
                        str皮试结果 = str皮试结果 & ",过敏时间：" & Format(rsTmp!过敏时间, "yyyy-MM-dd") & IIf(Nvl(rsTmp!过敏反应) = "", "", ",过敏反应：" & rsTmp!过敏反应)
                    End If
                End If
            End If
            
            strSQL = "Select 医嘱内容 From 病人医嘱记录 Where ID=[1]"
            If rptPati.SelectedRows(0).Record(COL_数据转出).Value = 1 Then
                strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
            End If
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
            If Not rsTmp.EOF Then strTmp = Nvl(rsTmp!医嘱内容)
        Else
            '给药途径
            For i = 1 To rsTmp.RecordCount
                strTmp = strTmp & "," & rsTmp!医嘱内容 & IIf(Not IsNull(rsTmp!单次用量), " " & FormatEx(rsTmp!单次用量, 5) & rsTmp!计算单位, "")
                rsTmp.MoveNext
            Next
            rsTmp.Filter = "ID=" & lng医嘱ID
            strTmp = rsTmp!名称 & "," & rsTmp!执行频次 & "(" & rsTmp!执行时间方案 & "):每" & rsTmp!计算单位 & " " & Mid(strTmp, 2)
        End If
    End If
    
    '读取发送数次
    strSQL = "Select A.发送数次,Nvl(D.计算单位,C.计算单位) as 计算单位" & _
        " From 病人医嘱发送 A,病人医嘱记录 B,诊疗项目目录 C,收费项目目录 D" & _
        " Where A.医嘱ID=[1] And A.发送号=[2]" & _
        " And A.医嘱ID=B.ID And B.诊疗项目ID=C.ID And B.收费细目ID=D.ID(+)"
    If rptPati.SelectedRows(0).Record(COL_数据转出).Value = 1 Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
    End If
    'Set rsTmp = New ADODB.Recordset
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID, lng发送号)
    If Not rsTmp.EOF Then
        If IsNull(rsTmp!发送数次) Then
            Get执行内容 = " 执行内容：" & strTmp & str皮试结果
        Else
            Get执行内容 = " 发送数次：" & FormatEx(rsTmp!发送数次, 5) & " " & Nvl(rsTmp!计算单位) & "，执行内容：" & strTmp & str皮试结果
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function LoadExecList(ByVal lng医嘱ID As Long, ByVal lng发送号 As Long) As Boolean
'功能：读取指定医嘱的执行情况表
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strPre As String
    Dim rs血库 As ADODB.Recordset
    Dim bln输血 As Boolean
    Dim int血袋数 As Integer
    
    On Error GoTo errH
    
    '检验项目一并执行时，执行情况登记到第一个项目上。分散单独执行时，登记到各个项目上。
    strSQL = "Select A.要求时间,A.执行时间,A.本次数次,D.计算单位,A.执行摘要,A.执行人,A.登记时间,A.登记人,DECODE(NVL(A.执行结果,1),0,'未执行',1,'完成',2,'拒绝',3,'外出') As 执行结果,a.核对人,a.核对时间,d.操作类型,d.类别,a.说明,a.记录来源 as 来源" & _
        " From 病人医嘱执行 A,病人医嘱发送 B,病人医嘱记录 C,诊疗项目目录 D" & _
        " Where A.医嘱ID=[1] And A.发送号=[2]" & _
        " And A.医嘱ID=B.医嘱ID And A.发送号=B.发送号 And B.医嘱ID=C.ID And C.诊疗项目ID=D.ID" & _
        " Order by A.登记时间 Desc"
    With rptPati.SelectedRows(0)
        If .Record(COL_数据转出).Value = 1 Then
            strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
            strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
            strSQL = Replace(strSQL, "病人医嘱执行", "H病人医嘱执行")
        End If
    End With
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID, lng发送号)
    With vsExec
        strPre = .Cell(flexcpData, .Row, 0)
        .Redraw = flexRDNone
        .Rows = vsExec.FixedRows
        .Rows = vsExec.FixedRows + 1
        .Row = .FixedRows: .Col = .FixedCols
        If Not rsTmp.EOF Then
            .Rows = rsTmp.RecordCount + 1
            '输血医嘱处理流程变动 70823
            If gbln血库系统 And Val(rsTmp!操作类型 & "") = 8 And rsTmp!类别 = "E" Then
                strSQL = "select zl_Get_输血执行次数(相关id) as 数量 from 病人医嘱记录 where id = [1]"
                Set rs血库 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
                If Not rs血库.EOF Then int血袋数 = Val(rs血库!数量 & "")
                bln输血 = True
            End If
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, 0) = Format(rsTmp!要求时间, "yyyy-MM-dd HH:mm")
                .TextMatrix(i, 1) = Format(rsTmp!执行时间, "yyyy-MM-dd HH:mm")
                If bln输血 Then
                    .TextMatrix(i, 2) = FormatEx(Val(rsTmp!本次数次 & "") * int血袋数, 0) & " 袋"
                Else
                    .TextMatrix(i, 2) = FormatEx(rsTmp!本次数次, 5) & " " & Nvl(rsTmp!计算单位)
                End If
                .TextMatrix(i, 3) = Nvl(rsTmp!执行摘要)
                .TextMatrix(i, 4) = Nvl(rsTmp!执行人)
                .TextMatrix(i, 5) = Format(rsTmp!登记时间, "yyyy-MM-dd HH:mm")
                .TextMatrix(i, 6) = Nvl(rsTmp!登记人)
                .TextMatrix(i, 7) = rsTmp!执行结果 & ""
                .TextMatrix(i, 8) = Nvl(rsTmp!核对人)
                .TextMatrix(i, 9) = Format(rsTmp!核对时间, "yyyy-MM-dd HH:mm")
		.TextMatrix(i, 10) = NVL(rsTmp!说明)
                .TextMatrix(i, 11) = IIf(1 = Val(rsTmp!来源 & ""), "移动端", "PC端")
                .Cell(flexcpData, i, 0) = Format(rsTmp!要求时间, "yyyy-MM-dd HH:mm:ss")
                .Cell(flexcpData, i, 1) = Format(rsTmp!执行时间, "yyyy-MM-dd HH:mm:ss")
                If .Cell(flexcpData, i, 0) = strPre Then .Row = i
                rsTmp.MoveNext
            Next
            rsTmp.MoveFirst
        End If
        Call .ShowCell(.Row, .Col)
        .Redraw = flexRDDirect
    End With
    LoadExecList = True
    With rptPati
        If .SelectedRows.Count > 0 Then
            If Not .SelectedRows(0).GroupRow Then
            
                If Not (.SelectedRows(0).Record(COL_诊疗类别).Value = "E" And .SelectedRows(0).Record(COL_操作类型).Value = "1" And Mid(gstr医嘱核对, 2, 1) = "1" Or _
                    .SelectedRows(0).Record(COL_诊疗类别).Value = "E" And .SelectedRows(0).Record(COL_操作类型).Value = "8" And Mid(gstr医嘱核对, 1, 1) = "1" Or _
                    .SelectedRows(0).Record(COL_诊疗类别).Value = "K" And Mid(gstr医嘱核对, 1, 1) = "1") Then
                    
                    vsExec.ColHidden(8) = True
                    vsExec.ColHidden(9) = True
                Else
                    vsExec.ColHidden(8) = False
                    vsExec.ColHidden(9) = False
                End If
                
            End If
        End If
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub vsExec_GotFocus()
    vsExec.BackColorSel = COLOR_FOCUS
End Sub

Private Sub vsExec_LostFocus()
    vsExec.BackColorSel = COLOR_LOST
End Sub

Private Function ShowBillList(objPopup As CommandBarPopup) As Boolean
'功能：显示当前执行医嘱可以打印的诊疗单据在菜单上
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objControl As CommandBarControl
        
    If mlng病人ID = 0 Then
        objPopup.CommandBar.Controls.DeleteAll
        ShowBillList = True: Exit Function
    End If
        
    With rptPati.SelectedRows(0)
        '主项才显示诊疗单据
        If Not .ParentRow.GroupRow Then
            objPopup.CommandBar.Controls.DeleteAll
            ShowBillList = True: Exit Function
        End If
        
        If .Record(col_医嘱ID).Value & "_" & .Record(col_发送号).Value = objPopup.Parameter Then
            ShowBillList = True: Exit Function
        Else
            objPopup.Parameter = .Record(col_医嘱ID).Value & "_" & .Record(col_发送号).Value
            objPopup.CommandBar.Controls.DeleteAll
        End If
    End With
        
    On Error GoTo errH
    
    With rptPati.SelectedRows(0)
        strSQL = "Select Distinct D.编号,D.名称,D.说明" & _
            " From 病人医嘱发送 A,病人医嘱记录 B,病历单据应用 C,病历文件列表 D" & _
            " Where A.发送号=[1] And A.NO=[2]" & _
            " And A.医嘱ID=B.ID And B.诊疗项目ID=C.诊疗项目ID" & _
            " And C.应用场合=[3] And C.病历文件ID=D.ID And D.种类=7" & _
            " Order by D.编号"
        If .Record(COL_数据转出).Value = 1 Then
            strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
            strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .Record(col_发送号).Value, .Record(col_单据号).Value, .Record(col_记录性质).Value)
    End With
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            With objPopup.CommandBar.Controls
                Set objControl = .Add(xtpControlButton, conMenu_Manage_RequestPrint * 10# + i, rsTmp!名称)
                If i <= 10 Then
                    objControl.Caption = objControl.Caption & "(&" & i - 1 & ")"
                ElseIf i <= 36 Then
                    objControl.Caption = objControl.Caption & "(&" & Chr(i - 11 + Asc("A")) & ")"
                End If
                objControl.Parameter = "ZLCISBILL" & Format(rsTmp!编号, "00000") & "-1" '对应的自定义报表编号
                'If i > 1 Then objControl.Enabled = False '一个项目只能设置一个诊疗单据
            End With
            rsTmp.MoveNext
        Next
        
        cbsMain.KeyBindings.Add 0, vbKeyF2, conMenu_Manage_RequestPrint * 10# + 1
    End If
    
    ShowBillList = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FuncBillPrint(objControl As CommandBarControl)
'功能：打印诊疗单据
    Dim strNO As String, int性质 As Integer
    Dim lng医嘱ID As Long, lng发送号 As Long
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    If rptPati.SelectedRows.Count = 0 Then Exit Sub
    If objControl.Parameter = "" Then '奇怪，直接按F2时，是一个空的Control
        Set objControl = cbsMain.FindControl(, conMenu_Manage_RequestPrint * 10# + 1, , True)
        If objControl Is Nothing Then Exit Sub
    End If
    If objControl.Parameter = "" Then Exit Sub
    
    With rptPati.SelectedRows(0)
        If ReportPrintSet(gcnOracle, glngSys, objControl.Parameter, Me) Then
            '是否是采集方式，如果是要特殊处理
            strSQL = "Select A.诊疗类别,B.操作类型 From 病人医嘱记录 A,诊疗项目目录 B Where A.诊疗项目ID=B.ID And A.ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .Record(col_医嘱ID).Value)
            If Not rsTmp.EOF Then
                '是采集方式
                If Nvl(rsTmp(0)) = "E" And Nvl(rsTmp(1)) = "6" Then
                    Print采集方式 .Record(col_单据号).Value, .Record(col_记录性质).Value, .Record(col_医嘱ID).Value, objControl.Parameter
                Else
                    Call ReportOpen(gcnOracle, glngSys, objControl.Parameter, Me, "NO=" & .Record(col_单据号).Value, "性质=" & .Record(col_记录性质).Value, "项目=Untitled", 2)
                End If
            Else
                '为了打印条码，报表增加了“项目”参数。By：赵彤宇
                Call ReportOpen(gcnOracle, glngSys, objControl.Parameter, Me, "NO=" & .Record(col_单据号).Value, "性质=" & .Record(col_记录性质).Value, "项目=Untitled", 2)
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Print采集方式(ByVal strNO As String, ByVal intAttribute As Integer, ByVal lngAdviceID As Long, ByVal strReport As String)
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo DataError
    Me.MousePointer = vbHourglass
    
    '同一标本再按仪器分别打印
'    strSQL = "Select 病人ID,标本,执行部门,NO," & _
'        " Trim(内容1||' '||内容2||' '||内容3||' '||内容4||' '||内容5) As 项目,仪器" & _
'        " From" & _
'        " (Select B.病人ID,B.标本部位 As 标本,F.名称 As 执行部门,S.仪器," & _
'        "  Max(Decode(Mod(Rownum,5),0,B.医嘱内容,'')) As 内容1," & _
'        "  Max(Decode(Mod(Rownum,5),1,B.医嘱内容,'')) As 内容2," & _
'        "  Max(Decode(Mod(Rownum,5),2,B.医嘱内容,'')) As 内容3," & _
'        "  Max(Decode(Mod(Rownum,5),3,B.医嘱内容,'')) As 内容4," & _
'        "  Max(Decode(Mod(Rownum,5),4,B.医嘱内容,'')) As 内容5," & _
'        "  Max(S.NO||','||S.记录性质) As NO" & _
'        "  From 病人医嘱记录 B,部门表 F," & _
'        "   (Select DISTINCT 医嘱ID,NO,记录性质,仪器,诊疗项目ID FROM " & _
'        "    (Select A.医嘱ID,A.NO,A.记录性质,B.诊疗项目ID,I.报告项目ID,MAX(Decode(M.名称,NULL,'手工',M.名称)) AS 仪器 " & _
'        "     From 病人医嘱发送 A,病人医嘱记录 B,病人医嘱记录 D,诊疗项目目录 C,检验报告项目 I,检验仪器项目 J,检验仪器 M," & _
'        "     (SELECT A.病人ID,B.发送时间,B.执行部门ID FROM 病人医嘱记录 A,病人医嘱发送 B" & _
'        "      WHERE A.ID=B.医嘱ID AND B.NO=[1] AND B.记录性质=[2]) N Where a.医嘱ID+0 = B.ID And B.诊疗项目ID = C.ID" & _
'        "      AND D.相关ID = B.ID AND D.诊疗项目ID=I.诊疗项目ID(+) AND I.报告项目ID=J.项目ID(+) AND J.仪器ID=M.ID(+)" & _
'        "      And C.类别='E' And Nvl(C.操作类型,'0')='6'" & _
'        "      And B.病人ID=N.病人ID And A.执行部门ID+0= N.执行部门ID And A.发送时间 BETWEEN to_Date(to_Char(N.发送时间,'YYYY-MM-DD'),'YYYY-MM-DD HH24:MI:SS') AND to_Date(to_Char(N.发送时间,'YYYY-MM-DD')||' 23:59:59','YYYY-MM-DD HH24:MI:SS')" & _
'        "      And Nvl(A.执行状态,0)=0 " & _
'        "     GROUP BY A.医嘱ID,A.NO,A.记录性质,B.诊疗项目ID,I.报告项目ID)" & _
'        "   ) S" & _
'        "  Where B.执行科室ID = F.ID And B.相关ID = S.医嘱ID" & _
'        "  Group By B.病人ID, B.标本部位,F.名称,S.仪器,S.诊疗项目ID)" & _
'        " Order By 病人ID"
    '同一标本只打一张
    strSQL = "Select 病人ID,标本,执行部门," & _
        " Trim(内容1||' '||内容2||' '||内容3||' '||内容4||' '||内容5) As 项目" & _
        " From" & _
        " (Select B.病人ID,B.标本部位 As 标本,F.名称 As 执行部门," & _
        "  Max(Decode(Mod(Rownum,5),0,B.医嘱内容,'')) As 内容1," & _
        "  Max(Decode(Mod(Rownum,5),1,B.医嘱内容,'')) As 内容2," & _
        "  Max(Decode(Mod(Rownum,5),2,B.医嘱内容,'')) As 内容3," & _
        "  Max(Decode(Mod(Rownum,5),3,B.医嘱内容,'')) As 内容4," & _
        "  Max(Decode(Mod(Rownum,5),4,B.医嘱内容,'')) As 内容5 " & _
        "  From 病人医嘱记录 B,部门表 F" & _
        "  Where B.执行科室ID = F.ID And B.相关ID = [1]" & _
        "  Group By B.病人ID, B.标本部位,F.名称)" & _
        " Order By 病人ID"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngAdviceID)
    
    If rsTmp.EOF Then
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    Call ReportOpen(gcnOracle, glngSys, strReport, Me, "NO=" & strNO, "性质=" & intAttribute, "项目=" & Nvl(rsTmp("项目")), 2)
    
    Me.MousePointer = vbDefault
    Exit Sub
DataError:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
        
    Me.MousePointer = vbDefault
End Sub

Private Function FuncShowReport(ByVal intOption As Integer) As Boolean
'功能：报告填写/查阅/打印
'参数：intOption=0-填写,1-查阅,2-打印,3-预览
    Dim rsTmp As ADODB.Recordset
    Dim lng医嘱ID As Long, strBill As String
    Dim strSQL As String
    
    On Error GoTo errH
    
    With rptPati.SelectedRows(0)
        If intOption = 0 And .Record(COL_数据转出).Value = 1 Then
            MsgBox "该病人的本次" & Decode(.Record(col_来源).Value, "住院", "住院", "就诊") & "数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
                "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
            Exit Function
        End If
    
        If (.Record(COL_诊疗类别).Value = "C" Or .Record(COL_诊疗类别).Value = "D") And .Record(col_相关ID).Value <> 0 Then
            lng医嘱ID = .Record(col_相关ID).Value '检验组合和多部位检查取相关ID
        Else
            lng医嘱ID = .Record(col_医嘱ID).Value
        End If
        
        Set mclsEPRReport = New zlRichEPR.cEPRDocument
        
        If intOption = 0 Then
            If .Record(col_报告ID).Value = 0 Then
                Call mclsEPRReport.InitEPRDoc(cprEM_新增, cprET_单病历编辑, .Record(col_文件ID).Value, _
                    Decode(.Record(col_来源).Value, "门诊", cprPF_门诊, "住院", cprPF_住院, "体检", cprPF_体检, "外来", cprPF_外来), _
                    mlng病人ID, IIf(.Record(col_挂号ID).Value <> 0, .Record(col_挂号ID).Value, .Record(col_主页ID).Value), _
                    .Record(col_婴儿).Value, mlngDept, lng医嘱ID)
            Else
                Call mclsEPRReport.InitEPRDoc(cprEM_修改, cprET_单病历编辑, .Record(col_报告ID).Value, _
                    Decode(.Record(col_来源).Value, "门诊", cprPF_门诊, "住院", cprPF_住院, "体检", cprPF_体检, "外来", cprPF_外来), _
                    mlng病人ID, IIf(.Record(col_挂号ID).Value <> 0, .Record(col_挂号ID).Value, .Record(col_主页ID).Value), _
                    .Record(col_婴儿).Value, mlngDept, lng医嘱ID)
            End If
            Call mclsEPRReport.ShowEPREditor(Me) '是非模态显示
        ElseIf intOption = 1 Then
            Call gobjRichEPR.ViewDocument(Me, .Record(col_报告ID).Value, True)
        ElseIf intOption = 2 Or intOption = 3 Then
            If .Record(col_报告项).Value = 1 Then
                '按编辑格式打印
                Call gobjRichEPR.PrintOrPreviewDoc(Me, cpr诊疗报告, .Record(col_报告ID).Value, IIf(intOption = 2, True, False))
            ElseIf .Record(col_报告项).Value = 2 Then
                '按报表格式打印
                strSQL = "Select 编号 From 病历文件列表 Where ID=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.Record(col_文件ID).Value))
                strBill = "ZLCISBILL" & Format(rsTmp!编号, "00000") & "-2"
                
                If intOption = 2 Then
                    If Not ReportPrintSet(gcnOracle, glngSys, strBill, Me) Then Exit Function
                End If
                Call ReportOpen(gcnOracle, glngSys, strBill, Me, "NO=" & .Record(col_单据号).Value, "性质=" & .Record(col_记录性质).Value, "医嘱ID=" & lng医嘱ID, IIf(intOption = 2, 2, 1))
            End If
        End If
    End With
    
    FuncShowReport = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FuncBatchPrint()
'功能：批量打印条码
    Dim strPatiSource As String
    
    If InStr(mstrPrivs, "门诊病人") > 0 And InStr(mstrPrivs, "住院病人") > 0 Then
        strPatiSource = "1,2,3"
    ElseIf InStr(mstrPrivs, "门诊病人") > 0 Then
        strPatiSource = "1"
    ElseIf InStr(mstrPrivs, "住院病人") > 0 Then
        strPatiSource = "2"
    Else
        strPatiSource = "3"
    End If
    frmLISBillPrint.ShowMe Me, strPatiSource, cboDept.ItemData(cboDept.ListIndex)
End Sub

Private Sub FuncExecPlanTime()
'功能：时间安排
    Dim lng医嘱ID As Long, lng发送号 As Long, lng执行科室ID As Long
    
    With rptPati.SelectedRows(0)
        If .Record(col_执行过程).Value > 1 Then
            MsgBox "该项目已由影像检查工作站处理，不允许再操作。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If .Record(COL_数据转出).Value = 1 Then
            MsgBox "该病人的本次" & Decode(.Record(col_来源).Value, "住院", "住院", "就诊") & "数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
                "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
            Exit Sub
        End If
        lng医嘱ID = .Record(col_医嘱ID).Value
        lng发送号 = .Record(col_发送号).Value
        If .Record(col_执行科室).Value <> cboDept.ItemData(cboDept.ListIndex) Then
            lng执行科室ID = cboDept.ItemData(cboDept.ListIndex)
        End If
    End With
    With frmTechnicPlanTime
        If .ShowMe(Me, mclsMipModule, lng医嘱ID, lng发送号, lng执行科室ID) Then Call LoadPatients
    End With
End Sub

Private Sub FuncExecPlan()
'功能：执行报到
    Dim lng医嘱ID As Long, lng发送号 As Long, lng执行科室ID As Long, lng病人ID As Long, lng卡类别ID As Long
    
    With rptPati.SelectedRows(0)
        If .Record(col_执行过程).Value > 1 Then
            MsgBox "该项目已由影像检查工作站处理，不允许再操作。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If .Record(COL_数据转出).Value = 1 Then
            MsgBox "该病人的本次" & Decode(.Record(col_来源).Value, "住院", "住院", "就诊") & "数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
                "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If .Record(col_执行过程).Value = 1 Then
            If MsgBox("该病人已经报到，需要调整吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
        lng医嘱ID = .Record(col_医嘱ID).Value
        lng发送号 = .Record(col_发送号).Value
        lng病人ID = .Record(col_病人Id).Value
        If .Record(col_执行科室).Value <> cboDept.ItemData(cboDept.ListIndex) Then
            lng执行科室ID = cboDept.ItemData(cboDept.ListIndex)
        End If
        '取就诊卡类型
        lng卡类别ID = Val(PatiIdentify.objIDKind.GetCurCard.接口序号)
    End With
    With frmTechnicPlan
        If .ShowMe(Me, lng医嘱ID, lng发送号, lng执行科室ID, lng卡类别ID, lng病人ID, mstrPrivs, mobjSquareCard) Then Call LoadPatients
    End With
End Sub

Private Sub FuncExecErase()
'功能：取消报到
    Dim lng医嘱ID As Long, lng发送号 As Long, strSQL As String
        
    With rptPati.SelectedRows(0)
        If .Record(col_执行过程).Value > 1 Then
            MsgBox "该项目已由影像检查工作站处理，不允许再操作。", vbInformation, gstrSysName
            Exit Sub
        End If
        If .Record(col_执行过程).Value = 0 Then
            MsgBox "该病人还没有报到。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If .Record(COL_数据转出).Value = 1 Then
            MsgBox "该病人的本次" & Decode(.Record(col_来源).Value, "住院", "住院", "就诊") & "数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
                "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If MsgBox("确实要取消报到吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        lng医嘱ID = .Record(col_医嘱ID).Value
        lng发送号 = .Record(col_发送号).Value
    End With
    
    strSQL = "ZL_病人医嘱执行_Plan(" & lng医嘱ID & "," & lng发送号 & ",0)"
    
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    err.Clear: On Error GoTo 0
    
    Call LoadPatients '要更新执行状态
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncExecRefuse()
'功能：拒绝执行
    Dim lng医嘱ID As Long, lng发送号 As Long, lng执行科室ID As Long
    Dim strSQL As String, blnTrans As Boolean
    Dim str结果 As String, strTextInput As String
    
    With rptPati.SelectedRows(0)
        '正在执行或已执行不允许拒绝
        If .Record(col_执行状态).Value = 2 Then
            MsgBox "该执行项目当前已经拒绝执行。", vbInformation, gstrSysName
            Exit Sub
        End If
        If .Record(col_执行状态).Value = 3 Then
            MsgBox "该执行项目当前正在执行，不能拒绝。", vbInformation, gstrSysName
            Exit Sub
        End If
        If .Record(col_执行状态).Value = 1 Then
            MsgBox "该执行项目当前已经执行，不能拒绝。", vbInformation, gstrSysName
            Exit Sub
        End If
        '已报到的病人不允许拒绝
        If .Record(col_执行过程).Value <> 0 Then
            MsgBox "该病人已经报到，不能拒绝。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If .Record(COL_数据转出).Value = 1 Then
            MsgBox "该病人的本次" & Decode(.Record(col_来源).Value, "住院", "住院", "就诊") & "数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
                "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        str结果 = zlCommFun.ShowMsgBox("拒绝执行", "请根据填写拒绝执行的原因。", _
            "确定(&O),?取消(&C)", Me, vbQuestion, , , , , , _
            "拒绝原因(&B)", 50, strTextInput, , True)
        If str结果 = "" Then Exit Sub
            
        lng医嘱ID = .Record(col_医嘱ID).Value
        lng发送号 = .Record(col_发送号).Value
        If .Record(col_执行科室).Value <> cboDept.ItemData(cboDept.ListIndex) Then
            lng执行科室ID = cboDept.ItemData(cboDept.ListIndex)
        End If
    End With
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    
        If lng执行科室ID <> 0 Then
            strSQL = "Zl_病人医嘱发送_科室变更(" & lng医嘱ID & "," & lng发送号 & "," & lng执行科室ID & ")"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        End If
        
        strSQL = "ZL_病人医嘱执行_拒绝执行(" & lng医嘱ID & "," & lng发送号 & ",NULL,NULL,NULL,'" & strTextInput & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    gcnOracle.CommitTrans: blnTrans = False
    err.Clear: On Error GoTo 0
     
    With rptPati.SelectedRows(0)
        Call ZLHIS_CIS_015(mclsMipModule, .Record(col_病人Id).Value, .Record(col_姓名).Value, .Record(col_标识号).Value, , 2, .Record(col_主页ID).Value, _
         .Record(col_病区id).Value, .Record(col_科室).Value, .Record(col_科室).Caption, , .Record(col_床号).Value, lng医嘱ID, .Record(col_期效).Value, _
         .Record(COL_诊疗类别).Value, .Record(COL_操作类型).Value, .Record(COL_诊疗项目ID).Value, .Record(col_内容).Value)
    End With
    
    Call LoadPatients '要更新执行状态
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncExecRestore()
'功能：取消拒绝执行
    Dim lng医嘱ID As Long, lng发送号 As Long
    Dim strSQL As String, i As Long
    
    With rptPati.SelectedRows(0)
        '正在执行或已执行不允许拒绝
        If .Record(col_执行状态).Value <> 2 Then
            MsgBox "该执行项目没有被拒绝执行。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If .Record(COL_数据转出).Value = 1 Then
            MsgBox "该病人的本次" & Decode(.Record(col_来源).Value, "住院", "住院", "就诊") & "数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
                "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If MsgBox("确实要取消拒绝执行该项目吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        lng医嘱ID = .Record(col_医嘱ID).Value
        lng发送号 = .Record(col_发送号).Value
    End With
    
    strSQL = "ZL_病人医嘱执行_取消拒绝(" & lng医嘱ID & "," & lng发送号 & ")"
    
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    err.Clear: On Error GoTo 0
    
    Call LoadPatients '要更新执行状态
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function CheckExecuteLog(ByVal lng医嘱ID As Long, ByVal lng发送号 As Long) As String
'功能：检查对应发送医嘱的执行情况记录
'返回：如果无执行情况记录或者次数未达到要求次数，则返回提示信息
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select A.发送数次,Sum(B.本次数次) as 已有数次" & _
        " From 病人医嘱发送 A,病人医嘱执行 B Where A.医嘱ID=B.医嘱ID And A.发送号=B.发送号 And A.医嘱ID=[1] And A.发送号=[2]" & _
        " Group by A.发送数次"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckExecuteLog", lng医嘱ID, lng发送号)
    If rsTmp.EOF Then
        CheckExecuteLog = "还没有记录执行情况"
    ElseIf Nvl(rsTmp!已有数次, 0) < Nvl(rsTmp!发送数次, 0) Then
        CheckExecuteLog = "已执行数次 " & Nvl(rsTmp!已有数次, 0) & " 没有达到要求的数次 " & Nvl(rsTmp!发送数次, 0)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FuncExecFinish()
'功能：确认执行完成
    Dim rsTmp As New ADODB.Recordset
    Dim lng医嘱ID As Long, lng发送号 As Long, lng病人ID As Long, strNos As String, str医嘱IDs As String
    Dim lng相关ID As Long, lng报告ID As Long, blnTmp As Boolean
    Dim strSQL As String, strTest As String
    Dim str报告Del As String, str执行 As String
    Dim str结果 As String, int结果 As Integer, strLabel As String
    Dim cnNew As ADODB.Connection, i As Long
    Dim strUserName As String, strOwner As String, blnTrans As Boolean
    Dim rptRowChild As ReportRow
    Dim blnIsAbnormal As Boolean
    Dim lng卡类别ID As Long
    Dim dateInput As Date
    Dim strSelect As String
    Dim strSelectInput As String
    Dim strTextInput As String
    Dim dat完成时间 As Date
    Dim datTmp As Date
    Dim lng新门诊服务 As Long
    
    Dim curMoney As Currency, str类别 As String, str类别名 As String
    
    '判断是否批量执行模式，是则在单独的流程中调用
    If rptPati.Columns(col_选择).Visible Then
        For i = 0 To rptPati.Rows.Count - 1
            If Not rptPati.Rows(i).GroupRow Then
                If rptPati.Rows(i).Record(col_选择).Checked Then Exit For
            End If
        Next
        If i <= rptPati.Rows.Count - 1 Then
            If MsgBox("要对当前选择的一个或多个项目执行完成吗？", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                Call FuncExecFinishBatch
            End If
            Exit Sub
        End If
    End If
    
    With rptPati.SelectedRows(0)
        lng医嘱ID = .Record(col_医嘱ID).Value
        lng发送号 = .Record(col_发送号).Value
        lng相关ID = .Record(col_相关ID).Value
        lng病人ID = .Record(col_病人Id).Value
        
        '可以不填写执行情况直接完成执行
        If .Record(col_综合状态).Value = 1 Then
            MsgBox "该执行项目当前已经执行完成。", vbInformation, gstrSysName
            Exit Sub
        End If
                
        '检查病人是否开始审核
        If Val(.Record(col_审核标志).Value & "") >= 1 And mbyt病人审核方式 = 1 Then
            MsgBox "该病人的费用正在审核阶段，不允许操作医嘱和费用。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If .Record(COL_诊疗类别).Value = "E" And .Record(COL_操作类型).Value = "1" And Mid(gstr医嘱核对, 2, 1) = "1" Or _
            .Record(COL_诊疗类别).Value = "E" And .Record(COL_操作类型).Value = "8" And Mid(gstr医嘱核对, 1, 1) = "1" Or _
            .Record(COL_诊疗类别).Value = "K" And Mid(gstr医嘱核对, 1, 1) = "1" Then
            '输血和皮试医嘱没核对不允许完成
            If .Record(COL_核对人).Value & "" = "" Then
                MsgBox "该项目是" & IIf(.Record(COL_操作类型).Value = "1", "皮试", "输血") & "医嘱，必须核对了才能完成。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        If .Record(COL_数据转出).Value = 1 Then
            MsgBox "该病人的本次" & Decode(.Record(col_来源).Value, "住院", "住院", "就诊") & "数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
                "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
            Exit Sub
        End If
        
	'病人收费判断，两类，新门诊病人，非新门诊病人 0-走以前逻辑，1-提示并禁止，2-成功验证通过
        If Val(.Record(COL_附加标志).Value) = 3 Then '新门诊病人
            lng新门诊服务 = NewOut收费(lng医嘱ID)
            If lng新门诊服务 = 1 Then Exit Sub
        Else
            lng新门诊服务 = 0
        End If
        
        If lng新门诊服务 = 0 Then

        blnIsAbnormal = False
        '是否允许完成未收费病人的项目
        If .Record(col_记录性质).Value = 1 Then
            '记帐划价,除非执行后不自动审核
            If Not ItemHaveCash(IIf(.Record(col_来源).Value = "住院", 2, 1), Not .ParentRow.GroupRow, _
                lng医嘱ID, .Record(col_相关ID).Value, .Record(col_发送号).Value, _
                .Record(COL_诊疗类别).Value, .Record(col_单据号).Value, .Record(col_记录性质).Value, .Record(col_门诊记帐).Value, _
                0, .Record(COL_数据转出).Value = 1, .Record(col_发送时间).Value, , , blnIsAbnormal) Then
                
                '判断单据是否异常
                If blnIsAbnormal Then MsgBox "该病人还存在异常费用，请检查。", vbInformation, gstrSysName: Exit Sub
                
                If Not mbln未收费完成 Then
                    If gbln执行前先结算 Then
                        '获得一组执行或者单独执行的医嘱字符串
                        If Not .ParentRow.GroupRow Or (.Childs.Count = 0 And .ParentRow.GroupRow) Then
                            str医嘱IDs = lng医嘱ID
                        Else
                            For Each rptRowChild In .Childs
                                str医嘱IDs = str医嘱IDs & IIf(str医嘱IDs = "" Or rptRowChild.Record(col_医嘱ID).Value & "" = "", "", ",") & rptRowChild.Record(col_医嘱ID).Value
                            Next
                        End If
                        '取就诊卡类型
                        lng卡类别ID = Val(PatiIdentify.objIDKind.GetCurCard.接口序号)
                    Else
                    
                        MsgBox "该病人还存在未收费的费用，请检查。", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
            End If
            '判断是否有销帐的费给出提示
            If Not mbln未收费完成 Then
                blnTmp = Check销帐费用(Not .ParentRow.GroupRow, lng医嘱ID, IIf(lng相关ID = 0, lng医嘱ID, lng相关ID), .Record(COL_诊疗类别).Value, IIf(.Record(col_来源).Value = "住院", 2, 1), .Record(col_单据号).Value)
                If blnTmp Then
                    If MsgBox("该病人存在销帐或退费的费用，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                End If
            End If
        End If
        End If
        
        
        '检查报告的填写情况：对应了病历单据，并且有报告的情况才需要填写
        If .Record(col_文件ID).Value <> 0 And .Record(col_报告项).Value <> 0 Then
            i = CheckEPRReport(IIf(.Record(COL_诊疗类别).Value = "C" Or (.Record(COL_诊疗类别).Value = "D" And lng相关ID <> 0), lng相关ID, lng医嘱ID), lng报告ID, True, .Record(col_综合状态).Value)
            If InStr(mstrPrivs, "直接执行完成") > 0 Then
                If i = 2 Then
                    If MsgBox("该项目的报告填写了内容但还没有完成，请先完成报告后再继续。" & _
                        vbCrLf & vbCrLf & "或者可以删除掉该份未完成的报告并继续，要删除吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                    str报告Del = "Zl_电子病历记录_Delete(" & lng报告ID & ")"
                End If
            Else
                If i = 0 Then
                    MsgBox "该项目的报告还没有填写，请先填写报告再继续。", vbInformation, gstrSysName
                    Exit Sub
                ElseIf i = 2 Then
                    MsgBox "该项目的报告填写了内容但还没有完成，请先完成报告后再继续。", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        End If
        
        If .Record(col_执行科室).Value <> cboDept.ItemData(cboDept.ListIndex) Then
            str执行 = "Zl_病人医嘱发送_科室变更(" & lng医嘱ID & "," & lng发送号 & "," & cboDept.ItemData(cboDept.ListIndex) & ")"
        End If
    End With
        
    On Error GoTo errH
    
    '判断是否皮试,再填写结果
    strSQL = "Select A.诊疗类别,A.皮试结果,B.操作类型,Nvl(B.标本部位,'阳性(+);阴性(-)') as 标本部位 From 病人医嘱记录 A,诊疗项目目录 B Where A.诊疗项目ID=B.ID And A.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
    If Not rsTmp.EOF Then
        '已经填写了皮试结果则不再填写
        If rsTmp!诊疗类别 = "E" And Nvl(rsTmp!操作类型) = "1" And IsNull(rsTmp!皮试结果) Then
            '先作身份验证
            If mbln皮试验证 Then
                Set cnNew = New ADODB.Connection
                strUserName = zlDatabase.UserIdentify(Me, "在填写皮试结果前，请您先输入用户名和密码进行身份验证。", glngSys, p医技工作站, "确认执行完成", cnNew)
                If strUserName = "" Then Exit Sub
            End If
            '阳性
            For i = 0 To UBound(Split(Split(rsTmp!标本部位 & "", ";")(0), ","))
                strSelect = strSelect & "," & Split(Split(rsTmp!标本部位 & "", ";")(0), ",")(i) & "|0"
            Next
            '阴性
            For i = 0 To UBound(Split(Split(rsTmp!标本部位 & "", ";")(1), ","))
                strSelect = strSelect & "," & Split(Split(rsTmp!标本部位 & "", ";")(1), ",")(i) & "|0|2"
            Next
            strSelect = Mid(strSelect, 2)
            
            '填写皮试结果
            str结果 = zlCommFun.ShowMsgBox("皮试结果", rptPati.SelectedRows(0).Record(col_内容).Value & "：^^请根据过敏试验结果选择相应的按钮操作。", _
            "确定(&O),?取消(&C)", Me, vbQuestion, "皮试时间", dateInput, "yyyy-MM-dd HH:mm", "皮试结果(&P):" & strSelect, strSelectInput, _
            "过敏反应(&F)", 50, strTextInput, , True)
            
            If str结果 = "" Then Exit Sub
            If strSelectInput = "" Then Exit Sub
            Call GetTestLabel(rsTmp!标本部位, strSelectInput, strLabel, int结果)
            strTest = "ZL_病人医嘱记录_皮试(" & lng医嘱ID & ",'" & strLabel & "'," & int结果 & _
                        ",'',to_date('" & dateInput & "','YYYY-MM-DD HH24:MI:SS')" & ",'" & strTextInput & "')"
        End If
    Else
        MsgBox "对应的医嘱记录不存在，无法完成操作。", vbInformation, gstrSysName
        Exit Sub
    End If

    '----
    With rptPati.SelectedRows(0)
        '只检查记帐费用
        If .Record(col_记录性质).Value = 2 Then
            curMoney = GetAdviceMoney(IIf(lng相关ID = 0, lng医嘱ID, lng相关ID), lng医嘱ID, lng发送号, str类别, str类别名, Not .ParentRow.GroupRow, _
                   IIf(.Record(col_来源).Value = "住院" And .Record(col_门诊记帐).Value = 0, 2, 1))
            If curMoney > 0 Then
                '住院出院病人费用控制
                If .Record(col_来源).Value = "住院" Then
                    If Not PatiCanBilling(.Record(col_病人Id).Value, .Record(col_主页ID).Value, GetInsidePrivs(p医嘱附费管理), p医嘱附费管理) Then Exit Sub
                End If
                '记帐报警
                If InitObjPublicExpense Then
                    If gobjPublicExpense.zlBillingWarn.zlBillingVerfyWarnCheck(Me, p医技工作站, "", .Record(col_单据号).Value, GetInsidePrivs(p医嘱附费管理), Val(.Record(col_病区id).Value)) = False Then Exit Sub
                End If
                    
                '门诊一卡通消费身份验证,只检查门诊记帐费用
                If gdbl预存款消费验卡 <> 0 And _
                    (.Record(col_来源).Value <> "住院" Or .Record(col_来源).Value = "住院" And .Record(col_门诊记帐).Value = 1) Then
                    If Not zlDatabase.PatiIdentify(Me, glngSys, .Record(col_病人Id).Value, curMoney, , , , IIf(-1 * gdbl预存款消费验卡 >= Val(curMoney), False, True), , , (gdbl预存款消费验卡 <> 0), (2 = gdbl预存款消费验卡)) Then Exit Sub
                End If
            End If
        End If
    
        '严格要求记录执行的情况
        If mblnExeLog Then
            strSQL = CheckExecuteLog(lng医嘱ID, lng发送号)
            If strSQL <> "" Then
                MsgBox "该执行项目" & strSQL & "，不能完成执行。", vbInformation, gstrSysName
                Exit Sub
            Else
                If strTest = "" And str报告Del = "" Then
                    If MsgBox("确认该执行项目执行完成吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                End If
            End If
        Else
            If strTest = "" And str报告Del = "" Then
                If MsgBox("确认该执行项目执行完成吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
        End If
        
        dat完成时间 = zlDatabase.Currentdate
        datTmp = dat完成时间
        blnTmp = frmSelectTime.ShowMe(Me, dat完成时间, datTmp, Me, 1)
        If Not blnTmp Then
            Exit Sub
        End If
        
        
        '门诊一卡通,项目执行前必须先收费或先记帐审核,不传单据号，根据医嘱ID读取所有未收费单据或未审核的记帐单
        If gbln执行前先结算 And str医嘱IDs <> "" Then
            If mobjSquareCard.zlSquareAffirm(Me, p医技工作站, mstrPrivs, lng病人ID, lng卡类别ID, False, , , str医嘱IDs) = False Then
                Exit Sub
            End If
        End If
        
        strSQL = "ZL_病人医嘱执行_Finish(" & lng医嘱ID & "," & lng发送号 & "," & _
            "Null," & IIf(Not .ParentRow.GroupRow, 1, 0) & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlngDept & ",0,to_date('" & dat完成时间 & "','YYYY-MM-DD HH24:MI:SS'))"
   
    End With
    
    If strTest <> "" And Not cnNew Is Nothing Then
        strOwner = SystemOwner(glngSys)
        
        On Error GoTo errNew
        cnNew.BeginTrans: blnTrans = True
        
        If str执行 <> "" Then
            Call SQLTest(App.ProductName, Me.Caption, str执行)
            cnNew.Execute strOwner & "." & str执行, , adCmdStoredProc
            Call SQLTest
        End If
        
        Call SQLTest(App.ProductName, Me.Caption, strTest)
        cnNew.Execute strOwner & "." & strTest, , adCmdStoredProc
        Call SQLTest
        
        If str报告Del <> "" Then
            Call SQLTest(App.ProductName, Me.Caption, str报告Del)
            cnNew.Execute strOwner & "." & str报告Del, , adCmdStoredProc
            Call SQLTest
        End If
        
        Call SQLTest(App.ProductName, Me.Caption, strSQL)
        cnNew.Execute strOwner & "." & strSQL, , adCmdStoredProc
        Call SQLTest
        
        cnNew.CommitTrans: blnTrans = False
        cnNew.Close: Set cnNew = Nothing
        On Error GoTo errH
    Else
        gcnOracle.BeginTrans: blnTrans = True
            If str执行 <> "" Then
                Call zlDatabase.ExecuteProcedure(str执行, Me.Caption)
            End If
            If strTest <> "" Then
                Call zlDatabase.ExecuteProcedure(strTest, Me.Caption)
            End If
            If str报告Del <> "" Then
                Call zlDatabase.ExecuteProcedure(str报告Del, Me.Caption)
            End If
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        gcnOracle.CommitTrans: blnTrans = False
    End If
    
    Call LoadPatients '要更新执行状态
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
errNew:
    If blnTrans Then cnNew.RollbackTrans
    MsgBox err.Number & vbCrLf & vbCrLf & err.Description, vbInformation, gstrSysName
    cnNew.Close: Set cnNew = Nothing
End Sub

Private Sub FuncExecFinishBatch()
'功能：对一个病人选择的多个项目确认执行完成
    Dim rsTmp As New ADODB.Recordset
    Dim arrSQL As Variant, strSQL As String, i As Long
    Dim str结果 As String, int结果 As Integer, strLabel As String, strNos As String, str医嘱IDs As String
    Dim lng报告ID As Long, intThing As Integer
    Dim blnTmp As Boolean, blnTest As Boolean, blnTrans As Boolean
    
    Dim cnNew As ADODB.Connection
    Dim strUserName As String, strOwner As String
    
    Dim rsPati As ADODB.Recordset
    Dim curMoney As Currency, str类别 As String, str类别名 As String
    Dim strPatiIDs As String, blnIsMany As Boolean         'blnIsMany是否勾选了多个病人的单据
    Dim blnIsAbnormal As Boolean
    Dim lng卡类别ID  As Long
    Dim dateInput As Date
    Dim strMsgAduit As String
    Dim strMsg As String
    Dim strSelect As String, j As Long
    Dim strSelectInput As String
    Dim strTextInput As String
    Dim dat完成时间 As Date
    Dim datTmp As Date
    Dim lng新门诊服务 As Long
    
    On Error GoTo errH
    
    Set rsPati = New ADODB.Recordset
    rsPati.Fields.Append "来源", adVarChar, 10
    rsPati.Fields.Append "记录性质", adBigInt
    rsPati.Fields.Append "门诊记帐", adBigInt
    rsPati.Fields.Append "病人ID", adBigInt
    rsPati.Fields.Append "主页ID", adBigInt
    rsPati.Fields.Append "病区ID", adBigInt
    rsPati.Fields.Append "组ID", adVarChar, 2000
    rsPati.Fields.Append "医嘱ID", adVarChar, 2000
    rsPati.Fields.Append "发送号", adVarChar, 2000
    rsPati.Fields.Append "NO", adVarChar, 4000
    
    rsPati.CursorLocation = adUseClient
    rsPati.LockType = adLockOptimistic
    rsPati.CursorType = adOpenStatic
    rsPati.Open
    
    arrSQL = Array()
    
    dat完成时间 = zlDatabase.Currentdate
    datTmp = dat完成时间
    blnTmp = frmSelectTime.ShowMe(Me, dat完成时间, datTmp, Me, 1)
    If Not blnTmp Then
        Exit Sub
    End If
    
    '取就诊卡类型
    If gbln执行前先结算 Then
        lng卡类别ID = Val(PatiIdentify.objIDKind.GetCurCard.接口序号)
    End If
    
    For i = 0 To rptPati.Rows.Count - 1
        If Not rptPati.Rows(i).GroupRow Then
            If rptPati.Rows(i).Record(col_选择).Checked Then
                With rptPati.Rows(i).Record
                    '可以不填写执行情况直接完成执行
                    If .Item(col_综合状态).Value = 1 Then
                        MsgBox "病人""" & .Item(col_姓名).Value & """的项目""" & .Item(col_内容).Value & """当前已经执行完成。", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    
                    If .Item(COL_诊疗类别).Value = "E" And .Item(COL_操作类型).Value = "1" And Mid(gstr医嘱核对, 2, 1) = "1" Or _
                        .Item(COL_诊疗类别).Value = "E" And .Item(COL_操作类型).Value = "8" And Mid(gstr医嘱核对, 1, 1) = "1" Or _
                        .Item(COL_诊疗类别).Value = "K" And Mid(gstr医嘱核对, 1, 1) = "1" Then
                        '输血和皮试医嘱没核对不允许完成
                        If .Item(COL_核对人).Value & "" = "" Then
                            strMsgAduit = strMsgAduit & "," & .Item(col_单据号).Value
                        End If
                    End If
                    
                    '检查病人是否开始审核
                    If Val(.Item(col_审核标志).Value & "") >= 1 And mbyt病人审核方式 = 1 Then
                        strMsg = strMsg & "," & .Item(col_单据号).Value
                    Else
                        '严格要求记录执行的情况
                        If mblnExeLog Then
                            strSQL = CheckExecuteLog(.Item(col_医嘱ID).Value, .Item(col_发送号).Value)
                            If strSQL <> "" Then
                                MsgBox "病人""" & .Item(col_姓名).Value & """的项目""" & .Item(col_内容).Value & """" & strSQL & "，不能完成执行。", vbInformation, gstrSysName
                                Exit Sub
                            End If
                        End If
                            
                        If .Item(COL_数据转出).Value = 1 Then
                            MsgBox "病人""" & .Item(col_姓名).Value & """的项目""" & .Item(col_内容).Value & """的数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
                                "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
                            Exit Sub
                        End If
                        
			'病人收费判断，两类，新门诊病人，非新门诊病人 0-走以前逻辑，1-提示并禁止，2-成功验证通过
                        If Val(.Item(COL_附加标志).Value) = 3 Then '新门诊病人
                            lng新门诊服务 = NewOut收费(Val(.Item(col_医嘱ID).Value))
                            If lng新门诊服务 = 1 Then Exit Sub
                        Else
                            lng新门诊服务 = 0
                        End If
        
                        If lng新门诊服务 = 0 Then

                        If .Item(col_记录性质).Value = 1 Then
                            '不管记帐划价,除非执行后自动审核
                            If Not ItemHaveCash(IIf(.Item(col_来源).Value = "住院", 2, 1), Not rptPati.Rows(i).ParentRow.GroupRow, _
                                .Item(col_医嘱ID).Value, .Item(col_相关ID).Value, .Item(col_发送号).Value, .Item(COL_诊疗类别).Value, _
                                .Item(col_单据号).Value, .Item(col_记录性质).Value, .Item(col_门诊记帐).Value, 0, .Item(COL_数据转出).Value = 1, .Item(col_发送时间).Value, , , blnIsAbnormal) Then
                                
                                '判断单据是否异常
                                If blnIsAbnormal Then MsgBox "该病人还存在异常费用，请检查。", vbInformation, gstrSysName: Exit Sub
                                
                                '是否允许完成未收费病人的项目
                                If Not mbln未收费完成 Then
                                    '门诊一卡通,项目执行前必须先收费或先记帐审核
                                    If gbln执行前先结算 Then
                                        '获取病人的所有选中医嘱ID字符串(这里判断病人ID是为了一个病人多张单据只调用一次接口）
                                        If InStr("," & strPatiIDs & ",", "," & .Item(col_病人Id).Value & ",") = 0 Then
                                            str医嘱IDs = GetSelectAdviceIDs(Val(.Item(col_病人Id).Value), blnIsMany)
                                            '如果是多个病人批量结算，给予提示，不允许结算
                                            If blnIsMany Then
                                                MsgBox "由于涉及费用结算，不允许多个病人批量完成。", vbInformation, Me.Caption
                                                Exit Sub
                                            End If
                                            
                                            strPatiIDs = strPatiIDs & "," & .Item(col_病人Id).Value
                                        End If
                                    Else
                                        MsgBox "病人""" & .Item(col_姓名).Value & """的项目""" & .Item(col_内容).Value & """还存在未收费的费用，请检查。", vbInformation, gstrSysName
                                        Exit Sub
                                    End If
                                End If
                            End If
                            '判断是否有销帐的费给出提示
                            If Not mbln未收费完成 Then
                                blnTmp = Check销帐费用(Not rptPati.Rows(i).ParentRow.GroupRow, Val(.Item(col_医嘱ID).Value), _
                                            IIf(Val(.Item(col_相关ID).Value) = 0, Val(.Item(col_医嘱ID).Value), Val(.Item(col_相关ID).Value)), _
                                            .Item(COL_诊疗类别).Value, IIf(.Item(col_来源).Value = "住院", 2, 1), .Item(col_单据号).Value)
                                If blnTmp Then
                                    If MsgBox("病人""" & .Item(col_姓名).Value & "【" & .Item(col_单据号).Value & "】" & """的项目""" & .Item(col_内容).Value & """存在销帐或退费的费用，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                                End If
                            End If
                        End If
                              End If                        

                        '检查报告的填写情况：对应了病历单据，并且有报告的情况才需要填写
                        If .Item(col_文件ID).Value <> 0 And .Item(col_报告项).Value <> 0 Then
                            intThing = CheckEPRReport(IIf(.Item(COL_诊疗类别).Value = "C" Or (.Item(COL_诊疗类别).Value = "D" And Val(.Item(col_相关ID).Value) <> 0), .Item(col_相关ID).Value, .Item(col_医嘱ID).Value), lng报告ID, True, .Item(col_综合状态).Value)
                            If InStr(mstrPrivs, "直接执行完成") > 0 Then
                                If intThing = 2 Then
                                    If MsgBox("病人""" & .Item(col_姓名).Value & """的项目""" & .Item(col_内容).Value & """的报告填写了内容但还没有完成，请先完成报告后再继续。" & _
                                        vbCrLf & vbCrLf & "或者可以删除掉该份未完成的报告并继续，要删除吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                    arrSQL(UBound(arrSQL)) = "Zl_电子病历记录_Delete(" & lng报告ID & ")"
                                End If
                            Else
                                If intThing = 0 Then
                                    MsgBox "病人""" & .Item(col_姓名).Value & """的项目""" & .Item(col_内容).Value & """的报告还没有填写，请先填写报告再继续。", vbInformation, gstrSysName
                                    Exit Sub
                                ElseIf intThing = 2 Then
                                    MsgBox "病人""" & .Item(col_姓名).Value & """的项目""" & .Item(col_内容).Value & """的报告填写了内容但还没有完成，请先完成报告后再继续。", vbInformation, gstrSysName
                                    Exit Sub
                                End If
                            End If
                        End If
                        
                        '判断是否皮试,再填写结果
                        strSQL = "Select A.诊疗类别,A.皮试结果,B.操作类型,Nvl(B.标本部位,'阳性(+);阴性(-)') as 标本部位 From 病人医嘱记录 A,诊疗项目目录 B Where A.诊疗项目ID=B.ID And A.ID=[1]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .Item(col_医嘱ID).Value)
                        If Not rsTmp.EOF Then
                            '已经填写了皮试结果则不再填写
                            If rsTmp!诊疗类别 = "E" And Nvl(rsTmp!操作类型) = "1" And IsNull(rsTmp!皮试结果) Then
                                '先作身份验证
                                If mbln皮试验证 And cnNew Is Nothing Then
                                    Set cnNew = New ADODB.Connection
                                    strUserName = zlDatabase.UserIdentify(Me, "在填写皮试结果前，请您先输入用户名和密码进行身份验证。", glngSys, p医技工作站, "确认执行完成", cnNew)
                                    If strUserName = "" Then Exit Sub
                                End If
                                strSelect = "": strSelectInput = ""
                                '阳性
                                For j = 0 To UBound(Split(Split(rsTmp!标本部位 & "", ";")(0), ","))
                                    strSelect = strSelect & "," & Split(Split(rsTmp!标本部位 & "", ";")(0), ",")(j) & "|0"
                                Next
                                '阴性
                                For j = 0 To UBound(Split(Split(rsTmp!标本部位 & "", ";")(1), ","))
                                    strSelect = strSelect & "," & Split(Split(rsTmp!标本部位 & "", ";")(1), ",")(j) & "|0|2"
                                Next
                                strSelect = Mid(strSelect, 2)
                                '填写皮试结果
                                str结果 = zlCommFun.ShowMsgBox("皮试结果", .Item(col_内容).Value & "：^^请根据过敏试验结果选择相应的按钮操作。", _
                                        "确定(&O),?取消(&C)", Me, vbQuestion, "皮试时间", dateInput, "yyyy-MM-dd HH:mm", "皮试结果(&P):" & strSelect, strSelectInput, _
                                        "过敏反应(&F)", 50, strTextInput, , True)
                                If str结果 = "" Then Exit Sub
                                
                                blnTest = True
                                If strSelectInput = "" Then Exit Sub
                                Call GetTestLabel(rsTmp!标本部位, strSelectInput, strLabel, int结果)
                                
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_皮试(" & .Item(col_医嘱ID).Value & ",'" & strLabel & "'," & int结果 & _
                                                    ",'',to_date('" & dateInput & "','YYYY-MM-DD HH24:MI:SS')" & ",'" & strTextInput & "')"
                            End If
                        Else
                            MsgBox "病人""" & .Item(col_姓名).Value & """的项目""" & .Item(col_内容).Value & """对应的医嘱记录不存在，无法完成操作。", vbInformation, gstrSysName
                            Exit Sub
                        End If
                        
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "ZL_病人医嘱执行_Finish(" & .Item(col_医嘱ID).Value & "," & _
                            .Item(col_发送号).Value & ",Null," & IIf(Not rptPati.Rows(i).ParentRow.GroupRow, 1, 0) & "," & _
                            "'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlngDept & ",0,to_date('" & dat完成时间 & "','YYYY-MM-DD HH24:MI:SS'))"
                           
                        '收集不同的病人信息
                        rsPati.Filter = "来源='" & .Item(col_来源).Value & "' And 门诊记帐=" & .Item(col_门诊记帐).Value & _
                            " And 记录性质=" & .Item(col_记录性质).Value & " And 病人ID=" & .Item(col_病人Id).Value & _
                            " And 主页ID=" & .Item(col_主页ID).Value & " And 病区ID=" & .Item(col_病区id).Value
                        If rsPati.EOF Then
                            rsPati.AddNew
                            rsPati!来源 = CStr(.Item(col_来源).Value)
                            rsPati!记录性质 = Val(.Item(col_记录性质).Value)
                            rsPati!门诊记帐 = Val(.Item(col_门诊记帐).Value)
                            rsPati!病人ID = Val(.Item(col_病人Id).Value)
                            rsPati!主页ID = Val(.Item(col_主页ID).Value)
                            rsPati!病区ID = Val(.Item(col_病区id).Value)
                            rsPati.Update
                        End If
                        rsPati!组ID = Nvl(rsPati!组ID) & "," & IIf(.Item(col_相关ID).Value = 0, .Item(col_医嘱ID).Value, .Item(col_相关ID).Value)
                        rsPati!医嘱ID = Nvl(rsPati!医嘱ID) & "," & .Item(col_医嘱ID).Value
                        rsPati!发送号 = Nvl(rsPati!发送号) & "," & .Item(col_发送号).Value
                        rsPati!NO = Nvl(rsPati!NO) & "," & .Item(col_单据号).Value
                        rsPati.Update
                    End If
                        
                    
                End With
            End If
        End If
    Next
    
    If strMsgAduit <> "" Or strMsg <> "" Then
        If strMsg <> "" And strMsgAduit <> "" Then
            strMsg = "以下单据号的病人费用正在审核阶段，不允许操作医嘱和费用：" & vbCrLf & Mid(strMsg, 2) & "。" & vbCrLf & _
                    "以下单据号由于是输血或是皮试项目，必须核对后再执行完成：" & vbCrLf & Mid(strMsgAduit, 2) & "。"
        ElseIf strMsgAduit <> "" Then
            strMsg = "以下单据号由于是输血或是皮试项目，必须核对后再执行完成：" & vbCrLf & Mid(strMsgAduit, 2) & "。"
        Else
            strMsg = "以下单据号的病人费用正在审核阶段，不允许操作医嘱和费用：" & vbCrLf & Mid(strMsg, 2) & "。"
        End If
        MsgBox strMsg, vbInformation, gstrSysName
    End If
    
    If UBound(arrSQL) = -1 Then Exit Sub
    
    '多个病人的费用检查和报警
    rsPati.Filter = "记录性质 = 2"
    Do While Not rsPati.EOF
        curMoney = GetAdviceMoney(Mid(rsPati!组ID, 2), Mid(rsPati!医嘱ID, 2), Mid(rsPati!发送号, 2), str类别, str类别名, False, _
                IIf(rsPati!来源 = "住院" And rsPati!门诊记帐 = 0, 2, 1))
        If curMoney > 0 Then
            '住院出院病人费用控制
            If rsPati!来源 = "住院" Then
                If Not PatiCanBilling(rsPati!病人ID, rsPati!主页ID, GetInsidePrivs(p医嘱附费管理), p医嘱附费管理) Then Exit Sub
            End If
            '记帐报警
            If InitObjPublicExpense Then
                If gobjPublicExpense.zlBillingWarn.zlBillingVerfyWarnCheck(Me, p医技工作站, "", Mid(rsPati!NO & "", 2), GetInsidePrivs(p医嘱附费管理), Val(rsPati!病区ID & "")) = False Then Exit Sub
            End If
            '门诊一卡通消费身份验证
            If gdbl预存款消费验卡 <> 0 And (rsPati!来源 <> "住院" Or rsPati!来源 = "住院" And rsPati!门诊记帐 = 1) Then
                If Not zlDatabase.PatiIdentify(Me, glngSys, rsPati!病人ID, curMoney, , , , IIf(-1 * gdbl预存款消费验卡 >= Val(curMoney), False, True), , , (gdbl预存款消费验卡 <> 0), (2 = gdbl预存款消费验卡)) Then Exit Sub
            End If
        End If
        rsPati.MoveNext
    Loop
    
    If gbln执行前先结算 And str医嘱IDs <> "" Then
        If mobjSquareCard.zlSquareAffirm(Me, p医技工作站, mstrPrivs, Val(Mid(strPatiIDs, 2)), lng卡类别ID, False, , , str医嘱IDs) = False Then
            Exit Sub
        End If
    End If

    '提交SQL执行
    If blnTest And Not cnNew Is Nothing Then
        strOwner = SystemOwner(glngSys)
        
        On Error GoTo errNew
        cnNew.BeginTrans: blnTrans = True
            
        For i = 0 To UBound(arrSQL)
            Call SQLTest(App.ProductName, Me.Caption, arrSQL(i))
            cnNew.Execute strOwner & "." & arrSQL(i), , adCmdStoredProc
            Call SQLTest
        Next
        
        cnNew.CommitTrans: blnTrans = False
        cnNew.Close: Set cnNew = Nothing
        On Error GoTo errH
    Else
        gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
        Next
        gcnOracle.CommitTrans: blnTrans = False
    End If
    Call LoadPatients '要更新执行状态
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Exit Sub
errNew:
    cnNew.RollbackTrans
    MsgBox err.Number & vbCrLf & vbCrLf & err.Description, vbInformation, gstrSysName
    cnNew.Close: Set cnNew = Nothing
End Sub

Private Function GetSelectAdviceIDs(ByVal lngPatiID As Long, ByRef blnIsMany As Boolean) As String
'功能：根据病人获得批量执行所有医嘱ID字符串
    Dim i As Long, strSelectAdvices As String
    Dim rptRowChild As ReportRow
    
    For i = 0 To rptPati.Rows.Count - 1
        If Not rptPati.Rows(i).GroupRow Then
            If rptPati.Rows(i).Record(col_选择).Checked Then
                With rptPati.Rows(i)
                    '判断是否选中多个病人的单据
                    If lngPatiID <> Val(.Record(col_病人Id).Value) Then blnIsMany = True: Exit Function
                    
                    If .Record(col_记录性质).Value = 1 Or .Record(col_记录性质).Value = 2 And (.Record(col_来源).Value <> "住院" _
                            Or .Record(col_来源).Value = "住院" And .Record(col_门诊记帐).Value = 1) Then
                        If Val(.Record(col_病人Id).Value) = lngPatiID Then
                            '获得一组执行或者单独执行的医嘱字符串
                            If Not .ParentRow.GroupRow Or (.Childs.Count = 0 And .ParentRow.GroupRow) Then
                                strSelectAdvices = strSelectAdvices & IIf(.Record(col_医嘱ID).Value & "" = "", "", ",") & .Record(col_医嘱ID).Value
                            Else
                                For Each rptRowChild In .Childs
                                    strSelectAdvices = strSelectAdvices & IIf(rptRowChild.Record(col_医嘱ID).Value & "" = "", "", ",") & rptRowChild.Record(col_医嘱ID).Value
                                Next
                            End If
                        End If
                    End If
                End With
            End If
        End If
    Next
    GetSelectAdviceIDs = Mid(strSelectAdvices, 2)
End Function

Private Sub FuncExecCancel()
'功能：取消执行完成
    Dim lng组ID As Long, lng医嘱ID As Long, lng发送号 As Long
    Dim str诊疗类别 As String, strSQL As String, byt来源 As Byte
    Dim strOwner As String, strUserName As String
    Dim cnNew As ADODB.Connection, rsTmp As ADODB.Recordset
    Dim i As Long
    
    '判断是否批量执行模式，是则在单独的流程中调用
    If rptPati.Columns(col_选择).Visible Then
        For i = 0 To rptPati.Rows.Count - 1
            If Not rptPati.Rows(i).GroupRow Then
                If rptPati.Rows(i).Record(col_选择).Checked Then Exit For
            End If
        Next
        If i <= rptPati.Rows.Count - 1 Then
            If MsgBox("要对当前选择的一个或多个项目取消执行完成吗？", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                Call FuncExecCancelBatch
            End If
            Exit Sub
        End If
    End If

    With rptPati.SelectedRows(0)
        
        '必须是已执行才可以取消
        If .Record(col_综合状态).Value <> 1 Then
            MsgBox "该执行项目当前不处于已执行状态，不能取消执行。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '检查病人是否开始审核
        If Val(.Record(col_审核标志).Value & "") >= 1 And mbyt病人审核方式 = 1 Then
            MsgBox "该病人的费用正在审核阶段，不允许操作医嘱和费用。", vbInformation, gstrSysName
            Exit Sub
        End If
                
        If .Record(COL_数据转出).Value = 1 Then
            MsgBox "该病人的本次" & Decode(.Record(col_来源).Value, "住院", "住院", "就诊") & "数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
                "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        lng医嘱ID = .Record(col_医嘱ID).Value
        lng发送号 = .Record(col_发送号).Value
        str诊疗类别 = .Record(COL_诊疗类别).Value
        lng组ID = IIf(.Record(col_相关ID).Value = 0, .Record(col_医嘱ID).Value, .Record(col_相关ID).Value)
        
        
        If Val(.Record(col_记录性质).Value) <> 1 Then
            If .Record(col_来源).Value = "住院" And Val(.Record(col_门诊记帐).Value) = 0 Then
                byt来源 = 2
            Else
                byt来源 = 1
            End If
            '费用结帐判断
            If Not ItemCanCancel(lng医嘱ID, lng发送号, lng组ID, str诊疗类别, Not .ParentRow.GroupRow, .Record(COL_数据转出).Value = 1, byt来源) Then Exit Sub
        End If
    End With
            
    If MsgBox("确实要将该执行项目取消执行完成吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    '判断是否皮试,再填写结果
    strSQL = "Select A.诊疗类别,A.皮试结果,B.操作类型,Nvl(B.标本部位,'阳性(+);阴性(-)') as 标本部位 From 病人医嘱记录 A,诊疗项目目录 B Where A.诊疗项目ID=B.ID And A.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
    If Not rsTmp.EOF Then
        '已经填写了皮试结果则不再填写
        If rsTmp!诊疗类别 = "E" And Nvl(rsTmp!操作类型) = "1" And Not IsNull(rsTmp!皮试结果) Then
            '先作身份验证
            If mbln皮试验证 Then
                Set cnNew = New ADODB.Connection
                strUserName = zlDatabase.UserIdentify(Me, "在取消完成皮试医嘱前，请您先输入用户名和密码进行身份验证。", glngSys, p住院医嘱发送, "皮试医嘱结果", cnNew)
                If strUserName = "" Then Exit Sub
            End If
            strSQL = "ZL_病人医嘱执行_Cancel(" & lng医嘱ID & "," & lng发送号 & ",1," & _
                IIf(Not rptPati.SelectedRows(0).ParentRow.GroupRow, 1, 0) & "," & mlngDept & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
        Else
            strSQL = "ZL_病人医嘱执行_Cancel(" & lng医嘱ID & "," & lng发送号 & ",Null," & _
                IIf(Not rptPati.SelectedRows(0).ParentRow.GroupRow, 1, 0) & "," & mlngDept & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
        End If
 
    End If
    If Not cnNew Is Nothing Then
        strOwner = SystemOwner(glngSys)

        On Error GoTo errNew
        Call SQLTest(App.ProductName, Me.Caption, strSQL)
        cnNew.Execute strOwner & "." & strSQL, , adCmdStoredProc
        Call SQLTest
        cnNew.Close: Set cnNew = Nothing
        On Error GoTo errH
    Else
        On Error GoTo errH
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        On Error GoTo 0
    End If
    
    Call LoadPatients '要更新执行状态
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
errNew:
    MsgBox err.Number & vbCrLf & vbCrLf & err.Description, vbInformation, gstrSysName
    cnNew.Close: Set cnNew = Nothing
End Sub

Private Function FuncThingNew(Optional ByVal blnRefresh As Boolean = True) As Boolean
    Dim lng科室ID As Long, lng执行科室ID As Long
    Dim lng医嘱ID As Long, lng发送号 As Long
    
    With rptPati.SelectedRows(0)
        If .Record(col_综合状态).Value = 1 Then '子项和独项同执行状态
            MsgBox "该执行项目已经执行完成，不能再继续操作。", vbInformation, gstrSysName
            Exit Function
        End If
        
        If .Record(COL_数据转出).Value = 1 Then
            MsgBox "该病人的本次" & Decode(.Record(col_来源).Value, "住院", "住院", "就诊") & "数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
                "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
            Exit Function
        End If
        
        lng科室ID = cboDept.ItemData(cboDept.ListIndex)
        lng医嘱ID = .Record(col_医嘱ID).Value
        lng发送号 = .Record(col_发送号).Value
        If .Record(col_执行科室).Value <> cboDept.ItemData(cboDept.ListIndex) Then
            lng执行科室ID = cboDept.ItemData(cboDept.ListIndex)
        End If
    
        On Error Resume Next
        If frmTechnicLog.ShowMe(Me, p医技工作站, lng科室ID, lng医嘱ID, lng发送号, Not .ParentRow.GroupRow, , lng执行科室ID, .Record(col_完成人).Value, mstrPrivs) Then
            err.Clear: On Error GoTo 0
            If blnRefresh Then Call LoadPatients '可能要更新执行状态
            FuncThingNew = True
        End If
    End With
End Function

Private Sub FuncThingModi()
    Dim lng科室ID As Long, lng医嘱ID As Long, lng发送号 As Long
    Dim str执行时间 As String, lng执行科室ID As Long
        
    If vsExec.TextMatrix(1, 0) = "" Then Exit Sub
    If vsExec.Row <> vsExec.FixedRows Then Exit Sub '只能操作最近一次执行
    
    If Val(gstr医嘱核对) > 0 And vsExec.TextMatrix(vsExec.FixedRows, COLExec("核对人")) <> "" Then
        MsgBox "该医嘱还已经核对，请取消核对后再试。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    With rptPati.SelectedRows(0)
        If .Record(col_综合状态).Value = 1 Then '子项和独项同执行状态
            MsgBox "该执行项目已经执行完成，不能再继续操作。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If .Record(COL_数据转出).Value = 1 Then
            MsgBox "该病人的本次" & Decode(.Record(col_来源).Value, "住院", "住院", "就诊") & "数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
                "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        lng科室ID = cboDept.ItemData(cboDept.ListIndex)
        lng医嘱ID = .Record(col_医嘱ID).Value
        lng发送号 = .Record(col_发送号).Value
        str执行时间 = vsExec.Cell(flexcpData, vsExec.Row, 1)
        If .Record(col_执行科室).Value <> cboDept.ItemData(cboDept.ListIndex) Then
            lng执行科室ID = cboDept.ItemData(cboDept.ListIndex)
        End If
        
        On Error Resume Next
        If frmTechnicLog.ShowMe(Me, p医技工作站, lng科室ID, lng医嘱ID, lng发送号, Not .ParentRow.GroupRow, str执行时间, lng执行科室ID, .Record(col_完成人).Value, mstrPrivs) Then
            err.Clear: On Error GoTo 0
            Call LoadExecList(lng医嘱ID, lng发送号)
        End If
    End With
End Sub

Private Sub FuncThingDel()
    Dim lng医嘱ID As Long, lng发送号 As Long
    Dim str执行时间 As String, strSQL As String
    
    If vsExec.TextMatrix(1, 0) = "" Then Exit Sub
    If vsExec.Row <> vsExec.FixedRows Then Exit Sub '只能操作最近一次执行
    
    With rptPati.SelectedRows(0)
        If .Record(col_综合状态).Value = 1 Then '子项和独项同执行状态
            MsgBox "该执行项目已经执行完成，不能再继续操作。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If Val(gstr医嘱核对) > 0 And vsExec.TextMatrix(vsExec.FixedRows, COLExec("核对人")) <> "" Then
            MsgBox "该医嘱还已经核对，请取消核对后再试。", vbInformation, gstrSysName
            Exit Sub
        End If
            
        If .Record(COL_数据转出).Value = 1 Then
            MsgBox "该病人的本次" & Decode(.Record(col_来源).Value, "住院", "住院", "就诊") & "数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
                "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
            Exit Sub
        End If
            
        If MsgBox("确实要删除该条执行情况吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        lng医嘱ID = .Record(col_医嘱ID).Value
        lng发送号 = .Record(col_发送号).Value
        str执行时间 = vsExec.Cell(flexcpData, vsExec.Row, 1)
    
        strSQL = "ZL_病人医嘱执行_Delete(" & lng医嘱ID & "," & lng发送号 & "," & _
            "To_Date('" & str执行时间 & "','YYYY-MM-DD HH24:MI:SS')," & IIf(Not .ParentRow.GroupRow, 1, 0) & ",0," & mlngDept & ")"
        
        On Error GoTo errH
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        err.Clear: On Error GoTo 0
        
        Call LoadPatients '可能要更新执行状态
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsExec_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBar
        
    If Button = 2 Then
        Set objPopup = cbsMain.Add("Popup", xtpBarPopup)
        With objPopup.Controls
            .Add xtpControlButton, conMenu_Manage_ThingAdd, "记录执行情况(&A)"
            .Add xtpControlButton, conMenu_Manage_ThingModi, "调整执行情况(&M)"
            .Add xtpControlButton, conMenu_Manage_ThingDel, "删除执行情况(&D)"
            .Add xtpControlButton, conMenu_Manage_ThingAudit, "核对"
            .Add xtpControlButton, conMenu_Manage_ThingDelAudit, "取消核对"
        End With
        
        vsExec.SetFocus
        objPopup.ShowPopup
    End If
End Sub

Private Sub Set诊疗项目费用设置()
     On Error Resume Next
    If gobjCISBase Is Nothing Then
        Set gobjCISBase = CreateObject("zl9CISBase.clsCISBase")
        If gobjCISBase Is Nothing Then
            MsgBox "诊疗基础部件(ZLCISBase)没有正确安装，该功能无法执行。", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    err.Clear: On Error GoTo 0
    
    Call gobjCISBase.CallSetClinicCharge(Val(cboDept.ItemData(cboDept.ListIndex)), 1, Me, gcnOracle, glngSys, gstrDBUser, E门诊调用, InStr(mstrPrivs, "诊疗项目费用设置") = 0)
End Sub

Private Function ShowBillAppend(ByVal lngRow As Long, ByRef blnExist As Boolean) As Boolean
'功能：显示指定行医嘱的单据附项内容
'返回：blnExist=医嘱是否存在单据附项内容
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lngidx As Long

    blnExist = False
    rtfAppend.Text = "": rtfAppend.SelStart = 0
    
    On Error GoTo errH
    
    strSQL = "Select 项目,内容 From 病人医嘱附件 Where 医嘱ID=[1] Order by 排列"
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人医嘱附件", "H病人医嘱附件")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIf(Val(rptPati.SelectedRows(0).Record(col_相关ID).Value) = 0, Val(rptPati.SelectedRows(0).Record(col_医嘱ID).Value), Val(rptPati.SelectedRows(0).Record(col_相关ID).Value)))
    If Not rsTmp.EOF Then
        With rtfAppend
            Do While Not rsTmp.EOF
                .SelBold = False
                .SelText = IIf(.Text = "", "", vbCrLf) & rsTmp!项目 & "：" & Nvl(rsTmp!内容)
                lngidx = .Find(rsTmp!项目 & "：", , , rtfNoHighlight Or rtfMatchCase)
                If lngidx <> -1 Then
                    .SelStart = lngidx
                    .SelLength = Len(rsTmp!项目 & "：")
                    .SelBold = True
                    .SelIndent = 100
                End If
                .SelStart = Len(.Text)
                
                rsTmp.MoveNext
            Loop
            
            '光标定位在第一个输入附项
            rsTmp.MoveFirst
            lngidx = .Find(rsTmp!项目 & "：", 0, , rtfNoHighlight Or rtfMatchCase)
            If lngidx <> -1 Then .SelStart = lngidx + Len(rsTmp!项目 & "：")
            
            Call zlControl.RTFSetFontSize(rtfAppend, IIf(mbytSize = 0, 9, 12))
        End With
        blnExist = True
    End If
    
    ShowBillAppend = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetFontSize(ByVal blnSetMainFont As Boolean)
'功能：进行界面字体的统一设置
'参数：blnSetMainFont  是否设置主界面字体 （用以区分子界面切换）
    If blnSetMainFont Then
        Call zlControl.SetPubFontSize(Me, mbytSize, "fraExec")
        Call SetControlPosition
        If Not mobjFrmBloodExe Is Nothing Then
            If mobjFrmBloodExe.Visible = True Then Call mobjFrmBloodExe.SetFontSize(IIf(mbytSize = 0, 9, 12))
        End If
    End If
        
    Select Case tbcSub.Selected.Tag
        Case "医嘱附费"
            Call mclsExpenses.SetFontSize(mbytSize)
        Case "门诊医嘱"
            Call mclsOutAdvices.SetFontSize(mbytSize)
        Case "住院医嘱"
            Call mclsInAdvices.SetFontSize(mbytSize)
        Case "住院病历"
            Call mclsInEPRs.SetFontSize(mbytSize)
        Case "门诊病历"
            Call mclsOutEPRs.SetFontSize(mbytSize)
        Case "护理"
            Call mclsTends.SetFontSize(mbytSize)
        Case "护理病历"
            Call mclsTendEPRs.SetFontSize(mbytSize)
        Case "新版护理"
            Call mclsTendsNew.SetFontSize(mbytSize)
                Case "新病历"
            On Error Resume Next
            Call mclsEMR.SetFontSize(mbytSize)
            err.Clear: On Error GoTo 0
    End Select
End Sub

Private Sub SetControlPosition()
'功能：设置主界面的空间位置大小以及字体
    Dim lngVcDis As Long
    lngVcDis = IIf(mbytSize = 0, 20, 50)
    lblAdvice.Font.Size = IIf(mbytSize = 0, 9, 12)
    
    fraDiag.Height = Me.TextHeight("字") + 120
    
    fraExec.Top = fraDiag.Top + fraDiag.Height + 10
    lblCash.Height = fraExec.Height - lblCash.Top - 30
    lblRec.Height = lblCash.Height
    
    vsExec.Top = fraExec.Top + fraExec.Height + 20
    '根据字高设置VSFlexGrid设置为两行显示
    vsExec.Height = Me.TextHeight("字") * 5
    
    picApplyInfo.Top = vsExec.Top
    picApplyInfo.Height = vsExec.Height
    rtfAppend.Top = lblApply.Top + lblApply.Height + 10
    rtfAppend.Height = picApplyInfo.Height - rtfAppend.Top
    '调用事件picExec_Resize
    picExec.Height = vsExec.Top + vsExec.Height
    
    Call zlControl.SetPubCtrlPos(False, 0, lblDept, 20, cboDept)
    Call zlControl.SetPubCtrlPos(False, 0, lblFind, 20, PatiIdentify)
    PatiIdentify.Left = IIf(mbytSize = 0, 990, 1050)
    chkFilter.Height = PatiIdentify.Height
    Call zlControl.SetPubCtrlPos(True, 1, chk执行状态(0), lngVcDis, chk执行状态(2), lngVcDis, chk执行状态(4))
    Call zlControl.SetPubCtrlPos(False, 0, Image1(0), 10, chk执行状态(0), IIf(mbytSize = 0, 30, 20), Image1(1), 10, chk执行状态(1))
    Call zlControl.SetPubCtrlPos(False, 0, Image1(2), 10, chk执行状态(2), IIf(mbytSize = 0, 30, 20), Image1(3), 10, chk执行状态(3))
    Call zlControl.SetPubCtrlPos(False, 0, Image1(4), 20, chk执行状态(4))
    Call picPati_Resize
End Sub

Private Sub AddMsgToLis(ByVal rsMsg As ADODB.Recordset)
'功能：将接收到的消息加入提醒列表中
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim i As Long
    
    On Error GoTo errH
    
    If Mid(rsMsg!提醒场合, 4, 1) <> "1" Then Exit Sub
    
    If InStr("," & rsMsg!部门IDs & ",", "," & cboDept.ItemData(cboDept.ListIndex) & ",") > 0 Or _
        InStr("," & rsMsg!提醒人员 & ",", "," & UserInfo.姓名 & ",") > 0 Then
        
        '判断列表是否已经有这类消息了
        For i = 0 To rptNotify.Rows.Count - 1
            If Not rptNotify.Rows(i).GroupRow Then
                If rptNotify.Rows(i).Record(C_消息).Value = rsMsg!类型编码 And rptNotify.Rows(i).Record.Tag = CStr(rsMsg!病人ID & "," & rsMsg!就诊id) Then
                    Exit Sub
                End If
            End If
        Next
        strSQL = "Select a.住院号, a.姓名, a.性别, a.年龄, a.当前床号 As 床号, a.险类 From 病人信息 A Where a.病人id =[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsMsg!病人ID))
        
        Call AddReportRow(rsMsg!病人ID & "," & rsMsg!就诊id, rsMsg!病人ID, rsMsg!就诊id, Nvl(rsTmp!姓名), Nvl(rsMsg!消息内容), rsMsg!类型编码 & "", _
                rsMsg!优先程度 & "", Format(rsMsg!登记时间 & "", "yyyy-MM-dd HH:mm:ss"), rsMsg!业务标识 & "")
        rptNotify.Populate
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub AddReportRow(ParamArray arrInput() As Variant)
'功能：向消息提配列表中增加一行
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim Index As Integer
    
    Set objRecord = Me.rptNotify.Records.Add()
    objRecord.Tag = arrInput(Index): Index = Index + 1         'Tag值
    Set objItem = objRecord.AddItem(""): objItem.Icon = 3
    objRecord.AddItem Val(arrInput(Index)): Index = Index + 1  '病人id
    objRecord.AddItem Val(arrInput(Index)): Index = Index + 1  '就诊id
    objRecord.AddItem CStr(arrInput(Index)): Index = Index + 1 '姓名
    
    Set objItem = objRecord.AddItem(CStr(arrInput(Index)))     '状态，内容
    objItem.Caption = CStr(arrInput(Index)): Index = Index + 1
    
    objRecord.AddItem CStr(arrInput(Index)): Index = Index + 1  '消息编号
    objRecord.AddItem Val(arrInput(Index)): Index = Index + 1   '序号
    objRecord.AddItem CStr(arrInput(Index)): Index = Index + 1  '日期
    objRecord.AddItem CStr(arrInput(Index)): Index = Index + 1  '业务标识
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function LocatePati(ByVal strTag As String) As Boolean
'功能：通过reportControl的Record.Tag值定位病人
'参数：strTag 医嘱id_发送号

    Dim blnEnabled As Boolean
    Dim objRow As ReportRow
    
    For Each objRow In rptPati.Rows
        If objRow.GroupRow Then objRow.Expanded = True
        If Not objRow.GroupRow Then
            If InStr(objRow.Record.Tag & "_", "_" & strTag & "_") > 0 Then
                blnEnabled = timRefresh.Enabled
                timRefresh.Enabled = False '避免连锁引起刷新提醒内容
                Set rptPati.FocusedRow = objRow '选中,显示,[激活Change事件]
                timRefresh.Enabled = blnEnabled
                LocatePati = True: Exit Function
            End If
        End If
    Next
End Function

Private Sub FuncExecCancelBatch()
'功能：对一个病人选择的多个项目取消执行完成
    Dim lng组ID As Long, lng医嘱ID As Long, lng发送号 As Long
    Dim str诊疗类别 As String, strSQL As String, byt来源 As Byte
    Dim strOwner As String, strUserName As String
    Dim cnNew As ADODB.Connection, rsTmp As ADODB.Recordset
    Dim i As Long
    Dim arrSQL As Variant
    Dim strMsg As String
    Dim blnTrans As Boolean
    Dim blnGroupRow As Boolean
    
    On Error GoTo errH
    
    arrSQL = Array()
    For i = 0 To rptPati.Rows.Count - 1
        If Not rptPati.Rows(i).GroupRow Then
            If rptPati.Rows(i).Record(col_选择).Checked Then
                With rptPati.Rows(i).Record
                    '可以不填写执行情况直接完成执行
                    If Val(.Item(col_综合状态).Value) <> 1 Then
                        MsgBox "病人""" & .Item(col_姓名).Value & """的项目""" & .Item(col_内容).Value & """当前不处于已执行状态，不能取消执行。", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    
                    '检查病人是否开始审核
                    If Val(.Item(col_审核标志).Value & "") >= 1 And mbyt病人审核方式 = 1 Then
                        strMsg = strMsg & "," & .Item(col_单据号).Value
                    Else
                        blnGroupRow = Not rptPati.Rows(i).ParentRow.GroupRow
                        '数据是否转出
                        If .Item(COL_数据转出).Value = 1 Then
                            MsgBox "病人""" & .Item(col_姓名).Value & """的项目""" & .Item(col_内容).Value & """的数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
                                "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
                            Exit Sub
                        End If
                      
                        lng医嘱ID = .Item(col_医嘱ID).Value
                        lng发送号 = .Item(col_发送号).Value
                        str诊疗类别 = .Item(COL_诊疗类别).Value
                        lng组ID = IIf(.Item(col_相关ID).Value = 0, .Item(col_医嘱ID).Value, .Item(col_相关ID).Value)
                        If Val(.Item(col_记录性质).Value) <> 1 Then
                            '费用结帐判断
                            byt来源 = IIf(.Item(col_来源).Value = "住院" And Val(.Item(col_门诊记帐).Value) = 0, 2, 1)
                            If Not ItemCanCancel(lng医嘱ID, lng发送号, lng组ID, str诊疗类别, blnGroupRow, .Item(COL_数据转出).Value = 1, byt来源) Then Exit Sub
                        End If
                        
                        
                        '判断是否皮试,再填写结果
                        strSQL = "Select A.诊疗类别,A.皮试结果,B.操作类型,Nvl(B.标本部位,'阳性(+);阴性(-)') as 标本部位 From 病人医嘱记录 A,诊疗项目目录 B Where A.诊疗项目ID=B.ID And A.ID=[1]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .Item(col_医嘱ID).Value)
                        If Not rsTmp.EOF Then
                            '已经填写了皮试结果则不再填写
                            If rsTmp!诊疗类别 = "E" And Nvl(rsTmp!操作类型) = "1" And Not IsNull(rsTmp!皮试结果) Then
                                '先作身份验证
                                If mbln皮试验证 Then
                                    Set cnNew = New ADODB.Connection
                                    strUserName = zlDatabase.UserIdentify(Me, "在取消完成皮试医嘱前，请您先输入用户名和密码进行身份验证。", glngSys, p住院医嘱发送, "皮试医嘱结果", cnNew)
                                    If strUserName = "" Then Exit Sub
                                End If
                                strSQL = "ZL_病人医嘱执行_Cancel(" & lng医嘱ID & "," & lng发送号 & "," & IIf(mbln皮试验证, 1, 0) & "," & IIf(blnGroupRow, 1, 0) & "," & mlngDept & ")"
                            Else
                                strSQL = "ZL_病人医嘱执行_Cancel(" & lng医嘱ID & "," & lng发送号 & ",Null ," & IIf(blnGroupRow, 1, 0) & "," & mlngDept & ")"
                            End If
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = strSQL
                        End If
                    End If
                End With
            End If
        End If
    Next
    
    If strMsg <> "" Then
        strMsg = "以下单据号的病人费用正在审核阶段，不允许操作医嘱和费用：" & vbCrLf & Mid(strMsg, 2) & "。"
        MsgBox strMsg, vbInformation, gstrSysName
    End If
    
    If UBound(arrSQL) = -1 Then Exit Sub
 
    If Not cnNew Is Nothing Then
        strOwner = SystemOwner(glngSys)
        
        On Error GoTo errNew
        cnNew.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            Call SQLTest(App.ProductName, Me.Caption, arrSQL(i))
            cnNew.Execute strOwner & "." & arrSQL(i), , adCmdStoredProc
            Call SQLTest
        Next
        
        cnNew.CommitTrans: blnTrans = False
        cnNew.Close: Set cnNew = Nothing
        On Error GoTo errH
    Else
        gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
        Next
        gcnOracle.CommitTrans: blnTrans = False
    End If
     
    Call LoadPatients
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Exit Sub
errNew:
    cnNew.RollbackTrans
    MsgBox err.Number & vbCrLf & vbCrLf & err.Description, vbInformation, gstrSysName
    cnNew.Close: Set cnNew = Nothing
End Sub

Private Function Check销帐费用(ByVal bln单独执行 As Boolean, ByVal lng医嘱ID As Long, ByVal lng组ID As Long, ByVal str诊疗类别 As String, ByVal int费用性质 As Integer, ByVal strNO As String) As Boolean
'功能：获取某条医嘱，或某组医嘱的是否存在已经销帐的费用
'       bln单独执行 是否单独执行，检验检查类存在单据的医嘱的单独执行某一部位，某一部分检查
'       lng医嘱ID 该条医嘱ID
'       lng组ID 没有父医嘱，或者父医嘱时为医嘱ID,子医嘱为相关ID
'       str诊疗类别 该医嘱的诊疗类别
'       int费用性质 1-门诊费用，2-住院费用'
    Dim rsTmp As ADODB.Recordset, strSQL As String, strTable As String
    strTable = IIf(int费用性质 = 1, "门诊费用记录", "住院费用记录")
    On Error GoTo errH
    If bln单独执行 Then
        lng组ID = lng医嘱ID
        strSQL = "Select -1 * Sum(Nvl(a.付数, 1) * a.数次 / b.数量) As 最大已销数" & vbNewLine & _
                "From " & strTable & " A, 病人医嘱计价 B" & vbNewLine & _
                "Where a.医嘱序号 = [1] And A.NO=[3] And b.医嘱id = a.医嘱序号 And b.收费细目id = a.收费细目id And Nvl(B.费用性质,0)=0 And a.记录状态 = 2 And a.记录性质 in(1,2,11) And a.价格父号 Is Null And" & vbNewLine & _
                "      a.收费类别 Not In ('5', '6', '7') And Not Exists" & vbNewLine & _
                " (Select 1 From 材料特性 Where 材料id = a.收费细目id And 跟踪在用 = 1)"

    Else
        strSQL = "Select Max(c.已销数) 最大已销数" & vbNewLine & _
                "From (Select -1 * Sum(Nvl(a.付数, 1) * a.数次 / b.数量) As 已销数" & vbNewLine & _
                "       From " & strTable & " A, 病人医嘱计价 B" & vbNewLine & _
                "       Where a.医嘱序号 In (Select ID From 病人医嘱记录 Where (ID = [1] Or 相关id = [1]) And A.NO=[3] And 诊疗类别 = [2]) And b.医嘱id = a.医嘱序号 And" & vbNewLine & _
                "             b.收费细目id = a.收费细目id And Nvl(B.费用性质,0)=0 And a.记录状态 = 2 And a.记录性质 in(1,2) And a.价格父号 Is Null And a.收费类别 Not In ('5', '6', '7') And" & vbNewLine & _
                "             Not Exists" & vbNewLine & _
                "        (Select 1 From 材料特性 Where 材料id = a.收费细目id And 跟踪在用 = 1) " & vbNewLine & _
                "       Group By  a.医嘱序号,a.收费细目id) C"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng组ID, str诊疗类别, strNO)
    If rsTmp.RecordCount <> 0 Then
        Check销帐费用 = (Val(rsTmp!最大已销数 & "") > 0)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Set收费标记(ByVal int病人来源 As Integer, ByVal bln单独执行 As Boolean, ByVal lng医嘱ID As Long, ByVal lng相关ID As Long, ByVal lng发送号 As Long, _
    ByVal str类别 As String, ByVal str单据号 As String, ByVal int记录性质 As Integer, ByVal int门诊记帐 As Integer, ByVal blnMove As Boolean, ByVal dat发送时间 As Date) As Boolean
'功能：判断界面上的"收"字是否显示，调用费用公共部件接口进行判断
    Dim strSQL As String, strTab As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, strNos As String
    Dim bln门诊 As Boolean
    Dim bytState As Byte
    
    On Error GoTo errH
    
    If int病人来源 = 2 And int记录性质 = 2 And int门诊记帐 = 0 Then
        strTab = "住院费用记录"
    Else
        strTab = "门诊费用记录"
        bln门诊 = True
    End If
    
    strSQL = "select a.no from (" & _
        " Select a.no" & _
        " From " & strTab & " A,病人医嘱记录 B" & _
        " Where A.NO=[4] And A.医嘱序号+0=B.ID And MOD(A.记录性质,10)=[5]" & IIf(bln单独执行, " And B.ID=[2]", "") & _
        " Union ALL " & _
        " Select A.NO" & _
        " From 病人医嘱记录 C," & strTab & " B,病人医嘱附费 A" & _
        " Where A.NO=B.NO And A.记录性质=MOD(B.记录性质,10) And A.医嘱ID=B.医嘱序号+0" & IIf(bln单独执行, " And A.医嘱ID=[2]", _
        " And A.医嘱ID IN (Select ID From 病人医嘱记录 Where (ID=[1] Or 相关ID=[1]) And 诊疗类别=[6])") & _
        " And A.发送号=[3] And A.医嘱ID=C.ID And A.记录性质=[5]) a group by a.no"
    If blnMove Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        strSQL = Replace(strSQL, "病人医嘱附费", "H病人医嘱附费")
        strSQL = Replace(strSQL, strTab, "H" & strTab)
    ElseIf zlDatabase.DateMoved(dat发送时间) Then
        strSQL = strSQL & " Union ALL " & Replace(strSQL, strTab, "H" & strTab)
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ItemHaveCash", IIf(lng相关ID <> 0, lng相关ID, lng医嘱ID), lng医嘱ID, lng发送号, str单据号, int记录性质, str类别)
    
    For i = 1 To rsTmp.RecordCount
        strNos = strNos & "," & rsTmp!NO
        rsTmp.MoveNext
    Next
    If strNos <> "" Then
        strNos = Mid(strNos, 2)
        If int记录性质 = 2 Then
            Call mclsPExp.zlGetBalanceStatus(strNos, bytState, bln门诊)
        Else
            Call mclsPExp.zlGetBillChargeStatus(strNos, bytState)
        End If
    End If
    Set收费标记 = (bytState = 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub picMsg_Resize()
'
    Dim lngTmp As Long
   
    On Error Resume Next
    
    lngTmp = picMsg.Height
    
    If mbytSize = 0 Then
        If lngTmp < 1010 Then
            lngTmp = 1010
        End If
    Else
        If lngTmp < 1130 Then
            lngTmp = 1130
        End If
    End If
    
    rptNotify.Top = 0
    rptNotify.Left = 0
    rptNotify.Width = picMsg.Width
    rptNotify.Height = lngTmp
End Sub

Private Sub FuncAppendBill()
'功能：调三方附费窗口

    Dim rsTmp As ADODB.Recordset
    Dim lng医嘱ID As Long
    Dim strSQL As String
    Dim lngTmp As Long
    
    Dim strPar As String
    Dim str来源系统 As String
    Dim str病人来源 As String
    Dim str病人标识 As String
    Dim str就诊标识 As String
    Dim str医嘱编号 As String
    Dim str医嘱发送号 As String
    Dim str当前科室标识 As String
    Dim str当前科室编码 As String
    Dim str当前科室名称 As String
    Dim str操作员标识 As String
    Dim str操作员编码 As String
    Dim str操作员姓名 As String
    Dim str院区编码 As String '站点
    Dim str院区名称 As String
    Dim str用户名 As String
    Dim str用户密码 As String
    
    
    On Error GoTo errH
    
    With rptPati.SelectedRows(0)
    
        lng医嘱ID = .Record(col_医嘱ID).Value
        
        '院区参数--ZLHIS系统中的站点信息
        If gstrNodeNo <> "" And gstrNodeNo <> "-" Then
            strSQL = "Select 编号,名称 From Zlnodelist Where 编号=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, gstrNodeNo)
            If Not rsTmp.EOF Then
                str院区编码 = rsTmp!编号 & ""
                str院区名称 = rsTmp!名称 & ""
            End If
        End If
        
        '病人标识-- 病人ID
        str病人标识 = mlng病人ID
        
        '就诊标识--门诊病人 传空,住院病人 主页ID
        If .Record(col_来源).Value = "门诊" Then
            str就诊标识 = ""
        Else
            str就诊标识 = .Record(col_主页ID).Value
        End If
        
        '来源系统-- 01 ZLHIS中的病人，02 新病人
        str来源系统 = "01"
        lngTmp = Val(.Record(col_挂号ID).Value)
        If lngTmp <> 0 Then
            strSQL = "Select 1 From 病人挂号记录 a where a.附加标志=3 and a.id=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngTmp)
            If Not rsTmp.EOF Then '找到数据说明是新门诊病人
                str来源系统 = "02"
                str就诊标识 = ""
            End If
        End If
        
        '病人来源 -- 0-门诊/1-住院/2-体检
        str病人来源 = Decode(.Record(col_来源).Value, "门诊", 0, "住院", 1, "体检", 2, "外来", 3)
 
        '发送号
        str医嘱发送号 = .Record(col_发送号).Value
        
        str医嘱编号 = lng医嘱ID
        
        strSQL = "Select id,资源id,编码 As 当前科室编码,名称 as 当前科室名称 From 部门表 Where ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDept)
        
        If str来源系统 = "02" Then
            str当前科室标识 = rsTmp!资源id & ""
	    str操作员标识 = Sys.RowValue("人员表", UserInfo.ID, "资源id")
        Else
            str当前科室标识 = rsTmp!ID & ""
	    str操作员标识 = UserInfo.ID
        End If
        str当前科室编码 = rsTmp!当前科室编码 & ""
        str当前科室名称 = rsTmp!当前科室名称 & ""
          
        str操作员编码 = UserInfo.编号
        str操作员姓名 = UserInfo.姓名
    End With
    
    str用户名 = UserInfo.用户名 ' "ZLHIS"
    
    If mstr密码 = "" Then
        mstr密码 = GetConnPassword
    End If
    str用户密码 = mstr密码
    
    strPar = _
        "{" & _
            """来源系统"":""" & str来源系统 & """," & _
            """病人来源"":" & str病人来源 & "," & _
            """病人标识"":""" & str病人标识 & """," & _
            """就诊标识"":""" & str就诊标识 & """," & _
            """医嘱编号"":""" & str医嘱编号 & """," & _
            """医嘱发送号"":""" & str医嘱发送号 & """," & _
            """当前科室标识"":""" & str当前科室标识 & """," & _
            """当前科室编码"":""" & str当前科室编码 & """," & _
            """当前科室名称"":""" & str当前科室名称 & """," & _
            """操作员标识"":""" & str操作员标识 & """," & _
            """操作员编码"":""" & str操作员编码 & """," & _
            """操作员姓名"":""" & str操作员姓名 & """," & _
            """院区编码"":""" & str院区编码 & """," & _
            """院区名称"":""" & str院区名称 & """," & _
            """用户名"":""" & str用户名 & """," & _
            """用户密码"":""" & str用户密码 & """" & _
        "}"
    '调用补费
    Call mobjAppendBill.EditChargeBill(strPar)
 
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetConnPassword()
    '获取当前用户登录密码
    Dim objLogin As Object
    
    On Error Resume Next
    Set objLogin = CreateObject("zlLogin.clsLogin")
    If objLogin Is Nothing Then
        err.Clear
        MsgBox "创建zlLogin部件对象失败，请检查文件是否存在并且正确注册。", vbExclamation, gstrSysName
        Exit Function
    End If
    
    GetConnPassword = objLogin.InputPwd
End Function

Private Function NewOut收费(ByVal lng医嘱ID As Long) As Long
'功能：判断新门诊病人是否收费，调用新门诊系统的服务
'返回：0-走以前逻辑，1-提示并禁止，2-成功验证通过
    Dim strJsIn As String
    Dim strJsOut As String
    Dim strErr As String
    Dim int已收费 As Integer
    Dim blnTmp As Boolean
    Dim lngRes As Long
    
    Screen.MousePointer = 11
    strJsIn = "{""input"":{""head"":{""bizno"":""RJ001"",""sysno"":""ZLDAYROOM"",""time"":"""",""action_no"":"""",""tarno"":""03""},""apply_id"":" & lng医嘱ID & "}}"
    blnTmp = Sys.NewSystemSvr("新门诊系统", "判断医嘱是否收费", strJsIn, strJsOut, strErr)
    Screen.MousePointer = 0
    If strErr <> "" Then
        MsgBox strErr, vbInformation, gstrSysName
        lngRes = 1
        NewOut收费 = lngRes
        Exit Function
    End If
    
    If blnTmp Then
        If strJsOut <> "" Then
            If Val(zlStr.JSONParse("result", strJsOut) & "") <> 1 Then
                MsgBox zlStr.JSONParse("errmsg", strJsOut) & "", vbInformation, gstrSysName
                lngRes = 1
            End If
            int已收费 = Val(zlStr.JSONParse("kacnt_sign", strJsOut) & "")
        End If
        If int已收费 <> 1 Then
            MsgBox "该病人还存在未收费的费用，请检查。", vbInformation, gstrSysName
            lngRes = 1
        Else
            lngRes = 2
        End If
    End If
    NewOut收费 = lngRes
End Function