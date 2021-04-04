VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Begin VB.Form frmMakeupPrintBill 
   Caption         =   "住院结账补打发票"
   ClientHeight    =   8415
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11790
   Icon            =   "frmMakeupPrintBill.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8415
   ScaleWidth      =   11790
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picBalance 
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   6480
      ScaleHeight     =   4695
      ScaleWidth      =   4380
      TabIndex        =   17
      Top             =   1140
      Width           =   4380
      Begin VSFlex8Ctl.VSFlexGrid vsBalance 
         Height          =   2685
         Left            =   525
         TabIndex        =   18
         Top             =   300
         Width           =   8505
         _cx             =   15002
         _cy             =   4736
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
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmMakeupPrintBill.frx":030A
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
      Begin VB.Label lblSum 
         AutoSize        =   -1  'True
         Caption         =   "补打合计:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   465
         TabIndex        =   19
         Top             =   3165
         Width           =   1155
      End
   End
   Begin VB.PictureBox PicDetail 
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   75
      ScaleHeight     =   2775
      ScaleWidth      =   5535
      TabIndex        =   15
      Top             =   4290
      Width           =   5535
      Begin VSFlex8Ctl.VSFlexGrid vsDetail 
         Height          =   2685
         Left            =   180
         TabIndex        =   16
         Top             =   330
         Width           =   8505
         _cx             =   15002
         _cy             =   4736
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
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmMakeupPrintBill.frx":0384
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
   Begin VB.PictureBox picList 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   90
      ScaleHeight     =   2295
      ScaleWidth      =   5400
      TabIndex        =   13
      Top             =   1290
      Width           =   5400
      Begin VSFlex8Ctl.VSFlexGrid vsList 
         Height          =   2685
         Left            =   105
         TabIndex        =   14
         Top             =   75
         Width           =   4650
         _cx             =   8202
         _cy             =   4736
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
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmMakeupPrintBill.frx":039A
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
   Begin VB.PictureBox picCon 
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
      Height          =   675
      Left            =   180
      ScaleHeight     =   675
      ScaleWidth      =   14475
      TabIndex        =   5
      Top             =   135
      Width           =   14475
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   375
         Left            =   615
         TabIndex        =   24
         Top             =   150
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   661
         Appearance      =   2
         IDKindStr       =   "姓|姓名|0;医|医保号|0;身|身份证号|0;IC|IC卡号|1;门|门诊号|0;住|住院号|0;就|就诊卡|0"
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
         MustSelectItems =   "姓名"
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
         Left            =   1305
         MaxLength       =   100
         TabIndex        =   9
         Top             =   150
         Width           =   2040
      End
      Begin VB.CommandButton cmdBrush 
         Caption         =   "刷新(&N)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9195
         TabIndex        =   6
         Top             =   150
         Width           =   1100
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   345
         Left            =   5520
         TabIndex        =   7
         Top             =   180
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   93323267
         CurrentDate     =   36588
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   345
         Left            =   7515
         TabIndex        =   10
         Top             =   180
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   93323267
         CurrentDate     =   36588
      End
      Begin VB.CheckBox chkDate 
         Caption         =   "忽略发生时间"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3690
         TabIndex        =   8
         Top             =   210
         Value           =   1  'Checked
         Width           =   1890
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
         Left            =   150
         TabIndex        =   12
         Top             =   195
         Width           =   480
      End
      Begin VB.Label lblArang 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "～"
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
         Left            =   7200
         TabIndex        =   11
         Top             =   225
         Width           =   240
      End
   End
   Begin VB.PictureBox picDown 
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   45
      ScaleHeight     =   660
      ScaleWidth      =   11700
      TabIndex        =   0
      Top             =   7170
      Width           =   11700
      Begin VB.TextBox txtInvoice 
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
         IMEMode         =   3  'DISABLE
         Left            =   2985
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "补打(&O)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   8955
         TabIndex        =   4
         Top             =   225
         Width           =   1100
      End
      Begin VB.CommandButton cmdSelAll 
         Caption         =   "全选(&A)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   105
         TabIndex        =   3
         ToolTipText     =   "热键：Ctrl+A"
         Top             =   225
         Width           =   1100
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "全清(&R)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1245
         TabIndex        =   2
         ToolTipText     =   "热键：Ctrl+R"
         Top             =   225
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   10095
         TabIndex        =   1
         Top             =   225
         Width           =   1100
      End
      Begin VB.Label lblFormat 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   240
         Left            =   5325
         TabIndex        =   23
         Top             =   300
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label lblFact 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "票号"
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
         Height          =   240
         Left            =   2490
         TabIndex        =   22
         Top             =   300
         Width           =   480
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   20
      Top             =   8055
      Width           =   11790
      _ExtentX        =   20796
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
            Picture         =   "frmMakeupPrintBill.frx":03B0
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15716
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
   Begin MSComctlLib.ImageList imlPaneIcons 
      Left            =   1500
      Top             =   210
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakeupPrintBill.frx":0C44
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakeupPrintBill.frx":0F98
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   30
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmMakeupPrintBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum mPanel
    Pane_Search = 1
    Pane_List = 2
    Pane_Detail = 3
    Pane_Balance = 4
End Enum
'-----------------------------------------------------------------------------------
'结算卡相关
Private mSquareCard As SquareCard '结算卡相关
Private mstrPassWord As String
Private mbytInvoiceKind As Byte
'-----------------------------------------------------------------------------------
Private mrsInfo As ADODB.Recordset
Private mstrFindNO As String, mstrFindFpNo As String
Private mrsList As ADODB.Recordset  '单据列表
Private mrsDetail As ADODB.Recordset
Private mrsBalance As ADODB.Recordset
Private mstrNOs As String
Private mlngModule As Long
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private WithEvents mobjIDCard As zlIDCard.clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private mblnValid As Boolean
Private mblnSel As Boolean
Private mstrPrivs As String
Private mintSucces As Integer  '成功打印次数
Private mlng病人ID As Long
Private mblnFirst As Boolean
Private mblnNotClick As Boolean
Private mobjInvoice As clsInvoice
Private mobjFact As clsFactProperty

Private mlng领用ID As Long
Private mintInsure As Integer
Private mintPrintNums As Boolean '打印票据张数
Private mblnNOMoved As Boolean

Public Function zlRePrintBill(ByVal frmMain As Object, ByVal lngModule As Long, _
    ByVal strPrivs As String, Optional lng病人ID As Long = 0) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重打票据入口
    '返回:打印成功,返回true,否则返回Flase
    '编制:刘兴洪
    '日期:2013-01-05 15:21:03
    '问题:56283
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    mlngModule = lngModule: mstrPrivs = strPrivs: mintSucces = 0
    mlng病人ID = lng病人ID
    mbytInvoiceKind = IIf(gbytInvoiceKind = 0, 3, 1)
    
    If frmMain Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
    zlRePrintBill = mintSucces > 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function InitPanel()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化条件区域
    '编制:刘兴洪
    '日期:2013-01-05 15:22:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane, strKey As String
    Dim objTemp As Object
    With dkpMan
        .ImageList = imlPaneIcons
        Set objPane = .CreatePane(mPanel.Pane_Search, 200, 100, DockLeftOf, Nothing)
        objPane.Tag = mPanel.Pane_Search
        objPane.Title = "条件设置": objPane.Options = PaneNoCloseable Or PaneNoHideable Or PaneNoCaption
        objPane.MaxTrackSize.Height = 675 \ Screen.TwipsPerPixelY
        objPane.MinTrackSize.Height = 675 \ Screen.TwipsPerPixelY
        objPane.Handle = picCon.hWnd
        Set objTemp = .CreatePane(mPanel.Pane_List, 300, 100, DockBottomOf, objPane)
        objTemp.Tag = mPanel.Pane_List
        objTemp.Title = "结账单据列表": objTemp.Options = PaneNoCloseable Or PaneNoHideable
        objTemp.Handle = picList.hWnd
        
        Set objPane = .CreatePane(mPanel.Pane_Balance, 100, 100, DockRightOf, objTemp)
        objPane.Tag = mPanel.Pane_Balance
        objPane.Title = "结算信息列表": objPane.Options = PaneNoCloseable Or PaneNoHideable
        objPane.Handle = picBalance.hWnd
       '
        Set objPane = .CreatePane(mPanel.Pane_Detail, 300, 100, DockBottomOf, objTemp)
        objPane.Tag = mPanel.Pane_Detail
        objPane.Title = "结账明细列表": objPane.Options = PaneNoCloseable Or PaneNoHideable
        objPane.Handle = PicDetail.hWnd
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
    dkpMan.RecalcLayout: DoEvents
    zlRestoreDockPanceToReg Me, dkpMan, "区域"
End Function
Private Sub chkDate_Click()
    dtpBegin.Enabled = chkDate.Value <> 1
    dtpEnd.Enabled = chkDate.Value <> 1
End Sub
Private Sub cmdBrush_Click()
        Call ReadListData
End Sub
Private Sub cmdCancel_Click()
     Unload Me
End Sub

Private Sub cmdClear_Click()
    On Error GoTo errHandle
    With vsList
        If .ColIndex("选择") < 0 Then Exit Sub
        .Cell(flexcpChecked, 1, .ColIndex("选择"), .Rows - 1, .ColIndex("选择")) = 0
    End With
    Call SetBlanceShow
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function zlMakeupPrint(ByVal lng病人ID As Long, lng主页ID As Long, _
    ByVal strNO As String, ByVal lng结帐ID As String, ByVal intInsure As Integer, _
    Optional strInvoice As String = "", Optional ByVal bytFunc As Byte = 1) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:补打结帐票据
    '返回:补打成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-01-05 15:24:07
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objFact As New clsFactProperty
    
    On Error GoTo errHandle
    If lng结帐ID <= 0 Or lng病人ID = 0 Then Exit Function
 
    mobjFact.LastUseID = mlng领用ID
    Call frmPrint.ReportPrint(2, strNO, lng结帐ID, mobjFact, strInvoice, , , , lng病人ID, mobjFact.打印格式)
        
    '银医一卡通写卡，85950
    Call WriteInforToCard(Me, mlngModule, mstrPrivs, gobjSquare.objSquareCard, 0, bytFunc, lng结帐ID, lng病人ID)
    
    mintSucces = mintSucces + 1
    zlMakeupPrint = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckFP(ByVal lng结帐ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查发票是否正确
    '返回: 发票合法 返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-07-12 11:30:22
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intNum As Integer
    
     On Error GoTo errHandle
    intNum = 1
    If lng结帐ID = 0 Then
        MsgBox "不存在需要补打的票据", vbInformation, gstrSysName
        Exit Function
    End If
     If Not gblnStrictCtrl Then
        If Len(txtInvoice.Text) <> gbytFactLength And txtInvoice.Text <> "" Then
            MsgBox "票据号码长度应该为 " & gbytFactLength & " 位！", vbInformation, gstrSysName
            txtInvoice.SetFocus: Exit Function
        End If
        CheckFP = True
        Exit Function
     End If
     
    If Trim(txtInvoice.Text) = "" Then
        MsgBox "必须输入一个有效的票据号码！", vbInformation, gstrSysName
        txtInvoice.SetFocus: Exit Function
    End If
    
InvoiceHandle:
    If zlGetInvoiceGroupUseID(mlng领用ID, intNum, txtInvoice.Text) = False Then
        Exit Function
    End If
    '并发操作检查,票号是否已用
    If CheckBillRepeat(mlng领用ID, mbytInvoiceKind, txtInvoice.Text) Then
            'Tag：问题：24363:刘兴洪：主要是解决自动生成的号是否被用户更改，主要解决：
            If txtInvoice.Locked = False And txtInvoice.Tag <> Trim(txtInvoice.Text) Then
                MsgBox "票据号""" & txtInvoice.Text & """已经被使用，请重新输入。", vbInformation, gstrSysName
                txtInvoice.SetFocus: Exit Function
            Else
                Call RefreshFact
                If txtInvoice.Text = "" Then
                    txtInvoice.SetFocus: Exit Function
                Else
                    MsgBox "当前票据号已经被使用，已重新获取票据号:" & txtInvoice.Text, vbInformation, gstrSysName
                    txtInvoice.SetFocus: Exit Function
                End If
            End If
    End If
   CheckFP = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub cmdOK_Click()

     Dim str结帐IDs As String, lng结帐ID As Long, strNO As String
     Dim lngRow As Long, intInsure As Integer
     Dim lng病人ID As Long, lng主页ID As Long
     Dim bytFunc As Byte '结帐类型
     
    On Error GoTo errHandle

     If mrsInfo Is Nothing Then Exit Sub
     If mrsInfo.State <> 1 Then Exit Sub
     If mrsInfo.RecordCount = 0 Then Exit Sub
     
     lng病人ID = Val(Nvl(mrsInfo!病人ID))
     lng主页ID = Val(Nvl(mrsInfo!主页ID))
     
     With vsList
        str结帐IDs = ""
        For lngRow = 1 To .Rows - 1
            If GetVsGridBoolColVal(vsList, lngRow, .ColIndex("选择")) Then
                 lng结帐ID = Val(.TextMatrix(lngRow, .ColIndex("结帐ID")))
                 strNO = Trim(.TextMatrix(lngRow, .ColIndex("单据号")))
                 intInsure = Val(.TextMatrix(lngRow, .ColIndex("险类ID")))
                 bytFunc = IIf(.TextMatrix(lngRow, .ColIndex("结帐类型")) = "门诊结帐", 0, 1)
                 If lng结帐ID <> 0 Then
                    If Not CheckFP(lng结帐ID) Then Exit Sub
                    
                     str结帐IDs = str结帐IDs & "," & lng结帐ID
                    If lng病人ID <> Val(.TextMatrix(lngRow, .ColIndex("病人ID"))) _
                        Or Val(.TextMatrix(lngRow, .ColIndex("主页ID"))) <> lng主页ID _
                        Or intInsure <> mintInsure Then
                        
                        lng病人ID = Val(.TextMatrix(lngRow, .ColIndex("病人ID"))): lng主页ID = Val(.TextMatrix(lngRow, .ColIndex("主页ID")))
                        mintInsure = intInsure
                        
                        '处理发票相关信息
                        Call ReInitPatiInvoice(Val(Nvl(mrsInfo!病人ID)), Val(Nvl(mrsInfo!主页ID)), intInsure)
                        Call RefreshFact    '重新处理发票
                    End If
                    Call zlMakeupPrint(lng病人ID, lng主页ID, strNO, lng结帐ID, intInsure, Trim(txtInvoice.Text), bytFunc)
                    '重新取发票
                    Call RefreshFact
                 End If
            End If
        Next
     End With
     If str结帐IDs = "" Then
        MsgBox "未选择要补打的票据", vbOKOnly, gstrSysName
        Exit Sub
     End If
     Call zlClearPatiInfor
     If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdSelAll_Click()
    On Error GoTo errHandle
    With vsList
        If .ColIndex("选择") < 0 Then Exit Sub
        .Cell(flexcpChecked, 1, .ColIndex("选择"), .Rows - 1, .ColIndex("选择")) = -1
    End With
    Call SetBlanceShow
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case Pane_Search    '1
        Item.Handle = picCon.hWnd
    Case Pane_List      ' 2
        Item.Handle = picList.hWnd
    Case Pane_Detail    '3
        Item.Handle = PicDetail.hWnd
    Case Pane_Balance  ' 4
        Item.Handle = picBalance.hWnd
    End Select
End Sub
Private Sub dkpMan_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Err = 0: On Error Resume Next
     Bottom = stbThis.Height + picDown.Height
End Sub

Private Sub Form_Activate()
    If Not mblnFirst Then Exit Sub
    On Error GoTo errHandle
    Call zlClearPatiInfor
    mblnFirst = False
    If mlng病人ID <> 0 Then
        If GetPatient(IDKind.GetCurCard, "-" & mlng病人ID, False) = False Then
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
            zlControl.TxtSelAll txtPatient
            Exit Sub
        End If
         vsList.SetFocus: Exit Sub
    End If
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub Form_Load()

    On Error GoTo errHandle
    cmdBrush.Enabled = False
    mblnFirst = True
    lblFormat.Alignment = 0
    dtpBegin.MaxDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd") & " 23:59:59")
    dtpEnd.MaxDate = dtpBegin.MaxDate
    dtpBegin.Value = Format(DateAdd("d", -7, dtpBegin.MaxDate), "yyyy-mm-dd")
    dtpEnd.Value = Format(dtpEnd.MaxDate, "yyyy-mm-dd")
    txtInvoice.Locked = Not zlStr.IsHavePrivs(mstrPrivs, "修改票据号") And gblnStrictCtrl '89302
    Call InitPanel
    Call zlCardSquareObject
    Call zlCreateObject
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
 


Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With picDown
        .Top = ScaleHeight - stbThis.Height - .Height
        .Width = ScaleWidth
        .Left = ScaleLeft
    End With
End Sub
Private Sub UnloadWinSaveInfor()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:关闭窗口时,保存相关信息
    '编制:刘兴洪
    '日期:2013-01-05 15:27:20
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Call zlCardSquareObject(True)
    zlSaveDockPanceToReg Me, dkpMan, "区域"
    zl_vsGrid_Para_Save mlngModule, vsBalance, Me.Caption, "结算列表", False
    zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "单据列表", False
    zl_vsGrid_Para_Save mlngModule, vsDetail, Me.Caption, "单据明细列表", False
    Call zlCloseObject
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    '保存窗体个性化信息
    Call UnloadWinSaveInfor
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng卡类别ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    
    On Error GoTo errHandle
    
    If txtPatient.Locked Then Exit Sub
    If objCard.名称 Like "IC卡*" And objCard.系统 Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If mobjICCard Is Nothing Then Exit Sub
        txtPatient.Text = mobjICCard.Read_Card()
        If txtPatient.Text <> "" Then Call FindPati(objCard, False, txtPatient.Text)
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
    If mSquareCard.objSquareCard.zlReadCard(Me, mlngModule, lng卡类别ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    If txtPatient.Text <> "" Then Call FindPati(objCard, False, txtPatient.Text)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
 

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub
Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If txtPatient.Locked Then Exit Sub
    txtPatient.Text = objPatiInfor.卡号
    If txtPatient.Text <> "" Then Call FindPati(objCard, False, txtPatient.Text)
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNo As String)
    If txtPatient.Locked Or txtPatient.Text <> "" Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("IC卡", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    
    txtPatient.Text = strCardNo
    If txtPatient.Text <> "" Then Call FindPati(objCard, False, txtPatient.Text)
End Sub

Private Sub picCon_Resize()
    Err = 0: On Error Resume Next
    With picCon
        cmdBrush.Left = .ScaleWidth - cmdBrush.Width - 50
    End With
End Sub
Private Sub picDown_Resize()
    With picDown
        cmdCancel.Left = .ScaleLeft + .ScaleWidth - cmdCancel.Width - 50
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 50
    End With
End Sub

Private Sub picList_Resize()
    Err = 0: On Error Resume Next
    With picList
        vsList.Left = .ScaleLeft
        vsList.Width = .ScaleWidth
        vsList.Height = .ScaleHeight
        vsList.Top = .ScaleTop
    End With
End Sub
Private Sub picDetail_Resize()
    Err = 0: On Error Resume Next
    With PicDetail
        vsDetail.Left = .ScaleLeft
        vsDetail.Width = .ScaleWidth
        vsDetail.Height = .ScaleHeight
        vsDetail.Top = .ScaleTop
    End With
End Sub

Private Sub picBalance_Resize()
    Err = 0: On Error Resume Next
    With picBalance
        vsBalance.Left = .ScaleLeft
        vsBalance.Width = .ScaleWidth
        vsBalance.Height = .ScaleHeight - lblSum.Height - 50
        vsBalance.Top = .ScaleTop
        lblSum.Top = .ScaleHeight - lblSum.Height - 10
        lblSum.Left = .ScaleLeft
    End With
End Sub
 Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthday As Date, ByVal strAddress As String)
                            
    On Error GoTo errHandle
    
    If txtPatient.Locked Or txtPatient.Text <> "" Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("身份证", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = strID
    If txtPatient.Text <> "" Then Call FindPati(objCard, False, txtPatient.Text)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub zlCardSquareObject(Optional blnClosed As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建或关闭结算卡对象
    '入参:blnClosed:关闭对象
    '编制:刘兴洪
    '日期:2010-01-05 14:51:23
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    
    On Error GoTo errHandle
    
    If mSquareCard Is Nothing Then
         Set mSquareCard = New SquareCard
    End If
    '只有:执行或退费时,才可能管结算卡的
    If blnClosed Then
       If Not mSquareCard.objSquareCard Is Nothing Then
            Call mSquareCard.objSquareCard.CloseWindows
            Set mSquareCard.objSquareCard = Nothing
        End If
        Set mSquareCard = Nothing
        Exit Sub
    End If
    
    '创建对象
    '刘兴洪:增加结算卡的结算:执行或退费时
    Err = 0: On Error Resume Next
    Set mSquareCard.objSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    If Err <> 0 Then
        Err = 0: On Error GoTo 0:      Exit Sub
    End If
    
    If mSquareCard.objSquareCard Is Nothing Then Exit Sub
    Dim objCard As Card
    
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
       
   '安装了结算卡的部件
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   '功能:zlInitComponents (初始化接口部件)
   '    ByVal frmMain As Object, _
   '        ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
   '        ByVal cnOracle As ADODB.Connection, _
   '        Optional blnDeviceSet As Boolean = False, _
   '        Optional strExpand As String
   '出参:
   '返回:   True:调用成功,False:调用失败
   '编制:刘兴洪
   '日期:2009-12-15 15:16:22
   'HIS调用说明.
   '   1.进入门诊收费时调用本接口
   '   2.进入住院结帐时调用本接口
   '   3.进入预交款时
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '初始部件不成功,则作为不存在处理
   If mSquareCard.objSquareCard.zlInitComponents(Me, mlngModule, glngSys, gstrDBUser, gcnOracle, False, strExpend) = False Then Exit Sub
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, ByVal blnCard As Boolean, Optional blnOutMsg As Boolean = False) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息
    '入参:blnCard=是否就诊卡刷卡
    '出参:blnOutMsg-已经提示,不用再外部再提示
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-01-05 15:30:46
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, lng卡类别ID As Long, strPassWord As String, strErrMsg As String
    Dim lng病人ID As Long, blnHavePassWord As Boolean
    Dim strIF As String, strTemp As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errH
    mstrFindNO = "": mstrFindFpNo = ""
    mintPrintNums = 0
    strSQL = _
        "Select A.病人ID,Nvl(B.主页ID,0) as 主页ID,A.门诊号 as 门诊号,A.当前床号,B.出院病床," & _
        "       Nvl(b.姓名, a.姓名) As 姓名, Nvl(b.性别, a.性别) As 性别,Nvl(B.年龄,A.年龄) as 年龄,A.IC卡号,A.就诊卡号,A.卡验证码," & _
        "       Nvl(B.费别,A.费别) as 费别,C.名称 as 当前科室,A.当前科室ID,D.名称 as 出院科室,B.出院科室ID, A.险类 as 险类,E.卡号,E.医保号,E.密码," & _
        "       A.登记时间,Nvl(B.状态,0) as 状态,Nvl(B.医疗付款方式,A.医疗付款方式) as 医疗付款方式,Nvl(B.审核标志,0) as 审核标志,B.入院日期,B.出院日期,B.病人性质,B.病人类型" & _
        " From 病人信息 A,病案主页 B,部门表 C,部门表 D,医保病人档案 E,医保病人关联表 F" & _
        " Where A.停用时间 is NULL And A.病人ID=B.病人ID(+) And A.主页ID=B.主页ID(+) " & _
        "           And A.病人ID=F.病人ID(+) And F.标志(+)=1 And F.医保号=E.医保号(+) And F.险类=E.险类(+) And F.中心 = E.中心(+)" & _
        "           And A.当前科室ID=C.ID(+) And B.出院科室ID=D.ID(+)" & _
        "           And A.停用时间 is NULL "
    
    If blnCard = True And objCard.名称 Like "姓名*" Then    '刷卡
        lng卡类别ID = IDKind.GetDefaultCardTypeID
        '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户);…
        If mSquareCard.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg, lng卡类别ID) = False Then GoTo NotFoundPati:
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
    ElseIf Left(strInput, 1) = "." Or objCard.名称 = "单据号" Then
        '单据号查找
        If Left(strInput, 1) = "." Then
            strTemp = UCase(GetFullNO(Mid(strInput, 2), 15))
        Else
            strTemp = UCase(GetFullNO(strInput, 15))
        End If
        
        txtPatient.Text = strTemp
        gstrSQL = "" & _
        "   Select  distinct A.病人ID " & _
        "   From 病人结帐记录A " & _
        "   Where A.NO=[1]  And Rownum=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strTemp, 1)
        If rsTemp.EOF Then
            MsgBox "注意:" & vbCrLf & "  单据号为『" & strTemp & "』不存在,请检查输入的单据是否正确!", vbInformation + vbOKOnly, gstrSysName
            Call zlClearPatiInfor
            Exit Function
        End If
        If Val(Nvl(rsTemp!病人ID)) = 0 Then
            MsgBox "该结账单据是合约单位结账!", vbInformation, gstrSysName
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
            Case "姓名", "姓名或就诊卡"
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
            Case Else
                '其他类别的,获取相关的病人ID
                If objCard.接口序号 > 0 Then
                    lng卡类别ID = objCard.接口序号
                    If mSquareCard.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    If lng病人ID = 0 Then GoTo NotFoundPati:
                Else
                    If mSquareCard.objSquareCard.zlGetPatiID(objCard.名称, strInput, False, lng病人ID, _
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
        txtPatient.Text = Nvl(mrsInfo!姓名)
        'txtOld.Text = Nvl(mrsInfo!年龄): txtSex.Text = Nvl(mrsInfo!性别)
        ' txt住院号.Text = Nvl(mrsInfo!门诊号)
        txtPatient.PasswordChar = ""
        '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
        txtPatient.IMEMode = 0
        mstrPassWord = strPassWord
        If Not blnHavePassWord Then mstrPassWord = Nvl(mrsInfo!卡验证码)
        GetPatient = True
        Exit Function
    Else
        Call zlClearPatiInfor
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
NotFoundPati:
  Call zlClearPatiInfor
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
End Function

Private Sub zlClearPatiInfor()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检除病人信息
    '编制:刘兴洪
    '日期:2013-01-05 15:31:23
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    txtPatient.Text = ""
    Set mrsInfo = New ADODB.Recordset
    vsList.Clear 1: vsList.Rows = 2: vsDetail.Clear 1: vsDetail.Rows = 2
    vsBalance.Clear 1: vsBalance.Rows = 2
    txtInvoice.Text = ""
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub txtPatient_Change()
    txtPatient.Tag = ""
    cmdBrush.Enabled = False
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    IDKind.SetAutoReadCard (txtPatient.Text <> "" And Me.ActiveControl Is txtPatient)
    stbThis.Panels(2).Text = ""
End Sub

Private Sub txtPatient_GotFocus()
    zlControl.TxtSelAll txtPatient
     If txtPatient.Locked Then Exit Sub
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtPatient.Text = "")
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (txtPatient.Text = "")
    IDKind.SetAutoReadCard (txtPatient.Text = "")
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
    '问题:51488
    If (IDKind.Cards.读卡快键 = "空格键" Or IDKind.Cards.读卡快键 = " ") And Chr(KeyAscii) = " " Then KeyAscii = 0: Exit Sub
        
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
                Call zlClearPatiInfor
                Exit Sub
            End If
            Call ReadListData
            Exit Sub
        End If
        KeyAscii = 0
        Call FindPati(IDKind.GetCurCard, blnCard, Trim(txtPatient.Text))
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
    '日期:2013-01-05 15:32:23
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long, lngUnit As Long, i As Long
    Dim strSQL As String, curTotal As Currency, blnIDCard As Boolean
    Dim blnICCard As Boolean, blnMsg As Boolean
    
    On Error GoTo errHandle
    
    If objCard.名称 Like "IC卡*" And objCard.系统 And mstrPassWord <> "" Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    If objCard.名称 Like "*身份证*" And objCard.系统 And mstrPassWord <> "" Then blnIDCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    If Not GetPatient(objCard, strInput, blnCard, blnMsg) Then
        cmdBrush.Enabled = False
        If blnCard Then
            If Not blnMsg Then MsgBox "不能确定病人信息，请检查是否正确刷卡！", vbInformation, gstrSysName
            Call zlClearPatiInfor
            Exit Sub
        End If
        If Not blnMsg Then MsgBox "不能读取病人信息！", vbInformation, gstrSysName
        Call zlClearPatiInfor
        Exit Sub
    End If
    '读取成功
    '就诊卡密码检查
    If (objCard.名称 Like "IC卡*" Or objCard.名称 Like "*身份证*") And objCard.系统 = True And blnCard Then blnCard = False
    If Mid(gstrCardPass, 6, 1) = "1" And (blnCard Or blnICCard Or blnIDCard) Then
        If Not zlCommFun.VerifyPassWord(Me, mstrPassWord, mrsInfo!姓名, mrsInfo!性别, "" & mrsInfo!年龄) Then
            Call zlClearPatiInfor
             txtPatient.SetFocus: Exit Sub
        End If
    End If
    cmdBrush.Enabled = True
    Call ReadListData
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub txtPatient_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
End Sub

Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtPatient.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub
Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub
Private Sub txtPatient_Validate(Cancel As Boolean)
    If IsNumeric(txtPatient.Tag) And mrsInfo.State = 1 Then
        mblnValid = True
        Call txtPatient_KeyPress(13)
        mblnValid = False
    End If
End Sub
  
Private Sub ShowDetail(ByVal strNO As String, Optional lng结帐ID As Long)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示明细数据
    '入参:strNO-结账单据号
    '       lng结帐ID-结帐ID
    '编制:刘兴洪
    '日期:2013-01-05 15:33:04
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim i As Long, j As Long, strSQL As String
    Dim lngCol As Long
    
    On Error GoTo errH
    
    '是否已转入后备数据表中
    mblnNOMoved = zlDatabase.NOMoved("病人结帐记录", strNO, , , Me.Caption)
    strSQL = "" & _
    "   Select  '门诊' as 住院,A.发生时间,A.NO,A.序号,A.收费细目ID,A.收据费目,A.婴儿费,A.结帐金额,A.开单部门ID " & _
    "   From " & IIf(mblnNOMoved, "H", "") & "门诊费用记录 A" & _
    "   Where A.结帐ID=[1]" & _
    "    Union ALL " & _
    "   Select  Decode(A.主页ID,NULL,'门诊','第'||A.主页ID||'次') as 住院,A.发生时间,A.NO,A.序号,A.收费细目ID,A.收据费目,A.婴儿费,A.结帐金额,A.开单部门ID " & _
    "   From " & IIf(mblnNOMoved, "H", "") & "住院费用记录 A" & _
    "   Where A.结帐ID=[1] " & _
    "   "
    strSQL = _
    "  Select   A.住院," & _
    "            Nvl(B.名称,'未知') as 科室,To_Char(A.发生时间,'YYYY-MM-DD') as 时间," & _
    "            A.NO as 单据号,Nvl(E.名称,D.名称) as 项目,A.收据费目 as 费目," & _
    "            Decode(Nvl(A.婴儿费,0),0,'','√') as 婴儿费,A.结帐金额" & _
    " From (" & strSQL & ") A,部门表 B,收费项目目录 D,收费项目别名 E" & _
    " Where A.开单部门ID=B.ID(+) And A.收费细目ID=D.ID" & _
    "           And A.收费细目ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
    " Order by 住院 Desc,时间 Desc,单据号 Desc,序号"
    Set mrsDetail = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng结帐ID)
    If mrsDetail.EOF Then Exit Sub
 
    With vsDetail
        .Clear 1
        .Redraw = flexRDNone
        Set .DataSource = mrsDetail
        For lngCol = 0 To .Cols - 1
             .ColAlignment(lngCol) = flexAlignLeftCenter
             .FixedAlignment(lngCol) = flexAlignCenterCenter
              .ColKey(lngCol) = Trim(.TextMatrix(0, lngCol))
              If .ColKey(lngCol) Like "*ID" Or InStr(1, ",符号", "," & .ColKey(lngCol) & ",") > 0 Then
                    .ColHidden(lngCol) = True
              ElseIf .ColKey(lngCol) Like "*数*" Or .ColKey(lngCol) Like "*价*" Or .ColKey(lngCol) Like "*额" Then
                    .ColAlignment(lngCol) = flexAlignRightCenter
              End If
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .Cols - 1)
        zl_vsGrid_Para_Restore mlngModule, vsDetail, Me.Caption, "单据明细列表", False
        .Redraw = flexRDBuffered
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Function InitBlanceData(ByVal str结帐IDs As String) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化结算数据
    '入参:str结帐IDs-结帐ID(多个用逗号分离)
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-01-05 15:33:40
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
   Dim strSQL As String
    Err = 0: On Error GoTo errHandle
    
    If str结帐IDs = "" Then InitBlanceData = True: Exit Function
    
    strSQL = _
    " Select   M.NO,M.ID as 结帐ID" & _
    " From 病人结帐记录 M , Table(f_num2list([1]))  J" & _
    " Where  M.ID=J.Column_Value"
    
    strSQL = _
    " Select /*+ rule */ A.NO,A.结帐ID, A.结算方式,Nvl(B.性质,1) as 性质,B.应付款,A.金额,A.摘要,A.结算号码" & _
    " From (  Select B.NO ,A.结帐ID,Decode(A.记录性质,2,A.结算方式,12,A.结算方式,NULL) as 结算方式,A.摘要,A.结算号码," & _
    "               Sum(A.冲预交) as 金额" & _
    "         From 病人预交记录 A, (" & strSQL & ")  B" & _
    "         Where A.结帐ID=B.结帐ID And A.记录性质 IN(1,11,2,12) And Nvl(A.冲预交,0)<>0" & _
    "         Group by B.NO,A.结帐ID, Decode(A.记录性质,2,A.结算方式,12,A.结算方式,NULL),A.摘要,A.结算号码" & _
    "       ) A,结算方式 B " & _
    " Where A.结算方式=B.名称(+) " & _
    " "
    
    Set mrsBalance = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Replace(str结帐IDs, "'", ""))
    InitBlanceData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub SetBlanceShow()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示结算方式
    '入参:blnAllSel-选择所有的单据
    '编制:刘兴洪
    '日期:2013-01-05 15:35:36
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strBalance As String, lngRow As Long, i As Long, str结算 As String
    Dim bln全选 As Boolean, bln未选 As Boolean
    Dim strFilter As String, bln退款 As Boolean
    Dim lng结帐ID As Long, str结帐IDs As String
    Dim dblMoney As Double
    
    On Error GoTo errHandle
    
    lblSum.Caption = "补打合计:" & Format(0, "0.00")
    If mrsBalance Is Nothing Then Exit Sub
    
    With vsList
        bln全选 = True: bln未选 = True
        For lngRow = 1 To .Rows - 1
            lng结帐ID = Val(.TextMatrix(lngRow, .ColIndex("结帐ID")))
            
            If GetVsGridBoolColVal(vsList, lngRow, .ColIndex("选择")) Then
                If InStr(1, str结帐IDs & ",", "," & lng结帐ID & ",") = 0 Then
                    str结帐IDs = str结帐IDs & "," & lng结帐ID
                    bln未选 = False
                End If
            End If
             If InStr(1, str结帐IDs & ",", "," & lng结帐ID & ",") = 0 Then bln全选 = False
        Next
    End With
    If str结帐IDs <> "" Then str结帐IDs = Mid(str结帐IDs, 2)
    bln退款 = False
    '显示所有选择的单据的结算方式之和
    If bln全选 Or bln未选 Then
        mrsBalance.Filter = 0
        If bln全选 Then bln退款 = True
    Else
        strFilter = Replace(str结帐IDs, ",", " Or 结帐ID=")
        strFilter = " 结帐ID=" & strFilter & ""
        mrsBalance.Filter = strFilter
        bln退款 = True
    End If
    
    mrsBalance.Sort = "NO,结算方式"
    With vsBalance
         .Redraw = flexRDNone
        .Rows = IIf(mrsBalance.RecordCount = 0, 1, mrsBalance.RecordCount) + 1
        i = 1
        dblMoney = 0
        Do While Not mrsBalance.EOF
            .TextMatrix(i, .ColIndex("NO")) = Nvl(mrsBalance!NO)
            .TextMatrix(i, .ColIndex("结算方式")) = Nvl(mrsBalance!结算方式)
            .TextMatrix(i, .ColIndex("结算金额")) = Format(Val(Nvl(mrsBalance!金额)), "0.00")
            dblMoney = dblMoney + Val(Nvl(mrsBalance!金额))
            i = i + 1
            mrsBalance.MoveNext
        Loop
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .Cols - 1)
        zl_vsGrid_Para_Restore mlngModule, vsBalance, Me.Caption, "结算列表", False
        .Redraw = flexRDBuffered
        If bln未选 Then
            lblSum.Caption = "未打合计:" & Format(dblMoney, "0.00")
        Else
            lblSum.Caption = "补打合计:" & Format(dblMoney, "0.00")
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function ReadListData() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取需要明细数据
    '返回:读取成功,返回true,否则返回Flase
    '编制:刘兴洪
    '日期:2013-01-05 15:36:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng病人ID As Long, strTable As String, rsTemp As ADODB.Recordset
    Dim lngCol As Long, strSQL As String, lngRow As Long
    Dim strFilter As String, str结帐IDs As String
    Dim strWhere As String, strTable1 As String, dtStartDate As Date, dtEndDate As Date
    Dim strNO As String, i As Long, lng结帐序号 As Long
    Dim lng主页ID  As Long
    
    On Error GoTo errHandle
    
    dtStartDate = CDate("1901-01-01")
    dtEndDate = dtStartDate
    lng主页ID = 0
    If mrsInfo Is Nothing Then
        lng病人ID = 0
    ElseIf mrsInfo.State <> 1 Then
        lng病人ID = 0
    Else
        lng病人ID = Val(Nvl(mrsInfo!病人ID))
        lng主页ID = Val(Nvl(mrsInfo!主页ID))
    End If
    If mstrFindNO <> "" Then
        strWhere = "  And A.NO=[2]"
    Else
        strTable1 = ""
        strWhere = "  And A.病人ID=[1]"
    End If
    
    If chkDate.Value = 0 Then
        strWhere = strWhere & " And A.收费时间 betWeen [3] and [4]"
        dtStartDate = CDate(Format(dtpBegin.Value, "yyyy-mm-dd") & " 00:00:00")
        dtEndDate = CDate(Format(dtpEnd.Value, "yyyy-mm-dd") & " 23:59:59")
    End If
    mblnSel = False
    zlCommFun.ShowFlash "正在读取单据数据,请稍候 ..."
    Screen.MousePointer = 11
    DoEvents
    strSQL = "" & _
    "   Select  -1 as  选择 ,  max(Decode(Y.名称,NULL,NULL,'√')) as 医保,a.Id as 结帐ID, a.No as 单据号, a.实际票号, a.病人id, Max(B.主页ID) as 主页ID, a.操作员编号, a.操作员姓名, a.备注, a.原因, To_Char(a.收费时间, 'yyyy-mm-dd hh24:mi:ss') As 结帐时间, " & _
    "          Decode(a.结帐类型, 1, '门诊结帐', '住院结帐') As 结帐类型, To_Char(Sum(b.实收金额), '99999990.00') As 实际金额, " & _
    "          nvl(Max(X.险类),0) as 险类ID " & _
    "   From 病人结帐记录 A, 住院费用记录 B,保险结算记录 X,保险类别 Y " & _
    "   Where   a.Id = b.结帐id And a.实际票号 Is Null And a.记录状态 = 1 " & strWhere & _
    "         And A.id=X.记录ID(+) And X.性质(+)=2 And X.险类=Y.序号(+) And Nvl(X.序号(+),1)=1 " & _
    "   Group By a.Id, a.No, a.实际票号, a.病人id , a.操作员编号, a.操作员姓名, a.备注, a.原因, " & _
    "           To_Char(a.收费时间, 'yyyy-mm-dd hh24:mi:ss'),  Decode(a.结帐类型, 1, '门诊结帐', '住院结帐') " & _
    "   order by 结帐时间,单据号"
     
    Set mrsList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, mstrFindNO, dtStartDate, dtEndDate)
    vsList.Redraw = flexRDNone
    vsList.Clear: vsList.Cols = 2
    Set vsList.DataSource = mrsList
    If vsList.Rows <= 1 Then vsList.Rows = 2
    With vsList
        mstrNOs = ""
        For lngCol = 0 To .Cols - 1
             .ColAlignment(lngCol) = flexAlignLeftCenter
             .FixedAlignment(lngCol) = flexAlignCenterCenter
              .ColKey(lngCol) = Trim(.TextMatrix(0, lngCol))
              If .ColKey(lngCol) Like "*ID" Or InStr(1, ",符号,", "," & .ColKey(lngCol) & ",") > 0 Then
                    .ColHidden(lngCol) = True
              ElseIf .ColKey(lngCol) Like "*数*" Or .ColKey(lngCol) Like "*价*" Or .ColKey(lngCol) Like "*额" Then
                    .ColAlignment(lngCol) = flexAlignRightCenter
              End If
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        .ColDataType(.ColIndex("选择")) = flexDTBoolean
        Call .AutoSize(0, .Cols - 1)
        zl_vsGrid_Para_Restore mlngModule, vsList, Me.Caption, "单据列表", False
        .Editable = flexEDKbdMouse

        .Redraw = flexRDBuffered
        vsList_AfterRowColChange 0, 0, .Row, .Col
    
    End With
    mrsList.Filter = "险类ID>0"
    If Not mrsList.EOF Then
        mintInsure = Val(Nvl(mrsList!险类ID))
    End If
    mrsList.Filter = 0
    If mrsList.RecordCount <> 0 Then mrsList.MoveFirst
    str结帐IDs = ""
    With mrsList
        Do While Not .EOF
            str结帐IDs = str结帐IDs & "," & Val(Nvl(!结帐ID))
            .MoveNext
        Loop
        str结帐IDs = "-1" & str结帐IDs
        If str结帐IDs <> "" Then str结帐IDs = Mid(str结帐IDs, 2)
    End With
    '加载结算方式
    Call InitBlanceData(str结帐IDs)
    Call SetBlanceShow
    '处理发票相关信息
    Call ReInitPatiInvoice(lng病人ID, lng主页ID, mintInsure)
    Call RefreshFact    '重新处理发票
    zlCommFun.StopFlash
    Screen.MousePointer = 0
    ReadListData = True
    Exit Function
errHandle:
    vsList.Redraw = flexRDBuffered
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
   zlCommFun.StopFlash
End Function
Private Sub vsBalance_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsBalance, Me.Caption, "结算列表", False
End Sub

Private Sub vsBalance_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsBalance, Me.Caption, "结算列表", False
End Sub

Private Sub vsDetail_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsDetail, Me.Caption, "单据明细列表", False
End Sub

Private Sub vsDetail_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsDetail, Me.Caption, "单据明细列表", False
End Sub
Private Sub vsList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call SetBlanceShow
End Sub

Private Sub vsList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "单据列表", False
End Sub

Private Sub vsList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strNO As String, lng结帐ID As Long
    
    On Error GoTo errHandle
    
    If OldRow = NewRow Then Exit Sub
    With vsList
        If NewRow < 0 Then Exit Sub
        strNO = Trim(.TextMatrix(NewRow, .ColIndex("单据号")))
        lng结帐ID = Val(.TextMatrix(NewRow, .ColIndex("结帐ID")))
    End With
    ShowDetail strNO, lng结帐ID
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "单据列表", False
End Sub
 
Private Sub vsList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsList
        Select Case Col
        Case .ColIndex("选择")
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub ReInitPatiInvoice(ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
     Optional ByVal intInsure As Integer = 0)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新初始化病人发票信息
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-01-05 15:38:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bytInvoiceKind     As Byte, intFormat As Integer
    Dim intPrintMode As Integer, lngShareUseID As Long
    On Error GoTo errHandle

    bytInvoiceKind = Val(zlDatabase.GetPara("住院结帐票据类型", glngSys, 1137, "0"))

    Set mobjInvoice = New clsInvoice: Set mobjFact = New clsFactProperty
    mobjInvoice.zlInitCommon glngSys, gcnOracle, gstrDBUser
    
    Call mobjInvoice.zlGetInvoicePreperty(1137, IIf(bytInvoiceKind = 0, 3, 1), 0, 0, 0, mobjFact, , , 2)
    
    mobjFact.使用类别 = zlDatabase.GetPara("合约单位结帐打印", glngSys, 1137)
    mobjFact.票种 = IIf(bytInvoiceKind = 0, 3, 1)
    Call mobjInvoice.zlGetInvoicePrintFormat(1137, mobjFact.票种, mobjFact.使用类别, intFormat, 2)
    mobjFact.打印格式 = intFormat
    If mobjInvoice.zlGetInvoicePrintMode(1137, mobjFact.票种, mobjFact.使用类别, intPrintMode) = False Then Exit Sub
    mobjFact.打印方式 = intPrintMode
    If mobjInvoice.zlGetInvoiceShareID(1137, mobjFact.票种, mobjFact.使用类别, lngShareUseID) = False Then Exit Sub
    mobjFact.共享批次ID = lngShareUseID
    
    Call ZlShowBillFormat(bytInvoiceKind, lblFormat, mobjFact.打印格式)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub RefreshFact()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:刷新收费票据号
    '编制:刘兴洪
    '日期:2013-01-05 15:39:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
       Dim strFactNO As String
       
    On Error GoTo errHandle
        
    If mobjFact.打印方式 = 0 Then Exit Sub
    If Not mobjFact.严格控制 Then
        '非严格控制下
        '松散：取下一个号码
        txtInvoice.Text = zlCommFun.IncStr(UCase(zlDatabase.GetPara("当前结帐票据号", glngSys, 1137, "")))
        txtInvoice.Tag = txtInvoice.Text
        txtInvoice.SelStart = Len(txtInvoice.Text)
        Exit Sub
    End If
    If zlGetInvoiceGroupUseID(mlng领用ID, 1, "") = False Then
          txtInvoice.Text = "": txtInvoice.Tag = ""
        Exit Sub
    End If
    
    '严格：取下一个号码
    If mobjInvoice.zlGetNextBill(1137, mlng领用ID, strFactNO) = False Then strFactNO = ""
    txtInvoice.Text = strFactNO
    'Tag：问题：24363:刘兴洪：主要是解决自动生成的号是否被用户更改，主要解决：
    '    1.更改的票据号需要检查是否重复，重复后直接返回不更改发票号
    '    2.并发操作，不更改的情况下，检查是否重复，如果重复，自动取下一个号码！
    txtInvoice.Tag = txtInvoice.Text
    lblFact.Tag = txtInvoice.Tag
    txtInvoice.SelStart = Len(txtInvoice.Text)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
     
End Sub

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
    
    If mobjInvoice.zlGetInvoiceGroupID(1137, UserInfo.姓名, mobjFact.票种, _
        mobjFact.使用类别, lng领用ID, mobjFact.共享批次ID, lng领用ID, intNum, strInvoiceNO) = False Then Exit Function
    
    If lng领用ID > 0 Then zlGetInvoiceGroupUseID = True: Exit Function
    
    Select Case lng领用ID
        Case 0 '操作失败
        Case -1
            If Trim(mobjFact.使用类别) = "" Then
                MsgBox "你没有自用和共用的结帐票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
            Else
                MsgBox "你没有自用和共用的『" & mobjFact.使用类别 & "』结帐票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
            End If
            Exit Function
        Case -2
            If Trim(mobjFact.使用类别) = "" Then
                MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
            Else
                MsgBox "本地的共用票据的『" & mobjFact.使用类别 & "』结帐票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
            End If
            Exit Function
        Case -3
            MsgBox "当前票据号码不在可用领用批次的有效票据号范围内,请重新输入！", vbInformation, gstrSysName
            If txtInvoice.Enabled And txtInvoice.Visible Then txtInvoice.SetFocus
            Exit Function
    End Select
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub zlCreateObject()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建公共事件对象
    '编制:刘兴洪
    '日期:2013-01-05 15:41:19
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
Private Function CheckBillRepeat(lng领用ID As Long, byt票种 As Byte, strNO As String) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:在使用新票号之前，检查是否重复
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-01-05 15:42:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 号码 From 票据使用明细" & _
        " Where 领用ID=[1] And 票种=[2] And 号码=[3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lng领用ID, byt票种, strNO)
    CheckBillRepeat = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

