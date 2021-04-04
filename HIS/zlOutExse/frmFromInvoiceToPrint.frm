VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFromInvoiceToPrint 
   Caption         =   "根据发票号重打发票"
   ClientHeight    =   8325
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11865
   Icon            =   "frmFromInvoiceToPrint.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmFromInvoiceToPrint.frx":058A
   ScaleHeight     =   8325
   ScaleWidth      =   11865
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picDown 
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   45
      ScaleHeight     =   660
      ScaleWidth      =   11700
      TabIndex        =   14
      Top             =   7170
      Width           =   11700
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
         TabIndex        =   17
         Top             =   225
         Width           =   1440
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "重打(&O)"
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
         Left            =   8565
         TabIndex        =   16
         Top             =   195
         Width           =   1440
      End
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
         Left            =   765
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   180
         Width           =   2175
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
         Left            =   270
         TabIndex        =   19
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lblFormat 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C000C0&
         Height          =   240
         Left            =   3195
         TabIndex        =   18
         Top             =   285
         Visible         =   0   'False
         Width           =   120
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
      Begin VB.CommandButton cmdBrush 
         Caption         =   "刷新(&N)"
         Height          =   375
         Left            =   9195
         TabIndex        =   8
         Top             =   150
         Width           =   1245
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
         MaxLength       =   64
         TabIndex        =   7
         Top             =   150
         Width           =   2040
      End
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   375
         Left            =   615
         TabIndex        =   6
         Top             =   150
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   661
         Appearance      =   2
         IDKindStr       =   $"frmFromInvoiceToPrint.frx":0B14
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
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   345
         Left            =   5505
         TabIndex        =   9
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
         Format          =   146800643
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
         Format          =   146800643
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
         TabIndex        =   11
         Top             =   210
         Value           =   1  'Checked
         Width           =   1890
      End
      Begin VB.Label lbl至 
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
         TabIndex        =   13
         Top             =   225
         Width           =   240
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
         Left            =   150
         TabIndex        =   12
         Top             =   195
         Width           =   480
      End
   End
   Begin VB.PictureBox picList 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   5025
      ScaleHeight     =   2295
      ScaleWidth      =   5400
      TabIndex        =   3
      Top             =   2280
      Width           =   5400
      Begin VSFlex8Ctl.VSFlexGrid vsList 
         Height          =   2685
         Left            =   105
         TabIndex        =   4
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
         FormatString    =   $"frmFromInvoiceToPrint.frx":0BC7
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
   Begin VB.PictureBox PicDetail 
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   5250
      ScaleHeight     =   2775
      ScaleWidth      =   5535
      TabIndex        =   1
      Top             =   3840
      Width           =   5535
      Begin VSFlex8Ctl.VSFlexGrid vsDetail 
         Height          =   2685
         Left            =   180
         TabIndex        =   2
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
         FormatString    =   $"frmFromInvoiceToPrint.frx":0BDD
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
   Begin VB.PictureBox picInvoice 
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   150
      ScaleHeight     =   4695
      ScaleWidth      =   4380
      TabIndex        =   0
      Top             =   1065
      Width           =   4380
      Begin VB.TextBox txtFilterInvoice 
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
         Left            =   900
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   0
         Width           =   2175
      End
      Begin VSFlex8Ctl.VSFlexGrid vsInvoiceList 
         Height          =   2685
         Left            =   300
         TabIndex        =   23
         Top             =   660
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
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmFromInvoiceToPrint.frx":0BF3
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
      Begin VB.Label lblInVoice 
         AutoSize        =   -1  'True
         Caption         =   "发票号"
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
         Left            =   255
         TabIndex        =   21
         Top             =   60
         Width           =   630
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   20
      Top             =   7965
      Width           =   11865
      _ExtentX        =   20929
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
            Picture         =   "frmFromInvoiceToPrint.frx":0C96
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15849
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
            Picture         =   "frmFromInvoiceToPrint.frx":152A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFromInvoiceToPrint.frx":187E
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
Attribute VB_Name = "frmFromInvoiceToPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum mPanel
    Pane_Search = 1
    Pane_List = 2
    Pane_Detail = 3
    Pane_inVoiceList = 4
End Enum
'-----------------------------------------------------------------------------------
'结算卡相关
Private mSquareCard As SquareCard '结算卡相关
Private mstrPassWord As String
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
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private mblnValid As Boolean
Private mblnSel As Boolean
Private mstrPrivs As String
Private mintSucces As Integer  '成功打印次数
Private mlng病人ID As Long
Private mblnFirst As Boolean
Private mblnNotClick As Boolean
Private mlngShareUseID As Long '共享领用批次ID
Private mstrUseType As String '使用类别
Private mintInvoiceFormat As Integer  '打印的发票格式,发票格式序号
Private mintOldInvoiceFormat As Integer '旧发票打印的格式
Private mblnStartFactUseType As Boolean   '是否启用了使用类别
Private mintInvoicePrint As Integer  '0-不打印;1-自动打印;2-提示打印
Private mlng领用ID As Long
Private mintInsure As Integer
Private mrsInVoice As ADODB.Recordset
Private mblnNotChange As Boolean

Public Function zlRePrintBill(ByVal frmMain As Object, ByVal lngModule As Long, _
    ByVal strPrivs As String, Optional lng病人ID As Long = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重打票据入口
    '返回:打印成功,返回true,否则返回Flase
    '编制:刘兴洪
    '日期:2011-09-04 22:39:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs: mintSucces = 0
    mlng病人ID = lng病人ID
    If frmMain Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
    zlRePrintBill = mintSucces > 0
End Function
Private Function InitPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化条件区哉
    '编制:刘兴洪
    '日期:2009-09-09 15:04:30
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
        
        Set objTemp = .CreatePane(mPanel.Pane_inVoiceList, 100, 100, DockBottomOf, objPane)
        objTemp.Tag = mPanel.Pane_inVoiceList
        objTemp.Title = "发票信息列表": objTemp.Options = PaneNoCloseable Or PaneNoHideable
        objTemp.Handle = picInvoice.hWnd
        
        Set objPane = .CreatePane(mPanel.Pane_List, 300, 100, DockRightOf, objTemp)
        objPane.Tag = mPanel.Pane_List
        objPane.Title = "收费单据列表": objPane.Options = PaneNoCloseable Or PaneNoHideable
        objPane.Handle = picList.hWnd
 
        Set objPane = .CreatePane(mPanel.Pane_Detail, 300, 100, DockBottomOf, objPane)
        objPane.Tag = mPanel.Pane_Detail
        objPane.Title = "单据明细列表": objPane.Options = PaneNoCloseable Or PaneNoHideable
        objPane.Handle = PicDetail.hWnd
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
    dkpMan.RecalcLayout: DoEvents
    zlRestoreDockPanceToReg Me, dkpMan, "区域"
    'Call GetRegInFor(g私有模块, Me.Name, "隐藏", strKey)
    'If Val(strKey) = 1 Then mPanSearch.Hide
        
End Function
Private Sub chkDate_Click()
    dtpBegin.Enabled = chkDate.Value <> 1
    dtpEnd.Enabled = chkDate.Value <> 1
End Sub
Private Sub cmdBrush_Click()
        mblnNotChange = True
        Call ReadInVoice(False)
        mblnNotChange = False
End Sub
Private Sub cmdCancel_Click()
     Unload Me
End Sub
 
Private Function zlMakeupPrint(ByVal strNos As String, ByVal strReclaimInvoice As String, _
    Optional ByVal blnMediCare As Boolean, _
    Optional ByVal blnDel As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:补打票据
    '入参:strReclaimInvoice-回收的票据号
    '返回:补打成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-09-04 22:58:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNo As String
    Dim intInsure As Integer, blnVirtualPrint As Boolean, lng结帐ID As Long, lng病人ID As Long
    Dim strUseType  As String, lngShareUseID As Long, intInvoiceFormat As Integer
    Dim intOldInvoiceFormat As Integer
    If strNos = "" Then Exit Function
    If mrsInfo Is Nothing Then Exit Function
    If mrsInfo.State <> 1 Then Exit Function
    
    lng病人ID = Val(Nvl(mrsInfo!病人ID))
    If lng病人ID = 0 Then Exit Function
    If strNos = "" Then
        MsgBox "当前没有单据可以重打票据！", vbInformation, gstrSysName
        Exit Function
    End If
    If blnMediCare Then
        intInsure = ChargeExistInsure(strNos, lng病人ID, lng结帐ID, , blnDel)
        If intInsure <> 0 Then
            blnVirtualPrint = gclsInsure.GetCapability(support医保接口打印票据, lng病人ID, intInsure, CStr(lng结帐ID))
        End If
    End If
    strUseType = zl_GetInvoiceUserType(lng病人ID, 0, intInsure)
    lngShareUseID = zl_GetInvoiceShareID(mlngModule, strUseType)
    intInvoiceFormat = zl_GetInvoicePrintFormat(mlngModule, strUseType, intOldInvoiceFormat)
    '单据有剩余数量的才可以重打，北京医保，即使退完了也可以重新打印
    If Not blnVirtualPrint Then
        If Not BillExistMoney(strNos, 1) Then
            MsgBox "单据中的项目已经全部退费，不能进行打印！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    '冉俊明,2014-12-17,补结算后的收费单据不允许重打票据
    If CheckBillExistReplenishData(1, , strNos) = True Then
        MsgBox "单据中的项目已经进行了保险补充结算，不能进行打印！", vbInformation, gstrSysName
        Exit Function
    End If
    
    Dim dtDate As Date, strPriceGrade As String
    dtDate = zlDatabase.Currentdate
    '获取价格等级
    If gintPriceGradeStartType >= 2 Then
        Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, lng病人ID, 0, "", , , strPriceGrade)
    Else
        strPriceGrade = gstr普通价格等级
    End If
    Call frmPrint.ReportPrint(2, strNos, "", strReclaimInvoice, mlng领用ID, mlngShareUseID, txtInvoice.Text, dtDate, "", "", _
        gTy_Module_Para.bln分别打印, mintInvoiceFormat, blnVirtualPrint, , mstrUseType, , , , strPriceGrade)
    
    '银医一卡通写卡，85950
    Call WriteInforToCard(Me, mlngModule, mstrPrivs, mSquareCard.objSquareCard, 0, strNos)
    
    mintSucces = mintSucces + 1
    zlMakeupPrint = True
End Function

Private Function CheckFP() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查发票是否正确
    '返回: 发票合法 返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-07-12 11:30:22
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intNum As Integer, varData As Variant
     On Error GoTo errHandle
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
    intNum = 1
    If Not gTy_Module_Para.bln分别打印 Then intNum = 1
 
InvoiceHandle:
    If zlCheckInvoiceValied(mlng领用ID, intNum, txtInvoice.Text, mlngShareUseID, mstrUseType) = False Then
        Exit Function
    End If
    '并发操作检查,票号是否已用
    If CheckBillRepeat(mlng领用ID, 1, txtInvoice.Text) Then
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
     Dim strReclaimInvoice As String '回收票据
     Dim strNos As String, strNo As String, blnYb As Boolean
     Dim lngRow As Long
     
     With vsInvoiceList
        strReclaimInvoice = .TextMatrix(.Row, .ColIndex("发票号"))
        If strReclaimInvoice = "" Then
            MsgBox "未选中需要重打的票据", vbInformation + vbOKOnly, gstrSysName
            Exit Sub
        End If
        If MsgBox("你确认要重打发票号为『" & strReclaimInvoice & "』的票据吗?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
     End With
     With vsList
        For lngRow = 1 To .Rows - 1
            strNo = Trim(.TextMatrix(lngRow, .ColIndex("单据号")))
            If strNo <> "" And InStr(strNos & ",", "," & strNo & ",") = 0 Then
               If blnYb = False Then blnYb = .TextMatrix(lngRow, .ColIndex("医保")) = "√"
                strNos = strNos & "," & strNo
            End If
        Next
     End With
    If Not CheckFP() Then Exit Sub
    Call zlMakeupPrint(strNos, strReclaimInvoice, blnYb, False)
    Call RefreshFact
     Call zlClearPatiInfor
     gblnOK = True
End Sub
 

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True: Exit Sub
    If Action = PaneActionFloating Then Cancel = True: Exit Sub
    If Action = PaneActionPinning Then Cancel = True: Exit Sub
    If Action = PaneActionCollapsing Then Cancel = True: Exit Sub
    If Action = PaneActionAttaching Then Cancel = True: Exit Sub
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case Pane_Search    '1
        Item.Handle = picCon.hWnd
    Case Pane_List      ' 2
        Item.Handle = picList.hWnd
    Case Pane_Detail    '3
        Item.Handle = PicDetail.hWnd
    Case Pane_inVoiceList  ' 4
        Item.Handle = picInvoice.hWnd
    End Select
End Sub

Private Sub dkpMan_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Err = 0: On Error Resume Next
     Bottom = stbThis.Height + picDown.Height
End Sub

Private Sub Form_Activate()
    If Not mblnFirst Then Exit Sub
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
End Sub

Private Sub Form_Load()
    mblnFirst = True
    mblnStartFactUseType = zlStartFactUseType(1)
    mlng领用ID = 0
    lblFormat.Alignment = 0

    dtpBegin.MaxDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd") & " 23:59:59")
    dtpEnd.MaxDate = dtpBegin.MaxDate
    dtpBegin.Value = Format(DateAdd("d", -7, dtpBegin.MaxDate), "yyyy-mm-dd")
    dtpEnd.Value = Format(dtpEnd.MaxDate, "yyyy-mm-dd")
    Call InitPanel
    Call zlCardSquareObject
    Call zlCreateObject
End Sub
Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With picDown
        .Top = ScaleHeight - stbThis.Height - .Height
        .Width = ScaleWidth
        .Left = ScaleLeft
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call zlCardSquareObject(True)
    zlSaveDockPanceToReg Me, dkpMan, "区域"
    zl_vsGrid_Para_Save mlngModule, vsInvoiceList, Me.Caption, "发票列表", False
    zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "单据列表", False
    zl_vsGrid_Para_Save mlngModule, vsDetail, Me.Caption, "单据明细列表", False
    Call zlCloseObject
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng卡类别ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
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
Private Sub PicDetail_Resize()
    Err = 0: On Error Resume Next
    With PicDetail
        vsDetail.Left = .ScaleLeft
        vsDetail.Width = .ScaleWidth
        vsDetail.Height = .ScaleHeight
        vsDetail.Top = .ScaleTop
    End With
End Sub

Private Sub picInvoice_Resize()
    Err = 0: On Error Resume Next
    With picInvoice
        vsInvoiceList.Top = txtFilterInvoice.Top + txtFilterInvoice.Height + 20
        vsInvoiceList.Left = .ScaleLeft
        vsInvoiceList.Width = .ScaleWidth
        vsInvoiceList.Height = .ScaleHeight - vsInvoiceList.Top
    End With
End Sub
 Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthday As Date, ByVal strAddress As String)
    If txtPatient.Locked Or txtPatient.Text <> "" Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("身份证", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = strID
    If txtPatient.Text <> "" Then Call FindPati(objCard, False, txtPatient.Text)
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
   If mSquareCard.objSquareCard.zlInitComponents(Me, mlngModule, glngSys, gstrDBUser, gcnOracle, False, strExpend) = False Then
        '初始部件不成功,则作为不存在处理
        Exit Sub
   End If
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
        "Select A.病人ID,Nvl(B.主页ID,0) as 主页ID,A.门诊号 as 门诊号,A.当前床号,B.出院病床,A.姓名,A.性别,Nvl(B.年龄,A.年龄) as 年龄,A.IC卡号,A.就诊卡号,A.卡验证码," & _
        "       Nvl(B.费别,A.费别) as 费别,C.名称 as 当前科室,A.当前科室ID,D.名称 as 出院科室,B.出院科室ID, A.险类 as 险类,E.卡号,E.医保号,E.密码," & _
        "       A.登记时间,Nvl(B.状态,0) as 状态,Nvl(B.医疗付款方式,A.医疗付款方式) as 医疗付款方式,Nvl(B.审核标志,0) as 审核标志,B.入院日期,B.出院日期,B.病人性质,NVL(A.病人类型,B.病人类型) as 病人类型" & _
        " From 病人信息 A,病案主页 B,部门表 C,部门表 D,医保病人档案 E,医保病人关联表 F" & _
        " Where A.停用时间 is NULL And A.病人ID=B.病人ID(+) And A.主页ID=B.主页ID(+) " & _
        "           And A.病人ID=F.病人ID(+) And F.标志(+)=1 And F.医保号=E.医保号(+) And F.险类=E.险类(+) And F.中心 = E.中心(+)" & _
        "           And A.当前科室ID=C.ID(+) And B.出院科室ID=D.ID(+)" & _
        "           And A.停用时间 is NULL "
    
    If blnCard = True And objCard.名称 Like "姓名*" And InStr("-+*", Left(strInput, 1)) = 0 Then  '103563
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
        strSQL = strSQL & " And A.住院号=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '门诊号
        strSQL = strSQL & " And A.门诊号=[1]"
        '75087,冉俊明,2014-7-29,门诊病人收费时,不需要输入完整的门诊号,只需要输入门诊号的最后顺序号即能找到当天就诊的病人信息、费用
        strInput = "*" & zlCommFun.GetFullNO(Mid(strInput, 2), 3)
    ElseIf Left(strInput, 1) = "." Or objCard.名称 = "单据号" Then
        '单据号查找
        If Left(strInput, 1) = "." Then
            strTemp = UCase(GetFullNO(Mid(strInput, 2), 13))
        Else
            strTemp = UCase(GetFullNO(strInput, 13))
        End If
        txtPatient.Text = strTemp
        gstrSQL = "" & _
        "   Select  distinct A.病人ID " & _
        "   From 门诊费用记录 A " & _
        "   Where A.NO=[1] and A.记录性质=[2] " & _
        "              And Rownum=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strTemp, 1)
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
                '75087,冉俊明,2014-7-29,门诊病人收费时,不需要输入完整的门诊号,只需要输入门诊号的最后顺序号即能找到当天就诊的病人信息、费用
                strInput = zlCommFun.GetFullNO(strInput, 3)
            Case "住院号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And A.住院号=[2]"
            Case "发票号"
                strSQL = "" & _
                "   Select distinct A.病人ID " & _
                "   From 门诊费用记录 A,票据打印内容 B,票据使用明细 C" & _
                "   Where A.NO=B.NO and A.记录性质=1 and A.记录状态=1  " & _
                "               and  B.数据性质=1 And B.ID=C.打印ID and C.票种=1 And C.性质=1 And C.号码=[1] And Rownum=1 " & _
                "   "
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput)
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
                txtInvoice.Text = strInput
                GetPatient = True
                Exit Function
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
        '75259：李南春,2014-7-10，病人姓名颜色处理
        Call SetPatiColor(txtPatient, Nvl(mrsInfo!病人类型), IIf(IsNull(mrsInfo!险类), txtPatient.ForeColor, vbRed))
        txtPatient.Text = Nvl(mrsInfo!姓名)
        txtPatient.PasswordChar = ""
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
    '日期:2011-09-04 18:09:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    txtPatient.Text = ""
    ': txtOld.Text = "": txtSex.Text = ""
    'txt住院号.Text = "":
    Set mrsInfo = New ADODB.Recordset
    vsList.Clear 1: vsList.Rows = 2: vsDetail.Clear 1: vsDetail.Rows = 2
    vsInvoiceList.Clear 1: vsInvoiceList.Rows = 2
End Sub

 
Private Sub txtFilterInvoice_Change()
    If mblnNotChange Then Exit Sub
   Call ReadInVoice(True)

End Sub

Private Sub txtFilterInvoice_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then Exit Sub
        If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub
Private Sub txtPatient_Change()
    txtPatient.Tag = ""
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    IDKind.SetAutoReadCard (txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
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
            '103563,只要输入的第一个字符是“-+*”，后面是全数字，都认为不是刷卡
            If Not (InStr("-+*", Left(txtPatient.Text, 1)) > 0 And IsNumeric(Mid(txtPatient.Text, 2))) Then
                blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
            End If
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
            mblnNotChange = True
            Call ReadInVoice(False)
            mblnNotChange = False
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
    '日期:2012-09-03 09:32:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long, lngUnit As Long, i As Long
    Dim strSQL As String, curTotal As Currency, blnIDCard As Boolean
    Dim blnICCard As Boolean, blnMsg As Boolean
    If objCard.名称 Like "IC卡*" And objCard.系统 And mstrPassWord <> "" Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    If objCard.名称 Like "*身份证*" And objCard.系统 And mstrPassWord <> "" Then blnIDCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    
    If Not GetPatient(objCard, strInput, blnCard, blnMsg) Then
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
    If Mid(gstrCardPass, 3, 1) = "1" And (blnCard Or blnICCard Or blnIDCard Or IDKind.GetCurCard.接口序号 <> 0) And mstrPassWord <> "" Then
        If Not zlCommFun.VerifyPassWord(Me, mstrPassWord, mrsInfo!姓名, mrsInfo!性别, "" & mrsInfo!年龄) Then
            Call zlClearPatiInfor
             txtPatient.SetFocus: Exit Sub
        End If
    End If
    mblnNotChange = True
    Call ReadInVoice(False)
    mblnNotChange = False
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
 
Private Sub ShowDetail(Optional strNo As String, Optional strDate As String, _
            Optional ByVal blnDel As Boolean, Optional blnSort As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示明细数据
    '参数:strDate:单据的登记时间
    '编制:刘兴洪
    '日期:2011-09-04 20:18:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, strSQL As String
    Dim lngCol As Long
    
    On Error GoTo errH
    strSQL = _
    " Select C.名称 as 类别,Nvl(E.名称,B.名称) as 名称," & IIf(gTy_System_Para.byt药品名称显示 = 2, "E1.名称 as 商品名,", "") & "B.规格," & _
            IIf(gbln药房单位, "Decode(X.药品ID,NULL,A.计算单位,X." & gstr药房单位 & ")", "A.计算单位") & " as 单位," & _
    "       To_Char(Avg(Nvl(A.付数,1)*" & IIf(blnDel, "-1*", "") & "A.数次)" & _
            IIf(gbln药房单位, "/Nvl(X." & gstr药房包装 & ",1)", "") & ",'9999990.00000') as 数量, " & _
    "       A.费别,To_Char(Sum(A.标准单价)" & _
            IIf(gbln药房单位, "*Nvl(X." & gstr药房包装 & ",1)", "") & ",'999999" & gstrFeePrecisionFmt & "') as 单价, " & _
    "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.应收金额),'9999999" & gstrDec & "') as 应收金额, " & _
    "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.实收金额),'9999999" & gstrDec & "') as 实收金额, " & _
    "       D.名称 as 执行科室,Max(Nvl(A.费用类型,B.费用类型)) as 类型," & _
    "       Decode(Nvl(A.执行状态,0),0,'未执行',1,'完全执行',2,'部分执行',9,'异常收费单','第'||ABS(A.执行状态)||'次退费') as 说明," & _
    "       A.记录状态" & _
    " From  门诊费用记录 A,收费项目目录 B,收费项目类别 C,部门表 D,收费项目别名 E,药品规格 X" & _
              IIf(gTy_System_Para.byt药品名称显示 = 2, ",收费项目别名 E1", "") & _
    " Where A.收费细目ID=B.ID and A.收费类别=C.编码 And A.执行部门ID=D.ID(+) And A.收费细目ID=X.药品ID(+)" & _
    "       And A.记录性质=1 And A.NO=[1] And A.记录状态" & IIf(blnDel, "=2", " IN(1,3)") & _
            IIf(strDate <> "", " And A.登记时间=[2]", "") & _
    "       And A.收费细目ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
            IIf(gTy_System_Para.byt药品名称显示 = 2, "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3", "") & _
            IIf(Not gblnShowErr, " And Nvl(A.附加标志,0)<>9", "") & _
    " Group by Nvl(A.价格父号,A.序号),C.名称,Nvl(E.名称,B.名称)," & IIf(gTy_System_Para.byt药品名称显示 = 2, "E1.名称 ,", "") & " B.规格,A.计算单位,A.费别,D.名称," & _
    "       A.执行状态,A.记录状态,X.药品ID,X." & gstr药房单位 & ",Nvl(X." & gstr药房包装 & ",1)" & _
    " Order by Nvl(A.价格父号,A.序号)"
    If strDate <> "" Then
        Set mrsDetail = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo, CDate(strDate))
    Else
        Set mrsDetail = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo, "")
    End If
    With vsDetail
        .Clear 1
        .Redraw = flexRDNone
        Set .DataSource = mrsDetail
        For lngCol = 0 To .COLS - 1
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
        Call .AutoSize(0, .COLS - 1)
        zl_vsGrid_Para_Restore mlngModule, vsDetail, Me.Caption, "单据明细列表", False
        .Redraw = flexRDBuffered
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Function ReadInVoice(ByVal blnFilter As Boolean) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载发票信息
    '入参:
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-04-02 15:33:06
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng病人ID As Long, dtStartDate As Date, dtEndDate As Date
    Dim strWhere As String
    On Error GoTo errHandle
    dtStartDate = CDate("1901-01-01")
    dtEndDate = dtStartDate
    If mrsInfo Is Nothing Then
        lng病人ID = 0
    ElseIf mrsInfo.State <> 1 Then
        lng病人ID = 0
    Else
        lng病人ID = Val(Nvl(mrsInfo!病人ID))
    End If
    If chkDate.Value = 0 Then
        strWhere = strWhere & " And A.发生时间+0 betWeen [2] and [3]"
        dtStartDate = CDate(Format(dtpBegin.Value, "yyyy-mm-dd") & " 00:00:00")
        dtEndDate = CDate(Format(dtpEnd.Value, "yyyy-mm-dd") & " 23:59:59")
    End If
    
    If blnFilter = False Or mrsInVoice Is Nothing Then
        gstrSQL = "" & _
        "   Select  C.号码 as 发票号,C.使用人,to_char(C.使用时间,'yyyy-mm-dd hh24:mi:ss') as 使用时间," & _
        "               Sum(nvl(Q.实收金额,0)) as  发票金额" & _
        "   From (     Select  A.NO,nvl(A.价格父号,A.序号) as 序号,sum(实收金额) as 实收金额" & _
        "                   From 门诊费用记录   A" & _
        "                    Where Mod(a.记录性质,0)=1 And A.病人ID=[1]  " & strWhere & _
        "                   Group by NO,nvl(A.价格父号,A.序号) " & _
        "               )  Q,票据打印明细 B,票据使用明细 C" & _
        "   Where   q.NO=B.NO  And  instr(','||B.序号||',',','||Q.序号||',')>0  And B.使用ID=C.ID" & _
        "               And  B.票种=1 And nvl(B.是否回收,0)=0 " & _
        "    Group by  C.号码  ,C.使用人,to_char(C.使用时间,'yyyy-mm-dd hh24:mi:ss') " & _
        "   Order by 发票号 "
        Set mrsInVoice = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng病人ID, dtStartDate, dtEndDate)
    End If
    mrsInVoice.Filter = 0
    With vsInvoiceList
        .Clear 1
        .Rows = 2
        .Row = -1
        If mrsInVoice.RecordCount <> 0 Then mrsInVoice.MoveFirst
        Do While Not mrsInVoice.EOF
            If Nvl(mrsInVoice!发票号) Like "*" & txtFilterInvoice.Text & "*" _
                Or Trim(txtFilterInvoice.Text) = "" Then
                If .TextMatrix(.Rows - 1, .ColIndex("发票号")) <> "" Then .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, .ColIndex("发票号")) = Nvl(mrsInVoice!发票号)
                .TextMatrix(.Rows - 1, .ColIndex("使用人")) = Nvl(mrsInVoice!使用人)
                .TextMatrix(.Rows - 1, .ColIndex("使用日期")) = Nvl(mrsInVoice!使用时间)
                .TextMatrix(.Rows - 1, .ColIndex("发票金额")) = Format(Nvl(mrsInVoice!发票金额), "#######" & gstrDec)
            End If
            mrsInVoice.MoveNext
        Loop
        .Row = 1
    End With
    ReadInVoice = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function



Private Function ReadListData(ByVal strInvoiceNO As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取需要明细数据
    '入参:strInvoiceNO-发票号
    '返回:读取成功,返回true,否则返回Flase
    '编制:刘兴洪
    '日期:2011-01-25 17:10:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng病人ID As Long, strTable As String, rsTemp As ADODB.Recordset
    Dim lngCol As Long, strSQL As String, lngRow As Long
    Dim strFilter As String, strNos As String
    Dim strWhere As String, strTable1 As String, dtStartDate As Date, dtEndDate As Date
    Dim strNo As String, i As Long, lng结帐序号 As Long
    
    mblnSel = False
    On Error GoTo errHandle
    zlCommFun.ShowFlash "正在读取单据数据,请稍候 ..."
    Screen.MousePointer = 11
    strTable1 = " Select distinct NO From 票据打印明细  Where 票种=1 And 票号=[1] "
    DoEvents
    strTable = "" & _
    "   Select  A.NO,A.实际票号,A.姓名,A.性别,A.年龄, " & _
    "         Decode(A.门诊标志,2,'',A.标识号) as 门诊号,  " & _
    "         Decode(A.门诊标志,2,A.标识号,'') as 住院号, " & _
    "         Min(A.费别) as 费别,  " & _
    "        Sum(Decode(A.记录状态,2,-1,1)*A.应收金额) as 应收金额," & _
    "        Sum(Decode(A.记录状态,2,-1,1)*A.实收金额) as 实收金额, " & _
    "         Max(A.记录状态) as 符号,A.结帐ID," & _
    "        A.开单人,A.开单部门ID,A.付款方式,A.划价人,A.操作员姓名,A.登记时间" & _
    "   From 门诊费用记录 A,( " & strTable1 & ") B" & _
    "   Where A.记录性质 =1 and A.NO=B.NO  " & _
    "   Group by A.NO,A.实际票号,A.姓名,A.性别,A.年龄, " & _
    "               Decode(A.门诊标志,2,'',A.标识号),Decode(A.门诊标志,2,A.标识号,'')," & _
    "               A.开单人,A.开单部门ID,A.付款方式,A.划价人,A.操作员姓名,A.登记时间,A.结帐ID"
           
    strSQL = _
    " Select Decode(Nvl(Max(t.险类),0),0,NULL,'√') as 医保,A.NO as 单据号, " & _
    "       Min(A.实际票号) as 票据号,Max(B.名称) as 开单科室," & _
    "       max(A.开单人) as 开单人,max(A.门诊号) as 门诊号,max(A.住院号) as 住院号, " & _
    "       max(C.名称) as  医疗付款方式, max(A.姓名) as 姓名,max(A.性别) as 性别,max(A.年龄) as 年龄," & _
    "       min(A.费别) as 费别, " & _
    "       To_Char(max(A.应收金额),'9999999" & gstrDec & "')   as 应收金额," & _
    "       To_Char(max(A.实收金额),'9999999" & gstrDec & "')  as 实收金额," & _
    "       max(A.划价人) as 划价人,max(A.操作员姓名) as 操作员, " & _
    "       To_Char(max(A.登记时间) ,'YYYY-MM-DD HH24:MI:SS') as 登记时间," & _
    "       max(A.符号) as 符号,A.结帐ID,Max(nvl(M.结算序号,A.结帐ID)) as 结算序号ID,nvl(Max(t.险类),0) as 险类ID" & _
    " From (" & strTable & ") A,病人预交记录 M, 部门表 B,医疗付款方式 C,保险结算记录 T" & _
    " Where  A.开单部门ID=B.ID And A.付款方式=C.编码(+)  " & _
    "       And A.结帐id=t.记录ID(+) And t.性质(+)=1 And A.结帐id=M.结帐ID(+) " & _
    "       And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & vbNewLine & _
    " Group by  A.No,A.结帐ID" & _
    " Order by 结算序号ID desc,登记时间 Desc,单据号 Desc"
    Set mrsList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInvoiceNO)
    vsList.Redraw = flexRDNone
    vsList.Clear: vsList.COLS = 2
    Set vsList.DataSource = mrsList
    If vsList.Rows <= 1 Then vsList.Rows = 2
    With vsList
        mstrNOs = ""
        For lngCol = 0 To .COLS - 1
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
        Call .AutoSize(0, .COLS - 1)
        zl_vsGrid_Para_Restore mlngModule, vsList, Me.Caption, "单据列表", False
        .Redraw = flexRDBuffered
        lng结帐序号 = 0
        For i = 1 To .Rows - 1
            strNo = Trim(.TextMatrix(i, .ColIndex("单据号")))
            If InStr(1, mstrNOs & ",", "," & strNo & ",") = 0 Then
                mstrNOs = mstrNOs & "," & strNo
            End If
            If lng结帐序号 <> Trim(.TextMatrix(lngRow, .ColIndex("结算序号ID"))) _
                 And lng结帐序号 <> 0 Then
                '画出分隔线
                .Select lngRow, .FixedCols, lngRow, .COLS - 1
                .CellBorder vbBlue, 0, 1, 0, 0, 0, 0
            End If
        Next
        vsList_AfterRowColChange 0, 0, .Row, .Col
    End With
    mrsList.Filter = "险类ID>0"
    If Not mrsList.EOF Then
        mintInsure = Val(Nvl(mrsList!险类ID))
    Else
        mintInsure = 0
    End If
    zlCommFun.StopFlash
    Call ReInitPatiInvoice
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
Private Sub vsInvoiceList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsInvoiceList, Me.Caption, "发票列表", False
End Sub

Private Sub vsInvoiceList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsInvoiceList, Me.Caption, "发票列表", False
End Sub

Private Sub vsDetail_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsDetail, Me.Caption, "单据明细列表", False
End Sub

Private Sub vsDetail_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsDetail, Me.Caption, "单据明细列表", False
End Sub
 
Private Sub vsInvoiceList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strInvoiceNO As String
    If OldRow = NewRow Then Exit Sub
    If NewRow = -1 Then Exit Sub
    With vsInvoiceList
        strInvoiceNO = .TextMatrix(NewRow, .ColIndex("发票号"))
    End With
    Call ReadListData(strInvoiceNO)
End Sub
 
 

Private Sub vsList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "单据列表", False
End Sub

Private Sub vsList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strNo As String
    If OldRow = NewRow Then Exit Sub
    With vsList
        If NewRow < 0 Then Exit Sub
        strNo = Trim(.TextMatrix(NewRow, .ColIndex("单据号")))
    End With
    ShowDetail strNo
    
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

Private Sub ReInitPatiInvoice(Optional blnFact As Boolean = True, _
    Optional ByVal intInsure_IN As Integer = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新初始化病人发票信息
    '入参:blnFact-是否重新取发票号
    '编制:刘兴洪
    '日期:2011-04-29 14:17:33
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInvoiceFormat As String, lng病人ID As Long
    Dim intInsure As Integer
  
    If Not mrsInfo Is Nothing Then
      If mrsInfo.State = 1 Then lng病人ID = Val(Nvl(mrsInfo!病人ID))
    End If
    intInsure = IIf(intInsure_IN <> 0, intInsure_IN, mintInsure)
    mstrUseType = "": mlngShareUseID = 0: mintInvoiceFormat = 0
    mstrUseType = zl_GetInvoiceUserType(lng病人ID, 0, intInsure)
    mlngShareUseID = zl_GetInvoiceShareID(mlngModule, mstrUseType)
    mintInvoiceFormat = zl_GetInvoicePrintFormat(mlngModule, mstrUseType, mintOldInvoiceFormat)
    mintInvoicePrint = zl_GetInvoicePrintMode(mlngModule, mstrUseType)
    Call ZlShowBillFormat(mlngModule, lblFormat, mintInvoiceFormat)
    If blnFact Then Call RefreshFact
End Sub

Private Sub RefreshFact()
    '功能：刷新收费票据号
  '  If mintInvoicePrint = 0 Then Exit Sub
    If gblnStrictCtrl Then
        'lblFact.tag主要是检查发票号是否手工输入的.手工输入的,发票号为空,否则是自动产生的发票号
        If (lblFact.Tag <> "" And txtInvoice.Text <> "") Or Trim(txtInvoice.Text) = "" Then
            If zlCheckInvoiceValied(mlng领用ID, 1, , mlngShareUseID, mstrUseType) = False Then
                txtInvoice.Text = "": txtInvoice.Tag = "": Exit Sub
            End If
            '严格：取下一个号码
            txtInvoice.Text = GetNextBill(mlng领用ID)
            'Tag：问题：24363:刘兴洪：主要是解决自动生成的号是否被用户更改，主要解决：
            '    1.更改的票据号需要检查是否重复，重复后直接返回不更改发票号
            '    2.并发操作，不更改的情况下，检查是否重复，如果重复，自动取下一个号码！
            txtInvoice.Tag = txtInvoice.Text
            lblFact.Tag = txtInvoice.Tag
            If mblnStartFactUseType Then Call zlCheckFactIsEnough
        End If
    Else
        If (lblFact.Tag <> "" And txtInvoice.Text <> "") Or Trim(txtInvoice.Text) = "" Then
            '松散：取下一个号码
            txtInvoice.Text = zlStr.Increase(UCase(zlDatabase.GetPara("当前收费票据号", glngSys, mlngModule)))
            'Tag：问题：24363:刘兴洪：主要是解决自动生成的号是否被用户更改，主要解决：
            '    1.更改的票据号需要检查是否重复，重复后直接返回不更改发票号
            '    2.并发操作，不更改的情况下，检查是否重复，如果重复，自动取下一个号码！
        End If
        txtInvoice.Tag = txtInvoice.Text
        lblFact.Tag = txtInvoice.Tag
    End If
    txtInvoice.SelStart = Len(txtInvoice.Text)
End Sub

Private Sub zlCheckFactIsEnough()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查当前票据是否允足
    '编制:刘兴洪
    '日期:2011-05-10 17:54:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng剩余数量 As Long
    '刘兴洪 问题:26948 日期:2009-12-28 17:43:00
    '需要检查剩余数量是否充足:
    If zlCheckInvoiceOverplusEnough(1, gTy_Module_Para.int提醒剩余票据张数, lng剩余数量, mlng领用ID, mstrUseType) = False Then
        MsgBox "注意:" & vbCrLf & _
               "    当前剩余票据(" & lng剩余数量 & ") 小于了报警的张数(" & gTy_Module_Para.int提醒剩余票据张数 & "),请注意更换发票!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
    End If
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



