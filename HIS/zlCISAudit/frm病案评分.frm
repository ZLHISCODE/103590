VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm病案评分 
   AutoRedraw      =   -1  'True
   Caption         =   "电子病案评分"
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13425
   Icon            =   "frm病案评分.frx":0000
   ScaleHeight     =   8985
   ScaleWidth      =   13425
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chk结果 
      Caption         =   "未评分"
      Height          =   240
      Index           =   0
      Left            =   270
      TabIndex        =   24
      Top             =   675
      Value           =   1  'Checked
      Width           =   915
   End
   Begin VB.CheckBox chk结果 
      Caption         =   "未审核"
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   1
      Left            =   1215
      TabIndex        =   23
      Top             =   675
      Value           =   1  'Checked
      Width           =   915
   End
   Begin VB.CheckBox chk结果 
      Caption         =   "已审核"
      ForeColor       =   &H00808080&
      Height          =   240
      Index           =   2
      Left            =   2160
      TabIndex        =   22
      Top             =   675
      Value           =   1  'Checked
      Width           =   915
   End
   Begin zl9CISAudit.tipPopup tipPopup1 
      Height          =   420
      Left            =   135
      Top             =   9240
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid fg列表 
      Height          =   2865
      Left            =   345
      TabIndex        =   16
      Top             =   3630
      Visible         =   0   'False
      Width           =   1635
      _cx             =   2884
      _cy             =   5054
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm病案评分.frx":08CA
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
      Ellipsis        =   1
      ExplorerBar     =   7
      PicturesOver    =   -1  'True
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
      WallPaperAlignment=   1
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSComctlLib.TreeView tvw科室 
      Height          =   1170
      Left            =   2145
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3585
      Visible         =   0   'False
      Width           =   3450
      _ExtentX        =   6085
      _ExtentY        =   2064
      _Version        =   393217
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "Img小图标"
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.PictureBox picRight 
      BackColor       =   &H00FAFAFA&
      ClipControls    =   0   'False
      Height          =   4665
      Left            =   7095
      Picture         =   "frm病案评分.frx":0917
      ScaleHeight     =   4605
      ScaleWidth      =   6015
      TabIndex        =   5
      Top             =   1530
      Width           =   6075
      Begin VSFlex8Ctl.VSFlexGrid fg结果_S 
         Height          =   1425
         Left            =   210
         TabIndex        =   1
         Top             =   1920
         Width           =   4920
         _cx             =   8678
         _cy             =   2514
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   16777215
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frm病案评分.frx":0E14
         ScrollTrack     =   -1  'True
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
         Ellipsis        =   1
         ExplorerBar     =   0
         PicturesOver    =   -1  'True
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
         WallPaperAlignment=   4
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lbl病理类型 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病理类型:"
         Height          =   180
         Left            =   2550
         TabIndex        =   25
         Top             =   870
         Width           =   810
      End
      Begin VB.Label lbl备注 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "备注:"
         Height          =   180
         Left            =   180
         TabIndex        =   18
         Top             =   1590
         Width           =   450
      End
      Begin VB.Label lbl返回修改 
         BackStyle       =   0  'Transparent
         Caption         =   "√返回修改"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2565
         TabIndex        =   17
         Top             =   645
         Width           =   2580
      End
      Begin VB.Label lbl评分时间 
         BackStyle       =   0  'Transparent
         Caption         =   "评分时间:"
         Height          =   195
         Left            =   2565
         TabIndex        =   15
         Top             =   1113
         Width           =   2580
      End
      Begin VB.Label lbl审核时间 
         BackStyle       =   0  'Transparent
         Caption         =   "审核时间:"
         Height          =   195
         Left            =   2565
         TabIndex        =   14
         Top             =   1350
         Width           =   2580
      End
      Begin VB.Label lbl病人信息 
         BackStyle       =   0  'Transparent
         Caption         =   "病人信息"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   405
         Width           =   2580
      End
      Begin VB.Label lbl等级 
         BackStyle       =   0  'Transparent
         Caption         =   "等级:"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   877
         Width           =   2580
      End
      Begin VB.Label lbl评分人 
         BackStyle       =   0  'Transparent
         Caption         =   "评分人:"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   1113
         Width           =   2580
      End
      Begin VB.Label lbl审核人 
         BackStyle       =   0  'Transparent
         Caption         =   "审核人:"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   1350
         Width           =   2580
      End
      Begin VB.Label lbl总分 
         BackStyle       =   0  'Transparent
         Caption         =   "总分:"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   641
         Width           =   2580
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "评分结果"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F0F0F0&
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   90
         Width           =   1095
      End
   End
   Begin VB.PictureBox picV 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2115
      Left            =   4950
      MousePointer    =   9  'Size W E
      ScaleHeight     =   2115
      ScaleWidth      =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2460
      Width           =   0
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   8625
      Width           =   13425
      _ExtentX        =   23680
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm病案评分.frx":0F6D
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20770
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
   Begin MSComctlLib.ImageList Img小图标 
      Left            =   5145
      Top             =   6825
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm病案评分.frx":1801
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm病案评分.frx":1978
            Key             =   "Dot"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picLeft_S 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1800
      Left            =   300
      ScaleHeight     =   1800
      ScaleWidth      =   4650
      TabIndex        =   12
      Top             =   1440
      Width           =   4650
      Begin VSFlex8Ctl.VSFlexGrid fg病案_S 
         Height          =   1020
         Left            =   135
         TabIndex        =   0
         Top             =   345
         Width           =   4365
         _cx             =   7699
         _cy             =   1799
         Appearance      =   0
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   14737632
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   26
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frm病案评分.frx":1A50
         ScrollTrack     =   -1  'True
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
         Ellipsis        =   1
         ExplorerBar     =   7
         PicturesOver    =   -1  'True
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
         WallPaperAlignment=   1
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Image imgSelCols_S 
         Height          =   195
         Left            =   4275
         MouseIcon       =   "frm病案评分.frx":1DC8
         MousePointer    =   99  'Custom
         Picture         =   "frm病案评分.frx":1F1A
         ToolTipText     =   "选择需要显示的列"
         Top             =   90
         Width           =   195
      End
      Begin VB.Image imgRefresh 
         Height          =   195
         Left            =   4005
         MouseIcon       =   "frm病案评分.frx":1F6D
         MousePointer    =   99  'Custom
         Picture         =   "frm病案评分.frx":20BF
         ToolTipText     =   "刷新数据"
         Top             =   90
         Width           =   195
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "病案信息"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F0F0F0&
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   90
         Width           =   1095
      End
   End
   Begin VB.TextBox txt内容 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2970
      TabIndex        =   19
      Top             =   105
      Width           =   2760
   End
   Begin VB.TextBox txt科室 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   3390
      TabIndex        =   21
      Top             =   645
      Width           =   2250
   End
   Begin VB.Image imgLujinPic 
      Height          =   240
      Left            =   11070
      MouseIcon       =   "frm病案评分.frx":22FF
      MousePointer    =   99  'Custom
      Picture         =   "frm病案评分.frx":2451
      ToolTipText     =   "刷新数据"
      Top             =   8040
      Width           =   240
   End
   Begin VB.Image imgLujin 
      Height          =   240
      Left            =   11085
      MouseIcon       =   "frm病案评分.frx":8CA3
      MousePointer    =   99  'Custom
      Picture         =   "frm病案评分.frx":8DF5
      ToolTipText     =   "刷新数据"
      Top             =   8265
      Width           =   240
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frm病案评分.frx":F647
      Left            =   465
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
      ScaleMode       =   1
   End
   Begin VB.Label Label4 
      Caption         =   "内容(&T)"
      Height          =   195
      Left            =   2295
      TabIndex        =   20
      Top             =   165
      Width           =   870
   End
   Begin VB.Image imgBGBlue 
      Height          =   1530
      Left            =   165
      Picture         =   "frm病案评分.frx":F65B
      Top             =   9750
      Visible         =   0   'False
      Width           =   2790
   End
   Begin VB.Image imgBG_fg 
      Height          =   2970
      Index           =   0
      Left            =   3090
      Picture         =   "frm病案评分.frx":F81B
      Top             =   9735
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.Image imgBG_fg 
      Height          =   2970
      Index           =   1
      Left            =   6345
      Picture         =   "frm病案评分.frx":1003F
      Top             =   9810
      Visible         =   0   'False
      Width           =   3105
   End
End
Attribute VB_Name = "frm病案评分"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Private WithEvents mfrm评分结果编辑 As frm评分结果编辑
Attribute mfrm评分结果编辑.VB_VarHelpID = -1
Private mstrPrivs               As String               '权限串
Private mobjFindKey             As CommandBarPopup      '查询
Private mstrFindKey             As String               '查询串
Private mlngModule              As Long                 '模块号
Private m_lngOldRow             As Long
Private mcbrPopupBar            As CommandBar           '弹出窗口
Private mRecordRating           As Boolean              '评分否
Private mRecordAudit            As Boolean              '审核否
Private mRecordMyAudit          As Boolean              '是否本人记录
Private mRecordReturn           As Boolean              '重新评分否
Private mvarPara                As Variant              '查询参数
Private rsM                     As ADODB.Recordset      '病案数据集
Private mstrWhere               As String               '当前的查询条件
Private mblnSetDept             As Boolean              '是否限定科室

Dim m_lng病人ID                 As Long                 '当前病人ID
Dim m_lng主页ID                 As Long                 '当前主页ID
Dim m_lng结果ID                 As Long                 '当前结果ID
Dim m_lng方案ID                 As Long                 '当前评分方案ID
Dim m_str列标题                 As String               '
Dim mfrm查找                    As frm病案评分查询      '用于查找病案的窗体，作为一个局部变量使用。（在Form_QueryUnload中关闭，否则只是隐藏之！）
Dim cbrPopupItem                As CommandBarControl    '弹出项
'查询窗口变量
Private mlngSickID              As Long             '病人ID
Private mlngHospitalID          As Long             '住院号
Private mlngHospitalTimes       As Long             '住院次数
Private mstrSickName            As String           '病人姓名
Private mstrMainDoctor          As String           '主治医师
Private mstrOutpatientDoctor    As String           '门诊医师
Private mstrNurses              As String           '责任护士
Private mstrRatingMan           As String           '评分人
Private mstrAuditMan            As String           '审核人
Private mstrOutDept             As String           '出院科室
Private mstrInDept              As String           '入院科室
Private mdatStarOutDate         As Date             '出院开始日期
Private mdatEndOutDate          As Date             '出院开始日期
Private mdatStarInDate          As Date             '入院开始日期
Private mdatEndInDate           As Date             '入院开始日期
Private mstrSickType            As String           '病理类型
Private mfrmArchiveView         As frmArchiveView   '病案查阅

'==============================================================================
'=功能： 控件初始化
'==============================================================================
Private Sub InitControl()
    On Error GoTo ErrH
    
    '菜单控制
    Call InitCommandBar
    '加载区域
    Call InitDockPannel
    '初始化网格
    Call InitVsf
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 初始区域划分
'==============================================================================
Private Sub InitDockPannel()
    Dim objPane             As Pane

    On Error GoTo ErrH
    
    Set objPane = dkpMain.CreatePane(1, 200, 100, DockLeftOf, Nothing)
    objPane.Title = "病案信息"
    objPane.Options = PaneNoCaption
    
    Set objPane = dkpMain.CreatePane(2, 200, 100, DockRightOf, Nothing)
    objPane.Title = "评分结果"
    objPane.Options = PaneNoCaption
    
    dkpMain.SetCommandBars cbsMain
    
    Call DockPannelInit(dkpMain)
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 初始菜单工具栏
'==============================================================================
Private Sub InitCommandBar()
    Dim objMenu         As CommandBarPopup
    Dim objBar          As CommandBar
    Dim objExtendedBar  As CommandBar
    Dim objPopup        As CommandBarPopup
    Dim objControl      As CommandBarControl
    Dim cbrCustom       As CommandBarControlCustom
    
    On Error GoTo ErrH
    
    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    Call CommandBarInit(cbsMain)

    '------------------------------------------------------------------------------------------------------------------
    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    '------------------------------------------------------------------------------------------------------------------
    '文件
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)...")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Preview, "预览(&V)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Print, "打印(&P)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Excel, "输出到&Excel")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_BatPrint, "全部打印(&A)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Exit, "退出(&X)", True)
    
    '------------------------------------------------------------------------------------------------------------------
    '编辑
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_NewParent, "病案评分(&A)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_ModifyParent, "修改结果(&M)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Insert, "重新评分(&R)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_DeleteParent, "删除结果(&D)")
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_ReportView, "查阅病案(&V)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_Audit, "通过审核(&P)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Leave_UndoPost, "取消审核(&C)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Select, "全部选中(&L)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_DeSelect, "取消选中(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_UnAudit, "反向选择(&B)")
    
    '------------------------------------------------------------------------------------------------------------------
    '查看
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Find, "过滤(&F)...")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Refresh, "刷新(&R)", True)
    
    '------------------------------------------------------------------------------------------------------------------
    '帮助
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的" & ParamInfo.产品名称)
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Home, ParamInfo.产品名称 & "主页(&H)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Forum, ParamInfo.产品名称 & "论坛(&F)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Help_About, "关于(&A)…", True)
    
    '主菜单右侧的查找
    '------------------------------------------------------------------------------------------------------------------
    cbsMain.ActiveMenuBar.SetIconSize 16, 16
    If Len(mstrFindKey) <= 2 Then mstrFindKey = "住院号"
    Set mobjFindKey = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_View_LocationItem, mstrFindKey)
    mobjFindKey.IconId = conMenu_View_Find
    mobjFindKey.Flags = xtpFlagRightAlign
    mobjFindKey.STYLE = xtpButtonIconAndCaption
    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&1.住院号", , , "住院号")
    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&2.病人姓名", , , "病人姓名")
'    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&3.病人ID", , , "病人ID")
'    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&4.就诊卡号", , , "就诊卡号")
'    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&5.床位、门诊号", , , "床位、门诊号")
    Set cbrCustom = cbsMain.ActiveMenuBar.Controls.Add(xtpControlCustom, conMenu_View_Location, "")
    cbrCustom.Handle = txt内容.hWnd
    cbrCustom.Flags = xtpFlagRightAlign
    Set objControl = cbsMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_View_Forward, "前一条")
    objControl.Flags = xtpFlagRightAlign
    objControl.STYLE = xtpButtonIcon
    Set objControl = cbsMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_View_Backward, "后一条")
    objControl.Flags = xtpFlagRightAlign
    objControl.STYLE = xtpButtonIcon
    
    '标准工具栏
    '------------------------------------------------------------------------------------------------------------------
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_ReportView, "查阅病案(&V)", True)
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_Audit, "通过审核(&P)", True)
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Leave_UndoPost, "取消审核(&C)")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Print, "打印")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Preview, "预览")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_NewParent, "评分", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_ModifyParent, "修改")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Insert, "重评")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_DeleteParent, "删除")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Manage_ReportView, "查阅", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Manage_Audit, "审核", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Leave_UndoPost, "消审")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_View_Find, "过滤", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Help_Help, "帮助", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "退出")
    '标准过滤栏
    '------------------------------------------------------------------------------------------------------------------
    Set objBar = cbsMain.Add("过滤栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    If ScaleWidth - 4000 > 0 Then txt科室.Width = 4000
    Set cbrCustom = objBar.Controls.Add(xtpControlCustom, 10000, "")
    cbrCustom.Handle = chk结果(0).hWnd
    cbrCustom.Flags = xtpFlagLeftPopup
    Set cbrCustom = objBar.Controls.Add(xtpControlCustom, 10001, "")
    cbrCustom.Handle = chk结果(1).hWnd
    cbrCustom.Flags = xtpFlagLeftPopup
    Set cbrCustom = objBar.Controls.Add(xtpControlCustom, 10002, "")
    cbrCustom.Handle = chk结果(2).hWnd
    cbrCustom.Flags = xtpFlagLeftPopup
    chk结果(2).ForeColor = vbRed
    Set objControl = NewToolBar(objBar, xtpControlLabel, conMenu_Help_Help, "科室", True)
    Set cbrCustom = objBar.Controls.Add(xtpControlCustom, 10003, "")
    cbrCustom.Handle = txt科室.hWnd
    cbrCustom.Flags = xtpFlagLeftPopup
    objBar.Controls.Add xtpControlButton, 10004, ""
    '------------------------------------------------------------------------------------------------------------------
    '命令的快键绑定:公共部份主界面已处理
    With cbsMain.KeyBindings
        .Add 0, vbKeyF11, conMenu_Manage_ReportView         '查阅
        .Add 0, vbKeyF5, conMenu_View_Refresh               '刷新
        .Add 0, vbKeyF1, conMenu_Help_Help                  '帮助
        .Add FCONTROL, vbKeyF, conMenu_View_Find            '查找
        .Add FCONTROL, vbKeyP, conMenu_File_Print           '打印
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem         '新增
        .Add FCONTROL, vbKeyI, conMenu_Edit_CopyNewItem     '插入
        .Add FCONTROL, vbKeyE, conMenu_Edit_Modify          '修改
        .Add FSHIFT, vbKeyDelete, conMenu_Edit_Delete       '删除
        .Add FCONTROL, vbKeyS, conMenu_Edit_Transf_Save     '保存
        .Add 0, vbKeyF3, conMenu_View_Location              '定位
        .Add 0, vbKeyF4, conMenu_View_Option                '选择定位依据
        .Add FCONTROL, vbKeyLeft, conMenu_View_Forward      '前一条
        .Add FCONTROL, vbKeyRight, conMenu_View_Backward    '后一条
    End With
    '------------------------------------------------------------------------------------------------------------------
    '弹出菜单分类
    Set mcbrPopupBar = cbsMain.Add("弹出菜单", xtpBarPopup)
    Set cbrPopupItem = mcbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_NewParent, "病案评分(&A)")
    Set cbrPopupItem = mcbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_ModifyParent, "修改结果(&R)")
    Set cbrPopupItem = mcbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Insert, "重新评分(&M)")
    Set cbrPopupItem = mcbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_DeleteParent, "删除结果(&D)")
    Set cbrPopupItem = mcbrPopupBar.Controls.Add(xtpControlButton, conMenu_Manage_ReportView, "查阅病案(&V)", True)
    Set cbrPopupItem = mcbrPopupBar.Controls.Add(xtpControlButton, conMenu_Manage_Audit, "通过审核(&P)", True)
    Set cbrPopupItem = mcbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Leave_UndoPost, "取消审核(&C)")
    Set cbrPopupItem = mcbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Select, "全部选中(&L)", True)
    Set cbrPopupItem = mcbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_DeSelect, "取消选中(&S)")
    Set cbrPopupItem = mcbrPopupBar.Controls.Add(xtpControlButton, conMenu_Manage_UnAudit, "反向选择(&B)")
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 界面分割
'==============================================================================
Private Sub InitVsf()
    Dim i           As Long
    On Error GoTo ErrH
    With fg列表
        .Rows = 24
        .Cell(flexcpText, 1, 1) = "病人ID"
        .Cell(flexcpText, 2, 1) = "路径图标"
        .Cell(flexcpText, 3, 1) = "住院号"
        .Cell(flexcpText, 4, 1) = "住院次数"
        .Cell(flexcpText, 5, 1) = "姓名"
        .Cell(flexcpText, 6, 1) = "性别"
        .Cell(flexcpText, 7, 1) = "住院医师"
        .Cell(flexcpText, 8, 1) = "门诊医师"
        .Cell(flexcpText, 9, 1) = "责任护士"
        .Cell(flexcpText, 10, 1) = "出院科室"
        .Cell(flexcpText, 11, 1) = "出院日期"
        .Cell(flexcpText, 12, 1) = "入院科室"
        .Cell(flexcpText, 13, 1) = "入院日期"
        .Cell(flexcpText, 14, 1) = "编目日期"
        .Cell(flexcpText, 15, 1) = "评分人"
        .Cell(flexcpText, 16, 1) = "评分时间"
        .Cell(flexcpText, 17, 1) = "审核人"
        .Cell(flexcpText, 18, 1) = "审核时间"
        .Cell(flexcpText, 19, 1) = "总分"
        .Cell(flexcpText, 20, 1) = "等级"
        .Cell(flexcpText, 21, 1) = "返回修改"
        .Cell(flexcpText, 22, 1) = "备注"
        .Cell(flexcpText, 23, 1) = "病理类型"
        .Cell(flexcpChecked, 1, 0, .Rows - 1, 0) = flexUnchecked
        .Editable = flexEDKbdMouse
    End With
    For i = 1 To fg列表.Rows - 1
        If fg病案_S.ColWidth(fg病案_S.ColIndex(fg列表.Cell(flexcpText, i, 1))) < 100 Then
            fg列表.Cell(flexcpChecked, i, 0) = flexUnchecked
        Else
            fg列表.Cell(flexcpChecked, i, 0) = flexChecked
        End If
    Next
    fg列表.ZOrder 0
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 界面分割
'==============================================================================
Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    On Error GoTo ErrH
    
    Select Case Item.ID
        Case 1
            Item.Handle = picLeft_S.hWnd
        Case 2
            Item.Handle = picRight.hWnd
    End Select
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 加载科室数据
'==============================================================================
Private Sub InitTvw()
    Dim rsTemp      As ADODB.Recordset
    Dim strTmp      As String
    Dim varTmp      As Variant
    Dim varAry      As Variant
    Dim lngCount    As Long
    Dim strDept     As String
    Dim intCol      As Integer
    
    On Error GoTo ErrH
    
    '列出部门表和对应人员
    strTmp = GetPara("评分科室范围", mlngModule)
    varTmp = Split(strTmp, ";")
    strDept = ""
    For lngCount = 0 To UBound(varTmp)
        varAry = Split(varTmp(lngCount), ",")
        If UserInfo.ID = varAry(0) Then
            strDept = varTmp(lngCount)
            strDept = Mid(strDept, InStr(1, strDept, ",")) & ","
            Exit For
        End If
    Next
    
    tvw科室.Nodes.Clear
    gstrSQL = "select id,编码,名称,上级id From 部门表 where ( TO_CHAR (撤档时间, 'yyyy-MM-dd') = '3000-01-01' or 撤档时间 is null)" & _
             " start with 上级id is null connect by prior id = 上级id"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If IsPrivs(mstrPrivs, "所有科室") Then tvw科室.Nodes.Add , , "C0", "【+】所有科室", "Search", "Search"
    If strDept = "" Then
        mblnSetDept = False
        Do Until rsTemp.EOF
            If IsNull(rsTemp("上级id")) Then
                tvw科室.Nodes.Add , , "C" & rsTemp("id"), "【" & rsTemp("编码") & "】" & rsTemp("名称"), "Search", "Search"
            Else
                tvw科室.Nodes.Add "C" & rsTemp("上级id"), tvwChild, "C" & rsTemp("id").Value, "【" & rsTemp("编码") & "】" & rsTemp("名称"), "Dot", "Dot"
            End If
            rsTemp.MoveNext
        Loop
    ElseIf IsPrivs(mstrPrivs, "所有科室") Then
        mblnSetDept = True
        Do Until rsTemp.EOF
            If InStr(1, strDept, "," & rsTemp("id").Value & ",") Then
                tvw科室.Nodes.Add "C0", tvwChild, "C" & rsTemp("id").Value, "【" & rsTemp("编码") & "】" & rsTemp("名称"), "Dot", "Dot"
            End If
            rsTemp.MoveNext
        Loop
    End If
    mblnSetDept = False
    rsTemp.Close
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 更新数据
'==============================================================================
Private Sub mfrm评分结果编辑_AferSaveData()
    Call 更新ID
    Call Fill评分结果
End Sub

'==============================================================================
'=功能： 病案评分
'==============================================================================
Private Sub RecordRating()
    On Error GoTo ErrH
    Call 更新ID
    mfrm评分结果编辑.ShowForm "新增", m_lng结果ID, m_lng病人ID, m_lng主页ID, m_lng方案ID, Val(txt科室.Tag)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 修改评分
'==============================================================================
Private Sub RecordEdit()
    
    On Error GoTo ErrH

    Call 更新ID
    mfrm评分结果编辑.ShowForm "修改", m_lng结果ID, m_lng病人ID, m_lng主页ID, m_lng方案ID, Val(txt科室.Tag)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 重新评分
'==============================================================================
Private Sub RecordReturn()
 
    On Error GoTo ErrH

    Call 更新ID
    mfrm评分结果编辑.ShowForm "重评", m_lng结果ID, m_lng病人ID, m_lng主页ID, m_lng方案ID, Val(txt科室.Tag)
 
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 删除评分
'==============================================================================
Private Sub RecordDel()
    Dim msgReturn       As VbMsgBoxResult '保存对话框返回值
    
    On Error GoTo ErrH
    
    Call 更新ID
    msgReturn = MsgBox("你确认要删除" & fg病案_S.Cell(flexcpText, fg病案_S.Row, 5) & "号病人的第(" & fg病案_S.Cell(flexcpText, fg病案_S.Row, 6) & ")次住院病案评分结果记录？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName)
    If msgReturn = vbNo Then Exit Sub
    gstrSQL = "ZL_病案评分结果_Delete (" & m_lng结果ID & ")"
    '注意：此处使用了事务处理。
    gcnOracle.BeginTrans
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    '若病人列表选择项为空或主页Tab选择项为空，则退出
    If fg病案_S.Row < 1 Then Exit Sub
    '提交事务
    gcnOracle.CommitTrans
    '刷新主页TAB
    Call 更新ID
    Call Fill评分结果
    Exit Sub
ErrH:
    '回滚事务
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 查看首页
'==============================================================================
Private Sub RecordLook()
    
    On Error GoTo ErrH
                                                    
    Call 更新ID
    '开始查阅病案首页
    If fg病案_S.Row < 1 Then Exit Sub
    If mfrmArchiveView Is Nothing Then Set mfrmArchiveView = New frmArchiveView
    Call mfrmArchiveView.ShowArchive(Me, m_lng病人ID, m_lng主页ID, False)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 评分审核
'==============================================================================
Private Sub RecordAudit()
    Dim rs              As ADODB.Recordset
    Dim msgReturn       As VbMsgBoxResult '保存对话框返回值
    
    On Error GoTo ErrH
    
    Call 更新ID
    msgReturn = MsgBox("请确认通过评分审核：" & fg病案_S.Cell(flexcpText, fg病案_S.Row, 5) & "号病人的第(" & fg病案_S.Cell(flexcpText, fg病案_S.Row, 6) & ")次住院病案评分结果记录？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName)
    If msgReturn = vbNo Then Exit Sub
    gstrSQL = "select A.病人ID,A.主页ID,A.信息名,A.信息值,B.等级 from 病案主页从表 A,病案评分结果 B Where A.病人ID=" & m_lng病人ID & " and A.主页ID=" & m_lng主页ID & " and A.信息名='病案质量' " & _
        " and B.病人ID=A.病人ID and B.主页ID=A.主页ID "
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If Not rs.EOF Then
        If rs("信息值") <> rs("等级") Then
            If MsgBox("已编目的病案等级与评分等级不相同，确认保存吗？" + vbCrLf + _
                "----------------------------------------------------------- " + vbCrLf + _
                "前者为: [" + rs("信息值") + "]，后者为: [" + IIf(rs("等级") = "否", "不合格", rs("等级")) + "]", vbOKCancel + vbInformation + vbDefaultButton1, gstrSysName) = vbOK Then
                gstrSQL = "ZL_病案评分结果_审核" & _
                    "(" & m_lng结果ID & ",'" & gstrUserName & "'," & glngSys & ")"
            Else
                Exit Sub
            End If
        Else
            gstrSQL = "ZL_病案评分结果_审核(" & m_lng结果ID & ",'" & gstrUserName & "'," & glngSys & ")"
        End If
    Else
        '等级相同，则直接审核通过！
        gstrSQL = "ZL_病案评分结果_审核" & _
            "(" & m_lng结果ID & ",'" & gstrUserName & "'," & glngSys & ")"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    '刷新主页TAB
    Call 更新ID
    Call Fill评分结果
    Call SetMenu
    
    Exit Sub
ErrH:
    If gcnOracle.Errors.count > 0 Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 取消审查
'==============================================================================
Private Sub RecordUnAudit()
    Dim msgReturn As VbMsgBoxResult '保存对话框返回值
On Error GoTo ErrH
    Call 更新ID
    msgReturn = MsgBox("请确认取消评分审核：" & fg病案_S.Cell(flexcpText, fg病案_S.Row, 5) & "号病人的第(" & fg病案_S.Cell(flexcpText, fg病案_S.Row, 6) & ")次住院病案评分结果记录？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName)
    If msgReturn = vbNo Then Exit Sub
    gstrSQL = "ZL_病案评分结果_取消审核" & "(" & m_lng结果ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    '刷新主页TAB
    Call 更新ID
    Call Fill评分结果
    Call SetMenu
    Exit Sub
ErrH:
    If gcnOracle.Errors.count > 0 Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 全部选中
'==============================================================================
Private Sub RecordSelect()
    On Error GoTo ErrH
    
    fg病案_S.Cell(flexcpChecked, 0, 0, fg病案_S.Rows - 1, 0) = flexChecked
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 全部清除
'==============================================================================
Private Sub RecordUnSelect()
    
    On Error GoTo ErrH
    
    fg病案_S.Cell(flexcpChecked, 0, 0, fg病案_S.Rows - 1, 0) = flexUnchecked
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 反向选择
'==============================================================================
Private Sub RecordSelectOther()
    Dim i As Long

    On Error GoTo ErrH
    
    fg病案_S.Cell(flexcpChecked, 0, 0) = flexUnchecked
    For i = 1 To fg病案_S.Rows - 1
        If fg病案_S.Cell(flexcpChecked, i, 0) = flexUnchecked Then
            fg病案_S.Cell(flexcpChecked, i, 0) = flexChecked
        Else
            fg病案_S.Cell(flexcpChecked, i, 0) = flexUnchecked
        End If
    Next
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 过滤查询
'==============================================================================
Private Sub RecordFind()
    Dim strTemp         As String
    
    On Error GoTo ErrH
    With mfrm查找
    
        strTemp = .GetFilter(mstrPrivs, txt科室.Text)
        If .mblnCancel Then Exit Sub
        mlngSickID = .lngSickID                         '病人ID
        mlngHospitalID = .lngHospitalID                 '住院号
        mlngHospitalTimes = .lngHospitalTimes           '住院次数
        mstrSickName = .strSickName                     '病人姓名
        mstrMainDoctor = .strMainDoctor                 '主治医师
        mstrOutpatientDoctor = .strOutpatientDoctor     '门诊医师
        mstrNurses = .strNurses                         '责任护士
        mstrRatingMan = .strRatingMan                   '评分人
        mstrAuditMan = .strAuditMan                     '审核人
        mstrOutDept = .strOutDept                       '出院科室
        mstrInDept = .strInDept                         '入院科室
        mdatStarOutDate = .datStarOutDate               '出院开始日期
        mdatEndOutDate = .datEndOutDate                 '出院开始日期
        mdatStarInDate = .datStarInDate                 '入院开始日期
        mdatEndInDate = .datEndInDate                   '入院开始日期
        mstrSickType = .strSickType                     '病理类型
    End With
    mstrWhere = "Where 1=1 " & strTemp
    Call mDataLoad
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 批量打印
'==============================================================================
Private Sub RecordAllPrint()
    Dim i               As Long
    Dim rs              As ADODB.Recordset
    Dim lngID           As Long
    Dim lngNum          As Long
    On Error GoTo ErrH

    If MsgBox("是否打印当前选中的所有病案评分结果报表？", vbOKCancel + vbInformation + vbDefaultButton2, gstrSysName) = vbCancel Then Exit Sub
    lngNum = 0
    For i = 1 To fg病案_S.Rows - 1
        If fg病案_S.Cell(flexcpChecked, i, 0) = flexChecked Then
            lngID = Val(fg病案_S.Cell(flexcpText, i, 1))
            If lngID <> 0 Then
                lngNum = lngNum + 1
                stbThis.Panels(2) = "打印进度:" & CStr(lngNum)
                ReportOpen gcnOracle, glngSys, "ZL1_INSIDE_1562_1", Me, "结果ID=" & lngID, 2
            End If
        End If
    Next i
    stbThis.Panels(2) = "打印任务发送完毕，共发送" & CStr(lngNum) & "份。"
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：条件变化时重新加载数据
'==============================================================================
Private Sub chk结果_Click(Index As Integer)
    On Error GoTo ErrH
    Call mDataLoad
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：chk结果 回车相当于Tab键
'==============================================================================
Private Sub chk结果_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo ErrH
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：cmb范围 回车相当于Tab键
'==============================================================================
Private Sub cmb范围_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrH
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：加载科室选择数据
'==============================================================================
Private Sub DeptSelect()
    Dim i As Integer
    On Error GoTo ErrH
    If tvw科室.Nodes.count = 0 Then
        Exit Sub
    End If
    tvw科室.Visible = True
    tvw科室.ZOrder (0)
    If tvw科室.Visible Then
        '显示当前人员
        If txt科室.Tag = "" Then
            tvw科室.Nodes(1).Expanded = True
            tvw科室.Nodes(1).Selected = True
        Else
            tvw科室.Nodes("C" & txt科室.Tag).Selected = True
            tvw科室.SelectedItem.EnsureVisible
        End If
        tvw科室.SetFocus
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：右键菜单（评分快捷菜单）
'==============================================================================
Private Sub fg病案_S_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrH
    If Button = 2 Then
        If fg病案_S.MouseRow = -1 And fg病案_S.Rows >= 1 Then
            fg病案_S.Row = fg病案_S.Rows - 1
        ElseIf fg病案_S.MouseRow = 0 And fg病案_S.Rows > 1 Then
            fg病案_S.Row = 1
        Else
            fg病案_S.Row = fg病案_S.MouseRow
        End If
        fg病案_S.Col = fg病案_S.MouseCol
        mcbrPopupBar.ShowPopup
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：首列不能拖动位置
'==============================================================================
Private Sub fg病案_S_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    On Error GoTo ErrH
    If Col = 0 Then
        Position = -1
    Else
        If Position <= 0 Then Position = Col
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 某列不能拖动大小 fg病案_S[图标]
'==============================================================================
Private Sub fg病案_S_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    On Error GoTo ErrH
    If Col = 0 Then Cancel = True
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：fg病案 行列变动后更新状态栏
'==============================================================================
Private Sub fg病案_S_RowColChange()
    On Error GoTo ErrH
    Call fg病案_S_SelChange
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：fg病案 选择列变动后更新状态栏
'==============================================================================
Private Sub fg病案_S_SelChange()
    Dim strTag As String
    On Error GoTo ErrH
    If fg病案_S.Rows < 1 Then
        fg结果_S.Rows = 1
        lbl总分.Caption = "总分:"
        lbl等级.Caption = "等级:"
        lbl评分人.Caption = "评分人:"
        lbl审核人.Caption = "审核人:"
        lbl返回修改.Caption = ""
        lbl病理类型.Caption = "病理类型:"
        lbl备注.Caption = "备注:"
        lbl评分时间.Caption = "评分时间:"
        lbl审核时间.Caption = "审核时间:"
        Exit Sub
    End If
    strTag = fg病案_S.Cell(flexcpText, fg病案_S.Row, 4) & "_" & fg病案_S.Cell(flexcpText, fg病案_S.Row, 6)
    If fg病案_S.Tag <> strTag Then
        fg病案_S.Tag = strTag
        Call Fill评分结果
    End If
    m_lngOldRow = fg病案_S.Row
    mRecordRating = (fg病案_S.TextMatrix(fg病案_S.Row, fg病案_S.ColIndex("评分时间")) <> "")
    mRecordAudit = (fg病案_S.TextMatrix(fg病案_S.Row, fg病案_S.ColIndex("审核时间")) <> "")
    mRecordMyAudit = (fg病案_S.TextMatrix(fg病案_S.Row, fg病案_S.ColIndex("评分人")) = UserInfo.姓名)
    gstrSQL = "select Count(1) from 病案评分方案 where 选用=1 and ID = [1]"
    mRecordReturn = (zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(fg病案_S.TextMatrix(fg病案_S.Row, fg病案_S.ColIndex("方案ID")))).Fields(0) > 0)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：全选或全清
'==============================================================================
Private Sub fg病案_S_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    On Error GoTo ErrH
    
    If Col = 0 And Row = 0 Then
        If fg病案_S.Cell(flexcpChecked, 0, 0, 0, 0) = flexUnchecked Then
            fg病案_S.Cell(flexcpChecked, 0, 0, fg病案_S.Rows - 1, 0) = flexChecked
        Else
            fg病案_S.Cell(flexcpChecked, 0, 0, fg病案_S.Rows - 1, 0) = flexUnchecked
        End If
        Cancel = True
    ElseIf Col <> 0 Then
        Cancel = True
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：右键菜单（评分快捷菜单）
'==============================================================================
Private Sub fg结果_S_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrH
    If Button = 2 Then
        If fg结果_S.MouseRow = -1 And fg结果_S.Rows >= 1 Then
            fg结果_S.Row = fg结果_S.Rows - 1
        ElseIf fg结果_S.MouseRow = 0 And fg结果_S.Rows > 1 Then
            fg结果_S.Row = 1
        Else
            fg结果_S.Row = fg结果_S.MouseRow
        End If
        fg结果_S.Col = fg结果_S.MouseCol
        mcbrPopupBar.ShowPopup
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：fg列表 分数值的判断
'==============================================================================
Private Sub fg列表_Click()
    Dim i           As Long
    On Error GoTo ErrH
    For i = 1 To fg列表.Rows - 1
        If fg列表.Cell(flexcpChecked, i, 0) = flexChecked Then
            If fg病案_S.ColWidth(fg病案_S.ColIndex(fg列表.Cell(flexcpText, i, 1))) < 100 Then
                
                If fg病案_S.ColIndex(fg列表.Cell(flexcpText, i, 1)) = 3 Then
                    fg病案_S.ColWidth(fg病案_S.ColIndex(fg列表.Cell(flexcpText, i, 1))) = 300
                Else
                    fg病案_S.ColWidth(fg病案_S.ColIndex(fg列表.Cell(flexcpText, i, 1))) = 1000
                End If
            End If
        Else
            If fg病案_S.ColWidth(fg病案_S.ColIndex(fg列表.Cell(flexcpText, i, 1))) > 100 Then
                fg病案_S.ColWidth(fg病案_S.ColIndex(fg列表.Cell(flexcpText, i, 1))) = 0
            End If
        End If
    Next
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：fg列表 按ESC时隐藏列表
'==============================================================================
Private Sub fg列表_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrH
    If KeyAscii = 27 Then fg列表.Visible = False
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：fg列表 按ESC时隐藏列表
'==============================================================================
Private Sub fg列表_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    On Error GoTo ErrH
    If Col <> 0 Then
        Cancel = True
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：页面控件初始化
'==============================================================================
Private Sub Form_Initialize()
    On Error GoTo ErrH
    Call InitCommonControls
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：页面初始化
'==============================================================================
Private Sub Form_Load()
    Dim strKey              As String
    On Error GoTo ErrH
    m_lngOldRow = -1
    mstrPrivs = UserInfo.模块权限
    mlngModule = ParamInfo.模块号
    mblnSetDept = False
    Set mfrm评分结果编辑 = New frm评分结果编辑
    '控件初始化
    mstrFindKey = "住院号"
    If GetPersonSet Then
        mstrFindKey = Trim(GetPara("定位范围", mlngModule, "住院号", True))
        chk结果(0).Value = zlDatabase.GetPara("未评分", glngSys, mlngModule, vbChecked)
        chk结果(1).Value = zlDatabase.GetPara("未审核", glngSys, mlngModule, vbChecked)
        chk结果(2).Value = zlDatabase.GetPara("已审核", glngSys, mlngModule, vbChecked)
    End If
                                       
    Call InitControl
    
    '初始化查找窗体
    Set mfrm查找 = New frm病案评分查询
    Load mfrm查找
    mdatStarOutDate = DateAdd("M", -1, Date)
    If mfrm查找.mbln编目后评分 Then
        mstrWhere = " Where 编目日期 is not null and 出院日期 >= [12]"
    Else
        mstrWhere = " Where 出院日期 >= [12]"
    End If
    If IsPrivs(mstrPrivs, "所有科室") Then
        txt科室.Text = "所有科室"
        Call InitTvw
    Else
        
        txt科室.Text = Get所属部门(UserInfo.ID, 0)
        txt科室.Locked = True
        txt科室.BackColor = &H80000000
        mstrOutDept = UserInfo.部门名称
        
        If txt科室.Text <> mstrOutDept Then
            mstrWhere = mstrWhere & " And 出院科室 In (" & Get所属部门(UserInfo.ID, 1) & ")"
        Else
            mstrWhere = mstrWhere & " And 出院科室 = [11]"
        End If
    End If
    Call mDataLoad
    stbThis.Panels(2) = "当前显示有" & fg病案_S.Rows - 1 & "份病案。"
    strKey = Me.fg病案_S.Tag
    Me.fg病案_S.Tag = ""
    RestoreWinState Me, App.ProductName
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs, "ZL1_INSIDE_1562_1")
    fg病案_S.Tag = strKey
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：窗口关闭时关闭子窗口
'==============================================================================
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo ErrH
    If Not (mfrmArchiveView Is Nothing) Then Unload mfrmArchiveView
    Set mfrmArchiveView = Nothing
    mfrm查找.mblnForce = True   '强制关闭（正式关闭，否则只是隐藏之）
    Unload mfrm查找
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：窗口大小变动时控件变化
'==============================================================================
Private Sub Form_Resize()
    On Error GoTo ErrH

    Call SetPaneRange(dkpMain, 1, 100, 100, ScaleHeight - 200, ScaleHeight)
    Call SetPaneRange(dkpMain, 2, 400, 100, ScaleHeight - 200, ScaleHeight)
    With tvw科室
        .Move txt科室.Left, stbThis.Height * 2 + txt科室.Top + txt科室.Height + 70, txt科室.Width, 4000
    End With
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：页面关闭时参数保存
'==============================================================================
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrH
    If Not (mfrm评分结果编辑 Is Nothing) Then Unload mfrm评分结果编辑
    If Not (mfrmArchiveView Is Nothing) Then Unload mfrmArchiveView
    Set mfrmArchiveView = Nothing
    Call SetPara("未评分", chk结果(0).Value, mlngModule)
    Call SetPara("未审核", chk结果(1).Value, mlngModule)
    Call SetPara("已审核", chk结果(2).Value, mlngModule)
    Call SetPara("定位范围", mstrFindKey, mlngModule)
    Me.fg病案_S.Tag = ""
    SaveWinState Me, App.ProductName
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：fg列表 显示字段列表
'==============================================================================
Private Sub imgSelCols_S_Click()
    On Error GoTo ErrH
    Call InitTvw
    fg列表.Visible = Not fg列表.Visible
    If fg列表.Visible Then fg列表.SetFocus
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：fg列表 失去焦点时隐藏
'==============================================================================
Private Sub fg列表_LostFocus()
    On Error GoTo ErrH
    fg列表.Visible = False
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：批量打印
'==============================================================================
Private Sub mnuFilePrintALL_Click()
    Dim i               As Long
    Dim rs              As ADODB.Recordset
    Dim lngID           As Long
    Dim lngNum          As Long
    On Error GoTo ErrH
    If MsgBox("是否打印当前选中的所有病案评分结果报表？", vbOKCancel + vbInformation + vbDefaultButton2, gstrSysName) = vbCancel Then Exit Sub
    lngNum = 0
    For i = 1 To fg病案_S.Rows - 1
        If fg病案_S.Cell(flexcpChecked, i, 0) = flexChecked Then
            lngID = Val(fg病案_S.Cell(flexcpText, i, 1))
            If lngID <> 0 Then
                lngNum = lngNum + 1
                stbThis.Panels(2) = "打印进度:" & CStr(lngNum)
                ReportOpen gcnOracle, glngSys, "ZL1_INSIDE_1562_1", Me, "结果ID=" & lngID, 2
            End If
        End If
    Next i
    stbThis.Panels(2) = "打印任务发送完毕，共发送" & CStr(lngNum) & "份。"
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：左侧位置位置变化界面调整
'==============================================================================
Private Sub picLeft_S_Resize()
On Error Resume Next
    imgSelCols_S.Move IIf(imgSelCols_S.Left < 1350, 1350, picLeft_S.ScaleWidth - imgSelCols_S.Width - 100)
    fg列表.Move Abs(picLeft_S.Width - fg列表.Width), picLeft_S.Top + imgSelCols_S.Top + imgSelCols_S.Height + 45
    imgRefresh.Move imgSelCols_S.Left - imgRefresh.Width - 175
    picLeft_S.Cls
    picLeft_S.PaintPicture imgBGBlue.Picture, Screen.TwipsPerPixelX, 0, picLeft_S.Width, 360, 0, 0, imgBGBlue.Width, 360
    picLeft_S.PaintPicture imgBGBlue.Picture, Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, picLeft_S.Height - 360, 0, 360, Screen.TwipsPerPixelX, imgBGBlue.Height - 360
    picLeft_S.PaintPicture imgBGBlue.Picture, picLeft_S.ScaleWidth - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, picLeft_S.Height - 360, imgBGBlue.Width - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, imgBGBlue.Height - 360
    picLeft_S.PaintPicture imgBGBlue.Picture, Screen.TwipsPerPixelX, picLeft_S.ScaleHeight - Screen.TwipsPerPixelY, picLeft_S.Width, Screen.TwipsPerPixelY, 0, imgBGBlue.Height - Screen.TwipsPerPixelY, imgBGBlue.Width, Screen.TwipsPerPixelY
    fg病案_S.Move picLeft_S.Left + 40, fg病案_S.Top, picLeft_S.Width - 60, picLeft_S.Height - fg病案_S.Top - 420
    Refresh
End Sub

'==============================================================================
'=功能：右击时弹出菜单
'==============================================================================
Private Sub picRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrH
    If Button = 2 Then mcbrPopupBar.ShowPopup
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：右侧控件位置调整
'==============================================================================
Private Sub picRight_Resize()
    On Error GoTo ErrH
    fg结果_S.Move fg结果_S.Left, fg结果_S.Top, picRight.Width - 400, picRight.Height - fg结果_S.Top - 450
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：病案数据加载
'==============================================================================
Private Sub mDataLoad()
    Dim strWhereDept        As String
    Dim strWhere            As String
    Dim blnA                As Boolean
    Dim blnB                As Boolean
    Dim blnC                As Boolean
    Dim i                   As Integer
    
    On Error GoTo ErrH
    
    strWhereDept = "A.出院日期 is not null"
    blnA = (chk结果(0).Value = vbChecked)
    blnB = (chk结果(1).Value = vbChecked)
    blnC = (chk结果(2).Value = vbChecked)
    
    strWhere = "("
    If blnA Then
        strWhere = strWhere & "A.评分时间 is null"
    End If
    If blnB Then
        If strWhere = "(" Then
            strWhere = strWhere & "A.审核时间 is null"
        Else
            strWhere = strWhere & " or A.审核时间 is null"
        End If
    End If
    If blnC Then
        If strWhere = "(" Then
            strWhere = strWhere & "A.审核时间 is not null"
        Else
            strWhere = strWhere & " or A.审核时间 is not null"
        End If
    End If
 
    If strWhere <> "(" Then
        strWhere = strWhereDept & " And " & strWhere & ")"
    Else
        strWhere = strWhereDept
    End If
    If Trim(mstrWhere) = "" Then mstrWhere = "1=1"
    strWhere = IIf(InStr(LCase(mstrWhere), "where") > 0, mstrWhere, " where " & mstrWhere) & " And " & strWhere
    
    gstrSQL = "" & _
        "   Select A.住院号, A.姓名, A.性别,  Decode((select Count(*) from 病人临床路径 where 病人ID = A.病人ID and 主页id = A.主页ID),0,'','lujin') as 路径,A.病人id, A.主页id, A.入院日期, A.出院日期, A.入院科室, A.出院科室, A.门诊医师, A.责任护士, A.住院医师," & _
        "           A.编目日期, A.结果id, A.方案id, A.总分, A.等级, A.评分人, A.评分时间, A.审核人, A.审核时间, A.返回修改, A.备注,A.病理类型 " & _
        "   from 病案质量报表视图 A " & strWhere

    Set rsM = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngSickID, mlngHospitalID, mlngHospitalTimes, mstrSickName, mstrMainDoctor, mstrOutpatientDoctor, mstrNurses, mstrRatingMan, mstrInDept, mstrAuditMan, mstrOutDept, CDate(mdatStarOutDate), CDate(mdatEndOutDate), CDate(mdatStarInDate), CDate(mdatEndInDate), mstrSickType)
    rsM.Sort = "审核时间 desc,评分时间 desc,出院日期 desc,住院医师,姓名"
    
    If (txt科室.Text = "所有科室" Or txt科室.Text = "") And Not mblnSetDept Then
        rsM.Filter = ""
    Else
        If txt科室.Text = UserInfo.部门名称 Then
            rsM.Filter = "出院科室='" & txt科室.Text & "'"
        End If
    End If
    
    Call Fill病案
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：加载评分相关信息
'==============================================================================
Private Sub Fill病案()
    Dim lngIndex            As Long
    Dim lng评分状态         As Long
    Dim sngForeColor        As ColorConstants
    Dim bln提醒             As Boolean
    Dim str等级             As String
    Dim i                   As Long
    Dim j                   As Long
    
    On Error GoTo ErrH
    
    With fg病案_S
        .Editable = flexEDKbdMouse
        .Redraw = flexRDNone
        .Rows = 1
        .Cols = 26
        .Clear
        
        .Cell(flexcpText, 0, 0) = ""
        .Cell(flexcpText, 0, 1) = "结果ID"
        .Cell(flexcpText, 0, 2) = "方案ID"
        .Cell(flexcpPicture, 0, 3) = imgLujinPic.Picture
        .Cell(flexcpText, 0, 4) = "病人ID"
        .Cell(flexcpText, 0, 5) = "住院号"
        .Cell(flexcpText, 0, 6) = "住院次数"
        .Cell(flexcpText, 0, 7) = "姓名"
        .Cell(flexcpText, 0, 8) = "性别"
        .Cell(flexcpText, 0, 9) = "住院医师"
        .Cell(flexcpText, 0, 10) = "门诊医师"
        .Cell(flexcpText, 0, 11) = "责任护士"
        .Cell(flexcpText, 0, 12) = "出院科室"
        .Cell(flexcpText, 0, 13) = "出院日期"
        .Cell(flexcpText, 0, 14) = "入院科室"
        .Cell(flexcpText, 0, 15) = "入院日期"
        .Cell(flexcpText, 0, 16) = "编目日期"
        .Cell(flexcpText, 0, 17) = "评分人"
        .Cell(flexcpText, 0, 18) = "评分时间"
        .Cell(flexcpText, 0, 19) = "审核人"
        .Cell(flexcpText, 0, 20) = "审核时间"
        .Cell(flexcpText, 0, 21) = "总分"
        .Cell(flexcpText, 0, 22) = "等级"
        .Cell(flexcpText, 0, 23) = "返回修改"
        .Cell(flexcpText, 0, 24) = "备注"
        .Cell(flexcpText, 0, 25) = "病理类型"
        DoEvents
        .FocusRect = flexFocusSolid
        '数据填入
        .Rows = IIf(rsM.RecordCount < 1000, rsM.RecordCount + 1, 1001)
        i = 1
        Do Until rsM.EOF
        
            If i >= 1001 And bln提醒 = False Then
                If MsgBox("已经装入1000份病案，还有" & rsM.RecordCount - .Rows + 1 & "份待装。" & vbCrLf & _
                    "是否继续？", vbQuestion Or vbYesNo, gstrSysName) = vbNo Then
                    
                    Exit Do
                End If
                .Rows = rsM.RecordCount + 1
                bln提醒 = True
            End If
            
            If Trim(rsM("审核时间")) <> "" Then
                lng评分状态 = 2
                sngForeColor = RGB(180, 180, 180)
            ElseIf Trim(rsM("评分时间")) <> "" Then
                lng评分状态 = 1
                sngForeColor = RGB(0, 0, 255)
            Else
                lng评分状态 = 0
                sngForeColor = vbBlack
            End If
            
            .Cell(flexcpText, i, 1) = NVL(rsM("结果ID"), 0)
            .Cell(flexcpText, i, 2) = NVL(rsM("方案ID"), 0)
            .Cell(flexcpPicture, i, 3) = IIf(NVL(rsM("路径")) = "", "", imgLujin.Picture)
            .Cell(flexcpText, i, 4) = NVL(rsM("病人ID"), 0)
            
            .Cell(flexcpText, i, 5) = NVL(rsM("住院号"), 0)
            .Cell(flexcpText, i, 6) = NVL(rsM("主页ID"))
            .Cell(flexcpText, i, 7) = NVL(rsM("姓名"))
            .Cell(flexcpText, i, 8) = NVL(rsM("性别"))
            .Cell(flexcpText, i, 9) = NVL(rsM("住院医师"))
            .Cell(flexcpText, i, 10) = NVL(rsM("门诊医师"))
            .Cell(flexcpText, i, 11) = NVL(rsM("责任护士"))
            .Cell(flexcpText, i, 12) = NVL(rsM("出院科室"))
            .Cell(flexcpText, i, 13) = IIf(IsNull(rsM("出院日期")), "", Format(rsM("出院日期"), "YYYY-MM-DD HH:mm"))
            .Cell(flexcpText, i, 14) = NVL(rsM("入院科室"))
            .Cell(flexcpText, i, 15) = IIf(IsNull(rsM("入院日期")), "", Format(rsM("入院日期"), "YYYY-MM-DD HH:mm"))
            .Cell(flexcpText, i, 16) = IIf(IsNull(rsM("编目日期")), "", Format(rsM("编目日期"), "YYYY-MM-DD HH:mm"))
            .Cell(flexcpText, i, 17) = NVL(rsM("评分人"))
            .Cell(flexcpText, i, 18) = IIf(IsNull(rsM("评分时间")), "", Format(rsM("评分时间"), "YYYY-MM-DD HH:mm"))
            .Cell(flexcpText, i, 19) = NVL(rsM("审核人"))
            .Cell(flexcpText, i, 20) = IIf(IsNull(rsM("审核时间")), "", Format(rsM("审核时间"), "YYYY-MM-DD HH:mm"))
            .Cell(flexcpText, i, 21) = IIf(NVL(rsM("等级")) = "否", "", NVL(rsM("总分")))
            .Cell(flexcpText, i, 24) = NVL(rsM!备注)
            .Cell(flexcpText, i, 25) = NVL(rsM!病理类型)
            Select Case NVL(rsM("等级"))
                Case "甲"
                    str等级 = "甲级"
                Case "乙"
                    str等级 = "乙级"
                Case "丙"
                    str等级 = "丙级"
                Case "否"
                    str等级 = "不合格"
                Case Else
                    str等级 = ""
            End Select
            .Cell(flexcpText, i, 22) = str等级
            If NVL(rsM("返回修改"), 0) = 0 Then
                .Cell(flexcpText, i, 23) = ""
            Else
                .Cell(flexcpText, i, 23) = "√"
            End If
             
            For j = 1 To 25
                .Cell(flexcpForeColor, i, j) = sngForeColor
            Next
            
            rsM.MoveNext
            i = i + 1
        Loop
        .Cell(flexcpChecked, 0, 0, .Rows - 1, 0) = flexUnchecked
            
        If Me.Tag = "" Then
            .ColWidth(.ColIndex("ICON")) = 300
            .ColWidth(.ColIndex("结果ID")) = 0
            .ColWidth(.ColIndex("方案ID")) = 0
            .ColWidth(.ColIndex("路径图标")) = 300
            
            .ColWidth(.ColIndex("病人ID")) = 0
            .ColWidth(.ColIndex("住院次数")) = 400
            .ColWidth(.ColIndex("住院号")) = 650
            .ColWidth(.ColIndex("姓名")) = 900
            .ColWidth(.ColIndex("性别")) = 600
            .ColWidth(.ColIndex("住院医师")) = 900
            .ColWidth(.ColIndex("门诊医师")) = 0
            .ColWidth(.ColIndex("责任护士")) = 0
            .ColWidth(.ColIndex("出院科室")) = 900
            .ColWidth(.ColIndex("出院日期")) = 1600
            .ColWidth(.ColIndex("入院科室")) = 0
            .ColWidth(.ColIndex("入院日期")) = 0
            .ColWidth(.ColIndex("编目日期")) = 0
            .ColWidth(.ColIndex("评分人")) = 0
            .ColWidth(.ColIndex("评分时间")) = 0
            .ColWidth(.ColIndex("审核人")) = 0
            .ColWidth(.ColIndex("审核时间")) = 0
            .ColWidth(.ColIndex("总分")) = 0
            .ColWidth(.ColIndex("等级")) = 0
            .ColWidth(.ColIndex("返回修改")) = 0
            .ColWidth(.ColIndex("备注")) = 0
            .ColWidth(.ColIndex("病理类型")) = 0
            .ColAlignment(.ColIndex("返回修改")) = flexAlignCenterCenter
            Me.Tag = "已经调整列宽"
        End If
        
        .ColAlignment(.ColIndex("性别")) = flexAlignCenterCenter
        '行高设置
        .RowHeightMin = 300
        .AllowBigSelection = False
        .Redraw = flexRDBuffered
        
        '选中先前的行
        If m_lngOldRow > 0 And m_lngOldRow < i Then
            .Row = m_lngOldRow
            .Col = 2
            .ShowCell m_lngOldRow, 2
            On Error Resume Next
            If .Visible = True Then .SetFocus
            Call fg病案_S_SelChange
        ElseIf .Tag = "" And i > 1 And .Rows > 1 Then
            m_lngOldRow = 1
            .Tag = "选中第一行"
            .Row = 1
            .Col = 2
            .ShowCell m_lngOldRow, 2
            If .Visible = True Then .SetFocus
            Call fg病案_S_SelChange
        Else
            If .Rows > 1 Then .Row = 1
            Call fg病案_S_SelChange
        End If
     End With

    Call SetMenu
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：取得当前评分结果ID和方案ID
'==============================================================================
Private Sub 更新ID()
    Dim i                   As Long
    Dim sngForeColor        As ColorConstants
    Dim lng评分状态         As Long
    On Error GoTo ErrH
    With fg病案_S
        m_lng病人ID = Val(.Cell(flexcpText, .Row, 4))
        m_lng主页ID = Val(.Cell(flexcpText, .Row, 6))
        
        Dim rs As New ADODB.Recordset, str等级 As String
        gstrSQL = " " & _
            "   Select 住院号, 姓名, 性别, 病人id, 主页id, 入院日期, 出院日期, 入院科室, 出院科室, 门诊医师, 责任护士, 住院医师," & _
            "           编目日期, 结果id, 方案id, 总分, 等级, 评分人, 评分时间, 审核人, 审核时间, 返回修改, 备注,病理类型 " & _
            "   From 病案质量报表视图 " & _
            "   where 病人ID=[1] and 主页ID=[2]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lng病人ID, m_lng主页ID)
        
        If Not rs.EOF Then
            If Trim(rs("审核时间")) <> "" Then
                lng评分状态 = 2
                sngForeColor = RGB(180, 180, 180)
            ElseIf Trim(rs("评分时间")) <> "" Then
                lng评分状态 = 1
                sngForeColor = RGB(0, 0, 255)
            Else
                lng评分状态 = 0
                sngForeColor = vbBlack
            End If
            
            For i = 1 To 21
                .Cell(flexcpForeColor, .Row, i) = sngForeColor
            Next
            
            .Cell(flexcpText, .Row, 1) = NVL(rs("结果ID"))
            .Cell(flexcpText, .Row, 2) = NVL(rs("方案ID"))
            .Cell(flexcpText, .Row, 17) = NVL(rs("评分人"))
            .Cell(flexcpText, .Row, 18) = NVL(rs("评分时间"))
            .Cell(flexcpText, .Row, 19) = NVL(rs("审核人"))
            .Cell(flexcpText, .Row, 20) = NVL(rs("审核时间"))
            .Cell(flexcpText, .Row, 21) = IIf(NVL(rs("等级")) = "否", "", NVL(rs("总分")))
            .Cell(flexcpText, .Row, 23) = IIf(NVL(rs("返回修改"), 0) = 0, "", "√")
            .Cell(flexcpText, .Row, 24) = NVL(rs("备注"))
            .Cell(flexcpText, .Row, 25) = NVL(rs("病理类型"))
            
            Select Case NVL(rs("等级"))
                Case "甲"
                    str等级 = "甲级"
                Case "乙"
                    str等级 = "乙级"
                Case "丙"
                    str等级 = "丙级"
                Case "否"
                    str等级 = "不合格"
                Case Else
                    str等级 = ""
            End Select
            .Cell(flexcpText, .Row, 22) = str等级
            '颜色控制
        End If
        rs.Close
        m_lng结果ID = Val(.Cell(flexcpText, .Row, 1))
        m_lng方案ID = Val(.Cell(flexcpText, .Row, 2))
    End With
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能:将数据表进行打印,预览和输出到EXCEL
'=参数:bytMode=1 打印;2 预览;3 输出到EXCEL
'==============================================================================
Private Sub subPrint(ByVal bytMode As Byte)
    On Error GoTo ErrH
    Select Case bytMode
        Case 1  'Print
            ReportOpen gcnOracle, glngSys, "ZL1_INSIDE_1562_1", Me, "结果ID=" & m_lng结果ID, 2
        Case 2  'Preview
            ReportOpen gcnOracle, glngSys, "ZL1_INSIDE_1562_1", Me, "结果ID=" & m_lng结果ID, 1
        Case 3  'Excel
            ReportOpen gcnOracle, glngSys, "ZL1_INSIDE_1562_1", Me, "结果ID=" & m_lng结果ID, 3
    End Select
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：根据选择的对象设置当前菜单项，屏蔽不需要的菜单。
'==============================================================================
Public Sub SetMenu()
    On Error GoTo ErrH
    Call 更新ID
    If Trim(fg病案_S.Cell(flexcpText, fg病案_S.Row, 18)) <> "" Then
        mRecordRating = True
    End If
    If Trim(fg病案_S.Cell(flexcpText, fg病案_S.Row, 20)) <> "" Then
        mRecordAudit = True
    End If
    stbThis.Panels(2) = "当前显示有" & fg病案_S.Rows - 1 & "份病案。"
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：根据评分方案ID，填入评分标准网格。
'==============================================================================
Private Sub Fill评分标准()
    Dim rsTemp          As ADODB.Recordset
    Dim lngIndex        As Long
    Dim lng可否修改     As Long
    Dim i               As Long
    On Error GoTo ErrH
    Call 更新ID
    With fg结果_S
        .Redraw = flexRDNone
        .Rows = 1
        .Clear
        .Cell(flexcpText, 0, 0) = "项目"
        .Cell(flexcpAlignment, 0, 0) = flexAlignCenterCenter
        .Cell(flexcpText, 0, 1) = "标准分值"
        .Cell(flexcpAlignment, 0, 1) = flexAlignCenterCenter
        .Cell(flexcpText, 0, 2) = "缺陷内容"
        .Cell(flexcpAlignment, 0, 2) = flexAlignCenterCenter
        .Cell(flexcpText, 0, 3) = "评分标准"
        .Cell(flexcpAlignment, 0, 3) = flexAlignCenterCenter
        .Cell(flexcpText, 0, 4) = "评分"
        .Cell(flexcpAlignment, 0, 4) = flexAlignCenterCenter
        .Cell(flexcpText, 0, 5) = "可否修改"
        .Cell(flexcpText, 0, 6) = "ID"
        .Cell(flexcpText, 0, 7) = "上级ID"
        .Cell(flexcpText, 0, 8) = "方案ID"
        .Cell(flexcpText, 0, 9) = "备注"

        
        '确定方案名称
        If m_lng方案ID < 1 Then .Redraw = flexRDDirect: Exit Sub
        gstrSQL = "" & _
            "   Select 上级序号, 序号, Id, 上级id, 方案id, 项目, 标准分值, 基本要求, 缺陷内容, 扣分标准, 隐藏 " & _
            "   From 病案评分标准视图 " & _
            "   Where 隐藏='否' and 方案ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lng方案ID)
        
        .FocusRect = flexFocusSolid
        '数据填入
        .Cols = 10
        .Rows = rsTemp.RecordCount + 1
        i = 1
        Do Until rsTemp.EOF
            .Cell(flexcpText, i, 0) = IIf(IsNull(rsTemp.Fields("项目")), "", rsTemp.Fields("项目"))
            .Cell(flexcpAlignment, i, 0) = flexAlignCenterCenter
            .Cell(flexcpText, i, 1) = IIf(IsNull(rsTemp.Fields("标准分值")), " ", Format(rsTemp.Fields("标准分值"), "####分"))
            .Cell(flexcpAlignment, i, 1) = flexAlignCenterCenter
            .Cell(flexcpText, i, 2) = IIf(IsNull(rsTemp.Fields("缺陷内容")), "", rsTemp.Fields("缺陷内容"))
            .Cell(flexcpAlignment, i, 2) = flexAlignLeftTop
            .Cell(flexcpText, i, 3) = IIf(IsNull(rsTemp.Fields("扣分标准")), "", IIf(rsTemp.Fields("扣分标准") = "甲", "甲级", IIf(rsTemp.Fields("扣分标准") = "乙", "乙级", IIf(rsTemp.Fields("扣分标准") = "丙", "丙级", IIf(rsTemp.Fields("扣分标准") = "否", "单项否决", rsTemp.Fields("扣分标准"))))))
            .Cell(flexcpAlignment, i, 3) = flexAlignCenterCenter
            .Cell(flexcpText, i, 4) = ""
            .Cell(flexcpAlignment, i, 4) = flexAlignCenterCenter
            .Cell(flexcpText, i, 5) = ""
            .Cell(flexcpText, i, 6) = IIf(IsNull(rsTemp.Fields("ID")), "", rsTemp.Fields("ID"))
            .Cell(flexcpText, i, 7) = IIf(IsNull(rsTemp.Fields("上级ID")), "", rsTemp.Fields("上级ID"))
            .Cell(flexcpText, i, 8) = IIf(IsNull(rsTemp.Fields("方案ID")), "", rsTemp.Fields("方案ID"))
            .Cell(flexcpText, i, 9) = ""
            
            rsTemp.MoveNext
            i = i + 1
        Loop
        '自动换行
        .WordWrap = True
        '合并单元格
        .MergeCells = 2
        .MergeCol(.ColIndex("项目")) = True
        .MergeCol(.ColIndex("标准分值")) = True
        '对齐设置
        .ColAlignment(.ColIndex("项目")) = flexAlignLeftCenter
        .ColAlignment(.ColIndex("标准分值")) = flexAlignCenterCenter
        .ColAlignment(.ColIndex("评分标准")) = flexAlignCenterCenter
        .ColAlignment(.ColIndex("评分")) = flexAlignCenterCenter
        .ColAlignment(.ColIndex("可否修改")) = flexAlignCenterCenter
        .ColAlignment(.ColIndex("备注")) = flexAlignLeftCenter
        
        '隐藏单元格
        .ColWidth(.ColIndex("ID")) = 0
        .ColWidth(.ColIndex("上级ID")) = 0
        .ColWidth(.ColIndex("方案ID")) = 0
        '宽度设置
'        .ColWidth(.ColIndex("项目")) = 1500
'        .ColWidth(.ColIndex("标准分值")) = 850
'        .ColWidth(.ColIndex("缺陷内容")) = 3000
'        .ColWidth(.ColIndex("评分标准")) = 1100
'        .ColWidth(.ColIndex("评分")) = 800
'        .ColWidth(.ColIndex("可否修改")) = 800
        '行高设置
        .RowHeightMin = 300
        '最大宽度设置
        .ColWidthMax = 7000
        '自动适应行高、列宽
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize .ColIndex("缺陷内容")
        .AllowBigSelection = False
    End With
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：装入对应主页的评分结果
'==============================================================================
Private Function Fill评分结果() As Boolean
    Dim rs              As ADODB.Recordset
    Dim i               As Long
    Dim bln扣分制       As Boolean
    Dim intSign         As Long

    On Error GoTo ErrH

    Call Fill评分标准
    With fg病案_S
        lbl病人信息 = "姓名:" & .Cell(flexcpText, .Row, 7) & ",第" & .Cell(flexcpText, .Row, 6) & "次住院"
        lbl总分 = "总分:" & .Cell(flexcpText, .Row, 21)
        lbl等级 = "等级:" & .Cell(flexcpText, .Row, 22)
        lbl评分人 = "评分人:" & .Cell(flexcpText, .Row, 17)
        lbl评分时间 = "评分时间:" & .Cell(flexcpText, .Row, 18)
        lbl审核人 = "审核人:" & .Cell(flexcpText, .Row, 19)
        lbl审核时间 = "审核时间:" & .Cell(flexcpText, .Row, 0)
        lbl返回修改 = IIf(.Cell(flexcpText, .Row, 23) = "", "", "【返回修改】")
        lbl备注.Caption = "备注:" & .Cell(flexcpText, .Row, .ColIndex("备注"))
        lbl病理类型.Caption = "病理类型:" & .Cell(flexcpText, .Row, .ColIndex("病理类型"))
    End With
    fg结果_S.Redraw = flexRDNone
    For i = 1 To fg结果_S.Rows - 1
        fg结果_S.Cell(flexcpText, i, 4) = ""
        fg结果_S.Cell(flexcpText, i, 5) = ""
        fg结果_S.Cell(flexcpText, i, 9) = ""
    Next
    '确定分制
    gstrSQL = "select 分制 from 病案评分方案 where ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lng方案ID)
    bln扣分制 = True
    If Not rs.EOF Then
        bln扣分制 = IIf(NVL(rs("分制"), "加分制") = "加分制", False, True)
    End If
    rs.Close
    If bln扣分制 Then
        intSign = -1
    Else
        intSign = 1
    End If
    gstrSQL = "" & _
        "   Select A.ID,A.项目,A.标准分值,A.基本要求,A.缺陷内容,A.扣分标准," & _
        "           (select decode(缺陷等级,null,to_CHAR(单项分数),缺陷等级) from 病案评分明细 where 评分标准ID=A.ID and 主表ID=[1]) as 评分," & _
        "           (select 可否修改 from 病案评分明细 where 评分标准ID=A.ID and 主表ID=[1]) as 可否修改," & _
        "           (select 备注 from 病案评分明细 where 评分标准ID=A.ID and 主表ID=[1]) as 备注 " & _
        "   from 病案评分标准视图 A " & _
        "   where A.隐藏='否' and A.方案ID=(select B.方案ID from 病案评分结果 B where B.ID=[1]) " & _
        "   order by A.上级ID,A.ID "
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lng结果ID)
    If Not rs.EOF Then
        For i = 1 To fg结果_S.Rows - 1
            rs.MoveFirst
            rs.Find "ID=" & Val(fg结果_S.Cell(flexcpText, i, 6))
            If Not rs.EOF Then
                If Not IsNull(rs("评分")) Then
                    Select Case rs("评分")
                    Case "甲", "乙", "丙"
                        fg结果_S.Cell(flexcpText, i, 4) = rs("评分").Value + "级"
                    Case "否"
                        fg结果_S.Cell(flexcpText, i, 4) = "单项否决"
                    Case Else
                        fg结果_S.Cell(flexcpText, i, 4) = IIf(Abs(NVL(rs("评分").Value, 0)) < 1, Format(Abs(NVL(rs("评分").Value, 0)), "0.0"), Abs(NVL(rs("评分").Value, 0)))
                    End Select
                    If intSign = -1 Then
                        fg结果_S.Cell(flexcpForeColor, i, 4) = RGB(255, 0, 0)
                    Else
                        fg结果_S.Cell(flexcpForeColor, i, 4) = RGB(0, 0, 255)
                    End If
                End If
                If Not IsNull(rs("可否修改")) Then
                    If rs("可否修改") = 1 Then
                        fg结果_S.Cell(flexcpText, i, 5) = "√"
                    End If
                End If
                fg结果_S.Cell(flexcpText, i, 9) = NVL(rs!备注)
            End If
        Next
    End If
    fg结果_S.Redraw = flexRDBuffered
    If fg结果_S.Rows <= 1 Then
        '无数据
        fg结果_S.WallPaper = imgBG_fg(0).Picture
    ElseIf Trim(fg病案_S.Cell(flexcpText, fg病案_S.Row, 20)) <> "" Then
        '已审核
        fg结果_S.WallPaper = imgBG_fg(1).Picture
    Else
        '未审核
        fg结果_S.WallPaper = LoadPicture("")
    End If
    Fill评分结果 = True
    Call SetMenu
    Call fg病案_S_SelChange
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Fill评分结果 = False
End Function

'==============================================================================
'=功能：快速查找病案
'==============================================================================
Private Sub 查找病案(范围 As String, strID As String)
    Dim lngBRID             As Long
    Dim lngZYID             As Long
    Dim strSQL              As String
    Dim i                   As Long
    Dim rs                  As ADODB.Recordset
    Dim blnFinded           As Boolean
    Dim lngCurRowTMP        As Long
    On Error GoTo ErrH
    If 范围 = "1-就诊卡号" Then
        strSQL = _
            "Select A.病人ID,B.主页ID " & _
            " From 病人信息 A,病案主页 B " & _
            " Where A.病人ID=B.病人ID " & _
            " And Nvl(B.主页ID,0)<>0 " & _
            " And A.就诊卡号=[1]"
    ElseIf 范围 = "2-病人ID" Then '病人ID
        strSQL = _
            "Select A.病人ID,B.主页ID " & _
            " From 病人信息 A,病案主页 B " & _
            " Where A.病人ID=B.病人ID " & _
            " And Nvl(B.主页ID,0)<>0 " & _
            " And A.病人ID=[1]"
    ElseIf 范围 = "3-住院号" Then '住院号(病人在院)
        strSQL = _
            " Select A.病人ID,B.主页ID " & _
            " From 病人信息 A,病案主页 B " & _
            " Where A.病人ID=B.病人ID  And Nvl(B.主页ID,0)<>0 And B.住院号=[1]"
    ElseIf 范围 = "4-门诊号" Then '门诊号(医技记帐)
        strSQL = _
            " Select A.病人ID,B.主页ID " & _
            " From 病人信息 A,病案主页 B " & _
            " Where A.病人ID=B.病人ID   And Nvl(B.主页ID,0)<>0 And A.门诊号=[1]"
    Else '当作姓名
        strSQL = _
            " Select A.病人ID,B.主页ID " & _
            " From 病人信息 A,病案主页 B " & _
            " Where A.病人ID=B.病人ID  And Nvl(B.主页ID,0)<>0 And Upper(A.姓名)=[1]"
    End If
    gstrSQL = strSQL
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CStr(UCase(strID)))
    If Not rs.EOF Then
        lngBRID = rs("病人ID")
        If lngBRID <= 0 Then Exit Sub
        With fg病案_S
            lngCurRowTMP = .Row
            For i = lngCurRowTMP + 1 To .Rows - 1
                If Val(.Cell(flexcpText, i, 3)) = lngBRID Then
                    .Row = i
                    .ShowCell i, 2
                    blnFinded = True
                    Exit For
                End If
            Next
            If blnFinded = False Then '如果当前行下面没有匹配项，则从第一行开始重新查询。
                For i = 1 To lngCurRowTMP
                    If Val(.Cell(flexcpText, i, 3)) = lngBRID Then
                        .Row = i
                        .ShowCell i, 2
                        blnFinded = True
                        Exit For
                    End If
                Next
            End If
        End With
    End If
    rs.Close
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：树双击时加载数据
'==============================================================================
Private Sub tvw科室_DblClick()
    Dim i       As Integer
    On Error GoTo ErrH
    If tvw科室.SelectedItem Is Nothing Then Exit Sub
    With tvw科室.SelectedItem
        txt科室.Tag = Mid(.Key, 2)
        txt科室.Text = Mid(.Text, InStr(.Text, "】") + 1)
        tvw科室.Visible = False
        '出院科室快速选择
        If txt科室.Text = "所有科室" Then
            rsM.Filter = ""
        Else
            If mblnSetDept Then
                rsM.Filter = ""
            Else
                rsM.Filter = "出院科室='" & Mid(tvw科室.SelectedItem.Text, InStr(tvw科室.SelectedItem.Text, "】") + 1) & "'"
            End If
        End If
        Call Fill病案
    End With
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：树按键处理
'==============================================================================
Private Sub tvw科室_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrH
    If KeyAscii = vbKeyReturn Then
        '出院科室快速选择
        txt科室.Tag = Mid(tvw科室.SelectedItem.Key, 2)
        txt科室.Text = Mid(tvw科室.SelectedItem.Text, InStr(tvw科室.SelectedItem.Text, "】") + 1)
        If txt科室.Text = "所有科室" Then
            rsM.Filter = ""
        Else
            rsM.Filter = "出院科室='" & Mid(tvw科室.SelectedItem.Text, InStr(tvw科室.SelectedItem.Text, "】") + 1) & "'"
        End If
        Call Fill病案
    ElseIf KeyAscii = 27 Or KeyAscii = vbKeySpace Then
        tvw科室.Visible = False
        txt科室.SetFocus
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：树失去焦点时隐藏
'==============================================================================
Private Sub tvw科室_LostFocus()
    On Error GoTo ErrH
    tvw科室.Visible = False
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txt科室_Change()
    Dim lngStart        As Long
    Dim lngLength       As Long
On Error GoTo ErrH

    '控件录入控制字符等等
    lngLength = Len(txt科室.Text)
    lngStart = txt科室.SelStart
    txt科室.Text = ConvertString(txt科室.Text)
    If lngStart - (lngLength - Len(txt科室.Text)) >= 0 Then txt科室.SelStart = lngStart - (lngLength - Len(txt科室.Text))
 
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：文本框按键处理
'==============================================================================
Private Sub txt科室_KeyPress(KeyAscii As Integer)
    Dim zlInputNot      As String
    On Error GoTo ErrH
    '不允许输入字符
    zlInputNot = "'|"
    If Len(zlInputNot) > 0 Then
        If InStr(1, zlInputNot, Chr(KeyAscii)) > 0 Then
            KeyAscii = 0
        End If
    End If
    If txt科室.Locked Then Exit Sub
    
    If KeyAscii = vbKeyReturn Then
        Call DeptSelect
    ElseIf KeyAscii = 27 Then
        tvw科室.Visible = False
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：文本得到焦点选中内容
'==============================================================================
Private Sub txt内容_GotFocus()
    On Error GoTo ErrH
    Call zlControl.TxtSelAll(txt内容)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：定位内容回车刷新
'==============================================================================
Private Sub txt内容_KeyPress(KeyAscii As Integer)
    Dim lngRow      As Long
    Dim intCol      As Integer
    Dim bytMatch    As Byte
    Dim lngLoop     As Long
    
    On Error GoTo ErrH
    
    lngRow = 0
    If txt内容.Locked Then Exit Sub
    If mstrFindKey = "病人姓名" Then mstrFindKey = "姓名"
    If fg病案_S.ColIndex(mstrFindKey) = -1 Then Exit Sub
    If KeyAscii = vbKeyReturn Then
        '读取大于当前行的记录数据
        For lngLoop = fg病案_S.Row + 1 To fg病案_S.Rows - 1
            If InStr(UCase(fg病案_S.TextMatrix(lngLoop, fg病案_S.ColIndex(mstrFindKey))), UCase(txt内容.Text)) > 0 Then
                lngRow = lngLoop
                Exit For
            End If
        Next
        '读取小于当前行的记录数据
        If lngRow = 0 Then
            For lngLoop = 0 To fg病案_S.Row
                If InStr(UCase(fg病案_S.TextMatrix(lngLoop, fg病案_S.ColIndex(mstrFindKey))), UCase(txt内容.Text)) > 0 Then
                    lngRow = lngLoop
                    Exit For
                End If
            Next
        End If
        If fg病案_S.Rows > 1 And lngRow >= 1 Then fg病案_S.Row = lngRow
        Call LocationObj(txt内容)
    End If
    If mstrFindKey = "姓名" Then mstrFindKey = "病人姓名"
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：取得个性化设置
'==============================================================================
Private Function GetPersonSet() As Boolean
    
    On Error GoTo ErrH
    
    GetPersonSet = False
    If Val(zlDatabase.GetPara("使用个性化风格")) = 1 Then GetPersonSet = True

    Exit Function
ErrH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

'==============================================================================
'=功能：菜单按钮处理
'==============================================================================
Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnNewCancel        As Boolean
    
    On Error GoTo ErrH
    
    Select Case Control.ID
        Case conMenu_View_Refresh           '刷新数据
            Call mDataLoad
        Case conMenu_File_Preview           '预览
            subPrint 2
        Case conMenu_File_Print             '打印
            subPrint 1
        Case conMenu_File_Excel             '输出到&Excel
            subPrint 3
        Case conMenu_File_BatPrint          '全部打印
            Call RecordAllPrint
        Case conMenu_Edit_NewParent         '病案评分
            Call RecordRating
        Case conMenu_Edit_ModifyParent      '修改结果
            Call RecordEdit
        Case conMenu_Edit_Insert            '重新评分
            Call RecordReturn
        Case conMenu_Edit_DeleteParent      '删除结果
            Call RecordDel
        Case conMenu_Manage_ReportView      '查阅首页
            Call RecordLook
        Case conMenu_Manage_Audit           '通过审核
            Call RecordAudit
        Case conMenu_Edit_Leave_UndoPost    '取消审核
            Call RecordUnAudit
        Case conMenu_Edit_Select            '全部选中
            Call RecordSelect
        Case conMenu_Edit_DeSelect          '取消选中
            Call RecordUnSelect
        Case conMenu_Manage_UnAudit         '反向选择
            Call RecordSelectOther
        Case 10004                          '科室选择
            Call DeptSelect
        Case conMenu_View_Find              '过滤查询
            Call RecordFind
        Case conMenu_View_Forward           '上一条
            With fg病案_S
                If .Row > 1 Then
                    .Row = .Row - 1
                    .ShowCell .Row, .Col
                End If
            End With
        Case conMenu_View_Backward          '下一条
            With fg病案_S
                If .Row < .Rows - 1 Then
                    .Row = .Row + 1
                    .ShowCell .Row, .Col
                End If
            End With
        Case conMenu_View_Option
            mobjFindKey.Execute
        Case conMenu_View_LocationItem
            mstrFindKey = Control.Parameter
            mobjFindKey.Caption = mstrFindKey
            cbsMain.RecalcLayout
        Case conMenu_View_Location
            LocationObj txt内容
        Case Else
            If Control.ID > 400 And Control.ID < 500 Then
                Call ReportOpen(gcnOracle, Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1), Me)
            Else
                 '与业务无关的功能，公共的功能
                Call CommandBarExecutePublic(Control, Me, fg病案_S, "病案评分")
            End If
    End Select
    Exit Sub
ErrH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：菜单权限控制
'==============================================================================
Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)

    On Error GoTo ErrH
    
    Select Case Control.ID
        Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel, conMenu_File_BatPrint '预览,打印,输出到Excel,全部打印
            Control.Enabled = ((fg病案_S.Rows > 1) And IsPrivs(mstrPrivs, "评分"))
        Case conMenu_Edit_Select, conMenu_Edit_DeSelect, conMenu_Manage_UnAudit
            Control.Enabled = ((fg病案_S.Rows > 1))
        Case conMenu_Manage_ReportView
            Control.Enabled = ((fg病案_S.Rows > 1))
        Case conMenu_View_Refresh
            Control.Enabled = IsPrivs(mstrPrivs, "评分")
        Case conMenu_Edit_NewParent
            Control.Visible = (InStr(mstrPrivs, "评分") > 0)
            Control.Enabled = (Not mRecordRating) And (fg病案_S.Rows > 1)
        Case conMenu_Edit_ModifyParent      '修改[有评分权限且（能修改他人评分或自己的记录）且未审核]
            Control.Visible = (InStr(mstrPrivs, "评分") > 0)
            Control.Enabled = (InStr(mstrPrivs, "修改他人评分") > 0 Or mRecordMyAudit) And (fg病案_S.Rows > 1) And mRecordRating And Not mRecordAudit
        Case conMenu_Edit_DeleteParent
            Control.Visible = (InStr(mstrPrivs, "评分") > 0)
            Control.Enabled = (InStr(mstrPrivs, "修改他人评分") > 0 Or mRecordMyAudit) And (fg病案_S.Rows > 1) And mRecordRating And Not mRecordAudit
        Case conMenu_Edit_Insert
            Control.Visible = (InStr(mstrPrivs, "评分") > 0)
            Control.Enabled = (InStr(mstrPrivs, "修改他人评分") > 0 Or mRecordMyAudit) And (fg病案_S.Rows > 1) And mRecordRating And Not mRecordAudit And mRecordReturn
        Case conMenu_Manage_Audit   '审核
            Control.Visible = (InStr(mstrPrivs, "审核") > 0)
            Control.Enabled = (InStr(mstrPrivs, "修改他人评分") > 0 Or mRecordMyAudit) And (fg病案_S.Rows > 1) And mRecordRating And Not mRecordAudit
        Case conMenu_Edit_Leave_UndoPost
            Control.Visible = (InStr(mstrPrivs, "审核") > 0)
            Control.Enabled = (InStr(mstrPrivs, "修改他人评分") > 0 Or mRecordMyAudit) And (fg病案_S.Rows > 1) And mRecordRating And mRecordAudit
        Case 10004              '无所有科室则不能选择
            Control.Visible = IsPrivs(mstrPrivs, "所有科室")
        Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel, conMenu_File_BatPrint  '预览,打印,输出到Excel,全部打印
            Control.Enabled = ((fg病案_S.Rows > 1) And IsPrivs(mstrPrivs, "打印评分结果表"))
        Case conMenu_View_LocationItem, conMenu_View_LocationItem, conMenu_View_LocationItem
            If InStr(Control.Caption, mstrFindKey) > 0 Then
                Control.Checked = True
            Else
                Control.Checked = False
            End If
        Case conMenu_View_Forward
            Control.Enabled = (Control.Visible And fg病案_S.Row > 1)
        Case conMenu_View_Backward
                Control.Enabled = (Control.Visible And fg病案_S.Row + 1 < fg病案_S.Rows)
        Case Else
            Call CommandBarUpdatePublic(Control, Me)
    End Select
    Exit Sub
ErrH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub

Private Function Get所属部门(ByVal lng人员Id As Long, ByVal lngMode As Long) As String
    Dim strSQL As String
    Dim strTmp As String
    Dim rsTemp As ADODB.Recordset
    On Error GoTo ErrH
    ' lngMode =0 显示方式 lngMode =1 用于查询方式
    
    strSQL = "SELECT  distinct C.名称 AS 科室" & vbNewLine & _
                "      FROM 人员表 A,人员性质说明 B,部门表 C,部门人员 D" & vbNewLine & _
                "      WHERE A.ID=B.人员id AND C.ID=D.部门id AND D.人员id=A.ID And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null)" & vbNewLine & _
                "      AND A.id =[1]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng人员Id)
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        Do Until rsTemp.EOF
            If lngMode = 0 Then
                If Len(strTmp) = 0 Then
                    strTmp = NVL(rsTemp!科室)
                Else
                    strTmp = strTmp & "," & NVL(rsTemp!科室)
                End If
            Else
                If Len(strTmp) = 0 Then
                    strTmp = "'" & NVL(rsTemp!科室) & "'"
                Else
                    strTmp = strTmp & ",'" & NVL(rsTemp!科室) & "'"
                End If
            End If
            
            rsTemp.MoveNext
        Loop
        
        Get所属部门 = strTmp
    Else
        Get所属部门 = UserInfo.部门名称
    End If
    Exit Function
ErrH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    Exit Function
End Function
