VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmMediPriceCard 
   Caption         =   "药品调价单"
   ClientHeight    =   9075
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15150
   Icon            =   "frmMediPriceCard.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9075
   ScaleWidth      =   15150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picSplit 
      BorderStyle     =   0  'None
      Height          =   100
      Left            =   240
      MousePointer    =   7  'Size N S
      ScaleHeight     =   105
      ScaleWidth      =   2775
      TabIndex        =   32
      Top             =   4200
      Width           =   2775
   End
   Begin VB.PictureBox picOtherSelect 
      Height          =   3255
      Left            =   3360
      ScaleHeight     =   3195
      ScaleWidth      =   4875
      TabIndex        =   15
      Top             =   1080
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CommandButton cmdFilterOk 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   2640
         Picture         =   "frmMediPriceCard.frx":6852
         TabIndex        =   28
         Top             =   2760
         Width           =   1100
      End
      Begin VB.CommandButton cmdFilterCan 
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   3720
         Picture         =   "frmMediPriceCard.frx":699C
         TabIndex        =   27
         Top             =   2760
         Width           =   1100
      End
      Begin VB.Frame fra辅助选项 
         Caption         =   "辅助选项（成本价调价相关）"
         Height          =   2535
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   4695
         Begin VB.CheckBox chk加成率 
            Caption         =   "指定加成率"
            Height          =   180
            Left            =   120
            TabIndex        =   22
            Top             =   1125
            Width           =   1215
         End
         Begin VB.CheckBox chk供应商 
            Caption         =   "指定供应商"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox chk应付记录 
            Caption         =   "产生成本价调价带来的应付款修正记录"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   1920
            Width           =   3495
         End
         Begin VB.TextBox txt加成率 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   270
            Left            =   1440
            TabIndex        =   19
            Text            =   "15.0000"
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox txt供应商 
            Enabled         =   0   'False
            Height          =   270
            Left            =   1440
            TabIndex        =   18
            Top             =   360
            Width           =   2655
         End
         Begin VB.CommandButton cmd供应商 
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   270
            Left            =   4080
            TabIndex        =   17
            Top             =   350
            Width           =   375
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshProvider 
            Height          =   1695
            Left            =   120
            TabIndex        =   23
            Top             =   2280
            Visible         =   0   'False
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   2990
            _Version        =   393216
            FixedCols       =   0
            GridColor       =   32768
            FocusRect       =   0
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label lblComment加成率 
            Caption         =   "（指定加成率，则统一默认按该加成率计算成本价；不指定，则默认显示实际加成率）"
            ForeColor       =   &H00FF0000&
            Height          =   540
            Left            =   240
            TabIndex        =   26
            Top             =   1440
            Width           =   4260
         End
         Begin VB.Label lblComment供应商 
            AutoSize        =   -1  'True
            Caption         =   "（指定供应商，则只调整该供应商的库存药品成本价）"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   240
            TabIndex        =   25
            Top             =   720
            Width           =   4320
         End
         Begin VB.Label lblPercent 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   180
            Left            =   2415
            TabIndex        =   24
            Top             =   1125
            Width           =   90
         End
      End
   End
   Begin VB.PictureBox picInfo 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      ScaleHeight     =   495
      ScaleWidth      =   14175
      TabIndex        =   10
      Top             =   8160
      Width           =   14175
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   600
         TabIndex        =   35
         Top             =   120
         Width           =   1365
      End
      Begin VB.TextBox txtSummary 
         Height          =   300
         Left            =   5040
         MaxLength       =   100
         TabIndex        =   13
         Top             =   120
         Width           =   8835
      End
      Begin VB.TextBox txtValuer 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         Left            =   2790
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   120
         Width           =   1125
      End
      Begin VB.Label lblFind 
         BackColor       =   &H80000003&
         Caption         =   "查找"
         Height          =   180
         Left            =   120
         TabIndex        =   36
         Top             =   180
         Width           =   540
      End
      Begin VB.Label lblSummary 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "调价说明"
         Height          =   180
         Left            =   4200
         TabIndex        =   14
         Top             =   180
         Width           =   720
      End
      Begin VB.Label lblValuer 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "调价人"
         Height          =   180
         Left            =   2160
         TabIndex        =   12
         Top             =   180
         Width           =   540
      End
   End
   Begin VB.Frame fraCondition 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   16575
      Begin VB.PictureBox picAdjustTime 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   3840
         ScaleHeight     =   375
         ScaleWidth      =   5535
         TabIndex        =   39
         Top             =   120
         Width           =   5535
         Begin VB.OptionButton opt时间 
            BackColor       =   &H80000003&
            Caption         =   "指定日期"
            Height          =   255
            Index           =   1
            Left            =   1920
            TabIndex        =   41
            Top             =   15
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton opt时间 
            BackColor       =   &H80000003&
            Caption         =   "立即执行"
            Height          =   255
            Index           =   0
            Left            =   840
            TabIndex        =   40
            Top             =   15
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker dtpRunDate 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "yyyy-MM-dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
            Height          =   300
            Left            =   3000
            TabIndex        =   42
            Top             =   0
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
            Format          =   127729667
            CurrentDate     =   36846.5833333333
         End
         Begin VB.Label lbl执行时间 
            BackColor       =   &H80000003&
            Caption         =   "执行时间"
            Height          =   180
            Left            =   0
            TabIndex        =   43
            Top             =   45
            Width           =   855
         End
      End
      Begin VB.TextBox txtNO 
         Enabled         =   0   'False
         Height          =   300
         Left            =   14640
         TabIndex        =   37
         Top             =   120
         Width           =   1695
      End
      Begin VB.CheckBox chkAutoPay 
         BackColor       =   &H80000003&
         Caption         =   "自动计算应付款变动记录"
         Height          =   210
         Left            =   3360
         TabIndex        =   29
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CheckBox chkAotuCost 
         BackColor       =   &H80000003&
         Caption         =   "调售价时自动按加成率调整成本价"
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   3000
      End
      Begin VB.CommandButton cmdPriceMethod 
         Caption         =   "…"
         Height          =   300
         Left            =   3360
         TabIndex        =   5
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox cboPriceMethod 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   120
         Width           =   2415
      End
      Begin VB.CheckBox chk按批次 
         Caption         =   "成本价按库房批次调整"
         Height          =   210
         Left            =   10560
         TabIndex        =   3
         Top             =   -225
         Width           =   2175
      End
      Begin VB.CheckBox chk自动计算应付款变动 
         Caption         =   "自动计算应付款变动"
         Height          =   210
         Left            =   12840
         TabIndex        =   2
         Top             =   -225
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.ComboBox cbo售价计算方式 
         Height          =   300
         Left            =   10800
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label lblNO 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "调价流水号"
         Height          =   180
         Left            =   13560
         TabIndex        =   38
         Top             =   180
         Width           =   900
      End
      Begin VB.Label lbl调价方式 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "售价计算方式"
         Height          =   180
         Left            =   9480
         TabIndex        =   9
         Top             =   180
         Width           =   1080
      End
      Begin VB.Label lblMethod 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "调价方式"
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   180
         Width           =   720
      End
   End
   Begin XtremeSuiteControls.TabControl TabCtlDetails 
      Height          =   975
      Left            =   240
      TabIndex        =   6
      Top             =   4920
      Width           =   1815
      _Version        =   589884
      _ExtentX        =   3201
      _ExtentY        =   1720
      _StockProps     =   64
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfStore 
      Height          =   975
      Left            =   2880
      TabIndex        =   30
      Top             =   4680
      Width           =   3495
      _cx             =   6165
      _cy             =   1720
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   10526880
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
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
   Begin VSFlex8Ctl.VSFlexGrid vsfPay 
      Height          =   975
      Left            =   8040
      TabIndex        =   31
      Top             =   4680
      Width           =   3495
      _cx             =   6165
      _cy             =   1720
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   10526880
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
   Begin VSFlex8Ctl.VSFlexGrid vsfPrice 
      Height          =   2295
      Left            =   480
      TabIndex        =   33
      Top             =   2040
      Width           =   11055
      _cx             =   19500
      _cy             =   4048
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   10526880
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   34
      Top             =   8715
      Width           =   15150
      _ExtentX        =   26723
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMediPriceCard.frx":6AE6
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20955
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1402
            MinWidth        =   1411
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
            Object.ToolTipText     =   "当前数字键状态"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
            Object.ToolTipText     =   "当前大写键状态"
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
   Begin XtremeCommandBars.ImageManager imgList 
      Left            =   480
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmMediPriceCard.frx":737A
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmMediPriceCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'各种全局变量
Private Const mlngRowHeight As Long = 300 '表格中各行行高
Private mintUnit As Integer     '用来记录启用的是什么单位
Private mint调价 As Integer     '0-调售价;1-调成本价;2-调售价及成本价
Private mlng供应商ID As Long  '用来记录供应商id
Private mdbl加成率 As Double
Private mbln应付记录 As Boolean '记录是否产生应付记录

Private Enum typeAdjust
    AdjustPriceAndCost = 0
    AdjustPrice = 1
    AdjustCost = 2
End Enum

Private mintCostDigit As Integer        '成本价小数位数
Private mintPriceDigit As Integer       '售价小数位数
Private mintNumberDigit As Integer      '数量小数位数
Private mintMoneyDigit As Integer       '金额小数位数
Private mstrMoneyFormat As String
Private mintSalePriceDigit As Integer
'颜色方案
Private Const mconlngColor As Long = &HFFFFFF        '不能修改列颜色为白色
Private Const mconlngCanColColor As Long = &HE7CFBA    '能修改列颜色为淡蓝色
Private Const mlngBorderColor As Long = &H0&    '选中行边框颜色
Private Const mlngNoneBorderColor As Long = &HE0E0E0    ' 没选中行边框颜色

Private mbln时价药品按批次调价 As Boolean '时价药品按照批次调价
Private mbln成本价按库房批次调整 As Boolean '成本价按库房批次调整
Private mbln现价提示 As Boolean         '限价药品提示 true-提示 false-不提示
Private mdbl分段加成率 As Double    '用来记录分段加成率
Private mdbl成本价 As Double            '记录修改之前的成本价
Private mrs分段加成 As ADODB.Recordset  '记录分段加成率集合
Private mstrNo As String            '调价单No
Private mintModal As Integer        '本次是什么状态 0-新增 1-修改 2-查阅
Private mintMethod As Integer   '调价方式 0-调售价;1-调成本价;2-调售价及成本价
Private mstr调价汇总号 As String
Private mblnLoad As Boolean     '是否加载完成
Private mrsReturn As ADODB.Recordset '批量选择返回的数据集
Private mblnOK As Boolean
Private mrsFindName As ADODB.Recordset '查询的数据集
Private mBlnClick As Boolean
Private mblnUpdateAdd As Boolean    '修改情况下的新增卫材
Private mlngOldDrugID As Long '检查原始行是否有药品
Private mdblOldPrice As Double   '原售价
Private mblnBatchItem As Boolean   '记录是否点击了批量选择按钮
Private mstrPrivs As String     '操作员权限
Private Const MStrCaption As String = "药品调价单"

'功能按钮
Private Const mconMenu_Save = 100 '确定(&A)
Private Const mconMenu_Quit = 101 '取消(&Q)
Private Const mconMenu_PrintStore = 102 '打印库存变动表(&P)
Private Const mconMenu_ClearAll = 103 '清空列表(&C)
Private Const mconMenu_BatchSelect = 104 '批量选择项目
Private Const mconMenu_Find = 105 '查找
Private Const mconMenu_ModifyPrice = 106 '调价方式
Private Const mconMenu_CostPrice = 107 '仅调成本价
Private Const mconMenu_RetailPrice = 108  '仅调售价
Private Const mconMenu_Together = 109  '成本价售价一起调
Private Enum menuPriceCol
    药品id = 0
    原价id = 1
    药品 = 2
    规格 = 3
    药价属性 = 4
    是否变价
    产地
    单位
    包装系数
    加成率
    差价让利比
    是否有库存
    收入项目ID
    原成本价
    现成本价
    原零售价
    现零售价
    原采购限价
    现采购限价
    原指导售价
    现指导售价
    总列数
End Enum
Private Enum menuStoreCol
    药品id = 0
    药品 = 1
    规格 = 2
    库房 = 3
    库房id = 4
    供应商
    供应商id
    批号
    效期
    产地
    批次
    变价
    数量
    单位
    包装系数
    原成本价
    现成本价
    成本盈亏
    加成率
    原零售价
    现零售价
    售价盈亏
    总列数
End Enum

Private Enum menuPayCol
    药品id = 0
    药品 = 1
    发票号 = 2
    发票日期
    发票金额
    总列数
End Enum

Public Sub ShowME(ByVal frmParent As Form, ByVal intModal As Integer, ByVal str调价汇总号 As String, ByVal intMethod As Integer)
    mintModal = intModal
    mstr调价汇总号 = str调价汇总号
    mintMethod = intMethod

    Me.Show vbModal, frmParent
End Sub

Private Sub cboPriceMethod_Click()
    Dim intCol As Integer
    Dim intTemp As Integer

    With cboPriceMethod
        If .Text = "仅调售价" Then
            intTemp = 0
            lbl调价方式.Visible = False
            cbo售价计算方式.Visible = False
        ElseIf .Text = "仅调成本价" Then
            intTemp = 1
            lbl调价方式.Visible = False
            cbo售价计算方式.Visible = False
        Else
            intTemp = 2
            lbl调价方式.Visible = True
            cbo售价计算方式.Visible = True
        End If
    End With


    If mblnLoad = True And intTemp <> Val(lblMethod.Tag) Then
        If vsfPrice.TextMatrix(1, menuPriceCol.药品id) <> "" Then
            If MsgBox("调价方式改变将清空列表中数据，是否继续？", vbYesNo, gstrSysName) = vbNo Then
                cboPriceMethod.ListIndex = mint调价
                Exit Sub
            Else
                vsfPrice.rows = 2
                For intCol = 0 To vsfPrice.Cols - 1
                    vsfPrice.TextMatrix(1, intCol) = ""
                Next
                vsfStore.rows = 1
                vsfPay.rows = 1
            End If
        End If
    End If
    With cboPriceMethod
        If .Text = "仅调售价" Then
            mint调价 = 0
            lblMethod.Tag = 0
            opt时间(0).Value = False
            opt时间(1).Value = True
            opt时间(0).Enabled = True
            opt时间(1).Enabled = True
            dtpRunDate.Enabled = True
            chkAutoPay.Visible = False
            chkAutoPay.Value = 0
            chkAotuCost.Visible = False
            chkAotuCost.Value = False
        ElseIf .Text = "仅调成本价" Then
            mint调价 = 1
            lblMethod.Tag = 1
'            opt时间(0).Value = True
            opt时间(0).Enabled = True
            opt时间(1).Enabled = True
            dtpRunDate.Enabled = True
            If mbln应付记录 = True Then
                chkAutoPay.Visible = True
                chkAutoPay.Value = 1
            End If
            chkAotuCost.Visible = False
            chkAotuCost.Value = False
        ElseIf .Text = "售价成本价一起调价" Then
            mint调价 = 2
            lblMethod.Tag = 2
            opt时间(0).Value = False
            opt时间(1).Value = True
            opt时间(0).Enabled = True
            opt时间(1).Enabled = True
            dtpRunDate.Enabled = True
            If mbln应付记录 = True Then
                chkAutoPay.Visible = True
                chkAutoPay.Value = 1
            Else
                chkAutoPay.Visible = False
                chkAutoPay.Value = 0
            End If
            chkAotuCost.Visible = True
        End If
        If .Text = "仅调售价" Then
            cmdPriceMethod.Visible = False
            picOtherSelect.Visible = cmdPriceMethod.Visible
        Else
            cmdPriceMethod.Visible = True
        End If
    End With
    vsfStore.Cols = menuStoreCol.总列数
    vsfPay.Cols = menuPayCol.总列数
    vsfPrice.Cols = menuPriceCol.总列数
    Call setColEdit
    Call setColHiddenVsf
End Sub

Private Sub cboPriceMethod_DropDown()
    With cboPriceMethod
        If .Text = "仅调售价" Then
            mint调价 = 0
        ElseIf .Text = "仅调成本价" Then
            mint调价 = 1
        ElseIf .Text = "售价成本价一起调价" Then
            mint调价 = 2
        End If
    End With
End Sub

Private Sub cbo售价计算方式_Click()
    On Error GoTo errHandle
    Set mrs分段加成 = Nothing
    If cbo售价计算方式.Text = "售价按分段加成计算" Then
        gstrSQL = "select 序号, 最低价, 最高价, 加成率, 差价额, 说明, 类型 from 药品加成方案 order by 序号"
        Set mrs分段加成 = zlDatabase.OpenSQLRecord(gstrSQL, "药品加成方案")
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
        Case mconMenu_Save  '保存
            Call Save
        Case mconMenu_PrintStore    '打印库存变动表
            Call PrintStore
        Case mconMenu_ClearAll  '清空
            Call ClearAll
        Case mconMenu_Find '查找
            txtFind.SetFocus
            If Trim(txtFind.Text) = "" Then Exit Sub
            Call FindGridRow(UCase(Trim(txtFind.Text)))
        Case mconMenu_Quit  '取消
            Call Quit
        Case mconMenu_BatchSelect  '批量选择项目
            Call BatchSelect
    End Select
End Sub

Private Sub chkAotuCost_Click()
    If chkAotuCost.Value = 1 Then
        cbo售价计算方式.Visible = False
        cbo售价计算方式.ListIndex = 0
        lbl调价方式.Visible = False
    Else
        cbo售价计算方式.Visible = True
        lbl调价方式.Visible = True
    End If
End Sub


Private Sub Chk供应商_Click()
    If chk供应商.Value = 1 Then
        cmd供应商.Enabled = True
        txt供应商.Enabled = True
        chk应付记录.Enabled = True
    Else
        cmd供应商.Enabled = False
        txt供应商.Enabled = False
        chk应付记录.Enabled = False
        chk应付记录.Value = 0
    End If
End Sub

Private Sub chk加成率_Click()
    If chk加成率.Value = 1 Then
        txt加成率.Enabled = True
    Else
        txt加成率.Enabled = False
    End If
End Sub

Private Sub Quit()
    Call ReleaseSelectorRS '卸载数据集
    Unload Me
End Sub

Private Sub ClearAll()
    Dim intCol As Integer

    If MsgBox("你确定要清空所有数据？", vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        vsfPrice.rows = 2
        For intCol = 0 To vsfPrice.Cols - 1
            vsfPrice.TextMatrix(1, intCol) = ""
        Next
        vsfStore.rows = 1
        vsfPay.rows = 1
    End If
End Sub

Private Sub cmdFilterCan_Click()
    picOtherSelect.Visible = False
End Sub

Private Sub cmdFilterOk_Click()
    Dim i As Integer

    If chk供应商.Value = 1 Then
        If Val(Split(txt供应商.Tag, "|")(0)) = 0 Then
            MsgBox "请选择供应商。", vbInformation, gstrSysName
            txt供应商.SetFocus
            Exit Sub
        End If
    End If
    With vsfPrice
        If Val(.TextMatrix(1, menuPriceCol.药品id)) <> 0 Then
            If MsgBox("将清空表格中的数据，是否继续？", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            Else
                vsfPrice.rows = 2
                For i = 0 To vsfPrice.Cols - 1
                    .TextMatrix(1, i) = ""
                Next
                vsfStore.rows = 1
                vsfPay.rows = 1
            End If
        End If
    End With

    mlng供应商ID = IIf(chk供应商.Value = 1, Val(Split(txt供应商.Tag, "|")(0)), 0)
    mdbl加成率 = IIf(chk加成率.Value = 1, Val(Trim(txt加成率.Text)), 0)
    mbln应付记录 = (chk应付记录.Enabled And chk应付记录.Value = 1)
    picOtherSelect.Visible = False
    If mbln应付记录 = True Then
        TabCtlDetails.Item(1).Visible = True
    Else
        TabCtlDetails.Item(1).Visible = False
    End If

    With cboPriceMethod
        If .Text = "仅调售价" Then
            mint调价 = 0
            lblMethod.Tag = 0
            opt时间(0).Value = False
            opt时间(1).Value = True
            opt时间(0).Enabled = True
            opt时间(1).Enabled = True
            dtpRunDate.Enabled = True
            chkAutoPay.Visible = False
            chkAutoPay.Value = 0
            chkAotuCost.Visible = False
            chkAotuCost.Value = False
        ElseIf .Text = "仅调成本价" Then
            mint调价 = 1
            lblMethod.Tag = 1
            opt时间(0).Value = True
            opt时间(0).Enabled = False
            opt时间(1).Enabled = False
            dtpRunDate.Enabled = False
            If mbln应付记录 = True Then
                chkAutoPay.Visible = True
                chkAutoPay.Value = 1
            Else
                chkAutoPay.Visible = False
                chkAutoPay.Value = 0
            End If
            chkAotuCost.Visible = False
            chkAotuCost.Value = False
        ElseIf .Text = "售价成本价一起调价" Then
            mint调价 = 2
            lblMethod.Tag = 2
            opt时间(0).Value = False
            opt时间(1).Value = True
            opt时间(0).Enabled = True
            opt时间(1).Enabled = True
            dtpRunDate.Enabled = True
            If mbln应付记录 = True Then
                chkAutoPay.Visible = True
                chkAutoPay.Value = 1
            Else
                chkAutoPay.Visible = False
                chkAutoPay.Value = 0
            End If
            chkAotuCost.Visible = True
        End If
    End With

End Sub

Private Sub CmdHelp_Click()

End Sub

Private Sub BatchSelect()
    Dim intRow As Integer

    frmBatchSelect.ShowME Me, mrsReturn, mblnOK

    On Error GoTo errHandle
    If mblnOK = False Then Exit Sub
    If mrsReturn.RecordCount = 0 Then Exit Sub

    With vsfPrice
        If .TextMatrix(.rows - 1, menuPriceCol.药品id) = "" Then
            intRow = .rows - 1
        Else
            .rows = .rows + 1
            intRow = .rows - 1
        End If
    End With
    mblnBatchItem = True

    Call GetDrugPirce(mrsReturn, intRow)
    mblnBatchItem = False
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub deleteNotExecutePirce()
    '清除未执行价格
    Dim intRow As Integer
    Dim int删除类型 As Integer
    
    'Private mint调价 As Integer     '0-调售价;1-调成本价;2-调售价及成本价
    '删除方式_In   In Number := 0 --0-所有;1-售价;2-成本价
    On Error GoTo errHandle
    
    If mint调价 = 0 Then
        int删除类型 = 1
    ElseIf mint调价 = 1 Then
        int删除类型 = 2
    Else
        int删除类型 = 0
    End If
    
    With vsfPrice
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, menuPriceCol.药品id) <> "" Then
                gstrSQL = "Zl_药品未执行价格_Delete(" & Val(.TextMatrix(intRow, menuPriceCol.药品id)) & "," & int删除类型 & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, MStrCaption)
            End If
        Next
    End With

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Save()
    Dim intRow As Integer
    Dim intCol As Integer
    Dim dtToday As Date
    Dim lngAdjId As Long
    Dim LngCurID As Long
    Dim strID As String
    Dim intCount As Integer
    Dim dbl包装 As Double
    Dim strTmp As String
    Dim lngCurrBatch As Long
    Dim str批次价格 As String
    Dim blnPrint As Boolean '是否打印调价通知单
    Dim blnOne As Boolean   '检查是否是第一行
    Dim n As Integer
    Dim intProc As Integer
    Dim blnIgnore As Boolean
    Dim blnPrice As Boolean '记录是否售价调价了
    Dim blnCost As Boolean  '记录是否成本价调价了
    Dim intUpdateModel As Integer '调价模式 0-售价调价 1-成本价调价 2-成本价售价一起调价
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim ArrayID
    Dim Array批次价格
    Dim strUpdate As String

    Dim lng库房ID As Long
    Dim lng供应商ID As Long
    Dim lng药品id As Long
    Dim lng批次  As Long
    Dim str批号 As String
    Dim str效期 As String
    Dim str产地 As String
    Dim dblOldCost As Double
    Dim dblNewCost As Double
    Dim Str发票号 As String
    Dim str发票日期 As String
    Dim dbl发票金额 As Double
    Dim strInfo As String
    Dim strMsg As String '记录提示信息
    Dim intCount2 As Integer '用来计数
    Dim lngDouID As Long
    
    Dim str执行时间 As String
    Dim str终止时间 As String
    Dim strDrugs As String
    
    If vsfPrice.rows > 1 Then   '只有有数据的情况下才能保存
        If Val(vsfPrice.TextMatrix(1, menuPriceCol.药品id)) = 0 Then Exit Sub
    End If
    If CheckPrice = False Then Exit Sub
    
    On Error GoTo ErrHand
    
    dtToday = Sys.Currentdate()
    If opt时间(0).Value = True Then
        str执行时间 = Format(dtToday, "YYYY-MM-DD HH:mm:ss")
        str终止时间 = Format(DateAdd("s", -1, dtToday), "YYYY-MM-DD HH:mm:ss")
    Else
        str执行时间 = Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss")
        str终止时间 = Format(DateAdd("s", -1, Me.dtpRunDate.Value), "YYYY-MM-DD HH:mm:ss")
    End If
                    
    gstrSQL = "select 收费价目_ID.nextval from dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取收费价目序号")
    lngAdjId = rsTemp.Fields(0).Value

    gcnOracle.BeginTrans
    If mintModal = 1 Then '修改 在修改模式下先删除原来的调价信息，然后插入新的调价信息
        Call deleteNotExecutePirce
    End If

    '检查是否存在未执行的价格
    If checkNotExecutePrice(, strInfo) = True Then
        MsgBox strInfo, vbInformation, gstrSysName
        Exit Sub
    End If
    
    '获取调价NO
    mstrNo = Sys.GetNextNo(9)
    '获取调价汇总NO
    gstrSQL = "select nextno(135) as 流水号 from dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "调价流水号")
    If rsTemp.RecordCount = 0 Then
        MsgBox "调价流水号未能初始化成功，请与管理员联系！", vbInformation, gstrSysName
        Exit Sub
    End If
    txtNO.Text = rsTemp!流水号

    With Me.vsfPrice
        '售价调价
        strID = ""
        For intCount = 1 To IIf(Trim(.TextMatrix(.rows - 1, 0)) = "", .rows - 2, .rows - 1)
            If mint调价 <> 1 Then
                LngCurID = Sys.NextId("收费价目")
                
                strID = strID & IIf(strID = "", "", ",") & LngCurID
                
                If InStr(1, "," & strDrugs & ",", "," & Val(.TextMatrix(intCount, menuPriceCol.药品id)) & ",") = 0 Then
                    strDrugs = IIf(strDrugs = "", "", strDrugs & ",") & Val(.TextMatrix(intCount, menuPriceCol.药品id))
                End If
                
                dbl包装 = Val(.TextMatrix(intCount, menuPriceCol.包装系数))

                If .TextMatrix(intCount, menuPriceCol.是否变价) = "1" And mbln时价药品按批次调价 And mint调价 <> 1 Then
                    strTmp = ""
                    lngCurrBatch = -1
                    For n = 1 To vsfStore.rows - 1
                        If Val(.TextMatrix(intCount, menuPriceCol.药品id)) = Val(vsfStore.TextMatrix(n, menuStoreCol.药品id)) Then
                            If InStr(1, "|" & strTmp, "|" & vsfStore.TextMatrix(n, menuStoreCol.批次) & ",") = 0 Then
                                lngCurrBatch = vsfStore.TextMatrix(n, menuStoreCol.批次)
                                strTmp = strTmp & IIf(strTmp = "", "", "|") & vsfStore.TextMatrix(n, menuStoreCol.批次) & "," & vsfStore.TextMatrix(n, menuStoreCol.现零售价) / dbl包装
                            End If
                        End If
                    Next
                    str批次价格 = str批次价格 & strTmp
                End If
                str批次价格 = str批次价格 & ";"
                             
                If CLng(.TextMatrix(intCount, menuPriceCol.原价id)) <> 0 Then
                    '设置上一次的价格记录终止执行
                    gstrSQL = "zl_收费价目_stop(" & .TextMatrix(intCount, menuPriceCol.药品id) & ","
                    If opt时间(0).Value = True Then
                        gstrSQL = gstrSQL & "to_date('" & Format(DateAdd("s", -1, dtToday), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    Else
                        gstrSQL = gstrSQL & "to_date('" & Format(DateAdd("s", -1, Me.dtpRunDate.Value), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    End If
                    gstrSQL = gstrSQL & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, MStrCaption)

                    '产生价格记录
                    gstrSQL = "zl_收费价目_Insert(" & LngCurID & "," & IIf(.TextMatrix(intCount, menuPriceCol.原价id) = "", "NUll", Val(.TextMatrix(intCount, menuPriceCol.原价id))) & _
                              "," & .TextMatrix(intCount, menuPriceCol.药品id) & "," & Val(.TextMatrix(intCount, menuPriceCol.收入项目ID)) & "," & _
                              Round(Val(.TextMatrix(intCount, menuPriceCol.原零售价)) / dbl包装, gtype_UserDrugDigits.Digit_零售价) & "," & _
                              Round(Val(.TextMatrix(intCount, menuPriceCol.现零售价)) / dbl包装, gtype_UserDrugDigits.Digit_零售价) & _
                              ",NULL,NULL,'" & Me.txtSummary.Text & "'," & lngAdjId & ",'" & Trim(Me.txtValuer.Text) & "',"
                    If opt时间(0).Value = True Then
                        gstrSQL = gstrSQL & "to_date('" & Format(dtToday, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    Else
                        gstrSQL = gstrSQL & "to_date('" & Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    End If
                    gstrSQL = gstrSQL & ",0,'" & mstrNo & "'," & intCount & ",Null," & txtNO & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, MStrCaption)
                    blnPrice = True
                    blnPrint = True
                End If
                
                If .TextMatrix(intCount, menuPriceCol.是否变价) = "1" And mint调价 <> 1 Then
                    If .TextMatrix(intCount, menuPriceCol.是否有库存) = "0" Then
                        If Val(.TextMatrix(intCount, menuPriceCol.原零售价)) <> Val(.TextMatrix(intCount, menuPriceCol.现零售价)) Then
                            '时价药品无库存调价
                            dbl包装 = Val(.TextMatrix(intCount, menuPriceCol.包装系数))
                            lng药品id = Val(.TextMatrix(intCount, menuPriceCol.药品id))
                            dblOldCost = Val(.TextMatrix(intCount, menuPriceCol.原零售价)) / dbl包装
                            dblNewCost = Val(.TextMatrix(intCount, menuPriceCol.现零售价)) / dbl包装
                            
                            gstrSQL = "Zl_药品价格记录_Stop("
                            '价格类型_In
                            gstrSQL = gstrSQL & 1
                            '库房id_In
                            gstrSQL = gstrSQL & ",Null"
                            '药品id_In
                            gstrSQL = gstrSQL & "," & lng药品id
                            '批次_In
                            gstrSQL = gstrSQL & ",0"
                            '终止日期_In
                            gstrSQL = gstrSQL & "," & "to_date('" & str终止时间 & "','YYYY-MM-DD HH24:MI:SS')"
                            gstrSQL = gstrSQL & ")"
                            Call zlDatabase.ExecuteProcedure(gstrSQL, MStrCaption)
                        
                            gstrSQL = "Zl_药品价格记录_Insert("
                            '调价类型_In
                            gstrSQL = gstrSQL & 1
                            '价格类型_In
                            gstrSQL = gstrSQL & ",1"
                            '库房id_In
                            gstrSQL = gstrSQL & ",Null"
                            '药品id_In
                            gstrSQL = gstrSQL & "," & lng药品id
                            '批次_In
                            gstrSQL = gstrSQL & ",0"
                            
                            '原价_In
                            gstrSQL = gstrSQL & "," & dblOldCost
                            '现价_In
                            gstrSQL = gstrSQL & "," & dblNewCost
                            '执行日期_In
                            gstrSQL = gstrSQL & "," & "to_date('" & str执行时间 & "','YYYY-MM-DD HH24:MI:SS')"
                            '调价说明_In
                            gstrSQL = gstrSQL & ",'" & Me.txtSummary.Text & "'"
                            '调价人_In
                            gstrSQL = gstrSQL & ",'" & Trim(Me.txtValuer.Text) & "'"
                            
                            '调价汇总号_In
                            gstrSQL = gstrSQL & ",'" & txtNO.Text & "'"
                            
                            gstrSQL = gstrSQL & ")"
                            Call zlDatabase.ExecuteProcedure(gstrSQL, MStrCaption)
                            
                            blnPrice = True
                         End If
                    Else
                        '时价药品有库存调价
                        For n = 1 To vsfStore.rows - 1
                            If Val(.TextMatrix(intCount, menuPriceCol.药品id)) = Val(vsfStore.TextMatrix(n, menuStoreCol.药品id)) Then
                                lng库房ID = Val(vsfStore.TextMatrix(n, menuStoreCol.库房id))
                                lng药品id = Val(vsfStore.TextMatrix(n, menuStoreCol.药品id))
                                lng批次 = Val(vsfStore.TextMatrix(n, menuStoreCol.批次))
                                lng供应商ID = Val(vsfStore.TextMatrix(n, menuStoreCol.供应商id))
                                str批号 = vsfStore.TextMatrix(n, menuStoreCol.批号)
                                str效期 = IIf(Trim(vsfStore.TextMatrix(n, menuStoreCol.效期)) = "", "", vsfStore.TextMatrix(n, menuStoreCol.效期))
                                str产地 = vsfStore.TextMatrix(n, menuStoreCol.产地)
                                dblOldCost = Val(vsfStore.TextMatrix(n, menuStoreCol.原零售价)) / Val(vsfStore.TextMatrix(n, menuStoreCol.包装系数))
                                dblNewCost = Val(vsfStore.TextMatrix(n, menuStoreCol.现零售价)) / Val(vsfStore.TextMatrix(n, menuStoreCol.包装系数))
                                
                                gstrSQL = "Zl_药品价格记录_Stop("
                                '价格类型_In
                                gstrSQL = gstrSQL & 1
                                '库房id_In
                                gstrSQL = gstrSQL & "," & lng库房ID
                                '药品id_In
                                gstrSQL = gstrSQL & "," & lng药品id
                                '批次_In
                                gstrSQL = gstrSQL & "," & lng批次
                                '终止日期_In
                                gstrSQL = gstrSQL & "," & "to_date('" & str终止时间 & "','YYYY-MM-DD HH24:MI:SS')"
                                gstrSQL = gstrSQL & ")"
                                Call zlDatabase.ExecuteProcedure(gstrSQL, MStrCaption)
                                
                                gstrSQL = "Zl_药品价格记录_Insert("
                                '调价类型_In
                                gstrSQL = gstrSQL & 1
                                '价格类型_In
                                gstrSQL = gstrSQL & ",1"
                                '库房id_In
                                gstrSQL = gstrSQL & "," & lng库房ID
                                '药品id_In
                                gstrSQL = gstrSQL & "," & lng药品id
                                '批次_In
                                gstrSQL = gstrSQL & "," & lng批次
                                
                                '原价_In
                                gstrSQL = gstrSQL & "," & dblOldCost
                                '现价_In
                                gstrSQL = gstrSQL & "," & dblNewCost
                                '执行日期_In
                                gstrSQL = gstrSQL & "," & "to_date('" & str执行时间 & "','YYYY-MM-DD HH24:MI:SS')"
                                '调价说明_In
                                gstrSQL = gstrSQL & ",'" & Me.txtSummary.Text & "'"
                                '调价人_In
                                gstrSQL = gstrSQL & ",'" & Trim(Me.txtValuer.Text) & "'"
                                
                                '调价汇总号_In
                                gstrSQL = gstrSQL & ",'" & txtNO.Text & "'"
                                '供药单位id_In
                                gstrSQL = gstrSQL & "," & IIf(lng供应商ID = 0, "Null", lng供应商ID)
                                '批号_In
                                gstrSQL = gstrSQL & ",'" & str批号 & "'"
                                '效期_In
                                gstrSQL = gstrSQL & "," & IIf(str效期 = "", "Null", "to_date('" & Format(str效期, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                                '产地_In
                                gstrSQL = gstrSQL & ",'" & str产地 & "'"
                                
                                gstrSQL = gstrSQL & ")"
                                Call zlDatabase.ExecuteProcedure(gstrSQL, MStrCaption)
                                
                                blnPrice = True
                                blnPrint = True
                            End If
                        Next
                    End If
                End If
            End If
        Next
    End With

    '成本价调价处理
    If mint调价 = 1 Or mint调价 = 2 Then
        If vsfStore.rows > 1 Then
            If vsfStore.TextMatrix(1, menuStoreCol.药品id) <> "" Then
'                lngDouID = 0
'                For n = 1 To vsfStore.rows - 1
'                    If vsfStore.TextMatrix(n, menuStoreCol.药品id) = "" Then Exit For
'
'                    '检查未审核单据
'                    If CheckUnVerify(Val(vsfStore.TextMatrix(n, menuStoreCol.药品id))) = True And Val(vsfStore.TextMatrix(n, menuStoreCol.药品id)) <> lngDouID Then
'                        lngDouID = Val(vsfStore.TextMatrix(n, menuStoreCol.药品id))
'                        strMsg = vsfStore.TextMatrix(n, menuStoreCol.药品) & ","
'                        intCount2 = intCount2 + 1
'                        If intCount2 > 3 Then Exit For '只判断3个
'                    End If
'                Next
'
'                If strMsg <> "" Then
'                    If MsgBox(strMsg & "存在未审核单据，调整成本价可能会造成差价误差。" & _
'                        vbCrLf & Space(4) & "建议先处理未审核单据。是否还继续调价？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
'                        gcnOracle.RollbackTrans
'                        Exit Sub
'                    End If
'                End If

                For n = 1 To vsfStore.rows - 1
                    For i = 1 To vsfPay.rows - 1
                        If vsfPay.TextMatrix(i, 0) = "" Then Exit For
                        If Val(vsfStore.TextMatrix(n, menuStoreCol.药品id)) = Val(vsfPay.TextMatrix(i, menuPayCol.药品id)) Then
                            lng库房ID = Val(vsfStore.TextMatrix(n, menuStoreCol.库房id))
                            lng供应商ID = Val(vsfStore.TextMatrix(n, menuStoreCol.供应商id))
                            lng药品id = Val(vsfStore.TextMatrix(n, menuStoreCol.药品id))
                            lng批次 = Val(vsfStore.TextMatrix(n, menuStoreCol.批次))
                            str批号 = vsfStore.TextMatrix(n, menuStoreCol.批号)
                            str效期 = IIf(Trim(vsfStore.TextMatrix(n, menuStoreCol.效期)) = "", "", vsfStore.TextMatrix(n, menuStoreCol.效期))
                            str产地 = vsfStore.TextMatrix(n, menuStoreCol.产地)
                            dblOldCost = zlStr.FormatEx(Val(vsfStore.TextMatrix(n, menuStoreCol.原成本价)) / Val(vsfStore.TextMatrix(n, menuStoreCol.包装系数)), gtype_UserDrugDigits.Digit_成本价, , True)
                            dblNewCost = zlStr.FormatEx(Val(vsfStore.TextMatrix(n, menuStoreCol.现成本价)) / Val(vsfStore.TextMatrix(n, menuStoreCol.包装系数)), gtype_UserDrugDigits.Digit_成本价, , True)
                            Str发票号 = vsfPay.TextMatrix(i, menuPayCol.发票号)
                            str发票日期 = Format(vsfPay.TextMatrix(i, menuPayCol.发票日期), "yyyy-mm-dd")
                            dbl发票金额 = Val(vsfPay.TextMatrix(i, menuPayCol.发票金额))
                            
'                            gstrSQL = "Zl_成本价调价信息_Insert(" & IIf(lng供应商ID = 0, "Null", lng供应商ID) & "," & lng库房ID & "," & lng药品ID & "," & lng批次 & ",'" & str批号 & "'" & _
'                                    "," & IIf(str效期 = "", "Null", "to_date('" & Format(str效期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ",'" & str产地 & "',Null," & dblOldCost & ", " & dblNewCost & "," & _
'                                    IIf(Str发票号 <> "", "'" & Str发票号 & "'", "NULL") & "," & IIf(str发票日期 = "", "Null", "to_date('" & Format(str发票日期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ", " & dbl发票金额 & "," & IIf(mbln应付记录 = True, 1, 0) & "," & txtNo.Text & ")"
'                            Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
                            
                            gstrSQL = "Zl_药品价格记录_Stop("
                            '价格类型_In
                            gstrSQL = gstrSQL & 2
                            '库房id_In
                            gstrSQL = gstrSQL & "," & lng库房ID
                            '药品id_In
                            gstrSQL = gstrSQL & "," & lng药品id
                            '批次_In
                            gstrSQL = gstrSQL & "," & lng批次
                            '终止日期_In
                            gstrSQL = gstrSQL & "," & "to_date('" & str终止时间 & "','YYYY-MM-DD HH24:MI:SS')"
                            gstrSQL = gstrSQL & ")"
                            Call zlDatabase.ExecuteProcedure(gstrSQL, MStrCaption)
                            
'                            调价类型_In   In 药品价格记录.调价类型%Type,
'                            价格类型_In   In 药品价格记录.价格类型%Type,
'                            库房id_In     In 药品价格记录.库房id%Type,
'                            药品id_In     In 药品价格记录.药品id%Type,
'                            批次_In       In 药品价格记录.批次%Type := Null,
'
'                            原价_In       In 药品价格记录.原价%Type := Null,
'                            现价_In       In 药品价格记录.现价%Type := Null,
'                            执行日期_In   In 药品价格记录.执行日期%Type := Null,
'                            调价说明_In   In 药品价格记录.调价说明%Type := Null,
'                            调价人_In     In 药品价格记录.调价人%Type := Null,
'
'                            调价汇总号_In In 药品价格记录.调价汇总号%Type := Null,
'                            供药单位id_In In 药品价格记录.供药单位id%Type := Null,
'                            批号_In       In 药品价格记录.批号%Type := Null,
'                            效期_In       In 药品价格记录.效期%Type := Null,
'                            产地_In       In 药品价格记录.产地%Type := Null
'
'                            灭菌效期_In   In 药品价格记录.灭菌效期%Type := Null,
'                            发票号_In     In 药品价格记录.发票号%Type := Null,
'                            发票日期_In   In 药品价格记录.发票日期%Type := Null,
'                            发票金额_In   In 药品价格记录.发票金额%Type := Null,
'                            应付款变动_In In 药品价格记录.应付款变动%Type := 0
  
                            
                            gstrSQL = "Zl_药品价格记录_Insert("
                            '调价类型_In
                            gstrSQL = gstrSQL & 1
                            '价格类型_In
                            gstrSQL = gstrSQL & ",2"
                            '库房id_In
                            gstrSQL = gstrSQL & "," & lng库房ID
                            '药品id_In
                            gstrSQL = gstrSQL & "," & lng药品id
                            '批次_In
                            gstrSQL = gstrSQL & "," & lng批次
                            
                            '原价_In
                            gstrSQL = gstrSQL & "," & dblOldCost
                            '现价_In
                            gstrSQL = gstrSQL & "," & dblNewCost
                            '执行日期_In
                            gstrSQL = gstrSQL & "," & "to_date('" & str执行时间 & "','YYYY-MM-DD HH24:MI:SS')"
                            '调价说明_In
                            gstrSQL = gstrSQL & ",'" & Me.txtSummary.Text & "'"
                            '调价人_In
                            gstrSQL = gstrSQL & ",'" & Trim(Me.txtValuer.Text) & "'"
                            
                            '调价汇总号_In
                            gstrSQL = gstrSQL & ",'" & txtNO.Text & "'"
                            '供药单位id_In
                            gstrSQL = gstrSQL & "," & IIf(lng供应商ID = 0, "Null", lng供应商ID)
                            '批号_In
                            gstrSQL = gstrSQL & ",'" & str批号 & "'"
                            '效期_In
                            gstrSQL = gstrSQL & "," & IIf(str效期 = "", "Null", "to_date('" & Format(str效期, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                            '产地_In
                            gstrSQL = gstrSQL & ",'" & str产地 & "'"
                            
                            '灭菌效期_In
                            gstrSQL = gstrSQL & ",Null"
                            '发票号_In
                            gstrSQL = gstrSQL & ",'" & Str发票号 & "'"
                            '发票日期_In
                            gstrSQL = gstrSQL & "," & IIf(str发票日期 = "", "Null", "to_date('" & Format(str发票日期, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                            '发票金额_In
                            gstrSQL = gstrSQL & "," & dbl发票金额
                            '应付款变动_In
                            gstrSQL = gstrSQL & "," & IIf(mbln应付记录 = True, 1, 0)
                            
                            gstrSQL = gstrSQL & ")"
                            Call zlDatabase.ExecuteProcedure(gstrSQL, MStrCaption)
                            
                            blnCost = True
                            blnPrint = True
                        End If
                    Next
                Next
            End If
        End If
    End If

    '无库存时调整成本价
    If mint调价 = 1 Or mint调价 = 2 Then
        With Me.vsfPrice
            For intCount = 1 To IIf(Trim(.TextMatrix(.rows - 1, 0)) = "", .rows - 2, .rows - 1)
                If .TextMatrix(intCount, menuPriceCol.是否有库存) = "0" And Val(.TextMatrix(intCount, menuPriceCol.原成本价)) <> Val(.TextMatrix(intCount, menuPriceCol.现成本价)) Then
                    dbl包装 = Val(.TextMatrix(intCount, menuPriceCol.包装系数))

                    lng药品id = Val(.TextMatrix(intCount, menuPriceCol.药品id))
                    dblOldCost = Val(Round(Val(.TextMatrix(intCount, menuPriceCol.原成本价)) / dbl包装, gtype_UserDrugDigits.Digit_成本价))
                    dblNewCost = Val(Round(Val(.TextMatrix(intCount, menuPriceCol.现成本价)) / dbl包装, gtype_UserDrugDigits.Digit_成本价))

'                    gstrSQL = "Zl_成本价调价信息_Insert(Null,Null," & lng药品ID & ",0,Null,Null,Null,Null," & dblOldCost & ", " & dblNewCost & ",NULL,Null,0,0, " & txtNO.Text & ")"
'                    Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
                    
                    gstrSQL = "Zl_药品价格记录_Stop("
                    '价格类型_In
                    gstrSQL = gstrSQL & 2
                    '库房id_In
                    gstrSQL = gstrSQL & ",Null"
                    '药品id_In
                    gstrSQL = gstrSQL & "," & lng药品id
                    '批次_In
                    gstrSQL = gstrSQL & ",0"
                    '终止日期_In
                    gstrSQL = gstrSQL & "," & "to_date('" & str终止时间 & "','YYYY-MM-DD HH24:MI:SS')"
                    gstrSQL = gstrSQL & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, MStrCaption)
                    
                    gstrSQL = "Zl_药品价格记录_Insert("
                    '调价类型_In
                    gstrSQL = gstrSQL & 1
                    '价格类型_In
                    gstrSQL = gstrSQL & ",2"
                    '库房id_In
                    gstrSQL = gstrSQL & ",Null"
                    '药品id_In
                    gstrSQL = gstrSQL & "," & lng药品id
                    '批次_In
                    gstrSQL = gstrSQL & ",0"
                    
                    '原价_In
                    gstrSQL = gstrSQL & "," & dblOldCost
                    '现价_In
                    gstrSQL = gstrSQL & "," & dblNewCost
                    '执行日期_In
                    gstrSQL = gstrSQL & "," & "to_date('" & str执行时间 & "','YYYY-MM-DD HH24:MI:SS')"
                    '调价说明_In
                    gstrSQL = gstrSQL & ",'" & Me.txtSummary.Text & "'"
                    '调价人_In
                    gstrSQL = gstrSQL & ",'" & Trim(Me.txtValuer.Text) & "'"
                    
                    '调价汇总号_In
                    gstrSQL = gstrSQL & ",'" & txtNO.Text & "'"
                    
                    gstrSQL = gstrSQL & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, MStrCaption)
                    
                    blnCost = True
                End If
            Next
        End With
    End If

    '立即执行
    If mint调价 = 1 Then
        '单独成本价调价时
        If opt时间(0).Value = True Then
            With Me.vsfPrice
                For intCount = 1 To IIf(Trim(.TextMatrix(.rows - 1, 0)) = "", .rows - 2, .rows - 1)
                    gstrSQL = "zl_药品收发记录_Adjust(" & Val(.TextMatrix(intCount, menuPriceCol.药品id)) & "," & typeAdjust.AdjustCost & " )"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, MStrCaption)
                Next
            End With
        End If
    Else
        '调售价
        If opt时间(0).Value = True Then
            ArrayID = Split(strDrugs, ",")
            For intCount = 0 To UBound(ArrayID)
                gstrSQL = "zl_药品收发记录_Adjust(" & ArrayID(intCount) & "," & IIf(mint调价 = 0, typeAdjust.AdjustPrice, typeAdjust.AdjustPriceAndCost) & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, MStrCaption)
            Next
        End If
    End If

    '调整指导价格
    With Me.vsfPrice
        For intCount = 1 To IIf(Trim(.TextMatrix(.rows - 1, 0)) = "", .rows - 2, .rows - 1)
            dbl包装 = Val(.TextMatrix(intCount, menuPriceCol.包装系数))

            '更新指导零售价
            If Val(.TextMatrix(intCount, menuPriceCol.原指导售价)) <> Val(.TextMatrix(intCount, menuPriceCol.现指导售价)) And Val(.TextMatrix(intCount, menuPriceCol.现指导售价)) <> 0 Then
                strUpdate = Val(Round(Val(.TextMatrix(intCount, menuPriceCol.现指导售价)) / dbl包装, mintSalePriceDigit))

                gstrSQL = "zl_药品目录_UpdateCustom(" & Val(.TextMatrix(intCount, menuPriceCol.药品id)) & ",'指导零售价=" & strUpdate & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, MStrCaption)
            End If

            '更新采购限价
            If Val(.TextMatrix(intCount, menuPriceCol.原采购限价)) <> Val(.TextMatrix(intCount, menuPriceCol.现采购限价)) And Val(.TextMatrix(intCount, menuPriceCol.现采购限价)) <> 0 Then
                strUpdate = Val(Round(Val(.TextMatrix(intCount, menuPriceCol.现采购限价)) / dbl包装, mintSalePriceDigit))

                gstrSQL = "zl_药品目录_UpdateCustom(" & Val(.TextMatrix(intCount, menuPriceCol.药品id)) & ",'指导批发价=" & strUpdate & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, MStrCaption)
            End If
        Next
    End With

    '产生调价汇总记录
    If blnPrice = True And blnCost = True Then
        intUpdateModel = 2
    ElseIf blnPrice = True And blnCost = False Then
        intUpdateModel = 0
    ElseIf blnPrice = False And blnCost = True Then
        intUpdateModel = 1
    End If

    gstrSQL = "Zl_调价汇总记录_Insert(" & txtNO.Text & "," & intUpdateModel & ","
    If opt时间(0).Value = True Then
        gstrSQL = gstrSQL & "sysdate" & ","
    Else
        gstrSQL = gstrSQL & "to_date('" & Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
    End If
    gstrSQL = gstrSQL & IIf(txtSummary.Text = "", "Null", "'" & txtSummary.Text & "'") & ",0,'" & UserInfo.用户姓名 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, MStrCaption)

    gcnOracle.CommitTrans

    If blnPrint = True Then
        If MsgBox("你需要打印调价通知单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1333", Me, "NO=" & txtNO.Text, "包装单位=" & mintUnit, 2)
        End If
    End If

    '清空列表中数据
    With vsfPrice
        .rows = 2
        For intCol = 0 To .Cols - 1
            .TextMatrix(1, intCol) = ""
        Next
    End With
    vsfStore.rows = 1
    vsfPay.rows = 1
    txtNO.Text = ""
    txtSummary.Text = ""

    Exit Sub

ErrHand:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Function CheckUnVerify(ByVal lng药品id As Long) As Boolean
    '检查药品是否存在未审核单据
    Dim rsTemp As ADODB.Recordset

    On Error GoTo errHandle
    gstrSQL = "Select 1 From 药品收发记录 Where 药品id = [1] And Rownum = 1 And 审核日期 Is Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查药品是否存在未审核单据", lng药品id)

    If rsTemp.RecordCount > 0 Then
        CheckUnVerify = True
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function checkNotExecutePrice(Optional ByVal lngDrugID As Long = 0, Optional ByRef strInfo As String) As Boolean
    '功能 ：检查是否存在未执行的价格
    Dim RecCheck As New ADODB.Recordset
    Dim LngmediIDThis As Long, IntCheck As Integer

    Err = 0
    On Error GoTo ErrHand

    If lngDrugID = 0 Then
        '循环判断所有药品
        For IntCheck = 1 To vsfPrice.rows - 1
            LngmediIDThis = Val(vsfPrice.TextMatrix(IntCheck, menuPriceCol.药品id))
            If LngmediIDThis <> 0 Then
                If mint调价 = 0 Or mint调价 = 2 Then
                    '判断是否有未执行的历史价格
                    gstrSQL = " Select Count(*) Records From 收费价目 Where 变动原因=0 And 执行日期 > Sysdate And 收费细目ID=[1]" & _
                            GetPriceClassString("")
                    
                    Set RecCheck = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption, LngmediIDThis)

                    With RecCheck
                        If Not .EOF Then
                            If Not IsNull(!Records) Then
                                If !Records <> 0 Then
                                    strInfo = "药品" & vsfPrice.TextMatrix(IntCheck, menuPriceCol.药品) & "存在未执行价格，未执行药品不能调价！"
                                    checkNotExecutePrice = True
                                    Exit Function
                                End If
                            End If
                        End If
                    End With
                End If

                If mint调价 = 1 Or mint调价 = 2 Then
                    '检查是否还有未执行的成本价调价计划
                    gstrSQL = "Select 1 From 药品价格记录 Where 价格类型=2 And 记录状态=0 And 药品id = [1] And Rownum < 2 "
                    Set RecCheck = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption, LngmediIDThis)

                    If RecCheck.RecordCount > 0 Then
                        strInfo = "药品" & vsfPrice.TextMatrix(IntCheck, menuPriceCol.药品) & "存在未执行成本价，未执行药品不能调价！"
                        checkNotExecutePrice = True
                        Exit Function
                    End If
                End If
            End If
        Next
    Else
        If mint调价 = 0 Or mint调价 = 2 Then
            '判断是否有未执行的历史价格
            gstrSQL = " Select Count(*) Records From 收费价目 Where 变动原因=0 And 执行日期 > Sysdate And 收费细目ID=[1]" & _
                    GetPriceClassString("")
            
            Set RecCheck = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption, lngDrugID, gstrPriceClass)

            With RecCheck
                If Not .EOF Then
                    If Not IsNull(!Records) Then
                        If !Records <> 0 Then
                            strInfo = "还存在未执行的售价调价记录，未执行药品不能调价！"
                            checkNotExecutePrice = True
                            Exit Function
                        End If
                    End If
                End If
            End With
        End If

        If mint调价 = 1 Or mint调价 = 2 Then
            '检查是否还有未执行的成本价调价计划
            gstrSQL = "Select 1 From 药品价格记录 Where 价格类型=2 And 记录状态=0 And 药品id = [1] And Rownum < 2 "
            Set RecCheck = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption, lngDrugID)

            If RecCheck.RecordCount > 0 Then
                strInfo = "还存在未执行的成本价调价，未执行药品不能调价！"
                checkNotExecutePrice = True
                Exit Function
            End If
        End If
    End If


    checkNotExecutePrice = False
    Exit Function
ErrHand:
    Call ErrCenter
    Call SaveErrLog
    Me.vsfPrice.SetFocus

End Function

Private Function CheckPrice() As Boolean
    Dim IntCheck As Integer
    Dim n As Integer
    Dim strTmp As String
    Dim bln无库存 As Boolean
    Dim dbl包装 As Double
    Dim bln有无库存 As Boolean
    Dim lngDouID As Long
    Dim strMsg As String '记录提示信息
    Dim intCount2 As Integer '用来计数
    
    '检测各执行价格是否正确
    '以及收入项目相同的情况下现价是否与原价相同
    CheckPrice = False
    With vsfPrice
        For IntCheck = 1 To .rows - 1
            If Val(.TextMatrix(IntCheck, menuPriceCol.药品id)) <> 0 Then
                If Not IsNumeric(Trim(.TextMatrix(IntCheck, menuPriceCol.现零售价))) Then
                    MsgBox "第" & IntCheck & "行的药品售价中含有非法字符！", vbInformation, gstrSysName
                    .Row = IntCheck
                    .Col = menuPriceCol.现零售价
                    vsfPrice.SetFocus
                    .Select IntCheck, 0, IntCheck, .Cols - 1
                    .TopRow = IntCheck
                    Exit Function
                End If
                
                '检查价格是否为空
                If .TextMatrix(IntCheck, menuPriceCol.现零售价) = "" Or .TextMatrix(IntCheck, menuPriceCol.原零售价) = "" Or .TextMatrix(IntCheck, menuPriceCol.现成本价) = "" Or .TextMatrix(IntCheck, menuPriceCol.原成本价) = "" Then
                    MsgBox "第" & IntCheck & "行的药品有价格为空，不能执行调价！", vbInformation, gstrSysName
                    .Row = IntCheck
                    vsfPrice.SetFocus
                    .Select IntCheck, 0, IntCheck, .Cols - 1
                    .TopRow = IntCheck
                    Exit Function
                End If
                For n = 1 To vsfStore.rows - 1
                    If Val(.TextMatrix(IntCheck, menuPriceCol.药品id)) = Val(vsfStore.TextMatrix(n, menuStoreCol.药品id)) Then
                        If vsfStore.TextMatrix(n, menuStoreCol.现零售价) = "" Or vsfStore.TextMatrix(n, menuStoreCol.原零售价) = "" Or vsfStore.TextMatrix(n, menuStoreCol.现成本价) = "" Or vsfStore.TextMatrix(n, menuStoreCol.原成本价) = "" Then
                            MsgBox "第" & IntCheck & "行的药品有价格为空，不能执行调价！", vbInformation, gstrSysName
                            .Row = IntCheck
                            vsfPrice.SetFocus
                            .Select IntCheck, 0, IntCheck, .Cols - 1
                            .TopRow = IntCheck
                            Exit Function
                        End If
                    End If
                Next
                
                '检查售价是否相同
                If mint调价 = 0 Or mint调价 = 2 Then
                    strTmp = ""
                    bln有无库存 = False
                    dbl包装 = Val(.TextMatrix(IntCheck, menuPriceCol.包装系数))
                    If .TextMatrix(IntCheck, menuPriceCol.是否变价) = "1" Then
                        For n = 1 To vsfStore.rows - 1
                            If Val(.TextMatrix(IntCheck, menuPriceCol.药品id)) = Val(vsfStore.TextMatrix(n, menuStoreCol.药品id)) Then
                                bln有无库存 = True
                                If InStr(1, "|" & strTmp, "|" & vsfStore.TextMatrix(n, menuStoreCol.批次) & ",") = 0 And vsfStore.TextMatrix(n, menuStoreCol.现零售价) <> vsfStore.TextMatrix(n, menuStoreCol.原零售价) Then
                                    strTmp = strTmp & IIf(strTmp = "", "", "|") & vsfStore.TextMatrix(n, menuStoreCol.批次) & "," & vsfStore.TextMatrix(n, menuStoreCol.现零售价) / dbl包装
                                End If
                            End If
                        Next
                        If strTmp = "" And bln有无库存 = True Then
                            MsgBox "第" & IntCheck & "行的药品现零售价与原零售价相同，不能执行调价！", vbInformation, gstrSysName
                            .Row = IntCheck
                            .Col = menuPriceCol.现零售价
                            vsfPrice.SetFocus
                            .Select IntCheck, 0, IntCheck, .Cols - 1
                            .TopRow = IntCheck
                            Exit Function
                        End If
                        If bln有无库存 = False And .TextMatrix(IntCheck, menuPriceCol.现零售价) = .TextMatrix(IntCheck, menuPriceCol.原零售价) Then
                            MsgBox "第" & IntCheck & "行的药品现零售价与原零售价相同，不能执行调价！", vbInformation, gstrSysName
                            .Row = IntCheck
                            .Col = menuPriceCol.现零售价
                            vsfPrice.SetFocus
                            .Select IntCheck, 0, IntCheck, .Cols - 1
                            .TopRow = IntCheck
                            Exit Function
                        End If
                    End If
                    If .TextMatrix(IntCheck, menuPriceCol.是否变价) <> "1" And .TextMatrix(IntCheck, menuPriceCol.现零售价) = .TextMatrix(IntCheck, menuPriceCol.原零售价) Then
                        MsgBox "第" & IntCheck & "行的药品现零售价与原零售价相同，不能执行调价！", vbInformation, gstrSysName
                        .Row = IntCheck
                        .Col = menuPriceCol.现零售价
                        vsfPrice.SetFocus
                        .Select IntCheck, 0, IntCheck, .Cols - 1
                        .TopRow = IntCheck
                        Exit Function
                    End If
                End If
                
                '检查成本价是否相同
                If mint调价 = 1 Or mint调价 = 2 Then
                    bln有无库存 = False
                    strTmp = ""
                    For n = 1 To vsfStore.rows - 1
                        If Val(.TextMatrix(IntCheck, menuPriceCol.药品id)) = Val(vsfStore.TextMatrix(n, menuStoreCol.药品id)) Then
                            bln有无库存 = True
                            If vsfStore.TextMatrix(n, menuStoreCol.现成本价) <> vsfStore.TextMatrix(n, menuStoreCol.原成本价) Then
                                strTmp = "调过成本价"
                            End If
                        End If
                    Next
                    If bln有无库存 = True And strTmp = "" Then
                        MsgBox "第" & IntCheck & "行的药品现成本价与原成本价相同，不能执行调价！", vbInformation, gstrSysName
                        .Row = IntCheck
                        .Col = menuPriceCol.现成本价
                        vsfPrice.SetFocus
                        .Select IntCheck, 0, IntCheck, .Cols - 1
                        .TopRow = IntCheck
                        Exit Function
                    End If
                    If bln有无库存 = False And .TextMatrix(IntCheck, menuPriceCol.现成本价) = .TextMatrix(IntCheck, menuPriceCol.原成本价) Then
                        MsgBox "第" & IntCheck & "行的药品现成本价与原成本价相同，不能执行调价！", vbInformation, gstrSysName
                        .Row = IntCheck
                        .Col = menuPriceCol.现成本价
                        vsfPrice.SetFocus
                        .Select IntCheck, 0, IntCheck, .Cols - 1
                        .TopRow = IntCheck
                        Exit Function
                    End If
                End If
                
                '零差价管理：检查调价后售价和成本价是否一致
                If gtype_UserSysParms.P275_零差价管理模式 > 0 Then
                    If IsPriceAdjustMod(Val(.TextMatrix(IntCheck, menuPriceCol.药品id))) = True Then
                        If .TextMatrix(IntCheck, menuPriceCol.是否有库存) = 0 Then
                            '无库存，直接比较价格表单中的售价和成本价

                            If Val(.TextMatrix(IntCheck, menuPriceCol.现零售价)) <> Val(.TextMatrix(IntCheck, menuPriceCol.现成本价)) Then
                                MsgBox "第" & IntCheck & "行的定价药品已启用零差价管理，新售价必须和库存成本价一致！", vbInformation, gstrSysName
                                .Row = IntCheck
                                .Col = menuPriceCol.现零售价
                                vsfPrice.SetFocus
                                .Select IntCheck, 0, IntCheck, .Cols - 1
                                .TopRow = IntCheck
                                Exit Function
                            End If
       
                        Else
                            If .TextMatrix(IntCheck, menuPriceCol.是否变价) = "0" Then
                                '定价，检查规格列表中的最新售价是否和库存表中的最新成本价一致
                                For n = 1 To vsfStore.rows - 1
                                    If Val(.TextMatrix(IntCheck, menuPriceCol.药品id)) = Val(vsfStore.TextMatrix(n, menuStoreCol.药品id)) Then
                                        If mint调价 = 0 Then
                                            '仅调售价方式
                                            If Val(.TextMatrix(IntCheck, menuPriceCol.现零售价)) <> Val(vsfStore.TextMatrix(n, menuStoreCol.原成本价)) Then
                                                MsgBox "第" & IntCheck & "行的定价药品已启用零差价管理，新售价必须和库存成本价一致！", vbInformation, gstrSysName
                                                .Row = IntCheck
                                                .Col = menuPriceCol.现零售价
                                                vsfPrice.SetFocus
                                                .Select IntCheck, 0, IntCheck, .Cols - 1
                                                .TopRow = IntCheck
                                                Exit Function
                                            End If
                                        ElseIf mint调价 = 1 Then
                                            '仅调成本价方式
                                            If Val(.TextMatrix(IntCheck, menuPriceCol.原零售价)) <> Val(vsfStore.TextMatrix(n, menuStoreCol.现成本价)) Then
                                                MsgBox "第" & IntCheck & "行的定价药品已启用零差价管理，新成本价必须和售价一致！", vbInformation, gstrSysName
                                                .Row = IntCheck
                                                .Col = menuPriceCol.现零售价
                                                vsfPrice.SetFocus
                                                .Select IntCheck, 0, IntCheck, .Cols - 1
                                                .TopRow = IntCheck
                                                Exit Function
                                            End If
                                        Else
                                            '售价和成本价一起调方式
                                            If Val(.TextMatrix(IntCheck, menuPriceCol.现零售价)) <> Val(vsfStore.TextMatrix(n, menuStoreCol.现成本价)) Then
                                                MsgBox "第" & IntCheck & "行的定价药品已启用零差价管理，新售价必须和库存新成本价一致！", vbInformation, gstrSysName
                                                .Row = IntCheck
                                                .Col = menuPriceCol.现零售价
                                                vsfPrice.SetFocus
                                                .Select IntCheck, 0, IntCheck, .Cols - 1
                                                .TopRow = IntCheck
                                                Exit Function
                                            End If
                                        End If
                                    End If
                                Next
                            Else
                                '时价，检查库存列表中的最新售价是否和库存表中的最新成本价一致
                                For n = 1 To vsfStore.rows - 1
                                    If Val(.TextMatrix(IntCheck, menuPriceCol.药品id)) = Val(vsfStore.TextMatrix(n, menuStoreCol.药品id)) Then
                                        If mint调价 = 0 Then
                                            '仅调售价方式
                                            If Val(vsfStore.TextMatrix(n, menuStoreCol.现零售价)) <> Val(vsfStore.TextMatrix(n, menuStoreCol.原成本价)) Then
                                                MsgBox "第" & IntCheck & "行的时价药品已启用零差价管理，新售价必须和库存成本价一致！", vbInformation, gstrSysName
                                                .Row = IntCheck
                                                .Col = menuPriceCol.现零售价
                                                vsfPrice.SetFocus
                                                .Select IntCheck, 0, IntCheck, .Cols - 1
                                                .TopRow = IntCheck
                                                Exit Function
                                            End If
                                        ElseIf mint调价 = 1 Then
                                            '仅调成本价方式
                                            If Val(vsfStore.TextMatrix(n, menuStoreCol.现成本价)) <> Val(vsfStore.TextMatrix(n, menuStoreCol.原零售价)) Then
                                                MsgBox "第" & IntCheck & "行的时价药品已启用零差价管理，新成本价必须和售价一致！", vbInformation, gstrSysName
                                                .Row = IntCheck
                                                .Col = menuPriceCol.现零售价
                                                vsfPrice.SetFocus
                                                .Select IntCheck, 0, IntCheck, .Cols - 1
                                                .TopRow = IntCheck
                                                Exit Function
                                            End If
                                        Else
                                            '售价和成本价一起调方式
                                            If Val(vsfStore.TextMatrix(n, menuStoreCol.现零售价)) <> Val(vsfStore.TextMatrix(n, menuStoreCol.现成本价)) Then
                                                MsgBox "第" & IntCheck & "行的时价药品已启用零差价管理，新售价必须和库存新成本价一致！", vbInformation, gstrSysName
                                                .Row = IntCheck
                                                .Col = menuPriceCol.现零售价
                                                vsfPrice.SetFocus
                                                .Select IntCheck, 0, IntCheck, .Cols - 1
                                                .TopRow = IntCheck
                                                Exit Function
                                            End If
                                        End If
                                    End If
                                Next
        
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End With

    '检查未审核单据
    If vsfStore.rows > 1 And (mint调价 = 1 Or mint调价 = 2) Then
        If vsfStore.TextMatrix(1, menuStoreCol.药品id) <> "" Then
            lngDouID = 0
            For n = 1 To vsfStore.rows - 1
                If vsfStore.TextMatrix(n, menuStoreCol.药品id) = "" Then Exit For
    
                If CheckUnVerify(Val(vsfStore.TextMatrix(n, menuStoreCol.药品id))) = True And Val(vsfStore.TextMatrix(n, menuStoreCol.药品id)) <> lngDouID Then
                    lngDouID = Val(vsfStore.TextMatrix(n, menuStoreCol.药品id))
                    strMsg = strMsg & vsfStore.TextMatrix(n, menuStoreCol.药品) & ","
                    intCount2 = intCount2 + 1
                    If intCount2 > 3 Then Exit For '只判断3个
                End If
            Next
    
            If strMsg <> "" Then
                If MsgBox(strMsg & "存在未审核单据，调整成本价可能会造成差价误差。" & _
                    vbCrLf & Space(4) & "建议先处理未审核单据。是否还继续调价？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            End If
        End If
    End If
                
    CheckPrice = True
End Function


Private Sub cmdPriceMethod_Click()
    If txt供应商.Tag = "" Then
        Me.txt供应商.Tag = "0|"
    End If
    picOtherSelect.Visible = True
End Sub

Private Sub PrintStore()
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    If vsfStore.rows = 1 Then
        MsgBox "没有库存变动记录！", vbInformation, gstrSysName
        Exit Sub
    End If
    If Trim(Me.vsfStore.TextMatrix(1, menuStoreCol.库房)) = "" Then Exit Sub

    objPrint.Title.Text = "调价库存变动表"

    Set objRow = New zlTabAppRow
    objRow.Add "调价说明:" & Me.txtSummary.Text
    objPrint.UnderAppRows.Add objRow

    Set objRow = New zlTabAppRow
    objRow.Add "执行时间:" & Format(IIf(opt时间(0).Value = True, Sys.Currentdate, Me.dtpRunDate.Value), "yyyy年MM月DD日 HH:mm:ss")
    objRow.Add "调价人:" & Me.txtValuer.Text
    objPrint.UnderAppRows.Add objRow

    Set objRow = New zlTabAppRow
    objRow.Add "打印人:" & gstrUserName
    objRow.Add "打印时间:" & Format(Sys.Currentdate, "yyyy年MM月DD日 HH:mm:ss")
    objPrint.BelowAppRows.Add objRow

    Set objPrint.Body = Me.vsfStore.Object
    objPrint.PageFooter = 2

    Select Case zlPrintAsk(objPrint)
    Case 1
         zlPrintOrView1Grd objPrint, 1
    Case 2
        zlPrintOrView1Grd objPrint, 2
    Case 3
        zlPrintOrView1Grd objPrint, 3
    End Select
    Set objPrint = Nothing
End Sub

Private Sub Cmd供应商_Click()
    Dim rsTemp As ADODB.Recordset

    On Error GoTo errHandle
    gstrSQL = "Select 编码,名称,简码,id" & _
        " From 供应商" & _
        " where 末级=1 And substr(类型,1,1) = '1' And (撤档时间 is null or 撤档时间=to_date('3000-01-01','YYYY-MM-DD')) " & _
        " Order By 编码 "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取供应商信息")
    If rsTemp.EOF Then
        MsgBox "请初始化供应商（字典管理）！", vbInformation, gstrSysName
        Exit Sub
    End If

    With Me.mshProvider
        .Left = chk供应商.Left
        .Top = txt供应商.Top + txt供应商.Height
        .Clear
        Set .DataSource = rsTemp
        .ColWidth(0) = 800: .ColWidth(1) = 2500: .ColWidth(2) = 800: .ColWidth(3) = 0
        .Row = 1: .ColSel = .Cols - 1
        .ZOrder 0: .Visible = True: .SetFocus
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub Form_Activate()
    If mblnLoad = False Then
        vsfPrice.SetFocus
    End If
    If mBlnClick = False Then
        vsfPrice.Row = 1
        vsfPrice.Col = menuPriceCol.药品
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        picOtherSelect.Visible = False
    End If
End Sub

Private Sub Form_Load()
    Dim StrToday As String
    Dim intUnitTemp As Integer
    Dim blnOldAjuset As Boolean '判断35.70之前的老调价模式
    Dim rsTemp As ADODB.Recordset
    
    Me.Height = 768 * 15
    Me.Width = 1024 * 15
    '获取设置的单位
    mintUnit = Val(zlDatabase.GetPara("药品单位", glngSys, 1333, "1"))
    mstrPrivs = GetPrivFunc(glngSys, 1333)
    Select Case mintUnit
        Case 0 '药库
            intUnitTemp = 4
        Case 2 '住院
            intUnitTemp = 3
        Case 1 '门诊
            intUnitTemp = 2
        Case 3 '售价
            intUnitTemp = 1
    End Select
    '获取各级单位精度
    mintCostDigit = GetDigitTiaoJia(1, 1, intUnitTemp)
    mintPriceDigit = GetDigitTiaoJia(1, 2, intUnitTemp)
    mintNumberDigit = GetDigitTiaoJia(1, 3, intUnitTemp)
    mintMoneyDigit = GetDigitTiaoJia(1, 4)
    mstrMoneyFormat = "0." & String(mintMoneyDigit, "0")
    mintSalePriceDigit = GetDigitTiaoJia(1, 2, 1)
    '初始化时间为当前时间+1天
    StrToday = Format(Sys.Currentdate(), "yyyy-MM-dd hh:mm:ss")

    If mintModal = 0 Then '新增的时候最小时间设置为当前时间+1天
        Me.dtpRunDate.MinDate = DateAdd("s", 1, CDate(StrToday))
    End If
    Me.dtpRunDate.Value = DateAdd("d", 1, CDate(StrToday))

    mbln时价药品按批次调价 = Val(zlDatabase.GetPara("时价药品按批次调价", glngSys, 1333, 0))
    mbln成本价按库房批次调整 = Val(zlDatabase.GetPara("成本价按库房批次调整", glngSys, 1333, 0))
    mbln现价提示 = Val(zlDatabase.GetPara("限价提示", glngSys, 1333, 1))

    txtValuer.Text = UserInfo.用户姓名  'gstrUserName

    txtNO.Text = IIf(mintModal = 0, "", mstr调价汇总号)
    If mintModal = 0 Then
        lblNO.Visible = False
        txtNO.Visible = False
    End If

    Call initComboBox '初始化下拉控件
    If mintModal = 1 Then '修改
        If (InStr(1, ";" & mstrPrivs & ";", ";成本价调价;") > 0 And InStr(1, ";" & mstrPrivs & ";", ";售价调价;") = 0) Or (InStr(1, ";" & mstrPrivs & ";", ";成本价调价;") = 0 And InStr(1, ";" & mstrPrivs & ";", ";售价调价;") > 0) Then
            cboPriceMethod.ListIndex = 0
        ElseIf (InStr(1, ";" & mstrPrivs & ";", ";成本价调价;") > 0 And InStr(1, ";" & mstrPrivs & ";", ";售价调价;") > 0) Then
            cboPriceMethod.ListIndex = mintMethod
        End If
    ElseIf mintModal = 2 Then '查阅
        cboPriceMethod.ListIndex = mintMethod
    End If

    Call initCommandBars
    
    Call InitTabControl
    Call InitVsfGridFlex

    Call RestoreWinState(Me, App.ProductName, MStrCaption)
    If mbln应付记录 = False Then
        TabCtlDetails.Item(1).Visible = False
    End If
    
    If mintModal <> 0 Then
        gstrSQL = "Select 1 from 药品价格记录 where 调价汇总号=[1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption, mstr调价汇总号)
        
        '判断是否35.70更新前数据，老数据调用老方法来显示；只调售价也用来方法显示
        If Not rsTemp.EOF Then
            Call initGrid
        Else
            Call initGrid_Old
        End If
    End If

    If mintModal = 2 Then '查阅
        Dim cbrControls As CommandBarControl
        Set cbrControls = cbsMain.FindControl(, mconMenu_Save)
        cbrControls.Enabled = False
        Set cbrControls = cbsMain.FindControl(, mconMenu_BatchSelect)
        cbrControls.Enabled = False
        Set cbrControls = cbsMain.FindControl(, mconMenu_ClearAll)
        cbrControls.Enabled = False
        Set cbrControls = cbsMain.FindControl(, mconMenu_Find)
        cbrControls.Enabled = False
    
        cboPriceMethod.Enabled = False
        cmdPriceMethod.Enabled = False
        opt时间(0).Enabled = False
        opt时间(1).Enabled = False
        dtpRunDate.Enabled = False
        cbo售价计算方式.Visible = False
        lbl调价方式.Visible = False
        chkAotuCost.Visible = False
        chkAotuCost.Enabled = False
        chkAutoPay.Enabled = False
        txtSummary.Enabled = False

        vsfPrice.Cell(flexcpBackColor, 1, 0, vsfPrice.rows - 1, vsfPrice.Cols - 1) = mconlngColor
        If vsfStore.rows > 1 Then
            vsfStore.Cell(flexcpBackColor, 1, 0, vsfStore.rows - 1, vsfStore.Cols - 1) = mconlngColor
        End If
        If vsfPay.rows > 1 Then
            vsfPay.Cell(flexcpBackColor, 0, 0, vsfPay.rows - 1, vsfPay.Cols - 1) = mconlngColor
        End If
    End If
    mblnLoad = True
End Sub

Private Sub initGrid_Old()
    '如果是修改或者查阅则提取相应的记录并填充到表格中
    Dim rsTemp As ADODB.Recordset
    Dim lngRow As Long
    Dim i As Long
    Dim lngDrugID As Long
    Dim db包装系数 As Double
    Dim strUnit As String
    Dim StrToday As String
    Dim rs产地 As ADODB.Recordset

    On Error GoTo errHandle
    '调价方式 0-调售价;1-调成本价;2-调售价及成本价
    If mintMethod = 0 Then
        gstrSQL = "Select Distinct p.原价id, i.是否变价, Nvl(s.指导批发价, 0) As 指导批价, Nvl(s.扣率, 0) As 扣率, Nvl(s.指导零售价, 0) As 指导售价," & vbNewLine & _
            "                s.加成率/100 As 加成率, i.编码, b.名称 As 商品名, i.名称 As 通用名, i.规格, i.产地 As 产地, i.计算单位 As 单位," & vbNewLine & _
            "                s.门诊单位, s.门诊包装, s.住院单位, s.住院包装, s.药库单位, Nvl(s.药库包装, 1) 药库包装, s.成本价 As 原成本价, s.成本价 As 新成本价, p.原价, p.现价," & vbNewLine & _
            "                p.收入项目id, p.调价人, p.调价说明, s.差价让利比, To_Char(a.执行日期, 'YYYY-MM-DD HH24:MI:SS') As 执行日期, i.Id 药品id," & vbNewLine & _
            "                Decode(k.药品id, Null, 0, 1) 是否有库存" & vbNewLine & _
            "From (Select s.药品id From 药品库存 s where s.性质=1 And Not (zl_fun_getbatchpro(s.库房id,s.药品id)=1 And Nvl(S.批次,0) = 0 And S.可用数量 < 0 And S.实际数量 = 0 And S.实际金额 = 0 And S.实际差价 = 0)) K, 调价汇总记录 A, 收费项目别名 B, 药品规格 S, 收费项目目录 I, 收费价目 P" & vbNewLine & _
            "Where a.调价号 = p.调价汇总号 And b.收费细目id(+) = s.药品id And s.药品id = i.Id And i.Id = k.药品id(+) And i.Id = p.收费细目id And" & vbNewLine & _
            "      p.调价汇总号 = [1] And a.分类 = 0 And b.性质(+) = 3 And a.调价号 = [1] " & vbNewLine & _
            GetPriceClassString("P") & vbNewLine & _
            IIf(mintModal = 2, "", "  And (i.撤档时间 Is Null Or i.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))") & vbNewLine & _
            "Order By 药品id"
    ElseIf mintMethod = 1 Then
        gstrSQL = "Select Distinct i.是否变价, Nvl(s.指导批发价, 0) As 指导批价, Nvl(s.扣率, 0) As 扣率, Nvl(s.指导零售价, 0) As 指导售价," & vbNewLine & _
            "                s.加成率/100 As 加成率, i.编码, b.名称 As 商品名, i.名称 As 通用名, i.规格, i.产地 As 产地, i.计算单位 As 单位," & vbNewLine & _
            "                s.门诊单位, s.门诊包装, s.住院单位, s.住院包装, s.药库单位, Nvl(s.药库包装, 1) 药库包装, m.原成本价, m.新成本价, p.现价 as 原价, p.现价, p.收入项目id," & vbNewLine & _
            "                a.填制人 As 调价人, a.说明 As 调价说明, s.差价让利比, To_Char(m.执行日期, 'YYYY-MM-DD HH24:MI:SS') As 执行日期, i.Id 药品id," & vbNewLine & _
            "                Decode(k.药品id, Null, 0, 1) 是否有库存" & vbNewLine & _
            "From (Select Min(原成本价) As 原成本价, Min(新成本价) As 新成本价, min(产地) as 产地,调价汇总号,药品id,min(执行日期) as 执行日期 From 成本价调价信息 Where 调价汇总号 = [1] Group By 调价汇总号,药品id) M, (Select s.药品id From 药品库存 s where s.性质=1 And Not (zl_fun_getbatchpro(s.库房id,s.药品id)=1 And Nvl(S.批次,0) = 0 And S.可用数量 < 0 And S.实际数量 = 0 And S.实际金额 = 0 And S.实际差价 = 0)) K, 调价汇总记录 A, 收费项目别名 B, 药品规格 S, 收费项目目录 I, 收费价目 P" & vbNewLine & _
            "Where m.调价汇总号(+) = a.调价号 And b.收费细目id(+) = s.药品id And s.药品id = i.Id And i.Id = k.药品id(+) And m.药品id = i.Id And" & vbNewLine & _
            "      i.Id = p.收费细目id And Sysdate Between p.执行日期 And p.终止日期 And m.调价汇总号 = [1] And a.分类 = 0 And b.性质(+) = 3 And" & vbNewLine & _
            "      a.调价号 = [1] " & IIf(mintModal = 2, "", " And (i.撤档时间 Is Null Or i.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))") & vbNewLine & _
            GetPriceClassString("P") & vbNewLine & _
            "Order By 药品id"
    ElseIf mintMethod = 2 Then
        gstrSQL = "Select distinct p.原价id, i.是否变价, Nvl(s.指导批发价, 0) As 指导批价, Nvl(s.扣率, 0) As 扣率, Nvl(s.指导零售价, 0) As 指导售价," & vbNewLine & _
            "       s.加成率/100 As 加成率, i.编码, b.名称 As 商品名, i.名称 As 通用名, i.规格, i.产地 As 产地, i.计算单位 As 单位, s.门诊单位," & vbNewLine & _
            "       s.门诊包装, s.住院单位, s.住院包装, s.药库单位, Nvl(s.药库包装, 1) 药库包装, m.原成本价, m.新成本价, p.原价, p.现价, p.收入项目id, p.调价人, p.调价说明, s.差价让利比," & vbNewLine & _
            "       To_Char(p.执行日期, 'YYYY-MM-DD HH24:MI:SS') As 执行日期, i.Id 药品id, Decode(k.药品id, Null, 0, 1) 是否有库存" & vbNewLine & _
            "From (Select 药品id,Min(原成本价) As 原成本价, Min(新成本价) As 新成本价, min(产地) as 产地,调价汇总号 From 成本价调价信息 Where 调价汇总号 = [1] Group By 药品id,调价汇总号) M, 收费价目 P, 调价汇总记录 A, (Select s.药品id From 药品库存 s where s.性质=1 And Not (zl_fun_getbatchpro(s.库房id,s.药品id)=1 And Nvl(S.批次,0) = 0 And S.可用数量 < 0 And S.实际数量 = 0 And S.实际金额 = 0 And S.实际差价 = 0)) K, 收费项目别名 B, 药品规格 S, 收费项目目录 I" & vbNewLine & _
            "Where m.调价汇总号 = a.调价号 and m.药品id=i.id And p.调价汇总号 = a.调价号 And p.收费细目id = k.药品id(+) And p.收费细目id = b.收费细目id(+) And p.收费细目id = s.药品id And" & vbNewLine & _
            "      s.药品id = i.Id And a.调价号 =[1] And b.性质(+) = 3 " & vbNewLine & _
            GetPriceClassString("P") & vbNewLine & _
            IIf(mintModal = 2, "", "  And (i.撤档时间 Is Null Or i.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption, mstr调价汇总号)
    If rsTemp.RecordCount = 0 Then
        MsgBox "该调价记录已经被删除了！", vbInformation, gstrSysName
        Exit Sub
    End If

    With vsfPrice
        .rows = 2
        rsTemp.MoveFirst
        For i = 0 To rsTemp.RecordCount - 1
            If rsTemp!药品id <> lngDrugID Then
                Select Case mintUnit
                    Case 0
                        db包装系数 = rsTemp!药库包装
                        strUnit = rsTemp!药库单位
                    Case 2
                        db包装系数 = rsTemp!住院包装
                        strUnit = rsTemp!住院单位
                    Case 1
                        db包装系数 = rsTemp!门诊包装
                        strUnit = rsTemp!门诊单位
                    Case 3
                        db包装系数 = 1
                        strUnit = rsTemp!单位
                End Select

                lngDrugID = rsTemp!药品id
                If mintMethod = 0 Or mintMethod = 2 Then
                    .TextMatrix(.rows - 1, menuPriceCol.原价id) = IIf(IsNull(rsTemp!原价id), "", rsTemp!原价id)
                End If
                .TextMatrix(.rows - 1, menuPriceCol.药品id) = rsTemp!药品id

                If gint药品名称显示 = 1 Then
                    .TextMatrix(.rows - 1, menuPriceCol.药品) = "[" & rsTemp!编码 & "]" & IIf(IsNull(rsTemp!商品名), rsTemp!通用名, rsTemp!商品名)
                Else
                    .TextMatrix(.rows - 1, menuPriceCol.药品) = "[" & rsTemp!编码 & "]" & rsTemp!通用名
                End If
                .TextMatrix(.rows - 1, menuPriceCol.规格) = IIf(IsNull(rsTemp!规格), "", rsTemp!规格)
                .TextMatrix(.rows - 1, menuPriceCol.是否变价) = rsTemp!是否变价
                
'                If mintMethod = 1 Or mintMethod = 2 Then
'                    gstrSQL = "select min(产地) as 产地 from 成本价调价信息 where 调价汇总号=[1] and 药品id=[2]"
'                    Set rs产地 = zldatabase.OpenSQLRecord(gstrSQL, "产地查询", mstr调价汇总号, rsTemp!药品id)
'                    If rs产地.RecordCount > 0 Then
'                        .TextMatrix(.rows - 1, menuPriceCol.产地) = IIf(IsNull(rs产地!厂牌), "", rs产地!厂牌)
'                    End If
'                Else
                    .TextMatrix(.rows - 1, menuPriceCol.产地) = IIf(IsNull(rsTemp!产地), "", rsTemp!产地)
'                End If
                
                .TextMatrix(.rows - 1, menuPriceCol.单位) = strUnit
                .TextMatrix(.rows - 1, menuPriceCol.包装系数) = db包装系数

                .TextMatrix(.rows - 1, menuPriceCol.加成率) = rsTemp!加成率
                .TextMatrix(.rows - 1, menuPriceCol.差价让利比) = Nvl(rsTemp!差价让利比, 0)
                .TextMatrix(.rows - 1, menuPriceCol.是否有库存) = rsTemp!是否有库存
                .TextMatrix(.rows - 1, menuPriceCol.收入项目ID) = IIf(IsNull(rsTemp!收入项目ID), "", rsTemp!收入项目ID)
                .TextMatrix(.rows - 1, menuPriceCol.原成本价) = zlStr.FormatEx(Nvl(rsTemp!原成本价, 0) * db包装系数, mintCostDigit, , True)
                .TextMatrix(.rows - 1, menuPriceCol.现成本价) = zlStr.FormatEx(rsTemp!新成本价 * db包装系数, mintCostDigit, , True)
                .TextMatrix(.rows - 1, menuPriceCol.原零售价) = zlStr.FormatEx(IIf(IsNull(rsTemp!原价), rsTemp!现价, rsTemp!原价) * db包装系数, mintPriceDigit, , True)
                .TextMatrix(.rows - 1, menuPriceCol.现零售价) = zlStr.FormatEx(rsTemp!现价 * db包装系数, mintPriceDigit, , True)
                .TextMatrix(.rows - 1, menuPriceCol.原采购限价) = zlStr.FormatEx(rsTemp!指导批价 * db包装系数, mintCostDigit, , True)
                .TextMatrix(.rows - 1, menuPriceCol.现采购限价) = zlStr.FormatEx(rsTemp!指导批价 * db包装系数, mintCostDigit, , True)
                .TextMatrix(.rows - 1, menuPriceCol.原指导售价) = zlStr.FormatEx(rsTemp!指导售价 * db包装系数, mintPriceDigit, , True)
                .TextMatrix(.rows - 1, menuPriceCol.现指导售价) = zlStr.FormatEx(rsTemp!指导售价 * db包装系数, mintPriceDigit, , True)

                txtValuer.Text = IIf(IsNull(rsTemp!调价人), "", rsTemp!调价人)
                txtSummary.Text = IIf(IsNull(rsTemp!调价说明), "", rsTemp!调价说明)
                If mintModal = 1 Then
                    Me.dtpRunDate.MinDate = CDate(rsTemp!执行日期)
                End If
                If IsNull(rsTemp!执行日期) Then
                    StrToday = Format(Sys.Currentdate(), "yyyy-MM-dd hh:mm:ss")
                Else
                    StrToday = Format(rsTemp!执行日期, "yyyy-MM-dd hh:mm:ss")
                End If
                Me.dtpRunDate.Value = CDate(StrToday)

                .rows = .rows + 1
                Call setColEdit
                .RowHeight(.rows - 1) = mlngRowHeight
            End If
            rsTemp.MoveNext
        Next
        Call GetDrugStore_Old(Val(.TextMatrix(1, menuPriceCol.药品id)), 1)
    End With

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetDrugStore_Old(ByVal lngDrugID As Long, ByVal intRow As Integer)
    Dim rsTemp As ADODB.Recordset
    Dim dblOldCost As Double
    Dim dblOldPrice As Double
    Dim dblNewCost As Double
    Dim dblNewPrice As Double
    Dim dbl加成率 As Double
    Dim lngCurRow As Long     '当前行
    Dim i As Long
    Dim dbl发票金额 As Double
    Dim str药品名称 As String
    Dim str发票 As String
    Dim str发票日期 As String
    Dim rsPirce As ADODB.Recordset
    Dim rsCost As ADODB.Recordset
    Dim dbl包装换算 As Double
    Dim bln相同药品 As Boolean
    Dim lng药品id As Long
    Dim str单位 As String

    '功能：为库存列表填充数据
    '参数：药品id

    On Error GoTo errHandle
    '先检查是否有重复的数据，如果有就先清除掉重复的数据
    With vsfStore
        For i = .rows - 1 To 1 Step -1
            If Val(.TextMatrix(i, menuStoreCol.药品id)) = mlngOldDrugID And mlngOldDrugID <> 0 Then
                .RemoveItem i
            End If
        Next
    End With

    With vsfPay
        For i = .rows - 1 To 1 Step -1
            If Val(.TextMatrix(i, menuPayCol.药品id)) = mlngOldDrugID And mlngOldDrugID <> 0 Then
                .RemoveItem i
            End If
        Next
    End With

    If mintModal = 0 Or mblnUpdateAdd = True Or mblnBatchItem = True Then
        gstrSQL = "Select s.库房id,s.药品id, d.名称 As 库房, '[' || m.编码 || ']' || m.名称 As 药品, m.规格, m.产地, m.计算单位 售价单位, p.药库单位, s.上次批号 As 批号, nvl(s.实际数量,0) As 数量," & vbNewLine & _
            "       s.批次, Nvl(m.是否变价, 0) 变价, m.Id, Decode(Nvl(m.是否变价, 0), 0, e.现价, Nvl(S.零售价,0)) 时价售价, p.加成率," & vbNewLine & _
            "       nvl(s.平均成本价,p.成本价) As 成本价, s.上次供应商id, n.名称 As 供应商, s.效期, s.上次产地 As 产地" & vbNewLine & _
            " From 药品库存 S, 部门表 D, 收费项目目录 M, 药品规格 P, 供应商 N, 收费价目 E" & vbNewLine & _
            " Where d.Id = s.库房id And s.药品id = m.Id And m.Id = p.药品id And Nvl(s.上次供应商id, 0) = n.Id(+) And m.Id = e.收费细目id And" & vbNewLine & _
            " s.性质 = 1 And s.药品id = [1] And Sysdate Between e.执行日期 And e.终止日期  " & vbNewLine & _
            GetPriceClassString("E") & vbNewLine & _
            " Order By  s.药品id,s.库房id, s.上次批号,s.批次 "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption, lngDrugID)

        If mlng供应商ID > 0 Then
            rsTemp.Filter = "上次供应商ID=" & mlng供应商ID
        End If
    Else '修改，查阅
        If mintModal = 2 Then   '查阅
            If cboPriceMethod.Text = "仅调成本价" Or cboPriceMethod.Text = "售价成本价一起调价" Then
                gstrSQL = "select (sysdate-执行日期 ) as 是否执行 from 调价汇总记录 where 调价号=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "是否执行", txtNO.Text)
                If rsTemp!是否执行 > 0 Then
                    gstrSQL = "Select Distinct a.库房id, c.名称 As 库房, b.药品id, b.供药单位id As 上次供应商id, '[' || e.编码 || ']' || e.名称 As 药品, e.规格, d.名称 As 供应商," & vbNewLine & _
                            "                b.新成本价, b.原成本价, b.发票号, b.发票日期, b.发票金额, b.产地, b.批次, b.批号, e.是否变价 As 变价, e.计算单位 As 售价单位, f.药库单位," & vbNewLine & _
                            "                nvl(a.填写数量,0) As 数量, f.加成率, b.效期" & vbNewLine & _
                            "From 药品收发记录 A,成本价调价信息 B, 部门表 C, 供应商 D, 收费项目目录 E, 药品规格 F" & vbNewLine & _
                            "Where a.id=b.收发id And a.库房id = c.Id And b.供药单位id = d.Id(+) And" & vbNewLine & _
                            "      a.药品id = e.Id And e.Id = f.药品id And b.调价汇总号 = [1] and a.单据 = 5"
                Else
                    gstrSQL = "Select Distinct a.库房id,c.名称 as 库房, b.药品id,a.上次供应商id, '[' || e.编码 || ']' ||e.名称 as 药品,e.规格,d.名称 as 供应商, b.新成本价, b.原成本价, b.发票号, b.发票日期, b.发票金额" & _
                            " ,a.上次产地 as 产地,a.批次,a.上次批号 as 批号,e.是否变价 as 变价,e.计算单位 as 售价单位,f.药库单位,nvl(a.实际数量,0) as 数量,f.加成率,a.效期" & _
                            " From 药品库存 A,部门表 C,供应商 D,收费项目目录 E,药品规格 F," & _
                                 " (Select Distinct 药品id, 库房id, 批次, 批号, 效期, 产地, 原成本价, 新成本价, 发票号, 发票日期, 发票金额, 应付款变动, 执行日期" & _
                                   " From 成本价调价信息" & _
                                   " Where 调价汇总号 = [1]) B" & _
                            " Where a.药品id = b.药品id And a.库房id = b.库房id and nvl(a.批次,0)=nvl(b.批次,0) and a.库房id=c.id and a.上次供应商id=d.id(+) and a.药品id=e.id and e.id=f.药品id and a.性质=1 "
                End If
            ElseIf cboPriceMethod.Text = "仅调售价" Then
                gstrSQL = "select (sysdate-执行日期 ) as 是否执行 from 调价汇总记录 where 调价号=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "是否执行", txtNO.Text)
                If rsTemp!是否执行 > 0 Then
                    gstrSQL = "Select Distinct a.库房id, c.名称 As 库房, b.收费细目id As 药品id, a.供药单位id As 上次供应商id, '[' || e.编码 || ']' || e.名称 As 药品, e.规格," & vbNewLine & _
                            "                d.名称 As 供应商, f.成本价 As 新成本价, f.成本价 As 原成本价, '' 发票号, '' 发票日期, '' 发票金额, a.产地, a.批次, a.批号, e.是否变价 As 变价," & vbNewLine & _
                            "                e.计算单位 As 售价单位, f.药库单位, nvl(a.填写数量,0) As 数量, f.加成率, a.效期" & vbNewLine & _
                            "From 药品收发记录 A, 收费价目 B, 部门表 C, 供应商 D, 收费项目目录 E, 药品规格 F" & vbNewLine & _
                            "Where a.价格id = b.Id And a.库房id = c.Id And a.供药单位id = d.Id(+) And a.药品id = e.Id And e.Id = f.药品id And" & vbNewLine & _
                            "      b.调价汇总号 = [1] and a.单据=13 And a.费用id Is Null " & GetPriceClassString("B")
                Else
                    gstrSQL = "Select Distinct a.库房id, c.名称 As 库房, b.收费细目id As 药品id, a.上次供应商id, '[' || e.编码 || ']' || e.名称 As 药品, e.规格, d.名称 As 供应商," & _
                                            " nvl(a.平均成本价,f.成本价) As 新成本价, nvl(a.平均成本价,f.成本价) As 原成本价, '' 发票号, '' 发票日期, '' 发票金额, a.上次产地 As 产地, a.批次, a.上次批号 As 批号," & _
                                            " e.是否变价 As 变价, e.计算单位 As 售价单位, f.药库单位, nvl(a.实际数量,0) As 数量, f.加成率, a.效期" & _
                            " From 药品库存 A, 收费价目 B, 部门表 C, 供应商 D, 收费项目目录 E, 药品规格 F" & _
                            " Where a.药品id = b.收费细目id And a.库房id = c.Id And a.上次供应商id = d.Id(+) And a.药品id = e.Id And e.Id = f.药品id And a.性质 = 1  And" & _
                                  " b.调价汇总号 = [1]" & GetPriceClassString("B")
                End If
            End If
        Else '修改
            If cboPriceMethod.Text = "仅调成本价" Or cboPriceMethod.Text = "售价成本价一起调价" Then
                gstrSQL = "Select Distinct a.库房id,c.名称 as 库房, b.药品id,a.上次供应商id, '[' || e.编码 || ']' ||e.名称 as 药品,e.规格,d.名称 as 供应商, b.新成本价, b.原成本价, b.发票号, b.发票日期, b.发票金额" & _
                            " ,a.上次产地 as 产地,a.批次,a.上次批号 as 批号,e.是否变价 as 变价,e.计算单位 as 售价单位,f.药库单位,nvl(a.实际数量,0) as 数量,f.加成率,a.效期" & _
                            " From 药品库存 A,部门表 C,供应商 D,收费项目目录 E,药品规格 F," & _
                                 " (Select Distinct 药品id, 库房id, 批次, 批号, 效期, 产地, 原成本价, 新成本价, 发票号, 发票日期, 发票金额, 应付款变动, 执行日期" & _
                                   " From 成本价调价信息" & _
                                   " Where 调价汇总号 = [1]) B" & _
                            " Where a.药品id = b.药品id And a.库房id = b.库房id and nvl(a.批次,0)=nvl(b.批次,0) and a.库房id=c.id and a.上次供应商id=d.id(+) and a.药品id=e.id and e.id=f.药品id and a.性质=1 "
            ElseIf cboPriceMethod.Text = "仅调售价" Then
                gstrSQL = "Select Distinct a.库房id, c.名称 As 库房, b.收费细目id As 药品id, a.上次供应商id, '[' || e.编码 || ']' || e.名称 As 药品, e.规格, d.名称 As 供应商," & _
                                            " nvl(a.平均成本价,f.成本价) As 新成本价, nvl(a.平均成本价,f.成本价) As 原成本价, '' 发票号, '' 发票日期, '' 发票金额, a.上次产地 As 产地, a.批次, a.上次批号 As 批号," & _
                                            " e.是否变价 As 变价, e.计算单位 As 售价单位, f.药库单位, nvl(a.实际数量,0) As 数量, f.加成率, a.效期" & _
                            " From 药品库存 A, 收费价目 B, 部门表 C, 供应商 D, 收费项目目录 E, 药品规格 F" & _
                            " Where a.药品id = b.收费细目id And a.库房id = c.Id And a.上次供应商id = d.Id(+) And a.药品id = e.Id And e.Id = f.药品id And a.性质 = 1  And" & _
                                  " b.调价汇总号 = [1]" & GetPriceClassString("B")
            End If
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption, txtNO.Text)
    End If
    
    With vsfStore
        Do While Not rsTemp.EOF
            dbl包装换算 = 0
            dbl发票金额 = 0
            dblOldPrice = 0
            dblNewPrice = 0
            For i = 0 To vsfPrice.rows - 1
                If rsTemp!药品id = vsfPrice.TextMatrix(i, menuPriceCol.药品id) Then
                    dbl包装换算 = vsfPrice.TextMatrix(i, menuPriceCol.包装系数)
                    dblOldPrice = Val(vsfPrice.TextMatrix(i, menuPriceCol.原零售价))
                    dblNewPrice = Val(vsfPrice.TextMatrix(i, menuPriceCol.现零售价))
                    str单位 = vsfPrice.TextMatrix(i, menuPriceCol.单位)
                    Exit For
                End If
            Next
            .rows = .rows + 1
            Call setColEdit
            .RowHeight(.rows - 1) = mlngRowHeight

            '从空白行开始插入数据
            .TextMatrix(.rows - 1, menuStoreCol.药品id) = rsTemp!药品id
            .TextMatrix(.rows - 1, menuStoreCol.库房) = rsTemp!库房
            .TextMatrix(.rows - 1, menuStoreCol.库房id) = rsTemp!库房id
            .TextMatrix(.rows - 1, menuStoreCol.供应商) = Nvl(rsTemp!供应商, "")
            .TextMatrix(.rows - 1, menuStoreCol.供应商id) = IIf(mlng供应商ID > 0, mlng供应商ID, Nvl(rsTemp!上次供应商ID))
            .TextMatrix(.rows - 1, menuStoreCol.药品) = rsTemp!药品
            str药品名称 = rsTemp!药品

            .TextMatrix(.rows - 1, menuStoreCol.规格) = IIf(IsNull(rsTemp!规格), "", rsTemp!规格)
            .TextMatrix(.rows - 1, menuStoreCol.单位) = str单位
            .TextMatrix(.rows - 1, menuStoreCol.批号) = IIf(IsNull(rsTemp!批号), "", rsTemp!批号)
            .TextMatrix(.rows - 1, menuStoreCol.效期) = Format(IIf(IsNull(rsTemp!效期), "", rsTemp!效期), "YYYY-MM-DD")
            .TextMatrix(.rows - 1, menuStoreCol.产地) = IIf(IsNull(rsTemp!产地), "", rsTemp!产地)
            .TextMatrix(.rows - 1, menuStoreCol.数量) = zlStr.FormatEx(rsTemp!数量 / dbl包装换算, mintNumberDigit, , True)
            .TextMatrix(.rows - 1, menuStoreCol.包装系数) = dbl包装换算
            .TextMatrix(.rows - 1, menuStoreCol.批次) = Nvl(rsTemp!批次, 0)
            .TextMatrix(.rows - 1, menuStoreCol.变价) = rsTemp!变价


            If mintModal = 0 Or mblnUpdateAdd = True Or mblnBatchItem = True Then
                dblOldCost = IIf(IsNull(rsTemp!成本价), 0, rsTemp!成本价) * dbl包装换算

                If mdbl加成率 > 0 Then
                    dbl加成率 = Round(mdbl加成率 / 100, 7)
                ElseIf dblOldCost > 0 Then
                    dbl加成率 = Round(IIf(rsTemp!变价 = 1, rsTemp!时价售价 * dbl包装换算, dblOldPrice) / dblOldCost - 1, 7)
                Else
                    dbl加成率 = Round(rsTemp!加成率 / 100, 2)
                End If
                If 1 + dbl加成率 = 0 Then
                    dblNewCost = 0
                Else
                    dblNewCost = rsTemp!时价售价 * dbl包装换算 / (1 + dbl加成率)
                End If
                If dbl加成率 = -1 Then dbl加成率 = 0

                .TextMatrix(.rows - 1, menuStoreCol.原零售价) = zlStr.FormatEx(IIf(rsTemp!变价 = 1, rsTemp!时价售价 * dbl包装换算, dblOldPrice), mintPriceDigit, , True)
                .TextMatrix(.rows - 1, menuStoreCol.现零售价) = zlStr.FormatEx(IIf(rsTemp!变价 = 1, rsTemp!时价售价 * dbl包装换算, dblOldPrice), mintPriceDigit, , True)
                .TextMatrix(.rows - 1, menuStoreCol.售价盈亏) = Format(rsTemp!数量 / dbl包装换算 * (Val(.TextMatrix(.rows - 1, menuStoreCol.现零售价)) - Val(.TextMatrix(.rows - 1, menuStoreCol.原零售价))), mstrMoneyFormat)
                .TextMatrix(.rows - 1, menuStoreCol.加成率) = dbl加成率 * 100
                .TextMatrix(.rows - 1, menuStoreCol.原成本价) = zlStr.FormatEx(dblOldCost, mintCostDigit, , True)
                .TextMatrix(.rows - 1, menuStoreCol.现成本价) = zlStr.FormatEx(dblNewCost, mintCostDigit, , True)
                .TextMatrix(.rows - 1, menuStoreCol.成本盈亏) = Format((Val(.TextMatrix(.rows - 1, menuStoreCol.现成本价)) - Val(.TextMatrix(.rows - 1, menuStoreCol.原成本价))) * Val(.TextMatrix(.rows - 1, menuStoreCol.数量)), mstrMoneyFormat)
                dbl发票金额 = dbl发票金额 + (dblNewCost - dblOldCost) * Val(.TextMatrix(.rows - 1, menuStoreCol.数量))
                
                '为应付记录表赋值
                If mint调价 = 1 Or mint调价 = 2 Then
                    If vsfPay.rows > 1 Then
                        bln相同药品 = False
                        For i = 1 To vsfPay.rows - 1
                            If vsfPay.TextMatrix(i, menuPayCol.药品id) = rsTemp!药品id Then
                                bln相同药品 = True
                                Exit For
                            End If
                        Next
                        If bln相同药品 = True Then
                            vsfPay.TextMatrix(i, menuPayCol.发票金额) = zlStr.FormatEx(Val(vsfPay.TextMatrix(i, menuPayCol.发票金额)) + dbl发票金额, mintMoneyDigit, , True)
                        Else
                            vsfPay.rows = vsfPay.rows + 1
                            vsfPay.RowHeight(vsfPay.rows - 1) = mlngRowHeight
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.药品id) = rsTemp!药品id
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.药品) = str药品名称
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.发票号) = str发票
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.发票日期) = Format(str发票日期, "yyyy-mm-dd")
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.发票金额) = zlStr.FormatEx(dbl发票金额, mintMoneyDigit, , True)
                        End If
                    Else
                        vsfPay.rows = vsfPay.rows + 1
                        vsfPay.RowHeight(vsfPay.rows - 1) = mlngRowHeight
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.药品id) = rsTemp!药品id
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.药品) = str药品名称
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.发票号) = str发票
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.发票日期) = Format(str发票日期, "yyyy-mm-dd")
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.发票金额) = zlStr.FormatEx(dbl发票金额, mintMoneyDigit, , True)
                    End If
                End If
            Else
                If mintModal = 2 And (cboPriceMethod.Text = "仅调售价" Or cboPriceMethod.Text = "售价成本价一起调价") Then   '查阅
                    gstrSQL = "Select a.成本价 As 原价, a.零售价 As 现价" & vbNewLine & _
                        "From 药品收发记录 A, 收费价目 B" & vbNewLine & _
                        "Where a.价格id = b.Id And b.调价汇总号 = [1] And a.库房id = [2] And a.药品id = [3] And Nvl(a.批次, 0) = [4]" & _
                        GetPriceClassString("B")
                        
                    Set rsPirce = zlDatabase.OpenSQLRecord(gstrSQL, "获取售价", txtNO.Text, rsTemp!库房id, rsTemp!药品id, Nvl(rsTemp!批次, 0))
                    
                    If Not rsPirce.EOF Then
                        .TextMatrix(.rows - 1, menuStoreCol.原零售价) = zlStr.FormatEx(Val(rsPirce!原价) * dbl包装换算, mintPriceDigit, , True)
                        .TextMatrix(.rows - 1, menuStoreCol.现零售价) = zlStr.FormatEx(Val(rsPirce!现价) * dbl包装换算, mintPriceDigit, , True)
                        .TextMatrix(.rows - 1, menuStoreCol.售价盈亏) = Format(rsTemp!数量 / dbl包装换算 * (Val(.TextMatrix(.rows - 1, menuStoreCol.现零售价)) - Val(.TextMatrix(.rows - 1, menuStoreCol.原零售价))), mstrMoneyFormat)
                    Else
                        .TextMatrix(.rows - 1, menuStoreCol.原零售价) = zlStr.FormatEx(dblOldPrice, mintPriceDigit, , True)
                        .TextMatrix(.rows - 1, menuStoreCol.现零售价) = zlStr.FormatEx(dblNewPrice, mintPriceDigit, , True)
                        .TextMatrix(.rows - 1, menuStoreCol.售价盈亏) = Format(rsTemp!数量 / dbl包装换算 * (Val(.TextMatrix(.rows - 1, menuStoreCol.现零售价)) - Val(.TextMatrix(.rows - 1, menuStoreCol.原零售价))), mstrMoneyFormat)
                    End If
                    If cboPriceMethod.Text = "仅调售价" Then
                        gstrSQL = "Select 成本价" & vbNewLine & _
                                    "      From (Select 平均成本价 As 成本价" & vbNewLine & _
                                    "             From 药品库存" & vbNewLine & _
                                    "             Where 性质=1 And 库房id = [1] And 药品id = [2] And nvl(批次,0) = [3]" & vbNewLine & _
                                    "             Union All" & vbNewLine & _
                                    "             Select 成本价 From 药品规格 Where 药品id = [2])" & vbNewLine & _
                                    "      Where Rownum <= 1"

                        Set rsCost = zlDatabase.OpenSQLRecord(gstrSQL, "获取成本价", rsTemp!库房id, rsTemp!药品id, Nvl(rsTemp!批次, 0))
                        .TextMatrix(.rows - 1, menuStoreCol.原成本价) = zlStr.FormatEx(rsCost!成本价 * dbl包装换算, mintCostDigit, , True)
                        .TextMatrix(.rows - 1, menuStoreCol.现成本价) = zlStr.FormatEx(rsCost!成本价 * dbl包装换算, mintCostDigit, , True)
                        .TextMatrix(.rows - 1, menuStoreCol.成本盈亏) = Format(0, mstrMoneyFormat)
                    Else
                        .TextMatrix(.rows - 1, menuStoreCol.原成本价) = zlStr.FormatEx(Nvl(rsTemp!原成本价, 0) * dbl包装换算, mintCostDigit, , True)
                        .TextMatrix(.rows - 1, menuStoreCol.现成本价) = zlStr.FormatEx(rsTemp!新成本价 * dbl包装换算, mintCostDigit, , True)
                        .TextMatrix(.rows - 1, menuStoreCol.成本盈亏) = Format((rsTemp!新成本价 * dbl包装换算 - Nvl(rsTemp!原成本价, 0) * dbl包装换算) * Val(.TextMatrix(.rows - 1, menuStoreCol.数量)), mstrMoneyFormat)
                    End If
                Else '修改或者成本价调价
                    '定价直接从收费价目取现价，时价优先从库存取，如果没有则从收费价目取
                    If Nvl(rsTemp!变价, 0) = 1 Then
                        gstrSQL = "Select Nvl(s.零售价, Decode(Nvl(s.实际数量, 0), 0, 0, Nvl(s.实际金额, 0) / s.实际数量)) 时价售价" & vbNewLine & _
                        "From 药品库存 S" & vbNewLine & _
                        "Where s.性质=1 And s.库房id = [1] And s.药品id = [2] And nvl(s.批次,0) = [3]"
                        
                        Set rsPirce = zlDatabase.OpenSQLRecord(gstrSQL, "获取售价", rsTemp!库房id, rsTemp!药品id, Nvl(rsTemp!批次, 0))
                        If rsPirce.RecordCount > 0 Then
                            If rsPirce!时价售价 > 0 Then
                                .TextMatrix(.rows - 1, menuStoreCol.原零售价) = zlStr.FormatEx(rsPirce!时价售价 * dbl包装换算, mintPriceDigit, , True)
                                .TextMatrix(.rows - 1, menuStoreCol.现零售价) = zlStr.FormatEx(rsPirce!时价售价 * dbl包装换算, mintPriceDigit, , True)
                            Else
                                .TextMatrix(.rows - 1, menuStoreCol.原零售价) = zlStr.FormatEx(dblOldPrice, mintPriceDigit, , True)
                                .TextMatrix(.rows - 1, menuStoreCol.现零售价) = zlStr.FormatEx(dblNewPrice, mintPriceDigit, , True)
                            End If
                        Else
                            .TextMatrix(.rows - 1, menuStoreCol.原零售价) = zlStr.FormatEx(dblOldPrice, mintPriceDigit, , True)
                            .TextMatrix(.rows - 1, menuStoreCol.现零售价) = zlStr.FormatEx(dblNewPrice, mintPriceDigit, , True)
                        End If
                    Else
                        .TextMatrix(.rows - 1, menuStoreCol.原零售价) = zlStr.FormatEx(dblOldPrice, mintPriceDigit, , True)
                        .TextMatrix(.rows - 1, menuStoreCol.现零售价) = zlStr.FormatEx(dblNewPrice, mintPriceDigit, , True)
                    End If
                    .TextMatrix(.rows - 1, menuStoreCol.售价盈亏) = Format(rsTemp!数量 / dbl包装换算 * (Val(.TextMatrix(.rows - 1, menuStoreCol.现零售价)) - Val(.TextMatrix(.rows - 1, menuStoreCol.原零售价))), mstrMoneyFormat)
                    .TextMatrix(.rows - 1, menuStoreCol.原成本价) = zlStr.FormatEx(Nvl(rsTemp!原成本价, 0) * dbl包装换算, mintCostDigit, , True)
                    .TextMatrix(.rows - 1, menuStoreCol.现成本价) = zlStr.FormatEx(rsTemp!新成本价 * dbl包装换算, mintCostDigit, , True)
                    .TextMatrix(.rows - 1, menuStoreCol.成本盈亏) = Format((rsTemp!新成本价 * dbl包装换算 - Nvl(rsTemp!原成本价, 0) * dbl包装换算) * Val(.TextMatrix(.rows - 1, menuStoreCol.数量)), mstrMoneyFormat)
                End If
                 
                If cboPriceMethod.Text = "仅调成本价" Or cboPriceMethod.Text = "售价成本价一起调价" Then
                    If rsTemp!新成本价 = 0 Then
                        dbl加成率 = 0
                    Else
                        dbl加成率 = Round(Val(.TextMatrix(.rows - 1, menuStoreCol.现零售价)) / (rsTemp!新成本价 * dbl包装换算) - 1, 7)
                    End If
                    .TextMatrix(.rows - 1, menuStoreCol.加成率) = dbl加成率 * 100
                    .TextMatrix(.rows - 1, menuStoreCol.原成本价) = zlStr.FormatEx(Nvl(rsTemp!原成本价, 0) * dbl包装换算, mintCostDigit, , True)
                    .TextMatrix(.rows - 1, menuStoreCol.现成本价) = zlStr.FormatEx(rsTemp!新成本价 * dbl包装换算, mintCostDigit, , True)
                    .TextMatrix(.rows - 1, menuStoreCol.成本盈亏) = Format((rsTemp!新成本价 * dbl包装换算 - Nvl(rsTemp!原成本价, 0) * dbl包装换算) * Val(.TextMatrix(.rows - 1, menuStoreCol.数量)), mstrMoneyFormat)
                    dbl发票金额 = dbl发票金额 + (rsTemp!新成本价 * dbl包装换算 - Nvl(rsTemp!原成本价, 0) * dbl包装换算) * Val(.TextMatrix(.rows - 1, menuStoreCol.数量))
                    str发票 = IIf(IsNull(rsTemp!发票号), "", rsTemp!发票号)
                    str发票日期 = IIf(IsNull(rsTemp!发票日期), "", rsTemp!发票日期)
                    
                    '为付款记录列表赋值
                    If vsfPay.rows > 1 Then
                        bln相同药品 = False
                        For i = 1 To vsfPay.rows - 1
                            If vsfPay.TextMatrix(i, menuPayCol.药品id) = rsTemp!药品id Then
                                bln相同药品 = True
                                Exit For
                            End If
                        Next
                        If bln相同药品 = True Then
                            vsfPay.TextMatrix(i, menuPayCol.发票金额) = zlStr.FormatEx(Val(vsfPay.TextMatrix(i, menuPayCol.发票金额)) + dbl发票金额, mintMoneyDigit, , True)
                        Else
                            vsfPay.rows = vsfPay.rows + 1
                            vsfPay.RowHeight(vsfPay.rows - 1) = mlngRowHeight
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.药品id) = rsTemp!药品id
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.药品) = str药品名称
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.发票号) = str发票
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.发票日期) = Format(str发票日期, "yyyy-mm-dd")
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.发票金额) = zlStr.FormatEx(dbl发票金额, mintMoneyDigit, , True)
                        End If
                    Else
                        vsfPay.rows = vsfPay.rows + 1
                        vsfPay.RowHeight(vsfPay.rows - 1) = mlngRowHeight
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.药品id) = rsTemp!药品id
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.药品) = str药品名称
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.发票号) = str发票
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.发票日期) = Format(str发票日期, "yyyy-mm-dd")
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.发票金额) = zlStr.FormatEx(dbl发票金额, mintMoneyDigit, , True)
                    End If
                End If
            End If
            rsTemp.MoveNext
        Loop
    End With
    '修改和查阅时重算规格列表平均成本价，售价
    'mintModal 0-新增 1-修改 2-查阅
    If mintModal = 1 Or mintModal = 2 Then
        With vsfStore
            For i = 1 To .rows - 1
                If lng药品id <> .TextMatrix(i, menuStoreCol.药品id) Then
                    Call CaluateAverCost(Val(.TextMatrix(i, menuStoreCol.药品id)))
                    Call CaluateAverOldCost(Val(.TextMatrix(i, menuStoreCol.药品id)))
                    Call CaculateAverPirce(Val(.TextMatrix(i, menuStoreCol.药品id)))
                    Call CaculateAverOldPirce(Val(.TextMatrix(i, menuStoreCol.药品id)))
                    lng药品id = Val(.TextMatrix(i, menuStoreCol.药品id))
                End If
            Next
        End With
    End If

    If mint调价 = 1 Or mint调价 = 2 Then
        If rsTemp.RecordCount = 0 Then Exit Sub
        TabCtlDetails.Item(1).Visible = True
    End If

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub initComboBox()
    With cbo售价计算方式
        .AddItem "售价与成本价不关联计算"
        .AddItem "售价按固定比例计算"
        .AddItem "售价按分段加成计算"
        .ListIndex = 0
    End With

    With cboPriceMethod
        If mintModal <> 2 Then  '非查阅
            If InStr(1, ";" & mstrPrivs & ";", ";成本价调价;") > 0 And InStr(1, ";" & mstrPrivs & ";", ";售价调价;") = 0 Then
                .AddItem "仅调成本价"
                .ListIndex = 0
                lblMethod.Tag = 0
            ElseIf InStr(1, ";" & mstrPrivs & ";", ";成本价调价;") = 0 And InStr(1, ";" & mstrPrivs & ";", ";售价调价;") > 0 Then
                .AddItem "仅调售价"
                .ListIndex = 0
                lblMethod.Tag = 0
            ElseIf InStr(1, ";" & mstrPrivs & ";", ";成本价调价;") > 0 And InStr(1, ";" & mstrPrivs & ";", ";售价调价;") > 0 Then
                .AddItem "仅调售价"
                .AddItem "仅调成本价"
                .AddItem "售价成本价一起调价"
                .ListIndex = 0
                lblMethod.Tag = 0
            End If
        Else
            .AddItem "仅调售价"
            .AddItem "仅调成本价"
            .AddItem "售价成本价一起调价"
            .ListIndex = 0
            lblMethod.Tag = 0
        End If
    End With
End Sub

Private Sub InitTabControl()
    '初始化TabControl控件
    Dim objtabctl As TabControlItem

    picSplit.Left = 0
    picSplit.Top = vsfPrice.Top + vsfPrice.Height + 5
    With TabCtlDetails
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        .InsertItem 0, "库存变动表", vsfStore.hWnd, 0
        .InsertItem 1, "应付款变动表", vsfPay.hWnd, 0
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - vsfPrice.Height - vsfPrice.Top - 20
        .Top = picSplit.Height + picSplit.Top + 20
        .Item(1).Selected = True
        .Item(0).Selected = True
    End With
End Sub

Private Sub Form_Resize()

    On Error Resume Next

    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.WindowState <> vbMaximized Then
        If Me.Height < 8145 Then
            Me.Height = 8145
        End If
    End If

    With fraCondition
        .Width = Me.ScaleWidth
    End With
    txtNO.Move fraCondition.Width - 2000
    lblNO.Move fraCondition.Width - lblNO.Width - 2100
    
    vsfPrice.Move 20, fraCondition.Top + fraCondition.Height + 20, Me.ScaleWidth, 3000
    picSplit.Left = 50
    picSplit.Top = vsfPrice.Top + vsfPrice.Height + 5
    picSplit.Width = Me.ScaleWidth
'    txtSummary.Width = Me.ScaleWidth - lblSummary.Left - lblSummary.Width - 300
    TabCtlDetails.Move 20, picSplit.Height + picSplit.Top, Me.ScaleWidth, Me.ScaleHeight - picSplit.Top - picSplit.Height - picInfo.Height - stbThis.Height
    picInfo.Move 0, TabCtlDetails.Top + TabCtlDetails.Height, Me.ScaleWidth
End Sub

Private Sub picInfo_Resize()
    On Error Resume Next
    
    With txtSummary
        .Width = picInfo.Width - .Left - 300
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ReleaseSelectorRS
    Call SaveWinState(Me, App.ProductName, MStrCaption)
    mblnLoad = False
    mbln应付记录 = False
    mlng供应商ID = 0
    mblnUpdateAdd = False
End Sub

Private Sub mshProvider_DblClick()
    With Me.mshProvider
        Me.txt供应商.Text = .TextMatrix(.Row, 1)
        Me.txt供应商.Tag = .TextMatrix(.Row, 3) & "|" & .TextMatrix(.Row, 1)
        .Visible = False
    End With

    Me.txt供应商.SetFocus
End Sub

Private Sub opt时间_Click(Index As Integer)
    If Index = 0 Then
        dtpRunDate.Enabled = False
    Else
        dtpRunDate.Enabled = True
    End If
End Sub

Private Sub InitVsfGridFlex()
    With vsfPrice

        .Cols = menuPriceCol.总列数
        .rows = 2
        .RowHeight(1) = mlngRowHeight
        .ColWidth(0) = 200
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = 50
        .RowHeight(0) = mlngRowHeight
        .AllowSelection = False '不能多选
'        .SelectionMode = flexSelectionByRow '整行选择
        .ExplorerBar = flexExMoveRows '拖动
        .AllowUserResizing = flexResizeBoth  '可以改变行列宽度
        .Editable = flexEDNone
'        .GridLineWidth = 2
'        .GridLines = flexGridInset
'        .GridColor = &H80000011
'        .GridColorFixed = &H80000011
'        .ForeColorFixed = &H80000012
'        .BackColorSel = &HF4F4EA

        .TextMatrix(0, menuPriceCol.药品id) = "药品ID"
        .TextMatrix(0, menuPriceCol.原价id) = "原价id"
        .TextMatrix(0, menuPriceCol.药价属性) = "药价属性"
        .TextMatrix(0, menuPriceCol.药品) = "药品"
        .TextMatrix(0, menuPriceCol.规格) = "规格"
        .TextMatrix(0, menuPriceCol.是否变价) = "是否变价"
        .TextMatrix(0, menuPriceCol.产地) = "生产商"
        .TextMatrix(0, menuPriceCol.单位) = "单位"
        .TextMatrix(0, menuPriceCol.包装系数) = "包装系数"
        .TextMatrix(0, menuPriceCol.加成率) = "加成率"
        .TextMatrix(0, menuPriceCol.差价让利比) = "差价让利比"
        .TextMatrix(0, menuPriceCol.是否有库存) = "是否有库存"
        .TextMatrix(0, menuPriceCol.收入项目ID) = "收入项目id"
        .TextMatrix(0, menuPriceCol.原成本价) = "原成本价"
        .TextMatrix(0, menuPriceCol.现成本价) = "现成本价"
        .TextMatrix(0, menuPriceCol.原零售价) = "原零售价"
        .TextMatrix(0, menuPriceCol.现零售价) = "现零售价"
        .TextMatrix(0, menuPriceCol.原采购限价) = "原采购限价"
        .TextMatrix(0, menuPriceCol.现采购限价) = "现采购限价"
        .TextMatrix(0, menuPriceCol.原指导售价) = "原指导售价"
        .TextMatrix(0, menuPriceCol.现指导售价) = "现指导售价"

        '设置列宽
        .ColWidth(menuPriceCol.药品id) = 0
        .ColWidth(menuPriceCol.原价id) = 0
        .ColWidth(menuPriceCol.药价属性) = 1000
        .ColWidth(menuPriceCol.药品) = 3000
        .ColWidth(menuPriceCol.规格) = 1500
        .ColWidth(menuPriceCol.是否变价) = 0
        .ColWidth(menuPriceCol.产地) = 2000
        .ColWidth(menuPriceCol.单位) = 800
        .ColWidth(menuPriceCol.包装系数) = 0
        .ColWidth(menuPriceCol.加成率) = 0
        .ColWidth(menuPriceCol.差价让利比) = 0
        .ColWidth(menuPriceCol.是否有库存) = 0
        .ColWidth(menuPriceCol.收入项目ID) = 0
        .ColWidth(menuPriceCol.原成本价) = 1000
        .ColWidth(menuPriceCol.现成本价) = 1000
        .ColWidth(menuPriceCol.原零售价) = 1000
        .ColWidth(menuPriceCol.现零售价) = 1000
        .ColWidth(menuPriceCol.原采购限价) = 0
        .ColWidth(menuPriceCol.现采购限价) = 0
        .ColWidth(menuPriceCol.原指导售价) = 0
        .ColWidth(menuPriceCol.现指导售价) = 0
        '设置对齐方式
        .ColAlignment(menuPriceCol.药价属性) = flexAlignLeftCenter
        .ColAlignment(menuPriceCol.药品) = flexAlignLeftCenter
        .ColAlignment(menuPriceCol.规格) = flexAlignLeftCenter
        .ColAlignment(menuPriceCol.产地) = flexAlignLeftCenter
        .ColAlignment(menuPriceCol.单位) = flexAlignCenterCenter
        .ColAlignment(menuPriceCol.原成本价) = flexAlignRightCenter
        .ColAlignment(menuPriceCol.现成本价) = flexAlignRightCenter
        .ColAlignment(menuPriceCol.原零售价) = flexAlignRightCenter
        .ColAlignment(menuPriceCol.现零售价) = flexAlignRightCenter
        .ColAlignment(menuPriceCol.原采购限价) = flexAlignRightCenter
        .ColAlignment(menuPriceCol.原指导售价) = flexAlignRightCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter '列头居中对齐
        .ColComboList(menuPriceCol.药品) = "|..."
    End With

    With vsfStore
        .Editable = flexEDNone
        .Cols = menuStoreCol.总列数
        .rows = 1
        .ColWidth(0) = 200
'        .RowHeight(1) = mlngRowHeight
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = 50
        .RowHeight(0) = mlngRowHeight
        .AllowSelection = False '不能多选
'        .SelectionMode = flexSelectionByRow '整行选择
        .ExplorerBar = flexExMoveRows '拖动
        .AllowUserResizing = flexResizeBoth  '可以改变行列宽度
        .GridLineWidth = 2
        .GridLines = flexGridInset
        .GridColor = &H0&

        '设置列名
        .TextMatrix(0, menuStoreCol.药品id) = "药品id"
        .TextMatrix(0, menuStoreCol.库房) = "库房"
        .TextMatrix(0, menuStoreCol.库房id) = "库房id"
        .TextMatrix(0, menuStoreCol.供应商) = "供应商"
        .TextMatrix(0, menuStoreCol.供应商id) = "供应商id"
        .TextMatrix(0, menuStoreCol.药品) = "药品"
        .TextMatrix(0, menuStoreCol.规格) = "规格"
        .TextMatrix(0, menuStoreCol.单位) = "单位"
        .TextMatrix(0, menuStoreCol.批号) = "批号"
        .TextMatrix(0, menuStoreCol.效期) = "效期"
        .TextMatrix(0, menuStoreCol.产地) = "生产商"
        .TextMatrix(0, menuStoreCol.数量) = "数量"
        .TextMatrix(0, menuStoreCol.包装系数) = "包装系数"
        .TextMatrix(0, menuStoreCol.批次) = "批次"
        .TextMatrix(0, menuStoreCol.变价) = "变价"
        .TextMatrix(0, menuStoreCol.原零售价) = "原零售价"
        .TextMatrix(0, menuStoreCol.现零售价) = "现零售价"
        .TextMatrix(0, menuStoreCol.售价盈亏) = "售价盈亏"
        .TextMatrix(0, menuStoreCol.加成率) = "加成率"
        .TextMatrix(0, menuStoreCol.原成本价) = "原成本价"
        .TextMatrix(0, menuStoreCol.现成本价) = "现成本价"
        .TextMatrix(0, menuStoreCol.成本盈亏) = "成本盈亏"
        '设置列宽
        .ColWidth(0) = 0
        .ColWidth(menuStoreCol.库房) = 1500
        .ColWidth(menuStoreCol.库房id) = 0
        .ColWidth(menuStoreCol.供应商) = 2000
        .ColWidth(menuStoreCol.供应商id) = 0
        .ColWidth(menuStoreCol.药品) = 3000
        .ColWidth(menuStoreCol.规格) = 1500
        .ColWidth(menuStoreCol.单位) = 800
        .ColWidth(menuStoreCol.批号) = 1500
        .ColWidth(menuStoreCol.效期) = 2000
        .ColWidth(menuStoreCol.产地) = 1500
        .ColWidth(menuStoreCol.数量) = 1500
        .ColWidth(menuStoreCol.包装系数) = 0
        .ColWidth(menuStoreCol.批次) = 0
        .ColWidth(menuStoreCol.变价) = 0
        .ColWidth(menuStoreCol.原零售价) = 1000
        .ColWidth(menuStoreCol.现零售价) = 1000
        .ColWidth(menuStoreCol.售价盈亏) = 1000
        .ColWidth(menuStoreCol.加成率) = 1000
        .ColWidth(menuStoreCol.原成本价) = 1000
        .ColWidth(menuStoreCol.现成本价) = 1000
        .ColWidth(menuStoreCol.成本盈亏) = 1000
        '对齐方式
        .ColAlignment(menuStoreCol.库房) = flexAlignLeftCenter
        .ColAlignment(menuStoreCol.供应商) = flexAlignLeftCenter
        .ColAlignment(menuStoreCol.药品) = flexAlignLeftCenter
        .ColAlignment(menuStoreCol.规格) = flexAlignLeftCenter
        .ColAlignment(menuStoreCol.单位) = flexAlignCenterCenter
        .ColAlignment(menuStoreCol.批号) = flexAlignLeftCenter
        .ColAlignment(menuStoreCol.效期) = flexAlignLeftCenter
        .ColAlignment(menuStoreCol.产地) = flexAlignLeftCenter
        .ColAlignment(menuStoreCol.数量) = flexAlignRightCenter
        .ColAlignment(menuStoreCol.原零售价) = flexAlignRightCenter
        .ColAlignment(menuStoreCol.现零售价) = flexAlignRightCenter
        .ColAlignment(menuStoreCol.售价盈亏) = flexAlignRightCenter
        .ColAlignment(menuStoreCol.加成率) = flexAlignRightCenter
        .ColAlignment(menuStoreCol.原成本价) = flexAlignRightCenter
        .ColAlignment(menuStoreCol.现成本价) = flexAlignRightCenter
        .ColAlignment(menuStoreCol.成本盈亏) = flexAlignRightCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter '列头居中对齐
    End With

    With vsfPay
        .Editable = flexEDNone
        .Cols = menuPayCol.总列数
        .rows = 1
        .ColWidth(0) = 200
'        .RowHeight(1) = mlngRowHeight
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = 50
        .RowHeight(0) = mlngRowHeight
        .AllowSelection = False '不能多选
'        .SelectionMode = flexSelectionByRow '整行选择
        .ExplorerBar = flexExMoveRows '拖动
        .AllowUserResizing = flexResizeBoth  '可以改变行列宽度
        .GridLineWidth = 2
        .GridLines = flexGridInset
        .GridColor = &H0&

        .TextMatrix(0, menuPayCol.药品id) = "药品id"
        .TextMatrix(0, menuPayCol.药品) = "药品"
        .TextMatrix(0, menuPayCol.发票号) = "发票号"
        .TextMatrix(0, menuPayCol.发票日期) = "发票日期"
        .TextMatrix(0, menuPayCol.发票金额) = "发票金额"
        '设置列宽
        .ColWidth(menuPayCol.药品id) = 0
        .ColWidth(menuPayCol.药品) = 2000
        .ColWidth(menuPayCol.发票号) = 1500
        .ColWidth(menuPayCol.发票日期) = 2000
        .ColWidth(menuPayCol.发票金额) = 1500
        '对齐方式
        .ColAlignment(menuPayCol.药品) = flexAlignLeftCenter
        .ColAlignment(menuPayCol.发票号) = flexAlignLeftCenter
        .ColAlignment(menuPayCol.发票日期) = flexAlignLeftCenter
        .ColAlignment(menuPayCol.发票金额) = flexAlignRightCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter '列头居中对齐
    End With
End Sub

Private Sub initGrid()
    '如果是修改或者查阅则提取相应的记录并填充到表格中
    Dim rsTemp As ADODB.Recordset
    Dim lngRow As Long
    Dim i As Long
    Dim lngDrugID As Long
    Dim db包装系数 As Double
    Dim strUnit As String
    Dim StrToday As String
    Dim rs产地 As ADODB.Recordset

    On Error GoTo errHandle
    '调价方式 0-调售价;1-调成本价;2-调售价及成本价
    If mintMethod = 0 Then
        gstrSQL = "Select Distinct p.原价id, i.是否变价, Nvl(s.指导批发价, 0) As 指导批价, Nvl(s.扣率, 0) As 扣率, Nvl(s.指导零售价, 0) As 指导售价," & vbNewLine & _
            "                s.加成率/100 As 加成率, i.编码, b.名称 As 商品名, i.名称 As 通用名, i.规格, i.产地 As 产地, i.计算单位 As 单位," & vbNewLine & _
            "                s.门诊单位, s.门诊包装, s.住院单位, s.住院包装, s.药库单位, Nvl(s.药库包装, 1) 药库包装, s.成本价 As 原成本价, s.成本价 As 新成本价, p.原价, p.现价," & vbNewLine & _
            "                p.收入项目id, p.调价人, p.调价说明, s.差价让利比, To_Char(a.执行日期, 'YYYY-MM-DD HH24:MI:SS') As 执行日期, i.Id 药品id," & vbNewLine & _
            "                Decode(k.药品id, Null, 0, 1) 是否有库存" & vbNewLine & _
            "From (Select s.药品id From 药品库存 s where s.性质=1 And Not (zl_fun_getbatchpro(s.库房id,s.药品id)=1 And Nvl(S.批次,0) = 0 And S.可用数量 < 0 And S.实际数量 = 0 And S.实际金额 = 0 And S.实际差价 = 0)) K, 调价汇总记录 A, 收费项目别名 B, 药品规格 S, 收费项目目录 I, 收费价目 P" & vbNewLine & _
            "Where a.调价号 = p.调价汇总号 And b.收费细目id(+) = s.药品id And s.药品id = i.Id And i.Id = k.药品id(+) And i.Id = p.收费细目id And" & vbNewLine & _
            "      p.调价汇总号 = [1] And a.分类 = 0 And b.性质(+) = 3 And a.调价号 = [1] " & vbNewLine & _
            GetPriceClassString("P") & vbNewLine & _
            IIf(mintModal = 2, "", "  And (i.撤档时间 Is Null Or i.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))") & vbNewLine & _
            "Order By 药品id"
    ElseIf mintMethod = 1 Then
        gstrSQL = "Select Distinct i.是否变价, Nvl(s.指导批发价, 0) As 指导批价, Nvl(s.扣率, 0) As 扣率, Nvl(s.指导零售价, 0) As 指导售价," & vbNewLine & _
            "                s.加成率/100 As 加成率, i.编码, b.名称 As 商品名, i.名称 As 通用名, i.规格, i.产地 As 产地, i.计算单位 As 单位," & vbNewLine & _
            "                s.门诊单位, s.门诊包装, s.住院单位, s.住院包装, s.药库单位, Nvl(s.药库包装, 1) 药库包装, m.原成本价, m.新成本价, p.原价, p.现价, p.收入项目id," & vbNewLine & _
            "                a.填制人 As 调价人, a.说明 As 调价说明, s.差价让利比, To_Char(m.执行日期, 'YYYY-MM-DD HH24:MI:SS') As 执行日期, i.Id 药品id," & vbNewLine & _
            "                Decode(k.药品id, Null, 0, 1) 是否有库存" & vbNewLine & _
            "From (Select Min(原价) As 原成本价, Min(现价) As 新成本价, min(产地) as 产地,调价汇总号,药品id,min(执行日期) as 执行日期 From 药品价格记录 Where 价格类型=2 and 调价汇总号 = [1] Group By 调价汇总号,药品id) M, (Select s.药品id From 药品库存 s where s.性质=1 And Not (zl_fun_getbatchpro(s.库房id,s.药品id)=1 And Nvl(S.批次,0) = 0 And S.可用数量 < 0 And S.实际数量 = 0 And S.实际金额 = 0 And S.实际差价 = 0)) K, 调价汇总记录 A, 收费项目别名 B, 药品规格 S, 收费项目目录 I, 收费价目 P" & vbNewLine & _
            "Where m.调价汇总号(+) = a.调价号 And b.收费细目id(+) = s.药品id And s.药品id = i.Id And i.Id = k.药品id(+) And m.药品id = i.Id And" & vbNewLine & _
            "      i.Id = p.收费细目id And Sysdate Between p.执行日期 And p.终止日期 And m.调价汇总号 = [1] And a.分类 = 0 And b.性质(+) = 3 And" & vbNewLine & _
            "      a.调价号 = [1] " & IIf(mintModal = 2, "", " And (i.撤档时间 Is Null Or i.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))") & vbNewLine & _
            GetPriceClassString("P") & vbNewLine & _
            "Order By 药品id"
    ElseIf mintMethod = 2 Then
        gstrSQL = "Select distinct p.原价id, i.是否变价, Nvl(s.指导批发价, 0) As 指导批价, Nvl(s.扣率, 0) As 扣率, Nvl(s.指导零售价, 0) As 指导售价," & vbNewLine & _
            "       s.加成率/100 As 加成率, i.编码, b.名称 As 商品名, i.名称 As 通用名, i.规格, i.产地 As 产地, i.计算单位 As 单位, s.门诊单位," & vbNewLine & _
            "       s.门诊包装, s.住院单位, s.住院包装, s.药库单位, Nvl(s.药库包装, 1) 药库包装, m.原成本价, m.新成本价, p.原价, p.现价, p.收入项目id, p.调价人, p.调价说明, s.差价让利比," & vbNewLine & _
            "       To_Char(p.执行日期, 'YYYY-MM-DD HH24:MI:SS') As 执行日期, i.Id 药品id, Decode(k.药品id, Null, 0, 1) 是否有库存" & vbNewLine & _
            "From (Select 药品id,Min(原价) As 原成本价, Min(现价) As 新成本价, min(产地) as 产地,调价汇总号 From 药品价格记录 Where 价格类型=2 and 调价汇总号 = [1] Group By 药品id,调价汇总号) M, 收费价目 P, 调价汇总记录 A, (Select s.药品id From 药品库存 s where s.性质=1 And Not (zl_fun_getbatchpro(s.库房id,s.药品id)=1 And Nvl(S.批次,0) = 0 And S.可用数量 < 0 And S.实际数量 = 0 And S.实际金额 = 0 And S.实际差价 = 0)) K, 收费项目别名 B, 药品规格 S, 收费项目目录 I" & vbNewLine & _
            "Where m.调价汇总号 = a.调价号 and m.药品id=i.id And p.调价汇总号 = a.调价号 And p.收费细目id = k.药品id(+) And p.收费细目id = b.收费细目id(+) And p.收费细目id = s.药品id And" & vbNewLine & _
            "      s.药品id = i.Id And a.调价号 =[1] And b.性质(+) = 3 " & vbNewLine & _
            GetPriceClassString("P") & vbNewLine & _
            IIf(mintModal = 2, "", "  And (i.撤档时间 Is Null Or i.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))") & "Order By 药品id "
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption, mstr调价汇总号)
    If rsTemp.RecordCount = 0 Then
        MsgBox "该调价记录已经被删除了！", vbInformation, gstrSysName
        Exit Sub
    End If

    With vsfPrice
        .rows = 2
        rsTemp.MoveFirst
        For i = 0 To rsTemp.RecordCount - 1
            If rsTemp!药品id <> lngDrugID Then
                Select Case mintUnit
                    Case 0
                        db包装系数 = rsTemp!药库包装
                        strUnit = rsTemp!药库单位
                    Case 2
                        db包装系数 = rsTemp!住院包装
                        strUnit = rsTemp!住院单位
                    Case 1
                        db包装系数 = rsTemp!门诊包装
                        strUnit = rsTemp!门诊单位
                    Case 3
                        db包装系数 = 1
                        strUnit = rsTemp!单位
                End Select

                lngDrugID = rsTemp!药品id
                If mintMethod = 0 Or mintMethod = 2 Then
                    .TextMatrix(.rows - 1, menuPriceCol.原价id) = IIf(IsNull(rsTemp!原价id), "", rsTemp!原价id)
                End If
                .TextMatrix(.rows - 1, menuPriceCol.药品id) = rsTemp!药品id

                If gint药品名称显示 = 1 Then
                    .TextMatrix(.rows - 1, menuPriceCol.药品) = "[" & rsTemp!编码 & "]" & IIf(IsNull(rsTemp!商品名), rsTemp!通用名, rsTemp!商品名)
                Else
                    .TextMatrix(.rows - 1, menuPriceCol.药品) = "[" & rsTemp!编码 & "]" & rsTemp!通用名
                End If
                .TextMatrix(.rows - 1, menuPriceCol.规格) = IIf(IsNull(rsTemp!规格), "", rsTemp!规格)
                .TextMatrix(.rows - 1, menuPriceCol.是否变价) = rsTemp!是否变价
                .TextMatrix(.rows - 1, menuPriceCol.药价属性) = IIf(rsTemp!是否变价 = 0, "定价", "时价")
                
'                If mintMethod = 1 Or mintMethod = 2 Then
'                    gstrSQL = "select min(产地) as 产地 from 成本价调价信息 where 调价汇总号=[1] and 药品id=[2]"
'                    Set rs产地 = zldatabase.OpenSQLRecord(gstrSQL, "产地查询", mstr调价汇总号, rsTemp!药品id)
'                    If rs产地.RecordCount > 0 Then
'                        .TextMatrix(.rows - 1, menuPriceCol.产地) = IIf(IsNull(rs产地!厂牌), "", rs产地!厂牌)
'                    End If
'                Else
                    .TextMatrix(.rows - 1, menuPriceCol.产地) = IIf(IsNull(rsTemp!产地), "", rsTemp!产地)
'                End If
                
                .TextMatrix(.rows - 1, menuPriceCol.单位) = strUnit
                .TextMatrix(.rows - 1, menuPriceCol.包装系数) = db包装系数

                .TextMatrix(.rows - 1, menuPriceCol.加成率) = rsTemp!加成率
                .TextMatrix(.rows - 1, menuPriceCol.差价让利比) = Nvl(rsTemp!差价让利比, 100)
                .TextMatrix(.rows - 1, menuPriceCol.是否有库存) = rsTemp!是否有库存
                .TextMatrix(.rows - 1, menuPriceCol.收入项目ID) = IIf(IsNull(rsTemp!收入项目ID), "", rsTemp!收入项目ID)
                .TextMatrix(.rows - 1, menuPriceCol.原成本价) = zlStr.FormatEx(Nvl(rsTemp!原成本价, 0) * db包装系数, mintCostDigit, , True)
                .TextMatrix(.rows - 1, menuPriceCol.现成本价) = zlStr.FormatEx(rsTemp!新成本价 * db包装系数, mintCostDigit, , True)
                .TextMatrix(.rows - 1, menuPriceCol.原零售价) = zlStr.FormatEx(IIf(IsNull(rsTemp!原价), rsTemp!现价, rsTemp!原价) * db包装系数, mintPriceDigit, , True)
                .TextMatrix(.rows - 1, menuPriceCol.现零售价) = zlStr.FormatEx(rsTemp!现价 * db包装系数, mintPriceDigit, , True)
                .TextMatrix(.rows - 1, menuPriceCol.原采购限价) = zlStr.FormatEx(rsTemp!指导批价 * db包装系数, mintCostDigit, , True)
                .TextMatrix(.rows - 1, menuPriceCol.现采购限价) = zlStr.FormatEx(rsTemp!指导批价 * db包装系数, mintCostDigit, , True)
                .TextMatrix(.rows - 1, menuPriceCol.原指导售价) = zlStr.FormatEx(rsTemp!指导售价 * db包装系数, mintPriceDigit, , True)
                .TextMatrix(.rows - 1, menuPriceCol.现指导售价) = zlStr.FormatEx(rsTemp!指导售价 * db包装系数, mintPriceDigit, , True)

                txtValuer.Text = IIf(IsNull(rsTemp!调价人), "", rsTemp!调价人)
                txtSummary.Text = IIf(IsNull(rsTemp!调价说明), "", rsTemp!调价说明)
                If mintModal = 1 Then
                    Me.dtpRunDate.MinDate = CDate(rsTemp!执行日期)
                End If
                If IsNull(rsTemp!执行日期) Then
                    StrToday = Format(Sys.Currentdate(), "yyyy-MM-dd hh:mm:ss")
                Else
                    StrToday = Format(rsTemp!执行日期, "yyyy-MM-dd hh:mm:ss")
                End If
                Me.dtpRunDate.Value = CDate(StrToday)

                .rows = .rows + 1
                Call setColEdit
                .RowHeight(.rows - 1) = mlngRowHeight
            End If
            rsTemp.MoveNext
        Next
        
        .colHidden(menuPriceCol.原零售价) = False
        .colHidden(menuPriceCol.原成本价) = False
        If mintMethod = 1 Then
            '调成本价
            .colHidden(menuPriceCol.原零售价) = True
        ElseIf mintMethod = 0 Then
            '调售价
            .colHidden(menuPriceCol.原成本价) = True
        End If
        
        Call GetDrugStore(Val(.TextMatrix(1, menuPriceCol.药品id)), 1)
    End With

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FindGridRow(ByVal strInput As String)
    Dim n As Integer
    Dim lngFindRow As Long
    Dim str药名 As String
    Dim lngRow As Long

    '查找药品
    On Error GoTo errHandle
    If strInput <> txtFind.Tag Then
        '表示新的查找
        txtFind.Tag = strInput

        gstrSQL = "Select Distinct A.Id,'[' || A.编码 || ']' As 药品编码, A.名称 As 通用名, B.名称 As 商品名 " & _
                  "From 收费项目目录 A,收费项目别名 B " & _
                  "Where (A.站点 = [3] Or A.站点 is Null) And A.Id =B.收费细目id And A.类别 In ('5','6','7') " & _
                  "  And (A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2] ) " & _
                  "Order By 药品编码 "
        Set mrsFindName = zlDatabase.OpenSQLRecord(gstrSQL, "取匹配的药品ID", strInput & "%", "%" & strInput & "%", gstrNodeNo)

        If mrsFindName.RecordCount = 0 Then Exit Sub
        mrsFindName.MoveFirst
    End If

    '开始查找
    If mrsFindName.State <> adStateOpen Then Exit Sub
    If mrsFindName.RecordCount = 0 Then Exit Sub

    For n = 1 To mrsFindName.RecordCount
        '如果到底了，则返回第1条记录
        If mrsFindName.EOF Then mrsFindName.MoveFirst

        If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
            str药名 = mrsFindName!药品编码 & mrsFindName!通用名
        Else
            str药名 = mrsFindName!药品编码 & IIf(IsNull(mrsFindName!商品名), mrsFindName!通用名, mrsFindName!商品名)
        End If

        For lngRow = 1 To vsfPrice.rows - 1
            lngFindRow = vsfPrice.FindRow(str药名, lngRow, CLng(menuPriceCol.药品), True, True)
            If lngFindRow > 0 Then
'                vsfPrice.Select lngFindRow, 1, lngFindRow, vsfPrice.Cols - 1
                vsfPrice.Row = lngFindRow
                vsfPrice.TopRow = lngFindRow
                Exit For
            End If
        Next

        If lngFindRow > 0 Then  '查询到数据后就移动下下一条并退出本次查询
            mrsFindName.MoveNext
            Exit For
        Else
            mrsFindName.MoveNext '未查询到数据则移动到下一条数据集继续查询
        End If
    Next
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    If vsfPrice.Height + y <= 800 Then Exit Sub
    If TabCtlDetails.Height - y <= 1000 Then Exit Sub
    picSplit.Move 0, picSplit.Top + y
    vsfPrice.Move 0, fraCondition.Top + fraCondition.Height + 20, Me.ScaleWidth, vsfPrice.Height + y

    With TabCtlDetails
        .Top = picSplit.Top + picSplit.Height
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = TabCtlDetails.Height - y
    End With
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Trim(txtFind.Text) = "" Then Exit Sub

    Call FindGridRow(UCase(Trim(txtFind.Text)))
End Sub

Private Sub txtSummary_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If LenB(StrConv(txtSummary.Text, vbFromUnicode)) >= 100 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtSummary_Validate(Cancel As Boolean)
    If LenB(StrConv(txtSummary.Text, vbFromUnicode)) > 100 Then
        MsgBox "说明太长！", vbInformation, gstrSysName
        txtSummary.SelStart = 0
        txtSummary.SelLength = LenB(StrConv(txtSummary.Text, vbFromUnicode))
        Cancel = True
    End If
End Sub

Private Sub txt供应商_GotFocus()
    Me.txt供应商.SelStart = 0: Me.txt供应商.SelLength = Len(Me.txt供应商.Text)
End Sub

Private Sub txt供应商_KeyPress(KeyAscii As Integer)
    Dim strTmp As String
    Dim rsTemp As ADODB.Recordset

    On Error GoTo errHandle
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub

    strTmp = UCase(Trim(Me.txt供应商.Text))

    If strTmp = "" Then
        Me.txt供应商.Tag = "|"
        Exit Sub
    ElseIf strTmp = Split(Me.txt供应商.Tag, "|")(1) Then
        Exit Sub
    End If

    gstrSQL = "Select 编码,名称,简码,id" & _
            " From 供应商" & _
            " where (编码 Like [1] " & _
            "       Or 名称 Like [2] " & _
            "       Or 简码 Like [2])" & _
            " And 末级=1 And substr(类型,1,1) = '1' And (撤档时间 is null or 撤档时间=to_date('3000-01-01','YYYY-MM-DD')) " & _
            " Order By 编码 "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption, strTmp & "%", IIf(gstrMatchMethod = "0", "%", "") & strTmp & "%")

    With rsTemp
        If .EOF Then
            MsgBox "没有找到匹配的供应商，请在供应商管理中增加供应商！", vbInformation, gstrSysName
            Me.txt供应商.Text = Split(Me.txt供应商.Tag, "|")(1)
            Me.txt供应商.SelStart = 0: Me.txt供应商.SelLength = Len(Me.txt供应商.Text)
            Exit Sub
        End If

        If .RecordCount = 1 Then
            Me.txt供应商.Text = Trim(rsTemp!名称): Me.txt供应商.Tag = rsTemp!Id & "|" & rsTemp!名称
            Exit Sub
        Else
            With Me.mshProvider
                .Left = Me.chk供应商.Left
                .Top = Me.txt供应商.Top + Me.txt供应商.Height
                .Clear
                Set .DataSource = rsTemp
                .ColWidth(0) = 800: .ColWidth(1) = 2500: .ColWidth(2) = 800: .ColWidth(3) = 0
                .Row = 1: .ColSel = .Cols - 1
                .ZOrder 0: .Visible = True: .SetFocus
            End With
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub get分段加成售价(ByVal lng药品id As Long, ByVal lng比例系数 As Long, ByVal dbl采购价 As Double, ByRef dbl售价 As Double)
'功能：通过成本价按分段加成方式计算售价
'参数：成本价,售价
    Dim dbl差价额 As Double
    Dim blnData As Boolean
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    mdbl分段加成率 = 0
    dbl差价额 = 0
    
    gstrSQL = "select 类别 from  收费项目目录 a where a.id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取得药品材质分类", lng药品id)
    If rsTemp!类别 = 7 Then
        mrs分段加成.Filter = "类型=1"
    Else
        mrs分段加成.Filter = "类型=0"
    End If
    
    If mrs分段加成.RecordCount <> 0 Then
        mrs分段加成.MoveFirst
        Do While Not mrs分段加成.EOF
            With mrs分段加成
                If dbl采购价 > !最低价 And dbl采购价 <= !最高价 Then
                    mdbl分段加成率 = IIf(IsNull(!加成率), 0, !加成率) / 100
                    dbl差价额 = IIf(IsNull(!差价额), 0, !差价额)
                    blnData = True
                    Exit Do
                End If
            End With
            mrs分段加成.MoveNext
        Loop
    End If
    
    If blnData = False Then
        MsgBox "没有设置金额段为：" & dbl采购价 & "  的分段加成数据，请在药品目录管理（分段加成率）中设置！", vbInformation, gstrSysName
        dbl售价 = 0
        Exit Sub
    End If
    
    dbl售价 = dbl采购价 * (1 + mdbl分段加成率) + dbl差价额
    
    Set rsTemp = Nothing
    gstrSQL = "Select 指导零售价 From 药品规格 Where 药品ID=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption & "[读取指导零售价]", lng药品id)
    If rsTemp!指导零售价 * lng比例系数 < dbl售价 Then
        dbl售价 = rsTemp!指导零售价 * lng比例系数
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub txt供应商_Validate(Cancel As Boolean)
    If Me.txt供应商.Text = "" Then
        Me.txt供应商.Tag = "|"
    ElseIf Me.txt供应商.Text <> Split(Me.txt供应商.Tag, "|")(1) Then
        txt供应商_KeyPress (vbKeyReturn)
    End If
End Sub


Private Sub vsfPay_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsfPay
        .Move 0, 360, TabCtlDetails.Width, TabCtlDetails.Height - 370
    End With
End Sub

Private Sub vsfPay_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfPay
        If .Cell(flexcpBackColor, Row, Col, Row, Col) = mconlngColor Then
            Cancel = True
            .Editable = flexEDNone
        Else
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub


Private Sub vsfPay_DblClick()
    With vsfPay
        If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mconlngCanColColor Then
            .EditCell
            .EditSelStart = 0
            .EditSelLength = Len(.EditText)
        End If
    End With
End Sub

Private Sub vsfPay_EnterCell()
    With vsfPay
        If .CellBackColor = mconlngColor Then
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
        End If
    End With
End Sub

Private Sub vsfPay_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsfPay
        If KeyCode = vbKeyReturn Then
            If .Col = menuPayCol.药品 Then
                .Col = menuPayCol.发票号
            ElseIf .Col = menuPayCol.发票号 Then
                .Col = menuPayCol.发票日期
            ElseIf .Col = menuPayCol.发票日期 Then
                .Col = menuPayCol.发票金额
            ElseIf .Col = menuPayCol.发票金额 And .Row <> .rows - 1 Then
                .Col = menuPayCol.药品
                .Row = .Row + 1
            End If
        End If
    End With
End Sub

Private Sub vsfPay_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        With vsfPay
            If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mconlngCanColColor Then
                .Editable = flexEDKbdMouse
            Else
                .Editable = flexEDNone
            End If
        End With
    End If
End Sub

Private Sub vsfPay_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strkey As String
    Dim intDigit As Integer
    
    If KeyAscii = vbKeyReturn Then Exit Sub
    If KeyAscii <> vbKeyBack Then
        With vsfPay
            If Col = menuPayCol.发票金额 Then
                strkey = .EditText
                intDigit = mintMoneyDigit
                If KeyAscii = vbKeyDelete Then
                    If InStr(1, .EditText, ".") > 0 Then
                        KeyAscii = 0
                    End If
                ElseIf KeyAscii = Asc(".") Or (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
                    If .EditSelLength = Len(strkey) Then Exit Sub
                    If InStr(strkey, ".") <> 0 And Chr(KeyAscii) = "." Then   '只能存在一个小数点
                        KeyAscii = 0
                        Exit Sub
                    End If
                    If Len(Mid(strkey, InStr(1, strkey, ".") + 1)) >= intDigit And strkey Like "*.*" Then
                        KeyAscii = 0
                        Exit Sub
                    Else
                        Exit Sub
                    End If
                Else
                    KeyAscii = 0
                End If
            ElseIf Col = menuPayCol.发票号 Then
                If InStr("`~!@#$%^&*()_-+={[}]|\:;""'<,>.?/", Chr(KeyAscii)) > 0 Then
                    KeyAscii = 0
                End If
            End If
        End With
    End If
End Sub

Private Sub vsfPay_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strkey As String

    With vsfPay
        If Col = menuPayCol.发票日期 Then
            strkey = .EditText
            If strkey <> "" Then
                If Len(strkey) = 8 And InStr(1, strkey, "-") = 0 Then
                    strkey = TranNumToDate(strkey)
                    If strkey = "" Then
                        MsgBox "对不起，发票日期必须为日期型,格式(20000101或者2000-01-01)！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                    .EditText = strkey
                    .TextMatrix(Row, menuPayCol.发票日期) = .EditText
                End If
                
                If Not IsDate(strkey) Then
                    MsgBox "对不起，发票日期必须为日期型(20000101或者2000-01-01)！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    Exit Sub
                End If
            End If
        End If
    End With
End Sub

Private Sub vsfprice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsfPrice
        If Col = menuPriceCol.现成本价 Then
            If Val(.TextMatrix(Row, Col)) <> Val(.TextMatrix(Row, menuPriceCol.原成本价)) Then
                .Cell(flexcpFontBold, Row, Col, Row, Col) = 10
                .Cell(flexcpForeColor, Row, Col, Row, Col) = vbRed
            End If
        ElseIf Col = menuPriceCol.现零售价 Then
            If Val(.TextMatrix(Row, Col)) <> Val(.TextMatrix(Row, menuPriceCol.原零售价)) Then
                .Cell(flexcpFontBold, Row, Col, Row, Col) = 10
                .Cell(flexcpForeColor, Row, Col, Row, Col) = vbRed
            End If
        End If
    End With
End Sub

Private Sub vsfPrice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Then Exit Sub
'    Call SetRowHidden(Val(vsfPrice.TextMatrix(NewRow, menuPriceCol.药品id)))
End Sub

Private Sub SetRowHidden(ByVal lngDrugID As Long)
    '功能：行的显示与隐藏
    '参数：药品id
    Dim intRow As Integer

    If lngDrugID = 0 Then Exit Sub
    With vsfStore
        For intRow = 1 To .rows - 1
            If Val(.TextMatrix(intRow, menuStoreCol.药品id)) = lngDrugID Then
                .RowHidden(intRow) = False
            Else
                .RowHidden(intRow) = True
            End If
        Next
    End With

    With vsfPay
        For intRow = 1 To .rows - 1
            If Val(.TextMatrix(intRow, menuPayCol.药品id)) = lngDrugID Then
                .RowHidden(intRow) = False
            Else
                .RowHidden(intRow) = True
            End If
        Next
    End With
End Sub

'Private Sub vsfPrice_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'    With vsfPrice
'        If .Cell(flexcpBackColor, Row, Col, Row, Col) = mconlngColor Then
'            Cancel = True
'            .Editable = flexEDNone
'        Else
'            .Editable = flexEDKbdMouse
'        End If
'    End With
'End Sub

Private Sub vsfPrice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim mrsReturn As Recordset
    Dim vRect As RECT
    Dim dblLeft As Double
    Dim dblTop As Double

    mBlnClick = True
    vRect = zlControl.GetControlRect(vsfPrice.hWnd) '获取位置
    dblLeft = vsfPrice.CellLeft
    dblTop = vRect.Top + vsfPrice.CellTop + vsfPrice.CellHeight


    On Error GoTo errHandle
    If grsMaster.State = adStateClosed Then
        Call SetSelectorRS(1, "", 0, , , , , , , , , True)
    End If
    Set mrsReturn = frmSelector.ShowME(Me, 0, 1, , dblLeft, dblTop, , , , , , , , , False, mstrPrivs)

    If mrsReturn.RecordCount = 0 Then Exit Sub
    mblnUpdateAdd = True
    Call GetDrugPirce(mrsReturn, Row)
    mblnUpdateAdd = False
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetDrugPirce(ByVal rsReturn As ADODB.Recordset, ByVal Row As Integer)
    '用来获取药品信息
    Dim rsTemp As Recordset
    Dim lngDrugID As Long
    Dim intRow As Long
    Dim i As Long
    Dim intCurrentPrice As Integer '是否是时价
    Dim strUnit As String
    Dim db包装系数 As Double
    Dim strInfo As String

    On Error GoTo errHandle

    mlngOldDrugID = Val(vsfPrice.TextMatrix(Row, menuPriceCol.药品id))
    Set rsReturn = CheckDoubleDrug(rsReturn)
    If rsReturn.RecordCount = 0 Then Exit Sub

    rsReturn.MoveFirst
    For i = 0 To rsReturn.RecordCount - 1
        With vsfPrice
            lngDrugID = rsReturn!药品id

            '检查是否存在为执行的价格
            If checkNotExecutePrice(lngDrugID, strInfo) = True Then
                MsgBox strInfo, vbInformation, gstrSysName
                Exit Sub
            End If

            Select Case mintUnit
                Case 0
                    db包装系数 = rsReturn!药库包装
                    strUnit = rsReturn!药库单位
                Case 2
                    db包装系数 = rsReturn!住院包装
                    strUnit = rsReturn!住院单位
                Case 1
                    db包装系数 = rsReturn!门诊包装
                    strUnit = rsReturn!门诊单位
                Case 3
                    db包装系数 = 1
                    strUnit = rsReturn!售价单位
            End Select

            .TextMatrix(Row, menuPriceCol.药品id) = lngDrugID

            If gint药品名称显示 = 1 Then
                .TextMatrix(Row, menuPriceCol.药品) = "[" & rsReturn!药品编码 & "]" & IIf(IsNull(rsReturn!商品名), rsReturn!通用名, rsReturn!商品名)
            Else
                .TextMatrix(Row, menuPriceCol.药品) = "[" & rsReturn!药品编码 & "]" & rsReturn!通用名
            End If

            .TextMatrix(Row, menuPriceCol.规格) = IIf(IsNull(rsReturn!规格), "", rsReturn!规格)
            .TextMatrix(Row, menuPriceCol.是否变价) = rsReturn!时价
            .TextMatrix(Row, menuPriceCol.药价属性) = IIf(rsReturn!时价 = 0, "定价", "时价")
            intCurrentPrice = rsReturn!时价
            .TextMatrix(Row, menuPriceCol.产地) = IIf(IsNull(rsReturn!产地), "", rsReturn!产地)
            .TextMatrix(Row, menuPriceCol.单位) = strUnit
            .TextMatrix(Row, menuPriceCol.包装系数) = db包装系数
            gstrSQL = "select 药品id from 药品库存 s where s.药品id=[1] and s.性质=1 And Not (zl_fun_getbatchpro(s.库房id,[1])=1 And Nvl(S.批次,0) = 0 And S.可用数量 < 0 And S.实际数量 = 0 And S.实际金额 = 0 And S.实际差价 = 0) "
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查库存", lngDrugID)
            If rsTemp.RecordCount = 0 Then
                .TextMatrix(Row, menuPriceCol.是否有库存) = 0
            Else
                .TextMatrix(Row, menuPriceCol.是否有库存) = 1
            End If

            If intCurrentPrice = 0 Then '定价药品
                '表示定价药品调价，成本价取平均价格，售价取收费价目现价
                gstrSQL = "Select b.Id, Decode(k.成本价, Null, a.成本价*" & db包装系数 & ", k.成本价) As 成本价, a.指导批发价, a.指导零售价, b.现价*" & db包装系数 & " as 现价, a.差价让利比, a.加成率 / 100 As 加成率," & vbNewLine & _
                    "            b.收入项目id" & vbNewLine & _
                    "     From 药品规格 A, 收费价目 B," & vbNewLine & _
                    "          (Select Decode(Sum(Nvl(实际数量, 0)), 0, Null, Sum(Round(平均成本价*" & db包装系数 & ", " & mintCostDigit & ") * round(实际数量/" & db包装系数 & "," & mintNumberDigit & ")) / Sum(round(实际数量/" & db包装系数 & "," & mintNumberDigit & "))) As 成本价" & vbNewLine & _
                    "            From 药品库存" & vbNewLine & _
                    "            Where 性质 = 1 And 药品id = [1] ) K" & vbNewLine & _
                    "     Where a.药品id = b.收费细目id And a.药品id = [1] And Sysdate Between 执行日期 And 终止日期" & GetPriceClassString("B")
            Else '时价药品
                '表示时价药品调价，取库存金额/库存数量做为其价格
                gstrSQL = "Select p.Id, Nvl(k.现价, Nvl(j.上次售价*" & db包装系数 & ",p.现价*" & db包装系数 & ")) as 现价, j.加成率 / 100 As 加成率, Nvl(k.成本价, j.成本价*" & db包装系数 & ") As 成本价, j.指导批发价," & vbNewLine & _
                    "       j.指导零售价, j.差价让利比, p.收入项目id, p.执行日期, p.收入项目id, i.名称 As 收入名称" & vbNewLine & _
                    "From 收费价目 P, 收入项目 I, 药品规格 J," & vbNewLine & _
                    "     (Select Decode(Sum(Nvl(实际数量, 0)), 0, Null, Sum(Round(零售价*" & db包装系数 & ", " & mintPriceDigit & ") * round(实际数量/" & db包装系数 & "," & mintNumberDigit & ")) / Sum(round(实际数量/" & db包装系数 & "," & mintNumberDigit & "))) As 现价," & vbNewLine & _
                    "              Decode(Sum(Nvl(实际数量, 0)), 0, Null, Sum(Round(平均成本价*" & db包装系数 & ", " & mintCostDigit & ") * round(实际数量/" & db包装系数 & "," & mintNumberDigit & ")) / Sum(round(实际数量/" & db包装系数 & "," & mintNumberDigit & "))) As 成本价" & vbNewLine & _
                    "       From 药品库存" & vbNewLine & _
                    "       Where 性质 = 1 And 药品id = [1] ) K" & vbNewLine & _
                    "Where p.收入项目id = i.Id And p.收费细目id = j.药品id And p.收费细目id = [1] And" & vbNewLine & _
                    "      (p.终止日期 Is Null Or Sysdate Between p.执行日期 And p.终止日期) " & GetPriceClassString("P")

            End If
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "查询药品", lngDrugID)
            If rsTemp.RecordCount = 0 Then
                MsgBox "该药品不存在，请重新建立该药品卡片！", vbInformation, gstrSysName
                Exit Sub
            End If
            .TextMatrix(Row, menuPriceCol.原价id) = rsTemp!Id
            .TextMatrix(Row, menuPriceCol.收入项目ID) = IIf(IsNull(rsTemp!收入项目ID), 0, rsTemp!收入项目ID)
            .TextMatrix(Row, menuPriceCol.加成率) = zlStr.FormatEx(IIf(IsNull(rsTemp!加成率), 0, rsTemp!加成率), 5, , True)
            .TextMatrix(Row, menuPriceCol.差价让利比) = IIf(IsNull(rsTemp!差价让利比), 100, rsTemp!差价让利比)
            
            '成本价，售价不用包装换算，在之前的SQL中已经换算了
            .TextMatrix(Row, menuPriceCol.原成本价) = zlStr.FormatEx(IIf(IsNull(rsTemp!成本价), 0, rsTemp!成本价), mintCostDigit, , True)
            .TextMatrix(Row, menuPriceCol.现成本价) = zlStr.FormatEx(IIf(IsNull(rsTemp!成本价), 0, rsTemp!成本价), mintCostDigit, , True)
            .TextMatrix(Row, menuPriceCol.原零售价) = zlStr.FormatEx(IIf(IsNull(rsTemp!现价), 0, rsTemp!现价), mintPriceDigit, , True)
            .TextMatrix(Row, menuPriceCol.现零售价) = zlStr.FormatEx(IIf(IsNull(rsTemp!现价), 0, rsTemp!现价), mintPriceDigit, , True)
            
            .TextMatrix(Row, menuPriceCol.原采购限价) = zlStr.FormatEx(IIf(IsNull(rsTemp!指导批发价), 0, rsTemp!指导批发价) * db包装系数, mintCostDigit, , True)
            .TextMatrix(Row, menuPriceCol.现采购限价) = .TextMatrix(Row, menuPriceCol.原采购限价)
            .TextMatrix(Row, menuPriceCol.原指导售价) = zlStr.FormatEx(IIf(IsNull(rsTemp!指导零售价), 0, rsTemp!指导零售价) * db包装系数, mintPriceDigit, , True)
            .TextMatrix(Row, menuPriceCol.现指导售价) = .TextMatrix(Row, menuPriceCol.原指导售价)

            Call GetDrugStore(lngDrugID, Row)
            If Row = .rows - 1 Then '最后一行才新增行
                .rows = .rows + 1
                .RowHeight(.rows - 1) = mlngRowHeight
                Row = Row + 1
            End If
        End With
'        If mint调价 = 0 And mbln时价药品按批次调价 = True Then '售价调价
'            Call GetDrugStore(lngDrugID, db包装系数)
'        ElseIf mint调价 <> 0 Then

'        End If
'        Call SetRowHidden(lngDrugID)

        rsReturn.MoveNext
    Next
    Call setColEdit

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetDrugStore(ByVal lngDrugID As Long, ByVal intRow As Integer)
    Dim rsTemp As ADODB.Recordset
    Dim dblOldCost As Double
    Dim dblOldPrice As Double
    Dim dblNewCost As Double
    Dim dblNewPrice As Double
    Dim dbl加成率 As Double
    Dim lngCurRow As Long     '当前行
    Dim i As Long
    Dim dbl发票金额 As Double
    Dim str药品名称 As String
    Dim str发票 As String
    Dim str发票日期 As String
    Dim rsPirce As ADODB.Recordset
    Dim rsCost As ADODB.Recordset
    Dim dbl包装换算 As Double
    Dim bln相同药品 As Boolean
    Dim lng药品id As Long
    Dim str单位 As String
    Dim bln是否执行 As Boolean
    
    '功能：为库存列表填充数据
    '参数：药品id

    On Error GoTo errHandle
    '先检查是否有重复的数据，如果有就先清除掉重复的数据
    With vsfStore
        For i = .rows - 1 To 1 Step -1
            If Val(.TextMatrix(i, menuStoreCol.药品id)) = mlngOldDrugID And mlngOldDrugID <> 0 Then
                .RemoveItem i
            End If
        Next
    End With

    With vsfPay
        For i = .rows - 1 To 1 Step -1
            If Val(.TextMatrix(i, menuPayCol.药品id)) = mlngOldDrugID And mlngOldDrugID <> 0 Then
                .RemoveItem i
            End If
        Next
    End With

    If mintModal = 0 Or mblnUpdateAdd = True Or mblnBatchItem = True Then
        gstrSQL = "Select s.库房id,s.药品id, d.名称 As 库房, '[' || m.编码 || ']' || m.名称 As 药品, m.规格, m.产地, m.计算单位 售价单位, p.药库单位, s.上次批号 As 批号, nvl(s.实际数量,0) As 数量," & vbNewLine & _
            "       s.批次, Nvl(m.是否变价, 0) 变价, m.Id, Decode(Nvl(m.是否变价, 0), 0, e.现价, Decode(s.零售价,null,Decode(Nvl(s.实际数量, 0), 0, e.现价, s.实际金额 / s.实际数量),s.零售价)) As 时价售价, p.加成率," & vbNewLine & _
            "       Decode(s.平均成本价, null, p.成本价, s.平均成本价) As 成本价, s.上次供应商id, n.名称 As 供应商, s.效期, s.上次产地 As 产地" & vbNewLine & _
            " From 药品库存 S, 部门表 D, 收费项目目录 M, 药品规格 P, 供应商 N, 收费价目 E" & vbNewLine & _
            " Where d.Id = s.库房id And s.药品id = m.Id And m.Id = p.药品id And Nvl(s.上次供应商id, 0) = n.Id(+) And m.Id = e.收费细目id And" & vbNewLine & _
            " s.性质 = 1 And s.药品id = [1] And Sysdate Between e.执行日期 And e.终止日期  " & vbNewLine & _
            " And Not (zl_fun_getbatchpro(s.库房id,[1])=1 And Nvl(S.批次,0) = 0 And S.可用数量 < 0 And S.实际数量 = 0 And S.实际金额 = 0 And S.实际差价 = 0) " & vbNewLine & _
            GetPriceClassString("E") & vbNewLine & _
            " Order By s.药品id,s.库房id, s.上次批号,s.批次 "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption, lngDrugID)

        If mlng供应商ID > 0 Then
            rsTemp.Filter = "上次供应商ID=" & mlng供应商ID
        End If
    Else '修改，查阅
        gstrSQL = "select (sysdate-执行日期 ) as 是否执行 from 调价汇总记录 where 调价号=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "是否执行", txtNO.Text)
          
        bln是否执行 = rsTemp!是否执行 > 0
        
        If cboPriceMethod.Text = "售价成本价一起调价" Then
            If bln是否执行 = True Then
                gstrSQL = "Select Distinct b.库房id, c.名称 As 库房, b.药品id, b.供药单位id as 上次供应商id, '[' || e.编码 || ']' || e.名称 As 药品, e.规格, d.名称 As 供应商, b.新成本价," & vbNewLine & _
                                "                b.原成本价, b.发票号, b.发票日期, b.发票金额, b.产地, b.批次, b.批号, e.是否变价 As 变价, e.计算单位 As 售价单位, f.药库单位," & vbNewLine & _
                                "                Nvl(b.实际数量, Nvl(b.实际数量, 0)) As 数量, f.加成率, b.效期, Decode(Nvl(e.是否变价, 0), 0, h.原价, b.原零售价) As 原零售价," & vbNewLine & _
                                "                Decode(Nvl(e.是否变价, 0), 0, h.现价, Nvl(b.现零售价, h.现价)) As 现零售价" & vbNewLine & _
                                "From 部门表 C, 供应商 D, 收费项目目录 E, 药品规格 F, 收费价目 H," & vbNewLine & _
                                "     (Select Distinct b.药品id, b.库房id, b.批次, b.批号, b.效期, b.产地, b.原价 As 原成本价, b.现价 As 新成本价, b.发票号, b.发票日期, b.发票金额, b.应付款变动," & vbNewLine & _
                                "                       b.执行日期, g.原价 As 原零售价, g.现价 As 现零售价, b.调价汇总号, i.填写数量 As 实际数量, b.供药单位id" & vbNewLine & _
                                "       From 药品价格记录 B, 药品价格记录 G, 药品收发记录 I" & vbNewLine & _
                                "       Where b.价格类型 = 2 And b.调价汇总号 =[1] And" & vbNewLine & _
                                "             Decode(b.库房id, Null, 1, b.库房id) = Decode(b.库房id, Null, 1, g.库房id(+)) And b.药品id = g.药品id(+) And" & vbNewLine & _
                                "             Decode(b.库房id, Null, 1, Nvl(b.批次,0)) = Decode(b.库房id, Null, 1, Nvl(g.批次(+),0)) And b.调价汇总号 = g.调价汇总号(+) And g.价格类型(+) = 1 And b.收发id = i.Id) B" & vbNewLine & _
                                "Where b.库房id = c.Id And b.供药单位id = d.Id(+) And b.药品id = e.Id And e.Id = f.药品id And b.药品id = h.收费细目id And h.调价汇总号 = b.调价汇总号" & vbNewLine & _
                                "Order By b.药品id, b.库房id, b.批号, b.批次"
            Else
                gstrSQL = "Select Distinct a.库房id,c.名称 as 库房, b.药品id,a.上次供应商id, '[' || e.编码 || ']' ||e.名称 as 药品,e.规格,d.名称 as 供应商, b.新成本价, b.原成本价, b.发票号, b.发票日期, b.发票金额" & _
                                    " ,a.上次产地 as 产地,a.批次,a.上次批号 as 批号,e.是否变价 as 变价,e.计算单位 as 售价单位,f.药库单位,Nvl(b.实际数量, Nvl(a.实际数量, 0)) as 数量,f.加成率,a.效期, " & _
                                    " Decode(Nvl(e.是否变价, 0), 0, h.原价, b.原零售价) As 原零售价,Decode(Nvl(e.是否变价, 0), 0, h.现价, Nvl(b.现零售价, h.现价)) As 现零售价 " & _
                                    " From 药品库存 A,部门表 C,供应商 D,收费项目目录 E,药品规格 F,收费价目 H, " & _
                                         " (Select Distinct b.药品id, b.库房id, b.批次, b.批号, b.效期, b.产地, b.原价 as 原成本价, b.现价 as 新成本价, b.发票号, b.发票日期," & _
                                         " b.发票金额, b.应付款变动, b.执行日期, g.原价 As 原零售价, g.现价 As 现零售价,b.调价汇总号, i.填写数量 As 实际数量 " & _
                                           " From 药品价格记录 B, 药品价格记录 G, 药品收发记录 I " & _
                                           " Where B.价格类型=2 And B.调价汇总号 = [1] " & _
                                    " And Decode(b.库房id, Null, 1,b.库房id) = Decode(b.库房id, Null, 1,g.库房id(+)) And " & _
                                    " b.药品id = g.药品id(+) And Decode(b.库房id, Null, 1,Nvl(b.批次,0)) = Decode(b.库房id, Null, 1,Nvl(g.批次(+),0)) And b.调价汇总号 = g.调价汇总号(+) And g.价格类型(+) = 1 And b.收发id = i.Id(+)) B" & _
                                    " Where a.药品id = b.药品id And Decode(b.库房id, Null, 1,a.库房id) = Decode(b.库房id, Null, 1,b.库房id) and " & _
                                    " Decode(b.库房id, Null, 1,nvl(a.批次,0))=Decode(b.库房id, Null, 1,nvl(b.批次,0)) and a.库房id=c.id and a.上次供应商id=d.id(+) and " & _
                                    " a.药品id=e.id and e.id=f.药品id and a.性质=1 And a.药品id = h.收费细目id And h.调价汇总号 = b.调价汇总号 " & _
                                    " And Not (zl_fun_getbatchpro(a.库房id,a.药品id)=1 And Nvl(a.批次,0) = 0 And a.可用数量 < 0 And a.实际数量 = 0 And a.实际金额 = 0 And a.实际差价 = 0) " & _
                                    " Order By b.药品id, a.库房id, a.上次批号,a.批次 "
            End If
        ElseIf cboPriceMethod.Text = "仅调成本价" Then
            If bln是否执行 = True Then
                '已经执行了取已产生的收发记录数据
                gstrSQL = "Select Distinct a.库房id, c.名称 As 库房, b.药品id, b.供药单位id As 上次供应商id, '[' || e.编码 || ']' || e.名称 As 药品, e.规格, d.名称 As 供应商," & vbNewLine & _
                        "                b.现价 as 新成本价, b.原价 as 原成本价, b.发票号, b.发票日期, b.发票金额, b.产地, b.批次, b.批号, e.是否变价 As 变价, e.计算单位 As 售价单位, f.药库单位," & vbNewLine & _
                        "                nvl(a.填写数量,0) As 数量, f.加成率, b.效期, g.原价 As 原零售价, g.现价 As 现零售价 " & vbNewLine & _
                        "From 药品收发记录 A, 药品价格记录 B, 部门表 C, 供应商 D, 收费项目目录 E, 药品规格 F, 收费价目 G " & vbNewLine & _
                        "Where a.id=b.收发id And a.库房id = c.Id And b.供药单位id = d.Id(+) And" & vbNewLine & _
                        "      a.药品id = e.Id And e.Id = f.药品id And b.价格类型=2 And b.调价汇总号 = [1] and a.单据 = 5 " & vbNewLine & _
                        " And b.药品id = g.收费细目id And Sysdate Between g.执行日期 And g.终止日期 " & _
                        " Order By b.药品id, a.库房id, b.批号,b.批次 "
            Else
                '未执行取价格记录表，库存表等信息
                gstrSQL = "Select Distinct a.库房id, c.名称 As 库房, b.收费细目id As 药品id, a.上次供应商id, '[' || e.编码 || ']' || e.名称 As 药品, e.规格, d.名称 As 供应商," & _
                        " g.现价 As 新成本价, g.原价 As 原成本价, '' 发票号, '' 发票日期, '' 发票金额, a.上次产地 As 产地, a.批次, a.上次批号 As 批号," & _
                        " e.是否变价 As 变价, e.计算单位 As 售价单位, f.药库单位, nvl(a.实际数量,0) As 数量, f.加成率, a.效期, " & _
                        " Decode(Nvl(e.是否变价, 0), 0, b.原价, a.零售价) As 原零售价,Decode(Nvl(e.是否变价, 0), 0, b.现价, a.零售价)  As 现零售价 " & _
                        " From 药品库存 A, 收费价目 B, 部门表 C, 供应商 D, 收费项目目录 E, 药品规格 F, 药品价格记录 G " & _
                        " Where a.药品id = b.收费细目id And a.库房id = c.Id And a.上次供应商id = d.Id(+) And a.药品id = e.Id And e.Id = f.药品id And a.性质 = 1 And" & _
                        " g.调价汇总号 = [1] " & GetPriceClassString("B") & _
                        " And Decode(g.库房id, Null, 1,a.库房id) = Decode(g.库房id, Null, 1,g.库房id) And " & _
                        " a.药品id = g.药品id And Decode(g.库房id, Null, 1,Nvl(a.批次,0)) = Decode(g.库房id, Null, 1,Nvl(g.批次,0)) " & _
                        " And Sysdate Between b.执行日期 And b.终止日期 And g.价格类型 = 2 " & _
                        " And Not (zl_fun_getbatchpro(a.库房id,a.药品id)=1 And Nvl(a.批次,0) = 0 And a.可用数量 < 0 And a.实际数量 = 0 And a.实际金额 = 0 And a.实际差价 = 0) " & _
                        " Order By 药品id, 库房id, 批号,批次 "
            
            End If
        ElseIf cboPriceMethod.Text = "仅调售价" Then
            If bln是否执行 = True Then
                '已经执行了取已产生的收发记录数据
                gstrSQL = "Select Distinct a.库房id, c.名称 As 库房, b.收费细目id As 药品id, a.供药单位id As 上次供应商id, '[' || e.编码 || ']' || e.名称 As 药品, e.规格," & vbNewLine & _
                        "                d.名称 As 供应商, nvl(h.平均成本价,f.成本价) As 新成本价, nvl(h.平均成本价,f.成本价) As 原成本价, '' 发票号, '' 发票日期, '' 发票金额, a.产地, a.批次, a.批号, e.是否变价 As 变价," & vbNewLine & _
                        "                e.计算单位 As 售价单位, f.药库单位, nvl(a.填写数量,0) As 数量, f.加成率, a.效期,a.成本价 As 原零售价, a.零售价 As 现零售价 " & vbNewLine & _
                        "From 药品收发记录 A, 收费价目 B, 部门表 C, 供应商 D, 收费项目目录 E, 药品规格 F, 药品库存 H " & vbNewLine & _
                        "Where a.价格id = b.Id And a.库房id = c.Id And a.供药单位id = d.Id(+) And a.药品id = e.Id And e.Id = f.药品id And" & vbNewLine & _
                        "      b.调价汇总号 = [1] and a.单据=13 And a.费用id Is Null and a.库房id=h.库房id(+) and a.药品id=h.药品id(+) and Nvl(a.批次,0)=nvl(h.批次(+),0) and h.性质(+)=1 " & GetPriceClassString("B") & _
                        " Order By b.收费细目id, a.库房id, a.批号,a.批次 "
            Else
                '未执行取价格记录表，库存表等信息
                gstrSQL = "Select Distinct a.库房id, c.名称 As 库房, b.收费细目id As 药品id, a.上次供应商id, '[' || e.编码 || ']' || e.名称 As 药品, e.规格, d.名称 As 供应商," & _
                                        " nvl(a.平均成本价,f.成本价) As 新成本价, nvl(a.平均成本价,f.成本价) As 原成本价, '' 发票号, '' 发票日期, '' 发票金额, a.上次产地 As 产地, a.批次, a.上次批号 As 批号," & _
                                        " e.是否变价 As 变价, e.计算单位 As 售价单位, f.药库单位, nvl(a.实际数量,0) As 数量, f.加成率, a.效期,b.原价 As 原零售价,b.现价 As 现零售价 " & _
                        " From 药品库存 A, 收费价目 B, 部门表 C, 供应商 D, 收费项目目录 E, 药品规格 F " & _
                        " Where a.药品id = b.收费细目id And a.库房id = c.Id And a.上次供应商id = d.Id(+) And a.药品id = e.Id And e.Id = f.药品id And a.性质 = 1 And" & _
                              " b.调价汇总号 = [1] And Nvl(e.是否变价, 0) = 0 " & GetPriceClassString("B") & _
                              " And Not (zl_fun_getbatchpro(a.库房id,a.药品id)=1 And Nvl(a.批次,0) = 0 And a.可用数量 < 0 And a.实际数量 = 0 And a.实际金额 = 0 And a.实际差价 = 0) "
                gstrSQL = gstrSQL & " Union All " & _
                        "Select Distinct a.库房id, c.名称 As 库房, b.收费细目id As 药品id, a.上次供应商id, '[' || e.编码 || ']' || e.名称 As 药品, e.规格, d.名称 As 供应商," & _
                                        " nvl(a.平均成本价,f.成本价) As 新成本价, nvl(a.平均成本价,f.成本价) As 原成本价, '' 发票号, '' 发票日期, '' 发票金额, a.上次产地 As 产地, a.批次, a.上次批号 As 批号," & _
                                        " e.是否变价 As 变价, e.计算单位 As 售价单位, f.药库单位, nvl(a.实际数量,0) As 数量, f.加成率, a.效期,g.原价 As 原零售价,g.现价 As 现零售价 " & _
                        " From 药品库存 A, 收费价目 B, 部门表 C, 供应商 D, 收费项目目录 E, 药品规格 F, 药品价格记录 G " & _
                        " Where a.药品id = b.收费细目id And a.库房id = c.Id And a.上次供应商id = d.Id(+) And a.药品id = e.Id And e.Id = f.药品id And a.性质 = 1 And" & _
                        " b.调价汇总号 = [1] And Nvl(e.是否变价, 0) = 1 " & GetPriceClassString("B") & _
                        " And Decode(g.库房id, Null, 1,a.库房id) = Decode(g.库房id, Null, 1,g.库房id) And " & _
                        " a.药品id = g.药品id And Decode(g.库房id, Null, 1,Nvl(a.批次,0)) = Decode(g.库房id, Null, 1,Nvl(g.批次,0)) And b.调价汇总号 = g.调价汇总号 And g.价格类型 = 1 " & _
                        " And Not (zl_fun_getbatchpro(a.库房id,a.药品id)=1 And Nvl(a.批次,0) = 0 And a.可用数量 < 0 And a.实际数量 = 0 And a.实际金额 = 0 And a.实际差价 = 0) " & _
                        " Order By 药品id, 库房id, 批号,批次 "
            End If
        End If
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption, txtNO.Text)
    End If

    With vsfStore
        Do While Not rsTemp.EOF
            dbl包装换算 = 0
            dbl发票金额 = 0
            dblOldPrice = 0
            dblNewPrice = 0
            For i = 0 To vsfPrice.rows - 1
                If rsTemp!药品id = vsfPrice.TextMatrix(i, menuPriceCol.药品id) Then
                    dbl包装换算 = vsfPrice.TextMatrix(i, menuPriceCol.包装系数)
                    dblOldPrice = Val(vsfPrice.TextMatrix(i, menuPriceCol.原零售价))
                    dblNewPrice = Val(vsfPrice.TextMatrix(i, menuPriceCol.现零售价))
                    str单位 = vsfPrice.TextMatrix(i, menuPriceCol.单位)
                    Exit For
                End If
            Next
            .rows = .rows + 1
            .TextMatrix(.rows - 1, menuStoreCol.变价) = rsTemp!变价
            Call setColEdit
            .RowHeight(.rows - 1) = mlngRowHeight

            '从空白行开始插入数据
            .TextMatrix(.rows - 1, menuStoreCol.药品id) = rsTemp!药品id
            .TextMatrix(.rows - 1, menuStoreCol.库房) = rsTemp!库房
            .TextMatrix(.rows - 1, menuStoreCol.库房id) = rsTemp!库房id
            .TextMatrix(.rows - 1, menuStoreCol.供应商) = Nvl(rsTemp!供应商, "")
            .TextMatrix(.rows - 1, menuStoreCol.供应商id) = IIf(mlng供应商ID > 0, mlng供应商ID, Nvl(rsTemp!上次供应商ID))
            .TextMatrix(.rows - 1, menuStoreCol.药品) = rsTemp!药品
            str药品名称 = rsTemp!药品

            .TextMatrix(.rows - 1, menuStoreCol.规格) = IIf(IsNull(rsTemp!规格), "", rsTemp!规格)
            .TextMatrix(.rows - 1, menuStoreCol.单位) = str单位
            .TextMatrix(.rows - 1, menuStoreCol.批号) = IIf(IsNull(rsTemp!批号), "", rsTemp!批号)
            .TextMatrix(.rows - 1, menuStoreCol.效期) = Format(IIf(IsNull(rsTemp!效期), "", rsTemp!效期), "YYYY-MM-DD")
            .TextMatrix(.rows - 1, menuStoreCol.产地) = IIf(IsNull(rsTemp!产地), "", rsTemp!产地)
            .TextMatrix(.rows - 1, menuStoreCol.数量) = zlStr.FormatEx(rsTemp!数量 / dbl包装换算, mintNumberDigit, , True)
            .TextMatrix(.rows - 1, menuStoreCol.包装系数) = dbl包装换算
            .TextMatrix(.rows - 1, menuStoreCol.批次) = Nvl(rsTemp!批次, 0)
'            .TextMatrix(.rows - 1, menuStoreCol.变价) = rsTemp!变价


            If mintModal = 0 Or mblnUpdateAdd = True Or mblnBatchItem = True Then
                dblOldCost = IIf(IsNull(rsTemp!成本价), 0, rsTemp!成本价) * dbl包装换算

                If mdbl加成率 > 0 Then
                    dbl加成率 = Round(mdbl加成率 / 100, 7)
                ElseIf dblOldCost > 0 Then
                    dbl加成率 = Round(IIf(rsTemp!变价 = 1, rsTemp!时价售价 * dbl包装换算, dblOldPrice) / dblOldCost - 1, 7)
                Else
                    dbl加成率 = Round(rsTemp!加成率 / 100, 2)
                End If
                If 1 + dbl加成率 = 0 Then
                    dblNewCost = 0
                Else
                    dblNewCost = rsTemp!时价售价 * dbl包装换算 / (1 + dbl加成率)
                End If
                If dbl加成率 = -1 Then dbl加成率 = 0

                .TextMatrix(.rows - 1, menuStoreCol.原零售价) = zlStr.FormatEx(IIf(rsTemp!变价 = 1, rsTemp!时价售价 * dbl包装换算, dblOldPrice), mintPriceDigit, , True)
                .TextMatrix(.rows - 1, menuStoreCol.现零售价) = zlStr.FormatEx(IIf(rsTemp!变价 = 1, rsTemp!时价售价 * dbl包装换算, dblOldPrice), mintPriceDigit, , True)
                .TextMatrix(.rows - 1, menuStoreCol.售价盈亏) = Format(Format(rsTemp!数量 / dbl包装换算 * Val(.TextMatrix(.rows - 1, menuStoreCol.现零售价)), mstrMoneyFormat) - Format(rsTemp!数量 / dbl包装换算 * Val(.TextMatrix(.rows - 1, menuStoreCol.原零售价)), mstrMoneyFormat), mstrMoneyFormat)
                
                .TextMatrix(.rows - 1, menuStoreCol.加成率) = zlStr.FormatEx(zlStr.FormatEx(dbl加成率, 5, , True) * 100, 5, , True)
                .TextMatrix(.rows - 1, menuStoreCol.原成本价) = zlStr.FormatEx(dblOldCost, mintCostDigit, , True)
                .TextMatrix(.rows - 1, menuStoreCol.现成本价) = zlStr.FormatEx(dblNewCost, mintCostDigit, , True)
                .TextMatrix(.rows - 1, menuStoreCol.成本盈亏) = Format(Format(Val(.TextMatrix(.rows - 1, menuStoreCol.现成本价)) * Val(.TextMatrix(.rows - 1, menuStoreCol.数量)), mstrMoneyFormat) - Format(Val(.TextMatrix(.rows - 1, menuStoreCol.原成本价)) * Val(.TextMatrix(.rows - 1, menuStoreCol.数量)), mstrMoneyFormat), mstrMoneyFormat)
                dbl发票金额 = dbl发票金额 + (dblNewCost - dblOldCost) * Val(.TextMatrix(.rows - 1, menuStoreCol.数量))
                
                '为应付记录表赋值
                If mint调价 = 1 Or mint调价 = 2 Then
                    If vsfPay.rows > 1 Then
                        bln相同药品 = False
                        For i = 1 To vsfPay.rows - 1
                            If vsfPay.TextMatrix(i, menuPayCol.药品id) = rsTemp!药品id Then
                                bln相同药品 = True
                                Exit For
                            End If
                        Next
                        If bln相同药品 = True Then
                            vsfPay.TextMatrix(i, menuPayCol.发票金额) = zlStr.FormatEx(Val(vsfPay.TextMatrix(i, menuPayCol.发票金额)) + dbl发票金额, mintMoneyDigit, , True)
                        Else
                            vsfPay.rows = vsfPay.rows + 1
                            vsfPay.RowHeight(vsfPay.rows - 1) = mlngRowHeight
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.药品id) = rsTemp!药品id
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.药品) = str药品名称
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.发票号) = str发票
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.发票日期) = Format(str发票日期, "yyyy-mm-dd")
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.发票金额) = zlStr.FormatEx(dbl发票金额, mintMoneyDigit, , True)
                        End If
                    Else
                        vsfPay.rows = vsfPay.rows + 1
                        vsfPay.RowHeight(vsfPay.rows - 1) = mlngRowHeight
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.药品id) = rsTemp!药品id
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.药品) = str药品名称
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.发票号) = str发票
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.发票日期) = Format(str发票日期, "yyyy-mm-dd")
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.发票金额) = zlStr.FormatEx(dbl发票金额, mintMoneyDigit, , True)
                    End If
                End If
            Else
                .TextMatrix(.rows - 1, menuStoreCol.原零售价) = zlStr.FormatEx(Val(rsTemp!原零售价) * dbl包装换算, mintPriceDigit, , True)
                .TextMatrix(.rows - 1, menuStoreCol.现零售价) = zlStr.FormatEx(Val(rsTemp!现零售价) * dbl包装换算, mintPriceDigit, , True)
                .TextMatrix(.rows - 1, menuStoreCol.售价盈亏) = Format(Format(rsTemp!数量 / dbl包装换算 * Val(.TextMatrix(.rows - 1, menuStoreCol.现零售价)), mstrMoneyFormat) - Format(rsTemp!数量 / dbl包装换算 * Val(.TextMatrix(.rows - 1, menuStoreCol.原零售价)), mstrMoneyFormat), mstrMoneyFormat)
                .TextMatrix(.rows - 1, menuStoreCol.原成本价) = zlStr.FormatEx(Nvl(rsTemp!原成本价, 0) * dbl包装换算, mintCostDigit, , True)
                .TextMatrix(.rows - 1, menuStoreCol.现成本价) = zlStr.FormatEx(rsTemp!新成本价 * dbl包装换算, mintCostDigit, , True)
                .TextMatrix(.rows - 1, menuStoreCol.成本盈亏) = Format(Format((rsTemp!新成本价 * dbl包装换算) * Val(.TextMatrix(.rows - 1, menuStoreCol.数量)), mstrMoneyFormat) - Format((Nvl(rsTemp!原成本价, 0) * dbl包装换算) * Val(.TextMatrix(.rows - 1, menuStoreCol.数量)), mstrMoneyFormat), mstrMoneyFormat)
                 
                If cboPriceMethod.Text = "仅调成本价" Or cboPriceMethod.Text = "售价成本价一起调价" Then
                    If rsTemp!新成本价 = 0 Then
                        dbl加成率 = 0
                    Else
                        dbl加成率 = Round(Val(.TextMatrix(.rows - 1, menuStoreCol.现零售价)) / (rsTemp!新成本价 * dbl包装换算) - 1, 7)
                    End If
                    .TextMatrix(.rows - 1, menuStoreCol.加成率) = zlStr.FormatEx(zlStr.FormatEx(dbl加成率, 5, , True) * 100, 5, , True)
                    .TextMatrix(.rows - 1, menuStoreCol.原成本价) = zlStr.FormatEx(Nvl(rsTemp!原成本价, 0) * dbl包装换算, mintCostDigit, , True)
                    .TextMatrix(.rows - 1, menuStoreCol.现成本价) = zlStr.FormatEx(rsTemp!新成本价 * dbl包装换算, mintCostDigit, , True)
                    .TextMatrix(.rows - 1, menuStoreCol.成本盈亏) = Format(Format((rsTemp!新成本价 * dbl包装换算) * Val(.TextMatrix(.rows - 1, menuStoreCol.数量)), mstrMoneyFormat) - Format((Nvl(rsTemp!原成本价, 0) * dbl包装换算) * Val(.TextMatrix(.rows - 1, menuStoreCol.数量)), mstrMoneyFormat), mstrMoneyFormat)
                    dbl发票金额 = dbl发票金额 + (rsTemp!新成本价 * dbl包装换算 - Nvl(rsTemp!原成本价, 0) * dbl包装换算) * Val(.TextMatrix(.rows - 1, menuStoreCol.数量))
                    str发票 = IIf(IsNull(rsTemp!发票号), "", rsTemp!发票号)
                    str发票日期 = IIf(IsNull(rsTemp!发票日期), "", rsTemp!发票日期)
                    
                    '为付款记录列表赋值
                    If vsfPay.rows > 1 Then
                        bln相同药品 = False
                        For i = 1 To vsfPay.rows - 1
                            If vsfPay.TextMatrix(i, menuPayCol.药品id) = rsTemp!药品id Then
                                bln相同药品 = True
                                Exit For
                            End If
                        Next
                        If bln相同药品 = True Then
                            vsfPay.TextMatrix(i, menuPayCol.发票金额) = zlStr.FormatEx(Val(vsfPay.TextMatrix(i, menuPayCol.发票金额)) + dbl发票金额, mintMoneyDigit, , True)
                        Else
                            vsfPay.rows = vsfPay.rows + 1
                            vsfPay.RowHeight(vsfPay.rows - 1) = mlngRowHeight
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.药品id) = rsTemp!药品id
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.药品) = str药品名称
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.发票号) = str发票
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.发票日期) = Format(str发票日期, "yyyy-mm-dd")
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.发票金额) = zlStr.FormatEx(dbl发票金额, mintMoneyDigit, , True)
                        End If
                    Else
                        vsfPay.rows = vsfPay.rows + 1
                        vsfPay.RowHeight(vsfPay.rows - 1) = mlngRowHeight
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.药品id) = rsTemp!药品id
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.药品) = str药品名称
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.发票号) = str发票
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.发票日期) = Format(str发票日期, "yyyy-mm-dd")
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.发票金额) = zlStr.FormatEx(dbl发票金额, mintMoneyDigit, , True)
                    End If
                End If
            End If
            rsTemp.MoveNext
        Loop
        
    End With
    '修改和查阅时重算规格列表平均成本价，售价
    'mintModal 0-新增 1-修改 2-查阅
    If mintModal = 1 Or mintModal = 2 Then
        With vsfStore
            For i = 1 To .rows - 1
                If lng药品id <> .TextMatrix(i, menuStoreCol.药品id) Then
                    Call CaluateAverCost(Val(.TextMatrix(i, menuStoreCol.药品id)))
                    Call CaluateAverOldCost(Val(.TextMatrix(i, menuStoreCol.药品id)))
                    
                    If Val(.TextMatrix(i, menuStoreCol.变价)) = 1 Then
                        Call CaculateAverPirce(Val(.TextMatrix(i, menuStoreCol.药品id)))
                        Call CaculateAverOldPirce(Val(.TextMatrix(i, menuStoreCol.药品id)))
                    End If
                    
                    lng药品id = Val(.TextMatrix(i, menuStoreCol.药品id))
                End If
            Next
        End With
    End If

    If mint调价 = 1 Or mint调价 = 2 Then
        If rsTemp.RecordCount = 0 Then Exit Sub
        TabCtlDetails.Item(1).Visible = True
    End If
            
    Call setColHiddenVsf
    
    '合并单元格
    vsfStore.MergeCol(menuStoreCol.药品) = True
    vsfStore.MergeCol(menuStoreCol.规格) = True
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsfPrice_DblClick()
    With vsfPrice
        If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mconlngCanColColor Then
            .EditCell
            .EditSelStart = 0
            .EditSelLength = Len(.EditText)
        End If
    End With
End Sub

Private Sub vsfPrice_EnterCell()
    Dim i As Integer, j As Integer
    Dim intRow As Integer

    With vsfPrice
        .Editable = flexEDNone
        If .CellBackColor = mconlngColor Then
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
        End If

        If .Col = menuPriceCol.现零售价 Then
            mdblOldPrice = Val(vsfPrice.TextMatrix(.Row, menuPriceCol.现零售价))
        ElseIf .Col = menuPriceCol.现成本价 Then
            mdblOldPrice = Val(vsfPrice.TextMatrix(.Row, menuPriceCol.现成本价))
        End If
    End With
    With vsfStore
        If Val(vsfPrice.TextMatrix(vsfPrice.Row, menuPriceCol.药品id)) = 0 Then Exit Sub
        If .rows > 1 Then
            .Select 0, 0, 0, 0
            For i = 1 To .rows - 1
                If Val(vsfPrice.TextMatrix(vsfPrice.Row, menuPriceCol.药品id)) = Val(.TextMatrix(i, menuStoreCol.药品id)) Then
                    If j = 0 Then j = i
                    .Select j, 3, j, .Cols - 1
                    .TopRow = j
                    intRow = intRow + 1
                End If
                .CellBorderRange i, 0, i, .Cols - 1, mlngNoneBorderColor, 0, 0, 0, 0, 0, 0
            Next
            
            For i = j To j + intRow - 1
                If i = j Then .CellBorderRange i, 0, i, .Cols - 1, mlngBorderColor, 0, 2, 0, 0, 0, 2
                If i = j + intRow - 1 Then .CellBorderRange i, 0, i, .Cols - 1, mlngBorderColor, 0, 0, 0, 2, 0, 2
                If i = j And i = j + intRow - 1 Then .CellBorderRange i, 0, i, .Cols - 1, mlngBorderColor, 0, 2, 0, 2, 0, 2
            Next
        End If
    End With
    
    Call SetBorder '设置行选中边框
End Sub

Private Sub SetBorder()
    '设置行选中边框
    Dim intRow As Integer
    
    With vsfPrice
        If .rows <> 1 Then
            For intRow = 1 To .rows - 1
                .CellBorderRange intRow, 0, intRow, .Cols - 1, mlngNoneBorderColor, 0, 0, 0, 0, 0, 0
            Next
            
            .CellBorderRange .Row, menuPriceCol.药品, .Row, menuPriceCol.现零售价, mlngBorderColor, 0, 2, 0, 2, 0, 2
        End If
    End With
End Sub

Private Sub vsfPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intRow As Integer
    Dim intCol As Integer
    Dim lngDrugID As Long
    Dim strRow As String
    Dim int删除类型 As Integer
    
    With vsfPrice
        If KeyCode = vbKeyReturn Then
            If .Col <> menuPriceCol.现零售价 Then '成本价调价
                If .Col = menuPriceCol.药品 And cboPriceMethod.Text = "仅调成本价" Then
                    .Col = menuPriceCol.现成本价
'                    .EditCell
                ElseIf .Col = menuPriceCol.药品 And cboPriceMethod.Text = "仅调售价" Then
                    .Col = menuPriceCol.现零售价
'                    .EditCell
                ElseIf .Col = menuPriceCol.现成本价 And cboPriceMethod.Text = "仅调成本价" Then
                    If .Row = .rows - 1 And Val(.TextMatrix(.Row, menuPriceCol.药品id)) <> 0 Then
                        .rows = .rows + 1
                        .Row = .Row + 1
                        .Col = menuPriceCol.药品
                        .RowHeight(.rows - 1) = mlngRowHeight
'                        .EditCell
                        Call setColEdit
                    ElseIf Val(.TextMatrix(.Row, menuPriceCol.药品id)) <> 0 Then
                        .ColComboList(menuPriceCol.药品) = ""
                        .Row = .Row + 1
                        .Col = menuPriceCol.药品
                    End If
                ElseIf .Col = menuPriceCol.药品 And cboPriceMethod.Text = "售价成本价一起调价" Then
                    .Col = menuPriceCol.现成本价
'                    .EditCell
                ElseIf .Col = menuPriceCol.现成本价 And cboPriceMethod.Text = "售价成本价一起调价" Then
                    .Col = menuPriceCol.现零售价
'                    .EditCell
                ElseIf .Col = menuPriceCol.现零售价 And cboPriceMethod.Text = "售价成本价一起调价" Then
                    If .Row = .rows - 1 Then
                        .rows = .rows + 1
                        .Row = .Row + 1
                        .Col = menuPriceCol.药品
                        .RowHeight(.rows - 1) = mlngRowHeight
'                        .EditCell
                        Call setColEdit
                    ElseIf Val(.TextMatrix(.Row, menuPriceCol.药品id)) <> 0 Then
                        .ColComboList(menuPriceCol.药品) = ""
                        .Row = .Row + 1
                        .Col = menuPriceCol.药品
'                        .EditCell
                    End If
                Else
                    .Col = .Col + 1
'                    .EditCell
                End If
            Else
                If Val(.TextMatrix(.Row, menuPriceCol.药品id)) <> 0 And .Row = .rows - 1 Then
                    .ColComboList(menuPriceCol.药品) = ""
                    .rows = .rows + 1
                    .Row = .Row + 1
                    .Col = menuPriceCol.药品
                    .RowHeight(.rows - 1) = mlngRowHeight
'                    .EditCell
                    Call setColEdit
                ElseIf Val(.TextMatrix(.Row, menuPriceCol.药品id)) <> 0 Then
                    .ColComboList(menuPriceCol.药品) = ""
                    .Row = .Row + 1
                    .Col = menuPriceCol.药品
'                    .EditCell
                End If
            End If
        ElseIf KeyCode = vbKeyDelete Then
            lngDrugID = Val(vsfPrice.TextMatrix(vsfPrice.Row, menuPriceCol.药品id))
            
            '修改模式时删除一条价格表中数据，则清楚未执行价格
            If mintModal = 1 Then
                'Private mint调价 As Integer     '0-调售价;1-调成本价;2-调售价及成本价
                '删除方式_In   In Number := 0 --0-所有;1-售价;2-成本价
                If mint调价 = 0 Then
                    int删除类型 = 1
                ElseIf mint调价 = 1 Then
                    int删除类型 = 2
                Else
                    int删除类型 = 0
                End If
                
                gstrSQL = "Zl_药品未执行价格_Delete(" & lngDrugID & "," & int删除类型 & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, MStrCaption)
            End If
            
            If .rows > 2 Then
                .RemoveItem .Row
            Else
                For intCol = 0 To .Cols - 1
                    .TextMatrix(.Row, intCol) = ""
                Next
            End If

            With vsfStore
                If lngDrugID = 0 Then Exit Sub
                For intRow = .rows - 1 To 1 Step -1
                    If Val(.TextMatrix(intRow, menuStoreCol.药品id)) = lngDrugID Then
                        .RemoveItem intRow
                    End If
                Next
            End With

            With vsfPay
                If lngDrugID = 0 Then Exit Sub
                For intRow = .rows - 1 To 1 Step -1
                    If Val(.TextMatrix(intRow, menuPayCol.药品id)) = lngDrugID Then
                        .RemoveItem intRow
                    End If
                Next
            End With
        End If
    End With
End Sub

Private Sub vsfPrice_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim mrsReturn As Recordset
    Dim rsTemp As Recordset
    Dim vRect As RECT
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim strkey As String
    Dim lngDrugID As Long
    Dim intCurrentPirce As Integer '是否是时价

    On Error GoTo errHandle
    If KeyCode <> vbKeyReturn Then Exit Sub
    mBlnClick = True
    vRect = zlControl.GetControlRect(vsfPrice.hWnd) '获取位置
    dblLeft = vRect.Left + vsfPrice.CellLeft
    dblTop = vRect.Top + vsfPrice.CellTop + vsfPrice.CellHeight

    With vsfPrice
        strkey = .EditText
        Select Case Col
        Case menuPriceCol.药品
            If grsMaster.State = adStateClosed Then
                Call SetSelectorRS(1, "", 0, , , , , , , , , True)
            End If
            Set mrsReturn = frmSelector.ShowME(Me, 1, 1, strkey, dblLeft, dblTop, , , , , , , , , False, mstrPrivs)
            If mrsReturn.RecordCount = 0 Then Exit Sub
            mblnUpdateAdd = True
            Call GetDrugPirce(mrsReturn, Row)
            mblnUpdateAdd = False
        End Select
    End With

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CheckDoubleDrug(ByVal rsTemp As ADODB.Recordset) As ADODB.Recordset
    '检查是否有重复的药品
    'lngDrugId 药品id
    '返回值 true-存在重复值 false-不存在重复值
    Dim i As Integer
    Dim j As Integer
    Dim strTemp As String
    Dim strName As String
    Dim intCount As Integer
    Dim intLength As Integer

    If rsTemp.RecordCount = 0 Then Exit Function
    rsTemp.MoveFirst
    With vsfPrice
        For i = 0 To rsTemp.RecordCount - 1
            For j = 1 To .rows - 1
                If Val(.TextMatrix(j, menuPriceCol.药品id)) = rsTemp!药品id Then
                    strTemp = strTemp & " 药品id <> " & rsTemp!药品id & " and "
                    intCount = intCount + 1
                    If intCount < 5 Then
                        strName = strName & rsTemp!通用名 & " "
                    End If
                End If
            Next
            rsTemp.MoveNext
        Next
    End With

    If strTemp <> "" Then
        intLength = LenB(StrConv(strTemp, vbFromUnicode)) '得到字符串长度
        Do Until Mid(strTemp, intLength, 3) = "and" '从后向前查找倒数第一个"and"
           intLength = intLength - 1
        Loop
        strTemp = Left(strTemp, intLength - 1) '倒数第一个"and"之前的字符串

        rsTemp.Filter = strTemp
        MsgBox strName & "等" & intCount & "种药品在列表中已经存在，已存在药品不再添加！", vbInformation, gstrSysName
    End If

    Set CheckDoubleDrug = rsTemp
End Function

Private Sub vsfPrice_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        With vsfPrice
            If .Col = menuPriceCol.药品 Then
                .Editable = flexEDKbdMouse
                Exit Sub
            End If
            If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mconlngCanColColor Then
                .Editable = flexEDKbdMouse
            Else
                .Editable = flexEDNone
            End If
        End With
    End If
End Sub

Private Sub vsfPrice_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strkey As String
    Dim intDigit As Integer

    With vsfPrice
        strkey = .EditText
        If .Col = menuPriceCol.现成本价 Then
            mdbl成本价 = Val(.TextMatrix(Row, Col))
        End If
    End With

    If Col = menuPriceCol.现成本价 Or Col = menuPriceCol.现零售价 Then
        If KeyAscii = vbKeyReturn Then Exit Sub
        If KeyAscii <> vbKeyBack Then
            Select Case Col
                Case menuPriceCol.现成本价
                    intDigit = mintCostDigit
                Case menuPriceCol.现零售价
                    intDigit = mintPriceDigit
            End Select

            If KeyAscii = vbKeyDelete Then
                If InStr(1, strkey, ".") > 0 Then
                    KeyAscii = 0
                End If
            ElseIf KeyAscii = Asc(".") Or (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
                If vsfPrice.EditSelLength = Len(strkey) Then Exit Sub
                If InStr(strkey, ".") <> 0 And Chr(KeyAscii) = "." Then   '只能存在一个小数点
                    KeyAscii = 0
                    Exit Sub
                End If
                If Len(Mid(strkey, InStr(1, strkey, ".") + 1)) >= intDigit And strkey Like "*.*" Then
                    KeyAscii = 0
                    Exit Sub
                Else
                    Exit Sub
                End If
            Else
                KeyAscii = 0
            End If
        End If
    ElseIf Col = menuPriceCol.药品 Then
        If InStr("`~!@#$%^&*()_-+={[}]|\:;""'<,>.?/", Chr(KeyAscii)) > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub vsfPrice_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If Col = menuPriceCol.药品 Then
        vsfPrice.ColComboList(menuPriceCol.药品) = "|..."
    End If
End Sub

Private Sub setColEdit()
    '功能：设置列是否可以修改
    '不能修改的列颜色为灰色，能修改的列颜色为白色
    Dim intCol As Integer
    Dim intRow As Integer

    With vsfPrice
        .Cell(flexcpBackColor, 1, 1, .rows - 1, .Cols - 1) = mconlngColor
        If cboPriceMethod.Text = "仅调售价" Then
            .Cell(flexcpBackColor, 1, menuPriceCol.药品, .rows - 1, menuPriceCol.药品) = mconlngCanColColor
            .Cell(flexcpBackColor, 1, menuPriceCol.现零售价, .rows - 1, menuPriceCol.现零售价) = mconlngCanColColor
        ElseIf cboPriceMethod.Text = "仅调成本价" Then
            .Cell(flexcpBackColor, 1, menuPriceCol.药品, .rows - 1, menuPriceCol.药品) = mconlngCanColColor
            .Cell(flexcpBackColor, 1, menuPriceCol.现成本价, .rows - 1, menuPriceCol.现成本价) = mconlngCanColColor
        Else
            .Cell(flexcpBackColor, 1, menuPriceCol.药品, .rows - 1, menuPriceCol.药品) = mconlngCanColColor
            .Cell(flexcpBackColor, 1, menuPriceCol.现成本价, .rows - 1, menuPriceCol.现成本价) = mconlngCanColColor
            .Cell(flexcpBackColor, 1, menuPriceCol.现零售价, .rows - 1, menuPriceCol.现零售价) = mconlngCanColColor
        End If

    End With

    With vsfStore
        If .rows = 1 Then Exit Sub
        .Cell(flexcpBackColor, 1, 0, .rows - 1, .Cols - 1) = mconlngColor
        If cboPriceMethod.Text = "仅调售价" Then
            .Cell(flexcpBackColor, 1, menuStoreCol.现零售价, .rows - 1, menuStoreCol.现零售价) = mconlngCanColColor
        ElseIf cboPriceMethod.Text = "仅调成本价" Then
'            .Cell(flexcpBackColor, 1, menuStoreCol.加成率, .rows - 1, menuStoreCol.加成率) = mconlngCanColColor
            .Cell(flexcpBackColor, 1, menuStoreCol.现成本价, .rows - 1, menuStoreCol.现成本价) = mconlngCanColColor
        Else
            .Cell(flexcpBackColor, 1, menuStoreCol.加成率, .rows - 1, menuStoreCol.加成率) = mconlngCanColColor
            .Cell(flexcpBackColor, 1, menuStoreCol.现成本价, .rows - 1, menuStoreCol.现成本价) = mconlngCanColColor
            .Cell(flexcpBackColor, 1, menuStoreCol.现零售价, .rows - 1, menuStoreCol.现零售价) = mconlngCanColColor
        End If
        If .rows > 1 Then
            For intRow = 1 To .rows - 1
                If Val(.TextMatrix(intRow, menuStoreCol.变价)) = 1 And mbln时价药品按批次调价 = True And mint调价 <> 1 Then
                    .Cell(flexcpBackColor, intRow, menuStoreCol.现零售价, intRow, menuStoreCol.现零售价) = mconlngCanColColor
                Else
                    .Cell(flexcpBackColor, intRow, menuStoreCol.现零售价, intRow, menuStoreCol.现零售价) = mconlngColor
                End If
                If mbln成本价按库房批次调整 = True And mint调价 <> 0 Then
                    .Cell(flexcpBackColor, intRow, menuStoreCol.现成本价, intRow, menuStoreCol.现成本价) = mconlngCanColColor
                Else
                    .Cell(flexcpBackColor, intRow, menuStoreCol.现成本价, intRow, menuStoreCol.现成本价) = mconlngColor
                End If
            Next
        End If
    End With

    With vsfPay
        If .rows = 1 Then Exit Sub
        .Cell(flexcpBackColor, 1, 0, .rows - 1, .Cols - 1) = mconlngColor
        .Cell(flexcpBackColor, 1, menuPayCol.发票号, .rows - 1, menuPayCol.发票号) = mconlngCanColColor
        .Cell(flexcpBackColor, 1, menuPayCol.发票日期, .rows - 1, menuPayCol.发票日期) = mconlngCanColColor
        .Cell(flexcpBackColor, 1, menuPayCol.发票金额, .rows - 1, menuPayCol.发票金额) = mconlngCanColColor
    End With
End Sub


Private Sub vsfPrice_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        vsfPrice.Editable = flexEDNone
        If vsfPrice.Col = menuPriceCol.药品 And mintModal <> 2 Then
            vsfPrice.ColComboList(menuPriceCol.药品) = "|..."
            vsfPrice.Editable = flexEDKbdMouse
        End If
    End If
End Sub

Private Sub vsfPrice_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lngDrugID As Long
    Dim dblSalePrice As Double
    Dim intRow As Integer
    Dim dbl加成率 As Double

    With vsfPrice
        If .EditText = "" Then Exit Sub
        lngDrugID = Val(.TextMatrix(Row, menuPriceCol.药品id))
        If lngDrugID = 0 Then Exit Sub

        Select Case Col
            Case menuPriceCol.现成本价
                If Val(.EditText) < 0 Then
                    MsgBox "成本价不能为负数！", vbExclamation, gstrSysName
                    Cancel = True
                End If
                If Not IsNumeric(.EditText) Then
                    Cancel = True
                    Exit Sub
                End If
                If .EditText > 9999999 Then
                    MsgBox "成本价过大，请重新输入！", vbInformation, gstrSysName
                    Cancel = True
                    Exit Sub
                End If
                .EditText = zlStr.FormatEx(.EditText, mintPriceDigit, , True)
                If mbln现价提示 = True Then
                    If Val(.EditText) > Val(.TextMatrix(Row, menuPriceCol.原采购限价)) Then
                        If MsgBox("现成本价高于采购价限价" & Val(.TextMatrix(.Row, menuPriceCol.原采购限价)) & "。" & vbCrLf & "继续吗？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                            Cancel = True
                            Exit Sub
                        Else
                            .TextMatrix(.Row, menuPriceCol.现采购限价) = zlStr.FormatEx(.EditText, mintCostDigit, , True)
                        End If
                    End If
                Else
                    If Val(.EditText) > Val(.TextMatrix(Row, menuPriceCol.原采购限价)) Then
                        .TextMatrix(.Row, menuPriceCol.现采购限价) = zlStr.FormatEx(.EditText, mintCostDigit, , True)
                    End If
                End If

                If cbo售价计算方式.Text = "售价按分段加成计算" And .TextMatrix(.Row, menuPriceCol.是否变价) = "1" And mint调价 = 2 Then
                    Call get分段加成售价(lngDrugID, Val(.TextMatrix(.Row, menuPriceCol.包装系数)), Val(.EditText), dblSalePrice)
                    If dblSalePrice = 0 Then
                        .EditText = mdbl成本价
                        .TextMatrix(vsfPrice.Row, menuPriceCol.现成本价) = zlStr.FormatEx(.EditText, mintCostDigit, , True)
                        Exit Sub
                    End If
                    dblSalePrice = dblSalePrice + (Val(.TextMatrix(.Row, menuPriceCol.原指导售价)) - dblSalePrice) * (1 - Val(.TextMatrix(.Row, menuPriceCol.差价让利比)) / 100)
                    .TextMatrix(.Row, menuPriceCol.现零售价) = zlStr.FormatEx(dblSalePrice, mintPriceDigit, , True)
                    
                    '调了售价应该同步更新库存列表价格信息
                    If vsfStore.rows > 1 Then
                        For intRow = 1 To vsfStore.rows - 1
                            If vsfStore.TextMatrix(intRow, menuStoreCol.药品id) = .TextMatrix(.Row, menuPriceCol.药品id) Then
                                vsfStore.TextMatrix(intRow, menuStoreCol.现零售价) = zlStr.FormatEx(dblSalePrice, mintPriceDigit, , True)
'                                vsfStore.TextMatrix(intRow, menuStoreCol.售价盈亏) = Format(Val(vsfStore.TextMatrix(intRow, menuStoreCol.数量)) * (Val(vsfStore.TextMatrix(intRow, menuStoreCol.现零售价)) - Val(vsfStore.TextMatrix(intRow, menuStoreCol.原零售价))), mstrMoneyFormat)
                                vsfStore.TextMatrix(intRow, menuStoreCol.售价盈亏) = Format(Format(Val(vsfStore.TextMatrix(intRow, menuStoreCol.数量)) * Val(vsfStore.TextMatrix(intRow, menuStoreCol.现零售价)), mstrMoneyFormat) - Format(Val(vsfStore.TextMatrix(intRow, menuStoreCol.数量)) * Val(vsfStore.TextMatrix(intRow, menuStoreCol.原零售价)), mstrMoneyFormat), mstrMoneyFormat)
                                
                                If Val(vsfStore.TextMatrix(intRow, menuStoreCol.现成本价)) <> 0 Then
                                    dbl加成率 = zlStr.FormatEx(zlStr.FormatEx(((Val(vsfStore.TextMatrix(intRow, menuStoreCol.现零售价))) / Val(vsfStore.TextMatrix(intRow, menuStoreCol.现成本价)) - 1), 5, , True) * 100, 5, , True)
                                Else
                                    dbl加成率 = 0
                                End If
                                vsfStore.TextMatrix(intRow, menuStoreCol.加成率) = dbl加成率
                            End If
                        Next
                    End If
                ElseIf cbo售价计算方式 = "售价按固定比例计算" And .TextMatrix(.Row, menuPriceCol.是否变价) = "1" And mint调价 = 2 Then
                    dblSalePrice = Val(.EditText) * (1 + Val(.TextMatrix(.Row, menuPriceCol.加成率)))
                    If dblSalePrice > Val(.TextMatrix(.Row, menuPriceCol.原指导售价)) Then dblSalePrice = Val(.TextMatrix(.Row, menuPriceCol.原指导售价))
                    .TextMatrix(.Row, menuPriceCol.现零售价) = zlStr.FormatEx(dblSalePrice, mintPriceDigit, , True)
                    
                    '调了售价应该同步更新库存列表价格信息
                    If vsfStore.rows > 1 Then
                        For intRow = 1 To vsfStore.rows - 1
                            If vsfStore.TextMatrix(intRow, menuStoreCol.药品id) = .TextMatrix(.Row, menuPriceCol.药品id) Then
                                vsfStore.TextMatrix(intRow, menuStoreCol.现零售价) = zlStr.FormatEx(dblSalePrice, mintPriceDigit, , True)
'                                vsfStore.TextMatrix(intRow, menuStoreCol.售价盈亏) = Format(Val(vsfStore.TextMatrix(intRow, menuStoreCol.数量)) * (Val(vsfStore.TextMatrix(intRow, menuStoreCol.现零售价)) - Val(vsfStore.TextMatrix(intRow, menuStoreCol.原零售价))), mstrMoneyFormat)
                                vsfStore.TextMatrix(intRow, menuStoreCol.售价盈亏) = Format(Format(Val(vsfStore.TextMatrix(intRow, menuStoreCol.数量)) * Val(vsfStore.TextMatrix(intRow, menuStoreCol.现零售价)), mstrMoneyFormat) - Format(Val(vsfStore.TextMatrix(intRow, menuStoreCol.数量)) * Val(vsfStore.TextMatrix(intRow, menuStoreCol.原零售价)), mstrMoneyFormat), mstrMoneyFormat)

                                If Val(vsfStore.TextMatrix(intRow, menuStoreCol.现成本价)) <> 0 Then
                                    dbl加成率 = zlStr.FormatEx(zlStr.FormatEx(((Val(vsfStore.TextMatrix(intRow, menuStoreCol.现零售价))) / Val(vsfStore.TextMatrix(intRow, menuStoreCol.现成本价)) - 1), 5, , True) * 100, 5, , True)
                                Else
                                    dbl加成率 = 0
                                End If
                                vsfStore.TextMatrix(intRow, menuStoreCol.加成率) = dbl加成率
                            End If
                        Next
                    End If
                End If

                Call CaculateCost(lngDrugID, .EditText) '重新计算成本价
            Case menuPriceCol.现零售价
                If Val(.EditText) < 0 Then
                    MsgBox "售价不能为负数！", vbExclamation, gstrSysName
                    Cancel = True
                End If
                If Not IsNumeric(.EditText) Then
                    Cancel = True
                    Exit Sub
                End If

                If .EditText > 9999999 Then
                    MsgBox "零售价过大，请重新输入！", vbInformation, gstrSysName
                    Cancel = True
                    Exit Sub
                End If

                .EditText = zlStr.FormatEx(.EditText, mintPriceDigit, , True)
'                If mdblOldPrice = .EditText Then '未做修改直接退出
'                    Exit Sub
'                End If

                If mbln现价提示 = True Then
                    If Val(.EditText) > Val(.TextMatrix(Row, menuPriceCol.原指导售价)) Then
                        If MsgBox("现零售价高于指导售价" & Val(.TextMatrix(.Row, menuPriceCol.原指导售价)) & "。" & vbCrLf & "继续吗？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                            Cancel = True
                            Exit Sub
                        Else
                            .TextMatrix(.Row, menuPriceCol.现指导售价) = zlStr.FormatEx(.EditText, mintPriceDigit, , True)
                        End If
                    End If
                Else
                    If Val(.EditText) > Val(.TextMatrix(Row, menuPriceCol.原指导售价)) Then
                        .TextMatrix(.Row, menuPriceCol.现指导售价) = zlStr.FormatEx(.EditText, mintPriceDigit, , True)
                    End If
                End If
                
                If chkAotuCost.Value = 1 Then '修改售价后自动计算成本价
                    .TextMatrix(.Row, menuPriceCol.现成本价) = zlStr.FormatEx(.EditText / (1 + Val(.TextMatrix(.Row, menuPriceCol.加成率))), mintCostDigit, , True)
                    If vsfStore.rows > 1 Then
                        For intRow = 1 To vsfStore.rows - 1
                            If vsfStore.TextMatrix(intRow, menuStoreCol.药品id) = .TextMatrix(.Row, menuPriceCol.药品id) Then
                                vsfStore.TextMatrix(intRow, menuStoreCol.现成本价) = zlStr.FormatEx(.TextMatrix(.Row, menuPriceCol.现成本价), mintCostDigit, , True)
                                
                                If Val(vsfStore.TextMatrix(intRow, menuStoreCol.现成本价)) <> 0 Then
                                    dbl加成率 = zlStr.FormatEx((.EditText / Val(vsfStore.TextMatrix(intRow, menuStoreCol.现成本价)) - 1), 5, , True)
                                Else
                                    dbl加成率 = 0
                                End If
                                vsfStore.TextMatrix(intRow, menuStoreCol.加成率) = zlStr.FormatEx(dbl加成率 * 100, 5, , True)
'                                vsfStore.TextMatrix(intRow, menuStoreCol.成本盈亏) = Format((Val(vsfStore.TextMatrix(intRow, menuStoreCol.现成本价)) - Val(vsfStore.TextMatrix(intRow, menuStoreCol.原成本价))) * Val(vsfStore.TextMatrix(intRow, menuStoreCol.数量)), mstrMoneyFormat)
                                vsfStore.TextMatrix(intRow, menuStoreCol.成本盈亏) = Format(Format(Val(vsfStore.TextMatrix(intRow, menuStoreCol.现成本价)) * Val(vsfStore.TextMatrix(intRow, menuStoreCol.数量)), mstrMoneyFormat) - Format(Val(vsfStore.TextMatrix(intRow, menuStoreCol.原成本价)) * Val(vsfStore.TextMatrix(intRow, menuStoreCol.数量)), mstrMoneyFormat), mstrMoneyFormat)
                           
                            End If
                        Next
                    End If
                End If

                Call ChangeDrugStore(Row, lngDrugID, .EditText)
        End Select
    End With
End Sub

Private Sub ChangeDrugStore(ByVal intRow As Integer, ByVal lngDrugID As Long, ByVal dblNewPrice As Double)
    '功能：通过修改价格表中的零售价修改库存列表中相对应的零售价
    Dim dblOldPrice As Double
    Dim dblOldCost As Double
    Dim dblNewCost As Double
    Dim dblNum As Double
    Dim dbl包装 As Double
    Dim n As Integer
    Dim dbl发票金额 As Double
    Dim dbl加成率 As Double

    If intRow = 0 Or mint调价 = 1 Then Exit Sub

    dbl包装 = Val(vsfPrice.TextMatrix(vsfPrice.Row, menuPriceCol.包装系数))

    With vsfStore
        For n = 1 To .rows - 1
            If .TextMatrix(n, 0) <> "" Then
                If Val(.TextMatrix(n, menuStoreCol.药品id)) = lngDrugID Then
                    dblNum = Val(.TextMatrix(n, menuStoreCol.数量))
                    dblOldPrice = Val(vsfStore.TextMatrix(n, menuStoreCol.原零售价))

                    .TextMatrix(n, menuStoreCol.现零售价) = zlStr.FormatEx(dblNewPrice, mintPriceDigit, , True)
'                    .TextMatrix(n, menuStoreCol.售价盈亏) = Format(Val(.TextMatrix(n, menuStoreCol.数量)) * (dblNewPrice - dblOldPrice), mstrMoneyFormat)
                    .TextMatrix(n, menuStoreCol.售价盈亏) = Format(Format(Val(.TextMatrix(n, menuStoreCol.数量)) * dblNewPrice, mstrMoneyFormat) - Format(Val(.TextMatrix(n, menuStoreCol.数量)) * dblOldPrice, mstrMoneyFormat), mstrMoneyFormat)

                    If Val(.TextMatrix(n, menuStoreCol.现成本价)) <> 0 Then
                        dbl加成率 = zlStr.FormatEx(((Val(.TextMatrix(n, menuStoreCol.现零售价))) / Val(.TextMatrix(n, menuStoreCol.现成本价)) - 1), 5, , True)
                    Else
                        dbl加成率 = 0
                    End If
                    .TextMatrix(n, menuStoreCol.加成率) = zlStr.FormatEx(dbl加成率 * 100, 5, , True)
                    
                    If mint调价 = 2 And chkAotuCost.Value = 1 Then
                        dblOldCost = .TextMatrix(n, menuStoreCol.原成本价)
                        dblNewCost = dblNewPrice / (1 + Round(Val(.TextMatrix(n, menuStoreCol.加成率)) / 100, 7))
                        .TextMatrix(n, menuStoreCol.现成本价) = zlStr.FormatEx(dblNewCost, mintCostDigit, , True)
'                        .TextMatrix(n, menuStoreCol.成本盈亏) = Format((dblNewCost - dblOldCost) * dblNum, mstrMoneyFormat)
                        .TextMatrix(n, menuStoreCol.成本盈亏) = Format(Format(.TextMatrix(n, menuStoreCol.现成本价) * dblNum, mstrMoneyFormat) - Format(dblOldCost * dblNum, mstrMoneyFormat), mstrMoneyFormat)

                    End If
                    dbl发票金额 = dbl发票金额 + Val(.TextMatrix(n, menuStoreCol.成本盈亏))
                End If
            End If
        Next
    End With

    If chkAutoPay.Value = 1 Then
        With vsfPay
            For n = 1 To .rows - 1
                If .TextMatrix(1, 0) <> "" Then
                    If Val(.TextMatrix(n, menuPayCol.药品id)) = lngDrugID Then
                        .TextMatrix(n, menuPayCol.发票金额) = zlStr.FormatEx(dbl发票金额, mintMoneyDigit, , True)
                    End If
                End If
            Next
        End With
    End If

    If mint调价 = 2 Then
        CaluateAverCost lngDrugID
    End If
End Sub

Private Sub CaluateAverCost(ByVal lng药品id As Long)
    '计算平均成本价
    Dim i As Integer
    Dim dblSumCost As Double
    Dim dblSumNumber As Double

    With vsfStore
        For i = 1 To .rows - 1
            If .TextMatrix(i, menuStoreCol.药品id) <> "" Then
                If Val(.TextMatrix(i, menuStoreCol.药品id)) = lng药品id Then
                    dblSumCost = dblSumCost + Val(.TextMatrix(i, menuStoreCol.现成本价)) * Val(.TextMatrix(i, menuStoreCol.数量))
                    dblSumNumber = dblSumNumber + Val(.TextMatrix(i, menuStoreCol.数量))
                End If
            End If
        Next
    End With

    With vsfPrice
        If dblSumNumber > 0 Then
            For i = 1 To .rows - 1
                If .TextMatrix(i, menuPriceCol.药品id) <> "" Then
                    If Val(.TextMatrix(i, menuPriceCol.药品id)) = lng药品id Then
                        .TextMatrix(i, menuPriceCol.现成本价) = zlStr.FormatEx(dblSumCost / dblSumNumber, mintCostDigit, , True)
                        Exit For
                    End If
                End If
            Next
        End If
    End With
End Sub

Private Sub CaluateAverOldCost(ByVal lng药品id As Long)
    '计算原始平均成本价
    Dim i As Integer
    Dim dblSumCost As Double
    Dim dblSumNumber As Double

    With vsfStore
        For i = 1 To .rows - 1
            If .TextMatrix(i, menuStoreCol.药品id) <> "" Then
                If Val(.TextMatrix(i, menuStoreCol.药品id)) = lng药品id Then
                    dblSumCost = dblSumCost + Val(.TextMatrix(i, menuStoreCol.原成本价)) * Val(.TextMatrix(i, menuStoreCol.数量))
                    dblSumNumber = dblSumNumber + Val(.TextMatrix(i, menuStoreCol.数量))
                End If
            End If
        Next
    End With

    With vsfPrice
        If dblSumNumber > 0 Then
            For i = 1 To .rows - 1
                If .TextMatrix(i, menuPriceCol.药品id) <> "" Then
                    If Val(.TextMatrix(i, menuPriceCol.药品id)) = lng药品id Then
                        .TextMatrix(i, menuPriceCol.原成本价) = zlStr.FormatEx(dblSumCost / dblSumNumber, mintCostDigit, , True)
                        Exit For
                    End If
                End If
            Next
        End If
    End With
End Sub

Private Sub CaculateCost(ByVal lng药品id As Long, ByVal dbl现成本价 As Double)
    '功能：通过修改价格表中的成本价修改库存列表中相对应的成本价

    Dim n As Integer
    Dim dbl发票金额 As Double

    With vsfStore
        For n = 1 To .rows - 1
            If .TextMatrix(n, menuStoreCol.药品id) <> "" Then
                If Val(.TextMatrix(n, menuStoreCol.药品id)) = lng药品id Then
                    .TextMatrix(n, menuStoreCol.现成本价) = zlStr.FormatEx(dbl现成本价, mintCostDigit, , True)
                    If (cbo售价计算方式.Text = "售价按分段加成计算" Or cbo售价计算方式.Text = "售价按固定比例计算") And vsfPrice.TextMatrix(vsfPrice.Row, menuPriceCol.是否变价) = "1" And mint调价 = 2 Then
                        .TextMatrix(n, menuStoreCol.现零售价) = vsfPrice.TextMatrix(vsfPrice.Row, menuPriceCol.现零售价)
                    End If
                    If dbl现成本价 <> 0 Then
                        .TextMatrix(n, menuStoreCol.加成率) = zlStr.FormatEx(zlStr.FormatEx((Val(.TextMatrix(n, menuStoreCol.现零售价)) / dbl现成本价 - 1), 5, , True) * 100, 5, , True)
                    End If
                    If cbo售价计算方式 = "售价按分段加成计算" Then
                        .TextMatrix(n, menuStoreCol.加成率) = zlStr.FormatEx(zlStr.FormatEx(mdbl分段加成率, 5, , True) * 100, 5, , True)
                    End If
'                    .TextMatrix(n, menuStoreCol.成本盈亏) = Format((dbl现成本价 - Val(.TextMatrix(n, menuStoreCol.原成本价))) * Val(.TextMatrix(n, menuStoreCol.数量)), mstrMoneyFormat)
                    .TextMatrix(n, menuStoreCol.成本盈亏) = Format(Format(dbl现成本价 * Val(.TextMatrix(n, menuStoreCol.数量)), mstrMoneyFormat) - Format(Val(.TextMatrix(n, menuStoreCol.原成本价)) * Val(.TextMatrix(n, menuStoreCol.数量)), mstrMoneyFormat), mstrMoneyFormat)
                    dbl发票金额 = dbl发票金额 + (dbl现成本价 - .TextMatrix(n, menuStoreCol.原成本价)) * Val(.TextMatrix(n, menuStoreCol.数量))
                    .TextMatrix(n, menuStoreCol.售价盈亏) = Format(Format(Val(.TextMatrix(n, menuStoreCol.现零售价)) * Val(.TextMatrix(n, menuStoreCol.数量)), mstrMoneyFormat) - Format(Val(.TextMatrix(n, menuStoreCol.原零售价)) * Val(.TextMatrix(n, menuStoreCol.数量)), mstrMoneyFormat), mstrMoneyFormat)
                
                End If
            End If
        Next
    End With

    If chkAutoPay.Value = 1 Then
        For n = 1 To vsfPay.rows - 1
            If vsfPay.TextMatrix(1, 0) <> "" Then
                If Val(vsfPay.TextMatrix(n, menuPayCol.药品id)) = lng药品id Then
                    vsfPay.TextMatrix(n, menuPayCol.发票金额) = Format(dbl发票金额, mstrMoneyFormat)
                End If
            End If
        Next
    End If
End Sub


Private Sub vsfStore_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsfStore
        .Move 0, 360, TabCtlDetails.Width, TabCtlDetails.Height - 370
    End With
End Sub

Private Sub vsfStore_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfStore
        If .Cell(flexcpBackColor, Row, Col, Row, Col) = mconlngColor Then
            Cancel = True
            .Editable = flexEDNone
        Else
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub

Private Sub setColHiddenVsf()
    '不同模式下面，列显示不一样
    With vsfStore
        If cboPriceMethod.Text = "仅调售价" Then
            .colHidden(menuStoreCol.批次) = True
            .colHidden(menuStoreCol.变价) = True
            .colHidden(menuStoreCol.加成率) = True
            .colHidden(menuStoreCol.原成本价) = True
            .colHidden(menuStoreCol.现成本价) = False
            .colHidden(menuStoreCol.成本盈亏) = True
            .colHidden(menuStoreCol.原零售价) = False
            .colHidden(menuStoreCol.现零售价) = False
        ElseIf cboPriceMethod.Text = "仅调成本价" Then
            .colHidden(menuStoreCol.原零售价) = True
            .colHidden(menuStoreCol.现零售价) = False
            .colHidden(menuStoreCol.售价盈亏) = True
            .colHidden(menuStoreCol.加成率) = False
            .colHidden(menuStoreCol.原成本价) = False
            .colHidden(menuStoreCol.现成本价) = False
            .colHidden(menuStoreCol.成本盈亏) = False
        ElseIf cboPriceMethod.Text = "售价成本价一起调价" Then
            .colHidden(menuStoreCol.原零售价) = False
            .colHidden(menuStoreCol.现零售价) = False
            .colHidden(menuStoreCol.售价盈亏) = False
            .colHidden(menuStoreCol.加成率) = False
            .colHidden(menuStoreCol.原成本价) = False
            .colHidden(menuStoreCol.现成本价) = False
            .colHidden(menuStoreCol.成本盈亏) = False
        End If
    End With
End Sub

Private Sub vsfStore_Click()
    Dim i As Integer
    With vsfStore
        For i = 1 To vsfPrice.rows - 1
            If Val(.TextMatrix(.Row, menuStoreCol.药品id)) = Val(vsfPrice.TextMatrix(i, menuPriceCol.药品id)) Then
                vsfPrice.Tag = i
            End If
        Next
    End With
End Sub

Private Sub vsfStore_DblClick()
    With vsfStore
        If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mconlngCanColColor Then
            .EditCell
            .EditSelStart = 0
            .EditSelLength = Len(.EditText)
        End If
    End With
End Sub

Private Sub vsfStore_EnterCell()
    With vsfStore
        If .CellBackColor = mconlngColor Then
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
        End If
        If .Col = menuStoreCol.加成率 Then
            mdblOldPrice = Val(.TextMatrix(.Row, menuStoreCol.加成率))
        ElseIf .Col = menuStoreCol.现成本价 Then
            mdblOldPrice = Val(.TextMatrix(.Row, menuStoreCol.现成本价))
        ElseIf .Col = menuStoreCol.现零售价 Then
            mdblOldPrice = Val(.TextMatrix(.Row, menuStoreCol.现零售价))
        End If
    End With
End Sub

Private Sub vsfStore_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsfStore
        If KeyCode = vbKeyReturn Then
            If .Col < vsfStore.Cols - 1 Then
                .Col = .Col + 1
            Else
                If .Row <> .rows - 1 Then
                    .Row = .Row + 1
                    .Col = menuStoreCol.规格
                End If
            End If
        End If
    End With
End Sub

Private Sub vsfStore_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        With vsfStore
            If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mconlngCanColColor Then
                .Editable = flexEDKbdMouse
            Else
                .Editable = flexEDNone
            End If
        End With
    End If
End Sub

Private Sub vsfStore_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strkey As String
    Dim intDigit As Integer

    If KeyAscii = vbKeyReturn Then Exit Sub
    If KeyAscii <> vbKeyBack Then
        With vsfStore
            If Col = menuStoreCol.现成本价 Or Col = menuStoreCol.现零售价 Or Col = menuStoreCol.加成率 Then
                strkey = .EditText
                Select Case Col
                    Case menuStoreCol.现成本价
                        intDigit = mintCostDigit
                    Case menuStoreCol.现零售价
                        intDigit = mintPriceDigit
                    Case menuStoreCol.加成率
                        intDigit = 5
                End Select
                If KeyAscii = vbKeyDelete Then
                    If InStr(1, .EditText, ".") > 0 Then
                        KeyAscii = 0
                    End If
                ElseIf KeyAscii = Asc(".") Or (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
                    If .EditSelLength = Len(strkey) Then Exit Sub
                    If InStr(strkey, ".") <> 0 And Chr(KeyAscii) = "." Then   '只能存在一个小数点
                        KeyAscii = 0
                        Exit Sub
                    End If
                    If Len(Mid(strkey, InStr(1, strkey, ".") + 1)) >= intDigit And strkey Like "*.*" Then
                        KeyAscii = 0
                        Exit Sub
                    Else
                        Exit Sub
                    End If
                Else
                    KeyAscii = 0
                End If
            End If
        End With
    End If
End Sub

Private Sub vsfStore_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strInput As String
    Dim n As Integer
    Dim intRow As Integer
    Dim dbl发票金额 As Double
    Dim Dbl数量 As Double
    Dim Dbl金额 As Double
    Dim Dbl成本金额 As Double
    Dim dbl现采购价 As Double
    Dim dblTempNum As Double

    With vsfStore
        If .EditText = "" Then Exit Sub
        intRow = .Row
        Select Case .Col
            Case menuStoreCol.现零售价
                If Not IsNumeric(.EditText) Then
                    MsgBox "请输入新的售价。", vbInformation, gstrSysName
                    Exit Sub
                Else
                    .EditText = zlStr.FormatEx(.EditText, mintPriceDigit, , True)
                End If

                If .EditText > 9999999 Then
                    MsgBox "零售价过大，请重新输入！", vbInformation, gstrSysName
                    Cancel = True
                    Exit Sub
                End If

'                If mdblOldPrice = .EditText Then Exit Sub

                If chkAotuCost.Value = 1 Then '修改售价后自动计算成本价
                    .TextMatrix(intRow, menuStoreCol.现成本价) = zlStr.FormatEx(.EditText / (1 + Val(.TextMatrix(intRow, menuStoreCol.加成率)) / 100), mintCostDigit, , True)
                    .TextMatrix(intRow, menuStoreCol.成本盈亏) = Format(Format(Val(.TextMatrix(intRow, menuStoreCol.数量)) * Val(.TextMatrix(intRow, menuStoreCol.现成本价)), mstrMoneyFormat) - Format(Val(.TextMatrix(intRow, menuStoreCol.数量)) * Val(.TextMatrix(intRow, menuStoreCol.原成本价)), mstrMoneyFormat), mstrMoneyFormat)
                End If
                
                .TextMatrix(intRow, menuStoreCol.售价盈亏) = Format(Format(Val(.TextMatrix(intRow, menuStoreCol.数量)) * Val(.EditText), mstrMoneyFormat) - Format(Val(.TextMatrix(intRow, menuStoreCol.数量)) * Val(.TextMatrix(intRow, menuStoreCol.原零售价)), mstrMoneyFormat), mstrMoneyFormat)

'                .TextMatrix(intRow, menuStoreCol.售价盈亏) = Format(Val(.TextMatrix(intRow, menuStoreCol.数量)) * (Val(.EditText) - Val(.TextMatrix(intRow, menuStoreCol.原零售价))), mstrMoneyFormat)
                .TextMatrix(intRow, menuStoreCol.现零售价) = zlStr.FormatEx(Val(.EditText), mintPriceDigit, , True)
'                .TextMatrix(intRow, menuStoreCol.现成本价) = zlStr.FormatEx(Val(.TextMatrix(intRow, menuStoreCol.现零售价)) / (1 + Val(.TextMatrix(intRow, menuStoreCol.加成率)) / 100), mintCostDigit)
'                .TextMatrix(intRow, menuStoreCol.成本盈亏) = Format((Val(.TextMatrix(intRow, menuStoreCol.现成本价)) - Val(.TextMatrix(intRow, menuStoreCol.原成本价))) * Val(.TextMatrix(intRow, menuStoreCol.数量)), mstrMoneyFormat)
                If chkAotuCost.Value <> 1 Then
                    If Val(.TextMatrix(intRow, menuStoreCol.现成本价)) <> 0 Then
                        .TextMatrix(intRow, menuStoreCol.加成率) = zlStr.FormatEx(zlStr.FormatEx((Val(.TextMatrix(intRow, menuStoreCol.现零售价)) / Val(.TextMatrix(intRow, menuStoreCol.现成本价)) - 1), 5, , True) * 100, 5, , True)
                    Else
                        .TextMatrix(intRow, menuStoreCol.加成率) = zlStr.FormatEx(0, 5, , True)
                    End If
                End If
                
                For n = 1 To .rows - 1
                    If .TextMatrix(intRow, menuStoreCol.药品id) = .TextMatrix(n, menuStoreCol.药品id) Then
                        If Val(.TextMatrix(intRow, menuStoreCol.批次)) <> 0 And Val(.TextMatrix(intRow, menuStoreCol.批次)) = Val(.TextMatrix(n, menuStoreCol.批次)) Then
                            .TextMatrix(n, menuStoreCol.现零售价) = .TextMatrix(intRow, menuStoreCol.现零售价)
'                            .TextMatrix(n, menuStoreCol.售价盈亏) = Format(Val(.TextMatrix(n, menuStoreCol.数量)) * (Val(.EditText) - Val(.TextMatrix(n, menuStoreCol.原零售价))), mstrMoneyFormat)
                            .TextMatrix(n, menuStoreCol.售价盈亏) = Format(Format(Val(.TextMatrix(n, menuStoreCol.数量)) * Val(.EditText), mstrMoneyFormat) - Format(Val(.TextMatrix(n, menuStoreCol.数量)) * Val(.TextMatrix(n, menuStoreCol.原零售价)), mstrMoneyFormat), mstrMoneyFormat)
                            If chkAotuCost.Value <> 1 Then
                                If Val(.TextMatrix(n, menuStoreCol.现成本价)) <> 0 Then
                                    .TextMatrix(n, menuStoreCol.加成率) = zlStr.FormatEx(zlStr.FormatEx((Val(.TextMatrix(n, menuStoreCol.现零售价)) / Val(.TextMatrix(n, menuStoreCol.现成本价)) - 1), 5, , True) * 100, 5, , True)
                                Else
                                    .TextMatrix(n, menuStoreCol.加成率) = zlStr.FormatEx(0, 5, , True)
                                End If
                            End If
                        End If
                        Dbl数量 = Dbl数量 + .TextMatrix(n, menuStoreCol.数量)
                        Dbl金额 = Dbl金额 + .TextMatrix(n, menuStoreCol.数量) * Val(.TextMatrix(n, menuStoreCol.现零售价))
                        Dbl成本金额 = Dbl成本金额 + .TextMatrix(n, menuStoreCol.数量) * Val(.TextMatrix(n, menuStoreCol.现成本价))
                    End If
                Next
                For n = 1 To vsfPrice.rows - 1
                    If .TextMatrix(intRow, menuStoreCol.药品id) = vsfPrice.TextMatrix(n, menuPriceCol.药品id) Then
                        If Dbl数量 <> 0 Then
                            If chkAotuCost.Value = 1 Then
                                vsfPrice.TextMatrix(n, menuPriceCol.现成本价) = zlStr.FormatEx(Dbl成本金额 / Dbl数量, mintPriceDigit, , True)
                            End If
                            vsfPrice.TextMatrix(n, menuPriceCol.现零售价) = zlStr.FormatEx(Dbl金额 / Dbl数量, mintPriceDigit, , True)
                        Else
                            If chkAotuCost.Value = 1 Then
                                vsfPrice.TextMatrix(n, menuPriceCol.现成本价) = vsfStore.TextMatrix(intRow, menuStoreCol.现成本价)
                            End If
                            vsfPrice.TextMatrix(n, menuPriceCol.现零售价) = vsfStore.TextMatrix(intRow, menuStoreCol.现零售价)
                        End If
                    End If
                Next

                If mint调价 > 0 Then
                    For n = 1 To .rows - 1
                        If .TextMatrix(n, menuStoreCol.药品id) <> "" Then
                            If Val(.TextMatrix(n, menuStoreCol.药品id)) = Val(.TextMatrix(intRow, menuStoreCol.药品id)) Then
                                dbl发票金额 = dbl发票金额 + (Val(.TextMatrix(n, menuStoreCol.现成本价)) - Val(.TextMatrix(n, menuStoreCol.原成本价))) * Val(.TextMatrix(n, menuStoreCol.数量))
                            End If
                        End If
                    Next

                    If chkAutoPay.Value = 1 Then
                        For n = 1 To vsfPay.rows - 1
                            If vsfPay.TextMatrix(1, 0) <> "" Then
                                If Val(vsfPay.TextMatrix(n, menuPayCol.药品id)) = Val(vsfStore.TextMatrix(intRow, menuStoreCol.药品id)) Then
                                    vsfPay.TextMatrix(n, menuPayCol.发票金额) = zlStr.FormatEx(dbl发票金额, mintMoneyDigit, , True)
                                End If
                            End If
                        Next
                    End If
                End If
            Case menuStoreCol.加成率
                If Val(.EditText) < 0 Then Exit Sub
                If Not IsNumeric(.EditText) Then
                    Cancel = True
                    Exit Sub
                End If
'                If mdblOldPrice = .EditText Then Exit Sub
                .EditText = zlStr.FormatEx(.EditText, 5, , True)
                .TextMatrix(intRow, menuStoreCol.加成率) = zlStr.FormatEx(Val(.EditText), 5, , True)
                .TextMatrix(intRow, menuStoreCol.现零售价) = zlStr.FormatEx(Val(.TextMatrix(intRow, menuStoreCol.现成本价)) * (1 + Val(.TextMatrix(intRow, menuStoreCol.加成率)) / 100), mintCostDigit, , True)
'                .TextMatrix(intRow, menuStoreCol.售价盈亏) = Format(Val(.TextMatrix(intRow, menuStoreCol.数量)) * (Val(.TextMatrix(intRow, menuStoreCol.现零售价)) - Val(.TextMatrix(intRow, menuStoreCol.原零售价))), mstrMoneyFormat)
                .TextMatrix(intRow, menuStoreCol.售价盈亏) = Format(Format(Val(.TextMatrix(intRow, menuStoreCol.数量)) * Val(.TextMatrix(intRow, menuStoreCol.现零售价)), mstrMoneyFormat) - Format(Val(.TextMatrix(intRow, menuStoreCol.数量)) * Val(.TextMatrix(intRow, menuStoreCol.原零售价)), mstrMoneyFormat), mstrMoneyFormat)

                For n = 1 To .rows - 1
                    If vsfPrice.TextMatrix(Val(vsfPrice.Tag), menuPriceCol.药品id) = .TextMatrix(n, menuStoreCol.药品id) Then
                        If Val(.TextMatrix(intRow, menuStoreCol.变价)) = 0 Or mbln时价药品按批次调价 = False Then
                            .TextMatrix(n, menuStoreCol.加成率) = zlStr.FormatEx(Val(.EditText), 5, , True)
                            .TextMatrix(n, menuStoreCol.现零售价) = zlStr.FormatEx(Val(.TextMatrix(n, menuStoreCol.现成本价)) * (1 + zlStr.FormatEx(Val(.EditText), 5) / 100), mintCostDigit, , True)
'                            .TextMatrix(n, menuStoreCol.售价盈亏) = Format(Val(.TextMatrix(n, menuStoreCol.数量)) * (Val(.TextMatrix(n, menuStoreCol.现零售价)) - Val(.TextMatrix(n, menuStoreCol.原零售价))), mstrMoneyFormat)
                            .TextMatrix(n, menuStoreCol.售价盈亏) = Format(Format(Val(.TextMatrix(n, menuStoreCol.数量)) * Val(.TextMatrix(n, menuStoreCol.现零售价)), mstrMoneyFormat) - Format(Val(.TextMatrix(n, menuStoreCol.数量)) * Val(.TextMatrix(n, menuStoreCol.原零售价)), mstrMoneyFormat), mstrMoneyFormat)
    
                        End If
                        Dbl数量 = Dbl数量 + .TextMatrix(n, menuStoreCol.数量)
                        Dbl金额 = Dbl金额 + .TextMatrix(n, menuStoreCol.数量) * Val(.TextMatrix(n, menuStoreCol.现零售价))
                    End If
                Next
                If Dbl数量 <> 0 Then
                    vsfPrice.TextMatrix(Val(vsfPrice.Tag), menuPriceCol.现零售价) = zlStr.FormatEx(Dbl金额 / Dbl数量, mintPriceDigit, , True)
                Else
                    vsfPrice.TextMatrix(Val(vsfPrice.Tag), menuPriceCol.现零售价) = .TextMatrix(intRow, menuStoreCol.现零售价)
                End If
            Case menuStoreCol.现成本价
                If Val(.EditText) > Val(.TextMatrix(.Row, menuStoreCol.现零售价)) Then
                    MsgBox "注意，新成本价大于了新售价！", vbExclamation, gstrSysName
                End If

                If Val(.EditText) < 0 Then
                    MsgBox "成本价不能为负数！", vbExclamation, gstrSysName
                    Cancel = True
                End If
                If .EditText > 9999999 Then
                    MsgBox "采购价过大，请重新输入！", vbInformation, gstrSysName
                    Cancel = True
                    Exit Sub
                End If
'                If mdblOldPrice = .EditText Then Exit Sub
                .EditText = zlStr.FormatEx(.EditText, mintCostDigit, , True)
                .TextMatrix(intRow, menuStoreCol.现成本价) = zlStr.FormatEx(Val(.EditText), mintCostDigit, , True)
'                If Val(.EditText) <> 0 Then
'                    .TextMatrix(intRow, menuStoreCol.加成率) = zlStr.FormatEx((Val(.TextMatrix(intRow, menuStoreCol.现零售价)) / Val(.EditText) - 1) * 100, 5)
'                End If
'                .TextMatrix(intRow, menuStoreCol.成本盈亏) = Format((Val(.EditText) - .TextMatrix(intRow, menuStoreCol.原成本价)) * Val(.TextMatrix(intRow, menuStoreCol.数量)), mstrMoneyFormat)
                .TextMatrix(intRow, menuStoreCol.成本盈亏) = Format(Format(Val(.EditText) * Val(.TextMatrix(intRow, menuStoreCol.数量)), mstrMoneyFormat) - Format(.TextMatrix(intRow, menuStoreCol.原成本价) * Val(.TextMatrix(intRow, menuStoreCol.数量)), mstrMoneyFormat), mstrMoneyFormat)
                
                If Val(.TextMatrix(intRow, menuStoreCol.变价)) = 1 And mbln时价药品按批次调价 = True And mint调价 <> 1 Then
                    .TextMatrix(intRow, menuStoreCol.现零售价) = zlStr.FormatEx(zlStr.FormatEx(Val(.EditText), mintCostDigit) * (1 + (Val(.TextMatrix(intRow, menuStoreCol.加成率)) / 100)), mintPriceDigit, , True)
'                    .TextMatrix(intRow, menuStoreCol.售价盈亏) = Format(Val(.TextMatrix(intRow, menuStoreCol.数量)) * (Val(.TextMatrix(intRow, menuStoreCol.现零售价)) - Val(.TextMatrix(intRow, menuStoreCol.原零售价))), mstrMoneyFormat)
                    .TextMatrix(intRow, menuStoreCol.售价盈亏) = Format(Format(Val(.TextMatrix(intRow, menuStoreCol.数量)) * Val(.TextMatrix(intRow, menuStoreCol.现零售价)), mstrMoneyFormat) - Format(Val(.TextMatrix(intRow, menuStoreCol.数量)) * Val(.TextMatrix(intRow, menuStoreCol.原零售价)), mstrMoneyFormat), mstrMoneyFormat)

                End If
                
                dbl发票金额 = (Val(.EditText) - .TextMatrix(intRow, menuStoreCol.原成本价)) * Val(.TextMatrix(intRow, menuStoreCol.数量))

                For n = 1 To .rows - 1
                    If .TextMatrix(n, menuStoreCol.药品id) <> "" Then
                        If Val(.TextMatrix(n, menuStoreCol.药品id)) = Val(.TextMatrix(intRow, menuStoreCol.药品id)) And n <> intRow Then
                            If mbln成本价按库房批次调整 = False Or (Val(.TextMatrix(intRow, menuStoreCol.批次)) <> 0 And Val(.TextMatrix(intRow, menuStoreCol.批次)) = Val(.TextMatrix(n, menuStoreCol.批次))) Then
                                dbl现采购价 = Val(.EditText)
                                .TextMatrix(n, menuStoreCol.现成本价) = zlStr.FormatEx(dbl现采购价, mintCostDigit, , True)
'                                If dbl现采购价 <> 0 Then
'                                    .TextMatrix(n, menuStoreCol.加成率) = zlStr.FormatEx((Val(.TextMatrix(n, menuStoreCol.现零售价)) / dbl现采购价 - 1) * 100, 5)
'                                End If
'                                .TextMatrix(n, menuStoreCol.成本盈亏) = Format((dbl现采购价 - .TextMatrix(n, menuStoreCol.原成本价)) * Val(.TextMatrix(n, menuStoreCol.数量)), mstrMoneyFormat)
                                 .TextMatrix(n, menuStoreCol.成本盈亏) = Format(Format(dbl现采购价 * Val(.TextMatrix(n, menuStoreCol.数量)), mstrMoneyFormat) - Format(.TextMatrix(n, menuStoreCol.原成本价) * Val(.TextMatrix(n, menuStoreCol.数量)), mstrMoneyFormat), mstrMoneyFormat)
                               
                                If Val(.TextMatrix(intRow, menuStoreCol.变价)) = 1 And mbln时价药品按批次调价 = True And mint调价 <> 1 Then
                                    .TextMatrix(n, menuStoreCol.现零售价) = zlStr.FormatEx(zlStr.FormatEx(dbl现采购价, mintCostDigit) * (1 + (Val(.TextMatrix(n, menuStoreCol.加成率)) / 100)), mintPriceDigit, , True)
'                                    .TextMatrix(n, menuStoreCol.售价盈亏) = Format(Val(.TextMatrix(n, menuStoreCol.数量)) * (Val(.TextMatrix(n, menuStoreCol.现零售价)) - Val(.TextMatrix(n, menuStoreCol.原零售价))), mstrMoneyFormat)
                                    .TextMatrix(n, menuStoreCol.售价盈亏) = Format(Format(Val(.TextMatrix(n, menuStoreCol.数量)) * Val(.TextMatrix(n, menuStoreCol.现零售价)), mstrMoneyFormat) - Format(Val(.TextMatrix(n, menuStoreCol.数量)) * Val(.TextMatrix(n, menuStoreCol.原零售价)), mstrMoneyFormat), mstrMoneyFormat)

                                End If
                            Else
                                dbl现采购价 = Val(.TextMatrix(n, menuStoreCol.现成本价))
                            End If
                            dbl发票金额 = dbl发票金额 + (dbl现采购价 - .TextMatrix(n, menuStoreCol.原成本价)) * Val(.TextMatrix(n, menuStoreCol.数量))
                        End If
                    End If
                Next

                If chkAutoPay.Value = 1 Then
                    For n = 1 To vsfPay.rows - 1
                        If vsfPay.TextMatrix(1, 0) <> "" Then
                            If Val(vsfPay.TextMatrix(n, menuPayCol.药品id)) = Val(vsfStore.TextMatrix(intRow, menuStoreCol.药品id)) Then
                                vsfPay.TextMatrix(n, menuPayCol.发票金额) = Format(dbl发票金额, mstrMoneyFormat)
                            End If
                        End If
                    Next
                End If

                If mbln成本价按库房批次调整 = False Then
                    For n = 1 To vsfPrice.rows - 1
                        If Val(.TextMatrix(intRow, menuStoreCol.药品id)) = Val(vsfPrice.TextMatrix(n, menuPriceCol.药品id)) Then
                            vsfPrice.TextMatrix(n, menuPriceCol.现成本价) = .TextMatrix(intRow, menuStoreCol.现成本价)
                            Exit For
                        End If
                    Next
                Else
                    CaluateAverCost Val(.TextMatrix(intRow, menuStoreCol.药品id))
                End If
                Call CaculateAverPirce(Val(.TextMatrix(intRow, menuStoreCol.药品id)))  '价格变动，计算平均售价
        End Select
    End With
End Sub

Private Sub CaculateAverPirce(ByVal lng药品id As Long)
    '自动计算平均售价
    Dim i As Integer
    Dim dblSumPrice As Double
    Dim dblSumNumber As Double
    
    With vsfStore
        For i = 1 To .rows - 1
            If .TextMatrix(i, menuStoreCol.药品id) <> "" Then
                If Val(.TextMatrix(i, menuStoreCol.药品id)) = lng药品id Then
                    dblSumPrice = dblSumPrice + Val(.TextMatrix(i, menuStoreCol.现零售价)) * Val(.TextMatrix(i, menuStoreCol.数量))
                    dblSumNumber = dblSumNumber + Val(.TextMatrix(i, menuStoreCol.数量))
                End If
            End If
        Next
    End With

    With vsfPrice
        If dblSumNumber > 0 Then
            For i = 1 To .rows - 1
                If .TextMatrix(i, menuPriceCol.药品id) <> "" Then
                    If Val(.TextMatrix(i, menuPriceCol.药品id)) = lng药品id Then
                        .TextMatrix(i, menuPriceCol.现零售价) = zlStr.FormatEx(dblSumPrice / dblSumNumber, mintPriceDigit, , True)
                        Exit For
                    End If
                End If
            Next
        End If
    End With
End Sub

Private Sub CaculateAverOldPirce(ByVal lng药品id As Long)
    '自动原始计算平均售价
    Dim i As Integer
    Dim dblSumPrice As Double
    Dim dblSumNumber As Double
    
    With vsfStore
        For i = 1 To .rows - 1
            If .TextMatrix(i, menuStoreCol.药品id) <> "" Then
                If Val(.TextMatrix(i, menuStoreCol.药品id)) = lng药品id Then
                    dblSumPrice = dblSumPrice + Val(.TextMatrix(i, menuStoreCol.原零售价)) * Val(.TextMatrix(i, menuStoreCol.数量))
                    dblSumNumber = dblSumNumber + Val(.TextMatrix(i, menuStoreCol.数量))
                End If
            End If
        Next
    End With

    With vsfPrice
        If dblSumNumber > 0 Then
            For i = 1 To .rows - 1
                If .TextMatrix(i, menuPriceCol.药品id) <> "" Then
                    If Val(.TextMatrix(i, menuPriceCol.药品id)) = lng药品id Then
                        .TextMatrix(i, menuPriceCol.原零售价) = zlStr.FormatEx(dblSumPrice / dblSumNumber, mintPriceDigit, , True)
                        Exit For
                    End If
                End If
            Next
        End If
    End With
End Sub

Private Sub initCommandBars()
    Dim cbrToolBar As CommandBar
    Dim cbrControl As CommandBarControl
    Dim cbrControlPopu As CommandBarControl
    Dim lngCount As Integer
    
    With CommandBarsGlobalSettings
        .App = App
        .CompanyName = "重庆中联信息产业有限责任公司" '公司名称
        .ResourceFile = .OcxPath & "\XTPResourceZhCn.dll" '设置中文语言资源文件
        .ColorManager.SystemTheme = xtpSystemThemeAuto  '控件整体的颜色方案
    End With

    With cbsMain.Options
        .ShowExpandButtonAlways = False '总是在工具栏右侧显示选项按钮,即使窗体宽度足够。
        .ToolBarAccelTips = True '显示按钮提示
        .AlwaysShowFullMenus = False '不常用的菜单项先隐藏
        .UseFadedIcons = True '图标显示为褪色效果
        .IconsWithShadow = True '鼠标指向的命令图标显示阴影效果
        .UseDisabledIcons = True '工具栏按钮禁用时图标显示为禁用样式
        .LargeIcons = True '工具栏显示为大图标
        .SetIconSize True, 24, 24 '设置大图标的尺寸
        .SetIconSize False, 16, 16 '设置小图标的尺寸
    End With

    With cbsMain
        .VisualTheme = xtpThemeOffice2003 '设置控件显示风格
        .EnableCustomization False '是否允许自定义设置
        Set .Icons = imgList.Icons '设置关联的图标控件
        .ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap '窗体变化时，如果显示不完菜单也不换行
        .ActiveMenuBar.Title = "菜单"
    End With
    
    '删除现在的工具栏及顶级菜单项
    For lngCount = cbsMain.ActiveMenuBar.Controls.count To 1 Step -1
        cbsMain.ActiveMenuBar.Controls(lngCount).Delete
    Next
    For lngCount = cbsMain.count To 1 Step -1
        cbsMain(lngCount).Delete
    Next
    
    '创建工具栏
    Set cbrToolBar = cbsMain.Add("工具栏", xtpBarTop)
    cbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    cbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    cbrToolBar.ContextMenuPresent = False

    With cbrToolBar
        Set cbrControl = .Controls.Add(xtpControlButton, mconMenu_PrintStore, "打印库存变动表")
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Controls.Add(xtpControlButton, mconMenu_ClearAll, "清空")
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Controls.Add(xtpControlButton, mconMenu_Find, "查找")
        cbrControl.Visible = False
    
        Set cbrControl = .Controls.Add(xtpControlButton, mconMenu_BatchSelect, "批量选择项目")
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Controls.Add(xtpControlButton, mconMenu_Save, "确定")
        cbrControl.BeginGroup = True
        Set cbrControl = .Controls.Add(xtpControlButton, mconMenu_Quit, "退出")
                
    End With

    For Each cbrControl In cbrToolBar.Controls  '让工具栏中按钮同时显示图标和文字
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    With Me.cbsMain.KeyBindings
        .Add 0, VK_F3, mconMenu_Find
    End With

End Sub
