VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmStuffPriceCard 
   Caption         =   "卫材调价单"
   ClientHeight    =   8550
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12810
   Icon            =   "frmStuffPriceCard.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   12810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picSplit 
      BorderStyle     =   0  'None
      Height          =   100
      Left            =   240
      MousePointer    =   7  'Size N S
      ScaleHeight     =   105
      ScaleWidth      =   2775
      TabIndex        =   45
      Top             =   4200
      Width           =   2775
   End
   Begin VB.TextBox txtFind 
      Height          =   300
      Left            =   840
      TabIndex        =   41
      Top             =   7440
      Width           =   1965
   End
   Begin VB.PictureBox picOtherSelect 
      Height          =   3135
      Left            =   3600
      ScaleHeight     =   3075
      ScaleWidth      =   4755
      TabIndex        =   25
      Top             =   1200
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CommandButton cmdFilterOk 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   2400
         TabIndex        =   38
         Top             =   2640
         Width           =   1100
      End
      Begin VB.CommandButton cmdFilterCan 
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   3480
         TabIndex        =   37
         Top             =   2640
         Width           =   1100
      End
      Begin VB.Frame fra辅助选项 
         Caption         =   "辅助选项（成本价调价相关）"
         Height          =   2535
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   4695
         Begin VB.CheckBox chk加成率 
            Caption         =   "指定加成率"
            Height          =   180
            Left            =   120
            TabIndex        =   32
            Top             =   1125
            Width           =   1215
         End
         Begin VB.CheckBox chk供应商 
            Caption         =   "指定供应商"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox chk应付记录 
            Caption         =   "产生成本价调价带来的应付款修正记录"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   1920
            Width           =   3495
         End
         Begin VB.TextBox txt加成率 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   270
            Left            =   1440
            TabIndex        =   29
            Text            =   "15.0000"
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox txt供应商 
            Enabled         =   0   'False
            Height          =   270
            Left            =   1440
            TabIndex        =   28
            Top             =   360
            Width           =   2655
         End
         Begin VB.CommandButton cmd供应商 
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   270
            Left            =   4080
            TabIndex        =   27
            Top             =   350
            Width           =   375
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshProvider 
            Height          =   1695
            Left            =   120
            TabIndex        =   33
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
            TabIndex        =   36
            Top             =   1440
            Width           =   4260
         End
         Begin VB.Label lblComment供应商 
            AutoSize        =   -1  'True
            Caption         =   "（指定供应商，则只调整该供应商的库存卫材成本价）"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   240
            TabIndex        =   35
            Top             =   720
            Width           =   4320
         End
         Begin VB.Label lblPercent 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   180
            Left            =   2415
            TabIndex        =   34
            Top             =   1125
            Width           =   90
         End
      End
   End
   Begin VB.PictureBox picInfo 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   -120
      ScaleHeight     =   495
      ScaleWidth      =   10575
      TabIndex        =   20
      Top             =   6600
      Width           =   10575
      Begin VB.TextBox txtSummary 
         Height          =   300
         Left            =   4320
         MaxLength       =   100
         TabIndex        =   23
         Top             =   120
         Width           =   5565
      End
      Begin VB.TextBox txtValuer 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   120
         Width           =   1965
      End
      Begin VB.Label lblSummary 
         AutoSize        =   -1  'True
         Caption         =   "调价说明"
         Height          =   180
         Left            =   3360
         TabIndex        =   24
         Top             =   180
         Width           =   720
      End
      Begin VB.Label lblValuer 
         AutoSize        =   -1  'True
         Caption         =   "调价人"
         Height          =   180
         Left            =   360
         TabIndex        =   22
         Top             =   180
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "清空(&D)"
      Height          =   350
      Left            =   6720
      TabIndex        =   14
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   12360
      TabIndex        =   13
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton cmdCanc 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   13680
      TabIndex        =   12
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印库存变动表(&P)…"
      Height          =   350
      Left            =   9960
      TabIndex        =   11
      Top             =   7440
      Width           =   1935
   End
   Begin VB.CommandButton cmdItem 
      Caption         =   "批量选择项目(&I)"
      Height          =   350
      Left            =   8160
      TabIndex        =   10
      Top             =   7440
      Width           =   1695
   End
   Begin VB.Frame fraCondition 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   16335
      Begin VB.CheckBox chkAppAllColumn 
         Caption         =   "修改价格应用于所有列"
         Height          =   255
         Left            =   11040
         TabIndex        =   48
         Top             =   23
         Width           =   2295
      End
      Begin VB.CheckBox chkAutoPay 
         Caption         =   "自动计算应付款变动记录"
         Height          =   210
         Left            =   4560
         TabIndex        =   40
         Top             =   480
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CheckBox chkCostBatch 
         Caption         =   "成本价按库房批次调整"
         Height          =   210
         Left            =   2160
         TabIndex        =   39
         Top             =   480
         Width           =   2370
      End
      Begin VB.CheckBox Chk定价 
         Caption         =   "时价卫材改为定价"
         Height          =   210
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   1770
      End
      Begin VB.CommandButton cmdPriceMethod 
         Caption         =   "…"
         Height          =   300
         Left            =   3360
         TabIndex        =   16
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox cboPriceMethod 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   0
         Width           =   2415
      End
      Begin VB.CheckBox chk按批次 
         Caption         =   "成本价按库房批次调整"
         Height          =   210
         Left            =   10560
         TabIndex        =   7
         Top             =   -225
         Width           =   2175
      End
      Begin VB.CheckBox chk自动计算应付款变动 
         Caption         =   "自动计算应付款变动"
         Height          =   210
         Left            =   12840
         TabIndex        =   6
         Top             =   -225
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.OptionButton opt时间 
         Caption         =   "立即执行"
         Height          =   255
         Index           =   0
         Left            =   5040
         TabIndex        =   5
         Top             =   8
         Width           =   1095
      End
      Begin VB.OptionButton opt时间 
         Caption         =   "指定日期执行"
         Height          =   255
         Index           =   1
         Left            =   6240
         TabIndex        =   4
         Top             =   8
         Width           =   1455
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
         Left            =   8040
         TabIndex        =   8
         Top             =   0
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
         Format          =   125829123
         CurrentDate     =   36846.5833333333
      End
      Begin VB.Label lblMethod 
         AutoSize        =   -1  'True
         Caption         =   "调价方式"
         Height          =   180
         Left            =   120
         TabIndex        =   19
         Top             =   60
         Width           =   720
      End
      Begin VB.Label lbl执行时间 
         Caption         =   "执行时间"
         Height          =   180
         Left            =   4200
         TabIndex        =   9
         Top             =   45
         Width           =   855
      End
   End
   Begin VB.TextBox txtNO 
      Enabled         =   0   'False
      Height          =   300
      Left            =   13200
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin XtremeSuiteControls.TabControl TabCtlDetails 
      Height          =   975
      Left            =   240
      TabIndex        =   17
      Top             =   5040
      Width           =   1815
      _Version        =   589884
      _ExtentX        =   3201
      _ExtentY        =   1720
      _StockProps     =   64
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfStore 
      Height          =   975
      Left            =   2880
      TabIndex        =   43
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
   Begin VSFlex8Ctl.VSFlexGrid vsfPay 
      Height          =   975
      Left            =   8040
      TabIndex        =   44
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
      TabIndex        =   46
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
      TabIndex        =   47
      Top             =   8190
      Width           =   12810
      _ExtentX        =   22595
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16828
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
   Begin VB.Label lblFind 
      Caption         =   "查找"
      Height          =   255
      Left            =   240
      TabIndex        =   42
      Top             =   7488
      Width           =   495
   End
   Begin VB.Label lblNO 
      AutoSize        =   -1  'True
      Caption         =   "调价流水号"
      Height          =   180
      Left            =   12120
      TabIndex        =   1
      Top             =   180
      Width           =   900
   End
   Begin VB.Label lblDrugName 
      AutoSize        =   -1  'True
      Caption         =   "卫材调价单"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6120
      TabIndex        =   0
      Top             =   120
      Width           =   1875
   End
End
Attribute VB_Name = "frmStuffPriceCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'各种全局变量
Private Const mconlngRowHeight As Long = 300 '表格中各行行高
Private mintUnit As Integer     '用来记录启用的是什么单位
Private mint调价 As Integer     '0-调售价;1-调成本价;2-调售价及成本价
Private mlng供应商ID As Long  '用来记录供应商id
Private mdbl加成率 As Double
Private mbln应付记录 As Boolean '记录是否产生应付记录
Private mbln时价卫材按批次调价 As Boolean '时价卫材按照批次调价
Private mint参数 As Integer  'mint参数=1代表从卫材目录进入
Private mlng规格ID As Long '从卫材目录界面获取的规格ID
Private mstr调整额 As String
Private mintSalePriceDigit As Integer
'颜色方案
Private Const mconlngColor As Long = &HFFFFFF        '不能修改列颜色为白色
Private Const mconlngCanColor As Long = &HE7CFBA    '能修改列颜色为淡蓝色

Private mbln现价提示 As Boolean         '限价卫材提示 true-提示 false-不提示
Private mdbl分段加成率 As Double    '用来记录分段加成率
Private mdbl成本价 As Double            '记录修改之前的成本价
Private mstrNo As String            '调价单No
Private mintModal As Integer        '本次是什么状态 0-新增 1-修改 2-查阅
Private mintMethod As Integer   '调价方式 0-调售价;1-调成本价;2-调售价及成本价
Private mstr调价汇总号 As String
Private mblnLoad As Boolean     '是否加载完成
Private mrsReturn As ADODB.Recordset '批量选择返回的数据集
Private mblnOk As Boolean
Private mrsFindName As ADODB.Recordset '查询的数据集
Private mblnClick As Boolean
Private mintType As Integer      '调整方式
Private mdbl比率 As Double      '调整方式中填写的调整额度
Private mlngPrice As Long       '记录价格
Private mblnUpdateAdd As Boolean    '修改情况下的新增卫材
Private mlngOldStuffID As Long '检查原始行是否有药品
Private mdblOldPrice As Double     '记录原始价格
Private mblnBatchItem As Boolean   '记录是否点击了批量选择按钮
Private mstrPrivs As String       '模块权限
Private Const mstrCaption As String = "卫材调价单"

Private mFMT As g_FmtString
Private mOraFMT As g_FmtString

Private Enum menuPriceCol
    材料ID = 0
    原价id = 1
    品名 = 2
    规格 = 3
    是否变价
    厂牌
    单位
    包装系数
    是否跟踪在用
    加成率
    差价让利比
    是否有库存
    收入项目id
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
    材料ID = 0
    库房 = 1
    库房ID = 2
    供应商
    供应商ID
    药品
    规格
    批号
    效期
    产地
    批次
    变价
    数量
    单位
    包装系数
    原零售价
    现零售价
    调整金额
    加成率
    原采购价
    现采购价
    差价差
    总列数
End Enum

Private Enum menuPayCol
    材料ID = 0
    品名 = 1
    供应商
    供应商ID
    规格
    产地
    发票号
    发票日期
    发票金额
    总列数
End Enum

Public Sub ShowMe(ByVal frmParent As Form, ByVal intModal As Integer, ByVal str调价汇总号 As String, ByVal intMethod As Integer, Optional int参数 As Integer, Optional lng规格ID As Long)
    mintModal = intModal
    mstr调价汇总号 = str调价汇总号
    mintMethod = intMethod
    mstrPrivs = GetPrivFunc(glngSys, 1726)
    mint参数 = int参数
    mlng规格ID = lng规格ID
    
    Me.Show vbModal, frmParent
End Sub

Private Sub cboPriceMethod_Click()
    Dim intCol As Integer
    Dim intTemp As Integer

    With cboPriceMethod
        If .Text = "仅调售价" Then
            intTemp = 0
        ElseIf .Text = "仅调成本价" Then
            intTemp = 1
        Else
            intTemp = 2
        End If
    End With

    If mint参数 = 1 Then
        If mblnLoad = True And intTemp <> Val(lblMethod.Tag) Then
            If vsfPrice.TextMatrix(1, menuPriceCol.材料ID) <> "" Then
                If MsgBox("调价方式改变将恢复列表已修改的价格，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    cboPriceMethod.ListIndex = mint调价
                    Exit Sub
                Else
                    mdbl比率 = 0
                    mstr调整额 = ""
                    vsfPrice.Rows = 2
                    For intCol = 0 To vsfPrice.Cols - 1
                        vsfPrice.TextMatrix(1, intCol) = ""
                    Next
                    vsfStore.Rows = 1
                    vsfPay.Rows = 1
                    Call CatalogModifyPrice
                End If
            End If
        End If
    Else
        If mblnLoad = True And intTemp <> Val(lblMethod.Tag) Then
            If vsfPrice.TextMatrix(1, menuPriceCol.材料ID) <> "" Then
                If MsgBox("调价方式改变将清空列表中数据，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    cboPriceMethod.ListIndex = mint调价
                    Exit Sub
                Else
                    mdbl比率 = 0
                    mstr调整额 = ""
                    vsfPrice.Rows = 2
                    For intCol = 0 To vsfPrice.Cols - 1
                        vsfPrice.TextMatrix(1, intCol) = ""
                    Next
                    vsfStore.Rows = 1
                    vsfPay.Rows = 1
                End If
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
            chkCostBatch.Visible = False
            chkCostBatch.Value = False
            chkAutoPay.Visible = False
            chkAutoPay.Value = 0
            TabCtlDetails.Item(1).Visible = False
        ElseIf .Text = "仅调成本价" Then
            mint调价 = 1
            lblMethod.Tag = 1
            opt时间(0).Value = True
            opt时间(0).Enabled = False
            opt时间(1).Enabled = False
            dtpRunDate.Enabled = False
            chkCostBatch.Visible = True
            If mbln应付记录 = True Then
                chkAutoPay.Visible = True
                chkAutoPay.Value = 1
                TabCtlDetails.Item(1).Visible = True
            End If
        ElseIf .Text = "售价成本价一起调价" Then
            mint调价 = 2
            lblMethod.Tag = 2
            opt时间(0).Value = False
            opt时间(1).Value = True
            opt时间(0).Enabled = True
            opt时间(1).Enabled = True
            dtpRunDate.Enabled = True
            chkCostBatch.Visible = True
            If mbln应付记录 = True Then
                chkAutoPay.Visible = True
                chkAutoPay.Value = 1
                TabCtlDetails.Item(1).Visible = True
            End If
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

Private Sub Chk供应商_Click()
    If chk供应商.Value = 1 Then
        Cmd供应商.Enabled = True
        txt供应商.Enabled = True
        chk应付记录.Enabled = True
    Else
        Cmd供应商.Enabled = False
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

Private Sub cmdCanc_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    Dim intCol As Integer

    If MsgBox("你确定要清空所有数据？", vbYesNo, gstrSysName) = vbYes Then
        mdbl比率 = 0
        mstr调整额 = ""
        vsfPrice.Rows = 2
        For intCol = 0 To vsfPrice.Cols - 1
            vsfPrice.TextMatrix(1, intCol) = ""
        Next
        vsfStore.Rows = 1
        vsfPay.Rows = 1
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
        If Val(.TextMatrix(1, menuPriceCol.材料ID)) <> 0 Then
            If MsgBox("将清空表格中的数据，是否继续？", vbYesNo, gstrSysName) = vbNo Then
                Exit Sub
            Else
                vsfPrice.Rows = 2
                For i = 0 To vsfPrice.Cols - 1
                    .TextMatrix(1, i) = ""
                Next
                vsfStore.Rows = 1
                vsfPay.Rows = 1
            End If
        End If
    End With

    mlng供应商ID = IIf(chk供应商.Value = 1, Val(Split(txt供应商.Tag, "|")(0)), 0)
    mdbl加成率 = IIf(chk加成率.Value = 1, Val(Trim(txt加成率.Text)), 0)
    mbln应付记录 = (chk应付记录.Enabled And chk应付记录.Value = 1)
    picOtherSelect.Visible = False
    If mbln应付记录 = True Then
        TabCtlDetails.Item(1).Visible = True
        chkAutoPay.Visible = True
        chkAutoPay.Value = 1
    Else
        TabCtlDetails.Item(1).Visible = False
        chkAutoPay.Visible = False
        chkAutoPay.Value = 0
    End If
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Sub cmdItem_Click()
    Dim intRow As Integer

    frmBatchSelect.ShowMe Me, mrsReturn, mblnOk, mintType, mdbl比率, mint调价, mstr调整额

    On Error GoTo ErrHandle
    If mblnOk = False Then Exit Sub
    If mrsReturn.RecordCount = 0 Then Exit Sub

    With vsfPrice
        If .TextMatrix(.Rows - 1, menuPriceCol.材料ID) = "" Then
            intRow = .Rows - 1
        Else
            .Rows = .Rows + 1
            intRow = .Rows - 1
        End If
    End With
    mblnBatchItem = True
    Call GetDrugPirce(mrsReturn, intRow)
    mblnBatchItem = False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub deleteNotExecutePirce()
    '清除未执行价格
    Dim intRow As Integer

    On Error GoTo ErrHandle
    With vsfPrice
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, menuPriceCol.材料ID) <> "" Then
                gstrSQL = "Zl_删除材料未执行价格_Delete(" & Val(.TextMatrix(intRow, menuPriceCol.材料ID)) & "," & 0 & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
            End If
        Next
    End With

    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdOk_Click()
    Dim intRow As Integer
    Dim intCol As Integer
    Dim dtToDay As Date
    Dim lngAdjId As Long
    Dim lngId As Long
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
    Dim lng材料ID As Long
    Dim lng批次  As Long
    Dim str批号 As String
    Dim str效期 As String
    Dim str产地 As String
    Dim dblOldCost As Double
    Dim dblNewCost As Double
    Dim str发票号 As String
    Dim str发票日期 As String
    Dim dbl发票金额 As Double

    Dim lng序号 As Long
    Dim cllProc As Collection
    Dim strTemp As String
    Dim j As Integer
    Dim dbl成本价 As Double

    Set cllProc = New Collection

    If vsfPrice.Rows > 1 Then
        If Val(vsfPrice.TextMatrix(1, menuPriceCol.材料ID)) = 0 Then Exit Sub
    End If
    If CheckPrice = False Then Exit Sub

    On Error GoTo ErrHand
    dtToDay = sys.Currentdate()

    gstrSQL = "select 收费价目_ID.nextval from dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取收费价目序号")
    lngAdjId = rsTemp.Fields(0).Value

    If mintModal = 1 Then '修改 在修改模式下先删除原来的调价信息，然后插入新的调价信息
        Call deleteNotExecutePirce
    End If

    '检查是否存在未执行的价格
    If checkNotExecutePrice = True Then Exit Sub
    '获取调价NO
    mstrNo = sys.GetNextNo(9)
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
        lng序号 = 1
        For intCount = 1 To IIf(Trim(.TextMatrix(.Rows - 1, 0)) = "", .Rows - 2, .Rows - 1)
            lng材料ID = Val(.TextMatrix(intCount, menuPriceCol.材料ID))
            dbl包装 = Val(.TextMatrix(intCount, menuPriceCol.包装系数))
            
            If lng材料ID <> 0 Then
                If Val(.TextMatrix(intCount, menuPriceCol.现零售价)) <> Val(.TextMatrix(intCount, menuPriceCol.原零售价)) Then
                    lngId = sys.NextId("收费价目")
                    If opt时间(0).Value = True Then
                        strID = strID & "," & lngId
                    ElseIf lng材料ID = -1 Then
                        strID = strID & "," & lngId
                    End If
                    
                    If .TextMatrix(intCount, menuPriceCol.是否变价) = "1" And mbln时价卫材按批次调价 And mint调价 <> 1 Then
                        strTmp = ""
                        lngCurrBatch = -1
                        For n = 1 To vsfStore.Rows - 1
                            If Val(.TextMatrix(intCount, menuPriceCol.材料ID)) = Val(vsfStore.TextMatrix(n, menuStoreCol.材料ID)) Then
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
                        gstrSQL = "zl_收费价目_stop("
                                    '    收费细目ID_IN IN 收费价目.收费细目ID%TYPE,
                                    gstrSQL = gstrSQL & "" & lng材料ID & ","
                                    '    终止日期_IN IN 收费价目.终止日期%TYPE := NULL
                                    If opt时间(0).Value Then
                                        gstrSQL = gstrSQL & "to_date('" & Format(DateAdd("s", -1, dtToDay), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                                    Else
                                        gstrSQL = gstrSQL & "to_date('" & Format(DateAdd("s", -1, Me.dtpRunDate.Value), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                                    End If
                                    gstrSQL = gstrSQL & ")"
                                    AddArray cllProc, gstrSQL

                                    'Zl_收费价目_Insert
                                    gstrSQL = "zl_收费价目_Insert("
                                    '  Id_In         In 收费价目.ID%Type,
                                    gstrSQL = gstrSQL & "" & lngId & ","
                                    '  原价id_In     In 收费价目.原价id%Type := Null,
                                    gstrSQL = gstrSQL & "" & IIf(.TextMatrix(intCount, menuPriceCol.原价id) = "", "NUll", Val(.TextMatrix(intCount, menuPriceCol.原价id))) & ","
                                    '  收费细目id_In In 收费价目.收费细目id%Type := Null,
                                    gstrSQL = gstrSQL & "" & lng材料ID & ","
                                    '  收入项目id_In In 收费价目.收入项目id%Type := Null,
                                    gstrSQL = gstrSQL & "" & Val(.TextMatrix(intCount, menuPriceCol.收入项目id)) & ","
                                    '  原价_In       In 收费价目.原价%Type := Null,
                                    If .TextMatrix(intCount, menuPriceCol.是否变价) = "1" And Val(.TextMatrix(intCount, menuPriceCol.是否跟踪在用)) = 0 Then
                                        '非跟踪卫生材料的实价卫材，是以范围决定的，（主要是医嘱应用),始终填为零
                                        gstrSQL = gstrSQL & "" & 0 & ","
                                    Else
                                        gstrSQL = gstrSQL & "" & Round(Val(.TextMatrix(intCount, menuPriceCol.原零售价)) / dbl包装, g_小数位数.obj_最大小数.零售价小数) & ","
                                    End If

                                    '  现价_In       In 收费价目.现价%Type := Null,
                                    gstrSQL = gstrSQL & "" & Round(Val(.TextMatrix(intCount, menuPriceCol.现零售价)) / dbl包装, g_小数位数.obj_最大小数.零售价小数) & ","
                                    '  附术收费率_In In 收费价目.附术收费率%Type := Null,
                                    gstrSQL = gstrSQL & "NULL,"
                                    '  加班加价率_In In 收费价目.加班加价率%Type := Null,
                                    gstrSQL = gstrSQL & "NULL,"
                                    '  调价说明_In   In 收费价目.调价说明%Type := Null,
                                    gstrSQL = gstrSQL & "'" & Me.txtSummary.Text & "',"
                                    '  调价id_In     In 收费价目.调价id%Type := Null,
                                    gstrSQL = gstrSQL & "" & lngAdjId & ","
                                    '  调价人_In     In 收费价目.调价人%Type := Null,
                                    gstrSQL = gstrSQL & "'" & Me.txtValuer.Text & "',"
                                    '  执行日期_In   In 收费价目.执行日期%Type := Null,
                                    If Me.opt时间(0).Value Then
                                        gstrSQL = gstrSQL & "to_date('" & Format(dtToDay, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
                                    Else
                                        gstrSQL = gstrSQL & "to_date('" & Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
                                    End If
                                    '  变动原因_In   In 收费价目.变动原因%Type := 1,
                                    gstrSQL = gstrSQL & "" & 0 & ","
                                    '  No_In         In 收费价目.NO%Type := Null,
                                    gstrSQL = gstrSQL & "'" & mstrNo & "',"
                                    '  序号_In       In 收费价目.序号%Type := 1
                                    gstrSQL = gstrSQL & "" & lng序号 & ","
                                    '缺省价格_In
                                    If .TextMatrix(intCount, menuPriceCol.是否变价) = "1" And Val(.TextMatrix(intCount, menuPriceCol.是否跟踪在用)) = 0 Then
                                            gstrSQL = gstrSQL & "" & Round(Val(.TextMatrix(intCount, menuPriceCol.现零售价)) / dbl包装, g_小数位数.obj_最大小数.零售价小数) & ","
                                    Else
                                            gstrSQL = gstrSQL & "NULL,"
                                    End If
                                    '调价汇总号
                                    gstrSQL = gstrSQL & "" & txtNO.Text & ")"
                                    AddArray cllProc, gstrSQL
                                    lng序号 = lng序号 + 1
                        blnPrice = True
                        blnPrint = True
                    End If
                End If
                If lng材料ID <> 0 Then
                    If Val(.TextMatrix(intCount, menuPriceCol.原指导售价)) <> Val(.TextMatrix(intCount, menuPriceCol.现指导售价)) Then
                        strTemp = Round(Val(.TextMatrix(intCount, menuPriceCol.现指导售价)) / dbl包装, g_小数位数.obj_最大小数.零售价小数)
                        'zl_材料特性_UpdateCustom ( 材料ID_IN ,SQL_IN)
                        gstrSQL = "zl_材料特性_UpdateCustom(" & lng材料ID & ",'指导零售价=" & strTemp & "')"
                        AddArray cllProc, gstrSQL
                    End If
                    '更新采购限价
                    If Val(.TextMatrix(intCount, menuPriceCol.原采购限价)) <> Val(.TextMatrix(intCount, menuPriceCol.现采购限价)) Then
                        strTemp = Round(Val(.TextMatrix(intCount, menuPriceCol.现采购限价)) / dbl包装, g_小数位数.obj_最大小数.成本价小数)
                        'zl_材料特性_UpdateCustom ( 材料ID_IN ,SQL_IN)
                        gstrSQL = "zl_材料特性_UpdateCustom(" & lng材料ID & ",'指导批发价=" & strTemp & "')"
                        AddArray cllProc, gstrSQL
                    End If
                End If
            End If
        Next
    End With

    '成本价调价处理
    If mint调价 = 1 Or mint调价 = 2 Then
        With vsfStore
            For i = 1 To .Rows - 1
                lng库房ID = Val(.TextMatrix(i, menuStoreCol.库房ID))
                lng供应商ID = Val(.TextMatrix(i, menuStoreCol.供应商ID))
                lng材料ID = Val(.TextMatrix(i, menuStoreCol.材料ID))
                lng批次 = Val(.TextMatrix(i, menuStoreCol.批次))
                str批号 = .TextMatrix(i, menuStoreCol.批号)
                dbl包装 = Val(.TextMatrix(i, menuStoreCol.包装系数))
                If lng材料ID <> 0 Then
                    str发票号 = "": str发票日期 = "": dbl发票金额 = 0
                    If chkAutoPay.Value = 1 Then
                        With vsfPay
                            For j = 1 To .Rows - 1
                                If Val(.TextMatrix(j, menuPayCol.材料ID)) = lng材料ID And _
                                    Val(.TextMatrix(j, menuPayCol.供应商ID)) = lng供应商ID Then
                                    '看是否有此卫生材料库存变动情况
                                    str发票号 = Trim(.TextMatrix(j, menuPayCol.发票号))
                                    str发票日期 = Trim(.TextMatrix(j, menuPayCol.发票日期))
                                    dbl发票金额 = Val(.TextMatrix(j, menuPayCol.发票金额))
                                    Exit For
                                End If
                            Next
                        End With
                    End If

                    dbl成本价 = Round(Val(.TextMatrix(i, menuStoreCol.现采购价)) / dbl包装, g_小数位数.obj_最大小数.成本价小数)

                    ' Zl_材料成本调价_Insert
                    gstrSQL = "Zl_材料成本调价_Insert("
                    '  供药单位id_In In 成本价调价信息.供药单位id%Type,
                    gstrSQL = gstrSQL & IIf(lng供应商ID = 0, "Null", lng供应商ID) & ","
                    '  库房id_In     In 成本价调价信息.库房id%Type,
                    gstrSQL = gstrSQL & "" & lng库房ID & ","
                    '  材料id_In     In 成本价调价信息.药品id%Type,
                    gstrSQL = gstrSQL & "" & lng材料ID & ","
                    '  批次_In       In 成本价调价信息.批次%Type := Null,
                    gstrSQL = gstrSQL & "" & lng批次 & ","
                    '批号_in
                    gstrSQL = gstrSQL & "" & IIf(str批号 = "", "NULL", "'" & str批号 & "'") & ","
                    '  原成本价_In   In 成本价调价信息.原成本价%Type := Null,
                    gstrSQL = gstrSQL & "" & Round(Val(.TextMatrix(i, menuStoreCol.原采购价)) / dbl包装, g_小数位数.obj_最大小数.成本价小数) & ","
                    '  新成本价_In   In 成本价调价信息.新成本价%Type := Null,
                    gstrSQL = gstrSQL & "" & dbl成本价 & ","
                    '  发票号_In     In 成本价调价信息.发票号%Type := Null,
                    gstrSQL = gstrSQL & "'" & str发票号 & "',"
                    '  发票日期_In   In 成本价调价信息.发票日期%Type := Null,
                    gstrSQL = gstrSQL & "" & IIf(str发票日期 = "", "NULL", "to_date('" & str发票日期 & "','yyyy-mm-dd') ") & ","
                    '  发票金额_In   In 成本价调价信息.发票金额%Type := Null,
                    gstrSQL = gstrSQL & "" & dbl发票金额 & ","
                    '  应付款变动_In In 成本价调价信息.应付款变动%Type := 0
                    gstrSQL = gstrSQL & "" & IIf(chkAutoPay.Value = 1 And lng供应商ID <> 0 And dbl发票金额 <> 0, 1, 0) & ","
                    gstrSQL = gstrSQL & "'" & txtNO.Text & "')"
                    AddArray cllProc, gstrSQL
                    blnCost = True
                End If
            Next
        End With
    End If

    '无库存时调整成本价
    If mint调价 = 1 Or mint调价 = 2 Then
        With Me.vsfPrice
            For intCount = 1 To .Rows - 1
                lng材料ID = Val(.TextMatrix(intCount, menuPriceCol.材料ID))
                dbl包装 = Val(.TextMatrix(intCount, menuStoreCol.包装系数))
                If lng材料ID <> 0 Then
                    If .TextMatrix(intCount, menuPriceCol.是否有库存) = "0" And Val(.TextMatrix(intCount, menuPriceCol.原成本价)) <> Val(.TextMatrix(intCount, menuPriceCol.现成本价)) Then
                        dbl包装 = Val(.TextMatrix(intCount, menuPriceCol.包装系数))
    
                        lng材料ID = Val(.TextMatrix(intCount, menuPriceCol.材料ID))
                        dblOldCost = Val(Round(Val(.TextMatrix(intCount, menuPriceCol.原成本价)) / dbl包装, g_小数位数.obj_最大小数.成本价小数))
                        dblNewCost = Val(Round(Val(.TextMatrix(intCount, menuPriceCol.现成本价)) / dbl包装, g_小数位数.obj_最大小数.成本价小数))
    
                        gstrSQL = "Zl_材料成本调价_Insert(Null,Null," & lng材料ID & ",0,NULL" & "," & dblOldCost & ", " & dblNewCost & ",NULL,Null,0,0, '" & txtNO.Text & "')"
                        AddArray cllProc, gstrSQL
                        blnCost = True
                    End If
                End If
            Next
        End With
    End If

   '分两种情况下对成本价进行调整:
    '1.当仅为成本价调价及立即执行时，立即对成本价进行调整
    '2.当非立即执行和非成本价(即成本价调价方式)调价时，在卫生材料调价时，再执行。
     '单独成本价调价时
    If mint调价 = 1 Then
        If Me.opt时间(0).Value = True Then
            With vsfPrice
                For i = 1 To .Rows - 1
                    lng材料ID = Val(.TextMatrix(i, menuPriceCol.材料ID))
                    If lng材料ID <> 0 Then
                      ' Zl_材料收发记录_Adjust
                      gstrSQL = "Zl_材料收发记录_Adjust("
                      '  调价id_In In Number, --调价记录的ID
                      gstrSQL = gstrSQL & "" & 0 & ","
                      '  定价_In   In Number := 0, --是否转为定价销售（更新材料特性、收费细目中的变价）
                      gstrSQL = gstrSQL & "" & 0 & ","
                      '  材料id_In In Number := 0 --当不为0时表示是成本价调价，不处理售价相关内容
                        gstrSQL = gstrSQL & "" & lng材料ID & ")"
                      AddArray cllProc, gstrSQL
                    End If
                Next
            End With
        End If
    Else
        '调售价
        If strID <> "" Then strID = Mid(strID, 2)
        ArrayID = Split(strID, ",")
        Array批次价格 = Split(str批次价格, ";")
        For intCount = 0 To UBound(ArrayID)
            If opt时间(0).Value = True Or vsfPrice.TextMatrix(intCount + 1, menuPriceCol.原价id) = "" Then
                gstrSQL = "zl_材料收发记录_Adjust(" & ArrayID(intCount) & "," & Me.Chk定价.Value & ",0,'" & Array批次价格(intCount) & "')"
                AddArray cllProc, gstrSQL
            End If
        Next
    End If

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
    gstrSQL = gstrSQL & IIf(txtSummary.Text = "", "Null", "'" & txtSummary.Text & "'") & ",1,'" & UserInfo.用户名 & "')"

    AddArray cllProc, gstrSQL

'    gcnOracle.BeginTrans
    ExecuteProcedureArrAy cllProc, mstrCaption
'    gcnOracle.CommitTrans

    If blnPrint = True Then
        If MsgBox("你需要打印调价通知单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1726_1", Me, "调价号=" & txtNO.Text, "计算单位=" & mintUnit, 2)
        End If
    End If

    '清空列表中数据
    With vsfPrice
        .Rows = 2
        For intCol = 0 To .Cols - 1
            .TextMatrix(1, intCol) = ""
        Next
    End With
    vsfStore.Rows = 1
    vsfPay.Rows = 1
    txtNO.Text = ""
    txtSummary.Text = ""
    
    If mint参数 = 1 Then
        Unload Me
    End If
    
    Exit Sub

ErrHand:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Function CheckUnVerify(ByVal lng材料ID As Long) As Boolean
    '检查卫材是否存在未审核单据
    Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrHandle
    gstrSQL = "Select 1 From 药品收发记录 Where 材料id = [1] And Rownum = 1 And 审核日期 Is Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查卫材是否存在未审核单据", lng材料ID)

    If rsTemp.RecordCount > 0 Then
        CheckUnVerify = True
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function checkNotExecutePrice(Optional ByVal lngDrugID As Long = 0) As Boolean
    '功能 ：检查是否存在未执行的价格
    Dim RecCheck As New ADODB.Recordset
    Dim LngmediIDThis As Long, IntCheck As Integer

    err = 0
    On Error GoTo ErrHand

    If lngDrugID = 0 Then
        '循环判断所有卫材
        For IntCheck = 1 To vsfPrice.Rows - 1
            LngmediIDThis = Val(vsfPrice.TextMatrix(IntCheck, menuPriceCol.材料ID))
            If LngmediIDThis <> 0 Then
                If mint调价 = 0 Or mint调价 = 2 Then
                    '判断是否有未执行的历史价格
                    gstrSQL = " Select Count(*) Records From 收费价目 Where 变动原因=0 And 执行日期 > Sysdate And 收费细目ID=[1]" & _
                            GetPriceClassString("")
                    
                    Set RecCheck = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, LngmediIDThis)

                    With RecCheck
                        If Not .EOF Then
                            If Not IsNull(!Records) Then
                                If !Records <> 0 Then
                                    MsgBox "卫材" & vsfPrice.TextMatrix(IntCheck, menuPriceCol.品名) & "存在未执行价格，未执行卫材不能调价！", vbInformation, gstrSysName
                                    checkNotExecutePrice = True
                                    Exit Function
                                End If
                            End If
                        End If
                    End With
                End If

                If mint调价 = 1 Or mint调价 = 2 Then
                    '检查是否还有未执行的成本价调价计划
                    gstrSQL = "Select 1 From 成本价调价信息 Where 药品id = [1] And 执行日期 Is Null And Rownum = 1 "
                    Set RecCheck = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, LngmediIDThis)

                    If RecCheck.RecordCount > 0 Then
                        MsgBox "卫材" & vsfPrice.TextMatrix(IntCheck, menuPriceCol.品名) & "存在未执行成本价，未执行卫材不能调价！", vbInformation, gstrSysName
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
            
            Set RecCheck = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lngDrugID)

            With RecCheck
                If Not .EOF Then
                    If Not IsNull(!Records) Then
                        If !Records <> 0 Then
                            MsgBox "还存在未执行的售价调价记录，未执行卫材不能调价！", vbInformation, gstrSysName
                            checkNotExecutePrice = True
                            Exit Function
                        End If
                    End If
                End If
            End With
        End If

        If mint调价 = 1 Or mint调价 = 2 Then
            '检查是否还有未执行的成本价调价计划
            gstrSQL = "Select 1 From 成本价调价信息 Where 药品id = [1] And 执行日期 Is Null And Rownum = 1 "
            Set RecCheck = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lngDrugID)

            If RecCheck.RecordCount > 0 Then
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
    Dim bln无库存 As Boolean

    '检测各执行价格是否正确
    '以及收入项目相同的情况下现价是否与原价相同
    CheckPrice = False
    With vsfPrice
        For IntCheck = 1 To .Rows - 1
            If Val(.TextMatrix(IntCheck, menuPriceCol.材料ID)) <> 0 Then
                If Not IsNumeric(Trim(.TextMatrix(IntCheck, menuPriceCol.现零售价))) Then
                    MsgBox "第" & IntCheck & "行的卫材售价中含有非法字符！", vbInformation, gstrSysName
                    .Row = IntCheck
                    .Col = menuPriceCol.现零售价
                    vsfPrice.SetFocus
                    .Select IntCheck, 0, IntCheck, .Cols - 1
                    .TopRow = IntCheck
                    Exit Function
                End If

                If mint调价 <> 1 Then
                    If Val(.TextMatrix(IntCheck, menuPriceCol.现零售价)) = Val(.TextMatrix(IntCheck, menuPriceCol.原零售价)) Then
                        MsgBox "第" & IntCheck & "行的卫材现价与原价相同，不能执行调价！", vbInformation, gstrSysName
                        .Row = IntCheck
                        .Col = menuPriceCol.现零售价
                        vsfPrice.SetFocus
                        .Select IntCheck, 0, IntCheck, .Cols - 1
                        .TopRow = IntCheck
                        Exit Function
                    End If
                End If

'                If mint调价 <> 0 Then
'                    If Val(.TextMatrix(IntCheck, menuPriceCol.现成本价)) = Val(.TextMatrix(IntCheck, menuPriceCol.原成本价)) Then
'                        MsgBox "第" & IntCheck & "行的药品现成本价与原成本价相同，不能执行调价！", vbInformation, gstrSysName
'                        .Row = IntCheck
'                        .Col = menuPriceCol.现成本价
'                        vsfPrice.SetFocus
'                        .Select IntCheck, 0, IntCheck, .Cols - 1
'                        .TopRow = IntCheck
'                        Exit Function
'                    End If
'                End If

                If .TextMatrix(IntCheck, menuPriceCol.是否变价) = "1" And opt时间(0).Value <> True And mint调价 <> 1 Then
                    MsgBox "第" & IntCheck & "行为时价卫材，必须设置为立即执行！", vbInformation, gstrSysName
                    .Row = IntCheck
                    .Col = menuPriceCol.现零售价
                    vsfPrice.SetFocus
                    .Select IntCheck, 0, IntCheck, .Cols - 1
                    .TopRow = IntCheck
                    Exit Function
                End If
            End If
        Next
    End With

    CheckPrice = True
End Function


Private Sub cmdPriceMethod_Click()
    If txt供应商.Tag = "" Then
        Me.txt供应商.Tag = "0|"
    End If
    picOtherSelect.Visible = True
End Sub

Private Sub CmdPrint_Click()
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim i As Long


    If vsfStore.Rows = 1 Then Exit Sub
    If Trim(vsfStore.TextMatrix(1, menuStoreCol.材料ID)) = "" Then Exit Sub

    objPrint.Title.Text = "调价库存变动表"

    Set objRow = New zlTabAppRow
    objRow.Add "调价说明:" & Me.txtSummary.Text
    objPrint.UnderAppRows.Add objRow

    Set objRow = New zlTabAppRow
    objRow.Add "执行时间:" & Format(IIf(Me.opt时间(0).Value, sys.Currentdate, Me.dtpRunDate.Value), "yyyy年MM月DD日 HH:mm:ss")
    objRow.Add "调价人:" & Me.txtValuer.Text
    objPrint.UnderAppRows.Add objRow

    Set objRow = New zlTabAppRow
    objRow.Add "打印人:" & gstrUserName
    objRow.Add "打印时间:" & Format(sys.Currentdate, "yyyy年MM月DD日 HH:mm:ss")
    objPrint.BelowAppRows.Add objRow

    Set objPrint.Body = vsfStore
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

    On Error GoTo ErrHandle
    gstrSQL = "Select 编码,名称,简码,id" & _
        " From 供应商" & _
        " where 末级=1 And substr(类型,5,1) = '1' And (撤档时间 is null or 撤档时间=to_date('3000-01-01','YYYY-MM-DD')) " & _
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
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub Form_Activate()
    If mblnLoad = False Then
        vsfPrice.SetFocus
    End If
    If mblnClick = False Then
        vsfPrice.Row = 1
        vsfPrice.Col = menuPriceCol.品名
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

    Me.Height = 768 * 15
    Me.Width = 1024 * 15
    '获取设置的单位
    mintUnit = Val(zlDatabase.GetPara("卫材单位", glngSys, 1726, 1))
    mbln时价卫材按批次调价 = Val(zlDatabase.GetPara("时价卫材按批次调价", glngSys, 1726, 0))
    
    With mFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价)
        .FM_金额 = GetFmtString(mintUnit, g_金额)
        .FM_零售价 = GetFmtString(mintUnit, g_售价)
        .FM_数量 = GetFmtString(mintUnit, g_数量)
    End With
    With mOraFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价, True)
        .FM_金额 = GetFmtString(mintUnit, g_金额, True)
        .FM_零售价 = GetFmtString(mintUnit, g_售价, True)
        .FM_数量 = GetFmtString(mintUnit, g_数量, True)
    End With

    '初始化时间为当前时间+1天
    StrToday = Format(sys.Currentdate(), "yyyy-MM-dd hh:mm:ss")

    If mintModal = 0 Then '新增的时候最小时间设置为当前时间+1天
        Me.dtpRunDate.MinDate = DateAdd("s", 1, CDate(StrToday))
    End If
    Me.dtpRunDate.Value = DateAdd("d", 1, CDate(StrToday))

    txtValuer.Text = gstrUserName

    txtNO.Text = IIf(mintModal = 0, "", mstr调价汇总号)
    If mintModal = 0 Then
        LblNo.Visible = False
        txtNO.Visible = False
    End If

    Call InitTabControl
    Call initComboBox '初始化下拉控件

    If mintModal = 1 Then '修改
        If (InStr(1, ";" & gstrPrivs & ";", ";成本价调价;") > 0 And InStr(1, ";" & gstrPrivs & ";", ";售价调价;") = 0) Or (InStr(1, ";" & gstrPrivs & ";", ";成本价调价;") = 0 And InStr(1, ";" & gstrPrivs & ";", ";售价调价;") > 0) Then
            cboPriceMethod.ListIndex = 0
        ElseIf (InStr(1, ";" & gstrPrivs & ";", ";成本价调价;") > 0 And InStr(1, ";" & gstrPrivs & ";", ";售价调价;") > 0) Then
            cboPriceMethod.ListIndex = mintMethod
        End If
    ElseIf mintModal = 2 Then '查阅
        cboPriceMethod.ListIndex = mintMethod
    End If

    Call InitVsfGridFlex

    Call RestoreWinState(Me, App.ProductName, mstrCaption)
    If mbln应付记录 = False Then
        TabCtlDetails.Item(1).Visible = False
    End If
    If mintModal <> 0 Then
        Call initGrid
    End If

    If mintModal = 2 Then '查阅
        cboPriceMethod.Enabled = False
        cmdPriceMethod.Enabled = False
        opt时间(0).Enabled = False
        opt时间(1).Enabled = False
        dtpRunDate.Enabled = False
        Chk定价.Enabled = False
        chkCostBatch.Enabled = False
        chkAutoPay.Enabled = False
        txtSummary.Enabled = False
        cmdClear.Visible = False
        cmdItem.Visible = False
        cmdOk.Visible = False
        vsfPrice.Cell(flexcpBackColor, 1, 0, vsfPrice.Rows - 1, vsfPrice.Cols - 1) = mconlngColor
        If vsfStore.Rows > 1 Then
            vsfStore.Cell(flexcpBackColor, 1, 0, vsfStore.Rows - 1, vsfStore.Cols - 1) = mconlngColor
        End If
        If vsfPay.Rows > 1 Then
            vsfPay.Cell(flexcpBackColor, 0, 0, vsfPay.Rows - 1, vsfPay.Cols - 1) = mconlngColor
        End If
    End If
    mblnLoad = True
    If mint参数 = 1 Then
        Call CatalogModifyPrice
    End If
End Sub

Private Sub initComboBox()
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
        .InsertItem 0, "库存变动表", vsfStore.hwnd, 0
        .InsertItem 1, "应付款变动表", vsfPay.hwnd, 0
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
    txtNO.Left = Me.ScaleWidth - txtNO.Width
    LblNo.Left = txtNO.Left - LblNo.Width - 200
    lblDrugName.Left = Me.ScaleWidth / 2 - lblDrugName.Width / 2
    vsfPrice.Move 20, fraCondition.Top + fraCondition.Height + 20, Me.ScaleWidth, 3000
    picSplit.Left = 50
    picSplit.Top = vsfPrice.Top + vsfPrice.Height + 5
    picSplit.Width = Me.ScaleWidth
    txtSummary.Width = Me.ScaleWidth - lblSummary.Left - lblSummary.Width - 300
    TabCtlDetails.Move 20, picSplit.Height + picSplit.Top, Me.ScaleWidth, Me.ScaleHeight - picSplit.Top - picSplit.Height - picInfo.Height - cmdClear.Height - 300 - stbThis.Height
    picInfo.Move 0, TabCtlDetails.Top + TabCtlDetails.Height, Me.ScaleWidth
    lblFind.Top = picInfo.Top + picInfo.Height + 180
    lblFind.Left = 380
    txtFind.Top = lblFind.Top - 50
    txtFind.Left = lblFind.Left + lblFind.Width + 95
    cmdClear.Top = txtFind.Top
    cmdItem.Top = txtFind.Top
    cmdPrint.Top = txtFind.Top
    cmdOk.Top = txtFind.Top
    cmdCanc.Top = txtFind.Top
    cmdCanc.Left = Me.ScaleWidth - cmdCanc.Width - 300
    cmdOk.Left = cmdCanc.Left - cmdOk.Width - 200
    cmdPrint.Left = cmdOk.Left - cmdPrint.Width - 500
    cmdItem.Left = cmdPrint.Left - cmdPrint.Width - 20
    cmdClear.Left = cmdItem.Left - cmdItem.Width - 20
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName, mstrCaption)
    mdbl比率 = 0
    mstr调整额 = ""
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
        .Rows = 2
        .RowHeight(1) = mconlngRowHeight
        .ColWidth(0) = 200
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = 50
        .RowHeight(0) = mconlngRowHeight
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

        .TextMatrix(0, menuPriceCol.材料ID) = "材料id"
        .TextMatrix(0, menuPriceCol.原价id) = "原价id"
        .TextMatrix(0, menuPriceCol.品名) = "品名"
        .TextMatrix(0, menuPriceCol.规格) = "规格"
        .TextMatrix(0, menuPriceCol.是否变价) = "是否变价"
        .TextMatrix(0, menuPriceCol.厂牌) = "厂牌"
        .TextMatrix(0, menuPriceCol.单位) = "单位"
        .TextMatrix(0, menuPriceCol.包装系数) = "包装系数"
        .TextMatrix(0, menuPriceCol.是否跟踪在用) = "是否跟踪在用"
        .TextMatrix(0, menuPriceCol.加成率) = "加成率"
        .TextMatrix(0, menuPriceCol.差价让利比) = "差价让利比"
        .TextMatrix(0, menuPriceCol.是否有库存) = "是否有库存"
        .TextMatrix(0, menuPriceCol.收入项目id) = "收入项目id"
        .TextMatrix(0, menuPriceCol.原成本价) = "原成本价"
        .TextMatrix(0, menuPriceCol.现成本价) = "现成本价"
        .TextMatrix(0, menuPriceCol.原零售价) = "原零售价"
        .TextMatrix(0, menuPriceCol.现零售价) = "现零售价"
        .TextMatrix(0, menuPriceCol.原采购限价) = "原采购限价"
        .TextMatrix(0, menuPriceCol.现采购限价) = "现采购限价"
        .TextMatrix(0, menuPriceCol.原指导售价) = "原指导售价"
        .TextMatrix(0, menuPriceCol.现指导售价) = "现指导售价"

        '设置列宽
        .ColWidth(menuPriceCol.材料ID) = 0
        .ColWidth(menuPriceCol.原价id) = 0
        .ColWidth(menuPriceCol.品名) = 3000
        .ColWidth(menuPriceCol.规格) = 1500
        .ColWidth(menuPriceCol.是否变价) = 0
        .ColWidth(menuPriceCol.厂牌) = 2000
        .ColWidth(menuPriceCol.单位) = 800
        .ColWidth(menuPriceCol.包装系数) = 0
        .ColWidth(menuPriceCol.加成率) = 0
        .ColWidth(menuPriceCol.是否跟踪在用) = 0
        .ColWidth(menuPriceCol.差价让利比) = 0
        .ColWidth(menuPriceCol.是否有库存) = 0
        .ColWidth(menuPriceCol.收入项目id) = 0
        .ColWidth(menuPriceCol.原成本价) = 1000
        .ColWidth(menuPriceCol.现成本价) = 1000
        .ColWidth(menuPriceCol.原零售价) = 1000
        .ColWidth(menuPriceCol.现零售价) = 1000
        .ColWidth(menuPriceCol.原采购限价) = 0
        .ColWidth(menuPriceCol.现采购限价) = 0
        .ColWidth(menuPriceCol.原指导售价) = 0
        .ColWidth(menuPriceCol.现指导售价) = 0
        '设置对齐方式
        .ColAlignment(menuPriceCol.品名) = flexAlignLeftCenter
        .ColAlignment(menuPriceCol.规格) = flexAlignLeftCenter
        .ColAlignment(menuPriceCol.厂牌) = flexAlignLeftCenter
        .ColAlignment(menuPriceCol.单位) = flexAlignCenterCenter
        .ColAlignment(menuPriceCol.原成本价) = flexAlignRightCenter
        .ColAlignment(menuPriceCol.现成本价) = flexAlignRightCenter
        .ColAlignment(menuPriceCol.原零售价) = flexAlignRightCenter
        .ColAlignment(menuPriceCol.现零售价) = flexAlignRightCenter
        .ColAlignment(menuPriceCol.原采购限价) = flexAlignRightCenter
        .ColAlignment(menuPriceCol.原指导售价) = flexAlignRightCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter '列头居中对齐
        .ColComboList(menuPriceCol.品名) = "|..."
    End With

    With vsfStore
        .Editable = flexEDNone
        .Cols = menuStoreCol.总列数
        .Rows = 1
        .ColWidth(0) = 200
'        .RowHeight(1) = mconlngRowHeight
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = 50
        .RowHeight(0) = mconlngRowHeight
        .AllowSelection = False '不能多选
'        .SelectionMode = flexSelectionByRow '整行选择
        .ExplorerBar = flexExMoveRows '拖动
        .AllowUserResizing = flexResizeBoth  '可以改变行列宽度
        .GridLineWidth = 2
        .GridLines = flexGridInset
        .GridColor = &H0&

        '设置列名
        .TextMatrix(0, menuStoreCol.材料ID) = "材料id"
        .TextMatrix(0, menuStoreCol.库房) = "库房"
        .TextMatrix(0, menuStoreCol.库房ID) = "库房id"
        .TextMatrix(0, menuStoreCol.供应商) = "供应商"
        .TextMatrix(0, menuStoreCol.供应商ID) = "供应商id"
        .TextMatrix(0, menuStoreCol.药品) = "卫材"
        .TextMatrix(0, menuStoreCol.规格) = "规格"
        .TextMatrix(0, menuStoreCol.单位) = "单位"
        .TextMatrix(0, menuStoreCol.批号) = "批号"
        .TextMatrix(0, menuStoreCol.效期) = "效期"
        .TextMatrix(0, menuStoreCol.产地) = "产地"
        .TextMatrix(0, menuStoreCol.数量) = "数量"
        .TextMatrix(0, menuStoreCol.包装系数) = "包装系数"
        .TextMatrix(0, menuStoreCol.批次) = "批次"
        .TextMatrix(0, menuStoreCol.变价) = "变价"
        .TextMatrix(0, menuStoreCol.原零售价) = "原零售价"
        .TextMatrix(0, menuStoreCol.现零售价) = "现零售价"
        .TextMatrix(0, menuStoreCol.调整金额) = "调整金额"
        .TextMatrix(0, menuStoreCol.加成率) = "加成率"
        .TextMatrix(0, menuStoreCol.原采购价) = "原采购价"
        .TextMatrix(0, menuStoreCol.现采购价) = "现采购价"
        .TextMatrix(0, menuStoreCol.差价差) = "差价差"
        '设置列宽
        .ColWidth(0) = 0
        .ColWidth(menuStoreCol.库房) = 1500
        .ColWidth(menuStoreCol.库房ID) = 0
        .ColWidth(menuStoreCol.供应商) = 2000
        .ColWidth(menuStoreCol.供应商ID) = 0
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
        .ColWidth(menuStoreCol.调整金额) = 1000
        .ColWidth(menuStoreCol.加成率) = 1000
        .ColWidth(menuStoreCol.原采购价) = 1000
        .ColWidth(menuStoreCol.现采购价) = 1000
        .ColWidth(menuStoreCol.差价差) = 1000
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
        .ColAlignment(menuStoreCol.调整金额) = flexAlignRightCenter
        .ColAlignment(menuStoreCol.加成率) = flexAlignRightCenter
        .ColAlignment(menuStoreCol.原采购价) = flexAlignRightCenter
        .ColAlignment(menuStoreCol.现采购价) = flexAlignRightCenter
        .ColAlignment(menuStoreCol.差价差) = flexAlignRightCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter '列头居中对齐
    End With

    With vsfPay
        .Editable = flexEDNone
        .Cols = menuPayCol.总列数
        .Rows = 1
        .ColWidth(0) = 200
'        .RowHeight(1) = mconlngRowHeight
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = 50
        .RowHeight(0) = mconlngRowHeight
        .AllowSelection = False '不能多选
'        .SelectionMode = flexSelectionByRow '整行选择
        .ExplorerBar = flexExMoveRows '拖动
        .AllowUserResizing = flexResizeBoth  '可以改变行列宽度
        .GridLineWidth = 2
        .GridLines = flexGridInset
        .GridColor = &H0&

        .TextMatrix(0, menuPayCol.材料ID) = "材料id"
        .TextMatrix(0, menuPayCol.供应商) = "供应商"
        .TextMatrix(0, menuPayCol.供应商ID) = "供应商id"
        .TextMatrix(0, menuPayCol.品名) = "品名"
        .TextMatrix(0, menuPayCol.发票号) = "发票号"
        .TextMatrix(0, menuPayCol.发票日期) = "发票日期"
        .TextMatrix(0, menuPayCol.发票金额) = "发票金额"
        .TextMatrix(0, menuPayCol.规格) = "规格"
        .TextMatrix(0, menuPayCol.产地) = "产地"
        '设置列宽
        .ColWidth(menuPayCol.材料ID) = 0
        .ColWidth(menuPayCol.供应商) = 1500
        .ColWidth(menuPayCol.品名) = 2000
        .ColWidth(menuPayCol.发票号) = 1500
        .ColWidth(menuPayCol.发票日期) = 2000
        .ColWidth(menuPayCol.发票金额) = 1500
        .ColHidden(menuPayCol.供应商ID) = True
        .ColHidden(menuPayCol.规格) = True
        .ColHidden(menuPayCol.产地) = True
        '对齐方式
        .ColAlignment(menuPayCol.品名) = flexAlignLeftCenter
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

    On Error GoTo ErrHandle
    '调价方式 0-调售价;1-调成本价;2-调售价及成本价
    If mintMethod = 0 Then
        gstrSQL = "Select Distinct p.原价id, i.是否变价, Nvl(s.指导批发价, 0) As 指导批价, Nvl(s.扣率, 0) As 扣率, Nvl(s.指导零售价, 0) As 指导售价," & vbNewLine & _
            "                 nvl(s.加成率,0) / 100 As 加成率, i.编码, b.名称 As 商品名, i.名称 As 通用名, i.规格, i.产地 As 厂牌, i.计算单位 As 单位," & vbNewLine & _
            "                s.包装单位,s.换算系数, s.成本价 As 原成本价, s.成本价 As 新成本价, p.原价, p.现价," & vbNewLine & _
            "                p.收入项目id, p.调价人, p.调价说明, s.差价让利比, To_Char(a.执行日期, 'YYYY-MM-DD HH24:MI:SS') As 执行日期, i.Id 材料id," & vbNewLine & _
            "                Decode(k.药品id, Null, 0, 1) 是否有库存" & vbNewLine & _
            "From (Select 药品id From 药品库存 where 性质=1) K, 调价汇总记录 A, 收费项目别名 B, 材料特性 S, 收费项目目录 I, 收费价目 P" & vbNewLine & _
            "Where a.调价号 = p.调价汇总号 And b.收费细目id(+) = s.材料id And s.材料id = i.Id And i.Id = k.药品id(+) And i.Id = p.收费细目id And" & vbNewLine & _
            "      p.调价汇总号 = [1] And a.分类 = 1 And b.性质(+) = 3 And a.调价号 = [1] " & vbNewLine & _
            IIf(mintModal = 2, "", "  And (i.撤档时间 Is Null Or i.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))") & vbNewLine & _
            GetPriceClassString("P") & " Order By 材料id"
    ElseIf mintMethod = 1 Then
        gstrSQL = "Select Distinct i.是否变价, Nvl(s.指导批发价, 0) As 指导批价, Nvl(s.扣率, 0) As 扣率, Nvl(s.指导零售价, 0) As 指导售价," & vbNewLine & _
            "                nvl(s.加成率,0) / 100 As 加成率, i.编码, b.名称 As 商品名, i.名称 As 通用名, i.规格, m.产地 As 厂牌, i.计算单位 As 单位," & vbNewLine & _
            "                s.包装单位,s.换算系数, m.原成本价, m.新成本价, p.现价 as 原价, p.现价, p.收入项目id," & vbNewLine & _
            "                a.填制人 As 调价人, a.说明 As 调价说明, s.差价让利比, To_Char(m.执行日期, 'YYYY-MM-DD HH24:MI:SS') As 执行日期, i.Id 材料id," & vbNewLine & _
            "                Decode(k.药品id, Null, 0, 1) 是否有库存" & vbNewLine & _
            "From (Select Min(原成本价) As 原成本价, Min(新成本价) As 新成本价, min(产地) as 产地,调价汇总号,药品id,min(执行日期) as 执行日期 From 成本价调价信息 Where 调价汇总号 = [1] Group By 调价汇总号,药品id) M, (Select 药品id From 药品库存 where 性质=1) K, 调价汇总记录 A, 收费项目别名 B, 材料特性 S, 收费项目目录 I, 收费价目 P" & vbNewLine & _
            "Where m.调价汇总号(+) = a.调价号 And b.收费细目id(+) = s.材料id And s.材料id = i.Id And i.Id = k.药品id(+) And m.药品id = i.Id And" & vbNewLine & _
            "      i.Id = p.收费细目id And Sysdate Between p.执行日期 And p.终止日期 And m.调价汇总号 = [1] And a.分类 = 1 And b.性质(+) = 3 And" & vbNewLine & _
            "      a.调价号 = [1] " & IIf(mintModal = 2, "", " And (i.撤档时间 Is Null Or i.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))") & vbNewLine & _
            GetPriceClassString("P") & " Order By 材料id"
    ElseIf mintMethod = 2 Then
        gstrSQL = "Select distinct p.原价id, i.是否变价, Nvl(s.指导批发价, 0) As 指导批价, Nvl(s.扣率, 0) As 扣率, Nvl(s.指导零售价, 0) As 指导售价," & vbNewLine & _
            "       nvl(s.加成率,0) / 100 As 加成率, i.编码, b.名称 As 商品名, i.名称 As 通用名, i.规格, decode(m.产地,null,i.产地,m.产地) As 厂牌, i.计算单位 As 单位," & vbNewLine & _
            "       s.包装单位,s.换算系数, m.原成本价, m.新成本价, p.原价, p.现价, p.收入项目id, p.调价人, p.调价说明, s.差价让利比," & vbNewLine & _
            "       To_Char(p.执行日期, 'YYYY-MM-DD HH24:MI:SS') As 执行日期, i.Id 材料id, Decode(k.药品id, Null, 0, 1) 是否有库存" & vbNewLine & _
            "From (Select 药品id,Min(原成本价) As 原成本价, Min(新成本价) As 新成本价, min(产地) as 产地,调价汇总号 From 成本价调价信息 Where 调价汇总号 = [1] Group By 药品id,调价汇总号) M, 收费价目 P, 调价汇总记录 A, (Select 药品id From 药品库存 where 性质=1) K, 收费项目别名 B, 材料特性 S, 收费项目目录 I" & vbNewLine & _
            "Where m.调价汇总号 = a.调价号 and m.药品id=i.id And p.调价汇总号 = a.调价号 And p.收费细目id = k.药品id(+) And p.收费细目id = b.收费细目id(+) And p.收费细目id = s.材料id And" & vbNewLine & _
            "      s.材料id = i.Id And a.调价号 =[1] And b.性质(+) = 3 And a.分类 = 1 " & vbNewLine & _
            GetPriceClassString("P") & vbNewLine & _
            IIf(mintModal = 2, "", "  And (i.撤档时间 Is Null Or i.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))") & " order by 材料id"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr调价汇总号)
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "该调价记录已经被删除了！", vbInformation, gstrSysName
        Exit Sub
    End If

    With vsfPrice
        .Rows = 2
        rsTemp.MoveFirst
        For i = 0 To rsTemp.RecordCount - 1
            If rsTemp!材料ID <> lngDrugID Then
                Select Case mintUnit
                    Case 0
                        db包装系数 = 1
                        strUnit = rsTemp!单位
                    Case 1
                        db包装系数 = rsTemp!换算系数
                        strUnit = rsTemp!包装单位
                End Select

                lngDrugID = rsTemp!材料ID

                If mintMethod = 0 Or mintMethod = 2 Then
                    .TextMatrix(.Rows - 1, menuPriceCol.原价id) = IIf(IsNull(rsTemp!原价id), "", rsTemp!原价id)
                End If
                .TextMatrix(.Rows - 1, menuPriceCol.材料ID) = lngDrugID

                .TextMatrix(.Rows - 1, menuPriceCol.品名) = "[" & rsTemp!编码 & "]" & IIf(IsNull(rsTemp!商品名), rsTemp!通用名, rsTemp!商品名)
                .TextMatrix(.Rows - 1, menuPriceCol.规格) = rsTemp!规格
                .TextMatrix(.Rows - 1, menuPriceCol.是否变价) = rsTemp!是否变价
                
                If mintMethod = 1 Or mintMethod = 2 Then
                    gstrSQL = "select min(产地) as 厂牌 from 成本价调价信息 where 调价汇总号=[1] and 药品id=[2]"
                    Set rs产地 = zlDatabase.OpenSQLRecord(gstrSQL, "产地查询", mstr调价汇总号, lngDrugID)
                    If rs产地.RecordCount > 0 Then
                        .TextMatrix(.Rows - 1, menuPriceCol.厂牌) = IIf(IsNull(rs产地!厂牌), "", rs产地!厂牌)
                    End If
                Else
                    .TextMatrix(.Rows - 1, menuPriceCol.厂牌) = IIf(IsNull(rsTemp!厂牌), "", rsTemp!厂牌)
                End If
                
                .TextMatrix(.Rows - 1, menuPriceCol.单位) = strUnit
                .TextMatrix(.Rows - 1, menuPriceCol.包装系数) = db包装系数

                .TextMatrix(.Rows - 1, menuPriceCol.加成率) = rsTemp!加成率
                .TextMatrix(.Rows - 1, menuPriceCol.差价让利比) = rsTemp!差价让利比
                .TextMatrix(.Rows - 1, menuPriceCol.是否有库存) = rsTemp!是否有库存
                .TextMatrix(.Rows - 1, menuPriceCol.收入项目id) = IIf(IsNull(rsTemp!收入项目id), "", rsTemp!收入项目id)
                .TextMatrix(.Rows - 1, menuPriceCol.原成本价) = Format(rsTemp!原成本价 * db包装系数, mFMT.FM_成本价)
                .TextMatrix(.Rows - 1, menuPriceCol.现成本价) = Format(rsTemp!新成本价 * db包装系数, mFMT.FM_成本价)
                .TextMatrix(.Rows - 1, menuPriceCol.原零售价) = Format(IIf(IsNull(rsTemp!原价), rsTemp!现价, rsTemp!原价) * db包装系数, mFMT.FM_零售价)
                .TextMatrix(.Rows - 1, menuPriceCol.现零售价) = Format(rsTemp!现价 * db包装系数, mFMT.FM_零售价)
                .TextMatrix(.Rows - 1, menuPriceCol.原采购限价) = Format(rsTemp!指导批价 * db包装系数, mFMT.FM_成本价)
                .TextMatrix(.Rows - 1, menuPriceCol.现采购限价) = Format(rsTemp!指导批价 * db包装系数, mFMT.FM_成本价)
                .TextMatrix(.Rows - 1, menuPriceCol.原指导售价) = Format(rsTemp!指导售价 * db包装系数, mFMT.FM_零售价)
                .TextMatrix(.Rows - 1, menuPriceCol.现指导售价) = Format(rsTemp!指导售价 * db包装系数, mFMT.FM_零售价)

                txtValuer.Text = IIf(IsNull(rsTemp!调价人), "", rsTemp!调价人)
                txtSummary.Text = IIf(IsNull(rsTemp!调价说明), "", rsTemp!调价说明)
                If mintModal = 1 Then
                    Me.dtpRunDate.MinDate = CDate(rsTemp!执行日期)
                End If
                If IsNull(rsTemp!执行日期) Then
                    StrToday = Format(sys.Currentdate(), "yyyy-MM-dd hh:mm:ss")
                Else
                    StrToday = Format(rsTemp!执行日期, "yyyy-MM-dd hh:mm:ss")
                End If
                Me.dtpRunDate.Value = CDate(StrToday)

                .Rows = .Rows + 1
                Call setColEdit
                .RowHeight(.Rows - 1) = mconlngRowHeight
            End If
            rsTemp.MoveNext
        Next
        Call GetDrugStore(Val(.TextMatrix(1, menuPriceCol.材料ID)), 1)
    End With

    Exit Sub
ErrHandle:
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

    '查找卫材
    On Error GoTo ErrHandle
    If strInput <> txtFind.Tag Then
        '表示新的查找
        txtFind.Tag = strInput

        gstrSQL = "Select Distinct A.Id,'[' || A.编码 || ']' As 药品编码, A.名称 As 通用名, B.名称 As 商品名 " & _
                  "From 收费项目目录 A,收费项目别名 B " & _
                  "Where (A.站点 = [3] Or A.站点 is Null) And A.Id =B.收费细目id And A.类别='4' " & _
                  "  And (A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2] ) " & _
                  "Order By 药品编码 "
        Set mrsFindName = zlDatabase.OpenSQLRecord(gstrSQL, "取匹配的材料id", strInput & "%", "%" & strInput & "%", gstrNodeNo)

        If mrsFindName.RecordCount = 0 Then Exit Sub
        mrsFindName.MoveFirst
    End If

    '开始查找
    If mrsFindName.State <> adStateOpen Then Exit Sub
    If mrsFindName.RecordCount = 0 Then Exit Sub

    For n = 1 To mrsFindName.RecordCount
        '如果到底了，则返回第1条记录
        If mrsFindName.EOF Then mrsFindName.MoveFirst

        str药名 = mrsFindName!药品编码 & IIf(IsNull(mrsFindName!商品名), mrsFindName!通用名, mrsFindName!商品名)

        For lngRow = 1 To vsfPrice.Rows - 1
            lngFindRow = vsfPrice.FindRow(str药名, lngRow, CLng(menuPriceCol.品名), True, True)
            If lngFindRow > 0 Then
                vsfPrice.Select lngFindRow, 1, lngFindRow, vsfPrice.Cols - 1
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
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    If vsfPrice.Height + Y <= 800 Then Exit Sub
    If TabCtlDetails.Height - Y <= 1000 Then Exit Sub
    picSplit.Move 0, picSplit.Top + Y
    vsfPrice.Move 0, fraCondition.Top + fraCondition.Height + 20, Me.ScaleWidth, vsfPrice.Height + Y

    With TabCtlDetails
        .Top = picSplit.Top + picSplit.Height + 5
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = TabCtlDetails.Height - Y
    End With
End Sub

Private Sub txtfind_KeyPress(KeyAscii As Integer)
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

    On Error GoTo ErrHandle
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
            " And 末级=1 And substr(类型,5,1) = '1' And (撤档时间 is null or 撤档时间=to_date('3000-01-01','YYYY-MM-DD')) " & _
            " Order By 编码 "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, strTmp & "%", IIf(gstrMatchMethod = "0", "%", "") & strTmp & "%")

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
ErrHandle:
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
        If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mconlngCanColor Then
            .EditCell
            .EditSelStart = 0
            .EditSelLength = Len(.EditText)
        End If
    End With
End Sub

Private Sub vsfPay_EnterCell()
    With vsfPrice
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
            If .Col = menuPayCol.品名 Then
                .Col = menuPayCol.供应商
            ElseIf .Col = menuPayCol.供应商 Then
                .Col = menuPayCol.发票号
            ElseIf .Col = menuPayCol.发票号 Then
                .Col = menuPayCol.发票日期
            ElseIf .Col = menuPayCol.发票日期 Then
                .Col = menuPayCol.发票金额
            ElseIf .Col = menuPayCol.发票金额 And .Row <> .Rows - 1 Then
                .Col = menuPayCol.品名
                .Row = .Row + 1
            End If
        End If
    End With
End Sub

Private Sub vsfPay_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        With vsfPay
            If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mconlngCanColor Then
                .Editable = flexEDKbdMouse
            Else
                .Editable = flexEDNone
            End If

        End With
    End If
End Sub

Private Sub vsfPay_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strKey As String
    Dim intDigit As Integer

    If KeyAscii = vbKeyReturn Then Exit Sub
    If KeyAscii <> vbKeyBack Then
        With vsfPay
            If Col = menuPayCol.发票金额 Then
                strKey = .EditText
                intDigit = Len(Mid(mFMT.FM_金额, InStr(1, mFMT.FM_金额, ".") + 1))
                If KeyAscii = vbKeyDelete Then
                    If InStr(1, .EditText, ".") > 0 Then
                        KeyAscii = 0
                    End If
                ElseIf KeyAscii = Asc(".") Or (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
                    If .EditSelLength = Len(strKey) Then Exit Sub
                    If InStr(strKey, ".") <> 0 And Chr(KeyAscii) = "." Then   '只能存在一个小数点
                        KeyAscii = 0
                        Exit Sub
                    End If
                    If Len(Mid(strKey, InStr(1, strKey, ".") + 1)) >= intDigit And strKey Like "*.*" Then
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
    Dim strKey As String

    With vsfPay
        If Col = menuPayCol.发票日期 Then
            strKey = .EditText
            If strKey <> "" Then
                If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
                    strKey = TranNumToDate(strKey)
                    If strKey = "" Then
                        MsgBox "对不起，发票日期必须为日期型,格式(20000101或者2000-01-01)！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                    .EditText = strKey
                    .TextMatrix(Row, menuPayCol.发票日期) = .EditText
                End If

                If Not IsDate(strKey) Then
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

        If chkAppAllColumn.Value = 1 Then
            Call AutoCalc所有库存价格
        End If
    End With
End Sub

Private Sub AutoCalc所有库存价格()
    '-----------------------------------------------------------------------------------------------------------
    '功能:自动计算所有库存的价格
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, dbl现成本价 As Double, dbl现价 As Double, dbl加成率 As Double, dbl成本差价 As Double, dbl差价调整额 As Double, dbl调整额 As Double
    Dim lng材料ID As Long, bln库房分批 As Boolean, lng供应商ID As Long, lngTemp As Long, i As Long
    Dim blnHaveData As Boolean, lngStep As Long, lngSteps As Long
    Dim intCol As Integer
    Dim cllData As New Collection

    err = 0: On Error GoTo ErrHand:

    '因为存在包装换算问题，因此，目前按最小单位进行设置单价
    dbl现成本价 = Val(vsfPrice.TextMatrix(vsfPrice.Row, menuPriceCol.现成本价))
    dbl现价 = Val(vsfPrice.TextMatrix(vsfPrice.Row, menuPriceCol.现零售价))

    With vsfStore
        For lngRow = 1 To .Rows - 1
            If vsfPrice.Col = menuPriceCol.现成本价 Then
                .TextMatrix(lngRow, menuStoreCol.现采购价) = dbl现成本价
                '加成率=现零售价/现成本价-1
                If dbl现成本价 <> 0 Then
                    dbl加成率 = Round(Val(.TextMatrix(lngRow, menuStoreCol.现零售价)) / dbl现成本价 - 1, 7)
                Else
                    dbl加成率 = 0

                End If
                '差价调整额=(现成本价-原成本价)
                dbl成本差价 = Round((Val(.TextMatrix(lngRow, menuStoreCol.原采购价)) - dbl现成本价), 7)
            ElseIf vsfPrice.Col = menuPriceCol.现零售价 Then
                .TextMatrix(lngRow, menuStoreCol.现零售价) = dbl现价
                '现价发生改变时,需要重新根据加成率计算相关的现成本价
                dbl加成率 = Round(Val(.TextMatrix(lngRow, menuStoreCol.加成率)) / 100, 7)
                If dbl加成率 = -1 Then dbl加成率 = 0
                '现成本价=现零售价/(1+加成率)
                dbl现成本价 = Round(dbl现价 / (1 + dbl加成率), 7)
                '差价调整额=(现成本价-原成本价)
                dbl成本差价 = (dbl现成本价 - Val(.TextMatrix(lngRow, menuStoreCol.原采购价)))

                '调整额=数量*(现价-原价)
                dbl调整额 = (dbl现价 - Val(.TextMatrix(lngRow, menuStoreCol.原零售价))) * Val(.TextMatrix(lngRow, menuStoreCol.数量))
                .TextMatrix(lngRow, menuStoreCol.调整金额) = Format(dbl调整额, mFMT.FM_金额)
            End If

            lng材料ID = Val(.TextMatrix(lngRow, menuStoreCol.材料ID))
            lng供应商ID = Val(.TextMatrix(lngRow, menuStoreCol.供应商ID))

            If dbl加成率 = -1 Then dbl加成率 = 0
            .TextMatrix(lngRow, menuStoreCol.加成率) = Format(dbl加成率 * 100, GFM_VBJCL)
            dbl成本差价 = (Val(.TextMatrix(lngRow, menuStoreCol.原采购价)) - dbl现成本价)
             '差价调整额=(现成本价-原成本价)*数量
             dbl差价调整额 = Round(dbl成本差价 * Val(.TextMatrix(lngRow, menuStoreCol.数量)), 7)
            .TextMatrix(lngRow, menuStoreCol.差价差) = Format(dbl差价调整额, mFMT.FM_金额)
            lngTemp = Val(.TextMatrix(lngRow, menuStoreCol.材料ID))
            lng供应商ID = Val(.TextMatrix(lngRow, menuStoreCol.供应商ID))

            If lng供应商ID <> 0 Then
                err = 0: On Error Resume Next
                cllData.Add Array(lngTemp, lng供应商ID, dbl差价调整额, .TextMatrix(lngRow, menuStoreCol.供应商ID), .TextMatrix(lngRow, menuStoreCol.材料ID), .TextMatrix(lngRow, menuStoreCol.规格), .TextMatrix(lngRow, menuStoreCol.产地)), "K" & lng供应商ID & "_" & lngTemp
                If err <> 0 Then
                    '累计差价调整额
                    dbl差价调整额 = Val(cllData("K" & lng供应商ID & "_" & lngTemp)(2)) + dbl差价调整额
                    cllData.Remove "K" & lng供应商ID & "_" & lngTemp
                     err = 0: On Error GoTo ErrHand:
                    cllData.Add Array(lngTemp, lng供应商ID, dbl差价调整额, .TextMatrix(lngRow, menuStoreCol.供应商ID), .TextMatrix(lngRow, menuStoreCol.材料ID), .TextMatrix(lngRow, menuStoreCol.规格), .TextMatrix(lngRow, menuStoreCol.产地)), "K" & lng供应商ID & "_" & lngTemp

                End If
                On Error GoTo ErrHand:
            End If
        Next

        If chkAutoPay.Value = 1 Then
            '需要自动计算相关的应付变动记录
            For i = 1 To cllData.Count
                With vsfPay
                    blnHaveData = False
                    For lngRow = 1 To .Rows - 1
                        lngTemp = Val(.TextMatrix(lngRow, menuPayCol.材料ID))
                        lng供应商ID = Val(.TextMatrix(lngRow, menuPayCol.供应商ID))
                        If lngTemp = Val(cllData(i)(0)) _
                            And lng供应商ID = Val(cllData(i)(1)) Then
                            '卫材及供应商相同,清空相关的值
                            .TextMatrix(lngRow, menuPayCol.发票金额) = Format(Val(cllData(i)(2)), mFMT.FM_金额)
                             blnHaveData = True
                        End If
                    Next
                    If blnHaveData = False Then
                        '需要增加该项供应商的物资
                        If Val(.TextMatrix(.Rows - 1, menuStoreCol.材料ID)) <> 0 Then
                            .Rows = .Rows + 1
                        End If
                        lngRow = .Rows - 1
                        .TextMatrix(lngRow, menuPayCol.供应商ID) = cllData(i)(3)
                        .TextMatrix(lngRow, menuPayCol.材料ID) = cllData(i)(0)
                        .TextMatrix(lngRow, menuPayCol.规格) = cllData(i)(5)
                        .TextMatrix(lngRow, menuPayCol.产地) = cllData(i)(6)
                        .TextMatrix(lngRow, menuPayCol.发票金额) = Format(Val(cllData(i)(2)), mFMT.FM_金额)
                    End If
                End With
            Next
        End If
    End With

    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsfPrice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Then Exit Sub
'    Call SetRowHidden(Val(vsfPrice.TextMatrix(NewRow, menuPriceCol.材料id)))
End Sub

Private Sub SetRowHidden(ByVal lngDrugID As Long)
    '功能：行的显示与隐藏
    '参数：材料id
    Dim intRow As Integer

    If lngDrugID = 0 Then Exit Sub
    With vsfStore
        For intRow = 1 To .Rows - 1
            If Val(.TextMatrix(intRow, menuStoreCol.材料ID)) = lngDrugID Then
                .RowHidden(intRow) = False
            Else
                .RowHidden(intRow) = True
            End If
        Next
    End With

    With vsfPay
        For intRow = 1 To .Rows - 1
            If Val(.TextMatrix(intRow, menuPayCol.材料ID)) = lngDrugID Then
                .RowHidden(intRow) = False
            Else
                .RowHidden(intRow) = True
            End If
        Next
    End With
End Sub

'Private Sub vsfPrice_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'    With vsfPrice
'        mlngPrice = Val(.TextMatrix(Row, Col))
'        If .Cell(flexcpBackColor, Row, Col, Row, Col) = mconlngColor Then
'            Cancel = True
'        End If
'    End With
'End Sub

Private Sub vsfPrice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim mrsReturn As Recordset

    mblnClick = True
    mblnUpdateAdd = True

    On Error GoTo ErrHandle
    Set mrsReturn = SelectStuff("")
    If mrsReturn Is Nothing Then Exit Sub
    If mrsReturn.RecordCount = 0 Then Exit Sub

    Call GetDrugPirce(mrsReturn, Row)
    mblnUpdateAdd = False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SelectStuff(ByVal strKey As String) As ADODB.Recordset
    '-----------------------------------------------------------------------------------------------------------
    '功能:选择指定的卫生材料
    '参数:strKey-多选择的条件
    '返回:选择成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/09/17
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim rsDrugInfo As ADODB.Recordset
    Dim blnCancel As Boolean, i As Long
    Dim vRect As RECT, sngX As Single, sngY As Single
    Dim int系数 As Integer
    'zlDatabase.ShowSelect
    '功能：多功能选择器
    '参数：
    '     frmParent=显示的父窗体
    '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
    '     bytStyle=选择器风格
    '       为0时:列表风格:ID,…
    '       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
    '       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
    '     strTitle=选择器功能命名,也用于个性化区分
    '     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
    '     strSeek=当bytStyle<>2时有效,缺省定位的项目。
    '             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
    '             bytStyle=1时,可以是编码或名称
    '     strNote=选择器的说明文字
    '     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
    '     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
    '     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
    '     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
    '     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
    '     blnSearch=是否显示行号,并可以输入行号定位
    '返回：取消=Nothing,选择=SQL源的单行记录集
    '说明：
    '     1.ID和上级ID可以为字符型数据
    '     2.末级等字段不要带空值
    '应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    err = 0: On Error GoTo ErrHand:
    Call CalcPosition(sngX, sngY, vsfPrice)

    Set rsDrugInfo = New ADODB.Recordset
    With rsDrugInfo
        If .State = 1 Then .Close
        .Fields.Append "id", adDouble, 20, adFldIsNullable
        .Fields.Append "编码", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "商品名", adLongVarChar, 18, adFldIsNullable
        .Fields.Append "通用名", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "规格", adLongVarChar, 18, adFldIsNullable
        .Fields.Append "产地", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "计算单位", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "换算系数", adDouble, 40, adFldIsNullable
        .Fields.Append "包装单位", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "类型", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "时价", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "成本价", adDouble, 40, adFldIsNullable
        .Fields.Append "指导批发价", adDouble, 40, adFldIsNullable
        .Fields.Append "指导零售价", adDouble, 40, adFldIsNullable
        .Fields.Append "跟踪在用", adDouble, 1, adFldIsNullable

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With

    If strKey <> "" Then
        strKey = GetMatchingSting(strKey)
        gstrSQL = "" & _
            "   Select distinct I.ID,I.编码,b.名称 As 商品名, i.名称 As 通用名,I.规格,I.产地,I.计算单位,P.换算系数,P.包装单位," & _
            "         decode(I.是否变价,1,'时价','定价') 类型,Decode(i.是否变价, 0, '定价', 1, '时价') As 时价," & _
            "         to_char(p.成本价,'9999999999990.9999999') as 成本价," & _
            "         to_char(p.指导批发价,'9999999999990.9999999') 指导批发价," & _
            "         to_char(p.指导零售价,'9999999999990.9999999') 指导零售价," & _
            "          P.跟踪在用" & _
            "   From 收费项目目录 I,收费项目别名 N,材料特性 P,收费项目别名 B" & _
            "   Where I.ID=N.收费细目ID And I.ID=P.材料ID  and i.Id = b.收费细目id(+) and b.性质(+) = 3 And i.类别 = '4' " & _
            "       and (I.编码 like [1] or N.简码 Like [1] or N.名称 Like [1])" & _
            "       and (I.撤档时间 Is Null Or I.撤档时间=To_Date('3000-01-01','yyyy-MM-dd'))"
     Else
        gstrSQL = "" & _
            "   Select distinct  I.ID,I.编码,b.名称 As 商品名, i.名称 As 通用名,I.规格,I.产地,I.计算单位,P.换算系数,P.包装单位, " & _
            "           decode(I.是否变价,1,'时价','定价') 类型,Decode(i.是否变价, 0, '定价', 1, '时价') As 时价," & _
            "           to_char(p.成本价,'9999999999990.9999999') as 成本价," & _
            "           to_char(p.指导批发价,'9999999999990.9999999') 指导批发价," & _
            "           to_char(p.指导零售价,'9999999999990.9999999') 指导零售价," & _
            "           P.跟踪在用" & _
            "   From 收费项目目录 I,材料特性 P,收费项目别名 B" & _
            "   Where I.ID=P.材料ID and i.Id = b.收费细目id(+) And" & _
            "   b.性质(+) = 3 And i.类别 = '4'" & _
            "           and (I.撤档时间 Is Null Or I.撤档时间=To_Date('3000-01-01','yyyy-MM-dd'))"

    End If

    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "卫生材料选择", False, "", "", False, False, True, sngX, sngY - vsfPrice.CellHeight, vsfPrice.CellHeight, blnCancel, False, False, strKey)
    If blnCancel = True Then Exit Function

    If rsTemp Is Nothing Then
        ShowMsgBox "不存在指定的卫生材料,请检查!"
        Exit Function
    End If

    With rsDrugInfo
        .AddNew
        !Id = rsTemp!Id
        !编码 = rsTemp!编码
        !商品名 = rsTemp!商品名
        !通用名 = rsTemp!通用名
        !规格 = rsTemp!规格
        !产地 = rsTemp!产地
        !计算单位 = rsTemp!计算单位
        !换算系数 = rsTemp!换算系数
        !包装单位 = rsTemp!包装单位
        !时价 = rsTemp!时价
        !成本价 = rsTemp!成本价
        !指导批发价 = rsTemp!指导批发价
        !指导零售价 = rsTemp!指导零售价
        !跟踪在用 = rsTemp!跟踪在用

        .Update
    End With

    Set SelectStuff = rsDrugInfo

    Exit Function
ErrHand:
    vsfPrice.Redraw = flexRDBuffered
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub GetDrugPirce(ByVal rsReturn As ADODB.Recordset, ByVal Row As Integer)
    '用来获取药品信息
    Dim rsTemp As Recordset
    Dim lngDrugID As Long
    Dim lngRow As Long
    Dim i As Long
    Dim intCurrentPrice As Integer '是否是时价
    Dim strUnit As String
    Dim db包装系数 As Double
    Dim dbl比率 As Double

    On Error GoTo ErrHandle

    mlngOldStuffID = Val(vsfPrice.TextMatrix(Row, menuPriceCol.材料ID))
    Set rsReturn = CheckDoubleDrug(rsReturn)
    If rsReturn.RecordCount = 0 Then Exit Sub

    rsReturn.MoveFirst
    For i = 0 To rsReturn.RecordCount - 1
        With vsfPrice
            lngDrugID = rsReturn!Id

            '检查是否存在未执行的价格
            If checkNotExecutePrice(lngDrugID) = True Then Exit Sub

            Select Case mintUnit
                Case 0  '散装单位
                    db包装系数 = 1
                    strUnit = rsReturn!计算单位
                Case 1  '包装单位
                    db包装系数 = rsReturn!换算系数
                    strUnit = rsReturn!包装单位
            End Select

            .TextMatrix(Row, menuPriceCol.材料ID) = lngDrugID

            .EditText = "[" & rsReturn!编码 & "]" & IIf(IsNull(rsReturn!商品名) Or rsReturn!商品名 = "", rsReturn!通用名, rsReturn!商品名)
            .TextMatrix(Row, menuPriceCol.品名) = IIf(.EditText = "", "[" & rsReturn!编码 & "]" & IIf(IsNull(rsReturn!商品名) Or rsReturn!商品名 = "", rsReturn!通用名, rsReturn!商品名), .EditText)

            .TextMatrix(Row, menuPriceCol.规格) = IIf(IsNull(rsReturn!规格), "", rsReturn!规格)
            .TextMatrix(Row, menuPriceCol.是否变价) = IIf(rsReturn!时价 = "时价", 1, 0)
            intCurrentPrice = IIf(rsReturn!时价 = "时价", 1, 0)
            .TextMatrix(Row, menuPriceCol.厂牌) = IIf(IsNull(rsReturn!产地), "", rsReturn!产地)
            .TextMatrix(Row, menuPriceCol.单位) = strUnit
            .TextMatrix(Row, menuPriceCol.包装系数) = db包装系数
            .TextMatrix(Row, menuPriceCol.是否跟踪在用) = zlStr.nvl(rsReturn!跟踪在用)
            .TextMatrix(Row, menuPriceCol.现成本价) = Format(Val(zlStr.nvl(rsReturn!成本价)) * db包装系数, mFMT.FM_成本价)
            .TextMatrix(Row, menuPriceCol.原采购限价) = Format(Val(zlStr.nvl(rsReturn!指导批发价)) * db包装系数, mFMT.FM_成本价)
            .TextMatrix(Row, menuPriceCol.现采购限价) = .TextMatrix(Row, menuPriceCol.原采购限价)
            .TextMatrix(Row, menuPriceCol.原指导售价) = Format(Val(zlStr.nvl(rsReturn!指导零售价)) * db包装系数, mFMT.FM_零售价)
            .TextMatrix(Row, menuPriceCol.现指导售价) = .TextMatrix(Row, menuPriceCol.原指导售价)

            gstrSQL = "select 药品id from 药品库存 where 药品id=[1] and 性质=1"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查库存", lngDrugID)
            If rsTemp.RecordCount = 0 Then
                .TextMatrix(Row, menuPriceCol.是否有库存) = 0
            Else
                .TextMatrix(Row, menuPriceCol.是否有库存) = 1
            End If

            If intCurrentPrice = 0 Then '定价卫材
                '表示定价药品调价，成本价取平均价格，售价取收费价目现价
                gstrSQL = "Select b.Id, Decode(Nvl(k.库存数量, 0), 0, a.成本价, (k.库存金额 - k.库存差价) / k.库存数量) As 成本价, a.指导批发价, a.指导零售价, b.现价, a.差价让利比," & vbNewLine & _
                            "       nvl(a.加成率,0) / 100 As 加成率, b.收入项目id" & vbNewLine & _
                            "From 材料特性 A, 收费价目 B," & vbNewLine & _
                            "     (Select Sum(实际金额) 库存金额, Sum(实际差价) As 库存差价, Sum(实际数量) 库存数量" & vbNewLine & _
                            "       From 药品库存" & vbNewLine & _
                            "       Where 性质 = 1 And 药品id = [1]) K" & vbNewLine & _
                            "Where a.材料id = b.收费细目id And a.材料id = [1] And (b.终止日期 Is Null Or b.终止日期 = To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
                            GetPriceClassString("B")
            Else '时价卫材
                '表示时价卫材调价，取库存金额/库存数量做为其价格
                gstrSQL = "" & _
                        "   Select  P.id,Decode(Nvl(K.库存数量,0),0,P.现价,K.库存金额/Nvl(K.库存数量,1)) 现价,nvl(m.加成率,0) / 100 as 加成率," & _
                        "           P.执行日期,P.收入项目id,I.名称 as 收入名称, " & IIf(mintUnit = 0, "1", " Nvl(M.换算系数,1)") & " as  系数,decode(nvl(k.库存数量,0),0,m.成本价,(k.库存金额-k.库存差价)/k.库存数量) as 成本价,m.跟踪在用,m.指导批发价,m.指导零售价,m.差价让利比" & _
                        "   From 收费价目 P,收入项目 I,材料特性 M," & _
                        "       (   Select Sum(实际金额) 库存金额,Sum(实际数量) 库存数量,Sum(实际差价) As 库存差价" & _
                        "           From 药品库存 " & _
                        "           Where  性质=1 and 药品ID=[1] " & _
                        "        ) K" & _
                        " where p.收费细目id=M.材料id and P.收入项目id=I.id and P.收费细目id=[1] " & _
                        "       and (P.终止日期 is null or P.终止日期=to_date('3000-01-01','YYYY-MM-DD'))" & _
                        GetPriceClassString("P")
            End If
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "查询药品", lngDrugID)
            If rsTemp.RecordCount > 0 Then

                .TextMatrix(Row, menuPriceCol.原价id) = rsTemp!Id
                .TextMatrix(Row, menuPriceCol.收入项目id) = IIf(IsNull(rsTemp!收入项目id), 0, rsTemp!收入项目id)
                .TextMatrix(Row, menuPriceCol.加成率) = GetFormat(IIf(IsNull(rsTemp!加成率), 0, rsTemp!加成率), 2)
                .TextMatrix(Row, menuPriceCol.差价让利比) = IIf(IsNull(rsTemp!差价让利比), 0, rsTemp!差价让利比)
                .TextMatrix(Row, menuPriceCol.原成本价) = Format(IIf(IsNull(rsTemp!成本价), 0, rsTemp!成本价) * db包装系数, mFMT.FM_成本价)
                .TextMatrix(Row, menuPriceCol.现成本价) = Format(IIf(IsNull(rsTemp!成本价), 0, rsTemp!成本价) * db包装系数, mFMT.FM_成本价)
                .TextMatrix(Row, menuPriceCol.原零售价) = Format(IIf(IsNull(rsTemp!现价), 0, rsTemp!现价) * db包装系数, mFMT.FM_零售价)
                If mstr调整额 = "" Or mint调价 = 1 Then
                    .TextMatrix(Row, menuPriceCol.现零售价) = Format(IIf(IsNull(rsTemp!现价), 0, rsTemp!现价) * db包装系数, mFMT.FM_零售价)
                Else
                    Select Case mintType
                        Case 1      '根据成本价加成
                            dbl比率 = 1 + Val(mdbl比率) / 100
                            .TextMatrix(Row, menuPriceCol.现零售价) = Format(Val(zlStr.nvl(rsTemp!成本价)) * dbl比率 * db包装系数, mFMT.FM_零售价)
                        Case 2      '根据零售价按比例
                            dbl比率 = 1 + Val(mdbl比率) / 100
                            .TextMatrix(Row, menuPriceCol.现零售价) = Format(Val(zlStr.nvl(rsTemp!现价)) * dbl比率 * db包装系数, mFMT.FM_零售价)
                        Case 3      '根据零售价按固定金额加减
                            dbl比率 = Val(mdbl比率)
                            .TextMatrix(Row, menuPriceCol.现零售价) = Format((Val(zlStr.nvl(rsTemp!现价)) * db包装系数) + dbl比率, mFMT.FM_零售价)
                    End Select
                End If

                If Val(.TextMatrix(Row, menuPriceCol.现零售价)) > Val(.TextMatrix(Row, menuPriceCol.现指导售价)) And Val(.TextMatrix(Row, menuPriceCol.现指导售价)) <> 0 Then
                    .TextMatrix(Row, menuPriceCol.现零售价) = Format(Val(.TextMatrix(Row, menuPriceCol.现指导售价)), mFMT.FM_零售价)
                End If
            Else
                .TextMatrix(Row, menuPriceCol.原价id) = 0
                If Row > 1 Then
                    .TextMatrix(Row, menuPriceCol.收入项目id) = .TextMatrix(Row - 1, menuPriceCol.收入项目id)
                End If
                .TextMatrix(Row, menuPriceCol.原零售价) = Format(0, mFMT.FM_零售价)
                .TextMatrix(Row, menuPriceCol.现零售价) = Format(0, mFMT.FM_零售价)
                .TextMatrix(Row, menuPriceCol.原成本价) = Format(0, mFMT.FM_成本价)
                .TextMatrix(Row, menuPriceCol.现成本价) = Format(0, mFMT.FM_成本价)

                If mstr调整额 = "" Or mint调价 = 1 Then
                    .TextMatrix(Row, menuPriceCol.现零售价) = Format(0, mFMT.FM_零售价)
                Else
                    Select Case mintType
                        Case 1      '根据成本价加成
                            dbl比率 = 1 + Val(mdbl比率) / 100
                            .TextMatrix(Row, menuPriceCol.现零售价) = Format(0 * dbl比率 * db包装系数, mFMT.FM_零售价)
                        Case 2      '根据零售价按比例
                            dbl比率 = 1 + Val(mdbl比率) / 100
                            .TextMatrix(Row, menuPriceCol.现零售价) = Format(0 * dbl比率 * db包装系数, mFMT.FM_零售价)
                        Case 3      '根据零售价按固定金额加减
                            dbl比率 = Val(mdbl比率)
                            .TextMatrix(Row, menuPriceCol.现零售价) = Format(0 + dbl比率 * db包装系数, mFMT.FM_零售价)
                    End Select
                End If

                If Val(.TextMatrix(Row, menuPriceCol.现零售价)) > Val(.TextMatrix(Row, menuPriceCol.现指导售价)) And Val(.TextMatrix(Row, menuPriceCol.现指导售价)) <> 0 Then
                    .TextMatrix(Row, menuPriceCol.现零售价) = Format(Val(.TextMatrix(Row, menuPriceCol.现指导售价)), mFMT.FM_零售价)
                End If
            End If

            Call GetDrugStore(lngDrugID, Row)
            If Row = .Rows - 1 Then '最后一行才新增行
                .Rows = .Rows + 1
                .RowHeight(.Rows - 1) = mconlngRowHeight
                Row = Row + 1
            End If
        End With

        rsReturn.MoveNext
    Next
    Call setColEdit
    mstr调整额 = ""
    mdbl比率 = 0
    Exit Sub
ErrHandle:
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
    Dim i As Long, n As Long
    Dim dbl发票金额 As Double
    Dim str药品名称 As String
    Dim str发票 As String
    Dim str发票日期 As String
    Dim rsPirce As ADODB.Recordset
    Dim rsCost As ADODB.Recordset
    Dim dbl包装换算 As Double
    Dim bln相同药品 As Boolean
    Dim lng材料ID As Long
    Dim str单位 As String
    Dim dbl比率 As Double

    '功能：为库存列表填充数据
    '参数：材料id

    On Error GoTo ErrHandle

    '先检查是否有重复的数据，如果有就先清除掉重复的数据
    With vsfStore
        For i = .Rows - 1 To 1 Step -1
            If Val(.TextMatrix(i, menuStoreCol.材料ID)) = mlngOldStuffID And mlngOldStuffID <> 0 Then
                .RemoveItem i
            End If
        Next
    End With

    With vsfPay
        For i = .Rows - 1 To 1 Step -1
            If Val(.TextMatrix(i, menuPayCol.材料ID)) = mlngOldStuffID And mlngOldStuffID <> 0 Then
                .RemoveItem i
            End If
        Next
    End With

    If mintModal = 0 Or mblnUpdateAdd = True Or mblnBatchItem = True Then
        gstrSQL = "Select s.库房id,s.药品id as 材料id, d.名称 As 库房, '[' || m.编码 || ']' || m.名称 As 药品, m.规格, m.产地, m.计算单位 售价单位, p.包装单位, s.上次批号 As 批号, s.实际数量 As 数量," & vbNewLine & _
                    "       s.批次, Nvl(m.是否变价, 0) 变价, m.Id," & vbNewLine & _
                    "       Decode(Nvl(m.是否变价, 0), 0, e.现价, Decode(s.零售价,null,Decode(Nvl(s.实际数量, 0), 0, e.现价, s.实际金额 / s.实际数量),s.零售价)) As 时价售价, p.指导差价率 As 差价率,nvl(p.加成率,0) as 加成率," & vbNewLine & _
                    "       Decode(s.平均成本价, null, p.成本价, s.平均成本价) As 成本价, s.上次供应商id, n.名称 As 供应商, s.效期, s.上次产地 As 产地" & vbNewLine & _
                    "From 药品库存 S, 部门表 D, 收费项目目录 M, 材料特性 P, 供应商 N, 收费价目 E" & vbNewLine & _
                    "Where d.Id = s.库房id And s.药品id = m.Id And m.Id = p.材料id And Nvl(s.上次供应商id, 0) = n.Id(+) And m.Id = e.收费细目id And" & vbNewLine & _
                    "      s.性质 = 1 And s.药品id = [1] And Sysdate Between e.执行日期 And e.终止日期  " & vbNewLine & _
                    GetPriceClassString("E") & "Order By 库房, s.上次批号"

        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lngDrugID)

        If mlng供应商ID > 0 Then
            rsTemp.Filter = "上次供应商ID=" & mlng供应商ID
        End If
    Else '修改，查阅
        If mintModal = 2 Then '查阅
            If cboPriceMethod.Text = "仅调成本价" Or cboPriceMethod.Text = "售价成本价一起调价" Then
                gstrSQL = "select (sysdate-执行日期 ) as 是否执行 from 调价汇总记录 where 调价号=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "是否执行", txtNO.Text)
                If rsTemp!是否执行 > 0 Then
                    gstrSQL = "Select Distinct a.库房id, c.名称 As 库房, b.药品id as 材料id, b.供药单位id As 上次供应商id, '[' || e.编码 || ']' || e.名称 As 药品, e.规格, d.名称 As 供应商," & vbNewLine & _
                        "                b.新成本价, b.原成本价, b.发票号, b.发票日期, b.发票金额, b.产地, b.批次, b.批号, e.是否变价 As 变价, e.计算单位 As 售价单位, f.包装单位 As 药库单位," & vbNewLine & _
                        "                a.填写数量 As 数量, f.指导差价率 As 差价率,nvl(f.加成率,0) as 加成率, b.效期" & vbNewLine & _
                        "From 药品收发记录 A,成本价调价信息 B, 部门表 C, 供应商 D, 收费项目目录 E, 材料特性 F" & vbNewLine & _
                        "Where a.id=b.收发id And a.库房id = c.Id And b.供药单位id = d.Id(+) And" & vbNewLine & _
                        "      a.药品id = e.Id And e.Id = f.材料id And b.调价汇总号 = [1] And a.单据 = 18 order by 库房,批号"
                Else
                    gstrSQL = "Select Distinct a.库房id,c.名称 as 库房, b.药品id as 材料id,a.上次供应商id, '[' || e.编码 || ']' ||e.名称 as 药品,e.规格,d.名称 as 供应商, b.新成本价, b.原成本价, b.发票号, b.发票日期, b.发票金额" & _
                            " ,a.上次产地 as 产地,a.批次,a.上次批号 as 批号,e.是否变价 as 变价,e.计算单位 as 售价单位,f.包装单位 as 药库单位,a.实际数量 as 数量,f.指导差价率 as 差价率,nvl(f.加成率,0) as 加成率,a.效期" & _
                            " From 药品库存 A,部门表 C,供应商 D,收费项目目录 E,材料特性 F," & _
                                 " (Select Distinct 药品id, 库房id, 批次, 批号, 效期, 产地, 原成本价, 新成本价, 发票号, 发票日期, 发票金额, 应付款变动, 执行日期" & _
                                   " From 成本价调价信息" & _
                                   " Where 调价汇总号 = [1]) B" & _
                            " Where a.药品id = b.药品id And a.库房id = b.库房id and nvl(a.批次,0)=nvl(b.批次,0) and a.库房id=c.id and a.上次供应商id=d.id(+) and a.药品id=e.id and e.id=f.材料id and a.性质=1 order by 库房,批号"
                End If

            ElseIf cboPriceMethod.Text = "仅调售价" Then
                gstrSQL = "select (sysdate-执行日期 ) as 是否执行 from 调价汇总记录 where 调价号=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "是否执行", txtNO.Text)
                If rsTemp!是否执行 > 0 Then
                    gstrSQL = "Select Distinct a.单据, a.库房id, c.名称 As 库房, b.收费细目id As 材料id, a.供药单位id As 上次供应商id, '[' || e.编码 || ']' || e.名称 As 药品, e.规格," & vbNewLine & _
                        "                d.名称 As 供应商, f.成本价 As 新成本价, f.成本价 As 原成本价, '' 发票号, '' 发票日期, '' 发票金额, a.产地, a.批次, a.批号, e.是否变价 As 变价," & vbNewLine & _
                        "                e.计算单位 As 售价单位, f.包装单位 As 药库单位, a.填写数量 As 数量, f.指导差价率 As 差价率, nvl(f.加成率,0) as 加成率,a.效期" & vbNewLine & _
                        "From 药品收发记录 A, 收费价目 B, 部门表 C, 供应商 D, 收费项目目录 E, 材料特性 F" & vbNewLine & _
                        "Where a.价格id = b.Id And a.库房id = c.Id And a.供药单位id = d.Id(+) And a.药品id = e.Id And e.Id = f.材料id And" & vbNewLine & _
                        "      b.调价汇总号 = [1] And 单据 = 13 " & GetPriceClassString("B") & "order by 库房,批号"
                Else
                    gstrSQL = "Select Distinct a.库房id, c.名称 As 库房, b.收费细目id As 材料id, a.上次供应商id, '[' || e.编码 || ']' || e.名称 As 药品, e.规格, d.名称 As 供应商," & _
                                            " a.平均成本价 As 新成本价, a.平均成本价 As 原成本价, '' 发票号, '' 发票日期, '' 发票金额, a.上次产地 As 产地, a.批次, a.上次批号 As 批号," & _
                                            " e.是否变价 As 变价, e.计算单位 As 售价单位, f.包装单位 as 药库单位, a.实际数量 As 数量, f.指导差价率 As 差价率, nvl(f.加成率,0) as 加成率,a.效期" & _
                            " From 药品库存 A, 收费价目 B, 部门表 C, 供应商 D, 收费项目目录 E, 材料特性 F" & _
                            " Where a.药品id = b.收费细目id And a.库房id = c.Id And a.上次供应商id = d.Id(+) And a.药品id = e.Id And e.Id = f.材料id And a.性质 = 1 And" & _
                                  " b.调价汇总号 = [1]" & GetPriceClassString("B") & " order by 库房,批号"
                End If
            End If
        Else '修改
            If cboPriceMethod.Text = "仅调成本价" Or cboPriceMethod.Text = "售价成本价一起调价" Then
                gstrSQL = "Select Distinct a.库房id,c.名称 as 库房, b.药品id as 材料id,a.上次供应商id, '[' || e.编码 || ']' ||e.名称 as 药品,e.规格,d.名称 as 供应商, b.新成本价, b.原成本价, b.发票号, b.发票日期, b.发票金额" & _
                            " ,a.上次产地 as 产地,a.批次,a.上次批号 as 批号,e.是否变价 as 变价,e.计算单位 as 售价单位,f.包装单位 as 药库单位,a.实际数量 as 数量,f.指导差价率 as 差价率,nvl(f.加成率,0) as 加成率,a.效期" & _
                            " From 药品库存 A,部门表 C,供应商 D,收费项目目录 E,材料特性 F," & _
                                 " (Select Distinct 药品id, 库房id, 批次, 批号, 效期, 产地, 原成本价, 新成本价, 发票号, 发票日期, 发票金额, 应付款变动, 执行日期" & _
                                   " From 成本价调价信息" & _
                                   " Where 调价汇总号 = [1]) B" & _
                            " Where a.药品id = b.药品id And a.库房id = b.库房id and nvl(a.批次,0)=nvl(b.批次,0) and a.库房id=c.id and a.上次供应商id=d.id(+) and a.药品id=e.id and e.id=f.材料id and a.性质=1 order by 库房,批号"
            ElseIf cboPriceMethod.Text = "仅调售价" Then
                gstrSQL = "Select Distinct a.库房id, c.名称 As 库房, b.收费细目id As 材料id, a.上次供应商id, '[' || e.编码 || ']' || e.名称 As 药品, e.规格, d.名称 As 供应商," & _
                                            " a.平均成本价 As 新成本价, a.平均成本价 As 原成本价, '' 发票号, '' 发票日期, '' 发票金额, a.上次产地 As 产地, a.批次, a.上次批号 As 批号," & _
                                            " e.是否变价 As 变价, e.计算单位 As 售价单位, f.包装单位 as 药库单位, a.实际数量 As 数量, f.指导差价率 As 差价率, nvl(f.加成率,0) as 加成率,a.效期" & _
                            " From 药品库存 A, 收费价目 B, 部门表 C, 供应商 D, 收费项目目录 E, 材料特性 F" & _
                            " Where a.药品id = b.收费细目id And a.库房id = c.Id And a.上次供应商id = d.Id(+) And a.药品id = e.Id And e.Id = f.材料id And a.性质 = 1 And" & _
                                  " b.调价汇总号 = [1] " & GetPriceClassString("B") & "order by 库房,批号"
            End If
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, txtNO.Text)
    End If
    
    With vsfStore
        Do While Not rsTemp.EOF
            dbl包装换算 = 0
            dbl发票金额 = 0
            dblOldPrice = 0
            dblNewPrice = 0
            For i = 0 To vsfPrice.Rows - 1
                If rsTemp!材料ID = vsfPrice.TextMatrix(i, menuPriceCol.材料ID) Then
                    dbl包装换算 = vsfPrice.TextMatrix(i, menuPriceCol.包装系数)
                    dblOldPrice = Val(vsfPrice.TextMatrix(i, menuPriceCol.原零售价))
                    dblNewPrice = Val(vsfPrice.TextMatrix(i, menuPriceCol.现零售价))
                    str单位 = vsfPrice.TextMatrix(i, menuPriceCol.单位)
                    Exit For
                End If
            Next
        
            .Rows = .Rows + 1
            Call setColEdit
            .RowHeight(.Rows - 1) = mconlngRowHeight

            '从空白行开始插入数据
            .TextMatrix(.Rows - 1, menuStoreCol.材料ID) = rsTemp!材料ID
            .TextMatrix(.Rows - 1, menuStoreCol.库房) = rsTemp!库房
            .TextMatrix(.Rows - 1, menuStoreCol.库房ID) = rsTemp!库房ID
            .TextMatrix(.Rows - 1, menuStoreCol.供应商) = zlStr.nvl(rsTemp!供应商, "")
            .TextMatrix(.Rows - 1, menuStoreCol.供应商ID) = IIf(mlng供应商ID > 0, mlng供应商ID, zlStr.nvl(rsTemp!上次供应商id))
            .TextMatrix(.Rows - 1, menuStoreCol.药品) = rsTemp!药品
            str药品名称 = rsTemp!药品

            .TextMatrix(.Rows - 1, menuStoreCol.规格) = rsTemp!规格
            .TextMatrix(.Rows - 1, menuStoreCol.单位) = str单位
            .TextMatrix(.Rows - 1, menuStoreCol.批号) = IIf(IsNull(rsTemp!批号), "", rsTemp!批号)
            .TextMatrix(.Rows - 1, menuStoreCol.效期) = Format(IIf(IsNull(rsTemp!效期), "", rsTemp!效期), "YYYY-MM-DD")
            .TextMatrix(.Rows - 1, menuStoreCol.产地) = IIf(IsNull(rsTemp!产地), "", rsTemp!产地)
            .TextMatrix(.Rows - 1, menuStoreCol.数量) = Format(rsTemp!数量 / dbl包装换算, mFMT.FM_数量)
            .TextMatrix(.Rows - 1, menuStoreCol.包装系数) = dbl包装换算
            .TextMatrix(.Rows - 1, menuStoreCol.批次) = zlStr.nvl(rsTemp!批次, 0)
            .TextMatrix(.Rows - 1, menuStoreCol.变价) = rsTemp!变价


            If mintModal = 0 Or mblnUpdateAdd = True Or mblnBatchItem = True Then
                dblOldCost = IIf(IsNull(rsTemp!成本价), 0, rsTemp!成本价) * dbl包装换算

                If mdbl加成率 > 0 Then
                    dbl加成率 = Round(mdbl加成率 / 100, 7)
                ElseIf dblOldCost > 0 Then
                    dbl加成率 = Round(IIf(rsTemp!变价 = 1, rsTemp!时价售价 * dbl包装换算, dblOldPrice) / dblOldCost - 1, 7)
                Else
                   dbl加成率 = nvl(rsTemp!加成率, 0) / 100
                End If

                If 1 + dbl加成率 = 0 Then
                    dblNewCost = 0
                Else
                    dblNewCost = rsTemp!时价售价 * dbl包装换算 / (1 + dbl加成率)
                End If
                If dbl加成率 = -1 Then dbl加成率 = 0

                 .TextMatrix(.Rows - 1, menuStoreCol.原零售价) = Format(IIf(rsTemp!变价 = 1, rsTemp!时价售价 * dbl包装换算, dblOldPrice), mFMT.FM_零售价)

                n = n + 1
                If (mbln时价卫材按批次调价 = False Or Val(.TextMatrix(.Rows - 1, menuStoreCol.变价)) = 0) _
                                                                                    And n <> 1 And mstr调整额 <> "" And mint调价 <> 1 Then
                    .TextMatrix(.Rows - 1, menuStoreCol.现零售价) = .TextMatrix(.Rows - 2, menuStoreCol.现零售价)
                Else
                    If mstr调整额 = "" Or mint调价 = 1 Then
                        .TextMatrix(.Rows - 1, menuStoreCol.现零售价) = Format(IIf(rsTemp!变价 = 1, rsTemp!时价售价 * dbl包装换算, dblOldPrice), mFMT.FM_零售价)
                    Else
                        Select Case mintType
                            Case 1      '根据成本价加成
                                dbl比率 = 1 + Val(mdbl比率) / 100
                                .TextMatrix(.Rows - 1, menuStoreCol.现零售价) = Format(dblOldCost * dbl比率, mFMT.FM_零售价)
                            Case 2      '根据零售价按比例
                                dbl比率 = 1 + Val(mdbl比率) / 100
                                .TextMatrix(.Rows - 1, menuStoreCol.现零售价) = Format(IIf(rsTemp!变价 = 1, rsTemp!时价售价 * dbl比率 * dbl包装换算, dblOldPrice * dbl比率), mFMT.FM_零售价)
                            Case 3      '根据零售价按固定金额加减
                                dbl比率 = Val(mdbl比率)
                                .TextMatrix(.Rows - 1, menuStoreCol.现零售价) = Format(IIf(rsTemp!变价 = 1, rsTemp!时价售价 * dbl包装换算 + dbl比率, dblOldPrice + dbl比率), mFMT.FM_零售价)
                        End Select
                    End If
                End If
                 
                 .TextMatrix(.Rows - 1, menuStoreCol.调整金额) = Format(rsTemp!数量 / dbl包装换算 * (Val(.TextMatrix(.Rows - 1, menuStoreCol.现零售价)) - Val(.TextMatrix(.Rows - 1, menuStoreCol.原零售价))), mFMT.FM_金额)
                 .TextMatrix(.Rows - 1, menuStoreCol.加成率) = GetFormat(dbl加成率 * 100, 2)
                 .TextMatrix(.Rows - 1, menuStoreCol.原采购价) = Format(dblOldCost, mFMT.FM_成本价)
                 .TextMatrix(.Rows - 1, menuStoreCol.现采购价) = Format(dblNewCost, mFMT.FM_成本价)
                 .TextMatrix(.Rows - 1, menuStoreCol.差价差) = Format((Val(.TextMatrix(.Rows - 1, menuStoreCol.现采购价)) - Val(.TextMatrix(.Rows - 1, menuStoreCol.原采购价))) * Val(.TextMatrix(.Rows - 1, menuStoreCol.数量)), mFMT.FM_金额)
                 dbl发票金额 = dbl发票金额 + (dblNewCost - dblOldCost) * Val(.TextMatrix(.Rows - 1, menuStoreCol.数量))
                 
                 Call RefreshPayData("", dbl发票金额)
            Else
                If mintModal = 2 And (cboPriceMethod.Text = "仅调售价" Or cboPriceMethod.Text = "售价成本价一起调价") Then   '查阅
                    gstrSQL = "Select a.成本价 As 原价, a.零售价 As 现价" & vbNewLine & _
                        "From 药品收发记录 A, 收费价目 B" & vbNewLine & _
                        "Where a.价格id = b.Id And b.调价汇总号 = [1] And a.库房id = [2] And a.药品id = [3] And Nvl(a.批次, 0) = [4]" & _
                        GetPriceClassString("B")
                        
                    Set rsPirce = zlDatabase.OpenSQLRecord(gstrSQL, "获取售价", txtNO.Text, rsTemp!库房ID, rsTemp!材料ID, zlStr.nvl(rsTemp!批次, 0))
                    
                    If Not rsPirce.EOF Then
                        .TextMatrix(.Rows - 1, menuStoreCol.原零售价) = Format(Val(rsPirce!原价) * dbl包装换算, mFMT.FM_零售价)
                        .TextMatrix(.Rows - 1, menuStoreCol.现零售价) = Format(Val(rsPirce!现价) * dbl包装换算, mFMT.FM_零售价)
                        .TextMatrix(.Rows - 1, menuStoreCol.调整金额) = Format(rsTemp!数量 / dbl包装换算 * (Val(.TextMatrix(.Rows - 1, menuStoreCol.现零售价)) - Val(.TextMatrix(.Rows - 1, menuStoreCol.原零售价))), mFMT.FM_金额)
                    Else
                        .TextMatrix(.Rows - 1, menuStoreCol.原零售价) = Format(dblOldPrice, mFMT.FM_零售价)
                        .TextMatrix(.Rows - 1, menuStoreCol.现零售价) = Format(dblNewPrice, mFMT.FM_零售价)
                        .TextMatrix(.Rows - 1, menuStoreCol.调整金额) = Format(rsTemp!数量 / dbl包装换算 * (dblNewPrice - IIf(rsTemp!变价 = 1, dblNewPrice * dbl包装换算, dblOldPrice)), mFMT.FM_金额)
                    End If
                    If cboPriceMethod.Text = "仅调售价" Then
                        gstrSQL = "Select 成本价" & vbNewLine & _
                                    "      From (Select 平均成本价 As 成本价" & vbNewLine & _
                                    "             From 药品库存" & vbNewLine & _
                                    "             Where 性质=1 And 库房id = [1] And 药品id = [2] And nvl(批次,0) = [3]" & vbNewLine & _
                                    "             Union All" & vbNewLine & _
                                    "             Select 成本价 From 材料特性 Where 材料id = [2])" & vbNewLine & _
                                    "      Where Rownum <= 1"

                        Set rsCost = zlDatabase.OpenSQLRecord(gstrSQL, "获取成本价", rsTemp!库房ID, rsTemp!材料ID, nvl(rsTemp!批次, 0))
                        .TextMatrix(.Rows - 1, menuStoreCol.原采购价) = Format(rsCost!成本价 * dbl包装换算, mFMT.FM_成本价)
                        .TextMatrix(.Rows - 1, menuStoreCol.现采购价) = Format(rsCost!成本价 * dbl包装换算, mFMT.FM_成本价)
                        .TextMatrix(.Rows - 1, menuStoreCol.差价差) = Format(0, mFMT.FM_金额)
                    Else
                        .TextMatrix(.Rows - 1, menuStoreCol.原采购价) = Format(rsTemp!原成本价 * dbl包装换算, mFMT.FM_成本价)
                        .TextMatrix(.Rows - 1, menuStoreCol.现采购价) = Format(rsTemp!新成本价 * dbl包装换算, mFMT.FM_成本价)
                        .TextMatrix(.Rows - 1, menuStoreCol.差价差) = Format((rsTemp!新成本价 * dbl包装换算 - rsTemp!原成本价 * dbl包装换算) * Val(.TextMatrix(.Rows - 1, menuStoreCol.数量)), mFMT.FM_金额)
                    End If
                Else '修改或者成本价调价
                    '定价直接从收费价目取现价，时价优先从库存取，如果没有则从收费价目取
                    If nvl(rsTemp!变价, 0) = 1 Then
                        gstrSQL = "Select Nvl(s.零售价, Decode(Nvl(s.实际数量, 0), 0, 0, Nvl(s.实际金额, 0) / s.实际数量)) 时价售价" & vbNewLine & _
                        "From 药品库存 S" & vbNewLine & _
                        "Where s.性质=1 And s.库房id = [1] And s.药品id = [2] And nvl(s.批次,0) = [3]"
                        
                        Set rsPirce = zlDatabase.OpenSQLRecord(gstrSQL, "获取售价", rsTemp!库房ID, rsTemp!材料ID, nvl(rsTemp!批次, 0))
                        If rsPirce.RecordCount > 0 Then
                            If rsPirce!时价售价 > 0 Then
                                .TextMatrix(.Rows - 1, menuStoreCol.原零售价) = Format(rsPirce!时价售价 * dbl包装换算, mFMT.FM_零售价)
                                .TextMatrix(.Rows - 1, menuStoreCol.现零售价) = Format(rsPirce!时价售价 * dbl包装换算, mFMT.FM_零售价)
                            Else
                                .TextMatrix(.Rows - 1, menuStoreCol.原零售价) = Format(dblOldPrice, mFMT.FM_零售价)
                                .TextMatrix(.Rows - 1, menuStoreCol.现零售价) = Format(dblNewPrice, mFMT.FM_零售价)
                            End If
                        Else
                            .TextMatrix(.Rows - 1, menuStoreCol.原零售价) = Format(dblOldPrice, mFMT.FM_零售价)
                            .TextMatrix(.Rows - 1, menuStoreCol.现零售价) = Format(dblNewPrice, mFMT.FM_零售价)
                        End If
                    Else
                        .TextMatrix(.Rows - 1, menuStoreCol.原零售价) = Format(dblOldPrice, mFMT.FM_零售价)
                        .TextMatrix(.Rows - 1, menuStoreCol.现零售价) = Format(dblNewPrice, mFMT.FM_零售价)
                    End If
                    .TextMatrix(.Rows - 1, menuStoreCol.调整金额) = Format(rsTemp!数量 / dbl包装换算 * (dblNewPrice - IIf(rsTemp!变价 = 1, dblNewPrice * dbl包装换算, dblOldPrice)), mFMT.FM_金额)
                    .TextMatrix(.Rows - 1, menuStoreCol.原采购价) = Format(rsTemp!原成本价 * dbl包装换算, mFMT.FM_成本价)
                    .TextMatrix(.Rows - 1, menuStoreCol.现采购价) = Format(rsTemp!新成本价 * dbl包装换算, mFMT.FM_成本价)
                    .TextMatrix(.Rows - 1, menuStoreCol.差价差) = Format((Val(.TextMatrix(.Rows - 1, menuStoreCol.现采购价)) - Val(.TextMatrix(.Rows - 1, menuStoreCol.原采购价))) * Val(.TextMatrix(.Rows - 1, menuStoreCol.数量)), mFMT.FM_金额)
                End If
                 
                If cboPriceMethod.Text = "仅调成本价" Or cboPriceMethod.Text = "售价成本价一起调价" Then
                    If rsTemp!原成本价 = 0 Then
                        dbl加成率 = 0
                    Else
                        dbl加成率 = Round(dblNewPrice / (rsTemp!新成本价 * dbl包装换算) - 1, 7)
                    End If
                    .TextMatrix(.Rows - 1, menuStoreCol.加成率) = GetFormat(dbl加成率 * 100, 2)
                    .TextMatrix(.Rows - 1, menuStoreCol.原采购价) = Format(rsTemp!原成本价 * dbl包装换算, mFMT.FM_成本价)
                    .TextMatrix(.Rows - 1, menuStoreCol.现采购价) = Format(rsTemp!新成本价 * dbl包装换算, mFMT.FM_成本价)
                    .TextMatrix(.Rows - 1, menuStoreCol.差价差) = Format((Val(.TextMatrix(.Rows - 1, menuStoreCol.现采购价)) - Val(.TextMatrix(.Rows - 1, menuStoreCol.原采购价))) * Val(.TextMatrix(.Rows - 1, menuStoreCol.数量)), mFMT.FM_金额)
                    dbl发票金额 = dbl发票金额 + (Val(.TextMatrix(.Rows - 1, menuStoreCol.现采购价)) - Val(.TextMatrix(.Rows - 1, menuStoreCol.原采购价))) * Val(.TextMatrix(.Rows - 1, menuStoreCol.数量))
                    str发票 = IIf(IsNull(rsTemp!发票号), "", rsTemp!发票号)
                    str发票日期 = IIf(IsNull(rsTemp!发票日期), "", rsTemp!发票日期)
                    
                    Call RefreshPayData(str发票, dbl发票金额)
                End If
            End If

            rsTemp.MoveNext
        Loop
    End With
    
    '修改和查阅时重算规格列表平均成本价，售价
    'mintModal 0-新增 1-修改 2-查阅
    If mintModal = 1 Or mintModal = 2 Then
        With vsfStore
            For i = 1 To .Rows - 1
                If lng材料ID <> .TextMatrix(i, menuStoreCol.材料ID) Then
                    Call CaluateAverCost(Val(.TextMatrix(i, menuStoreCol.材料ID)))
                    Call CaluateAverOldCost(Val(.TextMatrix(i, menuStoreCol.材料ID)))
                    Call CaculateAverPirce(Val(.TextMatrix(i, menuStoreCol.材料ID)))
                    Call CaculateAverOldPirce(Val(.TextMatrix(i, menuStoreCol.材料ID)))
                    lng材料ID = Val(.TextMatrix(i, menuStoreCol.材料ID))
                End If
            Next
        End With
    End If

    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function RefreshPayData(ByVal str发票号 As String, ByVal str发票日期 As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:重新获取应付情况变动数据
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, dbl发票金额 As Double
    Dim lng供应商ID As Long, lng材料ID As Long, blnData As Boolean

    err = 0: On Error GoTo ErrHand:
    If cboPriceMethod.Text = "仅调成本价" Or cboPriceMethod.Text = "售价成本价一起调价" Then
        TabCtlDetails.Item(1).Visible = True
        chkAutoPay.Visible = True
        chkAutoPay.Value = 1
    End If
    
    If chkAutoPay.Value <> 1 Then RefreshPayData = True: Exit Function

    With vsfPay
        .Rows = 2
        .RowHeight(.Rows - 1) = mconlngRowHeight
        .Clear 1
    End With

    With vsfStore
        For i = 1 To .Rows - 1
            lng供应商ID = Val(.TextMatrix(i, menuStoreCol.供应商ID))

            lng材料ID = Val(.TextMatrix(i, menuStoreCol.材料ID))

            If lng供应商ID <> 0 And lng材料ID <> 0 Then
                dbl发票金额 = Val(.TextMatrix(i, menuStoreCol.差价差))
'                If dbl发票金额 <> 0 Then
                    '先找相关的供应商是否存在
                    With vsfPay
                        blnData = False
                        For j = 1 To .Rows - 1
                            If lng材料ID = Val(.TextMatrix(j, menuPayCol.材料ID)) And _
                               lng供应商ID = Val(.TextMatrix(j, menuPayCol.供应商ID)) Then
                                .TextMatrix(j, menuPayCol.发票金额) = Format(Val(.TextMatrix(j, menuPayCol.发票金额)) + dbl发票金额, mFMT.FM_金额)
                               blnData = True
                               Exit For
                            End If
                        Next
                        If blnData = False Then
                            '没有此供应商或材料,因此需要额外增加
                            If Val(.TextMatrix(.Rows - 1, menuPayCol.供应商ID)) <> 0 Then
                                .Rows = .Rows + 1
                                .RowHeight(.Rows - 1) = mconlngRowHeight
                                Call setColEdit
                            End If
                            .TextMatrix(.Rows - 1, menuPayCol.供应商) = vsfStore.TextMatrix(i, menuStoreCol.供应商)
                            .TextMatrix(.Rows - 1, menuPayCol.供应商ID) = vsfStore.TextMatrix(i, menuStoreCol.供应商ID)
                            .TextMatrix(.Rows - 1, menuPayCol.材料ID) = vsfStore.TextMatrix(i, menuStoreCol.材料ID)
                            .TextMatrix(.Rows - 1, menuPayCol.品名) = vsfStore.TextMatrix(i, menuStoreCol.药品)
                            .TextMatrix(.Rows - 1, menuPayCol.规格) = vsfStore.TextMatrix(i, menuStoreCol.规格)
                            .TextMatrix(.Rows - 1, menuPayCol.产地) = vsfStore.TextMatrix(i, menuStoreCol.产地)
                            .TextMatrix(.Rows - 1, menuPayCol.发票号) = str发票号
                            If str发票日期 <> "" Then
                                .TextMatrix(.Rows - 1, menuPayCol.发票日期) = Format(str发票日期, "YYYY-MM-DD")
                            End If
                            .TextMatrix(.Rows - 1, menuPayCol.发票金额) = Format(dbl发票金额, mFMT.FM_金额)
                        End If
                    End With
'                End If
            End If
        Next
    End With

    RefreshPayData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub vsfPrice_DblClick()
    With vsfPrice
        If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mconlngCanColor Then
            .EditCell
            .EditSelStart = 0
            .EditSelLength = Len(.EditText)
        End If
    End With
End Sub

Private Sub vsfPrice_EnterCell()
    Dim i As Integer

    With vsfPrice
        .Editable = flexEDNone
        If .CellBackColor = mconlngColor Then
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
        End If
        If .Col = menuPriceCol.现成本价 Then
            mdblOldPrice = Val(.TextMatrix(.Row, menuPriceCol.现成本价))
        ElseIf .Col = menuPriceCol.现零售价 Then
            mdblOldPrice = Val(.TextMatrix(.Row, menuPriceCol.现零售价))
        End If
    End With

    With vsfStore
        If Val(vsfPrice.TextMatrix(vsfPrice.Row, menuPriceCol.材料ID)) = 0 Then Exit Sub

        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If Val(vsfPrice.TextMatrix(vsfPrice.Row, menuPriceCol.材料ID)) = Val(.TextMatrix(i, menuStoreCol.材料ID)) Then
                    .Select i, 0, i, .Cols - 1
                    .TopRow = i
                End If
            Next
        End If
    End With
End Sub

Private Sub vsfPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intRow As Integer
    Dim intCol As Integer
    Dim lngDrugID As Long
    Dim strRow As String

    With vsfPrice
        If KeyCode = vbKeyReturn Then
            If .Col <> menuPriceCol.现零售价 Then
                If .Col = menuPriceCol.品名 And cboPriceMethod.Text = "仅调成本价" Then
                    .Col = menuPriceCol.现成本价
'                    .EditCell
                ElseIf .Col = menuPriceCol.品名 And cboPriceMethod.Text = "仅调售价" Then
                    .Col = menuPriceCol.现零售价
'                    .EditCell
                ElseIf .Col = menuPriceCol.现成本价 And cboPriceMethod.Text = "仅调成本价" Then
                    If .Row = .Rows - 1 And Val(.TextMatrix(.Row, menuPriceCol.材料ID)) <> 0 Then
                        .Rows = .Rows + 1
                        .Row = .Row + 1
                        .Col = menuPriceCol.品名
                        .RowHeight(.Rows - 1) = mconlngRowHeight
'                        .EditCell
                        Call setColEdit
                    ElseIf Val(.TextMatrix(.Row, menuPriceCol.材料ID)) <> 0 Then
                        .ColComboList(menuPriceCol.品名) = ""
                        .Row = .Row + 1
                        .Col = menuPriceCol.品名
                    End If
                ElseIf .Col = menuPriceCol.品名 And cboPriceMethod.Text = "售价成本价一起调价" Then
                    .Col = menuPriceCol.现成本价
'                    .EditCell
                ElseIf .Col = menuPriceCol.现成本价 And cboPriceMethod.Text = "售价成本价一起调价" Then
                    .Col = menuPriceCol.现零售价
'                    .EditCell
                ElseIf .Col = menuPriceCol.现零售价 And cboPriceMethod.Text = "售价成本价一起调价" Then
                    If .Row = .Rows - 1 Then
                        .Rows = .Rows + 1
                        .Row = .Row + 1
                        .Col = menuPriceCol.品名
                        .RowHeight(.Rows - 1) = mconlngRowHeight
'                        .EditCell
                        Call setColEdit
                    ElseIf Val(.TextMatrix(.Row, menuPriceCol.材料ID)) <> 0 Then
                        .ColComboList(menuPriceCol.品名) = ""
                        .Row = .Row + 1
                        .Col = menuPriceCol.品名
'                        .EditCell
                    End If
                Else
                    .Col = .Col + 1
'                    .EditCell
                End If
            Else
                If Val(.TextMatrix(.Row, menuPriceCol.材料ID)) <> 0 And .Row = .Rows - 1 Then
                    .ColComboList(menuPriceCol.品名) = ""
                    .Rows = .Rows + 1
                    .Row = .Row + 1
                    .Col = menuPriceCol.品名
                    .RowHeight(.Rows - 1) = mconlngRowHeight
'                    .EditCell
                    Call setColEdit
                ElseIf Val(.TextMatrix(.Row, menuPriceCol.材料ID)) <> 0 Then
                    .ColComboList(menuPriceCol.品名) = ""
                    .Row = .Row + 1
                    .Col = menuPriceCol.品名
'                    .EditCell
                End If
            End If
        ElseIf KeyCode = vbKeyDelete Then
            lngDrugID = Val(vsfPrice.TextMatrix(vsfPrice.Row, menuPriceCol.材料ID))

            If .Rows > 2 Then
                mdbl比率 = 0
                mstr调整额 = ""
                .RemoveItem .Row
            Else
                For intCol = 0 To .Cols - 1
                    .TextMatrix(.Row, intCol) = ""
                Next
            End If

            With vsfStore
                If lngDrugID = 0 Then Exit Sub
                For intRow = .Rows - 1 To 1 Step -1
                    If Val(.TextMatrix(intRow, menuStoreCol.材料ID)) = lngDrugID Then
                        .RemoveItem intRow
                    End If
                Next
            End With

            With vsfPay
                If lngDrugID = 0 Then Exit Sub
                For intRow = .Rows - 1 To 1 Step -1
                    If Val(.TextMatrix(intRow, menuPayCol.材料ID)) = lngDrugID Then
                        .RemoveItem intRow
                    End If
                Next
            End With
        End If
    End With
End Sub

Private Sub vsfPrice_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim rsReturn As Recordset
    Dim strKey As String

    On Error GoTo ErrHandle
    If KeyCode <> vbKeyReturn Then Exit Sub

    With vsfPrice
        strKey = .EditText
        Select Case Col
        Case menuPriceCol.品名
            mblnUpdateAdd = True
            Set rsReturn = SelectStuff(strKey)
            If rsReturn Is Nothing Then Exit Sub
            If rsReturn.RecordCount = 0 Then Exit Sub
            Call GetDrugPirce(rsReturn, Row)
            mblnUpdateAdd = False
        End Select
    End With

    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CheckDoubleDrug(ByVal rsTemp As ADODB.Recordset) As ADODB.Recordset
    '检查是否有重复的药品
    'lngDrugId 材料id
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
            For j = 1 To .Rows - 1
                If Val(.TextMatrix(j, menuPriceCol.材料ID)) = rsTemp!Id Then
                    strTemp = strTemp & " id <> " & rsTemp!Id & " and "
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
        MsgBox strName & "等" & intCount & "种卫材在列表中已经存在，已存在卫材不再添加！", vbInformation, gstrSysName
    End If

    Set CheckDoubleDrug = rsTemp
End Function

Private Sub vsfPrice_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        With vsfPrice
            If .Col = menuPriceCol.品名 Then
                .Editable = flexEDKbdMouse
                Exit Sub
            End If
            If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mconlngCanColor Then
                .Editable = flexEDKbdMouse
            Else
                .Editable = flexEDNone
            End If
        End With
    End If
End Sub

Private Sub vsfPrice_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strKey As String
    Dim intDigit As Integer

    With vsfPrice
        strKey = .EditText
        If .Col = menuPriceCol.现成本价 Then
            mdbl成本价 = Val(.TextMatrix(Row, Col))
        End If
    End With

    If Col = menuPriceCol.现成本价 Or Col = menuPriceCol.现零售价 Then
        If KeyAscii = vbKeyReturn Then Exit Sub
        If KeyAscii <> vbKeyBack Then
            Select Case Col
                Case menuPriceCol.现成本价
                    intDigit = Len(Mid(mFMT.FM_成本价, InStr(1, mFMT.FM_成本价, ".") + 1))
                Case menuPriceCol.现零售价
                    intDigit = Len(Mid(mFMT.FM_零售价, InStr(1, mFMT.FM_零售价, ".") + 1))
            End Select

            If KeyAscii = vbKeyDelete Then
                If InStr(1, strKey, ".") > 0 Then
                    KeyAscii = 0
                End If
            ElseIf KeyAscii = Asc(".") Or (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
                If vsfPrice.EditSelLength = Len(strKey) Then Exit Sub
                If InStr(strKey, ".") <> 0 And Chr(KeyAscii) = "." Then   '只能存在一个小数点
                    KeyAscii = 0
                    Exit Sub
                End If
                If Len(Mid(strKey, InStr(1, strKey, ".") + 1)) >= intDigit And strKey Like "*.*" Then
                    KeyAscii = 0
                    Exit Sub
                Else
                    Exit Sub
                End If
            Else
                KeyAscii = 0
            End If
        End If
    ElseIf Col = menuPriceCol.品名 Then
        If InStr("`~!@#$%^&*()_-+={[}]|\:;""'<,>.?/", Chr(KeyAscii)) > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub vsfPrice_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If Col = menuPriceCol.品名 Then
        vsfPrice.ColComboList(menuPriceCol.品名) = "|..."
    End If
End Sub

Private Sub setColEdit()
    '功能：设置列是否可以修改
    '不能修改的列颜色为灰色，能修改的列颜色为白色
    Dim intCol As Integer
    Dim intRow As Integer

    With vsfPrice
        .Cell(flexcpBackColor, 1, 1, .Rows - 1, .Cols - 1) = mconlngColor
        If cboPriceMethod.Text = "仅调售价" Then
            .Cell(flexcpBackColor, 1, menuPriceCol.品名, .Rows - 1, menuPriceCol.品名) = mconlngCanColor
            .Cell(flexcpBackColor, 1, menuPriceCol.现零售价, .Rows - 1, menuPriceCol.现零售价) = mconlngCanColor
        ElseIf cboPriceMethod.Text = "仅调成本价" Then
            .Cell(flexcpBackColor, 1, menuPriceCol.品名, .Rows - 1, menuPriceCol.品名) = mconlngCanColor
            .Cell(flexcpBackColor, 1, menuPriceCol.现成本价, .Rows - 1, menuPriceCol.现成本价) = mconlngCanColor
        Else
            .Cell(flexcpBackColor, 1, menuPriceCol.品名, .Rows - 1, menuPriceCol.品名) = mconlngCanColor
            .Cell(flexcpBackColor, 1, menuPriceCol.现成本价, .Rows - 1, menuPriceCol.现成本价) = mconlngCanColor
            .Cell(flexcpBackColor, 1, menuPriceCol.现零售价, .Rows - 1, menuPriceCol.现零售价) = mconlngCanColor
        End If

    End With

    With vsfStore
        If .Rows = 1 Then Exit Sub
        .Cell(flexcpBackColor, 1, 0, .Rows - 1, .Cols - 1) = mconlngColor
        If cboPriceMethod.Text = "仅调售价" Then
            .Cell(flexcpBackColor, 1, 0, .Rows - 1, .Cols - 1) = mconlngColor
        ElseIf cboPriceMethod.Text = "仅调成本价" Then
            .Cell(flexcpBackColor, 1, menuStoreCol.加成率, .Rows - 1, menuStoreCol.加成率) = mconlngCanColor
            .Cell(flexcpBackColor, 1, menuStoreCol.现采购价, .Rows - 1, menuStoreCol.现采购价) = mconlngCanColor
        Else
            .Cell(flexcpBackColor, 1, menuStoreCol.加成率, .Rows - 1, menuStoreCol.加成率) = mconlngCanColor
            .Cell(flexcpBackColor, 1, menuStoreCol.现采购价, .Rows - 1, menuStoreCol.现采购价) = mconlngCanColor
            .Cell(flexcpBackColor, 1, menuStoreCol.现零售价, .Rows - 1, menuStoreCol.现零售价) = mconlngCanColor
        End If
        If .Rows > 1 Then
            For intRow = 1 To .Rows - 1
                If Val(.TextMatrix(intRow, menuStoreCol.变价)) = 1 And mbln时价卫材按批次调价 = True And mint调价 <> 1 Then
                    .Cell(flexcpBackColor, intRow, menuStoreCol.现零售价, intRow, menuStoreCol.现零售价) = mconlngCanColor
                Else
                    .Cell(flexcpBackColor, intRow, menuStoreCol.现零售价, intRow, menuStoreCol.现零售价) = mconlngColor
                End If
            Next
        End If
    End With

    With vsfPay
        If .Rows = 1 Then Exit Sub
        .Cell(flexcpBackColor, 1, 0, .Rows - 1, .Cols - 1) = mconlngColor
        .Cell(flexcpBackColor, 1, menuPayCol.发票号, .Rows - 1, menuPayCol.发票号) = mconlngCanColor
        .Cell(flexcpBackColor, 1, menuPayCol.发票日期, .Rows - 1, menuPayCol.发票日期) = mconlngCanColor
        .Cell(flexcpBackColor, 1, menuPayCol.发票金额, .Rows - 1, menuPayCol.发票金额) = mconlngCanColor
    End With
End Sub

Private Sub vsfPrice_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        vsfPrice.Editable = flexEDNone
        If vsfPrice.Col = menuPriceCol.品名 And mintModal <> 2 Then
            vsfPrice.ColComboList(menuPriceCol.品名) = "|..."
            vsfPrice.Editable = flexEDKbdMouse
        End If
    End If
End Sub

Private Sub vsfPrice_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lngDrugID As Long
    Dim intRow As Integer
    Dim strKey As String

    strKey = Trim(vsfPrice.EditText)
    strKey = Replace(strKey, Chr(vbKeyReturn), "")
    strKey = Replace(strKey, Chr(10), "")

    With vsfPrice
        If .EditText = "" Then Exit Sub
        lngDrugID = Val(.TextMatrix(Row, menuPriceCol.材料ID))
        If lngDrugID = 0 Then Exit Sub

        Select Case Col
            Case menuPriceCol.现成本价
                If Not IsNumeric(strKey) Then
                    Cancel = True
                    Exit Sub
                End If
                If .EditText > 9999999 Then
                    MsgBox "成本价过大，请重新输入！", vbInformation, gstrSysName
                    Cancel = True
                    Exit Sub
                End If
                If mdblOldPrice = .EditText Then Exit Sub '没有做修改时，直接退出

                If strKey <> "" Then
                    If Val(strKey) > Val(.TextMatrix(Row, menuPriceCol.原采购限价)) And Val(.TextMatrix(Row, menuPriceCol.原采购限价)) <> 0 Then
                        If MsgBox("成本价不能大于指导零售价，（" & Format(Val(.TextMatrix(Row, menuPriceCol.原采购限价)), mFMT.FM_成本价) & "），继续吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                            Cancel = True
                            Exit Sub
                        Else
                            vsfPrice.EditText = Format(Val(strKey), mFMT.FM_成本价)
                            vsfPrice.TextMatrix(Row, menuPriceCol.现采购限价) = Format(Val(strKey), mFMT.FM_成本价)
                        End If
                    Else
                        vsfPrice.EditText = Format(Val(strKey), mFMT.FM_成本价)
                    End If
                    If chkAppAllColumn.Value = 1 And mlngPrice <> vsfPrice.EditText Then
                        For intRow = 1 To .Rows - 1
                            If .TextMatrix(intRow, menuPriceCol.材料ID) <> "" Then
                                .TextMatrix(intRow, menuPriceCol.现成本价) = vsfPrice.EditText
                            End If
                        Next
                    End If
                End If
                If chkAppAllColumn.Value = 0 Then
                    Call FullStoce成本价(Val(.TextMatrix(Row, menuPriceCol.材料ID)), vsfPrice.EditText)
                End If
            Case menuPriceCol.现零售价
                If Not IsNumeric(strKey) Then
                    Cancel = True
                    Exit Sub
                End If
                If .EditText > 9999999 Then
                    MsgBox "零售价过大，请重新输入！", vbInformation, gstrSysName
                    Cancel = True
                    Exit Sub
                End If
                If mdblOldPrice = .EditText Then Exit Sub

                If strKey <> "" Then
'                    If zlCommFun.DblIsValid(strkey, 12, , False, , "现价") = False Then Cancel = True: Exit Sub
                    If Val(.TextMatrix(Row, menuPriceCol.材料ID)) = 0 Then
                        vsfPrice.EditText = Format(Val(strKey), mFMT.FM_零售价)
                        Exit Sub
                    End If
                    If Val(strKey) > Val(.TextMatrix(Row, menuPriceCol.原指导售价)) And Val(.TextMatrix(Row, menuPriceCol.原指导售价)) <> 0 Then
                        If MsgBox("现价不能大于指导零售价，（" & Format(Val(.TextMatrix(Row, menuPriceCol.原指导售价)), mFMT.FM_零售价) & "），继续吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                            Cancel = True
                            Exit Sub
                        Else
                            vsfPrice.EditText = Format(Val(strKey), mFMT.FM_零售价)
                            vsfPrice.TextMatrix(Row, menuPriceCol.现指导售价) = Format(Val(strKey), mFMT.FM_零售价)
                        End If
                    Else
                        vsfPrice.EditText = Format(Val(strKey), mFMT.FM_零售价)
                    End If

                End If
                If chkAppAllColumn.Value = 1 And mlngPrice <> vsfPrice.EditText Then
                    For intRow = 1 To .Rows - 1
                        If .TextMatrix(intRow, menuPriceCol.材料ID) <> "" Then
                            .TextMatrix(intRow, menuPriceCol.现零售价) = vsfPrice.EditText
                        End If
                    Next
                End If
                If chkAppAllColumn.Value = 0 Then
                    Call FullStoce现价(Val(.TextMatrix(Row, menuPriceCol.材料ID)), vsfPrice.EditText)
                End If
        End Select
    End With
End Sub

Private Sub FullStoce成本价(ByVal lng材料ID, ByVal dbl成本价 As Double)
    '成本价
    Dim lngRow As Long, dbl调整额 As Double
    With vsfStore
        For lngRow = 1 To .Rows - 1
            If Val(.TextMatrix(lngRow, menuStoreCol.材料ID)) = lng材料ID Then
                .TextMatrix(lngRow, menuStoreCol.现采购价) = Format(dbl成本价, mFMT.FM_成本价)
                 Call AutoCalcStoce(lngRow, menuStoreCol.现采购价)
            End If
        Next
    End With
End Sub

Private Sub FullStoce现价(ByVal lng材料ID As Long, ByVal dbl现价 As Double)
    '-----------------------------------------------------------------------------------------------------------
    '功能:根据现价,填充库存变动的现价及调整额
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-11-07 10:32:13
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, dbl调整额 As Double
    With vsfStore
        For lngRow = 1 To .Rows - 1
            If Val(.TextMatrix(lngRow, menuStoreCol.材料ID)) = lng材料ID Then
                .TextMatrix(lngRow, menuStoreCol.现零售价) = Format(dbl现价, mFMT.FM_零售价)
                '调整额=数量*(现价-原价)
                dbl调整额 = (dbl现价 - Val(.TextMatrix(lngRow, menuStoreCol.原零售价))) * Val(.TextMatrix(lngRow, menuStoreCol.数量))
                .TextMatrix(lngRow, menuStoreCol.调整金额) = Format(dbl调整额, mFMT.FM_金额)
                '需要根据加成率重新计算调整的成本价
'                 Call AutoCalcStoce(lngRow, menuStoreCol.现零售价)
                If Val(.TextMatrix(lngRow, menuStoreCol.现采购价)) <> 0 Then
                    .TextMatrix(lngRow, menuStoreCol.加成率) = Format(Val((.TextMatrix(lngRow, menuStoreCol.现零售价)) / Val(.TextMatrix(lngRow, menuStoreCol.现采购价)) - 1) * 100, "#0.000")
                Else
                    .TextMatrix(lngRow, menuStoreCol.加成率) = 0
                End If
            End If
        Next
    End With
End Sub

Private Sub AutoCalcStoce(ByVal lngEditRow As Long, ByVal lngEditCol As Long)
    '-----------------------------------------------------------------------------------------------------------
    '功能:自动计算相关信息(根据加成率计算现成本价及差额,根据现成本价计算差额及加成率)
    '入参:lngEditRow-当前编辑的行
    '     lngEditCol-当前编辑的列
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-11-06 17:03:02
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, dbl现成本价 As Double, dbl加成率 As Double, dbl成本差价 As Double, dbl差价调整额 As Double
    Dim lng材料ID As Long, bln库房分批 As Boolean, lng供应商ID As Long, lngTemp As Long, i As Long
    Dim blnHaveData As Boolean, lngStep As Long, lngSteps As Long

    err = 0: On Error GoTo ErrHand:
    With vsfStore
        bln库房分批 = chkCostBatch.Value = 1
        lngStep = IIf(bln库房分批, lngEditRow, 1)
        lngSteps = IIf(bln库房分批, lngEditRow, .Rows - 1)
        Select Case lngEditCol
        Case menuStoreCol.加成率
            dbl加成率 = Val(.TextMatrix(lngEditRow, lngEditCol)) / 100
            If dbl加成率 = -1 Then dbl加成率 = 0
            '现成本价=现零售价/(1+加成率)
            dbl现成本价 = Format(Val(.TextMatrix(lngEditRow, menuStoreCol.现零售价)) / (1 + dbl加成率), mFMT.FM_成本价)
            '差价调整额=(原成本价-现成本价)
            dbl成本差价 = dbl现成本价 - Val(.TextMatrix(lngEditRow, menuStoreCol.原采购价))
        Case menuStoreCol.现采购价
            '因为存在包装换算问题，因此，目前按最小单位进行设置单价
            dbl现成本价 = Val(.TextMatrix(lngEditRow, lngEditCol))
            '加成率=现零售价/现成本价-1
            If dbl现成本价 <> 0 Then
                dbl加成率 = Round(Val(.TextMatrix(lngEditRow, menuStoreCol.现零售价)) / dbl现成本价 - 1, 7)
            Else
                dbl加成率 = 0
            End If
            '差价调整额=(现成本价-原成本价)
            dbl成本差价 = Format((dbl现成本价 - Val(.TextMatrix(lngEditRow, menuStoreCol.原采购价))), mFMT.FM_成本价)
        Case menuStoreCol.差价差
            Exit Sub
        Case menuStoreCol.现零售价
            '现价发生改变时,需要重新根据加成率计算相关的现成本价
'            dbl加成率 = Round(Val(.TextMatrix(lngEditRow, menuStoreCol.加成率)) / 100, 7)
'            If dbl加成率 = -1 Then dbl加成率 = 0
'            '现成本价=现零售价/(1+加成率)
'            dbl现成本价 = Format(Val(.TextMatrix(lngEditRow, menuStoreCol.现零售价)) / (1 + dbl加成率), mFMT.FM_成本价)
'            '差价调整额=(现成本价-原成本价)
'            dbl成本差价 = (dbl现成本价 - Val(.TextMatrix(lngEditRow, menuStoreCol.原采购价)))


            '现价改变时，需要重新根据成本价计算相关的加成率
            dbl现成本价 = Val(.TextMatrix(lngEditRow, menuStoreCol.现采购价))
            If dbl现成本价 = 0 Then
                dbl加成率 = 0
            Else
                dbl加成率 = Round(Val(.TextMatrix(lngEditRow, menuStoreCol.现零售价)) / dbl现成本价 - 1, 7)
            End If
            lngStep = lngEditRow
            lngSteps = lngEditRow
        Case Else
            Exit Sub
        End Select

        lng材料ID = Val(.TextMatrix(lngEditRow, menuStoreCol.材料ID))
        lng供应商ID = Val(.TextMatrix(lngEditRow, menuStoreCol.供应商ID))
        Dim cllData As New Collection
        For lngRow = lngStep To lngSteps
            If lng材料ID = Val(.TextMatrix(lngRow, menuStoreCol.材料ID)) Then
                If dbl加成率 = -1 Then dbl加成率 = 0
                .TextMatrix(lngRow, menuStoreCol.加成率) = Format(dbl加成率 * 100, GFM_VBJCL)
                '该成本价是以最小单位为准的，因此要乘小换算系数.
                .TextMatrix(lngRow, menuStoreCol.现采购价) = Format(dbl现成本价, mFMT.FM_成本价)
                dbl成本差价 = dbl现成本价 - Val(.TextMatrix(lngRow, menuStoreCol.原采购价))
                 '差价调整额=(现成本价-原成本价)*数量
                 dbl差价调整额 = Round(dbl成本差价 * Val(.TextMatrix(lngRow, menuStoreCol.数量)), 7)
                .TextMatrix(lngRow, menuStoreCol.差价差) = Format(dbl差价调整额, mFMT.FM_金额)
'                .TextMatrix(lngRow, menuStoreCol.差价差) = dbl差价调整额
                lngTemp = Val(.TextMatrix(lngRow, menuStoreCol.材料ID))
                lng供应商ID = Val(.TextMatrix(lngRow, menuStoreCol.供应商ID))

                If lng供应商ID <> 0 Then
                    err = 0: On Error Resume Next
                    cllData.Add Array(lngTemp, lng供应商ID, dbl差价调整额, .TextMatrix(lngRow, menuStoreCol.供应商), .TextMatrix(lngRow, menuStoreCol.药品), .TextMatrix(lngRow, menuStoreCol.规格), .TextMatrix(lngRow, menuStoreCol.产地)), "K" & lng供应商ID & "_" & lngTemp
                    If err <> 0 Then
                        '累计差价调整额
                        dbl差价调整额 = Val(cllData("K" & lng供应商ID & "_" & lngTemp)(2)) + dbl差价调整额
                        cllData.Remove "K" & lng供应商ID & "_" & lngTemp
                         err = 0: On Error GoTo ErrHand:
                        cllData.Add Array(lngTemp, lng供应商ID, dbl差价调整额, .TextMatrix(lngRow, menuStoreCol.供应商), .TextMatrix(lngRow, menuStoreCol.药品), .TextMatrix(lngRow, menuStoreCol.规格), .TextMatrix(lngRow, menuStoreCol.产地)), "K" & lng供应商ID & "_" & lngTemp

                    End If
                    On Error GoTo ErrHand:
                End If
            End If
        Next
        If chkAutoPay.Value = 1 Then
            '需要自动计算相关的应付变动记录
            For i = 1 To cllData.Count
                With vsfPay
                    blnHaveData = False
                    For lngRow = 1 To .Rows - 1
                        lngTemp = Val(.TextMatrix(lngRow, menuPayCol.材料ID))
                        lng供应商ID = Val(.TextMatrix(lngRow, menuPayCol.供应商ID))
                        If lngTemp = Val(cllData(i)(0)) And lng供应商ID = Val(cllData(i)(1)) Then
                            '卫材及供应商相同,清空相关的值
                            .TextMatrix(lngRow, menuPayCol.发票金额) = Format(Val(cllData(i)(2)), mFMT.FM_金额)
                             blnHaveData = True
                        End If
                    Next
                    If blnHaveData = False Then
                        '需要增加该项供应商的物资
                        If Val(.TextMatrix(.Rows - 1, menuPayCol.材料ID)) <> 0 Or .Rows = 1 Then
                            .Rows = .Rows + 1
                            .RowHeight(.Rows - 1) = mconlngRowHeight
                            Call setColEdit
                        End If
                        lngRow = .Rows - 1
                        .TextMatrix(lngRow, menuPayCol.供应商) = cllData(i)(3)
                        .TextMatrix(lngRow, menuPayCol.供应商ID) = cllData(i)(1)
                        .TextMatrix(lngRow, menuPayCol.材料ID) = cllData(i)(0)
                        .TextMatrix(lngRow, menuPayCol.品名) = cllData(i)(4)
                        .TextMatrix(lngRow, menuPayCol.规格) = cllData(i)(5)
                        .TextMatrix(lngRow, menuPayCol.产地) = cllData(i)(6)
                        .TextMatrix(lngRow, menuPayCol.发票金额) = Format(Val(cllData(i)(2)), mFMT.FM_金额)
                    End If
                End With
            Next
        End If
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
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

    If intRow = 0 Or mint调价 = 1 Then Exit Sub

'    dblOldPrice = Val(vsfPrice.TextMatrix(intRow, menuPriceCol.原零售价))
'    dbl包装 = Val(vsfPrice.TextMatrix(vsfPrice.Row, menuPriceCol.包装系数))
'
'    With vsfStore
'        For n = 1 To .Rows - 1
'            If .TextMatrix(n, 0) <> "" Then
'                If Val(.TextMatrix(n, menuStoreCol.材料id)) = lngDrugID Then
'                    dblNum = Val(.TextMatrix(n, menuStoreCol.数量))
'
'                    .TextMatrix(n, menuStoreCol.现零售价) = format(dblNewPrice, mFMT.FM_零售价)
'                    .TextMatrix(n, menuStoreCol.调整金额) = Format(Val(.TextMatrix(n, menuStoreCol.数量)) * (dblNewPrice - dblOldPrice), mFMT.fm_金额)
'
'                    If mint调价 = 2 And chkAotuCost.Value = 1 Then
'                        dblOldCost = .TextMatrix(n, menuStoreCol.原采购价)
'                        dblNewCost = dblNewPrice / (1 + Round(Val(.TextMatrix(n, menuStoreCol.加成率)) / 100, 7))
'                        .TextMatrix(n, menuStoreCol.现采购价) = format(dblNewCost, mFMT.FM_成本价)
'                        .TextMatrix(n, menuStoreCol.差价差) = Format((dblNewCost - dblOldCost) * dblNum, mFMT.fm_金额)
'                        dbl发票金额 = dbl发票金额 + (dblNewCost - dblOldCost) * dblNum
'                    End If
'                End If
'            End If
'        Next
'    End With
'
'    If chkAutoPay.Value = 1 Then
'        With vsfPay
'            For n = 1 To .Rows - 1
'                If .TextMatrix(1, 0) <> "" Then
'                    If Val(.TextMatrix(n, menuPayCol.材料id)) = lngDrugID Then
'                        .TextMatrix(n, menuPayCol.发票金额) = format(dbl发票金额, 2)
'                    End If
'                End If
'            Next
'        End With
'    End If
'
'    CaluateAverCost lngDrugID
End Sub

Private Sub CaluateAverOldCost(ByVal lng材料ID As Long)
    '计算原始平均成本价
    Dim i As Integer
    Dim dblSumCost As Double
    Dim dblSumNumber As Double

    With vsfStore
        For i = 1 To .Rows - 1
            If .TextMatrix(i, menuStoreCol.材料ID) <> "" Then
                If Val(.TextMatrix(i, menuStoreCol.材料ID)) = lng材料ID Then
                    dblSumCost = dblSumCost + Val(.TextMatrix(i, menuStoreCol.原采购价)) * Val(.TextMatrix(i, menuStoreCol.数量))
                    dblSumNumber = dblSumNumber + Val(.TextMatrix(i, menuStoreCol.数量))
                End If
            End If
        Next
    End With

    With vsfPrice
        If dblSumNumber > 0 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, menuPriceCol.材料ID) <> "" Then
                    If Val(.TextMatrix(i, menuPriceCol.材料ID)) = lng材料ID Then
                        .TextMatrix(i, menuPriceCol.原成本价) = Format(dblSumCost / dblSumNumber, mFMT.FM_成本价)
                        Exit For
                    End If
                End If
            Next
        End If
    End With
End Sub

Private Sub CaluateAverCost(ByVal lng材料ID As Long)
    '计算平均成本价
    Dim i As Integer
    Dim dblSumCost As Double
    Dim dblSumNumber As Double

    With vsfStore
        For i = 1 To .Rows - 1
            If .TextMatrix(i, menuStoreCol.材料ID) <> "" Then
                If Val(.TextMatrix(i, menuStoreCol.材料ID)) = lng材料ID Then
                    dblSumCost = dblSumCost + Val(.TextMatrix(i, menuStoreCol.现采购价)) * Val(.TextMatrix(i, menuStoreCol.数量))
                    dblSumNumber = dblSumNumber + Val(.TextMatrix(i, menuStoreCol.数量))
                End If
            End If
        Next
    End With

    With vsfPrice
        If dblSumNumber > 0 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, menuPriceCol.材料ID) <> "" Then
                    If Val(.TextMatrix(i, menuPriceCol.材料ID)) = lng材料ID Then
                        .TextMatrix(i, menuPriceCol.现成本价) = Format(dblSumCost / dblSumNumber, mFMT.FM_成本价)
                        Exit For
                    End If
                End If
            Next
        End If
    End With
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
            .ColHidden(menuStoreCol.批次) = True
            .ColHidden(menuStoreCol.变价) = True
            .ColHidden(menuStoreCol.加成率) = True
            .ColHidden(menuStoreCol.原采购价) = True
            .ColHidden(menuStoreCol.现采购价) = True
            .ColHidden(menuStoreCol.差价差) = True
            .ColHidden(menuStoreCol.原零售价) = False
            .ColHidden(menuStoreCol.现零售价) = False
        ElseIf cboPriceMethod.Text = "仅调成本价" Then
            .ColHidden(menuStoreCol.原零售价) = True
            .ColHidden(menuStoreCol.现零售价) = True
            .ColHidden(menuStoreCol.调整金额) = True
            .ColHidden(menuStoreCol.加成率) = False
            .ColHidden(menuStoreCol.原采购价) = False
            .ColHidden(menuStoreCol.现采购价) = False
            .ColHidden(menuStoreCol.差价差) = False
        ElseIf cboPriceMethod.Text = "售价成本价一起调价" Then
            .ColHidden(menuStoreCol.原零售价) = False
            .ColHidden(menuStoreCol.现零售价) = False
            .ColHidden(menuStoreCol.调整金额) = False
            .ColHidden(menuStoreCol.加成率) = False
            .ColHidden(menuStoreCol.原采购价) = False
            .ColHidden(menuStoreCol.现采购价) = False
            .ColHidden(menuStoreCol.差价差) = False
        End If
    End With
End Sub

Private Sub vsfStore_Click()
    Dim i As Integer
    With vsfStore
        For i = 1 To vsfPrice.Rows - 1
            If Val(.TextMatrix(.Row, menuStoreCol.材料ID)) = Val(vsfPrice.TextMatrix(i, menuPriceCol.材料ID)) Then
                vsfPrice.Tag = i
            End If
        Next
    End With
End Sub

Private Sub vsfStore_DblClick()
    With vsfStore
        If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mconlngCanColor Then
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
        ElseIf .Col = menuStoreCol.现采购价 Then
            mdblOldPrice = Val(.TextMatrix(.Row, menuStoreCol.现采购价))
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
                If .Row <> .Rows - 1 Then
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
            If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mconlngCanColor Then
                .Editable = flexEDKbdMouse
            Else
                .Editable = flexEDNone
            End If
        End With
    End If
End Sub

Private Sub vsfStore_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strKey As String
    Dim intDigit As Integer

    If KeyAscii = vbKeyReturn Then Exit Sub
    If KeyAscii <> vbKeyBack Then
        With vsfStore
            If Col = menuStoreCol.现采购价 Or Col = menuStoreCol.现零售价 Or Col = menuStoreCol.加成率 Then
                strKey = .EditText
                Select Case Col
                    Case menuStoreCol.现采购价
                        intDigit = Len(Mid(mFMT.FM_成本价, InStr(1, mFMT.FM_成本价, ".") + 1))
                    Case menuStoreCol.现零售价
                        intDigit = Len(Mid(mFMT.FM_成本价, InStr(1, mFMT.FM_零售价, ".") + 1))
                    Case menuStoreCol.加成率
                        intDigit = 5
                End Select
                If KeyAscii = vbKeyDelete Then
                    If InStr(1, .EditText, ".") > 0 Then
                        KeyAscii = 0
                    End If
                ElseIf KeyAscii = Asc(".") Or (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
                    If .EditSelLength = Len(strKey) Then Exit Sub
                    If InStr(strKey, ".") <> 0 And Chr(KeyAscii) = "." Then   '只能存在一个小数点
                        KeyAscii = 0
                        Exit Sub
                    End If
                    If Len(Mid(strKey, InStr(1, strKey, ".") + 1)) >= intDigit And strKey Like "*.*" Then
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
    Dim dbl数量 As Double
    Dim Dbl金额 As Double
    Dim dbl现采购价 As Double

    With vsfStore
        If .EditText = "" Then Exit Sub
        intRow = .Row
        Select Case .Col
            Case menuStoreCol.现零售价
                If Not IsNumeric(.EditText) Then
                    MsgBox "请输入数字！", vbInformation, gstrSysName
                    Cancel = True
                    Exit Sub
                Else
                    .EditText = Format(.EditText, mFMT.FM_零售价)
                End If

'                If mdblOldPrice = .EditText Then Exit Sub

                If .EditText > 9999999 Then
                    MsgBox "零售价过大，请重新输入！", vbInformation, gstrSysName
                    Cancel = True
                    Exit Sub
                End If
                .TextMatrix(intRow, menuStoreCol.调整金额) = Format(Val(.TextMatrix(intRow, menuStoreCol.数量)) * (Val(.EditText) - Val(.TextMatrix(intRow, menuStoreCol.原零售价))), mFMT.FM_金额)
                .TextMatrix(intRow, menuStoreCol.现零售价) = Format(Val(.EditText), mFMT.FM_零售价)
                If Val(.TextMatrix(intRow, menuStoreCol.现采购价)) <> 0 Then
                    .TextMatrix(intRow, menuStoreCol.加成率) = GetFormat((Val(.TextMatrix(intRow, menuStoreCol.现零售价)) / Val(.TextMatrix(intRow, menuStoreCol.现采购价)) - 1) * 100, 3)
                End If
                
                For n = 1 To .Rows - 1
                    If .TextMatrix(intRow, menuStoreCol.材料ID) = .TextMatrix(n, menuStoreCol.材料ID) Then
                        If Val(.TextMatrix(intRow, menuStoreCol.批次)) <> 0 And Val(.TextMatrix(intRow, menuStoreCol.批次)) = Val(.TextMatrix(n, menuStoreCol.批次)) Then
                            .TextMatrix(n, menuStoreCol.现零售价) = .TextMatrix(intRow, menuStoreCol.现零售价)
                            .TextMatrix(n, menuStoreCol.调整金额) = Format(Val(.TextMatrix(n, menuStoreCol.数量)) * (Val(.EditText) - Val(.TextMatrix(n, menuStoreCol.原零售价))), mFMT.FM_金额)
                            If Val(.TextMatrix(n, menuStoreCol.现采购价)) <> 0 Then
                                .TextMatrix(n, menuStoreCol.加成率) = GetFormat((Val(.TextMatrix(n, menuStoreCol.现零售价)) / Val(.TextMatrix(n, menuStoreCol.现采购价)) - 1) * 100, 3)
                            End If
                        End If
                        dbl数量 = dbl数量 + .TextMatrix(n, menuStoreCol.数量)
                        Dbl金额 = Dbl金额 + .TextMatrix(n, menuStoreCol.数量) * Val(.TextMatrix(n, menuStoreCol.现零售价))
                    End If
                Next
                For n = 1 To vsfPrice.Rows - 1
                    If .TextMatrix(intRow, menuStoreCol.材料ID) = vsfPrice.TextMatrix(n, menuPriceCol.材料ID) Then
                        If dbl数量 <> 0 Then
                            vsfPrice.TextMatrix(n, menuPriceCol.现零售价) = Format(Dbl金额 / dbl数量, mFMT.FM_零售价)
                        Else
                            vsfPrice.TextMatrix(n, menuPriceCol.现零售价) = .TextMatrix(intRow, menuStoreCol.现零售价)
                        End If
                    End If
                Next

                If mint调价 > 0 Then
                    For n = 1 To .Rows - 1
                        If .TextMatrix(n, menuStoreCol.材料ID) <> "" Then
                            If Val(.TextMatrix(n, menuStoreCol.材料ID)) = Val(.TextMatrix(intRow, menuStoreCol.材料ID)) Then
                                dbl发票金额 = dbl发票金额 + (Val(.TextMatrix(n, menuStoreCol.现采购价)) - Val(.TextMatrix(n, menuStoreCol.原采购价))) * Val(.TextMatrix(n, menuStoreCol.数量))
                            End If
                        End If
                    Next

                    If chkAutoPay.Value = 1 Then
                        For n = 1 To vsfPay.Rows - 1
                            If vsfPay.TextMatrix(1, 0) <> "" Then
                                If Val(vsfPay.TextMatrix(n, menuPayCol.材料ID)) = Val(vsfStore.TextMatrix(intRow, menuStoreCol.材料ID)) Then
                                    vsfPay.TextMatrix(n, menuPayCol.发票金额) = Format(dbl发票金额, 2)
                                End If
                            End If
                        Next
                    End If
                End If
            Case menuStoreCol.加成率
                If Val(.EditText) < 0 Then Exit Sub

                If Not IsNumeric(.EditText) Then
                    MsgBox "请输入数字！", vbInformation, gstrSysName
                    Cancel = True
                    Exit Sub
                End If
'                If mdblOldPrice = .EditText Then Exit Sub

                .TextMatrix(intRow, menuStoreCol.加成率) = Format(Val(.EditText), "#0.000")
                .TextMatrix(intRow, menuStoreCol.现零售价) = Format(Val(.TextMatrix(intRow, menuStoreCol.现采购价)) * (1 + (Format(Val(.EditText), "#0.000")) / 100), mFMT.FM_零售价)
                .TextMatrix(intRow, menuStoreCol.调整金额) = Format(Val(.TextMatrix(intRow, menuStoreCol.数量)) * (Val(.TextMatrix(intRow, menuStoreCol.现零售价)) - Val(.TextMatrix(intRow, menuStoreCol.原零售价))), mFMT.FM_金额)

                For n = 1 To .Rows - 1
                    If vsfPrice.TextMatrix(Val(vsfPrice.Tag), menuPriceCol.材料ID) = .TextMatrix(n, menuStoreCol.材料ID) Then
                        If Val(.TextMatrix(intRow, menuStoreCol.变价)) = 0 Or mbln时价卫材按批次调价 = False Then
                            .TextMatrix(n, menuStoreCol.加成率) = Format(.TextMatrix(intRow, menuStoreCol.加成率), "#0.000")
                            .TextMatrix(n, menuStoreCol.现零售价) = Format(Val(.TextMatrix(n, menuStoreCol.现采购价)) * (1 + (Format(Val(.EditText), "#0.000")) / 100), mFMT.FM_零售价)
                            .TextMatrix(n, menuStoreCol.调整金额) = Format(Val(.TextMatrix(n, menuStoreCol.数量)) * (Val(.TextMatrix(n, menuStoreCol.现零售价)) - Val(.TextMatrix(n, menuStoreCol.原零售价))), mFMT.FM_金额)
                        End If
                        dbl数量 = dbl数量 + .TextMatrix(n, menuStoreCol.数量)
                        Dbl金额 = Dbl金额 + .TextMatrix(n, menuStoreCol.数量) * Val(.TextMatrix(n, menuStoreCol.现零售价))
                    End If
                Next
                If dbl数量 <> 0 Then
                    vsfPrice.TextMatrix(Val(vsfPrice.Tag), menuPriceCol.现零售价) = Format(Dbl金额 / dbl数量, mFMT.FM_零售价)
                Else
                    vsfPrice.TextMatrix(Val(vsfPrice.Tag), menuPriceCol.现零售价) = .TextMatrix(intRow, menuStoreCol.现零售价)
                End If
            Case menuStoreCol.现采购价
                If Not IsNumeric(.EditText) Then
                    MsgBox "请输入数字！", vbInformation, gstrSysName
                    Cancel = True
                    Exit Sub
                End If
                If .EditText > 9999999 Then
                    MsgBox "采购价过大，请重新输入！", vbInformation, gstrSysName
                    Cancel = True
                    Exit Sub
                End If
                If Val(.EditText) < 0 Then
                    MsgBox "成本价不能为负数！", vbExclamation, gstrSysName
                    Cancel = True
                End If

                .EditText = Format(Val(.EditText), mFMT.FM_成本价)
                .TextMatrix(intRow, menuStoreCol.现采购价) = Format(Val(.EditText), mFMT.FM_成本价)
                .TextMatrix(intRow, menuStoreCol.差价差) = Format((Val(.EditText) - .TextMatrix(intRow, menuStoreCol.原采购价)) * Val(.TextMatrix(intRow, menuStoreCol.数量)), mFMT.FM_金额)
                If Val(.TextMatrix(intRow, menuStoreCol.变价)) = 1 And mbln时价卫材按批次调价 = True And mint调价 <> 1 Then
                    .TextMatrix(intRow, menuStoreCol.现零售价) = Format(Val(.TextMatrix(intRow, menuStoreCol.现采购价)) * (1 + (Val(.TextMatrix(intRow, menuStoreCol.加成率)) / 100)), mFMT.FM_零售价)
                    .TextMatrix(intRow, menuStoreCol.调整金额) = Format(Val(.TextMatrix(intRow, menuStoreCol.数量)) * (Val(.TextMatrix(intRow, menuStoreCol.现零售价)) - Val(.TextMatrix(intRow, menuStoreCol.原零售价))), mFMT.FM_金额)
                End If
                
                dbl发票金额 = (Val(.EditText) - .TextMatrix(intRow, menuStoreCol.原采购价)) * Val(.TextMatrix(intRow, menuStoreCol.数量))

                For n = 1 To .Rows - 1
                    If .TextMatrix(n, menuStoreCol.材料ID) <> "" Then
                        If Val(.TextMatrix(n, menuStoreCol.材料ID)) = Val(.TextMatrix(intRow, menuStoreCol.材料ID)) And n <> intRow Then
                            If chkCostBatch.Value = 0 Or (Val(.TextMatrix(intRow, menuStoreCol.批次)) <> 0 And Val(.TextMatrix(intRow, menuStoreCol.批次)) = Val(.TextMatrix(n, menuStoreCol.批次))) Then
                                dbl现采购价 = Format(Val(.EditText), mFMT.FM_成本价)
                                .TextMatrix(n, menuStoreCol.现采购价) = Format(dbl现采购价, mFMT.FM_成本价)
                                .TextMatrix(n, menuStoreCol.差价差) = Format((dbl现采购价 - .TextMatrix(n, menuStoreCol.原采购价)) * Val(.TextMatrix(n, menuStoreCol.数量)), mFMT.FM_金额)
                                If Val(.TextMatrix(intRow, menuStoreCol.变价)) = 1 And mbln时价卫材按批次调价 = True And mint调价 <> 1 Then
                                    .TextMatrix(n, menuStoreCol.现零售价) = Format(dbl现采购价 * (1 + (Val(.TextMatrix(n, menuStoreCol.加成率)) / 100)), mFMT.FM_零售价)
                                    .TextMatrix(n, menuStoreCol.调整金额) = Format(Val(.TextMatrix(n, menuStoreCol.数量)) * (Val(.TextMatrix(n, menuStoreCol.现零售价)) - Val(.TextMatrix(n, menuStoreCol.原零售价))), mFMT.FM_金额)
                                End If
                            Else
                                dbl现采购价 = Val(.TextMatrix(n, menuStoreCol.现采购价))
                            End If
                            dbl发票金额 = dbl发票金额 + (dbl现采购价 - .TextMatrix(n, menuStoreCol.原采购价)) * Val(.TextMatrix(n, menuStoreCol.数量))
                        End If
                    End If
                Next

                If chkAutoPay.Value = 1 Then
                    For n = 1 To vsfPay.Rows - 1
                        If vsfPay.TextMatrix(1, 0) <> "" Then
                            If Val(vsfPay.TextMatrix(n, menuPayCol.材料ID)) = Val(vsfStore.TextMatrix(intRow, menuStoreCol.材料ID)) Then
                                vsfPay.TextMatrix(n, menuPayCol.发票金额) = Format(dbl发票金额, mFMT.FM_金额)
                            End If
                        End If
                    Next
                End If

                If chkCostBatch.Value = 0 Then
                    For n = 1 To vsfPrice.Rows - 1
                        If Val(.TextMatrix(intRow, menuStoreCol.材料ID)) = Val(vsfPrice.TextMatrix(n, menuPriceCol.材料ID)) Then
                            vsfPrice.TextMatrix(n, menuPriceCol.现成本价) = Format(.TextMatrix(intRow, menuStoreCol.现采购价), mFMT.FM_成本价)
                            Exit For
                        End If
                    Next
                Else
                    CaluateAverCost Val(.TextMatrix(intRow, menuStoreCol.材料ID))
                End If
                Call CaculateAverPirce(Val(.TextMatrix(intRow, menuStoreCol.材料ID)))   '售价变动，计算平均售价
        End Select
    End With
End Sub

Private Sub CaculateAverPirce(ByVal lng材料ID As Long)
    '自动计算平均售价
    Dim i As Integer
    Dim dblSumPrice As Double
    Dim dblSumNumber As Double
    
    With vsfStore
        For i = 1 To .Rows - 1
            If .TextMatrix(i, menuStoreCol.材料ID) <> "" Then
                If Val(.TextMatrix(i, menuStoreCol.材料ID)) = lng材料ID Then
                    dblSumPrice = dblSumPrice + Val(.TextMatrix(i, menuStoreCol.现零售价)) * Val(.TextMatrix(i, menuStoreCol.数量))
                    dblSumNumber = dblSumNumber + Val(.TextMatrix(i, menuStoreCol.数量))
                End If
            End If
        Next
    End With

    With vsfPrice
        If dblSumNumber > 0 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, menuPriceCol.材料ID) <> "" Then
                    If Val(.TextMatrix(i, menuPriceCol.材料ID)) = lng材料ID Then
                        .TextMatrix(i, menuPriceCol.现零售价) = Format(dblSumPrice / dblSumNumber, mFMT.FM_零售价)
                        Exit For
                    End If
                End If
            Next
        End If
    End With
End Sub

Private Sub CaculateAverOldPirce(ByVal lng材料ID As Long)
    '自动原始计算平均售价
    Dim i As Integer
    Dim dblSumPrice As Double
    Dim dblSumNumber As Double
    
    With vsfStore
        For i = 1 To .Rows - 1
            If .TextMatrix(i, menuStoreCol.材料ID) <> "" Then
                If Val(.TextMatrix(i, menuStoreCol.材料ID)) = lng材料ID Then
                    dblSumPrice = dblSumPrice + Val(.TextMatrix(i, menuStoreCol.原零售价)) * Val(.TextMatrix(i, menuStoreCol.数量))
                    dblSumNumber = dblSumNumber + Val(.TextMatrix(i, menuStoreCol.数量))
                End If
            End If
        Next
    End With

    With vsfPrice
        If dblSumNumber > 0 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, menuPriceCol.材料ID) <> "" Then
                    If Val(.TextMatrix(i, menuPriceCol.材料ID)) = lng材料ID Then
                        .TextMatrix(i, menuPriceCol.原零售价) = Format(dblSumPrice / dblSumNumber, mFMT.FM_零售价)
                        Exit For
                    End If
                End If
            Next
        End If
    End With
End Sub

Private Sub CatalogModifyPrice()
'卫材目录直接进入调价
    Dim rsprice As New ADODB.Recordset
    gstrSQL = "Select Distinct i.Id,i.编码,b.名称 As 商品名,i.名称 As 通用名,i.规格,i.产地,i.计算单位,p.换算系数,p.包装单位," & vbNewLine & _
        "                Decode(i.是否变价, 0, '定价', 1, '时价') As 时价," & vbNewLine & _
        "                To_Char(p.成本价, '9999999999990.9999999') As 成本价," & vbNewLine & _
        "                To_Char(p.指导批发价, '9999999999990.9999999') 指导批发价," & vbNewLine & _
        "                To_Char(p.指导零售价, '9999999999990.9999999') 指导零售价," & vbNewLine & _
        "                p.跟踪在用" & vbNewLine & _
        "From 收费项目目录 i, 材料特性 p, 收费项目别名 b" & vbNewLine & _
        "Where i.Id = p.材料id And i.Id = b.收费细目id(+) And b.性质(+) = 3 And i.类别 = '4' and i.id=[1] And" & vbNewLine & _
        "      (i.撤档时间 Is Null Or i.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd'))"
    
    Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, "卫材直接进入调价", mlng规格ID)
    
    Call GetDrugPirce(rsprice, 1)
    chkAppAllColumn.Enabled = False
End Sub
