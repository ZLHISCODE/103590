VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frm批量冲销 
   Caption         =   "批量冲销"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11790
   Icon            =   "frm批量冲销.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   11790
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fra药品基本信息 
      Height          =   615
      Left            =   3720
      TabIndex        =   33
      Top             =   600
      Width           =   7215
      Begin VB.Label lbl药品基本信息 
         Caption         =   "药品信息"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   6975
      End
   End
   Begin VB.Frame fraColSel 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7920
      TabIndex        =   29
      Top             =   3960
      Width           =   195
      Begin VB.Image imgColSel 
         Height          =   195
         Left            =   0
         Picture         =   "frm批量冲销.frx":6852
         Top             =   0
         Width           =   195
      End
   End
   Begin VB.PictureBox pic其他信息 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   120
      ScaleHeight     =   3255
      ScaleWidth      =   3375
      TabIndex        =   15
      Top             =   3840
      Width           =   3375
      Begin VB.TextBox txt填制人 
         Height          =   300
         Left            =   960
         TabIndex        =   8
         Top             =   1320
         Width           =   2055
      End
      Begin VB.ComboBox cbo填制时间 
         Height          =   300
         Left            =   960
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   1720
         Width           =   2055
      End
      Begin VB.TextBox txt审核人 
         Height          =   300
         Left            =   960
         TabIndex        =   7
         Top             =   920
         Width           =   2055
      End
      Begin VB.TextBox txt开始NO 
         Height          =   300
         Left            =   960
         TabIndex        =   5
         Top             =   120
         Width           =   2055
      End
      Begin VB.TextBox txt结束NO 
         Height          =   300
         Left            =   960
         TabIndex        =   6
         Top             =   520
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker DTP结束填制 
         Height          =   300
         Left            =   960
         TabIndex        =   23
         Top             =   2520
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   159449091
         CurrentDate     =   40750
      End
      Begin MSComCtl2.DTPicker DTP开始填制 
         Height          =   300
         Left            =   960
         TabIndex        =   24
         Top             =   2115
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   159449091
         CurrentDate     =   40750
      End
      Begin VB.Label lbl填制人 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "填制人"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   1350
         Width           =   615
      End
      Begin VB.Label lbl结束填制 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "结束时间"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   2550
         Width           =   735
      End
      Begin VB.Label lbl开始填制 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "开始时间"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2145
         Width           =   735
      End
      Begin VB.Label lbl填制时间 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "填制时间"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1755
         Width           =   735
      End
      Begin VB.Label lbl审核人 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "审核人"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   945
         Width           =   615
      End
      Begin VB.Label lbl开始NO 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "开始NO"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   150
         Width           =   615
      End
      Begin VB.Label lbl结束NO 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "结束NO"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   555
         Width           =   615
      End
   End
   Begin VB.PictureBox pic基本信息 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   120
      ScaleHeight     =   2295
      ScaleWidth      =   3375
      TabIndex        =   11
      Top             =   1080
      Width           =   3375
      Begin VB.ComboBox cbo库房 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   97
         Width           =   2055
      End
      Begin VB.ComboBox cbo时间范围 
         Height          =   300
         Left            =   960
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   933
         Width           =   2055
      End
      Begin VB.CommandButton cmd药品 
         Caption         =   "…"
         Height          =   300
         Left            =   2720
         TabIndex        =   3
         Top             =   517
         Width           =   255
      End
      Begin VB.TextBox txt药品 
         Height          =   300
         Left            =   960
         TabIndex        =   2
         Top             =   517
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTP结束时间 
         Height          =   300
         Left            =   960
         TabIndex        =   16
         Top             =   1770
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   159449091
         CurrentDate     =   40750
      End
      Begin MSComCtl2.DTPicker DTP开始时间 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "gg yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         Height          =   300
         Left            =   960
         TabIndex        =   31
         Top             =   1357
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   159449091
         CurrentDate     =   40750
      End
      Begin VB.Label lbl结束时间 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "结束时间"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lbl开始时间 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "开始时间"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1380
         Width           =   735
      End
      Begin VB.Label lbl时间范围 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "审核时间"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lbl库房id 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "库房名称"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   735
      End
      Begin VB.Label lbl药品 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "药品名称"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   540
         Width           =   735
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   2295
      Left            =   3840
      TabIndex        =   32
      Top             =   1680
      Width           =   7215
      _cx             =   12726
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16777215
      ForeColorSel    =   -2147483641
      BackColorBkg    =   -2147483644
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483641
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
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
      FrozenCols      =   1
      AllowUserFreezing=   0
      BackColorFrozen =   -2147483633
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Frame fraEW 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5415
      Left            =   3840
      MousePointer    =   9  'Size W E
      TabIndex        =   10
      Top             =   240
      Width           =   45
   End
   Begin XtremeSuiteControls.TaskPanel tkpMain 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3375
      _Version        =   589884
      _ExtentX        =   5953
      _ExtentY        =   1508
      _StockProps     =   64
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   13
      Top             =   7080
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm批量冲销.frx":6DA0
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14446
            Text            =   "药品的可用库存"
            TextSave        =   "药品的可用库存"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frm批量冲销.frx":7634
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frm批量冲销.frx":7B36
            Key             =   "WB"
            Object.ToolTipText     =   "五笔(F7)"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin VSFlex8Ctl.VSFlexGrid vsfColSel 
      Height          =   1335
      Left            =   4440
      TabIndex        =   30
      Top             =   3840
      Visible         =   0   'False
      Width           =   1470
      _cx             =   2593
      _cy             =   2355
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
      FormatString    =   $"frm批量冲销.frx":8038
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   2880
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm批量冲销.frx":8086
            Key             =   "当前"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.ImageManager imgicon 
      Left            =   1680
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frm批量冲销.frx":E8E8
   End
   Begin XtremeCommandBars.CommandBars combars 
      Left            =   480
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane docpane 
      Bindings        =   "frm批量冲销.frx":1435E
      Left            =   480
      Top             =   720
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frm批量冲销"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const INTPANETYPE As Integer = 0
Private Const INTPANEDETAIL As Integer = 1
Private mfrmMain As Form
Private MStrCaption As String
Private mintUnit As Integer
Private mintCostDigit As Integer  '成本价的小数点位数
Private mintPriceDigit As Integer
Private mintNumberDigit As Integer
Private mintMoneyDigit As Integer
Private mint库存检查 As Integer
Private mint库存检查入库库房 As Integer
Private mbln下可用数量 As Boolean
Private mlng模块号 As Long
Private mstr列名列宽 As String  '保存列的位置和列宽
Private mrsReturn As Recordset  '保存药品选择器选择的药品
Private mInt单据号 As Integer   '各个业务的序号

Private Const MCST_INVALIDCHAR As String = "'" '禁止输入的字符

'定义单位级别
Private Const MCONINTPRICEUNIT As Integer = 1   '售价单位
Private Const MCONINTOUTUNIT As Integer = 2     '门诊单位
Private Const MCONINTINUNIT As Integer = 3      '住院单位
Private Const MCONINTSTOREUNIT As Integer = 4   '药库单位

'定义颜色常量
Private Const CSTCOLOR_FIXED = &H808080        '灰色，列选择中不能编辑的字体颜色
Private Const CSTCOLOR_MODIFY = &HE0E0E0       '深灰色，表格行修改之后的背景色
Private Const CSTCOLOR_FONT = vbRed            '红色，冲销单元格数量不为0的字体颜色
Private Const CSTCOLOR_NOFONT = vbBlack        '黑色，冲销单元格数量为0的字体颜色
Private Const CSTCOLOR_NOMODIFY = vbWhite      '白色，表格行修改之前的背景色
Private Const CSTCOLOR_ENTERCELL = &HFF0000    '蓝色，进入冲销数量单元格的边框颜色
Private Const CSTCOLOR_LOSTFORCE = &H80000005  '表格失去焦点之后，表格选中的颜色

'定义列常量
Private mintcol选择 As Integer
Private mintcol药品id As Integer
Private mintcol行号 As Integer
Private mIntColNO As Integer
Private mintcol药品名称与编码 As Integer
Private mintcol商品名 As Integer
Private mintcol药品来源 As Integer
Private mintcol基本药物 As Integer
Private mintcol药价级别 As Integer
Private mintcol规格 As Integer
Private mintcol单位 As Integer
Private mintcol数量 As Integer
Private mintcol冲销数量 As Integer
Private mintcol产地 As Integer
Private mintcol批号 As Integer
Private mintcol生产日期 As Integer
Private mintcol有效期至 As Integer
Private mintcol部门名称 As Integer
Private mintcol批准文号 As Integer
Private mintcol其他外观 As Integer   '其他入库的字段

Private mintcol采购限价 As Integer
Private mintcol采购价 As Integer
Private mintcol扣率 As Integer
Private mintcol成本价 As Integer    '外购入库则为结算价
Private mintcol成本金额 As Integer  '外购入库则为结算金额
Private mintcol加成率 As Integer
Private mintcol售价 As Integer
Private mintcol售价金额 As Integer
Private mintcol差价 As Integer
Private mintcol填制人 As Integer
Private mintcol填制日期 As Integer
Private mintcol审核人 As Integer
Private mintcol审核日期 As Integer

'外购入库需要的字段
Private mintcol零售价 As Integer
Private mintcol零售单位 As Integer
Private mintcol零售金额 As Integer
Private mintcol零售差价 As Integer
Private mintcol外购批准文号 As Integer
Private mintcol随货单号 As Integer
Private mintcol发票号 As Integer
Private mintcol发票代码 As Integer
Private mintcol发票信息 As Integer
Private mintcol发票金额 As Integer

'需要隐藏的列
Private mintcol真实数量 As Integer
Private mintcol序号 As Integer
Private mintcol比例系数 As Integer
Private mintcol药名 As Integer
Private mintcol批次 As Integer
Private mintcol记录状态 As Integer
Private mintcol分批核算 As Integer
Private mintcol可用数量 As Integer
Private mintcol最大效期 As Integer
Private mintcol实际差价 As Integer
Private mintcol实际金额 As Integer
Private mintcol上次供应商ID As Integer
Private mintcol摘要 As Integer
Private mintcol对方部门 As Integer
Private mintcol是否变价 As Integer
Private Const MINTCOL总列数 As Integer = 59

'工具栏按钮的定义
Private Const MINTBTNFILTER As Integer = 1          '过滤按钮
Private Const MINTBTNALLWRITEOFF As Integer = 2     '全冲按钮
Private Const MINTBTNALLELIMINATE As Integer = 3    '全清按钮
Private Const MINTBTNDEL As Integer = 4             '删除按钮
Private Const MINTBTNWRITEOFF As Integer = 5        '冲销按钮
Private Const MINTBTNHELP  As Integer = 6           '帮助按钮
Private Const MINTBTNEXIT  As Integer = 7           '退出按钮
Private Const MINTBTNSIMPLE  As Integer = 8         '简洁按钮
Private Const MINTBTNCONPLETE   As Integer = 9      '完整按钮
Private mrecSort As Recordset

Private Sub cbo库房_Click()
    mint库存检查 = MediWork_GetCheckStockRule(Me.cbo库房.ItemData(cbo库房.ListIndex))
    mint库存检查入库库房 = MediWork_GetCheckStockRule(cbo库房.ItemData(cbo库房.ListIndex))
End Sub

Private Sub cbo库房_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub
Private Sub cbo时间范围_Click()
    If Me.cbo时间范围.ListIndex = 0 Then
        Me.DTP开始时间.Value = Date
        Me.DTP结束时间.Value = Date
        Me.DTP开始时间.Enabled = False
        Me.DTP结束时间.Enabled = False
    ElseIf Me.cbo时间范围.ListIndex = 1 Then
        Me.DTP开始时间.Value = Date - 1
        Me.DTP结束时间.Value = Date
        Me.DTP开始时间.Enabled = False
        Me.DTP结束时间.Enabled = False
    ElseIf Me.cbo时间范围.ListIndex = 2 Then
        Me.DTP开始时间.Value = Date - 2
        Me.DTP结束时间.Value = Date
        Me.DTP开始时间.Enabled = False
        Me.DTP结束时间.Enabled = False
    Else
        Me.DTP开始时间.Value = Date - 30
        Me.DTP结束时间.Value = Date
        Me.DTP开始时间.Enabled = True
        Me.DTP结束时间.Enabled = True
    End If
End Sub
Private Sub cbo时间范围_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub
Private Sub cbo填制时间_Click()
    If Me.cbo填制时间.ListIndex = 0 Then
        Me.DTP开始填制.Value = Date
        Me.DTP结束填制.Value = Date
        Me.DTP开始填制.Enabled = False
        Me.DTP结束填制.Enabled = False
    ElseIf Me.cbo填制时间.ListIndex = 1 Then
        Me.DTP开始填制.Value = Date - 1
        Me.DTP结束填制.Value = Date
        Me.DTP开始填制.Enabled = False
        Me.DTP结束填制.Enabled = False
    ElseIf Me.cbo填制时间.ListIndex = 2 Then
        Me.DTP开始填制.Value = Date - 2
        Me.DTP结束填制.Value = Date
        Me.DTP开始填制.Enabled = False
        Me.DTP结束填制.Enabled = False
    Else
        Me.DTP开始填制.Value = Date - 30
        Me.DTP结束填制.Value = Date
        Me.DTP开始填制.Enabled = True
        Me.DTP结束填制.Enabled = True
    End If
End Sub
Private Sub cbo填制时间_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub
'药品使用药品的通用选择器
Private Sub cmd药品_Click()
    Dim vRect As RECT
    Dim strsql As String
    
    On Error GoTo errRow
    
    If grsMaster.State = adStateClosed Then
        Call SetSelectorRS(6, "药品外购入库管理", cbo库房.ItemData(cbo库房.ListIndex), cbo库房.ItemData(cbo库房.ListIndex))
    End If
    
'    Set mrsReturn = Frm药品选择器.ShowME(Me, 6, cbo库房.ItemData(cbo库房.ListIndex), cbo库房.ItemData(cbo库房.ListIndex), , True, True, False, False, True, 0)
    Set mrsReturn = frmSelector.showMe(Me, 0, 6, , , , cbo库房.ItemData(cbo库房.ListIndex), cbo库房.ItemData(cbo库房.ListIndex), , 0, True, True, True, , False)

    If Not mrsReturn.EOF Then
        Me.txt药品.Text = mrsReturn!通用名
        Me.txt药品.Tag = mrsReturn!药品ID
    End If
    Exit Sub
errRow:
    If ErrCenter <> 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub cmd药品_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub
Private Sub combars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
Dim i As Integer
Dim strKey As String
Dim count As Integer
Dim dblSum As Integer
Dim dblOldSum As Integer

Select Case Control.Id
    Case MINTBTNFILTER            '执行过滤操作
        Call Filter
    Case MINTBTNALLWRITEOFF       '执行全冲操作
        Call AllWriteOff
    Case MINTBTNALLELIMINATE      '执行全清操作
        Call AllEliminate
    Case MINTBTNWRITEOFF          '执行冲销操作
        Call WriteOff
    Case MINTBTNDEL               '执行删除操作
        Call DelRow
    Case MINTBTNHELP              '执行帮助操作
        Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
    Case MINTBTNEXIT              '执行退出操作
        Unload Me
    Case MINTBTNSIMPLE            '执行设置为简洁
        Call SetSimple(Control)
    Case MINTBTNCONPLETE          '执行设置为完整
        Call SetConplete(Control)
End Select
End Sub

Private Sub InitData()
'----------------------------------------------
'执行数据初始化操作，主要是根据条件得出sql语句，然后将查询结果保存到vsflexfrid表格中
'----------------------------------------------
    Dim strsql As String
    Dim rsData As Recordset
    Dim strOrder As String
    Dim strCompare As String
    Dim strSqlOrder As String
    Dim strUnitQuantity As String
    Dim i As Long
    Dim j As Long
    Dim int库房id  As Long
    Dim int包装系数 As String
    Dim str药品信息 As String
    
    On Error GoTo errRow
    strOrder = zldatabase.GetPara("排序", glngSys, mlng模块号)
    strCompare = Mid(strOrder, 1, 1)
    strSqlOrder = "序号"
    
    '排序的依据
    If strCompare = "0" Then
        strSqlOrder = "序号"
    ElseIf strCompare = "1" Then
        strSqlOrder = "药品编码"
    ElseIf strCompare = "2" Then
        If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
            strSqlOrder = "通用名"
        Else
            strSqlOrder = "Nvl(商品名, 通用名)"
        End If
    End If
    
    '排序的方式
    strSqlOrder = strSqlOrder & IIf(Right(strOrder, 1) = "0", " ASC", " DESC")
    int库房id = cbo库房.ItemData(cbo库房.ListIndex)
    
    '取各种价格的精度
    Call GetDrugDigit(int库房id, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    
    Select Case mlng模块号
        Case 模块号.药品移库
            Select Case mintUnit
                Case MCONINTPRICEUNIT
                    strUnitQuantity = "C.计算单位 AS 单位, A.填写数量,a.实际数量,a.成本价,a.零售价,'1' as 比例系数,"
                Case MCONINTOUTUNIT
                    strUnitQuantity = "B.门诊单位 AS 单位,(A.填写数量 / B.门诊包装) AS 填写数量,(A.实际数量 / B.门诊包装) AS 实际数量,a.成本价*B.门诊包装 as 成本价,a.零售价*B.门诊包装 as 零售价,B.门诊包装 as 比例系数,"
                Case MCONINTINUNIT
                    strUnitQuantity = "B.住院单位 AS 单位,(A.填写数量 / B.住院包装) AS 填写数量,(A.实际数量 / B.住院包装) AS 实际数量,a.成本价*B.住院包装 as 成本价,a.零售价*B.住院包装 as 零售价,B.住院包装 as 比例系数,"
                Case MCONINTSTOREUNIT
                    strUnitQuantity = "B.药库单位 AS 单位,(A.填写数量 / B.药库包装) AS 填写数量,(A.实际数量 / B.药库包装) AS 实际数量,a.成本价*B.药库包装 as 成本价,a.零售价*B.药库包装 as 零售价,B.药库包装 as 比例系数,"
            End Select
            
            strsql = "SELECT W.*,Z.可用数量/W.比例系数 AS  可用数量,Z.实际金额,Z.实际差价 " & _
                        " FROM " & _
                        "     (SELECT DISTINCT a.NO,d.记录状态,A.药品ID,A.序号,'[' || C.编码 || ']' As 药品编码, C.名称 As 通用名, E.名称 As 商品名," & _
                        "     B.药品来源,B.基本药物,C.规格,C.产地 AS 原产地,A.产地, A.批号,A.批次,B.加成率,B.药库分批 AS 分批核算," & _
                        "     B.最大效期,A.效期," & strUnitQuantity & _
                        "     A.成本金额,0 零售金额, 0 差价,D.摘要,A.库房ID,A.对方部门ID,C.是否变价,B.药房分批 AS 药房分批核算,A.上次供应商ID,A.批准文号,A.填写数量 真实数量 " & _
                        " ,D.填制人,D.填制日期,D.审核人,D.审核日期,Y.名称 对方库房" & _
                        "     FROM " & _
                        "         (SELECT MIN(ID) AS ID, SUM(实际数量) AS 填写数量,0 实际数量,SUM(成本金额) AS 成本金额,NO,药品ID,序号,产地," & _
                        "   批号,效期,NVL(批次,0) 批次,扣率,成本价,零售价,库房ID,对方部门ID,入出类别ID,NVL(供药单位ID,0) 上次供应商ID,批准文号" & _
                        "          FROM 药品收发记录 X " & _
                        "          WHERE 药品id=[1] AND 单据=6 AND 入出系数=1 " & _
                        "          GROUP BY NO,药品ID,序号,产地,批号,效期,NVL(批次,0),扣率,成本价,零售价,库房ID,对方部门ID,入出类别ID,NVL(供药单位ID,0),批准文号" & _
                        "          HAVING SUM(实际数量)<>0) A," & _
                        "     药品规格 B,收费项目目录 C,收费项目别名 E,部门表 Y, " & _
                        " (Select NO,序号,摘要,记录状态,填制人,填制日期,审核人,审核日期 From 药品收发记录 " & _
                        "  Where 单据 = 6 And 药品ID = [1] and 库房ID=[2] And 入出系数 = 1 And (记录状态 = 1 Or Mod(记录状态, 3) = 0)) D " & _
                        "     WHERE A.药品ID = B.药品ID AND B.药品ID=E.收费细目ID(+) AND A.对方部门id=Y.id AND E.性质(+)=3 AND B.药品ID=C.ID And A.序号 = D.序号 and d.no=a.no) W," & _
                        "     (SELECT  药品ID,NVL(批次,0) 批次,可用数量,实际金额,实际差价 " & _
                        "     FROM 药品库存 WHERE 库房ID=[2] AND 性质=1) Z " & _
                        " WHERE W.药品ID=Z.药品ID(+) AND NVL(W.批次,0)=Nvl(Z.批次(+),0)" & _
                        " And w.No Not in (Select NO From 药品收发记录 Where 单据 = 6 And 入出系数 = 1 And" & _
                        " 药品id = [1] And 库房id = [2] Having Sum(实际数量) = 0 Group By NO, 序号, 药品id)"
        Case 模块号.药品领用
            Select Case mintUnit
                Case MCONINTPRICEUNIT
                    strUnitQuantity = "F.计算单位 AS 单位, A.填写数量,a.实际数量,a.成本价,a.零售价,'1' as 比例系数,"
                Case MCONINTOUTUNIT
                    strUnitQuantity = "B.门诊单位 AS 单位,(A.填写数量 / B.门诊包装) AS 填写数量,(A.实际数量 / B.门诊包装) AS 实际数量,a.成本价*B.门诊包装 as 成本价,a.零售价*B.门诊包装 as 零售价,B.门诊包装 as 比例系数,"
                Case MCONINTINUNIT
                    strUnitQuantity = "B.住院单位 AS 单位,(A.填写数量 / B.住院包装) AS 填写数量,(A.实际数量 / B.住院包装) AS 实际数量,a.成本价*B.住院包装 as 成本价,a.零售价*B.住院包装 as 零售价,B.住院包装 as 比例系数,"
                Case MCONINTSTOREUNIT
                    strUnitQuantity = "B.药库单位 AS 单位,(A.填写数量 / B.药库包装) AS 填写数量,(A.实际数量 / B.药库包装) AS 实际数量,a.成本价*B.药库包装 as 成本价,a.零售价*B.药库包装 as 零售价,B.药库包装 as 比例系数,"
            End Select
        
            strsql = "Select w.*, z.可用数量 / w.比例系数 可用数量, z.实际金额, z.实际差价" & _
                " From" & _
                " (Select Distinct a.No, x.记录状态, 填制人, 填制日期, 审核人, 审核日期, a.药品id," & _
                "a.序号, '[' || f.编码 || ']' As 药品编码, f.名称 As 通用名,e.名称 As 商品名," & _
                "Nvl(e.名称, f.名称) 名称, b.药品来源, b.基本药物, f.规格, f.产地 As 原产地, a.产地," & _
                "a.批号, Nvl(a.批次, 0) 批次,b.加成率, a.效期," & strUnitQuantity & _
                "a.成本金额, 0 零售金额, 0 差价, a.摘要, a.库房id,a.对方部门id, c.名称 As 对方库房," & _
                "f.是否变价, b.药房分批 As 药房分批核算, a.领用人, a.批准文号, a.发药方式, a.填写数量 原始数量" & _
                " From " & _
                " (Select Min(ID) As ID, Sum(实际数量) As 填写数量, 0 实际数量, Sum(成本金额) As 成本金额," & _
                "NO, 药品id, 序号, 产地, 批号, 效期, Nvl(批次, 0) 批次,扣率, 成本价, 零售价, 摘要, 库房id," & _
                "对方部门id, 入出类别id, Nvl(x.领用人, '') As 领用人, x.批准文号, x.发药方式 From 药品收发记录 X" & _
                " Where 单据 = 7 And 药品id = [1] Group By NO, 药品id, 序号, 产地, 批号, 效期, Nvl(批次, 0), 扣率," & _
                "成本价, 零售价, 摘要, 库房id, 对方部门id, 入出类别id, 领用人, 批准文号, 发药方式" & _
                " Having Sum(实际数量) <> 0) A, 药品规格 B, 收费项目别名 E, 收费项目目录 F, 部门表 C," & _
                "(Select NO, 序号, 摘要, 记录状态, 填制人, 填制日期, 审核人, 审核日期 From 药品收发记录" & _
                " Where 单据 = 7 And 药品id = [1] And (记录状态 = 1 Or Mod(记录状态, 3) = 0)) X" & _
                " Where a.No = x.No And a.序号 = x.序号 And a.药品id = b.药品id And b.药品id = f.Id" & _
                " And a.对方部门id = c.Id And b.药品id = e.收费细目id(+) And e.性质(+) = 3) W, 药品库存 Z" & _
                " Where w.药品id = z.药品id(+) And Nvl(w.批次, 0) = Nvl(z.批次(+), 0) And z.库房id(+) = [2] And z.性质(+) = 1"

        Case 模块号.其他入库
             Select Case mintUnit
                Case MCONINTPRICEUNIT
                    strUnitQuantity = "F.计算单位 AS 售价单位,F.计算单位 AS 单位, A.填写数量 AS 填写数量,b.指导批发价 as 指导批发价, a.成本价,A.零售价,1 as 比例系数,"
                Case MCONINTOUTUNIT
                    strUnitQuantity = "F.计算单位 AS 售价单位,B.门诊单位 AS 单位,(A.填写数量 / B.门诊包装) AS 填写数量,b.指导批发价*B.门诊包装 as 指导批发价 , a.成本价*B.门诊包装 as 成本价,A.零售价*B.门诊包装 as 零售价 ,B.门诊包装 as 比例系数,"
                Case MCONINTINUNIT
                    strUnitQuantity = "F.计算单位 AS 售价单位,B.住院单位 AS 单位,(A.填写数量 / B.住院包装) AS 填写数量,b.指导批发价*B.住院包装 as 指导批发价 , a.成本价*B.住院包装 as 成本价,A.零售价*B.住院包装 as 零售价 ,  B.住院包装 as 比例系数,"
                Case MCONINTSTOREUNIT
                    strUnitQuantity = "F.计算单位 AS 售价单位,B.药库单位 AS 单位,(A.填写数量 / B.药库包装) AS 填写数量,b.指导批发价*B.药库包装 as 指导批发价 , a.成本价*B.药库包装 as 成本价,A.零售价*B.药库包装 as 零售价 ,B.药库包装 as 比例系数,"
            End Select
            
            strsql = " Select w.*,z.可用数量 / w.比例系数 可用数量" & _
                " From (Select Distinct a.no,a.药品id, a.序号, x.记录状态,  x.填制人, x. 填制日期,  x.审核人,  x.审核日期,'[' || f.编码 || ']' As 药品编码," & _
                "f.名称 As 通用名, e.名称 As 商品名, b.药品来源, b.基本药物,b.药价级别,f.规格,f.产地 As 原产地, a.产地, a.批号, b.最大效期, a.效期," & _
                strUnitQuantity & "a.成本金额, 0 零售金额, 0 差价, b.加成率 ," & _
                "f.是否变价, b.药房分批 As 药房分批核算, a.摘要, a.库房id, g.名称 As 部门, a.入出类别id," & _
                "a.生产日期, a.批准文号, a.外观,a.填写数量 真实数量, a.金额差,a.批次" & _
                " From (Select Min(ID) As ID, Sum(实际数量) As 填写数量, Sum(成本金额) As 成本金额, Sum(To_Number(Nvl(用法, 0))) As 金额差, no,药品id," & _
                "序号 , 产地, 批号, 效期, 扣率, 成本价, 零售价, 摘要, 库房id, 入出类别id, x.生产日期, x.批准文号, x.外观,nvl(批次,0) 批次" & _
                " From 药品收发记录 X Where 单据 = 4 And 药品id = [1]" & _
                " Group By no,药品id, nvl(批次,0),序号, 产地, 批号, 效期, 扣率, 成本价, 零售价, 摘要, 库房id, 入出类别id, x.生产日期, x.批准文号, x.外观" & _
                " Having Sum(实际数量) <> 0) A, 药品规格 B, 收费项目别名 E, 收费项目目录 F, 部门表 G," & _
                "(Select NO, 序号, 摘要, 记录状态, 填制人, 填制日期, 审核人, 审核日期 From 药品收发记录" & _
                " Where 单据 = 4 And 药品id = [1] And (记录状态 = 1 Or Mod(记录状态, 3) = 0)) X" & _
                " Where a.no=x.no and a.序号=x.序号 and a.药品id = b.药品id And b.药品id = f.Id And " & _
                "a.库房id = g.Id And b.药品id = e.收费细目id(+) And e.性质(+) = 3 And e.码类(+) = 1 and a.库房id=[2]) w, 药品库存 Z" & _
                " Where w.药品id = z.药品id(+) And Nvl(w.批次, 0) = Nvl(z.批次(+), 0) And z.库房id(+) = [2] And z.性质(+) = 1"
        Case 模块号.其他出库
            Select Case mintUnit
                Case MCONINTPRICEUNIT
                    strUnitQuantity = "F.计算单位 AS 单位, A.填写数量 as 填写数量,a.成本价,a.零售价,nvl(a.单量,0) As 外调价,'1' as 比例系数,"
                Case MCONINTOUTUNIT
                    strUnitQuantity = "B.门诊单位 AS 单位,(A.填写数量 / B.门诊包装) AS 填写数量,a.成本价*B.门诊包装 as 成本价,a.零售价*B.门诊包装 as 零售价,nvl(a.单量,0)*B.门诊包装 As 外调价,B.门诊包装 as 比例系数,"
                Case MCONINTINUNIT
                    strUnitQuantity = "B.住院单位 AS 单位,(A.填写数量 / B.住院包装) AS 填写数量,a.成本价*B.住院包装 as 成本价,a.零售价*B.住院包装 as 零售价,nvl(a.单量,0)*B.住院包装 As 外调价,B.住院包装 as 比例系数,"
                Case MCONINTSTOREUNIT
                    strUnitQuantity = "B.药库单位 AS 单位,(A.填写数量 / B.药库包装) AS 填写数量,a.成本价*B.药库包装 as 成本价,a.零售价*B.药库包装 as 零售价,nvl(a.单量,0)*B.药库包装 As 外调价,B.药库包装 as 比例系数,"
            End Select
        
            strsql = "Select w.*, z.可用数量, z.实际金额, z.实际差价" & _
                    " From (Select Distinct a.no,a.药品id,x.记录状态,x.填制人, x.填制日期,x. 审核人,x. 审核日期, a.序号, '[' || f.编码 || ']' As 药品编码, f.名称 As 通用名," & _
                    "e.名称 As 商品名, b.药品来源, b.基本药物, f.规格,f.产地 As 原产地, a.产地, a.批号, a.批次," & _
                    "b.加成率, a.效期, g.名称 As 外调单位, h.名称 As 外销单位, a.增值税率," & strUnitQuantity & _
                    "a.成本金额, 0 零售金额,0 差价, a.摘要, a.库房id,a.入出类别id, f.是否变价, b.药房分批 As 药房分批核算," & _
                    "a.批准文号 From (Select Min(ID) As ID, Sum(实际数量) As 填写数量, Sum(成本金额) As 成本金额, no,药品id," & _
                    "序号, 产地, 批号, 效期, Nvl(批次, 0) 批次, 扣率,成本价, 零售价, 摘要, 库房id, 入出类别id, 单量, 发药窗口," & _
                    "批准文号,To_Number(Trim(To_Char(Nvl(频次, '0'), '999999999999.0000'))) As 增值税率 From 药品收发记录 X" & _
                    " Where 单据 = 11 And 药品id = [1] Group By no,药品id, 序号, 产地, 批号, 效期, Nvl(批次, 0), 扣率, 成本价," & _
                    "零售价, 摘要, 库房id, 入出类别id, 单量, 发药窗口, 批准文号," & _
                    "To_Number(Trim(To_Char(Nvl(频次, '0'), '999999999999.0000')))" & _
                    " Having Sum(实际数量) <> 0) A, 药品规格 B, 收费项目别名 E, 收费项目目录 F, 药品外调单位 G, 药品外销单位 H," & _
                    "(Select NO, 序号, 摘要, 记录状态, 填制人, 填制日期, 审核人, 审核日期 From 药品收发记录" & _
                    " Where 单据 = 11 And 药品id = [1] And (记录状态 = 1 Or Mod(记录状态, 3) = 0)) X" & _
                    " Where a.no=x.no and a.序号=x.序号 and a.药品id = b.药品id And b.药品id = f.Id And a.发药窗口 = g.编码(+)" & _
                    " And a.发药窗口 = h.编码(+) And b.药品id = e.收费细目id(+) And e.性质(+) = 3 And e.码类(+) = 1) W," & _
                    " (Select 药品id, Nvl(批次, 0) 批次, 可用数量, 实际金额, 实际差价 From 药品库存" & _
                    " Where 库房id = [2] And 性质 = 1) Z" & _
                    " Where w.药品id = z.药品id(+) And Nvl(w.批次, 0) = Nvl(z.批次(+), 0)"
        Case 模块号.外购入库
            Select Case mintUnit
                Case MCONINTPRICEUNIT
                    strUnitQuantity = "D.计算单位 AS 售价单位,D.计算单位 AS 单位, A.填写数量 AS 填写数量,'1' as 比例系数, "
                    int包装系数 = "1"
                Case MCONINTOUTUNIT
                    strUnitQuantity = "D.计算单位 AS 售价单位,B.门诊单位 AS 单位,(A.填写数量 / B.门诊包装) AS 填写数量,B.门诊包装 as 比例系数,"
                    int包装系数 = "B.门诊包装"
                Case MCONINTINUNIT
                    strUnitQuantity = "D.计算单位 AS 售价单位,B.住院单位 AS 单位,(A.填写数量 / B.住院包装) AS 填写数量,B.住院包装 as 比例系数,"
                    int包装系数 = "B.住院包装"
                Case MCONINTSTOREUNIT
                    strUnitQuantity = "D.计算单位 AS 售价单位,B.药库单位 AS 单位,(A.填写数量 / B.药库包装) AS 填写数量,B.药库包装 as 比例系数,"
                    int包装系数 = "B.药库包装"
            End Select
        
            strsql = "Select w.*,z.可用数量 / w.比例系数 可用数量" & _
                    " From (Select Distinct a.no,a.药品id, a.序号, x.记录状态, x.填制人, x.填制日期, x.审核人, x.审核日期, '[' || d.编码 || ']' As 药品编码," & _
                    "d.名称 As 通用名, e.名称 As 商品名, b.药品来源, b.基本药物, d.规格,d.产地 As 原产地, a.产地, a.批号, Nvl(b.招标药品, 0) 招标药品," & _
                    "Nvl(b.差价让利比, 0) 差价让利比, b.最大效期, a.效期," & strUnitQuantity & _
                    " nvl(A.单量,b.指导批发价)*" & int包装系数 & " AS 指导批发价 ,A.成本价*" & int包装系数 & " AS 成本价," & _
                    " A.成本金额 AS 采购金额,D.是否变价,B.药房分批 药房分批核算,  " & _
                    " DECODE(A.扣率, NULL, 0, A.扣率) AS 扣率, A.零售价*" & int包装系数 & " AS 零售价 ,0 AS 零售金额,0 AS 差价,A.金额差, " & _
                    "a.批准文号, a.随货单号, a.发票号,a.发票代码, a.发票日期, a.发票金额,a.供药单位id, f.名称 As 供应商, a.库房id," & _
                    "g.名称 As 部门, Nvl(a.付款序号, 0) As 付款序号, a.退货, a.生产日期, a.批次,a.配药人 As 核查人," & _
                    "a.配药日期 As 核查日期, b.药价级别, a.加成率 From (Select Min(x.Id) As ID, Sum(实际数量) As 填写数量," & _
                    "Sum(成本金额) As 成本金额, 随货单号, 发票号,发票代码, 发票日期, Sum(发票金额) As 发票金额,x.no,x.药品id, x.序号," & _
                    "x.产地, x.批号, x.效期, x.扣率, x.成本价, x.零售价, x.单量, x.供药单位id, 库房id," & _
                    "Nvl(y.付款序号, 0) As 付款序号, Nvl(x.发药方式, 0) As 退货, x.生产日期, x.批准文号, Nvl(x.批次, 0) 批次," & _
                    "x.配药人, x.配药日期,Sum(To_Number(Nvl(用法, 0))) As 金额差, 频次 As 加成率 From 药品收发记录 X," & _
                    "(Select 收发id, 付款序号, 随货单号, 发票号,发票代码, 发票日期, 发票金额 From 应付记录" & _
                    " Where 系统标识 = 1 And 记录性质 =0) Y" & _
                    " Where x.Id = y.收发id(+) And  单据 = 1 and 药品id=[1]" & _
                    " Group By x.no,x.药品id, x.序号, x.产地, x.批号, x.效期, x.扣率, x.成本价, x.零售价, x.单量, x.供药单位id, x.库房id, 随货单号, 发票号,发票代码, 发票日期," & _
                    "Nvl(y.付款序号, 0), Nvl(x.发药方式, 0), x.生产日期, x.批准文号, Nvl(x.批次, 0), x.配药人, x.配药日期, x.频次" & _
                    " Having Sum(实际数量) <> 0) A, 药品规格 B, 收费项目别名 E, 收费项目目录 D, 供应商 F, 部门表 G," & _
                    "(Select NO, 序号, 摘要, 记录状态, 填制人, 填制日期, 审核人, 审核日期 From 药品收发记录" & _
                    " Where 单据 = 1 And 药品id = [1] And (记录状态 = 1 Or Mod(记录状态, 3) = 0)) X" & _
                    " Where a.No = x.No And a.序号 = x.序号 And a.药品id = b.药品id And b.药品id = d.Id And a.库房id = g.Id And" & _
                    " b.药品id = e.收费细目id(+) and a.库房id=[2] And e.性质(+) = 3 And a.供药单位id = f.Id And Substr(f.类型, 1, 1) = 1) w, 药品库存 Z" & _
                    " Where w.药品id = z.药品id(+) And Nvl(w.批次, 0) = Nvl(z.批次(+), 0) And z.库房id(+) = [2] And z.性质(+) = 1"
    End Select

    If Me.DTP开始时间.Value <> "" And Me.DTP结束时间.Value <> "" Then
        strsql = strsql + " and w.审核日期>=to_date('" & Me.DTP开始时间.Value & "','yyyy-mm-dd') and w.审核日期<=to_date('" & Me.DTP结束时间.Value & " 23:59:59','yyyy-mm-dd HH24:MI:SS')"
    End If
    
    If Me.txt审核人.Text <> "" Then
        strsql = strsql + " and w.审核人 like '%" & Me.txt审核人.Text & "%'"
    End If
    
    If Me.txt开始NO.Text <> "" Then
        If Me.txt结束NO.Text <> "" Then
            strsql = strsql + " and w.NO>='" & Me.txt开始NO.Text & "'  and w.NO<='" & Me.txt结束NO.Text & "'"
        Else
            strsql = strsql + " and w.NO='" & Me.txt开始NO.Text & "'"
        End If
    Else
        If Me.txt结束NO.Text <> "" Then
            strsql = strsql + " and w.NO='" & Me.txt结束NO.Text & "'"
        End If
    End If
    
    

    If Me.txt填制人.Text <> "" Then
        strsql = strsql + " and w.填制人 like '%" & Me.txt填制人.Text & "%'"
    End If
    
    If Me.DTP开始填制.Value <= Date And Me.DTP结束填制.Value <= Date Then
        strsql = strsql + " and w.填制日期>=to_date('" & Me.DTP开始填制.Value & "','yyyy-mm-dd') and w.填制日期<=to_date('" & Me.DTP结束填制.Value & " 23:59:59','yyyy-mm-dd HH24:MI:SS')"
    End If
    
    strsql = strsql + " ORDER BY NO," & strSqlOrder
    Set rsData = zldatabase.OpenSQLRecord(strsql, Me.Caption, Me.txt药品.Tag, Me.cbo库房.ItemData(Me.cbo库房.ListIndex))
    
    If rsData.RecordCount > 0 Then
       '加载药品的基本信息
        str药品信息 = rsData!药品编码
        str药品信息 = str药品信息 & IIf(gint药品名称显示 <> 1, rsData!通用名, "")
        str药品信息 = str药品信息 & IIf(gint药品名称显示 <> 0 And zlStr.Nvl(rsData!商品名) <> "", "(" & zlStr.Nvl(rsData!商品名) & ")", "")
        str药品信息 = str药品信息 & "   " & zlStr.Nvl(rsData!规格)
        str药品信息 = str药品信息 & "   (" & zlStr.Nvl(rsData!单位) & ")"
        
        Me.lbl药品基本信息.Caption = str药品信息
        
        Me.vsfList.rows = rsData.RecordCount + 1
        For i = 1 To rsData.RecordCount
            With Me.vsfList
                .TextMatrix(i, mintcol药品id) = rsData!药品ID
                .TextMatrix(i, mIntColNO) = rsData!NO
                
               '药品名称显示方式
                If gint药品名称显示 = 1 Then
                    .TextMatrix(i, mintcol药品名称与编码) = rsData!药品编码
                Else
                    .TextMatrix(i, mintcol药品名称与编码) = rsData!药品编码 & rsData!通用名
                End If
                
                .TextMatrix(i, mintcol药名) = zlStr.Nvl(rsData!通用名)
                .TextMatrix(i, mintcol商品名) = zlStr.Nvl(rsData!商品名)
                .TextMatrix(i, mintcol药品来源) = zlStr.Nvl(rsData!药品来源)
                .TextMatrix(i, mintcol基本药物) = zlStr.Nvl(rsData!基本药物)
                .TextMatrix(i, mintcol规格) = zlStr.Nvl(rsData!规格)
                .TextMatrix(i, mintcol产地) = zlStr.Nvl(rsData!产地)
                .TextMatrix(i, mintcol单位) = zlStr.Nvl(rsData!单位)
                .TextMatrix(i, mintcol批号) = zlStr.Nvl(rsData!批号)
                .TextMatrix(i, mintcol有效期至) = Format(zlStr.Nvl(rsData!效期), "yyyy-mm-dd")
                .TextMatrix(i, mintcol填制人) = zlStr.Nvl(rsData!填制人)
                .TextMatrix(i, mintcol填制日期) = Format(zlStr.Nvl(rsData!填制日期), "yyyy-MM-dd hh:mm:ss")
                .TextMatrix(i, mintcol审核人) = zlStr.Nvl(rsData!审核人)
                .TextMatrix(i, mintcol审核日期) = Format(zlStr.Nvl(rsData!审核日期), "yyyy-MM-dd hh:mm:ss")
                .TextMatrix(i, mintcol数量) = zlStr.FormatEx(Val(rsData!填写数量), mintNumberDigit, , True)
                .TextMatrix(i, mintcol冲销数量) = 0
                .TextMatrix(i, mintcol成本价) = zlStr.FormatEx(Val(rsData!成本价), mintCostDigit, , True)
                .TextMatrix(i, mintcol成本金额) = 0
                .TextMatrix(i, mintcol售价) = zlStr.FormatEx(Val(rsData!零售价), mintPriceDigit, , True)
                .TextMatrix(i, mintcol售价金额) = 0
                .TextMatrix(i, mintcol差价) = 0
                .TextMatrix(i, mintcol序号) = Val(rsData!序号)
                .TextMatrix(i, mintcol比例系数) = Val(rsData!比例系数)
                .TextMatrix(i, mintcol记录状态) = Val(rsData!记录状态)
                .TextMatrix(i, mintcol可用数量) = zlStr.FormatEx(zlStr.Nvl(rsData!可用数量, 0), mintNumberDigit, , True)
                .TextMatrix(i, mintcol是否变价) = zlStr.Nvl(rsData!是否变价, 0)
                
                If gtype_UserSysParms.P149_效期显示方式 = 0 And .TextMatrix(i, mintcol有效期至) <> "" Then
                    '换算为失效期
                    .TextMatrix(i, mintcol有效期至) = Format(DateAdd("D", 1, .TextMatrix(i, mintcol有效期至)), "yyyy-mm-dd")
                End If
                
                If mlng模块号 = 模块号.外购入库 Or mlng模块号 = 模块号.其他入库 Then .TextMatrix(i, mintcol药价级别) = zlStr.Nvl(rsData!药价级别)
                If mlng模块号 = 模块号.外购入库 Or mlng模块号 = 模块号.其他入库 Then .TextMatrix(i, mintcol生产日期) = Format(zlStr.Nvl(rsData!生产日期), "yyyy-mm-dd")

                If mlng模块号 = 模块号.药品领用 Or mlng模块号 = 模块号.药品移库 Then .TextMatrix(i, mintcol部门名称) = zlStr.Nvl(rsData!对方库房)
                If mlng模块号 <> 模块号.外购入库 Then .TextMatrix(i, mintcol批准文号) = zlStr.Nvl(rsData!批准文号)
                If mlng模块号 = 模块号.其他入库 Then .TextMatrix(i, mintcol其他外观) = zlStr.Nvl(rsData!外观)

                If mlng模块号 = 模块号.外购入库 Then
                    .TextMatrix(i, mintcol采购价) = zlStr.FormatEx(Val(rsData!成本价) / (Val(rsData!扣率 / 100)), mintCostDigit, , True)
                    .TextMatrix(i, mintcol采购限价) = zlStr.FormatEx(Val(rsData!指导批发价), mintCostDigit, , True)
                    .TextMatrix(i, mintcol扣率) = Val(rsData!扣率)
                    .TextMatrix(i, mintcol加成率) = Val(rsData!加成率) * 100 & "%"
                End If
                
                If mlng模块号 <> 模块号.其他入库 Then .TextMatrix(i, mintcol批次) = Val(rsData!批次)
                If mlng模块号 = 模块号.药品移库 Or mlng模块号 = 模块号.其他入库 Then .TextMatrix(i, mintcol真实数量) = zlStr.FormatEx(Val(rsData!真实数量), mintNumberDigit, , True)
                
                If mlng模块号 = 模块号.外购入库 Or mlng模块号 = 模块号.其他入库 Then
                    .TextMatrix(i, mintcol实际差价) = zlStr.FormatEx(.TextMatrix(i, mintcol售价金额) - .TextMatrix(i, mintcol成本金额), mintMoneyDigit, , True)
                Else
                    .TextMatrix(i, mintcol实际差价) = zlStr.FormatEx(zlStr.Nvl(rsData!实际差价, 0), mintMoneyDigit, , True)
                End If
                If mlng模块号 <> 模块号.外购入库 And mlng模块号 <> 模块号.其他入库 Then .TextMatrix(i, mintcol实际金额) = zlStr.FormatEx(zlStr.Nvl(rsData!实际金额, 0), mintMoneyDigit, , True)
                
                If mlng模块号 = 模块号.药品移库 Then .TextMatrix(i, mintcol对方部门) = Val(rsData!对方部门id)
                
                If mlng模块号 = 模块号.外购入库 Then
                    If Val(.TextMatrix(i, mintcol是否变价)) = 1 And zlStr.Nvl(.TextMatrix(i, mintcol批次), 0) <> 0 Then
                        .TextMatrix(i, mintcol零售价) = zlStr.FormatEx(rsData!零售价 / Val(rsData!比例系数), gtype_UserDrugDigits.Digit_零售价, , True)
                        .TextMatrix(i, mintcol零售单位) = rsData!售价单位
                        .TextMatrix(i, mintcol零售金额) = 0
                        .TextMatrix(i, mintcol零售差价) = 0
                     End If
                    .TextMatrix(i, mintcol外购批准文号) = zlStr.Nvl(rsData!批准文号)
                    .TextMatrix(i, mintcol随货单号) = zlStr.Nvl(rsData!随货单号)
                    .TextMatrix(i, mintcol发票号) = zlStr.Nvl(rsData!发票号)
                    .TextMatrix(i, mintcol发票代码) = zlStr.Nvl(rsData!发票代码)
                    .TextMatrix(i, mintcol发票信息) = zlStr.Nvl(rsData!发票日期)
                    .TextMatrix(i, mintcol发票金额) = zlStr.FormatEx(zlStr.Nvl(rsData!发票金额, 0), mintMoneyDigit, , True)
                End If
                .TextMatrix(i, mintcol摘要) = ""
            End With
            
            If Not rsData.EOF Then rsData.MoveNext
        Next
    End If
    
    If vsfList.rows > 1 Then
        vsfList.Row = 1
        Me.combars.Item(1).Controls.Item(MINTBTNALLWRITEOFF).Enabled = True
        Me.combars.Item(1).Controls.Item(MINTBTNALLELIMINATE).Enabled = True
        Me.combars.Item(1).Controls.Item(MINTBTNDEL).Enabled = True
        Me.vsfList.Cell(flexcpFontBold, 1, mintcol冲销数量, Me.vsfList.rows - 1, mintcol冲销数量) = True
    Else
        Me.combars.Item(1).Controls.Item(MINTBTNALLWRITEOFF).Enabled = False
        Me.combars.Item(1).Controls.Item(MINTBTNALLELIMINATE).Enabled = False
        Me.combars.Item(1).Controls.Item(MINTBTNDEL).Enabled = False
    End If
    
    Exit Sub
errRow:
    If ErrCenter <> 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub combars_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Control.Id = MINTBTNCONPLETE Then
        If Control.Checked Then
            Control.IconId = 90003
        Else
            Control.IconId = 90004
        End If
    ElseIf Control.Id = MINTBTNSIMPLE Then
        If Control.Checked Then
            Control.IconId = 90003
        Else
            Control.IconId = 90004
        End If
    End If
End Sub

Private Sub Form_Load()
    Call SetTitle
    '恢复窗口
    If Val(zldatabase.GetPara("使用个性化风格")) = 1 Then
        mstr列名列宽 = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Caption, "列名列宽", "")
    End If
    
    '由于药品申领记录和药品移库的记录相同，所以当在申领冲销时，内部就用移库的方式处理
    If mlng模块号 = 1341 Then
        mlng模块号 = 模块号.药品移库
    End If
    
    Call InitComman
    Call InitTool
    Call InitTask
    Call InitCbo
    Call initGrid
    Call InitVSFColSel
    
    '因为没有数据，所以操作数据的按钮不可用
    Me.combars.Item(1).Controls.Item(MINTBTNALLWRITEOFF).Enabled = False
    Me.combars.Item(1).Controls.Item(MINTBTNALLELIMINATE).Enabled = False
    Me.combars.Item(1).Controls.Item(MINTBTNWRITEOFF).Enabled = False
    Me.combars.Item(1).Controls.Item(MINTBTNDEL).Enabled = False

    Me.DTP开始时间.Value = DateSerial(Year(Sys.Currentdate), Month(Sys.Currentdate), Day(Sys.Currentdate))
    Me.DTP结束时间.Value = DateSerial(Year(Sys.Currentdate), Month(Sys.Currentdate), Day(Sys.Currentdate))
    Me.DTP开始时间.Enabled = False
    Me.DTP结束时间.Enabled = False
    
    Me.DTP开始填制.Value = DateSerial(Year(Sys.Currentdate), Month(Sys.Currentdate), Day(Sys.Currentdate)) + 1
    Me.DTP结束填制.Value = DateSerial(Year(Sys.Currentdate), Month(Sys.Currentdate), Day(Sys.Currentdate)) + 1
    Me.DTP开始填制.Enabled = False
    Me.DTP结束填制.Enabled = False
    
    Call combars_Execute(Me.combars.Item(1).Controls.Item(MINTBTNSIMPLE))
    mbln下可用数量 = (gtype_UserSysParms.P96_药品填单下可用库存 = 1)
End Sub

Private Sub SetTitle()
'------------------------------
'根据不同的业务设置窗体的标题
'-------------------------------
    Select Case mlng模块号
        Case 模块号.药品移库
            Me.Caption = "药品移库批量冲销"
            mInt单据号 = 单据号.药品移库
            MStrCaption = "药品移库管理"
        Case 模块号.其他出库
            Me.Caption = "药品其他出库批量冲销"
            mInt单据号 = 单据号.其他出库
            MStrCaption = "药品其他出库管理"
        Case 模块号.其他入库
            Me.Caption = "药品其他入库批量冲销"
            mInt单据号 = 单据号.其他入库
            MStrCaption = "药品其他入库管理"
        Case 模块号.外购入库
            Me.Caption = "药品外购入库批量冲销"
            mInt单据号 = 单据号.外购入库
            MStrCaption = "药品外购入库管理"
        Case 模块号.药品领用
            Me.Caption = "药品领用批量冲销"
            mInt单据号 = 单据号.药品领用
            MStrCaption = "药品领用管理"
        Case 1341
            Me.Caption = "药品申领批量冲销"
            mInt单据号 = 单据号.药品移库
            MStrCaption = "药品申领管理"
    End Select
    
End Sub


Private Sub InitVSFColSel()
'-----------------------------------------
'设置列选择的列，以及不能进行选择的列
'-----------------------------------------
    Dim rows As Integer
    Dim i As Integer
    Dim sum As Integer
    
    For i = 1 To Me.vsfList.Cols - 1
        If Me.vsfList.ColHidden(i) = False Then
            With vsfColSel
                .rows = .rows + 1
                .TextMatrix(.rows - 2, 1) = Me.vsfList.TextMatrix(0, i)
                .RowData(.rows - 2) = i
            End With
        End If
    Next
    
    sum = 6
    If Me.vsfList.ColHidden(mintcol部门名称) = False Then sum = 7
    
    Me.vsfColSel.rows = vsfColSel.rows - 1
    For i = 1 To sum
        Me.vsfColSel.Cell(flexcpForeColor, i, 0, i, 1) = CSTCOLOR_FIXED
    Next
End Sub

Private Sub InitComman()
'--------------------------------------
'初始化CommandBars1控件

'--------------------------------------
    With CommandBarsGlobalSettings
        Set .App = App
        .ResourceFile = .OcxPath & "\XTPResourceZhCn.dll" '设置中文语言资源文件
        .ColorManager.SystemTheme = xtpSystemThemeAuto  '控件整体的颜色方案，根据系统自动识别
    End With

    With combars.Options
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

    With combars
        .VisualTheme = xtpThemeOffice2003 '设置控件显示风格
        .EnableCustomization False '是否允许自定义设置
        .Item(1).Delete
        .Icons = Me.imgicon.Icons
    End With
End Sub


Private Sub InitTool()
'-----------------------------------------------------
'设置工具栏
'----------------------------------------------------
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    
    Set objBar = combars.Add("工具栏1", xtpBarTop)
    objBar.ContextMenuPresent = False '工具栏上点击鼠标右键时不弹出设置菜单
    objBar.ShowTextBelowIcons = False '工具栏中的按钮文字显示在图标右侧
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, MINTBTNFILTER, "过滤")
        objControl.Style = xtpButtonIconAndCaption '同时显示图标和文字
        Set objControl = .Add(xtpControlButton, MINTBTNALLWRITEOFF, "全冲")
        objControl.BeginGroup = True
        objControl.Style = xtpButtonIconAndCaption '同时显示图标和文字
        Set objControl = .Add(xtpControlButton, MINTBTNALLELIMINATE, "全清")
        objControl.Style = xtpButtonIconAndCaption '同时显示图标和文字
        Set objControl = .Add(xtpControlButton, MINTBTNDEL, "删除")
        objControl.Style = xtpButtonIconAndCaption '同时显示图标和文字
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, MINTBTNWRITEOFF, "冲销")
        objControl.Style = xtpButtonIconAndCaption '同时显示图标和文字
        Set objControl = .Add(xtpControlButton, MINTBTNHELP, "帮助")
        objControl.Style = xtpButtonIconAndCaption '同时显示图标和文字
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, MINTBTNEXIT, "退出")
        objControl.Style = xtpButtonIconAndCaption '同时显示图标和文字
        
        Set objControl = .Add(xtpControlButton, MINTBTNSIMPLE, "简洁")
            objControl.BeginGroup = True
            objControl.Flags = xtpFlagRightAlign
            objControl.Style = xtpButtonIconAndCaption '同时显示图标和文字
            objControl.Checked = True
        Set objControl = .Add(xtpControlButton, MINTBTNCONPLETE, "完整")
            objControl.Flags = xtpFlagRightAlign
            objControl.Style = xtpButtonIconAndCaption '同时显示图标和文字
    End With
End Sub
Private Sub InitTask()
'---------------------------------------
'初始化任务面板
'----------------------------------------
    Dim objGroup As TaskPanelGroup
    Dim objItem As TaskPanelGroupItem
   
    Call tkpMain.SetMargins(0, 0, 0, 0, 0)
    Call tkpMain.SetItemInnerMargins(0, 0, 0, 0)
    Call tkpMain.SetItemOuterMargins(0, 0, 0, 0)
    Call tkpMain.SetGroupInnerMargins(0, 0, 0, 0)
    Call tkpMain.SetGroupOuterMargins(3, 3, 3, 0)
        
    Set objGroup = tkpMain.Groups.Add(1, "基本条件")
    objGroup.Expandable = False '不能收缩
    Set objItem = objGroup.Items.Add(0, "", xtpTaskItemTypeControl)
    Set objItem.Control = pic基本信息
    pic基本信息.BackColor = objItem.BackColor
   
    Set objGroup = tkpMain.Groups.Add(2, "其他条件")
    objGroup.Expandable = True
    Set objItem = objGroup.Items.Add(0, "", xtpTaskItemTypeControl)
    Set objItem.Control = pic其他信息
    pic其他信息.BackColor = objItem.BackColor
    objGroup.Expanded = False  '没有打开

End Sub
Private Sub Form_Resize()
    On Error Resume Next
    
    Me.tkpMain.Move 0, 530, Me.tkpMain.Width, Me.ScaleHeight - 530 - Me.staThis.Height
    Me.fraEW.Move Me.tkpMain.Left + Me.tkpMain.Width, Me.tkpMain.Top, 45, Me.tkpMain.Height
    Me.vsfList.Move Me.fraEW.Left + Me.fraEW.Width, Me.fraEW.Top + Me.fra药品基本信息.Height + 50, Me.ScaleWidth - (Me.fraEW.Left + Me.fraEW.Width), Me.tkpMain.Height - Me.fra药品基本信息.Height
    Me.fra药品基本信息.Move vsfList.Left, Me.fraEW.Top - 30, vsfList.Width, Me.fra药品基本信息.Height

    fraColSel.Left = Me.tkpMain.Width + Me.tkpMain.Left - fraColSel.Width + 265
    fraColSel.Top = (vsfList.RowHeight(0) - fraColSel.Height) / 2 + 540 + Me.fra药品基本信息.Height + 50
    fraColSel.ZOrder
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim str列名列宽 As String
    Dim i As Integer
    
    If Val(zldatabase.GetPara("使用个性化风格")) = 1 Then
        For i = 0 To MINTCOL总列数 - 1
            str列名列宽 = str列名列宽 & vsfList.ColKey(i) & "," & vsfList.ColWidth(i) & "|"
        Next
        Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Caption, "列名列宽", str列名列宽)
    End If
    mstr列名列宽 = ""
    Call ReleaseSelectorRS
End Sub

Private Sub fraEW_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'------------------------------------------
'条件区和单据区的拉动
'------------------------------------------
    On Error Resume Next
    If Me.tkpMain.Width + x < 1700 Or Me.vsfList.Width - x < 500 Then
        Exit Sub
    End If
    
    If Button = 1 Then
        Me.fraEW.Move Me.fraEW.Left + x, Me.fraEW.Top, Me.fraEW.Width, Me.fraEW.Height
        Me.pic基本信息.Move Me.pic基本信息.Left, Me.pic基本信息.Top, Me.pic基本信息.Width + x, Me.pic基本信息.Height
        Me.tkpMain.Move Me.tkpMain.Left, Me.tkpMain.Top, Me.tkpMain.Width + x, Me.tkpMain.Height
        Me.vsfList.Move Me.vsfList.Left + x, Me.vsfList.Top, Me.vsfList.Width - x, Me.vsfList.Height
        Me.fra药品基本信息.Move Me.fra药品基本信息.Left + x, Me.fra药品基本信息.Top, Me.fra药品基本信息.Width - x, Me.fra药品基本信息.Height
        fraColSel.Left = fraColSel.Left + x
        Me.cbo库房.Width = Me.cbo库房.Width + x
        Me.txt结束NO.Width = Me.txt结束NO.Width + x
        Me.txt开始NO.Width = Me.txt开始NO.Width + x
        Me.txt审核人.Width = Me.txt审核人.Width + x
        Me.txt填制人.Width = Me.txt填制人.Width + x
        Me.txt药品.Width = Me.txt药品.Width + x
        Me.cmd药品.Left = Me.cmd药品.Left + x
        Me.DTP结束时间.Width = Me.DTP结束时间.Width + x
        Me.DTP结束填制.Width = Me.DTP结束填制.Width + x
        Me.DTP开始时间.Width = Me.DTP开始时间.Width + x
        Me.DTP开始填制.Width = Me.DTP开始填制.Width + x
        Me.cbo时间范围.Width = Me.cbo时间范围.Width + x
        Me.cbo填制时间.Width = Me.cbo填制时间.Width + x
    End If
End Sub
Private Sub InitCbo()
    '-----------------
    '初始化下拉框
    '-----------------
    With Me.cbo时间范围
        .Clear
        .AddItem "一天内"
        .AddItem "两天内"
        .AddItem "三天内"
        .AddItem "指定时间范围"
    End With
    
    With Me.cbo填制时间
        .Clear
        .AddItem "一天内"
        .AddItem "两天内"
        .AddItem "三天内"
        .AddItem "指定时间范围"
    End With
    
    Me.cbo时间范围.ListIndex = 0
End Sub
Private Sub initGrid()
'----------------------------------
'初始化表格列
'----------------------------------
    Dim i As Integer
    Dim arr列设置
    
    mintcol选择 = 0
    mintcol药品id = 1
    mintcol行号 = 2
    mIntColNO = 3
    mintcol部门名称 = 4
    mintcol药品名称与编码 = 5
    mintcol商品名 = 6
    mintcol药品来源 = 7
    mintcol基本药物 = 8
    mintcol药价级别 = 9
    mintcol规格 = 10
    mintcol单位 = 11
    mintcol数量 = 12
    mintcol冲销数量 = 13
    mintcol产地 = 14
    mintcol批号 = 15
    mintcol摘要 = 16
    mintcol生产日期 = 17
    mintcol有效期至 = 18
    mintcol批准文号 = 19
    mintcol其他外观 = 20 '其他入库的字段
    mintcol填制人 = 21
    mintcol填制日期 = 22
    mintcol审核人 = 23
    mintcol审核日期 = 24
    mintcol采购限价 = 25
    mintcol采购价 = 26
    mintcol扣率 = 27
    mintcol成本价 = 28 '外购入库则为结算价
    mintcol成本金额 = 29 '外购入库则为结算金额
    mintcol加成率 = 30
    mintcol售价 = 31
    mintcol售价金额 = 32
    mintcol差价 = 33
    '外购入库需要的字段
    mintcol零售价 = 34
    mintcol零售单位 = 35
    mintcol零售金额 = 36
    mintcol零售差价 = 37
    mintcol外购批准文号 = 38
    mintcol随货单号 = 39
    mintcol发票号 = 40
    mintcol发票代码 = 41
    mintcol发票信息 = 42
    mintcol发票金额 = 43
    '需要隐藏的列
    mintcol真实数量 = 44
    mintcol序号 = 45
    mintcol比例系数 = 46
    mintcol批次 = 47
    mintcol记录状态 = 48
    mintcol分批核算 = 49
    mintcol可用数量 = 50
    mintcol最大效期 = 51
    mintcol实际差价 = 52
    mintcol实际金额 = 53
    mintcol上次供应商ID = 54
    mintcol对方部门 = 55
    mintcol药名 = 56
    mintcol是否变价 = 57
    
    With Me.vsfList
        .rows = 1
        .Cols = MINTCOL总列数
        If mstr列名列宽 <> "" Then
            arr列设置 = Split(mstr列名列宽, "|")
            If UBound(arr列设置) <> MINTCOL总列数 Then
                mstr列名列宽 = ""
            Else
                For i = 0 To UBound(arr列设置) - 1
                    SetColValue Split(arr列设置(i), ",")(0), i, Split(arr列设置(i), ",")(1)
                Next
            End If
        End If
    
        .TextMatrix(0, mintcol选择) = ""
        .TextMatrix(0, mintcol药品id) = "药品id"
        .TextMatrix(0, mintcol行号) = "行号"
        .TextMatrix(0, mIntColNO) = "NO"
        .TextMatrix(0, mintcol药品名称与编码) = "药品名称与编码"
        .TextMatrix(0, mintcol商品名) = "商品名"
        .TextMatrix(0, mintcol药品来源) = "药品来源"
        .TextMatrix(0, mintcol基本药物) = "基本药物"
        .TextMatrix(0, mintcol药价级别) = "药价级别"
        .TextMatrix(0, mintcol规格) = "规格"
        .TextMatrix(0, mintcol产地) = "产地"
        .TextMatrix(0, mintcol单位) = "单位"
        .TextMatrix(0, mintcol批号) = "批号"
        .TextMatrix(0, mintcol生产日期) = "生产日期"
        .TextMatrix(0, mintcol有效期至) = IIf(gtype_UserSysParms.P149_效期显示方式 = 1, "有效期至", "失效期")
        .TextMatrix(0, mintcol部门名称) = "对方库房"
        .TextMatrix(0, mintcol填制人) = "填制人"
        .TextMatrix(0, mintcol填制日期) = "填制日期"
        .TextMatrix(0, mintcol审核人) = "审核人"
        .TextMatrix(0, mintcol审核日期) = "审核日期"
        .TextMatrix(0, mintcol其他外观) = "外观"
        .TextMatrix(0, mintcol批准文号) = "批准文号"
        .TextMatrix(0, mintcol数量) = "数量"
        .TextMatrix(0, mintcol冲销数量) = "冲销数量"
        .TextMatrix(0, mintcol采购限价) = "采购限价"
        .TextMatrix(0, mintcol采购价) = "采购价"
        .TextMatrix(0, mintcol扣率) = "扣率"
        .TextMatrix(0, mintcol成本价) = "成本价"
        .TextMatrix(0, mintcol成本金额) = "成本金额"
        .TextMatrix(0, mintcol加成率) = "加成率"
        .TextMatrix(0, mintcol售价) = "售价"
        .TextMatrix(0, mintcol售价金额) = "售价金额"
        .TextMatrix(0, mintcol差价) = "差价"
        .TextMatrix(0, mintcol零售价) = "零售价"
        .TextMatrix(0, mintcol零售单位) = "零售单位"
        .TextMatrix(0, mintcol零售金额) = "零售金额"
        .TextMatrix(0, mintcol零售差价) = "零售差价"
        .TextMatrix(0, mintcol外购批准文号) = "批准文号"
        .TextMatrix(0, mintcol随货单号) = "随货单号"
        .TextMatrix(0, mintcol发票号) = "发票号"
        .TextMatrix(0, mintcol发票代码) = "发票代码"
        .TextMatrix(0, mintcol发票信息) = "发票信息"
        .TextMatrix(0, mintcol发票金额) = "发票金额"

        .TextMatrix(0, mintcol真实数量) = "真实数量"
        .TextMatrix(0, mintcol序号) = "序号"
        .TextMatrix(0, mintcol比例系数) = "比例系数"
        .TextMatrix(0, mintcol药名) = "药名"
        .TextMatrix(0, mintcol批次) = "批次"
        .TextMatrix(0, mintcol记录状态) = "记录状态"
        .TextMatrix(0, mintcol分批核算) = "分批核算"
        .TextMatrix(0, mintcol可用数量) = "可用数量"
        .TextMatrix(0, mintcol最大效期) = "最大效期"
        .TextMatrix(0, mintcol实际差价) = "实际差价"
        .TextMatrix(0, mintcol实际金额) = "实际金额"
        .TextMatrix(0, mintcol上次供应商ID) = "上次供应商ID"
        .TextMatrix(0, mintcol摘要) = "摘要"
        .TextMatrix(0, mintcol对方部门) = "对方部门id"
        .TextMatrix(0, mintcol是否变价) = "是否变价"

        .ColKey(mintcol选择) = "选择"
        .ColKey(mintcol药品id) = "药品id"
        .ColKey(mintcol行号) = "行号"
        .ColKey(mIntColNO) = "NO"
        .ColKey(mintcol药品名称与编码) = "药品名称与编码"
        .ColKey(mintcol商品名) = "商品名"
        .ColKey(mintcol药品来源) = "药品来源"
        .ColKey(mintcol基本药物) = "基本药物"
        .ColKey(mintcol药价级别) = "药价级别"
        .ColKey(mintcol产地) = "产地"
        .ColKey(mintcol规格) = "规格"
        .ColKey(mintcol单位) = "单位"
        .ColKey(mintcol批号) = "批号"
        .ColKey(mintcol生产日期) = "生产日期"
        .ColKey(mintcol有效期至) = "有效期至"
        .ColKey(mintcol部门名称) = "部门名称"
        .ColKey(mintcol其他外观) = "其他外观"
        .ColKey(mintcol填制人) = "填制人"
        .ColKey(mintcol填制日期) = "填制日期"
        .ColKey(mintcol审核人) = "审核人"
        .ColKey(mintcol审核日期) = "审核日期"
        .ColKey(mintcol批准文号) = "批准文号"
        .ColKey(mintcol数量) = "数量"
        .ColKey(mintcol冲销数量) = "冲销数量"
        .ColKey(mintcol采购限价) = "采购限价"
        .ColKey(mintcol采购价) = "采购价"
        .ColKey(mintcol扣率) = "扣率"
        .ColKey(mintcol成本价) = "成本价"
        .ColKey(mintcol成本金额) = "成本金额"
        .ColKey(mintcol加成率) = "加成率"
        .ColKey(mintcol售价) = "售价"
        .ColKey(mintcol售价金额) = "售价金额"
        .ColKey(mintcol差价) = "差价"
        .ColKey(mintcol零售价) = "零售价"
        .ColKey(mintcol零售单位) = "零售单位"
        .ColKey(mintcol零售金额) = "零售金额"
        .ColKey(mintcol零售差价) = "零售差价"
        .ColKey(mintcol外购批准文号) = "外购批准文号"
        .ColKey(mintcol随货单号) = "随货单号"
        .ColKey(mintcol发票号) = "发票号"
        .ColKey(mintcol发票代码) = "发票代码"
        .ColKey(mintcol发票信息) = "发票信息"
        .ColKey(mintcol发票金额) = "发票金额"
        .ColKey(mintcol真实数量) = "真实数量"
        .ColKey(mintcol序号) = "序号"
        .ColKey(mintcol比例系数) = "比例系数"
        .ColKey(mintcol药名) = "药名"
        .ColKey(mintcol批次) = "批次"
        .ColKey(mintcol记录状态) = "记录状态"
        .ColKey(mintcol分批核算) = "分批核算"
        .ColKey(mintcol可用数量) = "可用数量"
        .ColKey(mintcol最大效期) = "最大效期"
        .ColKey(mintcol实际差价) = "实际差价"
        .ColKey(mintcol实际金额) = "实际金额"
        .ColKey(mintcol上次供应商ID) = "上次供应商ID"
        .ColKey(mintcol摘要) = "摘要"
        .ColKey(mintcol对方部门) = "对方部门"
        .ColKey(mintcol是否变价) = "是否变价"
        .ColKey(mintcol部门名称) = "部门名称"

        If mstr列名列宽 = "" Then
            .ColWidth(mintcol选择) = 270
            .ColWidth(mintcol行号) = 800
            .ColWidth(mIntColNO) = 900
            .ColWidth(mintcol药品名称与编码) = 2500
            .ColWidth(mintcol商品名) = 1100
            .ColWidth(mintcol规格) = 1100
            .ColWidth(mintcol产地) = 1100
            .ColWidth(mintcol单位) = 600
            .ColWidth(mintcol生产日期) = 1100
            .ColWidth(mintcol摘要) = 2300
            .ColWidth(mintcol有效期至) = 1100
            .ColWidth(mintcol部门名称) = 1100
            .ColWidth(mintcol批准文号) = 1100
            .ColWidth(mintcol其他外观) = 1100
            .ColWidth(mintcol填制日期) = 1100
            .ColWidth(mintcol审核日期) = 1100
            .ColWidth(mintcol数量) = 1100
            .ColWidth(mintcol冲销数量) = 1200
            .ColWidth(mintcol采购限价) = 1100
            .ColWidth(mintcol采购价) = 1100
            .ColWidth(mintcol扣率) = 1100
            .ColWidth(mintcol加成率) = 1100
            .ColWidth(mintcol成本价) = 1100
            .ColWidth(mintcol成本金额) = 1100
            .ColWidth(mintcol售价) = 1100
            .ColWidth(mintcol售价金额) = 1100
            .ColWidth(mintcol差价) = 1100
            .ColWidth(mintcol批号) = 1100
            .ColWidth(mintcol零售价) = 1100
            .ColWidth(mintcol零售单位) = 1100
            .ColWidth(mintcol零售金额) = 1100
            .ColWidth(mintcol零售差价) = 1100
            .ColWidth(mintcol外购批准文号) = 1100
            .ColWidth(mintcol随货单号) = 1100
            .ColWidth(mintcol发票号) = 1100
            .ColWidth(mintcol发票代码) = 1100
            .ColWidth(mintcol发票信息) = 1100
            .ColWidth(mintcol发票金额) = 1100
        End If
        
        '是否显示商品名
        If gint药品名称显示 <> 0 Then
            .ColHidden(mintcol商品名) = False
        Else
            .ColHidden(mintcol商品名) = True
        End If
        
        '其他入库显示的列和不显示的列
        If mlng模块号 = 模块号.其他入库 Then
            .ColHidden(mintcol其他外观) = False
        Else
            .ColHidden(mintcol其他外观) = True
        End If
        
        If mlng模块号 = 模块号.药品领用 Or mlng模块号 = 模块号.药品移库 Then
            .ColHidden(mintcol部门名称) = False
        Else
            .ColHidden(mintcol部门名称) = True
        End If
        '只有外购才显示的列
        If mlng模块号 = 模块号.外购入库 Then
            .ColHidden(mintcol采购限价) = False
            .ColHidden(mintcol采购价) = False
            .ColHidden(mintcol扣率) = False
            .ColHidden(mintcol加成率) = False
            .ColHidden(mintcol零售价) = False
            .ColHidden(mintcol零售单位) = False
            .ColHidden(mintcol零售金额) = False
            .ColHidden(mintcol零售差价) = False
            .ColHidden(mintcol外购批准文号) = False
            .ColHidden(mintcol随货单号) = False
            .ColHidden(mintcol发票号) = False
            .ColHidden(mintcol发票代码) = False
            .ColHidden(mintcol发票信息) = False
            .ColHidden(mintcol发票金额) = False
            .ColHidden(mintcol药价级别) = False
            .ColHidden(mintcol生产日期) = False
            .ColHidden(mintcol批准文号) = True
        Else
            .ColHidden(mintcol采购限价) = True
            .ColHidden(mintcol采购价) = True
            .ColHidden(mintcol扣率) = True
            .ColHidden(mintcol加成率) = True
            .ColHidden(mintcol零售价) = True
            .ColHidden(mintcol零售单位) = True
            .ColHidden(mintcol零售金额) = True
            .ColHidden(mintcol零售差价) = True
            .ColHidden(mintcol外购批准文号) = True
            .ColHidden(mintcol随货单号) = True
            .ColHidden(mintcol发票号) = True
            .ColHidden(mintcol发票代码) = True
            .ColHidden(mintcol发票信息) = True
            .ColHidden(mintcol发票金额) = True
            .ColHidden(mintcol药价级别) = True
            .ColHidden(mintcol生产日期) = True
            .ColHidden(mintcol批准文号) = False
        End If
        
        If mlng模块号 = 模块号.外购入库 Or mlng模块号 = 模块号.药品移库 Then
            .ColHidden(mintcol摘要) = False
        Else
            .ColHidden(mintcol摘要) = True
        End If
        
        '全部业务隐藏的列
        .ColHidden(mintcol药品名称与编码) = True
        .ColHidden(mintcol商品名) = True
        .ColHidden(mintcol规格) = True
        .ColHidden(mintcol单位) = True
        .ColHidden(mintcol行号) = True
        .ColHidden(mintcol药品id) = True
        .ColHidden(mintcol真实数量) = True
        .ColHidden(mintcol序号) = True
        .ColHidden(mintcol比例系数) = True
        .ColHidden(mintcol药名) = True
        .ColHidden(mintcol批次) = True
        .ColHidden(mintcol记录状态) = True
        .ColHidden(mintcol分批核算) = True
        .ColHidden(mintcol可用数量) = True
        .ColHidden(mintcol最大效期) = True
        .ColHidden(mintcol实际差价) = True
        .ColHidden(mintcol实际金额) = True
        .ColHidden(mintcol上次供应商ID) = True
        .ColHidden(mintcol对方部门) = True
        .ColHidden(mintcol是否变价) = True
        .ColHidden(mintcol药品来源) = True
        .ColHidden(mintcol药价级别) = True
        .ColHidden(mintcol基本药物) = True
        
        '列内容对齐方式
        .ColAlignment(mintcol药品id) = flexAlignRightCenter
        .ColAlignment(mIntColNO) = flexAlignLeftCenter
        .ColAlignment(mintcol药品名称与编码) = flexAlignLeftCenter
        .ColAlignment(mintcol商品名) = flexAlignLeftCenter
        .ColAlignment(mintcol药品来源) = flexAlignLeftCenter
        .ColAlignment(mintcol基本药物) = flexAlignLeftCenter
        .ColAlignment(mintcol药价级别) = flexAlignLeftCenter
        .ColAlignment(mintcol规格) = flexAlignLeftCenter
        .ColAlignment(mintcol产地) = flexAlignLeftCenter
        .ColAlignment(mintcol单位) = flexAlignLeftCenter
        .ColAlignment(mintcol批号) = flexAlignLeftCenter
        .ColAlignment(mintcol生产日期) = flexAlignLeftCenter
        .ColAlignment(mintcol有效期至) = flexAlignLeftCenter
        .ColAlignment(mintcol批准文号) = flexAlignLeftCenter
        .ColAlignment(mintcol其他外观) = flexAlignLeftCenter
        .ColAlignment(mintcol数量) = flexAlignRightCenter
        .ColAlignment(mintcol冲销数量) = flexAlignRightCenter
        .ColAlignment(mintcol成本价) = flexAlignRightCenter
        .ColAlignment(mintcol成本金额) = flexAlignRightCenter
        .ColAlignment(mintcol售价) = flexAlignRightCenter
        .ColAlignment(mintcol售价金额) = flexAlignRightCenter
        .ColAlignment(mintcol差价) = flexAlignRightCenter
        .ColAlignment(mintcol零售价) = flexAlignRightCenter
        .ColAlignment(mintcol零售单位) = flexAlignLeftCenter
        .ColAlignment(mintcol零售金额) = flexAlignRightCenter
        .ColAlignment(mintcol零售差价) = flexAlignRightCenter
        .ColAlignment(mintcol外购批准文号) = flexAlignLeftCenter
        .ColAlignment(mintcol随货单号) = flexAlignLeftCenter
        .ColAlignment(mintcol发票号) = flexAlignLeftCenter
        .ColAlignment(mintcol发票代码) = flexAlignLeftCenter
        .ColAlignment(mintcol发票信息) = flexAlignLeftCenter
        .ColAlignment(mintcol发票金额) = flexAlignLeftCenter
        .ColAlignment(mintcol摘要) = flexAlignLeftCenter
        
        '列标题对齐方式
        .FixedAlignment(mIntColNO) = flexAlignCenterCenter
        .FixedAlignment(mintcol药品名称与编码) = flexAlignCenterCenter
        .FixedAlignment(mintcol商品名) = flexAlignCenterCenter
        .FixedAlignment(mintcol药品来源) = flexAlignCenterCenter
        .FixedAlignment(mintcol基本药物) = flexAlignCenterCenter
        .FixedAlignment(mintcol药价级别) = flexAlignCenterCenter
        .FixedAlignment(mintcol规格) = flexAlignCenterCenter
        .FixedAlignment(mintcol产地) = flexAlignCenterCenter
        .FixedAlignment(mintcol单位) = flexAlignCenterCenter
        .FixedAlignment(mintcol批号) = flexAlignCenterCenter
        .FixedAlignment(mintcol生产日期) = flexAlignCenterCenter
        .FixedAlignment(mintcol有效期至) = flexAlignCenterCenter
        .FixedAlignment(mintcol批准文号) = flexAlignCenterCenter
        .FixedAlignment(mintcol其他外观) = flexAlignCenterCenter
        .FixedAlignment(mintcol部门名称) = flexAlignCenterCenter
        .FixedAlignment(mintcol填制人) = flexAlignCenterCenter
        .FixedAlignment(mintcol填制日期) = flexAlignCenterCenter
        .FixedAlignment(mintcol审核人) = flexAlignCenterCenter
        .FixedAlignment(mintcol审核日期) = flexAlignCenterCenter
        .FixedAlignment(mintcol数量) = flexAlignCenterCenter
        .FixedAlignment(mintcol冲销数量) = flexAlignCenterCenter
        .FixedAlignment(mintcol成本价) = flexAlignCenterCenter
        .FixedAlignment(mintcol成本金额) = flexAlignCenterCenter
        .FixedAlignment(mintcol售价) = flexAlignCenterCenter
        .FixedAlignment(mintcol售价金额) = flexAlignCenterCenter
        .FixedAlignment(mintcol差价) = flexAlignCenterCenter
        .FixedAlignment(mintcol零售价) = flexAlignCenterCenter
        .FixedAlignment(mintcol零售单位) = flexAlignCenterCenter
        .FixedAlignment(mintcol零售金额) = flexAlignCenterCenter
        .FixedAlignment(mintcol零售差价) = flexAlignCenterCenter
        .FixedAlignment(mintcol外购批准文号) = flexAlignCenterCenter
        .FixedAlignment(mintcol随货单号) = flexAlignCenterCenter
        .FixedAlignment(mintcol发票号) = flexAlignCenterCenter
        .FixedAlignment(mintcol发票代码) = flexAlignCenterCenter
        .FixedAlignment(mintcol发票信息) = flexAlignCenterCenter
        .FixedAlignment(mintcol发票金额) = flexAlignCenterCenter
        .FixedAlignment(mintcol摘要) = flexAlignCenterCenter
        
        .RowHeight(0) = 300
        .AllowUserResizing = flexResizeBoth
        .ExplorerBar = flexExSortShowAndMove
        
        .Cell(flexcpForeColor, 0, mintcol冲销数量) = &HFF0000
        .Cell(flexcpForeColor, 0, mintcol摘要) = &HFF0000
        .Cell(flexcpFontBold, 0, mintcol冲销数量) = True
        .Cell(flexcpFontBold, 0, mintcol摘要) = True
        
    End With
End Sub
Public Sub showMe(ByVal int模块号 As Integer, Optional ByVal FrmMain As Form, Optional strtock As String, Optional int库房Index As Integer)
    '该过程用于其他窗体打开该窗体
    '参数：int模块号：业务模块号
    'FrmMain:主窗体
    'strtock:库房字符串(格式:“库房名称1,库房id|库房名称2,库房id|......”
    'int库房Index:主窗体中库房下拉列表的listindex
    Dim i As Integer
    Dim strsql As String
    Dim rsDepend As Recordset
    Dim arr库房
    
    On Error Resume Next
    Set mfrmMain = FrmMain
    mlng模块号 = int模块号
     
    If strtock <> "" Then
        arr库房 = Split(strtock, "|")
        
        For i = 0 To UBound(arr库房) - 1
            Me.cbo库房.AddItem Split(arr库房(i), ",")(0)
            Me.cbo库房.ItemData(i) = Split(arr库房(i), ",")(1)
        Next
    End If
    
    Me.cbo库房.ListIndex = int库房Index
    
    Me.Show 1
End Sub
Private Sub imgColSel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Long
    
    If Button = 1 Then '列选择器
        '根据当前状态直接确定勾选状态
        With vsfColSel
            If .Visible Then
                .Visible = False
                vsfList.SetFocus
            Else
                For i = .FixedRows To .rows - 1
                    If vsfList.ColHidden(.RowData(i)) Or vsfList.ColWidth(.RowData(i)) = 0 Then
                        .TextMatrix(i, 0) = 0
                    Else
                        .TextMatrix(i, 0) = 1
                    End If
                Next
                
                .Height = .RowHeightMin * .rows + 150
                .Top = fraColSel.Top + fraColSel.Height
                If .Top + .Height > Me.ScaleHeight - vsfList.Top Then
                    .Height = Me.ScaleHeight - .Top - vsfList.Top
                    .Width = 1750
                Else
                    .Width = 1470
                End If
                .Left = fraColSel.Left
                .ZOrder
                .Visible = True
                .SetFocus
            End If
        End With
    End If
End Sub
Private Sub Txt结束NO_KeyDown(KeyCode As Integer, Shift As Integer)
     Dim lng库房id As Long
    Dim intNO As Integer, strNo As String
    
    '初始准备
    Select Case mlng模块号
        Case 模块号.外购入库
            intNO = 21
        Case 模块号.其他入库
            intNO = 24
        Case 模块号.药品移库
            intNO = 26
        Case 模块号.其他出库
            intNO = 28
        Case 模块号.药品领用
            intNO = 27
    End Select
    
    lng库房id = Me.cbo库房.ItemData(Me.cbo库房.ListIndex)
    If KeyCode = vbKeyReturn Then
        If Len(txt结束NO) < 8 And Len(txt结束NO) > 0 Then
            txt结束NO.Text = zlCommFun.GetFullNO(txt结束NO.Text, intNO, lng库房id)
        End If
        OS.PressKey (vbKeyTab)
    End If
    
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub
Private Sub Txt开始NO_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng库房id As Long
    Dim intNO As Integer, strNo As String
    
    '初始准备
    Select Case mlng模块号
        Case 模块号.外购入库
            intNO = 21
        Case 模块号.其他入库
            intNO = 24
        Case 模块号.药品移库
            intNO = 26
        Case 模块号.其他出库
            intNO = 28
        Case 模块号.药品领用
            intNO = 27
    End Select
    
    lng库房id = Me.cbo库房.ItemData(Me.cbo库房.ListIndex)
    
    If KeyCode = vbKeyReturn Then
        If Len(txt开始NO) < 8 And Len(txt开始NO) > 0 Then
            txt开始NO.Text = zlCommFun.GetFullNO(txt开始NO.Text, intNO, lng库房id)
        End If
        Me.txt结束NO.SetFocus
    End If
    
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub
Private Function SaveStrike() As Boolean
'-------------------------------------------
'处理冲销的过程，返回Boolean类型的值
'-------------------------------------------
    Dim 行次_IN As Integer
    Dim 原记录状态_IN As Integer
    Dim NO_IN As String
    Dim 序号_IN As Integer
    Dim 药品ID_IN As Long
    Dim 冲销数量_IN As Double
    Dim 填制人_IN As String
    Dim 填制日期_IN  As String
    Dim 发票号_IN As String
    Dim 发票代码_In As String
    Dim 发票日期_IN As Date
    Dim 发票金额_IN As Double
    Dim intRow As Integer
    Dim rstemp As New ADODB.Recordset
    Dim i As Integer
    Dim 摘要_IN As String
    Dim str药品id As String
    Dim lastNO As String
    Dim int全部冲销 As Integer
    Dim arrSql As Variant
    
    arrSql = Array()
    SaveStrike = False
    With Me.vsfList
        If Val(.TextMatrix(1, mintcol批次)) = 0 Then
            For intRow = 1 To .rows - 1
                冲销数量_IN = 冲销数量_IN + zlStr.FormatEx(.TextMatrix(intRow, mintcol冲销数量) * .TextMatrix(intRow, mintcol比例系数), gtype_UserSaleDigits.Digit_数量, , True)
            Next
        End If
        For intRow = 1 To .rows - 1
            '检查冲销数量，不能小于零
            If Val(.TextMatrix(intRow, mintcol冲销数量)) <> 0 Then
                If Not 相同符号(Val(.TextMatrix(intRow, mintcol数量)), Val(.TextMatrix(intRow, mintcol冲销数量))) Then
                    MsgBox "请输入合法的冲销数量（第" & intRow & "行）！", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            If mlng模块号 <> 模块号.其他出库 And mlng模块号 <> 模块号.药品领用 Then
                '检查可用数量是否足够，参数设置为不检查库存时不进行
                If mint库存检查入库库房 <> 0 And .TextMatrix(intRow, 1) <> "" Then
                    If Val(.TextMatrix(intRow, mintcol冲销数量)) = Val(.TextMatrix(intRow, mintcol数量)) Then
                        int全部冲销 = 1
                        If Val(.TextMatrix(1, mintcol批次)) <> 0 Then
                            冲销数量_IN = zlStr.FormatEx(同批次冲销数量(Val(.TextMatrix(intRow, mintcol批次))), gtype_UserSaleDigits.Digit_数量, , True) 'Val(.TextMatrix(intRow, mintcol数量)) * Val(.TextMatrix(intRow, mintcol比例系数))
                        End If
                    Else
                        int全部冲销 = 0
                        If Val(.TextMatrix(1, mintcol批次)) <> 0 Then
                            冲销数量_IN = zlStr.FormatEx(同批次冲销数量(Val(.TextMatrix(intRow, mintcol批次))), gtype_UserSaleDigits.Digit_数量, , True) 'zlStr.FormatEx(.TextMatrix(intRow, mintcol冲销数量) * .TextMatrix(intRow, mintcol比例系数), gtype_UserSaleDigits.Digit_数量, , True)
                        End If
                    End If

                    If CheckStrickUsable(mInt单据号, Me.cbo库房.ItemData(Me.cbo库房.ListIndex), Val(.TextMatrix(intRow, 1)), .TextMatrix(intRow, mintcol药名), _
                        IIf(mlng模块号 = 模块号.其他入库, 0, (.TextMatrix(intRow, mintcol批次))), Val(冲销数量_IN), mint库存检查入库库房, Trim(.TextMatrix(intRow, mIntColNO)), Val(.TextMatrix(intRow, mintcol序号))) = False Then
                        Exit Function
                    End If
                End If
            End If
        Next
        
        填制人_IN = UserInfo.用户姓名
        填制日期_IN = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")

        On Error GoTo errHandle
        行次_IN = 0
        '按药品ID顺序更新数据
        Call SetSortRecord
        mrecSort.Sort = "药品id,批次,序号"
        mrecSort.MoveFirst
        
        For i = 1 To mrecSort.RecordCount
            intRow = mrecSort!行号
            If .TextMatrix(intRow, 1) <> "" And Val(.TextMatrix(intRow, mintcol冲销数量)) <> 0 Then
                NO_IN = .TextMatrix(intRow, mIntColNO)
                If lastNO <> NO_IN Then
                    lastNO = NO_IN
                    行次_IN = 0
                End If
                
                行次_IN = 行次_IN + 1
                药品ID_IN = .TextMatrix(intRow, 1)
                str药品id = IIf(str药品id = "", "", str药品id & ",") & 药品ID_IN
                If Val(.TextMatrix(intRow, mintcol冲销数量)) = Val(.TextMatrix(intRow, mintcol数量)) Then
                    冲销数量_IN = Val(.TextMatrix(intRow, mintcol数量)) * Val(.TextMatrix(intRow, mintcol比例系数))
                Else
                    冲销数量_IN = zlStr.FormatEx(.TextMatrix(intRow, mintcol冲销数量) * .TextMatrix(intRow, mintcol比例系数), gtype_UserSaleDigits.Digit_数量, , True)
                End If
                
                冲销数量_IN = 冲销数量_IN
                原记录状态_IN = .TextMatrix(intRow, mintcol记录状态)
                摘要_IN = .TextMatrix(intRow, mintcol摘要)
                序号_IN = IIf(mlng模块号 <> 模块号.药品移库, .TextMatrix(intRow, mintcol序号), Val(.TextMatrix(intRow, mintcol序号)) - 1)
                
                If mlng模块号 = 模块号.外购入库 Then
                    发票号_IN = Trim(.TextMatrix(intRow, mintcol发票号))
                    发票代码_In = Trim(.TextMatrix(intRow, mintcol发票代码))
                    发票金额_IN = Val(IIf(.TextMatrix(intRow, mintcol发票金额) = "", "", .TextMatrix(intRow, mintcol发票金额)))
                End If
                
                Select Case mlng模块号
                    Case 模块号.其他出库
                        gstrSQL = "ZL_药品其他出库_STRIKE("
                    Case 模块号.其他入库
                        gstrSQL = "ZL_药品其他入库_STRIKE("
                    Case 模块号.外购入库
                        gstrSQL = "ZL_药品外购_STRIKE("
                    Case 模块号.药品领用
                        gstrSQL = "ZL_药品领用_STRIKE("
                    Case 模块号.药品移库
                        gstrSQL = "ZL_药品移库_STRIKE("
                End Select
                
                '行次
                gstrSQL = gstrSQL & 行次_IN
                '原记录状态
                gstrSQL = gstrSQL & "," & 原记录状态_IN
                'NO
                gstrSQL = gstrSQL & ",'" & NO_IN & "'"
                '序号
                gstrSQL = gstrSQL & "," & 序号_IN
                '药品ID
                gstrSQL = gstrSQL & "," & 药品ID_IN
                If mlng模块号 = 模块号.其他入库 Then
                    gstrSQL = gstrSQL & "," & IIf(摘要_IN = "", "Null", "'" & 摘要_IN & "'")
                End If
                '冲销数量
                gstrSQL = gstrSQL & "," & 冲销数量_IN
                '填制人
                gstrSQL = gstrSQL & ",'" & 填制人_IN & "'"
                '填制日期
                gstrSQL = gstrSQL & ",to_date('" & Format(填制日期_IN, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS')"
                
                If mlng模块号 = 模块号.外购入库 Then
                    '发票号
                    gstrSQL = gstrSQL & "," & IIf(发票号_IN = "", "Null", "'" & 发票号_IN & "'")
                    '发票金额
                    gstrSQL = gstrSQL & "," & 发票金额_IN
                    '是否全部冲销
                    gstrSQL = gstrSQL & "," & int全部冲销
                    '是否财务审核
                    gstrSQL = gstrSQL & "," & 0
                End If
                
                If mlng模块号 = 模块号.外购入库 Or mlng模块号 = 模块号.药品移库 Then
                    '摘要
                    gstrSQL = gstrSQL & "," & IIf(摘要_IN = "", "Null", "'" & 摘要_IN & "'")
                End If
                
                If mlng模块号 = 模块号.外购入库 Then
                    '发票代码
                    gstrSQL = gstrSQL & "," & IIf(发票代码_In = "", "Null", "'" & 发票代码_In & "'")
                End If
                
                If mlng模块号 = 模块号.药品移库 Then
                    '冲销方式
                    gstrSQL = gstrSQL & ",0"
                End If
                gstrSQL = gstrSQL & ")"

                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            End If
            mrecSort.MoveNext
        Next
        gcnOracle.BeginTrans
        For i = 0 To UBound(arrSql)
            Call zldatabase.ExecuteProcedure(CStr(arrSql(i)), "SaveCard")
        Next
        gcnOracle.CommitTrans
        
        If 行次_IN = 0 Then
            MsgBox "没有选择一行药品来冲销，请录入冲销数量！", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End With
    SaveStrike = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    'MsgBox "存盘失败！请检查！", vbInformation + vbOKOnly, gstrSysName
    Call SaveErrLog
End Function
Private Sub SetSortRecord()
'------------------------------------------------
'要冲销单据排序
'------------------------------------------------
    Dim n As Integer
    
    If Me.vsfList.rows < 2 Then Exit Sub
    If vsfList.TextMatrix(1, 1) = "" Then Exit Sub
    
    Set mrecSort = New ADODB.Recordset
    With mrecSort
        If .State = 1 Then .Close
        .Fields.Append "行号", adDouble, 18, adFldIsNullable
        .Fields.Append "序号", adDouble, 18, adFldIsNullable
        .Fields.Append "药品ID", adDouble, 18, adFldIsNullable
        .Fields.Append "批次", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        For n = 1 To vsfList.rows - 1
            If vsfList.TextMatrix(n, 1) <> "" Then
                .AddNew
                !行号 = n
                !序号 = IIf(Val(vsfList.TextMatrix(n, mintcol序号)) = 0, n, Val(vsfList.TextMatrix(n, mintcol序号)))
                !药品ID = Val(vsfList.TextMatrix(n, 0))
                !批次 = Val(vsfList.TextMatrix(n, mintcol批次))
                
                .Update
            End If
        Next
    End With
End Sub
Private Sub Txt审核人_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub
Private Sub Txt填制人_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub
Private Sub Txt药品_KeyDown(KeyCode As Integer, Shift As Integer)
'--------------------------------------------------------------------
'调用药品通用选择器，选择药品
'--------------------------------------------------------------------
    Dim vRect As RECT
    Dim strsql As String
    Dim sngLeft As Single
    Dim sngTop As Single
    
    If KeyCode = 13 Then
        sngLeft = Me.Left + Me.txt药品.Left + Screen.TwipsPerPixelX + 100
        sngTop = Me.Top + Me.Height - Me.ScaleHeight + Me.txt药品.Top + Me.pic基本信息.Top + Me.txt药品.Height + 400
        
        If grsMaster.State = adStateClosed Then
            Call SetSelectorRS(6, "药品外购入库管理", cbo库房.ItemData(cbo库房.ListIndex), cbo库房.ItemData(cbo库房.ListIndex))
        End If
        
'        Set mrsReturn = Frm药品多选选择器.ShowME(Me, 6, cbo库房.ItemData(cbo库房.ListIndex), cbo库房.ItemData(cbo库房.ListIndex), , Me.txt药品.Text, sngLeft, sngTop, True, True, False, False, True, 0)
        Set mrsReturn = frmSelector.showMe(Me, 1, 6, UCase(Me.txt药品.Text), sngLeft, sngTop, cbo库房.ItemData(cbo库房.ListIndex), cbo库房.ItemData(cbo库房.ListIndex), , 0, True, True, True, , False)
        If Not (mrsReturn Is Nothing) Then
            If Not mrsReturn.EOF Then
                Me.txt药品.Text = mrsReturn!通用名
                Me.txt药品.Tag = mrsReturn!药品ID
            Else
                Me.txt药品.SetFocus
                Me.txt药品.SelStart = 0
                Me.txt药品.SelLength = Len(Me.txt药品.Text)
            End If
        End If
    End If
End Sub
Private Sub vsfColSel_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '设置列的隐藏和显示
    If Me.vsfColSel.TextMatrix(Row, 0) <> 0 Then
        Me.vsfList.ColHidden(Me.vsfColSel.RowData(Row)) = False
    Else
        Me.vsfList.ColHidden(Me.vsfColSel.RowData(Row)) = True
    End If
End Sub
Private Sub vsfColSel_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim sum As Integer
    
    sum = 6
    If Me.vsfList.ColHidden(mintcol部门名称) = False Then sum = 7
    
    If Row > sum Then
        Cancel = False
    Else
        Cancel = True
    End If
End Sub

Private Sub vsfColSel_LostFocus()
    Me.vsfColSel.Visible = False
End Sub

Private Sub vsfList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '-------------------------------------------------------------
    '编辑冲销数量后调整价格
    '-------------------------------------------------------------
    Dim strKey As String
    Dim i As Integer
    Dim dblNum As Double
    Dim count As Integer
    
    If Col <> mintcol冲销数量 Then Exit Sub
    
    If vsfList.TextMatrix(Row, Col) = "" And strKey = "" Then
        vsfList.TextMatrix(Row, mintcol冲销数量) = 0
        vsfList.Cell(flexcpForeColor, Row, mintcol冲销数量) = CSTCOLOR_NOFONT
        vsfList.Cell(flexcpBackColor, Row, 1, Row, MINTCOL总列数 - 1) = CSTCOLOR_NOMODIFY
        Exit Sub
    End If
    
    If Not IsNumeric(strKey) And strKey <> "" Then
        MsgBox "对不起，冲销数量必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
        vsfList.TextMatrix(Row, mintcol冲销数量) = 0
        vsfList.Cell(flexcpForeColor, Row, mintcol冲销数量) = CSTCOLOR_NOFONT
        vsfList.Cell(flexcpBackColor, Row, 1, Row, MINTCOL总列数 - 1) = CSTCOLOR_NOMODIFY
        Exit Sub
    End If
    
    If CDbl(Me.vsfList.TextMatrix(Row, mintcol冲销数量)) > CDbl(Me.vsfList.TextMatrix(Row, mintcol数量)) Then
        Me.vsfList.TextMatrix(Row, mintcol冲销数量) = Me.vsfList.TextMatrix(Row, mintcol数量)
    End If
    
    For i = 1 To Me.vsfList.rows - 1
        If zlStr.Nvl(Me.vsfList.TextMatrix(i, mintcol冲销数量)) <> 0 Then
            count = 1
            Me.combars.Item(1).Controls.Item(MINTBTNWRITEOFF).Enabled = True
            Exit For
        End If
    Next
    
    If count <> 1 Then
        Me.combars.Item(1).Controls.Item(MINTBTNWRITEOFF).Enabled = False
    End If
    
    strKey = vsfList.TextMatrix(Row, mintcol冲销数量)
    If Val(strKey) <> 0 Then
        If InStr(1, strKey, ".") <> 0 Then vsfList.TextMatrix(Row, mintcol冲销数量) = zlStr.FormatEx(vsfList.TextMatrix(Row, mintcol冲销数量), mintNumberDigit, , True)
    Else
        vsfList.TextMatrix(Row, mintcol冲销数量) = 0
    End If
    If Me.vsfList.TextMatrix(Row, Col) <> 0 Then
        With Me.vsfList
            If .TextMatrix(Row, mintcol售价) <> "" Then
                .TextMatrix(Row, mintcol售价金额) = zlStr.FormatEx(.TextMatrix(Row, mintcol售价) * strKey, mintMoneyDigit, , True)
            End If

            .TextMatrix(Row, mintcol成本金额) = zlStr.FormatEx(Val(.TextMatrix(Row, mintcol成本价)) * strKey, mintMoneyDigit, , True)
            .TextMatrix(Row, mintcol差价) = zlStr.FormatEx(Val(.TextMatrix(Row, mintcol售价金额)) - Val(.TextMatrix(Row, mintcol成本金额)), mintMoneyDigit, , True)
            
            If mlng模块号 = 模块号.外购入库 And Val(.TextMatrix(.Row, mintcol是否变价)) = 1 And zlStr.Nvl(.TextMatrix(.Row, mintcol批次), 0) <> 0 Then
                 .TextMatrix(.Row, mintcol零售金额) = zlStr.FormatEx(Val(.TextMatrix(.Row, mintcol零售价)) * Val(strKey), mintMoneyDigit, , True)
                 .TextMatrix(.Row, mintcol零售差价) = zlStr.FormatEx(Val(.TextMatrix(.Row, mintcol零售金额)) - Val(.TextMatrix(.Row, mintcol成本金额)), mintMoneyDigit, , True)
            End If
            
            .Cell(flexcpForeColor, Row, mintcol冲销数量) = CSTCOLOR_FONT
            .Cell(flexcpBackColor, Row, 1, Row, MINTCOL总列数 - 1) = CSTCOLOR_MODIFY
        End With
    Else
        With Me.vsfList
            .TextMatrix(Row, mintcol售价金额) = 0
            .TextMatrix(Row, mintcol差价) = 0
            .TextMatrix(Row, mintcol成本金额) = 0
            
            If mlng模块号 = 模块号.外购入库 And Val(.TextMatrix(Row, mintcol是否变价)) = 1 And zlStr.Nvl(.TextMatrix(Row, mintcol批次), 0) <> 0 Then
                 .TextMatrix(Row, mintcol零售金额) = 0
                 .TextMatrix(Row, mintcol零售差价) = 0
            End If
            
            .Cell(flexcpForeColor, Row, mintcol冲销数量) = CSTCOLOR_NOFONT
            .Cell(flexcpBackColor, Row, 1, Row, MINTCOL总列数 - 1) = CSTCOLOR_NOMODIFY
        End With
    End If
    If mlng模块号 = 模块号.外购入库 Or mlng模块号 = 模块号.其他入库 Or mlng模块号 = 模块号.药品移库 Then
        If IIf(vsfList.TextMatrix(Row, mintcol批次) = "", 0, vsfList.TextMatrix(Row, mintcol批次)) = 0 Then
            dblNum = Val(vsfList.TextMatrix(Row, mintcol可用数量)) - Val(vsfList.TextMatrix(Row, mintcol冲销数量))
            For i = 1 To Me.vsfList.rows - 1
                vsfList.TextMatrix(i, mintcol可用数量) = zlStr.FormatEx(dblNum, mintNumberDigit, , True)
            Next
        End If
    End If
    
End Sub
Private Sub vsfList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '----------------------------------------------------------------------
    '控制可以编辑的列
    '只有冲销数量列可以编辑
    '----------------------------------------------------------------------
    If Col = mintcol冲销数量 Or Col = mintcol摘要 Or Row = 0 Then
        Cancel = False
    Else
        Cancel = True
        Exit Sub
    End If
End Sub
Private Sub vsfList_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    If Col = mintcol选择 Or Col = mIntColNO Or Col = mintcol药品名称与编码 Or Position = mintcol药品名称与编码 Or Position = mIntColNO Or Position = mintcol选择 Then
        Position = Col
    End If
End Sub
Private Sub vsfList_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = mintcol选择 Or Col = mIntColNO Or Col = mintcol药品名称与编码 Then
        Cancel = True
    Else
        Cancel = False
    End If
End Sub
Private Sub vsfList_DblClick()
'--------------------------------------------
'将数量的值付给冲销数量
'--------------------------------------------
    Dim strKey As String
    Dim dblNum As Double
    Dim i As Integer
    Dim count As Integer

    If vsfList.Row = 0 Or vsfList.Col <> mintcol数量 Then Exit Sub
    
    If mlng模块号 = 模块号.外购入库 Or mlng模块号 = 模块号.其他入库 Or mlng模块号 = 模块号.药品移库 Then
        If Val(Me.vsfList.TextMatrix(vsfList.Row, mintcol批次)) = 0 Then
            Me.vsfList.TextMatrix(vsfList.Row, mintcol可用数量) = zlStr.FormatEx(Val(Me.vsfList.TextMatrix(vsfList.Row, mintcol可用数量)) + Val(Me.vsfList.TextMatrix(vsfList.Row, mintcol冲销数量)), mintNumberDigit, , True)
        End If
    End If
    
    If zlStr.Nvl(Me.vsfList.TextMatrix(Me.vsfList.Row, mintcol冲销数量), 0) = 0 Then
        Me.vsfList.TextMatrix(Me.vsfList.Row, mintcol冲销数量) = zlStr.FormatEx(Me.vsfList.TextMatrix(Me.vsfList.Row, mintcol数量), mintNumberDigit, , True)
    Else
        Me.vsfList.TextMatrix(Me.vsfList.Row, mintcol冲销数量) = 0
    End If
    
    For i = 1 To Me.vsfList.rows - 1
        If Me.vsfList.TextMatrix(i, mintcol冲销数量) <> 0 Then
            count = 1
            Me.combars.Item(1).Controls.Item(MINTBTNWRITEOFF).Enabled = True
            Exit For
        End If
    Next
    
    If count <> 1 Then
        Me.combars.Item(1).Controls.Item(MINTBTNWRITEOFF).Enabled = False
    End If
    
    If Me.vsfList.TextMatrix(Me.vsfList.Row, mintcol冲销数量) <> 0 Then
        With Me.vsfList
            strKey = .TextMatrix(.Row, mintcol冲销数量)
            If .TextMatrix(.Row, mintcol售价) <> "" Then
                .TextMatrix(.Row, mintcol售价金额) = zlStr.FormatEx(.TextMatrix(.Row, mintcol售价) * strKey, mintMoneyDigit, , True)
            End If
            
'            .TextMatrix(.Row, mintcol成本价) =Str.FormatEx((Val(.TextMatrix(.Row, mintcol售价金额)) - Val(.TextMatrix(.Row, mintcol差价))) / strkey, mintCostDigit)
            .TextMatrix(.Row, mintcol成本金额) = zlStr.FormatEx(Val(.TextMatrix(.Row, mintcol成本价)) * strKey, mintMoneyDigit, , True)
            .TextMatrix(.Row, mintcol差价) = zlStr.FormatEx(Val(.TextMatrix(.Row, mintcol售价金额)) - Val(.TextMatrix(.Row, mintcol成本金额)), mintMoneyDigit, , True)
            
            If mlng模块号 = 模块号.外购入库 And Val(.TextMatrix(.Row, mintcol是否变价)) = 1 And zlStr.Nvl(.TextMatrix(.Row, mintcol批次), 0) <> 0 Then
                 .TextMatrix(.Row, mintcol零售金额) = zlStr.FormatEx(Val(.TextMatrix(.Row, mintcol零售价)) * Val(strKey), mintMoneyDigit, , True)
                 .TextMatrix(.Row, mintcol零售差价) = zlStr.FormatEx(Val(.TextMatrix(.Row, mintcol零售金额)) - Val(.TextMatrix(.Row, mintcol成本金额)), mintMoneyDigit, , True)
            End If
            
            .Cell(flexcpForeColor, .Row, mintcol冲销数量) = CSTCOLOR_FONT
            .Cell(flexcpBackColor, .Row, 1, .Row, MINTCOL总列数 - 1) = CSTCOLOR_MODIFY
        End With
    Else
        With Me.vsfList
            .TextMatrix(.Row, mintcol售价金额) = 0
            .TextMatrix(.Row, mintcol差价) = 0
            .TextMatrix(.Row, mintcol成本金额) = 0
            
            If mlng模块号 = 模块号.外购入库 And Val(.TextMatrix(.Row, mintcol是否变价)) = 1 And zlStr.Nvl(.TextMatrix(.Row, mintcol批次), 0) <> 0 Then
                 .TextMatrix(.Row, mintcol零售金额) = 0
                 .TextMatrix(.Row, mintcol零售差价) = 0
            End If
            
            .Cell(flexcpForeColor, .Row, mintcol冲销数量) = CSTCOLOR_NOFONT
            .Cell(flexcpBackColor, .Row, 1, .Row, MINTCOL总列数 - 1) = CSTCOLOR_NOMODIFY
        End With
    End If
    
    If mlng模块号 = 模块号.外购入库 Or mlng模块号 = 模块号.其他入库 Or mlng模块号 = 模块号.药品移库 Then
        If IIf(vsfList.TextMatrix(vsfList.Row, mintcol批次) = "", 0, vsfList.TextMatrix(vsfList.Row, mintcol批次)) = 0 Then
            dblNum = Val(vsfList.TextMatrix(vsfList.Row, mintcol可用数量)) - Val(vsfList.TextMatrix(vsfList.Row, mintcol冲销数量))
            For i = 1 To Me.vsfList.rows - 1
                vsfList.TextMatrix(i, mintcol可用数量) = zlStr.FormatEx(dblNum, mintNumberDigit, , True)
            Next
        End If
    End If
End Sub
Private Sub vsfList_EnterCell()
    Dim i As Integer
    
    If Me.vsfList.Row = 0 Then Exit Sub
    Me.staThis.Panels(2).Text = "当前批次的可用库存为" & Me.vsfList.TextMatrix(Me.vsfList.Row, mintcol可用数量) & Me.vsfList.TextMatrix(Me.vsfList.Row, mintcol单位)
    
    With Me.vsfList
        .Cell(flexcpPicture, 1, 0, .rows - 1, 0) = Nothing
        .Cell(flexcpPicture, .Row, 0, .Row, 0) = Me.imgList.ListImages(1).Picture
            
        '加粗编辑列的边框
        If .MouseCol = mintcol冲销数量 Or .MouseCol = mintcol摘要 Then
            .BackColorSel = CSTCOLOR_ENTERCELL
            .FocusRect = flexFocusSolid
        Else
            .FocusRect = flexFocusLight
        End If
        
        If .MouseCol = mintcol摘要 Then
            i = .Row
            Do While i >= 1
                If .TextMatrix(i, mintcol摘要) <> "" Then
                    .TextMatrix(.Row, mintcol摘要) = .TextMatrix(i, mintcol摘要)
                    Exit Sub
                End If
                i = i - 1
            Loop
        End If
    End With
End Sub
Private Sub vsfList_GotFocus()
    Me.vsfList.BackColorSel = CSTCOLOR_ENTERCELL
    If Me.vsfList.MouseCol = mintcol冲销数量 Then Me.vsfList.FocusRect = flexFocusSolid
End Sub
Private Sub vsfList_KeyDown(KeyCode As Integer, Shift As Integer)
    '-----------------------------------------------------------
    '通过回车键进入下一行冲销数量的编辑,删除选中行
    '-----------------------------------------------------------
    Dim strText As String
    Dim count As Integer
    
    With Me.vsfList
        If .Row = 0 Then Exit Sub
        If KeyCode = 46 Then
            .RemoveItem (.Row)
            
            If .rows = 1 Then
                Me.combars.Item(1).Controls.Item(MINTBTNALLWRITEOFF).Enabled = False
                Me.combars.Item(1).Controls.Item(MINTBTNALLELIMINATE).Enabled = False
                Me.combars.Item(1).Controls.Item(MINTBTNWRITEOFF).Enabled = False
                Me.combars.Item(1).Controls.Item(MINTBTNDEL).Enabled = False
            End If
        End If
        
        If KeyCode = 13 And .Col = mintcol摘要 Then
            If .Row <> .rows - 1 Then
                .Row = .Row + 1
                .Col = mintcol冲销数量
            End If
        ElseIf KeyCode = 13 And .Col = mintcol冲销数量 Then
            If .ColHidden(mintcol摘要) Then
                If .Row <> .rows - 1 Then
                    .Row = .Row + 1
                    .Col = mintcol冲销数量
                End If
            Else
                .Col = mintcol摘要
            End If
        End If
        
    End With
End Sub
Private Sub vsfList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strText As String
    Dim count As Integer
    
    If Col <> mintcol冲销数量 Or Row = 0 Then Exit Sub
    
    If InStr(MCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
    
    If InStr("1234567890" + Chr(46) + Chr(8) + Chr(13) + IIf(Val(vsfList.TextMatrix(Row, mintcol数量)) > 0, "", Chr(45)), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    ElseIf KeyAscii = Asc(".") Then
        If InStr(vsfList.EditText, ".") <> 0 Then     '只能存在一个小数点
            KeyAscii = 0
        End If
    ElseIf KeyAscii = Asc("-") Then
        If InStr(vsfList.EditText, "-") <> 0 Then     '只能存在一个-
            KeyAscii = 0
        End If
    End If
    strText = ""
End Sub

Private Sub vsfList_LostFocus()
    Me.vsfList.BackColorSel = CSTCOLOR_LOSTFORCE
    Me.vsfList.FocusRect = flexFocusLight
End Sub

Private Sub vsfList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    '-----------------------------------------------------------------
    '当鼠标移上数量列，提示双击该列可改变冲销数量的值
    '------------------------------------------------------------
    '为标题列则结束过程
    If Me.vsfList.MouseRow <= 0 Then Exit Sub
    
    If Me.vsfList.MouseCol = mintcol数量 Then
        If Me.vsfList.TextMatrix(Me.vsfList.MouseRow, mintcol冲销数量) = 0 Then
            Me.vsfList.ToolTipText = "双击该列，该行冲销数量等于" & Me.vsfList.TextMatrix(Me.vsfList.MouseRow, mintcol数量)
        Else
            Me.vsfList.ToolTipText = "双击该列，该行冲销数量等于0"
        End If
    Else
        Me.vsfList.ToolTipText = ""
    End If
End Sub
Private Sub SetSimple(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '---------------------------------------------------------------------
    '设置简洁列的形式
    '---------------------------------------------------------------------
    Dim i As Integer
    
    For i = mintcol生产日期 To Me.vsfList.Cols - 1
        If vsfList.ColHidden(i) = False Then
            vsfList.ColData(i) = i
            vsfList.ColHidden(i) = True
        End If
    Next
    
    If Control.Checked = False Then
        Control.Checked = True
        Me.combars.Item(1).Controls.Item(MINTBTNCONPLETE).Checked = False
    End If
End Sub
Private Sub SetConplete(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '---------------------------------------------------------------------
    '设置完整列的形式
    '---------------------------------------------------------------------
    Dim i As Integer
    If Control.Checked = False Then
        For i = mintcol生产日期 To Me.vsfList.Cols - 1
            If vsfList.ColData(i) Then vsfList.ColHidden(i) = False
        Next
        
        Control.Checked = True
        Me.combars.Item(1).Controls.Item(MINTBTNSIMPLE).Checked = False
    End If
End Sub
Private Sub SetColValue(ByVal str列名 As String, ByVal intValue As Integer, ByVal intW As Integer)
'-----------------------------------------------------------------
'在启用个性化设置的前提下，将列的顺序和列的宽度恢复为之前的状态
'-----------------------------------------------------------------
    Select Case str列名
        Case "选择"
            mintcol选择 = intValue
        Case "药品id"
            mintcol药品id = intValue
        Case "行号"
            mintcol行号 = intValue
        Case "NO"
            mIntColNO = intValue
        Case "药品名称与编码"
            mintcol药品名称与编码 = intValue
        Case "商品名"
            mintcol商品名 = intValue
        Case "药品来源"
            mintcol药品来源 = intValue
        Case "基本药物"
            mintcol基本药物 = intValue
        Case "药价级别"
            mintcol药价级别 = intValue
        Case "规格"
            mintcol规格 = intValue
        Case "单位"
            mintcol单位 = intValue
        Case "数量"
            mintcol数量 = intValue
        Case "冲销数量"
            mintcol冲销数量 = intValue
        Case "产地"
            mintcol产地 = intValue
        Case "批号"
            mintcol批号 = intValue
        Case "生产日期"
            mintcol生产日期 = intValue
        Case "有效期至"
            mintcol有效期至 = intValue
        Case "其他外观"
            mintcol其他外观 = intValue
        Case "采购限价"
            mintcol采购限价 = intValue
        Case "采购价"
            mintcol采购价 = intValue
        Case "扣率"
            mintcol扣率 = intValue
        Case "成本价"
            mintcol成本价 = intValue
        Case "成本金额"
            mintcol成本金额 = intValue
        Case "加成率"
            mintcol加成率 = intValue
        Case "售价"
            mintcol售价 = intValue
        Case "售价金额"
            mintcol售价金额 = intValue
        Case "差价"
            mintcol差价 = intValue
        Case "零售价"
            mintcol零售价 = intValue
        Case "零售单位"
            mintcol零售单位 = intValue
        Case "零售金额"
            mintcol零售金额 = intValue
        Case "零售差价"
            mintcol零售差价 = intValue
        Case "外购批准文号"
            mintcol外购批准文号 = intValue
        Case "随货单号"
            mintcol随货单号 = intValue
        Case "发票号"
            mintcol发票号 = intValue
        Case "发票代码"
            mintcol发票代码 = intValue
        Case "发票信息"
            mintcol发票信息 = intValue
        Case "发票金额"
            mintcol发票金额 = intValue
        Case "真实数量"
            mintcol真实数量 = intValue
        Case "序号"
            mintcol序号 = intValue
        Case "比例系数"
            mintcol比例系数 = intValue
        Case "药名"
            mintcol药名 = intValue
        Case "批次"
            mintcol批次 = intValue
        Case "记录状态"
            mintcol记录状态 = intValue
        Case "分批核算"
            mintcol分批核算 = intValue
        Case "可用数量"
            mintcol可用数量 = intValue
        Case "最大效期"
            mintcol最大效期 = intValue
        Case "实际差价"
            mintcol实际差价 = intValue
        Case "实际金额"
            mintcol实际金额 = intValue
        Case "上次供应商ID"
            mintcol上次供应商ID = intValue
        Case "摘要"
            mintcol摘要 = intValue
        Case "对方部门"
            mintcol对方部门 = intValue
        Case "是否变价"
            mintcol是否变价 = intValue
        Case "填制人"
            mintcol填制人 = intValue
        Case "填制日期"
            mintcol填制日期 = intValue
        Case "审核人"
            mintcol审核人 = intValue
        Case "审核日期"
            mintcol审核日期 = intValue
        Case "部门名称"
            mintcol部门名称 = intValue
    End Select
    
    vsfList.ColWidth(intValue) = intW
End Sub
Private Sub vsfList_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mlng模块号 = 模块号.外购入库 Or mlng模块号 = 模块号.其他入库 Or mlng模块号 = 模块号.药品移库 Then
        If Val(Me.vsfList.TextMatrix(Row, mintcol批次)) = 0 Then
            Me.vsfList.TextMatrix(Row, mintcol可用数量) = Val(Me.vsfList.TextMatrix(Row, mintcol可用数量)) + Val(Me.vsfList.TextMatrix(Row, mintcol冲销数量))
        End If
    End If
End Sub
Private Sub AllWriteOff()
    Dim i As Integer
    Dim dblOldSum As Double
    Dim dblSum As Double
    Dim strKey As String
    
    With Me.vsfList
        For i = 1 To .rows - 1
            dblOldSum = .TextMatrix(i, mintcol冲销数量) + dblOldSum
            .TextMatrix(i, mintcol冲销数量) = zlStr.FormatEx(.TextMatrix(i, mintcol数量), mintNumberDigit, , True)
            
            If i = .Row Then
                .EditText = .TextMatrix(i, mintcol数量)
            End If
            
            If mlng模块号 = 模块号.外购入库 Or mlng模块号 = 模块号.其他入库 Or mlng模块号 = 模块号.药品移库 Then
                If Val(.TextMatrix(i, mintcol批次)) = 0 Then
                    dblSum = dblSum + Val(.TextMatrix(i, mintcol冲销数量))
                    .Cell(flexcpText, 1, mintcol可用数量, .rows - 1, mintcol可用数量) = zlStr.FormatEx(.TextMatrix(1, mintcol可用数量) + dblOldSum - dblSum, mintNumberDigit, , True)
                End If
            End If
            
            If .TextMatrix(i, mintcol冲销数量) <> 0 Then
                strKey = .TextMatrix(i, mintcol冲销数量)
                If .TextMatrix(i, mintcol售价) <> "" Then
                    .TextMatrix(i, mintcol售价金额) = zlStr.FormatEx(.TextMatrix(i, mintcol售价) * strKey, mintMoneyDigit, , True)
                End If
                
                .TextMatrix(i, mintcol成本金额) = zlStr.FormatEx(Val(.TextMatrix(i, mintcol成本价)) * strKey, mintMoneyDigit, , True)
                .TextMatrix(i, mintcol差价) = zlStr.FormatEx(Val(.TextMatrix(i, mintcol售价金额)) - Val(.TextMatrix(i, mintcol成本金额)), mintMoneyDigit, , True)
                
                If mlng模块号 = 模块号.外购入库 And Val(.TextMatrix(i, mintcol是否变价)) = 1 And zlStr.Nvl(.TextMatrix(i, mintcol批次), 0) <> 0 Then
                     .TextMatrix(i, mintcol零售金额) = zlStr.FormatEx(Val(.TextMatrix(i, mintcol零售价)) * Val(.TextMatrix(i, mintcol冲销数量)), mintMoneyDigit, , True)
                     .TextMatrix(i, mintcol零售差价) = zlStr.FormatEx(Val(.TextMatrix(i, mintcol零售金额)) - Val(.TextMatrix(i, mintcol成本金额)), mintMoneyDigit, , True)
                End If
                
                .Cell(flexcpForeColor, i, mintcol冲销数量) = CSTCOLOR_FONT
                .Cell(flexcpBackColor, i, 1, i, MINTCOL总列数 - 1) = CSTCOLOR_MODIFY
            End If
        Next
    End With
    
    For i = 1 To Me.vsfList.rows - 1
        If Me.vsfList.TextMatrix(i, mintcol冲销数量) <> 0 Then
            Me.combars.Item(1).Controls.Item(MINTBTNWRITEOFF).Enabled = True
        End If
    Next
End Sub
Private Sub AllEliminate()
    Dim i As Integer
    Dim dblOldSum As Double
    Dim dblSum As Double
    
    Me.combars.Item(1).Controls.Item(MINTBTNWRITEOFF).Enabled = False
        
    With Me.vsfList
        If .rows <= 1 Then Exit Sub
        For i = 1 To .rows - 1
            If mlng模块号 = 模块号.外购入库 Or mlng模块号 = 模块号.其他入库 Or mlng模块号 = 模块号.药品移库 Then
                If Val(.TextMatrix(i, mintcol批次)) = 0 Then
                    dblSum = dblSum + Val(.TextMatrix(i, mintcol冲销数量))
                End If
            End If
            
            .TextMatrix(i, mintcol冲销数量) = 0
            .TextMatrix(i, mintcol售价金额) = 0
            .TextMatrix(i, mintcol差价) = 0
            .TextMatrix(i, mintcol成本金额) = 0
            If i = .Row Then
                .EditText = 0
            End If
            
            If mlng模块号 = 模块号.外购入库 And Val(.TextMatrix(.Row, mintcol是否变价)) = 1 And zlStr.Nvl(.TextMatrix(.Row, mintcol批次), 0) <> 0 Then
                 .TextMatrix(i, mintcol零售金额) = 0
                 .TextMatrix(i, mintcol零售差价) = 0
            End If
            
            .Cell(flexcpForeColor, i, mintcol冲销数量) = CSTCOLOR_NOFONT
            .Cell(flexcpBackColor, i, 1, i, MINTCOL总列数 - 1) = CSTCOLOR_NOMODIFY
        Next
        
        If .rows > 1 And Val(.TextMatrix(1, mintcol批次)) = 0 Then .Cell(flexcpText, 1, mintcol可用数量, .rows - 1, mintcol可用数量) = zlStr.FormatEx(.TextMatrix(1, mintcol可用数量) + dblSum, mintNumberDigit, , True)
    End With
End Sub
Private Sub Filter()
'-----------------------
'过滤操作
'-----------------------
    Dim i As Integer
    
    '清除表格数据
    For i = 1 To Me.vsfList.rows - 1
        Me.vsfList.RemoveItem (1)
    Next
    
    Me.lbl药品基本信息.Caption = "药品信息"
    
    Call InitData

    '设置行高
    For i = 1 To Me.vsfList.rows - 1
        Me.vsfList.RowHeight(i) = 300
    Next
End Sub
Private Sub WriteOff()
'-----------------------
'冲销操作
'-----------------------
    Dim i As Integer
    
    With Me.vsfList
        If .rows = 1 Then
            Exit Sub
        End If
    End With

    Call SetSortRecord
    
    If mlng模块号 = 模块号.外购入库 Then
        If CheckPay = True Then Exit Sub
    End If
    
    If SaveStrike = True Then
        Call combars_Execute(Me.combars.Item(1).Controls.Item(MINTBTNFILTER))
        If Me.vsfList.rows = 1 Then
            '因为没有数据，所以操作数据的按钮不可用
            Me.combars.Item(1).Controls.Item(MINTBTNALLWRITEOFF).Enabled = False
            Me.combars.Item(1).Controls.Item(MINTBTNALLELIMINATE).Enabled = False
            Me.combars.Item(1).Controls.Item(MINTBTNDEL).Enabled = False
        End If
        Me.combars.Item(1).Controls.Item(MINTBTNWRITEOFF).Enabled = False
    End If
End Sub

Private Function CheckPay() As Boolean
    '检查是否存在已经或者部分付款的单据
    Dim rs As New ADODB.Recordset
    Dim n As Integer
    
    On Error GoTo errHandle
    For n = 1 To vsfList.rows - 1
        If vsfList.TextMatrix(n, mintcol序号) <> "" Then
            gstrSQL = "Select Nvl(Max(付款序号), 0) 付款序号 From 应付记录 " & _
                " where 收发id=(Select Id From 药品收发记录 Where 单据=1 And No=[1] And (Mod(记录状态,3)=0 Or 记录状态=1) " & _
                " And 序号=[2]) "
            Set rs = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[取付款序号]", vsfList.TextMatrix(n, mIntColNO), Val(vsfList.TextMatrix(n, mintcol序号)))

            If rs.EOF Then CheckPay = False: Exit Function

            If rs!付款序号 = 0 Then
                CheckPay = False
            Else
                CheckPay = True
                MsgBox "第" & n & "行药品已经付款或者部分付款，不能冲销！", vbInformation, gstrSysName
                vsfList.Row = n
                vsfList.Col = 2
                Exit Function
            End If
        End If
    Next
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

Private Sub DelRow()
'-----------------------
'删除行操作
'-----------------------
    Dim count As Integer
    Dim i As Integer
    
    With Me.vsfList
        If .Row = 0 Then Exit Sub
        .RemoveItem (.Row)
        
        For i = 1 To Me.vsfList.rows - 1
            If zlStr.Nvl(Me.vsfList.TextMatrix(i, mintcol冲销数量)) <> 0 Then
                count = 1
                Me.combars.Item(1).Controls.Item(MINTBTNWRITEOFF).Enabled = True
                Exit For
            End If
        Next
    
        If count <> 1 Then
            Me.combars.Item(1).Controls.Item(MINTBTNWRITEOFF).Enabled = False
        End If
        
        If .rows = 1 Then
            Me.combars.Item(1).Controls.Item(MINTBTNALLWRITEOFF).Enabled = False
            Me.combars.Item(1).Controls.Item(MINTBTNALLELIMINATE).Enabled = False
            Me.combars.Item(1).Controls.Item(MINTBTNWRITEOFF).Enabled = False
            Me.combars.Item(1).Controls.Item(MINTBTNDEL).Enabled = False
        End If
    End With
End Sub

Private Function 同批次冲销数量(ByVal lng批次 As Long) As Double
    '批量冲销指定了库房和药品
    '获取列表中相同批次的冲销数量和
    Dim dbl冲销数量 As Double
    Dim intRow As Integer
    
    For intRow = 1 To vsfList.rows - 1
        If lng批次 = Val(vsfList.TextMatrix(intRow, mintcol批次)) Then
            dbl冲销数量 = dbl冲销数量 + (Val(vsfList.TextMatrix(intRow, mintcol冲销数量)) * Val(vsfList.TextMatrix(intRow, mintcol比例系数)))
        End If
    Next
    
    同批次冲销数量 = dbl冲销数量
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

