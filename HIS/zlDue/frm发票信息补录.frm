VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frm发票信息补录 
   Caption         =   "发票信息补录"
   ClientHeight    =   11265
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19980
   Icon            =   "frm发票信息补录.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11265
   ScaleWidth      =   19980
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picColor 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   10440
      ScaleHeight     =   255
      ScaleWidth      =   3375
      TabIndex        =   37
      Top             =   9023
      Width           =   3375
      Begin VB.PictureBox pic预减 
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1820
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   41
         Top             =   0
         Width           =   260
      End
      Begin VB.PictureBox pic停用 
         BackColor       =   &H000000C0&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   40
         Top             =   0
         Width           =   260
      End
      Begin VB.PictureBox pic近效期 
         BackColor       =   &H00C00000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   910
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   39
         Top             =   0
         Width           =   260
      End
      Begin VB.PictureBox pic库存不足 
         BackColor       =   &H00C000C0&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2730
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   38
         Top             =   0
         Width           =   260
      End
      Begin VB.Label lblColor1 
         AutoSize        =   -1  'True
         Caption         =   "物资"
         Height          =   180
         Index           =   0
         Left            =   2100
         TabIndex        =   45
         Top             =   37
         Width           =   360
      End
      Begin VB.Label lblColor1 
         AutoSize        =   -1  'True
         Caption         =   "药品"
         Height          =   180
         Index           =   2
         Left            =   285
         TabIndex        =   44
         Top             =   37
         Width           =   360
      End
      Begin VB.Label lblColor2 
         AutoSize        =   -1  'True
         Caption         =   "卫材"
         Height          =   180
         Left            =   1200
         TabIndex        =   43
         Top             =   37
         Width           =   360
      End
      Begin VB.Label lblColor1 
         AutoSize        =   -1  'True
         Caption         =   "设备"
         Height          =   180
         Index           =   4
         Left            =   3015
         TabIndex        =   42
         Top             =   37
         Width           =   360
      End
   End
   Begin VB.PictureBox pic其他信息 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   0
      ScaleHeight     =   3255
      ScaleWidth      =   3375
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4080
      Width           =   3375
      Begin VB.CommandButton cmd药品 
         Caption         =   "…"
         Enabled         =   0   'False
         Height          =   300
         Left            =   2820
         TabIndex        =   18
         Top             =   1035
         Width           =   255
      End
      Begin VB.CheckBox chkDept 
         BackColor       =   &H8000000B&
         Caption         =   "卫材(&W)"
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   2040
         TabIndex        =   14
         Tag             =   "4"
         Top             =   240
         Width           =   1035
      End
      Begin VB.CheckBox chkDept 
         BackColor       =   &H8000000B&
         Caption         =   "药品(&D)"
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   13
         Tag             =   "1"
         Top             =   240
         Width           =   1035
      End
      Begin VB.CheckBox chkDept 
         BackColor       =   &H8000000B&
         Caption         =   "物资(&M)"
         Enabled         =   0   'False
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   15
         Tag             =   "2"
         Top             =   600
         Width           =   1035
      End
      Begin VB.CheckBox chkDept 
         BackColor       =   &H8000000B&
         Caption         =   "设备(&S)"
         Enabled         =   0   'False
         Height          =   195
         Index           =   3
         Left            =   2040
         TabIndex        =   16
         Tag             =   "4"
         Top             =   600
         Width           =   1035
      End
      Begin VB.ComboBox cbo填制日期 
         Height          =   300
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1470
         Width           =   1725
      End
      Begin MSComCtl2.DTPicker dtp填制结束时间 
         Height          =   315
         Left            =   1380
         TabIndex        =   25
         Top             =   2340
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   193003523
         CurrentDate     =   43522
      End
      Begin MSComCtl2.DTPicker dtp填制开始时间 
         Height          =   315
         Left            =   1380
         TabIndex        =   23
         Top             =   1905
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   193003523
         CurrentDate     =   43522
      End
      Begin VB.TextBox txt项目名称 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1380
         MaxLength       =   20
         TabIndex        =   19
         Top             =   1035
         Width           =   1725
      End
      Begin VB.Label lbl药品 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "品    名"
         Height          =   180
         Left            =   360
         TabIndex        =   17
         Top             =   1095
         Width           =   720
      End
      Begin VB.Label lbl填制日期 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "填制日期"
         Height          =   180
         Left            =   360
         TabIndex        =   20
         Top             =   1530
         Width           =   720
      End
      Begin VB.Label lbl填制开始日期 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "开始日期"
         Height          =   180
         Left            =   360
         TabIndex        =   22
         Top             =   1965
         Width           =   720
      End
      Begin VB.Label lbl填制结束日期 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "结束日期"
         Height          =   180
         Left            =   360
         TabIndex        =   24
         Top             =   2400
         Width           =   720
      End
   End
   Begin VB.PictureBox pic基本信息 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   0
      ScaleHeight     =   2295
      ScaleWidth      =   3375
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1440
      Width           =   3375
      Begin VB.ComboBox cbo审核日期 
         Height          =   300
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   570
         Width           =   1725
      End
      Begin VB.CommandButton Cmd供应商 
         Caption         =   "…"
         Height          =   300
         Left            =   2820
         TabIndex        =   5
         Top             =   120
         Width           =   255
      End
      Begin VB.TextBox txt供应商 
         Height          =   300
         Left            =   1380
         TabIndex        =   4
         Top             =   120
         Width           =   1485
      End
      Begin MSComCtl2.DTPicker dtp开始时间 
         Height          =   315
         Left            =   1380
         TabIndex        =   9
         Top             =   1050
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   193003523
         CurrentDate     =   40848
      End
      Begin MSComCtl2.DTPicker dtp结束时间 
         Height          =   315
         Left            =   1380
         TabIndex        =   11
         Top             =   1538
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   193003523
         CurrentDate     =   40848
      End
      Begin VB.Label lbl开始日期 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "开始日期"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lbl结束日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "结束日期"
         Height          =   180
         Left            =   360
         TabIndex        =   10
         Top             =   1605
         Width           =   720
      End
      Begin VB.Label lbl审核日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "审核日期"
         Height          =   180
         Left            =   360
         TabIndex        =   6
         Top             =   630
         Width           =   720
      End
      Begin VB.Label lbl供应商 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "供 应 商"
         Height          =   180
         Left            =   360
         TabIndex        =   3
         Top             =   180
         Width           =   720
      End
   End
   Begin VB.PictureBox picDetails 
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   3480
      ScaleHeight     =   5295
      ScaleWidth      =   13335
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   600
      Width           =   13335
      Begin VB.Frame fra单据信息 
         Height          =   855
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   13335
         Begin VB.TextBox txt发票代码 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3390
            TabIndex        =   31
            Top             =   300
            Width           =   1485
         End
         Begin VB.TextBox txt发票号 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   840
            TabIndex        =   29
            Top             =   300
            Width           =   1485
         End
         Begin MSComCtl2.DTPicker dtp发票日期 
            Height          =   315
            Left            =   5955
            TabIndex        =   33
            Top             =   300
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy/MM/dd"
            Format          =   193003523
            CurrentDate     =   40848
         End
         Begin VB.Label lbl强调 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H000000FF&
            Height          =   180
            Left            =   120
            TabIndex        =   34
            Top             =   360
            Width           =   90
         End
         Begin VB.Label lbl发票日期 
            AutoSize        =   -1  'True
            Caption         =   "发票日期"
            Height          =   180
            Left            =   5174
            TabIndex        =   32
            Top             =   360
            Width           =   720
         End
         Begin VB.Label lbl发票代码 
            AutoSize        =   -1  'True
            Caption         =   "发票代码"
            Height          =   180
            Left            =   2617
            TabIndex        =   30
            Top             =   360
            Width           =   720
         End
         Begin VB.Label lbl发票号 
            AutoSize        =   -1  'True
            Caption         =   "发票号"
            Height          =   180
            Left            =   240
            TabIndex        =   28
            Top             =   360
            Width           =   540
         End
         Begin VB.Label lbl提示 
            Caption         =   $"frm发票信息补录.frx":6852
            ForeColor       =   &H000000FF&
            Height          =   540
            Left            =   8520
            TabIndex        =   46
            Top             =   187
            Width           =   4695
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   2805
         Left            =   0
         TabIndex        =   36
         Top             =   1200
         Width           =   12060
         _cx             =   21272
         _cy             =   4948
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16769992
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   315
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frm发票信息补录.frx":68F3
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
         ExplorerBar     =   5
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
   Begin XtremeSuiteControls.TaskPanel tkpMain 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   600
      Width           =   3375
      _Version        =   589884
      _ExtentX        =   5953
      _ExtentY        =   1508
      _StockProps     =   64
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   35
      Top             =   10905
      Width           =   19980
      _ExtentX        =   35243
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   2356
            Picture         =   "frm发票信息补录.frx":6968
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   30154
            MinWidth        =   600
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
   Begin VSFlex8Ctl.VSFlexGrid mshSelect 
      Height          =   2535
      Left            =   3480
      TabIndex        =   26
      Top             =   6240
      Width           =   4695
      _cx             =   8281
      _cy             =   4471
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16769992
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
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
      ExplorerBar     =   1
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
      VirtualData     =   0   'False
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
   Begin XtremeCommandBars.ImageManager imgPicture 
      Left            =   1320
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frm发票信息补录.frx":71FC
   End
   Begin XtremeDockingPane.DockingPane dkpPanel 
      Left            =   720
      Top             =   120
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
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
Attribute VB_Name = "frm发票信息补录"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const menuToolSave As Integer = 101
Private Const menuToolGetData As Integer = 102
Private Const menuToolExit As Integer = 103
Private Const menuToolCheck As Integer = 104
Private Const menuToolCheckCancel As Integer = 105
Private Const menuToolHelp As Integer = 108

Private Enum mColumn
    选择 = 0
    No = 1
    库房
    序号
    记录状态
    项目id
    项目信息
    规格
    批号
    单位
    数量
    采购价
    采购金额
    发票号
    发票代码
    发票金额
    标识
    收发ID
    随货单号
    Count = 19
    
End Enum
'用于区分药品、卫材、物资和设备
Private Const glngColor卫材 As Long = &HC00000
Private Const glngColor药品 As Long = &HC0
Private Const glngColor物资 As Long = &H8000&
Private Const glngColor设备 As Long = &HC000C0
'用于可否编辑颜色设置
Private Const glngColorGray = &H80000004
Private Const glngColorWhite = &H80000005
Private Const gint药品Index = 0
Private Const gint卫材Index = 1
Private Const gint物资Index = 2
Private Const gint设备Index = 3
Private mfrmMain As Form
Private mstrPrivs As String
Private mstr供应商Type As String '药品、卫材、物资、设备
Private mstrSelectTag As String
Private Const mintShowPriceDigit = 5           '价格小数位数
Private Const mintShowAmountDigit = 5         '金额小数位数
Private mbln待更新 As Boolean '勾选的发票金额合计是否需要更新，修改分批金额后需要更新
Private mbln付款标志 As Boolean

Private Sub cbo审核日期_Click()
    Dim dateCurrentDate As Date
    
    If cbo审核日期.Text = "自定义日期" Then
        dtp开始时间.Enabled = True
        dtp结束时间.Enabled = True
        
    Else
        dtp开始时间.Enabled = False
        dtp结束时间.Enabled = False
    End If
    
    '根据选择改变时间
    dateCurrentDate = Sys.Currentdate
    Select Case cbo审核日期.ListIndex
        Case 0, 1
            dtp开始时间.Value = CDate(Format(dateCurrentDate, "yyyy-mm-dd") & " 00:00:00")
            dtp结束时间.Value = dateCurrentDate
        Case 2
            dtp开始时间.Value = CDate(Format(DateAdd("d", -7, dateCurrentDate), "yyyy-mm-dd") & " 00:00:00")
            dtp结束时间.Value = dateCurrentDate
        Case 3
            dtp开始时间.Value = CDate(Format(DateAdd("d", -30, dateCurrentDate), "yyyy-mm-dd") & " 00:00:00")
            dtp结束时间.Value = dateCurrentDate
        Case 4
            dtp开始时间.Value = CDate(Format(DateAdd("d", -90, dateCurrentDate), "yyyy-mm-dd") & " 00:00:00")
            dtp结束时间.Value = dateCurrentDate
    End Select
End Sub



Private Sub cbo填制日期_Click()
    Dim dateCurrentDate As Date
    
    If cbo填制日期.Text = "自定义日期" Then
        dtp填制开始时间.Enabled = True
        dtp填制结束时间.Enabled = True
        
    Else
        dtp填制开始时间.Enabled = False
        dtp填制结束时间.Enabled = False
    End If
    
    '根据选择改变时间
    dateCurrentDate = Sys.Currentdate
    Select Case cbo填制日期.ListIndex
        Case 0, 1
            dtp填制开始时间.Value = CDate(Format(dateCurrentDate, "yyyy-mm-dd") & " 00:00:00")
            dtp填制结束时间.Value = dateCurrentDate
        Case 2
            dtp填制开始时间.Value = CDate(Format(DateAdd("d", -7, dateCurrentDate), "yyyy-mm-dd") & " 00:00:00")
            dtp填制结束时间.Value = dateCurrentDate
        Case 3
            dtp填制开始时间.Value = CDate(Format(DateAdd("d", -30, dateCurrentDate), "yyyy-mm-dd") & " 00:00:00")
            dtp填制结束时间.Value = dateCurrentDate
        Case 4
            dtp填制开始时间.Value = CDate(Format(DateAdd("d", -90, dateCurrentDate), "yyyy-mm-dd") & " 00:00:00")
            dtp填制结束时间.Value = dateCurrentDate
    End Select
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case menuToolGetData '提取数据
            If ValidData = False Then Exit Sub
            GetData
        Case menuToolSave '保存数据
            If Not SaveCard Then Exit Sub '保存失败退出
            txt发票号.SetFocus: txt发票号.Text = "": txt发票代码.Text = ""
            '重新提取数据
            If ValidData = False Then Exit Sub
            GetData
        Case menuToolCheck  '全选
            cbsCheck
        Case menuToolCheckCancel   '全清
            cbsCheckCancel
        Case menuToolHelp '帮助
            Call ShowHelp(App.ProductName, Me.hwnd, Me.Name)
        Case menuToolExit '退出
            Unload Me
    End Select
    
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If mbln待更新 Then
        AmountSum
        mbln待更新 = False
    End If
End Sub

Private Sub chkDept_Click(Index As Integer)
    Dim intSum As Integer
    Dim i As Integer
    
    For i = chkDept.LBound To chkDept.UBound
        If chkDept(i).Value = 1 Then intSum = intSum + 1
    Next
    
    If intSum = 1 Then
        txt项目名称.Enabled = True
        cmd药品.Enabled = True
        txt项目名称.BackColor = glngColorWhite
    Else
        txt项目名称.Enabled = False
        cmd药品.Enabled = False
        txt项目名称.BackColor = glngColorGray
    End If
    
    txt项目名称.Text = ""
    txt项目名称.Tag = ""
End Sub

Private Sub cbsCheck()
    Dim i As Integer
    
    With vsfList
        For i = 1 To .Rows - 1
            .TextMatrix(i, mColumn.选择) = "√"
            .TextMatrix(i, mColumn.发票号) = txt发票号.Text
            .TextMatrix(i, mColumn.发票代码) = txt发票代码.Text
        Next
    End With
    
    AmountSum
End Sub

Private Sub cbsCheckCancel()
    Dim i As Integer
    
    With vsfList
        For i = 1 To .Rows - 1
            .TextMatrix(i, mColumn.选择) = ""
            .TextMatrix(i, mColumn.发票号) = ""
            .TextMatrix(i, mColumn.发票代码) = ""
        Next
    End With
    
    AmountSum
End Sub

Private Sub Cmd供应商_Click()
    Dim rsTemp As ADODB.Recordset
    Dim strTemp As String
    strTemp = frm供应商选择.SelDept(mstrPrivs)
    If strTemp = "" Then
        Unload frm供应商选择
        If txt供应商.Enabled Then txt供应商.SetFocus
        Exit Sub
    End If
    txt供应商.Text = Mid(strTemp, InStr(strTemp, ",") + 1)
    txt供应商.Tag = Val(Left(strTemp, InStr(strTemp, ",") - 1))
    On Error GoTo errHandle
    Set rsTemp = zlDatabase.OpenSQLRecord("select 类型 from 供应商 where id=[1] ", Caption & "-提取供应商类型", txt供应商.Tag)
    If Not rsTemp.EOF Then
        mstr供应商Type = Nvl(rsTemp!类型)
    End If
    rsTemp.Close
    Call SetClass
    
    zlCommFun.PressKey vbKeyTab
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmd药品_Click()
    Call GetItem("")
    zlCommFun.PressKey vbKeyTab
End Sub


Private Sub Form_Load()
    Call initComandbar  '初始化工具栏
    Call InitTask  '初始化面板
    Call initComboBox
    Call initColumn
    dtp发票日期.Value = Sys.Currentdate
    mbln付款标志 = Val(zlDatabase.GetPara("外购入库需要经过标记付款后才能进行付款管理", glngSys, 0)) = 1
    
    RestoreWinState Me, App.ProductName
    stbThis.Panels(2).Picture = picColor
End Sub


Private Sub initComandbar()
    '初始化工具栏
    Dim cbrControlMain As CommandBarControl
    Dim cbrToolBar As CommandBar
    
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    Me.cbsMain.VisualTheme = xtpThemeOffice2003

    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsMain.EnableCustomization False
    Me.cbsMain.Icons = Me.imgPicture.Icons
    
    '工具栏定义
    Set cbrToolBar = Me.cbsMain.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched Or xtpFlagFloating Or xtpFlagAlignAny
    
    With cbrToolBar.Controls    '
        Set cbrControlMain = .Add(xtpControlButton, menuToolGetData, "提取数据")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        Set cbrControlMain = .Add(xtpControlButton, menuToolSave, "确定")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        Set cbrControlMain = .Add(xtpControlButton, menuToolCheck, "全选")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, menuToolCheckCancel, "全清")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        Set cbrControlMain = .Add(xtpControlButton, menuToolHelp, "帮助")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, menuToolExit, "退出")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        
    End With
    '快键绑定
    With Me.cbsMain.KeyBindings
        .Add 0, VK_F5, menuToolGetData
        .Add 0, VK_ESCAPE, menuToolExit
    End With
    
    
    cbsMain.Item(1).Delete
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

Private Sub initComboBox()
    With cbo填制日期
        .Clear
        .AddItem ""
        .AddItem "今日"
        .AddItem "一星期内"
        .AddItem "一个月内"
        .AddItem "三个月内"
        .AddItem "自定义日期"
    End With
    
    With cbo审核日期
        .Clear
        .AddItem ""
        .AddItem "今日"
        .AddItem "一星期内"
        .AddItem "一个月内"
        .AddItem "三个月内"
        .AddItem "自定义日期"
        .ListIndex = 1
    End With
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Me.tkpMain.Move 0, 530, Me.tkpMain.Width, Me.ScaleHeight - stbThis.Height - 530
    Me.picDetails.Move tkpMain.Width, 530, Me.ScaleWidth - tkpMain.Width, tkpMain.Height
    fra单据信息.Move 0, 0, picDetails.Width, fra单据信息.Height
    lbl提示.Left = fra单据信息.Width - lbl提示.Width - 50
    
    vsfList.Move 0, fra单据信息.Height, picDetails.Width, picDetails.Height - fra单据信息.Height
    
    With picColor
        .Top = Me.ScaleHeight - .Height - 30
        .Left = Me.ScaleWidth - stbThis.Panels(3).Width - stbThis.Panels(4).Width - .Width - 400
    End With
End Sub


Public Sub ShowCard(frmMain As Form, ByVal strPrivs As String)
    
    mstrPrivs = strPrivs
    
    Set mfrmMain = frmMain

    Me.Show vbModal, frmMain
End Sub

Private Sub SetClass()
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo errHandle
    For i = 0 To chkDept.Count - 1
        '系统
        If i >= 2 Then
            Set rsTemp = zlDatabase.OpenSQLRecord("Select Count(1) Rec From zlSystems Where 编号 = [1]", Caption, IIf(i = 2, 400, 600))
            chkDept(i).Enabled = rsTemp!rec > 0
            rsTemp.Close
        Else
            chkDept(i).Enabled = True
        End If
        '权限
        Select Case i
            Case 0
                chkDept(i).Enabled = chkDept(i).Enabled And InStr(mstrPrivs, ";药品;") > 0
            Case 1
                chkDept(i).Enabled = chkDept(i).Enabled And InStr(mstrPrivs, ";卫材;") > 0
            Case 2
                chkDept(i).Enabled = chkDept(i).Enabled And InStr(mstrPrivs, ";物资;") > 0
            Case 3
                chkDept(i).Enabled = chkDept(i).Enabled And InStr(mstrPrivs, ";设备;") > 0
        End Select
    Next
    '供应商
    If Len(mstr供应商Type) >= 1 Then  '药品
        chkDept(gint药品Index).Enabled = chkDept(gint药品Index).Enabled And Mid(mstr供应商Type, 1, 1) = "1"
    Else
        chkDept(gint药品Index).Enabled = False
    End If
    If Len(mstr供应商Type) >= 5 Then  '卫材
        chkDept(gint卫材Index).Enabled = chkDept(gint卫材Index).Enabled And Mid(mstr供应商Type, 5, 1) = "1"
    Else
        chkDept(gint卫材Index).Enabled = False
    End If
    If Len(mstr供应商Type) >= 2 Then  '物资
        chkDept(gint物资Index).Enabled = chkDept(gint物资Index).Enabled And Mid(mstr供应商Type, 2, 1) = "1"
    Else
        chkDept(gint物资Index).Enabled = False
    End If
    If Len(mstr供应商Type) >= 3 Then  '设备
        chkDept(gint设备Index).Enabled = chkDept(gint设备Index).Enabled And Mid(mstr供应商Type, 3, 1) = "1"
    Else
        chkDept(gint设备Index).Enabled = False
    End If
    
    Exit Sub
    
errHandle:
    Call ErrCenter
End Sub

Private Sub GetData()
    Dim rsRecord As ADODB.Recordset
    Dim str单据 As String
    Dim strSQL As String
    Dim str物资主SQL As String
    Dim str设备主SQL As String
    Dim str药品卫材主SQL As String
    Dim str组合SQL As String
    
    On Error GoTo errHandle
    
    Me.MousePointer = vbHourglass
    
    If cbo填制日期.Text <> "" And cbo审核日期.Text <> "" Then '填制审核日期都不为空
        strSQL = " And ((x.填制日期 between [2] and [3] And x.审核日期 is Null) Or x.审核日期 between [4] and [5] ) "
    ElseIf cbo填制日期.Text <> "" Then
        strSQL = " And x.填制日期 between [2] and [3] And x.审核日期 is Null "
    ElseIf cbo审核日期.Text <> "" Then
        strSQL = " And x.审核日期 between [4] and [5] "
    End If
    
    '需要加载药品或卫材
    If chkDept(gint药品Index).Value = 1 Or chkDept(gint卫材Index).Value = 1 Or _
        (chkDept(gint药品Index).Value = 1 And chkDept(gint卫材Index).Value = 1 And chkDept(gint物资Index).Value = 1 And chkDept(gint设备Index).Value = 1) Or _
        (chkDept(gint药品Index).Value <> 1 And chkDept(gint卫材Index).Value <> 1 And chkDept(gint物资Index).Value <> 1 And chkDept(gint设备Index).Value <> 1) Then
        
        '需要查询哪些类型
        If chkDept(gint药品Index).Value = 1 Then str单据 = "1"
        If chkDept(gint卫材Index).Value = 1 Then str单据 = IIf(str单据 = "", "", str单据 & ",") & "15"
        If str单据 = "" Then str单据 = "1,15" '都未勾选当作是都勾选
        
        If txt项目名称.Text <> "" Then strSQL = strSQL & " And x.药品ID = [6]"
        '读取记录状态作用：判断是否被冲销过所以用Min(x.记录状态)的方法
        '取原始单据的收发ID：用Min(x.id)的方法
        str药品卫材主SQL = "Select Distinct a.No 入库单据号, a.序号, a.记录状态, a.药品id 项目ID, '[' || d.编码 || ']' || d.名称 As 项目信息, d.规格, a.批号, d.计算单位 As 单位," & vbNewLine & _
                        "                       a.填写数量 As 数量, a.成本价 * 1 As 采购价, a.成本金额 As 采购金额, e.名称 库房,decode(a.单据,1,1,15,5) 标识,a.收发ID,Null 随货单号" & vbNewLine & _
                        "       From (Select x.No, Min(x.记录状态) 记录状态, Sum(实际数量) As 填写数量, Sum(成本金额) As 成本金额, x.药品id, x.序号," & vbNewLine & _
                        "                     x.批号, x.成本价,  x.供药单位id, x.库房id,x.单据,Min(x.id) 收发ID " & vbNewLine & _
                        "              From 药品收发记录 X" & vbNewLine & _
                        "              Where Not Exists" & vbNewLine & _
                        "               (Select 1" & vbNewLine & _
                        "                     From 应付记录 Y" & vbNewLine & _
                        "                     Where x.Id = y.收发id And y.系统标识 In (1, 5) And y.记录性质 = 0 And y.发票号 Is Not Null)" & vbNewLine & _
                        "                     And 单据 in (" & str单据 & ") " & vbNewLine & _
                        "             " & strSQL & vbNewLine & _
                        "              Group By x.No, x.药品id, x.序号, x.批号, x.成本价, x.供药单位id, x.库房id,x.单据" & vbNewLine & _
                        "              Having Sum(实际数量) <> 0) A, 收费项目目录 D, 部门表 E, 供应商 F" & vbNewLine & _
                        "       Where a.药品id = d.Id And a.供药单位id  + 0 = f.Id And a.库房id = e.Id And (Substr(f.类型, 1, 1) = 1 or Substr(f.类型, 5, 1) = 1) and f.id = [1]"

    End If
    
    '需要加载物资:物资系统必须安装
    If chkDept(gint物资Index).Enabled = True And (chkDept(gint物资Index).Value = 1 Or (chkDept(gint药品Index).Value = 1 And chkDept(gint卫材Index).Value = 1 And chkDept(gint物资Index).Value = 1 And chkDept(gint设备Index).Value = 1) Or _
    (chkDept(gint药品Index).Value <> 1 And chkDept(gint卫材Index).Value <> 1 And chkDept(gint物资Index).Value <> 1 And chkDept(gint设备Index).Value <> 1)) Then
        
        If txt项目名称.Text <> "" Then strSQL = strSQL & " And x.物资id = [6]"
        '读取记录状态作用：判断是否被冲销过所以用Min(x.记录状态)的方法
        '取原始单据的收发ID：用Min(x.id)的方法
        str物资主SQL = "Select Distinct a.No 入库单据号, a.序号, a.记录状态, a.物资id 项目ID, '[' || d.编码 || ']' || d.名称 As 项目信息, d.规格, a.批号, d.散装单位 As 单位," & vbNewLine & _
                    "                       a.填写数量 As 数量, a.成本价 * 1 As 采购价, a.成本金额 As 采购金额, e.名称 库房,2 标识,a.收发ID,Null 随货单号" & vbNewLine & _
                    "       From (Select x.No, Min(x.记录状态) 记录状态, Sum(实际数量) As 填写数量, Sum(金额) As 成本金额, x.物资id, x.序号," & vbNewLine & _
                    "                     x.批号, x.单价 成本价,  x.供货单位id, x.库房id,x.单据,Min(x.id) 收发ID" & vbNewLine & _
                    "              From 物资收发记录 X " & vbNewLine & _
                    "              Where Not Exists" & vbNewLine & _
                    "               (Select 1" & vbNewLine & _
                    "                     From 应付记录 Y" & vbNewLine & _
                    "                     Where x.Id = y.收发id And y.系统标识 = 2 And y.记录性质 In (0, -1) And y.发票号 Is Not Null)" & vbNewLine & _
                    "                     And x.单据 = 1 " & vbNewLine & _
                    "             " & strSQL & vbNewLine & _
                    "              Group By x.No, x.物资id, x.序号, x.批号, x.单价, x.供货单位id, x.库房id,x.单据" & vbNewLine & _
                    "              Having Sum(实际数量) <> 0) A, 物资目录 D, 部门表 E, 供应商 F" & vbNewLine & _
                    "       Where a.物资id = d.Id And a.供货单位id + 0 = f.Id And a.库房id = e.Id And  Substr(f.类型, 2, 1) = 1 and f.id = [1]"

    End If
    
    '需要加载设备:设备系统必须安装
    If chkDept(gint设备Index).Enabled = True And (chkDept(gint设备Index).Value = 1 Or (chkDept(gint药品Index).Value = 1 And chkDept(gint卫材Index).Value = 1 And chkDept(gint物资Index).Value = 1 And chkDept(gint设备Index).Value = 1) Or _
    (chkDept(gint药品Index).Value <> 1 And chkDept(gint卫材Index).Value <> 1 And chkDept(gint物资Index).Value <> 1 And chkDept(gint设备Index).Value <> 1)) Then
        
        If txt项目名称.Text <> "" Then strSQL = strSQL & " And x.设备id = [6]"
        '读取记录状态作用：判断是否被冲销过所以用Min(x.记录状态)的方法
        '取原始单据的收发ID：用Min(x.id)的方法
        str设备主SQL = "Select Distinct a.No 入库单据号, a.序号, a.记录状态, a.设备id 项目id, '[' || d.编码 || ']' || d.名称 As 项目信息, d.规格, a.批号, d.单位 ," & vbNewLine & _
                    "                a.填写数量 As 数量, a.成本价 * 1 As 采购价, a.成本金额 As 采购金额, e.名称 库房, 3 标识, a.收发id,a.随货单号" & vbNewLine & _
                    "From (Select x.No, Min(x.记录状态) 记录状态, Sum(实际数量) As 填写数量, Sum(金额) As 成本金额, x.设备id, x.序号, Null 批号, x.单价 成本价, x.供货单位id, x.库房id," & vbNewLine & _
                    "              x.单据, Min(x.Id) 收发ID,x.随货单号" & vbNewLine & _
                    "       From 设备收发记录 X " & vbNewLine & _
                    "              Where Not Exists" & vbNewLine & _
                    "               (Select 1" & vbNewLine & _
                    "                     From 应付记录 Y" & vbNewLine & _
                    "                     Where x.Id = y.收发id And y.系统标识 = 3 And y.记录性质 In (0, -1) And y.发票号 Is Not Null)" & vbNewLine & _
                    "                     And x.单据 =1 " & vbNewLine & _
                    "             " & strSQL & vbNewLine & _
                    "       Group By x.No, x.设备id, x.序号, x.批次, x.单价, x.供货单位id, x.库房id, x.单据,x.随货单号" & vbNewLine & _
                    "       Having Sum(实际数量) <> 0) A, 设备目录 D, 部门表 E, 供应商 F" & vbNewLine & _
                    "Where a.设备id = d.Id And a.供货单位id + 0 = f.Id And a.库房id = e.Id And Substr(f.类型, 3, 1) = 1 and f.id = [1]"


    End If
    
    If str药品卫材主SQL <> "" And str物资主SQL <> "" And str设备主SQL <> "" Then
        str组合SQL = str药品卫材主SQL & vbNewLine & " Union All " & vbNewLine & str物资主SQL & vbNewLine & " Union All" & vbNewLine & str设备主SQL & vbNewLine
    ElseIf str药品卫材主SQL <> "" And str物资主SQL <> "" Then
        str组合SQL = str药品卫材主SQL & vbNewLine & " Union All " & vbNewLine & str物资主SQL & vbNewLine
    ElseIf str药品卫材主SQL <> "" And str设备主SQL <> "" Then
        str组合SQL = str药品卫材主SQL & vbNewLine & " Union All" & vbNewLine & str设备主SQL & vbNewLine
    ElseIf str物资主SQL <> "" And str设备主SQL <> "" Then
        str组合SQL = str物资主SQL & vbNewLine & " Union All" & vbNewLine & str设备主SQL & vbNewLine
    ElseIf str药品卫材主SQL <> "" Then
        str组合SQL = str药品卫材主SQL & vbNewLine
    ElseIf str物资主SQL <> "" Then
        str组合SQL = str物资主SQL & vbNewLine
    ElseIf str设备主SQL <> "" Then
        str组合SQL = str设备主SQL & vbNewLine
    End If
    
    
    gstrSQL = "Select * " & vbNewLine & _
            "   From ( " & vbNewLine & _
            "  " & str组合SQL & _
            ")" & vbNewLine & _
            "Order By 入库单据号,库房,序号 Asc"
    
    

    Set rsRecord = zlDatabase.OpenSQLRecord(gstrSQL, "提取数据", Val(txt供应商.Tag), CDate(Format(dtp填制开始时间, "yyyy-mm-dd") & " 00:00:00"), CDate(Format(dtp填制结束时间, "yyyy-mm-dd") & " 23:59:59"), CDate(Format(dtp开始时间, "yyyy-mm-dd") & " 00:00:00"), CDate(Format(dtp结束时间, "yyyy-mm-dd") & " 23:59:59"), Val(txt项目名称.Tag))
    
    SetColumn rsRecord
    Me.MousePointer = vbDefault
    
    stbThis.Panels(2).Text = ""
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function ValidData() As Boolean
    
    If Val(txt供应商.Tag) = 0 Then
        ShowMsgbox "供应商未选择，不能继续！"
        If txt供应商.Enabled Then txt供应商.SetFocus
        Exit Function
    End If
    
    If Trim(cbo审核日期.Text) = "" And Trim(cbo填制日期.Text) = "" Then
        ShowMsgbox "审核日期和填制日期都为空，不能继续！"
        If cbo审核日期.Enabled Then cbo审核日期.SetFocus
        Exit Function
    End If
    
    ValidData = True
End Function

Private Sub Form_Unload(Cancel As Integer)
    If mshSelect.Visible = True Then
        mshSelect.Visible = False
        Cancel = 1
        Exit Sub
    End If
    
    SaveWinState Me, App.ProductName
    '卸载窗体对象
    If Not mfrmMain Is Nothing Then
        Set mfrmMain = Nothing
    End If
End Sub

Private Sub mshSelect_DblClick()
    With mshSelect
        If .Row > 0 And .TextMatrix(.Row, 0) <> "" Then
            mshSelect_KeyPress 13
        End If
    End With
End Sub

Private Sub mshSelect_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    With mshSelect
        Select Case mstrSelectTag
            Case "Provide"
                If KeyAscii = vbKeyReturn Then
                    If .Row = 0 Then Exit Sub
                    txt供应商.Text = "[" & .TextMatrix(.Row, 1) & "]" & .TextMatrix(.Row, 2)
                    
                    mstr供应商Type = .TextMatrix(.Row, 4)
                    Call SetClass
                    
                    txt供应商.Tag = Val(.TextMatrix(.Row, 0))
                    zlCommFun.PressKey vbKeyTab
                ElseIf KeyAscii = 27 Then
                    If txt供应商.Enabled Then txt供应商.SetFocus
                End If
            Case Else
        End Select
        .Visible = False
    End With
End Sub

Private Sub mshSelect_LostFocus()
    mshSelect.Visible = False
End Sub

Private Sub txt发票代码_GotFocus()
    txt发票代码.SelStart = 0
    txt发票代码.SelLength = Len(txt发票代码.Text)
End Sub

Private Sub txt发票代码_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then dtp发票日期.SetFocus
End Sub

Private Sub txt发票代码_LostFocus()
    Dim i As Integer
    
    With vsfList
        For i = 1 To .Rows - 1
            If .TextMatrix(i, mColumn.选择) = "√" Then .TextMatrix(i, mColumn.发票代码) = txt发票代码.Text
        Next
    End With
End Sub

Private Sub txt发票号_GotFocus()
    txt发票号.SelStart = 0
    txt发票号.SelLength = Len(txt发票号.Text)
End Sub

Private Sub txt发票号_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txt发票代码.SetFocus
End Sub

Private Sub txt发票号_LostFocus()
    Dim i As Integer
    
    With vsfList
        For i = 1 To .Rows - 1
            If .TextMatrix(i, mColumn.选择) = "√" Then .TextMatrix(i, mColumn.发票号) = txt发票号.Text
        Next
    End With
    
End Sub

Private Sub txt供应商_GotFocus()
    txt供应商.SelStart = 0
    txt供应商.SelLength = Len(txt供应商.Text)
End Sub

Private Sub txt供应商_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If SelMltProvide = False And mshSelect.Visible = False Then
            If txt供应商.Enabled Then txt供应商.SetFocus: txt供应商.SelStart = 0: txt供应商.SelLength = Len(txt供应商.Text)
        Else
            If mshSelect.Visible = False Then
                zlCommFun.PressKey vbKeyTab
            End If
        End If
    End If
End Sub

Private Function SelMltProvide() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取供应商数据
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim strTmp As String
    Dim str权限 As String
    
    If Trim(txt供应商.Text) = "" Then Exit Function
    
    strTmp = GetMatchingSting(UCase(txt供应商.Text), False)
    
    str权限 = " and " & Get分类权限(mstrPrivs)
    
    SelMltProvide = False
    
    strSQL = "" & _
        "  Select   ID,编码,名称,简码,类型" & _
        "  From  供应商 " & _
        "  Where (撤档时间 is null or To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01') " & _
        "       " & zl_获取站点限制 & "  and 末级=1  " & _
        "       And ( 编码 Like [1] or 名称 like [1] or 简码  like upper([1])) " & str权限
    On Error GoTo errHandle
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strTmp)
    
    If rsTemp.EOF Then
        ShowMsgbox "未找到指定的供应商!"
        Exit Function
    End If
    With rsTemp
        If .RecordCount > 1 Then
            mstrSelectTag = "Provide"
            Set mshSelect.DataSource = rsTemp
            With mshSelect
                .Top = tkpMain.Top + pic基本信息.Top + txt供应商.Top + txt供应商.Height + 10
                .Left = pic基本信息.Left + txt供应商.Left
                .Visible = True
                .ColWidth(0) = 0
                .ColWidth(1) = 800
                .ColWidth(2) = 2000
                .ColWidth(3) = 800
                .Row = 1
                .Col = 0
                .ColSel = .Cols - 1
                .ZOrder
                .SetFocus
                Exit Function
            End With
        Else
            txt供应商.Text = "[" & Nvl(rsTemp!编码) & "]" & rsTemp!名称
            txt供应商.Tag = Nvl(rsTemp!ID, 0)
            mstr供应商Type = rsTemp!类型
            Call SetClass
            SelMltProvide = True
        End If
    End With
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then Resume
End Function

Private Sub GetItem(ByVal strkey As String)
    Dim intClass As Integer
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim blnCancel As Boolean
    Dim vRect As RECT
    Dim sngX As Single, sngY As Single, sngH As Single
    Dim intSysParam As Integer
    Dim strMatch As String
    
    intClass = GetClassValue()
    vRect = zlControl.GetControlRect(txt项目名称.hwnd)
    sngX = vRect.Left
    sngY = vRect.Bottom
    
    On Error GoTo errHandle
    Select Case intClass
    Case 0
        '药品
        If strkey = "" Then
            strSQL = "Select ID, 上级id, 编码, 名称, '' 规格, '' 产地, '' 药库单位, '' 住院单位, '' 门诊单位, 0 As 末级 " & _
                     "From 诊疗分类目录 " & vbLf & _
                     "Where 类型 in ('1','2','3') " & vbLf & _
                     "Start With 上级id Is Null Connect By Prior ID = 上级id " & vbLf & _
                     "Union all " & vbLf & _
                     "Select a.Id, c.分类id As 上级id, a.编码, a.名称, a.规格, a.产地, b.药库单位, b.住院单位, b.门诊单位, 1 As 末级 " & vbLf & _
                     "From 收费项目目录 A, 药品规格 B, 诊疗项目目录 C " & vbLf & _
                     "Where a.Id = b.药品id And b.药名id = c.Id And a.类别 in ('5','6','7') " & vbLf & _
                     "  And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) "
            
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, Caption & "-药品" _
                    , False, "", "选择", False, False, False, sngX, sngY, sngH, blnCancel, False, False)
        Else
            strSQL = "Select Distinct a.ID, null 上级ID, a.编码, a.名称, a.规格, a.产地, b.药库单位, b.住院单位, b.门诊单位 " & vbLf & _
                     "From 收费项目目录 A, 药品规格 B, 收费项目别名 C " & vbLf & _
                     "Where a.Id = b.药品id And a.id = c.收费细目id And A.类别 in ('5','6','7') " & vbLf & _
                     "  And (to_char(A.撤档时间, 'yyyy-mm-dd') = '3000-01-01' or A.撤档时间 is null) " & _
                     "  And C.性质 = 1 "
            intSysParam = Val(zlDatabase.GetPara("简码方式"))
            strMatch = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = "0", "%", "")
            
            If IsNumeric(strkey) Then
                strSQL = strSQL & " And (a.编码 Like [1] Or C.简码 Like [2] And C.码类=3) "
            ElseIf zlCommFun.IsCharAlpha(strkey) Then
                strSQL = strSQL & " And C.简码 Like [2] and c.码类=" & IIf(intSysParam = 0, 1, 2) & " "
            ElseIf zlCommFun.IsCharChinese(strkey) Then
                strSQL = strSQL & " And C.名称 Like [2] "
            Else
                strSQL = strSQL & " And (a.编码 = [1] And C.名称 Like [2] Or C.简码 LIKE [2]) and c.码类=" & IIf(intSysParam = 0, 1, 2) & " "
            End If
            strSQL = strSQL & vbNewLine & "Order by a.编码 "
            
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, Caption & "-药品" _
                    , False, "", "选择", False, False, True, sngX, sngY, sngH, blnCancel, False, False _
                    , strkey & "%" _
                    , strMatch & strkey & "%")
        End If
        
    Case 1
        '卫材
        If strkey = "" Then
            strSQL = "Select ID, 上级id, 编码, 名称, '' 规格, '' 产地, '' As 计算单位, 0 As 末级 " & _
                     "From 诊疗分类目录 " & vbLf & _
                     "Where 类型 = '7' " & vbLf & _
                     "Start With 上级id Is Null Connect By Prior ID = 上级id " & vbLf & _
                     "Union all " & vbLf & _
                     "Select i.Id, b.分类id As 上级id, i.编码, i.名称, i.规格, i.产地, i.计算单位, 1 As 末级 " & vbLf & _
                     "From 收费项目目录 I, 材料特性 T, 诊疗项目目录 B " & vbLf & _
                     "Where i.Id = t.材料id And t.诊疗id = b.Id And i.类别 = '4' " & vbLf & _
                     "  And (i.撤档时间 Is Null Or i.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) "
            
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, Caption & "-卫材" _
                    , False, "", "选择", False, False, False, sngX, sngY, sngH, blnCancel, False, False)
        Else
            strSQL = "Select Distinct i.Id, i.编码, i.名称, i.规格, i.产地, i.计算单位, 1 As 末级 " & vbLf & _
                     "From 收费项目目录 I, 材料特性 T, 收费项目别名 B " & vbLf & _
                     "Where i.Id = t.材料id And i.Id = b.收费细目id And i.类别 = '4' " & vbLf & _
                     "  And (i.撤档时间 Is Null Or i.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) "
            intSysParam = Val(zlDatabase.GetPara("简码方式"))
            strMatch = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = "0", "%", "")
            
            If IsNumeric(strkey) Then
                strSQL = strSQL & " And (i.编码 Like [1] Or b.简码 Like [2] And b.码类=3) "
            ElseIf zlCommFun.IsCharAlpha(strkey) Then
                strSQL = strSQL & " And b.简码 Like [2] And b.码类 = [3] "
            ElseIf zlCommFun.IsCharChinese(strkey) Then
                strSQL = strSQL & " And b.名称 Like [2] "
            Else
                strSQL = strSQL & " And (i.编码 = [1] And b.名称 Like [2] Or b.简码 LIKE [2]) And b.码类 = [3] "
            End If
            strSQL = strSQL & vbLf & "Order by i.编码 "
        
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, Caption & "-卫材" _
                    , False, "", "选择", False, False, True, sngX, sngY, sngH, blnCancel, False, False _
                    , strkey & "%" _
                    , strMatch & strkey & "%" _
                    , IIf(intSysParam = 0, 1, 2))
        End If
    Case 2
        '物资
        If strkey = "" Then
            strSQL = "Select ID, 0 末级, 上级id, 编码, 名称, '' 规格, '' 产地, '' 散装单位, '' 包装单位 " & _
                     "From 物资分类 " & _
                     "Where 物资类别 in ('普通物资', '医用物资') " & _
                     "Start With 上级id Is Null Connect By Prior ID = 上级id " & _
                     "Union All " & _
                     "Select ID, 1 末级, 分类id 上级id, 编码, 名称, 规格, 产地, 散装单位, 包装单位 " & _
                     "From 物资目录 " & _
                     "Where (to_char(撤档时间,'yyyy-MM-DD') = '3000-01-01' or 撤档时间 is null) And 物资类别 in ('普通物资', '医用物资') "
            
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, Caption & "-物资" _
                    , False, "", "选择", False, False, False, sngX, sngY, sngH, blnCancel, False, False)
        Else
            strSQL = "Select ID, 编码, 名称, 规格, 产地, 散装单位, 包装单位 " & _
                     "From 物资目录 " & _
                     "Where (to_char(撤档时间,'yyyy-MM-DD') = '3000-01-01' or 撤档时间 is null) And 物资类别 in ('普通物资', '医用物资') "
            
            If IsNumeric(strkey) Then
                strSQL = strSQL & " And (编码 Like [1] Or 简码 Like [2]) "
            ElseIf zlCommFun.IsCharAlpha(strkey) Then
                strSQL = strSQL & " And 简码 Like [2] "
            ElseIf zlCommFun.IsCharChinese(strkey) Then
                strSQL = strSQL & " And 名称 Like [2] "
            Else
                strSQL = strSQL & " And (编码 = [1] And 名称 Like [2] Or 简码 LIKE [2]) "
            End If
            strSQL = strSQL & vbLf & "Order by 编码 "
        
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, Caption & "-物资" _
                    , False, "", "选择", False, False, True, sngX, sngY, sngH, blnCancel, False, False _
                    , strkey & "%" _
                    , "%" & strkey & "%")
        End If
    Case 3
        '设备
        If strkey = "" Then
            strSQL = "Select ID, 0 末级, 上级id, 编码, 名称, '' 规格, '' 产地, '' 单位 " & _
                     "From 设备分类 " & _
                     "Start With 上级id Is Null Connect By Prior ID = 上级id " & _
                     "Union All " & _
                     "Select ID, 1 末级, 分类id 上级id, 编码, 名称, 规格, 产地, 单位 " & _
                     "From 设备目录 " & _
                     "Where (to_char(撤档时间,'yyyy-MM-DD') = '3000-01-01' or 撤档时间 is null) "
            
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, Caption & "-设备" _
                    , False, "", "选择", False, False, False, sngX, sngY, sngH, blnCancel, False, False)
        Else
            strSQL = "Select ID, 编码, 名称, 规格, 产地, 单位 " & _
                     "From 设备目录 " & _
                     "Where (to_char(撤档时间,'yyyy-MM-DD') = '3000-01-01' or 撤档时间 is null) "
            
            If IsNumeric(strkey) Then
                strSQL = strSQL & " And (编码 Like [1] Or 简码 Like [2]) "
            ElseIf zlCommFun.IsCharAlpha(strkey) Then
                strSQL = strSQL & " And 简码 Like [2] "
            ElseIf zlCommFun.IsCharChinese(strkey) Then
                strSQL = strSQL & " And 名称 Like [2] "
            Else
                strSQL = strSQL & " And (编码 = [1] And 名称 Like [2] Or 简码 LIKE [2]) "
            End If
            strSQL = strSQL & vbLf & "Order by 编码 "
        
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, Caption & "-设备" _
                    , False, "", "选择", False, False, True, sngX, sngY, sngH, blnCancel, False, False _
                    , strkey & "%" _
                    , "%" & strkey & "%")
        End If
    End Select
    
    If blnCancel = False And Not rsTemp Is Nothing Then
        txt项目名称.Text = Nvl(rsTemp!名称)
        txt项目名称.Tag = Nvl(rsTemp!ID)
    End If
    
    If Not rsTemp Is Nothing Then
        rsTemp.Close
    ElseIf rsTemp Is Nothing And Not blnCancel Then
        MsgBox "未找到该项目!", vbInformation + vbDefaultButton1, gstrSysName
        txt项目名称.SetFocus
        txt项目名称.SelStart = 0
        txt项目名称.SelLength = Len(txt项目名称.Text)
    End If
    Exit Sub
    
errHandle:
    Call ErrCenter
End Sub

Private Function GetClassValue() As Integer
    Dim i As Integer
    For i = 0 To chkDept.Count - 1
        If chkDept(i).Value And chkDept(i).Enabled Then
            GetClassValue = i
            Exit Function
        End If
    Next
    GetClassValue = -1
End Function


Private Sub initColumn()
    Dim i As Integer
    '初始化表格
    With vsfList
        .Rows = 1
        .Cols = mColumn.Count
    End With
    
    'vsf列设置：列名，列宽，列对齐方式，固定列对齐方式（默认为居中对齐
    VsfGridColFormat vsfList, mColumn.选择, "选择", 500, flexAlignCenterCenter, "选择"
    VsfGridColFormat vsfList, mColumn.No, "入库单据号", 1500, flexAlignLeftCenter, "入库单据号"
    VsfGridColFormat vsfList, mColumn.库房, "库房", 1000, flexAlignLeftCenter, "库房"
    VsfGridColFormat vsfList, mColumn.序号, "序号", 640, flexAlignCenterCenter, "序号"
    VsfGridColFormat vsfList, mColumn.记录状态, "记录状态", 0, flexAlignLeftCenter, "记录状态"
    VsfGridColFormat vsfList, mColumn.项目id, "项目ID", 0, flexAlignLeftCenter, "项目ID"
    VsfGridColFormat vsfList, mColumn.项目信息, "项目信息", 2500, flexAlignLeftCenter, "项目信息"
    VsfGridColFormat vsfList, mColumn.规格, "规格", 1500, flexAlignLeftCenter, "规格"
    VsfGridColFormat vsfList, mColumn.批号, "批号", 600, flexAlignLeftCenter, "批号"
    VsfGridColFormat vsfList, mColumn.单位, "单位", 600, flexAlignLeftCenter, "单位"
    VsfGridColFormat vsfList, mColumn.数量, "数量", 1000, flexAlignRightCenter, "数量"
    VsfGridColFormat vsfList, mColumn.采购价, "采购价", 1000, flexAlignRightCenter, "采购价"
    VsfGridColFormat vsfList, mColumn.采购金额, "采购金额", 1000, flexAlignRightCenter, "采购金额"
    VsfGridColFormat vsfList, mColumn.发票号, "发票号", 1500, flexAlignLeftCenter, "发票号"
    VsfGridColFormat vsfList, mColumn.发票代码, "发票代码", 1500, flexAlignLeftCenter, "发票代码"
    VsfGridColFormat vsfList, mColumn.发票金额, "发票金额", 1500, flexAlignRightCenter, "发票金额"
    VsfGridColFormat vsfList, mColumn.标识, "标识", 0, flexAlignLeftCenter, "标识"
    VsfGridColFormat vsfList, mColumn.收发ID, "收发ID", 0, flexAlignLeftCenter, "收发ID"
    VsfGridColFormat vsfList, mColumn.随货单号, "随货单号", 0, flexAlignLeftCenter, "随货单号"
    
End Sub

Public Sub VsfGridColFormat(ByVal objGrid As VSFlexGrid, ByVal intCol As Integer, ByVal strColName As String, _
    ByVal lngColWidth As Long, ByVal intColAlignment As Integer, _
    Optional ByVal strColKey As String = "", Optional ByVal intFixedColAlignment As Integer = 4)
    'vsf列设置：列名，列宽，列对齐方式，固定列对齐方式（默认为居中对齐）
    With objGrid
        .TextMatrix(0, intCol) = strColName
        .ColWidth(intCol) = lngColWidth: If lngColWidth = 0 Then .ColHidden(intCol) = True
        .ColAlignment(intCol) = intColAlignment
        .ColKey(intCol) = strColKey
        .FixedAlignment(intCol) = intFixedColAlignment
    End With
End Sub


Private Sub SetColumn(ByVal rsRecord As ADODB.Recordset)
    Dim lngLoop As Long
    With vsfList
        .Redraw = flexRDNone
        .Rows = 1
        .Rows = rsRecord.RecordCount + 1
        For lngLoop = 1 To rsRecord.RecordCount
            .TextMatrix(lngLoop, mColumn.No) = rsRecord!入库单据号
            .Cell(flexcpForeColor, lngLoop, mColumn.项目信息, lngLoop, mColumn.项目信息) = IIf(rsRecord!标识 = 1, glngColor药品, IIf(rsRecord!标识 = 2, glngColor物资, IIf(rsRecord!标识 = 3, glngColor设备, glngColor卫材)))
            .TextMatrix(lngLoop, mColumn.库房) = rsRecord!库房
            .TextMatrix(lngLoop, mColumn.序号) = rsRecord!序号
            .TextMatrix(lngLoop, mColumn.记录状态) = rsRecord!记录状态
            .TextMatrix(lngLoop, mColumn.项目id) = rsRecord!项目id
            .TextMatrix(lngLoop, mColumn.项目信息) = rsRecord!项目信息
            .TextMatrix(lngLoop, mColumn.规格) = "" & rsRecord!规格
            .TextMatrix(lngLoop, mColumn.批号) = "" & rsRecord!批号
            .TextMatrix(lngLoop, mColumn.单位) = rsRecord!单位
            .TextMatrix(lngLoop, mColumn.数量) = zlStr.FormatEx(IIf(IsNull(rsRecord!数量), 0, rsRecord!数量), mintShowPriceDigit, , True)
            .ColFormat(mColumn.数量) = "#0.00000"
            .TextMatrix(lngLoop, mColumn.采购价) = zlStr.FormatEx(rsRecord!采购价, mintShowPriceDigit, , True)
            .ColFormat(mColumn.采购价) = "#0.00000"
            .TextMatrix(lngLoop, mColumn.采购金额) = zlStr.FormatEx(rsRecord!采购金额, mintShowAmountDigit, , True)
            .ColFormat(mColumn.采购金额) = "#0.00000"
            .TextMatrix(lngLoop, mColumn.发票金额) = zlStr.FormatEx(rsRecord!采购金额, mintShowAmountDigit, , True)
            .ColFormat(mColumn.发票金额) = "#0.00000"
            .TextMatrix(lngLoop, mColumn.标识) = rsRecord!标识
            .TextMatrix(lngLoop, mColumn.收发ID) = rsRecord!收发ID
            .TextMatrix(lngLoop, mColumn.随货单号) = "" & rsRecord!随货单号
            
            rsRecord.MoveNext
        Next
        
        If .Rows > 1 Then
            .Cell(flexcpFontBold, 1, mColumn.发票金额, .Rows - 1, mColumn.发票金额) = True '发票金额加粗
            .Cell(flexcpFontBold, 1, mColumn.选择, .Rows - 1, mColumn.选择) = True '选择加粗
        End If
        .Redraw = flexRDDirect
    End With

    If vsfList.Rows > 1 Then
        vsfList.Select 1, 1
    End If
End Sub

Private Sub txt项目名称_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then GetItem UCase(Trim(txt项目名称))
End Sub

Private Sub vsfList_DblClick()
    With vsfList
        If .Rows < 2 Then Exit Sub
        If .MouseRow = 0 Or .MouseCol = mColumn.发票金额 Then Exit Sub
            
        If .TextMatrix(.Row, mColumn.选择) = "√" Then
            .TextMatrix(.Row, mColumn.选择) = ""
            .TextMatrix(.Row, mColumn.发票号) = ""
            .TextMatrix(.Row, mColumn.发票代码) = ""
        Else
            .TextMatrix(.Row, mColumn.选择) = "√"
            .TextMatrix(.Row, mColumn.发票号) = txt发票号.Text
            .TextMatrix(.Row, mColumn.发票代码) = txt发票代码.Text
        End If
    End With
    
    AmountSum
End Sub

Private Sub vsfList_EnterCell()

    With vsfList
        .Editable = flexEDNone
        .FocusRect = flexFocusLight
        
        Select Case .Col
            Case mColumn.发票金额
                If Val(.TextMatrix(.Row, mColumn.标识)) <> 3 Then .Editable = flexEDKbdMouse
        End Select
        
    End With
    
End Sub

Private Sub vsfList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strkey As String
    
    If InStr(" '", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    
    With vsfList
        Select Case Col
            Case mColumn.发票金额
                If InStr("1234567890" + Chr(46) + Chr(8) + Chr(13), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                ElseIf KeyAscii = Asc(".") Then
                    If InStr(.EditText, ".") <> 0 Then     '只能存在一个小数点
                        KeyAscii = 0
                    End If
                End If
                
        End Select
    End With
End Sub

Private Function SaveCard() As Boolean
    Dim lngLoop As Long
    Dim strNO As String
    Dim lng序号 As Long
    Dim Str发票号 As String
    Dim str发票代码 As String
    Dim dat发票日期 As String
    Dim dbl发票金额 As Double
    Dim int操作标志 As Integer '1、未冲销单据修改发票信息; 2、部分冲销单据修改发票信息
    Dim arrSql As Variant
    
    
    arrSql = Array()
    SaveCard = False
    If vsfList.Rows < 2 Then Exit Function
    '检查是否输入供药单位
    If Trim(txt发票号.Text) = "" Then
        MsgBox "发票号不能为空！", vbInformation, gstrSysName
        txt发票号.SetFocus
        Exit Function
    End If
    
    With vsfList
        '检查是否全都未勾选
        For lngLoop = 1 To .Rows - 1
            If Trim(.TextMatrix(lngLoop, mColumn.选择)) = "√" Then Exit For
        Next
        If lngLoop = .Rows Then
            MsgBox "未选择单据信息，请检查！", vbInformation, gstrSysName
            vsfList.SetFocus
            Exit Function
        End If
        
        On Error GoTo errHandle
        Str发票号 = Trim(txt发票号.Text)
        str发票代码 = Trim(txt发票代码.Text)
        dat发票日期 = dtp发票日期.Value
        
        If MsgBox("即将保存选择的发票信息，是否继续？" & _
        vbCrLf & "发票号：" & Str发票号 & "    发票代码：" & IIf(str发票代码 = "", "无", str发票代码) _
        , vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName) = vbNo Then Exit Function
        
        For lngLoop = 1 To .Rows - 1
            If Trim(.TextMatrix(lngLoop, mColumn.选择)) = "√" Then '勾选的
                strNO = Trim(.TextMatrix(lngLoop, mColumn.No))
                lng序号 = Val(.TextMatrix(lngLoop, mColumn.序号))
                dbl发票金额 = IIf(Trim(.TextMatrix(lngLoop, mColumn.发票金额)) = "", 0, .TextMatrix(lngLoop, mColumn.发票金额))
                int操作标志 = IIf(Val(.TextMatrix(lngLoop, mColumn.记录状态)) = 1, 1, 2)
                
                If Val(.TextMatrix(lngLoop, mColumn.标识)) = 1 Then '药品
                    gstrSQL = "zl_药品外购发票信息_UPDATE("
                    'NO
                    gstrSQL = gstrSQL & "'" & strNO & "'"
                    '序号
                    gstrSQL = gstrSQL & "," & lng序号
                    '发票号
                    gstrSQL = gstrSQL & ",'" & Str发票号 & "'"
                    '发票日期
                    gstrSQL = gstrSQL & "," & IIf(dat发票日期 = "", "Null", "to_date('" & Format(dat发票日期, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                    '发票金额
                    gstrSQL = gstrSQL & "," & dbl发票金额
                    '供药单位ID
                    gstrSQL = gstrSQL & "," & Val(txt供应商.Tag)
                    '操作标志
                    gstrSQL = gstrSQL & "," & int操作标志
                    '发票代码
                    gstrSQL = gstrSQL & ",'" & str发票代码 & "'"
                    '自动付款标记
                    gstrSQL = gstrSQL & "," & IIf(mbln付款标志, 1, 0) & ""
                    gstrSQL = gstrSQL & ")"
                ElseIf Val(.TextMatrix(lngLoop, mColumn.标识)) = 5 Then '卫材
                    gstrSQL = "zl_材料外购发票信息_UPDATE( "
                    gstrSQL = gstrSQL & "'" & strNO & "',"
                    gstrSQL = gstrSQL & "" & Val(.TextMatrix(lngLoop, mColumn.记录状态)) & ","
                    gstrSQL = gstrSQL & "" & lng序号 & ","
                    gstrSQL = gstrSQL & "'" & Str发票号 & "',"
                    gstrSQL = gstrSQL & "" & IIf(dat发票日期 = "", "Null", "to_date('" & Format(dat发票日期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                    gstrSQL = gstrSQL & "" & dbl发票金额 & ","
                    gstrSQL = gstrSQL & "" & Val(txt供应商.Tag) & ","
                    gstrSQL = gstrSQL & IIf(str发票代码 = "", "NULL", "'" & str发票代码 & "'") & ")"
                ElseIf Val(.TextMatrix(lngLoop, mColumn.标识)) = 2 Then '物资
                    gstrSQL = "ZL_物资外购入库_Invoice( "
                    '收发记录ID（冲销过的单据取原始收发ID）
                    gstrSQL = gstrSQL & "'" & Val(.TextMatrix(lngLoop, mColumn.收发ID)) & "',"
                    '发票号
                    gstrSQL = gstrSQL & "'" & Str发票号 & "',"
                    '发票日期
                    gstrSQL = gstrSQL & "" & IIf(dat发票日期 = "", "Null", "to_date('" & Format(dat发票日期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                    '发票金额
                    gstrSQL = gstrSQL & "" & dbl发票金额 & ","
                    '发票代码
                    gstrSQL = gstrSQL & "" & IIf(str发票代码 = "", "NULL", "'" & str发票代码 & "'") & ","
                    '供应商
                    gstrSQL = gstrSQL & Val(txt供应商.Tag) & ")"
                ElseIf Val(.TextMatrix(lngLoop, mColumn.标识)) = 3 Then '设备
                    gstrSQL = "ZL_设备外购入库_ModifyFP("
                    '   no_in        IN  设备收发记录.no%TYPE := NULL,
                    gstrSQL = gstrSQL & "'" & strNO & "',"
                    '   设备id_IN        IN  设备收发记录.设备id%type:=null,
                    gstrSQL = gstrSQL & "" & Val(.TextMatrix(lngLoop, mColumn.项目id)) & ","
                    '   序号_IN      IN  设备收发记录.序号%type:=null,
                    gstrSQL = gstrSQL & "" & lng序号 & ","
                    '   随货单号_In
                    gstrSQL = gstrSQL & "" & IIf(Trim(.TextMatrix(lngLoop, mColumn.随货单号)) = "", "NULL", "'" & Trim(.TextMatrix(lngLoop, mColumn.随货单号)) & "'") & ","
                    '   发票号_IN        IN  设备收发记录.发票号码%TYPE := NULL,
                    gstrSQL = gstrSQL & "" & Str发票号 & ","
                    '   发票日期_IN      IN  设备收发记录.发票日期%TYPE := NULL
                    gstrSQL = gstrSQL & "" & IIf(dat发票日期 = "", "Null", "to_date('" & Format(dat发票日期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                    '   发票代码
                    gstrSQL = gstrSQL & IIf(str发票代码 = "", "NULL", "'" & str发票代码 & "'") & ")"

                End If
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            End If
        Next
        
        gcnOracle.BeginTrans
        For lngLoop = 0 To UBound(arrSql)
            Call zlDatabase.ExecuteProcedure(CStr(arrSql(lngLoop)), "SaveCard")
        Next
        gcnOracle.CommitTrans
        
    End With
    SaveCard = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub AmountSum()
    Dim lngLoop As Long
    Dim dblAmountSum As Double
    
    With vsfList
        For lngLoop = 1 To .Rows - 1
            If Trim(.TextMatrix(lngLoop, mColumn.选择)) = "√" Then '勾选的
                dblAmountSum = dblAmountSum + Val(.TextMatrix(lngLoop, mColumn.发票金额))
            End If
        Next
    End With
    
    stbThis.Panels(2).Text = "已选发票金额合计：" & dblAmountSum
End Sub

Private Sub vsfList_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strkey As String
    
    With vsfList
        If Trim(.TextMatrix(.Row, mColumn.选择)) = "√" Then mbln待更新 = True
        
        If Col = mColumn.发票金额 Then
            .EditText = Trim(.EditText)
            strkey = Trim(.EditText)
        
            If .TextMatrix(Row, Col) = "" Or strkey = "" Then
                MsgBox "对不起，金额必须输入！", vbOKOnly + vbInformation, gstrSysName
                Cancel = True
                Exit Sub
            End If
            If Not IsNumeric(strkey) And strkey <> "" Then
                MsgBox "对不起，金额必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                Cancel = True
                Exit Sub
            End If
            If Val(strkey) < 0 Then
                MsgBox "对不起，金额不能为负数,请重输！", vbInformation + vbOKOnly, gstrSysName
                Cancel = True
                Exit Sub
            End If
                
           strkey = zlStr.FormatEx(strkey, mintShowAmountDigit, , True)
            .EditText = strkey
        End If
    End With
End Sub
