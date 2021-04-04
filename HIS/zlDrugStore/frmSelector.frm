VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelector 
   Caption         =   "药品选择器"
   ClientHeight    =   6465
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9975
   Icon            =   "frmSelector.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6465
   ScaleWidth      =   9975
   Begin VB.CheckBox chkView 
      Caption         =   "显示停用药品(&V)"
      Height          =   255
      Left            =   5640
      TabIndex        =   29
      Top             =   143
      Width           =   1650
   End
   Begin MSComctlLib.ImageList imgsDrug 
      Left            =   2160
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelector.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelector.frx":180C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelector.frx":1DA6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "刷新(F5)"
      Height          =   350
      Left            =   8280
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "F5功能键刷新缓存"
      Top             =   90
      Width           =   975
   End
   Begin VB.CheckBox chkContinue 
      Caption         =   "连续选择(&M)"
      Height          =   180
      Left            =   6840
      TabIndex        =   4
      Top             =   180
      Width           =   1335
   End
   Begin VB.TextBox txtFilterFind 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2400
      TabIndex        =   3
      Top             =   120
      Width           =   3135
   End
   Begin VB.OptionButton optFilterFind 
      Caption         =   "查找(&I)"
      Height          =   180
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   160
      Width           =   950
   End
   Begin VB.OptionButton optFilterFind 
      Caption         =   "过滤(&F)"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   160
      Value           =   -1  'True
      Width           =   950
   End
   Begin MSComctlLib.ImageList imgsMain 
      Left            =   9000
      Top             =   120
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
            Picture         =   "frmSelector.frx":2340
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelector.frx":2692
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplit02_S 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   40
      Left            =   2880
      ScaleHeight     =   45
      ScaleWidth      =   2535
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4080
      Width           =   2535
   End
   Begin VB.PictureBox pic选定区 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   3120
      ScaleHeight     =   1455
      ScaleWidth      =   4815
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4560
      Width           =   4815
      Begin VB.PictureBox picOK 
         BackColor       =   &H00FFEDDD&
         BorderStyle     =   0  'None
         Height          =   220
         Left            =   3240
         Picture         =   "frmSelector.frx":29E4
         ScaleHeight     =   225
         ScaleWidth      =   270
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "选定"
         Top             =   0
         Width           =   270
      End
      Begin VB.PictureBox picUpDown01 
         BackColor       =   &H00FFEDDD&
         BorderStyle     =   0  'None
         Height          =   220
         Left            =   3600
         Picture         =   "frmSelector.frx":2D26
         ScaleHeight     =   225
         ScaleWidth      =   270
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   0
         Width           =   270
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf选定 
         Height          =   1125
         Left            =   0
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   360
         Width           =   4275
         _cx             =   7541
         _cy             =   1984
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   15724527
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   1
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSelector.frx":3068
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
         ExplorerBar     =   7
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
      Begin VB.Label lbl选定 
         BackColor       =   &H00FFEDDD&
         Caption         =   "选定药品"
         Height          =   255
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   3885
      End
   End
   Begin VB.PictureBox pic药品区 
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   3120
      ScaleHeight     =   3375
      ScaleWidth      =   4695
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   600
      Width           =   4695
      Begin VB.PictureBox picSetCols 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   220
         Index           =   0
         Left            =   120
         Picture         =   "frmSelector.frx":30DD
         ScaleHeight     =   225
         ScaleWidth      =   255
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox picSplit04_S 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   40
         Left            =   2040
         MousePointer    =   7  'Size N S
         ScaleHeight     =   45
         ScaleWidth      =   2535
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1560
         Width           =   2535
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf批次 
         Height          =   1485
         Left            =   0
         TabIndex        =   9
         Top             =   1800
         Width           =   4275
         _cx             =   7541
         _cy             =   2619
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   15724527
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
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
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSelector.frx":360F
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
         ExplorerBar     =   7
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
         Begin VB.PictureBox picSetCols 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   220
            Index           =   1
            Left            =   0
            Picture         =   "frmSelector.frx":3684
            ScaleHeight     =   225
            ScaleWidth      =   255
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   0
            Width           =   255
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf规格 
         Height          =   1485
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   4275
         _cx             =   7541
         _cy             =   2619
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   15724527
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
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
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSelector.frx":3BB6
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
         ExplorerBar     =   7
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
   Begin VB.PictureBox picSplit01_S 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   2760
      MousePointer    =   9  'Size W E
      ScaleHeight     =   4215
      ScaleWidth      =   45
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1200
      Width           =   40
   End
   Begin VB.PictureBox pic类型区 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   0
      ScaleHeight     =   5775
      ScaleWidth      =   2655
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   600
      Width           =   2655
      Begin VB.Frame fra形态 
         Caption         =   "中药形态"
         Height          =   495
         Left            =   0
         TabIndex        =   25
         Top             =   3120
         Width           =   2655
         Begin VB.CheckBox chk形态 
            BackColor       =   &H00FFEDDD&
            Caption         =   "散装"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   28
            Top             =   200
            Width           =   700
         End
         Begin VB.CheckBox chk形态 
            BackColor       =   &H00FFEDDD&
            Caption         =   "饮片"
            Height          =   180
            Index           =   1
            Left            =   960
            TabIndex        =   27
            Top             =   200
            Width           =   700
         End
         Begin VB.CheckBox chk形态 
            BackColor       =   &H00FFEDDD&
            Caption         =   "免煎剂"
            Height          =   180
            Index           =   2
            Left            =   1710
            TabIndex        =   26
            Top             =   200
            Width           =   900
         End
      End
      Begin VB.CheckBox chkChoose 
         BackColor       =   &H00FFEDDD&
         Caption         =   "全选"
         Height          =   180
         Left            =   1560
         TabIndex        =   22
         Top             =   3600
         Width           =   700
      End
      Begin VB.PictureBox picSplit03_S 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   40
         Left            =   0
         MousePointer    =   7  'Size N S
         ScaleHeight     =   45
         ScaleWidth      =   2535
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   3000
         Width           =   2535
      End
      Begin MSComctlLib.ListView lvw剂型 
         Height          =   1995
         Left            =   0
         TabIndex        =   7
         Top             =   3840
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   3519
         View            =   1
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         _Version        =   393217
         Icons           =   "imgsDrug"
         SmallIcons      =   "imgsDrug"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "名称"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.TreeView tvw类别 
         Height          =   2925
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   5159
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "imgsDrug"
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label lbl剂型 
         BackColor       =   &H00FFEDDD&
         Caption         =   "剂型"
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   3600
         Width           =   2565
      End
   End
   Begin VB.Label lblFilter 
      AutoSize        =   -1  'True
      Caption         =   "过滤(&F)"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   630
   End
End
Attribute VB_Name = "frmSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'格式： "剂型,,3,1000,r|..."
'   元素1：Key值；
'   元素2：Caption值（默认为Key值）；
'   元素3：列属性（0：内部显示，可移动；1：内部隐藏，不可移动，不可显示；2：用户隐藏；3：用户显示(默认值)）
'   元素4：列宽度（默认0）；
'   元素5：显示格式；s：字符串； n：数字； d：日期； t：时间； dt：日期时间
Private Enum enmColProperty
    cpKey = 0
    cpCaption
    cpDisplay
    cpWidth
    cpFormat
End Enum

Private Const MCON_规格 = _
    "剂型,,,1000|中药形态,,,1000,s|药名编码,,1,0|来源,,,1000|基本药物,,,1000|药典ID,,1,0|领用标志,,1,0|用途分类ID,,1,0|剂量单位,,1,1000" & _
    "|药品编码,,,1000|通用名称,,,1000|药品名称,,1,1000|商品名,,,1000|规格,,0,1000|生产商,,,1000|原产地,,,1000|药名ID,,1,0" & _
    "|药品ID,,1,0|上次采购价,,1,1000,n|售价,,,1000,n|售价单位,,0,1000|售价包装,,0,1000,n|门诊单位,,0,1000" & _
    "|门诊包装,,0,1000,n|住院单位,,0,1000|住院包装,,0,1000,n|药库单位,,0,1000|药库包装,,0,1000,n|可用数量,,,1000,n" & _
    "|库存数量,,1,1000,n|库存金额,,1,1000,n|库存差价,,1,1000,n|有效期,,,1000,n|药库分批,,,1000|药房分批,,,1000" & _
    "|时价,,,1000|指导批发价,,1,1000,n|加成率,,1,1000,n|库房货位,,,1000|批准文号,,,1000|实际数量,,1,1000,n" & _
    "|合同单位,,0,1000|药价级别,,,1000|留存数量,,1,,n|简码,,1,0|数字简码,,1,0|五笔码,,1,0"

Private Const MCON_批次 = _
    "RID,,1,0|库房,,,1000|批次,,1,1000|入库日期,,0,1000,d|批号,,,1000|生产日期,,,1000,d|有效期,,,1000,d|生产商,,,1000|原产地,,,1000" & _
    "|成本价,,,1000,n|售价,,,1000,n|可用数量,,,1000,n|库存数量,,,1000,n|库存金额,,,1000,n|库存差价,,,1000,n" & _
    "|上次供应商ID,,1,0|实际数量,,1,0,n|批准文号,,,1000|供应商,,,1000"

Private Const MCON_选定 = _
    "剂型,,1,1000|药名编码,,1,0|来源,,1,1000|基本药物,,1,1000|药品编码,,0,1000|通用名称,,,1000|药典ID,,1,0|用途分类ID,,1,0" & _
    "|剂量单位,,1,1000|药品名称,,1,1000|商品名,,,1000|规格,,0,1000|生产商,,0,1000|原产地,,0,1000|药名ID,,1,0|药品ID,,1,0|批号,,0,1000" & _
    "|上次采购价,,1,1000,n|售价,,0,1000,n|售价单位,,1,1000|售价包装,,1,1000,n|门诊单位,,1,1000" & _
    "|门诊包装,,1,1000,n|住院单位,,1,1000|住院包装,,1,1000,n|药库单位,,1,1000|药库包装,,1,1000,n|可用数量,,1,1000,n" & _
    "|库存数量,,1,1000,n|库存金额,,1,1000,n|库存差价,,1,1000,n|最大效期,,1,1000,n|药库分批,,1,1000|药房分批,,1,1000" & _
    "|时价,,1,1000|指导批发价,,1,1000,n|加成率,,1,1000,n|批准文号,,1,1000" & _
    "|合同单位,,1,1000|药价级别,,1,1000|留存数量,,1,,n" & _
    "|批次,,1,1000|生产日期,,1,1000,d|有效期,有效期至,1,1000,d|实际数量,,1,1000,n" & _
    "|上次供应商ID,,1,0|成本价,,1,1000,n"

Private Const mconint售价单位 As Integer = 1
Private Const mconint门诊单位 As Integer = 2
Private Const mconint住院单位 As Integer = 3
Private Const mconint药库单位 As Integer = 4

Private mintUnit As Integer             '单位系数：1-售价;2-门诊;3-住院;4-药库
Private mintCostDigit As Integer        '成本价小数位数
Private mintPriceDigit As Integer       '售价小数位数
Private mintNumberDigit As Integer      '数量小数位数
Private mintMoneyDigit As Integer       '金额小数位数

Private mstrCostFormat As String
Private mstrPriceFormat As String
Private mstrNumberFormat As String
Private mstrMoneyFormat As String

Private mStr成本价 As String
Private mStr单价 As String
Private mStr数量 As String
Private mStr金额 As String

Private Type WinLocate
    Left As Double
    Top As Double
End Type
Private WindowPosition As WinLocate         '窗体位置

Private MStrCaption As String
Private mbytStyle As Long                   '选择器显示模式。   0：查找选择模式； 1：模糊录入模式
'Private mlngSys As Long                     '系统号
'Private mlngMode As Long                    '模块号
Private mstrPrivs As String                 '当前操作模块权限
Private mfrmMain As Form                    '主窗体
Private mbyt编辑模式 As Byte                '1：入库； 2：出库
Private mstr简码 As String                  '查询录入的简码
Private mlng来源库房 As Long                '来源库房ID
Private mlng目标库房 As Long                '目标库房ID
Private mlng使用部门 As Long                '使用部门ID
Private mlng供应商 As Long                  '供应商ID
Private mbyt包含停用药品 As Byte         '是否显示停用药品（0-不显示停用药品，1-显示停用药品，2-根据注册表参数来确定）
Private mbln中药库房 As Boolean             '是否中药库房   True：是  False：否
Private mintStockCheck As Integer           '库存检测       0-不检查；1-检查，不足提醒；2-检查，不足禁止
Private mbyt库房性质 As Byte                '库房性质       1-药库；2-药房；3-制剂室
Private mrsReturn As ADODB.Recordset        '返回选定药品数据
Private mstrFilterClass As String           '药品分类的过滤条件
Private mstrMatch As String                 '简码匹配方式   0-双向匹配； 1-单向右匹配
'Private mbln明确申领批次 As Boolean         '明确申领药品批次
'Private mbln申领单 As Boolean               'True申领单；False移库单
Private mblnCheck As Boolean                '是否检测库存(盘点用)
Private mblnPrice As Boolean                '是否允许时价或批次药品零出库
Private mblnStore As Boolean                '显示库存
Private mbln空批次 As Boolean               '空批次记录显示
Private mstr空批次库房 As String            '空批次记录显示的库房
Private mblnMultiSel As Boolean             '可多选记录
Private mblnOK As Boolean
Private mblnCostView As Boolean             '查看成本价 true-允许查看 false-不允许查看

Private mstr规格 As String, mstr批次 As String, mstr选定 As String      '用户自定义的列头名、顺序
Private mlngLast As Long    '最后一次选中药品
Private mblnLoad As Boolean     '布尔型，判断窗体是否正在加载 true-正在，false-已经加载完成
Private mint按批次出库 As Integer           '0-不按批次出库,1-按批次出库
Public Function ShowMe( _
    ByVal FrmMain As Form, _
    ByVal bytStyle As Byte, _
    ByVal byt编辑模式 As Byte, _
    Optional ByVal str简码 As String, _
    Optional ByVal lngWinLeft As Long = 0, _
    Optional ByVal lngWinTop As Long = 0, _
    Optional ByVal lng来源库房 As Long = 0, _
    Optional ByVal lng目标库房 As Long = 0, _
    Optional ByVal lng使用部门 As Long = 0, _
    Optional ByVal lng供应商 As Long = 0, _
    Optional ByVal bln检测库存 As Boolean = True, _
    Optional ByVal bln检查批次或时价 As Boolean = True, _
    Optional ByVal bln显示库存 As Boolean = True, _
    Optional ByVal byt包含停用药品 As Byte = 0, _
    Optional ByVal bln可多选 As Boolean = True, _
    Optional ByVal strPrivs As String = "" _
) As ADODB.Recordset
    Dim strKeyName As String
    
    If grsMaster Is Nothing Or grsSlave Is Nothing Then
        MsgBox "药品选择器的数据未生成！", vbInformation, gstrSysName
        Exit Function
    End If
    
    WindowPosition.Left = lngWinLeft
    WindowPosition.Top = lngWinTop
    
    mbytStyle = bytStyle
    Set mfrmMain = FrmMain
'    mlngSys = lngSys
'    mlngMode = lngMode
    mbyt编辑模式 = byt编辑模式
    mstr简码 = str简码 'VerifyFilterStr(str简码)
    mlng来源库房 = lng来源库房
    mlng目标库房 = lng目标库房
    mlng使用部门 = lng使用部门
    mlng供应商 = lng供应商
    mblnCheck = bln检测库存
    mblnPrice = bln检查批次或时价
    mblnStore = bln显示库存
    mbyt包含停用药品 = byt包含停用药品
    mblnMultiSel = bln可多选
    mstrPrivs = strPrivs
    
    '恢复批次选择状态
    Select Case UCase(mfrmMain.Name)
        Case UCase("frmTransferCard")
            strKeyName = "药品移库管理"
        Case UCase("frmRequestDrugCard")
            strKeyName = "药品申领管理"
        Case UCase("frmDrawCard")
            strKeyName = "药品领用管理"
    End Select
    
    If UCase(mfrmMain.Name) = UCase("frmRequestDrugCard") Then
        mint按批次出库 = Val(zldatabase.GetPara("药品按批次出库", glngSys, 1343, 0))
    ElseIf UCase(mfrmMain.Name) = UCase("frmTransferCard") Then
        mint按批次出库 = Val(zldatabase.GetPara("药品按批次出库", glngSys, 1304, 1))
    ElseIf UCase(mfrmMain.Name) = UCase("frmDrawCard") Then
        mint按批次出库 = Val(zldatabase.GetPara("药品按批次出库", glngSys, 1305, 1))
    Else
        mint按批次出库 = 1
    End If
    
    mbln空批次 = False
    '盘点单要记录
    If UCase(FrmMain.Name) = UCase("frmCheckCard") Or UCase(FrmMain.Name) = UCase("frmCheckCourseCard") Then
        mbln空批次 = True
        If grsSlave.State = adStateOpen And grsSlave.RecordCount > 0 Then
            grsSlave.MoveFirst
            mstr空批次库房 = zlStr.Nvl(grsSlave!库房)
        Else
            mstr空批次库房 = Get部门名称(IIf(lng来源库房 = 0, lng目标库房, lng来源库房))
        End If
    ElseIf UCase(mfrmMain.Name) = UCase("frmTransferCard") Or UCase(mfrmMain.Name) = UCase("frmRequestDrugCard") Or UCase(mfrmMain.Name) = UCase("frmDrawCard") Then
'        chkBatch.Visible = True
    End If
    
    If txtFilterFind.Tag = "1" Then txtFilterFind.SetFocus
    
    Show vbModal, FrmMain
    Set ShowMe = mrsReturn
    Unload Me
End Function

Private Sub chkBatch_Click()
    If vsf选定.rows > 1 And mint按批次出库 = 0 Then
        If MsgBox("已经有选定药品存在，取消“按批次出库”将清除已选定的药品，你确定吗？" _
            , vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            vsf选定.rows = 1
            chkContinue.Tag = ""
            Call Form_Resize
        End If
    End If

    ViewVSF批次 vsf规格
End Sub

Private Sub chkChoose_Click()
    Dim i As Integer
    If chkChoose.Value = 2 Then Exit Sub
    For i = 1 To lvw剂型.ListItems.count
        lvw剂型.ListItems(i).Checked = chkChoose.Value
    Next
    '处理停用药品是否显示
    myFilter
    SetColor
End Sub

Private Sub chkContinue_Click()
    pic选定区.Visible = chkContinue.Value
    pic选定区.TabStop = chkContinue.Value
    picSplit02_S.Visible = chkContinue.Value
    If chkContinue.Value = 1 And chkContinue.Tag <> "msg" Then
        vsf选定.rows = 1
        chkContinue.Tag = ""
        lbl选定.Caption = "选定药品"
    Else
        If vsf选定.rows > 1 And chkContinue.Tag <> "msg" Then
            If MsgBox("已经有选定药品存在，取消“连续选择”将清除已选定的药品，你确定吗？" _
                , vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                vsf选定.rows = 1
                chkContinue.Tag = ""
                Call Form_Resize
                Exit Sub
            End If
            chkContinue.Tag = "msg"
            chkContinue.Value = 1
            Exit Sub
        Else
            chkContinue.Tag = ""
        End If
    End If
    Call Form_Resize
End Sub

Private Sub chkView_Click()
    If chkView.Visible = False Then Exit Sub
    
    myFilter
    SetColor
End Sub

Private Sub chk形态_Click(index As Integer)
    Call myFilter
    SetColor
End Sub

'
'
'
Private Sub cmdRefresh_Click()
    Dim strFind As String
    
    cmdRefresh.Enabled = False
    Me.MousePointer = vbHourglass
    On Error GoTo errHandle
    If mfrmMain.Caption Like "药品移库单*" Then
        Call SetSelectorRS(mbyt编辑模式, "药品移库管理", mlng来源库房, mlng目标库房, , mlng供应商, , False)
    
    ElseIf mfrmMain.Caption Like "药品盘点记录单*" Or mfrmMain.Caption Like "药品盘点表*" Then
        Call SetSelectorRS(mbyt编辑模式, "药品盘点管理", mlng来源库房, mlng目标库房, , mlng供应商, , IIf(mbyt包含停用药品 = 1, True, False))
    
    ElseIf mfrmMain.Caption Like "库存差价调整单*" Then
        Call SetSelectorRS(mbyt编辑模式, "库存差价调整管理", mlng来源库房, mlng目标库房, , mlng供应商, , False)
    
    ElseIf mfrmMain.Caption Like "药品领用单*" Then
        Call SetSelectorRS(mbyt编辑模式, "药品领用管理", mlng来源库房, mlng目标库房, , mlng供应商, , False)
    
    ElseIf mfrmMain.Caption Like "药品其他入库单*" Then
        Call SetSelectorRS(mbyt编辑模式, "药品其他入库管理", mlng来源库房, mlng目标库房, , mlng供应商, , False)
    
    ElseIf mfrmMain.Caption Like "药品其他出库单*" Then
        Call SetSelectorRS(mbyt编辑模式, "药品其他出库管理", mlng来源库房, mlng目标库房, , mlng供应商, , False)
    
    ElseIf mfrmMain.Caption Like "药品外购入库单*" Then
        Call SetSelectorRS(mbyt编辑模式, "药品外购入库管理", mlng来源库房, mlng目标库房, , mlng供应商, , False)
    
    ElseIf mfrmMain.Caption Like "药品移库单*" Then
        Call SetSelectorRS(mbyt编辑模式, "药品移库管理", mlng来源库房, mlng目标库房, , mlng供应商, , False)
    
    ElseIf mfrmMain.Caption Like "药品申领单*" Then
        Call SetSelectorRS(mbyt编辑模式, "药品申领管理", mlng来源库房, mlng目标库房, , mlng供应商, , False)
    
    ElseIf mfrmMain.Caption Like "药品计划单*" Then
        Call SetSelectorRS(mbyt编辑模式, "药品计划管理", mlng来源库房, mlng目标库房, , mlng供应商, , False)
    Else '过滤
        Call SetSelectorRS(mbyt编辑模式, "查找药品", mlng来源库房, mlng目标库房, , mlng供应商, , True)
    End If
    Me.MousePointer = vbDefault
    '重新填充数据
    If mbytStyle = 0 Then '选择模式
        mstrFilterClass = GetFilterClass(tvw类别.SelectedItem)
        If tvw类别.SelectedItem.Children = 0 And tvw类别.SelectedItem.Key Like "Root*" Then '如果是根节点且该节点下面无数据的话则没有数据
            mstrFilterClass = "用途分类id=99999999999999"
        Else
            mstrFilterClass = Left(mstrFilterClass, Len(mstrFilterClass) - 4)
        End If
        grsMaster.Filter = mstrFilterClass
        Call FillVSF(grsMaster, vsf规格)
    Else '录入模式
        strFind = Trim(txtFilterFind.Text)
        txtFilterFind.Tag = ""
        mstrFilterClass = GetFilterSimpleCode(strFind)
        grsMasterInput.Filter = mstrFilterClass
        Call FillVSF(grsMasterInput, vsf规格)
    End If
    ViewVSF批次 vsf规格
    cmdRefresh.Enabled = True
    
    myFilter
    
    SetColor
    
    Exit Sub
    
errHandle:
    Me.MousePointer = vbDefault
    cmdRefresh.Enabled = True
    Call ErrCenter
End Sub

Private Sub Form_Activate()
    '简码录入匹配只有一条数据
    If vsf规格.rows = 2 And mbytStyle = 1 Then
        If vsf批次.Visible = True Then
            If vsf批次.rows = 2 Then
                Call vsf规格_DblClick
                If mblnOK Then Unload Me
            End If
        Else
            Call vsf规格_DblClick
            If mblnOK Then Unload Me
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    ElseIf KeyCode = vbKeyF5 Then
        If cmdRefresh.Enabled Then Call cmdRefresh_Click
    End If
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo errHandle
    mblnLoad = True
    mblnCostView = zlStr.IsHavePrivs(mstrPrivs, "查看成本价")
    picSplit03_S.Top = Me.ScaleHeight - lvw剂型.Height - lbl剂型.Height - picSplit03_S.Height
    picSplit04_S.Top = Me.ScaleHeight - vsf批次.Height - picSplit04_S.Height - pic药品区.Top
    
    If mbytStyle = 1 Then Height = 4000
    
    MStrCaption = GetText(GetParentWindow(mfrmMain.hWnd))
    Call RestoreWinState(Me, App.ProductName, mfrmMain.Caption & mbytStyle)
    
    picSplit02_S.Visible = False
    picSplit04_S.Visible = False
    pic选定区.Visible = False
    vsf批次.Visible = False
    vsf批次.TabStop = False
    
    optFilterFind(0).Visible = False: optFilterFind(1).Visible = False  '已取消，控件和代码暂时保留
    optFilterFind(0).TabStop = False: optFilterFind(1).TabStop = False
    
    pic类型区.Visible = mbytStyle = 0
    lblFilter.Visible = mbytStyle = 1
    txtFilterFind.Visible = mbytStyle = 1: txtFilterFind.TabStop = mbytStyle = 1
    
    '简码匹配方式
    If mbytStyle = 1 Then mstrMatch = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = "0", "%", "")
    
    '获取是否库存检查设置
    gstrSQL = "Select Nvl(检查方式,0) 库存检查 From 药品出库检查 Where 库房ID=[1]"
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "获取是否库存检查设置", mlng来源库房)
    If Not rsTmp.EOF Then
        mintStockCheck = rsTmp!库存检查
    End If
    rsTmp.Close
    
    '检查源库房是否为药库
    mbyt库房性质 = GetStockType(mlng来源库房)
        
    '数量单位
    If MStrCaption Like "药品申领管理*" Then
        Call GetDrugDigit(mlng使用部门, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
        mlng使用部门 = 0
    ElseIf MStrCaption Like "药品移库管理*" Then
        Call GetDrugDigit(mlng使用部门, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
        mlng使用部门 = 0
    Else
        Call GetDrugDigit(IIf(mlng来源库房 = 0, mlng目标库房, mlng来源库房), MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    End If

    '单价、金额格式
'    mstrCostFormat = "'999999999990." & String(gtype_UserDrugDigits.Digit_成本价, "0") & "'"
'    mstrPriceFormat = "'999999999990." & String(gtype_UserDrugDigits.Digit_零售价, "0") & "'"
'    mstrNumberFormat = "'999999999990." & String(gtype_UserDrugDigits.Digit_数量, "0") & "'"
'    mstrMoneyFormat = "'999999999990." & String(gtype_UserDrugDigits.Digit_金额, "0") & "'"
'    mStr成本价 = "####0." & String(gtype_UserDrugDigits.Digit_成本价, "0") & ";-####0." & String(gtype_UserDrugDigits.Digit_成本价, "0") & "; ;"
'    mStr单价 = "####0." & String(gtype_UserDrugDigits.Digit_零售价, "0") & ";-####0." & String(gtype_UserDrugDigits.Digit_零售价, "0") & "; ;"
'    mStr数量 = "####0." & String(gtype_UserDrugDigits.Digit_数量, "0") & ";-####0." & String(gtype_UserDrugDigits.Digit_数量, "0") & "; ;"
'    mStr金额 = "####0." & String(gtype_UserDrugDigits.Digit_金额, "0") & ";-####0." & String(gtype_UserDrugDigits.Digit_金额, "0") & "; ;"
    mStr成本价 = "####0." & String(mintCostDigit, "0") & ";-####0." & String(mintCostDigit, "0") & "; ;"
    mStr单价 = "####0." & String(mintPriceDigit, "0") & ";-####0." & String(mintPriceDigit, "0") & "; ;"
    mStr数量 = "####0." & String(mintNumberDigit, "0") & ";-####0." & String(mintNumberDigit, "0") & "; ;"
    mStr金额 = "####0." & String(mintMoneyDigit, "0") & ";-####0." & String(mintMoneyDigit, "0") & "; ;"
    
    'VSF表格头单独重复处理
    mstr规格 = Trim(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name & mfrmMain.Caption & mbytStyle & "\VSFlexGrid", _
        vsf规格.Name & vsf规格.Tag & "名称", ""))
    mstr批次 = Trim(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name & mfrmMain.Caption & mbytStyle & "\VSFlexGrid", _
        vsf批次.Name & vsf批次.Tag & "名称", ""))
    mstr选定 = Trim(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name & mfrmMain.Caption & mbytStyle & "\VSFlexGrid", _
        vsf选定.Name & vsf选定.Tag & "名称", ""))
    
    '参数控制显示
    Call ParamsColsHead
    
    InitVSF vsf规格
    InitVSF vsf批次
    InitVSF vsf选定
    SetVSFHead vsf规格, mstr规格
    SetVSFHead vsf批次, mstr批次
    SetVSFHead vsf选定, mstr选定
    vsf选定.rows = 1
    
    '装载数据 tvw类别、lvw剂型
    If mbytStyle = 0 Then Call Fill_TVW类别
    
    If mbytStyle = 1 Then
        'txtFilterFind.Tag = "1"
        txtFilterFind.Text = mstr简码
        'txtFilterFind.Tag = ""
        'WindowPosition在ShowMe()已经赋值
    Else
        If Not tvw类别.SelectedItem Is Nothing Then
            tvw类别_NodeClick tvw类别.SelectedItem
        End If
        '屏幕居中
        'Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
        '所有者居中
        WindowPosition.Left = mfrmMain.Left + (mfrmMain.Width - Me.Width) \ 2
        If WindowPosition.Left < 0 Then WindowPosition.Left = 0
        WindowPosition.Top = mfrmMain.Top + (mfrmMain.Height - Me.Height) \ 2
        If WindowPosition.Top < 0 Then WindowPosition.Top = 0
    End If
    Move WindowPosition.Left, WindowPosition.Top
    
    vsf规格.TabIndex = 0: vsf批次.TabIndex = 1
    
    '初始化返回数据集对象
    Call InitReturnRecord
    
    chkContinue.Visible = mblnMultiSel
    chkView.Visible = mbyt包含停用药品 = 2
    
    chkView.Value = GetSetting("ZLSOFT", "私有模块\ZLHIS\zl9MediStore", "显示停用药品", 0)
    
    myFilter
    SetColor
    
    '显示批次
    If vsf规格.rows > 1 Then
        For i = 1 To vsf规格.rows - 1
            If vsf规格.RowHidden(i) = False Then
                vsf规格.Row = i
                Exit For
            End If
        Next
        
        
        ViewVSF批次 vsf规格
'        If mbln空批次 Then
'            If vsf批次.Rows > 2 Then
'                vsf批次.Row = 2
'            ElseIf vsf批次.Rows > 1 Then
'                vsf批次.Row = 1
'            End If
'        Else
'            If vsf批次.Rows > 1 Then vsf批次.Row = 1
'        End If
    End If
    
    mblnLoad = False
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub




Private Sub myFilter()
'功能：根据“显示停用药品”、“剂型”和“中药形态”过滤数据
    Dim i As Integer
    Dim str形态 As String
    Dim str剂型 As String
    
     '1、判断选择的形态
    str形态 = ";"
    For i = 0 To chk形态.count - 1
        If chk形态(i).Value = 1 Then str形态 = str形态 & chk形态(i).Caption & ";"
    Next
    If str形态 = ";" Then '都不勾选，默认为全选
        For i = 0 To chk形态.count - 1
            str形态 = str形态 & chk形态(i).Caption & ";"
        Next
    End If
    '2、判断选择的剂型
    If chkChoose.Value = 2 Then '部分选择
        With lvw剂型
            str剂型 = ";"
            For i = 1 To .ListItems.count
                If .ListItems(i).Checked Then str剂型 = str剂型 & .ListItems(i).Text & ";"
            Next
        End With
    Else '全选
        With lvw剂型
            str剂型 = ";"
            For i = 1 To .ListItems.count
                str剂型 = str剂型 & .ListItems(i).Text & ";"
            Next
        End With
    End If
    
    '3、过滤

    With vsf规格
         .Redraw = flexRDNone
        For i = 1 To .rows - 1
        
            .RowHidden(i) = False '每条过滤前，是显示
            
            '停用药品不显示
            If mbyt包含停用药品 = 2 Then
                If chkView.Value = 0 Then
                    .RowHidden(i) = .TextMatrix(i, .ColIndex("停用")) = "是"
                End If
            Else
                If mbyt包含停用药品 <> 1 Then .RowHidden(i) = .TextMatrix(i, .ColIndex("停用")) = "是"
            End If
                
            '模糊查询模式不过滤剂型
            If mbytStyle <> 1 And .RowHidden(i) = False Then .RowHidden(i) = InStr(str剂型, ";" & .TextMatrix(i, .ColIndex("剂型")) & ";") = 0 '剂型过滤
            
            If .RowHidden(i) = False And fra形态.Visible = True Then '只针对中草药，形态过滤
                .RowHidden(i) = InStr(str形态, ";" & .TextMatrix(i, .ColIndex("中药形态")) & ";") = 0
            End If
 
        Next
        .Redraw = flexRDDirect
    End With
    
End Sub

Private Function GetStockType(ByVal lngStockid As Long) As Byte
'------------------------------------------------------------------------
'功能：获取药品库房的性质
'参数：
'   lngStockID：库房ID
'返回：0：未找到  1：药库  2：药房  3：制剂室
'------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strsql As String
    
    If lngStockid <= 0 Then Exit Function
    
    On Error GoTo errHandle
    GetStockType = 3
    strsql = "select count(部门ID) rec from 部门性质说明 where (工作性质 like '%制剂室' or 工作性质 like '%药房') And 部门id=[1] "
    Set rsTmp = zldatabase.OpenSQLRecord(strsql, "获取指定部门性质", lngStockid)
    If rsTmp!Rec > 0 Then
        GetStockType = 2
    Else
        strsql = "select count(部门ID) rec from 部门性质说明 where 工作性质 like '%药库' And 部门id=[1] "
        Set rsTmp = zldatabase.OpenSQLRecord(strsql, "获取指定部门性质", lngStockid)
        If rsTmp!Rec > 0 Then
            GetStockType = 1
        End If
    End If
    rsTmp.Close
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Resize()
    Dim bln中药 As Boolean
    
    If WindowState = 1 Then Exit Sub
    
    On Error Resume Next
    If mbytStyle = 1 Then
        If Me.Height < 3000 Then Me.Height = 3000
    Else
        If Me.Height < 5835 Then Me.Height = 5835
    End If
    If Me.Width < 8415 Then Me.Width = 8415
    
    cmdRefresh.Left = ScaleWidth - cmdRefresh.Width - 150
    If mbytStyle = 1 Then
        chkContinue.Left = cmdRefresh.Left - chkContinue.Width - 150
        chkView.Left = cmdRefresh.Left - chkView.Width
    Else
        chkContinue.Left = picSplit01_S.Left + picSplit01_S.Width + 75
        chkView.Left = picSplit01_S.Left + picSplit01_S.Width + 75
    End If

'    If optFilterFind(1).Visible Then
'        txtFilterFind.Left = optFilterFind(1).Left + optFilterFind(1).Width + 50
'        txtFilterFind.Width = ScaleWidth - txtFilterFind.Left - chkContinue.Width - cmdRefresh.Width - 500
'    Else
'        txtFilterFind.Left = optFilterFind(1).Left
'        txtFilterFind.Width = ScaleWidth - optFilterFind(1).Left - chkContinue.Width - cmdRefresh.Width - 500
'    End If
    If lblFilter.Visible Then
        lblFilter.Top = txtFilterFind.Top + 30
        txtFilterFind.Left = lblFilter.Left + lblFilter.Width + 50
        txtFilterFind.Width = ScaleWidth - txtFilterFind.Left - chkContinue.Width - cmdRefresh.Width - 500
    End If
    
    'pic类型区.Visible = mbytStyle <> 1
    picSplit01_S.Visible = mbytStyle <> 1       '不等于模糊录入模式
    picSplit03_S.Visible = mbytStyle <> 1
    
    If pic类型区.Visible Then
        If Not tvw类别.SelectedItem Is Nothing Then
            If tvw类别.SelectedItem.Tag = "3" Or tvw类别.SelectedItem.Tag = "Root7" Then
                bln中药 = True
            End If
        End If
        With pic类型区
            .Top = 20
            .Left = 0
            .Height = ScaleHeight - .Top
            .Width = picSplit01_S.Left
        End With
        With tvw类别
            .Top = 0
            .Left = 0
            .Width = pic类型区.Width
            .Height = picSplit03_S.Top
        End With
        With picSplit03_S
            If .Top > ScaleHeight - 2000 Then .Top = ScaleHeight - 2000
            .Left = 0
            .Width = pic类型区.Width
        End With
        With fra形态
            If bln中药 = True Then
                .Visible = True
                .Top = picSplit03_S.Top + picSplit03_S.Height
                .Left = 0
                .Width = pic类型区.Width
            Else
                .Visible = False
            End If
        End With
        With lbl剂型
            If bln中药 = True Then
                .Top = fra形态.Top + fra形态.Height
            Else
                .Top = picSplit03_S.Top + picSplit03_S.Height
            End If
            .Left = 0
            .Width = pic类型区.Width
        End With
        With chkChoose
            .Top = lbl剂型.Top + 10
            .Left = lbl剂型.Width - chkChoose.Width
        End With
        With lvw剂型
            .Top = lbl剂型.Height + lbl剂型.Top
            .Left = 0
            .Height = pic类型区.Height - lbl剂型.Top - lbl剂型.Height
            .Width = pic类型区.Width
        End With
        
        With picSplit01_S
            .Visible = mbytStyle <> 1
            .Top = pic类型区.Top
            '.Left = pic类型区.Width
            .Height = pic类型区.Height
        End With
    End If
    
    With pic药品区
        .Top = 550
        .Left = IIf(pic类型区.Visible, picSplit01_S.Left + picSplit01_S.Width, 0)
        .Width = IIf(pic类型区.Visible, ScaleWidth - pic类型区.Width - picSplit01_S.Width, ScaleWidth)
        If chkContinue.Value Then
            .Height = ScaleHeight - .Top - IIf(pic类型区.Tag = "展开", pic选定区.Height, lbl选定.Height + picSplit02_S.Height)
        Else
            .Height = ScaleHeight - .Top
        End If
    End With
    With vsf规格
        .Top = 0
        .Left = 0
        .Height = IIf(vsf批次.Visible, picSplit04_S.Top, pic药品区.Height)
        .Width = pic药品区.Width
    End With
    With picSetCols(0)
        .Top = 30
        .Left = 40
        .Height = 220
        .Width = 220
    End With
    
    If picSplit04_S.Visible Then
        With picSplit04_S
            If .Top > pic药品区.ScaleHeight - 1000 Then .Top = pic药品区.ScaleHeight - 1000
            .Left = 0
            .Width = pic药品区.Width
        End With
        With vsf批次
            If .Visible Then
                .Top = picSplit04_S.Top + picSplit04_S.Height
                .Left = 0
                .Width = pic药品区.Width
                .Height = pic药品区.Height - picSplit04_S.Top
            End If
        End With
        With picSetCols(1)
            If .Visible Then
                .Top = 30
                .Left = 40
                .Height = 220
                .Width = 220
            End If
        End With
    End If
    
    If picSplit02_S.Visible Then
        With picSplit02_S
            .Top = pic药品区.Top + pic药品区.Height
            .Left = pic药品区.Left
            .Width = pic药品区.Width
        End With
    End If
    
    If pic选定区.Visible Then
        With pic选定区
            .Tag = "收缩"
            picSplit02_S.MousePointer = 0
            Set picUpDown01.Picture = imgsMain.ListImages(2).Picture
            If .Tag = "展开" Then
                .Height = ScaleHeight - pic药品区.Top - pic药品区.Height - picSplit02_S.Height
            Else
                .Height = lbl选定.Height
            End If
            .Top = ScaleHeight - .Height
            .Left = pic药品区.Left 'IIf(pic类型区.Visible, pic类型区.Width, 0)
            .Width = ScaleWidth - IIf(pic类型区.Visible, pic类型区.Width + picSplit01_S.Width, 0)
        End With
        With lbl选定
            .Top = 0
            .Left = 0
            .Width = pic选定区.Width
        End With
        With picUpDown01
            .Left = pic选定区.Width - .Width
            .Top = 0
        End With
        With picOK
            .Left = picUpDown01.Left - .Width
            .Top = 0
        End With
        With vsf选定
            .Visible = pic选定区.Tag = "展开"
            .Top = lbl选定.Height
            .Left = 0
            .Width = lbl选定.Width
            If .Visible Then
                .Height = pic选定区.Height - lbl选定.Height
            End If
        End With
    End If
    err.Clear: On Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strKeyName As String
    
    Call SaveWinState(Me, App.ProductName, mfrmMain.Caption & mbytStyle)
    '单独保存VSF表格头状态
    mstr规格 = GetVSFHead(vsf规格)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name & mfrmMain.Caption & mbytStyle & "\VSFlexGrid", _
        vsf规格.Name & vsf规格.Tag & "名称", _
        IIf(mstr规格 = "", MCON_规格, mstr规格)

    mstr批次 = GetVSFHead(vsf批次)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name & mfrmMain.Caption & mbytStyle & "\VSFlexGrid", _
        vsf批次.Name & vsf批次.Tag & "名称", _
        IIf(mstr批次 = "", MCON_批次, mstr批次)
    
    mstr选定 = GetVSFHead(vsf选定)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name & mfrmMain.Caption & mbytStyle & "\VSFlexGrid", _
        vsf选定.Name & vsf选定.Tag & "名称", _
        IIf(mstr选定 = "", MCON_选定, mstr选定)
    
    '保存注册表信息(是否显示停用药品)
    SaveSetting "ZLSOFT", "私有模块\ZLHIS\zl9MediStore", "显示停用药品", chkView.Value
    
'    If chkBatch.Visible = True Then
'        '批次选择状态保存到注册表
'        Select Case UCase(mfrmMain.Name)
'            Case UCase("frmTransferCard")
'                strKeyName = "药品移库管理"
'            Case UCase("frmRequestDrugCard")
'                strKeyName = "药品申领管理"
'            Case UCase("frmDrawCard")
'                strKeyName = "药品领用管理"
'        End Select
'
'        SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & strKeyName, "按批次填单", IIf(chkBatch.Value = 1, 1, 0)
'    End If
    
    mlngLast = 0
End Sub

Private Sub Lvw剂型_ItemCheck(ByVal Item As MSComctlLib.listItem)
    Dim i As Integer
    
    If Item.Checked Then
        chkChoose.Value = 2
        For i = 1 To lvw剂型.ListItems.count
            If lvw剂型.ListItems(i).Checked = False Then
               
                myFilter
                SetColor
                Exit Sub
            End If
        Next
        chkChoose.Value = 1
    Else
        For i = 1 To lvw剂型.ListItems.count
            If lvw剂型.ListItems(i).Checked Then
                chkChoose.Value = 2
                
                myFilter
                SetColor
                Exit Sub
            End If
        Next
        chkChoose.Value = 0
    End If
    
    myFilter
    SetColor
End Sub

Private Sub optFilterFind_Click(index As Integer)
    txtFilterFind.SetFocus
End Sub

Private Sub picSetCols_Click(index As Integer)
    Dim frm列设置 As New frmVsColSel
    Dim vRect As RECT
    
    If index = 0 Then
        vRect = zlControl.GetControlRect(vsf规格.hWnd)
        frm列设置.ShowColSet Me, "列设置", vsf规格, _
            vRect.Left, vRect.Top + picSetCols(0).Top + picSetCols(0).Height + 10, _
            Me.Top + Me.Height - (vRect.Top + picSetCols(0).Top + picSetCols(0).Height + 120)
    Else
        vRect = zlControl.GetControlRect(vsf批次.hWnd)
        frm列设置.ShowColSet Me, "列设置", vsf批次, _
            vRect.Left, vRect.Top - 4000, _
            4000
    End If
End Sub

Private Sub picSetCols_MouseDown(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    picSetCols(index).BorderStyle = 1
End Sub

Private Sub picSetCols_MouseUp(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    picSetCols(index).BorderStyle = 0
End Sub

Private Sub picSplit01_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    With picSplit01_S
        If .Left + x < 2000 Then Exit Sub
        If .Left + x > ScaleWidth - 2000 Then Exit Sub
        .Move .Left + x, .Top
    End With
    With pic类型区
        .Width = .Width + x
    End With
    With pic药品区
        .Left = .Left + x
        .Width = .Width + x
    End With
    With picSplit02_S
        .Left = pic药品区.Left
        .Width = pic药品区.Width
    End With
    With pic选定区
        .Left = pic药品区.Left
        .Width = pic药品区.Width
    End With
    Call Form_Resize
End Sub

Private Sub picSplit02_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    If Not (picSplit02_S.MousePointer = 7 Or picSplit02_S.MousePointer = 0) Then Exit Sub
    With picSplit02_S
        If .Top + y < 1500 Then Exit Sub
        If .Top + y > Me.ScaleHeight - picSplit02_S.Height - lbl选定.Height Then Exit Sub
        .Move .Left, .Top + y
    End With
    With pic选定区
        .Top = picSplit02_S.Top + picSplit02_S.Height
        .Height = Me.ScaleHeight - .Top
    End With
    With vsf选定
        .Top = lbl选定.Height
        .Height = pic选定区.Height - lbl选定.Height
    End With
End Sub

Private Sub picSplit02_S_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If picSplit02_S.Top >= Me.ScaleHeight - picSplit02_S.Height - lbl选定.Height Then
        Call Form_Resize
    End If
End Sub

Private Sub picSplit03_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    With picSplit03_S
        If .Top + y < tvw类别.Top + 1000 Then Exit Sub
        If .Top + y > ScaleHeight - 2000 Then Exit Sub
        .Move .Left, .Top + y
    End With
    With tvw类别
        .Height = picSplit03_S.Top - .Top
    End With
    With fra形态
        .Top = picSplit03_S.Top + picSplit03_S.Height
    End With
    With lbl剂型
         If tvw类别.SelectedItem.Tag = "3" Then
            .Top = fra形态.Top + fra形态.Height
        Else
            .Top = picSplit03_S.Top + picSplit03_S.Height
        End If
        .Left = 0
        .Width = pic类型区.Width
    End With
    With chkChoose
        .Top = lbl剂型.Top + 10
    End With
    With lvw剂型
        .Top = lbl剂型.Height + lbl剂型.Top
        .Left = 0
        .Height = pic类型区.Height - lbl剂型.Top - lbl剂型.Height
        .Width = pic类型区.Width
    End With
End Sub

Private Sub picSplit04_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    With picSplit04_S
        If .Top + y < vsf规格.Top + 1000 Then Exit Sub
        If .Top + y > pic药品区.ScaleHeight - 1000 Then Exit Sub
        .Move .Left, .Top + y
    End With
    With vsf规格
        .Height = picSplit04_S.Top - .Top
    End With
    If vsf批次.Visible Then
        With vsf批次
            .Top = picSplit04_S.Top + picSplit04_S.Height
            .Height = pic药品区.Height - .Top
        End With
    End If
End Sub

Private Sub picOK_Click()
    Call CombinateRec
    Unload Me
End Sub

Private Sub picUpDown01_Click()
    If pic选定区.Tag = "展开" Then
        pic选定区.Tag = "收缩"
        Set picUpDown01.Picture = imgsMain.ListImages(2).Picture
    Else
        pic选定区.Tag = "展开"
        Set picUpDown01.Picture = imgsMain.ListImages(1).Picture
    End If
    ViewVSF选定 pic选定区.Tag = "展开"
End Sub

Private Sub ViewVSF选定(ByVal blnDisp As Boolean)
    Dim i As Integer
    Dim y As Single
    
    vsf选定.Visible = blnDisp: vsf选定.TabStop = blnDisp
    If blnDisp Then
        picSplit02_S.MousePointer = 7
        For i = picSplit02_S.Top To Me.ScaleHeight \ 2 Step -100
            picSplit02_S.Top = i
            picSplit02_S_MouseMove 1, 0, 0, y
        Next
    Else
        picSplit02_S.MousePointer = 0
        For i = picSplit02_S.Top To Me.ScaleHeight - picSplit02_S.Height - lbl选定.Height Step 100
            picSplit02_S.Top = i
            picSplit02_S_MouseMove 1, 0, 0, y
        Next
    End If
End Sub

Private Sub Fill_TVW类别()
    Dim nodTmp As Node
    Dim rsTmp As ADODB.Recordset
    Dim lng库房id As Long
    Dim Int末级 As Integer
    
    On Error GoTo errHandle
    gstrSQL = "Select 编码, 名称 From 诊疗项目类别 " & _
              "Where Instr([1], 编码, 1) > 0 " & _
              "Order by 编码 "
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, "567")
    
    With tvw类别
        .Nodes.Clear
'        Set nodTmp = .Nodes.Add(, , "Root", "所有", 2, 2)
        Do While Not rsTmp.EOF
            Set nodTmp = .Nodes.Add(, , "Root" & rsTmp!名称, rsTmp!名称, 2, 2)
            nodTmp.Tag = "Root" & rsTmp!编码
            rsTmp.MoveNext
        Loop
        rsTmp.Close
    End With
    
    '如果是入库，以入库库房为准，否则以出库库房为准
    lng库房id = IIf(mbyt编辑模式 = 1, mlng目标库房, mlng来源库房)
    If lng库房id <> 0 Then
        '提取该库房现有剂型，供用户选择
        gstrSQL = "Select 1 From 部门性质说明 " & _
                 " Where 工作性质 Like '中药%' And 部门ID = [1] "
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[检查部门性质]", lng库房id)
        
        mbln中药库房 = Not rsTmp.EOF
        
        gstrSQL = "Select Distinct J.编码,J.名称 " & _
                  "From 诊疗执行科室 A, 药品特性 B, 药品剂型 J " & _
                  "Where A.诊疗项目ID=B.药名ID And B.药品剂型=J.名称 And A.执行科室ID=[1] " & _
                  "Order by J.名称 "
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[提取该库房现在剂型]", lng库房id)
    Else
        gstrSQL = "Select 编码,名称 From 药品剂型 order by 名称 "
        Call zldatabase.OpenRecordset(rsTmp, gstrSQL, "提取所有药品剂型")
    End If
    
    With rsTmp
        lvw剂型.ListItems.Clear
        Do While Not .EOF
            lvw剂型.ListItems.Add , "K" & !编码, !名称, 1, 1
            .MoveNext
        Loop
        If .State = 1 Then .Close
        
        gstrSQL = "Select ID, 上级ID, 名称, 1 as 末级, decode(类型,1,'西成药',2,'中成药','中草药') as 材质, 类型 " & _
                  "From 诊疗分类目录 " & _
                  "Where 类型 in (1,2,3) " & _
                  "Start With 上级ID IS NULL Connect By Prior ID=上级ID Order by level,ID "
    End With
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "提取药品用途分类")
    With rsTmp
        If .EOF Then
            MsgBox "请初始化药品用途分类（药品用途分类）！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '将药品用途分类数据装入
        Do While Not .EOF
            Int末级 = IIf(!末级 = 1, 3, 2)
            If IsNull(!上级ID) Then
                Set nodTmp = tvw类别.Nodes.Add("Root" & !材质, 4, "K_" & !Id, !名称, Int末级, Int末级)
            Else
                Set nodTmp = tvw类别.Nodes.Add("K_" & !上级ID, 4, "K_" & !Id, !名称, Int末级, Int末级)
            End If
            nodTmp.Tag = !类型   '存放分类类型:1-西成药,2-中成药,3-中草药
            .MoveNext
        Loop
    End With

    With tvw类别
        .Nodes(1).Selected = True
        If .Nodes(1).Children <> 0 Then
            Int末级 = 1
            .Nodes(Int末级).Child.Selected = True
            .SelectedItem.Selected = True
        ElseIf .Nodes(2).Children <> 0 Then
            Int末级 = 2
            .Nodes(Int末级).Child.Selected = True
            .SelectedItem.Selected = True
        ElseIf .Nodes(3).Children <> 0 Then
            Int末级 = 3
            .Nodes(Int末级).Child.Selected = True
            .SelectedItem.Selected = True
        Else
            Int末级 = 0
            .Nodes(1).Selected = True
            .SelectedItem.Selected = True
        End If
        If Int末级 <> 0 Then .Nodes(Int末级).Expanded = True
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetVSFHead(ByVal vsfObject As VSFlexGrid, ByVal strHead As String)
'--------------------------------
'功能：初始化VSFlexGrid控件表格头
'参数：
'  vsfObject：目标控件；
'  strHead：表格头的初始化字串
'--------------------------------
    Dim arrCols As Variant, arrRows As Variant
    Dim i As Integer
    
    On Error GoTo errHandle
    
    arrRows = Split(strHead, "|")
    With vsfObject
        If .rows = 0 Then .rows = 1
        .Cols = UBound(arrRows) + 1
        For i = LBound(arrRows) To UBound(arrRows)
            If arrRows(i) <> "" Then
                arrCols = Split(arrRows(i), ",")
                '第1元素：Key值
                .ColKey(i) = arrCols(0)
                '第2元素：Caption值
                If arrCols(1) = "" Then
                    .TextMatrix(0, i) = arrCols(0)
                Else
                    .TextMatrix(0, i) = arrCols(1)
                End If
                '第3元素：列属性
                If arrCols(2) = "" Then
                    .ColData(i) = 3
                Else
                    .ColData(i) = Val(arrCols(2))
                End If
                '第4元素：宽度
                .ColWidth(i) = Val(arrCols(3))
                '第5元素：显示格式
                If UBound(arrCols) > 3 Then
                    If UCase(arrCols(4)) = "D" Then
                        .ColFormat(i) = "yyyy-mm-dd"
                        .ColAlignment(i) = flexAlignCenterCenter
                    ElseIf UCase(arrCols(4)) = "T" Then
                        .ColFormat(i) = "hh:mm:ss"
                        .ColAlignment(i) = flexAlignCenterCenter
                    ElseIf UCase(arrCols(4)) = "DT" Then
                        .ColFormat(i) = "yyyy-mm-dd hh:mm:ss"
                        .ColAlignment(i) = flexAlignCenterCenter
                    ElseIf UCase(arrCols(4)) = "N" Then
                        .ColAlignment(i) = flexAlignRightCenter
                    Else
                        .ColAlignment(i) = flexAlignLeftCenter
                    End If
                Else
                    .ColAlignment(i) = flexAlignLeftCenter
                End If
                '隐藏列
                If Val(arrCols(2)) = 1 Or Val(arrCols(2)) = 2 Then
                    .ColHidden(i) = True
                Else
                    .ColHidden(i) = False
                End If
                
            End If
        Next
        If .Cols > 0 Then .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
    End With
    Exit Sub
    
errHandle:
    Call ErrCenter
End Sub

Private Function GetVSFHead(ByVal vsfObject As VSFlexGrid) As String
'---------------------------------
'功能：获取VSF目标控件的表格头字串
'参数：vsfObject：目标控件
'返回：表格头字串
'---------------------------------
    Dim i As Integer
    Dim strHead As String, strCol As String
    
    With vsfObject
        strHead = ""
        For i = 0 To .Cols - 1
            '第1元素：Key
            strCol = .ColKey(i) & ","
            '第2元素：Caption
            If strCol = .TextMatrix(0, i) & "," Then
                strCol = strCol & ","
            Else
                strCol = strCol & .TextMatrix(0, i) & ","
            End If
            '第3元素：列属性
            If Val(.ColData(i)) = 3 Then
                If .ColHidden(i) Then
                    strCol = strCol & "2,"
                Else
                    strCol = strCol & ","
                End If
            Else
                If .ColHidden(i) = False And Val(.ColData(i)) = 2 Then
                    strCol = strCol & "3,"
                Else
                    strCol = strCol & .ColData(i) & ","
                End If
            End If
            '第4元素：列宽
            If Val(.ColWidth(i)) = 0 Then
                strCol = strCol & ","
            Else
                strCol = strCol & .ColWidth(i) & ","
            End If
            '第5元素：显示格式
            If Trim(.ColFormat(i)) = "" Then
                If .ColAlignment(i) = flexAlignRightCenter Then
                    strCol = strCol & "n"
                Else
                    strCol = Left(strCol, Len(strCol) - 1)
                End If
            Else
                If .ColFormat(i) = "yyyy-mm-dd" Then
                    strCol = strCol & "d"
                ElseIf .ColFormat(i) = "hh:mm:ss" Then
                    strCol = strCol & "t"
                ElseIf .ColFormat(i) = "yyyy-mm-dd hh:mm:ss" Then
                    strCol = strCol & "dt"
                End If
            End If
            '各列组合
            strHead = strHead & strCol & IIf(i = .Cols - 1, "", "|")
        Next
    End With
    GetVSFHead = strHead
End Function

Private Function GetVSFRow(ByVal vsfVal As VSFlexGrid) As Long
    Dim i As Long
    With vsfVal
        For i = 1 To vsfVal.rows - 1
            If vsfVal.RowHidden(i) = False Then
                GetVSFRow = i
                Exit Function
            End If
        Next
    End With
End Function

Private Sub SetColor()
    '设置表格颜色
    '如果是停用药品将其字体设置为红色
    Dim lngRow As Long
    
    With vsf规格
        If .rows > 1 Then
            For lngRow = 1 To .rows - 1
                If .TextMatrix(lngRow, .ColIndex("停用")) = "是" Then
                    .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = vbRed
                End If
            Next
        End If
    End With
End Sub

Private Sub tvw类别_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim strCols As String
    On Error GoTo errHandle
    If tvw类别.Tag <> Node.Key Then
        Me.MousePointer = vbHourglass
        strCols = GetVSFHead(vsf规格)
        mstrFilterClass = GetFilterClass(Node)
        
        If Node.Children = 0 And Node.Key Like "Root*" Then '如果是根节点且该节点下面无数据的话则没有数据
            mstrFilterClass = "用途分类id=99999999999999"
        Else
            mstrFilterClass = Left(mstrFilterClass, Len(mstrFilterClass) - 4)
        End If
        
        grsMaster.Filter = mstrFilterClass
        
        vsf规格.rows = 1
        Set vsf规格.DataSource = grsMaster
        '设置ColKey值
        SetColKey vsf规格
        '格式化VSF的列
        FormatCols strCols
        '根据条件过滤数据
        With fra形态
            If Node.Tag = "3" Then
                .Visible = True
            Else
                .Visible = False
            End If
        End With
        Call myFilter
        '设置颜色
        Call SetColor
    
        If Node.Tag = "3" Or Node.Tag = "Root7" Then '非中草药不显示中药形态和原产地
            vsf规格.ColHidden(vsf规格.ColIndex("中药形态")) = False
            vsf规格.ColHidden(vsf规格.ColIndex("原产地")) = False
            vsf规格.ColData(vsf规格.ColIndex("原产地")) = 3
            vsf批次.ColData(vsf批次.ColIndex("原产地")) = 3
            vsf规格.ColData(vsf规格.ColIndex("中药形态")) = 3
            vsf批次.ColWidth(vsf批次.ColIndex("原产地")) = 1000
        Else
            vsf规格.ColHidden(vsf规格.ColIndex("中药形态")) = True
            vsf规格.ColHidden(vsf规格.ColIndex("原产地")) = True
            vsf规格.ColData(vsf规格.ColIndex("原产地")) = 1
            vsf批次.ColData(vsf批次.ColIndex("原产地")) = 1
            vsf规格.ColData(vsf规格.ColIndex("中药形态")) = 1
            vsf批次.ColWidth(vsf批次.ColIndex("原产地")) = 0
        End If
        
        Call Form_Resize
        
        '刷新VSF批次
        If vsf规格.rows > 1 Then
            vsf规格.Row = 1
            ViewVSF批次 vsf规格
        Else
            vsf批次.rows = 1
            picSplit04_S.Visible = False
            vsf批次.Visible = False: vsf批次.TabStop = False
            Call Form_Resize
        End If
        
        If chkChoose.Value = 2 Then
            Call myFilter
            SetColor
            If GetVSFlexRows(vsf规格) <= 1 Then
                picSplit04_S.Visible = False
                vsf批次.Visible = False: vsf批次.TabStop = False
                Call Form_Resize
            Else
                vsf规格.Row = GetVSFRow(vsf规格)
            End If
        End If
        tvw类别.Tag = Node.Key
    End If
    Me.MousePointer = vbDefault
    Exit Sub

errHandle:
    Me.MousePointer = vbDefault
    Call ErrCenter
End Sub

Private Sub txtFilterFind_Change()
    Dim strCols As String, strFind As String
    Dim i As Integer
    Dim rstemp As ADODB.Recordset
    
    strCols = GetVSFHead(vsf规格)
    strFind = Trim(txtFilterFind.Text)
    txtFilterFind.Tag = ""
    
    mstrFilterClass = GetFilterSimpleCode(strFind)
    err.Clear: On Error GoTo errHandle
    grsMasterInput.Filter = mstrFilterClass
    err.Clear: On Error GoTo 0
    vsf规格.rows = 1
    Set vsf规格.DataSource = grsMasterInput
    '设置ColKey值
    SetColKey vsf规格
    '格式化VSF的列
    FormatCols strCols
    '设置颜色
    Call SetColor
    '刷新批次
    ViewVSF批次 vsf规格
    '中药不隐藏"原产地"、"中药形态"列
    Call HiddenColumns
    
    If mblnLoad = True Then
        If mbytStyle = 1 Then
            '录入模式才检查
            gstrSQL = "Select a.编码 as 药品编码,a.名称 as 药品名称, b.名称 As 输入名称, b.简码, b.数字简码, b.五笔码" & vbNewLine & _
                            "From 收费项目目录 A," & vbNewLine & _
                            "     (Select 收费细目id, Max(Decode(码类, '3', 简码, Null)) 数字简码, Max(Decode(码类, '1', 简码, Null)) 简码," & vbNewLine & _
                            "              Max(Decode(码类, '2', 简码, Null)) 五笔码, 名称" & vbNewLine & _
                            "       From 收费项目别名" & vbNewLine & _
                            "       Where 码类 In (1, 2, 3) And 性质 In (1, 3, 9)" & vbNewLine & _
                            "       Group By 收费细目id, 名称) B" & vbNewLine & _
                            "Where a.Id = b.收费细目id And a.类别 In ('5', '6', '7') And (a.站点 = '0' Or a.站点 Is Null)"
    
            Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "查询简码")
            
            If mbyt编辑模式 = 1 Then
                '入库
                rstemp.Filter = mstrFilterClass
                If grsMasterInput.RecordCount = 0 And rstemp.RecordCount = 0 Then
                    MsgBox "无此药品！", vbInformation, gstrSysName
                End If
            Else
                '出库
                rstemp.Filter = mstrFilterClass
                If grsMasterInput.RecordCount = 0 And rstemp.RecordCount > 0 And mint按批次出库 = 1 Then
                    MsgBox "此药品无库存了！", vbInformation, gstrSysName
                ElseIf grsMasterInput.RecordCount = 0 And rstemp.RecordCount = 0 Then
                    MsgBox "无此药品！", vbInformation, gstrSysName
                End If
            End If
        End If
    End If
    Exit Sub
    
errHandle:
    MsgBox "文本包含非法字符！", vbInformation, gstrSysName
    txtFilterFind.Text = "": txtFilterFind.Tag = "1"
    If txtFilterFind.Enabled And txtFilterFind.Visible Then txtFilterFind.SetFocus
End Sub

Private Sub txtFilterFind_GotFocus()
    txtFilterFind.SelStart = 0
    txtFilterFind.SelLength = Len(txtFilterFind.Text)
End Sub

Private Sub txtFilterFind_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = vbKeyReturn Then zlcommfun.PressKey vbKeyTab
End Sub

Private Sub txtFilterFind_KeyPress(KeyAscii As Integer)
    If InStr("`~!@#$%^&*()_-+={[}]|\:;""'<,>.?/", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    ElseIf KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    Else
    End If
End Sub

Private Sub vsf规格_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <> NewRow Then
        ViewVSF批次 vsf规格
    End If
End Sub

Private Function SetColPropStr(ByVal strCol As String, ByVal intProp As Integer, ByVal strParam As String) As String
'----------------------------------
'功能：改变指定的列属性值
'参数：
'  strCol：列的所有属性字符串
'  intProp：列的属性列号
'  strParam：要改变列的目标值
'返回：新的列属性字符串
'----------------------------------
    Dim arrElement As Variant
    Dim i As Integer, n As Integer
    Dim strTmp As String, strReturn As String
    
    arrElement = Split(strCol, ",")
    n = UBound(arrElement)
    For i = 0 To 4
        If intProp > i Then
            strTmp = "," & arrElement(i)
        ElseIf intProp = i Then
            strTmp = "," & strParam
        Else
            If i > n Then
                strTmp = ""
            Else
                strTmp = "," & arrElement(i)
            End If
        End If
        strReturn = strReturn & strTmp
    Next
    SetColPropStr = Right(strReturn, Len(strReturn) - 1)
End Function

Private Sub ParamsColsHead()
'----------------------------------
'功能：根据参数设置，对应的特殊处理
'----------------------------------
    Dim intBegin As Integer, intLen As Integer
    Dim strColHead As String
    Dim arrCols规格 As Variant, arrCols批次 As Variant, arrCols选定 As Variant
    Dim i As Integer
    Dim strTmp As String

    'VSF列头更新
    SyncColumns MCON_规格, mstr规格
    SyncColumns MCON_批次, mstr批次
    SyncColumns MCON_选定, mstr选定

    On Error GoTo errHandle
    
    arrCols规格 = Split(mstr规格, "|")
    arrCols批次 = Split(mstr批次, "|")
    arrCols选定 = Split(mstr选定, "|")
    
    strTmp = ""
    For i = LBound(arrCols规格) To UBound(arrCols规格)
        If InStr(";" & arrCols规格(i) & ";", ";通用名称,") > 0 Then
            If mbytStyle = 1 Then
                '录入药品名称   0：匹配显示； 1：同时显示通用名和商品名
                strTmp = strTmp & SetColPropStr(arrCols规格(i), enmColProperty.cpDisplay, IIf(gint输入药品显示 = 0, "1", "0")) & "|"
            Else
                '显示药品名称   0：显示通用名； 1：显示商品名； 2：同时显示通用名和商品名
                strTmp = strTmp & SetColPropStr(arrCols规格(i), enmColProperty.cpDisplay, IIf(gint药品名称显示 = 1, "1", "0")) & "|"
            End If
        ElseIf InStr(";" & arrCols规格(i) & ";", ";商品名,") > 0 Then
            If mbytStyle = 1 Then
                '录入药品名称   0：匹配显示； 1：同时显示通用名和商品名
                strTmp = strTmp & SetColPropStr(arrCols规格(i), enmColProperty.cpDisplay, IIf(gint输入药品显示 = 0, "1", "0")) & "|"
            Else
                '显示药品名称   0：显示通用名； 1：显示商品名； 2：同时显示通用名和商品名
                strTmp = strTmp & SetColPropStr(arrCols规格(i), enmColProperty.cpDisplay, IIf(gint药品名称显示 = 0, "1", "0")) & "|"
            End If
        ElseIf InStr(";" & arrCols规格(i) & ";", ";药品名称,") > 0 Then
            If mbytStyle = 1 Then
                '录入药品名称   0：匹配显示； 1：同时显示通用名和商品名
                strTmp = strTmp & SetColPropStr(arrCols规格(i), enmColProperty.cpDisplay, IIf(gint输入药品显示 = 0, "0", "1")) & "|"
            Else
                strTmp = strTmp & arrCols规格(i) & "|"
            End If
        ElseIf InStr(";" & arrCols规格(i) & ";", ";售价单位,") > 0 Then
            strTmp = strTmp & SetColPropStr(arrCols规格(i), enmColProperty.cpDisplay, IIf(mintUnit = 1, "0", "1")) & "|"
        ElseIf InStr(";" & arrCols规格(i) & ";", ";门诊单位,") > 0 Then
            strTmp = strTmp & SetColPropStr(arrCols规格(i), enmColProperty.cpDisplay, IIf(mintUnit = 2, "0", "1")) & "|"
        ElseIf InStr(";" & arrCols规格(i) & ";", ";住院单位,") > 0 Then
            strTmp = strTmp & SetColPropStr(arrCols规格(i), enmColProperty.cpDisplay, IIf(mintUnit = 3, "0", "1")) & "|"
        ElseIf InStr(";" & arrCols规格(i) & ";", ";药库单位,") > 0 Then
            strTmp = strTmp & SetColPropStr(arrCols规格(i), enmColProperty.cpDisplay, IIf(mintUnit = 4, "0", "1")) & "|"
        ElseIf InStr(";" & arrCols规格(i) & ";", ";售价包装,") > 0 Then
            strTmp = strTmp & SetColPropStr(arrCols规格(i), enmColProperty.cpDisplay, IIf(mintUnit = 1, "0", "1")) & "|"
        ElseIf InStr(";" & arrCols规格(i) & ";", ";门诊包装,") > 0 Then
            strTmp = strTmp & SetColPropStr(arrCols规格(i), enmColProperty.cpDisplay, IIf(mintUnit = 2, "0", "1")) & "|"
        ElseIf InStr(";" & arrCols规格(i) & ";", ";住院包装,") > 0 Then
            strTmp = strTmp & SetColPropStr(arrCols规格(i), enmColProperty.cpDisplay, IIf(mintUnit = 3, "0", "1")) & "|"
        ElseIf InStr(";" & arrCols规格(i) & ";", ";药库包装,") > 0 Then
            strTmp = strTmp & SetColPropStr(arrCols规格(i), enmColProperty.cpDisplay, IIf(mintUnit = 4, "0", "1")) & "|"
        ElseIf InStr(";" & arrCols规格(i) & ";", ";上次采购价,") > 0 Then
            strTmp = strTmp & SetColPropStr(arrCols规格(i), enmColProperty.cpDisplay, IIf(mblnCostView = False, "1", "0")) & "|"
        ElseIf InStr(";" & arrCols规格(i) & ";", ";库存差价,") > 0 Then
            strTmp = strTmp & SetColPropStr(arrCols规格(i), enmColProperty.cpDisplay, IIf(mblnCostView = False, "1", "0")) & "|"
        ElseIf InStr(";" & arrCols规格(i) & ";", ";指导批发价,") > 0 Then
            strTmp = strTmp & SetColPropStr(arrCols规格(i), enmColProperty.cpDisplay, IIf(mblnCostView = False, "1", "0")) & "|"
        ElseIf InStr(";" & arrCols规格(i) & ";", ";合同单位,") > 0 Then
            '1300：药品外购入库； 1330：药品计划管理
            strTmp = strTmp & SetColPropStr(arrCols规格(i), enmColProperty.cpDisplay, IIf(glngModul = 1300 Or glngModul = 1330, "0", "1")) & "|"
        Else
            strTmp = strTmp & arrCols规格(i) & "|"
        End If
    Next
    mstr规格 = Left(strTmp, Len(strTmp) - 1)
    
    '批次
    strTmp = ""
    For i = LBound(arrCols批次) To UBound(arrCols批次)
        If InStr(";" & arrCols批次(i) & ";", ";有效期,") > 0 Then
            '效期     0：失效期；  1：有效期
            If gtype_UserSysParms.P149_效期显示方式 = 0 Then
                strTmp = strTmp & SetColPropStr(arrCols批次(i), enmColProperty.cpCaption, "失效期") & "|"
            Else
                strTmp = strTmp & SetColPropStr(arrCols批次(i), enmColProperty.cpCaption, "有效期至") & "|"
            End If
        ElseIf InStr(";" & arrCols批次(i) & ";", ";入库日期,") > 0 Then
            '1304：药品移库管理； 1343：药品申领管理
            strTmp = strTmp & SetColPropStr(arrCols批次(i), enmColProperty.cpDisplay, IIf(glngModul = 1304 Or glngModul = 1343, "0", "1")) & "|"
        ElseIf InStr(";" & arrCols批次(i) & ";", ";成本价,") > 0 Then
            strTmp = strTmp & SetColPropStr(arrCols批次(i), enmColProperty.cpDisplay, IIf(mblnCostView = False, "1", "0")) & "|"
        ElseIf InStr(";" & arrCols批次(i) & ";", ";库存差价,") > 0 Then
            strTmp = strTmp & SetColPropStr(arrCols批次(i), enmColProperty.cpDisplay, IIf(mblnCostView = False, "1", "0")) & "|"
        Else
            strTmp = strTmp & arrCols批次(i) & "|"
        End If
    Next
    mstr批次 = Left(strTmp, Len(strTmp) - 1)
    
    '选定
    strTmp = ""
    For i = LBound(arrCols选定) To UBound(arrCols选定)
        If InStr(";" & arrCols选定(i) & ";", ";通用名称,") > 0 Then
            If mbytStyle = 1 Then
                '录入药品名称   0：匹配显示； 1：同时显示通用名和商品名
                strTmp = strTmp & SetColPropStr(arrCols选定(i), enmColProperty.cpDisplay, IIf(gint输入药品显示 = 0, "1", "0")) & "|"
            Else
                '显示药品名称   0：显示通用名； 1：显示商品名； 2：同时显示通用名和商品名
                strTmp = strTmp & SetColPropStr(arrCols选定(i), enmColProperty.cpDisplay, IIf(gint药品名称显示 = 1, "1", "0")) & "|"
            End If
        ElseIf InStr(";" & arrCols选定(i) & ";", ";商品名,") > 0 Then
            If mbytStyle = 1 Then
                '录入药品名称   0：匹配显示； 1：同时显示通用名和商品名
                strTmp = strTmp & SetColPropStr(arrCols选定(i), enmColProperty.cpDisplay, IIf(gint输入药品显示 = 0, "1", "0")) & "|"
            Else
                '显示药品名称   0：显示通用名； 1：显示商品名； 2：同时显示通用名和商品名
                strTmp = strTmp & SetColPropStr(arrCols选定(i), enmColProperty.cpDisplay, IIf(gint药品名称显示 = 0, "1", "0")) & "|"
            End If
        ElseIf InStr(";" & arrCols选定(i) & ";", ";药品名称,") > 0 Then
            If mbytStyle = 1 Then
                '录入药品名称   0：匹配显示； 1：同时显示通用名和商品名
                strTmp = strTmp & SetColPropStr(arrCols选定(i), enmColProperty.cpDisplay, IIf(gint输入药品显示 = 0, "0", "1")) & "|"
            Else
                strTmp = strTmp & arrCols选定(i) & "|"
            End If
        ElseIf InStr(";" & arrCols选定(i) & ";", ";上次采购价,") > 0 Then
            strTmp = strTmp & SetColPropStr(arrCols选定(i), enmColProperty.cpDisplay, IIf(mblnCostView = False, "1", "0")) & "|"
        ElseIf InStr(";" & arrCols选定(i) & ";", ";库存差价,") > 0 Then
            strTmp = strTmp & SetColPropStr(arrCols选定(i), enmColProperty.cpDisplay, IIf(mblnCostView = False, "1", "0")) & "|"
        ElseIf InStr(";" & arrCols选定(i) & ";", ";指导批发价,") > 0 Then
            strTmp = strTmp & SetColPropStr(arrCols选定(i), enmColProperty.cpDisplay, IIf(mblnCostView = False, "1", "0")) & "|"
        ElseIf InStr(";" & arrCols选定(i) & ";", ";合同单位,") > 0 Then
            '1300：药品外购入库； 1330：药品计划管理
            strTmp = strTmp & SetColPropStr(arrCols选定(i), enmColProperty.cpDisplay, IIf(glngModul = 1300 Or glngModul = 1330, "0", "1")) & "|"
        ElseIf InStr(";" & arrCols选定(i) & ";", ";有效期,") > 0 Then
            '效期     0：失效期；  1：有效期
            If gtype_UserSysParms.P149_效期显示方式 = 0 Then
                strTmp = strTmp & SetColPropStr(arrCols选定(i), enmColProperty.cpCaption, "失效期") & "|"
            Else
                strTmp = strTmp & SetColPropStr(arrCols选定(i), enmColProperty.cpCaption, "有效期至") & "|"
            End If
        Else
            strTmp = strTmp & arrCols选定(i) & "|"
        End If
    Next
    mstr选定 = Left(strTmp, Len(strTmp) - 1)
    
    Exit Sub
    
errHandle:
    Call ErrCenter
End Sub

Private Function GetFilterSimpleCode(ByVal strSimpleCode As String) As String
'--------------------------------------------
'功能：得到简码匹配的过滤条件
'参数：strSimpleCode：录入的简码
'返回：过滤条件字符串
'--------------------------------------------
    Dim strReturn As String, strTemp As String
    
    If strSimpleCode = "" Then Exit Function
    
    If gint简码方式 = 1 Then
        '五笔码
        strTemp = "五笔码"
    Else
        '拼音码
        strTemp = "简码"
    End If
    
    If IsNumeric(strSimpleCode) Then
        '纯数字
        strReturn = "药品编码 like '" & mstrMatch & strSimpleCode & "%'" & _
                    " Or 数字简码 like '" & mstrMatch & strSimpleCode & "%'" & _
                    " or 药品名称 like '" & mstrMatch & strSimpleCode & "%'"
    ElseIf zlStr.IsCharAlpha(strSimpleCode) Then
        '纯字母
        strReturn = strTemp & " like '" & mstrMatch & strSimpleCode & "%'"
    ElseIf zlStr.IsCharChinese(strSimpleCode) Then
        '纯汉字
        strReturn = "药品名称 like '" & mstrMatch & strSimpleCode & "%'"
    Else
        strReturn = "药品编码 like '" & mstrMatch & strSimpleCode & "%'" & _
                    " Or " & strTemp & " like '" & mstrMatch & strSimpleCode & "%'" & _
                    " Or 药品名称 like '" & mstrMatch & strSimpleCode & "%'"
    End If
    
    GetFilterSimpleCode = strReturn
End Function

Private Function GetFilterClass(ByVal objNode As Node) As String
'--------------------------------------------
'功能：得到当前节点下所有的子孙节点的过滤条件
'参数：objNode：当前节点对象
'返回：过滤条件字符串
'--------------------------------------------
    Dim i, n As Integer
    Dim strReturn As String
    Dim objTmp As Node
    Dim strsql As String
    Dim bln库存 As Boolean

    n = objNode.Children
        
    If Left(objNode.Key, 2) = "K_" Then
        strReturn = strReturn & "用途分类id=" & Mid(objNode.Key, 3) & " or "
    End If
    If n > 0 Then
        Set objTmp = objNode.Child
        strReturn = strReturn & GetFilterClass(objTmp)
        For i = 2 To n
            Set objTmp = objTmp.Next
            strReturn = strReturn & GetFilterClass(objTmp)
        Next
    End If
    
    GetFilterClass = strReturn
End Function

Private Sub FillVSF(ByVal rsVal As ADODB.Recordset, ByVal vsfVal As VSFlexGrid)
'------------------------------
'功能：为VSF控件填充数据
'参数：
'  rsVal：填充的数据集
'------------------------------
    Dim i, j As Long
    Dim strData As String
    
    With rsVal
        vsfVal.rows = 1
        If .RecordCount > 0 Then
            .MoveFirst
        End If
        vsfVal.Redraw = False
        vsfVal.rows = .RecordCount + 1
        For i = 1 To .RecordCount
            For j = 0 To .Fields.count - 1
                If vsfVal.ColIndex(.Fields(j).Name) > -1 Then
                    vsfVal.TextMatrix(i, vsfVal.ColIndex(.Fields(j).Name)) = FieldValueDisp(.Fields, j, vsfVal.Name)
                Else
                    'Debug.Print .Fields(j).Name & "(无)"
                End If
            Next
            .MoveNext
        Next
        vsfVal.Redraw = True
    End With
    If glngModul = 1305 Then
        With vsf批次
            If InStr(1, gstrprivs, "显示对方库存") = 0 Then
                .ColData(.ColIndex("库存数量")) = 1
                .ColData(.ColIndex("库存金额")) = 1
                .ColData(.ColIndex("库存差价")) = 1
            Else
                .ColData(.ColIndex("库存数量")) = IIf(.ColData(.ColIndex("库存数量")) = 2, 2, 3)
                .ColData(.ColIndex("库存金额")) = IIf(.ColData(.ColIndex("库存金额")) = 2, 2, 3)
                .ColData(.ColIndex("库存差价")) = IIf(.ColData(.ColIndex("库存差价")) = 2, 2, 3)
            End If

        
            If InStr(1, gstrprivs, "显示对方库存") = 0 Then
                For i = 1 To .rows - 1
                    If Val(.TextMatrix(i, .ColIndex("库存数量"))) > 0 Then
                        .TextMatrix(i, .ColIndex("库存数量")) = "有"
                    Else
                        .TextMatrix(i, .ColIndex("库存数量")) = "无"
                    End If
                Next
                .ColData(.ColIndex("库存金额")) = 1
                .ColData(.ColIndex("库存差价")) = 1
                .ColHidden(.ColIndex("库存金额")) = True
                .ColHidden(.ColIndex("库存差价")) = True
            Else
'                .ColData(.ColIndex("库存数量")) = IIf(.ColData(.ColIndex("库存数量")) = 2, 2, 3)
                .ColData(.ColIndex("库存金额")) = IIf(.ColData(.ColIndex("库存金额")) = 2, 2, 3)
                .ColData(.ColIndex("库存差价")) = IIf(.ColData(.ColIndex("库存差价")) = 2, 2, 3)
                .ColHidden(.ColIndex("库存金额")) = IIf(.ColData(.ColIndex("库存金额")) = 2, True, False)
                .ColHidden(.ColIndex("库存差价")) = IIf(.ColData(.ColIndex("库存差价")) = 2, True, False)
            End If
        End With
    End If
    
    If glngModul = 1343 Then '申领
        With vsf批次
            If InStr(1, gstrprivs, "显示对方库存") = 0 Then
                For i = 1 To .rows - 1
                    If Val(.TextMatrix(i, .ColIndex("可用数量"))) > 0 Then
                        .TextMatrix(i, .ColIndex("可用数量")) = "有"
                    Else
                        .TextMatrix(i, .ColIndex("可用数量")) = "无"
                    End If
                Next
            End If
        End With
    End If
    
End Sub

Private Function FieldValueDisp(ByVal objFields As ADODB.Fields, ByVal intCol As Integer, ByVal strVSFName As String) As String
'--------------------------------
'功能：根据参数设置，调整显示数据
'参数：
'  objFields：列集合
'  intCol：列序号
'  strVSFName：VSF控件Name值
'返回：调整后的数据(字符型)
'--------------------------------
    Dim strReturn As String
    Dim dblUnit As Double
    
    Select Case mintUnit
    Case mconint门诊单位
        dblUnit = zlStr.Nvl(objFields("门诊包装").Value, 0)
    Case mconint住院单位
        dblUnit = zlStr.Nvl(objFields("住院包装").Value, 0)
    Case mconint药库单位
        dblUnit = zlStr.Nvl(objFields("药库包装").Value, 0)
    Case Else
        dblUnit = 1
    End Select
    
    If objFields(intCol).Name = "上次采购价" Or objFields(intCol).Name = "成本价" Then
        '上次采购价、售价：vsf规格控件的列； 成本价：vsf批次控件的列
        If strVSFName <> "vsf规格" Then
            strReturn = Format(zlStr.Nvl(objFields(intCol).Value, 0) * dblUnit, mStr成本价)
        Else
            strReturn = Format(zlStr.Nvl(objFields(intCol).Value, 0), mStr成本价)
        End If
    ElseIf objFields(intCol).Name = "售价" Then
        If strVSFName <> "vsf规格" Then
            strReturn = Format(zlStr.Nvl(objFields(intCol).Value, 0) * dblUnit, mStr单价)
        Else
            strReturn = Format(zlStr.Nvl(objFields(intCol).Value, 0), mStr单价)
        End If
    ElseIf objFields(intCol).Name = "可用数量" Or objFields(intCol).Name = "库存数量" Or objFields(intCol).Name = "实际数量" Then
        If strVSFName <> "vsf规格" Then
            strReturn = Format(zlStr.Nvl(objFields(intCol).Value, 0) / dblUnit, mStr数量)
        Else
            strReturn = Format(zlStr.Nvl(objFields(intCol).Value, 0), mStr数量)
        End If
    ElseIf objFields(intCol).Name = "库存金额" Or objFields(intCol).Name = "库存差价" Then
        strReturn = Format(zlStr.Nvl(objFields(intCol).Value, 0), mStr金额)
    Else
        strReturn = zlStr.Nvl(objFields(intCol).Value)
    End If
    
    '盘点记录单“查看盘点单库存”参数
    If glngModul <> 1343 Then '非申领
        If mblnStore = False And (objFields(intCol).Name = "售价" Or objFields(intCol).Name = "成本价" _
            Or objFields(intCol).Name = "可用数量" Or objFields(intCol).Name = "库存数量" Or objFields(intCol).Name = "库存金额" _
            Or objFields(intCol).Name = "库存差价") Then
            strReturn = ""
        End If
    Else '申领
        If mblnStore = False And (objFields(intCol).Name = "售价" Or objFields(intCol).Name = "成本价" _
            Or objFields(intCol).Name = "库存数量" Or objFields(intCol).Name = "库存金额" _
            Or objFields(intCol).Name = "库存差价") Then
            strReturn = ""
        End If
    End If
    
    FieldValueDisp = Trim(strReturn)
End Function

Private Sub ViewVSF批次(ByVal vsfVal As VSFlexGrid)
    Dim strFilter As String
    Dim blnVisible As Boolean
    Dim int分批 As Integer
    Dim lngRow As Long
    
    If mbyt编辑模式 <> 2 Then Exit Sub
    blnVisible = vsf批次.Visible
    If grsSlave.State = adStateClosed Then
        picSplit04_S.Visible = False
        vsf批次.Visible = False: vsf批次.TabStop = False
        If blnVisible <> vsf批次.Visible Then Call Form_Resize
        Exit Sub
    End If
    
    If GetVSFlexRows(vsfVal) <= 1 Then Exit Sub
    
    strFilter = "药品ID=" & Val(vsfVal.TextMatrix(vsfVal.Row, vsfVal.ColIndex("药品ID")))
    grsSlave.Filter = strFilter
    FillVSF grsSlave, vsf批次
    
    int分批 = Get库房分批()
    With vsfVal
        '如果该药品不分批
        If Not ((int分批 = 3 And mbyt库房性质 <> 3) Or (int分批 = 1 And mbyt库房性质 = 1) Or (int分批 = 2 And mbyt库房性质 = 2)) Then
            picSplit04_S.Visible = False
            vsf批次.Visible = False: vsf批次.TabStop = False
        Else
            '移库可以根据按批次复选框控制
            If UCase(mfrmMain.Name) = UCase("frmTransferCard") Or UCase(mfrmMain.Name) = UCase("frmRequestDrugCard") Or UCase(mfrmMain.Name) = UCase("frmDrawCard") Then
                If mint按批次出库 = 1 Then
                    picSplit04_S.Visible = True
                    vsf批次.Visible = True: vsf批次.TabStop = True
                Else
                    picSplit04_S.Visible = False
                    vsf批次.Visible = False: vsf批次.TabStop = False
                End If
            ElseIf UCase(mfrmMain.Name) = UCase("frmcheckCoursecard") Or UCase(mfrmMain.Name) = UCase("frmCheckCard") Then
                picSplit04_S.Visible = True
                vsf批次.Visible = True: vsf批次.TabStop = True
                '增加空批次记录
                If mbln空批次 Then
                    With vsf批次
                        .rows = .rows + 1
                        lngRow = .rows - 1
                        .TextMatrix(lngRow, .ColIndex("RID")) = "1"
                        .TextMatrix(lngRow, .ColIndex("库房")) = mstr空批次库房
                        .TextMatrix(lngRow, .ColIndex("批次")) = "-1"
                        .TextMatrix(lngRow, .ColIndex("批号")) = "新增批次药品"
                        .TextMatrix(lngRow, .ColIndex("有效期")) = zldatabase.Currentdate
                        .RowPosition(lngRow) = 1
                    End With
                End If
            Else
                If grsSlave.RecordCount > 0 Then
                    picSplit04_S.Visible = True
                    vsf批次.Visible = True: vsf批次.TabStop = True
                Else
                    '增加空批次记录，只有盘点有效
                    If mbln空批次 Then
                        picSplit04_S.Visible = True
                        vsf批次.Visible = True: vsf批次.TabStop = True
                        With vsf批次
                            .rows = .rows + 1
                            lngRow = .rows - 1
                            .TextMatrix(lngRow, .ColIndex("RID")) = "1"
                            .TextMatrix(lngRow, .ColIndex("库房")) = mstr空批次库房
                            .TextMatrix(lngRow, .ColIndex("批次")) = "-1"
                            .TextMatrix(lngRow, .ColIndex("批号")) = "新增批次药品"
                            .TextMatrix(lngRow, .ColIndex("有效期")) = Sys.Currentdate
                            .RowPosition(lngRow) = 1
                        End With
                    Else
                        picSplit04_S.Visible = False
                        vsf批次.Visible = False: vsf批次.TabStop = False
                    End If
                End If
            End If
        End If
    End With
    
    '刷新
    If blnVisible <> vsf批次.Visible Then Call Form_Resize
    If mbln空批次 Then
        If vsf批次.rows > 2 Then
            vsf批次.Row = 2
        ElseIf vsf批次.rows > 1 Then
            vsf批次.Row = 1
        End If
    Else
        If vsf批次.rows > 1 Then vsf批次.Row = 1
    End If
End Sub




Private Sub SyncColumns(ByVal strInit As String, ByRef strRegister As String)
'-----------------------------------------------
'功能：VSF列头与注册表保存的列头名、数量保持一致
'参数：
'  strInit：列头的初始值
'  strRegister：保存在注册表的列头，并更新它
'-----------------------------------------------
    Dim i As Integer, j As Integer, intOrder As Integer
    Dim arrInit As Variant, arrRegister As Variant
    Dim blnFind As Boolean
    Dim strTmp As String
    Dim bytInit As Byte, bytReg As Byte
    Dim rstemp As ADODB.Recordset
    
    Set rstemp = New ADODB.Recordset
    With rstemp
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Fields.Append "SN", adInteger
        .Fields.Append "VAL", adVarChar, 100
        .Open
    End With
    
    arrInit = Split(strInit, "|")
    arrRegister = Split(strRegister, "|")
    For i = LBound(arrInit) To UBound(arrInit)
        blnFind = False
        For j = LBound(arrRegister) To UBound(arrRegister)
            If Split(arrInit(i), ",")(0) = Split(arrRegister(j), ",")(0) Then
                rstemp.AddNew
                rstemp!SN = j
                '比较Key值
                '列头程序常量的属性改变，注意"0"与Val()的0是不同的，Val()=0是属性3，即默认显示属性
                If Split(arrRegister(j), ",")(2) = "0" Or Split(arrRegister(j), ",")(2) = "1" Then bytReg = 1 Else bytReg = 0
                If Split(arrInit(i), ",")(2) = "0" Or Split(arrInit(i), ",")(2) = "1" Then bytInit = 1 Else bytInit = 0
                If bytInit = bytReg Then
                    rstemp!Val = arrRegister(j)
                Else
                    strTmp = arrRegister(j)
                    strTmp = SetColPropStr(strTmp, enmColProperty.cpDisplay, Split(arrInit(i), ",")(2))
                    rstemp!Val = strTmp
                End If
                '比较Format值
                If UBound(Split(arrInit(i), ",")) >= 4 Then
                    If UBound(Split(arrRegister(j), ",")) >= 3 Then
                        strTmp = rstemp!Val
                        If UBound(Split(arrRegister(j), ",")) >= 4 Then
                            strTmp = SetColPropStr(strTmp, enmColProperty.cpFormat, Split(arrInit(i), ",")(4))
                        Else
                            strTmp = strTmp & "," & Split(arrInit(i), ",")(4)
                        End If
                        rstemp!Val = strTmp
                    End If
                End If
                rstemp.Update
                blnFind = True
                Exit For
            End If
        Next
        If blnFind = False Then
            '新增的列
            rstemp.AddNew
            rstemp!SN = i
            rstemp!Val = arrInit(i)
            rstemp.Update
        End If
    Next
    
    '排序
    rstemp.Sort = "SN"
    strTmp = ""
    rstemp.MoveFirst
    Do While Not rstemp.EOF
        strTmp = strTmp & rstemp!Val & "|"
        rstemp.MoveNext
    Loop
    rstemp.Close
    strRegister = Left(strTmp, Len(strTmp) - 1)
End Sub

Private Function VerifyFilterStr(ByVal strFilter As String) As String
'----------------------------------------------------
'功能：审核录入简码过滤字符串有无特殊字符，并过滤特殊
'----------------------------------------------------
    Dim i As Integer
    Dim strTmp As String
    
    If Len(strFilter) < 1 Then Exit Function
                                                                                                                                                                                                                                                               
    For i = 1 To Len(strFilter)
        If InStr("`~!@#$%^&*()_-+={[}]|\:;""'<,>.?/", Mid(strFilter, i, 1)) = 0 Then
            strTmp = strTmp & Mid(strFilter, i, 1)
        End If
    Next
    VerifyFilterStr = strTmp
End Function

Private Function CheckData() As Boolean
    Dim DblCurStock As Double       '当前库存数
    '检测是否允许选择
    CheckData = False
    
'    If BlnSelect = False Then Exit Function
    
    'lng供应商ID不为零，表示退货，无库存时不准继续
    If mlng供应商 <> 0 Then
        If vsf批次.Visible Then
            If Val(vsf批次.TextMatrix(vsf批次.Row, vsf批次.ColIndex("上次供应商ID"))) <> 0 _
                And mlng供应商 <> Val(vsf批次.TextMatrix(vsf批次.Row, vsf批次.ColIndex("上次供应商ID"))) Then
                MsgBox "你选择的退货商不是该药品的供应商，不能继续操作！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If

    '如果源库房与目库房为空，则表明是药品目录自己在进行常规设置，不判断
    If (mlng来源库房 = 0 And mlng目标库房 = 0) Then
        CheckData = True
        Exit Function
    End If
    
    '如果是盘点单调用药品选择器，则不需判断，直接退出
    'If bln盘点单 Then
    If glngModul = 1307 Or glngModul = 1303 Then   '药品盘点管理、药品库存差价调整
        CheckData = True
        Exit Function
    End If
    
    If vsf批次.Visible Then
        If Trim(vsf批次.TextMatrix(vsf批次.Row, vsf批次.ColIndex("可用数量"))) = "有" Then
            CheckData = True
            Exit Function
        End If
        
        DblCurStock = Val(vsf批次.TextMatrix(vsf批次.Row, vsf批次.ColIndex("可用数量")))
    Else
        DblCurStock = Val(vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("可用数量")))
    End If
    
    If DblCurStock > 0 Or mint按批次出库 = 0 Then
        '不分批的不检查库存(申领/移库/领用)
        CheckData = True
        Exit Function
    Else
        Select Case mintStockCheck
        Case 1
            If MsgBox("该药品已经没有库存，是否继续！", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        Case 2
            MsgBox "该药品已经没有库存，不能继续操作！", vbInformation, gstrSysName
            Exit Function
        End Select
    End If
        
    CheckData = True
End Function

Private Sub vsf规格_Click()
    If pic选定区.Tag = "展开" Then picUpDown01_Click
End Sub

Private Sub vsf规格_DblClick()
    Dim int分批 As Integer
    
    mblnOK = False
    
    If GetVSFlexRows(vsf规格) <= 1 Then Exit Sub
    
    If glngModul = 1305 Then
        With vsf规格
            If .TextMatrix(.Row, .ColIndex("领用标志")) = "0" Or .TextMatrix(.Row, .ColIndex("领用标志")) = "" Then
                MsgBox "该部门不允许领用该药品，请到药品储备限额中设置！", vbInformation, gstrSysName
                Exit Sub
            End If
            
            '零差价管理：选定药品不满足零差价要求，则不能进行出库业务
            If gtype_UserSysParms.P275_零差价管理模式 = 2 And IsPriceAdjustMod(Val(vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("药品ID")))) = True Then
                If CheckPriceAdjust(Val(vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("药品ID"))), mlng来源库房, -1) = False Then
                    MsgBox "该药品启用零差价管理模式，但成本价和售价不一致，不能开展业务。请先调整价格！", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        End With
    End If
    If vsf批次.Visible Then
        '库房分批
        If GetVSFlexRows(vsf批次) > 2 Then Exit Sub
        Call vsf批次_DblClick
        Exit Sub
    Else
        If glngModul = 1304 Or glngModul = 1305 Or glngModul = 1306 Or glngModul = 1307 Or glngModul = 1343 Then
            '零差价管理：选定药品不满足零差价要求，则不能进行出库业务
            If gtype_UserSysParms.P275_零差价管理模式 = 2 And IsPriceAdjustMod(Val(vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("药品ID")))) = True Then
                If CheckPriceAdjust(Val(vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("药品ID"))), mlng来源库房, -1) = False Then
                    MsgBox "该药品启用零差价管理模式，但成本价和售价不一致，不能开展业务。请先调整价格！", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        End If
                
        If UCase(mfrmMain.Name) <> UCase("frmOtherOutputCard") And UCase(mfrmMain.Name) <> UCase("frmTransferCard") And UCase(mfrmMain.Name) <> UCase("frmPurchaseCard") Then
            If FillVSF选定(Val(vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("药品ID")))) = False Then Exit Sub
        Else
            '单独处理其他出库、移库、外购退货，分批的无库存不能移库和出库
            If mbyt编辑模式 = 2 Then
                '出库才判断分批属性
                int分批 = Get库房分批()
                If (int分批 = 3 And mbyt库房性质 <> 3) Or (int分批 = 1 And mbyt库房性质 = 1) Or (int分批 = 2 And mbyt库房性质 = 2) Then
                    '库房分批
                    If grsSlave.RecordCount > 0 Or mint按批次出库 = 0 Then
                        If FillVSF选定(Val(vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("药品ID")))) = False Then Exit Sub
                    Else
                        MsgBox "该药品是分批药品且没有库存，不能继续操作！", vbInformation, gstrSysName
                        Exit Sub
                    End If
                Else
                    '不分批
                    If FillVSF选定(Val(vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("药品ID")))) = False Then Exit Sub
                End If
            Else
                If FillVSF选定(Val(vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("药品ID")))) = False Then Exit Sub
            End If
        End If
    End If
    
    If chkContinue.Value <> 1 Then
        Call CombinateRec
        mblnOK = True
        Unload Me
    End If
End Sub

Private Sub vsf规格_EnterCell()
    '如果已到执行日期而价格未执行，执行计算过程
    Dim lng药品id As Long
    
    lng药品id = Val(vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("药品ID")))
    If lng药品id = 0 Then Exit Sub
    
    If mlngLast <> lng药品id Then
        Call AutoAdjustPrice_ByID(lng药品id)
    End If
    mlngLast = lng药品id
End Sub

Private Sub vsf规格_GotFocus()
    SetGridFocus vsf规格, True
End Sub

Private Sub vsf规格_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If KeyCode = vbKeyReturn Then
        If vsf规格.RowHidden(vsf规格.Row) Then Exit Sub '当前行隐藏则退出过程
        
        Call vsf规格_DblClick
    End If
    
End Sub

Private Sub vsf规格_LostFocus()
    SetGridFocus vsf规格, False
End Sub

Private Sub vsf批次_DblClick()
    Dim intRow As Integer
    If GetVSFlexRows(vsf批次) <= 1 Then
        If mint按批次出库 = 1 Then
            MsgBox "分批药品按批次出库，无库存不能继续操作！"
        End If
        Exit Sub
    End If
    If glngModul = 1305 And mint按批次出库 = 0 Then '领用
        With vsf规格
            If .TextMatrix(.Row, .ColIndex("领用标志")) = "0" Or .TextMatrix(.Row, .ColIndex("领用标志")) = "" Then
                MsgBox "该部门不允许领用该药品，请到药品储备限额中设置！", vbInformation, gstrSysName
                Exit Sub
            End If
            
            For intRow = 1 To vsf选定.rows - 1
                If .TextMatrix(.Row, .ColIndex("药品id")) = vsf选定.TextMatrix(intRow, vsf选定.ColIndex("药品id")) Then
                    MsgBox "该药品已经在药品选择器中，不允许在同张单据中领用多次！", vbInformation, gstrSysName
                    Exit Sub
                End If
            Next
        End With
    End If
    
    If glngModul = 1304 And mint按批次出库 = 0 Then  '移库
        With vsf规格
            For intRow = 1 To vsf选定.rows - 1
                If .TextMatrix(.Row, .ColIndex("药品id")) = vsf选定.TextMatrix(intRow, vsf选定.ColIndex("药品id")) Then
                    MsgBox "该药品已经在药品选择器中，不允许在同张单据中移库多次！", vbInformation, gstrSysName
                    Exit Sub
                End If
            Next
        End With
    End If
    
    If glngModul = 1343 And mint按批次出库 = 0 Then    '申领
        With vsf规格
            For intRow = 1 To vsf选定.rows - 1
                If .TextMatrix(.Row, .ColIndex("药品id")) = vsf选定.TextMatrix(intRow, vsf选定.ColIndex("药品id")) Then
                    MsgBox "该药品已经在药品选择器中，不允许在同张单据中申领多次！", vbInformation, gstrSysName
                    Exit Sub
                End If
            Next
        End With
    End If
    
    '零差价管理：选定药品不满足零差价要求，则不能进行入出库业务
'    If glngModul <> 1307 Or (glngModul = 1307 And vsf批次.TextMatrix(vsf批次.Row, vsf批次.ColIndex("批号")) <> "新增批次药品") Then
        If gtype_UserSysParms.P275_零差价管理模式 = 2 And IsPriceAdjustMod(Val(vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("药品ID")))) = True Then
            If CheckPriceAdjust(Val(vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("药品ID"))), mlng来源库房, vsf批次.TextMatrix(vsf批次.Row, vsf批次.ColIndex("批次"))) = False Then
                MsgBox "该药品启用零差价管理模式，但成本价和售价不一致，不能开展业务。请先调整价格！", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
'    End If
    
    
    If FillVSF选定(Val(vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("药品ID"))), vsf批次.TextMatrix(vsf批次.Row, vsf批次.ColIndex("批次"))) = False Then Exit Sub
    
    If chkContinue.Value <> 1 Then
        If CombinateRec = False Then Exit Sub
        Unload Me
    End If
End Sub

Private Function Get库房分批() As Integer
'-------------------------------------------------------------------------------------------
'功能：计算库房分批属性
'返回：0：库房不分批； 1：药库分批，药房不分批； 2：药库不分批，药房分批； 3：库房都分批
'-------------------------------------------------------------------------------------------
    Dim intReturn As Integer
    With vsf规格
        If .TextMatrix(.Row, .ColIndex("药库分批")) = "是" Or .TextMatrix(.Row, .ColIndex("药房分批")) = "是" Then
            If .TextMatrix(.Row, .ColIndex("药库分批")) = "是" And .TextMatrix(.Row, .ColIndex("药房分批")) = "是" Then
                intReturn = 3
            ElseIf .TextMatrix(.Row, .ColIndex("药库分批")) = "是" Then
                intReturn = 1
            Else
                intReturn = 2
            End If
        End If
    End With
    Get库房分批 = intReturn
End Function

Private Function FillVSF选定(ByVal lngDrugID As Long, Optional ByVal str批次 As String) As Boolean
    Dim blnValid As Boolean
    Dim lngRow As Long, i As Long
    Dim int分批 As Integer
    Dim dblPrice As Double
    Dim rstemp As ADODB.Recordset
    
    On Error GoTo errHandle
    '检查药品重复
    If chkContinue.Value = 1 Then
        For i = 1 To vsf选定.rows - 1
            If Val(vsf选定.TextMatrix(i, vsf选定.ColIndex("药品ID"))) = lngDrugID Then
                If vsf批次.Visible Then
                    If vsf选定.TextMatrix(i, vsf选定.ColIndex("批次")) = str批次 Then
                        Exit Function
                    End If
                Else
                    Exit Function
                End If
            End If
        Next
    End If
    
    '出库类型的数据检查
    If mbyt编辑模式 = 2 Then If CheckData = False Then Exit Function
'
'    '检查分批属性与库存数据是否一致
'    blnValid = 检查库存数据(IIf(mbyt编辑模式 = 2, mlng来源库房, mlng目标库房), lngDrugID)
'    If blnValid = False Then
'        MsgBox "发现该药品在当前库房中的库存记录存在错误（可能是基础数据设置错误，请检查当前库房的部门性质及该药品的分批属性）！", vbInformation, gstrSysName
'        Exit Function
'    End If
    
    '分批
'    int分批 = Get库房分批()
        
    '填充vsf选定
    With vsf选定
        .rows = .rows + 1
        lngRow = .rows - 1
        .TextMatrix(lngRow, .ColIndex("剂型")) = vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("剂型"))
        .TextMatrix(lngRow, .ColIndex("药名编码")) = vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("药名编码"))
        .TextMatrix(lngRow, .ColIndex("来源")) = vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("来源"))
        .TextMatrix(lngRow, .ColIndex("基本药物")) = vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("基本药物"))
        .TextMatrix(lngRow, .ColIndex("通用名称")) = vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("通用名称"))
        .TextMatrix(lngRow, .ColIndex("药典ID")) = vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("药典ID"))
        .TextMatrix(lngRow, .ColIndex("用途分类ID")) = vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("用途分类ID"))
        .TextMatrix(lngRow, .ColIndex("剂量单位")) = vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("剂量单位"))
        .TextMatrix(lngRow, .ColIndex("药品编码")) = vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("药品编码"))
        .TextMatrix(lngRow, .ColIndex("药品名称")) = vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("药品名称"))
        .TextMatrix(lngRow, .ColIndex("商品名")) = vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("商品名"))
        .TextMatrix(lngRow, .ColIndex("规格")) = vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("规格"))
        .TextMatrix(lngRow, .ColIndex("药名ID")) = vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("药名ID"))
        .TextMatrix(lngRow, .ColIndex("药品ID")) = vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("药品ID"))
        
        If Get售价(Val(.TextMatrix(lngRow, .ColIndex("药品ID"))), dblPrice) = False Then
            .RemoveItem .rows - 1
            Exit Function
        End If
        .TextMatrix(lngRow, .ColIndex("售价")) = vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("售价"))
        
        .TextMatrix(lngRow, .ColIndex("售价单位")) = vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("售价单位"))
        .TextMatrix(lngRow, .ColIndex("售价包装")) = vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("售价包装"))
        .TextMatrix(lngRow, .ColIndex("最大效期")) = vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("有效期"))
        .TextMatrix(lngRow, .ColIndex("门诊单位")) = vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("门诊单位"))
        .TextMatrix(lngRow, .ColIndex("门诊包装")) = vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("门诊包装"))
        .TextMatrix(lngRow, .ColIndex("住院单位")) = vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("住院单位"))
        .TextMatrix(lngRow, .ColIndex("住院包装")) = vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("住院包装"))
        .TextMatrix(lngRow, .ColIndex("药库单位")) = vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("药库单位"))
        .TextMatrix(lngRow, .ColIndex("药库包装")) = vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("药库包装"))
        .TextMatrix(lngRow, .ColIndex("药库分批")) = vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("药库分批"))
        .TextMatrix(lngRow, .ColIndex("药房分批")) = vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("药房分批"))
        .TextMatrix(lngRow, .ColIndex("时价")) = vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("时价"))
        .TextMatrix(lngRow, .ColIndex("批准文号")) = vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("批准文号"))
        .TextMatrix(lngRow, .ColIndex("指导批发价")) = vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("指导批发价"))
        .TextMatrix(lngRow, .ColIndex("加成率")) = vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("加成率"))
        .TextMatrix(lngRow, .ColIndex("生产商")) = vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("生产商"))
        .TextMatrix(lngRow, .ColIndex("原产地")) = vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("原产地"))
        
        '成本价
        dblPrice = Val(vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("上次采购价")))
        If dblPrice = 0 Then
            dblPrice = Val(vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("指导批发价")))
        End If
'        Select Case mintUnit
'        Case mconint门诊单位
'            dblPrice = dblPrice / IIf(Val(.TextMatrix(lngRow, .ColIndex("门诊包装"))) = 0, 1, Val(.TextMatrix(lngRow, .ColIndex("门诊包装"))))
'        Case mconint住院单位
'            dblPrice = dblPrice / IIf(Val(.TextMatrix(lngRow, .ColIndex("住院包装"))) = 0, 1, Val(.TextMatrix(lngRow, .ColIndex("住院包装"))))
'        Case mconint药库单位
'            dblPrice = dblPrice / IIf(Val(.TextMatrix(lngRow, .ColIndex("药库包装"))) = 0, 1, Val(.TextMatrix(lngRow, .ColIndex("药库包装"))))
'        End Select
        .TextMatrix(lngRow, .ColIndex("成本价")) = dblPrice
        
        'If mbyt编辑模式 = 2 And (int分批 = 3 And mbyt库房性质 <> 3) Or (int分批 = 1 And mbyt库房性质 = 1) Or (int分批 = 2 And mbyt库房性质 = 2) Then
        If vsf批次.Visible Then
            .TextMatrix(lngRow, .ColIndex("批号")) = vsf批次.TextMatrix(vsf批次.Row, vsf批次.ColIndex("批号"))
            .TextMatrix(lngRow, .ColIndex("生产日期")) = vsf批次.TextMatrix(vsf批次.Row, vsf批次.ColIndex("生产日期"))
            .TextMatrix(lngRow, .ColIndex("上次供应商ID")) = vsf批次.TextMatrix(vsf批次.Row, vsf批次.ColIndex("上次供应商ID"))
            .TextMatrix(lngRow, .ColIndex("有效期")) = vsf批次.TextMatrix(vsf批次.Row, vsf批次.ColIndex("有效期"))
            .TextMatrix(lngRow, .ColIndex("批准文号")) = vsf批次.TextMatrix(vsf批次.Row, vsf批次.ColIndex("批准文号"))
            .TextMatrix(lngRow, .ColIndex("生产商")) = vsf批次.TextMatrix(vsf批次.Row, vsf批次.ColIndex("生产商"))
            .TextMatrix(lngRow, .ColIndex("原产地")) = vsf批次.TextMatrix(vsf批次.Row, vsf批次.ColIndex("原产地"))
            
            .TextMatrix(lngRow, .ColIndex("批次")) = vsf批次.TextMatrix(vsf批次.Row, vsf批次.ColIndex("批次"))
            '库存盘点记录单，无“查看盘点单库存”参数处理
            If Not mblnStore Then
                ProcQuantity vsf选定, lngRow, .TextMatrix(lngRow, .ColIndex("药品ID"))
            Else
                .TextMatrix(lngRow, .ColIndex("可用数量")) = vsf批次.TextMatrix(vsf批次.Row, vsf批次.ColIndex("可用数量"))
                .TextMatrix(lngRow, .ColIndex("库存数量")) = vsf批次.TextMatrix(vsf批次.Row, vsf批次.ColIndex("库存数量"))
                .TextMatrix(lngRow, .ColIndex("库存金额")) = vsf批次.TextMatrix(vsf批次.Row, vsf批次.ColIndex("库存金额"))
                .TextMatrix(lngRow, .ColIndex("库存差价")) = vsf批次.TextMatrix(vsf批次.Row, vsf批次.ColIndex("库存差价"))
                .TextMatrix(lngRow, .ColIndex("实际数量")) = vsf批次.TextMatrix(vsf批次.Row, vsf批次.ColIndex("实际数量"))
            End If
        Else
            '提取不分批药品的批号与效期信息
            gstrSQL = "Select 上次批号,效期,上次供应商id,上次生产日期 AS 生产日期,批准文号,上次产地 From 药品库存 " & _
                     " Where 库房ID=[1] And 药品ID=[2] And 性质=1 "
            Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[提取不分批药品的批号与效期信息]", mlng来源库房, CLng(lngDrugID))
            If rstemp.RecordCount > 0 Then
                .TextMatrix(lngRow, .ColIndex("批号")) = zlStr.Nvl(rstemp!上次批号)
                If Not IsNull(rstemp!生产日期) Then
                    .TextMatrix(lngRow, .ColIndex("生产日期")) = zlStr.Nvl(rstemp!生产日期)
                End If
                .TextMatrix(lngRow, .ColIndex("上次供应商ID")) = zlStr.Nvl(rstemp!上次供应商ID)
                .TextMatrix(lngRow, .ColIndex("生产商")) = zlStr.Nvl(rstemp!上次产地)
                
                If Not IsNull(rstemp!效期) Then
                    'If gtype_UserSysParms.P149_效期显示方式 = 1 And Nvl(!效期) <> "" Then
                    If .TextMatrix(0, .ColIndex("有效期")) = "有效期至" Then
                        '换算为有效期
                        .TextMatrix(lngRow, .ColIndex("有效期")) = Format(DateAdd("D", -1, rstemp!效期), "yyyy-mm-dd")
                    Else
                        .TextMatrix(lngRow, .ColIndex("有效期")) = zlStr.Nvl(rstemp!效期)
                    End If
                End If
                .TextMatrix(lngRow, .ColIndex("批准文号")) = zlStr.Nvl(rstemp!批准文号)
            End If
            rstemp.Close
            
            .TextMatrix(lngRow, .ColIndex("批次")) = "0"
            '库存盘点记录单，无“查看盘点单库存”参数处理
            If Not mblnStore Then
                ProcQuantity vsf选定, lngRow, .TextMatrix(lngRow, .ColIndex("药品ID"))
            Else
                .TextMatrix(lngRow, .ColIndex("可用数量")) = vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("可用数量"))
                .TextMatrix(lngRow, .ColIndex("库存数量")) = vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("库存数量"))
                .TextMatrix(lngRow, .ColIndex("库存金额")) = vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("库存金额"))
                .TextMatrix(lngRow, .ColIndex("库存差价")) = vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("库存差价"))
                .TextMatrix(lngRow, .ColIndex("实际数量")) = vsf规格.TextMatrix(vsf规格.Row, vsf规格.ColIndex("实际数量"))
            End If
        End If
    End With
    
    lbl选定.Caption = "选定药品（" & vsf选定.rows - 1 & "条）"
    
    '选定完成
    FillVSF选定 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ProcQuantity(ByVal vsfVal As VSFlexGrid, ByVal lngRow As Long, ByVal lngDrugID As Long)
'----------------------------------------------------------
'功能：单独处理库存盘点记录单，无“查看盘点单库存”参数处理
'----------------------------------------------------------
    Dim dbl可用数量 As Double, dbl库存数量 As Double, dbl实际数量 As Double
    Dim dbl库存金额 As Double, dbl库存差价 As Double
    With grsMaster
        .Find "药品id=" & lngDrugID
        If Not .EOF Then
            dbl可用数量 = zlStr.Nvl(!可用数量, 0)
            dbl库存数量 = zlStr.Nvl(!库存数量, 0)
            dbl库存金额 = zlStr.Nvl(!库存金额, 0)
            dbl库存差价 = zlStr.Nvl(!库存差价, 0)
            dbl实际数量 = zlStr.Nvl(!实际数量, 0)
        End If
    End With
    With vsfVal
        .TextMatrix(lngRow, .ColIndex("可用数量")) = dbl可用数量
        .TextMatrix(lngRow, .ColIndex("库存数量")) = dbl库存数量
        .TextMatrix(lngRow, .ColIndex("库存金额")) = dbl库存金额
        .TextMatrix(lngRow, .ColIndex("库存差价")) = dbl库存差价
        .TextMatrix(lngRow, .ColIndex("实际数量")) = dbl实际数量
    End With
End Sub

Private Function Get售价(ByVal lngDrugID As Long, ByRef dblPrice As Double) As Boolean
'-----------------------------------
'功能：提取指定药品的零售单位价格
'返回：True成功；False失败
'-----------------------------------
    Dim rstemp As ADODB.Recordset
    Dim strMsg As String
    
    On Error GoTo errHandle
    gstrSQL = "Select A.现价, B.指导批发价, B.指导零售价, C.编码 药名编码, C.名称 通用名称 " & _
              "From 收费价目 A, 药品规格 B, 收费项目目录 C " & _
              "Where A.收费细目id = B.药品id And B.药品ID = C.ID " & _
              "  And Sysdate Between A.执行日期 And Nvl(A.终止日期,Sysdate) And A.收费细目ID=[1] " & _
              GetPriceClassString("A")
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[提取该药品的零售单位价格]", lngDrugID)
    
    If Not rstemp.EOF Then
        dblPrice = zlStr.Nvl(rstemp!现价, 0)
    Else
        dblPrice = 0
    End If
    
    '检查指导批发价，指导零售价，为0时不允许对该药品操作
    strMsg = ""
    If Not rstemp.EOF Then
        If rstemp!指导批发价 = 0 And rstemp!指导零售价 = 0 Then
            strMsg = "采购限价和指导售价为0，请先设置价格。"
        ElseIf rstemp!指导批发价 = 0 Then
            strMsg = "采购限价为0，请先设置价格。"
        ElseIf rstemp!指导零售价 = 0 Then
            strMsg = "指导售价为0，请先设置价格。"
        End If
        If strMsg <> "" Then strMsg = "[" & zlStr.Nvl(rstemp!药名编码) & zlStr.Nvl(rstemp!通用名称) & "]" & strMsg
    End If
    rstemp.Close
    
    If strMsg <> "" Then
        MsgBox strMsg, vbInformation, gstrSysName
        Get售价 = False
        Exit Function
    End If
    
    Get售价 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CombinateRec() As Boolean
    Dim i As Long
    Dim dblPrice As Double
    
    With vsf选定
        For i = 1 To .rows - 1
            If Val(.TextMatrix(i, .ColIndex("药品ID"))) > 0 Then
                mrsReturn.AddNew
                mrsReturn!剂型 = IIf(.TextMatrix(i, .ColIndex("剂型")) = "", Null, .TextMatrix(i, .ColIndex("剂型")))
                mrsReturn!药品编码 = .TextMatrix(i, .ColIndex("药品编码"))
                mrsReturn!药品来源 = .TextMatrix(i, .ColIndex("来源"))
                mrsReturn!基本药物 = .TextMatrix(i, .ColIndex("基本药物"))
                mrsReturn!通用名 = .TextMatrix(i, .ColIndex("通用名称"))
                mrsReturn!药典ID = Val(.TextMatrix(i, .ColIndex("药典ID")))
                mrsReturn!用途分类id = Val(.TextMatrix(i, .ColIndex("用途分类ID")))
                mrsReturn!剂量单位 = .TextMatrix(i, .ColIndex("剂量单位"))
                mrsReturn!商品名 = .TextMatrix(i, .ColIndex("商品名"))
                mrsReturn!规格 = .TextMatrix(i, .ColIndex("规格"))
                mrsReturn!产地 = .TextMatrix(i, .ColIndex("生产商"))
                mrsReturn!原产地 = .TextMatrix(i, .ColIndex("原产地"))
                mrsReturn!药名ID = Val(.TextMatrix(i, .ColIndex("药名ID")))
                mrsReturn!药品ID = Val(.TextMatrix(i, .ColIndex("药品ID")))
                
                mrsReturn!售价单位 = .TextMatrix(i, .ColIndex("售价单位"))
                mrsReturn!剂量系数 = Val(.TextMatrix(i, .ColIndex("售价包装")))
                mrsReturn!最大效期 = .TextMatrix(i, .ColIndex("最大效期"))
                mrsReturn!门诊单位 = .TextMatrix(i, .ColIndex("门诊单位"))
                mrsReturn!门诊包装 = Val(.TextMatrix(i, .ColIndex("门诊包装")))
                mrsReturn!住院单位 = .TextMatrix(i, .ColIndex("住院单位"))
                mrsReturn!住院包装 = Val(.TextMatrix(i, .ColIndex("住院包装")))
                mrsReturn!药库单位 = .TextMatrix(i, .ColIndex("药库单位"))
                mrsReturn!药库包装 = Val(.TextMatrix(i, .ColIndex("药库包装")))
                mrsReturn!药库分批 = IIf(.TextMatrix(i, .ColIndex("药库分批")) = "是", 1, 0)
                mrsReturn!药房分批 = IIf(.TextMatrix(i, .ColIndex("药房分批")) = "是", 1, 0)
                mrsReturn!时价 = IIf(.TextMatrix(i, .ColIndex("时价")) = "是", 1, 0)
                mrsReturn!上次供应商ID = Val(.TextMatrix(i, .ColIndex("上次供应商ID")))
                mrsReturn!批准文号 = .TextMatrix(i, .ColIndex("批准文号"))
                mrsReturn!批次 = IIf(.TextMatrix(i, .ColIndex("批次")) = "", "0", .TextMatrix(i, .ColIndex("批次")))
                mrsReturn!批号 = IIf(.TextMatrix(i, .ColIndex("批号")) = "", Null, .TextMatrix(i, .ColIndex("批号")))
                mrsReturn!生产日期 = IIf(.TextMatrix(i, .ColIndex("生产日期")) = "", Null, .TextMatrix(i, .ColIndex("生产日期")))
                mrsReturn!效期 = IIf(.TextMatrix(i, .ColIndex("有效期")) = "", Null, .TextMatrix(i, .ColIndex("有效期")))
                mrsReturn!可用数量 = Val(.TextMatrix(i, .ColIndex("可用数量")))
                mrsReturn!实际数量 = Val(.TextMatrix(i, .ColIndex("实际数量")))
                mrsReturn!实际金额 = Val(.TextMatrix(i, .ColIndex("库存金额")))
                mrsReturn!实际差价 = Val(.TextMatrix(i, .ColIndex("库存差价")))
                mrsReturn!库存数量 = Val(.TextMatrix(i, .ColIndex("库存数量")))
                
                dblPrice = Val(.TextMatrix(i, .ColIndex("指导批发价")))
                Select Case mintUnit
                    Case mconint门诊单位
                        dblPrice = dblPrice / IIf(Val(.TextMatrix(i, .ColIndex("门诊包装"))) = 0, 1, Val(.TextMatrix(i, .ColIndex("门诊包装"))))
                    Case mconint住院单位
                        dblPrice = dblPrice / IIf(Val(.TextMatrix(i, .ColIndex("住院包装"))) = 0, 1, Val(.TextMatrix(i, .ColIndex("住院包装"))))
                    Case mconint药库单位
                        dblPrice = dblPrice / IIf(Val(.TextMatrix(i, .ColIndex("药库包装"))) = 0, 1, Val(.TextMatrix(i, .ColIndex("药库包装"))))
                End Select
                mrsReturn!指导批发价 = dblPrice
                                
                mrsReturn!加成率 = Val(.TextMatrix(i, .ColIndex("加成率")))
                
                dblPrice = Val(.TextMatrix(i, .ColIndex("售价")))
                Select Case mintUnit
                    Case mconint门诊单位
                        dblPrice = dblPrice / IIf(Val(.TextMatrix(i, .ColIndex("门诊包装"))) = 0, 1, Val(.TextMatrix(i, .ColIndex("门诊包装"))))
                    Case mconint住院单位
                        dblPrice = dblPrice / IIf(Val(.TextMatrix(i, .ColIndex("住院包装"))) = 0, 1, Val(.TextMatrix(i, .ColIndex("住院包装"))))
                    Case mconint药库单位
                        dblPrice = dblPrice / IIf(Val(.TextMatrix(i, .ColIndex("药库包装"))) = 0, 1, Val(.TextMatrix(i, .ColIndex("药库包装"))))
                End Select
                mrsReturn!售价 = dblPrice
                
                dblPrice = Val(.TextMatrix(i, .ColIndex("成本价")))
                Select Case mintUnit
                    Case mconint门诊单位
                        dblPrice = dblPrice / IIf(Val(.TextMatrix(i, .ColIndex("门诊包装"))) = 0, 1, Val(.TextMatrix(i, .ColIndex("门诊包装"))))
                    Case mconint住院单位
                        dblPrice = dblPrice / IIf(Val(.TextMatrix(i, .ColIndex("住院包装"))) = 0, 1, Val(.TextMatrix(i, .ColIndex("住院包装"))))
                    Case mconint药库单位
                        dblPrice = dblPrice / IIf(Val(.TextMatrix(i, .ColIndex("药库包装"))) = 0, 1, Val(.TextMatrix(i, .ColIndex("药库包装"))))
                End Select
                mrsReturn!成本价 = dblPrice
                
                mrsReturn.Update
            End If
        Next
    End With
    CombinateRec = True
End Function

Private Sub InitReturnRecord()
    Set mrsReturn = New ADODB.Recordset
    With mrsReturn
        .Fields.Append "剂型", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "药名编码", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "药品来源", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "基本药物", adVarChar, 30, adFldIsNullable
        .Fields.Append "通用名", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "药典ID", adDouble, 18, adFldIsNullable
        .Fields.Append "用途分类ID", adDouble, 18, adFldIsNullable
        .Fields.Append "剂量单位", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "药品编码", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "商品名", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "规格", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "产地", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "原产地", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "药名ID", adDouble, 18, adFldIsNullable
        .Fields.Append "药品ID", adDouble, 18, adFldIsNullable
        .Fields.Append "售价", adDouble, 18, adFldIsNullable
        .Fields.Append "售价单位", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "剂量系数", adDouble, 11, adFldIsNullable
        .Fields.Append "最大效期", adDouble, 5, adFldIsNullable
        .Fields.Append "门诊单位", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "门诊包装", adDouble, 11, adFldIsNullable
        .Fields.Append "住院单位", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "住院包装", adDouble, 11, adFldIsNullable
        .Fields.Append "药库单位", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "药库包装", adDouble, 11, adFldIsNullable
        .Fields.Append "药库分批", adDouble, 2, adFldIsNullable
        .Fields.Append "药房分批", adDouble, 2, adFldIsNullable
        .Fields.Append "时价", adDouble, 2, adFldIsNullable
        .Fields.Append "批次", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "批号", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "生产日期", adDate, , adFldIsNullable
        .Fields.Append "效期", adDate, , adFldIsNullable
        .Fields.Append "可用数量", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "实际数量", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "实际金额", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "实际差价", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "指导批发价", adDouble, 11, adFldIsNullable
        .Fields.Append "加成率", adDouble, 11, adFldIsNullable
        .Fields.Append "上次供应商ID", adDouble, 18, adFldIsNullable
        .Fields.Append "库存数量", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "批准文号", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "成本价", adDouble, 11, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

'Private Function Get可用库存(ByVal lng药品ID As Long, Optional ByVal lng批次 As Long = 0) As Single
'    Dim rsStock As New ADODB.Recordset
'
'    gstrSQL = " Select Sum(A.可用数量" & StrUnitString & ") 可用数量,Sum(A.实际数量" & StrUnitString & ") 实际数量,sum(A.实际金额) 实际金额,sum(A.实际差价) 实际差价,Sum(A.实际数量) 库存数量 " & _
'              " From 药品库存 A,药品规格 B " & _
'              " Where A.药品ID=B.药品ID And A.性质=1 And A.药品ID=[1] " & IIf(lng批次 = 0, "", " And Nvl(A.批次,0)=[2] ")
'    If mlng来源库房 <> 0 Or mlng目标库房 <> 0 Then
'        gstrSQL = gstrSQL & " And A.库房ID=[3]"
'    End If
'    gstrSQL = gstrSQL & " Group By A.药品id"
'
'    Set rsStock = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[获取可用库存]", lng药品ID, lng批次, IIf(mlng来源库房 = 0, mlng目标库房, mlng来源库房))
'
'    mdbl可用数量 = 0
'    mdbl实际差价 = 0
'    mdbl实际金额 = 0
'    mdbl实际数量 = 0
'    mdbl库存数量 = 0
'    If Not rsStock.EOF Then
'        mdbl可用数量 = IIf(IsNull(rsStock!可用数量), 0, rsStock!可用数量)
'        mdbl实际差价 = IIf(IsNull(rsStock!实际差价), 0, rsStock!实际差价)
'        mdbl实际金额 = IIf(IsNull(rsStock!实际金额), 0, rsStock!实际金额)
'        mdbl实际数量 = IIf(IsNull(rsStock!实际数量), 0, rsStock!实际数量)
'        mdbl库存数量 = IIf(IsNull(rsStock!库存数量), 0, rsStock!库存数量)
'    End If
'    Get可用库存 = mdbl可用数量
'End Function

Private Sub vsf批次_GotFocus()
    SetGridFocus vsf批次, True
End Sub

Private Sub vsf批次_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call vsf批次_DblClick
End Sub

Private Sub vsf批次_LostFocus()
    SetGridFocus vsf批次, False
End Sub

Private Sub vsf选定_GotFocus()
    SetGridFocus vsf选定, True
End Sub

Private Sub vsf选定_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If vsf选定.rows > 1 Then
            vsf选定.RemoveItem vsf选定.Row
            If vsf选定.rows = 1 Then
                lbl选定.Caption = "选定药品"
            Else
                lbl选定.Caption = "选定药品（" & vsf选定.rows - 1 & "条）"
            End If
        End If
    End If
End Sub

Private Function Get部门名称(ByVal lngDeptId As Long) As String
    Dim strsql As String
    Dim rstemp As ADODB.Recordset
    
    On Error GoTo errHandle
    strsql = "select 名称 from 部门表 where ID=[1] "
    Set rstemp = zldatabase.OpenSQLRecord(strsql, "获取部门名称", lngDeptId)
    If Not rstemp.EOF Then
        Get部门名称 = zlStr.Nvl(rstemp!名称)
    End If
    rstemp.Close
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub FormatCols(ByVal strCols As String)
'功能：VSF绑定数据集对象后，对VSF列的调整
    Dim i As Integer
    Dim arrCols As Variant, arrColumn As Variant

    arrCols = Split(strCols, "|")
    With vsf规格
        .Redraw = False
        For i = LBound(arrCols) To UBound(arrCols)
            arrColumn = Split(arrCols(i), ",")
            If .ColIndex(arrColumn(0)) >= 0 Then
                '列顺序
                .ColPosition(.ColIndex(arrColumn(0))) = i
                '列属性
                If UBound(arrColumn) > 1 Then
                    .ColData(i) = IIf(arrColumn(2) = "", 3, Val(arrColumn(2)))
                Else
                    .ColData(i) = 3
                End If
                If .ColData(i) = 1 Or .ColData(i) = 2 Then
                    .ColHidden(i) = True
                Else
                    .ColHidden(i) = False
                End If
                '列宽度
                If UBound(arrColumn) > 2 Then
                    .ColWidth(i) = Val(arrColumn(3))
                Else
                    .ColWidth(i) = 0
                End If
                '显示格式
                If UBound(arrColumn) > 3 Then
                    If UCase(arrColumn(4)) = "D" Then
                        .ColFormat(i) = "yyyy-mm-dd"
                        .ColAlignment(i) = flexAlignCenterCenter
                    ElseIf UCase(arrColumn(4)) = "T" Then
                        .ColFormat(i) = "hh:mm:ss"
                        .ColAlignment(i) = flexAlignCenterCenter
                    ElseIf UCase(arrColumn(4)) = "DT" Then
                        .ColFormat(i) = "yyyy-mm-dd hh:mm:ss"
                        .ColAlignment(i) = flexAlignCenterCenter
                    ElseIf UCase(arrColumn(4)) = "N" Then
                        .ColAlignment(i) = flexAlignRightCenter
                    Else
                        .ColAlignment(i) = flexAlignLeftCenter
                    End If
                Else
                    .ColAlignment(i) = flexAlignLeftCenter
                End If
            End If
        Next
        '列头文本居中显示
        If .Cols > 0 Then .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .Redraw = True
    End With
End Sub

Private Sub SetColKey(ByVal vsfVal As VSFlexGrid)
'功能：VSF控件绑定数据集对象时，ColKey值没有，用0行的列值设为ColKey值
    Dim i As Integer
    For i = 0 To vsfVal.Cols - 1
        vsfVal.ColKey(i) = vsfVal.TextMatrix(0, i)
    Next
End Sub

Private Sub InitVSF(ByVal vsfVal As VSFlexGrid)
  With vsfVal
    .AllowUserResizing = flexResizeColumns
    .Appearance = flexFlat
    .BackColorAlternate = .BackColor
    .BackColorSel = glngRowByNotFocus
    .BackColorBkg = &H8000000C
    .ForeColorSel = vbBlack
    .ExplorerBar = flexExSortShowAndMove
    '.FixedCols = 0
    .GridColor = &H80000010
    .GridLinesFixed = flexGridFlat
    .SelectionMode = flexSelectionByRow
    .SheetBorder = &H80000005
  End With
End Sub

Private Sub vsf选定_LostFocus()
    SetGridFocus vsf选定, True
End Sub
Private Sub HiddenColumns()
    Dim rsDetail As New Recordset
    Dim bln中药库房 As Boolean
    
    On Error GoTo errHandle
        
    With vsf规格
        If .rows > 1 Then
            If mlng目标库房 = 0 And mlng来源库房 = 0 Then
                .ColWidth(.ColIndex("原产地")) = 0
                vsf批次.ColWidth(vsf批次.ColIndex("原产地")) = 0
                .ColData(.ColIndex("原产地")) = 1
                vsf批次.ColData(vsf批次.ColIndex("原产地")) = 1
                Exit Sub
            End If
            
            gstrSQL = "select 类别 from 收费项目目录  where id=[1]"
            Set rsDetail = zldatabase.OpenSQLRecord(gstrSQL, "判断库房性质", .TextMatrix(1, .ColIndex("药品id")))
    
            If rsDetail!类别 = "7" Then bln中药库房 = True
            If bln中药库房 Then
                .ColWidth(.ColIndex("原产地")) = 1000
                vsf批次.ColWidth(vsf批次.ColIndex("原产地")) = 1000
                .ColData(.ColIndex("原产地")) = 3
                vsf批次.ColData(vsf批次.ColIndex("原产地")) = 3
            Else
                vsf规格.ColWidth(vsf规格.ColIndex("原产地")) = 0
                vsf批次.ColWidth(vsf批次.ColIndex("原产地")) = 0
                .ColData(.ColIndex("原产地")) = 1
                vsf批次.ColData(vsf批次.ColIndex("原产地")) = 1
            End If
        End If
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
