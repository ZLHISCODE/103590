VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl UsrPriceQuery 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   8340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10440
   KeyPreview      =   -1  'True
   ScaleHeight     =   8340
   ScaleWidth      =   10440
   Begin VB.PictureBox picKey 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3075
      Left            =   5055
      MouseIcon       =   "UsrPriceQuery.ctx":0000
      MousePointer    =   99  'Custom
      ScaleHeight     =   3075
      ScaleWidth      =   5310
      TabIndex        =   25
      Top             =   645
      Width           =   5310
      Begin zl9NewQuery.ctlKeyBoard cmdKey 
         Height          =   3090
         Left            =   0
         TabIndex        =   46
         Top             =   0
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   5450
      End
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   3
      Left            =   2310
      ScaleHeight     =   285
      ScaleWidth      =   2580
      TabIndex        =   44
      Top             =   6240
      Width           =   2580
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "合计:10000000.00元"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   240
         Index           =   1
         Left            =   15
         TabIndex        =   45
         Top             =   30
         Width           =   2385
      End
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2460
      Left            =   75
      MouseIcon       =   "UsrPriceQuery.ctx":030A
      MousePointer    =   99  'Custom
      ScaleHeight     =   2460
      ScaleWidth      =   1845
      TabIndex        =   15
      Top             =   4875
      Width           =   1845
      Begin VB.CommandButton cmdBtn 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   0
         Left            =   30
         TabIndex        =   31
         Top             =   45
         Width           =   600
      End
      Begin VB.CommandButton cmdBtn 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   1
         Left            =   630
         TabIndex        =   30
         Top             =   45
         Width           =   600
      End
      Begin VB.CommandButton cmdBtn 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   2
         Left            =   1230
         TabIndex        =   29
         Top             =   45
         Width           =   600
      End
      Begin VB.CommandButton cmdBtn 
         Caption         =   "清除"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   12
         Left            =   1230
         TabIndex        =   24
         Top             =   1845
         Width           =   600
      End
      Begin VB.CommandButton cmdBtn 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   11
         Left            =   630
         TabIndex        =   23
         Top             =   1845
         Width           =   600
      End
      Begin VB.CommandButton cmdBtn 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   9
         Left            =   30
         TabIndex        =   22
         Top             =   1845
         Width           =   600
      End
      Begin VB.CommandButton cmdBtn 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   8
         Left            =   1230
         TabIndex        =   21
         Top             =   1245
         Width           =   600
      End
      Begin VB.CommandButton cmdBtn 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   7
         Left            =   630
         TabIndex        =   20
         Top             =   1245
         Width           =   600
      End
      Begin VB.CommandButton cmdBtn 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   6
         Left            =   30
         TabIndex        =   19
         Top             =   1245
         Width           =   600
      End
      Begin VB.CommandButton cmdBtn 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   5
         Left            =   1230
         TabIndex        =   18
         Top             =   645
         Width           =   600
      End
      Begin VB.CommandButton cmdBtn 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   4
         Left            =   630
         TabIndex        =   17
         Top             =   645
         Width           =   600
      End
      Begin VB.CommandButton cmdBtn 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   3
         Left            =   30
         TabIndex        =   16
         Top             =   645
         Width           =   600
      End
      Begin VB.CommandButton cmdBtn 
         Caption         =   "确定"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   10
         Left            =   900
         TabIndex        =   28
         Top             =   1455
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   930
      Index           =   0
      Left            =   15
      ScaleHeight     =   930
      ScaleWidth      =   10425
      TabIndex        =   0
      Top             =   3915
      Width           =   10425
      Begin VB.Frame fra2 
         Caption         =   "Frame1"
         Height          =   960
         Left            =   4800
         TabIndex        =   41
         Top             =   -60
         Width           =   30
      End
      Begin VB.Frame fra 
         Height          =   30
         Left            =   5430
         TabIndex        =   33
         Top             =   435
         Width           =   4920
      End
      Begin VB.TextBox txt 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   7065
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   555
         Width           =   1500
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   1
         Left            =   8640
         TabIndex        =   26
         Top             =   480
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   741
         Caption         =   "病人自助查询"
         BackColor       =   16777215
         FontSize        =   10.5
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   14
         Left            =   3780
         TabIndex        =   42
         Top             =   30
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   741
         Caption         =   "上翻"
         BackColor       =   16777215
         ForeColor       =   12583104
         FontSize        =   10.5
         AutoSize        =   0   'False
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   15
         Left            =   3780
         TabIndex        =   43
         Top             =   480
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   741
         Caption         =   "下翻"
         BackColor       =   16777215
         ForeColor       =   12583104
         FontSize        =   10.5
         AutoSize        =   0   'False
         TextAligment    =   0
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "项目说明:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   4
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Width           =   3225
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "价格依据:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   3
         Left            =   6615
         TabIndex        =   32
         Top             =   90
         Width           =   945
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请输入查询简码:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   210
         Index           =   2
         Left            =   5415
         TabIndex        =   2
         Top             =   615
         Width           =   1575
      End
   End
   Begin VB.Timer tmrScroll 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2640
      Top             =   2640
   End
   Begin VB.Timer tmrInfo 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   2685
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   2
      Left            =   75
      ScaleHeight     =   480
      ScaleWidth      =   7920
      TabIndex        =   12
      Top             =   7350
      Width           =   7920
      Begin VB.TextBox txt 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   495
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   37
         Top             =   90
         Width           =   645
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   12
         Left            =   5670
         TabIndex        =   13
         Top             =   15
         Width           =   540
         _ExtentX        =   1482
         _ExtentY        =   741
         Caption         =   "上翻"
         BackColor       =   16777215
         FontSize        =   10.5
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   13
         Left            =   6225
         TabIndex        =   14
         Top             =   15
         Width           =   540
         _ExtentX        =   1482
         _ExtentY        =   741
         Caption         =   "下翻"
         BackColor       =   16777215
         FontSize        =   10.5
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   0
         Left            =   1845
         TabIndex        =   38
         Top             =   30
         Width           =   540
         _ExtentX        =   1482
         _ExtentY        =   741
         Caption         =   "删除"
         BackColor       =   16777215
         FontSize        =   10.5
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   3
         Left            =   1215
         TabIndex        =   39
         Top             =   30
         Width           =   540
         _ExtentX        =   1482
         _ExtentY        =   741
         Caption         =   "添加"
         BackColor       =   16777215
         FontSize        =   10.5
         TextAligment    =   0
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "数量"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   210
         Index           =   0
         Left            =   45
         TabIndex        =   40
         Top             =   150
         Width           =   420
      End
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   1
      Left            =   450
      ScaleHeight     =   450
      ScaleWidth      =   9540
      TabIndex        =   3
      Top             =   0
      Width           =   9540
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   4
         Left            =   45
         TabIndex        =   4
         Top             =   15
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   741
         Caption         =   "药疗"
         BackColor       =   16777215
         FontSize        =   10.5
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   5
         Left            =   930
         TabIndex        =   5
         Top             =   15
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   741
         Caption         =   "检验"
         BackColor       =   16777215
         FontSize        =   10.5
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   6
         Left            =   1815
         TabIndex        =   6
         Top             =   15
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   741
         Caption         =   "检查"
         BackColor       =   16777215
         FontSize        =   10.5
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   7
         Left            =   2700
         TabIndex        =   7
         Top             =   15
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   741
         Caption         =   "治疗"
         BackColor       =   16777215
         FontSize        =   10.5
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   8
         Left            =   3585
         TabIndex        =   8
         Top             =   15
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   741
         Caption         =   "手术"
         BackColor       =   16777215
         FontSize        =   10.5
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   9
         Left            =   4485
         TabIndex        =   9
         Top             =   15
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   741
         Caption         =   "其他所有"
         BackColor       =   16777215
         FontSize        =   10.5
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   10
         Left            =   6660
         TabIndex        =   10
         Top             =   15
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   741
         Caption         =   "上翻"
         BackColor       =   16777215
         FontSize        =   10.5
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   11
         Left            =   7560
         TabIndex        =   11
         Top             =   15
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   741
         Caption         =   "下翻"
         BackColor       =   16777215
         FontSize        =   10.5
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   2
         Left            =   8625
         TabIndex        =   27
         Top             =   15
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   741
         Caption         =   "教您怎么查价格?"
         BackColor       =   16777215
         FontSize        =   10.5
         TextAligment    =   0
      End
   End
   Begin MSComctlLib.ImageList ilsImage 
      Left            =   960
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrPriceQuery.ctx":0614
            Key             =   "search"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrPriceQuery.ctx":09AE
            Key             =   "hide"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrPriceQuery.ctx":2AE8
            Key             =   "add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrPriceQuery.ctx":2E82
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrPriceQuery.ctx":321C
            Key             =   "remove"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrPriceQuery.ctx":35B6
            Key             =   "select"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrPriceQuery.ctx":3950
            Key             =   "unselect"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrPriceQuery.ctx":3CEA
            Key             =   "down"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrPriceQuery.ctx":4084
            Key             =   "up"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrPriceQuery.ctx":441E
            Key             =   "help"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid msfResult 
      Height          =   3285
      Left            =   525
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   795
      Width           =   4740
      _cx             =   8361
      _cy             =   5794
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483648
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16772055
      ForeColorSel    =   16711680
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   16761024
      GridColorFixed  =   16761024
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   345
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
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
      OwnerDraw       =   0
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
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid msfCalc 
      Height          =   1230
      Left            =   2355
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   4770
      Width           =   4740
      _cx             =   8361
      _cy             =   2170
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
      BackColorFixed  =   -2147483648
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16772055
      ForeColorSel    =   16711680
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   16761024
      GridColorFixed  =   16761024
      TreeColor       =   -2147483643
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
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
      OwnerDraw       =   0
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
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "UsrPriceQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mvarCurPos1 As Long
Private mvarRows1 As Long

Private mvarCurPos2 As Long
Private mvarRows2 As Long

Private mblnNumber As Boolean

Private mvarStop As Long                '用户查询信息停留间隔
Private mvarScroll As Long

Private mvarRs As New ADODB.Recordset
Private mrsPrice As New ADODB.Recordset         '时价药品价格（平均价）
Private mrs分类id As ADODB.Recordset
Private mblnUnSelect  As Boolean
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event ClickOK(ByVal strQuery As String, blnCancel As Boolean)

Private Enum mCol
    编码
    名称
    规格
    剂型
    单位
    价格
    指导售价
    产地
    标识主码
    标识子码
    费用类型
    价格依据
    项目说明
End Enum

Public Sub InitLoad()
    '初始化进入
    
    Dim i As Long
    
    mvarCurPos1 = 1
    mvarCurPos2 = 1
    
    UsrCmd(10).Enabled = False
    UsrCmd(11).Enabled = False
    UsrCmd(12).Enabled = False
    UsrCmd(13).Enabled = False
    UsrCmd(14).ShowPicture = False
    UsrCmd(15).ShowPicture = False
    
    msfResult.Rows = 50
    For i = 0 To msfResult.Cols - 1
        msfResult.TextMatrix(1, i) = ""
    Next
    
    msfCalc.Rows = 2
    For i = 0 To msfCalc.Cols - 1
        msfCalc.TextMatrix(1, i) = ""
    Next
                
    Call DrawMsfHeader
    
    UsrCmd(0).Picture = ilsImage.ListImages("delete")
    UsrCmd(1).Picture = ilsImage.ListImages("hide")
    UsrCmd(2).Picture = ilsImage.ListImages("help")
    UsrCmd(3).Picture = ilsImage.ListImages("add")
    
    UsrCmd(11).Picture = ilsImage.ListImages("down")
    UsrCmd(10).Picture = ilsImage.ListImages("up")
    
    UsrCmd(13).Picture = ilsImage.ListImages("down")
    UsrCmd(12).Picture = ilsImage.ListImages("up")
    
    UsrCmd(3).ShowPicture = False
    UsrCmd(0).ShowPicture = False
    UsrCmd(13).ShowPicture = False
    UsrCmd(12).ShowPicture = False
    
    Dim blnHave As Boolean
    Dim strTmp As String
    
    strTmp = Trim(zlDatabase.GetPara("价格显示类别", glngSys, 1536, "000000"))
    
    For i = 4 To 9
        If Val(Mid(strTmp, i - 3, 1)) = 1 Then
'        If Val(GetPara(UsrCmd(i).Caption)) = 1 Then
            UsrCmd(i).Picture = ilsImage.ListImages("select")
            UsrCmd(i).Tag = "1"
            blnHave = True
        Else
            UsrCmd(i).Picture = ilsImage.ListImages("unselect")
            UsrCmd(i).Tag = ""
        End If
    Next
    If blnHave = False Then
        For i = 4 To 9
            UsrCmd(i).Picture = ilsImage.ListImages("select")
            UsrCmd(i).Tag = "1"
        Next
    End If
    
    
    Dim varTmp As Variant
    Dim lngLoop As Long
    
    Set mvarRs = New ADODB.Recordset
    Set mrs分类id = New ADODB.Recordset
    mrs分类id.Fields.Append "分类id", adVarChar, 30, adFldKeyColumn
    mrs分类id.Open
    
    mblnUnSelect = False
    strTmp = ""
    strTmp = GetPara("允许显示的收费分类")
    If strTmp <> "" Then
        varTmp = Split(strTmp, ",")
        For lngLoop = 0 To UBound(varTmp)
            If CStr(varTmp(lngLoop)) <> "" Then
                
                mrs分类id.AddNew
                
                If Left(CStr(varTmp(lngLoop)), 1) = "-" Then
                    mblnUnSelect = True
                    mrs分类id("分类id").Value = Mid(CStr(varTmp(lngLoop)), 2)
                Else
                    mrs分类id("分类id").Value = CStr(varTmp(lngLoop))
                End If
                
                
            End If
        Next
    End If
    
    txt(0).Text = ""
    txt(1).Text = ""
                
    UsrCmd(1).Caption = "病人自助查询"
    Call UserControl_Resize
    
    Call CalcMoney
    Call SearchItem("")
    
    tmrInfo.Enabled = True
    
    mvarStop = Val(GetPara("价格查询停留时间", "30"))
    mvarStop = IIf(mvarStop <= 0, 30, mvarStop)
    
    mvarScroll = Val(GetPara("价格查询滚动间隔", "10"))
    mvarScroll = IIf(mvarScroll <= 0, 10, mvarScroll)
    
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    Dim intPos As Long
    
    Select Case cmdBtn(Index).Caption
    Case "确定"
        If msfResult.TextMatrix(msfResult.Row, mCol.名称) <> "" And Val(txt(0).Text) > 0 Then
            mvarCurPos2 = 1
            
            If msfCalc.Rows = 2 And msfCalc.TextMatrix(1, 0) = "" Then
                
            Else
                msfCalc.Rows = msfCalc.Rows + 1
            End If
                                                
            mvarRows2 = msfCalc.Rows
            
            msfCalc.TextMatrix(msfCalc.Rows - 1, 0) = msfResult.TextMatrix(msfResult.Row, mCol.名称)
            msfCalc.TextMatrix(msfCalc.Rows - 1, 1) = Val(txt(0).Text)
            
            intPos = InStr(msfResult.TextMatrix(msfResult.Row, mCol.价格), "(指导价)")
            
            If intPos > 0 Then
                msfCalc.TextMatrix(msfCalc.Rows - 1, 2) = Format(Val(txt(0).Text) * Val(msfResult.RowData(msfResult.Row)), "0.00")
            Else
                msfCalc.TextMatrix(msfCalc.Rows - 1, 2) = Format(Val(txt(0).Text) * Val(msfResult.TextMatrix(msfResult.Row, mCol.价格)), "0.00")
            End If
            
            Call CalcMoney
            Call EnablePageButton(msfCalc, mvarCurPos2, mvarRows2, UsrCmd(12), UsrCmd(13))
        ElseIf Val(txt(0).Text) <= 0 Then
            MsgBox "您必须先输入数量，才能添加！", vbInformation, gstrSysName
        ElseIf msfResult.TextMatrix(msfResult.Row, 0) = "" Then
            MsgBox "没有选中要添加的收费项目或当前没有收费项目！", vbInformation, gstrSysName
        End If

        EnterFocus msfResult
    Case "清除"
        txt(0).Text = ""
'        msfCalc.SetFocus
        EnterFocus msfCalc
    Case Else
        txt(0).Text = txt(0).Text & Trim(cmdBtn(Index).Caption)
        'msfResult.SetFocus
        EnterFocus msfResult
    End Select
    
End Sub

Private Sub cmdBtn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub cmdKey_CommandClick(Caption As String)
    Dim strTmp As String
    Dim blnCancel As Boolean

    tmrScroll.Enabled = False
    mvarStop = Val(GetPara("价格查询停留时间", 30))
    mvarStop = IIf(mvarStop <= 0, 30, mvarStop)
    
    Select Case Caption
    Case "确定"
        strTmp = txt(1).Text
        tmrScroll.Enabled = False
        lbl(3).Caption = "价格依据:"
        lbl(4).Caption = "项目说明:"
        '修改编号2667
        '判断输入的是不是adminexitnewquery如果用户取消就直接退出过程
        RaiseEvent ClickOK(strTmp, blnCancel)
        If blnCancel = True Then Exit Sub
        
        Call SearchItem(strTmp)
        
        msfResult.Row = 1
        Call msfResult_RowColChange
        txt(1).Text = ""
        mvarInfo = Val(GetPara("价格查询停留时间", "30"))
        tmrInfo.Enabled = IIf(mvarInfo = 0, False, True)
    Case "清除"
        txt(1).Text = ""
    Case Else
        txt(1).Text = txt(1).Text & Trim(Caption)
        txt(1).SelStart = Len(txt(1).Text & Trim(Caption))
    End Select
    'msfResult.SetFocus
    EnterFocus msfResult
End Sub

Private Sub cmdKey_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub


Private Sub msfCalc_Click()
    tmrScroll.Enabled = False
    mvarStop = Val(GetPara("价格查询停留时间", "30"))
    mvarStop = IIf(mvarStop <= 0, 30, mvarStop)
End Sub

Private Sub msfCalc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub msfResult_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call CalcAutoColWidth(msfResult, mCol.名称)
    Call SaveFlexState(msfResult, App.ProductName)
End Sub

Private Sub msfResult_Click()
    tmrScroll.Enabled = False
    mvarStop = Val(GetPara("价格查询停留时间", "30"))
    mvarStop = IIf(mvarStop <= 0, 30, mvarStop)
    
    If UsrCmd(1).Caption = "病人自助查询" Then
        Call UsrCmd_CommandClick(1)
    End If
End Sub

Private Sub msfResult_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub msfResult_RowColChange()
    
    On Error Resume Next
    lbl(3).Caption = "价格依据:" & msfResult.TextMatrix(msfResult.Row, mCol.价格依据)
    lbl(4).Caption = msfResult.TextMatrix(msfResult.Row, mCol.项目说明)
    If lbl(4).Caption = "" Then lbl(4).Caption = "项目说明:"
'    lbl(4).Caption = "用于整进零出和负数出库等原因导致的库存差价不准确时的调整处理（只允许一种入库类别）用于整进零出和负数出库等原因导致的库存差价不准确时的调整处理（只允许一种入库类别）用于整进零出和负数出库等原因导致的库存差价不准确时的调整处理（只允许一种入库类别）用于整进零出和负数出库等原因导致的库存差价不准确时的调整处理（只允许一种入库类别）用于整进零出和负数出库等原因导致的库存差价不准确时的调整处理（只允许一种入库类别）用于整进零出和负数出库等原因导致的库存差价不准确时的调整处理（只允许一种入库类别）用于整进零出和负数出库等原因导致的库存差价不准确时的调整处理（只允许一种入库类别）用于整进零出和负数出库等原因导致的库存差价不准确时的调整处理（只允许一种入库类别）用于整进零出和负数出库等原因导致的库存差价不准确时的调整处理（只允许一种入库类别）用于整进零出和负数出库等原因导致的库存差价不准确时的调整处理（只允许一种入库类别）用于整进零出和负数出库等原因导致的库存差价不准确时的调整处理（只允许一种入库类别）用于整进零出和负数出库等原因导致的库存差价不准确时的调整处理（只允许一种入库类别）" & _
'                        "用于整进零出和负数出库等原因导致的库存差价不准确时的调整处理（只允许一种入库类别）用于整进零出和负数出库等原因导致的库存差价不准确时的调整处理（只允许一种入库类别）用于整进零出和负数出库等原因导致的库存差价不准确时的调整处理（只允许一种入库类别）用于整进零出和负数出库等原因导致的库存差价不准确时的调整处理（只允许一种入库类别）用于整进零出和负数出库等原因导致的库存差价不准确时的调整处理（只允许一种入库类别）用于整进零出和负数出库等原因导致的库存差价不准确时的调整处理（只允许一种入库类别）用于整进零出和负数出库等原因导致的库存差价不准确时的调整处理（只允许一种入库类别）用于整进零出和负数出库等原因导致的库存差价不准确时的调整处理（只允许一种入库类别）用于整进零出和负数出库等原因导致的库存差价不准确时的调整处理（只允许一种入库类别）用于整进零出和负数出库等原因导致的库存差价不准确时的调整处理（只允许一种入库类别）用于整进零出和负数出库等原因导致的库存差价不准确时的调整处理（只允许一种入库类别）"
'    lbl(4).Caption = "项目说明:" & msfResult.TextMatrix(msfResult.Row, mCol.项目说明)
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub picBack_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub picKey_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub tmrInfo_Timer()
    If mvarStop > 0 Then
        mvarStop = mvarStop - 1
    Else
        tmrScroll.Enabled = True
    End If
End Sub

Private Sub tmrScroll_Timer()
    
    If mvarScroll > 0 Then
        mvarScroll = mvarScroll - 1
    Else
        If UsrCmd(11).Enabled Then
            If UsrCmd(1).Caption <> "病人自助查询" Then
                Call UsrCmd_CommandClick(1)
            End If
            Call UsrCmd_CommandClick(11)
        Else
            If UsrCmd(1).Caption <> "病人自助查询" Then
                Call UsrCmd_CommandClick(1)
            End If
            Call SearchItem("")
        End If
        mvarScroll = Val(GetPara("价格查询滚动间隔", "10"))
        mvarScroll = IIf(mvarScroll <= 0, 10, mvarScroll)
    End If
End Sub

Private Sub txt_GotFocus(Index As Integer)
    If Index = 0 Then
        mblnNumber = True
    Else
        mblnNumber = False
    End If
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then Exit Sub
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If CheckIsInclude(UCase(Chr(KeyAscii)), "'") = True Then KeyAscii = 0
End Sub

Private Sub txt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    tmrScroll.Enabled = False
    mvarStop = Val(GetPara("价格查询停留时间", "30"))
    mvarStop = IIf(mvarStop <= 0, 30, mvarStop)
    
    Select Case KeyCode
    Case vbKeyA
        txt(1).Text = txt(1).Text & "A"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyB
        txt(1).Text = txt(1).Text & "B"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyC
        txt(1).Text = txt(1).Text & "C"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyD
        txt(1).Text = txt(1).Text & "D"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyE
        txt(1).Text = txt(1).Text & "E"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyF
        txt(1).Text = txt(1).Text & "F"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyG
        txt(1).Text = txt(1).Text & "G"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyH
        txt(1).Text = txt(1).Text & "H"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyI
        txt(1).Text = txt(1).Text & "I"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyJ
        txt(1).Text = txt(1).Text & "J"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyK
        txt(1).Text = txt(1).Text & "K"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyL
        txt(1).Text = txt(1).Text & "L"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyM
        txt(1).Text = txt(1).Text & "M"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyN
        txt(1).Text = txt(1).Text & "N"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyO
        txt(1).Text = txt(1).Text & "O"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyP
        txt(1).Text = txt(1).Text & "P"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyQ
        txt(1).Text = txt(1).Text & "Q"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyR
        txt(1).Text = txt(1).Text & "R"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyS
        txt(1).Text = txt(1).Text & "S"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyT
        txt(1).Text = txt(1).Text & "T"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyU
        txt(1).Text = txt(1).Text & "U"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyV
        txt(1).Text = txt(1).Text & "V"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyW
        txt(1).Text = txt(1).Text & "W"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyX
        txt(1).Text = txt(1).Text & "X"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyY
        txt(1).Text = txt(1).Text & "Y"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyZ
        txt(1).Text = txt(1).Text & "Z"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKey0, vbKeyNumpad0
        txt(0).Text = txt(0).Text & "0"
    Case vbKey1, vbKeyNumpad1
        txt(0).Text = txt(0).Text & "1"
    Case vbKey2, vbKeyNumpad2
        txt(0).Text = txt(0).Text & "2"
    Case vbKey3, vbKeyNumpad3
        txt(0).Text = txt(0).Text & "3"
    Case vbKey4, vbKeyNumpad4
        txt(0).Text = txt(0).Text & "4"
    Case vbKey5, vbKeyNumpad5
        txt(0).Text = txt(0).Text & "5"
    Case vbKey6, vbKeyNumpad6
        txt(0).Text = txt(0).Text & "6"
    Case vbKey7, vbKeyNumpad7
        txt(0).Text = txt(0).Text & "7"
    Case vbKey8, vbKeyNumpad8
        txt(0).Text = txt(0).Text & "8"
    Case vbKey9, vbKeyNumpad9
        txt(0).Text = txt(0).Text & "9"
    Case vbKeyReturn, vbKeySeparator
        If txt(0).Text <> "" Then
            Call cmdBtn_Click(10)
        Else
            Call cmdKey_CommandClick("确定")
        End If
    Case vbKeyDecimal
        Call cmdBtn_Click(11)
    Case vbKeyDelete
        Call cmdKey_CommandClick("清除")
        Call cmdBtn_Click(12)
    End Select
    
    If KeyCode <> 27 Then KeyCode = 0
    
    If KeyCode = 27 Then
        RaiseEvent KeyDown(KeyCode, Shift)
    End If
    
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    
    Dim vHeight As Single
    
    If UsrCmd(1).Caption = "隐藏查询区" Then
        vHeight = picKey.Height + picBack(0).Height + 45
        txt(0).Visible = True
        txt(1).Visible = True
        lbl(0).Visible = True
        lbl(2).Visible = True
        UsrCmd(0).Visible = True
        UsrCmd(3).Visible = True
    Else
        vHeight = picBack(0).Height
        txt(0).Visible = False
        txt(1).Visible = False
        lbl(0).Visible = False
        lbl(2).Visible = False
        UsrCmd(0).Visible = False
        UsrCmd(3).Visible = False
    End If
    
    Call ResizeControl(picBack(1), 0, 0, UserControl.Width, picBack(1).Height)
    Call ResizeControl(msfResult, 0, picBack(1).Top + picBack(1).Height + 15, UserControl.Width - 15, UserControl.Height - picBack(1).Height - vHeight)
    Call ResizeControl(picBack(0), 0, msfResult.Top + msfResult.Height + 15, UserControl.Width, picBack(0).Height)

    
'
    Call ResizeControl(msfCalc, pic.Width, picBack(0).Top + picBack(0).Height + 15, UserControl.Width - picKey.Width - pic.Width, picKey.Height - picBack(2).Height - picBack(3).Height)
    Call ResizeControl(picBack(3), msfCalc.Left, msfCalc.Top + msfCalc.Height, msfCalc.Width, picBack(3).Height)
    Call ResizeControl(pic, 0, msfCalc.Top, pic.Width, pic.Height)
    
    Call ResizeControl(picKey, msfCalc.Left + msfCalc.Width + 30, msfCalc.Top, picKey.Width, picKey.Height)
    Call ResizeControl(picBack(2), pic.Left, picBack(3).Top + picBack(3).Height, msfCalc.Width + pic.Width, picBack(2).Height)
    
    
    
    UsrCmd(2).Left = picBack(1).ScaleWidth - UsrCmd(2).Width - 30
    UsrCmd(11).Left = UsrCmd(2).Left - UsrCmd(11).Width - 30
    UsrCmd(10).Left = UsrCmd(11).Left - UsrCmd(10).Width - 30

    UsrCmd(13).Left = picBack(2).ScaleWidth - UsrCmd(13).Width - 30
    UsrCmd(12).Left = UsrCmd(13).Left - UsrCmd(12).Width - 30

    lbl(2).Left = picKey.Left
    

    txt(0).Left = lbl(0).Left + lbl(0).Width
    txt(0).Top = lbl(0).Top - 60
    txt(1).Top = UsrCmd(1).Top
    lbl(2).Top = txt(1).Top + 60
    lbl(3).Left = lbl(2).Left

    UsrCmd(1).Left = picBack(0).ScaleWidth - UsrCmd(1).Width - 30
    txt(1).Left = UsrCmd(1).Left - txt(1).Width - 30
    lbl(2).Left = txt(1).Left - lbl(2).Width - 15
    
    
    fra.Move lbl(3).Left - 30, fra.Top, picKey.Width
    fra2.Move lbl(3).Left - 30, -90, fra2.Width, picBack(0).Height + 90
    
    lbl(4).Move 0, 0, fra.Left - UsrCmd(14).Width - 30 - 60, picBack(0).Height
    
    UsrCmd(14).Left = fra.Left - UsrCmd(14).Width - 60
    
    UsrCmd(15).Left = fra.Left - UsrCmd(15).Width - 60
    
End Sub

Private Sub DrawMsfHeader()

    msfResult.Cols = 0
        
    Call AddColumn(msfResult, "编码", 1080, 1)
    Call AddColumn(msfResult, "名称", 4020, 1)
    Call AddColumn(msfResult, "规格", 2700, 1)
    Call AddColumn(msfResult, "剂型", 900, 1)
    Call AddColumn(msfResult, "单位", 600, 1)
    Call AddColumn(msfResult, "价格", 1800, 7)
    Call AddColumn(msfResult, "指导售价", 1800, 7)
    
    Call AddColumn(msfResult, "产地", 2100, 1)
    Call AddColumn(msfResult, "标识主码", 1080, 1)
    Call AddColumn(msfResult, "标识子码", 1080, 1)
    
    Call AddColumn(msfResult, "费用类别", 1200, 1)
    Call AddColumn(msfResult, "价格依据", 0, 1)
    Call AddColumn(msfResult, "项目说明", 0, 1)
    Call AddColumn(msfResult, "", 1200, 1)
    
    Call RestoreFlexState(msfResult, App.ProductName)
    
    Dim strTmp As String
    
    strTmp = Trim(zlDatabase.GetPara("价格显示信息", glngSys, 1536, "0000011"))
    If Len(strTmp) = 6 Then strTmp = strTmp & "1"
    
    
    If Val(Mid(strTmp, 1, 1)) = 1 Then msfResult.ColHidden(mCol.费用类型) = True
    If Val(Mid(strTmp, 2, 1)) = 1 Then msfResult.ColHidden(mCol.编码) = True
    If Val(Mid(strTmp, 3, 1)) = 1 Then msfResult.ColHidden(mCol.产地) = True
    If Val(Mid(strTmp, 4, 1)) = 1 Then msfResult.ColHidden(mCol.标识主码) = True
    If Val(Mid(strTmp, 5, 1)) = 1 Then msfResult.ColHidden(mCol.标识子码) = True
    If Val(Mid(strTmp, 6, 1)) = 1 Then msfResult.ColHidden(mCol.指导售价) = True
    If Val(Mid(strTmp, 7, 1)) = 1 Then msfResult.ColHidden(mCol.剂型) = True
    
    msfCalc.Cols = 0
    Call AddColumn(msfCalc, "名称", 3030, 1)
    Call AddColumn(msfCalc, "数量", 540, 7)
    Call AddColumn(msfCalc, "金额", 810, 7)
    Call AddColumn(msfCalc, "", 15, 1)
    
    Call CalcAutoColWidth(msfResult, mCol.名称)
    Call CalcAutoColWidth(msfCalc, 0)
           
End Sub


Private Sub UserControl_Show()
    cmdKey.KeyMode = 1
End Sub

Private Sub UsrCmd_CommandClick(Index As Integer)
    Dim i As Long
    
    If Index >= 4 And Index < 10 Then
        UsrCmd(Index).Picture = ilsImage.ListImages(IIf(UsrCmd(Index).Tag = "1", "unselect", "select"))
        UsrCmd(Index).Tag = IIf(UsrCmd(Index).Tag = "1", "", "1")
        DoEvents
        Call SearchItem(txt(1).Text)
        Exit Sub
    End If
    
    Select Case Index
    Case 0
        If msfCalc.Row < 1 Then Exit Sub
        If msfCalc.TextMatrix(msfCalc.Row, 0) <> "" Then
            mvarRows2 = mvarRows2 - 1
            mvarCurPos2 = 1
            If msfCalc.Rows <= 2 Then
                Call ClearSpecRowCol(msfCalc, 1, Array())
            Else
                msfCalc.RemoveItem msfCalc.Row
            End If
            'msfCalc.Rows = mvarRows2 + 10
            Call CalcMoney
            Call EnablePageButton(msfCalc, mvarCurPos2, mvarRows2, UsrCmd(12), UsrCmd(13))
        End If
    Case 1
        If UsrCmd(Index).Caption = "隐藏查询区" Then
            '隐藏查询区
            UsrCmd(Index).Caption = "病人自助查询"
        Else
            UsrCmd(Index).Caption = "隐藏查询区"
        End If
        Call UserControl_Resize
    Case 2
        Call frmHelp.ShowHelp(Me, -1, UserControl.Width, UserControl.Height)
    Case 3
        Call cmdBtn_Click(10)
    Case 10
        Call TurnToPage(msfResult, -1, mvarCurPos1)
        Call EnablePageButton(msfResult, mvarCurPos1, mvarRows1, UsrCmd(10), UsrCmd(11))
    Case 11             '下一页
        Call TurnToPage(msfResult, 1, mvarCurPos1)
        Call EnablePageButton(msfResult, mvarCurPos1, mvarRows1, UsrCmd(10), UsrCmd(11))
    Case 12
        Call TurnToPage(msfCalc, -1, mvarCurPos2)
        Call EnablePageButton(msfCalc, mvarCurPos2, mvarRows2, UsrCmd(12), UsrCmd(13))
    Case 13             '下一页
        Call TurnToPage(msfCalc, 1, mvarCurPos2)
        Call EnablePageButton(msfCalc, mvarCurPos2, mvarRows2, UsrCmd(12), UsrCmd(13))
    Case 14
        
        If (lbl(4).Top + lbl(4).Height) > picBack(0).Height Then lbl(4).Top = lbl(4).Top - 210
    Case 15
        If lbl(4).Top < 0 Then lbl(4).Top = lbl(4).Top + 210
    End Select
End Sub


Private Sub SearchItem(ByVal strKey As String)
    Dim strLike As String
    Dim strInput As String
    Dim strSort As String
    Dim i As Long
    Dim lngSvrRow As Long
    Dim sgl价格 As Single
    Dim rs As New ADODB.Recordset
    Dim lngBkColor As Long
    Dim strTmp As String
    Dim blnAllow As Boolean
    
    On Error GoTo errHand
    
    lngBkColor = 15987699
    
    strTmp = GetPara("允许显示的收费类别")
    
    strSort = " 类别='999' "
    If UsrCmd(4).Tag = "1" Then strSort = strSort & " OR 类别='5' OR 类别='6' OR 类别='7'"
    If UsrCmd(5).Tag = "1" Then strSort = strSort & " OR 类别='C'"
    If UsrCmd(6).Tag = "1" Then strSort = strSort & " OR 类别='D'"
    If UsrCmd(7).Tag = "1" Then strSort = strSort & " OR 类别='E'"
    If UsrCmd(8).Tag = "1" Then strSort = strSort & " OR 类别='F'"
    
    If mrsPrice.State <> adStateOpen Then
        gstrSQL = "Select a.药品id,Sum(实际金额)/Sum(实际数量) As 均价 from 药品库存 a,收费项目目录 b where a.药品id=b.ID And Nvl(b.是否变价,0)=1 And " & GetNodeCheckSQL("b.站点") & " Group By a.药品id Having Sum(实际数量)<>0"
        Set mrsPrice = zlDatabase.OpenSQLRecord(gstrSQL, "价格查询")
    End If
    
    If UsrCmd(9).Tag = "1" Then
        '其他，这里的其他是指除了前面的几个类别外的所有类别
        strSort = strSort & " OR (类别<>'5' AND 类别<>'6' AND 类别<>'7' AND 类别<>'C' AND 类别<>'D' AND 类别<>'E' AND 类别<>'F')"
    End If

    
    strSearchSQL = ""
    If strKey <> "" Then
        strInput = "%" & strKey & "%"
        strSearchSQL = " AND (Y.编码 Like [1] OR Upper(Y.名称) Like [1] OR Y.ID IN (SELECT 收费细目ID From 收费项目别名 WHERE UPPER(简码) Like [1] OR Upper(名称) Like [1]))"
    End If
            
    '非变价的收收费项目,药品变价项目，显示价格为药品指导零售价,其他变价项目不显示
    gstrSQL = "" & _
        "Select A.类别,A.编码,A.名称,A.规格,A.剂型,A.单位,A.产地,A.标识子码,A.标识主码,A.价格依据,A.费用类型,A.项目说明,A.从项数次,A.是否从属,A.是否变价,A.现价,A.指导零售价,A.从项ID,A.主项ID,A.分类ID,DECODE(现价,0,NULL,DECODE(a.从项数次, 0, 1, NULL, 1, a.从项数次) * a.现价) AS 价格 " & _
        "From ( " & _
        "Select 主项类别 As 类别, " & _
               "编码, " & _
               "名称, " & _
               "规格,剂型, " & _
               "单位,产地,标识子码,标识主码, " & _
               "价格依据, " & _
               "费用类型, " & _
               "项目说明, " & _
               "从项数次, " & _
               "Decode(从项id,0,0,1) As 是否从属, " & _
               "Decode(类别, '5', 1, '6', 1, '7', 1, 0) *Decode(是否变价, 1, 1, 0) As 是否变价, " & _
               "Decode(Decode(类别, '5', 1, '6', 1, '7', 1, 0) * Decode(是否变价, 1, 1, 0),1,指导零售价,现价) as 现价,指导零售价,从项id,主项id,分类id " & _
        "From ( " & _
        "Select X.主项id,X.从项id, " & _
               "Y.类别 As 主项类别, " & _
               "Decode(X.从项id,0,Y.类别,Z.类别) As 类别, " & _
               "Decode(X.从项id,0,Y.编码,Z.编码) As 编码, " & _
               "Decode(X.从项id,0,Y.名称,'  '||Z.名称||X.从项说明) As 名称, " & _
               "Decode(X.从项id,0,Y.规格,Z.规格) As 规格, " & _
               "Decode(X.从项id,0,Y.计算单位,Z.计算单位) As 单位, " & _
               "Decode(X.从项id,0,Y.标识主码,Z.标识主码) As 标识主码,Decode(X.从项id,0,Y.标识子码,Z.标识子码) As 标识子码, " & _
               "P.现价,y.产地,m.剂型, "
                       
    gstrSQL = gstrSQL & _
               "p.指导零售价," & _
               "x.从项数次," & _
               "x.从项说明," & _
               "P.调价说明 AS 价格依据," & _
               "Decode(X.从项id,0,Y.费用类型,Z.费用类型) As 费用类型," & _
               "Decode(X.从项id,0,Y.说明,Z.说明) As 项目说明," & _
               "Decode(X.从项id, 0, Y.是否变价, Z.是否变价) As 是否变价,Decode(m.分类id,Null,'P'||To_Char(y.分类id),'K'||To_Char(m.分类id)) As 分类id " & _
        "From ( Select ID,ID As 主项id,0 As 从项id,'' As 从项说明,0 As 从项数次 From 收费项目目录 Where 服务对象 In (1,2,3) And " & GetNodeCheckSQL("站点") & " " & _
               "Union All " & _
               "Select 从项id AS ID,主项ID,从项id,DECODE(固有从属, 2, '[比', 1, '[固', '[活') ||to_char(从项数次) || ']' AS 从项说明,从项数次 From 收费从属项目 a,收费项目目录 b where b.id=a.从项id and b.服务对象 In (1,2,3) And " & GetNodeCheckSQL("b.站点") & " " & _
              ") x, " & _
              "( Select k.收费细目ID,K.调价说明,K.现价,Decode(t.指导零售价,Null,s.指导零售价,t.指导零售价) As 指导零售价 " & _
                "From 药品规格 t,材料特性 s," & _
                     "( Select 收费细目ID,调价说明,Sum(现价) As 现价 " & _
                     "From 收费价目 " & _
                     "Where (终止日期 is Null OR 终止日期 = TO_DATE('3000-01-01', 'YYYY-MM-DD')) Group By 收费细目ID,调价说明 " & _
                     ") k " & _
                "Where k.收费细目ID=t.药品id(+) And k.收费细目ID=s.材料ID(+) " & _
              ") p, " & _
              "收费项目目录 y, " & _
              "收费项目目录 z, " & _
              "(Select t.药品id,w.分类id,p.药品剂型 As 剂型 From 药品规格 t,诊疗项目目录 w,药品特性 p Where t.药名id=w.ID And p.药名id=w.ID And " & GetNodeCheckSQL("w.站点") & ") m " & _
        "Where x.主项ID = y.ID and y.服务对象 In (1,2,3) And " & GetNodeCheckSQL("y.站点") & " And " & GetNodeCheckSQL("z.站点") & " AND x.ID=p.收费细目ID(+) AND x.从项id=z.ID(+) And y.ID=m.药品id(+) "
    
    If strTmp <> "" Then
        gstrSQL = gstrSQL & " And y.类别 In (" & strTmp & ")"
    End If
    
    gstrSQL = gstrSQL & _
              "AND (Y.撤档时间 is null OR Y.撤档时间 = TO_DATE('3000-01-01', 'YYYY-MM-DD')) " & _
              "AND (Z.撤档时间 is null OR Z.撤档时间 = TO_DATE('3000-01-01', 'YYYY-MM-DD')) " & strSearchSQL & _
        "Order By Y.名称,X.从项id " & _
        ")) a"

    mvarCurPos1 = 1
    mvarRows1 = 0
    UsrCmd(10).Enabled = False
    UsrCmd(11).Enabled = False
                
    i = 1
    msfResult.Rows = 2
    Call ClearSpecRowCol(msfResult, 1, Array())
    
    ShowFlatFlash "正在提取价格信息..."
    DoEvents
    
    If strKey = "" Then
        If mvarRs.State = adStateOpen Then
            mvarRs.Filter = ""
            mvarRs.Filter = strSort
        Else
            Set mvarRs = zlDatabase.OpenSQLRecord(gstrSQL, "价格查询")
            mvarRs.Filter = strSort
        End If
        Set rs = mvarRs
    Else
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "价格查询", strInput)
    End If
    
    If rs.BOF = False Then
        Do While Not rs.EOF
            blnAllow = False
            mrs分类id.Filter = ""
            If mrs分类id.RecordCount > 0 And zlCommFun.Nvl(rs("分类id").Value) <> "P" Then
                If mblnUnSelect Then
                    mrs分类id.Filter = "分类id='" & zlCommFun.Nvl(rs("分类id").Value) & "'"
                    blnAllow = Not (mrs分类id.RecordCount > 0)
                Else
                    mrs分类id.Filter = "分类id='" & zlCommFun.Nvl(rs("分类id").Value) & "'"
                    blnAllow = (mrs分类id.RecordCount > 0)
                End If
            Else
                blnAllow = True
            End If
            
            If blnAllow Then
                msfResult.TextMatrix(i, mCol.编码) = IIf(IsNull(rs!编码), "", rs!编码)
                msfResult.TextMatrix(i, mCol.名称) = IIf(IsNull(rs!名称), "", rs!名称)
                msfResult.TextMatrix(i, mCol.规格) = IIf(IsNull(rs!规格), "", rs!规格)
                msfResult.TextMatrix(i, mCol.剂型) = IIf(IsNull(rs!剂型), "", rs!剂型)
                msfResult.TextMatrix(i, mCol.单位) = IIf(IsNull(rs!单位), "", rs!单位)
                msfResult.TextMatrix(i, mCol.产地) = IIf(IsNull(rs!产地), "", rs!产地)
                msfResult.TextMatrix(i, mCol.标识主码) = IIf(IsNull(rs!标识主码), "", rs!标识主码)
                msfResult.TextMatrix(i, mCol.标识子码) = IIf(IsNull(rs!标识子码), "", rs!标识子码)
                msfResult.TextMatrix(i, mCol.价格) = IIf(IsNull(rs!价格), "", Format(rs!价格, "0.00##"))
                msfResult.TextMatrix(i, mCol.指导售价) = IIf(IsNull(rs!指导零售价), "", Format(rs!指导零售价, "0.00##"))
                msfResult.TextMatrix(i, mCol.费用类型) = IIf(IsNull(rs!费用类型), "", rs!费用类型)
                msfResult.TextMatrix(i, mCol.价格依据) = IIf(IsNull(rs!价格依据), "", rs!价格依据)
                msfResult.TextMatrix(i, mCol.项目说明) = IIf(IsNull(rs!项目说明), "", rs!项目说明)
                msfResult.RowData(i) = Val(msfResult.TextMatrix(i, mCol.价格))
                
                '计算时价药品的价格
                msfResult.Cell(flexcpData, i, mCol.价格, i, mCol.价格) = 0
                Select Case zlCommFun.Nvl(rs("类别").Value)
                Case "4", "5", "6", "7"
                    '药品,材料
                    
                    If zlCommFun.Nvl(rs("是否变价").Value, 0) = 1 Then
                        msfResult.Cell(flexcpData, i, mCol.价格, i, mCol.价格) = 1
                        mrsPrice.Filter = ""
                        
                        If zlCommFun.Nvl(rs("从项id").Value, 0) > 0 Then
                            mrsPrice.Filter = "药品id=" & zlCommFun.Nvl(rs("从项id").Value, 0)
                        Else
                            mrsPrice.Filter = "药品id=" & zlCommFun.Nvl(rs("主项id").Value, 0)
                        End If
                        
                        If mrsPrice.RecordCount > 0 Then
                            If zlCommFun.Nvl(mrsPrice("均价").Value, 0) > 0 Then
                                msfResult.TextMatrix(i, mCol.价格) = Format(zlCommFun.Nvl(mrsPrice("均价").Value, 0), "0.00##")
                                msfResult.Cell(flexcpData, i, mCol.价格, i, mCol.价格) = 0
                            End If
                        End If
                    End If
                End Select
            
            
                If zlCommFun.Nvl(rs("是否从属"), 0) = 1 Then
                
                    '累加价格
                    sgl价格 = sgl价格 + Val(msfResult.TextMatrix(i, mCol.价格))
                    
                Else
                    If sgl价格 > 0 Then
                        If InStr(msfResult.TextMatrix(lngSvrRow, mCol.价格), "(指导价)") > 0 Then
                            msfResult.TextMatrix(lngSvrRow, mCol.价格) = Left(msfResult.TextMatrix(lngSvrRow, mCol.价格), Len(msfResult.TextMatrix(lngSvrRow, mCol.价格)) - 5)
                            msfResult.TextMatrix(lngSvrRow, mCol.价格) = Format(Val(msfResult.TextMatrix(lngSvrRow, mCol.价格)) + sgl价格, "0.00##")
                            If Val(msfResult.TextMatrix(lngSvrRow, mCol.价格)) = 0 Then
                                msfResult.TextMatrix(lngSvrRow, mCol.价格) = ""
                            Else
                                msfResult.TextMatrix(lngSvrRow, mCol.价格) = msfResult.TextMatrix(lngSvrRow, mCol.价格) & "(指导价)"
                            End If
                        Else
                            msfResult.TextMatrix(lngSvrRow, mCol.价格) = Format(Val(msfResult.TextMatrix(lngSvrRow, mCol.价格)) + sgl价格, "0.00##")
                        End If
                        
                        sgl价格 = 0
                    End If
                    lngSvrRow = i
                    
                End If
            
                If i Mod 2 = 0 Then msfResult.Cell(flexcpBackColor, i, 0, i, msfResult.Cols - 1) = lngBkColor
    
                If msfResult.Cell(flexcpData, i, mCol.价格, i, mCol.价格) = 1 Then
                    msfResult.TextMatrix(i, mCol.价格) = msfResult.TextMatrix(i, mCol.价格) & "(指导价)"
                End If
                
                i = i + 1
                msfResult.Rows = i + 1
            End If
            
            rs.MoveNext
            If msfResult.Rows = 30 Then DoEvents
        Loop
        
        If sgl价格 > 0 Then
            If InStr(msfResult.TextMatrix(lngSvrRow, mCol.价格), "(指导价)") > 0 Then
                msfResult.TextMatrix(lngSvrRow, mCol.价格) = Left(msfResult.TextMatrix(lngSvrRow, mCol.价格), Len(msfResult.TextMatrix(lngSvrRow, mCol.价格)) - 5)
                msfResult.TextMatrix(lngSvrRow, mCol.价格) = Format(Val(msfResult.TextMatrix(lngSvrRow, mCol.价格)) + sgl价格, "0.00##")
                If Val(msfResult.TextMatrix(lngSvrRow, mCol.价格)) = 0 Then
                    msfResult.TextMatrix(lngSvrRow, mCol.价格) = ""
                Else
                    msfResult.TextMatrix(lngSvrRow, mCol.价格) = msfResult.TextMatrix(lngSvrRow, mCol.价格) & "(指导价)"
                End If
            Else
                msfResult.TextMatrix(lngSvrRow, mCol.价格) = Format(Val(msfResult.TextMatrix(lngSvrRow, mCol.价格)) + sgl价格, "0.00##")
            End If
            
            sgl价格 = 0
        End If
        If msfResult.Rows > 2 Then msfResult.Rows = msfResult.Rows - 1
        mvarCurPos1 = 1
        mvarRows1 = msfResult.Rows - 1
    End If
        
    StopFlatFlash
    
    msfResult.Rows = msfResult.Rows + 50
    Call EnablePageButton(msfResult, mvarCurPos1, mvarRows1, UsrCmd(10), UsrCmd(11))
    Call msfResult_RowColChange
    
    Exit Sub
    
errHand:
    StopFlatFlash
End Sub

Private Sub CalcMoney()
    Dim i As Long
    Dim vTmp As Single

    vTmp = 0
    For i = 1 To msfCalc.Rows - 1
        vTmp = vTmp + Val(msfCalc.TextMatrix(i, 2))
    Next
    lbl(1).Caption = "合计:" & Format(vTmp, "0.00元")
End Sub

Private Sub UsrCmd_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Public Property Let Enabled(ByVal vData As Boolean)
    UserControl.Enabled = vData
End Property

