VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frm清单管理 
   Caption         =   "应付款查询"
   ClientHeight    =   6255
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9630
   Icon            =   "frm清单管理.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chk包含付款 
      Caption         =   "包含重置时间段之后付款的未付清单"
      Height          =   255
      Left            =   6720
      TabIndex        =   36
      Top             =   2400
      Width           =   3495
   End
   Begin VB.PictureBox picFind 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6165
      ScaleHeight     =   315
      ScaleWidth      =   3945
      TabIndex        =   33
      Top             =   2835
      Width           =   3945
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   1365
         TabIndex        =   35
         Top             =   0
         Width           =   1785
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         Caption         =   "按名称查找"
         Height          =   180
         Left            =   420
         TabIndex        =   34
         Top             =   60
         Width           =   900
      End
      Begin VB.Image imgSearch 
         Height          =   360
         Left            =   15
         Picture         =   "frm清单管理.frx":08CA
         Top             =   -15
         Width           =   360
      End
   End
   Begin MSComctlLib.ImageList ilt24 
      Left            =   2505
      Top             =   825
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清单管理.frx":0FB4
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清单管理.frx":16AE
            Key             =   "FindH"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsList 
      Height          =   2100
      Left            =   2790
      TabIndex        =   31
      Top             =   3195
      Width           =   6780
      _cx             =   11959
      _cy             =   3704
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
      BackColorFixed  =   -2147483644
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483644
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   12632256
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   12632256
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm清单管理.frx":1DA8
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
   Begin VB.Frame fraSearch 
      Height          =   5010
      Left            =   30
      TabIndex        =   17
      Top             =   765
      Visible         =   0   'False
      Width           =   2370
      Begin VB.Frame fraSplit 
         Height          =   30
         Left            =   0
         TabIndex        =   30
         Top             =   465
         Width           =   2355
      End
      Begin VB.PictureBox picClear 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   45
         ScaleHeight     =   315
         ScaleWidth      =   750
         TabIndex        =   29
         ToolTipText     =   "清除当前搜索，重新开始!"
         Top             =   525
         Width           =   750
         Begin VB.Image img清除 
            Height          =   285
            Left            =   0
            Picture         =   "frm清单管理.frx":1DE4
            Top             =   15
            Width           =   300
         End
      End
      Begin VB.PictureBox picHelp 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   960
         ScaleHeight     =   330
         ScaleWidth      =   435
         TabIndex        =   28
         ToolTipText     =   "清除当前搜索，重新开始!"
         Top             =   510
         Width           =   435
         Begin VB.Image imgHelp 
            Height          =   270
            Left            =   45
            Picture         =   "frm清单管理.frx":229A
            Top             =   30
            Width           =   300
         End
      End
      Begin VB.PictureBox PicClose 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2010
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   27
         ToolTipText     =   "清除当前搜索，重新开始!"
         Top             =   135
         Width           =   300
      End
      Begin VB.PictureBox PicSearchBack 
         BackColor       =   &H8000000E&
         Height          =   4125
         Left            =   0
         ScaleHeight     =   4065
         ScaleWidth      =   2280
         TabIndex        =   18
         Top             =   870
         Width           =   2340
         Begin VB.PictureBox picSearch 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   4530
            Left            =   0
            ScaleHeight     =   4530
            ScaleWidth      =   2205
            TabIndex        =   20
            Top             =   -135
            Width           =   2205
            Begin VB.TextBox txt编码 
               Height          =   300
               Left            =   30
               MaxLength       =   13
               TabIndex        =   3
               Top             =   420
               Width           =   1965
            End
            Begin VB.CommandButton cmd搜索 
               Caption         =   "立即过滤(&S)"
               Height          =   350
               Left            =   45
               TabIndex        =   6
               Top             =   1380
               Width           =   1245
            End
            Begin VB.TextBox Txt名称 
               Height          =   300
               Left            =   45
               MaxLength       =   50
               TabIndex        =   5
               Top             =   960
               Width           =   1935
            End
            Begin VB.TextBox TxtOther 
               Height          =   300
               Index           =   0
               Left            =   240
               MaxLength       =   50
               TabIndex        =   23
               Top             =   2400
               Visible         =   0   'False
               Width           =   1710
            End
            Begin MSComCtl2.DTPicker DtpOther 
               Height          =   300
               Index           =   0
               Left            =   120
               TabIndex        =   21
               Top             =   3525
               Visible         =   0   'False
               Width           =   1740
               _ExtentX        =   3069
               _ExtentY        =   529
               _Version        =   393216
               CheckBox        =   -1  'True
               CustomFormat    =   "yyyy年MM月dd日"
               DateIsNull      =   -1  'True
               Format          =   292356099
               CurrentDate     =   37131
            End
            Begin VB.CheckBox chkOther 
               BackColor       =   &H8000000E&
               Caption         =   "末级"
               Height          =   225
               Index           =   0
               Left            =   165
               TabIndex        =   22
               Top             =   2820
               Visible         =   0   'False
               Width           =   1515
            End
            Begin VB.Label lbl 
               BackStyle       =   0  'Transparent
               Caption         =   "供应商编码"
               Height          =   180
               Index           =   0
               Left            =   45
               TabIndex        =   2
               Top             =   195
               Width           =   1980
            End
            Begin VB.Label lblHit 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "其他选项>>"
               ForeColor       =   &H8000000D&
               Height          =   240
               Left            =   45
               TabIndex        =   26
               Top             =   1800
               Width           =   1905
            End
            Begin VB.Shape shpHit 
               Height          =   2505
               Left            =   45
               Top             =   1785
               Visible         =   0   'False
               Width           =   1920
            End
            Begin VB.Label lbl 
               BackStyle       =   0  'Transparent
               Caption         =   "供应商名称或简码"
               Height          =   180
               Index           =   2
               Left            =   60
               TabIndex        =   4
               Top             =   735
               Width           =   1980
            End
            Begin VB.Label lblOther 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "简码"
               Height          =   180
               Index           =   0
               Left            =   150
               TabIndex        =   25
               Top             =   2145
               Visible         =   0   'False
               Width           =   360
            End
            Begin VB.Label lblDate 
               BackStyle       =   0  'Transparent
               Caption         =   "建档时间"
               Height          =   240
               Index           =   0
               Left            =   150
               TabIndex        =   24
               Top             =   3180
               Visible         =   0   'False
               Width           =   1695
            End
         End
         Begin VB.VScrollBar Scr 
            Height          =   4125
            Left            =   2055
            TabIndex        =   19
            Top             =   0
            Width           =   225
         End
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "过滤条件"
         Height          =   165
         Index           =   1
         Left            =   60
         TabIndex        =   1
         Top             =   195
         Width           =   1860
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         DrawMode        =   9  'Not Mask Pen
         X1              =   900
         X2              =   900
         Y1              =   555
         Y2              =   795
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000006&
         DrawMode        =   16  'Merge Pen
         X1              =   915
         X2              =   915
         Y1              =   540
         Y2              =   810
      End
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   2595
      Top             =   3330
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清单管理.frx":2714
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清单管理.frx":2B6C
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清单管理.frx":2FC4
            Key             =   "ItemNo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清单管理.frx":3418
            Key             =   "No"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清单管理.frx":3870
            Key             =   "Write"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip tabSelect 
      Height          =   315
      Left            =   2850
      TabIndex        =   16
      Top             =   2820
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   556
      TabWidthStyle   =   1
      MultiRow        =   -1  'True
      Style           =   2
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "付款明细帐"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "已付清单"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "未付清单"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwList 
      Height          =   5355
      Left            =   0
      TabIndex        =   0
      Top             =   750
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   9446
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   5895
      Width           =   9630
      _ExtentX        =   16986
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm清单管理.frx":3CC8
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11906
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
   Begin MSComctlLib.ImageList ilsCold 
      Left            =   6150
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清单管理.frx":455C
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清单管理.frx":477C
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清单管理.frx":499C
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清单管理.frx":4BB8
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清单管理.frx":4DD8
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清单管理.frx":4FF8
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清单管理.frx":5214
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清单管理.frx":5430
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清单管理.frx":564A
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清单管理.frx":57A4
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清单管理.frx":59C0
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清单管理.frx":5BE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清单管理.frx":5DFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清单管理.frx":6014
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清单管理.frx":622E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHot 
      Left            =   6750
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清单管理.frx":6448
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清单管理.frx":6668
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清单管理.frx":6888
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清单管理.frx":6AA4
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清单管理.frx":6CC4
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清单管理.frx":6EE4
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清单管理.frx":7100
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清单管理.frx":731C
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清单管理.frx":7536
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清单管理.frx":7690
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清单管理.frx":78B0
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清单管理.frx":7AD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清单管理.frx":7CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清单管理.frx":7F04
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清单管理.frx":811E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrTool 
      Height          =   780
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   1376
      BandCount       =   2
      BandBorders     =   0   'False
      _CBWidth        =   9570
      _CBHeight       =   780
      _Version        =   "6.7.8988"
      Child1          =   "tlbThis"
      MinHeight1      =   720
      Width1          =   11040
      NewRow1         =   0   'False
      MinHeight2      =   0
      NewRow2         =   0   'False
      BandStyle2      =   1
      Begin MSComctlLib.Toolbar tlbThis 
         Height          =   720
         Left            =   165
         TabIndex        =   9
         Top             =   30
         Width           =   9315
         _ExtentX        =   16431
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsCold"
         HotImageList    =   "ilsHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "PrintView"
               Description     =   "预览"
               Object.ToolTipText     =   "预览"
               Object.Tag             =   "预览"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "Print"
               Description     =   "打印"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "PrintSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "重置"
               Key             =   "Search"
               Description     =   "重置"
               Object.ToolTipText     =   "条件重置"
               Object.Tag             =   "重置"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "过滤"
               Key             =   "Find"
               Description     =   "定位"
               Object.ToolTipText     =   "单据定位"
               Object.Tag             =   "定位"
               ImageIndex      =   14
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "Filter"
               Object.ToolTipText     =   "过滤"
               Object.Tag             =   "过滤"
               ImageKey        =   "Search"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "刷新"
               Key             =   "Refresh"
               Description     =   "刷新"
               Object.ToolTipText     =   "刷新"
               Object.Tag             =   "刷新"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FindSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助主题"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Exit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   11
            EndProperty
         EndProperty
         MouseIcon       =   "frm清单管理.frx":8338
         Begin MSComctlLib.ImageList iltHelp 
            Left            =   3825
            Top             =   210
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   20
            ImageHeight     =   18
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   4
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm清单管理.frx":8652
                  Key             =   "HELPB"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm清单管理.frx":8ADC
                  Key             =   "HELPC"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm清单管理.frx":8F66
                  Key             =   "SEARCHB"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm清单管理.frx":942C
                  Key             =   "SEARCHC"
               EndProperty
            EndProperty
         End
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsHead 
      Height          =   1245
      Left            =   2865
      TabIndex        =   32
      Top             =   1140
      Width           =   6765
      _cx             =   11933
      _cy             =   2196
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
      BackColorFixed  =   -2147483644
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483644
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   12632256
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   12632256
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm清单管理.frx":98F2
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
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "单位名称"
      Height          =   180
      Left            =   3945
      TabIndex        =   13
      Top             =   2475
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lblTemp 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   1
      Left            =   2820
      TabIndex        =   12
      Top             =   2400
      Visible         =   0   'False
      Width           =   3690
   End
   Begin VB.Label lblVsc_s 
      Height          =   75
      Left            =   2820
      MousePointer    =   7  'Size N S
      TabIndex        =   15
      Top             =   2730
      Width           =   6750
   End
   Begin VB.Label lblHsc_s 
      Height          =   5355
      Left            =   2745
      MousePointer    =   9  'Size W E
      TabIndex        =   14
      Top             =   750
      Width           =   60
   End
   Begin VB.Label lblTemp 
      AutoSize        =   -1  'True
      Caption         =   "汇总信息"
      Height          =   180
      Index           =   2
      Left            =   5595
      TabIndex        =   11
      Top             =   825
      Width           =   720
   End
   Begin VB.Label lblTemp 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Index           =   0
      Left            =   2820
      TabIndex        =   10
      Top             =   750
      Width           =   6750
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFilePreView 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileLocalSet 
         Caption         =   "本地参数设置(&R)"
      End
      Begin VB.Menu mnuFileLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Visible         =   0   'False
      Begin VB.Menu mnuEditAdd 
         Caption         =   "增加(&A)"
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "修改(&M)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "删除(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSplit 
         Caption         =   "计划(&S)"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "报表(&R)"
      Visible         =   0   'False
      Begin VB.Menu mnuReportItem 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "标准按钮(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu menuViewLine1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "文本标签(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewUnit 
         Caption         =   "药品供应商(&L)"
         Index           =   0
      End
      Begin VB.Menu mnuViewUnit 
         Caption         =   "物资供应商(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuViewUnit 
         Caption         =   "设备供应商(&E)"
         Index           =   2
      End
      Begin VB.Menu mnuViewUnit 
         Caption         =   "其他供应商(&O)"
         Index           =   3
      End
      Begin VB.Menu mnuViewUnit 
         Caption         =   "卫材供应商(&W)"
         Index           =   4
      End
      Begin VB.Menu mnuViewLine4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewOpen 
         Caption         =   "条件重置(&J)"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "单据定位(&F)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewFilter 
         Caption         =   "过滤(&S)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewLine5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "刷新(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "帮助主题(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&Web上的中联"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "中联主页(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "中联论坛(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "发送反馈(&K)"
         End
      End
      Begin VB.Menu mnuHelpLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)"
      End
   End
End
Attribute VB_Name = "frm清单管理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mdtBegin As Date, mdtEnd As Date
Private mstrKey As String, mstrData As String
Private mstrType As String    '供应商类型
Private msngDownX As Single, msngDownY As Single, mintOldSel As Integer, mstrDeptWhere As String
Private mlngModule As Long

Dim mlng单位ID As Long      '上次选择的供应商
Dim mstrPrivs As String
Dim mblnFirst As Boolean
Private mstrFiler As String     '过滤
Private mstrOthers() As String
Private mint单位 As Integer
Private mintFlag As Integer

Private Sub chkOther_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
        ScrCtl chkOther(Index)
    End If
End Sub
'问题26224 by lesfeng 2010-02-08
Private Sub chk包含付款_Click()
    mintFlag = 1
End Sub

Private Sub DtpOther_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
        ScrCtl DtpOther(Index)
    End If
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    Call LoadOtherCon
    Me.imgHelp.Visible = False
    Me.picHelp.Visible = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)        '
        If KeyAscii = Asc("'") Then
            KeyAscii = 0
        End If
End Sub

Private Sub Form_Load()
    Dim strOthers(0 To 16) As String
    
    mintOldSel = 1
    mstrOthers = strOthers
    mstrPrivs = gstrPrivs: mlngModule = glngModul: mstrKey = "": mblnFirst = True
    
    mlng单位ID = Val(zlDatabase.GetPara("上次选择单位ID", glngSys, mlngModule))
    '问题26224 by lesfeng 2010-02-08
    chk包含付款.Value = IIf(Val(zlDatabase.GetPara("包含时间段之后付款", glngSys, mlngModule)) = 1, 1, 0)
    '问题27878 by lesfeng 2010-02-25
    mint单位 = Val(zlDatabase.GetPara("单位", glngSys, mlngModule))
    mintFlag = 0
    Call initvsHeadHead(True)
    
    mdtEnd = CDate(Format(zlDatabase.Currentdate, "yyyy-MM-dd"))
    mdtBegin = DateAdd("m", -1, mdtEnd) + 1
    mstrData = "00000"
    
    If Check相关权限(mstrPrivs, "药品") = False Then
        mnuViewUnit(0).Checked = False
        mnuViewUnit(0).Enabled = False
    End If
    If Check相关权限(mstrPrivs, "物资") = False Then
        mnuViewUnit(1).Checked = False
        mnuViewUnit(1).Enabled = False
    End If
    If Check相关权限(mstrPrivs, "设备") = False Then
        mnuViewUnit(2).Checked = False
        mnuViewUnit(2).Enabled = False
    End If
    If Check相关权限(mstrPrivs, "其他") = False Then
        mnuViewUnit(3).Checked = False
        mnuViewUnit(3).Enabled = False
    End If
    If Check相关权限(mstrPrivs, "卫材") = False Then
        mnuViewUnit(4).Checked = False
        mnuViewUnit(4).Enabled = False
    End If
    '恢复相关参数
    RestoreWinState Me, App.ProductName
    InitColHead tabSelect.SelectedItem.Index
    FullDept
    '2006-04-25:刘兴宏,统一增加报表发布到模块的功能
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    
    Call vsList_AfterRowColChange(1, 0, 1, 0)
End Sub
'问题27878 by lesfeng 2010-02-25
Private Sub mnuFileLocalSet_Click()
    '本地参数设置
    frm清单管理Set.参数设置 Me, mlngModule, mstrPrivs
    mint单位 = Val(zlDatabase.GetPara("单位", glngSys, mlngModule))
    '重新初使化表
    Select Case tabSelect.SelectedItem.Index
        Case 1
            zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "付款明细列表"
            InitColHead 1
            Full付款明细
        Case 2
            zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "已付明细列表"
            InitColHead 2
            Full已付清单
        Case 3
             zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "未付明细列表"
             InitColHead 3
            Full未付清单
    End Select
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim lng分类id As Long
    Dim lng供应商ID As Long
    
    If Not tvwList.SelectedItem Is Nothing Then
        lng分类id = Val(Mid(Me.tvwList.SelectedItem.Key, 2))
    End If
    
    lng供应商ID = vsHead.RowData(vsHead.Row)
    '2006-04-25:刘兴宏:增加自定义报表发布到模块的功能
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, "分类=" & lng分类id, "供应商=" & lng供应商ID)
End Sub

Private Sub FullCount()
    Dim rstList As New ADODB.Recordset, strSQL As String
'    '问题26224 by lesfeng 2010-02-08
    If Me.mnuViewFilter.Checked Then
    Else
        If tvwList.SelectedItem.Key = mstrKey And mintFlag = 0 Then Exit Sub
    End If
        
    FillSum
    vsHead_EnterCell
End Sub

Private Sub initvsHeadHead(Optional bln初始 As Boolean = False)
    '初始列头
    Dim i As Long
    With vsHead
        .Redraw = False
        .Clear 1
        .Rows = 2
        .Cols = 5
        .TextMatrix(0, 0) = "供应商名称": .ColKey(0) = "供应商名称"
        .TextMatrix(0, 1) = "期初应付": .ColKey(1) = "期初应付"
        .TextMatrix(0, 2) = "本期赊购": .ColKey(2) = "本期赊购"
        .TextMatrix(0, 3) = "本期支付": .ColKey(3) = "本期支付"
        .TextMatrix(0, 4) = "期末应付": .ColKey(4) = "期末应付"
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColIndex("供应商名称") = i Then
                .ColAlignment(i) = 1
            Else
                .ColAlignment(i) = 7
            End If
        Next
        
        If bln初始 Then
            .ColWidth(.ColIndex("供应商名称")) = 3000
            .ColWidth(.ColIndex("期初应付")) = 1200
            .ColWidth(.ColIndex("本期赊购")) = 1200
            .ColWidth(.ColIndex("本期支付")) = 1200
            .ColWidth(.ColIndex("期末应付")) = 1200
        End If
        zl_vsGrid_Para_Restore mlngModule, vsHead, Me.Caption, "余额信息列表", True
        .Redraw = True
    End With
End Sub

Private Sub InitColHead(ByVal intType As Integer, Optional bln初始 As Boolean = True)
    '初始列头
    Dim i As Long
    If intType = 1 Then
        With vsList
                .Redraw = False
                .Clear
                .ExplorerBar = flexExMove
                .Rows = 2
                .FormatString = "^日期|^单据号|^摘要|^单位|^批号|^采购数量|^采购价|^应付金额|^已付金额|^余额"
                .MergeCells = flexMergeNever
                .SelectionMode = flexSelectionByRow
                For i = 0 To .Cols - 1
                     .ColKey(i) = .TextMatrix(0, i)
                     .FixedAlignment(i) = flexAlignCenterCenter
                     Select Case i
                     Case .ColIndex("采购数量"), .ColIndex("采购价"), .ColIndex("应付金额"), .ColIndex("已付金额"), .ColIndex("余额")
                         .ColAlignment(i) = flexAlignRightCenter
                         If bln初始 Then .ColWidth(i) = 1000
                     Case .ColIndex("日期"), .ColIndex("单据号"), .ColIndex("单位")
                         .ColAlignment(i) = flexAlignCenterCenter
                         If bln初始 Then .ColWidth(i) = 1400
                     Case Else
                         .ColAlignment(i) = flexAlignLeftCenter
                         If bln初始 Then .ColWidth(i) = 1400
                     End Select
                Next
                zl_vsGrid_Para_Restore mlngModule, vsList, Me.Caption, "付款明细列表", True
               .Redraw = True
           End With
           Exit Sub
    End If
    If intType = 2 Then
        With vsList
            .Redraw = False
            .Clear
            .Rows = 2
            .ExplorerBar = flexExMove
            .FormatString = "^入库单据号|^发票号|^发票日期|^发票金额|^付款单据号|^日期|^品名|^规格|^单位|^批号|^数量|^金额"
            .MergeCells = flexMergeNever
            .SelectionMode = flexSelectionByRow
            For i = 0 To .Cols - 1
                 .ColKey(i) = .TextMatrix(0, i)
                 .FixedAlignment(i) = flexAlignCenterCenter
                 Select Case i
                 Case .ColIndex("数量"), .ColIndex("金额"), .ColIndex("发票金额")
                     .ColAlignment(i) = flexAlignRightCenter
                     If bln初始 Then .ColWidth(i) = 1000
                 Case .ColIndex("入库单据号"), .ColIndex("发票日期"), .ColIndex("单位"), .ColIndex("日期")
                     .ColAlignment(i) = flexAlignCenterCenter
                     If bln初始 Then .ColWidth(i) = 1400
                 Case Else
                     .ColAlignment(i) = flexAlignLeftCenter
                     If bln初始 Then .ColWidth(i) = 1400
                 End Select
            Next
            zl_vsGrid_Para_Restore mlngModule, vsList, Me.Caption, "已付明细列表", True
            .Redraw = True
        End With
        Exit Sub
    End If
    With vsList
        .Redraw = False
        .Clear
        .Rows = 2
        .FormatString = "^日期|^单据号|^品名|^规格|^单位|^批号|^入库单据号|^数量|^单据金额|^金额|^发票号|^发票日期|^计划序号|^计划付款日期"
        .ExplorerBar = flexExMove
        .MergeCells = flexMergeNever
        .SelectionMode = flexSelectionByRow
        For i = 0 To .Cols - 1
             .ColKey(i) = .TextMatrix(0, i)
             .FixedAlignment(i) = flexAlignCenterCenter
             Select Case i
             Case .ColIndex("数量"), .ColIndex("单据金额"), .ColIndex("金额")
                 .ColAlignment(i) = flexAlignRightCenter
                 If bln初始 Then .ColWidth(i) = 1000
             Case .ColIndex("日期"), .ColIndex("单据号"), .ColIndex("单位"), .ColIndex("发票日期")
                 .ColAlignment(i) = flexAlignCenterCenter
                 If bln初始 Then .ColWidth(i) = 1400
             Case Else
                 .ColAlignment(i) = flexAlignLeftCenter
                 If bln初始 Then .ColWidth(i) = 1400
             End Select
        Next
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeCol(2) = True
        If bln初始 Then
            .ColWidth(.ColIndex("品名")) = 2000
        End If
        zl_vsGrid_Para_Restore mlngModule, vsList, Me.Caption, "未付明细列表", True
        .Redraw = True
    End With
End Sub

Private Sub GetTypeCon()
    '获取类型条件
    Dim intIndex As Integer
    Dim strTmp As String
    strTmp = ""
    For intIndex = 0 To 4
        If mnuViewUnit(intIndex).Checked Then
            strTmp = strTmp & "1"
        Else
            strTmp = strTmp & "0"
        End If
    Next
    mstrType = strTmp ' Bin2Dec(strTmp)
End Sub

Private Sub FullDept()
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:填充供应商
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim intFilt As Integer
    Dim i As Long
    Dim str类型 As String
    Dim strFilt As String
    
    Call GetTypeCon
    
    str类型 = ""
    For i = 1 To Len(mstrType)
        If Mid(mstrType, i, 1) = 1 Then
            str类型 = str类型 & " or substr(类型," & i & ",1)=1"
        End If
    Next
    If str类型 <> "" Then
        str类型 = " And (" & Mid(str类型, 4) & ") "
    End If

    Dim str权限 As String
    str权限 = " and  " & Get分类权限(gstrPrivs)
                
    strSQL = "" & _
        "   Select ID,上级ID,编码,名称,末级 " & _
        "   From 供应商 " & _
        "   Where (撤档时间=TO_DATE('3000-1-1','yyyy-MM-dd') or 撤档时间 is null ) " & _
        "           and (末级=0 Or (末级=1 " & zl_获取站点限制() & " " & str类型 & str权限 & "))" & _
        "   Start with 上级ID is null Connect by prior ID =上级ID"
    On Error GoTo errHandle
    zlDatabase.OpenRecordset rsTemp, strSQL, Me.Caption
    Dim curNode As Node
    tvwList.Nodes.Clear
    tvwList.Nodes.Add , , "Root", "所有供应商", 1
    tvwList.Nodes("Root").Selected = True
    tvwList.Nodes("Root").Expanded = True
    tvwList.Nodes("Root").Sorted = True
    While Not rsTemp.EOF
        If IsNull(rsTemp!上级ID) Then
            Set curNode = tvwList.Nodes.Add("Root", tvwChild, "K" & rsTemp!ID, "【" & rsTemp!编码 & "】" & rsTemp!名称, IIf(rsTemp!末级 <> 1, 5, 2))
        Else
            Set curNode = tvwList.Nodes.Add("K" & rsTemp!上级ID, tvwChild, "K" & rsTemp!ID, "【" & rsTemp!编码 & "】" & rsTemp!名称, IIf(rsTemp!末级 <> 1, 5, 2))
        End If
        If Nvl(rsTemp!ID) = mlng单位ID Then
            curNode.Selected = True
            curNode.Expanded = True
        End If
        curNode.Sorted = True
        rsTemp.MoveNext
    Wend
    FullCount
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Full付款明细()
    Dim rsTemp As New ADODB.Recordset
    Dim strBegin As String, strEnd As String
    Dim dblSum(1 To 2) As Double, dblBalance As Double
    Dim lngRow As Long, lngID As Long
    Dim strSelect As String
  
    '得到查询条件
    lngID = Val(vsHead.Cell(flexcpData, vsHead.Row, vsHead.ColIndex("供应商名称")))
    If lngID = 0 Then Exit Sub
    
    '开始查询
    strBegin = Format(mdtBegin, "yyyy-MM-dd")
    strEnd = Format(mdtEnd + 1, "yyyy-MM-dd")
    On Error GoTo errHandle
    '首先得到期末余额
    gstrSQL = "" & _
        "   Select sum(nvl(金额,0)) as 余额 " & _
        "   From(   Select 金额  From 付款记录  " & _
        "           Where 审核日期>=[2] and 单位ID=[1]" & _
        "           Union All " & _
        "           Select -1 * nvl(发票金额,0) as 金额 from 应付记录 " & _
        "           Where 审核日期>=[2] and 单位ID=[1] " & _
        "           Union All " & _
        "           Select 金额  From 应付余额 " & _
        "           Where 性质=1 and 单位ID=[1]) "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngID, CDate(strEnd))
    If Not rsTemp.EOF Then
        dblBalance = IIf(IsNull(rsTemp("余额")), 0, rsTemp("余额"))
    Else
        dblBalance = 0
    End If
    
    '再得到明细帐
    '问题27878 by lesfeng 2010-02-25
    Select Case mint单位
    Case 0
        strSelect = "A.计量单位,nvl(A.数量,0) as 采购数量,nvl(A.采购价,0) as 采购价,"
    Case 1
        strSelect = "B.门诊单位 as 计量单位,nvl(A.数量,0)/decode(B.门诊包装,null,1,0,1,B.门诊包装) as 采购数量 ,nvl(A.采购价,0)*nvl(B.门诊包装,1) as 采购价,"
    Case 2
        strSelect = "B.住院单位 as 计量单位,nvl(A.数量,0)/decode(B.住院包装,null,1,0,1,B.住院包装) as 采购数量 ,nvl(A.采购价,0)*nvl(B.住院包装,1) as 采购价,"
    Case 3
        strSelect = "B.药库单位 as 计量单位,nvl(A.数量,0)/decode(B.药库包装,null,1,0,1,B.药库包装) as 采购数量 ,nvl(A.采购价,0)*nvl(B.药库包装,1) as 采购价,"
    End Select
    '问题27930 by lesfeng 2010-03-23
    gstrSQL = " Select * From ( " & _
              "  Select to_char(审核日期,'yyyy-MM-dd') as 日期,decode(拒付标志,0,'付','标')||NO as NO, " & _
              "       decode(预付款,1,'预付款',decode(mod(记录状态,3),2,'冲销记录',摘要))||decode(结算方式,'','','('||结算方式||')') as 摘要, " & _
              "       '' as 批号,'' as 单位,0 as 采购数量,0 as 采购价,0 as 应付金额,nvl(金额,0) as 已付金额 " & _
              "       From 付款记录 " & _
              "       where 审核日期>=[1] and 审核日期<[2] and 单位ID=[3]" & _
              " "
    '药品部分
    gstrSQL = gstrSQL & "  Union All  " & _
              "  select to_char(A.审核日期,'yyyy-MM-dd') as 日期,A.入库单据号,A.品名||decode(A.规格,null,'','('||A.规格||')') as 摘要, " & _
              "       A.批号," & strSelect & "nvl(A.发票金额,0) as 应付金额,0 as 已付金额 " & _
              "       from 应付记录 A,药品规格 B " & _
              "       where not A.记录性质 in (-1,2) And A.审核日期>=[1] and A.审核日期<[2] and A.单位ID=[3]" & _
              "         And A.项目ID = B.药品id And A.系统标识 = 1 "
    '除药品其他
    gstrSQL = gstrSQL & "  Union All  " & _
              "  select to_char(审核日期,'yyyy-MM-dd') as 日期,入库单据号,品名||decode(规格,null,'','('||规格||')') as 摘要, " & _
              "       批号,计量单位,nvl(数量,0),nvl(采购价,0),nvl(发票金额,0) as 应付金额,0 as 已付金额 " & _
              "       from 应付记录  " & _
              "       where not 记录性质 in (-1,2) And 审核日期>=[1] and 审核日期<[2] and 单位ID=[3]  And 系统标识 <> 1 " & _
              ")  order by 日期,no"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CDate(strBegin), CDate(strEnd), lngID)
    
    If rsTemp.RecordCount = 0 Then
        Exit Sub
    End If
    vsList.Rows = rsTemp.RecordCount + 3
    lngRow = 2
    With vsList
        .Redraw = False
        '"^日期|^单据号|^摘要|^单位|^批号|^采购数量|^采购价|^应付金额|^已付金额|^余额"
        
        .TextMatrix(1, .ColIndex("日期")) = Format(mdtBegin, "yyyy-MM-dd")
        .TextMatrix(1, .ColIndex("摘要")) = "期初余额"
        
        Do Until rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("日期")) = Nvl(rsTemp("日期"))
            .TextMatrix(lngRow, .ColIndex("单据号")) = Nvl(rsTemp("NO"))
            .TextMatrix(lngRow, .ColIndex("摘要")) = Nvl(rsTemp("摘要"))
            .TextMatrix(lngRow, .ColIndex("单位")) = Nvl(rsTemp("单位"))
            .TextMatrix(lngRow, .ColIndex("批号")) = Nvl(rsTemp("批号"))
            .TextMatrix(lngRow, .ColIndex("采购数量")) = Format(Val(Nvl(rsTemp("采购数量"))), gVbFmtString.FM_数量)
            .TextMatrix(lngRow, .ColIndex("采购价")) = Format(Val(Nvl(rsTemp("采购价"))), gVbFmtString.FM_成本价)
            .TextMatrix(lngRow, .ColIndex("应付金额")) = Format(Val(Nvl(rsTemp("应付金额"))), gVbFmtString.FM_金额)
            .TextMatrix(lngRow, .ColIndex("已付金额")) = Format(Val(Nvl(rsTemp("已付金额"))), gVbFmtString.FM_金额)
            dblSum(1) = dblSum(1) + Nvl(rsTemp("应付金额"), 0)
            dblSum(2) = dblSum(2) + Nvl(rsTemp("已付金额"), 0)
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        
        .TextMatrix(lngRow, .ColIndex("日期")) = Format(mdtEnd, "yyyy-MM-dd")
        .TextMatrix(lngRow, .ColIndex("摘要")) = "合计"
        .TextMatrix(lngRow, .ColIndex("应付金额")) = Format(dblSum(1), gVbFmtString.FM_金额)
        .TextMatrix(lngRow, .ColIndex("已付金额")) = Format(dblSum(2), gVbFmtString.FM_金额)
        .TextMatrix(lngRow, .ColIndex("余额")) = Format(dblBalance, gVbFmtString.FM_金额)
        
        Do Until lngRow = 1
            lngRow = lngRow - 1
            .TextMatrix(lngRow, .ColIndex("余额")) = Format(dblBalance, gVbFmtString.FM_金额)
            dblBalance = dblBalance + Val(.TextMatrix(lngRow, .ColIndex("已付金额"))) - Val(.TextMatrix(lngRow, .ColIndex("应付金额")))
        Loop
        If .Rows - 1 >= 2 Then .Row = 1
        .Col = 0: .LeftCol = 0
        .Redraw = True
    End With
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Full已付清单()
    Dim rsTemp As New ADODB.Recordset, intFilt As Integer, strFilt As String
    Dim strBegin As String, strEnd As String
    Dim dblSum As Double
    Dim lngRow As Long, lngCount As Long, lngTemp As Long
    Dim lngID As Long
    Dim strSelect As String
    
    
    '得到查询条件
    lngID = Val(vsHead.Cell(flexcpData, vsHead.Row, vsHead.ColIndex("供应商名称")))
    If lngID = 0 Then Exit Sub
    '开始查询
    strBegin = Format(mdtBegin, "yyyy-MM-dd") & " 00:00:00"
    strEnd = Format(mdtEnd, "yyyy-MM-dd") & " 23:59:59"
    '得到已付清单
    'by lesfeng 2009-12-2 性能优化
    '问题27878 by lesfeng 2010-02-25
    Select Case mint单位
    Case 0
        strSelect = "Max(A.计量单位) As 单位,Sum(Nvl(A.数量, 0)) as 数量,"
    Case 1
        strSelect = "Max(C.门诊单位) as 单位,Sum(nvl(A.数量,0)/decode(C.门诊包装,null,1,0,1,C.门诊包装)) as 数量,"
    Case 2
        strSelect = "Max(C.住院单位) as 单位,Sum(nvl(A.数量,0)/decode(C.住院包装,null,1,0,1,C.住院包装)) as 数量,"
    Case 3
        strSelect = "Max(C.药库单位) as 单位,Sum(nvl(A.数量,0)/decode(C.药库包装,null,1,0,1,C.药库包装)) as 数量,"
    End Select
    '问题27930 by lesfeng 2010-03-23
    gstrSQL = "Select * From ( " & _
             " Select A.序号,Max(A.入库单据号) || Decode(A.计划序号, Null, ' ', 0, ' ', '(计划:' || a.计划序号 || ')') As 入库单据号, " & _
             "       sum(Nvl(A.发票金额, 0)) As 金额,B.NO   As 付款单据号, " & _
             "       To_Char(Max(B.审核日期), 'yyyy-MM-dd') As 日期, Max(A.品名) As 名称, Max(A.规格) 规格, " & _
             "       Max(A.发票号) 发票号, To_Char(Max(A.发票日期), 'yyyy-mm-dd') 发票日期, Max(A.批号) 批号," & strSelect & _
             "       Sum(Decode(A.记录性质, -1, Nvl(A.计划金额, 0), Nvl(A.发票金额, 0))) As 付款金额 " & _
             " From　应付记录 A,药品规格 C," & _
             "    (Select Distinct 付款序号,NO||'('||decode(预付款,1,'冲预付',decode(拒付标志,1,'标记','付款'))||')' As No,审核日期 " & _
             "     From 付款记录 " & _
             "     Where nvl(预付款,0)<>1 And 审核日期 Between [1] And [2] " & _
             "           And 单位id=[3] )  B " & _
             " Where a.付款序号=b.付款序号 And A.单位id=[3] And A.项目ID = C.药品id And A.系统标识 = 1 and a.记录性质<>2 " & _
             " Group By A.系统标识, A.记录性质, A.NO, A.项目id, A.序号, A.计划序号, B.NO "
             
    gstrSQL = gstrSQL & "  Union All  " & _
             " Select A.序号,Max(A.入库单据号) || Decode(A.计划序号, Null, ' ', 0, ' ', '(计划:' || a.计划序号 || ')') As 入库单据号, " & _
             "       sum(Nvl(A.发票金额, 0)) As 金额,B.NO   As 付款单据号, " & _
             "       To_Char(Max(B.审核日期), 'yyyy-MM-dd') As 日期, Max(A.品名) As 名称, Max(A.规格) 规格, " & _
             "       Max(A.发票号) 发票号, To_Char(Max(A.发票日期), 'yyyy-mm-dd') 发票日期, Max(A.批号) 批号," & _
             "       Max(A.计量单位) As 单位,Sum(Nvl(A.数量, 0)) 数量," & _
             "       Sum(Decode(A.记录性质, -1, Nvl(A.计划金额, 0), Nvl(A.发票金额, 0))) As 付款金额 " & _
             " From　应付记录 A," & _
             "    (Select Distinct 付款序号,NO||'('||decode(预付款,1,'冲预付', decode(拒付标志,1,'标记','付款'))||')' As No,审核日期 " & _
             "     From 付款记录 " & _
             "     Where nvl(预付款,0)<>1 And 审核日期 Between [1] And [2] " & _
             "           And 单位id=[3] )  B " & _
             " Where a.付款序号=b.付款序号 And A.单位id=[3] And A.系统标识 <> 1 and a.记录性质<>2 " & _
             " Group By A.系统标识, A.记录性质, A.NO, A.项目id, A.序号, A.计划序号, B.NO)" & _
             " Order by 入库单据号,序号"
                 
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CDate(strBegin), CDate(strEnd), lngID)
    If rsTemp.RecordCount = 0 Then
        Exit Sub
    Else
        vsList.Rows = rsTemp.RecordCount + 2
    End If
    lngRow = 1
    With vsList
        .Redraw = False
        Do Until rsTemp.EOF
            ' "^入库单据号|^发票号|^发票日期|^发票金额|^付款单据号|^日期|^品名|^规格|^单位|^批号|^数量|^金额"
            .TextMatrix(lngRow, .ColIndex("入库单据号")) = Nvl(rsTemp!入库单据号)
            .TextMatrix(lngRow, .ColIndex("发票号")) = Nvl(rsTemp!发票号)
            .TextMatrix(lngRow, .ColIndex("发票日期")) = Nvl(rsTemp!发票日期)
            .TextMatrix(lngRow, .ColIndex("发票金额")) = Format(Val(Nvl(rsTemp!金额)), gVbFmtString.FM_金额)
            .TextMatrix(lngRow, .ColIndex("付款单据号")) = Nvl(rsTemp!付款单据号)
            .TextMatrix(lngRow, .ColIndex("日期")) = Nvl(rsTemp!日期)
            .TextMatrix(lngRow, .ColIndex("品名")) = Nvl(rsTemp!名称)
            .TextMatrix(lngRow, .ColIndex("规格")) = Nvl(rsTemp!规格)
            .TextMatrix(lngRow, .ColIndex("单位")) = Nvl(rsTemp!单位)
            .TextMatrix(lngRow, .ColIndex("批号")) = Nvl(rsTemp!批号)
            .TextMatrix(lngRow, .ColIndex("数量")) = Format(Val(Nvl(rsTemp!数量)), gVbFmtString.FM_数量)
            .TextMatrix(lngRow, .ColIndex("金额")) = Format(Val(Nvl(rsTemp!付款金额)), gVbFmtString.FM_金额)
            dblSum = dblSum + Val(Nvl(rsTemp!付款金额))
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If .Rows > 2 Then
            .TextMatrix(lngRow, .ColIndex("入库单据号")) = "合计"
            .TextMatrix(lngRow, .ColIndex("金额")) = Format(dblSum, gVbFmtString.FM_金额)
        End If
        If .Rows - 1 >= 2 Then .Row = 1
        .Col = 0: .LeftCol = 0
        .Redraw = True
    End With
End Sub

Private Sub Full未付清单()
    Dim rsTemp As New ADODB.Recordset
    Dim dtStartdate As Date, dtEndDate As Date
    Dim dblSum As Double
    Dim lngRow As Long, lngCount As Long, lngTemp As Long
    Dim lngID As Long
    Dim strTemp As String
    Dim strSelect As String
    Dim strDSelect As String
    '得到查询条件
    lngID = Val(vsHead.Cell(flexcpData, vsHead.Row, vsHead.ColIndex("供应商名称")))
    If lngID = 0 Then Exit Sub
    '开始查询
    dtStartdate = CDate(Format(mdtBegin, "yyyy-mm-dd") & " 00:00:00")
    dtEndDate = CDate(Format(mdtEnd + 1, "yyyy-mm-dd") & " 00:00:00")
    '问题26224 by lesfeng 2010-02-08
    If chk包含付款.Value = 1 Then
        strTemp = " or 审核日期 >= [3]"
    Else
        strTemp = ""
    End If
    '得到未付清单
    '问题27878 by lesfeng 2010-02-25
    Select Case mint单位
    Case 0
        strSelect = "Max(A.计量单位) As 单位,Sum(Nvl(A.数量, 0)) as 数量,"
        strDSelect = "A.计量单位 As 单位, Nvl(A.数量, 0) as 数量,"
    Case 1
        strSelect = "Max(C.门诊单位) as 单位,Sum(nvl(A.数量,0)/decode(C.门诊包装,null,1,0,1,C.门诊包装)) as 数量,"
        strDSelect = "C.门诊单位 as 单位,nvl(A.数量,0)/decode(C.门诊包装,null,1,0,1,C.门诊包装) as 数量,"
    Case 2
        strSelect = "Max(C.住院单位) as 单位,Sum(nvl(A.数量,0)/decode(C.住院包装,null,1,0,1,C.住院包装)) as 数量,"
        strDSelect = "C.住院单位 as 单位,nvl(A.数量,0)/decode(C.住院包装,null,1,0,1,C.住院包装) as 数量,"
    Case 3
        strSelect = "Max(C.药库单位) as 单位,Sum(nvl(A.数量,0)/decode(C.药库包装,null,1,0,1,C.药库包装)) as 数量,"
        strDSelect = "C.药库单位 as 单位,nvl(A.数量,0)/decode(C.药库包装,null,1,0,1,C.药库包装) as 数量,"
    End Select
    
    gstrSQL = " " & _
             "  Select  no as 序号,0 as 标志,a.NO,max(a.入库单据号) 入库单据号,            " & _
             "      to_char(max(A.审核日期),'yyyy-MM-dd') as 日期,max(A.发票号) 发票号,to_char(max(A.发票日期),'yyyy-MM-dd') as 发票日期,            " & _
             "      null as 计划日期,max(a.计划序号) 计划序号,null,max(decode(记录状态,3,A.ID,1,A.ID,0)) ID,             " & _
             "      max(A.品名) as 名称,max(规格) 规格,max(批号) 批号," & strSelect & "sum(nvl(A.发票金额,0)) as 金额 ,sum(nvl(A.单据金额,0)) 单据金额    " & _
             "  From    应付记录 a,药品规格 C" & _
             "  Where  a.审核日期 between [2] and [3] and 计划日期 is null and (A.付款序号 is  null or A.付款序号 is not null and   " & _
             "      A.付款序号  in (Select 付款序号 From 付款记录 where (审核日期 is null" & strTemp & ") and 单位ID=[1]))" & _
             "      and A.单位ID=[1] And A.项目ID = C.药品id And A.系统标识 = 1 and a.记录性质 <> 2 " & _
             "  group by a.系统标识,a.记录性质,A.NO,A.项目id,A.序号 " & _
             "  having sum(nvl(A.发票金额,0))<>0  "
             
    gstrSQL = gstrSQL & " UNION ALL " & _
             "  Select  no as 序号,0 as 标志,a.NO,max(a.入库单据号) 入库单据号,            " & _
             "      to_char(max(A.审核日期),'yyyy-MM-dd') as 日期,max(A.发票号) 发票号,to_char(max(A.发票日期),'yyyy-MM-dd') as 发票日期,            " & _
             "      null as 计划日期,max(a.计划序号) 计划序号,null,max(decode(记录状态,3,A.ID,1,A.ID,0)) ID,             " & _
             "      max(A.品名) as 名称,max(规格) 规格,max(批号) 批号,max(计量单位) As 单位,sum(nvl(数量,0)) 数量, sum(nvl(A.发票金额,0)) as 金额 ,sum(nvl(A.单据金额,0)) 单据金额    " & _
             "  From    应付记录 a " & _
             "  Where  a.审核日期 between [2] and [3] and 计划日期 is null and (A.付款序号 is  null or A.付款序号 is not null and   " & _
             "      A.付款序号  in (Select 付款序号 From 付款记录 where (审核日期 is null" & strTemp & ") and 单位ID=[1]))" & _
             "      and A.单位ID=[1] And A.系统标识 <> 1 and a.记录性质 <> 2 " & _
             "  group by a.系统标识,a.记录性质,A.NO,A.项目id,A.序号 " & _
             "  having sum(nvl(A.发票金额,0))<>0 "

    gstrSQL = gstrSQL & " UNION ALL " & _
            " Select A.NO As 序号, 0 As 标志, A.NO, Max(A.入库单据号) 入库单据号, To_Char(Max(A.审核日期), 'yyyy-MM-dd') As 日期, " & _
            "       Max(A.发票号) 发票号, To_Char(Max(A.发票日期), 'yyyy-MM-dd') As 发票日期, Null As 计划日期, " & _
            "       -1 As 计划序号, Null, Max(Decode(a.记录状态, 3, A.ID, 1, A.ID, 0)) ID, Max(A.品名) As 名称, " & _
            "       Max(a.规格) 规格, Max(a.批号) 批号,Max(a.单位) As 单位, Max(Nvl(a.数量, 0)) 数量, " & _
            "       Max( Nvl(A.发票金额, 0))-Sum( Nvl(B.计划金额, 0)) As 金额, Max(Nvl(A.单据金额, 0)) 单据金额 " & _
            " From ( Select A.ID,A.No,A.入库单据号,A.审核日期,A.发票号,A.发票日期,A.品名,A.规格,A.批号," & strDSelect & "A.发票金额," & _
            "               A.单据金额,A.记录状态,A.系统标识,A.记录性质,A.项目id,A.序号 " & _
            "       From 应付记录 A,药品规格 C " & _
            "       Where A.审核日期 Between [2] And [3] And 计划日期 Is Not Null And A.单位id =[1]  And A.项目ID = C.药品id And A.系统标识 = 1" & _
            "       ) A,应付记录 B " & _
            " Where A.ID = B.ID " & _
            " Group By A.系统标识, A.记录性质, A.NO, A.项目id, A.序号 " & _
            " Having Max(Nvl(A.发票金额, 0)) - Sum(Nvl(B.计划金额, 0)) <> 0 "
            
    gstrSQL = gstrSQL & " UNION ALL " & _
            " Select A.NO As 序号, 0 As 标志, A.NO, Max(A.入库单据号) 入库单据号, To_Char(Max(A.审核日期), 'yyyy-MM-dd') As 日期, " & _
            "       Max(A.发票号) 发票号, To_Char(Max(A.发票日期), 'yyyy-MM-dd') As 发票日期, Null As 计划日期, " & _
            "       -1 As 计划序号, Null, Max(Decode(a.记录状态, 3, A.ID, 1, A.ID, 0)) ID, Max(A.品名) As 名称, " & _
            "       Max(a.规格) 规格, Max(a.批号) 批号, Max(a.计量单位) As 单位, Max(Nvl(a.数量, 0)) 数量, " & _
            "       Max( Nvl(A.发票金额, 0))-Sum( Nvl(B.计划金额, 0)) As 金额, Max(Nvl(A.单据金额, 0)) 单据金额 " & _
            " From ( Select ID,No,入库单据号,审核日期,发票号,发票日期,品名,规格,批号,计量单位,数量,发票金额,单据金额,记录状态,系统标识,记录性质,项目id,序号 " & _
            "       From 应付记录 A " & _
            "       Where A.审核日期 Between [2] And [3] And 计划日期 Is Not Null And A.单位id =[1] And A.系统标识 <> 1" & _
            "       ) A,应付记录 B " & _
            " Where A.ID = B.ID " & _
            " Group By A.系统标识, A.记录性质, A.NO, A.项目id, A.序号 " & _
            " Having Max(Nvl(A.发票金额, 0)) - Sum(Nvl(B.计划金额, 0)) <> 0 "
      
    gstrSQL = gstrSQL & _
           "  UNION ALL  " & _
           "  Select  B.no as 序号,1 as 标志,decode(b.no,null,a.no,b.no) as NO,a.入库单据号,            " & _
           "      to_char(b.审核日期,'yyyy-MM-dd') as 日期,A.发票号,to_char(A.发票日期,'yyyy-MM-dd') as 发票日期,            " & _
           "      to_char(a.计划日期,'yyyy-mm-dd') as 计划日期,a.计划序号,null,A.ID,             " & _
           "      A.品名 as 名称,a.规格,a.批号," & strDSelect & "nvl(A.计划金额,0) as 金额 ,nvl(A.单据金额,0) 单据金额    " & _
           "  From   应付记录 a,药品规格 C, " & _
           "          (Select * From 应付记录   " & _
           "           WHERE 单位ID=[1]  AND 审核日期 between [2] and [3] and (记录状态=1 or mod(记录状态,3)=0) AND 记录性质<>-1 AND (付款序号 IS NULL OR 付款序号 IS NOT NULL AND  " & _
           "              付款序号 IN (Select 付款序号 From 付款记录  where (审核日期 is null" & strTemp & ") and 单位ID=[1]))  " & _
           "              AND 计划日期 IS NOT  NULL  " & _
           "          ) b " & _
           "  Where  a.记录性质=-1  and (a.付款序号 is  null or a.付款序号 in (Select 付款序号 from 付款记录 where (审核日期 is null" & strTemp & ") and 单位id=[1]))  and A.单位ID=[1] AND a.ID=b.id " & _
           "     And A.项目ID = C.药品id And A.系统标识 = 1 "
    
    gstrSQL = gstrSQL & _
           "  UNION ALL  " & _
           "  Select  B.no as 序号,1 as 标志,decode(b.no,null,a.no,b.no) as NO,a.入库单据号,            " & _
           "      to_char(b.审核日期,'yyyy-MM-dd') as 日期,A.发票号,to_char(A.发票日期,'yyyy-MM-dd') as 发票日期,            " & _
           "      to_char(a.计划日期,'yyyy-mm-dd') as 计划日期,a.计划序号,null,A.ID,             " & _
           "      A.品名 as 名称,a.规格,a.批号,a.计量单位 As 单位,nvl(a.数量,0), nvl(A.计划金额,0) as 金额 ,nvl(A.单据金额,0) 单据金额    " & _
           "  From   应付记录 a, " & _
           "          (Select * From 应付记录   " & _
           "           WHERE 单位ID=[1]  AND 审核日期 between [2] and [3] and (记录状态=1 or mod(记录状态,3)=0) AND 记录性质<>-1 AND (付款序号 IS NULL OR 付款序号 IS NOT NULL AND  " & _
           "              付款序号 IN (Select 付款序号 From 付款记录  where (审核日期 is null" & strTemp & ") and 单位ID=[1]))  " & _
           "              AND 计划日期 IS NOT  NULL  " & _
           "          ) b " & _
           "  Where  a.记录性质=-1  and (a.付款序号 is  null or a.付款序号 in (Select 付款序号 from 付款记录 where (审核日期 is null" & strTemp & ") and 单位id=[1]))  and A.单位ID=[1] AND a.ID=b.id  And A.系统标识 <> 1 " & _
           "  ORDER BY 序号,名称,标志,计划序号"

    Err = 0: On Error GoTo ErrHand:
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngID, dtStartdate, dtEndDate)
    
    If rsTemp.RecordCount = 0 Then
        Exit Sub
    Else
        vsList.Rows = rsTemp.RecordCount + 2
    End If
    lngRow = 1
    With vsList
        .Redraw = False
        Do Until rsTemp.EOF
            '"^日期|^单据号|^品名|^规格|^单位|^批号|^入库单据号|^数量|^单据金额|^金额|^发票日期|^计划序号|^计划付款日期"
            .TextMatrix(lngRow, .ColIndex("日期")) = Nvl(rsTemp!日期)
            .TextMatrix(lngRow, .ColIndex("单据号")) = Nvl(rsTemp!NO)
            .TextMatrix(lngRow, .ColIndex("品名")) = Nvl(rsTemp!名称)
            .TextMatrix(lngRow, .ColIndex("规格")) = Nvl(rsTemp!规格)
            .TextMatrix(lngRow, .ColIndex("单位")) = Nvl(rsTemp!单位)
            .TextMatrix(lngRow, .ColIndex("批号")) = Nvl(rsTemp!批号)
            
            .TextMatrix(lngRow, .ColIndex("入库单据号")) = Nvl(rsTemp!入库单据号)
            
            .TextMatrix(lngRow, .ColIndex("数量")) = Format(rsTemp("数量"), gVbFmtString.FM_数量)
            .TextMatrix(lngRow, .ColIndex("单据金额")) = Format(rsTemp("单据金额"), gVbFmtString.FM_金额)
            .TextMatrix(lngRow, .ColIndex("金额")) = Format(rsTemp("金额"), gVbFmtString.FM_金额)
                
            .TextMatrix(lngRow, .ColIndex("发票号")) = Nvl(rsTemp!发票号)
            .TextMatrix(lngRow, .ColIndex("发票日期")) = Nvl(rsTemp!发票日期)
            If Val(Nvl(rsTemp!计划序号)) = 0 Then
                .TextMatrix(lngRow, .ColIndex("计划序号")) = ""
            Else
                .TextMatrix(lngRow, .ColIndex("计划序号")) = IIf(Val(Nvl(rsTemp!计划序号)) = -1, "未编制计划", Nvl(rsTemp!计划序号))
            End If
            .TextMatrix(lngRow, .ColIndex("计划付款日期")) = Nvl(rsTemp!计划日期)
            dblSum = dblSum + Val(Nvl(rsTemp("金额")))
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If .Rows > 2 Then
            .TextMatrix(lngRow, .ColIndex("日期")) = "合计"
            .TextMatrix(lngRow, .ColIndex("金额")) = Format(dblSum, gVbFmtString.FM_金额)
        End If
        If .Rows - 1 >= 2 Then .Row = 1
        .Col = 0: .LeftCol = 0
        
        .Redraw = True
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    
    If Me.WindowState <> vbMaximized Then
        If Me.Height < 5000 Then
            Me.Height = 5000
        End If
        If Me.Width < 4500 Then
            Me.Width = 4500
        End If
    End If
    
    cbrTool.Move 0, 0, Me.ScaleWidth
    
    If lblHsc_s.Left > Me.ScaleWidth - 2000 Then lblHsc_s.Left = Me.ScaleWidth - 2000
    
    lblHsc_s.Top = IIf(cbrTool.Visible, cbrTool.Height + 30, 0)
    lblHsc_s.Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - lblHsc_s.Top - 15
    tvwList.Move 0, lblHsc_s.Top, lblHsc_s.Left, lblHsc_s.Height
    With fraSearch
        .Left = tvwList.Left
        .Top = tvwList.Top
        .Height = tvwList.Height
        .Width = tvwList.Width
    End With
    With fraSplit
        .Left = 0
        .Width = fraSearch.Width
    End With
    With PicSearchBack
        .Left = 0
        .Width = fraSearch.Width - 10
        .Height = IIf(fraSearch.Height - .Top < 0, 0, fraSearch.Height - .Top)
    End With
    Dim sngTmp As Single
    With PicClose
        sngTmp = PicSearchBack.Left + PicSearchBack.Width - .Width - 50
        .Left = sngTmp
    End With
    
    lblVsc_s.Left = lblHsc_s.Left + lblHsc_s.Width
    lblVsc_s.Width = Me.ScaleWidth - lblVsc_s.Left
    
    If lblVsc_s.Top > Me.ScaleHeight - 2000 Then lblVsc_s.Top = Me.ScaleHeight - 2000
    
    lblTemp(0).Move lblVsc_s.Left, lblHsc_s.Top, lblVsc_s.Width
    lblTemp(2).Move lblVsc_s.Left + (lblVsc_s.Width - lblTemp(2).Width) / 2, lblTemp(0).Top + (lblTemp(0).Height - lblTemp(2).Height) / 2
    vsHead.Move lblVsc_s.Left, lblTemp(0).Top + lblTemp(0).Height + 15, lblTemp(0).Width, lblVsc_s.Top - (lblTemp(0).Top + lblTemp(0).Height + 15)
    
    tabSelect.Move lblVsc_s.Left, lblVsc_s.Top + lblVsc_s.Height
    '问题26224 by lesfeng 2010-02-08
    chk包含付款.Top = tabSelect.Top
    chk包含付款.Left = tabSelect.Left + tabSelect.Width + 200
    
    picFind.Move Me.ScaleWidth - picFind.Width - 5, tabSelect.Top
    If picFind.Left < lblVsc_s.Left Then picFind.Left = lblVsc_s.Left
       
    vsList.Move lblVsc_s.Left, tabSelect.Top + tabSelect.Height + 15, lblVsc_s.Width
    vsList.Height = Me.ScaleHeight - vsList.Top - IIf(stbThis.Visible, stbThis.Height, 0) - 30
     
    Call picFind_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
    zlDatabase.SetPara "上次选择单位ID", mlng单位ID, glngSys, mlngModule
    '问题26224 by lesfeng 2010-02-08
    zlDatabase.SetPara "包含时间段之后付款", IIf(chk包含付款.Value, 1, 0), glngSys, mlngModule
    zl_vsGrid_Para_Save mlngModule, vsHead, Me.Caption, "余额信息列表", True
    Select Case tabSelect.SelectedItem.Index
        Case 1
            zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "付款明细列表", True
        Case 2
            zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "已付明细列表", True
        Case 3
            zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "未付明细列表", True
    End Select
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFilePreView_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub mnuFilePrintSet_Click()
    zlPrintSet
End Sub

Private Sub lblHsc_s_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    msngDownX = X
End Sub

Private Sub lblHsc_s_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        With lblHsc_s
            If .Left + X - msngDownX < 2000 Then Exit Sub
            If .Left + X - msngDownX > ScaleWidth - 2000 Then Exit Sub
            .Left = .Left + X - msngDownX
        End With
        Call Form_Resize
        
    End If
End Sub

Private Sub lblVsc_s_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    msngDownY = Y
End Sub

Private Sub lblVsc_s_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        With lblVsc_s
            If .Top + Y - msngDownY < 2000 Then Exit Sub
            If .Top + Y - msngDownY > ScaleHeight - 2000 Then Exit Sub
            .Top = .Top + Y - msngDownY
        End With
        Call Form_Resize
    End If
End Sub

Private Sub mnuViewFilter_Click()
    mnuViewFilter.Checked = Not mnuViewFilter.Checked
    
    If Not mnuViewFilter.Checked Then
        tlbThis.Buttons("Filter").Value = tbrUnpressed
        Me.fraSearch.Visible = False
    Else
        tlbThis.Buttons("Filter").Value = tbrPressed
        Me.fraSearch.Visible = True
        fraSearch.ZOrder
        Me.txt编码.SetFocus
    End If
End Sub

Private Sub mnuViewFind_Click()
    '单据定位
    '按供应商与单据号定位
    Dim str单据号 As String, str供应商ID As String
    Dim rsTemp As New ADODB.Recordset
    Dim nod As MSComctlLib.Node, lngRow As Long, lngCol As Long
    
    If frm应付款定位.Get定位条件(mstrPrivs, str单据号, str供应商ID) = False Then
        Exit Sub
    End If
    
    If str单据号 <> "" Then
        '根据单据号找到供应商
        On Error GoTo errHandle
        gstrSQL = "select 单位ID from 应付记录 where 入库单据号=[1] and 单位ID is not null"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str单据号)
        
        If rsTemp.EOF = True Then
            MsgBox "单据号为 " & str单据号 & " 的记录没有找到。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        str供应商ID = rsTemp("单位ID")
        rsTemp.Close
    End If
    
    On Error Resume Next
    Set nod = tvwList.Nodes("K" & str供应商ID)
    If Err <> 0 Then
        MsgBox "没有发现指定供应商，可能已经被停用。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    nod.Selected = True
    nod.EnsureVisible
    Call FullCount
    
    If str单据号 <> "" Then
        '找到单据所在列
        If tabSelect.SelectedItem.Index = 1 Then
            lngCol = 1
        Else
            lngCol = 0
        End If
        
        With vsList
            For lngRow = .FixedRows To .Rows - 1
                If .TextMatrix(lngRow, lngCol) = str单据号 Then
                    .TopRow = lngRow
                    Exit For
                End If
            Next
        End With
    End If
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub mnuViewOpen_Click()
    If frmTimeSet.GetTimeScope(mdtBegin, mdtEnd, mstrData, Me) = True Then
        mstrKey = ""
        Call FullCount
    End If
End Sub

Private Sub mnuViewRefresh_Click()
    FullDept
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrTool.Visible = mnuViewToolButton.Checked
    cbrTool.Bands(1).MinHeight = tlbThis.Height
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim buttTemp As Button
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For Each buttTemp In tlbThis.Buttons
        If mnuViewToolText.Checked Then
            buttTemp.Caption = buttTemp.Tag
        Else
            buttTemp.Caption = ""
        End If
    Next
    cbrTool.Bands(1).MinHeight = tlbThis.Height
    Form_Resize
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Form_Resize
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
       ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(hwnd)
End Sub

Private Sub mnuViewUnit_Click(Index As Integer)
    mnuViewUnit(Index).Checked = Not mnuViewUnit(Index).Checked
    FullDept
End Sub

Private Sub picFind_Resize()
    Err = 0: On Error Resume Next
    With picFind
        txtFind.Left = lblFind.Width + lblFind.Left + 10
        txtFind.Width = .ScaleWidth - txtFind.Left
    End With
End Sub

Private Sub vsHead_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsHead, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vsHead_EnterCell()
    lblName = vsHead.TextMatrix(vsHead.Row, 0)
    lblName.Left = lblTemp(1).Left + (lblTemp(1).Width - lblName.Width) / 2

    Select Case tabSelect.SelectedItem.Index
        Case 1
            zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "付款明细列表"
            InitColHead 1
            Full付款明细
        Case 2
            zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "已付明细列表"
            InitColHead 2
            Full已付清单
        Case 3
             zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "未付明细列表"
             InitColHead 3
            Full未付清单
    End Select
    '问题26224 by lesfeng 2010-02-08
    mintFlag = 0
End Sub

Private Sub picClear_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClearSearchData
    RaisEffect picClear, 0, "清除", mRightAgnmt
    picClear.Tag = ""
    img清除.Tag = ""
    ReleaseCapture
End Sub

Private Sub tabSelect_Click()
    '问题26224 by lesfeng 2010-02-08
    If tabSelect.SelectedItem.Index = mintOldSel And mintFlag = 0 Then Exit Sub
    mintFlag = 0
    '保存历史结构
    Select Case mintOldSel
        Case 1
            zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "付款明细列表", True
        Case 2
            zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "已付明细列表", True
        Case 3
            zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "未付明细列表", True
    End Select
    
    '恢复历史结构
    vsList.Cols = 1
    Select Case tabSelect.SelectedItem.Index
        Case 1
            InitColHead 1
            Full付款明细
        Case 2
            InitColHead 2
            Full已付清单
        Case 3
            InitColHead 3
            Full未付清单
    End Select
    mintOldSel = tabSelect.SelectedItem.Index
    lblFind.Caption = "按" & vsList.ColKey(vsList.Col) & "查找"
    lblFind.Tag = vsList.ColKey(vsList.Col)
    txtFind.Text = ""
    Call picFind_Resize
End Sub

Private Sub tlbthis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Find"
            mnuViewFind_Click
        Case "Search"
            mnuViewOpen_Click
        Case "PrintView"
            mnuFilePreView_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Filter"
            mnuViewFilter_Click
        Case "Refresh"
            Call mnuViewRefresh_Click
        Case "Help"
            Call mnuHelpTitle_Click
        Case "Exit"
            Call mnuFileExit_Click
    End Select
End Sub

Private Sub tlbthis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then Me.PopupMenu mnuViewTool
End Sub

Private Sub tvwList_NodeClick(ByVal Node As MSComctlLib.Node)
    If mlng单位ID = Val(Mid(Node.Key, 2)) Then Exit Sub
    mlng单位ID = Val(Mid(Node.Key, 2))
    FullCount
End Sub

Private Sub FillSum()
'功能:装入各种统计数据
    Dim rsTemp As New ADODB.Recordset
    Dim strBegin As String, strEnd As String
    Dim dblSum(1 To 4) As Double
    Dim lngRow As Long
    Dim blnSum As Boolean        '合计的显示
    Dim i As Long
    Dim str类型 As String
    Dim lng上级id As Long
    
    str类型 = ""
    For i = 1 To Len(mstrType)
        If Mid(mstrType, i, 1) = 1 Then
            str类型 = str类型 & " or substr(b.类型," & i & ",1)=1"
        End If
    Next
    If str类型 <> "" Then
        str类型 = " And (" & Mid(str类型, 4) & ") "
    End If

    Dim str权限 As String
    str权限 = " and  " & Get分类权限(gstrPrivs)
    
    stbThis.Panels(2).Text = "时间范围：" & Format(mdtBegin, "yyyy-MM-dd") & " 至 " & Format(mdtEnd, "yyyy-MM-dd")

    If tvwList.SelectedItem Is Nothing Then Exit Sub
    If mnuViewFilter.Checked Then
        mstrKey = ""
    Else
        If mstrKey = tvwList.SelectedItem.Key Then Exit Sub
        mstrKey = tvwList.SelectedItem.Key
    End If
    '开始查询
    'by lesfeng 2009-12-2 性能优化
    strBegin = Format(mdtBegin, "yyyyMMdd")
    strEnd = Format(mdtEnd, "yyyyMMdd")
    If UCase(mstrKey) = "ROOT" Then
        lng上级id = 0
    Else
        lng上级id = Val(Mid(mstrKey, 2))
    End If
    MousePointer = 11
    '首先得到子查询的SQL语句
    If mnuViewFilter.Checked Then
        '过滤:
        gstrSQL = mstrFiler & " and  " & Get分类权限(gstrPrivs, "B.")
    Else
        If tvwList.SelectedItem.Image = "2" Then
            gstrSQL = " and A.单位ID=" & Mid(mstrKey, 2)
        ElseIf tvwList.SelectedItem.Image = "1" Then
            gstrSQL = " and A.单位ID in (select ID from 供应商 where 1=1 " & zl_获取站点限制() & "  " & Replace(str类型, "b.类型", "类型") & str权限 & " start with 上级ID is null connect by prior id=上级ID )"
        Else
            gstrSQL = " and A.单位ID in (select ID from 供应商 where 1=1 " & zl_获取站点限制() & "  " & Replace(str类型, "b.类型", "类型") & str权限 & " start with 上级ID =[2] connect by prior id=上级ID  )"
        End If
    End If
    
    If Mid(mstrData, 1, 1) = "1" Then
        gstrSQL = gstrSQL & " And A.期初应付<>0 "
    End If
    If Mid(mstrData, 2, 1) = "1" Then
        gstrSQL = gstrSQL & " And A.本期赊购<>0 "
    End If
    If Mid(mstrData, 3, 1) = "1" Then
        gstrSQL = gstrSQL & " And A.本期支付<>0 "
    End If
    If Mid(mstrData, 4, 1) = "1" Then
        gstrSQL = gstrSQL & " And A.期末应付<>0 "
    End If
    
    '再得到完整的SQL语句
    gstrSQL = "select '【'||B.编码||'】'|| B.名称 as 名称,B.ID,A.期初应付,A.本期赊购,A.本期支付,A.期末应付 from " & _
            "(select 单位ID,sum(余额-期初应付+期初付款) as 期初应付,sum(期初应付-期末应付) as 本期赊购 " & _
            "            ,sum(期初付款-期末付款) as 本期支付,sum(余额-期末应付+期末付款) as 期末应付 " & _
            "from( " & _
            "select 单位ID,nvl(金额,0)  as 期初付款, " & _
            "    decode(sign(to_char(审核日期,'yyyymmdd')-'" & strEnd & "'),1,nvl(金额,0),0) as 期末付款, " & _
            "    0 as 期初应付,0 as 期末应付,0 as 余额 from 付款记录 " & _
            "    where 审核日期>=[1] " & _
            "Union All " & _
            "select 单位ID 单位ID,0 as 期初付款,0 as 期末付款, " & _
            "    发票金额 as 期初应付,decode(sign(to_char(审核日期,'yyyymmdd')-'" & strEnd & "'),1,nvl(发票金额,0),0) as 期末应付,0 as 余额 from 应付记录 " & _
            "    where 记录性质<>-1 And 审核日期>=[1] " & _
            "Union All " & _
            "select 单位ID 单位ID,0 as 期初付款,0 as 期末付款,0 as 期初应付,0 as 期末应付,nvl(金额,0) as 余额 from 应付余额 " & _
            "    where 性质=1) " & _
            "group by 单位ID)A,供应商 B " & _
            "where A.单位ID=B.ID  " & str类型 & gstrSQL
    On Error GoTo errHandle
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CDate(Format(mdtBegin, "yyyy-MM-dd")), lng上级id, mstrOthers(0), mstrOthers(1), _
                            mstrOthers(2), mstrOthers(3), mstrOthers(4), mstrOthers(5), mstrOthers(6), mstrOthers(7), mstrOthers(8), mstrOthers(9), _
                            mstrOthers(10), mstrOthers(11), mstrOthers(12), mstrOthers(13), mstrOthers(14), mstrOthers(15))
    
    initvsHeadHead
    vsHead.Redraw = False
    If rsTemp.RecordCount = 0 Then
        vsHead.Rows = 2
        vsHead.RowData(1) = 0
    Else
        If rsTemp.RecordCount = 1 Then
            '只有一行，就不显示合计了
            vsHead.Rows = 2
            blnSum = False
        Else
            vsHead.Rows = rsTemp.RecordCount + 2
            blnSum = True
        End If
    End If
    lngRow = 1
    With vsHead
        Do Until rsTemp.EOF
            .RowData(lngRow) = rsTemp("ID")
            .TextMatrix(lngRow, .ColIndex("供应商名称")) = Nvl(rsTemp!名称)
            .Cell(flexcpData, lngRow, .ColIndex("供应商名称")) = Nvl(rsTemp!ID)
            .TextMatrix(lngRow, .ColIndex("期初应付")) = Format(Val(Nvl(rsTemp!期初应付)), gVbFmtString.FM_金额)
            .TextMatrix(lngRow, .ColIndex("本期赊购")) = Format(Val(Nvl(rsTemp!本期赊购)), gVbFmtString.FM_金额)
            .TextMatrix(lngRow, .ColIndex("本期支付")) = Format(Val(Nvl(rsTemp!本期支付)), gVbFmtString.FM_金额)
            .TextMatrix(lngRow, .ColIndex("期末应付")) = Format(Val(Nvl(rsTemp!期末应付)), gVbFmtString.FM_金额)
            If blnSum = True Then
                dblSum(1) = dblSum(1) + Nvl(rsTemp("期初应付"), 0)
                dblSum(2) = dblSum(2) + Nvl(rsTemp("本期赊购"), 0)
                dblSum(3) = dblSum(3) + Nvl(rsTemp("本期支付"), 0)
                dblSum(4) = dblSum(4) + Nvl(rsTemp("期末应付"), 0)
            End If
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If blnSum = True Then
            .TextMatrix(lngRow, 0) = "  合计"
            .Cell(flexcpData, lngRow, .ColIndex("供应商名称")) = 0
            .TextMatrix(lngRow, .ColIndex("期初应付")) = Format(dblSum(1), gVbFmtString.FM_金额)
            .TextMatrix(lngRow, .ColIndex("本期赊购")) = Format(dblSum(2), gVbFmtString.FM_金额)
            .TextMatrix(lngRow, .ColIndex("本期支付")) = Format(dblSum(3), gVbFmtString.FM_金额)
            .TextMatrix(lngRow, .ColIndex("期末应付")) = Format(dblSum(4), gVbFmtString.FM_金额)
        End If
                
    End With
    vsHead.Redraw = True
    
    MousePointer = 0
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then
        Resume
    Else
        MousePointer = 0
    End If
End Sub

Private Sub subPrint(bytMode As Byte)
'功能:进行打印,预览和输出到EXCEL
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    
    If vsHead Is ActiveControl Then
        Set objPrint.Body = vsHead
        objPrint.Title.Text = "应付款汇总信息"
        objRow.Add " "
        objRow.Add "查询时间：" & Format(mdtBegin, "yyyy-MM-dd") & " 至 " & Format(mdtEnd, "yyyy-MM-dd")
        objPrint.UnderAppRows.Add objRow
        
        Set objRow = New zlTabAppRow
        objRow.Add "打印人：" & UserInfo.姓名
        objRow.Add "打印时间：" & Format(zlDatabase.Currentdate, "yyyy-MM-dd")
        objPrint.BelowAppRows.Add objRow
    Else
        Set objPrint.Body = vsList
        objPrint.Title.Text = tabSelect.SelectedItem.Caption
        objRow.Add "供应商：" & Mid(lblName.Caption, InStr(lblName.Caption, "】") + 1)
        objRow.Add "查询时间：" & Format(mdtBegin, "yyyy-MM-dd") & " 至 " & Format(mdtEnd, "yyyy-MM-dd")
        objPrint.UnderAppRows.Add objRow
        
        Set objRow = New zlTabAppRow
        objRow.Add "打印人：" & UserInfo.姓名
        objRow.Add "打印时间：" & Format(zlDatabase.Currentdate, "yyyy-MM-dd")
        objPrint.BelowAppRows.Add objRow
    End If
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub PicSearchBack_Resize()
    Dim i As Long
    Dim CtlWidth As Single
    Dim sngBottom As Single
    Dim blnOther As Boolean
    
    If InStr(1, Me.lblHit.Caption, ">>") <> 0 Then
        sngBottom = Me.lblHit.Top + Me.lblHit.Height
    Else
        sngBottom = shpHit.Top + shpHit.Height
    End If
    
    With picSearch
        .Left = PicSearchBack.ScaleLeft
        .Top = PicSearchBack.ScaleTop
    End With
    
    If PicSearchBack.ScaleHeight < sngBottom Then
        Scr.Visible = True
        picSearch.Width = IIf(PicSearchBack.ScaleWidth - Me.Scr.Width < 0, 0, PicSearchBack.ScaleWidth - Me.Scr.Width)
    Else
        Scr.Visible = False
        picSearch.Width = PicSearchBack.ScaleWidth
    End If
    With Scr
        .Left = picSearch.Left + picSearch.Width
        .Top = picSearch.Top
        .Height = PicSearchBack.ScaleHeight
    End With
    shpHit.Width = IIf(picSearch.Width - 100 < 0, 0, picSearch.Width - 100)
    CtlWidth = IIf(shpHit.Width - 100 < 0, 0, shpHit.Width - 100)
    lblHit.Width = shpHit.Width
    For i = 0 To lblOther.UBound
        lblOther(i).Width = CtlWidth
        TxtOther(i).Width = CtlWidth
    Next
    For i = 0 To chkOther.UBound
        chkOther(i).Width = CtlWidth
    Next
    For i = 0 To lblDate.UBound
        lblDate(i).Width = CtlWidth
        DtpOther(i).Width = CtlWidth
    Next
    Scr.Max = Int(Me.picSearch.Height / Me.PicSearchBack.Height + 0.5) * 12
End Sub

Private Sub Scr_Change()
    Scr_Scroll
End Sub

Private Sub Scr_Scroll()
    picSearch.Top = -Scr.Value * (Me.PicSearchBack.Height / 12) + 400
End Sub


Private Sub lblHit_Click()
    Dim i As Long
    Dim blnTrue As Boolean
    
    If InStr(1, lblHit.Caption, "<<") <> 0 Then
        blnTrue = False
        lblHit.Caption = Replace(Me.lblHit.Caption, "<<", ">>")
        lblHit.BackStyle = 0
        lblHit.ForeColor = &H8000000D
        shpHit.Visible = False
    Else
        blnTrue = True
        lblHit.Caption = Replace(Me.lblHit.Caption, ">>", "<<")
        lblHit.BackStyle = 1
        lblHit.ForeColor = &H8000000E
        shpHit.Visible = True
    End If
    For i = 0 To lblOther.UBound
        lblOther(i).Visible = shpHit.Visible
        TxtOther(i).Visible = shpHit.Visible
    Next
    For i = 0 To chkOther.UBound
        chkOther(i).Visible = chkOther(i).Visible And shpHit.Visible
    Next
    For i = 0 To lblDate.UBound
        lblDate(i).Visible = shpHit.Visible
        DtpOther(i).Visible = shpHit.Visible
    Next
    PicSearchBack_Resize
End Sub

Private Function ClearSearchData()
    '------------------------------------------------------------------
    '功能:清除条件及相关数据
    '------------------------------------------------------------------
    Dim i As Long
    initvsHeadHead True
    Call InitColHead(1, True)
    Call InitColHead(2, True)
    
    Me.txt编码.Text = ""
    Me.Txt名称.Text = ""
    For i = 0 To TxtOther.UBound
        TxtOther(i).Text = ""
    Next
    For i = 0 To chkOther.UBound
        chkOther(i).Value = 0
    Next
    For i = 0 To DtpOther.UBound
        DtpOther(i).Value = ""
    Next
End Function

Private Sub PicClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaisEffect PicClose, 2, "×", mCenterAgnmt, True
End Sub

Private Sub PicClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If PicClose.Tag = "In" Then
        If X < 0 Or Y < 0 Or X > PicClose.Width Or Y > PicClose.Height Then
            PicClose.Tag = ""
            ReleaseCapture
            RaisEffect PicClose, 0, "×", mCenterAgnmt, True
        End If
    Else
        PicClose.Tag = "In"
        SetCapture PicClose.hwnd
        MousePointer = 99
        RaisEffect PicClose, 1, "×", mCenterAgnmt, True
    End If
End Sub

Private Sub PicClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PicClose.Tag = ""
    RaisEffect PicClose, 0, "×", mCenterAgnmt, True
    ReleaseCapture
    mnuViewFilter_Click
    Call FullCount
End Sub

Private Sub picHelp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaisEffect picHelp, 2
End Sub

Private Sub picHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If picHelp.Tag = "In" Then
        If X < 0 Or Y < 0 Or X > picHelp.Width Or Y > picHelp.Height Then
            picHelp.Tag = ""
            ReleaseCapture
            RaisEffect picHelp, 0
            Set Me.imgHelp.Picture = iltHelp.ListImages("HELPB").Picture
        End If
    Else
        picHelp.Tag = "In"
        SetCapture picHelp.hwnd
        MousePointer = 99
        RaisEffect picHelp, 1
        Set Me.imgHelp.Picture = iltHelp.ListImages("HELPC").Picture      'LoadResPicture("HELPC", 0)
    End If
End Sub

Private Sub picHelp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
    RaisEffect picHelp, 0
    picHelp.Tag = ""
    imgHelp.Tag = ""
    ReleaseCapture
End Sub

Private Sub picClear_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaisEffect picClear, 2, "清除", mRightAgnmt
End Sub

Private Sub picClear_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If picClear.Tag = "In" Then
        If X < 0 Or Y < 0 Or X > picClear.Width Or Y > picClear.Height Then
            picClear.Tag = ""
            ReleaseCapture
            RaisEffect picClear, 0, "清除", mRightAgnmt
            Set Me.img清除.Picture = iltHelp.ListImages("SEARCHB").Picture   'LoadResPicture("SEARCHB", 0)
        End If
    Else
        picClear.Tag = "In"
        SetCapture picClear.hwnd
        MousePointer = 99
        RaisEffect picClear, 1, "清除", mRightAgnmt
        Set Me.img清除.Picture = iltHelp.ListImages("SEARCHC").Picture ' LoadResPicture("SEARCHC", 0)
    End If
End Sub

Private Sub imgHelp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picHelp_MouseDown Button, Shift, X, Y
End Sub

Private Sub imgHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If imgHelp.Tag = "In" Then
        If X < 0 Or Y < 0 Or X > imgHelp.Width Or Y > imgHelp.Height Then
            imgHelp.Tag = ""
            ReleaseCapture
            RaisEffect picHelp, 0
            Set Me.imgHelp.Picture = iltHelp.ListImages("HELPB").Picture
        End If
    Else
        imgHelp.Tag = "In"
        SetCapture picHelp.hwnd
        MousePointer = 99
        RaisEffect picHelp, 1
        Set Me.imgHelp.Picture = iltHelp.ListImages("HELPC").Picture      'LoadResPicture("HELPC", 0)
    End If
End Sub

Private Sub imgHelp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picHelp_MouseUp Button, Shift, X, Y
End Sub

Private Function IsValitSearchCon() As Boolean
    '------------------------------------------------------
    '功能:检查搜索条件是否有效
    '------------------------------------------------------
    Dim i As Long
    IsValitSearchCon = False
    If InStr(1, Me.txt编码.Text, "'") <> 0 Then
        MsgBox "编码或简码中含用非法字符！", vbInformation, gstrSysName
        Exit Function
    End If
    If InStr(1, Me.Txt名称.Text, "'") <> 0 Then
        MsgBox "名称中含用非法字符！", vbInformation, gstrSysName
        Exit Function
    End If
    For i = 0 To TxtOther.UBound
        If InStr(1, Me.TxtOther(i).Text, "'") <> 0 Then
            MsgBox Me.TxtOther(i).Tag & "中含用非法字符！", vbInformation, gstrSysName
            Exit Function
        End If
    Next
    IsValitSearchCon = True
End Function

Private Sub TxtOther_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
        ScrCtl TxtOther(Index)
    End If
End Sub

Private Sub TxtOther_LostFocus(Index As Integer)
    Dim strIme As String
    zlCommFun.OpenIme (False)
End Sub

Private Sub txt编码_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Txt名称_GotFocus()
    Dim strIme As String
    Txt名称.SelStart = 0
    Txt名称.SelLength = Len(Txt名称)
    zlCommFun.OpenIme (True)
End Sub

Private Sub cmd搜索_Click()
    Dim strWhere As String
    If Not IsValitSearchCon Then Exit Sub
    strWhere = Trim(GetSearchCon)
    If strWhere = "" Then
        ShowMsgbox "未输入过滤条件,请输入！"
        Exit Sub
    End If
    mstrFiler = " AND (" & strWhere & " )"
   '条件
   Call FullCount
End Sub

Private Sub LoadOtherCon()
    '---------------------------------------------------------------------------------
    '功能:加载其它条件选择项
    '----------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strFind As String
    Dim CurTop As Single
    Dim CtlWidth As Single
    Dim CtlLeft As Single
    CtlWidth = IIf(lblHit.Width - 50 < 0, 0, lblHit.Width - 50)
    CtlLeft = lblHit.Left + 50
    Dim sngTabIndex As Long
    sngTabIndex = 4
    RaisEffect PicClose, 0, "×", mCenterAgnmt, True
    RaisEffect picClear, 0, "清除", mRightAgnmt

    '加载文本条件
    For i = 0 To 9
        strFind = Switch(i = 0, "地址", i = 1, "许可证号", i = 2, "执照号", i = 3, "税务登记号", i = 4, "帐号", _
             i = 5, "联系人", i = 6, "开户银行", i = 7, "销售委托人", i = 8, "质量认证号", i = 9, "药监局备案号")
             
        If i <> 0 Then
            Load lblOther(i)
            Load TxtOther(i)
            lblOther(i).Top = CurTop
            CurTop = CurTop + lblOther(i).Height + 50
            TxtOther(i).Top = CurTop
            CurTop = CurTop + TxtOther(i).Height + 100
        Else
            CurTop = TxtOther(i).Top + TxtOther(i).Height + 100
        End If
        lblOther(i).TabIndex = sngTabIndex
        sngTabIndex = sngTabIndex + 1
        TxtOther(i).TabIndex = sngTabIndex
        sngTabIndex = sngTabIndex + 1
        lblOther(i) = strFind
        TxtOther(i).Tag = strFind
        lblOther(i).Left = CtlLeft
        TxtOther(i).Left = CtlLeft
        TxtOther(i).Width = CtlWidth
        lblOther(i).Width = CtlWidth
    Next
    
    '加载选择条件
'    For i = 0 To 2
'        strFind = Switch(i = 0, "无菌性材料", i = 1, "一次性材料", i = 2, "定额材料")
'        If i <> 0 Then
'            Load chkOther(i)
'        End If
'        chkOther(i).Top = CurTop
'        chkOther(i).TabIndex = sngTabIndex
'        sngTabIndex = sngTabIndex + 1
'        CurTop = CurTop + chkOther(i).Height + 100
'        chkOther(i).Caption = strFind
'        chkOther(i).Tag = strFind
'        chkOther(i).Left = CtlLeft
'        chkOther(i).Width = CtlWidth
'        If marblnSelectWare And (marblnOnlyRation Or marBln无菌性材料) Then
'            chkOther(i).Enabled = False
'        Else
'            chkOther(i).Enabled = True
'        End If
'    Next
    '加载时间选择控件
    For i = 0 To 3
        strFind = Switch(i = 0, "建档时间", i = 1, "撤档时间", i = 2, "许可证效期", i = 3, "执照效期")
        If i <> 0 Then
            Load lblDate(i)
            Load DtpOther(i)
        End If
        lblDate(i).TabIndex = sngTabIndex
        sngTabIndex = sngTabIndex + 1
        DtpOther(i).TabIndex = sngTabIndex
        sngTabIndex = sngTabIndex + 1
        lblDate(i) = strFind
        DtpOther(i).Tag = strFind
        lblDate(i).Left = CtlLeft
        DtpOther(i).Left = CtlLeft
        DtpOther(i).Width = CtlWidth
        lblDate(i).Width = CtlWidth
        lblDate(i).Top = CurTop
        CurTop = CurTop + lblDate(i).Height + 50
        DtpOther(i).Top = CurTop
        If i < 2 Then
            DtpOther(i).MaxDate = zlDatabase.Currentdate()
        Else
            DtpOther(i).MaxDate = CDate("3000-01-01")
        End If
        DtpOther(i).Value = zlDatabase.Currentdate()
        CurTop = CurTop + DtpOther(i).Height + 100
        DtpOther(i).Value = Null
        DtpOther(i).Enabled = True
    Next
    picSearch.Height = CurTop
    shpHit.Height = IIf(CurTop - shpHit.Top < 0, 0, CurTop - shpHit.Top)
    chkOther(0).Visible = False
End Sub

Private Function GetSearchCon() As String
    '---------------------------------------------------------------------------------------------------------
    '功能:提取查询条件
    '---------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strWhere As String
    Dim strTemp As String
    Dim strField As String
    Dim LfPBF As String
    Dim RgPbf As String
    Dim strOthers(0 To 16) As String
    
    If gstrMatchMethod = "0" Then
        LfPBF = "%"
        RgPbf = "%"
    Else
        LfPBF = ""
        RgPbf = "%"
    End If
    
    strWhere = ""
    strTemp = Trim(txt编码.Text)
    If strTemp <> "" Then
        If InStr(1, strTemp, "%") <> 0 Then
            strWhere = strWhere & "   or  (B.编码 Like [3]) "
            strOthers(0) = strTemp
        Else
            strWhere = strWhere & "   or  (B.编码 Like [3]) "
            strOthers(0) = LfPBF & strTemp & RgPbf
        End If
    End If
    
    strTemp = UCase(Trim(Txt名称.Text))
    If strTemp <> "" Then
        If InStr(1, strTemp, "%") <> 0 Then
            strWhere = strWhere & "   or  (B.名称 Like [4])  "
            strWhere = strWhere & "   or  (B.简码 Like [4])  "
            strOthers(1) = strTemp
        Else
            strWhere = strWhere & "   or  (B.名称 Like [4]) "
            strWhere = strWhere & "   or  (B.简码 Like [4]) "
            strOthers(1) = LfPBF & strTemp & RgPbf
        End If
    End If
    
    If shpHit.Visible Then
        For i = 0 To TxtOther.UBound
            strField = " upper(B." & TxtOther(i).Tag & ")"
            strTemp = UCase(Trim(TxtOther(i).Text))
            If strTemp <> "" Then
                If InStr(1, strTemp, "%") <> 0 Then
                    strWhere = strWhere & "   or  (" & strField & "  Like [" & i + 5 & "]) "
                    strOthers(i + 2) = strTemp
                Else
                    strWhere = strWhere & "   or  ( " & strField & "  Like [" & i + 5 & "]) "
                    strOthers(i + 2) = LfPBF & strTemp & RgPbf
                End If
            End If
        Next
        For i = 0 To DtpOther.UBound
            strField = DtpOther(i).Tag
            If Not IsNull(DtpOther(i).Value) Then
'                strWhere = strWhere & "   or  ( to_char(B." & strField & ",'yyyy-MM-DD')  = '" & Format(DtpOther(i).Value, "yyyy-MM-DD") & "' ) "
                strWhere = strWhere & "   or  ( to_char(B." & strField & ",'yyyy-MM-DD')  = [" & i + 15 & "] ) "
                strOthers(i + 12) = Format(DtpOther(i).Value, "yyyy-MM-DD")
            End If
        Next
    End If
    mstrOthers = strOthers
    strWhere = Mid(strWhere, 6)
    GetSearchCon = strWhere
End Function

Private Sub Txt名称_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub ScrCtl(ByVal ctlObject As Object)
'    Err = 0: On Error Resume Next
'    If (ctlObject.Top + ctlObject.Height) + picSearch.Top + 1800 > PicSearchBack.ScaleHeight Then
'            If Scr.Value + 1 < Scr.Max Then
'                 Scr.Value = Scr.Value + 3
'            End If
'    End If
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

Private Sub vsHead_GotFocus()
    zl_VsGridGotFocus vsHead
End Sub

Private Sub vsHead_LostFocus()
    zl_VsGridLOSTFOCUS vsHead
End Sub

Private Sub vsList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    lblFind.Caption = "按" & vsList.ColKey(NewCol) & "查找"
    If lblFind.Tag <> vsList.ColKey(NewCol) Then
        txtFind.Text = ""
    End If
    lblFind.Tag = vsList.ColKey(NewCol)
    
    Call picFind_Resize
    zl_VsGridRowChange vsList, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vsList_GotFocus()
    zl_VsGridGotFocus vsList
End Sub

Private Sub vsList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyF3 Then Exit Sub
    If Trim(txtFind.Text) = "" Then Exit Sub
    FindRow Trim(txtFind.Text), IIf(vsList.Row + 1 >= vsList.Rows - 1, 1, vsList.Row + 1)
End Sub

Private Sub vsList_LostFocus()
    zl_VsGridLOSTFOCUS vsList
End Sub

Private Sub FindRow(ByVal strFind As String, Optional lngRow As Long = 1)
    '功能:查找指列的数据是否满足相关的条件
    '参数:intMachType:0-左匹配,1-完全匹配
    Dim i As Long, lngCol As Long
    Dim blnAll As Boolean
Redo:
    With vsList
        lngCol = .ColIndex(lblFind.Tag)
        '未找到列退出
        If lngCol < 0 Then Exit Sub
        
        If InStr(1, lblFind.Tag, "数量") > 0 Then
            blnAll = True
            strFind = Format(Val(strFind), gVbFmtString.FM_数量)
        ElseIf InStr(1, lblFind.Tag, "金额") > 0 Then
            blnAll = True
            strFind = Format(Val(strFind), gVbFmtString.FM_金额)
        ElseIf InStr(1, lblFind.Tag, "采购价") > 0 Then
            blnAll = True
            strFind = Format(Val(strFind), gVbFmtString.FM_成本价)
        ElseIf InStr(1, lblFind.Tag, "零售价") > 0 Then
            blnAll = True
            strFind = Format(Val(strFind), gVbFmtString.FM_零售价)
        ElseIf InStr(1, lblFind.Tag, "价") > 0 Then
            blnAll = True
            strFind = Format(Val(strFind), gVbFmtString.FM_金额)
        ElseIf InStr(1, lblFind.Tag, "日期") > 0 Then
            blnAll = False
            strFind = CheckIsDate(strFind, lblFind.Tag)
            If strFind = "" Then Exit Sub
        Else
            blnAll = False
        End If
       i = .FindRow(strFind, lngRow, lngCol, False, blnAll)
       If i > 0 Then
            .Row = i: .TopRow = i
       Else
            If lngRow = 1 Then
                ShowMsgbox "已经查到末尾,没有发现满足条件的数据,请检查"
            Else
                If MsgBox("已经查到末尾,没有发现满足条件的数据,是否重新进行查找!", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                    lngRow = 1
                    GoTo Redo:
                End If
            End If
       End If
    End With
End Sub

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
    If InStr(1, lblFind.Tag, "号") > 0 Or _
        InStr(1, lblFind.Tag, "日期") > 0 Or _
        InStr(1, lblFind.Tag, "额") > 0 Then
        zlCommFun.OpenIme False
    Else
        zlCommFun.OpenIme True
    End If
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strNO As String
    If KeyCode <> vbKeyReturn Then
        If KeyCode = vbKeyF3 Then
            Call vsList_KeyDown(vbKeyF3, 0)
            Exit Sub
        End If
        Exit Sub
    End If
    If Trim(txtFind) = "" Then Exit Sub
    FindRow Trim(txtFind.Text)
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
        
    If InStr(1, lblFind.Tag, "数量") > 0 Then
        zlControl.TxtCheckKeyPress txtFind, KeyAscii, m负金额式
    ElseIf InStr(1, lblFind.Tag, "金额") > 0 Then
        zlControl.TxtCheckKeyPress txtFind, KeyAscii, m负金额式
    ElseIf InStr(1, lblFind.Tag, "采购价") > 0 Then
        zlControl.TxtCheckKeyPress txtFind, KeyAscii, m负金额式
    ElseIf InStr(1, lblFind.Tag, "零售价") > 0 Then
        zlControl.TxtCheckKeyPress txtFind, KeyAscii, m负金额式
    ElseIf InStr(1, lblFind.Tag, "价") > 0 Then
        zlControl.TxtCheckKeyPress txtFind, KeyAscii, m负金额式
    ElseIf InStr(1, lblFind.Tag, "日期") > 0 Then
        zlControl.TxtCheckKeyPress txtFind, KeyAscii, m文本式
    Else
        zlControl.TxtCheckKeyPress txtFind, KeyAscii, m文本式
    End If
        
End Sub

