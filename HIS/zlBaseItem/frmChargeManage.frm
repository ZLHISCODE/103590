VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "CODEJO~1.OCX"
Begin VB.Form frmChargeManage 
   Caption         =   "收费项目管理"
   ClientHeight    =   7890
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   11250
   Icon            =   "frmChargeManage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picPage 
      BorderStyle     =   0  'None
      Height          =   1620
      Index           =   1
      Left            =   5490
      ScaleHeight     =   1620
      ScaleWidth      =   960
      TabIndex        =   28
      Top             =   2520
      Width           =   960
      Begin VSFlex8Ctl.VSFlexGrid msh价目 
         Height          =   1695
         Left            =   60
         TabIndex        =   30
         Top             =   570
         Width           =   3465
         _cx             =   6112
         _cy             =   2990
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
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
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
      Begin VB.CheckBox chk价格 
         Caption         =   "显示历史价格"
         Height          =   315
         Left            =   4770
         TabIndex        =   29
         Top             =   60
         Width           =   1425
      End
      Begin VB.Image img价目 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   540
         Left            =   0
         Picture         =   "frmChargeManage.frx":0442
         Top             =   0
         Width           =   540
      End
      Begin VB.Label lbl价目 
         Caption         =   "    此处显示收费项目的价格，背景颜色为黄色的那几行是当前价格。"
         Height          =   435
         Left            =   780
         TabIndex        =   31
         Top             =   60
         Width           =   3795
      End
   End
   Begin VB.PictureBox pic停用原因 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   150
      ScaleHeight     =   225
      ScaleWidth      =   3945
      TabIndex        =   42
      Top             =   5580
      Width           =   3945
      Begin VB.Label lbl停用原因 
         Caption         =   "停用原因："
         Height          =   225
         Left            =   0
         TabIndex        =   43
         Top             =   0
         Width           =   2955
      End
   End
   Begin VB.PictureBox picPage 
      Height          =   1815
      Index           =   6
      Left            =   8700
      ScaleHeight     =   1755
      ScaleWidth      =   810
      TabIndex        =   11
      Top             =   4860
      Width           =   870
      Begin VSFlex8Ctl.VSFlexGrid vsWholeSet 
         Height          =   4680
         Left            =   0
         TabIndex        =   38
         Top             =   225
         Width           =   11355
         _cx             =   20029
         _cy             =   8255
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
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   16
         FixedRows       =   1
         FixedCols       =   2
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmChargeManage.frx":067C
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
   Begin MSComctlLib.ImageList iltdept 
      Left            =   3180
      Top             =   3345
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
            Picture         =   "frmChargeManage.frx":08AA
            Key             =   "Dept"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picPage 
      Height          =   1815
      Index           =   7
      Left            =   9870
      ScaleHeight     =   1755
      ScaleWidth      =   810
      TabIndex        =   32
      Top             =   4860
      Width           =   870
      Begin MSComctlLib.ListView lvwUseDept 
         Height          =   1230
         Left            =   45
         TabIndex        =   36
         Top             =   165
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   2170
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "iltdept"
         SmallIcons      =   "iltdept"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "科室"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.PictureBox picPage 
      BorderStyle     =   0  'None
      Height          =   1815
      Index           =   5
      Left            =   7500
      ScaleHeight     =   1815
      ScaleWidth      =   870
      TabIndex        =   33
      Top             =   4860
      Width           =   870
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh费别 
         Height          =   2475
         Left            =   0
         TabIndex        =   34
         Top             =   240
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   4366
         _Version        =   393216
         FixedCols       =   0
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483630
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.PictureBox picPage 
      BorderStyle     =   0  'None
      Height          =   1650
      Index           =   3
      Left            =   7650
      ScaleHeight     =   1650
      ScaleWidth      =   900
      TabIndex        =   14
      Top             =   2520
      Width           =   900
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh从属 
         Height          =   1455
         Left            =   0
         TabIndex        =   15
         Top             =   810
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   2566
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
      End
      Begin VB.Image img从属 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   480
         Left            =   0
         Picture         =   "frmChargeManage.frx":0BC6
         Top             =   0
         Width           =   480
      End
      Begin VB.Label lbl从属 
         Caption         =   "    从属项目是指用户在进行单据录入中，会随着主收费项目的增加而自动增加的收费项目。"
         Height          =   435
         Left            =   870
         TabIndex        =   16
         Top             =   90
         Width           =   3795
      End
   End
   Begin VB.PictureBox picPage 
      BorderStyle     =   0  'None
      Height          =   1650
      Index           =   4
      Left            =   8730
      ScaleHeight     =   1650
      ScaleWidth      =   900
      TabIndex        =   12
      Top             =   2520
      Width           =   900
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshAlias 
         Height          =   2475
         Left            =   120
         TabIndex        =   13
         Top             =   0
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   4366
         _Version        =   393216
         FixedCols       =   0
         BackColorSel    =   -2147483639
         ForeColorSel    =   -2147483630
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.PictureBox picTreeClass_S 
      BorderStyle     =   0  'None
      Height          =   1125
      Left            =   120
      ScaleHeight     =   1125
      ScaleWidth      =   2700
      TabIndex        =   9
      Top             =   4380
      Width           =   2700
      Begin XtremeSuiteControls.TabControl tbClassPage 
         Height          =   2700
         Left            =   -270
         TabIndex        =   10
         Top             =   -495
         Width           =   2175
         _Version        =   589884
         _ExtentX        =   3836
         _ExtentY        =   4762
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picTreeWholeSet 
      BorderStyle     =   0  'None
      Height          =   1605
      Left            =   60
      ScaleHeight     =   1605
      ScaleWidth      =   2700
      TabIndex        =   7
      Top             =   2520
      Width           =   2700
      Begin MSComctlLib.TreeView tvwWholeSet 
         Height          =   1485
         Left            =   0
         TabIndex        =   8
         Top             =   240
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   2619
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         ImageList       =   "ils16"
         Appearance      =   1
      End
   End
   Begin VB.PictureBox picTreeItem 
      BorderStyle     =   0  'None
      Height          =   1560
      Left            =   90
      ScaleHeight     =   1560
      ScaleWidth      =   2745
      TabIndex        =   5
      Top             =   825
      Width           =   2745
      Begin MSComctlLib.TreeView tvwMainItem 
         Height          =   1485
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   2619
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         ImageList       =   "ils16"
         Appearance      =   1
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   7530
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   635
      SimpleText      =   $"frmChargeManage.frx":1008
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmChargeManage.frx":104F
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14764
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
   Begin VB.PictureBox picNS 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   4050
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   3000
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2040
      Width           =   3000
   End
   Begin MSComctlLib.ListView lvwMain_S 
      Height          =   1065
      Left            =   8040
      TabIndex        =   2
      Top             =   810
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   1879
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils32"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList Ilsrw 
      Left            =   3645
      Top             =   1710
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":18E3
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":1B03
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":1D23
            Key             =   "Parent"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":1F3F
            Key             =   "Child"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":215B
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":237B
            Key             =   "Raise"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":2597
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":27B7
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":29D7
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":2BF7
            Key             =   "View"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":2E17
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":3037
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":3257
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":3571
            Key             =   "verify"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Ilscolor 
      Left            =   2925
      Top             =   1665
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":378B
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":39AB
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":3BCB
            Key             =   "Parent"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":3DE7
            Key             =   "Child"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":4003
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":4223
            Key             =   "Raise"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":443F
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":465F
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":487F
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":4A9F
            Key             =   "View"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":4CBF
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":4EDF
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":50FF
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":5419
            Key             =   "verify"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3225
      Left            =   4500
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3225
      ScaleWidth      =   45
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1920
      Width           =   45
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   11250
      _CBHeight       =   780
      _Version        =   "6.7.8988"
      Child1          =   "Toolbar1"
      MinHeight1      =   720
      Width1          =   9795
      Key1            =   "only"
      NewRow1         =   0   'False
      Caption2        =   "查找"
      Child2          =   "txtFind"
      MinHeight2      =   300
      Width2          =   1080
      NewRow2         =   0   'False
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   10410
         TabIndex        =   41
         Top             =   240
         Width           =   750
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   720
         Left            =   165
         TabIndex        =   40
         Top             =   30
         Width           =   9600
         _ExtentX        =   16933
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "Ilsrw"
         HotImageList    =   "Ilscolor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   19
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "Preview"
               Object.ToolTipText     =   "预览"
               Object.Tag             =   "预览"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "Print"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split0"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "分类"
               Key             =   "Parent"
               Object.ToolTipText     =   "增加分类"
               Object.Tag             =   "分类"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "项目"
               Key             =   "Child"
               Object.ToolTipText     =   "增加项目"
               Object.Tag             =   "项目"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "Modify"
               Object.ToolTipText     =   "修改"
               Object.Tag             =   "修改"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "Delete"
               Object.ToolTipText     =   "删除"
               Object.Tag             =   "删除"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split1"
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "调价"
               Key             =   "Raise"
               Description     =   "调价"
               Object.ToolTipText     =   "调价"
               Object.Tag             =   "调价"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "审核"
               Key             =   "RaiseVerify"
               Object.ToolTipText     =   "审核"
               Object.Tag             =   "审核"
               ImageIndex      =   14
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split2"
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "启用"
               Key             =   "Start"
               Object.ToolTipText     =   "启用"
               Object.Tag             =   "启用"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "停用"
               Key             =   "Stop"
               Object.ToolTipText     =   "停用"
               Object.Tag             =   "停用"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split3"
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查找"
               Key             =   "Find"
               Object.ToolTipText     =   "查找"
               Object.Tag             =   "查找"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查看"
               Key             =   "View"
               Object.ToolTipText     =   "查看方式"
               Object.Tag             =   "查看"
               ImageIndex      =   10
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "  大图标"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "  小图标"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "  列表"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "  详细资料"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split4"
               Style           =   3
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Quit"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   12
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   3615
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":5633
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":5A8B
            Key             =   "ItemNo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":5EDF
            Key             =   "ItemR"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":6D31
            Key             =   "ItemRNo"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   2940
      Top             =   1020
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
            Picture         =   "frmChargeManage.frx":7B83
            Key             =   "RootS"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":7CDD
            Key             =   "Exp"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":7E37
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":8289
            Key             =   "RootR"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":86DB
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":8B33
            Key             =   "ItemNo"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":8F87
            Key             =   "Write"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":93DB
            Key             =   "Read"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":982F
            Key             =   "ItemR"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeManage.frx":A681
            Key             =   "ItemRNo"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picPage 
      BorderStyle     =   0  'None
      Height          =   1650
      Index           =   2
      Left            =   6660
      ScaleHeight     =   1650
      ScaleWidth      =   900
      TabIndex        =   17
      Top             =   2520
      Width           =   900
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   3000
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   5025
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'None
            Height          =   2130
            Left            =   120
            TabIndex        =   26
            Top             =   720
            Width           =   4875
            Begin MSComctlLib.ListView lvwOutIn 
               Height          =   1230
               Left            =   255
               TabIndex        =   27
               Top             =   360
               Width           =   4605
               _ExtentX        =   8123
               _ExtentY        =   2170
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   0
            End
         End
         Begin VB.OptionButton opt科室 
            Caption         =   "无明确执行科室"
            Height          =   255
            Index           =   0
            Left            =   45
            TabIndex        =   25
            Top             =   105
            Width           =   1590
         End
         Begin VB.OptionButton opt科室 
            Caption         =   "病人所在病区"
            Height          =   255
            Index           =   2
            Left            =   45
            TabIndex        =   24
            Top             =   405
            Width           =   1470
         End
         Begin VB.OptionButton opt科室 
            Caption         =   "操作员所在科室"
            Height          =   255
            Index           =   3
            Left            =   1725
            TabIndex        =   23
            Top             =   435
            Width           =   1665
         End
         Begin VB.OptionButton opt科室 
            Caption         =   "指定科室"
            Height          =   255
            Index           =   4
            Left            =   210
            TabIndex        =   22
            Top             =   750
            Width           =   1170
         End
         Begin VB.OptionButton opt科室 
            Caption         =   "病人所在科室"
            Height          =   255
            Index           =   1
            Left            =   1725
            TabIndex        =   21
            Top             =   120
            Width           =   1530
         End
         Begin VB.OptionButton opt科室 
            Caption         =   "院外执行"
            Height          =   195
            Index           =   5
            Left            =   3480
            TabIndex        =   20
            Top             =   135
            Visible         =   0   'False
            Width           =   1860
         End
         Begin VB.OptionButton opt科室 
            Caption         =   "开单人所在科室"
            Height          =   195
            Index           =   6
            Left            =   3480
            TabIndex        =   19
            Top             =   465
            Width           =   1860
         End
         Begin VB.Label lblMsg 
            Caption         =   "本页面仅供查看，如需修改请双击。"
            Height          =   375
            Left            =   1560
            TabIndex        =   39
            Top             =   0
            Width           =   2895
         End
      End
   End
   Begin XtremeSuiteControls.TabControl tbPage 
      Height          =   990
      Left            =   4860
      TabIndex        =   35
      Top             =   5370
      Width           =   2055
      _Version        =   589884
      _ExtentX        =   3625
      _ExtentY        =   1746
      _StockProps     =   64
   End
   Begin MSComctlLib.ListView lvwWholeSetItem_S 
      Height          =   1095
      Left            =   6960
      TabIndex        =   37
      Top             =   810
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1931
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils32"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFileset 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFilepre 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuFileSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileStdImp 
         Caption         =   "标准导入(&I)"
      End
      Begin VB.Menu mnuFileStdCheck 
         Caption         =   "标准核查(&C)"
      End
      Begin VB.Menu mnuFileSplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileParameter 
         Caption         =   "参数设置(&R)"
      End
      Begin VB.Menu mnuFileSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileexit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEditWholeSet 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditWholeSetClassAdd 
         Caption         =   "增加成套分类(&N)"
      End
      Begin VB.Menu mnuEditWholeSetClassModify 
         Caption         =   "修改成套分类(&M)"
      End
      Begin VB.Menu mnuEditWholeSetClassDelete 
         Caption         =   "删除成套分类(&L)"
      End
      Begin VB.Menu mnuEditWholeSplit 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuEditWholeSetItemAdd 
         Caption         =   "增加成套项目(&C)"
      End
      Begin VB.Menu mnuEditWholeSetItemModify 
         Caption         =   "修改成套项目(&I)"
      End
      Begin VB.Menu mnuEditWholeSetItemDelete 
         Caption         =   "删除成套项目(&D)"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditParent 
         Caption         =   "增加分类(&N)"
      End
      Begin VB.Menu mnuEditModifyAssort 
         Caption         =   "修改分类(&M)"
      End
      Begin VB.Menu mnuEditDeleteAssort 
         Caption         =   "删除分类(&L)"
      End
      Begin VB.Menu mnuEditSplit0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditChild 
         Caption         =   "增加项目(&C)"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "复制新增(&O)"
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "修改项目(&I)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "删除项目(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDept 
         Caption         =   "执行科室(&P)"
      End
      Begin VB.Menu mnuEditSlave 
         Caption         =   "从属项目(&V)"
      End
      Begin VB.Menu mnuEditItemGroup 
         Caption         =   "项目组成(&G)"
      End
      Begin VB.Menu mnuEditSplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClassEdit 
         Caption         =   "类别编辑(&A)"
      End
      Begin VB.Menu mnuEditExcel 
         Caption         =   "导入项目"
      End
      Begin VB.Menu mnuEditSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditStart 
         Caption         =   "启用(&S)"
      End
      Begin VB.Menu mnuEditStop 
         Caption         =   "停用(&T)"
      End
   End
   Begin VB.Menu mnuPrice 
      Caption         =   "价目管理(&T)"
      Begin VB.Menu mnuPriceRaise 
         Caption         =   "调价(&R)"
      End
      Begin VB.Menu mnuPriceRaiseMass 
         Caption         =   "批量调价(&P)"
      End
      Begin VB.Menu mnuPriceRaiseVerify 
         Caption         =   "调价审核(&V)"
      End
      Begin VB.Menu mnuPriceHistory 
         Caption         =   "删除未执行价格(&E)"
      End
      Begin VB.Menu mnuPriceChargeSet 
         Caption         =   "费别设置(&C)"
      End
      Begin VB.Menu mnuPriceReport 
         Caption         =   "价目表(&J)"
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
         Begin VB.Menu mnuViewToolspilt1 
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
      Begin VB.Menu mnuviewsplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "大图标(&G)"
         Index           =   0
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "小图标(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "列表(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "详细资料(&D)"
         Checked         =   -1  'True
         Index           =   3
      End
      Begin VB.Menu mnuViewSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewShowAll 
         Caption         =   "显示所有下级(&H)"
      End
      Begin VB.Menu mnuViewShowStop 
         Caption         =   "显示停用项目(&P)"
      End
      Begin VB.Menu mnuViewSplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewSelect 
         Caption         =   "选择列(&C)"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "查找(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "刷新(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpTopic 
         Caption         =   "帮助主题(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "Web上的中联"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "中联主页(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "中联论坛(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "发送反馈(&K)..."
         End
      End
      Begin VB.Menu mnuHelpSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
   Begin VB.Menu mnuShort1 
      Caption         =   "快捷菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuShortMenu1 
         Caption         =   "增加分类(&P)"
         Index           =   1
      End
      Begin VB.Menu mnuShortMenu1 
         Caption         =   "修改(&M)"
         Index           =   2
      End
      Begin VB.Menu mnuShortMenu1 
         Caption         =   "删除(&D)"
         Index           =   3
      End
   End
   Begin VB.Menu mnuShort2 
      Caption         =   "快捷菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "增加项目(&C)"
         Index           =   0
      End
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "复制新增(&O)"
         Index           =   1
      End
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "修改(&M)"
         Index           =   2
      End
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "执行科室(&P)"
         Index           =   3
      End
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "从属项目(&V)"
         Index           =   4
      End
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "删除(&D)"
         Index           =   5
      End
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "项目组成"
         Index           =   6
      End
      Begin VB.Menu mnuShortsplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "大图标(&G)"
         Index           =   0
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "小图标(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "列表(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "详细资料(&D)"
         Checked         =   -1  'True
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmChargeManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng编码长度 As Long
Dim mintColumn1 As Integer
Dim msngStartX As Single, msngStartY As Single    '移动前鼠标的位置
Dim mblnItem As Boolean  '为真表示单击到ListView某一项上
Dim mblnLoad As Boolean
Dim mintColumn As Integer
Dim mstrKey As String       '前一个树节点的关键值
Dim mstr上级Key As String   '当前项目的上级树节点的Key值，主要是由于显示所有下级项目
Dim mstrClass As String     '用来记录类别编码
Dim mstrClassName As String '用来记录类别名称
Dim mbln启动医价系统 As Boolean '是否已启用医价系统
Private Const mstrLvw As String = "名称,1500,0,1;编码,1000,0,2;标识主码,1400,0,0;标识子码,900,0,0;备选码,1400,0,0;" & _
                                "规格,550,0,0;计算单位,900,0,0;所属类别,1000,0,2;费用类型,900,0,0;服务对象,900,0,0;" & _
                                "说明,1440,0,0;屏蔽费别,900,0,0;是否变价,900,0,0;加班加价,900,0,0;补充摘要,900,0,0;" & _
                                "项目特性,1100,0,2;最高限价,1000,1,0;最低限价,1000,1,0;建档时间,1100,0,0;撤档时间,1100,0,0;" & _
                                "所属分类,1300,0,2;病案费目,1300,0,2;院区,0,0,2"
Private Const mstrLvwWholeSet As String = "名称,1500,0,1;编码,800,0,2;拼音,1400,0,0;五笔,1400,0,0;使用范围,1000,0,0;所属分类,2400,0,0"
Private mlngMode As Long
Public mstrPrivs As String                              '权限串
Private mint上次细目页 As Integer '
Private mint上次成套页 As Integer '
Private mfrmEarnRS As New frmEarnRS
Private mblnVerifyFlow As Boolean   '调价是否启用了审核流程，true-启用，false-未启用
Private mblnVerifyPris As Boolean   '审核调价单权限 true-有权限，false-无权限

Private Enum mCalssPage
    pg_细目 = 1
    pg_成套 = 2
End Enum
Private Enum mItemPage
    pg_价目 = 1
    pg_执行科室 = 2
    pg_从属项目 = 3
    pg_别名 = 4
    pg_费别等级 = 5
    pg_成套组成 = 6
    pg_成套使用科室 = 7
End Enum

Private Enum mIndex价目
    Col_价格等级 = 0
    Col_单据号
    Col_执行日期
    Col_终止日期
    Col_收入项目
    Col_原价
    Col_现价
    Col_附加手术收费率
    Col_加班加价率
    Col_调价说明
    Col_缺省价格
    Col_调价人
End Enum

Private mblnNotClick As Boolean
Private mblnCanUpdateAll As Boolean '是否允许操作所有项目：未启用价格等级或启用了价格等级有“所有院区”权限

Private Sub zlInitClassPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化分类页面
    '编制:刘兴洪
    '日期:2010-08-24 10:15:11
    '说明:27327
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, ObjItem As TabControlItem
    Err = 0: On Error GoTo ErrHand:
    mblnNotClick = True
    Set ObjItem = tbClassPage.InsertItem(mCalssPage.pg_细目, "收费细目", picTreeItem.hwnd, 0)
    ObjItem.Tag = mCalssPage.pg_细目
    ObjItem.Selected = True
    Set ObjItem = tbClassPage.InsertItem(mCalssPage.pg_成套, "成套项目", picTreeWholeSet.hwnd, 0)
    ObjItem.Tag = mCalssPage.pg_成套
     With tbClassPage
        
        .PaintManager.Appearance = xtpTabAppearanceVisio
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = True
        .PaintManager.ClientFrame = xtpTabFrameBorder
    End With
    
    mblnNotClick = False
    Set ObjItem = tbPage.InsertItem(mItemPage.pg_价目, "收费价目", picPage(1).hwnd, 0)
    ObjItem.Tag = mItemPage.pg_价目
    Set ObjItem = tbPage.InsertItem(mItemPage.pg_执行科室, "执行科室", picPage(2).hwnd, 0)
    ObjItem.Tag = mItemPage.pg_执行科室
    Set ObjItem = tbPage.InsertItem(mItemPage.pg_从属项目, "从属项目", picPage(3).hwnd, 0)
    ObjItem.Tag = mItemPage.pg_从属项目
    Set ObjItem = tbPage.InsertItem(mItemPage.pg_别名, "别名", picPage(4).hwnd, 0)
    ObjItem.Tag = mItemPage.pg_别名
    Set ObjItem = tbPage.InsertItem(mItemPage.pg_费别等级, "费别等级", picPage(5).hwnd, 0)
    ObjItem.Tag = mItemPage.pg_费别等级
    Set ObjItem = tbPage.InsertItem(mItemPage.pg_成套组成, "成套项目组成", picPage(6).hwnd, 0)
    ObjItem.Tag = mItemPage.pg_成套组成
    Set ObjItem = tbPage.InsertItem(mItemPage.pg_成套使用科室, "成套项目使用科室", picPage(7).hwnd, 0)
    ObjItem.Tag = mItemPage.pg_成套使用科室
    tbPage.Item(0).Selected = True
    mint上次成套页 = mItemPage.pg_成套组成
    mint上次细目页 = mItemPage.pg_价目
    
     With tbPage
        .PaintManager.Appearance = xtpTabAppearancePropertyPage
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = True
        .PaintManager.ClientFrame = xtpTabFrameBorder
    End With
    Call SetPageVisible
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function Check收费项目(ByVal ID As Long, ByRef strMsg As String) As Boolean
    '检查收费项目存在的各种依赖关系
    Dim rs As New ADODB.Recordset
        
    On Error GoTo ErrHandle

    '1.该收费项目是否被设置在诊疗收费对照中。
    gstrSQL = "Select 1 From 诊疗收费关系 where RowNum=1 and 收费项目id=[1] "
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "检查诊疗收费对照。", ID)
    
    strMsg = IIF(rs.RecordCount = 0, "", "[该项目存在诊疗收费对照关系！]" & vbCrLf)
        
    '2.该收费项目是否被设置为其它项目的从属项目。
    gstrSQL = "Select 1 From 收费从属项目 where RowNum=1 and (主项id=[1] or 从项id=[1] )"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "检查收费从属项目。", ID)
    
    strMsg = strMsg & IIF(rs.RecordCount = 0, "", "[该项目存在收费从属关系！]" & vbCrLf)
    
    '3.该收费项目是否特定收费项目。
    gstrSQL = "Select 1 From 收费特定项目 where RowNum=1 and 收费细目id=[1] "
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "检查收费特定项目。", ID)
    
    strMsg = strMsg & IIF(rs.RecordCount = 0, "", "[该项目是收费特定项目！]" & vbCrLf)
    
    '4.该收费项目是否被设置为自动计价项目。
    gstrSQL = "Select 1 From 自动计价项目 where RowNum=1 and 收费细目id=[1] "
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "检查自动计价项目。", ID)
    
    strMsg = strMsg & IIF(rs.RecordCount = 0, "", "[该项目是自动计价项目！]" & vbCrLf)
    
    Check收费项目 = True
    Exit Function
    
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    
End Function


Private Sub Form_Activate()
    On Error GoTo ErrHandle
    
    mblnVerifyFlow = IIF(Val(zlDatabase.GetPara("调价需要审核", glngSys, 1009, 0)) = 0, False, True)
    mblnVerifyPris = IIF(InStr(1, ";" & gstrPrivs & ";", ";收费价目调价审核;") > 0, True, False)
    If mblnLoad = False Then Exit Sub
    Call Form_Resize '为了使CoolBar自适应高度
    Call FillTree:
    Call FillWholeSetTree
    mblnLoad = False
    
    If checkNotPrice(0) = False Then
        MsgBox "收费细目中还存在未审核的价格，请注意审核！", vbInformation, gstrSysName
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub zlInitLvwHeadCol()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化ListView列头
    '编制:刘兴洪
    '日期:2010-08-24 14:21:48
    '说明:27327
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '如果ListView的还未被设置，比如第一次使用，那就调用缺省的初始化
    lvwWholeSetItem_S.ListItems.Clear
    zlControl.LvwSelectColumns lvwWholeSetItem_S, mstrLvwWholeSet, True
    lvwMain_S.ListItems.Clear
    zlControl.LvwSelectColumns lvwMain_S, mstrLvw, True
End Sub

Private Sub Form_Load()
    Dim intType  As Integer
    On Error GoTo ErrHandle
    mblnLoad = True
    
    mlngMode = glngModul
    mstrPrivs = gstrPrivs
    mblnCanUpdateAll = IsPriceGradeEnabled() = False Or zlStr.IsHavePrivs(mstrPrivs, "所有院区")   '110070
    
    Call GetPriceGrade(gstr药品价格等级, gstr卫材价格等级, gstr普通价格等级)
    Call zlInitClassPage
    
    '权限
    Call 权限控制

    '允许进行列删除的ListView须做标记
    lvwMain_S.Tag = "可变化的"
    lvwWholeSetItem_S.Tag = "可变化的"
    mnuViewShowAll.Checked = (Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "显示所有", 0)) = 1)
    mnuViewShowStop.Checked = (Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "显示停用", 0)) = 1)
    chk价格.value = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "显示历史价格", "0"))
    
    Call zlInitLvwHeadCol
    
    '根据lvwMain_S显示设置对应菜单
     mnuViewIcon_Click lvwMain_S.View
    
    '初始化右下角的执行科室栏
    zlControl.LvwSelectColumns lvwOutIn, "执行科室,3000,0,0;病人科室,8000,0,0", True
    zlControl.LvwFlatColumnHeader lvwOutIn
    
    RestoreWinState Me, App.ProductName
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs, "ZL1_INSIDE_1009")
    
    Call InitTable
    Call GetDefineSize
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitTable()
On Error GoTo ErrHandle
    '初始化收费价目表和从属项目表
    msh价目.Cols = 12
    
    msh价目.ColWidth(Col_价格等级) = 250
    msh价目.ColWidth(Col_单据号) = 1000
    msh价目.ColWidth(Col_执行日期) = 2000
    msh价目.ColWidth(Col_终止日期) = 2000
    msh价目.ColWidth(Col_收入项目) = 1000
    msh价目.ColWidth(Col_原价) = 1000
    msh价目.ColWidth(Col_现价) = 1000
    msh价目.ColWidth(Col_附加手术收费率) = 1000
    msh价目.ColWidth(Col_加班加价率) = 1000
    msh价目.ColWidth(Col_调价说明) = 3000
    msh价目.ColWidth(Col_缺省价格) = 1000
    msh价目.ColWidth(Col_调价人) = 800
    
    msh从属.ColWidth(0) = 1000
    msh从属.ColWidth(1) = 3000
    msh从属.ColWidth(2) = 1000
    msh从属.ColWidth(3) = 1500
    
    msh价目.TextMatrix(0, Col_价格等级) = ""
    msh价目.TextMatrix(0, Col_单据号) = "单据号"
    msh价目.TextMatrix(0, Col_执行日期) = "执行日期"
    msh价目.TextMatrix(0, Col_终止日期) = "终止日期"
    msh价目.TextMatrix(0, Col_收入项目) = "收入项目"
    msh价目.TextMatrix(0, Col_原价) = "原价"
    msh价目.TextMatrix(0, Col_现价) = "现价"
    msh价目.TextMatrix(0, Col_附加手术收费率) = "附加手术收费率"
    msh价目.TextMatrix(0, Col_加班加价率) = "加班加价率"
    msh价目.TextMatrix(0, Col_调价说明) = "调价说明"
    msh价目.TextMatrix(0, Col_缺省价格) = "缺省价格"
    msh价目.TextMatrix(0, Col_调价人) = "调价人"
   
    mshAlias.Cols = 4
    mshAlias.ColWidth(0) = 1000
    mshAlias.ColWidth(1) = 4000
    mshAlias.ColWidth(2) = 800
    mshAlias.ColWidth(3) = 3000
    
    mshAlias.TextMatrix(0, 0) = "名称种类"
    mshAlias.TextMatrix(0, 1) = "名称"
    mshAlias.TextMatrix(0, 2) = "码类"
    mshAlias.TextMatrix(0, 3) = "简码"
    
    msh从属.TextMatrix(0, 0) = "收费类别"
    msh从属.TextMatrix(0, 1) = "收费项目"
    msh从属.TextMatrix(0, 2) = "次数"
    msh从属.TextMatrix(0, 3) = "固定"
    msh从属.TextMatrix(0, 4) = "状态"
    msh从属.Col = 0
    msh从属.Row = 0
    msh从属.ColSel = 3
    msh从属.RowSel = 0
    msh从属.FillStyle = flexFillRepeat
    msh从属.CellAlignment = 4
    msh从属.FillStyle = flexFillSingle
    msh从属.AllowBigSelection = False
    msh从属.Row = 1
    msh从属.ColAlignment(3) = 1
    msh从属.ColAlignment(0) = 1
    msh从属.ColAlignment(1) = 1
    
    msh价目.ColAlignment(1) = 1
    msh价目.FillStyle = flexFillRepeat
    msh价目.CellAlignment = 4
    msh价目.FillStyle = flexFillSingle
    msh价目.AllowBigSelection = False
    msh价目.Row = 1
    msh价目.FixedAlignment(-1) = flexAlignCenterCenter
    msh价目.HighLight = flexHighlightNever
    
    mshAlias.Col = 0
    mshAlias.Row = 1
    mshAlias.ColSel = 1
    mshAlias.RowSel = 0
    mshAlias.FillStyle = flexFillRepeat
    mshAlias.CellAlignment = 4
    mshAlias.FillStyle = flexFillSingle
    mshAlias.Row = 1
    
    With msh费别
        .Cols = 4
        .ColWidth(0) = 1500
        .ColWidth(1) = 3000
        .ColWidth(2) = 1050
        .ColWidth(3) = 2000
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .TextMatrix(0, 0) = "费别"
        .TextMatrix(0, 1) = "应收金额(元)"
        .TextMatrix(0, 2) = "实收比率(%)"
        .TextMatrix(0, 3) = "计算方法"
        
        .MergeCol(0) = True
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
 

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandle
    mstrKey = ""
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "显示所有", IIF(mnuViewShowAll.Checked, 1, 0)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "显示停用", IIF(mnuViewShowStop.Checked, 1, 0)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "显示历史价格", chk价格.value
    zl_vsGrid_Para_Save mlngMode, vsWholeSet, Me.Caption, "成套项目组成表列-主界面", True, True
    SaveWinState Me, App.ProductName
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
 
Private Sub lvwWholeSetItem_S_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo ErrHandle
    If mintColumn1 = ColumnHeader.Index - 1 Then '仍是刚才那列
        lvwWholeSetItem_S.SortOrder = IIF(lvwWholeSetItem_S.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn1 = ColumnHeader.Index - 1
        lvwWholeSetItem_S.SortKey = mintColumn1
        lvwWholeSetItem_S.SortOrder = lvwAscending
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwWholeSetItem_S_DblClick()
    Dim strID As String, lng上级ID As Long
    If Not mblnItem Then Exit Sub
    If mnuEditWholeSetClassModify.Enabled And mnuEditWholeSetClassModify.Visible Then
        Call mnuEditWholeSetItemModify_Click
    Else
        If Me.lvwWholeSetItem_S.SelectedItem Is Nothing Then Exit Sub
        With lvwWholeSetItem_S
            strID = Val(Mid(.SelectedItem.Key, 2))
        End With
        If frmChargeWholeSetItemEdit.ShowCard(Me, EdI_查看, mstrPrivs, mlngMode, lng上级ID, strID) = False Then Exit Sub
    End If
End Sub

Private Sub lvwWholeSetItem_S_GotFocus()
    Call MenuSet
End Sub

Private Sub lvwWholeSetItem_S_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '刘兴洪:27327
    '为成套项目维护时,需要单独处理
    mblnItem = True
    If lvwWholeSetItem_S.Tag <> Item.Key Then
        Call FillWholeSetItemChildData(Val(Mid(Item.Key, 2)))
    End If
    lvwWholeSetItem_S.Tag = Item.Key
End Sub

Private Sub lvwWholeSetItem_S_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If mnuEditWholeSetItemModify.Enabled And mnuEditWholeSetItemModify.Visible Then mnuEditWholeSetItemModify_Click
    End If
End Sub

Private Sub lvwWholeSetItem_S_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnItem = False
End Sub

Private Sub lvwWholeSetItem_S_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mnuEditWholeSet.Visible = False Then Exit Sub
    If Button <> 2 Then Exit Sub
    PopupMenu mnuEditWholeSet, vbPopupMenuRightButton
End Sub

Private Sub mnuEditExcel_Click()
    frmItemImport.ShowMe 1, Me
    Call FillTree
    
End Sub

Private Sub mnuEditWholeSetClassAdd_Click()
    Dim lng上级ID As Long
    With tvwWholeSet
        If .SelectedItem Is Nothing Then Exit Sub
        lng上级ID = Val(Mid(.SelectedItem.Key, 2))
    End With
    If InStr(1, mstrPrivs, ";增加成套项目;") = 0 Then Exit Sub
    If frmChargeWholeSetClassEdit.EditCard(Me, Ed_增加, mstrPrivs, mlngMode, lng上级ID, "") = False Then Exit Sub
    Call FillWholeSetTree

End Sub

Private Sub mnuEditWholeSetClassDelete_Click()
    Dim strKey As String, intIndex As Long
    On Error GoTo ErrHandle
    With tvwWholeSet
        If .SelectedItem Is Nothing Then Exit Sub
        If .SelectedItem.Key = "Root" Then Exit Sub
        If MsgBox("你确认要删除名称为“" & .SelectedItem.Text & "”的分类吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
            Me.MousePointer = 11
            gstrSQL = "Zl_成套项目分类_Delete(" & Val(Mid(.SelectedItem.Key, 2)) & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            Me.MousePointer = 0
             
            strKey = .SelectedItem.Key
            If Not .SelectedItem.Next Is Nothing Then
                 .SelectedItem.Next.Selected = True
                tvwWholeSet_NodeClick tvwWholeSet.SelectedItem
            Else
                If Not .SelectedItem.Parent Is Nothing Then
                     .SelectedItem.Parent.Selected = True
                End If
                If Not .SelectedItem Is Nothing Then
                    tvwWholeSet_NodeClick tvwWholeSet.SelectedItem
                End If
            End If
             .Nodes.Remove strKey
        End If
    End With
    MenuSet
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Me.MousePointer = 0
End Sub

Private Sub mnuEditWholeSetClassModify_Click()
    Dim lng上级ID As Long, lngId As Long
    If InStr(1, mstrPrivs, ";修改成套项目;") = 0 Then Exit Sub
    With tvwWholeSet
        If .SelectedItem Is Nothing Then Exit Sub
        If .SelectedItem.Key = "Root" Then
            lng上级ID = 0
            lngId = 0
        Else
            lng上级ID = Val(Mid(.SelectedItem.Parent.Key, 2))
            lngId = Mid(.SelectedItem.Key, 2)
        End If
    End With
    If frmChargeWholeSetClassEdit.EditCard(Me, Ed_修改, mstrPrivs, mlngMode, lng上级ID, lngId) = False Then Exit Sub
    Call FillWholeSetTree
End Sub

Private Sub mnuEditWholeSetItemAdd_Click()
    '成套项目增加
    Dim lng上级ID As Long
    With tvwWholeSet
        If .SelectedItem Is Nothing Then Exit Sub
        lng上级ID = Val(Mid(.SelectedItem.Key, 2))
    End With
    If InStr(1, mstrPrivs, ";增加成套项目;") = 0 Then Exit Sub
    If frmChargeWholeSetItemEdit.ShowCard(Me, EdI_增加, mstrPrivs, mlngMode, lng上级ID, "") = False Then Exit Sub
    Call FillWholeItem(lng上级ID)
End Sub

Private Sub mnuEditWholeSetItemDelete_Click()
    '删除项目
    Dim lngId As Long, strKey As String
    Dim intIndex As Long
     
    '修改项目
    If Not (mnuEditWholeSetItemModify.Enabled And mnuEditWholeSetItemModify.Visible) Then Exit Sub
    If Me.lvwWholeSetItem_S.SelectedItem Is Nothing Then Exit Sub
    
    If InStr(1, mstrPrivs, ";修改成套项目;") = 0 Then Exit Sub
    With lvwWholeSetItem_S
        lngId = Val(Mid(.SelectedItem.Key, 2))
        strKey = .SelectedItem.Key
    End With
    If MsgBox("你确认要删除名称为“" & lvwWholeSetItem_S.SelectedItem.Text & "”的成套收费项目吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    On Error GoTo ErrHandle
    Me.MousePointer = 11
    'Zl_成套收费项目_Delete(Id_In In 成套收费项目.ID%Type)
    gstrSQL = "Zl_成套收费项目_Delete(" & lngId & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Me.MousePointer = 0
    With lvwWholeSetItem_S
        intIndex = .SelectedItem.Index
        .ListItems.Remove .SelectedItem.Key
        If .ListItems.Count > 0 Then
            intIndex = IIF(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
            .ListItems(intIndex).Selected = True
            .ListItems(intIndex).EnsureVisible
            FillWholeSetItemChildData Val(Mid(.SelectedItem.Key, 2))
        Else
            FillWholeSetItemChildData 0
        End If
    End With
    MenuSet
    Me.MousePointer = 0
    Exit Sub
ErrHandle:
    Me.MousePointer = 0
    If ErrCenter() = 1 Then
        Me.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuEditWholeSetItemModify_Click()
    Dim strID As String, lng上级ID As Long
    '修改项目
    If Not (mnuEditWholeSetItemModify.Enabled And mnuEditWholeSetItemModify.Visible) Then Exit Sub
    If Me.lvwWholeSetItem_S.SelectedItem Is Nothing Then Exit Sub
    
    If InStr(1, mstrPrivs, ";修改成套项目;") = 0 Then Exit Sub
    With lvwWholeSetItem_S
        strID = Val(Mid(.SelectedItem.Key, 2))
    End With
    With tvwWholeSet
        If .SelectedItem Is Nothing Then
            lng上级ID = 0
        Else
            lng上级ID = Val(Mid(.SelectedItem.Key, 2))
        End If
    End With
    If frmChargeWholeSetItemEdit.ShowCard(Me, EdI_修改, mstrPrivs, mlngMode, lng上级ID, strID) = False Then Exit Sub
    Call FillWholeItem(lng上级ID)
End Sub

Private Sub mnuFileParameter_Click()
    frmParSetFeeItem.ShowMe Me
    
    mblnVerifyFlow = IIF(Val(zlDatabase.GetPara("调价需要审核", glngSys, 1009, 0)) = 0, False, True)
    mnuPriceRaiseVerify.Visible = mblnVerifyFlow
    Toolbar1.Buttons("RaiseVerify").Visible = mblnVerifyFlow   '调价审核
End Sub

Private Sub mnuPriceRaiseVerify_Click()
    frmChargePriceVerify.ShowMe Me, mblnCanUpdateAll
End Sub

Private Sub msh价目_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = Col_价格等级 Then Cancel = True
End Sub

Private Sub picPage_DblClick(Index As Integer)
    
    On Error GoTo ErrHandle
    Select Case Index
        Case 1
            If mnuPriceRaise.Enabled = True And mnuPriceRaise.Visible = True Then Call mnuPriceRaise_Click
        Case 2
            If mnuEditDept.Enabled = True And mnuEditDept.Visible = True Then Call mnuEditDept_Click
        Case 3
            If mnuEditSlave.Enabled = True And mnuEditSlave.Visible = True Then Call mnuEditSlave_Click
        Case 4
            If mnuEditModify.Enabled = True And mnuEditModify.Visible = True Then Call mnuEditModify_Click
        Case 5
            If mnuPriceChargeSet.Enabled = True And mnuPriceChargeSet.Visible = True Then Call mnuPriceChargeSet_Click
    End Select
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwMain_S_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo ErrHandle
    If mintColumn = ColumnHeader.Index - 1 Then '仍是刚才那列
        lvwMain_S.SortOrder = IIF(lvwMain_S.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwMain_S.SortKey = mintColumn
        lvwMain_S.SortOrder = lvwAscending
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwMain_S_DblClick()
On Error GoTo ErrHandle
    If mblnItem = True And mnuEditModify.Enabled And mnuEditModify.Visible Then mnuEditModify_Click
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwMain_S_KeyPress(KeyAscii As Integer)
On Error GoTo ErrHandle
    If KeyAscii = vbKeyReturn Then
        If mnuEditModify.Enabled And mnuEditModify.Visible Then mnuEditModify_Click
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwMain_S_GotFocus()
On Error GoTo ErrHandle
    Call MenuSet
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub lvwMain_S_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long
    Dim strTips As String
    Dim intName As Integer
    Dim intArrit As Integer
    Dim intAssort As Integer
    Dim intCode As Integer

On Error GoTo ErrHandle
    mstrClass = ""
    mstrClassName = ""
    If Item Is Nothing Then Exit Sub
    
    '当类别不为挂号、护理、床位的
    strTips = ""
    For i = 1 To lvwMain_S.ColumnHeaders.Count - 1
        If lvwMain_S.ColumnHeaders(i).Text = "项目特性" Then
            intArrit = i - 1
        ElseIf lvwMain_S.ColumnHeaders(i).Text = "名称" Then
            intName = i - 1
        ElseIf lvwMain_S.ColumnHeaders(i).Text = "病案费目" Then
            intAssort = i - 1
        ElseIf lvwMain_S.ColumnHeaders(i).Text = "编码" Then
            intCode = i - 1
        End If
    Next
    
    mblnItem = True
    FillItem Item.Key
    
    If lvwMain_S.ListItems.Count > 0 Then
        strTips = "当前分类下查找到 " & lvwMain_S.ListItems.Count & " 条项目"
    Else
        strTips = "当前分类无项目"
    End If
    Me.stbThis.Panels(2).Text = strTips
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwMain_S_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnItem = False
End Sub

Private Sub lvwMain_S_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrHandle
    Dim i As Long
    Dim intArrit As Integer
    For i = 1 To lvwMain_S.ColumnHeaders.Count - 1
        If lvwMain_S.ColumnHeaders(i).Text = "项目特性" Then
            intArrit = i - 1
            Exit For
        End If
    Next
    
    If Button = 2 Then
        If lvwMain_S.ListItems.Count < 1 Then
            mnuEditCopy.Enabled = False
            mnuEditModify.Enabled = False
        End If
        mnuShortMenu2(0).Enabled = mnuEditChild.Enabled
        mnuShortMenu2(1).Enabled = mnuEditCopy.Enabled
        mnuShortMenu2(2).Enabled = mnuEditModify.Enabled
        mnuShortMenu2(3).Enabled = mnuEditDept.Enabled
        mnuShortMenu2(4).Enabled = mnuEditSlave.Enabled
        mnuShortMenu2(5).Enabled = mnuEditDelete.Enabled
        mnuShortMenu2(6).Enabled = mnuEditItemGroup.Enabled
        
        For i = 0 To 3
            mnuShortIcon(i).Checked = mnuViewIcon(i).Checked
        Next
        PopupMenu mnuShort2, vbPopupMenuRightButton
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuClassEdit_Click()
    frmChargeClassEdit.Show 1, Me
End Sub

Private Sub mnuPriceChargeSet_Click()
    If Me.lvwMain_S.ListItems.Count > 0 Then
        If frmChargeSortItemEdit.ShowMe(Me, 2, "", Val(Me.lvwMain_S.SelectedItem.Tag), Me.lvwMain_S.SelectedItem.Text) Then
            Call frmChargeManage.FillTree
        End If
    End If
End Sub

Private Sub mnuEditItemGroup_Click()
    If Me.lvwMain_S.ListItems.Count > 0 Then
        frmChargeGroupItem.ShowMe Me, Me.lvwMain_S.SelectedItem.Tag
    End If
End Sub

Private Sub mnuFileStdCheck_Click()
    frmStdCheck.Show 1, Me
End Sub

Private Sub mnuFileStdImp_Click()
    frmPriceImp.Show 1, Me
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuEditDept_Click()
On Error GoTo ErrHandle
    If mnuEdit.Visible = False Then Exit Sub
    
    ModifyMode 3  'editDept
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuEditSlave_Click()
On Error GoTo ErrHandle
    If mnuEdit.Visible = False Then Exit Sub
    ModifyMode 4    'editSlave
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuPriceHistory_Click()
'删除未执行价格
    Dim strNodeNo As String
    
    If lvwMain_S.SelectedItem Is Nothing Then Exit Sub
    '110133
    '最后一列是院区编码
    strNodeNo = lvwMain_S.SelectedItem.ListSubItems(lvwMain_S.SelectedItem.ListSubItems.Count).Tag
    If mblnCanUpdateAll Or strNodeNo = gstrNodeNo Then
        If MsgBox("你确认要删除最近一次的未执行价格吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Else
        If MsgBox("你确认要删除最近一次的未执行价格吗？" & vbCrLf & vbCrLf & _
            "注意：由于你没有“所有院区”权限，所以只会删除价格等级“" & gstr普通价格等级 & "”最近一次的未执行价格。", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    On Error GoTo ErrHandle
    MousePointer = 11
    'Zl_收费价目_Delete(
    gstrSQL = "zl_收费价目_Delete("
    '  细目id_In In 收费价目.收费细目id%Type,
    gstrSQL = gstrSQL & "" & Val(lvwMain_S.SelectedItem.Tag) & ","
    '  站点_In   In 收费项目目录.站点%Type := Null
    gstrSQL = gstrSQL & "" & IIF(mblnCanUpdateAll, "NULL", "'" & gstrNodeNo & "'") & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    FillItem lvwMain_S.SelectedItem.Key
    MousePointer = 0
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    MousePointer = 0
End Sub

Private Sub mnuPriceRaiseMass_Click()
'批量调价
On Error GoTo ErrHandle
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim datToday As Date
    
    datToday = sys.Currentdate
    If tvwMainItem.SelectedItem Is Nothing Then Exit Sub
    With frmChargeBatchPrice
        .mstrPrivs = mstrPrivs
        .mblnCanUpdateAll = mblnCanUpdateAll
        
        .txtType.Text = mstrClassName
        .lbl类别.Tag = mstrClass
        
        .txtChargeType.Text = tvwMainItem.SelectedItem.Text
        .lbl分类.Tag = tvwMainItem.SelectedItem.Tag
        '求最大执行日期
        strSQL = "select max(B.执行日期) as 最大日期 from " & _
                "(select id from 收费项目目录 where (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) And 是否变价=0 and " & IIF(.lbl分类.Tag = "", "类别=[1] and 分类ID is null", "分类ID=[2] ") & ") A" & _
                ",收费价目 B  Where A.ID = B.收费细目ID "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .lbl类别.Tag, .lbl分类.Tag)
                
        If rsTemp("最大日期") > datToday Then
            .datSingle = rsTemp("最大日期") + 1
        Else
            .datSingle = datToday + 1
        End If
        
        '求最小金额（但不包含变价项目）
        strSQL = "select min(B.现价) as 最小金额 from " & _
                "(select id from 收费项目目录 where (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) And 是否变价=0 and " & IIF(.lbl分类.Tag = "", "类别=[1] and 分类ID is null", "分类ID=[2] ") & ") A" & _
                ",收费价目 B  Where A.ID = B.收费细目ID And Sysdate Between b.执行日期 And Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .lbl类别.Tag, .lbl分类.Tag)
        
        .dblSingle = IIF(IsNull(rsTemp("最小金额")), 0, rsTemp("最小金额"))
        
        strSQL = "select max(B.执行日期) as 最大日期 from " & _
                "(select c.id from 收费项目目录 c,(Select id From 收费分类目录  start with  " & IIF(.lbl分类.Tag = "", "类别=[1] and 上级ID is null", "上级ID=[2] ") & " connect by prior id=上级ID) d  " & _
                " where c.是否变价=0 And c.分类id=d.Id And (c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.撤档时间 Is Null)) A" & _
                ",收费价目 B  Where A.ID = B.收费细目ID "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .lbl类别.Tag, .lbl分类.Tag)
        
        If rsTemp("最大日期") > datToday Then
            .datAll = rsTemp("最大日期") + 1
        Else
            .datAll = datToday + 1
        End If
        
        strSQL = "select min(B.现价) as 最小金额 from " & _
                "(select c.id from 收费项目目录 c,(Select id From 收费分类目录  start with  " & IIF(.lbl分类.Tag = "", "类别=[1] and 上级ID is null", "上级ID=[2] ") & " connect by prior id=上级ID) d  " & _
                " where c.是否变价=0 And c.分类id=d.Id And (c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.撤档时间 Is Null)) A" & _
                ",收费价目 B  Where A.ID = B.收费细目ID And Sysdate Between b.执行日期 And Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .lbl类别.Tag, .lbl分类.Tag)
        
        .dblAll = IIF(IsNull(rsTemp("最小金额")), 0, rsTemp("最小金额"))
                
        rsTemp.Close
        
        .dtpBegin.value = .datSingle
        .dtpBegin.MinDate = .datSingle
    End With
    frmChargeBatchPrice.Show vbModal, Me
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuPriceReport_Click()
'价目表
On Error GoTo ErrHandle
    Dim str类别 As String
    Dim lng分类id As Long
    Dim strCaption As String
    
    If tvwMainItem.Nodes.Count > 0 Then
        lng分类id = Val(Mid(tvwMainItem.SelectedItem.Key, 2))
    End If
    
    On Error Resume Next
    ReportOpen gcnOracle, glngSys, "ZL1_INSIDE_1009", Me, strCaption, gstrUserName, "分类=" & lng分类id, _
        "站点='" & IIF(mblnCanUpdateAll, "全院", gstrNodeNo) & "'"
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub mnuReportItem_Click(Index As Integer)
    '默认参数：分类=分类id，项目=项目id，类别=收费类别名称
    Dim lng分类id As Long
    Dim lng项目id As Long
    Dim str收费类别 As String
    
    If Not tvwMainItem.SelectedItem Is Nothing Then
        lng分类id = Mid(tvwMainItem.SelectedItem.Key, 2)
    End If
    
    If Not lvwMain_S.SelectedItem Is Nothing Then
        lng项目id = Mid(lvwMain_S.SelectedItem.Key, 3)
        str收费类别 = Replace(Replace(lvwMain_S.SelectedItem.SubItems(lvwMain_S.ColumnHeaders("_所属类别").Index - 1), "[", ""), "]", "")
    End If
    
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
        "分类=" & IIF(lng分类id = 0, "", lng分类id), _
        "项目=" & IIF(lng项目id = 0, "", lng项目id), _
        "类别=" & str收费类别)
End Sub

Private Sub mnuViewFind_Click()
    frmChargeItemFind.Show , Me
End Sub

Private Sub mnuViewRefresh_Click()
On Error GoTo ErrHandle
    FillTree
    Call FillWholeSetTree
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuViewShowAll_Click()
On Error GoTo ErrHandle
    mnuViewShowAll.Checked = Not mnuViewShowAll.Checked
    If tvwMainItem.SelectedItem Is Nothing Then
        If tvwMainItem.Nodes.Count > 0 Then
            MsgBox "请选择一下分类！", vbInformation, gstrSysName
        Else
            MsgBox "无任何分类可显示！", vbInformation, gstrSysName
        End If
        Exit Sub
    End If
    FillList tvwMainItem.SelectedItem.Tag
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuViewShowStop_Click()
On Error GoTo ErrHandle
    mnuViewShowStop.Checked = Not mnuViewShowStop.Checked
    If tvwMainItem.Nodes.Count > 0 Then
        FillList tvwMainItem.SelectedItem.Tag
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuShortMenu1_Click(Index As Integer)
On Error GoTo ErrHandle
    Select Case Index
        Case 1
            mnuEditParent_Click
        Case 2
            mnuEditModifyAssort_Click
        Case 3
            mnuEditDeleteAssort_Click
    End Select
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuShortMenu2_Click(Index As Integer)
On Error GoTo ErrHandle
    Select Case Index
        Case 0
            mnuEditChild_Click
        Case 1
            mnuEditCopy_Click
        Case 2
            mnuEditModify_Click
        Case 3
            mnuEditDept_Click
        Case 4
            mnuEditSlave_Click
        Case 5
            mnuEditDelete_Click
        Case 6
            mnuEditItemGroup_Click
        Case 7
            mnuPriceChargeSet_Click
    End Select
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuShortIcon_Click(Index As Integer)
On Error GoTo ErrHandle
    mnuViewIcon_Click Index
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuViewIcon_Click(Index As Integer)
On Error GoTo ErrHandle
    Dim i As Integer
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
        Toolbar1.Buttons("View").ButtonMenus(i + 1).Text = Replace(Toolbar1.Buttons("View").ButtonMenus(i + 1).Text, "●", "  ")
    Next
    mnuViewIcon(Index).Checked = True
    Toolbar1.Buttons("View").ButtonMenus(Index + 1).Text = Replace(Toolbar1.Buttons("View").ButtonMenus(Index + 1).Text, "  ", "●")
    lvwMain_S.View = Index
    lvwWholeSetItem_S.View = Index
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuViewSelect_Click()
On Error GoTo ErrHandle
    If zlControl.LvwSelectColumns(lvwMain_S, mstrLvw) = True Then
        '列有变化就要重新刷新
        If tvwMainItem.Nodes.Count > 0 Then
            FillList tvwMainItem.SelectedItem.Tag
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mshAlias_DblClick()
On Error GoTo ErrHandle
    If mnuEditModify.Enabled And mnuEditModify.Visible = True Then mnuEditModify_Click
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub msh从属_DblClick()
On Error GoTo ErrHandle
    If mnuEditSlave.Enabled = True And mnuEditSlave.Visible Then Call mnuEditSlave_Click
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub msh费别_DblClick()
    Dim strCharge As String
    On Error GoTo ErrHandle
    
    If InStr(mstrPrivs, "费别设置") = 0 Then Exit Sub
    
    If msh费别.Rows > 1 Then
        If msh费别.TextMatrix(msh费别.Rows - 1, 0) <> "" Then
            strCharge = msh费别.TextMatrix(msh费别.Row, 0)
        End If
    End If
    
    If mnuPriceChargeSet.Enabled = True Then
        If lvwMain_S.SelectedItem Is Nothing Then Exit Sub
        If frmChargeSortItemEdit.ShowMe(Me, 2, strCharge, Val(Me.lvwMain_S.SelectedItem.Tag), Me.lvwMain_S.SelectedItem.Text) Then
            Call frmChargeManage.FillTree
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub msh价目_DblClick()
On Error GoTo ErrHandle
    If mnuPriceRaise.Enabled = True Then Call mnuPriceRaise_Click
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub msh价目_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrHandle
    If Button = 2 Then
        If InStr(mstrPrivs, "价目管理") > 0 Then PopupMenu mnuPrice, vbPopupMenuRightButton
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

 
Private Sub picTreeClass_S_Resize()
    Err = 0: On Error Resume Next
    With tbClassPage
        .Top = picTreeClass_S.ScaleTop
        .Height = picTreeClass_S.ScaleHeight
        .Left = picTreeClass_S.ScaleLeft
        .Width = picTreeClass_S.ScaleWidth
    End With
End Sub

Private Sub picTreeItem_Resize()
    Err = 0: On Error Resume Next
    With tvwMainItem
        .Top = picTreeItem.ScaleTop
        .Height = picTreeItem.ScaleHeight
        .Left = picTreeItem.ScaleLeft
        .Width = picTreeItem.ScaleWidth
    End With
End Sub

Private Sub picTreeWholeSet_Resize()
    Err = 0: On Error Resume Next
    With tvwWholeSet
        .Top = picTreeWholeSet.ScaleTop
        .Height = picTreeWholeSet.ScaleHeight
        .Left = picTreeWholeSet.ScaleLeft
        .Width = picTreeWholeSet.ScaleWidth
    End With
End Sub

Private Sub tbClassPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If mblnNotClick = True Then Exit Sub
    Select Case Val(Item.Tag)
    Case mCalssPage.pg_细目
        If Val(tbClassPage.Tag) <> mCalssPage.pg_细目 Then
            '改变了,需要重新处理数据
            '先保存上次选择的列
           ' SaveListViewState Me.lvwMain_S, Me.Name & "_" & Val(tbClassPage.Tag), lvwMain_S.View
            tbClassPage.Tag = mCalssPage.pg_细目: mstrKey = ""
            If Not tvwMainItem.SelectedItem Is Nothing Then
                 Call tvwMainItem_NodeClick(tvwMainItem.SelectedItem)
            End If
        End If
         If tvwMainItem.Enabled And tvwMainItem.Visible Then tvwMainItem.SetFocus
    Case mCalssPage.pg_成套
        If Val(tbClassPage.Tag) <> mCalssPage.pg_成套 Then
            '改变了,需要重新处理数据
            '先保存上次选择的列
            'SaveListViewState Me.lvwMain_S, Me.Name & "_" & Val(tbClassPage.Tag), lvwMain_S.View
            tbClassPage.Tag = mCalssPage.pg_成套: tvwWholeSet.Tag = ""
            If Not tvwWholeSet.SelectedItem Is Nothing Then
                 Call tvwWholeSet_NodeClick(tvwWholeSet.SelectedItem)
            End If
        End If
         If tvwWholeSet.Enabled And tvwWholeSet.Visible Then tvwWholeSet.SetFocus
    End Select
    Call SetPageVisible
    
End Sub

Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If tbClassPage.Selected Is Nothing Then Exit Sub
    If Val(tbClassPage.Selected.Tag) = mCalssPage.pg_成套 Then
        mint上次成套页 = Val(Item.Tag)
    Else
        mint上次细目页 = Val(Item.Tag)
    End If
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
On Error GoTo ErrHandle
    Dim i As Integer
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
    Next
    mnuViewIcon(ButtonMenu.Index - 1).Checked = True
    lvwMain_S.View = ButtonMenu.Index - 1
    lvwWholeSetItem_S.View = lvwMain_S.View
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Toolbar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrHandle
    If Button = 2 Then
        PopupMenu mnuViewTool, 2
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub tvwMainItem_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo ErrHandle
    If Node Is Nothing Then Exit Sub
    If mstrKey = Node.Key Then Exit Sub
    mstrKey = Node.Key
    
    FillList Format(Node.Tag)
    If lvwMain_S.ListItems.Count > 0 Then
        If Not lvwMain_S.SelectedItem Is Nothing Then
            lvwMain_S_ItemClick lvwMain_S.SelectedItem
            Exit Sub
        End If
    End If
    
    msh价目.Clear 1
    msh价目.Rows = 2
    
    chk价格.value = 0
    opt科室(0).value = 1
    opt科室(1).value = 0
    opt科室(2).value = 0
    opt科室(3).value = 0
    opt科室(4).value = 0
    opt科室(5).value = 0
    opt科室(6).value = 0
    lvwOutIn.ListItems.Clear
    msh从属.Rows = 2
    msh从属.TextMatrix(1, 0) = ""
    msh从属.TextMatrix(1, 1) = ""
    msh从属.TextMatrix(1, 2) = ""
    msh从属.TextMatrix(1, 3) = ""
    
    mshAlias.Rows = 2
    mshAlias.TextMatrix(1, 0) = ""
    mshAlias.TextMatrix(1, 1) = ""
    mshAlias.TextMatrix(1, 2) = ""
    mshAlias.TextMatrix(1, 3) = ""
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub tvwMainItem_GotFocus()
    Call MenuSet
End Sub

Private Sub tvwMainItem_LostFocus()
    Call MenuSet
End Sub

Private Sub tvwMainItem_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrHandle
    If Button = 2 Then
        If mnuShortMenu1(1).Visible = False Then Exit Sub
        mnuShortMenu1(1).Enabled = mnuEditParent.Enabled
        mnuShortMenu1(2).Enabled = mnuEditModifyAssort.Enabled
        mnuShortMenu1(3).Enabled = mnuEditDeleteAssort.Enabled
        PopupMenu mnuShort1, vbPopupMenuRightButton
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    Dim sngTop As Single, sngBottom As Single
    On Error Resume Next
    sngTop = IIF(CoolBar1.Visible, CoolBar1.Top + CoolBar1.Height, 0) + Screen.TwipsPerPixelY
    sngBottom = Me.ScaleHeight - IIF(stbThis.Visible, stbThis.Height, 0)
    'picTreeClass_S的位置
    picTreeClass_S.Top = sngTop
    picTreeClass_S.Height = IIF(sngBottom - picTreeClass_S.Top > 0, sngBottom - picTreeClass_S.Top, 0)
    picTreeClass_S.Left = 0
    'picSplit的位置
    picSplit.Top = sngTop
    picSplit.Height = picTreeClass_S.Height
    picSplit.Left = picTreeClass_S.Left + picTreeClass_S.Width
    'lvwMain_S的位置
    lvwMain_S.Top = sngTop - Screen.TwipsPerPixelY
    lvwMain_S.Left = picSplit.Left + picSplit.Width

    
    If Me.ScaleWidth - lvwMain_S.Left > 0 Then lvwMain_S.Width = Me.ScaleWidth - lvwMain_S.Left
    'picNS的位置
    picNS.Left = lvwMain_S.Left
    picNS.Top = lvwMain_S.Top + lvwMain_S.Height
    picNS.Width = lvwMain_S.Width
    'picTreeClass_S的位置
    tbPage.Left = lvwMain_S.Left
    tbPage.Top = picNS.Top + picNS.Height
    tbPage.Width = lvwMain_S.Width
    tbPage.Height = IIF(sngBottom - tbPage.Top > 0, sngBottom - tbPage.Top, 0)
    Me.pic停用原因.Left = Me.tbPage.Left + 4300
    Me.pic停用原因.Top = Me.tbPage.Top + 100
    lbl停用原因.ZOrder
    With lvwWholeSetItem_S
        .Top = lvwMain_S.Top
        .Left = lvwMain_S.Left
        .Width = lvwMain_S.Width
        .Height = lvwMain_S.Height
    End With
    lblMsg.Move opt科室(1).Left + opt科室(1).Width + 2500, opt科室(1).Top
    CoolBar1.Bands(1).Width = Me.Width - 2000
    Me.Refresh
End Sub


Private Sub mnuEditChild_Click()
'新增项目
    On Error GoTo ErrHandle
    Dim strSQL As String

    If mnuEdit.Visible = False Then Exit Sub
    If tvwMainItem.SelectedItem Is Nothing Then Exit Sub
    If lvwMain_S.ListItems.Count > 0 And Not lvwMain_S.SelectedItem Is Nothing Then
        If IsNumeric(lvwMain_S.SelectedItem.Tag) Then
            If CLng(lvwMain_S.SelectedItem.Tag) > 0 Then
                Call frmChargeItem.编辑项目(mstrPrivs, mblnCanUpdateAll, "C" & IIF(mstrClass = "", " ", mstrClass) & tvwMainItem.SelectedItem.Tag, , , , mbln启动医价系统)
                Exit Sub
            End If
        End If
                
        Call frmChargeItem.编辑项目(mstrPrivs, mblnCanUpdateAll, "C" & IIF(mstrClass = "", " ", mstrClass) & tvwMainItem.SelectedItem.Tag, , , , mbln启动医价系统)
    Else
        '此时的类别为空
        Call frmChargeItem.编辑项目(mstrPrivs, mblnCanUpdateAll, "C " & tvwMainItem.SelectedItem.Tag, , , , mbln启动医价系统)
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuEditCopy_Click()
'复制新增
On Error GoTo ErrHandle
    If lvwMain_S.SelectedItem Is Nothing Then Exit Sub
    Call frmChargeItem.编辑项目(mstrPrivs, mblnCanUpdateAll, "C" & IIF(mstrClass = "", " ", mstrClass) & IIF(tvwMainItem.SelectedItem.Key = "Root", "", tvwMainItem.SelectedItem.Tag), lvwMain_S.SelectedItem.Tag, , 5, mbln启动医价系统)  'editCopy
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuEditParent_Click()
'新增分类
On Error GoTo ErrHandle
    With frmChargeSort
        If Me.tvwMainItem.SelectedItem Is Nothing Then
            .txtParent.Tag = 0
        Else
            .txtParent.Tag = Mid(Me.tvwMainItem.SelectedItem.Key, 2)
        End If
        .Tag = "增加"
        If .ShowMe(1, Me) Then Call FillTree
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuEditModifyAssort_Click()
'修改分类
Dim i As Long
Dim strSQL As String
Dim rsTmp As New ADODB.Recordset

On Error GoTo ErrHandle
    If tvwMainItem.SelectedItem Is Nothing Then Exit Sub
    With frmChargeSort
        .mblnCancel = True
        If Me.tvwMainItem.SelectedItem.Parent Is Nothing Then
            .txtParent.Tag = 0
            .txtParent.Text = "(无)"
            .txtUpCode.Text = ""
            .txtCode.Text = Mid(Split(Me.tvwMainItem.SelectedItem.Text, "]")(0), 2)
            .txtCode.MaxLength = Len(.txtCode.Text)
            .txtCode.Tag = .txtCode.MaxLength
        Else
            .txtParent.Tag = Mid(Me.tvwMainItem.SelectedItem.Parent.Key, 2)
            .txtParent.Text = Me.tvwMainItem.SelectedItem.Parent.Text
            .txtUpCode.Text = Mid(Split(Me.tvwMainItem.SelectedItem.Parent.Text, "]")(0), 2)
            .txtCode.Text = Mid(Split(Me.tvwMainItem.SelectedItem.Text, "]")(0), Len(.txtUpCode.Text) + 2)
            .txtCode.MaxLength = Len(.txtCode.Text)
            .txtCode.Tag = .txtCode.MaxLength
        End If
        .txtName = Split(Me.tvwMainItem.SelectedItem.Text, "]")(1)
        strSQL = "Select 简码 from 收费分类目录 where id=[1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Me.tvwMainItem.SelectedItem.Tag))
        
        If rsTmp.RecordCount > 0 Then
            .txtSymbol = Nvl(rsTmp!简码)
        Else
            .txtSymbol = ""
        End If
        .Tag = Mid(Me.tvwMainItem.SelectedItem.Key, 2)
        .mblnCancel = False
        If .ShowMe(1, Me) Then Call FillTree
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuEditModify_Click()
'修改项目
On Error GoTo ErrHandle
    If mnuEdit.Visible = False Then Exit Sub
    If lvwMain_S.SelectedItem Is Nothing Then Exit Sub
    Call frmChargeItem.编辑项目(mstrPrivs, mblnCanUpdateAll, "C" & IIF(mstrClass = "", " ", mstrClass) & tvwMainItem.SelectedItem.Tag, lvwMain_S.SelectedItem.Tag, , 1, mbln启动医价系统) 'EditMode.editModify
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuPriceRaise_Click()
'调价
On Error GoTo ErrHandle
    If mnuPrice.Visible = False Then Exit Sub
    ModifyMode 2    'editRaise
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ModifyMode(ByVal edit方式 As EditMode)
On Error GoTo ErrHandle
    If ActiveControl Is tvwMainItem And edit方式 < 2 Then  'editRaise
        With tvwMainItem.SelectedItem
            Call frmChargeItem.编辑项目(mstrPrivs, mblnCanUpdateAll, "C" & IIF(mstrClass = "", " ", mstrClass) & .Tag, .Tag, 0, edit方式, mbln启动医价系统)
        End With
    Else
        If lvwMain_S.SelectedItem Is Nothing Then Exit Sub
        If checkNotPrice(lvwMain_S.SelectedItem.Tag) = False Then
            MsgBox "该收费细目还存在未审核的价格，请审核后再来调价！", vbInformation, gstrSysName
            Exit Sub
        End If
        Call frmChargeItem.编辑项目(mstrPrivs, mblnCanUpdateAll, "C" & IIF(mstrClass = "", " ", mstrClass) & IIF(tvwMainItem.SelectedItem.Key = "Root", "", tvwMainItem.SelectedItem.Tag), lvwMain_S.SelectedItem.Tag, , edit方式, mbln启动医价系统)
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function checkNotPrice(ByVal lng收费细目ID As Long) As Boolean
    '检查是否还存在未生效的价格
    Dim rsData As ADODB.Recordset
    Dim strWhere As String
    
    On Error GoTo ErrHandle
    If mblnCanUpdateAll = False Then
        strWhere = " And (b.站点=[2]" & vbNewLine & _
                "       Or b.站点 Is Null And a.价格等级 In(" & vbNewLine & _
                "           Select m.名称" & vbNewLine & _
                "           From 收费价格等级 M, 收费价格等级应用 N" & vbNewLine & _
                "           Where m.名称 = n.价格等级 And Nvl(m.是否适用普通项目, 0) = 1 And n.站点 = [2]" & vbNewLine & _
                "                 And (m.撤档时间 Is Null Or m.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd'))))"
    End If
    If lng收费细目ID <> 0 Then
        gstrSQL = "Select 1 From 收费调价记录 A,收费项目目录 B Where a.收费细目ID = b.ID And a.审核标志 = 0 And a.收费细目id=[1]" & strWhere & " And Rownum < 2"
    Else
        gstrSQL = "Select 1 From 收费调价记录 A,收费项目目录 B" & _
                " Where a.收费细目ID = b.ID And a.审核标志 = 0" & strWhere & " And Rownum < 2"
    End If
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "未生效单据查询", lng收费细目ID, gstrNodeNo)
    If rsData.RecordCount > 0 Then
        checkNotPrice = False
    Else
        checkNotPrice = True
    End If
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub mnuEditDeleteAssort_Click()
'删除
    Dim strKey As String
    Dim intIndex As Long
    
    On Error GoTo ErrHandle
    If tvwMainItem.SelectedItem Is Nothing Then Exit Sub
    
    If MsgBox("你确认要删除名称为“" & tvwMainItem.SelectedItem.Text & "”的分类吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
        Me.MousePointer = 11
        gstrSQL = "ZL_收费分类目录_DELETE(" & tvwMainItem.SelectedItem.Tag & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Me.MousePointer = 0
        
        strKey = tvwMainItem.SelectedItem.Key
        If Not tvwMainItem.SelectedItem.Next Is Nothing Then
            tvwMainItem.SelectedItem.Next.Selected = True
            tvwMainItem_NodeClick tvwMainItem.SelectedItem
        Else
            If Not tvwMainItem.SelectedItem.Parent Is Nothing Then
                tvwMainItem.SelectedItem.Parent.Selected = True
            End If
            If Not tvwMainItem.SelectedItem Is Nothing Then
                tvwMainItem_NodeClick tvwMainItem.SelectedItem
            End If
        End If
        tvwMainItem.Nodes.Remove strKey
    End If
    MenuSet
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Me.MousePointer = 0
End Sub

Private Sub mnuEditDelete_Click()
'删除
    Dim strKey As String
    Dim intIndex As Long
    
    On Error GoTo ErrHandle
    If lvwMain_S.SelectedItem Is Nothing Then Exit Sub
    If MsgBox("你确认要删除名称为“" & lvwMain_S.SelectedItem.Text & "”的项目吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
        Me.MousePointer = 11
        gstrSQL = "zl_收费细目_delete(" & lvwMain_S.SelectedItem.Tag & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Me.MousePointer = 0
        
        With lvwMain_S
            intIndex = .SelectedItem.Index
            .ListItems.Remove .SelectedItem.Key
            If .ListItems.Count > 0 Then
                intIndex = IIF(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
                .ListItems(intIndex).Selected = True
                .ListItems(intIndex).EnsureVisible
                FillItem .ListItems(intIndex).Key
            Else
                FillItem ""
            End If
        End With
    End If
    MenuSet
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Me.MousePointer = 0
End Sub

Private Sub mnuEditStart_Click()
    Dim str原因 As String
    On Error GoTo ErrHandle
    
    If lvwMain_S.SelectedItem Is Nothing Then Exit Sub
    
    mfrmEarnRS.ShowMe 1, str原因
    
    If str原因 = "" Then Exit Sub
    
    gstrSQL = "zl_收费细目_reuse(" & lvwMain_S.SelectedItem.Tag & ",'" & str原因 & "')"
    '执行启用过程
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    '改变图标和颜色
    With lvwMain_S.SelectedItem
        .Icon = "Item"
        .SmallIcon = "Item"
        .ForeColor = RGB(0, 0, 0)
        
        Dim i As Integer
        For i = 1 To lvwMain_S.ColumnHeaders.Count
            If i < lvwMain_S.ColumnHeaders.Count Then
                .ListSubItems(i).ForeColor = RGB(0, 0, 0)
            End If
            '更新撤档时间
            If lvwMain_S.ColumnHeaders(i).Text = "撤档时间" Then
                .SubItems(i - 1) = "3000-01-01"
            End If
        Next
    End With
    '改变状态栏和菜单
    MenuSet
    lvwMain_S_ItemClick lvwMain_S.SelectedItem
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuEditStop_Click()
    On Error GoTo ErrHandle
    Dim strKey As String
    Dim intIndex As Integer
    Dim strTmp As String
    Dim str原因 As String
    
    If lvwMain_S.SelectedItem Is Nothing Then Exit Sub
    
    strKey = lvwMain_S.SelectedItem.Tag
    
    If Not Check收费项目(Val(strKey), strTmp) Then
        Exit Sub
    End If
    If strTmp <> "" Then
        If MsgBox("该项目还存在以下依赖关系：" & vbCrLf & strTmp & "是否停用？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    
    mfrmEarnRS.ShowMe 2, str原因
    
    If str原因 = "" Then Exit Sub
    
    gstrSQL = "zl_收费细目_stop(" & strKey & ",'" & str原因 & "')"
    '执行启用过程
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    '改变图标和颜色
    If mnuViewShowStop.Checked = True Then '要显示停用部门
        With lvwMain_S.SelectedItem
            .Icon = "ItemNo"
            .SmallIcon = "ItemNo"
            .ForeColor = RGB(255, 0, 0)
            
            Dim i As Integer
            For i = 1 To lvwMain_S.ColumnHeaders.Count
                If i < lvwMain_S.ColumnHeaders.Count Then
                    .ListSubItems(i).ForeColor = RGB(255, 0, 0)
                End If
                '更新撤档时间
                If lvwMain_S.ColumnHeaders(i).Text = "撤档时间" Then
                    .SubItems(i - 1) = Format(Date, "yyyy-MM-dd")
                End If
            Next
        End With
    Else '不显示停用部门
        With lvwMain_S
            intIndex = .SelectedItem.Index
            .ListItems.Remove .SelectedItem.Key
            If .ListItems.Count > 0 Then
                intIndex = IIF(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
                .ListItems(intIndex).Selected = True
                .ListItems(intIndex).EnsureVisible
                FillItem .ListItems(intIndex).Key
            Else
                FillItem ""
            End If
        End With
    End If
    MenuSet
    lvwMain_S_ItemClick lvwMain_S.SelectedItem
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetCodeLength(ByVal strTable As String) As Long
'功能:从表中得到字段的长度
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo ErrHandle
    
    GetCodeLength = 0
    gstrSQL = "select nvl(max(Vsize(编码)),0) as LenCode from " & strTable
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rsTmp.RecordCount > 0 Then
        GetCodeLength = rsTmp!lencode
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub picsplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        msngStartX = X
    End If
End Sub

Private Sub picsplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Dim sngTemp As Single
    If Button = 1 Then
        sngTemp = picSplit.Left + X - msngStartX
        If sngTemp > 500 And Me.ScaleWidth - (sngTemp + picSplit.Width) > 600 Then
            picSplit.Left = sngTemp
            picTreeClass_S.Width = picSplit.Left - picTreeClass_S.Left
            Form_Resize
        End If
    End If
End Sub

Private Sub picNS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        msngStartY = Y
    End If
End Sub

Private Sub picNS_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Dim sngTemp As Single
    If Button = 1 Then
        sngTemp = picNS.Top + Y - msngStartY
        If sngTemp > lvwMain_S.Top + 600 And Me.ScaleHeight - (sngTemp + picNS.Height) > 600 Then
            picNS.Top = sngTemp
            lvwMain_S.Height = picNS.Top - lvwMain_S.Top
            lvwWholeSetItem_S.Height = lvwMain_S.Height
            Form_Resize
        End If
    End If
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnufilepre_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub mnufileset_Click()
    zlPrintSet
End Sub

'Private Sub tbPage_Click()
'    Dim i As Integer
'    tbPage.ZOrder 1
'    For i = 1 To 5
'        fra(i).ZOrder 1
'    Next
'    fra(tbPage.SelectedItem.Index).ZOrder
'End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Child"
            If tbClassPage.Selected Is Nothing Then Exit Sub
            If Val(tbClassPage.Selected.Tag) = mCalssPage.pg_成套 Then
                '成套项目增加
                Call mnuEditWholeSetItemAdd_Click
            Else
                 mnuEditChild_Click
            End If
        Case "Parent"
            If tbClassPage.Selected Is Nothing Then Exit Sub
            If Val(tbClassPage.Selected.Tag) = mCalssPage.pg_成套 Then
                '成套分类
                Call mnuEditWholeSetClassAdd_Click
            Else
                mnuEditParent_Click
            End If
        Case "Modify"
            If tbClassPage.Selected Is Nothing Then Exit Sub
            If Val(tbClassPage.Selected.Tag) = mCalssPage.pg_成套 Then
                '成套
                If ActiveControl Is tvwWholeSet Then
                    mnuEditWholeSetClassModify_Click
                Else
                    mnuEditWholeSetItemModify_Click
                End If
            Else
                If ActiveControl Is tvwMainItem Then
                    If InStr(mstrPrivs, "类别管理") > 0 Then
                        mnuEditModifyAssort_Click
                    End If
                Else
                    mnuEditModify_Click
                End If
            End If
        Case "Delete"
            If tbClassPage.Selected Is Nothing Then Exit Sub
            If Val(tbClassPage.Selected.Tag) = mCalssPage.pg_成套 Then
                '成套
                If ActiveControl Is tvwWholeSet Then
                    mnuEditWholeSetClassDelete_Click
                Else
                    mnuEditWholeSetItemDelete_Click
                End If
            Else
                If ActiveControl Is tvwMainItem Then
                    If InStr(mstrPrivs, "类别管理") > 0 Then
                        mnuEditDeleteAssort_Click
                    End If
                Else
                    mnuEditDelete_Click
                End If
            End If
        Case "Raise"
            mnuPriceRaise_Click
        Case "RaiseVerify"
            frmChargePriceVerify.ShowMe Me, mblnCanUpdateAll
        Case "Start"
            mnuEditStart_Click
        Case "Stop"
            mnuEditStop_Click
        Case "Quit"
            mnuFileExit_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Find"
            mnuViewFind_Click
        Case "Preview"
            mnufilepre_Click
        Case "Help"
            mnuhelptopic_Click
        Case "View"
            mnuViewIcon(lvwMain_S.View).Checked = False
            If lvwMain_S.View = 3 Then
                mnuViewIcon(0).Checked = True
                lvwMain_S.View = 0
            Else
                mnuViewIcon(lvwMain_S.View + 1).Checked = True
                lvwMain_S.View = lvwMain_S.View + 1
            End If
    End Select
    
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    CoolBar1.Visible = mnuViewToolButton.Checked
    CoolBar1.Bands("only").MinHeight = Toolbar1.Height
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim buttTemp As Button
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For Each buttTemp In Toolbar1.Buttons
        If mnuViewToolText.Checked Then
            buttTemp.Caption = buttTemp.Tag
        Else
            buttTemp.Caption = ""
        End If
    Next
    CoolBar1.Bands("only").MinHeight = Toolbar1.Height
    Form_Resize
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Form_Resize
End Sub

Private Sub mnuhelptopic_Click()
      ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hwnd)
End Sub
Private Sub subPrint(bytMode As Byte)
    '功能:进行打印,预览和输出到EXCEL
    '参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    Dim objPrint As New zlPrintLvw
    If gstrUserName = "" Then Call GetUserInfo
    If tbClassPage.Selected Is Nothing Then Exit Sub
    If Val(tbClassPage.Selected.Tag) = mCalssPage.pg_成套 Then
        If tvwWholeSet.SelectedItem Is Nothing Then Exit Sub
        If lvwWholeSetItem_S.ListItems.Count = 0 Then Exit Sub
        objPrint.Title.Text = "成套收费项目"
        Set objPrint.Body.objData = lvwWholeSetItem_S
        objPrint.UnderAppItems.Add "分类：" & tvwWholeSet.SelectedItem.Text
    Else
        If tvwMainItem.SelectedItem Is Nothing Then Exit Sub
        If lvwMain_S.ListItems.Count = 0 Then Exit Sub
        objPrint.Title.Text = "收费项目"
        Set objPrint.Body.objData = lvwMain_S
        objPrint.UnderAppItems.Add "分类：" & tvwMainItem.SelectedItem.Text
    End If
    objPrint.BelowAppItems.Add "打印人：" & gstrUserName
    objPrint.BelowAppItems.Add "打印时间：" & Format(sys.Currentdate, "yyyy年MM月dd日")
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrViewLvw objPrint, 1
          Case 2
              zlPrintOrViewLvw objPrint, 2
          Case 3
              zlPrintOrViewLvw objPrint, 3
      End Select
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
End Sub
Public Sub FillWholeSetTree()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:填充成套分类数据
    '编制:刘兴洪
    '日期:2010-08-24 14:55:07
    '说明:27327
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, objNode As Node
    Dim strPreKey As String
    Err = 0: On Error GoTo ErrHand:
    strSQL = "" & _
    "   Select id,上级ID,编码,名称 " & _
    "   From 成套项目分类  " & _
    "   Start with 上级id is null Connect by Prior   Id=上级ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With tvwWholeSet
        If Not .SelectedItem Is Nothing Then strPreKey = .SelectedItem.Key
        .Nodes.Clear
       Set objNode = .Nodes.Add(, , "Root", "所有成套", "RootS", "Exp")
       objNode.Expanded = True
       objNode.Sorted = True
       Do While Not rsTemp.EOF
            If IsNull(rsTemp!上级id) Then
                Set objNode = .Nodes.Add("Root", tvwChild, "K" & Nvl(rsTemp!ID), Nvl(rsTemp!编码) & "-" & Nvl(rsTemp!名称), "RootS", "Exp")
            Else
                Set objNode = .Nodes.Add("K" & rsTemp!上级id, tvwChild, "K" & Nvl(rsTemp!ID), Nvl(rsTemp!编码) & "-" & Nvl(rsTemp!名称), "RootS", "Exp")
            End If
            objNode.Sorted = True
            If objNode.Key = strPreKey Then
                objNode.EnsureVisible
                objNode.Selected = True
                objNode.Expanded = True
            End If
            objNode.Sorted = True
            rsTemp.MoveNext
       Loop
       tvwWholeSet.Tag = ""
       If .SelectedItem Is Nothing Then .Nodes("Root").Selected = True
       If Not .SelectedItem Is Nothing Then
            
            Call tvwWholeSet_NodeClick(.SelectedItem)
       End If
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Sub
Private Function FillWholeItem(ByVal lng分类id As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:填充成套项目
    '入参:lng分类id-分类ID,0-所有分类
    '出参:
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-08-25 15:41:48
    '问题:27327
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, strWhere As String, strOwner As String
    Dim strPreKey As String, objListItem As ListItem, lngCol As Long
    On Error GoTo ErrHandle
    
    Screen.MousePointer = vbHourglass
    Me.stbThis.Panels(2).Text = "正在读取收费成套项目列表数据,请稍候 ．．．"
    Me.stbThis.Refresh
    
    If Not tbClassPage.Selected Is Nothing Then
        If Not lvwWholeSetItem_S.SelectedItem Is Nothing And Val(tbClassPage.Selected.Tag) = mCalssPage.pg_成套 Then
            strPreKey = lvwWholeSetItem_S.SelectedItem.Key
        End If
    End If
    
    strSQL = "select 所有者 from zlsystems where 编号=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "读所有者", glngSys)
    If rsTemp.RecordCount = 1 Then
        strOwner = IIF(IsNull(rsTemp!所有者), "", rsTemp!所有者)
    End If
    rsTemp.Close
    
    If strOwner <> gstrDbUser Then
        strWhere = " And ( A.人员ID=[2] "
        If InStr(1, mstrPrivs, ";本科成套方案;") > 0 Then
            strWhere = strWhere & " OR Exists(Select 1 From 成套项目使用科室 A1 ,部门人员 B1 Where A1.成套ID=A.ID And A1.科室ID=B1.部门Id and B1.人员id=[2]) "
        End If
        If InStr(1, mstrPrivs, ";全院成套方案;") > 0 Then
            strWhere = strWhere & " OR nvl(A.范围,0)=0 "
        End If
        strWhere = strWhere & ")"
    End If
    
    strSQL = "" & _
    "   Select  A.Id,A.分类ID,A.编码,A.名称,A.拼音,A.五笔,decode(nvl(范围,0),0,'全院',1,'指定科室',decode(A.人员id,Null,'指定操作员',B.姓名)) As 使用范围," & _
    "              C.名称 as 所属分类 " & _
    "   From 成套收费项目 A,人员表 B " & _
            IIF(lng分类id = 0, ",成套项目分类 C", " ,(Select ID,上级ID,编码,名称 From  成套项目分类  Start With Id =[1] Connect By Prior Id=上级id ) C") & _
    "   Where a.人员id=b.Id(+) And A.分类id=C.ID " & strWhere
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng分类id, glngUserId)
    zlControl.FormLock lvwWholeSetItem_S.hwnd
    mblnNotClick = True
    With lvwWholeSetItem_S
        .ListItems.Clear
        Do While Not rsTemp.EOF
            '添加节点
            Set objListItem = .ListItems.Add(, "K" & rsTemp!ID, Nvl(rsTemp!编码) & "-" & Nvl(rsTemp!名称), "Item", "Item")
            objListItem.Tag = Nvl(rsTemp!分类id)
            ' "名称,1500,0,1;编码,800,0,2;简码,1400,0,0;使用范围,400,0,0;所属分类,2400,0,0"
            '根据ListView的列名从数据库取数
            For lngCol = 2 To lvwWholeSetItem_S.ColumnHeaders.Count
                objListItem.SubItems(lngCol - 1) = Nvl(rsTemp.Fields(lvwWholeSetItem_S.ColumnHeaders(lngCol).Text))
            Next
            If rsTemp.AbsolutePosition = 1 Then '缺省为第一行选中
                objListItem.Selected = True
            End If
            If objListItem.Key = strPreKey Then
                objListItem.Selected = True
                objListItem.EnsureVisible
            End If
            rsTemp.MoveNext
        Loop
        If .ListItems.Count > 0 Then
                Me.stbThis.Panels(2).Text = "收费成套项目数据读取完成！"
        Else
                Me.stbThis.Panels(2).Text = ""
        End If
    End With
    mblnNotClick = False
    lvwWholeSetItem_S.Tag = ""
    
    If Not lvwWholeSetItem_S.SelectedItem Is Nothing Then
        Call lvwWholeSetItem_S_ItemClick(lvwWholeSetItem_S.SelectedItem)
    Else
        '清除成套项目的一些数据
        Call zlClearDownWholeSetItem
    End If
    zlControl.FormLock 0
    Screen.MousePointer = vbDefault
    FillWholeItem = True
    Exit Function
ErrHandle:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Screen.MousePointer = vbHourglass
        Resume
    End If
    Me.stbThis.Panels(2).Text = ""
    Me.stbThis.Refresh
    mblnNotClick = False
    zlControl.FormLock 0
End Function
Private Function FillWholeSetItemChildData(ByVal lng成套ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载成套项目子数据
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-08-25 17:42:49
    '说明:27327
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim objListItem As ListItem, lng父号 As Long, j As Long, i As Long
    Dim strWherePriceGrade As String
    
    On Error GoTo ErrHandle

    'strSQL = "" & _
    "   Select '' as 标志,A.序号, A.成套id, A.收费细目id, B.编码, B.名称, B.计算单位, B.规格, A.从属父号, A.数量, A.单价, A.执行科室id, " & _
    "          decode(C.编码,NULL,'',C.编码||'-') ||C.名称 As 执行科室 " & _
    "   From 成套收费项目组合 A, 收费项目目录 B, 部门表 C " & _
    "   Where A.收费细目id = B.ID And A.执行科室id = C.ID(+)  And A.成套id = [1]" & _
    "   Order By A.序号"

    If gstr普通价格等级 = "" And gstr药品价格等级 = "" And gstr卫材价格等级 = "" Then
        strWherePriceGrade = " And j.价格等级 Is Null"
    Else
        strWherePriceGrade = "" & _
            " And ((Instr(';5;6;7;', ';' || k.类别 || ';') > 0 And j.价格等级 = [2])" & vbNewLine & _
            "      Or (Instr(';4;', ';' || k.类别 || ';') > 0 And j.价格等级 = [3])" & vbNewLine & _
            "      Or (Instr(';4;5;6;7;', ';' || k.类别 || ';') = 0 And j.价格等级 = [4])" & vbNewLine & _
            "      Or (j.价格等级 Is Null" & vbNewLine & _
            "          And Not Exists (Select 1" & vbNewLine & _
            "                          From 收费价目" & vbNewLine & _
            "                          Where j.收费细目id = 收费细目id And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            "                                And ((Instr(';5;6;7;', ';' || k.类别 || ';') > 0 And 价格等级 = [2])" & vbNewLine & _
            "                                      Or (Instr(';4;', ';' || k.类别 || ';') > 0 And 价格等级 = [3])" & vbNewLine & _
            "                                      Or (Instr(';4;5;6;7;', ';' || k.类别 || ';') = 0 And 价格等级 = [4])))))"
    End If
    
    gstrSQL = "" & _
    "   Select  /*+Rule */ A.成套ID,A.收费细目ID,A.序号,A.从属父号,A.付数,A.数量,A.单价,A.执行科室ID, " & _
    "              B.类别,B.编码,B.名称,B.计算单位,B.规格,C.中药形态,D.编码 as 诊疗编码, " & _
    "              D.名称 as 诊疗名称,D.计算单位 as 剂量单位,C.剂量系数, " & _
    "              E.编码 As 执行科室编码,E.名称 As 执行科室名称, " & _
    "              M.编码 As 成套编码,M.名称 As 成套名称,M.拼音,M.五笔,M.备注,M.范围, " & _
    "              M.分类ID,M.人员ID,G.姓名,J.编码 As 分类编码,J.名称 As 分类名称 ,B.是否变价,B.执行科室,C.药名ID," & _
    "              Decode(B.是否变价,1,'时价',LTrim(To_Char(J1.现价,'999999999.9999999'))) as 现价  " & _
    "   From 成套收费项目 M,成套项目分类 J,成套收费项目组合 A,收费项目目录 B,药品规格 C,诊疗项目目录 D, " & _
    "             部门表 E,人员表 G," & _
    "             (Select j.收费细目id, Sum(j.现价) as 现价" & vbNewLine & _
    "              From 收费价目 J,收费项目目录 K" & vbNewLine & _
    "              Where j.收费细目ID = k.ID And Sysdate Between J.执行日期 And Nvl(J.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
                         strWherePriceGrade & vbNewLine & _
    "              Group By j.收费细目id ) J1 " & _
    "   Where   M.分类ID=J.Id And  M.人员ID=G.Id(+) And M.Id=A.成套ID(+)  " & _
    "               And A.收费细目id=b.Id(+)  And a.收费细目ID=C.药品ID(+) And C.药名ID=D.Id(+) " & _
    "               And A.收费细目id=J1.收费细目ID(+)  And A.执行科室ID=E.Id(+)  " & _
    "               And M.ID=[1] Order by A.序号"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng成套ID, gstr药品价格等级, gstr卫材价格等级, gstr普通价格等级)
    With vsWholeSet
        .redraw = flexRDNone
        .Clear 1
        .OutlineBar = flexOutlineBarSimple
        .OutlineCol = .ColIndex("标志"): .SubtotalPosition = flexSTAbove
        .Rows = IIF(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        i = 1
        lng父号 = 0
        Do While Not rsTemp.EOF
            If Val(Nvl(rsTemp!从属父号)) = 0 Then
                    lng父号 = Nvl(rsTemp!序号)
            End If
            .TextMatrix(i, .ColIndex("序号")) = Nvl(rsTemp!序号)
            .TextMatrix(i, .ColIndex("从属父号")) = Nvl(rsTemp!从属父号)
            .TextMatrix(i, .ColIndex("收费项目")) = Nvl(rsTemp!编码) & "-" & Nvl(rsTemp!名称)
            .TextMatrix(i, .ColIndex("规格")) = Nvl(rsTemp!规格)
            .TextMatrix(i, .ColIndex("缺省付数")) = IIF(Val(Nvl(rsTemp!付数)) = 0, 1, Val(Nvl(rsTemp!付数)))
            .TextMatrix(i, .ColIndex("缺省数量")) = FormatEx(Val(Nvl(rsTemp!数量)), 5)
            .TextMatrix(i, .ColIndex("缺省价格")) = FormatEx(Val(Nvl(rsTemp!单价)), 8)
            .TextMatrix(i, .ColIndex("缺省执行科室")) = Nvl(rsTemp!执行科室名称)
            .TextMatrix(i, .ColIndex("单位")) = Nvl(rsTemp!计算单位)
            .TextMatrix(i, .ColIndex("现价")) = IIF(Nvl(rsTemp!现价) = "实价", "实价", FormatEx(Val(Nvl(rsTemp!现价)), 5))
            If Nvl(rsTemp!类别) = "7" Then
                '草药,显示诊疗名称
                .TextMatrix(.Row, .ColIndex("药名")) = Nvl(rsTemp!诊疗编码) & "-" & Nvl(rsTemp!诊疗名称)
                .TextMatrix(.Row, .ColIndex("单位")) = Nvl(rsTemp!剂量单位)
                .TextMatrix(.Row, .ColIndex("缺省数量")) = FormatEx(Val(.TextMatrix(.Row, .ColIndex("缺省数量"))) * Val(Nvl(rsTemp!剂量系数)), 5)
                .TextMatrix(.Row, .ColIndex("缺省价格")) = FormatEx(Val(.TextMatrix(.Row, .ColIndex("缺省价格"))) / Val(Nvl(rsTemp!剂量系数)), 8)
              '  .TextMatrix(.Row, .ColIndex("中药形态")) = Val(Nvl(rsTemp!中药形态))
            End If
        
            If Val(Nvl(rsTemp!从属父号)) = 0 Then
                    .IsSubtotal(i) = True
                    .RowOutlineLevel(i) = 1
            ElseIf lng父号 = Val(.TextMatrix(i, .ColIndex("从属父号"))) Then
                    If i > 2 Then
                        If Val(.TextMatrix(i - 1, .ColIndex("从属父号"))) <> 0 Then
                            .IsSubtotal(i - 1) = False
                            .RowOutlineLevel(i - 1) = 1
                        End If
                    End If
                    .IsSubtotal(i) = True
                    .RowOutlineLevel(i) = 2
            End If
            i = i + 1
            rsTemp.MoveNext
        Loop
        
        zl_vsGrid_Para_Restore mlngMode, vsWholeSet, Me.Caption, "成套项目组成表列-主界面", True, True
        .redraw = flexRDBuffered
    End With
    strSQL = "" & _
    "   Select A.科室ID,B.编码,b.名称 " & _
    "   From 成套项目使用科室 A,部门表  B  " & _
    "   Where a.科室id=b.Id And a.成套ID=[1]" & _
    "   Order By 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng成套ID)
   With lvwUseDept
        .ListItems.Clear
        Do While Not rsTemp.EOF
            .ListItems.Add , "K" & rsTemp!科室ID, Nvl(rsTemp!编码) & "-" & Nvl(rsTemp!名称), "Dept", "Dept"
            rsTemp.MoveNext
        Loop
   End With
    FillWholeSetItemChildData = True
    Exit Function
ErrHandle:
    vsWholeSet.redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
        vsWholeSet.redraw = flexRDNone
    End If
End Function
Private Sub zlClearDownWholeSetItem()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除指定成套项目的组成和使用科室数据
    '编制:刘兴洪
    '日期:2010-08-25 16:35:03
    '问题:27327
    '---------------------------------------------------------------------------------------------------------------------------------------------

    With vsWholeSet
        .Rows = 2
        .Clear 1
    End With
    With lvwUseDept
        .ListItems.Clear
    End With
End Sub

Public Sub FillTree()
'功能:装入收费类别和收费细目的所有分类到tvwMainItem
    '本程序中树节点比其它程序的KEY值多一个字符，即第二位的类别编码

    Dim rs分类 As New ADODB.Recordset
    Dim strTemp As String
    Dim strKey As String
    Dim i As Long
    Dim objNode As Node
    
    mstrKey = ""     '全面刷新时就相当于用户没点过任何节点
    
    If Not tvwMainItem.SelectedItem Is Nothing Then
    '记录以前的节点
        strKey = tvwMainItem.SelectedItem.Key
    End If
    
    Screen.MousePointer = vbHourglass
    
    Me.stbThis.Panels(2).Text = "正在读取分类数据,请稍候 ．．．"
    Me.stbThis.Refresh
    On Error GoTo ErrHandle
    zlControl.FormLock tvwMainItem.hwnd
    
    tvwMainItem.Nodes.Clear
    tvwMainItem.Sorted = False
    
    '显示分类
    gstrSQL = _
        "Select ID,上级id,编码,名称,简码 " & vbCrLf & _
        "From 收费分类目录" & vbCrLf & _
        "Start With 上级id Is Null" & vbCrLf & _
        "Connect By Prior id=上级id "
        
    rs分类.CursorLocation = adUseClient
    Call zlDatabase.OpenRecordset(rs分类, gstrSQL, Me.Caption)
    If rs分类.RecordCount > 0 Then
        With rs分类
            .MoveFirst
            For i = 0 To .RecordCount - 1
                If IsNull(rs分类!上级id) Then
                    Set objNode = tvwMainItem.Nodes.Add(, , "R" & rs分类!ID, "[" & rs分类("编码") & "]" & rs分类("名称"), "RootS", "Exp")
                Else
                    Set objNode = tvwMainItem.Nodes.Add("R" & rs分类!上级id, tvwChild, "R" & rs分类!ID, "[" & rs分类("编码") & "]" & rs分类("名称"), "RootS", "Exp")
                End If
                'objNode.ExpandedImage = "Exp"
                objNode.Tag = rs分类!ID
                objNode.Sorted = True
                .MoveNext
            Next
        End With
        Me.stbThis.Panels(2).Text = ""
    Else
        Me.stbThis.Panels(2).Text = "无任何分类!"
    End If
    tvwMainItem.Sorted = True
    
    zlControl.FormLock 0
    
    Dim nod As Node
    On Error Resume Next
    Set nod = tvwMainItem.Nodes(strKey)
    If Err <> 0 Then
        Set nod = tvwMainItem.Nodes(1)
        nod.Selected = True
        nod.Expanded = True
        tvwMainItem_NodeClick nod
    Else
        Err.Clear
        nod.Selected = True
        nod.Expanded = True
        nod.EnsureVisible
        tvwMainItem_NodeClick nod
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    zlControl.FormLock 0
End Sub

Public Sub FillList(ByVal str分类 As String)
'功能:装入对应分类的项目到lvwMain_S
'参数:str分类 分类的标识
    Dim rs项目 As New ADODB.Recordset
    Dim strTemp As String
    Dim i As Integer
    Dim lst As ListItem
    Dim strKey As String
    Dim str名称 As String
    Dim j As Long
    
    On Error GoTo errHandleList
    
    If Not lvwMain_S.SelectedItem Is Nothing Then
        '保留原有键值
        strKey = lvwMain_S.SelectedItem.Key
    End If
    rs项目.CursorLocation = adUseClient
    rs项目.CursorType = adOpenKeyset
    rs项目.LockType = adLockReadOnly
    
    Screen.MousePointer = vbHourglass
    Me.stbThis.Panels(2).Text = "正在读取收费项目列表数据,请稍候 ．．．"
    Me.stbThis.Refresh
    If mnuViewShowStop.Checked = False Then
        strTemp = " and (A.撤档时间 = to_date('3000-01-01','YYYY-MM-DD') or A.撤档时间 is null) "
    End If
    If mblnCanUpdateAll = False Then
        strTemp = strTemp & " And (a.站点 Is Null Or a.站点 = [2])"
    End If
    If mnuViewShowAll.Checked = True Then
            gstrSQL = _
                "Select A.ID,A.分类ID,A.类别,C.名称 所属类别,C.固定 类别固定,A.编码,A.标识主码,A.标识子码,A.最高限价,A.最低限价,A.备选码,A.名称,A.规格,A.计算单位,A.费用类型, " & vbCrLf & _
                "       Decode(A.服务对象,1,'门诊',2,'住院',3,'门诊与住院','无') as 服务对象," & vbCrLf & _
                "       decode(A.补充摘要,1,'√','') as 补充摘要, A.说明,decode(A.屏蔽费别,1,'√','') as 屏蔽费别," & vbCrLf & _
                "       decode(A.是否变价,1,'√','') as 是否变价,decode(A.加班加价,1,'√','') as 加班加价,decode(A.执行科室,1,4,2,1,3,3,0) AS 执行科室, " & vbCrLf & _
                "       to_char(A.建档时间,'yyyy-mm-dd') as 建档时间,to_char(A.撤档时间,'yyyy-mm-dd') as 撤档时间," & vbCrLf & _
                "       decode(A.类别,'1',decode(A.项目特性,1,'急诊项目','挂号项目'),'H',decode(A.项目特性,1,'护理等级',2,'基本护理','')) As 项目特性," & vbCrLf & _
                "        Nvl(B.名称,'') As 所属分类,a.病案费目,d.编号 As 站点编号, d.名称 As 站点名称" & vbCrLf & _
                " From " & vbCrLf & _
                "   (Select Id,名称 From 收费分类目录" & vbCrLf & _
                "   Start With 上级id  = [1]" & vbCrLf & _
                "    Connect By Prior id=上级id) B," & vbCrLf & _
                "    收费项目目录 A,收费项目类别 C, zlnodelist D " & vbCrLf & _
                "Where  A.类别=C.编码 And a.站点 = d.编号(+) and A.分类id=B.Id And A.类别<>'5'  And  A.类别<>'6'  And  A.类别<>'7'" & strTemp & vbCrLf & _
                "Union" & vbCrLf & _
                "Select A.ID,A.分类ID,A.类别,C.名称 所属类别,C.固定 类别固定,A.编码,A.标识主码,A.标识子码,A.最高限价,A.最低限价,A.备选码,A.名称,A.规格,A.计算单位,A.费用类型, " & vbCrLf & _
                "       Decode(A.服务对象,1,'门诊',2,'住院',3,'门诊与住院','无') as 服务对象," & vbCrLf & _
                "       decode(A.补充摘要,1,'√','') as 补充摘要, A.说明,decode(A.屏蔽费别,1,'√','') as 屏蔽费别," & vbCrLf & _
                "       decode(A.是否变价,1,'√','') as 是否变价,decode(A.加班加价,1,'√','') as 加班加价,decode(A.执行科室,1,4,2,1,3,3,0) AS 执行科室, " & vbCrLf & _
                "       to_char(A.建档时间,'yyyy-mm-dd') as 建档时间,to_char(A.撤档时间,'yyyy-mm-dd') as 撤档时间," & vbCrLf & _
                "       decode(A.类别,'1',decode(A.项目特性,1,'急诊项目','挂号项目'),'H',decode(A.项目特性,1,'护理等级',2,'基本护理','')) As 项目特性," & vbCrLf & _
                "        Nvl(B.名称,'') As 所属分类,a.病案费目,d.编号 As 站点编号, d.名称 As 站点名称" & vbCrLf & _
                " From 收费项目目录 A,收费分类目录 B,收费项目类别 C, zlnodelist D " & vbCrLf & _
                "Where A.类别=C.编码 And a.站点 = d.编号(+) and A.分类id  = [1] And A.分类id=B.ID And A.类别<>'5'  And  A.类别<>'6'  And  A.类别<>'7'" & strTemp
    Else
            gstrSQL = _
                "Select A.ID,A.分类ID,A.类别,C.名称 所属类别,C.固定 类别固定,A.编码,A.标识主码,A.标识子码,A.最高限价,A.最低限价,A.备选码,A.名称,A.规格,A.计算单位,A.费用类型, " & vbCrLf & _
                "       decode(A.服务对象,1,'门诊',2,'住院',3,'门诊与住院','无') as 服务对象," & vbCrLf & _
                "       decode(A.补充摘要,1,'√','') as 补充摘要, A.说明,decode(A.屏蔽费别,1,'√','') as 屏蔽费别," & vbCrLf & _
                "       decode(A.是否变价,1,'√','') as 是否变价,decode(A.加班加价,1,'√','') as 加班加价,decode(A.执行科室,1,4,2,1,3,3,0) AS 执行科室, " & vbCrLf & _
                "       to_char(A.建档时间,'yyyy-mm-dd') as 建档时间,to_char(A.撤档时间,'yyyy-mm-dd') as 撤档时间," & vbCrLf & _
                "       decode(A.类别,'1',decode(A.项目特性,1,'急诊项目','挂号项目'),'H',decode(A.项目特性,1,'护理等级',2,'基本护理','')) As 项目特性," & vbCrLf & _
                "        Nvl(B.名称,'') As 所属分类,a.病案费目,d.编号 As 站点编号, d.名称 As 站点名称" & vbCrLf & _
                "From 收费项目目录 A  ,收费分类目录 B,收费项目类别 C, zlnodelist D " & vbCrLf & _
                "Where A.类别=C.编码 And a.站点 = d.编号(+) and A.分类id=B.Id And  A.类别<>'5'  And A.类别<>'6' And A.类别<>'7'" & strTemp & " And  A.分类id  = [1] "
    End If
    
    Set rs项目 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(str分类), gstrNodeNo)
    'Call zlInitLvwHeadCol(1 )
    zlControl.FormLock lvwMain_S.hwnd
    With lvwMain_S.ListItems
        .Clear
        '名称;编码;规格;计算单位;费用类型;服务对象;说明;屏蔽费别;是否变价;加班加价;补充摘要;项目特性;建档时间;撤档时间;所属分类
        If rs项目.RecordCount > 0 Then
            rs项目.MoveFirst
            Dim lngCol  As Long
            Dim varValue As Variant
            For i = 0 To rs项目.RecordCount - 1
                '得出正确的图标
                strTemp = "Item"
                If Not CDate(IIF(IsNull(rs项目("撤档时间")), CDate("3000/1/1"), rs项目("撤档时间"))) = CDate("3000/1/1") Then
                    strTemp = strTemp & "No"
                End If
                '添加节点
                Set lst = .Add(, "C" & rs项目("类别") & rs项目("id"), rs项目("名称"), strTemp, strTemp)
                If InStr(strTemp, "No") > 0 Then lst.ForeColor = RGB(255, 0, 0)
                lst.Tag = rs项目!ID
                
                '根据ListView的列名从数据库取数
                For lngCol = 2 To lvwMain_S.ColumnHeaders.Count
                    If Trim(lvwMain_S.ColumnHeaders(lngCol).Text) = "所属类别" And rs项目!类别固定 = 1 Then
                            varValue = "[" & rs项目(lvwMain_S.ColumnHeaders(lngCol).Text).value & "]"
                    ElseIf Trim(lvwMain_S.ColumnHeaders(lngCol).Text) = "最高限价" Or Trim(lvwMain_S.ColumnHeaders(lngCol).Text) = "最低限价" Then
                        If IsNull(rs项目(lvwMain_S.ColumnHeaders(lngCol).Text).value) Then
                            varValue = " "
                        Else
                            varValue = CStr(Format(rs项目(lvwMain_S.ColumnHeaders(lngCol).Text).value, "0.00"))
                        End If
                    ElseIf lvwMain_S.ColumnHeaders(lngCol).Text = "院区" Then
                        varValue = rs项目("站点名称").value
                    Else
                        varValue = rs项目(lvwMain_S.ColumnHeaders(lngCol).Text).value
                    End If
                    lst.SubItems(lngCol - 1) = IIF(IsNull(varValue), "", varValue)
                    If InStr(strTemp, "No") > 0 Then lst.ListSubItems(lngCol - 1).ForeColor = RGB(255, 0, 0)
                Next
                '记录该项目的类别
                lst.ListSubItems(2).Tag = Nvl(rs项目("类别"))
                '记录该项目的站点
                lst.ListSubItems(lst.ListSubItems.Count).Tag = Nvl(rs项目("站点编号"))
                rs项目.MoveNext
            Next
        End If
    End With
    If rs项目.RecordCount > 0 Then
        Me.stbThis.Panels(2).Text = "收费项目数据读取完成！"
        Me.mnuFileExcel.Enabled = True
        Me.mnuFilePrint.Enabled = True
        Me.mnuFilepre.Enabled = True
        Me.mnuFilePrint.Enabled = True
    Else
        Me.stbThis.Panels(2).Text = "当前分类无项目"
        Me.mnuFileExcel.Enabled = False
        Me.mnuFilePrint.Enabled = False
        Me.mnuFilepre.Enabled = False
        Me.mnuFilePrint.Enabled = False
    End If
    Toolbar1.Buttons("Print").Enabled = mnuFileExcel.Enabled
    Toolbar1.Buttons("Preview").Enabled = mnuFileExcel.Enabled
    If Me.ActiveControl Is tvwMainItem Then
        If tvwMainItem.SelectedItem Is Nothing Then
            mnuEditModifyAssort.Enabled = False
            mnuEditDeleteAssort.Enabled = False
        End If
        mnuEditModify.Enabled = False
        mnuEditDelete.Enabled = False
        mnuEditStart.Enabled = False
        mnuEditStop.Enabled = mnuEditStart.Enabled
        mnuEditDept.Enabled = mnuEditStart.Enabled
        mnuEditSlave.Enabled = mnuEditStart.Enabled
        mnuEditItemGroup.Enabled = mnuEditStart.Enabled
        mnuPriceChargeSet.Enabled = mnuEditStart.Enabled
        mnuPriceHistory.Enabled = mnuEditStart.Enabled
        mnuPriceRaise.Enabled = mnuEditStart.Enabled
        mnuEditCopy.Enabled = mnuEditStart.Enabled
        
        Toolbar1.Buttons("Modify").Enabled = True
        Toolbar1.Buttons("Delete").Enabled = True
        Toolbar1.Buttons("Start").Enabled = mnuEditStart.Enabled
        Toolbar1.Buttons("Stop").Enabled = mnuEditStop.Enabled
    Else
        If lvwMain_S.ListItems.Count > 0 Then
            tbPage.Enabled = True
            
            Dim Item As ListItem
            On Error Resume Next
            Set Item = lvwMain_S.ListItems(strKey)
            If Err <> 0 Then
                Set Item = lvwMain_S.ListItems(1)
                Item.Selected = True
                Item.EnsureVisible
                lvwMain_S_ItemClick Item
            Else
                Err.Clear
                Item.Selected = True
                Item.EnsureVisible
                lvwMain_S_ItemClick Item
            End If
        Else
            mnuEditStart.Enabled = False
            mnuEditStop.Enabled = mnuEditStart.Enabled
            mnuEditDelete.Enabled = mnuEditStart.Enabled
            mnuEditDept.Enabled = mnuEditStart.Enabled
            mnuEditModify.Enabled = mnuEditStart.Enabled
            mnuEditSlave.Enabled = mnuEditStart.Enabled
            mnuEditItemGroup.Enabled = mnuEditStart.Enabled
            mnuPriceChargeSet.Enabled = mnuEditStart.Enabled
            mnuPriceHistory.Enabled = mnuEditStart.Enabled
            mnuPriceRaise.Enabled = mnuEditStart.Enabled
            mnuEditCopy.Enabled = mnuEditStart.Enabled
            
            Toolbar1.Buttons("Modify").Enabled = mnuEditStart.Enabled
            Toolbar1.Buttons("Delete").Enabled = mnuEditStart.Enabled
            Toolbar1.Buttons("Start").Enabled = mnuEditStart.Enabled
            Toolbar1.Buttons("Stop").Enabled = mnuEditStop.Enabled
            
            FillItem ""
        End If
    End If
    zlControl.FormLock 0
    Screen.MousePointer = vbDefault
    Exit Sub
errHandleList:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Screen.MousePointer = vbHourglass
        Resume
    End If
    Call SaveErrLog
    zlControl.FormLock 0
End Sub

Public Sub FillItem(ByVal str项目 As String)
'功能:显示末级收费细目的价目,从属项目和执行科室
'参数:str项目 项目的标识
    Dim rsTemp As New ADODB.Recordset
    Dim strID As String
    Dim lst As ListItem
    Dim i As Integer, j As Integer
    Dim datCurr As Date
    Dim strSQL As String
    Dim iRow As Integer, icol As Integer
    Dim strTmp As String, str价格等级 As String
    Dim ObjItem As ListItem
    
    On Error GoTo ErrHandle
    
    MenuSet
    If str项目 = "" Then
        mstr上级Key = ""
        tbPage.Enabled = False
        
        msh价目.Clear 1
        msh价目.Rows = 2
        
        msh从属.Rows = 2
        For i = 0 To msh从属.Cols - 1
            msh从属.TextMatrix(1, i) = ""
        Next
        mshAlias.Rows = 2
        For i = 0 To mshAlias.Cols - 1
            mshAlias.TextMatrix(1, i) = ""
        Next
        opt科室(4).value = True
        
        If Mid(tvwMainItem.SelectedItem.Key, 2, 1) = "F" Then
            msh价目.ColWidth(Col_附加手术收费率) = 1500
            msh价目.TextMatrix(0, Col_附加手术收费率) = "附加手术收费率"
        Else
            msh价目.ColWidth(Col_附加手术收费率) = 0
        End If
    Else
        tbPage.Enabled = True
    End If
    rsTemp.CursorLocation = adUseClient
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockReadOnly
    
    If lvwMain_S.ListItems.Count = 0 Then Exit Sub
    Set lst = lvwMain_S.ListItems(str项目)
    strID = Mid(str项目, 3)
    
    '调整表格
    If Mid(str项目, 2, 1) = "F" Then
        msh价目.ColWidth(Col_附加手术收费率) = 1500
        msh价目.TextMatrix(0, Col_附加手术收费率) = "附加手术收费率"
    Else
        msh价目.ColWidth(Col_附加手术收费率) = 0
    End If
    
    gstrSQL = "select a.是否变价,a.加班加价,a.执行科室,b.编码,b.名称 类别,a.分类ID,a.启用原因,a.停用原因,a.撤档时间 from 收费项目目录   A,收费项目类别 B  where   a.类别=b.编码  AND a.ID=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strID))
    
    If Not CDate(IIF(IsNull(rsTemp("撤档时间")), CDate("3000/1/1"), rsTemp("撤档时间"))) = CDate("3000/1/1") Then
        If Nvl(rsTemp!停用原因) = "" Then
            Me.pic停用原因.Visible = False
        Else
            Me.pic停用原因.Visible = True
            Me.lbl停用原因.Caption = "停用原因：" & rsTemp!停用原因
        End If
    Else
        If Nvl(rsTemp!启用原因) = "" Then
            Me.pic停用原因.Visible = False
        Else
            Me.pic停用原因.Visible = True
            Me.lbl停用原因.Caption = "启用原因：" & rsTemp!启用原因
        End If
    End If
    
    If rsTemp.RecordCount > 0 Then
        If IsNull(rsTemp("分类ID")) Then
            mstr上级Key = "R" & rsTemp("类别")
        Else
            mstr上级Key = "C" & rsTemp("类别") & rsTemp("分类ID")
        End If
        mstrClass = Nvl(rsTemp!编码)
        mstrClassName = Nvl(rsTemp!类别)
    Else
        mstrClass = ""
        mstrClassName = ""
        mstr上级Key = ""
        MsgBox "该项目不存在！", vbExclamation, gstrSysName
        Call mnuViewRefresh_Click
        Exit Sub
    End If
    
    If rsTemp("是否变价") = 1 Then
        msh价目.Rows = 2
        msh价目.TextMatrix(0, Col_原价) = "最低限价"
        msh价目.TextMatrix(0, Col_现价) = "最高限价"
        msh价目.ColWidth(Col_缺省价格) = 1000
    Else
        msh价目.TextMatrix(0, Col_原价) = "原价"
        msh价目.TextMatrix(0, Col_现价) = "现价"
        msh价目.ColWidth(Col_缺省价格) = 0
    End If
    
    If rsTemp("加班加价") = 1 Then
        msh价目.ColWidth(Col_加班加价率) = 1500
        msh价目.TextMatrix(0, Col_加班加价率) = "加班加价率"
    Else
        msh价目.ColWidth(Col_加班加价率) = 0
    End If
    '显示科室
    opt科室(IIF(rsTemp("执行科室") < 7, rsTemp("执行科室"), 0)).value = True
    lvwOutIn.ListItems.Clear
    
    rsTemp.Close
    If opt科室(4).value = True Then
        gstrSQL = " Select  " & _
            "   decode(b.编码,null,'','['||b.编码||']'|| b.名称) As 开单科室,  " & vbCrLf & _
            "    '['||c.编码||']'|| c.名称 As 执行科室" & vbCrLf & _
            " from 收费执行科室 A,部门表 B,部门表 C" & vbCrLf & _
            " Where a.开单科室id=b.Id(+) And a.执行科室id=C.id and 病人来源 is null and A.收费细目id=[1] " & _
            " order by  c.名称 "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strID))
                
        Do Until rsTemp.EOF
            If strTmp <> rsTemp!执行科室 Then
                i = i + 1
                Set ObjItem = Me.lvwOutIn.ListItems.Add(, "A" & i, rsTemp!执行科室)
                ObjItem.SubItems(1) = IIF(IsNull(rsTemp!开单科室), "（所有部门）", rsTemp!开单科室)
            Else
                Me.lvwOutIn.ListItems("A" & i).SubItems(1) = Me.lvwOutIn.ListItems("A" & i).SubItems(1) & "," & rsTemp!开单科室
            End If
            strTmp = rsTemp!执行科室
            rsTemp.MoveNext
        Loop
        If lvwOutIn.ListItems.Count > 0 Then
            lvwOutIn.ListItems(1).Selected = True
            lvwOutIn.ListItems(1).EnsureVisible
        End If
    ElseIf opt科室(0).value = True Then
        '无明确执行科室显示已设置的手工记帐缺省执行科室
        gstrSQL = "" & _
            " Select '[' || b.编码 || ']' || b.名称 As 执行科室" & vbNewLine & _
            " From 收费执行科室 A, 部门表 B" & vbNewLine & _
            " Where a.执行科室id = b.Id And a.病人来源 = 2 And a.收费细目id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strID))
        If Not rsTemp.EOF Then
            lvwOutIn.ListItems.Add , , Nvl(rsTemp!执行科室)
        End If
    End If
    
    '显示收费价目
    Call Fill价目(Val(strID))
    
    '显示从属项目
    gstrSQL = "select a.主项ID,a.从项ID,a.固有从属,a.从项数次,b.名称,b.编码 项目编码,c.编码,c.名称 类别," & _
        "Nvl(B.撤档时间,to_Date('3000-01-01','YYYY-MM-DD')) As 撤档时间 from 收费从属项目 a,收费项目目录 b ,收费项目类别 c where c.编码=b.类别 and a.从项ID=b.id and 主项ID=[1] " & _
        " ORDER BY a.ROWID "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strID))
        
    msh从属.Rows = IIF(rsTemp.RecordCount = 0, 2, rsTemp.RecordCount + 1)
    If rsTemp.RecordCount = 0 Then
        For i = 0 To 3
            msh从属.TextMatrix(1, i) = ""
        Next
    Else
        i = 1
        Do Until rsTemp.EOF
            msh从属.TextMatrix(i, 0) = "(" & rsTemp("编码") & ")" & rsTemp("类别")
            msh从属.TextMatrix(i, 1) = "[" & rsTemp("项目编码") & "]" & rsTemp("名称")
            msh从属.TextMatrix(i, 2) = rsTemp("从项数次")
            If rsTemp("固有从属") = 0 Then
                msh从属.TextMatrix(i, 3) = "0-不固定"
            ElseIf rsTemp("固有从属") = 2 Then
                msh从属.TextMatrix(i, 3) = "2-按比例计算"
            Else
                msh从属.TextMatrix(i, 3) = "1-固定"
            End If
            msh从属.TextMatrix(i, 4) = IIF(Format(rsTemp!撤档时间, "YYYY-MM-DD") <> "3000-01-01", "停用", "")

            iRow = msh从属.Row: icol = msh从属.Col
            msh从属.Row = i
            For j = 0 To msh从属.Cols - 1
                msh从属.Col = j
                msh从属.CellForeColor = IIF(Format(rsTemp!撤档时间, "YYYY-MM-DD") <> "3000-01-01", &HFF&, vbBlack)
            Next
            msh从属.Row = iRow: msh从属.Col = icol
            
            i = i + 1
            rsTemp.MoveNext
        Loop
    End If
    '显示别名
    gstrSQL = "select decode( 性质,1,'正名',2,'英文名',3,'拉丁名',4,'化学名',5,'商品名',9,'其他别名','') 名称种类,名称," & _
        "   decode(码类,1,'拼音码',2,'五笔码',3,'数字码','')  码类,nvl(简码,'') 简码" & _
        " from 收费项目别名 where 收费细目ID=[1]" & _
        " order by 名称种类 ,名称 "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strID))
        
    mshAlias.Clear
    mshAlias.Rows = IIF(rsTemp.RecordCount = 0, 2, rsTemp.RecordCount + 1)
    
    If rsTemp.RecordCount = 0 Then
        For i = 0 To 3
            mshAlias.TextMatrix(1, i) = ""
        Next
    Else
        mshAlias.TextMatrix(0, 0) = "名称种类"
        mshAlias.TextMatrix(0, 1) = "名称"
        mshAlias.TextMatrix(0, 2) = "码类"
        mshAlias.TextMatrix(0, 3) = "简码"
        
        i = 1
        Do Until rsTemp.EOF
            mshAlias.TextMatrix(i, 0) = rsTemp("名称种类")
            mshAlias.TextMatrix(i, 1) = rsTemp("名称")
            mshAlias.TextMatrix(i, 2) = rsTemp("码类")
            mshAlias.TextMatrix(i, 3) = Nvl(rsTemp("简码"))
            i = i + 1
            rsTemp.MoveNext
        Loop
        mshAlias.ColAlignment(0) = flexAlignLeftCenter
        mshAlias.ColAlignment(1) = 4
        mshAlias.ColAlignment(2) = flexAlignLeftCenter
        mshAlias.ColAlignment(3) = flexAlignLeftCenter
        mshAlias.MergeCells = flexMergeRestrictColumns
        mshAlias.MergeCol(0) = True
        mshAlias.MergeCol(1) = True
    End If
    
    '显示费别等级
    gstrSQL = "Select A.费别, A.段号, 应收段首值, 应收段尾值, 实收比率, Decode(计算方法, 1, '1-成本价加收比例计算', '0-分段比例计算') As 计算方法 " & _
            " From 费别明细 A, 收费项目目录 B " & _
            " Where A.收费细目id = B.ID And A.收费细目id = [1] " & _
            " Order By A.费别, A.段号, A.应收段首值"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strID))
    
    msh费别.Clear
    msh费别.Rows = IIF(rsTemp.RecordCount = 0, 2, rsTemp.RecordCount + 1)
    
    With msh费别
        If rsTemp.RecordCount = 0 Then
            .TextMatrix(0, 0) = "费别"
            .TextMatrix(0, 1) = "应收金额(元)"
            .TextMatrix(0, 2) = "实收比率(%)"
            .TextMatrix(0, 3) = "计算方法"
            For i = 0 To 3
                msh费别.TextMatrix(1, i) = ""
            Next
        Else
            .TextMatrix(0, 0) = "费别"
            .TextMatrix(0, 1) = "应收金额(元)"
            .TextMatrix(0, 2) = "实收比率(%)"
            .TextMatrix(0, 3) = "计算方法"
            
            i = 1
            Do Until rsTemp.EOF
                .TextMatrix(i, 0) = rsTemp.Fields("费别").value
                .TextMatrix(i, 1) = Format(rsTemp.Fields("应收段首值").value, "##########0.00;-#########0.00;0.00;0.00") & _
                    " ～ " & Format(rsTemp.Fields("应收段尾值").value, "##########0.00;-#########0.00;0.00;0.00")
                .TextMatrix(i, 2) = Format(rsTemp.Fields("实收比率").value, "###0.000;-##0.000;0.000;0.000")
                .TextMatrix(i, 3) = rsTemp.Fields("计算方法").value
                i = i + 1
                rsTemp.MoveNext
            Loop
            .ColAlignment(1) = 1
            .ColAlignment(2) = 1
            .MergeCells = flexMergeRestrictColumns
            .MergeCol(0) = True
        End If
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    msh从属.redraw = True
    msh价目.redraw = flexRDBuffered
    mshAlias.redraw = True
End Sub

Private Function Fill价目(ByVal lng细目ID As Long) As Boolean
    '显示收费价目
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer
    Dim datCurr As Date
    Dim str价格等级 As String, strWhere As String
    
    On Error GoTo ErrHandle
    With msh价目
        .redraw = flexRDNone
        .Clear 1
        .Rows = 2
        .Cell(flexcpBackColor, 1, 0, 1, .Cols - 1) = &H80000005
        .Subtotal flexSTClear
        
        datCurr = sys.Currentdate
        If mblnCanUpdateAll Then
            strWhere = "      And (a.价格等级 Is Null Or Exists (Select 1 From 收费价格等级应用 Where 价格等级 = a.价格等级))"
        Else
            strWhere = _
                "      And (a.价格等级 Is Null" & vbNewLine & _
                "           Or Exists(Select 1" & vbNewLine & _
                "               From 收费项目目录 M, 收费价格等级应用 N" & vbNewLine & _
                "               Where m.Id = a.收费细目id And c.名称 = n.价格等级 And n.站点 = [2]))"
        End If
        gstrSQL = "" & _
                "Select a.价格等级, a.No, a.Id, a.原价id, a.收入项目id, a.原价, a.现价, Nvl(a.缺省价格, 0) As 缺省价格," & vbNewLine & _
                "       a.收费细目id, b.名称, a.加班加价率, a.附术收费率, a.变动原因, a.调价说明, a.执行日期, a.终止日期, a.调价人" & vbNewLine & _
                "From 收费价目 A, 收入项目 B, 收费价格等级 C" & vbNewLine & _
                "Where a.收入项目id = b.Id And a.价格等级 = c.名称(+) And a.收费细目id = [1]" & vbNewLine & _
                    IIF(chk价格.value = 1, "", "And (a.终止日期 Is Null Or a.终止日期 > Sysdate)") & vbNewLine & _
                "      And (c.撤档时间 Is Null Or c.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd'))" & vbNewLine & _
                    strWhere & vbNewLine & _
                "Order By Nvl(c.编码, ' '), a.执行日期 Desc"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng细目ID, gstrNodeNo)
        
        .Rows = IIF(rsTemp.RecordCount = 0, 2, rsTemp.RecordCount + 1)
        i = 1
        Do Until rsTemp.EOF
            If InStr("," & str价格等级 & ",", "," & Nvl(rsTemp!价格等级, "缺省") & ",") = 0 Then
                str价格等级 = str价格等级 & "," & Nvl(rsTemp!价格等级, "缺省")
            End If
            .TextMatrix(i, Col_价格等级) = Nvl(rsTemp!价格等级, "缺省")
            .TextMatrix(i, Col_单据号) = Nvl(rsTemp!NO)
            .TextMatrix(i, Col_执行日期) = Format(Nvl(rsTemp!执行日期), "yyyy-MM-dd hh:mm:ss")
            .TextMatrix(i, Col_终止日期) = Format(Nvl(rsTemp!终止日期), "yyyy-MM-dd hh:mm:ss")
            .TextMatrix(i, Col_收入项目) = Nvl(rsTemp!名称)
            .TextMatrix(i, Col_原价) = Format(Nvl(rsTemp!原价), "###########0.000;-##########0.000; ; ")
            .TextMatrix(i, Col_现价) = Format(Nvl(rsTemp!现价), "###########0.000;-##########0.000; ; ")
            .TextMatrix(i, Col_附加手术收费率) = Val(Nvl(rsTemp!附术收费率))
            .TextMatrix(i, Col_加班加价率) = Val(Nvl(rsTemp!加班加价率))
            .TextMatrix(i, Col_调价说明) = Nvl(rsTemp!调价说明)
            .TextMatrix(i, Col_缺省价格) = Format(Nvl(rsTemp!缺省价格), "###########0.000;-##########0.000; ; ")
            .TextMatrix(i, Col_调价人) = Nvl(rsTemp!调价人)
            .RowData(i) = rsTemp("ID")
            '看它是不是现行价格
            If rsTemp("执行日期") <= datCurr Then
                If CDate(Nvl(rsTemp!终止日期, "3000-01-01")) >= datCurr Then
                    .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HDFFFFF
                End If
            End If
            i = i + 1
            rsTemp.MoveNext
        Loop
        '分组显示
        If str价格等级 <> "" Then str价格等级 = Mid(str价格等级, 2)
        If UBound(Split(str价格等级, ",")) <= 0 Then
            .ColHidden(Col_价格等级) = True
        Else
            .ColHidden(Col_价格等级) = False
            .OutlineBar = flexOutlineBarComplete
            .MultiTotals = True
    
            .Subtotal flexSTNone, Col_价格等级, , , , , True, "%s", , True
            .SubtotalPosition = flexSTAbove
    
            .Outline Col_价格等级
            .OutlineCol = Col_价格等级
    
            .MergeCells = flexMergeRestrictRows
            .MergeRow(-1) = False
            
            For i = 1 To .Rows - 1
                If .IsSubtotal(i) Then
                    .Cell(flexcpText, i, 0, i, .Cols - 1) = .TextMatrix(i + 1, Col_价格等级)
                    .MergeRow(i) = True '该行合并
                    .IsCollapsed(i) = flexOutlineExpanded  '是否展开状态
                End If
            Next
        End If
        .redraw = flexRDBuffered
    End With
    Fill价目 = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    msh价目.redraw = flexRDBuffered
End Function

Private Sub chk价格_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim strID As String
    Dim i As Integer
    Dim datCurr As Date
    
    On Error GoTo ErrHandle
    msh价目.Clear 1
    msh价目.Rows = 2
    msh价目.Cell(flexcpBackColor, 1, 0, 1, msh价目.Cols - 1) = &H80000005
    msh价目.Subtotal flexSTClear
    
    If lvwMain_S.SelectedItem Is Nothing Then Exit Sub
    strID = Mid(lvwMain_S.SelectedItem.Key, 3)
    Call Fill价目(Val(strID))
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub SetPageVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置成套项目页面数据或设置细目数据页的显示
    '编制:刘兴洪
    '日期:2010-08-27 16:39:39
    '问题:27327
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnShowWholeSet As Boolean  '是否成套显示
    Dim i As Long
    If tbClassPage.Selected Is Nothing Then
        blnShowWholeSet = False
    Else
        blnShowWholeSet = Val(tbClassPage.Selected.Tag) <> mCalssPage.pg_细目
    End If
    With tbPage
        For i = 0 To .ItemCount - 1
            If Val(.Item(i).Tag) = mItemPage.pg_成套使用科室 Or Val(.Item(i).Tag) = mItemPage.pg_成套组成 Then
                .Item(i).Visible = blnShowWholeSet
                If Val(.Item(i).Tag) = mint上次成套页 And .Item(i).Visible Then
                        .Item(i).Selected = True
                End If
            Else
                .Item(i).Visible = Not blnShowWholeSet
                If Val(.Item(i).Tag) = mint上次细目页 And .Item(i).Visible Then
                        .Item(i).Selected = True
                End If
            End If
        Next
    End With

End Sub
Private Sub zlSetWholeSetMenu()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置成套项目的相关菜单
    '编制:刘兴洪
    '日期:2010-08-25 17:01:26
    '问题:27327
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '先设置菜单,屏蔽所有的非成套项目编辑
    Dim blnAdd As Boolean, blnModify As Boolean, blnDelete As Boolean
    Dim blnEdit As Boolean
    If Val(tbClassPage.Selected.Tag) = mCalssPage.pg_细目 Then
        '是细目,重新设置权限
        mnuEdit.Visible = True: mnuPrice.Visible = True
        mnuViewSelect.Visible = True: mnuViewShowStop.Visible = True
        mnuViewFind.Visible = True
        Call 权限控制
        mnuEditWholeSet.Visible = False
        lvwWholeSetItem_S.Visible = False
        lvwMain_S.Visible = True
        mnuPriceRaiseVerify.Visible = mblnVerifyFlow
        Toolbar1.Buttons("RaiseVerify").Visible = mblnVerifyFlow   '调价审核
        If mblnVerifyPris = False And mnuPriceRaiseVerify.Visible = True Then
            mnuPriceRaiseVerify.Enabled = False
            Toolbar1.Buttons("RaiseVerify").Enabled = False
        End If
        Exit Sub
    End If
    lvwWholeSetItem_S.Visible = True
    lvwMain_S.Visible = False
    
    mnuEdit.Visible = False: mnuPrice.Visible = False
    '列不多,不用设置
    mnuViewSelect.Visible = False: mnuViewShowStop.Visible = False  '没有显示停用项
    mnuPriceRaiseVerify.Visible = mblnVerifyFlow
    mnuViewFind.Visible = False '没有查找
    blnAdd = InStr(1, mstrPrivs, ";增加成套项目;") > 0
    blnModify = InStr(1, mstrPrivs, ";修改成套项目;") > 0
    blnDelete = InStr(1, mstrPrivs, ";删除成套项目;") > 0
    If blnAdd Or blnModify Or blnDelete Then
        mnuEditWholeSet.Visible = True
        mnuEditWholeSetClassAdd.Visible = blnAdd
        mnuEditWholeSetClassModify.Visible = blnModify
        mnuEditWholeSetClassDelete.Visible = blnDelete
        mnuEditWholeSetItemAdd.Visible = blnAdd
        mnuEditWholeSetItemModify.Visible = blnModify
        mnuEditWholeSetItemDelete.Visible = blnDelete
        Toolbar1.Buttons("Split4").Visible = True
    Else
        mnuEditWholeSet.Visible = False
        Toolbar1.Buttons("Split4").Visible = False
    End If
    Toolbar1.Buttons("Parent").Visible = blnAdd
    Toolbar1.Buttons("Child").Visible = blnAdd
    Toolbar1.Buttons("Child").Enabled = blnAdd
    Toolbar1.Buttons("Modify").Visible = blnModify
    Toolbar1.Buttons("Delete").Visible = blnDelete
    
    Toolbar1.Buttons("Split1").Visible = False
    Toolbar1.Buttons("Raise").Visible = False   '调价
    Toolbar1.Buttons("RaiseVerify").Visible = mblnVerifyFlow   '调价审核
    
    If mblnVerifyPris = False And mnuPriceRaiseVerify.Visible = True Then
        mnuPriceRaiseVerify.Enabled = False
        Toolbar1.Buttons("RaiseVerify").Enabled = False
    End If
        
    Toolbar1.Buttons("Split2").Visible = False
    Toolbar1.Buttons("Start").Visible = False   '启用
    Toolbar1.Buttons("Stop").Visible = False   '停用
    Toolbar1.Buttons("Split3").Visible = False
    Toolbar1.Buttons("Find").Visible = False  '查找
    Toolbar1.Buttons("Split4").Visible = False
    '控制是否可编辑
    blnEdit = Not lvwWholeSetItem_S.SelectedItem Is Nothing
    
    mnuEditWholeSetItemModify.Enabled = blnEdit
    mnuEditWholeSetItemDelete.Enabled = blnEdit
    If Me.ActiveControl Is tvwWholeSet Then
       '当前选中的是分类
       If tvwWholeSet.SelectedItem Is Nothing Then
            blnEdit = False
       Else
            blnEdit = Val(Mid(tvwWholeSet.SelectedItem.Key, 2)) <> 0
       End If
       mnuEditWholeSetClassModify.Enabled = blnEdit
       mnuEditWholeSetClassDelete.Enabled = blnEdit
       Toolbar1.Buttons("Modify").Enabled = blnEdit
       Toolbar1.Buttons("Delete").Enabled = blnEdit
       With tvwWholeSet
            If .SelectedItem Is Nothing Then
                stbThis.Panels(2).Text = ""
            Else
                stbThis.Panels(2).Text = "该分类共有" & .SelectedItem.Children & "个下级分类。"
            End If
        End With
       Exit Sub
    End If
    With lvwWholeSetItem_S
        stbThis.Panels(2).Text = "成套项目列表中共显示有" & .ListItems.Count & "个成套项目。"
    End With
    Toolbar1.Buttons("Modify").Enabled = mnuEditWholeSetItemModify.Enabled
    Toolbar1.Buttons("Delete").Enabled = mnuEditWholeSetItemDelete.Enabled
End Sub
Public Sub MenuSet()
'功能:显示菜单和工具栏的状态
'对于类别：
'   它是独立编辑时,不允许删改类别、也不允许类别和项目的增删除
'   它是系统标志时,不允许删改类别、但允许类别和项目的增删除
'对于分类：
'   永远允许增删除改
'对于项目：
'   当它是独立编辑时，不能修改和删除
'   当它是处于停用时，修改也不能做
'   其它情况就看权限的控制了
On Error GoTo ErrHandle
    Dim blnClassModify As Boolean, blnItemModify As Boolean '代表各种修改、删除
    Dim blnPrice As Boolean '代表各种调价、批量调价等
    Dim blnStart As Boolean '启用的状态
    Dim blnStop As Boolean '停用的状态
    Dim blnPrint As Boolean '打印
    Dim blnCanModify As Boolean
    
    '刘兴洪:27327
    Call zlSetWholeSetMenu
    If Val(tbClassPage.Selected.Tag) = mCalssPage.pg_成套 Then Exit Sub
    
    If ActiveControl Is tvwMainItem Then
        With tvwMainItem
            If Not tvwMainItem.SelectedItem Is Nothing Then
                blnClassModify = True
                stbThis.Panels(2).Text = "该分类共有" & .SelectedItem.Children & "个下级分类。"
            End If
        End With
    Else
        With lvwMain_S
            stbThis.Panels(2).Text = "项目列表中共显示有" & .ListItems.Count & "个项目。"
            blnPrint = .ListItems.Count > 0
            
            If Not .SelectedItem Is Nothing Then
                blnCanModify = mblnCanUpdateAll _
                    Or .SelectedItem.ListSubItems(.SelectedItem.ListSubItems.Count).Tag = gstrNodeNo
            
                If InStr(.SelectedItem.Icon, "No") > 0 Then
                    '停用
                    blnStop = (.SelectedItem.Icon = "ItemNo") And blnCanModify
                Else
                    blnStart = (.SelectedItem.Icon = "Item") And blnCanModify
                    blnPrice = True
                End If
                blnItemModify = blnStart
            End If
        End With
    End If
    
    '编辑
    mnuEditParent.Enabled = True '新增分类
    mnuEditModifyAssort.Enabled = blnClassModify '修改分类
    mnuEditDeleteAssort.Enabled = blnClassModify '删除分类
    
    mnuEditChild.Enabled = True '新增项目
    mnuEditCopy.Enabled = Not lvwMain_S.SelectedItem Is Nothing '复制新增
    mnuEditModify.Enabled = blnItemModify '修改项目
    mnuEditDelete.Enabled = blnItemModify '删除项目
    
    Toolbar1.Buttons("Modify").Enabled = blnClassModify Or blnItemModify
    Toolbar1.Buttons("Delete").Enabled = blnClassModify Or blnItemModify
    
    mnuEditDept.Enabled = blnItemModify   '执行科室
    mnuEditSlave.Enabled = blnItemModify '从属项目
    mnuEditItemGroup.Enabled = blnItemModify '项目组成
    
    mnuEditStart.Enabled = blnStop   '启用
    mnuEditStop.Enabled = blnStart  '停用
    Toolbar1.Buttons("Start").Enabled = blnStop
    Toolbar1.Buttons("Stop").Enabled = blnStart
    
    '价目管理
    mnuPriceRaise.Enabled = blnPrice '调价
    Toolbar1.Buttons("Raise").Enabled = blnPrice
    If gstr医价接口编号 <> "" And gbln允许医价收费项目 Then
        mnuPriceRaiseMass.Enabled = False '批量调价
    Else
        mnuPriceRaiseMass.Enabled = blnPrice
    End If
    mnuPriceHistory.Enabled = blnPrice '删除未执行价格
    
    mnuPriceChargeSet.Enabled = (InStr(mstrPrivs, "费别设置") > 0) And Not lvwMain_S.SelectedItem Is Nothing '费别设置
    mnuEditItemGroup.Enabled = (InStr(mstrPrivs, "价目管理") > 0)
    
    '打印
    mnuFilepre.Enabled = blnPrint
    mnuFilePrint.Enabled = blnPrint
    mnuFileExcel.Enabled = blnPrint
    Toolbar1.Buttons("Preview").Enabled = blnPrint
    Toolbar1.Buttons("Print").Enabled = blnPrint
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub 权限控制()
    Dim rsTmp As New ADODB.Recordset
    
    '初始化工具栏按钮
    On Error GoTo ErrHandle
    Toolbar1.Buttons("Split1").Visible = True
    Toolbar1.Buttons("Start").Visible = True
    Toolbar1.Buttons("Stop").Visible = True
    Toolbar1.Buttons("Raise").Visible = True
    Toolbar1.Buttons("Split2").Visible = True
    Toolbar1.Buttons("Find").Visible = True
    Toolbar1.Buttons("Split3").Visible = True
    Toolbar1.Buttons("Split4").Visible = True
    
    '功能:由于有的用户权限不够,故使一些菜单项或按钮不可见
    If InStr(mstrPrivs, "项目管理") = 0 Then
        mnuEditChild.Visible = False  '增加项目
        mnuEditCopy.Visible = False   '复制拷贝
        mnuEditModify.Visible = False '修改
        mnuEditDelete.Visible = False '删除
        mnuEditDept.Visible = False   '执行科室
        mnuEditStart.Visible = False  '启用
        mnuEditStop.Visible = False   '停用
        mnuEditSplit0.Visible = False '第一个分隔
        mnuEditSplit1.Visible = False '第二个分隔
        
        mnuShortMenu2(0).Visible = False  '项目的快捷菜单编辑功能不可见
        mnuShortMenu2(1).Visible = False
        mnuShortMenu2(2).Visible = False
        mnuShortMenu2(3).Visible = False
        mnuShortMenu2(5).Visible = False
        
        mnuShortsplit1.Visible = False
        Toolbar1.Buttons("Child").Visible = False
        Toolbar1.Buttons("Modify").Visible = False
        Toolbar1.Buttons("Delete").Visible = False
        Toolbar1.Buttons("Split1").Visible = False
        Toolbar1.Buttons("Start").Visible = False
        Toolbar1.Buttons("Stop").Visible = False
        Toolbar1.Buttons("Split3").Visible = False
    End If
    
    If InStr(mstrPrivs, "类别管理") = 0 Then
        mnuClassEdit.Visible = False
        mnuEditSplit3.Visible = False
        Toolbar1.Buttons("Parent").Enabled = False
        Me.mnuEditParent.Enabled = False
        Me.mnuEditModifyAssort.Enabled = False
        Me.mnuEditDeleteAssort.Enabled = False
        Me.mnuShort1.Visible = False
        mnuEditParent.Visible = False '增加分类
        Toolbar1.Buttons("Parent").Visible = False
    End If
    
    If InStr(mstrPrivs, "项目组合设置") = 0 Then
        mnuEditSlave.Visible = False  '从属项目
        mnuEditItemGroup.Visible = False    '项目组成
        mnuShortMenu2(4).Visible = False
        mnuShortMenu2(6).Visible = False
    End If
    
    If InStr(mstrPrivs, "价目管理") = 0 Then
        mnuPrice.Visible = False
        mnuPriceRaise.Visible = False
        mnuPriceRaiseMass.Visible = False
        mnuPriceHistory.Visible = False
        Toolbar1.Buttons("Raise").Visible = False
        Toolbar1.Buttons("Split2").Visible = False
    End If
    
    If InStr(mstrPrivs, "费别设置") = 0 Then
        mnuPriceChargeSet.Visible = False
    End If
    
    If InStr(mstrPrivs, "医价接口") = 0 Then
        Me.mnuFileStdImp.Visible = False
        Me.mnuFileStdCheck.Visible = False
        Me.mnuFileSplit1.Visible = False
    Else
        Me.mnuFileStdImp.Visible = True
        Me.mnuFileStdCheck.Visible = True
        Me.mnuFileSplit1.Visible = True
        
        '得到医价接口编码
        gstrSQL = "select 编号,医疗 from 医价接口 where nvl(选用,0)=1"
        Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
        If rsTmp.RecordCount > 0 Then
            gstr医价接口编号 = Nvl(rsTmp!编号)
            gbln允许医价收费项目 = CStr(Nvl(rsTmp!医疗)) = "1"
            mnuFileStdImp.Enabled = True
            mnuFileStdCheck.Enabled = True
            mbln启动医价系统 = True
        Else
            gstr医价接口编号 = ""
            gbln允许医价收费项目 = False
            mnuFileStdImp.Enabled = False
            mnuFileStdCheck.Enabled = False
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub GetDefineSize()
'功能：得到数据库的表字段的长度
On Error GoTo ErrHandle
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select 编码 From 收费项目目录 Where Rownum<0"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "收费项目目录")
    
    mlng编码长度 = rsTemp.Fields("编码").DefinedSize
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub


Private Sub picPage_Resize(Index As Integer)
        Err = 0: On Error Resume Next
        With picPage(Index)
                Select Case Index
                Case 1
                    msh价目.Left = 50
                    msh价目.Height = .ScaleHeight - 100 - msh价目.Top
                    msh价目.Width = .ScaleWidth - msh价目.Left - 100
                Case 2
                    Frame2.Top = 0
                    Frame2.Left = 0
                    Frame2.Width = .ScaleWidth
                    Frame2.Height = .ScaleHeight
                    
                    Frame1.Width = .ScaleWidth - 300
                    Frame1.Height = .ScaleHeight - Frame1.Top - 200
                    lvwOutIn.Top = 0
                    lvwOutIn.Left = 0
                    lvwOutIn.Width = Frame1.Width
                    lvwOutIn.Height = Frame1.Height - lvwOutIn.Top
                Case 3
                    msh从属.Top = 800
                    msh从属.Height = .ScaleHeight - 100 - msh从属.Top
                    msh从属.Width = .ScaleWidth - 300
                Case 4
                    mshAlias.Left = 0
                    mshAlias.Width = .ScaleWidth
                    mshAlias.Top = 0
                    mshAlias.Height = .ScaleHeight
                Case 5
                    msh费别.Left = 0
                    msh费别.Width = .ScaleWidth
                    msh费别.Top = 0
                    msh费别.Height = .ScaleHeight
                Case 6
                    vsWholeSet.Left = 0
                    vsWholeSet.Top = 0
                    vsWholeSet.Width = .ScaleWidth
                    vsWholeSet.Height = .ScaleHeight
                Case 7
                    lvwUseDept.Left = 0
                    lvwUseDept.Top = 0
                    lvwUseDept.Width = .ScaleWidth
                    lvwUseDept.Height = .ScaleHeight
                End Select
        End With
End Sub

Private Sub tvwWholeSet_GotFocus()
    Call MenuSet
End Sub

Private Sub tvwWholeSet_LostFocus()
    Call MenuSet
End Sub

Private Sub tvwWholeSet_NodeClick(ByVal Node As MSComctlLib.Node)
        '加载成套项目数据
        If tvwWholeSet.Tag <> Node.Key Then
            tvwWholeSet.Tag = Node.Key
            Call FillWholeItem(Val(Mid(Node.Key, 2)))
        End If
        Call MenuSet
End Sub

Private Sub txtFind_GotFocus()
    Call zlControl.TxtSelAll(txtFind)
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSQL As String
    Dim rsTmp As Recordset
    Dim vRect As RECT, blnCancel As Boolean
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    On Error GoTo ErrHandle
    strSQL = "Select Distinct a.Id, a.类别, b.名称, a.编码, a.标识主码, a.标识子码, b.简码, c.名称 As 分类, a.分类id, a.撤档时间" & vbNewLine & _
            "From (Select ID, 类别, 分类id, 名称, 编码, 标识主码, 标识子码, 撤档时间" & vbNewLine & _
            "       From 收费项目目录" & vbNewLine & _
            "       Where 类别 <> '5' And 类别 <> '6' And 类别 <> '7'" & _
            IIF(mnuViewShowStop.Checked, "", " And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null)") & vbNewLine & _
            ") A, (Select a.收费细目id, a.名称, a.简码 || '/' || b.简码 As 简码" & vbNewLine & _
            "       From 收费项目别名 A, 收费项目别名 B" & vbNewLine & _
            "       Where a.收费细目id = b.收费细目id And a.码类 = 1 And b.码类 = 2) B, 收费分类目录 C" & vbNewLine & _
            "Where a.分类id = c.Id(+) And a.Id = b.收费细目id And c.名称 Is Not Null"
    If txtFind.Text = "" Then Exit Sub
    If zlStr.IsCharChinese(txtFind.Text) Then
        strSQL = strSQL & " And b.名称 Like [1]"
    ElseIf IsNumeric(txtFind.Text) Then
        strSQL = strSQL & " And a.编码 Like [2]"
    Else
        strSQL = strSQL & " And (b.名称 Like [1] Or b.简码 Like [3])"
    End If
    
    vRect = zlControl.GetControlRect(txtFind.hwnd)
    If vRect.Left + 7350 > Screen.Width Then vRect.Left = Screen.Width - 7350
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "收费细目选择", False, "", "", False, False, True, _
                        vRect.Left, vRect.Top, txtFind.Height, blnCancel, False, True, gstrLike & txtFind.Text & "%", txtFind.Text & "%", gstrLike & UCase(txtFind.Text) & "%")
    If blnCancel = True Then Exit Sub
    If Not rsTmp Is Nothing Then
        Call FindLocate(rsTmp)
    Else
        MsgBox "没有找到您所查找的收费项目。", vbInformation, Me.Caption
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FindLocate(ByVal rsTmp As Recordset)
'查找定位
    Dim strKey As String
    
    On Error Resume Next
        strKey = "R" & rsTmp!分类id
        If rsTmp!分类id & "" <> "" Then
            Me.tvwMainItem.Nodes(strKey).Selected = True
            Me.tvwMainItem.Nodes(strKey).EnsureVisible
            Me.tvwMainItem_NodeClick Me.tvwMainItem.SelectedItem
            Err.Clear
            Me.lvwMain_S.ListItems("C" & rsTmp!类别 & rsTmp!ID).Selected = True
            Me.lvwMain_S.ListItems("C" & rsTmp!类别 & rsTmp!ID).EnsureVisible
            If Err.Number = 35601 Then
                MsgBox "你找到的这条记录可能已被删除或停用，请刷新列表。", vbInformation, gstrSysName
                Err.Clear
                Exit Sub
            End If
            Me.lvwMain_S_ItemClick Me.lvwMain_S.SelectedItem
        Else
            Me.tvwMainItem.Nodes("Root").Selected = True
            Me.tvwMainItem.Nodes(strKey).EnsureVisible
            Me.tvwMainItem_NodeClick Me.tvwMainItem.SelectedItem
            Err.Clear
            Me.lvwMain_S.ListItems("C" & rsTmp!类别 & rsTmp!ID).Selected = True
            Me.lvwMain_S.ListItems("C" & rsTmp!类别 & rsTmp!ID).EnsureVisible
            If Err.Number = 35601 Then
                MsgBox "你找到的这条记录可能已被删除或停用，请刷新列表。", vbInformation, gstrSysName
                Err.Clear
                Exit Sub
            End If
            Me.lvwMain_S_ItemClick Me.lvwMain_S.SelectedItem
        End If
    Err.Clear
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0: Exit Sub
End Sub

Private Sub vsWholeSet_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngMode, vsWholeSet, Me.Caption, "成套项目组成表列-主界面", True, True
End Sub

Private Sub vsWholeSet_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngMode, vsWholeSet, Me.Caption, "成套项目组成表列-主界面", True, True
End Sub

Private Sub vsWholeSet_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsWholeSet
        Select Case Col
        Case .ColIndex("标志")
            Cancel = True
        Case Else
        End Select
    End With
End Sub
