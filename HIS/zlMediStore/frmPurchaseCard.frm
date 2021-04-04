VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmPurchaseCard 
   Caption         =   "药品外购入库单"
   ClientHeight    =   7320
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12495
   Icon            =   "frmPurchaseCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   12495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdAddProducer 
      Caption         =   "新增生产商(&P)"
      Height          =   350
      Left            =   2520
      TabIndex        =   54
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "批量复制(&C)"
      Height          =   350
      Left            =   6600
      TabIndex        =   53
      ToolTipText     =   "复制当前行发票信息应用于其他无发票信息行"
      Top             =   6120
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdALLDel 
      Caption         =   "全清(&D)"
      Height          =   350
      Left            =   4680
      TabIndex        =   52
      ToolTipText     =   "清除所有行的发票相关数据"
      Top             =   6120
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.PictureBox picInputCost 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   1665
      Left            =   5760
      ScaleHeight     =   1635
      ScaleWidth      =   5415
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   5445
      Begin VB.CommandButton cmdGetData 
         Caption         =   "提取(&G)"
         Height          =   300
         Left            =   4550
         TabIndex        =   51
         Top             =   0
         Width           =   855
      End
      Begin VB.ComboBox cboInputDate 
         Height          =   300
         Left            =   855
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   0
         Width           =   1440
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfInputCost 
         Height          =   1300
         Left            =   0
         TabIndex        =   48
         Top             =   300
         Width           =   5415
         _cx             =   9551
         _cy             =   2293
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
         Rows            =   1
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPurchaseCard.frx":014A
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
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入库时间"
         Height          =   180
         Left            =   0
         TabIndex        =   50
         Top             =   60
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdGetInputCost 
      Caption         =   "…"
      Height          =   300
      Left            =   1320
      TabIndex        =   46
      Top             =   6120
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CheckBox chk转入移库 
      Caption         =   "本张入库单药品移库到"
      Height          =   270
      Left            =   4680
      TabIndex        =   41
      Top             =   5700
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.ComboBox cboEnterStock 
      Enabled         =   0   'False
      Height          =   300
      ItemData        =   "frmPurchaseCard.frx":0204
      Left            =   6915
      List            =   "frmPurchaseCard.frx":020D
      TabIndex        =   40
      Text            =   "cboEnterStock"
      Top             =   5685
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.PictureBox PicInput 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   1665
      Left            =   210
      ScaleHeight     =   1635
      ScaleWidth      =   2775
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   1050
      Visible         =   0   'False
      Width           =   2805
      Begin VB.TextBox Txt加价率 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   960
         MaxLength       =   8
         TabIndex        =   37
         Text            =   "15.0000"
         Top             =   690
         Width           =   1725
      End
      Begin VB.CommandButton CmdNO 
         Caption         =   "取消"
         Height          =   345
         Left            =   1800
         TabIndex        =   39
         Top             =   1140
         Width           =   855
      End
      Begin VB.CommandButton CmdYes 
         Caption         =   "确定"
         Height          =   345
         Left            =   810
         TabIndex        =   38
         Top             =   1140
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "    请输入加成率，零售价的计算公式：零售价=成本价*(1+加成率%)"
         ForeColor       =   &H00400000&
         Height          =   585
         Left            =   0
         TabIndex        =   35
         Top             =   150
         Width           =   2805
      End
      Begin VB.Label Lbl加价率 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "加成率(&J)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   90
         TabIndex        =   36
         Top             =   750
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdAllCls 
      Caption         =   "全清(&L)"
      Height          =   350
      Left            =   9975
      TabIndex        =   33
      Top             =   6135
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdAllSel 
      Caption         =   "全冲(&A)"
      Height          =   350
      Left            =   8655
      TabIndex        =   32
      Top             =   6135
      Visible         =   0   'False
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh产地 
      Height          =   2175
      Left            =   2520
      TabIndex        =   31
      Top             =   1680
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3836
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
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   3030
      TabIndex        =   14
      Top             =   5685
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "查找(&F)"
      Height          =   350
      Left            =   1290
      TabIndex        =   13
      Top             =   5640
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   -15
      TabIndex        =   12
      Top             =   5640
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   8655
      TabIndex        =   10
      Top             =   5655
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   9975
      TabIndex        =   11
      Top             =   5655
      Width           =   1100
   End
   Begin VB.PictureBox Pic单据 
      BackColor       =   &H80000004&
      Height          =   5325
      Left            =   0
      ScaleHeight     =   5265
      ScaleWidth      =   12375
      TabIndex        =   15
      Top             =   0
      Width           =   12435
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshProvider 
         Height          =   1815
         Left            =   4800
         TabIndex        =   26
         Top             =   1080
         Visible         =   0   'False
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   3201
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
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   2805
         Left            =   195
         TabIndex        =   7
         Top             =   1020
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   4948
         Appearance      =   0
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Active          =   -1  'True
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483634
      End
      Begin VB.TextBox txt摘要 
         Height          =   300
         Left            =   900
         MaxLength       =   40
         TabIndex        =   9
         Top             =   4080
         Width           =   10410
      End
      Begin VB.CommandButton cmdProvider 
         Caption         =   "…"
         Height          =   300
         Left            =   11010
         TabIndex        =   20
         Top             =   660
         Width           =   300
      End
      Begin VB.TextBox txtProvider 
         Height          =   300
         Left            =   8055
         TabIndex        =   2
         Top             =   660
         Width           =   2895
      End
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   660
         Width           =   1515
      End
      Begin VB.TextBox txtNO 
         Height          =   315
         IMEMode         =   2  'OFF
         Left            =   9870
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   210
         Width           =   1425
      End
      Begin VB.Label lbl修改日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "修改日期"
         Height          =   180
         Left            =   3240
         TabIndex        =   58
         Top             =   4860
         Width           =   720
      End
      Begin VB.Label lbl修改人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  修改人"
         Height          =   180
         Left            =   3285
         TabIndex        =   57
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label Txt修改日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4050
         TabIndex        =   56
         Top             =   4800
         Width           =   1890
      End
      Begin VB.Label Txt修改人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4050
         TabIndex        =   55
         Top             =   4440
         Width           =   1890
      End
      Begin VB.Label txt核查人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7290
         TabIndex        =   45
         Top             =   4440
         Width           =   1890
      End
      Begin VB.Label txt核查日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7290
         TabIndex        =   44
         Top             =   4800
         Width           =   1890
      End
      Begin VB.Label lbl核查人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  核查人"
         Height          =   180
         Left            =   6525
         TabIndex        =   43
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label lbl核查日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "核查日期"
         Height          =   180
         Left            =   6480
         TabIndex        =   42
         Top             =   4860
         Width           =   720
      End
      Begin VB.Label lblDifference 
         AutoSize        =   -1  'True
         Caption         =   "差价合计:"
         Height          =   180
         Left            =   4920
         TabIndex        =   29
         Top             =   3840
         Width           =   810
      End
      Begin VB.Label lblSalePrice 
         AutoSize        =   -1  'True
         Caption         =   "售价金额合计:"
         Height          =   180
         Left            =   2040
         TabIndex        =   28
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "结算金额合计:"
         Height          =   180
         Left            =   240
         TabIndex        =   27
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label Txt审核人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   10410
         TabIndex        =   24
         Top             =   4440
         Width           =   1890
      End
      Begin VB.Label Txt审核日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   10410
         TabIndex        =   23
         Top             =   4800
         Width           =   1890
      End
      Begin VB.Label Txt填制日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   22
         Top             =   4800
         Width           =   1890
      End
      Begin VB.Label Txt填制人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   21
         Top             =   4440
         Width           =   1890
      End
      Begin VB.Label LblNo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NO."
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   9390
         TabIndex        =   1
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lbl摘要 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "摘要(&M)"
         Height          =   180
         Left            =   240
         TabIndex        =   8
         Top             =   4155
         Width           =   650
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "药品外购入库单"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   30
         TabIndex        =   0
         Top             =   120
         Width           =   11535
      End
      Begin VB.Label LblStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "库房(&S)"
         Height          =   180
         Left            =   195
         TabIndex        =   3
         Top             =   720
         Width           =   630
      End
      Begin VB.Label Lbl填制人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "填制人"
         Height          =   180
         Left            =   300
         TabIndex        =   19
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label Lbl填制日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "填制日期"
         Height          =   180
         Left            =   120
         TabIndex        =   18
         Top             =   4860
         Width           =   720
      End
      Begin VB.Label Lbl审核人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审核人"
         Height          =   180
         Left            =   9765
         TabIndex        =   17
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label Lbl审核日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审核日期"
         Height          =   180
         Left            =   9600
         TabIndex        =   16
         Top             =   4860
         Width           =   720
      End
      Begin VB.Label LblProvider 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "供药单位(&G)"
         Height          =   180
         Left            =   7035
         TabIndex        =   5
         Top             =   720
         Width           =   990
      End
   End
   Begin MSComctlLib.ImageList imghot 
      Left            =   840
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":021B
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":0435
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":064F
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":0869
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":0A83
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":0C9D
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":0EB7
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":10D1
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgcold 
      Left            =   120
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":12EB
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":1505
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":171F
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":1939
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":1B53
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":1D6D
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":1F87
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":21A1
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   30
      Top             =   6960
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPurchaseCard.frx":23BB
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15690
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmPurchaseCard.frx":2C4F
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmPurchaseCard.frx":3151
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
   Begin VB.Label lblCode 
      Caption         =   "编码"
      Height          =   255
      Left            =   2550
      TabIndex        =   25
      Top             =   5730
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Menu mnuCol 
      Caption         =   "列名"
      Visible         =   0   'False
      Begin VB.Menu mnuColDrug 
         Caption         =   "药名(编码和名称)"
         Index           =   0
      End
      Begin VB.Menu mnuColDrug 
         Caption         =   "药名(仅编码)"
         Index           =   1
      End
      Begin VB.Menu mnuColDrug 
         Caption         =   "药名(仅名称)"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmPurchaseCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng供药单位ID As Long              '供药单位ID
Private mint编辑状态 As Integer             '1.新增；2、修改；3、验收；4、查看；5、修改发票；6、冲销；
                                            '7、财务审核（冲销、产生新单据并审核；已付款的单据不允许财
                                            '务审核，同样，财务审核后的单据不允许冲销）;8-药库退货;9-核查
Private mstr单据号 As String                '具体的单据号;
Private mint记录状态 As Integer             '1:正常记录;2-冲销记录;3-已经冲销的原记录
Private mblnSuccess As Boolean              '只要有一张成功，即为True，否则为False
Private mblnSave As Boolean                 '是否存盘和审核   TURE：成功。
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEdit As Boolean                 '是否可以修改
Private mblnChange As Boolean               '是否进行过编辑
Private mintParallelRecord As Integer       '对于新增后单据并行执行的处理： 1、代表正常情况；2、已经删除的记录；3、已经审核的记录；7、已全部付款
Private mstrPrivs As String                 '权限
Private mintBatchNoLen As Integer           '数据库中批号定义长度
Private mstr外观 As String                  '用来记录默认外观的值
Private mstr验收结论 As String                  '用来记录默认验收结论
Private mbln提示 As Boolean                 '在药品选择器中选择的药品与界面中已有数据的比较看是否重复，对于重复的数据只提示一次，true 已经提示了，false还没有提示
Private mrs分段加成 As ADODB.Recordset      '分段加成集合
'Private mbln时价取上次售价 As Boolean        '时价药品直接去上次售价
Private mlng包装系数 As Long                '记录包装系数
Private mbln提示方式 As Boolean             '提示方式 true-只提示一次，false-连续提示
Private mbln效期提示 As Boolean             '是否提示失效期的药品,主要用于在加载单据时繁琐的过期药品提示。true-提示;false-不提示

Private marrFrom As Variant                   '纪录用户恢复窗体表列格宽度
Private marrInitGrid As Variant                '纪录初始化窗体表列格宽度

Private mblnEnter As Boolean                '是否进入单元格
Private mstr审核日期 As String              '审核日期

Private mbln修改批发价 As Boolean           '允许修改批发价
Private mdbl加价率 As Double
Private mblnUpdate As Boolean               '表示是否已根据最新价格更新单据内容
Private mbln退货 As Boolean                 '表示是否是退货单
Private mbln招标药品可选择非中标单位入库 As Boolean      '本地参数控制
Private mstr存储库房提示 As String
Private mbln所有药品无存储库房 As Boolean
Private mint取上次采购价方式 As Integer     '0-优先从药品库存取;1-优先从药品规格取
Private mbln供应商校验 As Boolean

Private mbln加价率 As Boolean               '时价药品是否必须输入加价率
Private mint时价入库售价加成方式 As Integer '系统参数：时价药品外购入库时售价计算方式：0－按折扣后的采购价计算售价;1－按折扣前的采购价计算售价。
Private mint时价分段加成方式 As Integer     ' 0-不按分段加成（默认） 1-按分段加成
Private mint库存检查 As Integer             '表示药品出库时是否进行库存检查：0-不检查;1-检查，不足提醒；2-检查，不足禁止

Private mbln允许手工输入加成率 As Boolean
Private mstrColumn_UnSelected As String     '记录哪些列被设置为不显示
Public RecReturn As Recordset
Private mblnViewCost As Boolean             '查看成本价 true-可以查看成本价 false-不可以查看成本价

Private recSort As ADODB.Recordset          '按药品ID排序的专用记录集

Private mstrTime_Start As String                      '进入单据编辑界面时，待编辑单据的最大修改时间
Private mstrTime_End As String                        '此刻该编辑单据的最大修改时间

Private mlng入库库房 As Long
Private mintUnit As Integer                 '单位系数：1-售价;2-门诊;3-住院;4-药库

Private mstrControlItem As String           '核查、审核、财务审核环节允许修改的项目

Private mintLastCol As Integer              '用户的列设置中的最后可见列的列号

Private mintDrugNameShow As Integer         '药品显示：0－显示编码和名称；1－仅显示编码；2－仅显示名称

Private mblnMSH_GetFocus As Boolean         '控制只一次提示
Private mlng生产商长度 As Long                 '生产商字段长度
Private mlng原产地长度 As Long                 '原产地字段长度

Private mbln取目录中产地信息 As Boolean

Private Const MStrCaption As String = "药品外购入库管理"

Private Enum 环节

    核查 = 1
    审核 = 2
    财务审核 = 3
End Enum

Private mblnLoad As Boolean              '记录是否执行完成Form_Load事件

'从参数表中取药品价格、数量、金额小数位数
Private mintCostDigit As Integer        '成本价小数位数
Private mintPriceDigit As Integer       '售价小数位数
Private mintNumberDigit As Integer      '数量小数位数
Private mintMoneyDigit As Integer       '金额小数位数

Private Const mconint售价单位 As Integer = 1
Private Const mconint门诊单位 As Integer = 2
Private Const mconint住院单位 As Integer = 3
Private Const mconint药库单位 As Integer = 4


'=========================================================================================
Private mconIntCol行号 As Integer
Private mconIntCol药名 As Integer
Private mconIntCol商品名 As Integer
Private mconIntCol来源 As Integer
Private mconIntCol基本药物 As Integer
Private mconIntCol序号 As Integer
Private mconIntCol规格 As Integer
Private mconIntCol药价级别 As Integer
Private mconIntCol原生产商 As Integer
Private mconIntCol原销期 As Integer
Private mconIntCol比例系数 As Integer
Private mconintcol简码 As Integer
Private mconIntCol产地 As Integer
Private mconIntCol原产地 As Integer
Private mconIntCol单位 As Integer
Private mconIntCol批号 As Integer
Private mconIntCol生产日期 As Integer
Private mconIntCol效期 As Integer
Private mconIntCol数量 As Integer
Private mconIntCol冲销数量 As Integer
Private mconIntCol批次 As Integer
Private mconIntCol指导批发价 As Integer
Private mconIntCol扣率 As Integer
Private mconIntCol成本价 As Integer
Private mconIntCol成本金额 As Integer
Private mconIntCol售价 As Integer
Private mconIntCol售价金额 As Integer
Private mconintCol差价 As Integer
Private mconintCol零售价 As Integer
Private mconintCol零售单位 As Integer
Private mconintCol零售金额 As Integer
Private mconintCol零售差价 As Integer
Private mconIntCol批准文号 As Integer
Private mconIntCol外观 As Integer
Private mconIntCol验收结论 As Integer
Private mconintcol产品合格证 As Integer
Private mconintcol随货单号 As Integer
Private mconintcol发票号 As Integer
Private mconintcol发票代码 As Integer
Private mconIntCol发票日期 As Integer
Private mconintcol发票金额 As Integer
Private mconIntCol采购价 As Integer
Private mconIntCol分批属性 As Integer
Private mconIntCol是否新行 As Integer
Private mconIntcol加成率 As Integer
Private mconIntCol药品编码和名称 As Integer
Private mconIntCol药品编码 As Integer
Private mconIntCol药品名称 As Integer
Private mconIntCol付款标志 As Integer
Private mconIntCol计划id As Integer
Private mconintcol随货日期 As Integer
Private Const mconIntColS As Integer = 52
'=========================================================================================



Private Function CheckQualifications(ByVal strInput As String) As Boolean
    '校验供应商信息和资质效期
    'strInput：字符串时为名称；数字时为ID
    Dim rsTmp As ADODB.Recordset
    Dim strMsgInfo As String
    Dim strMsgDate As String
    Dim dateCurrent As Date
    Dim strMsg As String
    
    Dim intCheckType As Integer
    Dim arrColumn
    Dim strCheck As String
    Dim strCheck_供应商 As String
    Dim n As Integer
    Dim strTmp As String
    
    On Error GoTo errHandle
    If strInput = "" Then
        CheckQualifications = True
        Exit Function
    End If
        
    '资质校验项目和方式的保存格式：校验方式|类别1,项目1,是否校验;类别1,项目2,是否校验;类别2,项目1,是否校验;类别2,项目2....
    strCheck = zlDataBase.GetPara("资质校验", glngSys, 1300, "")
    
    '保存的参数格式不正确时退出
    If InStr(1, strCheck, "|") = 0 Then
        CheckQualifications = True
        Exit Function
    End If

    '取校验方式：0-不检查；1－提醒；2－禁止
    intCheckType = Val(Mid(strCheck, 1, InStr(1, strCheck, "|") - 1))
    
    '不检查时退出
    If intCheckType = 0 Then
        CheckQualifications = True
        Exit Function
    End If

    '取校验内容：
    strCheck = Mid(strCheck, InStr(1, strCheck, "|") + 1)

    If strCheck = "" Then
        CheckQualifications = True
        Exit Function
    End If

    '分别取卫材，生产商，供应商需要校验的内容
    strCheck = strCheck & ";"
    arrColumn = Split(strCheck, ";")
    For n = 0 To UBound(arrColumn)
        If arrColumn(n) <> "" Then
'            If Split(arrColumn(n), ",")(0) = "卫材" And Split(arrColumn(n), ",")(2) = 1 Then
'                strCheck_卫材 = IIf(strCheck_卫材 = "", "", strCheck_卫材 & ";") & Split(arrColumn(n), ",")(1)
'            End If
'
'            If Split(arrColumn(n), ",")(0) = "卫材生产商" And Split(arrColumn(n), ",")(2) = 1 Then
'                strCheck_生产商 = IIf(strCheck_生产商 = "", "", strCheck_生产商 & ";") & Split(arrColumn(n), ",")(1)
'            End If

            If Split(arrColumn(n), ",")(0) = "药品供应商" And Split(arrColumn(n), ",")(2) = 1 Then
                strCheck_供应商 = IIf(strCheck_供应商 = "", "", strCheck_供应商 & ";") & Split(arrColumn(n), ",")(1)
            End If
        End If
    Next
    
    '无校验内容时退出
    If strCheck_供应商 = "" Then
        CheckQualifications = True
        Exit Function
    End If
    
    dateCurrent = CDate(Format(Sys.Currentdate, "yyyy-mm-dd"))
    
    gstrSQL = "Select ('[' || 编码 || ']' || 名称) AS 供应商, 税务登记号, 许可证号, 执照号, 授权号, 质量认证号, 质量认证日期, 药监局备案号, 药监局备案日期, 许可证效期, 执照效期, 授权期 " & _
              "From 供应商 " & _
              "Where (撤档时间 Is Null Or 撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And ID = [1] "
    Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, "供应商信息", Val(strInput))
    
    strTmp = ""
    
    If Not rsTmp.EOF Then
        If nvl(rsTmp!税务登记号) = "" And InStr(strCheck_供应商, "税务登记号") > 0 Then
            strTmp = rsTmp!供应商 & "：" & "无税务登记号"
        End If
        
        If nvl(rsTmp!许可证号) = "" And InStr(strCheck_供应商, "许可证号") > 0 Then
            strTmp = IIf(strTmp = "", rsTmp!供应商 & "：", strTmp & ",") & "无许可证号"
        End If
        
        If nvl(rsTmp!执照号) = "" And InStr(strCheck_供应商, "执照号") > 0 Then
            strTmp = IIf(strTmp = "", rsTmp!供应商 & "：", strTmp & ",") & "无执照号"
        End If
        
        If nvl(rsTmp!授权号) = "" And InStr(strCheck_供应商, "授权号") > 0 Then
            strTmp = IIf(strTmp = "", rsTmp!供应商 & "：", strTmp & ",") & "无授权号"
        End If
        
        If nvl(rsTmp!质量认证号) = "" And InStr(strCheck_供应商, "质量认证号") > 0 Then
            strTmp = IIf(strTmp = "", rsTmp!供应商 & "：", strTmp & ",") & "无质量认证号"
        End If
        
        If nvl(rsTmp!质量认证日期) <> "" Then
            If DateDiff("d", rsTmp!质量认证日期, dateCurrent) > 0 And InStr(strCheck_供应商, "质量认证日期") > 0 Then
                strTmp = IIf(strTmp = "", rsTmp!供应商 & "：", strTmp & ",") & "质量认证号已过期"
            End If
        End If
        
        If nvl(rsTmp!药监局备案号) = "" And InStr(strCheck_供应商, "药监局备案号") > 0 Then
            strTmp = IIf(strTmp = "", rsTmp!供应商 & "：", strTmp & ",") & "无药监局备案号"
        End If
        
        If nvl(rsTmp!药监局备案日期) <> "" Then
            If DateDiff("d", rsTmp!药监局备案日期, dateCurrent) > 0 And InStr(strCheck_供应商, "药监局备案日期") > 0 Then
                strTmp = IIf(strTmp = "", rsTmp!供应商 & "：", strTmp & ",") & "药监局备案号已过期"
            End If
        End If
        
        If nvl(rsTmp!许可证效期) <> "" Then
            If DateDiff("d", rsTmp!许可证效期, dateCurrent) > 0 And InStr(strCheck_供应商, "许可证效期") > 0 Then
                strTmp = IIf(strTmp = "", rsTmp!供应商 & "：", strTmp & ",") & "许可证已过期"
            End If
        End If
        
        If nvl(rsTmp!执照效期) <> "" Then
            If DateDiff("d", rsTmp!执照效期, dateCurrent) > 0 And InStr(strCheck_供应商, "执照效期") > 0 Then
                strTmp = IIf(strTmp = "", rsTmp!供应商 & "：", strTmp & ",") & "执照已过期"
            End If
        End If
        
        If nvl(rsTmp!授权期) <> "" Then
            If DateDiff("d", rsTmp!执照效期, dateCurrent) > 0 And InStr(strCheck_供应商, "授权期") > 0 Then
                strTmp = IIf(strTmp = "", rsTmp!供应商 & "：", strTmp & ",") & "授权已过期"
            End If
        End If
    End If
    
    '提示或禁止
    If strTmp <> "" Then
        If intCheckType = 1 Then
            If MsgBox("未通过资质校验，是否继续？" & vbCrLf & strTmp, vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
                CheckQualifications = True
                Exit Function
            Else
                Exit Function
            End If
        ElseIf intCheckType = 2 Then
            MsgBox "未通过资质校验，不能入库！" & vbCrLf & strTmp, vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    
    CheckQualifications = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Check人员部门(lngUserId As Long, lngDeptId As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select 1 From 部门人员 Where 人员id=[1] And 部门id=[2] "
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[判断用户是否是属于入库部门]", lngUserId, lngDeptId)
    Check人员部门 = (rsTemp.RecordCount > 0)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Check合同单位() As Boolean
    Dim n As Integer
    Dim rs As New ADODB.Recordset
    Dim strTmp As String
    
    On Error GoTo errHandle
    For n = 1 To mshBill.rows - 1
        If Trim(mshBill.TextMatrix(n, 0)) <> "" Then
            gstrSQL = "select nvl(合同单位id,0) 合同单位id from 药品规格 where 药品id=[1] "
            Set rs = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[判断是否是存在合同单位]", Val(mshBill.TextMatrix(n, 0)))
            
            If Not rs.EOF Then
                gstrSQL = "select id,名称 from 供应商 " & _
                          "where (站点 = [2] Or 站点 is Null) And 末级=1 And substr(类型,1,1) = '1' And (撤档时间 is null or 撤档时间=to_date('3000-01-01','YYYY-MM-DD')) " & _
                          "  and id=(select nvl(合同单位id,0) id from 药品规格 where 药品id=[1]) "
                Set rs = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[判断是否是在合同单位处采购]", Val(mshBill.TextMatrix(n, 0)), gstrNodeNo)
                
                If Not rs.EOF Then
                    If rs!id <> txtProvider.Tag Then
                        strTmp = strTmp & mshBill.TextMatrix(n, mconIntCol药名) & "[" & rs!名称 & "]" & vbCrLf
                    End If
                End If
            End If
        End If
    Next
    
    If strTmp <> "" Then
        MsgBox "该供药单位不是以下药品的合同单位：" & vbCrLf & strTmp, vbInformation, gstrSysName
        Exit Function
    End If
    
    Check合同单位 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Check存储库房() As Boolean
    Dim n As Integer
    Dim rs As New ADODB.Recordset
    Dim strTmp As String
        
    mbln所有药品无存储库房 = True
    On Error GoTo errHandle
    For n = 1 To mshBill.rows - 1
        If Trim(mshBill.TextMatrix(n, 0)) <> "" Then
            gstrSQL = "select 收费细目ID from 收费执行科室 where 收费细目ID=[1] and 执行科室ID=[2]  "
            Set rs = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[判断药品存储库房]", Val(mshBill.TextMatrix(n, 0)), cboEnterStock.ItemData(cboEnterStock.ListIndex))
            
            If rs.RecordCount = 0 Then
                 strTmp = strTmp & mshBill.TextMatrix(n, mconIntCol药名) & vbCrLf
            Else
                mbln所有药品无存储库房 = False
            End If
        End If
    Next
    
    If strTmp <> "" Then
        If mbln所有药品无存储库房 Then
            mstr存储库房提示 = "本次入库药品没有设置存储库房，将不能移库到[" & cboEnterStock.Text & "]。"
        Else
            mstr存储库房提示 = "以下药品没有设置存储库房，将不能移库到[" & cboEnterStock.Text & "] ：" & vbCrLf & strTmp & vbCrLf & "其余药品可以导入移库。"
        End If
        Check存储库房 = False
    Else
        Check存储库房 = True
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetNextEnableCol(ByVal intCurrCol As Integer) As Integer
    '返回下一个可见并可用的列号
    Dim n As Integer
    Dim intNextCol As Integer
    
    If mshBill.TextMatrix(mshBill.Row, 0) <> "" Then
        If intCurrCol > mshBill.Cols Or intCurrCol + 1 >= mintLastCol Then
            If mshBill.Row = mshBill.rows - 1 Then
                mshBill.rows = mshBill.rows + 1
            End If
            
            mshBill.Row = mshBill.Row + 1
            GetNextEnableCol = 2
            Exit Function
        End If
        
        With mshBill
            For n = intCurrCol + 1 To .Cols - 1
                If .ColWidth(n) > 0 And .ColData(n) <> 5 Then
                    intNextCol = n
                    Exit For
                End If
            Next
        End With
        
        GetNextEnableCol = IIf(intNextCol = 0, mintLastCol, intNextCol)
    End If
End Function
Private Sub GetSysParm()
    Dim int环节 As Integer
    
    'mint编辑状态：1.新增；2、修改；3、验收；4、查看；5、修改发票；6、冲销；
    '7、财务审核（冲销、产生新单据并审核；已付款的单据不允许财
    '务审核，同样，财务审核后的单据不允许冲销）;8-药库退货;9-核查

    If mint编辑状态 = 9 Then
        int环节 = 环节.核查
    ElseIf mint编辑状态 = 3 Then
        int环节 = 环节.审核
    ElseIf mint编辑状态 = 7 Then
        int环节 = 环节.财务审核
    End If
    
    If int环节 > 0 Then
        mstrControlItem = "," & GetControlItem(1, int环节) & ","
    End If
End Sub

Private Sub Get药品分批属性(ByVal intBillRow As Integer)
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim int分批属性 As Integer      '0-不分批;1-分批
    Dim int药库分批 As Integer      '0-不分批;1-分批
    Dim int药房分批 As Integer      '0-不分批;1-分批
    Dim bln是否具有药房性质 As Boolean  'True-具有药房性质;False-不具有药房性质
    
    If Val(mshBill.TextMatrix(intBillRow, 0)) = 0 Then Exit Sub
    On Error GoTo errHandle
    strSQL = "SELECT NVL(药库分批, 0) 药库分批,NVL(药房分批, 0) 药房分批 " & _
            " From 药品规格 WHERE 药品ID = [1] "
    Set rsTemp = zlDataBase.OpenSQLRecord(strSQL, "取药品库房分批属性", Val(mshBill.TextMatrix(intBillRow, 0)))
    
    If rsTemp.RecordCount > 0 Then
        int药库分批 = rsTemp!药库分批
        int药房分批 = rsTemp!药房分批
    End If
    
    If int药房分批 = 1 Then     '如果药房分批，则分批属性为1
        int分批属性 = 1
    Else
        If int药库分批 = 1 Then
            strSQL = "SELECT 部门ID From 部门性质说明 " & _
                    " WHERE ((工作性质 LIKE '%药房') OR (工作性质 LIKE '制剂室')) AND 部门ID = [1] "
            Set rsTemp = zlDataBase.OpenSQLRecord(strSQL, "取部门性质", cboStock.ItemData(Me.cboStock.ListIndex))
            
            bln是否具有药房性质 = (rsTemp.RecordCount > 0)
                    
            If bln是否具有药房性质 Then
                int分批属性 = 0
            Else
                int分批属性 = 1
            End If
        End If
    End If
    
    mshBill.TextMatrix(intBillRow, mconIntCol分批属性) = int分批属性
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Get摘要(ByVal strNo As String, ByVal int编辑状态 As Integer) As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    Select Case int编辑状态
        Case 6          '冲销(取最后一次冲销的摘要)
            gstrSQL = "Select 摘要 From 药品收发记录 Where 单据=1 And No=[1] Order By 审核日期 Desc "
        Case 5, 7       '修改发票、财务审核
            gstrSQL = "Select 摘要 From 药品收发记录 Where 单据 = 1 And NO = [1] And (Mod(记录状态, 3) = 0 Or 记录状态 = 1) "
    End Select
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "取摘要信息", strNo)
    
    If Not rsTemp.EOF Then
        Get摘要 = nvl(rsTemp!摘要)
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub Refresh付款标志()
    Dim rs As New ADODB.Recordset
    Dim n As Integer
    
    On Error GoTo errHandle
    For n = 1 To mshBill.rows - 1
        If mshBill.TextMatrix(n, mconIntCol序号) <> "" Then
            gstrSQL = "Select Nvl(Max(付款序号), 0) 付款序号 From 应付记录 " & _
                " where 收发id=(Select Id From 药品收发记录 Where 单据=1 And No=[1] And (Mod(记录状态,3)=0 Or 记录状态=1) " & _
                " And 序号=[2]) "
            Set rs = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[取付款序号]", txtNO.Text, Val(mshBill.TextMatrix(n, mconIntCol序号)))
            
            If rs.EOF Then
                mshBill.RowData(n) = 0
            Else
                mshBill.RowData(n) = rs!付款序号
            End If
        End If
    Next
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Sub


'检查数据依赖性
Private Function GetDepend() As Boolean
    Dim rsDepend As New Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    GetDepend = False
    strSQL = "SELECT B.Id " _
            & "FROM 药品单据性质 A, 药品入出类别 B " _
            & "Where A.类别id = B.ID AND A.单据 = 1 and rownum=1 "
    Set rsDepend = zlDataBase.OpenSQLRecord(strSQL, MStrCaption)
    If rsDepend.EOF Then
        MsgBox "没有设置药品外购入库的入出类别，请检查药品入出分类！", vbInformation + vbOKOnly, gstrSysName
        rsDepend.Close
        Exit Function
    End If
    rsDepend.Close
        
    strSQL = "Select (id) from 供应商 " & _
              "Where (站点 = [1] Or 站点 is Null) And (撤档时间=to_date('3000-01-01','yyyy-mm-dd') or 撤档时间 is null) " & _
              "  And (substr(类型,1,1)=1 Or Nvl(末级,0)=0) and rownum=1 "
    Set rsDepend = zlDataBase.OpenSQLRecord(strSQL, MStrCaption & "-供应商", gstrNodeNo)
    If rsDepend.EOF Then
        MsgBox "没有设置药品供药单位，请检查药品供药单位管理！", vbInformation + vbOKOnly, gstrSysName
        rsDepend.Close
        Exit Function
    End If
    rsDepend.Close
        
    GetDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function


Private Sub SetDrugName(ByVal intType As Integer)
    '药品名称显示：
    'intType：0－显示编码和名称；1－仅显示编码；2－仅显示名称
    Dim lngRow As Long
    
    If intType = mintDrugNameShow Then Exit Sub
    
    mintDrugNameShow = intType
    
    With mshBill
        For lngRow = 1 To .rows - 1
            If .TextMatrix(lngRow, mconIntCol药名) <> "" Then
                If mintDrugNameShow = 1 Then
                    .TextMatrix(lngRow, mconIntCol药名) = .TextMatrix(lngRow, mconIntCol药品编码)
                ElseIf mintDrugNameShow = 2 Then
                    .TextMatrix(lngRow, mconIntCol药名) = .TextMatrix(lngRow, mconIntCol药品名称)
                Else
                    .TextMatrix(lngRow, mconIntCol药名) = .TextMatrix(lngRow, mconIntCol药品编码和名称)
                End If
            End If
        Next
    End With
End Sub

Private Sub SetSortRecord()
    Dim n As Integer
    
    If mshBill.rows < 2 Then Exit Sub
    If mshBill.TextMatrix(1, 0) = "" Then Exit Sub
    
    Set recSort = New ADODB.Recordset
    With recSort
        If .State = 1 Then .Close
        .Fields.Append "行号", adDouble, 18, adFldIsNullable
        .Fields.Append "序号", adDouble, 18, adFldIsNullable
        .Fields.Append "药品ID", adDouble, 18, adFldIsNullable
        .Fields.Append "批次", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        For n = 1 To mshBill.rows - 1
            If mshBill.TextMatrix(n, 0) <> "" Then
                .AddNew
                !行号 = n
                !序号 = IIf(Val(mshBill.TextMatrix(n, mconIntCol序号)) = 0, n, Val(mshBill.TextMatrix(n, mconIntCol序号)))
                !药品ID = Val(mshBill.TextMatrix(n, 0))
                !批次 = Val(mshBill.TextMatrix(n, mconIntCol批次))
                
                .Update
            End If
        Next
        
    End With
End Sub

Private Sub Set时价分批药品零售价(ByVal intRow As Integer, ByVal dblPrice As Double)
    Dim Dbl数量 As Double

    With mshBill
        If .TextMatrix(intRow, mconIntCol原销期) = "" Then Exit Sub
        If Val(Split(.TextMatrix(intRow, mconIntCol原销期), "||")(2)) <> 1 Or Val(.TextMatrix(intRow, mconIntCol分批属性)) <> 1 Then Exit Sub
        
       .TextMatrix(intRow, mconintCol零售价) = zlStr.FormatEx(dblPrice, gtype_UserDrugDigits.Digit_零售价, , True) '零售价字段本来都是最小单位，因此按照最大位数进行显示
        
        If mint编辑状态 = 6 Then
            Dbl数量 = Val(.TextMatrix(intRow, mconIntCol冲销数量)) * Val(.TextMatrix(intRow, mconIntCol比例系数))
        Else
            Dbl数量 = Val(.TextMatrix(intRow, mconIntCol数量)) * Val(.TextMatrix(intRow, mconIntCol比例系数))
        End If
        
        If Val(.TextMatrix(intRow, mconIntCol成本价)) = Val(.TextMatrix(intRow, mconIntCol售价)) Then
            '通过技术手段单独处理零差价销售情况下零售价和售价不等的情况
            .TextMatrix(intRow, mconintCol零售金额) = .TextMatrix(intRow, mconIntCol售价金额)
        Else
            .TextMatrix(intRow, mconintCol零售金额) = zlStr.FormatEx(Dbl数量 * Val(.TextMatrix(intRow, mconintCol零售价)), mintMoneyDigit, , True)
        End If
        .TextMatrix(intRow, mconintCol零售差价) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol零售金额)) - Val(.TextMatrix(intRow, mconIntCol成本金额)), mintMoneyDigit, , True)
    End With
End Sub

Public Sub ShowCard(FrmMain As Form, ByVal str单据号 As String, ByVal int编辑状态 As Integer, Optional int记录状态 As Integer = 1, Optional BlnSuccess As Boolean = False)
    mblnSave = False
    mblnSuccess = False
    mstr单据号 = str单据号
    mint编辑状态 = int编辑状态
    
    mint记录状态 = int记录状态
    mblnSuccess = BlnSuccess
    mblnChange = False
    mbln效期提示 = False
    mintParallelRecord = 1
    mstrPrivs = GetPrivFunc(glngSys, 1300)

    mbln修改批发价 = (Val(zlDataBase.GetPara("修改采购限价", glngSys, 模块号.外购入库)) = 1)
    mbln招标药品可选择非中标单位入库 = (Val(zlDataBase.GetPara("招标药品可选择非中标单位入库", glngSys, 模块号.外购入库)) = 1)
    mint取上次采购价方式 = Val(zlDataBase.GetPara("取上次采购价方式", glngSys, 模块号.外购入库))
    mbln供应商校验 = (Val(zlDataBase.GetPara("校验供应商资质", glngSys, 模块号.外购入库)) = 1)
    mint时价分段加成方式 = gtype_UserSysParms.P181_药品入库按分段加成
    Set mfrmMain = FrmMain
    mblnEdit = False
    If mint编辑状态 = 1 Or mint编辑状态 = 8 Then
        mblnEdit = True

        txtNO.Locked = True
        txtNO.TabStop = True
        If Not GetDepend Then Exit Sub
    ElseIf mint编辑状态 = 2 Or mint编辑状态 = 7 Then
        mblnEdit = True
        If mint编辑状态 = 2 Then
            txtNO.Locked = True
            txtNO.TabStop = True
        End If
    ElseIf mint编辑状态 = 3 Then
        mblnEdit = False
        CmdSave.Caption = "审核(&V)"
        
        If Not mbln退货 Then
            Me.chk转入移库.Visible = True
            Me.cboEnterStock.Visible = True
        End If
    ElseIf mint编辑状态 = 4 Then
        mblnEdit = False
        CmdSave.Caption = "打印(&P)"
        If Not zlStr.IsHavePrivs(mstrPrivs, "单据打印") Then
            CmdSave.Visible = False
        Else
            CmdSave.Visible = True
        End If
    ElseIf mint编辑状态 = 5 Then
        mblnEdit = False
        
    ElseIf mint编辑状态 = 6 Then
        mblnEdit = False
        CmdSave.Caption = "冲销(&O)"
        cmdAllSel.Visible = True
        cmdAllCls.Visible = True
    End If
    
    
    LblTitle.Caption = GetUnitName & IIf(mint编辑状态 = 8, "药品退货单", LblTitle.Caption)
    mbln效期提示 = True
    Me.Show vbModal, FrmMain
    BlnSuccess = mblnSuccess
End Sub

Private Sub zlPrintBill_Check()
    Dim lng上次药品ID As Long
    
    If Val(zlDataBase.GetPara("审核打印", glngSys, 模块号.外购入库)) = 1 Then
        '打印
        If zlStr.IsHavePrivs(mstrPrivs, "单据打印") Then
            printbill
            
            If Val(zlDataBase.GetPara("打印药品条码", glngSys, 模块号.外购入库)) = 1 And zlStr.IsHavePrivs(mstrPrivs, "药品条码打印") Then
                '按药品ID顺序更新数据
                recSort.Sort = "药品id"
                recSort.MoveFirst
                '打印药品条码
                Do While Not recSort.EOF
                    If lng上次药品ID <> Val(recSort!药品ID) Then
                        ReportOpen gcnOracle, glngSys, "ZL1_INSIDE_1300_1", Me, "药品=" & Val(recSort!药品ID), 2
                        lng上次药品ID = recSort!药品ID
                    End If
                    recSort.MoveNext
                Loop
            End If
                
        End If
    End If
End Sub

Private Sub cboEnterStock_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cboEnterStock.ListCount = 0 Then Exit Sub
    If cboEnterStock.ListIndex >= 0 Then
        If Val(cboEnterStock.Tag) = cboEnterStock.ItemData(cboEnterStock.ListIndex) Then
            Exit Sub
        End If
    End If

    Dim rsEnterDept As ADODB.Recordset
    Dim vRect As RECT
    Dim blnCancel As Boolean
    Dim i As Long
        
    vRect = zlControl.GetControlRect(cboEnterStock.hWnd)
    gstrSQL = "Select Distinct a.Id, a.编码, a.名称 " & vbNewLine & _
              "From 部门性质说明 C, 部门性质分类 B, 部门表 A, " & vbNewLine & _
              "     (Select 对方库房id ID " & vbNewLine & _
              "       From 药品流向控制 " & vbNewLine & _
              "       Where 所在库房id = [1] And 流向 In (1, 3) " & vbNewLine & _
              "       Union " & vbNewLine & _
              "       Select 所在库房id ID From 药品流向控制 Where 对方库房id = [1] And 流向 In (2, 3)) D " & vbNewLine & _
              "Where c.工作性质 = b.名称 And b.编码 || '' In ('H', 'I', 'J', 'K', 'L', 'M', 'N') " & vbNewLine & _
              "    And a.Id = c.部门id And a.Id = d.Id " & vbNewLine & _
              "    And To_Char(a.撤档时间, 'yyyy-MM-dd') = '3000-01-01' " & vbNewLine & _
              "    And (A.编码 like [2] or A.名称 like [2] or A.简码 like [2] ) " & vbNewLine & _
              "Order By a.编码 "
    Set rsEnterDept = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, MStrCaption, False, "", "", _
            False, False, True, vRect.Left - 15, vRect.Top, 3000, blnCancel, False, False, _
            cboStock.ItemData(cboStock.ListIndex), _
            IIf(gstrMatchMethod = "0", "%", "") & UCase(Trim(cboEnterStock.Text)) & "%")
    If blnCancel = False Then
        If rsEnterDept Is Nothing Then Exit Sub
        If Not rsEnterDept.EOF Then
            For i = 0 To cboEnterStock.ListCount - 1
                If cboEnterStock.ItemData(i) = nvl(rsEnterDept!id, -1) Then
                    cboEnterStock.ListIndex = i
                    cboEnterStock.Tag = nvl(rsEnterDept!id, 0)
                    Exit For
                End If
            Next
        End If
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cboStock_Change()
    mblnChange = True
End Sub

Private Sub cboStock_Click()
    Dim rsDetail As New Recordset
    Dim bln中药库房 As Boolean
    Dim str库房性质 As String
    Dim str屏蔽列 As String
    
    On Error GoTo errHandle
    
    str库房性质 = ""
    
    If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
        gstrSQL = "Select 工作性质 From 部门性质说明 Where 部门id =[1]"
        Set rsDetail = zlDataBase.OpenSQLRecord(gstrSQL, "判断库房性质", cboStock.ItemData(cboStock.ListIndex))
        Do While Not rsDetail.EOF
            str库房性质 = str库房性质 & "," & rsDetail!工作性质
            rsDetail.MoveNext
        Loop
        If str库房性质 Like "*中药*" Or str库房性质 Like "*制剂室*" Then bln中药库房 = True

        str屏蔽列 = zlDataBase.GetPara("屏蔽列", glngSys, 模块号.外购入库)
        
        If InStr(1, "|" & str屏蔽列 & "|", "|原产地|") = 0 Then mshBill.ColWidth(mconIntCol原产地) = IIf(bln中药库房, 800, 0)
        
        If mblnLoad = True Then Call SetSelectorRS(IIf(mint编辑状态 = 8 Or mbln退货, 2, 1), MStrCaption, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , IIf(mint编辑状态 = 8 Or mbln退货, Val(txtProvider.Tag), 0))
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboStock_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboStock_Validate False
        OS.PressKey (vbKeyTab)
    End If
End Sub


Private Sub cboStock_Validate(Cancel As Boolean)
    Dim i As Integer
        
    With cboStock
        If .ListIndex <> mintcboIndex Then
            For i = 1 To mshBill.rows - 1
                If mshBill.TextMatrix(i, 0) <> "" Then
                    Exit For
                End If
            Next
            If i <> mshBill.rows Then
                If MsgBox("如果改变库房，有可能要改变相应药品的单位，且要清除现有单据内容，你是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    '处理药品单位改变
                    mintcboIndex = .ListIndex
                    mshBill.ClearBill
                    mint库存检查 = MediWork_GetCheckStockRule(cboStock.ItemData(cboStock.ListIndex))
                    
                    mlng入库库房 = Me.cboStock.ItemData(Me.cboStock.ListIndex)
                    Call GetDrugDigit(mlng入库库房, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
                Else
                    .ListIndex = mintcboIndex
                End If
            Else
                mintcboIndex = .ListIndex
            End If
        End If
        
        
    End With
End Sub

Private Sub chk转入移库_Click()
    If chk转入移库.Value = 1 Then
        cboEnterStock.Enabled = True
    Else
        cboEnterStock.Enabled = False
    End If
End Sub

Private Sub cmdAddProducer_Click()
    frmDrugProducer.Show 1, Me
End Sub

Private Sub cmdAllCls_Click()
    Dim intRow As Integer
    
    With mshBill
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                .TextMatrix(intRow, mconIntCol冲销数量) = zlStr.FormatEx(0, mintNumberDigit, , True)
                .TextMatrix(intRow, mconIntCol成本金额) = zlStr.FormatEx(0, mintMoneyDigit, , True)
                .TextMatrix(intRow, mconIntCol售价金额) = zlStr.FormatEx(0, mintMoneyDigit, , True)
                .TextMatrix(intRow, mconintCol差价) = zlStr.FormatEx(0, mintMoneyDigit, , True)
                If Trim(.TextMatrix(intRow, mconintcol发票号)) <> "" Then
                    .TextMatrix(intRow, mconintcol发票金额) = zlStr.FormatEx(0, mintMoneyDigit, , True)
                End If
                
                Call Set时价分批药品零售价(intRow, Val(.TextMatrix(intRow, mconintCol零售价)))
            End If
        Next
    End With
    mblnChange = False
End Sub

Private Sub cmdAllSel_Click()
    Dim rsDrug As New Recordset
    Dim intRow As Integer
    
    On Error GoTo errHandle
    For intRow = 1 To mshBill.rows - 1
        If (gtype_UserSysParms.P173_经过标记付款后才能进行付款管理 = 1 And mshBill.TextMatrix(intRow, mconIntCol付款标志) <> "是") Or gtype_UserSysParms.P173_经过标记付款后才能进行付款管理 <> 1 Then
            If mshBill.TextMatrix(intRow, 0) <> "" And mshBill.RowData(intRow) = 0 Then
                mshBill.TextMatrix(intRow, mconIntCol冲销数量) = mshBill.TextMatrix(intRow, mconIntCol数量)
                mshBill.TextMatrix(intRow, mconIntCol成本金额) = zlStr.FormatEx(mshBill.TextMatrix(intRow, mconIntCol数量) * mshBill.TextMatrix(intRow, mconIntCol成本价), mintMoneyDigit, , True)
                mshBill.TextMatrix(intRow, mconIntCol售价金额) = zlStr.FormatEx(mshBill.TextMatrix(intRow, mconIntCol数量) * mshBill.TextMatrix(intRow, mconIntCol售价), mintMoneyDigit, , True)
                mshBill.TextMatrix(intRow, mconintCol差价) = zlStr.FormatEx(mshBill.TextMatrix(intRow, mconIntCol售价金额) - mshBill.TextMatrix(intRow, mconIntCol成本金额), mintMoneyDigit, , True)
                  
                If Trim(mshBill.TextMatrix(intRow, mconintcol发票号)) <> "" Then
                    gstrSQL = "select sum(nvl(发票金额,0)) as 发票金额 " _
                        & " From 药品收发记录 x,(Select 收发id,发票金额 From 应付记录 Where 系统标识=1 And 记录性质=0) y " _
                        & " WHERE x.id=y.收发id(+) and X.NO=[1] AND 单据=1 " _
                        & " and x.药品id+0=[2] " _
                        & " and x.序号=[3] "
                    Set rsDrug = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, mstr单据号, Val(mshBill.TextMatrix(intRow, 0)), Val(mshBill.TextMatrix(intRow, mconIntCol序号)))
                    
                    If rsDrug.EOF Then
                        mshBill.TextMatrix(intRow, mconintcol发票金额) = mshBill.TextMatrix(intRow, mconIntCol成本金额)
                    Else
                        mshBill.TextMatrix(intRow, mconintcol发票金额) = zlStr.FormatEx(rsDrug.Fields(0), mintMoneyDigit, , True)
                    End If
                End If
                
                Call Set时价分批药品零售价(intRow, Val(mshBill.TextMatrix(intRow, mconintCol零售价)))
            End If
        End If
    Next
    mblnChange = True
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCopy_Click()
    Dim i As Integer
    
    With mshBill
        '1、都有发票号则不做第2步
        For i = 1 To .rows - 1
           If Trim(.TextMatrix(i, mconintcol发票号)) = "" Or .TextMatrix(i, 0) = "" Then Exit For
        Next
        
        If i = .rows - 1 Then Exit Sub
        
        '2、发票代码或发票日期为空，则提示
        If Trim(.TextMatrix(.Row, mconintcol发票代码)) = "" Or .TextMatrix(.Row, mconIntCol发票日期) = "" Then
            If MsgBox("发票代码或发票日期为空，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Else
            If MsgBox("是否将该行的发票信息批量复制到发票号为空的行？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        
        '3、复制
        For i = 1 To .rows - 1
            If i <> .Row And Trim(.TextMatrix(i, mconintcol发票号)) = "" And .TextMatrix(i, 0) <> "" Then    '不是编辑行且发票号为空的批量修改
                
                .TextMatrix(i, mconintcol发票号) = .TextMatrix(.Row, mconintcol发票号)
                .TextMatrix(i, mconintcol发票代码) = .TextMatrix(.Row, mconintcol发票代码)
                .TextMatrix(i, mconIntCol发票日期) = .TextMatrix(.Row, mconIntCol发票日期)
                If mint记录状态 = 1 Then .TextMatrix(i, mconintcol发票金额) = .TextMatrix(i, mconIntCol成本金额)
                
            End If
        Next
    End With
End Sub

Private Sub cmdALLDel_Click()
    Dim i As Integer
    
    With mshBill
        '1、都无发票号则不做第2步
        For i = 1 To .rows - 1
           If Trim(.TextMatrix(i, mconintcol发票号)) <> "" Or .TextMatrix(i, 0) = "" Then Exit For
        Next
        
        If i = .rows - 1 Then Exit Sub
    
        If MsgBox("该操作将清除所有行的发票相关数据，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            For i = 1 To .rows - 1
            
                If Trim(.TextMatrix(i, mconintcol发票号)) <> "" And .TextMatrix(i, 0) <> "" Then
                    .TextMatrix(i, mconintcol发票号) = ""
                    .TextMatrix(i, mconintcol发票代码) = ""
                    .TextMatrix(i, mconIntCol发票日期) = ""
                    .TextMatrix(i, mconintcol发票金额) = ""
                End If
                
            Next
            
            cmdCopy.Enabled = False
        End If
    End With
End Sub

'查找
Private Sub cmdFind_Click()
    If lblCode.Visible = False Then
        lblCode.Visible = True
        txtCode.Visible = True
        txtCode.SetFocus
    Else
        FindRow mshBill, mconIntCol药品编码和名称, txtCode.Text, True
        lblCode.Visible = False
        txtCode.Visible = False
        
    End If
    
    Form_Resize
End Sub

Private Sub cmdGetData_Click()
    If cboInputDate.Text <> "三个月内" Then
        If MsgBox("查询时间过长可能很慢，是否继续？", vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call getInputData
        End If
    End If
End Sub

Private Sub cmdGetInputCost_Click()
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim dblRowStation As Double
    Dim lngCurentRow As Long
    Dim lng药品ID As Long
    Dim dbl换算系数 As Double
    
    dblLeft = mshBill.Left + mshBill.MsfObj.CellLeft
    dblTop = mshBill.Top + mshBill.MsfObj.CellTop
    '通过控件高度获取行位置
    dblRowStation = mshBill.MsfObj.CellTop
    dblRowStation = dblRowStation / mshBill.MsfObj.CellHeight
    lngCurentRow = CLng(dblRowStation) 'Clng保证取到的行为整数
    
    If mshBill.TextMatrix(lngCurentRow, 0) <> "" Then
        cboInputDate.Clear
        '初始化下拉列表
        cboInputDate.AddItem "三个月内"
        cboInputDate.AddItem "半年内"
        cboInputDate.AddItem "一年内"
        cboInputDate.ListIndex = 0
        
        picInputCost.Visible = True
        vsfInputCost.SetFocus
        picInputCost.Top = dblTop
        picInputCost.Left = dblLeft
        
        lng药品ID = mshBill.TextMatrix(lngCurentRow, 0)
        dbl换算系数 = mshBill.TextMatrix(lngCurentRow, mconIntCol比例系数)
        picInputCost.Tag = lng药品ID
        cmdGetData.Tag = dbl换算系数
        lblDate.Tag = mintCostDigit
        vsfInputCost.Tag = lngCurentRow
        
        Call getInputData
    End If
End Sub

Private Sub getInputData()
    Dim dbBeginDate As Date
    Dim dbEndDate As Date
    Dim rsTemp As ADODB.Recordset
    
    If cboInputDate.Text = "三个月内" Then
        dbBeginDate = CDate(Format(DateAdd("M", -3, Date), "yyyy-mm-dd") & " 00:00:00")
        dbEndDate = CDate(Format(Date, "yyyy-mm-dd") & " 23:59:59")
    End If
    If cboInputDate.Text = "半年内" Then
        dbBeginDate = CDate(Format(DateAdd("M", -6, Date), "yyyy-mm-dd") & " 00:00:00")
        dbEndDate = CDate(Format(Date, "yyyy-mm-dd") & " 23:59:59")
    End If
    If cboInputDate.Text = "一年内" Then
        dbBeginDate = CDate(Format(DateAdd("yyyy", -1, Date), "yyyy-mm-dd") & " 00:00:00")
        dbEndDate = CDate(Format(Date, "yyyy-mm-dd") & " 23:59:59")
    End If
    
    gstrSQL = "Select a.No, a.审核日期, a.成本价, b.发票号, b.发票代码" & vbNewLine & _
                "From 药品收发记录 A, 应付记录 B" & vbNewLine & _
                "Where a.Id = b.收发id(+) And a.单据 = 1 And a.药品id + 0 = [1] And a.库房id+0=[2]  And (a.记录状态 = 1 Or Mod(a.记录状态, 3) = 0) And" & vbNewLine & _
                "      Nvl(a.费用id, 0) = 0 And nvl(a.发药方式,0)=0 And a.审核日期 Between [3] And [4]" & vbNewLine & _
                "Order By a.审核日期 Desc"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "入库信息查询", picInputCost.Tag, cboStock.ItemData(cboStock.ListIndex), dbBeginDate, dbEndDate)
    vsfInputCost.rows = 1
    Do While Not rsTemp.EOF
        With vsfInputCost
            .rows = .rows + 1
            .TextMatrix(.rows - 1, .ColIndex("no")) = rsTemp!NO
            .TextMatrix(.rows - 1, .ColIndex("入库时间")) = Format(rsTemp!审核日期, "yyyy-mm-dd hh:mm:ss")
            .TextMatrix(.rows - 1, .ColIndex("成本价")) = zlStr.FormatEx(rsTemp!成本价 * cmdGetData.Tag, lblDate.Tag, , True)
            .TextMatrix(.rows - 1, .ColIndex("发票号")) = IIf(IsNull(rsTemp!发票号), "", rsTemp!发票号)
            .TextMatrix(.rows - 1, .ColIndex("发票代码")) = IIf(IsNull(rsTemp!发票代码), "", rsTemp!发票代码)
            
            rsTemp.MoveNext
        End With
    Loop
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub cmdNo_Click()
    Dim mdbl加价率 As Double
    Dim dblTemp售价 As Double
    
    With mshBill
        mdbl加价率 = Val(Txt加价率.Tag)
                
        '重新计算零售价、差价
        If gtype_UserSysParms.P183_时价取上次售价 <> 1 Then
            .TextMatrix(.Row, mconIntCol售价) = zlStr.FormatEx(时价药品零售价(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol成本价)), mdbl加价率 / 100, Val(.TextMatrix(.Row, mconIntCol成本价)) * (1 + (mdbl加价率 / 100))), mintPriceDigit, , True)
        End If
        .TextMatrix(.Row, mconIntCol售价金额) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol售价)) * Val(.TextMatrix(.Row, mconIntCol数量)), mintMoneyDigit, , True)
        .TextMatrix(.Row, mconintCol差价) = zlStr.FormatEx(IIf(.TextMatrix(.Row, mconIntCol售价金额) = "", 0, .TextMatrix(.Row, mconIntCol售价金额)) - IIf(.TextMatrix(.Row, mconIntCol成本金额) = "", 0, .TextMatrix(.Row, mconIntCol成本金额)), mintMoneyDigit, , True)
    
        Call Set时价分批药品零售价(.Row, Val(.TextMatrix(.Row, mconIntCol售价)) / Val(.TextMatrix(.Row, mconIntCol比例系数)))
    End With
    PicInput.Visible = False
End Sub

Private Sub CmdNO_LostFocus()
    Call PicInput_LostFocus
End Sub

Private Sub cmdYes_Click()
    Dim dbl成本价 As Double
    
    If Val(Txt加价率) > 9900 Or Val(Txt加价率) < 0 Then
        MsgBox "请输入合法的加成率！（0-9900）", vbInformation, gstrSysName
        Txt加价率.SetFocus
        Exit Sub
    End If
    
    With mshBill
        '根据参数决定时价药品售价公式中成本价的算法
        dbl成本价 = IIf(mint时价入库售价加成方式 = 0, Val(.TextMatrix(.Row, mconIntCol成本价)), Val(.TextMatrix(.Row, mconIntCol采购价)))
                        
        '重新计算零售价、差价
        If gtype_UserSysParms.P183_时价取上次售价 <> 1 Then
            .TextMatrix(.Row, mconIntCol售价) = zlStr.FormatEx(时价药品零售价(Val(.TextMatrix(.Row, 0)), dbl成本价, Val(Txt加价率) / 100, dbl成本价 * (1 + (Val(Txt加价率) / 100))), mintPriceDigit, , True)
        End If
        .TextMatrix(.Row, mconIntCol售价金额) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol售价)) * Val(.TextMatrix(.Row, mconIntCol数量)), mintMoneyDigit, , True)
        .TextMatrix(.Row, mconintCol差价) = zlStr.FormatEx(IIf(.TextMatrix(.Row, mconIntCol售价金额) = "", 0, .TextMatrix(.Row, mconIntCol售价金额)) - IIf(.TextMatrix(.Row, mconIntCol成本金额) = "", 0, .TextMatrix(.Row, mconIntCol成本金额)), mintMoneyDigit, , True)
        .TextMatrix(.Row, mconIntcol加成率) = zlStr.FormatEx(Val(Txt加价率), 2) & "%"
        
        Call Set时价分批药品零售价(.Row, Val(.TextMatrix(.Row, mconIntCol售价)) / Val(.TextMatrix(.Row, mconIntCol比例系数)))
    End With
    
    PicInput.Visible = False
    mshBill.SetFocus
End Sub
Private Sub CmdYes_LostFocus()
    Call PicInput_LostFocus
End Sub

Private Sub Form_Activate()
    Select Case mintParallelRecord
        Case 1
            '正常
        Case 2
            '单据已被删除
            If mint编辑状态 = 5 Then
                MsgBox "该单据已被冲销，不能修改发票信息，请检查！", vbOKOnly, gstrSysName
            ElseIf mint编辑状态 = 6 Then
                MsgBox "该单据已没有可以冲销的药品，请检查！", vbOKOnly, gstrSysName
            Else
                MsgBox "该单据已被删除，请检查！", vbOKOnly, gstrSysName
            End If
            Unload Me
            Exit Sub
        Case 3
            '修改的单据已被审核
            MsgBox "该单据已被其他人审核，请检查！", vbOKOnly, gstrSysName
            Unload Me
            Exit Sub
        Case 4
            MsgBox "该单据已被其他人付过款，不能修改发票信息，请检查！", vbOKOnly, gstrSysName
            Unload Me
            Exit Sub
        Case 5
            MsgBox "该单据已被其他人付过款，不能进行财务审核！", vbOKOnly, gstrSysName
            Unload Me
            Exit Sub
        Case 6
            MsgBox "你不是[" & cboStock.Text & "]人员，不能进行财务审核。", vbOKOnly, gstrSysName
            Unload Me
            Exit Sub
        Case 7
            MsgBox "该单据已全部或者部分付款，不能冲销。", vbOKOnly, gstrSysName
            Unload Me
            Exit Sub
    End Select
    '初始化简码方式
    If (mint编辑状态 = 1 Or mint编辑状态 = 2) And gbytSimpleCodeTrans = 1 Then
        staThis.Panels("PY").Visible = True
        staThis.Panels("WB").Visible = True
        gint简码方式 = Val(zlDataBase.GetPara("简码方式", , , 0))    '默认拼音简码
        Logogram staThis, gint简码方式
    Else
        staThis.Panels("PY").Visible = False
        staThis.Panels("WB").Visible = False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sngLeft As Single, sngTop As Single
    If KeyCode = 70 Or KeyCode = 102 Then
        If Shift = vbCtrlMask Then   'Ctrl+F
            cmdFind_Click
        End If
    ElseIf KeyCode = vbKeyF3 Then
        FindRow mshBill, mconIntCol药名, txtCode.Text, False
    ElseIf KeyCode = vbKeyF4 Then
        '如果系统参数为真，则提示用户输入加价率
        If mbln加价率 And (mint编辑状态 = 1 Or mint编辑状态 = 2) Then
            If PicInput.Visible Then PicInput.SetFocus: Exit Sub
            If mshBill.TextMatrix(mshBill.Row, mconIntCol药名) = "" Then Exit Sub
            If Split(mshBill.TextMatrix(mshBill.Row, mconIntCol原销期), "||")(2) <> 1 Then Exit Sub
            sngLeft = Pic单据.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
            sngTop = Pic单据.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
            If sngTop + 1700 > Screen.Height Then
                sngTop = sngTop - mshBill.MsfObj.CellHeight - 1700
            End If
            
            With PicInput
                .Top = sngTop
                .Left = sngLeft
                .Visible = True
            End With
            Txt加价率 = "15.00000"
            With mshBill
                If Val(.TextMatrix(.Row, mconIntCol售价)) <> 0 And Val(.TextMatrix(.Row, mconIntCol成本价)) <> 0 Then
                    Txt加价率 = zlStr.FormatEx((Val(.TextMatrix(.Row, mconIntCol售价)) / Val(.TextMatrix(.Row, mconIntCol成本价)) - 1) * 100, 5, , True)
                End If
            End With
            Txt加价率.Tag = Txt加价率
            Txt加价率.SetFocus
        End If
    ElseIf KeyCode = vbKeyF7 Then
        If staThis.Panels("PY").Bevel = sbrRaised Then
            Logogram staThis, 0
        Else
            Logogram staThis, 1
        End If
    End If
End Sub

Private Sub cmdProvider_Click()
    Dim rsProvider As New Recordset
    Dim vRect As RECT
    
    vRect = zlControl.GetControlRect(txtProvider.hWnd)
    
    On Error GoTo errHandle
    gstrSQL = "Select id,上级ID,末级,编码,简码,名称 From 供应商 " & _
              "Where (站点 = [1] Or 站点 is Null) And (To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01' or 撤档时间 is null) " & _
              "  And (substr(类型,1,1)=1 Or Nvl(末级,0)=0) " & _
              "Start with 上级ID is null connect by prior ID =上级ID order by level,ID"
    'Set rsProvider = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption & "-药品供应商", gstrNodeNo)
    Set rsProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 1, "供药单位", True, "", "", False, False, _
        True, vRect.Left, vRect.Top, 300, False, False, True, gstrNodeNo)
    
    If rsProvider.State = 0 Then
        txtProvider.SetFocus
        Exit Sub
    End If
    
    If rsProvider.EOF Then
        rsProvider.Close
        Exit Sub
    End If
    
    Me.txtProvider.Tag = rsProvider!id
    Me.txtProvider = rsProvider!名称
    mblnChange = True
    
    mshBill.SetFocus
    
    If CheckQualifications(Val(txtProvider.Tag)) = False Then
        txtProvider.Text = ""
        txtProvider.Tag = "0"
        mblnChange = True
        Exit Sub
    End If
    
    mblnChange = True
    If Val(txtProvider.Tag) <> mlng供药单位ID And (mint编辑状态 = 8 Or mbln退货) Then
        mlng供药单位ID = Val(txtProvider.Tag)
        mshBill.ClearBill
        mshBill.TextMatrix(1, mconIntCol行号) = "1"
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'打印单据
Private Sub printbill()
    Dim int单位系数 As Integer
    Dim strNo As String
    
    Select Case mintUnit
        Case mconint售价单位
            int单位系数 = 4
        Case mconint门诊单位
            int单位系数 = 2
        Case mconint住院单位
            int单位系数 = 1
        Case mconint药库单位
            int单位系数 = 3
    End Select
    
    strNo = txtNO.Tag
    FrmBillPrint.ShowME Me, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1300", "zl8_bill_1300"), _
        mint记录状态, int单位系数, "1300", IIf(mint编辑状态 = 8 Or mbln退货, "药品退货单", "药品外购入库单"), strNo
End Sub

Private Sub CmdSave_Click()
    On Error GoTo ErrHand
    Dim strNewNO As String
    Dim BlnSuccess As Boolean, blnTrans As Boolean, bln退货单 As Boolean
    Dim str药品 As String
    Dim intLop As Integer
    Dim lng上次药品ID As Long
    
    '设置排序数据集
    Call SetSortRecord
    
    mstr审核日期 = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
    If mint编辑状态 = 4 Then    '查看
        '打印
        printbill
        '退出
        Unload Me
        Exit Sub
    End If
    
    '检查界面上药品进行预调价处理
    For intLop = 1 To Me.mshBill.rows - 1
        If mshBill.TextMatrix(intLop, 0) <> "" Then '有药品
            Call AutoAdjustPrice_ByID(Val(mshBill.TextMatrix(intLop, 0)))
        End If
    Next
    
    If mint编辑状态 = 9 Then    '核查
        mstrTime_End = GetBillInfo(1, mstr单据号)
        If mstrTime_End = "" Then
            MsgBox "该单据已经被其他操作员删除！", vbInformation, gstrSysName
            Exit Sub
        End If
   
        If mstrTime_End > mstrTime_Start Then
            MsgBox "该单据已经被其他操作员编辑，请退出后重试！", vbInformation, gstrSysName
            Exit Sub
        End If

        If Not SaveCard Then Exit Sub
        
        Unload Me
        Exit Sub
    End If
    
    If mint编辑状态 = 7 Then    '财务审核
        '先冲销，再新增单据并审核
        gcnOracle.BeginTrans
        blnTrans = True
        '产生新NO
        strNewNO = Sys.GetNextNo(21, Me.cboStock.ItemData(Me.cboStock.ListIndex))
        If IsNull(strNewNO) Then Exit Sub
        '产生未审核新单据
        BlnSuccess = SaveNewCard(strNewNO)
        '向财务审核记录表中插入数据
        If BlnSuccess Then BlnSuccess = SaveVerifyCard(strNewNO)
        '冲销原单据
        If BlnSuccess Then BlnSuccess = SaveStrike
        '审核新单据
        If BlnSuccess Then BlnSuccess = SaveCheck(strNewNO)
        
        '重新将提示方式设置为false
        mbln提示方式 = False
        
        If BlnSuccess Then
            gcnOracle.CommitTrans
        Else
            gcnOracle.RollbackTrans
            Exit Sub
        End If
        blnTrans = False
        Unload Me
        Exit Sub
    End If
    
    If mint编辑状态 = 3 Then        '审核
        Dim rsTemp As New ADODB.Recordset
        
        If ValidData = False Then Exit Sub '主要是检查计划单生成的入库单
        
        mstrTime_End = GetBillInfo(1, mstr单据号)
        If mstrTime_End = "" Then
            MsgBox "该单据已经被其他操作员删除！", vbInformation, gstrSysName
            Exit Sub
        End If
   
        If mstrTime_End > mstrTime_Start Then
            MsgBox "该单据已经被其他操作员编辑，请退出后重试！", vbInformation, gstrSysName
            Exit Sub
        End If
       
        If chk转入移库.Value = 1 Then
            If cboEnterStock.ListIndex < 0 Then
                MsgBox "要移库的部门不正确！", vbInformation, gstrSysName
                cboEnterStock.SetFocus
                Exit Sub
            End If
            If cboStock.ItemData(cboStock.ListIndex) = cboEnterStock.ItemData(cboEnterStock.ListIndex) Then
                MsgBox "移入部门与移出部门不能相同！", vbInformation, gstrSysName
                cboEnterStock.SetFocus
                Exit Sub
            End If
        End If
       
        'Modified by ZYB 2004-05-16 昆明处理
        '在审核时，新的价格已生效，因此需要将单据删除后重新产生
        '因为只影响库存和药品数据，对应付款相关数据无影响
        If Not 检查单价(1, txtNO, False) And Not mblnUpdate Then
            '以最新的价格更新单据体，退出的目的是让用户看一下最终的单据
            MsgBox "有记录未使用最新售价，程序将自动完成更新（售价、售价金额、差价），更新后请检查！", vbInformation, gstrSysName
            Call RefreshBill
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
        If Not 药品单据审核(Txt填制人.Caption) Then Exit Sub
        
        '零差价管理：检查是否存在不满足零差价的药品
        For intLop = 1 To mshBill.rows - 1
            If mshBill.TextMatrix(intLop, 0) <> "" And gtype_UserSysParms.P275_零差价管理模式 = 2 Then
                If IsPriceAdjustMod(Val(mshBill.TextMatrix(intLop, 0))) = True Then
                    If Val(mshBill.TextMatrix(intLop, mconIntCol成本价)) <> Val(mshBill.TextMatrix(intLop, mconIntCol售价)) Then
                        MsgBox "第" & intLop & "行药品已启用零差价管理，但入库单据售价和成本价不一致，不能进行业务，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        mshBill.Row = intLop
                        mshBill.MsfObj.TopRow = intLop
                        Exit Sub
                    End If
                End If
            End If
        Next
                
        '检查本单据是否为退货单（发药方式=1）
        gstrSQL = "Select nvl(发药方式,0) 退货 " & _
                  "From 药品收发记录 " & _
                  "Where 单据 =1 and NO=[1] AND ROWNUM<2 "
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[检查是否是退货单]", txtNO.Text)
        If Not rsTemp.EOF Then
            bln退货单 = (rsTemp!退货 = 1)
        End If
        
        '只有不是退货单时，才进行以下操作（因为允许审核时修改发票信息），否则直接审核
        If Not bln退货单 Then
            blnTrans = True
            gcnOracle.BeginTrans
            '如果审核时修改了单据，则重新生成单据保存
            If mblnChange Then
                If Not SaveCard(True) Then
                    gcnOracle.RollbackTrans: Exit Sub
                End If
            End If
            If Not SaveCheck() Then
                gcnOracle.RollbackTrans: Exit Sub
            End If
            
            '转入移库窗口
            If chk转入移库.Value = 1 And Me.cboEnterStock.ListIndex >= 0 Then
                If Check是否存在负数量 Then
                    If MsgBox("入库药品数量存在负数，不能使用导入移库的功能。确认入库，选择<是>；放弃审核，选择<否>。", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        gcnOracle.CommitTrans
                        
                        zlPrintBill_Check
                    Else
                        gcnOracle.RollbackTrans: Exit Sub
                    End If
                Else
                    If Not Check存储库房 Then
                        If mbln所有药品无存储库房 Then
                            gcnOracle.CommitTrans
                            
                            MsgBox mstr存储库房提示
                            
                            zlPrintBill_Check
                        Else
                            gcnOracle.CommitTrans
                        
                            MsgBox mstr存储库房提示
                            
                            zlPrintBill_Check
                            
                            frmTransferCard.ShowCard Me, txtNO.Text, 11, , BlnSuccess
                        End If
                    Else
                        gcnOracle.CommitTrans
                        
                        zlPrintBill_Check
                        
                        frmTransferCard.ShowCard Me, txtNO.Text, 11, , BlnSuccess
                    End If
                End If
            Else
                gcnOracle.CommitTrans
                
                zlPrintBill_Check
            End If
        Else
            If SaveCheck Then
                If Val(zlDataBase.GetPara("审核打印", glngSys, 模块号.外购入库)) = 1 Then
                    '打印
                    If zlStr.IsHavePrivs(mstrPrivs, "单据打印") Then
                        printbill
                        
                        If Val(zlDataBase.GetPara("打印药品条码", glngSys, 模块号.外购入库)) = 1 And zlStr.IsHavePrivs(mstrPrivs, "药品条码打印") Then
                            '按药品ID顺序更新数据
                            recSort.Sort = "药品id"
                            recSort.MoveFirst
                            '打印药品条码
                            Do While Not recSort.EOF
                                If lng上次药品ID <> Val(recSort!药品ID) Then
                                    ReportOpen gcnOracle, glngSys, "ZL1_INSIDE_1300_1", Me, "药品=" & Val(recSort!药品ID), 2
                                    lng上次药品ID = recSort!药品ID
                                End If
                                recSort.MoveNext
                            Loop
                        End If
                        
                    End If
                End If
            End If
        End If
        blnTrans = False
        Unload Me
        Exit Sub
    End If
            
    If mint编辑状态 = 5 Then      '修改发票信息
        If SaveRecipe = True Then
            Unload Me
        End If
        Exit Sub
    End If
    
    If mint编辑状态 = 6 Then
        If mblnChange = False Then
            MsgBox "请录入冲销数量！", vbInformation + vbOKOnly, gstrSysName
            Exit Sub
        End If
        
        If MsgBox("你确实要冲销单据吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            If SaveStrike = True Then
                Unload Me
            End If
        End If
        Exit Sub
    End If
    
    If mint编辑状态 = 8 Then
        If SaveRestore Then
            If Val(zlDataBase.GetPara("存盘打印", glngSys, 模块号.外购入库)) = 1 Then
                '打印
                If zlStr.IsHavePrivs(mstrPrivs, "单据打印") Then
                    printbill
                    
                    If Val(zlDataBase.GetPara("打印药品条码", glngSys, 模块号.外购入库)) = 1 And zlStr.IsHavePrivs(mstrPrivs, "药品条码打印") Then
                        '按药品ID顺序更新数据
                        recSort.Sort = "药品id"
                        recSort.MoveFirst
                        '打印药品条码
                        Do While Not recSort.EOF
                            If lng上次药品ID <> Val(recSort!药品ID) Then
                                ReportOpen gcnOracle, glngSys, "ZL1_INSIDE_1300_1", Me, "药品=" & Val(recSort!药品ID), 2
                                lng上次药品ID = recSort!药品ID
                            End If
                            recSort.MoveNext
                        Loop
                    End If
                    
                End If
            End If
            Unload Me
        End If
        Exit Sub
    End If
    
    If mint编辑状态 = 2 Then
        If Not 检查单价(1, txtNO, False) And Not mblnUpdate Then
            '以最新的价格更新单据体，退出的目的是让用户看一下最终的单据
            MsgBox "有记录未使用最新售价，程序将自动完成更新（售价、售价金额、差价），更新后请检查！", vbInformation, gstrSysName
            Call RefreshBill
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
    End If
    
    If mint编辑状态 = 1 Then '新增保存时，判断售价是否已经更新

        If 检查售价 Then
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
    End If
            
    If ValidData = False Then Exit Sub
    BlnSuccess = SaveCard
        
    If BlnSuccess = True Then
        If Val(zlDataBase.GetPara("存盘打印", glngSys, 模块号.外购入库)) = 1 Then
            '打印
            If zlStr.IsHavePrivs(mstrPrivs, "单据打印") Then
                printbill
                
                If Val(zlDataBase.GetPara("打印药品条码", glngSys, 模块号.外购入库)) = 1 And zlStr.IsHavePrivs(mstrPrivs, "药品条码打印") Then
                    '按药品ID顺序更新数据
                    recSort.Sort = "药品id"
                    recSort.MoveFirst
                    '打印药品条码
                    Do While Not recSort.EOF
                        If lng上次药品ID <> Val(recSort!药品ID) Then
                            ReportOpen gcnOracle, glngSys, "ZL1_INSIDE_1300_1", Me, "药品=" & Val(recSort!药品ID), 2
                            lng上次药品ID = recSort!药品ID
                        End If
                        recSort.MoveNext
                    Loop
                End If
                    
            End If
        End If
        If mint编辑状态 = 2 Then   '修改
            Unload Me
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    
    mblnSave = False
    mblnEdit = True
    mshBill.ClearBill
    Call RefreshRowNO(mshBill, mconIntCol行号, 1)
    SetEdit
    
'    txtProvider.Text = ""
'    txtProvider.Tag = "0"
    txt摘要.Text = ""
    txtProvider.SetFocus
    mblnChange = False
    If txtNO.Tag <> "" Then Me.staThis.Panels(2).Text = "上一张单据的NO号：" & txtNO.Tag
    Exit Sub
ErrHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
End Sub

Private Function 检查售价() As Boolean
    '功能：外购新增时，判断定价药品是否是最新售价，是则修改后提示
    Dim strMsg As String '保存提示信息
    Dim i As Integer, intSum As Integer, intPriceDigit As Integer
    Dim rsPrice As New ADODB.Recordset
    Dim Dbl数量 As Double, dbl成本价 As Double, dbl成本金额 As Double, dbl零售价 As Double, dbl零售金额 As Double, dbl差价 As Double
    
    On Error GoTo errHandle
    
    检查售价 = False
    
    With mshBill
        For i = 1 To .rows - 1
            If mshBill.TextMatrix(i, 0) <> "" Then
                
                If Val(Split(.TextMatrix(i, mconIntCol原销期), "||")(2)) = 0 Then '判断定价

                    dbl零售价 = zlStr.FormatEx(Get售价(False, Val(.TextMatrix(i, 0)), cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(i, mconIntCol批次))) * Val(.TextMatrix(i, mconIntCol比例系数)), mintPriceDigit)
                    
                    If .TextMatrix(i, mconIntCol售价) <> dbl零售价 Then
                        intSum = intSum + 1 '记录更新了几条数据
                        
                        dbl成本价 = Val(.TextMatrix(i, mconIntCol成本价))
                        Dbl数量 = Val(.TextMatrix(i, mconIntCol数量))
                        dbl成本金额 = dbl成本价 * Dbl数量
                        dbl零售金额 = dbl零售价 * Dbl数量
                        dbl差价 = dbl零售金额 - dbl成本金额
                        
                        '更新售价相关数据
                        .TextMatrix(i, mconIntCol售价) = zlStr.FormatEx(dbl零售价, mintPriceDigit, , True)
                        .TextMatrix(i, mconIntCol售价金额) = zlStr.FormatEx(dbl零售金额, mintMoneyDigit, , True)
                        .TextMatrix(i, mconintCol差价) = zlStr.FormatEx(dbl差价, mintMoneyDigit, , True)
                        .TextMatrix(i, mconIntcol加成率) = zlStr.FormatEx((Val(.TextMatrix(i, mconIntCol售价)) / dbl成本价 - 1) * 100, 2) & "%"
                        
                    End If
                End If
            End If
        Next
        
        If intSum > 0 Then
            MsgBox "有记录未使用最新售价，程序已自动完成更新（售价、售价金额、差价），更新后请检查！", vbInformation, gstrSysName
            检查售价 = True
        End If
    End With
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Load()
    Dim rsTemp As ADODB.Recordset
    Dim rsDetail As New Recordset
    Dim bln中药库房 As Boolean
    Dim str库房性质 As String
    Dim i As Integer, j As Integer
    Dim str屏蔽列 As String
    
    On Error GoTo errHandle
    mblnLoad = False
    marrFrom = Array()
    marrInitGrid = Array()
    mblnUpdate = False
    mintBatchNoLen = GetBatchNoLen()
    mbln加价率 = Get加价率()
'    mbln时价取上次售价 = IIf(Val(zldatabase.GetPara(183, 100)) = 0, False, True)
    mblnViewCost = zlStr.IsHavePrivs(mstrPrivs, "查看成本价")
    
    mbln取目录中产地信息 = gtype_UserSysParms.P294_优先取目录中产地信息 = 1
    
    txtNO = mstr单据号
    txtNO.Tag = txtNO
    Call Get时价药品直接确定售价
    Call GetSysParm
    Call GetDefineSize
    mblnEnter = True
        
    If glngModul = 1300 Then '外购入库单退货
        gstrSQL = "select 发药方式 from 药品收发记录 where no=[1] and 记录状态=[2] and 单据=1 and rownum=1 "
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "判断是否是退库单", mstr单据号, mint记录状态)
        
        If rsTemp.RecordCount > 0 Then
            mbln退货 = IIf(IsNull(rsTemp!发药方式), False, rsTemp!发药方式)
        Else
            mbln退货 = False
        End If
    Else
        mbln退货 = False
    End If
        
    Set mrs分段加成 = Nothing
    If mint时价分段加成方式 = 1 Then
        gstrSQL = "select 序号, 最低价, 最高价, 加成率, 差价额, 说明, 类型 from 药品加成方案 order by 序号"
        Set mrs分段加成 = zlDataBase.OpenSQLRecord(gstrSQL, "查询分段加成")
    End If
    mshBill.Value = Format(Sys.Currentdate, "YYYY-MM-DD")
    
    mintDrugNameShow = Int(Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品外购入库管理", "药品名称显示方式", 0)))
    If mintDrugNameShow > 2 Or mintDrugNameShow < 0 Then mintDrugNameShow = 0
    mnuColDrug.Item(mintDrugNameShow).Checked = True
    
    mlng入库库房 = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    
    '取入库单位的小数位数
    Call GetDrugDigit(mlng入库库房, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)

    Call initCard
    
    mstrTime_Start = GetBillInfo(1, mstr单据号)
    mint库存检查 = MediWork_GetCheckStockRule(cboStock.ItemData(cboStock.ListIndex))
    
    '只有中药类库房才显示"原产地"列
    str库房性质 = ""
    gstrSQL = "Select 工作性质 From 部门性质说明 Where 部门id =[1]"
    Set rsDetail = zlDataBase.OpenSQLRecord(gstrSQL, "判断库房性质", cboStock.ItemData(cboStock.ListIndex))
    Do While Not rsDetail.EOF
        str库房性质 = str库房性质 & "," & rsDetail!工作性质
        rsDetail.MoveNext
    Loop
    If str库房性质 Like "*中药*" Or str库房性质 Like "*制剂室*" Then bln中药库房 = True
    str屏蔽列 = zlDataBase.GetPara("屏蔽列", glngSys, 模块号.外购入库)
    If InStr(1, "|" & str屏蔽列 & "|", "|原产地|") = 0 Then mshBill.ColWidth(mconIntCol原产地) = IIf(bln中药库房, 800, 0)
    
    For i = 1 To mconIntColS - 1
        ReDim Preserve marrInitGrid(UBound(marrInitGrid) + 1)
        marrInitGrid(UBound(marrInitGrid)) = mshBill.TextMatrix(0, i) & "|" & mshBill.ColWidth(i)
    Next
    
    RestoreWinState Me, App.ProductName, MStrCaption
    
    For i = 1 To mconIntColS - 1
        ReDim Preserve marrFrom(UBound(marrFrom) + 1)
        marrFrom(UBound(marrFrom)) = mshBill.TextMatrix(0, i) & "|" & mshBill.ColWidth(i)
    Next
    
    For i = 0 To UBound(marrInitGrid)
        For j = 0 To UBound(marrFrom)
            If Split(marrInitGrid(i), "|")(0) = Split(marrFrom(j), "|")(0) And Split(marrInitGrid(i), "|")(1) * Split(marrFrom(j), "|")(1) = 0 Then
                mshBill.ColWidth(i + 1) = Split(marrInitGrid(i), "|")(1)
            End If
        Next
    Next

    If mint编辑状态 <> 6 Then
        If mshBill.ColWidth(mconIntCol冲销数量) > 0 Then
            mshBill.ColWidth(mconIntCol冲销数量) = 0
        End If
    Else
        If mshBill.ColWidth(mconIntCol冲销数量) = 0 Then
            mshBill.ColWidth(mconIntCol冲销数量) = 1000
        End If
    End If

    mshBill.ColWidth(mconIntCol序号) = 0
    mshBill.ColWidth(mconIntCol分批属性) = 0

    If mint编辑状态 = 8 Or mbln退货 = True Or mintUnit = mconint售价单位 Then
        mshBill.ColWidth(mconintCol零售价) = 0
        mshBill.ColWidth(mconintCol零售单位) = 0
        mshBill.ColWidth(mconintCol零售金额) = 0
        mshBill.ColWidth(mconintCol零售差价) = 0
    Else
        mshBill.ColWidth(mconintCol零售价) = 0
        mshBill.ColWidth(mconintCol零售单位) = 0
        mshBill.ColWidth(mconintCol零售金额) = 0
        mshBill.ColWidth(mconintCol零售差价) = 0
        
        If InStr(1, "|" & mstrColumn_UnSelected & "|", "|零售价|") = 0 Then mshBill.ColWidth(mconintCol零售价) = 1000
        If InStr(1, "|" & mstrColumn_UnSelected & "|", "|零售单位|") = 0 Then mshBill.ColWidth(mconintCol零售单位) = 1000
        If InStr(1, "|" & mstrColumn_UnSelected & "|", "|零售金额|") = 0 Then mshBill.ColWidth(mconintCol零售金额) = 1000
        If InStr(1, "|" & mstrColumn_UnSelected & "|", "|零售差价|") = 0 Then mshBill.ColWidth(mconintCol零售差价) = 1000
    End If
    
'    根据系统参数决定药房人员查看单据时，是否显示成本价
    If InStr(1, "|" & mstrColumn_UnSelected & "|", "|指导批发价|") = 0 Then mshBill.ColWidth(mconIntCol指导批发价) = IIf((mblnViewCost Or mint编辑状态 = 7), 1000, 0)
    If InStr(1, "|" & mstrColumn_UnSelected & "|", "|成本价|") = 0 Then mshBill.ColWidth(mconIntCol成本价) = IIf(mblnViewCost Or mint编辑状态 = 7, 1000, 0)
    If InStr(1, "|" & mstrColumn_UnSelected & "|", "|采购价|") = 0 Then mshBill.ColWidth(mconIntCol采购价) = IIf((mblnViewCost Or mint编辑状态 = 7), 1000, 0)
    If InStr(1, "|" & mstrColumn_UnSelected & "|", "|成本金额|") = 0 Then mshBill.ColWidth(mconIntCol成本金额) = IIf(mblnViewCost Or mint编辑状态 = 7, 900, 0)
    If InStr(1, "|" & mstrColumn_UnSelected & "|", "|差价|") = 0 Then mshBill.ColWidth(mconintCol差价) = IIf(mblnViewCost, 900, 0)
    If InStr(1, "|" & mstrColumn_UnSelected & "|", "|零售差价|") = 0 Then mshBill.ColWidth(mconintCol零售差价) = IIf(mblnViewCost, 1000, 0)
    
    '商品名列处理
    If gint药品名称显示 = 2 Then
        '显示商品名列
        mshBill.ColWidth(mconIntCol商品名) = IIf(mshBill.ColWidth(mconIntCol商品名) = 0, 2000, mshBill.ColWidth(mconIntCol商品名))
    Else
        '不单独显示商品名列
        mshBill.ColWidth(mconIntCol商品名) = 0
    End If
    
    mblnLoad = True
    If Check人员部门(UserInfo.用户ID, cboStock.ItemData(cboStock.ListIndex)) = False And mint编辑状态 = 7 Then
        mintParallelRecord = 6
        Exit Sub
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub initCard()
    Dim rsInitCard As New Recordset
    Dim strUnitQuantity As String
    Dim intRow As Integer
    Dim num包装系数 As String
    Dim strOrder As String, strCompare As String
    Dim i As Long, j As Long
    Dim rsEnterStock As New ADODB.Recordset
    Dim numUseAbleCount As Double
    Dim vardrug As Variant
    Dim str批次 As String
    Dim strArray As String
    Dim dbl成本价 As Double
    Dim intCostDigit As Integer        '成本价小数位数
    Dim intPriceDigit As Integer       '售价小数位数
    Dim intNumberDigit As Integer      '数量小数位数
    Dim intMoneyDigit As Integer       '金额小数位数
    Dim blnAllPay As Boolean
    Dim str药名 As String
    Dim strSqlOrder As String
    Dim rs As ADODB.Recordset
    
    blnAllPay = True
    On Error GoTo errHandle
    strOrder = zlDataBase.GetPara("排序", glngSys, 模块号.外购入库)
    strCompare = Mid(strOrder, 1, 1)
    
    strSqlOrder = "序号"
    
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
    
    strSqlOrder = strSqlOrder & IIf(Right(strOrder, 1) = "0", " ASC", " DESC")
    
    intCostDigit = mintCostDigit
    intPriceDigit = mintPriceDigit
    intNumberDigit = mintNumberDigit
    intMoneyDigit = mintMoneyDigit
        
    If mint编辑状态 <> 4 Then
        With mfrmMain.cboStock
            cboStock.Clear
            For i = 0 To .ListCount - 1
                cboStock.AddItem .List(i)
                cboStock.ItemData(cboStock.NewIndex) = .ItemData(i)
            Next
            mintcboIndex = .ListIndex
            cboStock.ListIndex = .ListIndex
            cboStock.Enabled = .Enabled
        End With
    End If
    
    If mint编辑状态 = 3 Then
        With cboEnterStock
            .Clear
            Set rsEnterStock = ReturnSQL(cboStock.ItemData(cboStock.ListIndex), MStrCaption & "[药品入库管理取移库库房]", True)
            
            Do While Not rsEnterStock.EOF
                .AddItem rsEnterStock.Fields(2)
                .ItemData(.NewIndex) = rsEnterStock.Fields(0)
                rsEnterStock.MoveNext
            Loop
    
            If .ListCount > 0 Then
                .ListIndex = 0
            End If
                                
        End With
    End If
    
    Select Case mint编辑状态
        Case 1, 8
            Txt填制人 = UserInfo.用户姓名
            Txt填制日期 = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
'            Txt修改人 = UserInfo.用户姓名
'            Txt修改日期 = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            initGrid
        Case 2, 3, 4, 5, 6, 7, 9
            initGrid
            If mint编辑状态 = 4 Then
                gstrSQL = "select b.id,b.名称 from 药品收发记录 a,部门表 b where a.库房id=b.id and A.单据 = 1 and a.no=[1] "
                Set rsInitCard = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, mstr单据号)
                
                If rsInitCard.EOF Then
                    mintParallelRecord = 2
                    Exit Sub
                End If
                
                With cboStock
                    .AddItem rsInitCard!名称
                    .ItemData(.NewIndex) = rsInitCard!id
                    .ListIndex = 0
                End With
                rsInitCard.Close
            End If
            
            Select Case mintUnit
                Case mconint售价单位
                    strUnitQuantity = "D.计算单位 AS 售价单位,D.计算单位 AS 单位, A.填写数量 AS 数量,'1' as 比例系数, "
                    num包装系数 = "1"
                Case mconint门诊单位
                    strUnitQuantity = "D.计算单位 AS 售价单位,B.门诊单位 AS 单位,(A.填写数量 / B.门诊包装) AS 数量,B.门诊包装 as 比例系数,"
                    num包装系数 = "B.门诊包装"
                Case mconint住院单位
                    strUnitQuantity = "D.计算单位 AS 售价单位,B.住院单位 AS 单位,(A.填写数量 / B.住院包装) AS 数量,B.住院包装 as 比例系数,"
                    num包装系数 = "B.住院包装"
                Case mconint药库单位
                    strUnitQuantity = "D.计算单位 AS 售价单位,B.药库单位 AS 单位,(A.填写数量 / B.药库包装) AS 数量,B.药库包装 as 比例系数,"
                    num包装系数 = "B.药库包装"
            End Select
            
            Select Case mint编辑状态
            Case 5, 7     '修改发票,财务审核
                If mint记录状态 = 1 Then
                    gstrSQL = "SELECT * FROM (SELECT DISTINCT A.药品ID,A.序号,'[' || D.编码 || ']' As 药品编码, D.名称 As 通用名, E.名称 As 商品名," & _
                        " B.药品来源,B.基本药物,D.规格,D.产地 AS 原生产商,A.产地, A.原产地,A.批号,NVL(A.批次,0) 批次," & _
                        " NVL(B.招标药品,0) 招标药品,NVL(B.差价让利比,0) 差价让利比,B.最大效期,A.效期," & strUnitQuantity & _
                        " nvl(A.单量,b.指导批发价)*" & num包装系数 & " AS 指导批发价 ,A.成本价*" & num包装系数 & " AS 采购价, " & _
                        " A.成本金额 AS 采购金额,D.是否变价,B.药房分批 药房分批核算," & _
                        " DECODE(A.扣率, NULL, 0, A.扣率) AS 扣率, A.零售价*" & num包装系数 & " AS 零售价,A.零售金额,A.差价, " & _
                        " A.批准文号,C.随货单号,C.随货日期, C.发票号 ,c.发票代码, C.发票日期, C.发票金额,A.供药单位ID,F.名称 AS 供应商, A.填制人,A.填制日期," & _
                        " A.修改人,A.修改日期,A.审核人,A.审核日期,A.库房ID,G.名称 AS 部门,NVL(C.付款序号,0) AS 付款序号,Nvl(A.发药方式,0) 退货,A.外观,A.验收结论," & _
                        " A.产品合格证,A.生产日期,A.配药人 As 核查人,A.配药日期 As 核查日期,B.药价级别, Nvl(A.用法, 0) As 金额差,A.频次 As 加成率,c.付款标志,a.计划id " & _
                        " FROM 药品收发记录 A, 药品规格 B,收费项目目录 D,收费项目别名 E ,应付记录 C,供应商 F,部门表 G " & _
                        " WHERE A.药品ID = B.药品ID AND B.药品ID=D.ID AND A.库房ID=G.ID" & _
                        " AND A.供药单位ID=F.ID AND SUBSTR(F.类型,1,1)=1" & _
                        " AND A.ID = C.收发ID(+) AND C.系统标识(+)=1 AND C.记录性质(+)=0 " & _
                        " AND B.药品ID=E.收费细目ID(+) AND E.性质(+)=3 " & _
                        " AND A.记录状态 =[2] " & _
                        " AND A.单据 = 1 AND A.NO =[1] " & _
                        " ) ORDER BY " & strSqlOrder
                Else
                    gstrSQL = "SELECT * FROM (SELECT DISTINCT A.药品ID,A.序号,'[' || D.编码 || ']' As 药品编码, D.名称 As 通用名, E.名称 As 商品名," & _
                        " B.药品来源,B.基本药物,D.规格,D.产地 AS 原生产商,A.产地,A.原产地,A.批号,A.批次," & _
                        " NVL(B.招标药品,0) 招标药品,NVL(B.差价让利比,0) 差价让利比,B.最大效期,A.效期," & strUnitQuantity & _
                        " nvl(A.单量,b.指导批发价)*" & num包装系数 & " AS 指导批发价 ,A.成本价*" & num包装系数 & " AS 采购价," & _
                        " A.成本金额 AS 采购金额,D.是否变价,B.药房分批 药房分批核算,  " & _
                        " DECODE(A.扣率, NULL, 0, A.扣率) AS 扣率, A.零售价*" & num包装系数 & " AS 零售价 ,A.零售金额,A.差价," & _
                        " A.批准文号, A.随货单号, A.随货日期, A.发票号,a.发票代码,A.发票日期,A.发票金额,A.供药单位ID,F.名称 AS 供应商, A.库房ID,G.名称 AS 部门,NVL(A.付款序号,0) AS 付款序号,A.退货,A.生产日期,A.配药人 As 核查人,A.配药日期 As 核查日期,B.药价级别,A.金额差,A.加成率,付款标志,a.计划id " & _
                        " FROM " & _
                        "     (SELECT X.NO, SUM(实际数量) AS 填写数量,SUM(成本金额) AS 成本金额,随货单号,随货日期,发票号,发票代码,发票日期," & _
                        "      X.药品ID,X.序号,X.产地, X.原产地,X.批号,NVL(X.批次,0) 批次,X.效期,X.扣率,X.成本价,X.零售价,X.单量," & _
                        "      X.供药单位ID,库房ID,NVL(Y.付款序号,0) AS 付款序号,Nvl(X.发药方式,0) As 退货,X.生产日期,X.批准文号,发票金额,X.配药人,X.配药日期,Sum(零售金额) 零售金额,Sum(差价) 差价,Sum(To_Number(Nvl(用法, 0))) As 金额差,频次 As 加成率,y.付款标志,x.计划id " & _
                        "      FROM 药品收发记录 X,(Select 序号,项目ID,随货单号,随货日期,发票号,发票代码,发票日期,付款序号,Sum(发票金额) 发票金额,付款标志  From 应付记录 " & _
                        "      Where 系统标识 = 1 And 记录性质 =0 And 入库单据号=[1] Group By 序号,项目ID,随货单号,随货日期,发票号,发票代码,发票日期,付款序号,付款标志) Y " & _
                        "      WHERE X.序号 = Y.序号(+) And X.药品ID=Y.项目ID(+) AND X.NO=[1] AND 单据=1 " & _
                        "      GROUP BY X.NO,X.药品ID,X.序号,X.产地,X.原产地,X.批号,NVL(X.批次,0),X.效期,X.扣率,X.成本价,X.零售价,X.单量," & _
                        "      X.供药单位ID,X.库房ID,随货单号,随货日期,发票号,发票代码,发票日期,NVL(Y.付款序号,0),Nvl(X.发药方式,0),X.生产日期,X.批准文号,发票金额,X.配药人,X.配药日期,X.频次,y.付款标志,x.计划id " & _
                        "      HAVING SUM(实际数量)<>0 ) A," & _
                        "      药品规格 B,收费项目别名 E ,收费项目目录 D,供应商 F,部门表 G " & _
                        " WHERE A.药品ID = B.药品ID AND B.药品ID=D.ID AND A.库房ID=G.ID" & _
                        " AND B.药品ID=E.收费细目ID(+) AND E.性质(+)=3 " & _
                        " AND A.供药单位ID=F.ID AND SUBSTR(F.类型,1,1)=1 " & _
                        " ) ORDER BY " & strSqlOrder
                End If
                
            Case 6      '冲销
                gstrSQL = "SELECT * FROM (SELECT DISTINCT A.药品ID,A.序号,'[' || D.编码 || ']' As 药品编码, D.名称 As 通用名, E.名称 As 商品名," & _
                    " B.药品来源,B.基本药物,D.规格,D.产地 AS 原生产商,A.产地,A.原产地,A.批号," & _
                    " NVL(B.招标药品,0) 招标药品,NVL(B.差价让利比,0) 差价让利比,B.最大效期,A.效期," & strUnitQuantity & _
                    " nvl(A.单量,b.指导批发价)*" & num包装系数 & " AS 指导批发价 ,A.成本价*" & num包装系数 & " AS 采购价," & _
                    " A.成本金额 AS 采购金额,D.是否变价,B.药房分批 药房分批核算,  " & _
                    " DECODE(A.扣率, NULL, 0, A.扣率) AS 扣率, A.零售价*" & num包装系数 & " AS 零售价 ,0 AS 零售金额,0 AS 差价,A.金额差, " & _
                    " A.批准文号,A.随货单号,A.随货日期, A.发票号,a.发票代码,A.发票日期,A.发票金额,A.供药单位ID,F.名称 AS 供应商, A.库房ID,G.名称 AS 部门,NVL(A.付款序号,0) AS 付款序号,A.退货,A.生产日期,A.批次,A.配药人 As 核查人,A.配药日期 As 核查日期,B.药价级别,A.加成率,a.付款标志,a.计划id,a.外观 " & _
                    " FROM " & _
                    "     (SELECT MIN(X.ID) AS ID, SUM(实际数量) AS 填写数量,SUM(成本金额) AS 成本金额,随货单号,随货日期,发票号,发票代码,发票日期,Sum(发票金额) As 发票金额," & _
                    "      X.药品ID,X.序号,X.产地, X.原产地,X.批号,X.效期,X.扣率,X.成本价,X.零售价,X.单量," & _
                    "      X.供药单位ID,库房ID,NVL(Y.付款序号,0) AS 付款序号,Nvl(X.发药方式,0) As 退货,X.生产日期,X.批准文号,NVL(X.批次,0) 批次,X.配药人,X.配药日期,Sum(To_Number(Nvl(用法, 0))) As 金额差,频次 As 加成率,y.付款标志,x.计划id,x.外观  " & _
                    "      FROM 药品收发记录 X,(SELECT 收发id,付款序号,随货单号,随货日期,发票号,发票代码,发票日期,发票金额,付款标志 FROM 应付记录 WHERE 系统标识=1 AND 记录性质=0) Y " & _
                    "      WHERE X.ID=Y.收发ID(+) AND X.NO=[1] AND 单据=1 " & _
                    "      GROUP BY X.药品ID,X.序号,X.产地,X.原产地,X.批号,X.效期,X.扣率,X.成本价,X.零售价,X.单量," & _
                    "      X.供药单位ID,X.库房ID,随货单号,随货日期,发票号,发票代码,发票日期,NVL(Y.付款序号,0),Nvl(X.发药方式,0),X.生产日期,X.批准文号,NVL(X.批次,0),X.配药人,X.配药日期,X.频次,付款标志,x.计划id,x.外观 " & _
                    "      HAVING SUM(实际数量)<>0 ) A," & _
                    "      药品规格 B,收费项目别名 E ,收费项目目录 D,供应商 F,部门表 G " & _
                    " WHERE A.药品ID = B.药品ID AND B.药品ID=D.ID AND A.库房ID=G.ID" & _
                    " AND B.药品ID=E.收费细目ID(+) AND E.性质(+)=3 " & _
                    " AND A.供药单位ID=F.ID AND SUBSTR(F.类型,1,1)=1 " & _
                    " ) ORDER BY " & strSqlOrder
            Case Else
                gstrSQL = "SELECT * FROM (SELECT DISTINCT A.药品ID,A.序号,'[' || D.编码 || ']' As 药品编码, D.名称 As 通用名, E.名称 As 商品名," & _
                    " B.药品来源,B.基本药物,D.规格,D.产地 AS 原生产商,A.产地, A.原产地,A.批号,NVL(A.批次,0) 批次," & _
                    " NVL(B.招标药品,0) 招标药品,NVL(B.差价让利比,0) 差价让利比,B.最大效期,A.效期," & strUnitQuantity & _
                    " nvl(A.单量,b.指导批发价)*" & num包装系数 & " AS 指导批发价 ,A.成本价*" & num包装系数 & " AS 采购价, " & _
                    " A.成本金额 AS 采购金额,D.是否变价,B.药房分批 药房分批核算," & _
                    " DECODE(A.扣率, NULL, 0, A.扣率) AS 扣率, A.零售价*" & num包装系数 & " AS 零售价,A.零售金额,A.差价, " & _
                    " A.批准文号,C.随货单号,C.随货日期, C.发票号 ,c.发票代码, C.发票日期, C.发票金额,A.供药单位ID,F.名称 AS 供应商, A.摘要,A.填制人,A.填制日期," & _
                    " A.修改人,A.修改日期,A.审核人,A.审核日期,A.库房ID,G.名称 AS 部门,NVL(C.付款序号,0) AS 付款序号,Nvl(A.发药方式,0) 退货,A.外观,A.验收结论,A.产品合格证," & _
                    " A.生产日期,A.配药人 As 核查人,A.配药日期 As 核查日期,B.药价级别, Nvl(A.用法, 0) As 金额差,A.频次 As 加成率,A.对方部门ID,a.计划id " & _
                    " FROM 药品收发记录 A, 药品规格 B,收费项目目录 D,收费项目别名 E ,应付记录 C,供应商 F,部门表 G " & _
                    " WHERE A.药品ID = B.药品ID AND B.药品ID=D.ID AND A.库房ID=G.ID" & _
                    " AND A.供药单位ID=F.ID AND SUBSTR(F.类型,1,1)=1" & _
                    " AND A.ID = C.收发ID(+) AND C.系统标识(+)=1 AND C.记录性质(+)=0 " & _
                    " AND B.药品ID=E.收费细目ID(+) AND E.性质(+)=3 " & _
                    " AND A.记录状态 =[2] " & _
                    " AND A.单据 = 1 AND A.NO = [1] " & _
                    " ) ORDER BY " & strSqlOrder
            End Select
             
            Set rsInitCard = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, mstr单据号, mint记录状态)
            
            If rsInitCard.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            
            Select Case mint编辑状态
                Case 2, 6, 9 '修改、冲销、核查
                    If mint编辑状态 = 2 Then
                        Txt填制人 = rsInitCard!填制人
                        Txt填制日期 = Format(rsInitCard!填制日期, "yyyy-mm-dd hh:mm:ss")
                        Txt修改人 = IIf(IsNull(rsInitCard!修改人), "", rsInitCard!修改人)
                        Txt修改日期 = IIf(IsNull(rsInitCard!修改日期), "", Format(rsInitCard!修改日期, "yyyy-mm-dd hh:mm:ss"))
                        Txt审核人 = ""
                        Txt审核日期 = ""
                    Else
                        Txt填制人 = UserInfo.用户姓名
                        Txt填制日期 = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
'                        Txt修改人 = UserInfo.用户姓名
'                        Txt修改日期 = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                        Txt审核人 = UserInfo.用户姓名
                        Txt审核日期 = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                        If mint编辑状态 = 9 Then
                            Txt填制人 = rsInitCard!填制人
                            Txt填制日期 = Format(rsInitCard!填制日期, "yyyy-mm-dd hh:mm:ss")
                            Txt修改人 = IIf(IsNull(rsInitCard!修改人), "", rsInitCard!修改人)
                            Txt修改日期 = IIf(IsNull(rsInitCard!修改日期), "", Format(rsInitCard!修改日期, "yyyy-mm-dd hh:mm:ss"))
                            Lbl审核人.Caption = "核查人"
                            Lbl审核日期.Caption = "核查日期"
                            lbl核查人.Visible = False
                            txt核查人.Visible = False
                            lbl核查日期.Visible = False
                            txt核查日期.Visible = False
                        End If
                    End If
                Case Else '3：验收；4：查看；5:修改发票；7：财务审核
                    If (mint编辑状态 = 5 Or mint编辑状态 = 7) And mint记录状态 <> 1 Then
                        Txt填制人 = UserInfo.用户姓名
                        Txt填制日期 = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                        Txt修改人 = UserInfo.用户姓名
                        Txt修改日期 = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                    Else
                        Txt填制人 = rsInitCard!填制人
                        Txt填制日期 = Format(rsInitCard!填制日期, "yyyy-mm-dd hh:mm:ss")
                        Txt修改人 = IIf(IsNull(rsInitCard!修改人), "", rsInitCard!修改人)
                        Txt修改日期 = IIf(IsNull(rsInitCard!修改日期), "", Format(rsInitCard!修改日期, "yyyy-mm-dd hh:mm:ss"))
                        Txt审核人 = IIf(IsNull(rsInitCard!审核人), "", rsInitCard!审核人)
                        Txt审核日期 = IIf(IsNull(rsInitCard!审核日期), "", Format(rsInitCard!审核日期, "yyyy-mm-dd hh:mm:ss"))
                    End If
            End Select
            
            '大医附二需求，自动锁定移库部门
            If mint编辑状态 = 3 Then        '审核
                If nvl(rsInitCard!对方部门id, 0) > 0 Then
                    chk转入移库.Tag = rsInitCard!对方部门id
                    chk转入移库.Value = 1
                    chk转入移库.Enabled = False
                    For i = 0 To cboEnterStock.ListCount
                        If Val(cboEnterStock.ItemData(i)) = rsInitCard!对方部门id Then
                            cboEnterStock.ListIndex = i
                            Exit For
                        End If
                    Next
                    cboEnterStock.Enabled = False
                End If
            ElseIf mint编辑状态 = 9 Then    '核查
                chk转入移库.Tag = nvl(rsInitCard!对方部门id)
            End If
            
            txt核查人.Caption = IIf(IsNull(rsInitCard!核查人), "", rsInitCard!核查人)
            txt核查日期.Caption = IIf(IsNull(rsInitCard!核查日期), "", Format(rsInitCard!核查日期, "yyyy-mm-dd hh:mm:ss"))

            
            txtProvider.Tag = rsInitCard!供药单位ID
            txtProvider.Text = rsInitCard!供应商
            
            If mint编辑状态 = 5 Or mint编辑状态 = 6 Or mint编辑状态 = 7 Then
                txt摘要.Text = Get摘要(mstr单据号, mint编辑状态)
            Else
                txt摘要.Text = IIf(IsNull(rsInitCard!摘要), "", rsInitCard!摘要)
            End If
            
            If (mint编辑状态 = 2 Or mint编辑状态 = 3) And Txt审核人 <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            If mint编辑状态 = 7 Then
                '只要有一笔付了款，就不允许进行财务审核
                With rsInitCard
                    Do While Not .EOF
                        If !付款序号 <> 0 Then
                            mintParallelRecord = 5        '已被其他人付款
                            Exit Sub
                        Else
                            '检查是否存在部分付款的情况
                            gstrSQL = "Select Nvl(Max(付款序号), 0) 付款序号 From 应付记录 " & _
                                " where 收发id=(Select Id From 药品收发记录 Where 单据=1 And No=[1] And (Mod(记录状态,3)=0 Or 记录状态=1) " & _
                                " And 序号=[2]) "
                            strOrder = rsInitCard!序号
                            Set rs = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[取付款序号]", txtNO.Text, strOrder)
                            
                            If rs!付款序号 <> 0 Then
                                mintParallelRecord = 5
                                Exit Sub
                            End If
                        End If
                        .MoveNext
                    Loop
                    If .RecordCount <> 0 Then .MoveFirst
                End With
            End If
            
            intRow = 0
            mbln退货 = (rsInitCard!退货 = 1)
            If mbln退货 Then LblTitle.Caption = Mid(LblTitle.Caption, 1, Len(LblTitle.Caption) - 5) & "退货单"
            With mshBill
                Do While Not rsInitCard.EOF
                    intRow = intRow + 1
                    .rows = .rows + 1
                    
                    .TextMatrix(intRow, 0) = rsInitCard.Fields(0)
                    
                    If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
                        str药名 = rsInitCard!通用名
                    Else
                        str药名 = IIf(IsNull(rsInitCard!商品名), rsInitCard!通用名, rsInitCard!商品名)
                    End If
                    
                    .TextMatrix(intRow, mconIntCol药品编码和名称) = rsInitCard!药品编码 & str药名
                    .TextMatrix(intRow, mconIntCol药品编码) = rsInitCard!药品编码
                    .TextMatrix(intRow, mconIntCol药品名称) = str药名
                    .TextMatrix(intRow, 0) = rsInitCard!药品ID
                    If mintDrugNameShow = 1 Then
                        .TextMatrix(intRow, mconIntCol药名) = .TextMatrix(intRow, mconIntCol药品编码)
                    ElseIf mintDrugNameShow = 2 Then
                        .TextMatrix(intRow, mconIntCol药名) = .TextMatrix(intRow, mconIntCol药品名称)
                    Else
                        .TextMatrix(intRow, mconIntCol药名) = .TextMatrix(intRow, mconIntCol药品编码和名称)
                    End If
                    
                    .TextMatrix(intRow, mconIntCol商品名) = IIf(IsNull(rsInitCard!商品名), "", rsInitCard!商品名)
                    
                    .TextMatrix(intRow, mconIntCol来源) = nvl(rsInitCard!药品来源)
                    .TextMatrix(intRow, mconIntCol基本药物) = nvl(rsInitCard!基本药物)
                    .TextMatrix(intRow, mconIntCol药价级别) = IIf(IsNull(rsInitCard!药价级别), "", rsInitCard!药价级别)
                    .TextMatrix(intRow, mconIntCol规格) = IIf(IsNull(rsInitCard!规格), "", rsInitCard!规格)
                    .TextMatrix(intRow, mconIntCol产地) = IIf(IsNull(rsInitCard!产地), "", rsInitCard!产地)
                    .TextMatrix(intRow, mconIntCol原产地) = IIf(IsNull(rsInitCard!原产地), "", rsInitCard!原产地)
                    .TextMatrix(intRow, mconIntCol单位) = rsInitCard!单位
                    .TextMatrix(intRow, mconIntCol批号) = IIf(IsNull(rsInitCard!批号), "", rsInitCard!批号)
                    .TextMatrix(intRow, mconIntCol效期) = IIf(IsNull(rsInitCard!效期), "", rsInitCard!效期)
                    
                    If gtype_UserSysParms.P149_效期显示方式 = 1 And .TextMatrix(intRow, mconIntCol效期) <> "" Then
                        '换算为有效期
                        .TextMatrix(intRow, mconIntCol效期) = Format(DateAdd("D", -1, .TextMatrix(intRow, mconIntCol效期)), "yyyy-mm-dd")
                    End If
                
                    If mbln退货 Then
                        .TextMatrix(intRow, mconIntCol数量) = zlStr.FormatEx(rsInitCard!数量 * IIf(mint编辑状态 = 6, 1, -1), intNumberDigit, , True)
                        .TextMatrix(intRow, mconIntCol成本金额) = zlStr.FormatEx(rsInitCard!采购金额 * IIf(mint编辑状态 = 6, 1, -1), intMoneyDigit, , True)
                        .TextMatrix(intRow, mconIntCol售价金额) = zlStr.FormatEx(rsInitCard!零售金额 * IIf(mint编辑状态 = 6, 1, -1), intMoneyDigit, , True)
                        .TextMatrix(intRow, mconintCol差价) = zlStr.FormatEx(rsInitCard!差价 * IIf(mint编辑状态 = 6, 1, -1), intMoneyDigit, , True)
                        .TextMatrix(intRow, mconintcol发票金额) = IIf(zlStr.FormatEx(nvl(rsInitCard!发票金额, 0) * IIf(mint编辑状态 = 6, 1, -1), intMoneyDigit) = "0.00", "", zlStr.FormatEx(nvl(rsInitCard!发票金额, 0) * IIf(mint编辑状态 = 6, 1, -1), intMoneyDigit, , True))
                    Else
                        .TextMatrix(intRow, mconIntCol数量) = zlStr.FormatEx(rsInitCard!数量, intNumberDigit, , True)
                        .TextMatrix(intRow, mconIntCol成本金额) = zlStr.FormatEx(IIf(mint编辑状态 = 6, 0, rsInitCard!采购金额), intMoneyDigit, , True)
                        .TextMatrix(intRow, mconIntCol售价金额) = zlStr.FormatEx(rsInitCard!零售金额, intMoneyDigit, , True)
                        .TextMatrix(intRow, mconintCol差价) = zlStr.FormatEx(rsInitCard!差价, intMoneyDigit, , True)
                        .TextMatrix(intRow, mconintcol发票金额) = IIf(zlStr.FormatEx(IIf(IsNull(rsInitCard!发票金额), "0", rsInitCard!发票金额), intMoneyDigit) = "0.00", "", zlStr.FormatEx(IIf(IsNull(rsInitCard!发票金额), "0", rsInitCard!发票金额), intMoneyDigit, , True))
                    End If
                    .TextMatrix(intRow, mconIntCol成本价) = zlStr.FormatEx(rsInitCard!采购价, intCostDigit, , True)
                    .TextMatrix(intRow, mconIntCol售价) = zlStr.FormatEx(rsInitCard!零售价, intPriceDigit, , True)
                    .TextMatrix(intRow, mconIntCol扣率) = rsInitCard!扣率
                    .TextMatrix(intRow, mconIntCol批准文号) = IIf(IsNull(rsInitCard!批准文号), "", rsInitCard!批准文号)
                    .TextMatrix(intRow, mconintcol随货单号) = IIf(IsNull(rsInitCard!随货单号), "", rsInitCard!随货单号)
                    .TextMatrix(intRow, mconintcol随货日期) = IIf(IsNull(rsInitCard!随货日期), "", rsInitCard!随货日期)
                    .TextMatrix(intRow, mconintcol发票号) = IIf(IsNull(rsInitCard!发票号), "", rsInitCard!发票号)
                    .TextMatrix(intRow, mconintcol发票代码) = IIf(IsNull(rsInitCard!发票代码), "", rsInitCard!发票代码)
                    .TextMatrix(intRow, mconIntCol发票日期) = IIf(IsNull(rsInitCard!发票日期), "", rsInitCard!发票日期)
                    .TextMatrix(intRow, mconIntCol指导批发价) = zlStr.FormatEx(rsInitCard!指导批发价, intCostDigit, , True)
                    .TextMatrix(intRow, mconIntCol原生产商) = IIf(IsNull(rsInitCard!原生产商), "!", rsInitCard!原生产商)
                    
                    .TextMatrix(intRow, mconIntCol原销期) = IIf(IsNull(rsInitCard!最大效期), "0", rsInitCard!最大效期) & "||" & rsInitCard!加成率 * 100 & "||" & IIf(IsNull(rsInitCard!是否变价), 0, rsInitCard!是否变价) & "||" & IIf(IsNull(rsInitCard!药房分批核算), 0, rsInitCard!药房分批核算)
                    
                    '分批属性
                    Call Get药品分批属性(intRow)
                    
                    '时价分批药品处理，需要重算界面的售价、售价金额、差价
                    If .TextMatrix(intRow, mconIntCol原销期) <> "" Then
                        If Val(Split(.TextMatrix(intRow, mconIntCol原销期), "||")(2)) = 1 And Val(.TextMatrix(intRow, mconIntCol分批属性)) = 1 Then
                            .TextMatrix(intRow, mconintCol零售单位) = rsInitCard!售价单位
                            .TextMatrix(intRow, mconintCol零售价) = zlStr.FormatEx(rsInitCard!零售价 / Val(rsInitCard!比例系数), gtype_UserDrugDigits.Digit_零售价, , True)

                            If mbln退货 Then
                                .TextMatrix(intRow, mconintCol零售金额) = zlStr.FormatEx(-1 * rsInitCard!零售金额, intMoneyDigit, , True)
                                .TextMatrix(intRow, mconintCol零售差价) = zlStr.FormatEx(-1 * rsInitCard!差价, intMoneyDigit, , True)
                            Else
                                .TextMatrix(intRow, mconintCol零售金额) = zlStr.FormatEx(rsInitCard!零售金额, intMoneyDigit, , True)
                                .TextMatrix(intRow, mconintCol零售差价) = zlStr.FormatEx(rsInitCard!差价, intMoneyDigit, , True)
                            End If

                            If mint编辑状态 <> 6 Then   '不是冲销时
                                If mbln退货 Then
                                    .TextMatrix(intRow, mconIntCol售价金额) = zlStr.FormatEx(-1 * (rsInitCard!零售金额 - rsInitCard!金额差), intMoneyDigit, , True)
                                    .TextMatrix(intRow, mconintCol差价) = zlStr.FormatEx(-1 * (rsInitCard!差价 - rsInitCard!金额差), intMoneyDigit, , True)
                                Else
                                    .TextMatrix(intRow, mconIntCol售价金额) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol零售金额)) - Val(rsInitCard!金额差), intMoneyDigit, , True)
                                    .TextMatrix(intRow, mconintCol差价) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol零售差价)) - Val(rsInitCard!金额差), intMoneyDigit, , True)
                                End If
                                If Val(.TextMatrix(intRow, mconIntCol数量)) <> 0 And rsInitCard!金额差 <> 0 Then
                                    .TextMatrix(intRow, mconIntCol售价) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol售价金额)) / Val(.TextMatrix(intRow, mconIntCol数量)), intPriceDigit, , True)
                                Else
                                    .TextMatrix(intRow, mconIntCol售价) = zlStr.FormatEx(rsInitCard!零售价, intPriceDigit, , True)
                                End If
                            Else
                                '冲销时
                                .TextMatrix(intRow, mconIntCol售价金额) = zlStr.FormatEx(0, intMoneyDigit, , True)
                                .TextMatrix(intRow, mconintCol差价) = zlStr.FormatEx(0, intMoneyDigit, , True)
                                .TextMatrix(intRow, mconIntCol售价) = zlStr.FormatEx((Val(.TextMatrix(intRow, mconintCol零售价)) * Val(rsInitCard!比例系数) * Val(rsInitCard!数量) - Val(rsInitCard!金额差)) / Val(rsInitCard!数量), intPriceDigit, , True)
                            End If
                        End If
                    End If
                    
                    .TextMatrix(intRow, mconintcol简码) = ""
                    .TextMatrix(intRow, mconIntCol比例系数) = rsInitCard!比例系数
                    .TextMatrix(intRow, mconIntCol序号) = rsInitCard!序号
                    .TextMatrix(intRow, mconIntCol生产日期) = IIf(IsNull(rsInitCard!生产日期), "", rsInitCard!生产日期)
                    .TextMatrix(intRow, mconIntCol采购价) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol成本价)) * 100 / IIf(Val(.TextMatrix(intRow, mconIntCol扣率)) = 0, 1, Val(.TextMatrix(intRow, mconIntCol扣率))), intCostDigit, , True)
                    .TextMatrix(intRow, mconIntCol是否新行) = "否"
                    If mint编辑状态 = 6 Then
                        .TextMatrix(intRow, mconIntCol外观) = nvl(rsInitCard!外观)
                    End If
                    If mint编辑状态 <> 6 Then
                        If (mint编辑状态 = 5 Or mint编辑状态 = 7) And mint记录状态 <> 1 Then
                            .TextMatrix(intRow, mconIntCol外观) = ""
                            .TextMatrix(intRow, mconIntCol验收结论) = ""
                            .TextMatrix(intRow, mconintcol产品合格证) = ""
                        Else
                            .TextMatrix(intRow, mconIntCol外观) = nvl(rsInitCard!外观)
                            .TextMatrix(intRow, mconIntCol验收结论) = nvl(rsInitCard!验收结论)
                            .TextMatrix(intRow, mconintcol产品合格证) = nvl(rsInitCard!产品合格证)
                        End If
                    End If
                    
                    '根据参数决定时价药品售价公式中成本价的算法
                    dbl成本价 = IIf(mint时价入库售价加成方式 = 0, Val(.TextMatrix(intRow, mconIntCol成本价)), Val(.TextMatrix(intRow, mconIntCol采购价)))
                    
                    '计算加成率
                    If Val(.TextMatrix(intRow, mconIntCol售价)) <> 0 And dbl成本价 <> 0 Then
                        If IIf(IsNull(rsInitCard!加成率), "", rsInitCard!加成率) <> "" Then
                            .TextMatrix(intRow, mconIntcol加成率) = zlStr.FormatEx(Val(rsInitCard!加成率) * 100, 2) & "%"
                        Else
                            .TextMatrix(intRow, mconIntcol加成率) = zlStr.FormatEx((Val(.TextMatrix(intRow, mconIntCol售价)) / dbl成本价 - 1) * 100, 2) & "%"
                        End If
                    End If
                    .TextMatrix(intRow, mconIntCol计划id) = IIf(IsNull(rsInitCard!计划id), "", rsInitCard!计划id)

                    '招标药品需要上色
                    mblnEnter = False
                    .Row = intRow
                    For i = mconIntCol药名 To .Cols - 1
                        j = .ColData(i)
                        If .ColData(i) = 5 Then .ColData(i) = 0
                        .Col = i
                        If rsInitCard!招标药品 = 1 Then
                            .MsfObj.CellForeColor = IIf(rsInitCard!差价让利比 = 0, &H800000, &H800080)
                        Else
                            .MsfObj.CellForeColor = IIf(rsInitCard!差价让利比 = 0, &H0, &H40&)     ' &H40C0&
                        End If
                        .ColData(i) = j
                    Next
                    mblnEnter = True
                    
                    If mint编辑状态 = 6 Then
                        .TextMatrix(intRow, mconIntCol冲销数量) = zlStr.FormatEx(0, intNumberDigit, , True)
                        .RowData(intRow) = rsInitCard!付款序号
                        .TextMatrix(intRow, mconIntCol批次) = rsInitCard!批次
                        
                        If rsInitCard!付款序号 = 0 Then
                            '检查是否存在部分付款的情况
                            gstrSQL = "Select Nvl(Max(付款序号), 0) 付款序号 From 应付记录 " & _
                                " where 收发id=(Select Id From 药品收发记录 Where 单据=1 And No=[1] And (Mod(记录状态,3)=0 Or 记录状态=1) " & _
                                " And 序号=[2]) "
                            Set rs = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[取付款序号]", txtNO.Text, Val(.TextMatrix(intRow, mconIntCol序号)))
                            
                            If rs!付款序号 = 0 Then
                                blnAllPay = False
                            End If
                        End If
                    Else
'                        If (mint编辑状态 = 5 Or mint编辑状态 = 7) And mint记录状态 <> 1 Then
'                            .TextMatrix(intRow, mconIntCol批次) = 0
'                        Else
                            .TextMatrix(intRow, mconIntCol批次) = rsInitCard!批次
'                        End If
                        .RowData(intRow) = nvl(rsInitCard!付款序号, 0)
                    End If
                    
                    If mint编辑状态 = 5 Or mint编辑状态 = 6 Or mint编辑状态 = 7 Then
                        .TextMatrix(intRow, mconIntCol付款标志) = IIf(IsNull(rsInitCard!付款标志), "否", IIf(rsInitCard!付款标志 = 0, "否", "是"))
                    End If
                    
                    rsInitCard.MoveNext
                Loop
                .Col = mconIntCol药名
                .CmdVisible = False
            End With
            rsInitCard.Close
    End Select
    SetEdit         '设置编辑属性
    Call RefreshRowNO(mshBill, mconIntCol行号, 1)
    Call 显示合计金额
    
    If mint编辑状态 = 6 And blnAllPay = True Then
        mintParallelRecord = 7
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetEdit()
    Dim intCol As Integer
    
    With mshBill
        If mblnEdit = False Then
            
            cboStock.Enabled = False
            txtProvider.Enabled = False
            cmdProvider.Enabled = False
            txt摘要.Enabled = False
            
            For intCol = 0 To .Cols - 1
                .ColData(intCol) = 0
            Next
            If mint编辑状态 = 3 Then
                .ColData(mconIntCol成本价) = IIf(InStr(1, mstrControlItem, ",成本价,") > 0, 4, 5)
                .ColData(mconIntCol采购价) = IIf(InStr(1, mstrControlItem, ",采购价,") > 0, 4, 5)
                .ColData(mconIntCol售价) = IIf(InStr(1, mstrControlItem, ",售价,") > 0, 4, 5)
                .ColData(mconIntCol扣率) = IIf(InStr(1, mstrControlItem, ",扣率,") > 0, 4, 5)
                .ColData(mconIntCol成本金额) = IIf(InStr(1, mstrControlItem, ",成本金额,") > 0, 4, 5)
                .ColData(mconIntCol外观) = IIf(InStr(1, mstrControlItem, ",外观,") > 0, 1, 5)
                .ColData(mconIntCol验收结论) = IIf(InStr(1, mstrControlItem, ",验收结论,") > 0, 1, 5)
                .ColData(mconintcol发票号) = IIf(InStr(1, mstrControlItem, ",发票号,") > 0, 4, 5)
                .ColData(mconintcol发票代码) = IIf(InStr(1, mstrControlItem, ",发票代码,") > 0, 4, 5)
                .ColData(mconIntCol发票日期) = IIf(InStr(1, mstrControlItem, ",发票日期,") > 0, 2, 5)
                .ColData(mconintcol发票金额) = IIf(InStr(1, mstrControlItem, ",发票金额,") > 0, 4, 5)

                txtProvider.Enabled = True
                cmdProvider.Enabled = True

                If mint记录状态 <> 1 Then
                    txtProvider.Enabled = False
                    cmdProvider.Enabled = False
                End If
            ElseIf mint编辑状态 = 5 Then
                .ColData(mconintcol发票号) = 4
                .ColData(mconintcol发票代码) = 4
                .ColData(mconIntCol发票日期) = 2

                txtProvider.Enabled = True
                cmdProvider.Enabled = True

                If mint记录状态 <> 1 Then
                    txtProvider.Enabled = False
                    cmdProvider.Enabled = False
                End If
            ElseIf mint编辑状态 = 6 Then
                For intCol = 0 To .Cols - 1
                    .ColData(intCol) = 5
                Next
                mshBill.ColData(mconIntCol药名) = 0
                mshBill.ColData(mconIntCol冲销数量) = 4
                txt摘要.Enabled = True
            ElseIf mint编辑状态 = 9 Then
                .ColData(mconIntCol成本价) = IIf(InStr(1, mstrControlItem, ",成本价,") > 0, 4, 5)
                .ColData(mconIntCol采购价) = IIf(InStr(1, mstrControlItem, ",采购价,") > 0, 4, 5)
                .ColData(mconIntCol售价) = IIf(InStr(1, mstrControlItem, ",售价,") > 0, 4, 5)
                .ColData(mconIntCol扣率) = IIf(InStr(1, mstrControlItem, ",扣率,") > 0, 4, 5)
                .ColData(mconIntCol成本金额) = IIf(InStr(1, mstrControlItem, ",成本金额,") > 0, 4, 5)
                .ColData(mconIntCol外观) = IIf(InStr(1, mstrControlItem, ",外观,") > 0, 1, 5)
                .ColData(mconIntCol验收结论) = IIf(InStr(1, mstrControlItem, ",验收结论,") > 0, 1, 5)
                .ColData(mconintcol发票号) = IIf(InStr(1, mstrControlItem, ",发票号,") > 0, 4, 5)
                .ColData(mconintcol发票代码) = IIf(InStr(1, mstrControlItem, ",发票代码,") > 0, 4, 5)
                .ColData(mconIntCol发票日期) = IIf(InStr(1, mstrControlItem, ",发票日期,") > 0, 2, 5)
                .ColData(mconintcol发票金额) = IIf(InStr(1, mstrControlItem, ",发票金额,") > 0, 4, 5)

                'Modifed by ZYB 20050104
                .ColData(mconIntCol指导批发价) = IIf(mbln修改批发价, 4, 0)
                'Modifed by ZYB 20050104 END
'                .ColData(mconintcol外观) = 4
'                .ColData(mconintcol产品合格证) = 4
                txt摘要.Enabled = True
            End If
            
            If mbln退货 Then
                txtProvider.Enabled = False
                cmdProvider.Enabled = False
            End If
        Else
            If mint编辑状态 = 7 Then
                txtProvider.Enabled = False
                cmdProvider.Enabled = False
                txt摘要.Enabled = False
                cboStock.Enabled = False
                
                For intCol = 0 To .Cols - 1
                    .ColData(intCol) = 5
                Next
                .ColData(mconIntCol成本价) = IIf(InStr(1, mstrControlItem, ",成本价,") > 0, 4, 5)
                .ColData(mconIntCol采购价) = IIf(InStr(1, mstrControlItem, ",采购价,") > 0, 4, 5)
                .ColData(mconIntCol售价) = IIf(InStr(1, mstrControlItem, ",售价,") > 0, 4, 5)
                .ColData(mconIntCol扣率) = IIf(InStr(1, mstrControlItem, ",扣率,") > 0, 4, 5)
                .ColData(mconIntCol成本金额) = IIf(InStr(1, mstrControlItem, ",成本金额,") > 0, 4, 5)
                .ColData(mconIntCol外观) = IIf(InStr(1, mstrControlItem, ",外观,") > 0, 1, 5)
                .ColData(mconIntCol验收结论) = IIf(InStr(1, mstrControlItem, ",验收结论,") > 0, 1, 5)
                .ColData(mconintcol发票号) = IIf(InStr(1, mstrControlItem, ",发票号,") > 0, 4, 5)
                .ColData(mconintcol发票代码) = IIf(InStr(1, mstrControlItem, ",发票代码,") > 0, 4, 5)
                .ColData(mconIntCol发票日期) = IIf(InStr(1, mstrControlItem, ",发票日期,") > 0, 2, 5)
                .ColData(mconintcol发票金额) = IIf(InStr(1, mstrControlItem, ",发票金额,") > 0, 4, 5)

'                .LocateCol = mconIntCol成本价
                Exit Sub
            ElseIf mint编辑状态 = 8 Or mbln退货 Then
                .ColData(mconIntCol批号) = 5
                .ColData(mconIntCol生产日期) = 5
                .ColData(mconIntCol效期) = 5
                .ColData(mconIntCol扣率) = 5
                .ColData(mconIntCol指导批发价) = IIf(mbln修改批发价, 4, 5)
                '.ColData(mconIntCol成本价) = 5
                .ColData(mconIntCol成本金额) = 5
                If mbln退货 Then
                    txtProvider.Enabled = False
                    cmdProvider.Enabled = False
                End If
                '退货单不允许选择库房
                cboStock.Enabled = False
                Exit Sub
            End If
            .ColData(0) = 5
            .ColData(mconIntCol药名) = 1
            .ColData(mconIntCol序号) = 5
            .ColData(mconIntCol规格) = 5
            If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
                .ColData(mconIntCol产地) = 1
                .ColData(mconIntCol原产地) = 1
            Else
                .ColData(mconIntCol产地) = 5
                .ColData(mconIntCol原产地) = 5
            End If
            .ColData(mconIntCol单位) = 5
            .ColData(mconIntCol批号) = 4
            .ColData(mconIntCol生产日期) = 2
            .ColData(mconIntCol效期) = 5
            .ColData(mconIntCol数量) = 4
            
            .ColData(mconIntCol售价) = 5
            .ColData(mconIntCol售价金额) = 5
            .ColData(mconintCol差价) = 5
            
            .ColData(mconintcol发票号) = 4
            .ColData(mconintcol发票代码) = 4
            .ColData(mconIntCol发票日期) = 2
            
            .ColData(mconIntCol指导批发价) = IIf(mbln修改批发价, 4, 5)
            .ColData(mconIntCol原生产商) = 5
            .ColData(mconIntCol原销期) = 5
            .ColData(mconintcol简码) = 5
            .ColData(mconIntCol比例系数) = 5
            .ColData(mconIntCol批准文号) = 4
            .ColData(mconIntCol付款标志) = 5

            .ColData(mconIntCol成本价) = 4
            .ColData(mconIntCol采购价) = 4
            .ColData(mconIntCol成本金额) = 4
            .ColData(mconIntCol扣率) = 4
            .ColData(mconintcol发票金额) = 4
            
            .ColAlignment(mconIntCol药名) = flexAlignLeftCenter
            .ColAlignment(mconIntCol规格) = flexAlignLeftCenter
            .ColAlignment(mconIntCol产地) = flexAlignLeftCenter
            .ColAlignment(mconIntCol原产地) = flexAlignLeftCenter
            .ColAlignment(mconIntCol单位) = flexAlignCenterCenter
            .ColAlignment(mconIntCol批号) = flexAlignLeftCenter
            .ColAlignment(mconIntCol生产日期) = flexAlignLeftCenter
            .ColAlignment(mconIntCol效期) = flexAlignLeftCenter
            .ColAlignment(mconIntCol数量) = flexAlignRightCenter
            .ColAlignment(mconIntCol成本价) = flexAlignRightCenter
            .ColAlignment(mconIntCol成本金额) = flexAlignRightCenter
            .ColAlignment(mconIntCol售价) = flexAlignRightCenter
            .ColAlignment(mconIntCol售价金额) = flexAlignRightCenter
            .ColAlignment(mconintCol差价) = flexAlignRightCenter
            .ColAlignment(mconIntCol扣率) = flexAlignRightCenter
            .ColAlignment(mconintcol发票号) = flexAlignLeftCenter
            .ColAlignment(mconintcol发票代码) = flexAlignLeftCenter
            .ColAlignment(mconIntCol发票日期) = flexAlignLeftCenter
            .ColAlignment(mconintcol发票金额) = flexAlignRightCenter
            .ColAlignment(mconIntCol付款标志) = flexAlignLeftCenter
            
            cboStock.Enabled = True
           
            txtProvider.Enabled = True
            cmdProvider.Enabled = True
            txt摘要.Enabled = True
        End If
    End With
End Sub


Private Sub initGrid()
    '表格初始化、初始化摘要文本框的长度
    With mshBill
        .Active = True
        .Cols = mconIntColS
        
        .MsfObj.FixedCols = 1
        Call SetColumnByUserDefine
        .TextMatrix(0, mconIntCol行号) = ""
        .TextMatrix(0, mconIntCol药名) = "药品名称与编码"
        .TextMatrix(0, mconIntCol商品名) = "商品名"
        .TextMatrix(0, mconIntCol来源) = "药品来源"
        .TextMatrix(0, mconIntCol基本药物) = "基本药物"
        .TextMatrix(0, mconIntCol序号) = "序号"
        .TextMatrix(0, mconIntCol药价级别) = "药价级别"
        .TextMatrix(0, mconIntCol规格) = "规格"
        .TextMatrix(0, mconIntCol产地) = "生产商"
        .TextMatrix(0, mconIntCol原产地) = "原产地"
        .TextMatrix(0, mconIntCol单位) = "单位"
        .TextMatrix(0, mconIntCol批号) = "批号"
        .TextMatrix(0, mconIntCol生产日期) = "生产日期"
        .TextMatrix(0, mconIntCol效期) = IIf(gtype_UserSysParms.P149_效期显示方式 = 1, "有效期至", "失效期")
        .TextMatrix(0, mconIntCol数量) = "数量"
        .TextMatrix(0, mconIntCol冲销数量) = "冲销数量"
        .TextMatrix(0, mconIntCol批次) = "批次"
        .TextMatrix(0, mconIntCol成本价) = "成本价"
        .TextMatrix(0, mconIntCol成本金额) = "成本金额"
        .TextMatrix(0, mconIntCol售价) = "售价"
        .TextMatrix(0, mconIntCol售价金额) = "售价金额"
        .TextMatrix(0, mconintCol差价) = "差价"
        .TextMatrix(0, mconintCol零售价) = "零售价"
        .TextMatrix(0, mconintCol零售单位) = "零售单位"
        .TextMatrix(0, mconintCol零售金额) = "零售金额"
        .TextMatrix(0, mconintCol零售差价) = "零售差价"
        .TextMatrix(0, mconIntCol扣率) = "扣率"
        .TextMatrix(0, mconIntCol批准文号) = "批准文号"
        .TextMatrix(0, mconintcol随货单号) = "随货单号"
        .TextMatrix(0, mconintcol随货日期) = "随货日期"
        .TextMatrix(0, mconintcol发票号) = "发票号"
        .TextMatrix(0, mconintcol发票代码) = "发票代码"
        .TextMatrix(0, mconIntCol发票日期) = "发票日期"
        .TextMatrix(0, mconintcol发票金额) = "发票金额"
        .TextMatrix(0, mconIntCol指导批发价) = "采购限价"
        .TextMatrix(0, mconIntCol原生产商) = "原生产商"
        .TextMatrix(0, mconIntCol原销期) = "原效期"
        .TextMatrix(0, mconintcol简码) = "简码"
        .TextMatrix(0, mconIntCol比例系数) = "比例系数"
        .TextMatrix(0, mconIntCol外观) = "外观"
        .TextMatrix(0, mconIntCol验收结论) = "验收结论"
        .TextMatrix(0, mconintcol产品合格证) = "产品合格证"
        .TextMatrix(0, mconIntCol采购价) = "采购价"
        .TextMatrix(0, mconIntCol分批属性) = "分批属性"
        .TextMatrix(0, mconIntCol是否新行) = "是否新行"
        .TextMatrix(0, mconIntcol加成率) = "加成率"
        .TextMatrix(0, mconIntCol药品编码和名称) = "药品编码和名称"
        .TextMatrix(0, mconIntCol药品编码) = "药品编码"
        .TextMatrix(0, mconIntCol药品名称) = "药品名称"
        .TextMatrix(0, mconIntCol付款标志) = "付款标志"
        .TextMatrix(0, mconIntCol计划id) = "计划id"
                
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, mconIntCol行号) = "1"
        
        .ColWidth(0) = 0
        .ColWidth(mconIntCol行号) = 300
        .ColWidth(mconIntCol药名) = 2000
        .ColWidth(mconIntCol商品名) = 2000
        .ColWidth(mconIntCol来源) = 900
        .ColWidth(mconIntCol基本药物) = 900
        .ColWidth(mconIntCol序号) = 0
        .ColWidth(mconIntCol药价级别) = 1200
        .ColWidth(mconIntCol规格) = 900
        .ColWidth(mconIntCol产地) = 800
        .ColWidth(mconIntCol原产地) = 0
        .ColWidth(mconIntCol单位) = 500
        .ColWidth(mconIntCol批号) = 800
        .ColWidth(mconIntCol生产日期) = 1000
        .ColWidth(mconIntCol效期) = 1000
        .ColWidth(mconIntCol数量) = 1100
        .ColWidth(mconIntCol冲销数量) = IIf(mint编辑状态 = 6, 1100, 0)
        .ColWidth(mconIntCol批次) = 0
        .ColWidth(mconIntCol成本价) = 1000
        .ColWidth(mconIntCol成本金额) = 900
        .ColWidth(mconIntCol售价) = 1000
        .ColWidth(mconIntCol售价金额) = 900
        .ColWidth(mconintCol差价) = 800
        .ColWidth(mconintCol零售价) = IIf(mintUnit = mconint售价单位, 0, 1000)
        .ColWidth(mconintCol零售单位) = IIf(mintUnit = mconint售价单位, 0, 1000)
        .ColWidth(mconintCol零售金额) = IIf(mintUnit = mconint售价单位, 0, 1000)
        .ColWidth(mconintCol零售差价) = IIf(mintUnit = mconint售价单位, 0, 1000)
        .ColWidth(mconIntCol扣率) = 800
        .ColWidth(mconIntCol批准文号) = 1000
        .ColWidth(mconintcol随货单号) = 1200
        .ColWidth(mconintcol随货日期) = 1000
        .ColWidth(mconintcol发票号) = 800
        .ColWidth(mconintcol发票代码) = 1000
        .ColWidth(mconIntCol发票日期) = 1000
        .ColWidth(mconintcol发票金额) = 900
        .ColWidth(mconIntCol指导批发价) = 1000
        .ColWidth(mconIntCol原生产商) = 0
        .ColWidth(mconIntCol原销期) = 0
        .ColWidth(mconintcol简码) = 0
        .ColWidth(mconIntCol比例系数) = 0
        .ColWidth(mconIntCol外观) = 1000
        .ColWidth(mconIntCol验收结论) = 4500
        .ColWidth(mconintcol产品合格证) = 1000
        .ColWidth(mconIntCol采购价) = 1000
        .ColWidth(mconIntCol分批属性) = 0
        .ColWidth(mconIntCol是否新行) = 0
        .ColWidth(mconIntcol加成率) = 1000
        .ColWidth(mconIntCol药品编码和名称) = 0
        .ColWidth(mconIntCol药品编码) = 0
        .ColWidth(mconIntCol药品名称) = 0
        If mint编辑状态 = 6 Then
            .ColWidth(mconIntCol付款标志) = 800
        Else
            .ColWidth(mconIntCol付款标志) = 0
        End If
        .ColWidth(mconIntCol计划id) = 0
                
        '-1：表示该列可以选择，是布尔型［"√"，" "］
        ' 0：表示该列可以选择，但不能修改
        ' 1：表示该列可以输入，外部显示为按钮选择
        ' 2：表示该列是日期列，外部显示为按钮选择，弹出是日期选择框
        ' 3：表示该列是选择列，外部显示为下拉框选择
        '4:  表示该列为单纯的文本框供用户输入
        '5:  表示该列不允许选择

        .ColData(0) = 5
        .ColData(mconIntCol行号) = 5
        .ColData(mconIntCol药名) = 1
        .ColData(mconIntCol商品名) = 5
        .ColData(mconIntCol来源) = 5
        .ColData(mconIntCol基本药物) = 5
        .ColData(mconIntCol序号) = 5
        .ColData(mconIntCol药价级别) = 5
        .ColData(mconIntCol规格) = 5
        If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
            .ColData(mconIntCol产地) = 1
            .ColData(mconIntCol原产地) = 1
        Else
            .ColData(mconIntCol产地) = 5
            .ColData(mconIntCol原产地) = 5
        End If
        .ColData(mconIntCol单位) = 5
        .ColData(mconIntCol批号) = 4
        .ColData(mconIntCol生产日期) = 2
        .ColData(mconIntCol效期) = 5
        .ColData(mconIntCol数量) = 4
        .ColData(mconIntCol冲销数量) = 5
        .ColData(mconIntCol批次) = 5
        
        .ColData(mconIntCol售价) = 5
        .ColData(mconIntCol售价金额) = 5
        .ColData(mconintCol差价) = 5
        
        .ColData(mconintCol零售价) = 5
        .ColData(mconintCol零售单位) = 5
        .ColData(mconintCol零售金额) = 5
        .ColData(mconintCol零售差价) = 5
        
        .ColData(mconIntCol批准文号) = 5
        .ColData(mconintcol随货单号) = 4
        .ColData(mconintcol随货日期) = 2
        .ColData(mconintcol发票号) = 4
        .ColData(mconintcol发票代码) = 4
        .ColData(mconIntCol发票日期) = 2
        
        .ColData(mconIntCol指导批发价) = IIf(mbln修改批发价, 4, 5)
        .ColData(mconIntCol原生产商) = 5
        .ColData(mconIntCol原销期) = 5
        .ColData(mconintcol简码) = 5
        .ColData(mconIntCol比例系数) = 5
        .ColData(mconIntCol外观) = 1
        .ColData(mconIntCol验收结论) = 1
        .ColData(mconintcol产品合格证) = 4
        .ColData(mconIntCol是否新行) = 5
        .ColData(mconIntcol加成率) = 5
        .ColData(mconIntCol药品编码和名称) = 5
        .ColData(mconIntCol药品编码) = 5
        .ColData(mconIntCol药品名称) = 5

        .ColData(mconIntCol成本价) = 4
        .ColData(mconIntCol成本金额) = 4
        .ColData(mconIntCol扣率) = 4
        .ColData(mconintcol发票金额) = 4
        .ColData(mconIntCol采购价) = 4
        
        .ColAlignment(mconIntCol药名) = flexAlignLeftCenter
        .ColAlignment(mconIntCol商品名) = flexAlignLeftCenter
        .ColAlignment(mconIntCol来源) = flexAlignLeftCenter
        .ColAlignment(mconIntCol基本药物) = flexAlignLeftCenter
        .ColAlignment(mconIntCol药价级别) = flexAlignLeftCenter
        .ColAlignment(mconIntCol规格) = flexAlignLeftCenter
        .ColAlignment(mconIntCol产地) = flexAlignLeftCenter
        .ColAlignment(mconIntCol原产地) = flexAlignLeftCenter
        .ColAlignment(mconIntCol单位) = flexAlignCenterCenter
        .ColAlignment(mconIntCol批号) = flexAlignLeftCenter
        .ColAlignment(mconIntCol生产日期) = flexAlignLeftCenter
        .ColAlignment(mconIntCol效期) = flexAlignLeftCenter
        .ColAlignment(mconIntCol数量) = flexAlignRightCenter
        .ColAlignment(mconIntCol冲销数量) = flexAlignRightCenter
        .ColAlignment(mconIntCol成本价) = flexAlignRightCenter
        .ColAlignment(mconIntCol成本金额) = flexAlignRightCenter
        .ColAlignment(mconIntCol售价) = flexAlignRightCenter
        .ColAlignment(mconIntCol售价金额) = flexAlignRightCenter
        .ColAlignment(mconintCol差价) = flexAlignRightCenter
        .ColAlignment(mconintCol零售价) = flexAlignRightCenter
        .ColAlignment(mconintCol零售单位) = flexAlignRightCenter
        .ColAlignment(mconintCol零售金额) = flexAlignRightCenter
        .ColAlignment(mconintCol零售差价) = flexAlignRightCenter
        .ColAlignment(mconIntCol扣率) = flexAlignRightCenter
        .ColAlignment(mconIntCol批准文号) = flexAlignLeftCenter
        .ColAlignment(mconintcol随货单号) = flexAlignLeftCenter
        .ColAlignment(mconintcol随货日期) = flexAlignLeftCenter
        .ColAlignment(mconintcol发票号) = flexAlignLeftCenter
        .ColAlignment(mconintcol发票代码) = flexAlignLeftCenter
        .ColAlignment(mconIntCol发票日期) = flexAlignLeftCenter
        .ColAlignment(mconintcol发票金额) = flexAlignRightCenter
        .ColAlignment(mconIntCol采购价) = flexAlignRightCenter
        .ColAlignment(mconIntcol加成率) = flexAlignRightCenter
        
        .PrimaryCol = mconIntCol药名
        .LocateCol = mconIntCol药名
    End With
    
    Call SetColumnByUserDefine
    txt摘要.MaxLength = Sys.FieldsLength("药品收发记录", "摘要")
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    If Me.Width < 12735 Then Me.Width = 12735
    
    With Pic单据
        
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - IIf(staThis.Visible, staThis.Height, 0) - .Top - 100 - CmdCancel.Height - 200

    End With
    

    
    With LblTitle
        .Left = 0
        .Top = 150
        .Width = Pic单据.Width
    End With
    
    
    With mshBill
        .Left = 200
        .Width = Pic单据.Width - .Left * 2
    End With
    With txtNO
        .Left = mshBill.Left + mshBill.Width - .Width
        LblNo.Left = .Left - LblNo.Width - 100
        .Top = LblTitle.Top
        LblNo.Top = .Top
    End With
    
    LblStock.Left = mshBill.Left
    cboStock.Left = LblStock.Left + LblStock.Width + 100
    
    cmdProvider.Left = mshBill.Left + mshBill.Width - cmdProvider.Width
    txtProvider.Left = cmdProvider.Left - txtProvider.Width
    LblProvider.Left = txtProvider.Left - LblProvider.Width - 100
    
    
    With Lbl填制日期
        .Top = Pic单据.Height - 200 - .Height
        .Left = mshBill.Left + 100
    End With
    
    With Txt填制日期
        .Top = Lbl填制日期.Top - 60
        .Left = Lbl填制日期.Left + Lbl填制日期.Width + 100
    End With
    
    With Lbl填制人
        .Top = Lbl填制日期.Top - Lbl填制日期.Height - 180
        .Left = Lbl填制日期.Left
    End With
    
    With Txt填制人
        .Top = Lbl填制人.Top - 60
        .Left = Txt填制日期.Left
    End With
    
    With lbl修改日期
        .Top = Lbl填制日期.Top
        .Left = Txt填制日期.Left + Txt填制日期.Width + 400
    End With
    
    With Txt修改日期
        .Top = Txt填制日期.Top
        .Left = lbl修改日期.Left + lbl修改日期.Width + 100
    End With
    
    With lbl修改人
        .Top = Lbl填制人.Top
        .Left = lbl修改日期.Left
    End With
    
    With Txt修改人
        .Top = Txt填制人.Top
        .Left = Txt修改日期.Left
    End With

    With Txt审核日期
        .Top = Txt填制日期.Top
        .Left = mshBill.Left + mshBill.Width - .Width
    End With
    
    With Lbl审核日期
        .Top = Lbl填制日期.Top
        .Left = Txt审核日期.Left - 100 - .Width
    End With
    
    With Txt审核人
        .Top = Txt填制人.Top
        .Left = Txt审核日期.Left
    End With
    
    With Lbl审核人
        .Top = Lbl填制人.Top
        .Left = Lbl审核日期.Left
    End With
    
    With txt核查日期
        .Top = Txt填制日期.Top
        .Left = Lbl审核日期.Left - 400 - .Width
    End With
    
    With lbl核查日期
        .Top = Lbl填制日期.Top
        .Left = txt核查日期.Left - .Width - 100
    End With
    
    With txt核查人
        .Top = Txt填制人.Top
        .Left = txt核查日期.Left
    End With
    
    With lbl核查人
        .Top = Lbl填制人.Top
        .Left = lbl核查日期.Left
    End With
    
    With txt摘要
        .Top = Lbl填制人.Top - 140 - .Height
        .Left = Txt填制人.Left
        .Width = mshBill.Left + mshBill.Width - .Left
    End With
    
    With lbl摘要
        .Top = txt摘要.Top + 50
        .Left = txt摘要.Left - .Width - 180
    End With
    
    
    With lblPurchasePrice
        .Left = mshBill.Left
        .Top = txt摘要.Top - 60 - .Height
        .Width = mshBill.Width
        lblSalePrice.Top = .Top
        lblDifference.Top = .Top
    End With
    
    With lblSalePrice
        .Left = lblPurchasePrice.Left + mshBill.Width / 4
    End With
    With lblDifference
        .Left = lblPurchasePrice.Left + mshBill.Width / 3 * 2
    End With
    If mblnViewCost = False Then
        lblPurchasePrice.Visible = False
        lblDifference.Visible = False
    End If
    
    With mshBill
        .Height = lblPurchasePrice.Top - .Top - 60
    End With
    
    With CmdCancel
        .Left = Pic单据.Left + mshBill.Left + mshBill.Width - .Width
        .Top = Pic单据.Top + Pic单据.Height + 100
    End With
    
    With CmdSave
        .Left = CmdCancel.Left - .Width - 100
        .Top = CmdCancel.Top
    End With
    
    With cmdAllCls
        .Left = CmdSave.Left - .Width - 500
        .Top = CmdCancel.Top
    End With
    
    With cmdAllSel
        .Left = cmdAllCls.Left - .Width - 100
        .Top = CmdCancel.Top
    End With
    
    
    With cmdHelp
        .Left = Pic单据.Left + mshBill.Left
        .Top = CmdCancel.Top
    End With
    
    With cmdFind
        .Top = CmdCancel.Top
    End With
    
    cmdAddProducer.Visible = (mint编辑状态 = 1 Or mint编辑状态 = 2)
    With cmdAddProducer
        .Top = cmdFind.Top
        .Left = IIf(lblCode.Visible, txtCode.Left + txtCode.Width + 50, cmdFind.Left + cmdFind.Width + 50)
    End With
    
    If mint编辑状态 = 5 Then '修改发票信息该按钮才可用
        With cmdCopy
            cmdCopy.Visible = True
            .Top = cmdFind.Top
            If txtCode.Visible Then
                .Left = txtCode.Left + txtCode.Width + 100
            Else
                .Left = cmdFind.Left + cmdFind.Width + 100
            End If
        End With
        
        With cmdALLDel
            .Visible = True
            .Left = cmdCopy.Left + cmdCopy.Width + 100
            .Top = cmdCopy.Top
        End With
    End If
    
    With lblCode
        .Top = CmdCancel.Top + 50
    End With
    With txtCode
        .Top = CmdCancel.Top + 30
    End With
    
    Me.chk转入移库.Top = txtCode.Top
    Me.cboEnterStock.Top = txtCode.Top
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\药品外购入库管理", "药品名称显示方式", mintDrugNameShow)
    
    If msh产地.Visible = True Then
        msh产地.Visible = False
        mshBill.SetFocus
        mshBill.Col = mconIntCol产地
        Cancel = True
        Exit Sub
    End If
    
    If mblnChange = False Or mint编辑状态 = 4 Or mint编辑状态 = 3 Then
        SaveWinState Me, App.ProductName, MStrCaption
        Call ReleaseSelectorRS  '卸载数据集
        Exit Sub
    End If
    If MsgBox("数据可能已改变，但未存盘，真要退出吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
        Exit Sub
    Else
        SaveWinState Me, App.ProductName, MStrCaption
    End If
    mlng供药单位ID = 0
    Call ReleaseSelectorRS '卸载数据集
End Sub

Private Function SaveCheck(Optional ByVal strNo As String = "", Optional ByVal blnTrans As Boolean = False) As Boolean
    'blnTrans:表示是否开始单独的事务控制
    mblnSave = False
    SaveCheck = False
    
    Dim n As Integer
    Dim m As Integer
    Dim dbl合计数量 As Double
    Dim lng药品ID As Long
    Dim lng批次 As Long
    Dim str药品 As String
    Dim intNumCol As Integer
    
    '检查库存,只有出库业务才检查，如冲销，退库,分批和时价库存不足也禁止
    If mint编辑状态 = 6 Or mint编辑状态 = 8 Or mbln退货 = True Then
        If mint编辑状态 = 6 Then
            intNumCol = mconIntCol冲销数量
        Else
            intNumCol = mconIntCol数量
        End If
        str药品 = CheckNumStock(mshBill, cboStock.ItemData(cboStock.ListIndex), 0, mconIntCol批次, intNumCol, mconIntCol比例系数, IIf(mint编辑状态 = 3, IIf(mbln退货 = True, 3, 1), 3), , mintNumberDigit)
        If str药品 <> "" Then
            If mbln提示方式 = False Then
                If mint库存检查 = 1 Then '不足提醒
                    If MsgBox("药品【" & str药品 & "】库存不足，是否继续？", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                ElseIf mint库存检查 = 2 Then '不足禁止
                    MsgBox "药品【" & str药品 & "】库存不足，不能审核！", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    End If
    
    gstrSQL = "zl_药品外购_Verify('" & IIf(mint编辑状态 = 7, strNo, txtNO.Tag) & "'," & IIf(mint编辑状态 = 7, "'" & txtNO.Tag & "'", "Null") & ",'" & UserInfo.用户姓名 & "',to_date('" & Format(mstr审核日期, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS'))"
    On Error GoTo errHandle
    If blnTrans Then gcnOracle.BeginTrans
    Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
    
    If blnTrans Then gcnOracle.CommitTrans
    SaveCheck = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function



Private Sub mnuColDrug_Click(Index As Integer)
    Dim n As Integer
    
    With mnuColDrug
        For n = 0 To .count - 1
            .Item(n).Checked = False
        Next
        
        .Item(Index).Checked = True
        
        Call SetDrugName(Index)
    End With
End Sub

Private Sub mshBill_AfterAddRow(Row As Long)
    Call RefreshRowNO(mshBill, mconIntCol行号, Row)
    mshBill.Value = Format(Sys.Currentdate, "YYYY-MM-DD")
    mshBill.TextMatrix(Row, mconIntCol是否新行) = "是"
End Sub

Private Sub mshBill_AfterDeleteRow()
    Call 显示合计金额
    Call RefreshRowNO(mshBill, mconIntCol行号, mshBill.Row)
End Sub

Private Sub mshBill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    If InStr(1, "345679", mint编辑状态) <> 0 Then
        Cancel = True
        Exit Sub
    End If
    With mshBill
        If .TextMatrix(.Row, 0) <> "" Then
            If MsgBox("你确实要删除该行药品？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True
            End If
        End If
    End With
            
End Sub

Private Sub mshbill_CommandClick()
    Dim str药品ID As String
    Dim i As Integer
    Dim intRow As Integer
    Dim intOldRow As Integer
    
    On Error GoTo errHandle
    intOldRow = mshBill.Row
    Select Case mshBill.Col
    Case mconIntCol药名
        Dim RecReturn As Recordset
        
        mblnChange = True
        mshBill.CmdEnable = False
'        Set RecReturn = Frm药品选择器.ShowME(Me, IIf(mint编辑状态 = 8 Or mbln退货, 2, 1), cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , True, True, False, False, True, IIf(mint编辑状态 = 8 Or mbln退货, Val(txtProvider.Tag), 0))
        If grsMaster.State = adStateClosed Then
            Call SetSelectorRS(IIf(mint编辑状态 = 8 Or mbln退货, 2, 1), MStrCaption, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , IIf(mint编辑状态 = 8 Or mbln退货, Val(txtProvider.Tag), 0))
        End If
        Set RecReturn = frmSelector.ShowME(Me, 0, IIf(mint编辑状态 = 8 Or mbln退货, 2, 1), , , , cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , IIf(mint编辑状态 = 8 Or mbln退货, Val(txtProvider.Tag), 0), True, True, True, , , mstrPrivs)
        If RecReturn.RecordCount > 0 And (mint编辑状态 = 8 Or mbln退货 = True) Then
            Set RecReturn = CheckRedo(RecReturn) '检查重复记录，将重复的记录过滤掉然后返回过滤后的数据集
        End If
        
        mshBill.CmdEnable = True
        
        mshBill.Redraw = False
        If RecReturn.RecordCount > 0 Then
            
            RecReturn.MoveFirst
            For i = 1 To RecReturn.RecordCount
                With mshBill
                    .Redraw = False
                    mlng包装系数 = Choose(mintUnit, 1, RecReturn!门诊包装, RecReturn!住院包装, RecReturn!药库包装)
                    intRow = .Row
                    .TextMatrix(intRow, mconIntCol行号) = .Row
                    SetColValue .Row, RecReturn!药品ID, "[" & RecReturn!药品编码 & "]", RecReturn!通用名, IIf(IsNull(RecReturn!商品名), "", RecReturn!商品名), _
                        nvl(RecReturn!药品来源), "" & RecReturn!基本药物, _
                        IIf(IsNull(RecReturn!规格), "", RecReturn!规格), IIf(IsNull(RecReturn!产地), "", RecReturn!产地), _
                        Choose(mintUnit, RecReturn!售价单位, RecReturn!门诊单位, RecReturn!住院单位, RecReturn!药库单位), _
                        IIf(IsNull(RecReturn!售价), 0, RecReturn!售价) * Choose(mintUnit, 1, RecReturn!门诊包装, RecReturn!住院包装, RecReturn!药库包装), _
                        RecReturn!指导批发价 * Choose(mintUnit, 1, RecReturn!门诊包装, RecReturn!住院包装, RecReturn!药库包装), _
                        IIf(IsNull(RecReturn!产地), "!", RecReturn!产地), RecReturn!最大效期, "", _
                        Choose(mintUnit, 1, RecReturn!门诊包装, RecReturn!住院包装, RecReturn!药库包装), IIf(IsNull(RecReturn!批次), 0, RecReturn!批次), RecReturn!时价, _
                        RecReturn!药房分批, RecReturn!加成率 / 100, IIf(IsNull(RecReturn!生产日期), "", Format(RecReturn!生产日期, "yyyy-mm-dd")), _
                        RecReturn!售价单位, RecReturn!原产地
                    If .TextMatrix(.Row, mconIntCol原生产商) = "!" Then
                        .Col = mconIntCol产地
                    Else
                        .Col = mconIntCol批号
                    End If
                                            
                    If (.TextMatrix(intRow, 0) = "" Or intRow = 1 Or .Row = .rows - 1) And .TextMatrix(.rows - 1, 0) <> "" Then
                        .rows = .rows + 1
                    End If

                    .Row = .rows - 1
                    RecReturn.MoveNext
                End With
            Next
            mshBill.Row = intOldRow
            RecReturn.Close
        End If
        mshBill.Redraw = True
    Case mconIntCol产地
        Dim rsProvider As Recordset
        Dim vRect As RECT, blnCancel As Boolean
        vRect = zlControl.GetControlRect(mshBill.MsfObj.hWnd)
        
        gstrSQL = "Select 编码 as id,名称,简码 From 药品生产商 Where 站点 = [1] Or 站点 is Null Order By 编码"
        Set rsProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "产地", False, "", "", False, False, _
        True, vRect.Left + 8200, vRect.Top, 300, blnCancel, False, True, gstrNodeNo)
        
        If rsProvider Is Nothing Then
            Exit Sub
        End If
        If Not rsProvider.EOF Then
            mshBill.TextMatrix(mshBill.Row, mconIntCol产地) = rsProvider!名称
            gstrSQL = "select 批准文号 from 药品生产商对照 where 厂家名称=[1] and 药品id=[2]"
                        Set rsProvider = zlDataBase.OpenSQLRecord(gstrSQL, "mshbill_CommandClick", mshBill.TextMatrix(mshBill.Row, mconIntCol产地), mshBill.TextMatrix(mshBill.Row, 0))
            If Not rsProvider.EOF Then
                mshBill.TextMatrix(mshBill.Row, mconIntCol批准文号) = IIf(IsNull(rsProvider!批准文号), "", rsProvider!批准文号)
            Else
                mshBill.TextMatrix(mshBill.Row, mconIntCol批准文号) = ""
            End If
        End If
    Case mconIntCol原产地
        Dim vRects As RECT, blnCancels As Boolean
        vRects = zlControl.GetControlRect(mshBill.MsfObj.hWnd)
        
        gstrSQL = "Select 编码 as id,名称,简码 From 药品生产商 Where 站点 = [1] Or 站点 is Null Order By 编码"
        Set rsProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "产地", False, "", "", False, False, _
        True, vRects.Left + 9000, vRects.Top, 300, blnCancels, False, True, gstrNodeNo)
        
        If rsProvider Is Nothing Then
            Exit Sub
        End If
        If Not rsProvider.EOF Then
            mshBill.TextMatrix(mshBill.Row, mconIntCol原产地) = rsProvider!名称
        End If
    Case mconIntCol外观
        Dim rs外观 As New Recordset
                    
        gstrSQL = "Select 编码,名称,简码 From 药品外观 Order By 编码"
        Set rs外观 = zlDataBase.OpenSQLRecord(gstrSQL, "药品外观")
                
        If rs外观.EOF Then
            rs外观.Close
            Exit Sub
        End If
        With FrmSelect
            Set .TreeRec = rs外观
            .StrNode = "所有药品外观"
            .lngMode = 1
            .Show 1, Me
            If .BlnSuccess = True Then
                mshBill.TextMatrix(mshBill.Row, mconIntCol外观) = .CurrentName
                If mconIntCol外观 <> mintLastCol And mconIntCol外观 < mconintcol发票号 Then
                    mshBill.Col = mconintcol发票号
                End If
            End If
        End With
        Unload FrmSelect
    Case mconIntCol验收结论
        Dim rs验收结论 As New Recordset
                    
        gstrSQL = "Select 编码,名称 From 入库验收结论 Order By 编码"
        Set rs验收结论 = zlDataBase.OpenSQLRecord(gstrSQL, "入库验收结论")
                
        If rs验收结论.EOF Then
            rs验收结论.Close
            Exit Sub
        End If
        With FrmSelect
            Set .TreeRec = rs验收结论
            .StrNode = "所有验收结论"
            .lngMode = 1
            .Show 1, Me
            If .BlnSuccess = True Then
                mshBill.TextMatrix(mshBill.Row, mconIntCol验收结论) = .CurrentName
                If mconIntCol验收结论 <> mintLastCol And mconIntCol验收结论 < mconintcol发票号 Then
                    mshBill.Col = mconintcol发票号
                End If
            End If
        End With
        Unload FrmSelect
    End Select
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mshbill_EditChange(curText As String)
    mblnChange = True
End Sub

Private Sub mshBill_EditKeyPress(KeyAscii As Integer)
    Dim strKey As String
    Dim intDigit As Integer
    
    With mshBill
        If .Col = mconIntCol数量 Or .Col = mconIntCol冲销数量 Or .Col = mconIntCol采购价 Or .Col = mconIntCol成本价 Or .Col = mconIntCol成本金额 Or .Col = mconintcol发票金额 Or .Col = mconIntCol售价 Or .Col = mconintCol零售价 Then
            strKey = .Text
            If strKey = "" Then
                strKey = .TextMatrix(.Row, .Col)
            End If
            Select Case .Col
                Case mconIntCol数量, mconIntCol冲销数量
                    intDigit = mintNumberDigit
                Case mconIntCol采购价, mconIntCol成本价
                   intDigit = mintCostDigit
                Case mconIntCol成本金额, mconintcol发票金额
                    intDigit = mintMoneyDigit
                Case mconIntCol售价
                    intDigit = mintPriceDigit
                Case mconintCol零售价
                    intDigit = gtype_UserDrugDigits.Digit_零售价 '零售价本来是最小单位，因此按照最大单位控制显示和输入
            End Select
            
            If InStr(strKey, ".") <> 0 And Chr(KeyAscii) = "." Then   '只能存在一个小数点
                KeyAscii = 0
                Exit Sub
            End If
            
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
                If .SelLength = Len(strKey) Then Exit Sub
                If Len(Mid(strKey, InStr(1, strKey, ".") + 1)) >= intDigit And strKey Like "*.*" Then
                    KeyAscii = 0
                    Exit Sub
                Else
                    Exit Sub
                End If
            End If
        End If
    End With
End Sub

Private Sub mshbill_EnterCell(Row As Long, Col As Long)
    Dim lngRow As Long
    Dim strxq As String
    Dim dbl售价 As Double
    Dim dblLeft As Double
    Dim dblTop As Double
    
    If mint编辑状态 = 5 And Trim(mshBill.TextMatrix(mshBill.Row, mconintcol发票号)) <> "" Then
        cmdCopy.Enabled = True
    Else
        cmdCopy.Enabled = False
    End If
    
    
    If mint编辑状态 = 8 Then
        cmdGetInputCost.Visible = False
        picInputCost.Visible = False
    End If
    If Not mblnEnter Then Exit Sub
    
    If Trim(txtProvider.Text) = "" And (mint编辑状态 = 8 Or mbln退货) Then
        If mblnMSH_GetFocus Then
            mblnMSH_GetFocus = False
            MsgBox "请先选择供应商！", vbInformation, gstrSysName
        End If
        SendMessage txtProvider.hWnd, 7, 0, 0   '直接用txtprovider.setfocus会报错
        Exit Sub
    End If
    
    With mshBill
        If Row > 0 Then
            .SetRowColor CLng(Row), &HFFCECE, True
        End If
        If .Row <> .LastRow Then
            lngRow = .LastRow
            If PicInput.Visible Then
                '重新计算零售价、差价
                dbl售价 = Val(.TextMatrix(lngRow, mconIntCol成本价)) * (1 + (Val(Txt加价率) / 100))
                .TextMatrix(lngRow, mconIntCol售价) = zlStr.FormatEx(时价药品零售价(Val(.TextMatrix(lngRow, 0)), Val(.TextMatrix(lngRow, mconIntCol成本价)), Val(Txt加价率) / 100, dbl售价, lngRow), mintPriceDigit, , True)
                .TextMatrix(lngRow, mconIntCol售价金额) = zlStr.FormatEx(Val(.TextMatrix(lngRow, mconIntCol售价)) * Val(.TextMatrix(lngRow, mconIntCol数量)), mintMoneyDigit, , True)
                .TextMatrix(lngRow, mconintCol差价) = zlStr.FormatEx(IIf(.TextMatrix(lngRow, mconIntCol售价金额) = "", 0, .TextMatrix(lngRow, mconIntCol售价金额)) - IIf(.TextMatrix(lngRow, mconIntCol成本金额) = "", 0, .TextMatrix(lngRow, mconIntCol成本金额)), mintMoneyDigit, , True)
                PicInput.Visible = False
            End If
        End If
        SetInputFormat .Row
        
        'Modified by zyb 2002-10-30
        If mbln允许手工输入加成率 = False Then
            PicInput.Visible = False
        ElseIf PicInput.Visible = True Then
            If Txt加价率.Visible And Txt加价率.Enabled Then
                Txt加价率.SetFocus
            End If
            Exit Sub
        End If
                
        Select Case .Col
            Case mconIntCol药名
                .txtCheck = False
                .MaxLength = 40
                '只在药名列才显示合计信息和库存数
                Call 显示合计金额
                Call 提示库存数
                
            Case mconIntCol产地
                OS.OpenIme True
    
                .txtCheck = False
                .MaxLength = mlng生产商长度
                .TxtSetFocus
                
            Case mconIntCol原产地
                OS.OpenIme True
    
                .txtCheck = False
                .MaxLength = mlng原产地长度
                .TxtSetFocus
                
            Case mconIntCol批号
                .txtCheck = False
                '.TextMask = "1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
                .MaxLength = mintBatchNoLen
            Case mconIntCol生产日期
                .txtCheck = True
                .TextMask = "1234567890-"
                .MaxLength = 10
                If .TextMatrix(.Row, mconIntCol批号) <> "" And Len(.TextMatrix(.Row, mconIntCol批号)) = 8 Then
                    strxq = UCase(.TextMatrix(.Row, mconIntCol批号))
                    If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                        If IsNumeric(strxq) Then
                            If Trim(.TextMatrix(.Row, mconIntCol生产日期)) = "" Then
                                strxq = TranNumToDate(strxq)
                                If Trim(strxq) = "" Then Exit Sub
                                .TextMatrix(.Row, mconIntCol生产日期) = Format(strxq, "yyyy-mm-dd")
                            End If
                         End If
                    End If
                End If
            Case mconIntCol效期
                .txtCheck = True
                .TextMask = "1234567890-"
                .MaxLength = 10
                If Trim(.TextMatrix(.Row, mconIntCol原销期)) = "" Then
                    Exit Sub
                End If
                If Split(.TextMatrix(.Row, mconIntCol原销期), "||")(0) = "0" Then
                    Exit Sub
                End If
                If .TextMatrix(.Row, mconIntCol生产日期) <> "" Then
                    If Trim(.TextMatrix(.Row, mconIntCol效期)) = "" Then
                        strxq = UCase(.TextMatrix(.Row, mconIntCol生产日期))
                    End If
                ElseIf .TextMatrix(.Row, mconIntCol批号) <> "" And Len(.TextMatrix(.Row, mconIntCol批号)) = 8 Then
                    strxq = UCase(.TextMatrix(.Row, mconIntCol批号))
                    If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                        If IsNumeric(strxq) Then
                            If Trim(.TextMatrix(.Row, mconIntCol效期)) = "" Then
                                strxq = TranNumToDate(strxq)
                            Else
                                Exit Sub
                            End If
                        Else
                            strxq = ""
                        End If
                    Else
                        strxq = ""
                    End If
                End If
                If Trim(strxq) = "" Then Exit Sub
                
                .TextMatrix(.Row, mconIntCol效期) = Format(DateAdd("M", Split(.TextMatrix(.Row, mconIntCol原销期), "||")(0), strxq), "yyyy-mm-dd")
                
                If gtype_UserSysParms.P149_效期显示方式 = 1 Then
                    '换算为有效期
                    .TextMatrix(.Row, mconIntCol效期) = Format(DateAdd("D", -1, .TextMatrix(.Row, mconIntCol效期)), "yyyy-mm-dd")
                End If
                
                Call CheckLapse(.TextMatrix(.Row, mconIntCol效期))
            Case mconIntCol扣率
                .txtCheck = True
                .MaxLength = 5
                .TextMask = ".1234567890"
                staThis.Panels.Item(2) = .TextMatrix(.Row, mconIntCol药名) & "的指导批发价为：" & .TextMatrix(.Row, mconIntCol指导批发价)
                
                If mint编辑状态 = 7 Then
                    Call SetState
                End If
            Case mconIntCol成本价, mconIntCol指导批发价, mconIntCol采购价, mconintCol零售价
                .txtCheck = True
                .MaxLength = 11
                .TextMask = ".1234567890"
                If mint编辑状态 = 7 Then
                    Call SetState
                ElseIf mint编辑状态 = 8 And .Col = mconIntCol采购价 Then
                    cmdGetInputCost.Visible = True
                    dblLeft = mshBill.Left + mshBill.MsfObj.CellLeft + mshBill.MsfObj.CellWidth - cmdGetInputCost.Width + 20
                    dblTop = mshBill.Top + mshBill.MsfObj.CellTop
                    cmdGetInputCost.Top = dblTop
                    cmdGetInputCost.Left = dblLeft
                End If
            Case mconIntCol成本金额
                .txtCheck = True
                .MaxLength = 14
                .TextMask = "-.1234567890"
                If mint编辑状态 = 7 Then
                    Call SetState
                End If
            Case mconIntCol数量
                .txtCheck = True
                .MaxLength = 11
                .TextMask = ".1234567890"
            Case mconIntCol冲销数量
                .txtCheck = True
                .MaxLength = 11
                If mint编辑状态 = 6 And mbln退货 = True Then
                    .TextMask = "-.1234567890"
                Else
                    .TextMask = ".1234567890"
                End If
                
                If .TextMatrix(.Row, mconIntCol付款标志) = "是" And mint编辑状态 = 6 And gtype_UserSysParms.P173_经过标记付款后才能进行付款管理 = 1 Then
                    .ColData(mconIntCol冲销数量) = 5
                ElseIf mint编辑状态 = 6 Then
                    .ColData(mconIntCol冲销数量) = 4
                End If
            Case mconintcol发票号
                .txtCheck = False
                .MaxLength = 200
                
                If .TextMatrix(.Row, mconIntCol付款标志) = "是" And (mint编辑状态 = 5 Or mint编辑状态 = 7) And .Col = mconintcol发票号 And gtype_UserSysParms.P173_经过标记付款后才能进行付款管理 = 1 Then
                    .ColData(mconintcol发票号) = 5
                ElseIf mint编辑状态 = 5 Then
                    .ColData(mconintcol发票号) = 4
                End If
            Case mconintcol发票代码
                .txtCheck = True
                .MaxLength = 20
                .TextMask = "1234567890"
                
                If .TextMatrix(.Row, mconIntCol付款标志) = "是" And (mint编辑状态 = 5 Or mint编辑状态 = 7) And .Col = mconintcol发票代码 And gtype_UserSysParms.P173_经过标记付款后才能进行付款管理 = 1 Then
                    .ColData(mconintcol发票代码) = 5
                ElseIf mint编辑状态 = 5 And Trim(.TextMatrix(.Row, mconintcol发票号)) <> "" Then
                    .ColData(mconintcol发票代码) = 4
                End If
            Case mconintcol发票金额
                .txtCheck = True
                .MaxLength = 14
                .TextMask = "-.1234567890"
                
                If .TextMatrix(.Row, mconIntCol付款标志) = "是" And (mint编辑状态 = 5 Or mint编辑状态 = 7) And .Col = mconintcol发票金额 And gtype_UserSysParms.P173_经过标记付款后才能进行付款管理 = 1 Then
                    .ColData(mconintcol发票金额) = 5
                ElseIf mint编辑状态 = 5 Then
                    .ColData(mconintcol发票金额) = 4
                End If
            Case mconIntCol发票日期
                .txtCheck = True
                .TextMask = "1234567890-"
                .Value = Sys.Currentdate
                .MaxLength = 10
                
                If .TextMatrix(.Row, mconIntCol付款标志) = "是" And (mint编辑状态 = 5 Or mint编辑状态 = 7) And .Col = mconIntCol发票日期 And gtype_UserSysParms.P173_经过标记付款后才能进行付款管理 = 1 Then
                    .ColData(mconIntCol发票日期) = 5
                ElseIf mint编辑状态 = 5 Then
                    .ColData(mconIntCol发票日期) = 2
                End If
            Case mconIntCol外观, mconintcol产品合格证, mconIntCol验收结论
                .txtCheck = True
                .MaxLength = 100
            Case mconIntCol售价
                .txtCheck = True
                .MaxLength = 11
                .TextMask = ".1234567890"
                If mint编辑状态 = 7 Then
                    Call SetState
                End If
            Case mconIntCol批准文号
                .txtCheck = False
                .MaxLength = 40
            Case mconintcol随货单号
                .txtCheck = False
                .MaxLength = 200
            Case mconintcol随货日期
                .txtCheck = True
                .TextMask = "1234567890-"
                .Value = Sys.Currentdate
                .MaxLength = 10
                
                If .TextMatrix(.Row, mconintcol随货单号) <> "" And (mint编辑状态 = 1 Or mint编辑状态 = 2) Then
                    .ColData(mconintcol随货日期) = 2
                Else
                    .ColData(mconintcol随货日期) = 5
                End If
            End Select
    End With
End Sub

Private Sub mshBill_GotFocus()
    
    With mshBill
        
    End With
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim rsDrug As New Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim strUnitQuantity As String
    Dim dbl加成率 As Double, dbl指导零售价 As Double
    Dim dbl成本价 As Double
    Dim strxq As String
    Dim intRow As Integer
    Dim i As Integer
    Dim str药品ID As String
    Dim intOldRow As Integer
    Dim dbl售价 As Double
    Dim dblTemp售价 As Double
    Dim str药品 As String
    Dim rsProvider As Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim rsMaxs As New Recordset
    Dim ints编码 As Integer, strCodes As String
                    
    intOldRow = mshBill.Row
        
    If KeyCode <> vbKeyReturn Then Exit Sub
    On Error GoTo errHandle
    
    With mshBill
'        .Text = UCase(Trim(.Text))
        strKey = Trim(.Text)
        
        If Mid(strKey, 1, 1) = "[" Then
            If InStr(2, strKey, "]") <> 0 Then
                strKey = Mid(strKey, 2, InStr(2, strKey, "]") - 2)
            Else
                strKey = Mid(strKey, 2)
            End If
        End If
        Select Case .Col
            Case mconIntCol药名
                If strKey <> "" Then
                    Dim RecReturn As Recordset
                    Dim sngLeft As Single
                    Dim sngTop As Single
                    
                    .Redraw = False
                    sngLeft = Me.Left + Pic单据.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                    sngTop = Me.Top + Me.Height - Me.ScaleHeight + Pic单据.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                    If sngTop + 3630 > Screen.Height Then
                        sngTop = sngTop - mshBill.MsfObj.CellHeight - 3630
                    End If
                    
                    If grsMaster.State = adStateClosed Then '获取数据集
                        Call SetSelectorRS(IIf(mint编辑状态 = 8 Or mbln退货, 2, 1), MStrCaption, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , IIf(mint编辑状态 = 8 Or mbln退货, Val(txtProvider.Tag), 0))
                    End If
'                    Set RecReturn = Frm药品多选选择器.ShowME(Me, IIf(mint编辑状态 = 8 Or mbln退货, 2, 1), cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , strkey, sngLeft, sngTop, True, True, False, False, True, IIf(mint编辑状态 = 8 Or mbln退货, Val(txtProvider.Tag), 0))
                    Set RecReturn = frmSelector.ShowME(Me, 1, IIf(mint编辑状态 = 8 Or mbln退货, 2, 1), strKey, sngLeft, sngTop, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , IIf(mint编辑状态 = 8 Or mbln退货, Val(txtProvider.Tag), 0), True, True, True, , , mstrPrivs)
                                                        
                    If RecReturn.RecordCount > 0 And (mint编辑状态 = 8 Or mbln退货 = True) Then
                        Set RecReturn = CheckRedo(RecReturn) '检查重复记录 并将重复记录的药品id返回回来
                    End If
'                    If str药品id <> "" And mint编辑状态 = 8 Then
'                        mbln提示 = False
'                        Set RecReturn = GetRs(str药品id, RecReturn) '过滤重复的数据
'                    End If
                                                        
                    If RecReturn.RecordCount > 0 Then
                        RecReturn.MoveFirst
                        
                        For i = 1 To RecReturn.RecordCount
                            mlng包装系数 = Choose(mintUnit, 1, RecReturn!门诊包装, RecReturn!住院包装, RecReturn!药库包装)
                            
                            intRow = .Row
                            .TextMatrix(intRow, mconIntCol行号) = .Row
                            If SetColValue(.Row, RecReturn!药品ID, "[" & RecReturn!药品编码 & "]", RecReturn!通用名, IIf(IsNull(RecReturn!商品名), "", RecReturn!商品名), _
                                 nvl(RecReturn!药品来源), "" & RecReturn!基本药物, IIf(IsNull(RecReturn!规格), "", RecReturn!规格), _
                                 IIf(IsNull(RecReturn!产地), "", RecReturn!产地), Choose(mintUnit, RecReturn!售价单位, RecReturn!门诊单位, RecReturn!住院单位, RecReturn!药库单位), _
                                 IIf(IsNull(RecReturn!售价), 0, RecReturn!售价) * Choose(mintUnit, 1, RecReturn!门诊包装, RecReturn!住院包装, RecReturn!药库包装), _
                                 RecReturn!指导批发价 * Choose(mintUnit, 1, RecReturn!门诊包装, RecReturn!住院包装, RecReturn!药库包装), _
                                 IIf(IsNull(RecReturn!产地), "!", RecReturn!产地), RecReturn!最大效期, "", _
                                 Choose(mintUnit, 1, RecReturn!门诊包装, RecReturn!住院包装, RecReturn!药库包装), IIf(IsNull(RecReturn!批次), 0, RecReturn!批次), RecReturn!时价, _
                                 RecReturn!药房分批, RecReturn!加成率 / 100, IIf(IsNull(RecReturn!生产日期), "", Format(RecReturn!生产日期, "yyyy-mm-dd")), RecReturn!售价单位, RecReturn!原产地) = False Then ' RecReturn!简码
                                 Cancel = True
                                 Exit Sub
                             End If
                            .Text = .TextMatrix(.Row, .Col)
                            If (.TextMatrix(intRow, 0) = "" Or intRow = 1 Or .Row = .rows - 1) And .TextMatrix(.rows - 1, 0) <> "" Then
                                .rows = .rows + 1
                            End If
                            .Row = .rows - 1
                            RecReturn.MoveNext
                        Next
                        .Row = intOldRow
                    Else
                        Cancel = True
                    End If
                    Call 提示库存数
                End If
                .Redraw = True
            Case mconIntCol产地
                '无处理
                If Trim(.Text) = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mconIntCol产地) = ""
                    End If
                    
                    If mconIntCol产地 <> mintLastCol And mconIntCol产地 < mconIntCol批号 Then
'                        .Col = mconIntCol批号
                        .Col = GetNextEnableCol(mconIntCol产地)
                        Cancel = True
                    End If
                    Exit Sub
                Else
                    vRect = zlControl.GetControlRect(mshBill.MsfObj.hWnd)
                    
                    .Text = UCase(Trim(.Text))
                    strKey = Trim(.Text)
                    
                    If Trim(.Text) = "" Then Exit Sub
                    
                    gstrSQL = "Select 编码 as id,编码 ,名称,简码 From 药品生产商 " _
                            & "Where (站点 = '" & gstrNodeNo & "' Or 站点 is Null) And (upper(名称) like '" & IIf(gstrMatchMethod = "0", "%", "") & strKey & "%' or Upper(编码) like '" & IIf(gstrMatchMethod = "0", "%", "") & strKey & "%' or Upper(简码) like '" & strKey & "%') " _
                                & "Order By 编码 "
                                
                    Set rsProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "生产商", False, "", "生产商选择", False, False, _
                    True, vRect.Left + 8200, vRect.Top, 300, blnCancel, False, True, gstrNodeNo)
                    
                    If blnCancel = True Then mshBill.Text = "": .TextMatrix(.Row, mconIntCol产地) = "": Exit Sub '打开选择器时，点Esc不做以下处理
                    
                    If rsProvider Is Nothing Then
                        If MsgBox("药品生产商没有找到你输入的生产商，你要把它加入药品生产商中吗？", vbYesNo + vbQuestion, MStrCaption) = vbNo Then
                            mshBill.TextMatrix(mshBill.Row, mconIntCol产地) = ""
                            mshBill.Text = ""
                            Cancel = True
                            Exit Sub
                        Else
                            If LenB(StrConv(strKey, vbFromUnicode)) > mlng生产商长度 Then
                                MsgBox "生产商名称过长(最多" & mlng生产商长度 & "个字符或" & Int(mlng生产商长度 / 2) & "个汉字)!", vbInformation + vbOKOnly, gstrSysName
                                Cancel = True
                                Exit Sub
                            End If
                            
                            If rsMaxs.State = 1 Then rsMaxs.Close
                            gstrSQL = "SELECT Nvl(MAX(LENGTH(编码)),2) As Length FROM 药品生产商"
                            Set rsMaxs = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "-药品生产商编码长度")
                            ints编码 = rsMaxs!length
                            
                            gstrSQL = "SELECT Nvl(MAX(LPAD(编码," & ints编码 & ",'0')),'00') As Code FROM 药品生产商"
                            Set rsMaxs = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "-药品生产商编码")
                            strCodes = rsMaxs!Code
                            
                            ints编码 = Len(strCodes)
                            strCodes = strCodes + 1
                            If ints编码 >= Len(strCodes) Then
                                strCodes = String(ints编码 - Len(strCodes), "0") & strCodes
                            End If

                            gstrSQL = "ZL_药品生产商_INSERT('" & strCodes & "','" & strKey & "',zlSpellCode('" & strKey & "',10))"
                            
                            Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
                        End If
                    Else
                        mshBill.TextMatrix(mshBill.Row, mconIntCol产地) = rsProvider!名称
                        mshBill.Text = rsProvider!名称
                        
                        gstrSQL = "select 批准文号 from 药品生产商对照 where 厂家名称=[1] and 药品id=[2]"
                        Set rsProvider = zlDataBase.OpenSQLRecord(gstrSQL, "mshbill_CommandClick", mshBill.TextMatrix(mshBill.Row, mconIntCol产地), mshBill.TextMatrix(mshBill.Row, 0))
                        If Not rsProvider.EOF Then
                            mshBill.TextMatrix(mshBill.Row, mconIntCol批准文号) = IIf(IsNull(rsProvider!批准文号), "", rsProvider!批准文号)
                        Else
                            mshBill.TextMatrix(mshBill.Row, mconIntCol批准文号) = ""
                        End If
                    End If
                End If
                OS.OpenIme
            Case mconIntCol原产地
                '无处理
                If Trim(.Text) = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mconIntCol原产地) = ""
                    End If
                    
                    If mconIntCol原产地 <> mintLastCol And mconIntCol原产地 < mconIntCol批号 Then
'                        .Col = mconIntCol批号
                        .Col = GetNextEnableCol(mconIntCol原产地)
                        Cancel = True
                    End If
                    Exit Sub
                Else
                
                    vRect = zlControl.GetControlRect(mshBill.MsfObj.hWnd)
                    .Text = UCase(Trim(.Text))
                    strKey = Trim(.Text)
                    
                    gstrSQL = "Select 编码 as id,编码 ,名称,简码 From 药品生产商 " _
                            & "Where (站点 = '" & gstrNodeNo & "' Or 站点 is Null) And (upper(名称) like '" & IIf(gstrMatchMethod = "0", "%", "") & strKey & "%' or Upper(编码) like '" & IIf(gstrMatchMethod = "0", "%", "") & strKey & "%' or Upper(简码) like '" & strKey & "%') " _
                                & "Order By 编码 "
                                
                    Set rsProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "原产地", False, "", "原产地选择", False, False, _
                    True, vRect.Left + 9000, vRect.Top, 300, blnCancel, False, True, gstrNodeNo)
                    
                    If blnCancel = True Then .Text = "": .TextMatrix(.Row, mconIntCol原产地) = "": Exit Sub '打开选择器时，点Esc不做以下处理
                    
                    If rsProvider Is Nothing Then
                        If MsgBox("药品生产商没有找到你输入的原产地，你要把它加入药品生产商中吗？", vbYesNo + vbQuestion, MStrCaption) = vbNo Then
                            mshBill.TextMatrix(mshBill.Row, mconIntCol原产地) = ""
                            mshBill.Text = ""
                            Cancel = True
                            Exit Sub
                        Else
                            If LenB(StrConv(strKey, vbFromUnicode)) > mlng原产地长度 Then
                                MsgBox "原产地名称过长(最多" & mlng原产地长度 & "个字符或" & Int(mlng原产地长度 / 2) & "个汉字)!", vbInformation + vbOKOnly, gstrSysName
                                Cancel = True
                                Exit Sub
                            End If
                        
                            If rsMaxs.State = 1 Then rsMaxs.Close
                            gstrSQL = "SELECT Nvl(MAX(LENGTH(编码)),2) As Length FROM 药品生产商"
                            Set rsMaxs = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "-药品生产商编码长度")
                            ints编码 = rsMaxs!length
                            
                            gstrSQL = "SELECT Nvl(MAX(LPAD(编码," & ints编码 & ",'0')),'00') As Code FROM 药品生产商"
                            Set rsMaxs = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "-药品生产商编码")
                            strCodes = rsMaxs!Code
                            
                            ints编码 = Len(strCodes)
                            strCodes = strCodes + 1
                            If ints编码 >= Len(strCodes) Then
                                strCodes = String(ints编码 - Len(strCodes), "0") & strCodes
                            End If
                            
                            gstrSQL = "ZL_药品生产商_INSERT('" & strCodes & "','" & strKey & "',zlSpellCode('" & strKey & "',10))"
 
                            Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
                        End If
                    Else
                        mshBill.TextMatrix(mshBill.Row, mconIntCol原产地) = rsProvider!名称
                        mshBill.Text = rsProvider!名称
                    End If
                End If
                OS.OpenIme
            Case mconIntCol验收结论
                If Trim(.TextMatrix(.Row, mconIntCol药名)) = "" Then
                    .Text = ""
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If .Text = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mconIntCol验收结论) = ""
                    End If
                    .Col = GetNextEnableCol(mconIntCol验收结论)
                    Cancel = True
                    Exit Sub
                Else
                    Dim rs结论 As New Recordset
                    
                    gstrSQL = "" & _
                        "   Select 编码,名称 From 入库验收结论 " & _
                        "   Where upper(名称) like [1] or Upper(编码) like [1] "
                    Set rs结论 = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, IIf(gstrMatchMethod = "0", "%", "") & strKey & "%")
                    
                    If rs结论.EOF Then
                        .TextMatrix(.Row, mconIntCol验收结论) = .Text
                        .Col = GetNextEnableCol(mconIntCol验收结论)
                        Cancel = True
                        Exit Sub
                    Else
                        If rs结论.RecordCount = 1 Then
                            .TextMatrix(.Row, mconIntCol验收结论) = rs结论.Fields("名称")
                            .Text = rs结论.Fields("名称")
                            .Col = GetNextEnableCol(mconIntCol验收结论)
                        Else
                            Set msh产地.Recordset = rs结论
                            With msh产地
                                .Redraw = False
                                .Left = Pic单据.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                                .Top = Pic单据.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight
                                .Visible = True
                                .SetFocus
                                .ColWidth(0) = 800
                                .ColWidth(1) = 5000
                                .Row = 1
                                .Col = 0
                                .TopRow = 1
                                .ColSel = .Cols - 1
                                .Redraw = True
                                Cancel = True
                                Exit Sub
                            End With
                        End If
                    End If
                End If
                OS.OpenIme
            Case mconIntCol批号
                '无处理
                If strKey = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mconIntCol批号) = ""
                    End If
                    If mconIntCol批号 <> mintLastCol And mconIntCol批号 < mconIntCol生产日期 Then
                        .Col = mconIntCol生产日期
                        Cancel = True
                    End If
                    Exit Sub
                End If
            Case mconIntCol生产日期
                '有处理
                If strKey <> "" Then
                    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
                        strKey = TranNumToDate(strKey)
                        If strKey = "" Then
                            MsgBox "对不起，生产日期必须为日期型！", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        .Text = strKey
                        .TextMatrix(.Row, mconIntCol生产日期) = .Text
                        
                        '设置效期
                        If Trim(.TextMatrix(.Row, mconIntCol原销期)) = "" Then
                            Exit Sub
                        End If
                        If Split(.TextMatrix(.Row, mconIntCol原销期), "||")(0) = "0" Then
                            Exit Sub
                        End If
                        If .TextMatrix(.Row, mconIntCol生产日期) <> "" Then
                            strxq = UCase(.TextMatrix(.Row, mconIntCol生产日期))
                        ElseIf .TextMatrix(.Row, mconIntCol批号) <> "" And Len(.TextMatrix(.Row, mconIntCol批号)) = 8 Then
                            strxq = UCase(.TextMatrix(.Row, mconIntCol批号))
                            If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                                If IsNumeric(strxq) Then
                                    If Trim(.TextMatrix(.Row, mconIntCol效期)) = "" Then
                                        strxq = TranNumToDate(strxq)
                                    Else
                                        Exit Sub
                                    End If
                                Else
                                    strxq = ""
                                End If
                            Else
                                strxq = ""
                            End If
                        End If
                        If Trim(strxq) = "" Then Exit Sub
                        
                        .TextMatrix(.Row, mconIntCol效期) = Format(DateAdd("M", Split(.TextMatrix(.Row, mconIntCol原销期), "||")(0), strxq), "yyyy-mm-dd")
                        
                        If gtype_UserSysParms.P149_效期显示方式 = 1 Then
                            '换算为有效期
                            .TextMatrix(.Row, mconIntCol效期) = Format(DateAdd("D", -1, .TextMatrix(.Row, mconIntCol效期)), "yyyy-mm-dd")
                        End If
                        
                        Call CheckLapse(.TextMatrix(.Row, mconIntCol效期))
                        Exit Sub
                    End If
                    If Not IsDate(strKey) Then
                        MsgBox "对不起，生产日期必须为日期型如(2000-10-10) 或（20001010）,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                ElseIf strKey = "" And strKey <> .TextMatrix(.Row, mconIntCol生产日期) Then
                    If .TxtVisible = True Then
                        .Text = " "
                        Exit Sub
                    End If
                    If .ColData(mconIntCol效期) = 2 Then
                        If mconIntCol生产日期 <> mintLastCol And mconIntCol生产日期 < mconIntCol效期 Then
                            .Col = mconIntCol效期
                        End If
                    Else
                        If mconIntCol生产日期 <> mintLastCol And mconIntCol生产日期 < mconIntCol数量 Then
                            .Col = mconIntCol数量
                        End If
                    End If
                    Exit Sub
                End If
            Case mconIntCol效期
                '有处理
                If strKey <> "" Then
                    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
                        strKey = TranNumToDate(strKey)
                        If strKey = "" Then
                            MsgBox "对不起，失效期必须为日期型！", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        .Text = strKey
                        Exit Sub
                    End If
                    If Not IsDate(strKey) Then
                        MsgBox "对不起，失效期必须为日期型如(2000-10-10) 或（20001010）,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                ElseIf strKey = "" And strKey <> .TextMatrix(.Row, mconIntCol效期) Then
                
                    If .TxtVisible = True Then
                        .Text = " "
                        Exit Sub
                    End If
                    
                    Exit Sub
                End If
            Case mconIntCol扣率
                If Trim(.TextMatrix(.Row, mconIntCol药名)) = "" Then
                    .Text = ""
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "对不起，扣率必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" And strKey <> .TextMatrix(.Row, mconIntCol扣率) Then
                    SetDisCount .Row, strKey
                End If
                
                Call 检查成本价
                Call 显示合计金额
                Call Set时价分批药品零售价(.Row, Val(.TextMatrix(.Row, mconIntCol售价)) / IIf(Val(.TextMatrix(.Row, mconIntCol比例系数)) = 0, 1, Val(.TextMatrix(.Row, mconIntCol比例系数))))
            Case mconIntCol指导批发价
                If .TxtVisible Then
                    If Not IsNumeric(strKey) And strKey <> "" Then
                        MsgBox "对不起，采购限价必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    If mintUnit = mconint药库单位 Then
                        If Val(strKey) < 0.01 Then
                            MsgBox "对不起，采购限价必须大于0.01,请重输！", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    Else
                        If Val(strKey) < 0.001 Then
                            MsgBox "对不起，采购限价必须大于0.001,请重输！", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                    
                    If strKey <> "" And strKey <> .TextMatrix(.Row, mconIntCol指导批发价) Then
                        strKey = zlStr.FormatEx(strKey, mintCostDigit, , True)
                        .Text = strKey
                        'Modifed by ZYB 20050104
                        .TextMatrix(.Row, mconIntCol指导批发价) = .Text
                        'Modifed by ZYB 20050104 END
                        SetDisCount .Row, strKey
                    End If
                    
                    Call 检查成本价
                    Call 显示合计金额
                End If
            Case mconIntCol采购价
                If Trim(.TextMatrix(.Row, mconIntCol药名)) = "" Then
                    .Text = ""
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "对不起，采购价必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Val(strKey) < 0 Then
                        MsgBox "对不起，采购价不能为负数,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strKey) >= 10 ^ 11 - 1 Then
                        MsgBox "采购价必须小于" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    .Text = zlStr.FormatEx(strKey, 7, , True)
                    .TextMatrix(.Row, .Col) = .Text
                End If
                
                If Val(strKey) > Val(.TextMatrix(.Row, mconIntCol指导批发价)) Then
                    MsgBox "你输入的采购价大于了采购限价。", vbInformation + vbOKOnly, gstrSysName
                End If
                
                '返回设置扣率
                If strKey <> "" Then
                    strKey = zlStr.FormatEx(strKey, mintCostDigit, , True)
                    .Text = strKey
                    .TextMatrix(.Row, mconIntCol采购价) = .Text
                End If
               
                '计算成本价：成本价=采购价*扣率
                .TextMatrix(.Row, mconIntCol成本价) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol采购价)) * Val(.TextMatrix(.Row, mconIntCol扣率)) / 100, mintCostDigit, , True)
                
                '对时价药品的处理
                If strKey <> "" And .TextMatrix(.Row, mconIntCol原销期) <> "" And mint编辑状态 <> 8 And mbln退货 = False Then
                    '根据参数决定时价药品售价公式中成本价的算法
                    dbl成本价 = IIf(mint时价入库售价加成方式 = 0, Val(.TextMatrix(.Row, mconIntCol成本价)), Val(.TextMatrix(.Row, mconIntCol采购价)))
                    
                    If Split(.TextMatrix(.Row, mconIntCol原销期), "||")(2) = 1 Then
                        '零差价控制：时价药品，售价直接等于成本价
                        If gtype_UserSysParms.P275_零差价管理模式 = 2 And IsPriceAdjustMod(Val(.TextMatrix(.Row, 0))) = True Then
                            .TextMatrix(.Row, mconIntCol售价) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol成本价)), mintPriceDigit, , True)
                            If .TextMatrix(.Row, mconIntCol数量) <> "" Then
                                .TextMatrix(.Row, mconIntCol售价金额) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol数量) * .TextMatrix(.Row, mconIntCol售价), mintMoneyDigit, , True)
                            End If
                        Else
                            '由于存在差价让利比的存在,需要按加成率计算,因此将指导差价率转换成加成率计算 公式：加成率=1/(1-差价率)-1
                            '如果系统参数为真，则提示用户输入加价率
                            If mbln加价率 And mint时价入库售价加成方式 = 1 Then
                                mbln允许手工输入加成率 = True
                                sngLeft = Pic单据.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                                sngTop = Pic单据.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                                If sngTop + 1700 > Screen.Height Then
                                    sngTop = sngTop - mshBill.MsfObj.CellHeight - 1700
                                End If
                                
                                With PicInput
                                    .Top = sngTop
                                    .Left = sngLeft
                                    .Visible = True
                                End With
                                 Txt加价率 = Val(Replace(.TextMatrix(.Row, mconIntcol加成率), "%", "")) '"15.00000"
                                 .TextMatrix(.Row, mconIntCol售价) = zlStr.FormatEx(时价药品零售价(Val(.TextMatrix(.Row, 0)), dbl成本价, Val(Txt加价率) / 100, dbl成本价 * (1 + (Val(Txt加价率) / 100))), mintPriceDigit)
                                 
'                                If Val(.TextMatrix(.Row, mconIntCol售价)) <> 0 And dbl成本价 <> 0 Then
'                                    Txt加价率 = zlStr.FormatEx(计算加成率(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol售价)), dbl成本价), 5, , True)
'                                End If
                                Txt加价率.Tag = Txt加价率
                                Txt加价率.SetFocus
                            Else
                                If mint时价分段加成方式 = 1 Then
                                    If get分段加成售价(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol比例系数)), dbl成本价, dbl加成率, dblTemp售价) = False Then
                                        Cancel = True
                                        .TxtSetFocus
                                        Exit Sub
                                    End If
                                Else
                                    dbl加成率 = Val(Replace(.TextMatrix(.Row, mconIntcol加成率), "%", "")) / 100
                                    dblTemp售价 = dbl成本价 * (1 + dbl加成率)
                                End If
                                If gtype_UserSysParms.P183_时价取上次售价 <> 1 Then  '没有勾选时价取上次售价参数
                                    .TextMatrix(.Row, mconIntCol售价) = zlStr.FormatEx(时价药品零售价(Val(.TextMatrix(.Row, 0)), dbl成本价, dbl加成率, dblTemp售价), mintPriceDigit, , True)
                                End If
                                
                                .TextMatrix(.Row, mconIntcol加成率) = zlStr.FormatEx(dbl加成率 * 100, 2) & "%"
                                If .TextMatrix(.Row, mconIntCol数量) <> "" Then
                                    .TextMatrix(.Row, mconIntCol售价金额) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol数量) * .TextMatrix(.Row, mconIntCol售价), mintMoneyDigit, , True)
                                End If
                            End If
                        End If
                    Else
                        '定价药品计算加成率，无实际意义，仅显示
                        If Val(.TextMatrix(.Row, mconIntCol售价)) <> 0 And dbl成本价 <> 0 Then
                            .TextMatrix(.Row, mconIntcol加成率) = zlStr.FormatEx((Val(.TextMatrix(.Row, mconIntCol售价)) / dbl成本价 - 1) * 100, 2) & "%"
                        End If
                        
                        '零差价控制：定价药品，检查录入的成本价是否等于售价
                        If gtype_UserSysParms.P275_零差价管理模式 = 2 And Val(.TextMatrix(.Row, mconIntCol成本价)) <> Val(.TextMatrix(.Row, mconIntCol售价)) Then
                            If IsPriceAdjustMod(Val(.TextMatrix(.Row, 0))) = True Then
                                MsgBox "该定价药品已启用零差价管理模式，入库成本价应和售价(" & zlStr.FormatEx(.TextMatrix(.Row, mconIntCol售价), mintPriceDigit, , True) & ")相等！", vbInformation + vbOKOnly, gstrSysName
                                
                                .TextMatrix(.Row, mconIntCol成本价) = .TextMatrix(.Row, mconIntCol售价)
                                strKey = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol成本价)) / (Val(.TextMatrix(.Row, mconIntCol扣率)) / 100), mintPriceDigit, , True)
                                .TextMatrix(.Row, mconIntCol采购价) = strKey
                                .Text = strKey
                                .TextMatrix(.Row, mconIntcol加成率) = "0%"
                            End If
                        End If
                    End If
                End If
                
                '退货时
                If strKey <> "" And (mint编辑状态 = 8 Or mbln退货 = True) Then
                    If .TextMatrix(.Row, mconIntCol原销期) <> "" Then
                        If Split(.TextMatrix(.Row, mconIntCol原销期), "||")(2) = 1 Then
                            '零差价控制：时价药品，售价直接等于成本价
                            If gtype_UserSysParms.P275_零差价管理模式 = 2 And IsPriceAdjustMod(Val(.TextMatrix(.Row, 0))) = True Then
                                .TextMatrix(.Row, mconIntCol售价) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol成本价)), mintPriceDigit, , True)
                                If .TextMatrix(.Row, mconIntCol数量) <> "" Then
                                    .TextMatrix(.Row, mconIntCol售价金额) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol数量) * .TextMatrix(.Row, mconIntCol售价), mintMoneyDigit, , True)
                                End If
                            End If
                        Else
                            '零差价控制：定价药品，检查录入的成本价是否等于售价
                            If gtype_UserSysParms.P275_零差价管理模式 = 2 And Val(.TextMatrix(.Row, mconIntCol成本价)) <> Val(.TextMatrix(.Row, mconIntCol售价)) Then
                                If IsPriceAdjustMod(Val(.TextMatrix(.Row, 0))) = True Then
                                MsgBox "该定价药品已启用零差价管理模式，入库成本价应和售价(" & zlStr.FormatEx(.TextMatrix(.Row, mconIntCol售价), mintPriceDigit, , True) & ")相等！", vbInformation + vbOKOnly, gstrSysName
                                
                                .TextMatrix(.Row, mconIntCol成本价) = .TextMatrix(.Row, mconIntCol售价)
                                strKey = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol成本价)) / (Val(.TextMatrix(.Row, mconIntCol扣率)) / 100), mintPriceDigit, , True)
                                .TextMatrix(.Row, mconIntCol采购价) = strKey
                                .Text = strKey
                                .TextMatrix(.Row, mconIntcol加成率) = "0%"
                                End If
                            End If
                        End If
                    End If
                 End If
                
                '设置金额
                If strKey <> "" And .TextMatrix(.Row, mconIntCol数量) <> "" Then
                    .TextMatrix(.Row, mconIntCol成本金额) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol数量) * Val(.TextMatrix(.Row, mconIntCol成本价)), mintMoneyDigit, , True)
                    .TextMatrix(.Row, mconIntCol售价金额) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol数量) * Val(.TextMatrix(.Row, mconIntCol售价)), mintMoneyDigit, , True)
                    .TextMatrix(.Row, mconintcol发票金额) = IIf(Trim(.TextMatrix(.Row, mconintcol发票号)) = "", "", .TextMatrix(.Row, mconIntCol成本金额))
                    .TextMatrix(.Row, mconintCol差价) = zlStr.FormatEx(IIf(.TextMatrix(.Row, mconIntCol售价金额) = "", 0, .TextMatrix(.Row, mconIntCol售价金额)) - IIf(.TextMatrix(.Row, mconIntCol成本金额) = "", 0, .TextMatrix(.Row, mconIntCol成本金额)), mintMoneyDigit, , True)
                End If
                
                Call 检查成本价
                Call 显示合计金额
                Call Set时价分批药品零售价(.Row, Val(.TextMatrix(.Row, mconIntCol售价)) / Val(.TextMatrix(.Row, mconIntCol比例系数)))
            Case mconIntCol成本价
                If Trim(.TextMatrix(.Row, mconIntCol药名)) = "" Then
                    .Text = ""
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "对不起，成本价必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Val(strKey) < 0 Then
                        MsgBox "对不起，成本价不能为负数,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strKey) >= 10 ^ 11 - 1 Then
                        MsgBox "成本价必须小于" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    .Text = zlStr.FormatEx(strKey, 7, , True)
                    .TextMatrix(.Row, .Col) = .Text
                End If
                
                If strKey <> "" Then
                    strKey = zlStr.FormatEx(strKey, mintCostDigit, , True)
                    .Text = strKey
                    .TextMatrix(.Row, mconIntCol成本价) = .Text
                End If
                          
                If Val(strKey) > Val(.TextMatrix(.Row, mconIntCol指导批发价)) Then
                    MsgBox "你输入的成本价大于了采购限价。", vbInformation + vbOKOnly, gstrSysName
                End If
                           
                If Val(.TextMatrix(.Row, mconIntCol扣率)) = 0 Then
                    .TextMatrix(.Row, mconIntCol扣率) = "100"
                End If
                
                '计算扣率：扣率=成本价/采购价
                If Val(.TextMatrix(.Row, mconIntCol采购价)) <> 0 Then
                    .TextMatrix(.Row, mconIntCol扣率) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol成本价)) / Val(.TextMatrix(.Row, mconIntCol采购价)) * 100, 7, , True)
                Else
                    .TextMatrix(.Row, mconIntCol扣率) = "100"
                End If
                
                '对时价药品的处理
                If strKey <> "" And .TextMatrix(.Row, mconIntCol原销期) <> "" And mint编辑状态 <> 8 And mbln退货 = False Then
                    '根据参数决定时价药品售价公式中成本价的算法
                    dbl成本价 = IIf(mint时价入库售价加成方式 = 0, Val(.TextMatrix(.Row, mconIntCol成本价)), Val(.TextMatrix(.Row, mconIntCol采购价)))
                        
                    If Split(.TextMatrix(.Row, mconIntCol原销期), "||")(2) = 1 Then
                        '零差价控制：时价药品，售价直接等于成本价
                        If gtype_UserSysParms.P275_零差价管理模式 = 2 And IsPriceAdjustMod(Val(.TextMatrix(.Row, 0))) = True Then
                            .TextMatrix(.Row, mconIntCol售价) = zlStr.FormatEx(strKey, mintPriceDigit, , True)
                            If .TextMatrix(.Row, mconIntCol数量) <> "" Then
                                .TextMatrix(.Row, mconIntCol售价金额) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol数量) * .TextMatrix(.Row, mconIntCol售价), mintMoneyDigit, , True)
                            End If
                            
                            .TextMatrix(.Row, mconIntcol加成率) = "0%"
                        Else
                            '由于存在差价让利比的存在,需要按加成率计算,因此将指导差价率转换成加成率计算 公式：加成率=1/(1-差价率)-1
                            '如果系统参数为真，则提示用户输入加价率
                            If mbln加价率 And mint时价入库售价加成方式 = 0 Then
                                mbln允许手工输入加成率 = True
                                If gtype_UserSysParms.P183_时价取上次售价 <> 1 Then  '如果未勾选取上次售价，且勾选了手工录入加成率参数则弹出加成率框，让用户选择
                                    sngLeft = Pic单据.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                                    sngTop = Pic单据.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                                    If sngTop + 1700 > Screen.Height Then
                                        sngTop = sngTop - mshBill.MsfObj.CellHeight - 1700
                                    End If
                                    
                                    With PicInput
                                        .Top = sngTop
                                        .Left = sngLeft
                                        .Visible = True
                                    End With
                                    Txt加价率 = Val(Replace(.TextMatrix(.Row, mconIntcol加成率), "%", "")) '"15.00000"
                                    .TextMatrix(.Row, mconIntCol售价) = zlStr.FormatEx(时价药品零售价(Val(.TextMatrix(.Row, 0)), dbl成本价, Val(Txt加价率) / 100, dbl成本价 * (1 + (Val(Txt加价率) / 100))), mintPriceDigit)


'                                    If Val(.TextMatrix(.Row, mconIntCol售价)) <> 0 And dbl成本价 <> 0 Then
'                                        Txt加价率 = zlStr.FormatEx(计算加成率(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol售价)), dbl成本价), 5, , True)
'                                    End If
                                    Txt加价率.Tag = Txt加价率
                                    Txt加价率.SetFocus
                                End If
                            Else
                                If mint时价分段加成方式 = 1 Then
                                    If get分段加成售价(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol比例系数)), dbl成本价, dbl加成率, dblTemp售价) = False Then
                                        Cancel = True
                                        .TxtSetFocus
                                        Exit Sub
                                    End If
                                Else
                                    dbl加成率 = Val(Replace(.TextMatrix(.Row, mconIntcol加成率), "%", "")) / 100
                                    dblTemp售价 = dbl成本价 * (1 + dbl加成率)
                                End If
                                
                                If gtype_UserSysParms.P183_时价取上次售价 <> 1 Then  '没有勾选时价取上次售价参数
                                    .TextMatrix(.Row, mconIntCol售价) = zlStr.FormatEx(时价药品零售价(Val(.TextMatrix(.Row, 0)), dbl成本价, dbl加成率, dblTemp售价), mintPriceDigit, , True)
                                End If
                                .TextMatrix(.Row, mconIntcol加成率) = zlStr.FormatEx(dbl加成率 * 100, 2) & "%"
                                If .TextMatrix(.Row, mconIntCol数量) <> "" Then
                                    .TextMatrix(.Row, mconIntCol售价金额) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol数量) * .TextMatrix(.Row, mconIntCol售价), mintMoneyDigit, , True)
                                End If
                            End If
                        End If
                    Else
                        '定价药品计算加成率，无实际意义，仅显示
                        If Val(.TextMatrix(.Row, mconIntCol售价)) <> 0 And dbl成本价 <> 0 Then
                            .TextMatrix(.Row, mconIntcol加成率) = zlStr.FormatEx((Val(.TextMatrix(.Row, mconIntCol售价)) / dbl成本价 - 1) * 100, 2) & "%"
                        End If
                        
                        '零差价控制：定价药品，检查录入的成本价是否等于售价
                        If gtype_UserSysParms.P275_零差价管理模式 = 2 And Val(strKey) <> Val(.TextMatrix(.Row, mconIntCol售价)) Then
                            If IsPriceAdjustMod(Val(.TextMatrix(.Row, 0))) = True Then
                                MsgBox "该定价药品已启用零差价管理模式，入库成本价应和售价(" & zlStr.FormatEx(.TextMatrix(.Row, mconIntCol售价), mintPriceDigit, , True) & ")相等！", vbInformation + vbOKOnly, gstrSysName
                                strKey = .TextMatrix(.Row, mconIntCol售价)
                                .TextMatrix(.Row, mconIntCol成本价) = zlStr.FormatEx(strKey, mintPriceDigit, , True)
                                .Text = strKey
                                .TextMatrix(.Row, mconIntcol加成率) = "0%"
                                .TextMatrix(.Row, mconIntCol扣率) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol成本价)) / Val(.TextMatrix(.Row, mconIntCol采购价)) * 100, 7, , True)
'                                Cancel = True
'                                .TxtSetFocus
'                                Exit Sub
                            End If
                        End If
                    End If
                End If
                
                '退货时
                If strKey <> "" And (mint编辑状态 = 8 Or mbln退货 = True) Then
                    If .TextMatrix(.Row, mconIntCol原销期) <> "" Then
                        If Split(.TextMatrix(.Row, mconIntCol原销期), "||")(2) = 1 Then
                            '零差价控制：时价药品，售价直接等于成本价
                            If gtype_UserSysParms.P275_零差价管理模式 = 2 And IsPriceAdjustMod(Val(.TextMatrix(.Row, 0))) = True Then
                                .TextMatrix(.Row, mconIntCol售价) = zlStr.FormatEx(strKey, mintPriceDigit, , True)
                                If .TextMatrix(.Row, mconIntCol数量) <> "" Then
                                    .TextMatrix(.Row, mconIntCol售价金额) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol数量) * .TextMatrix(.Row, mconIntCol售价), mintMoneyDigit, , True)
                                End If

                                .TextMatrix(.Row, mconIntcol加成率) = "0%"
                            End If
                        Else
                            '零差价控制：定价药品，检查录入的成本价是否等于售价
                            If gtype_UserSysParms.P275_零差价管理模式 = 2 And Val(strKey) <> Val(.TextMatrix(.Row, mconIntCol售价)) Then
                                If IsPriceAdjustMod(Val(.TextMatrix(.Row, 0))) = True Then
                                    MsgBox "该定价药品已启用零差价管理模式，入库成本价应和售价(" & zlStr.FormatEx(.TextMatrix(.Row, mconIntCol售价), mintPriceDigit, , True) & ")相等！", vbInformation + vbOKOnly, gstrSysName
                                    strKey = .TextMatrix(.Row, mconIntCol售价)
                                    .TextMatrix(.Row, mconIntCol成本价) = zlStr.FormatEx(strKey, mintPriceDigit, , True)
                                    .Text = strKey
                                    .TextMatrix(.Row, mconIntcol加成率) = "0%"
                                    .TextMatrix(.Row, mconIntCol扣率) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol成本价)) / Val(.TextMatrix(.Row, mconIntCol采购价)) * 100, 7, , True)
                                End If
                            End If
                        End If
                    End If
                 End If
                
                '设置金额
                If strKey <> "" And .TextMatrix(.Row, mconIntCol数量) <> "" Then
                    .TextMatrix(.Row, mconIntCol成本金额) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol数量) * Val(.TextMatrix(.Row, mconIntCol成本价)), mintMoneyDigit, , True)
                    .TextMatrix(.Row, mconIntCol售价金额) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol数量) * Val(.TextMatrix(.Row, mconIntCol售价)), mintMoneyDigit, , True)
                    .TextMatrix(.Row, mconintcol发票金额) = IIf(Trim(.TextMatrix(.Row, mconintcol发票号)) = "", "", .TextMatrix(.Row, mconIntCol成本金额))
                    .TextMatrix(.Row, mconintCol差价) = zlStr.FormatEx(IIf(.TextMatrix(.Row, mconIntCol售价金额) = "", 0, .TextMatrix(.Row, mconIntCol售价金额)) - IIf(.TextMatrix(.Row, mconIntCol成本金额) = "", 0, .TextMatrix(.Row, mconIntCol成本金额)), mintMoneyDigit, , True)
                End If
                
                Call 检查成本价
                Call 显示合计金额
                Call Set时价分批药品零售价(.Row, Val(.TextMatrix(.Row, mconIntCol售价)) / Val(.TextMatrix(.Row, mconIntCol比例系数)))
            Case mconIntCol售价
                '输入的售价不能大于指导零售价
                If Val(.TextMatrix(.Row, 0)) = 0 Then Exit Sub
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "售价必须为数字型，请重输！", vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If .TxtVisible = False Then strKey = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol售价), mintPriceDigit, , True)
                
                '判断输入的零售价与指导零售价
                gstrSQL = "Select 指导零售价 From 药品目录 Where 药品ID=[1] "
                Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[读取指导零售价]", Val(.TextMatrix(.Row, 0)))
                
                dbl指导零售价 = Round(rsTemp!指导零售价 * Val(.TextMatrix(.Row, mconIntCol比例系数)), mintPriceDigit)
                strKey = Round(strKey, 5)
                If Val(strKey) > dbl指导零售价 Then
                    MsgBox "售价不能大于指导零售价（指导零售价：￥" & dbl指导零售价 & "）", vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                '零差价控制：时价药品，售价直接等于成本价；只有时价药品才能修改售价
                If gtype_UserSysParms.P275_零差价管理模式 = 2 And Split(.TextMatrix(.Row, mconIntCol原销期), "||")(2) = 1 And Val(strKey) <> Val(.TextMatrix(.Row, mconIntCol成本价)) Then
                    If IsPriceAdjustMod(Val(.TextMatrix(.Row, 0))) = True Then
                        MsgBox "该时价药品已启用零差价管理模式，售价应和成本价(" & zlStr.FormatEx(.TextMatrix(.Row, mconIntCol成本价), mintPriceDigit, , True) & ")相等！", vbInformation + vbOKOnly, gstrSysName
                        strKey = .TextMatrix(.Row, mconIntCol成本价)
'                        Cancel = True
'                        .TxtSetFocus
'                        Exit Sub
                    End If
                End If
                
                .Text = zlStr.FormatEx(strKey, mintPriceDigit, , True)
                .TextMatrix(.Row, .Col) = .Text
                
'                If Len(Mid(.Text, InStr(1, .Text, ".") + 1)) > Get精度(2, mintUnit) Then
'                    MsgBox "售价精度大于了设置的计算精度，请重输！", vbInformation, gstrSysName
'                    Cancel = True
'                    .TxtSetFocus
'                    Exit Sub
'                End If
                                                
                dbl成本价 = IIf(mint时价入库售价加成方式 = 0, Val(.TextMatrix(.Row, mconIntCol成本价)), Val(.TextMatrix(.Row, mconIntCol采购价)))
                .TextMatrix(.Row, mconIntcol加成率) = zlStr.FormatEx(计算加成率(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, .Col)), dbl成本价), 2) & "%"
                '重算差价
                .TextMatrix(.Row, mconIntCol售价金额) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol售价)) * Val(.TextMatrix(.Row, mconIntCol数量)), mintMoneyDigit, , True)
                .TextMatrix(.Row, mconintCol差价) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol售价金额)) - Val(.TextMatrix(.Row, mconIntCol成本金额)), mintMoneyDigit, , True)
                
                Call Set时价分批药品零售价(.Row, Val(.TextMatrix(.Row, mconIntCol售价)) / Val(.TextMatrix(.Row, mconIntCol比例系数)))
            Case mconintCol零售价
                '用于时价分批药品按零售价入库
                '输入的零售价不能大于指导零售价
                If Val(.TextMatrix(.Row, 0)) = 0 Then Exit Sub
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "零售价必须为数字型，请重输！", vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If .TxtVisible = False Then strKey = zlStr.FormatEx(.TextMatrix(.Row, mconintCol零售价), gtype_UserDrugDigits.Digit_零售价, , True)
                
                '判断输入的零售价与指导零售价
                gstrSQL = "Select 指导零售价 From 药品目录 Where 药品ID=[1] "
                Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[读取指导零售价]", Val(.TextMatrix(.Row, 0)))
                
                dbl指导零售价 = Round(rsTemp!指导零售价, gtype_UserDrugDigits.Digit_零售价)
                
                If Val(strKey) <> 0 Then
                    strKey = Round(strKey, gtype_UserDrugDigits.Digit_零售价)
                End If
                If Val(strKey) > dbl指导零售价 Then
                    MsgBox "零售价不能大于指导零售价（指导零售价：￥" & dbl指导零售价 & "）", vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                                
                .Text = zlStr.FormatEx(strKey, gtype_UserDrugDigits.Digit_零售价, , True)
                .TextMatrix(.Row, .Col) = .Text
                
                .TextMatrix(.Row, mconIntCol售价) = zlStr.FormatEx(Val(.TextMatrix(.Row, .Col)) * Val(.TextMatrix(.Row, mconIntCol比例系数)), mintPriceDigit, , True)
                .TextMatrix(.Row, mconIntCol售价金额) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol售价)) * Val(.TextMatrix(.Row, mconIntCol数量)), mintMoneyDigit, , True)
                
                If gtype_UserSysParms.P275_零差价管理模式 = 2 And IsPriceAdjustMod(Val(.TextMatrix(.Row, 0))) = True Then
                    .TextMatrix(.Row, mconIntCol成本价) = .TextMatrix(.Row, mconIntCol售价)
                    .TextMatrix(.Row, mconIntCol成本金额) = .TextMatrix(.Row, mconIntCol售价金额)
                End If
                
                .TextMatrix(.Row, mconintCol差价) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol售价金额)) - Val(.TextMatrix(.Row, mconIntCol成本金额)), mintMoneyDigit, , True)
                
                dbl成本价 = IIf(mint时价入库售价加成方式 = 0, Val(.TextMatrix(.Row, mconIntCol成本价)), Val(.TextMatrix(.Row, mconIntCol采购价)))
                .TextMatrix(.Row, mconIntcol加成率) = zlStr.FormatEx(计算加成率(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol售价)), dbl成本价), 2) & "%"
                
                Call Set时价分批药品零售价(.Row, Val(.Text))
                Call 显示合计金额
            Case mconIntCol成本金额
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "对不起，成本金额必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Val(strKey) >= 10 ^ 14 - 1 Then
                        MsgBox "成本金额必须小于" & (10 ^ 14 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strKey) * Val(.TextMatrix(.Row, mconIntCol数量)) < 0 Then
                        MsgBox "成本金额符号应与数量符号一致！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                End If
                
                '格式化金额
                If strKey <> "" Then
                    strKey = zlStr.FormatEx(strKey, mintMoneyDigit, , True)
                    .Text = strKey
                End If
                
                If strKey <> "" And strKey <> .TextMatrix(.Row, mconIntCol成本金额) Then
                    If .TextMatrix(.Row, mconIntCol数量) <> "" Then
                        '零差价控制：定价药品，不能调整成本金额（因为售价固定，售价金额也固定）
                        If gtype_UserSysParms.P275_零差价管理模式 = 2 And IsPriceAdjustMod(Val(.TextMatrix(.Row, 0))) = True Then
                            If Split(.TextMatrix(.Row, mconIntCol原销期), "||")(2) = 0 And strKey <> .TextMatrix(.Row, mconIntCol售价金额) Then
                                MsgBox "该定价药品已启用零差价管理模式，不能调整成本金额！", vbInformation + vbOKOnly, gstrSysName
                                strKey = .TextMatrix(.Row, mconIntCol售价金额)
                                .Text = strKey
                                Cancel = True
'                                .TxtSetFocus
                                Exit Sub
                            End If
                        Else
                            If mbln加价率 Then
                                '取得改变采购金额前的加价率
                                dbl成本价 = IIf(mint时价入库售价加成方式 = 0, Val(.TextMatrix(.Row, mconIntCol成本价)), Val(.TextMatrix(.Row, mconIntCol采购价)))
                                mdbl加价率 = 15
                                If Val(.TextMatrix(.Row, mconIntCol售价)) <> 0 And dbl成本价 <> 0 Then
                                    mdbl加价率 = 计算加成率(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol售价)), dbl成本价)
                                End If
                            End If
                            
                            '计算成本价、采购价：成本价=采购金额/数量;采购价=(采购金额/数量)/扣率
                            If Val(.TextMatrix(.Row, mconIntCol扣率)) = 0 Then
                                .TextMatrix(.Row, mconIntCol扣率) = "100"
                            End If
                            .TextMatrix(.Row, mconIntCol成本价) = zlStr.FormatEx(strKey / .TextMatrix(.Row, mconIntCol数量), mintCostDigit, , True)
                            .TextMatrix(.Row, mconIntCol采购价) = zlStr.FormatEx((strKey / .TextMatrix(.Row, mconIntCol数量)) * 100 / Val(.TextMatrix(.Row, mconIntCol扣率)), mintCostDigit, , True)
                            
                            '根据参数决定时价药品售价公式中成本价的算法
                            dbl成本价 = IIf(mint时价入库售价加成方式 = 0, Val(.TextMatrix(.Row, mconIntCol成本价)), Val(.TextMatrix(.Row, mconIntCol采购价)))
                            
                            '对时价药品的处理
                            If .TextMatrix(.Row, mconIntCol原销期) <> "" Then
                                '重新计算零售价、差价
                                If Split(.TextMatrix(.Row, mconIntCol原销期), "||")(2) = 1 Then
                                    '由于存在差价让利比的存在,需要按加成率计算,因此将指导差价率转换成加成率计算 公式：加成率=1/(1-差价率)-1
                                    If mbln加价率 Then
                                        dbl加成率 = (mdbl加价率 / 100)
                                        dblTemp售价 = dbl成本价 * (1 + (mdbl加价率 / 100))
                                        
                                        If gtype_UserSysParms.P183_时价取上次售价 <> 1 Then  '没有勾选时价取上次售价参数
                                            .TextMatrix(.Row, mconIntCol售价) = zlStr.FormatEx(时价药品零售价(Val(.TextMatrix(.Row, 0)), dbl成本价, dbl加成率, dblTemp售价), mintPriceDigit, , True)
                                        End If
                                        .TextMatrix(.Row, mconIntCol售价金额) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol售价)) * Val(.TextMatrix(.Row, mconIntCol数量)), mintMoneyDigit, , True)
                                        .TextMatrix(.Row, mconintCol差价) = zlStr.FormatEx(IIf(.TextMatrix(.Row, mconIntCol售价金额) = "", 0, .TextMatrix(.Row, mconIntCol售价金额)) - IIf(.TextMatrix(.Row, mconIntCol成本金额) = "", 0, .TextMatrix(.Row, mconIntCol成本金额)), mintMoneyDigit, , True)
                                    Else
                                        If mint时价分段加成方式 = 1 Then
                                            If get分段加成售价(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol比例系数)), dbl成本价, dbl加成率, dblTemp售价) = False Then
                                                Cancel = True
                                                .TxtSetFocus
                                                Exit Sub
                                            End If
                                        Else
                                            dbl加成率 = Val(Replace(.TextMatrix(.Row, mconIntcol加成率), "%", "")) / 100
                                            dblTemp售价 = dbl成本价 * (1 + dbl加成率)
                                        End If
                                                                            
                                        If gtype_UserSysParms.P183_时价取上次售价 <> 1 Then  '没有勾选时价取上次售价参数
                                            .TextMatrix(.Row, mconIntCol售价) = zlStr.FormatEx(时价药品零售价(Val(.TextMatrix(.Row, 0)), dbl成本价, dbl加成率, dblTemp售价), mintPriceDigit, , True)
                                        End If
                                        .TextMatrix(.Row, mconIntCol售价金额) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol数量) * .TextMatrix(.Row, mconIntCol售价), mintMoneyDigit, , True)
                                        .TextMatrix(.Row, mconIntcol加成率) = zlStr.FormatEx(dbl加成率 * 100, 2) & "%"
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                    If Val(.TextMatrix(.Row, mconIntCol数量)) <> 0 Then
                        .TextMatrix(.Row, mconIntCol成本价) = zlStr.FormatEx(strKey / Val(.TextMatrix(.Row, mconIntCol数量)), mintCostDigit, , True)
                        .TextMatrix(.Row, mconIntCol采购价) = zlStr.FormatEx(strKey / Val(.TextMatrix(.Row, mconIntCol数量)) * 100 / Val(.TextMatrix(.Row, mconIntCol扣率)), mintCostDigit, , True)
                    End If
                    .TextMatrix(.Row, mconintcol发票金额) = IIf(Trim(.TextMatrix(.Row, mconintcol发票号)) = "", "", zlStr.FormatEx(strKey, mintMoneyDigit, , True))
                    
                    '零差价控制：定价药品，不能调整成本金额（因为售价固定，售价金额也固定）
                    If gtype_UserSysParms.P275_零差价管理模式 = 2 And IsPriceAdjustMod(Val(.TextMatrix(.Row, 0))) = True Then
                        .TextMatrix(.Row, mconIntCol售价) = zlStr.FormatEx(strKey / Val(.TextMatrix(.Row, mconIntCol数量)), mintCostDigit, , True)
                        .TextMatrix(.Row, mconIntCol售价金额) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol数量) * .TextMatrix(.Row, mconIntCol售价), mintMoneyDigit, , True)
                    End If
                    
                    .TextMatrix(.Row, mconintCol差价) = zlStr.FormatEx(IIf(.TextMatrix(.Row, mconIntCol售价金额) = "", 0, .TextMatrix(.Row, mconIntCol售价金额)) - strKey, mintMoneyDigit, , True)
                    .TextMatrix(.Row, mconIntCol成本金额) = zlStr.FormatEx(strKey, mintMoneyDigit, , True)
                End If
                
                Call 检查成本价
                Call 显示合计金额
                Call Set时价分批药品零售价(.Row, Val(.TextMatrix(.Row, mconIntCol售价)) / IIf(Val(.TextMatrix(.Row, mconIntCol比例系数)) = 0, 1, Val(.TextMatrix(.Row, mconIntCol比例系数))))
            Case mconIntCol数量
                If .TextMatrix(.Row, 0) = "" Then
                    .Text = ""
                    Exit Sub
                End If
                
                If .TextMatrix(.Row, .Col) = "" And strKey = "" Then
                    MsgBox "对不起，数量必须输入！", vbOKOnly + vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "对不起，数量必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Val(strKey) = 0 Then
                        MsgBox "对不起，数量不能为零,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    If mint编辑状态 = 2 And Val(.TextMatrix(.Row, mconIntCol数量)) <> 0 And .TextMatrix(.Row, mconIntCol是否新行) = "否" Then
                        If Not 相同符号(Val(strKey), Val(.TextMatrix(.Row, mconIntCol数量))) Then
                            MsgBox "对不起，数量的符号应该与原单据数量的符号一致！", vbInformation, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                    If Val(strKey) < 0 Then
                        If mint编辑状态 = 8 Or mbln退货 Then
                            MsgBox "退库单不能输入负数，请重输！", vbInformation, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        If Not zlStr.IsHavePrivs(mstrPrivs, "负数开单") Then
                            MsgBox "对不起，你没有负数开单的权限，请重输！", vbInformation, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        If .TextMatrix(.Row, mconIntCol分批属性) = 1 Then
                            MsgBox "分批药品不允许负数入库，请重输", vbInformation, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                    
                    If Val(strKey) >= 10 ^ 11 - 1 Then
                        MsgBox "数量必须小于" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    strKey = zlStr.FormatEx(strKey, mintNumberDigit, , True)
                    '1 检查是否有足够的库存可以退货;2 检查负数退库时库存是否足够
                    If mint编辑状态 = 8 Or mbln退货 Or Val(strKey) < 0 Then
                        If Not CheckUsableNum(cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol批次)), Val(.Text), Val(.TextMatrix(.Row, mconIntCol比例系数)), Trim(txtNO.Text), 1, mint库存检查, mintNumberDigit) Then
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                    
                    .Text = strKey
                    
'                    If Len(Mid(strkey, InStr(1, strkey, ".") + 1)) > Get精度(3, mintUnit) Then
'                        MsgBox "数量精度大于了设置的计算精度，请重输！", vbInformation, gstrSysName
'                        Cancel = True
'                        .TxtSetFocus
'                        Exit Sub
'                    End If
                    If .TextMatrix(.Row, mconIntCol成本价) <> "" Then
                        .TextMatrix(.Row, mconIntCol成本金额) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol成本价) * strKey, mintMoneyDigit, , True)
                        
                        '零差价控制
                        If gtype_UserSysParms.P275_零差价管理模式 = 2 And IsPriceAdjustMod(Val(.TextMatrix(.Row, 0))) = True Then
                            '如果启用了零差价管理不用再重算售价
                        Else
                            '根据参数决定时价药品售价公式中成本价的算法
                            dbl成本价 = IIf(mint时价入库售价加成方式 = 0, Val(.TextMatrix(.Row, mconIntCol成本价)), Val(.TextMatrix(.Row, mconIntCol采购价)))
                            
                            '时价药品的处理
                            If .TextMatrix(.Row, mconIntCol原销期) <> "" And mint编辑状态 <> 8 And mbln退货 <> True Then
                                If Split(.TextMatrix(.Row, mconIntCol原销期), "||")(2) = 1 Then
                                    '由于存在差价让利比的存在,需要按加成率计算,因此将指导差价率转换成加成率计算 公式：加成率=1/(1-差价率)-1
                                    If mbln加价率 Then
                                        mdbl加价率 = 15
                                        If Val(.TextMatrix(.Row, mconIntCol售价)) <> 0 And dbl成本价 <> 0 Then
                                            mdbl加价率 = 计算加成率(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol售价)), dbl成本价)
                                        End If
                                        
    '                                    mdbl加价率 = mdbl加价率 / 100
                                        dblTemp售价 = dbl成本价 * (1 + (mdbl加价率 / 100))
                                        
                                        If gtype_UserSysParms.P183_时价取上次售价 <> 1 Then  '没有勾选时价取上次售价参数
                                            .TextMatrix(.Row, mconIntCol售价) = zlStr.FormatEx(时价药品零售价(Val(.TextMatrix(.Row, 0)), dbl成本价, mdbl加价率, dblTemp售价), mintPriceDigit, , True)
                                        End If
                                        
                                        .TextMatrix(.Row, mconIntCol售价金额) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol售价)) * strKey, mintMoneyDigit, , True)
                                        .TextMatrix(.Row, mconintCol差价) = zlStr.FormatEx(IIf(.TextMatrix(.Row, mconIntCol售价金额) = "", 0, .TextMatrix(.Row, mconIntCol售价金额)) - IIf(.TextMatrix(.Row, mconIntCol成本金额) = "", 0, .TextMatrix(.Row, mconIntCol成本金额)), mintMoneyDigit, , True)
                                    Else
                                        If mint时价分段加成方式 = 1 Then
                                            If get分段加成售价(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol比例系数)), dbl成本价, dbl加成率, dblTemp售价) = False Then
                                                Cancel = True
                                                .TxtSetFocus
                                                Exit Sub
                                            End If
                                        Else
                                            dbl加成率 = Val(Replace(.TextMatrix(.Row, mconIntcol加成率), "%", "")) / 100
                                            dblTemp售价 = dbl成本价 * (1 + dbl加成率)
                                        End If
                                                                            
                                        If gtype_UserSysParms.P183_时价取上次售价 <> 1 Then  '没有勾选时价取上次售价参数
                                            .TextMatrix(.Row, mconIntCol售价) = zlStr.FormatEx(时价药品零售价(Val(.TextMatrix(.Row, 0)), dbl成本价, dbl加成率, dblTemp售价), mintPriceDigit, , True)
                                        End If
                                        .TextMatrix(.Row, mconIntcol加成率) = zlStr.FormatEx(dbl加成率 * 100, 2) & "%"
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                    If .TextMatrix(.Row, mconIntCol售价) <> "" Then
                        .TextMatrix(.Row, mconIntCol售价金额) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol售价) * strKey, mintMoneyDigit, , True)
                    End If
                    .TextMatrix(.Row, mconintCol差价) = zlStr.FormatEx(IIf(.TextMatrix(.Row, mconIntCol售价金额) = "", 0, .TextMatrix(.Row, mconIntCol售价金额)) - IIf(.TextMatrix(.Row, mconIntCol成本金额) = "", 0, .TextMatrix(.Row, mconIntCol成本金额)), mintMoneyDigit, , True)
                    .TextMatrix(.Row, mconintcol发票金额) = .TextMatrix(.Row, mconIntCol成本金额)
                                    
                    .TextMatrix(.Row, mconIntCol数量) = strKey
                    Call Set时价分批药品零售价(.Row, Val(.TextMatrix(.Row, mconIntCol售价)) / Val(.TextMatrix(.Row, mconIntCol比例系数)))
                End If
                显示合计金额
            Case mconIntCol冲销数量
                If .TextMatrix(.Row, .Col) = "" And strKey = "" Then
                    MsgBox "对不起，数量必须输入！", vbOKOnly + vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "对不起，数量必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Val(strKey) < 0 Then
                        If Not zlStr.IsHavePrivs(mstrPrivs, "负数开单") Then
                            MsgBox "对不起，你没有负数开单的权限，请重输！", vbInformation, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                    
                    If Val(strKey) <> 0 And Not 相同符号(Val(strKey), Val(.TextMatrix(.Row, mconIntCol数量))) Then
                        MsgBox "对不起，冲销数量的符号应该与原有数量一致！", vbInformation, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strKey) >= 0 Then
                        If Val(strKey) > Val(.TextMatrix(.Row, mconIntCol数量)) Then
                            MsgBox "对不起，冲销数量不能大于原有数量,请重输！", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    Else
                        If Val(strKey) < Val(.TextMatrix(.Row, mconIntCol数量)) Then
                            MsgBox "对不起，冲销数量不能大于原有数量,请重输！", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                                        
                    If Val(strKey) >= 10 ^ 11 - 1 Then
                        MsgBox "冲销数量必须小于" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    strKey = zlStr.FormatEx(strKey, mintNumberDigit, , True)
                    .Text = strKey
                    
                    If .TextMatrix(.Row, mconIntCol成本价) <> "" Then
                        .TextMatrix(.Row, mconIntCol成本金额) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol成本价) * strKey, mintMoneyDigit, , True)
                    End If
                    If .TextMatrix(.Row, mconIntCol售价) <> "" Then
                        .TextMatrix(.Row, mconIntCol售价金额) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol售价) * strKey, mintMoneyDigit, , True)
                    End If
                    .TextMatrix(.Row, mconintCol差价) = zlStr.FormatEx(IIf(.TextMatrix(.Row, mconIntCol售价金额) = "", 0, .TextMatrix(.Row, mconIntCol售价金额)) - IIf(.TextMatrix(.Row, mconIntCol成本金额) = "", 0, .TextMatrix(.Row, mconIntCol成本金额)), mintMoneyDigit, , True)
                    
                    gstrSQL = "select sum(nvl(发票金额,0)) as 发票金额 " _
                        & " From 药品收发记录 x,(Select 收发id,发票金额 From 应付记录 Where 系统标识=1 And 记录性质=0)  y " _
                        & " WHERE x.id=y.收发id(+) and x.NO=[1] AND 单据=1 " _
                        & " and x.药品id=[2] " _
                        & " and x.序号=[3] "
                    Set rsDrug = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, mstr单据号, Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol序号)))
                    
                    If rsDrug.EOF Then
                        .TextMatrix(.Row, mconintcol发票金额) = zlStr.FormatEx(0, mintMoneyDigit, , True)
                    Else
                        .TextMatrix(.Row, mconintcol发票金额) = zlStr.FormatEx(strKey / .TextMatrix(.Row, mconIntCol数量) * rsDrug.Fields(0), mintMoneyDigit, , True)
                    End If
                    
                    .TextMatrix(.Row, mconIntCol冲销数量) = strKey
                    Call Set时价分批药品零售价(.Row, Val(.TextMatrix(.Row, mconintCol零售价)))
                End If
                显示合计金额
            Case mconintcol发票号
                If Trim(.TextMatrix(.Row, mconIntCol药名)) = "" Then
                    .Text = ""
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If Trim(.Text) = "" Then
                    If .TxtVisible = True Then
                        .ColData(mconIntCol发票日期) = 5
                        .ColData(mconintcol发票金额) = 5
                        .ColData(mconintcol发票代码) = 5
                        .TextMatrix(.Row, mconintcol发票代码) = ""
                        .TextMatrix(.Row, mconintcol发票金额) = ""
                        .TextMatrix(.Row, mconIntCol发票日期) = ""
                        .TextMatrix(.Row, .Col) = " "
                        .Text = " "
                    ElseIf .TxtVisible = False Then
                        If Trim(.TextMatrix(.Row, mconintcol发票号)) = "" Then
                           .ColData(mconIntCol发票日期) = 5
                           .ColData(mconintcol发票金额) = 5
                           .ColData(mconintcol发票代码) = 5
                           .TextMatrix(.Row, mconintcol发票代码) = ""
                           .TextMatrix(.Row, mconintcol发票金额) = ""
                           .TextMatrix(.Row, mconIntCol发票日期) = ""
                           .TextMatrix(.Row, .Col) = " "
                           .Text = " "
                        Else
                            .Text = .TextMatrix(.Row, .Col)
                           
                            If mint编辑状态 = 9 Or mint编辑状态 = 3 Or mint编辑状态 = 7 Then
                                .ColData(mconIntCol发票日期) = IIf(InStr(1, mstrControlItem, ",发票日期,") > 0, 2, 5)
                                .ColData(mconintcol发票代码) = IIf(InStr(1, mstrControlItem, ",发票代码,") > 0, 4, 5)
                                .ColData(mconintcol发票金额) = IIf(InStr(1, mstrControlItem, ",发票金额,") > 0, 4, 5)
                            Else
                                .ColData(mconIntCol发票日期) = 2
                                .ColData(mconintcol发票代码) = 4
                                .ColData(mconintcol发票金额) = 4
                           End If
                        End If
                    End If
                ElseIf mint记录状态 = 1 Then
                    If mint编辑状态 = 9 Or mint编辑状态 = 3 Or mint编辑状态 = 7 Then
                         .ColData(mconIntCol发票日期) = IIf(InStr(1, mstrControlItem, ",发票日期,") > 0, 2, 5)
                         .ColData(mconintcol发票代码) = IIf(InStr(1, mstrControlItem, ",发票代码,") > 0, 4, 5)
                         .ColData(mconintcol发票金额) = IIf(InStr(1, mstrControlItem, ",发票金额,") > 0, 4, 5)
                     Else
                         .ColData(mconIntCol发票日期) = 2
                         .ColData(mconintcol发票代码) = 4
                         .ColData(mconintcol发票金额) = 4
                    End If
                    .TextMatrix(.Row, mconintcol发票金额) = .TextMatrix(.Row, mconIntCol成本金额)
                End If
                    
                Exit Sub
            Case mconintcol发票代码
                If Trim(.Text) = "" Then
                   If mconintcol发票代码 <> mintLastCol Then
                       .Col = GetNextEnableCol(mconintcol发票代码)
                       .Text = ""
                       Cancel = True
                       Exit Sub
                   End If
                End If
            Case mconintcol发票金额
                If Trim(.TextMatrix(.Row, mconIntCol药名)) = "" Then
                    .Text = ""
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "对不起，发票金额必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Abs(Val(strKey)) < 0.001 Then
                        MsgBox "对不起，发票金额必须大于0.001,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strKey) >= 10 ^ 14 - 1 Then
                        MsgBox "发票金额必须小于" & (10 ^ 14 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                End If
                
                If strKey <> "" Then
                    strKey = zlStr.FormatEx(strKey, 2, , True)
                    .Text = strKey
                ElseIf .TxtVisible = True Then
                    .Text = " "
                ElseIf .TxtVisible = False Then
                    If .TextMatrix(.Row, .Col) = "" Then
                        .Text = " "
                    Else
                        .Text = .TextMatrix(.Row, .Col)
                    End If
                    
                End If
            Case mconIntCol发票日期
                If strKey <> "" Then
                    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
                        strKey = TranNumToDate(strKey)
                        
                        If strKey = "" Then
                            MsgBox "对不起，效期必须为日期型！", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            Exit Sub
                        End If
                        .Text = strKey
                        Exit Sub
                    End If
                    If Not IsDate(strKey) Then
                        MsgBox "对不起，发票日期必须为日期型如(2000-10-10) 或 （20001010）,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                End If
            Case mconIntCol批准文号
                If strKey = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mconIntCol批准文号) = ""
                    End If
                    If mconIntCol批准文号 <> mintLastCol Then
                        .Col = GetNextEnableCol(mconIntCol批准文号)
                        Cancel = True
                    End If
                    Exit Sub
                End If
            Case mconIntCol外观
                '无处理
                If .Text = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mconIntCol外观) = ""
                    End If
                    If mconIntCol外观 <> mintLastCol Then
                        .Col = GetNextEnableCol(mconIntCol外观)
                        Cancel = True
                        Exit Sub
                    End If
                Else
                    Dim rs外观 As New Recordset
                    
                    gstrSQL = "Select 编码,简码,名称 From 药品外观 " _
                            & "Where upper(名称) like [1] or Upper(编码) like [1] or Upper(简码) like [2] "
                    Set rs外观 = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, IIf(gstrMatchMethod = "0", "%", "") & strKey & "%", strKey & "%")
                    
                    If rs外观.EOF Then
                        .TextMatrix(.Row, mconIntCol外观) = .Text
                        If mconIntCol外观 <> mintLastCol And mconIntCol外观 < mconintcol发票号 Then
                            .Col = mconintcol发票号
                            Cancel = True
                            Exit Sub
                        End If
                    Else
                        If rs外观.RecordCount = 1 Then
                            .TextMatrix(.Row, mconIntCol外观) = rs外观.Fields("名称")
                            .Text = rs外观.Fields("名称")
                        Else
                            Set msh产地.Recordset = rs外观
                            With msh产地
                                .Redraw = False
                                .Left = Pic单据.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                                .Top = Pic单据.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight
                                .Visible = True
                                .SetFocus
                                .ColWidth(0) = 800
                                .ColWidth(1) = 800
                                .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(0)
                                .Row = 1
                                .Col = 0
                                .TopRow = 1
                                .ColSel = .Cols - 1
                                .Redraw = True
                                Cancel = True
                                Exit Sub
                            End With
                        End If
                    End If
                End If
            Case mconintcol随货单号
                '无处理
                If Trim(.Text) = "" Then
                    If .TxtVisible = True Then
                        .ColData(mconintcol随货日期) = 5
                        .TextMatrix(.Row, mconintcol随货日期) = ""
                        .TextMatrix(.Row, .Col) = ""
                        .Text = ""
                    ElseIf .TxtVisible = False Then
                        If Trim(.TextMatrix(.Row, mconintcol随货单号)) = "" Then
                           .ColData(mconintcol随货日期) = 5
                           .TextMatrix(.Row, mconintcol随货日期) = ""
                           .TextMatrix(.Row, .Col) = ""
                           .Text = ""
                        Else
                            .Text = .TextMatrix(.Row, .Col)
                            .ColData(mconintcol随货日期) = 2
                        End If
                    End If
                    
                Else
                    .TextMatrix(.Row, .Col) = .Text
                    .ColData(mconintcol随货日期) = 2
                End If
                
                If mconintcol随货单号 <> mintLastCol Then
                    .Col = GetNextEnableCol(mconintcol随货单号)
                    Cancel = True
                    Exit Sub
                End If
            Case mconintcol随货日期
                If .Text = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mconintcol随货日期) = ""
                    End If
                    If mconintcol随货日期 <> mintLastCol Then
                        .Col = GetNextEnableCol(mconintcol随货日期)
                        Cancel = True
                        Exit Sub
                    End If
                Else
                    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
                        strKey = TranNumToDate(strKey)
                        
                        If strKey = "" Then
                            MsgBox "对不起，随货日期必须为日期型！", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            Exit Sub
                        End If
                        .Text = strKey
                        Exit Sub
                    End If
                    If Not IsDate(strKey) Then
                        MsgBox "对不起，随货日期必须为日期型如(2000-10-10) 或 （20001010）,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                End If
            Case mconintcol产品合格证
                '无处理
                If .Text = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mconintcol产品合格证) = ""
                    End If
                    If mconintcol产品合格证 <> mintLastCol Then
                        .Col = GetNextEnableCol(mconintcol产品合格证)
                        Cancel = True
                        Exit Sub
                    End If
                End If
        End Select
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'从药品目录中取值并附给相应的列
Private Function SetColValue(ByVal intRow As Integer, ByVal lng药品ID As Long, ByVal str药品编码 As String, ByVal str通用名 As String, _
    ByVal str商品名 As String, ByVal str药品来源 As String, ByVal str基本药物 As String, _
    ByVal str规格 As String, ByVal str产地 As String, ByVal str单位 As String, ByVal num售价 As Double, _
    ByVal num指导批发价 As Double, ByVal str原生产商 As String, ByVal int原效期 As Integer, _
    ByVal str简码 As String, ByVal num比例系数 As Double, ByVal lng批次 As Long, _
    ByVal int是否变价 As Integer, ByVal int药房分批 As Integer, ByVal dbl加成率 As Double, ByVal str生产日期 As String, _
    ByVal str售价单位 As String, ByVal str原产地 As String) As Boolean
    
    Dim intCount As Integer
    Dim intCol As Integer
    Dim rsPrice As New Recordset
    Dim lngDepartid As Long
    Dim dblRate As Double, dbl成本价 As Double
    Dim bln招标药品 As Boolean, dbl差价让利比 As Double
    Dim rsTemp As New ADODB.Recordset
    Dim i As Long, j As Long
    Dim dbl时价成本价 As Double
    Dim str药名 As String
    Dim rsRecord As ADODB.Recordset
    Dim rsProvider As ADODB.Recordset
    Dim dblTemp售价 As Double
    Dim rs售价 As ADODB.Recordset
    
    SetColValue = False
    On Error GoTo errHandle
'    If mint编辑状态 = 8 Then
'        '检查是否重复
'        If Not CheckRepeatMedicine(mshBill, lng药品ID & "," & "0" & "|" & lng批次 & "," & mconIntCol批次, introw) Then
'            Exit Function
'        End If
'    End If
    
    With mshBill
        For intCol = 0 To .Cols - 1
            If intCol <> mconIntCol行号 Then .TextMatrix(intRow, intCol) = ""
        Next
        
        gstrSQL = "SELECT Nvl(a.差价让利比,0) 差价让利比,nvl(a.扣率,0) 扣率,Nvl(a.招标药品,0) 招标药品,nvl(a.成本价,0) 成本价, a.批准文号,a.上次批准文号,a.上次产地,b.产地,a.原产地,a.上次生产日期,a.药价级别 " & _
                  "from 药品规格 a,收费项目目录 b  where a.药品id=b.id and 药品id=[1]"
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[读取扣率]", lng药品ID)
            
        If rsTemp!扣率 = 0 Then
            dblRate = 100
        Else
            dblRate = rsTemp!扣率
        End If
        bln招标药品 = (rsTemp!招标药品 = 1)
        dbl差价让利比 = rsTemp!差价让利比
        dbl成本价 = rsTemp!成本价
        
        .TextMatrix(intRow, 0) = lng药品ID
        
        If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
            str药名 = str通用名
        Else
            str药名 = IIf(str商品名 <> "", str商品名, str通用名)
        End If
        
        .TextMatrix(intRow, mconIntCol药品编码和名称) = str药品编码 & str药名
        .TextMatrix(intRow, mconIntCol药品编码) = str药品编码
        .TextMatrix(intRow, mconIntCol药品名称) = str药名
        
        If mintDrugNameShow = 1 Then
            .TextMatrix(intRow, mconIntCol药名) = .TextMatrix(intRow, mconIntCol药品编码)
        ElseIf mintDrugNameShow = 2 Then
            .TextMatrix(intRow, mconIntCol药名) = .TextMatrix(intRow, mconIntCol药品名称)
        Else
            .TextMatrix(intRow, mconIntCol药名) = .TextMatrix(intRow, mconIntCol药品编码和名称)
        End If
        
        .TextMatrix(intRow, mconIntCol商品名) = str商品名
        
        .TextMatrix(intRow, mconIntCol来源) = str药品来源
        .TextMatrix(intRow, mconIntCol基本药物) = str基本药物
        .TextMatrix(intRow, mconIntCol药价级别) = IIf(IsNull(rsTemp!药价级别), "", rsTemp!药价级别)
        .TextMatrix(intRow, mconIntCol规格) = str规格
        .TextMatrix(intRow, mconIntCol产地) = IIf(IsNull(str产地), "", str产地)
        .TextMatrix(intRow, mconIntCol原产地) = IIf(IsNull(str原产地), nvl(rsTemp!原产地), str原产地)
        
        '产地、批准文号、生产日期规则，根据参数设置取
        '参数：优先从上次入库取
        '产地：直接从规格表中取上次产地，如果没有则从收费项目中取产地，没有则不填产地
        '批准文号：优先从规格表中取上次批准文号，如果没有则从规格表中取批准文号，还没有则不填批准文号
        '生产日期：优先从规格表中取上次生产日期，如果没有则不填
        '成本价：从规格表中取成本价
        
        '参数：优先从最近库存批次取
        '产地：优先从库存表最近批次中取产地，如果没有则从收费项目中取产地，没有则不填产地
        '批准文号：优先从库存表最近批次中取批准文号，如果没有则从规格表中取批准文号，还没有则不填批准文号
        '生产日期：优先从库存表最近批次中取生产日期，如果没有则不填
        '成本价：优先从药品库存表最近批次中取上次采购价，没有则从规格表中取成本价
        If IIf(IsNull(rsTemp!上次产地), "", rsTemp!上次产地) <> "" Then
            .TextMatrix(intRow, mconIntCol产地) = rsTemp!上次产地
        Else
            .TextMatrix(intRow, mconIntCol产地) = IIf(IsNull(rsTemp!产地), "", rsTemp!产地)
        End If
        If IIf(IsNull(rsTemp!上次批准文号), "", rsTemp!上次批准文号) <> "" Then
            .TextMatrix(intRow, mconIntCol批准文号) = rsTemp!上次批准文号
        Else
            .TextMatrix(intRow, mconIntCol批准文号) = IIf(IsNull(rsTemp!批准文号), "", rsTemp!批准文号)
        End If
        
        If IIf(IsNull(rsTemp!上次生产日期), "", rsTemp!上次生产日期) <> "" Then
            .TextMatrix(intRow, mconIntCol生产日期) = Format(rsTemp!上次生产日期, "yyyy-mm-dd")
        Else
            .TextMatrix(intRow, mconIntCol生产日期) = ""
        End If
        
        .TextMatrix(intRow, mconIntCol单位) = str单位
        .TextMatrix(intRow, mconIntCol售价) = zlStr.FormatEx(num售价, mintPriceDigit, , True)
        .TextMatrix(intRow, mconIntCol指导批发价) = zlStr.FormatEx(num指导批发价, mintCostDigit, , True)
        .TextMatrix(intRow, mconIntCol原生产商) = IIf(IsNull(str原生产商), "", str原生产商)
        .TextMatrix(intRow, mconIntCol批次) = lng批次
        
        '取出该药品的批号及效期，以及原采购价
        If mint编辑状态 = 8 Or mbln退货 Then
            gstrSQL = " Select 上次批号 批号,效期,上次生产日期,上次产地,原产地,批准文号,上次采购价 From 药品库存" & _
                    " Where 库房ID=[1] And 药品ID=[2] " & _
                    " And 性质=1 And nvl(批次,0)=[3] "
            Set rsPrice = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[提取该药品的批号和效期]", cboStock.ItemData(cboStock.ListIndex), lng药品ID, lng批次)
            If rsPrice.RecordCount <> 0 Then
                .TextMatrix(intRow, mconIntCol批号) = IIf(IsNull(rsPrice!批号), "", rsPrice!批号)
                .TextMatrix(intRow, mconIntCol效期) = IIf(IsNull(rsPrice!效期), "", rsPrice!效期)
                .TextMatrix(intRow, mconIntCol产地) = IIf(IsNull(rsPrice!上次产地), "", rsPrice!上次产地)
                .TextMatrix(intRow, mconIntCol原产地) = IIf(IsNull(rsPrice!原产地), "", rsPrice!原产地)
                .TextMatrix(intRow, mconIntCol批准文号) = IIf(IsNull(rsPrice!批准文号), "", rsPrice!批准文号)
                .TextMatrix(intRow, mconIntCol生产日期) = Format(rsPrice!上次生产日期, "yyyy-mm-dd")
                
                If gtype_UserSysParms.P149_效期显示方式 = 1 And .TextMatrix(intRow, mconIntCol效期) <> "" Then
                    '换算为有效期
                    .TextMatrix(intRow, mconIntCol效期) = Format(DateAdd("D", -1, .TextMatrix(intRow, mconIntCol效期)), "yyyy-mm-dd")
                End If
                
                dbl成本价 = nvl(rsPrice!上次采购价, 0)
                If dbl成本价 > 0 Then
                    .TextMatrix(intRow, mconIntCol采购价) = zlStr.FormatEx(dbl成本价 * num比例系数, mintCostDigit, , True)
                    .TextMatrix(intRow, mconIntCol成本价) = zlStr.FormatEx(dbl成本价 * num比例系数 * dblRate / 100, mintCostDigit, , True)
                End If
            End If
        End If
        
        '原效期字段下面保存原效期，指导差价，是否变价，药房分批等，格式为：原效期||指导差价率||是否变价||药房分批
        .TextMatrix(intRow, mconIntCol原销期) = IIf(IsNull(int原效期), "0", int原效期) & "||" & dbl加成率 & "||" & int是否变价 & "||" & int药房分批
       
        .TextMatrix(intRow, mconintcol简码) = str简码
        .TextMatrix(intRow, mconIntCol比例系数) = num比例系数
        If intRow > 1 Then
            .TextMatrix(intRow, mconintcol随货单号) = .TextMatrix(intRow - 1, mconintcol随货单号)
            .TextMatrix(intRow, mconintcol随货日期) = .TextMatrix(intRow - 1, mconintcol随货日期)
            .TextMatrix(intRow, mconintcol发票号) = .TextMatrix(intRow - 1, mconintcol发票号)
            .TextMatrix(intRow, mconintcol发票代码) = .TextMatrix(intRow - 1, mconintcol发票代码)
            .TextMatrix(intRow, mconIntCol发票日期) = .TextMatrix(intRow - 1, mconIntCol发票日期)
        End If
        
        SetInputFormat intRow
        SetDisCount intRow, dblRate
        lngDepartid = cboStock.ItemData(cboStock.ListIndex)
        
        '分批属性
        Call Get药品分批属性(intRow)
        
        '说明：这里区分分批核算和不分批核算的目的是提高运行速度。
        '本来可以不分这些，直接用第一条SQL语句实现，但不分批的药品就多在数据库中扫描一次。
        
        '对定价采购，不用取上次的采购价和扣率
        If Not (mint编辑状态 = 8 Or mbln退货) Then
            If mint取上次采购价方式 = 0 Then
                If Val(.TextMatrix(intRow, mconIntCol分批属性)) = 1 Then
                    gstrSQL = "select 上次采购价,上次产地,批准文号,上次生产日期 from 药品库存 where 性质=1 and 库房id=[1] and 药品id=[2] " & _
                            " and nvl(批次,0) =(select max(nvl(批次,0)) from 药品库存 where 性质=1 and 库房id=[1] )"
                Else
                    gstrSQL = "select 上次采购价,上次产地,批准文号,上次生产日期 from 药品库存 where 性质=1 and 库房id=[1] and 药品id=[2]"
                End If
                Set rsPrice = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[取上次采购价]", lngDepartid, lng药品ID)
                
                If Not rsPrice.EOF Then
                    .TextMatrix(intRow, mconIntCol产地) = IIf(IsNull(rsPrice!上次产地), IIf(IsNull(rsTemp!产地), "", rsTemp!产地), rsPrice!上次产地)
                    'mint时价入库售价加成方式
                    If nvl(rsPrice.Fields(0), 0) = 0 Then
                        If dbl成本价 > 0 Then
                            .TextMatrix(intRow, mconIntCol采购价) = zlStr.FormatEx(dbl成本价 * num比例系数, mintCostDigit, , True)
                            .TextMatrix(intRow, mconIntCol成本价) = zlStr.FormatEx(dbl成本价 * num比例系数 * dblRate / 100, mintCostDigit, , True)
                        End If
                    Else
                        .TextMatrix(intRow, mconIntCol采购价) = zlStr.FormatEx(rsPrice.Fields(0) * num比例系数, mintCostDigit, , True)
                        .TextMatrix(intRow, mconIntCol成本价) = zlStr.FormatEx(rsPrice.Fields(0) * num比例系数 * dblRate / 100, mintCostDigit, , True)
                    End If
                    .TextMatrix(intRow, mconIntCol批准文号) = IIf(IsNull(rsPrice!批准文号), IIf(IsNull(rsTemp!批准文号), "", rsTemp!批准文号), rsPrice!批准文号)
                    .TextMatrix(intRow, mconIntCol生产日期) = IIf(IsNull(rsPrice!上次生产日期), "", Format(rsPrice!上次生产日期, "yyyy-mm-dd"))
                Else
                    .TextMatrix(intRow, mconIntCol生产日期) = ""
                    If dbl成本价 > 0 Then
                        .TextMatrix(intRow, mconIntCol采购价) = zlStr.FormatEx(dbl成本价 * num比例系数, mintCostDigit, , True)
                        .TextMatrix(intRow, mconIntCol成本价) = zlStr.FormatEx(dbl成本价 * num比例系数 * dblRate / 100, mintCostDigit, , True)
                    End If
                End If
                If Val(.TextMatrix(intRow, mconIntCol采购价)) <> 0 Then
                    .TextMatrix(intRow, mconIntCol扣率) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol成本价)) / Val(.TextMatrix(intRow, mconIntCol采购价)) * 100, 7, , True)
                End If
            Else
                If dbl成本价 > 0 Then
                    .TextMatrix(intRow, mconIntCol采购价) = zlStr.FormatEx(dbl成本价 * num比例系数, mintCostDigit, , True)
                    .TextMatrix(intRow, mconIntCol成本价) = zlStr.FormatEx(dbl成本价 * num比例系数 * dblRate / 100, mintCostDigit, , True)
                End If
            End If
        End If
        
        If mbln取目录中产地信息 = True Then '取药品目录中的产地批准文号
            If IIf(IsNull(rsTemp!产地), "", rsTemp!产地) <> "" Then
                .TextMatrix(intRow, mconIntCol产地) = rsTemp!产地
            End If
            If IIf(IsNull(rsTemp!批准文号), "", rsTemp!批准文号) <> "" Then
                .TextMatrix(intRow, mconIntCol批准文号) = rsTemp!批准文号
            End If
        End If
        
        '根据参数决定时价药品售价公式中成本价的算法
        dbl时价成本价 = IIf(mint时价入库售价加成方式 = 0, Val(.TextMatrix(intRow, mconIntCol成本价)), Val(.TextMatrix(intRow, mconIntCol采购价)))
        
        '时价药品处理
        If int是否变价 = 1 Then
            '零差价控制：时价药品，售价直接等于成本价
            If gtype_UserSysParms.P275_零差价管理模式 = 2 And IsPriceAdjustMod(Val(.TextMatrix(intRow, 0))) = True Then
                .TextMatrix(intRow, mconIntCol售价) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol成本价), mintPriceDigit, , True)
                If .TextMatrix(intRow, mconIntCol数量) <> "" Then
                    .TextMatrix(intRow, mconIntCol售价金额) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol数量) * .TextMatrix(.Row, mconIntCol售价), mintMoneyDigit, , True)
                End If
            Else
                If mint编辑状态 <> 8 And mbln退货 = False Then
                    dblTemp售价 = dbl时价成本价 * (1 + dbl加成率)
                    '由于存在差价让利比的存在,需要按加成率计算,因此将指导差价率转换成加成率计算 公式：加成率=1/(1-差价率)-1
                    If gtype_UserSysParms.P183_时价取上次售价 = 1 Then
                        gstrSQL = "select nvl(上次售价,0) 上次售价 from 药品规格 where 药品id=[1]"
                                         
                        Set rs售价 = zlDataBase.OpenSQLRecord(gstrSQL, "查询售价", lng药品ID)
                        If rs售价!上次售价 > 0 Then
                            .TextMatrix(intRow, mconIntCol售价) = zlStr.FormatEx(rs售价!上次售价 * mlng包装系数, mintPriceDigit, , True)
                        Else
                            .TextMatrix(intRow, mconIntCol售价) = zlStr.FormatEx(时价药品零售价(Val(.TextMatrix(intRow, 0)), dbl时价成本价, dbl加成率, dblTemp售价), mintPriceDigit, , True)
                        End If
                    Else
                        .TextMatrix(intRow, mconIntCol售价) = zlStr.FormatEx(时价药品零售价(Val(.TextMatrix(intRow, 0)), dbl时价成本价, dbl加成率, dblTemp售价), mintPriceDigit, , True)
                    End If
                    .TextMatrix(intRow, mconIntcol加成率) = zlStr.FormatEx(dbl加成率 * 100, 2) & "%"
                Else
                    '按出库的形式计算售价
                    gstrSQL = " Select Decode(Nvl(批次,0),0,nvl(实际金额,0)/Nvl(实际数量,0),Nvl(零售价,nvl(实际金额,0)/Nvl(实际数量,0))) 售价 From 药品库存" & _
                              " Where 库房ID=[1] And 药品ID=[2] And 性质=1 And NVL(批次,0)=[3]"
                    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[按出库的形式计算售价]", cboStock.ItemData(cboStock.ListIndex), lng药品ID, lng批次)
                    
                    If Not rsTemp.EOF Then
                        .TextMatrix(intRow, mconIntCol售价) = zlStr.FormatEx(rsTemp!售价 * num比例系数, mintPriceDigit, , True)
                    Else
                        gstrSQL = "select nvl(上次售价,0) 售价 from 药品规格 where 药品id=[1]"
                        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "查询售价", lng药品ID)
                        
                        If Not rsTemp.EOF Then
                            .TextMatrix(intRow, mconIntCol售价) = zlStr.FormatEx(rsTemp!售价 * num比例系数, mintPriceDigit, , True)
                        End If
                    End If
                    
                    If Val(.TextMatrix(intRow, mconIntCol售价)) <> 0 And dbl时价成本价 <> 0 Then
                        .TextMatrix(intRow, mconIntcol加成率) = zlStr.FormatEx((Val(.TextMatrix(intRow, mconIntCol售价)) / dbl时价成本价 - 1) * 100, 2) & "%"
                    End If
                End If
            End If
        Else
            '定价药品显示加成率，无实际意义，仅显示
            If Val(.TextMatrix(intRow, mconIntCol售价)) <> 0 And dbl时价成本价 <> 0 Then
                .TextMatrix(intRow, mconIntcol加成率) = zlStr.FormatEx((Val(.TextMatrix(intRow, mconIntCol售价)) / dbl时价成本价 - 1) * 100, 2) & "%"
            End If
                        
            '零差价控制：定价药品，成本价默认等于售价
            If gtype_UserSysParms.P275_零差价管理模式 = 2 And IsPriceAdjustMod(Val(.TextMatrix(intRow, 0))) = True Then
                .TextMatrix(intRow, mconIntCol成本价) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol售价), mintPriceDigit, , True)
                .TextMatrix(intRow, mconIntcol加成率) = "0%"
                .TextMatrix(intRow, mconIntCol采购价) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol成本价) / .TextMatrix(intRow, mconIntCol扣率) * 100, mintPriceDigit, , True)
            End If
        End If
        
        If mstr外观 = "" Then
            gstrSQL = "Select 名称  From 药品外观 where 缺省标志=1"
            Set rsPrice = zlDataBase.OpenSQLRecord(gstrSQL, "SetColValue")
            
            If Not rsPrice.EOF Then
                .TextMatrix(intRow, mconIntCol外观) = IIf(IsNull(rsPrice!名称), "", rsPrice!名称)
                mstr外观 = rsPrice!名称
            End If
        Else
            .TextMatrix(intRow, mconIntCol外观) = mstr外观
        End If
        
        If mstr验收结论 = "" Then
            gstrSQL = "Select 名称  From 入库验收结论 where 缺省标志=1"
            Set rsPrice = zlDataBase.OpenSQLRecord(gstrSQL, "SetColValue")
            
            If Not rsPrice.EOF Then
                .TextMatrix(intRow, mconIntCol验收结论) = IIf(IsNull(rsPrice!名称), "", rsPrice!名称)
                mstr验收结论 = rsPrice!名称
            End If
        Else
            .TextMatrix(intRow, mconIntCol验收结论) = mstr验收结论
        End If
        
        If .TextMatrix(intRow, mconIntCol原销期) <> "" Then
            If mintUnit <> mconint售价单位 And Val(Split(.TextMatrix(intRow, mconIntCol原销期), "||")(2)) = 1 And Val(.TextMatrix(intRow, mconIntCol分批属性)) = 1 Then
                .TextMatrix(intRow, mconintCol零售单位) = str售价单位
            End If
        End If
        
        '招标药品需要上色
        mblnEnter = False
        intCol = .Col
        For i = mconIntCol药名 To .Cols - 1
            j = .ColData(i)
            If .ColData(i) = 5 Then .ColData(i) = 0
            .Col = i
            If bln招标药品 Then
                mshBill.MsfObj.CellForeColor = IIf(dbl差价让利比 = 0, &H800000, &H800080)
            Else
                mshBill.MsfObj.CellForeColor = IIf(dbl差价让利比 = 0, &H0, &H40&)     ' &H40C0&
            End If
            .ColData(i) = j
        Next
        .Col = intCol
        
        If (.TextMatrix(intRow, mconIntCol产地) <> "" And .TextMatrix(intRow, mconIntCol批准文号) <> "") Then
        Else
            If .TextMatrix(intRow, mconIntCol产地) <> "" And .TextMatrix(intRow, mconIntCol批准文号) = "" Then  '产地不为空，批准文号为空时
                gstrSQL = "select 批准文号,厂家名称 from 药品生产商对照 where  药品id=[1] and 厂家名称=[2]"
                Set rsProvider = zlDataBase.OpenSQLRecord(gstrSQL, "mshbill_CommandClick", mshBill.TextMatrix(mshBill.Row, 0), mshBill.TextMatrix(mshBill.Row, mconIntCol产地))
                Do While Not rsProvider.EOF
                    .TextMatrix(intRow, mconIntCol批准文号) = IIf(IsNull(rsProvider!批准文号), "", rsProvider!批准文号)
                    Exit Do
                Loop
            ElseIf (.TextMatrix(intRow, mconIntCol产地) = "" And .TextMatrix(intRow, mconIntCol批准文号) <> "") Then '产地为空，批准文号不为空时
                gstrSQL = "select 批准文号,厂家名称 from 药品生产商对照 where  药品id=[1]"
                Set rsProvider = zlDataBase.OpenSQLRecord(gstrSQL, "mshbill_CommandClick", mshBill.TextMatrix(mshBill.Row, 0))
                Do While Not rsProvider.EOF
                    .TextMatrix(mshBill.Row, mconIntCol产地) = IIf(IsNull(rsProvider!厂家名称), "", rsProvider!厂家名称)
                    Exit Do
                Loop
            ElseIf .TextMatrix(intRow, mconIntCol产地) = "" And .TextMatrix(intRow, mconIntCol批准文号) = "" Then '产地为空，批准文号为空时
                gstrSQL = "select 批准文号,厂家名称 from 药品生产商对照 where  药品id=[1]"
                Set rsProvider = zlDataBase.OpenSQLRecord(gstrSQL, "mshbill_CommandClick", mshBill.TextMatrix(mshBill.Row, 0))
                Do While Not rsProvider.EOF
                    .TextMatrix(intRow, mconIntCol批准文号) = IIf(IsNull(rsProvider!批准文号), "", rsProvider!批准文号)
                    .TextMatrix(mshBill.Row, mconIntCol产地) = IIf(IsNull(rsProvider!厂家名称), "", rsProvider!厂家名称)
                    Exit Do
                Loop
            End If

        End If
        
        If mint编辑状态 = 8 Then
            If Val(.TextMatrix(intRow, mconIntCol采购价)) = 0 Then
                MsgBox "第" & intRow & "行药品成本价为空了，请注意确认！", vbInformation, gstrSysName
            End If
        End If
        mblnEnter = True
    End With
    SetColValue = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function get分段加成售价(ByVal lng药品ID As Long, ByVal lng比例系数 As Long, ByVal dbl采购价 As Double, ByRef dblR加成率 As Double, ByRef dbl售价 As Double) As Boolean
    '功能:在启用时价药品分段加成入库后，根据采购价计算出相应的售价
    '售价计算公式：购进价格在2000元/支、瓶或盒（含2000元）以下的药品，最高零售价格=实际购进价×（1+差价率）+差价额；
    '               购进价格在2000元/支、瓶或盒（不含2000元）以上的药品：最高零售价格 = 实际购进价 + 差价额（此段已经调整，不再适用）

    '参数：采购价
    Dim dbl加成率 As Double
    Dim dbl差价额 As Double
    Dim blnData As Boolean
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    dbl加成率 = 0
    dbl差价额 = 0
    
    gstrSQL = "select 类别 from  收费项目目录 a where a.id=[1]"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "取得药品材质分类", lng药品ID)
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
                    dbl加成率 = IIf(IsNull(!加成率), 0, !加成率) / 100
                    dblR加成率 = dbl加成率
                    dbl差价额 = IIf(IsNull(!差价额), 0, !差价额)
                    blnData = True
                    Exit Do
                End If
            End With
            mrs分段加成.MoveNext
        Loop
    End If
    
    If blnData = False Then
        If rsTemp!类别 = 7 Then
            MsgBox "【草药】未设置金额段为：" & dbl采购价 & " " & "的分段加成数据，请到药品目录管理中分段加成率设置！", vbInformation, gstrSysName
        Else
            MsgBox "【西药/成药】未设置金额段为：" & dbl采购价 & " " & "的分段加成数据，请到药品目录管理中分段加成率设置！", vbInformation, gstrSysName
        End If
        get分段加成售价 = False
    End If
    
    dbl售价 = dbl采购价 * (1 + dbl加成率) + dbl差价额
    
    Set rsTemp = Nothing
    gstrSQL = "Select 指导零售价 From 药品规格 Where 药品ID=[1] "
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[读取指导零售价]", lng药品ID)
    If rsTemp!指导零售价 * lng比例系数 < dbl售价 Then
        dbl售价 = rsTemp!指导零售价 * lng比例系数
    End If
    
    get分段加成售价 = True
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub SetInputFormat(ByVal intRow As Integer)
    If mint编辑状态 = 5 Then '修改发票信息
        '未付款的记录允许修改
        mshBill.ColData(mconintcol发票号) = IIf(mshBill.RowData(intRow) = 0, 4, 0)
        mshBill.ColData(mconintcol发票代码) = IIf(mshBill.RowData(intRow) = 0, 4, 0)
        mshBill.ColData(mconIntCol发票日期) = IIf(mshBill.RowData(intRow) = 0, 2, 0)
        mshBill.ColData(mconintcol发票金额) = IIf(mshBill.RowData(intRow) = 0, 4, 0)

        Exit Sub
    End If
    
    If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
        If mshBill.TextMatrix(intRow, mconIntCol原销期) <> "" Then
            mshBill.ColData(mconintCol零售价) = 5
            If Val(Split(mshBill.TextMatrix(intRow, mconIntCol原销期), "||")(2)) = 1 And Val(mshBill.TextMatrix(intRow, mconIntCol分批属性)) = 1 Then
                mshBill.ColData(mconintCol零售价) = 4
            End If
        End If
    End If
    
    If mint编辑状态 = 9 Or mint编辑状态 = 3 Or mint编辑状态 = 7 Then
        If mshBill.TextMatrix(intRow, mconIntCol原销期) <> "" Then
            mshBill.ColData(mconIntCol效期) = 2                '日期输入框
            '如果是时价药品，则允许输入售价
            If InStr(1, mstrControlItem, ",售价,") > 0 Then
                If Split(mshBill.TextMatrix(intRow, mconIntCol原销期), "||")(2) = 1 Then
                    mshBill.ColData(mconIntCol售价) = IIf(Get时价药品直接确定售价, 4, 5)
                Else
                    mshBill.ColData(mconIntCol售价) = 5
                End If
            End If
        Else
            mshBill.ColData(mconIntCol效期) = 5
        End If
    End If
    
    If mblnEdit = False Then Exit Sub
    With mshBill
        If mint编辑状态 = 9 Or mint编辑状态 = 3 Or mint编辑状态 = 7 Or mint编辑状态 = 8 Or mbln退货 Then Exit Sub
        
        If mint编辑状态 = 1 Then
            .ColData(mconIntCol产地) = 1
            .ColData(mconIntCol原产地) = 1
        End If
        
        If .TextMatrix(intRow, mconIntCol原销期) <> "" Then
            .ColData(mconIntCol效期) = 2                '日期输入框
            '如果是时价药品，则允许输入售价
            If Split(.TextMatrix(intRow, mconIntCol原销期), "||")(2) = 1 Then
                .ColData(mconIntCol售价) = IIf(Get时价药品直接确定售价, 4, 5)
            Else
                .ColData(mconIntCol售价) = 5
            End If
        Else
            .ColData(mconIntCol效期) = 5
        End If
        
        If Trim(.TextMatrix(intRow, mconintcol发票号)) = "" Then
            .ColData(mconintcol发票代码) = 5
            .ColData(mconIntCol发票日期) = 5
            .ColData(mconintcol发票金额) = 5
        Else
            .ColData(mconintcol发票代码) = 4
            .ColData(mconIntCol发票日期) = 2
            .ColData(mconintcol发票金额) = 4
        End If
        
    End With
End Sub

'设置折扣
Private Sub SetDisCount(ByVal intRow As Integer, ByVal intDisCount As Double)
    Dim dbl加成率 As Double
    Dim dbl成本价 As Double
    
    With mshBill
        '取原来成本价
        dbl成本价 = IIf(mint时价入库售价加成方式 = 0, Val(.TextMatrix(.Row, mconIntCol成本价)), Val(.TextMatrix(.Row, mconIntCol采购价)))
        
        If mbln加价率 Then
            mdbl加价率 = 15
            If Val(.TextMatrix(.Row, mconIntCol售价)) <> 0 And dbl成本价 <> 0 Then
                mdbl加价率 = 计算加成率(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol售价)), dbl成本价)
            End If
        End If
        If mshBill.Col = mconIntCol指导批发价 Then
            intDisCount = Val(.TextMatrix(intRow, mconIntCol扣率))
        Else
            .TextMatrix(intRow, mconIntCol扣率) = intDisCount
        End If
        
        If .TextMatrix(intRow, mconIntCol指导批发价) <> "" Then
'            If .TextMatrix(intRow, mconIntCol采购价) = "" Then
'                .TextMatrix(intRow, mconIntCol采购价) = .TextMatrix(intRow, mconIntCol指导批发价)
'            End If
            If Not (mint编辑状态 = 8 Or mbln退货) Then
                .TextMatrix(intRow, mconIntCol成本价) = zlStr.FormatEx((Val(.TextMatrix(intRow, mconIntCol采购价)) * intDisCount / 100), mintCostDigit, , True)
            End If
            If .TextMatrix(intRow, mconIntCol数量) <> "" Then
               .TextMatrix(intRow, mconIntCol成本金额) = zlStr.FormatEx((.TextMatrix(intRow, mconIntCol数量) * .TextMatrix(intRow, mconIntCol成本价)), mintMoneyDigit, , True)
               .TextMatrix(intRow, mconintcol发票金额) = IIf(Trim(.TextMatrix(intRow, mconintcol发票号)) = "", "", .TextMatrix(intRow, mconIntCol成本金额))
            End If
            .TextMatrix(intRow, mconintCol差价) = zlStr.FormatEx(IIf(.TextMatrix(intRow, mconIntCol售价金额) = "", 0, .TextMatrix(intRow, mconIntCol售价金额)) - IIf(.TextMatrix(intRow, mconIntCol成本金额) = "", 0, .TextMatrix(intRow, mconIntCol成本金额)), mintMoneyDigit, , True)
            
            '根据参数决定时价药品售价公式中成本价的算法
            dbl成本价 = IIf(mint时价入库售价加成方式 = 0, Val(.TextMatrix(.Row, mconIntCol成本价)), Val(.TextMatrix(.Row, mconIntCol采购价)))
            
            '对时价药品的处理
            If .TextMatrix(intRow, mconIntCol原销期) <> "" Then
                If Split(.TextMatrix(intRow, mconIntCol原销期), "||")(2) = 1 Then
                    '由于存在差价让利比的存在,需要按加成率计算,因此将指导差价率转换成加成率计算 公式：加成率=1/(1-差价率)-1
                    If mbln加价率 Then
                        If gtype_UserSysParms.P183_时价取上次售价 <> 1 Then  '没有勾选时价取上次售价参数
                            .TextMatrix(intRow, mconIntCol售价) = zlStr.FormatEx(时价药品零售价(Val(.TextMatrix(.Row, 0)), dbl成本价, (mdbl加价率 / 100), dbl成本价 * (1 + (mdbl加价率 / 100))), mintPriceDigit, , True)
                        End If
                    Else
                        dbl加成率 = Val(Replace(.TextMatrix(intRow, mconIntcol加成率), "%", "")) / 100
                        If gtype_UserSysParms.P183_时价取上次售价 <> 1 Then  '没有勾选时价取上次售价参数
                            .TextMatrix(intRow, mconIntCol售价) = zlStr.FormatEx(时价药品零售价(Val(.TextMatrix(.Row, 0)), dbl成本价, dbl加成率, dbl成本价 * (1 + dbl加成率)), mintPriceDigit, , True)
                        End If
                    End If
                    If .TextMatrix(intRow, mconIntCol数量) <> "" Then
                        .TextMatrix(intRow, mconIntCol售价金额) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol数量) * .TextMatrix(intRow, mconIntCol售价), mintMoneyDigit, , True)
                        'Modified by ZYB  ##2002-10-24
                        '###########################################################
                        .TextMatrix(intRow, mconintCol差价) = zlStr.FormatEx(IIf(.TextMatrix(intRow, mconIntCol售价金额) = "", 0, .TextMatrix(intRow, mconIntCol售价金额)) - IIf(.TextMatrix(intRow, mconIntCol成本金额) = "", 0, .TextMatrix(intRow, mconIntCol成本金额)), mintMoneyDigit, , True)
                        '###########################################################
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub mshBill_LeaveCell(Row As Long, Col As Long)
    If mblnEnter Then OS.OpenIme
    mshBill.Redraw = False
    If mbln效期提示 Then
        With mshBill
            If .Col = mconIntCol效期 Then
                CheckLapse (mshBill.TextMatrix(mshBill.Row, mconIntCol效期))
            End If
        End With
    End If
End Sub

Private Sub mshBill_LostFocus()
    OS.OpenIme
End Sub

Private Sub mshBill_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        With mshBill
           If .Col = mconIntCol药名 Then
                PopupMenu mnuCol, 2
            End If
        End With
    End If
End Sub



Private Sub msh产地_DblClick()
    msh产地_KeyDown vbKeyReturn, 0
End Sub

Private Sub msh产地_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsProvider As ADODB.Recordset
    
    On Error GoTo errHandle
    With mshBill
        If KeyCode = vbKeyEscape Then
            msh产地.Visible = False
            .SetFocus
        End If
        
        If KeyCode = vbKeyReturn Then
            If .Col = mconIntCol验收结论 Then
                .TextMatrix(.Row, .Col) = msh产地.TextMatrix(msh产地.Row, 1)
                msh产地.Visible = False
                .SetFocus
                .Col = GetNextEnableCol(.Col)
                Exit Sub
            End If

            .TextMatrix(.Row, .Col) = msh产地.TextMatrix(msh产地.Row, 2)
            
            gstrSQL = "select 批准文号 from 药品生产商对照 where 厂家名称=[1] and 药品id=[2]"
            Set rsProvider = zlDataBase.OpenSQLRecord(gstrSQL, "mshbill_CommandClick", mshBill.TextMatrix(mshBill.Row, mconIntCol产地), mshBill.TextMatrix(mshBill.Row, 0))
            If Not rsProvider.EOF Then
                mshBill.TextMatrix(mshBill.Row, mconIntCol批准文号) = IIf(IsNull(rsProvider!批准文号), "", rsProvider!批准文号)
            Else
                mshBill.TextMatrix(mshBill.Row, mconIntCol批准文号) = ""
            End If
            msh产地.Visible = False
            .Col = mconIntCol批号
            .SetFocus
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub msh产地_LostFocus()
    If msh产地.Visible Then
        msh产地.Visible = False
    End If
End Sub

Private Sub PicInput_LostFocus()
    Dim strActive As String
    strActive = UCase(Me.ActiveControl.Name)
    
    If InStr(1, "CMDYES,CMDNO,TXT加价率", strActive) <> 0 Then
        Exit Sub
    Else
        If strActive = "MSHBILL" Then
            If mbln允许手工输入加成率 = True Then Exit Sub
        End If
    End If
    mbln允许手工输入加成率 = False
    PicInput.Visible = False
End Sub

Private Sub staThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "PY" And staThis.Tag <> "PY" Then
        Logogram staThis, 0
        staThis.Tag = Panel.Key
    ElseIf Panel.Key = "WB" And staThis.Tag <> "WB" Then
        Logogram staThis, 1
        staThis.Tag = Panel.Key
    End If
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
    If KeyAscii = 13 Then
        cmdFind_Click
    End If
End Sub


Private Sub txtNO_Change()
    If txtNO.Locked = True Then
        If mstr单据号 <> "" And mstr单据号 <> txtNO.Text Then
            txtNO.Text = mstr单据号
        End If
    End If
End Sub

Private Sub txtNO_GotFocus()
    If txtNO.Locked = False Then
        txtNO.SelStart = 0
        txtNO.SelLength = Len(txtNO.Text)
    End If
End Sub

Private Sub TxtNO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey (vbKeyTab)
End Sub

'------------------------------------------------------------------
'------------------------------------------------------------------
'-1：表示该列可以选择，是布尔型［"√"，" "］
' 0：表示该列可以选择，但不能修改
' 1：表示该列可以输入，外部显示为按钮选择
' 2：表示该列是日期列，外部显示为按钮选择，弹出是日期选择框
' 3：表示该列是选择列，外部显示为下拉框选择
'4:  表示该列为单纯的文本框供用户输入
'5:  表示该列不允许选择
'-----------------------------------------------------------------
'-----------------------------------------------------------------
Private Sub txtProvider_Change()
    With txtProvider
        .Text = UCase(.Text)
        .SelStart = Len(.Text)
    End With
End Sub

Private Sub txtProvider_GotFocus()
    txtProvider.SelStart = 0
    txtProvider.SelLength = Len(txtProvider.Text)
End Sub

Private Sub txtProvider_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strProviderText As String
    Dim adoProvider As New Recordset
    Dim vRect As RECT, blnCancel As Boolean
    
    vRect = zlControl.GetControlRect(txtProvider.hWnd)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If mint编辑状态 = 4 Then Exit Sub
    
    On Error GoTo errHandle
    With txtProvider
        If Trim(.Text) = "" Then Exit Sub
        strProviderText = UCase(.Text)
        gstrSQL = "Select id,编码,名称,简码 From 供应商 " & _
                  "Where (站点 = [2] Or 站点 is Null) And (To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01' or 撤档时间 is null) " & _
                  "  And 末级=1 And (substr(类型,1,1)=1 Or Nvl(末级,0)=0) " & _
                  "  And (简码 like [1] Or 编码 like [1] or 名称 like [1] )"
        'Set adoProvider = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, _
                            IIf(gstrMatchMethod = "0", "%", "") & strProviderText & "%", gstrNodeNo)
        Set adoProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "供药单位", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, 300, blnCancel, False, True, IIf(gstrMatchMethod = "0", "%", "") & strProviderText & "%", gstrNodeNo)

        If blnCancel = True Then .SetFocus: Exit Sub  '打开选择器时，点Esc不做以下处理
        
        If adoProvider.State = 0 Then
            MsgBox "没有你输入的供药单位，请重输！", vbOKOnly + vbInformation, gstrSysName
            KeyCode = 0
            .SelStart = 0
            .SelLength = Len(.Text)
            .Tag = 0
            Exit Sub
        End If

        .Text = adoProvider!名称
        .Tag = adoProvider!id
        mblnChange = True
        
        
        adoProvider.Close
        mshBill.SetFocus
        mshBill.Col = 1
        mshBill.Row = 1
        
        If CheckQualifications(Val(txtProvider.Tag)) = False Then
            txtProvider.Text = ""
            txtProvider.Tag = "0"
            Exit Sub
        End If
        
        If Val(.Tag) <> mlng供药单位ID And (mint编辑状态 = 8 Or mbln退货) Then
            mlng供药单位ID = Val(txtProvider.Tag)
            mshBill.ClearBill
            mshBill.TextMatrix(1, mconIntCol行号) = "1"
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Function ValidData() As Boolean
    Dim intLop As Integer
    Dim rsStock As New Recordset
    Dim blnStock As Boolean
    Dim strNo As String
    
    ValidData = False
    On Error GoTo errHandle
    gstrSQL = "SELECT count(*) " _
            & "From 部门性质说明 " _
            & "WHERE ((工作性质 LIKE '%药房') OR (工作性质 LIKE '制剂室')) AND 部门id =[1] "
    Set rsStock = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[检测]", cboStock.ItemData(cboStock.ListIndex))
               
    If rsStock.Fields(0) > 0 Then
        blnStock = False
    Else
        blnStock = True
    End If
    
    If txtNO.Locked = False Then
        '新增，且允许修改单据号
        strNo = txtNO.Text
        If strNo = "" Then
            MsgBox "请输入单据号。", vbInformation, gstrSysName
            txtNO.SetFocus
            Exit Function
        End If
        
        If InStr(strNo, "'") > 0 Then
            MsgBox "单据号所输内容中含有非法字符。", vbExclamation, gstrSysName
            txtNO.SetFocus
            Exit Function
        End If
        If LenB(StrConv(strNo, vbFromUnicode)) > 8 Then
            MsgBox "单据号长度不能超过8个字母。", vbExclamation, gstrSysName
            txtNO.SetFocus
            Exit Function
        End If
    Else
        '防止用户强制修改
'        If mstr单据号 <> "" And mstr单据号 <> txtNO.Text Then
'            txtNO.Text = mstr单据号
'        End If
    End If
    
    If CheckQualifications(Val(txtProvider.Tag)) = False Then Exit Function
    
    With mshBill
        If .TextMatrix(1, 0) <> "" Then         '先判有否数据
            If Val(txtProvider.Tag) = 0 Then
                MsgBox "对不起，供药单位不能为空！", vbOKOnly + vbInformation, gstrSysName
                txtProvider.SetFocus
                Exit Function
            End If
            
            If LenB(StrConv(txt摘要.Text, vbFromUnicode)) > txt摘要.MaxLength Then
                MsgBox "摘要超长,最多能输入" & CInt(txt摘要.MaxLength / 2) & "个汉字或" & txt摘要.MaxLength & "个字符!", vbInformation + vbOKOnly, gstrSysName
                txt摘要.SetFocus
                Exit Function
            End If
        
            For intLop = 1 To .rows - 1
                If Trim(.TextMatrix(intLop, mconIntCol药名)) <> "" Then
                    If Trim(Trim(.TextMatrix(intLop, mconIntCol数量))) = "" Then
                        MsgBox "第" & intLop & "行药品的数量为空了，请检查！", vbInformation, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol数量
                        Exit Function
                    End If
                    
                    If Trim(Trim(.TextMatrix(intLop, mconIntCol成本价))) = "" Then
                        MsgBox "第" & intLop & "行药品的采购价为空了，请检查！", vbInformation, gstrSysName
                        .SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol成本价
                        Exit Function
                    End If
                    
                    If LenB(StrConv(Trim(Trim(.TextMatrix(intLop, mconIntCol批号))), vbFromUnicode)) > mintBatchNoLen Then
                        MsgBox "第" & intLop & "行药品的批号超长,最多能输入" & Int(mintBatchNoLen / 2) & "个汉字或" & mintBatchNoLen & "个字符!", vbInformation + vbOKOnly, gstrSysName
                        .SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol批号
                        Exit Function
                    End If
                    
                    If Trim(Trim(.TextMatrix(intLop, mconIntCol成本金额))) = "" Then
                        MsgBox "第" & intLop & "行药品的采购金额为空了，请检查！", vbInformation, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol成本金额
                        Exit Function
                    End If
                    
                    If Trim(Trim(.TextMatrix(intLop, mconIntCol扣率))) = "" Then
                        MsgBox "第" & intLop & "行药品的扣率为空了，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol扣率
                        Exit Function
                    End If
                    
                    If Val(Trim(Trim(.TextMatrix(intLop, mconIntCol扣率)))) >= 1000# Then
                        MsgBox "第" & intLop & "行药品的扣率太大了，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol扣率
                        Exit Function
                    End If
                    
                    If Split(.TextMatrix(intLop, mconIntCol原销期), "||")(0) <> "0" Then
                        If Trim(.TextMatrix(intLop, mconIntCol批号)) = "" Or Trim(.TextMatrix(intLop, mconIntCol效期)) = "" Then
                            MsgBox "第" & intLop & "行的药品是效期药品,请把它的批号及效期" & vbCrLf & "信息完整输入单据中！", vbInformation, gstrSysName
                            mshBill.SetFocus
                            .Row = intLop
                            .MsfObj.TopRow = intLop
                            If .TextMatrix(intLop, mconIntCol批号) = "" Then
                                .Col = mconIntCol批号
                            Else
                                .Col = mconIntCol效期
                            End If
                            Exit Function
                        End If
                    End If
                    
                    '分批药品必须录入产地和批号
                    If Val(.TextMatrix(intLop, mconIntCol分批属性)) = 1 And (.TextMatrix(intLop, mconIntCol产地) = "" Or .TextMatrix(intLop, mconIntCol批号) = "") Then
                        MsgBox "第" & intLop & "行的药品是分批药品,请把它的生产商和批号" & vbCrLf & "信息输入单据中！", vbInformation, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        If .TextMatrix(intLop, mconIntCol产地) = "" Then
                            .Col = mconIntCol产地
                        Else
                            .Col = mconIntCol批号
                        End If
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mconIntCol成本价)) > 9999999999# Then
                        MsgBox "  第" & intLop & "行药品的采购价大于了数据库能够保存的" & vbCrLf & "最大范围9999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol成本价
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mconIntCol数量)) > 9999999999# Then
                        MsgBox "第" & intLop & "行药品的数量大于了数据库能够保存的" & vbCrLf & "最大范围9999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol数量
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mconIntCol成本金额)) > 9999999999999# Then
                        MsgBox "第" & intLop & "行药品的采购金额大于了数据库能够保存的" & vbCrLf & "最大范围9999999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol成本金额
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mconIntCol售价金额)) > 9999999999999# Then
                        MsgBox "第" & intLop & "行药品的售价金额大于了数据库能够保存的" & vbCrLf & "最大范围9999999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol数量
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mconintcol发票金额)) > 1E+15 Then
                        MsgBox "第" & intLop & "行药品的售价金额大于了数据库能够保存的" & vbCrLf & "最大范围999999999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconintcol发票金额
                        Exit Function
                    End If
                    
                    If Trim(.TextMatrix(intLop, mconintcol发票号)) <> "" Then
                        If Trim(.TextMatrix(intLop, mconIntCol发票日期)) = "" Then
                            MsgBox "第" & intLop & "行药品没有输入发票日期，请检查！", vbInformation + vbOKOnly, gstrSysName
                            mshBill.SetFocus
                            .Row = intLop
                            .MsfObj.TopRow = intLop
                            .Col = mconIntCol发票日期
                            .ColData(mconIntCol发票日期) = 2
                            Exit Function
                        End If
                    End If
                    
                    '零差价管理：检查是否存在不满足零差价的药品
                    If gtype_UserSysParms.P275_零差价管理模式 = 2 Then
                        If IsPriceAdjustMod(Val(.TextMatrix(intLop, 0))) = True Then
                            If Val(.TextMatrix(intLop, mconIntCol成本价)) <> Val(.TextMatrix(intLop, mconIntCol售价)) Then
                                MsgBox "第" & intLop & "行药品已启用零差价管理，但入库单据售价和成本价不一致，不能进行业务，请检查！", vbInformation + vbOKOnly, gstrSysName
                                mshBill.SetFocus
                                .Row = intLop
                                .MsfObj.TopRow = intLop
                                Exit Function
                            End If
                        End If
                    End If
                End If
            Next
        Else
            Exit Function
        End If
    End With
    
    
    ValidData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub txtProvider_LostFocus()
    If txtProvider.Text = "" Then
        txtProvider.Tag = "0"
        Exit Sub
    End If
End Sub

Private Sub txtProvider_Validate(Cancel As Boolean)
    If Trim(txtProvider.Text) = "" Then
        If mint编辑状态 = 8 Or mbln退货 Then
            mblnMSH_GetFocus = True
        End If
        txtProvider.Tag = "0"
        Exit Sub
    End If
    
    If CheckQualifications(Val(txtProvider.Tag)) = False Then
        txtProvider.Text = ""
        txtProvider.Tag = "0"
        Exit Sub
    End If
    
    If Val(txtProvider.Tag) <> mlng供药单位ID And (mint编辑状态 = 8 Or mbln退货) Then
        mlng供药单位ID = Val(txtProvider.Tag)
        mshBill.ClearBill
        mshBill.TextMatrix(1, mconIntCol行号) = "1"
    End If
End Sub
Private Function SaveVerifyCard(ByVal strNo As String) As Boolean
    '功能：财务审核时向财务审核记录表中插入数据
    '返回值:true-执行成功 false-执行失败
    Dim str审核日期 As String
    
    On Error GoTo ErrHand
    
    SaveVerifyCard = False
    str审核日期 = Format(Txt审核日期.Caption, "yyyy-mm-dd hh:mm:ss")
    gstrSQL = "zl_药品财务审核_insert("
    '库房id
    gstrSQL = gstrSQL & cboStock.ItemData(cboStock.ListIndex)
    '单据
    gstrSQL = gstrSQL & ",1"
    '冲销no
    gstrSQL = gstrSQL & ",'" & txtNO.Text & "'"
    'newNO
    gstrSQL = gstrSQL & ",'" & strNo & "'"
    '审核人
    gstrSQL = gstrSQL & ",'" & UserInfo.用户姓名 & "'"
    '审核日期
    gstrSQL = gstrSQL & ",to_date('" & Format(mstr审核日期, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS')"
    '备注
    If Trim(txt摘要.Text) = "" Then
        gstrSQL = gstrSQL & "," & "Null" & ")"
    Else
        gstrSQL = gstrSQL & ",'" & txt摘要.Text & "')"
    End If
    
    Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
    SaveVerifyCard = True
    Exit Function
    
ErrHand:
If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SaveNewCard(ByVal strNo As String) As Boolean
    '功能：财务审核产生新单据用到的函数
    '参数strNO：新单据的No
    '返回值：true-新单据产生成功 false-新单据产生失败
    Dim chrNo As Variant
    Dim lngSerial As Long
    Dim lngStockid As Long
    Dim lng对方部门id As Long
    Dim lngProviderId As Long
    Dim lngDrugID As Long
    Dim strBatchNo As String
    Dim lngBatchNO As Long
    Dim strProducingArea As String
    Dim strOldProducingArea As String
    Dim datTimeLimit As String
    Dim dblQuantity As Double
    Dim dblPurchasePrice As Double
    Dim dblPurchaseMoney As Double
    Dim dblDiscount As Double
    Dim dblSalePrice As Double
    Dim dblSaleMoney As Double
    Dim dblMistakePrice As Double
    Dim strBrief As String
    Dim strBooker As String
    Dim datBookDate As String
    Dim strModifier As String
    Dim datModifyDate As String
    Dim strAssessor As String
    Dim datAssessDate As String
    Dim dbl金额差 As Double
    Dim dbl加成率 As Double
    
    Dim str外观 As String
    Dim str验收结论 As String
    Dim str产品合格证 As String
    Dim Str发票号 As String
    Dim dat发票日期 As String
    Dim dbl发票金额 As Double
    Dim str指导批发价 As String
    Dim intRow As Integer
    Dim str批准文号 As String
    Dim str随货单号 As String
    Dim str随货日期 As String
    Dim str发票代码 As String
    
    Dim str核查人 As String
    Dim str核查日期 As String
    
    Dim datTimeProduct As String
    
    Dim lngRow As Integer
    Dim m As Integer
    Dim lng药品ID As Long
    Dim lng批次 As Long
    Dim i As Integer
    Dim arrSql As Variant
    
    SaveNewCard = False
    arrSql = Array()
    
    On Error GoTo errHandle
    With mshBill
        lngStockid = cboStock.ItemData(cboStock.ListIndex)
        lng对方部门id = Val(chk转入移库.Tag)
        lngProviderId = txtProvider.Tag
        strBrief = txt摘要.Text
        strBooker = Txt填制人.Caption
        datBookDate = Format(Txt填制日期.Caption, "yyyy-mm-dd hh:mm:ss")
'        strModifier = Txt修改人.Caption
'        datModifyDate = Format(Txt修改日期.Caption, "yyyy-mm-dd hh:mm:ss")
        str核查人 = txt核查人.Caption
        str核查日期 = Format(txt核查日期.Caption, "yyyy-mm-dd hh:mm:ss")
    
        '按药品ID顺序更新数据
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                lngSerial = .TextMatrix(intRow, mconIntCol序号)
                lngDrugID = .TextMatrix(intRow, 0)
                strProducingArea = .TextMatrix(intRow, mconIntCol产地)
                strOldProducingArea = .TextMatrix(intRow, mconIntCol原产地)
                strBatchNo = .TextMatrix(intRow, mconIntCol批号)
                lngBatchNO = Val(.TextMatrix(intRow, mconIntCol批次))
                datTimeProduct = IIf(Trim(.TextMatrix(intRow, mconIntCol生产日期)) = "", "", .TextMatrix(intRow, mconIntCol生产日期))
                datTimeLimit = IIf(Trim(.TextMatrix(intRow, mconIntCol效期)) = "", "", .TextMatrix(intRow, mconIntCol效期))
                
                If gtype_UserSysParms.P149_效期显示方式 = 1 And datTimeLimit <> "" Then
                    '换算为失效期来保存
                    datTimeLimit = Format(DateAdd("D", 1, datTimeLimit), "yyyy-mm-dd")
                End If
                
                dblQuantity = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol数量) * .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_数量, , True)
                dblDiscount = .TextMatrix(intRow, mconIntCol扣率)
                dbl加成率 = Val(Replace(.TextMatrix(intRow, mconIntcol加成率), "%", "")) / 100
                dblPurchasePrice = Round(.TextMatrix(intRow, mconIntCol成本价) / .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_成本价)
                dblPurchaseMoney = .TextMatrix(intRow, mconIntCol成本金额)
                dblSalePrice = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol售价) / .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_零售价)
                dblSaleMoney = .TextMatrix(intRow, mconIntCol售价金额)
                dblMistakePrice = .TextMatrix(intRow, mconintCol差价)
                
                str外观 = Trim(.TextMatrix(intRow, mconIntCol外观))
                str验收结论 = Trim(.TextMatrix(intRow, mconIntCol验收结论))
                str产品合格证 = Trim(.TextMatrix(intRow, mconintcol产品合格证))
                Str发票号 = Trim(.TextMatrix(intRow, mconintcol发票号))
                str发票代码 = Trim(.TextMatrix(intRow, mconintcol发票代码))
                dat发票日期 = IIf(.TextMatrix(intRow, mconIntCol发票日期) = "", "", .TextMatrix(intRow, mconIntCol发票日期))
                dbl发票金额 = IIf(Trim(.TextMatrix(intRow, mconintcol发票金额)) = "", 0, .TextMatrix(intRow, mconintcol发票金额))
                str批准文号 = IIf(Trim(.TextMatrix(intRow, mconIntCol批准文号)) = "", "", .TextMatrix(intRow, mconIntCol批准文号))
                str随货单号 = IIf(Trim(.TextMatrix(intRow, mconintcol随货单号)) = "", "", .TextMatrix(intRow, mconintcol随货单号))
                str随货日期 = IIf(Trim(.TextMatrix(intRow, mconintcol随货日期)) = "", "", .TextMatrix(intRow, mconintcol随货日期))
                
                '时价分批药品处理
                If Val(Split(.TextMatrix(intRow, mconIntCol原销期), "||")(2)) = 1 And Val(.TextMatrix(intRow, mconIntCol分批属性)) = 1 And Trim(.TextMatrix(intRow, mconintCol零售价)) <> "" Then
                    dblSalePrice = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol零售价)), gtype_UserDrugDigits.Digit_零售价)
                    dblSaleMoney = Val(.TextMatrix(intRow, mconintCol零售金额))
                    dblMistakePrice = Val(.TextMatrix(intRow, mconintCol零售差价))
                    dbl金额差 = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol零售金额)) - Val(.TextMatrix(intRow, mconIntCol售价金额)), mintMoneyDigit, , True)
                End If
                
                '更新药品目录中的指导批发价
                If mbln修改批发价 Then
                    str指导批发价 = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol指导批发价)) / Val(.TextMatrix(intRow, mconIntCol比例系数)), gtype_UserDrugDigits.Digit_成本价)
                    gstrSQL = "zl_药品目录_UpdateCustom(" & lngDrugID & ",'指导批发价=" & str指导批发价 & "')"
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSQL
                End If
                
                gstrSQL = "zl_药品外购_INSERT("
                'NO
                gstrSQL = gstrSQL & "'" & strNo & "'"
                '序号
                gstrSQL = gstrSQL & "," & lngSerial
                '库房ID
                gstrSQL = gstrSQL & "," & lngStockid
                '对方部门ID
                gstrSQL = gstrSQL & "," & IIf(lng对方部门id <= 0, "null", lng对方部门id)
                '供药单位ID
                gstrSQL = gstrSQL & "," & lngProviderId
                '药品ID
                gstrSQL = gstrSQL & "," & lngDrugID
                '产地
                gstrSQL = gstrSQL & ",'" & strProducingArea & "'"
                '批号
                gstrSQL = gstrSQL & ",'" & strBatchNo & "'"
                '效期
                gstrSQL = gstrSQL & "," & IIf(datTimeLimit = "", "Null", "to_date('" & Format(datTimeLimit, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                '实际数量
                gstrSQL = gstrSQL & "," & dblQuantity
                '成本价
                gstrSQL = gstrSQL & "," & dblPurchasePrice
                '成本金额
                gstrSQL = gstrSQL & "," & dblPurchaseMoney
                '扣率
                gstrSQL = gstrSQL & "," & dblDiscount
                '零售价
                gstrSQL = gstrSQL & "," & dblSalePrice
                '零售金额
                gstrSQL = gstrSQL & "," & dblSaleMoney
                '差价
                gstrSQL = gstrSQL & "," & dblMistakePrice
                '摘要
                gstrSQL = gstrSQL & ",'" & strBrief & "'"
                '填制人
                gstrSQL = gstrSQL & ",'" & strBooker & "'"
                '发票号
                gstrSQL = gstrSQL & ",'" & Str发票号 & "'"
                '发票日期
                gstrSQL = gstrSQL & "," & IIf(dat发票日期 = "", "Null", "to_date('" & Format(dat发票日期, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                '发票金额
                gstrSQL = gstrSQL & "," & dbl发票金额
                '填制日期
                gstrSQL = gstrSQL & ",to_date('" & datBookDate & "','yyyy-mm-dd HH24:MI:SS')"
                '外观
                gstrSQL = gstrSQL & ",'" & str外观 & "'"
                '产品合格证
                gstrSQL = gstrSQL & ",'" & str产品合格证 & "'"
                '核查人
                gstrSQL = gstrSQL & "," & IIf(str核查人 <> "", "'" & str核查人 & "'", "NULL")
                '核查日期
                gstrSQL = gstrSQL & "," & IIf(str核查人 <> "", "to_date('" & str核查日期 & "','yyyy-mm-dd HH24:MI:SS')", "NULL")
                '批次
                gstrSQL = gstrSQL & "," & lngBatchNO
                '是否退货
                gstrSQL = gstrSQL & "," & IIf(mbln退货, -1, 1)
                '生产日期
                gstrSQL = gstrSQL & "," & IIf(datTimeProduct = "", "Null", "to_date('" & Format(datTimeProduct, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                '批准文号
                gstrSQL = gstrSQL & ",'" & str批准文号 & "'"
                '随货单号
                gstrSQL = gstrSQL & ",'" & str随货单号 & "'"
                '金额差
                gstrSQL = gstrSQL & "," & IIf(dbl金额差 <> 0, dbl金额差, "NULL")
                '加成率
                gstrSQL = gstrSQL & "," & dbl加成率
                '发票代码
                gstrSQL = gstrSQL & "," & IIf(str发票代码 = "", "NULL", "'" & str发票代码 & "'")
                '计划id
                gstrSQL = gstrSQL & ",NULL"
                '财务审核
                gstrSQL = gstrSQL & ",2"
                '原产地
                gstrSQL = gstrSQL & ",'" & strOldProducingArea & "'"
                '随货日期
                gstrSQL = gstrSQL & "," & IIf(str随货日期 <> "", "to_date('" & str随货日期 & "','yyyy-mm-dd HH24:MI:SS')", "Null")
                '验收结论
                gstrSQL = gstrSQL & ",'" & str验收结论 & "'"
                '修改人
                gstrSQL = gstrSQL & ",'" & strModifier & "'"
                '修改日期
                gstrSQL = gstrSQL & "," & IIf(datModifyDate = "", "Null", "to_date('" & datModifyDate & "','yyyy-mm-dd HH24:MI:SS')")
                gstrSQL = gstrSQL & ")"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            End If
        Next
        
        For i = 0 To UBound(arrSql)
            Call zlDataBase.ExecuteProcedure(CStr(arrSql(i)), "SaveNewCard")
        Next
        
        SaveNewCard = True
        mstr单据号 = chrNo
        
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SaveCard(Optional ByVal bln强制保存 As Boolean = False) As Boolean
    Dim chrNo As Variant
    Dim lngSerial As Long
    Dim lngStockid As Long
    Dim lng对方部门id As Long
    Dim lngProviderId As Long
    Dim lngDrugID As Long
    Dim strBatchNo As String
    Dim lngBatchNO As Long
    Dim strProducingArea As String
    Dim strOldProducingArea As String
    Dim datTimeLimit As String
    Dim dblQuantity As Double
    Dim dblPurchasePrice As Double
    Dim dblPurchaseMoney As Double
    Dim dblDiscount As Double
    Dim dblSalePrice As Double
    Dim dblSaleMoney As Double
    Dim dblMistakePrice As Double
    Dim strBrief As String
    Dim strBooker As String
    Dim datBookDate As String
    Dim strModifier As String
    Dim datModifyDate As String
    Dim strAssessor As String
    Dim datAssessDate As String
    Dim dbl金额差 As Double
    Dim dbl加成率 As Double
    
    Dim str外观 As String
    Dim str验收结论 As String
    Dim str产品合格证 As String
    Dim Str发票号 As String
    Dim dat发票日期 As String
    Dim dbl发票金额 As Double
    Dim str指导批发价 As String
    Dim intRow As Integer
    Dim str批准文号 As String
    Dim str随货单号 As String
    Dim str随货日期 As String
    Dim str发票代码 As String
    Dim lng计划id As Long
    
    Dim str核查人 As String
    Dim str核查日期 As String
    Dim rsTemp As New ADODB.Recordset
    
    Dim datTimeProduct As String
    
    Dim n As Integer
    Dim m As Integer
    Dim dbl合计数量 As Double
    Dim lng药品ID As Long
    Dim lng批次 As Long
    Dim i As Integer
    Dim arrSql As Variant
    
    SaveCard = False
    arrSql = Array()
    If Not Check合同单位 Then Exit Function
    If Not CheckProvider Then Exit Function
    
    On Error GoTo errHandle
    With mshBill
        chrNo = Trim(txtNO)
        If chrNo = "" Then chrNo = Sys.GetNextNo(21, Me.cboStock.ItemData(Me.cboStock.ListIndex))
        If IsNull(chrNo) Then Exit Function
        
        If mint编辑状态 = 1 Then
            If CheckNOExists(1, chrNo) Then
                MsgBox "存在相同单据号的外购入库单，请检查单据号是否正确！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        Me.txtNO.Tag = chrNo
        lngStockid = cboStock.ItemData(cboStock.ListIndex)
        lngProviderId = txtProvider.Tag
        
        strBrief = Trim(txt摘要.Text)
        strBooker = Txt填制人
        

        If mint编辑状态 = 9 Then '9-核查
            '修改信息
            strModifier = Txt修改人
            datModifyDate = Format(Txt修改日期, "yyyy-mm-dd hh:mm:ss")
            
            If IsDate(Txt填制日期) Then
                datBookDate = Format(Txt填制日期.Caption, "yyyy-mm-dd hh:mm:ss")
            Else
                datBookDate = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            End If
        Else
            datBookDate = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        End If
        
        
        strAssessor = Txt审核人
        
        '取原始单据的核查人
        If mint编辑状态 <> 9 Then
            gstrSQL = "Select 配药人,to_Char(配药日期,'yyyy-MM-dd hh24:mi:ss') 配药日期 " & _
                " From 药品收发记录 Where 单据=1 And NO=[1] "
            Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[取原始单据的核查人]", chrNo)
                
            If Not rsTemp.EOF Then
                str核查人 = nvl(rsTemp!配药人)
                str核查日期 = nvl(rsTemp!配药日期)
            End If
            If mint编辑状态 = 2 Then
                '修改单据则清除核查人，需要再次核查
                str核查人 = ""
                str核查日期 = ""
            End If
        Else
            str核查人 = Txt审核人.Caption
            str核查日期 = Txt审核日期.Caption
        End If
                
        If mint编辑状态 = 2 Or mint编辑状态 = 9 Or bln强制保存 Then        '修改
            gstrSQL = "zl_药品外购_Delete('" & mstr单据号 & "')"
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = gstrSQL
            
            strBooker = Txt填制人
            datBookDate = Format(Txt填制日期.Caption, "yyyy-mm-dd hh:mm:ss")
            '修改信息
            If mint编辑状态 = 9 Then
                strModifier = Txt修改人
                datModifyDate = Format(Txt修改日期.Caption, "yyyy-mm-dd hh:mm:ss")
            Else
                strModifier = UserInfo.用户姓名
                datModifyDate = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            End If
        End If
            
        lng对方部门id = Val(chk转入移库.Tag)
    
        '按药品ID顺序更新数据
        recSort.Sort = "药品id,批次,序号"
        recSort.MoveFirst
        For n = 1 To recSort.RecordCount
            intRow = recSort!行号
            If .TextMatrix(intRow, 0) <> "" Then
'                If Val(.TextMatrix(intRow, mconIntCol序号)) = 0 Then
'                    lngSerial = intRow
'                Else
'                    lngSerial = Val(.TextMatrix(intRow, mconIntCol序号))
'                End If
                lngSerial = intRow
                
                lngDrugID = .TextMatrix(intRow, 0)
                strProducingArea = .TextMatrix(intRow, mconIntCol产地)
                strOldProducingArea = .TextMatrix(intRow, mconIntCol原产地)
                strBatchNo = .TextMatrix(intRow, mconIntCol批号)
                lngBatchNO = Val(.TextMatrix(intRow, mconIntCol批次))
                datTimeProduct = IIf(Trim(.TextMatrix(intRow, mconIntCol生产日期)) = "", "", .TextMatrix(intRow, mconIntCol生产日期))
                datTimeLimit = IIf(Trim(.TextMatrix(intRow, mconIntCol效期)) = "", "", .TextMatrix(intRow, mconIntCol效期))
                
                If gtype_UserSysParms.P149_效期显示方式 = 1 And datTimeLimit <> "" Then
                    '换算为失效期来保存
                    datTimeLimit = Format(DateAdd("D", 1, datTimeLimit), "yyyy-mm-dd")
                End If
                
                dblQuantity = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol数量) * .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_数量, , True)
                dblDiscount = .TextMatrix(intRow, mconIntCol扣率)
                dbl加成率 = Val(Replace(.TextMatrix(intRow, mconIntcol加成率), "%", "")) / 100
                dblPurchasePrice = Round(.TextMatrix(intRow, mconIntCol成本价) / .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_成本价)
                dblPurchaseMoney = .TextMatrix(intRow, mconIntCol成本金额)
                dblSalePrice = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol售价) / .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_零售价)
                dblSaleMoney = .TextMatrix(intRow, mconIntCol售价金额)
                dblMistakePrice = .TextMatrix(intRow, mconintCol差价)
                
                If Val(Split(.TextMatrix(intRow, mconIntCol原销期), "||")(2)) = 0 And mintUnit <> 4 Then
                    '如果是定价药品，则售价取原始价格保存
                    dblSalePrice = Get售价(Val(Split(.TextMatrix(intRow, mconIntCol原销期), "||")(2)) = 1, lngDrugID, lngStockid, 0)
                                    
                    If gtype_UserSysParms.P275_零差价管理模式 = 2 And IsPriceAdjustMod(lngDrugID) = True Then
                        '如果是实行零差价管理的药品，成本价也要和售价一致
                        dblPurchasePrice = dblSalePrice
                    End If
                End If
                
                str外观 = Trim(.TextMatrix(intRow, mconIntCol外观))
                str验收结论 = Trim(.TextMatrix(intRow, mconIntCol验收结论))
                str产品合格证 = Trim(.TextMatrix(intRow, mconintcol产品合格证))
                Str发票号 = Trim(.TextMatrix(intRow, mconintcol发票号))
                str发票代码 = Trim(.TextMatrix(intRow, mconintcol发票代码))
                dat发票日期 = IIf(.TextMatrix(intRow, mconIntCol发票日期) = "", "", .TextMatrix(intRow, mconIntCol发票日期))
                dbl发票金额 = IIf(Trim(.TextMatrix(intRow, mconintcol发票金额)) = "", 0, .TextMatrix(intRow, mconintcol发票金额))
                str批准文号 = IIf(Trim(.TextMatrix(intRow, mconIntCol批准文号)) = "", "", .TextMatrix(intRow, mconIntCol批准文号))
                str随货单号 = IIf(Trim(.TextMatrix(intRow, mconintcol随货单号)) = "", "", .TextMatrix(intRow, mconintcol随货单号))
                str随货日期 = IIf(Trim(.TextMatrix(intRow, mconintcol随货日期)) = "", "", .TextMatrix(intRow, mconintcol随货日期))
                lng计划id = IIf(Trim(.TextMatrix(intRow, mconIntCol计划id)) = "", 0, Val(.TextMatrix(intRow, mconIntCol计划id)))
                
                '时价分批药品处理
                If Val(Split(.TextMatrix(intRow, mconIntCol原销期), "||")(2)) = 1 And Val(.TextMatrix(intRow, mconIntCol分批属性)) = 1 And Trim(.TextMatrix(intRow, mconintCol零售价)) <> "" Then
                    dblSalePrice = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol零售价)), gtype_UserDrugDigits.Digit_零售价)
                    dblSaleMoney = Val(.TextMatrix(intRow, mconintCol零售金额))
                    dblMistakePrice = Val(.TextMatrix(intRow, mconintCol零售差价))
                    dbl金额差 = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol零售金额)) - Val(.TextMatrix(intRow, mconIntCol售价金额)), mintMoneyDigit, , True)
                End If
  
                '更新药品目录中的指导批发价
                If mbln修改批发价 Then
                    str指导批发价 = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol指导批发价)) / Val(.TextMatrix(intRow, mconIntCol比例系数)), gtype_UserDrugDigits.Digit_成本价)
                    gstrSQL = "zl_药品目录_UpdateCustom(" & lngDrugID & ",'指导批发价=" & str指导批发价 & "')"
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSQL
                End If
                
                gstrSQL = "zl_药品外购_INSERT("
                'NO
                gstrSQL = gstrSQL & "'" & chrNo & "'"
                '序号
                gstrSQL = gstrSQL & "," & lngSerial
                '库房ID
                gstrSQL = gstrSQL & "," & lngStockid
                '对方部门ID
                gstrSQL = gstrSQL & "," & IIf(lng对方部门id <= 0, "null", lng对方部门id)
                '供药单位ID
                gstrSQL = gstrSQL & "," & lngProviderId
                '药品ID
                gstrSQL = gstrSQL & "," & lngDrugID
                '产地
                gstrSQL = gstrSQL & ",'" & strProducingArea & "'"
                '批号
                gstrSQL = gstrSQL & ",'" & strBatchNo & "'"
                '效期
                gstrSQL = gstrSQL & "," & IIf(datTimeLimit = "", "Null", "to_date('" & Format(datTimeLimit, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                '实际数量
                gstrSQL = gstrSQL & "," & dblQuantity
                '成本价
                gstrSQL = gstrSQL & "," & dblPurchasePrice
                '成本金额
                gstrSQL = gstrSQL & "," & dblPurchaseMoney
                '扣率
                gstrSQL = gstrSQL & "," & dblDiscount
                '零售价
                gstrSQL = gstrSQL & "," & dblSalePrice
                '零售金额
                gstrSQL = gstrSQL & "," & dblSaleMoney
                '差价
                gstrSQL = gstrSQL & "," & dblMistakePrice
                '摘要
                gstrSQL = gstrSQL & ",'" & strBrief & "'"
                '填制人
                gstrSQL = gstrSQL & ",'" & strBooker & "'"
                '发票号
                gstrSQL = gstrSQL & ",'" & Str发票号 & "'"
                '发票日期
                gstrSQL = gstrSQL & "," & IIf(dat发票日期 = "", "Null", "to_date('" & Format(dat发票日期, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                '发票金额
                gstrSQL = gstrSQL & "," & dbl发票金额
                '填制日期
                gstrSQL = gstrSQL & ",to_date('" & datBookDate & "','yyyy-mm-dd HH24:MI:SS')"
                '外观
                gstrSQL = gstrSQL & ",'" & str外观 & "'"
                '产品合格证
                gstrSQL = gstrSQL & ",'" & str产品合格证 & "'"
                '核查人
                gstrSQL = gstrSQL & "," & IIf(str核查人 <> "", "'" & str核查人 & "'", "NULL")
                '核查日期
                gstrSQL = gstrSQL & "," & IIf(str核查人 <> "", "to_date('" & str核查日期 & "','yyyy-mm-dd HH24:MI:SS')", "NULL")
                '批次
                gstrSQL = gstrSQL & "," & lngBatchNO
                '是否退货
                gstrSQL = gstrSQL & "," & IIf(mbln退货, -1, 1)
                '生产日期
                gstrSQL = gstrSQL & "," & IIf(datTimeProduct = "", "Null", "to_date('" & Format(datTimeProduct, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                '批准文号
                gstrSQL = gstrSQL & ",'" & str批准文号 & "'"
                '随货单号
                gstrSQL = gstrSQL & ",'" & str随货单号 & "'"
                '金额差
                gstrSQL = gstrSQL & "," & IIf(dbl金额差 <> 0, dbl金额差, "NULL")
                '加成率
                gstrSQL = gstrSQL & "," & dbl加成率
                '发票代码
                gstrSQL = gstrSQL & "," & IIf(str发票代码 = "", "NULL", "'" & str发票代码 & "'")
                '计划id
                gstrSQL = gstrSQL & "," & IIf(lng计划id = 0, "NULL", lng计划id)
                '财务审核
                gstrSQL = gstrSQL & "," & 0
                '原产地
                gstrSQL = gstrSQL & ",'" & strOldProducingArea & "'"
                '随货日期
                gstrSQL = gstrSQL & "," & IIf(str随货单号 <> "", "to_date('" & str随货日期 & "','yyyy-mm-dd HH24:MI:SS')", "NULL")
                '验收结论
                gstrSQL = gstrSQL & ",'" & str验收结论 & "'"
                '修改人
                gstrSQL = gstrSQL & ",'" & strModifier & "'"
                '修改日期
                gstrSQL = gstrSQL & "," & IIf(datModifyDate = "", "Null", "to_date('" & datModifyDate & "','yyyy-mm-dd HH24:MI:SS')")
                gstrSQL = gstrSQL & ")"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            End If
            recSort.MoveNext
        Next
        
        If Not bln强制保存 Then gcnOracle.BeginTrans
        For i = 0 To UBound(arrSql)
            Call zlDataBase.ExecuteProcedure(CStr(arrSql(i)), "SaveCard")
        Next
        If Not bln强制保存 Then gcnOracle.CommitTrans
        mstr单据号 = chrNo
        
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveCard = True
    Exit Function
errHandle:
    If Not bln强制保存 Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'退货
Private Function SaveRestore() As Boolean
    Dim chrNo As Variant
    Dim lngSerial As Long
    Dim lngStockid As Long
    Dim lngProviderId As Long
    Dim lngDrugID As Long
    Dim strBatchNo As String
    Dim strProducingArea As String
    Dim strOldProducingArea As String
    Dim datTimeLimit As String
    Dim dblQuantity As Double
    Dim dblPurchasePrice As Double
    Dim dblPurchaseMoney As Double
    Dim dblDiscount As Double
    Dim dblSalePrice As Double
    Dim dblSaleMoney As Double
    Dim dblMistakePrice As Double
    Dim strBrief As String
    Dim strBooker As String
    Dim datBookDate As String
    Dim strModifier As String
    Dim datModifyDate As String
    Dim strAssessor As String
    Dim datAssessDate As String
    Dim str外观 As String
    Dim str验收结论 As String
    Dim str产品合格证 As String
    Dim Str发票号 As String
    Dim str发票代码 As String
    Dim dat发票日期 As String
    Dim dbl发票金额 As Double
    Dim str指导批发价 As String
    Dim intRow As Integer
    Dim dbl金额差 As Double
    Dim dbl加成率 As Double
    Dim str批准文号 As String
    Dim str核查人 As String
    Dim str核查日期 As String
    Dim rsTemp As New ADODB.Recordset
    
    Dim datTimeProduct As String
    Dim n As Integer
    Dim m As Integer
    Dim dbl合计数量 As Double
    Dim lng药品ID As Long
    Dim lng批次 As Long
    Dim str随货单号 As String
    Dim str随货日期 As String
    Dim i As Integer
    Dim arrSql As Variant
    Dim blnTran As Boolean  '是否开始了事物
    Dim intLop As Integer
    
    On Error GoTo errHandle
    
    SaveRestore = False
    '只有药库才允许使用退货功能
    arrSql = Array()
    If Val(txtProvider.Tag) = 0 Then
        MsgBox "请选择供应商！", vbInformation, gstrSysName
        txtProvider.SetFocus
        Exit Function
    End If
    
    With mshBill
        If .TextMatrix(1, 0) = "" Then Exit Function
        
        chrNo = Trim(txtNO)
        If chrNo = "" Then chrNo = Sys.GetNextNo(21, Me.cboStock.ItemData(Me.cboStock.ListIndex))
        If IsNull(chrNo) Then Exit Function
        Me.txtNO.Tag = chrNo
        lngStockid = cboStock.ItemData(cboStock.ListIndex)
        lngProviderId = Val(txtProvider.Tag)
        strBrief = Trim(txt摘要.Text)
        strBooker = Txt填制人
        datBookDate = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
'        strModifier = Txt修改人
'        datModifyDate = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        strAssessor = Txt审核人
        
        '取原核查人
        gstrSQL = "Select 配药人,to_Char(配药日期,'yyyy-MM-dd hh24:mi:ss') 配药日期 From 药品收发记录 Where 单据=1 And NO=[1] "
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[取原始单据的核查人]", chrNo)
        
        If Not rsTemp.EOF Then
            str核查人 = nvl(rsTemp!配药人)
            str核查日期 = nvl(rsTemp!配药日期)
        End If
        
        On Error GoTo errHandle
            
        '按药品ID顺序更新数据
        recSort.Sort = "药品id,批次,序号"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!行号
            If .TextMatrix(intRow, 0) <> "" Then
                lngDrugID = .TextMatrix(intRow, 0)
                strProducingArea = .TextMatrix(intRow, mconIntCol产地)
                strOldProducingArea = .TextMatrix(intRow, mconIntCol原产地)
                strBatchNo = .TextMatrix(intRow, mconIntCol批号)
                datTimeProduct = IIf(Trim(.TextMatrix(intRow, mconIntCol生产日期)) = "", "", .TextMatrix(intRow, mconIntCol生产日期))
                datTimeLimit = IIf(Trim(.TextMatrix(intRow, mconIntCol效期)) = "", "", .TextMatrix(intRow, mconIntCol效期))
                str批准文号 = IIf(Trim(.TextMatrix(intRow, mconIntCol批准文号)) = "", "", .TextMatrix(intRow, mconIntCol批准文号))
                dblQuantity = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol数量)) * .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_数量, , True)
                dblDiscount = .TextMatrix(intRow, mconIntCol扣率)
                dbl加成率 = Val(Replace(.TextMatrix(intRow, mconIntcol加成率), "%", "")) / 100
                dblPurchasePrice = Round(Val(.TextMatrix(intRow, mconIntCol成本价)) / .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_成本价)
                dblPurchaseMoney = Val(.TextMatrix(intRow, mconIntCol成本金额))
                dblSalePrice = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol售价) / .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_零售价)
                dblSaleMoney = Val(.TextMatrix(intRow, mconIntCol售价金额))
                dblMistakePrice = Val(.TextMatrix(intRow, mconintCol差价))
                lngSerial = intRow
                
                str外观 = Trim(.TextMatrix(intRow, mconIntCol外观))
                str验收结论 = Trim(.TextMatrix(intRow, mconIntCol验收结论))
                str产品合格证 = Trim(.TextMatrix(intRow, mconintcol产品合格证))
                Str发票号 = Trim(.TextMatrix(intRow, mconintcol发票号))
                str发票代码 = Trim(.TextMatrix(intRow, mconintcol发票代码))
                dat发票日期 = IIf(.TextMatrix(intRow, mconIntCol发票日期) = "", "", .TextMatrix(intRow, mconIntCol发票日期))
                dbl发票金额 = IIf(Trim(.TextMatrix(intRow, mconintcol发票金额)) = "", 0, .TextMatrix(intRow, mconintcol发票金额))
                str随货单号 = IIf(Trim(.TextMatrix(intRow, mconintcol随货单号)) = "", "", .TextMatrix(intRow, mconintcol随货单号))
                str随货日期 = IIf(Trim(.TextMatrix(intRow, mconintcol随货日期)) = "", "", .TextMatrix(intRow, mconintcol随货日期))
                
                '时价分批药品处理
                If Val(Split(.TextMatrix(intRow, mconIntCol原销期), "||")(2)) = 1 And Val(.TextMatrix(intRow, mconIntCol分批属性)) = 1 And Trim(.TextMatrix(intRow, mconintCol零售价)) <> "" Then
                    dblSalePrice = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol零售价)), gtype_UserDrugDigits.Digit_零售价)
                    dblSaleMoney = Val(.TextMatrix(intRow, mconintCol零售金额))
                    dblMistakePrice = Val(.TextMatrix(intRow, mconintCol零售差价))
                    dbl金额差 = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol零售金额)) - Val(.TextMatrix(intRow, mconIntCol售价金额)), mintMoneyDigit, , True)
                End If
                
                '更新药品目录中的指导批发价
                If mbln修改批发价 Then
                    str指导批发价 = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol指导批发价)) / Val(.TextMatrix(intRow, mconIntCol比例系数)), gtype_UserDrugDigits.Digit_成本价)
                    gstrSQL = "zl_药品目录_UpdateCustom(" & lngDrugID & ",'指导批发价=" & str指导批发价 & "')"
                    
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSQL
                End If
                
                If dblQuantity = 0 Then
                    MsgBox "第" & lngSerial & "行的退货数量为零，不允许保存单据！", vbInformation, gstrSysName
                    Exit Function
                End If
                
                gstrSQL = "zl_药品外购_INSERT("
                'NO
                gstrSQL = gstrSQL & "'" & chrNo & "'"
                '序号
                gstrSQL = gstrSQL & "," & lngSerial
                '库房ID
                gstrSQL = gstrSQL & "," & lngStockid
                '对方部门ID
                gstrSQL = gstrSQL & ",null"
                '供药单位ID
                gstrSQL = gstrSQL & "," & lngProviderId
                '药品ID
                gstrSQL = gstrSQL & "," & lngDrugID
                '产地
                gstrSQL = gstrSQL & ",'" & strProducingArea & "'"
                '批号
                gstrSQL = gstrSQL & ",'" & strBatchNo & "'"
                '效期
                gstrSQL = gstrSQL & "," & IIf(datTimeLimit = "", "Null", "to_date('" & Format(datTimeLimit, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                '实际数量
                gstrSQL = gstrSQL & "," & dblQuantity
                '成本价
                gstrSQL = gstrSQL & "," & dblPurchasePrice
                '成本金额
                gstrSQL = gstrSQL & "," & dblPurchaseMoney
                '扣率
                gstrSQL = gstrSQL & "," & dblDiscount
                '零售价
                gstrSQL = gstrSQL & "," & dblSalePrice
                '零售金额
                gstrSQL = gstrSQL & "," & dblSaleMoney
                '差价
                gstrSQL = gstrSQL & "," & dblMistakePrice
                '摘要
                gstrSQL = gstrSQL & ",'" & strBrief & "'"
                '填制人
                gstrSQL = gstrSQL & ",'" & strBooker & "'"
                '发票号
                gstrSQL = gstrSQL & ",'" & Str发票号 & "'"
                '发票日期
                gstrSQL = gstrSQL & "," & IIf(dat发票日期 = "", "Null", "to_date('" & Format(dat发票日期, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                '发票金额
                gstrSQL = gstrSQL & "," & dbl发票金额
                '填制日期
                gstrSQL = gstrSQL & ",to_date('" & datBookDate & "','yyyy-mm-dd HH24:MI:SS')"
                '外观
                gstrSQL = gstrSQL & ",'" & str外观 & "'"
                '产品合格证
                gstrSQL = gstrSQL & ",'" & str产品合格证 & "'"
                '核查人
                gstrSQL = gstrSQL & "," & IIf(str核查人 <> "", "'" & str核查人 & "'", "NULL")
                '核查日期
                gstrSQL = gstrSQL & "," & IIf(str核查人 <> "", "to_date('" & str核查日期 & "','yyyy-mm-dd HH24:MI:SS')", "NULL")
                '批次
                gstrSQL = gstrSQL & "," & Val(.TextMatrix(intRow, mconIntCol批次))
                '是否退货
                gstrSQL = gstrSQL & ",-1"
                '生产日期
                gstrSQL = gstrSQL & "," & IIf(datTimeProduct = "", "Null", "to_date('" & Format(datTimeProduct, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                '批准文号
                gstrSQL = gstrSQL & ",'" & str批准文号 & "'"
                '随货单号
                gstrSQL = gstrSQL & ",'" & str随货单号 & "'"
                '金额差
                gstrSQL = gstrSQL & "," & IIf(dbl金额差 <> 0, dbl金额差, "NULL")
                '加成率
                gstrSQL = gstrSQL & "," & dbl加成率
                '发票代码
                gstrSQL = gstrSQL & "," & IIf(str发票代码 = "", "NULL", "'" & str发票代码 & "'")
                '计划id
                gstrSQL = gstrSQL & ",NULL"
                '财务审核
                gstrSQL = gstrSQL & ",0"
                '原产地
                gstrSQL = gstrSQL & ",'" & strOldProducingArea & "'"
                '随货日期
                gstrSQL = gstrSQL & ",NULL"
                '验收结论
                gstrSQL = gstrSQL & ",NULL"
                '修改人
                gstrSQL = gstrSQL & ",'" & strModifier & "'"
                '修改日期
                gstrSQL = gstrSQL & "," & IIf(datModifyDate = "", "Null", "to_date('" & datModifyDate & "','yyyy-mm-dd HH24:MI:SS')")
                gstrSQL = gstrSQL & ")"
                    
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            End If
            recSort.MoveNext
        Next
        
        gcnOracle.BeginTrans
        blnTran = True
        For i = 0 To UBound(arrSql)
            Call zlDataBase.ExecuteProcedure(CStr(arrSql(i)), "SaveRestore")
        Next
        gcnOracle.CommitTrans
        
        mstr单据号 = chrNo
        
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveRestore = True
    Exit Function
errHandle:
    If blnTran = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'保存冲销
Private Function SaveStrike() As Boolean
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
    Dim 发票日期_IN As String
    Dim 发票金额_IN As Double
    Dim intRow As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim n As Integer
    Dim int全部冲销 As Integer
    Dim 摘要_IN As String
    Dim str药品ID As String
    Dim i As Integer
    Dim arrSql As Variant
    Dim str药品 As String
    Dim intNumCol As Integer
    
    arrSql = Array()
    SaveStrike = False
    With mshBill
        '为避免并发操作，重新更新付款标志
        Call Refresh付款标志
        
        '检查冲销数量（符号必须与原始数量相同；已付款的记录不允许冲销；财务审核的单据也不允许冲销）
        For intRow = 1 To .rows - 1
            If Val(.TextMatrix(intRow, mconIntCol冲销数量)) <> 0 Then
                If Not 相同符号(Val(.TextMatrix(intRow, mconIntCol数量)), Val(.TextMatrix(intRow, mconIntCol冲销数量))) Then
                    MsgBox "请输入合法的冲销数量（第" & intRow & "行）！", vbInformation, gstrSysName
                    .MsfObj.TopRow = intRow
                    Exit Function
                End If
                If .RowData(intRow) <> 0 Then
                    MsgBox "第" & intRow & "行的药品已经付款，不允许冲销！", vbInformation, gstrSysName
                    .MsfObj.TopRow = intRow
                    Exit Function
                End If
            End If
        Next
        
        If mint编辑状态 = 6 Then '冲销
            intNumCol = mconIntCol冲销数量
        Else
            intNumCol = mconIntCol数量
        End If
        '检查库存
        If mint库存检查 <> 0 And mint编辑状态 <> 7 And mbln退货 = False Then '退货冲销和财务审核不用检查
            str药品 = CheckNumStock(mshBill, cboStock.ItemData(cboStock.ListIndex), 0, mconIntCol批次, intNumCol, mconIntCol比例系数, 2, , mintNumberDigit)
            If str药品 <> "" Then
                If mbln提示方式 = False Then
                    If mint库存检查 = 1 Then '不足提醒
                        If MsgBox("药品【" & str药品 & "】库存不足，是否继续？", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Exit Function
                        Else
                            mbln提示方式 = True
                        End If
                    ElseIf mint库存检查 = 2 Then '不足禁止
                        MsgBox "药品【" & str药品 & "】库存不足，不能审核！", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            End If
        End If
        
        NO_IN = Trim(txtNO.Tag)
        填制人_IN = UserInfo.用户姓名
        原记录状态_IN = mint记录状态
        摘要_IN = Trim(txt摘要.Text)
        
        On Error GoTo errHandle
        
        行次_IN = 0
        
        '按药品ID顺序更新数据
        recSort.Sort = "药品id,批次,序号"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!行号
            If .TextMatrix(intRow, 0) <> "" And (Val(.TextMatrix(intRow, mconIntCol冲销数量)) <> 0 Or mint编辑状态 = 7) Then
                行次_IN = 行次_IN + 1
                药品ID_IN = .TextMatrix(intRow, 0)
                str药品ID = IIf(str药品ID = "", "", str药品ID & ",") & 药品ID_IN
                If Val(.TextMatrix(intRow, mconIntCol冲销数量)) = Val(.TextMatrix(intRow, mconIntCol数量)) Then
                    int全部冲销 = 1
                Else
                    int全部冲销 = 0
                End If
                冲销数量_IN = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol冲销数量)) * .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_数量, , True)
                发票号_IN = Trim(.TextMatrix(intRow, mconintcol发票号))
                发票代码_In = Trim(.TextMatrix(intRow, mconintcol发票代码))
                发票日期_IN = IIf(.TextMatrix(intRow, mconIntCol发票日期) = "", "", .TextMatrix(intRow, mconIntCol发票日期))
                发票金额_IN = IIf(Trim(.TextMatrix(intRow, mconintcol发票金额)) = "", 0, .TextMatrix(intRow, mconintcol发票金额))
                序号_IN = .TextMatrix(intRow, mconIntCol序号)
                
                gstrSQL = "ZL_药品外购_STRIKE("
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
                '冲销数量
                gstrSQL = gstrSQL & "," & 冲销数量_IN
                '填制人
                gstrSQL = gstrSQL & ",'" & 填制人_IN & "'"
                '填制日期
                gstrSQL = gstrSQL & ",to_date('" & Format(mstr审核日期, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS')"
                '发票号
                gstrSQL = gstrSQL & "," & IIf(发票号_IN = "", "Null", "'" & 发票号_IN & "'")
                '发票金额
                gstrSQL = gstrSQL & "," & 发票金额_IN
                '是否全部冲销
                gstrSQL = gstrSQL & "," & IIf(mint编辑状态 = 7 Or int全部冲销 = 1, 1, 0)
                '是否财务审核
                gstrSQL = gstrSQL & "," & IIf(mint编辑状态 = 7, 1, 0)
                '摘要
                gstrSQL = gstrSQL & ",'" & 摘要_IN & "'"
                '发票代码
                gstrSQL = gstrSQL & "," & IIf(发票代码_In = "", "NULL", "'" & 发票代码_In & "'")
                gstrSQL = gstrSQL & ")"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            End If
            recSort.MoveNext
        Next
        
        If mint编辑状态 <> 7 Then gcnOracle.BeginTrans
        For i = 0 To UBound(arrSql)
            Call zlDataBase.ExecuteProcedure(CStr(arrSql(i)), "SaveCard")
        Next
        If mint编辑状态 <> 7 Then gcnOracle.CommitTrans
        
        If 行次_IN = 0 Then
            MsgBox "没有选择一行药品来冲销，请录入冲销数量！", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        
        '提示停用药品
        If str药品ID <> "" Then
            Call CheckStopMedi(str药品ID)
        End If
        
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveStrike = True
    Exit Function
errHandle:
    If mint编辑状态 <> 7 Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SaveRecipe(Optional ByVal strNewNO As String = "") As Boolean
    Dim chrNo As String
    Dim lng序号 As Long
    Dim Str发票号 As String
    Dim str发票代码 As String
    Dim dat发票日期 As String
    Dim dbl发票金额 As Double
    Dim int操作标志 As Integer '1、未冲销单据修改发票信息; 2、部分冲销单据修改发票信息
    Dim intRow As Integer
    Dim n As Integer
    Dim i As Integer
    Dim arrSql As Variant
    
    arrSql = Array()
    SaveRecipe = False
    '检查是否输入供药单位
    If Val(txtProvider.Tag) = 0 Then
        MsgBox "请选择药品供应商！", vbInformation, gstrSysName
        txtProvider.SetFocus
        Exit Function
    End If
    If Not Check合同单位 Then Exit Function
    If Not CheckProvider Then Exit Function
        
    With mshBill
        If strNewNO = "" Then
            chrNo = Trim(txtNO)
        Else
            chrNo = strNewNO
        End If
        
        On Error GoTo errHandle
        
        '为避免并发操作，重新更新付款标志
        Call Refresh付款标志
        
        '检查入库单据是否已付款
        For intRow = 1 To .rows - 1
            If Val(.TextMatrix(intRow, mconIntCol数量)) <> 0 Then
                If .RowData(intRow) <> 0 Then
                    MsgBox "第" & intRow & "行的药品已经付款，不能修改该药品的发票信息！", vbInformation, gstrSysName
                End If
            End If
        Next
                
        
        If mint编辑状态 = 5 Then
            If mint记录状态 = 1 Then
                int操作标志 = 1
            Else
                int操作标志 = 2
            End If
        ElseIf mint编辑状态 = 7 Then
            int操作标志 = 1
        End If
        
        '按药品ID顺序更新数据
        recSort.Sort = "药品id,批次,序号"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!行号
            If .TextMatrix(intRow, 0) <> "" Then
                If .RowData(intRow) = 0 Then
'                    If strNewNO = "" Then
'                        lng序号 = Val(.TextMatrix(intRow, mconIntCol序号))
'                    Else
'                        lng序号 = intRow
'                    End If
                    lng序号 = Val(.TextMatrix(intRow, mconIntCol序号))
                    Str发票号 = Trim(.TextMatrix(intRow, mconintcol发票号))
                    str发票代码 = Trim(.TextMatrix(intRow, mconintcol发票代码))
                    dat发票日期 = IIf(.TextMatrix(intRow, mconIntCol发票日期) = "", "", .TextMatrix(intRow, mconIntCol发票日期))
                    dbl发票金额 = IIf(mbln退货, -1, 1) * IIf(Trim(.TextMatrix(intRow, mconintcol发票金额)) = "", 0, .TextMatrix(intRow, mconintcol发票金额))
                    
                    gstrSQL = "zl_药品外购发票信息_UPDATE("
                    'NO
                    gstrSQL = gstrSQL & "'" & chrNo & "'"
                    '序号
                    gstrSQL = gstrSQL & "," & lng序号
                    '发票号
                    gstrSQL = gstrSQL & ",'" & Str发票号 & "'"
                    '发票日期
                    gstrSQL = gstrSQL & "," & IIf(dat发票日期 = "", "Null", "to_date('" & Format(dat发票日期, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                    '发票金额
                    gstrSQL = gstrSQL & "," & dbl发票金额
                    '供药单位ID
                    gstrSQL = gstrSQL & "," & Val(txtProvider.Tag)
                    '操作标志
                    gstrSQL = gstrSQL & "," & int操作标志
                    '发票代码
                    gstrSQL = gstrSQL & ",'" & str发票代码 & "'"
                    gstrSQL = gstrSQL & ")"

                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSQL
                End If
            End If
            
            recSort.MoveNext
        Next
        
        If mint编辑状态 <> 7 Then gcnOracle.BeginTrans
        For i = 0 To UBound(arrSql)
            Call zlDataBase.ExecuteProcedure(CStr(arrSql(i)), "SaveCard")
        Next
        If mint编辑状态 <> 7 Then gcnOracle.CommitTrans
        
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveRecipe = True
    Exit Function
errHandle:
    If mint编辑状态 <> 7 Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub 显示合计金额()
    Dim curTotal As Double, Cur记帐金额 As Double, Cur记帐差价 As Double
    Dim intLop As Integer
    Dim dbl时价分批 As Boolean
    
    curTotal = 0: Cur记帐金额 = 0: Cur记帐差价 = 0:
    
    With mshBill
        For intLop = 1 To .rows - 1
            curTotal = curTotal + Val(.TextMatrix(intLop, mconIntCol成本金额))
'            Cur记帐金额 = Cur记帐金额 + Val(.TextMatrix(intLop, mconIntCol售价金额))
            If .TextMatrix(intLop, mconIntCol原销期) <> "" Then
                If Val(Split(.TextMatrix(intLop, mconIntCol原销期), "||")(2)) = 1 And Val(.TextMatrix(intLop, mconIntCol分批属性)) = 1 Then
                    dbl时价分批 = True
                    Cur记帐金额 = Cur记帐金额 + Val(.TextMatrix(intLop, mconintCol零售金额))
                Else
                    Cur记帐金额 = Cur记帐金额 + Val(.TextMatrix(intLop, mconIntCol售价金额))
                End If
            Else
                Cur记帐金额 = Cur记帐金额 + Val(.TextMatrix(intLop, mconIntCol售价金额))
            End If
        Next
    End With
    
    Cur记帐差价 = Cur记帐金额 - curTotal
    lblPurchasePrice.Caption = "成本金额合计：" & zlStr.FormatEx(curTotal, mintMoneyDigit, , True)
    
    If dbl时价分批 = True Then
        lblSalePrice.Caption = "售价金额(时价分批按零售金额)合计：" & zlStr.FormatEx(Cur记帐金额, mintMoneyDigit, , True)
        lblDifference.Caption = "差价(时价分批按零售差价)合计：" & zlStr.FormatEx(Cur记帐差价, mintMoneyDigit, , True)
    Else
        lblDifference.Caption = "差价合计：" & zlStr.FormatEx(Cur记帐差价, mintMoneyDigit, , True)
        lblSalePrice.Caption = "售价金额合计：" & zlStr.FormatEx(Cur记帐金额, mintMoneyDigit, , True)
    End If
        
End Sub
Private Sub 提示库存数()
    Dim RecTmp As New ADODB.Recordset
    Dim Dbl数量 As Double
    Dim str单位 As String
    Dim intID As Long
    Dim strUnit As String
    Dim strQuantity As String
    Dim bln显示批次库存 As Boolean
    Dim strSQL As String
    
    On Error GoTo errHandle
    If mint编辑状态 = 6 Then
        bln显示批次库存 = True
    End If
    
    If mshBill.TextMatrix(mshBill.Row, mconIntCol药名) = "" Then
        staThis.Panels(2).Text = ""
        Exit Sub
    End If
    If mshBill.TextMatrix(mshBill.Row, 0) = "" Then Exit Sub
    intID = mshBill.TextMatrix(mshBill.Row, 0)
 
    If RecTmp.State = 1 Then RecTmp.Close
    
    Select Case mintUnit
        Case mconint售价单位
            strUnit = "C.计算单位"
            strQuantity = "可用数量"
        Case mconint门诊单位
            strUnit = "B.门诊单位"
            strQuantity = "可用数量/门诊包装"
        Case mconint住院单位
            strUnit = "B.住院单位"
            strQuantity = "可用数量/住院包装"
        Case mconint药库单位
            strUnit = "B.药库单位"
            strQuantity = "可用数量/药库包装"
    End Select

    gstrSQL = " SELECT B.药品ID," & strUnit & " AS 单位,SUM(" & strQuantity & ") AS 数量 " & _
              " FROM 药品库存 A,药品规格 B,收费项目目录 C " & _
              " WHERE A.性质=1 AND A.可用数量<>0 AND A.库房ID=[1] " & _
              " AND A.药品ID=B.药品ID AND B.药品ID=C.ID AND B.药品ID=[2] "
    '如果是冲销，财务审核，退库，则显示该批次的库存
    If bln显示批次库存 = True Then
        gstrSQL = gstrSQL & " AND NVL(A.批次,0)=[3] "
    End If
    gstrSQL = gstrSQL & " GROUP BY B.药品ID," & strUnit
    Set RecTmp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[提示库存数]", cboStock.ItemData(cboStock.ListIndex), intID, Val(mshBill.TextMatrix(mshBill.Row, mconIntCol批次)))
    
    If RecTmp.EOF Then
        staThis.Panels(2).Text = "该" & IIf(bln显示批次库存 = True, "批次", "") & "药品当前库存数为[0]"
        Exit Sub
    End If
    Dbl数量 = IIf(IsNull(RecTmp!数量), 0, RecTmp!数量)
    
    With mshBill
        strSQL = ""
        If .TextMatrix(.Row, mconIntCol付款标志) = "是" And mint编辑状态 = 6 And gtype_UserSysParms.P173_经过标记付款后才能进行付款管理 = 1 Then
            strSQL = "启用【标记后才能付款管理】参数后，已经标记的药品不能冲销！"
        End If
    End With
    
    staThis.Panels(2).Text = "该" & IIf(bln显示批次库存 = True, "批次", "") & "药品当前库存数为[" & FormatEx(Dbl数量, mintNumberDigit) & "]" & RecTmp!单位 & "  " & strSQL
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Txt加价率_GotFocus()
    Txt加价率.SelStart = 0
    Txt加价率.SelLength = Len(Txt加价率)
End Sub

Private Sub Txt加价率_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call cmdYes_Click
End Sub

Private Sub Txt加价率_KeyPress(KeyAscii As Integer)
    If Not (Chr(KeyAscii) >= 0 And Chr(KeyAscii) <= 9 Or KeyAscii = vbKeyBack Or KeyAscii = 46) Then KeyAscii = 0
End Sub

Private Sub Txt加价率_LostFocus()
    Call PicInput_LostFocus
End Sub


Private Sub txt摘要_Change()
    mblnChange = True
End Sub

Private Sub txt摘要_GotFocus()
    OS.OpenIme True
    With txt摘要
        .SelStart = 0
        .SelLength = Len(txt摘要.Text)
    End With
End Sub

Private Sub txt摘要_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey (vbKeyTab)
        KeyCode = 0
    End If
End Sub

Private Sub txt摘要_LostFocus()
    OS.OpenIme
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub


'取指导批发价定价单位的设置值，缺省为0-按售价单位定价，可选为1-按药库单位定价；
Private Function GetUnit() As Integer
   GetUnit = gtype_UserSysParms.P29_指导批发价定价单位
    
End Function

'取时价药品入库时，是否必须输入加价率
Private Function Get时价药品直接确定售价() As Boolean
    Get时价药品直接确定售价 = (gtype_UserSysParms.P76_时价药品直接确定售价 = 1)
    mint时价入库售价加成方式 = gtype_UserSysParms.P126_时价药品售价加成方式
End Function

'取时价药品入库时，是否必须输入加价率
Private Function Get加价率() As Boolean
    Get加价率 = (gtype_UserSysParms.P54_时价药品以加价率入库 = 1)
End Function

'取数据库中批号的长度，这样，程序中的批号长度与数据库中保持一致了
Private Function GetBatchNoLen() As Integer
    Dim rsBatchNolen As New Recordset
    
    On Error GoTo errHandle
    gstrSQL = "select 批号 from 药品收发记录 where rownum<1 "
    Set rsBatchNolen = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "-取批号长度")
    GetBatchNoLen = rsBatchNolen.Fields(0).DefinedSize
    rsBatchNolen.Close
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub 检查成本价()
    Dim dbl成本价 As Double, dbl零售价 As Double
    '如果成本价比零售价还高，提示用户
    With mshBill
        If Val(.TextMatrix(.Row, 0)) = 0 Then Exit Sub
        dbl成本价 = Format(Val(.TextMatrix(.Row, mconIntCol成本价)), "#####0.00000;-#####0.00000;0;")
        dbl零售价 = Format(Val(.TextMatrix(.Row, mconIntCol售价)), "#####0.00000;-#####0.00000;0;")
    End With
    If dbl成本价 > dbl零售价 Then
        MsgBox "提醒：该药品的成本价比零售价还高！", vbInformation, gstrSysName
    End If
End Sub

Private Function CopyCard() As String
    Dim intRow As Integer, intUpdate As Integer
    Dim sin原数量 As Double, sin现数量 As Double
    Dim dbl采购价 As Double, dbl采购金额 As Double, dbl差价 As Double, dbl零售金额 As Double, dbl扣率 As Double
    Dim strNo As Variant
    Dim dbl售价 As Double
    On Error GoTo ErrHand
    
    strNo = Sys.GetNextNo(21, Me.cboStock.ItemData(Me.cboStock.ListIndex))
    If IsNull(strNo) Then Exit Function
    intUpdate = 0
    CopyCard = ""
    
    '复制产生新单据
    gstrSQL = "zl_billcopy(1,'" & txtNO.Tag & "','" & strNo & "')"
    Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
    
'    修改采购价、采购金额及差价（要考虑到存在审核冲销的单据，这时需要修改采购价、采购金额，差价）
    With mshBill
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                dbl采购价 = Val(.TextMatrix(intRow, mconIntCol成本价))
                dbl采购金额 = Val(.TextMatrix(intRow, mconIntCol成本金额))
                dbl差价 = Val(.TextMatrix(intRow, mconintCol差价))
                dbl零售金额 = Val(.TextMatrix(intRow, mconIntCol售价金额))
                dbl扣率 = Val(.TextMatrix(intRow, mconIntCol扣率))
                dbl售价 = Val(.TextMatrix(intRow, mconIntCol售价))
                Call Get数量(txtNO.Tag, Val(.TextMatrix(intRow, mconIntCol序号)), sin原数量)
                If Get数量(strNo, Val(.TextMatrix(intRow, mconIntCol序号)), sin现数量) Then
                    If Abs(sin现数量) > 0 Then
                        '修正数量
                        dbl采购价 = Round(dbl采购价 / Val(.TextMatrix(intRow, mconIntCol比例系数)), gtype_UserDrugDigits.Digit_成本价)
                        dbl售价 = Round(dbl售价 / Val(.TextMatrix(intRow, mconIntCol比例系数)), gtype_UserDrugDigits.Digit_零售价)
                        dbl采购金额 = Val(IIf(mbln退货, -1, 1)) * dbl采购金额
                        dbl差价 = Val(IIf(mbln退货, -1, 1)) * dbl差价
                        dbl零售金额 = Val(IIf(mbln退货, -1, 1)) * dbl零售金额
                        
                        '更新药品收发记录
                        gstrSQL = "zl_Bill_更新信息(1,'" & strNo & "'," & Val(.TextMatrix(intRow, mconIntCol序号)) & ",'成本价','" & dbl采购价 & "')"
                        Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
                        gstrSQL = "zl_Bill_更新信息(1,'" & strNo & "'," & Val(.TextMatrix(intRow, mconIntCol序号)) & ",'成本金额','" & dbl采购金额 & "')"
                        Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
                        gstrSQL = "zl_Bill_更新信息(1,'" & strNo & "'," & Val(.TextMatrix(intRow, mconIntCol序号)) & ",'差价','" & dbl差价 & "')"
                        Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
                        gstrSQL = "zl_Bill_更新信息(1,'" & strNo & "'," & Val(.TextMatrix(intRow, mconIntCol序号)) & ",'零售金额','" & dbl零售金额 & "')"
                        Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
                        gstrSQL = "zl_Bill_更新信息(1,'" & strNo & "'," & Val(.TextMatrix(intRow, mconIntCol序号)) & ",'扣率','" & dbl扣率 & "')"
                        Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
                        gstrSQL = "zl_Bill_更新信息(1,'" & strNo & "'," & Val(.TextMatrix(intRow, mconIntCol序号)) & ",'零售价','" & dbl售价 & "')"
                        Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
                        '更新应付记录
                        gstrSQL = "zl_Bill_更新应付记录('" & strNo & "'," & Val(.TextMatrix(intRow, mconIntCol序号)) & ",'采购价','" & dbl采购价 & "')"
                        Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
                        gstrSQL = "zl_Bill_更新应付记录('" & strNo & "'," & Val(.TextMatrix(intRow, mconIntCol序号)) & ",'采购金额','" & dbl采购金额 & "')"
                        Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)

                        intUpdate = intUpdate + 1
                    End If
                End If
            End If
        Next
    End With

    If intUpdate = 0 Then
        MsgBox "无法完成财务审核，因为该单据已被全部冲销！", vbInformation, gstrSysName
        Exit Function
    End If
    
    CopyCard = strNo
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function Get数量(ByVal strNo As String, ByVal int序号 As Integer, sin数量 As Double) As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = " Select Nvl(实际数量,0) 数量 From 药品收发记录" & _
              " Where 单据=1 And NO=[1] And 序号=[2] ANd (记录状态=1 Or Mod(记录状态,3)=0)"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[读取数量]", strNo, int序号)
    If rsTemp.EOF Then Exit Function
    sin数量 = rsTemp!数量
    Get数量 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckProvider() As Boolean
    Dim lngRow As Long
    Dim str药品 As String
    Dim str招标药品 As String
    Dim rsTemp As New ADODB.Recordset
    '检查供应商是否是招标药品的中标单位
    str药品 = ""
    
    On Error GoTo errHandle
    With mshBill
        For lngRow = 1 To .rows - 1
            If Val(.TextMatrix(lngRow, 0)) <> 0 Then
                str药品 = str药品 & "," & Val(.TextMatrix(lngRow, 0))
            End If
        Next
        If str药品 <> "" Then str药品 = Mid(str药品, 2)
    End With
    
    '以招标药品数减去共有同一个中标单位的招标药品数，如果无记录，则说明正确，否则按记录中的药品ID提示不是合法的中标单位
    gstrSQL = " Select a.药品ID From 药品规格 a " & _
              " Where a.药品ID in (select * from Table(Cast(f_Num2List([2]) As zlTools.t_NumList))) And Nvl(a.招标药品,0)=1" & _
              " Minus" & _
              " Select A.药品ID From " & _
              "     (Select a.药品ID From 药品规格 a " & _
              "     Where a.药品ID in (select * from Table(Cast(f_Num2List([2]) As zlTools.t_NumList))) And Nvl(a.招标药品,0)=1) A,药品中标单位 B" & _
              " Where A.药品ID=B.药品ID And B.单位ID=[1] And (B.撤档时间 is null or B.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
    gstrSQL = " Select '['||A.编码||']'||A.名称 药品名称 " & _
              " From " & _
              "     (Select A.药品ID,C.编码,Nvl(B.名称,C.名称) 名称" & _
              "     From (" & gstrSQL & ") A,收费项目别名 B,收费项目目录 C" & _
              "     Where A.药品ID=B.收费细目ID(+) and A.药品ID=C.ID" & _
              "     and B.性质(+)=3 and B.码类(+)=1) A"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[判断是否是中标单位处采购]", Val(txtProvider.Tag), str药品)
    
    With rsTemp
        str药品 = ""
        Do While Not .EOF
            str药品 = str药品 & "、" & rsTemp!药品名称
            .MoveNext
        Loop
        If str药品 <> "" Then str药品 = Mid(str药品, 2)
    End With
    
    If str药品 <> "" Then
        If mbln招标药品可选择非中标单位入库 = True Then
            If MsgBox("该供药单位不是以下招标药品的中标单位：是否继续？" & vbCrLf & str药品, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        Else
            MsgBox "该供药单位不是以下招标药品的中标单位：" & vbCrLf & str药品, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    CheckProvider = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function 时价药品零售价(ByVal lng药品ID As Long, ByVal sin采购价 As Double, ByVal sin加成率 As Double, ByVal sin售价 As Double, Optional ByVal lngLastRow As Long = -1) As Double
    Dim sin零售价 As Double, sin指导零售价 As Double, sin差价让利比 As Double
    Dim sinTemp指导零售价 As Double
    Dim rsTemp As New ADODB.Recordset
    Dim sin差价让利 As Double
    '时价药品零售价计算公式:采购价*(1+加成率)
    '改为:采购价*(1+加成率)+(指导零售价-采购价*(1+加成率))*(1-差价让利比)
    '由于差价让利比的存在,以前所有按指导差价率计算的地方,均需要将差价率转换成加成率进行计算,此函数用于返回本次公式增加的部分金额：(指导零售价-采购价*(1+加成率))*(1-差价让利比)+售价部分
    If lngLastRow = -1 Then lngLastRow = mshBill.Row
    On Error GoTo errHandle
    gstrSQL = "Select 指导零售价,Nvl(差价让利比,100) 差价让利比 From 药品规格 Where 药品ID=[1] "
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[读取指导零售价]", lng药品ID)
    
    sin指导零售价 = rsTemp!指导零售价
    sinTemp指导零售价 = rsTemp!指导零售价 * Val(mshBill.TextMatrix(lngLastRow, mconIntCol比例系数))
    sin差价让利比 = rsTemp!差价让利比
    
    时价药品零售价 = 0
    If sin差价让利比 = 100 Then
        时价药品零售价 = sin售价
        Exit Function
    End If
    
    If (mint编辑状态 = 8 Or mbln退货) Then
        '如果是退货，则按出库的方式计算售价
        gstrSQL = " Select Nvl(实际数量,0) 实际数量,Nvl(实际金额,0) 实际金额 From 药品库存 " & _
                " Where 性质=1 And 药品ID=[2] And 库房ID=[1] And Nvl(批次,0)=[3] "
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[计算零售价]", cboStock.ItemData(cboStock.ListIndex), lng药品ID, Val(mshBill.TextMatrix(mshBill.Row, mconIntCol批次)))
        
        If rsTemp.RecordCount = 0 Then
            MsgBox "药品库存数据错误（未找到指定药品的库存记录）！", vbInformation, gstrSysName
            时价药品零售价 = sin售价
            Exit Function
        End If
        '肯定有数量，没有数量的话，无法到达此处
        sin差价让利 = rsTemp!实际金额 / rsTemp!实际数量 * Val(mshBill.TextMatrix(lngLastRow, mconIntCol比例系数))
    Else
        sin零售价 = sin采购价 * (1 + sin加成率)
        If sin零售价 / Val(mshBill.TextMatrix(lngLastRow, mconIntCol比例系数)) >= sin指导零售价 Then
            时价药品零售价 = sin售价
            Exit Function
        End If
        sin指导零售价 = sin指导零售价 * Val(mshBill.TextMatrix(lngLastRow, mconIntCol比例系数))
        sin差价让利 = (sin指导零售价 - sin零售价) * (1 - sin差价让利比 / 100)
    End If
    
    时价药品零售价 = IIf(sin差价让利 + sin售价 > sinTemp指导零售价, sinTemp指导零售价, sin差价让利 + sin售价)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function 计算加成率(ByVal lng药品ID As Long, ByVal sin零售价 As Double, ByVal sin成本价 As Double) As Double
    Dim sin指导零售价 As Double, sin差价让利比 As Double
    Dim rsTemp As New ADODB.Recordset
    '根据零售价反算成本价,由于时价药品公式的变化,导致原来计算加成率的公式无效,需重新计算
    '原公式:(零售价/成本价-1)*100
    '现公式的理论:由于零售价是按加成率算出来后,再加上了让利外那部分金额,因此实际按加成率算出的零售价=指导零售价-(指导零售价-零售价)/差价让利比
    '再套用原公式算出实际的加成率
    计算加成率 = 0.15
    
    On Error GoTo errHandle
    gstrSQL = " Select A.指导零售价,Nvl(A.差价让利比,100) 差价让利比,Nvl(B.是否变价,0) 时价 " & _
          " From 药品规格 A,收费项目目录 B " & _
          " Where A.药品ID=B.ID AND A.药品ID=[1] "
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[读取指导零售价]", lng药品ID)
    
    sin指导零售价 = rsTemp!指导零售价
    sin差价让利比 = rsTemp!差价让利比
    If rsTemp!时价 = 0 Then Exit Function
    
    '指导零售价-(指导零售价-零售价)/差价让利比
    sin指导零售价 = sin指导零售价 * Val(mshBill.TextMatrix(mshBill.Row, mconIntCol比例系数))
    If sin差价让利比 <> 100 And sin差价让利比 > 0 Then
        sin零售价 = sin指导零售价 - (sin指导零售价 - sin零售价) / sin差价让利比 * 100
    Else
        sin零售价 = sin指导零售价 - (sin指导零售价 - sin零售价)
    End If
    If sin成本价 = 0 Then
        
        计算加成率 = (sin零售价 / IIf(sin成本价 = 0, 1, sin成本价)) * 100
    Else
        计算加成率 = (Val(sin零售价) / Val(sin成本价) - 1) * 100
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function 校正零售价(ByVal sin零售价 As Double, Optional ByVal lngLastRow As Long = -1) As Double
    '功能：得到按当前单位系数计算出来的指导零售价，如果时价药品计算出来的零售价大于指导零售价，以指导零售价为准
    Dim sin指导零售价 As Double
    Dim rsTemp As New ADODB.Recordset
       
    On Error GoTo errHandle
    If lngLastRow = -1 Then lngLastRow = mshBill.Row
    
    gstrSQL = " Select 指导零售价,Nvl(差价让利比,100) 差价让利比 " & _
              " From 药品规格 Where 药品ID=[1] "
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[读取指导零售价]", Val(mshBill.TextMatrix(lngLastRow, 0)))
    
    sin指导零售价 = rsTemp!指导零售价
    sin指导零售价 = sin指导零售价 * Val(mshBill.TextMatrix(lngLastRow, mconIntCol比例系数))
    
    校正零售价 = IIf(sin零售价 > sin指导零售价, sin指导零售价, sin零售价)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetColumnByUserDefine()
    Dim intColumns As Integer
    Dim strColumn_Selected As String
    Dim strColumn_All As String
    Dim arrColumn_All, arrColumn_Selected, arrColumn_UnSelected, arr总列, arr可设置列
    Dim intCol As Integer, intCols As Integer
    Dim strAllCol As String
    Dim strChange As String
    Dim strOldColName As String, strNewColName As String
    On Error GoTo ErrHand
    
    strColumn_Selected = zlDataBase.GetPara("选择列", glngSys, 模块号.外购入库)
    mstrColumn_UnSelected = zlDataBase.GetPara("屏蔽列", glngSys, 模块号.外购入库)
    strColumn_All = "药名,2|药品来源,4|基本药物,5|药价级别,7|规格,8|生产商,13|原产地,14|单位,15|批号,16|生产日期,17|效期,18|数量,19|指导批发价,22|采购价,23|扣率,24|" & _
                    "成本价,25|成本金额,26|加成率,27|售价,28|售价金额,29|差价,30|零售价,31|零售单位,32|零售金额,33|零售差价,34|批准文号, 35|外观,36|" & _
                    "产品合格证,37|随货单号,38|随货日期,39|验收结论,40|发票号,41|发票代码,42|发票日期,43|发票金额,44"
    If strColumn_Selected <> "" Then
        '兼容老版本处理，列名称变化，格式：老列名,新列名|老列名,新列名...
        strChange = "产地,生产商|结算价,成本价|结算金额,成本金额"
        
        For intCol = 0 To UBound(Split(strChange, "|"))
            strOldColName = Split(Split(strChange, "|")(intCol), ",")(0)
            strNewColName = Split(Split(strChange, "|")(intCol), ",")(1)
            
            If InStr(1, "|" & strColumn_Selected & "|", "|" & strOldColName & "|") <> 0 Then
                strColumn_Selected = Replace("|" & strColumn_Selected & "|", "|" & strOldColName & "|", "|" & strNewColName & "|")
                strColumn_Selected = Left(strColumn_Selected, Len(strColumn_Selected) - 1)
                strColumn_Selected = Mid(strColumn_Selected, 2)
            End If
            
            If InStr(1, "|" & mstrColumn_UnSelected & "|", "|" & strOldColName & "|") <> 0 Then
                mstrColumn_UnSelected = Replace("|" & mstrColumn_UnSelected & "|", "|" & strOldColName & "|", "|" & strNewColName & "|")
                mstrColumn_UnSelected = Left(mstrColumn_UnSelected, Len(mstrColumn_UnSelected) - 1)
                mstrColumn_UnSelected = Mid(mstrColumn_UnSelected, 2)
            End If
        Next
        
        If mstrColumn_UnSelected <> "" Then
            strAllCol = strColumn_Selected & "|" & mstrColumn_UnSelected
        Else
            strAllCol = strColumn_Selected
        End If
        arr总列 = Split(strColumn_All, "|")
        arr可设置列 = Split(strAllCol, "|")
        If UBound(arr总列) <> UBound(arr可设置列) Or InStr(1, "|" & strColumn_Selected & "|", "|生产商|") = 0 Or InStr(1, "|" & mstrColumn_UnSelected & "|", "|生产商|") <> 0 Then
            strColumn_Selected = "药名|药品来源|基本药物|药价级别|规格|生产商|原产地|批号|生产日期|效期|单位|数量|指导批发价|采购价|扣率|成本价|成本金额|加成率|售价|售价金额|差价|批准文号|外观|产品合格证|随货单号|随货日期|验收结论|发票号|发票代码|发票日期|发票金额"
            mstrColumn_UnSelected = "零售价|零售单位|零售金额|零售差价"
            zlDataBase.SetPara "选择列", strColumn_Selected, glngSys, 模块号.外购入库
            zlDataBase.SetPara "屏蔽列", mstrColumn_UnSelected, glngSys, 模块号.外购入库
        End If
    Else
        strColumn_Selected = "药名|药品来源|基本药物|药价级别|规格|生产商|原产地|批号|生产日期|效期|单位|数量|指导批发价|采购价|扣率|成本价|成本金额|加成率|售价|售价金额|差价|批准文号|外观|产品合格证|随货单号|随货日期|验收结论|发票号|发票代码|发票日期|发票金额"
        mstrColumn_UnSelected = "零售价|零售单位|零售金额|零售差价"
        zlDataBase.SetPara "选择列", strColumn_Selected, glngSys, 模块号.外购入库
        zlDataBase.SetPara "屏蔽列", mstrColumn_UnSelected, glngSys, 模块号.外购入库
    End If
    
    '先装入缺省设置
    mconIntCol行号 = 1
    mconIntCol药名 = 2
    mconIntCol商品名 = 3
    mconIntCol来源 = 4
    mconIntCol基本药物 = 5
    mconIntCol序号 = 6
    mconIntCol药价级别 = 7
    mconIntCol规格 = 8
    mconIntCol原生产商 = 9
    mconIntCol原销期 = 10
    mconIntCol比例系数 = 11
    mconintcol简码 = 12
    mconIntCol产地 = 13
    mconIntCol原产地 = 14
    mconIntCol单位 = 15
    mconIntCol批号 = 16
    mconIntCol生产日期 = 17
    mconIntCol效期 = 18
    mconIntCol数量 = 19
    mconIntCol冲销数量 = 20
    mconIntCol批次 = 21
    mconIntCol指导批发价 = 22
    mconIntCol采购价 = 23
    mconIntCol扣率 = 24
    mconIntCol成本价 = 25
    mconIntCol成本金额 = 26
    mconIntcol加成率 = 27
    mconIntCol售价 = 28
    mconIntCol售价金额 = 29
    mconintCol差价 = 30
    mconintCol零售价 = 31
    mconintCol零售单位 = 32
    mconintCol零售金额 = 33
    mconintCol零售差价 = 34
    mconIntCol批准文号 = 35
    mconIntCol外观 = 36
    mconintcol产品合格证 = 37
    mconintcol随货单号 = 38
    mconintcol随货日期 = 39
    mconIntCol验收结论 = 40
    mconintcol发票号 = 41
    mconintcol发票代码 = 42
    mconIntCol发票日期 = 43
    mconintcol发票金额 = 44
    mconIntCol分批属性 = 45
    mconIntCol是否新行 = 46
    mconIntCol药品编码和名称 = 47
    mconIntCol药品编码 = 48
    mconIntCol药品名称 = 49
    mconIntCol付款标志 = 50
    mconIntCol计划id = 51
    
    mintLastCol = 51

    
    '商品名列处理
    If gint药品名称显示 = 2 Then
        '显示商品名列
        mshBill.ColWidth(mconIntCol商品名) = IIf(mshBill.ColWidth(mconIntCol商品名) = 0, 2000, mshBill.ColWidth(mconIntCol商品名))
    Else
        '不单独显示商品名列
        mshBill.ColWidth(mconIntCol商品名) = 0
    End If
    
    '根据用户设置调整列顺序
    arrColumn_All = Split(strColumn_All, "|")
    arrColumn_Selected = Split(strColumn_Selected, "|")
    intCols = UBound(arrColumn_Selected)
    For intCol = 0 To intCols
        Call SetColumnValue(arrColumn_Selected(intCol), Split(arrColumn_All(intCol), ",")(1))
    Next
    
    '将未选择的列的列宽设置为零，且列数据为5——不可选择
    If mstrColumn_UnSelected = "" Then Exit Sub
    intCol = intCols + 1
    intColumns = 0
    arrColumn_UnSelected = Split(mstrColumn_UnSelected, "|")
    intCols = UBound(arrColumn_All)
    For intCol = intCol To intCols
        If UBound(arrColumn_UnSelected) >= intColumns Then
            Call SetColumnValue(arrColumn_UnSelected(intColumns), Split(arrColumn_All(intCol), ",")(1), False)
            intColumns = intColumns + 1
        Else
            Call SetColumnValue(Split(arrColumn_All(intCol), ",")(0), Split(arrColumn_All(intCol), ",")(1), False)
        End If
    Next
    Exit Sub
ErrHand:
    MsgBox "恢复列设置时发生错误，请重新进行列设置！", vbInformation, gstrSysName
End Sub

Private Sub SetColumnValue(ByVal str列名 As String, ByVal intValue As Integer, Optional ByVal blnShow As Boolean = True)
    Select Case str列名
    Case "行号"
        mconIntCol行号 = intValue
    Case "药名"
        mconIntCol药名 = intValue
    Case "药品来源"
        mconIntCol来源 = intValue
    Case "基本药物"
        mconIntCol基本药物 = intValue
    Case "序号"
        mconIntCol序号 = intValue
    Case "规格"
        mconIntCol规格 = intValue
    Case "药价级别"
        mconIntCol药价级别 = intValue
    Case "原生产商"
        mconIntCol原生产商 = intValue
    Case "原销期"
        mconIntCol原销期 = intValue
    Case "比例系数"
        mconIntCol比例系数 = intValue
    Case "简码"
        mconintcol简码 = intValue
    Case "生产商"
        mconIntCol产地 = intValue
    Case "原产地"
        mconIntCol原产地 = intValue
    Case "单位"
        mconIntCol单位 = intValue
    Case "批号"
        mconIntCol批号 = intValue
    Case "生产日期"
        mconIntCol生产日期 = intValue
    Case "效期"
        mconIntCol效期 = intValue
    Case "数量"
        mconIntCol数量 = intValue
    Case "冲销数量"
        mconIntCol冲销数量 = intValue
    Case "指导批发价"
        mconIntCol指导批发价 = intValue
    Case "扣率"
        mconIntCol扣率 = intValue
    Case "成本价"
        mconIntCol成本价 = intValue
    Case "成本金额"
        mconIntCol成本金额 = intValue
    Case "售价"
        mconIntCol售价 = intValue
    Case "售价金额"
        mconIntCol售价金额 = intValue
    Case "差价"
        mconintCol差价 = intValue
    Case "零售价"
        mconintCol零售价 = intValue
    Case "零售单位"
        mconintCol零售单位 = intValue
    Case "零售金额"
        mconintCol零售金额 = intValue
    Case "零售差价"
        mconintCol零售差价 = intValue
    Case "批准文号"
        mconIntCol批准文号 = intValue
    Case "外观"
        mconIntCol外观 = intValue
    Case "产品合格证"
        mconintcol产品合格证 = intValue
    Case "随货单号"
        mconintcol随货单号 = intValue
    Case "随货日期"
        mconintcol随货日期 = intValue
    Case "发票号"
        mconintcol发票号 = intValue
    Case "发票代码"
        mconintcol发票代码 = intValue
    Case "发票日期"
        mconIntCol发票日期 = intValue
    Case "发票金额"
        mconintcol发票金额 = intValue
    Case "采购价"
        mconIntCol采购价 = intValue
    Case "加成率"
        mconIntcol加成率 = intValue
    Case "验收结论"
        mconIntCol验收结论 = intValue
    End Select
    
    If Not blnShow Then
        mshBill.ColWidth(intValue) = 0
        mshBill.ColData(intValue) = 5
    Else
        mintLastCol = intValue
    End If
End Sub

Private Function Check是否存在负数量() As Boolean
    Dim n As Integer
    
    With mshBill
        For n = 1 To .rows - 1
            If Val(.TextMatrix(n, 0)) <> 0 Then
                If Val(.TextMatrix(n, mconIntCol数量)) < 0 Then
                    Check是否存在负数量 = True
                    Exit Function
                End If
            End If
        Next

    End With
End Function

Private Sub RefreshBill()
    '以最新价格最新单据相关数据，用于单据审核时
    Dim lngRow As Long, lngRows As Long, lng药品ID As Long
    Dim Dbl数量 As Double, dbl成本价 As Double, dbl成本金额 As Double, dbl零售价 As Double, dbl零售金额 As Double, dbl差价 As Double
    Dim rsPrice As New ADODB.Recordset
    Dim intPriceDigit As Integer
        
    On Error GoTo errHandle
    intPriceDigit = GetDigit(0, 1, 2, 1)
        
    gstrSQL = " Select 收费细目ID,nvl(现价,0) 现价 From 收费价目 " & _
            " Where (终止日期 Is NULL Or sysdate Between 执行日期 And nvl(终止日期,to_date('3000-01-01','yyyy-MM-dd')))" & _
            GetPriceClassString("")
    gstrSQL = "Select A.序号,A.药品ID,B.现价 From 药品收发记录 A,(" & gstrSQL & ") B,收费项目目录 C" & _
            " Where A.单据=1 And A.NO=[1] And A.药品ID=B.收费细目ID And C.ID=B.收费细目ID And Round(A.零售价," & intPriceDigit & ")<>Round(B.现价," & intPriceDigit & ") And Nvl(C.是否变价,0)=0" & _
            " Union All " & _
            " Select A.序号, A.药品id, decode(x.现价,null,decode(nvl(b.零售价,0),0,b.实际金额 / b.实际数量,b.零售价),x.现价) 现价 " & _
            " From 药品收发记录 A, 药品库存 B, 收费项目目录 C, 药品规格 D , " & _
            "      (Select x.药品id,x.库房id,x.批次,x.现价 from 药品价格记录 x where x.价格类型 = 1 and (x.终止日期 Is Null Or Sysdate Between x.执行日期 And Nvl(x.终止日期, To_Date('3000-01-01', 'yyyy-MM-dd')))) X" & _
            " Where A.单据 = 1 And A.NO = [1] And C.ID = A.药品id And Round(A.零售价, " & intPriceDigit & ") <> Round(decode(x.现价,null,decode(nvl(b.零售价,0),0,b.实际金额 / b.实际数量,b.零售价),x.现价), " & intPriceDigit & ") And " & _
            " Nvl(C.是否变价, 0) = 1 And D.药品id = A.药品id And B.性质 = 1 And B.库房id = A.库房id And B.药品id = A.药品id And " & _
            " a.药品id = x.药品id(+) And a.库房id = x.库房id(+) And Nvl(a.批次, 0) = Nvl(x.批次(+), 0) AND " & _
            " Nvl(B.批次, 0) = Nvl(A.批次, 0) And NVL(b.实际数量, 0) <> 0 And decode(x.现价,null,decode(nvl(b.零售价,0),0,b.实际金额 / b.实际数量,b.零售价),x.现价) > 0 " & _
            " Order by 药品id,序号"
    Set rsPrice = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[取当前价格]", txtNO.Text)
    
    If rsPrice.EOF Then Exit Sub
    
    lngRows = mshBill.rows - 1
    For lngRow = 1 To lngRows
        lng药品ID = Val(mshBill.TextMatrix(lngRow, 0))
        If lng药品ID <> 0 Then
            rsPrice.Filter = "药品ID=" & lng药品ID
            If rsPrice.RecordCount <> 0 Then
                '以当前最新价格最新单据相关数据（单价、零售金额、差价）
                dbl零售价 = rsPrice!现价 * Val(mshBill.TextMatrix(lngRow, mconIntCol比例系数))
                dbl成本价 = Val(mshBill.TextMatrix(lngRow, mconIntCol成本价))
                Dbl数量 = Val(mshBill.TextMatrix(lngRow, mconIntCol数量))
                dbl成本金额 = dbl成本价 * Dbl数量
                dbl零售金额 = dbl零售价 * Dbl数量
                dbl差价 = dbl零售金额 - dbl成本金额
                
                mshBill.TextMatrix(lngRow, mconIntCol售价) = zlStr.FormatEx(dbl零售价, intPriceDigit, , True)
                mshBill.TextMatrix(lngRow, mconIntCol售价金额) = zlStr.FormatEx(dbl零售金额, mintMoneyDigit, , True)
                mshBill.TextMatrix(lngRow, mconintCol差价) = zlStr.FormatEx(dbl差价, mintMoneyDigit, , True)
                mshBill.TextMatrix(lngRow, mconIntcol加成率) = zlStr.FormatEx((Val(mshBill.TextMatrix(lngRow, mconIntCol售价)) / dbl成本价 - 1) * 100, 2) & "%"
            End If
        End If
    Next
    rsPrice.Filter = 0
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetState()
    '控制某个单元格是否可以进行修改
    Dim strTemp As String
    Dim i As Integer
        
    With mshBill
        If .TextMatrix(.Row, mconIntCol付款标志) = "是" And gtype_UserSysParms.P173_经过标记付款后才能进行付款管理 = 1 Then
            If InStr(1, mstrControlItem, "采购价") > 0 Then
                .ColData(mconIntCol采购价) = 5
            Else
                .ColData(mconIntCol采购价) = 4
            End If
            
            If InStr(1, mstrControlItem, "扣率") > 0 Then
                .ColData(mconIntCol扣率) = 5
            Else
                .ColData(mconIntCol扣率) = 4
            End If
            
            If InStr(1, mstrControlItem, "成本价") > 0 Then
                .ColData(mconIntCol成本价) = 5
            Else
                .ColData(mconIntCol成本价) = 4
            End If
            
            If InStr(1, mstrControlItem, "成本金额") > 0 Then
                .ColData(mconIntCol成本金额) = 5
            Else
                .ColData(mconIntCol成本金额) = 4
            End If
            
            If InStr(1, mstrControlItem, "售价") > 0 Then
                .ColData(mconIntCol售价) = 5
            Else
                .ColData(mconIntCol售价) = 4
            End If
        End If
    End With
End Sub

Private Function CheckRedo(ByVal rsTemp As ADODB.Recordset) As ADODB.Recordset
    '功能：将重复的记录过滤掉，并返回过滤后的数据集合

    Dim i As Integer
    Dim strTemp As String
    Dim str批次 As String
    Dim str药品ID As String
    Dim str重复药名 As String
    Dim strDub As String
    Dim strSQL As String
    
    rsTemp.MoveFirst
    str批次 = ""
    Do While Not rsTemp.EOF
        str批次 = IIf(IsNull(rsTemp!批次), "0", rsTemp!批次)
        If InStr(1, strTemp, rsTemp!药品ID & "," & str批次) = 0 Then
            strTemp = strTemp & rsTemp!药品ID & "," & str批次 & "|"
        End If
        rsTemp.MoveNext
    Loop
    
    With mshBill
        For i = 1 To .rows - 1
            If InStr(1, strTemp, .TextMatrix(i, 0) & "," & .TextMatrix(i, mconIntCol批次)) > 0 And .TextMatrix(i, 0) <> "" Then
                str药品ID = str药品ID & .TextMatrix(i, 0) & "," & .TextMatrix(i, mconIntCol药名) & "|"
            End If
        Next
        
        If str药品ID <> "" Then   '为过滤数据拼接sql
            strDub = ""
            For i = 0 To UBound(Split(str药品ID, "|")) - 1
                strDub = strDub & "药品id<>" & Split(Split(str药品ID, "|")(i), ",")(0) & " and "
                If UBound(Split(str重复药名, ",")) <= 2 Then
                    str重复药名 = str重复药名 & Split(Split(str药品ID, "|")(i), ",")(1) & ","
                End If
            Next
            If strDub <> "" Then
                strDub = Mid(strDub, 1, Len(strDub) - 4)
            End If
        End If
        
        If str重复药名 <> "" Then
            MsgBox str重复药名 & "列表中已经含有了！" & vbCrLf & "以上药品不再添加！", vbInformation, gstrSysName
            strSQL = strDub
        End If
        rsTemp.Filter = strSQL
        Set CheckRedo = rsTemp
    End With
End Function

Private Sub vsfInputCost_DblClick()
    If vsfInputCost.rows = 1 Then Exit Sub
    With mshBill
        .SetFocus
        .Row = vsfInputCost.Tag
        .Col = mconIntCol采购价
        .TextMatrix(vsfInputCost.Tag, mconIntCol采购价) = vsfInputCost.TextMatrix(vsfInputCost.Row, vsfInputCost.ColIndex("成本价"))
        .TextMatrix(vsfInputCost.Tag, mconIntCol成本价) = zlStr.FormatEx(Val(.TextMatrix(vsfInputCost.Tag, mconIntCol采购价)) * Val(.TextMatrix(vsfInputCost.Tag, mconIntCol扣率)) / 100, mintCostDigit, , True)
        '设置金额
        If .TextMatrix(vsfInputCost.Tag, mconIntCol数量) <> "" Then
            .TextMatrix(vsfInputCost.Tag, mconIntCol成本金额) = zlStr.FormatEx(.TextMatrix(vsfInputCost.Tag, mconIntCol数量) * Val(.TextMatrix(vsfInputCost.Tag, mconIntCol成本价)), mintMoneyDigit, , True)
            .TextMatrix(vsfInputCost.Tag, mconintcol发票金额) = IIf(Trim(.TextMatrix(vsfInputCost.Tag, mconintcol发票号)) = "", "", .TextMatrix(vsfInputCost.Tag, mconIntCol成本金额))
            .TextMatrix(vsfInputCost.Tag, mconintCol差价) = zlStr.FormatEx(IIf(.TextMatrix(vsfInputCost.Tag, mconIntCol售价金额) = "", 0, .TextMatrix(vsfInputCost.Tag, mconIntCol售价金额)) - IIf(.TextMatrix(vsfInputCost.Tag, mconIntCol成本金额) = "", 0, .TextMatrix(vsfInputCost.Tag, mconIntCol成本金额)), mintMoneyDigit, , True)
            .TextMatrix(vsfInputCost.Tag, mconintCol零售差价) = zlStr.FormatEx(Val(.TextMatrix(vsfInputCost.Tag, mconintCol零售金额)) - Val(.TextMatrix(vsfInputCost.Tag, mconIntCol成本金额)), mintMoneyDigit, , True)
        End If
        
        Call 显示合计金额
        picInputCost.Visible = False
    End With
End Sub

Private Sub GetDefineSize()
    '功能：得到数据库的表字段的长度
    On Error GoTo errHandle
    Dim rsTmp As New ADODB.Recordset
     
    gstrSQL = "Select t.上次产地 as 生产商, t.原产地 as 原产地 From 药品规格 T Where Rownum < 1"
    Call zlDataBase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    
    mlng生产商长度 = rsTmp.Fields("生产商").DefinedSize
    mlng原产地长度 = rsTmp.Fields("原产地").DefinedSize
   
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
