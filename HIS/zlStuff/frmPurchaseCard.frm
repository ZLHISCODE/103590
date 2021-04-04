VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmPurchaseCard 
   Caption         =   "卫材外购入库单"
   ClientHeight    =   6990
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11430
   Icon            =   "frmPurchaseCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   11430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdALLDel 
      Caption         =   "全清(&D)"
      Height          =   350
      Left            =   4080
      TabIndex        =   63
      ToolTipText     =   "清除所有行的发票相关数据"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdBulkCopy 
      Caption         =   "批量复制(&C)"
      Height          =   350
      Left            =   6000
      TabIndex        =   62
      ToolTipText     =   "复制当前行发票信息应用于其他无发票信息行"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "批量复制(&C)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6000
      TabIndex        =   60
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox txtCopy 
      Enabled         =   0   'False
      Height          =   270
      Left            =   7320
      MaxLength       =   100
      TabIndex        =   59
      Text            =   "1"
      Top             =   5805
      Width           =   600
   End
   Begin VB.CommandButton cmdExtractData 
      Caption         =   "提取数据(&E)"
      Height          =   350
      Left            =   1440
      TabIndex        =   58
      Top             =   5760
      Width           =   1215
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
      TabIndex        =   42
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
         TabIndex        =   45
         Text            =   "15.0000"
         Top             =   690
         Width           =   1725
      End
      Begin VB.CommandButton CmdNO 
         Caption         =   "取消"
         Height          =   345
         Left            =   1800
         TabIndex        =   47
         Top             =   1140
         Width           =   855
      End
      Begin VB.CommandButton CmdYes 
         Caption         =   "确定"
         Height          =   345
         Left            =   810
         TabIndex        =   46
         Top             =   1140
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "    请输入加成率，零售价的计算公式：零售价=成本价*(1+加成率%)"
         ForeColor       =   &H00400000&
         Height          =   585
         Left            =   0
         TabIndex        =   43
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
         TabIndex        =   44
         Top             =   750
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdAllCls 
      Caption         =   "全清(&L)"
      Height          =   350
      Left            =   10095
      TabIndex        =   41
      Top             =   6225
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdAllSel 
      Caption         =   "全冲(&A)"
      Height          =   350
      Left            =   8775
      TabIndex        =   16
      Top             =   6225
      Visible         =   0   'False
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh产地 
      Height          =   2175
      Left            =   2520
      TabIndex        =   40
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
      Left            =   4425
      TabIndex        =   23
      Top             =   5775
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "查找(&F)"
      Height          =   350
      Left            =   2760
      TabIndex        =   22
      Top             =   5745
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   225
      TabIndex        =   21
      Top             =   5745
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   8760
      TabIndex        =   15
      Top             =   5745
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   10095
      TabIndex        =   34
      Top             =   5745
      Width           =   1100
   End
   Begin VB.PictureBox Pic单据 
      BackColor       =   &H80000004&
      Height          =   5565
      Left            =   0
      ScaleHeight     =   5505
      ScaleWidth      =   11655
      TabIndex        =   24
      Top             =   0
      Width           =   11715
      Begin VB.PictureBox picCostly 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   6120
         ScaleHeight     =   495
         ScaleWidth      =   5175
         TabIndex        =   55
         Top             =   3120
         Width           =   5175
         Begin VB.TextBox txtTypeVar 
            Height          =   270
            Left            =   3720
            TabIndex        =   57
            Top             =   80
            Width           =   2000
         End
         Begin VB.Label lblType 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000A&
            Caption         =   "检索类型↓"
            Height          =   180
            Left            =   2400
            TabIndex        =   56
            Top             =   120
            Width           =   1140
         End
         Begin VB.Label lblCostly 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "高值材料信息(&1)"
            Height          =   180
            Left            =   120
            TabIndex        =   52
            Top             =   120
            Width           =   1350
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshProvider 
         Height          =   1815
         Left            =   4800
         TabIndex        =   35
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
      Begin VB.TextBox txt摘要 
         Height          =   300
         Left            =   900
         MaxLength       =   1000
         TabIndex        =   6
         Top             =   4080
         Width           =   10410
      End
      Begin VB.CommandButton cmdProvider 
         Caption         =   "…"
         Height          =   300
         Left            =   11160
         TabIndex        =   4
         Top             =   660
         Width           =   300
      End
      Begin VB.TextBox txtProvider 
         Height          =   300
         Left            =   8280
         TabIndex        =   3
         Top             =   660
         Width           =   2895
      End
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   660
         Width           =   1515
      End
      Begin VB.TextBox txtNO 
         Height          =   300
         IMEMode         =   2  'OFF
         Left            =   9960
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   158
         Width           =   1425
      End
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   2805
         Left            =   240
         TabIndex        =   5
         Top             =   1005
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
      Begin VSFlex8Ctl.VSFlexGrid vsfCostlyInfo 
         Height          =   615
         Left            =   1200
         TabIndex        =   54
         Top             =   5200
         Visible         =   0   'False
         Width           =   3375
         _cx             =   5953
         _cy             =   1085
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
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
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
      Begin VB.Label lbl核查日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "核查日期"
         Height          =   180
         Left            =   4035
         TabIndex        =   51
         Top             =   4920
         Width           =   720
      End
      Begin VB.Label lbl核查人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  核查人"
         Height          =   180
         Left            =   4080
         TabIndex        =   50
         Top             =   4515
         Width           =   720
      End
      Begin VB.Label txt核查日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4845
         TabIndex        =   49
         Top             =   4860
         Width           =   1890
      End
      Begin VB.Label txt核查人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4845
         TabIndex        =   48
         Top             =   4455
         Width           =   1890
      End
      Begin VB.Label lblDifference 
         AutoSize        =   -1  'True
         Caption         =   "差价合计:"
         Height          =   180
         Left            =   4920
         TabIndex        =   38
         Top             =   3840
         Width           =   810
      End
      Begin VB.Label lblSalePrice 
         AutoSize        =   -1  'True
         Caption         =   "售价金额合计:"
         Height          =   180
         Left            =   2040
         TabIndex        =   37
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "采购金额合计:"
         Height          =   180
         Left            =   240
         TabIndex        =   36
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label Txt审核人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   9450
         TabIndex        =   32
         Top             =   4500
         Width           =   1890
      End
      Begin VB.Label Txt审核日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   9420
         TabIndex        =   31
         Top             =   4905
         Width           =   1890
      End
      Begin VB.Label Txt填制日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   30
         Top             =   4830
         Width           =   1890
      End
      Begin VB.Label Txt填制人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   29
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
         Left            =   9480
         TabIndex        =   18
         Top             =   195
         Width           =   480
      End
      Begin VB.Label lbl摘要 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "摘要(&M)"
         Height          =   180
         Left            =   240
         TabIndex        =   20
         Top             =   4155
         Width           =   650
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "卫生材料外购入库单"
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
         TabIndex        =   17
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
         TabIndex        =   0
         Top             =   720
         Width           =   630
      End
      Begin VB.Label Lbl填制人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  填制人"
         Height          =   180
         Left            =   120
         TabIndex        =   28
         Top             =   4485
         Width           =   720
      End
      Begin VB.Label Lbl填制日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "填制日期"
         Height          =   180
         Left            =   135
         TabIndex        =   27
         Top             =   4890
         Width           =   720
      End
      Begin VB.Label lbl审核人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  审核人"
         Height          =   180
         Left            =   8610
         TabIndex        =   26
         Top             =   4560
         Width           =   720
      End
      Begin VB.Label Lbl审核日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审核日期"
         Height          =   180
         Left            =   8610
         TabIndex        =   25
         Top             =   4965
         Width           =   720
      End
      Begin VB.Label LblProvider 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "供货单位(&G)"
         Height          =   180
         Left            =   7200
         TabIndex        =   2
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
            Picture         =   "frmPurchaseCard.frx":014A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":0364
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":057E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":0798
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":09B2
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":0BCC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":0DE6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":1000
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
            Picture         =   "frmPurchaseCard.frx":121A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":1434
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":164E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":1868
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":1A82
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":1C9C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":1EB6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":20D0
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   39
      Top             =   6630
      Width           =   11430
      _ExtentX        =   20161
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPurchaseCard.frx":22EA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13811
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmPurchaseCard.frx":2B7E
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmPurchaseCard.frx":3080
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
   Begin VB.Frame fraMoveNO 
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   1200
      TabIndex        =   53
      Top             =   6000
      Width           =   7605
      Begin VB.ComboBox cboType 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1245
         Style           =   2  'Dropdown List
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   15
         Width           =   885
      End
      Begin VB.CheckBox chk转入移库 
         Caption         =   "本张入库单移库到"
         Height          =   270
         Left            =   90
         TabIndex        =   7
         Top             =   30
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.ComboBox cboEnterStock 
         Enabled         =   0   'False
         Height          =   300
         ItemData        =   "frmPurchaseCard.frx":3582
         Left            =   2160
         List            =   "frmPurchaseCard.frx":358B
         TabIndex        =   9
         Text            =   "cboEnterStock"
         Top             =   15
         Visible         =   0   'False
         Width           =   2200
      End
      Begin VB.CommandButton cmdDraw 
         Caption         =   "…"
         Height          =   300
         Left            =   4575
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   15
         Width           =   300
      End
      Begin VB.TextBox txtDraw 
         Height          =   300
         Left            =   2190
         TabIndex        =   10
         Top             =   15
         Width           =   2415
      End
      Begin VB.CommandButton cmdDrawPerson 
         Caption         =   "…"
         Height          =   300
         Left            =   7230
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   15
         Width           =   300
      End
      Begin VB.TextBox txtDrawPerson 
         Height          =   300
         Left            =   5790
         TabIndex        =   13
         Top             =   15
         Width           =   1425
      End
      Begin VB.Label lbl领用人 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "领用人(&L)"
         Height          =   180
         Left            =   4920
         TabIndex        =   12
         Top             =   75
         Width           =   825
      End
   End
   Begin VB.Label lblCopy 
      AutoSize        =   -1  'True
      Caption         =   "(最大9999)"
      Enabled         =   0   'False
      Height          =   180
      Left            =   7920
      TabIndex        =   61
      Top             =   5850
      Width           =   900
   End
   Begin VB.Label lblCode 
      Caption         =   "材料"
      Height          =   255
      Left            =   3945
      TabIndex        =   33
      Top             =   5865
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "检索类型"
      Visible         =   0   'False
      Begin VB.Menu mnuSearch01 
         Caption         =   "病人ID"
      End
      Begin VB.Menu mnuSearch02 
         Caption         =   "病人姓名"
      End
      Begin VB.Menu mnuSearch03 
         Caption         =   "住院号"
      End
      Begin VB.Menu mnuSearch04 
         Caption         =   "门诊号"
      End
      Begin VB.Menu mnuSearch05 
         Caption         =   "床号"
      End
   End
End
Attribute VB_Name = "frmPurchaseCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'----------------------------------------------------------------------------------------------------------
'刘兴宏:增加小数位数的格式串
'修改:2007/03/06
Private mFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------
Private mrs环节控制 As ADODB.Recordset
Private mlng供货单位ID As Long              '供药单位ID
Private mint编辑状态 As Integer             '1.新增；2、修改；3、验收；4、查看；5、修改发票；6、冲销；
                                            '7、财务审核（冲销、产生新单据并审核；已付款的单据不允许财务审核，同样，财务审核后的单据不允许冲销）;
                                            '8、卫材库退货,9-核查,10-修改注册证号
Private mstr单据号 As String                '具体的单据号;
Private mint记录状态 As Integer             '1:正常记录;2-冲销记录;3-已经冲销的原记录
Private mblnSuccess As Boolean              '只要有一张成功，即为True，否则为False
Private mblnSave As Boolean                 '是否存盘和审核   TURE：成功。
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEdit As Boolean                 '是否可以修改
Private mblnChange As Boolean               '是否进行过编辑
Private mintParallelRecord As Integer       '对于新增后单据并行执行的处理： 1、代表正常情况；2、已经删除的记录；3、已经审核的记录
Private mstrPrivs As String                 '权限
Private mintBatchNoLen As Integer           '数据库中批号定义长度
Private mint发票号Len As Integer            '数据库中的发票号长度

Private mbln修改批发价 As Boolean               '允许修改批发价
Private mbln加价率 As Boolean                   '时价卫材是否必须输入加价率
Private mdbl加价率 As Double
Private mbln单据增加    As Boolean              '进入时单据号累加1
Private mintUnit  As Integer                    '显示单位:0-散装单位,1-包装单位
Private mbln退货 As Boolean
Private mblnFirst As Boolean
Private mbln库房  As Boolean                    '该库房是否为卫材库!
Private mstr审核日期 As String                  '财务审核时间
Private mstr验收结论 As String                  '用来记录默认验收结论

Public mrsReturn As Recordset
Private mbln不强制控制指导价格 As Boolean       '控制是否不强制采用指导价格
Private mbln分段加成率 As Boolean               '以分段加成率为依据
Private mbln时价购前销售 As Boolean             '外购入时，时价按扣前加成算售价
Private mbln时价卫材直接确定售价 As Boolean     '外购入库时,时价卫材直接确定售价，直接确定售价的意思是可以手动输入
Private mblnUpdate As Boolean                   '是否按新的售价更新单据，主要是针对审核时定价格不致的情况
Private mblnCheckPrice As Boolean
Private mbln非中标单位入库 As Boolean           '如果为true,招标卫生材料可以在非中标单位中入库
Private mCllBillData As Collection              '已使用的数量集合,目前主要是退货单的修改,以材料ID & "-" & 批次为主键
Private mbln需要核查 As Boolean
Private mbln供应商校验 As Boolean
Private mintCheckType As Integer                '资质校验方式：0-不检查；1－提醒；2－禁止
Private mintProduceDate As Integer              '生产日期注册证效期检查，1-检查，0-不检查

Private mrsCostlyInfo As ADODB.Recordset        '高值材料
Private mlngLastRow As Long
Private mbln时价卫材取上次售价 As Boolean        '外购入库时，时价卫材取上次售价 true-取上次售价 false-其他方式计算
Private mblnCostView As Boolean                 '查看成本价 true-允许查看 false-不允许查看
Private Const mstrCaption As String = "卫材外购入库单"
Private mintLastCol As Integer                  '用户的列设置中的最后可见列的列号
Private mbln高值卫材 As Boolean                 'true-当前行是高值卫材，false-不是高值卫材

Private mbln分批卫材批号产地控制 As Boolean  '是否检查分批卫材批号产地是否录入

'刘兴宏:2007/06/10:问题10813
Private mstrTime_Start As String                '进入单据编辑的单据时间 ,主要判断是否单据被他人更改过,如果编辑过,则不能进行审核
Private mstrTime_End As String
Private Const mlngModule = 1712

Private recSort As ADODB.Recordset          '按药品ID排序的专用记录集

'=========================================================================================
Private mCol行号 As Integer
Private mCol诊疗 As Integer
Private mCol序号 As Integer
Private mCol商品名 As Integer
Private mCol规格 As Integer
Private mCol原产地 As Integer
Private mCol原销期 As Integer
Private mCol比例系数 As Integer
Private mCol简码 As Integer
Private mCol产地 As Integer
Private mcol批准文号 As Integer
Private mCol单位 As Integer
Private mCol批号 As Integer
Private mcol生产日期 As Integer
Private mCol效期 As Integer
Private mCol数量 As Integer
Private mCol冲销数量 As Integer
Private mCol批次 As Integer
Private mCol指导批发价 As Integer
Private mCol采购价 As Integer
Private mCol扣率 As Integer
Private mcol加成率 As Integer
Private mCol结算价 As Integer
Private mCol结算金额 As Integer
Private mCol售价 As Integer
Private mCol售价金额 As Integer
Private mCol差价 As Integer
Private mcol零售价 As Integer
Private mcol零售单位 As Integer
Private mcol零售金额 As Integer
Private mcol零售差价 As Integer
Private mCol随货单号 As Integer
Private mCol验收结论 As Integer
Private mCol发票号 As Integer
Private mcol发票代码 As Integer
Private mCol发票日期 As Integer
Private mCol发票金额 As Integer
Private mcol一次性材料 As Integer
Private mcol灭菌效期 As Integer
Private mcol灭菌日期   As Integer
Private mcol灭菌失效期 As Integer
Private mcol注册证号 As Integer
Private mcol商品条码 As Integer
Private mcol条码管理 As Integer
Private mcol内部条码 As Integer
Private mcol费用ID As Integer
Private mcol注册证有效期 As Integer
Private Const mCols As Integer = 48
'=========================================================================================
Private Function CheckValuePrice(ByVal int单据 As Integer, ByVal strNo As String) As Boolean
    '检查高值卫材虚拟入库产生的入库单的价格，有价格变动时更新界面价格，金额
    '查找在入库单填制日期后是否存在同批次的调价记录，如果有调价记录，找最近的调价记录和当前入库单的价格进行比较
    '只检查时价卫材的售价和成本价
    '返回：true-检查通过,false-有价格变动
    Dim rsData As ADODB.Recordset
    Dim rsprice As ADODB.Recordset
    Dim lng材料ID As Long
    Dim lng批次 As Long
    Dim str填制日期 As String
    Dim dbl原价 As Double
    Dim dbl现售价 As Double
    Dim dbl现成本价 As Double
    Dim strAdjustList As String '需要变动的清单：材料id,批次,现售价(为0表示价格无变化),现成本价(为0表示价格无变化)
    Dim lngRow As Long
    Dim lngRows As Long
    Dim dbl数量 As Double
    Dim dbl成本金额 As Double
    Dim dbl零售金额 As Double
    Dim dbl差价 As Double
    Dim blnUpdate As Boolean
    
    gstrSQL = "Select '售价' As 类型, a.药品id As 材料id, Nvl(a.批次, 0) As 批次, a.零售价 As 原价, a.填制日期 " & vbNewLine & _
        " From 药品收发记录 A, 收费项目目录 C" & vbNewLine & _
        " Where a.单据 = [1] And a.No = [2] And c.Id = a.药品id And Nvl(c.是否变价, 0) = 1 And a.费用id > 0 And Exists" & vbNewLine & _
        " (Select 1" & vbNewLine & _
        "       From 药品收发记录 B" & vbNewLine & _
        "       Where a.药品id = b.药品id And a.批次 = b.批次 And b.单据 = 13 And b.审核日期 > a.填制日期 And b.摘要 = '卫材调价')" & vbNewLine & _
        " Union All" & vbNewLine & _
        " Select '成本价' As 类型, a.药品id As 材料id, Nvl(a.批次, 0) As 批次, a.成本价 As 原价, a.填制日期 " & vbNewLine & _
        " From 药品收发记录 A" & vbNewLine & _
        " Where a.单据 = [1] And a.No = [2] And a.费用id > 0 And Exists" & vbNewLine & _
        " (Select 1" & vbNewLine & _
        "       From 药品收发记录 B" & vbNewLine & _
        "       Where a.药品id = b.药品id And a.批次 = b.批次 And b.单据 = 18 And b.审核日期 > a.填制日期 And b.摘要 = '卫生材料成本价调价') "
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "CheckValuePrice", int单据, strNo)
        
    If rsData.RecordCount = 0 Then
        CheckValuePrice = True
        Exit Function
    End If
    
    '检查到有调价记录则比较价格，由于调价记录可能有多条，取最近一条价格来比较
    Do While Not rsData.EOF
        lng材料ID = rsData!材料ID
        lng批次 = rsData!批次
        str填制日期 = Format(rsData!填制日期, "yyyy-mm-dd hh:mm:ss")
        dbl原价 = rsData!原价
        
        dbl现售价 = 0
        dbl现成本价 = 0
        
        If rsData!类型 = "售价" Then
            gstrSQL = "Select 零售价 As 现价 " & _
                " From 药品收发记录 " & _
                " Where ID = (Select Max(ID) " & _
                " From 药品收发记录 B " & _
                " Where b.药品id = [1] And b.批次 = [2] And b.单据 = 13 And b.审核日期 > [3] And b.摘要 = '卫材调价') "
            Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, "CheckValuePrice", lng材料ID, lng批次, CDate(str填制日期))
            
            If rsprice.RecordCount > 0 Then
                If Round(rsprice!现价, 2) <> Round(dbl原价, 2) Then
                    dbl现售价 = rsprice!现价
                    blnUpdate = True
                End If
            End If
        End If
        
        If rsData!类型 = "成本价" Then
            gstrSQL = "Select 单量 As 现价 " & _
                " From 药品收发记录 " & _
                " Where ID = (Select Max(ID) " & _
                " From 药品收发记录 B " & _
                " Where b.药品id = [1] And b.批次 = [2] And b.单据 = 18 And b.审核日期 > [3] And b.摘要 = '卫生材料成本价调价') "
            Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, "CheckValuePrice", lng材料ID, lng批次, CDate(str填制日期))
            
            If rsprice.RecordCount > 0 Then
                If Round(rsprice!现价, 2) <> Round(dbl原价, 2) Then
                    dbl现成本价 = rsprice!现价
                    blnUpdate = True
                End If
            End If
        End If
        
        '以当前最新价格最新单据相关数据（单价、零售金额、差价）
        lngRows = mshBill.Rows - 1
        For lngRow = 1 To lngRows
            If lng材料ID = Val(mshBill.TextMatrix(lngRow, 0)) And (dbl现售价 <> 0 Or dbl现成本价 <> 0) Then
                dbl数量 = Val(mshBill.TextMatrix(lngRow, mCol数量))
                If dbl现售价 <> 0 Then
                    dbl现售价 = Val(Format(dbl现售价 * Val(mshBill.TextMatrix(lngRow, mCol比例系数)), mFMT.FM_零售价))
                    dbl零售金额 = dbl现售价 * dbl数量
                Else
                    dbl现售价 = Val(mshBill.TextMatrix(lngRow, mCol售价))
                    dbl零售金额 = Val(mshBill.TextMatrix(lngRow, mCol售价金额))
                End If
                
                If dbl现成本价 <> 0 Then
                    dbl现成本价 = Val(Format(dbl现成本价 * Val(mshBill.TextMatrix(lngRow, mCol比例系数)), mFMT.FM_成本价))
                    dbl成本金额 = dbl现成本价 * dbl数量
                Else
                    dbl现成本价 = Val(mshBill.TextMatrix(lngRow, mCol结算价))
                    dbl成本金额 = Val(mshBill.TextMatrix(lngRow, mCol结算金额))
                End If
                
                dbl差价 = dbl零售金额 - dbl成本金额
                
                mshBill.TextMatrix(lngRow, mCol结算价) = Format(dbl现成本价, mFMT.FM_成本价)
                mshBill.TextMatrix(lngRow, mCol结算金额) = Format(dbl成本金额, mFMT.FM_金额)
                mshBill.TextMatrix(lngRow, mCol售价) = Format(dbl现售价, mFMT.FM_零售价)
                mshBill.TextMatrix(lngRow, mCol售价金额) = Format(dbl零售金额, mFMT.FM_金额)
                mshBill.TextMatrix(lngRow, mCol差价) = Format(dbl差价, mFMT.FM_金额)
                
                ''刘兴宏:零售价处理
                Call 计算零售价及零售差价(lngRow)
            End If
        Next
        
        rsData.MoveNext
    Loop
    
    CheckValuePrice = Not blnUpdate
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
'检查数据依赖性
Private Function GetDepend() As Boolean
    Dim rstemp As New Recordset
    
    On Error GoTo ErrHandle
    GetDepend = False
    
    gstrSQL = "" & _
        "   SELECT B.Id " & _
        "   FROM 药品单据性质 A, 药品入出类别 B " & _
        "   Where A.类别id = B.ID " & _
        "           AND A.单据 = 30 and rownum=1 "
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否存卫生材料入出类别")
    If rstemp.EOF Then
        ShowMsgBox "没有设置卫生材料外购入库的入出类别，请检查卫生材料入出分类！"
        rstemp.Close
        Exit Function
    End If
    rstemp.Close
   
    gstrSQL = "" & _
        "   Select id " & _
        "   From 供应商 " & _
        "   where (撤档时间=to_date('3000-01-01','yyyy-mm-dd') or 撤档时间 is null)  " & _
        "           And substr(类型,5,1)=1 and Nvl(末级,0)=1 and (站点=[1] or 站点 is null) and rownum=1 "
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否存供应商", gstrNodeNo)
    If rstemp.EOF Then
        ShowMsgBox "没有设置卫生材料供应单位，请在供应商管理中设置！"
        rstemp.Close
        Exit Function
    End If
    rstemp.Close
    GetDepend = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

Public Sub ShowCard(FrmMain As Form, ByVal str单据号 As String, ByVal int编辑状态 As Integer, _
    Optional int记录状态 As Integer = 1, Optional strPrivs As String, Optional blnSuccess As Boolean = False)
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:单据入口
    '--入参数:frmnMain-调用的窗体
    '--       str单据号-单据号;int编辑状态;int记录状态;strPrivs-权限串
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim strReg As String
    
    Set mfrmMain = FrmMain
    
    mblnSave = False
    mblnSuccess = False
    
    mstr单据号 = str单据号
    mint编辑状态 = int编辑状态
    
    mint记录状态 = int记录状态
        
    mblnSuccess = blnSuccess
    mblnChange = False
    mintParallelRecord = 1
    mstrPrivs = strPrivs
     
    mbln修改批发价 = IIf(Val(zlDatabase.GetPara("修改采购限价", glngSys, mlngModule, "0")) = 1, 1, 0) = 1
    mbln供应商校验 = (Val(zlDatabase.GetPara("校验供应商资质", glngSys, mlngModule, "0")) = 1)
   
    
    Call GetRegInFor(g私有模块, "卫材外购入库管理", "单据号累加", strReg)
    mbln单据增加 = IIf(strReg = "", True, Val(strReg) = 1)
    
    If Not GetDepend Then Exit Sub
    
    cmdCopy.Visible = (mint编辑状态 = 1 Or mint编辑状态 = 2)
    txtCopy.Visible = (mint编辑状态 = 1 Or mint编辑状态 = 2)
    lblCopy.Visible = (mint编辑状态 = 1 Or mint编辑状态 = 2)
    
    If mint编辑状态 = 1 Or mint编辑状态 = 8 Then '新增或退货
        mblnEdit = True

        txtNO.Locked = True
        txtNO.TabStop = True
        txtNO = mstr单据号
        txtNO.Tag = txtNO
    ElseIf mint编辑状态 = 2 Or mint编辑状态 = 7 Then    '修改或财务审核
        mblnEdit = True
        
        If mint编辑状态 = 2 Then
            txtNO.Locked = True
            txtNO.TabStop = True
        End If
    ElseIf mint编辑状态 = 3 Then            '审核
        mblnEdit = False
        CmdSave.Caption = "审核(&V)"
        chk转入移库.Visible = True
    ElseIf mint编辑状态 = 4 Then            '查阅
        mblnEdit = False
        CmdSave.Caption = "打印(&P)"
        If InStr(mstrPrivs, "单据打印") = 0 Then
            CmdSave.Visible = False
        Else
            CmdSave.Visible = True
        End If
        vsfCostlyInfo.Editable = flexEDNone
    ElseIf mint编辑状态 = 5 Then        '修改发票
        mblnEdit = False
    ElseIf mint编辑状态 = 6 Then        '冲销
        mblnEdit = False
        CmdSave.Caption = "冲销(&O)"
        cmdAllSel.Visible = True
        cmdAllCls.Visible = True
    ElseIf mint编辑状态 = 9 Then
        '核查功能
        mblnEdit = False
    End If
    fraMoveNO.Visible = mint编辑状态 = 3    '审核

    LblTitle.Caption = GetUnitName & IIf(mint编辑状态 = 8, "卫生材料退货单", LblTitle.Caption)
    Me.Show vbModal, FrmMain
    blnSuccess = mblnSuccess
    str单据号 = mstr单据号
    
End Sub

Private Sub cboEnterStock_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cboEnterStock.ListCount = 0 Then Exit Sub
    If cboEnterStock.ListIndex >= 0 Then
        If Val(cboEnterStock.Tag) = cboEnterStock.ItemData(cboEnterStock.ListIndex) Then
            Exit Sub
        End If
    End If
    
    Dim blnOptionerPrivs As Boolean
    
    blnOptionerPrivs = Not zlStr.IsHavePrivs(mstrPrivs, "所有库房")
    If Select部门选择器(Me, cboEnterStock, Trim(cboEnterStock.Text), "V,W,K,12", blnOptionerPrivs) = False Then
        Exit Sub
    End If
    If cboEnterStock.ListIndex >= 0 Then
        cboEnterStock.Tag = cboEnterStock.ItemData(cboEnterStock.ListIndex)
    End If
End Sub

Private Sub cboStock_Change()
    mblnChange = True
End Sub

Private Sub cboStock_Click()
    Call 当前仅为库房
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
            For i = 1 To mshBill.Rows - 1
                If mshBill.TextMatrix(i, 0) <> "" Then
                    Exit For
                End If
            Next
            If i <> mshBill.Rows Then
                If MsgBox("如果改变库房，有可能要改变相应卫生材料的单位，且要清除现有单据内容，你是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    '处理卫生材料单位改变
                    mintcboIndex = .ListIndex
                    mshBill.ClearBill
                Else
                    .ListIndex = mintcboIndex
                End If
            Else
                mintcboIndex = .ListIndex
            End If
        End If
    End With
End Sub

Private Sub cboType_Change()
    mblnChange = True
End Sub

Private Sub cboType_Click()
    Dim bln移库 As Boolean
        
    bln移库 = cboType.ItemData(cboType.ListIndex) = 0

    cboEnterStock.Visible = bln移库
    txtDraw.Visible = Not bln移库
    cmdDraw.Visible = Not bln移库
    txtDrawPerson.Visible = Not bln移库
    cmdDrawPerson.Visible = Not bln移库
    lbl领用人.Visible = Not bln移库
    lbl领用人.Enabled = lbl领用人.Visible
    
End Sub

Private Sub cboType_GotFocus()
    OS.OpenIme False
End Sub

Private Sub cboType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub chk转入移库_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub cmdAllCls_Click()
    Dim intRow As Integer
    With mshBill
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                .TextMatrix(intRow, mCol冲销数量) = Format(0, mFMT.FM_数量)
                .TextMatrix(intRow, mCol结算金额) = Format(0, mFMT.FM_金额)
                .TextMatrix(intRow, mCol售价金额) = Format(0, mFMT.FM_金额)
                .TextMatrix(intRow, mCol差价) = Format(0, mFMT.FM_金额)
                '刘兴宏:零售价处理
                .TextMatrix(intRow, mcol零售金额) = Format(0, mFMT.FM_金额)
                .TextMatrix(intRow, mcol零售差价) = Format(0, mFMT.FM_金额)
                If Trim(.TextMatrix(intRow, mCol发票号)) <> "" Then
                    .TextMatrix(intRow, mCol发票金额) = Format(0, mFMT.FM_金额)
                End If
            End If
        Next
    End With
End Sub

Private Sub cmdAllSel_Click()
    Dim rstemp As New Recordset
    Dim intRow As Integer, dbl发票金额 As Double
    
    With mshBill
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" And .RowData(intRow) = 0 Then
                .TextMatrix(intRow, mCol冲销数量) = .TextMatrix(intRow, mCol数量)
                .TextMatrix(intRow, mCol结算金额) = Format(Val(.TextMatrix(intRow, mCol数量)) * Val(.TextMatrix(intRow, mCol结算价)), mFMT.FM_金额)
                .TextMatrix(intRow, mCol售价金额) = Format(Val(.TextMatrix(intRow, mCol数量)) * Val(.TextMatrix(intRow, mCol售价)), mFMT.FM_金额)
                .TextMatrix(intRow, mCol差价) = Format(Val(.TextMatrix(intRow, mCol售价金额)) - Val(.TextMatrix(intRow, mCol结算金额)), mFMT.FM_金额)
                
                '刘兴宏:零售价处理,主要是确定时价定价问题
                Call 计算零售价及零售差价(intRow, False)
                If Trim(.TextMatrix(intRow, mCol发票号)) <> "" Or Trim(.TextMatrix(intRow, mCol随货单号)) <> "" Then
                
                    dbl发票金额 = GetTotale发票金额(mstr单据号, Val(.TextMatrix(intRow, 0)), Val(.TextMatrix(.Row, mCol序号)))
                    If dbl发票金额 = 0 Then dbl发票金额 = Val(.TextMatrix(intRow, mCol结算金额))
                    .TextMatrix(intRow, mCol发票金额) = Format(dbl发票金额, mFMT.FM_金额)
                End If
            End If
        Next
    End With
End Sub
Public Function GetTotale发票金额(ByVal strNo As String, ByVal lng材料ID As Long, ByVal lng序号 As Long) As Double
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取发票金额
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-11-19 14:43:30
    '-----------------------------------------------------------------------------------------------------------
    Dim rstemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "" & _
        "   Select sum(nvl(Q.发票金额,0)) as 发票金额 " & _
        "   From 药品收发记录 x," & _
        "        ( Select B.ID, B.收发id,Sum(B.发票金额) as 发票金额, Max(B.发票号) As 发票号,Max(B.随货单号) as 随货单号, Max(B.发票日期) As 发票日期, Max(B.付款序号) As 付款序号 " & _
        "          From 药品收发记录 A,应付记录 B " & _
        "          Where A.ID = B.收发id And A.NO =[1] and A.药品ID=[2] And A.序号=[3]  And A.单据 = 15 And B.系统标识 = 5 And B.记录性质 In (0, -1) " & _
        "          Group By B.ID,B.收发id ) Q " & _
        "   WHERE x.id=q.收发id(+) AND X.单据=15" & _
        "         and X.NO=[1] and X.药品id=[2] and x.序号=[3]"
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, strNo, lng材料ID, lng序号)
    If rstemp.EOF Then
        GetTotale发票金额 = 0
    Else
        GetTotale发票金额 = IIf(mbln退货, -1, 1) * Val(zlStr.Nvl(rstemp!发票金额))
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

Private Sub cmdBulkCopy_Click()
    Dim i As Integer
    
    With mshBill
        '1、都有发票号则不做第2步
        For i = 1 To .Rows - 1
           If Trim(.TextMatrix(i, mCol发票号)) = "" Or .TextMatrix(i, 0) = "" Then Exit For
        Next
        
        If i = .Rows - 1 Then Exit Sub
        
        '2、发票代码或发票日期为空，则提示
        If Trim(.TextMatrix(.Row, mcol发票代码)) = "" Or .TextMatrix(.Row, mCol发票日期) = "" Then
            If MsgBox("发票代码或发票日期为空，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Else
            If MsgBox("是否将该行的发票信息批量复制到发票号为空的行？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        
        '3、复制
        For i = 1 To .Rows - 1
            If i <> .Row And Trim(.TextMatrix(i, mCol发票号)) = "" And .TextMatrix(i, 0) <> "" Then    '不是编辑行且发票号为空的批量修改
                
                .TextMatrix(i, mCol发票号) = .TextMatrix(.Row, mCol发票号)
                .TextMatrix(i, mcol发票代码) = .TextMatrix(.Row, mcol发票代码)
                .TextMatrix(i, mCol发票日期) = .TextMatrix(.Row, mCol发票日期)
                .TextMatrix(i, mCol发票金额) = .TextMatrix(i, mCol结算金额)
                
            End If
        Next
    End With
End Sub

Private Sub cmdALLDel_Click()
    Dim i As Integer
    
    With mshBill
        '1、都无发票号则不做第2步
        For i = 1 To .Rows - 1
           If Trim(.TextMatrix(i, mCol发票号)) <> "" Or .TextMatrix(i, 0) = "" Then Exit For
        Next
        
        If i = .Rows - 1 Then Exit Sub
    
        If MsgBox("该操作将清除所有行的发票相关数据，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            For i = 1 To .Rows - 1
            
                If Trim(.TextMatrix(i, mCol发票号)) <> "" And .TextMatrix(i, 0) <> "" Then
                    .TextMatrix(i, mCol发票号) = ""
                    .TextMatrix(i, mcol发票代码) = ""
                    .TextMatrix(i, mCol发票日期) = ""
                    .TextMatrix(i, mCol发票金额) = ""
                End If
                
            Next
            
            cmdBulkCopy.Enabled = False
        End If
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCopy_Click()
    '按录入数量复制当前行数据到该行数据后面
    Dim lngCopyNum As Long
    Dim lngMoveRowStart As Long, lngMoveRowEnd As Long
    Dim i As Long
    Dim intCol As Integer
    Dim lngRow As Long
    Dim str名称 As String
    Dim str单位 As String
    
    With mshBill
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        
        If mbln高值卫材 = False Then Exit Sub
        
        lngCopyNum = Val(Trim(txtCopy.Text))
        If lngCopyNum = 0 Then
            MsgBox "请录入复制的数量！", vbInformation, gstrSysName
            txtCopy.SetFocus
            Exit Sub
        End If
        
        str名称 = .TextMatrix(.Row, mCol诊疗)
        str单位 = .TextMatrix(.Row, mCol单位)
        
        '提醒
        If MsgBox("是否复制新增名称为“" & str名称 & "”的卫材共计" & lngCopyNum & "" & str单位 & " ?", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                
        lngRow = .Row
        
        '记录当前行后面数据的起始行号
        If .Row = .Rows - 1 Then
            '当前行是最后一行时
            lngMoveRowStart = 0
            lngMoveRowEnd = 0
        Else
            '当前行不是最后一行时
            lngMoveRowStart = .Row + 1
            lngMoveRowEnd = .Rows - 1
        End If
        
        '新增行数
        .Rows = .Rows + lngCopyNum
                
        '把当前行后面的数据依次往后移动
        If lngMoveRowStart <> 0 Then
            For i = lngMoveRowEnd To lngMoveRowStart Step -1
                For intCol = 0 To .Cols - 1
                    If mCol行号 = intCol Then
                        .TextMatrix(i + lngCopyNum, intCol) = i + lngCopyNum
                    Else
                        .TextMatrix(i + lngCopyNum, intCol) = .TextMatrix(i, intCol)
                    End If
                Next
            Next
        End If
        
        '复制当前行
        For i = 1 To lngCopyNum
            For intCol = 0 To .Cols - 1
                If mCol行号 = intCol Then
                    .TextMatrix(i + lngRow, intCol) = i + lngRow
                Else
                    .TextMatrix(i + lngRow, intCol) = .TextMatrix(lngRow, intCol)
                End If
            Next
        Next
    End With
End Sub

Private Sub cmdDrawPerson_Click()
    If SelectItem(txtDrawPerson, "", True) = False Then Exit Sub
End Sub

Private Sub cmdExtractData_Click()
    If Val(txtProvider.Tag) <= 0 Then
        txtProvider.SetFocus
        MsgBox "请录入供商单位信息！", vbInformation, gstrSysName
        Exit Sub
    End If
    With frmPurchaseCardExtract
        .EntryPort cboStock.ItemData(cboStock.ListIndex) & ";" & cboStock.Text, txtProvider.Tag
        .Show vbModal, Me
    End With
    With mshBill
        .Row = 1
        .SetFocus
    End With
End Sub

'查找
Private Sub cmdFind_Click()
    If lblCode.Visible = False Then
        lblCode.Visible = True
        txtCode.Visible = True
        txtCode.SetFocus
    Else
        FindRownew mshBill, mCol诊疗, txtCode.Text, True
        lblCode.Visible = False
        txtCode.Visible = False
    End If
    
    If mint编辑状态 = 5 Then '修改发票信息该按钮才可用
        With cmdALLDel
            cmdALLDel.Visible = True
            .Top = cmdFind.Top
            If txtCode.Visible Then
                .Left = txtCode.Left + txtCode.Width + 100
            Else
                .Left = cmdFind.Left + cmdFind.Width + 100
            End If
        End With
    End If
    
    With cmdCopy
        .Left = IIf(txtCode.Visible, txtCode.Left + txtCode.Width, cmdFind.Left + cmdFind.Width) + 100
        .Top = cmdFind.Top
    End With
    
    With txtCopy
        .Left = cmdCopy.Left + cmdCopy.Width + 50
        .Top = txtCode.Top
    End With
    
    With lblCopy
        .Left = txtCopy.Left + txtCopy.Width + 25
        .Top = txtCopy.Top + 45
    End With
        
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Sub CmdNO_Click()

    Dim mdbl加价率 As Double
    Dim dbl成本价 As Double
    
    With mshBill
        mdbl加价率 = Val(Txt加价率.Tag)
        If mbln时价购前销售 Then
            dbl成本价 = Val(.TextMatrix(.Row, mCol采购价))
        Else
            dbl成本价 = Val(.TextMatrix(.Row, mCol结算价))
        End If
        
        '重新计算零售价、差价
        If mint编辑状态 = 8 And Val(.TextMatrix(.Row, mCol售价)) <> 0 Then
        Else
            .TextMatrix(.Row, mCol售价) = Format(校正零售价(dbl成本价 * (1 + (mdbl加价率 / 100)) + _
             时价材料零售价(Val(.TextMatrix(.Row, 0)), dbl成本价, mdbl加价率 / 100)), mFMT.FM_零售价)
        End If
        .TextMatrix(.Row, mCol售价金额) = Format(Val(.TextMatrix(.Row, mCol售价)) * Val(.TextMatrix(.Row, mCol数量)), mFMT.FM_金额)
        .TextMatrix(.Row, mCol差价) = Format(IIf(.TextMatrix(.Row, mCol售价金额) = "", 0, .TextMatrix(.Row, mCol售价金额)) - IIf(.TextMatrix(.Row, mCol结算金额) = "", 0, .TextMatrix(.Row, mCol结算金额)), mFMT.FM_金额)
        
        '刘兴宏:零售价处理
        Call 计算零售价及零售差价(.Row, True)
    End With
    PicInput.Visible = False
End Sub

Private Sub CmdNO_LostFocus()
    Call PicInput_LostFocus
End Sub

Private Sub CmdYes_Click()
    If Val(Txt加价率) > 9900 Or Val(Txt加价率) < 0 Then
        MsgBox "请输入合法的加成率！（0-9900）", vbInformation, gstrSysName
        Txt加价率.SetFocus
        Exit Sub
    End If
    Dim dbl成本价 As Double
    With mshBill
        If mbln时价购前销售 Then
            dbl成本价 = Val(.TextMatrix(.Row, mCol采购价))
        Else
            dbl成本价 = Val(.TextMatrix(.Row, mCol结算价))
        End If
        '重新计算零售价、差价
        If mint编辑状态 = 8 And Val(.TextMatrix(.Row, mCol售价)) <> 0 Then
        Else
            .TextMatrix(.Row, mCol售价) = Format(校正零售价(dbl成本价 * (1 + (Val(Txt加价率) / 100)) + _
            时价材料零售价(Val(.TextMatrix(.Row, 0)), dbl成本价, Val(Txt加价率) / 100)), mFMT.FM_零售价)
        End If
         
        .TextMatrix(.Row, mcol加成率) = zlStr.FormatEx(Val(Txt加价率), 2) & "%"
        .TextMatrix(.Row, mCol售价金额) = Format(Val(.TextMatrix(.Row, mCol售价)) * Val(.TextMatrix(.Row, mCol数量)), mFMT.FM_金额)
        .TextMatrix(.Row, mCol差价) = Format(IIf(.TextMatrix(.Row, mCol售价金额) = "", 0, .TextMatrix(.Row, mCol售价金额)) - IIf(.TextMatrix(.Row, mCol结算金额) = "", 0, .TextMatrix(.Row, mCol结算金额)), mFMT.FM_金额)
        
        '刘兴宏:零售价处理
        Call 计算零售价及零售差价(.Row, True)
    End With
    
    PicInput.Visible = False
    mshBill.SetFocus
End Sub

Private Sub CmdYes_LostFocus()
    Call PicInput_LostFocus
End Sub

Private Sub Form_Activate()
'    mblnChange = False
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    Select Case mintParallelRecord
        Case 1
            '正常
        Case 2
            '单据已被删除
            If mint编辑状态 = 5 Then
               MsgBox "该单据已被全部冲销，不能修改发票信息，请检查！", vbOKOnly, gstrSysName
            ElseIf mint编辑状态 = 6 Then
                MsgBox "该单据已没有可以冲销的卫生材料，请检查！", vbOKOnly, gstrSysName
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
        Case 6  '冲销
            MsgBox "该单据已付过款，不能进行冲销！", vbOKOnly, gstrSysName
            Unload Me
            Exit Sub
    End Select
    SetEdit
    If IsCtrlSetFocus(txtProvider) Then
        zlControl.ControlSetFocus txtProvider
    Else
        If mint编辑状态 = 3 And IsCtrlSetFocus(chk转入移库) Then
             zlControl.ControlSetFocus chk转入移库
        Else
            zlControl.ControlSetFocus mshBill
        End If
    End If
    '初始化简码方式
    If (mint编辑状态 = 1 Or mint编辑状态 = 2) And gbytSimpleCodeTrans = 1 Then
        stbThis.Panels("PY").Visible = True
        stbThis.Panels("WB").Visible = True
        gSystem_Para.int简码方式 = Val(zlDatabase.GetPara("简码方式", , , 0))    '默认拼音简码
        Logogram stbThis, gSystem_Para.int简码方式
    Else
        stbThis.Panels("PY").Visible = False
        stbThis.Panels("WB").Visible = False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sngLeft As Single, sngTop As Single
    
    If KeyCode = 70 Or KeyCode = 102 Then
        If Shift = vbCtrlMask Then   'Ctrl+F
            cmdFind_Click
        End If
    ElseIf KeyCode = vbKeyF3 Then
        FindRownew mshBill, mCol诊疗, txtCode.Text, False
    ElseIf KeyCode = vbKeyF4 Then
        '如果系统参数为真，则提示用户输入加价率
        If mbln加价率 And (mint编辑状态 = 1 Or mint编辑状态 = 2) Then
            If PicInput.Visible Then PicInput.SetFocus: Exit Sub
            If mshBill.TextMatrix(mshBill.Row, mCol诊疗) = "" Then Exit Sub
            '存储格式值:最大效期||指导差价率||是否变价||在用分批||库房分批
            If Split(mshBill.TextMatrix(mshBill.Row, mCol原销期), "||")(2) <> 1 Then Exit Sub
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
            Txt加价率 = "15.0000"
            With mshBill
                If Val(.TextMatrix(.Row, mCol售价)) <> 0 And Val(.TextMatrix(.Row, mCol结算价)) <> 0 Then
                    Txt加价率 = Format((Val(.TextMatrix(.Row, mCol售价)) / Val(.TextMatrix(.Row, mCol结算价)) - 1) * 100, "####0.0000000;-####0.0000000;0;0")
                End If
            End With
            Txt加价率.Tag = Txt加价率
            Txt加价率.SetFocus
        End If
    ElseIf KeyCode = vbKeyF7 Then
        If stbThis.Panels("PY").Bevel = sbrRaised Then
            Logogram stbThis, 0
        Else
            Logogram stbThis, 1
        End If
    End If
End Sub

Private Sub cmdProvider_Click()
    Dim rstemp As New Recordset
    Dim blnCancel As Boolean
    Dim vRect As RECT
    
    vRect = zlControl.GetControlRect(txtProvider.hwnd)
    
    gstrSQL = "" & _
        "   Select id,上级ID,编码,简码,名称,末级 " & _
        "   From 供应商 " & _
        "   Where  (To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01' or 撤档时间 is null) " & _
        "       And (substr(类型,5,1)=1 And (站点=[1] or 站点 is null)  Or Nvl(末级,0)=0) " & _
        "   Start with 上级ID is null connect by prior ID =上级ID " & _
        "   Order by level,ID"
    
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
    
    Set rstemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 2, "供应商选择", True, "", "请选择符合卫生材料的供应商", True, True, False, vRect.Left - 15, vRect.Top, txtProvider.Height, blnCancel, False, False, gstrNodeNo)
        
    If rstemp Is Nothing Or blnCancel Then Exit Sub
    If rstemp.State <> 1 Then Exit Sub
    
    With rstemp
        Me.txtProvider = "[" & zlStr.Nvl(!编码) & "] " & zlStr.Nvl(!名称)
        Me.txtProvider.Tag = zlStr.Nvl(!Id)
    End With
    If mshBill.Col = 1 Then mshBill.Col = mCol诊疗
    mshBill.SetFocus
    
    If CheckQualifications(mlngModule, 2, Val(txtProvider.Tag)) = False Then
        txtProvider.Text = ""
        txtProvider.Tag = "0"
        Exit Sub
    End If
    
    If Val(txtProvider.Tag) <> mlng供货单位ID And (mint编辑状态 = 8 Or mbln退货) Then     '退货
        mlng供货单位ID = Val(txtProvider.Tag)
        mshBill.ClearBill
        mshBill.TextMatrix(1, mCol行号) = "1"
    End If
End Sub

'打印单据
Private Sub printbill()
    Dim strUnit As String
    Dim strNo As String
    strNo = txtNO.Tag
    
    FrmBillPrint.ShowMe Me, glngSys, "zl1_bill_1712", _
        mint记录状态, mintUnit, "1712", IIf(mint编辑状态 = 8 Or mbln退货, "卫材退货单", "卫材外购入库单"), strNo
End Sub

Private Function SaveNewCard(ByVal strNo As String) As Boolean
    '功能：财务审核产生新单据
    '参数strNO ：财务审核新单据no
    Dim chrNo As Variant
    Dim lng序号 As Long
    Dim lngStockID As Long
    Dim lng供货单位id As Long
    Dim lng材料ID As Long
    Dim str批号 As String
    Dim str产地 As String
    Dim str效期 As String
    Dim dbl实际数量 As Double
    Dim dbl成本价 As Double
    Dim dbl成本金额 As Double
    Dim dbl扣率 As Double
    Dim dbl零售价 As Double
    Dim dbl零售金额 As Double
    Dim dbl差价 As Double
    Dim str零售差价 As String '收发记录表中用法字段保存的是外购入库差价差，由于用法字段类型是字符型因此如果采用double类型会出现 -.00x现象
    Dim str摘要 As String
    Dim str填制人 As String
    Dim str填制日期 As String
    Dim str审核人 As String
    Dim datAssessDate As String
    Dim str发票号 As String
    Dim str发票代码 As String
    Dim str发票日期 As String
    Dim str灭菌日期 As String
    Dim str灭菌失效期 As String
    Dim dbl发票金额 As Double
    Dim str生产日期  As String
    Dim str核查人 As String
    Dim str核查日期 As String
    Dim str注册证号 As String
    Dim intUnit As Integer
    Dim strUnit As String
    Dim str指导批发价 As String
    Dim str随货单号 As String
    Dim str验收结论 As String
    Dim str商品条码 As String
    Dim str内部条码 As String
    Dim str批准文号 As String
    Dim lng费用ID As Long
    Dim intRow As Integer
    Dim i As Integer
    Dim arrSQL As Variant
    Dim n As Long
    
    SaveNewCard = False
    arrSQL = Array()
    With mshBill

        chrNo = Trim(txtNO)
        lngStockID = cboStock.ItemData(cboStock.ListIndex)
        lng供货单位id = txtProvider.Tag
        str摘要 = Trim(txt摘要.Text)

        str填制人 = Txt填制人
        str填制日期 = Txt填制日期.Caption
        str核查人 = txt核查人
        str核查日期 = txt核查日期.Caption
        str审核人 = Txt审核人

        On Error GoTo ErrHandle

        '取该库房的单位，更新指导批发价时使用
        '按药品ID顺序更新数据
        recSort.Sort = "药品id,批次,序号"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!行号
'        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                lng材料ID = .TextMatrix(intRow, 0)
                str产地 = .TextMatrix(intRow, mCol产地)
                str批号 = .TextMatrix(intRow, mCol批号)
                str批准文号 = .TextMatrix(intRow, mcol批准文号)
                str效期 = IIf(.TextMatrix(intRow, mCol效期) = "", "", .TextMatrix(intRow, mCol效期))
                dbl实际数量 = GetFormat(.TextMatrix(intRow, mCol数量) * .TextMatrix(intRow, mCol比例系数), g_小数位数.obj_最大小数.数量小数)
                dbl扣率 = Val(.TextMatrix(intRow, mCol扣率))
                dbl成本价 = GetFormat(Val(.TextMatrix(intRow, mCol结算价)) / .TextMatrix(intRow, mCol比例系数), g_小数位数.obj_最大小数.成本价小数)
                dbl成本金额 = GetFormat(Val(.TextMatrix(intRow, mCol结算金额)), g_小数位数.obj_最大小数.金额小数)

                'dbl零售价 = Round(Val(.TextMatrix(intRow, mCol售价)) / .TextMatrix(intRow, mCol比例系数), g_小数位数.obj_散装小数.零售价小数)
                'dbl零售金额 = Round(Val(.TextMatrix(intRow, mCol售价金额)), g_小数位数.obj_散装小数.金额小数)
                '数据库中的:差价 = 零售金额 - 结算金额
                '数据库中的:用法 = 零售金额-售价金额或零售差价-差价(库房单位的差价)

                dbl零售价 = GetFormat(Val(.TextMatrix(intRow, mcol零售价)), g_小数位数.obj_最大小数.零售价小数)
                dbl零售金额 = GetFormat(Val(.TextMatrix(intRow, mcol零售金额)), g_小数位数.obj_最大小数.零售价小数)
                dbl差价 = GetFormat(Val(.TextMatrix(intRow, mcol零售差价)), g_小数位数.obj_最大小数.零售价小数)
                str零售差价 = GetFormat(Val(.TextMatrix(intRow, mcol零售差价)) - Val(.TextMatrix(intRow, mCol差价)), g_小数位数.obj_最大小数.零售价小数)
                'dbl差价 = Round(Val(.TextMatrix(intRow, mCol差价)), g_小数位数.obj_散装小数.金额小数)
                lng序号 = .TextMatrix(intRow, mCol序号)

                str随货单号 = Trim(.TextMatrix(intRow, mCol随货单号))
                str验收结论 = Trim(.TextMatrix(intRow, mCol验收结论))
                str发票号 = Trim(.TextMatrix(intRow, mCol发票号))
                str发票代码 = Trim(.TextMatrix(intRow, mcol发票代码))
                str发票日期 = Trim(IIf(.TextMatrix(intRow, mCol发票日期) = "", "", .TextMatrix(intRow, mCol发票日期)))
                dbl发票金额 = Round(Val(.TextMatrix(intRow, mCol发票金额)), g_小数位数.obj_散装小数.金额小数)

                str灭菌日期 = Trim(IIf(.TextMatrix(intRow, mcol灭菌日期) = "", "", .TextMatrix(intRow, mcol灭菌日期)))
                str灭菌失效期 = Trim(IIf(.TextMatrix(intRow, mcol灭菌失效期) = "", "", .TextMatrix(intRow, mcol灭菌失效期)))
                str生产日期 = Trim(IIf(.TextMatrix(intRow, mcol生产日期) = "", "", .TextMatrix(intRow, mcol生产日期)))
                str注册证号 = Trim(.TextMatrix(intRow, mcol注册证号))

                str内部条码 = Trim(.TextMatrix(intRow, mcol内部条码))
                '财务审核新单据费用id等于2
                lng费用ID = 2

                str商品条码 = Trim(.TextMatrix(intRow, mcol商品条码))

                ' Zl_材料外购_Insert
                gstrSQL = "zl_材料外购_INSERT("
                '  No_In         In 药品收发记录.NO%Type,
                gstrSQL = gstrSQL & "'" & strNo & "',"
                '  序号_In       In 药品收发记录.序号%Type,
                gstrSQL = gstrSQL & "" & lng序号 & ","
                '  库房id_In     In 药品收发记录.库房id%Type,
                gstrSQL = gstrSQL & "" & lngStockID & ","
                '  供药单位id_In In 药品收发记录.供药单位id%Type,
                gstrSQL = gstrSQL & "" & lng供货单位id & ","
                '  材料id_In     In 药品收发记录.药品id%Type,
                gstrSQL = gstrSQL & "" & lng材料ID & ","
                '  产地_In       In 药品收发记录.产地%Type := Null,
                gstrSQL = gstrSQL & "'" & str产地 & "',"
                '  批号_In       In 药品收发记录.批号%Type := Null,
                gstrSQL = gstrSQL & "'" & str批号 & "',"
                '  生产日期_In   In 药品收发记录.生产日期%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str生产日期 = "", "Null", "to_date('" & Format(str生产日期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '  效期_In       In 药品收发记录.效期%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str效期 = "", "Null", "to_date('" & Format(str效期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '  灭菌日期_In   In 药品收发记录.灭菌日期%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str灭菌日期 = "", "Null", "to_date('" & Format(str灭菌日期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '  灭菌效期_In   In 药品收发记录.灭菌效期%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str灭菌失效期 = "", "Null", "to_date('" & Format(str灭菌失效期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '  实际数量_In   In 药品收发记录.实际数量%Type := Null,
                gstrSQL = gstrSQL & "" & dbl实际数量 & ","
                '  成本价_In     In 药品收发记录.成本价%Type := Null,
                gstrSQL = gstrSQL & "" & dbl成本价 & ","
                '  成本金额_In   In 药品收发记录.成本金额%Type := Null,
                gstrSQL = gstrSQL & "" & dbl成本金额 & ","
                '  扣率_In       In 药品收发记录.扣率%Type := Null,
                gstrSQL = gstrSQL & "" & dbl扣率 & ","
                '  零售价_In     In 药品收发记录.零售价%Type := Null,
                gstrSQL = gstrSQL & "" & dbl零售价 & ","
                '  零售金额_In   In 药品收发记录.零售金额%Type := Null,
                gstrSQL = gstrSQL & "" & dbl零售金额 & ","
                '  差价_In       In 药品收发记录.差价%Type := Null,
                gstrSQL = gstrSQL & "" & dbl差价 & ","
                '  零售差价_In   In 药品收发记录.差价%Type := Null,目前存放在用法字段
                gstrSQL = gstrSQL & "" & str零售差价 & ","
                '  摘要_In       In 药品收发记录.摘要%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str摘要 = "", "NULL", "'" & str摘要 & "'") & ","
                '   注册证号_In   In 药品收发记录.注册证号%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str注册证号 = "", "NULL", "'" & str注册证号 & "'") & ","
                '  填制人_In     In 药品收发记录.填制人%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str填制人 = "", "NULL", "'" & str填制人 & "'") & ","
                '  随货单号_In   In 应付记录.随货单号%Type := Null
                gstrSQL = gstrSQL & "" & IIf(str随货单号 = "", "NULL", "'" & str随货单号 & "'") & ","
                '  发票号_In     In 应付记录.发票号%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str发票号 = "", "NULL", "'" & str发票号 & "'") & ","
                '  发票日期_In   In 应付记录.发票日期%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str发票日期 = "", "Null", "to_date('" & Format(str发票日期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '  发票金额_In   In 应付记录.发票金额%Type := Null,
                gstrSQL = gstrSQL & "" & dbl发票金额 & ","
                '  填制日期_In   In 药品收发记录.填制日期%Type := Null,
                gstrSQL = gstrSQL & "to_date('" & str填制日期 & "','yyyy-mm-dd HH24:MI:SS'),"
                '  核查人_In     In 药品收发记录.配药人%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str核查人 = "", "NULL", "'" & str核查人 & "'") & ","
                '  核查日期_In   In 药品收发记录.配药日期%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str核查日期 = "", "Null", "to_date('" & str核查日期 & "','yyyy-mm-dd hh24:mi:ss')") & ","
                '  批次_In       In 药品收发记录.批次%Type := 0,
                gstrSQL = gstrSQL & "" & Val(.TextMatrix(intRow, mCol批次)) & ","
                '  退货_In       In Number := 1
                gstrSQL = gstrSQL & "" & IIf(mbln退货, -1, 1) & ","
                '  高值材料_In   In varchar2(250)
                gstrSQL = gstrSQL & "'" & GetCostlyInfoStr(intRow) & "'" & ","
                '  商品条码_In   In 药品收发记录.商品条码%Type :=Null
                gstrSQL = gstrSQL & "" & IIf(str商品条码 = "", "NULL", "'" & str商品条码 & "'") & ","
                '  内部条码
                gstrSQL = gstrSQL & IIf(str内部条码 = "", "Null", "'" & str内部条码 & "'") & ","
                '  费用ID
                gstrSQL = gstrSQL & lng费用ID & ","
                '  发票代码
                gstrSQL = gstrSQL & IIf(str发票代码 = "", "NULL", "'" & str发票代码 & "'")
                '  财务审核
                gstrSQL = gstrSQL & ",1,"
                '  批准文号
                gstrSQL = gstrSQL & IIf(str批准文号 = "", "NULL", "'" & str批准文号 & "'") & ","
                '  验收结论
                gstrSQL = gstrSQL & "'" & str验收结论 & "'"
                gstrSQL = gstrSQL & ")"

                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = gstrSQL
            End If
            
            recSort.MoveNext
        Next
        
        For i = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "SaveCard")
        Next

        mstr单据号 = chrNo
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveNewCard = True
    Exit Function

ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SaveVerifyCard(ByVal strNo As String) As Boolean
    '功能：财务审核时向财务审核记录表中插入数据
    '返回值:true-执行成功 false-执行失败
    Dim str审核日期 As String
    
    On Error GoTo ErrHand
    
    SaveVerifyCard = False
    
    gstrSQL = "Zl_药品财务审核_Insert("
    '库房id
    gstrSQL = gstrSQL & cboStock.ItemData(cboStock.ListIndex)
    '单据
    gstrSQL = gstrSQL & ",15"
    '冲销no
    gstrSQL = gstrSQL & ",'" & txtNO.Text & "'"
    'newNO
    gstrSQL = gstrSQL & ",'" & strNo & "'"
    '审核人
    gstrSQL = gstrSQL & ",'" & UserInfo.用户名 & "'"
    '审核日期
    gstrSQL = gstrSQL & ",to_date('" & Format(mstr审核日期, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS')"
    '备注
    If Trim(txt摘要.Text) = "" Then
        gstrSQL = gstrSQL & "," & "Null" & ")"
    Else
        gstrSQL = gstrSQL & ",'" & txt摘要.Text & "')"
    End If
    
    Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
    SaveVerifyCard = True
    Exit Function
    
ErrHand:
If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub CmdSave_Click()
    On Error GoTo ErrHand
    Dim strNewNO As String
    Dim strReg As String
    Dim blnSuccess As Boolean, blnTrans As Boolean
    
    '设置排序数据集
    Call SetSortRecord
    
    mstr审核日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
    If mint编辑状态 = 4 Then    '查看
        '打印
        printbill
        '退出
        Unload Me
        Exit Sub
    End If
    
    If mint编辑状态 = 7 Then
        ' '财务审核（冲销、产生新单据并审核；已付款的单据不允许财务审核，同样，财务审核后的单据不允许冲销）
        '先冲销，再新增单据并审核
        gcnOracle.BeginTrans
        
        blnTrans = True
        '产生新的no
        strNewNO = sys.GetNextNo(68, cboStock.ItemData(cboStock.ListIndex))
        blnSuccess = (strNewNO <> "")
        '产生新的未审核的财务审核单据
        If blnSuccess Then blnSuccess = SaveNewCard(strNewNO)
        '冲销原始单据
        If blnSuccess Then blnSuccess = SaveStrike
        '审核新产生的财务审核单据
        If blnSuccess Then blnSuccess = SaveCheck(strNewNO)
        '向财务审核记录表中插入数据
        If blnSuccess Then blnSuccess = SaveVerifyCard(strNewNO)
        
        If blnSuccess Then
            gcnOracle.CommitTrans
        Else
            gcnOracle.RollbackTrans
            Exit Sub
        End If
        blnTrans = False
        Unload Me
        Exit Sub
    End If
    
    '刘兴宏:2007/05/14:增加核查人
    If mint编辑状态 = 9 Then    '核查
        '刘兴宏:由于是定价,需要先对定价的价格进行调整.
        If mblnUpdate = False Then
            If Not 检查单价(15, txtNO.Tag, False) Then
                '以最新的价格更新单据体，退出的目的是让用户看一下最终的单据
                ShowMsgBox "有记录未使用最新售价，程序将自动完成更新（售价、售价金额、差价），更新后请检查！"
                Call RefreshBill
                mblnUpdate = True
                Exit Sub
            End If
        End If
        
        If mblnCheckPrice = False Then
            '处理高值卫材虚拟入库产生的入库单检查价格
            If Not CheckValuePrice(15, txtNO.Tag) Then
                ShowMsgBox "高值卫材入库单中价格已调价，程序已自动完成更新（售价、售价金额、，成本价、成本金额、差价）,请检查！"
                mblnCheckPrice = True
                Exit Sub
            End If
        End If
        
        '刘兴宏:2007/06/10:问题10813
        mstrTime_End = GetBillInfo(15, txtNO.Tag)
        If mstrTime_End = "" Then
            MsgBox "注意:" & vbCrLf & "  该单据已经被其他操作员删除,不能继续！", vbInformation, gstrSysName
            Exit Sub
        End If
        If mstrTime_End <> mstrTime_Start Then
            If MsgBox("注意:" & vbCrLf & "  该单据已经被其他操作员编辑，不能继续!" & vbCrLf & "  是否重新刷新单据?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                mshBill.ClearBill
                Call initCard
            End If
            Exit Sub
        End If
        
        If Not SaveCard Then Exit Sub
        Unload Me
        Exit Sub
    End If
    
    If mint编辑状态 = 3 Then      '审核
        If chk转入移库.Value = 1 And mbln退货 = False Then
            If cboType.ItemData(cboType.ListIndex) <> 0 Then
                If Val(txtDraw.Tag) = 0 Then
                    ShowMsgBox "未输入相关的领用部门,请检查!"
                    zlControl.ControlSetFocus txtDraw, True
                    Exit Sub
                End If
                If Trim(txtDrawPerson.Tag) = "" And Trim(txtDrawPerson.Text) <> "" Then
                    If MsgBox("领用人不是当前所属部门的相关人员,是否继续?", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        zlControl.ControlSetFocus txtDrawPerson, True
                        Exit Sub
                    End If
                End If
                If Trim(txtDrawPerson.Tag) = "" And Trim(txtDrawPerson.Text) = "" Then
                    If MsgBox("未输入相关的领用人,是否继续?", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        zlControl.ControlSetFocus txtDrawPerson, True
                        Exit Sub
                    End If
                End If
            Else
                If cboEnterStock.ListIndex < 0 Then
                    ShowMsgBox "要移库的部门不正确！"
                    cboEnterStock.SetFocus
                    Exit Sub
                End If
                If cboStock.ItemData(cboStock.ListIndex) = cboEnterStock.ItemData(cboEnterStock.ListIndex) Then
                    ShowMsgBox "移入部门与移出部门不能相同！"
                    cboEnterStock.SetFocus
                    Exit Sub
                End If
            End If
        End If
        If mblnUpdate = False Then
            If Not 检查单价(15, txtNO.Tag, False) Then
                '以最新的价格更新单据体，退出的目的是让用户看一下最终的单据
                ShowMsgBox "有记录未使用最新售价，程序将自动完成更新（售价、售价金额、差价），更新后请检查！"
                Call RefreshBill
                mblnUpdate = True
                Exit Sub
            End If
        End If
        
        If mblnCheckPrice = False Then
            '处理高值卫材虚拟入库产生的入库单检查价格
            If Not CheckValuePrice(15, txtNO.Tag) Then
                ShowMsgBox "高值卫材入库单中价格已调价，程序已自动完成更新（售价、售价金额、，成本价、成本金额、差价）,请检查！"
                mblnCheckPrice = True
                Exit Sub
            End If
        End If
        
        If Not 材料单据审核(Txt填制人.Caption) Then Exit Sub
                
        If mbln退货 = False Then
            '不是退货单需要重新更新价格,并且价格与原来价格一致时,不用更新价价，因此没必要保存
            blnTrans = True
            gcnOracle.BeginTrans
            
            If mblnUpdate Or mblnChange Or mblnCheckPrice Then
                '刘兴宏:2007/05/15
                '0.与原来价格不一致，需重新保存单据
                '1.更改了原单据,也需要重新保存单据
                If Not SaveCard(True) Then
                    gcnOracle.RollbackTrans: Exit Sub
                End If
                If mbln需要核查 Then
                    '刘兴宏:2007/06/10:问题10813
                    mstrTime_Start = GetBillInfo(15, mstr单据号, False, True)
                Else
                    '刘兴宏:2007/06/10:问题10813
                    mstrTime_Start = GetBillInfo(15, mstr单据号)
                End If
            End If
            
            '刘兴宏:2007/06/10:问题10813
            If mbln需要核查 Then
                mstrTime_End = GetBillInfo(15, txtNO.Tag, False, True)
            Else
                mstrTime_End = GetBillInfo(15, txtNO.Tag)
            End If
            
            If mstrTime_End = "" Then
                gcnOracle.RollbackTrans
                MsgBox "注意:" & vbCrLf & "  该单据已经被其他操作员删除,不能继续！", vbInformation, gstrSysName
                Exit Sub
            End If
            
            If mstrTime_End <> mstrTime_Start Then
                If MsgBox("注意:" & vbCrLf & "  该单据已经被其他操作员编辑，不能继续!" & vbCrLf & "  是否重新刷新单据?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    gcnOracle.RollbackTrans
                    mshBill.ClearBill
                    Call initCard
                    Exit Sub
                Else
                    gcnOracle.RollbackTrans: Exit Sub
                End If
            End If
            
            If SaveCheck = True Then
                Dim blnTemp As Boolean
                strReg = IIf(Val(zlDatabase.GetPara("审核打印", glngSys, mlngModule, "0")) = 1, 1, 0)
                
                If chk转入移库.Value = 1 Then
                    If cboType.ItemData(cboType.ListIndex) <> 0 Then
                        '保存数据
                        gcnOracle.CommitTrans
                        blnTemp = False
                        If mbln退货 = False Then
                            frmDrawCard.ShowCard Me, txtNO.Text, 7, , , blnSuccess, Val(txtDraw.Tag), Trim(txtDrawPerson.Text)
                        End If
                    Else
                        If Check移库(blnTemp) = False Then
                            If blnTemp = True Then gcnOracle.RollbackTrans: Exit Sub
                            gcnOracle.CommitTrans
                            blnTemp = False
                        Else
                            gcnOracle.CommitTrans
                            blnTemp = True
                        End If
                    End If
                Else
                    gcnOracle.CommitTrans
                    blnTemp = False
                End If
                
                If Val(strReg) = 1 Then
                    '打印
                    If InStr(mstrPrivs, "单据打印") <> 0 Then
                        printbill
                    End If
                End If
                If blnTemp Then
                    frmTransferCard.ShowCard Me, txtNO.Text, 11, , , blnSuccess
                End If
                Unload Me
                Exit Sub
            Else
                gcnOracle.RollbackTrans
            End If
        
        Else
            If Not 库存检查_退货 Then Exit Sub
            If SaveCheck = True Then
                strReg = IIf(Val(zlDatabase.GetPara("审核打印", glngSys, mlngModule, "0")) = 1, 1, 0)
                If Val(strReg) = 1 Then
                    '打印
                    If InStr(mstrPrivs, "单据打印") <> 0 Then
                        printbill
                    End If
                End If
                Unload Me
                Exit Sub
            End If
        End If
        blnTrans = False
        mblnUpdate = False
        mblnCheckPrice = False
        Exit Sub
    End If
            
    If mint编辑状态 = 5 Then      '修改发票信息
        If SaveRecipe = True Then
            Unload Me
        End If
        Exit Sub
    End If
    
    If mint编辑状态 = 10 Then      '修改注册证号
        If SaveRegist = True Then
            Unload Me
        End If
        Exit Sub
    End If
    
    If mint编辑状态 = 6 Then
        If SaveStrike = True Then
            Unload Me
        End If
        Exit Sub
    End If
    
    If mint编辑状态 = 8 Then    '退货
        If ValidData = False Then Exit Sub
        If SaveRestore Then
            strReg = IIf(Val(zlDatabase.GetPara("存盘打印", glngSys, mlngModule, "0")) = 1, 1, 0)
            If Val(strReg) = 1 Then
                '打印
                If InStr(mstrPrivs, "单据打印") <> 0 Then
                    printbill
                End If
            End If
            Unload Me
            Exit Sub
        End If
    End If
            
    If ValidData = False Then Exit Sub
    If Not CheckProvider Then Exit Sub

    blnSuccess = SaveCard
        
    If blnSuccess = True Then
        '清空本地数据集
        If mrsCostlyInfo.RecordCount > 0 Then
            mrsCostlyInfo.MoveFirst
            Do While Not mrsCostlyInfo.EOF
                mrsCostlyInfo.Delete
                mrsCostlyInfo.MoveNext
            Loop
        End If
        strReg = IIf(Val(zlDatabase.GetPara("存盘打印", glngSys, mlngModule, "0")) = 1, 1, 0)
        If Val(strReg) = 1 Then
            '打印
            If InStr(mstrPrivs, "单据打印") <> 0 Then
                printbill
            End If
        End If
        If mint编辑状态 = 2 Then   '修改
            Unload Me
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    
'    If mbln单据增加 Then
'        mstr单据号 = NextNo(68)
'    End If
    txtNO = ""
    mblnSave = False
    mblnEdit = True
    mshBill.ClearBill
    
    Call RefreshRowNO(mshBill, mCol行号, 1)
    
    SetEdit
    
    txtProvider.Text = ""
    txtProvider.Tag = "0"
    txt摘要.Text = ""
    txtProvider.SetFocus
    mblnChange = False
    vsfCostlyInfo.Visible = False: Call Form_Resize
    If txtNO.Tag <> "" Then Me.stbThis.Panels(2).Text = "上一张单据的NO号：" & txtNO.Tag
    Exit Sub
ErrHand:
    If blnTrans Then
        gcnOracle.RollbackTrans
    End If
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function ProduceDateCheck(ByVal strDate As String) As Boolean
    '功能：生产日期，注册证效期检查
    'strdate 生产日期
    '返回值：true-检查通过，false-检查未通过
    If mintProduceDate = 1 Then
        With mshBill
            If .TextMatrix(.Row, mcol注册证有效期) = "" Then
                ProduceDateCheck = True '注册证有效期为空则不检查
                Exit Function
            Else
                If CDate(strDate) > CDate(.TextMatrix(.Row, mcol注册证有效期)) Then
                    If mintCheckType = 1 Then
                        If MsgBox("生产日期大于注册证有效期，此卫材为无证生产卫材，是否继续？", vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                            ProduceDateCheck = True
                        Else
                            ProduceDateCheck = False
                        End If
                    ElseIf mintCheckType = 2 Then
                        MsgBox "生产日期大于注册证有效期，此卫材为无证生产卫材！", vbInformation, gstrSysName
                        ProduceDateCheck = False
                    Else
                        ProduceDateCheck = True
                    End If
                Else
                    ProduceDateCheck = True
                End If
            End If
            
        End With
    Else
        ProduceDateCheck = True
    End If
End Function

Private Sub Form_Load()
    Dim strReg As String
    Dim strCheck As String
    
    strCheck = zlDatabase.GetPara("资质校验", glngSys, mlngModule, 0)
    If InStr(1, strCheck, "|") > 0 Then
        '校验方式：0-不检查；1－提醒；2－禁止
        mintCheckType = Val(Mid(strCheck, 1, InStr(1, strCheck, "|") - 1))
    End If
    mintProduceDate = Val(zlDatabase.GetPara("生产日期效期检查", glngSys, mlngModule, "0"))
    
    Me.lblType.Caption = "病人ID↓": Me.lblType.Tag = 1
    
    strReg = Val(zlDatabase.GetPara("卫材单位", glngSys, mlngModule, "0"))
    mintUnit = Val(strReg)
    strReg = Val(zlDatabase.GetPara("审核产生单据", glngSys, mlngModule, "0"))
    mblnCostView = zlStr.IsHavePrivs(mstrPrivs, "查看成本价")
    mblnFirst = True
    With cboType
        .AddItem "移库到"
        .ItemData(.NewIndex) = 0
        If Val(strReg) = 0 Then .ListIndex = 0
        .AddItem "领用到"
        .ItemData(.NewIndex) = 1
        If Val(strReg) = 1 Then .ListIndex = 1
        If .ListIndex < 0 Then .ListIndex = 0
    End With
    Call chk转入移库_Click
    
    mbln非中标单位入库 = IIf(Val(zlDatabase.GetPara("招标卫材可选择非中标单位入库", glngSys, mlngModule, "0")) = 1, 1, 0) = 1
   
    '刘兴宏:增加小数格式化串
    With mFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价)
        .FM_金额 = GetFmtString(mintUnit, g_金额)
        .FM_零售价 = GetFmtString(mintUnit, g_售价)
        .FM_数量 = GetFmtString(mintUnit, g_数量)
        .FM_散装零售价 = GetFmtString(2, g_售价)
    End With
    
    '刘兴宏:增加核查功能,2007/05/13
    '外购入库，需要确定是否需要核查功能
    mbln需要核查 = Val(zlDatabase.GetPara("卫材外购需要核查", glngSys, 0)) = 1
     
    lbl核查人.Visible = mbln需要核查
    lbl核查日期.Visible = mbln需要核查
    txt核查人.Visible = mbln需要核查
    txt核查日期.Visible = mbln需要核查
    
    mintBatchNoLen = GetBatchNoLen()
    mint发票号Len = Get发票号Len
    
    mbln加价率 = Get加价率()
    mbln分段加成率 = IS分段加成率()
    mbln时价购前销售 = ISCHECK外购扣前销售()
    mbln不强制控制指导价格 = ISCHECK不强制控制指导价格()
    mbln时价卫材直接确定售价 = is时价卫材直接确定售价()
    mbln时价卫材取上次售价 = is时价卫材取上次售价()
    
    mbln分批卫材批号产地控制 = Val(zlDatabase.GetPara(305, glngSys, 0)) = 1
    
    mbln退货 = False
    
    txtNO = mstr单据号
    txtNO.Tag = txtNO.Text
    '恢复个性化参数设置
    RestoreWinState Me, App.ProductName, mstrCaption
    
    Call initCard
    If mint编辑状态 <> 6 Then
        If mshBill.ColWidth(mCol冲销数量) > 0 Then
            mshBill.ColWidth(mCol冲销数量) = 0
        End If
    Else
        If mshBill.ColWidth(mCol冲销数量) = 0 Then
            mshBill.ColWidth(mCol冲销数量) = 800
        End If
    End If
    Call RestoreBILLWidthSet
    mblnUpdate = False
    mblnChange = False
    mblnCheckPrice = False
    mshBill.ColWidth(mCol序号) = 0
    '恢复个性化参数设置后，还需要对权限控制的列进一步设置
    With mshBill
        .ColWidth(mCol结算价) = IIf(mblnCostView = False, 0, 900)
        .ColWidth(mCol采购价) = IIf(mblnCostView = False, 0, 900)
        .ColWidth(mCol结算金额) = IIf(mblnCostView = False, 0, 900)
        .ColWidth(mCol差价) = IIf(mblnCostView = False, 0, 900)
    End With
End Sub

Private Sub initCard()
    '------------------------------------------------------------------------------------------------
    '功能:初始化卡片信息
    '------------------------------------------------------------------------------------------------
    Dim rstemp As New Recordset
    Dim strUnit As String, strUnitQuantity As String, str批次 As String
    Dim num包装系数 As String, strOrder As String, strCompare As String, strReg As String
    Dim dblSum As Double, intRow As Integer, i As Integer
    Dim varStuff As Variant
    Dim lngProviderID As Long
    Dim strDateBegin As String, strDateEnd As String, strIVNO As String
    Dim dtIVDate As Date
    Dim rs As ADODB.Recordset
    Dim dbl采购价 As Double
     
    strReg = zlDatabase.GetPara("单据排序", glngSys, mlngModule, "00")
    strOrder = IIf(strReg = "", "00", strReg)
    On Error GoTo ErrHandle
    '库房
    strCompare = Mid(strOrder, 1, 1)
    
    '卫材退货
    If mint编辑状态 = 8 Then
        cmdExtractData.Visible = True
    Else
        cmdExtractData.Visible = False
        lblCode.Left = lblCode.Left - cmdExtractData.Width - 100
        txtCode.Left = txtCode.Left - cmdExtractData.Width - 100
        cmdFind.Left = cmdFind.Left - cmdExtractData.Width - 100
    End If
    
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
            Set rstemp = ReturnSQL(cboStock.ItemData(cboStock.ListIndex), mstrCaption, True)
            Do While Not rstemp.EOF
                .AddItem rstemp.Fields(2)
                .ItemData(.NewIndex) = rstemp.Fields(0)
                rstemp.MoveNext
            Loop
    
            If .ListCount > 0 Then
                .ListIndex = 0
            End If
                                
        End With
    End If
    
    Select Case mint编辑状态
        Case 1, 8           '新增或退货
        
            Txt填制人 = UserInfo.用户名
            Txt填制日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            initGrid
        Case 2, 3, 4, 5, 6, 7, 9, 10
                
            initGrid
            
            If mint编辑状态 = 4 Then
                gstrSQL = "" & _
                    "   Select b.id,b.名称 " & _
                    "   From 药品收发记录 a,部门表 b " & _
                    "   where a.库房id=b.id and A.单据 = 15 and a.no=[1] "
                
                Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr单据号)
                If rstemp.EOF Then
                    mintParallelRecord = 2
                    Exit Sub
                End If
                
                With cboStock
                    .AddItem rstemp!名称
                    .ItemData(.NewIndex) = rstemp!Id
                    .ListIndex = 0
                    mintcboIndex = 0
                End With
                rstemp.Close
            End If
            
             
            Select Case mintUnit
                Case 0
                    strUnitQuantity = "D.计算单位 AS 单位,D.计算单位  as 散装单位, decode(nvl(A.发药方式,0),1,-1,1)* A.填写数量  AS 数量,1 as 换算系数,"
                    num包装系数 = "1"
                Case 1
                    strUnitQuantity = "B.包装单位 AS 单位,D.计算单位  as  散装单位, decode(nvl(A.发药方式,0),1,-1,1)* (A.填写数量 / B.换算系数) AS 数量,B.换算系数 ,"
                    num包装系数 = "B.换算系数"
            End Select
            
            
            Select Case mint编辑状态
            Case 5
                '修改发票信息
                If mint记录状态 = 1 Then GoTo Go正常:
                    gstrSQL = "" & _
                        "   Select * " & _
                        "   From (  SELECT distinct a.药品id as 材料id,A.序号,('[' || D.编码 || ']' || D.名称) AS 卫材信息,zlSpellCode(D.名称) 名称,E.名称 商品名,D.规格,D.产地 as 原产地,A.产地,a.批准文号, A.批号,A.批次,to_char(A.生产日期,'yyyy-mm-dd') 生产日期," & _
                        "                   b.最大效期,A.效期,b.一次性材料,Nvl(b.是否条码管理,0) as 条码管理,b.库房分批,b.灭菌效期,A.灭菌日期,a.灭菌效期 as 灭菌失效期,Nvl(b.加成率,0)/100 as 加成率," & strUnitQuantity & _
                        "                   b.指导批发价*" & num包装系数 & " as 指导批发价 ,a.成本价*" & num包装系数 & " AS 结算价, decode(nvl(A.发药方式,0),1,-1,1)*a.成本金额 as 结算金额,b.指导差价率/100 as 指导差价率,d.是否变价,b.在用分批,nvl(a.发药方式,0) 退货," & _
                        "                   DECODE(A.扣率, NULL, 0, A.扣率) AS 扣率, A.零售价 ,a.零售金额,a.差价,a.零售差价," & _
                        "                   a.随货单号,a.发票号 ,a.发票代码, a.发票日期,a.发票金额,a.供药单位id,f.名称 as 供应商,a.注册证号,a.商品条码,a.填制人,a.填制日期," & _
                        "                   a.审核人,a.审核日期,a.库房id,g.名称 as 部门,nvl(a.付款序号,0) as 付款序号,A.核查人,A.核查日期,a.内部条码,a.费用ID,b.注册证有效期 " & _
                        "           FROM (  select min(X.id) as id,max(Nvl(x.发药方式,0)) 发药方式, " & _
                        "                           sum(x.实际数量) as 填写数量,sum(x.成本金额) as 成本金额 ,sum(x.零售金额) as 零售金额,sum(x.差价) as 差价,sum(to_number(nvl(to_char(x.用法," & gOraFmt_Max.FM_金额 & " ),0), " & gOraFmt_Max.FM_金额 & ")) as 零售差价," & _
                        "                           y.随货单号,y.发票号,y.发票代码,y.发票日期,sum(y.发票金额) as 发票金额," & _
                        "                           X.药品ID,X.序号,X.产地,X.批准文号, X.批号,NVL(X.批次,0) 批次,X.生产日期,X.效期,X.灭菌效期,X.灭菌日期,X.扣率,X.成本价,X.零售价," & _
                        "                           x.供药单位ID,X.注册证号,X.商品条码,x.库房ID,max(x.填制人) as 填制人,max(x.填制日期) as 填制日期,max(x.审核人) as 审核人," & _
                        "                           max(x.审核日期) as 审核日期,max(x.配药人) as 核查人,x.内部条码,x.费用ID," & _
                        "                           max(x.配药日期) as 核查日期,Nvl(Y.付款序号,0) as 付款序号 " & _
                        "                   From 药品收发记录 x,应付记录 y " & _
                        "                   WHERE x.id=y.收发id(+)  and y.系统标识(+)=5 and y.记录性质(+)=0 and X.NO=[1] AND 单据=15  " & _
                        "                   group by X.药品ID,X.序号,X.产地,X.批准文号,X.批号,NVL(X.批次,0)  ,x.生产日期,X.效期,X.灭菌效期,X.灭菌日期,X.扣率,X.成本价,X.零售价," & _
                        "                            x.供药单位ID,X.注册证号,X.商品条码,X.库房ID,X.内部条码,X.费用ID,y.随货单号,y.发票号,y.发票代码,y.发票日期,NVL(Y.付款序号,0) " & _
                        "                   having sum(实际数量)<>0 ) A," & _
                        "                   材料特性 B,收费项目目录 D,供应商 f,部门表 g,收费项目别名 e  " & _
                        "           Where A.药品id = B.材料id and a.药品id=D.id and a.供药单位id=f.id and a.库房id=g.id And d.id  = e.收费细目id(+) And e.性质(+) = 3  " & _
                        "          ) " & _
                        "   ORDER BY " & IIf(strCompare = "0", "序号", IIf(strCompare = "1", "卫材信息", "名称")) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
            Case 6
                '冲销
                gstrSQL = "" & _
                    "   Select * " & _
                    "   From (  SELECT distinct a.药品id as 材料id,A.序号,('[' || D.编码 || ']' || D.名称) AS 卫材信息,zlSpellCode(D.名称) 名称,e.名称 商品名,D.规格,D.产地 as 原产地,A.产地,a.批准文号, A.批号,to_char(A.生产日期,'yyyy-mm-dd') 生产日期," & _
                    "                   b.最大效期,A.效期,b.一次性材料,nvl(b.是否条码管理,0) as 条码管理,b.库房分批,b.灭菌效期,A.灭菌日期,a.灭菌效期 as 灭菌失效期,Nvl(b.加成率,0)/100 as 加成率," & strUnitQuantity & _
                    "                   b.指导批发价*" & num包装系数 & " as 指导批发价 ,a.成本价*" & num包装系数 & " AS 结算价, " & _
                    "                   decode(nvl(A.发药方式,0),1,-1,1)*a.成本金额 as 结算金额,b.指导差价率/100 as 指导差价率,d.是否变价,b.在用分批,nvl(a.发药方式,0) 退货," & _
                    "                   DECODE(A.扣率, NULL, 0, A.扣率) AS 扣率,A.零售价,decode(nvl(A.发药方式,0),1,-1,1)* nvl(A.零售金额,0) as 零售金额,decode(nvl(A.发药方式,0),1,-1,1)*nvl(A.差价,0)  as 差价,decode(nvl(A.发药方式,0),1,-1,1)*nvl(A.零售差价,0) as 零售差价, " & _
                    "                   a.随货单号,a.发票号 ,a.发票代码, a.发票日期,0 as 发票金额,a.供药单位id,f.名称 as 供应商,a.注册证号,a.商品条码,a.库房id,g.名称 as 部门," & _
                    "                   nvl(a.付款序号,0) as 付款序号,A.核查人,A.核查日期,a.内部条码,a.费用ID,b.注册证有效期 " & _
                    "           FROM (  select min(X.id) as id,max(Nvl(x.发药方式,0)) 发药方式, sum(填写数量) as 填写数量,sum(成本金额) as 成本金额," & _
                    "                           y.随货单号,y.发票号,y.发票代码,y.发票日期,X.药品ID,X.序号,X.产地,X.批准文号, X.批号,X.生产日期,X.效期,X.灭菌效期,X.灭菌日期,X.扣率,X.成本价,X.零售价," & _
                    "                           sum(x.零售金额) as 零售金额,sum(x.差价) as 差价,sum(to_number(nvl(to_char(x.用法," & gOraFmt_Max.FM_金额 & " ),0), " & gOraFmt_Max.FM_金额 & ")) as 零售差价," & _
                    "                           x.供药单位ID,X.注册证号,X.商品条码,x.库房ID,max(x.配药人) as 核查人,max(x.配药日期) as 核查日期," & _
                    "                           x.内部条码, x.费用ID, Nvl(Y.付款序号,0) as 付款序号 " & _
                    "                   From 药品收发记录 x,(Select Id, 记录性质, 记录状态, No, 项目id, 序号, 收发id, 单位id, 品名, 规格, 产地, 批号, 计量单位, 入库单据号, 单据金额, 数量, 采购价, 采购金额, 随货单号, 发票号,发票代码, 发票日期, 发票金额, 制定日期, 计划金额, 计划人, 计划日期, 填制人, 填制日期, 审核人, 审核日期, 付款序号, 计划序号, 系统标识 From 应付记录 Where 系统标识=5 And 记录性质=0) y " & _
                    "                   WHERE x.id=y.收发id(+) and X.NO=[1] AND 单据=15  " & _
                    "                   group by X.药品ID,X.序号,X.产地,X.批准文号,X.批号,x.生产日期,X.效期,X.灭菌效期,X.灭菌日期,X.扣率,X.成本价,X.零售价," & _
                    "                            x.供药单位ID,X.注册证号,X.商品条码,X.库房ID,x.内部条码,x.费用ID,y.随货单号,y.发票号,y.发票代码,y.发票日期,NVL(Y.付款序号,0) " & _
                    "                   having sum(填写数量)<>0 ) A," & _
                    "                   材料特性 B,收费项目目录 D,供应商 f,部门表 g,收费项目别名 e " & _
                    "           Where A.药品id = B.材料id and a.药品id=D.id and a.供药单位id=f.id and a.库房id=g.id And d.id  = e.收费细目id(+) And e.性质(+) = 3  " & _
                    "          ) " & _
                    "   ORDER BY " & IIf(strCompare = "0", "序号", IIf(strCompare = "1", "卫材信息", "名称")) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
            Case 10
                '修改注册证号
                gstrSQL = "" & _
                    "   Select * " & _
                    "   From (  SELECT distinct a.药品id 材料id,A.序号,b.一次性材料,nvl(b.是否条码管理,0) as 条码管理,b.库房分批,('[' || D.编码 || ']' || D.名称) AS 卫材信息,zlSpellCode(D.名称) 名称,e.名称 商品名, D.规格,D.产地 as 原产地,A.产地,A.验收结论,A.批准文号, A.批号,Nvl(A.批次,0) 批次,to_char(A.生产日期,'yyyy-mm-dd') 生产日期," & _
                    "                   b.最大效期,A.效期,b.灭菌效期,A.灭菌日期,a.灭菌效期 as 灭菌失效期,Nvl(b.加成率,0)/100 as 加成率," & strUnitQuantity & _
                    "                   b.指导批发价*" & num包装系数 & " as 指导批发价 ,A.成本价*" & num包装系数 & " AS 结算价, decode(nvl(A.发药方式,0),1,-1,1)*A.成本金额 AS 结算金额,Nvl(A.发药方式,0) 退货,b.指导差价率/100 as 指导差价率,D.是否变价,b.在用分批," & _
                    "                   DECODE(A.扣率, NULL, 0, A.扣率) AS 扣率, A.零售价, decode(nvl(A.发药方式,0),1,-1,1)*A.零售金额 零售金额, " & _
                    "                   decode(nvl(A.发药方式,0),1,-1,1)* A.差价 差价,  decode(nvl(A.发药方式,0),1,-1,1)*to_number(nvl(to_char(A.用法," & gOraFmt_Max.FM_金额 & " ),0), " & gOraFmt_Max.FM_金额 & ")   as 零售差价, " & _
                    "                   C.随货单号,C.发票号 ,c.发票代码, C.发票日期, decode(nvl(A.发药方式,0),1,-1,1)*C.发票金额 发票金额,a.供药单位id,f.名称 as 供应商,a.注册证号,a.商品条码, a.摘要,A.填制人,A.填制日期,A.配药人 as 核查人,A.配药日期 as 核查日期,A.审核人,A.审核日期," & _
                    "                   a.库房id,g.名称 as 部门,nvl(c.付款序号,0) as 付款序号, a.内部条码, a.费用id,b.注册证有效期 " & _
                    "           FROM 药品收发记录 A, 材料特性 B,收费项目目录 D, (Select Id, 记录性质, 记录状态, No, 项目id, 序号, 收发id, 单位id, 品名, 规格, 产地, 批号, 计量单位, 入库单据号, 单据金额, 数量, 采购价, 采购金额, 随货单号, 发票号,发票代码, 发票日期, 发票金额, 制定日期, 计划金额, 计划人, 计划日期, 填制人, 填制日期, 审核人, 审核日期, 摘要, 付款序号, 计划序号, 系统标识 From 应付记录 Where 系统标识=5 And 记录性质=0) C,供应商 f,部门表 g,收费项目别名 e  " & _
                    "           Where A.药品id = B.材料id  and a.药品id=d.ID AND A.Id = C.收发id (+)   and a.供药单位id=f.id and a.库房id=g.id And d.id  = e.收费细目id(+) And e.性质(+) = 3  " & _
                    "                   AND (A.记录状态 = 1 Or Mod(A.记录状态, 3) = 0) And A.审核日期 Is Not Null " & _
                    "                   AND A.单据 = 15 AND A.No = [1] " & _
                    "       ) " & _
                    "   ORDER BY " & IIf(strCompare = "0", "序号", IIf(strCompare = "1", "卫材信息", "名称")) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
            
            Case Else
            
Go正常:
                gstrSQL = "" & _
                    "   Select * " & _
                    "   From (  SELECT distinct a.药品id 材料id,A.序号,b.一次性材料,nvl(b.是否条码管理,0) as 条码管理,b.库房分批,('[' || D.编码 || ']' || D.名称) AS 卫材信息,zlSpellCode(D.名称) 名称,E.名称 商品名,D.规格,D.产地 as 原产地,A.产地,A.验收结论,A.批准文号, A.批号,Nvl(A.批次,0) 批次,to_char(A.生产日期,'yyyy-mm-dd') 生产日期," & _
                    "               b.最大效期,A.效期,b.灭菌效期,A.灭菌日期,a.灭菌效期 as 灭菌失效期,Nvl(b.加成率,0)/100 as 加成率," & strUnitQuantity & _
                    "               b.指导批发价*" & num包装系数 & " as 指导批发价 ,A.成本价*" & num包装系数 & " AS 结算价, decode(nvl(A.发药方式,0),1,-1,1)*A.成本金额 AS 结算金额,Nvl(A.发药方式,0) 退货,b.指导差价率/100 as 指导差价率,D.是否变价,b.在用分批," & _
                    "               DECODE(A.扣率, NULL, 0, A.扣率) AS 扣率, A.零售价, decode(nvl(A.发药方式,0),1,-1,1)*A.零售金额 零售金额, " & _
                    "               decode(nvl(A.发药方式,0),1,-1,1)* A.差价 差价,  decode(nvl(A.发药方式,0),1,-1,1)*to_number(nvl(to_char(A.用法," & gOraFmt_Max.FM_金额 & " ),0), " & gOraFmt_Max.FM_金额 & ")   as 零售差价, " & _
                    "               C.随货单号,C.发票号 ,c.发票代码, C.发票日期, decode(nvl(A.发药方式,0),1,-1,1)*C.发票金额 发票金额,a.供药单位id,f.名称 as 供应商,a.注册证号,a.商品条码, a.摘要,A.填制人,A.填制日期,A.配药人 as 核查人,A.配药日期 as 核查日期,A.审核人,A.审核日期," & _
                    "               a.库房id,g.名称 as 部门,nvl(c.付款序号,0) as 付款序号, a.内部条码, a.费用id,b.注册证有效期 " & _
                    "           FROM 药品收发记录 A, 材料特性 B,收费项目目录 D, (Select Id, 记录性质, 记录状态, No, 项目id, 序号, 收发id, 单位id, 品名, 规格, 产地, 批号, 计量单位, 入库单据号, 单据金额, 数量, 采购价, 采购金额, 随货单号, 发票号,发票代码, 发票日期, 发票金额, 制定日期, 计划金额, 计划人, 计划日期, 填制人, 填制日期, 审核人, 审核日期, 摘要, 付款序号, 计划序号, 系统标识 From 应付记录 Where 系统标识=5 And 记录性质=0) C,供应商 f,部门表 g,收费项目别名 e " & _
                    "           Where A.药品id = B.材料id  and a.药品id=d.ID AND A.Id = C.收发id (+)   and a.供药单位id=f.id and a.库房id=g.id And d.id  = e.收费细目id(+) And e.性质(+) = 3 " & _
                    "                   AND A.记录状态 =[2] " & _
                    "                   AND A.单据 = 15 AND A.No = [1] " & _
                    "       ) " & _
                    "   ORDER BY " & IIf(strCompare = "0", "序号", IIf(strCompare = "1", "卫材信息", "名称")) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
            End Select
            
            
            Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr单据号, mint记录状态, _
                            cboStock.ItemData(cboStock.ListIndex), _
                            lngProviderID, strDateBegin, strDateEnd)
            
            If rstemp.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            If mint编辑状态 = 9 Then
                '刘兴宏:2007/06/10:问题10813
                mstrTime_Start = GetBillInfo(15, mstr单据号)
            ElseIf mint编辑状态 = 3 And mbln需要核查 Then
                '刘兴宏:2007/06/10:问题10813
                mstrTime_Start = GetBillInfo(15, mstr单据号, False, True)
            Else
                '刘兴宏:2007/06/10:问题10813
                mstrTime_Start = GetBillInfo(15, mstr单据号)
            End If
            Select Case mint编辑状态
                Case 2, 6
                    Txt填制人 = UserInfo.用户名
                    Txt填制日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                    If mint编辑状态 = 2 Then
                        Txt审核人 = ""
                        Txt审核日期 = ""
                        txt核查人 = ""
                        txt核查日期 = ""
                    Else
                        Txt审核人 = UserInfo.用户名
                        Txt审核日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                        txt核查人 = UserInfo.用户名
                        txt核查日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                    End If
                Case 9
                    Txt填制人 = zlStr.Nvl(rstemp!填制人)
                    Txt填制日期 = IIf(zlStr.Nvl(rstemp!填制日期) = "", "", Format(rstemp!填制日期, "yyyy-mm-dd hh:mm:ss"))
                    txt核查人 = UserInfo.用户名
                    txt核查日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                    Txt审核人 = ""
                    Txt审核日期 = ""
                Case Else
                    Txt填制人 = zlStr.Nvl(rstemp!填制人)
                    Txt填制日期 = Format(rstemp!填制日期, "yyyy-mm-dd hh:mm:ss")
                    txt核查人 = zlStr.Nvl(rstemp!核查人) 'UserInfo.用户姓名
                    txt核查日期 = IIf(IsNull(rstemp!核查日期), "", Format(rstemp!核查日期, "yyyy-mm-dd hh:mm:ss"))
                    Txt审核人 = IIf(IsNull(rstemp!审核人), "", rstemp!审核人)
                    Txt审核日期 = IIf(IsNull(rstemp!审核日期), "", Format(rstemp!审核日期, "yyyy-mm-dd hh:mm:ss"))
            End Select
            txtProvider.Text = rstemp!供应商
            txtProvider.Tag = rstemp!供药单位ID
            mbln退货 = (rstemp!退货 = 1)
'            txt摘要.Text = IIf(IsNull(rsTemp!摘要), "", rsTemp!摘要)
            If mint编辑状态 = 5 Or mint编辑状态 = 6 Then
                txt摘要.Text = Get摘要(mstr单据号, mint编辑状态)
            Else
                txt摘要.Text = IIf(IsNull(rstemp!摘要), "", rstemp!摘要)
            End If
            
            If (mint编辑状态 = 2 Or mint编辑状态 = 3) And Txt审核人 <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            If mint编辑状态 = 5 Or mint编辑状态 = 7 Then
                If rstemp!付款序号 <> 0 Then
                    mintParallelRecord = IIf(mint编辑状态 = 5, 4, 5)        '已被其他人付款
                    Exit Sub
                ElseIf mint编辑状态 = 7 Then
                    '检查是否存在部分付款的情况
                    gstrSQL = "Select Nvl(Max(付款序号), 0) 付款序号 From 应付记录 " & _
                        " where 收发id=(Select Id From 药品收发记录 Where 单据=15 And No=[1] And (Mod(记录状态,3)=0 Or 记录状态=1) " & _
                        " And 序号=[2]) "
                    strOrder = rstemp!序号
                    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "[取付款序号]", txtNO.Text, strOrder)
                    
                    If rs!付款序号 <> 0 Then
                        mintParallelRecord = 5
                    End If
                End If
            End If
            
            If mbln退货 Then LblTitle.Caption = "卫生材料退货单"
            intRow = 0
            If mbln退货 Or mint编辑状态 = 3 Then
                Set mCllBillData = New Collection
            End If
                        
            With mshBill
                Do While Not rstemp.EOF
                    intRow = intRow + 1
                    .Rows = .Rows + 1
                    
                    .TextMatrix(intRow, 0) = rstemp.Fields(0)
                    .TextMatrix(intRow, mCol诊疗) = rstemp!卫材信息
                    .TextMatrix(intRow, mCol商品名) = IIf(IsNull(rstemp!商品名), "", rstemp!商品名)
                    .TextMatrix(intRow, mCol规格) = IIf(IsNull(rstemp!规格), "", rstemp!规格)
                    .TextMatrix(intRow, mCol产地) = IIf(IsNull(rstemp!产地), "", rstemp!产地)
                    .TextMatrix(intRow, mCol单位) = rstemp!单位
                    .TextMatrix(intRow, mCol批号) = IIf(IsNull(rstemp!批号), "", rstemp!批号)
                    .TextMatrix(intRow, mcol批准文号) = IIf(IsNull(rstemp!批准文号), "", rstemp!批准文号)
                    .TextMatrix(intRow, mCol效期) = IIf(IsNull(rstemp!效期), "", rstemp!效期)
                    .TextMatrix(intRow, mCol数量) = Format(rstemp!数量, mFMT.FM_数量)
                    .TextMatrix(intRow, mCol结算价) = Format(rstemp!结算价, mFMT.FM_成本价)
                    .TextMatrix(intRow, mCol采购价) = Format(Val(.TextMatrix(intRow, mCol结算价)) * 100 / IIf(Val(zlStr.Nvl(rstemp!扣率)) = 0, 1, Val(zlStr.Nvl(rstemp!扣率))), mFMT.FM_成本价)
                    .TextMatrix(intRow, mCol结算金额) = Format(IIf(mint编辑状态 = 6, 0, rstemp!结算金额), mFMT.FM_金额)
                    
                    '刘兴宏:零售价处理:零售价-->零售价;零售金额-->零售金额;差价-->零售差价;用途-->库房单位差价
                    ' 零售金额＝入库数量×零售价；
                    ' 零售差价"＝售价金额－零售金额，即按入库单位计算的金额和按零售单位计算的金额的差值；
                    .TextMatrix(intRow, mcol零售价) = Format(Val(zlStr.Nvl(rstemp!零售价)), mFMT.FM_散装零售价)          'If Val(.TextMatrix(.row, mcol零售价)) = 0 Then
                    .TextMatrix(intRow, mCol售价) = Format(Val(zlStr.Nvl(rstemp!零售价)) * Val(zlStr.Nvl(rstemp!换算系数)), mFMT.FM_零售价)
                    
                    '反算售价
'                    .TextMatrix(intRow, mCol售价) = Format((Val(NVL(rsTemp!零售金额)) - Val(NVL(rsTemp!零售差价))) / Val(NVL(rsTemp!数量)), mFMT.FM_零售价)
                    .TextMatrix(intRow, mcol零售单位) = zlStr.Nvl(rstemp!散装单位)
                    
                    If mint编辑状态 = 6 Then
                        '冲销没有相关的差价
                        .TextMatrix(intRow, mcol零售差价) = ""
                        .TextMatrix(intRow, mcol零售金额) = ""
                        .TextMatrix(intRow, mCol差价) = ""
                        .TextMatrix(intRow, mCol售价金额) = ""
                    Else
                        .TextMatrix(intRow, mcol零售差价) = Format(Val(zlStr.Nvl(rstemp!差价)), mFMT.FM_金额)
                        .TextMatrix(intRow, mcol零售金额) = Format(Val(zlStr.Nvl(rstemp!零售金额)), mFMT.FM_金额)
                        '反算售价及售价金额
'                        .TextMatrix(intRow, mCol差价) = Format(Val(NVL(rsTemp!差价)) - Val(NVL(rsTemp!零售差价)), mFMT.FM_金额)
'                        .TextMatrix(intRow, mCol售价金额) = Format(Val(NVL(rsTemp!零售金额)) - Val(NVL(rsTemp!零售差价)), mFMT.FM_金额)
                        .TextMatrix(intRow, mCol差价) = Format(Val(zlStr.Nvl(rstemp!差价)), mFMT.FM_金额)
                        .TextMatrix(intRow, mCol售价金额) = Format(Val(zlStr.Nvl(rstemp!零售金额)), mFMT.FM_金额)
                    End If
                    
                    .TextMatrix(intRow, mCol扣率) = rstemp!扣率
                    .TextMatrix(intRow, mCol随货单号) = IIf(IsNull(rstemp!随货单号), "", rstemp!随货单号)
                    
                    If mint编辑状态 <> 6 Then '冲销不显示验收结论
                        If (mint编辑状态 = 5 And mint记录状态 <> 1) Then '修改发票信息非政策记录不显示验收结论
                        Else
                            .TextMatrix(intRow, mCol验收结论) = IIf(IsNull(rstemp!验收结论), "", rstemp!验收结论)
                        End If
                    End If
                    
                    .TextMatrix(intRow, mCol发票号) = IIf(IsNull(rstemp!发票号), "", rstemp!发票号)
                    .TextMatrix(intRow, mcol发票代码) = IIf(IsNull(rstemp!发票代码), "", rstemp!发票代码)
                    .TextMatrix(intRow, mCol发票日期) = IIf(IsNull(rstemp!发票日期), "", rstemp!发票日期)
                    .TextMatrix(intRow, mCol发票金额) = IIf(Format(IIf(IsNull(rstemp!发票金额), "0", rstemp!发票金额), mFMT.FM_金额) = "0.00", "", Format(IIf(IsNull(rstemp!发票金额), "0", rstemp!发票金额), mFMT.FM_金额))
                    
                    .TextMatrix(intRow, mcol一次性材料) = zlStr.Nvl(rstemp!一次性材料)
                    .TextMatrix(intRow, mcol条码管理) = zlStr.Nvl(rstemp!条码管理)
                    .TextMatrix(intRow, mcol灭菌效期) = zlStr.Nvl(rstemp!灭菌效期)
                    .TextMatrix(intRow, mcol灭菌日期) = zlStr.Nvl(rstemp!灭菌日期)
                    .TextMatrix(intRow, mcol灭菌失效期) = zlStr.Nvl(rstemp!灭菌失效期)
                    .TextMatrix(intRow, mcol生产日期) = zlStr.Nvl(rstemp!生产日期)
                    .TextMatrix(intRow, mcol注册证有效期) = IIf(IsNull(rstemp!注册证有效期), "", Format(rstemp!注册证有效期, "yyyy-mm-dd"))
                    .TextMatrix(intRow, mCol指导批发价) = Format(rstemp!指导批发价, mFMT.FM_成本价)
                    .TextMatrix(intRow, mCol原产地) = IIf(IsNull(rstemp!原产地), "!", rstemp!原产地)
                    
                    '存储格式值:最大效期||指导差价率||是否变价||在用分批||库房分批
                    .TextMatrix(intRow, mCol原销期) = IIf(IsNull(rstemp!最大效期), "0", rstemp!最大效期) & "||" & rstemp!指导差价率 & "||" & IIf(IsNull(rstemp!是否变价), 0, rstemp!是否变价) & "||" & IIf(IsNull(rstemp!在用分批), 0, rstemp!在用分批) & "||" & zlStr.Nvl(rstemp!库房分批, 0)
                    .TextMatrix(intRow, mCol简码) = ""
                    .TextMatrix(intRow, mCol比例系数) = zlStr.Nvl(rstemp!换算系数)
                    .TextMatrix(intRow, mcol注册证号) = zlStr.Nvl(rstemp!注册证号)
                    .TextMatrix(intRow, mcol商品条码) = zlStr.Nvl(rstemp!商品条码)
                    .TextMatrix(intRow, mCol序号) = zlStr.Nvl(rstemp!序号)
                    
                    .TextMatrix(intRow, mcol内部条码) = zlStr.Nvl(rstemp!内部条码)
                    .TextMatrix(intRow, mcol费用ID) = zlStr.Nvl(rstemp!费用ID)
                    
                    If (mbln退货 Or mint编辑状态 = 3) And mint编辑状态 <> 6 Then
                        dblSum = 0
                        For Each varStuff In mCllBillData
                            If varStuff(0) = CStr(rstemp!材料ID & "_" & IIf(IsNull(rstemp!批次), "0", rstemp!批次)) Then
                                dblSum = varStuff(1)
                                mCllBillData.Remove varStuff(0)
                                Exit For
                            End If
                        Next
                        str批次 = rstemp!材料ID & "_" & IIf(IsNull(rstemp!批次), "0", rstemp!批次)
                        dblSum = dblSum + Val(zlStr.Nvl(rstemp!数量)) * Val(zlStr.Nvl(rstemp!换算系数))
                        mCllBillData.Add Array(str批次, dblSum), str批次
                    End If
                    
                    If mint编辑状态 = 6 Then
                        .TextMatrix(intRow, mCol冲销数量) = Format(0, mFMT.FM_数量)
                        .RowData(intRow) = rstemp!付款序号
                        
                        '检查是否存在部分付款的情况
                        gstrSQL = "Select Nvl(Max(付款序号), 0) 付款序号 From 应付记录 " & _
                            " where 收发id=(Select Id From 药品收发记录 Where 单据=15 And No=[1] And (Mod(记录状态,3)=0 Or 记录状态=1) " & _
                            " And 序号=[2]) "
                        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "[取付款序号]", txtNO.Text, Val(.TextMatrix(intRow, mCol序号)))
                        
                        If rs!付款序号 <> 0 Then
                            mintParallelRecord = 6
                        End If
                    Else
                        .TextMatrix(intRow, mCol批次) = rstemp!批次
                    End If
                    
                    '计算加成率
                    If mbln时价购前销售 Then
                        dbl采购价 = Val(.TextMatrix(intRow, mCol采购价))
                    Else
                        dbl采购价 = Val(.TextMatrix(intRow, mCol结算价))
                    End If
                    
                    If Val(rstemp!是否变价) = 1 Then '时价
                        If Val(.TextMatrix(intRow, mCol售价)) <> 0 And dbl采购价 <> 0 Then
                            .TextMatrix(intRow, mcol加成率) = zlStr.FormatEx((Val(.TextMatrix(intRow, mCol售价)) / dbl采购价 - 1) * 100, 2) & "%"
                        End If
                    Else '定价
                        .TextMatrix(intRow, mcol加成率) = zlStr.FormatEx(Val(rstemp!加成率) * 100, 2) & "%"
                    End If
                    
                    rstemp.MoveNext
                Loop
            End With
            rstemp.Close
    End Select
    SetEdit         '设置编辑属性
    Call RefreshRowNO(mshBill, mCol行号, 1)
    Call 显示合计金额
    '高值材料VsflexGrid控件初始化
    With vsfCostlyInfo
        .Editable = flexEDKbd
        .BackColorBkg = vbWhite
        .AllowUserResizing = flexResizeColumns
        .Rows = 2
        .Cols = 6
        .RowHeight(0) = 300
        '.RowHeight(1) = 245
        .TextMatrix(0, 0) = "SN"
        .TextMatrix(0, 1) = "科室"
        .TextMatrix(0, 2) = "病人姓名"
        .TextMatrix(0, 3) = "住院号"
        .TextMatrix(0, 4) = "床号"
        .TextMatrix(0, 5) = "科室ID"
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColKey(0) = "SN"
        .ColKey(1) = "科室"
        .ColKey(2) = "病人姓名"
        .ColKey(3) = "住院号"
        .ColKey(4) = "床号"
        .ColKey(5) = "科室ID"
        .ColHidden(0) = True
        .ColHidden(5) = True
        .ColWidth(1) = 2000
        .ColWidth(2) = 2000
        .ColWidth(3) = 2000
        .ColWidth(4) = 1000
        .Visible = False
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Get摘要(ByVal strNo As String, ByVal int编辑状态 As Integer) As String
    Dim rstemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    Select Case int编辑状态
        Case 6          '冲销(取最后一次冲销的摘要)
            gstrSQL = "Select 摘要 From 药品收发记录 Where 单据=15 And No=[1] Order By 审核日期 asc "
        Case 5, 7       '修改发票、财务审核
            gstrSQL = "Select 摘要 From 药品收发记录 Where 单据 = 15 And NO = [1] And (Mod(记录状态, 3) = 0 Or 记录状态 = 1) order by 审核日期 asc"
    End Select
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "取摘要信息", strNo)
    
    If Not rstemp.EOF Then
        Get摘要 = zlStr.Nvl(rstemp!摘要)
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetEdit()
    Dim intCol As Integer
    Dim intRow As Integer
    
    With mshBill
        If mblnEdit = False Then
            
            cboStock.Enabled = False
            txtProvider.Enabled = False
            cmdProvider.Enabled = False
            txt摘要.Enabled = False
            
            For intCol = 0 To .Cols - 1
                .ColData(intCol) = 0
            Next
            
            If mint编辑状态 = 5 Then
                '修改发票信息
                mshBill.ColData(mCol发票号) = 4
                mshBill.ColData(mcol发票代码) = 4
                mshBill.ColData(mCol发票日期) = 2
                .ColData(mCol发票金额) = 4
  
                txtProvider.Enabled = True
                cmdProvider.Enabled = True
            ElseIf mint编辑状态 = 10 Then
                .ColData(mcol注册证号) = 4
            ElseIf mint编辑状态 = 6 Then
                '冲销
                For intCol = 0 To .Cols - 1
                    .ColData(intCol) = 5
                Next
                mshBill.ColData(mCol诊疗) = 0
                mshBill.ColData(mCol冲销数量) = 4
                txt摘要.Enabled = True
                
            ElseIf mint编辑状态 = 9 Then
                '核查
                '刘兴宏:增加核查人2007/05/13:10557
                '刘兴宏:20070530,增加核查流程时根据各流程设置相应的编辑项目
                Call Set操作流程Update
                txt摘要.Enabled = True
            ElseIf mint编辑状态 = 3 Then
                '审核
                '审核,要求更改成本价
                '刘兴宏:20070530,增加核查流程时根据各流程设置相应的编辑项目
                Call Set操作流程Update
                For intRow = 1 To .Rows - 1
                    If Val(.TextMatrix(intRow, mcol费用ID)) > 0 Then    '如果是备货卫材，已经在过程中处理了此处不再处理，所以不显示
                        chk转入移库.Visible = False
                        cboEnterStock.Visible = False
                        cboType.Visible = False
                        Exit For
                    End If
                Next
            End If
        Else
            '1.新增；2、修改；3、验收；4、查看；5、修改发票；6、冲销；
            '7、财务审核（冲销、产生新单据并审核；已付款的单据不允许财务审核，同样，财务审核后的单据不允许冲销）;
            '8、卫材库退货,9-核查
            If mint编辑状态 = 7 Then
            
                txtProvider.Enabled = False
                cmdProvider.Enabled = False
                txt摘要.Enabled = False
                cboStock.Enabled = False
                '刘兴宏:20070530,增加核查流程时根据各流程设置相应的编辑项目
                Call Set操作流程Update
                '                For intCol = 0 To .Cols - 1
                '                    .ColData(intCol) = 5
                '                Next
                '                .ColData(mCol结算价) = 4
                '                .ColData(mCol结算金额) = 4
                '                .LocateCol = mCol结算价
                Exit Sub
            ElseIf mint编辑状态 = 8 Or mbln退货 Then
                .ColData(mCol批号) = 5
                .ColData(mCol采购价) = 5
                .ColData(mCol效期) = 5
                .ColData(3) = 5
                .ColData(mCol结算价) = 5
                .ColData(mCol结算金额) = 5
                .ColData(mcol注册证号) = 5
                .ColData(mcol商品条码) = 5
                .ColData(mcol灭菌效期) = 5
                .ColData(mcol灭菌日期) = 5
                
                .ColData(mcol生产日期) = 5
                .ColData(mCol扣率) = 5
                
                If mbln退货 Then
                    txtProvider.Enabled = False
                    cmdProvider.Enabled = False
                End If
                '退货单不允许选择库房
                cboStock.Enabled = False
                Exit Sub
            ElseIf mint编辑状态 = 3 Then
                For intRow = 1 To .Rows - 1
                    If Val(.TextMatrix(intRow, mcol费用ID)) > 0 Then    '如果是备货卫材，已经在过程中处理了此处不再处理，所以不显示
                        chk转入移库.Visible = False
                        cboEnterStock.Visible = False
                        cboType.Visible = False
                        Exit For
                    End If
                Next
            End If
            .ColData(0) = 5
            .ColData(mCol行号) = 5
            .ColData(mCol诊疗) = 1
            .ColData(mCol序号) = 5
            .ColData(mCol规格) = 5
            .ColData(mCol产地) = 5
            .ColData(mCol单位) = 5
            .ColData(mCol批号) = 4
            .ColData(mcol生产日期) = 2
            .ColData(mcol注册证号) = 4
            .ColData(mCol效期) = 5
            .ColData(mCol数量) = 4
            
            .ColData(mCol售价) = 5
            .ColData(mCol售价金额) = 5
            .ColData(mCol差价) = 5
            '刘兴宏:零售价处理
            .ColData(mcol零售价) = 5
            .ColData(mcol零售金额) = 5
            .ColData(mcol零售差价) = 5
 
            .ColData(mCol验收结论) = IIf(mint编辑状态 = 1 Or mint编辑状态 = 2, 1, 5)
            .ColData(mCol随货单号) = 4
            .ColData(mCol发票号) = 4
            .ColData(mcol发票代码) = 4
            .ColData(mCol发票日期) = 2
            
            If mbln不强制控制指导价格 Then
                .ColData(mCol指导批发价) = 5
            Else
                .ColData(mCol指导批发价) = IIf(mbln修改批发价, 4, 5)
            End If
            
            .ColData(mCol原产地) = 5
            .ColData(mCol原销期) = 5
            .ColData(mCol简码) = 5
            .ColData(mCol比例系数) = 5
            
            .ColData(mcol一次性材料) = 5
            .ColData(mcol条码管理) = 5
            .ColData(mcol灭菌效期) = 5
            .ColData(mcol灭菌日期) = 2
            .ColData(mcol灭菌失效期) = 5
            .ColData(mcol注册证号) = 4
            .ColData(mcol注册证有效期) = 5

            .ColData(mCol采购价) = 4
            .ColData(mCol结算价) = 4
            .ColData(mCol结算金额) = 4
            
            .ColData(mCol采购价) = 4
            .ColData(mCol扣率) = 4
            .ColData(mCol发票金额) = 4
                  
            .ColAlignment(mCol诊疗) = flexAlignLeftCenter
            .ColAlignment(mCol规格) = flexAlignLeftCenter
            .ColAlignment(mCol产地) = flexAlignLeftCenter
            .ColAlignment(mCol单位) = flexAlignCenterCenter
            .ColAlignment(mCol批号) = flexAlignLeftCenter
            .ColAlignment(mCol效期) = flexAlignLeftCenter
            .ColAlignment(mCol数量) = flexAlignRightCenter
            .ColAlignment(mCol采购价) = flexAlignRightCenter
            .ColAlignment(mCol结算价) = flexAlignRightCenter
            .ColAlignment(mCol结算金额) = flexAlignRightCenter
            .ColAlignment(mCol售价) = flexAlignRightCenter
            .ColAlignment(mCol售价金额) = flexAlignRightCenter
            .ColAlignment(mCol差价) = flexAlignRightCenter

            
            '刘兴宏:零售价处理
            .ColAlignment(mcol零售价) = flexAlignRightCenter
            .ColAlignment(mcol零售金额) = flexAlignRightCenter
            .ColAlignment(mcol零售差价) = flexAlignRightCenter
            .ColAlignment(mcol零售单位) = flexAlignCenterCenter
            
            .ColAlignment(mCol扣率) = flexAlignRightCenter
            .ColAlignment(mCol随货单号) = flexAlignLeftCenter
            .ColAlignment(mCol发票号) = flexAlignLeftCenter
            .ColAlignment(mcol发票代码) = flexAlignLeftCenter
            .ColAlignment(mCol发票日期) = flexAlignLeftCenter
            .ColAlignment(mCol发票金额) = flexAlignRightCenter
            
            .ColAlignment(mcol灭菌日期) = flexAlignCenterCenter
            .ColAlignment(mcol灭菌失效期) = flexAlignCenterCenter
            .ColAlignment(mcol注册证号) = flexAlignLeftCenter
            .ColAlignment(mcol商品条码) = flexAlignLeftCenter
            
            cboStock.Enabled = True
                        
            txtProvider.Enabled = True
            cmdProvider.Enabled = True
            txt摘要.Enabled = True
        End If
    End With
End Sub


Private Sub initGrid()
    '表格初始化、初始化摘要文本框的长度
    On Error GoTo ErrHandle
    Dim intCol As Integer
    
    With mshBill
        .Active = True
        .Cols = mCols
        .Value = Format(sys.Currentdate, "yyyy-mm-dd")
        .MsfObj.FixedCols = 1
        Call SetColumnByUserDefine
        .TextMatrix(0, mCol行号) = ""
        .TextMatrix(0, mCol诊疗) = "名称与编码"
        .TextMatrix(0, mCol序号) = "序号"
        .TextMatrix(0, mCol商品名) = "商品名"
        .TextMatrix(0, mCol规格) = "规格"
        .TextMatrix(0, mCol产地) = "产地"
        .TextMatrix(0, mcol批准文号) = "批准文号"
        .TextMatrix(0, mCol单位) = "单位"
        .TextMatrix(0, mCol批号) = "批号"
        .TextMatrix(0, mcol生产日期) = "生产日期"
        .TextMatrix(0, mCol效期) = "失效期"
        .TextMatrix(0, mCol数量) = "数量"
        .TextMatrix(0, mCol冲销数量) = "冲销数量"
        .TextMatrix(0, mCol批次) = "批次"
        
        .TextMatrix(0, mCol采购价) = "采购价"
        .TextMatrix(0, mCol结算价) = "结算价"
        .TextMatrix(0, mCol结算金额) = "结算金额"
        .TextMatrix(0, mCol售价) = "售价"
        .TextMatrix(0, mCol售价金额) = "售价金额"
        .TextMatrix(0, mCol差价) = "差价"
        
        .TextMatrix(0, mcol零售价) = "零售价"
        .TextMatrix(0, mcol零售单位) = "零售单位"
        .TextMatrix(0, mcol零售金额) = "零售金额"
        .TextMatrix(0, mcol零售差价) = "零售差价"
        
        .TextMatrix(0, mCol扣率) = "扣率"
        .TextMatrix(0, mcol加成率) = "加成率"
        
        .TextMatrix(0, mCol验收结论) = "验收结论"
        .TextMatrix(0, mCol随货单号) = "随货单号"
        .TextMatrix(0, mCol发票号) = "发票号"
        .TextMatrix(0, mcol发票代码) = "发票代码"
        .TextMatrix(0, mCol发票日期) = "发票日期"
        .TextMatrix(0, mCol发票金额) = "发票金额"
        .TextMatrix(0, mCol指导批发价) = "采购限价"
        .TextMatrix(0, mCol原产地) = "原产地"
        .TextMatrix(0, mCol原销期) = "原效期"
        .TextMatrix(0, mCol简码) = "简码"
        .TextMatrix(0, mCol比例系数) = "比例系数"
        
        .TextMatrix(0, mcol一次性材料) = "一次性材料"
        .TextMatrix(0, mcol条码管理) = "条码管理"
        .TextMatrix(0, mcol灭菌效期) = "灭菌效期"
        .TextMatrix(0, mcol灭菌日期) = "灭菌日期"
        .TextMatrix(0, mcol灭菌失效期) = "灭菌失效期"
        .TextMatrix(0, mcol注册证号) = "注册证号"
        .TextMatrix(0, mcol商品条码) = "商品条码"
        .TextMatrix(0, mcol内部条码) = "内部条码"
        .TextMatrix(0, mcol费用ID) = "费用ID"
        .TextMatrix(0, mcol注册证有效期) = "注册证有效期"

        .TextMatrix(1, 0) = ""
        .TextMatrix(1, mCol行号) = "1"
        
        .ColWidth(0) = 0
        
        .ColWidth(mCol行号) = 300
        .ColWidth(mCol诊疗) = 2000
        .ColWidth(mCol序号) = 0
        .ColWidth(mCol商品名) = 900
        .ColWidth(mCol规格) = 900
        .ColWidth(mCol产地) = 800
        .ColWidth(mcol批准文号) = 1000
        .ColWidth(mCol单位) = 500
        .ColWidth(mCol批号) = 800
        .ColWidth(mcol生产日期) = 1000
        .ColWidth(mCol效期) = 1000
        .ColWidth(mCol数量) = 800
        .ColWidth(mCol冲销数量) = IIf(mint编辑状态 = 6, 800, 0)
        .ColWidth(mCol批次) = 0
        
        .ColWidth(mCol结算价) = IIf(mblnCostView = False, 0, 900)
        .ColWidth(mCol采购价) = IIf(mblnCostView = False, 0, 900)
        
        .ColWidth(mCol结算金额) = IIf(mblnCostView = False, 0, 900)
        .ColWidth(mCol售价) = 900
        .ColWidth(mCol售价金额) = 900
        .ColWidth(mCol差价) = IIf(mblnCostView = False, 0, 900)
            
        '刘兴宏:零售价处理
        .ColWidth(mcol零售价) = 900
        .ColWidth(mcol零售金额) = 900
        .ColWidth(mcol零售差价) = 800
        .ColWidth(mcol零售单位) = 800
        If mbln不强制控制指导价格 Then
            .ColWidth(mCol指导批发价) = 0
        Else
            .ColWidth(mCol指导批发价) = 900
        End If
        
        .ColWidth(mCol扣率) = 800
        
        If mbln分段加成率 = True Then
            .ColWidth(mcol加成率) = 0
        Else
            .ColWidth(mcol加成率) = 1200
        End If
        .ColWidth(mCol验收结论) = 4500
        .ColWidth(mCol随货单号) = 800
        .ColWidth(mCol发票号) = 800
        .ColWidth(mcol发票代码) = 1000
        .ColWidth(mCol发票日期) = 1000
        .ColWidth(mCol发票金额) = 900
        .ColWidth(mCol原产地) = 0
        .ColWidth(mCol原销期) = 0
        .ColWidth(mCol简码) = 0
        .ColWidth(mCol比例系数) = 0
        .ColWidth(mcol一次性材料) = 0
        .ColWidth(mcol条码管理) = 0
        .ColWidth(mcol灭菌效期) = 0
        .ColWidth(mcol灭菌日期) = 1200
        .ColWidth(mcol灭菌失效期) = 1200
        .ColWidth(mcol注册证号) = 1600
        .ColWidth(mcol商品条码) = IIf(gblnCode = True, 2000, 0)
        .ColWidth(mcol内部条码) = 0
        .ColWidth(mcol费用ID) = 0
        .ColWidth(mcol注册证有效期) = 0
        
        '-1：表示该列可以选择，是布尔型［"√"，" "］
        ' 0：表示该列可以选择，但不能修改
        ' 1：表示该列可以输入，外部显示为按钮选择
        ' 2：表示该列是日期列，外部显示为按钮选择，弹出是日期选择框
        ' 3：表示该列是选择列，外部显示为下拉框选择
        '4:  表示该列为单纯的文本框供用户输入
        '5:  表示该列不允许选择

        .ColData(0) = 5
        .ColData(mCol行号) = 5
        .ColData(mCol诊疗) = 1
        .ColData(mCol序号) = 5
        .ColData(mCol商品名) = 5
        .ColData(mCol规格) = 5
        .ColData(mCol产地) = 5
        .ColData(mCol单位) = 5
        .ColData(mCol批号) = 4
        .ColData(mcol生产日期) = 2
        .ColData(mcol批准文号) = 4
        .ColData(mCol效期) = 5
        .ColData(mCol数量) = 4
        .ColData(mCol冲销数量) = 5
        .ColData(mCol批次) = 5
        .ColData(mcol加成率) = 5
        .ColData(mCol售价) = 5
        .ColData(mCol售价金额) = 5
        .ColData(mCol差价) = 5
        '刘兴宏:零售价处理
        .ColData(mcol零售单位) = 5
        .ColData(mcol零售价) = IIf(mbln退货 Or mint编辑状态 = 8, 5, 4)
        .ColData(mcol零售金额) = 5
        .ColData(mcol零售差价) = 5
        
        .ColData(mCol验收结论) = IIf(mint编辑状态 = 1 Or mint编辑状态 = 2, 1, 5)
        .ColData(mCol发票号) = 4
        .ColData(mcol发票代码) = 5
        .ColData(mCol随货单号) = 4
        .ColData(mCol发票日期) = 2
        
        If mbln不强制控制指导价格 Then
            .ColData(mCol指导批发价) = 5
        Else
            .ColData(mCol指导批发价) = IIf(mbln修改批发价, 4, 5)
        End If
        .ColData(mCol原产地) = 5
        .ColData(mCol原销期) = 5
        .ColData(mCol简码) = 5
        .ColData(mCol比例系数) = 5
        

        .ColData(mcol一次性材料) = 5
        .ColData(mcol条码管理) = 5
        .ColData(mcol灭菌效期) = 5
        .ColData(mcol灭菌日期) = 2
        .ColData(mcol灭菌失效期) = 5
        .ColData(mcol注册证号) = 5
        .ColData(mcol商品条码) = 4
        .ColData(mcol内部条码) = 5
        .ColData(mcol费用ID) = 5
        .ColData(mcol注册证有效期) = 5

         .ColData(mCol结算价) = 4
         .ColData(mCol结算金额) = 4
        
        .ColData(mCol扣率) = 4
        .ColData(mCol采购价) = 4
        .ColData(mCol发票金额) = 4
        
        .ColAlignment(mCol诊疗) = flexAlignLeftCenter
        .ColAlignment(mCol商品名) = flexAlignLeftCenter
        .ColAlignment(mCol规格) = flexAlignLeftCenter
        .ColAlignment(mCol产地) = flexAlignLeftCenter
        .ColAlignment(mCol验收结论) = flexAlignLeftCenter
        .ColAlignment(mCol单位) = flexAlignCenterCenter
        .ColAlignment(mCol批号) = flexAlignLeftCenter
        .ColAlignment(mCol效期) = flexAlignLeftCenter
        .ColAlignment(mCol数量) = flexAlignRightCenter
        .ColAlignment(mCol冲销数量) = flexAlignRightCenter
        .ColAlignment(mCol结算价) = flexAlignRightCenter
        .ColAlignment(mCol结算金额) = flexAlignRightCenter
        .ColAlignment(mCol售价) = flexAlignRightCenter
        .ColAlignment(mCol售价金额) = flexAlignRightCenter
        .ColAlignment(mCol差价) = flexAlignRightCenter
        '刘兴宏:零售价处理
        .ColAlignment(mcol零售单位) = flexAlignCenterCenter
        .ColAlignment(mcol零售价) = flexAlignRightCenter
        .ColAlignment(mcol零售金额) = flexAlignRightCenter
        .ColAlignment(mcol零售差价) = flexAlignRightCenter
        
        .ColAlignment(mCol扣率) = flexAlignRightCenter
        .ColAlignment(mCol发票号) = flexAlignLeftCenter
        .ColAlignment(mcol发票代码) = flexAlignLeftCenter
        .ColAlignment(mCol发票日期) = flexAlignLeftCenter
        .ColAlignment(mCol发票金额) = flexAlignRightCenter
                 

        .ColAlignment(mcol灭菌日期) = flexAlignLeftCenter
        .ColAlignment(mcol生产日期) = flexAlignLeftCenter
        .ColAlignment(mcol灭菌失效期) = flexAlignLeftCenter
        .ColAlignment(mcol注册证号) = flexAlignLeftCenter
        .ColAlignment(mcol商品条码) = flexAlignLeftCenter
        
        .PrimaryCol = mCol诊疗
        .LocateCol = mCol诊疗
        Call SetColumnByUserDefine
        '刘兴宏:如果是最小单位,则要屏蔽列:
        If mintUnit = 0 Then
            .ColWidth(mcol零售单位) = 0
            .ColWidth(mcol零售价) = 0
            .ColWidth(mcol零售金额) = 0
            .ColWidth(mcol零售差价) = 0
            .ColWidth(mcol零售单位) = 5
            .ColWidth(mcol零售价) = 5
            .ColWidth(mcol零售金额) = 5
            .ColWidth(mcol零售差价) = 5
        End If
    End With
    
    
    txt摘要.MaxLength = sys.FieldsLength("药品收发记录", "摘要")
    
    '高值材料
    Select Case mint编辑状态
    Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10
        '1.新增; 2.编辑; 3.审核; 4:查询; 6:冲销; 7:财务审核; 8:退货; 9:核查; 10:修改注册证号
        Set mrsCostlyInfo = New ADODB.Recordset
        With mrsCostlyInfo
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .CursorType = adOpenStatic
            .Fields.Append "SN", adInteger, , adFldIsNullable
            .Fields.Append "id", adInteger, , adFldIsNullable
            .Fields.Append "科室", adLongVarChar, 100, adFldIsNullable
            .Fields.Append "病人姓名", adVarChar, 64, adFldIsNullable
            .Fields.Append "住院号", adVarChar, 20, adFldIsNullable
            .Fields.Append "床号", adVarChar, 10, adFldIsNullable
            .Open
        End With
        If mint编辑状态 <> 1 Then
            Dim rsTmp As ADODB.Recordset
            Dim strTmp As String
            
            strTmp = "select a.序号 SN, c.id 科室id, b.科室, b.病人姓名, b.住院号, b.床号 " _
                   & "from 药品收发记录 a, 收发记录补充信息 b, 部门表 c " _
                   & "where a.id=b.收发id and b.科室=c.名称(+) and a.no=[1] " _
                   & "order by a.序号"
            Set rsTmp = zlDatabase.OpenSQLRecord(strTmp, mstrCaption, mstr单据号)
            Do While Not rsTmp.EOF
                mrsCostlyInfo.AddNew
                mrsCostlyInfo!sn = rsTmp!sn
                mrsCostlyInfo!Id = rsTmp!科室id
                mrsCostlyInfo!科室 = rsTmp!科室
                mrsCostlyInfo!病人姓名 = rsTmp!病人姓名
                mrsCostlyInfo!住院号 = rsTmp!住院号
                mrsCostlyInfo!床号 = rsTmp!床号
                mrsCostlyInfo.Update
                rsTmp.MoveNext
            Loop
            rsTmp.Close
        End If
       
    End Select
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    
    With Pic单据
        
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - .Top - 100 - CmdCancel.Height - 200

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
        .Top = Lbl填制日期.Top - 80
        .Left = Lbl填制日期.Left + Lbl填制日期.Width + 100
    End With
    
    With Lbl填制人
        .Top = Lbl填制日期.Top - .Height - 140
        .Left = mshBill.Left + 100
    End With
    
    With Txt填制人
        .Top = Lbl填制人.Top - 80
        .Left = Lbl填制人.Left + Lbl填制人.Width + 100
    End With
    
    With lbl核查人
        .Top = Lbl填制人.Top
        .Left = Abs(mshBill.Width - .Width - txt核查人.Width - 100) / 2
    End With
    With txt核查人
        .Top = lbl核查人.Top - 80
        .Left = lbl核查人.Left + lbl核查人.Width + 100
    End With
    
    With lbl核查日期
        .Top = Lbl填制日期.Top
        .Left = lbl核查人.Left
    End With
    With txt核查日期
        .Top = Txt填制日期.Top
        .Left = txt核查人.Left
    End With
    
    
    With Txt审核日期
        .Top = Lbl填制日期.Top - 80
        .Left = mshBill.Left + mshBill.Width - .Width
    End With
    
    With Lbl审核日期
        .Top = Lbl填制日期.Top
        .Left = Txt审核日期.Left - 100 - .Width
    End With
    
    With Txt审核人
        .Top = Lbl填制人.Top - 80
        .Left = mshBill.Left + mshBill.Width - .Width
    End With
    
    With lbl审核人
        .Top = Lbl填制人.Top
        .Left = Txt审核人.Left - 100 - .Width
    End With
    
    With txt摘要
        .Top = Lbl填制人.Top - 140 - .Height
        .Left = Txt填制人.Left
        .Width = mshBill.Left + mshBill.Width - .Left
    End With
    
    With lbl摘要
        .Top = txt摘要.Top + 50
        .Left = txt摘要.Left - .Width - 100
    End With
    
    With lblPurchasePrice
        .Left = mshBill.Left
        .Top = txt摘要.Top - 60 - .Height
        .Width = mshBill.Width
        lblSalePrice.Top = .Top
        lblDifference.Top = .Top
    End With
    If mblnCostView = False Then
        lblPurchasePrice.Visible = False
    End If
    
    With lblSalePrice
        .Left = lblPurchasePrice.Left + mshBill.Width / 3
    End With
    
    With lblDifference
        .Left = lblPurchasePrice.Left + mshBill.Width / 3 * 2
    End With
    If mblnCostView = False Then
        lblDifference.Visible = False
    End If
    
    With mshBill
        '高值材料
        picCostly.Visible = vsfCostlyInfo.Visible
        If vsfCostlyInfo.Visible Then
            picCostly.Height = 400
            picCostly.Left = .Left
            picCostly.Width = .Width
            vsfCostlyInfo.Height = 650
            vsfCostlyInfo.Left = .Left
            vsfCostlyInfo.Width = .Width
            .Height = lblPurchasePrice.Top - .Top - 60 - vsfCostlyInfo.Height - picCostly.Height
            picCostly.Top = .Top + .Height + 40
            vsfCostlyInfo.Top = picCostly.Top + picCostly.Height + 10
        Else
            .Height = lblPurchasePrice.Top - .Top - 60
        End If
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
        
    With cmdExtractData
        .Top = CmdCancel.Top
    End With
    
    With cmdFind
        .Top = CmdCancel.Top
    End With
    
    If mint编辑状态 = 5 Then '修改发票信息该按钮才可用
        With cmdBulkCopy
            cmdBulkCopy.Visible = True
            .Top = cmdFind.Top
            If txtCode.Visible Then
                .Left = txtCode.Left + txtCode.Width + 100
            Else
                .Left = cmdFind.Left + cmdFind.Width + 100
            End If
        End With
        
        With cmdALLDel
            .Visible = True
            .Left = cmdBulkCopy.Left + cmdBulkCopy.Width + 100
            .Top = cmdBulkCopy.Top
        End With
    End If
    
    With lblCode
        .Top = CmdCancel.Top + 50
    End With
    With txtCode
        .Top = CmdCancel.Top + 30
    End With
    Me.fraMoveNO.Top = txtCode.Top
    Me.fraMoveNO.Left = txtCode.Left + txtCode.Width + 50
    
    With cmdCopy
        .Left = IIf(txtCode.Visible, txtCode.Left + txtCode.Width, cmdFind.Left + cmdFind.Width) + 100
        .Top = cmdFind.Top
    End With
    
    With txtCopy
        .Left = cmdCopy.Left + cmdCopy.Width + 50
        .Top = txtCode.Top
    End With
    
    With lblCopy
        .Left = txtCopy.Left + txtCopy.Width + 25
        .Top = txtCopy.Top + 45
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mshProvider.Visible = True Then
        mshProvider.Visible = False
        txtProvider.SetFocus
        txtProvider.SelLength = Len(txtProvider.Text)
        txtProvider.SelStart = 0
        Cancel = True
        Exit Sub
    End If
    
    If msh产地.Visible = True Then
        msh产地.Visible = False
        mshBill.SetFocus
        mshBill.Col = mCol产地
        Cancel = True
        Exit Sub
    End If
    
    If mblnChange = False Or mint编辑状态 = 4 Or mint编辑状态 = 3 Then
        SaveWinState Me, App.ProductName, mstrCaption
        Call SaveBILLWidth
        Call zlDatabase.SetPara("审核产生单据", cboType.ItemData(cboType.ListIndex), glngSys, mlngModule)
        Exit Sub
    End If
    
    If MsgBox("数据可能已改变，但未存盘，真要退出吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
        Exit Sub
    Else
        SaveWinState Me, App.ProductName, mstrCaption
    End If
    Call SaveBILLWidth
    Call zlDatabase.SetPara("审核产生单据", cboType.ItemData(cboType.ListIndex), glngSys, mlngModule)
    '高值材料
    If mrsCostlyInfo Is Nothing Then Exit Sub
    If mrsCostlyInfo.State = adStateOpen Then mrsCostlyInfo.Close
End Sub

Private Function SaveCheck(Optional ByVal strNo As String = "") As Boolean
    mblnSave = False
    SaveCheck = False
    
    gstrSQL = "zl_材料外购_Verify('" & IIf(mint编辑状态 = 7, strNo, txtNO.Tag) & "','" & UserInfo.用户名 & "',to_date('" & Format(mstr审核日期, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS'))"
    
    On Error GoTo ErrHandle
    'If mint编辑状态 <> 7 Then gcnOracle.BeginTrans
    Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
   ' If mint编辑状态 <> 7 Then gcnOracle.CommitTrans
    
    SaveCheck = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    Exit Function
ErrHandle:
    'If mint编辑状态 <> 7 Then gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Function
 




Private Sub lblType_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Me.PopupMenu mnuSearch
End Sub

Private Sub mnuSearch01_Click()
    lblType.Caption = "病人ID↓"
    If lblType.Tag <> 1 Then txtTypeVar.Text = ""
    lblType.Tag = 1
End Sub

Private Sub mnuSearch02_Click()
    lblType.Caption = "病人姓名↓"
    If lblType.Tag <> 2 Then txtTypeVar.Text = ""
    lblType.Tag = 2
End Sub

Private Sub mnuSearch03_Click()
    lblType.Caption = "住院号↓"
    If lblType.Tag <> 3 Then txtTypeVar.Text = ""
    lblType.Tag = 3
End Sub

Private Sub mnuSearch04_Click()
    lblType.Caption = "门诊号↓"
    If lblType.Tag <> 4 Then txtTypeVar.Text = ""
    lblType.Tag = 4
End Sub

Private Sub mnuSearch05_Click()
    lblType.Caption = "床号↓"
    If lblType.Tag <> 5 Then txtTypeVar.Text = ""
    lblType.Tag = 5
End Sub

Private Sub mshBill_AfterAddRow(Row As Long)
    Call RefreshRowNO(mshBill, mCol行号, Row)
End Sub

Private Sub mshBill_AfterDeleteRow()
    Call 显示合计金额
    Call RefreshRowNO(mshBill, mCol行号, mshBill.Row)
    If mshBill.TextMatrix(mshBill.Row, 0) <> "" Then
        CostlyInfo_Refresh Val(mshBill.TextMatrix(mshBill.Row, 1)), IsCostly(mshBill.TextMatrix(mshBill.Row, 0))
    End If
End Sub

Private Sub mshBill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    If InStr(1, ",3,4,5,6,7,9,10,", "," & mint编辑状态 & ",") <> 0 Then
        Cancel = True
        Exit Sub
    End If
    With mshBill
        If .TextMatrix(.Row, 0) <> "" Then
            If MsgBox("你确实要删除该行卫生材料？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True
                Exit Sub
            End If
            CostlyInfo_Refresh Val(mshBill.TextMatrix(mshBill.Row, 1)), False
            '调整对应高值材料的SN
            RecountSN mshBill.Row
        End If
    End With
End Sub
Private Sub ColMoveNextCol(ByVal lngCol As Long)
    '------------------------------------------------------------------------------
    '功能:列移动
    '返回:
    '编制:刘兴宏
    '日期:2007/08/14
    '------------------------------------------------------------------------------
    Dim i As Long
    With mshBill
        For i = lngCol + 1 To .Cols - 1
            Select Case .ColData(i)
            Case -1, 1, 2, 3, 4
                '-1：表示该列可以选择，是布尔型［"√"，" "］
                ' 0：表示该列可以选择，但不能修改
                ' 1：表示该列可以输入，外部显示为按钮选择
                ' 2：表示该列是日期列，外部显示为按钮选择，弹出是日期选择框
                ' 3：表示该列是选择列，外部显示为下拉框选择
                '4:  表示该列为单纯的文本框供用户输入
                '5:  表示该列不允许选择
                .Col = i
                Exit For
            End Select
        Next
        
        If i - 1 = .Cols - 1 Then
            If .Row = .Rows - 1 Then
                .Rows = .Rows + 1
            End If
            .Row = .Row + 1
            .Col = mCol诊疗
        End If
    End With
    
End Sub

Private Sub mshbill_CommandClick()
    Dim i As Integer
    Dim int点击行 As Integer
    Dim rs验收结论 As Recordset
    
    On Error GoTo ErrHandle
    
    int点击行 = mshBill.Row

     If mshBill.Col = mCol诊疗 Then
        Dim mrsReturn As Recordset
        If mint编辑状态 = 8 Or mbln退货 Then
            If Val(txtProvider.Tag) = 0 Then
                ShowMsgBox "未选择退货单位!"
                If txtProvider.Enabled Then txtProvider.SetFocus
                Exit Sub
            End If
        End If
        Set mrsReturn = Frm材料选择器.ShowMe(Me, IIf(mint编辑状态 = 8 Or mbln退货, 2, 1), cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , True, True, False, False, True, IIf(mint编辑状态 = 8 Or mbln退货, Val(txtProvider.Tag), 0), IIf(mintUnit = 0, True, False), , , , 1712, , mstrPrivs, , False)
        
        If mrsReturn.RecordCount > 0 And (mint编辑状态 = 8 Or mbln退货 = True) Then
            Set mrsReturn = CheckRedo(mrsReturn) '检查重复记录，将重复的记录过滤掉然后返回过滤后的数据集
        End If
        
        If mrsReturn.RecordCount > 0 Then
            With mshBill
                mrsReturn.MoveFirst
                For i = 1 To mrsReturn.RecordCount
                    If CheckQualifications(mlngModule, 0, Val(mrsReturn!材料ID)) = False Then Exit Sub
                    
                    SetColValue .Row, mrsReturn!材料ID, "[" & mrsReturn!编码 & "]" & mrsReturn!名称, IIf(IsNull(mrsReturn!规格), "", mrsReturn!规格), _
                        IIf(IsNull(mrsReturn!产地), "", mrsReturn!产地), _
                        IIf(mintUnit = 0, mrsReturn!散装单位, mrsReturn!包装单位), _
                        IIf(IsNull(mrsReturn!售价), 0, mrsReturn!售价) * IIf(mintUnit = 0, 1, mrsReturn!换算系数), _
                         mrsReturn!指导批发价 * IIf(mintUnit = 0, 1, mrsReturn!换算系数), _
                        IIf(IsNull(mrsReturn!产地), "!", mrsReturn!产地), mrsReturn!最大效期, "", _
                        IIf(mintUnit = 0, 1, mrsReturn!换算系数), IIf(IsNull(mrsReturn!批次), 0, mrsReturn!批次), mrsReturn!时价, _
                        mrsReturn!在用分批, mrsReturn!指导差价率 / 100, IIf(IsNull(mrsReturn!批准文号), "", mrsReturn!批准文号), IIf(IsNull(mrsReturn!商品名), "", mrsReturn!商品名)
                    
                    Call ColMoveNextCol(.Col)
                    
                    '高值材料新增
                    If .TextMatrix(.Row, 0) = "" Then
                        vsfCostlyInfo.Visible = False
                    Else
                        vsfCostlyInfo.Visible = IsCostly(.TextMatrix(.Row, 0))
                    End If
                    CostlyInfo_Refresh Val(.TextMatrix(.Row, 1)), vsfCostlyInfo.Visible
                    Call Form_Resize
                    
                    mblnChange = True

                    If .Row = .Rows - 1 Then .Rows = .Rows + 1 '只有当前行是最后一行时才新增行
                    .Row = .Row + 1
                    
                    mrsReturn.MoveNext
                Next
                
                mshBill.Row = int点击行
            
                
'                If mrsReturn.RecordCount = 1 Then
'                    If CheckQualifications(mlngModule, 0, Val(mrsReturn!材料ID)) = False Then Exit Sub
'
'                    SetColValue .Row, mrsReturn!材料ID, "[" & mrsReturn!编码 & "]" & mrsReturn!名称, IIf(IsNull(mrsReturn!规格), "", mrsReturn!规格), _
'                        IIf(IsNull(mrsReturn!产地), "", mrsReturn!产地), _
'                        IIf(mintUnit = 0, mrsReturn!散装单位, mrsReturn!包装单位), _
'                        IIf(IsNull(mrsReturn!售价), 0, mrsReturn!售价) * IIf(mintUnit = 0, 1, mrsReturn!换算系数), _
'                         mrsReturn!指导批发价 * IIf(mintUnit = 0, 1, mrsReturn!换算系数), _
'                        IIf(IsNull(mrsReturn!产地), "!", mrsReturn!产地), mrsReturn!最大效期, "", _
'                        IIf(mintUnit = 0, 1, mrsReturn!换算系数), IIf(IsNull(mrsReturn!批次), 0, mrsReturn!批次), mrsReturn!时价, _
'                        mrsReturn!在用分批, mrsReturn!指导差价率 / 100, IIf(IsNull(mrsReturn!批准文号), "", mrsReturn!批准文号)
'
'                    Call ColMoveNextCol(.Col)
'
'                    '高值材料新增
'                    If .TextMatrix(.Row, 0) = "" Then
'                        vsfCostlyInfo.Visible = False
'                    Else
'                        vsfCostlyInfo.Visible = IsCostly(.TextMatrix(.Row, 0))
'                    End If
'                    CostlyInfo_Refresh Val(.TextMatrix(.Row, 1)), vsfCostlyInfo.Visible
'                    Call Form_Resize
'
'                    mblnChange = True
'                End If
            End With
            mrsReturn.Close
        End If
    ElseIf mshBill.Col = mCol验收结论 Then
        gstrSQL = "Select 编码 as id,null as 上级id,编码,名称,1 as 末级 From 入库验收结论 Order By 编码 "
        Set rs验收结论 = zlDatabase.ShowSelect(Me, gstrSQL, 1, "入库验收结论", True, , "选择入库验收结论")
        If rs验收结论 Is Nothing Then Exit Sub
        If rs验收结论.State <> 1 Then Exit Sub
        
        With rs验收结论
            mshBill.TextMatrix(mshBill.Row, mCol验收结论) = zlStr.Nvl(!名称)
        End With
        
        Call ColMoveNextCol(mshBill.Col)
    Else
        Dim rstemp As New Recordset
        
        gstrSQL = "Select rownum as id,null as 上级id,编码,名称,简码,1 as 末级 From 材料生产商 "
        Set rstemp = zlDatabase.ShowSelect(Me, gstrSQL, 1, "材料生产商选择", True, , "选择卫生材料生产商或厂牌")
        
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
        If rstemp Is Nothing Then Exit Sub
        If rstemp.State <> 1 Then Exit Sub
        
        With rstemp
            If CheckQualifications(mlngModule, 1, CStr(zlStr.Nvl(!名称))) = False Then Exit Sub
            mshBill.TextMatrix(mshBill.Row, mCol产地) = zlStr.Nvl(!名称)
        End With
        
        gstrSQL = "select 批准文号 from 药品生产商对照 where 厂家名称=[1] and 药品id=[2]"
        Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "mshbill_CommandClick", mshBill.TextMatrix(mshBill.Row, mCol产地), mshBill.TextMatrix(mshBill.Row, 0))
        If rstemp.RecordCount > 0 Then
            mshBill.TextMatrix(mshBill.Row, mcol批准文号) = IIf(IsNull(rstemp!批准文号), "", rstemp!批准文号)
        Else
            mshBill.TextMatrix(mshBill.Row, mcol批准文号) = ""
        End If
        Call ColMoveNextCol(mshBill.Col)
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub mshbill_EditChange(curText As String)
    mblnChange = True
End Sub
Private Function 计算零售价及零售差价(ByVal lngRow As Long, Optional bln零售价 As Boolean = True) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:根据库房单位计算散装单位的零售价及金额
    '入参:lngRow -指定计算的行
    '     bln零售价-零售价为售价
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-11-28 12:09:04
    '-----------------------------------------------------------------------------------------------------------
    Dim dbl比例系数 As Double, arrSplit As Variant, dbl数量 As Double
    
    
    With mshBill
    
        dbl比例系数 = Val(.TextMatrix(lngRow, mCol比例系数))
        
        dbl数量 = IIf(mint编辑状态 = 6, IIf(Val(.TextMatrix(lngRow, mCol冲销数量)) = 0, Val(.Text), Val(.TextMatrix(lngRow, mCol冲销数量))), Val(.TextMatrix(lngRow, mCol数量)))
        If dbl数量 = 0 Or Val(.TextMatrix(lngRow, 0)) = 0 Then
            .TextMatrix(lngRow, mcol零售金额) = 0
            .TextMatrix(lngRow, mcol零售差价) = 0
            .TextMatrix(lngRow, mCol差价) = 0
            .TextMatrix(lngRow, mCol售价金额) = 0
            .TextMatrix(lngRow, mcol零售价) = Format(Val(.TextMatrix(lngRow, mCol售价)) / IIf(dbl比例系数 = 0, 1, dbl比例系数), mFMT.FM_散装零售价)
            Exit Function
        End If
        
        '存储格式:最大效期||指导差价率||是否变价||在用分批||库房分批
        If .TextMatrix(lngRow, mCol原销期) <> "" Then
           arrSplit = Split(.TextMatrix(lngRow, mCol原销期), "||")
           If Val(arrSplit(2)) = 1 And (IIf(mbln库房, arrSplit(4) = 1, arrSplit(3) = 1)) Then
                '实价卫材
                '刘兴宏:零售价处理
                If bln零售价 Then
                    .TextMatrix(lngRow, mcol零售价) = Format(Val(.TextMatrix(lngRow, mCol售价)) / dbl比例系数, mFMT.FM_散装零售价)
                End If
                .TextMatrix(lngRow, mcol零售金额) = Format(Val(.TextMatrix(lngRow, mcol零售价)) * (dbl数量 * dbl比例系数), mFMT.FM_金额)
                '零售差价=零售金额-结算金额
                .TextMatrix(lngRow, mcol零售差价) = Format(Val(.TextMatrix(lngRow, mcol零售金额)) - Val(.TextMatrix(lngRow, mCol结算金额)), mFMT.FM_金额)
           Else '定价
                '刘兴宏:零售价处理
                .TextMatrix(lngRow, mcol零售价) = Format(Val(.TextMatrix(lngRow, mCol售价)) / dbl比例系数, mFMT.FM_散装零售价)
                .TextMatrix(lngRow, mcol零售金额) = Format(Val(.TextMatrix(lngRow, mcol零售价)) * (dbl数量 * dbl比例系数), mFMT.FM_金额)
                '零售差价=零售金额-结算金额
                .TextMatrix(lngRow, mcol零售差价) = Format(Val(.TextMatrix(lngRow, mcol零售金额)) - Val(.TextMatrix(lngRow, mCol结算金额)), mFMT.FM_金额)
           End If
        Else
            .TextMatrix(lngRow, mcol零售价) = Format(Val(.TextMatrix(lngRow, mCol售价)) / dbl比例系数, mFMT.FM_散装零售价)
            .TextMatrix(lngRow, mcol零售金额) = Format(Val(.TextMatrix(lngRow, mcol零售价)) * (dbl数量 * dbl比例系数), mFMT.FM_金额)
            '零售差价=零售金额-结算金额
            .TextMatrix(lngRow, mcol零售差价) = Format(Val(.TextMatrix(lngRow, mcol零售金额)) - Val(.TextMatrix(lngRow, mCol结算金额)), mFMT.FM_金额)
        End If
    End With
    计算零售价及零售差价 = True
End Function


Private Sub mshBill_EditKeyPress(KeyAscii As Integer)
    Dim strKey As String
    Dim intDigit As Integer
    
    With mshBill
        strKey = .Text
        If strKey = "" Then
            strKey = .TextMatrix(.Row, .Col)
        End If
        Select Case .Col
            Case mcol商品条码
                Select Case KeyAscii
                    Case vbKeyBack, vbKeyEscape, 3, 22
                        Exit Sub
                    Case vbKeyReturn
'                        Call OS.PressKey(vbKeyTab)
                        Exit Sub
                    Case Else
                        '仅能录入数字和字母
                        If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z")) Or (KeyAscii >= Asc("a") And KeyAscii <= Asc("z")) Or (InStr(1, "`·！……；’、，。!@#$￥%^&*()_-——：”？?'+{}《》<>｛｝（）~[]:;'\|,./", Chr(KeyAscii)) > 0) Then Exit Sub
                End Select
                KeyAscii = 0
                Exit Sub
            Case mCol数量, mCol冲销数量
                intDigit = IIf(mintUnit = 1, g_小数位数.obj_包装小数.数量小数, g_小数位数.obj_散装小数.数量小数)
            Case mCol采购价, mCol结算价
               intDigit = IIf(mintUnit = 1, g_小数位数.obj_包装小数.成本价小数, g_小数位数.obj_散装小数.成本价小数)
            Case mCol结算金额, mCol发票金额
                intDigit = IIf(mintUnit = 1, g_小数位数.obj_包装小数.金额小数, g_小数位数.obj_散装小数.金额小数)
            Case mCol售价, mcol零售价
                intDigit = IIf(mintUnit = 1, g_小数位数.obj_包装小数.零售价小数, g_小数位数.obj_散装小数.零售价小数)
        End Select
        
        If .Col = mCol数量 Or .Col = mCol冲销数量 Or .Col = mCol采购价 Or .Col = mCol结算价 Or .Col = mCol结算金额 Or .Col = mCol发票金额 Or .Col = mCol售价 Or .Col = mcol零售价 Then
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
    Dim str批号 As String
    Dim strxq As String
    
    If mint编辑状态 = 5 And Trim(mshBill.TextMatrix(mshBill.Row, mCol发票号)) <> "" Then '当前行发票号不为空才可用
        cmdBulkCopy.Enabled = True
    Else
        cmdBulkCopy.Enabled = False
    End If

    
    With mshBill
        If Row > 0 Then
            .SetRowColor CLng(Row), &HFFCECE, True
        End If
        If .Row <> .LastRow Then
            lngRow = .LastRow
            If PicInput.Visible Then
                '重新计算零售价、差价
                .TextMatrix(lngRow, mCol售价) = Format(校正零售价(Val(.TextMatrix(lngRow, mCol结算价)) * (1 + (Val(Txt加价率) / 100)) + _
                时价材料零售价(Val(.TextMatrix(lngRow, 0)), Val(.TextMatrix(lngRow, mCol结算价)), Val(Txt加价率) / 100, lngRow), lngRow), mFMT.FM_成本价)
                .TextMatrix(lngRow, mCol售价金额) = Format(Val(.TextMatrix(lngRow, mCol售价)) * Val(.TextMatrix(lngRow, mCol数量)), mFMT.FM_金额)
                .TextMatrix(lngRow, mCol差价) = Format(IIf(.TextMatrix(lngRow, mCol售价金额) = "", 0, .TextMatrix(lngRow, mCol售价金额)) - IIf(.TextMatrix(lngRow, mCol结算金额) = "", 0, .TextMatrix(lngRow, mCol结算金额)), mFMT.FM_金额)
                Call 计算零售价及零售差价(lngRow)
                PicInput.Visible = False
            End If
        End If
        
        SetInputFormat .Row
        
        If Not (.Col = mCol结算价 Or .Col = mCol采购价 Or .Col = mCol扣率 Or .Col = mCol结算金额) Then PicInput.Visible = False
        
        If .Col = mCol结算金额 And PicInput.Visible Then Txt加价率.SetFocus: Exit Sub
        If .Col = mCol扣率 And PicInput.Visible Then Txt加价率.SetFocus: Exit Sub
        
        Select Case .Col
            Case mCol诊疗
                .TxtCheck = False
                .MaxLength = 80
                
                '只在诊疗列才显示合计信息和库存数
                Call 显示合计金额
                Call 提示库存数
                
            Case mCol产地
                ImeLanguage True
                .TxtCheck = False
                .MaxLength = 30
                .TxtSetFocus
                
            Case mCol批号
                .TxtCheck = False
                .MaxLength = mintBatchNoLen
            Case mcol注册证号
                .TxtCheck = False
                .MaxLength = 50
            Case mcol商品条码
                .TxtCheck = False
                .MaxLength = 50
            Case mCol效期
                .TxtCheck = True
                .TextMask = "1234567890-"
                .MaxLength = 10
                .Value = Format(sys.Currentdate, "yyyy-mm-dd")
                
'                If Trim(.TextMatrix(.Row, mCol批号)) = "" Or IsNumeric(.TextMatrix(.Row, mCol批号)) = False Then
                    If Not IsDate(Trim(.TextMatrix(.Row, mcol生产日期))) Then
                        str批号 = ""
                    Else
                        str批号 = Format(.TextMatrix(.Row, mcol生产日期), "yyyymmdd")
                    End If
'                Else
'                    str批号 = Trim(.TextMatrix(.Row, mCol批号))
'                End If
                
                If str批号 <> "" And Trim(.TextMatrix(.Row, mCol原销期)) <> "" Then
                    '存储格式值:最大效期||指导差价率||是否变价||在用分批||库房分批

                    If IsNumeric(str批号) And Split(.TextMatrix(.Row, mCol原销期), "||")(0) <> "0" Then
                        strxq = UCase(str批号)
                        If Trim(.TextMatrix(.Row, mCol效期)) = "" Then
                            If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                                strxq = TranNumToDate(strxq, True)
                                If strxq = "" Then Exit Sub
                                
                                .TextMatrix(.Row, mCol效期) = Format(DateAdd("M", Split(.TextMatrix(.Row, mCol原销期), "||")(0), strxq), "yyyy-mm-dd")
                                Call CheckLapse(.TextMatrix(.Row, mCol效期))
                            End If
                        End If
                    End If
                End If
            Case mcol生产日期
                .TxtCheck = True
                .TextMask = "1234567890-"
                .MaxLength = 10
                .Value = Format(sys.Currentdate, "yyyy-mm-dd")
                If Trim(.TextMatrix(.Row, mCol批号)) = "" Or IsNumeric(.TextMatrix(.Row, mCol批号)) = False Then
                    If Not IsDate(Trim(.TextMatrix(.Row, mcol生产日期))) Then
                        str批号 = ""
                    Else
                        str批号 = Format(.TextMatrix(.Row, mcol生产日期), "yyyymmdd")
                    End If
                Else
                    str批号 = Trim(.TextMatrix(.Row, mCol批号))
                End If
                
                If str批号 <> "" Then
                    
                    '存储格式值:最大效期||指导差价率||是否变价||在用分批||库房分批
                    If IsNumeric(str批号) And Split(.TextMatrix(.Row, mCol原销期), "||")(0) <> "0" Then
                        strxq = UCase(str批号)
                        If Trim(.TextMatrix(.Row, mcol生产日期)) = "" Then
                            If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                                strxq = TranNumToDate(strxq)
                                If strxq = "" Then Exit Sub
                                .TextMatrix(.Row, mcol生产日期) = Format(strxq, "yyyy-mm-dd")
                            End If
                        End If
                    End If
                End If
            Case mcol灭菌日期
                .TxtCheck = True
                .Value = Format(sys.Currentdate, "yyyy-mm-dd")
                .TextMask = "1234567890-"
                .MaxLength = 10
            Case mcol灭菌失效期
                .TxtCheck = True
                .Value = Format(sys.Currentdate, "yyyy-mm-dd")
                .TextMask = "1234567890-"
                .MaxLength = 10
            Case mCol扣率
                .TxtCheck = True
                .MaxLength = 3
                .TextMask = ".1234567890"
                stbThis.Panels.Item(2) = .TextMatrix(.Row, mCol诊疗) & "的指导批发价为：" & .TextMatrix(.Row, mCol指导批发价)
                
            Case mCol结算价, mCol指导批发价, mCol采购价
                .TxtCheck = True
                .MaxLength = 11
                .TextMask = ".1234567890"
            Case mCol结算金额
                .TxtCheck = True
                .MaxLength = 14
                .TextMask = ".1234567890"
            Case mcol零售价
                .TxtCheck = True
                .MaxLength = 11
                .TextMask = ".1234567890"
                
            Case mCol售价
                .TxtCheck = True
                .MaxLength = 11
                .TextMask = ".1234567890"
            Case mCol数量
                .TxtCheck = True
                .MaxLength = 11
                .TextMask = ".1234567890"
            Case mCol冲销数量
                .TxtCheck = True
                .MaxLength = 11
                .TextMask = ".1234567890"
            Case mCol随货单号
                .TxtCheck = False
                .MaxLength = 200
            Case mCol验收结论
                .MaxLength = 100
            Case mCol发票号
                .TxtCheck = False
                .MaxLength = mint发票号Len
            Case mcol发票代码
                .TxtCheck = True
                .MaxLength = 20
                .TextMask = "1234567890"
            Case mCol发票金额
                .TxtCheck = True
                .MaxLength = 14
                .TextMask = "-.1234567890"
            Case mCol发票日期
                .TxtCheck = True
                .TextMask = "1234567890-"
                .Value = sys.Currentdate
                .MaxLength = 10
        End Select
        
        '高值材料
        Select Case mint编辑状态
            Case 3, 4, 5, 6, 7, 10
                vsfCostlyInfo.Enabled = False
                picCostly.Enabled = False
        End Select
        
        If .Row <> .LastRow Then
            '状态切换
            If .TextMatrix(.Row, 0) = "" Then
                vsfCostlyInfo.Visible = False
                Call Form_Resize
                Exit Sub
            Else
                vsfCostlyInfo.Visible = IsCostly(.TextMatrix(.Row, 0))
                '显示数据
                If vsfCostlyInfo.Visible Then
                    '定位
                    If mrsCostlyInfo Is Nothing Then Exit Sub
                    If mrsCostlyInfo.RecordCount > 0 Then mrsCostlyInfo.MoveFirst
                    mrsCostlyInfo.Find "SN=" & .TextMatrix(.Row, 1)
                    If Not mrsCostlyInfo.EOF Then
                        vsfCostlyInfo.TextMatrix(1, 1) = IIf(IsNull(mrsCostlyInfo!科室), "", mrsCostlyInfo!科室)
                        vsfCostlyInfo.TextMatrix(1, 2) = IIf(IsNull(mrsCostlyInfo!病人姓名), "", mrsCostlyInfo!病人姓名)
                        vsfCostlyInfo.TextMatrix(1, 3) = IIf(IsNull(mrsCostlyInfo!住院号), "", mrsCostlyInfo!住院号)
                        vsfCostlyInfo.TextMatrix(1, 4) = IIf(IsNull(mrsCostlyInfo!床号), "", mrsCostlyInfo!床号)
                        vsfCostlyInfo.TextMatrix(1, 5) = IIf(IsNull(mrsCostlyInfo!Id), "", mrsCostlyInfo!Id)
                    Else
                        vsfCostlyInfo.TextMatrix(1, 1) = ""
                        vsfCostlyInfo.TextMatrix(1, 2) = ""
                        vsfCostlyInfo.TextMatrix(1, 3) = ""
                        vsfCostlyInfo.TextMatrix(1, 4) = ""
                        vsfCostlyInfo.TextMatrix(1, 5) = ""
                    End If
                    vsfCostlyInfo.Col = vsfCostlyInfo.ColIndex("科室")
                End If
            End If
            Call Form_Resize
        End If
        
    End With
    
End Sub

Private Sub mshBill_GotFocus()
'    If mint编辑状态 <> 1 Then
        '高值材料
        '状态切换
        With mshBill
            If .TextMatrix(.Row, 0) = "" Then
                vsfCostlyInfo.Visible = False
                Call Form_Resize
                Exit Sub
            Else
                vsfCostlyInfo.Visible = IsCostly(.TextMatrix(.Row, 0))
                '显示数据
                If vsfCostlyInfo.Visible Then
                    '定位
                    Call Form_Resize
                    If mrsCostlyInfo Is Nothing Then
                        Exit Sub
                    End If
                    If mrsCostlyInfo.RecordCount > 0 Then
                        mrsCostlyInfo.MoveFirst
                    Else
                        Exit Sub
                    End If
                    mrsCostlyInfo.Find "SN=" & Val(.TextMatrix(.Row, 1))
                    If Not mrsCostlyInfo.EOF Then
                        vsfCostlyInfo.TextMatrix(1, 1) = IIf(IsNull(mrsCostlyInfo!科室), "", mrsCostlyInfo!科室)
                        vsfCostlyInfo.TextMatrix(1, 2) = IIf(IsNull(mrsCostlyInfo!病人姓名), "", mrsCostlyInfo!病人姓名)
                        vsfCostlyInfo.TextMatrix(1, 3) = IIf(IsNull(mrsCostlyInfo!住院号), "", mrsCostlyInfo!住院号)
                        vsfCostlyInfo.TextMatrix(1, 4) = IIf(IsNull(mrsCostlyInfo!床号), "", mrsCostlyInfo!床号)
                        vsfCostlyInfo.TextMatrix(1, 5) = IIf(IsNull(mrsCostlyInfo!Id), "", mrsCostlyInfo!Id)
                    End If
                    vsfCostlyInfo.Col = vsfCostlyInfo.ColIndex("科室")
                End If
            End If
        End With
        Call Form_Resize
'    End If
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim rstemp As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim dbl加成率 As Double, dbl发票金额 As Double
    Dim dbl指导零售价 As Double
    Dim dbl售价 As Double, dbl采购价 As Double, dbl结算价 As Double, dbl扣率 As Double, dbl数量 As Double
    Dim lng材料ID As Long
    Dim sng分段售价 As Double
    Dim dblCostPrice As Double, dblPrice As Double
    Dim strBidMess As String
    Dim dbl成本价 As Double
    Dim intCol As Integer
    Dim i As Integer
    Dim int点击行 As Integer
    
    int点击行 = mshBill.Row

    On Error GoTo ErrHandle
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With mshBill
        .Text = Trim(.Text)
        strKey = Trim(.Text)
        
        If Mid(strKey, 1, 1) = "[" Then
            If InStr(2, strKey, "]") <> 0 Then
                strKey = Mid(strKey, 2, InStr(2, strKey, "]") - 2)
            Else
                strKey = Mid(strKey, 2)
            End If
        End If
        Select Case .Col
            Case mCol诊疗
                If strKey <> "" Then
                    Dim mrsReturn As Recordset
                    Dim sngLeft As Single
                    Dim sngTop As Single
                    
                    sngLeft = Me.Left + Pic单据.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                    sngTop = Me.Top + Me.Height - Me.ScaleHeight + Pic单据.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                    
'                    If sngTop + 3630 > Screen.Height Then
'                        sngTop = sngTop - mshBill.MsfObj.CellHeight - 3630
'                    End If
                    If mint编辑状态 = 8 Or mbln退货 Then
                        If Val(txtProvider.Tag) = 0 Then
                            ShowMsgBox "未选择退货单位!"
                            Cancel = True
                            If txtProvider.Enabled Then txtProvider.SetFocus
                            Exit Sub
                        End If
                    End If
                    Set mrsReturn = FrmMulitSel.ShowSelect(Me, IIf(mint编辑状态 = 8 Or mbln退货, 2, 1) _
                                  , cboStock.ItemData(cboStock.ListIndex) _
                                  , cboStock.ItemData(cboStock.ListIndex) _
                                  , cboStock.ItemData(cboStock.ListIndex) _
                                  , strKey, sngLeft, sngTop, mshBill.MsfObj.CellWidth, mshBill.MsfObj.CellHeight _
                                  , True, True, False, False, True _
                                  , IIf(mint编辑状态 = 8 Or mbln退货, Val(txtProvider.Tag), 0), IIf(mintUnit = 0, True, False), , , 1712, , mstrPrivs, , False)
                    
                    If mrsReturn.RecordCount > 0 And (mint编辑状态 = 8 Or mbln退货 = True) Then
                        Set mrsReturn = CheckRedo(mrsReturn) '检查重复记录，将重复的记录过滤掉然后返回过滤后的数据集
                    End If
                    
                    If mrsReturn.RecordCount <= 0 Then
                        Cancel = True
                        Exit Sub
                    End If
                    
                    mrsReturn.MoveFirst
                    For i = 1 To mrsReturn.RecordCount
                        If CheckQualifications(mlngModule, 0, Val(mrsReturn!材料ID)) = False Then Exit Sub
                        
                        If SetColValue(.Row, mrsReturn!材料ID, "[" & mrsReturn!编码 & "]" & mrsReturn!名称, IIf(IsNull(mrsReturn!规格), "", mrsReturn!规格), _
                                    IIf(IsNull(mrsReturn!产地), "", mrsReturn!产地), _
                                    IIf(mintUnit = 0, mrsReturn!散装单位, mrsReturn!包装单位), _
                                    IIf(IsNull(mrsReturn!售价), 0, mrsReturn!售价) * IIf(mintUnit = 0, 1, mrsReturn!换算系数), _
                                    mrsReturn!指导批发价 * IIf(mintUnit = 0, 1, mrsReturn!换算系数), _
                                    IIf(IsNull(mrsReturn!产地), "!", mrsReturn!产地), mrsReturn!最大效期, "", _
                                    IIf(mintUnit = 0, 1, mrsReturn!换算系数), IIf(IsNull(mrsReturn!批次), 0, mrsReturn!批次), mrsReturn!时价, _
                                    mrsReturn!在用分批, mrsReturn!指导差价率 / 100, IIf(IsNull(mrsReturn!批准文号), "", mrsReturn!批准文号), IIf(IsNull(mrsReturn!商品名), "", mrsReturn!商品名)) Then
                            
                            If .Row = .Rows - 1 Then .Rows = .Rows + 1 '只有当前行是最后一行时才新增行
                            .Row = .Row + 1
                            
                            .Text = .TextMatrix(.Row, .Col)
                        Else
                            Cancel = True
                        End If
                        
                        mrsReturn.MoveNext
                    Next
                    
                    mshBill.Row = int点击行
                    
'                    If mrsReturn.RecordCount = 1 Then
'                        If CheckQualifications(mlngModule, 0, Val(mrsReturn!材料ID)) = False Then Exit Sub
'
'                        If SetColValue(.Row, mrsReturn!材料ID, "[" & mrsReturn!编码 & "]" & mrsReturn!名称, IIf(IsNull(mrsReturn!规格), "", mrsReturn!规格), _
'                                    IIf(IsNull(mrsReturn!产地), "", mrsReturn!产地), _
'                                    IIf(mintUnit = 0, mrsReturn!散装单位, mrsReturn!包装单位), _
'                                    IIf(IsNull(mrsReturn!售价), 0, mrsReturn!售价) * IIf(mintUnit = 0, 1, mrsReturn!换算系数), _
'                                    mrsReturn!指导批发价 * IIf(mintUnit = 0, 1, mrsReturn!换算系数), _
'                                    IIf(IsNull(mrsReturn!产地), "!", mrsReturn!产地), mrsReturn!最大效期, "", _
'                                    IIf(mintUnit = 0, 1, mrsReturn!换算系数), IIf(IsNull(mrsReturn!批次), 0, mrsReturn!批次), mrsReturn!时价, _
'                                    mrsReturn!在用分批, mrsReturn!指导差价率 / 100, IIf(IsNull(mrsReturn!批准文号), "", mrsReturn!批准文号)) = False Then ' mrsReturn!简码
'                             Cancel = True
'                             Exit Sub
'                         End If
'                        .Text = .TextMatrix(.Row, .Col)
'                    Else
'                        Cancel = True
'                    End If
                    Call 提示库存数
                    '高值材料新增
                    If .TextMatrix(.Row, 0) = "" Then
                        vsfCostlyInfo.Visible = False
                    Else
                        vsfCostlyInfo.Visible = IsCostly(.TextMatrix(.Row, 0))
                    End If
                    CostlyInfo_Refresh Val(.TextMatrix(.Row, 1)), vsfCostlyInfo.Visible
                    Call Form_Resize
                End If
            Case mCol产地
                '无处理
                If .Text = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mCol产地) = ""
                    End If
                    Call ColMoveNextCol(.Col)
                    '.Col = mCol批号
                    Cancel = True
                    Exit Sub
                Else
                    Dim rs产地 As New Recordset
                    
                    .Text = UCase(Trim(.Text))
                    strKey = Trim(.Text)
                    
                    gstrSQL = "" & _
                        "   Select 编码,简码,名称 From 材料生产商 " & _
                        "   Where upper(名称) like [1] or Upper(编码) like [1] or Upper(简码) like [1]"
                    
                    Set rs产地 = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, IIf(gstrMatchMethod = "0", "%", "") & strKey & "%")
                    
                    
                    If rs产地.EOF Then
                        If MsgBox("卫生材料生产商没有找到你输入的产地，你要把它加入卫生材料生产商中吗？", vbYesNo + vbQuestion, mstrCaption) = vbNo Then
                            Cancel = True
                            Exit Sub
                        Else
                            Dim rsMax As New Recordset
                            Dim int编码 As Integer, strCode As String, strSpecify As String
                            
                            If rsMax.State = 1 Then rsMax.Close
                            gstrSQL = "SELECT Nvl(MAX(LENGTH(编码)),2) As Length FROM 材料生产商"
                            zlDatabase.OpenRecordset rsMax, gstrSQL, mstrCaption
                            int编码 = rsMax!Length
                            
                            gstrSQL = "SELECT Nvl(MAX(LPAD(编码," & int编码 & ",'0')),'00') As Code FROM 材料生产商"
                            rsMax.Close
                            zlDatabase.OpenRecordset rsMax, gstrSQL, mstrCaption
                            strCode = rsMax!Code
                            
                            int编码 = Len(strCode)
                            strCode = strCode + 1
                            
                            If int编码 >= Len(strCode) Then
                                strCode = String(int编码 - Len(strCode), "0") & strCode
                            End If
                            strSpecify = zlStr.GetCodeByVB(strKey)
                            
                            
                            gstrSQL = "ZL_材料生产商_INSERT('" & strCode & "','" & strKey & "','" & strSpecify & "')"
                            Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
                        End If
                    Else
                        If rs产地.RecordCount = 1 Then
                            If CheckQualifications(mlngModule, 1, rs产地.Fields("名称")) = False Then
                                Exit Sub
                            End If
                            
                            .TextMatrix(.Row, mCol产地) = rs产地.Fields("名称")
                            .Text = rs产地.Fields("名称")
                            
                            gstrSQL = "select 批准文号 from 药品生产商对照 where 厂家名称=[1] and 药品id=[2]"
                            Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "mshbill_CommandClick", .TextMatrix(.Row, mCol产地), .TextMatrix(mshBill.Row, 0))
                            If rstemp.RecordCount > 0 Then
                                .TextMatrix(.Row, mcol批准文号) = IIf(IsNull(rstemp!批准文号), "", rstemp!批准文号)
                            Else
                                .TextMatrix(.Row, mcol批准文号) = ""
                            End If
                        Else
                            Set msh产地.Recordset = rs产地
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
                OS.OpenIme False
            
            Case mCol验收结论
                If Val(.TextMatrix(.Row, 0)) = 0 Then Exit Sub
                
                If .Text = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mCol验收结论) = ""
                    End If
                    Call ColMoveNextCol(.Col)
                    Cancel = True
                    Exit Sub
                Else
                    Dim rs结论 As New Recordset
                    
                    gstrSQL = "" & _
                        "   Select 编码,名称 From 入库验收结论 " & _
                        "   Where upper(名称) like [1] or Upper(编码) like [1] "
                    Set rs结论 = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, IIf(gstrMatchMethod = "0", "%", "") & strKey & "%")
                    
                    If rs结论.EOF Then
                        MsgBox "入库验收结论没有找到，请重新输入！", vbInformation, mstrCaption
                        Cancel = True
                        Exit Sub
                    Else
                        If rs结论.RecordCount = 1 Then
                            .TextMatrix(.Row, mCol验收结论) = rs结论.Fields("名称")
                            .Text = rs结论.Fields("名称")
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
                OS.OpenIme False
            Case mcol批准文号
                If strKey = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mcol批准文号) = ""
                    End If
                    Call ColMoveNextCol(.Col)
                    Cancel = True
                    Exit Sub
                End If
            Case mCol批号
                If Len(strKey) > mintBatchNoLen Then
                    ShowMsgBox "批号不能大于" & mintBatchNoLen & "位"
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
               
                '无处理
                If strKey = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mCol批号) = ""
                    End If
                    Call ColMoveNextCol(.Col)
                    '.Col = mcol生产日期
                    Cancel = True
                    Exit Sub
                End If
            Case mcol注册证号
                If LenB(StrConv(strKey, vbFromUnicode)) > 50 Then
                    ShowMsgBox "注册证号不能大于50个字符或25个汉字,请检查!"
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                '无处理
                If strKey = "" Then
                    If .TxtVisible = True Then
                        .Text = " "
                    Else
                        If Trim(.TextMatrix(.Row, mcol注册证号)) = "" Then
                            .TextMatrix(.Row, mcol注册证号) = " "
                        End If
                        
                    End If
                    Exit Sub
                End If
            Case mcol商品条码
                If Len(Trim(strKey)) > 50 Then
                    ShowMsgBox "商品条码不能大于50个字符,请检查!"
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                '无处理
                If strKey = "" Then
                    If .TxtVisible = True Then
                        .Text = " "
                    Else
                        If Trim(.TextMatrix(.Row, mcol商品条码)) = "" Then
                            .TextMatrix(.Row, mcol商品条码) = " "
                        End If
                        
                    End If
                    Exit Sub
                End If
                .Text = UCase(.Text)
            Case mCol效期
                '有处理
                If strKey <> "" Then
                    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
                        strKey = TranNumToDate(strKey)
                        If strKey = "" Then
                            MsgBox "失效期必须为日期型！", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        .Text = strKey
                        Exit Sub
                    End If
                    If Not IsDate(strKey) Then
                        MsgBox "失效期必须为日期型如(2000-10-10) 或（20001010）,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                ElseIf strKey = "" And strKey <> .TextMatrix(.Row, mCol效期) Then
                
                    If .TxtVisible = True Then
                        .Text = " "
                        Exit Sub
                    End If
                    
                    Exit Sub
                End If
            Case mcol生产日期
                '有处理
                Dim str批号 As String, strxq As String
                If strKey <> "" Then
                    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
                        strKey = TranNumToDate(strKey)
                        If strKey = "" Then
                            MsgBox "生产日期必须为日期型！", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        .Text = strKey
                        Exit Sub
                    End If
                    If Not IsDate(strKey) Then
                        MsgBox "生产日期必须为日期型如（2000-10-10）或（20001010）,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                
                    If .ColData(mCol效期) = 5 Then
                        If Not IsDate(Trim(strKey)) Then
                            str批号 = ""
                        Else
                            str批号 = Format(strKey, "yyyymmdd")
                        End If
                        If str批号 <> "" And Trim(.TextMatrix(.Row, mCol原销期)) <> "" Then
                            '存储格式值:最大效期||指导差价率||是否变价||在用分批||库房分批
                            If IsNumeric(str批号) And Split(.TextMatrix(.Row, mCol原销期), "||")(0) <> "0" Then
                                strxq = UCase(str批号)
                                If Trim(.TextMatrix(.Row, mCol效期)) = "" Then
                                    If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                                        strxq = TranNumToDate(strxq, True)
                                        If strxq = "" Then Exit Sub
                                        
                                        .TextMatrix(.Row, mCol效期) = Format(DateAdd("M", Split(.TextMatrix(.Row, mCol原销期), "||")(0), strxq), "yyyy-mm-dd")
                                        Call CheckLapse(.TextMatrix(.Row, mCol效期))
                                    End If
                                End If
                            End If
                        End If
                    End If
                
                ElseIf strKey = "" And strKey <> .TextMatrix(.Row, mcol生产日期) Then
                
                    If .ColData(mCol效期) = 5 And .TextMatrix(.Row, mcol生产日期) <> "" Then
                        If Not IsDate(Trim(.TextMatrix(.Row, mcol生产日期))) Then
                            str批号 = ""
                        Else
                            str批号 = Format(.TextMatrix(.Row, mcol生产日期), "yyyymmdd")
                        End If
                        If str批号 <> "" And Trim(.TextMatrix(.Row, mCol原销期)) <> "" Then
                            '存储格式值:最大效期||指导差价率||是否变价||在用分批||库房分批
                            If IsNumeric(str批号) And Split(.TextMatrix(.Row, mCol原销期), "||")(0) <> "0" Then
                                strxq = UCase(str批号)
                                If Trim(.TextMatrix(.Row, mCol效期)) = "" Then
                                    If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                                        strxq = TranNumToDate(strxq, True)
                                        If strxq = "" Then Exit Sub
                                        
                                        .TextMatrix(.Row, mCol效期) = Format(DateAdd("M", Split(.TextMatrix(.Row, mCol原销期), "||")(0), strxq), "yyyy-mm-dd")
                                        Call CheckLapse(.TextMatrix(.Row, mCol效期))
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                    If .TxtVisible = True Then
                        .Text = " "
                        Exit Sub
                    Else
                        
                    End If
                    
                    Exit Sub
                End If

            Case mCol扣率
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "扣率必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" And strKey <> .TextMatrix(.Row, mCol扣率) Then
                    SetDisCount .Row, strKey
                End If
                
                Call 检查成本价
                Call 显示合计金额
            Case mcol灭菌日期
                '有处理
                If strKey <> "" Then
                    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
                        strKey = TranNumToDate(strKey)
                        If strKey = "" Then
                            MsgBox "灭菌日期必须为日期型！", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        .Text = strKey
                        'Exit Sub
                    End If
                    If Not IsDate(strKey) Then
                        MsgBox "灭菌日期必须为日期型如(2000-10-10) 或（20001010）,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                                
                        Exit Sub
                    End If
                    If Format(sys.Currentdate, "yyyy-mm-dd") >= Format(DateAdd("m", Val(.TextMatrix(.Row, mcol灭菌效期)), CDate(strKey)), "yyyy-mm-dd") Then
                        If MsgBox("该卫生材料已经过了灭菌失效期(" & Format(DateAdd("m", Val(.TextMatrix(.Row, mcol灭菌效期)), CDate(strKey)), "yyyy-mm-dd") & "),是否还要进行入库!", vbQuestion + vbDefaultButton2 + vbYesNo) = vbNo Then
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                    
                    .Text = strKey
                    '计算失效期
                    .TextMatrix(.Row, mcol灭菌失效期) = Format(DateAdd("m", Val(.TextMatrix(.Row, mcol灭菌效期)), CDate(strKey)), "yyyy-mm-dd")
                ElseIf strKey = "" And strKey <> .TextMatrix(.Row, mcol灭菌日期) Then
                    If .TxtVisible = True Then
                        .Text = " "
                        Exit Sub
                    End If
                    Exit Sub
                End If
                
            Case mcol灭菌失效期
                If strKey <> "" Then
                    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
                        strKey = TranNumToDate(strKey)
                        If strKey = "" Then
                            MsgBox "灭菌失效期必须为日期型！", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        .Text = strKey
                        'Exit Sub
                    End If
                    If Not IsDate(strKey) Then
                        MsgBox "灭菌失效期必须为日期型如(2000-10-10) 或（20001010）,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    If Format(sys.Currentdate, "yyyy-mm-dd") >= Format(DateAdd("m", Val(.TextMatrix(.Row, mcol灭菌效期)), CDate(strKey)), "yyyy-mm-dd") Then
                        If MsgBox("该卫生材料已经过了灭菌失效期(" & Format(DateAdd("m", Val(.TextMatrix(.Row, mcol灭菌效期)), CDate(strKey)), "yyyy-mm-dd") & "),是否还要进行入库!", vbQuestion + vbDefaultButton2 + vbYesNo) = vbNo Then
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If

                    .Text = strKey
                ElseIf strKey = "" And strKey <> .TextMatrix(.Row, mcol灭菌失效期) Then
                    If .TxtVisible = True Then
                        .Text = " "
                        Exit Sub
                    End If
                    Exit Sub
                End If
                
            Case mCol指导批发价
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "指导批发价必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" And strKey <> .TextMatrix(.Row, mCol指导批发价) Then
                    SetDisCount .Row, strKey
                End If
                
                Call 检查成本价
                Call 显示合计金额
            Case mCol采购价
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "采购价必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Val(strKey) < 0 Then
                        MsgBox "采购价必须大于0,请重输！", vbInformation + vbOKOnly, gstrSysName
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
                    .Text = Format(strKey, mFMT.FM_成本价)
                    .TextMatrix(.Row, .Col) = .Text
                    If mbln不强制控制指导价格 = False Then
                        If Val(strKey) > Val(.TextMatrix(.Row, mCol指导批发价)) Then
                            MsgBox "你输入的采购价(" & strKey & ")大于了采购限价(" & .TextMatrix(.Row, mCol指导批发价) & ")。", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                    
                    '中标成本价判断
                    dblCostPrice = Get中标单位成本价(.TextMatrix(.Row, lng材料ID))
                    dblPrice = CDbl(IIf(.Text <> "", .Text, IIf(.TextMatrix(.Row, mCol采购价) = "", 0, .TextMatrix(.Row, mCol采购价))))
                    If dblCostPrice < dblPrice And dblCostPrice <> 0 Then
                        strBidMess = zlDatabase.GetPara("入库单价超中标单价", glngSys, mlngModule)
                        If Val(strBidMess) = 0 Then     '禁止入库单价超中标单价
                            MsgBox "禁止采购价（" & dblPrice & "）超 中标单价（" & dblCostPrice & "）。", vbCritical, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        ElseIf Val(strBidMess) = 1 Then '提示
                            MsgBox "采购价（" & dblPrice & "）超 中标单价（" & dblCostPrice & "）！", vbInformation, gstrSysName
                        End If
                    End If
                End If
                If strKey <> "" Then
                    strKey = Format(strKey, mFMT.FM_成本价)
                    .Text = strKey
                    .TextMatrix(.Row, mCol采购价) = .Text
                End If
                .TextMatrix(.Row, mCol结算价) = Format(Val(.TextMatrix(.Row, mCol采购价)) * Val(.TextMatrix(.Row, mCol扣率)) / 100, mFMT.FM_成本价)
                
                If ISCheckScalc售价(False, .Row) = False Then
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                End If
                Call 检查成本价
                Call 显示合计金额
            Case mCol结算价
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "结算价必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Val(strKey) < 0 Then
                        MsgBox "结算价必须大于0,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strKey) >= 10 ^ 11 - 1 Then
                        MsgBox "结算价必须小于" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    .Text = Format(strKey, mFMT.FM_成本价)
                    .TextMatrix(.Row, .Col) = .Text
                End If
                '中标成本价判断
                dblCostPrice = Get中标单位成本价(.TextMatrix(.Row, lng材料ID))
                dblPrice = CDbl(IIf(.Text <> "", .Text, IIf(.TextMatrix(.Row, mCol采购价) = "", 0, .TextMatrix(.Row, mCol采购价))))
                If dblCostPrice < dblPrice And dblCostPrice <> 0 Then
                    strBidMess = zlDatabase.GetPara("入库单价超中标单价", glngSys, mlngModule)
                    If Val(strBidMess) = 0 Then     '禁止入库单价超中标单价
                        MsgBox "禁止采购价（" & dblPrice & "）超 中标单价（" & dblCostPrice & "）。", vbCritical, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    ElseIf Val(strBidMess) = 1 Then '提示
                        MsgBox "采购价（" & dblPrice & "）超 中标单价（" & dblCostPrice & "）！", vbInformation, gstrSysName
                    End If
                End If
                
                '返回设置扣率
                If Val(.TextMatrix(.Row, mCol扣率)) = 0 Then
                
                    If strKey <> "" And Val(.TextMatrix(.Row, mCol指导批发价)) <> 0 Then
                        .TextMatrix(.Row, mCol扣率) = Format((strKey / .TextMatrix(.Row, mCol指导批发价)) * 100, mFMT.FM_成本价)
                    Else
                        .TextMatrix(.Row, mCol扣率) = 100
                    End If
                End If
                If strKey <> "" Then
                    strKey = Format(strKey, mFMT.FM_成本价)
                    .Text = strKey
                    .TextMatrix(.Row, mCol结算价) = .Text
                End If
                If Val(.TextMatrix(.Row, mCol采购价)) <> 0 Then
                    .TextMatrix(.Row, mCol采购价) = Format((Val(.TextMatrix(.Row, mCol结算价)) / .TextMatrix(.Row, mCol扣率)) * 100, mFMT.FM_成本价)
                Else
                    .TextMatrix(.Row, mCol采购价) = Format(Val(.TextMatrix(.Row, mCol结算价)), mFMT.FM_成本价)
                    .TextMatrix(.Row, mCol扣率) = "100"
                End If
        
                If ISCheckScalc售价(True, .Row) = False Then
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If

                Call 检查成本价
                Call 显示合计金额
            Case mCol结算金额
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "结算金额必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Abs(Val(strKey)) < 0 Then
                        MsgBox "结算金额的必须大于0,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strKey) >= 10 ^ 14 - 1 Then
                        MsgBox "结算金额必须小于" & (10 ^ 14 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                End If
                
                dbl数量 = Val(.TextMatrix(.Row, mCol数量))
                dbl扣率 = Val(.TextMatrix(.Row, mCol扣率))
                dbl售价 = Val(.TextMatrix(.Row, mCol售价))
                
                lng材料ID = Val(.TextMatrix(.Row, 0))
                If strKey <> "" And strKey <> .TextMatrix(.Row, mCol结算金额) Then
                    If dbl数量 <> 0 Then
                        dbl结算价 = Val(strKey) / dbl数量
                        dbl采购价 = dbl结算价 * 100 / IIf(dbl扣率 = 0, 1, dbl扣率)
                        .TextMatrix(.Row, mCol结算价) = Format(dbl结算价, mFMT.FM_成本价)
                        .TextMatrix(.Row, mCol采购价) = Format(dbl采购价, mFMT.FM_成本价)
                        dbl成本价 = IIf(mbln时价购前销售 = True, dbl采购价, dbl结算价)
                        
                        If mbln加价率 = True Then
                            '取得改变结算金额前的加价率
                            mdbl加价率 = 15
                            If dbl售价 <> 0 And dbl结算价 <> 0 Then
                                If Val(Replace(.TextMatrix(.Row, mcol加成率), "%", "")) >= 0 Then
                                    mdbl加价率 = Val(Replace(.TextMatrix(.Row, mcol加成率), "%", ""))
                                Else
                                    mdbl加价率 = 计算加成率(lng材料ID, dbl售价, dbl成本价)
                                End If
                            End If
                        End If
                        
                        '对时价材料的处理
                        If .TextMatrix(.Row, mCol原销期) <> "" Then
                            '存储格式值:最大效期||指导差价率||是否变价||在用分批||库房分批
                            
                            '重新计算零售价、差价
                            If Split(.TextMatrix(.Row, mCol原销期), "||")(2) = 1 Then
                                '由于存在差价让利比的存在,需要按加成率计算,因此将指导差价率转换成加成率计算 公式：加成率=1/(1-差价率)-1
                                If mbln加价率 = True Then
                                    If mint编辑状态 = 8 And dbl售价 <> 0 Then
                                    Else
                                        .TextMatrix(.Row, mCol售价) = Format(校正零售价(dbl成本价 * (1 + (mdbl加价率 / 100)) + _
                                        时价材料零售价(lng材料ID, dbl成本价, (mdbl加价率 / 100))), mFMT.FM_零售价)
                                    End If
                                    .TextMatrix(.Row, mCol售价金额) = Format(Val(.TextMatrix(.Row, mCol售价)) * dbl数量, mFMT.FM_金额)
                                    .TextMatrix(.Row, mCol差价) = Format(IIf(.TextMatrix(.Row, mCol售价金额) = "", 0, .TextMatrix(.Row, mCol售价金额)) - IIf(.TextMatrix(.Row, mCol结算金额) = "", 0, .TextMatrix(.Row, mCol结算金额)), mFMT.FM_金额)
                                    '刘兴宏:零售价处理
                                    Call 计算零售价及零售差价(.Row)
                                ElseIf mbln分段加成率 = True Then
                                    dbl加成率 = 0
                                    If mint编辑状态 = 8 And dbl售价 <> 0 Then
                                    Else
                                        If Get分段加成售价(dbl成本价, Val(.TextMatrix(.Row, mCol比例系数)), mstrCaption, sng分段售价) = False Then
                                            Cancel = True
                                            .TxtSetFocus
                                            Exit Sub
                                        End If
                                        .TextMatrix(.Row, mCol售价) = Format(校正零售价(sng分段售价 + _
                                                                      时价材料零售价(lng材料ID, dbl成本价, dbl加成率, -1, sng分段售价)) _
                                                                      , mFMT.FM_零售价)
                                    End If
                                    .TextMatrix(.Row, mCol售价金额) = Format(dbl数量 * Val(.TextMatrix(.Row, mCol售价)), mFMT.FM_金额)
                                    '刘兴宏:零售价处理
                                    Call 计算零售价及零售差价(.Row)
                                Else 'mbln时价卫材取上次售价 = True或者3种取售价方式都没有设置时，优先从上次取，如果没有则按照加成率方式取
                                    If mbln时价卫材取上次售价 = True Then
                                        gstrSQL = "Select Nvl(上次售价, 0) As 上次售价 From 材料特性 Where 材料id = [1]"
                                        Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng材料ID)
                                        If rstemp!上次售价 > 0 Then
                                            .TextMatrix(.Row, mCol售价) = Format(zlStr.Nvl(rstemp!上次售价, 0) * Val(.TextMatrix(.Row, mCol比例系数)), mFMT.FM_零售价)
                                            If dbl成本价 <> 0 Then
                                                .TextMatrix(.Row, mcol加成率) = Format((Val(.TextMatrix(.Row, mCol售价)) / dbl成本价 - 1) * 100, "###0.00") & "%"
                                            End If
                                        Else
                                            '由于存在差价让利比的存在,需要按加成率计算,因此将指导差价率转换成加成率计算 公式：加成率=1/(1-差价率)-1
                                            If dbl成本价 <> 0 Then
                                                mdbl加价率 = Val(Replace(.TextMatrix(.Row, mcol加成率), "%", "")) '计算加成率(lng材料ID, dbl售价, dbl成本价)
                                                .TextMatrix(.Row, mcol加成率) = Format(mdbl加价率, "####0.00") & "%"
                                            End If
                                            If mint编辑状态 = 8 And dbl售价 <> 0 Then
                                            Else
                                                .TextMatrix(.Row, mCol售价) = Format(校正零售价(dbl成本价 * (1 + (mdbl加价率 / 100)) + _
                                                时价材料零售价(lng材料ID, dbl成本价, (mdbl加价率 / 100))), mFMT.FM_零售价)
                                            End If
                                        End If
                                    Else
                                        If dbl成本价 <> 0 Then
                                            mdbl加价率 = Val(Replace(.TextMatrix(.Row, mcol加成率), "%", "")) '计算加成率(lng材料ID, dbl售价, dbl成本价)
                                            .TextMatrix(.Row, mcol加成率) = Format(mdbl加价率, "####0.00") & "%"
                                        End If
                                        If mint编辑状态 = 8 And dbl售价 <> 0 Then
                                        Else
                                            .TextMatrix(.Row, mCol售价) = Format(校正零售价(dbl成本价 * (1 + (mdbl加价率 / 100)) + _
                                            时价材料零售价(lng材料ID, dbl成本价, (mdbl加价率 / 100))), mFMT.FM_零售价)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                        
                    If .TextMatrix(.Row, mCol指导批发价) <> "" Then
                        If Val(.TextMatrix(.Row, mCol指导批发价)) = 0 Then
                             .TextMatrix(.Row, mCol扣率) = 100
'                        Else
'                             .TextMatrix(.Row, mCol扣率) = Format(Val(.TextMatrix(.Row, mCol结算价)) / Val(.TextMatrix(.Row, mCol指导批发价)) * 100, mFMT.FM_成本价)
                        End If
                    End If
                    If strKey <> "" Then
                         .Text = Format(strKey, mFMT.FM_金额)

                    End If
                    .TextMatrix(.Row, mCol发票金额) = IIf(Trim(.TextMatrix(.Row, mCol发票号)) = "" And Trim(.TextMatrix(.Row, mcol发票代码)) = "", "", Format(strKey, mFMT.FM_金额))
                    .TextMatrix(.Row, mCol差价) = Format(IIf(.TextMatrix(.Row, mCol售价金额) = "", 0, .TextMatrix(.Row, mCol售价金额)) - strKey, mFMT.FM_金额)
                    .TextMatrix(.Row, mCol结算金额) = Format(strKey, mFMT.FM_金额)
                    '刘兴宏:零售价处理
                    Call 计算零售价及零售差价(.Row)
                End If
                
                Call 检查成本价
                Call 显示合计金额
                
            Case mCol数量
                If .TextMatrix(.Row, .Col) = "" And strKey = "" Then
                    MsgBox "数量必须输入！", vbOKOnly + vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "数量必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Val(strKey) = 0 Then
                        MsgBox "数量不能为零,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Abs(Val(strKey)) < 0.001 Then
                            MsgBox "数量的必须大于0.001,请重输！", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        
                    If Val(strKey) >= 10 ^ 11 - 1 Then
                        MsgBox "数量必须小于" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    strKey = Format(strKey, mFMT.FM_数量)
                    '检查是否有足够的库存可以退货
                    If mint编辑状态 = 8 Or mbln退货 Then
                        If Not CheckStock(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mCol批次)), Val(.Text) * Val(.TextMatrix(.Row, mCol比例系数))) Then
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                    
                    .Text = strKey
                    lng材料ID = Val(.TextMatrix(.Row, 0))
                    dbl数量 = Val(strKey)
                    dbl结算价 = Val(.TextMatrix(.Row, mCol结算价))
                    dbl采购价 = Val(.TextMatrix(.Row, mCol采购价))
                    dbl售价 = Val(.TextMatrix(.Row, mCol售价))
                    dbl成本价 = IIf(mbln时价购前销售 = True, dbl采购价, dbl结算价)
                    
                    .TextMatrix(.Row, mCol结算金额) = Format(dbl结算价 * Val(strKey), mFMT.FM_金额)
                    
                    '时价材料的处理
                    If .TextMatrix(.Row, mCol原销期) <> "" Then
                        '存储格式值:最大效期||指导差价率||是否变价||在用分批||库房分批
                        If Split(.TextMatrix(.Row, mCol原销期), "||")(2) = 1 Then
                            '由于存在差价让利比的存在,需要按加成率计算,因此将指导差价率转换成加成率计算 公式：加成率=1/(1-差价率)-1
                            If mbln加价率 = True Then
                                mdbl加价率 = 15
                                
                                If dbl成本价 <> 0 Then
                                    mdbl加价率 = 计算加成率(lng材料ID, dbl售价, dbl成本价)
                                    .TextMatrix(.Row, mcol加成率) = zlStr.FormatEx(mdbl加价率, 2) & "%"
                                End If
                                If mint编辑状态 = 8 And dbl售价 <> 0 Then
                                Else
                                    .TextMatrix(.Row, mCol售价) = Format(校正零售价(dbl成本价 * (1 + (mdbl加价率 / 100)) + _
                                    时价材料零售价(lng材料ID, dbl成本价, (mdbl加价率 / 100))), mFMT.FM_零售价)
                                End If
                                .TextMatrix(.Row, mCol售价金额) = Format(Val(.TextMatrix(.Row, mCol售价)) * strKey, mFMT.FM_金额)
                                .TextMatrix(.Row, mCol差价) = Format(IIf(.TextMatrix(.Row, mCol售价金额) = "", 0, .TextMatrix(.Row, mCol售价金额)) - IIf(.TextMatrix(.Row, mCol结算金额) = "", 0, .TextMatrix(.Row, mCol结算金额)), mFMT.FM_金额)
                                '刘兴宏:零售价处理
                                Call 计算零售价及零售差价(.Row)
                            ElseIf mbln分段加成率 = True Then
                                dbl加成率 = 0
                                If mint编辑状态 = 8 And dbl售价 <> 0 Then
                                Else
                                    If Get分段加成售价(dbl成本价, Val(.TextMatrix(.Row, mCol比例系数)), mstrCaption, sng分段售价) = False Then
                                        Cancel = True
                                        .TxtSetFocus
                                        Exit Sub
                                    End If
                                    .TextMatrix(.Row, mCol售价) = Format(校正零售价(sng分段售价 + _
                                                                  时价材料零售价(lng材料ID, dbl成本价, dbl加成率, -1, sng分段售价)) _
                                                                  , mFMT.FM_零售价)
                                End If
                                .TextMatrix(.Row, mcol加成率) = Format(dbl加成率 * 100, "####0.00") & "%" '因为是分段加成的所以加成率不准确，取一个模糊值即可
                            Else  'mbln时价卫材取上次售价 = True或者3种取售价方式都没有设置时，优先从上次取，如果没有则按照加成率方式取
                                If mbln时价卫材取上次售价 = True Then
                                    gstrSQL = "Select Nvl(上次售价, 0) As 上次售价 From 材料特性 Where 材料id = [1]"
                                    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng材料ID)
                                    If rstemp!上次售价 > 0 Then
                                        .TextMatrix(.Row, mCol售价) = Format(zlStr.Nvl(rstemp!上次售价, 0) * Val(.TextMatrix(.Row, mCol比例系数)), mFMT.FM_零售价)
                                        If dbl成本价 <> 0 Then
                                            .TextMatrix(.Row, mcol加成率) = Format((Val(.TextMatrix(.Row, mCol售价)) / dbl成本价 - 1) * 100, "###0.00") & "%"
                                        End If
                                    Else
                                        '由于存在差价让利比的存在,需要按加成率计算,因此将指导差价率转换成加成率计算 公式：加成率=1/(1-差价率)-1
                                        If dbl成本价 <> 0 Then
                                            mdbl加价率 = Val(Replace(.TextMatrix(.Row, mcol加成率), "%", "")) '计算加成率(lng材料ID, dbl售价, dbl成本价)
                                            .TextMatrix(.Row, mcol加成率) = Format(mdbl加价率, "####0.00") & "%"
                                        End If
                                        If mint编辑状态 = 8 And dbl售价 <> 0 Then
                                        Else
                                            .TextMatrix(.Row, mCol售价) = Format(校正零售价(dbl成本价 * (1 + (mdbl加价率 / 100)) + _
                                            时价材料零售价(lng材料ID, dbl成本价, (mdbl加价率 / 100))), mFMT.FM_零售价)
                                        End If
                                    End If
                                Else
                                    If dbl成本价 <> 0 Then
                                        mdbl加价率 = Val(Replace(.TextMatrix(.Row, mcol加成率), "%", "")) '计算加成率(lng材料ID, dbl售价, dbl成本价)
                                        .TextMatrix(.Row, mcol加成率) = Format(mdbl加价率, "####0.00") & "%"
                                    End If
                                    If mint编辑状态 = 8 And dbl售价 <> 0 Then
                                    Else
                                        .TextMatrix(.Row, mCol售价) = Format(校正零售价(dbl成本价 * (1 + (mdbl加价率 / 100)) + _
                                        时价材料零售价(lng材料ID, dbl成本价, (mdbl加价率 / 100))), mFMT.FM_零售价)
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                    If Val(.TextMatrix(.Row, mCol售价)) <> 0 Then
                        .TextMatrix(.Row, mCol售价金额) = Format(Val(.TextMatrix(.Row, mCol售价)) * Val(strKey), mFMT.FM_金额)
                    End If
                    .TextMatrix(.Row, mCol差价) = Format(IIf(.TextMatrix(.Row, mCol售价金额) = "", 0, .TextMatrix(.Row, mCol售价金额)) - IIf(.TextMatrix(.Row, mCol结算金额) = "", 0, .TextMatrix(.Row, mCol结算金额)), mFMT.FM_金额)
                    .TextMatrix(.Row, .Col) = strKey
                    '刘兴宏:零售价处理
                    Call 计算零售价及零售差价(.Row)
                    If mint编辑状态 = 8 Or (mint编辑状态 = 2 And mbln退货 = True) Then
                        .TextMatrix(.Row, mCol发票金额) = IIf(Trim(.TextMatrix(.Row, mCol发票号)) = "" And Trim(.TextMatrix(.Row, mcol发票代码)) = "", "", .TextMatrix(.Row, mCol售价金额))
                    Else
                        .TextMatrix(.Row, mCol发票金额) = IIf(Trim(.TextMatrix(.Row, mCol发票号)) = "" And Trim(.TextMatrix(.Row, mcol发票代码)) = "", "", .TextMatrix(.Row, mCol结算金额))
                    End If
                End If
                显示合计金额
            Case mCol冲销数量
                If .TextMatrix(.Row, .Col) = "" And strKey = "" Then
                    MsgBox "数量必须输入！", vbOKOnly + vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "数量必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Abs(Val(strKey)) > Abs(Val(.TextMatrix(.Row, mCol数量))) Then
                        MsgBox "冲销数量的绝对值不能大于原有数量的绝对值,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strKey) >= 10 ^ 11 - 1 Then
                        MsgBox "冲销数量必须小于" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    strKey = Format(strKey, mFMT.FM_数量)
                    .Text = strKey
                    
                    If Val(.TextMatrix(.Row, mCol结算价)) <> 0 Then
                        .TextMatrix(.Row, mCol结算金额) = Format(Val(.TextMatrix(.Row, mCol结算价)) * strKey, mFMT.FM_金额)
                    End If
                    If .TextMatrix(.Row, mCol售价) <> "" Then
                        .TextMatrix(.Row, mCol售价金额) = Format(Val(.TextMatrix(.Row, mCol售价)) * strKey, mFMT.FM_金额)
                    End If
                    .TextMatrix(.Row, mCol差价) = Format(IIf(.TextMatrix(.Row, mCol售价金额) = "", 0, .TextMatrix(.Row, mCol售价金额)) - IIf(.TextMatrix(.Row, mCol结算金额) = "", 0, .TextMatrix(.Row, mCol结算金额)), mFMT.FM_金额)
                    '刘兴宏:零售价处理
                    Call 计算零售价及零售差价(.Row, False)
                    If Trim(.TextMatrix(.Row, mCol发票号)) <> "" Or Trim(.TextMatrix(.Row, mCol随货单号)) <> "" Or Trim(.TextMatrix(.Row, mcol发票代码)) <> "" Then
                    
                        dbl发票金额 = GetTotale发票金额(mstr单据号, Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mCol序号)))
                        If Val(.TextMatrix(.Row, mCol数量)) = 0 Then
                            .TextMatrix(.Row, mCol发票金额) = Format(0, mFMT.FM_金额)
                        Else
                            .TextMatrix(.Row, mCol发票金额) = Format(Val(strKey) / Val(.TextMatrix(.Row, mCol数量)) * dbl发票金额, mFMT.FM_金额)
                        End If
                    End If
                End If
                
                显示合计金额
                
            Case mCol发票号
                
                If Trim(.Text) = "" Then
                    If .TxtVisible = True Then
                        .ColData(mCol发票日期) = 5
                        .ColData(mCol发票金额) = 5
                        .ColData(mcol发票代码) = 5
                        .TextMatrix(.Row, mCol发票金额) = ""
                        .TextMatrix(.Row, mCol发票日期) = ""
                        .TextMatrix(.Row, mcol发票代码) = ""
                        .TextMatrix(.Row, .Col) = " "
                        .Text = " "
                    ElseIf .TxtVisible = False Then
                        If Trim(.TextMatrix(.Row, mCol发票号)) = "" Then
                            .ColData(mCol发票日期) = 5
                            .ColData(mCol发票金额) = 5
                            .ColData(mcol发票代码) = 5
                            .TextMatrix(.Row, mCol发票金额) = ""
                            .TextMatrix(.Row, mCol发票日期) = ""
                            .TextMatrix(.Row, mcol发票代码) = ""
                            .TextMatrix(.Row, .Col) = " "
                            .Text = " "
                        Else
                           .Text = .TextMatrix(.Row, .Col)
                           .ColData(mCol发票日期) = 2
                           .ColData(mcol发票代码) = 4
                           .ColData(mCol发票金额) = 4
                                    
                        End If
                    End If
                Else
                    If zlCommFun.ActualLen(.Text) > mint发票号Len Then
                        ShowMsgBox "发票号最多只能输入" & mint发票号Len & "个字符!"
                            Cancel = True
                        Exit Sub
                    End If

                    .ColData(mCol发票日期) = 2
                    .ColData(mcol发票代码) = 4
                    .ColData(mCol发票金额) = 4
                   
                    If mint编辑状态 = 8 Or (mint编辑状态 = 3 And mbln退货 = True) Then
                        .TextMatrix(.Row, mCol发票金额) = .TextMatrix(.Row, mCol结算金额)
                    Else
                        If mint记录状态 <> 1 Then
                            If Val(.TextMatrix(.Row, mCol发票金额)) = 0 Then
                                .TextMatrix(.Row, mCol发票金额) = .TextMatrix(.Row, mCol结算金额)
                            End If
                        Else
                            .TextMatrix(.Row, mCol发票金额) = .TextMatrix(.Row, mCol结算金额)
                        End If
                    End If
                End If
                显示合计金额
                Exit Sub
            Case mcol发票代码
                If Trim(.Text) = "" Then
                    If mcol发票代码 <> mintLastCol Then
                        .Col = GetNextEnableCol(mcol发票代码)
                        .Text = ""
                        Cancel = True
                        Exit Sub
                    End If
                Else
                    If zlCommFun.ActualLen(.Text) > 20 Then
                        ShowMsgBox "发票代码最多只能输入" & 20 & "个字符!"
                            Cancel = True
                        Exit Sub
                    End If
                End If
                Exit Sub
            Case mCol发票金额
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "发票金额必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Abs(Val(strKey)) < 0 Then
                        MsgBox "发票金额必须大于0,请重输！", vbInformation + vbOKOnly, gstrSysName
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
                    strKey = Format(Val(strKey), mFMT.FM_金额)
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
                显示合计金额
            Case mcol零售价
                '检查条件:
                ' 1.售价不能大于指导零售价(根据参数:不强制控制指导价格决定)
                ' 2.检查了结算价与售价
                If Val(.TextMatrix(.Row, 0)) = 0 Then Exit Sub
                If strKey <> "" Then
                    If Not IsNumeric(strKey) Then
                        ShowMsgBox "零售价必须为数字型，请重输！"
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    If Val(strKey) < 0 Then
                        ShowMsgBox "零售价必须大于等于0,请重输！"
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    If Val(strKey) >= 10 ^ 11 - 1 Then
                        ShowMsgBox "零售价必须小于" & (10 ^ 11 - 1)
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    If mbln不强制控制指导价格 = False Then
                        '判断输入的零售价与指导零售价
                        gstrSQL = "Select 指导零售价 From 材料特性 Where 材料ID=[1] "
                        Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "[读取指导零售价]", Val(.TextMatrix(.Row, 0)))
                        dbl指导零售价 = Val(zlStr.Nvl(rstemp!指导零售价))
                        dbl指导零售价 = Val(Format(dbl指导零售价, mFMT.FM_散装零售价))
                        If Val(Format(Val(strKey), mFMT.FM_散装零售价)) > dbl指导零售价 Then
                            ShowMsgBox "零售价不能大于指导零售价（指导零售价：￥" & dbl指导零售价 & "）"
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                    
                    If Val(.TextMatrix(.Row, mCol比例系数)) = 0 Then
                        dbl采购价 = Val(.TextMatrix(.Row, mCol结算价))
                    Else
                        dbl采购价 = Val(.TextMatrix(.Row, mCol结算价)) / Val(.TextMatrix(.Row, mCol比例系数))
                    End If
                    
                    If Val(strKey) < dbl采购价 Then
                        If MsgBox("注意：" & vbCrLf & "     零售价(￥" & Format(Val(strKey), mFMT.FM_散装零售价) & " 小于了" & vbCrLf & "     结算价（￥" & Format(dbl采购价, mFMT.FM_成本价) & "）,是否继续?", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        
                    End If
                End If
                
                If strKey <> "" Then
                    strKey = Format(Val(strKey), mFMT.FM_散装零售价)
                    .Text = strKey
                    .TextMatrix(.Row, .Col) = strKey
                ElseIf .TxtVisible = True Then
                    .Text = " "
                ElseIf .TxtVisible = False Then
                    If .TextMatrix(.Row, .Col) = "" Then
                        .Text = " "
                    Else
                        .Text = .TextMatrix(.Row, .Col)
                    End If
                End If
                '刘兴宏:零售价处理
                Call 计算零售价及零售差价(.Row, False)
                If strKey <> "" Then
                    .TextMatrix(.Row, mCol售价) = Format(Val(strKey) * Val(.TextMatrix(.Row, mCol比例系数)), mFMT.FM_零售价)
                    .TextMatrix(.Row, mCol售价金额) = Format(Val(.TextMatrix(.Row, mCol售价)) * Val(.TextMatrix(.Row, mCol数量)), mFMT.FM_金额)
                    .TextMatrix(.Row, mCol差价) = Format(Val(.TextMatrix(.Row, mCol售价金额)) - Val(.TextMatrix(.Row, mCol结算金额)), mFMT.FM_金额)



                End If
                显示合计金额
                
            Case mCol售价
                '检查条件:
                ' 1.售价不能大于指导零售价(根据参数:不强制控制指导价格决定)
                ' 2.检查了结算价与售价
                            
                If Val(.TextMatrix(.Row, 0)) = 0 Then Exit Sub
                If strKey <> "" Then
                    If Not IsNumeric(strKey) Then
                        ShowMsgBox "售价必须为数字型，请重输！"
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    If Val(strKey) < 0 Then
                        ShowMsgBox "售价必须大于等于0,请重输！"
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strKey) >= 10 ^ 11 - 1 Then
                        ShowMsgBox "售价必须小于" & (10 ^ 11 - 1)
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If mbln不强制控制指导价格 = False Then
                        '判断输入的零售价与指导零售价
                        gstrSQL = "Select 指导零售价 From 材料特性 Where 材料ID=[1] "
                        Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "[读取指导零售价]", Val(.TextMatrix(.Row, 0)))
                        dbl指导零售价 = Val(zlStr.Nvl(rstemp!指导零售价))
                        dbl指导零售价 = Val(Format(dbl指导零售价 * Val(.TextMatrix(.Row, mCol比例系数)), mFMT.FM_零售价))
                        
                        If Val(Format(Val(strKey), mFMT.FM_零售价)) > dbl指导零售价 Then
                            ShowMsgBox "售价不能大于指导零售价（指导零售价：￥" & dbl指导零售价 & "）"
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                    If Val(strKey) < Val(.TextMatrix(.Row, mCol结算价)) Then
                        If MsgBox("注意：" & vbCrLf & "     售价(￥" & Format(Val(strKey), mFMT.FM_零售价) & " 小于了" & vbCrLf & "     结算价（￥" & Format(Val(.TextMatrix(.Row, mCol结算价)), mFMT.FM_成本价) & "）,是否继续?", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        
                    End If
                End If
                
                If strKey <> "" Then
                    strKey = Format(Val(strKey), mFMT.FM_零售价)
                    .Text = strKey
                    .TextMatrix(.Row, .Col) = strKey
                ElseIf .TxtVisible = True Then
                    .Text = " "
                    .TextMatrix(.Row, .Col) = " "
                ElseIf .TxtVisible = False Then
                    If .TextMatrix(.Row, .Col) = "" Then
                        .Text = " "
                        .TextMatrix(.Row, .Col) = " "
                    Else
                        .Text = .TextMatrix(.Row, .Col)
                    End If
                End If
                If mbln时价购前销售 Then
                    dbl采购价 = Val(.TextMatrix(.Row, mCol采购价))
                Else
                    dbl采购价 = Val(.TextMatrix(.Row, mCol结算价))
                End If
                
                '重算差价
                If strKey <> "" Then
                    .TextMatrix(.Row, mcol加成率) = zlStr.FormatEx(计算加成率(Val(.TextMatrix(.Row, 0)), Val(strKey), dbl采购价), 2) & "%"
                    .TextMatrix(.Row, mCol售价金额) = Format(Val(strKey) * Val(.TextMatrix(.Row, mCol数量)), mFMT.FM_金额)
                    .TextMatrix(.Row, mCol差价) = Format(Val(.TextMatrix(.Row, mCol售价金额)) - Val(.TextMatrix(.Row, mCol结算金额)), mFMT.FM_金额)
                End If
                '刘兴宏:零售价处理
                Call 计算零售价及零售差价(.Row)
                显示合计金额

            Case mCol发票日期
                If strKey <> "" Then
                    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
                        strKey = TranNumToDate(strKey)
                        
                        If strKey = "" Then
                            MsgBox "效期必须为日期型！", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            Exit Sub
                        End If
                        .Text = strKey
                        Exit Sub
                    End If
                    If Not IsDate(strKey) Then
                        MsgBox "发票日期必须为日期型如(2000-10-10) 或 （20001010）,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                End If
            Case mCol随货单号
                If Trim(.Text) = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, .Col) = " "
                        .Text = " "
                    ElseIf .TxtVisible = False Then
                        If Trim(.TextMatrix(.Row, mCol随货单号)) = "" Then
                            .TextMatrix(.Row, .Col) = " "
                            .Text = " "
                        Else
                           .Text = .TextMatrix(.Row, .Col)
                        End If
                    End If
                Else
                    If zlCommFun.ActualLen(.Text) > 200 Then
                        ShowMsgBox "随货单号最多只能输入" & 200 & "个字符!"
                        Cancel = True
                        Exit Sub
                    End If
                End If
        End Select
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetNextEnableCol(ByVal intCurrCol As Integer) As Integer
    '返回下一个可见并可用的列号
    Dim n As Integer
    Dim intNextCol As Integer

    If intCurrCol > mshBill.Cols Or intCurrCol + 1 >= mintLastCol Then
        GetNextEnableCol = mintLastCol
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
End Function

'从材料目录中取值并附给相应的列
Public Function SetColValue(ByVal intRow As Integer, ByVal lng材料ID As Long, ByVal str诊疗 As String, ByVal str规格 As String, _
    ByVal str产地 As String, ByVal str单位 As String, ByVal num售价 As Double, _
    ByVal num指导批发价 As Double, ByVal str原产地 As String, ByVal int原效期 As Integer, _
    ByVal str简码 As String, ByVal num比例系数 As Double, ByVal lng批次 As Long, _
    ByVal int是否变价 As Integer, ByVal int在用分批 As Integer, ByVal dbl指导差价率 As Double, ByVal str批准文号 As String, ByVal str商品名 As String) As Boolean
    
    Dim intCount As Integer
    Dim intCol As Integer
    Dim rsprice As New Recordset
    Dim lngDepartid As Long
    Dim dblrate As Double, dbl差价让利比 As Double, dbl加成率 As Double
    Dim dbl成本价 As Double
    Dim rstemp As New ADODB.Recordset
    Dim int库房分批 As Integer
    Dim str散装单位 As String
    Dim bln虚拟库房 As Boolean
    Dim bln高值材料 As Boolean
    Dim bln跟踪病人 As Boolean
    Dim bln跟踪在用 As Boolean
    Dim bln在用分批 As Boolean
    Dim strMsg As String
    Dim sng分段售价 As Double
    Dim dbl定价加成率 As Double
    
    On Error GoTo ErrHandle
    SetColValue = False
    With mshBill
        For intCol = 0 To .Cols - 1
            If intCol <> mCol行号 Then .TextMatrix(intRow, intCol) = ""
        Next
        
        gstrSQL = "SELECT a.加成率 from 材料特性 a where a.材料id=[1]"
        Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "加成率", lng材料ID)
        dbl加成率 = Nvl(rstemp!加成率, 0) / 100
        
        gstrSQL = "select count(*) rec from 部门性质说明 where 部门id=[1] and 工作性质='虚拟库房'"
        Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "部门性质说明", cboStock.ItemData(cboStock.ListIndex))
        If rstemp!rec = 1 Then
            bln虚拟库房 = True
        End If
        rstemp.Close
        
        gstrSQL = "SELECT nvl(A.扣率,0) 扣率,nvl(a.加成率,0)/100 as 加成率,A.灭菌效期,A.一次性材料,A.成本价,A.库房分批,A.在用分批,A.注册证号,B.计算单位 散装单位" & _
                  ",Nvl(A.是否条码管理,0) As 条码管理, a.高值材料, a.跟踪病人, a.跟踪在用,a.注册证有效期 " & _
                  "From 材料特性 A,收费项目目录 B " & _
                  "Where a.材料ID=b.id and A.材料id=[1] "
        Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, "读取扣率", lng材料ID)
        
        dbl定价加成率 = Val(rsprice!加成率)
        
        If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
            bln高值材料 = (zlStr.Nvl(rsprice!高值材料, 0) = 1)
            bln跟踪病人 = (zlStr.Nvl(rsprice!跟踪病人, 0) = 1)
            bln跟踪在用 = (zlStr.Nvl(rsprice!跟踪在用, 0) = 1)
            bln在用分批 = (zlStr.Nvl(rsprice!在用分批, 0) = 1)
            
            strMsg = ""
            If bln虚拟库房 Then
                If bln高值材料 = False Then
                    strMsg = IIf(strMsg = "", "", strMsg & "、") & """高值材料"""
                End If
                If bln跟踪病人 = False Then
                    strMsg = IIf(strMsg = "", "", strMsg & "、") & """跟踪病人"""
                End If
                If bln跟踪在用 = False Then
                    strMsg = IIf(strMsg = "", "", strMsg & "、") & """跟踪在用"""
                End If
                If bln在用分批 = False Then
                    strMsg = IIf(strMsg = "", "", strMsg & "、") & """在用分批"""
                End If
                
                If strMsg <> "" Then
                    MsgBox "(" & str诊疗 & ")进行虚拟入库必须具备" & strMsg & "的属性。", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
        
        .TextMatrix(intRow, mcol注册证有效期) = IIf(IsNull(rsprice!注册证有效期), "", Format(rsprice!注册证有效期, "yyyy-mm-dd"))
        
        If rsprice!扣率 = 0 Then
            dblrate = 100
        Else
            dblrate = rsprice!扣率
        End If
        int库房分批 = Val(zlStr.Nvl(rsprice!库房分批))
        dbl成本价 = rsprice!成本价
        
        str散装单位 = zlStr.Nvl(rsprice!散装单位)
        
        .TextMatrix(intRow, mcol灭菌效期) = zlStr.Nvl(rsprice!灭菌效期, 0)
        .TextMatrix(intRow, mcol一次性材料) = zlStr.Nvl(rsprice!一次性材料, 0)
        .TextMatrix(intRow, mcol条码管理) = zlStr.Nvl(rsprice!条码管理, 0)
        
        .TextMatrix(intRow, mCol行号) = intRow
        .TextMatrix(intRow, 0) = lng材料ID
        .TextMatrix(intRow, mCol诊疗) = str诊疗
        .TextMatrix(intRow, mCol商品名) = str商品名
        .TextMatrix(intRow, mCol规格) = str规格
        
        If CheckQualifications(mlngModule, 1, IIf(IsNull(str产地), "", str产地)) = False Then
            .TextMatrix(intRow, mCol产地) = ""
        Else
            .TextMatrix(intRow, mCol产地) = IIf(IsNull(str产地), "", str产地)
        End If
        .TextMatrix(intRow, mcol批准文号) = IIf(IsNull(str批准文号), "", str批准文号)
        
        .TextMatrix(intRow, mCol单位) = str单位
        
        .TextMatrix(intRow, mCol售价) = Format(num售价, mFMT.FM_零售价)
        .TextMatrix(intRow, mCol指导批发价) = Format(num指导批发价, mFMT.FM_成本价)
        
        .TextMatrix(intRow, mCol原产地) = IIf(IsNull(str原产地), "", str原产地)
        .TextMatrix(intRow, mCol批次) = lng批次
        .TextMatrix(intRow, mcol注册证号) = zlStr.Nvl(rsprice!注册证号)   '取默认值
        
        
        '取出该材料的批号及效期
        If mint编辑状态 = 8 Or mbln退货 Then
            gstrSQL = "" & _
                " Select 上次批号 批号,效期,上次生产日期,上次采购价 From 药品库存" & _
                " Where 库房ID=[1] And 药品ID=[2]" & _
                "       And 性质=1 And nvl(批次,0)=[3]"
            Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex), lng材料ID, lng批次)
            
            If rsprice.RecordCount <> 0 Then
                .TextMatrix(intRow, mCol批号) = IIf(IsNull(rsprice!批号), "", rsprice!批号)
                .TextMatrix(intRow, mCol效期) = IIf(IsNull(rsprice!效期), "", rsprice!效期)
                If IsNull(rsprice!上次生产日期) Then
                    .TextMatrix(intRow, mcol生产日期) = ""
                Else
                    .TextMatrix(intRow, mcol生产日期) = Format(rsprice!上次生产日期, "yyyy-mm-dd")
                End If
                
                dbl成本价 = zlStr.Nvl(rsprice!上次采购价, 0)
                
                If dbl成本价 > 0 Then
                    .TextMatrix(intRow, mCol采购价) = Format(dbl成本价 * num比例系数, mFMT.FM_成本价)
                    .TextMatrix(intRow, mCol结算价) = Format(dbl成本价 * num比例系数 * dblrate / 100, mFMT.FM_成本价)
                End If
            End If
        End If
        
        '原效期字段下面保存原效期，指导差价，是否变价，在用分批等，格式为：最大效期||指导差价率||是否变价||在用分批||库房分批
        .TextMatrix(intRow, mCol原销期) = IIf(IsNull(int原效期), "0", int原效期) & "||" & dbl指导差价率 & "||" & int是否变价 & "||" & int在用分批 & "||" & int库房分批
        
        .TextMatrix(intRow, mCol简码) = str简码
        .TextMatrix(intRow, mCol比例系数) = num比例系数
        If intRow > 1 Then
            .TextMatrix(intRow, mCol随货单号) = .TextMatrix(intRow - 1, mCol随货单号)
            .TextMatrix(intRow, mCol发票号) = .TextMatrix(intRow - 1, mCol发票号)
            .TextMatrix(intRow, mcol发票代码) = .TextMatrix(intRow - 1, mcol发票代码)
            .TextMatrix(intRow, mCol发票日期) = .TextMatrix(intRow - 1, mCol发票日期)
        End If
        
        SetInputFormat intRow
        SetDisCount intRow, dblrate
        lngDepartid = cboStock.ItemData(cboStock.ListIndex)
        
        '说明：这里区分分批核算和不分批核算的目的是提高运行速度。
        '本来可以不分这些，直接用第一条SQL语句实现，但不分批的卫材就多在数据库中扫描一次。
        
        If Not (mint编辑状态 = 8 Or mbln退货) Then
            '对定价采购，不用取上次的结算价和扣率
                        
            '存储格式值:最大效期||指导差价率||是否变价||在用分批||库房分批
            If Val(Split(.TextMatrix(intRow, mCol原销期), "||")(4)) > 0 Then
                gstrSQL = "" & _
                    "   Select 上次采购价,上次产地,上次生产日期 " & _
                    "   From 药品库存 " & _
                    "   where 性质=1 and 库房id=[1] and 药品id=" & lng材料ID & _
                    "       and nvl(批次,0) =(  Select max(nvl(批次,0)) " & _
                    "                            From 药品库存 " & _
                    "                           Where 性质=1 and 库房id=[1]" & _
                    "                               and 药品id=[2] )"
            Else
                gstrSQL = "select 上次采购价,上次产地,上次生产日期 from 药品库存 where 性质=1 and 库房id= [1] and 药品id=[2]"
            End If
            Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lngDepartid, lng材料ID)
            
            If Not rsprice.EOF Then
                If .TextMatrix(intRow, mCol产地) = "" Then
                    If CheckQualifications(mlngModule, 1, IIf(IsNull(rsprice.Fields(1)), "", rsprice.Fields(1))) = False Then
                        .TextMatrix(intRow, mCol产地) = ""
                    Else
                        .TextMatrix(intRow, mCol产地) = IIf(IsNull(rsprice.Fields(1)), "", rsprice.Fields(1))
                    End If
                End If
                If IsNull(rsprice!上次生产日期) Then
                    .TextMatrix(intRow, mcol生产日期) = ""
                Else
                    .TextMatrix(intRow, mcol生产日期) = Format(rsprice!上次生产日期, "yyyy-mm-dd")
                    If ProduceDateCheck(.TextMatrix(intRow, mcol生产日期)) = False Then
                        '不合格的数据清除掉
                        For intCol = 0 To .Cols - 1
                            .TextMatrix(intRow, intCol) = ""
                        Next
                        .Row = intRow
                        .Col = mCol诊疗
                        Exit Function
                    End If
                End If
                If Val(zlStr.Nvl(rsprice.Fields(0))) = 0 Then
                    .TextMatrix(intRow, mCol采购价) = Format(dbl成本价 * num比例系数, mFMT.FM_成本价)
                    .TextMatrix(intRow, mCol结算价) = Format(dbl成本价 * num比例系数 * dblrate / 100, mFMT.FM_成本价)
                Else
                    .TextMatrix(intRow, mCol采购价) = Format(Val(zlStr.Nvl(rsprice.Fields(0)) * num比例系数), mFMT.FM_成本价)
                    .TextMatrix(intRow, mCol结算价) = Format(Val(zlStr.Nvl(rsprice.Fields(0)) * num比例系数) * dblrate / 100, mFMT.FM_成本价)
                End If
            Else
                If dbl成本价 > 0 Then
                    .TextMatrix(intRow, mCol采购价) = Format(dbl成本价 * num比例系数, mFMT.FM_成本价)
                    .TextMatrix(intRow, mCol结算价) = Format(dbl成本价 * num比例系数 * dblrate / 100, mFMT.FM_成本价)
                End If
            End If
        End If
        
        If .TextMatrix(intRow, mCol产地) <> "" Then
            gstrSQL = "select 批准文号 from 药品生产商对照 where 厂家名称=[1] and 药品id=[2]"
            Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "mshbill_CommandClick", .TextMatrix(mshBill.Row, mCol产地), lng材料ID)
            If Not rstemp.EOF Then
               .TextMatrix(intRow, mcol批准文号) = IIf(IsNull(rstemp!批准文号), "", rstemp!批准文号)
            End If
        End If
        
        Dim dbl结算价 As Double, dbl采购价 As Double
        
        dbl结算价 = Val(.TextMatrix(intRow, mCol结算价))
        dbl采购价 = Val(.TextMatrix(intRow, mCol采购价))
        dbl成本价 = IIf(mbln时价购前销售, dbl采购价, dbl结算价)
        '时价材料处理
        If int是否变价 = 1 Then
            If mint编辑状态 = 8 Or mbln退货 Then
                gstrSQL = "" & _
                "   Select 实际金额/实际数量*" & num比例系数 & " as  售价 " & _
                "   From 药品库存 " & _
                "   Where 库房id=[1]" & _
                "           and 药品id=[2]" & _
                "           and 性质=1 and 实际数量>0 and " & _
                "           nvl(批次,0)=[3]"
                Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex), lng材料ID, lng批次)
                If rstemp.EOF Then
                    MsgBox "时价材料没有库存，不能出库，请检查！", vbOKOnly, gstrSysName
                    Exit Function
                End If
               .TextMatrix(intRow, mCol售价) = Format(Nvl(rstemp!售价, 0), mFMT.FM_零售价)
'               .TextMatrix(intRow, mcol加成率) = Format(Val(.TextMatrix(intRow, mCol售价)) / Val(), "###0.00") & "%"
            Else
                If mbln加价率 = True Then
                    '由于存在差价让利比的存在,需要按加成率计算,因此将指导差价率转换成加成率计算 公式：加成率=1/(1-差价率)-1
                    .TextMatrix(intRow, mCol售价) = Format(校正零售价(dbl成本价 * (1 + dbl加成率) + _
                                                    时价材料零售价(lng材料ID, dbl成本价, dbl加成率)) _
                                                    , mFMT.FM_零售价)
                    .TextMatrix(intRow, mcol加成率) = Format(dbl加成率 * 100, "###0.00") & "%"
                ElseIf mbln分段加成率 = True Then
                    dbl加成率 = 0
                    If Get分段加成售价(dbl成本价, Val(.TextMatrix(intRow, mCol比例系数)), mstrCaption, sng分段售价) = False Then
                        .TextMatrix(intRow, mCol售价) = Format(校正零售价(dbl成本价 * (1 + dbl加成率) + _
                                                        时价材料零售价(lng材料ID, dbl成本价, dbl加成率)) _
                                                        , mFMT.FM_零售价)
                        .TextMatrix(intRow, mcol加成率) = Format(dbl加成率 * 100, "###0.00") & "%"
                    Else
                        .TextMatrix(intRow, mCol售价) = Format(校正零售价(sng分段售价 + _
                                                      时价材料零售价(lng材料ID, dbl成本价, dbl加成率, -1, sng分段售价)) _
                                                      , mFMT.FM_零售价)
                        .TextMatrix(intRow, mcol加成率) = Format(dbl加成率 * 100, "####0.00") & "%" '因为是分段加成的所以加成率不准确，取一个模糊值即可
                    End If
                Else '取上次售价模式和没有勾选任何取售价方式
                    If mbln时价卫材取上次售价 = True Then
                        gstrSQL = "Select Nvl(上次售价, 0) As 上次售价 From 材料特性 Where 材料id = [1]"
                        Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng材料ID)
                        If rstemp!上次售价 > 0 Then
                            .TextMatrix(intRow, mCol售价) = Format(zlStr.Nvl(rstemp!上次售价, 0) * num比例系数, mFMT.FM_零售价)
                            If dbl成本价 <> 0 Then
                                .TextMatrix(intRow, mcol加成率) = Format((Val(.TextMatrix(intRow, mCol售价)) / dbl成本价 - 1) * 100, "###0.00") & "%"
                            End If
                        Else
                            '由于存在差价让利比的存在,需要按加成率计算,因此将指导差价率转换成加成率计算 公式：加成率=1/(1-差价率)-1
                            .TextMatrix(intRow, mCol售价) = Format(校正零售价(dbl成本价 * (1 + dbl加成率) + _
                                                            时价材料零售价(lng材料ID, dbl成本价, dbl加成率)) _
                                                            , mFMT.FM_零售价)
                            .TextMatrix(intRow, mcol加成率) = Format(dbl加成率 * 100, "###0.00") & "%"
                        End If
                    Else
                        '由于存在差价让利比的存在,需要按加成率计算,因此将指导差价率转换成加成率计算 公式：加成率=1/(1-差价率)-1
                        .TextMatrix(intRow, mCol售价) = Format(校正零售价(dbl成本价 * (1 + dbl加成率) + _
                                                        时价材料零售价(lng材料ID, dbl成本价, dbl加成率)) _
                                                        , mFMT.FM_零售价)
                        .TextMatrix(intRow, mcol加成率) = Format(dbl加成率 * 100, "###0.00") & "%"
                    End If
                End If
            End If
        Else
            .TextMatrix(intRow, mcol加成率) = Format(dbl定价加成率 * 100, "###0.00") & "%" '"15.00%"'定价取规格中加成率
        End If
        .TextMatrix(intRow, mcol零售单位) = str散装单位
        '刘兴宏:零售价处理
        Call 计算零售价及零售差价(intRow)
        mshBill.MsfObj.CellForeColor = IIf(int是否变价 = 0, &H0, &H40&)     ' &H40C0&
        
        If mstr验收结论 = "" Then
            gstrSQL = "Select 名称  From 入库验收结论 where 缺省标志=1"
            Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "SetColValue")
            
            If Not rstemp.EOF Then
                .TextMatrix(intRow, mCol验收结论) = IIf(IsNull(rstemp!名称), "", rstemp!名称)
                mstr验收结论 = rstemp!名称
            End If
        Else
            .TextMatrix(intRow, mCol验收结论) = mstr验收结论
        End If
    End With
    Call 提示库存数
    SetColValue = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetInputFormat(ByVal intRow As Integer)
    '--------------------------------------------------------------------------------------------------------
    '功能:设置当前行的编辑格式
    '参数:introw-当前行
    '返回:
    '编制:刘兴宏
    '日期:2007/05/15
    '--------------------------------------------------------------------------------------------------------
    
    With mshBill
    
        '1.新增；2、修改；3、验收；4、查看；5、修改发票；6、冲销；
        '7、财务审核（冲销、产生新单据并审核；已付款的单据不允许财务审核，同样，财务审核后的单据不允许冲销）;
        '8、卫材库退货,9-核查
        If mint编辑状态 = 9 Or mint编辑状态 = 3 Or mint编辑状态 = 7 Then
            '刘兴宏:2007/05/30:增加流程控制时编制相关的属性
            Call Set操作流程Update(True)
        End If
        If mint编辑状态 = 7 Then
                '如果是时价卫材，则允许输入售价
                '存储格式值:最大效期||指导差价率||是否变价||在用分批||库房分批
                If Split(.TextMatrix(intRow, mCol原销期) & "||||||||", "||")(2) = 1 Then
                    If Split(.TextMatrix(intRow, mCol原销期) & "||||||||", "||")(4) = 1 Then
                        .ColData(mcol零售价) = IIf(mbln退货 Or mint编辑状态 = 8, 5, 4)
                    Else
                        .ColData(mcol零售价) = 5
                    End If
                Else
                    .ColData(mcol零售价) = 5
                End If
        End If
        If mblnEdit = False Then Exit Sub
        
        If mint编辑状态 = 7 Or mint编辑状态 = 8 Or mbln退货 Then Exit Sub
'        If .TextMatrix(intRow, mCol原产地) = "!" Then
            .ColData(mCol产地) = 1              '纯文本输入
'        Else
'            .ColData(mCol产地) = 5              '禁止
'        End If

        If .TextMatrix(intRow, mcol一次性材料) = "1" Then
            .ColData(mcol灭菌日期) = 2
            .ColData(mcol灭菌失效期) = 2
        Else
            .ColData(mcol灭菌日期) = 5              '禁止
            .ColData(mcol灭菌失效期) = 5
        End If
        
        If .TextMatrix(intRow, mcol条码管理) = "1" Then
            .ColData(mcol商品条码) = 4
        Else
            .ColData(mcol商品条码) = 5
        End If

        .ColData(mCol效期) = 2

        '存储格式值:最大效期||指导差价率||是否变价||在用分批||库房分批
        If .TextMatrix(intRow, mCol原销期) <> "" Then
            If mint编辑状态 <> 9 And mint编辑状态 <> 3 And mint编辑状态 <> 7 Then
                '如果是时价卫材，则允许输入售价
                '存储格式值:最大效期||指导差价率||是否变价||在用分批||库房分批
                If Split(.TextMatrix(intRow, mCol原销期), "||")(2) = 1 Then
                    If Split(.TextMatrix(intRow, mCol原销期), "||")(4) = 1 Then
                        .ColData(mcol零售价) = IIf(mbln退货 Or mint编辑状态 = 8, 5, 4)
                    Else
                        .ColData(mcol零售价) = 5
                    End If
                    .ColData(mCol售价) = IIf(mbln时价卫材直接确定售价, 4, 5)
                Else
                    .ColData(mCol售价) = 5
                    .ColData(mcol零售价) = 5
                End If
            Else
                '核查\审核\财务审核已经设置了该售价的输入
                '20070530:刘兴宏
            End If
            
        Else
            .ColData(mCol售价) = 5
            .ColData(mcol零售价) = 5
        End If
        
        If Trim(.TextMatrix(intRow, mCol发票号)) = "" Then
            .ColData(mCol发票日期) = 5
            .ColData(mCol发票金额) = 5
            .ColData(mcol发票代码) = 5
        Else
            .ColData(mCol发票日期) = 2
            .ColData(mcol发票代码) = 4
            .ColData(mCol发票金额) = 4
        End If
        
    End With
End Sub


'设置折扣
Private Sub SetDisCount(ByVal intRow As Integer, ByVal intDisCount As Double)
    Dim dbl加成率 As Double, dbl售价 As Double, dbl采购价 As Double, dbl数量 As Double, dbl结算价 As Double
    Dim lng材料ID  As Long
    Dim bln是否计算 As Boolean
    
    With mshBill
        dbl售价 = Val(.TextMatrix(intRow, mCol售价))
        dbl结算价 = Val(.TextMatrix(intRow, mCol结算价))
        dbl采购价 = Val(.TextMatrix(intRow, mCol采购价))
        lng材料ID = Val(.TextMatrix(intRow, 0))
        
        If mbln加价率 Then
            mdbl加价率 = 15
            If mbln时价购前销售 Then
                If dbl售价 <> 0 And dbl采购价 <> 0 Then
                    mdbl加价率 = 计算加成率(lng材料ID, dbl售价, dbl采购价)
                    bln是否计算 = True
                End If
            Else
                If dbl售价 <> 0 And dbl结算价 <> 0 Then
                    mdbl加价率 = 计算加成率(lng材料ID, dbl售价, dbl结算价)
                    bln是否计算 = True
                End If
            End If
        End If
        
        If mshBill.Col = mCol指导批发价 Then
            .Text = Format(intDisCount, mFMT.FM_金额)
            
            .TextMatrix(intRow, mCol指导批发价) = .Text
            If Val(.TextMatrix(intRow, mCol采购价)) = 0 Then
                .TextMatrix(intRow, mCol采购价) = .Text
                dbl采购价 = Val(.TextMatrix(intRow, mCol采购价))
            End If
            intDisCount = Val(.TextMatrix(intRow, mCol扣率))
        Else
            .TextMatrix(intRow, mCol扣率) = intDisCount
        End If
        
        If .TextMatrix(intRow, mCol指导批发价) <> "" Then
            If .TextMatrix(intRow, mCol采购价) = "" Then
                .TextMatrix(intRow, mCol采购价) = .TextMatrix(intRow, mCol指导批发价)
                dbl采购价 = Val(.TextMatrix(intRow, mCol采购价))
            End If
            If Not (mint编辑状态 = 8 Or mbln退货) Then
                .TextMatrix(intRow, mCol结算价) = Format((Val(.TextMatrix(intRow, mCol采购价)) * intDisCount / 100), mFMT.FM_成本价)
            End If
            dbl结算价 = Val(.TextMatrix(intRow, mCol结算价))
            If .TextMatrix(intRow, mCol数量) <> "" Then
               .TextMatrix(intRow, mCol结算金额) = Format((Val(.TextMatrix(intRow, mCol数量)) * Val(.TextMatrix(intRow, mCol结算价))), mFMT.FM_金额)
               .TextMatrix(intRow, mCol发票金额) = IIf(Trim(.TextMatrix(intRow, mCol发票号)) = "" And Trim(.TextMatrix(intRow, mcol发票代码)) = "", "", .TextMatrix(intRow, mCol结算金额))
            End If
            .TextMatrix(intRow, mCol差价) = Format(IIf(.TextMatrix(intRow, mCol售价金额) = "", 0, .TextMatrix(intRow, mCol售价金额)) - IIf(.TextMatrix(intRow, mCol结算金额) = "", 0, .TextMatrix(intRow, mCol结算金额)), mFMT.FM_金额)
            '对时价卫材的处理
            '存储格式值:最大效期||指导差价率||是否变价||在用分批||库房分批
            If .TextMatrix(intRow, mCol原销期) <> "" Then
                If Split(.TextMatrix(intRow, mCol原销期), "||")(2) = 1 Then
                    '由于存在差价让利比的存在,需要按加成率计算,因此将指导差价率转换成加成率计算 公式：加成率=1/(1-差价率)-1
                    If mbln加价率 Then
                        If bln是否计算 Then
                            .TextMatrix(intRow, mCol售价) = Format(校正零售价(IIf(mbln时价购前销售, dbl采购价, dbl结算价) * (1 + (mdbl加价率 / 100)) + _
                                时价材料零售价(lng材料ID, IIf(mbln时价购前销售, dbl采购价, dbl结算价), (mdbl加价率 / 100))), mFMT.FM_零售价)
                        Else
                            dbl加成率 = Val(Replace(.TextMatrix(intRow, mcol加成率), "%", "")) / 100
                            .TextMatrix(intRow, mCol售价) = Format(校正零售价(IIf(mbln时价购前销售, dbl采购价, dbl结算价) * (1 + dbl加成率) + _
                                时价材料零售价(lng材料ID, IIf(mbln时价购前销售, dbl采购价, dbl结算价), dbl加成率)), mFMT.FM_零售价)
                        End If
                    Else
                        dbl加成率 = Val(Replace(.TextMatrix(intRow, mcol加成率), "%", "")) / 100
                        .TextMatrix(intRow, mCol售价) = Format(校正零售价(IIf(mbln时价购前销售, dbl采购价, dbl结算价) * (1 + dbl加成率) + _
                            时价材料零售价(lng材料ID, IIf(mbln时价购前销售, dbl采购价, dbl结算价), dbl加成率)), mFMT.FM_零售价)
                    End If
                    If .TextMatrix(intRow, mCol数量) <> "" Then
                        .TextMatrix(intRow, mCol售价金额) = Format(.TextMatrix(intRow, mCol数量) * Val(.TextMatrix(intRow, mCol售价)), mFMT.FM_金额)
                        .TextMatrix(intRow, mCol差价) = Format(IIf(.TextMatrix(intRow, mCol售价金额) = "", 0, .TextMatrix(intRow, mCol售价金额)) - IIf(.TextMatrix(intRow, mCol结算金额) = "", 0, .TextMatrix(intRow, mCol结算金额)), mFMT.FM_金额)
                    End If
                End If
            End If
            Call 计算零售价及零售差价(intRow)
        End If
    End With
End Sub

Private Sub mshBill_LeaveCell(Row As Long, Col As Long)
     ImeLanguage False
     
     With mshBill
        If .Col = mcol生产日期 And .TextMatrix(.Row, mcol生产日期) <> "" Then
            If ProduceDateCheck(.TextMatrix(.Row, mcol生产日期)) = False Then
                .TextMatrix(.Row, mcol生产日期) = ""
                Exit Sub
            End If
        End If
    End With
End Sub

Private Sub mshBill_LostFocus()
     ImeLanguage False
End Sub

Private Sub mshBill_Validate(Cancel As Boolean)
    mshBill.LastRow = 0
End Sub

Private Sub mshProvider_DblClick()
    mshProvider_KeyDown vbKeyReturn, 0
End Sub

Private Sub mshProvider_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        mshProvider.Visible = False
        txtProvider.SetFocus
        txtProvider.SelStart = 0
        txtProvider.SelLength = Len(txtProvider.Text)
    End If
    
    If KeyCode = vbKeyReturn Then
        txtProvider.Text = mshProvider.TextMatrix(mshProvider.Row, 2)
        txtProvider.Tag = mshProvider.TextMatrix(mshProvider.Row, 0)
        mshProvider.Visible = False
        mshBill.SetFocus
    End If
    
    If CheckQualifications(mlngModule, 2, Val(txtProvider.Tag)) = False Then
        txtProvider.Text = ""
        txtProvider.Tag = "0"
        Exit Sub
    End If

    If Val(txtProvider.Tag) <> mlng供货单位ID And (mint编辑状态 = 8 Or mbln退货) Then
        mshBill.ClearBill
        mlng供货单位ID = Val(txtProvider.Tag)
        mshBill.TextMatrix(1, mCol行号) = "1"
    End If
End Sub

Private Sub mshProvider_LostFocus()
    If mshProvider.Visible Then
        mshProvider.Visible = False
    End If
End Sub

Private Sub msh产地_DblClick()
    msh产地_KeyDown vbKeyReturn, 0
End Sub

Private Sub msh产地_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsProvider As ADODB.Recordset
    
    With mshBill
    
        If KeyCode = vbKeyEscape Then
            msh产地.Visible = False
            .SetFocus
        End If
        
        If .Col = mCol验收结论 And KeyCode = vbKeyReturn Then
            .TextMatrix(.Row, .Col) = msh产地.TextMatrix(msh产地.Row, 1)
            .Text = .TextMatrix(.Row, .Col)
            msh产地.Visible = False
            .SetFocus
            Call ColMoveNextCol(.Col)
            Exit Sub
        End If
        
        If CheckQualifications(mlngModule, 1, msh产地.TextMatrix(msh产地.Row, 2)) = False Then
            Exit Sub
        End If
        
        If KeyCode = vbKeyReturn Then
            .TextMatrix(.Row, .Col) = msh产地.TextMatrix(msh产地.Row, 2)
            msh产地.Visible = False
            
            gstrSQL = "select 批准文号 from 药品生产商对照 where 厂家名称=[1] and 药品id=[2]"
            Set rsProvider = zlDatabase.OpenSQLRecord(gstrSQL, "msh产地_KeyDown", .TextMatrix(.Row, .Col), .TextMatrix(.Row, 0))
            If rsProvider.RecordCount > 0 Then
                .TextMatrix(.Row, mcol批准文号) = IIf(IsNull(rsProvider!批准文号), "", rsProvider!批准文号)
            Else
                .TextMatrix(.Row, mcol批准文号) = ""
            End If
            
            .Col = mCol批号
            .SetFocus
        End If
    
    End With
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
            If mshBill.Col = mCol结算价 Or mshBill.Col = mCol结算金额 Then Exit Sub
        End If
    End If
    PicInput.Visible = False
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "PY" And stbThis.Tag <> "PY" Then
        Logogram stbThis, 0
        stbThis.Tag = Panel.Key
    ElseIf Panel.Key = "WB" And stbThis.Tag <> "WB" Then
        Logogram stbThis, 1
        stbThis.Tag = Panel.Key
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


Private Sub txtCopy_Change()
    txtCopy.Text = Val(txtCopy.Text)
    If Val(txtCopy.Text) > 9999 Then txtCopy.Text = 9999
End Sub

Private Sub txtCopy_KeyPress(KeyAscii As Integer)
    If Not (Chr(KeyAscii) >= 0 And Chr(KeyAscii) <= 9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txtDrawPerson_Change()
    mblnChange = True
    txtDrawPerson.Tag = ""
End Sub

Private Sub txtDrawPerson_GotFocus()
    OS.OpenIme False
    zlControl.TxtSelAll txtDrawPerson
End Sub

Private Sub txtDrawPerson_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If txtDrawPerson.Tag <> "" Then OS.PressKey vbKeyTab: Exit Sub
    If SelectItem(txtDrawPerson, Trim(txtDrawPerson.Text), True) = False Then Exit Sub
End Sub

Private Sub txtNO_Change()
    If txtNO.Locked = True Then
'        If mstr单据号 <> "" And mstr单据号 <> txtNO.Text Then
'            txtNO.Text = mstr单据号
'        End If
    End If
End Sub

Private Sub TxtNo_GotFocus()
    If txtNO.Locked = False Then
        txtNO.SelStart = 0
        txtNO.SelLength = Len(txtNO.Text)
    End If
End Sub

Private Sub TxtNo_KeyDown(KeyCode As Integer, Shift As Integer)
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
    mblnChange = True
End Sub

Private Sub txtProvider_GotFocus()
    txtProvider.SelStart = 0
    txtProvider.SelLength = Len(txtProvider.Text)
End Sub

Private Sub txtProvider_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strProviderText As String
    Dim adoProvider As New Recordset
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If mint编辑状态 = 3 Or mint编辑状态 = 4 Then Exit Sub
    
    On Error GoTo ErrHandle
    With txtProvider
        If Trim(.Text) = "" Then Exit Sub
        strProviderText = GetMatchingSting(UCase(.Text))
        
        
        gstrSQL = "" & _
            "   Select id,编码,名称,简码 " & _
            "   From 供应商 " & _
            "   Where (To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01' or 撤档时间 is null) " & _
            "       And (站点=[2] or 站点 is null) And 末级=1 And (substr(类型,5,1)=1 ) " & _
            "       And (简码 like [1] Or 编码 like [1] or 名称 like [1]) "
        Set adoProvider = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, strProviderText, gstrNodeNo)
        
        If adoProvider.EOF Then
            MsgBox "没有你输入的供货单位，请重输！", vbOKOnly + vbInformation, gstrSysName
            KeyCode = 0
            .SelStart = 0
            .SelLength = Len(.Text)
            .Tag = 0
            Exit Sub
        End If
        
        If adoProvider.RecordCount > 1 Then
            Set mshProvider.Recordset = adoProvider
            Dim intCol As Integer
            Dim intRow As Integer
            
            With mshProvider
                If .Visible = False Then .Visible = True
                .Redraw = False
                .SetFocus
                
                For intRow = 0 To .Rows - 1
                    .Row = intRow
                    For intCol = 0 To .Cols - 1
                        .Col = intCol
                        If .Row = 0 Then
                            .CellFontBold = True
                        Else
                            .CellFontBold = False
                        End If
                    Next
                Next
                .Font.Bold = False
                .FontFixed.Bold = True
                .ColWidth(0) = 0
                .ColWidth(1) = 1000
                .ColWidth(2) = 2700
                .ColWidth(3) = 1200
                .Row = 1
                .TopRow = 1
                .Col = 0
                .ColSel = .Cols - 1
                
                .Top = txtProvider.Top + txtProvider.Height
                .Left = cmdProvider.Left + cmdProvider.Width - .Width
                .Redraw = True
                Exit Sub
            End With
        Else
            .Text = adoProvider!名称
            .Tag = adoProvider!Id
        End If
        adoProvider.Close
        mshBill.SetFocus
        mshBill.Col = 1
        mshBill.Row = 1
        
        If CheckQualifications(mlngModule, 2, Val(txtProvider.Tag)) = False Then
            txtProvider.Text = ""
            txtProvider.Tag = "0"
            Exit Sub
        End If
        
        If Val(.Tag) <> mlng供货单位ID And mint编辑状态 = 8 Then
            mlng供货单位ID = Val(txtProvider.Tag)
            mshBill.ClearBill
            mshBill.TextMatrix(1, mCol行号) = "1"
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Function ValidData() As Boolean
    Dim intLop As Integer
    Dim rsStock As New Recordset
    Dim blnStock As Boolean
    Dim bln高值卫材录入 As Boolean
    Dim strNo As String
    Dim bln高值卫材 As Boolean
    
    On Error GoTo ErrHandle
    ValidData = False
    
    gstrSQL = "" & _
        "   SELECT count(*)" & _
        "   From 部门性质说明 " & _
        "   WHERE ((工作性质 LIKE '发料部门') OR (工作性质 LIKE '制剂室')) " & _
        "           AND 部门id =[1]"
        
    Set rsStock = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex))
    
    If rsStock.Fields(0) > 0 Then
        blnStock = False
    Else
        blnStock = True
    End If
    
    If txtNO.Locked = False Then
        '新增，且允许修改单据号
        strNo = txtNO.Text
        If strNo = "" Then
            ShowMsgBox "请输入单据号。"
            txtNO.SetFocus
            Exit Function
        End If
        
        If InStr(strNo, "'") > 0 Then
            ShowMsgBox "单据号所输内容中含有非法字符。"
            txtNO.SetFocus
            Exit Function
        End If
        
        If LenB(StrConv(strNo, vbFromUnicode)) > 8 Then
            ShowMsgBox "单据号长度不能超过8个字母。"
            txtNO.SetFocus
            Exit Function
        End If
    Else
'        '防止用户强制修改
'        If mstr单据号 <> "" And mstr单据号 <> txtNO.Text Then
'            txtNO.Text = mstr单据号
'        End If
    End If
    
    If CheckQualifications(mlngModule, 2, Val(txtProvider.Tag)) = False Then Exit Function
    
    bln高值卫材录入 = IIf(Val(zlDatabase.GetPara("高值卫材必须填写详细信息", glngSys, mlngModule)) = 1, True, False)
    
    With mshBill
        If .TextMatrix(1, 0) <> "" Then         '先判有否数据
            If Val(txtProvider.Tag) = 0 Then
                ShowMsgBox "供货单位不能为空！"
                txtProvider.SetFocus
                Exit Function
            End If
            
            '29679 处理办法
            If Val(.TextMatrix(.Row, mCol数量)) < 0 Then
                ShowMsgBox "数量不能小于0！"
                .Col = mCol数量
                .SetFocus
                Exit Function
            End If
            
            If LenB(StrConv(txt摘要.Text, vbFromUnicode)) > txt摘要.MaxLength Then
                ShowMsgBox "摘要超长,最多能输入" & CInt(txt摘要.MaxLength / 2) & "个汉字或" & txt摘要.MaxLength & "个字符!"
                txt摘要.SetFocus
                Exit Function
            End If
        
            For intLop = 1 To .Rows - 1
                If Trim(.TextMatrix(intLop, mCol诊疗)) <> "" Then
                    If Trim(Trim(.TextMatrix(intLop, mCol数量))) = "" Then
                        ShowMsgBox "第" & intLop & "行卫生材料的数量为空了，请检查！"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mCol数量
                        Exit Function
                    End If
                    
'                    If Val(Trim(.TextMatrix(intLop, mCol结算价))) = 0 Then
'                        ShowMsgbox "第" & intLop & "行卫生材料的结算价为空了，请检查！"
'                        .SetFocus
'                        .Row = intLop
'                        .MsfObj.TopRow = intLop
'                        .Col = mCol结算价
'                        Exit Function
'                    End If
                    If mbln不强制控制指导价格 = False Then
                        If Val(.TextMatrix(intLop, mCol指导批发价)) < 0 Then
                            ShowMsgBox "第" & intLop & "行卫生材料的采购限价必需大于等于0，请检查！"
                            .SetFocus
                            .Row = intLop
                            .MsfObj.TopRow = intLop
                            .Col = mCol结算价
                            Exit Function
                        End If
                    End If
                    If LenB(StrConv(Trim(Trim(.TextMatrix(intLop, mCol批号))), vbFromUnicode)) > mintBatchNoLen Then
                        ShowMsgBox "第" & intLop & "行卫生材料的批号超长,最多能输入" & Int(mintBatchNoLen / 2) & "个汉字或" & mintBatchNoLen & "个字符!"
                        .SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mCol批号
                        Exit Function
                    End If
                    
                    If LenB(StrConv(Trim(Trim(.TextMatrix(intLop, mcol注册证号))), vbFromUnicode)) > 50 Then
                        ShowMsgBox "第" & intLop & "行卫生材料的注册证号超长,最多能输入25个汉字或50个字符!"
                        .SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mcol注册证号
                        Exit Function
                    End If
                    
                    If Len(Trim(.TextMatrix(intLop, mcol商品条码))) > 50 Then
                        ShowMsgBox "第" & intLop & "行卫生材料的商品条码超长,最多能输入50个字符!"
                        .SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mcol商品条码
                        Exit Function
                    End If
                    
                    If Trim(Trim(.TextMatrix(intLop, mCol扣率))) = "" Then
                        ShowMsgBox "第" & intLop & "行卫生材料的扣率为空了，请检查！"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mCol扣率
                        Exit Function
                    End If
                    
                    If Val(Trim(Trim(.TextMatrix(intLop, mCol扣率)))) >= 1000# Then
                        ShowMsgBox "第" & intLop & "行卫生材料的扣率太大了，请检查！"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mCol扣率
                        Exit Function
                    End If
                    
                    If blnStock = True Then
                        '存储格式值:最大效期||指导差价率||是否变价||在用分批||库房分批
                        If Split(.TextMatrix(intLop, mCol原销期), "||")(0) <> "0" Then
                                        
                            If Trim(.TextMatrix(intLop, mCol批号)) = "" Or Trim(.TextMatrix(intLop, mCol效期)) = "" Then
                                ShowMsgBox "第" & intLop & "行的卫生材料是效期材料,请把它的批号及效期" & vbCrLf & "信息完整输入单据中！"
                                mshBill.SetFocus
                                .Row = intLop
                                .MsfObj.TopRow = intLop
                                If .TextMatrix(intLop, mCol批号) = "" Then
                                    .Col = mCol批号
                                Else
                                    .Col = mCol效期
                                End If
                                Exit Function
                            End If
                        End If
                        
                        '分批药品必须录入产地和批号
                        If mbln分批卫材批号产地控制 = True And Not (mint编辑状态 = 8 Or mbln退货 = True) Then '退货不检查
                            '分批药品必须录入产地和批号
                            If Split(.TextMatrix(intLop, mCol原销期), "||")(4) <> "0" And (.TextMatrix(intLop, mCol产地) = "" Or .TextMatrix(intLop, mCol批号) = "") Then
                                MsgBox "第" & intLop & "行的卫材是分批卫材,请把它的产地和批号" & vbCrLf & "信息输入单据中！", vbInformation, gstrSysName
                                mshBill.SetFocus
                                .Row = intLop
                                .MsfObj.TopRow = intLop
                                If .TextMatrix(intLop, mCol产地) = "" Then
                                    .Col = mCol产地
                                Else
                                    .Col = mCol批号
                                End If
                                Exit Function
                            End If
                        End If
                    Else '性质是“发料部门”
                        If Split(.TextMatrix(intLop, mCol原销期), "||")(3) <> "0" Then
                            '存储格式值:最大效期||指导差价率||是否变价||在用分批||库房分批
                            If Split(.TextMatrix(intLop, mCol原销期), "||")(0) <> "0" Then
                                            
                                If Trim(.TextMatrix(intLop, mCol批号)) = "" Or Trim(.TextMatrix(intLop, mCol效期)) = "" Then
                                    ShowMsgBox "第" & intLop & "行的卫生材料是效期材料,请把它的批号及效期" & vbCrLf & "信息完整输入单据中！"
                                    mshBill.SetFocus
                                    .Row = intLop
                                    .MsfObj.TopRow = intLop
                                    If .TextMatrix(intLop, mCol批号) = "" Then
                                        .Col = mCol批号
                                    Else
                                        .Col = mCol效期
                                    End If
                                    Exit Function
                                End If
                            End If
                        End If
                        
                        '分批药品必须录入产地和批号
                        If mbln分批卫材批号产地控制 = True And Not (mint编辑状态 = 8 Or mbln退货 = True) Then '退货不检查
                        '分批药品必须录入产地和批号
                            If Split(.TextMatrix(intLop, mCol原销期), "||")(3) <> "0" And (.TextMatrix(intLop, mCol产地) = "" Or .TextMatrix(intLop, mCol批号) = "") Then
                                MsgBox "第" & intLop & "行的卫材是分批卫材,请把它的产地和批号" & vbCrLf & "信息输入单据中！", vbInformation, gstrSysName
                                mshBill.SetFocus
                                .Row = intLop
                                .MsfObj.TopRow = intLop
                                If .TextMatrix(intLop, mCol产地) = "" Then
                                    .Col = mCol产地
                                Else
                                    .Col = mCol批号
                                End If
                                Exit Function
                            End If
                        End If
                    End If
                    
                    If Val(.TextMatrix(intLop, mCol结算价)) > 9999999999# Then
                        ShowMsgBox "  第" & intLop & "行卫生材料的结算价大于了数据库能够保存的" & vbCrLf & "最大范围9999999999，请检查！"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mCol结算价
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mCol结算价)) < 0 Then
                        ShowMsgBox "  第" & intLop & "行卫生材料的结算价必需大于等于0，请检查！"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mCol结算价
                        Exit Function
                    End If
                                        
                    If Val(.TextMatrix(intLop, mCol数量)) > 9999999999# Then
                        ShowMsgBox "第" & intLop & "行卫生材料的数量大于了数据库能够保存的" & vbCrLf & "最大范围9999999999，请检查！"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mCol数量
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mCol结算金额)) > 9999999999999# Then
                        ShowMsgBox "第" & intLop & "行卫生材料的结算金额大于了数据库能够保存的" & vbCrLf & "最大范围9999999999999，请检查！"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mCol结算金额
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mCol售价金额)) > 9999999999999# Then
                        ShowMsgBox "第" & intLop & "行卫生材料的售价金额大于了数据库能够保存的" & vbCrLf & "最大范围9999999999999，请检查！"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mCol数量
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mcol零售金额)) > 9999999999999# Then
                        ShowMsgBox "第" & intLop & "行卫生材料的零售金额大于了数据库能够保存的" & vbCrLf & "最大范围9999999999999，请检查！"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mcol零售金额
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mcol零售价)) > 9999999999999# Then
                        ShowMsgBox "第" & intLop & "行卫生材料的零售价大于了数据库能够保存的" & vbCrLf & "最大范围9999999999999，请检查！"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mcol零售价
                        Exit Function
                    End If
                    
                    
                    If Val(.TextMatrix(intLop, mcol零售差价)) > 9999999999999# Then
                        ShowMsgBox "第" & intLop & "行卫生材料的零售差价大于了数据库能够保存的" & vbCrLf & "最大范围9999999999999，请检查！"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mcol零售差价
                        Exit Function
                    End If
                    
                    If LenB(StrConv(.TextMatrix(intLop, mCol随货单号), vbFromUnicode)) > 200 Then
                        MsgBox "第" & intLop & "行的物资是随货单号不能大于200个字符或100个汉字！", vbInformation, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mCol随货单号
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mCol发票金额)) > 1E+15 Then
                        ShowMsgBox "第" & intLop & "行卫生材料的售价金额大于了数据库能够保存的" & vbCrLf & "最大范围999999999999999，请检查！"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mCol发票金额
                        Exit Function
                    End If
                    If zlCommFun.ActualLen(.TextMatrix(intLop, mCol发票号)) > mint发票号Len Then
                        ShowMsgBox "第" & intLop & "行卫生材料的发票号最多能输入" & mint发票号Len & "个字符和" & mint发票号Len / 2 & "个汉字，请检查！"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mCol发票号
                        Exit Function
                    End If
                    If zlCommFun.ActualLen(.TextMatrix(intLop, mcol发票代码)) > 20 Then
                        ShowMsgBox "第" & intLop & "行卫生材料的发票代码最多能输入" & 20 & "个字符和" & 20 / 2 & "个汉字，请检查！"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mcol发票代码
                        Exit Function
                    End If
                    
                    bln高值卫材 = IsCostly(.TextMatrix(intLop, 0))
                    '是否强制录入高值卫材信息
                    If bln高值卫材录入 = True And bln高值卫材 = True Then
                        If Trim(.TextMatrix(intLop, mcol注册证号)) = "" Then
                            ShowMsgBox "第" & intLop & "行未录入“注册证号”信息，请检查！"
                            mshBill.SetFocus
                            .Row = intLop
                            .MsfObj.TopRow = intLop
                            .Col = mcol注册证号
                            Exit Function
                        End If
                        If mrsCostlyInfo.RecordCount > 0 Then mrsCostlyInfo.MoveFirst
                        mrsCostlyInfo.Find "SN=" & .TextMatrix(.Row, 1)
                        If mrsCostlyInfo.EOF Then
                            ShowMsgBox "第" & intLop & "行未录入高值卫材信息，请检查！"
                            mshBill.SetFocus
                            .Row = intLop
                            .MsfObj.TopRow = intLop
                            .Col = mCol诊疗
                            Exit Function
                        Else
                            Dim blnCostlyOK As Boolean
                            blnCostlyOK = True
                            If IIf(IsNull(mrsCostlyInfo!科室), "", mrsCostlyInfo!科室) = "" Then
                                blnCostlyOK = False
                                ShowMsgBox "第" & intLop & "行未录入高值卫材的“科室”信息，请检查！"
                            End If
                            If blnCostlyOK And IIf(IsNull(mrsCostlyInfo!病人姓名), "", mrsCostlyInfo!病人姓名) = "" Then
                                blnCostlyOK = False
                                ShowMsgBox "第" & intLop & "行未录入高值卫材的“病人姓名”信息，请检查！"
                            End If
                            If blnCostlyOK And IIf(IsNull(mrsCostlyInfo!住院号), "", mrsCostlyInfo!住院号) = "" Then
                                blnCostlyOK = False
                                ShowMsgBox "第" & intLop & "行未录入高值卫材的“住院号”信息，请检查！"
                            End If
                            If blnCostlyOK And IIf(IsNull(mrsCostlyInfo!床号), "", mrsCostlyInfo!床号) = "" Then
                                blnCostlyOK = False
                                ShowMsgBox "第" & intLop & "行未录入高值卫材的“床号”信息，请检查！"
                            End If
                            If blnCostlyOK = False Then
                                mshBill.SetFocus
                                .Row = intLop
                                .MsfObj.TopRow = intLop
                                .Col = mCol诊疗
                                Exit Function
                            End If
                        End If
                    End If
                    
                    '中标单价判断
                    Dim dblCostPrice As Double, dblPrice As Double
                    Dim strBidMess As String
                    dblCostPrice = Get中标单位成本价(.TextMatrix(intLop, 0))
                    dblPrice = CDbl(IIf(.TextMatrix(intLop, mCol采购价) <> "", .TextMatrix(intLop, mCol采购价), _
                                    IIf(.TextMatrix(intLop, mCol采购价) = "", 0, .TextMatrix(intLop, mCol采购价))))
                    If dblCostPrice < dblPrice And dblCostPrice <> 0 Then
                        strBidMess = zlDatabase.GetPara("入库单价超中标单价", glngSys, mlngModule)
                        If Val(strBidMess) = 0 Then     '禁止入库单价超中标单价
                            ShowMsgBox "第" & intLop & "行禁止采购价（" & dblPrice & "）超 中标单价（" & dblCostPrice & "）。"
                            mshBill.SetFocus
                            .Row = intLop
                            .MsfObj.TopRow = intLop
                            .Col = mCol采购价
                            Exit Function
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
ErrHandle:
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
    If txtProvider.Text = "" Then
        txtProvider.Tag = "0"
        Exit Sub
    End If
    
    If CheckQualifications(mlngModule, 2, Val(txtProvider.Tag)) = False Then
        txtProvider.Text = ""
        txtProvider.Tag = "0"
        Exit Sub
    End If
    
    If Val(txtProvider.Tag) <> mlng供货单位ID And (mint编辑状态 = 8 Or mbln退货) Then
        mlng供货单位ID = Val(txtProvider.Tag)
        mshBill.ClearBill
        mshBill.TextMatrix(1, mCol行号) = "1"
    End If
End Sub

Private Function SaveCard(Optional ByVal bln强制保存 As Boolean = False) As Boolean
'----------------------------------------------------------------------------
'修改该过程时，注意 frmPurchaseVerifyBatch(批量审核)窗体的过程是否涉及
'----------------------------------------------------------------------------
    Dim chrNo As Variant
    Dim lng序号 As Long
    Dim lngStockID As Long
    Dim lng供货单位id As Long
    Dim lng材料ID As Long
    Dim str批号 As String
    Dim str产地 As String
    Dim str效期 As String
    Dim dbl实际数量 As Double
    Dim dbl成本价 As Double
    Dim dbl成本金额 As Double
    Dim dbl扣率 As Double
    Dim dbl零售价 As Double
    Dim dbl零售金额 As Double
    Dim dbl差价 As Double
    Dim str零售差价 As String '收发记录表中用法字段保存的是外购入库差价差，由于用法字段类型是字符型因此如果采用double类型会出现 -.00x现象
    Dim str摘要 As String
    Dim str填制人 As String
    Dim str填制日期 As String
    Dim str审核人 As String
    Dim datAssessDate As String
    Dim str发票号 As String
    Dim str发票代码 As String
    Dim str发票日期 As String
    Dim str灭菌日期 As String
    Dim str灭菌失效期 As String
    Dim dbl发票金额 As Double
    Dim str生产日期  As String
    Dim str核查人 As String
    Dim str核查日期 As String
    Dim str注册证号 As String
    Dim intUnit As Integer
    Dim strUnit As String
    Dim str指导批发价 As String
    Dim str随货单号 As String
    Dim str验收结论 As String
    Dim str商品条码 As String
    Dim str内部条码 As String
    Dim str批准文号 As String
    Dim lng费用ID As Long
    Dim intRow As Integer
    Dim i As Integer
    Dim arrSQL As Variant
    Dim blnTrans As Boolean
    Dim n As Long
    
    SaveCard = False
    arrSQL = Array()
    With mshBill
        
        chrNo = Trim(txtNO)
        lngStockID = cboStock.ItemData(cboStock.ListIndex)
        
        If mint编辑状态 = 1 Then
            If chrNo <> "" Then
                If CheckNOExists(68, chrNo) Then Exit Function
            End If
            If chrNo = "" Then
                chrNo = sys.GetNextNo(68, lngStockID)
            End If
            If IsNull(chrNo) Then Exit Function
        End If
        txtNO.Tag = chrNo
        lng供货单位id = txtProvider.Tag
        str摘要 = Trim(txt摘要.Text)
        
        
        '刘兴宏:2007/05/15:加入核查人
        str填制人 = Txt填制人
        str核查人 = IIf(txt核查人.Visible, txt核查人, "")
        str审核人 = Txt审核人
        
        If mint编辑状态 = 9 Then
            str填制日期 = Trim(Txt填制日期.Caption)
            str核查日期 = IIf(txt核查人.Visible, Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss"), "")
        Else
            str填制日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            If bln强制保存 Then
                str核查日期 = Trim(txt核查日期.Caption)
            Else
                str核查日期 = ""
            End If
        End If
        
        On Error GoTo ErrHandle
        
        If mint编辑状态 = 2 Or mint编辑状态 = 9 Or bln强制保存 Then         '2:修改;  9:核查;
            gstrSQL = "zl_材料外购_Delete('" & mstr单据号 & "')"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = gstrSQL
        End If
        
        '取该库房的单位，更新指导批发价时使用
        '按药品ID顺序更新数据
        recSort.Sort = "药品id,批次,序号"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!行号
'        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                lng材料ID = .TextMatrix(intRow, 0)
                str产地 = .TextMatrix(intRow, mCol产地)
                str批号 = .TextMatrix(intRow, mCol批号)
                
                str批准文号 = .TextMatrix(intRow, mcol批准文号)
                str效期 = IIf(Trim(.TextMatrix(intRow, mCol效期)) = "", "", .TextMatrix(intRow, mCol效期))
                dbl实际数量 = GetFormat(.TextMatrix(intRow, mCol数量) * .TextMatrix(intRow, mCol比例系数), g_小数位数.obj_最大小数.数量小数)
                dbl扣率 = Val(.TextMatrix(intRow, mCol扣率))
                dbl成本价 = GetFormat(Val(.TextMatrix(intRow, mCol结算价)) / .TextMatrix(intRow, mCol比例系数), g_小数位数.obj_最大小数.成本价小数)
                dbl成本金额 = GetFormat(Val(.TextMatrix(intRow, mCol结算金额)), g_小数位数.obj_最大小数.金额小数)
                
                '刘兴宏:零售价处理
                
                'dbl零售价 = Round(Val(.TextMatrix(intRow, mCol售价)) / .TextMatrix(intRow, mCol比例系数), g_小数位数.obj_散装小数.零售价小数)
                'dbl零售金额 = Round(Val(.TextMatrix(intRow, mCol售价金额)), g_小数位数.obj_散装小数.金额小数)
                '数据库中的:差价 = 零售金额 - 结算金额
                '数据库中的:用法 = 零售金额-售价金额或零售差价-差价(库房单位的差价)

                dbl零售价 = GetFormat(Val(.TextMatrix(intRow, mcol零售价)), g_小数位数.obj_最大小数.零售价小数)
                dbl零售金额 = GetFormat(Val(.TextMatrix(intRow, mcol零售金额)), g_小数位数.obj_最大小数.零售价小数)
                dbl差价 = GetFormat(Val(.TextMatrix(intRow, mcol零售差价)), g_小数位数.obj_最大小数.零售价小数)
                str零售差价 = GetFormat(Val(.TextMatrix(intRow, mcol零售差价)) - Val(.TextMatrix(intRow, mCol差价)), g_小数位数.obj_最大小数.零售价小数)
                'dbl差价 = Round(Val(.TextMatrix(intRow, mCol差价)), g_小数位数.obj_散装小数.金额小数)
                lng序号 = intRow
                
                str验收结论 = Trim(.TextMatrix(intRow, mCol验收结论))
                str随货单号 = Trim(.TextMatrix(intRow, mCol随货单号))
                str发票号 = Trim(.TextMatrix(intRow, mCol发票号))
                str发票代码 = Trim(.TextMatrix(intRow, mcol发票代码))
                str发票日期 = Trim(IIf(.TextMatrix(intRow, mCol发票日期) = "", "", .TextMatrix(intRow, mCol发票日期)))
                dbl发票金额 = Round(Val(.TextMatrix(intRow, mCol发票金额)), g_小数位数.obj_散装小数.金额小数)
                  
                str灭菌日期 = Trim(IIf(.TextMatrix(intRow, mcol灭菌日期) = "", "", .TextMatrix(intRow, mcol灭菌日期)))
                str灭菌失效期 = Trim(IIf(.TextMatrix(intRow, mcol灭菌失效期) = "", "", .TextMatrix(intRow, mcol灭菌失效期)))
                str生产日期 = Trim(IIf(.TextMatrix(intRow, mcol生产日期) = "", "", .TextMatrix(intRow, mcol生产日期)))
                str注册证号 = Trim(.TextMatrix(intRow, mcol注册证号))
                
                str内部条码 = Trim(.TextMatrix(intRow, mcol内部条码))
                lng费用ID = Val(.TextMatrix(intRow, mcol费用ID))
                
                If gblnCode = True Then str商品条码 = Trim(.TextMatrix(intRow, mcol商品条码))
                
                '更新材料特性中的指导批发价
                str指导批发价 = Val(.TextMatrix(intRow, mCol指导批发价)) & "/" & IIf(mintUnit = 0, "1", "换算系数")
                
                '参数:材料ID_IN,SQL_IN
                If mbln不强制控制指导价格 = False Then
                    gstrSQL = "zl_材料特性_UpdateCustom(" & lng材料ID & ",'指导批发价=" & str指导批发价 & "')"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = gstrSQL
                End If
                ' Zl_材料外购_Insert
                gstrSQL = "zl_材料外购_INSERT("
                '  No_In         In 药品收发记录.NO%Type,
                gstrSQL = gstrSQL & "'" & chrNo & "',"
                '  序号_In       In 药品收发记录.序号%Type,
                gstrSQL = gstrSQL & "" & lng序号 & ","
                '  库房id_In     In 药品收发记录.库房id%Type,
                gstrSQL = gstrSQL & "" & lngStockID & ","
                '  供药单位id_In In 药品收发记录.供药单位id%Type,
                gstrSQL = gstrSQL & "" & lng供货单位id & ","
                '  材料id_In     In 药品收发记录.药品id%Type,
                gstrSQL = gstrSQL & "" & lng材料ID & ","
                '  产地_In       In 药品收发记录.产地%Type := Null,
                gstrSQL = gstrSQL & "'" & str产地 & "',"
                '  批号_In       In 药品收发记录.批号%Type := Null,
                gstrSQL = gstrSQL & "'" & str批号 & "',"
                '  生产日期_In   In 药品收发记录.生产日期%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str生产日期 = "", "Null", "to_date('" & Format(str生产日期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '  效期_In       In 药品收发记录.效期%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str效期 = "", "Null", "to_date('" & Format(str效期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '  灭菌日期_In   In 药品收发记录.灭菌日期%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str灭菌日期 = "", "Null", "to_date('" & Format(str灭菌日期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '  灭菌效期_In   In 药品收发记录.灭菌效期%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str灭菌失效期 = "", "Null", "to_date('" & Format(str灭菌失效期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '  实际数量_In   In 药品收发记录.实际数量%Type := Null,
                gstrSQL = gstrSQL & "" & dbl实际数量 & ","
                '  成本价_In     In 药品收发记录.成本价%Type := Null,
                gstrSQL = gstrSQL & "" & dbl成本价 & ","
                '  成本金额_In   In 药品收发记录.成本金额%Type := Null,
                gstrSQL = gstrSQL & "" & dbl成本金额 & ","
                '  扣率_In       In 药品收发记录.扣率%Type := Null,
                gstrSQL = gstrSQL & "" & dbl扣率 & ","
                '  零售价_In     In 药品收发记录.零售价%Type := Null,
                gstrSQL = gstrSQL & "" & dbl零售价 & ","
                '  零售金额_In   In 药品收发记录.零售金额%Type := Null,
                gstrSQL = gstrSQL & "" & dbl零售金额 & ","
                '  差价_In       In 药品收发记录.差价%Type := Null,
                gstrSQL = gstrSQL & "" & dbl差价 & ","
                '  零售差价_In   In 药品收发记录.差价%Type := Null,目前存放在用法字段
                gstrSQL = gstrSQL & "" & str零售差价 & ","
                '  摘要_In       In 药品收发记录.摘要%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str摘要 = "", "NULL", "'" & str摘要 & "'") & ","
                '   注册证号_In   In 药品收发记录.注册证号%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str注册证号 = "", "NULL", "'" & str注册证号 & "'") & ","
                '  填制人_In     In 药品收发记录.填制人%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str填制人 = "", "NULL", "'" & str填制人 & "'") & ","
                '  随货单号_In   In 应付记录.随货单号%Type := Null
                gstrSQL = gstrSQL & "" & IIf(str随货单号 = "", "NULL", "'" & str随货单号 & "'") & ","
                '  发票号_In     In 应付记录.发票号%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str发票号 = "", "NULL", "'" & str发票号 & "'") & ","
                '  发票日期_In   In 应付记录.发票日期%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str发票日期 = "", "Null", "to_date('" & Format(str发票日期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '  发票金额_In   In 应付记录.发票金额%Type := Null,
                gstrSQL = gstrSQL & "" & dbl发票金额 & ","
                '  填制日期_In   In 药品收发记录.填制日期%Type := Null,
                gstrSQL = gstrSQL & "to_date('" & str填制日期 & "','yyyy-mm-dd HH24:MI:SS'),"
                '  核查人_In     In 药品收发记录.配药人%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str核查人 = "", "NULL", "'" & str核查人 & "'") & ","
                '  核查日期_In   In 药品收发记录.配药日期%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str核查日期 = "", "Null", "to_date('" & str核查日期 & "','yyyy-mm-dd hh24:mi:ss')") & ","
                '  批次_In       In 药品收发记录.批次%Type := 0,
                gstrSQL = gstrSQL & "" & Val(.TextMatrix(intRow, mCol批次)) & ","
                '  退货_In       In Number := 1
                gstrSQL = gstrSQL & "" & IIf(mbln退货, -1, 1) & ","
                '  高值材料_In   In varchar2(250)
                gstrSQL = gstrSQL & "'" & GetCostlyInfoStr(intRow) & "'" & ","
                '  商品条码_In   In 药品收发记录.商品条码%Type :=Null
                gstrSQL = gstrSQL & "" & IIf(str商品条码 = "", "NULL", "'" & str商品条码 & "'") & ","
                '  内部条码
                gstrSQL = gstrSQL & IIf(str内部条码 = "", "Null", "'" & str内部条码 & "'") & ","
                '  费用ID
                gstrSQL = gstrSQL & IIf(lng费用ID = 0, "Null", lng费用ID) & ","
                '  发票代码
                gstrSQL = gstrSQL & IIf(str发票代码 = "", "NULL", "'" & str发票代码 & "'") & ","
                '   财务审核
                gstrSQL = gstrSQL & "0" & ","
                '   批准文号
                gstrSQL = gstrSQL & IIf(str批准文号 = "", "NULL", "'" & str批准文号 & "'") & ","
                '   验收结论
                gstrSQL = gstrSQL & "'" & str验收结论 & "'"
                gstrSQL = gstrSQL & ")"
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = gstrSQL
            End If
            
            recSort.MoveNext
        Next
        
        If bln强制保存 = False Then gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "SaveCard")
        Next
         If bln强制保存 = False Then gcnOracle.CommitTrans: blnTrans = False
        
        mstr单据号 = chrNo
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveCard = True
    Exit Function
    
ErrHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'退货
Private Function SaveRestore() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:退货保存
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-11-28 14:15:40
    '-----------------------------------------------------------------------------------------------------------

    Dim lng序号 As Long, lngStockID As Long, lng供货单位id As Long, lng材料ID As Long
    Dim str批号 As String, str产地 As String, str效期 As String, chrNo As String
    Dim dbl实际数量 As Double, dbl成本价 As Double, dbl成本金额 As Double, dbl扣率 As Double
    Dim dbl零售价 As Double, dbl零售金额 As Double, dbl差价 As Double, dbl零售差价 As Double
    Dim str摘要 As String, str填制人 As String, str填制日期 As String, str审核人 As String
    Dim datAssessDate As String, str生产日期 As String, str注册证号 As String
    Dim str发票号 As String, str发票日期 As String, dbl发票金额 As Double
    Dim intUnit As Integer, strUnit As String, str指导批发价 As String
    Dim intRow As Integer, str灭菌日期 As String, str灭菌失效期 As String, str核查人 As String
    Dim str核查日期 As String, str随货单号 As String
    Dim str商品条码 As String
    Dim i As Integer
    Dim arrSQL As Variant
    Dim n As Long
    
    SaveRestore = False
    arrSQL = Array()
    '只有库房才允许使用退货功能
    
    If Val(txtProvider.Tag) = 0 Then
        MsgBox "请选择供应商！", vbInformation, gstrSysName
        txtProvider.SetFocus
        Exit Function
    End If
    
    With mshBill
        chrNo = Trim(txtNO.Tag)
        lngStockID = cboStock.ItemData(cboStock.ListIndex)
        If chrNo = "" Then chrNo = sys.GetNextNo(68, lngStockID)
        If IsNull(chrNo) Then Exit Function
        
        txtNO.Tag = chrNo
        lng供货单位id = Val(txtProvider.Tag)
        str摘要 = Trim(txt摘要.Text)
        str填制人 = Txt填制人
        str填制日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        
        str审核人 = Txt审核人
        str核查人 = txt核查人
        str核查日期 = txt核查日期
        
        On Error GoTo ErrHandle
        
        '按药品ID顺序更新数据
        recSort.Sort = "药品id,批次,序号"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!行号
'        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                lng材料ID = Val(.TextMatrix(intRow, 0))
                str产地 = .TextMatrix(intRow, mCol产地)
                str批号 = .TextMatrix(intRow, mCol批号)
                str效期 = IIf(.TextMatrix(intRow, mCol效期) = "", "", .TextMatrix(intRow, mCol效期))
                dbl实际数量 = Round(Val(.TextMatrix(intRow, mCol数量)) * .TextMatrix(intRow, mCol比例系数), g_小数位数.obj_最大小数.数量小数)
                dbl扣率 = Val(.TextMatrix(intRow, mCol扣率))
                dbl成本价 = Round(Val(.TextMatrix(intRow, mCol结算价)) / .TextMatrix(intRow, mCol比例系数), g_小数位数.obj_最大小数.成本价小数)
                dbl成本金额 = Round(Val(.TextMatrix(intRow, mCol结算金额)), g_小数位数.obj_最大小数.金额小数)
                
                '刘兴宏:零售价不处理
                dbl零售价 = Round(Val(.TextMatrix(intRow, mCol售价)) / .TextMatrix(intRow, mCol比例系数), g_小数位数.obj_最大小数.零售价小数)
                dbl零售金额 = Round(Val(.TextMatrix(intRow, mCol售价金额)), g_小数位数.obj_最大小数.金额小数)
                dbl差价 = Round(Val(.TextMatrix(intRow, mCol差价)), g_小数位数.obj_最大小数.金额小数)
                
                'dbl零售价 = Round(Val(.TextMatrix(intRow, mCol售价)) / .TextMatrix(intRow, mCol比例系数), g_小数位数.obj_散装小数.零售价小数)
'                'dbl零售金额 = Round(Val(.TextMatrix(intRow, mCol售价金额)), g_小数位数.obj_散装小数.金额小数)
'                dbl零售价 = Round(Val(.TextMatrix(intRow, mcol零售价)), g_小数位数.obj_散装小数.零售价小数)
'                dbl零售金额 = Round(Val(.TextMatrix(intRow, mcol零售金额)), g_小数位数.obj_散装小数.零售价小数)
'                dbl零售差价 = Round(Val(.TextMatrix(intRow, mcol零售差价)), g_小数位数.obj_散装小数.零售价小数)
  
                dbl差价 = Round(Val(.TextMatrix(intRow, mCol差价)), g_小数位数.obj_最大小数.金额小数)
                lng序号 = intRow
                
                str随货单号 = Trim(.TextMatrix(intRow, mCol随货单号))
                str发票号 = Trim(.TextMatrix(intRow, mCol发票号))
                str注册证号 = Trim(.TextMatrix(intRow, mcol注册证号))
                str商品条码 = Trim(.TextMatrix(intRow, mcol商品条码))
                str发票日期 = IIf(.TextMatrix(intRow, mCol发票日期) = "", "", .TextMatrix(intRow, mCol发票日期))
                dbl发票金额 = Round(Val(.TextMatrix(intRow, mCol发票金额)), g_小数位数.obj_最大小数.金额小数)
                
                str生产日期 = IIf(.TextMatrix(intRow, mcol生产日期) = "", "", .TextMatrix(intRow, mcol生产日期))
                str灭菌日期 = IIf(.TextMatrix(intRow, mcol灭菌日期) = "", "", .TextMatrix(intRow, mcol灭菌日期))
                str灭菌失效期 = IIf(.TextMatrix(intRow, mcol灭菌失效期) = "", "", .TextMatrix(intRow, mcol灭菌失效期))
                ' Zl_材料外购_Insert
                gstrSQL = "Zl_材料外购_Insert("
                '    No_In         In 药品收发记录.NO%Type,
                gstrSQL = gstrSQL & "'" & chrNo & "',"
                '    序号_In       In 药品收发记录.序号%Type,
                gstrSQL = gstrSQL & "" & lng序号 & ","
                '    库房id_In     In 药品收发记录.库房id%Type,
                gstrSQL = gstrSQL & "" & lngStockID & ","
                '    供药单位id_In In 药品收发记录.供药单位id%Type,
                gstrSQL = gstrSQL & "" & lng供货单位id & ","
                '    材料id_In     In 药品收发记录.药品id%Type,
                gstrSQL = gstrSQL & "" & lng材料ID & ","
                '    产地_In       In 药品收发记录.产地%Type := Null,
                gstrSQL = gstrSQL & "'" & str产地 & "',"
                '    批号_In       In 药品收发记录.批号%Type := Null,
                gstrSQL = gstrSQL & "'" & str批号 & "',"
                '    生产日期_In   In 药品收发记录.生产日期%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str生产日期 = "", "Null", "to_date('" & Format(str生产日期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '    效期_In       In 药品收发记录.效期%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str效期 = "", "Null", "to_date('" & Format(str效期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '    灭菌日期_In   In 药品收发记录.灭菌日期%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str灭菌日期 = "", "Null", "to_date('" & Format(str灭菌日期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '    灭菌效期_In   In 药品收发记录.灭菌效期%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str灭菌失效期 = "", "Null", "to_date('" & Format(str灭菌失效期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '    实际数量_In   In 药品收发记录.实际数量%Type := Null,
                gstrSQL = gstrSQL & "" & dbl实际数量 & ","
                '    成本价_In     In 药品收发记录.成本价%Type := Null,
                gstrSQL = gstrSQL & "" & dbl成本价 & ","
                '    成本金额_In   In 药品收发记录.成本金额%Type := Null,
                gstrSQL = gstrSQL & "" & dbl成本金额 & ","
                '    扣率_In       In 药品收发记录.扣率%Type := Null,
                gstrSQL = gstrSQL & "" & dbl扣率 & ","
                '    零售价_In     In 药品收发记录.零售价%Type := Null,
                gstrSQL = gstrSQL & "" & dbl零售价 & ","
                '    零售金额_In   In 药品收发记录.零售金额%Type := Null,
                gstrSQL = gstrSQL & "" & dbl零售金额 & ","
                '    差价_In       In 药品收发记录.差价%Type := Null,
                gstrSQL = gstrSQL & "" & dbl差价 & ","
                '   零售差价_In   In 药品收发记录.差价%Type := Null,目前存放在用法字段
                gstrSQL = gstrSQL & "" & dbl零售差价 & ","
                '    摘要_In       In 药品收发记录.摘要%Type := Null,
                gstrSQL = gstrSQL & "'" & str摘要 & "',"
                '    注册证号_In   In 药品收发记录.注册证号%Type := Null,
                gstrSQL = gstrSQL & "'" & str注册证号 & "',"
                '    填制人_In     In 药品收发记录.填制人%Type := Null,
                gstrSQL = gstrSQL & "'" & str填制人 & "',"
                '    随货单号_In   In 应付记录.随货单号%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str随货单号 = "", "NULL", "'" & str随货单号 & "'") & ","
                '    发票号_In     In 应付记录.发票号%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str发票号 = "", "NULL", "'" & str发票号 & "'") & ","
                '    发票日期_In   In 应付记录.发票日期%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str发票日期 = "", "Null", "to_date('" & Format(str发票日期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '    发票金额_In   In 应付记录.发票金额%Type := Null,
                gstrSQL = gstrSQL & "" & dbl发票金额 & ","
                '    填制日期_In   In 药品收发记录.填制日期%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str填制日期 = "", "Null", "to_date('" & str填制日期 & "','yyyy-mm-dd HH24:MI:SS')") & ","
                '    核查人_In     In 药品收发记录.配药人%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str核查人 = "", "NULL", "'" & str核查人 & "'") & ","
                '    核查日期_In   In 药品收发记录.配药日期%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str核查日期 = "", "Null", "to_date('" & str核查日期 & "','yyyy-mm-dd HH24:MI:SS')") & ","
                '    批次_In       In 药品收发记录.批次%Type := 0,
                gstrSQL = gstrSQL & "" & Val(.TextMatrix(intRow, mCol批次)) & ","
                '    退货_In       In Number := 1
                gstrSQL = gstrSQL & "-1,"
                '  高值材料_In   In varchar2(250)
                gstrSQL = gstrSQL & "'" & GetCostlyInfoStr(intRow) & "'" & ","
                '  商品条码_In   In 药品收发记录.商品条码%Type :=Null
                gstrSQL = gstrSQL & "" & IIf(str商品条码 = "", "NULL", "'" & str商品条码 & "'")
                gstrSQL = gstrSQL & ")"
               
               ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = gstrSQL
            End If
            recSort.MoveNext
        Next
        gcnOracle.BeginTrans
        For i = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "SaveCard")
        Next
        gcnOracle.CommitTrans
        
        mstr单据号 = chrNo
        
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveRestore = True
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'保存冲销
Private Function SaveStrike() As Boolean
    Dim int行次 As Integer
    Dim int原记录状态 As Integer
    Dim strNo As String
    Dim int序号 As Integer
    Dim lng材料ID As Long
    Dim dbl冲销数量 As Double
    Dim str填制人 As String
    Dim str填制日期  As String
    Dim str发票号 As String
    Dim str发票日期 As String
    Dim dbl发票金额 As Double
    Dim str随货单号 As String
    Dim str发票代码 As String
    
    Dim intRow As Integer
    Dim rstemp As New ADODB.Recordset
    Dim bln全冲 As Boolean
    Dim lng库房id As Long, int库存检查 As Integer, lng批次 As Long
    Dim i As Integer
    Dim arrSQL As Variant
    Dim n As Long
    
    arrSQL = Array()
    SaveStrike = False
    With mshBill
        '检查冲销数量（符号必须与原始数量相同；已付款的记录不允许冲销；财务审核的单据也不允许冲销）
        strNo = Trim(txtNO.Tag)
        lng库房id = cboStock.ItemData(cboStock.ListIndex)
        int库存检查 = Get出库检查(cboStock.ItemData(cboStock.ListIndex))
        
        For intRow = 1 To .Rows - 1
            If Val(.TextMatrix(intRow, mCol冲销数量)) <> 0 Then
                If Not 相同符号(Val(.TextMatrix(intRow, mCol数量)), Val(.TextMatrix(intRow, mCol冲销数量))) Then
                    MsgBox "请输入合法的冲销数量（第" & intRow & "行）！", vbInformation, gstrSysName
                    .MsfObj.TopRow = intRow
                    Exit Function
                End If
                If .RowData(intRow) <> 0 Then
                    MsgBox "第" & intRow & "行的卫生材料已经付款，不允许冲销！", vbInformation, gstrSysName
                    .MsfObj.TopRow = intRow
                    Exit Function
                End If
                If int库存检查 <> 0 And mint编辑状态 <> 7 And mbln退货 = False Then
                    dbl冲销数量 = Round(Val(.TextMatrix(intRow, mCol冲销数量)) * Val(.TextMatrix(intRow, mCol比例系数)), g_小数位数.obj_散装小数.数量小数)
                    lng批次 = 取单据批次(15, strNo, Val(.TextMatrix(intRow, 0)), Val(.TextMatrix(intRow, mCol序号)))
                    If Check可用数量(lng库房id, Val(.TextMatrix(intRow, 0)), lng批次, dbl冲销数量, int库存检查) = False Then Exit Function
                End If
            End If
        Next
        
        str填制人 = UserInfo.用户名
        str填制日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        
        int原记录状态 = mint记录状态
        
        On Error GoTo ErrHandle
        
        int行次 = 0
        '按药品ID顺序更新数据
        recSort.Sort = "药品id,批次,序号"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!行号

'        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" And (Val(.TextMatrix(intRow, mCol冲销数量)) <> 0 Or mint编辑状态 = 7) Then
                int行次 = int行次 + 1
                int序号 = Val(.TextMatrix(intRow, mCol序号))
                
                lng材料ID = Val(.TextMatrix(intRow, 0))
                dbl冲销数量 = Round(IIf(mbln退货, -1, 1) * Val(.TextMatrix(intRow, mCol冲销数量)) * Val(.TextMatrix(intRow, mCol比例系数)), g_小数位数.obj_散装小数.数量小数)
                
                str随货单号 = Trim(.TextMatrix(intRow, mCol随货单号))
                str发票号 = Trim(.TextMatrix(intRow, mCol发票号))
                str发票代码 = Trim(.TextMatrix(intRow, mcol发票代码))
                str发票日期 = IIf(.TextMatrix(intRow, mCol发票日期) = "", "", .TextMatrix(intRow, mCol发票日期))
                If mint编辑状态 = 7 Then
                    '找因原单据的发票
                    dbl发票金额 = GetTotale发票金额(strNo, lng材料ID, int序号)
                Else
                    dbl发票金额 = Round(IIf(mbln退货, -1, 1) * Val(.TextMatrix(intRow, mCol发票金额)), g_小数位数.obj_散装小数.金额小数)
                End If
                bln全冲 = False
                If dbl冲销数量 = Round(IIf(mbln退货, -1, 1) * Val(.TextMatrix(intRow, mCol数量)) * Val(.TextMatrix(intRow, mCol比例系数)), g_小数位数.obj_散装小数.数量小数) Then
                    bln全冲 = True
                End If
                
                If mint编辑状态 = 7 Then
                    bln全冲 = True
                End If
                
                
                ' Zl_材料外购_Strike
                gstrSQL = "ZL_材料外购_STRIKE("
                '  行次_In       In Integer,
                gstrSQL = gstrSQL & "" & int行次 & ","
                '  原记录状态_In In 药品收发记录.记录状态%Type,
                gstrSQL = gstrSQL & "" & int原记录状态 & ","
                '  No_In         In 药品收发记录.NO%Type,
                gstrSQL = gstrSQL & "'" & strNo & "',"
                '  序号_In       In 药品收发记录.序号%Type,
                gstrSQL = gstrSQL & "" & int序号 & ","
                '  材料id_In     In 药品收发记录.药品id%Type,
                gstrSQL = gstrSQL & "" & lng材料ID & ","
                '  冲销数量_In   In 药品收发记录.实际数量%Type,
                gstrSQL = gstrSQL & "" & dbl冲销数量 & ","
                '  填制人_In     In 药品收发记录.填制人%Type,
                gstrSQL = gstrSQL & "'" & str填制人 & "',"
                '  填制日期_In   In 药品收发记录.填制日期%Type,
                gstrSQL = gstrSQL & "to_date('" & Format(mstr审核日期, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS'),"
                '  随货单号_In   In 应付记录.随货单号%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str随货单号 = "", "null", "'" & str随货单号 & "'") & ","
                '  发票号_In     In 应付记录.发票号%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str发票号 = "", "null", "'" & str发票号 & "'") & ","
                '  发票日期_In   In 应付记录.发票日期%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str发票日期 = "", "Null", "to_date('" & Format(str发票日期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '  发票金额_In   In 应付记录.发票金额%Type := Null,
                gstrSQL = gstrSQL & "" & dbl发票金额 & ","
                '  全部冲销_In   In 药品收发记录.实际数量%Type := 0 --用于财务审核
                gstrSQL = gstrSQL & "" & IIf(bln全冲, 1, 0) & ","
                '  财务审核_In   In Number := 0 --财务审核标志:1-财务审核,0-冲销
                gstrSQL = gstrSQL & "" & IIf(mint编辑状态 = 7, 1, 0) & ","
                '摘要_in in 药品收发记录.摘要%type
                gstrSQL = gstrSQL & "'" & txt摘要.Text & "',"
                '  发票代码_In      应付记录.发票代码%type :=null
                gstrSQL = gstrSQL & IIf(str发票代码 = "", "NULL)", "'" & str发票代码 & "')")
                Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
            End If
            recSort.MoveNext
        Next
        If mint编辑状态 <> 7 Then gcnOracle.BeginTrans
            For i = 0 To UBound(arrSQL)
                Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "SaveCard")
            Next
        If mint编辑状态 <> 7 Then gcnOracle.CommitTrans
        If int行次 = 0 Then
            ShowMsgBox "没有选择一行卫生材料来冲销，不能冲销，请检查！"
            Exit Function
        End If
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveStrike = True
    Exit Function
ErrHandle:
    If mint编辑状态 <> 7 Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveRecipe() As Boolean
    Dim chrNo As String
    Dim lng序号 As Long
    Dim str发票号 As String
    Dim str发票代码 As String
    Dim str发票日期 As String
    Dim dbl发票金额 As Double
    Dim cllTemp As New Collection
    Dim intRow As Integer
    Dim n As Long
    
    SaveRecipe = False
    '检查是否输入供药单位
    If Val(txtProvider.Tag) = 0 Then
        MsgBox "请选择卫生材料的供应商！", vbInformation, gstrSysName
        txtProvider.SetFocus
        Exit Function
    End If

    With mshBill
        chrNo = Trim(txtNO.Tag)
        
        '按药品ID顺序更新数据
        recSort.Sort = "药品id,批次,序号"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!行号
'        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
            
                lng序号 = Val(.TextMatrix(intRow, mCol序号))
                
                str发票号 = .TextMatrix(intRow, mCol发票号)
                str发票代码 = Trim(.TextMatrix(intRow, mcol发票代码))
                str发票日期 = IIf(.TextMatrix(intRow, mCol发票日期) = "", "", .TextMatrix(intRow, mCol发票日期))
                dbl发票金额 = Round(Val(.TextMatrix(intRow, mCol发票金额)), g_小数位数.obj_散装小数.金额小数)
                
                '    NO_IN       IN 药品收发记录.NO%TYPE := NULL,
                '    记录状态_IN     IN 药品收发记录.记录状态%type:=NULL,
                '    序号_IN     IN 药品收发记录.序号%TYPE:=NULL,
                '    发票号_IN       IN 应付记录.发票号%TYPE := NULL,
                '    发票日期_IN     IN 应付记录.发票日期%TYPE := NULL,
                '    发票金额_IN     IN 应付记录.发票金额%TYPE := NULL,
                '    供药单位_IN     in 应付记录.单位ID%TYPE:=0,
                '    发票代码_in     in 应付记录.发票代码%type := null
                
                gstrSQL = "zl_材料外购发票信息_UPDATE( "
                gstrSQL = gstrSQL & "'" & chrNo & "',"
                gstrSQL = gstrSQL & "" & mint记录状态 & ","
                gstrSQL = gstrSQL & "" & lng序号 & ","
                gstrSQL = gstrSQL & "'" & str发票号 & "',"
                gstrSQL = gstrSQL & "" & IIf(str发票日期 = "", "Null", "to_date('" & Format(str发票日期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                gstrSQL = gstrSQL & "" & dbl发票金额 & ","
                gstrSQL = gstrSQL & "" & Val(txtProvider.Tag) & ","
                gstrSQL = gstrSQL & IIf(str发票代码 = "", "NULL", "'" & str发票代码 & "'") & ")"
                AddArray cllTemp, gstrSQL
            End If
            recSort.MoveNext
        Next
        err = 0: On Error GoTo ErrHandle
        ExecuteProcedureArrAy cllTemp, mstrCaption
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveRecipe = True
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveRegist() As Boolean
    Dim chrNo As String
    Dim lng材料ID As Long
    Dim str注册证号 As String
    Dim cllTemp As New Collection
    Dim intRow As Integer
    Dim n As Long
    
    SaveRegist = False

    With mshBill
        chrNo = Trim(txtNO.Tag)
        
        '按药品ID顺序更新数据
        recSort.Sort = "药品id,批次,序号"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!行号
'        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
            
                lng材料ID = .TextMatrix(intRow, 0)
                str注册证号 = .TextMatrix(intRow, mcol注册证号)
                
                gstrSQL = "Zl_材料外购_修改注册证号( "
                gstrSQL = gstrSQL & "'" & chrNo & "',"
                gstrSQL = gstrSQL & "" & lng材料ID & ","
                gstrSQL = gstrSQL & "'" & str注册证号 & "')"
                AddArray cllTemp, gstrSQL
            End If
            recSort.MoveNext
        Next
        err = 0: On Error GoTo ErrHandle
        ExecuteProcedureArrAy cllTemp, mstrCaption
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveRegist = True
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub 显示合计金额()
    Dim curTotal As Double, Cur记帐金额 As Double, Cur记帐差价 As Double
    Dim intLop As Integer
    
    curTotal = 0: Cur记帐金额 = 0: Cur记帐差价 = 0
    
    With mshBill
        For intLop = 1 To .Rows - 1
            curTotal = curTotal + Val(.TextMatrix(intLop, mCol结算金额))
            Cur记帐金额 = Cur记帐金额 + Val(.TextMatrix(intLop, mCol售价金额))
            Cur记帐差价 = Cur记帐差价 + Val(.TextMatrix(intLop, mCol差价))
        Next
    End With
    
'    Cur记帐差价 = Cur记帐金额 - curTotal
    lblPurchasePrice.Caption = "结算金额合计：" & Format(curTotal, mFMT.FM_金额)
    lblSalePrice.Caption = "售价金额合计：" & Format(Cur记帐金额, mFMT.FM_金额)
    lblDifference.Caption = "差价合计：" & Format(Cur记帐差价, mFMT.FM_金额)
End Sub

Private Sub 提示库存数()
    Dim rstemp As New ADODB.Recordset
    Dim dbl数量 As Double
    Dim str单位 As String, strUnit As String, strQuantity As String
    Dim intID As Long, lng批次 As Long
    
    On Error GoTo ErrHandle
    If mshBill.TextMatrix(mshBill.Row, mCol诊疗) = "" Then
        stbThis.Panels(2).Text = ""
        Exit Sub
    End If
    If mshBill.TextMatrix(mshBill.Row, 0) = "" Then Exit Sub
    intID = mshBill.TextMatrix(mshBill.Row, 0)
    lng批次 = Val(mshBill.TextMatrix(mshBill.Row, mCol批次))
    If mintUnit = 0 Then
            strQuantity = "a.可用数量"
    Else
            strQuantity = "a.可用数量/b.换算系数"
    End If

    
    gstrSQL = "" & _
    "   Select b.材料ID," & IIf(mintUnit = 0, "c.计算单位", "b.包装单位") & " as 单位, Sum(nvl(" & strQuantity & ",0)) as 数量 " & _
    "   From 药品库存 a,材料特性 b,收费项目目录 c " & _
    "   Where a.性质=1 and a.药品id=b.材料id and b.材料id=c.id " & _
    "         and a.可用数量<>0 And a.库房ID=[1] and b.材料ID=[2]  " & IIf(mint编辑状态 = 8 Or mbln退货, " and nvl(批次,0)=[3]", "") & _
    "   Group by b.材料ID," & _
            IIf(mintUnit = 0, "c.计算单位", "b.包装单位")
    
   Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex), intID, lng批次)
   With rstemp
        If .EOF Then
            stbThis.Panels(2).Text = ""
            Exit Sub
        End If
        
        dbl数量 = IIf(IsNull(!数量), 0, !数量)
        stbThis.Panels(2).Text = "该卫生材料当前库存数为[" & Format(dbl数量, mFMT.FM_数量) & "]" & zlStr.Nvl(!单位)
        
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtTypeVar_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Select Case Me.lblType.Tag
        Case 1, 3, 4
            Call Comm_Selecter(Me.txtTypeVar.Text, Me.lblType.Tag + 2)
        Case Else
            Call Comm_Selecter("%" & Me.txtTypeVar.Text & "%", Me.lblType.Tag + 2)
        End Select
        Me.vsfCostlyInfo.SetFocus
    Else
        If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
        If InStr("1,3,4", Me.lblType.Tag) Then
            If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
        End If
    End If
End Sub

Private Sub Txt加价率_GotFocus()
    Txt加价率.SelStart = 0
    Txt加价率.SelLength = Len(Txt加价率)
End Sub

Private Sub Txt加价率_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then Call CmdYes_Click
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
    ImeLanguage True
    zlControl.TxtSelAll txt摘要
End Sub

Private Sub txt摘要_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey (vbKeyTab)
        KeyCode = 0
    End If
End Sub

Private Sub txt摘要_LostFocus()
    ImeLanguage False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

'取数据库中发票号的长度，这样，程序中的批号长度与数据库中保持一致了
Private Function Get发票号Len() As Integer
    Dim rstemp As New Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "select 发票号 from 应付记录 where rownum<1 "
    zlDatabase.OpenRecordset rstemp, gstrSQL, "取字段长度"
    Get发票号Len = rstemp.Fields(0).DefinedSize
    rstemp.Close
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function




'取数据库中批号的长度，这样，程序中的批号长度与数据库中保持一致了
Private Function GetBatchNoLen() As Integer
    Dim rstemp As New Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "select 批号 from 药品收发记录 where rownum<1 "
    zlDatabase.OpenRecordset rstemp, gstrSQL, "取字段长度"
    GetBatchNoLen = rstemp.Fields(0).DefinedSize
    rstemp.Close
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub 检查成本价()
    Dim dbl成本价 As Double, dbl零售价 As Double, dbl售价 As Double
    
    '如果成本价比零售价还高，提示用户
    With mshBill
        If Val(.TextMatrix(.Row, 0)) = 0 Then Exit Sub
        dbl成本价 = Format(Val(.TextMatrix(.Row, mCol结算价)), "#####0.00;-#####0.00;0;")
        dbl售价 = Format(Val(.TextMatrix(.Row, mCol售价)), "#####0.00;-#####0.00;0;")
        dbl零售价 = Format(Val(.TextMatrix(.Row, mcol零售价)) * Val(.TextMatrix(.Row, mCol比例系数)), "#####0.00;-#####0.00;0;")
    End With
    If dbl成本价 > dbl售价 Then
        MsgBox "提醒：该卫生材料的成本价比售价还高！", vbInformation, gstrSysName
    End If
    If dbl成本价 > dbl零售价 Then
        MsgBox "提醒：该卫生材料的成本价比零售价还高！", vbInformation, gstrSysName
    End If
End Sub

Private Function CopyCard() As String
    Dim intRow As Integer, intUpdate As Integer, str随货单号 As String
    Dim dbl原数量 As Double, dbl现数据 As Double
    Dim dbl结算价 As Double, dbl结算金额 As Double, dbl差价 As Double, dbl零售金额 As Double, dbl扣率 As Double
    Dim dbl采购价 As Double, dbl售价 As Double, dbl发票金额 As Double, dbl零售差价 As Double
    Dim str发票号 As String, str发票日期 As String, str发票代码 As String
    Dim lng序号 As Long
    
    
    Dim strNo As Variant
    On Error GoTo ErrHand
    
    strNo = sys.GetNextNo(68, cboStock.ItemData(cboStock.ListIndex))
    If IsNull(strNo) Then Exit Function
    
    intUpdate = 0
    CopyCard = ""
    
    '复制产生新单据
    ' 单据_IN,NO_IN,NewNO_IN
    
'    gstrSQL = "zl_卫生材料_billcopy(15,'" & txtNO.Text & "','" & StrNo & "','" & UserInfo.用户名 & "')"
    gstrSQL = "zl_卫生材料_billcopy(15,'" & txtNO.Tag & "','" & strNo & "')"
    zlDatabase.ExecuteProcedure gstrSQL, mstrCaption
    '采购价，扣率，结算价，结算金额，售价，发票号，发票日期，发票金额
    
    '修改结算价、结算金额及差价（要考虑到存在审核冲销的单据，这时需要修改结算价、结算金额，差价）
    With mshBill
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                dbl结算价 = Val(.TextMatrix(intRow, mCol结算价))
                dbl结算金额 = IIf(mbln退货 = True, "-" & Val(.TextMatrix(intRow, mCol结算金额)), Val(.TextMatrix(intRow, mCol结算金额)))
                dbl差价 = Val(.TextMatrix(intRow, mCol差价))
                
                '刘兴宏:零售价处理
                dbl售价 = Val(.TextMatrix(intRow, mcol零售价))
                dbl零售金额 = IIf(mbln退货 = True, (-1) * Val(.TextMatrix(intRow, mcol零售金额)), Val(.TextMatrix(intRow, mcol零售金额)))
                dbl零售差价 = IIf(mbln退货 = True, (-1) * Val(.TextMatrix(intRow, mcol零售差价)), Val(.TextMatrix(intRow, mcol零售差价)))
                dbl差价 = Round(IIf(mbln退货 = True, (-1) * Val(.TextMatrix(intRow, mcol零售差价)), Val(.TextMatrix(intRow, mcol零售差价))), g_小数位数.obj_散装小数.零售价小数)
                dbl零售差价 = Round(dbl零售差价 - dbl差价, g_小数位数.obj_散装小数.零售价小数)
                
                dbl扣率 = Val(.TextMatrix(intRow, mCol扣率))
                
                str随货单号 = Trim(.TextMatrix(intRow, mCol随货单号))
                str发票号 = Trim(.TextMatrix(intRow, mCol发票号))
                str发票代码 = Trim(.TextMatrix(intRow, mcol发票代码))
                str发票日期 = Trim(IIf(.TextMatrix(intRow, mCol发票日期) = "", "", .TextMatrix(intRow, mCol发票日期)))
                dbl发票金额 = IIf(mbln退货 = True, (-1) * Val(.TextMatrix(intRow, mCol发票金额)), Val(.TextMatrix(intRow, mCol发票金额)))
                If dbl发票金额 = 0 Then dbl发票金额 = IIf(str随货单号 <> "", dbl结算金额, 0)
                Call Get数量(txtNO.Tag, Val(.TextMatrix(intRow, mCol序号)), dbl原数量)
                lng序号 = intRow
                
                If Get数量(strNo, Val(.TextMatrix(intRow, mCol序号)), dbl现数据) Then
                    If Abs(dbl现数据) > 0 Then
                        '修正数量
                        dbl结算价 = Round(dbl结算价 / Val(.TextMatrix(intRow, mCol比例系数)), g_小数位数.obj_最大小数.成本价小数)
                        'dbl售价 = Round(dbl售价 / Val(.TextMatrix(intRow, mCol比例系数)), g_小数位数.obj_散装小数.零售价小数)
                        dbl结算金额 = Round(dbl结算金额 * dbl现数据 / dbl原数量, g_小数位数.obj_散装小数.金额小数)
                        dbl差价 = Round(dbl差价 * dbl现数据 / dbl原数量, g_小数位数.obj_散装小数.金额小数)
                        dbl零售金额 = Round(dbl零售金额 * dbl现数据 / dbl原数量, g_小数位数.obj_散装小数.金额小数)
                        dbl零售差价 = Round(dbl零售差价 * dbl现数据 / dbl原数量, g_小数位数.obj_散装小数.金额小数)
                        
                        dbl发票金额 = Round(dbl发票金额 * dbl现数据 / dbl原数量, g_小数位数.obj_散装小数.金额小数)
                        
                        
                        '更新
                        gstrSQL = "zl_Bill_更新信息(15,'" & strNo & "'," & Val(.TextMatrix(intRow, mCol序号)) & ",'成本价','" & dbl结算价 & "')"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
                        gstrSQL = "zl_Bill_更新信息(15,'" & strNo & "'," & Val(.TextMatrix(intRow, mCol序号)) & ",'成本金额','" & dbl结算金额 & "')"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
                        gstrSQL = "zl_Bill_更新信息(15,'" & strNo & "'," & Val(.TextMatrix(intRow, mCol序号)) & ",'差价','" & dbl差价 & "')"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
                        gstrSQL = "zl_Bill_更新信息(15,'" & strNo & "'," & Val(.TextMatrix(intRow, mCol序号)) & ",'零售价','" & dbl售价 & "')"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
                        gstrSQL = "zl_Bill_更新信息(15,'" & strNo & "'," & Val(.TextMatrix(intRow, mCol序号)) & ",'零售金额','" & dbl零售金额 & "')"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
                        gstrSQL = "zl_Bill_更新信息(15,'" & strNo & "'," & Val(.TextMatrix(intRow, mCol序号)) & ",'用法','" & dbl零售差价 & "')"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
                        
                        gstrSQL = "zl_Bill_更新信息(15,'" & strNo & "'," & Val(.TextMatrix(intRow, mCol序号)) & ",'扣率','" & dbl扣率 & "')"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
                        
                        
                        gstrSQL = "zl_Bill_更新应付记录('" & strNo & "'," & Val(.TextMatrix(intRow, mCol序号)) & ",'发票金额','" & dbl发票金额 & "',5)"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
                        gstrSQL = "zl_Bill_更新应付记录('" & strNo & "'," & Val(.TextMatrix(intRow, mCol序号)) & ",'随货单号'," & IIf(str随货单号 = "", "NULL", "''" & str随货单号 & "''") & ",5)"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
                        gstrSQL = "zl_Bill_更新应付记录('" & strNo & "'," & Val(.TextMatrix(intRow, mCol序号)) & ",'发票号'," & IIf(str发票号 = "", "NULL", "''" & str发票号 & "''") & ",5)"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
                        gstrSQL = "zl_Bill_更新应付记录('" & strNo & "'," & Val(.TextMatrix(intRow, mCol序号)) & ",'发票代码'," & IIf(str发票号 = "", "NULL", "''" & str发票代码 & "''") & ",5)"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
                        gstrSQL = "zl_Bill_更新应付记录('" & strNo & "'," & Val(.TextMatrix(intRow, mCol序号)) & ",'发票日期'," & IIf(str发票日期 = "", "NULL", "to_date('" & str发票日期 & "','yyyy-mm-dd')") & ",5)"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
                        
                        intUpdate = intUpdate + 1
                    End If
                End If
            End If
        Next
    End With
    gstrSQL = "zl_材料财务审核_update(15,'" & strNo & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
    
    If intUpdate = 0 Then
        MsgBox "无法完成财务审核，因为该单据已被全部冲销！", vbInformation, gstrSysName
        Exit Function
    End If
    CopyCard = strNo
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function Get数量(ByVal strNo As String, ByVal int序号 As Integer, dbl数量 As Double) As Boolean
    Dim rstemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "" & _
        "   Select Nvl(实际数量,0) 数量 " & _
        "   From 药品收发记录" & _
        "   Where 单据=15 And NO=[1]  And 序号=[2]"
    
    
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, strNo, int序号)
    
    If rstemp.EOF Then Exit Function
    dbl数量 = rstemp!数量
    Get数量 = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function 时价材料零售价(ByVal lng材料ID As Long, ByVal sin采购价 As Double, ByVal sin加成率 As Double, _
    Optional LngLastRow As Long = -1, Optional sng售价 As Double = -99999999) As Double
    '------------------------------------------------------------------------------------------------------
    '功能:根据指导价格或差价比计算出时价材料的差价让利情况
    '入参:lng材料ID-材料ID
    '     sin采购价-采购价格
    '     sin加成率-加成率(如果传入0,同时又传入dbl零售价,则将按传入的零售价进行计算)
    '     LngLastRow-单据的行号
    '     sng售价-传入的零售价
    '出参:
    '返回:零售价
    '修改人:刘兴宏
    '修改时间:2007/2/25
    '------------------------------------------------------------------------------------------------------
       '时价材料零售价计算公式:采购价*(1+加成率)
    '改为:采购价*(1+加成率)+(指导零售价-采购价*(1+加成率))*(1-差价让利比)
    '由于差价让利比的存在,以前所有按指导差价率计算的地方,均需要将差价率转换成加成率进行计算,此函数用于返回本次公式增加的部分金额：(指导零售价-采购价*(1+加成率))*(1-差价让利比)
    
    Dim sin零售价 As Double, sin指导零售价 As Double, sin差价让利比 As Double
    Dim rstemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "Select 指导零售价,Nvl(差价让利比,100) 差价让利比 From 材料特性 Where 材料ID=[1]"
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取指导零售价", lng材料ID)
    If rstemp.EOF Then Exit Function
    
    sin指导零售价 = rstemp!指导零售价
    sin差价让利比 = rstemp!差价让利比
    
    时价材料零售价 = 0
    If sin差价让利比 = 100 Then Exit Function
    
    '如果未用指导价，就不存在让利问题
    If sin指导零售价 = 0 Then Exit Function
    If LngLastRow = -1 Then LngLastRow = mshBill.Row
    
    If mint编辑状态 = 8 Or mbln退货 Then
        '如果是退货，则按出库的方式计算售价
        gstrSQL = " Select Nvl(实际数量,0) 实际数量,Nvl(实际金额,0) 实际金额 From 药品库存 " & _
                  " Where 性质=1 And 药品ID=[1] And 库房ID=[2] And Nvl(批次,0)=[3]"
        Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "计算零售价", lng材料ID, cboStock.ItemData(cboStock.ListIndex), Val(mshBill.TextMatrix(LngLastRow, mCol批次)))
        
        
        If rstemp.RecordCount = 0 Then
            ShowMsgBox "卫生材料库存数据错误（未找到指定卫生材料的库存记录）！"
            Exit Function
        End If
        '肯定有数量，没有数量的话，无法到达此处
        时价材料零售价 = rstemp!实际金额 / rstemp!实际数量 * Val(mshBill.TextMatrix(LngLastRow, mCol比例系数))
    Else
        If sng售价 <> -99999999 And sin加成率 = 0 Then
            sin零售价 = sng售价
        Else
            sin零售价 = sin采购价 * (1 + sin加成率)
        End If
        
        If sin零售价 / Val(mshBill.TextMatrix(LngLastRow, mCol比例系数)) >= sin指导零售价 Then Exit Function
        sin指导零售价 = sin指导零售价 * Val(mshBill.TextMatrix(LngLastRow, mCol比例系数))
        
        时价材料零售价 = (sin指导零售价 - sin零售价) * (1 - sin差价让利比 / 100)
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function 计算加成率(ByVal lng材料ID As Long, ByVal sin零售价 As Double, ByVal sin成本价 As Double) As Double
    Dim sin指导零售价 As Double, sin差价让利比 As Double
    Dim rstemp As New ADODB.Recordset
    '根据零售价反算成本价,由于时价卫材公式的变化,导致原来计算加成率的公式无效,需重新计算
    '原公式:(零售价/成本价-1)*100
    '现公式的理论:由于零售价是按加成率算出来后,再加上了让利外那部分金额,因此实际按加成率算出的零售价=指导零售价-(指导零售价-零售价)/差价让利比
    '再套用原公式算出实际的加成率
    计算加成率 = 0.15
    
    On Error GoTo ErrHandle
    gstrSQL = "Select a.指导零售价,Nvl(a.差价让利比,100) 差价让利比,Nvl(b.是否变价,0) 时价 From 材料特性 A, 收费项目目录 b Where a.材料ID=b.id  and b.ID=[1]"
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取指导零售价", lng材料ID)
    If rstemp.EOF Then Exit Function
    
    sin指导零售价 = rstemp!指导零售价
    sin差价让利比 = rstemp!差价让利比
    If rstemp!时价 = 0 Then Exit Function
    
'    If mbln分段加成率 Then
'            计算加成率 = Get分段加成率(sin成本价)
'    Else
        
        '指导零售价-(指导零售价-零售价)/差价让利比
        sin指导零售价 = sin指导零售价 * Val(mshBill.TextMatrix(mshBill.Row, mCol比例系数))
        If sin差价让利比 <> 100 And sin差价让利比 > 0 Then
            sin零售价 = sin指导零售价 - (sin指导零售价 - sin零售价) / sin差价让利比 * 100
        Else
            sin零售价 = sin指导零售价 - (sin指导零售价 - sin零售价)
        End If
        计算加成率 = (sin零售价 / sin成本价 - 1) * 100
   ' End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function 校正零售价(ByVal sin零售价 As Double, Optional LngLastRow As Long = -1) As Double
    '得到按当前单位系数计算出来的指导零售价，如果时价卫材计算出来的零售价大于指导零售价，以指导零售价为准
    Dim sin指导零售价 As Double
    Dim rstemp As New ADODB.Recordset
    If LngLastRow = -1 Then LngLastRow = mshBill.Row
    
    On Error GoTo ErrHandle
    gstrSQL = "Select 指导零售价,Nvl(差价让利比,100) 差价让利比 From 材料特性 Where 材料ID=[1]"
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取指导零售价", Val(mshBill.TextMatrix(LngLastRow, 0)))
    If rstemp.EOF Then Exit Function
    
    sin指导零售价 = Val(zlStr.Nvl(rstemp!指导零售价))
    sin指导零售价 = sin指导零售价 * Val(mshBill.TextMatrix(LngLastRow, mCol比例系数))
    If sin指导零售价 = 0 Then sin指导零售价 = sin零售价
    
    If Val(sin零售价) > Val(sin指导零售价) And Not mbln不强制控制指导价格 Then
'        MsgBox "售价￥" & sin零售价 & "超过" & "指导售价￥" & sin指导零售价 & "，强行改为指导售价。", vbInformation, gstrSysName
        校正零售价 = sin指导零售价
    Else
        校正零售价 = sin零售价
    End If
    '校正零售价 = IIf(sin零售价 > sin指导零售价, sin指导零售价, sin零售价)
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetColumnByUserDefine()
    Dim intColumns As Integer
    Dim strColumn_UnSelected As String
    Dim strColumn_Selected As String
    Dim strColumn_All As String
    Dim arrColumn_All, arrColumn_Selected, arrColumn_UnSelected
    Dim strCol As String
    Dim intCol As Integer, intCols As Integer
    Dim i As Integer
    
    strColumn_Selected = zlDatabase.GetPara("选择列", glngSys, mlngModule)
    strColumn_UnSelected = zlDatabase.GetPara("屏蔽列", glngSys, mlngModule)
     
    'strColumn_All = "卫材,0|规格,1|产地,1|批号,0|生产日期,1|灭菌日期,1|灭菌失效期,1|效期,0|单位,1|数量,0|指导批发价,1|采购价,1|扣率,1|" & _
                "加成率,0|结算价,0|结算金额,0|售价,0|售价金额,0|差价,0|发票号,0|发票代码,0|发票日期,0|发票金额,0"
    
    strColumn_All = ""
    '先装入缺省设置
    i = 1: mCol行号 = i:
    i = i + 1: mCol诊疗 = i: strColumn_All = strColumn_All & "卫材," & i & "|"
    i = i + 1: mCol序号 = i:
    i = i + 1: mCol商品名 = i: strColumn_All = strColumn_All & "商品名," & i & "|"
    i = i + 1: mCol规格 = i: strColumn_All = strColumn_All & "规格," & i & "|"
    i = i + 1: mCol原产地 = i:
    i = i + 1: mCol原销期 = i:
    i = i + 1: mCol比例系数 = i:
    i = i + 1: mCol简码 = i:
    i = i + 1: mCol产地 = i: strColumn_All = strColumn_All & "产地," & i & "|"
    i = i + 1: mcol批准文号 = i: strColumn_All = strColumn_All & "批准文号," & i & "|"
    i = i + 1: mCol单位 = i: strColumn_All = strColumn_All & "单位," & i & "|"
    i = i + 1: mCol批号 = i: strColumn_All = strColumn_All & "批号," & i & "|"
    i = i + 1: mcol生产日期 = i: strColumn_All = strColumn_All & "生产日期," & i & "|"
    i = i + 1: mCol效期 = i: strColumn_All = strColumn_All & "效期," & i & "|"
    i = i + 1: mcol一次性材料 = i:
    i = i + 1: mcol条码管理 = i:
    i = i + 1: mcol灭菌效期 = i:
    i = i + 1: mcol灭菌日期 = i: strColumn_All = strColumn_All & "灭菌日期," & i & "|"
    i = i + 1: mcol灭菌失效期 = i: strColumn_All = strColumn_All & "灭菌失效期," & i & "|"
    i = i + 1: mcol注册证号 = i: strColumn_All = strColumn_All & "注册证号," & i & "|"
    i = i + 1: mcol注册证有效期 = i: strColumn_All = strColumn_All & "注册证有效期," & i & "|"
    i = i + 1: mcol内部条码 = i
    i = i + 1: mcol费用ID = i
    i = i + 1: mcol商品条码 = i
    If gblnCode = True Then
        strColumn_All = strColumn_All & "商品条码," & i & "|"
    End If
    
    i = i + 1: mCol数量 = i: strColumn_All = strColumn_All & "数量," & i & "|"
    i = i + 1: mCol冲销数量 = i
    i = i + 1: mCol批次 = i:
    i = i + 1: mCol指导批发价 = i: strColumn_All = strColumn_All & "指导批发价," & i & "|"
    i = i + 1: mCol采购价 = i: strColumn_All = strColumn_All & "采购价," & i & "|"
    i = i + 1: mCol扣率 = i: strColumn_All = strColumn_All & "扣率," & i & "|"
    i = i + 1: mCol结算价 = i: strColumn_All = strColumn_All & "结算价," & i & "|"
    i = i + 1: mCol结算金额 = i: strColumn_All = strColumn_All & "结算金额," & i & "|"
    i = i + 1: mcol加成率 = i: strColumn_All = strColumn_All & "加成率," & i & "|"
    i = i + 1: mCol售价 = i: strColumn_All = strColumn_All & "售价," & i & "|"
    i = i + 1: mCol售价金额 = i: strColumn_All = strColumn_All & "售价金额," & i & "|"
    i = i + 1: mCol差价 = i: strColumn_All = strColumn_All & "差价," & i & "|"
    
    '刘兴宏:零售价处理
    i = i + 1: mcol零售价 = i: strColumn_All = strColumn_All & "零售价," & i & "|"
    i = i + 1: mcol零售单位 = i: strColumn_All = strColumn_All & "零售单位," & i & "|"
    i = i + 1: mcol零售金额 = i: strColumn_All = strColumn_All & "零售金额," & i & "|"
    i = i + 1: mcol零售差价 = i: strColumn_All = strColumn_All & "零售差价," & i & "|"
    i = i + 1: mCol随货单号 = i: strColumn_All = strColumn_All & "随货单号," & i & "|"
    i = i + 1: mCol验收结论 = i: strColumn_All = strColumn_All & "验收结论," & i & "|"
    i = i + 1: mCol发票号 = i: strColumn_All = strColumn_All & "发票号," & i & "|"
    i = i + 1: mcol发票代码 = i: strColumn_All = strColumn_All & "发票代码," & i & "|"
    i = i + 1: mCol发票日期 = i: strColumn_All = strColumn_All & "发票日期," & i & "|"
    i = i + 1: mCol发票金额 = i: strColumn_All = strColumn_All & "发票金额," & i
    
    
    If strColumn_Selected = "" Then Exit Sub
    
    '根据用户设置调整列顺序
    arrColumn_All = Split(strColumn_All, "|")
    arrColumn_Selected = Split(strColumn_Selected, "|")
    intCols = UBound(arrColumn_Selected)
    
    For intCol = 0 To intCols
        Call SetColumnValue(arrColumn_Selected(intCol), Split(arrColumn_All(intCol), ",")(1))
    Next
    
    For intCol = 0 To UBound(arrColumn_All)
       strCol = "|" & Split(arrColumn_All(intCol), ",")(0) & "|"
       If InStr("|" & strColumn_Selected & "|", strCol) = 0 Then
           '在选择列中不存在,则肯定是隐藏,并且在未选择列中没有的,只表示新增的列,需要增加
           If InStr("|" & strColumn_UnSelected & "|", strCol) = 0 Then
               strColumn_UnSelected = strColumn_UnSelected & "|" & Split(arrColumn_All(intCol), ",")(0)
           End If
       End If
    Next
     
    '将未选择的列的列宽设置为零，且列数据为5——不可选择
    If strColumn_UnSelected = "" Then Exit Sub
    If Left(strColumn_UnSelected, 1) = "|" Then strColumn_UnSelected = Mid(strColumn_UnSelected, 2)
    intCol = intCols + 1
    intColumns = 0
    arrColumn_UnSelected = Split(strColumn_UnSelected, "|")
    intCols = UBound(arrColumn_All)
    For intCol = intCol To intCols
        If UBound(arrColumn_UnSelected) >= intColumns Then
            Call SetColumnValue(arrColumn_UnSelected(intColumns), Split(arrColumn_All(intCol), ",")(1), False)
                intColumns = intColumns + 1
        Else
            Call SetColumnValue(Split(arrColumn_All(intCol), ",")(0), Split(arrColumn_All(intCol), ",")(1), False)
        End If
    Next
End Sub

Private Sub SetColumnValue(ByVal str列名 As String, ByVal intValue As Integer, Optional ByVal blnShow As Boolean = True)
    Select Case str列名
    Case "行号"
        mCol行号 = intValue
    Case "诊疗", "卫材"
        mCol诊疗 = intValue
    Case "序号"
        mCol序号 = intValue
    Case "商品名"
        mCol商品名 = intValue
    Case "规格"
        mCol规格 = intValue
    Case "原产地"
        mCol原产地 = intValue
    Case "原销期"
        mCol原销期 = intValue
    Case "比例系数"
        mCol比例系数 = intValue
    Case "简码"
        mCol简码 = intValue
    Case "产地"
        mCol产地 = intValue
    Case "单位"
        mCol单位 = intValue
    Case "批号"
        mCol批号 = intValue
    Case "生产日期"
        mcol生产日期 = intValue
    Case "效期"
        mCol效期 = intValue
    Case "批准文号"
        mcol批准文号 = intValue
    Case "数量"
        mCol数量 = intValue
    Case "冲销数量"
        mCol冲销数量 = intValue
    Case "指导批发价"
        mCol指导批发价 = intValue
    Case "扣率"
        mCol扣率 = intValue
    Case "采购价"
        mCol采购价 = intValue
    Case "结算价"
        mCol结算价 = intValue
    Case "结算金额"
        mCol结算金额 = intValue
    Case "售价"
        mCol售价 = intValue
    Case "售价金额"
        mCol售价金额 = intValue
    Case "差价"
        mCol差价 = intValue
    Case "零售价"
        mcol零售价 = intValue
    Case "零售单位"
        mcol零售单位 = intValue
    Case "零售金额"
        mcol零售金额 = intValue
    Case "零售差价"
        mcol零售差价 = intValue
    Case "随货单号"
        mCol随货单号 = intValue
    Case "验收结论"
        mCol验收结论 = intValue
    Case "发票号"
        mCol发票号 = intValue
    Case "发票代码"
        mcol发票代码 = intValue
    Case "发票日期"
        mCol发票日期 = intValue
    Case "发票金额"
        mCol发票金额 = intValue
    Case "一次性材料 "
        mcol一次性材料 = intValue
    Case "条码管理"
        mcol条码管理 = intValue
    Case "灭菌效期 "
        mcol灭菌效期 = intValue
    Case "灭菌日期"
        mcol灭菌日期 = intValue
    Case "灭菌失效期"
        mcol灭菌失效期 = intValue
    Case "注册证号"
        mcol注册证号 = intValue
    Case "注册证有效期"
        mcol注册证有效期 = intValue
    Case "商品条码"
        mcol商品条码 = intValue
    Case "加成率"
        mcol加成率 = intValue
    Case Else
        blnShow = False
    End Select
'--    Debug.Print str列名 & vbTab & intValue
    
    If Not blnShow Then
        mshBill.ColWidth(intValue) = 0
        mshBill.ColData(intValue) = 5
    Else
        mintLastCol = intValue
    End If
End Sub

Private Function CheckStock(ByVal lng材料ID As Long, ByVal lng批次 As Long, ByVal dbl数量 As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------
    '功能:检查库存的可用数量是否充足
    '---------------------------------------------------------------------------------------------------------
    Dim lng库房id As Long, intRow As Integer, intLop As Integer
    Dim blnMsg As Boolean
    Dim dblSum As Double, dbltotal As Double
    Dim varStuff As Variant
    
    Dim rsCheck As New ADODB.Recordset
    '退货时使用本函数，用以检查输入的退货数量是否足够
    If mint编辑状态 <> 8 And mbln退货 = False Then CheckStock = True: Exit Function
    
    On Error GoTo ErrHandle
    
    '计算原单据中的原始数量
    With mshBill
        dblSum = 0
        intRow = .Row
        If mint编辑状态 <> 8 Then
            For Each varStuff In mCllBillData
                If varStuff(0) = Val(.TextMatrix(.Row, 0)) & "_" & Val(.TextMatrix(.Row, mCol批次)) Then
                    dblSum = varStuff(1)
                    Exit For
                End If
            Next
        End If
        dbltotal = 0
        For intLop = 1 To .Rows - 1
            If .TextMatrix(intLop, 0) <> "" Then
                If intLop <> intRow And Trim(.TextMatrix(intLop, 0)) = Trim(.TextMatrix(intRow, 0)) And Val(.TextMatrix(intRow, mCol批次)) = Val(.TextMatrix(intLop, mCol批次)) Then
                    dbltotal = dbltotal + Val(.TextMatrix(intLop, mCol数量)) * Val(.TextMatrix(intLop, mCol比例系数))
                End If
            End If
        Next
    End With
                
    
    
    lng库房id = cboStock.ItemData(cboStock.ListIndex)
    
    gstrSQL = "Select 可用数量 From 药品库存 Where 库房ID=[1] And Nvl(批次,0)=[2]  And 性质=1 And 药品ID=[3] "
    
    Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "检查库存是否足够——退货", lng库房id, lng批次, lng材料ID)
    If rsCheck.RecordCount <> 0 Then
        dblSum = dblSum + Val(zlStr.Nvl(rsCheck!可用数量))
    Else
    End If
    dbltotal = dbltotal + dbl数量
    If dbltotal > dblSum Then
        ShowMsgBox "退货数量不能大于现有的库存数量（当前库存数量为：" & dblSum / Val(mshBill.TextMatrix(mshBill.Row, mCol比例系数)) & "）！"
        Exit Function
    End If
    CheckStock = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function 库存检查_退货() As Boolean
    Dim lngRow As Long, lngRows As Long, lng材料ID As Long, lng库房id As Long, lng批次 As Long
    Dim dbl库存数量 As Double, dbl退货数量 As Double
    
    Dim rstemp As New ADODB.Recordset
    Dim blnExit As Boolean
    
    On Error GoTo ErrHandle

    '用于审核退货单据时
    lngRows = mshBill.Rows - 1
    For lngRow = 1 To lngRows
        blnExit = False
        lng材料ID = Val(mshBill.TextMatrix(lngRow, 0))
        
        If lng材料ID <> 0 And Val(mshBill.TextMatrix(lngRow, mCol数量)) < 0 Then
        
            lng批次 = Val(mshBill.TextMatrix(lngRow, mCol批次))
            
            lng库房id = cboStock.ItemData(cboStock.ListIndex)
            
            gstrSQL = "" & _
                "   Select Nvl(A.实际数量,0)/" & Choose(mintUnit + 1, "1", "B.换算系数") & " As 数量 " & _
                "   From 药品库存 A,材料特性 B " & _
                "   Where A.药品ID=[1] And A.性质=1 And A.药品ID=B.材料ID And Nvl(A.批次,0)=[2]  And A.库房ID=[3]"
            
            Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查退货记录的库存是否足够!", lng材料ID, lng批次, lng库房id)
            
            If rstemp.EOF Then
                blnExit = True
            Else
                blnExit = (rstemp!数量 < Abs(Val(mshBill.TextMatrix(lngRow, mCol数量))))
            End If
            
            If blnExit Then
                MsgBox "第" & lngRow & "行的卫材库存数不够，不允许审核！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    Next
    库存检查_退货 = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function ISCheckScalc售价(ByVal bln结算价 As Boolean, ByVal lngRow As Long) As Boolean
    '功能:检查及计算时价的售价
    '参数:bln结算价:true-结算价,False-采购价
    '返回:成功:返回ture,否则返回False
    Dim dbl购价 As Double
    Dim dbl售价 As Double
    Dim dbl加成率 As Double
    Dim dbl数量 As Double, dbl零售金额 As Double, dbl结算金额 As Double
    Dim sng分段售价 As Double
    Dim lng材料ID As Long
    Dim dbl指导差价率 As Double
    Dim dbl比例系数 As Double
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim rstemp As ADODB.Recordset
    
    err = 0: On Error GoTo ErrHand:
    With mshBill
        '存储格式值:最大效期||指导差价率||是否变价||在用分批||库房分批
        If Not (Trim(.TextMatrix(lngRow, mCol原销期)) <> "") Then
            '未输入
            GoTo Calc:
        End If
        If Val(Split(.TextMatrix(lngRow, mCol原销期), "||")(2)) <> 1 Then
            '不是时价卫材,退出
            GoTo Calc:
        End If
        lng材料ID = Val(.TextMatrix(lngRow, 0))
        If lng材料ID = 0 Then
            Exit Function
        End If
        If mbln时价购前销售 Then
            If bln结算价 Then
                Call 重新计算售价(lngRow)
                GoTo Calc:
            End If
        Else
            If bln结算价 = False Then
                Call 重新计算售价(lngRow)
                GoTo Calc:
            End If
        End If
        
        dbl售价 = Val(.TextMatrix(lngRow, mCol售价))
        '如果是按购前价格的话,则需计算加成率不一样
        If mbln时价购前销售 Then
            dbl购价 = Val(.TextMatrix(lngRow, mCol采购价))
        Else
            dbl购价 = Val(.TextMatrix(lngRow, mCol结算价))
        End If
        dbl比例系数 = Val(.TextMatrix(lngRow, mCol比例系数))
        
        '原效期字段下面保存原效期，指导差价，是否变价，在用分批等，格式为：最大效期||指导差价率||是否变价||在用分批||库房分批
        dbl指导差价率 = Val(Split(.TextMatrix(lngRow, mCol原销期), "||")(1))
            
        '由于存在差价让利比的存在,需要按加成率计算,因此将指导差价率转换成加成率计算 公式：加成率=1/(1-差价率)-1
        '如果系统参数为真，则提示用户输入加价率
        If mbln加价率 = True Then
            sngLeft = Pic单据.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
            sngTop = Pic单据.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
            If sngLeft + PicInput.Width > Screen.Width Then
                sngLeft = sngLeft + mshBill.MsfObj.CellWidth - PicInput.Width
            End If
            
            If sngTop + 1700 > Screen.Height Then
                sngTop = sngTop - mshBill.MsfObj.CellHeight - 1700
            End If
            
            With PicInput
                .Top = sngTop
                .Left = sngLeft
                .Visible = True
                .Tag = IIf(bln结算价, "1", "0")
            End With
            Txt加价率 = Val(Replace(.TextMatrix(.Row, mcol加成率), "%", "")) '"15.0000"
            .TextMatrix(lngRow, mCol售价) = Format(校正零售价(dbl购价 * (1 + (Val(Txt加价率) / 100)) + 时价材料零售价(lng材料ID, dbl购价, (Val(Txt加价率) / 100))), mFMT.FM_零售价)
            
'            If dbl售价 <> 0 And dbl购价 <> 0 Then
'                Txt加价率 = Format(计算加成率(lng材料ID, dbl售价, dbl购价), "###0.0000000;-###0.0000000;0;0")
'            End If
            
            Txt加价率.Tag = Txt加价率
            Txt加价率.SetFocus
        ElseIf mbln分段加成率 = True Then
            dbl加成率 = 0 ' Get分段加成率(dbl购价) / 100
            If mint编辑状态 = 8 And dbl售价 <> 0 Then
            Else
                If Get分段加成售价(dbl购价, dbl比例系数, mstrCaption, sng分段售价) = True Then
                    
                    .TextMatrix(lngRow, mCol售价) = Format(校正零售价(sng分段售价 + _
                                                    时价材料零售价(lng材料ID, dbl购价, 0, -1, sng分段售价)) _
                                                    , mFMT.FM_零售价)
                Else
                    ISCheckScalc售价 = False
                    Exit Function
                End If
            End If
        Else 'mbln时价卫材取上次售价 = True优先从上次取，如果没有则按照加成率方式取
            If mbln时价卫材取上次售价 = True Then
                gstrSQL = "Select Nvl(上次售价, 0) As 上次售价 From 材料特性 Where 材料id = [1]"
                Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng材料ID)
                If rstemp!上次售价 > 0 Then
                    .TextMatrix(lngRow, mCol售价) = Format(zlStr.Nvl(rstemp!上次售价, 0) * Val(.TextMatrix(lngRow, mCol比例系数)), mFMT.FM_零售价)
                    If dbl购价 <> 0 Then
                        .TextMatrix(lngRow, mcol加成率) = Format((Val(.TextMatrix(lngRow, mCol售价)) / dbl购价 - 1) * 100, "###0.00") & "%"
                    End If
                Else
                    '由于存在差价让利比的存在,需要按加成率计算,因此将指导差价率转换成加成率计算 公式：加成率=1/(1-差价率)-1
                    If dbl购价 <> 0 Then
                        Txt加价率 = Val(Replace(.TextMatrix(.Row, mcol加成率), "%", "")) '计算加成率(lng材料ID, dbl售价, dbl购价)
                        .TextMatrix(lngRow, mcol加成率) = Format(Txt加价率, "####0.00") & "%"
                    End If
                    If mint编辑状态 = 8 And dbl售价 <> 0 Then
                    Else
                        .TextMatrix(lngRow, mCol售价) = Format(校正零售价(dbl购价 * (1 + (Txt加价率 / 100)) + _
                        时价材料零售价(lng材料ID, dbl购价, (Txt加价率 / 100))), mFMT.FM_零售价)
                    End If
                End If
            Else '3种取售价方式都没有设置时，则按照加成率方式取
                If dbl购价 <> 0 Then
                    Txt加价率 = Val(Replace(.TextMatrix(.Row, mcol加成率), "%", "")) '计算加成率(lng材料ID, dbl售价, dbl购价)
                    .TextMatrix(lngRow, mcol加成率) = Format(Txt加价率, "####0.00") & "%"
                End If
                If mint编辑状态 = 8 And dbl售价 <> 0 Then
                Else
                    .TextMatrix(lngRow, mCol售价) = Format(校正零售价(dbl购价 * (1 + (Txt加价率 / 100)) + _
                    时价材料零售价(lng材料ID, dbl购价, (Txt加价率 / 100))), mFMT.FM_零售价)
                End If
            End If
            
        End If
Calc:
        dbl数量 = Val(.TextMatrix(lngRow, mCol数量))
        dbl售价 = Val(.TextMatrix(lngRow, mCol售价))
        dbl购价 = Val(.TextMatrix(lngRow, mCol结算价))
        dbl零售金额 = dbl数量 * dbl售价
        dbl结算金额 = dbl数量 * dbl购价
        .TextMatrix(lngRow, mCol售价金额) = Format(dbl零售金额, mFMT.FM_金额)
        .TextMatrix(lngRow, mCol结算金额) = Format(dbl结算金额, mFMT.FM_金额)
        .TextMatrix(lngRow, mCol发票金额) = IIf(Trim(Trim(.TextMatrix(lngRow, mCol发票号))) = "", "", .TextMatrix(lngRow, mCol结算金额))
        .TextMatrix(lngRow, mCol差价) = Format(dbl零售金额 - dbl结算金额, mFMT.FM_金额)
        
        ''刘兴宏:零售价处理
        Call 计算零售价及零售差价(lngRow)
    End With
    ISCheckScalc售价 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 重新计算售价(ByVal lngRow As Long) As Boolean
    Dim dbl数量 As Double, dbl扣率 As Double, dbl售价 As Double, lng材料ID As Long
    Dim dbl采购价 As Double, dbl结算价 As Double, dbl加成率 As Double
    Dim sng分段售价 As Double
    Dim rstemp As ADODB.Recordset
    
    With mshBill
            dbl数量 = Val(.TextMatrix(lngRow, mCol数量))
            dbl扣率 = Val(.TextMatrix(lngRow, mCol扣率))
            dbl售价 = Val(.TextMatrix(lngRow, mCol售价))
            dbl结算价 = Val(.TextMatrix(lngRow, mCol结算价))
            dbl采购价 = Val(.TextMatrix(lngRow, mCol采购价))
            lng材料ID = Val(.TextMatrix(lngRow, 0))
                
            If mbln加价率 Then
                mdbl加价率 = 15
                If dbl售价 <> 0 And dbl结算价 <> 0 Then
                    If Val(Replace(.TextMatrix(lngRow, mcol加成率), "%", "")) >= 0 Then '"X.XXX%"小数不为0就会报类型不匹配错误
                        mdbl加价率 = Val(Replace(.TextMatrix(lngRow, mcol加成率), "%", "")) 'Val(Split(.TextMatrix(lngRow, mcol加成率), "%")(0))
                    Else
                        mdbl加价率 = 计算加成率(lng材料ID, dbl售价, IIf(mbln时价购前销售, dbl采购价, dbl结算价))
                    End If
                End If
             End If
                
            '对时价材料的处理
            If .TextMatrix(lngRow, mCol原销期) <> "" Then
                '重新计算零售价、差价
                '
                '存储格式值:最大效期||指导差价率||是否变价||在用分批||库房分批
                If Split(.TextMatrix(lngRow, mCol原销期), "||")(2) = 1 Then
                    '由于存在差价让利比的存在,需要按加成率计算,因此将指导差价率转换成加成率计算 公式：加成率=1/(1-差价率)-1
                    If mbln加价率 Then
                        If mint编辑状态 = 8 And dbl售价 <> 0 Then
                        Else
                            If mbln分段加成率 Then
                                If Get分段加成售价(IIf(mbln时价购前销售, dbl采购价, dbl结算价), Val(.TextMatrix(.Row, mCol比例系数)), mstrCaption, sng分段售价) = True Then
                                    .TextMatrix(lngRow, mCol售价) = Format(校正零售价(sng分段售价 + _
                                                                    时价材料零售价(lng材料ID, IIf(mbln时价购前销售, dbl采购价, dbl结算价), 0, -1, sng分段售价)) _
                                                                    , mFMT.FM_零售价)
                                End If
                            Else
                                .TextMatrix(lngRow, mCol售价) = Format(校正零售价(IIf(mbln时价购前销售, dbl采购价, dbl结算价) * (1 + (mdbl加价率 / 100)) + _
                                                                时价材料零售价(lng材料ID, IIf(mbln时价购前销售, dbl采购价, dbl结算价), (mdbl加价率 / 100))) _
                                                                , mFMT.FM_零售价)
                            End If
                        End If
                        .TextMatrix(lngRow, mCol售价金额) = Format(Val(.TextMatrix(lngRow, mCol售价)) * dbl数量, mFMT.FM_金额)
                        .TextMatrix(lngRow, mCol差价) = Format(IIf(.TextMatrix(lngRow, mCol售价金额) = "", 0, .TextMatrix(lngRow, mCol售价金额)) - IIf(.TextMatrix(lngRow, mCol结算金额) = "", 0, .TextMatrix(lngRow, mCol结算金额)), mFMT.FM_金额)
                    Else
                        If mbln分段加成率 Then
                            dbl加成率 = 0 ' Get分段加成率(IIf(mbln时价购前销售, dbl采购价, dbl结算价)) / 100
                        Else
                            dbl加成率 = Val(Replace(.TextMatrix(lngRow, mcol加成率), "%", "")) / 100
                        End If
                        If mint编辑状态 = 8 And dbl售价 <> 0 Then
                        Else
                            If mbln分段加成率 Then
                                If Get分段加成售价(IIf(mbln时价购前销售, dbl采购价, dbl结算价), Val(.TextMatrix(.Row, mCol比例系数)), mstrCaption, sng分段售价) = True Then
                                    .TextMatrix(lngRow, mCol售价) = Format(校正零售价(sng分段售价 + _
                                                                    时价材料零售价(lng材料ID, IIf(mbln时价购前销售, dbl采购价, dbl结算价), 0, -1, sng分段售价)) _
                                                                    , mFMT.FM_零售价)
                                End If
                            Else
                                If mbln时价卫材取上次售价 = True Then '取上次售价
                                    gstrSQL = "Select Nvl(上次售价, 0) As 上次售价 From 材料特性 Where 材料id = [1]"
                                    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng材料ID)
                                    If rstemp!上次售价 > 0 Then
                                        .TextMatrix(lngRow, mCol售价) = Format(zlStr.Nvl(rstemp!上次售价, 0) * Val(.TextMatrix(lngRow, mCol比例系数)), mFMT.FM_零售价)
                                        If IIf(mbln时价购前销售, dbl采购价, dbl结算价) <> 0 Then
                                            .TextMatrix(lngRow, mcol加成率) = Format((Val(.TextMatrix(lngRow, mCol售价)) / IIf(mbln时价购前销售, dbl采购价, dbl结算价) - 1) * 100, "###0.00") & "%"
                                        End If
                                    Else
                                        dbl加成率 = Val(Replace(.TextMatrix(lngRow, mcol加成率), "%", "")) / 100
                                        
                                        .TextMatrix(lngRow, mCol售价) = Format(校正零售价(IIf(mbln时价购前销售, dbl采购价, dbl结算价) * (1 + dbl加成率) + _
                                                                        时价材料零售价(lng材料ID, IIf(mbln时价购前销售, dbl采购价, dbl结算价), dbl加成率)) _
                                                                        , mFMT.FM_零售价)
                                    End If
                                Else '参数都为选择按加成率计算
                                    dbl加成率 = Val(Replace(.TextMatrix(lngRow, mcol加成率), "%", "")) / 100
                                        
                                    .TextMatrix(lngRow, mCol售价) = Format(校正零售价(IIf(mbln时价购前销售, dbl采购价, dbl结算价) * (1 + dbl加成率) + _
                                                                    时价材料零售价(lng材料ID, IIf(mbln时价购前销售, dbl采购价, dbl结算价), dbl加成率)) _
                                                                    , mFMT.FM_零售价)
                                End If
                            End If
                        End If
                        .TextMatrix(lngRow, mCol售价金额) = Format(dbl数量 * Val(.TextMatrix(lngRow, mCol售价)), mFMT.FM_金额)
                    End If
                End If
            End If
            ''刘兴宏:零售价处理
            Call 计算零售价及零售差价(lngRow)
    End With
End Function


Private Sub RefreshBill()
    '以最新价格最新单据相关数据，用于单据审核时
    Dim lngRow As Long, lngRows As Long, lng材料ID As Long
    Dim dbl数量 As Double, dbl成本价 As Double, dbl成本金额 As Double, dbl零售价 As Double, dbl零售金额 As Double, dbl差价 As Double
    Dim rsprice As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = " Select 收费细目ID,nvl(现价,0) 现价 From 收费价目 " & _
            " Where (终止日期 Is NULL Or sysdate Between 执行日期 And nvl(终止日期,to_date('3000-01-01','yyyy-MM-dd')))" & _
            GetPriceClassString("")
            
    gstrSQL = "Select A.序号,A.药品ID ,B.现价 From 药品收发记录 A,(" & gstrSQL & ") B,收费项目目录 C" & _
            " Where A.单据=15  And A.NO=[1] And A.药品ID=B.收费细目ID And C.ID=B.收费细目ID And Round(A.零售价,7)<>Round(B.现价,7) And Nvl(C.是否变价,0)=0" & _
            " Order by A.序号"
    
    Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "[取当前价格]", txtNO.Text)
    
    If rsprice.EOF Then Exit Sub
    
    lngRows = mshBill.Rows - 1
    For lngRow = 1 To lngRows
        lng材料ID = Val(mshBill.TextMatrix(lngRow, 0))
        If lng材料ID <> 0 Then
            rsprice.Filter = "药品ID=" & lng材料ID
            If rsprice.RecordCount <> 0 Then
                '以当前最新价格最新单据相关数据（单价、零售金额、差价）
                dbl零售价 = rsprice!现价 * Val(mshBill.TextMatrix(lngRow, mCol比例系数))
                dbl成本价 = Val(mshBill.TextMatrix(lngRow, mCol结算价))
                dbl数量 = Val(mshBill.TextMatrix(lngRow, mCol数量))
                
                dbl成本金额 = dbl成本价 * dbl数量
                dbl零售金额 = dbl零售价 * dbl数量
                dbl差价 = dbl零售金额 - dbl成本金额
                
                mshBill.TextMatrix(lngRow, mCol售价) = Format(dbl零售价, mFMT.FM_零售价)
                mshBill.TextMatrix(lngRow, mCol售价金额) = Format(dbl零售金额, mFMT.FM_金额)
                mshBill.TextMatrix(lngRow, mCol差价) = Format(dbl差价, mFMT.FM_金额)
                ''刘兴宏:零售价处理
                Call 计算零售价及零售差价(lngRow)
            End If
        End If
    Next
    rsprice.Filter = 0
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Function CheckProvider() As Boolean
    Dim lngRow As Long
    Dim str材料 As String
    Dim str招标材料 As String
    Dim rstemp As New ADODB.Recordset
    '检查供应商是否是招标材料的中标单位
    On Error GoTo ErrHandle
    str材料 = ""
    With mshBill
        For lngRow = 1 To .Rows - 1
            If Val(.TextMatrix(lngRow, 0)) <> 0 Then
                str材料 = str材料 & "," & Val(.TextMatrix(lngRow, 0))
            End If
        Next
        If str材料 <> "" Then str材料 = Mid(str材料, 2)
    End With
    
    '以招标材料数减去共有同一个中标单位的招标材料数，如果无记录，则说明正确，否则按记录中的材料ID提示不是合法的中标单位
    gstrSQL = " Select a.材料ID From 材料特性 a,Table(Cast(f_Num2List([2]) As zlTools.t_NumList)) b " & _
              " Where a.材料ID=b.Column_Value And Nvl(a.招标材料,0)=1" & _
              " Minus" & _
              " Select A.材料ID From " & _
              "     (Select a.材料ID From 材料特性 a,Table(Cast(f_Num2List([2]) As zlTools.t_NumList)) b  Where a.材料ID=b.Column_Value And Nvl(a.招标材料,0)=1) A,材料中标单位 B" & _
              " Where A.材料ID=B.材料ID And B.单位ID=[1]"
    gstrSQL = " Select '['||A.编码||']'||A.名称 材料名称 " & _
              " From " & _
              "     (Select A.材料ID,C.编码,Nvl(B.名称,C.名称) 名称" & _
              "     From (" & gstrSQL & ") A,收费项目别名 B,收费项目目录 C" & _
              "     Where A.材料ID=B.收费细目ID(+) and A.材料ID=C.ID" & _
              "     and B.性质(+)=3 and B.码类(+)=1) A"
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "[判断是否是中标单位处采购]", Val(txtProvider.Tag), str材料)
    
    With rstemp
        str材料 = ""
        Do While Not .EOF
            str材料 = str材料 & "、" & rstemp!材料名称
            .MoveNext
        Loop
        If str材料 <> "" Then str材料 = Mid(str材料, 2)
    End With
    
    If str材料 <> "" Then
        If mbln非中标单位入库 Then
            If MsgBox("该供货单位不是以下招标卫材的中标单位,是否继续？" & vbCrLf & str材料, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        Else
            MsgBox "该供货单位不是以下招标卫材的中标单位：" & vbCrLf & str材料, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    CheckProvider = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub chk转入移库_Click()
    Dim blnEnabled As Boolean
    blnEnabled = (chk转入移库.Value = 1)
    
    cboType.Enabled = blnEnabled
    cboEnterStock.Enabled = blnEnabled
    txtDraw.Enabled = blnEnabled
    cmdDraw.Enabled = blnEnabled
    txtDrawPerson.Enabled = blnEnabled
    cmdDrawPerson.Enabled = blnEnabled
    lbl领用人.Enabled = blnEnabled
End Sub

Private Function Check移库(ByRef blnExit As Boolean) As Boolean
    '--------------------------------------------------------------------------------------------------
    '功能:检查外购入库单转移到其他库房时，对材料的移库条件进行检查:
    '       1)对存储库房进行检查
    '       2)对跟踪在用进行检查
    '参数:blnExit-必需退出
    '--------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim rstemp As New ADODB.Recordset
    Dim strTmp As String
    Dim bln无移库材料 As Boolean
    Dim bln具备跟踪材料 As Boolean
    
    On Error GoTo ErrHand
    bln无移库材料 = True
    '检查是否退货单
    If mbln退货 Then
        If MsgBox("退库单，不能使用导入移库的功能。确认入库，选择<是>；放弃审核，选择<否>。", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            blnExit = True
        Else
            blnExit = False
        End If
        Check移库 = False
        Exit Function
    End If
    If chk转入移库.Value <> 1 Then Check移库 = False: Exit Function
    If cboEnterStock.ListIndex < 0 Then Check移库 = False: Exit Function
    
    bln具备跟踪材料 = 判断只具备发料部门(cboEnterStock.ItemData(cboEnterStock.ListIndex))
    
    strTmp = ""
    With mshBill
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, 0)) <> "" Then
                '检查负数出库部分,不能进行导入
                If Val(.TextMatrix(i, mCol数量)) < 0 Then
                    If MsgBox("名称为“" & .TextMatrix(i, mCol诊疗) & "”的卫生材料为负数，不能使用导入移库的功能。确认入库，选择<是>；放弃审核，选择<否>。", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        blnExit = True
                    Else
                        blnExit = False
                    End If
                    Check移库 = False
                    Exit Function
                End If
                
                '检查是否设置存储库房
                gstrSQL = "select 收费细目ID from 收费执行科室 where 收费细目ID=[1] and 执行科室ID=[2]  "
                Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "[判断存储库房]", Val(.TextMatrix(i, 0)), cboEnterStock.ItemData(cboEnterStock.ListIndex))
                If rstemp.RecordCount = 0 Then
                     strTmp = strTmp & "名称:" & mshBill.TextMatrix(i, mCol诊疗) & " 规格:" & mshBill.TextMatrix(i, mCol规格) & vbCrLf
                Else
                    If bln具备跟踪材料 Then
                        '判断跟踪材料
                        gstrSQL = "Select 跟踪在用 From 材料特性 where 材料id=[1]"
                        Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "[判断跟踪在用]", Val(.TextMatrix(i, 0)))
                        If Not rstemp.EOF Then
                            If Val(Nvl(rstemp!跟踪在用)) = 1 Then
                                bln无移库材料 = False
                            Else
                                strTmp = strTmp & "名称:" & mshBill.TextMatrix(i, mCol诊疗) & " 规格:" & mshBill.TextMatrix(i, mCol规格) & vbCrLf
                            End If
                        End If
                    Else
                        bln无移库材料 = False
                    End If
                End If
            End If
        Next
    End With
    
    If strTmp <> "" Then
        If bln无移库材料 Then
            ShowMsgBox "本次入库材料没有设置存储库房或跟踪材料，将不能移库到[" & cboEnterStock.Text & "]。"
            Check移库 = False
            Exit Function
        End If
        ShowMsgBox "以下材料没有设置存储库房，将不能移库到[" & cboEnterStock.Text & "] ：" & vbCrLf & strTmp & vbCrLf & "其余材料可以导入移库。"
    End If
    Check移库 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Function RestoreBILLWidthSet() As Boolean
    Dim strWidth As String, strText As String
    Dim arrText As Variant, i As Integer
    Dim arrWidth As Variant
    
    On Error Resume Next
    
    If Val(zlDatabase.GetPara("使用个性化风格", 0, 0)) = 0 Then
        RestoreBILLWidthSet = True: Exit Function
    End If
   
    '检查是否需要恢复
    
    strWidth = zlDatabase.GetPara("单据列宽", glngSys, mlngModule)
    strText = zlDatabase.GetPara("单据列头文本", glngSys, mlngModule)
    
    If strText <> "" Then
        '固定行变了,不恢复而使用缺省
        '列数变了,不恢复而使用缺省
        arrText = Split(strText, ",")
        arrWidth = Split(strWidth, ",")
        For i = 0 To UBound(arrText) + 1
            Call SetBillColWidth(arrText(i), arrWidth(i))
        Next
    End If
    RestoreBILLWidthSet = True
End Function
Private Sub SetBillColWidth(ByVal strColName As String, ByVal lngWidth As Long)
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置列宽度
    '参数:strName-表头列名字
    '     lngwidth-列宽
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCol As Integer, i As Integer
    If lngWidth <= 0 Then Exit Sub
    With mshBill
        intCol = -1
        For i = 0 To .Cols - 1
            If .TextMatrix(0, i) = strColName Then
                intCol = i
                Exit For
            End If
        Next
        If intCol = -1 Then Exit Sub
        If .ColWidth(intCol) <= 0 Then Exit Sub
        .ColWidth(intCol) = lngWidth
    End With
End Sub
Private Sub SaveBILLWidth()
    Dim strWidth As String, strText As String, i As Integer
        
    On Error Resume Next
    
    strWidth = "": strText = ""
    For i = 0 To mshBill.Cols - 1
        strWidth = strWidth & "," & mshBill.ColWidth(i)
        strText = strText & "," & mshBill.TextMatrix(0, i)
    Next
    zlDatabase.SetPara "单据列宽", Mid(strWidth, 2), glngSys, mlngModule
    zlDatabase.SetPara "单据列头文本", Mid(strText, 2), glngSys, mlngModule
End Sub
Private Function Set操作流程Update(Optional blnRowInput As Boolean = False) As Boolean
    '----------------------------------------------------------------------------------------------------------
    '功能:根据流程,设置相应的修改项目
    '参数:blnRowInput-是否根据行来进行动态判断输入情况
    '返回:设置成功,返回True,否则返回False
    '编制:刘兴宏
    '日期:2007/05/30
    '----------------------------------------------------------------------------------------------------------
    Dim int环节 As Integer, intCol As Integer
    Dim mrs环节控制 As New ADODB.Recordset
    Dim arr内容 As Variant
    Dim str内容 As String
    
    On Error GoTo ErrHandle
    '1.新增；2、修改；3、验收；4、查看；5、修改发票；6、冲销；
    '7、财务审核（冲销、产生新单据并审核；已付款的单据不允许财务审核，同样，财务审核后的单据不允许冲销）;
    '8、卫材库退货,9-核查
    
    If mint编辑状态 = 3 Or mint编辑状态 = 9 Or mint编辑状态 = 7 Then
        '  1.核查，2.审核，3.帐务审核
        int环节 = Decode(mint编辑状态, 3, 2, 9, 1, 3)
    Else
        Set操作流程Update = True
        Exit Function
    End If
    
    gstrSQL = "Select 环节,','||内容||',' as 内容 From 单据环节控制 where 单据=[1] and 环节=[2] order by 环节"
    '内容:存放该环节可修改的项目，格式为"项目1,项目2,..."，可选项目为"采购价，扣率，结算价，结算金额，售价，发票号，发票日期，发票金额,发票代码"。以后可扩充。
    
    If mrs环节控制 Is Nothing Then
        Set mrs环节控制 = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, 15, int环节)
    ElseIf mrs环节控制.State <> 1 Then
        Set mrs环节控制 = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, 15, int环节)
    End If
    '如果没有数据控制则都不能修改
    If mrs环节控制.RecordCount = 0 Then
        For intCol = 0 To mshBill.Cols - 1
            mshBill.ColData(intCol) = 0
        Next
        Exit Function
    End If
    With mshBill
        If blnRowInput Then
            str内容 = zlStr.Nvl(mrs环节控制!内容)
            If InStr(1, str内容, ",售价,") > 0 Then
                '如果是时价卫材，则允许输入售价
                '存储格式值:最大效期||指导差价率||是否变价||在用分批||库房分批
                If .TextMatrix(.Row, mCol原销期) <> "" Then
                    If Split(.TextMatrix(.Row, mCol原销期), "||")(2) = 1 Then
                        .ColData(mCol售价) = IIf(mbln时价卫材直接确定售价, 4, 5)
                        .ColData(mcol零售价) = IIf(mbln退货 Or mint编辑状态 = 8, 5, 4)
                    Else
                        .ColData(mCol售价) = 0
                        .ColData(mcol零售价) = 0
                    End If
                Else
                     .ColData(mCol售价) = 0
                     .ColData(mcol零售价) = 0
                End If
            Else
                .ColData(mCol售价) = 0
                .ColData(mcol零售价) = 0
            End If
            If Trim(.TextMatrix(.Row, mCol发票号)) = "" Then
                .ColData(mCol发票日期) = 0
                .ColData(mCol发票金额) = 0
                .ColData(mcol发票代码) = 0
            Else
                If InStr(1, str内容, ",发票日期,") > 0 Then
                    .ColData(mCol发票日期) = 2
                Else
                    .ColData(mCol发票日期) = 0
                End If
                If InStr(1, str内容, ",发票代码,") > 0 Then
                    .ColData(mcol发票代码) = 4
                Else
                    .ColData(mcol发票代码) = 0
                End If
                If InStr(1, str内容, ",发票金额,") > 0 Then
                    .ColData(mCol发票金额) = 4
                Else
                    .ColData(mCol发票金额) = 0
                End If
            End If
            Set操作流程Update = True
            Exit Function
        End If
        
        For intCol = 0 To .Cols - 1
            .ColData(intCol) = 0
        Next
        If mrs环节控制.EOF = False Then
            '"采购价，扣率，结算价，结算金额，售价，发票号，发票日期，发票金额,发票代码
            arr内容 = Split(zlStr.Nvl(mrs环节控制!内容), ",")
            For intCol = 0 To UBound(arr内容)
                Select Case arr内容(intCol)
                Case "采购价"
                    If mbln不强制控制指导价格 = False Then
                        .ColData(mCol指导批发价) = IIf(mbln修改批发价, 4, 0)
                    End If
                    .ColData(mCol采购价) = 4
                Case "扣率"
                    .ColData(mCol扣率) = 4
                Case "结算价"
                    .ColData(mCol结算价) = 4
                Case "结算金额"
                    .ColData(mCol结算金额) = 4
                Case "售价"
                    .ColData(mCol售价) = 4
                Case "零售价"
                    .ColData(mcol零售价) = IIf(mbln退货 Or mint编辑状态 = 8, 5, 4)
                Case "发票号"
                    mshBill.ColData(mCol发票号) = 4
                Case "发票代码"
                    mshBill.ColData(mcol发票代码) = 4
                Case "发票日期"
                    mshBill.ColData(mCol发票日期) = 2
                Case "发票金额"
                   .ColData(mCol发票金额) = 4
                End Select
            Next
        End If
        '重新定位
        For int环节 = 0 To mshBill.Cols - 1
            If .ColData(int环节) = 4 Or .ColData(int环节) = 2 Then
                .LocateCol = int环节
                Exit For
            End If
        Next
    End With
    Set操作流程Update = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function SelectItem(ByVal objCtl As Control, ByVal strKey As String, Optional bln人员 As Boolean = False) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:多功能选择器
    '入参:objCtl-文本框控件
    '     strKey-要搜索的值
    '     bln人员-是否人员选择
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-10-27 10:37:40
    '-----------------------------------------------------------------------------------------------------------
    Dim blnCancel As Boolean, lngH As Long, strTittle As String
    Dim vRect As RECT, sngX As Single, sngY As Single
    Dim rstemp  As ADODB.Recordset
    Dim bytStyle As Byte
    Dim str站点限制 As String
    
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
    
    str站点限制 = GetDeptStationNode(cboStock.ItemData(cboStock.ListIndex))
    
    bytStyle = 0
    If bln人员 Then
        strTittle = "人员选择器"
        If strKey = "" Then
            gstrSQL = "" & _
                    "   Select ID, 编号,简码,姓名 From 人员表 a " & _
                    "   Where exists(select 1 from 部门人员 where 人员id=a.id and 部门id=[2]) " & _
                    "       and (a.撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) " & _
                    "       and (a.站点=[3] or a.站点 is null) " & _
                    "   order by 编号"
        Else
            gstrSQL = "" & _
                    "   Select ID, 编号,简码,姓名 From 人员表 a " & _
                    "   Where (姓名 like [1] or  编号  like [1] or  简码  like  upper([1])) " & _
                    "       and (a.站点=[3] or a.站点 is null) " & _
                    "       and exists(select 1 from 部门人员 where 人员id=a.id and 部门id=[2]) " & _
                    "       and (a.撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null)" & _
                    "   order by 编号"
        End If
    Else
        strTittle = "部门选择器"
        gstrSQL = "SELECT a.id,a.上级ID,a.编码,a.名称,a.简码 " & _
                  "FROM 部门表 a " & _
                  "Where ( TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01' or a.撤档时间 is null ) " & _
                  IIf(str站点限制 <> "", " And a.站点 = [3] ", "")
    
        If strKey <> "" Then
            gstrSQL = gstrSQL & _
                      " And (a.简码 like upper([1]) Or a.编码 like upper([1]) or a.名称 like [1])" & _
                      " Order by 编码 "
        Else
            If gstrNodeNo = "-" Then
                '没有站点号,以树型显示
                gstrSQL = gstrSQL & " start with 上级id is null connect by prior id=上级id "
                bytStyle = 1
            Else
                '存在站点，主要是可能上级设置了站点编号，而下级未设置的情况，因此只能用列表方式进行处理
                gstrSQL = gstrSQL & " Order by 编码 "
            End If
        End If
    End If
    
    strKey = GetMatchingSting(strKey, False)
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Then
        Call CalcPosition(sngX, sngY, objCtl)
        lngH = objCtl.CellHeight
        sngY = sngY - lngH
    Else
        vRect = zlControl.GetControlRect(objCtl.hwnd)
        lngH = objCtl.Height
        sngX = vRect.Left - 15
        sngY = vRect.Top
    End If
    
    Set rstemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, bytStyle, strTittle, False, "", "", _
                    False, False, True, sngX, sngY, lngH, blnCancel, False, False, _
                    strKey, Val(txtDraw.Tag), str站点限制)
    If blnCancel = True Then
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    
    If rstemp Is Nothing Then
        ShowMsgBox "没有找到满足条件的内容,请检查!"
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Then
        With objCtl
            If bln人员 Then
                .TextMatrix(.Row, .Col) = zlStr.Nvl(rstemp!姓名)
                .Cell(flexcpData, .Row, .Col) = zlStr.Nvl(rstemp!姓名)
            Else
                .TextMatrix(.Row, .Col) = zlStr.Nvl(rstemp!编码) & "-" & zlStr.Nvl(rstemp!名称)
                .Cell(flexcpData, .Row, .Col) = zlStr.Nvl(rstemp!Id)
            End If
        End With
    Else
        If bln人员 Then
            objCtl.Text = zlStr.Nvl(rstemp!姓名)
            objCtl.Tag = zlStr.Nvl(rstemp!姓名)
        Else
            objCtl.Text = zlStr.Nvl(rstemp!编码) & "-" & zlStr.Nvl(rstemp!名称)
            objCtl.Tag = IIf(bln人员, zlStr.Nvl(rstemp!名称), zlStr.Nvl(rstemp!Id))
        End If
        zlControl.ControlSetFocus objCtl, True
        OS.PressKey vbKeyTab
    End If
    SelectItem = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdDraw_Click()
    If SelectItem(txtDraw, "") = False Then Exit Sub
End Sub
Private Sub txtDraw_Change()
    txtDraw.Tag = ""
    txtDrawPerson.Text = ""
    mblnChange = True
End Sub

Private Sub txtDraw_GotFocus()
    OS.OpenIme False
    zlControl.TxtSelAll txtDraw
End Sub
Private Sub txtDraw_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If txtDraw.Tag <> "" Then OS.PressKey vbKeyTab: Exit Sub
    If SelectItem(txtDraw, Trim(txtDraw.Text)) = False Then Exit Sub
End Sub

Private Function 当前仅为库房() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:判断当前库房仅为库房
    '入参:
    '出参:
    '返回:返回true表示仅为库房,否则为(发料部门或制剂室)
    '编制:刘兴洪
    '日期:2008-12-03 11:23:18
    '-----------------------------------------------------------------------------------------------------------
    Dim rstemp As New ADODB.Recordset
    On Error GoTo ErrHandle
    gstrSQL = "" & _
        "   SELECT count(*)" & _
        "   From 部门性质说明 " & _
        "   WHERE ((工作性质 LIKE '发料部门') OR (工作性质 LIKE '制剂室')) " & _
        "        AND 部门id =[1]"
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex))
    If rstemp.Fields(0) > 0 Then
        当前仅为库房 = False
        mbln库房 = False
    Else
        当前仅为库房 = True
        mbln库房 = True
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub vsfCostlyInfo_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsfCostlyInfo
        Select Case .Col
            Case .ColIndex("科室")
                .ColComboList(.Col) = "..."
            Case .ColIndex("病人姓名")
                .ColComboList(.Col) = "..."
        End Select
    End With
End Sub

Private Sub vsfCostlyInfo_BeforeDataRefresh(Cancel As Boolean)
    '保存到数据集
    If vsfCostlyInfo.Visible = False Then Exit Sub
    If vsfCostlyInfo.Rows < 2 Then Exit Sub
    If vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("科室")) = "" Then
        Cancel = True
        MsgBox "'科室'信息未录入！", vbCritical, gstrSysName
        Exit Sub
    End If
    If vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("病人姓名")) = "" Then
        Cancel = True
        MsgBox "'病人姓名'信息未录入！", vbCritical, gstrSysName
        Exit Sub
    End If
    If vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("住院号")) = "" Then
        Cancel = True
        MsgBox "'住院号'信息未录入！", vbCritical, gstrSysName
        Exit Sub
    End If
    If vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("床号")) = "" Then
        Cancel = True
        MsgBox "'床号'信息未录入！", vbCritical, gstrSysName
        Exit Sub
    End If
End Sub

Private Sub vsfCostlyInfo_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If NewRow > 1 Then Cancel = True
End Sub

Private Sub CostlyInfo_Refresh(ByVal lngId As Long, ByVal blnCostly As Boolean)
'高值材料信息刷新
    If lngId < 1 Then Exit Sub
    If mshBill.TextMatrix(lngId, 0) = "" Then Exit Sub
    If mrsCostlyInfo Is Nothing Then Exit Sub
    If mrsCostlyInfo.RecordCount <= 0 Then
        If blnCostly Then
            mrsCostlyInfo.AddNew
            mrsCostlyInfo!sn = lngId
            mrsCostlyInfo.Update
        Else
            Exit Sub
        End If
    End If
    If mrsCostlyInfo.RecordCount > 0 Then mrsCostlyInfo.MoveFirst
    mrsCostlyInfo.Find "SN=" & lngId
    If mrsCostlyInfo.EOF Then
        If blnCostly Then
            mrsCostlyInfo.AddNew
            mrsCostlyInfo!sn = lngId
            mrsCostlyInfo.Update
        End If
        vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("科室ID")) = ""
        vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("科室")) = ""
        vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("病人姓名")) = ""
        vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("住院号")) = ""
        vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("床号")) = ""
    Else
        If blnCostly = False Then
            mrsCostlyInfo.Delete
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("科室ID")) = ""
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("科室")) = ""
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("病人姓名")) = ""
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("住院号")) = ""
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("床号")) = ""
        Else
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("科室ID")) = IIf(IsNull(mrsCostlyInfo!Id), "", mrsCostlyInfo!Id)
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("科室")) = IIf(IsNull(mrsCostlyInfo!科室), "", mrsCostlyInfo!科室)
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("病人姓名")) = IIf(IsNull(mrsCostlyInfo!病人姓名), "", mrsCostlyInfo!病人姓名)
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("住院号")) = IIf(IsNull(mrsCostlyInfo!住院号), "", mrsCostlyInfo!住院号)
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("床号")) = IIf(IsNull(mrsCostlyInfo!床号), "", mrsCostlyInfo!床号)
        End If
    End If
    
    vsfCostlyInfo.Visible = blnCostly
    Call Form_Resize
End Sub

Private Function IsCostly(ByVal lngMaterialID As Long) As Boolean
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    Set rsTmp = zlDatabase.OpenSQLRecord("select count(高值材料) 高值材料 from 材料特性 where 材料id=[1] and nvl(高值材料,'')='1'", mstrCaption, lngMaterialID)
    If rsTmp.RecordCount = 1 And rsTmp!高值材料 = "1" Then
        IsCostly = True
        mbln高值卫材 = True
    Else
        mbln高值卫材 = False
    End If
    rsTmp.Close
    
    cmdCopy.Enabled = mbln高值卫材
    txtCopy.Enabled = mbln高值卫材
    lblCopy.Enabled = mbln高值卫材
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub vsfCostlyInfo_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    With vsfCostlyInfo
        Select Case Col
            Case .ColIndex("科室")
                Call Comm_Selecter("%" & .EditText & "%", 1)
            Case .ColIndex("病人姓名")
                'If .EditText = "" Then
                    Call Comm_Selecter(.TextMatrix(1, .ColIndex("科室ID")), 2)
                'End If
        End Select
    End With
End Sub

Private Sub vsfCostlyInfo_EnterCell()
    With vsfCostlyInfo
        Select Case .Col
            Case .ColIndex("科室"), .ColIndex("病人姓名")
                .ColComboList(.Col) = "..."
        End Select
    End With
End Sub

Private Sub vsfCostlyInfo_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsfCostlyInfo
        Select Case .Col
            Case .ColIndex("科室")
                .ColComboList(.Col) = ""
            Case .ColIndex("病人姓名")
                .ColComboList(.Col) = ""
        End Select
    End With
End Sub

Private Sub vsfCostlyInfo_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With vsfCostlyInfo
            Select Case Col
                Case .ColIndex("科室")
                    Call Comm_Selecter("%" & .EditText & "%", 1)
                Case .ColIndex("病人姓名")
                    'If .EditText = "" Then
                        Call Comm_Selecter(.TextMatrix(1, .ColIndex("科室id")), 2)
                    'End If
            End Select
            If Col = .ColIndex("床号") Then
                .Col = 1
            Else
                .Col = Col + 1
            End If
        End With
        Exit Sub
    
    ElseIf vsfCostlyInfo.ColIndex("住院号") = Col Then
        If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            KeyAscii = 0
        End If
    Else
        If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then KeyAscii = KeyAscii - 32
    End If
End Sub

Private Sub vsfCostlyInfo_Validate(Cancel As Boolean)
    '保存数据
    With vsfCostlyInfo
        '定位
        If mrsCostlyInfo.RecordCount > 0 Then
            mrsCostlyInfo.MoveFirst
            mrsCostlyInfo.Find "SN=" & mshBill.TextMatrix(mshBill.Row, 1)
        End If
        If mrsCostlyInfo.EOF Then
            '新增
            mrsCostlyInfo.AddNew
            mrsCostlyInfo!sn = mshBill.TextMatrix(mshBill.Row, 1)
        End If
        
        mrsCostlyInfo!Id = IIf(.TextMatrix(1, 5) = "", 0, .TextMatrix(1, 5))
        mrsCostlyInfo!科室 = .TextMatrix(1, 1)
        mrsCostlyInfo!病人姓名 = .TextMatrix(1, 2)
        mrsCostlyInfo!住院号 = IIf(.TextMatrix(1, 3) = "", Null, .TextMatrix(1, 3))
        mrsCostlyInfo!床号 = .TextMatrix(1, 4)
        mrsCostlyInfo.Update
    End With
End Sub

Private Sub RecountSN(ByVal lngRow As Long)
'调整对应高值材料的SN
    Dim i As Long, lngMax As Long
    If mrsCostlyInfo.RecordCount <= 0 Then Exit Sub
    mrsCostlyInfo.MoveFirst
    Do While Not mrsCostlyInfo.EOF
        If lngMax < mrsCostlyInfo!sn Then
            lngMax = mrsCostlyInfo!sn
        End If
        mrsCostlyInfo.MoveNext
    Loop
    
    If lngRow >= lngMax Then Exit Sub
    
    For i = lngRow + 1 To lngMax
        With mrsCostlyInfo
            .MoveFirst
            .Find "SN=" & i
            If Not .EOF Then
                mrsCostlyInfo!sn = mrsCostlyInfo!sn - 1
                mrsCostlyInfo.Update
            End If
        End With
    Next
End Sub

Private Sub Comm_Selecter(ByVal strParam As String, ByVal bytIndex As Byte)
    Dim rstemp As ADODB.Recordset
    Dim sngX As Single, sngY As Single, lngH As Long
    Dim blnCancel As Boolean
    Dim strsql As String
    Dim rectTmp As RECT
    
    If bytIndex < 3 Then
        Call CalcPosition(sngX, sngY, vsfCostlyInfo)
        lngH = vsfCostlyInfo.CellHeight
    Else
        rectTmp = zlControl.GetControlRect(Me.txtTypeVar.hwnd)
        sngX = rectTmp.Left
        sngY = rectTmp.Top + Me.txtTypeVar.Height
        lngH = Me.txtTypeVar.Height
    End If
    sngY = sngY - lngH
    
    On Error GoTo ErrHandle
    Select Case bytIndex
    Case 1
        strsql = "SELECT a.id, a.编码, a.简码, a.名称 " _
                & "FROM 部门表 a " _
                & "Where TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'" _
                & "  and (a.编码 like [1] or a.简码 like [1])" _
                & "order by a.编码"
    Case 2
        strsql = "Select 病人ID id, 住院号, 姓名, 当前床号 From 病人信息 Where 出院时间 Is Null And (当前科室id = [1]) order by 姓名,当前床号,住院号"
    Case 3 '病人ID
        strsql = "select rownum ID,a.当前科室ID 科室ID,b.名称 科室名称,a.姓名,a.住院号,a.当前床号 床号 from 病人信息 a, 部门表 b " _
               & "where a.当前科室id=b.id and a.病人id=[1] order by a.姓名,a.当前床号,a.住院号"
    Case 4 '病人姓名
        strsql = "select rownum ID,a.当前科室ID 科室ID,b.名称 科室名称,a.姓名,a.住院号,a.当前床号 床号 from 病人信息 a, 部门表 b " _
               & "where a.当前科室id=b.id and a.姓名 like [1] order by a.姓名,a.当前床号,a.住院号"
    Case 5 '住院号
        strsql = "select rownum ID,a.当前科室ID 科室ID,b.名称 科室名称,a.姓名,a.住院号,a.当前床号 床号 from 病人信息 a, 部门表 b " _
               & "where a.当前科室id=b.id and a.住院号=[1] order by a.住院号,a.姓名,a.当前床号"
    Case 6 '门诊号
        strsql = "select rownum ID,a.当前科室ID 科室ID,b.名称 科室名称,a.姓名,a.住院号,a.当前床号 床号 from 病人信息 a, 部门表 b " _
               & "where a.当前科室id=b.id and a.门诊号=[1] order by a.门诊号,a.姓名,a.当前床号,a.住院号"
    Case 7 '床号
        strsql = "select rownum ID,a.当前科室ID 科室ID,b.名称 科室名称,a.姓名,a.住院号,a.当前床号 床号 from 病人信息 a, 部门表 b " _
               & "where a.当前科室id=b.id and a.当前床号 like [1] order by a.当前床号,a.姓名,a.住院号"
    End Select
    Set rstemp = zlDatabase.ShowSQLSelect(Me, strsql, 0, "选择器", False, "", "", False, False, True, sngX, sngY, lngH, blnCancel, False, False, strParam)
    
    If blnCancel = True Then
        GoTo gtEmpty
        Exit Sub
    End If
    
    If Not rstemp Is Nothing Then
        Select Case bytIndex
        Case 1
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("科室")) = zlStr.Nvl(rstemp!名称)
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("科室ID")) = zlStr.Nvl(rstemp!Id)
        Case 2
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("病人姓名")) = zlStr.Nvl(rstemp!姓名)
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("住院号")) = zlStr.Nvl(rstemp!住院号)
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("床号")) = zlStr.Nvl(rstemp!当前床号)
        Case Else
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("科室")) = IIf(IsNull(rstemp!科室名称), "", rstemp!科室名称)
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("科室ID")) = IIf(IsNull(rstemp!科室id), "", rstemp!科室id)
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("病人姓名")) = IIf(IsNull(rstemp!姓名), "", rstemp!姓名)
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("住院号")) = IIf(IsNull(rstemp!住院号), "", rstemp!住院号)
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("床号")) = IIf(IsNull(rstemp!床号), "", rstemp!床号)
        End Select
        rstemp.Close
    Else
        GoTo gtEmpty
    End If
    Exit Sub
    
gtEmpty:
    If bytIndex <> 2 Then
        vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("科室")) = ""
        vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("科室ID")) = ""
    End If
    vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("病人姓名")) = ""
    vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("住院号")) = ""
    vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("床号")) = ""
    Exit Sub
ErrHandle:
    MsgBox "录入数据有错！", vbCritical, gstrSysName
End Sub

Private Function GetCostlyInfoStr(ByVal intSN As Integer) As String
'高值材料字符串
    Dim strTmp As String
    If mrsCostlyInfo Is Nothing Then Exit Function
    With mrsCostlyInfo
        If .RecordCount <= 0 Then Exit Function
        .MoveFirst
        .Find "SN=" & intSN
        If Not .EOF Then
            strTmp = IIf(IsNull(mrsCostlyInfo!科室), "", Trim(mrsCostlyInfo!科室)) _
                   & "," & IIf(IsNull(mrsCostlyInfo!病人姓名), "", Trim(mrsCostlyInfo!病人姓名)) _
                   & "," & IIf(IsNull(mrsCostlyInfo!住院号), "", Trim(mrsCostlyInfo!住院号)) _
                   & "," & IIf(IsNull(mrsCostlyInfo!床号), "", Trim(mrsCostlyInfo!床号))
            If Replace(strTmp, ",", "") = "" Then Exit Function
            GetCostlyInfoStr = strTmp
        End If
    End With
End Function

Private Function Get中标单位成本价(ByVal lng物资ID As Long) As Double
    '----------------------------------------------------------------------------------------------
    '功能:获取中标单位的成本价
    '参数:物资id
    '返回:成功,返回成本价
    '编制:余智勇
    '日期:2010/11/19
    '问题:33718
    '----------------------------------------------------------------------------------------------
    Dim lng供应单位ID As Long
    Dim rstemp As New ADODB.Recordset
    lng供应单位ID = Val(txtProvider.Tag)
    err = 0: On Error GoTo ErrHand:
    gstrSQL = "Select 成本价 FROM 材料中标单位 where 材料id=[1] and 单位id=[2] "
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng物资ID, lng供应单位ID)
    If rstemp.EOF Then
        Get中标单位成本价 = 0
    Else
        Get中标单位成本价 = Val(zlStr.Nvl(rstemp!成本价))
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function


Private Sub SetSortRecord()
    Dim n As Integer
    
    If mshBill.Rows < 2 Then Exit Sub
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
        
        For n = 1 To mshBill.Rows - 1
            If mshBill.TextMatrix(n, 0) <> "" Then
                .AddNew
                !行号 = n
                !序号 = IIf(Val(mshBill.TextMatrix(n, mCol序号)) = 0, n, Val(mshBill.TextMatrix(n, mCol序号)))
                !药品id = Val(mshBill.TextMatrix(n, 0))
                !批次 = Val(mshBill.TextMatrix(n, mCol批次))
                
                .Update
            End If
        Next
        
    End With
End Sub

Private Function CheckRedo(ByVal rstemp As ADODB.Recordset) As ADODB.Recordset
    '功能：将重复的记录过滤掉，并返回过滤后的数据集合

    Dim i As Integer
    Dim strTemp As String
    Dim str批次 As String
    Dim str材料ID As String
    Dim str重复材料 As String
    Dim strDub As String
    Dim strsql As String
    
    rstemp.MoveFirst
    str批次 = ""
    Do While Not rstemp.EOF
        str批次 = IIf(IsNull(rstemp!批次), "0", rstemp!批次)
        If InStr(1, strTemp, rstemp!材料ID & "," & str批次) = 0 Then
            strTemp = strTemp & rstemp!材料ID & "," & str批次 & "|"
        End If
        rstemp.MoveNext
    Loop
    
    With mshBill
        For i = 1 To .Rows - 1
            If InStr(1, strTemp, .TextMatrix(i, 0) & "," & .TextMatrix(i, mCol批次)) > 0 And .TextMatrix(i, 0) <> "" Then
                str材料ID = str材料ID & .TextMatrix(i, 0) & "," & .TextMatrix(i, mCol诊疗) & "|"
            End If
        Next
        
        If str材料ID <> "" Then   '为过滤数据拼接sql
            strDub = ""
            For i = 0 To UBound(Split(str材料ID, "|")) - 1
                strDub = strDub & "材料id<>" & Split(Split(str材料ID, "|")(i), ",")(0) & " and "
                If UBound(Split(str重复材料, ",")) <= 2 Then
                    str重复材料 = str重复材料 & Split(Split(str材料ID, "|")(i), ",")(1) & ","
                End If
            Next
            If strDub <> "" Then
                strDub = Mid(strDub, 1, Len(strDub) - 4)
            End If
        End If
        
        If str重复材料 <> "" Then
            MsgBox str重复材料 & "列表中已经含有了！" & vbCrLf & "以上材料不再添加！", vbInformation, gstrSysName
            strsql = strDub
        End If
        rstemp.Filter = strsql
        Set CheckRedo = rstemp
    End With
End Function
