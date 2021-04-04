VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDesign 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "报表设计"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10635
   FillColor       =   &H80000012&
   Icon            =   "frmDesign.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MouseIcon       =   "frmDesign.frx":020A
   ScaleHeight     =   6870
   ScaleWidth      =   10635
   Begin VB.ComboBox CboTest 
      Height          =   300
      Left            =   -8888
      Style           =   2  'Dropdown List
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   1980
      Width           =   1695
   End
   Begin MSComctlLib.ImageList ImgTool 
      Left            =   9720
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":035C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":0A56
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1150
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":184A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1F44
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":263E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":2DB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":34B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":3BAC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picFormat 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   2340
      ScaleHeight     =   405
      ScaleWidth      =   6180
      TabIndex        =   34
      TabStop         =   0   'False
      ToolTipText     =   "X向标尺"
      Top             =   1410
      Width           =   6180
      Begin VB.CommandButton cmdDel 
         Height          =   375
         Left            =   5760
         Picture         =   "frmDesign.frx":49FE
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "删除报表格式"
         Top             =   15
         Width           =   405
      End
      Begin VB.CommandButton cmdAdd 
         Height          =   375
         Left            =   5325
         Picture         =   "frmDesign.frx":4D40
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "增加报表格式"
         Top             =   15
         Width           =   405
      End
      Begin MSComctlLib.ImageCombo cboFormat 
         Height          =   315
         Left            =   900
         TabIndex        =   0
         ToolTipText     =   "点击可以修改格式名称"
         Top             =   45
         Width           =   4410
         _ExtentX        =   7779
         _ExtentY        =   556
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         ImageList       =   "img16"
      End
      Begin VB.Label lblFormat 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "报表格式"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   105
         TabIndex        =   35
         Top             =   105
         Width           =   720
      End
   End
   Begin VB.PictureBox picSQL 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4980
      Left            =   0
      ScaleHeight     =   4980
      ScaleWidth      =   2280
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1530
      Width           =   2280
      Begin MSComctlLib.TreeView tvwSQL 
         Height          =   2085
         Left            =   90
         TabIndex        =   4
         Top             =   315
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   3678
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         PathSeparator   =   "."
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "img16"
         Appearance      =   1
      End
      Begin MSComctlLib.ListView lvwPar 
         DragIcon        =   "frmDesign.frx":5082
         Height          =   2325
         Left            =   90
         TabIndex        =   5
         Top             =   2775
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   4101
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "img16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "序号"
            Object.Width           =   961
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "名称"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "类型"
            Object.Width           =   961
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "缺省值"
            Object.Width           =   2646
         EndProperty
      End
      Begin MSComctlLib.ImageList img16 
         Left            =   1680
         Top             =   540
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   11
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDesign.frx":538C
               Key             =   "SQL_Custom"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDesign.frx":54E6
               Key             =   "SQL_Group"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDesign.frx":5640
               Key             =   "Root"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDesign.frx":579A
               Key             =   "Other"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDesign.frx":58F4
               Key             =   "String"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDesign.frx":5A4E
               Key             =   "Number"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDesign.frx":5BA8
               Key             =   "Date"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDesign.frx":5D02
               Key             =   "Bin"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDesign.frx":5E5C
               Key             =   "Format"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDesign.frx":5FB6
               Key             =   "Pars"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDesign.frx":C818
               Key             =   "ParsRoot"
            EndProperty
         EndProperty
      End
      Begin VB.Label lblPar 
         Alignment       =   2  'Center
         BackColor       =   &H009B6737&
         Caption         =   "数据参数"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         MousePointer    =   7  'Size N S
         TabIndex        =   26
         Top             =   2505
         Width           =   2040
      End
      Begin VB.Label lblSQL 
         Alignment       =   2  'Center
         BackColor       =   &H009B6737&
         Caption         =   "报表数据源"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   105
         TabIndex        =   25
         Top             =   60
         Width           =   2055
      End
   End
   Begin VB.PictureBox picL 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4980
      Left            =   2280
      MousePointer    =   9  'Size W E
      ScaleHeight     =   4980
      ScaleWidth      =   45
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1530
      Width           =   45
      Begin VB.Line Line2 
         BorderColor     =   &H80000015&
         X1              =   30
         X2              =   30
         Y1              =   0
         Y2              =   15360
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   15360
      End
   End
   Begin VB.PictureBox picR 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4980
      Left            =   8190
      MousePointer    =   9  'Size W E
      ScaleHeight     =   4980
      ScaleWidth      =   45
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1530
      Width           =   45
      Begin VB.Line Line4 
         BorderColor     =   &H80000015&
         X1              =   30
         X2              =   30
         Y1              =   -60
         Y2              =   15300
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   0
         Y1              =   -60
         Y2              =   15300
      End
   End
   Begin VB.PictureBox picAtt 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4980
      Left            =   8235
      ScaleHeight     =   4980
      ScaleWidth      =   2400
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1530
      Width           =   2400
      Begin VB.CommandButton cmdAtt 
         Caption         =   "…"
         Height          =   285
         Left            =   1725
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   2130
         Visible         =   0   'False
         Width           =   300
      End
      Begin MSComCtl2.DTPicker dtpAtt 
         Height          =   300
         Left            =   960
         TabIndex        =   48
         Top             =   2520
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "HH:mm:ss"
         Format          =   298254338
         CurrentDate     =   41766
      End
      Begin VB.ComboBox cboText 
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1080
         TabIndex        =   47
         TabStop         =   0   'False
         Text            =   "cboAttText"
         Top             =   2400
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.TextBox txtAtt 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   1155
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1890
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.PictureBox picM 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   45
         Left            =   150
         MousePointer    =   7  'Size N S
         ScaleHeight     =   45
         ScaleWidth      =   1740
         TabIndex        =   32
         Top             =   4125
         Width           =   1740
      End
      Begin VB.ComboBox cboAtt 
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1110
         Style           =   2  'Dropdown List
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   2475
         Width           =   960
      End
      Begin VSFlex8Ctl.VSFlexGrid mshAtt 
         Height          =   1095
         Left            =   120
         TabIndex        =   44
         Top             =   2400
         Width           =   1695
         _cx             =   1964641198
         _cy             =   1964640139
         Appearance      =   0
         BorderStyle     =   0
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
         BackColorSel    =   11103813
         ForeColorSel    =   16777215
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   0
         GridColorFixed  =   0
         TreeColor       =   -2147483644
         FloodColor      =   192
         SheetBorder     =   -2147483631
         FocusRect       =   1
         HighLight       =   0
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
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDesign.frx":1307A
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
      Begin MSComctlLib.Toolbar tbrTool 
         Height          =   720
         Left            =   90
         TabIndex        =   37
         Top             =   300
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         ToolTips        =   0   'False
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImgTool"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "选择"
               Key             =   "Point"
               Object.ToolTipText     =   "选择对象"
               Object.Tag             =   "选择"
               ImageIndex      =   1
               Style           =   2
               Value           =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "线条"
               Key             =   "Line"
               Object.ToolTipText     =   "线条"
               Object.Tag             =   "线条"
               ImageIndex      =   2
               Style           =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "框线"
               Key             =   "Frame"
               Object.ToolTipText     =   "框线"
               Object.Tag             =   "框线"
               ImageIndex      =   3
               Style           =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "标签"
               Key             =   "Note"
               Object.ToolTipText     =   "标签"
               Object.Tag             =   "标签"
               ImageIndex      =   4
               Style           =   2
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "图片"
               Key             =   "Picture"
               Object.ToolTipText     =   "图片"
               Object.Tag             =   "图片"
               ImageIndex      =   5
               Style           =   2
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "表格"
               Key             =   "Table"
               Object.ToolTipText     =   "表格"
               Object.Tag             =   "表格"
               ImageIndex      =   6
               Style           =   2
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "图表"
               Key             =   "Chart"
               Object.ToolTipText     =   "图表"
               Object.Tag             =   "图表"
               ImageIndex      =   7
               Style           =   2
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "条码"
               Key             =   "BarCode"
               Object.ToolTipText     =   "条码"
               Object.Tag             =   "条码"
               ImageIndex      =   8
               Style           =   2
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "卡片"
               Key             =   "Card"
               Object.ToolTipText     =   "卡片"
               Object.Tag             =   "卡片"
               ImageIndex      =   9
               Style           =   2
            EndProperty
         EndProperty
      End
      Begin VB.Label lblAtt 
         Alignment       =   2  'Center
         BackColor       =   &H009B6737&
         Caption         =   "元素属性"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   105
         TabIndex        =   28
         Top             =   1605
         Width           =   2040
      End
      Begin VB.Label lblNote 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFFF&
         Height          =   915
         Left            =   60
         TabIndex        =   29
         Top             =   4200
         Width           =   2250
      End
      Begin VB.Label lblTool 
         Alignment       =   2  'Center
         BackColor       =   &H009B6737&
         Caption         =   "基本元素"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   90
         TabIndex        =   27
         Top             =   105
         Width           =   2220
      End
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   2760
      MouseIcon       =   "frmDesign.frx":13157
      ScaleHeight     =   3855
      ScaleWidth      =   5355
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2265
      Width           =   5355
      Begin VB.PictureBox picPaperSize 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   60
         Index           =   2
         Left            =   5190
         MousePointer    =   8  'Size NW SE
         ScaleHeight     =   60
         ScaleWidth      =   60
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "拖动可改变纸张高度和宽度"
         Top             =   3690
         Width           =   60
      End
      Begin VB.PictureBox picPaperSize 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   60
         Index           =   1
         Left            =   105
         MousePointer    =   7  'Size N S
         ScaleHeight     =   60
         ScaleWidth      =   4935
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "拖动可改变纸张高度"
         Top             =   3705
         Width           =   4935
      End
      Begin VB.PictureBox picPaperSize 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3480
         Index           =   0
         Left            =   5190
         MousePointer    =   9  'Size W E
         ScaleHeight     =   3480
         ScaleWidth      =   60
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "拖动可改变纸张宽度"
         Top             =   105
         Width           =   60
      End
      Begin VB.PictureBox picPaper 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DrawStyle       =   2  'Dot
         ForeColor       =   &H00FF0000&
         Height          =   3525
         Left            =   105
         MouseIcon       =   "frmDesign.frx":132A9
         ScaleHeight     =   3525
         ScaleWidth      =   5025
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   120
         Width           =   5025
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   615
            Index           =   0
            Left            =   -8888
            ScaleHeight     =   585
            ScaleWidth      =   585
            TabIndex        =   42
            Top             =   1080
            Width           =   615
         End
         Begin C1Chart2D8.Chart2D Chart 
            Height          =   1440
            Index           =   0
            Left            =   -8888
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   1200
            Visible         =   0   'False
            Width           =   2100
            _Version        =   524288
            _Revision       =   7
            _ExtentX        =   3704
            _ExtentY        =   2540
            _StockProps     =   0
            ControlProperties=   "frmDesign.frx":133FB
         End
         Begin VB.PictureBox PicSplit 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1155
            Left            =   -8888
            ScaleHeight     =   1155
            ScaleWidth      =   15
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   150
            Visible         =   0   'False
            Width           =   15
            Begin VB.Line LineSplit 
               BorderStyle     =   3  'Dot
               X1              =   0
               X2              =   0
               Y1              =   0
               Y2              =   8000
            End
         End
         Begin VB.PictureBox PicFontTest 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   60
            Left            =   -8888
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.PictureBox LblSize 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   0
            Left            =   -8888
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   780
            Visible         =   0   'False
            Width           =   60
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh1 
            Height          =   585
            Index           =   0
            Left            =   -8888
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   2130
            _ExtentX        =   3757
            _ExtentY        =   1032
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            BackColorFixed  =   15724527
            ForeColorFixed  =   0
            ForeColorSel    =   16777215
            BackColorBkg    =   16777215
            BackColorUnpopulated=   16777215
            GridColor       =   0
            GridColorFixed  =   0
            GridColorUnpopulated=   16777215
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            FocusRect       =   0
            HighLight       =   0
            GridLinesFixed  =   1
            GridLinesUnpopulated=   1
            ScrollBars      =   0
            MergeCells      =   1
            AllowUserResizing=   1
            Appearance      =   0
            MousePointer    =   99
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "frmDesign.frx":13A5A
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VSFlex8Ctl.VSFlexGrid msh 
            Height          =   1575
            Index           =   0
            Left            =   360
            TabIndex        =   45
            Top             =   50000
            Visible         =   0   'False
            Width           =   3135
            _cx             =   1964643738
            _cy             =   1964640986
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
            MouseIcon       =   "frmDesign.frx":13D74
            BackColor       =   16777215
            ForeColor       =   0
            BackColorFixed  =   16777215
            ForeColorFixed  =   0
            BackColorSel    =   10251637
            ForeColorSel    =   16777215
            BackColorBkg    =   16777215
            BackColorAlternate=   16777215
            GridColor       =   0
            GridColorFixed  =   0
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   0
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   5
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   -1  'True
            ScrollBars      =   0
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
         Begin VB.Label lblshp 
            BackColor       =   &H8000000E&
            Height          =   735
            Index           =   0
            Left            =   -50000
            TabIndex        =   46
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Shape Shp 
            Height          =   1575
            Index           =   0
            Left            =   -50000
            Top             =   1200
            Width           =   2040
         End
         Begin VB.Image ImgCode 
            Appearance      =   0  'Flat
            Height          =   735
            Index           =   0
            Left            =   -8888
            MouseIcon       =   "frmDesign.frx":14C4E
            Stretch         =   -1  'True
            Top             =   1230
            Width           =   555
         End
         Begin VB.Image Img 
            Appearance      =   0  'Flat
            Height          =   735
            Index           =   0
            Left            =   -8888
            MouseIcon       =   "frmDesign.frx":14F58
            Stretch         =   -1  'True
            Top             =   390
            Width           =   555
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            BackColor       =   &H00EFEFEF&
            Caption         =   "标签"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   -2205
            MouseIcon       =   "frmDesign.frx":15262
            MousePointer    =   99  'Custom
            TabIndex        =   17
            Top             =   255
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Label lblLine 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            ForeColor       =   &H80000008&
            Height          =   15
            Index           =   0
            Left            =   -2235
            MouseIcon       =   "frmDesign.frx":1556C
            MousePointer    =   99  'Custom
            TabIndex        =   16
            Top             =   75
            Visible         =   0   'False
            Width           =   1410
         End
      End
   End
   Begin VB.VScrollBar scrVsc 
      DragIcon        =   "frmDesign.frx":15876
      Height          =   3870
      LargeChange     =   20
      Left            =   8205
      Max             =   100
      MouseIcon       =   "frmDesign.frx":15B80
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2280
      Width           =   250
   End
   Begin VB.HScrollBar scrHsc 
      DragIcon        =   "frmDesign.frx":15E8A
      Height          =   250
      LargeChange     =   20
      Left            =   2790
      Max             =   100
      MouseIcon       =   "frmDesign.frx":16194
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6165
      Width           =   5400
   End
   Begin VB.PictureBox picRulerH 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   350
      Left            =   2370
      ScaleHeight     =   345
      ScaleWidth      =   6105
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "X向标尺"
      Top             =   1875
      Width           =   6105
   End
   Begin VB.PictureBox picRulerV 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4140
      Left            =   2370
      ScaleHeight     =   4140
      ScaleWidth      =   345
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Y向标尺"
      Top             =   2265
      Width           =   350
   End
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   1530
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   2699
      BandCount       =   2
      _CBWidth        =   10635
      _CBHeight       =   1530
      _Version        =   "6.7.9782"
      BandForeColor1  =   255
      Caption1        =   "系统"
      Child1          =   "tbr1"
      MinHeight1      =   720
      Width1          =   1305
      Key1            =   "System"
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      BandForeColor2  =   16711680
      Caption2        =   "格式"
      Child2          =   "tbr2"
      MinHeight2      =   720
      Width2          =   915
      Key2            =   "Format"
      NewRow2         =   -1  'True
      Begin MSComctlLib.Toolbar tbr2 
         Height          =   720
         Left            =   585
         TabIndex        =   9
         Top             =   780
         Width           =   9960
         _ExtentX        =   17568
         _ExtentY        =   1270
         ButtonWidth     =   1138
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "img灰色"
         HotImageList    =   "img彩色"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   17
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "左对齐"
               Key             =   "Left"
               Description     =   "左对齐"
               Object.ToolTipText     =   "选中项目左对齐"
               Object.Tag             =   "左对齐"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "右对齐"
               Key             =   "Right"
               Description     =   "右对齐"
               Object.ToolTipText     =   "选中项目右对齐"
               Object.Tag             =   "右对齐"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "上对齐"
               Key             =   "Up"
               Description     =   "上对齐"
               Object.ToolTipText     =   "选中项目上对齐"
               Object.Tag             =   "上对齐"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "下对齐"
               Key             =   "Down"
               Description     =   "下对齐"
               Object.ToolTipText     =   "选中项目下对齐"
               Object.Tag             =   "下对齐"
               ImageIndex      =   14
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "横居中"
               Key             =   "Hsc"
               Description     =   "横居中"
               Object.ToolTipText     =   "选中项目横向居中"
               Object.Tag             =   "横居中"
               ImageIndex      =   15
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "竖居中"
               Key             =   "Vsc"
               Description     =   "竖居中"
               Object.ToolTipText     =   "选中项目竖向居中"
               Object.Tag             =   "竖居中"
               ImageIndex      =   16
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "同宽度"
               Key             =   "Width"
               Description     =   "同宽度"
               Object.ToolTipText     =   "选中项目宽度相同"
               Object.Tag             =   "同宽度"
               ImageIndex      =   17
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "同高度"
               Key             =   "Height"
               Description     =   "同高度"
               Object.ToolTipText     =   "选中项目高度相同"
               Object.Tag             =   "同高度"
               ImageIndex      =   18
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "同宽高"
               Key             =   "WH"
               Description     =   "同宽高"
               Object.ToolTipText     =   "选中项目宽高相同"
               Object.Tag             =   "同宽高"
               ImageIndex      =   19
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "竖间距"
               Key             =   "VscSpace"
               Description     =   "竖间距"
               Object.ToolTipText     =   "调整选中项目竖向间距"
               Object.Tag             =   "竖间距"
               ImageIndex      =   20
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "VscSame"
                     Object.Tag             =   "相同(&S)"
                     Text            =   "相同(&S)"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "VscAdd"
                     Object.Tag             =   "增加(&A)"
                     Text            =   "增加(&A)"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "VscDec"
                     Object.Tag             =   "减少(&D)"
                     Text            =   "减少(&D)"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "横间距"
               Key             =   "HscSpace"
               Description     =   "横间距"
               Object.ToolTipText     =   "调整选中项目横向间距"
               Object.Tag             =   "横间距"
               ImageIndex      =   21
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "HscSame"
                     Object.Tag             =   "相同(&S)"
                     Text            =   "相同(&S)"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "HscAdd"
                     Object.Tag             =   "增加(&A)"
                     Text            =   "增加(&A)"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "HscDec"
                     Object.Tag             =   "减少(&D)"
                     Text            =   "减少(&D)"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "比例"
               Key             =   "Scale"
               Description     =   "比例"
               Object.ToolTipText     =   "调整页面的显示比例"
               Object.Tag             =   "比例"
               ImageIndex      =   22
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   9
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Page"
                     Object.Tag             =   "整页显示(&P)"
                     Text            =   "整页显示(&P)"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Width"
                     Object.Tag             =   "适应宽度(&W)"
                     Text            =   "适应宽度(&W)"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Height"
                     Object.Tag             =   "适应高度(&H)"
                     Text            =   "适应高度(&H)"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "-Menu1"
                     Object.Tag             =   "-"
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Scale200"
                     Object.Tag             =   "200%"
                     Text            =   "200%"
                  EndProperty
                  BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Scale100"
                     Object.Tag             =   "100%"
                     Text            =   "100%"
                  EndProperty
                  BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Scale75"
                     Object.Tag             =   "75%"
                     Text            =   "75%"
                  EndProperty
                  BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Scale50"
                     Object.Tag             =   "50%"
                     Text            =   "50%"
                  EndProperty
                  BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Scale25"
                     Object.Tag             =   "25%"
                     Text            =   "25%"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "锁定"
               Key             =   "Lock"
               Description     =   "锁定"
               Object.ToolTipText     =   "锁定"
               Object.Tag             =   "锁定"
               ImageIndex      =   23
               Style           =   1
               Value           =   1
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbr1 
         Height          =   720
         Left            =   585
         TabIndex        =   8
         Top             =   30
         Width           =   9960
         _ExtentX        =   17568
         _ExtentY        =   1270
         ButtonWidth     =   1138
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "img灰色"
         HotImageList    =   "img彩色"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   21
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "执行"
               Key             =   "Report"
               Description     =   "执行"
               Object.ToolTipText     =   "执行报表"
               Object.Tag             =   "执行"
               ImageKey        =   "Report"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "保存"
               Key             =   "Save"
               Description     =   "保存"
               Object.ToolTipText     =   "保存报表"
               Object.Tag             =   "保存"
               ImageKey        =   "Save"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "向导"
               Key             =   "Guide"
               Description     =   "向导"
               Object.ToolTipText     =   "报表向导"
               Object.Tag             =   "向导"
               ImageKey        =   "Guide"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "页面"
               Key             =   "Page"
               Description     =   "页面"
               Object.ToolTipText     =   "页面设置"
               Object.Tag             =   "页面"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "新格式"
               Key             =   "AddFormat"
               Object.ToolTipText     =   "新增加一种报表格式"
               Object.Tag             =   "新格式"
               ImageKey        =   "AddFormat"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删格式"
               Key             =   "DelFormat"
               Object.ToolTipText     =   "删除当前报表格式"
               Object.Tag             =   "删除"
               ImageKey        =   "DelFormat"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "新数据"
               Key             =   "New"
               Description     =   "新数据"
               Object.ToolTipText     =   "增加新数据源"
               Object.Tag             =   "新数据"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "Modi"
               Description     =   "修改"
               Object.ToolTipText     =   "修改当前数据源"
               Object.Tag             =   "修改"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除源"
               Key             =   "Del"
               Description     =   "删除"
               Object.ToolTipText     =   "删除当前数据源"
               Object.Tag             =   "删除"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Data_"
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "新元素"
               Key             =   "Item"
               Description     =   "新元素"
               Object.ToolTipText     =   "增加报表元素"
               Object.Tag             =   "新元素"
               ImageIndex      =   6
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   8
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Line"
                     Object.Tag             =   "线条"
                     Text            =   "线条"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Frame"
                     Object.Tag             =   "框线"
                     Text            =   "框线"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "_1"
                     Object.Tag             =   "-"
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Label"
                     Object.Tag             =   "标签"
                     Text            =   "标签"
                  EndProperty
                  BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Picture"
                     Object.Tag             =   "图片"
                     Text            =   "图片"
                  EndProperty
                  BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Table"
                     Object.Tag             =   "表格"
                     Text            =   "表格"
                  EndProperty
                  BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Chart"
                     Object.Tag             =   "图表"
                     Text            =   "图表"
                  EndProperty
                  BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "BarCode"
                     Object.Tag             =   "条码"
                     Text            =   "条码"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删元素"
               Key             =   "Remove"
               Description     =   "删除"
               Object.ToolTipText     =   "删除票据项目"
               Object.Tag             =   "删除"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助"
               Object.ToolTipText     =   "帮助主题"
               Object.Tag             =   "帮助"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Quit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   10
            EndProperty
         EndProperty
         Begin MSComctlLib.Toolbar tb2 
            Height          =   720
            Left            =   6200
            TabIndex        =   43
            Top             =   0
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   1270
            ButtonWidth     =   1455
            ButtonHeight    =   1270
            AllowCustomize  =   0   'False
            Wrappable       =   0   'False
            Style           =   1
            ImageList       =   "img灰色"
            HotImageList    =   "img彩色"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   1
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "历史记录"
                  Key             =   "History"
                  Description     =   "历史记录"
                  Object.ToolTipText     =   "历史记录"
                  Object.Tag             =   "历史记录"
                  ImageKey        =   "Guide"
               EndProperty
            EndProperty
         End
      End
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
      ToolTipText     =   "重庆中联信息产业公司"
      Top             =   6510
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDesign.frx":162E6
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12912
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   953
            MinWidth        =   88
            Text            =   "位置"
            TextSave        =   "位置"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   35
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
            Object.ToolTipText     =   "当前数字键状态"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   35
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
   Begin MSComctlLib.ImageList imgLarge 
      Left            =   255
      Top             =   480
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
            Picture         =   "frmDesign.frx":16B7A
            Key             =   "Fields"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":16E94
            Key             =   "Field"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":171AE
            Key             =   "Bill"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":174C8
            Key             =   "Not"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgSmall 
      Left            =   885
      Top             =   375
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":18312
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":188AC
            Key             =   "Field"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":18E46
            Key             =   "Bill"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":18FA0
            Key             =   "Fields"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdg 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontName        =   "宋体"
      FontSize        =   9
      Min             =   9
   End
   Begin MSComctlLib.ImageList img彩色 
      Left            =   165
      Top             =   855
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   26
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1953A
            Key             =   "Page"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":19754
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1996E
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1A068
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1A762
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1AE5C
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1B556
            Key             =   "Remove"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1B770
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1B98A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1BBA4
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1BDBE
            Key             =   "Left"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1C4B8
            Key             =   "Right"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1CBB2
            Key             =   "Up"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1D2AC
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1D9A6
            Key             =   "Hsc"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1E0A0
            Key             =   "Vsc"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1E79A
            Key             =   "Width"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1EE94
            Key             =   "Height"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1F58E
            Key             =   "WH"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1FC88
            Key             =   "VscSpace"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":20382
            Key             =   "HscSpace"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":20A7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":20C96
            Key             =   "Lock"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":21390
            Key             =   "Guide"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":215AA
            Key             =   "AddFormat"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":217C4
            Key             =   "DelFormat"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img灰色 
      Left            =   810
      Top             =   885
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   26
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":219DE
            Key             =   "Page"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":21BF8
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":21E12
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":2250C
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":22C06
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":23300
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":239FA
            Key             =   "Remove"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":23C14
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":23E2E
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":24048
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":24262
            Key             =   "Left"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":2495C
            Key             =   "Right"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":25056
            Key             =   "Up"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":25750
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":25E4A
            Key             =   "Hsc"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":26544
            Key             =   "Vsc"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":26C3E
            Key             =   "Width"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":27338
            Key             =   "Height"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":27A32
            Key             =   "WH"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":2812C
            Key             =   "VscSpace"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":28826
            Key             =   "HscSpace"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":28F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":2913A
            Key             =   "Lock"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":29834
            Key             =   "Guide"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":29A4E
            Key             =   "AddFormat"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":29C68
            Key             =   "DelFormat"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFile_Report 
         Caption         =   "执行报表(&E)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFile_Save 
         Caption         =   "保存报表(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFile_Guide 
         Caption         =   "报表向导(&G)"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuEdit_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Page 
         Caption         =   "页面设置(&S)"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Quit 
         Caption         =   "退出(&X)"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEdit_Copy 
         Caption         =   "复制元素(&C)"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEdit_Paste 
         Caption         =   "粘贴元素(&P)"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuEdit_7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_SelAll 
         Caption         =   "全部选择(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEdit_Inverse 
         Caption         =   "反向选择(&I)"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuEdit_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_AddFormat 
         Caption         =   "增加报表格式(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEdit_DelFormat 
         Caption         =   "删除报表格式(&E)"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuEdit_5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_New 
         Caption         =   "增加数据源(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuEdit_Modi 
         Caption         =   "修改数据源(&M)"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuEdit_Del 
         Caption         =   "删除数据源(&D)"
      End
      Begin VB.Menu mnuEdit_History 
         Caption         =   "历史记录(&H)"
      End
      Begin VB.Menu mnuEdit_Data_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Item 
         Caption         =   "增加元素(&T)"
         Begin VB.Menu mnuEdit_ItemAdd 
            Caption         =   "线条(&L)"
            Index           =   0
            Shortcut        =   ^L
         End
         Begin VB.Menu mnuEdit_ItemAdd 
            Caption         =   "框线(&F)"
            Index           =   1
            Shortcut        =   ^T
         End
         Begin VB.Menu mnuEdit_ItemAdd 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnuEdit_ItemAdd 
            Caption         =   "标签(&D)"
            Index           =   3
            Shortcut        =   ^D
         End
         Begin VB.Menu mnuEdit_ItemAdd 
            Caption         =   "图片(&P)"
            Index           =   4
            Shortcut        =   ^P
         End
         Begin VB.Menu mnuEdit_ItemAdd 
            Caption         =   "表格(&B)"
            Index           =   5
            Shortcut        =   ^B
         End
         Begin VB.Menu mnuEdit_ItemAdd 
            Caption         =   "图表(&C)"
            Index           =   6
            Shortcut        =   ^H
         End
         Begin VB.Menu mnuEdit_ItemAdd 
            Caption         =   "条码(&R)"
            Index           =   7
            Shortcut        =   ^R
         End
      End
      Begin VB.Menu mnuEdit_Remove 
         Caption         =   "删除元素(&R)"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "格式(&R)"
      Begin VB.Menu mnuFormat_Order 
         Caption         =   "设计顺序(&O)"
         Begin VB.Menu mnuFormat_Front 
            Caption         =   "置前(&F)"
         End
         Begin VB.Menu mnuFormat_Back 
            Caption         =   "置后(&B)"
         End
      End
      Begin VB.Menu mnuFormat_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormat_Align 
         Caption         =   "对齐(&A)"
         Begin VB.Menu mnuFormat_DoAlign 
            Caption         =   "左对齐(&L)"
            Index           =   0
         End
         Begin VB.Menu mnuFormat_DoAlign 
            Caption         =   "右对齐(&R)"
            Index           =   1
         End
         Begin VB.Menu mnuFormat_DoAlign 
            Caption         =   "上对齐(&U)"
            Index           =   2
         End
         Begin VB.Menu mnuFormat_DoAlign 
            Caption         =   "下对齐(&D)"
            Index           =   3
         End
         Begin VB.Menu mnuFormat_DoAlign 
            Caption         =   "横居中(&C)"
            Index           =   4
         End
         Begin VB.Menu mnuFormat_DoAlign 
            Caption         =   "竖居中(&M)"
            Index           =   5
         End
         Begin VB.Menu mnuFormat_DoAlign 
            Caption         =   "-"
            Index           =   6
         End
         Begin VB.Menu mnuFormat_DoAlign 
            Caption         =   "水平居中(&H)"
            Index           =   7
         End
         Begin VB.Menu mnuFormat_DoAlign 
            Caption         =   "垂直居中(&V)"
            Index           =   8
         End
      End
      Begin VB.Menu mnuFormat_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormat_Size 
         Caption         =   "尺寸(&S)"
         Begin VB.Menu mnuFormat_Width 
            Caption         =   "同宽度(&W)"
         End
         Begin VB.Menu mnuFormat_Height 
            Caption         =   "同高度(&H)"
         End
         Begin VB.Menu mnuFormat_WH 
            Caption         =   "同宽高(&B)"
         End
      End
      Begin VB.Menu mnuFomrat_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormat_VscSapce 
         Caption         =   "竖间距(&V)"
         Begin VB.Menu mnuFormat_VscSpace_Same 
            Caption         =   "相同(&S)"
         End
         Begin VB.Menu mnuFormat_VscSpace_Add 
            Caption         =   "增加(&A)"
         End
         Begin VB.Menu mnuFormat_VscSpace_Dec 
            Caption         =   "减少(&D)"
         End
      End
      Begin VB.Menu mnuFormat_HscSpace 
         Caption         =   "横间距(&H)"
         Begin VB.Menu mnuFormat_HscSpace_Same 
            Caption         =   "相同(&S)"
         End
         Begin VB.Menu mnuFormat_HscSpace_Add 
            Caption         =   "增加(&A)"
         End
         Begin VB.Menu mnuFormat_HscSpace_Dec 
            Caption         =   "减少(&D)"
         End
      End
      Begin VB.Menu mnuFormat_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormat_Lock 
         Caption         =   "锁定元素(&L)"
         Checked         =   -1  'True
         Shortcut        =   ^K
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "视图(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu mnuViewToolSystem 
            Caption         =   "系统功能(&B)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolFormat 
            Caption         =   "格式调整(&F)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuView_1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "文本标签(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewScale 
         Caption         =   "显示比例(&C)"
         Begin VB.Menu mnuViewScaleMode 
            Caption         =   "整页显示(&P)"
            Index           =   0
         End
         Begin VB.Menu mnuViewScaleMode 
            Caption         =   "适应宽度(&W)"
            Index           =   1
         End
         Begin VB.Menu mnuViewScaleMode 
            Caption         =   "适应高度(&H)"
            Index           =   2
         End
         Begin VB.Menu mnuViewScaleMode 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnuViewScaleMode 
            Caption         =   "200%"
            Checked         =   -1  'True
            Index           =   4
         End
         Begin VB.Menu mnuViewScaleMode 
            Caption         =   "100%"
            Checked         =   -1  'True
            Index           =   5
         End
         Begin VB.Menu mnuViewScaleMode 
            Caption         =   "75%"
            Checked         =   -1  'True
            Index           =   6
         End
         Begin VB.Menu mnuViewScaleMode 
            Caption         =   "50%"
            Checked         =   -1  'True
            Index           =   7
         End
         Begin VB.Menu mnuViewScaleMode 
            Caption         =   "25%"
            Checked         =   -1  'True
            Index           =   8
         End
      End
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewToolAttrib 
         Caption         =   "属性表框(&A)"
         Checked         =   -1  'True
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuViewToolSQL 
         Caption         =   "数据源框(&L)"
         Checked         =   -1  'True
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuViewToolRuler 
         Caption         =   "报表标尺(&U)"
         Checked         =   -1  'True
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuView_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView_reFlash 
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
         Caption         =   "&WEB上的中联"
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
      Begin VB.Menu mnuHelp_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
   Begin VB.Menu mnuCustom 
      Caption         =   "任意表格"
      Visible         =   0   'False
      Begin VB.Menu mnuCustom_Head 
         Caption         =   "表头操作"
         Begin VB.Menu mnuCustom_Head_Insert 
            Caption         =   "插入表头行(&I)"
            Begin VB.Menu mnuCustom_Head_Insert_UP 
               Caption         =   "当前行上面(&U)"
            End
            Begin VB.Menu mnuCustom_Head_Insert_Down 
               Caption         =   "当前行下面(&D)"
            End
         End
         Begin VB.Menu mnuCustom_Head_Del 
            Caption         =   "删除表头行(&D)"
         End
         Begin VB.Menu mnuCustom_Head_Auto 
            Caption         =   "自动编列号(&N)"
         End
         Begin VB.Menu mnuCustom_Head_Clear 
            Caption         =   "清空表头(&C)"
         End
         Begin VB.Menu mnuCustom_Head_1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCustom_Head_Text 
            Caption         =   "单元格文字(&T)"
         End
         Begin VB.Menu mnuCustom_Head_Merge 
            Caption         =   "单元格合并(&M)"
         End
         Begin VB.Menu mnuCustom_Head_Split 
            Caption         =   "单元格拆分(&S)"
         End
      End
      Begin VB.Menu mnuCustom_Col 
         Caption         =   "表列操作"
         Begin VB.Menu mnuCustom_Col_Insert 
            Caption         =   "插入表列(&I)"
            Begin VB.Menu mnuCustom_Col_Insert_Left 
               Caption         =   "当前列左面(&L)"
            End
            Begin VB.Menu mnuCustom_Col_Insert_Right 
               Caption         =   "当前列右面(&R)"
            End
         End
         Begin VB.Menu mnuCustom_Col_Del 
            Caption         =   "删除表列(&R)"
         End
         Begin VB.Menu mnuCustom_Col_Clear 
            Caption         =   "清空表体(&C)"
         End
         Begin VB.Menu mnuCustom_Col_1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCustom_Col_Data 
            Caption         =   "表列数据(&D)"
         End
         Begin VB.Menu mnuCustom_Col_State 
            Caption         =   "表列汇总(&S)"
            Begin VB.Menu mnuCustom_Col_State_Style 
               Caption         =   "无(&0)"
               Checked         =   -1  'True
               Index           =   0
            End
            Begin VB.Menu mnuCustom_Col_State_Style 
               Caption         =   "求和(&1)"
               Index           =   1
            End
            Begin VB.Menu mnuCustom_Col_State_Style 
               Caption         =   "求平均值(&2)"
               Index           =   2
            End
            Begin VB.Menu mnuCustom_Col_State_Style 
               Caption         =   "求最大值(&3)"
               Index           =   3
            End
            Begin VB.Menu mnuCustom_Col_State_Style 
               Caption         =   "求最小值(&4)"
               Index           =   4
            End
            Begin VB.Menu mnuCustom_Col_State_Style 
               Caption         =   "求记录数(&5)"
               Index           =   5
            End
         End
         Begin VB.Menu mnuCustom_Col_Align 
            Caption         =   "表列对齐(&A)"
            Begin VB.Menu mnuCustom_Col_Align_Style 
               Caption         =   "左对齐(&L)"
               Checked         =   -1  'True
               Index           =   0
            End
            Begin VB.Menu mnuCustom_Col_Align_Style 
               Caption         =   "居中对齐(&M)"
               Index           =   1
            End
            Begin VB.Menu mnuCustom_Col_Align_Style 
               Caption         =   "右对齐(&R)"
               Index           =   2
            End
         End
      End
   End
   Begin VB.Menu mnuClass 
      Caption         =   "分类表格"
      Visible         =   0   'False
      Begin VB.Menu mnuClass_Insert 
         Caption         =   "插入新项目(&I)"
         Begin VB.Menu mnuClass_Insert_Before 
            Caption         =   "在当前项之前(&B)"
         End
         Begin VB.Menu mnuClass_Insert_After 
            Caption         =   "在当前项之后(&A)"
         End
      End
      Begin VB.Menu mnuClass_Data 
         Caption         =   "表项数据(&D)"
      End
      Begin VB.Menu mnuClass_ExChange 
         Caption         =   "行列对换(&E)"
      End
      Begin VB.Menu mnuClass_Del 
         Caption         =   "删除表项(&R)"
      End
      Begin VB.Menu mnuClass_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClass_State 
         Caption         =   "表项汇总(&S)"
         Begin VB.Menu mnuClass_State_Style 
            Caption         =   "无(&0)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuClass_State_Style 
            Caption         =   "求和(&1)"
            Index           =   1
         End
         Begin VB.Menu mnuClass_State_Style 
            Caption         =   "求平均值(&2)"
            Index           =   2
         End
         Begin VB.Menu mnuClass_State_Style 
            Caption         =   "求最大值(&3)"
            Index           =   3
         End
         Begin VB.Menu mnuClass_State_Style 
            Caption         =   "求最小值(&4)"
            Index           =   4
         End
         Begin VB.Menu mnuClass_State_Style 
            Caption         =   "求记录数(&5)"
            Index           =   5
         End
      End
      Begin VB.Menu mnuClass_Align 
         Caption         =   "表项对齐(&A)"
         Begin VB.Menu mnuClass_Align_Style 
            Caption         =   "左对齐(&L)"
            Index           =   0
         End
         Begin VB.Menu mnuClass_Align_Style 
            Caption         =   "中间对齐(&M)"
            Index           =   1
         End
         Begin VB.Menu mnuClass_Align_Style 
            Caption         =   "右对齐(&R)"
            Checked         =   -1  'True
            Index           =   2
         End
      End
   End
End
Attribute VB_Name = "frmDesign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public lngRPTID As Long '入：要设计的报表ID
Public mblnNotModiData As Boolean '入：是否不可以修改数据源

Private objReport As Report '要设计的报表对象
Private blnLock As Boolean
Private blnMax As Boolean
Private bytLine As Byte
Private blnDown As Boolean
Private lngPreX As Long, lngPreY As Long
Private preHsc As Long, preVsc As Long '标尺位置记录
Private bytCurTool As Byte '当前选中元素(0=选择;1=线条;2=框线;3=标签;4=图片;5=表格;6=图表;7=条码,8=卡片)
Private selArea As RECT '存放选择的矩形框
Private selCell As Cells '存放选择的表头范围
Private drgCell As Cells '拖动过程中的单元格范围,drgCell.Row无用
Private intCurCol As Integer '当前选择任意表格列
Private objFont As New clsRotateFont '旋转字体对象
Private intMaxID As Integer '当前最大控件索引(从1开始)
Private intCurID As Integer '当前选择控件索引(从1开始)
Private BlnSave As Boolean
Private objLastSel As Object '最后一个选中的元素控件
Private blnDrop As Boolean, blnHead As Boolean, blnSum As Boolean
Private strMenu As String
Private mblnFirst As Boolean
Private mobjMove As Object  '元素移入的父控件
Private mlngX As Long   '移入父控件的位置
Private mlngY As Long   '移入父控件的位置
Private mobjPicMERGE As IPictureDisp
Private mobjPicMove As IPictureDisp

'zyb#Add
Private blnDelReportFormat As Boolean   '报表是否是固定报表(固定报表不允许删除报表样式)
Private mbytCurrFmt As Byte            '选择的报表样式(用于修改报表样式的名称)
Private blnAllowIn As Boolean           '是否进入Change事件
Private blnRefresh As Boolean           '必须刷新
Private blnModify As Boolean            '是否允许更新
Private blnAdjustRowHeight As Boolean   '允许改变固定行的行高
Private blnAdjustColWidth As Boolean    '允许改变所有列的列宽
Private sgnMode As Single
Private sgnLastMode As Single
Private Type WindowProperty
        l As Single
        H As Single
        T As Single
        W As Single
End Type
Private WinProperty As WindowProperty

Private Sub cboAtt_Click()
    Dim ItemThis As RPTItem, ItemSend As RPTItem, ItemFmt As RPTFmt
    Dim str性质 As String, intID As Integer
    Dim strCurText As String, intType As Integer
    Dim objBarCode As StdPicture, lngSize As Long
    Dim strBarCode As String, sngWidth As Single
    Dim blnSeek As Boolean
    Dim k As Long, X As Long, Y As Long, tmpID As RelatID, ItemTmp As RPTItem
    Dim StrCompare As String, lngX As Long, lngY As Long
    Dim tmpItem As RPTItem, strSouse As String
    Dim j As Long, i As Long, blnYes As Boolean
    Dim tmpObj As PictureBox

    strCurText = mshAtt.TextMatrix(mshAtt.Row, 0)
    If intCurID = 0 And InStr(1, "输出图形,报表元素", strCurText) = 0 Then Exit Sub '2002-03-26
    
    If intCurID <> 0 Then
        intType = objReport.Items("_" & intCurID).类型
    End If
    Select Case strCurText
        Case "对齐"
            mshAtt.TextMatrix(mshAtt.Row, 1) = cboAtt.Text
            '如果被选中的所有元素都是相同的元素
            If lblSize.count > 9 Then
                For Each tmpObj In lblSize
                    If tmpObj.Index Mod 8 = 1 Then
                        lbl(tmpObj.Tag).Alignment = IIF(cboAtt.ListIndex <> 0, IIF(cboAtt.ListIndex = 1, 2, 1), 0)
                        objReport.Items("_" & tmpObj.Tag).对齐 = cboAtt.ListIndex
                    End If
                Next
            Else
                lbl(intCurID).Alignment = IIF(cboAtt.ListIndex <> 0, IIF(cboAtt.ListIndex = 1, 2, 1), 0)
                objReport.Items("_" & intCurID).对齐 = cboAtt.ListIndex
            End If
            BlnSave = False
        Case "报表元素"
            If cboAtt.Text = "" Then Exit Sub
            'Call SelClear
            Call SelItem(cboAtt.ItemData(cboAtt.ListIndex), True)
            Call ShowAttrib(cboAtt.ItemData(cboAtt.ListIndex))
            BlnSave = False
        Case "输出图形"
            '更改集合
            If blnModify = False Then Exit Sub
            For Each ItemFmt In objReport.Fmts
                If ItemFmt.序号 = mbytCurrFmt Then
                    ItemFmt.图样 = cboAtt.ItemData(cboAtt.ListIndex)
                    Exit For
                End If
            Next
            mshAtt.TextMatrix(mshAtt.Row, 1) = cboAtt.Text
            BlnSave = False
        Case "参照对象"
            If objReport.Items("_" & intCurID).参照 = cboAtt.Text Then Exit Sub
            If (objReport.Items("_" & intCurID).类型 = 4 Or objReport.Items("_" & intCurID).类型 = 5) And cboAtt.Text <> "" Then
                If objReport.Items("_" & GetDependID(cboAtt.Text)).父ID <> 0 Then
                    MsgBox "卡片内的表格不允许设置附加表格！", vbInformation, App.Title
                    cboAtt.ListIndex = -1: cboAtt.SetFocus: Exit Sub
                End If
            End If
            Call CopyItem(ItemSend, objReport.Items("_" & intCurID))
            mshAtt.TextMatrix(mshAtt.Row, 1) = cboAtt.Text
            objReport.Items("_" & intCurID).参照 = cboAtt.Text
            objReport.Items("_" & intCurID).性质 = 0
            
            '如果参照对象无,则其附属项目[方向,性质]为"独立"
            If cboAtt.Text = "" Then
                With mshAtt
                    .TextMatrix(GetRow("性质"), 1) = "独立"
                    objReport.Items("_" & intCurID).性质 = "0"
                End With
            Else
                With mshAtt
                    .TextMatrix(GetRow("参照对象"), 1) = objReport.Items("_" & intCurID).参照
                    Select Case objReport.Items("_" & intCurID).类型
                    Case 2
                        str性质 = IIF(objReport.Items("_" & intCurID).性质 = "0", "11", objReport.Items("_" & intCurID).性质)
                        objReport.Items("_" & intCurID).性质 = str性质
                        .TextMatrix(GetRow("方向"), 1) = IIF(Mid(str性质, 1, 1) <> "2", "表上项", "表下项")
                        str性质 = IIF(str性质 = "0" Or str性质 = "", "独立", IIF(Mid(str性质, 2) = "1", "靠左", IIF(Mid(str性质, 2) = "2", "靠中", "靠右")))
                        .TextMatrix(GetRow("性质"), 1) = str性质
                    Case 4, 5
                        str性质 = ""
                        For Each ItemThis In objReport.Items
                            If ItemThis.参照 <> "" And ItemThis.格式号 = mbytCurrFmt And ItemThis.参照 = mshAtt.TextMatrix(GetRow("参照对象"), 1) And InStr(1, "4,5", ItemThis.类型) <> 0 Then
                                If ItemThis.性质 = 0 Then
                                    str性质 = IIF(ItemThis.类型 = 4, "附加", "左联接")
                                Else
                                    str性质 = IIF(ItemThis.性质 = 1, "附加", "左联接")
                                End If
                                Exit For
                            End If
                        Next
                        
                        If str性质 = "" Then str性质 = IIF(objReport.Items("_" & intCurID).类型 = 4, "附加", "左联接")
                        .TextMatrix(GetRow("性质"), 1) = str性质
                        objReport.Items("_" & intCurID).性质 = IIF(str性质 = "附加", 1, 2)
                    End Select
                End With
                If objReport.Items("_" & GetDependID(cboAtt.Text)).父ID <> 0 Then
                    If objReport.Items("_" & intCurID).父ID = 0 Then
                        objReport.Items("_" & intCurID).父ID = objReport.Items("_" & GetDependID(cboAtt.Text)).父ID
                    End If
                Else
                    If objReport.Items("_" & intCurID).父ID <> 0 Then
                        objReport.Items("_" & intCurID).父ID = 0
                    End If
                End If
                Dim ParentItem As RPTItem
                For Each ParentItem In objReport.Items
                    If ParentItem.格式号 = mbytCurrFmt And ParentItem.名称 = objReport.Items("_" & intCurID).参照 And ParentItem.类型 = 5 And objReport.Items("_" & intCurID).类型 = 5 Then
                        Call SetGridLike(msh(ParentItem.Key), msh(intCurID))
                        Exit For
                    End If
                Next
            End If
            
            Call ReferTo(ItemSend)
            If objReport.Items("_" & intCurID).系统 Then Call AdjustCoordinate
            
            BlnSave = False
            
            If GetSelNum = 1 Then Call ShowAttrib(intCurID)
        Case "数据源"
            If objReport.Items("_" & intCurID).数据源 = cboAtt.Text Then Exit Sub
            If Trim(cboAtt.Text) <> "" Then
                For Each ItemTmp In objReport.Items
                    If ItemTmp.父ID <> 0 And ItemTmp.父ID = intCurID Then
                        If ItemTmp.类型 = 4 Then
                            For Each tmpID In ItemTmp.SubIDs
                                With objReport.Items("_" & tmpID.id)
                                    X = InStr(1, .内容, "]")
                                    Y = InStr(1, .内容, ".")
                                    k = InStr(1, .内容, "[")
                                    If X > k And X > Y And X <> 0 And k <> 0 Then
                                        If Mid(.内容, k + 1, Y - k - 1) <> cboAtt.Text Then
                                            MsgBox "卡片中的表格绑定的数据列必须属于选择数据源，请检查！", vbInformation, App.Title
                                            Call CboSetText(cboAtt, objReport.Items("_" & intCurID).数据源)
                                            Exit Sub
                                        End If
                                    End If
                                End With
                            Next
                        End If
                    End If
                Next
            End If

            
            mshAtt.TextMatrix(mshAtt.Row, 1) = cboAtt.Text
            objReport.Items("_" & intCurID).数据源 = cboAtt.Text
            
            BlnSave = False
            
            If GetSelNum = 1 Then Call ShowAttrib(intCurID)
        Case "容器"
            If objReport.Items("_" & intCurID).父ID = 0 Then
                StrCompare = "页面"
            Else
                StrCompare = objReport.Items("_" & objReport.Items("_" & intCurID).父ID).名称
            End If
            If StrCompare = cboAtt.Text Then Exit Sub
            '检查是否是分栏表格
            If objReport.Items("_" & intCurID).类型 = 4 And cboAtt.Text <> "页面" Then
                If objReport.Items("_" & intCurID).分栏 > 1 Then
                    MsgBox "卡片中不允许放入分栏的表格。", vbInformation, App.Title
                    Exit Sub
                End If
                '卡片内不允许附加表格
                For Each tmpItem In objReport.Items
                    If tmpItem.格式号 = mbytCurrFmt Then
                        If tmpItem.类型 = 5 Or tmpItem.类型 = 4 Then
                            If tmpItem.参照 = objReport.Items("_" & intCurID).名称 Then
                                 MsgBox "本表存在附加表格，不允许放入卡片中！", vbInformation, App.Title
                                 Exit Sub
                            End If
                        End If
                    End If
                Next
                '如果卡片有数据源，则检查表格的数据源是否匹配
                If objReport.Items("_" & Val(cboAtt.ItemData(cboAtt.ListIndex) & "")).数据源 <> "" Then
                    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                        With objReport.Items("_" & tmpID.id)
                            i = InStr(1, .内容, "]")
                            j = InStr(1, .内容, ".")
                            k = InStr(1, .内容, "[")
                            If i > k And i > j And i <> 0 And k <> 0 And j <> 0 Then
                                If Mid(.内容, k + 1, j - k - 1) <> objReport.Items("_" & Val(cboAtt.ItemData(cboAtt.ListIndex) & "")).数据源 Then
                                    If blnYes = False Then
                                        If MsgBox("当前卡片绑定了数据源，而表格中的数据列和卡片数据源不相同，移入将清空不匹配的列，是否继续?", vbQuestion + vbYesNo + vbDefaultButton1, App.Title) = vbYes Then
                                            .内容 = ""
                                            msh(intCurID).TextMatrix(1, .序号) = ""
                                            blnYes = True
                                        Else
                                            Exit Sub
                                        End If
                                    Else
                                        .内容 = ""
                                        msh(intCurID).TextMatrix(1, .序号) = ""
                                    End If
                                End If
                            End If
                        End With
                    Next
                Else
                    '卡片没有数据源，则提示用户是否添加数据源
                    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                        With objReport.Items("_" & tmpID.id)
                            i = InStr(1, .内容, "]")
                            j = InStr(1, .内容, ".")
                            k = InStr(1, .内容, "[")
                            If i > k And i > j And i <> 0 And k <> 0 And j <> 0 Then
                                If InStr(strSouse, Mid(.内容, k + 1, j - k - 1)) = 0 Then
                                    strSouse = strSouse & "," & Mid(.内容, k + 1, j - k - 1)
                                End If
                            End If
                        End With
                    Next
                    strSouse = Mid(strSouse, 2)
                    '只有一个数据源时才提示
                    If InStr(strSouse, ",") = 0 And strSouse <> "" Then
                        If MsgBox("当前卡片未绑定数据源，绑定后将分组打印多张卡片，数据源中存在""分组标识""字段则""分组标识""相同的为一组,否则一行数据为一组；" & vbCrLf & _
                             "不绑定则只打印一张卡片，是否绑定数据源""" & strSouse & """?", vbQuestion + vbYesNo + vbDefaultButton1, App.Title) = vbYes Then
                            objReport.Items("_" & mobjMove.Index).数据源 = strSouse
                        End If
                    End If
                End If
            End If
            mshAtt.TextMatrix(mshAtt.Row, 1) = cboAtt.Text
            If cboAtt.Text = "页面" Then
                objReport.Items("_" & intCurID).X = objReport.Items("_" & intCurID).X + objReport.Items("_" & objReport.Items("_" & intCurID).父ID).X
                objReport.Items("_" & intCurID).Y = objReport.Items("_" & intCurID).Y + objReport.Items("_" & objReport.Items("_" & intCurID).父ID).Y
                objReport.Items("_" & intCurID).父ID = 0
            Else
                If objReport.Items("_" & intCurID).父ID = 0 Then
                    lngX = objReport.Items("_" & Val(cboAtt.ItemData(cboAtt.ListIndex) & "")).X
                    lngY = objReport.Items("_" & Val(cboAtt.ItemData(cboAtt.ListIndex) & "")).Y
                Else
                    lngX = objReport.Items("_" & Val(cboAtt.ItemData(cboAtt.ListIndex) & "")).X - objReport.Items("_" & intCurID).X
                    lngY = objReport.Items("_" & Val(cboAtt.ItemData(cboAtt.ListIndex) & "")).Y - objReport.Items("_" & intCurID).Y
                End If
                objReport.Items("_" & intCurID).X = objReport.Items("_" & intCurID).X - lngX
                objReport.Items("_" & intCurID).Y = objReport.Items("_" & intCurID).Y - lngY
                objReport.Items("_" & intCurID).父ID = Val(cboAtt.ItemData(cboAtt.ListIndex) & "")
            End If
            If objReport.Items("_" & intCurID).类型 = 4 Then
                '处理子项
                For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                    objReport.Items("_" & tmpID.id).父ID = objReport.Items("_" & intCurID).父ID
                Next
            End If
            Call AdjustCoordinate(True)
            BlnSave = False
            If GetSelNum = 1 Then Call ShowAttrib(intCurID)
        Case "方向"
            If mshAtt.TextMatrix(mshAtt.Row, 1) = cboAtt.Text Then Exit Sub
            objReport.Items("_" & intCurID).性质 = IIF(cboAtt.Text = "表上项", "1", "2") & Mid(objReport.Items("_" & intCurID).性质, 2)
            mshAtt.TextMatrix(mshAtt.Row, 1) = cboAtt.Text
            
            Call CopyItem(ItemSend, objReport.Items("_" & intCurID))
            Call ReferTo(ItemSend)
            If objReport.Items("_" & intCurID).系统 Then Call AdjustCoordinate
            
            BlnSave = False
            
            If GetSelNum = 1 Then Call ShowAttrib(intCurID)
        Case "性质"
            If mshAtt.TextMatrix(mshAtt.Row, 1) = cboAtt.Text Then Exit Sub
            Call CopyItem(ItemSend, objReport.Items("_" & intCurID))
            Select Case objReport.Items("_" & intCurID).类型
            Case 2
                mshAtt.TextMatrix(mshAtt.Row, 1) = cboAtt.Text
                str性质 = IIF(cboAtt.Text = "靠左", "1", IIF(cboAtt.Text = "靠中", "2", "3"))
                objReport.Items("_" & intCurID).性质 = Mid(objReport.Items("_" & intCurID).性质, 1, 1) & str性质
            Case 5
                mshAtt.TextMatrix(mshAtt.Row, 1) = cboAtt.Text
                objReport.Items("_" & intCurID).性质 = IIF(cboAtt.Text = "附加", "1", "2")
            End Select
            
            Call ReferTo(ItemSend)
            If objReport.Items("_" & intCurID).系统 Then Call AdjustCoordinate
            
            BlnSave = False
            
            If GetSelNum = 1 Then Call ShowAttrib(intCurID)
        Case "条码类型"
            If mshAtt.TextMatrix(mshAtt.Row, 1) = cboAtt.Text Then Exit Sub
            mshAtt.TextMatrix(mshAtt.Row, 1) = cboAtt.Text
            objReport.Items("_" & intCurID).序号 = cboAtt.ItemData(cboAtt.ListIndex)
            BlnSave = False
            
            With objReport.Items("_" & intCurID)
                strBarCode = ReplaceBracket(.内容)
                If strBarCode = "" Then strBarCode = "1234567890"
                
                Unload frmFlash '强制初始Picture，不然切换绘制有问题
                If .序号 = 1 Then
                    Set objBarCode = DrawBarCode128(frmFlash.picTemp, 3, strBarCode, Mid(.表头, 1, 1) = "1")
                ElseIf .序号 = 2 Then
                    Set objBarCode = DrawBarCode39(frmFlash.picTemp, 3, strBarCode, Mid(.表头, 2, 1) = "1", Mid(.表头, 1, 1) = "1")
                ElseIf .序号 = 3 Then
                    If .行高 = 0 Then .行高 = 2
                    Set objBarCode = DrawBarCode128Auto(frmFlash.picTemp, strBarCode, sngWidth, .行高, Mid(.表头, 1, 1) = "1")
                ElseIf .序号 = 10 Then
                    Set objBarCode = DrawBarCode2D(strBarCode, frmFlash.picTemp, lngSize)
                End If
                If Val(Mid(.表头, 3, 1)) <> 0 Then
                    Set objBarCode = PictureSpin(objBarCode, Val(Mid(.表头, 3, 1)), frmFlash.picTemp)
                End If
                Set ImgCode(intCurID).Picture = objBarCode
                
                If .序号 = 3 Then
                    '128码自动调整宽度
                    If Val(Mid(.表头, 3, 1)) = 0 Then
                        ImgCode(intCurID).Width = Format(Me.ScaleX(sngWidth, vbMillimeters, vbTwips) * sgnMode, "0.00")
                        .W = Me.ScaleX(sngWidth, vbMillimeters, vbTwips)
                    Else
                        ImgCode(intCurID).Height = Format(Me.ScaleY(sngWidth, vbMillimeters, vbTwips) * sgnMode, "0.00")
                        .H = Me.ScaleY(sngWidth, vbMillimeters, vbTwips)
                    End If
                    Call SeekItem(ImgCode(intCurID), ImgCode(intCurID).Left, ImgCode(intCurID).Top)
                ElseIf .序号 = 10 Then
                    '二维条码缺省自动调整大小
                    .自调 = True
                    ImgCode(intCurID).Width = Format(lngSize * sgnMode, "0.00")
                    ImgCode(intCurID).Height = Format(lngSize * sgnMode, "0.00")
                    .W = lngSize: .H = lngSize
                    
                    Call SeekItem(ImgCode(intCurID), ImgCode(intCurID).Left, ImgCode(intCurID).Top)
                End If
            End With
            
            If GetSelNum = 1 Then Call ShowAttrib(intCurID)
        Case "条码线宽"
            If mshAtt.TextMatrix(mshAtt.Row, 1) = cboAtt.Text Then Exit Sub
            mshAtt.TextMatrix(mshAtt.Row, 1) = cboAtt.Text
            
            With objReport.Items("_" & intCurID)
                .行高 = Val(cboAtt.Text): BlnSave = False
                
                '重绘图形
                strBarCode = ReplaceBracket(.内容)
                If strBarCode = "" Then strBarCode = "1234567890"
                
                Unload frmFlash '强制初始Picture，不然切换绘制有问题
                If .序号 = 3 Then
                    Set objBarCode = DrawBarCode128Auto(frmFlash.picTemp, strBarCode, sngWidth, .行高, Mid(.表头, 1, 1) = "1")
                End If
                If Val(Mid(.表头, 3, 1)) <> 0 Then
                    Set objBarCode = PictureSpin(objBarCode, Val(Mid(.表头, 3, 1)), frmFlash.picTemp)
                End If
                Set ImgCode(intCurID).Picture = objBarCode
                
                If .序号 = 3 Then
                    '128码自动调整宽度
                    If Val(Mid(.表头, 3, 1)) = 0 Then
                        ImgCode(intCurID).Width = Format(Me.ScaleX(sngWidth, vbMillimeters, vbTwips) * sgnMode, "0.00")
                        .W = Me.ScaleX(sngWidth, vbMillimeters, vbTwips)
                    Else
                        ImgCode(intCurID).Height = Format(Me.ScaleY(sngWidth, vbMillimeters, vbTwips) * sgnMode, "0.00")
                        .H = Me.ScaleY(sngWidth, vbMillimeters, vbTwips)
                    End If
                    Call SeekItem(ImgCode(intCurID), ImgCode(intCurID).Left, ImgCode(intCurID).Top)
                End If
            End With
            
            If GetSelNum = 1 Then Call ShowAttrib(intCurID)
        Case "旋转方向"
            If mshAtt.TextMatrix(mshAtt.Row, 1) = cboAtt.Text Then Exit Sub
            mshAtt.TextMatrix(mshAtt.Row, 1) = cboAtt.Text
            
            LockWindowUpdate Me.hwnd
            
            With objReport.Items("_" & intCurID)
                blnSeek = False
                
                '旋转之后变换宽高
                If Val(Mid(.表头, 3, 1)) <> 0 And cboAtt.ListIndex = 0 _
                    Or Val(Mid(.表头, 3, 1)) = 0 And cboAtt.ListIndex <> 0 Then
                    lngSize = .W: .W = .H: .H = lngSize
                    
                    lngSize = ImgCode(intCurID).Width
                    ImgCode(intCurID).Width = ImgCode(intCurID).Height
                    ImgCode(intCurID).Height = lngSize
                    
                    blnSeek = True
                End If
                .表头 = SetBit(.表头, 3, cboAtt.ListIndex)
                BlnSave = False
                
                '重绘图形
                strBarCode = ReplaceBracket(.内容)
                If strBarCode = "" Then strBarCode = "1234567890"
                
                Unload frmFlash '强制初始Picture，不然切换绘制有问题
                If .序号 = 1 Then
                    Set objBarCode = DrawBarCode128(frmFlash.picTemp, 3, strBarCode, Mid(.表头, 1, 1) = "1")
                ElseIf .序号 = 2 Then
                    Set objBarCode = DrawBarCode39(frmFlash.picTemp, 3, strBarCode, Mid(.表头, 2, 1) = "1", Mid(.表头, 1, 1) = "1")
                ElseIf .序号 = 3 Then
                    Set objBarCode = DrawBarCode128Auto(frmFlash.picTemp, strBarCode, sngWidth, .行高, Mid(.表头, 1, 1) = "1")
                End If
                If Val(Mid(.表头, 3, 1)) <> 0 Then
                    Set objBarCode = PictureSpin(objBarCode, Val(Mid(.表头, 3, 1)), frmFlash.picTemp)
                End If
                Set ImgCode(intCurID).Picture = objBarCode
                
                If .序号 = 3 Then
                    '128码自动调整宽度
                    If Val(Mid(.表头, 3, 1)) = 0 Then
                        ImgCode(intCurID).Width = Format(Me.ScaleX(sngWidth, vbMillimeters, vbTwips) * sgnMode, "0.00")
                        .W = Me.ScaleX(sngWidth, vbMillimeters, vbTwips)
                    Else
                        ImgCode(intCurID).Height = Format(Me.ScaleY(sngWidth, vbMillimeters, vbTwips) * sgnMode, "0.00")
                        .H = Me.ScaleY(sngWidth, vbMillimeters, vbTwips)
                    End If
                    blnSeek = True
                End If
            End With
            
            If blnSeek Then Call SeekItem(ImgCode(intCurID), ImgCode(intCurID).Left, ImgCode(intCurID).Top)
            If GetSelNum = 1 Then Call ShowAttrib(intCurID)
            
            LockWindowUpdate 0
        Case "形状"
            If mshAtt.TextMatrix(mshAtt.Row, 1) = cboAtt.Text Then Exit Sub
            objReport.Items("_" & intCurID).边框 = IIF(cboAtt.Text = "方形", False, True)
            mshAtt.TextMatrix(mshAtt.Row, 1) = cboAtt.Text
            Shp(intCurID).Shape = IIF(cboAtt.Text = "方形", ShapeConstants.vbShapeRectangle, ShapeConstants.vbShapeOval)
            BlnSave = False

    End Select
    
    '重新定位属性行：因为选择属性值之后可能属性行发生了变化
    If InStr("参照对象,方向,性质,条码类型,条码线宽,旋转方向", strCurText) > 0 And GetSelNum = 1 Then
        For lngSize = 1 To mshAtt.Rows - 1
            If mshAtt.TextMatrix(lngSize, 0) = strCurText Then
                mshAtt.Row = lngSize: mshAtt.Col = 1: Exit For
            End If
        Next
    End If
End Sub

Private Sub cboFormat_Change()
    Dim tmpFmt As RPTFmt, bytOrder As Byte

    If blnAllowIn = False Then Exit Sub
    blnRefresh = False
    With cboFormat
        If Trim(.Text) = "" Then
            blnAllowIn = False
            Set .SelectedItem = .ComboItems("_" & mbytCurrFmt)
            blnAllowIn = True
            Exit Sub
        End If

        For Each tmpFmt In objReport.Fmts
            If tmpFmt.说明 = Trim(.Text) Then Exit Sub
        Next

        '修改报表样式名称
        For Each tmpFmt In objReport.Fmts
            If tmpFmt.序号 = mbytCurrFmt Then
                tmpFmt.说明 = Trim(.Text)
                Exit For
            End If
        Next
        blnAllowIn = False
        .ComboItems("_" & mbytCurrFmt).Text = Trim(.Text)
        blnAllowIn = True
        BlnSave = False
    End With
End Sub

Private Sub cboFormat_Validate(Cancel As Boolean)
    Dim tmpFmt As RPTFmt
        
    'zyb#Add
    With cboFormat
        If Trim(.Text) = "" Then
            blnAllowIn = False
            Set .SelectedItem = .ComboItems("_" & mbytCurrFmt)
            blnAllowIn = True
            Exit Sub
        End If
        
        For Each tmpFmt In objReport.Fmts
            If tmpFmt.序号 <> mbytCurrFmt And tmpFmt.说明 = Trim(.Text) Then
                MsgBox "当前输入的报表格式名称已经存在，请重新输入！"
                blnAllowIn = False
                Set .SelectedItem = .ComboItems("_" & mbytCurrFmt)
                blnAllowIn = True
                .SetFocus
                Cancel = True
                Exit Sub
            End If
        Next
    End With
End Sub

Private Sub cboText_Click()
    Dim strCurText As String, intType As Integer
    Dim ObjSel As Object
    
    strCurText = mshAtt.TextMatrix(mshAtt.Row, 0)
    Set ObjSel = GetInxObj(intCurID)
    
    Select Case strCurText
        Case "内容"
            mshAtt.TextMatrix(mshAtt.Row, 1) = cboText.Text
            If UCase(TypeName(ObjSel)) = "LABEL" Then ObjSel.Caption = cboText.Text
            objReport.Items("_" & intCurID).内容 = cboText.Text
        
            '自调后须调整LblSize控件的位置
            If UCase(TypeName(ObjSel)) = "LABEL" Then
                If ObjSel.AutoSize Then
                    'Call SelItem(ObjSel.Index, False)
                    '卸载会报错，移动位置即可
                    
                    Call SelMove(ObjSel.Index)
                End If
                objReport.Items("_" & intCurID).W = lbl(intCurID).Width / sgnMode
            End If
            
            cboText.Visible = False: mshAtt.SetFocus
            BlnSave = False
    End Select
    
End Sub

Private Sub cboText_KeyPress(KeyAscii As Integer)
    Dim ObjSel As Object
    Dim xx As Integer, yy As Integer, zz As Integer
    Dim strBarCode As String, objBarCode As StdPicture
    Dim strTemp As String
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        '非法或分隔字符:
        If InString(cboText.Text, "'|~^") Then
            MsgBox "输入了非法字符！", vbInformation, App.Title
            cboText.SetFocus: Exit Sub
        End If
        Set ObjSel = GetInxObj(intCurID)

        Select Case mshAtt.TextMatrix(mshAtt.Row, 0)
            Case "内容"
                If TLen(cboText.Text) > 255 Then
                    MsgBox "内容不能超过255个字符！", vbInformation, App.Title
                    cboText.SetFocus: Exit Sub
                End If
                
                Dim strNodeName As String, NodeThis As Node
                '如果是adLongVarBinary型字段,则不允许修改
                xx = InStr(1, cboText, "]")
                yy = InStr(1, cboText, ".")
                zz = InStr(1, cboText, "[")
                If xx > zz And xx > yy And xx <> 0 And zz <> 0 Then
                    strNodeName = Mid(cboText, yy + 1, xx - yy - 1)
                    For Each NodeThis In tvwSQL.Nodes
                        If mdlPublic.GetStdNodeText(NodeThis.Text) = strNodeName And IsType(Val(NodeThis.Tag), adLongVarBinary) Then
                            MsgBox "不能选择图型字段为标签的内容！", vbInformation, App.Title
                            mshAtt.TextMatrix(mshAtt.Row, 1) = objReport.Items("_" & intCurID).内容
                            Exit Sub
                        End If
                    Next
                End If
                
                mshAtt.TextMatrix(mshAtt.Row, 1) = cboText.Text
                If UCase(TypeName(ObjSel)) = "LABEL" Then ObjSel.Caption = cboText.Text
                objReport.Items("_" & intCurID).内容 = cboText.Text
            
                '自调后须调整LblSize控件的位置
                If UCase(TypeName(ObjSel)) = "LABEL" Then
                    If ObjSel.AutoSize Then
                        Call SelItem(ObjSel.Index, False)
                        Call SelItem(ObjSel.Index, True)
                    End If
                    objReport.Items("_" & intCurID).W = lbl(intCurID).Width / sgnMode
                End If
                
                cboText.Visible = False: mshAtt.SetFocus
                BlnSave = False
        End Select
    Else
        Select Case mshAtt.TextMatrix(mshAtt.Row, 0)
            Case "内容"
                If InStr("'|~^", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
        End Select
    End If
End Sub

Private Sub cbr_HeightChanged(ByVal NewHeight As Single)
    Form_Resize
End Sub

Private Sub CboFormat_Click()
    'zyb#Add
    If Trim(cboFormat.Text) = "" Then Exit Sub
    If mbytCurrFmt <> Mid(cboFormat.SelectedItem.Key, 2) Or blnRefresh Then
        blnRefresh = False
        mbytCurrFmt = Mid(cboFormat.SelectedItem.Key, 2)
        
        Call ShowSize: Call ShowScroll
        Call ReFlashReportBySelFormat
        Call picPaper_MouseDown(1, 0, 0, 0) '显示基本属性
    End If
End Sub

Private Sub CboFormat_KeyPress(KeyAscii As Integer)
    'zyb#Add
    If blnDelReportFormat = False Then KeyAscii = 0: Exit Sub
    If KeyAscii = 39 Or KeyAscii = 22 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Chart_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lngPreX = X: lngPreY = Y
    If Button = 1 Then
        If Shift = 2 Then
            If Mid(Chart(Index).Tag, 1, 2) = "" Then
                Call SelItem(Index, True) '加选
                If GetSelNum() = 1 Then
                    Call ShowAttrib(Index) '只选中一个则显示属性
                Else
                    Call ShowAttrib '多选时不显示属性
                End If
            Else
                Call SelItem(Index, False) '反选
                If GetSelNum() = 1 Then
                    Call ShowAttrib(intCurID) '只选中一个则显示属性(选中的不一定是该控件)
                Else
                    Call ShowAttrib '多选时不显示属性
                End If
            End If
        Else
            If Mid(Chart(Index).Tag, 1, 2) = "" Then
                Call SelClear
                Call SelItem(Index, True)
                Call ShowAttrib(Index) '只选中一个则显示属性
            End If
        End If
    ElseIf Button = 2 Then
        PopupMenu mnuFormat, 2
    End If
End Sub

Private Sub Chart_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ObjSel As Object
    
    If objReport.Items("_" & Index).系统 Then Exit Sub
    DrawXY X + Chart(Index).Left, Y + Chart(Index).Top
    Set ObjSel = Chart(Index)
    If Button = 1 And Mid(ObjSel.Tag, 1, 2) <> "" Then
        If blnLock Then Exit Sub
        Call MoveSelect(X - lngPreX, Y - lngPreY)
        If GetSelNum() = 1 Then ShowAttrib Index
    Else
        If objReport.Items("_" & Index).类型 <> 12 Then
            ObjSel.MousePointer = 99
        End If
    End If
End Sub

Private Sub cmdAdd_Click()
    Dim bytOrder As Byte, strRptFmtName As String
    Dim objFmt As RPTFmt
    
    '增加报表格式：纸张缺省与第一个格式相同
    With cboFormat
        bytOrder = .ComboItems.count + 1
        If bytOrder < 100 Then
            strRptFmtName = GetRPtFmtName
            Set objFmt = objReport.Fmts(1)
            objReport.Fmts.Add bytOrder, strRptFmtName, objFmt.W, objFmt.H, objFmt.纸张, objFmt.纸向, objFmt.动态纸张, 0, "_" & bytOrder
        
            blnAllowIn = False
            .ComboItems.Add , "_" & bytOrder, strRptFmtName, "Format"
            .ComboItems("_" & bytOrder).Selected = True
            Set .SelectedItem = .ComboItems("_" & bytOrder)
            .SetFocus
            
            blnAllowIn = True
            blnRefresh = True
            BlnSave = False
        Else
            MsgBox "报表格式太多，不能继续增加！（最多99种格式）", vbInformation, App.Title
            .SetFocus
            Exit Sub
        End If
    End With
    
    cmdDel.Enabled = (cboFormat.ComboItems.count > 1) And blnDelReportFormat
    tbr1.Buttons("DelFormat").Enabled = cmdDel.Enabled
    mnuEdit_DelFormat.Enabled = cmdDel.Enabled
    mbytCurrFmt = bytOrder
    
    Call CboFormat_Click
End Sub

Private Sub cmdDel_Click()
    Dim bytFmt As Byte, tmpFmt As RPTFmt, tmpItem As RPTItem
    Dim intModify As Integer, intDel As Integer
    
    '删除报表样式,并更新集合
    'zyb#Add
    With cboFormat
        If .ComboItems.count < 2 Then cmdDel.Enabled = False: Exit Sub
        If .SelectedItem Is Nothing Then Exit Sub
        If MsgBox("你确定要删除该格式吗？（删除后将不可恢复）", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
        
        bytFmt = Mid(.SelectedItem.Key, 2)
        .ComboItems.Remove "_" & bytFmt
        For intModify = bytFmt To .ComboItems.count
            If Mid(.ComboItems("_" & intModify + 1).Key, 2) > bytFmt Then
                .ComboItems("_" & intModify + 1).Key = "_" & (Mid(.ComboItems("_" & intModify + 1).Key, 2) - 1)
            End If
        Next
        Set .SelectedItem = .ComboItems("_" & IIF(bytFmt = 1, 1, bytFmt - 1))
    End With
    
    '删除并增减报表样式的序号
    objReport.Fmts.Remove "_" & bytFmt
    For intDel = bytFmt + 1 To objReport.Fmts.count + 1
        With objReport.Fmts("_" & intDel)
            If .序号 > bytFmt Then
                objReport.Fmts.Add intDel - 1, .说明, .W, .H, .纸张, .纸向, .动态纸张, .图样, "_" & intDel - 1
                objReport.Fmts.Remove "_" & intDel
            End If
        End With
    Next
    '删除该报表样式对应的所有报表元素
    For Each tmpItem In objReport.Items
        If tmpItem.格式号 = bytFmt Then
            objReport.Items.Remove "_" & tmpItem.Key
        End If
    Next
    '修改其余的报表元素所对应的格式号
    For intDel = 1 To objReport.Items.count
        Set tmpItem = objReport.Items(intDel)
        If tmpItem.格式号 > bytFmt Then
            tmpItem.格式号 = tmpItem.格式号 - 1
        End If
    Next
    
    blnRefresh = True
    BlnSave = False
    cmdDel.Enabled = (cboFormat.ComboItems.count > 1) And blnDelReportFormat
    Call CboFormat_Click
End Sub

Private Function GetInxObj(ByVal intIndex As Integer) As Object
'功能：根据ID，获得对应的元素对象
    Dim ObjSel As Object
    
    Select Case objReport.Items("_" & intIndex).类型
        Case 1
            Set ObjSel = lblLine(intIndex)
        Case 2, 3
            Set ObjSel = lbl(intIndex)
        Case 10
            Set ObjSel = Shp(intIndex)
        Case 4, 5
            Set ObjSel = msh(intIndex)
        Case 11
            Set ObjSel = img(intIndex)
        Case 12 '@@@
            Set ObjSel = Chart(intIndex)
        Case 13
            Set ObjSel = ImgCode(intIndex)
        Case 14
            Set ObjSel = pic(intIndex)
    End Select
    Set GetInxObj = ObjSel
End Function

Private Sub SetCmdAttBackColor(ByVal intIndex As Integer)
'功能：设置制定元素的背景色
    Dim ObjSel As Object
    On Error Resume Next
    Set ObjSel = GetInxObj(intIndex)
    '控件属性
    If objReport.Items("_" & intIndex).类型 = 10 Then
        objReport.Items("_" & intIndex).背景 = cdg.Color '先赋值
        Call DrawFrame(ObjSel)
    ElseIf objReport.Items("_" & intIndex).类型 = 12 Then '@@@
        ObjSel.Interior.BackgroundColor = IIF(cdg.Color = &HFFFFFF, lbl(0).BackColor, cdg.Color) '白色区分
        objReport.Items("_" & intIndex).背景 = cdg.Color
    Else
        If cdg.Color = &HFFFFFF Then
            If objReport.Items("_" & intIndex).类型 = 4 Or objReport.Items("_" & intIndex).类型 = 5 Then
                ObjSel.BackColor = cdg.Color
                ObjSel.BackColorFixed = lbl(0).BackColor '白色区分出是固定行列
            Else
                ObjSel.BackColor = lbl(0).BackColor '白色区分
            End If
        Else
            ObjSel.BackColor = cdg.Color
            If objReport.Items("_" & intIndex).类型 = 4 Or objReport.Items("_" & intIndex).类型 = 5 Then
                ObjSel.BackColorFixed = cdg.Color
            End If
        End If
        If objReport.Items("_" & intIndex).类型 = 4 Or objReport.Items("_" & intIndex).类型 = 5 Then
            Call ResetColor(ObjSel.Index) '很怪,必须逐个单元刷新
            If objReport.Items("_" & intIndex).类型 = 4 Then
                Call SetCopyGrid(intIndex)
            End If
        End If
        objReport.Items("_" & intIndex).背景 = cdg.Color
    End If
End Sub

Private Sub SetCmdAttForeColor(ByVal intIndex As Integer)
'功能：设置制定元素的前景色
    Dim ObjSel As Object
    On Error Resume Next
    
    Set ObjSel = GetInxObj(intIndex)
    If objReport.Items("_" & intIndex).类型 = 1 Then
        '线条是以背景色显示
        If cdg.Color = &HFFFFFF Then
            ObjSel.BackColor = lbl(0).BackColor '白色时显示谈色
        Else
            ObjSel.BackColor = cdg.Color
        End If
    ElseIf objReport.Items("_" & intIndex).类型 = 12 Then '@@@
        ObjSel.Interior.ForegroundColor = cdg.Color
        '不知为什么仅设置控件前景无效,但通过属性框就有效
        ObjSel.ChartArea.Axes("X").AxisStyle.LineStyle.Color = cdg.Color
        ObjSel.ChartArea.Axes("Y").AxisStyle.LineStyle.Color = cdg.Color
    Else
        ObjSel.ForeColor = cdg.Color
    End If
    
    '表格还有固定行列前景色
    If objReport.Items("_" & intIndex).类型 = 4 Or objReport.Items("_" & intIndex).类型 = 5 Then
        ObjSel.ForeColorFixed = cdg.Color

        Call ResetColor(ObjSel.Index) '很怪,必须逐个单元刷新
        If objReport.Items("_" & intIndex).类型 = 4 Then
            Call SetCopyGrid(intIndex)
        End If
    End If
    
    '对象值
    objReport.Items("_" & intIndex).前景 = cdg.Color
End Sub

Private Sub SetCmdAttFont(ByVal intIndex As Integer)
'功能：设置制定元素的字体
    Dim ObjSel As Object, sgnH As Single
    Dim i As Long
    
    On Error Resume Next
    Set ObjSel = GetInxObj(intIndex)
    '对象内容
    objReport.Items("_" & intIndex).字体 = cdg.FontName
    objReport.Items("_" & intIndex).字号 = Format(cdg.FontSize, "0.0") '允许小字号@@@
    objReport.Items("_" & intIndex).粗体 = cdg.FontBold
    objReport.Items("_" & intIndex).斜体 = cdg.FontItalic
    If objReport.Items("_" & intIndex).类型 <> 12 Then '@@@
        objReport.Items("_" & intIndex).下线 = cdg.FontUnderline
        objReport.Items("_" & intIndex).前景 = cdg.Color '允许白色前景@@@
    End If

    '控件属性
    If objReport.Items("_" & intIndex).类型 = 12 Then '@@@
        Call SetChartStyleAndData(ObjSel, objReport.Items("_" & intIndex), , sgnMode, True)
    Else
        '为测试字体控件装入字体属性
        PicFontTest.Font.name = cdg.FontName
        PicFontTest.Font.Size = cdg.FontSize * sgnMode '允许小字号@@@
        PicFontTest.Font.Bold = cdg.FontBold
        PicFontTest.Font.Italic = cdg.FontItalic
        PicFontTest.Font.Underline = cdg.FontUnderline
        sgnH = (PicFontTest.TextHeight("字") + 15) * sgnMode
        
        ObjSel.Font.name = cdg.FontName
        ObjSel.Font.Size = cdg.FontSize * sgnMode '允许小字号@@@
        ObjSel.Font.Bold = cdg.FontBold
        ObjSel.Font.Italic = cdg.FontItalic
        ObjSel.Font.Underline = cdg.FontUnderline
        ObjSel.ForeColor = cdg.Color '允许白色前景@@@
        If TypeName(ObjSel) = "VSFlexGrid" Then
            ObjSel.ForeColorFixed = ObjSel.ForeColor
        End If
    End If
    
    '根据字体及可显示固定列数自动调整表格行高(过矮时)
    Select Case objReport.Items("_" & intIndex).类型
        Case 4, 5
            If ObjSel.RowHeight(0) < sgnH Then
                If Abs(Int(-ObjSel.Height / sgnH)) >= ObjSel.FixedRows + 2 Then
                    For i = 0 To ObjSel.Rows - 1
                        ObjSel.RowHeight(i) = sgnH
                    Next
                Else
                    For i = 0 To ObjSel.Rows - 1
                        ObjSel.RowHeight(i) = Abs(Int(-ObjSel.Height / (ObjSel.FixedRows + 2)))
                    Next
                End If
                objReport.Items("_" & intIndex).行高 = Format(ObjSel.RowHeight(0) / sgnMode, "0.00")
                mshAtt.TextMatrix(GetRow("行高"), 1) = Format(ObjSel.RowHeight(0) / Twip_mm / sgnMode, "0.00")
                Call SetGridLine(intIndex)
            End If
            Call ResetColor(ObjSel.Index) '很怪,必须逐个单元刷新
            If objReport.Items("_" & intIndex).类型 = 4 Then
                Call SetCopyGrid(intIndex)
            End If
        Case 2, 3
            If ObjSel.Height < sgnH Then
                ObjSel.Height = sgnH: ObjSel.Width = PicFontTest.TextWidth("字") * TLen(ObjSel.Text) / 2
                objReport.Items("_" & intIndex).H = ObjSel.Height / sgnMode
                objReport.Items("_" & intIndex).W = ObjSel.Width / sgnMode
                mshAtt.TextMatrix(GetRow("高度"), 1) = Format(ObjSel.Height / Twip_mm / sgnMode, "0.00")
                mshAtt.TextMatrix(GetRow("宽度"), 1) = Format(ObjSel.Width / Twip_mm / sgnMode, "0.00")
            End If
    End Select
    
    i = intIndex
    intIndex = ObjSel.Index
    Call ReferTo
    intIndex = i
    
    Call SelItem(ObjSel.Index, False)
    Call SelItem(ObjSel.Index, True)
End Sub

Private Sub cmdAtt_Click()
    Dim ObjSel As Object, i As Integer
    Dim tmpItem As RPTItem
    Dim strInfo As String
    Dim lngReportID As Long
    Dim strReportID As String
    Dim X As Long, Y As Long, k As Long
    Dim tmpObj As PictureBox
    
    Set ObjSel = GetInxObj(intCurID)
    
    On Error Resume Next
    cdg.CancelError = True
    Select Case mshAtt.TextMatrix(mshAtt.Row, 0)
        Case "内容"
            Set ObjSel = img(intCurID)
            cdg.Flags = &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
            cdg.Filter = "图片文件(*.ico;*.cur;*.bmp;*.gif;*.jpg;*.rle;*.wmf;*.emf)|*.ico;*.cur;*.bmp;*.gif;*.jpg;*.rle;*.wmf;*.emf"
            cdg.InitDir = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name, "图片路径", "C:\")
            cdg.ShowOpen
            If Err.Number = 0 Then
                ObjSel.Picture = LoadPicture(cdg.FileName)
                If Err.Number <> 0 Then
                    MsgBox "选择的图片文件格式错误！", vbInformation, App.Title
                    Set ObjSel.Picture = Nothing
                    Exit Sub
                End If
                
                '先保存备用
                Set objReport.Items("_" & intCurID).图片 = ObjSel.Picture
                
                '路径保存至注册表
                SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name, "图片路径", Replace(cdg.FileName, cdg.FileTitle, "") 'Mid(cdg.FileName, 1, Len(cdg.FileName) - InStr(1, strReverse(cdg.FileName), "\"))
                
                '调整选择控件
                If objReport.Items("_" & intCurID).自调 Then
                    ObjSel.Stretch = False
                    ObjSel.Width = ObjSel.Width * sgnMode
                    ObjSel.Height = ObjSel.Height * sgnMode
                    Call SelItem(ObjSel.Index, False)
                    Call SelItem(ObjSel.Index, True)
                    ObjSel.Stretch = True
                End If
                '保持比例
                If objReport.Items("_" & intCurID).粗体 Then
                    Set ObjSel.Picture = ScalePicture(PicFontTest, objReport.Items("_" & intCurID).图片, ObjSel.Width, ObjSel.Height)
                End If
                
                '赋值
                With objReport.Items("_" & intCurID)
                    .内容 = cdg.FileName
                    .X = Format(ObjSel.Left / sgnMode, "0.00")
                    .Y = Format(ObjSel.Top / sgnMode, "0.00")
                    .H = Format(ObjSel.Height / sgnMode, "0.00")
                    .W = Format(ObjSel.Width / sgnMode, "0.00")
                    mshAtt.TextMatrix(mshAtt.Row, 1) = "[Picture]"
                    mshAtt.TextMatrix(GetRow("X坐标"), 1) = Format(.X / Twip_mm, "0.00")
                    mshAtt.TextMatrix(GetRow("Y坐标"), 1) = Format(.Y / Twip_mm, "0.00")
                    mshAtt.TextMatrix(GetRow("高度"), 1) = Format(.H / Twip_mm, "0.00")
                    mshAtt.TextMatrix(GetRow("宽度"), 1) = Format(.W / Twip_mm, "0.00")
                End With
                
                BlnSave = False
            Else
                Err.Clear
            End If
        Case "表体网格色", "网格色"
            cdg.CancelError = True
            cdg.Flags = &H1 Or &H2
            cdg.Color = objReport.Items("_" & intCurID).网格
            cdg.ShowColor
            If Err.Number = 0 Then
                '属性值更新
                mshAtt.Col = 1
                mshAtt.CellForeColor = cdg.Color
                
                ObjSel.GridColor = cdg.Color
                If mshAtt.TextMatrix(mshAtt.Row, 0) = "网格色" Then
                    ObjSel.GridColorFixed = cdg.Color
                End If
                
                Call ResetColor(ObjSel.Index) '很怪,必须逐个单元刷新
                If objReport.Items("_" & intCurID).类型 = 4 Then
                    Call SetCopyGrid(intCurID)
                End If
                
                '对象值
                objReport.Items("_" & intCurID).网格 = cdg.Color
                BlnSave = False
            Else
                Err.Clear
            End If
                
            '设置其子表与主表相关属性一致
            If objReport.Items("_" & intCurID).参照 = "" And objReport.Items("_" & intCurID).类型 = 5 Then
                For Each tmpItem In objReport.Items
                    If tmpItem.格式号 = mbytCurrFmt And tmpItem.参照 = objReport.Items("_" & intCurID).名称 And tmpItem.类型 = 5 Then
                        Call SetGridLike(msh(intCurID), msh(tmpItem.Key))
                    End If
                Next
            End If
        Case "表头网格色"
            cdg.CancelError = True
            cdg.Flags = &H1 Or &H2
            cdg.Color = IIF(objReport.Items("_" & intCurID).格式 = "", objReport.Items("_" & intCurID).网格, Val(objReport.Items("_" & intCurID).格式))
            cdg.ShowColor
            If Err.Number = 0 Then
                '属性值更新
                mshAtt.Col = 1
                mshAtt.CellForeColor = cdg.Color
                
                ObjSel.GridColorFixed = cdg.Color
                
                Call ResetColor(ObjSel.Index) '很怪,必须逐个单元刷新
                If objReport.Items("_" & intCurID).类型 = 4 Then
                    Call SetCopyGrid(intCurID)
                End If
                
                '对象值
                objReport.Items("_" & intCurID).格式 = cdg.Color
                BlnSave = False
            Else
                Err.Clear
            End If
                
        Case "前景色"
            cdg.Flags = &H1 Or &H2
            cdg.Color = objReport.Items("_" & intCurID).前景
            cdg.ShowColor
            If Err.Number = 0 Then
                '属性值更新
                mshAtt.Col = 1
                mshAtt.CellForeColor = cdg.Color
                
                '如果被选中的所有元素都是相同的元素
                If lblSize.count > 9 Then
                    For Each tmpObj In lblSize
                        If tmpObj.Index Mod 8 = 1 Then
                            Call SetCmdAttForeColor(tmpObj.Tag)
                        End If
                    Next
                Else
                    Call SetCmdAttForeColor(intCurID)
                End If
                
                BlnSave = False
            Else
                Err.Clear
            End If
                
            '设置其子表与主表相关属性一致
            If objReport.Items("_" & intCurID).参照 = "" And objReport.Items("_" & intCurID).类型 = 5 Then
                For Each tmpItem In objReport.Items
                    If tmpItem.格式号 = mbytCurrFmt And tmpItem.参照 = objReport.Items("_" & intCurID).名称 And tmpItem.类型 = 5 Then
                        Call SetGridLike(msh(intCurID), msh(tmpItem.Key))
                    End If
                Next
            End If
        Case "背景色"
            cdg.Flags = &H1 Or &H2
            cdg.Color = objReport.Items("_" & intCurID).背景
            cdg.ShowColor
            If Err.Number = 0 Then
                '属性值更新
                mshAtt.Col = 1
                mshAtt.CellForeColor = cdg.Color
                '如果被选中的所有元素都是相同的元素
                If lblSize.count > 9 Then
                    For Each tmpObj In lblSize
                        If tmpObj.Index Mod 8 = 1 Then
                            Call SetCmdAttBackColor(tmpObj.Tag)
                        End If
                    Next
                Else
                    Call SetCmdAttBackColor(intCurID)
                End If
                
                BlnSave = False
            Else
                Err.Clear
            End If
                
            '设置其子表与主表相关属性一致
            If objReport.Items("_" & intCurID).参照 = "" And objReport.Items("_" & intCurID).类型 = 5 Then
                For Each tmpItem In objReport.Items
                    If tmpItem.格式号 = mbytCurrFmt And tmpItem.参照 = objReport.Items("_" & intCurID).名称 And tmpItem.类型 = 5 Then
                        Call SetGridLike(msh(intCurID), msh(tmpItem.Key))
                    End If
                Next
            End If
        Case "字体" '可以改变字体,字号,粗体,斜体
            cdg.Flags = &H3 Or &H400 Or &H200 Or &H10000
            If objReport.Items("_" & intCurID).类型 <> 12 Then '@@@
                cdg.Flags = cdg.Flags Or &H100
            End If
            cdg.FontName = objReport.Items("_" & intCurID).字体
            cdg.FontSize = objReport.Items("_" & intCurID).字号
            cdg.FontBold = objReport.Items("_" & intCurID).粗体
            cdg.FontItalic = objReport.Items("_" & intCurID).斜体
            If objReport.Items("_" & intCurID).类型 <> 12 Then '@@@
                cdg.FontUnderline = objReport.Items("_" & intCurID).下线
                cdg.Color = objReport.Items("_" & intCurID).前景
            End If
            
            cdg.ShowFont
            If Err.Number = 0 Then
                mshAtt.TextMatrix(mshAtt.Row, Val("1-设置列")) = cdg.FontName
                '如果被选中的所有元素都是相同的元素
                If lblSize.count > 9 Then
                    For Each tmpObj In lblSize
                        If tmpObj.Index Mod 8 = 1 Then
                            Call SetCmdAttFont(tmpObj.Tag)
                        End If
                    Next
                Else
                    Call SetCmdAttFont(intCurID)
                End If
                BlnSave = False
            Else
                Err.Clear
            End If
            
            '设置其子表与主表相关属性一致
            If objReport.Items("_" & intCurID).参照 = "" And objReport.Items("_" & intCurID).类型 = 5 Then
                For Each tmpItem In objReport.Items
                    If tmpItem.格式号 = mbytCurrFmt And tmpItem.参照 = objReport.Items("_" & intCurID).名称 And tmpItem.类型 = 5 Then
                        Call SetGridLike(msh(intCurID), msh(tmpItem.Key))
                    End If
                Next
            End If
        Case "设置"
            Set tmpItem = objReport.Items("_" & intCurID)
            If frmChartSetup.ShowMe(Me, objReport.Datas, ObjSel, tmpItem) Then
                Call CopyItem(objReport.Items("_" & intCurID), tmpItem, False)
                Call ShowAttrib(intCurID)
                mshAtt.Row = 3
                mshAtt_AfterRowColChange 0, 0, mshAtt.Row, mshAtt.Col
                BlnSave = False
            End If
        Case "关联报表"
            X = InStr(1, objReport.Items("_" & intCurID).内容, "]")
            Y = InStr(1, objReport.Items("_" & intCurID).内容, ".")
            k = InStr(1, objReport.Items("_" & intCurID).内容, "[")
            If X > k And X > Y And X <> 0 And k <> 0 Then
                strReportID = FindReport("", txtAtt.hwnd, strInfo, objReport.Items("_" & intCurID).Relations.Item(1).关联报表ID, objReport, objReport.Items("_" & intCurID).Relations, 2, Me, intCurID)
                If strReportID <> "" Then
                    mshAtt.TextMatrix(mshAtt.Row, 1) = strInfo
                    mshAtt.RowData(mshAtt.Row) = strReportID
                    txtAtt.Visible = False: mshAtt.SetFocus
                    BlnSave = False
                Else
                    '判断是取消还是清除
                    If objReport.Items("_" & intCurID).Relations.count > 0 Then
                        txtAtt.SetFocus
                    Else
                        mshAtt.TextMatrix(mshAtt.Row, 1) = ""
                        mshAtt.RowData(mshAtt.Row) = 0
                        txtAtt.Text = ""
                        txtAtt.Visible = True
                        txtAtt.SetFocus
                        BlnSave = False
                    End If
                End If
            Else
                MsgBox "当前标签必须先绑定一个数据源，例如：[数据源.字段],绑定后再设置关联报表。", vbInformation, Me.Caption
            End If
    End Select
    
    mshAtt.SetFocus
End Sub

Private Sub dtpAtt_Change()
    Select Case mshAtt.TextMatrix(mshAtt.Row, 0)
        Case "禁止开始时间"
            objReport.禁止开始时间 = dtpAtt.Value
            mshAtt.TextMatrix(mshAtt.Row, 1) = Format(objReport.禁止开始时间, "HH:mm:ss")
            
            dtpAtt.Visible = False: If dtpAtt.Visible Then dtpAtt.SetFocus
            BlnSave = False
        Case "禁止结束时间"
            objReport.禁止结束时间 = dtpAtt.Value
            mshAtt.TextMatrix(mshAtt.Row, 1) = Format(objReport.禁止结束时间, "HH:mm:ss")
            
            dtpAtt.Visible = False: If dtpAtt.Visible Then dtpAtt.SetFocus
            BlnSave = False
    End Select
End Sub

Private Sub Form_Activate()
    If mblnFirst Then
        mblnFirst = False
        Me.Refresh
        On Error Resume Next
        picPaper.SetFocus
        Call picPaper_MouseDown(1, 0, 0, 0)
    End If
End Sub

Private Sub Form_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Call ClearXY
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnScroll As Boolean
    
    If Shift = 0 Then
        If ActiveControl Is Nothing Then
            blnScroll = True
        ElseIf TypeName(ActiveControl) <> "TextBox" _
            And TypeName(ActiveControl) <> "ComboBox" _
            And TypeName(ActiveControl) <> "ListView" _
            And UCase(ActiveControl.name) <> "MSHATT" Then
            blnScroll = True
        End If
    End If
    
    Select Case KeyCode
        Case vbKeyUp
            If blnScroll And scrVsc.Enabled And scrVsc.Value > scrVsc.Min Then scrVsc.Value = IIF(scrVsc.Value - scrVsc.LargeChange < scrVsc.Min, scrVsc.Min, scrVsc.Value - scrVsc.LargeChange)
            If Shift = 2 And GetSelNum > 0 Then
                MoveSelect 0, -15
                If GetSelNum = 1 Then ShowAttrib intCurID
            End If
        Case vbKeyDown
            If blnScroll And scrVsc.Enabled And scrVsc.Value < scrVsc.Max Then scrVsc.Value = IIF(scrVsc.Value + scrVsc.LargeChange > scrVsc.Max, scrVsc.Max, scrVsc.Value + scrVsc.LargeChange)
            If Shift = 2 And GetSelNum > 0 Then
                MoveSelect 0, 15
                If GetSelNum = 1 Then ShowAttrib intCurID
            End If
        Case vbKeyLeft
            If blnScroll And scrHsc.Enabled And scrHsc.Value > scrHsc.Min Then scrHsc.Value = IIF(scrHsc.Value - scrHsc.LargeChange < scrHsc.Min, scrHsc.Min, scrHsc.Value - scrHsc.LargeChange)
            If Shift = 2 And GetSelNum > 0 Then
                MoveSelect -15, 0
                If GetSelNum = 1 Then ShowAttrib intCurID
            End If
        Case vbKeyRight
            If blnScroll And scrHsc.Enabled And scrHsc.Value < scrHsc.Max Then scrHsc.Value = IIF(scrHsc.Value + scrHsc.LargeChange > scrHsc.Max, scrHsc.Max, scrHsc.Value + scrHsc.LargeChange)
            If Shift = 2 And GetSelNum > 0 Then
                MoveSelect 15, 0
                If GetSelNum = 1 Then ShowAttrib intCurID
            End If
    End Select
    
    If Shift = 4 Then
        If picPaper.MousePointer <> 99 Then
            Set picPaper.MouseIcon = picBack.MouseIcon
            picPaper.MousePointer = 99
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        '取消属性编辑状态
        KeyAscii = 0
        If txtAtt.Visible Or cmdAtt.Visible Or cboAtt.Visible Then
            txtAtt.Visible = False
            cmdAtt.Visible = False
            cboAtt.Visible = False
            cboAtt.Clear: txtAtt.Text = ""
            mshAtt.SetFocus
        End If
    Else
        If InStr("'&", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub Form_Keyup(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then picPaper.MousePointer = 0
End Sub

Private Sub Form_Load()
    Dim rsDelReportFormat As New ADODB.Recordset
    Dim strSQL As String
    
    Screen.MousePointer = vbHourglass
        
    RestoreWinState Me, App.ProductName
    tb2.ZOrder
    If mblnNotModiData Then
        mnuEdit_New.Visible = False
        mnuEdit_Modi.Visible = False
        mnuEdit_Del.Visible = False
        mnuEdit_Data_.Visible = False
        tbr1.Buttons("New").Visible = False
        tbr1.Buttons("Modi").Visible = False
        tbr1.Buttons("Del").Visible = False
        tbr1.Buttons("Data_").Visible = False
    End If
    
    '显示比例
    sgnMode = 1: sgnLastMode = 1
    mblnFirst = True
    
    picR.Visible = mnuViewToolAttrib.Checked
    picAtt.Visible = mnuViewToolAttrib.Checked

    picL.Visible = mnuViewToolSQL.Checked
    picSQL.Visible = mnuViewToolSQL.Checked

    picRulerH.Visible = mnuViewToolRuler.Checked
    picRulerV.Visible = mnuViewToolRuler.Checked

    '初始化旋转字体对象
    Set objFont = New clsRotateFont
    Set objFont.LogFont = New StdFont
    objFont.LogFont.name = "Times New Roman"
    objFont.LogFont.Size = 6.5
    objFont.Rotation = 90
    
    gblnModi = False
    intMaxID = 0: intCurID = 0
    blnLock = True
    bytCurTool = 0
    BlnSave = True
    Set objLastSel = Nothing
    
    selCell.Col1 = -1: selCell.Col2 = -1: selCell.Row1 = -1: selCell.Row2 = -1: selCell.Row = -1: sta.Panels(2).Text = ""
    intCurCol = -1
    
    'zyb#Add
    '初始化显示比例菜单
    mnuViewScaleMode(0).Checked = False
    mnuViewScaleMode(1).Checked = False
    mnuViewScaleMode(2).Checked = False
    mnuViewScaleMode(4).Checked = False
    mnuViewScaleMode(5).Checked = True
    mnuViewScaleMode(6).Checked = False
    mnuViewScaleMode(7).Checked = False
    mnuViewScaleMode(8).Checked = False
    
    '读取同时改变intMaxID
    blnAdjustRowHeight = False
    Set objReport = ReadReport(lngRPTID, intMaxID)
    Call GetInPaper
    
    If Not objReport Is Nothing Then
        '获取该报表是否为固定报表
        strSQL = "Select Nvl(系统,0) 系统 From zlReports Where ID=[1]"
        Set rsDelReportFormat = OpenSQLRecord(strSQL, Me.Caption, lngRPTID)
        blnDelReportFormat = (rsDelReportFormat!系统 = 0)
        blnDelReportFormat = True '200312:固定报表也允许
        
        '显示报表内容(数据源、纸张、元素)
        'zyb#Modify
        Call LoadReportFormat       '读取该报表的所有格式
        Call ReFlashReport(False)
        
        Caption = Caption & " - [" & objReport.编号 & "]" & objReport.名称 & IIF(objReport.说明 = "", "", "：" & objReport.说明)
    Else
        Screen.MousePointer = vbDefault
        MsgBox "不能正确读取报表内容！", vbInformation, App.Title
        Unload Me: Exit Sub
    End If
    
    Screen.MousePointer = vbDefault
    If objReport.Items.count = 0 Then mnuFormat_Lock_Click
End Sub

Private Sub ShowPaperInfo()
    Dim objFmt As RPTFmt
    
    If Not objReport Is Nothing Then
        Set objFmt = objReport.Fmts("_" & mbytCurrFmt)
        sta.Panels(2).Text = "打印机:" & objReport.打印机 & "   纸张:" & GetPaperName(objFmt.纸张, objFmt.W, objFmt.H) & " " & _
            IIF(objFmt.纸张 = 256, CInt(objFmt.W / Twip_mm) & "mm × " & CInt(objFmt.H / Twip_mm) & "mm", "") & _
            IIF(objFmt.纸向 = 1, "   纵向", "   横向")
    Else
        sta.Panels(2).Text = ""
    End If
End Sub

Private Sub ReFlashReport(Optional blnReload As Boolean = False)
'功能：重新刷新显示报表内容
'参数：blnReLoad=是否重新从数据库中加载数据
    Dim objTmp As Object, tmpReport As Report, intPreMax As Long
    
    If blnReload Then
        intPreMax = intMaxID
        Set tmpReport = ReadReport(lngRPTID, intMaxID)
        If tmpReport Is Nothing Then
            MsgBox "报表内容刷新失败！", vbInformation, App.Title
            intMaxID = intPreMax: Exit Sub
        End If
        Set objReport = tmpReport
        BlnSave = True
    End If
    
    For Each objTmp In lblSize
        If objTmp.Index <> 0 Then Unload objTmp
    Next
    For Each objTmp In lblLine
        If objTmp.Index <> 0 Then Unload objTmp
    Next
    For Each objTmp In lbl
        If objTmp.Index <> 0 Then Unload objTmp
    Next
    For Each objTmp In msh
        If objTmp.Index <> 0 Then Unload objTmp
    Next
    For Each objTmp In img
        If objTmp.Index <> 0 Then Unload objTmp
    Next
    For Each objTmp In ImgCode
        If objTmp.Index <> 0 Then Unload objTmp
    Next
    For Each objTmp In Chart
        If objTmp.Index <> 0 Then Unload objTmp
    Next
    For Each objTmp In pic
        If objTmp.Index <> 0 Then Unload objTmp
    Next
    For Each objTmp In Shp
        If objTmp.Index <> 0 Then Unload lblshp(objTmp.Index): Unload objTmp
    Next
    
    intCurID = 0
    Set objLastSel = Nothing
    
    '清空剪贴板
    If Me.Visible Then Set objClip = New RPTItems
    
    Call ShowReportDetail
    Call ShowAttrib
    
    Me.Refresh
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long    '工具条占用高度
    Dim staH As Long    '状态栏占用高度
    Dim attW As Long    '属性框相关部分占用宽度
    Dim sqlW As Long    '报表列表相关部分占用宽度
    Dim formatH As Long '格式高度
    Dim rulW As Long    '标尺宽度
    Dim rulH As Long    '标尺高度
    Dim i As Integer
    
    On Error Resume Next
    
    If WindowState = vbMinimized Then Exit Sub
    If blnMax Or WindowState = vbMaximized Then
        picSQL.Width = 3500
        picAtt.Width = 2400
        lvwPar.Height = 1700
        lblNote.Height = 900
        blnMax = False
    End If
    If WindowState = vbMaximized Then blnMax = True
    
    If Width < 8000 Then Width = 8000
    If Height < 5000 Then Height = 5000
    
    '靠齐控件宽度和高度
    cbrH = IIF(cbr.Visible, cbr.Height, 0)
    staH = IIF(sta.Visible, sta.Height, 0)
    attW = IIF(picR.Visible, picR.Width + picAtt.Width, 0)
    sqlW = IIF(picL.Visible, picL.Width + picSQL.Width, 0)
    formatH = picFormat.Height
    rulW = IIF(picRulerV.Visible, picRulerV.Width, 0)
    rulH = IIF(picRulerH.Visible, picRulerH.Height, 0)
    
    'zyb#Add
    picFormat.Top = ScaleTop + cbrH
    picFormat.Left = ScaleLeft + sqlW
    picFormat.Width = Me.ScaleWidth - picFormat.Left - attW
    
    'zyb#Add
    cmdDel.Left = picFormat.Width - cmdDel.Width - 15
    cmdAdd.Left = cmdDel.Left - cmdAdd.Width - 15
    If cmdAdd.Left - cboFormat.Left - 50 > 3000 Then
        cboFormat.Width = cmdAdd.Left - cboFormat.Left - 30
    End If
    
    picRulerV.Top = picFormat.Top + formatH + rulH  'zyb#Modify
    picRulerV.Left = ScaleLeft + sqlW
    picRulerV.Height = ScaleHeight - cbrH - staH - rulH - scrHsc.Height - formatH   'zyb#Modify
    
    picRulerH.Left = ScaleLeft + sqlW
    picRulerH.Top = picFormat.Top + formatH 'zyb#Modify
    picRulerH.Width = ScaleWidth - sqlW - attW - scrVsc.Width
    
    scrHsc.Top = picRulerV.Top + picRulerV.Height
    scrHsc.Left = picRulerV.Left + rulW
    scrHsc.Width = picRulerH.Width - rulW
    
    scrVsc.Top = picRulerV.Top
    scrVsc.Left = picRulerH.Left + picRulerH.Width
    scrVsc.Height = picRulerV.Height
    
    picBack.Left = picRulerV.Left + rulW
    picBack.Top = picRulerH.Top + rulH
    picBack.Width = scrHsc.Width
    picBack.Height = scrVsc.Height
    
    lblSQL.Top = 15: lblSQL.Left = 30
    lblSQL.Width = picSQL.ScaleWidth - 60
    
    tvwSQL.Left = picSQL.ScaleLeft
    tvwSQL.Top = lblSQL.Height + 30
    tvwSQL.Width = picSQL.ScaleWidth
    tvwSQL.Height = (ScaleHeight - staH - cbrH) - lblSQL.Height - lblPar.Height - lvwPar.Height - 60
    
    lblPar.Top = tvwSQL.Top + tvwSQL.Height + 15
    lblPar.Left = picSQL.ScaleLeft + 30
    lblPar.Width = lblSQL.Width
    
    lvwPar.Top = lblPar.Top + lblPar.Height + 15
    lvwPar.Left = picSQL.ScaleLeft + 15
    lvwPar.Width = tvwSQL.Width
    
    lblTool.Top = 15: lblTool.Left = 30
    lblTool.Width = picAtt.ScaleWidth - 60
    
    tbrTool.Top = lblTool.Top + lblTool.Height
    tbrTool.Left = lblTool.Left
    tbrTool.Width = lblTool.Width
    
    lblAtt.Top = tbrTool.Top + tbrTool.Height + 45
    lblAtt.Left = 30
    lblAtt.Width = lblTool.Width
    
    mshAtt.Top = lblAtt.Top + lblAtt.Height + 15
    mshAtt.Left = picAtt.ScaleLeft
    mshAtt.Width = picAtt.ScaleWidth
    mshAtt.Height = (ScaleHeight - cbrH - staH) - (lblTool.Height + 30) - (lblAtt.Height + 45) - lblNote.Height - picM.Height - tbrTool.Height
    
    picM.Left = picAtt.ScaleLeft
    picM.Top = mshAtt.Top + mshAtt.Height
    picM.Width = picAtt.ScaleWidth
    
    lblNote.Top = picM.Top + picM.Height
    lblNote.Left = picAtt.ScaleLeft
    lblNote.Width = picAtt.ScaleWidth
    
    Call ShowSize
    Call ShowScroll
    If Not scrHsc.Enabled Then DrawRuler picRulerH
    If Not scrVsc.Enabled Then DrawRuler picRulerV
    
    Call NoneEdit
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim intR As VbMsgBoxResult
    Dim strInfo As String
    
    If Not BlnSave Then
        intR = MsgBox("报表中当前修改内容尚未保存,要保存吗？", vbQuestion + vbYesNoCancel, App.Title)
        If intR = vbCancel Then '取消退出
            Cancel = 1: Exit Sub
        ElseIf intR = vbYes Then '退出前先保存
            strInfo = CheckData
            If strInfo <> "" Then MsgBox strInfo, vbInformation, App.Title: Cancel = 1: Exit Sub
            
            strInfo = CheckHead
            If strInfo <> "" Then MsgBox strInfo, vbInformation, App.Title: Cancel = 1: Exit Sub
            
            strInfo = CheckArea
            If strInfo <> "" Then MsgBox strInfo, vbInformation, App.Title: Cancel = 1: Exit Sub
            
            strInfo = CheckPars
            If strInfo <> "" Then MsgBox strInfo, vbInformation, App.Title: Exit Sub
    
            Call SelClear
            Refresh
            If Not SaveReport(lngRPTID, objReport, sta.Panels(2)) Then
                MsgBox "报表保存失败,请重试保存操作！", vbInformation, App.Title
                Cancel = 1: Exit Sub
            End If
            Call UpdatePriv
            
            BlnSave = True
            gblnModi = True
            Refresh
            
            If Not CheckReportPriv(lngRPTID) Then
                If MsgBox("你没有权限查询该报表某些数据源中的对象，虽然可以正常" & vbCrLf & _
                          "地保存，但在你修正这些问题之前你不能正常使用该报表！" & vbCrLf & _
                          "确实要退出设计环境吗？", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then
                    Cancel = 1: Exit Sub
                End If
            End If
        End If
    Else
        If Not CheckReportPriv(lngRPTID) Then
            If MsgBox("你没有权限查询该报表某些数据源中的对象，" & vbCrLf & _
                   "在你修正这些问题前你不能正常使用该报表！" & vbCrLf & _
                   "确实要退出设计环境吗？", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then
                Cancel = 1: Exit Sub
            End If
        End If
    End If
    
    lngRPTID = 0
    mblnNotModiData = False
    strMenu = ""
    Unload frmFlash
    
    SaveWinState Me, App.ProductName
End Sub

Private Sub lblshp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lngPreX = X: lngPreY = Y
    Set mobjMove = Nothing: mlngX = 0: mlngY = 0
    If Button = 1 Then
        If Shift = 2 Then
            If Mid(Shp(Index).Tag, 1, 2) = "" Then
                Call SelItem(Index, True) '加选
                If GetSelNum() = 1 Then
                    Call ShowAttrib(Index) '只选中一个则显示属性
                Else
                    Call ShowAttrib '多选时不显示属性
                End If
            Else
                Call SelItem(Index, False) '反选
                If GetSelNum() = 1 Then
                    Call ShowAttrib(intCurID) '只选中一个则显示属性(选中的不一定是该控件)
                Else
                    Call ShowAttrib '多选时不显示属性
                End If
            End If
        Else
            If Mid(Shp(Index).Tag, 1, 2) = "" Then
                Call SelClear
                Call SelItem(Index, True)
                Call ShowAttrib(Index) '只选中一个则显示属性
            End If
        End If
    ElseIf Button = 2 Then
        PopupMenu mnuFormat, 2
    End If
End Sub

Private Sub lblshp_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ObjSel As Shape

    If objReport.Items("_" & Index).系统 Then Exit Sub
    lblshp(Index).ZOrder 1
    DrawXY X + Shp(Index).Left, Y + Shp(Index).Top

    If Button = 1 And Mid(Shp(Index).Tag, 1, 2) <> "" Then
        If blnLock Then Exit Sub
        Call MoveSelect(X - lngPreX, Y - lngPreY)
        If objReport.Items("_" & Index).类型 = 10 Then Call DrawFrame(lblshp(Index))
        If GetSelNum() = 1 Then ShowAttrib Index
'    Else
'解决鼠标移动时，Shp范围内其他元素闪烁的问题
'        'zyb#Add
'        '允许在其上画报表元素
'        Set ObjSel = Shp(Index)
'
'        If X < 100 Or Y < 100 Or X > ObjSel.Width - 100 Or Y > ObjSel.Height - 100 Then
'            lblshp(Index).MousePointer = 99
'        Else
'            lblshp(Index).MousePointer = IIF(bytCurTool <> 0, 2, 0)
'            ObjSel.ZOrder 1
'            lblshp(Index).ZOrder 1
'            picPaper_MouseMove Button, Shift, X + Shp(Index).Left, Y + Shp(Index).Top
'        End If
    End If
End Sub

Private Sub lblshp_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    picPaper.Cls
    If Not mobjMove Is Nothing Then
        If Not mobjMove Is Shp(Index).Container Then
            If objReport.Items("_" & Index).参照 = "" Then
                If GetDataSouse(objReport.Items("_" & Index).内容) <> "" And UCase(mobjMove.name) = "PIC" Then
                    If objReport.Items("_" & mobjMove.Index).数据源 = "" Then
                        If MsgBox("当前卡片未绑定数据源，绑定后将分组打印多张卡片，数据源中存在""分组标识""字段则""分组标识""相同的为一组,否则一行数据为一组；" & vbCrLf & _
                             "不绑定则只打印一张卡片，是否绑定数据源""" & GetDataSouse(objReport.Items("_" & Index).内容) & """?", vbQuestion + vbYesNo + vbDefaultButton1, App.Title) = vbYes Then
                            objReport.Items("_" & mobjMove.Index).数据源 = GetDataSouse(objReport.Items("_" & Index).内容)
                        End If
                    End If
                End If
                Set Shp(Index).Container = mobjMove
                Shp(Index).Top = mlngY: Shp(Index).Left = mlngX
                Set lblshp(Index).Container = mobjMove
                lblshp(Index).Top = mlngY: lblshp(Index).Left = mlngX
                If UCase(mobjMove.name) = "PIC" Then
                    objReport.Items("_" & Index).父ID = mobjMove.Index
                Else
                    objReport.Items("_" & Index).父ID = 0
                End If
                objReport.Items("_" & Index).X = mlngX: objReport.Items("_" & Index).Y = mlngY
                Set mobjMove = Nothing: mlngX = 0: mlngY = 0
                Call ShowAttrib(Index)
            End If
        End If
    End If
End Sub

Private Sub Img_DblClick(Index As Integer)
    Dim ObjSel As Image, ObjLeft As Single, ObjTop As Single
    
    If GetSelNum <> 1 Then Exit Sub
    Set ObjSel = img(Index)
    With ObjSel
        ObjLeft = .Left
        ObjTop = .Top
        .Stretch = False
        .Stretch = True
        .Width = .Width * sgnMode
        .Height = .Height * sgnMode
        .Left = ObjLeft
        .Top = ObjTop
    End With
    
    With objReport.Items("_" & Index)
        .X = Format(ObjSel.Left / sgnMode, "0.00")
        .Y = Format(ObjSel.Top / sgnMode, "0.00")
        .H = Format(ObjSel.Height / sgnMode, "0.00")
        .W = Format(ObjSel.Width / sgnMode, "0.00")
    End With
    Call SelItem(Index, False)
    Call SelItem(Index, True)
End Sub

Private Sub Img_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not mobjMove Is Nothing Then
        If Not mobjMove Is img(Index).Container Then
            Set img(Index).Container = mobjMove
            img(Index).Top = mlngY: img(Index).Left = mlngX
            If UCase(mobjMove.name) = "PIC" Then
                objReport.Items("_" & Index).父ID = mobjMove.Index
            Else
                objReport.Items("_" & Index).父ID = 0
            End If
            objReport.Items("_" & Index).X = mlngX: objReport.Items("_" & Index).Y = mlngY
        End If
        Set mobjMove = Nothing: mlngX = 0: mlngY = 0
        Call ShowAttrib(Index)
    End If
End Sub

Private Sub ImgCode_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not mobjMove Is Nothing Then
        If Not mobjMove Is ImgCode(Index).Container Then
            If GetDataSouse(objReport.Items("_" & Index).内容) <> "" And UCase(mobjMove.name) = "PIC" Then
                If objReport.Items("_" & mobjMove.Index).数据源 = "" Then
                    If MsgBox("当前卡片未绑定数据源，绑定后将分组打印多张卡片，数据源中存在""分组标识""字段则""分组标识""相同的为一组,否则一行数据为一组；" & vbCrLf & _
                         "不绑定则只打印一张卡片，是否绑定数据源""" & GetDataSouse(objReport.Items("_" & Index).内容) & """?", vbQuestion + vbYesNo + vbDefaultButton1, App.Title) = vbYes Then
                        objReport.Items("_" & mobjMove.Index).数据源 = GetDataSouse(objReport.Items("_" & Index).内容)
                    End If
                End If
            End If
            Set ImgCode(Index).Container = mobjMove
            ImgCode(Index).Top = mlngY: ImgCode(Index).Left = mlngX
            If UCase(mobjMove.name) = "PIC" Then
                objReport.Items("_" & Index).父ID = mobjMove.Index
            Else
                objReport.Items("_" & Index).父ID = 0
            End If
            objReport.Items("_" & Index).X = mlngX: objReport.Items("_" & Index).Y = mlngY
        End If
        Set mobjMove = Nothing: mlngX = 0: mlngY = 0
        Call ShowAttrib(Index)
    End If
End Sub

Private Function GetDataSouse(ByVal str内容 As String) As String
'功能：根据内容获得检查数据源名称
    Dim i As Long, j As Long, k As Long
    
    i = InStr(str内容, "]")
    j = InStr(str内容, ".")
    k = InStr(str内容, "[")
    If i > k And i > j And i <> 0 And k <> 0 And j <> 0 Then
        GetDataSouse = Mid(str内容, k + 1, j - k - 1)
    End If
End Function

Private Sub lblLine_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not mobjMove Is Nothing Then
        If Not mobjMove Is lblLine(Index).Container Then
            Set lblLine(Index).Container = mobjMove
            lblLine(Index).Top = mlngY: lblLine(Index).Left = mlngX
            If UCase(mobjMove.name) = "PIC" Then
                objReport.Items("_" & Index).父ID = mobjMove.Index
            Else
                objReport.Items("_" & Index).父ID = 0
            End If
            objReport.Items("_" & Index).X = mlngX: objReport.Items("_" & Index).Y = mlngY
        End If
        Set mobjMove = Nothing: mlngX = 0: mlngY = 0
        Call ShowAttrib(Index)
    End If
End Sub

Private Sub LblSize_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not mobjMove Is Nothing Then
        If UCase(mobjMove.name) = "PIC" Then
            If mobjMove.Index <> objReport.Items("_" & intCurID).父ID Then
                objReport.Items("_" & intCurID).父ID = mobjMove.Index
                objReport.Items("_" & intCurID).X = mlngX: objReport.Items("_" & intCurID).Y = mlngY
                Call AdjustCoordinate(True)
            End If
        Else
            If objReport.Items("_" & intCurID).父ID <> 0 Then
                objReport.Items("_" & intCurID).父ID = 0
                objReport.Items("_" & intCurID).X = mlngX: objReport.Items("_" & intCurID).Y = mlngY
                Call AdjustCoordinate(True)
            End If
        End If
  
        Set mobjMove = Nothing: mlngX = 0: mlngY = 0
        Call ShowAttrib(intCurID)
    End If
End Sub

Private Sub mnuEdit_History_Click()
'功能：查看历史数据源
    Dim rsTmp As Recordset, strSQL As String
    Dim strKey As String, strInfo As String
    Dim strPreName As String, strDBName As String
    
    If tvwSQL.Nodes.count = 1 Then
        MsgBox "当前没有数据源！", vbInformation, App.Title: Exit Sub
    End If
    If tvwSQL.SelectedItem.Key = "Root" Then
        MsgBox "请选择要查看的数据源！", vbInformation, App.Title: Exit Sub
    End If
    
    If tvwSQL.SelectedItem.Parent.Key <> "Root" Then
        strKey = tvwSQL.SelectedItem.Parent.Key
    Else
        strKey = tvwSQL.SelectedItem.Key
    End If
    strPreName = objReport.Datas(strKey).名称
    strDBName = objReport.Datas(strKey).原名称
    
    On Error GoTo errH
    strSQL = "select 1 from zlRPTSQLsHistory Where 报表ID=[1] and 数据源名称=[2] And rownum<2"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngRPTID, IIF(strDBName = "", strPreName, strDBName))
    If rsTmp.RecordCount = 0 Then
        MsgBox "当前数据源没有历史记录！", vbInformation, App.Title: Exit Sub
    End If
    
    Call frmSQLEdit.ShowMe(Me, IIF(glngSys <> 0, glngSys, objReport.系统), objReport.Datas(strKey), objReport.Datas, 1)
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mshAtt_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim i As Integer, j As Integer, k As Integer, intRow As Integer
    Dim ItemThis As RPTItem, StrCompare As String
    Dim DataThis As RPTData
    Dim lngLeft As Long, lngTop As Long
    Dim StrPar As String
    
    Call NoneEdit
    
    '处理选中颜色
    mshAtt.Redraw = False
    mshAtt.Cell(flexcpBackColor, 1, 0, mshAtt.Rows - 1, 0) = mshAtt.BackColor
    mshAtt.Cell(flexcpForeColor, 1, 0, mshAtt.Rows - 1, 0) = mshAtt.ForeColor
    mshAtt.Cell(flexcpBackColor, 0, 0, 0, 1) = mshAtt.BackColorFixed
    mshAtt.Cell(flexcpForeColor, 0, 0, 0, 1) = mshAtt.ForeColorFixed
    

    mshAtt.Cell(flexcpBackColor, NewRow, 0) = mshAtt.BackColorSel
    mshAtt.Cell(flexcpForeColor, NewRow, 0) = mshAtt.ForeColorSel

    mshAtt.Redraw = True

    '处理注释
    Select Case mshAtt.TextMatrix(mshAtt.Row, 0)
        Case "类型"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "元素的类型,分为:线条,框线,标签,图片,表格,卡片."
        Case "名称"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "报表元素的名称，用于参照对象."
        Case "设置"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "对图表元素的格式，数据等内容进行设置"
        Case "内容"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "标签的文本内容或对应的数据项目."
        Case "X坐标"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "报表元素左上角的左边位置,以毫米为单位."
        Case "Y坐标"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "报表元素左上角的上边位置,以毫米为单位."
        Case "宽度"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "报表元素的输出宽度,以毫米为单位."
        Case "高度"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "报表元素的输出高度,以毫米为单位."
        Case "行高"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "表格中每一行的高度,以毫米为单位."
        Case "对齐"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "标签或表格列的文字在水平方向上的对齐方式."
        Case "表体网格色"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "表格表体的网格线条颜色."
        Case "表头网格色"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "表格表头的网格线条颜色."
        Case "网格色"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "表格的网格线条颜色."
        Case "前景色"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "报表元素的文字颜色或线条的颜色."
        Case "背景色"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "报表元素的背景颜色."
        Case "自动调整大小"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "设置标签、条码、图形的尺寸是否自动调整大小"
        Case "加粗"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "设置线条、框线的边框是否加粗"
        Case "表格线加粗"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "设置表格的网格线是否加粗"
        Case "形状"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "设置框线的边框形状"
        Case "字体"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "报表元素的文字字体相关属性."
        Case "自动字体"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "设置标签的内容过多时是否自动缩小字体尺寸进行打印."
        Case "边框"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "指定是否在标签文字背景的周围加一矩形框线."
        Case "分栏"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "设置任意表格的数据按几栏自动分列输出."
        Case "参照对象"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "设定标签与指定参照对象间的对齐关系."
        Case "方向"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "设定标签是表上项还是表下项."
        Case "性质"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "设定标签与参照对象间的参照关系."
        Case "格式"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "设置标签中数据字段的内容输出格式串,与VB格式字符兼容"
        Case "报表元素"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "当前报表格式中的所有报表元素"
        Case "输出图形"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "当前报表格式图形输出的模式"
        Case "票据"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "当前报表是否以票据的方式进行输出、打印"
        Case "禁止开始时间"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "当前报表不允许查询的时间段(开始时间)。"
        Case "禁止结束时间"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "当前报表不允许查询的时间段(结束时间)。"
        Case "空表打印"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "模块程序直接打印时,所有表格数据为空是否进行打印"
        Case "动态纸张"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "是否根据打印的内容自动调节纸张高度"
        Case "打印机"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "当前使用的打印机"
        Case "纸张"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "当前使用的纸张类型"
        Case "纸向"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "当前纸张的方向"
        Case "进纸方式"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "当前纸张的进纸方式"
        Case "边线"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "表格打印或预览时是否输出边线"
        Case "换行"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "表格数据输出时是否自动换行,但任意表头文字终始要换行"
        Case "保持比例"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "图象缩放时，是否保持原始的宽高比例"
        Case "报告图像"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "设置该图像是否用于影像检查报告"
        Case "条码类型"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "设置条码的类型"
        Case "条码线宽"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "设置条码线条的宽度(1-N)，它决定了条码的宽度"
        Case "显示数字"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "是否在条码图形下面显示条码的数字"
        Case "求校验和"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "是否对29码计算校验和"
        Case "旋转方向"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "条码图形输出时的旋转方向"
        Case "行距"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "换行输出的多行文字之间的点距(0-100)"
        Case "数据源"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "可选择当前任意一个任意表数据源，确定卡片动态打印的分页"
        Case "左右间距"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "分页动态打印卡片时，每张卡片的左右间距,以毫米为单位"
        Case "上下间距"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "分页动态打印卡片时，每张卡片的上下间距,以毫米为单位"
        Case "数据源行号"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "如果标签内容中绑定了数据源，控制标签显示第几行的数据"
        Case "横向分栏"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "分页动态打印卡片时，横向打印的卡片数，0为自适应"
        Case "纵向分栏"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "分页动态打印卡片时，纵向打印的卡片数，0为自适应"
        Case "容器"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "可选择当前元素所在区域的容器，元素的范围必须在容器(卡片)之内才允许设置"
        Case "关联报表"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "关联后查询此报表时，双击此元素将根据设置的参数来源作为关联报表的参数来执行关联报表。"
        Case "水平反转"
            lblNote.Caption = mshAtt.TextMatrix(mshAtt.Row, 0) & vbCrLf & "签名元素在预览、打印时水平反转。"
        Case Else
            lblNote.Caption = ""
    End Select
    
    If blnLock Then Exit Sub
    '如果是系统固有项目，则只允许设置字体及参照对象
    If InStr(1, "字体,参照对象,方向,性质", mshAtt.TextMatrix(mshAtt.Row, 0)) = 0 And intCurID <> 0 Then
        If objReport.Items("_" & intCurID).系统 Then Exit Sub
    End If
    
    '如果是左联接表格的子表,则不允许设置高度及行高
    If InStr(1, "高度,行高,字体,背景色,网格色", mshAtt.TextMatrix(mshAtt.Row, 0)) <> 0 And intCurID <> 0 Then
        If objReport.Items("_" & intCurID).参照 <> "" And objReport.Items("_" & intCurID).类型 = 5 Then Exit Sub
    End If
    
    '如果是图型字段，且未设置内容，则允许设置；否则退出
    If mshAtt.TextMatrix(mshAtt.Row, 0) = "内容" Then
        If objReport.Items("_" & intCurID).类型 = 11 Then
            cmdAtt.Top = mshAtt.Top + mshAtt.CellTop
            cmdAtt.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1) + mshAtt.CellWidth - cmdAtt.Width
            cmdAtt.Visible = True
            Exit Sub
        ElseIf objReport.Items("_" & intCurID).类型 = 2 Then
            Dim strNodeName As String, NodeThis As Node
            '如果是adLongVarBinary型字段,则不允许修改
            i = InStr(1, mshAtt.TextMatrix(mshAtt.Row, 1), "]")
            j = InStr(1, mshAtt.TextMatrix(mshAtt.Row, 1), ".")
            k = InStr(1, mshAtt.TextMatrix(mshAtt.Row, 1), "[")
            If i > k And i > j And i <> 0 And k <> 0 Then
                strNodeName = Mid(mshAtt.TextMatrix(mshAtt.Row, 1), j + 1, i - j - 1)
                
                For Each NodeThis In tvwSQL.Nodes
                    If mdlPublic.GetStdNodeText(NodeThis.Text) = strNodeName And IsType(Val(NodeThis.Tag), adLongVarBinary) Then Exit Sub
                Next
            End If
        End If
    End If
    
    mshAtt.ColWidth(1) = mshAtt.Cell(flexcpWidth, 0, 1)
    
    '编辑处理
    Select Case mshAtt.TextMatrix(mshAtt.Row, 0)
        Case "名称", "内容", "X坐标", "Y坐标", "宽度", "高度", "行高", "分栏", "字号", "格式", "行距" _
            , "左右间距", "上下间距", "数据源行号", "横向分栏", "纵向分栏"
            If intCurID = 0 Then Exit Sub
            If objReport.Items("_" & intCurID).类型 <> 2 Or mshAtt.TextMatrix(mshAtt.Row, 0) <> "内容" Then
                txtAtt.MaxLength = 0
                If InStr("X坐标,Y坐标,宽度,高度,行高,左右间距,上下间距,源行号,横向分栏,纵向分栏", mshAtt.TextMatrix(mshAtt.Row, 0)) > 0 Then
                    txtAtt.MaxLength = 7
                ElseIf mshAtt.TextMatrix(mshAtt.Row, 0) = "分栏" Then
                    '不是主表,则允许设置分栏
                    If objReport.Items("_" & intCurID).参照 <> "" Then Exit Sub
                    For Each ItemThis In objReport.Items
                        If ItemThis.格式号 = mbytCurrFmt And InStr(1, "4,5", ItemThis.类型) <> 0 Then
                            If ItemThis.格式号 = mbytCurrFmt And ItemThis.Key <> intCurID And ItemThis.参照 = objReport.Items("_" & intCurID).名称 And InStr(1, "4,5", ItemThis.类型) <> 0 Then Exit Sub
                        End If
                    Next
                    txtAtt.MaxLength = 2
                ElseIf mshAtt.TextMatrix(mshAtt.Row, 0) = "字号" Then
                    txtAtt.MaxLength = 7
                ElseIf mshAtt.TextMatrix(mshAtt.Row, 0) = "格式" Then
                    txtAtt.MaxLength = 50
                ElseIf mshAtt.TextMatrix(mshAtt.Row, 0) = "行距" Then
                    txtAtt.MaxLength = 3
                End If
                If InStr("X坐标,Y坐标", mshAtt.TextMatrix(mshAtt.Row, 0)) > 0 And objReport.Items("_" & intCurID).参照 <> "" Then
                    If Not (mshAtt.TextMatrix(mshAtt.Row, 0) = "Y坐标" And objReport.Items("_" & intCurID).类型 = 2) Then Exit Sub
                End If
                
                txtAtt.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1) + 30
                txtAtt.Top = mshAtt.Top + mshAtt.CellTop + (mshAtt.CellHeight - txtAtt.Height) / 2
                txtAtt.Width = mshAtt.ColWidth(1) - 60
                txtAtt.Text = mshAtt.TextMatrix(mshAtt.Row, 1)
                txtAtt.Visible = True: txtAtt.SetFocus
            Else
                '标签允许选择常用的内容
                cboText.Clear
                For i = 1 To objReport.Datas.count
                    For j = 1 To objReport.Datas(i).Pars.count
                        If InStr(StrPar, objReport.Datas(i).Pars(j).名称) = 0 Then
                            StrPar = StrPar & "|" & objReport.Datas(i).Pars(j).名称
                        End If
                    Next
                Next
                StrPar = Mid(StrPar, 2)
                If StrPar <> "" Then
                    For i = 0 To UBound(Split(StrPar, "|"))
                        cboAtt.AddItem "[=" & Split(StrPar, "|")(i) & "]"
                    Next
                End If
                cboText.AddItem "[操作员姓名]"
                cboText.AddItem "[操作员编号]"
                cboText.AddItem "[单位名称]"
                cboText.AddItem "[页号]"
                cboText.AddItem "[页数]"
                cboText.AddItem "[yyyy-mm-dd]"
                cboText.AddItem "[yyyy-mm-dd HH:MM]"
                cboText.AddItem "[yyyy-mm-dd HH:MM:SS]"
                
                cboText.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1)
                cboText.Top = mshAtt.Top + mshAtt.CellTop
                cboText.Width = mshAtt.ColWidth(1) - 60
                cboText.Text = objReport.Items("_" & intCurID).内容
                cboText.Visible = True:  cboText.SetFocus
            End If
        Case "对齐"
            cboAtt.Clear
            cboAtt.AddItem "左对齐"
            cboAtt.AddItem "中对齐"
            cboAtt.AddItem "右对齐"
            cboAtt.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1)
            cboAtt.Top = mshAtt.Top + mshAtt.CellTop
            cboAtt.Width = mshAtt.ColWidth(1)
            cboAtt.ListIndex = objReport.Items("_" & intCurID).对齐
            cboAtt.Visible = True:  cboAtt.SetFocus
        Case "形状"
            cboAtt.Clear
            cboAtt.AddItem "方形"
            cboAtt.AddItem "圆形"
            cboAtt.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1)
            cboAtt.Top = mshAtt.Top + mshAtt.CellTop
            cboAtt.Width = mshAtt.ColWidth(1)
            cboAtt.ListIndex = IIF(objReport.Items("_" & intCurID).边框, 1, 0)
            cboAtt.Visible = True:  cboAtt.SetFocus
        Case "前景色", "背景色", "字体", "表体网格色", "设置", "网格色", "表头网格色"
            cmdAtt.Top = mshAtt.Top + mshAtt.CellTop
            cmdAtt.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1) + mshAtt.ColWidth(1) - cmdAtt.Width
            cmdAtt.Visible = True
        Case "关联报表"
            cmdAtt.Top = mshAtt.Top + mshAtt.CellTop
            cmdAtt.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1) + mshAtt.ColWidth(1) - cmdAtt.Width
            cmdAtt.Visible = True
            txtAtt.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1) + 30
            txtAtt.Top = mshAtt.Top + mshAtt.CellTop + (mshAtt.CellHeight - txtAtt.Height) / 2
            txtAtt.Width = mshAtt.ColWidth(1) - 60
            txtAtt.Text = mshAtt.TextMatrix(mshAtt.Row, 1)
            If mshAtt.TextMatrix(mshAtt.Row, 1) <> "" Then
                txtAtt.Visible = False
            Else
                txtAtt.Visible = True: txtAtt.SetFocus
            End If
        Case "报表元素"
            Call GetAllElement
            '调整位置
            cboAtt.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1)
            cboAtt.Top = mshAtt.Top + mshAtt.CellTop
            cboAtt.Width = mshAtt.ColWidth(1)
            cboAtt.Visible = True:  cboAtt.SetFocus
        Case "禁止开始时间"
            dtpAtt.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1)
            dtpAtt.Top = mshAtt.Top + mshAtt.CellTop
            dtpAtt.Width = mshAtt.ColWidth(1)
            dtpAtt.Value = Format(mshAtt.TextMatrix(mshAtt.Row, 1), "HH:mm:ss")
            dtpAtt.Visible = True:  dtpAtt.SetFocus
        Case "禁止结束时间"
            dtpAtt.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1)
            dtpAtt.Top = mshAtt.Top + mshAtt.CellTop
            dtpAtt.Width = mshAtt.ColWidth(1)
            dtpAtt.Value = Format(mshAtt.TextMatrix(mshAtt.Row, 1), "HH:mm:ss")
            dtpAtt.Visible = True:  dtpAtt.SetFocus
        Case "输出图形"
            blnModify = False
            Call LoadOutChart
            
            '调整位置
            cboAtt.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1)
            cboAtt.Top = mshAtt.Top + mshAtt.CellTop
            cboAtt.Width = mshAtt.ColWidth(1)
            cboAtt.Visible = True:  cboAtt.SetFocus
            
            Call LocateOutChart
            blnModify = True
        Case "参照对象"
            '如果本表已分栏,则退出
            If objReport.Items("_" & intCurID).分栏 > 1 Then Exit Sub
            '已有表格附加于本表,则本表不能设置参照对象及分栏
            For Each ItemThis In objReport.Items
                If ItemThis.格式号 = mbytCurrFmt And InStr(1, "4,5", ItemThis.类型) <> 0 Then
                    If ItemThis.Key <> intCurID And ItemThis.参照 = objReport.Items("_" & intCurID).名称 Then Exit Sub
                End If
            Next

            cboAtt.Clear
            cboAtt.AddItem ""
            
            '填充独立表格
            For Each ItemThis In objReport.Items
                If ItemThis.格式号 = mbytCurrFmt Then
                    Select Case objReport.Items("_" & intCurID).类型
                    Case "2"
                        If InStr(1, "|4,|5,", "|" & ItemThis.类型 & ",") <> 0 And ItemThis.参照 = "" Then
                            cboAtt.AddItem ItemThis.名称
                        End If
                    Case "4", "5"
                        If ItemThis.分栏 < 2 And CheckTableProperty(ItemThis) Then cboAtt.AddItem ItemThis.名称
                    End Select
                End If
            Next
            
            '调整位置
            cboAtt.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1)
            cboAtt.Top = mshAtt.Top + mshAtt.CellTop
            cboAtt.Width = mshAtt.ColWidth(1)
            cboAtt.Visible = True:  cboAtt.SetFocus
            
            StrCompare = objReport.Items("_" & intCurID).参照
            For i = 0 To cboAtt.ListCount - 1
                If cboAtt.List(i) = StrCompare Then
                    cboAtt.ListIndex = i
                    mshAtt.TextMatrix(mshAtt.Row, 1) = StrCompare
                    Exit For
                End If
            Next
            If cboAtt.Text <> StrCompare Then cboAtt.ListIndex = 0
        Case "数据源"

            cboAtt.Clear
            cboAtt.AddItem ""
            
            '填充独立表格
            For Each DataThis In objReport.Datas
                If DataThis.类型 = 0 Then
                    cboAtt.AddItem DataThis.名称
                End If
            Next
            
            '调整位置
            cboAtt.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1)
            cboAtt.Top = mshAtt.Top + mshAtt.CellTop
            cboAtt.Width = mshAtt.ColWidth(1)
            cboAtt.Visible = True:  cboAtt.SetFocus
            
            StrCompare = objReport.Items("_" & intCurID).数据源
            For i = 0 To cboAtt.ListCount - 1
                If cboAtt.List(i) = StrCompare Then
                    cboAtt.ListIndex = i
                    mshAtt.TextMatrix(mshAtt.Row, 1) = StrCompare
                    Exit For
                End If
            Next
            If cboAtt.Text <> StrCompare Then cboAtt.ListIndex = 0
        Case "容器"
            cboAtt.Clear
            cboAtt.AddItem "页面"
            
            '填充独立表格
            If objReport.Items("_" & intCurID).父ID <> 0 Then
                lngLeft = objReport.Items("_" & objReport.Items("_" & intCurID).父ID).X
                lngTop = objReport.Items("_" & objReport.Items("_" & intCurID).父ID).Y
            End If
            For Each ItemThis In objReport.Items
                If ItemThis.类型 = 14 And ItemThis.格式号 = mbytCurrFmt Then
                    If objReport.Items("_" & intCurID).Y + lngTop >= ItemThis.Y And objReport.Items("_" & intCurID).X + lngLeft >= ItemThis.X And _
                            objReport.Items("_" & intCurID).H + objReport.Items("_" & intCurID).Y + lngTop <= ItemThis.Y + ItemThis.H And _
                            objReport.Items("_" & intCurID).W + objReport.Items("_" & intCurID).X + lngLeft <= ItemThis.X + ItemThis.W Then
                        cboAtt.AddItem ItemThis.名称
                        cboAtt.ItemData(cboAtt.NewIndex) = ItemThis.id
                    ElseIf objReport.Items("_" & intCurID).父ID = ItemThis.id Then
                        cboAtt.AddItem ItemThis.名称
                        cboAtt.ItemData(cboAtt.NewIndex) = ItemThis.id
                    End If
                End If
            Next
            
            '调整位置
            cboAtt.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1)
            cboAtt.Top = mshAtt.Top + mshAtt.CellTop
            cboAtt.Width = mshAtt.ColWidth(1)
            cboAtt.Visible = True:  cboAtt.SetFocus
            
            If objReport.Items("_" & intCurID).父ID = 0 Then
                StrCompare = "页面"
            Else
                StrCompare = objReport.Items("_" & objReport.Items("_" & intCurID).父ID).名称
            End If
            For i = 0 To cboAtt.ListCount - 1
                If cboAtt.List(i) = StrCompare Then
                    cboAtt.ListIndex = i
                    mshAtt.TextMatrix(mshAtt.Row, 1) = StrCompare
                    Exit For
                End If
            Next
            If cboAtt.Text <> StrCompare Then cboAtt.ListIndex = 0
        Case "方向"
            If mshAtt.TextMatrix(GetRow("参照对象"), 1) = "" Then
                mshAtt.TextMatrix(GetRow("方向"), 1) = "独立"
                Exit Sub
            End If
            
            cboAtt.Clear
            cboAtt.AddItem "表上项"
            cboAtt.AddItem "表下项"
            
            '调整位置
            cboAtt.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1)
            cboAtt.Top = mshAtt.Top + mshAtt.CellTop
            cboAtt.Width = mshAtt.ColWidth(1)
            cboAtt.Visible = True:  cboAtt.SetFocus
            
            StrCompare = Mid(objReport.Items("_" & intCurID).性质, 1, 1)
            StrCompare = IIF(StrCompare = "2", "表下项", "表上项")
            For i = 0 To cboAtt.ListCount - 1
                If cboAtt.List(i) = StrCompare Then
                    cboAtt.ListIndex = i
                    mshAtt.TextMatrix(mshAtt.Row, 1) = StrCompare
                    Exit For
                End If
            Next
            If cboAtt.Text <> StrCompare Then cboAtt.ListIndex = 0
        Case "性质"
            If mshAtt.TextMatrix(GetRow("参照对象"), 1) = "" Then
                mshAtt.TextMatrix(GetRow("性质"), 1) = "独立"
                Exit Sub
            End If
            
            cboAtt.Clear
            Select Case objReport.Items("_" & intCurID).类型
            Case 2
                cboAtt.AddItem "靠左"
                cboAtt.AddItem "靠中"
                cboAtt.AddItem "靠右"
            Case 4
                cboAtt.AddItem "附加"
            Case 5
                cboAtt.AddItem "左联接"
            End Select
            
            '调整位置
            cboAtt.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1)
            cboAtt.Top = mshAtt.Top + mshAtt.CellTop
            cboAtt.Width = mshAtt.ColWidth(1)
            cboAtt.Visible = True:  cboAtt.SetFocus
            
            Select Case objReport.Items("_" & intCurID).类型
            Case 2
                StrCompare = Mid(objReport.Items("_" & intCurID).性质, 2)
                StrCompare = IIF(StrCompare = "1", "靠左", IIF(StrCompare = "2", "靠中", "靠右"))
                For i = 0 To cboAtt.ListCount - 1
                    If cboAtt.List(i) = StrCompare Then
                        cboAtt.ListIndex = i
                        mshAtt.TextMatrix(mshAtt.Row, 1) = StrCompare
                        Exit For
                    End If
                Next
                If cboAtt.Text <> StrCompare Then cboAtt.ListIndex = 0
            Case 4, 5
                StrCompare = IIF(objReport.Items("_" & intCurID).性质 = "1", "附加", "左联接")
                For i = 0 To cboAtt.ListCount - 1
                    If cboAtt.List(i) = StrCompare Then
                        cboAtt.ListIndex = i
                        mshAtt.TextMatrix(mshAtt.Row, 1) = StrCompare
                        Exit For
                    End If
                Next
                If cboAtt.Text <> StrCompare Then cboAtt.ListIndex = 0
            End Select
        Case "条码类型"
            cboAtt.Clear
            cboAtt.AddItem "Code 128(遗留)": cboAtt.ItemData(cboAtt.NewIndex) = 1
            cboAtt.AddItem "Code 128 Auto": cboAtt.ItemData(cboAtt.NewIndex) = 3
            cboAtt.AddItem "Code 39": cboAtt.ItemData(cboAtt.NewIndex) = 2
            cboAtt.AddItem "QR Code": cboAtt.ItemData(cboAtt.NewIndex) = 10
            For i = 0 To cboAtt.ListCount - 1
                If cboAtt.ItemData(i) = objReport.Items("_" & intCurID).序号 Then
                    CboSetIndex cboAtt.hwnd, i: Exit For
                End If
            Next
            If cboAtt.ListIndex = -1 Then CboSetIndex cboAtt.hwnd, 0
            
            cboAtt.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1)
            cboAtt.Top = mshAtt.Top + mshAtt.CellTop
            cboAtt.Width = mshAtt.ColWidth(1)
            cboAtt.Visible = True:  cboAtt.SetFocus
        Case "条码线宽"
            cboAtt.Clear
            For i = 1 To 10
                cboAtt.AddItem i
            Next
            If objReport.Items("_" & intCurID).行高 <= 0 Then
                CboSetIndex cboAtt.hwnd, 1
            Else
                CboSetIndex cboAtt.hwnd, objReport.Items("_" & intCurID).行高 - 1
            End If
            
            cboAtt.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1)
            cboAtt.Top = mshAtt.Top + mshAtt.CellTop
            cboAtt.Width = mshAtt.ColWidth(1)
            cboAtt.Visible = True:  cboAtt.SetFocus
        Case "旋转方向"
            cboAtt.Clear
            cboAtt.AddItem "不旋转"
            cboAtt.AddItem "顺时针90度"
            cboAtt.AddItem "逆时针90度"
            CboSetIndex cboAtt.hwnd, Val(Mid(objReport.Items("_" & intCurID).表头, 3, 1))
            
            cboAtt.Left = mshAtt.Left + mshAtt.Cell(flexcpLeft, NewRow, 1)
            cboAtt.Top = mshAtt.Top + mshAtt.CellTop
            cboAtt.Width = mshAtt.ColWidth(1)
            cboAtt.Visible = True:  cboAtt.SetFocus
    End Select
End Sub

Private Sub mshAtt_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    txtAtt.Visible = False
    cmdAtt.Visible = False
    cboAtt.Visible = False
    mshAtt.SetFocus
End Sub

Private Sub pic_DblClick(Index As Integer)
    Dim ObjSel As PictureBox, ObjLeft As Single, ObjTop As Single
    
    If GetSelNum <> 1 Then Exit Sub
    Set ObjSel = pic(Index)
    With ObjSel
        ObjLeft = .Left
        ObjTop = .Top
        .Width = .Width * sgnMode
        .Height = .Height * sgnMode
        .Left = ObjLeft
        .Top = ObjTop
    End With
    
    With objReport.Items("_" & Index)
        .X = Format(ObjSel.Left / sgnMode, "0.00")
        .Y = Format(ObjSel.Top / sgnMode, "0.00")
        .H = Format(ObjSel.Height / sgnMode, "0.00")
        .W = Format(ObjSel.Width / sgnMode, "0.00")
    End With
    Call SelItem(Index, False)
    Call SelItem(Index, True)
End Sub

Private Sub Img_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lngPreX = X: lngPreY = Y
    Set mobjMove = Nothing: mlngX = 0: mlngY = 0
    If Button = 1 Then
        If Shift = 2 Then
            If Mid(img(Index).Tag, 1, 2) = "" Then
                Call SelItem(Index, True) '加选
                If GetSelNum() = 1 Then
                    Call ShowAttrib(Index) '只选中一个则显示属性
                Else
                    Call ShowAttrib '多选时不显示属性
                End If
            Else
                Call SelItem(Index, False) '反选
                If GetSelNum() = 1 Then
                    Call ShowAttrib(intCurID) '只选中一个则显示属性(选中的不一定是该控件)
                Else
                    Call ShowAttrib '多选时不显示属性
                End If
            End If
        Else
            If Mid(img(Index).Tag, 1, 2) = "" Then
                Call SelClear
                Call SelItem(Index, True)
                Call ShowAttrib(Index) '只选中一个则显示属性
            End If
        End If
    ElseIf Button = 2 Then
        PopupMenu mnuFormat, 2
    End If
End Sub

Private Sub pic_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    If UCase(Source.name) = "TVWSQL" Then
        selArea.Left = X: selArea.Top = Y
        Call AddReportItem(True, pic(Index))
        BlnSave = False
    End If

End Sub

Private Sub pic_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
    DrawXY CLng(X), CLng(Y)
    If UCase(Source.name) = "TVWSQL" Then
        If State = 1 Then
            Set tvwSQL.DragIcon = lvwPar.DragIcon
        ElseIf State = 0 Then
            If tvwSQL.SelectedItem.Children = 0 Then
                Set tvwSQL.DragIcon = scrHsc.DragIcon
            Else
                Set tvwSQL.DragIcon = scrVsc.DragIcon
            End If
        End If
    End If
End Sub

Private Sub pic_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lngPreX = X: lngPreY = Y
    If Button = 1 Then
        selArea.Left = X
        selArea.Top = Y
        blnDown = True
        If Shift = 2 Then
            If Mid(pic(Index).Tag, 1, 2) = "" Then
                Call SelItem(Index, True) '加选
                If GetSelNum() = 1 Then
                    Call ShowAttrib(Index) '只选中一个则显示属性
                Else
                    Call ShowAttrib '多选时不显示属性
                End If
            Else
                Call SelItem(Index, False) '反选
                If GetSelNum() = 1 Then
                    Call ShowAttrib(intCurID) '只选中一个则显示属性(选中的不一定是该控件)
                Else
                    Call ShowAttrib '多选时不显示属性
                End If
            End If
        Else
            If Mid(pic(Index).Tag, 1, 2) = "" Then
                Call SelClear
                Call SelItem(Index, True)
                Call ShowAttrib(Index) '只选中一个则显示属性
            End If
        End If
    ElseIf Button = 2 Then
        PopupMenu mnuFormat, 2
    End If
End Sub

Private Sub Img_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ObjSel As Image
    
    If objReport.Items("_" & Index).系统 Then Exit Sub
    DrawXY X + img(Index).Left, Y + img(Index).Top
    Set ObjSel = img(Index)
    
    If Button = 1 And Mid(ObjSel.Tag, 1, 2) <> "" Then
        If blnLock Then Exit Sub
        Call MoveSelect(X - lngPreX, Y - lngPreY)
        If GetSelNum() = 1 Then ShowAttrib Index
    Else
        ObjSel.MousePointer = 99
    End If
End Sub

Private Sub pic_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ObjSel As PictureBox
    Static PreX As Long, PreY As Long
    
    If bytCurTool <> 0 Then
        Call DrawXY(CLng(X), CLng(Y))
        
        '画选择虚框
        If Button = 1 And blnDown And Shift <> 4 Then
            If PreX = Empty And PreY = Empty Then
                PreX = selArea.Left
                PreY = selArea.Top
            End If
            If bytCurTool <> 1 Then
                pic(Index).Line (selArea.Left, selArea.Top)-(PreX, PreY), picPaper.BackColor, B
                pic(Index).Line (selArea.Left, selArea.Top)-(X, Y), , B
            Else
                If Abs(X - selArea.Left) >= Abs(Y - selArea.Top) Then
                    '画横线
                    If bytLine = 2 Then pic(Index).Cls
                    pic(Index).Line (selArea.Left, selArea.Top)-(PreX, selArea.Top), picPaper.BackColor
                    pic(Index).Line (selArea.Left, selArea.Top)-(X, selArea.Top)
                    bytLine = 1
                Else
                    '画竖线
                    If bytLine = 1 Then pic(Index).Cls
                    pic(Index).Line (selArea.Left, selArea.Top)-(selArea.Left, PreY), picPaper.BackColor
                    pic(Index).Line (selArea.Left, selArea.Top)-(selArea.Left, Y)
                    bytLine = 2
                End If
            End If
            PreX = X: PreY = Y
        End If
    Else
        If objReport.Items("_" & Index).系统 Then Exit Sub
        DrawXY X + pic(Index).Left, Y + pic(Index).Top
        Set ObjSel = pic(Index)
        If Button = 1 And Mid(ObjSel.Tag, 1, 2) <> "" Then
            If blnLock Then Exit Sub
            Call MoveSelect(X - lngPreX, Y - lngPreY)
            If GetSelNum() = 1 Then ShowAttrib Index
        Else
            ObjSel.MousePointer = 99
        End If
    End If
End Sub

Private Sub ImgCode_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lngPreX = X: lngPreY = Y
    Set mobjMove = Nothing: mlngX = 0: mlngY = 0
    If Button = 1 Then
        If Shift = 2 Then
            If Mid(ImgCode(Index).Tag, 1, 2) = "" Then
                Call SelItem(Index, True) '加选
                If GetSelNum() = 1 Then
                    Call ShowAttrib(Index) '只选中一个则显示属性
                Else
                    Call ShowAttrib '多选时不显示属性
                End If
            Else
                Call SelItem(Index, False) '反选
                If GetSelNum() = 1 Then
                    Call ShowAttrib(intCurID) '只选中一个则显示属性(选中的不一定是该控件)
                Else
                    Call ShowAttrib '多选时不显示属性
                End If
            End If
        Else
            If Mid(ImgCode(Index).Tag, 1, 2) = "" Then
                Call SelClear
                Call SelItem(Index, True)
                Call ShowAttrib(Index) '只选中一个则显示属性
            End If
        End If
    ElseIf Button = 2 Then
        PopupMenu mnuFormat, 2
    End If
End Sub

Private Sub ImgCode_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ObjSel As Image
    
    If objReport.Items("_" & Index).系统 Then Exit Sub
    DrawXY X + ImgCode(Index).Left, Y + ImgCode(Index).Top
    Set ObjSel = ImgCode(Index)
    If Button = 1 And Mid(ObjSel.Tag, 1, 2) <> "" Then
        If blnLock Then Exit Sub
        Call MoveSelect(X - lngPreX, Y - lngPreY)
        If GetSelNum() = 1 Then ShowAttrib Index
    Else
        ObjSel.MousePointer = 99
    End If
End Sub

Private Sub lbl_DblClick(Index As Integer)
'功能：自动调整标签的大小
    If Not blnLock And GetSelNum = 1 Then
        If objReport.Items("_" & Index).类型 = 10 Then Exit Sub
        lbl(Index).AutoSize = True
        lbl(Index).AutoSize = False
        objReport.Items("_" & Index).W = Format(lbl(Index).Width / sgnMode, "0.00")
        objReport.Items("_" & Index).H = Format(lbl(Index).Height / sgnMode, "0.00")
        SeekItem lbl(Index), lbl(Index).Left, lbl(Index).Top
        Call ShowAttrib(Index)
        Call ReferTo
        BlnSave = False
    End If
End Sub

Private Sub lbl_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lngPreX = X: lngPreY = Y
    Set mobjMove = Nothing: mlngX = 0: mlngY = 0
    If Button = 1 Then
        If lbl(Index).MousePointer <> 99 Then 'zyb#Modify
            'zyb#Add
            picPaper_MouseDown Button, Shift, X + lbl(Index).Left, Y + lbl(Index).Top
        Else 'zyb#Modify
            If Shift = 2 Then
                If Mid(lbl(Index).Tag, 1, 2) = "" Then
                    Call SelItem(Index, True) '加选
                    If GetSelNum() = 1 Then
                        Call ShowAttrib(Index) '只选中一个则显示属性
                    Else
                        Call ShowAttrib '多选时不显示属性
                    End If
                Else
                    Call SelItem(Index, False) '反选
                    If GetSelNum() = 1 Then
                        Call ShowAttrib(intCurID) '只选中一个则显示属性(选中的不一定是该控件)
                    Else
                        Call ShowAttrib '多选时不显示属性
                    End If
                End If
            Else
                If Mid(lbl(Index).Tag, 1, 2) = "" Then
                    Call SelClear
                    Call SelItem(Index, True)
                    Call ShowAttrib(Index) '只选中一个则显示属性
                End If
            End If
        End If 'zyb#Modify
    ElseIf Button = 2 Then
        PopupMenu mnuFormat, 2
    End If
End Sub

Private Sub lbl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ObjSel As Label

    If objReport.Items("_" & Index).系统 Then Exit Sub
    DrawXY X + lbl(Index).Left, Y + lbl(Index).Top
    
    If Button = 1 And Mid(lbl(Index).Tag, 1, 2) <> "" Then
        If blnLock Then Exit Sub
        Call MoveSelect(X - lngPreX, Y - lngPreY)
        If objReport.Items("_" & Index).类型 = 10 Then Call DrawFrame(lbl(Index))
        If GetSelNum() = 1 Then ShowAttrib Index
    Else
        'zyb#Add
        '允许在其上画报表元素
        If objReport.Items("_" & Index).类型 = 10 Then
            Set ObjSel = lbl(Index)
            
            If X < 100 Or Y < 100 Or X > ObjSel.Width - 100 Or Y > ObjSel.Height - 100 Then
                ObjSel.MousePointer = 99
            Else
                ObjSel.MousePointer = IIF(bytCurTool <> 0, 2, 0)
                ObjSel.ZOrder 1
                picPaper_MouseMove Button, Shift, X + lbl(Index).Left, Y + lbl(Index).Top
            End If
        End If
    End If
End Sub

Private Sub lbl_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lbl(Index).MousePointer <> 99 Then
        picPaper_MouseUp Button, Shift, X + lbl(Index).Left, Y + lbl(Index).Top
    End If
    picPaper.Cls
    If Not mobjMove Is Nothing Then
        If Not mobjMove Is lbl(Index).Container Then
            If objReport.Items("_" & Index).参照 = "" Then
                If GetDataSouse(objReport.Items("_" & Index).内容) <> "" And UCase(mobjMove.name) = "PIC" Then
                    If objReport.Items("_" & mobjMove.Index).数据源 = "" Then
                        If MsgBox("当前卡片未绑定数据源，绑定后将分组打印多张卡片，数据源中存在""分组标识""字段则""分组标识""相同的为一组,否则一行数据为一组；" & vbCrLf & _
                             "不绑定则只打印一张卡片，是否绑定数据源""" & GetDataSouse(objReport.Items("_" & Index).内容) & """?", vbQuestion + vbYesNo + vbDefaultButton1, App.Title) = vbYes Then
                            objReport.Items("_" & mobjMove.Index).数据源 = GetDataSouse(objReport.Items("_" & Index).内容)
                        End If
                    End If
                End If
                Set lbl(Index).Container = mobjMove
                lbl(Index).Top = mlngY: lbl(Index).Left = mlngX
                If UCase(mobjMove.name) = "PIC" Then
                    objReport.Items("_" & Index).父ID = mobjMove.Index
                Else
                    objReport.Items("_" & Index).父ID = 0
                End If
                objReport.Items("_" & Index).X = mlngX: objReport.Items("_" & Index).Y = mlngY
                Set mobjMove = Nothing: mlngX = 0: mlngY = 0
                Call ShowAttrib(Index)
            End If
        End If
    End If
End Sub

Private Sub lblLine_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lngPreX = X: lngPreY = Y
    Set mobjMove = Nothing: mlngX = 0: mlngY = 0
    If Button = 1 Then
        If Shift = 2 Then
            If Mid(lblLine(Index).Tag, 1, 2) = "" Then
                Call SelItem(Index, True) '加选
                If GetSelNum() = 1 Then
                    Call ShowAttrib(Index) '只选中一个则显示属性
                Else
                    Call ShowAttrib '多选时不显示属性
                End If
            Else
                Call SelItem(Index, False) '反选
                If GetSelNum() = 1 Then
                    Call ShowAttrib(intCurID) '只选中一个则显示属性(选中的不一定是该控件)
                Else
                    Call ShowAttrib '多选时不显示属性
                End If
            End If
        Else
            If Mid(lblLine(Index).Tag, 1, 2) = "" Then
                Call SelClear
                Call SelItem(Index, True)
                Call ShowAttrib(Index) '只选中一个则显示属性
            End If
        End If
    ElseIf Button = 2 Then
        PopupMenu mnuFormat, 2
    End If
End Sub

Private Sub lblLine_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If objReport.Items("_" & Index).系统 Then Exit Sub
    DrawXY X + lblLine(Index).Left, Y + lblLine(Index).Top
    If Button = 1 And Mid(lblLine(Index).Tag, 1, 2) <> "" Then
        If blnLock Then Exit Sub
        Call MoveSelect(X - lngPreX, Y - lngPreY)
        If GetSelNum() = 1 Then ShowAttrib Index
    End If
End Sub

Private Sub lblPar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lngPreY = Y
End Sub

Private Sub lblPar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ClearXY
    If Button = 1 Then
        Call NoneEdit
        If tvwSQL.Height + Y - lngPreY < 2000 Or lvwPar.Height - (Y - lngPreY) < 600 Then Exit Sub
        lblPar.Top = lblPar.Top + Y - lngPreY
        tvwSQL.Height = tvwSQL.Height + Y - lngPreY
        lvwPar.Top = lvwPar.Top + Y - lngPreY
        lvwPar.Height = lvwPar.Height - (Y - lngPreY)
        Refresh
    End If
End Sub

Private Sub lblSize_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ObjSel As Control, tmpID As RelatID
    Dim lngMinW As Long, lngMinH As Long, i As Integer
    Dim xx As Integer, yy As Integer, zz As Integer
    Dim lngTop As Long, lngLeft As Long
   
    DrawXY X + lblSize(Index).Left, Y + lblSize(Index).Top
    
    If Button = 1 And GetSelNum = 1 And Not blnLock Then
        If objReport.Items("_" & intCurID).系统 Then Exit Sub
        Select Case objReport.Items("_" & intCurID).类型
            Case 1
                Set ObjSel = lblLine(intCurID)
                lngMinW = lblSize(0).Width: lngMinH = lblSize(0).Height
            Case 2, 3
                Set ObjSel = lbl(intCurID)
                lngMinW = lblSize(0).Width: lngMinH = lblSize(0).Height
            Case 10
                Set ObjSel = Shp(intCurID)
                lngMinW = lblSize(0).Width: lngMinH = lblSize(0).Height
            Case 4 '任意表格
                Call ResetColor(intCurID)
                Set ObjSel = msh(intCurID)
                lngMinW = msh(intCurID).ColWidth(0) + 15
                
                lngMinH = 0
                For xx = 0 To msh(intCurID).FixedRows
                    If msh(intCurID).Rows - 1 >= msh(intCurID).FixedRows + 1 Then
                        lngMinH = lngMinH + msh(intCurID).RowHeight(xx)
                    Else
                        lngMinH = lngMinH + 255 * sgnMode
                    End If
                Next
                lngMinH = lngMinH + 15
                Call CustomColColor(intCurID, 0)
            Case 5 '汇总表格
                Call ResetColor(intCurID)
                Set ObjSel = msh(intCurID)
                xx = msh(intCurID).FixedCols '纵向分类项目数
                yy = msh(intCurID).FixedRows - 1 '横向分类项目数
                For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                    If objReport.Items("_" & tmpID.id).类型 = 9 Then zz = zz + 1 '统计项目数
                Next
                lngMinH = msh(intCurID).RowHeight(0) * (yy + 1) + 60 '最小表头+1行高度
                For i = 0 To xx + zz - 1
                    lngMinW = lngMinW + msh(intCurID).ColWidth(i)
                Next
                lngMinW = lngMinW + 60
            Case 11
                Set ObjSel = img(intCurID)
                lngMinW = lblSize(0).Width: lngMinH = lblSize(0).Height
            Case 12 '@@@
                Set ObjSel = Chart(intCurID)
                lngMinW = Chart(0).Width: lngMinH = Chart(0).Height
            Case 13
                Set ObjSel = ImgCode(intCurID)
                lngMinW = lblSize(0).Width: lngMinH = lblSize(0).Height
            Case 14
                Set ObjSel = pic(intCurID)
                lngMinW = lblSize(0).Width: lngMinH = lblSize(0).Height
        End Select
        
        '@@@
        lngMinW = lngMinW * sgnMode
        lngMinH = lngMinH * sgnMode
        If UCase(ObjSel.Container.name) = "PIC" Then
            lngTop = ObjSel.Container.Top
            lngLeft = ObjSel.Container.Left
        End If
        '改变控件尺寸大小
        Select Case Index Mod 8
            Case 1 '中上
                If ObjSel.Height - Y < lngMinH Then Exit Sub
                If objReport.Items("_" & intCurID).类型 = 12 Then '@@@
                    If ObjSel.Top + lngTop + Y < 0 Then Exit Sub
                End If
                
                lblSize(Index).Top = lblSize(Index).Top + Y
                lblSize(Index + 1).Top = lblSize(Index + 1).Top + Y
                lblSize(Index + 7).Top = lblSize(Index + 7).Top + Y
                
                ObjSel.Top = ObjSel.Top + Y
                ObjSel.Height = ObjSel.Height - Y
                If objReport.Items("_" & intCurID).类型 = 10 Then
                    lblshp(intCurID).Top = lblshp(intCurID).Top + Y
                    lblshp(intCurID).Height = lblshp(intCurID).Height - Y
                End If
                
                lblSize(Index + 2).Top = ObjSel.Top + lngTop + (ObjSel.Height - lblSize(Index + 2).Height) / 2
                lblSize(Index + 6).Top = lblSize(Index + 2).Top
            Case 2 '右上
                If ObjSel.Height - Y < lngMinH Or ObjSel.Width + X < lngMinW Then Exit Sub
                If objReport.Items("_" & intCurID).类型 = 12 Then '@@@
                    If ObjSel.Top + lngTop + Y < 0 Then Exit Sub
                End If
                
                lblSize(Index - 1).Top = lblSize(Index - 1).Top + Y
                lblSize(Index).Top = lblSize(Index).Top + Y
                lblSize(Index + 6).Top = lblSize(Index + 6).Top + Y

                lblSize(Index).Left = lblSize(Index).Left + X
                lblSize(Index + 1).Left = lblSize(Index + 1).Left + X
                lblSize(Index + 2).Left = lblSize(Index + 2).Left + X
                
                ObjSel.Top = ObjSel.Top + Y
                ObjSel.Height = ObjSel.Height - Y
                ObjSel.Width = ObjSel.Width + X
                If objReport.Items("_" & intCurID).类型 = 10 Then
                    lblshp(intCurID).Top = lblshp(intCurID).Top + Y
                    lblshp(intCurID).Height = lblshp(intCurID).Height - Y
                    lblshp(intCurID).Width = lblshp(intCurID).Width + X
                End If
                
                lblSize(Index + 1).Top = ObjSel.Top + lngTop + (ObjSel.Height - lblSize(Index + 1).Height) / 2
                lblSize(Index + 5).Top = lblSize(Index + 1).Top
                lblSize(Index - 1).Left = ObjSel.Left + lngLeft + (ObjSel.Width - lblSize(Index - 1).Width) / 2
                lblSize(Index + 3).Left = lblSize(Index - 1).Left
            Case 3 '右中
                If ObjSel.Width + X < lngMinW Then Exit Sub
                
                lblSize(Index - 1).Left = lblSize(Index - 1).Left + X
                lblSize(Index).Left = lblSize(Index).Left + X
                lblSize(Index + 1).Left = lblSize(Index + 1).Left + X
                
                ObjSel.Width = ObjSel.Width + X
                If objReport.Items("_" & intCurID).类型 = 10 Then
                    lblshp(intCurID).Width = lblshp(intCurID).Width + X
                End If
                
                lblSize(Index - 2).Left = ObjSel.Left + lngLeft + (ObjSel.Width - lblSize(Index - 2).Width) / 2
                lblSize(Index + 2).Left = lblSize(Index - 2).Left
            Case 4 '右下
                If ObjSel.Height + Y < lngMinH Or ObjSel.Width + X < lngMinW Then Exit Sub
                
                lblSize(Index).Left = lblSize(Index).Left + X
                lblSize(Index - 1).Left = lblSize(Index - 1).Left + X
                lblSize(Index - 2).Left = lblSize(Index - 2).Left + X
                lblSize(Index).Top = lblSize(Index).Top + Y
                lblSize(Index + 1).Top = lblSize(Index + 1).Top + Y
                lblSize(Index + 2).Top = lblSize(Index + 2).Top + Y
                
                ObjSel.Height = ObjSel.Height + Y
                ObjSel.Width = ObjSel.Width + X
                If objReport.Items("_" & intCurID).类型 = 10 Then
                    lblshp(intCurID).Height = lblshp(intCurID).Height + Y
                    lblshp(intCurID).Width = lblshp(intCurID).Width + X
                End If
                
                lblSize(Index - 1).Top = ObjSel.Top + lngTop + (ObjSel.Height - lblSize(Index - 1).Height) / 2
                lblSize(Index + 3).Top = lblSize(Index - 1).Top
                lblSize(Index - 3).Left = ObjSel.Left + lngLeft + (ObjSel.Width - lblSize(Index - 3).Width) / 2
                lblSize(Index + 1).Left = lblSize(Index - 3).Left
            Case 5 '中下
                If ObjSel.Height + Y < lngMinH Then Exit Sub
                
                lblSize(Index).Top = lblSize(Index).Top + Y
                lblSize(Index - 1).Top = lblSize(Index - 1).Top + Y
                lblSize(Index + 1).Top = lblSize(Index + 1).Top + Y
                
                ObjSel.Height = ObjSel.Height + Y
                If objReport.Items("_" & intCurID).类型 = 10 Then
                    lblshp(intCurID).Height = lblshp(intCurID).Height + Y
                End If
                
                lblSize(Index - 2).Top = ObjSel.Top + lngTop + (ObjSel.Height - lblSize(Index - 2).Height) / 2
                lblSize(Index + 2).Top = lblSize(Index - 2).Top
            Case 6 '左下
                If ObjSel.Height + Y < lngMinH Or ObjSel.Width - X < lngMinW Then Exit Sub
                If objReport.Items("_" & intCurID).类型 = 12 Then '@@@
                    If ObjSel.Left + lngLeft + X < 0 Then Exit Sub
                End If
                
                lblSize(Index).Top = lblSize(Index).Top + Y
                lblSize(Index - 1).Top = lblSize(Index - 1).Top + Y
                lblSize(Index - 2).Top = lblSize(Index - 2).Top + Y
                
                lblSize(Index).Left = lblSize(Index).Left + X
                lblSize(Index + 1).Left = lblSize(Index + 1).Left + X
                lblSize(Index + 2).Left = lblSize(Index + 2).Left + X
                
                ObjSel.Width = ObjSel.Width - X
                ObjSel.Height = ObjSel.Height + Y
                ObjSel.Left = ObjSel.Left + X
                If objReport.Items("_" & intCurID).类型 = 10 Then
                    lblshp(intCurID).Left = lblshp(intCurID).Left + X
                    lblshp(intCurID).Height = lblshp(intCurID).Height + Y
                    lblshp(intCurID).Width = lblshp(intCurID).Width - X
                End If
                
                lblSize(Index - 1).Left = ObjSel.Left + lngLeft + (ObjSel.Width - lblSize(Index - 1).Width) / 2
                lblSize(Index - 5).Left = lblSize(Index - 1).Left
                lblSize(Index + 1).Top = ObjSel.Top + lngTop + (ObjSel.Height - lblSize(Index + 1).Height) / 2
                lblSize(Index - 3).Top = lblSize(Index + 1).Top
            Case 7 '左中
                If ObjSel.Width - X < lngMinW Then Exit Sub
                If objReport.Items("_" & intCurID).类型 = 12 Then '@@@
                    If ObjSel.Left + lngLeft + X < 0 Then Exit Sub
                End If
                
                lblSize(Index).Left = lblSize(Index).Left + X
                lblSize(Index + 1).Left = lblSize(Index + 1).Left + X
                lblSize(Index - 1).Left = lblSize(Index - 1).Left + X
                
                ObjSel.Width = ObjSel.Width - X
                ObjSel.Left = ObjSel.Left + X
                If objReport.Items("_" & intCurID).类型 = 10 Then
                    lblshp(intCurID).Left = lblshp(intCurID).Left + X
                    lblshp(intCurID).Width = lblshp(intCurID).Width - X
                End If
                
                lblSize(Index - 6).Left = ObjSel.Left + lngLeft + (ObjSel.Width - lblSize(Index - 6).Width) / 2
                lblSize(Index - 2).Left = lblSize(Index - 6).Left
            Case 0 '左上
                If ObjSel.Height - Y < lngMinH Or ObjSel.Width - X < lngMinW Then Exit Sub
                If objReport.Items("_" & intCurID).类型 = 12 Then '@@@
                    If ObjSel.Top + lngTop + Y < 0 Then Exit Sub
                    If ObjSel.Left + lngLeft + X < 0 Then Exit Sub
                End If
                
                lblSize(Index).Left = lblSize(Index).Left + X
                lblSize(Index - 1).Left = lblSize(Index - 1).Left + X
                lblSize(Index - 2).Left = lblSize(Index - 2).Left + X
                
                lblSize(Index).Top = lblSize(Index).Top + Y
                lblSize(Index - 6).Top = lblSize(Index - 6).Top + Y
                lblSize(Index - 7).Top = lblSize(Index - 7).Top + Y
                
                ObjSel.Width = ObjSel.Width - X
                ObjSel.Height = ObjSel.Height - Y
                ObjSel.Left = ObjSel.Left + X
                ObjSel.Top = ObjSel.Top + Y
                If objReport.Items("_" & intCurID).类型 = 10 Then
                    lblshp(intCurID).Left = lblshp(intCurID).Left + X
                    lblshp(intCurID).Height = lblshp(intCurID).Height - Y
                    lblshp(intCurID).Width = lblshp(intCurID).Width - X
                    lblshp(intCurID).Top = lblshp(intCurID).Top + Y
                End If
                
                lblSize(Index - 1).Top = ObjSel.Top + lngTop + (ObjSel.Height - lblSize(Index - 1).Height) / 2
                lblSize(Index - 5).Top = lblSize(Index - 1).Top
                lblSize(Index - 7).Left = ObjSel.Left + lngLeft + (ObjSel.Width - lblSize(Index - 7).Width) / 2
                lblSize(Index - 3).Left = lblSize(Index - 7).Left
        End Select
        If objReport.Items("_" & intCurID).类型 <> 10 Then Me.Refresh
        
        '图片尺寸调整时保持比例
        If objReport.Items("_" & intCurID).类型 = 11 Then
            If Not objReport.Items("_" & intCurID).图片 Is Nothing Then
                If objReport.Items("_" & intCurID).粗体 Then
                    Set ObjSel.Picture = ScalePicture(PicFontTest, objReport.Items("_" & intCurID).图片, ObjSel.Width, ObjSel.Height)
                End If
            End If
        End If
        
        Dim MainItem As RPTItem
        If InStr(1, "4,5", objReport.Items("_" & intCurID).类型) > 0 And objReport.Items("_" & intCurID).参照 <> "" Then
            For Each MainItem In objReport.Items
                If MainItem.格式号 = mbytCurrFmt And MainItem.名称 = objReport.Items("_" & intCurID).参照 Then Exit For
            Next
        End If
        If InStr(1, "4,5", objReport.Items("_" & intCurID).类型) > 0 Then
            Call SetGridLine(intCurID)  '填充表格线
            If objReport.Items("_" & intCurID).类型 = 4 Then
                Call SetCopyGrid(intCurID) '处理分栏控件
            End If
            If Not MainItem Is Nothing Then
                If objReport.Items("_" & intCurID).性质 <> 1 Then
                    msh(MainItem.Key).Height = msh(intCurID).Height
                    objReport.Items("_" & MainItem.Key).H = msh(MainItem.Key).Height / sgnMode
                Else
                    msh(MainItem.Key).Width = msh(intCurID).Width
                    objReport.Items("_" & MainItem.Key).W = msh(MainItem.Key).Width / sgnMode
                End If
                Call SetMainWH(MainItem.Key)    '处理从表(左联接或附加)
            Else
                Call SetChildWH(intCurID)    '处理从表(左联接或附加)
            End If
        End If
        
        '更新数据对象
        objReport.Items("_" & intCurID).X = Format(ObjSel.Left / sgnMode, "0.00")
        objReport.Items("_" & intCurID).Y = Format(ObjSel.Top / sgnMode, "0.00")
        If objReport.Items("_" & intCurID).类型 = 1 Then
            If ObjSel.Width > ObjSel.Height Then
                objReport.Items("_" & intCurID).W = Format(ObjSel.Width / sgnMode, "0.00")
                objReport.Items("_" & intCurID).H = 0
            Else
                objReport.Items("_" & intCurID).W = 0
                objReport.Items("_" & intCurID).H = Format(ObjSel.Height / sgnMode, "0.00")
            End If
        Else
            objReport.Items("_" & intCurID).W = Format(ObjSel.Width / sgnMode, "0.00")
            objReport.Items("_" & intCurID).H = Format(ObjSel.Height / sgnMode, "0.00")
        End If
        
        If GetSelNum = 1 And InStr(1, "4,5", objReport.Items("_" & intCurID).类型) <> 0 Then
            Call AdjustSelCons(msh(intCurID))
        End If
        Call MoveSelect(0, 0, True)
        Call ShowAttrib(intCurID)
        Call AdjustAll(True)
        BlnSave = False
    End If
End Sub

Private Sub lvwPar_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Call ClearXY
End Sub

Private Sub lvwPar_GotFocus()
    Call NoneEdit
End Sub

Private Sub mnuClass_Align_Style_Click(Index As Integer)
'功能：对汇总表格,设置当前子项对齐方式
    Dim X As Integer, Y As Integer, Z As Integer, intDel As Integer
    Dim tmpID As RelatID, tmpItem As RPTItem
    
    If selCell.Col1 = -1 And selCell.Row1 = -1 Then sta.Panels(2).Text = "不能确定当前项目类型及位置！": PlayWarn: Exit Sub
    
    '统计项范围
    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
        Set tmpItem = objReport.Items("_" & tmpID.id)
        If tmpItem.类型 = 9 And tmpItem.序号 = selCell.Col1 - msh(intCurID).FixedCols Then
            tmpItem.对齐 = Index: Exit For
        End If
    Next
    BlnSave = False
End Sub

Private Sub mnuClass_Data_Click()
'功能：对汇总表格,设置当前项目的数据来源
    Dim intState As Integer
    Dim tmpID As RelatID, tmpItem As RPTItem
    Dim objRelations As New RPTRelations
    Dim objColProtertys As New RPTColProtertys
    Dim i As Long
    
    If selCell.Col1 = -1 Or selCell.Row1 = -1 Then
        sta.Panels(2).Text = "不能确定当前项目类型及位置！"
        Call PlayWarn
        Exit Sub
    End If
    
    '求汇总表格统计项个数
    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
        If objReport.Items("_" & tmpID.id).类型 = 9 Then intState = intState + 1
    Next
    
    If selCell.Col1 <= msh(intCurID).FixedCols - 1 And selCell.Row1 >= msh(intCurID).FixedRows - 1 Then
        '纵向分类范围
        frmData.I_strTitle = "纵向分类数据项"
        frmData.I_bytType = 0
        frmData.I_strClass = objReport.Items("_" & intCurID).内容
        frmData.mintEleID = intCurID
        For Each tmpID In objReport.Items("_" & intCurID).SubIDs
            Set tmpItem = objReport.Items("_" & tmpID.id)
            If tmpItem.序号 = selCell.Col1 And tmpItem.类型 = 7 Then
                frmData.IO_FontBold = IIF(tmpItem.粗体, 1, 0)
                frmData.IO_FontColor = tmpItem.前景
                frmData.I_strOrder = tmpItem.排序: Exit For
            End If
        Next
        Set frmData.objReport = objReport
        Call SetCopyRelations(tmpItem.Relations, objRelations)
        Set frmData.mobjRelations = objRelations
        Set frmData.frmParent = Me
        frmData.IO_strNode = GetDataName(msh(intCurID).TextMatrix(msh(intCurID).FixedRows - 1, selCell.Col1))
        frmData.Show 1, Me
        If gblnOK Then
            '各种项目不可重复
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                Set tmpItem = objReport.Items("_" & tmpID.id)
                If tmpItem.内容 = frmData.txtData.Text And (tmpItem.序号 <> selCell.Col1 Or tmpItem.类型 <> 7) Then
                    sta.Panels(2).Text = "你所选择的数据项在该汇总表格中已经存在！"
                    Unload frmData: PlayWarn: Exit Sub
                End If
            Next
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                Set tmpItem = objReport.Items("_" & tmpID.id)
                If tmpItem.序号 = selCell.Col1 And tmpItem.类型 = 7 Then
                    tmpItem.内容 = frmData.txtData.Text
                    If frmData.txtOrder.Text = "" Then
                        tmpItem.排序 = ""
                    Else
                        tmpItem.排序 = frmData.txtOrder.Text
                        If frmData.optDesc Then tmpItem.排序 = "," & tmpItem.排序
                    End If
                    tmpItem.粗体 = frmData.IO_FontBold
                    tmpItem.前景 = frmData.IO_FontColor
                    Exit For
                End If
            Next
            '关联报表处理
            Set tmpItem.Relations = objRelations
            Unload frmData
            Call ReShowGrid(intCurID)
            Call ClassColor(intCurID, selCell)
            BlnSave = False
        End If
    ElseIf selCell.Row1 <= msh(intCurID).FixedRows - 2 Then
        '横向分类范围
        frmData.I_strTitle = "横向分类数据项"
        frmData.I_bytType = 1
        frmData.I_strClass = objReport.Items("_" & intCurID).内容
        frmData.mintEleID = intCurID
        For Each tmpID In objReport.Items("_" & intCurID).SubIDs
            Set tmpItem = objReport.Items("_" & tmpID.id)
            If tmpItem.序号 = selCell.Row1 And tmpItem.类型 = 8 Then
                frmData.IO_FontBold = IIF(tmpItem.粗体, 1, 0)
                frmData.IO_FontColor = tmpItem.前景
                frmData.I_strOrder = tmpItem.排序: Exit For
            End If
        Next
        Set frmData.frmParent = Me
        Set frmData.objReport = objReport
        Call SetCopyRelations(tmpItem.Relations, objRelations)
        Set frmData.mobjRelations = objRelations
        frmData.IO_strNode = GetDataName(msh(intCurID).TextMatrix(selCell.Row1, msh(intCurID).FixedCols - 1))
        frmData.Show 1, Me
        If gblnOK Then
            '各种项目不可重复
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                Set tmpItem = objReport.Items("_" & tmpID.id)
                If tmpItem.内容 = frmData.txtData.Text And (tmpItem.序号 <> selCell.Row1 Or tmpItem.类型 <> 8) Then
                    sta.Panels(2).Text = "你所选择的数据项在该汇总表格中已经存在！"
                    Unload frmData: PlayWarn: Exit Sub
                End If
            Next
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                Set tmpItem = objReport.Items("_" & tmpID.id)
                If tmpItem.序号 = selCell.Row1 And tmpItem.类型 = 8 Then
                    tmpItem.内容 = frmData.txtData.Text
                    tmpItem.粗体 = frmData.IO_FontBold
                    tmpItem.前景 = frmData.IO_FontColor
                    If frmData.txtOrder.Text = "" Then
                        tmpItem.排序 = ""
                    Else
                        tmpItem.排序 = frmData.txtOrder.Text
                        If frmData.optDesc Then tmpItem.排序 = "," & tmpItem.排序
                    End If
                    Exit For
                End If
            Next
            '关联报表处理
            Set tmpItem.Relations = objRelations
            Unload frmData
            Call ReShowGrid(intCurID)
            Call ClassColor(intCurID, selCell)
            BlnSave = False
        End If
    ElseIf selCell.Col1 >= msh(intCurID).FixedCols And selCell.Col1 <= msh(intCurID).FixedCols + intState - 1 And selCell.Row1 >= msh(intCurID).FixedRows - 1 Then
        '统计项范围
        frmData.I_strTitle = "统计数据项"
        frmData.I_strOrder = ""
        frmData.I_bytType = 2
        frmData.I_strClass = objReport.Items("_" & intCurID).内容
        frmData.mintEleID = intCurID
        For Each tmpID In objReport.Items("_" & intCurID).SubIDs
            Set tmpItem = objReport.Items("_" & tmpID.id)
            If tmpItem.序号 = selCell.Col1 - msh(intCurID).FixedCols And tmpItem.类型 = 9 Then
                frmData.IO_FontBold = IIF(tmpItem.粗体, 1, 0)
                frmData.IO_FontColor = tmpItem.前景
                frmData.I_strFormat = tmpItem.格式: Exit For
            End If
        Next
        Set frmData.frmParent = Me
        Set frmData.objReport = objReport
        Call SetCopyRelations(tmpItem.Relations, objRelations)
        Set frmData.mobjRelations = objRelations
        Call SetCopyColProtertys(tmpItem.ColProtertys, objColProtertys)
        Set frmData.mobjColProtertys = objColProtertys
        frmData.IO_strNode = GetDataName(msh(intCurID).TextMatrix(msh(intCurID).FixedRows - 1, selCell.Col1))
        frmData.I_strSummaryFile = GetSummaryFile(intCurID)
        frmData.Show 1, Me
        If gblnOK Then
            '各种项目不可重复
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                Set tmpItem = objReport.Items("_" & tmpID.id)
                If tmpItem.内容 = frmData.txtData.Text And (tmpItem.序号 <> selCell.Col1 - msh(intCurID).FixedCols Or tmpItem.类型 <> 9) Then
                    sta.Panels(2).Text = "你所选择的数据项在该汇总表格中已经存在！"
                    Unload frmData: PlayWarn: Exit Sub
                End If
            Next
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                Set tmpItem = objReport.Items("_" & tmpID.id)
                If tmpItem.序号 = selCell.Col1 - msh(intCurID).FixedCols And tmpItem.类型 = 9 Then
                    tmpItem.内容 = frmData.txtData.Text
                    tmpItem.格式 = frmData.txtFormat.Text
                    tmpItem.粗体 = frmData.IO_FontBold
                    tmpItem.前景 = frmData.IO_FontColor
                    Exit For
                End If
            Next
            '关联报表处理
            Set tmpItem.Relations = objRelations
            '列特性设置
            Set tmpItem.ColProtertys = frmData.mobjColProtertys
            
            Unload frmData
            Call ReShowGrid(intCurID)
            Call ClassColor(intCurID, selCell, intState)
            BlnSave = False
        End If
    End If
End Sub

Private Function GetSummaryFile(ByVal intIndex As Integer)
'功能：获取汇总表数据项字段
    Dim i As Long, strReturn As String
    Dim strTmp As String
    
    For i = msh(intCurID).FixedCols To msh(intCurID).Cols - 1
        strTmp = GetDataName(msh(intIndex).TextMatrix(msh(intIndex).FixedRows - 1, i))
        If InStr(strReturn, strTmp) = 0 Then strReturn = strReturn & "," & strTmp
    Next
    GetSummaryFile = Mid(strReturn, 2)
End Function

Private Sub mnuClass_Del_Click()
'功能：对汇总表格,删除当前项目
    Dim X As Integer, Y As Integer, Z As Integer, intDel As Integer
    Dim tmpID As RelatID, tmpItem As RPTItem
    
    If selCell.Col1 = -1 And selCell.Row1 = -1 Then sta.Panels(2).Text = "不能确定当前项目类型及位置！": PlayWarn: Exit Sub
    
    '求汇总表格子项个数
    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
        Select Case objReport.Items("_" & tmpID.id).类型
            Case 7
                X = X + 1
            Case 8
                Y = Y + 1
            Case 9
                Z = Z + 1
        End Select
    Next
    If selCell.Col1 <= msh(intCurID).FixedCols - 1 And selCell.Row1 >= msh(intCurID).FixedRows - 1 Then
        '纵向分类范围
        If X = 1 Then sta.Panels(2).Text = "至少要有一个纵向分类项目！": PlayWarn: Exit Sub
        For Each tmpID In objReport.Items("_" & intCurID).SubIDs
            Set tmpItem = objReport.Items("_" & tmpID.id)
            If tmpItem.类型 = 7 Then
                If tmpItem.序号 = selCell.Col1 Then
                    intDel = tmpItem.id
                ElseIf tmpItem.序号 > selCell.Col1 Then
                    tmpItem.序号 = tmpItem.序号 - 1
                End If
            End If
        Next
    ElseIf selCell.Row1 <= msh(intCurID).FixedRows - 2 Then
        '横向分类范围
        If Y = 0 Then sta.Panels(2).Text = "已经没有横向分类项目！": PlayWarn: Exit Sub
        For Each tmpID In objReport.Items("_" & intCurID).SubIDs
            Set tmpItem = objReport.Items("_" & tmpID.id)
            If tmpItem.类型 = 8 Then
                If tmpItem.序号 = selCell.Row1 Then
                    intDel = tmpItem.id
                ElseIf tmpItem.序号 > selCell.Row1 Then
                    tmpItem.序号 = tmpItem.序号 - 1
                End If
            End If
        Next
    ElseIf selCell.Col1 >= msh(intCurID).FixedCols And selCell.Col1 <= msh(intCurID).FixedCols + Z - 1 And selCell.Row1 >= msh(intCurID).FixedRows - 1 Then
        '统计项范围
        If Z = 1 Then sta.Panels(2).Text = "至少要有一个统计项目！": PlayWarn: Exit Sub
        For Each tmpID In objReport.Items("_" & intCurID).SubIDs
            Set tmpItem = objReport.Items("_" & tmpID.id)
            If tmpItem.类型 = 9 Then
                If tmpItem.序号 = selCell.Col1 - msh(intCurID).FixedCols Then
                    intDel = tmpItem.id
                ElseIf tmpItem.序号 > selCell.Col1 - msh(intCurID).FixedCols Then
                    tmpItem.序号 = tmpItem.序号 - 1
                End If
            End If
        Next
    End If
    objReport.Items.Remove "_" & intDel
    objReport.Items("_" & intCurID).SubIDs.Remove "_" & intDel
    selCell.Col1 = -1: selCell.Row1 = -1
    Call ReShowGrid(intCurID)
    BlnSave = False
End Sub

Private Sub mnuClass_ExChange_Click()
'功能：将一个汇总表格行列对换
    Dim tmpID As RelatID, i As Integer
    
    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
        If objReport.Items("_" & tmpID.id).类型 = 8 Then i = i + 1
    Next
    If i = 0 Then sta.Panels(2).Text = "由于没有横向分类项目,不能切换！": PlayWarn: Exit Sub

    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
        If objReport.Items("_" & tmpID.id).类型 = 7 Then
            objReport.Items("_" & tmpID.id).类型 = 77
        ElseIf objReport.Items("_" & tmpID.id).类型 = 8 Then
            objReport.Items("_" & tmpID.id).类型 = 88
        End If
    Next
    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
        If objReport.Items("_" & tmpID.id).类型 = 77 Then
            objReport.Items("_" & tmpID.id).类型 = 8
        ElseIf objReport.Items("_" & tmpID.id).类型 = 88 Then
            objReport.Items("_" & tmpID.id).类型 = 7
            If objReport.Items("_" & tmpID.id).W = 0 Then objReport.Items("_" & tmpID.id).W = 1000
        End If
    Next
    selCell.Col1 = -1: selCell.Row1 = -1
    Call ReShowGrid(intCurID)
    BlnSave = False
End Sub

Private Sub mnuClass_Insert_After_Click()
'功能：对汇总表格,对插入各种子项目在后
    Dim intState As Integer
    Dim tmpID As RelatID, tmpItem As RPTItem
    Dim objRelations As New RPTRelations
    Dim objColProtertys As New RPTColProtertys
    Dim i As Long
    
    If selCell.Col1 = -1 And selCell.Row1 = -1 Then sta.Panels(2).Text = "不能确定当前项目类型及位置！": PlayWarn: Exit Sub
    
    '求汇总表格统计项个数
    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
        If objReport.Items("_" & tmpID.id).类型 = 9 Then intState = intState + 1
    Next
    
    If selCell.Col1 <= msh(intCurID).FixedCols - 1 And selCell.Row1 >= msh(intCurID).FixedRows - 1 Then
        '纵向分类范围
        Set frmData.frmParent = Me
        frmData.I_strTitle = "纵向分类数据项"
        frmData.I_bytType = 0
        frmData.I_strClass = objReport.Items("_" & intCurID).内容
        frmData.mintEleID = intCurID
        frmData.IO_strNode = ""
        Set frmData.objReport = objReport
        Set frmData.mobjRelations = objRelations
        frmData.Show 1, Me
        If gblnOK Then
            '各种项目不可重复
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                If objReport.Items("_" & tmpID.id).内容 = frmData.txtData.Text Then
                    sta.Panels(2).Text = "你所选择的数据项在该汇总表格中已经存在！"
                    Unload frmData: PlayWarn: Exit Sub
                End If
            Next
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                Set tmpItem = objReport.Items("_" & tmpID.id)
                If tmpItem.类型 = 7 And tmpItem.序号 > selCell.Col1 Then
                    tmpItem.序号 = tmpItem.序号 + 1
                End If
            Next
            intMaxID = intMaxID + 1
            Set tmpItem = objReport.Items.Add(intMaxID, mbytCurrFmt, "元素" & intMaxID, intCurID, 7, selCell.Col1 + 1 _
                , "", 0, frmData.txtData.Text, "", 0, 0, 1000, 0, 0, 0, False, "", 0, frmData.IO_FontBold = 1, False _
                , False, 0, frmData.IO_FontColor, 0, False, 0, IIF(frmData.optDesc, ",", "") & frmData.txtOrder.Text _
                , "", "", False, False, , False, , , , "_" & intMaxID)
                
            objReport.Items("_" & intCurID).SubIDs.Add intMaxID, "_" & intMaxID
            '关联报表处理
            Set tmpItem.Relations = objRelations
            Unload frmData
            selCell.Col1 = -1: selCell.Row1 = -1
            Call ReShowGrid(intCurID)
            BlnSave = False
        End If
    ElseIf selCell.Row1 <= msh(intCurID).FixedRows - 2 Then
        '横向分类范围
        Set frmData.frmParent = Me
        frmData.I_strTitle = "横向分类数据项"
        frmData.I_bytType = 1
        frmData.I_strClass = objReport.Items("_" & intCurID).内容
        frmData.mintEleID = intCurID
        frmData.IO_strNode = ""
        Set frmData.objReport = objReport
        Set frmData.mobjRelations = objRelations
        frmData.Show 1, Me
        If gblnOK Then
            '各种项目不可重复
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                If objReport.Items("_" & tmpID.id).内容 = frmData.txtData.Text Then
                    sta.Panels(2).Text = "你所选择的数据项在该汇总表格中已经存在！"
                    Unload frmData: PlayWarn: Exit Sub
                End If
            Next
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                Set tmpItem = objReport.Items("_" & tmpID.id)
                If tmpItem.类型 = 8 And tmpItem.序号 > selCell.Row1 Then
                    tmpItem.序号 = tmpItem.序号 + 1
                End If
            Next
            intMaxID = intMaxID + 1
            Set tmpItem = objReport.Items.Add(intMaxID, mbytCurrFmt, "元素" & intMaxID, intCurID, 8, selCell.Row1 + 1 _
                , "", 0, frmData.txtData.Text, "", 0, 0, 0, 0, 0, 0, False, "", 0, frmData.IO_FontBold = 1, False, False _
                , 0, frmData.IO_FontColor, 0, False, 0, IIF(frmData.optDesc, ",", "") & frmData.txtOrder.Text, "", "" _
                , False, False, , False, , , , "_" & intMaxID)
                
            objReport.Items("_" & intCurID).SubIDs.Add intMaxID, "_" & intMaxID
            '关联报表处理
            Set tmpItem.Relations = objRelations
            Unload frmData
            selCell.Col1 = -1: selCell.Row1 = -1
            Call ReShowGrid(intCurID)
            BlnSave = False
        End If
    ElseIf selCell.Col1 >= msh(intCurID).FixedCols And selCell.Col1 <= msh(intCurID).FixedCols + intState - 1 _
        And selCell.Row1 >= msh(intCurID).FixedRows - 1 Then
        '统计项范围
        Set frmData.frmParent = Me
        frmData.I_strTitle = "统计数据项"
        frmData.I_bytType = 2
        frmData.I_strClass = objReport.Items("_" & intCurID).内容
        frmData.mintEleID = intCurID
        frmData.IO_strNode = ""
        Set frmData.objReport = objReport
        Set frmData.mobjRelations = objRelations
        Set frmData.mobjColProtertys = objColProtertys
        frmData.I_strSummaryFile = GetSummaryFile(intCurID)
        frmData.Show 1, Me
        If gblnOK Then
            '各种项目不可重复
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                If objReport.Items("_" & tmpID.id).内容 = frmData.txtData.Text Then
                    sta.Panels(2).Text = "你所选择的数据项在该汇总表格中已经存在！"
                    Unload frmData: PlayWarn: Exit Sub
                End If
            Next
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                Set tmpItem = objReport.Items("_" & tmpID.id)
                If tmpItem.类型 = 9 And tmpItem.序号 > selCell.Col1 - msh(intCurID).FixedCols Then
                    tmpItem.序号 = tmpItem.序号 + 1
                End If
            Next
            intMaxID = intMaxID + 1
            Set tmpItem = objReport.Items.Add(intMaxID, mbytCurrFmt, "元素" & intMaxID, intCurID, 9 _
                , selCell.Col1 - msh(intCurID).FixedCols + 1, "", 0, frmData.txtData.Text, "", 0, 0, 1000, 0, 0, 0 _
                , False, "", 0, frmData.IO_FontBold = 1, False, False, 0, frmData.IO_FontColor, 0, False, 0, "" _
                , frmData.txtFormat.Text, "", False, False, , False, , , , "_" & intMaxID)
                
            objReport.Items("_" & intCurID).SubIDs.Add intMaxID, "_" & intMaxID
            '关联报表处理
            Set tmpItem.Relations = objRelations
            '列特性设置
            Set tmpItem.ColProtertys = frmData.mobjColProtertys
            Unload frmData
            selCell.Col1 = -1: selCell.Row1 = -1
            Call ReShowGrid(intCurID)
            BlnSave = False
        End If
    End If
End Sub

Private Sub mnuClass_Insert_Before_Click()
'功能：对汇总表格,对插入各种子项目在前
    Dim intState As Integer
    Dim tmpID As RelatID, tmpItem As RPTItem
    Dim objRelations As New RPTRelations
    Dim objColProtertys As New RPTColProtertys
    Dim i As Long
    
    If selCell.Col1 = -1 And selCell.Row1 = -1 Then sta.Panels(2).Text = "不能确定当前项目类型及位置！": PlayWarn: Exit Sub
    
    '求汇总表格统计项个数
    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
        If objReport.Items("_" & tmpID.id).类型 = 9 Then intState = intState + 1
    Next
    
    If selCell.Col1 <= msh(intCurID).FixedCols - 1 And selCell.Row1 >= msh(intCurID).FixedRows - 1 Then
        '纵向分类范围
        Set frmData.frmParent = Me
        frmData.I_strTitle = "纵向分类数据项"
        frmData.I_bytType = 0
        frmData.I_strClass = objReport.Items("_" & intCurID).内容
        frmData.mintEleID = intCurID
        frmData.IO_strNode = ""
        Set frmData.objReport = objReport
        Set frmData.mobjRelations = objRelations
        frmData.Show 1, Me
        If gblnOK Then
            '各种项目不可重复
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                If objReport.Items("_" & tmpID.id).内容 = frmData.txtData.Text Then
                    sta.Panels(2).Text = "你所选择的数据项在该汇总表格中已经存在！"
                    Unload frmData: PlayWarn: Exit Sub
                End If
            Next
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                Set tmpItem = objReport.Items("_" & tmpID.id)
                If tmpItem.类型 = 7 And tmpItem.序号 >= selCell.Col1 Then
                    tmpItem.序号 = tmpItem.序号 + 1
                End If
            Next
            intMaxID = intMaxID + 1
            Set tmpItem = objReport.Items.Add(intMaxID, mbytCurrFmt, "元素" & intMaxID, intCurID, 7, selCell.Col1, "", 0 _
                , frmData.txtData.Text, "", 0, 0, 1000, 0, 0, 0, False, "", 0, frmData.IO_FontBold = 1, False, False, 0 _
                , frmData.IO_FontColor, 0, False, 0, IIF(frmData.optDesc, ",", "") & frmData.txtOrder.Text, "", "" _
                , False, False, , False, , , , "_" & intMaxID)
            
            objReport.Items("_" & intCurID).SubIDs.Add intMaxID, "_" & intMaxID
            '关联报表处理
            Set tmpItem.Relations = objRelations
            Unload frmData
            selCell.Col1 = -1: selCell.Row1 = -1
            Call ReShowGrid(intCurID)
            BlnSave = False
        End If
    ElseIf selCell.Row1 <= msh(intCurID).FixedRows - 2 Then
        '横向分类范围
        Set frmData.frmParent = Me
        frmData.I_strTitle = "横向分类数据项"
        frmData.I_bytType = 1
        frmData.I_strClass = objReport.Items("_" & intCurID).内容
        frmData.mintEleID = intCurID
        frmData.IO_strNode = ""
        Set frmData.objReport = objReport
        Set frmData.mobjRelations = objRelations
        frmData.Show 1, Me
        If gblnOK Then
            '各种项目不可重复
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                If objReport.Items("_" & tmpID.id).内容 = frmData.txtData Then
                    sta.Panels(2).Text = "你所选择的数据项在该汇总表格中已经存在！"
                    Unload frmData: PlayWarn: Exit Sub
                End If
            Next
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                Set tmpItem = objReport.Items("_" & tmpID.id)
                If tmpItem.类型 = 8 And tmpItem.序号 >= selCell.Row1 Then
                    tmpItem.序号 = tmpItem.序号 + 1
                End If
            Next
            intMaxID = intMaxID + 1
            Set tmpItem = objReport.Items.Add(intMaxID, mbytCurrFmt, "元素" & intMaxID, intCurID, 8, selCell.Row1, "", 0 _
                , frmData.txtData.Text, "", 0, 0, 0, 0, 0, 0, False, "", 0, frmData.IO_FontBold = 1, False, False, 0 _
                , frmData.IO_FontColor, 0, False, 0, IIF(frmData.optDesc, ",", "") & frmData.txtOrder.Text, "", "" _
                , False, False, , False, , , , "_" & intMaxID)
                
            objReport.Items("_" & intCurID).SubIDs.Add intMaxID, "_" & intMaxID
            '关联报表处理
            Set tmpItem.Relations = objRelations
            Unload frmData
            selCell.Col1 = -1: selCell.Row1 = -1
            Call ReShowGrid(intCurID)
            BlnSave = False
        End If
    ElseIf selCell.Col1 >= msh(intCurID).FixedCols And selCell.Col1 <= msh(intCurID).FixedCols + intState - 1 _
        And selCell.Row1 >= msh(intCurID).FixedRows - 1 Then
        '统计项范围
        Set frmData.frmParent = Me
        frmData.I_strTitle = "统计数据项"
        frmData.I_bytType = 2
        frmData.I_strClass = objReport.Items("_" & intCurID).内容
        frmData.mintEleID = intCurID
        frmData.IO_strNode = ""
        Set frmData.objReport = objReport
        Set frmData.mobjRelations = objRelations
        Set frmData.mobjColProtertys = objColProtertys
        frmData.I_strSummaryFile = GetSummaryFile(intCurID)
        frmData.Show 1, Me
        If gblnOK Then
            '各种项目不可重复
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                If objReport.Items("_" & tmpID.id).内容 = frmData.txtData.Text Then
                    sta.Panels(2).Text = "你所选择的数据项在该汇总表格中已经存在！"
                    Unload frmData: PlayWarn: Exit Sub
                End If
            Next
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                Set tmpItem = objReport.Items("_" & tmpID.id)
                If tmpItem.类型 = 9 And tmpItem.序号 >= selCell.Col1 - msh(intCurID).FixedCols Then
                    tmpItem.序号 = tmpItem.序号 + 1
                End If
            Next
            intMaxID = intMaxID + 1
            Set tmpItem = objReport.Items.Add(intMaxID, mbytCurrFmt, "元素" & intMaxID, intCurID, 9 _
                , selCell.Col1 - msh(intCurID).FixedCols, "", 0, frmData.txtData.Text, "", 0, 0, 1000, 0, 0, 0, False, "", 0 _
                , frmData.IO_FontBold = 1, False, False, 0, frmData.IO_FontColor, 0, False, 0, "", frmData.txtFormat.Text, "" _
                , False, False, , False, , , , "_" & intMaxID)
                
            objReport.Items("_" & intCurID).SubIDs.Add intMaxID, "_" & intMaxID
            '关联报表处理
            Set tmpItem.Relations = objRelations
            '列特性设置
            Set tmpItem.ColProtertys = frmData.mobjColProtertys
            Unload frmData
            selCell.Col1 = -1: selCell.Row1 = -1
            Call ReShowGrid(intCurID)
            BlnSave = False
        End If
    End If
End Sub

Private Sub mnuClass_State_Style_Click(Index As Integer)
'功能：对汇总表格,设置当前子项汇总方式
    Dim tmpID As RelatID, tmpItem As RPTItem
    
    If selCell.Col1 = -1 And selCell.Row1 = -1 Then sta.Panels(2).Text = "不能确定当前项目类型及位置！": PlayWarn: Exit Sub
    
    If selCell.Col1 <= msh(intCurID).FixedCols - 1 And selCell.Row1 >= msh(intCurID).FixedRows - 1 Then
        '纵向分类范围
        For Each tmpID In objReport.Items("_" & intCurID).SubIDs
            Set tmpItem = objReport.Items("_" & tmpID.id)
            If tmpItem.类型 = 7 And tmpItem.序号 = selCell.Col1 Then
                Select Case Index
                    Case 0
                        tmpItem.汇总 = ""
                    Case 1
                        tmpItem.汇总 = "SUM"
                    Case 2
                        tmpItem.汇总 = "AVG"
                    Case 3
                        tmpItem.汇总 = "MAX"
                    Case 4
                        tmpItem.汇总 = "MIN"
                    Case 5
                        tmpItem.汇总 = "COUNT"
                End Select
                If tmpItem.汇总 = "" Then
                    msh(intCurID).TextMatrix(msh(intCurID).FixedRows, selCell.Col1) = msh(intCurID).TextMatrix(msh(intCurID).FixedRows + 1, selCell.Col1)
                Else
                    msh(intCurID).TextMatrix(msh(intCurID).FixedRows, selCell.Col1) = tmpItem.汇总
                End If
                Exit For
            End If
        Next
    ElseIf selCell.Row1 <= msh(intCurID).FixedRows - 2 Then
        '横向分类范围
        For Each tmpID In objReport.Items("_" & intCurID).SubIDs
            Set tmpItem = objReport.Items("_" & tmpID.id)
            If tmpItem.类型 = 8 And tmpItem.序号 = selCell.Row1 Then
                Select Case Index
                    Case 0
                        tmpItem.汇总 = ""
                    Case 1
                        tmpItem.汇总 = "SUM"
                    Case 2
                        tmpItem.汇总 = "AVG"
                    Case 3
                        tmpItem.汇总 = "MAX"
                    Case 4
                        tmpItem.汇总 = "MIN"
                    Case 5
                        tmpItem.汇总 = "COUNT"
                End Select
                If tmpItem.汇总 = "" Then
                    msh(intCurID).TextMatrix(selCell.Row1, msh(intCurID).FixedCols) = msh(intCurID).TextMatrix(selCell.Row1, msh(intCurID).FixedCols + 1)
                Else
                    msh(intCurID).TextMatrix(selCell.Row1, msh(intCurID).FixedCols) = tmpItem.汇总
                End If
                msh(intCurID).MergeCol(msh(intCurID).FixedCols) = False
                Exit For
            End If
        Next
    End If
    BlnSave = False
End Sub

Private Sub mnuCustom_Col_Align_Style_Click(Index As Integer)
'功能：对任意表格,设置当前列数据对齐方式
    Dim tmpItem As RPTItem, tmpID As RelatID
    
    If intCurCol = -1 Then sta.Panels(2).Text = "不能确定当前数据列！": PlayWarn: Exit Sub
    
    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
        Set tmpItem = objReport.Items("_" & tmpID.id)
        If tmpItem.序号 = intCurCol Then tmpItem.对齐 = Index
    Next
    msh(intCurID).ColAlignment(intCurCol) = Switch(Index = 0, 1, Index = 1, 4, Index = 2, 7)
    
    Call SetCopyGrid(intCurID)
    BlnSave = False
End Sub

Private Sub mnuCustom_Col_Clear_Click()
    Dim intCol As Integer, tmpRelatID As RelatID
    '清空表体(把公式清空)
    If intCurID = 0 Then Exit Sub
    For Each tmpRelatID In objReport.Items("_" & intCurID).SubIDs
        objReport.Items("_" & tmpRelatID.Key).内容 = ""
    Next
    Call ReShowGrid(intCurID)
    selCell.Col1 = -1: selCell.Col2 = -1: selCell.Row1 = -1: selCell.Row2 = -1: selCell.Row = -1
    Call ShowPaperInfo
    BlnSave = False
End Sub

Private Sub mnuCustom_Col_Data_Click()
'功能：对任意表格,设置当前列的数据计算方式
    Dim tmpID As RelatID, tmpItem As RPTItem
    Dim strFormula As String, strText As String, blnDo As Boolean
    Dim blnPreMerge As Boolean
    Dim k As Long, X As Long, Y As Long
    Dim objRelations As New RPTRelations
    Dim objColProtertys As New RPTColProtertys
    Dim i As Long
    
    If intCurCol = -1 Then sta.Panels(2).Text = "不能确定当前数据列！": PlayWarn: Exit Sub
    
    blnPreMerge = True
    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
        If objReport.Items("_" & tmpID.id).序号 = intCurCol Then
            Set tmpItem = objReport.Items("_" & tmpID.id)
        ElseIf objReport.Items("_" & tmpID.id).序号 = intCurCol - 1 Then
            blnPreMerge = objReport.Items("_" & tmpID.id).自调
        End If
    Next
    
    With frmFormula
        .strInit = tmpItem.内容
        .strFormat = tmpItem.格式
        .mbln换页 = tmpItem.边框
        .mblnCan换页 = True 'objReport.Items("_" & intCurID).分栏 <= 1
        .mblnMerge = tmpItem.自调
        .mblnPreMerge = blnPreMerge
        .mblnVisible = tmpItem.分栏 = 1
        .intCol = intCurCol
        .intCur = intCurID
        
        '行高（缩小字体）与自适应行高互斥
        If tmpItem.行高 = 1 Then
            .mblnAutoFont = True
            .mblnAutoRowHeight = False
        Else
            .mblnAutoFont = False
            .mblnAutoRowHeight = tmpItem.自适应行高
        End If
        
        Set .frmParent = Me
        Set .objReport = objReport
        Call SetCopyRelations(tmpItem.Relations, objRelations)
        Call SetCopyColProtertys(tmpItem.ColProtertys, objColProtertys)
        Set .mobjRelations = objRelations
        Set .mobjColProtertys = objColProtertys
        
        .Show vbModal, Me
    End With
    
    If gblnOK Then
        If tmpItem.上级ID <> 0 Then
            If objReport.Items("_" & tmpItem.上级ID).父ID <> 0 Then
                If objReport.Items("_" & objReport.Items("_" & tmpItem.上级ID).父ID).数据源 <> "" Then
                    X = InStr(1, frmFormula.txtFormula.Text, "]")
                    Y = InStr(1, frmFormula.txtFormula.Text, ".")
                    k = InStr(1, frmFormula.txtFormula.Text, "[")
                    If X > k And X > Y And X <> 0 And k <> 0 And Y <> 0 Then
                        If Mid(frmFormula.txtFormula.Text, k + 1, Y - k - 1) <> _
                            objReport.Items("_" & objReport.Items("_" & tmpItem.上级ID).父ID).数据源 Then
                            MsgBox "绑定的数据列必须属于当前卡片数据源，请检查！", vbInformation, App.Title
                            Unload frmFormula
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
        tmpItem.内容 = frmFormula.txtFormula.Text
        tmpItem.格式 = frmFormula.txtFormat.Text
        If frmFormula.chkAutoFont.Value = 1 Then
            tmpItem.行高 = 1
            tmpItem.自适应行高 = False
        Else
            tmpItem.行高 = 0
            tmpItem.自适应行高 = frmFormula.chkAutoRowHeight.Value = 1
        End If
        tmpItem.分栏 = frmFormula.chkVisible.Value
        tmpItem.自调 = (frmFormula.chkMerge.Value = 1)
        tmpItem.边框 = (frmFormula.chk换页.Value = 1)
        If tmpItem.内容 = "" Then tmpItem.汇总 = ""
        '关联报表处理
        Set tmpItem.Relations = objRelations
        
        '列特性设置
        Set tmpItem.ColProtertys = frmFormula.mobjColProtertys
        
        Unload frmFormula
        msh(intCurID).TextMatrix(msh(intCurID).FixedRows, intCurCol) = tmpItem.内容
        msh(intCurID).TextMatrix(msh(intCurID).FixedRows + 1, intCurCol) = tmpItem.汇总
        
        If Not tmpItem.自调 Then
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                If objReport.Items("_" & tmpID.id).序号 > intCurCol Then
                    objReport.Items("_" & tmpID.id).自调 = False
                End If
            Next
        End If
        If tmpItem.边框 Then
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                If objReport.Items("_" & tmpID.id).序号 <> intCurCol Then
                    objReport.Items("_" & tmpID.id).边框 = False
                End If
            Next
        End If
        
        '表头处理
        If Right(tmpItem.表头, 1) = "#" And tmpItem.内容 <> "" Then
            If Left(tmpItem.内容, 1) = "[" And Right(tmpItem.内容, 1) = "]" And InStr(tmpItem.内容, ".") > 0 _
                And InStr(Mid(tmpItem.内容, 2, Len(tmpItem.内容) - 2), "[") = 0 Then
                
                strText = Mid(tmpItem.内容, InStr(tmpItem.内容, ".") + 1, Len(tmpItem.内容) - 1 - InStr(tmpItem.内容, "."))
                
                blnDo = True
                On Error Resume Next
                If msh(intCurID).TextMatrix(msh(intCurID).FixedRows - 1, intCurCol - 1) = strText And _
                    msh(intCurID).TextMatrix(msh(intCurID).FixedRows - 2, intCurCol) = strText Then
                    If Err.Number = 0 Then blnDo = False
                End If
                If msh(intCurID).TextMatrix(msh(intCurID).FixedRows - 1, intCurCol + 1) = strText And _
                    msh(intCurID).TextMatrix(msh(intCurID).FixedRows - 2, intCurCol) = strText Then
                    If Err.Number = 0 Then blnDo = False
                End If
                On Error GoTo 0
                
                If blnDo Then
                    tmpItem.表头 = Left(tmpItem.表头, Len(tmpItem.表头) - 1) & strText
                    msh(intCurID).TextMatrix(msh(intCurID).FixedRows - 1, intCurCol) = strText
                End If
            End If
        End If
        Call SetCopyGrid(intCurID)
        BlnSave = False
    End If
    Call ShowAttrib(intCurID)       '刷新属性区域
End Sub

Private Sub mnuCustom_Col_Del_Click()
'功能：对任意表格,删除当前数据列
    Dim tmpItem As RPTItem, tmpID As RelatID
    Dim intDel As Integer
    
    If intCurCol = -1 Then sta.Panels(2).Text = "不能确定当前数据列！": PlayWarn: Exit Sub
    If objReport.Items("_" & intCurID).SubIDs.count < 2 Then sta.Panels(2).Text = "不能全部删除表格列！": PlayWarn: Exit Sub
    
    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
        Set tmpItem = objReport.Items("_" & tmpID.id)
        If tmpItem.序号 = intCurCol Then
            intDel = tmpItem.id
        ElseIf tmpItem.序号 > intCurCol Then
            tmpItem.序号 = tmpItem.序号 - 1
        End If
    Next
    
    objReport.Items.Remove "_" & intDel
    objReport.Items("_" & intCurID).SubIDs.Remove "_" & intDel
    
    Call ReShowGrid(intCurID)
    intCurCol = -1
    BlnSave = False
End Sub

Private Sub mnuCustom_Col_Insert_Left_Click()
'功能：对任意表格,在当前列左边插入一空列
    Dim tmpItem As RPTItem, tmpID As RelatID
    Dim strHead As String, i As Integer
    
    If intCurCol = -1 Then sta.Panels(2).Text = "不能确定当前数据列！": PlayWarn: Exit Sub
    
    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
        Set tmpItem = objReport.Items("_" & tmpID.id)
        If tmpItem.序号 >= intCurCol Then tmpItem.序号 = tmpItem.序号 + 1
    Next
    
    For i = 0 To msh(intCurID).FixedRows - 1
        strHead = strHead & "|4^" & msh(intCurID).RowHeight(i) & "^#"
    Next
    strHead = Mid(strHead, 2)
    i = objReport.Items("_" & intCurID).id
    intMaxID = intMaxID + 1
    
    objReport.Items.Add intMaxID, mbytCurrFmt, "元素" & intMaxID, i, 6, intCurCol, "", 0, "", strHead, 0, 0, 1000, 0, 0, 0 _
        , False, "", 0, False, False, False, 0, 0, 0, False, 0, "", "", "", False, False, , False, , , , "_" & intMaxID
    objReport.Items("_" & intCurID).SubIDs.Add intMaxID, "_" & intMaxID
    
    Call ReShowGrid(intCurID)
    intCurCol = -1
    BlnSave = False
End Sub

Private Sub mnuCustom_Col_Insert_Right_Click()
'功能：对任意表格,在当前列右边插入一空列
    Dim tmpItem As RPTItem, tmpID As RelatID
    Dim strHead As String, i As Integer
    
    If intCurCol = -1 Then sta.Panels(2).Text = "不能确定当前数据列！": PlayWarn: Exit Sub
    
    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
        Set tmpItem = objReport.Items("_" & tmpID.id)
        If tmpItem.序号 > intCurCol Then tmpItem.序号 = tmpItem.序号 + 1
    Next
    
    For i = 0 To msh(intCurID).FixedRows - 1
        strHead = strHead & "|4^" & msh(intCurID).RowHeight(i) & "^#"
    Next
    strHead = Mid(strHead, 2)
    i = objReport.Items("_" & intCurID).id
    intMaxID = intMaxID + 1
    
    objReport.Items.Add intMaxID, mbytCurrFmt, "元素" & intMaxID, i, 6, intCurCol + 1, "", 0, "", strHead, 0, 0, 1000, 0, 0, 0 _
        , False, "", 0, False, False, False, 0, 0, 0, False, 0, "", "", "", False, False, , False, , , , "_" & intMaxID
    objReport.Items("_" & intCurID).SubIDs.Add intMaxID, "_" & intMaxID
    
    Call ReShowGrid(intCurID)
    intCurCol = -1
    BlnSave = False
End Sub

Private Sub mnuCustom_Col_State_Style_Click(Index As Integer)
'功能：对任意表格,设置当前列的汇总方式
    Dim tmpItem As RPTItem, tmpID As RelatID
    
    If intCurCol = -1 Then sta.Panels(2).Text = "不能确定当前数据列！": PlayWarn: Exit Sub
    If msh(intCurID).TextMatrix(msh(intCurID).FixedRows, intCurCol) = "" Then sta.Panels(2).Text = "该列没有定义数据来源,不能汇总！": PlayWarn: Exit Sub
    
    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
        Set tmpItem = objReport.Items("_" & tmpID.id)
        If tmpItem.序号 = intCurCol Then
            tmpItem.汇总 = Switch(Index = 0, "", Index = 1, "SUM", Index = 2, "AVG", Index = 3, "MAX", Index = 4, "MIN", Index = 5, "COUNT")
        End If
    Next
    msh(intCurID).TextMatrix(msh(intCurID).FixedRows + 1, intCurCol) = Switch(Index = 0, "", Index = 1, "SUM", Index = 2, "AVG", Index = 3, "MAX", Index = 4, "MIN", Index = 5, "COUNT")
    
    Call SetCopyGrid(intCurID)
    BlnSave = False
End Sub

Private Sub mnuCustom_Head_Auto_Click()
'功能：自动将当前表头行作为列编号并填入
    Dim intBegin As Integer, i As Integer
    Dim arrHead() As String, IntAlig As Integer
    Dim tmpID As RelatID, tmpItem As RPTItem
    
    If selCell.Row = -1 Then sta.Panels(2).Text = "不能确定当前行！": PlayWarn: Exit Sub
    
    frmInput.I_blnAllowNULL = False
    frmInput.I_bytType = 1
    frmInput.I_intMaxLen = 2
    frmInput.I_strInfo = "请在下面的输入框中填写当前表头行开始的列编号！"
    frmInput.I_strTitle = "自动列编号"
    frmInput.I_strMask = "0123456789"
    frmInput.IO_strValue = "0"
    frmInput.Show 1, Me

    If gblnOK Then
        intBegin = CInt(frmInput.IO_strValue)
        IntAlig = frmInput.IO_IntAlig
        Unload frmInput
        
        For Each tmpID In objReport.Items("_" & intCurID).SubIDs
            Set tmpItem = objReport.Items("_" & tmpID.id)
            arrHead = Split(tmpItem.表头, "|")
            tmpItem.表头 = ""
            For i = 0 To UBound(arrHead)
                If i = selCell.Row Then
                    msh(intCurID).Row = i: msh(intCurID).Col = tmpItem.序号
                    msh(intCurID).CellAlignment = IntAlig
                    msh(intCurID).CellForeColor = frmInput.IO_FontColor
                    msh(intCurID).CellFontBold = frmInput.IO_FontBold
                    tmpItem.表头 = tmpItem.表头 & "|" & msh(intCurID).CellAlignment & "^" & msh(intCurID).RowHeight(i) & "^<" & intBegin + tmpItem.序号 & ">" & "^" & IIF(msh(intCurID).CellFontBold, 1, 0) & "^" & msh(intCurID).CellForeColor
                Else
                    tmpItem.表头 = tmpItem.表头 & "|" & arrHead(i)
                End If
            Next
            tmpItem.表头 = Mid(tmpItem.表头, 2)
        Next
        
        Call ReShowGrid(intCurID)
        selCell.Col1 = -1: selCell.Col2 = -1: selCell.Row1 = -1: selCell.Row2 = -1: selCell.Row = -1
        Call ShowPaperInfo
        BlnSave = False
    End If
End Sub

Private Sub mnuCustom_Head_Clear_Click()
    Dim intCol As Integer, tmpRelatID As RelatID
    '清空表头,只保留一行固定行
    If intCurID = 0 Then Exit Sub
    msh(intCurID).FixedRows = 1
    
    For intCol = 0 To msh(intCurID).Cols - 1
        msh(intCurID).TextMatrix(0, intCol) = ""
    Next
    For Each tmpRelatID In objReport.Items("_" & intCurID).SubIDs
        msh(intCurID).Row = 0: msh(intCurID).Col = objReport.Items("_" & tmpRelatID.Key).序号
        msh(intCurID).RowHeight(0) = 30
        objReport.Items("_" & tmpRelatID.Key).表头 = msh(intCurID).CellAlignment & "^" & msh(intCurID).RowHeight(0) & "^#"
    Next
    Call ReShowGrid(intCurID)
    selCell.Col1 = -1: selCell.Col2 = -1: selCell.Row1 = -1: selCell.Row2 = -1: selCell.Row = -1
    Call ShowPaperInfo
    BlnSave = False
End Sub

Private Sub mnuCustom_Head_Del_Click()
'功能：对任意表格,删除当前选择行
    Dim tmpID As RelatID, tmpItem As RPTItem, StrDelName As String
    Dim arrHead() As String, i As Integer
    
    If selCell.Row = -1 Then sta.Panels(2).Text = "不能确定当前行！": PlayWarn: Exit Sub
    If msh(intCurID).FixedRows = 1 Then sta.Panels(2).Text = "表头至少要保留一行！": PlayWarn: Exit Sub
    
    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
        Set tmpItem = objReport.Items("_" & tmpID.id)
        arrHead = Split(tmpItem.表头, "|")
        tmpItem.表头 = ""
        For i = 0 To UBound(arrHead)
            '删除当前行内容
            If i <> selCell.Row Then
                If i = selCell.Row + 1 Then
                    StrDelName = msh(intCurID).TextMatrix(i, tmpItem.序号)
                    msh(intCurID).Row = i: msh(intCurID).Col = tmpItem.序号
                    tmpItem.表头 = tmpItem.表头 & "|" & msh(intCurID).CellAlignment & "^" & msh(intCurID).RowHeight(i) & "^" & StrDelName
                Else
                    tmpItem.表头 = tmpItem.表头 & "|" & arrHead(i)
                End If
            End If
        Next
        tmpItem.表头 = Mid(tmpItem.表头, 2)
    Next
    Call ReShowGrid(intCurID)
    selCell.Col1 = -1: selCell.Col2 = -1: selCell.Row1 = -1: selCell.Row2 = -1: selCell.Row = -1
    Call ShowPaperInfo
    BlnSave = False
End Sub

Private Sub mnuCustom_Head_Insert_Down_Click()
'功能：对任意表格表头,在当前选择行下方插入一个空行
'说明：以SelCell.Row为参照行
    Dim tmpID As RelatID, tmpItem As RPTItem
    Dim arrHead() As String, i As Integer
    
    If selCell.Row = -1 Then sta.Panels(2).Text = "不能确定当前行！": PlayWarn: Exit Sub
    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
        Set tmpItem = objReport.Items("_" & tmpID.id)
        arrHead = Split(tmpItem.表头, "|")
        tmpItem.表头 = ""
        For i = 0 To UBound(arrHead)
            If i = selCell.Row Then
                tmpItem.表头 = tmpItem.表头 & "|" & arrHead(i) & "|4^255^#" '新插入行在下
            Else
                tmpItem.表头 = tmpItem.表头 & "|" & arrHead(i)
            End If
        Next
        tmpItem.表头 = Mid(tmpItem.表头, 2)
    Next
    Call ReShowGrid(intCurID)
    selCell.Col1 = -1: selCell.Col2 = -1: selCell.Row1 = -1: selCell.Row2 = -1: selCell.Row = -1
    Call ShowPaperInfo
    BlnSave = False
End Sub

Private Sub mnuCustom_Head_Insert_UP_Click()
'功能：对任意表格表头,在当前选择行上方插入一个空行
'说明：以SelCell.Row为参照行
    Dim tmpID As RelatID, tmpItem As RPTItem
    Dim arrHead() As String, i As Integer
    
    If selCell.Row = -1 Then sta.Panels(2).Text = "不能确定当前行！": PlayWarn: Exit Sub
    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
        Set tmpItem = objReport.Items("_" & tmpID.id)
        arrHead = Split(tmpItem.表头, "|")
        tmpItem.表头 = ""
        For i = 0 To UBound(arrHead)
            If i = selCell.Row Then
                tmpItem.表头 = tmpItem.表头 & "|4^255^#|" & arrHead(i) '新插入行在上
            Else
                tmpItem.表头 = tmpItem.表头 & "|" & arrHead(i)
            End If
        Next
        tmpItem.表头 = Mid(tmpItem.表头, 2)
    Next
    Call ReShowGrid(intCurID)
    selCell.Col1 = -1: selCell.Col2 = -1: selCell.Row1 = -1: selCell.Row2 = -1: selCell.Row = -1
    Call ShowPaperInfo
    BlnSave = False
End Sub

Private Sub mnuCustom_Head_Merge_Click()
'功能：对任意表格表头,合并当前选择范围的单元格
'说明：设置后不清除当前选择单元格
    Dim i As Integer, j As Integer, strText As String
    Dim tmpID As RelatID, tmpItem As RPTItem
    Dim arrHead() As String, blnDo As Boolean
    
    If selCell.Row1 = -1 Then sta.Panels(2).Text = "不能确定当前选择单元格！": PlayWarn: Exit Sub
    If selCell.Row1 = selCell.Row2 And selCell.Col1 = selCell.Col2 Then sta.Panels(2).Text = "一个单元格不用合并！": PlayWarn: Exit Sub
    If selCell.Row1 <> selCell.Row2 And selCell.Col1 <> selCell.Col2 Then sta.Panels(2).Text = "单元格同时只能在一个方向上合并！": PlayWarn: Exit Sub
    '如果当前选择范围单元格内容全部相同则不用合并
    For i = selCell.Row1 To selCell.Row2
        For j = selCell.Col1 To selCell.Col2
            If i = selCell.Row1 And j = selCell.Col1 Then
                strText = msh(intCurID).TextMatrix(i, j)
            Else
                If msh(intCurID).TextMatrix(i, j) <> strText Or strText = "" Then
                    blnDo = True: Exit For
                End If
            End If
        Next
    Next
    If Not blnDo Then sta.Panels(2).Text = "当前选择单元格已经合并！": PlayWarn: Exit Sub
    
    frmInput.I_strTitle = "合并单元格"
    frmInput.I_strInfo = "请在下面的输入框中输入任意表格的表头单元格合并后的文字。"
    frmInput.I_blnAllowNULL = False
    frmInput.I_intMaxLen = 50
    frmInput.IO_strValue = ""
    For i = selCell.Row1 To selCell.Row2
        For j = selCell.Col1 To selCell.Col2
            If msh(intCurID).TextMatrix(i, j) <> "" Then
                frmInput.IO_strValue = msh(intCurID).TextMatrix(i, j)
                msh(intCurID).Row = i: msh(intCurID).Col = j
                frmInput.IO_FontBold = IIF(msh(intCurID).CellFontBold, 1, 0)
                frmInput.IO_FontColor = msh(intCurID).CellForeColor
                Exit For
            End If
        Next
    Next
    frmInput.Show 1, Me
    If gblnOK Then
        '更新控件
        For i = selCell.Row1 To selCell.Row2
            For j = selCell.Col1 To selCell.Col2
                If CheckCell(intCurID, i, j, frmInput.IO_strValue) Then
                    msh(intCurID).Row = i: msh(intCurID).Col = j
                    msh(intCurID).CellAlignment = frmInput.IO_IntAlig
                    msh(intCurID).TextMatrix(i, j) = frmInput.IO_strValue
                    msh(intCurID).CellForeColor = frmInput.IO_FontColor
                    msh(intCurID).CellFontBold = frmInput.IO_FontBold
                Else
                    sta.Panels(2).Text = "表头单元格同时只能在一个方向上被合并,单元格文字不能全部写入！": PlayWarn
                End If
            Next
        Next
        Unload frmInput
        '更新对象
        For Each tmpID In objReport.Items("_" & intCurID).SubIDs
            Set tmpItem = objReport.Items("_" & tmpID.id)
            If tmpItem.序号 >= selCell.Col1 And tmpItem.序号 <= selCell.Col2 Then
                arrHead = Split(tmpItem.表头, "|")
                tmpItem.表头 = ""
                For i = 0 To UBound(arrHead)
                    If i >= selCell.Row1 And i <= selCell.Row2 Then
                        msh(intCurID).Row = i: msh(intCurID).Col = tmpItem.序号
                        tmpItem.表头 = tmpItem.表头 & "|" & msh(intCurID).CellAlignment & "^" & msh(intCurID).RowHeight(i) & "^" & IIF(msh(intCurID).TextMatrix(i, tmpItem.序号) = "", "#", msh(intCurID).TextMatrix(i, tmpItem.序号)) & "^" & IIF(msh(intCurID).CellFontBold, 1, 0) & "^" & msh(intCurID).CellForeColor
                    Else
                        tmpItem.表头 = tmpItem.表头 & "|" & arrHead(i)
                    End If
                Next
                tmpItem.表头 = Mid(tmpItem.表头, 2)
            End If
        Next
        Call SetCopyGrid(intCurID)
        BlnSave = False
    End If
End Sub

Private Sub mnuCustom_Head_Split_Click()
'功能：对任意表格表头,拆分当前选择单元格
    Dim i As Integer, j As Integer, strText As String
    Dim tmpID As RelatID, tmpItem As RPTItem, arrHead() As String
    
    If selCell.Row1 = -1 Then sta.Panels(2).Text = "不能确定当前选择的单元格！": PlayWarn: Exit Sub
    If selCell.Row1 = selCell.Row2 And selCell.Col1 = selCell.Col2 Then sta.Panels(2).Text = "一个单元格不用拆分！": PlayWarn: Exit Sub
    
    '当前选择范围单元格内容必须全部相同
    For i = selCell.Row1 To selCell.Row2
        For j = selCell.Col1 To selCell.Col2
            If i = selCell.Row1 And j = selCell.Col1 Then
                strText = msh(intCurID).TextMatrix(i, j)
            Else
                If msh(intCurID).TextMatrix(i, j) <> strText Or strText = "" Then
                    sta.Panels(2).Text = "当前选择的范围不只一个单元格,不能拆分！": PlayWarn: Exit Sub
                End If
            End If
        Next
    Next
    
    '拆分单元格后内容为空
    '更新控件内容
    For i = selCell.Row1 To selCell.Row2
        For j = selCell.Col1 To selCell.Col2
            If Not (i = selCell.Row1 And j = selCell.Col1) Then
                msh(intCurID).TextMatrix(i, j) = ""
            End If
        Next
    Next
    
    '更新对象内容
    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
        Set tmpItem = objReport.Items("_" & tmpID.id)
        If tmpItem.序号 >= selCell.Col1 And tmpItem.序号 <= selCell.Col2 Then
            arrHead = Split(tmpItem.表头, "|")
            tmpItem.表头 = ""
            For i = 0 To UBound(arrHead)
                If i >= selCell.Row1 And i <= selCell.Row2 Then
                    msh(intCurID).Row = i: msh(intCurID).Col = tmpItem.序号
                    tmpItem.表头 = tmpItem.表头 & "|" & msh(intCurID).CellAlignment & "^" & msh(intCurID).RowHeight(i) & "^" & IIF(msh(intCurID).TextMatrix(i, tmpItem.序号) = "", "#", msh(intCurID).TextMatrix(i, tmpItem.序号))
                Else
                    tmpItem.表头 = tmpItem.表头 & "|" & arrHead(i)
                End If
            Next
            tmpItem.表头 = Mid(tmpItem.表头, 2)
        End If
    Next
    Call SetCopyGrid(intCurID)
    BlnSave = False
End Sub

Private Sub mnuCustom_Head_Text_Click()
'功能：对任意表格表头,在当单元格(可能是合并的)内容输入文字
'说明：设置后不清除当前选择单元格
    Dim i As Integer, j As Integer, strText As String
    Dim tmpID As RelatID, tmpItem As RPTItem
    Dim arrHead() As String
    
    If selCell.Row1 = -1 Then sta.Panels(2).Text = "不能确定当前选择的单元格！": PlayWarn: Exit Sub
    
    '当前选择范围单元格内容必须全部相同
    For i = selCell.Row1 To selCell.Row2
        For j = selCell.Col1 To selCell.Col2
            If i = selCell.Row1 And j = selCell.Col1 Then
                strText = msh(intCurID).TextMatrix(i, j)
            Else
                If msh(intCurID).TextMatrix(i, j) <> strText Or strText = "" Then
                    sta.Panels(2).Text = "当前选择的范围不只一个单元格,请先合并！": PlayWarn: Exit Sub
                End If
            End If
        Next
    Next
    
    With frmInput
        .I_strTitle = "单元格内容"
        .I_strInfo = "请在下面的输入框中输入任意表格的表头当前单元格的文字。"
        .I_blnAllowNULL = True
        .I_intMaxLen = 200
        .IO_strValue = msh(intCurID).TextMatrix(selCell.Row1, selCell.Col1)
        msh(intCurID).Row = selCell.Row1: msh(intCurID).Col = selCell.Col1
        .IO_IntAlig = msh(intCurID).CellAlignment
        .IO_FontBold = IIF(msh(intCurID).CellFontBold, 1, 0)
        .IO_FontColor = msh(intCurID).CellForeColor
        .Show 1, Me
    End With
    If gblnOK Then
        '更新控件
        For i = selCell.Row1 To selCell.Row2
            For j = selCell.Col1 To selCell.Col2
                If CheckCell(intCurID, i, j, frmInput.IO_strValue) Then
                    msh(intCurID).TextMatrix(i, j) = frmInput.IO_strValue
                    msh(intCurID).Row = i: msh(intCurID).Col = j
                    msh(intCurID).CellAlignment = frmInput.IO_IntAlig
                    msh(intCurID).CellForeColor = frmInput.IO_FontColor
                    msh(intCurID).CellFontBold = frmInput.IO_FontBold
                Else
                    sta.Panels(2).Text = "表头单元格同时只能在一个方向上被合并,单元格文字不能全部写入！": PlayWarn
                End If
            Next
        Next
        Unload frmInput
        '更新对象
        For Each tmpID In objReport.Items("_" & intCurID).SubIDs
            Set tmpItem = objReport.Items("_" & tmpID.id)
            arrHead = Split(tmpItem.表头, "|")
            tmpItem.表头 = ""
            
            For i = 0 To UBound(arrHead)
                If i >= selCell.Row1 And i <= selCell.Row2 Then
                    msh(intCurID).Row = i: msh(intCurID).Col = tmpItem.序号
                    tmpItem.表头 = tmpItem.表头 & IIF(tmpItem.表头 = "", "", "|") & msh(intCurID).CellAlignment & "^" & msh(intCurID).RowHeight(i) & "^" & IIF(msh(intCurID).TextMatrix(i, tmpItem.序号) = "", "#", msh(intCurID).TextMatrix(i, tmpItem.序号)) & "^" & IIF(msh(intCurID).CellFontBold, 1, 0) & "^" & msh(intCurID).CellForeColor
                Else
                    tmpItem.表头 = tmpItem.表头 & IIF(tmpItem.表头 = "", "", "|") & arrHead(i)
                End If
            Next
        Next
        Call SetCopyGrid(intCurID)
        BlnSave = False
    End If
End Sub

Private Sub mnuEdit_AddFormat_Click()
    cmdAdd_Click
End Sub

Private Sub mnuEdit_Copy_Click()
'功能：将当前选择的报表元素复制到剪贴板对象(objClip)
'说明：
'     1.自动调整X,Y坐标
'     2.分栏索引不加入,粘贴时再处理
'     3.各个元素索引不处理(包括子项索引),粘贴时再处理
    Dim tmpObj As PictureBox, tmpItem As RPTItem, tmpID As RelatID
    Dim tmpItem1 As RPTItem
    
    If GetSelNum = 0 Then PlayWarn: Exit Sub
    
    Set objClip = New RPTItems
    For Each tmpObj In lblSize
        If tmpObj.Index Mod 8 = 1 Then
            With objReport.Items("_" & tmpObj.Tag)
                Set tmpItem = objClip.Add(.id, .格式号, .名称, .上级ID, .类型, .序号, .参照, .性质, .内容, .表头, .X, .Y, .W, .H, _
                .行高, .对齐, .自调, .字体, .字号, .粗体, .下线, .斜体, .网格, .前景, .背景, .边框, .分栏, .排序, .格式, .汇总, .表格线加粗, _
                .自适应行高, .图片, .系统, .父ID, .SubIDs, .CopyIDs, "_" & .id, .数据源, .上下间距, .左右间距, .源行号, .横向分栏, .纵向分栏)
                '处理子项
                If .SubIDs.count > 0 Then
                    For Each tmpID In .SubIDs
                        With objReport.Items("_" & tmpID.id)
                            objClip.Add .id, .格式号, .名称, .上级ID, .类型, .序号, .参照, .性质, .内容, .表头, .X, .Y, .W, .H, _
                                .行高, .对齐, .自调, .字体, .字号, .粗体, .下线, .斜体, .网格, .前景, .背景, .边框, .分栏, .排序, _
                                .格式, .汇总, .表格线加粗, .自适应行高, .图片, .系统, .父ID, .SubIDs, .CopyIDs, "_" & .id, .数据源, _
                                .上下间距, .左右间距, .源行号, .横向分栏, .纵向分栏
                        End With
                    Next
                End If
                If tmpItem.类型 = "14" Then
                    For Each tmpItem1 In objReport.Items
                        If tmpItem1.父ID = tmpItem.id Then
                            With tmpItem1
                                objClip.Add .id, .格式号, .名称, .上级ID, .类型, .序号, .参照, .性质, .内容, .表头, .X, .Y, .W, .H, _
                                    .行高, .对齐, .自调, .字体, .字号, .粗体, .下线, .斜体, .网格, .前景, .背景, .边框, .分栏, .排序, _
                                    .格式, .汇总, .表格线加粗, .自适应行高, .图片, .系统, .父ID, .SubIDs, .CopyIDs, "_" & .id, .数据源, _
                                    .上下间距, .左右间距, .源行号, .横向分栏, .纵向分栏
                            End With
                        End If
                    Next
                End If
                '处理分栏
                Set tmpItem.CopyIDs = New RelatIDs
            End With
        End If
    Next
    
    Call SelClear
End Sub

Private Sub mnuEdit_Del_Click()
    Dim strKey As String, tmpItem As RPTItem
    Dim blnDo As Boolean, tmpMain As RPTItem, tmpID As RelatID
    
    If tvwSQL.Nodes.count = 1 Then
        MsgBox "当前没有数据源可以删除！", vbInformation, App.Title: Exit Sub
    End If
    If tvwSQL.SelectedItem.Key = "Root" Then
        MsgBox "请选择要删除的数据源！", vbInformation, App.Title: Exit Sub
    End If
    
    If tvwSQL.SelectedItem.Parent.Key <> "Root" Then
        strKey = tvwSQL.SelectedItem.Parent.Key
    Else
        strKey = tvwSQL.SelectedItem.Key
    End If
    
    '检查报表元素中是否使用了该数据源,否则不能删除
    For Each tmpItem In objReport.Items
        If tmpItem.类型 = 5 And tmpItem.内容 = mdlPublic.GetStdNodeText(tvwSQL.Nodes(strKey).Text) Then  '汇总表格内容
            MsgBox "在报表中发现有汇总表格使用了该数据源,不能删除！", vbInformation, App.Title: Exit Sub
        ElseIf tmpItem.类型 = 6 And InStr(tmpItem.内容, mdlPublic.GetStdNodeText(tvwSQL.Nodes(strKey).Text) & ".") > 0 Then  '任意报表列
            MsgBox "在报表中发现有任意表格使用了该数据源的数据项,不能删除！", vbInformation, App.Title: Exit Sub
        ElseIf tmpItem.类型 = 3 And InStr(tmpItem.内容, mdlPublic.GetStdNodeText(tvwSQL.Nodes(strKey).Text) & ".") > 0 Then  '数据标签
            MsgBox "在报表中发现有数据标签使用了该数据源的数据项,不能删除！", vbInformation, App.Title: Exit Sub
        ElseIf tmpItem.类型 = 12 And InStr(tmpItem.内容, mdlPublic.GetStdNodeText(tvwSQL.Nodes(strKey).Text) & ".") > 0 Then  '图表
            MsgBox "在报表中发现有数据图表使用了该数据源的数据项,不能删除！", vbInformation, App.Title: Exit Sub
        End If
    Next
    
    If MsgBox("确实要删除数据源 " & tvwSQL.Nodes(strKey).Text & " 吗？", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
    
    On Error Resume Next
    If objClip.count > 0 Then
        For Each tmpItem In objClip
            If tmpItem.类型 = 5 And tmpItem.内容 = mdlPublic.GetStdNodeText(tvwSQL.Nodes(strKey).Text) Then  '汇总表格内容
                For Each tmpID In tmpItem.SubIDs
                    objClip.Remove "_" & tmpID.id
                Next
                objClip.Remove "_" & tmpItem.id
                blnDo = True
            ElseIf tmpItem.类型 = 6 And InStr(tmpItem.内容, mdlPublic.GetStdNodeText(tvwSQL.Nodes(strKey).Text) & ".") > 0 Then  '任意报表列
                Set tmpMain = objClip("_" & tmpItem.上级ID)
                For Each tmpID In tmpMain.SubIDs
                    objClip.Remove "_" & tmpID.id
                Next
                objClip.Remove "_" & tmpMain.id
                blnDo = True
            ElseIf tmpItem.类型 = 3 And InStr(tmpItem.内容, mdlPublic.GetStdNodeText(tvwSQL.Nodes(strKey).Text) & ".") > 0 Then  '数据标签
                objClip.Remove "_" & tmpItem.id
                blnDo = True
            ElseIf tmpItem.类型 = 12 And InStr(tmpItem.内容, mdlPublic.GetStdNodeText(tvwSQL.Nodes(strKey).Text) & ".") > 0 Then  '图表
                objClip.Remove "_" & tmpItem.id
                blnDo = True
            End If
        Next
    End If
    On Error GoTo 0
        
    If blnDo Then MsgBox "系统剪贴板中发现有使用了该数据源的元素，已经自动清除这些元素！", vbInformation, App.Title
    
    objReport.Datas.Remove tvwSQL.Nodes(strKey).Key
    tvwSQL.Nodes.Remove tvwSQL.Nodes(strKey).Key
    
    tvwSQL.Nodes(1).Selected = True
    tvwSQL_NodeClick tvwSQL.SelectedItem
    
    BlnSave = False
End Sub

Private Sub mnuEdit_DelFormat_Click()
    cmdDel_Click
End Sub

Private Sub mnuFile_Guide_Click()
    Dim tmpItem As RPTItem, tmpData As RPTData, tmpID As RelatID
    
    Set frmGuide.frmParent = Me
    Set frmGuide.objReport = objReport
    Set frmGuide.mobjFmt = objReport.Fmts("_" & mbytCurrFmt)
    frmGuide.Show 1, Me
    
    If gblnOK Then
        If frmGuide.objGuide.编号 = "" Then
            '清除现有内容
            Set objReport.Items = New RPTItems
            Set objReport.Datas = New RPTDatas
            Set objReport.Fmts = New RPTFmts
            Set objReport.Items = frmGuide.objGuide.Items
            Set objReport.Datas = frmGuide.objGuide.Datas
            Set objReport.Fmts = frmGuide.objGuide.Fmts
            mbytCurrFmt = 1
        Else
            '加入现有内容
            For Each tmpData In frmGuide.objGuide.Datas
                With tmpData
                    objReport.Datas.Add .名称, .数据连接编号, .SQL, .字段, .对象, .类型, .说明, .Pars, "_" & .名称
                End With
            Next
            For Each tmpItem In frmGuide.objGuide.Items
                With tmpItem
                    objReport.Items.Add .id, mbytCurrFmt, .名称, .上级ID, .类型, .序号, .参照, .性质, .内容, .表头 _
                        , .X, .Y, .W, .H, .行高, .对齐, .自调, .字体, .字号, .粗体, .下线, .斜体, .网格, .前景, .背景 _
                        , .边框, .分栏, .排序, .格式, .汇总, .表格线加粗, .自适应行高, .图片, .系统, .父ID, .SubIDs _
                        , .CopyIDs, "_" & .id, .数据源, .上下间距, .左右间距, .源行号, .横向分栏, .纵向分栏
                End With
            Next
        End If
        
        Unload frmGuide
        
        '重新计算最大控件索引
        intMaxID = 0
        For Each tmpItem In objReport.Items
            If tmpItem.id > intMaxID Then intMaxID = tmpItem.id
            '注意分栏
            For Each tmpID In tmpItem.CopyIDs
                If tmpID.id > intMaxID Then intMaxID = tmpID.id
            Next
        Next
        
        Set objLastSel = Nothing: intCurID = 0
        
        Call ReFlashReport
        Call LoadReportFormat
        BlnSave = False
    End If
End Sub

Private Sub mnuEdit_Inverse_Click()
'功能：返向选择报表元素控件
    Dim tmpItem As RPTItem, ObjSel As Object
    Me.MousePointer = 11
    For Each tmpItem In objReport.Items
        If tmpItem.格式号 = Mid(cboFormat.ComboItems("_" & mbytCurrFmt).Key, 2) Then
            If InStr(1, "|1,|2,|3,|4,|5,|10,|11,|12,|14,", "|" & tmpItem.类型) <> 0 Then
                Set ObjSel = GetInxObj(tmpItem.id)
                
                If ObjSel.Tag = "" Then
                    SelItem tmpItem.id, True
                Else
                    SelItem tmpItem.id, False
                End If
            End If
        End If
    Next
    Call ShowAttrib
    If GetSelNum = 1 Then
        Call ShowAttrib(intCurID)
    End If
    Me.MousePointer = 0
End Sub

Private Sub mnuEdit_ItemAdd_Click(Index As Integer)
    Dim curNode As Object, objNode As Object
    Dim i As Integer, j As Integer
    
    '虚拟动作
    selArea.Left = 500: selArea.Top = 300
    Select Case Index
        Case 0 '线条
            bytCurTool = 1
            selArea.Bottom = selArea.Top
            selArea.Right = IIF(picPaper.Width < picBack.Width, picPaper.Width, picBack.Width) - selArea.Left
        Case 1, 3 '框线,标签
            bytCurTool = IIF(Index = 1, 2, Index)
            selArea.Bottom = selArea.Top + lbl(0).Height
            selArea.Right = selArea.Left + 2000
            If Index = 1 Then selArea.Bottom = selArea.Bottom + 1000
        Case 4
            bytCurTool = Index
            selArea.Right = selArea.Left + 1500
            selArea.Bottom = selArea.Top + 1300
        Case 5 '表格
            bytCurTool = Index
            selArea.Right = selArea.Left + 5000
            selArea.Bottom = selArea.Top + 5000
        Case 6 '图表@@@
            bytCurTool = Index
            selArea.Right = selArea.Left + 4500
            selArea.Bottom = selArea.Top + 3000
        Case 7 '条码
            bytCurTool = Index
            selArea.Right = selArea.Left + 2500
            selArea.Bottom = selArea.Top + 1100
    End Select
    Call AddReportItem
    bytCurTool = 0
    
    BlnSave = False
End Sub

Private Sub mnuEdit_New_Click()
    Dim objData As RPTData
    If frmSQLEdit.ShowMe(Me, IIF(glngSys <> 0, glngSys, objReport.系统), objData, objReport.Datas, 0) Then
        With objData
            objReport.Datas.Add .名称, .数据连接编号, .SQL, .字段, .对象, .类型, .说明, .Pars, "_" & .Key
        End With
        Call tvwSQL_NodeClick(tvwSQL.SelectedItem)
        Unload frmSQLEdit
        BlnSave = False
    End If
End Sub

Private Sub mnuEdit_Modi_Click()
    Dim strKey As String, strInfo As String
    Dim strPreName As String, strNewName As String
    Dim strDBName As String  '数据库中的名称
    Dim objData As RPTData
    
    If tvwSQL.Nodes.count = 1 Then
        MsgBox "当前没有数据源可以修改,请先新增数据源！", vbInformation, App.Title: Exit Sub
    End If
    If tvwSQL.SelectedItem.Key = "Root" Then
        MsgBox "请选择要修改的数据源！", vbInformation, App.Title: Exit Sub
    End If
    
    If tvwSQL.SelectedItem.Parent.Key <> "Root" Then
        strKey = tvwSQL.SelectedItem.Parent.Key
    Else
        strKey = tvwSQL.SelectedItem.Key
    End If
    strPreName = objReport.Datas(strKey).名称
    
    Set objData = objReport.Datas(strKey)
    strDBName = objReport.Datas(strKey).原名称
    
    If frmSQLEdit.ShowMe(Me, IIF(glngSys <> 0, glngSys, objReport.系统), objData, objReport.Datas, 0) Then
        objReport.Datas.Remove strKey
        With objData
            objReport.Datas.Add .名称, .数据连接编号, .SQL, .字段, .对象, .类型, .说明, .Pars, "_" & .Key
            strNewName = .名称
        End With
        Call tvwSQL_NodeClick(tvwSQL.SelectedItem)
        
        '如果数据源名称更改,则更改涉及的报表元素的相应内容(字段名变更则无法处理)
        If strPreName <> strNewName Then
            Call ReplaceName(strPreName, strNewName)
            '如果是第一次修改，则给原名称赋值
            If strDBName = "" Then
                objReport.Datas("_" & objData.Key).原名称 = strPreName
            Else
                objReport.Datas("_" & objData.Key).原名称 = strDBName
            End If
            If GetSelNum = 1 Then Call ShowAttrib(intCurID)
        End If
        
        '检查数据对应关系
        Me.Refresh
        strInfo = CheckData
        If strInfo <> "" Then
            MsgBox strInfo, vbInformation, App.Title
        End If
        Unload frmSQLEdit
        BlnSave = False
    End If
End Sub

Private Sub mnuEdit_Paste_Click()
    Dim tmpCopy As RPTItem, tmpItem As RPTItem, tmpID As RelatID
    Dim i As Integer, j As Integer, strName As String, tmpChange As RPTItem
    Dim RectTest As RECT
    Dim Col As New Collection
    Dim objClipTmp As RPTItems '剪贴板对象
    Dim lng父ID As Integer, lng父IDTmp As Long
    Dim blnSouse As Boolean
    Dim k As Long, X As Long, Y As Long
    Dim strSouse As String
    Dim lngMinusX As Long, lngMinusY As Long, strCardIDs As String
    
    If objClip.count = 0 Then PlayWarn: Exit Sub
    On Error Resume Next
    If lblSize.count = 9 Then
        If objReport.Items("_" & Val(lblSize(1).Tag)).类型 = "14" Then
            lng父IDTmp = objReport.Items("_" & Val(lblSize(1).Tag)).id
        End If
    End If
    Err.Clear: On Error GoTo 0
    Call SelClear
    
    '只粘贴一个控件，则清除参照对象
    If objClip.count = 1 Then objClip(1).参照 = "": objClip(1).性质 = 0
    
    '粘贴多个控件,无主表,则清除参照及性质
    Call CheckClip
    '排序：非子元素的先加入
    Set objClipTmp = New RPTItems
    For j = 1 To objClip.count
        Set tmpCopy = objClip(j)
        If tmpCopy.类型 = 5 Then
            
        End If
        If tmpCopy.父ID = 0 Then
            With tmpCopy
                objClipTmp.Add .id, .格式号, .名称, .上级ID, .类型, .序号, .参照, .性质, .内容, .表头, .X, .Y, .W, .H, _
                .行高, .对齐, .自调, .字体, .字号, .粗体, .下线, .斜体, .网格, .前景, .背景, .边框, .分栏, .排序, .格式, _
                .汇总, .表格线加粗, .自适应行高, .图片, .系统, .父ID, .SubIDs, .CopyIDs, "_" & .id, .数据源, .上下间距, _
                .左右间距, .源行号, .横向分栏, .纵向分栏
                If .类型 = 14 Then strCardIDs = strCardIDs & "," & .id
            End With
        End If
        If lng父IDTmp <> 0 And tmpCopy.上级ID = 0 Then
            If tmpCopy.类型 = 14 Then
                MsgBox "卡片元素不允许复制到卡片元素中去。", vbInformation, App.Title
                Exit Sub
            ElseIf tmpCopy.类型 = 4 And tmpCopy.分栏 > 1 Then
                MsgBox "卡片中不允许放入分栏的表格。", vbInformation, App.Title
                Exit Sub
            End If
            If lngMinusX > tmpCopy.X Or lngMinusX = 0 Then lngMinusX = tmpCopy.X
            If lngMinusY > tmpCopy.Y Or lngMinusY = 0 Then lngMinusY = tmpCopy.Y
        End If
    Next
    strCardIDs = Mid(strCardIDs, 2)
    For j = 1 To objClip.count
        Set tmpCopy = objClip(j)
        If tmpCopy.父ID <> 0 Then
            With tmpCopy
                If lng父IDTmp = 0 Then
                    If InStr("," & strCardIDs & ",", "," & tmpCopy.父ID & ",") = 0 Then
                        .X = .X + pic(.父ID).Left
                        .Y = .Y + pic(.父ID).Top
                    Else
                        .X = .X - 200  '后面要加两百，所以保持不变
                        .Y = .Y - 200
                    End If
                End If
                objClipTmp.Add .id, .格式号, .名称, .上级ID, .类型, .序号, .参照, .性质, .内容, .表头, .X, .Y, .W, .H, _
                    .行高, .对齐, .自调, .字体, .字号, .粗体, .下线, .斜体, .网格, .前景, .背景, .边框, .分栏, .排序, _
                    .格式, .汇总, .表格线加粗, .自适应行高, .图片, .系统, .父ID, .SubIDs, .CopyIDs, "_" & .id, .数据源, _
                    .上下间距, .左右间距, .源行号, .横向分栏, .纵向分栏
            End With
        End If
    Next
    Set objClip = objClipTmp
    For j = 1 To objClip.count
        '存在参照对象的元素,如果发生改变名称,则循环改变其子表
        Set tmpCopy = objClip(j)
        If InStr(1, "|1,|2,|3,|4,|5,|10,|11,|12,|14,", "|" & tmpCopy.类型 & ",") <> 0 Then
            strName = GetNextName(tmpCopy.类型, True)
            If strName <> tmpCopy.名称 Then
                For Each tmpChange In objClip
                    If tmpChange.参照 = tmpCopy.名称 And InStr(1, "4,5", tmpChange.类型) > 0 Then tmpChange.参照 = strName
                    If tmpChange.内容 Like "标签*" And tmpChange.类型 = 2 Then tmpChange.内容 = strName
                Next
                tmpCopy.名称 = strName
            End If
        End If
    Next
    
    For j = 1 To objClip.count
        '存在参照对象的元素,如果发生改变名称,则循环改变其子表
        Set tmpCopy = objClip(j)
        blnSouse = False
        If tmpCopy.类型 = "4" And lng父IDTmp <> 0 Then
            If objReport.Items("_" & lng父IDTmp).数据源 <> "" Then
                For Each tmpID In tmpCopy.SubIDs
                    With objReport.Items("_" & tmpID.id)
                        X = InStr(1, .内容, "]")
                        Y = InStr(1, .内容, ".")
                        k = InStr(1, .内容, "[")
                        If X > k And X > Y And X <> 0 And k <> 0 And Y <> 0 Then
                            If Mid(.内容, k + 1, Y - k - 1) <> objReport.Items("_" & lng父IDTmp).数据源 Then
                                strSouse = strSouse & "," & tmpCopy.名称
                                blnSouse = True
                                Exit For
                            End If
                        End If
                    End With
                Next
            End If
        End If
        If blnSouse = False Then
            If InStr(1, "|1,|2,|3,|4,|5,|10,|11,|12,|13,|14,", "|" & tmpCopy.类型 & ",") <> 0 Then
                intMaxID = intMaxID + 1
                With tmpCopy
                    RectTest.Left = .X
                    RectTest.Top = .Y
                    RectTest.Right = .W
                    RectTest.Bottom = .H
                    lng父ID = 0
                    If .父ID <> 0 Then
                        On Error Resume Next
                        lng父ID = Val(Col("_" & .父ID) & "")
                        Err.Clear: On Error GoTo 0
                    End If
                    If lng父IDTmp <> 0 And InStr(",5,12,14,", "," & tmpCopy.类型 & ",") = 0 Then
                        lng父ID = lng父IDTmp
                    End If
                    If .格式号 = mbytCurrFmt Then
                        If .类型 = 2 Or .类型 = 12 Then '@@@
                            '在高度上变化,为避免参照的X坐标发生改变
                            If .参照 <> "" Then
                                Select Case Mid(.性质, 1, 1)
                                Case 1  '表上项
                                    If Not ((.Y - .H - 200) * sgnMode < 100) Then RectTest.Top = RectTest.Top - 200
                                    If .系统 Then Call GetCoordinate(RectTest)
                                    Set tmpItem = objReport.Items.Add(intMaxID, mbytCurrFmt, .名称, 0, .类型, .序号, _
                                        .参照, .性质, .内容, .表头, RectTest.Left, RectTest.Top, RectTest.Right, RectTest.Bottom, _
                                        .行高, .对齐, .自调, .字体, .字号, .粗体, .下线, .斜体, .网格, .前景, .背景, .边框, .分栏, _
                                        .排序, .格式, .汇总, .表格线加粗, .自适应行高, .图片, .系统, lng父ID, , , "_" & intMaxID, .数据源, _
                                        .上下间距, .左右间距, .源行号, .横向分栏, .纵向分栏)
                                Case Else
                                    If (.Y + .H + 200) * sgnMode < picPaper.Height - 100 Then
                                        RectTest.Top = RectTest.Top + RectTest.Bottom + 200
                                    End If
                                    If .系统 Then Call GetCoordinate(RectTest)
                                    Set tmpItem = objReport.Items.Add(intMaxID, mbytCurrFmt, .名称, 0, .类型, .序号, _
                                        .参照, .性质, .内容, .表头, RectTest.Left, RectTest.Top, RectTest.Right, RectTest.Bottom, _
                                        .行高, .对齐, .自调, .字体, .字号, .粗体, .下线, .斜体, .网格, .前景, .背景, .边框, .分栏, _
                                        .排序, .格式, .汇总, .表格线加粗, .自适应行高, .图片, .系统, lng父ID, , , "_" & intMaxID, .数据源, _
                                        .上下间距, .左右间距, .源行号, .横向分栏, .纵向分栏)
                                End Select
                            Else
                                If .系统 Then
                                    Call GetCoordinate(RectTest)
                                Else
                                    RectTest.Top = RectTest.Top + 200
                                    RectTest.Left = RectTest.Left + 200
                                End If
                                Set tmpItem = objReport.Items.Add(intMaxID, mbytCurrFmt, .名称, 0, .类型, .序号, _
                                    .参照, .性质, .内容, .表头, RectTest.Left, RectTest.Top, RectTest.Right, RectTest.Bottom, _
                                    .行高, .对齐, .自调, .字体, .字号, .粗体, .下线, .斜体, .网格, .前景, .背景, .边框, .分栏, _
                                    .排序, .格式, .汇总, .表格线加粗, .自适应行高, .图片, .系统, lng父ID, , , "_" & intMaxID, .数据源, _
                                    .上下间距, .左右间距, .源行号, .横向分栏, .纵向分栏)
                            End If
                        Else
                            If .系统 Then
                                Call GetCoordinate(RectTest)
                            Else
                                RectTest.Top = RectTest.Top + 200
                                RectTest.Left = RectTest.Left + 200
                            End If
                            Set tmpItem = objReport.Items.Add(intMaxID, mbytCurrFmt, .名称, 0, .类型, .序号, _
                                .参照, .性质, .内容, .表头, RectTest.Left, RectTest.Top, RectTest.Right, RectTest.Bottom, _
                                .行高, .对齐, .自调, .字体, .字号, .粗体, .下线, .斜体, .网格, .前景, .背景, .边框, .分栏, _
                                .排序, .格式, .汇总, .表格线加粗, .自适应行高, .图片, .系统, lng父ID, , , "_" & intMaxID, .数据源, _
                                .上下间距, .左右间距, .源行号, .横向分栏, .纵向分栏)
                        End If
                    Else
                        '坐标与原格式下的坐标一致
                        Set tmpItem = objReport.Items.Add(intMaxID, mbytCurrFmt, .名称, 0, .类型, .序号, _
                            .参照, .性质, .内容, .表头, RectTest.Left, RectTest.Top, RectTest.Right, RectTest.Bottom, _
                            .行高, .对齐, .自调, .字体, .字号, .粗体, .下线, .斜体, .网格, .前景, .背景, .边框, .分栏, _
                            .排序, .格式, .汇总, .表格线加粗, .自适应行高, .图片, .系统, lng父ID, , , "_" & intMaxID, .数据源, _
                            .上下间距, .左右间距, .源行号, .横向分栏, .纵向分栏)
                    End If
                    If .父ID = 0 Then
                        Col.Add intMaxID, "_" & .id
                    End If
                    If lng父ID <> 0 Then
                        tmpItem.X = tmpItem.X - lngMinusX
                        tmpItem.Y = tmpItem.Y - lngMinusY
                    End If
                    '处理子项
                    If (.类型 = 4 Or .类型 = 5) And .SubIDs.count > 0 Then
                        For Each tmpID In .SubIDs
                            intMaxID = intMaxID + 1
                            With objClip("_" & tmpID.id)
                                objReport.Items.Add intMaxID, mbytCurrFmt, strName, tmpItem.id, .类型, .序号, _
                                    .参照, .性质, .内容, .表头, .X + 300, .Y + 300, .W, .H, .行高, .对齐, .自调, _
                                    .字体, .字号, .粗体, .下线, .斜体, .网格, .前景, .背景, .边框, .分栏, .排序, _
                                    .格式, .汇总, .表格线加粗, .自适应行高, .图片, .系统, lng父ID, , , "_" & intMaxID, _
                                    .数据源, .上下间距, .左右间距, .源行号, .横向分栏, .纵向分栏
                                tmpItem.SubIDs.Add intMaxID, "_" & intMaxID
                            End With
                        Next
                    End If
                
                    '处理分栏
                    If .类型 = 4 And .分栏 > 1 Then
                        For i = 1 To .分栏 - 1
                            intMaxID = intMaxID + 1
                            tmpItem.CopyIDs.Add intMaxID, "_" & intMaxID
                        Next
                    End If
                End With
                Call ShowItem(tmpItem.id)
                Call SelItem(tmpItem.id, True)
            End If
        End If
    Next
    
    If strSouse <> "" Then
        MsgBox "表格：" & Mid(strSouse, 2) & " 中的列绑定的数据不是本卡片数据源中的数据，不能放入此卡片中。", vbInformation, App.Title
    End If
    
    If objClip.count = 1 Then
        Call ShowAttrib(tmpItem.id)
    End If
    Set objClip = New RPTItems
    
    BlnSave = False
End Sub

Private Sub mnuEdit_Remove_Click()
'功能：删除当前选择的报表元素(一个或多个)
    Dim tmpObj As PictureBox, tmpID As RelatID, ItemThis As RPTItem
    Dim objControl As Object, i As Long
    Dim tmpObj1 As PictureBox, blntmp As Boolean
    
    Select Case GetSelNum
    Case 0
        MsgBox "没有选择任何报表元素,无法删除！", vbInformation, App.Title: Exit Sub
    Case 1
        If objReport.Items("_" & intCurID).系统 Then
            MsgBox "当前选择的报表元素是系统固有元素,无法删除！", vbInformation, App.Title: Exit Sub
        End If
    End Select
    
    If MsgBox("共选择了 " & GetSelNum & " 个元素,确实要删除吗？", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
Res:
    For Each tmpObj In lblSize
        If lblSize.count > 1 Then
            On Error Resume Next
            If tmpObj.Index Mod 8 = 1 Then
                If blntmp = True Then blntmp = False: GoTo Res
                    If Not objReport.Items("_" & tmpObj.Tag).系统 Then
                        Select Case objReport.Items("_" & tmpObj.Tag).类型
                            Case 1
                                Unload lblLine(tmpObj.Tag)
                                objReport.Items.Remove "_" & tmpObj.Tag
                            Case 2, 3
                                Unload lbl(tmpObj.Tag)
                                objReport.Items.Remove "_" & tmpObj.Tag
                            Case 10
                                Unload Shp(tmpObj.Tag)
                                Unload lblshp(tmpObj.Tag)
                                objReport.Items.Remove "_" & tmpObj.Tag
                            Case 11
                                Unload img(tmpObj.Tag)
                                objReport.Items.Remove "_" & tmpObj.Tag
                            Case 14
                                For Each objControl In Me.Controls
                                     If InStr(";ImageList;CommonDialog;Menu;PictureBox;", ";" & TypeName(objControl) & ";") = 0 Then
                                        If objControl.Container Is pic(tmpObj.Tag) Then
                                            '删除子控件的位置点，再删除控件，然后重新循环
                                            For Each tmpObj1 In lblSize
                                                If tmpObj1.Index Mod 8 = 1 Then
                                                    If objReport.Items("_" & tmpObj1.Tag).id = objControl.Index Then
                                                        For i = tmpObj1.Index To tmpObj1.Index + 7
                                                            Unload lblSize(i)
                                                        Next
                                                        blntmp = True
                                                        Exit For
                                                    End If
                                                End If
                                            Next
                                            If objReport.Items("_" & objControl.Index).类型 = 4 Then
                                                '移除子项对象
                                                For Each tmpID In objReport.Items("_" & objControl.Index).SubIDs
                                                    objReport.Items.Remove "_" & tmpID.id
                                                Next
                                            End If
                                            objReport.Items.Remove "_" & objControl.Index
                                            Unload objControl
                                        End If
                                    End If
                                Next
                                Unload pic(tmpObj.Tag)
                                objReport.Items.Remove "_" & tmpObj.Tag
                            Case 4, 5
                                '如果删除的是一个子表（附加表格或左联接表格），则调整其余子表的位置
                                If objReport.Items("_" & tmpObj.Tag).参照 <> "" Then
                                    For Each ItemThis In objReport.Items
                                        If ItemThis.格式号 = mbytCurrFmt And ItemThis.名称 = objReport.Items("_" & tmpObj.Tag).参照 Then
                                            If InStr(1, "4,5", ItemThis.类型) <> 0 Then objReport.Items("_" & tmpObj.Tag).参照 = "": SetChildWH (ItemThis.Key)
                                        End If
                                    Next
                                Else
                                    For Each ItemThis In objReport.Items
                                        If ItemThis.格式号 = mbytCurrFmt And ItemThis.参照 = objReport.Items("_" & tmpObj.Tag).名称 Then ItemThis.参照 = "": ItemThis.性质 = 0
                                    Next
                                End If
                                
                                Unload msh(tmpObj.Tag)
                                '移除分栏控件
                                For Each tmpID In objReport.Items("_" & tmpObj.Tag).CopyIDs
                                    Unload msh(tmpID.id)
                                Next
                                '移除子项对象
                                For Each tmpID In objReport.Items("_" & tmpObj.Tag).SubIDs
                                    objReport.Items.Remove "_" & tmpID.id
                                Next
                                objReport.Items.Remove "_" & tmpObj.Tag
                            Case 12 '@@@
                                Unload Chart(tmpObj.Tag)
                                objReport.Items.Remove "_" & tmpObj.Tag
                            Case 13
                                Unload ImgCode(tmpObj.Tag)
                                objReport.Items.Remove "_" & tmpObj.Tag
                        End Select
                    Else
                        Select Case objReport.Items("_" & tmpObj.Tag).类型
                            Case 1
                                lblLine(tmpObj.Tag).Tag = ""
                            Case 2, 3
                                lbl(tmpObj.Tag).Tag = ""
                            Case 10
                                Shp(tmpObj.Tag).Tag = ""
                            Case 11
                                img(tmpObj.Tag).Tag = ""
                            Case 4, 5
                                msh(tmpObj.Tag).Tag = ""
                            Case 12 '@@@
                                Chart(tmpObj.Tag).Tag = ""
                            Case 13
                                ImgCode(tmpObj.Tag).Tag = ""
                            Case 14
                                pic(tmpObj.Tag).Tag = ""
                        End Select
                    End If
            End If
            If tmpObj.Index <> 0 Then
                Unload lblSize(tmpObj.Index)
            End If
        End If
    Next
    
    Set objLastSel = Nothing: intCurID = 0
    
    picPaper.SetFocus
    Call picPaper_MouseDown(1, 0, 0, 0)
    BlnSave = False
End Sub

Private Function CheckPars() As String
'功能：检查参数中的数据源绑定参数的情况是否正确
    Dim objData As RPTData, objPar As RPTPar
    Dim strSQL As String, strParName As String
    
    For Each objData In objReport.Datas
        For Each objPar In objData.Pars
            If objPar.明细SQL <> "" Then
                strSQL = objPar.明细SQL
                If CheckParsRela(strSQL, objReport.Datas, objPar.名称, , , , strParName) = False Then
                    CheckPars = "数据源[" & objData.名称 & "]中的参数[" & objPar.名称 & "]的明细SQL中绑定的参数[" & strParName & "]未保存或绑定的就是当前参数，请检查。"
                    Exit Function
                End If
            End If
            If objPar.分类SQL <> "" Then
                strSQL = objPar.分类SQL
                If CheckParsRela(strSQL, objReport.Datas, objPar.名称, , , , strParName) = False Then
                    CheckPars = "数据源[" & objData.名称 & "]中的参数[" & objPar.名称 & "]的分类SQL中绑定的参数[" & strParName & "]未保存或绑定的就是当前参数，请检查。"
                    Exit Function
                End If
            End If
        Next
    Next
    
End Function

Private Sub mnuFile_Save_Click()
    Dim strInfo As String, LngItemKey As Long, i As Integer
    
    For i = 1 To cboFormat.ComboItems.count
        If InStr(cboFormat.ComboItems(i).Text, "'") > 0 Then
            MsgBox "第 " & i & " 个报表格式名中输入了非法字符，请检查！", vbInformation, App.Title
            Exit Sub
        End If
    Next
    
    '检查数据源
    strInfo = CheckData
    If strInfo <> "" Then MsgBox strInfo, vbInformation, App.Title: Exit Sub
    
    '检查表头文字
    strInfo = CheckHead
    If strInfo <> "" Then MsgBox strInfo, vbInformation, App.Title: Exit Sub
    
    strInfo = CheckArea
    If strInfo <> "" Then MsgBox strInfo, vbInformation, App.Title: Exit Sub
    
    strInfo = CheckPars
    If strInfo <> "" Then MsgBox strInfo, vbInformation, App.Title: Exit Sub
    
    LngItemKey = CheckCoordinate
    If LngItemKey <> 0 Then
        '检查系统项是否被遮住，是则不允许保存
        MsgBox "系统固有元素[" & objReport.Items("_" & LngItemKey).名称 & "]被其它元素遮住，保存中止！", vbInformation, App.Title
        Exit Sub
    End If
    
    Call SelClear
    Refresh
    If Not SaveReport(lngRPTID, objReport, sta.Panels(2)) Then
        MsgBox "报表保存失败,请重试保存操作！", vbInformation, App.Title
        Exit Sub
    End If
    Call UpdatePriv
    
    gblnModi = True
    BlnSave = True
    
    If Not CheckReportPriv(lngRPTID) Then
        MsgBox "你没有权限查询该报表某些数据源中的对象，虽然可以正常" & vbCrLf & _
               "地保存，但在你修正这些问题之前你不能正常使用该报表！", vbInformation, App.Title
    End If
End Sub

Private Sub mnuEdit_SelAll_Click()
'功能：选择全部报表元素控件
    Dim tmpItem As RPTItem
    
    Me.MousePointer = 11
    For Each tmpItem In objReport.Items
        If tmpItem.格式号 = Mid(cboFormat.ComboItems("_" & mbytCurrFmt).Key, 2) Then
            If InStr(1, "|1,|2,|3,|4,|5,|10,|11,|12,|13,|14,", "|" & tmpItem.类型) <> 0 Then '@@@
                SelItem tmpItem.id, True
            End If
        End If
    Next
    Call ShowAttrib
    If GetSelNum = 1 Then
        Call ShowAttrib(intCurID)
    End If
    Me.MousePointer = 0
End Sub

Private Sub mnuFile_Page_Click()
    If Printers.count = 0 Then
        MsgBox "在系统中没有检测到任何打印设备,请先安装打印机后再重试该操作！" & vbCrLf & _
            "在你添加新的打印机之前,系统将按缺省纸张进行设置。", vbInformation, App.Title
        Exit Sub
    End If
    
    With frmPageSetup
        .strPrinter = objReport.打印机
        .intBin = objReport.进纸
        .intPage = objReport.Fmts("_" & mbytCurrFmt).纸张
        .lngWidth = objReport.Fmts("_" & mbytCurrFmt).W
        .lngHeight = objReport.Fmts("_" & mbytCurrFmt).H
        .bytOrient = objReport.Fmts("_" & mbytCurrFmt).纸向
    End With
    frmPageSetup.Show 1, Me
    If gblnOK Then
        With frmPageSetup
            objReport.打印机 = .strPrinter
            objReport.进纸 = .intBin
            objReport.Fmts("_" & mbytCurrFmt).纸张 = .intPage
            objReport.Fmts("_" & mbytCurrFmt).W = .lngWidth '当非自定义纸张时,也要存值
            objReport.Fmts("_" & mbytCurrFmt).H = .lngHeight
            objReport.Fmts("_" & mbytCurrFmt).纸向 = .bytOrient
            If objReport.Fmts("_" & mbytCurrFmt).纸向 = 2 Then
                objReport.Fmts("_" & mbytCurrFmt).动态纸张 = False
            End If
        End With
        Unload frmPageSetup
        Call ShowSize: Call ShowScroll: Call GetInPaper
        BlnSave = False
    End If
End Sub

Private Sub mnuFile_Quit_Click()
    Unload Me
End Sub

Private Sub mnuFile_Report_Click()
    Dim strInfo As String, LngItemKey As Long

    strInfo = CheckData
    If strInfo <> "" Then MsgBox strInfo, vbInformation, App.Title: Exit Sub
    
    strInfo = CheckHead
    If strInfo <> "" Then MsgBox strInfo, vbInformation, App.Title: Exit Sub
    
    strInfo = CheckArea
    If strInfo <> "" Then MsgBox strInfo, vbInformation, App.Title: Exit Sub
    
    strInfo = CheckPars
    If strInfo <> "" Then MsgBox strInfo, vbInformation, App.Title: Exit Sub
    
    LngItemKey = CheckCoordinate
    If LngItemKey <> 0 Then
        '检查系统项是否被遮住，是则不允许保存
        MsgBox "系统固有元素[" & objReport.Items("_" & LngItemKey).名称 & "]被其它元素遮住，保存中止！", vbInformation, App.Title
        Exit Sub
    End If
    
    If cboFormat.SelectedItem Is Nothing Then
        MsgBox "不能确定当前报表格式,请选定一种报表格式！", vbInformation, App.Title
        cboFormat.SetFocus: Exit Sub
    End If
    
    '执行报表
    glngGroup = 0
    CopyReport objReport, gobjReport
    garrPars = Array("ReportFormat=" & cboFormat.SelectedItem.Index) '强行使用当前格式
    If Not ShowReport(Me) Then MsgBox "报表打开失败！", vbInformation, App.Title
End Sub

Private Sub mnuFormat_Back_Click()
    Call SetLevel(1)
End Sub

Private Sub mnuFormat_DoAlign_Click(Index As Integer)
    If Index <= 5 Then
        Call SetSelAlign(Index + 1)
    Else
        If Index = 7 Then
            Call SetSelCenter(0)
        ElseIf Index = 8 Then
            Call SetSelCenter(1)
        End If
    End If
End Sub

Private Sub mnuFormat_Front_Click()
    Call SetLevel
End Sub

Private Sub SetLevel(Optional bytOrder As Byte)
'功能：按选择顺序设置选择控件的前后顺序
    Dim tmpObj As PictureBox, ObjSel As Object
    
    If GetSelNum = 0 Then Exit Sub
    
    For Each tmpObj In lblSize
        If tmpObj.Index Mod 8 = 1 Then
            Set ObjSel = GetInxObj(tmpObj.Tag)
        End If
    Next
    ObjSel.ZOrder bytOrder
End Sub

Private Sub mnuFormat_Height_Click()
    Call SetSelAlign(8)
End Sub

Private Sub mnuFormat_HscSpace_Add_Click()
    Call SetHscSpace(1)
End Sub

Private Sub mnuFormat_HscSpace_Dec_Click()
    Call SetHscSpace(-1)
End Sub

Private Sub mnuFormat_HscSpace_Same_Click()
    Call SetHscSpace(0)
End Sub

Private Sub mnuFormat_Lock_Click()
    mnuFormat_Lock.Checked = Not mnuFormat_Lock.Checked
    blnLock = Not blnLock
    If mnuFormat_Lock.Checked And tbr2.Buttons("Lock").Value = tbrUnpressed Then
        tbr2.Buttons("Lock").Value = tbrPressed
    ElseIf Not mnuFormat_Lock.Checked And tbr2.Buttons("Lock").Value = tbrPressed Then
        tbr2.Buttons("Lock").Value = tbrUnpressed
    End If
    Call SetLock(blnLock)
End Sub

Private Sub mnuFormat_VscSpace_Add_Click()
    Call SetVscSpace(1)
End Sub

Private Sub mnuFormat_VscSpace_Dec_Click()
    Call SetVscSpace(-1)
End Sub

Private Sub mnuFormat_VscSpace_Same_Click()
    Call SetVscSpace(0)
End Sub

Private Sub mnuFormat_WH_Click()
    Call SetSelAlign(9)
End Sub

Private Sub mnuFormat_Width_Click()
    Call SetSelAlign(7)
End Sub

Private Sub mnuHelpAbout_Click()
    Call ShowAbout(Me)
End Sub

Private Sub mnuHelpTitle_Click()
    Call ShowHelpRpt(Me.hwnd, "design", 0)
End Sub

Private Sub mnuView_reFlash_Click()
    If MsgBox("确实要刷新显示报表内容吗？", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
    Call ReFlashReport(True)
    Call LoadReportFormat
End Sub

Private Sub mnuViewScaleMode_Click(Index As Integer)
    Dim MenuSelect As Menu   '显示比例
    Dim tmpItem As RPTItem, ObjSel As Control
    Dim lngRow As Long
    'zyb#Add
    '按比例显示
    
    '取上次的显示比例
    For Each MenuSelect In Me.mnuViewScaleMode
        If MenuSelect.Checked Then
            Select Case MenuSelect.Index
            Case 0, 1, 2
                sgnLastMode = GetAutoTest(MenuSelect.Index)
            Case 4
                sgnLastMode = 2
            Case 5
                sgnLastMode = 1
            Case 6
                sgnLastMode = 0.75
            Case 7
                sgnLastMode = 0.5
            Case 8
                sgnLastMode = 0.25
            End Select
        End If
    Next
    
    '清除选择
    mnuViewScaleMode(0).Checked = False
    mnuViewScaleMode(1).Checked = False
    mnuViewScaleMode(2).Checked = False
    mnuViewScaleMode(4).Checked = False
    mnuViewScaleMode(5).Checked = False
    mnuViewScaleMode(6).Checked = False
    mnuViewScaleMode(7).Checked = False
    mnuViewScaleMode(8).Checked = False
    
    '获取显示比例并设置相应菜单
    Select Case Index
    Case 0, 1, 2
        mnuViewScaleMode(Index).Checked = True
        sgnMode = GetAutoTest(Index)
    Case 4  '200%
        mnuViewScaleMode(4).Checked = True
        sgnMode = 2
    Case 5  '100%
        mnuViewScaleMode(5).Checked = True
        sgnMode = 1
    Case 6  '75%
        mnuViewScaleMode(6).Checked = True
        sgnMode = 0.75
    Case 7  '50%
        mnuViewScaleMode(7).Checked = True
        sgnMode = 0.5
    Case 8  '25%
        mnuViewScaleMode(8).Checked = True
        sgnMode = 0.25
    End Select
    
    Call ShowSize
    Call ShowScroll
    Call ReFlashReportBySelFormat
    If Not scrHsc.Enabled Then DrawRuler picRulerH
    If Not scrVsc.Enabled Then DrawRuler picRulerV
    
    If GetSelNum = 1 Then ShowAttrib (intCurID)
End Sub

Private Sub msh_DblClick(Index As Integer)
    If Not (Left(msh(Index).Tag, 2) = "C_") Then
        If objReport.Items("_" & Index).系统 Then Exit Sub
    End If
    strMenu = "DO"
    msh_MouseDown Index, 2, 0, CDbl(lngPreX), CDbl(lngPreY)
    Select Case strMenu
        Case "mnuClass_Data"
            strMenu = ""
            mnuClass_Data_Click
        Case "mnuCustom_Col_Data"
            strMenu = ""
            mnuCustom_Col_Data_Click
        Case "mnuCustom_Head_Text"
            strMenu = ""
            mnuCustom_Head_Text_Click
    End Select
End Sub

Private Sub msh_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intIdx As Integer, intState As Integer
    Dim tmpID As RelatID, tmpItem As RPTItem
    
    lngPreX = X: lngPreY = Y
    Set mobjMove = Nothing: mlngX = 0: mlngY = 0
    blnHead = False
    
    '分栏控件Tag存放="C_主控件索引"
    If Left(msh(Index).Tag, 2) = "C_" Then
        Call SetGridLine(CInt(Mid(msh(Index).Tag, 3)))
        Call SetCopyGrid(CInt(Mid(msh(Index).Tag, 3)))
        intIdx = CInt(Mid(msh(Index).Tag, 3))
    Else
        Call SetGridLine(Index)
        Call SetCopyGrid(Index)
        intIdx = Index
        If msh(Index).MouseRow < msh(Index).FixedRows Then blnHead = True
    End If
    
    Call ReFlashWidth
    
    If Shift = 2 Then
        If Mid(msh(intIdx).Tag, 1, 2) = "" Then
            Call SelItem(intIdx, True) '加选
            If GetSelNum() = 1 Then
                Call ShowAttrib(intIdx) '只选中一个则显示属性
            Else
                Call ShowAttrib '多选时不显示属性
            End If
        Else
            Call SelItem(intIdx, False) '反选
            If GetSelNum() = 1 Then
                Call ShowAttrib(intCurID) '只选中一个则显示属性(选中的不一定是该控件)
            Else
                Call ShowAttrib '多选时不显示属性
            End If
        End If
    Else
        If Mid(msh(intIdx).Tag, 1, 2) = "" Then
            Call SelClear
            Call SelItem(intIdx, True)
        End If
    End If
    '表格编辑操作.只选中一个有效;分栏控件无效
    If Left(msh(Index).Tag, 2) <> "C_" And GetSelNum = 1 Then
        If objReport.Items("_" & intIdx).类型 = 4 Then '任意表格编辑操作
            If Button = 1 Then
                Call ResetColor(intIdx)
                selCell.Col1 = -1: selCell.Col2 = -1: selCell.Row1 = -1: selCell.Row2 = -1: selCell.Row = -1
                Call ShowPaperInfo
                drgCell.Col1 = -1: drgCell.Col2 = -1: drgCell.Row1 = -1: drgCell.Row2 = -1
                intCurCol = -1
            End If
            If msh(intIdx).MouseRow < msh(intIdx).FixedRows Then '表头范围
                If Button = 1 Then
                    selCell = GetCellRange(msh(intIdx), msh(intIdx).MouseRow, msh(intIdx).MouseCol)
                    selCell.Row = msh(intIdx).MouseRow
                    '拖动范初始值
                    drgCell = selCell
                    Call CustomCellColor(intIdx, selCell, True)
                ElseIf Button = 2 Then
                    Call ShowPaperInfo
                    If strMenu = "" Then
                        If Not objReport.Items("_" & Index).系统 Then PopupMenu mnuCustom_Head, 2
                    Else
                        strMenu = "mnuCustom_Head_Text"
                    End If
                End If
            Else '表列范围
                intCurCol = msh(intIdx).MouseCol
                If Button = 1 Then
                    Call CustomColColor(intIdx, intCurCol)
                Else
                    Call SetMenuDefault(intIdx, intCurCol)
                    If strMenu = "" Then
                        If Not objReport.Items("_" & Index).系统 Then PopupMenu mnuCustom_Col, 2
                    Else
                        strMenu = "mnuCustom_Col_Data"
                    End If
                End If
            End If
    
            '检查是否允许改变高度或宽度
            If blnAdjustRowHeight = False And msh(Index).MouseRow < msh(Index).FixedRows And Button = 1 And objReport.Items("_" & Index).系统 = False Then
                msh(Index).Row = msh(Index).MouseRow
                msh(Index).Col = msh(Index).MouseCol
                If msh(Index).Row < msh(Index).FixedRows Then
                    blnAdjustRowHeight = (Button = 1 And (Y > msh(Index).CellTop + msh(Index).CellHeight - 100 And Y < msh(Index).CellTop + msh(Index).CellHeight + 100))
                End If
            ElseIf blnAdjustColWidth = False And Button = 1 And msh(Index).MouseRow = msh(Index).FixedRows And msh(Index).Col = 0 And objReport.Items("_" & Index).系统 = False Then
                msh(Index).Row = msh(Index).MouseRow
                msh(Index).Col = msh(Index).MouseCol
                If msh(Index).Row = msh(Index).FixedRows And msh(Index).Col = 0 Then
                    blnAdjustColWidth = (Button = 1 And (X > msh(Index).CellLeft + msh(Index).CellWidth - 100 And X < msh(Index).CellLeft + msh(Index).CellWidth + 100))
                End If
                If blnAdjustColWidth Then
                    '显示分隔线,供调整列宽
                    msh(Index).MousePointer = 9
                    With PicSplit
                        .Left = X + msh(Index).Left
                        .Top = msh(Index).Top
                        .Height = msh(Index).CellTop + msh(Index).CellHeight
                        .ZOrder 0
                        .Visible = True
                    End With
                Else
                    PicSplit.Visible = False
                End If
            End If
        ElseIf objReport.Items("_" & intIdx).类型 = 5 Then '汇总表格编辑操作
            If Button = 1 Then
                Call ResetColor(intIdx)
                selCell.Col1 = -1: selCell.Row1 = -1
                Call ShowPaperInfo
            End If
            
            If msh(intIdx).MouseRow < msh(intIdx).FixedRows Then '表头范围
                If Button = 1 Then
                    selCell = GetCellRange(msh(intIdx), msh(intIdx).MouseRow, msh(intIdx).MouseCol)
                    selCell.Row = msh(intIdx).MouseRow
                    '拖动范初始值
                    drgCell = selCell
                    Call CustomCellColor(intIdx, selCell, True)
                End If
            End If
            
            '求统计项个数
            For Each tmpID In objReport.Items("_" & intIdx).SubIDs
                If objReport.Items("_" & tmpID.id).类型 = 9 Then intState = intState + 1
            Next
            If msh(intIdx).MouseCol <= msh(intIdx).FixedCols - 1 And msh(intIdx).MouseRow >= msh(intIdx).FixedRows - 1 Then
                 '纵向分类范围
                If Button = 1 Then
                    selCell.Row1 = msh(intIdx).MouseRow
                    selCell.Col1 = msh(intIdx).MouseCol
                    Call ClassColor(intIdx, selCell)
                ElseIf Button = 2 Then
                    If selCell.Col1 <> -1 And selCell.Row1 <> -1 Then
                        For Each tmpID In objReport.Items("_" & intIdx).SubIDs
                            Set tmpItem = objReport.Items("_" & tmpID.id)
                            If tmpItem.类型 = 7 And tmpItem.序号 = selCell.Col1 Then
                                Call SetDefaultState(tmpItem.汇总, tmpItem.对齐)
                            End If
                        Next
                    Else
                        Call SetDefaultState("", 0, True)
                    End If
                    mnuClass_Align.Visible = False '纵向分类,固定左对齐
                    mnuClass_State.Visible = True
                    If strMenu = "" And Not ReferObj(Index) Then
                        If Not objReport.Items("_" & Index).系统 Then PopupMenu mnuClass, 2
                    Else
                        strMenu = "mnuClass_Data"
                    End If
                End If
            ElseIf msh(intIdx).MouseRow <= msh(intIdx).FixedRows - 2 Then
                 '横向分类范围
                If Button = 1 Then
                    selCell.Row1 = msh(intIdx).MouseRow
                    selCell.Col1 = msh(intIdx).MouseCol
                    Call ClassColor(intIdx, selCell)
                ElseIf Button = 2 Then
                    If selCell.Col1 <> -1 And selCell.Row1 <> -1 Then
                        For Each tmpID In objReport.Items("_" & intIdx).SubIDs
                            Set tmpItem = objReport.Items("_" & tmpID.id)
                            If tmpItem.类型 = 8 And tmpItem.序号 = selCell.Row1 Then
                                Call SetDefaultState(tmpItem.汇总, tmpItem.对齐)
                            End If
                        Next
                    Else
                        Call SetDefaultState("", 0, True)
                    End If
                    mnuClass_Align.Visible = False '横向分类,固定中对齐
                    mnuClass_State.Visible = True
                    If strMenu = "" Then
                        If Not objReport.Items("_" & Index).系统 Then PopupMenu mnuClass, 2
                    Else
                        strMenu = "mnuClass_Data"
                    End If
                End If
            ElseIf msh(intIdx).MouseCol >= msh(intIdx).FixedCols And msh(intIdx).MouseRow >= msh(intIdx).FixedRows - 1 Then
                '统计项范围
                If Button = 1 Then
                    selCell.Row1 = msh(intIdx).MouseRow
                    If msh(intIdx).MouseCol <= msh(intIdx).FixedCols + intState - 1 Then
                        selCell.Col1 = msh(intIdx).MouseCol
                    Else
                        selCell.Col1 = msh(intIdx).FixedCols + (msh(intIdx).MouseCol - msh(intIdx).FixedCols) Mod intState
                    End If
                    Call ClassColor(intIdx, selCell, intState)
                ElseIf Button = 2 Then
                    If selCell.Col1 <> -1 And selCell.Row1 <> -1 Then
                        For Each tmpID In objReport.Items("_" & intIdx).SubIDs
                            Set tmpItem = objReport.Items("_" & tmpID.id)
                            If tmpItem.类型 = 9 And tmpItem.序号 = selCell.Col1 - msh(intIdx).FixedCols Then
                                Call SetDefaultState(tmpItem.汇总, tmpItem.对齐)
                            End If
                        Next
                    Else
                        Call SetDefaultState("", 0, True)
                    End If
                    mnuClass_Align.Visible = True
                    mnuClass_State.Visible = False '统计项,本身就是汇总,其它汇总形式依赖于横纵向汇总
                    If strMenu = "" Then
                        If Not objReport.Items("_" & Index).系统 Then PopupMenu mnuClass, 2
                    Else
                        strMenu = "mnuClass_Data"
                    End If
                End If
            End If
            
            '检查是否允许改变高度或宽度
            If blnAdjustRowHeight = False And msh(Index).MouseRow < msh(Index).FixedRows And Button = 1 And objReport.Items("_" & Index).系统 = False Then
                msh(Index).Row = msh(Index).MouseRow
                msh(Index).Col = msh(Index).MouseCol
                If msh(Index).Row < msh(Index).FixedRows Then
                    blnAdjustRowHeight = (Button = 1 And (Y > msh(Index).CellTop + msh(Index).CellHeight - 100 And Y < msh(Index).CellTop + msh(Index).CellHeight + 100))
                End If
            End If
        End If
    End If
    If GetSelNum = 1 Then Call ShowAttrib(intIdx) '只选中一个则显示属性
End Sub

Private Sub msh_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tmpCell As Cells, RowH As Single, lngRow As Long, sgnH As Single
    Dim LngLastRow As Long, LngLastCol As Long, arrHead, i As Integer
    Dim sgnAlig As Single, strCaption As String, tmpID As RelatID, tmpItem As RPTItem
    
    If Not (Left(msh(Index).Tag, 2) = "C_") Then
        If objReport.Items("_" & Index).系统 Then Exit Sub
    End If
    
    DrawXY X + msh(Index).Left, Y + msh(Index).Top
    If msh(Index).MouseRow >= 0 And msh(Index).MouseRow <= msh(Index).Rows - 1 And _
        msh(Index).MouseCol >= 0 And msh(Index).MouseCol <= msh(Index).Cols - 1 Then
        msh(Index).ToolTipText = msh(Index).TextMatrix(msh(Index).MouseRow, msh(Index).MouseCol)
    End If
    If Left(msh(Index).Tag, 2) = "C_" Then
        If Button = 1 And Mid(msh(CInt(Mid(msh(Index).Tag, 3))).Tag, 1, 2) <> "" Then
            If blnLock Then Exit Sub
            Call MoveSelect(X - lngPreX, Y - lngPreY)
            If GetSelNum() = 1 Then ShowAttrib CInt(Mid(msh(Index).Tag, 3))
        End If
    Else
        If objReport.Items("_" & Index).类型 = 5 Then
            If msh(Index).MouseRow <= msh(Index).FixedRows Then
                If blnAdjustRowHeight = False And blnAdjustColWidth = False And Button = 0 And objReport.Items("_" & Index).系统 = False Then
                    LngLastRow = msh(Index).Row: LngLastCol = msh(Index).Col
                    msh(Index).Row = msh(Index).MouseRow: msh(Index).Col = msh(Index).MouseCol
                    msh(Index).MousePointer = IIF(Y > msh(Index).CellTop + msh(Index).CellHeight - 100 And Y < msh(Index).CellTop + msh(Index).CellHeight + 100 And msh(Index).MouseRow <> msh(Index).FixedRows, 7, 99)
                    msh(Index).Row = LngLastRow: msh(Index).Col = LngLastCol
                Else
                    msh(Index).MousePointer = 99
                    If mobjPicMove Is Nothing Then Set mobjPicMove = LoadResPicture("MERGE", vbResCursor)
                    If Not msh(Index).MouseIcon = mobjPicMove Then
                        Set msh(Index).MouseIcon = LoadResPicture("MOVE", vbResCursor)
                        Set mobjPicMove = msh(Index).MouseIcon
                    End If
                End If
                
                If blnAdjustRowHeight Then
                    On Error Resume Next
                    If intCurID = 0 Then blnAdjustRowHeight = False: Exit Sub
                    If selCell.Row = -1 Then blnAdjustRowHeight = False: Exit Sub
                    msh(Index).MousePointer = 7
                    msh(Index).RowHeight(msh(Index).Row) = msh(Index).RowHeight(msh(Index).Row) + Y - lngPreY
                    lngPreY = Y

                    '至少显示两行数据行
                    LngLastCol = 0
                    For LngLastRow = 0 To msh(intCurID).FixedRows + 1
                        LngLastCol = LngLastCol + msh(intCurID).RowHeight(LngLastRow)
                    Next
                    If LngLastCol > msh(intCurID).Height - 100 * sgnMode Then
                        arrHead = Split(objReport.Items("_" & objReport.Items("_" & intCurID).SubIDs(1).Key).表头, "|")
                        msh(intCurID).RowHeight(selCell.Row) = Split(arrHead(selCell.Row), "^")(1) * sgnMode
                    End If
'
                    '保存行高
                    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                        Set tmpItem = objReport.Items("_" & tmpID.id)
                        arrHead = Split(tmpItem.表头, "|")
                        tmpItem.表头 = ""
                        For i = 0 To 2 'UBound(arrHead)
                            If i >= selCell.Row1 And i <= selCell.Row2 Then
                                msh(Index).Row = i: msh(Index).Col = tmpItem.序号
                                sgnH = msh(Index).RowHeight(i)
                                sgnAlig = msh(Index).CellAlignment
                                strCaption = msh(Index).TextMatrix(msh(Index).Row, msh(Index).Col)
                                If strCaption = "" Then strCaption = "#"
                                tmpItem.表头 = tmpItem.表头 & "|" & sgnAlig & "^" & sgnH / sgnMode & "^" & strCaption
                            Else
                                tmpItem.表头 = tmpItem.表头 & "|" & arrHead(i)
                            End If
                        Next
                        tmpItem.表头 = Mid(tmpItem.表头, 2)
                    Next
                    On Error GoTo 0
'                ElseIf msh(Index).MouseRow = -1 Then
'                    If Button = 1 And Mid(msh(Index).Tag, 1, 2) <> "" And blnAdjustColWidth = False Then
'                        If blnLock Then Exit Sub
'                        Call MoveSelect(X - lngPreX, Y - lngPreY)
'                        If GetSelNum() = 1 Then ShowAttrib Index
'                    End If
                End If
            ElseIf Button = 1 And Mid(msh(Index).Tag, 1, 2) <> "" Then
                If blnLock Then Exit Sub
                Call MoveSelect(X - lngPreX, Y - lngPreY)
                If GetSelNum() = 1 Then ShowAttrib Index
            Else
                msh(Index).MousePointer = 99
                If mobjPicMove Is Nothing Then Set mobjPicMove = LoadResPicture("MERGE", vbResCursor)
                If Not msh(Index).MouseIcon = mobjPicMove Then
                    Set msh(Index).MouseIcon = LoadResPicture("MOVE", vbResCursor)
                    Set mobjPicMove = msh(Index).MouseIcon
                End If
            End If
            
        ElseIf objReport.Items("_" & Index).类型 = 4 Then
            If msh(Index).MouseRow < msh(Index).FixedRows Then
                If blnAdjustRowHeight = False And blnAdjustColWidth = False And Button = 0 And objReport.Items("_" & Index).系统 = False Then
                    LngLastRow = msh(Index).Row: LngLastCol = msh(Index).Col
                    msh(Index).Row = msh(Index).MouseRow: msh(Index).Col = msh(Index).MouseCol
                    msh(Index).MousePointer = IIF(Y > msh(Index).CellTop + msh(Index).CellHeight - 100 And Y < msh(Index).CellTop + msh(Index).CellHeight + 100, 7, 99)
                    msh(Index).Row = LngLastRow: msh(Index).Col = LngLastCol
                Else
                    msh(Index).MousePointer = 99
                    If mobjPicMove Is Nothing Then Set mobjPicMove = LoadResPicture("MERGE", vbResCursor)
                    If Not msh(Index).MouseIcon = mobjPicMove Then
                        Set msh(Index).MouseIcon = LoadResPicture("MOVE", vbResCursor)
                        Set mobjPicMove = msh(Index).MouseIcon
                    End If
                End If
                If blnAdjustColWidth = False And msh(Index).MousePointer = 99 Then
                    If mobjPicMERGE Is Nothing Then Set mobjPicMERGE = LoadResPicture("MOVE", vbResCursor)
                    If Not msh(Index).MouseIcon = mobjPicMERGE Then
                        Set msh(Index).MouseIcon = LoadResPicture("MERGE", vbResCursor)
                        Set mobjPicMERGE = msh(Index).MouseIcon
                    End If
                End If
                If blnAdjustRowHeight Then
                    On Error Resume Next
                    If intCurID = 0 Then blnAdjustRowHeight = False: Exit Sub
                    If selCell.Row = -1 Then blnAdjustRowHeight = False: Exit Sub
                    
                    msh(Index).MousePointer = 7
                    PicFontTest.FontName = objReport.Items("_" & intCurID).字体
                    PicFontTest.FontSize = objReport.Items("_" & intCurID).字号
                    sgnH = PicFontTest.TextHeight("字") + 15
                    
                    msh(Index).RowHeight(msh(Index).Row) = msh(Index).RowHeight(msh(Index).Row) + Y - lngPreY
                    For lngRow = selCell.Row1 To selCell.Row2
                        If Abs(msh(Index).RowHeight(lngRow)) < sgnH Then msh(Index).RowHeight(lngRow) = sgnH * sgnMode
                    Next
                    lngPreY = Y
                    
                    '至少显示两行数据行
                    LngLastCol = 0
                    For LngLastRow = 0 To msh(intCurID).FixedRows + 1
                        LngLastCol = LngLastCol + msh(intCurID).RowHeight(LngLastRow)
                    Next
                    If LngLastCol > msh(intCurID).Height - 100 * sgnMode Then
                        arrHead = Split(objReport.Items("_" & objReport.Items("_" & intCurID).SubIDs(1).Key).表头, "|")
                        msh(intCurID).RowHeight(selCell.Row) = Split(arrHead(selCell.Row), "^")(1) * sgnMode
                    End If
                    
                    '保存行高
                    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                        Set tmpItem = objReport.Items("_" & tmpID.id)
                        arrHead = Split(tmpItem.表头, "|")
                        tmpItem.表头 = ""
                        For i = 0 To UBound(arrHead)
                            If i >= selCell.Row1 And i <= selCell.Row2 Then
                                msh(Index).Row = i: msh(Index).Col = tmpItem.序号
                                sgnH = msh(Index).RowHeight(i)
                                sgnAlig = msh(Index).CellAlignment
                                strCaption = msh(Index).TextMatrix(msh(Index).Row, msh(Index).Col)
                                If strCaption = "" Then strCaption = "#"
                                tmpItem.表头 = tmpItem.表头 & "|" & sgnAlig & "^" & sgnH / sgnMode & "^" & strCaption
                            Else
                                tmpItem.表头 = tmpItem.表头 & "|" & arrHead(i)
                            End If
                        Next
                        tmpItem.表头 = Mid(tmpItem.表头, 2)
                    Next
                ElseIf msh(Index).MouseRow = -1 Then
                    If Button = 1 And Mid(msh(Index).Tag, 1, 2) <> "" And blnAdjustColWidth = False Then
                        If blnLock Then Exit Sub
                        Call MoveSelect(X - lngPreX, Y - lngPreY)
                        If GetSelNum() = 1 Then ShowAttrib Index
                    End If
                End If
            ElseIf msh(Index).MouseRow = msh(Index).FixedRows Then
                '允许列拖动(同时改变所有列的宽度)
                If objReport.Items("_" & Index).系统 = False And blnAdjustColWidth = False And msh(Index).MouseCol = 0 And Button = 0 Then
                    LngLastRow = msh(Index).Row: LngLastCol = msh(Index).Col
                    msh(Index).Row = msh(Index).MouseRow: msh(Index).Col = msh(Index).MouseCol
                    If msh(Index).Row = msh(Index).FixedRows And msh(Index).Col = 0 Then
                        msh(Index).MousePointer = IIF(X > msh(Index).CellLeft + msh(Index).CellWidth - 100 And X < msh(Index).CellLeft + msh(Index).CellWidth + 100, 9, 99)
                    End If
                    msh(Index).Row = LngLastRow: msh(Index).Col = LngLastCol
                End If
                If blnAdjustColWidth Then
                    '显示分隔线,供调整列宽
                    msh(Index).MousePointer = 9
                    With PicSplit
                        .Left = X + msh(Index).Left
                        .Top = msh(Index).Top
                        .Height = msh(Index).CellTop + msh(Index).CellHeight
                        .ZOrder 0
                        .Visible = True
                    End With
                Else
                    PicSplit.Visible = False
                    If msh(Index).MouseCol <> 0 Then
                        msh(Index).MousePointer = 99
                        If mobjPicMove Is Nothing Then Set mobjPicMove = LoadResPicture("MERGE", vbResCursor)
                        If Not msh(Index).MouseIcon = mobjPicMove Then
                            Set msh(Index).MouseIcon = LoadResPicture("MOVE", vbResCursor)
                            Set mobjPicMove = msh(Index).MouseIcon
                        End If
                    End If
                End If
            Else
                msh(Index).MousePointer = 99
                If mobjPicMove Is Nothing Then Set mobjPicMove = LoadResPicture("MERGE", vbResCursor)
                If Not msh(Index).MouseIcon = mobjPicMove Then
                    Set msh(Index).MouseIcon = LoadResPicture("MOVE", vbResCursor)
                    Set mobjPicMove = msh(Index).MouseIcon
                End If
            End If
            If Not blnHead Then
                If Button = 1 And Mid(msh(Index).Tag, 1, 2) <> "" And blnAdjustColWidth = False Then
                    If blnLock Then Exit Sub
                    Call MoveSelect(X - lngPreX, Y - lngPreY)
                    If GetSelNum() = 1 Then ShowAttrib Index
                End If
            ElseIf Button = 1 And blnAdjustRowHeight = False Then '拖动选择单元格范围
                If msh(Index).MouseRow >= 0 And msh(Index).MouseRow <= msh(Index).FixedRows - 1 And _
                    msh(Index).MouseCol >= 0 And msh(Index).MouseCol <= msh(Index).Cols - 1 Then
                    drgCell.Row2 = msh(Index).MouseRow
                    drgCell.Col2 = msh(Index).MouseCol
                    
                    tmpCell = MergeCell(Index, selCell, drgCell)
                    If tmpCell.Row1 <> -1 And tmpCell.Col1 <> -1 And tmpCell.Row2 <> -1 And tmpCell.Col2 <> -1 Then
                        Call CustomCellColor(Index, tmpCell)
                    Else
                        Call CustomCellColor(Index, selCell)
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub msh_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tmpCell As Cells, i As Integer
    Dim tmpID As RelatID, tmpItem As RPTItem, arrHead, arrModify
    Dim sgnH As Long, sgnAlig As Long, strCaption As String
    Dim blnYes As Boolean, j As Long, k As Long
    Dim strSouse As String
    
    If blnAdjustColWidth Then
        PicSplit.Visible = False
        blnAdjustColWidth = False
        If PicSplit.Left < msh(Index).Left + 200 * sgnMode Or PicSplit.Left > msh(Index).Width - 200 * sgnMode Then Call VBA.Beep: Exit Sub
        sgnH = PicSplit.Left - msh(Index).Left
        For i = 0 To msh(Index).Cols - 1
            msh(Index).ColWidth(i) = sgnH
        Next
        
        '更改集合
        For Each tmpID In objReport.Items("_" & intCurID).SubIDs
            Set tmpItem = objReport.Items("_" & tmpID.id)
            tmpItem.W = sgnH / sgnMode
        Next
        If GetSelNum = 1 Then ShowAttrib (Index)
    End If
    If blnAdjustRowHeight Then
        '保存固定行的行高
        If selCell.Row <> -1 Then
            For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                Set tmpItem = objReport.Items("_" & tmpID.id)
                arrHead = Split(tmpItem.表头, "|")
                tmpItem.表头 = ""
                For i = 0 To UBound(arrHead)
                    If i >= selCell.Row1 And i <= selCell.Row2 Then
                        msh(Index).Row = i: msh(Index).Col = tmpItem.序号
                        sgnH = msh(Index).RowHeight(i)
                        sgnAlig = msh(Index).CellAlignment
                        strCaption = msh(Index).TextMatrix(msh(Index).Row, msh(Index).Col)
                        If strCaption = "" Then strCaption = "#"
                        tmpItem.表头 = tmpItem.表头 & "|" & sgnAlig & "^" & sgnH / sgnMode & "^" & strCaption
                    Else
                        tmpItem.表头 = tmpItem.表头 & "|" & arrHead(i)
                    End If
                Next
                tmpItem.表头 = Mid(tmpItem.表头, 2)
            Next
        End If
        blnAdjustRowHeight = False
        If GetSelNum = 1 Then ShowAttrib (Index)
    End If
    
    If Left(msh(Index).Tag, 2) <> "C_" Then
        If objReport.Items("_" & Index).类型 = 5 Then
        
        ElseIf objReport.Items("_" & Index).类型 = 4 Then
            If Button = 1 And blnHead Then
                tmpCell = MergeCell(Index, selCell, drgCell)
                tmpCell.Row = selCell.Row
                If tmpCell.Row1 <> -1 And tmpCell.Col1 <> -1 And tmpCell.Row2 <> -1 And tmpCell.Col2 <> -1 Then selCell = tmpCell
                selCell = AdjustCell(selCell)
                Call CustomCellColor(Index, selCell)
            End If
        End If
   
        If Not mobjMove Is Nothing And objReport.Items("_" & Index).类型 = 4 Then
            If Not mobjMove Is msh(Index).Container Or mobjMove Is picPaper Then
                If UCase(mobjMove.name) = "PIC" Then
                    If objReport.Items("_" & Index).分栏 > 1 Then
                        MsgBox "卡片中不允许放入分栏的表格。", vbInformation, App.Title
                        Set mobjMove = Nothing: mlngX = 0: mlngY = 0
                        Exit Sub
                    End If
                    '卡片内不允许附加表格
                    For Each tmpItem In objReport.Items
                        If tmpItem.格式号 = mbytCurrFmt Then
                            If tmpItem.类型 = 5 Or tmpItem.类型 = 4 Then
                                If tmpItem.参照 = objReport.Items("_" & Index).名称 Then
                                     MsgBox "本表存在附加表格，不允许放入卡片中！", vbInformation, App.Title
                                     Set mobjMove = Nothing: mlngX = 0: mlngY = 0
                                     Exit Sub
                                End If
                            End If
                        End If
                    Next
                    '如果卡片有数据源，则检查表格的数据源是否匹配
                    If objReport.Items("_" & mobjMove.Index).数据源 <> "" Then
                        For Each tmpID In objReport.Items("_" & Index).SubIDs
                            With objReport.Items("_" & tmpID.id)
                                i = InStr(1, .内容, "]")
                                j = InStr(1, .内容, ".")
                                k = InStr(1, .内容, "[")
                                If i > k And i > j And i <> 0 And k <> 0 And j <> 0 Then
                                    If Mid(.内容, k + 1, j - k - 1) <> objReport.Items("_" & mobjMove.Index).数据源 Then
                                        If blnYes = False Then
                                            If MsgBox("当前卡片绑定了数据源，而表格中的数据列和卡片数据源不相同，移入将清空不匹配的列，是否继续?", vbQuestion + vbYesNo + vbDefaultButton1, App.Title) = vbYes Then
                                                .内容 = ""
                                                msh(Index).TextMatrix(1, .序号) = ""
                                                blnYes = True
                                            Else
                                                Exit Sub
                                            End If
                                        Else
                                            .内容 = ""
                                            msh(Index).TextMatrix(1, .序号) = ""
                                        End If
                                    End If
                                End If
                            End With
                        Next
                    Else
                        '卡片没有数据源，则提示用户是否添加数据源
                        For Each tmpID In objReport.Items("_" & Index).SubIDs
                            With objReport.Items("_" & tmpID.id)
                                i = InStr(1, .内容, "]")
                                j = InStr(1, .内容, ".")
                                k = InStr(1, .内容, "[")
                                If i > k And i > j And i <> 0 And k <> 0 And j <> 0 Then
                                    If InStr(strSouse, Mid(.内容, k + 1, j - k - 1)) = 0 Then
                                        strSouse = strSouse & "," & Mid(.内容, k + 1, j - k - 1)
                                    End If
                                End If
                            End With
                        Next
                        strSouse = Mid(strSouse, 2)
                        '只有一个数据源时才提示
                        If InStr(strSouse, ",") = 0 And strSouse <> "" Then
                            If MsgBox("当前卡片未绑定数据源，绑定后将分组打印多张卡片，数据源中存在""分组标识""字段则""分组标识""相同的为一组,否则一行数据为一组；" & vbCrLf & _
                                 "不绑定则只打印一张卡片，是否绑定数据源""" & strSouse & """?", vbQuestion + vbYesNo + vbDefaultButton1, App.Title) = vbYes Then
                                objReport.Items("_" & mobjMove.Index).数据源 = strSouse
                            End If
                        End If
                    End If
                End If
                
                For Each tmpItem In objReport.Items
                    If tmpItem.格式号 = mbytCurrFmt Then
                        If tmpItem.类型 = 2 Then
                            If tmpItem.参照 = objReport.Items("_" & Index).名称 Then
                                If UCase(mobjMove.name) = "PIC" Then
                                    objReport.Items("_" & tmpItem.id).父ID = mobjMove.Index
                                    lbl(tmpItem.id).Left = lbl(tmpItem.id).Left + (mlngX - msh(Index).Left)
                                    lbl(tmpItem.id).Top = lbl(tmpItem.id).Top + (mlngY - msh(Index).Top)
                                    objReport.Items("_" & tmpItem.id).X = lbl(tmpItem.id).Left
                                    objReport.Items("_" & tmpItem.id).Y = lbl(tmpItem.id).Top
                                Else
                                    objReport.Items("_" & tmpItem.id).父ID = 0
                                    lbl(tmpItem.id).Left = lbl(tmpItem.id).Left + lbl(tmpItem.id).Container.Left
                                    lbl(tmpItem.id).Top = lbl(tmpItem.id).Top + lbl(tmpItem.id).Container.Top
                                    objReport.Items("_" & tmpItem.id).X = lbl(tmpItem.id).Left
                                    objReport.Items("_" & tmpItem.id).Y = lbl(tmpItem.id).Top
                                End If
                                Set lbl(tmpItem.id).Container = mobjMove
                            End If
                        End If
                    End If
                Next
                
                Set msh(Index).Container = mobjMove
                msh(Index).Top = mlngY: msh(Index).Left = mlngX
                If UCase(mobjMove.name) = "PIC" Then
                    objReport.Items("_" & Index).父ID = mobjMove.Index
                Else
                    objReport.Items("_" & Index).父ID = 0
                End If
                If objReport.Items("_" & Index).类型 = 4 Then
                    '处理子项
                    For Each tmpID In objReport.Items("_" & Index).SubIDs
                        objReport.Items("_" & tmpID.id).父ID = objReport.Items("_" & Index).父ID
                    Next
                End If
                objReport.Items("_" & Index).X = mlngX: objReport.Items("_" & Index).Y = mlngY
                
            End If
            Set mobjMove = Nothing: mlngX = 0: mlngY = 0
            Call ShowAttrib(Index)
        End If
    Else
        If Not mobjMove Is Nothing Then
            If UCase(mobjMove.name) = "PIC" Then
                MsgBox "卡片中不允许放入分栏的表格。", vbInformation, App.Title
                Set mobjMove = Nothing: mlngX = 0: mlngY = 0
            End If
        End If
    End If
End Sub

Private Function ItemIsGraph(ByVal intID As Integer) As Boolean
'功能：判断指定的报表元素(标签)是否图片字段
    Dim strNode As String, objNode As Node
    Dim i As Integer, j As Integer, k As Integer
    
    If objReport.Items("_" & intID).类型 = 2 Then
        i = InStr(objReport.Items("_" & intID).内容, "]")
        j = InStr(objReport.Items("_" & intID).内容, ".")
        k = InStr(objReport.Items("_" & intID).内容, "[")
        If i > k And i > j And i <> 0 And k <> 0 Then
            strNode = Mid(objReport.Items("_" & intID).内容, j + 1, i - j - 1)
            For Each objNode In tvwSQL.Nodes
                If mdlPublic.GetStdNodeText(objNode.Text) = strNode And IsType(Val(objNode.Tag), adLongVarBinary) Then
                    ItemIsGraph = True
                End If
            Next
        End If
    End If
End Function

Private Sub SetAttAutoSize(ByVal intIndex As Integer, ByVal blnFlag As Boolean)
'功能：用于设置标签自动调整大小
    Dim ObjSel As Object, intType As Integer, lngSize As Long, intID As Integer

    On Error Resume Next
    Set ObjSel = GetInxObj(intIndex)
    objReport.Items("_" & intIndex).自调 = blnFlag
            
    intType = objReport.Items("_" & intIndex).类型
    If intType = 13 Then 'QR二维条码
        If blnFlag Then
            Set ObjSel.Picture = DrawBarCode2D(ReplaceBracket(objReport.Items("_" & intIndex).内容), frmFlash.picTemp, lngSize)
            
            ObjSel.Height = Format(lngSize * sgnMode, "0.00")
            ObjSel.Width = Format(lngSize * sgnMode, "0.00")
            
            intID = intIndex
            Call SelItem(intID, False)
            Call SelItem(intID, True)
            
            objReport.Items("_" & intIndex).W = lngSize
            objReport.Items("_" & intIndex).H = lngSize
        End If
    Else
        '如果是标签，且字段为图型，则当作图型处理
        '因为字段图型这里无法自调,图型对象可以自调
        If ItemIsGraph(intIndex) Then intType = 11
        If intType = 11 Then '图片,及内容为图片的标签
            If blnFlag And Not objReport.Items("_" & intIndex).图片 Is Nothing Then
                
                Set ObjSel.Picture = objReport.Items("_" & intIndex).图片 '重新以原始图片为准
                
                ObjSel.Width = objReport.Items("_" & intIndex).图片.Width * (15 / 26.46) * sgnMode
                ObjSel.Height = objReport.Items("_" & intIndex).图片.Height * (15 / 26.46) * sgnMode
                
                intID = intIndex
                Call SelItem(intID, False)
                Call SelItem(intID, True)
                
                objReport.Items("_" & intIndex).X = ObjSel.Left / sgnMode
                objReport.Items("_" & intIndex).Y = ObjSel.Top / sgnMode
                objReport.Items("_" & intIndex).W = ObjSel.Width / sgnMode
                objReport.Items("_" & intIndex).H = ObjSel.Height / sgnMode
            End If
        ElseIf intType = 2 Then '标签
            ObjSel.AutoSize = blnFlag
            If ObjSel.AutoSize Then '自调后须调整LblSize控件的位置
                intID = intIndex
                Call SelItem(intID, False)
                Call SelItem(intID, True)
                Call ReferTo
                objReport.Items("_" & intIndex).X = ObjSel.Left / sgnMode
                objReport.Items("_" & intIndex).Y = ObjSel.Top / sgnMode
                objReport.Items("_" & intIndex).W = ObjSel.Width / sgnMode
                objReport.Items("_" & intIndex).H = ObjSel.Height / sgnMode
            End If
        End If
    End If
End Sub

Private Sub mshAtt_DblClick()
    Dim intType As Integer
    Dim blnFlag As Boolean, blnFlagOld As Boolean
    Dim ObjSel As Object, objSub As RelatID
    Dim objBarCode As StdPicture
    Dim strBarCode As String
    Dim tmpObj As PictureBox
    
    If blnLock Then Exit Sub
    
    If Not (intCurID = 0 And (mshAtt.TextMatrix(mshAtt.Row, 0) = "票据" _
        Or mshAtt.TextMatrix(mshAtt.Row, 0) = "空表打印" Or mshAtt.TextMatrix(mshAtt.Row, 0) = "动态纸张")) Then
        
        If GetSelNum = 0 Then Exit Sub
    
        '系统项目(标签)不允许编辑
        If InStr(1, "自动调整大小,边框", mshAtt.TextMatrix(mshAtt.Row, 0)) > 0 And objReport.Items("_" & intCurID).系统 Then Exit Sub
            
        Set ObjSel = GetInxObj(intCurID)
    End If
    
    blnFlagOld = mshAtt.TextMatrix(mshAtt.Row, 1) = "√"
    If mshAtt.TextMatrix(mshAtt.Row, 1) = "√" Then
        mshAtt.TextMatrix(mshAtt.Row, 1) = "×"
        blnFlag = False: BlnSave = False
    ElseIf mshAtt.TextMatrix(mshAtt.Row, 1) = "×" Then
        mshAtt.TextMatrix(mshAtt.Row, 1) = "√"
        blnFlag = True: BlnSave = False
    End If
    
    Select Case mshAtt.TextMatrix(mshAtt.Row, 0)
        Case "票据"
            objReport.票据 = blnFlag
        Case "空表打印"
            objReport.打印方式 = IIF(blnFlag, 0, 1)
        Case "动态纸张"
            objReport.Fmts("_" & mbytCurrFmt).动态纸张 = blnFlag
        Case "边线"
            objReport.Items("_" & intCurID).边框 = blnFlag
        Case "粗体"
            ObjSel.FontBold = blnFlag
            If objReport.Items("_" & intCurID).类型 = 4 Then Call SetCopyGrid(intCurID)
            objReport.Items("_" & intCurID).粗体 = blnFlag
        Case "斜体"
            ObjSel.FontItalic = blnFlag
            If objReport.Items("_" & intCurID).类型 = 4 Then Call SetCopyGrid(intCurID)
            objReport.Items("_" & intCurID).斜体 = blnFlag
        Case "下划线"
            ObjSel.FontUnderline = blnFlag
            If objReport.Items("_" & intCurID).类型 = 4 Then Call SetCopyGrid(intCurID)
            objReport.Items("_" & intCurID).下线 = blnFlag
        Case "边框"
            '如果被选中的所有元素都是相同的元素
            If lblSize.count > 9 Then
                For Each tmpObj In lblSize
                    If tmpObj.Index Mod 8 = 1 Then
                        Set ObjSel = GetInxObj(tmpObj.Tag)
                        ObjSel.BorderStyle = IIF(blnFlag, 1, 0)
                        objReport.Items("_" & tmpObj.Tag).边框 = blnFlag
                    End If
                Next
            Else
                ObjSel.BorderStyle = IIF(blnFlag, 1, 0)
                objReport.Items("_" & intCurID).边框 = blnFlag
            End If
        Case "换行"
            If blnFlag = False Then
                '检查表格对象列的“自适应行高”属性
                Set ObjSel = objReport.Items("_" & intCurID)
                If Not ObjSel Is Nothing Then
                    If ObjSel.类型 = Val("4-任意表") Then
                        For Each objSub In ObjSel.SubIDs
                            If objReport.Items("_" & objSub.id).自适应行高 Then
                                ObjSel.自调 = blnFlagOld
                                mshAtt.TextMatrix(mshAtt.Row, 1) = IIF(blnFlagOld, "√", "×")
                                MsgBox "请先将表格中所有列的“该列单元格的高度随内容自动调整”设置取消！"
                                Exit Sub
                            End If
                        Next
                    End If
                End If
            End If
            objReport.Items("_" & intCurID).自调 = blnFlag
        Case "对齐"
            If cboAtt.Visible Then
                cboAtt.ListIndex = (cboAtt.ListIndex + 1) Mod 3
                Call cboAtt_Click
            End If
        Case "形状"
            If cboAtt.Visible Then
                cboAtt.ListIndex = (cboAtt.ListIndex + 1) Mod 2
                Call cboAtt_Click
            End If
        Case "保持比例"
            objReport.Items("_" & intCurID).粗体 = blnFlag
            '保持比例
            If Not objReport.Items("_" & intCurID).图片 Is Nothing Then
                If objReport.Items("_" & intCurID).粗体 Then
                    Set ObjSel.Picture = ScalePicture(PicFontTest, objReport.Items("_" & intCurID).图片, ObjSel.Width, ObjSel.Height)
                Else
                    Set ObjSel.Picture = objReport.Items("_" & intCurID).图片
                End If
            End If
        Case "自动字体"
            '如果被选中的所有元素都是相同的元素
            If lblSize.count > 9 Then
                For Each tmpObj In lblSize
                    If tmpObj.Index Mod 8 = 1 Then
                        objReport.Items("_" & tmpObj.Tag).行高 = IIF(blnFlag, 1, 0)
                    End If
                Next
            Else
                objReport.Items("_" & intCurID).行高 = IIF(blnFlag, 1, 0)
            End If
        Case "自动调整大小"
            '如果被选中的所有元素都是相同的元素
            If lblSize.count > 9 Then
                On Error Resume Next
                For Each tmpObj In lblSize
                    If tmpObj.Index Mod 8 = 1 Then
                        Call SetAttAutoSize(tmpObj.Tag, blnFlag)
                    End If
                Next
                On Error GoTo 0
            Else
                Call SetAttAutoSize(intCurID, blnFlag)
            End If
        Case "水平反转"
            objReport.Items("_" & intCurID).水平反转 = blnFlag
        Case "加粗"
            objReport.Items("_" & intCurID).粗体 = blnFlag
            If objReport.Items("_" & intCurID).类型 = 10 Then
                ObjSel.BorderWidth = IIF(blnFlag, 2, 1)
            ElseIf objReport.Items("_" & intCurID).类型 = 1 Then
                ObjSel.Height = IIF(blnFlag, 30, 15)
            End If
        Case "表格线加粗"
            objReport.Items("_" & intCurID).表格线加粗 = blnFlag
            If objReport.Items("_" & intCurID).类型 = 4 Or objReport.Items("_" & intCurID).类型 = 5 Then
                ObjSel.GridLineWidth = IIF(blnFlag, 2, 1)
            End If
        Case "报告图像"
            objReport.Items("_" & intCurID).下线 = blnFlag
        Case "前景色", "背景色", "字体", "表体网格色", "设置", "关联报表", "表头网格色", "网格色"
            If objReport.Items("_" & intCurID).系统 And InStr(1, "4,5", objReport.Items("_" & intCurID).类型) <> 0 Then Exit Sub
            If cmdAtt.Visible Then cmdAtt_Click
        Case "显示数字" '用于条码
            objReport.Items("_" & intCurID).表头 = SetBit(objReport.Items("_" & intCurID).表头, 1, IIF(blnFlag, 1, 0))
            
            With objReport.Items("_" & intCurID)
                strBarCode = ReplaceBracket(.内容)
                If strBarCode = "" Then strBarCode = "1234567890"
                
                Unload frmFlash '强制初始Picture，不然切换绘制有问题
                If .序号 = 1 Then
                    Set objBarCode = DrawBarCode128(frmFlash.picTemp, 3, strBarCode, Mid(.表头, 1, 1) = "1")
                ElseIf .序号 = 2 Then
                    Set objBarCode = DrawBarCode39(frmFlash.picTemp, 3, strBarCode, Mid(.表头, 2, 1) = "1", Mid(.表头, 1, 1) = "1")
                ElseIf .序号 = 3 Then
                    Set objBarCode = DrawBarCode128Auto(frmFlash.picTemp, strBarCode, 0, .行高, Mid(.表头, 1, 1) = "1")
                End If
                If Val(Mid(.表头, 3, 1)) <> 0 Then
                    Set objBarCode = PictureSpin(objBarCode, Val(Mid(.表头, 3, 1)), frmFlash.picTemp)
                End If
                Set ObjSel.Picture = objBarCode
            End With
        Case "求校验和" '用于条码
            objReport.Items("_" & intCurID).表头 = SetBit(objReport.Items("_" & intCurID).表头, 2, IIF(blnFlag, 1, 0))
            With objReport.Items("_" & intCurID)
                If .序号 = 2 Then
                    Set objBarCode = DrawBarCode39(frmFlash.picTemp, 3, ReplaceBracket(.内容), Mid(.表头, 2, 1) = "1", Mid(.表头, 1, 1) = "1")
                End If
                If Val(Mid(.表头, 3, 1)) <> 0 Then
                    Set objBarCode = PictureSpin(objBarCode, Val(Mid(.表头, 3, 1)), frmFlash.picTemp)
                End If
                Set ObjSel.Picture = objBarCode
            End With
    End Select
End Sub

Private Sub mshAtt_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Call ClearXY
End Sub

Private Sub mshAtt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        KeyAscii = 0: mshAtt_DblClick
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0: mshAtt_AfterRowColChange 0, 0, mshAtt.Row, mshAtt.Col
    End If
End Sub

Private Sub mshAtt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mshAtt.MouseRow > 0 Then mshAtt_AfterRowColChange 0, 0, mshAtt.Row, mshAtt.Col
End Sub

Private Sub pic_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    If Button = 1 And bytCurTool <> 0 Then
        picPaper.Cls
        
        '确定选择区域
        selArea.Right = X
        selArea.Bottom = Y
        blnDown = False
        If bytCurTool = 1 Then
            '修正线条的作图区域(控件宽高自动至少为15)
            If Abs(selArea.Right - selArea.Left) >= Abs(selArea.Bottom - selArea.Top) Then
                selArea.Bottom = selArea.Top
            Else
                selArea.Right = selArea.Left
            End If
        End If
        
        TrueArea selArea '修正选择区域
        
        If bytCurTool = 5 Then
            '任意表格最小宽高(W=1000+15,H=255*3+15)
            If selArea.Right - selArea.Left < 1015 Then selArea.Right = selArea.Left + 1015
            If selArea.Bottom - selArea.Top < 780 Then selArea.Bottom = selArea.Top + 780
        ElseIf bytCurTool = 6 Then
            '图表最小尺寸
            If selArea.Right - selArea.Left < Chart(0).Width Then selArea.Right = selArea.Left + Chart(0).Width
            If selArea.Bottom - selArea.Top < Chart(0).Height Then selArea.Bottom = selArea.Top + Chart(0).Height
        End If
        
        If bytCurTool = 0 Then
            '选择区域元素
            Call SelAreaItem(selArea)
            i = GetSelNum
            If i = 1 Then
                Call ShowAttrib(intCurID)
            ElseIf i = 0 Then
                Call ShowAttrib(, True)
            Else
                Call ShowAttrib
            End If
        Else
            '增加元素
            If Not (Abs(selArea.Left - selArea.Right) = 0 And Abs(selArea.Top - selArea.Bottom) = 0) Then '不处理点击
                Call AddReportItem(, pic(Index))
                BlnSave = False
            End If
        End If
        blnDown = False
    End If
End Sub

Private Sub picBack_GotFocus()
    Call NoneEdit
End Sub

Private Sub picL_GotFocus()
    Call NoneEdit
End Sub

Private Sub picM_GotFocus()
    Call NoneEdit
End Sub

Private Sub picPaper_DragDrop(Source As Control, X As Single, Y As Single)
    If UCase(Source.name) = "TVWSQL" Then
        selArea.Left = X: selArea.Top = Y
        Call AddReportItem(True)
        BlnSave = False
    End If
End Sub

Private Sub picPaper_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    DrawXY CLng(X), CLng(Y)
    If UCase(Source.name) = "TVWSQL" Then
        If State = 1 Then
            Set tvwSQL.DragIcon = lvwPar.DragIcon
        ElseIf State = 0 Then
            If tvwSQL.SelectedItem.Children = 0 Then
                Set tvwSQL.DragIcon = scrHsc.DragIcon
            Else
                Set tvwSQL.DragIcon = scrVsc.DragIcon
            End If
        End If
    End If
End Sub

Private Sub picPaper_GotFocus()
    Oldwinproc = GetWindowLong(picPaper.hwnd, GWL_WNDPROC)
    SetWindowLong picPaper.hwnd, GWL_WNDPROC, AddressOf FlexScroll
    Call NoneEdit
End Sub

Private Sub picPaper_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown Then
        '下
        If scrVsc.Value + scrVsc.Max / 10 > scrVsc.Max Then
            scrVsc.Value = scrVsc.Max
        Else
            scrVsc.Value = scrVsc.Value + scrVsc.Max / 10
        End If
    ElseIf KeyCode = vbKeyPageUp Then
        '上
        If scrVsc.Value - scrVsc.Max / 10 < 0 Then
            scrVsc.Value = 0
        Else
            scrVsc.Value = scrVsc.Value - scrVsc.Max / 10
        End If
    End If
End Sub

Private Sub picPaper_LostFocus()
    SetWindowLong picPaper.hwnd, GWL_WNDPROC, Oldwinproc
End Sub

Private Sub picPaperSize_GotFocus(Index As Integer)
    Call NoneEdit
End Sub

Private Sub picPaperSize_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    '将外面的元素移到可见位置
    Dim objItem As RPTItem
    Dim ObjSel As Object
    
    For Each objItem In objReport.Items
        If objItem.格式号 = mbytCurrFmt Then
            If objItem.X + objItem.W > objReport.Fmts(mbytCurrFmt).W And objReport.Fmts(mbytCurrFmt).W - objItem.W >= 0 Then
                objItem.X = objReport.Fmts(mbytCurrFmt).W - objItem.W
                Set ObjSel = GetInxObj(objItem.id)
                ObjSel.Left = objItem.X
            End If
            If objItem.Y + objItem.H > objReport.Fmts(mbytCurrFmt).H And objReport.Fmts(mbytCurrFmt).H - objItem.H >= 0 Then
                objItem.Y = objReport.Fmts(mbytCurrFmt).H - objItem.H
                Set ObjSel = GetInxObj(objItem.id)
                ObjSel.Top = objItem.Y
            End If
        End If
    Next
    SelClear
End Sub

Private Sub picR_GotFocus()
    Call NoneEdit
End Sub

Private Sub picRulerH_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Call ClearXY
End Sub

Private Sub picRulerV_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Call ClearXY
End Sub

Private Sub scrHsc_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Call ClearXY
End Sub

Private Sub scrHsc_GotFocus()
    Call NoneEdit
End Sub

Private Sub scrVsc_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Call ClearXY
End Sub

Private Sub scrVsc_GotFocus()
    Call NoneEdit
End Sub

Private Sub sta_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Call ClearXY
End Sub

Private Sub sta_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ClearXY
End Sub

Private Sub tb2_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Key = "History" Then
        Call mnuEdit_History_Click
    End If
End Sub

Private Sub tbr1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Call NoneEdit
    Select Case ButtonMenu.Key
        Case "Line"
            Call mnuEdit_ItemAdd_Click(0)
        Case "Frame"
            Call mnuEdit_ItemAdd_Click(1)
        Case "Label"
            Call mnuEdit_ItemAdd_Click(3)
        Case "Picture"
            Call mnuEdit_ItemAdd_Click(4)
        Case "Table"
            Call mnuEdit_ItemAdd_Click(5)
        Case "Chart"
            Call mnuEdit_ItemAdd_Click(6)
        Case "BarCode"
            Call mnuEdit_ItemAdd_Click(7)
    End Select
End Sub

Private Sub tbr1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Call ClearXY
End Sub

Private Sub tbr2_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Call ClearXY
End Sub

Private Sub tbr2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ClearXY
End Sub

Private Sub tbr1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ClearXY
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ClearXY
End Sub

Private Sub lblAtt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ClearXY
End Sub

Private Sub lblNote_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ClearXY
End Sub

Private Sub lblSQL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ClearXY
End Sub

Private Sub lblTool_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ClearXY
End Sub

Private Sub lvwPar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ClearXY
End Sub

Private Sub tbrTool_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim i As Integer
    
    bytCurTool = Button.Index - 1
    If bytCurTool = 0 Then
        picPaper.ForeColor = &HFF0000: picPaper.MousePointer = 0
    Else
        picPaper.ForeColor = &HFF&: picPaper.MousePointer = 2
    End If
End Sub

Private Sub tbrTool_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Call ClearXY
End Sub

Private Sub tbrTool_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ClearXY
End Sub

Private Sub tvwSQL_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Call ClearXY
    If State = 1 Then
        Set tvwSQL.DragIcon = lvwPar.DragIcon
    ElseIf State = 0 Then
        If tvwSQL.SelectedItem.Children = 0 Then
            Set tvwSQL.DragIcon = scrHsc.DragIcon
        Else
            Set tvwSQL.DragIcon = scrVsc.DragIcon
        End If
    End If
End Sub

Private Sub tvwSQL_GotFocus()
    Call NoneEdit
End Sub

Private Sub tvwSQL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objNode As Object, tmpNode As Object
    Dim i As Integer, j As Integer
    
    lngPreY = Y: lngPreX = X
    
    Set objNode = tvwSQL.HitTest(X, Y)
    
    '决定是否可拖动
    blnDrop = True
    blnSum = True
    If objNode Is Nothing Then
        blnDrop = False
    ElseIf objNode.Key = "Root" Then
        blnDrop = False
    ElseIf objNode.Children <> 0 Then
        If objReport.Datas(objNode.Key).类型 = 0 Then
            '任意表格
            If Not objNode.Checked Then blnDrop = False
        Else
            '汇总表格
            Set tmpNode = objNode.Child
            Do While Not tmpNode Is Nothing
                If tmpNode.Checked Then
                    If IsType(Val(tmpNode.Tag), adLongVarBinary) Then
                        blnDrop = False: Exit Do '有图片字段不允许作汇总表格
                    ElseIf IsType(Val(tmpNode.Tag), adNumeric) Then
                        i = i + 1 'i表示数字型字段个数
                    Else
                        j = j + 1 'i表示其它字段个数
                    End If
                End If
                Set tmpNode = tmpNode.Next
            Loop
            If i < 1 Then blnSum = False
            If i < 1 Or j < 1 Or i + j < 2 Then blnDrop = False
        End If
    End If
    If blnDrop Then
        Set tvwSQL.SelectedItem = objNode
        tvwSQL_NodeClick objNode
    End If
End Sub

Private Sub tvwsql_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    Dim objNode As Node
    
    
    Call ClearXY
    ClipCursor ByVal vbNullString
    If blnDrop And Button = 1 And (Abs(lngPreY - Y) > 300 Or Abs(lngPreX - X) > 300) Then
        If tvwSQL.SelectedItem.Children = 0 Then
            Set tvwSQL.DragIcon = scrHsc.DragIcon
        Else
            Set objNode = tvwSQL.SelectedItem.Child
            Do While Not objNode Is Nothing
                If objNode.Checked Then i = i + 1
                Set objNode = objNode.Next
            Loop
            If i = 0 Then
                ClipCursor GetObjRECT(tvwSQL.hwnd)
                Exit Sub
            End If
            Set tvwSQL.DragIcon = scrVsc.DragIcon
        End If
        tvwSQL.Drag 1
    ElseIf Button = 1 Then
        If blnSum = False And (X > (tvwSQL.Width - 200) Or Y > tvwSQL.Height - 200) Then
            MsgBox "数据源类型为分类汇总表，缺少汇总字段！", vbInformation + vbOKOnly, Me.Caption
        End If
        ClipCursor GetObjRECT(tvwSQL.hwnd)
    End If
End Sub

Private Sub mshAtt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ClearXY
End Sub

Private Sub picback_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ClearXY
End Sub

Private Sub picRulerH_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ClearXY
End Sub

Private Sub picRulerV_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ClearXY
End Sub

Private Sub mnuViewToolAttrib_Click()
    mnuViewToolAttrib.Checked = Not mnuViewToolAttrib.Checked
    picR.Visible = Not picR.Visible
    picAtt.Visible = Not picAtt.Visible
    Call Form_Resize
End Sub

Private Sub mnuViewToolRuler_Click()
    mnuViewToolRuler.Checked = Not mnuViewToolRuler.Checked
    picRulerH.Visible = Not picRulerH.Visible
    picRulerV.Visible = Not picRulerV.Visible
    Form_Resize
End Sub

Private Sub mnuViewToolSQL_Click()
    mnuViewToolSQL.Checked = Not mnuViewToolSQL.Checked
    picL.Visible = Not picL.Visible
    picSQL.Visible = Not picSQL.Visible
    Call Form_Resize
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    sta.Visible = Not sta.Visible
    Call Form_Resize
End Sub

Private Sub picL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ClearXY
    If Button = 1 Then
        If picSQL.Width + X < 1000 Or picBack.Width - X < 2000 Then Exit Sub
        If cboFormat.Width - X < 3000 Then Exit Sub
        picSQL.Width = picSQL.Width + X
        lblSQL.Width = lblSQL.Width + X
        tvwSQL.Width = tvwSQL.Width + X
        lblPar.Width = lblPar.Width + X
        lvwPar.Width = lvwPar.Width + X
        picRulerV.Left = picRulerV.Left + X
        picRulerH.Left = picRulerH.Left + X
        picRulerH.Width = picRulerH.Width - X
        scrHsc.Left = scrHsc.Left + X
        scrHsc.Width = scrHsc.Width - X
        picBack.Left = picBack.Left + X
        picBack.Width = picBack.Width - X
    
        'zyb#Add
        picFormat.Left = picFormat.Left + X
        picFormat.Width = picFormat.Width - X
        cmdDel.Left = picFormat.Width - cmdDel.Width
        cmdAdd.Left = cmdDel.Left - cmdAdd.Width
        cboFormat.Width = cmdAdd.Left - cboFormat.Left - 50
        
        Call ShowSize
        Call ShowScroll
        If Not scrHsc.Enabled Then DrawRuler picRulerH
        
        Refresh
    End If
End Sub

Private Sub picPaper_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lngPreX = X: lngPreY = Y
    blnDown = True

    If Button = 1 Then
        selArea.Left = X
        selArea.Top = Y
        bytLine = 0
        If Shift <> 2 Then Call SelClear
        Call ShowAttrib(, True)
    End If
    If Button = 2 Then PopupMenu mnuFormat, 2
End Sub

Private Sub picPaper_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static PreX As Long, PreY As Long
    
    If Shift = 0 Then picPaper.MousePointer = 0
    Call DrawXY(CLng(X), CLng(Y))
    
    '画选择虚框
    If Button = 1 And blnDown And Shift <> 4 Then
        If PreX = Empty And PreY = Empty Then
            PreX = selArea.Left
            PreY = selArea.Top
        End If
        If bytCurTool <> 1 Then
            picPaper.Line (selArea.Left, selArea.Top)-(PreX, PreY), picPaper.BackColor, B
            picPaper.Line (selArea.Left, selArea.Top)-(X, Y), , B
        Else
            If Abs(X - selArea.Left) >= Abs(Y - selArea.Top) Then
                '画横线
                If bytLine = 2 Then picPaper.Cls
                picPaper.Line (selArea.Left, selArea.Top)-(PreX, selArea.Top), picPaper.BackColor
                picPaper.Line (selArea.Left, selArea.Top)-(X, selArea.Top)
                bytLine = 1
            Else
                '画竖线
                If bytLine = 1 Then picPaper.Cls
                picPaper.Line (selArea.Left, selArea.Top)-(selArea.Left, PreY), picPaper.BackColor
                picPaper.Line (selArea.Left, selArea.Top)-(selArea.Left, Y)
                bytLine = 2
            End If
        End If
        PreX = X: PreY = Y
    End If
    
    '移动纸张
    If Button = 1 And Shift = 4 And blnDown Then
        If scrVsc.Enabled Then
            If (Y - lngPreY) / Screen.TwipsPerPixelX > 0 Then
                scrVsc.Value = IIF(scrVsc.Value - (Y - lngPreY) / Screen.TwipsPerPixelX < scrVsc.Min, scrVsc.Min, scrVsc.Value - (Y - lngPreY) / Screen.TwipsPerPixelX)
            Else
                scrVsc.Value = IIF(scrVsc.Value - (Y - lngPreY) / Screen.TwipsPerPixelX > scrVsc.Max, scrVsc.Max, scrVsc.Value - (Y - lngPreY) / Screen.TwipsPerPixelX)
            End If
        End If
        If scrHsc.Enabled Then
            If (X - lngPreX) / Screen.TwipsPerPixelX > 0 Then
                scrHsc.Value = IIF(scrHsc.Value - (X - lngPreX) / Screen.TwipsPerPixelX < scrHsc.Min, scrHsc.Min, scrHsc.Value - (X - lngPreX) / Screen.TwipsPerPixelX)
            Else
                scrHsc.Value = IIF(scrHsc.Value - (X - lngPreX) / Screen.TwipsPerPixelX > scrHsc.Max, scrHsc.Max, scrHsc.Value - (X - lngPreX) / Screen.TwipsPerPixelX)
            End If
        End If
    End If
End Sub

Private Sub picPaper_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    If Button = 1 And blnDown Then
        picPaper.Cls
        
        '确定选择区域
        selArea.Right = X
        selArea.Bottom = Y
        
        If bytCurTool = 1 Then
            '修正线条的作图区域(控件宽高自动至少为15)
            If Abs(selArea.Right - selArea.Left) >= Abs(selArea.Bottom - selArea.Top) Then
                selArea.Bottom = selArea.Top
            Else
                selArea.Right = selArea.Left
            End If
        End If
        
        TrueArea selArea '修正选择区域
        
        If bytCurTool = 5 Then
            '任意表格最小宽高(W=1000+15,H=255*3+15)
            If selArea.Right - selArea.Left < 1015 Then selArea.Right = selArea.Left + 1015
            If selArea.Bottom - selArea.Top < 780 Then selArea.Bottom = selArea.Top + 780
        ElseIf bytCurTool = 6 Then
            '图表最小尺寸
            If selArea.Right - selArea.Left < Chart(0).Width Then selArea.Right = selArea.Left + Chart(0).Width
            If selArea.Bottom - selArea.Top < Chart(0).Height Then selArea.Bottom = selArea.Top + Chart(0).Height
        End If
        
        If bytCurTool = 0 Then
            '选择区域元素
            Call SelAreaItem(selArea)
            i = GetSelNum
            If i = 1 Then
                Call ShowAttrib(intCurID)
            ElseIf i = 0 Then
                Call ShowAttrib(, True)
            Else
                Call ShowAttrib
            End If
        Else
            '增加元素
            If Not (Abs(selArea.Left - selArea.Right) = 0 And Abs(selArea.Top - selArea.Bottom) = 0) Then '不处理点击
                Call AddReportItem
                BlnSave = False
            End If
        End If
        blnDown = False
    End If
End Sub

Private Sub picR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call ClearXY
    If Button = 1 Then
        If picAtt.Width - X < 2000 Or picBack.Width + X < 2000 Then Exit Sub
        If cboFormat.Width + X < 3000 Then Exit Sub
        lblTool.Width = lblTool.Width - X
        tbrTool.Width = lblTool.Width
        lblAtt.Top = tbrTool.Top + tbrTool.Height
        lblAtt.Width = lblAtt.Width - X
        mshAtt.Top = lblAtt.Top + lblAtt.Height
        mshAtt.Width = mshAtt.Width - X
        mshAtt.Height = picM.Top - mshAtt.Top
        lblNote.Width = lblNote.Width - X
        
        picAtt.Width = picAtt.Width - X
        picRulerH.Width = picRulerH.Width + X
        scrVsc.Left = scrVsc.Left + X
        scrHsc.Width = scrHsc.Width + X
        picBack.Width = picBack.Width + X
        picM.Width = picM.Width - X
        
        'zyb#Add
        picFormat.Width = picFormat.Width + X
        cmdDel.Left = picFormat.Width - cmdDel.Width
        cmdAdd.Left = cmdDel.Left - cmdAdd.Width
        cboFormat.Width = cmdAdd.Left - cboFormat.Left - 50
        
        Call ShowSize
        Call ShowScroll
        If Not scrHsc.Enabled Then DrawRuler picRulerH
        
        Refresh
    End If
End Sub

Private Sub picM_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If lblNote.Height - Y < 300 Or mshAtt.Height + Y < 2000 Then Exit Sub
        picM.Top = picM.Top + Y
        mshAtt.Height = mshAtt.Height + Y
        lblNote.Top = lblNote.Top + Y
        lblNote.Height = lblNote.Height - Y
        Refresh
    End If
End Sub

Private Sub tbr1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Call NoneEdit
    Select Case Button.Key
        Case "Quit"
            mnuFile_Quit_Click
        Case "New"
            mnuEdit_New_Click
        Case "Modi"
            mnuEdit_Modi_Click
        Case "Del"
            mnuEdit_Del_Click
        Case "Page"
            mnuFile_Page_Click
        Case "Save"
            mnuFile_Save_Click
        Case "Remove"
            mnuEdit_Remove_Click
        Case "Report"
            mnuFile_Report_Click
        Case "Guide"
            mnuFile_Guide_Click
        Case "Help"
            mnuHelpTitle_Click
        Case "AddFormat"
            mnuEdit_AddFormat_Click
        Case "DelFormat"
            mnuEdit_DelFormat_Click
    End Select
End Sub

Private Sub mnuViewToolText_Click()
    Dim But As Button
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    If mnuViewToolText.Checked Then
        For Each But In tbr1.Buttons
            But.Caption = But.Tag
        Next
        For Each But In tbr2.Buttons
            But.Caption = But.Tag
        Next
        For Each But In tbrTool.Buttons
            But.Caption = But.Tag
        Next
    Else
        For Each But In tbr1.Buttons
            But.Caption = ""
        Next
        For Each But In tbr2.Buttons
            But.Caption = ""
        Next
        For Each But In tbrTool.Buttons
            But.Caption = ""
        Next
    End If
    cbr.Bands("System").MinHeight = tbr1.ButtonHeight
    cbr.Bands("Format").MinHeight = tbr2.ButtonHeight
    tbrTool.Height = tbrTool.ButtonHeight
    Form_Resize
End Sub

Private Sub mnuViewToolFormat_Click()
    mnuViewToolFormat.Checked = Not mnuViewToolFormat.Checked
    cbr.Bands("Format").Visible = Not cbr.Bands("Format").Visible
    If cbr.Bands("Format").Visible And cbr.Bands("System").Visible Then
        If cbr.Bands("System").Position = 2 Then cbr.Bands("System").NewRow = True
        If cbr.Bands("Format").Position = 2 Then cbr.Bands("Format").NewRow = True
    End If
    If Not mnuViewToolFormat.Checked And Not mnuViewToolSystem.Checked Then
        cbr.Visible = False
        mnuViewToolText.Enabled = False
    Else
        cbr.Visible = True
        mnuViewToolText.Enabled = True
    End If
    Form_Resize
End Sub

Private Sub mnuViewToolSystem_Click()
    mnuViewToolSystem.Checked = Not mnuViewToolSystem.Checked
    cbr.Bands("System").Visible = Not cbr.Bands("System").Visible
    If cbr.Bands("Format").Visible And cbr.Bands("System").Visible Then
        If cbr.Bands("System").Position = 2 Then cbr.Bands("System").NewRow = True
        If cbr.Bands("Format").Position = 2 Then cbr.Bands("Format").NewRow = True
    End If
    If Not mnuViewToolFormat.Checked And Not mnuViewToolSystem.Checked Then
        cbr.Visible = False
        mnuViewToolText.Enabled = False
    Else
        cbr.Visible = True
        mnuViewToolText.Enabled = True
    End If
    Form_Resize
End Sub

Private Sub DrawRuler(picRuler As PictureBox, Optional lngBegin As Long = 0)
'功能:显示标尺内容
'参数:picRuler=标尺控件;lngBegin=起始坐标值(单位:Twip,X或Y),应该为负数(<=0)
    Dim X As Long, Y As Long, IntStep As Integer
    Const FaceColor = &HA8A8A8
    
    IntStep = 283 * sgnMode
    With picRuler
        .Cls
        .DrawMode = vbCopyPen
        .FontName = "Times New Roman"
        .FontSize = 7.5
        .ForeColor = &H800000
        If .Width > .Height Then
            '横向
            '底纹
            picRuler.Line (0, 0)-(Screen.Width, .ScaleHeight / 4), FaceColor, BF
            picRuler.Line (.ScaleHeight - .ScaleHeight / 4, .ScaleHeight - .ScaleHeight / 4)-(Screen.Width, .ScaleHeight), FaceColor, BF
            picRuler.Line (0, 0)-(.ScaleHeight / 4, .ScaleHeight), FaceColor, BF
            picRuler.Line (.ScaleWidth - .ScaleHeight / 4, 0)-(.ScaleWidth, .ScaleHeight), FaceColor, BF
            '标注
            For X = .ScaleHeight + lngBegin To .ScaleWidth Step IntStep  '0.5cm
                If ((X - .ScaleHeight - lngBegin) / IntStep) Mod 2 = 0 Then
                    '文字
                    .CurrentY = .ScaleHeight / 2 - .TextHeight("0") / 2
                    .CurrentX = X - .TextWidth(CStr(((X - .ScaleHeight - lngBegin) / IntStep) / 2) & "0") / 2
                    picRuler.Print ((X - .ScaleHeight - lngBegin) / IntStep) / 2
                    '刻度
                    picRuler.Line (X, .ScaleHeight - .ScaleHeight / 4)-(X, .ScaleHeight), &HFFFFFF
                    picRuler.Line (X, 0)-(X, .ScaleHeight / 4), &HFFFFFF
                ElseIf ((X - .ScaleHeight - lngBegin) / IntStep) Mod 2 = 1 Then
                    picRuler.Line (X, .ScaleHeight - .ScaleHeight / 8 - 15)-(X, .ScaleHeight - .ScaleHeight / 8 + 15), &HFFFFFF
                    picRuler.Line (X, .ScaleHeight / 8 - 15)-(X, .ScaleHeight / 8 + 15), &HFFFFFF
                End If
            Next
        Else
            '纵向
            '底纹
            picRuler.Line (0, 0)-(.ScaleWidth / 4, Screen.Height), FaceColor, BF
            picRuler.Line (.ScaleWidth - .ScaleWidth / 4, 0)-(.ScaleWidth, Screen.Height), FaceColor, BF
            picRuler.Line (0, .ScaleHeight - .ScaleWidth / 4)-(.ScaleWidth, .ScaleHeight), FaceColor, BF
            '标注
            For Y = lngBegin To .ScaleHeight Step IntStep  '0.5cm
                If ((Y - lngBegin) / IntStep) Mod 2 = 0 Then
                    '文字
                    .CurrentX = .ScaleWidth / 4
                    .CurrentY = Y + .TextWidth(CStr(((Y - lngBegin) / IntStep) / 2)) / 2
                    objFont.OutPut picRuler, .CurrentX, .CurrentY, ((Y - lngBegin) / IntStep) / 2
                    '刻度
                    picRuler.Line (.ScaleWidth - .ScaleWidth / 4, Y)-(.ScaleWidth, Y), &HFFFFFF
                    picRuler.Line (0, Y)-(.ScaleWidth / 4, Y), &HFFFFFF
                ElseIf ((Y - lngBegin) / IntStep) Mod 2 = 1 Then
                    picRuler.Line (.ScaleWidth - .ScaleWidth / 8 - 15, Y)-(.ScaleWidth - .ScaleWidth / 8 + 15, Y), &HFFFFFF
                    picRuler.Line (.ScaleWidth / 8 - 15, Y)-(.ScaleWidth / 8 + 15, Y), &HFFFFFF
                End If
            Next
        End If
        .ForeColor = &HFFFF00
        .DrawMode = vbXorPen
    End With
End Sub

Private Sub picPaperSize_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lngPreX = X: lngPreY = Y
End Sub

Private Sub picPaperSize_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'功能:改变报表纸张大小
    Dim lngX As Long, lngY As Long
    Dim lngW As Long, lngH As Long
    
    Call ClearXY
    
    If Button = 1 Then
        If blnLock Then Exit Sub
        lngX = X - lngPreX
        lngY = Y - lngPreY
        
        '一旦改变,就设为自定义纸张
        With objReport.Fmts("_" & mbytCurrFmt)
            .纸张 = 256
            If .纸向 = 2 Then '交换宽高
                .W = .W + .H
                .H = .W - .H
                .W = .W - .H
                .纸向 = 1
            End If
        End With
        
        Select Case Index
            Case 0 '改变宽度
                If picPaper.Width + lngX < 283 * sgnMode Then Exit Sub
                
                picPaperSize(0).Left = picPaperSize(0).Left + lngX
                picPaper.Width = picPaper.Width + lngX
                objReport.Fmts("_" & mbytCurrFmt).W = picPaper.Width / sgnMode
                picPaperSize(2).Left = picPaperSize(2).Left + lngX
                picPaperSize(1).Width = picPaperSize(1).Width + lngX
                
                lngW = picBack.ScaleWidth - (picPaper.Width + picPaperSize(2).Width * 2 - Abs(picPaper.Left))
                If lngW < 0 Then lngW = 0
                
                If picBack.ScaleWidth >= picPaper.Width + picPaperSize(2).Width * 2 + lngW Then
                    scrHsc.Enabled = False
                Else
                    scrHsc.Enabled = True
                    scrHsc.Max = (picPaper.Width + picPaperSize(2).Width * 2 + lngW - picBack.ScaleWidth) / Screen.TwipsPerPixelX
                End If
            Case 1 '改变高度
                If picPaper.Height + lngY < 283 * sgnMode Then Exit Sub
                
                picPaperSize(1).Top = picPaperSize(1).Top + lngY
                picPaper.Height = picPaper.Height + lngY
                objReport.Fmts("_" & mbytCurrFmt).H = picPaper.Height / sgnMode
                picPaperSize(0).Height = picPaperSize(0).Height + lngY
                picPaperSize(2).Top = picPaperSize(2).Top + lngY
                
                lngH = picBack.ScaleHeight - (picPaper.Height + picPaperSize(2).Width * 2 - Abs(picPaper.Top))
                If lngH < 0 Then lngH = 0
                
                If picBack.ScaleHeight >= picPaper.Height + picPaperSize(2).Width * 2 + lngH Then
                    scrVsc.Enabled = False
                Else
                    scrVsc.Enabled = True
                    scrVsc.Max = (picPaper.Height + picPaperSize(2).Width * 2 + lngH - picBack.ScaleHeight) / Screen.TwipsPerPixelX
                End If
            Case 2 '改变宽高
                If picPaper.Height + lngY >= 283 * sgnMode Then
                    picPaper.Height = picPaper.Height + lngY
                    objReport.Fmts("_" & mbytCurrFmt).H = picPaper.Height / sgnMode
                    
                    picPaperSize(2).Top = picPaperSize(2).Top + lngY
                    picPaperSize(0).Height = picPaperSize(0).Height + lngY
                    picPaperSize(1).Top = picPaperSize(1).Top + lngY
                    
                    lngH = picBack.ScaleHeight - (picPaper.Height + picPaperSize(2).Width * 2 - Abs(picPaper.Top))
                    If lngH < 0 Then lngH = 0
                    
                    If picBack.ScaleHeight >= picPaper.Height + picPaperSize(2).Width * 2 + lngH Then
                        scrVsc.Enabled = False
                    Else
                        scrVsc.Enabled = True
                        scrVsc.Max = (picPaper.Height + picPaperSize(2).Width * 2 + lngH - picBack.ScaleHeight) / Screen.TwipsPerPixelX
                    End If
                End If
                If picPaper.Width + lngX >= 283 * sgnMode Then
                    picPaper.Width = picPaper.Width + lngX
                    objReport.Fmts("_" & mbytCurrFmt).W = picPaper.Width / sgnMode
                    
                    picPaperSize(2).Left = picPaperSize(2).Left + lngX
                    picPaperSize(0).Left = picPaperSize(0).Left + lngX
                    picPaperSize(1).Width = picPaperSize(1).Width + lngX
                    
                    lngW = picBack.ScaleWidth - (picPaper.Width + picPaperSize(2).Width * 2 - Abs(picPaper.Left))
                    If lngW < 0 Then lngW = 0
                    
                    If picBack.ScaleWidth >= picPaper.Width + picPaperSize(2).Width * 2 + lngW Then
                        scrHsc.Enabled = False
                    Else
                        scrHsc.Enabled = True
                        scrHsc.Max = (picPaper.Width + picPaperSize(2).Width * 2 + lngW - picBack.ScaleWidth) / Screen.TwipsPerPixelX
                    End If
                End If
        End Select
        BlnSave = False
    End If
End Sub

Private Sub scrhsc_Change()
    Call NoneEdit
    Call ClearXY
    Call ShowSize(-scrVsc.Value * 15#, -scrHsc.Value * 15#)
    If scrHsc.Value = 0 Then Call ShowScroll(1)
End Sub

Private Sub scrhsc_Scroll()
    Call NoneEdit
    Call ClearXY
    Call ShowSize(-scrVsc.Value * 15#, -scrHsc.Value * 15#)
    If scrHsc.Value = 0 Then Call ShowScroll(1)
End Sub

Private Sub scrVsc_Change()
    Call NoneEdit
    Call ClearXY
    Call ShowSize(-scrVsc.Value * 15#, -scrHsc.Value * 15#)
    If scrVsc.Value = 0 Then Call ShowScroll(2)
End Sub

Private Sub scrVsc_Scroll()
    Call NoneEdit
    Call ClearXY
    Call ShowSize(-scrVsc.Value * 15#, -scrHsc.Value * 15#)
    If scrVsc.Value = 0 Then Call ShowScroll(2)
End Sub

Private Sub ShowReportDetail()
'功能：根据objReport对象显示报表内容
    Call ShowSQLs
    Call ShowSize
    Call ShowScroll
    Call ShowItems
End Sub

Private Sub ShowSize(Optional lngTop As Single = 0, Optional lngLeft As Single = 0)
'功能:显示报表纸张大小
    Dim lngW As Long, lngH As Long
    
    picPaper.Left = lngLeft
    picPaper.Top = lngTop
    
    '打印的纸向只是简单地将纸张宽度和高度对调
    With objReport.Fmts("_" & mbytCurrFmt)
        If .纸向 = 1 Then
            lngW = .W: lngH = .H
        Else
            lngH = .W: lngW = .H
        End If
    End With
    
    picPaper.Width = Format(lngW * sgnMode, "0.00")
    picPaper.Height = Format(lngH * sgnMode, "0.00")
    
    '阴影及调整线位置
    picPaperSize(0).Top = picPaper.Top + picPaperSize(0).Width
    picPaperSize(0).Left = picPaper.Left + picPaper.Width
    picPaperSize(0).Height = picPaper.Height - picPaperSize(0).Width
    
    picPaperSize(1).Top = picPaper.Top + picPaper.Height
    picPaperSize(1).Left = picPaper.Left + picPaperSize(1).Height
    picPaperSize(1).Width = picPaper.Width - picPaperSize(1).Height
    
    picPaperSize(2).Top = picPaperSize(1).Top
    picPaperSize(2).Left = picPaperSize(0).Left
        
    '标尺
    DrawRuler picRulerH, picPaper.Left + 15
    DrawRuler picRulerV, picPaper.Top + 15
    
    With objReport.Fmts("_" & mbytCurrFmt)
        sta.Panels(2).Text = "打印机:" & objReport.打印机 & "   纸张:" & GetPaperName(.纸张, .W, .H) & " " & _
            IIF(.纸张 = 256, CInt(.W / Twip_mm) & "mm × " & CInt(.H / Twip_mm) & "mm", "") & _
            IIF(.纸向 = 1, "   纵向", "   横向")
    End With
    
    Me.Refresh
End Sub

Private Sub ShowScroll(Optional bytType As Byte = 3)
'功能:设置滚动条
'参数:bytType=3-两者都调整(缺省值),1-仅调整Hsc,2-仅调整Vsc
    
    If bytType = 3 Or bytType = 2 Then
        If picBack.ScaleHeight >= picPaper.Height + picPaperSize(2).Width * 2 Then
            scrVsc.Enabled = False
        Else
            scrVsc.Max = (picPaper.Height + picPaperSize(2).Width * 2 - picBack.ScaleHeight) / Screen.TwipsPerPixelX '转换为像素为单位
            Call ShowSize(0, picPaper.Left)
            scrVsc.Value = 0
            scrVsc.Enabled = True
        End If
    End If
    If bytType = 3 Or bytType = 1 Then
        If picBack.ScaleWidth >= picPaper.Width + picPaperSize(2).Width * 2 Then
            scrHsc.Enabled = False
        Else
            scrHsc.Max = (picPaper.Width + picPaperSize(2).Width * 2 - picBack.ScaleWidth) / Screen.TwipsPerPixelX
            Call ShowSize(picPaper.Top, 0)
            scrHsc.Value = 0
            scrHsc.Enabled = True
        End If
    End If
End Sub

Private Sub ShowSQLs()
'功能：根据objReport对象显示报表数据源及参数
'说明：为加快速度,自动分析SQL字段,实际用到时再打开数据。
    Dim tmpData As New RPTData
    Dim objNode As Object
    Dim arrFields() As String
    Dim i As Integer
    Dim strSource As String
    
    '显示多个数据源及字段
    tvwSQL.Nodes.Clear
    Set objNode = tvwSQL.Nodes.Add(, , "Root", objReport.名称, "Root")
    objNode.Selected = True
    objNode.Expanded = True
    
    For Each tmpData In objReport.Datas
        If tmpData.数据连接编号 > 0 Then
            '其他数据连接显示连接的名称
            strSource = GetDBConnectInfo(tmpData.数据连接编号)
            If strSource = "" Then
                strSource = tmpData.名称
            Else
                strSource = tmpData.名称 & "（" & strSource & "）"
            End If
        Else
            strSource = tmpData.名称
        End If
        
        If tmpData.类型 = 0 Then
            Set objNode = tvwSQL.Nodes.Add("Root", 4, "_" & tmpData.名称, strSource, "SQL_Custom")
        Else
            Set objNode = tvwSQL.Nodes.Add("Root", 4, "_" & tmpData.名称, strSource, "SQL_Group")
        End If
        objNode.Expanded = True
        
        '处理字段子项
        If tmpData.字段 <> "" Then
            arrFields = Split(tmpData.字段, "|")
            For i = 0 To UBound(arrFields)
                Select Case Split(arrFields(i), ",")(1)
                    Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR '文本型(Varchar2,Long)
                        Set objNode = tvwSQL.Nodes.Add("_" & tmpData.Key, 4, , Split(arrFields(i), ",")(0), "String")
                    Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, _
                        adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, _
                        adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt  '数字型(Numeric(a,b),Sum)
                        Set objNode = tvwSQL.Nodes.Add("_" & tmpData.Key, 4, , Split(arrFields(i), ",")(0), "Number")
                    Case adDBTimeStamp, adDBTime, adDBDate, adDate '日期型(Date)
                        Set objNode = tvwSQL.Nodes.Add("_" & tmpData.Key, 4, , Split(arrFields(i), ",")(0), "Date")
                    Case adBinary, adVarBinary, adLongVarBinary '二进制(Long Raw)
                        Set objNode = tvwSQL.Nodes.Add("_" & tmpData.Key, 4, , Split(arrFields(i), ",")(0), "Bin")
                    Case Else '其它
                        Set objNode = tvwSQL.Nodes.Add("_" & tmpData.Key, 4, , Split(arrFields(i), ",")(0), "Other")
                End Select
                objNode.Tag = Split(arrFields(i), ",")(1) '存放字段的类型！！！
            Next
        End If
    Next
    If Not tvwSQL.SelectedItem.Child Is Nothing Then tvwSQL.SelectedItem.Child.Selected = True
    Call tvwSQL_NodeClick(tvwSQL.SelectedItem)
End Sub

Private Sub tbr2_ButtonClick(ByVal Button As MSComctlLib.Button)
    Call NoneEdit
    Select Case Button.Key
        Case "Lock"
            mnuFormat_Lock_Click
        Case "Left", "Right", "Up", "Down", "Hsc", "Vsc"
            Call mnuFormat_DoAlign_Click(Button.Index - 1)
        Case "Width"
            mnuFormat_Width_Click
        Case "Height"
            mnuFormat_Height_Click
        Case "WH"
            mnuFormat_WH_Click
        Case "Scale"
            tbr2_ButtonMenuClick tbr2.Buttons("Scale").ButtonMenus(1)
    End Select
End Sub

Private Sub ClearXY()
    '清除标尺位置状态线
    If preHsc <> Empty Then
        picRulerH.Line (preHsc, 0)-(preHsc, picRulerH.ScaleHeight)
        preHsc = Empty
    End If
    If preVsc <> Empty Then
        picRulerV.Line (0, preVsc)-(picRulerH.ScaleWidth, preVsc)
        preVsc = Empty
    End If
    sta.Panels(3) = "位置"
End Sub

Private Sub DrawXY(X As Long, Y As Long)
'功能:处理标尺上的位置线
'参数:相对于PicPaper的位置
    If preHsc <> Empty Then picRulerH.Line (preHsc, 0)-(preHsc, picRulerH.ScaleHeight)
    If preVsc <> Empty Then picRulerV.Line (0, preVsc)-(picRulerV.ScaleWidth, preVsc)
    
    picRulerH.Line (X + picRulerH.ScaleHeight + 15 - Abs(picPaper.Left), 0)-(X + picRulerH.ScaleHeight + 15 - Abs(picPaper.Left), picRulerH.ScaleHeight)
    picRulerV.Line (0, Y + 15 - Abs(picPaper.Top))-(picRulerV.ScaleWidth, Y + 15 - Abs(picPaper.Top))
    preHsc = X + picRulerH.ScaleHeight + 15 - Abs(picPaper.Left)
    preVsc = Y + 15 - Abs(picPaper.Top)
    
    sta.Panels(3) = "X=" & Format(X / sgnMode / Twip_mm, "0.00") & "mm Y=" & Format(Y / sgnMode / Twip_mm, "0.00") & "mm"
End Sub

Private Sub TrueArea(ByRef Area As RECT)
'功能:将选择区域方向调为向右和向下,但范围不变
'参数:选择范围
    Dim X1 As Long, Y1 As Long, X2 As Long, Y2 As Long
    X1 = Area.Left: Y1 = Area.Top
    X2 = Area.Right: Y2 = Area.Bottom
    If X2 < X1 Then Area.Left = X2: Area.Right = X1
    If Y2 < Y1 Then Area.Top = Y2: Area.Bottom = Y1
End Sub

Private Sub tvwSQL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'复选处理
    Dim objNode As Object
    
    Set objNode = tvwSQL.HitTest(X, Y)
    
    If Not objNode Is Nothing Then
        If objNode.Key = "Root" Then
            objNode.Checked = False
        ElseIf IsType(Val(objNode.Tag), adLongVarBinary) Then '二进制字段不许加入汇总表格
            If objReport.Datas(objNode.Parent.Key).类型 = 1 Then
                objNode.Checked = False
            End If
        ElseIf objNode.Image = "Other" Then '其它型不允许加入表格
            objNode.Checked = False
        End If
    End If
End Sub

Private Sub tvwSQL_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim objNode As Object, blnCheck As Boolean
    
    Set objNode = Node
    
    If objNode.Key <> "Root" Then
        If objNode.Parent.Key = "Root" Then
            blnCheck = objNode.Checked
            Set objNode = objNode.Child
            Do While Not objNode Is Nothing
                'If Not IsType(Val(objNode.Tag), adLongVarBinary) And objNode.Image <> "Other" Then
                If objNode.Image <> "Other" And Not (objReport.Datas(objNode.Parent.Key).类型 = 1 And IsType(Val(objNode.Tag), adLongVarBinary)) Then
                    '二进制字段不许加入汇总表格
                    objNode.Checked = blnCheck
                    If blnCheck = True And objNode.Text = "分组标识" Then
                        objNode.Checked = False
                    End If
                Else
                    objNode.Checked = False
                End If
                Set objNode = objNode.Next
            Loop
        Else
            If objNode.Checked Then
                'If Not IsType(Val(objNode.Tag), adLongVarBinary) And objNode.Image <> "Other" Then
                If objNode.Image <> "Other" Then
                    objNode.Parent.Checked = True
                End If
            Else
                blnCheck = False
                Set objNode = objNode.Parent.Child
                Do While Not objNode Is Nothing
                    blnCheck = blnCheck Or objNode.Checked
                    If objNode.Next Is Nothing Then objNode.Parent.Checked = blnCheck
                    Set objNode = objNode.Next
                Loop
            End If
        End If
    End If
End Sub

Private Sub tvwSQL_NodeClick(ByVal Node As MSComctlLib.Node)
'功能：显示当前数据源的参数清单
    Dim objItem As Object
    Dim tmpPar As RPTPar
    Dim strKey As String
    
    lvwPar.ListItems.Clear
    
    If Node.Key <> "Root" Then
        If Node.Children = 0 Then
            strKey = Node.Parent.Key
        Else
            strKey = Node.Key
        End If
        For Each tmpPar In objReport.Datas(strKey).Pars
            Set objItem = lvwPar.ListItems.Add(, "_" & tmpPar.序号, tmpPar.序号)
            objItem.SubItems(1) = tmpPar.名称
            Select Case tmpPar.类型
                Case 0
                    objItem.SubItems(2) = "字符"
                    objItem.SmallIcon = "String"
                Case 1
                    objItem.SubItems(2) = "数字"
                    objItem.SmallIcon = "Number"
                Case 2
                    objItem.SubItems(2) = "日期"
                    objItem.SmallIcon = "Date"
                Case 3
                    objItem.SubItems(2) = "无类型"
                    objItem.SmallIcon = "Other"
            End Select
            objItem.SubItems(3) = tmpPar.缺省值
        Next
    End If
End Sub

Private Sub AddReportItem(Optional ByVal blnDrop As Boolean = False, Optional ByVal objParent As Object)
'功能:添加一个报表项目
'使用参数:
'   SelArea=当前选择范围(从菜单或拖动增加要手动设置)
'   bytCurTool=当前要添加的元素类型/或tvwSQL.SelectedItem=当前要添加的数据表格或项目
    Dim newObj As Control, objNode As Object, tmpItem As RPTItem
    Dim intCols As Integer, i As Integer, j As Integer, k As Integer, l As Integer
    Dim X As Integer, Y As Integer, Z As Integer, sngWidth As Single
    Dim bytAlign As Byte, arrAlign() As Byte, Str表头 As String, tmpID As RelatID
    Dim intMaxIDtmp As Integer, intCurIDtmp As Integer
    
    If cboFormat.ComboItems("_" & mbytCurrFmt) Is Nothing Then
        MsgBox "请选择要添加报表元素的报表格式！", vbInformation, App.Title
        Exit Sub
    End If
    If (bytCurTool = 6 Or bytCurTool = 8) And Not objParent Is Nothing Then
        MsgBox "卡片中不允许放入" & IIF(bytCurTool = 6, "图表", "卡片") & "！", vbInformation, App.Title
        objParent.Refresh
        Exit Sub
    End If
    intMaxIDtmp = intMaxID: intCurIDtmp = intCurID
    intMaxID = intMaxID + 1
    intCurID = intMaxID
    
    If Not blnDrop Then
        Select Case bytCurTool
            Case 1 '线条(lblLine)
                Load lblLine(intCurID)
                Set newObj = lblLine(intCurID)
                
                '加入数据对象
                objReport.Items.Add intCurID, mbytCurrFmt, GetNextName(1), 0, 1, 0, "", 0, "", "", _
                    Format(selArea.Left / sgnMode, "0.00"), Format(selArea.Top / sgnMode, "0.00"), _
                    Format((selArea.Right - selArea.Left) / sgnMode, "0.00"), _
                    Format((selArea.Bottom - selArea.Top) / sgnMode, "0.00"), _
                    0, 0, True, "", 0, False, False, False, 0, 0, 0, False, 0, "", "", "", False, _
                    False, , False, , , , "_" & intCurID
            Case 3  '标签(lbl)
                Load lbl(intCurID)
                Set newObj = lbl(intCurID)
                If bytCurTool = 3 Then
                    newObj.Caption = GetNextName(2)
                Else
                    newObj.Caption = ""
                End If
                
                '加入数据对象
                objReport.Items.Add intCurID, mbytCurrFmt, GetNextName(IIF(bytCurTool = 2, 10, 2)), _
                    0, IIF(bytCurTool = 2, 10, 2), 0, "", 0, GetNextName(IIF(bytCurTool = 2, 10, 2)), _
                    "", Format(selArea.Left / sgnMode, "0.00"), Format(selArea.Top / sgnMode, "0.00"), _
                    Format((selArea.Right - selArea.Left) / sgnMode, "0.00"), _
                    Format((selArea.Bottom - selArea.Top) / sgnMode, "0.00"), _
                    0, 0, IIF(bytCurTool = 3, True, False), "宋体", 9, False, False, False, 0, 0, &HFFFFFF, _
                    False, 0, "", "", "", False, False, , False, , , , "_" & intCurID
            Case 2 '框线(shp)
                Load lblshp(intCurID)
                Load Shp(intCurID)
                Set newObj = Shp(intCurID)
                lblshp(intCurID).BackColor = picPaper.BackColor
                
                '加入数据对象
                objReport.Items.Add intCurID, mbytCurrFmt, GetNextName(10), 0, 10, 0, "", 0, GetNextName(10), _
                    "", Format(selArea.Left / sgnMode, "0.00"), Format(selArea.Top / sgnMode, "0.00"), _
                    Format((selArea.Right - selArea.Left) / sgnMode, "0.00"), _
                    Format((selArea.Bottom - selArea.Top) / sgnMode, "0.00"), _
                    0, 0, IIF(bytCurTool = 3, True, False), "宋体", 9, False, False, False, 0, 0, &HFFFFFF, _
                    False, 0, "", "", "", False, False, , False, , , , "_" & intCurID
            Case 4 '图片
                Load img(intCurID)
                Set newObj = img(intCurID)
                newObj.BorderStyle = 1
                
                '照片字段
                '加入数据对象
                objReport.Items.Add intCurID, mbytCurrFmt, GetNextName(11), 0, 11, 0, "", 0, "", "", _
                    Format(selArea.Left / sgnMode, "0.00"), Format(selArea.Top / sgnMode, "0.00"), _
                    Format((selArea.Right - selArea.Left) / sgnMode, "0.00"), _
                    Format((selArea.Bottom - selArea.Top) / sgnMode, "0.00"), _
                    0, 0, False, "宋体", 9, False, False, False, 0, 0, &HFFFFFF, True, 0, "", "", "", _
                    False, False, , False, , , , "_" & intCurID
            Case 8 '卡片
                Load pic(intCurID)
                Set newObj = pic(intCurID)
                newObj.BorderStyle = 1
                
                '照片字段
                '加入数据对象
                objReport.Items.Add intCurID, mbytCurrFmt, GetNextName(14), 0, 14, 0, "", 0, "", "", _
                    Format(selArea.Left / sgnMode, "0.00"), Format(selArea.Top / sgnMode, "0.00"), _
                    Format((selArea.Right - selArea.Left) / sgnMode, "0.00"), _
                    Format((selArea.Bottom - selArea.Top) / sgnMode, "0.00"), _
                    0, 0, False, "宋体", 9, False, False, False, 0, 0, &HFFFFFF, True, 0, "", "", "", _
                    False, False, , False, , , , "_" & intCurID
            Case 5 '任意表格(特殊)
                Load msh(intCurID)
                            
                '缺省1个表头行,数据行由高度计算(至少2行)。
                msh(intCurID).Rows = Abs(Int(-(selArea.Bottom - selArea.Top) / (255 * sgnMode))) - 1
                If Not objReport.票据 Then
                    msh(intCurID).FixedRows = msh(intCurID).Rows - 2
                Else
                    msh(intCurID).FixedRows = 1
                End If
                msh(intCurID).FixedCols = 0
                '列数由宽度计算,至少1列
                msh(intCurID).Cols = Abs(Int(-(selArea.Right - selArea.Left) / 1000))
                If msh(intCurID).Cols > 1 Then msh(intCurID).Cols = msh(intCurID).Cols - 1
                
                msh(intCurID).TextMatrix(0, 0) = "表格" & msh.count - 1
                                
                msh(intCurID).Row = 0: msh(intCurID).Col = 0
                SetHeadCenter msh(intCurID)
                                
                Set newObj = msh(intCurID)
                
                '加入数据对象(1:内容为空,无数据源,2:分栏=1,3:行高=255)
                Set tmpItem = objReport.Items.Add(intCurID, mbytCurrFmt, GetNextName(4), 0, 4, 0, "", 0, "", "", _
                    Format(selArea.Left / sgnMode, "0.00"), Format(selArea.Top / sgnMode, "0.00"), _
                    Format((selArea.Right - selArea.Left) / sgnMode, "0.00"), _
                    Format((selArea.Bottom - selArea.Top) / sgnMode, "0.00"), _
                    255, 0, False, "宋体", 9, False, False, False, 0, 0, &HFFFFFF, True, 1, "", "", "", False, _
                    False, , False, , , , "_" & intCurID)
                
                For i = 0 To msh(intCurID).Cols - 1
                    msh(intCurID).ColWidth(i) = 1000 * sgnMode
                    msh(intCurID).ColAlignment(i) = 1
                    
                    intMaxID = intMaxID + 1
                    Str表头 = ""
                    For j = 0 To msh(intCurID).FixedRows - 1
                        msh(intCurID).Row = j: msh(intCurID).Col = i
                        Str表头 = Str表头 & "|" & msh(intCurID).CellAlignment & "^" & msh(intCurID).RowHeight(j) & "^#"
                    Next
                    Str表头 = Mid(Str表头, 2)
                    
                    '表格子项的加入(1:内容为空,无数据项,2:不汇总)
                    objReport.Items.Add intMaxID, mbytCurrFmt, "表列" & intMaxID, intCurID, 6, i, "", 0, "", _
                        Str表头, _
                        0, 0, msh(intCurID).ColWidth(i) / sgnMode, 0, _
                        0, 0, False, "", 0, False, False, False, 0, 0, 0, False, 0, "", "", "", False, False, _
                         , False, , , , "_" & intMaxID
                    
                    tmpItem.SubIDs.Add intMaxID, "_" & intMaxID
                Next
                '可以合并
                For i = 0 To msh(intCurID).FixedRows - 1
                    msh(intCurID).MergeRow(i) = True
                Next
                For i = 0 To msh(intCurID).Cols - 1
                    msh(intCurID).MergeCol(i) = True
                Next
                
            Case 6 '图表@@@
                Load Chart(intCurID)
                Set newObj = Chart(intCurID)
                
                '缺省为折线图,图例在东
                objReport.Items.Add intCurID, mbytCurrFmt, GetNextName(12), 0, 12, 1, "", 0, "", "", _
                    Format(selArea.Left / sgnMode, "0.00"), Format(selArea.Top / sgnMode, "0.00"), _
                    Format((selArea.Right - selArea.Left) / sgnMode, "0.00"), _
                    Format((selArea.Bottom - selArea.Top) / sgnMode, "0.00"), _
                    0, 1, True, "宋体", 9, False, True, False, 0, 0, &HFFFFFF, False, 0, "", "", "", _
                    False, False, , False, , , , "_" & intCurID
            Case 7 '条码
                Load ImgCode(intCurID)
                Set newObj = ImgCode(intCurID)
                newObj.BorderStyle = 0
                
                Unload frmFlash '强制初始Picture，不然切换绘制有问题
                Set newObj.Picture = DrawBarCode128Auto(frmFlash.picTemp, "1234567890", sngWidth, 2, True)
                
                '加入数据对象
                objReport.Items.Add intCurID, mbytCurrFmt, GetNextName(13), 0, 13, 3, "", 0, "", "100", _
                    Format(selArea.Left / sgnMode, "0.00"), Format(selArea.Top / sgnMode, "0.00"), _
                    Format(Me.ScaleX(sngWidth, vbMillimeters, vbTwips), "0.00"), _
                    Format((selArea.Bottom - selArea.Top) / sgnMode, "0.00"), _
                    2, 0, False, "宋体", 9, False, False, False, 0, 0, &HFFFFFF, False, 0, "", "", "", _
                    False, False, , False, , , , "_" & intCurID
        End Select
                
        tbrTool.Buttons(1).Value = tbrPressed
        tbrTool_ButtonClick tbrTool.Buttons(1)
    Else
        If tvwSQL.SelectedItem.Children = 0 Then
            If Not objParent Is Nothing Then
                If objReport.Items("_" & objParent.Index).数据源 = "" Then
                    If MsgBox("当前卡片未绑定数据源，绑定后将分组打印多张卡片，数据源中存在""分组标识""字段则""分组标识""相同的为一组,否则一行数据为一组；" & vbCrLf & _
                         "不绑定则只打印一张卡片，是否绑定数据源""" & tvwSQL.SelectedItem.Parent.Text & """?" _
                        , vbQuestion + vbYesNo + vbDefaultButton1, App.Title) = vbYes Then
                        objReport.Items("_" & objParent.Index).数据源 = mdlPublic.GetStdNodeText(tvwSQL.SelectedItem.Parent.Text)
                    End If
                End If
            End If
            '单数据项
            Load lbl(intCurID)
            Set newObj = lbl(intCurID)
            newObj.Caption = "[" & LevelText(tvwSQL.SelectedItem) & "]" '[]号不存入数据库
            If IsType(Val(tvwSQL.SelectedItem.Tag), adLongVarBinary) Then
                '照片字段
                newObj.BorderStyle = 1
                selArea.Bottom = selArea.Top + 1500
                selArea.Right = selArea.Left + 1300
                
                '加入数据对象
                objReport.Items.Add intCurID, mbytCurrFmt, GetNextName(2), 0, 2, 0, "", 0, newObj.Caption, "", _
                    Format(selArea.Left / sgnMode, "0.00"), Format(selArea.Top / sgnMode, "0.00"), _
                    Format((selArea.Right - selArea.Left) / sgnMode, "0.00"), _
                    Format((selArea.Bottom - selArea.Top) / sgnMode, "0.00"), _
                    0, 0, False, "宋体", 9, False, False, False, 0, 0, &HFFFFFF, True, 0, "", "", "", _
                    False, False, , False, , , , "_" & intCurID
            Else
                newObj.Caption = mdlPublic.GetStdNodeText(tvwSQL.SelectedItem.Text) & ":" & newObj.Caption
                selArea.Bottom = selArea.Top + lbl(0).Height * sgnMode
                selArea.Right = selArea.Left + TextWidth(lbl(intCurID).Caption & "字") * sgnMode
                
                Select Case tvwSQL.SelectedItem.Tag
                Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger _
                    , adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                    bytAlign = 2 '数字缺省右对齐
                    lbl(intCurID).Alignment = 1
                Case adDBTimeStamp, adDBTime, adDBDate, adDate
                    bytAlign = 1 '日期缺省中对齐
                    lbl(intCurID).Alignment = 2
                End Select
                
                '加入数据对象
                objReport.Items.Add intCurID, mbytCurrFmt, GetNextName(2), 0, 2, 0, "", 0, newObj.Caption, "", _
                    Format(selArea.Left / sgnMode, "0.00"), Format(selArea.Top / sgnMode, "0.00"), _
                    Format((selArea.Right - selArea.Left) / sgnMode, "0.00"), _
                    Format((selArea.Bottom - selArea.Top) / sgnMode, "0.00"), _
                    0, bytAlign, True, "宋体", 9, False, False, False, 0, 0, &HFFFFFF, False, 0, "", "", "", _
                    False, False, , False, , , , "_" & intCurID
            End If
        Else
            If objReport.Datas(tvwSQL.SelectedItem.Key).类型 = 0 Then
                '任意表只能放入和卡片数据源相同的数据源
                If Not objParent Is Nothing Then
                    If mdlPublic.GetStdNodeText(tvwSQL.SelectedItem.Text) <> objReport.Items("_" & objParent.Index).数据源 _
                        And objReport.Items("_" & objParent.Index).数据源 <> "" Then
                        MsgBox "卡片绑定了数据源，所以只能加入和卡片相同数据源的表格！", vbInformation, App.Title
                        intMaxID = intMaxIDtmp: intCurID = intCurIDtmp
                        Exit Sub
                    End If
                    '如果卡片未指定数据源，则自动指定
                    If objReport.Items("_" & objParent.Index).数据源 = "" Then
                        If MsgBox("当前卡片未绑定数据源，绑定后将分组打印多张卡片，数据源中存在""分组标识""字段则""分组标识""相同的为一组,否则一行数据为一组；" & vbCrLf & _
                             "不绑定则只打印一张卡片，是否绑定数据源""" & tvwSQL.SelectedItem.Text & """?" _
                            , vbQuestion + vbYesNo + vbDefaultButton1, App.Title) = vbYes Then
                            objReport.Items("_" & objParent.Index).数据源 = mdlPublic.GetStdNodeText(tvwSQL.SelectedItem.Text)
                        End If
                    End If
                End If
                '任意表格
                Load msh(intCurID)
                            
                '缺省1个表头行,5个数据行。
                selArea.Bottom = selArea.Top + (1545 * sgnMode) '255 * 3 + 15
                msh(intCurID).Rows = 6
                msh(intCurID).FixedRows = 1
                msh(intCurID).FixedCols = 0
                msh(intCurID).SelectionMode = flexSelectionFree
                
                '宽度由列数计算,至少1列
                msh(intCurID).Cols = Abs(Int(-(selArea.Right - selArea.Left) / (1000 * sgnMode)))
                
                i = 0
                Set objNode = tvwSQL.SelectedItem.Child
                Do While Not objNode Is Nothing
                    If objNode.Checked Then
                        If i = 0 Then
                            ReDim arrAlign(i)
                        Else
                            ReDim Preserve arrAlign(i)
                        End If
                        Select Case objNode.Tag
                        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger _
                            , adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                            arrAlign(i) = 2
                        Case adDBTimeStamp, adDBTime, adDBDate, adDate
                            arrAlign(i) = 1
                        Case adBinary, adVarBinary, adLongVarBinary
                            arrAlign(i) = 1
                        Case Else
                            arrAlign(i) = 0
                        End Select
                        
                        i = i + 1
                        
                        msh(intCurID).Cols = i
                        msh(intCurID).ColWidth(i - 1) = 1000 * sgnMode
                        msh(intCurID).ColAlignment(i - 1) = 1
                        msh(intCurID).TextMatrix(0, i - 1) = objNode.Text
                        msh(intCurID).TextMatrix(1, i - 1) = "[" & LevelText(objNode) & "]" '[]号要存入数据库
                    End If
                    Set objNode = objNode.Next
                Loop
                selArea.Right = selArea.Left + msh(intCurID).Cols * (1000 * sgnMode) + 30
                
                msh(intCurID).Row = 0: msh(intCurID).Col = 0
                SetHeadCenter msh(intCurID)
                
                Set newObj = msh(intCurID)
                
                '加入数据对象(1:内容为数据源名)
                Set tmpItem = objReport.Items.Add(intCurID, mbytCurrFmt, GetNextName(4), 0, 4, 0, "", 0, _
                    mdlPublic.GetStdNodeText(tvwSQL.SelectedItem.Text), "", _
                    Format(selArea.Left / sgnMode, "0.00"), Format(selArea.Top / sgnMode, "0.00"), _
                    Format((selArea.Right - selArea.Left) / sgnMode, "0.00"), _
                    Format((selArea.Bottom - selArea.Top) / sgnMode, "0.00"), _
                    255, 0, False, "宋体", 9, False, False, False, 0, 0, &HFFFFFF, True, 1, "", "", "", _
                    False, False, , False, , , , "_" & intCurID)
                
                For i = 0 To msh(intCurID).Cols - 1
                    intMaxID = intMaxID + 1
                    '表格子项的加入(1:无汇总)
                    Str表头 = ""
                    For j = 0 To msh(intCurID).FixedRows - 1
                        msh(intCurID).Row = j: msh(intCurID).Col = i
                        Str表头 = Str表头 & "|" & msh(intCurID).CellAlignment & _
                                    "^" & msh(intCurID).RowHeight(j) & "^" & msh(intCurID).TextMatrix(j, i)
                    Next
                    Str表头 = Mid(Str表头, 2)
                    objReport.Items.Add intMaxID, mbytCurrFmt, "表列" & intMaxID, intCurID, 6, i, _
                        "", 0, msh(intCurID).TextMatrix(1, i), Str表头, 0, 0, msh(intCurID).ColWidth(i) / sgnMode, _
                        0, 0, arrAlign(i), False, "", 0, False, False, False, 0, 0, 0, False, 0, "", "", "", _
                        False, False, , False, , , , "_" & intMaxID
                    
                    tmpItem.SubIDs.Add intMaxID, "_" & intMaxID
                Next
                '可以合并
                For i = 0 To msh(intCurID).FixedRows - 1
                    msh(intCurID).MergeRow(i) = True
                Next
                For i = 0 To msh(intCurID).Cols - 1
                    msh(intCurID).MergeCol(i) = True
                Next
                
            Else
                '卡片不允许放入汇总表
                If Not objParent Is Nothing Then
                    MsgBox "卡片中不允许加入汇总表！", vbInformation, App.Title
                    Exit Sub
                End If
                '汇总表格
                Load msh(intCurID)
                
                'x,y,z分别表示纵向/横向/统计分类项数
                Set objNode = tvwSQL.SelectedItem.Child
                Do While Not objNode Is Nothing
                    If objNode.Checked Then
                        '不可能有二进制及其它型
                        Select Case objNode.Tag
                            Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR, _
                                adDBTimeStamp, adDBTime, adDBDate, adDate  '文本及日期型
                                X = X + 1
                            Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, _
                                adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, _
                                adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt  '数字型
                                Z = Z + 1
                        End Select
                    End If
                    Set objNode = objNode.Next
                Loop
                '如果可能,分配一项为横向分类
                If X >= 2 Then X = X - 1: Y = Y + 1
                
                '建立表格
                With msh(intCurID)
                    '表格框架
                    .Rows = Y + 1 + X * 5
                    .FixedRows = Y + 1
                    .Cols = X + IIF(Y = 0, Z, Z * 2)
                    .FixedCols = X
                    selArea.Bottom = selArea.Top + .Rows * (255 * sgnMode) + 60
                    selArea.Right = selArea.Left + .Cols * (1000 * sgnMode) + 60
                    For i = 0 To .Cols - 1
                        .ColWidth(i) = 1000 * sgnMode
                        .ColAlignment(i) = 1
                    Next
                    For i = 0 To .FixedCols - 1
                        .MergeCol(i) = True
                    Next
                    For i = 0 To .FixedRows - 2
                        .MergeRow(i) = True
                    Next
                    
                    '表格内容及报表对象
                    
                    '加入数据对象(1:内容为数据源名)
                    Set tmpItem = objReport.Items.Add(intCurID, mbytCurrFmt, GetNextName(5), 0, 5, 0, "", 0, _
                        mdlPublic.GetStdNodeText(tvwSQL.SelectedItem.Text), "", Format(selArea.Left / sgnMode, "0.00"), _
                        Format(selArea.Top / sgnMode, "0.00"), Format((selArea.Right - selArea.Left) / sgnMode, "0.00"), _
                        Format((selArea.Bottom - selArea.Top) / sgnMode, "0.00"), 255, 0, False, "宋体", 9, False, False, _
                        False, 0, 0, &HFFFFFF, True, 0, "", "", "", False, False, , False, , , , "_" & intCurID)
                    
                    i = 0: j = 0
                    Set objNode = tvwSQL.SelectedItem.Child
                    Do While Not objNode Is Nothing
                        If objNode.Checked Then
                            intMaxID = intMaxID + 1
                            Select Case objNode.Tag
                                Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR, _
                                    adDBTimeStamp, adDBTime, adDBDate, adDate  '字符及日期
                                    i = i + 1
                                    If i <= X Then
                                        .TextMatrix(.FixedRows - 1, i - 1) = "[" & mdlPublic.GetStdNodeText(objNode.Text) & "]"  '[]不存入数据库
                                        For k = .FixedRows To .Rows - 1
                                            .TextMatrix(k, i - 1) = mdlPublic.GetStdNodeText(objNode.Text)
                                        Next
                                        
                                        '表格子项的加入(1:无汇总,2:内容直接为字段名,不为A.B的形式,因主项已存放)
                                        objReport.Items.Add intMaxID, mbytCurrFmt, "纵列" & intMaxID, intCurID, 7, i - 1, _
                                            "", 0, mdlPublic.GetStdNodeText(objNode.Text), "", 0, 0, 1000, 0, 255, 0, False, _
                                            "", 0, False, False, False, 0, 0, 0, False, 0, "", "", "", False, False, , _
                                            False, , , , "_" & intMaxID
                                    Else
                                        For k = 0 To .FixedRows - 2
                                            For l = 0 To .FixedCols - 1
                                                .TextMatrix(k, l) = "[" & mdlPublic.GetStdNodeText(objNode.Text) & "]"
                                            Next
                                            For l = .FixedCols To .Cols - 1
                                                .TextMatrix(k, l) = mdlPublic.GetStdNodeText(objNode.Text)
                                            Next
                                        Next
                                        objReport.Items.Add intMaxID, mbytCurrFmt, "横列" & intMaxID, intCurID, 8, i - X - 1, _
                                            "", 0, mdlPublic.GetStdNodeText(objNode.Text), "", 0, 0, 1000, 0, 255, 0, False, _
                                            "", 0, False, False, False, 0, 0, 0, False, 0, "", "", "", False, False, , _
                                            False, , , , "_" & intMaxID
                                    End If
                                Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt _
                                    , adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt _
                                    , adUnsignedTinyInt
                                    For k = .FixedCols To .Cols - 1 Step Z
                                        .TextMatrix(.FixedRows - 1, k + j) = "[" & mdlPublic.GetStdNodeText(objNode.Text) & "]"
                                    Next
                                    '缺省右对齐
                                    objReport.Items.Add intMaxID, mbytCurrFmt, "统计项" & intMaxID, intCurID, 9, j, _
                                        "", 0, mdlPublic.GetStdNodeText(objNode.Text), "", 0, 0, 1000, 0, 255, 2, _
                                        False, "", 0, False, False, False, 0, 0, 0, False, 0, "", "", "", False, _
                                        False, , False, , , , "_" & intMaxID
                                    j = j + 1 'j为序号
                            End Select
                            tmpItem.SubIDs.Add intMaxID, "_" & intMaxID
                        End If
                        Set objNode = objNode.Next
                    Loop
                End With
                SetHeadCenter msh(intCurID)
                Set newObj = msh(intCurID)
            End If
        End If
    End If
    
    '显示报表项目
    On Error Resume Next
    If Not objParent Is Nothing Then
        Set newObj.Container = objParent
        objReport.Items("_" & newObj.Index).父ID = objParent.Index
        For Each tmpID In objReport.Items("_" & newObj.Index).SubIDs
            With objReport.Items("_" & tmpID.id)
                .父ID = objParent.Index
            End With
        Next
    End If
    newObj.Left = selArea.Left
    newObj.Top = selArea.Top
    newObj.Width = objReport.Items("_" & newObj.Index).W * sgnMode
    newObj.Height = objReport.Items("_" & newObj.Index).H * sgnMode
    'newObj.Width = Abs(selArea.Right - selArea.Left)
    'newObj.Height = Abs(selArea.Bottom - selArea.Top)
    
    If UCase(TypeName(newObj)) = "LABEL" And objReport.Items("_" & newObj.Index).类型 <> 1 Then
        newObj.AutoSize = objReport.Items("_" & newObj.Index).自调
        objReport.Items("_" & newObj.Index).X = Format(newObj.Left / sgnMode, "0.00")
        objReport.Items("_" & newObj.Index).Y = Format(newObj.Top / sgnMode, "0.00")
        objReport.Items("_" & newObj.Index).W = Format(newObj.Width / sgnMode, "0.00")
        objReport.Items("_" & newObj.Index).H = Format(newObj.Height / sgnMode, "0.00")
    End If
    
    If objReport.Items("_" & newObj.Index).类型 = 12 Then '@@@
        Call SetChartStyleAndData(newObj, objReport.Items("_" & newObj.Index), , sgnMode, True)
    Else
        newObj.FontSize = Format(newObj.FontSize * sgnMode, "0.0")
    End If
    
    '初始表格行高,便于以后设置字体及行高时调整(切记,Load时一定加入)。
    Select Case objReport.Items("_" & newObj.Index).类型
        Case 4, 5
            Call InitRowHeight(newObj.Index)
            Call SetGridLine(newObj.Index)
            Call ReShowGrid(newObj.Index)
            objReport.Items("_" & newObj.Index).W = Format(newObj.Width / sgnMode, "0.00")
            objReport.Items("_" & newObj.Index).H = Format(newObj.Height / sgnMode, "0.00")
        Case 10
            newObj.BorderStyle = 1
            newObj.BackStyle = 0
            '如果是框线则同步修改lblshp的位置
            lblshp(newObj.Index).Left = newObj.Left
            lblshp(newObj.Index).Top = newObj.Top
            lblshp(newObj.Index).Width = newObj.Width
            lblshp(newObj.Index).Height = newObj.Height
            lblshp(newObj.Index).Visible = True
            lblshp(newObj.Index).ZOrder 1
            Call DrawFrame(newObj)
    End Select
    
    If objReport.Items("_" & newObj.Index).类型 <> 10 Then newObj.ZOrder

    newObj.Visible = True
    
    If GetSelNum > 0 Then SelClear
    Call SelItem(newObj.Index, True)
    
    Call ShowAttrib(newObj.Index)
    
    '新增元素后取消元素的锁定
    If mnuFormat_Lock.Checked Then mnuFormat_Lock.Checked = False
    blnLock = False
    If mnuFormat_Lock.Checked And tbr2.Buttons("Lock").Value = tbrPressed Then
        tbr2.Buttons("Lock").Value = tbrUnpressed
    End If
    Call SetLock(blnLock)
    
    '新增元素后默认不锁定元素
    If mnuFormat_Lock.Checked Then mnuFormat_Lock.Checked = False
    blnLock = False
    If Not mnuFormat_Lock.Checked And tbr2.Buttons("Lock").Value = tbrPressed Then
        tbr2.Buttons("Lock").Value = tbrUnpressed
    End If
    Call SetLock(blnLock)
    
    picPaper.SetFocus
End Sub

Private Function SelClear() As Integer
'功能:清除报表中选中报表元素的选择状态
'返回:被清除的项目个数
    Dim tmpObj As PictureBox
    
    For Each tmpObj In lblSize
        If tmpObj.Index <> 0 Then
            If tmpObj.Index Mod 8 = 1 Then '上中标志
                Select Case objReport.Items("_" & tmpObj.Tag).类型
                    Case 1
                        lblLine(tmpObj.Tag).Tag = ""
                    Case 2, 3
                        lbl(tmpObj.Tag).Tag = ""
                    Case 10
                        Shp(tmpObj.Tag).Tag = ""
                    Case 4, 5
                        msh(tmpObj.Tag).Tag = ""
                        Call ResetColor(tmpObj.Tag)
                        Call SetGridLine(tmpObj.Tag)
                        If objReport.Items("_" & tmpObj.Tag).类型 = 4 Then
                            Call CustomColColor(tmpObj.Tag, -9)
                            Call SetCopyGrid(tmpObj.Tag)
                        End If
                    Case 11
                        img(tmpObj.Tag).Tag = ""
                    Case 12 '@@@
                        Chart(tmpObj.Tag).Tag = ""
                    Case 13
                        ImgCode(tmpObj.Tag).Tag = ""
                    Case 14
                        pic(tmpObj.Tag).Tag = ""
                End Select
                SelClear = SelClear + 1
            End If
            Unload lblSize(tmpObj.Index)
        End If
    Next
    Set objLastSel = Nothing: intCurID = 0

    Call ShowAttrib
    
    selCell.Col1 = -1: selCell.Col2 = -1: selCell.Row1 = -1: selCell.Row2 = -1: selCell.Row = -1
    Call ShowPaperInfo
    intCurCol = -1
    
    picPaper.SetFocus
End Function

Private Sub SelMove(idx As Integer)
'功能：移动选择元素的点到正确位置
    Dim i As Integer
    Dim intBeginIdx As Integer
    Dim ObjSel As Control, tmpID As RelatID
    Dim blnUse(7) As Boolean '控制各个尺寸标志是否显示
    Dim lngTmp As Long
    Dim lngtmp1 As Long
    
    Select Case objReport.Items("_" & idx).类型
        Case 2, 3 '文本标签及数据标签用同名控件
            Set ObjSel = lbl(idx)
        Case 10
            Set ObjSel = Shp(idx)
        Case 1 '线条
            With lblLine(idx)
                If .Width > .Height Then
                    '横向
                    For i = 0 To UBound(blnUse)
                        If Not (i = 2 Or i = 6) Then blnUse(i) = False
                    Next
                Else
                    '纵向
                    For i = 0 To UBound(blnUse)
                        If Not (i = 0 Or i = 4) Then blnUse(i) = False
                    Next
                End If
            End With
            Set ObjSel = lblLine(idx)
        Case 4, 5 '汇总表及任意表用同名控件
            Set ObjSel = msh(idx)
            '分栏控件置前(如果有)
            For Each tmpID In objReport.Items("_" & idx).CopyIDs
                msh(tmpID.id).ZOrder
            Next
        Case 11
            Set ObjSel = img(idx)
        Case 12 '@@@
            Set ObjSel = Chart(idx)
        Case 13
            Set ObjSel = ImgCode(idx)
        Case 14
            Set ObjSel = pic(idx)
    End Select

    '如果已经选中,才移动
    If Mid(ObjSel.Tag, 1, 2) = "S_" Then
        
        Set objLastSel = ObjSel: intCurID = idx
        
        
        '重新计算实际位置及相关坐标2002-03-26
        If ObjSel.Container.name <> "picPaper" Then
            lngTmp = ObjSel.Container.Top
            lngtmp1 = ObjSel.Container.Left
        End If
        
        intBeginIdx = CInt(Mid(ObjSel.Tag, 3))
        
         '元素控件的Tag记录选择标志的起始索引
        ObjSel.ZOrder IIF(objReport.Items("_" & ObjSel.Index).类型 <> 10, 0, 1)
        If objReport.Items("_" & ObjSel.Index).类型 = 10 Then lblshp(ObjSel.Index).ZOrder 1

        With WinProperty
            .H = lblSize(0).Height / Screen.TwipsPerPixelX
            .W = lblSize(0).Width / Screen.TwipsPerPixelX
        End With
        
        For i = intBeginIdx To intBeginIdx + 7 '选择标志从"上中"开始,"顺时针"处理
            Select Case IIF(i Mod 8 <> 0, i Mod 8, 8) '定位选择边框的位置
                'zyb#Modify
                '改用MoveWindow()解决
                
                Case 1 '上中
                    With WinProperty
                        .T = (ObjSel.Top + lngTmp - lblSize(i).Height) / Screen.TwipsPerPixelX
                        .l = (ObjSel.Left + lngtmp1 + (ObjSel.Width - lblSize(i).Width) / 2) / Screen.TwipsPerPixelX
                        Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                    End With
                Case 2 '上右
                    With WinProperty
                        .T = (ObjSel.Top + lngTmp - lblSize(i).Height) / Screen.TwipsPerPixelX
                        .l = (ObjSel.Left + lngtmp1 + ObjSel.Width) / Screen.TwipsPerPixelX
                        Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                    End With
                Case 3 '右中
                    With WinProperty
                        .T = (ObjSel.Top + lngTmp + (ObjSel.Height - lblSize(i).Height) / 2) / Screen.TwipsPerPixelX
                        .l = (ObjSel.Left + lngtmp1 + ObjSel.Width) / Screen.TwipsPerPixelX
                        Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                    End With
                Case 4 '右下
                    With WinProperty
                        .T = (ObjSel.Top + lngTmp + ObjSel.Height) / Screen.TwipsPerPixelX
                        .l = (ObjSel.Left + lngtmp1 + ObjSel.Width) / Screen.TwipsPerPixelX
                        Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                    End With
                Case 5 '下中
                    With WinProperty
                        .T = (ObjSel.Top + lngTmp + ObjSel.Height) / Screen.TwipsPerPixelX
                        .l = (ObjSel.Left + lngtmp1 + (ObjSel.Width - lblSize(i).Width) / 2) / Screen.TwipsPerPixelX
                        Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                    End With
                Case 6 '左下
                    With WinProperty
                        .T = (ObjSel.Top + lngTmp + ObjSel.Height) / Screen.TwipsPerPixelX
                        .l = (ObjSel.Left + lngtmp1 - lblSize(i).Width) / Screen.TwipsPerPixelX
                        Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                    End With
                Case 7 '左中
                    With WinProperty
                        .T = (ObjSel.Top + lngTmp + (ObjSel.Height - lblSize(i).Height) / 2) / Screen.TwipsPerPixelX
                        .l = (ObjSel.Left + lngtmp1 - lblSize(i).Width) / Screen.TwipsPerPixelX
                        Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                    End With
                Case 8 '左上
                    With WinProperty
                        .T = (ObjSel.Top + lngTmp - lblSize(i).Height) / Screen.TwipsPerPixelX
                        .l = (ObjSel.Left + lngtmp1 - lblSize(i).Width) / Screen.TwipsPerPixelX
                        Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                    End With
            End Select
            lblSize(i).ZOrder
        Next
        
        Call SetSelFlag
        If GetSelNum > 1 Then Call SetSelFlag(intBeginIdx)
    End If
End Sub

Private Sub SelItem(idx As Integer, blnSel As Boolean)
'功能:将指定索引元素选中/不选中
'参数:idx=控件索引,blnSel=是否选中
    Dim i As Integer
    Dim intBeginIdx As Integer
    Dim ObjSel As Control, tmpID As RelatID
    Dim blnUse(7) As Boolean '控制各个尺寸标志是否显示
    Dim lngTmp As Long
    Dim lngtmp1 As Long
    
    picPaper.SetFocus
    
    For i = 0 To UBound(blnUse)
        blnUse(i) = True
    Next
    
    Select Case objReport.Items("_" & idx).类型
        Case 2, 3 '文本标签及数据标签用同名控件
            Set ObjSel = lbl(idx)
        Case 10
            Set ObjSel = Shp(idx)
        Case 1 '线条
            With lblLine(idx)
                If .Width > .Height Then
                    '横向
                    For i = 0 To UBound(blnUse)
                        If Not (i = 2 Or i = 6) Then blnUse(i) = False
                    Next
                Else
                    '纵向
                    For i = 0 To UBound(blnUse)
                        If Not (i = 0 Or i = 4) Then blnUse(i) = False
                    Next
                End If
            End With
            Set ObjSel = lblLine(idx)
        Case 4, 5 '汇总表及任意表用同名控件
            Set ObjSel = msh(idx)
            '分栏控件置前(如果有)
            For Each tmpID In objReport.Items("_" & idx).CopyIDs
                msh(tmpID.id).ZOrder
            Next
        Case 11
            Set ObjSel = img(idx)
        Case 12 '@@@
            Set ObjSel = Chart(idx)
        Case 13
            Set ObjSel = ImgCode(idx)
        Case 14
            Set ObjSel = pic(idx)
    End Select

    If blnSel Then
        '如果已经选中,则不再重复处理
        If Mid(ObjSel.Tag, 1, 2) = "S_" Then Exit Sub
        
        Set objLastSel = ObjSel: intCurID = idx
        
        
        '重新计算实际位置及相关坐标2002-03-26
        If ObjSel.Container.name <> "picPaper" Then
            lngTmp = ObjSel.Container.Top
            lngtmp1 = ObjSel.Container.Left
        End If

        
        intBeginIdx = lblSize.UBound + 1 '索引为(1n-8n),n>0
        
         '元素控件的Tag记录选择标志的起始索引
        ObjSel.Tag = "S_" & intBeginIdx
        ObjSel.ZOrder IIF(objReport.Items("_" & ObjSel.Index).类型 <> 10, 0, 1)
        If objReport.Items("_" & ObjSel.Index).类型 = 10 Then lblshp(ObjSel.Index).ZOrder 1

        With WinProperty
            .H = lblSize(0).Height / Screen.TwipsPerPixelX
            .W = lblSize(0).Width / Screen.TwipsPerPixelX
        End With
        
        For i = intBeginIdx To intBeginIdx + 7 '选择标志从"上中"开始,"顺时针"处理
            Load lblSize(i)
            Select Case IIF(i Mod 8 <> 0, i Mod 8, 8) '定位选择边框的位置
                'zyb#Modify
                '改用MoveWindow()解决
                
                Case 1 '上中
                    lblSize(i).Tag = idx '第一个(上中)记录对应控件的索引
                    With WinProperty
                        .T = (ObjSel.Top + lngTmp - lblSize(i).Height) / Screen.TwipsPerPixelX
                        .l = (ObjSel.Left + lngtmp1 + (ObjSel.Width - lblSize(i).Width) / 2) / Screen.TwipsPerPixelX
                        Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                    End With
                    lblSize(i).MousePointer = 7
                Case 2 '上右
                    lblSize(i).Tag = ObjSel.name '第二个(右上)记录对应控件的名称
                    With WinProperty
                        .T = (ObjSel.Top + lngTmp - lblSize(i).Height) / Screen.TwipsPerPixelX
                        .l = (ObjSel.Left + lngtmp1 + ObjSel.Width) / Screen.TwipsPerPixelX
                        Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                    End With
                    lblSize(i).MousePointer = 6
                Case 3 '右中
                    With WinProperty
                        .T = (ObjSel.Top + lngTmp + (ObjSel.Height - lblSize(i).Height) / 2) / Screen.TwipsPerPixelX
                        .l = (ObjSel.Left + lngtmp1 + ObjSel.Width) / Screen.TwipsPerPixelX
                        Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                    End With
                    lblSize(i).MousePointer = 9
                Case 4 '右下
                    With WinProperty
                        .T = (ObjSel.Top + lngTmp + ObjSel.Height) / Screen.TwipsPerPixelX
                        .l = (ObjSel.Left + lngtmp1 + ObjSel.Width) / Screen.TwipsPerPixelX
                        Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                    End With
                    lblSize(i).MousePointer = 8
                Case 5 '下中
                    With WinProperty
                        .T = (ObjSel.Top + lngTmp + ObjSel.Height) / Screen.TwipsPerPixelX
                        .l = (ObjSel.Left + lngtmp1 + (ObjSel.Width - lblSize(i).Width) / 2) / Screen.TwipsPerPixelX
                        Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                    End With
                    lblSize(i).MousePointer = 7
                Case 6 '左下
                    With WinProperty
                        .T = (ObjSel.Top + lngTmp + ObjSel.Height) / Screen.TwipsPerPixelX
                        .l = (ObjSel.Left + lngtmp1 - lblSize(i).Width) / Screen.TwipsPerPixelX
                        Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                    End With
                    lblSize(i).MousePointer = 6
                Case 7 '左中
                    With WinProperty
                        .T = (ObjSel.Top + lngTmp + (ObjSel.Height - lblSize(i).Height) / 2) / Screen.TwipsPerPixelX
                        .l = (ObjSel.Left + lngtmp1 - lblSize(i).Width) / Screen.TwipsPerPixelX
                        Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                    End With
                    lblSize(i).MousePointer = 9
                Case 8 '左上
                    With WinProperty
                        .T = (ObjSel.Top + lngTmp - lblSize(i).Height) / Screen.TwipsPerPixelX
                        .l = (ObjSel.Left + lngtmp1 - lblSize(i).Width) / Screen.TwipsPerPixelX
                        Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                    End With
                    lblSize(i).MousePointer = 8
            End Select
            lblSize(i).ZOrder
            lblSize(i).Visible = blnUse(IIF(i Mod 8 <> 0, i Mod 8, 8) - 1)
        Next
        
        Call SetSelFlag
        If GetSelNum > 1 Then Call SetSelFlag(intBeginIdx)
    Else
        If Trim(ObjSel.Tag) = "" Then Exit Sub '已处于非选择状态
        
        If objReport.Items("_" & ObjSel.Index).类型 = 4 Or objReport.Items("_" & ObjSel.Index).类型 = 5 Then
            selCell.Col1 = -1: selCell.Col2 = -1: selCell.Row1 = -1: selCell.Row2 = -1: selCell.Row = -1
            Call ShowPaperInfo
            intCurCol = -1
        End If
        
        ObjSel.ZOrder IIF(objReport.Items("_" & ObjSel.Index).类型 <> 10, 0, 1)
        intBeginIdx = CInt(Mid(ObjSel.Tag, 3))
        ObjSel.Tag = ""
        For i = intBeginIdx To intBeginIdx + 7
            Unload lblSize(i)
        Next
        
        Call SetSelFlag
        
        If GetSelNum > 0 Then
            If GetSelNum > 1 Then Call SetSelFlag(lblSize.UBound - 7)
            Select Case objReport.Items("_" & CInt(lblSize(lblSize.UBound - 7).Tag)).类型
                Case 2, 3
                    Set objLastSel = lbl(CInt(lblSize(lblSize.UBound - 7).Tag))
                Case 10
                    Set objLastSel = Shp(CInt(lblSize(lblSize.UBound - 7).Tag))
                Case 1
                    Set objLastSel = lblLine(CInt(lblSize(lblSize.UBound - 7).Tag))
                Case 4, 5
                    Set objLastSel = msh(CInt(lblSize(lblSize.UBound - 7).Tag))
                Case 11
                    Set objLastSel = img(CInt(lblSize(lblSize.UBound - 7).Tag))
                Case 12 '@@@
                    Set objLastSel = Chart(CInt(lblSize(lblSize.UBound - 7).Tag))
                Case 13
                    Set objLastSel = ImgCode(CInt(lblSize(lblSize.UBound - 7).Tag))
                Case 14
                    Set objLastSel = pic(CInt(lblSize(lblSize.UBound - 7).Tag))
            End Select
            intCurID = CInt(lblSize(lblSize.UBound - 7).Tag)
        Else
            Set objLastSel = Nothing: intCurID = 0
        End If
    End If
End Sub

Private Sub SetSelFlag(Optional intBegin As Integer = 0)
'功能:设定最后一个选中的项目控制标志与前面不同,或恢复所有标志为正常色
'参数：intBegin=(指定控件元素)尺寸标志起始索引,为0表示恢复所有标志为正常色
    Dim tmpObj As PictureBox
    Dim i As Integer
    If intBegin = 0 Then
        For Each tmpObj In lblSize
            If tmpObj.Index <> 0 Then
                If blnLock Then
                    tmpObj.BackColor = &HFF
                Else
                    tmpObj.BackColor = &HFF0000
                End If
            End If
        Next
    Else
        For i = intBegin To intBegin + 7
            lblSize(i).BackColor = &HC000&
        Next
    End If
End Sub

Private Function GetSelNum() As Integer
'功能：返回当前选择元素控件个数
'说明：利用选择控件一定存在尺寸标志
    Dim tmpObj As PictureBox, i As Integer
    For Each tmpObj In lblSize
        If tmpObj.Index <> 0 And tmpObj.Index Mod 8 = 0 Then i = i + 1
    Next
    GetSelNum = i
End Function

Private Function SelAreaItem(Area As RECT) As Integer
'功能：选择在指点定区域内的报表元素控件(不清除其它已选择的)
'返回：选中的个数
    Dim tmpItem As RPTItem, ObjSel As Object
    Dim lngLeft As Long, lngTop As Long
    
    For Each tmpItem In objReport.Items '@@@
        If InStr(1, "|1,|2,|3,|4,|5,|10,|11,|12,|13,|14,", "|" & tmpItem.类型) <> 0 _
            And Mid(cboFormat.ComboItems("_" & mbytCurrFmt).Key, 2) = tmpItem.格式号 Then
            Set ObjSel = GetInxObj(tmpItem.id)
            
            If tmpItem.类型 <> 14 And tmpItem.父ID <> 0 Then
                lngLeft = pic(tmpItem.父ID).Left
                lngTop = pic(tmpItem.父ID).Top
            Else
                lngLeft = 0
                lngTop = 0
            End If
            If Not (ObjSel.Top + lngTop > Area.Bottom Or _
                ObjSel.Left + lngLeft > Area.Right Or _
                ObjSel.Top + lngTop + ObjSel.Height < Area.Top Or _
                ObjSel.Left + lngLeft + ObjSel.Width < Area.Left) Then
                Call SelItem(ObjSel.Index, True)
                SelAreaItem = SelAreaItem + 1
            End If
        End If
    Next
End Function

Private Sub SetLock(blnLock As Boolean)
'功能:根据控件是否锁定设定控件选定边框颜色
    Dim tmpObj As PictureBox

    For Each tmpObj In lblSize
        If tmpObj.Index <> 0 Then
            If tmpObj.BackColor <> &HC000& Then
                If blnLock Then
                    tmpObj.BackColor = &HFF
                Else
                    tmpObj.BackColor = &HFF0000
                End If
            End If
        End If
    Next
End Sub

Private Sub ShowAttrib(Optional idx As Integer, Optional blnBase As Boolean = False)
'功能：显示指定索引控件对应的对象属性或清除属性
'参数：idx=控件索引,为0时表示清除属性框
    Dim intType As Integer, RowH As Single, i As Integer
    Dim arrHead As Variant, arrRowH As Variant
    Dim tmpID As RelatID, tmpItem As RPTItem
    Dim objFmt As RPTFmt
    Dim lngRow As Long, lngCol As Long
    Dim tmpObj As PictureBox
    Dim lngType As Long
    
    Set objFmt = objReport.Fmts("_" & mbytCurrFmt)
    
    With mshAtt
        lngRow = .Row: lngCol = .Col
        
        .Redraw = False
        .Clear
        .Rows = 2
        .TextMatrix(0, 0) = "项目": .TextMatrix(0, 1) = "设置"
        .Row = 0: .Col = 0: .CellAlignment = 4: .Col = 1: .CellAlignment = 4
        .ColAlignment(0) = 1: .ColAlignment(1) = 1
        .ExtendLastCol = True
        .ColWidth(0) = 1200
        '.ColWidth(1) = 1200
        lblNote.Caption = ""
        If idx = 0 Then
            '如果被选中的所有元素都是相同的元素
            For Each tmpObj In lblSize
                If tmpObj.Index Mod 8 = 1 Then
                    If lngType <> 0 And lngType <> objReport.Items("_" & tmpObj.Tag).类型 Then lngType = 0: idx = 0: Exit For
                    lngType = objReport.Items("_" & tmpObj.Tag).类型
                    idx = objReport.Items("_" & tmpObj.Tag).id
                End If
            Next
        End If
        If idx <> 0 Then
            '如果是标签，且字段为图型，则当作图型处理
            intType = objReport.Items("_" & idx).类型
            If ItemIsGraph(idx) Then intType = 11
            Select Case intType
                Case 1, 10 '线条,框线
                    .Rows = 8
                    .TextMatrix(1, 0) = "类型": .TextMatrix(1, 1) = IIF(objReport.Items("_" & idx).类型 = 1, "线条", "框线")
                    .TextMatrix(2, 0) = "名称": .TextMatrix(2, 1) = objReport.Items("_" & idx).名称
                    .TextMatrix(3, 0) = "X坐标": .TextMatrix(3, 1) = Format(objReport.Items("_" & idx).X / Twip_mm, "0.00")
                    .TextMatrix(4, 0) = "Y坐标": .TextMatrix(4, 1) = Format(objReport.Items("_" & idx).Y / Twip_mm, "0.00")
                    If objReport.Items("_" & idx).类型 = 1 Then
                        If objReport.Items("_" & idx).W > objReport.Items("_" & idx).H Then .TextMatrix(5, 0) = "宽度": .TextMatrix(5, 1) = Format(objReport.Items("_" & idx).W / Twip_mm, "0.00")
                        If objReport.Items("_" & idx).W < objReport.Items("_" & idx).H Then .TextMatrix(5, 0) = "高度": .TextMatrix(5, 1) = Format(objReport.Items("_" & idx).H / Twip_mm, "0.00")
                        .TextMatrix(6, 0) = "前景色": .TextMatrix(6, 1) = "█": .Row = 6: .Col = 1: .CellForeColor = objReport.Items("_" & idx).前景
                    Else
                        .TextMatrix(5, 0) = "高度": .TextMatrix(5, 1) = Format(objReport.Items("_" & idx).H / Twip_mm, "0.00")
                        .TextMatrix(6, 0) = "宽度": .TextMatrix(6, 1) = Format(objReport.Items("_" & idx).W / Twip_mm, "0.00")
                    End If
                    .TextMatrix(7, 0) = "加粗": .TextMatrix(7, 1) = IIF(objReport.Items("_" & idx).粗体, "√", "×")
                    If intType = 10 Then
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, 0) = "形状": .TextMatrix(.Rows - 1, 1) = IIF(objReport.Items("_" & idx).边框 = 0, "方形", "圆形")
                    End If
                Case 2 '标签
                    Dim str性质 As String
                    Dim str关联报表 As String
                    .Rows = 22
                    .TextMatrix(1, 0) = "类型": .TextMatrix(1, 1) = "标签"
                    .TextMatrix(2, 0) = "名称": .TextMatrix(2, 1) = objReport.Items("_" & idx).名称
                    .TextMatrix(3, 0) = "内容": .TextMatrix(3, 1) = objReport.Items("_" & idx).内容
                    .TextMatrix(4, 0) = "对齐": .TextMatrix(4, 1) = IIF(objReport.Items("_" & idx).对齐 = 0, "左对齐", IIF(objReport.Items("_" & idx).对齐 = 1, "中对齐", "右对齐"))
                    .TextMatrix(5, 0) = "X坐标": .TextMatrix(5, 1) = Format(objReport.Items("_" & idx).X / Twip_mm, "0.00")
                    .TextMatrix(6, 0) = "Y坐标": .TextMatrix(6, 1) = Format(objReport.Items("_" & idx).Y / Twip_mm, "0.00")
                    .TextMatrix(7, 0) = "宽度": .TextMatrix(7, 1) = Format(objReport.Items("_" & idx).W / Twip_mm, "0.00")
                    .TextMatrix(8, 0) = "高度": .TextMatrix(8, 1) = Format(objReport.Items("_" & idx).H / Twip_mm, "0.00")
                    .TextMatrix(9, 0) = "自动调整大小": .TextMatrix(9, 1) = IIF(objReport.Items("_" & idx).自调, "√", "×")
                    .TextMatrix(10, 0) = "字体": .TextMatrix(10, 1) = objReport.Items("_" & idx).字体
                    .TextMatrix(11, 0) = "自动字体": .TextMatrix(11, 1) = IIF(objReport.Items("_" & idx).行高 = 1, "√", "×")
                    .TextMatrix(12, 0) = "边框": .TextMatrix(12, 1) = IIF(objReport.Items("_" & idx).边框, "√", "×")
                    .TextMatrix(13, 0) = "背景色": .TextMatrix(13, 1) = "█": .Row = 13: .Col = 1: .CellForeColor = objReport.Items("_" & idx).背景
                    .TextMatrix(14, 0) = "参照对象": .TextMatrix(14, 1) = objReport.Items("_" & idx).参照
                    str性质 = objReport.Items("_" & idx).性质
                    If str性质 = "0" Or str性质 = "" Then
                        str性质 = "独立"
                        .TextMatrix(15, 0) = "方向": .TextMatrix(15, 1) = "独立"
                    Else
                        str性质 = Mid(str性质, 2)
                        str性质 = IIF(str性质 = "1", "靠左", IIF(str性质 = "2", "靠中", "靠右"))
                        .TextMatrix(15, 0) = "方向": .TextMatrix(15, 1) = IIF(Mid(objReport.Items("_" & idx).性质, 1, 1) = "1", "表上项", "表下项")
                    End If
                    .TextMatrix(16, 0) = "性质": .TextMatrix(16, 1) = str性质
                    .TextMatrix(17, 0) = "格式": .TextMatrix(17, 1) = objReport.Items("_" & idx).格式
                    .TextMatrix(18, 0) = "行距": .TextMatrix(18, 1) = objReport.Items("_" & idx).网格
                    .TextMatrix(19, 0) = "数据源行号": .TextMatrix(19, 1) = objReport.Items("_" & idx).源行号
                    For i = 1 To objReport.Items("_" & idx).Relations.count
                        If InStr(str关联报表, "," & objReport.Items("_" & idx).Relations.Item(i).关联报表名称) = 0 Then
                            str关联报表 = str关联报表 & "," & objReport.Items("_" & idx).Relations.Item(i).关联报表名称
                        End If
                    Next
                    If str关联报表 <> "" Then
                        str关联报表 = Mid(str关联报表, 2)
                    End If
                    .TextMatrix(20, 0) = "关联报表": .TextMatrix(20, 1) = str关联报表
                    .TextMatrix(21, 0) = "水平反转": .TextMatrix(21, 1) = IIF(objReport.Items("_" & idx).水平反转, "√", "×")
                Case 4 '任意表格
                    .Rows = 18
                    '因为任意表格数据来源于多个数据源,所以内容不能确定
                    .TextMatrix(1, 0) = "类型": .TextMatrix(1, 1) = "任意表格"
                    .TextMatrix(2, 0) = "名称": .TextMatrix(2, 1) = objReport.Items("_" & idx).名称
                    .TextMatrix(3, 0) = "X坐标": .TextMatrix(3, 1) = Format(objReport.Items("_" & idx).X / Twip_mm, "0.00")
                    .TextMatrix(4, 0) = "Y坐标": .TextMatrix(4, 1) = Format(objReport.Items("_" & idx).Y / Twip_mm, "0.00")
                    .TextMatrix(5, 0) = "宽度": .TextMatrix(5, 1) = Format(objReport.Items("_" & idx).W / Twip_mm, "0.00")
                    .TextMatrix(6, 0) = "高度": .TextMatrix(6, 1) = Format(objReport.Items("_" & idx).H / Twip_mm, "0.00")
                    If msh(idx).Row < msh(idx).FixedRows Then
                        RowH = 0
                        arrHead = Split(objReport.Items("_" & objReport.Items("_" & idx).SubIDs(1).Key).表头, "|")
                        If selCell.Row <> -1 Then
                            For i = selCell.Row1 To selCell.Row2
                                arrRowH = Split(arrHead(i), "^")
                                On Error Resume Next
                                Err = 0
                                arrRowH(1) = arrRowH(1) + 0
                                If Err <> 0 Then
                                    RowH = RowH + 255
                                Else
                                    RowH = RowH + arrRowH(1)
                                End If
                            Next
                        Else
                            RowH = 255
                        End If
                        On Error GoTo 0
                        .TextMatrix(7, 0) = "行高": .TextMatrix(7, 1) = Format(RowH / Twip_mm, "0.00")
                    Else
                        .TextMatrix(7, 0) = "行高": .TextMatrix(7, 1) = Format(objReport.Items("_" & idx).行高 / Twip_mm, "0.00")
                    End If
                    .TextMatrix(8, 0) = "字体": .TextMatrix(8, 1) = objReport.Items("_" & idx).字体
                    .TextMatrix(9, 0) = "背景色": .TextMatrix(9, 1) = "█": .Row = 9: .Col = 1: .CellForeColor = objReport.Items("_" & idx).背景
                    .TextMatrix(10, 0) = "表体网格色": .TextMatrix(10, 1) = "█": .Row = 10: .Col = 1: .CellForeColor = objReport.Items("_" & idx).网格
                    .TextMatrix(11, 0) = "表头网格色": .TextMatrix(11, 1) = "█": .Row = 11: .Col = 1: .CellForeColor = IIF(objReport.Items("_" & idx).格式 = "", objReport.Items("_" & idx).网格, Val(objReport.Items("_" & idx).格式))
                    .TextMatrix(12, 0) = "参照对象": .TextMatrix(12, 1) = objReport.Items("_" & idx).参照
                    .TextMatrix(13, 0) = "性质": .TextMatrix(13, 1) = IIF(objReport.Items("_" & idx).参照 = "", "独立", IIF(objReport.Items("_" & idx).性质 = "1", "附加", "左联接"))
                    .TextMatrix(14, 0) = "分栏": .TextMatrix(14, 1) = IIF(objReport.Items("_" & idx).分栏 <= 1, 1, objReport.Items("_" & idx).分栏)
                    .TextMatrix(15, 0) = "边线": .TextMatrix(15, 1) = IIF(objReport.Items("_" & idx).边框, "√", "×")
                    .TextMatrix(16, 0) = "换行": .TextMatrix(16, 1) = IIF(objReport.Items("_" & idx).自调, "√", "×")
                    .TextMatrix(17, 0) = "表格线加粗": .TextMatrix(17, 1) = IIF(objReport.Items("_" & idx).表格线加粗, "√", "×")
                Case 5 '汇总表格
                    .Rows = 17
                    '汇总表格不能分栏
                    .TextMatrix(1, 0) = "类型": .TextMatrix(1, 1) = "汇总表格"
                    .TextMatrix(2, 0) = "名称": .TextMatrix(2, 1) = objReport.Items("_" & idx).名称
                    .TextMatrix(3, 0) = "内容": .TextMatrix(3, 1) = objReport.Items("_" & idx).内容
                    .TextMatrix(4, 0) = "X坐标": .TextMatrix(4, 1) = Format(objReport.Items("_" & idx).X / Twip_mm, "0.00")
                    .TextMatrix(5, 0) = "Y坐标": .TextMatrix(5, 1) = Format(objReport.Items("_" & idx).Y / Twip_mm, "0.00")
                    .TextMatrix(6, 0) = "宽度": .TextMatrix(6, 1) = Format(objReport.Items("_" & idx).W / Twip_mm, "0.00")
                    .TextMatrix(7, 0) = "高度": .TextMatrix(7, 1) = Format(objReport.Items("_" & idx).H / Twip_mm, "0.00")
                    .TextMatrix(8, 0) = "行高": .TextMatrix(8, 1) = Format(objReport.Items("_" & idx).行高 / Twip_mm, "0.00")
                    .TextMatrix(9, 0) = "字体": .TextMatrix(9, 1) = objReport.Items("_" & idx).字体
                    .TextMatrix(10, 0) = "背景色": .TextMatrix(10, 1) = "█": .Row = 10: .Col = 1: .CellForeColor = objReport.Items("_" & idx).背景
                    .TextMatrix(11, 0) = "网格色": .TextMatrix(11, 1) = "█": .Row = 11: .Col = 1: .CellForeColor = objReport.Items("_" & idx).网格
                    .TextMatrix(12, 0) = "参照对象": .TextMatrix(12, 1) = objReport.Items("_" & idx).参照
                    .TextMatrix(13, 0) = "性质": .TextMatrix(13, 1) = IIF(objReport.Items("_" & idx).参照 = "", "独立", IIF(objReport.Items("_" & idx).性质 = "1", "附加", "左联接"))
                    .TextMatrix(14, 0) = "边线": .TextMatrix(14, 1) = IIF(objReport.Items("_" & idx).边框, "√", "×")
                    .TextMatrix(15, 0) = "换行": .TextMatrix(15, 1) = IIF(objReport.Items("_" & idx).自调, "√", "×")
                    .TextMatrix(16, 0) = "表格线加粗": .TextMatrix(16, 1) = IIF(objReport.Items("_" & idx).表格线加粗, "√", "×")
                Case 11 '图片
                    If objReport.Items("_" & idx).类型 = 11 Then
                        .Rows = 12 '单独加入的图片元素
                    Else
                        .Rows = 11
                    End If
                    .TextMatrix(1, 0) = "类型": .TextMatrix(1, 1) = "图片"
                    .TextMatrix(2, 0) = "名称": .TextMatrix(2, 1) = objReport.Items("_" & idx).名称
                    If objReport.Items("_" & idx).类型 = 2 Then
                        .TextMatrix(3, 0) = "内容": .TextMatrix(3, 1) = objReport.Items("_" & idx).内容
                    Else
                        .TextMatrix(3, 0) = "内容": .TextMatrix(3, 1) = IIF(Not objReport.Items("_" & idx).图片 Is Nothing, "[Pictrue]", "")
                    End If
                    .TextMatrix(4, 0) = "X坐标": .TextMatrix(4, 1) = Format(objReport.Items("_" & idx).X / Twip_mm, "0.00")
                    .TextMatrix(5, 0) = "Y坐标": .TextMatrix(5, 1) = Format(objReport.Items("_" & idx).Y / Twip_mm, "0.00")
                    .TextMatrix(6, 0) = "宽度": .TextMatrix(6, 1) = Format(objReport.Items("_" & idx).W / Twip_mm, "0.00")
                    .TextMatrix(7, 0) = "高度": .TextMatrix(7, 1) = Format(objReport.Items("_" & idx).H / Twip_mm, "0.00")
                    .TextMatrix(8, 0) = "边框": .TextMatrix(8, 1) = IIF(objReport.Items("_" & idx).边框, "√", "×")
                    .TextMatrix(9, 0) = "保持比例": .TextMatrix(9, 1) = IIF(objReport.Items("_" & idx).粗体, "√", "×")
                    .TextMatrix(10, 0) = "自动调整大小": .TextMatrix(10, 1) = IIF(objReport.Items("_" & idx).自调, "√", "×")
                    If objReport.Items("_" & idx).类型 = 11 Then
                        .TextMatrix(11, 0) = "报告图像": .TextMatrix(11, 1) = IIF(objReport.Items("_" & idx).下线, "√", "×")
                    End If
                Case 14 '卡片
                    .Rows = 13
                    .TextMatrix(1, 0) = "类型": .TextMatrix(1, 1) = "卡片"
                    .TextMatrix(2, 0) = "名称": .TextMatrix(2, 1) = objReport.Items("_" & idx).名称
                    .TextMatrix(3, 0) = "X坐标": .TextMatrix(3, 1) = Format(objReport.Items("_" & idx).X / Twip_mm, "0.00")
                    .TextMatrix(4, 0) = "Y坐标": .TextMatrix(4, 1) = Format(objReport.Items("_" & idx).Y / Twip_mm, "0.00")
                    .TextMatrix(5, 0) = "宽度": .TextMatrix(5, 1) = Format(objReport.Items("_" & idx).W / Twip_mm, "0.00")
                    .TextMatrix(6, 0) = "高度": .TextMatrix(6, 1) = Format(objReport.Items("_" & idx).H / Twip_mm, "0.00")
                    .TextMatrix(7, 0) = "边框": .TextMatrix(7, 1) = IIF(objReport.Items("_" & idx).边框, "√", "×")
                    .TextMatrix(8, 0) = "数据源": .TextMatrix(8, 1) = objReport.Items("_" & idx).数据源
                    .TextMatrix(9, 0) = "上下间距": .TextMatrix(9, 1) = Format(objReport.Items("_" & idx).上下间距 / Twip_mm, "0.00")
                    .TextMatrix(10, 0) = "左右间距": .TextMatrix(10, 1) = Format(objReport.Items("_" & idx).左右间距 / Twip_mm, "0.00")
                    .TextMatrix(11, 0) = "横向分栏": .TextMatrix(11, 1) = objReport.Items("_" & idx).横向分栏
                    .TextMatrix(12, 0) = "纵向分栏": .TextMatrix(12, 1) = objReport.Items("_" & idx).纵向分栏
                Case 12 '图表@@@
                    .Rows = 11
                    .TextMatrix(1, 0) = "类型": .TextMatrix(1, 1) = "图表"
                    .TextMatrix(2, 0) = "名称": .TextMatrix(2, 1) = objReport.Items("_" & idx).名称
                    .TextMatrix(3, 0) = "设置": .TextMatrix(3, 1) = ""
                    .TextMatrix(4, 0) = "X坐标": .TextMatrix(4, 1) = Format(objReport.Items("_" & idx).X / Twip_mm, "0.00")
                    .TextMatrix(5, 0) = "Y坐标": .TextMatrix(5, 1) = Format(objReport.Items("_" & idx).Y / Twip_mm, "0.00")
                    .TextMatrix(6, 0) = "宽度": .TextMatrix(6, 1) = Format(objReport.Items("_" & idx).W / Twip_mm, "0.00")
                    .TextMatrix(7, 0) = "高度": .TextMatrix(7, 1) = Format(objReport.Items("_" & idx).H / Twip_mm, "0.00")
                    .TextMatrix(8, 0) = "字体": .TextMatrix(8, 1) = objReport.Items("_" & idx).字体
                    .TextMatrix(9, 0) = "前景色": .TextMatrix(9, 1) = "█": .Row = 9: .Col = 1: .CellForeColor = objReport.Items("_" & idx).前景
                    .TextMatrix(10, 0) = "背景色": .TextMatrix(10, 1) = "█": .Row = 10: .Col = 1: .CellForeColor = objReport.Items("_" & idx).背景
                Case 13 '条码
                    If objReport.Items("_" & idx).序号 = 1 Then 'Code 128(遗留)
                        .Rows = 12
                    ElseIf objReport.Items("_" & idx).序号 = 2 Then 'Code 39
                        .Rows = 13
                    ElseIf objReport.Items("_" & idx).序号 = 3 Then 'Code 128 Auto
                        .Rows = 13
                    ElseIf objReport.Items("_" & idx).序号 = 10 Then 'QR Code
                        .Rows = 11
                    End If
                    
                    .TextMatrix(1, 0) = "类型": .TextMatrix(1, 1) = "条码"
                    .TextMatrix(2, 0) = "名称": .TextMatrix(2, 1) = objReport.Items("_" & idx).名称
                    .TextMatrix(3, 0) = "内容": .TextMatrix(3, 1) = objReport.Items("_" & idx).内容
                    .TextMatrix(4, 0) = "X坐标": .TextMatrix(4, 1) = Format(objReport.Items("_" & idx).X / Twip_mm, "0.00")
                    .TextMatrix(5, 0) = "Y坐标": .TextMatrix(5, 1) = Format(objReport.Items("_" & idx).Y / Twip_mm, "0.00")
                    .TextMatrix(6, 0) = "宽度": .TextMatrix(6, 1) = Format(objReport.Items("_" & idx).W / Twip_mm, "0.00")
                    .TextMatrix(7, 0) = "高度": .TextMatrix(7, 1) = Format(objReport.Items("_" & idx).H / Twip_mm, "0.00")
                    .TextMatrix(8, 0) = "条码类型": .TextMatrix(8, 1) = Decode(objReport.Items("_" & idx).序号, 1, "Code 128(遗留)", 3, "Code 128 Auto", 2, "Code 39", 10, "QR Code", "")
                    .TextMatrix(9, 0) = "数据源行号": .TextMatrix(9, 1) = objReport.Items("_" & idx).源行号
                    If objReport.Items("_" & idx).序号 = 1 Then
                        .TextMatrix(10, 0) = "显示数字": .TextMatrix(10, 1) = IIF(Mid(objReport.Items("_" & idx).表头, 1, 1) = "1", "√", "×")
                        .TextMatrix(11, 0) = "旋转方向": .TextMatrix(11, 1) = Decode(Val(Mid(objReport.Items("_" & idx).表头, 3, 1)), 0, "不旋转", 1, "顺时针90度", 2, "逆时针90度", "")
                    ElseIf objReport.Items("_" & idx).序号 = 2 Then
                        .TextMatrix(10, 0) = "显示数字": .TextMatrix(10, 1) = IIF(Mid(objReport.Items("_" & idx).表头, 1, 1) = "1", "√", "×")
                        .TextMatrix(11, 0) = "求校验和": .TextMatrix(11, 1) = IIF(Mid(objReport.Items("_" & idx).表头, 2, 1) = "1", "√", "×")
                        .TextMatrix(12, 0) = "旋转方向": .TextMatrix(12, 1) = Decode(Val(Mid(objReport.Items("_" & idx).表头, 3, 1)), 0, "不旋转", 1, "顺时针90度", 2, "逆时针90度", "")
                    ElseIf objReport.Items("_" & idx).序号 = 3 Then
                        .TextMatrix(10, 0) = "条码线宽": .TextMatrix(10, 1) = objReport.Items("_" & idx).行高
                        .TextMatrix(11, 0) = "显示数字": .TextMatrix(11, 1) = IIF(Mid(objReport.Items("_" & idx).表头, 1, 1) = "1", "√", "×")
                        .TextMatrix(12, 0) = "旋转方向": .TextMatrix(12, 1) = Decode(Val(Mid(objReport.Items("_" & idx).表头, 3, 1)), 0, "不旋转", 1, "顺时针90度", 2, "逆时针90度", "")
                    ElseIf objReport.Items("_" & idx).序号 = 10 Then
                        .TextMatrix(10, 0) = "自动调整大小": .TextMatrix(10, 1) = IIF(objReport.Items("_" & idx).自调, "√", "×")
                    End If
            End Select
            If intType <> 12 And intType <> 14 And intType <> 5 Then
                '设置其是否放置于卡片中
                 .AddItem ""
                 .TextMatrix(.Rows - 1, 0) = "容器"
                 If objReport.Items("_" & idx).父ID = 0 Then
                    .TextMatrix(.Rows - 1, 1) = "页面"
                 Else
                    .TextMatrix(.Rows - 1, 1) = objReport.Items("_" & objReport.Items("_" & idx).父ID).名称
                 End If
            End If
            '如果是多选，则只显示常用项目
            If lngType <> 0 Then
                For i = 1 To .Rows - 1
                    If InStr("前景色;背景色;对齐;自动调整大小;字体;自动字体;边框", .TextMatrix(i, 0)) = 0 Then .RowHidden(i) = True
                Next
            End If
        End If
        
        '显示纸张等基本属性
        If blnBase Then
            '横向打印不支持动态纸张
            .Rows = IIF(objFmt.纸向 = 1, 14, 13)
            .TextMatrix(1, 0) = "报表元素": .TextMatrix(1, 1) = ""
            .TextMatrix(2, 0) = "输出图形": .TextMatrix(2, 1) = GetCurOutChart
            .TextMatrix(3, 0) = "票据": .TextMatrix(3, 1) = IIF(objReport.票据, "√", "×")
            .TextMatrix(4, 0) = "空表打印": .TextMatrix(4, 1) = IIF(objReport.打印方式 = 0, "√", "×")
            .TextMatrix(5, 0) = "打印机": .TextMatrix(5, 1) = objReport.打印机
            .TextMatrix(6, 0) = "纸张": .TextMatrix(6, 1) = GetPaperName(objFmt.纸张, objFmt.W, objFmt.H)
            .TextMatrix(7, 0) = "纸向": .TextMatrix(7, 1) = IIF(objFmt.纸向 = 1, "纵向", "横向")
            .TextMatrix(8, 0) = "高度": .TextMatrix(8, 1) = CLng(objFmt.H / Twip_mm) & "毫米"
            .TextMatrix(9, 0) = "宽度": .TextMatrix(9, 1) = CLng(objFmt.W / Twip_mm) & "毫米"
            .TextMatrix(10, 0) = "进纸方式": .TextMatrix(10, 1) = CboTest.Text
            .TextMatrix(11, 0) = "禁止开始时间": .TextMatrix(11, 1) = Format(objReport.禁止开始时间, "HH:mm:ss")
            .TextMatrix(12, 0) = "禁止结束时间": .TextMatrix(12, 1) = Format(objReport.禁止结束时间, "HH:mm:ss")
            If objFmt.纸向 = 1 Then
                .TextMatrix(13, 0) = "动态纸张": .TextMatrix(13, 1) = IIF(objFmt.动态纸张, "√", "×")
            End If
        End If
        If lngRow <= .Rows - 1 Then
            .Row = IIF(lngRow <= .Rows - 1, lngRow, 1)
        Else
            .Row = 1
        End If
         .Col = 1
         mshAtt_AfterRowColChange 0, 0, mshAtt.Row, mshAtt.Col
        .Redraw = True
    End With
End Sub

Private Sub NoneEdit()
    txtAtt.Text = "": txtAtt.Visible = False
    cmdAtt.Visible = False
    cboAtt.Clear: cboAtt.Visible = False
    cboText.Clear: cboText.Visible = False
    dtpAtt.Visible = False
End Sub

Private Function SelIndex() As Integer
'功能：当只有一个元素控件被选择时,返回其控件索引
    Dim tmpObj As PictureBox
    
    If lblSize.count > 9 Or lblSize.count = 1 Then SelIndex = 0: Exit Function
    For Each tmpObj In lblSize
        If tmpObj.Index Mod 8 = 1 Then SelIndex = CInt(tmpObj.Tag): Exit Function
    Next
End Function

Private Sub MoveSelect(lngX As Long, lngY As Long, Optional ByVal blnReSize As Boolean)
'功能:移动选中元素控件
'参数:lngX=X偏移量,lngY=Y偏移量,blnReSize=改变元素大小时调用
    Dim ObjSel As Control, tmpObj As PictureBox
    Dim tmpID As RelatID, objParent As VSFlexGrid
    Dim ItemThis As RPTItem, blnMove As Boolean
    Dim tmpObj1 As PictureBox
    Dim tmpItem As RPTItem, blntmp As Boolean
    Dim vPoint As PointAPI
    
    '为提高速度,改用MoveWindow函数
    If lngX = 0 And lngY = 0 And blnReSize = False Then Exit Sub
    If blnLock Then Exit Sub '锁定
    For Each tmpObj In lblSize
        If tmpObj.Index Mod 8 = 1 Then
            blnMove = Not objReport.Items("_" & tmpObj.Tag).系统
            If blnMove Then
                '如果是移动卡片，则卡片内部的元素不再移动
                If objReport.Items("_" & tmpObj.Tag).父ID <> 0 Then
                    For Each tmpObj1 In lblSize
                        If tmpObj1.Index Mod 8 = 1 Then
                            If objReport.Items("_" & tmpObj.Tag).父ID = objReport.Items("_" & tmpObj1.Tag).id Then
                                GoTo NextObj
                            End If
                        End If
                    Next
                End If
                Select Case objReport.Items("_" & tmpObj.Tag).类型
                    Case 1
                        Set ObjSel = lblLine(tmpObj.Tag)
                    Case 2, 3, 10, 12 '@@@
                        If objReport.Items("_" & tmpObj.Tag).类型 = 12 Then
                            Set ObjSel = Chart(tmpObj.Tag)
                            If ObjSel.Top + lngY < 0 Then lngY = 0
                            If ObjSel.Left + lngX < 0 Then lngX = 0
                        ElseIf objReport.Items("_" & tmpObj.Tag).类型 = 10 Then
                            Set ObjSel = Shp(tmpObj.Tag)
                        Else
                            Set ObjSel = lbl(tmpObj.Tag)
                        End If
                        
                        blnMove = (objReport.Items("_" & tmpObj.Tag).参照 = "")
                        
                        If Not blnMove Then
                            lngX = 0
                            For Each ItemThis In objReport.Items
                                If ItemThis.格式号 = mbytCurrFmt _
                                    And ItemThis.名称 = objReport.Items("_" & tmpObj.Tag).参照 _
                                    And InStr(1, "4,5", ItemThis.类型) <> 0 Then
                                    On Error Resume Next
                                    Set objParent = msh(ItemThis.Key)
                                    Exit For
                                End If
                            Next
                            
                            If Mid(objReport.Items("_" & tmpObj.Tag).性质, 1, 1) = 1 Then   '表上项
                                If ObjSel.Top + ObjSel.Height + lngY > objParent.Top - 100 Then lngY = 0
                            Else
                                If ObjSel.Top + lngY < objParent.Top + objParent.Height + 100 Then lngY = 0
                            End If
                            blnMove = True
                        End If
                    Case 11
                        Set ObjSel = img(tmpObj.Tag)
                    Case 14
                        Set ObjSel = pic(tmpObj.Tag)
                    Case 13
                        Set ObjSel = ImgCode(tmpObj.Tag)
                    Case 4, 5
                        '如果是附加表格或左联接表格,则退出
                        blnMove = (objReport.Items("_" & tmpObj.Tag).参照 = "")
                        
                        Set ObjSel = msh(tmpObj.Tag)
                        
                        '调整分栏
                        For Each tmpID In objReport.Items("_" & tmpObj.Tag).CopyIDs
                            msh(tmpID.id).Top = msh(tmpID.id).Top + lngY
                            msh(tmpID.id).Left = msh(tmpID.id).Left + lngX
                        Next
                        Call LinkMove(ObjSel.Index, lngX, lngY)
                End Select
            End If
            
            If blnMove Then
                '单个元素才判断是否移入
                ObjSel.Top = ObjSel.Top + lngY
                ObjSel.Left = ObjSel.Left + lngX
                If objReport.Items("_" & tmpObj.Tag).类型 = 10 Then
                    '框线同步移动lblshp
                    lblshp(tmpObj.Tag).Top = ObjSel.Top
                    lblshp(tmpObj.Tag).Left = ObjSel.Left
                End If
                blntmp = False
                If objReport.Items("_" & tmpObj.Tag).类型 <> 14 And lblSize.count = 9 And objReport.Items("_" & tmpObj.Tag).类型 <> 12 Then
                    If UCase(ObjSel.Container.name) = "PIC" Then
                        '如果是子元素，则判断是否移动到了卡片外面,或其他卡片
                        If ObjSel.Top > ObjSel.Container.Height Or ObjSel.Left > ObjSel.Container.Width Or ObjSel.Top < -1 * ObjSel.Height Or ObjSel.Left < -1 * ObjSel.Width Then
                            '移出了卡片
                            For Each tmpItem In objReport.Items
                                If tmpItem.类型 = 14 And tmpItem.id <> ObjSel.Container.Index And tmpItem.格式号 = mbytCurrFmt Then
                                    If ObjSel.Top + ObjSel.Container.Top >= tmpItem.Y And ObjSel.Left + ObjSel.Container.Left >= tmpItem.X And _
                                        ObjSel.Height + ObjSel.Top + ObjSel.Container.Top <= tmpItem.Y + tmpItem.H And ObjSel.Width + ObjSel.Left + ObjSel.Container.Left <= tmpItem.X + tmpItem.W Then
                                        blntmp = True
                                        mlngY = ObjSel.Top + ObjSel.Container.Top - tmpItem.Y
                                        mlngX = ObjSel.Left + ObjSel.Container.Left - tmpItem.X
                                        Set mobjMove = pic(tmpItem.id)
                                        Exit For
                                    Else
                                        mlngY = ObjSel.Top
                                        mlngX = ObjSel.Left
                                        Set mobjMove = picPaper
                                    End If
                                End If
                            Next
                            If blntmp = False Then
                                '没有移入其他卡片就放入纸张中
                                If objReport.Items("_" & tmpObj.Tag).类型 = 4 Then
                                    ObjSel.Top = ObjSel.Top + ObjSel.Container.Top
                                    ObjSel.Left = ObjSel.Left + ObjSel.Container.Left
                                    Set ObjSel.Container = picPaper
                                    objReport.Items("_" & ObjSel.Index).父ID = 0
                                    objReport.Items("_" & ObjSel.Index).X = ObjSel.Left + ObjSel.Container.Left
                                    objReport.Items("_" & ObjSel.Index).Y = ObjSel.Top + ObjSel.Container.Top
                                    mlngY = ObjSel.Top + ObjSel.Container.Top
                                    mlngX = ObjSel.Left + ObjSel.Container.Left
                                    Set mobjMove = picPaper
                                    '处理子项
                                    For Each tmpID In objReport.Items("_" & ObjSel.Index).SubIDs
                                        objReport.Items("_" & tmpID.id).父ID = 0
                                    Next
                                Else
                                    mlngY = ObjSel.Top + ObjSel.Container.Top
                                    mlngX = ObjSel.Left + ObjSel.Container.Left
                                    Set mobjMove = picPaper
                                End If
                            End If
                        Else
                            Set mobjMove = Nothing: mlngX = 0: mlngY = 0
                        End If
                    Else
                        '如果不是子元素，则判断是否移动到了卡片中
                        For Each tmpItem In objReport.Items
                            If tmpItem.类型 = 14 And tmpItem.格式号 = mbytCurrFmt Then
                                If ObjSel.Top >= tmpItem.Y And ObjSel.Left >= tmpItem.X And _
                                    ObjSel.Height + ObjSel.Top <= tmpItem.Y + tmpItem.H And ObjSel.Width + ObjSel.Left <= tmpItem.X + tmpItem.W Then
                                    mlngY = ObjSel.Top - tmpItem.Y
                                    mlngX = ObjSel.Left - tmpItem.X
                                    Set mobjMove = pic(tmpItem.id)
                                    Exit For
                                Else
                                    mlngY = ObjSel.Top
                                    mlngX = ObjSel.Left
                                    Set mobjMove = picPaper
                                End If
                            End If
                        Next
                    End If
                End If
            
                '更新数据对象集
                objReport.Items("_" & tmpObj.Tag).X = Format(ObjSel.Left / sgnMode, "0.00")
                objReport.Items("_" & tmpObj.Tag).Y = Format(ObjSel.Top / sgnMode, "0.00")
            End If
        End If
NextObj:
        If tmpObj.Index <> 0 Then
            If blnMove Then
                With WinProperty
                    .T = (tmpObj.Top + lngY) / Screen.TwipsPerPixelX
                    .l = (tmpObj.Left + lngX) / Screen.TwipsPerPixelX
                    .H = tmpObj.Height / Screen.TwipsPerPixelX
                    .W = tmpObj.Width / Screen.TwipsPerPixelX
                    Call MoveWindow(tmpObj.hwnd, .l, .T, .W, .H, 1)
                End With
            End If
        End If

    Next
    
    Me.Refresh
    BlnSave = False
End Sub

Private Sub ResetColor(idx As Integer)
'功能：将表格编辑颜色恢复为表格本色
    Dim i As Integer, j As Integer
    Dim lngRow As Integer, lngCol As Integer
    msh(idx).Redraw = False
    lngRow = msh(idx).Row: lngCol = msh(idx).Col
    For i = 0 To msh(idx).Rows - 1
        msh(idx).Row = i
        For j = 0 To msh(idx).Cols - 1
            msh(idx).Col = j
            If i < msh(idx).FixedRows Or j < msh(idx).FixedCols Then
                msh(idx).CellBackColor = msh(idx).BackColorFixed
                'msh(idx).CellForeColor = msh(idx).ForeColorFixed
            Else
                msh(idx).CellBackColor = msh(idx).BackColor
                msh(idx).CellForeColor = msh(idx).ForeColor
            End If
        Next
    Next
    msh(idx).Row = lngRow: msh(idx).Col = lngCol
    msh(idx).Redraw = True
End Sub

Private Sub SetGridSame(mshS As Control, mshO As Control)
'功能:设置两个网格控件处观相同
'说明：消耗时间与行列数成正比
    Dim i As Integer, j As Integer
    
    mshO.Redraw = False
    mshS.Redraw = False
    
    mshO.Width = mshS.Width
    mshO.Height = mshS.Height
    mshO.Rows = mshS.Rows
    mshO.Cols = mshS.Cols
    mshO.FixedCols = mshS.FixedCols
    mshO.FixedRows = mshS.FixedRows
    
    mshO.ForeColor = mshS.ForeColor
    mshO.BackColor = mshS.BackColor
    mshO.BackColorFixed = mshS.BackColorFixed
    mshO.ForeColorFixed = mshS.ForeColorFixed
    mshO.BackColorSel = mshS.BackColorSel
    mshO.ForeColorSel = mshS.ForeColorSel
    mshO.GridColor = mshS.GridColor
    mshO.GridColorFixed = mshS.GridColorFixed
    
    mshO.Font.Size = mshS.Font.Size
    mshO.Font.name = mshS.Font.name
    mshO.Font.Bold = mshS.Font.Bold
    mshO.Font.Underline = mshS.Font.Underline
    mshO.Font.Italic = mshS.Font.Italic
    
    For i = 0 To mshS.Rows - 1
        mshS.Row = i: mshO.Row = i
        mshO.RowHeight(i) = mshS.RowHeight(i)
        mshO.MergeRow(i) = mshS.MergeRow(i)
        For j = 0 To mshS.Cols - 1
            mshS.Col = j: mshO.Col = j
            mshO.CellAlignment = mshS.CellAlignment
            mshO.CellFontBold = mshS.CellFontBold
            mshO.CellFontName = mshS.CellFontName
            mshO.CellFontSize = mshS.CellFontSize
            mshO.CellFontItalic = mshS.CellFontItalic
            mshO.CellFontUnderline = mshS.CellFontUnderline
            mshO.TextMatrix(i, j) = mshS.TextMatrix(i, j)
            If i <= mshS.FixedRows - 1 Or j <= mshS.FixedCols - 1 Then
                mshO.CellBackColor = mshS.BackColorFixed
                mshO.CellForeColor = mshS.ForeColorFixed
            Else
                mshO.CellBackColor = mshS.BackColor
                mshO.CellForeColor = mshS.ForeColor
            End If
        Next
    Next
    For i = 0 To mshS.Cols - 1
        mshO.ColWidth(i) = mshS.ColWidth(i)
        mshO.ColAlignment(i) = mshS.ColAlignment(i)
        mshO.MergeCol(i) = mshS.MergeCol(i)
    Next
    
    mshO.Redraw = True
    mshS.Redraw = True
End Sub

Private Sub SetGridLike(mshS As Control, mshO As Control)
'功能:设置两个网格控件处观相同(仅字体，行高及颜色相同)
'说明：消耗时间与行列数成正比
    Dim i As Integer, j As Integer
    
    mshO.Redraw = False
    mshS.Redraw = False
    
    mshO.ForeColor = mshS.ForeColor
    mshO.BackColor = mshS.BackColor
    mshO.BackColorFixed = mshS.BackColorFixed
    mshO.ForeColorFixed = mshS.ForeColorFixed
    mshO.BackColorSel = mshS.BackColorSel
    mshO.ForeColorSel = mshS.ForeColorSel
    mshO.GridColor = mshS.GridColor
    mshO.GridColorFixed = mshS.GridColorFixed
    
    mshO.Font.Size = mshS.Font.Size
    mshO.Font.name = mshS.Font.name
    mshO.Font.Bold = mshS.Font.Bold
    mshO.Font.Underline = mshS.Font.Underline
    mshO.Font.Italic = mshS.Font.Italic
    mshO.RowHeightMin = mshS.RowHeightMin
    objReport.Items("_" & mshO.Index).字号 = objReport.Items("_" & mshS.Index).字号
    objReport.Items("_" & mshO.Index).字体 = objReport.Items("_" & mshS.Index).字体
    objReport.Items("_" & mshO.Index).斜体 = objReport.Items("_" & mshS.Index).斜体
    objReport.Items("_" & mshO.Index).粗体 = objReport.Items("_" & mshS.Index).粗体
    objReport.Items("_" & mshO.Index).下线 = objReport.Items("_" & mshS.Index).下线
    objReport.Items("_" & mshO.Index).背景 = objReport.Items("_" & mshS.Index).背景
    objReport.Items("_" & mshO.Index).前景 = objReport.Items("_" & mshS.Index).前景
    objReport.Items("_" & mshO.Index).网格 = objReport.Items("_" & mshS.Index).网格
    objReport.Items("_" & mshO.Index).行高 = objReport.Items("_" & mshS.Index).行高
    
    For i = 0 To mshS.Rows - 1
        If i <= mshO.Rows - 1 Then
            mshS.Row = i: mshO.Row = i
            mshO.RowHeight(i) = mshS.RowHeight(i)
            For j = 0 To mshS.Cols - 1
                If j <= mshO.Cols - 1 Then
                    mshS.Col = j
                    mshO.Col = j
                    mshO.CellFontBold = mshS.CellFontBold
                    mshO.CellFontName = mshS.CellFontName
                    mshO.CellFontSize = mshS.CellFontSize
                    mshO.CellFontItalic = mshS.CellFontItalic
                    mshO.CellFontUnderline = mshS.CellFontUnderline
                    If i <= mshS.FixedRows - 1 Or j <= mshS.FixedCols - 1 Then
                        mshO.CellBackColor = mshS.BackColorFixed
                        mshO.CellForeColor = mshS.ForeColorFixed
                    Else
                        mshO.CellBackColor = mshS.BackColor
                        mshO.CellForeColor = mshS.ForeColor
                    End If
                End If
            Next
        End If
    Next
    
    mshO.Redraw = True
    mshS.Redraw = True
End Sub

Private Sub SetCopyGrid(intIdx As Integer)
'功能:对指定数据表的分栏进行调整
    Dim i As Integer
    Dim tmpID As RelatID
    i = 0
    For Each tmpID In objReport.Items("_" & intIdx).CopyIDs
        i = i + 1
        Call SetGridSame(msh(intIdx), msh(tmpID.id))
        msh(tmpID.id).Top = msh(intIdx).Top
        msh(tmpID.id).Left = msh(intIdx).Left + (msh(intIdx).Width - 15) * i
    Next
End Sub

Private Sub SetSelAlign(bytAlign As Byte)
'功能:设置选中控件对齐
'参数:
'     bytAlign=1:左对齐,2:右对齐,3:上对齐,4:下对齐,5:水平居中对齐,6:垂直居中对齐,7:相同宽度,8:相同高度,9:宽高相同
'说明：设置7,8,9时,注意表格的最小宽度及高度

    Dim tmpObj As PictureBox, ObjSel As Control, tmpID As RelatID
    Dim ItemSend As RPTItem
    Dim lngPreX As Long, lngPreY As Long
    Dim lngOffX As Long, lngOffY As Long '前后偏移量
    Dim lngMinW As Long, lngMinH As Long, i As Integer
    Dim xx As Integer, yy As Integer, zz As Integer
    
    If GetSelNum < 2 Then Exit Sub
    If objLastSel Is Nothing Then Exit Sub
    
    For Each tmpObj In lblSize
        If tmpObj.Index Mod 8 = 1 And tmpObj.Tag <> objLastSel.Index Then
            lngMinW = 0: lngMinH = 0
            Select Case objReport.Items("_" & tmpObj.Tag).类型
                Case 1
                    Set ObjSel = lblLine(tmpObj.Tag)
                Case 2, 3
                    Set ObjSel = lbl(tmpObj.Tag)
                Case 10
                    Set ObjSel = Shp(tmpObj.Tag)
                Case 11
                    Set ObjSel = img(tmpObj.Tag)
                Case 14
                    Set ObjSel = pic(tmpObj.Tag)
                Case 4
                    Set ObjSel = msh(tmpObj.Tag)
                    lngMinW = msh(tmpObj.Tag).ColWidth(0) + 15
                    lngMinH = msh(tmpObj.Tag).RowHeight(0) * (msh(tmpObj.Tag).FixedRows + 2) + 15
                Case 5
                    Set ObjSel = msh(tmpObj.Tag)
                    xx = msh(tmpObj.Tag).FixedCols '纵向分类项目数
                    yy = msh(tmpObj.Tag).FixedRows - 1 '横向分类项目数
                    For Each tmpID In objReport.Items("_" & tmpObj.Tag).SubIDs
                        If objReport.Items("_" & tmpID.id).类型 = 9 Then zz = zz + 1 '统计项目数
                    Next
                    lngMinH = msh(tmpObj.Tag).RowHeight(0) * (yy + 1) + 15
                    For i = 0 To xx + zz - 1
                        lngMinW = lngMinW + msh(tmpObj.Tag).ColWidth(i)
                    Next
                    lngMinW = lngMinW + 60
                Case 12 '@@@
                    Set ObjSel = Chart(tmpObj.Tag)
                    lngMinW = Chart(0).Width: lngMinH = Chart(0).Height
                Case 13
                    Set ObjSel = ImgCode(tmpObj.Tag)
            End Select
                        
            '@@@
            lngMinW = lngMinW * sgnMode
            lngMinH = lngMinH * sgnMode
            
            If bytAlign < 7 Then '对齐设置
                lngPreX = ObjSel.Left: lngPreY = ObjSel.Top
                Select Case bytAlign
                    Case 1
                        ObjSel.Left = objLastSel.Left
                    Case 2
                        ObjSel.Left = objLastSel.Left + objLastSel.Width - ObjSel.Width
                    Case 3
                        ObjSel.Top = objLastSel.Top
                    Case 4
                        ObjSel.Top = objLastSel.Top + objLastSel.Height - ObjSel.Height
                    Case 5
                        ObjSel.Top = objLastSel.Top + (objLastSel.Height - ObjSel.Height) / 2
                    Case 6
                        ObjSel.Left = objLastSel.Left + (objLastSel.Width - ObjSel.Width) / 2
                End Select
                
                If objReport.Items("_" & tmpObj.Tag).类型 = 12 Then '@@@
                    If ObjSel.Left < 0 Then ObjSel.Left = 0
                    If ObjSel.Top < 0 Then ObjSel.Top = 0
                End If
                
                lngOffX = ObjSel.Left - lngPreX
                lngOffY = ObjSel.Top - lngPreY
                
                For i = CInt(Mid(ObjSel.Tag, 3)) To CInt(Mid(ObjSel.Tag, 3)) + 7
                    lblSize(i).Left = lblSize(i).Left + lngOffX
                    lblSize(i).Top = lblSize(i).Top + lngOffY
                Next
                
                '更改数据对象
                objReport.Items("_" & tmpObj.Tag).X = Format(ObjSel.Left / sgnMode, "0.00")
                objReport.Items("_" & tmpObj.Tag).Y = Format(ObjSel.Top / sgnMode, "0.00")
            Else '尺寸设置
                Select Case bytAlign
                    Case 7
                        If Not (objReport.Items("_" & tmpObj.Tag).类型 = 1 And objReport.Items("_" & tmpObj.Tag).W < objReport.Items("_" & tmpObj.Tag).H) Then
                            If objReport.Items("_" & tmpObj.Tag).类型 = 4 Or objReport.Items("_" & tmpObj.Tag).类型 = 5 Then
                                If objLastSel.Width < lngMinW Then
                                    ObjSel.Width = lngMinW
                                Else
                                    ObjSel.Width = objLastSel.Width
                                End If
                            Else
                                ObjSel.Width = objLastSel.Width
                            End If
                        End If
                    Case 8
                        If Not (objReport.Items("_" & tmpObj.Tag).类型 = 1 And objReport.Items("_" & tmpObj.Tag).H < objReport.Items("_" & tmpObj.Tag).W) Then
                            If objReport.Items("_" & tmpObj.Tag).类型 = 4 Or objReport.Items("_" & tmpObj.Tag).类型 = 5 Then
                                If objLastSel.Height < lngMinH Then
                                    ObjSel.Height = lngMinH
                                Else
                                    ObjSel.Height = objLastSel.Height
                                End If
                            Else
                                ObjSel.Height = objLastSel.Height
                            End If
                        End If
                    Case 9
                        If Not (objReport.Items("_" & tmpObj.Tag).类型 = 1 And objReport.Items("_" & tmpObj.Tag).W < objReport.Items("_" & tmpObj.Tag).H) Then
                            If objReport.Items("_" & tmpObj.Tag).类型 = 4 Or objReport.Items("_" & tmpObj.Tag).类型 = 5 Then
                                If objLastSel.Width < lngMinW Then
                                    ObjSel.Width = lngMinW
                                Else
                                    ObjSel.Width = objLastSel.Width
                                End If
                            Else
                                ObjSel.Width = objLastSel.Width
                            End If
                        End If
                        If Not (objReport.Items("_" & tmpObj.Tag).类型 = 1 And objReport.Items("_" & tmpObj.Tag).H < objReport.Items("_" & tmpObj.Tag).W) Then
                            If objReport.Items("_" & tmpObj.Tag).类型 = 4 Or objReport.Items("_" & tmpObj.Tag).类型 = 5 Then
                                If objLastSel.Height < lngMinH Then
                                    ObjSel.Height = lngMinH
                                Else
                                    ObjSel.Height = objLastSel.Height
                                End If
                            Else
                                ObjSel.Height = objLastSel.Height
                            End If
                        End If
                End Select
                
                If objReport.Items("_" & tmpObj.Tag).类型 = 12 Then '@@@
                    If ObjSel.Left < 0 Then ObjSel.Left = 0
                    If ObjSel.Top < 0 Then ObjSel.Top = 0
                End If
                
                For i = CInt(Mid(ObjSel.Tag, 3)) To CInt(Mid(ObjSel.Tag, 3)) + 7
                    Select Case IIF(i Mod 8 <> 0, i Mod 8, 8) '定位选择边框的位置
                        Case 1 '上中
                            lblSize(i).Top = ObjSel.Top - lblSize(i).Height
                            lblSize(i).Left = ObjSel.Left + (ObjSel.Width - lblSize(i).Width) / 2
                        Case 2 '上右
                            lblSize(i).Top = ObjSel.Top - lblSize(i).Height
                            lblSize(i).Left = ObjSel.Left + ObjSel.Width
                        Case 3 '右中
                            lblSize(i).Top = ObjSel.Top + (ObjSel.Height - lblSize(i).Height) / 2
                            lblSize(i).Left = ObjSel.Left + ObjSel.Width
                        Case 4 '右下
                            lblSize(i).Top = ObjSel.Top + ObjSel.Height
                            lblSize(i).Left = ObjSel.Left + ObjSel.Width
                        Case 5 '下中
                            lblSize(i).Top = ObjSel.Top + ObjSel.Height
                            lblSize(i).Left = ObjSel.Left + (ObjSel.Width - lblSize(i).Width) / 2
                        Case 6 '左下
                            lblSize(i).Top = ObjSel.Top + ObjSel.Height
                            lblSize(i).Left = ObjSel.Left - lblSize(i).Width
                        Case 7 '左中
                            lblSize(i).Top = ObjSel.Top + (ObjSel.Height - lblSize(i).Height) / 2
                            lblSize(i).Left = ObjSel.Left - lblSize(i).Width
                        Case 8 '左上
                            lblSize(i).Top = ObjSel.Top - lblSize(i).Height
                            lblSize(i).Left = ObjSel.Left - lblSize(i).Width
                    End Select
                Next
                
                '更改数据对象
                If Not (objReport.Items("_" & tmpObj.Tag).类型 = 1 And objReport.Items("_" & tmpObj.Tag).W < objReport.Items("_" & tmpObj.Tag).H) Then objReport.Items("_" & tmpObj.Tag).W = Format(ObjSel.Width / sgnMode, "0.00")
                If Not (objReport.Items("_" & tmpObj.Tag).类型 = 1 And objReport.Items("_" & tmpObj.Tag).H < objReport.Items("_" & tmpObj.Tag).W) Then objReport.Items("_" & tmpObj.Tag).H = Format(ObjSel.Height / sgnMode, "0.00")
            End If
            
            If objReport.Items("_" & tmpObj.Tag).类型 = 4 Or objReport.Items("_" & tmpObj.Tag).类型 = 5 Then
                Call SetGridLine(tmpObj.Tag)
            End If
            
            '调整分栏
            If objReport.Items("_" & tmpObj.Tag).类型 = 4 Then
                Call SetCopyGrid(tmpObj.Tag)
            End If
        End If
        
        If Not ObjSel Is Nothing Then
            If InStr(1, "4,5", objReport.Items("_" & ObjSel.Index).类型) <> 0 And objReport.Items("_" & ObjSel.Index).参照 = "" Then
                Call CopyItem(ItemSend, objReport.Items("_" & ObjSel.Index))
                Call SetChildWH(ObjSel.Index)
                
                Dim ResizeItem As RPTItem, IntLastCurID As Integer
                IntLastCurID = intCurID
                For Each ResizeItem In objReport.Items
                    If ResizeItem.格式号 = mbytCurrFmt And ResizeItem.参照 = ItemSend.名称 And ResizeItem.类型 = 2 Then
                        intCurID = ResizeItem.Key
                        Call ReferTo
                    End If
                Next
                intCurID = IntLastCurID
            ElseIf objReport.Items("_" & ObjSel.Index).类型 = 2 And objReport.Items("_" & ObjSel.Index).参照 <> "" Then
                IntLastCurID = intCurID
                intCurID = ObjSel.Index
                Call ReferTo
                intCurID = IntLastCurID
            End If
        End If
    Next
    BlnSave = False
End Sub

Private Sub tbr2_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Call NoneEdit
    Select Case ButtonMenu.Key
        Case "HscSame"
            mnuFormat_HscSpace_Same_Click
        Case "HscAdd"
            mnuFormat_HscSpace_Add_Click
        Case "HscDec"
            mnuFormat_HscSpace_Dec_Click
        Case "VscSame"
            mnuFormat_VscSpace_Same_Click
        Case "VscAdd"
            mnuFormat_VscSpace_Add_Click
        Case "VscDec"
            mnuFormat_VscSpace_Dec_Click
        Case "Page", "Width", "Height", "Scale200", "Scale100", "Scale75", "Scale50", "Scale25"
            mnuViewScaleMode_Click ButtonMenu.Index - 1
    End Select
End Sub

Private Sub tbr2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuView, 2
End Sub

Private Sub tbr1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuView, 2
End Sub

Private Sub SetVscSpace(bytType As Integer)
'功能:调整选中控件的垂直间隔
'参数:bytType=0:相同,-1:减少,1增加
    Dim i As Integer, j As Integer, lngH As Long
    Dim tmpObj As PictureBox, ObjSel As Control, arrObj() As Control '最终按从上到下的顺序存放选中控件
    Dim ItemSend As RPTItem
    
    Const SPACE_STEP As Long = 75 '一次增加或减少的间隔
    
    On Error Resume Next
    
    i = GetSelNum()
    If i < 2 Or (i = 2 And bytType = 0) Then Exit Sub '当要调整间隔相同时,选中控件数必须大于3
    
    '形成选中控件数组
    ReDim arrObj(i - 1)
    i = 0
    For Each tmpObj In lblSize
        If tmpObj.Index Mod 8 = 1 Then
            Select Case objReport.Items("_" & tmpObj.Tag).类型
                Case 1
                    Set arrObj(i) = lblLine(tmpObj.Tag)
                Case 2, 3
                    Set arrObj(i) = lbl(tmpObj.Tag)
                Case 10
                    Set arrObj(i) = Shp(tmpObj.Tag)
                Case 11
                    Set arrObj(i) = img(tmpObj.Tag)
                Case 4, 5
                    Set arrObj(i) = msh(tmpObj.Tag)
                Case 12 '@@@
                    Set arrObj(i) = Chart(tmpObj.Tag)
                Case 13
                    Set arrObj(i) = ImgCode(tmpObj.Tag)
                Case 14
                    Set arrObj(i) = pic(tmpObj.Tag)
            End Select
            i = i + 1
        End If
    Next
    '对控件数组按Top从小到大的顺序顺列
    For i = 0 To UBound(arrObj) - 1
        For j = i + 1 To UBound(arrObj)
            If arrObj(j).Top < arrObj(i).Top Then
                Set ObjSel = arrObj(j)
                Set arrObj(j) = arrObj(i)
                Set arrObj(i) = ObjSel
            End If
        Next
    Next
    Select Case bytType
        Case 0
            '求平均间隔
            lngH = 0
            For i = 0 To UBound(arrObj) - 1
                lngH = lngH + (arrObj(i + 1).Top - (arrObj(i).Top + arrObj(i).Height))
            Next
            lngH = lngH \ UBound(arrObj)
            '定位项目
            For i = 1 To UBound(arrObj)
                '当定位时移动项目后,数组中对应的项目值对相应发生变化(SET)
                Call SeekItem(arrObj(i), arrObj(i).Left, arrObj(i - 1).Top + arrObj(i - 1).Height + lngH)
            Next
        Case 1
            lngH = SPACE_STEP
            For i = 1 To UBound(arrObj)
                Call SeekItem(arrObj(i), arrObj(i).Left, arrObj(i).Top + lngH)
                lngH = lngH + SPACE_STEP
            Next
        Case -1
            lngH = SPACE_STEP
            For i = 1 To UBound(arrObj)
                Call SeekItem(arrObj(i), arrObj(i).Left, arrObj(i).Top - lngH)
                lngH = lngH + SPACE_STEP
            Next
    End Select
    
    For i = 0 To UBound(arrObj) - 1
        If Not arrObj(i) Is Nothing Then
            If InStr(1, "4,5", objReport.Items("_" & arrObj(i).Index).类型) <> 0 And objReport.Items("_" & arrObj(i).Index).参照 = "" Then
                Call CopyItem(ItemSend, objReport.Items("_" & arrObj(i).Index))
                Call SetChildWH(arrObj(i).Index)
                Dim ResizeItem As RPTItem, IntLastCurID As Integer
                IntLastCurID = intCurID
                For Each ResizeItem In objReport.Items
                    If ResizeItem.格式号 = mbytCurrFmt And ResizeItem.参照 = ItemSend.名称 And ResizeItem.类型 = 2 Then
                        intCurID = ResizeItem.Key
                        Call ReferTo
                    End If
                Next
                intCurID = IntLastCurID
            ElseIf objReport.Items("_" & arrObj(i).Index).类型 = 2 And objReport.Items("_" & arrObj(i).Index).参照 <> "" Then
                IntLastCurID = intCurID
                intCurID = arrObj(i).Index
                Call ReferTo
                intCurID = IntLastCurID
            End If
        End If
    Next
    BlnSave = False
End Sub

Private Sub SetHscSpace(bytType As Integer)
'功能:调整选中控件的水平间隔
'参数:bytType=0:相同,-1:减少,1增加
    Dim i As Integer, j As Integer, lngW As Long
    Dim tmpObj As PictureBox, ObjSel As Control, arrObj() As Control  '最终按从左到右的顺序存放选中控件
    Dim ItemSend As RPTItem
    
    Const SPACE_STEP As Long = 75 '一次增加或减少的间隔
    
    i = GetSelNum()
    If i < 2 Or (i = 2 And bytType = 0) Then Exit Sub '当要调整间隔相同时,选中控件数必须大于3
    
    '形成选中控件数组
    ReDim arrObj(i - 1)
    i = 0
    For Each tmpObj In lblSize
        If tmpObj.Index Mod 8 = 1 Then
            Select Case objReport.Items("_" & tmpObj.Tag).类型
                Case 1
                    Set arrObj(i) = lblLine(tmpObj.Tag)
                Case 2, 3
                    Set arrObj(i) = lbl(tmpObj.Tag)
                Case 10
                    Set arrObj(i) = Shp(tmpObj.Tag)
                Case 11
                    Set arrObj(i) = img(tmpObj.Tag)
                Case 4, 5
                    Set arrObj(i) = msh(tmpObj.Tag)
                Case 12 '@@@
                    Set arrObj(i) = Chart(tmpObj.Tag)
                Case 13
                    Set arrObj(i) = ImgCode(tmpObj.Tag)
                Case 14
                    Set arrObj(i) = pic(tmpObj.Tag)
            End Select
            i = i + 1
        End If
    Next
    '对控件数组按Left从小到大的顺序顺列
    For i = 0 To UBound(arrObj) - 1
        For j = i + 1 To UBound(arrObj)
            If arrObj(j).Left < arrObj(i).Left Then
                Set ObjSel = arrObj(j)
                Set arrObj(j) = arrObj(i)
                Set arrObj(i) = ObjSel
            End If
        Next
    Next
    Select Case bytType
        Case 0
            '求平均间隔
            lngW = 0
            For i = 0 To UBound(arrObj) - 1
                lngW = lngW + (arrObj(i + 1).Left - (arrObj(i).Left + arrObj(i).Width))
            Next
            lngW = lngW \ UBound(arrObj)
            '定位项目
            For i = 1 To UBound(arrObj)
                '当定位时移动项目后,数组中对应的项目值对相应发生变化(SET)
                Call SeekItem(arrObj(i), arrObj(i - 1).Left + arrObj(i - 1).Width + lngW, arrObj(i).Top)
            Next
        Case 1
            lngW = SPACE_STEP
            For i = 1 To UBound(arrObj)
                Call SeekItem(arrObj(i), arrObj(i).Left + lngW, arrObj(i).Top)
                lngW = lngW + SPACE_STEP
            Next
        Case -1
            lngW = SPACE_STEP
            For i = 1 To UBound(arrObj)
                Call SeekItem(arrObj(i), arrObj(i).Left - lngW, arrObj(i).Top)
                lngW = lngW + SPACE_STEP
            Next
    End Select
    For i = 0 To UBound(arrObj) - 1
        If Not arrObj(i) Is Nothing Then
            If InStr(1, "4,5", objReport.Items("_" & arrObj(i).Index).类型) <> 0 And objReport.Items("_" & arrObj(i).Index).参照 = "" Then
                Call CopyItem(ItemSend, objReport.Items("_" & arrObj(i).Index))
                Call SetChildWH(arrObj(i).Index)
                Dim ResizeItem As RPTItem, IntLastCurID As Integer
                IntLastCurID = intCurID
                For Each ResizeItem In objReport.Items
                    If ResizeItem.格式号 = mbytCurrFmt And ResizeItem.参照 = ItemSend.名称 And ResizeItem.类型 = 2 Then
                        intCurID = ResizeItem.Key
                        Call ReferTo
                    End If
                Next
                intCurID = IntLastCurID
            ElseIf objReport.Items("_" & arrObj(i).Index).类型 = 2 And objReport.Items("_" & arrObj(i).Index).参照 <> "" Then
                IntLastCurID = intCurID
                intCurID = arrObj(i).Index
                Call ReferTo
                intCurID = IntLastCurID
            End If
        End If
    Next
    BlnSave = False
End Sub

Private Sub SeekItem(objSeek As Control, X As Long, Y As Long)
'功能:定位项目
'参数:objSeek=报表项目
'说明:该函数主要被水平和垂直间隔调整函数所调用
    Dim i As Byte
    Dim lngTop As Long, lngLeft As Long
    
    objSeek.Top = Y: objSeek.Left = X
    If objReport.Items("_" & objSeek.Index).类型 = 12 Then '@@@
        If objSeek.Top < 0 Then objSeek.Top = 0
        If objSeek.Left < 0 Then objSeek.Left = 0
    End If
    If UCase(objSeek.Container.name) = "PIC" Then
        lngTop = objSeek.Container.Top
        lngLeft = objSeek.Container.Left
    End If
    
    If Mid(objSeek.Tag, 1, 2) = "S_" Then
        For i = CInt(Mid(objSeek.Tag, 3)) To CInt(Mid(objSeek.Tag, 3)) + 7 '移动Size标记
            Select Case IIF(i Mod 8 <> 0, i Mod 8, 8) '定位选择边框的位置
                Case 1 '上中
                    lblSize(i).Top = objSeek.Top + lngTop - lblSize(i).Height
                    lblSize(i).Left = objSeek.Left + lngLeft + (objSeek.Width - lblSize(i).Width) / 2
                Case 2 '上右
                    lblSize(i).Top = objSeek.Top + lngTop - lblSize(i).Height
                    lblSize(i).Left = objSeek.Left + lngLeft + objSeek.Width
                Case 3 '右中
                    lblSize(i).Top = objSeek.Top + lngTop + (objSeek.Height - lblSize(i).Height) / 2
                    lblSize(i).Left = objSeek.Left + lngLeft + objSeek.Width
                Case 4 '右下
                    lblSize(i).Top = objSeek.Top + lngTop + objSeek.Height
                    lblSize(i).Left = objSeek.Left + lngLeft + objSeek.Width
                Case 5 '下中
                    lblSize(i).Top = objSeek.Top + lngTop + objSeek.Height
                    lblSize(i).Left = objSeek.Left + lngLeft + (objSeek.Width - lblSize(i).Width) / 2
                Case 6 '左下
                    lblSize(i).Top = objSeek.Top + lngTop + objSeek.Height
                    lblSize(i).Left = objSeek.Left + lngLeft - lblSize(i).Width
                Case 7 '左中
                    lblSize(i).Top = objSeek.Top + lngTop + (objSeek.Height - lblSize(i).Height) / 2
                    lblSize(i).Left = objSeek.Left + lngLeft - lblSize(i).Width
                Case 8 '左上
                    lblSize(i).Top = objSeek.Top + lngTop - lblSize(i).Height
                    lblSize(i).Left = objSeek.Left + lngLeft - lblSize(i).Width
            End Select
        Next
    End If
    '调整分栏
    If objReport.Items("_" & objSeek.Index).类型 = 4 Then
        Call SetCopyGrid(objSeek.Index)
    End If
    
    '更新数据对象集
    objReport.Items("_" & objSeek.Index).X = Format(objSeek.Left / sgnMode, "0.00")
    objReport.Items("_" & objSeek.Index).Y = Format(objSeek.Top / sgnMode, "0.00")
End Sub

Private Sub SetSelCenter(bytStyle As Byte)
'功能：设置选择控件水平居中或垂直居中
'参数：bytStyle=0:水平居中,1:垂直居中
    Dim tmpObj As PictureBox, ObjSel As Object
    Dim ItemSend As RPTItem, objFmt As RPTFmt
    Dim lngW As Long, lngH As Long
    
    If GetSelNum = 0 Then Exit Sub
    
    Set objFmt = objReport.Fmts("_" & mbytCurrFmt)
    If objFmt.纸向 = 1 Then
        lngW = objFmt.W
        lngH = objFmt.H
    Else
        lngW = objFmt.H
        lngH = objFmt.W
    End If
    For Each tmpObj In lblSize
        If tmpObj.Index Mod 8 = 1 Then
            Set ObjSel = GetInxObj(tmpObj.Tag)
            If bytStyle = 0 Then
                SeekItem ObjSel, (lngW - ObjSel.Width) / 2, ObjSel.Top
            Else
                SeekItem ObjSel, ObjSel.Left, (lngH - ObjSel.Height) / 2
            End If
            If Not ObjSel Is Nothing Then
                If InStr(1, "4,5", objReport.Items("_" & ObjSel.Index).类型) <> 0 And objReport.Items("_" & ObjSel.Index).参照 = "" Then
                    Call CopyItem(ItemSend, objReport.Items("_" & ObjSel.Index))
                    Call SetChildWH(ObjSel.Index)
                    Dim ResizeItem As RPTItem, IntLastCurID As Integer
                    IntLastCurID = intCurID
                    For Each ResizeItem In objReport.Items
                        If ResizeItem.格式号 = mbytCurrFmt And ResizeItem.参照 = ItemSend.名称 And ResizeItem.类型 = 2 Then
                            intCurID = ResizeItem.Key
                            Call ReferTo
                        End If
                    Next
                    intCurID = IntLastCurID
                ElseIf objReport.Items("_" & ObjSel.Index).类型 = 2 And objReport.Items("_" & ObjSel.Index).参照 <> "" Then
                    IntLastCurID = intCurID
                    intCurID = ObjSel.Index
                    Call ReferTo
                    intCurID = IntLastCurID
                End If
            End If
        End If
    Next
    If GetSelNum = 1 Then Call ShowAttrib(intCurID)
    BlnSave = False
End Sub

Private Sub txtAtt_GotFocus()
    SelAll txtAtt
End Sub

Private Sub txtAtt_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            mshAtt.SetFocus: mshAtt.Refresh: SendKeys "{DOWN}"
        Case vbKeyUp
            mshAtt.SetFocus: mshAtt.Refresh: SendKeys "{UP}"
    End Select
End Sub

Private Sub txtAtt_KeyPress(KeyAscii As Integer)
    Dim ObjSel As Object, tmpID As RelatID, tmpItem As RPTItem
    Dim i As Integer, xx As Long, yy As Long, zz As Long
    Dim lngMinW As Long, lngMinH As Long, sgnH As Single
    Dim arrHead, arrModify
    Dim objBarCode As StdPicture, lngSize As Long
    Dim strBarCode As String, sngWidth As Single
    Dim lngL As Long, lngW As Long
    Dim strInfo As String
    Dim lngReportID As Long
    Dim strReportID As String
    Dim X As Long, Y As Long, k As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        '非法或分隔字符:
        If InString(txtAtt.Text, "'|~^") Then
            MsgBox "输入了非法字符！", vbInformation, App.Title
            txtAtt.SetFocus: Exit Sub
        End If
        Set ObjSel = GetInxObj(intCurID)

        Select Case mshAtt.TextMatrix(mshAtt.Row, 0)
            Case "名称"
                txtAtt = Trim(txtAtt.Text)
                If TLen(txtAtt.Text) > 50 Then
                    MsgBox "名称不能超过50个字符！", vbInformation, App.Title
                    txtAtt.SetFocus: Exit Sub
                End If
                If txtAtt.Text = "" Then
                    MsgBox "名称不能为空！", vbInformation, App.Title
                    txtAtt.SetFocus: Exit Sub
                End If
                
                '检查名称的合法性
                If CheckNameValid(txtAtt.Text) = False Then
                    MsgBox "在现有的报表格式中发现名称重复！", vbInformation, App.Title
                    txtAtt.SetFocus
                    Exit Sub
                End If
                Call ChangeReferTo(objReport.Items("_" & intCurID).名称, txtAtt.Text)      '修改所有子表的参照对象
                mshAtt.TextMatrix(mshAtt.Row, 1) = txtAtt.Text
                objReport.Items("_" & intCurID).名称 = txtAtt.Text
                
                txtAtt.Visible = False: mshAtt.SetFocus
                BlnSave = False
            Case "内容"
                If TLen(txtAtt.Text) > 255 Then
                    MsgBox "内容不能超过255个字符！", vbInformation, App.Title
                    txtAtt.SetFocus: Exit Sub
                End If
                
                '条码非法字符检查
                If objReport.Items("_" & intCurID).类型 = 13 Then
                    If InString(txtAtt.Text, "[]") And Not BracketMatch(txtAtt.Text, "[]") Then
                        MsgBox "条码数据的括号不配对！", vbInformation, App.Title
                        txtAtt.SetFocus: Exit Sub
                    End If
                    If ReplaceBracket(txtAtt.Text) <> "" Then
                        If objReport.Items("_" & intCurID).序号 = 1 Or objReport.Items("_" & intCurID).序号 = 3 Then
                            If Not MatchString(ReplaceBracket(txtAtt.Text), STR_CODE_128) Then
                                MsgBox "条码内容中包含非法字符！", vbInformation, App.Title
                                txtAtt.SetFocus: Exit Sub
                            End If
                        ElseIf objReport.Items("_" & intCurID).序号 = 2 Then
                            If Not MatchString(ReplaceBracket(txtAtt.Text), STR_CODE_39) Then
                                MsgBox "条码内容中包含非法字符！", vbInformation, App.Title
                                txtAtt.SetFocus: Exit Sub
                            End If
                        End If
                    End If
                End If
                
                Dim strNodeName As String, NodeThis As Node
                '如果是adLongVarBinary型字段,则不允许修改
                xx = InStr(1, txtAtt, "]")
                yy = InStr(1, txtAtt, ".")
                zz = InStr(1, txtAtt, "[")
                If xx > zz And xx > yy And xx <> 0 And zz <> 0 Then
                    strNodeName = Mid(txtAtt, yy + 1, xx - yy - 1)
                    For Each NodeThis In tvwSQL.Nodes
                        If mdlPublic.GetStdNodeText(NodeThis.Text) = strNodeName And IsType(Val(NodeThis.Tag), adLongVarBinary) Then
                            MsgBox "不能选择图型字段为标签的内容！", vbInformation, App.Title
                            mshAtt.TextMatrix(mshAtt.Row, 1) = objReport.Items("_" & intCurID).内容
                            Exit Sub
                        End If
                    Next
                End If
                
                mshAtt.TextMatrix(mshAtt.Row, 1) = txtAtt.Text
                If UCase(TypeName(ObjSel)) = "LABEL" Then ObjSel.Caption = txtAtt.Text
                objReport.Items("_" & intCurID).内容 = txtAtt.Text
            
                '自调后须调整LblSize控件的位置
                If UCase(TypeName(ObjSel)) = "LABEL" Then
                    If ObjSel.AutoSize Then
                        Call SelItem(ObjSel.Index, False)
                        Call SelItem(ObjSel.Index, True)
                    End If
                    objReport.Items("_" & intCurID).W = lbl(intCurID).Width / sgnMode
                ElseIf objReport.Items("_" & intCurID).类型 = 13 Then
                    With objReport.Items("_" & intCurID)
                        strBarCode = ReplaceBracket(.内容)
                        If strBarCode = "" Then strBarCode = "1234567890"
                        
                        Unload frmFlash '强制初始Picture，不然切换绘制有问题
                        If .序号 = 1 Then
                            Set objBarCode = DrawBarCode128(frmFlash.picTemp, 3, strBarCode, Mid(.表头, 1, 1) = "1")
                        ElseIf .序号 = 2 Then
                            Set objBarCode = DrawBarCode39(frmFlash.picTemp, 3, strBarCode, Mid(.表头, 2, 1) = "1", Mid(.表头, 1, 1) = "1")
                        ElseIf .序号 = 3 Then
                            Set objBarCode = DrawBarCode128Auto(frmFlash.picTemp, strBarCode, sngWidth, .行高, Mid(.表头, 1, 1) = "1")
                        ElseIf .序号 = 10 Then
                            Set objBarCode = DrawBarCode2D(strBarCode, frmFlash.picTemp, lngSize)
                        End If
                        If Val(Mid(.表头, 3, 1)) <> 0 Then
                            Set objBarCode = PictureSpin(objBarCode, Val(Mid(.表头, 3, 1)), frmFlash.picTemp)
                        End If
                        Set ObjSel.Picture = objBarCode
                        
                        If .序号 = 3 Then
                            '128码自动调整宽度
                            If Val(Mid(.表头, 3, 1)) = 0 Then
                                ObjSel.Width = Format(Me.ScaleX(sngWidth, vbMillimeters, vbTwips) * sgnMode, "0.00")
                                .W = Me.ScaleX(sngWidth, vbMillimeters, vbTwips)
                            Else
                                ObjSel.Height = Format(Me.ScaleY(sngWidth, vbMillimeters, vbTwips) * sgnMode, "0.00")
                                .H = Me.ScaleY(sngWidth, vbMillimeters, vbTwips)
                            End If
                            Call SeekItem(ObjSel, ObjSel.Left, ObjSel.Top)
                        ElseIf .序号 = 10 And .自调 Then
                            '二维条码缺省自动调整大小
                            .W = lngSize: .H = lngSize
                            
                            ObjSel.Width = Format(lngSize * sgnMode, "0.00")
                            ObjSel.Height = Format(lngSize * sgnMode, "0.00")
                            
                            Call SeekItem(ObjSel, ObjSel.Left, ObjSel.Top)
                        End If
                    End With
                End If
                
                txtAtt.Visible = False: mshAtt.SetFocus
                BlnSave = False
            Case "X坐标"
                If Not IsNumeric(txtAtt.Text) Then
                    MsgBox "请输入数字型数据！", vbInformation, App.Title
                    txtAtt.SetFocus: Exit Sub
                End If
                If Abs(CDbl(txtAtt.Text)) > 5000 Then
                    If CDbl(txtAtt.Text) > 0 Then
                        txtAtt.Text = 5000
                    Else
                        txtAtt.Text = -5000
                    End If
                End If
                If objReport.Items("_" & ObjSel.Index).类型 = 12 Then '@@@
                    If Val(txtAtt.Text) < 0 Then txtAtt.Text = "0.00"
                End If
                
                ObjSel.Left = CDbl(txtAtt.Text) * Twip_mm * sgnMode
                Call SeekItem(ObjSel, ObjSel.Left, ObjSel.Top)
                
                If objReport.Items("_" & intCurID).类型 = 4 Then
                    Call SetCopyGrid(intCurID)
                End If
                
                mshAtt.TextMatrix(mshAtt.Row, 1) = Format(txtAtt.Text, "0.00")
                objReport.Items("_" & intCurID).X = Format(ObjSel.Left / sgnMode, "0.00")
                
                txtAtt.Visible = False: mshAtt.SetFocus
                
                Select Case objReport.Items("_" & ObjSel.Index).类型
                Case 2
                    AdjustAll (True)
                Case 4, 5
                    SetMainWH (ObjSel.Index)
                End Select
                BlnSave = False
            Case "Y坐标"
                If Not IsNumeric(txtAtt.Text) Then
                    MsgBox "请输入数字型数据！", vbInformation, App.Title
                    txtAtt.SetFocus: Exit Sub
                End If
                If Abs(CDbl(txtAtt.Text)) > 5000 Then
                    If CDbl(txtAtt.Text) > 0 Then
                        txtAtt.Text = 5000
                    Else
                        txtAtt.Text = -5000
                    End If
                End If
                If objReport.Items("_" & ObjSel.Index).类型 = 12 Then '@@@
                    If Val(txtAtt.Text) < 0 Then txtAtt.Text = "0.00"
                End If
                
                Call MoveSelect(0, (CDbl(txtAtt.Text) * Twip_mm - ObjSel.Top / sgnMode) * sgnMode)
                
                mshAtt.TextMatrix(mshAtt.Row, 1) = Format(txtAtt.Text, "0.00")
                objReport.Items("_" & intCurID).Y = Format(ObjSel.Top / sgnMode, "0.00")
                
                txtAtt.Visible = False: mshAtt.SetFocus
                
                Select Case objReport.Items("_" & ObjSel.Index).类型
                Case 2
                    AdjustAll (True)
                Case 4, 5
                    SetMainWH (ObjSel.Index)
                End Select
                BlnSave = False
            Case "宽度"
                If Not IsNumeric(txtAtt.Text) Then
                    MsgBox "请输入数字型数据！", vbInformation, App.Title
                    txtAtt.SetFocus: Exit Sub
                End If
                txtAtt.Text = Abs(CDbl(txtAtt.Text))
                If CDbl(txtAtt.Text) > 5000 Then txtAtt.Text = 5000
                
                '检查最小宽度
                If objReport.Items("_" & intCurID).类型 = 4 Then
                    lngMinW = ObjSel.ColWidth(0) + 15
                ElseIf objReport.Items("_" & intCurID).类型 = 5 Then
                    xx = ObjSel.FixedCols '纵向分类项目数
                    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
                        If objReport.Items("_" & tmpID.id).类型 = 9 Then zz = zz + 1 '统计项目数
                    Next
                    For i = 0 To xx + zz - 1
                        lngMinW = lngMinW + ObjSel.ColWidth(i)
                    Next
                    lngMinW = lngMinW + 60
                ElseIf objReport.Items("_" & intCurID).类型 = 12 Then '@@@
                    lngMinW = Chart(0).Width
                End If
                If CDbl(txtAtt.Text) * Twip_mm < lngMinW Then txtAtt.Text = lngMinW / Twip_mm
                
                ObjSel.Width = CDbl(txtAtt.Text) * Twip_mm * sgnMode
                Call SeekItem(ObjSel, ObjSel.Left, ObjSel.Top)
                
                mshAtt.TextMatrix(mshAtt.Row, 1) = Format(txtAtt.Text, "0.00")
                objReport.Items("_" & intCurID).W = Format(ObjSel.Width / sgnMode, "0.00")
                
                '如果是表格,要调整网格线
                If InStr(1, "4,5", objReport.Items("_" & intCurID).类型) <> 0 Then
                    Call SetGridLine(intCurID)
                End If
                If objReport.Items("_" & intCurID).类型 = 4 Then
                    Call SetCopyGrid(intCurID)
                End If
                
                txtAtt.Visible = False: mshAtt.SetFocus
                
                Select Case objReport.Items("_" & ObjSel.Index).类型
                Case 2
                    Call AdjustAll(True)
                Case 4, 5
                    Call SetMainWH(ObjSel.Index)
                End Select
                BlnSave = False
            Case "高度"
                If Not IsNumeric(txtAtt.Text) Then
                    MsgBox "请输入数字型数据！", vbInformation, App.Title
                    txtAtt.SetFocus: Exit Sub
                End If
                txtAtt.Text = Abs(CDbl(txtAtt.Text))
                If CDbl(txtAtt.Text) > 5000 Then txtAtt.Text = 5000
                
                '检查最小高度
                If objReport.Items("_" & intCurID).类型 = 4 Then
                    lngMinH = ObjSel.RowHeight(0) * (ObjSel.FixedRows + 2) + 15
                ElseIf objReport.Items("_" & intCurID).类型 = 5 Then
                    yy = ObjSel.FixedRows - 1 '横向分类项目数
                    lngMinH = ObjSel.RowHeight(0) * (yy + 3) + 60
                ElseIf objReport.Items("_" & intCurID).类型 = 12 Then '@@@
                    lngMinH = Chart(0).Height
                End If
                If CDbl(txtAtt.Text) * Twip_mm < lngMinH Then txtAtt.Text = lngMinH / Twip_mm
                
                ObjSel.Height = CDbl(txtAtt.Text) * Twip_mm * sgnMode
                Call SeekItem(ObjSel, ObjSel.Left, ObjSel.Top)
                
                mshAtt.TextMatrix(mshAtt.Row, 1) = Format(txtAtt.Text, "0.00")
                objReport.Items("_" & intCurID).H = Format(ObjSel.Height / sgnMode, "0.00")
                
                '如果是表格,要调整网格线
                If InStr(1, "4,5", objReport.Items("_" & intCurID).类型) <> 0 Then
                    Call SetGridLine(intCurID)
                End If
                If objReport.Items("_" & intCurID).类型 = 4 Then
                    Call SetCopyGrid(intCurID)
                End If
                
                txtAtt.Visible = False: mshAtt.SetFocus
                
                Select Case objReport.Items("_" & ObjSel.Index).类型
                Case 2
                    Call AdjustAll(True)
                Case 4, 5
                    Call SetMainWH(ObjSel.Index)
                End Select
                
                '设置其子表与主表相关属性一致
                If objReport.Items("_" & intCurID).参照 = "" And objReport.Items("_" & intCurID).类型 = 5 Then
                    For Each tmpItem In objReport.Items
                        If tmpItem.格式号 = mbytCurrFmt And tmpItem.参照 = objReport.Items("_" & intCurID).名称 And tmpItem.类型 = 5 Then
                            Call SetGridLike(msh(intCurID), msh(tmpItem.Key))
                        End If
                    Next
                End If
                BlnSave = False
            Case "行高"
                On Error Resume Next
                If Not IsNumeric(txtAtt.Text) Then
                    MsgBox "请输入数字型数据！", vbInformation, App.Title
                    txtAtt.SetFocus: Exit Sub
                End If
                txtAtt.Text = Abs(CDbl(txtAtt.Text))
                If CDbl(txtAtt.Text) > 5000 Then txtAtt.Text = 5000
                
                '所要设置的行高必须保证能完整显示所有固定行数+2
                '如果当前选中的是固定行,则设置当前固定行的高度;否则设置所有数据行高度
                PicFontTest.FontName = objReport.Items("_" & intCurID).字体
                PicFontTest.FontSize = objReport.Items("_" & intCurID).字号
                sgnH = (PicFontTest.TextHeight("字") + 15) * sgnMode
                Dim SgnFixedRows As Single
                
                If ObjSel.Row >= ObjSel.FixedRows Then
                    If objReport.Items("_" & intCurID).类型 = 4 Then
                        SgnFixedRows = 0
                        For i = 0 To ObjSel.FixedRows - 1
                            SgnFixedRows = SgnFixedRows + ObjSel.RowHeight(i)
                        Next
                        If Abs(Int(-CDbl(txtAtt.Text) * Twip_mm * sgnMode)) < sgnH Then
                            If objReport.Items("_" & intCurID).类型 = 5 Then ObjSel.RowHeightMin = sgnH
                            For i = ObjSel.FixedRows To ObjSel.Rows - 1
                                ObjSel.RowHeight(i) = sgnH
                            Next
                        ElseIf Abs(Int((-ObjSel.Height + SgnFixedRows) / Abs(Int(-CDbl(txtAtt.Text) * Twip_mm * sgnMode)))) > 2 Then
                            If objReport.Items("_" & intCurID).类型 = 5 Then ObjSel.RowHeightMin = Abs(Int(-ObjSel.Height / (ObjSel.FixedRows + 2)))
                            For i = ObjSel.FixedRows To ObjSel.Rows - 1
                                ObjSel.RowHeight(i) = Abs(Int(-CDbl(txtAtt.Text) * Twip_mm * sgnMode))
                            Next
                        End If
                        mshAtt.TextMatrix(mshAtt.Row, 1) = Format(txtAtt.Text, "0.00")
                        objReport.Items("_" & intCurID).行高 = Format(ObjSel.RowHeight(ObjSel.Row) / sgnMode, "0.00")
                    Else
                        If Abs(Int(-CDbl(txtAtt.Text) * Twip_mm * sgnMode)) < sgnH Then
                            For i = 0 To ObjSel.Rows - 1
                                ObjSel.RowHeight(i) = sgnH
                            Next
                        ElseIf Abs(Int(-ObjSel.Height / Abs(Int(-CDbl(txtAtt.Text) * Twip_mm * sgnMode)))) < ObjSel.FixedRows + 2 Then
                            For i = 0 To ObjSel.Rows - 1
                                ObjSel.RowHeight(i) = Abs(Int(-ObjSel.Height / (ObjSel.FixedRows + 2)))
                            Next
                        Else
                            For i = 0 To ObjSel.Rows - 1
                                ObjSel.RowHeight(i) = Abs(Int(-CDbl(txtAtt.Text) * Twip_mm * sgnMode))
                            Next
                        End If
                        mshAtt.TextMatrix(mshAtt.Row, 1) = Format(txtAtt.Text, "0.00")
                        objReport.Items("_" & intCurID).行高 = Format(ObjSel.RowHeight(0) / sgnMode, "0.00")
                    End If
                Else
                    '固定行
                    If Abs(Int(-CDbl(txtAtt.Text) * Twip_mm * sgnMode)) / (selCell.Row2 - selCell.Row1 + 1) < sgnH Then
                        If objReport.Items("_" & intCurID).类型 = 5 Then ObjSel.RowHeightMin = sgnH
                        For i = selCell.Row1 To selCell.Row2
                            ObjSel.RowHeight(i) = sgnH
                        Next
                    Else
                        If objReport.Items("_" & intCurID).类型 = 5 Then ObjSel.RowHeightMin = sgnH
                        For i = selCell.Row1 To selCell.Row2
                            ObjSel.RowHeight(i) = Abs(Int(-CDbl(txtAtt.Text) * Twip_mm * sgnMode)) / (selCell.Row2 - selCell.Row1 + 1)
                        Next
                    End If
                    '修改固定行的行高
                    For Each tmpID In objReport.Items("_" & ObjSel.Index).SubIDs
                        Set tmpItem = objReport.Items("_" & tmpID.id)
                        arrHead = Split(tmpItem.表头, "|")
                        tmpItem.表头 = ""
                        For i = 0 To UBound(arrHead)
                            If i >= selCell.Row1 And i <= selCell.Row2 Then
                                arrModify = Split(arrHead(i), "^")
                                tmpItem.表头 = tmpItem.表头 & "|" & arrModify(0) & "^" & ObjSel.RowHeight(i) / sgnMode & "^" & arrModify(2)
                            Else
                                tmpItem.表头 = tmpItem.表头 & "|" & arrHead(i)
                            End If
                        Next
                        tmpItem.表头 = Mid(tmpItem.表头, 2)
                    Next
                    mshAtt.TextMatrix(mshAtt.Row, 1) = txtAtt.Text
                End If
                
                Call SetGridLine(intCurID) '表格要调整网格线
                If objReport.Items("_" & intCurID).类型 = 4 Then Call SetCopyGrid(intCurID)
                txtAtt.Visible = False: mshAtt.SetFocus
                
                If GetSelNum = 1 Then ShowAttrib (intCurID)
                SetMainWH (ObjSel.Index)
                
                '设置其子表与主表相关属性一致
                If objReport.Items("_" & intCurID).参照 = "" And objReport.Items("_" & intCurID).类型 = 5 Then
                    For Each tmpItem In objReport.Items
                        If tmpItem.格式号 = mbytCurrFmt And tmpItem.参照 = objReport.Items("_" & intCurID).名称 And tmpItem.类型 = 5 Then
                            Call SetGridLike(msh(intCurID), msh(tmpItem.Key))
                        End If
                    Next
                End If
                BlnSave = False
            Case "分栏"
                If Not IsNumeric(txtAtt.Text) Then
                    MsgBox "请输入数字型数据！", vbInformation, App.Title
                    txtAtt.SetFocus: Exit Sub
                End If
                If CDbl(txtAtt.Text) > 20 Then
                    MsgBox "输入的数值过大！", vbInformation, App.Title
                    txtAtt.SetFocus: Exit Sub
                End If
                If objReport.Items("_" & intCurID).父ID <> 0 And txtAtt.Text <> "1" Then
                    MsgBox "卡片内的表格不允许分栏！", vbInformation, App.Title
                    txtAtt.SetFocus: txtAtt.Text = "1": Exit Sub
                End If
                txtAtt.Text = Abs(CDbl(txtAtt.Text))
                
                objReport.Items("_" & intCurID).分栏 = IIF(CByte(txtAtt.Text) < 2, 1, CByte(txtAtt.Text))
                
                mshAtt.TextMatrix(mshAtt.Row, 1) = objReport.Items("_" & intCurID).分栏
                
                '删除原有分栏
                For Each tmpID In objReport.Items("_" & intCurID).CopyIDs
                    Unload msh(tmpID.id)
                Next
                Set objReport.Items("_" & intCurID).CopyIDs = New RelatIDs
                
                '创建分栏
                For i = 1 To objReport.Items("_" & intCurID).分栏 - 1
                    intMaxID = intMaxID + 1
                    Load msh(intMaxID)
                    msh(intMaxID).Tag = "C_" & intCurID
                    msh(intMaxID).ToolTipText = "第 " & i & " 栏"
                    
                    msh(intMaxID).Top = msh(intCurID).Top
                    msh(intMaxID).Left = msh(intCurID).Left + (msh(intCurID).Width - 15) * i
                    msh(intMaxID).Width = msh(intCurID).Width
                    msh(intMaxID).Height = msh(intCurID).Height
                    
                    Call SetGridSame(msh(intCurID), msh(intMaxID))
                    
                    msh(intMaxID).ZOrder
                    msh(intMaxID).Visible = True
                    objReport.Items("_" & intCurID).CopyIDs.Add intMaxID, "_" & intMaxID
                Next
                msh(intCurID).ZOrder
                lblSize(Mid(msh(intCurID).Tag, 3) + 2).ZOrder
                
                '调整标签(只会是标签与其参照)因分栏
                Dim ResizeItem As RPTItem, IntSaveCurID As Integer
                IntSaveCurID = intCurID
                For Each ResizeItem In objReport.Items
                    If ResizeItem.类型 = 2 And ResizeItem.参照 = objReport.Items("_" & IntSaveCurID).名称 And ResizeItem.格式号 = mbytCurrFmt Then
                        intCurID = ResizeItem.Key
                        Call ReferTo
                    End If
                Next
                intCurID = IntSaveCurID
                
                txtAtt.Visible = False: mshAtt.SetFocus
                BlnSave = False
            Case "行距"
                If Not IsNumeric(txtAtt.Text) Then
                    MsgBox "请输入数字型数据！", vbInformation, App.Title
                    txtAtt.SetFocus: Exit Sub
                End If
                If CDbl(txtAtt.Text) > 100 Then
                    MsgBox "输入的数值过大！", vbInformation, App.Title
                    txtAtt.SetFocus: Exit Sub
                End If
                txtAtt.Text = Abs(CDbl(txtAtt.Text))
                
                objReport.Items("_" & intCurID).网格 = CByte(txtAtt.Text)
                
                mshAtt.TextMatrix(mshAtt.Row, 1) = objReport.Items("_" & intCurID).网格
                
                txtAtt.Visible = False: mshAtt.SetFocus
                BlnSave = False
            Case "格式"
                If TLen(txtAtt.Text) > 50 Then
                    MsgBox "格式中长度不能超过50个字符！", vbInformation, App.Title
                    txtAtt.SetFocus: Exit Sub
                End If
                mshAtt.TextMatrix(mshAtt.Row, 1) = txtAtt.Text
                objReport.Items("_" & intCurID).格式 = txtAtt.Text
                txtAtt.Visible = False: mshAtt.SetFocus
                BlnSave = False
            Case "上下间距"
                If Not IsNumeric(txtAtt.Text) Then
                    MsgBox "请输入数字型数据！", vbInformation, App.Title
                    txtAtt.SetFocus: Exit Sub
                End If
                txtAtt.Text = Abs(CDbl(txtAtt.Text))
                
                objReport.Items("_" & intCurID).上下间距 = Val(txtAtt.Text) * Twip_mm
                mshAtt.TextMatrix(mshAtt.Row, 1) = Format(objReport.Items("_" & intCurID).上下间距 / Twip_mm, "0.00")
                
                txtAtt.Visible = False: mshAtt.SetFocus
                BlnSave = False
            Case "左右间距"
                If Not IsNumeric(txtAtt.Text) Then
                    MsgBox "请输入数字型数据！", vbInformation, App.Title
                    txtAtt.SetFocus: Exit Sub
                End If
                txtAtt.Text = Abs(CDbl(txtAtt.Text))
                
                objReport.Items("_" & intCurID).左右间距 = Val(txtAtt.Text) * Twip_mm
                mshAtt.TextMatrix(mshAtt.Row, 1) = Format(objReport.Items("_" & intCurID).左右间距 / Twip_mm, "0.00")
                
                txtAtt.Visible = False: mshAtt.SetFocus
                BlnSave = False
            Case "横向分栏"
                If Not IsNumeric(txtAtt.Text) Then
                    MsgBox "请输入数字型数据！", vbInformation, App.Title
                    txtAtt.SetFocus: Exit Sub
                End If
                
                txtAtt.Text = Abs(CDbl(txtAtt.Text))
                If Val(txtAtt.Text) > 0 Then
                    lngL = (objReport.Fmts("_" & mbytCurrFmt).W - objReport.Items("_" & intCurID).X + objReport.Items("_" & intCurID).左右间距) \ (objReport.Items("_" & intCurID).W + objReport.Items("_" & intCurID).左右间距)
                    If Val(txtAtt.Text) > lngL Then
                        MsgBox "根据左右间距，纸张横向最多分" & lngL & "栏。", vbInformation, App.Title
                        txtAtt.SetFocus: Exit Sub
                    End If
                End If
                
                objReport.Items("_" & intCurID).横向分栏 = Val(txtAtt.Text)
                mshAtt.TextMatrix(mshAtt.Row, 1) = objReport.Items("_" & intCurID).横向分栏
                
                txtAtt.Visible = False: mshAtt.SetFocus
                BlnSave = False
            Case "纵向分栏"
                If Not IsNumeric(txtAtt.Text) Then
                    MsgBox "请输入数字型数据！", vbInformation, App.Title
                    txtAtt.SetFocus: Exit Sub
                End If
                
                txtAtt.Text = Abs(CDbl(txtAtt.Text))
                If Val(txtAtt.Text) > 0 Then
                    lngW = (objReport.Fmts("_" & mbytCurrFmt).H - objReport.Items("_" & intCurID).Y + objReport.Items("_" & intCurID).上下间距) \ (objReport.Items("_" & intCurID).H + objReport.Items("_" & intCurID).上下间距)
                    If Val(txtAtt.Text) > lngW Then
                        MsgBox "根据上下间距，纸张纵向最多分" & lngW & "栏。", vbInformation, App.Title
                        txtAtt.SetFocus: Exit Sub
                    End If
                End If
                
                objReport.Items("_" & intCurID).纵向分栏 = Val(txtAtt.Text)
                mshAtt.TextMatrix(mshAtt.Row, 1) = objReport.Items("_" & intCurID).纵向分栏
                
                txtAtt.Visible = False: mshAtt.SetFocus
                BlnSave = False
            
            Case "数据源行号"
                If Not IsNumeric(txtAtt.Text) Then
                    MsgBox "请输入数字型数据！", vbInformation, App.Title
                    txtAtt.SetFocus: Exit Sub
                End If
                txtAtt.Text = Abs(CDbl(txtAtt.Text))
                
                objReport.Items("_" & intCurID).源行号 = Val(txtAtt.Text)
                mshAtt.TextMatrix(mshAtt.Row, 1) = objReport.Items("_" & intCurID).源行号
                
                txtAtt.Visible = False: mshAtt.SetFocus
                BlnSave = False
            Case "关联报表"
                X = InStr(1, objReport.Items("_" & intCurID).内容, "]")
                Y = InStr(1, objReport.Items("_" & intCurID).内容, ".")
                k = InStr(1, objReport.Items("_" & intCurID).内容, "[")
                If X > k And X > Y And X <> 0 And k <> 0 Then
                    strReportID = FindReport(txtAtt.Text, txtAtt.hwnd, strInfo, objReport.Items("_" & intCurID).Relations.Item(1).关联报表ID, objReport, objReport.Items("_" & intCurID).Relations, 2, Me, intCurID)
                    If strReportID <> "" Then
                        mshAtt.TextMatrix(mshAtt.Row, 1) = strInfo
                        mshAtt.RowData(mshAtt.Row) = strReportID
                        txtAtt.Visible = False: mshAtt.SetFocus
                        BlnSave = False
                    Else
                        txtAtt.SetFocus
                    End If
                Else
                    MsgBox "当前标签必须先绑定一个数据源，例如：[数据源.字段],绑定后再设置关联报表。", vbInformation, Me.Caption
                End If
        End Select
        Call AdjustAll
    Else
        Select Case mshAtt.TextMatrix(mshAtt.Row, 0)
            Case "内容"
                If InStr("'|~^", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
            Case "X坐标", "Y坐标", "宽度", "高度", "行高", "上下间距", "左右间距"
                If InStr("-0.123456789" & Chr(8) & Chr(3) & Chr(22) & Chr(24) & Chr(26), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
            Case "分栏", "行距", "数据源行号", "纵向分栏", "横向分栏"
                If InStr("0123456789" & Chr(8) & Chr(3) & Chr(22) & Chr(24) & Chr(26), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
            Case "格式"
                If InStr("'|~^", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
        End Select
    End If
End Sub

Private Function GetRow(str As String) As Integer
    Dim i As Integer
    For i = 1 To mshAtt.Rows - 1
        If UCase(Trim(mshAtt.TextMatrix(i, 0))) = UCase(Trim(str)) Then
            GetRow = i: Exit Function
        End If
    Next
    GetRow = 0
End Function

Private Sub InitRowHeight(intIdx As Integer)
'功能：初始表格行高为255,以便于程序手动调整行高(不然会根据字体自动调整)
'说明：在Load一个新的表格后,一定要调用该过程
    Dim i As Integer
    For i = 0 To msh(intIdx).Rows - 1
        msh(intIdx).RowHeight(i) = 255 * sgnMode
    Next
End Sub

Private Sub SetGridLine(idx As Integer)
'功能：根据指定网格的现有字体,字体行高情况,重新填充网格线
'说明：调整时控件对应的数据对象(Item)必须已经存在，且对应控件已经建立好行列头框架
    Dim blnPre As Boolean, SinH As Single
    Dim X As Integer, Y As Integer, Z As Integer
    Dim tmpID As RelatID, i As Integer, j As Integer
    Dim intFixHeight As Integer

    blnPre = msh(idx).Redraw
    msh(idx).Redraw = False
    
    If objReport.Items("_" & idx).类型 = 4 Then '汇总表格
        '任意表格纵向填满表格线
        If objReport.票据 Then
            SinH = 0: X = msh(idx).FixedRows
            For i = 0 To msh(idx).FixedRows - 1
                SinH = SinH + msh(idx).RowHeight(i)
            Next
            msh(idx).Rows = Abs(Int((-(msh(idx).Height - SinH)) / (objReport.Items("_" & idx).行高 * sgnMode))) + X
            If msh(idx).Rows = X Then msh(idx).Rows = msh(idx).Rows + 2
            msh(idx).FixedRows = X
            For i = msh(idx).FixedRows To msh(idx).Rows - 1
                msh(idx).RowHeight(i) = objReport.Items("_" & idx).行高 * sgnMode
            Next
        End If
    ElseIf objReport.Items("_" & idx).类型 = 5 Then '汇总表格
        X = msh(idx).FixedCols '纵向分类项目数
        Y = msh(idx).FixedRows - 1 '横向分类项目数
        For Each tmpID In objReport.Items("_" & idx).SubIDs
            If objReport.Items("_" & tmpID.id).类型 = 9 Then Z = Z + 1 '统计项目数
        Next
        
        For i = 0 To msh(idx).FixedRows - 1
            intFixHeight = intFixHeight + msh(idx).RowHeight(i)
        Next
        '汇总表格纵向填满表格线
        msh(idx).Rows = Abs(Int(-(msh(idx).Height - intFixHeight) / msh(idx).RowHeight(msh(idx).FixedRows))) + 2
        If msh(idx).Rows < msh(idx).FixedRows + 3 Then msh(idx).Rows = msh(idx).FixedRows + 3
        
        For i = msh(idx).FixedRows + 1 To msh(idx).Rows - 1
'            msh(idx).RowHeight(i) = msh(idx).RowHeight(0)
            For j = 0 To msh(idx).FixedCols - 1
                msh(idx).TextMatrix(i, j) = msh(idx).TextMatrix(msh(idx).FixedRows + 2, j)
            Next
        Next
        
        '如果有横向分类：汇总表格横向填满统计项
        If msh(idx).FixedRows > 1 Then
            X = 0
            For i = 0 To msh(idx).FixedCols - 1
                X = X + msh(idx).ColWidth(i) '纵向分类总宽度
            Next
            Y = 0
            For i = msh(idx).FixedCols To msh(idx).FixedCols + Z - 1
                Y = Y + msh(idx).ColWidth(i) '一组统计项总宽度
            Next
            '列数 = 统计组数 * 每组列数 + 纵向分类项数
            msh(idx).Cols = Abs(Int(-(msh(idx).Width - X) / Y)) * Z + msh(idx).FixedCols
            '每组宽度及标题相同
            For i = msh(idx).FixedCols + Z To msh(idx).Cols - 1
                For j = 0 To msh(idx).FixedRows - 2
                    msh(idx).TextMatrix(j, i) = msh(idx).TextMatrix(j, msh(idx).FixedCols + 1)
                Next
            Next
            For i = msh(idx).FixedCols + Z To msh(idx).Cols - 1 Step Z
                For j = 1 To Z
                    msh(idx).TextMatrix(msh(idx).FixedRows - 1, i + j - 1) = _
                    msh(idx).TextMatrix(msh(idx).FixedRows - 1, msh(idx).FixedCols + j - 1)
                    msh(idx).ColWidth(i + j - 1) = msh(idx).ColWidth(msh(idx).FixedCols + j - 1)
                    msh(idx).ColAlignment(i + j - 1) = msh(idx).ColAlignment(msh(idx).FixedCols + j - 1)
                Next
            Next
        End If
    End If
    
    msh(idx).Redraw = blnPre
End Sub

Private Sub ReplaceName(strPre As String, strNew As String)
'功能：将所有涉及strPre数据源名称的报表元素的相应内容替换成新数据源名称strNew
    Dim tmpItem As RPTItem
    Dim i As Integer, j As Integer
    
    If strPre = strNew Then Exit Sub
    
    For Each tmpItem In objReport.Items
        Select Case tmpItem.类型
            Case 2, 3, 11, 14 '标签
                If InStr(tmpItem.内容, strPre & ".") > 0 Then
                    tmpItem.内容 = Replace(tmpItem.内容, strPre & ".", strNew & ".")
                    If tmpItem.类型 <> 11 And tmpItem.格式号 = mbytCurrFmt Then lbl(tmpItem.id).Caption = Replace(lbl(tmpItem.id).Caption, strPre & ".", strNew & ".")
                End If
            Case 4 '任意表格
                If tmpItem.格式号 = mbytCurrFmt Then
                    For j = 0 To msh(tmpItem.id).Rows - 1
                        For i = 0 To msh(tmpItem.id).Cols - 1
                            If InStr(msh(tmpItem.id).TextMatrix(j, i), strPre & ".") > 0 Then
                                msh(tmpItem.id).TextMatrix(j, i) = _
                                    Replace(msh(tmpItem.id).TextMatrix(j, i), strPre & ".", strNew & ".")
                            End If
                        Next
                    Next
                End If
                
                For i = 1 To tmpItem.SubIDs.count
                    objReport.Items("_" & tmpItem.SubIDs(i).Key).表头 = Replace(objReport.Items("_" & tmpItem.SubIDs(i).Key).表头, strPre & ".", strNew & ".")
                    objReport.Items("_" & tmpItem.SubIDs(i).Key).内容 = Replace(objReport.Items("_" & tmpItem.SubIDs(i).Key).内容, strPre & ".", strNew & ".")
                Next
                If tmpItem.格式号 = mbytCurrFmt Then
                    Call SetCopyGrid(tmpItem.id)
                End If
            Case 5 '汇总表格
                If tmpItem.内容 = strPre Then
                    tmpItem.内容 = strNew
                End If
            Case 6 '任意表格列
                If InStr(tmpItem.内容, strPre & ".") > 0 Then
                    tmpItem.内容 = Replace(tmpItem.内容, strPre & ".", strNew & ".")
                End If
            Case 12
                If tmpItem.内容 <> "" Then
                    tmpItem.内容 = Mid(Replace("|" & tmpItem.内容, "|" & strPre & ".", "|" & strNew & "."), 2)
                End If
        End Select
    Next
End Sub

Private Function CheckData() As String
'功能：检查元素中的数据元素是否能找到对应的数据源
'返回：提示信息
    Dim tmpItem As RPTItem, objNode As Object
    Dim strFX As String, strFS As String, strFY As String, strDBConn As String
    Dim strData As String, blnExist As Boolean
    Dim lngL As Long, lngW As Long
    
    For Each tmpItem In objReport.Items
        blnExist = False
        Select Case tmpItem.类型
            Case 2, 3 '标签
                '只在末级结点(数据源="[数据源名.字段]"
                blnExist = CheckText(tmpItem)
                If Not blnExist Then CheckData = "在数据源中找不到数据标签的内容""" & tmpItem.内容 & """,请处理数据源或标签！": Exit Function
            Case 5  '汇总表格(数据源名)
                '只在中间结点
                blnExist = CheckText(tmpItem)
                If Not blnExist Then CheckData = "在数据源中找不到汇总表格的内容""" & tmpItem.内容 & """,请处理数据源或表格！": Exit Function
            Case 6 '任意表格子项
                '只在末级结点(数据内容中可能有公式)
                If GetItemCount(tmpItem.内容) > 0 Then
                    blnExist = CheckText(tmpItem)
                    If Not blnExist Then CheckData = "在数据源中找不到任意表格列的内容""" & tmpItem.内容 & """,请处理数据源或表格！": Exit Function
                End If
            Case 7, 8, 9 '汇总表格子项(只存了字段名)
                '只在末级结点
                For Each objNode In tvwSQL.Nodes
                    If objNode.Key <> "Root" And objNode.Children = 0 Then
                        strDBConn = objReport.Items("_" & tmpItem.上级ID).内容
                        If strDBConn Like "*（*）" Then
                            strDBConn = Left(strDBConn, InStrRev(strDBConn, "（") - 1)
                        End If
                        If LevelText(objNode) = strDBConn & "." & tmpItem.内容 Then blnExist = True: Exit For
                    End If
                Next
                If Not blnExist Then CheckData = "在数据源中找不到汇总表格数据项的内容""" & tmpItem.内容 & """,请处理数据源或表格！": Exit Function
                
                If tmpItem.排序 <> "" Then
                    blnExist = False
                    For Each objNode In tvwSQL.Nodes
                        If objNode.Key <> "Root" And objNode.Children = 0 Then
                            If LevelText(objNode) = mdlPublic.GetStdNodeText(objReport.Items("_" & tmpItem.上级ID).内容) & "." & _
                                IIF(Left(tmpItem.排序, 1) = ",", Mid(tmpItem.排序, 2), tmpItem.排序) Then blnExist = True: Exit For
                        End If
                    Next
                    If Not blnExist Then CheckData = "在数据源中找不到汇总表格排序项的内容""" & IIF(Left(tmpItem.排序, 1) = ",", Mid(tmpItem.排序, 2), tmpItem.排序) & """,请处理数据源或表格！": Exit Function
                End If
            Case 12 '@@@
                If tmpItem.内容 <> "" Then
                    Call GetChartDataName(tmpItem.内容, strFX, strFS, strFY, strData)
                    If strFX <> "" Then
                        blnExist = False
                        For Each objNode In tvwSQL.Nodes
                            If objNode.Key <> "Root" And objNode.Children = 0 Then
                                If LevelText(objNode) = strData & "." & strFX Then blnExist = True: Exit For
                            End If
                        Next
                        If Not blnExist Then CheckData = "在数据源中找不到图表""" & tmpItem.名称 & """的Ｘ值字段，请处理数据源或图表。": Exit Function
                    End If
                    
                    If strFS <> "" Then
                        blnExist = False
                        For Each objNode In tvwSQL.Nodes
                            If objNode.Key <> "Root" And objNode.Children = 0 Then
                                If LevelText(objNode) = strData & "." & strFS Then blnExist = True: Exit For
                            End If
                        Next
                        If Not blnExist Then CheckData = "在数据源中找不到图表""" & tmpItem.名称 & """的序列字段，请处理数据源或图表。": Exit Function
                    End If
                    
                    If strFY <> "" Then
                        blnExist = False
                        For Each objNode In tvwSQL.Nodes
                            If objNode.Key <> "Root" And objNode.Children = 0 Then
                                If LevelText(objNode) = strData & "." & strFY Then blnExist = True: Exit For
                            End If
                        Next
                        If Not blnExist Then CheckData = "在数据源中找不到图表""" & tmpItem.名称 & """的Ｙ值字段，请处理数据源或图表。": Exit Function
                    End If
                End If
            Case 13 '条码
                '只在末级结点(数据源="[数据源名.字段]"
                blnExist = CheckText(tmpItem)
                If Not blnExist Then CheckData = "在数据源中找不到条码的数据""" & tmpItem.内容 & """,请处理数据源或条码内容！": Exit Function
            Case 14 '卡片
                If tmpItem.纵向分栏 > 0 Then
                    lngW = (objReport.Fmts("_" & mbytCurrFmt).H - tmpItem.Y + tmpItem.上下间距) \ (tmpItem.H + tmpItem.上下间距)
                    If tmpItem.纵向分栏 > lngW Then
                        CheckData = "根据上下间距，纸张纵向最多分" & lngW & "栏，请检查卡片属性。"
                        Exit Function
                    End If
                End If
                If tmpItem.横向分栏 > 0 Then
                    lngL = (objReport.Fmts("_" & mbytCurrFmt).W - tmpItem.X + tmpItem.左右间距) \ (tmpItem.W + tmpItem.左右间距)
                    If tmpItem.横向分栏 > lngL Then
                        CheckData = "根据左右间距，纸张横向最多分" & lngL & "栏，请检查卡片属性。"
                        Exit Function
                    End If
                End If
        End Select
    Next
End Function

Private Function CheckArea() As String
'功能：检查是否所有元素都在报表宽高范围之内,以及任意表格列是否可以显示完。
    Dim tmpItem As RPTItem, bytFmt As Byte
    Dim StrFmt As String, objFmt As RPTFmt
    Dim lngW As Long, lngH As Long
    Dim strTmp As String
    
    Call ReFlashWidth
    
    For Each tmpItem In objReport.Items
        With tmpItem
            If InStr(1, "|1,|2,|3,|4,|5,|10,|11,|12,|14,", "|" & .类型 & ",") <> 0 Then
                Set objFmt = objReport.Fmts("_" & .格式号)
                If tmpItem.父ID = 0 Then
                    If objFmt.纸向 = 1 Then
                        lngW = objFmt.W
                        lngH = objFmt.H
                    Else
                        lngW = objFmt.H
                        lngH = objFmt.W
                    End If
                    strTmp = "纸张"
                Else
                    lngW = objReport.Items("_" & tmpItem.父ID).W
                    lngH = objReport.Items("_" & tmpItem.父ID).H
                    strTmp = "卡片"
                End If
                StrFmt = objFmt.说明
                If .X < 0 Or .Y < 0 Or (.X + .W) > lngW Or (.Y + .H) > lngH Then
                    
                    Select Case .类型
                        Case 1
                            CheckArea = "格式[" & StrFmt & "]中某个线条的位置超出了" & strTmp & "的尺寸范围,请作调整！"
                        Case 2, 3
                            CheckArea = "格式[" & StrFmt & "]中某个标签的位置超出了" & strTmp & "的尺寸范围,请作调整！"
                        Case 4
                            CheckArea = "格式[" & StrFmt & "]中某个任意表格的位置超出了" & strTmp & "的尺寸范围,请作调整！"
                        Case 5
                            CheckArea = "格式[" & StrFmt & "]中某个汇总表格的位置超出了" & strTmp & "的尺寸范围,请作调整！"
                        Case 10
                            CheckArea = "格式[" & StrFmt & "]中某个框线的位置超出了" & strTmp & "的尺寸范围,请作调整！"
                        Case 11
                            CheckArea = "格式[" & StrFmt & "]中某个图片的位置超出了" & strTmp & "的尺寸范围,请作调整！"
                        Case 12
                            CheckArea = "格式[" & StrFmt & "]中某个图表的位置超出了" & strTmp & "的尺寸范围,请作调整！"
                        Case 14
                            CheckArea = "格式[" & StrFmt & "]中某个卡片的位置超出了" & strTmp & "的尺寸范围,请作调整！"
                    End Select
                    Exit Function
                End If
                If .分栏 > 1 And .类型 = 4 Then
                    If .X + .W * .分栏 > lngW Then
                        CheckArea = "报表中某个任意表格的分栏过多,超出了" & strTmp & "的尺寸范围,请作调整！"
                        Exit Function
                    End If
                End If
            End If
        End With
    Next
End Function

Private Sub ShowItems()
'功能：根据objReport对象显示报表元素
    Dim tmpItem As RPTItem, bytFormat As Byte
    
    '先显示图表以减少闪铄
    For Each tmpItem In objReport.Items
        If tmpItem.格式号 = mbytCurrFmt And tmpItem.类型 = 12 Then
            Call ShowItem(tmpItem.id)
        End If
    Next

    For Each tmpItem In objReport.Items
        If tmpItem.格式号 = mbytCurrFmt And tmpItem.类型 <> 12 Then
            Call ShowItem(tmpItem.id)
        End If
    Next
End Sub

Private Sub LoadReportFormat()
    Dim tmpFmt As RPTFmt
    Dim objItem As ComboItem
    
    With cboFormat
        blnAllowIn = False
        .ComboItems.Clear
        
        For Each tmpFmt In objReport.Fmts
            Set objItem = .ComboItems.Add(, "_" & tmpFmt.序号, tmpFmt.说明, "Root")
            If tmpFmt.序号 = mbytCurrFmt Then
                objItem.Selected = True
                Set .SelectedItem = objItem
                .SelectedItem.Selected = True
            End If
        Next
        If .ComboItems.count > 0 And .SelectedItem Is Nothing Then
            .ComboItems(1).Selected = True
            Set .SelectedItem = .ComboItems(1)
            .SelectedItem.Selected = True
        End If
        mbytCurrFmt = Mid(cboFormat.SelectedItem.Key, 2)
        
        cmdAdd.Enabled = blnDelReportFormat
        cmdDel.Enabled = (.ComboItems.count > 1) And blnDelReportFormat
        tbr1.Buttons("AddFormat").Enabled = cmdAdd.Enabled
        tbr1.Buttons("DelFormat").Enabled = cmdDel.Enabled
        mnuEdit_AddFormat.Enabled = cmdAdd.Enabled
        mnuEdit_DelFormat.Enabled = cmdDel.Enabled
        
        blnAllowIn = True
    End With
End Sub

Private Sub ShowItem(idx As Integer)
'功能：显示指定的报表元素(ShowItems的子函数,也可单独调用)
'参数：idx=objReport中的元素索引
    Dim i As Integer, j As Integer, tmpID As RelatID
    Dim ObjSel As Control
    Dim objBarCode As StdPicture, strBarCode As String
    Dim lngSize As Long, sngWidth As Single
    
    With objReport.Items("_" & idx)
        Select Case .类型
            Case 1 '线条
                Load lblLine(.id)
                Set ObjSel = lblLine(.id)
                ObjSel.Top = Format(.Y * sgnMode, "0.00")
                ObjSel.Left = Format(.X * sgnMode, "0.00")
                ObjSel.Height = Format(.H * sgnMode, "0.00")
                ObjSel.Width = Format(.W * sgnMode, "0.00")
                ObjSel.BackColor = .前景
                If .粗体 Then ObjSel.Height = 30
                ObjSel.ZOrder
                ObjSel.Visible = True
            Case 2, 3 '标签
                Load lbl(.id)
                Set ObjSel = lbl(.id)
                ObjSel.Top = Format(.Y * sgnMode, "0.00")
                ObjSel.Left = Format(.X * sgnMode, "0.00")
                ObjSel.Height = Format(.H * sgnMode, "0.00")
                ObjSel.Width = Format(.W * sgnMode, "0.00")
                ObjSel.ForeColor = .前景
                ObjSel.BackColor = IIF(.背景 = &HFFFFFF, lbl(0).BackColor, .背景)
                ObjSel.Font.name = .字体
                ObjSel.Font.Size = Format(.字号 * sgnMode, "0.0")
                ObjSel.Font.Bold = .粗体
                ObjSel.Font.Italic = .斜体
                ObjSel.Font.Underline = .下线
                ObjSel.BorderStyle = IIF(.边框, 1, 0)
                ObjSel.Alignment = IIF(.对齐 <> 0, IIF(.对齐 = 1, 2, 1), 0)
                ObjSel.Caption = .内容
                
                If Not ItemIsGraph(.id) Then
                    ObjSel.AutoSize = .自调
                End If
                
                If InStr(1, "|11,", "|" & .类型 & ",") <> 0 Then
                    ObjSel.BorderStyle = 1
                    ObjSel.BackStyle = 0
                    If .类型 = 10 Then ObjSel.Caption = ""
                    
                    Call DrawFrame(ObjSel)
                End If
                ObjSel.ZOrder 0
                ObjSel.Visible = True
            Case 10 '框线
                Load Shp(.id)
                Set ObjSel = Shp(.id)
                Load lblshp(.id)
                lblshp(.id).BackColor = picPaper.BackColor
                ObjSel.Top = Format(.Y * sgnMode, "0.00")
                ObjSel.Left = Format(.X * sgnMode, "0.00")
                ObjSel.Height = Format(.H * sgnMode, "0.00")
                ObjSel.Width = Format(.W * sgnMode, "0.00")
                lblshp(.id).Top = ObjSel.Top
                lblshp(.id).Left = ObjSel.Left
                lblshp(.id).Width = ObjSel.Width
                lblshp(.id).Height = ObjSel.Height
                ObjSel.BorderColor = .前景
                ObjSel.BackColor = IIF(.背景 = &HFFFFFF, Shp(0).BackColor, .背景)
                ObjSel.BorderStyle = 1
                ObjSel.BackStyle = 0
                ObjSel.BorderWidth = IIF(.粗体, 2, 1)
                ObjSel.Shape = IIF(.边框, ShapeConstants.vbShapeOval, ShapeConstants.vbShapeRectangle)
                
                ObjSel.ZOrder 1
                ObjSel.Visible = True
                lblshp(.id).ZOrder 1
                lblshp(.id).Visible = True
            Case 4, 5 '任意表格,汇总表格
                Load msh(.id)
                Set ObjSel = msh(.id)
                '格式设置
                ObjSel.Top = Format(.Y * sgnMode, "0.00")
                ObjSel.Left = Format(.X * sgnMode, "0.00")
                ObjSel.Height = Format(.H * sgnMode, "0.00")
                ObjSel.Width = Format(.W * sgnMode, "0.00")
                ObjSel.Font.Size = Format(.字号 * sgnMode, "0.0")
                '分栏设置(对象CopyIDs已经设置)
                i = 0
                For Each tmpID In .CopyIDs
                    i = i + 1
                    Load msh(tmpID.id)
                    msh(tmpID.id).Width = ObjSel.Width
                    msh(tmpID.id).Height = ObjSel.Height
                    msh(tmpID.id).Top = ObjSel.Top
                    msh(tmpID.id).Left = ObjSel.Left + (ObjSel.Width - 15) * i
                    msh(tmpID.id).Font.Size = ObjSel.Font.Size
                    msh(tmpID.id).Tag = "C_" & .id
                    msh(tmpID.id).ZOrder
                    msh(tmpID.id).Visible = True
                Next
                
                Call ReShowGrid(.id)
                If .类型 = 4 Then Call CustomColColor(.id, -9)
                ObjSel.ZOrder
                ObjSel.Visible = True
            Case 11
                Load img(.id)
                Set ObjSel = img(.id)
                
                ObjSel.Top = Format(.Y * sgnMode, "0.00")
                ObjSel.Left = Format(.X * sgnMode, "0.00")
                ObjSel.Height = Format(.H * sgnMode, "0.00")
                ObjSel.Width = Format(.W * sgnMode, "0.00")
                ObjSel.BorderStyle = IIF(.边框, 1, 0)
                
                '保持比例
                If Not .图片 Is Nothing Then
                    If .粗体 Then
                        Set ObjSel.Picture = ScalePicture(PicFontTest, .图片, ObjSel.Width, ObjSel.Height)
                    Else
                        Set ObjSel.Picture = .图片
                    End If
                End If
                
                ObjSel.ZOrder
                ObjSel.Visible = True
            Case 14
                Load pic(.id)
                Set ObjSel = pic(.id)
                
                ObjSel.Top = Format(.Y * sgnMode, "0.00")
                ObjSel.Left = Format(.X * sgnMode, "0.00")
                ObjSel.Height = Format(.H * sgnMode, "0.00")
                ObjSel.Width = Format(.W * sgnMode, "0.00")
                ObjSel.BorderStyle = IIF(.边框, 1, 0)
                
                ObjSel.ZOrder
                ObjSel.Visible = True
            Case 12 '@@@
                Load Chart(.id)
                Set ObjSel = Chart(.id)
                
                ObjSel.Top = Format(.Y * sgnMode, "0.00")
                ObjSel.Left = Format(.X * sgnMode, "0.00")
                ObjSel.Height = Format(.H * sgnMode, "0.00")
                ObjSel.Width = Format(.W * sgnMode, "0.00")
                
                Call SetChartStyleAndData(ObjSel, objReport.Items("_" & idx), , sgnMode, True)
                
                ObjSel.ZOrder
                ObjSel.Visible = True
            Case 13
                Load ImgCode(.id)
                Set ObjSel = ImgCode(.id)
                
                ObjSel.Top = Format(.Y * sgnMode, "0.00")
                ObjSel.Left = Format(.X * sgnMode, "0.00")
                ObjSel.Height = Format(.H * sgnMode, "0.00")
                ObjSel.Width = Format(.W * sgnMode, "0.00")
                ObjSel.BorderStyle = 0
                
                '显示条码图象
                strBarCode = ReplaceBracket(.内容)
                If strBarCode = "" Then strBarCode = "1234567890"
                
                Unload frmFlash '强制初始Picture，不然切换绘制有问题
                If .序号 = 1 Then
                    Set objBarCode = DrawBarCode128(frmFlash.picTemp, 3, strBarCode, Mid(.表头, 1, 1) = "1")
                ElseIf .序号 = 2 Then
                    Set objBarCode = DrawBarCode39(frmFlash.picTemp, 3, strBarCode, Mid(.表头, 2, 1) = "1", Mid(.表头, 1, 1) = "1")
                ElseIf .序号 = 3 Then
                    Set objBarCode = DrawBarCode128Auto(frmFlash.picTemp, strBarCode, sngWidth, .行高, Mid(.表头, 1, 1) = "1")
                ElseIf .序号 = 10 Then
                    Set objBarCode = DrawBarCode2D(strBarCode, frmFlash.picTemp, lngSize)
                End If
                If Val(Mid(.表头, 3, 1)) <> 0 Then
                    Set objBarCode = PictureSpin(objBarCode, Val(Mid(.表头, 3, 1)), frmFlash.picTemp)
                End If
                Set ObjSel.Picture = objBarCode
                
                If .序号 = 3 Then
                    '128码自动调整宽度
                    If Val(Mid(.表头, 3, 1)) = 0 Then
                        ObjSel.Width = Format(Me.ScaleX(sngWidth, vbMillimeters, vbTwips) * sgnMode, "0.00")
                        .W = Me.ScaleX(sngWidth, vbMillimeters, vbTwips)
                    Else
                        ObjSel.Height = Format(Me.ScaleY(sngWidth, vbMillimeters, vbTwips) * sgnMode, "0.00")
                        .H = Me.ScaleY(sngWidth, vbMillimeters, vbTwips)
                    End If
                ElseIf .序号 = 10 And .自调 Then
                    '二维条码缺省自动调整大小
                    ObjSel.Width = Format(lngSize * sgnMode, "0.00")
                    ObjSel.Height = Format(lngSize * sgnMode, "0.00")
                    .W = lngSize: .H = lngSize
                End If

                ObjSel.ZOrder
                ObjSel.Visible = True
        End Select
        If .父ID <> 0 And InStr(",14,5,6,12,", "," & .类型 & ",") = 0 Then
            Set ObjSel.Container = pic(.父ID)
        End If
    End With
End Sub

Private Sub ReShowGrid(idx As Integer)
'功能：根据objReport的内容重新绘制表格内容,可时刷新分栏控件
'说明：1.objReport对象内容已存在,2.对应控件已存在

    Dim i As Integer, j As Integer, X As Integer, Y As Integer, Z As Integer
    Dim tmpID As RelatID, tmpItem As RPTItem, strCaption As String, sgnH As Long
    
    msh(idx).Redraw = False
    msh(idx).Clear
    With objReport.Items("_" & idx)
        If .类型 = 4 Then '任意表格
            '格式设置(位置及尺寸不动)
            msh(idx).ForeColor = .前景
            msh(idx).ForeColorFixed = .前景
            msh(idx).GridColor = .网格
            msh(idx).GridColorFixed = IIF(.格式 = "", .网格, Val(.格式))
            
            msh(idx).BackColor = .背景
            msh(idx).BackColorFixed = IIF(.背景 = &HFFFFFF, lbl(0).BackColor, .背景)
            
            msh(idx).Font.name = .字体
            msh(idx).Font.Size = Format(.字号 * sgnMode, "0.0")
            msh(idx).Font.Bold = .粗体
            msh(idx).Font.Italic = .斜体
            msh(idx).Font.Underline = .下线
            msh(idx).GridLineWidth = IIF(.表格线加粗, 2, 1)
            '行列设置
            '列数
            msh(idx).Cols = .SubIDs.count
            msh(idx).FixedCols = 0
            i = 0
            For Each tmpID In .SubIDs
                Set tmpItem = objReport.Items("_" & tmpID.id)
                
                If i = 0 Then '最小行数
                    If objReport.票据 = False Then
                        msh(idx).Rows = UBound(Split(tmpItem.表头, "|")) + 3
                        msh(idx).FixedRows = UBound(Split(tmpItem.表头, "|")) + 1
                    Else
                        msh(idx).Rows = UBound(Split(tmpItem.表头, "|")) + 3
                        msh(idx).FixedRows = UBound(Split(tmpItem.表头, "|")) + 1
                    End If
                End If

                '数据列内容
                msh(idx).ColWidth(tmpItem.序号) = tmpItem.W * sgnMode
                msh(idx).ColAlignment(tmpItem.序号) = Switch(tmpItem.对齐 = 0, 1, tmpItem.对齐 = 1, 4, tmpItem.对齐 = 2, 7)
                msh(idx).TextMatrix(msh(idx).FixedRows, tmpItem.序号) = tmpItem.内容
                msh(idx).TextMatrix(msh(idx).FixedRows + 1, tmpItem.序号) = tmpItem.汇总
                                    
                '自定义表头内容
                For i = 0 To msh(idx).FixedRows - 1
                    On Error Resume Next
                    
                    Err = 0
                    strCaption = Split(Split(tmpItem.表头, "|")(i), "^")(2)
                    If Err <> 0 Then strCaption = ""
                    If strCaption = "#" Then
                        msh(idx).TextMatrix(i, tmpItem.序号) = ""
                    ElseIf strCaption = "←" Then
                        msh(idx).TextMatrix(i, tmpItem.序号) = msh(idx).TextMatrix(i, tmpItem.序号 - 1)
                    ElseIf strCaption = "↑" Then
                        msh(idx).TextMatrix(i, tmpItem.序号) = msh(idx).TextMatrix(i - 1, tmpItem.序号)
                    Else
                        msh(idx).TextMatrix(i, tmpItem.序号) = strCaption
                    End If
                    
                    Err = 0
                    sgnH = Split(Split(tmpItem.表头, "|")(i), "^")(1)
                    If Err <> 0 Then sgnH = 250
                    msh(idx).RowHeight(i) = sgnH * sgnMode
                    msh(idx).Row = i
                    msh(idx).Col = tmpItem.序号
                    Err = 0
                    sgnH = Split(Split(tmpItem.表头, "|")(i), "^")(0)
                    If Err <> 0 Then sgnH = 4
                    msh(idx).CellAlignment = sgnH
                    If UBound(Split(Split(tmpItem.表头, "|")(i), "^")) > 2 Then
                        msh(idx).CellFontBold = Split(Split(tmpItem.表头, "|")(i), "^")(3)
                        msh(idx).CellForeColor = Split(Split(tmpItem.表头, "|")(i), "^")(4)
                    End If
                Next
            Next
            
            For i = msh(idx).FixedRows To msh(idx).Rows - 1
                msh(idx).RowHeight(i) = .行高 * sgnMode
            Next
            
            '合并特性
            For i = 0 To msh(idx).FixedRows - 1
                msh(idx).MergeRow(i) = True
            Next
            For i = 0 To msh(idx).Cols - 1
                msh(idx).MergeCol(i) = True
            Next
            
'            Call SetHeadCenter(msh(idx)) '表头内容居中
            Call SetGridLine(.id) '填充表格线
            
            '分栏设置(对象CopyIDs已经设置)
            For Each tmpID In .CopyIDs
                Call SetGridSame(msh(idx), msh(tmpID.id))
            Next
        ElseIf .类型 = 5 Then '汇总表格
            msh(idx).ForeColor = .前景
            msh(idx).ForeColorFixed = .前景
            msh(idx).GridColor = .网格
            msh(idx).GridColorFixed = .网格
            
            msh(idx).BackColor = .背景
            msh(idx).BackColorFixed = IIF(.背景 = &HFFFFFF, lbl(0).BackColor, .背景)
            
            msh(idx).Font.name = .字体
            msh(idx).Font.Size = Format(.字号 * sgnMode, "0.0")
            msh(idx).Font.Bold = .粗体
            msh(idx).Font.Italic = .斜体
            msh(idx).Font.Underline = .下线
            msh(idx).GridLineWidth = IIF(.表格线加粗, 2, 1)
            
            X = 0: Y = 0: Z = 0
            For Each tmpID In .SubIDs
                Select Case objReport.Items("_" & tmpID.id).类型
                    Case 7
                        X = X + 1 '纵向分类数
                    Case 8
                        Y = Y + 1 '横向分类数
                    Case 9
                        Z = Z + 1 '统计项数
                End Select
            Next
            '最小行列数
            msh(idx).Rows = Y + 4
            msh(idx).FixedRows = Y + 1
            If Y = 0 Then
                msh(idx).Cols = X + Z
            Else
                msh(idx).Cols = X + IIF(Z = 1, Z + 1, Z)
            End If
            msh(idx).FixedCols = X
            msh(idx).RowHeight(0) = .行高 * sgnMode '行高0是标准
            msh(idx).RowHeightMin = msh(idx).RowHeight(0)
            
            '基本行列内容
            For Each tmpID In .SubIDs
                Set tmpItem = objReport.Items("_" & tmpID.id)
                Select Case tmpItem.类型
                    Case 7 '纵向分类
                        msh(idx).TextMatrix(msh(idx).FixedRows - 1, tmpItem.序号) = "[" & tmpItem.内容 & "]"
                        msh(idx).Cell(flexcpFontBold, msh(idx).FixedRows - 1, tmpItem.序号) = tmpItem.粗体
                        msh(idx).Cell(flexcpForeColor, msh(idx).FixedRows - 1, tmpItem.序号) = tmpItem.前景
                        
                        For i = msh(idx).FixedRows To msh(idx).Rows - 1
                            msh(idx).TextMatrix(i, tmpItem.序号) = tmpItem.内容
                        Next
                        If tmpItem.汇总 <> "" Then
                            msh(idx).TextMatrix(msh(idx).FixedRows, tmpItem.序号) = tmpItem.汇总
                        End If
                        
                        msh(idx).ColWidth(tmpItem.序号) = tmpItem.W * sgnMode
                        msh(idx).ColAlignment(tmpItem.序号) = Switch(tmpItem.对齐 = 0, 1, tmpItem.对齐 = 1, 4, tmpItem.对齐 = 2, 7)
                    Case 8 '横向分类
                        For i = 0 To msh(idx).FixedCols - 1
                            msh(idx).TextMatrix(tmpItem.序号, i) = "[" & tmpItem.内容 & "]"
                            msh(idx).Cell(flexcpFontBold, tmpItem.序号, i) = tmpItem.粗体
                            msh(idx).Cell(flexcpForeColor, tmpItem.序号, i) = tmpItem.前景
                        Next
                        
                        For i = msh(idx).FixedCols To msh(idx).Cols - 1
                            msh(idx).TextMatrix(tmpItem.序号, i) = tmpItem.内容
                        Next
                        If tmpItem.汇总 <> "" Then
                            msh(idx).TextMatrix(tmpItem.序号, msh(idx).FixedCols) = tmpItem.汇总
                        End If
                    Case 9 '统计项
                        msh(idx).TextMatrix(msh(idx).FixedRows - 1, msh(idx).FixedCols + tmpItem.序号) = "[" & tmpItem.内容 & "]"
                        msh(idx).ColWidth(msh(idx).FixedCols + tmpItem.序号) = tmpItem.W * sgnMode
                        msh(idx).ColAlignment(msh(idx).FixedCols + tmpItem.序号) = Switch(tmpItem.对齐 = 0, 1, tmpItem.对齐 = 1, 4, tmpItem.对齐 = 2, 7)
                        msh(idx).Cell(flexcpFontBold, msh(idx).FixedRows - 1, msh(idx).FixedCols + tmpItem.序号, msh(idx).Rows - 1, msh(idx).FixedCols + tmpItem.序号) = tmpItem.粗体
                        msh(idx).Cell(flexcpForeColor, msh(idx).FixedRows - 1, msh(idx).FixedCols + tmpItem.序号, msh(idx).Rows - 1, msh(idx).FixedCols + tmpItem.序号) = tmpItem.前景
                End Select
            Next
            
            '合并特性
            For i = 0 To msh(idx).FixedRows - 2
                msh(idx).MergeRow(i) = True
            Next
            For i = 0 To msh(idx).FixedCols - 1
                msh(idx).MergeCol(i) = True
            Next
            
            '表头行行高
             '自定义表头内容
            On Error Resume Next
            For i = 0 To msh(idx).FixedRows - 1
                Err = 0
                sgnH = Split(Split(tmpItem.表头, "|")(i), "^")(1)
                If Err <> 0 Then sgnH = 250
                msh(idx).RowHeight(i) = sgnH * sgnMode
                msh(idx).Row = i
            Next
            On Error GoTo 0
            Call SetGridLine(.id)
'            Call SetHeadCenter(msh(idx))
        End If
    End With
    msh(idx).Redraw = True
End Sub

Private Sub CustomCellColor(idx As Integer, sCell As Cells, Optional blnClear As Boolean = True)
'功能：按单元格类型内容填充颜色
'参数：blnClear=是否清除原有填充
'      sCell=填充单元格范围,Row=-1是表示不处理行
    Dim i As Integer, j As Integer
    Dim sRow As Integer, sCol As Integer
    
    If sCell.Col1 = -1 Or sCell.Col2 = -1 Or sCell.Row1 = -1 Or sCell.Row2 = -1 Then Exit Sub
    
    msh(idx).Redraw = False
    
    If blnClear Then '先清除上次的着色,这次重新着色
        For i = 0 To msh(idx).FixedRows - 1
            msh(idx).Row = i
            For j = 0 To msh(idx).Cols - 1
                msh(idx).Col = j: msh(idx).CellBackColor = msh(idx).BackColorFixed
            Next
        Next
    End If
    
'    '参照行颜色
'    If sCell.Row <> -1 Then
'        msh(idx).Row = sCell.Row
'        For i = 0 To msh(idx).Cols - 1
'            msh(idx).Col = i: msh(idx).CellBackColor = &HC0C0C0
'        Next
'    End If
        
    '单元格(范围)颜色
    sRow = 1
    If sCell.Row2 < sCell.Row1 Then sRow = -1
    sCol = 1
    If sCell.Col2 < sCell.Col1 Then sCol = -1
    For i = sCell.Row1 To sCell.Row2 Step sRow
        msh(idx).Row = i
        For j = sCell.Col1 To sCell.Col2 Step sCol
            msh(idx).Col = j: msh(idx).CellBackColor = &HE7CFBA
        Next
    Next
    
    msh(idx).Redraw = True
End Sub

Private Function CheckCell(idx As Integer, Row As Integer, Col As Integer, Text As String) As Boolean
'功能：检查当前单元格的文字是否可以填入,用以限制任意表格表头只能在单方向上合并
'参数：Text=将要填充在Row,Col单元格的文字
    Dim tmpCell As Cells
    
    '上方单元格
    If Row - 1 >= 0 Then
        If Text = msh(idx).TextMatrix(Row - 1, Col) Then
            tmpCell = GetCellRange(msh(idx), Row - 1, Col)
            If Abs(tmpCell.Col1 - tmpCell.Col2) <> 0 Then Exit Function
        End If
    End If
    '下方单元格
    If Row + 1 <= msh(idx).FixedRows - 1 Then
        If Text = msh(idx).TextMatrix(Row + 1, Col) Then
            tmpCell = GetCellRange(msh(idx), Row + 1, Col)
            If Abs(tmpCell.Col1 - tmpCell.Col2) <> 0 Then Exit Function
        End If
    End If
    '左边单元格
    If Col - 1 >= 0 Then
        If Text = msh(idx).TextMatrix(Row, Col - 1) Then
            tmpCell = GetCellRange(msh(idx), Row, Col - 1)
            If Abs(tmpCell.Row1 - tmpCell.Row2) <> 0 Then Exit Function
        End If
    End If
    '右边单元格
    If Col + 1 <= msh(idx).Cols - 1 Then
        If Text = msh(idx).TextMatrix(Row, Col + 1) Then
            tmpCell = GetCellRange(msh(idx), Row, Col + 1)
            If Abs(tmpCell.Row1 - tmpCell.Row2) <> 0 Then Exit Function
        End If
    End If
    CheckCell = True
End Function

Private Function MergeCell(idx As Integer, sCell As Cells, dCell As Cells) As Cells
'功能：任意表格表头,对两个单元格选择范围进行合并,并返回一个合并后的单元格范围
'说明：如果不能合并,则返回一个无效范围
    Dim i As Integer
    Dim tmpCell As Cells
    
    MergeCell.Row1 = -1: MergeCell.Col1 = -1: MergeCell.Row2 = -1: MergeCell.Col2 = -1
    
    If sCell.Row1 = sCell.Row2 And sCell.Col1 = sCell.Col2 Then '从单独一个单元格开始拖动
        MergeCell.Row1 = sCell.Row1
        MergeCell.Col1 = sCell.Col1
        If dCell.Row1 <> dCell.Row2 And dCell.Col1 <> dCell.Col2 Then
            If Abs(dCell.Row1 - dCell.Row2) >= Abs(dCell.Col1 - dCell.Col2) Then
                MergeCell.Row2 = sCell.Row1
                MergeCell.Col2 = dCell.Col2
            Else
                MergeCell.Row2 = dCell.Row2
                MergeCell.Col2 = sCell.Col1
            End If
        Else
            MergeCell.Row2 = dCell.Row2
            MergeCell.Col2 = dCell.Col2
        End If
    Else '从已合并单元格开始拖动
        If sCell.Row1 = sCell.Row2 Then '同一行合并
            If dCell.Row1 = dCell.Row2 Then
                MergeCell.Row1 = sCell.Row1
                MergeCell.Row2 = sCell.Row2
                If dCell.Col2 > dCell.Col1 Then '向右拖
                    MergeCell.Col1 = sCell.Col1
                    MergeCell.Col2 = dCell.Col2
                Else '向左拖
                    MergeCell.Col1 = dCell.Col2
                    MergeCell.Col2 = sCell.Col2
                End If
            End If
        ElseIf sCell.Col1 = sCell.Col2 Then '同一列合并
            If dCell.Col1 = dCell.Col2 Then
                MergeCell.Col1 = sCell.Col1
                MergeCell.Col2 = sCell.Col2
                If dCell.Row2 > dCell.Row1 Then '向下拖
                    MergeCell.Row1 = sCell.Row1
                    MergeCell.Row2 = dCell.Row2
                Else '向上拖
                    MergeCell.Row1 = dCell.Row2
                    MergeCell.Row2 = sCell.Row2
                End If
            End If
        End If
    End If
    
    MergeCell = AdjustCell(MergeCell)

    '处理包含关系
    If MergeCell.Row1 >= sCell.Row1 And MergeCell.Row2 <= sCell.Row2 And MergeCell.Col1 >= sCell.Col1 And MergeCell.Col2 <= sCell.Col2 Then MergeCell = sCell
    
    '中途有相返方向合并单元
    If MergeCell.Col1 <> -1 And MergeCell.Col2 <> -1 And MergeCell.Row1 <> -1 And MergeCell.Row2 <> -1 Then
        If MergeCell.Row1 = MergeCell.Row2 Then
            For i = MergeCell.Col1 To MergeCell.Col2
                If i < sCell.Col1 Or i > sCell.Col2 Then
                    tmpCell = GetCellRange(msh(idx), MergeCell.Row1, i)
                    If tmpCell.Row1 <> tmpCell.Row2 Then MergeCell = sCell: Exit For
                End If
            Next
        End If
        If MergeCell.Col1 = MergeCell.Col2 Then
            For i = MergeCell.Row1 To MergeCell.Row2
                If i < sCell.Row1 Or i > sCell.Row2 Then
                    tmpCell = GetCellRange(msh(idx), i, MergeCell.Col1)
                    If tmpCell.Col1 <> tmpCell.Col2 Then MergeCell = sCell: Exit For
                End If
            Next
        End If
    End If

    '处理半截单元
    If MergeCell.Row1 <> -1 And MergeCell.Row2 <> -1 And MergeCell.Col1 <> -1 And MergeCell.Col2 <> -1 Then
        tmpCell = GetCellRange(msh(idx), MergeCell.Row1, MergeCell.Col1)
        If tmpCell.Col1 <> tmpCell.Col2 And MergeCell.Row1 = MergeCell.Row2 Then MergeCell.Col1 = tmpCell.Col1
        If tmpCell.Row1 <> tmpCell.Row2 And MergeCell.Col1 = MergeCell.Col2 Then MergeCell.Row1 = tmpCell.Row1
        tmpCell = GetCellRange(msh(idx), MergeCell.Row2, MergeCell.Col2)
        If tmpCell.Col1 <> tmpCell.Col2 And MergeCell.Row1 = MergeCell.Row2 Then MergeCell.Col2 = tmpCell.Col2
        If tmpCell.Row1 <> tmpCell.Row2 And MergeCell.Col1 = MergeCell.Col2 Then MergeCell.Row2 = tmpCell.Row2
    End If
End Function

Private Function AdjustCell(sCell As Cells) As Cells
    Dim i As Integer
    If sCell.Row1 > sCell.Row2 Then
        i = sCell.Row1
        sCell.Row1 = sCell.Row2
        sCell.Row2 = i
    End If
    If sCell.Col1 > sCell.Col2 Then
        i = sCell.Col1
        sCell.Col1 = sCell.Col2
        sCell.Col2 = i
    End If
    AdjustCell = sCell
End Function

Private Sub CustomColColor(idx As Integer, Col As Integer, Optional Clear As Boolean = True)
'功能：绘制任意表格表列颜色
    Dim i As Integer, j As Integer, LngCurRow As Long
    
    If Col = -1 Then Exit Sub
    
    LngCurRow = msh(idx).Row
    msh(idx).Redraw = False
    For i = msh(idx).FixedRows To msh(idx).Rows - 1
        msh(idx).Row = i
        For j = 0 To msh(idx).Cols - 1
            msh(idx).Col = j
            If j <> Col And Clear Then
                msh(idx).CellBackColor = msh(idx).BackColor
            ElseIf j = Col Then
                msh(idx).CellBackColor = &HE7CFBA
            End If
        Next
    Next
    msh(idx).Row = LngCurRow
    If msh(idx).Cols > 0 Then
        msh(idx).Col = IIF(Col = -9, 0, Col)
    End If
    msh(idx).Redraw = True
End Sub

Private Function CheckHead() As String
'功能：检查任意表格表头数据是否超长
'返回：提示信息
    Dim tmpItem As RPTItem, tmpID As RelatID
    
    For Each tmpItem In objReport.Items
        If tmpItem.类型 = 4 Then
            For Each tmpID In tmpItem.SubIDs
                If LenB(StrConv(objReport.Items("_" & tmpID.id).表头, vbFromUnicode)) > 4000 Then
                    CheckHead = "报表中某个任意表格的表头文字过长或空表头行过多,请检查！"
                    Exit Function
                End If
                If CheckText(objReport.Items("_" & tmpID.id), True) = False Then
                    CheckHead = "报表中某个任意表格的表头数据源不正确,请检查！"
                    Exit Function
                End If
            Next
        End If
    Next
End Function

Private Sub SetMenuDefault(idx As Integer, Col As Integer)
'功能：对任意表格,根据当前列对齐及汇总方式,设置菜单复选项
    Dim tmpID As RelatID, tmpItem As RPTItem
    Dim lngType As Long, i As Integer
    
    mnuCustom_Col_State.Enabled = True
    mnuCustom_Col_Align.Enabled = True
    
    For i = 0 To mnuCustom_Col_Align_Style.UBound
        mnuCustom_Col_Align_Style(i).Checked = False
    Next
    For i = 0 To mnuCustom_Col_State_Style.UBound
        mnuCustom_Col_State_Style(i).Checked = False
    Next
    If Col = -1 Then Exit Sub
    
    For Each tmpID In objReport.Items("_" & intCurID).SubIDs
        Set tmpItem = objReport.Items("_" & tmpID.id)
        If tmpItem.序号 = intCurCol Then
            Select Case tmpItem.汇总
                Case ""
                    mnuCustom_Col_State_Style(0).Checked = True
                Case "SUM"
                    mnuCustom_Col_State_Style(1).Checked = True
                Case "AVG"
                    mnuCustom_Col_State_Style(2).Checked = True
                Case "MAX"
                    mnuCustom_Col_State_Style(3).Checked = True
                Case "MIN"
                    mnuCustom_Col_State_Style(4).Checked = True
                Case "COUNT"
                    mnuCustom_Col_State_Style(5).Checked = True
            End Select
            mnuCustom_Col_Align_Style(tmpItem.对齐).Checked = True
            
            '根据单一字段类型设置可用菜单
            If tmpItem.内容 Like "*.*" And Left(tmpItem.内容, 1) = "[" And Right(tmpItem.内容, 1) = "]" Then
                lngType = GetNodeType(Mid(tmpItem.内容, 2, Len(tmpItem.内容) - 2), tvwSQL)
                If IsType(lngType, adVarChar) Then
                    '字符不能使用汇总
                    mnuCustom_Col_State.Enabled = False
                End If
                If IsType(lngType, adLongVarBinary) Then
                    '图片不使用汇总和对齐(固定中对齐)
                    mnuCustom_Col_State.Enabled = False
                    mnuCustom_Col_Align.Enabled = False
                End If
            End If
            
            Exit For
        End If
    Next
End Sub

Private Sub ClassColor(idx As Integer, sCell As Cells, Optional intState As Integer)
'功能：根据当前行列值填充汇总表格的颜色
'参数：sCell=仅Row1,Col1有效;intState=统计项数
    Dim i As Integer
    On Error Resume Next
    msh(idx).Redraw = False
    If sCell.Col1 <= msh(idx).FixedCols - 1 And sCell.Row1 >= msh(idx).FixedRows - 1 Then
         '纵向分类范围
         msh(idx).Col = sCell.Col1
         For i = msh(idx).FixedRows - 1 To msh(idx).Rows - 1
            msh(idx).Row = i: msh(idx).CellBackColor = &HE7CFBA
         Next
    ElseIf sCell.Row1 <= msh(idx).FixedRows - 2 Then
         '横向分类范围
         msh(idx).Row = sCell.Row1
         For i = 0 To msh(idx).Cols - 1
            msh(idx).Col = i: msh(idx).CellBackColor = &HE7CFBA
         Next
    ElseIf sCell.Col1 >= msh(idx).FixedCols And sCell.Col1 <= msh(idx).FixedCols + intState - 1 And sCell.Row1 >= msh(idx).FixedRows - 1 Then
        '统计项范围
        msh(idx).Col = sCell.Col1
        For i = msh(idx).FixedRows - 1 To msh(idx).Rows - 1
            msh(idx).Row = i: msh(idx).CellBackColor = &HE7CFBA
        Next
    End If
    msh(idx).Redraw = True
End Sub

Private Sub ReFlashWidth()
'功能：处理表格列宽(根据控件内容刷新对象内容)
    Dim tmpID As RelatID, tmpItem As RPTItem
    For Each tmpItem In objReport.Items
        If tmpItem.格式号 = Mid(cboFormat.ComboItems("_" & mbytCurrFmt).Key, 2) Then  'zyb#Add
            If tmpItem.类型 = 4 Then
                For Each tmpID In tmpItem.SubIDs
                    objReport.Items("_" & tmpID.id).W = msh(tmpItem.id).ColWidth(objReport.Items("_" & tmpID.id).序号) / sgnMode
                Next
            ElseIf tmpItem.类型 = 5 Then
                For Each tmpID In tmpItem.SubIDs
                    If objReport.Items("_" & tmpID.id).类型 = 7 Then '纵
                        objReport.Items("_" & tmpID.id).W = msh(tmpItem.id).ColWidth(objReport.Items("_" & tmpID.id).序号) / sgnMode
                    ElseIf objReport.Items("_" & tmpID.id).类型 = 9 Then '统
                        objReport.Items("_" & tmpID.id).W = msh(tmpItem.id).ColWidth(msh(tmpItem.id).FixedCols + objReport.Items("_" & tmpID.id).序号) / sgnMode
                    End If
                Next
            End If
        End If
    Next
End Sub

Private Sub SetDefaultState(strState As String, intAlign As Integer, Optional blnClear As Boolean)
'功能：对汇总表格,根据当前子项汇总方式设置菜单显示
    Dim i As Integer
    For i = 0 To mnuClass_State_Style.UBound
        mnuClass_State_Style(i).Checked = False
    Next
    For i = 0 To mnuClass_Align_Style.UBound
        mnuClass_Align_Style(i).Checked = False
    Next
    If Not blnClear Then
        mnuClass_Align_Style(intAlign).Checked = True
        Select Case strState
            Case ""
                mnuClass_State_Style(0).Checked = True
            Case "SUM"
                mnuClass_State_Style(1).Checked = True
            Case "AVG"
                mnuClass_State_Style(2).Checked = True
            Case "MAX"
                mnuClass_State_Style(3).Checked = True
            Case "MIN"
                mnuClass_State_Style(4).Checked = True
            Case "COUNT"
                mnuClass_State_Style(5).Checked = True
        End Select
    End If
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage hwnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo hwnd
End Sub

Private Function UpdatePriv() As Boolean
'功能：更新报表访问权限(独立报表的权限更新及报表组的权限更新)
'返回：成功=True,失败=False
    Dim rsTmp As ADODB.Recordset, tmpData As RPTData, tmpPar As RPTPar
    Dim strObject As String, strOwner As String, strName As String
    Dim lngProgID As Long, i As Integer, j As Integer, blnTran As Boolean
    Dim strSQL As String, strTmp As String
    
    On Error GoTo errH
    
    '非当前登录的数据库连接忽略
    strObject = ""
    For Each tmpData In objReport.Datas
        If tmpData.对象 <> "" And tmpData.数据连接编号 <= 0 Then
            For j = 0 To UBound(Split(tmpData.对象, ","))
                If InStr(strObject & ",", "," & Split(tmpData.对象, ",")(j) & ",") = 0 Then
                    strObject = strObject & "," & Split(tmpData.对象, ",")(j)
                End If
            Next
        End If
        If tmpData.数据连接编号 <= 0 Then
            For Each tmpPar In tmpData.Pars
                If tmpPar.对象 <> "" Then
                    For i = 0 To UBound(Split(tmpPar.对象, "|"))
                        strTmp = Split(tmpPar.对象, "|")(i)
                        If strTmp <> "" Then
                            For j = 0 To UBound(Split(strTmp, ","))
                                If InStr(strObject & ",", "," & Split(strTmp, ",")(j) & ",") = 0 Then
                                    strObject = strObject & "," & Split(strTmp, ",")(j)
                                End If
                            Next
                        End If
                    Next
                End If
            Next
        End If
    Next
    strObject = Mid(strObject, 2)
    
    '是否已发布(各种发布方式)
    strSQL = _
        " Select " & IIF(objReport.系统 = 0, "-Null", objReport.系统) & " as 系统,程序ID,功能" & _
        " From zlReports Where 程序ID is Not Null And ID=[1]" & _
        " Union" & _
        " Select " & IIF(objReport.系统 = 0, "-Null", objReport.系统) & " as 系统,A.程序ID,B.功能" & _
        " From zlRPTGroups A,zlRPTSubs B" & _
        " Where A.程序ID is Not Null And A.ID=B.组ID And B.报表ID=[1]" & _
        " Union " & _
        " Select 系统,程序ID,功能 From zlRPTPuts Where 报表ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngRPTID)
    If Not rsTmp.EOF Then
        On Error GoTo errH
        gcnOracle.BeginTrans: blnTran = True
        Do While Not rsTmp.EOF
            '重新填写权限
            If IsNull(rsTmp!功能) Then
                gcnOracle.Execute "Delete From zlProgPrivs Where 序号=" & rsTmp!程序id & _
                                  " And 功能 is Null And Nvl(系统,0)=" & Nvl(rsTmp!系统, 0)
            Else
                gcnOracle.Execute "Delete From zlProgPrivs Where 序号=" & rsTmp!程序id & _
                                  " And 功能='" & rsTmp!功能 & "' And Nvl(系统,0)=" & Nvl(rsTmp!系统, 0)
            End If
            If strObject <> "" Then '该表格有可能不访问数据库
                For i = 0 To UBound(Split(strObject, ","))
                    strOwner = Left(Split(strObject, ",")(i), InStr(Split(strObject, ",")(i), ".") - 1)
                    If strOwner <> "SYS" And strOwner <> "ZLTOOLS" And strOwner <> "SYSTEM" Then
                        strName = Mid(Split(strObject, ",")(i), InStr(Split(strObject, ",")(i), ".") + 1)
                        gcnOracle.Execute GetInsertProgPrivs(Nvl(rsTmp!系统, 0), Nvl(rsTmp!程序id, 0) _
                                                , rsTmp!功能, strName, strOwner, "SELECT")
                    End If
                Next
            End If
            rsTmp.MoveNext
        Loop
        gcnOracle.CommitTrans: blnTran = False
    End If
    UpdatePriv = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    If blnTran Then gcnOracle.RollbackTrans
    Call SaveErrLog
End Function

Private Sub ReFlashReportBySelFormat()
'zyb#Add
'功能：根据指定报表格式,重新刷新显示报表内容
    Dim objTmp As Object
    
    For Each objTmp In lblSize
        If objTmp.Index <> 0 Then Unload objTmp
    Next
    For Each objTmp In lblLine
        If objTmp.Index <> 0 Then Unload objTmp
    Next
    For Each objTmp In lbl
        If objTmp.Index <> 0 Then Unload objTmp
    Next
    For Each objTmp In msh
        If objTmp.Index <> 0 Then Unload objTmp
    Next
    For Each objTmp In img
        If objTmp.Index <> 0 Then Unload objTmp
    Next
    For Each objTmp In ImgCode
        If objTmp.Index <> 0 Then Unload objTmp
    Next
    For Each objTmp In Chart
        If objTmp.Index <> 0 Then Unload objTmp
    Next
    For Each objTmp In pic
        If objTmp.Index <> 0 Then Unload objTmp
    Next
    For Each objTmp In Shp
        If objTmp.Index <> 0 Then Unload lblshp(objTmp.Index): Unload objTmp
    Next
    
    intCurID = 0
    Set objLastSel = Nothing
    
    Call ShowItems
    Call ShowAttrib
    Call AdjustAll
    
    Me.Refresh
End Sub

Private Function GetRPtFmtName() As String
    Dim IntFmt As Integer, StrFmtName As String, CboItem As ComboItem, StrCompare As String
    'zyb#Add
    '产生缺省报表样式的名称(不重复)
    
    StrFmtName = ""
    For IntFmt = 1 To cboFormat.ComboItems.count
        Set CboItem = cboFormat.ComboItems(IntFmt)
        StrCompare = CStr(Val(StrReverse(Val(StrReverse(CboItem.Text & "9"))))) '加上一个数字,避免最后一位零的丢失
        StrCompare = Mid(StrCompare, 1, Len(StrCompare) - 1)
        
        If Len(StrCompare) > Len(StrFmtName) Then
            StrFmtName = StrCompare
        ElseIf Len(StrCompare) = Len(StrFmtName) Then
            If StrFmtName < StrCompare Then StrFmtName = StrCompare
        End If
    Next
    
    '对取得的数字加1
    If StrFmtName = "" Then
        GetRPtFmtName = objReport.名称 & "1"
    Else
        GetRPtFmtName = objReport.名称 & (CLng(StrFmtName) + 1)
    End If
End Function

Private Sub DrawFrame(ByVal ObjOper As Label)
    'zyb#Add
    '画边框线
    
'    With WinProperty
'        .L = ObjOper.Left - 10
'        .T = ObjOper.Top - 10
'        .H = ObjOper.Top + ObjOper.Height - 10
'        .W = ObjOper.Left + ObjOper.Width - 10
'
'        picPaper.Line (.L, .T)-(.W, .H), objReport.Items("_" & ObjOper.Index).前景, B
'    End With
End Sub

Private Function GetAutoTest(ByVal bytMode As Byte) As Single
    Dim sgnCompare As Single, sgnActure As Single
    Dim objFmt As RPTFmt
    
    '获取适应的比例
    Set objFmt = objReport.Fmts("_" & mbytCurrFmt)
    Select Case bytMode
    Case 0  '自适应
        sgnCompare = IIF(objFmt.H > objFmt.W, objFmt.H, objFmt.W)
        sgnActure = IIF(picPaper.Height > picPaper.Width, scrHsc.Top - picRulerH.Top - picRulerH.Height, scrVsc.Left - picRulerV.Left - picRulerV.Width)
    Case 1  '适应宽度
        sgnCompare = IIF(objFmt.纸向 = 1, objFmt.W, objFmt.H)
        sgnActure = scrVsc.Left - picRulerV.Left - picRulerV.Width
    Case 2  '适应高度
        sgnCompare = IIF(objFmt.纸向 = 1, objFmt.H, objFmt.W)
        sgnActure = scrHsc.Top - picRulerH.Top - picRulerH.Height
    End Select
    GetAutoTest = (sgnActure - 200) / sgnCompare
End Function

Private Function CheckNameValid(ByVal strName As String) As Boolean
    Dim ItemThis As RPTItem
    '检查名称的合法性
    
    CheckNameValid = False
    For Each ItemThis In objReport.Items
        If ItemThis.格式号 = mbytCurrFmt Then
            If ItemThis.Key <> intCurID And strName = ItemThis.名称 Then Exit Function
        End If
    Next
    CheckNameValid = True
End Function

Private Sub ReferTo(Optional ByVal ItemTest As RPTItem)
    Dim ItemThis As RPTItem
    Dim ObjSel As Control, BytWay As Byte, bytKind As Byte
    Dim TargetObj As VSFlexGrid, ReferToObjname As String
    Dim DblAdd As Double

    '设置标签的位置,随参照对象,方向及性质的不同而发生变化
    If intCurID = 0 Then Exit Sub
    If Val(objReport.Items("_" & intCurID).性质) = 0 Then Exit Sub
    
    Select Case objReport.Items("_" & intCurID).类型
        Case 2
            Set ObjSel = lbl(intCurID)
            If objReport.Items("_" & intCurID).父ID = 0 Then
                Set ObjSel.Container = picPaper
            Else
                Set ObjSel.Container = pic(objReport.Items("_" & intCurID).父ID)
            End If
            ReferToObjname = objReport.Items("_" & intCurID).参照   '参照
            
            For Each ItemThis In objReport.Items
                If ItemThis.名称 = ReferToObjname And ItemThis.格式号 = mbytCurrFmt Then Set TargetObj = msh(ItemThis.Key): Exit For
            Next
        
            '移动控件
            Select Case Mid(objReport.Items("_" & intCurID).性质, 1, 1)
            Case 1
                If Not (ObjSel.Top + ObjSel.Height + 100 * sgnMode < TargetObj.Top) Then
                    ObjSel.Top = TargetObj.Top - 100 * sgnMode - ObjSel.Height
                End If
            Case 2
                If Not (ObjSel.Top >= TargetObj.Top + GetTableHeight(TargetObj) + 50 * sgnMode And ObjSel.Top <= picPaper.Height - 200 * sgnMode) Then
                    ObjSel.Top = TargetObj.Top + GetTableHeight(TargetObj) + 100 * sgnMode
                End If
            End Select
            Select Case Mid(objReport.Items("_" & intCurID).性质, 2)
            Case 1
                ObjSel.Left = TargetObj.Left
            Case 2
                ObjSel.Left = GetTableWidth(TargetObj) / 2 + TargetObj.Left - (ObjSel.Width / 2)
            Case 3
                ObjSel.Left = GetTableWidth(TargetObj) + TargetObj.Left - ObjSel.Width
            End Select
            objReport.Items("_" & ObjSel.Index).X = ObjSel.Left / sgnMode
            objReport.Items("_" & ObjSel.Index).Y = ObjSel.Top / sgnMode
            Call AdjustSelCons(ObjSel)
        Case 4, 5
            '获取主表
            For Each ItemThis In objReport.Items
                If ItemThis.名称 = IIF(ReferToObjname = "", objReport.Items("_" & intCurID).参照, ReferToObjname) And ItemThis.格式号 = mbytCurrFmt Then Set TargetObj = msh(ItemThis.Key): Exit For
            Next
            
            If objReport.Items("_" & intCurID).参照 <> "" And ItemTest.参照 = "" Then
                BytWay = objReport.Items("_" & intCurID).性质
                ReferToObjname = objReport.Items("_" & intCurID).参照
                
                '新增则移至最后
                DblAdd = IIF(BytWay = 1, TargetObj.Top, TargetObj.Left) + IIF(BytWay = 1, TargetObj.Height, TargetObj.Width) - 15
                For Each ItemThis In objReport.Items
                    If ItemThis.参照 = ReferToObjname And ItemThis.格式号 = mbytCurrFmt And ItemThis.Key <> intCurID And InStr(1, "4,5", ItemThis.类型) <> 0 Then
                        Set ObjSel = msh(ItemThis.Key)
                        DblAdd = DblAdd + IIF(BytWay = 1, ObjSel.Height, ObjSel.Width) - 15 * sgnMode
                    End If
                Next
                
                Set ObjSel = msh(intCurID)
                If objReport.Items("_" & intCurID).父ID = 0 Then
                    Set ObjSel.Container = picPaper
                Else
                    Set ObjSel.Container = pic(objReport.Items("_" & intCurID).父ID)
                End If
                ObjSel.Left = IIF(BytWay = 1, TargetObj.Left, DblAdd)
                ObjSel.Top = IIF(BytWay = 1, DblAdd, TargetObj.Top)
                If BytWay = 2 Then
                    ObjSel.Height = TargetObj.Height
                Else
                    ObjSel.Width = TargetObj.Width
                End If
                objReport.Items("_" & ObjSel.Index).Y = ObjSel.Top / sgnMode
                objReport.Items("_" & ObjSel.Index).X = ObjSel.Left / sgnMode
                objReport.Items("_" & ObjSel.Index).W = ObjSel.Width / sgnMode
                objReport.Items("_" & ObjSel.Index).H = ObjSel.Height / sgnMode
            Else
                BytWay = ItemTest.性质
                ReferToObjname = ItemTest.参照
                
                For Each ItemThis In objReport.Items
                    If ItemThis.参照 = ReferToObjname And ItemThis.格式号 = mbytCurrFmt And InStr(1, "4,5", ItemThis.类型) <> 0 Then
                        Set ObjSel = msh(ItemThis.Key)
                        Select Case BytWay
                        Case 1
                            If ObjSel.Top > msh(intCurID).Top Then
                                ObjSel.Top = ObjSel.Top - msh(intCurID).Height
                                ItemThis.Y = ObjSel.Top / sgnMode
                            End If
                        Case 2
                            If ObjSel.Left > msh(intCurID).Left Then
                                ObjSel.Left = ObjSel.Left - msh(intCurID).Width
                                ItemThis.X = ObjSel.Left / sgnMode
                            End If
                        End Select
                    End If
                Next
                Set ObjSel = msh(intCurID)
                If objReport.Items("_" & intCurID).父ID = 0 Then
                    Set ObjSel.Container = picPaper
                Else
                    Set ObjSel.Container = pic(objReport.Items("_" & intCurID).父ID)
                End If
            
                BytWay = objReport.Items("_" & intCurID).性质
                ReferToObjname = objReport.Items("_" & intCurID).参照
                
                If ReferToObjname <> "" Then
                    '新增则移至最后
                    DblAdd = IIF(BytWay = 1, TargetObj.Top, TargetObj.Left) + IIF(BytWay = 1, TargetObj.Height, TargetObj.Width) - 15
                    For Each ItemThis In objReport.Items
                        If ItemThis.参照 = ReferToObjname And ItemThis.格式号 = mbytCurrFmt And ItemThis.Key <> intCurID And InStr(1, "4,5", ItemThis.类型) <> 0 Then
                            Set ObjSel = msh(ItemThis.Key)
                            DblAdd = DblAdd + IIF(BytWay = 1, ObjSel.Height, ObjSel.Width) - 15 * sgnMode
                        End If
                    Next
                    
                    Set ObjSel = msh(intCurID)
                    ObjSel.Left = IIF(BytWay = 1, TargetObj.Left, DblAdd)
                    ObjSel.Top = IIF(BytWay = 1, DblAdd, TargetObj.Top)
                    ObjSel.Height = TargetObj.Height
                    ObjSel.Width = TargetObj.Width
                    objReport.Items("_" & ObjSel.Index).Y = ObjSel.Top / sgnMode
                    objReport.Items("_" & ObjSel.Index).X = ObjSel.Left / sgnMode
                    objReport.Items("_" & ObjSel.Index).W = ObjSel.Width / sgnMode
                    objReport.Items("_" & ObjSel.Index).H = ObjSel.Height / sgnMode
                End If
            End If
            SetGridLine (ObjSel.Index)
            Call AdjustSelCons(ObjSel)
            ObjSel.ZOrder 0
    End Select
End Sub

Private Sub AdjustSelCons(ByVal ObjSel As Object)
    '调整选择控件LblSize的位置
    Dim i As Integer, lngTop As Long, lngLeft As Long
    With WinProperty
        .H = lblSize(0).Height / Screen.TwipsPerPixelX
        .W = lblSize(0).Width / Screen.TwipsPerPixelX
    End With
    If UCase(ObjSel.Container.name) = "PIC" Then
        lngTop = ObjSel.Container.Top
        lngLeft = ObjSel.Container.Left
    End If

    If Mid(ObjSel.Tag, 3) = "" Then Exit Sub
    For i = Mid(ObjSel.Tag, 3) To Mid(ObjSel.Tag, 3) + 7 '选择标志从"上中"开始,"顺时针"处理
        Select Case IIF(i Mod 8 <> 0, i Mod 8, 8) '定位选择边框的位置
            Case 1 '上中
                With WinProperty
                    .T = (ObjSel.Top + lngTop - lblSize(i).Height) / Screen.TwipsPerPixelX
                    .l = (ObjSel.Left + lngLeft + (ObjSel.Width - lblSize(i).Width) / 2) / Screen.TwipsPerPixelX
                    Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                End With
                lblSize(i).MousePointer = 7
            Case 2 '上右
                With WinProperty
                    .T = (ObjSel.Top + lngTop - lblSize(i).Height) / Screen.TwipsPerPixelX
                    .l = (ObjSel.Left + lngLeft + ObjSel.Width) / Screen.TwipsPerPixelX
                    Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                End With
                lblSize(i).MousePointer = 6
            Case 3 '右中
                With WinProperty
                    .T = (ObjSel.Top + lngTop + (ObjSel.Height - lblSize(i).Height) / 2) / Screen.TwipsPerPixelX
                    .l = (ObjSel.Left + lngLeft + ObjSel.Width) / Screen.TwipsPerPixelX
                    Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                End With
                lblSize(i).MousePointer = 9
            Case 4 '右下
                With WinProperty
                    .T = (ObjSel.Top + lngTop + ObjSel.Height) / Screen.TwipsPerPixelX
                    .l = (ObjSel.Left + lngLeft + ObjSel.Width) / Screen.TwipsPerPixelX
                    Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                End With
                lblSize(i).MousePointer = 8
            Case 5 '下中
                With WinProperty
                    .T = (ObjSel.Top + lngTop + ObjSel.Height) / Screen.TwipsPerPixelX
                    .l = (ObjSel.Left + lngLeft + (ObjSel.Width - lblSize(i).Width) / 2) / Screen.TwipsPerPixelX
                    Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                End With
                lblSize(i).MousePointer = 7
            Case 6 '左下
                With WinProperty
                    .T = (ObjSel.Top + lngTop + ObjSel.Height) / Screen.TwipsPerPixelX
                    .l = (ObjSel.Left + lngLeft - lblSize(i).Width) / Screen.TwipsPerPixelX
                    Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                End With
                lblSize(i).MousePointer = 6
            Case 7 '左中
                With WinProperty
                    .T = (ObjSel.Top + lngTop + (ObjSel.Height - lblSize(i).Height) / 2) / Screen.TwipsPerPixelX
                    .l = (ObjSel.Left + lngLeft - lblSize(i).Width) / Screen.TwipsPerPixelX
                    Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                End With
                lblSize(i).MousePointer = 9
            Case 8 '左上
                With WinProperty
                    .T = (ObjSel.Top + lngTop - lblSize(i).Height) / Screen.TwipsPerPixelX
                    .l = (ObjSel.Left + lngLeft - lblSize(i).Width) / Screen.TwipsPerPixelX
                    Call MoveWindow(lblSize(i).hwnd, .l, .T, .W, .H, 1)
                End With
                lblSize(i).MousePointer = 8
        End Select
        lblSize(i).ZOrder
        lblSize(i).Visible = True
    Next
End Sub

Private Sub AdjustAll(Optional ByVal BlnSelected As Boolean = False)
    Dim ItemThis As RPTItem, intTmpIdx As Integer
    
    '用于在重新装入后,调整所有控件
    intTmpIdx = intCurID
    For Each ItemThis In objReport.Items
        If ItemThis.格式号 = mbytCurrFmt Then
            Select Case ItemThis.类型
                Case 2
                    intCurID = ItemThis.Key
                    Call ReferTo(ItemThis)
            End Select
        End If
    Next
    intCurID = intTmpIdx
    If GetSelNum = 1 And intCurID <> 0 Then
        Call ShowAttrib(intCurID)
    End If
End Sub

Private Function TestReferTo(ByVal LngIdx As Long, ByVal lngX As Long, ByVal lngY As Long) As Boolean
    Dim ItemThis As RPTItem
    Dim ObjSel As Label, BytWay As Byte, bytKind As Byte
    Dim TargetObj As VSFlexGrid, ReferToObjname As String
    '设置标签的位置,随参照对象,方向及性质的不同而发生变化
    TestReferTo = False
    If LngIdx = 0 Then Exit Function
    If objReport.Items("_" & LngIdx).参照 = "" Then Exit Function
    If Val(objReport.Items("_" & LngIdx).性质) = 0 Then Exit Function
    
    Set ObjSel = lbl(LngIdx)
    ReferToObjname = objReport.Items("_" & LngIdx).参照
    bytKind = objReport.Items("_" & LngIdx).性质   '性质
    
    For Each ItemThis In objReport.Items
        If ItemThis.名称 = ReferToObjname And ItemThis.格式号 = mbytCurrFmt Then Set TargetObj = msh(ItemThis.Key): Exit For
    Next
    
    '移动控件
    Select Case Mid(bytKind, 1, 1)
    Case 1, 2
        If Mid(bytKind, 1, 1) = "1" Then
            If ObjSel.Top + ObjSel.Height + lngY + 100 * sgnMode >= TargetObj.Top Then TestReferTo = True
        Else
            If Not (ObjSel.Top + lngY >= TargetObj.Top + GetTableHeight(TargetObj) + 50 And ObjSel.Top + lngY <= picPaper.Height - 200) Then TestReferTo = True
        End If
        If TestReferTo = False Then
            Select Case Mid(bytKind, 2)
            Case 1
                If Not (ObjSel.Left + lngX >= TargetObj.Left And ObjSel.Left + lngX <= TargetObj.Left + 200 * sgnMode) Then TestReferTo = True
            Case 2
                If Not (ObjSel.Left + lngX >= GetTableWidth(TargetObj) / 2 + TargetObj.Left - (ObjSel.Width / 2) - 100 * sgnMode And ObjSel.Left + lngX <= GetTableWidth(TargetObj) / 2 + TargetObj.Left - (ObjSel.Width / 2) + 100 * sgnMode) Then TestReferTo = True
            Case 3
                If Not (ObjSel.Left + lngX >= GetTableWidth(TargetObj) + TargetObj.Left - ObjSel.Width - 100 * sgnMode And ObjSel.Left + lngX <= GetTableWidth(TargetObj) + TargetObj.Left - ObjSel.Width + 100 * sgnMode) Then TestReferTo = True
            End Select
        End If
    End Select
End Function

Private Function CheckTableProperty(ByVal ItemThis As RPTItem) As Boolean
    Dim ItemTest As RPTItem, IntTest As Integer, StrTest As String, StrThis As String
    '检查该表格是否允许被左联接或附加
    CheckTableProperty = False
    
    'Key值相同则退出
    Set ItemTest = objReport.Items("_" & intCurID)
    If ItemTest.Key = ItemThis.Key Then Exit Function
    If ItemThis.参照 <> "" Then Exit Function
    
    Select Case ItemTest.类型
    Case 4  '附加表格(只能是清册表格)
        '检查附加表格是否分栏(不能为栏)
        If InStr(1, "4,5", ItemThis.类型) = 0 Then Exit Function
        If Not (ItemThis.分栏 < 2) Then Exit Function
        
    Case 5  '左联接表格(只能是分类表格)
        If ItemTest.SubIDs.count = 0 Or ItemThis.SubIDs.count = 0 Then Exit Function
        If ItemThis.类型 = 4 Then Exit Function
        
        '检查被左联接的表格与左联接表格的纵向分类是否一致
        StrTest = "": StrThis = ""
        For IntTest = 1 To ItemTest.SubIDs.count
            If objReport.Items("_" & ItemTest.SubIDs(IntTest).Key).类型 = 7 Then
                StrTest = StrTest & "|" & objReport.Items("_" & ItemTest.SubIDs(IntTest).Key).内容
            End If
        Next
        For IntTest = 1 To ItemThis.SubIDs.count
            If objReport.Items("_" & ItemThis.SubIDs(IntTest).Key).类型 = 7 Then
                StrThis = StrThis & "|" & objReport.Items("_" & ItemThis.SubIDs(IntTest).Key).内容
            End If
        Next
        If StrTest <> StrThis Then Exit Function
    End Select
    
    CheckTableProperty = True
End Function

Private Sub LinkMove(ByVal LngIdx As Long, ByVal lngX As Long, ByVal lngY As Long)
    Dim ItemThis As RPTItem, ObjSel As Control
    '连锁移动附加表格及左联接表格
    
    For Each ItemThis In objReport.Items
        If ItemThis.参照 = objReport.Items("_" & LngIdx).名称 _
            And ItemThis.格式号 = mbytCurrFmt And ItemThis.参照 <> "" Then
            Select Case ItemThis.类型
                Case 2
                    Set ObjSel = lbl(ItemThis.Key)
                Case 4, 5
                    Set ObjSel = msh(ItemThis.Key)
                Case 12 '@@@
                    Set ObjSel = Chart(ItemThis.Key)
            End Select
            
            ObjSel.Left = ObjSel.Left + lngX
            ObjSel.Top = ObjSel.Top + lngY
            ItemThis.X = ObjSel.Left / sgnMode
            ItemThis.Y = ObjSel.Top / sgnMode
            ObjSel.ZOrder
            
            Call AdjustSelCons(ObjSel)
    
            '设置为选中状态
            On Error Resume Next '可能传入主表进来,而实际移动的是子表
            lblSize(Mid(msh(LngIdx).Tag, 3) + 2).ZOrder
            lblSize(Mid(msh(LngIdx).Tag, 3) + 4).ZOrder
        End If
    Next
End Sub

Private Sub SetMainWH(ByVal LngIdx As Long)
    Dim ItemThis As RPTItem, SelObj As Control
    '传入子表,处理主表与子表一致(左联接则高度一致,附加则宽度一致)
    If InStr(1, "4,5", objReport.Items("_" & LngIdx).类型) = 0 Then Exit Sub
    If objReport.Items("_" & LngIdx).参照 = "" Then Call SetChildWH(LngIdx): Exit Sub
    
    '调整子表,则相应调整主表(主要指宽度与高度)
    For Each ItemThis In objReport.Items
        If ItemThis.格式号 = mbytCurrFmt And InStr(1, "4,5", ItemThis.类型) <> 0 And ItemThis.名称 = objReport.Items("_" & LngIdx).参照 Then
            Set SelObj = msh(ItemThis.Key)
            If objReport.Items("_" & LngIdx).性质 = 1 Then
                SelObj.Width = msh(LngIdx).Width
                ItemThis.W = SelObj.Width / sgnMode
            Else
                SelObj.Height = msh(LngIdx).Height
                ItemThis.H = SelObj.Height / sgnMode
            End If
            Call SetGridLine(SelObj.Index)
            Call SetChildWH(ItemThis.Key)
            Exit Sub
        End If
    Next
End Sub

Private Sub SetChildWH(ByVal LngIdx As Long)
    Dim TargetObj As Control, SelObj As Control, ItemThis As RPTItem, Int性质 As Integer
    Dim OrderTable() As Long, StrParentName As String, ArrayCount As Integer, ArrayIn As Integer, ArrayOut As Integer
    '重新设置相关子表的宽度与高度
    
    If objReport.Items("_" & LngIdx).参照 <> "" Then Exit Sub
    Set TargetObj = msh(LngIdx)
    StrParentName = objReport.Items("_" & LngIdx).名称
    ArrayCount = 0
    '获取所有子表
    For Each ItemThis In objReport.Items
        If ItemThis.格式号 = mbytCurrFmt And ItemThis.参照 = StrParentName And InStr(1, "4,5", ItemThis.类型) <> 0 Then
            ReDim Preserve OrderTable(ArrayCount)
            OrderTable(ArrayCount) = ItemThis.Key
            Int性质 = ItemThis.性质 '所有子表的性质一致
            ArrayCount = ArrayCount + 1
        End If
    Next
    
    '使数组有序
    Dim lngTmp As Long
    For ArrayOut = 0 To ArrayCount - 1
        For ArrayIn = 0 To ArrayCount - 2
            Select Case Int性质
            Case 1
                If objReport.Items("_" & OrderTable(ArrayIn + 1)).Y < objReport.Items("_" & OrderTable(ArrayIn)).Y Then
                    lngTmp = OrderTable(ArrayIn)
                    OrderTable(ArrayIn) = OrderTable(ArrayIn + 1)
                    OrderTable(ArrayIn + 1) = lngTmp
                End If
            Case 2
                If objReport.Items("_" & OrderTable(ArrayIn + 1)).X < objReport.Items("_" & OrderTable(ArrayIn)).X Then
                    lngTmp = OrderTable(ArrayIn)
                    OrderTable(ArrayIn) = OrderTable(ArrayIn + 1)
                    OrderTable(ArrayIn + 1) = lngTmp
                End If
            End Select
        Next
    Next
    
    '按序排列子表
    lngTmp = IIF(Int性质 = 1, TargetObj.Top + TargetObj.Height, TargetObj.Left + TargetObj.Width)
    For ArrayOut = 0 To ArrayCount - 1
        Set SelObj = msh(OrderTable(ArrayOut))
        SelObj.Top = IIF(Int性质 = 1, lngTmp, TargetObj.Top)
        SelObj.Left = IIF(Int性质 = 1, TargetObj.Left, lngTmp)
        If Int性质 = 1 Then '宽度一致
            SelObj.Width = TargetObj.Width
        Else
            SelObj.Height = TargetObj.Height
        End If
        Set ItemThis = objReport.Items("_" & OrderTable(ArrayOut))
        ItemThis.X = SelObj.Left / sgnMode
        ItemThis.Y = SelObj.Top / sgnMode
        ItemThis.W = SelObj.Width / sgnMode
        ItemThis.H = SelObj.Height / sgnMode
        Call SetGridLine(SelObj.Index)
        lngTmp = lngTmp + IIF(Int性质 = 1, SelObj.Height, SelObj.Width)
    Next
    
    '设置为选中状态
    Set SelObj = msh(LngIdx)
    Call AdjustSelCons(SelObj)
End Sub

Private Sub ChangeReferTo(ByVal strOldName As String, ByVal strNewName As String)
    Dim ItemThis  As RPTItem
    '修改所有的子表的参照对象
    
    For Each ItemThis In objReport.Items
        If ItemThis.参照 = strOldName And ItemThis.格式号 = mbytCurrFmt Then ItemThis.参照 = strNewName
    Next
End Sub

Private Function GetTableWidth(ByVal TargetObj As Control) As Single
    Dim ItemThis As RPTItem, SgnWidth As Single, strName As String
    '返回指定主表的宽度(含子表的宽度)
    
    SgnWidth = TargetObj.Width
    If objReport.Items("_" & TargetObj.Index).分栏 > 1 Then SgnWidth = SgnWidth * objReport.Items("_" & TargetObj.Index).分栏
    strName = objReport.Items("_" & TargetObj.Index).名称
    For Each ItemThis In objReport.Items
        If ItemThis.参照 = strName And ItemThis.格式号 = mbytCurrFmt And InStr(1, "4,5", ItemThis.类型) <> 0 Then
            If ItemThis.性质 = 2 Then SgnWidth = SgnWidth + msh(ItemThis.Key).Width
        End If
    Next
    GetTableWidth = SgnWidth
End Function

Private Function GetTableHeight(ByVal TargetObj As Control) As Single
    Dim ItemThis As RPTItem, SgnHeight As Single, strName As String
    '返回指定主表的高度(含子表的高度)
    
    SgnHeight = TargetObj.Height
    strName = objReport.Items("_" & TargetObj.Index).名称
    For Each ItemThis In objReport.Items
        If ItemThis.参照 = strName And ItemThis.格式号 = mbytCurrFmt And InStr(1, "4,5", ItemThis.类型) <> 0 Then
            If ItemThis.性质 = 1 Then SgnHeight = SgnHeight + msh(ItemThis.Key).Height
        End If
    Next
    GetTableHeight = SgnHeight
End Function

Private Sub SetRowHeight(ByVal ObjSel As VSFlexGrid)
    Dim lngRow As Long, LngRows As Long
    Dim ItemThis As RPTItem, ArrayHeight
    '设置固定行的行高
    
    ArrayHeight = Split(objReport.Items("_" & objReport.Items("_" & ObjSel.Index).SubIDs(1).Key).表头, "|")
    LngRows = UBound(ArrayHeight)
    For lngRow = 0 To LngRows
        ObjSel.RowHeight(lngRow) = Split(ArrayHeight, "^")(1) * sgnMode
    Next
End Sub

Private Function GetNextName(ByVal intType As Integer, Optional ByVal BlnTestClip As Boolean = False) As String
    Dim intMax As Integer, ItemThis As RPTItem, strName As String
    '返回一个不重复的名称
    
    intMax = 0
    For Each ItemThis In objReport.Items
        If ItemThis.格式号 = mbytCurrFmt And ItemThis.类型 = intType Then
            strName = Val(StrReverse(Val(StrReverse(ItemThis.名称 & "9"))))
            strName = Mid(strName, 1, Len(strName) - 1)
            If intMax < CInt(Val(strName)) Then intMax = CInt(Val(strName))
        End If
    Next
    If BlnTestClip Then
        For Each ItemThis In objClip
            If ItemThis.类型 = intType Then
                strName = Val(StrReverse(Val(StrReverse(ItemThis.名称 & "9"))))
                strName = Mid(strName, 1, Len(strName) - 1)
                If intMax < CInt(Val(strName)) Then intMax = CInt(Val(strName))
            End If
        Next
    End If
    intMax = intMax + 1
    
    Select Case intType
    Case 1
        GetNextName = "线条"
    Case 2, 3
        GetNextName = "标签"
    Case 4
        GetNextName = "任意表"
    Case 5
        GetNextName = "汇总表"
    Case 10
        GetNextName = "框线"
    Case 11
        GetNextName = "图片"
    Case 12 '@@@
        GetNextName = "图表"
    Case 13
        GetNextName = "条码"
    Case 14
        GetNextName = "卡片"
    End Select
    GetNextName = GetNextName & intMax
End Function

Private Function UserRefer(ByVal LngIdx As Long) As Boolean
    Dim ItemThis As RPTItem
    '如果是其它表格的参照对象，则不允许删除
    UserRefer = True
    
    For Each ItemThis In objReport.Items
        If ItemThis.格式号 = mbytCurrFmt And ItemThis.Key <> LngIdx And ItemThis.参照 = objReport.Items("_" & LngIdx).名称 Then Exit Function
    Next
    
    UserRefer = False
End Function

Private Sub GetAllElement()
    Dim ItemThis As RPTItem
    
    '获取当前报表格式中的所有报表元素
    cboAtt.Clear
    cboAtt.AddItem ""
    
    For Each ItemThis In objReport.Items
        If ItemThis.格式号 = mbytCurrFmt And InStr(1, "|1,|2,|3,|4,|5,|10,|11,|12,|14,", "|" & ItemThis.类型 & ",") <> 0 Then
            cboAtt.AddItem ItemThis.名称
            cboAtt.ItemData(cboAtt.NewIndex) = ItemThis.Key
        End If
    Next
    cboAtt.ListIndex = 0
End Sub

Private Sub LoadOutChart()
    With cboAtt
        .Clear
        .AddItem "禁止输出"
        .ItemData(.NewIndex) = 0
        .AddItem "面积图"
        .ItemData(.NewIndex) = 1        'xlArea
        .AddItem "折线图"
        .ItemData(.NewIndex) = 4        'xlLine
        .AddItem "饼图"
        .ItemData(.NewIndex) = 5        'xlPie
        .AddItem "气泡图"
        .ItemData(.NewIndex) = 15       'xlBubble
        .AddItem "柱形图"
        .ItemData(.NewIndex) = 51       'xlColumnClustered
        .AddItem "条形图"
        .ItemData(.NewIndex) = 57       'xlBarClustered
        .AddItem "曲面图"
        .ItemData(.NewIndex) = 83       'xlSurface
        .AddItem "股价图"
        .ItemData(.NewIndex) = 88       'xlStockHLC
        .AddItem "圆柱图"
        .ItemData(.NewIndex) = 92       'xlCylinderColClustered
        .AddItem "圆锥图"
        .ItemData(.NewIndex) = 99       'xlConeColClustered
        .AddItem "棱锥图"
        .ItemData(.NewIndex) = 106      'xlPyramidColClustered
        .AddItem "散点图"
        .ItemData(.NewIndex) = -4169    'xlXYScatter
        .AddItem "圆环图"
        .ItemData(.NewIndex) = -4120    'xlDoughnut
        .AddItem "雷达图"
        .ItemData(.NewIndex) = -4151    'xlRadar
        .AddItem "三维柱形图"
        .ItemData(.NewIndex) = -4100    'xl3DColumn
        .ListIndex = 0
    End With
End Sub

Private Function GetCurOutChart() As String
    Select Case objReport.Fmts("_" & mbytCurrFmt).图样
        Case 0
            GetCurOutChart = "禁止输出"
        Case 1        'xlArea
            GetCurOutChart = "面积图"
        Case 4        'xlLine
            GetCurOutChart = "折线图"
        Case 5        'xlPie
            GetCurOutChart = "饼图"
        Case 15       'xlBubble
            GetCurOutChart = "气泡图"
        Case 51       'xlColumnClustered
            GetCurOutChart = "柱形图"
        Case 57       'xlBarClustered
            GetCurOutChart = "条形图"
        Case 83       'xlSurface
            GetCurOutChart = "曲面图"
        Case 88       'xlStockHLC
            GetCurOutChart = "股价图"
        Case 92       'xlCylinderColClustered
            GetCurOutChart = "圆柱图"
        Case 99       'xlConeColClustered
            GetCurOutChart = "圆锥图"
        Case 106      'xlPyramidColClustered
            GetCurOutChart = "棱锥图"
        Case -4169    'xlXYScatter
            GetCurOutChart = "散点图"
        Case -4120    'xlDoughnut
            GetCurOutChart = "圆环图"
        Case -4151    'xlRadar
            GetCurOutChart = "雷达图"
        Case -4100    'xl3DColumn
            GetCurOutChart = "三维柱形图"
    End Select
End Function

Private Sub LocateOutChart()
    Dim i As Integer, IntOutChart As Integer
    '显示该样式所对应的输出模式
    
    blnModify = False
    IntOutChart = objReport.Fmts("_" & mbytCurrFmt).图样
    For i = 0 To cboAtt.ListCount - 1
        If cboAtt.ItemData(i) = IntOutChart Then
            cboAtt.ListIndex = i
            Exit For
        End If
    Next
    If cboAtt.ListIndex < 0 Then cboAtt.ListIndex = 0
    If cboAtt.ItemData(cboAtt.ListIndex) <> IntOutChart Then cboAtt.ListIndex = 0
    blnModify = True
End Sub

Private Function CheckText(ByVal tmpItem As RPTItem, Optional ByVal blnHead As Boolean = False) As Boolean
    Dim objNode As Node, arrData As Variant
    Dim strFind As String, intFind As Integer, blnFind As Boolean
    
    CheckText = False
    strFind = GetLabelDataName(IIF(blnHead, tmpItem.表头, tmpItem.内容))
    arrData = Split(strFind, "|")
    
    For intFind = 0 To UBound(arrData)
        '每个数据源必须匹配，否则退出
        blnFind = False
        For Each objNode In tvwSQL.Nodes
            If objNode.Key <> "Root" And objNode.Children = 0 Then
                If InStr(1, arrData(intFind), LevelText(objNode)) <> 0 Then blnFind = True: Exit For
            End If
        Next
        If blnFind = False Then Exit Function
    Next
    CheckText = True
End Function

Private Sub GetInPaper()
    Dim i As Integer, j As Integer, k As Integer
    Dim IntPaper As Integer, strTmp As String
    Dim strPaperBinName As String * 1000, strPaperBin As String * 100
    '--------------------------------------------------------------------------------------------
    
    If Printers.count = 0 Then
        MsgBox "在系统中没有检测到任何打印设备,请尽快添加,否则部分操作不能正常执行。", vbInformation, App.Title
        Exit Sub
    End If
    
    IntPaper = DeviceCapabilities(Printer.DeviceName, Printer.Port, DC_BINNAMES, strPaperBinName, 0)
    IntPaper = DeviceCapabilities(Printer.DeviceName, Printer.Port, DC_BINS, strPaperBin, 0)
    
    CboTest.Clear
    j = 1
    For i = 1 To IntPaper
        k = 0
        '进纸名称
        Do
            If Mid(strPaperBinName, j, 1) = Chr(0) Then
                If Trim(strTmp) <> "" Then
                    CboTest.AddItem Trim(strTmp)
                    
                    '进纸编号
                    CboTest.ItemData(CboTest.ListCount - 1) = Asc(Mid(strPaperBin, i * 2, 1)) * 256# + Asc(Mid(strPaperBin, i * 2 - 1, 1))
                    If CboTest.ItemData(CboTest.ListCount - 1) = objReport.进纸 Then
                        CboTest.ListIndex = CboTest.ListCount - 1 '定位在原设置上
                    End If
                    If CboTest.ListIndex = -1 And CboTest.ItemData(CboTest.ListCount - 1) = Printer.PaperBin Then
                        CboTest.ListIndex = CboTest.ListCount - 1 '定位在打印机缺省设置上
                    End If
                End If
                
                j = 24 + j - LenB(StrConv(strTmp, vbFromUnicode))
                strTmp = ""
                Exit Do
            Else
                strTmp = strTmp & Mid(strPaperBinName, j, 1)
                j = j + 1
                k = k + 1
                If k > 24 Then Exit Do
            End If
        Loop
    Next
    '--------------------------------------------------------------------------------------------
    If CboTest.ListIndex = -1 And CboTest.ListCount > 0 Then CboTest.ListIndex = 0
End Sub

Private Function ReferObj(ByVal LngIdx As Long) As Boolean
    Dim ItemThis As RPTItem, StrMainObj As String
    '检测当前元素是否是主表或子表
    
    ReferObj = False
    StrMainObj = objReport.Items("_" & LngIdx).名称
    If objReport.Items("_" & LngIdx).参照 <> "" Then ReferObj = True: Exit Function
    For Each ItemThis In objReport.Items
        If ItemThis.格式号 = mbytCurrFmt And ItemThis.参照 = StrMainObj And ItemThis.类型 = 5 Then
            ReferObj = True
            Exit Function
        End If
    Next
End Function

Private Function CheckCoordinate() As Long
    Dim ItemCheck As RPTItem, blnCheck As Boolean
    Dim RectCheck As RECT
    '检查所有系统固有项目，如果被遮住，则不允许保存
    
    CheckCoordinate = 0
    For Each ItemCheck In objReport.Items
        If ItemCheck.系统 Then
            RectCheck.Left = ItemCheck.X
            RectCheck.Top = ItemCheck.Y
            RectCheck.Right = ItemCheck.W
            RectCheck.Bottom = ItemCheck.H
            blnCheck = GetCoordinate(RectCheck, ItemCheck.Key, True)
            If blnCheck Then CheckCoordinate = ItemCheck.Key: Exit Function
        End If
    Next
End Function

Private Function GetCoordinate(ByRef Area As RECT, Optional ByVal IntStyle As Integer = 1, _
Optional ByVal lngKey As Long, Optional ByVal blnCheck As Boolean = False) As Boolean
'功能：自动定位（当粘贴系统固有项目时）
'返回：选中的个数
    Dim tmpItem As RPTItem, ObjSel As Object, LngLoop As Long
    Dim ObjLeft As Single, ObjTop As Single, ObjHeight As Single, ObjWidth As Single
    
    For LngLoop = 1 To objReport.Items.count
        Set tmpItem = objReport.Items(LngLoop)
        If InStr(1, "|1,|2,|3,|4,|5,|10,|11,|12,|14,", "|" & tmpItem.类型) <> 0 _
            And Mid(cboFormat.ComboItems("_" & mbytCurrFmt).Key, 2) = tmpItem.格式号 And tmpItem.Key <> lngKey Then
            Set ObjSel = GetInxObj(tmpItem.id)
            
            ObjTop = objReport.Items("_" & tmpItem.id).Y
            ObjLeft = objReport.Items("_" & tmpItem.id).X
            ObjWidth = objReport.Items("_" & tmpItem.id).W
            ObjHeight = objReport.Items("_" & tmpItem.id).H
            
            If Not (ObjTop > Area.Bottom + Area.Top Or _
                ObjLeft > Area.Right + Area.Left Or _
                ObjTop + ObjHeight < Area.Top Or _
                ObjLeft + ObjWidth < Area.Left) Then
                If blnCheck Then GetCoordinate = True: Exit Function
                If IntStyle = 1 Then
                    Area.Top = ObjTop + ObjHeight + 100
                Else
                    Area.Top = ObjTop - Area.Bottom - 100
                End If
                LngLoop = 0
            End If
        End If
    Next
End Function

Private Function AdjustCoordinate(Optional ByVal BlnIn As Boolean)
'blnIn=设置容器时调用
    Dim ObjSel As Object, RectTest As RECT
    
    If objReport.Items("_" & intCurID).性质 <> 0 Or BlnIn Then
        RectTest.Left = objReport.Items("_" & intCurID).X
        RectTest.Top = objReport.Items("_" & intCurID).Y
        RectTest.Bottom = objReport.Items("_" & intCurID).H
        RectTest.Right = objReport.Items("_" & intCurID).W
        If BlnIn = False Then
            Call GetCoordinate(RectTest, IIF(Mid(objReport.Items("_" & intCurID).性质, 1, 1) = 1, 2, 1), intCurID)
        
            objReport.Items("_" & intCurID).X = RectTest.Left
            objReport.Items("_" & intCurID).Y = RectTest.Top
        End If
        Set ObjSel = GetInxObj(intCurID)
        
        With ObjSel
            .Left = RectTest.Left * sgnMode
            .Top = RectTest.Top * sgnMode
            If objReport.Items("_" & intCurID).父ID = 0 Then
                Set ObjSel.Container = picPaper
            Else
                Set ObjSel.Container = pic(objReport.Items("_" & intCurID).父ID)
            End If
        End With
        
        Call AdjustSelCons(ObjSel)
    End If
End Function

Private Function CheckClip()
    Dim ItemCheck As RPTItem, StrMainObj As String, ArrayRefer
    Dim i As Integer, blnFind As Boolean
    '如果粘贴控件中无主表,则清除所有子控件的参照对象及性质
    
    StrMainObj = ","
    For Each ItemCheck In objClip
        If ItemCheck.参照 <> "" And InStr(1, StrMainObj, ItemCheck.参照) = 0 Then
            StrMainObj = StrMainObj & IIF(StrMainObj = ",", "", ",") & ItemCheck.参照
        End If
    Next
    For Each ItemCheck In objClip
        If ItemCheck.类型 = 2 Then ItemCheck.参照 = "": ItemCheck.性质 = 0
    Next
    
    StrMainObj = Mid(StrMainObj, 2)
    ArrayRefer = Split(StrMainObj, ",")
    
    For i = 0 To UBound(ArrayRefer)
        blnFind = False
        For Each ItemCheck In objClip
            If ItemCheck.参照 = "" And ArrayRefer(i) = ItemCheck.名称 Then
                blnFind = True
                Exit For
            End If
        Next
        If blnFind = False Then
            '清除所有子表的参照对象及性质
            For Each ItemCheck In objClip
                If ItemCheck.参照 = ArrayRefer(i) Then ItemCheck.参照 = "": ItemCheck.性质 = 0
            Next
        End If
    Next
End Function

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

Private Function GetDependID(strName As String) As Integer
'功能：根据参照名称,获取其索引.
    Dim objItem As RPTItem
    
    For Each objItem In objReport.Items
        If objItem.格式号 = mbytCurrFmt And objItem.名称 = strName _
            And (objItem.类型 = 4 Or objItem.类型 = 5) And objItem.性质 = 0 Then
            GetDependID = objItem.id: Exit Function
        End If
    Next
End Function

