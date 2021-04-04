VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearchPatient 
   Caption         =   "病案查找"
   ClientHeight    =   7845
   ClientLeft      =   2835
   ClientTop       =   3825
   ClientWidth     =   15015
   Icon            =   "frmSearchPatient.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   15015
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   1290
      Index           =   4
      Left            =   4800
      ScaleHeight     =   1290
      ScaleWidth      =   3555
      TabIndex        =   32
      Top             =   6165
      Width           =   3555
      Begin VB.CommandButton cmd 
         Height          =   300
         Index           =   6
         Left            =   3165
         Picture         =   "frmSearchPatient.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "删除条件"
         Top             =   705
         Width           =   300
      End
      Begin VB.CommandButton cmd 
         Height          =   300
         Index           =   5
         Left            =   3165
         Picture         =   "frmSearchPatient.frx":685E
         Style           =   1  'Graphical
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "更新条件"
         Top             =   360
         Width           =   300
      End
      Begin VB.CommandButton cmd 
         Height          =   300
         Index           =   4
         Left            =   3165
         Picture         =   "frmSearchPatient.frx":D0B0
         Style           =   1  'Graphical
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   "新增条件"
         Top             =   15
         Width           =   300
      End
      Begin MSComctlLib.ListView lvw 
         Height          =   1290
         Left            =   0
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   0
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2275
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         Icons           =   "ils16"
         SmallIcons      =   "ils16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   4974
         EndProperty
      End
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   5160
      Index           =   3
      Left            =   3525
      ScaleHeight     =   5160
      ScaleWidth      =   5820
      TabIndex        =   21
      Top             =   570
      Width           =   5820
      Begin VB.ComboBox cmbLogical 
         Height          =   300
         ItemData        =   "frmSearchPatient.frx":13902
         Left            =   0
         List            =   "frmSearchPatient.frx":13904
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   2475
         Width           =   930
      End
      Begin VB.ListBox lstFields 
         Height          =   2400
         Left            =   0
         TabIndex        =   41
         ToolTipText     =   "打了勾的那一行会出现在查询结果中"
         Top             =   45
         Width           =   5280
      End
      Begin VB.OptionButton opt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "满足任一条件(&P)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   30
         TabIndex        =   28
         Top             =   4905
         Width           =   1650
      End
      Begin VB.OptionButton opt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "满足全部条件(&A)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   30
         TabIndex        =   27
         Top             =   4665
         Value           =   -1  'True
         Width           =   1650
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "刷新(&R)"
         Height          =   350
         Left            =   4185
         TabIndex        =   29
         Top             =   4635
         Width           =   1100
      End
      Begin VB.ComboBox cmbOperate 
         Height          =   300
         Left            =   945
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   2475
         Width           =   1005
      End
      Begin VB.ComboBox cmbList 
         Height          =   300
         Left            =   1965
         TabIndex        =   24
         Top             =   2475
         Width           =   3315
      End
      Begin VB.CommandButton cmd 
         Height          =   300
         Index           =   1
         Left            =   5310
         Picture         =   "frmSearchPatient.frx":13906
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2805
         Width           =   300
      End
      Begin VB.CommandButton cmd 
         Height          =   300
         Index           =   2
         Left            =   5310
         Picture         =   "frmSearchPatient.frx":1A158
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   3135
         Width           =   300
      End
      Begin MSComctlLib.ListView lvwCombine 
         Height          =   1800
         Left            =   15
         TabIndex        =   42
         Top             =   2805
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   3175
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "对象"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "条件"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.ComboBox cmbExample 
         Height          =   300
         ItemData        =   "frmSearchPatient.frx":209AA
         Left            =   1965
         List            =   "frmSearchPatient.frx":209AC
         TabIndex        =   43
         Top             =   2475
         Width           =   3300
      End
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2820
      Index           =   2
      Left            =   150
      ScaleHeight     =   2820
      ScaleWidth      =   4365
      TabIndex        =   20
      Top             =   3990
      Width           =   4365
      Begin VB.CommandButton cmdSearch 
         Caption         =   "刷新(&R)"
         Height          =   350
         Left            =   2850
         TabIndex        =   37
         Top             =   2460
         Width           =   1100
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   2
         Left            =   1065
         TabIndex        =   16
         Top             =   2100
         Width           =   2850
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1065
         TabIndex        =   14
         Top             =   1740
         Width           =   2850
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   6
         Left            =   1065
         TabIndex        =   8
         Top             =   1050
         Width           =   1050
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   0
         Left            =   1065
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   2850
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   4
         Left            =   2355
         TabIndex        =   10
         Top             =   1035
         Width           =   1365
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   1065
         TabIndex        =   1
         Top             =   15
         Width           =   2850
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   1
         Left            =   1065
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1380
         Width           =   2850
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   2
         Left            =   1065
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   705
         Width           =   2850
      End
      Begin VB.CommandButton cmd 
         Height          =   300
         Index           =   0
         Left            =   3930
         Picture         =   "frmSearchPatient.frx":209AE
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   15
         Width           =   300
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病 案 号(&H)"
         Height          =   180
         Index           =   4
         Left            =   45
         TabIndex        =   15
         Top             =   2145
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住 院 号(&Z)"
         Height          =   180
         Index           =   1
         Left            =   45
         TabIndex        =   13
         Top             =   1800
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病    种(&B)"
         Height          =   180
         Index           =   0
         Left            =   45
         TabIndex        =   0
         Top             =   75
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年    龄(&N)"
         Height          =   180
         Index           =   8
         Left            =   45
         TabIndex        =   7
         Top             =   1110
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "婚姻状况(&I)"
         Height          =   180
         Index           =   6
         Left            =   45
         TabIndex        =   11
         Top             =   1440
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性    别(&X)"
         Height          =   180
         Index           =   5
         Left            =   45
         TabIndex        =   3
         Top             =   420
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "科    室(&K)"
         Height          =   180
         Index           =   3
         Left            =   45
         TabIndex        =   5
         Top             =   750
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Index           =   2
         Left            =   2145
         TabIndex        =   9
         Top             =   1095
         Width           =   180
      End
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   2355
      Index           =   1
      Left            =   8490
      ScaleHeight     =   2355
      ScaleWidth      =   3555
      TabIndex        =   18
      Top             =   4320
      Width           =   3555
      Begin VSFlex8Ctl.VSFlexGrid vsfPatient 
         Height          =   1350
         Left            =   180
         TabIndex        =   19
         Top             =   240
         Width           =   2535
         _cx             =   4471
         _cy             =   2381
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
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   12698049
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
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
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   3000
      Index           =   0
      Left            =   150
      ScaleHeight     =   3000
      ScaleWidth      =   3030
      TabIndex        =   17
      Top             =   615
      Width           =   3030
      Begin XtremeSuiteControls.TaskPanel tpl 
         Height          =   2190
         Left            =   225
         TabIndex        =   30
         Top             =   405
         Width           =   1965
         _Version        =   589884
         _ExtentX        =   3466
         _ExtentY        =   3863
         _StockProps     =   64
         ItemLayout      =   2
         HotTrackStyle   =   1
      End
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   10170
      Top             =   1005
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearchPatient.frx":27200
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearchPatient.frx":2751A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearchPatient.frx":2836C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearchPatient.frx":2DB5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearchPatient.frx":2E570
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearchPatient.frx":3405A
            Key             =   "Query"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearchPatient.frx":3815C
            Key             =   "Attrib"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   31
      Top             =   7485
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   23574
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
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
            Picture         =   "frmSearchPatient.frx":3E9BE
            Key             =   "Record"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearchPatient.frx":45220
            Key             =   "RecordPart"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearchPatient.frx":4BA82
            Key             =   "RecordNO"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearchPatient.frx":522E4
            Key             =   "Page"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearchPatient.frx":52738
            Key             =   "Attrib"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   3315
      Index           =   5
      Left            =   9810
      ScaleHeight     =   3315
      ScaleWidth      =   3630
      TabIndex        =   38
      Top             =   450
      Width           =   3630
      Begin MSComctlLib.ListView lvw主页 
         Height          =   2835
         Left            =   15
         TabIndex        =   39
         Top             =   30
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   5001
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ils16"
         SmallIcons      =   "ils16"
         ForeColor       =   -2147483640
         BackColor       =   16446707
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "信息名"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "信息值"
            Object.Width           =   4410
         EndProperty
      End
      Begin MSComctlLib.TabStrip tabMain 
         Height          =   1335
         Left            =   30
         TabIndex        =   40
         Top             =   1905
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   2355
         MultiRow        =   -1  'True
         Placement       =   1
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "第1次入院"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.Image imgCelNo 
      Height          =   240
      Index           =   0
      Left            =   9495
      Picture         =   "frmSearchPatient.frx":5284A
      Top             =   3000
      Width           =   240
   End
   Begin VB.Image imgCelNo 
      Height          =   240
      Index           =   1
      Left            =   9540
      Picture         =   "frmSearchPatient.frx":52B8C
      Top             =   3495
      Width           =   240
   End
   Begin VB.Image imgCelNo 
      Height          =   240
      Index           =   2
      Left            =   9510
      Picture         =   "frmSearchPatient.frx":52ECE
      Top             =   3225
      Width           =   240
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmSearchPatient.frx":53210
      Left            =   1035
      Top             =   180
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmSearchPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'######################################################################################################################
Private mblnOK As Boolean
Private mfrmMain As Object
Private WithEvents mclsPatient As clsVsf
Attribute mclsPatient.VB_VarHelpID = -1
Private mrsPatient As ADODB.Recordset
Private mblnDataChanged As Boolean
Private mblnConditionChanged As Boolean
Private mlngLoop As Long
Private mlngMoudal As Long
Private mstrPrivs As String
Private Type Items
    疾病名称 As String
End Type

Private mlngOldRow As Long
Private mlngNewRow As Long
Private mstrFilter As String
Private mstr显示 As String
Private mbln病案系统 As Boolean '检查病案系统是否存在
Private mlngIndex As Long
Private mint简码方式 As Integer

Private mstr条件 As String
Private mstrReturn As String
Private mstrSQL As String
Private mstrExecel As String

Private usrSaveItem As Items

'######################################################################################################################

Public Function ShowEdit(ByVal frmMain As Object, ByRef rsPatient As ADODB.Recordset, ByVal lngMoudal As Long, ByVal strPrivs As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mlngMoudal = lngMoudal
    mblnOK = False
    mstrPrivs = strPrivs
    
    Set mfrmMain = frmMain
    If ExecuteCommand("初始控件") = False Then Exit Function
    If ExecuteCommand("初始数据") = False Then Exit Function
    
    DataChanged = False
    
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK
    If mblnOK Then
        Set rsPatient = CopyRecordStruct(mrsPatient)
        Call CopyRecordData(mrsPatient, rsPatient)
    End If
    
End Function

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    
    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    
    Call CommandBarInit(cbsMain)
    
    '------------------------------------------------------------------------------------------------------------------
    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值
    
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    '文件
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Preview, "预览(&V)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Print, "打印(&P)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Excel, "输出到&Excel")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Exit, "退出(&X)", True)
    
    '编辑
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_SelAll, "全部选择(&A)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_ClsAll, "全部不选(&D)")
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_SaveExit, "保存选择(&S)", True)
       
    '查看
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Refresh, "刷新(&R)", True)

            
    '帮助
    '------------------------------------------------------------------------------------------------------------------
    Call CreateHelpMenu(cbsMain)
    
    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义:包括公共部份

    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched

    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Print, "打印")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Preview, "预览")
    
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_SelAll, "全选(&A)", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_ClsAll, "全清(&D)")
    
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_SaveExit, "保存(&S)", True)

    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Help_Help, "帮助(&H)", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "退出(&X)")

    '------------------------------------------------------------------------------------------------------------------
    '命令的快键绑定:公共部份主界面已处理

    With cbsMain.KeyBindings
        .Add 0, vbKeyF5, conMenu_View_Refresh           '刷新
        .Add 0, vbKeyF1, conMenu_Help_Help              '帮助
        .Add 0, vbKeyF2, conMenu_Edit_SaveExit              '保存
        .Add FCONTROL, vbKeyA, conMenu_Edit_SelAll
    End With

End Function

Private Sub InitDockPannel()
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim objPane As Pane

    Set objPane = dkpMain.CreatePane(1, 100, 100, DockLeftOf, Nothing): objPane.Title = "条件": objPane.Options = PaneNoCaption
    Set objPane = dkpMain.CreatePane(2, 120, 100, DockRightOf, Nothing): objPane.Title = "病人": objPane.Options = PaneNoCaption
    Set objPane = dkpMain.CreatePane(3, 80, 100, DockRightOf, Nothing): objPane.Title = "信息": objPane.Options = PaneNoCaption
    
    dkpMain.SetCommandBars cbsMain
    Call DockPannelInit(dkpMain)

End Sub

Private Function InitToolBox() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    Dim objGrp As TaskPanelGroup
    Dim objItem As TaskPanelGroupItem
    Dim objIlsItem As Object
    

    '基本条件组
    Set objGrp = tpl.Groups.Add(0, "基本条件：")
    objGrp.Expandable = False
    objGrp.Expanded = True
    Set objItem = objGrp.Items.Add(0, "无", xtpTaskItemTypeControl, 1)
    objItem.Handle = picPane(2).hWnd
    picPane(2).BackColor = objItem.BackColor
    
    '高级条件
    Set objGrp = tpl.Groups.Add(1, "高级条件：")
    objGrp.Expandable = False
    objGrp.Expanded = True
    Set objItem = objGrp.Items.Add(0, "无", xtpTaskItemTypeControl, 1)
    objItem.Handle = picPane(3).hWnd
    picPane(3).BackColor = objItem.BackColor
    opt(0).BackColor = objItem.BackColor
    opt(1).BackColor = objItem.BackColor
    
    '历史条件
    Set objGrp = tpl.Groups.Add(1, "保存条件：")
    Set objItem = objGrp.Items.Add(0, "无", xtpTaskItemTypeControl, 1)
    objItem.Handle = picPane(4).hWnd
    picPane(4).BackColor = objItem.BackColor
    
    tpl.Animation = xtpTaskPanelAnimationNo
    tpl.VisualTheme = xtpTaskPanelThemeNativeWinXPPlain
    tpl.AllowDrag = False
    tpl.SelectItemOnFocus = True
    
    Call tpl.SetGroupOuterMargins(5, 5, 5, 5)
    Call tpl.SetMargins(1, 1, 1, 1, 1)
    
    
    InitToolBox = True
    
End Function

Private Sub InitListPatient()
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
     '初始化窗口控件的值
    
    '0  不处理
    '1  布尔型
    '2  数字
    '3  字符
    '4  日期
    '5  固定取值
    lstFields.Clear
    '一、常用检索条件
    lstFields.AddItem "住院号":   lstFields.ItemData(lstFields.NewIndex) = 2 ' Number(18)"
    lstFields.AddItem "病案号":   lstFields.ItemData(lstFields.NewIndex) = 3 ' varchar2(20)"
    lstFields.AddItem "档案号":   lstFields.ItemData(lstFields.NewIndex) = 3 ' varchar2(20)"
    lstFields.AddItem "住院次数": lstFields.ItemData(lstFields.NewIndex) = 2 ' Number(18)"
    lstFields.AddItem "姓名":     lstFields.ItemData(lstFields.NewIndex) = 3     ' VARCHAR2(10)"
    lstFields.AddItem "姓名简码": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(10)"
    lstFields.AddItem "出院日期": lstFields.ItemData(lstFields.NewIndex) = 4 ' Date"
    lstFields.AddItem "出院科室": lstFields.ItemData(lstFields.NewIndex) = 3 ' Number(18)"
    lstFields.AddItem "编目员姓名": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(10)"
    lstFields.AddItem "编目日期": lstFields.ItemData(lstFields.NewIndex) = 4 ' Date"
    '二、诊断条件
    lstFields.AddItem " "
    lstFields.AddItem "诊断类型": lstFields.ItemData(lstFields.NewIndex) = 5
    lstFields.AddItem "诊断编码": lstFields.ItemData(lstFields.NewIndex) = 3 ' Number(5)"
    lstFields.AddItem "诊断简码": lstFields.ItemData(lstFields.NewIndex) = 3
    lstFields.AddItem "诊断描述信息": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(6)"
    lstFields.AddItem "诊断出院情况": lstFields.ItemData(lstFields.NewIndex) = 3  ' VARCHAR2
    lstFields.AddItem "诊断次序": lstFields.ItemData(lstFields.NewIndex) = 2 ' Number
    lstFields.AddItem "诊断编码序号": lstFields.ItemData(lstFields.NewIndex) = 2 ' Number
'    If gSystemPara.bln购买中医 = True Then
        lstFields.AddItem "中医候诊": lstFields.ItemData(lstFields.NewIndex) = 3  ' VARCHAR2
        lstFields.AddItem "中医治疗类别": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2
'    End If
    lstFields.AddItem "未治": lstFields.ItemData(lstFields.NewIndex) = 1
    lstFields.AddItem "疑诊": lstFields.ItemData(lstFields.NewIndex) = 1
    lstFields.AddItem "符合类型": lstFields.ItemData(lstFields.NewIndex) = 5
    lstFields.AddItem "符合情况": lstFields.ItemData(lstFields.NewIndex) = 5
   
    '三、手术条件
    lstFields.AddItem " "
    lstFields.AddItem "手术编码": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2
    lstFields.AddItem "手术简码": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2
    lstFields.AddItem "手术已行手术": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(6)"
    lstFields.AddItem "手术切口": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2
    lstFields.AddItem "手术愈合": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2
    lstFields.AddItem "手术麻醉类型": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2
    lstFields.AddItem "手术日期": lstFields.ItemData(lstFields.NewIndex) = 4 ' VARCHAR2(10)"
    lstFields.AddItem "主刀医师": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(50)"
    lstFields.AddItem "麻醉医师": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(50)"
    '四、病人信息
    lstFields.AddItem " "
    lstFields.AddItem "身份证号": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(18)"
    lstFields.AddItem "年龄": lstFields.ItemData(lstFields.NewIndex) = 2 ' VARCHAR2(10)"
    lstFields.AddItem "性别": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(4)"
    lstFields.AddItem "婚姻状况": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(4)"
    lstFields.AddItem "医疗付款方式": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(20)"
    lstFields.AddItem "职业": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(40)"
    lstFields.AddItem "国籍": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(30)"
    lstFields.AddItem "血型": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(10)"
    lstFields.AddItem "单位电话": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(50)"
    lstFields.AddItem "单位邮编": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(6)"
    lstFields.AddItem "单位地址": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(50)"
    lstFields.AddItem "区域": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(10)"
    lstFields.AddItem "家庭地址": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(50)"
    lstFields.AddItem "家庭电话": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(20)"
    lstFields.AddItem "户口邮编": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(6)"
    lstFields.AddItem "联系人姓名": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(10)"
    lstFields.AddItem "联系人关系": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(10)"
    lstFields.AddItem "联系人地址": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(50)"
    lstFields.AddItem "联系人电话": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(20)"
     lstFields.AddItem "出生日期": lstFields.ItemData(lstFields.NewIndex) = 4 ' date"
    '五、病案主页
    lstFields.AddItem " "
    lstFields.AddItem "入院科室": lstFields.ItemData(lstFields.NewIndex) = 3 ' Number(18)"
    lstFields.AddItem "入院日期": lstFields.ItemData(lstFields.NewIndex) = 4 ' Date"
    lstFields.AddItem "入院病况": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(10)"
    lstFields.AddItem "住院天数": lstFields.ItemData(lstFields.NewIndex) = 2 ' Number(18)"
    lstFields.AddItem "确诊日期": lstFields.ItemData(lstFields.NewIndex) = 4 ' Date"
    lstFields.AddItem "随诊标志": lstFields.ItemData(lstFields.NewIndex) = 1
    lstFields.AddItem "随诊期限": lstFields.ItemData(lstFields.NewIndex) = 2 ' Number(18)"
    lstFields.AddItem "抢救次数": lstFields.ItemData(lstFields.NewIndex) = 2 ' Number(5)"
    lstFields.AddItem "成功次数": lstFields.ItemData(lstFields.NewIndex) = 2 ' Number(5)"
    lstFields.AddItem "尸检标志": lstFields.ItemData(lstFields.NewIndex) = 1
    lstFields.AddItem "费用和":   lstFields.ItemData(lstFields.NewIndex) = 2 ' Number(16, 5)"
    lstFields.AddItem "过敏药物": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(50)"
    lstFields.AddItem "病案质量": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(2)"
    lstFields.AddItem "科主任":   lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2
    lstFields.AddItem "主任医师": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2
    lstFields.AddItem "医保号":   lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2
    lstFields.AddItem "主治医师": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2
    lstFields.AddItem "住院医师": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2
    lstFields.AddItem "门诊医师": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2
    '问题31031 by lesfeng 2010-06-24 解决访问病案主页从表的性能，见下屏蔽的代码
    mlngIndex = lstFields.NewIndex
    
    lstFields.AddItem "进修医师": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "研究生实习医师": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "实习医师": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "质控医师": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "质控护士": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    '六、输血情况
    lstFields.AddItem " "
    lstFields.AddItem "HBsAg": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "HCV-Ab": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "HIV-Ab": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "Rh": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "输血检查": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "输液反应": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "输血反应": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "输红细胞": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "输其他": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "输血浆": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "输全血": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "输血小板": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    
'    If gSystemPara.bln购买中医 = True Then
        lstFields.AddItem " "
        lstFields.AddItem "中医抢救方法": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
        lstFields.AddItem "自制中药制剂": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
        lstFields.AddItem "中医危重": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
        lstFields.AddItem "中医疑难": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
        lstFields.AddItem "中医急症": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
'    End If

    '七、病室情况
    lstFields.AddItem " "
    lstFields.AddItem "入院病室": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "出院病室": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "收回日期": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "病理号": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "科研病案": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "示教病案": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "入院前经外院治疗": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "疑难病历": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "首例": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "转科记录": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "转科时间": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "病原学检查": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "出院方式": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "死亡根本原因": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "入院方式": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "感染与死亡关系": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "感染部位": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "离院方式_附页": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "转入机构名称": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "抢救病因": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "主页X线号": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "主页质量日期": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "呼吸机使用时间": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "昏迷时间": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "特级护理天数": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "一级护理天数": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "二级护理天数": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "三级护理天数": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "ICU天数": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "CCU天数": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "CT": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "MRI": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "彩色多普勒": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "特殊检查4": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "特殊检查5": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "特殊检查6": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    '八、分娩情况
    lstFields.AddItem " "
    lstFields.AddItem "分娩时间": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "产检次数": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "胎次": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "胎数": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "产程时间1": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "产程时间2": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "总产程时间": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "产后出血量": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "产科并发症": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "会阴Ⅲ度裂伤": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "住院死亡期间": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "分化程度": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "最高诊断依据": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "不足周岁年龄": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "新生儿出生体重": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "新生儿入院体重": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    
    '九、借阅情况
    lstFields.AddItem ""
    lstFields.AddItem "借阅人": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "借出人": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "批准人": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "借阅日期": lstFields.ItemData(lstFields.NewIndex) = 4 ' DATE"
    lstFields.AddItem "借阅部门": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "借阅用途": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    
    On Error GoTo errHandle
    strSQL = "Select Distinct(名称) 名称 From 病案项目" '自定义的数据不是很多，没有条件
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)

    lstFields.AddItem " "

    While Not rsTemp.EOF
        lstFields.AddItem rsTemp!名称: lstFields.ItemData(lstFields.NewIndex) = 3
        rsTemp.MoveNext
    Wend
    
    lstFields.ListIndex = 0
'    Call lstFields_Click
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub lstFields_Click()
    Dim strExample As String
    Dim lng值类型 As Long
    Dim str字段名 As String

    If cmbOperate.ListCount = 7 Then cmbOperate.RemoveItem 6: cmbOperate.ListIndex = 0
    With lstFields
        lng值类型 = .ItemData(.ListIndex)

        If lng值类型 = 0 Then
            '本行不能取值
            .ListIndex = .ListIndex - 1
            Exit Sub

        ElseIf lng值类型 = 1 Or lng值类型 = 5 Then
            cmbOperate.Clear
            cmbOperate.AddItem "等于                             ="
            cmbOperate.AddItem "不等于                          <>"
            cmbOperate.ListIndex = 0

            '固定取值
            cmbList.Clear
            If lng值类型 = 1 Then
                cmbList.AddItem "是": cmbList.ItemData(cmbList.NewIndex) = 1
                cmbList.AddItem "否": cmbList.ItemData(cmbList.NewIndex) = 0
            Else
                Set值列表 .List(.ListIndex)
            End If
            cmbList.AddItem ""
            cmbList.ListIndex = 0

            cmbList.Visible = True
            cmbExample.Visible = False
        Else
            cmbOperate.Clear
            cmbOperate.AddItem "等于                             ="
            cmbOperate.AddItem "不等于                          <>"
            cmbOperate.AddItem "小于                             <"
            cmbOperate.AddItem "小于等于                        <="
            cmbOperate.AddItem "大于                             >"
            cmbOperate.AddItem "大于等于                        >="
            If lng值类型 = 3 Then cmbOperate.AddItem "包含                          LIKE"
            If lng值类型 = 3 Then cmbOperate.AddItem "开头                          LIKE"
            cmbOperate.ListIndex = 0

            cmbList.Visible = False
            cmbExample.Visible = True

            Dim rsExample As New ADODB.Recordset
            rsExample.CursorLocation = adUseClient
            str字段名 = .List(.ListIndex)
            Select Case str字段名
                Case "住院号", "姓名", "身份证号", "出生日期"
                    strExample = "select distinct " & str字段名 & " from 病人信息 where " & str字段名 & " is not null and  rownum<51"
                Case "姓名简码"
                    strExample = "select distinct zlSpellcode(姓名) from 病人信息 where 姓名 is not null and  rownum<51"
                Case "性别"
                    strExample = "select distinct " & str字段名 & " from 病人信息 "
                Case "住院次数"
                    strExample = "select distinct 主页ID from 病案主页 where rownum<51"
                Case "病案号"
                    strExample = "select distinct 病案号 from 住院病案记录 where rownum<51"
                Case "档案号"
                    strExample = "select distinct 档案号 from 住院病案记录 where rownum<51"
                Case "入院科室", "出院科室"
                    strExample = "select A.名称 from 部门表 A,部门性质说明 B " & _
                                " where A.ID=B.部门ID And B.工作性质='临床' and (B.服务对象=2 or B.服务对象=3) and " & Where撤档时间("A") & zl_获取站点限制(True, "a") & " order by A.名称"
                Case "手术已行手术", "手术切口", "手术愈合", "手术麻醉类型"
                    strExample = "select distinct " & Mid(str字段名, 3) & " from 病人手麻记录 where 记录来源=4 and rownum<51"
                Case "手术日期", "主刀医师", "麻醉医师"
                    strExample = "select distinct " & str字段名 & " from 病人手麻记录 where 记录来源=4 and rownum<51"
                Case "手术编码"
                    strExample = "select distinct B.编码 from 病人手麻记录 A,疾病编码目录 B where A.手术操作ID=B.ID and A.记录来源=4 and rownum<51"
                Case "手术简码"
                    If mint简码方式 = 0 Then
                        strExample = "select distinct B.简码 from 病人手麻记录 A,疾病编码目录 B where A.手术操作ID=B.ID and A.记录来源=4 and rownum<51"
                    Else
                        strExample = "select distinct B.五笔码 from 病人手麻记录 A,疾病编码目录 B where A.手术操作ID=B.ID and A.记录来源=4 and rownum<51"
                    End If

                Case "诊断描述信息", "诊断出院情况", "诊断编码序号"
                    If str字段名 = "诊断描述信息" Then
                        strExample = "select distinct 诊断描述 from 病人诊断记录 where rownum<51"
                    ElseIf str字段名 = "诊断编码序号" Then
                        strExample = "select distinct " & Mid(str字段名, 3) & " from 病人诊断记录 where 记录来源=4 and  rownum<51"
                    Else
                        strExample = "Select 名称 From 治疗结果"
                    End If
                Case "诊断次序"
                    strExample = "select distinct 诊断次序 from 病人诊断记录 where   记录来源=4 and  rownum<51"
                Case "诊断编码"
                    strExample = "select distinct B.编码 from 病人诊断记录 A,疾病编码目录 B where A.记录来源=4 and  A.疾病ID=B.ID and rownum<51"
                Case "诊断简码"
                    If mint简码方式 = 0 Then
                        strExample = "select distinct  B.简码 from 病人诊断记录 A,疾病编码目录 B where A.记录来源=4 and  A.疾病ID=B.ID and rownum<51"
                    Else
                        strExample = "select distinct  B.五笔码 from 病人诊断记录 A,疾病编码目录 B where A.记录来源=4 and  A.疾病ID=B.ID and rownum<51"
                    End If
'                Case "中医诊断编码"
'                    strExample = "select distinct B.编码 from 病人诊断记录 A,疾病编码目录 B where A.疾病ID=B.Id And A.诊断类型 In (11,12,13) and rownum<51"
'                Case "中医诊断描述"
'                    strExample = "select distinct 诊断描述 from 病人诊断记录 where 诊断类型 In (11,12,13) and rownum<51"
'                Case "中医出院情况"
'                    strExample = "Select 名称 From 治疗结果"
                Case "中医候诊"
                    strExample = "select distinct B.编码 from 病人诊断记录 A,疾病编码目录 B where  A.记录来源=4 and  A.证候ID=B.ID and rownum<51"
                Case "过敏药物"
                    strExample = "select distinct 过敏药物 from 病人过敏药物 where rownum<51"
                Case "病案质量", "科主任", "主任医师", "主治医师", "医保号", "输红细胞", "输血小板", "输血浆", "输全血"
                    strExample = "select distinct 信息值 from 病案主页从表 where 信息名='" & str字段名 & "' and rownum<51"
                Case "年龄"
                    strExample = "select distinct trunc((sysdate-出生日期)/365) from 病人信息 where rownum<51"
                Case "借阅人", "借出人", "批准人", "借阅日期", "借阅部门", "借阅用途"
                    If str字段名 = "借阅日期" Then
                        strExample = "select distinct 借阅时间 from 病案主页 A ,借阅记录 O where A.病人ID=O.病人ID AND A.主页ID=O.主页ID And A.编目日期 is not null And 归还时间 Is Null And rownum<51"
                    Else
                        strExample = "select distinct " & str字段名 & " from 病案主页 A ,借阅记录 O where A.病人ID=O.病人ID AND A.主页ID=O.主页ID And A.编目日期 is not null And 归还时间 Is Null And rownum<51"
                    End If
                Case Else
                    If mlngIndex < lstFields.ListIndex Then
                        strExample = "select distinct 信息值 from 病案主页从表 where 信息名='" & str字段名 & "' and rownum<51"
                    Else
                        strExample = "select distinct " & str字段名 & " from 病案主页 where rownum<51"
                    End If
            End Select

            On Error GoTo errHandle

            '得到一些例子
            Set rsExample = zlDatabase.OpenSQLRecord(strExample, Me.Caption)

            cmbExample.Text = ""
            cmbExample.Clear
            Do Until rsExample.EOF
                If Not IsNull(rsExample(0)) Then
                    If .ItemData(.ListIndex) = 4 Then
                        cmbExample.AddItem Format(rsExample(0), "yyyy-MM-dd")
                    Else
                        cmbExample.AddItem rsExample(0)
                    End If
                End If
                rsExample.MoveNext
            Loop
            If cmbExample.ListCount > 0 Then cmbExample.ListIndex = 0
        End If
    End With
    CreateSQL
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Set值列表(ByVal str类型 As String)
    With cmbList
        Select Case str类型
            Case "诊断类型"
                .AddItem "门诊诊断": .ItemData(.NewIndex) = 1
                .AddItem "入院诊断": .ItemData(.NewIndex) = 2
                'by gzh
                .AddItem "出院主要诊断": .ItemData(.NewIndex) = 3
                .AddItem "出院次要诊断": .ItemData(.NewIndex) = 3
                .AddItem "院内感染": .ItemData(.NewIndex) = 5
                .AddItem "病理诊断": .ItemData(.NewIndex) = 6
                .AddItem "损伤中毒码": .ItemData(.NewIndex) = 7
'                If gSystemPara.bln购买中医 = True Then
                    .AddItem "中医门诊诊断": .ItemData(.NewIndex) = 11
                    .AddItem "中医入院诊断": .ItemData(.NewIndex) = 12
                    .AddItem "中医出院诊断": .ItemData(.NewIndex) = 13
                    .AddItem "中医主证诊断": .ItemData(.NewIndex) = 14
'                End If

            Case "符合类型"
                .AddItem "门诊与出院": .ItemData(.NewIndex) = 1
                .AddItem "入院与出院": .ItemData(.NewIndex) = 2
                .AddItem "放射与病理": .ItemData(.NewIndex) = 3
                .AddItem "临床与病理": .ItemData(.NewIndex) = 4
                .AddItem "临床与尸检": .ItemData(.NewIndex) = 5
                .AddItem "术前与术后": .ItemData(.NewIndex) = 6
'                If gSystemPara.bln购买中医 = True Then
                    .AddItem "中医门诊与出院": .ItemData(.NewIndex) = 11
                    .AddItem "中医入院与出院": .ItemData(.NewIndex) = 12
'                End If
            Case "符合情况"
                .AddItem "符合": .ItemData(.NewIndex) = 1
                .AddItem "不符合": .ItemData(.NewIndex) = 2
                .AddItem "不肯定": .ItemData(.NewIndex) = 3
        End Select
    End With
End Sub


Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '--------------------------------------------------------------------------------------------------------------
    '功能：
    '参数：
    '返回：
    '--------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strSQL As String
    Dim intRow As Integer
    Dim strTmp As String
    Dim objItem As ListItem
    Dim strCustom As String
    Dim strWhere As String
    Dim i As Long
    On Error GoTo errHand

    Call SQLRecord(rsSQL)

    Select Case strCommand
    '--------------------------------------------------------------------------------------------------------------
    Case "初始控件"
        Set mclsPatient = New clsVsf
        With mclsPatient
            Call .Initialize(Me.Controls, vsfPatient, True, True, frmPubResource.GetImageList(16))
            Call .ClearColumn
            Call .AppendColumn("序号", 500, flexAlignCenterCenter, flexDTString, "", "", False)
            Call .AppendColumn("选择", 500, flexAlignCenterCenter, flexDTBoolean, "", , True)
            Call .AppendColumn("住院号", 1500, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("姓名", 900, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("性别", 500, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("年龄", 500, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("出生日期", 1100, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("住院次数", 900, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("家庭地址", 1350, flexAlignLeftCenter, flexDTString, "", , True)
            
            Call .AppendColumn("婚姻状况", 900, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("家庭电话", 1000, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("工作单位", 1200, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("单位电话", 1000, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("联系人姓名", 1000, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("联系人关系", 1200, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("联系人地址", 1200, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("联系人电话", 1000, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("病案号", 1000, flexAlignLeftCenter, flexDTString, "", , True)
            
            Call .AppendColumn("病案号集合", 1200, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("病案存储状态", 1440, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("病案存放位置", 1440, flexAlignLeftCenter, flexDTString, "", , True)
            
            Call .AppendColumn("病人id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("主页id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            
            Call .AppendColumn("入院时间", 1670, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm", , True)
            Call .AppendColumn("出院时间", 1670, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm", , True)
            Call .AppendColumn("出院科室", 1200, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("借出状态", 0, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("借出申请人", 0, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("", 0, flexAlignLeftCenter, flexDTString, "", , True)
            
            Call .InitializeEdit(True, False, False)
            Call .InitializeEditColumn(.ColIndex("选择"), True, vbVsfEditCheck)
            
            .AppendRows = True
        End With
        
        Call InitToolBox
        
        mint简码方式 = Val(zlDatabase.GetPara("简码方式"))
        
    '--------------------------------------------------------------------------------------------------------------
    Case "初始数据"

        '性别
        cbo(0).Clear
        cbo(0).AddItem ""
        Set rs = gclsPackage.GetBaseCode("性别")
        If rs.BOF = False Then Call AddComboData(cbo(0), rs, "名称", , "缺省标志", False)
    
        '婚姻状况
        cbo(1).Clear
        cbo(1).AddItem ""
        Set rs = gclsPackage.GetBaseCode("婚姻状况")
        If rs.BOF = False Then Call AddComboData(cbo(1), rs, "名称", , "缺省标志", False)
        
        '临床科室
        cbo(2).Clear
        cbo(2).AddItem ""
        Set rs = gclsPackage.GetDept("临床")
        If rs.BOF = False Then Call AddComboData(cbo(2), rs, "名称", "ID", , False)
        
        cmbLogical.Clear
        cmbLogical.AddItem "或者"
        cmbLogical.AddItem "并且"
        cmbLogical.ListIndex = 0
        
        Set mrsPatient = New ADODB.Recordset
        With mrsPatient
            .Fields.Append "ID", adVarChar, 30
            .Fields.Append "病人id", adVarChar, 25
            .Fields.Append "主页id", adVarChar, 10
            .Fields.Append "性别", adVarChar, 50
            .Fields.Append "年龄", adVarChar, 30
            .Fields.Append "姓名", adVarChar, 30
            .Fields.Append "婚姻状况", adVarChar, 50
            .Fields.Append "入院时间", adVarChar, 30
            .Fields.Append "出院时间", adVarChar, 30
            .Fields.Append "出院科室", adVarChar, 50
            
            .Fields.Append "住院号", adVarChar, 50
            .Fields.Append "病案号", adVarChar, 50
            .Fields.Append "住院次数", adVarChar, 50
            
            .Open
        End With
        
        strTmp = zlDatabase.GetPara("常用条件", glngSys, mlngMoudal, "", Array(cmd(4), cmd(5), cmd(6)), IsPrivs(mstrPrivs, "参数设置"))
        If strTmp <> "" Then
            For intRow = 0 To UBound(Split(strTmp, "|")) Step 2
                Set objItem = lvw.ListItems.Add(, , Split(strTmp, "|")(intRow), 1, 1)
                objItem.Tag = Split(strTmp, "|")(intRow + 1)
            Next
        End If
        
        '检查病案系统是否存在
        Set rs = gclsPackage.GetMedicalExits
        If Not rs.EOF Then
            mbln病案系统 = True
        Else
            mbln病案系统 = False
        End If
        
    '--------------------------------------------------------------------------------------------------------------
    Case "刷新数据"
        
        If CreateSQL = False Then
            Exit Function
        End If
        strWhere = ""
        mstrFilter = mstrReturn
        mstr显示 = mstr条件
        Call FillData(strWhere)
        
        mclsPatient.AppendRows = True
        DataChanged = False
    '--------------------------------------------------------------------------------------------------------------
    Case "搜索数据"
        strCustom = " And (1=1)"
        mclsPatient.ClearGrid
        mstr显示 = ""
        strWhere = GetCustomWhere(Val(cmd(0).Tag), cbo(0).Text, cbo(2).ItemData(cbo(2).ListIndex), Val(txt(6).Text), Val(txt(4).Text), cbo(1).Text, txt(1).Text, txt(2).Text)
       
        '基本信息搜索
        If mbln病案系统 Then
            '病案系统已经安装,检查病案是否编目
            mstrFilter = ",(select A.病人ID,A.主页ID from 病案主页 A  where A.编目日期 is not null  and  A.封存时间 is  null ) Y where "
        Else
            mstrFilter = ",(select A.病人ID,A.主页ID from 病案主页 A  where A.封存时间 is  null ) Y where "
        End If
        
        Call FillData(strWhere)
        mclsPatient.AppendRows = True
        DataChanged = False
    '--------------------------------------------------------------------------------------------------------------
    Case "清空数据"
        
        DataChanged = False
        
    '--------------------------------------------------------------------------------------------------------------
    Case "校验数据"
        
        With vsfPatient
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("选择")) = True Then
                    If .TextMatrix(i, .ColIndex("病案存储状态")) = "在院" Then
                        ExecuteCommand = True
                    Else
                        MsgBox "选择的病案:[" & .TextMatrix(i, .ColIndex("姓名")) & "]已经被[" & .TextMatrix(i, .ColIndex("借出申请人")) & "]申请借出,请重新选择!", vbInformation, gstrSysName
                        ExecuteCommand = False
                        Exit Function
                    End If

                End If
            Next
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case "保存数据"
        
        
        Call DeleteRecordData(mrsPatient)
        With vsfPatient
            For intRow = 1 To .Rows - 1
                If Trim(.TextMatrix(intRow, .ColIndex("ID"))) <> "" And Trim(.TextMatrix(intRow, .ColIndex("ID"))) <> "0" And Abs(Val(.TextMatrix(intRow, .ColIndex("选择")))) = 1 Then
                    mrsPatient.AddNew
                    mrsPatient("ID").Value = Trim(.TextMatrix(intRow, .ColIndex("ID")))
                    mrsPatient("病人id").Value = Val(.TextMatrix(intRow, .ColIndex("病人id")))
                    mrsPatient("主页id").Value = Val(.TextMatrix(intRow, .ColIndex("主页id")))
                    mrsPatient("姓名").Value = Trim(.TextMatrix(intRow, .ColIndex("姓名")))
                    mrsPatient("性别").Value = Trim(.TextMatrix(intRow, .ColIndex("性别")))
                    mrsPatient("年龄").Value = Trim(.TextMatrix(intRow, .ColIndex("年龄")))
                    mrsPatient("婚姻状况").Value = Trim(.TextMatrix(intRow, .ColIndex("婚姻状况")))
                    mrsPatient("入院时间").Value = Format(.TextMatrix(intRow, .ColIndex("入院时间")), "yyyy-MM-dd HH:mm:ss")
                    mrsPatient("出院时间").Value = Format(.TextMatrix(intRow, .ColIndex("出院时间")), "yyyy-MM-dd HH:mm:ss")
                    mrsPatient("出院科室").Value = Trim(.TextMatrix(intRow, .ColIndex("出院科室")))
                    
                    mrsPatient("住院号").Value = Trim(.TextMatrix(intRow, .ColIndex("住院号")))
                    mrsPatient("病案号").Value = Trim(.TextMatrix(intRow, .ColIndex("病案号")))
                    mrsPatient("住院次数").Value = Trim(.TextMatrix(intRow, .ColIndex("住院次数")))
                End If
            Next
        End With
    
    '------------------------------------------------------------------------------------------------------------------
    Case "读注册表"
        
        If Val(GetPara("使用个性化风格")) = 1 Then
            '使用个性化设置
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "写注册表"
        If Val(GetPara("使用个性化风格")) = 1 Then
            '使用个性化设置

        End If
        
        strTmp = ""
        
        For intRow = 1 To lvw.ListItems.count
            strTmp = strTmp & "|" & lvw.ListItems(intRow).Text & "|" & lvw.ListItems(intRow).Tag
        Next
        If strTmp <> "" Then strTmp = Mid(strTmp, 2)
        
        strTmp = Replace(strTmp, "'", "123456789`1234567890")
        strTmp = Replace(strTmp, "123456789`1234567890", "''")
                        
        Call zlDatabase.SetPara("常用条件", strTmp, glngSys, mlngMoudal)
        
    End Select
    
    
    ExecuteCommand = True

    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

Private Property Let DataChanged(ByVal blnData As Boolean)
    mblnDataChanged = blnData
End Property

Private Property Get DataChanged() As Boolean
    DataChanged = mblnDataChanged
End Property

Private Property Let ConditionChanged(ByVal blnData As Boolean)
    mblnConditionChanged = blnData
End Property

Private Property Get ConditionChanged() As Boolean
    DataChanged = mblnConditionChanged
End Property

Private Sub cbo_Change(Index As Integer)
    Select Case Index
    Case 3, 4, 5
    Case Else
        ConditionChanged = True
    End Select
End Sub

Private Sub cbo_Click(Index As Integer)
    Dim blnSave As Boolean
    
    Select Case Index
    Case 3
        blnSave = ConditionChanged
        Select Case cbo(3).ItemData(cbo(3).ListIndex)
        Case 1
            With cbo(4)
                .Clear
                .AddItem "="
                .AddItem ">"
                .AddItem ">="
                .AddItem "<"
                .AddItem "<="
                .AddItem "<>"
                .AddItem "Like"
                .ListIndex = 0
            End With
        Case 2
            With cbo(4)
                .Clear
                .AddItem "="
                .AddItem ">"
                .AddItem ">="
                .AddItem "<"
                .AddItem "<="
                .AddItem "<>"
                .ListIndex = 0
            End With
        End Select
        ConditionChanged = blnSave
    Case 4
    
    Case 5
    
    Case Else
        ConditionChanged = True
    End Select
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cbo_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(cbo(Index).Text, 0)
    
    Select Case Index
    Case 5
        If cbo(3).ItemData(cbo(3).ListIndex) = 2 Then
            Cancel = Not IsDate(cbo(5).Text)
        End If
    End Select
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim lngLoop As Long
    
    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_SelAll
        
        With vsfPatient
            .Cell(flexcpText, 1, .ColIndex("选择"), .Rows - 1, .ColIndex("选择")) = 1
            DataChanged = True
        End With
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_ClsAll
        
        With vsfPatient
            .Cell(flexcpText, 1, .ColIndex("选择"), .Rows - 1, .ColIndex("选择")) = 0
            DataChanged = True
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Refresh
        
        Call ExecuteCommand("搜索数据")
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_SaveExit                  '保存数据
    
        If ExecuteCommand("校验数据") And DataChanged Then
            If ExecuteCommand("保存数据") Then
                
                DataChanged = False
                mblnOK = True
                
                Unload Me
                
            End If
        End If
                
    '------------------------------------------------------------------------------------------------------------------
    Case Else
    
        If Control.ID > 400 And Control.ID < 500 Then
            
        Else
             '与业务无关的功能，公共的功能
            Call CommandBarExecutePublic(Control, Me, vsfPatient, "查找病人结果清单")
        End If
        
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    On Error GoTo errHand
    
    Select Case Control.ID
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
            With vsfPatient
                Control.Enabled = (.TextMatrix(.Row, .ColIndex("ID")) <> "" And .TextMatrix(.Row, .ColIndex("ID")) <> "0") And .Rows > 1
            End With
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_SelAll, conMenu_Edit_ClsAll
            
            With vsfPatient
                Control.Enabled = (.TextMatrix(.Row, .ColIndex("ID")) <> "" And .TextMatrix(.Row, .ColIndex("ID")) <> "0") And .Rows > 1
            End With
            
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_Refresh
            
            Control.Enabled = Not DataChanged
            cmdRefresh.Enabled = Control.Enabled
            
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_SaveExit
        
            Control.Enabled = DataChanged
            
        '--------------------------------------------------------------------------------------------------------------
        Case Else
            Call CommandBarUpdatePublic(Control, Me)
    End Select
    '------------------------------------------------------------------------------------------------------------------
errHand:
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim rs              As New ADODB.Recordset
    Dim rsData          As New ADODB.Recordset
    Dim strTmp          As String
    Dim intRow          As Integer
    Dim objItem         As ListItem
    
    Select Case Index
        '------------------------------------------------------------------------------------------------------------------
        Case 0  '疾病编码选择
            Set rsData = gclsPackage.GetDisease
            If ShowPubSelect(Me, txt(0), 3, "编码,1200,0,;名称,2700,0,;简码,900,0,;附码,900,0,", Me.Name & "\疾病编码选择", "请从下表中选择一个疾病编码项目", rsData, rs, 8790, 4500, , Val(cmd(Index).Tag)) = 1 Then
                If Val(cmd(Index).Tag) <> zlCommFun.NVL(rs("ID").Value, 0) Then
                    txt(0).Text = zlCommFun.NVL(rs("名称").Value)
                    cmd(Index).Tag = zlCommFun.NVL(rs("ID").Value, 0)
                    usrSaveItem.疾病名称 = txt(0).Text
                    txt(0).Tag = ""
                End If
                DataChanged = True
            End If
        '------------------------------------------------------------------------------------------------------------------
        Case 1  '新增
            
            
            
            Call cmdAdd_Click
            
            
        '------------------------------------------------------------------------------------------------------------------
        Case 2  '删除
            
            lvwCombine.ListItems.Remove lvwCombine.SelectedItem.Index
        '    Call CreateSQL
            If lvwCombine.ListItems.count = 0 Then
                cmd(2).Enabled = False
            Else
                lvwCombine.SelectedItem.Selected = True
            End If
        '------------------------------------------------------------------------------------------------------------------
        Case 4          '新增条件
            
            strTmp = GetConditionString
            If strTmp <> "" Then
                
                Set objItem = lvw.ListItems.Add(, , "新条件", 1, 1)
                objItem.Tag = strTmp
                objItem.Selected = True
                
                lvw.SetFocus
                lvw.StartLabelEdit
            End If
        '------------------------------------------------------------------------------------------------------------------
        Case 5          '更新条件
            If Not (lvw.SelectedItem Is Nothing) Then
                lvw.SelectedItem.Tag = GetConditionString
            End If
        '------------------------------------------------------------------------------------------------------------------
        Case 6          '删除条件
            
            If Not (lvw.SelectedItem Is Nothing) Then
                lvw.ListItems.Remove lvw.SelectedItem.Index
            End If
            lvw.SetFocus
        
    End Select
End Sub

Private Function GetConditionString() As String
    Dim intRow As Integer
    Dim strTmp As String
    
    strTmp = txt(0).Text & "'" & Val(cmd(0).Tag) & "'" & cbo(0).Text & "'" & cbo(2).Text & "'" & txt(6).Text & "'" & txt(4).Text & "'" & cbo(1).Text & "'" & IIf(opt(0).Value, 0, 1)
    
    With lvwCombine
        For intRow = 1 To .ListItems.count
            If .ListItems(intRow).Text <> "" Then
                strTmp = strTmp & "'" & .ListItems(intRow).Text & "'" & .ListItems(intRow).SubItems(1) & "'" & .ListItems(intRow).Tag
            End If
        Next
    End With
    
    strTmp = strTmp & "`" & txt(1).Text & "`" & txt(2).Text
    
    GetConditionString = strTmp
    
End Function

Private Sub cmdRefresh_Click()
    Call ExecuteCommand("刷新数据")
End Sub

Private Sub cmdSearch_Click()
    Call ExecuteCommand("搜索数据")
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picPane(0).hWnd
    Case 2
        Item.Handle = picPane(1).hWnd
    Case 3
        Item.Handle = picPane(5).hWnd
    End Select
End Sub

Private Sub Form_Load()
    Call InitCommandBar
    Call InitDockPannel
    Call InitListPatient
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Call SetPaneRange(dkpMain, 1, 100, 100, 300, Me.ScaleHeight)
    dkpMain.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Call ExecuteCommand("写注册表")
    
    Call SaveWinState(Me, App.ProductName)
    
    Set mclsPatient = Nothing
    
End Sub

Private Sub lvw_DblClick()
    Dim intRow As Integer
    Dim varAry As Variant
    Dim strTmp As String
    Dim strTmp2 As String
    Dim sItem As MSComctlLib.ListItem
    
    If lvw.SelectedItem Is Nothing Then Exit Sub
    If lvw.SelectedItem.Tag = "" Then Exit Sub
    
    strTmp = lvw.SelectedItem.Tag
    If InStr(strTmp, "`") > 0 Then
        strTmp2 = Mid(strTmp, InStr(strTmp, "`") + 1)
        strTmp = Mid(strTmp, 1, InStr(strTmp, "`") - 1)
        
        varAry = Split(strTmp2, "`")
        txt(1).Text = varAry(0)
        txt(2).Text = varAry(1)
        
    End If
    
    varAry = Split(strTmp, "'")
    
    On Error Resume Next
    
    txt(0).Text = varAry(0)
    cmd(0).Tag = Val(varAry(1))
    
    Call zlControl.CboLocate(cbo(0), varAry(2))
    Call zlControl.CboLocate(cbo(2), varAry(3))
    
    txt(6).Text = varAry(4)
    txt(4).Text = varAry(5)
    
    Call zlControl.CboLocate(cbo(1), varAry(6))
    
    opt(Val(varAry(7))).Value = True
    
    With lvwCombine
        .ListItems.Clear
        For intRow = 8 To UBound(varAry) Step 4
            Set sItem = .ListItems.Add(, , varAry(intRow))
                sItem.SubItems(1) = varAry(intRow + 1)
                sItem.Tag = varAry(intRow + 2)
        Next
    End With
    
End Sub

Private Sub opt_Click(Index As Integer)
    ConditionChanged = True
End Sub

Private Sub opt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next

    Select Case Index
        Case 0
            tpl.Move 0, 0, picPane(Index).Width, picPane(Index).Height
        Case 1
            vsfPatient.Move 0, 0, picPane(Index).Width, picPane(Index).Height - 45
            mclsPatient.AppendRows = True
        Case 2
            txt(0).Move txt(0).Left, txt(0).Top, picPane(Index).Width - cmd(0).Width - 45 - txt(0).Left
            cmd(0).Move txt(0).Left + txt(0).Width + 30, cmd(0).Top
            cbo(0).Move cbo(0).Left, cbo(0).Top, txt(0).Width
            cbo(1).Move cbo(1).Left, cbo(1).Top, txt(0).Width
            cbo(2).Move cbo(0).Left, cbo(2).Top, txt(0).Width
            txt(4).Move txt(4).Left, txt(4).Top, picPane(Index).Width - cmd(0).Width - 45 - txt(4).Left
            txt(1).Move txt(0).Left, txt(1).Top, txt(0).Width, txt(0).Height
            txt(2).Move txt(0).Left, txt(2).Top, txt(0).Width, txt(0).Height
            cmdSearch.Move txt(2).Left + txt(2).Width - cmdSearch.Width, cmdSearch.Top, cmdSearch.Width, cmdSearch.Height
            
        Case 3
            'vsf(0).Move 0, vsf(0).Top, picPane(Index).Width - cmd(1).Width - 30, picPane(Index).Height - vsf(0).Top - cmdRefresh.Height - 90
            lvwCombine.Move 0, lvwCombine.Top, picPane(Index).Width - cmd(1).Width - 30, picPane(Index).Height - lvwCombine.Top - cmdRefresh.Height - 90
'            mclsVsf(0).AppendRows = True
            
            lstFields.Move lstFields.Left, lstFields.Top, picPane(Index).Width - cmbLogical.Left
            
            cmbLogical.Move cmbLogical.Left, cmbLogical.Top, cmbLogical.Width
            cmbList.Move cmbList.Left, cmbList.Top, picPane(Index).Width - cmbList.Left
            cmbExample.Move cmbList.Left, cmbList.Top, cmbList.Width
            cmd(1).Move lvwCombine.Left + lvwCombine.Width + 30, cmd(1).Top
            cmd(2).Move cmd(1).Left, cmd(2).Top
            cmdRefresh.Move lvwCombine.Width - cmdRefresh.Width, lvwCombine.Top + lvwCombine.Height + 30
            
            opt(0).Move 0, lvwCombine.Top + lvwCombine.Height + 30
            opt(1).Move 0, opt(0).Top + opt(0).Height + 30
        Case 4
            lvw.Move 0, 0, picPane(Index).Width - cmd(4).Width - 30, picPane(Index).Height
            cmd(4).Move lvw.Left + lvw.Width + 30, cmd(4).Top
            cmd(5).Move lvw.Left + lvw.Width + 30, cmd(5).Top
            cmd(6).Move lvw.Left + lvw.Width + 30, cmd(6).Top
        Case 5
            tabMain.Move 0, 0, picPane(5).Width, picPane(5).Height
            lvw主页.Move tabMain.ClientLeft, tabMain.ClientTop, tabMain.ClientWidth, tabMain.ClientHeight
    End Select
End Sub


Private Sub txt_Change(Index As Integer)

    ConditionChanged = True

    Select Case Index
        Case 0
            txt(Index).Tag = "Changed"
    End Select

End Sub

Private Sub txt_GotFocus(Index As Integer)

    zlControl.TxtSelAll txt(Index)

    Select Case Index
        Case 0
            zlCommFun.OpenIme True
    End Select

End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
        Case 0
            If KeyCode = vbKeyDelete Then
                KeyCode = 0
                txt(Index).Text = ""
                cmd(0).Tag = ""
                txt(Index).Tag = ""
                usrSaveItem.疾病名称 = ""
            End If
    End Select
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim StrText             As String
    Dim strTmp              As String
    Dim rs                  As New ADODB.Recordset
    Dim rsData              As New ADODB.Recordset
    Dim bytMode             As Byte

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Select Case Index
            Case 0
                If txt(Index).Tag <> "" Then
                    txt(Index).Tag = ""
    
                    Set rsData = gclsPackage.GetDisease(UCase(txt(Index).Text))
                    If ShowPubSelect(Me, txt(Index), 2, "编码,1200,0,;名称,2700,0,;简码,900,0,;附码,900,0,", Me.Name & "\疾病编码过滤", "请从下面选择一个疾病编目项目", rsData, rs) = 1 Then
                        If cmd(0).Tag <> zlCommFun.NVL(rs("ID").Value) Then
                            cmd(0).Tag = zlCommFun.NVL(rs("ID").Value)
                            txt(Index).Text = zlCommFun.NVL(rs("名称").Value)
                            txt(Index).Tag = ""
                            ConditionChanged = True
        
                            usrSaveItem.疾病名称 = txt(Index).Text
                        Else
                            txt(Index).Text = usrSaveItem.疾病名称
                            txt(Index).Tag = ""
                        End If
                    Else
                        txt(Index).Text = usrSaveItem.疾病名称
                        txt(Index).Tag = ""
                        Exit Sub
                    End If
    
                End If
            Case Else
                zlCommFun.PressKey vbKeyTab
        End Select
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)

    Select Case Index
    Case 0
        zlCommFun.OpenIme False
    End Select

End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        glngTXTProc = GetWindowLong(txt(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)

    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
    If Cancel Then Exit Sub

    Select Case Index
        Case 0
            If (txt(Index).Tag = "Changed") Then
                txt(Index).Text = usrSaveItem.疾病名称
                txt(Index).Tag = ""
            End If
    End Select

End Sub

Private Sub vsfPatient_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '编辑处理
    Call mclsPatient.AfterEdit(Row, Col)
    DataChanged = True
End Sub

Private Sub vsfPatient_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsPatient.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsfPatient_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsPatient.AppendRows = True
End Sub

Private Sub vsfPatient_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    mclsPatient.AppendRows = True
End Sub

Private Sub vsfPatient_DblClick()
    '编辑处理
    Call mclsPatient.DbClick
End Sub

Private Sub vsfPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    '编辑处理
    Call mclsPatient.KeyDown(KeyCode, Shift)
End Sub

Private Sub vsfPatient_KeyPress(KeyAscii As Integer)
    'ToDo...
    If KeyAscii = vbKeyReturn Then Call vsfPatient_DblClick
    
    '编辑处理,最后调用
    Call mclsPatient.KeyPress(KeyAscii)
End Sub

Private Sub vsfPatient_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    '编辑处理
    Call mclsPatient.KeyPressEdit(KeyAscii)
End Sub

Private Sub vsfPatient_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    '编辑处理
    Call mclsPatient.EditSelAll
End Sub

Private Sub vsfPatient_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '编辑处理
    Call mclsPatient.BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsfPatient_AfterMoveColumn(ByVal Col As Long, Position As Long)
    Call SaveHead(vsfPatient, 1)
End Sub

Private Sub vsfPatient_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vsfPatient.Editable = flexEDNone Then
        Cancel = True
        Exit Sub
    End If
    
    Select Case Col
        Case vsfPatient.ColIndex("选择")
            Cancel = False
            Exit Sub
        Case Else
            Cancel = True
            Exit Sub
    End Select
End Sub

Private Sub vsfPatient_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    Select Case Col
        Case vsfPatient.ColIndex("序号"), vsfPatient.ColIndex("选择")
            Position = -1
            Exit Sub
    End Select
    If Position = 0 Or Position = 1 Then
        Position = Col
    End If
End Sub

Private Sub vsfPatient_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call SaveHead(vsfPatient, 1)
End Sub

Private Sub vsfPatient_Click()
    With vsfPatient
        If .Row > 0 Then
'            Call FillTabData
            Call SetMenu
            Exit Sub
        End If
    End With
End Sub

Private Sub vsfPatient_EnterCell()
    Dim lngID As Long
     With vsfPatient
        If .Row > 0 Then mlngNewRow = .Row
        If mlngOldRow <> mlngNewRow And .Rows > 2 Then
            lngID = Val(vsfPatient.TextMatrix(vsfPatient.Row, vsfPatient.ColIndex("病人ID")))
            If lngID > 0 Then Call FillTabData
        End If
    End With
End Sub

Private Sub vsfPatient_LeaveCell()
    With vsfPatient
        If .Row > 0 Then mlngOldRow = .Row
'        MsgBox "LeaveCell" & mlngOldRow
    End With
End Sub

'Private Sub vsfPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    '    RightHead
'    Dim intGetHeight As Integer
'    Dim intGetWidth As Integer
'
'    intGetWidth = vsfPatient.ColWidth(0)
'    intGetHeight = vsfPatient.RowHeight(0)
'    If (Button = 2) Then
'        If X < intGetWidth And Y < intGetHeight Then
'            Call RightHead(vsfPatient, 1)
'        Else
''            PopupMenu mnuShortB, vbPopupMenuRightButton
'        End If
'    End If
'End Sub
'
'Private Sub RightHead(ByVal vsGrid As VSFlexGrid, intListOrDetail As Integer)
'    Dim lngLeft As Long
'    Dim lngTop As Long
'    Dim strHeadInfo As String
'    Dim vRect  As RECT
'    vRect = GetControlRect(vsGrid.hWnd)
'    lngLeft = vRect.Left + vsGrid.Left
'    lngTop = vRect.Top + vsGrid.RowHeight(0)
'    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsGrid, lngLeft, lngTop, vsGrid.RowHeight(0))
'    Call SaveHead(vsGrid, intListOrDetail)
'End Sub


Private Sub initVfgPatient()
    Dim strHead As String
    
    strHead = "序号,500,4,1;选择,500,4,1;住院号,1500,1,1;姓名,900,1,0;性别,500,4,0;年龄,500,7,0;出生日期,1100,1,0;住院次数,900,7,1;家庭地址,1350,1,0;婚姻状况,900,1,0;家庭电话,1000,1,0;" & _
              "工作单位,1200,1,0;单位电话,1000,1,0;联系人姓名,1000,1,0;联系人关系,1200,1,0;联系人地址,1200,1,0;联系人电话,1000,1,0;病案号,1000,1,1;病案号集合,1200,1,1;病案存储状态,1440,1,0;" & _
              "病案存放位置,1440,1,0;入院时间,1670,1,0;出院时间,1670,1,0;病人ID,0,7,-1;主页ID,0,7,-1;ID,0,7,-1;出院科室,1100,7,-1;借出状态,0,7,-1;借出申请人,0,7,-1"
'              mclsPatient.LoadStateFromString strHead
    Call SetVsFlexGridChangeHead(strHead, vsfPatient, 1)
End Sub

'问题24813
Private Sub SetInitVfgDataFormat(ByVal vsGrid As VSFlexGrid)
    Dim i As Long
    With vsGrid
        .ColDataType(.ColIndex("选择")) = flexDTBoolean
        .ForeColorSel = .CellForeColor
        .ExplorerBar = flexExSortShowAndMove
        .SelectionMode = flexSelectionByRow
        .Editable = flexEDKbdMouse
    End With
End Sub

'问题24813
Private Function FillData(ByVal strWhere As String) As Boolean
'装入符合条件的病案到VSFlexGrid中
    Dim rs病案 As New ADODB.Recordset
    Dim lngCol As Long, varValue As Variant
    Dim lst As ListItem, bln提醒 As Boolean
    Dim strKey As String, strTemp As String
    Dim strRecord As String
    Dim i As Long
     
    zlCommFun.ShowFlash "正在装入病案记录，请稍等 ……"

    gstrSQL = "" & _
         "   Select  Trim(To_Char(max(Y1.病人id)))||'-'||Trim(To_Char(max(Y1.主页id))) As ID,X.病人ID,max(Y1.主页id) as 主页ID,max(Y1.住院号) as 住院号,X.姓名,X.性别,X.年龄,to_char(X.出生日期,'YYYY-MM-DD HH24:MI') as 出生日期,Zl_获取住院次数或主页id(X.病人id,max(y.主页id),0) as 住院次数," & _
         "           X.出生地点,X.身份证号,X.职业,X.婚姻状况,X.家庭地址,max(Z.病案号) as  病案号集合," & _
         "           X.家庭电话,X.联系人姓名,X.联系人关系,X.联系人地址,X.联系人电话,X.工作单位,X.单位电话,max(Z.病案号) as 病案号 ,Decode(max(D.病人ID),'','在院','借出') as 病案存储状态,max(Z.存放位置) as 病案存放位置," & _
         "           max(to_char(y1.入院日期,'YYYY-MM-DD HH24:MI')) As 入院时间 ,to_char(y1.出院日期,'YYYY-MM-DD HH24:MI') as 出院时间,max(C.名称) As 出院科室,max(D.病人ID) as 借出状态,max(D.申请人) As 借出申请人 " & _
         "   From 病人信息 X ,部门表 C,(Select Max(A1.申请人) as 申请人,Max(B1.病人ID) as 病人ID,Max(B1.主页ID) as 主页ID From 病案借阅记录 A1 ,病案借阅内容 B1,病案借阅人员 C1 Where A1.ID = B1.借阅ID And A1.ID = C1.借阅ID And A1.记录状态=2 Group By B1.病人ID,B1.主页ID,A1.申请人) D,住院病案记录 Z,病案主页 Y1  " & _
                  mstrFilter & " Y.病人ID=X.病人ID and X.病人ID=Y1.病人ID  And y1.主页ID=y.主页id  And C.ID = Y1.出院科室ID  And D.病人ID(+) = X.病人ID And  Y1.出院日期 is Not null  and X.病人ID=Z.病人ID(+) and Y1.病案状态=5 " & strWhere & _
         "Group By  X.病人id,Y1.主页ID,X.住院号, X.姓名, X.性别, X.年龄,x.出生日期,X.出生地点, X.身份证号, X.职业, X.婚姻状况, X.家庭地址, X.家庭电话, X.联系人姓名, X.联系人关系, X.联系人地址, X.联系人电话, X.工作单位," & _
         "    X.单位电话 ,X.入院时间, y1.出院日期 " & _
         "   Order by  to_char(y1.出院日期,'YYYY-MM-DD HH24:MI') Desc"
      
    On Error GoTo errHandle
    Call zlDatabase.OpenRecordset(rs病案, gstrSQL, Me.Caption)
    
    With vsfPatient
        Call initVfgPatient
        .Rows = IIf(rs病案.EOF, 0, rs病案.RecordCount) + 1
        
        If Not rs病案.EOF Then
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("序号")) = i
                .TextMatrix(i, .ColIndex("选择")) = 0
                .TextMatrix(i, .ColIndex("住院号")) = IIf(IsNull(rs病案!住院号), "", rs病案!住院号)
                .TextMatrix(i, .ColIndex("姓名")) = IIf(IsNull(rs病案!姓名), "", rs病案!姓名)
                .TextMatrix(i, .ColIndex("性别")) = IIf(IsNull(rs病案!性别), "", rs病案!性别)
                .TextMatrix(i, .ColIndex("年龄")) = IIf(IsNull(rs病案!年龄), 0, rs病案!年龄)
                .TextMatrix(i, .ColIndex("出生日期")) = IIf(IsNull(rs病案!出生日期), "", rs病案!出生日期)
                .TextMatrix(i, .ColIndex("住院次数")) = IIf(IsNull(rs病案!住院次数), "", rs病案!住院次数)
                .TextMatrix(i, .ColIndex("家庭地址")) = IIf(IsNull(rs病案!家庭地址), "", rs病案!家庭地址)
                .TextMatrix(i, .ColIndex("婚姻状况")) = IIf(IsNull(rs病案!婚姻状况), "", rs病案!婚姻状况)
                .TextMatrix(i, .ColIndex("家庭电话")) = IIf(IsNull(rs病案!家庭电话), "", rs病案!家庭电话)
                .TextMatrix(i, .ColIndex("工作单位")) = IIf(IsNull(rs病案!工作单位), "", rs病案!工作单位)
                .TextMatrix(i, .ColIndex("单位电话")) = IIf(IsNull(rs病案!单位电话), "", rs病案!单位电话)
                .TextMatrix(i, .ColIndex("联系人姓名")) = IIf(IsNull(rs病案!联系人姓名), "", rs病案!联系人姓名)
                .TextMatrix(i, .ColIndex("联系人关系")) = IIf(IsNull(rs病案!联系人关系), "", rs病案!联系人关系)
                .TextMatrix(i, .ColIndex("联系人地址")) = IIf(IsNull(rs病案!联系人地址), "", rs病案!联系人地址)
                .TextMatrix(i, .ColIndex("联系人电话")) = IIf(IsNull(rs病案!联系人电话), "", rs病案!联系人电话)
                .TextMatrix(i, .ColIndex("病案号")) = IIf(IsNull(rs病案!病案号), "", rs病案!病案号)
                .TextMatrix(i, .ColIndex("病案号集合")) = IIf(IsNull(rs病案!病案号集合), "", rs病案!病案号集合)
'                    If gSystemPara.bln单独病案号 = True And mnuViewRecord.Checked = False Then
'                        '如果允许单独编号并且显示病人时获得病案号集合
'                        varValue = Get病案号集合(CLng(rs病案("病人id").Value))
'                        .TextMatrix(i, .ColIndex("病案号集合")) = IIf(IsNull(varValue), "", varValue)
'                    ElseIf gSystemPara.bln单独病案号 = False And mnuViewRecord.Checked = False Then
'                        '不允许病案单独编号并且显示病人列表时
'                        varValue = rs病案("病案号").Value
'                        .TextMatrix(i, .ColIndex("病案号集合")) = IIf(IsNull(varValue), "", varValue)
'                    ElseIf gSystemPara.bln单独病案号 = True And mnuViewRecord.Checked = True Then
'                        '允许病案号单独编号并且显示病案列表时
'                        varValue = rs病案("病案号")
'                        .TextMatrix(i, .ColIndex("病案号集合")) = IIf(IsNull(varValue), "", varValue)
'                    ElseIf gSystemPara.bln单独病案号 = False And mnuViewRecord.Checked = True Then
'                        '不允许病案号单独编号并且显示病案列表时
'                        varValue = rs病案("病案号").Value
'                        .TextMatrix(i, .ColIndex("病案号集合")) = IIf(IsNull(varValue), "", varValue)
'                    End If
                .TextMatrix(i, .ColIndex("病案存储状态")) = IIf(IsNull(rs病案!病案存储状态), "", rs病案!病案存储状态)
                .TextMatrix(i, .ColIndex("病案存放位置")) = IIf(IsNull(rs病案!病案存放位置), "", rs病案!病案存放位置)
                .TextMatrix(i, .ColIndex("入院时间")) = IIf(IsNull(rs病案!入院时间), "", rs病案!入院时间)
                .TextMatrix(i, .ColIndex("出院时间")) = IIf(IsNull(rs病案!出院时间), "", rs病案!出院时间)
                .TextMatrix(i, .ColIndex("出院科室")) = IIf(IsNull(rs病案!出院科室), "", rs病案!出院科室)
                .TextMatrix(i, .ColIndex("病人ID")) = IIf(IsNull(rs病案!病人ID), 0, rs病案!病人ID)
                .TextMatrix(i, .ColIndex("主页ID")) = IIf(IsNull(rs病案!主页ID), 0, rs病案!主页ID)
                
                .TextMatrix(i, .ColIndex("借出状态")) = IIf(IsNull(rs病案!借出状态), 0, rs病案!借出状态)
                .TextMatrix(i, .ColIndex("借出申请人")) = IIf(IsNull(rs病案!借出申请人), 0, rs病案!借出申请人)
                
'                strRecord = IIf(IsNull(rs病案!病案存储状态), "", rs病案!病案存储状态)
'                Select Case strRecord
'                Case "在院"
'                    .Cell(flexcpPicture, i, .ColIndex("住院号")) = imgCelNo(0)
'                Case "部分借出"
'                    .Cell(flexcpPicture, i, .ColIndex("住院号")) = imgCelNo(1)
'                Case "借出"
'                    .Cell(flexcpPicture, i, .ColIndex("住院号")) = imgCelNo(2)
'                End Select

                strRecord = IIf(IsNull(rs病案!借出状态), "", rs病案!借出状态)
                If strRecord = "" Then
                    .Cell(flexcpPicture, i, .ColIndex("住院号")) = imgCelNo(0)
                Else
                    .Cell(flexcpPicture, i, .ColIndex("住院号")) = imgCelNo(2)
                End If
                .TextMatrix(i, .ColIndex("ID")) = rs病案!ID
                
                
                rs病案.MoveNext
            Next
        End If
    End With
    Call SetInitVfgDataFormat(vsfPatient)
    Call RestoreHead(vsfPatient, 1)

    If mlngOldRow > 0 And (vsfPatient.Rows - 1) >= mlngOldRow Then vsfPatient.Select mlngOldRow, 1
    '显示病人信息
    Call FillTabData
  
    Call zlCommFun.StopFlash
    rs病案.Close
    FillData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call zlCommFun.StopFlash
    rs病案.Close
    FillData = False
    
End Function

Public Function FillTabData() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim i As Long
    Dim lngID As Long, lngPageID As Long
    
    On Error GoTo errHandle
    
    tabMain.Tabs.Clear
    
    With tabMain.Tabs
        If vsfPatient.Rows = 1 Then
            tabMain.Tag = "0"
            .Add , "T0", "无主页数据"
            .Item("T0").Tag = 0
        Else
            lngID = Val(vsfPatient.TextMatrix(vsfPatient.Row, vsfPatient.ColIndex("病人ID")))
            lngPageID = Val(vsfPatient.TextMatrix(vsfPatient.Row, vsfPatient.ColIndex("主页ID")))
            tabMain.Tag = lngID
            gstrSQL = "" & _
                "   select 主页ID as 住院次数,Zl_获取住院次数或主页id(病人id,主页id,0) as 实际住院次数 from 病案主页 where 编目日期 is not null and 病人ID=[1] Order by 主页ID Desc"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngID)
            Do Until rsTemp.EOF
                i = rsTemp("住院次数")
                .Add , "T" & i, "第 " & NVL(rsTemp!实际住院次数) & " 次住院" '由于某次为留观病人，中间可能出现中断
                .Item("T" & i).Tag = i
                If i = lngPageID Then
                    .Item("T" & i).Selected = True
                End If
                rsTemp.MoveNext
            Loop
              
            If .count = 0 Then
                .Add , "T0", "无主页数据"
                .Item("T0").Tag = 0
            End If
        End If
    End With
    
    Call FillDetail
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function FillDetail() As Boolean
'装入指定病案的所有主页到ListView中
    Dim rs主页 As New ADODB.Recordset
    Dim lngCount As Long
    Dim fld As Field
    Dim lst As ListItem
    Dim str医保号 As String
    On Error GoTo errHandle
    
    lvw主页.ListItems.Clear
    rs主页.CursorLocation = adUseClient
'    gstrSQL = "" & _
'        " Select A.病案号 as 病案号1,F.病案号 as 病案号,A.住院号,C.名称 as 入院科室,to_char(A.入院日期,'YYYY-MM-DD HH24:MI') AS 入院日期,A.入院病况 as 入院病情" & _
'        "           ,D.名称 as 出院科室,to_char(A.出院日期,'YYYY-MM-DD HH24:MI') AS 出院日期" & _
'        "           ,A.住院天数,A.费用和,B.诊断描述 as 主要诊断名,E.编码 as 主要诊断编码,B.出院情况 as 主要诊断出院情况 " & _
'        "           ,decode(A.随诊标志,1,'是',2,'是',3,'是','') as 是否随诊,A.编目员姓名 as 编目员,to_char(A.编目日期,'YYYY-MM-DD') as 编目日期" & _
'        " From 病案主页 A,(Select * From 病人诊断记录 where 记录来源=4 and 诊断类型 in (3,13) and 病人ID=[1] and 主页ID=[2]) B,部门表 C,部门表 D,疾病编码目录 E,住院病案记录 F" & _
'        " Where A.编目日期 is not null and A.病人ID=B.病人ID(+)  and A.主页ID=B.主页ID(+) and " & _
'        "       A.入院科室ID=C.ID(+) and A.出院科室ID=D.ID and B.疾病ID=E.ID(+) and A.病人ID=[1] and A.主页ID=[2] and A.病人id=F.病人ID(+) "
    
    gstrSQL = "" & _
    " Select     F.病案号 as 病案号,A.住院号,C.名称 as 入院科室,to_char(A.入院日期,'YYYY-MM-DD HH24:MI') AS 入院日期,A.入院病况 as 入院病情" & _
    "           ,D.名称 as 出院科室,to_char(A.出院日期,'YYYY-MM-DD HH24:MI') AS 出院日期" & _
    "           ,A.住院天数,A.费用和,B.诊断描述 as 主要诊断名,E.编码 as 主要诊断编码,B.出院情况 as 主要诊断出院情况 " & _
    "           ,decode(A.随诊标志,1,'是',2,'是',3,'是','') as 是否随诊,A.编目员姓名 as 编目员,to_char(A.编目日期,'YYYY-MM-DD') as 编目日期" & _
    " From 病案主页 A,(Select * From 病人诊断记录 where 记录来源=4 and 诊断类型 in (3,13) and 病人ID=[1] and 主页ID=[2]) B,部门表 C,部门表 D,疾病编码目录 E,住院病案记录 F" & _
    " Where A.编目日期 is not null and A.病人ID=B.病人ID(+)  and A.主页ID=B.主页ID(+) and " & _
    "       A.入院科室ID=C.ID(+) and A.出院科室ID=D.ID and B.疾病ID=E.ID(+) and A.病人ID=[1] and A.主页ID=[2] and A.病人id=F.病人ID(+) and a.主页id=F.主页ID(+) "
        
    Set rs主页 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(tabMain.Tag), Val(Mid(tabMain.SelectedItem.Key, 2)))
    
    If rs主页.EOF = False Then
        For lngCount = 0 To rs主页.Fields.count - 1
            Set fld = rs主页.Fields(lngCount)
             If fld.Name = "病案号" Then
                Set lst = lvw主页.ListItems.Add(, "病案号", "病案号", "Attrib", "Attrib")
                lst.SubItems(1) = IIf(IsNull(fld.Value), "", fld.Value)
'            'ElseIf fld.Name = "病案号1" And IsNull(fld.Value) = True And gSystemPara.bln单独病案号 = True Then
'                Set lst = lvw主页.ListItems.Add(, "病案号", "病案号", "Attrib", "Attrib")
'                lst.SubItems(1) = IIf(IsNull(rs主页("病案号")), "", rs主页("病案号"))
'            'ElseIf gSystemPara.bln单独病案号 = False And fld.Name = "病案号" Then
'                Set lst = lvw主页.ListItems.Add(, "病案号", "病案号", "Attrib", "Attrib")
'                lst.SubItems(1) = IIf(IsNull(rs主页("病案号")), "", rs主页("病案号"))
            ElseIf fld.Name <> "病案号" Then
                Set lst = lvw主页.ListItems.Add(, fld.Name, fld.Name, "Attrib", "Attrib")
                lst.SubItems(1) = IIf(IsNull(fld.Value), "", fld.Value)
           End If
            '特殊字段的显示效果
            Select Case fld.Name
                Case "主要诊断编码", "主要诊断名", "主要诊断出院情况"
                    lst.ForeColor = RGB(0, 0, 213)
                    lst.ListSubItems(1).ForeColor = RGB(0, 0, 213)
            End Select
        Next
    End If
    '确定医保号
    gstrSQL = "Select 信息值 From 病案主页从表 where 信息名='医保号' and 病人ID=[1] and 主页ID=[2]"
    Set rs主页 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(tabMain.Tag), Val(Mid(tabMain.SelectedItem.Key, 2)))
    If Not rs主页.EOF Then
        str医保号 = zlCommFun.NVL(rs主页!信息值)
    Else
        str医保号 = ""
    End If
    Set lst = lvw主页.ListItems.Add(, "医保号", "医保号", "Attrib", "Attrib")
    lst.SubItems(1) = str医保号
    rs主页.Close
    
    FillDetail = True
    Call SetMenu
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    FillDetail = False
End Function

Private Sub SaveHead(ByVal vsGrid As VSFlexGrid, intListOrDetail As Integer)
    Dim strHeadInfo As String
     If intListOrDetail = 1 Then
        strHeadInfo = "病人列头信息"
    Else
        strHeadInfo = "病案列头信息"
    End If
    zl_VsGrid_SaveToPara vsGrid, Me.Caption, mlngMoudal, strHeadInfo, True, True
End Sub

Private Sub RestoreHead(ByVal vsGrid As VSFlexGrid, intListOrDetail As Integer)
    Dim strHeadInfo As String
    If intListOrDetail = 1 Then
        strHeadInfo = "病人列头信息"
    Else
        strHeadInfo = "病案列头信息"
    End If
    zl_VsGrid_FromParaRestore vsGrid, Me.Caption, mlngMoudal, strHeadInfo, True, True
End Sub

Private Sub tabMain_Click()
    Call FillDetail
End Sub


Public Sub SetMenu()
    Dim i As Long
    Dim int在院 As Long, int借出 As Long
    Dim blnCount  As Boolean, blnEnble As Boolean
   
    '问题24813
    Dim lngCount As Long
    
    blnEnble = True
    If vsfPatient.Rows = 2 Then
        If vsfPatient.TextMatrix(1, vsfPatient.ColIndex("选择")) = "" Then
            stbThis.Panels(2) = ""
            Exit Sub
        End If
    End If
    
    stbThis.Panels(2) = "当前显示有" & vsfPatient.Rows - 1 & "份病案。查询条件是“" & mstr显示 & "”"
    With vsfPatient
        lngCount = .Rows - 1
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("选择")) = True Then
                If .TextMatrix(i, .ColIndex("病案存储状态")) = "在院" Then
                    int在院 = int在院 + 1
                Else
                    int借出 = int借出 + 1
                End If
                If int在院 > 0 And int借出 > 0 Then Exit For
            End If
        Next
        If lngCount > 0 Then
            blnCount = .TextMatrix(.Row, .ColIndex("病案存储状态")) = "在院"
        End If
    End With
End Sub

Function CreateSQL() As Boolean
    Dim lngTemp As Long
    Dim strTemp As String
    Dim strFilterX As String
    Dim lngCount As Long
    
    If ValidSQL = False Then
        Exit Function
    End If
    mstrFilter = ""
    mstr条件 = ""
    strTemp = ""
    strFilterX = ""
    If lvwCombine.ListItems.count = 0 Then
        If mbln病案系统 Then
            mstrFilter = " A.编目日期 is not null and  A.封存时间 is  null"
        Else
             mstrFilter = " A.封存时间 is null "
        End If
        mstr条件 = "所有病案"
    Else
        strTemp = IIf(opt(0).Value = True, " AND ", " OR ")
        For lngTemp = 1 To lvwCombine.ListItems.count
            If InStr(lvwCombine.ListItems(lngTemp).Tag, "X.") = 0 Then
                mstrFilter = mstrFilter & "(" & lvwCombine.ListItems(lngTemp).Tag & ")" & strTemp
            Else
                strFilterX = strFilterX & "(" & lvwCombine.ListItems(lngTemp).Tag & ")" & strTemp
            End If
            mstr条件 = mstr条件 & "(" & lvwCombine.ListItems(lngTemp) & " " & lvwCombine.ListItems(lngTemp).SubItems(1) & ")" & IIf(opt(0).Value = True, " 并且 ", " 或者 ")
        Next
        If Len(mstrFilter) > 4 Then mstrFilter = Left(mstrFilter, Len(mstrFilter) - 4)
        If Len(strFilterX) > 4 Then strFilterX = Left(strFilterX, Len(strFilterX) - 4)
        mstr条件 = Left(mstr条件, Len(mstr条件) - 4)
        '得到表之间的关系
        strTemp = ""
        '要查询的表与引用表之间的联系
        If InStr(mstrFilter, "【J.】") > 0 Then strTemp = strTemp & "【C.】疾病ID=【J.】ID AND "
        If InStr(mstrFilter, "【I.】") > 0 Then strTemp = strTemp & "【B.】手术操作ID=【I.】ID AND "
        If InStr(mstrFilter, "【G.】") > 0 Then strTemp = strTemp & "【A.】出院科室ID=【G.】ID AND "
        If InStr(mstrFilter, "【E.】") > 0 Then strTemp = strTemp & "【A.】入院科室ID=【E.】ID AND "
        '各表与病案主页表之间的联系
        If InStr(mstrFilter & strTemp, "【B.】") > 0 Then strTemp = strTemp & "【A.】病人ID=【B.】病人ID AND 【A.】主页ID=【B.】主页ID AND "
        If InStr(mstrFilter & strTemp, "【C.】") > 0 Then strTemp = strTemp & "【A.】病人ID=【C.】病人ID AND 【A.】主页ID=【C.】主页ID AND "
        If InStr(mstrFilter & strTemp, "【O.】") > 0 Then strTemp = strTemp & "【A.】病人ID=【O.】病人ID AND 【A.】主页ID=【O.】主页ID AND "
        If InStr(mstrFilter & strTemp, "【L1.】") > 0 Then strTemp = strTemp & "【A.】病人ID=【L1.】病人ID AND 【A.】主页ID=【L1.】主页ID AND "
        If InStr(mstrFilter & strTemp, "【L2.】") > 0 Then strTemp = strTemp & "【A.】病人ID=【L2.】病人ID AND 【A.】主页ID=【L2.】主页ID AND "
        If InStr(mstrFilter & strTemp, "【L3.】") > 0 Then strTemp = strTemp & "【A.】病人ID=【L3.】病人ID AND 【A.】主页ID=【L3.】主页ID AND "
        If InStr(mstrFilter & strTemp, "【L4.】") > 0 Then strTemp = strTemp & "【A.】病人ID=【L4.】病人ID AND 【A.】主页ID=【L4.】主页ID AND "
        If InStr(mstrFilter & strTemp, "【L5.】") > 0 Then strTemp = strTemp & "【A.】病人ID=【L5.】病人ID AND 【A.】主页ID=【L5.】主页ID AND "
        If InStr(mstrFilter & strTemp, "【L6.】") > 0 Then strTemp = strTemp & "【A.】病人ID=【L6.】病人ID AND 【A.】主页ID=【L6.】主页ID AND "
        If InStr(mstrFilter & strTemp, "【L7.】") > 0 Then strTemp = strTemp & "【A.】病人ID=【L7.】病人ID AND 【A.】主页ID=【L7.】主页ID AND "
        If InStr(mstrFilter & strTemp, "【L8.】") > 0 Then strTemp = strTemp & "【A.】病人ID=【L8.】病人ID AND 【A.】主页ID=【L8.】主页ID AND "
        For lngCount = mlngIndex + 1 To lstFields.ListCount - 1
            If InStr(mstrFilter & strTemp, "【L" & lngCount & ".】") > 0 Then
                strTemp = strTemp & "【A.】病人ID=【L" & lngCount & ".】病人ID AND 【A.】主页ID=【L" & lngCount & ".】主页ID AND "
            End If
        Next lngCount
        If InStr(mstrFilter & strTemp, "【M.】") > 0 Then strTemp = strTemp & "【A.】病人ID=【M.】病人ID AND 【A.】主页ID=【M.】主页ID  AND "
        If InStr(mstrFilter & strTemp, "【N.】") > 0 Then strTemp = strTemp & "【A.】病人ID=【N.】病人ID AND "
        If InStr(mstrFilter & strTemp, "【P.】") > 0 Then strTemp = strTemp & "【A.】病人ID=【P.】病人ID AND "
        If InStr(mstrFilter & strTemp, "【P1.】") > 0 Then strTemp = strTemp & "【A.】病人ID=【P1.】病人ID AND "
        mstrFilter = strTemp & "(" & mstrFilter & ")"
        
        
        '得出要引用的表
        strTemp = ""
        If InStr(mstrFilter, "【P.】") > 0 Then strTemp = ",住院病案记录 P" & strTemp
        If InStr(mstrFilter, "【P1.】") > 0 Then strTemp = ",病人信息 P1" & strTemp
        If InStr(mstrFilter, "【N.】") > 0 Then strTemp = ",病人过敏药物 N" & strTemp
        If InStr(mstrFilter, "【M.】") > 0 Then strTemp = ",诊断符合情况 M" & strTemp
        If InStr(mstrFilter, "【L1.】") > 0 Then strTemp = ",病案主页从表 L1" & strTemp
        If InStr(mstrFilter, "【L2.】") > 0 Then strTemp = ",病案主页从表 L2" & strTemp
        If InStr(mstrFilter, "【L3.】") > 0 Then strTemp = ",病案主页从表 L3" & strTemp
        If InStr(mstrFilter, "【L4.】") > 0 Then strTemp = ",病案主页从表 L4" & strTemp
        If InStr(mstrFilter, "【L5.】") > 0 Then strTemp = ",病案主页从表 L5" & strTemp
        If InStr(mstrFilter, "【L6.】") > 0 Then strTemp = ",病案主页从表 L6" & strTemp
        If InStr(mstrFilter, "【L7.】") > 0 Then strTemp = ",病案主页从表 L7" & strTemp
        If InStr(mstrFilter, "【L8.】") > 0 Then strTemp = ",病案主页从表 L8" & strTemp
        For lngCount = mlngIndex + 1 To lstFields.ListCount - 1
            If InStr(mstrFilter & strTemp, "【L" & lngCount & ".】") > 0 Then
                 strTemp = ",病案主页从表 L" & lngCount & strTemp
            End If
        Next lngCount
        If InStr(mstrFilter, "【J.】") > 0 Then strTemp = ",疾病编码目录 J" & strTemp
        If InStr(mstrFilter, "【I.】") > 0 Then strTemp = ",疾病编码目录 I" & strTemp
        If InStr(mstrFilter, "【G.】") > 0 Then strTemp = ",部门表 G" & strTemp
        If InStr(mstrFilter, "【E.】") > 0 Then strTemp = ",部门表 E" & strTemp
        If InStr(mstrFilter, "【C.】") > 0 Then strTemp = ",病人诊断记录 C" & strTemp
        If InStr(mstrFilter, "【B.】") > 0 Then strTemp = ",病人手麻记录 B" & strTemp
        If InStr(mstrFilter, "【O.】") > 0 Then strTemp = ",借阅记录 O" & strTemp
    End If
    strTemp = "病案主页 A " & strTemp
    
    '问题31176 by lesfeng 2010-07-19 IIf(InStr(1, strTemp, "病人手麻记录") > 0, " And B.记录来源 = 4 ", "")
    mstrReturn = ",(select A.病人ID,A.主页ID from " & strTemp & " where A.编目日期 is not null " & IIf(InStr(1, strTemp, "借阅记录") > 0, " And O.归还时间 Is Null ", "") & IIf(InStr(1, strTemp, "病人手麻记录") > 0, " And B.记录来源 = 4 ", "") & IIf(InStr(1, strTemp, "病人诊断记录") > 0, " And C.记录来源=4 ", "") & IIf(mstrFilter = "()", "", " and " & mstrFilter) & ") Y where " & IIf(Len(strFilterX) = 0, "", "(" & strFilterX & ") and ")
    mstrSQL = "select X.住院号 as 住院号, A.病人ID,X.姓名,A.年龄,A.婚姻状况,A.国籍 from 病人信息 X," & strTemp & " where X.病人ID=A.病人ID and A.编目日期 is not null " & IIf(InStr(1, strTemp, "借阅记录") > 0, " And O.归还时间 Is Null ", "") & IIf(InStr(1, strTemp, "病人诊断记录") > 0, " And C.记录来源=4 ", "") & IIf(mstrFilter = "()", "", " and " & mstrFilter) & IIf(Len(strFilterX) = 0, "", " and (" & strFilterX & ")  ")
    mstrExecel = "select A.病人ID,A.主页ID from 病人信息 X," & strTemp & " where X.病人ID=A.病人ID and A.编目日期 is not null " & IIf(InStr(1, strTemp, "借阅记录") > 0, " And O.归还时间 Is Null ", "") & IIf(InStr(1, strTemp, "病人诊断记录") > 0, " And C.记录来源=4 ", "") & IIf(mstrFilter = "()", "", " and " & mstrFilter) & IIf(Len(strFilterX) = 0, "", " and (" & strFilterX & ")  ")
    
    mstrReturn = Replace(mstrReturn, "【", "")
    mstrReturn = Replace(mstrReturn, "】", "")
    mstrSQL = Replace(mstrSQL, "【", "")
    mstrSQL = Replace(mstrSQL, "】", "")
    
    mstrExecel = Replace(mstrExecel, "【", "")
    mstrExecel = Replace(mstrExecel, "】", "")
    
'    mblnChange = False
    CreateSQL = True
End Function

Private Function ValidSQL() As Boolean
    Dim bln符合情况 As Boolean, bln符合类型 As Boolean
    Dim lst As ListItem
    
    For Each lst In lvwCombine.ListItems
        If lst.Text = "符合情况" Then
            bln符合情况 = True
        Else
            If lst.Text = "符合类型" Then
                bln符合类型 = True
            End If
        End If
    Next
    
    If bln符合类型 Xor bln符合情况 Then
        MsgBox "符合类型与符合情况必须要同时存在。", vbInformation, gstrSysName
        Exit Function
    End If
    
    ValidSQL = True
End Function

Private Sub cmdAdd_Click()
    Dim bln已有 As Boolean
    Dim lst As ListItem
    
    Dim str对象名 As String
    Dim strField As String
    Dim strTemp As String
    
    '得出在数据库中表示的值
    If ValidateDate = False Then Exit Sub

    str对象名 = lstFields.List(lstFields.ListIndex)
    
    '查出以前是否已经有过类似条件
    For Each lst In lvwCombine.ListItems
        If lst.Text = str对象名 Then
            bln已有 = True
            Exit For
        End If
    Next
    
    strTemp = LeftB(cmbOperate.Text, 10) & IIf(cmbList.Visible = False, IIf(Trim(cmbExample.Text) = "", "空", cmbExample.Text), cmbList.Text)
    If bln已有 = False Then
        '新增一个该条件
        Set lst = lvwCombine.ListItems.Add(, , str对象名)
        lst.SubItems(1) = strTemp
    Else
        If lst.ListSubItems(1).Tag = "" Then
            '也只是增加了一次
            If cmbLogical.Text = "并且" Then
                lst.SubItems(1) = "(" & lst.SubItems(1) & ") 并且 (" & strTemp & ")"
            Else
                lst.SubItems(1) = "(" & lst.SubItems(1) & ") 或者 (" & strTemp & ")"
            End If
            lst.ListSubItems(1).Tag = "多次"
        Else
            If cmbLogical.Text = "并且" Then
                lst.SubItems(1) = lst.SubItems(1) & " 并且 (" & strTemp & ")"
            Else
                lst.SubItems(1) = lst.SubItems(1) & " 或者 (" & strTemp & ")"
            End If
        End If
    End If
    
    '为了把别名与普通字符串区分开，使用以下表示，在产生SQL语句时去掉
    Select Case str对象名
        Case "姓名", "性别", "身份证号", "出生日期"
            strField = "【X.】" & str对象名
        Case "姓名简码"
            strField = "zlspellcode(【X.】姓名)"
        Case "住院次数"
            strField = "【A.】主页ID"
        Case "住院号"
            strField = "【A.】住院号"
        Case "病案号"
            strField = "【P.】病案号"
        Case "档案号"
            strField = "【P.】档案号"
        Case "入院科室"
            strField = "【E.】名称"
        Case "出院科室"
            strField = "【G.】名称"
        Case "手术已行手术", "手术切口", "手术愈合", "手术麻醉类型"
            strField = "【B.】" & Mid(str对象名, 3)
        Case "手术日期", "主刀医师", "麻醉医师"
            strField = "【B.】" & str对象名
        Case "手术编码"
            strField = "【I.】编码"
        Case "手术简码"
            If mint简码方式 = 0 Then
                strField = "【I.】简码"
            Else
                strField = "【I.】五笔码"
            End If
        Case "诊断编码"
            strField = "【J.】编码"
        Case "诊断简码"
            If mint简码方式 = 0 Then
                strField = "【J.】简码"
            Else
                strField = "【J.】五笔码"
            End If
        Case "诊断描述信息", "诊断出院情况", "诊断编码序号"
            If str对象名 = "诊断描述信息" Then
                strField = "【C.】诊断描述"
            Else
                strField = "【C.】" & Mid(str对象名, 3)
            End If
            
        Case "诊断次序"
            strField = "【C.】诊断次序"
        Case "诊断类型"
            strField = "【C.】诊断类型"
        Case "未治"
            strField = "【C.】是否未治"
        Case "疑诊"
            strField = "【C.】是否疑诊"
        Case "符合类型"
            strField = "【M.】符合类型"
        Case "符合情况"
            strField = "【M.】符合情况"
        Case "病案质量"
            strField = "【L1.】信息名='病案质量' AND 【L1.】信息值"
        Case "科主任"
            strField = "【L2.】信息名='科主任' AND 【L2.】信息值"
        Case "主任医师"
            strField = "【L3.】信息名='主任医师' AND 【L3.】信息值"
        Case "主治医师"
            strField = "【L4.】信息名='主治医师' AND 【L4.】信息值"
        Case "医保号"
            strField = "【L3.】信息名='医保号' AND 【L3.】信息值"
        Case "输红细胞"
            strField = "【L5.】信息名='输红细胞' AND 【L5.】信息值"
        Case "输血小板"
            strField = "【L6.】信息名='输血小板' AND 【L6.】信息值"
        Case "输血浆"
            strField = "【L7.】信息名='输血浆' AND 【L7.】信息值"
        Case "输全血"
             strField = "【L8.】信息名='输全血' AND 【L8.】信息值"
        Case "过敏药物"
            strField = "【N.】过敏药物"
        Case "年龄"
            strField = "trunc((sysdate-【X.】出生日期)/365)"
        Case "借阅人", "借出人", "批准人", "借阅日期", "借阅部门", "借阅用途"
            If str对象名 = "借阅日期" Then
                strField = "【O.】借阅时间"
            Else
                strField = "【O.】" & str对象名
            End If
        Case Else
            If mlngIndex < lstFields.ListIndex Then
                strField = "【L" & lstFields.ListIndex & ".】信息名='" & str对象名 & "' AND 【L" & lstFields.ListIndex & ".】信息值"
            Else
                strField = "【A.】" & str对象名
            End If
    End Select
    If cmbList.Visible = True Then
        '固定取值列表
        If cmbList.Text = "" Then
            strTemp = IIf(Left(cmbOperate.Text, 2) = "等于", " IS ", " IS NOT ") & " NULL "
        Else
            strTemp = Right(cmbOperate.Text, 5) & cmbList.ItemData(cmbList.ListIndex)
            '刘兴宏:2007/11/05加入
            If str对象名 = "诊断类型" Then
                If cmbList.Text = "出院主要诊断" Then
                    strTemp = strTemp & " And 【C.】诊断次序 " & Right(cmbOperate.Text, 5) & "1"
                ElseIf cmbList.Text = "出院次要诊断" Then
                    strTemp = strTemp & IIf(Left(cmbOperate.Text, 2) = "等于", " And 【C.】诊断次序> 1", " And 【C.】诊断次序<=1")
                End If
            End If
        End If
    Else
        
        If Trim(cmbExample.Text) = "" Then
            strTemp = IIf(Left(cmbOperate.Text, 2) = "等于", " IS ", " IS NOT ") & " NULL "
        Else
            Select Case lstFields.ItemData(lstFields.ListIndex)
                Case 2 'Number
                    strTemp = Right(cmbOperate.Text, 5) & cmbExample.Text & " "
                Case 3 'VarChar
                    strTemp = Replace(cmbExample.Text, "'", "''")
                    Select Case str对象名
                        Case "身份证号", "姓名简码", "手术编码", "诊断编码", "手术简码", "诊断简码"
                            strTemp = UCase(strTemp)
                    End Select
                    If Left(cmbOperate, 2) = "包含" Then
                        strTemp = " LIKE '%" & strTemp & "%'"
                    ElseIf Left(cmbOperate, 2) = "开头" Then
                        strTemp = " LIKE '" & strTemp & "%'"
                    Else
                        strTemp = Right(cmbOperate.Text, 5) & "'" & strTemp & "'"
                    End If
                Case 4 'Date
                    strTemp = Right(cmbOperate.Text, 5) & "TO_DATE('" & Format(CDate(cmbExample.Text), "yyyy-MM-dd") & "','YYYY-MM-DD')"
            End Select
        End If
    End If
    If bln已有 Then
        '在前一值上累加
        If cmbLogical.Text = "并且" Then
            If lstFields.ItemData(lstFields.ListIndex) = 4 Then
                    lst.Tag = lst.Tag & " AND (trunc(" & strField & ") " & strTemp & ")"
            ElseIf str对象名 = "病案号" Then
                    lst.Tag = lst.Tag & " AND (" & strField & strTemp & ") AND (【A.】病案号" & strTemp & ")"
            Else
                    lst.Tag = lst.Tag & " AND (" & strField & strTemp & ")"
            End If
        Else
            If lstFields.ItemData(lstFields.ListIndex) = 4 Then
                    lst.Tag = lst.Tag & " OR (trunc(" & strField & ") " & strTemp & ")"
            ElseIf str对象名 = "病案号" Then
                lst.Tag = lst.Tag & " OR (" & strField & strTemp & ") or (【A.】病案号" & strTemp & ")"
            Else
                lst.Tag = lst.Tag & " OR (" & strField & strTemp & ")"
            End If
        End If
    Else
        If lstFields.ItemData(lstFields.ListIndex) = 4 Then
            lst.Tag = "(trunc(" & strField & ") " & strTemp & ")"
        ElseIf str对象名 = "病案号" Then
            lst.Tag = lst.Tag & "(" & strField & strTemp & " or 【A.】病案号" & strTemp & ")"
        Else
            lst.Tag = "(" & strField & strTemp & ")"
        End If
    End If
    cmd(2).Enabled = True
End Sub

Private Function ValidateDate() As Boolean
    
    If cmbExample.Visible = True Then
        If (Left(cmbOperate.Text, 2) <> "等于" And Left(cmbOperate.Text, 2) <> "不等") And Trim(cmbExample.Text = "") Then
            MsgBox "当前条件下不允许空值。" & vbCrLf & "如果你要使用空值，那条件只能是“等于”或“不等于”。", vbInformation, gstrSysName
            cmbExample.SetFocus
            Exit Function
        End If
        If zlCommFun.ActualLen(cmbExample.Text) > 100 Then
            MsgBox "输入内容过长。", vbInformation, gstrSysName
            cmbExample.SetFocus
            Exit Function
        End If
        Select Case lstFields.ItemData(lstFields.ListIndex)
            Case 2 'Number
                If Not IsNumeric(cmbExample.Text) And Trim(cmbExample.Text) <> "" Then
                    MsgBox "请输入一个合法的数字。", vbInformation, gstrSysName
                    cmbExample.SetFocus
                    Exit Function
                End If
            Case 3 'VarChar
                If InStr(cmbExample.Text, "'") > 0 Then
                    MsgBox "输入了非法字符。", vbInformation, gstrSysName
                    cmbExample.SetFocus
                    Exit Function
                End If
            Case 4 'Date
                If Not IsDate(cmbExample.Text) And Trim(cmbExample.Text) <> "" Then
                    MsgBox "请输入一个合法的日期。" & vbCrLf & "比如：1997-07-01。", vbInformation, gstrSysName
                    cmbExample.SetFocus
                    Exit Function
                End If
        End Select
    End If
    ValidateDate = True
End Function