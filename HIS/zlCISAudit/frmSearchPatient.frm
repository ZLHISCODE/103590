VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearchPatient 
   Caption         =   "��������"
   ClientHeight    =   7845
   ClientLeft      =   2835
   ClientTop       =   3825
   ClientWidth     =   15015
   Icon            =   "frmSearchPatient.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   15015
   StartUpPosition =   2  '��Ļ����
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
         ToolTipText     =   "ɾ������"
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
         ToolTipText     =   "��������"
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
         ToolTipText     =   "��������"
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
         ToolTipText     =   "���˹�����һ�л�����ڲ�ѯ�����"
         Top             =   45
         Width           =   5280
      End
      Begin VB.OptionButton opt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "������һ����(&P)"
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
         Caption         =   "����ȫ������(&A)"
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
         Caption         =   "ˢ��(&R)"
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
            Text            =   "����"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "����"
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
         Caption         =   "ˢ��(&R)"
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
         Caption         =   "�� �� ��(&H)"
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
         Caption         =   "ס Ժ ��(&Z)"
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
         Caption         =   "��    ��(&B)"
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
         Caption         =   "��    ��(&N)"
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
         Caption         =   "����״��(&I)"
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
         Caption         =   "��    ��(&X)"
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
         Caption         =   "��    ��(&K)"
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
         Caption         =   "��"
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
            Name            =   "����"
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
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   23574
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      Begin MSComctlLib.ListView lvw��ҳ 
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
            Text            =   "��Ϣ��"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "��Ϣֵ"
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
               Caption         =   "��1����Ժ"
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
    �������� As String
End Type

Private mlngOldRow As Long
Private mlngNewRow As Long
Private mstrFilter As String
Private mstr��ʾ As String
Private mbln����ϵͳ As Boolean '��鲡��ϵͳ�Ƿ����
Private mlngIndex As Long
Private mint���뷽ʽ As Integer

Private mstr���� As String
Private mstrReturn As String
Private mstrSQL As String
Private mstrExecel As String

Private usrSaveItem As Items

'######################################################################################################################

Public Function ShowEdit(ByVal frmMain As Object, ByRef rsPatient As ADODB.Recordset, ByVal lngMoudal As Long, ByVal strPrivs As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mlngMoudal = lngMoudal
    mblnOK = False
    mstrPrivs = strPrivs
    
    Set mfrmMain = frmMain
    If ExecuteCommand("��ʼ�ؼ�") = False Then Exit Function
    If ExecuteCommand("��ʼ����") = False Then Exit Function
    
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
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    
    '------------------------------------------------------------------------------------------------------------------
    '��ʼ����
    
    Call CommandBarInit(cbsMain)
    
    '------------------------------------------------------------------------------------------------------------------
    '�˵�����:�����������ݣ����xtpControlPopup���͵�����ID���¸�ֵ
    
    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    '�ļ�
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Excel, "�����&Excel")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Exit, "�˳�(&X)", True)
    
    '�༭
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_SelAll, "ȫ��ѡ��(&A)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_ClsAll, "ȫ����ѡ(&D)")
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_SaveExit, "����ѡ��(&S)", True)
       
    '�鿴
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)", True)

            
    '����
    '------------------------------------------------------------------------------------------------------------------
    Call CreateHelpMenu(cbsMain)
    
    '------------------------------------------------------------------------------------------------------------------
    '����������:������������

    Set objBar = cbsMain.Add("������", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched

    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Print, "��ӡ")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Preview, "Ԥ��")
    
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_SelAll, "ȫѡ(&A)", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_ClsAll, "ȫ��(&D)")
    
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_SaveExit, "����(&S)", True)

    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Help_Help, "����(&H)", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "�˳�(&X)")

    '------------------------------------------------------------------------------------------------------------------
    '����Ŀ����:���������������Ѵ���

    With cbsMain.KeyBindings
        .Add 0, vbKeyF5, conMenu_View_Refresh           'ˢ��
        .Add 0, vbKeyF1, conMenu_Help_Help              '����
        .Add 0, vbKeyF2, conMenu_Edit_SaveExit              '����
        .Add FCONTROL, vbKeyA, conMenu_Edit_SelAll
    End With

End Function

Private Sub InitDockPannel()
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    Dim objPane As Pane

    Set objPane = dkpMain.CreatePane(1, 100, 100, DockLeftOf, Nothing): objPane.Title = "����": objPane.Options = PaneNoCaption
    Set objPane = dkpMain.CreatePane(2, 120, 100, DockRightOf, Nothing): objPane.Title = "����": objPane.Options = PaneNoCaption
    Set objPane = dkpMain.CreatePane(3, 80, 100, DockRightOf, Nothing): objPane.Title = "��Ϣ": objPane.Options = PaneNoCaption
    
    dkpMain.SetCommandBars cbsMain
    Call DockPannelInit(dkpMain)

End Sub

Private Function InitToolBox() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    Dim objGrp As TaskPanelGroup
    Dim objItem As TaskPanelGroupItem
    Dim objIlsItem As Object
    

    '����������
    Set objGrp = tpl.Groups.Add(0, "����������")
    objGrp.Expandable = False
    objGrp.Expanded = True
    Set objItem = objGrp.Items.Add(0, "��", xtpTaskItemTypeControl, 1)
    objItem.Handle = picPane(2).hWnd
    picPane(2).BackColor = objItem.BackColor
    
    '�߼�����
    Set objGrp = tpl.Groups.Add(1, "�߼�������")
    objGrp.Expandable = False
    objGrp.Expanded = True
    Set objItem = objGrp.Items.Add(0, "��", xtpTaskItemTypeControl, 1)
    objItem.Handle = picPane(3).hWnd
    picPane(3).BackColor = objItem.BackColor
    opt(0).BackColor = objItem.BackColor
    opt(1).BackColor = objItem.BackColor
    
    '��ʷ����
    Set objGrp = tpl.Groups.Add(1, "����������")
    Set objItem = objGrp.Items.Add(0, "��", xtpTaskItemTypeControl, 1)
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
     '��ʼ�����ڿؼ���ֵ
    
    '0  ������
    '1  ������
    '2  ����
    '3  �ַ�
    '4  ����
    '5  �̶�ȡֵ
    lstFields.Clear
    'һ�����ü�������
    lstFields.AddItem "סԺ��":   lstFields.ItemData(lstFields.NewIndex) = 2 ' Number(18)"
    lstFields.AddItem "������":   lstFields.ItemData(lstFields.NewIndex) = 3 ' varchar2(20)"
    lstFields.AddItem "������":   lstFields.ItemData(lstFields.NewIndex) = 3 ' varchar2(20)"
    lstFields.AddItem "סԺ����": lstFields.ItemData(lstFields.NewIndex) = 2 ' Number(18)"
    lstFields.AddItem "����":     lstFields.ItemData(lstFields.NewIndex) = 3     ' VARCHAR2(10)"
    lstFields.AddItem "��������": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(10)"
    lstFields.AddItem "��Ժ����": lstFields.ItemData(lstFields.NewIndex) = 4 ' Date"
    lstFields.AddItem "��Ժ����": lstFields.ItemData(lstFields.NewIndex) = 3 ' Number(18)"
    lstFields.AddItem "��ĿԱ����": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(10)"
    lstFields.AddItem "��Ŀ����": lstFields.ItemData(lstFields.NewIndex) = 4 ' Date"
    '�����������
    lstFields.AddItem " "
    lstFields.AddItem "�������": lstFields.ItemData(lstFields.NewIndex) = 5
    lstFields.AddItem "��ϱ���": lstFields.ItemData(lstFields.NewIndex) = 3 ' Number(5)"
    lstFields.AddItem "��ϼ���": lstFields.ItemData(lstFields.NewIndex) = 3
    lstFields.AddItem "���������Ϣ": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(6)"
    lstFields.AddItem "��ϳ�Ժ���": lstFields.ItemData(lstFields.NewIndex) = 3  ' VARCHAR2
    lstFields.AddItem "��ϴ���": lstFields.ItemData(lstFields.NewIndex) = 2 ' Number
    lstFields.AddItem "��ϱ������": lstFields.ItemData(lstFields.NewIndex) = 2 ' Number
'    If gSystemPara.bln������ҽ = True Then
        lstFields.AddItem "��ҽ����": lstFields.ItemData(lstFields.NewIndex) = 3  ' VARCHAR2
        lstFields.AddItem "��ҽ�������": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2
'    End If
    lstFields.AddItem "δ��": lstFields.ItemData(lstFields.NewIndex) = 1
    lstFields.AddItem "����": lstFields.ItemData(lstFields.NewIndex) = 1
    lstFields.AddItem "��������": lstFields.ItemData(lstFields.NewIndex) = 5
    lstFields.AddItem "�������": lstFields.ItemData(lstFields.NewIndex) = 5
   
    '������������
    lstFields.AddItem " "
    lstFields.AddItem "��������": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2
    lstFields.AddItem "��������": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2
    lstFields.AddItem "������������": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(6)"
    lstFields.AddItem "�����п�": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2
    lstFields.AddItem "��������": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2
    lstFields.AddItem "������������": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2
    lstFields.AddItem "��������": lstFields.ItemData(lstFields.NewIndex) = 4 ' VARCHAR2(10)"
    lstFields.AddItem "����ҽʦ": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(50)"
    lstFields.AddItem "����ҽʦ": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(50)"
    '�ġ�������Ϣ
    lstFields.AddItem " "
    lstFields.AddItem "���֤��": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(18)"
    lstFields.AddItem "����": lstFields.ItemData(lstFields.NewIndex) = 2 ' VARCHAR2(10)"
    lstFields.AddItem "�Ա�": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(4)"
    lstFields.AddItem "����״��": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(4)"
    lstFields.AddItem "ҽ�Ƹ��ʽ": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(20)"
    lstFields.AddItem "ְҵ": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(40)"
    lstFields.AddItem "����": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(30)"
    lstFields.AddItem "Ѫ��": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(10)"
    lstFields.AddItem "��λ�绰": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(50)"
    lstFields.AddItem "��λ�ʱ�": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(6)"
    lstFields.AddItem "��λ��ַ": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(50)"
    lstFields.AddItem "����": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(10)"
    lstFields.AddItem "��ͥ��ַ": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(50)"
    lstFields.AddItem "��ͥ�绰": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(20)"
    lstFields.AddItem "�����ʱ�": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(6)"
    lstFields.AddItem "��ϵ������": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(10)"
    lstFields.AddItem "��ϵ�˹�ϵ": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(10)"
    lstFields.AddItem "��ϵ�˵�ַ": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(50)"
    lstFields.AddItem "��ϵ�˵绰": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(20)"
     lstFields.AddItem "��������": lstFields.ItemData(lstFields.NewIndex) = 4 ' date"
    '�塢������ҳ
    lstFields.AddItem " "
    lstFields.AddItem "��Ժ����": lstFields.ItemData(lstFields.NewIndex) = 3 ' Number(18)"
    lstFields.AddItem "��Ժ����": lstFields.ItemData(lstFields.NewIndex) = 4 ' Date"
    lstFields.AddItem "��Ժ����": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(10)"
    lstFields.AddItem "סԺ����": lstFields.ItemData(lstFields.NewIndex) = 2 ' Number(18)"
    lstFields.AddItem "ȷ������": lstFields.ItemData(lstFields.NewIndex) = 4 ' Date"
    lstFields.AddItem "�����־": lstFields.ItemData(lstFields.NewIndex) = 1
    lstFields.AddItem "��������": lstFields.ItemData(lstFields.NewIndex) = 2 ' Number(18)"
    lstFields.AddItem "���ȴ���": lstFields.ItemData(lstFields.NewIndex) = 2 ' Number(5)"
    lstFields.AddItem "�ɹ�����": lstFields.ItemData(lstFields.NewIndex) = 2 ' Number(5)"
    lstFields.AddItem "ʬ���־": lstFields.ItemData(lstFields.NewIndex) = 1
    lstFields.AddItem "���ú�":   lstFields.ItemData(lstFields.NewIndex) = 2 ' Number(16, 5)"
    lstFields.AddItem "����ҩ��": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(50)"
    lstFields.AddItem "��������": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2(2)"
    lstFields.AddItem "������":   lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2
    lstFields.AddItem "����ҽʦ": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2
    lstFields.AddItem "ҽ����":   lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2
    lstFields.AddItem "����ҽʦ": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2
    lstFields.AddItem "סԺҽʦ": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2
    lstFields.AddItem "����ҽʦ": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2
    '����31031 by lesfeng 2010-06-24 ������ʲ�����ҳ�ӱ�����ܣ��������εĴ���
    mlngIndex = lstFields.NewIndex
    
    lstFields.AddItem "����ҽʦ": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "�о���ʵϰҽʦ": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "ʵϰҽʦ": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "�ʿ�ҽʦ": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "�ʿػ�ʿ": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    '������Ѫ���
    lstFields.AddItem " "
    lstFields.AddItem "HBsAg": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "HCV-Ab": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "HIV-Ab": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "Rh": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "��Ѫ���": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "��Һ��Ӧ": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "��Ѫ��Ӧ": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "���ϸ��": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "������": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "��Ѫ��": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "��ȫѪ": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "��ѪС��": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    
'    If gSystemPara.bln������ҽ = True Then
        lstFields.AddItem " "
        lstFields.AddItem "��ҽ���ȷ���": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
        lstFields.AddItem "������ҩ�Ƽ�": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
        lstFields.AddItem "��ҽΣ��": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
        lstFields.AddItem "��ҽ����": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
        lstFields.AddItem "��ҽ��֢": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
'    End If

    '�ߡ��������
    lstFields.AddItem " "
    lstFields.AddItem "��Ժ����": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "��Ժ����": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "�ջ�����": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "�����": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "���в���": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "ʾ�̲���": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "��Ժǰ����Ժ����": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "���Ѳ���": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "����": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "ת�Ƽ�¼": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "ת��ʱ��": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "��ԭѧ���": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "��Ժ��ʽ": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "��������ԭ��": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "��Ժ��ʽ": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "��Ⱦ��������ϵ": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "��Ⱦ��λ": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "��Ժ��ʽ_��ҳ": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "ת���������": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "���Ȳ���": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "��ҳX�ߺ�": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "��ҳ��������": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "������ʹ��ʱ��": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "����ʱ��": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "�ؼ���������": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "һ����������": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "������������": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "������������": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "ICU����": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "CCU����": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "CT": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "MRI": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "��ɫ������": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "������4": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "������5": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "������6": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    '�ˡ��������
    lstFields.AddItem " "
    lstFields.AddItem "����ʱ��": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "�������": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "̥��": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "̥��": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "����ʱ��1": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "����ʱ��2": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "�ܲ���ʱ��": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "�����Ѫ��": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "���Ʋ���֢": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "�����������": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "סԺ�����ڼ�": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "�ֻ��̶�": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "����������": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "������������": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "��������������": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "��������Ժ����": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    
    '�š��������
    lstFields.AddItem ""
    lstFields.AddItem "������": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "�����": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "��׼��": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "��������": lstFields.ItemData(lstFields.NewIndex) = 4 ' DATE"
    lstFields.AddItem "���Ĳ���": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    lstFields.AddItem "������;": lstFields.ItemData(lstFields.NewIndex) = 3 ' VARCHAR2"
    
    On Error GoTo errHandle
    strSQL = "Select Distinct(����) ���� From ������Ŀ" '�Զ�������ݲ��Ǻܶ࣬û������
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)

    lstFields.AddItem " "

    While Not rsTemp.EOF
        lstFields.AddItem rsTemp!����: lstFields.ItemData(lstFields.NewIndex) = 3
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
    Dim lngֵ���� As Long
    Dim str�ֶ��� As String

    If cmbOperate.ListCount = 7 Then cmbOperate.RemoveItem 6: cmbOperate.ListIndex = 0
    With lstFields
        lngֵ���� = .ItemData(.ListIndex)

        If lngֵ���� = 0 Then
            '���в���ȡֵ
            .ListIndex = .ListIndex - 1
            Exit Sub

        ElseIf lngֵ���� = 1 Or lngֵ���� = 5 Then
            cmbOperate.Clear
            cmbOperate.AddItem "����                             ="
            cmbOperate.AddItem "������                          <>"
            cmbOperate.ListIndex = 0

            '�̶�ȡֵ
            cmbList.Clear
            If lngֵ���� = 1 Then
                cmbList.AddItem "��": cmbList.ItemData(cmbList.NewIndex) = 1
                cmbList.AddItem "��": cmbList.ItemData(cmbList.NewIndex) = 0
            Else
                Setֵ�б� .List(.ListIndex)
            End If
            cmbList.AddItem ""
            cmbList.ListIndex = 0

            cmbList.Visible = True
            cmbExample.Visible = False
        Else
            cmbOperate.Clear
            cmbOperate.AddItem "����                             ="
            cmbOperate.AddItem "������                          <>"
            cmbOperate.AddItem "С��                             <"
            cmbOperate.AddItem "С�ڵ���                        <="
            cmbOperate.AddItem "����                             >"
            cmbOperate.AddItem "���ڵ���                        >="
            If lngֵ���� = 3 Then cmbOperate.AddItem "����                          LIKE"
            If lngֵ���� = 3 Then cmbOperate.AddItem "��ͷ                          LIKE"
            cmbOperate.ListIndex = 0

            cmbList.Visible = False
            cmbExample.Visible = True

            Dim rsExample As New ADODB.Recordset
            rsExample.CursorLocation = adUseClient
            str�ֶ��� = .List(.ListIndex)
            Select Case str�ֶ���
                Case "סԺ��", "����", "���֤��", "��������"
                    strExample = "select distinct " & str�ֶ��� & " from ������Ϣ where " & str�ֶ��� & " is not null and  rownum<51"
                Case "��������"
                    strExample = "select distinct zlSpellcode(����) from ������Ϣ where ���� is not null and  rownum<51"
                Case "�Ա�"
                    strExample = "select distinct " & str�ֶ��� & " from ������Ϣ "
                Case "סԺ����"
                    strExample = "select distinct ��ҳID from ������ҳ where rownum<51"
                Case "������"
                    strExample = "select distinct ������ from סԺ������¼ where rownum<51"
                Case "������"
                    strExample = "select distinct ������ from סԺ������¼ where rownum<51"
                Case "��Ժ����", "��Ժ����"
                    strExample = "select A.���� from ���ű� A,��������˵�� B " & _
                                " where A.ID=B.����ID And B.��������='�ٴ�' and (B.�������=2 or B.�������=3) and " & Where����ʱ��("A") & zl_��ȡվ������(True, "a") & " order by A.����"
                Case "������������", "�����п�", "��������", "������������"
                    strExample = "select distinct " & Mid(str�ֶ���, 3) & " from ���������¼ where ��¼��Դ=4 and rownum<51"
                Case "��������", "����ҽʦ", "����ҽʦ"
                    strExample = "select distinct " & str�ֶ��� & " from ���������¼ where ��¼��Դ=4 and rownum<51"
                Case "��������"
                    strExample = "select distinct B.���� from ���������¼ A,��������Ŀ¼ B where A.��������ID=B.ID and A.��¼��Դ=4 and rownum<51"
                Case "��������"
                    If mint���뷽ʽ = 0 Then
                        strExample = "select distinct B.���� from ���������¼ A,��������Ŀ¼ B where A.��������ID=B.ID and A.��¼��Դ=4 and rownum<51"
                    Else
                        strExample = "select distinct B.����� from ���������¼ A,��������Ŀ¼ B where A.��������ID=B.ID and A.��¼��Դ=4 and rownum<51"
                    End If

                Case "���������Ϣ", "��ϳ�Ժ���", "��ϱ������"
                    If str�ֶ��� = "���������Ϣ" Then
                        strExample = "select distinct ������� from ������ϼ�¼ where rownum<51"
                    ElseIf str�ֶ��� = "��ϱ������" Then
                        strExample = "select distinct " & Mid(str�ֶ���, 3) & " from ������ϼ�¼ where ��¼��Դ=4 and  rownum<51"
                    Else
                        strExample = "Select ���� From ���ƽ��"
                    End If
                Case "��ϴ���"
                    strExample = "select distinct ��ϴ��� from ������ϼ�¼ where   ��¼��Դ=4 and  rownum<51"
                Case "��ϱ���"
                    strExample = "select distinct B.���� from ������ϼ�¼ A,��������Ŀ¼ B where A.��¼��Դ=4 and  A.����ID=B.ID and rownum<51"
                Case "��ϼ���"
                    If mint���뷽ʽ = 0 Then
                        strExample = "select distinct  B.���� from ������ϼ�¼ A,��������Ŀ¼ B where A.��¼��Դ=4 and  A.����ID=B.ID and rownum<51"
                    Else
                        strExample = "select distinct  B.����� from ������ϼ�¼ A,��������Ŀ¼ B where A.��¼��Դ=4 and  A.����ID=B.ID and rownum<51"
                    End If
'                Case "��ҽ��ϱ���"
'                    strExample = "select distinct B.���� from ������ϼ�¼ A,��������Ŀ¼ B where A.����ID=B.Id And A.������� In (11,12,13) and rownum<51"
'                Case "��ҽ�������"
'                    strExample = "select distinct ������� from ������ϼ�¼ where ������� In (11,12,13) and rownum<51"
'                Case "��ҽ��Ժ���"
'                    strExample = "Select ���� From ���ƽ��"
                Case "��ҽ����"
                    strExample = "select distinct B.���� from ������ϼ�¼ A,��������Ŀ¼ B where  A.��¼��Դ=4 and  A.֤��ID=B.ID and rownum<51"
                Case "����ҩ��"
                    strExample = "select distinct ����ҩ�� from ���˹���ҩ�� where rownum<51"
                Case "��������", "������", "����ҽʦ", "����ҽʦ", "ҽ����", "���ϸ��", "��ѪС��", "��Ѫ��", "��ȫѪ"
                    strExample = "select distinct ��Ϣֵ from ������ҳ�ӱ� where ��Ϣ��='" & str�ֶ��� & "' and rownum<51"
                Case "����"
                    strExample = "select distinct trunc((sysdate-��������)/365) from ������Ϣ where rownum<51"
                Case "������", "�����", "��׼��", "��������", "���Ĳ���", "������;"
                    If str�ֶ��� = "��������" Then
                        strExample = "select distinct ����ʱ�� from ������ҳ A ,���ļ�¼ O where A.����ID=O.����ID AND A.��ҳID=O.��ҳID And A.��Ŀ���� is not null And �黹ʱ�� Is Null And rownum<51"
                    Else
                        strExample = "select distinct " & str�ֶ��� & " from ������ҳ A ,���ļ�¼ O where A.����ID=O.����ID AND A.��ҳID=O.��ҳID And A.��Ŀ���� is not null And �黹ʱ�� Is Null And rownum<51"
                    End If
                Case Else
                    If mlngIndex < lstFields.ListIndex Then
                        strExample = "select distinct ��Ϣֵ from ������ҳ�ӱ� where ��Ϣ��='" & str�ֶ��� & "' and rownum<51"
                    Else
                        strExample = "select distinct " & str�ֶ��� & " from ������ҳ where rownum<51"
                    End If
            End Select

            On Error GoTo errHandle

            '�õ�һЩ����
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

Private Sub Setֵ�б�(ByVal str���� As String)
    With cmbList
        Select Case str����
            Case "�������"
                .AddItem "�������": .ItemData(.NewIndex) = 1
                .AddItem "��Ժ���": .ItemData(.NewIndex) = 2
                'by gzh
                .AddItem "��Ժ��Ҫ���": .ItemData(.NewIndex) = 3
                .AddItem "��Ժ��Ҫ���": .ItemData(.NewIndex) = 3
                .AddItem "Ժ�ڸ�Ⱦ": .ItemData(.NewIndex) = 5
                .AddItem "�������": .ItemData(.NewIndex) = 6
                .AddItem "�����ж���": .ItemData(.NewIndex) = 7
'                If gSystemPara.bln������ҽ = True Then
                    .AddItem "��ҽ�������": .ItemData(.NewIndex) = 11
                    .AddItem "��ҽ��Ժ���": .ItemData(.NewIndex) = 12
                    .AddItem "��ҽ��Ժ���": .ItemData(.NewIndex) = 13
                    .AddItem "��ҽ��֤���": .ItemData(.NewIndex) = 14
'                End If

            Case "��������"
                .AddItem "�������Ժ": .ItemData(.NewIndex) = 1
                .AddItem "��Ժ���Ժ": .ItemData(.NewIndex) = 2
                .AddItem "�����벡��": .ItemData(.NewIndex) = 3
                .AddItem "�ٴ��벡��": .ItemData(.NewIndex) = 4
                .AddItem "�ٴ���ʬ��": .ItemData(.NewIndex) = 5
                .AddItem "��ǰ������": .ItemData(.NewIndex) = 6
'                If gSystemPara.bln������ҽ = True Then
                    .AddItem "��ҽ�������Ժ": .ItemData(.NewIndex) = 11
                    .AddItem "��ҽ��Ժ���Ժ": .ItemData(.NewIndex) = 12
'                End If
            Case "�������"
                .AddItem "����": .ItemData(.NewIndex) = 1
                .AddItem "������": .ItemData(.NewIndex) = 2
                .AddItem "���϶�": .ItemData(.NewIndex) = 3
        End Select
    End With
End Sub


Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '--------------------------------------------------------------------------------------------------------------
    '���ܣ�
    '������
    '���أ�
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
    Case "��ʼ�ؼ�"
        Set mclsPatient = New clsVsf
        With mclsPatient
            Call .Initialize(Me.Controls, vsfPatient, True, True, frmPubResource.GetImageList(16))
            Call .ClearColumn
            Call .AppendColumn("���", 500, flexAlignCenterCenter, flexDTString, "", "", False)
            Call .AppendColumn("ѡ��", 500, flexAlignCenterCenter, flexDTBoolean, "", , True)
            Call .AppendColumn("סԺ��", 1500, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("����", 900, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("�Ա�", 500, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("����", 500, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("��������", 1100, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("סԺ����", 900, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("��ͥ��ַ", 1350, flexAlignLeftCenter, flexDTString, "", , True)
            
            Call .AppendColumn("����״��", 900, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("��ͥ�绰", 1000, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("������λ", 1200, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("��λ�绰", 1000, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("��ϵ������", 1000, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("��ϵ�˹�ϵ", 1200, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("��ϵ�˵�ַ", 1200, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("��ϵ�˵绰", 1000, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("������", 1000, flexAlignLeftCenter, flexDTString, "", , True)
            
            Call .AppendColumn("�����ż���", 1200, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("�����洢״̬", 1440, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("�������λ��", 1440, flexAlignLeftCenter, flexDTString, "", , True)
            
            Call .AppendColumn("����id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("��ҳid", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            
            Call .AppendColumn("��Ժʱ��", 1670, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm", , True)
            Call .AppendColumn("��Ժʱ��", 1670, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm", , True)
            Call .AppendColumn("��Ժ����", 1200, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("���״̬", 0, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("���������", 0, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("", 0, flexAlignLeftCenter, flexDTString, "", , True)
            
            Call .InitializeEdit(True, False, False)
            Call .InitializeEditColumn(.ColIndex("ѡ��"), True, vbVsfEditCheck)
            
            .AppendRows = True
        End With
        
        Call InitToolBox
        
        mint���뷽ʽ = Val(zlDatabase.GetPara("���뷽ʽ"))
        
    '--------------------------------------------------------------------------------------------------------------
    Case "��ʼ����"

        '�Ա�
        cbo(0).Clear
        cbo(0).AddItem ""
        Set rs = gclsPackage.GetBaseCode("�Ա�")
        If rs.BOF = False Then Call AddComboData(cbo(0), rs, "����", , "ȱʡ��־", False)
    
        '����״��
        cbo(1).Clear
        cbo(1).AddItem ""
        Set rs = gclsPackage.GetBaseCode("����״��")
        If rs.BOF = False Then Call AddComboData(cbo(1), rs, "����", , "ȱʡ��־", False)
        
        '�ٴ�����
        cbo(2).Clear
        cbo(2).AddItem ""
        Set rs = gclsPackage.GetDept("�ٴ�")
        If rs.BOF = False Then Call AddComboData(cbo(2), rs, "����", "ID", , False)
        
        cmbLogical.Clear
        cmbLogical.AddItem "����"
        cmbLogical.AddItem "����"
        cmbLogical.ListIndex = 0
        
        Set mrsPatient = New ADODB.Recordset
        With mrsPatient
            .Fields.Append "ID", adVarChar, 30
            .Fields.Append "����id", adVarChar, 25
            .Fields.Append "��ҳid", adVarChar, 10
            .Fields.Append "�Ա�", adVarChar, 50
            .Fields.Append "����", adVarChar, 30
            .Fields.Append "����", adVarChar, 30
            .Fields.Append "����״��", adVarChar, 50
            .Fields.Append "��Ժʱ��", adVarChar, 30
            .Fields.Append "��Ժʱ��", adVarChar, 30
            .Fields.Append "��Ժ����", adVarChar, 50
            
            .Fields.Append "סԺ��", adVarChar, 50
            .Fields.Append "������", adVarChar, 50
            .Fields.Append "סԺ����", adVarChar, 50
            
            .Open
        End With
        
        strTmp = zlDatabase.GetPara("��������", glngSys, mlngMoudal, "", Array(cmd(4), cmd(5), cmd(6)), IsPrivs(mstrPrivs, "��������"))
        If strTmp <> "" Then
            For intRow = 0 To UBound(Split(strTmp, "|")) Step 2
                Set objItem = lvw.ListItems.Add(, , Split(strTmp, "|")(intRow), 1, 1)
                objItem.Tag = Split(strTmp, "|")(intRow + 1)
            Next
        End If
        
        '��鲡��ϵͳ�Ƿ����
        Set rs = gclsPackage.GetMedicalExits
        If Not rs.EOF Then
            mbln����ϵͳ = True
        Else
            mbln����ϵͳ = False
        End If
        
    '--------------------------------------------------------------------------------------------------------------
    Case "ˢ������"
        
        If CreateSQL = False Then
            Exit Function
        End If
        strWhere = ""
        mstrFilter = mstrReturn
        mstr��ʾ = mstr����
        Call FillData(strWhere)
        
        mclsPatient.AppendRows = True
        DataChanged = False
    '--------------------------------------------------------------------------------------------------------------
    Case "��������"
        strCustom = " And (1=1)"
        mclsPatient.ClearGrid
        mstr��ʾ = ""
        strWhere = GetCustomWhere(Val(cmd(0).Tag), cbo(0).Text, cbo(2).ItemData(cbo(2).ListIndex), Val(txt(6).Text), Val(txt(4).Text), cbo(1).Text, txt(1).Text, txt(2).Text)
       
        '������Ϣ����
        If mbln����ϵͳ Then
            '����ϵͳ�Ѿ���װ,��鲡���Ƿ��Ŀ
            mstrFilter = ",(select A.����ID,A.��ҳID from ������ҳ A  where A.��Ŀ���� is not null  and  A.���ʱ�� is  null ) Y where "
        Else
            mstrFilter = ",(select A.����ID,A.��ҳID from ������ҳ A  where A.���ʱ�� is  null ) Y where "
        End If
        
        Call FillData(strWhere)
        mclsPatient.AppendRows = True
        DataChanged = False
    '--------------------------------------------------------------------------------------------------------------
    Case "�������"
        
        DataChanged = False
        
    '--------------------------------------------------------------------------------------------------------------
    Case "У������"
        
        With vsfPatient
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("ѡ��")) = True Then
                    If .TextMatrix(i, .ColIndex("�����洢״̬")) = "��Ժ" Then
                        ExecuteCommand = True
                    Else
                        MsgBox "ѡ��Ĳ���:[" & .TextMatrix(i, .ColIndex("����")) & "]�Ѿ���[" & .TextMatrix(i, .ColIndex("���������")) & "]������,������ѡ��!", vbInformation, gstrSysName
                        ExecuteCommand = False
                        Exit Function
                    End If

                End If
            Next
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��������"
        
        
        Call DeleteRecordData(mrsPatient)
        With vsfPatient
            For intRow = 1 To .Rows - 1
                If Trim(.TextMatrix(intRow, .ColIndex("ID"))) <> "" And Trim(.TextMatrix(intRow, .ColIndex("ID"))) <> "0" And Abs(Val(.TextMatrix(intRow, .ColIndex("ѡ��")))) = 1 Then
                    mrsPatient.AddNew
                    mrsPatient("ID").Value = Trim(.TextMatrix(intRow, .ColIndex("ID")))
                    mrsPatient("����id").Value = Val(.TextMatrix(intRow, .ColIndex("����id")))
                    mrsPatient("��ҳid").Value = Val(.TextMatrix(intRow, .ColIndex("��ҳid")))
                    mrsPatient("����").Value = Trim(.TextMatrix(intRow, .ColIndex("����")))
                    mrsPatient("�Ա�").Value = Trim(.TextMatrix(intRow, .ColIndex("�Ա�")))
                    mrsPatient("����").Value = Trim(.TextMatrix(intRow, .ColIndex("����")))
                    mrsPatient("����״��").Value = Trim(.TextMatrix(intRow, .ColIndex("����״��")))
                    mrsPatient("��Ժʱ��").Value = Format(.TextMatrix(intRow, .ColIndex("��Ժʱ��")), "yyyy-MM-dd HH:mm:ss")
                    mrsPatient("��Ժʱ��").Value = Format(.TextMatrix(intRow, .ColIndex("��Ժʱ��")), "yyyy-MM-dd HH:mm:ss")
                    mrsPatient("��Ժ����").Value = Trim(.TextMatrix(intRow, .ColIndex("��Ժ����")))
                    
                    mrsPatient("סԺ��").Value = Trim(.TextMatrix(intRow, .ColIndex("סԺ��")))
                    mrsPatient("������").Value = Trim(.TextMatrix(intRow, .ColIndex("������")))
                    mrsPatient("סԺ����").Value = Trim(.TextMatrix(intRow, .ColIndex("סԺ����")))
                End If
            Next
        End With
    
    '------------------------------------------------------------------------------------------------------------------
    Case "��ע���"
        
        If Val(GetPara("ʹ�ø��Ի����")) = 1 Then
            'ʹ�ø��Ի�����
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "дע���"
        If Val(GetPara("ʹ�ø��Ի����")) = 1 Then
            'ʹ�ø��Ի�����

        End If
        
        strTmp = ""
        
        For intRow = 1 To lvw.ListItems.count
            strTmp = strTmp & "|" & lvw.ListItems(intRow).Text & "|" & lvw.ListItems(intRow).Tag
        Next
        If strTmp <> "" Then strTmp = Mid(strTmp, 2)
        
        strTmp = Replace(strTmp, "'", "123456789`1234567890")
        strTmp = Replace(strTmp, "123456789`1234567890", "''")
                        
        Call zlDatabase.SetPara("��������", strTmp, glngSys, mlngMoudal)
        
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
            .Cell(flexcpText, 1, .ColIndex("ѡ��"), .Rows - 1, .ColIndex("ѡ��")) = 1
            DataChanged = True
        End With
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_ClsAll
        
        With vsfPatient
            .Cell(flexcpText, 1, .ColIndex("ѡ��"), .Rows - 1, .ColIndex("ѡ��")) = 0
            DataChanged = True
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Refresh
        
        Call ExecuteCommand("��������")
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_SaveExit                  '��������
    
        If ExecuteCommand("У������") And DataChanged Then
            If ExecuteCommand("��������") Then
                
                DataChanged = False
                mblnOK = True
                
                Unload Me
                
            End If
        End If
                
    '------------------------------------------------------------------------------------------------------------------
    Case Else
    
        If Control.ID > 400 And Control.ID < 500 Then
            
        Else
             '��ҵ���޹صĹ��ܣ������Ĺ���
            Call CommandBarExecutePublic(Control, Me, vsfPatient, "���Ҳ��˽���嵥")
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
        Case 0  '��������ѡ��
            Set rsData = gclsPackage.GetDisease
            If ShowPubSelect(Me, txt(0), 3, "����,1200,0,;����,2700,0,;����,900,0,;����,900,0,", Me.Name & "\��������ѡ��", "����±���ѡ��һ������������Ŀ", rsData, rs, 8790, 4500, , Val(cmd(Index).Tag)) = 1 Then
                If Val(cmd(Index).Tag) <> zlCommFun.NVL(rs("ID").Value, 0) Then
                    txt(0).Text = zlCommFun.NVL(rs("����").Value)
                    cmd(Index).Tag = zlCommFun.NVL(rs("ID").Value, 0)
                    usrSaveItem.�������� = txt(0).Text
                    txt(0).Tag = ""
                End If
                DataChanged = True
            End If
        '------------------------------------------------------------------------------------------------------------------
        Case 1  '����
            
            
            
            Call cmdAdd_Click
            
            
        '------------------------------------------------------------------------------------------------------------------
        Case 2  'ɾ��
            
            lvwCombine.ListItems.Remove lvwCombine.SelectedItem.Index
        '    Call CreateSQL
            If lvwCombine.ListItems.count = 0 Then
                cmd(2).Enabled = False
            Else
                lvwCombine.SelectedItem.Selected = True
            End If
        '------------------------------------------------------------------------------------------------------------------
        Case 4          '��������
            
            strTmp = GetConditionString
            If strTmp <> "" Then
                
                Set objItem = lvw.ListItems.Add(, , "������", 1, 1)
                objItem.Tag = strTmp
                objItem.Selected = True
                
                lvw.SetFocus
                lvw.StartLabelEdit
            End If
        '------------------------------------------------------------------------------------------------------------------
        Case 5          '��������
            If Not (lvw.SelectedItem Is Nothing) Then
                lvw.SelectedItem.Tag = GetConditionString
            End If
        '------------------------------------------------------------------------------------------------------------------
        Case 6          'ɾ������
            
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
    Call ExecuteCommand("ˢ������")
End Sub

Private Sub cmdSearch_Click()
    Call ExecuteCommand("��������")
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
    
    Call ExecuteCommand("дע���")
    
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
            lvw��ҳ.Move tabMain.ClientLeft, tabMain.ClientTop, tabMain.ClientWidth, tabMain.ClientHeight
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
                usrSaveItem.�������� = ""
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
                    If ShowPubSelect(Me, txt(Index), 2, "����,1200,0,;����,2700,0,;����,900,0,;����,900,0,", Me.Name & "\�����������", "�������ѡ��һ��������Ŀ��Ŀ", rsData, rs) = 1 Then
                        If cmd(0).Tag <> zlCommFun.NVL(rs("ID").Value) Then
                            cmd(0).Tag = zlCommFun.NVL(rs("ID").Value)
                            txt(Index).Text = zlCommFun.NVL(rs("����").Value)
                            txt(Index).Tag = ""
                            ConditionChanged = True
        
                            usrSaveItem.�������� = txt(Index).Text
                        Else
                            txt(Index).Text = usrSaveItem.��������
                            txt(Index).Tag = ""
                        End If
                    Else
                        txt(Index).Text = usrSaveItem.��������
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
                txt(Index).Text = usrSaveItem.��������
                txt(Index).Tag = ""
            End If
    End Select

End Sub

Private Sub vsfPatient_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '�༭����
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
    '�༭����
    Call mclsPatient.DbClick
End Sub

Private Sub vsfPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    '�༭����
    Call mclsPatient.KeyDown(KeyCode, Shift)
End Sub

Private Sub vsfPatient_KeyPress(KeyAscii As Integer)
    'ToDo...
    If KeyAscii = vbKeyReturn Then Call vsfPatient_DblClick
    
    '�༭����,������
    Call mclsPatient.KeyPress(KeyAscii)
End Sub

Private Sub vsfPatient_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    '�༭����
    Call mclsPatient.KeyPressEdit(KeyAscii)
End Sub

Private Sub vsfPatient_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    '�༭����
    Call mclsPatient.EditSelAll
End Sub

Private Sub vsfPatient_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '�༭����
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
        Case vsfPatient.ColIndex("ѡ��")
            Cancel = False
            Exit Sub
        Case Else
            Cancel = True
            Exit Sub
    End Select
End Sub

Private Sub vsfPatient_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    Select Case Col
        Case vsfPatient.ColIndex("���"), vsfPatient.ColIndex("ѡ��")
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
            lngID = Val(vsfPatient.TextMatrix(vsfPatient.Row, vsfPatient.ColIndex("����ID")))
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
    
    strHead = "���,500,4,1;ѡ��,500,4,1;סԺ��,1500,1,1;����,900,1,0;�Ա�,500,4,0;����,500,7,0;��������,1100,1,0;סԺ����,900,7,1;��ͥ��ַ,1350,1,0;����״��,900,1,0;��ͥ�绰,1000,1,0;" & _
              "������λ,1200,1,0;��λ�绰,1000,1,0;��ϵ������,1000,1,0;��ϵ�˹�ϵ,1200,1,0;��ϵ�˵�ַ,1200,1,0;��ϵ�˵绰,1000,1,0;������,1000,1,1;�����ż���,1200,1,1;�����洢״̬,1440,1,0;" & _
              "�������λ��,1440,1,0;��Ժʱ��,1670,1,0;��Ժʱ��,1670,1,0;����ID,0,7,-1;��ҳID,0,7,-1;ID,0,7,-1;��Ժ����,1100,7,-1;���״̬,0,7,-1;���������,0,7,-1"
'              mclsPatient.LoadStateFromString strHead
    Call SetVsFlexGridChangeHead(strHead, vsfPatient, 1)
End Sub

'����24813
Private Sub SetInitVfgDataFormat(ByVal vsGrid As VSFlexGrid)
    Dim i As Long
    With vsGrid
        .ColDataType(.ColIndex("ѡ��")) = flexDTBoolean
        .ForeColorSel = .CellForeColor
        .ExplorerBar = flexExSortShowAndMove
        .SelectionMode = flexSelectionByRow
        .Editable = flexEDKbdMouse
    End With
End Sub

'����24813
Private Function FillData(ByVal strWhere As String) As Boolean
'װ����������Ĳ�����VSFlexGrid��
    Dim rs���� As New ADODB.Recordset
    Dim lngCol As Long, varValue As Variant
    Dim lst As ListItem, bln���� As Boolean
    Dim strKey As String, strTemp As String
    Dim strRecord As String
    Dim i As Long
     
    zlCommFun.ShowFlash "����װ�벡����¼�����Ե� ����"

    gstrSQL = "" & _
         "   Select  Trim(To_Char(max(Y1.����id)))||'-'||Trim(To_Char(max(Y1.��ҳid))) As ID,X.����ID,max(Y1.��ҳid) as ��ҳID,max(Y1.סԺ��) as סԺ��,X.����,X.�Ա�,X.����,to_char(X.��������,'YYYY-MM-DD HH24:MI') as ��������,Zl_��ȡסԺ��������ҳid(X.����id,max(y.��ҳid),0) as סԺ����," & _
         "           X.�����ص�,X.���֤��,X.ְҵ,X.����״��,X.��ͥ��ַ,max(Z.������) as  �����ż���," & _
         "           X.��ͥ�绰,X.��ϵ������,X.��ϵ�˹�ϵ,X.��ϵ�˵�ַ,X.��ϵ�˵绰,X.������λ,X.��λ�绰,max(Z.������) as ������ ,Decode(max(D.����ID),'','��Ժ','���') as �����洢״̬,max(Z.���λ��) as �������λ��," & _
         "           max(to_char(y1.��Ժ����,'YYYY-MM-DD HH24:MI')) As ��Ժʱ�� ,to_char(y1.��Ժ����,'YYYY-MM-DD HH24:MI') as ��Ժʱ��,max(C.����) As ��Ժ����,max(D.����ID) as ���״̬,max(D.������) As ��������� " & _
         "   From ������Ϣ X ,���ű� C,(Select Max(A1.������) as ������,Max(B1.����ID) as ����ID,Max(B1.��ҳID) as ��ҳID From �������ļ�¼ A1 ,������������ B1,����������Ա C1 Where A1.ID = B1.����ID And A1.ID = C1.����ID And A1.��¼״̬=2 Group By B1.����ID,B1.��ҳID,A1.������) D,סԺ������¼ Z,������ҳ Y1  " & _
                  mstrFilter & " Y.����ID=X.����ID and X.����ID=Y1.����ID  And y1.��ҳID=y.��ҳid  And C.ID = Y1.��Ժ����ID  And D.����ID(+) = X.����ID And  Y1.��Ժ���� is Not null  and X.����ID=Z.����ID(+) and Y1.����״̬=5 " & strWhere & _
         "Group By  X.����id,Y1.��ҳID,X.סԺ��, X.����, X.�Ա�, X.����,x.��������,X.�����ص�, X.���֤��, X.ְҵ, X.����״��, X.��ͥ��ַ, X.��ͥ�绰, X.��ϵ������, X.��ϵ�˹�ϵ, X.��ϵ�˵�ַ, X.��ϵ�˵绰, X.������λ," & _
         "    X.��λ�绰 ,X.��Ժʱ��, y1.��Ժ���� " & _
         "   Order by  to_char(y1.��Ժ����,'YYYY-MM-DD HH24:MI') Desc"
      
    On Error GoTo errHandle
    Call zlDatabase.OpenRecordset(rs����, gstrSQL, Me.Caption)
    
    With vsfPatient
        Call initVfgPatient
        .Rows = IIf(rs����.EOF, 0, rs����.RecordCount) + 1
        
        If Not rs����.EOF Then
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("���")) = i
                .TextMatrix(i, .ColIndex("ѡ��")) = 0
                .TextMatrix(i, .ColIndex("סԺ��")) = IIf(IsNull(rs����!סԺ��), "", rs����!סԺ��)
                .TextMatrix(i, .ColIndex("����")) = IIf(IsNull(rs����!����), "", rs����!����)
                .TextMatrix(i, .ColIndex("�Ա�")) = IIf(IsNull(rs����!�Ա�), "", rs����!�Ա�)
                .TextMatrix(i, .ColIndex("����")) = IIf(IsNull(rs����!����), 0, rs����!����)
                .TextMatrix(i, .ColIndex("��������")) = IIf(IsNull(rs����!��������), "", rs����!��������)
                .TextMatrix(i, .ColIndex("סԺ����")) = IIf(IsNull(rs����!סԺ����), "", rs����!סԺ����)
                .TextMatrix(i, .ColIndex("��ͥ��ַ")) = IIf(IsNull(rs����!��ͥ��ַ), "", rs����!��ͥ��ַ)
                .TextMatrix(i, .ColIndex("����״��")) = IIf(IsNull(rs����!����״��), "", rs����!����״��)
                .TextMatrix(i, .ColIndex("��ͥ�绰")) = IIf(IsNull(rs����!��ͥ�绰), "", rs����!��ͥ�绰)
                .TextMatrix(i, .ColIndex("������λ")) = IIf(IsNull(rs����!������λ), "", rs����!������λ)
                .TextMatrix(i, .ColIndex("��λ�绰")) = IIf(IsNull(rs����!��λ�绰), "", rs����!��λ�绰)
                .TextMatrix(i, .ColIndex("��ϵ������")) = IIf(IsNull(rs����!��ϵ������), "", rs����!��ϵ������)
                .TextMatrix(i, .ColIndex("��ϵ�˹�ϵ")) = IIf(IsNull(rs����!��ϵ�˹�ϵ), "", rs����!��ϵ�˹�ϵ)
                .TextMatrix(i, .ColIndex("��ϵ�˵�ַ")) = IIf(IsNull(rs����!��ϵ�˵�ַ), "", rs����!��ϵ�˵�ַ)
                .TextMatrix(i, .ColIndex("��ϵ�˵绰")) = IIf(IsNull(rs����!��ϵ�˵绰), "", rs����!��ϵ�˵绰)
                .TextMatrix(i, .ColIndex("������")) = IIf(IsNull(rs����!������), "", rs����!������)
                .TextMatrix(i, .ColIndex("�����ż���")) = IIf(IsNull(rs����!�����ż���), "", rs����!�����ż���)
'                    If gSystemPara.bln���������� = True And mnuViewRecord.Checked = False Then
'                        '�����������Ų�����ʾ����ʱ��ò����ż���
'                        varValue = Get�����ż���(CLng(rs����("����id").Value))
'                        .TextMatrix(i, .ColIndex("�����ż���")) = IIf(IsNull(varValue), "", varValue)
'                    ElseIf gSystemPara.bln���������� = False And mnuViewRecord.Checked = False Then
'                        '��������������Ų�����ʾ�����б�ʱ
'                        varValue = rs����("������").Value
'                        .TextMatrix(i, .ColIndex("�����ż���")) = IIf(IsNull(varValue), "", varValue)
'                    ElseIf gSystemPara.bln���������� = True And mnuViewRecord.Checked = True Then
'                        '�������ŵ�����Ų�����ʾ�����б�ʱ
'                        varValue = rs����("������")
'                        .TextMatrix(i, .ColIndex("�����ż���")) = IIf(IsNull(varValue), "", varValue)
'                    ElseIf gSystemPara.bln���������� = False And mnuViewRecord.Checked = True Then
'                        '���������ŵ�����Ų�����ʾ�����б�ʱ
'                        varValue = rs����("������").Value
'                        .TextMatrix(i, .ColIndex("�����ż���")) = IIf(IsNull(varValue), "", varValue)
'                    End If
                .TextMatrix(i, .ColIndex("�����洢״̬")) = IIf(IsNull(rs����!�����洢״̬), "", rs����!�����洢״̬)
                .TextMatrix(i, .ColIndex("�������λ��")) = IIf(IsNull(rs����!�������λ��), "", rs����!�������λ��)
                .TextMatrix(i, .ColIndex("��Ժʱ��")) = IIf(IsNull(rs����!��Ժʱ��), "", rs����!��Ժʱ��)
                .TextMatrix(i, .ColIndex("��Ժʱ��")) = IIf(IsNull(rs����!��Ժʱ��), "", rs����!��Ժʱ��)
                .TextMatrix(i, .ColIndex("��Ժ����")) = IIf(IsNull(rs����!��Ժ����), "", rs����!��Ժ����)
                .TextMatrix(i, .ColIndex("����ID")) = IIf(IsNull(rs����!����ID), 0, rs����!����ID)
                .TextMatrix(i, .ColIndex("��ҳID")) = IIf(IsNull(rs����!��ҳID), 0, rs����!��ҳID)
                
                .TextMatrix(i, .ColIndex("���״̬")) = IIf(IsNull(rs����!���״̬), 0, rs����!���״̬)
                .TextMatrix(i, .ColIndex("���������")) = IIf(IsNull(rs����!���������), 0, rs����!���������)
                
'                strRecord = IIf(IsNull(rs����!�����洢״̬), "", rs����!�����洢״̬)
'                Select Case strRecord
'                Case "��Ժ"
'                    .Cell(flexcpPicture, i, .ColIndex("סԺ��")) = imgCelNo(0)
'                Case "���ֽ��"
'                    .Cell(flexcpPicture, i, .ColIndex("סԺ��")) = imgCelNo(1)
'                Case "���"
'                    .Cell(flexcpPicture, i, .ColIndex("סԺ��")) = imgCelNo(2)
'                End Select

                strRecord = IIf(IsNull(rs����!���״̬), "", rs����!���״̬)
                If strRecord = "" Then
                    .Cell(flexcpPicture, i, .ColIndex("סԺ��")) = imgCelNo(0)
                Else
                    .Cell(flexcpPicture, i, .ColIndex("סԺ��")) = imgCelNo(2)
                End If
                .TextMatrix(i, .ColIndex("ID")) = rs����!ID
                
                
                rs����.MoveNext
            Next
        End If
    End With
    Call SetInitVfgDataFormat(vsfPatient)
    Call RestoreHead(vsfPatient, 1)

    If mlngOldRow > 0 And (vsfPatient.Rows - 1) >= mlngOldRow Then vsfPatient.Select mlngOldRow, 1
    '��ʾ������Ϣ
    Call FillTabData
  
    Call zlCommFun.StopFlash
    rs����.Close
    FillData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call zlCommFun.StopFlash
    rs����.Close
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
            .Add , "T0", "����ҳ����"
            .Item("T0").Tag = 0
        Else
            lngID = Val(vsfPatient.TextMatrix(vsfPatient.Row, vsfPatient.ColIndex("����ID")))
            lngPageID = Val(vsfPatient.TextMatrix(vsfPatient.Row, vsfPatient.ColIndex("��ҳID")))
            tabMain.Tag = lngID
            gstrSQL = "" & _
                "   select ��ҳID as סԺ����,Zl_��ȡסԺ��������ҳid(����id,��ҳid,0) as ʵ��סԺ���� from ������ҳ where ��Ŀ���� is not null and ����ID=[1] Order by ��ҳID Desc"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngID)
            Do Until rsTemp.EOF
                i = rsTemp("סԺ����")
                .Add , "T" & i, "�� " & NVL(rsTemp!ʵ��סԺ����) & " ��סԺ" '����ĳ��Ϊ���۲��ˣ��м���ܳ����ж�
                .Item("T" & i).Tag = i
                If i = lngPageID Then
                    .Item("T" & i).Selected = True
                End If
                rsTemp.MoveNext
            Loop
              
            If .count = 0 Then
                .Add , "T0", "����ҳ����"
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
'װ��ָ��������������ҳ��ListView��
    Dim rs��ҳ As New ADODB.Recordset
    Dim lngCount As Long
    Dim fld As Field
    Dim lst As ListItem
    Dim strҽ���� As String
    On Error GoTo errHandle
    
    lvw��ҳ.ListItems.Clear
    rs��ҳ.CursorLocation = adUseClient
'    gstrSQL = "" & _
'        " Select A.������ as ������1,F.������ as ������,A.סԺ��,C.���� as ��Ժ����,to_char(A.��Ժ����,'YYYY-MM-DD HH24:MI') AS ��Ժ����,A.��Ժ���� as ��Ժ����" & _
'        "           ,D.���� as ��Ժ����,to_char(A.��Ժ����,'YYYY-MM-DD HH24:MI') AS ��Ժ����" & _
'        "           ,A.סԺ����,A.���ú�,B.������� as ��Ҫ�����,E.���� as ��Ҫ��ϱ���,B.��Ժ��� as ��Ҫ��ϳ�Ժ��� " & _
'        "           ,decode(A.�����־,1,'��',2,'��',3,'��','') as �Ƿ�����,A.��ĿԱ���� as ��ĿԱ,to_char(A.��Ŀ����,'YYYY-MM-DD') as ��Ŀ����" & _
'        " From ������ҳ A,(Select * From ������ϼ�¼ where ��¼��Դ=4 and ������� in (3,13) and ����ID=[1] and ��ҳID=[2]) B,���ű� C,���ű� D,��������Ŀ¼ E,סԺ������¼ F" & _
'        " Where A.��Ŀ���� is not null and A.����ID=B.����ID(+)  and A.��ҳID=B.��ҳID(+) and " & _
'        "       A.��Ժ����ID=C.ID(+) and A.��Ժ����ID=D.ID and B.����ID=E.ID(+) and A.����ID=[1] and A.��ҳID=[2] and A.����id=F.����ID(+) "
    
    gstrSQL = "" & _
    " Select     F.������ as ������,A.סԺ��,C.���� as ��Ժ����,to_char(A.��Ժ����,'YYYY-MM-DD HH24:MI') AS ��Ժ����,A.��Ժ���� as ��Ժ����" & _
    "           ,D.���� as ��Ժ����,to_char(A.��Ժ����,'YYYY-MM-DD HH24:MI') AS ��Ժ����" & _
    "           ,A.סԺ����,A.���ú�,B.������� as ��Ҫ�����,E.���� as ��Ҫ��ϱ���,B.��Ժ��� as ��Ҫ��ϳ�Ժ��� " & _
    "           ,decode(A.�����־,1,'��',2,'��',3,'��','') as �Ƿ�����,A.��ĿԱ���� as ��ĿԱ,to_char(A.��Ŀ����,'YYYY-MM-DD') as ��Ŀ����" & _
    " From ������ҳ A,(Select * From ������ϼ�¼ where ��¼��Դ=4 and ������� in (3,13) and ����ID=[1] and ��ҳID=[2]) B,���ű� C,���ű� D,��������Ŀ¼ E,סԺ������¼ F" & _
    " Where A.��Ŀ���� is not null and A.����ID=B.����ID(+)  and A.��ҳID=B.��ҳID(+) and " & _
    "       A.��Ժ����ID=C.ID(+) and A.��Ժ����ID=D.ID and B.����ID=E.ID(+) and A.����ID=[1] and A.��ҳID=[2] and A.����id=F.����ID(+) and a.��ҳid=F.��ҳID(+) "
        
    Set rs��ҳ = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(tabMain.Tag), Val(Mid(tabMain.SelectedItem.Key, 2)))
    
    If rs��ҳ.EOF = False Then
        For lngCount = 0 To rs��ҳ.Fields.count - 1
            Set fld = rs��ҳ.Fields(lngCount)
             If fld.Name = "������" Then
                Set lst = lvw��ҳ.ListItems.Add(, "������", "������", "Attrib", "Attrib")
                lst.SubItems(1) = IIf(IsNull(fld.Value), "", fld.Value)
'            'ElseIf fld.Name = "������1" And IsNull(fld.Value) = True And gSystemPara.bln���������� = True Then
'                Set lst = lvw��ҳ.ListItems.Add(, "������", "������", "Attrib", "Attrib")
'                lst.SubItems(1) = IIf(IsNull(rs��ҳ("������")), "", rs��ҳ("������"))
'            'ElseIf gSystemPara.bln���������� = False And fld.Name = "������" Then
'                Set lst = lvw��ҳ.ListItems.Add(, "������", "������", "Attrib", "Attrib")
'                lst.SubItems(1) = IIf(IsNull(rs��ҳ("������")), "", rs��ҳ("������"))
            ElseIf fld.Name <> "������" Then
                Set lst = lvw��ҳ.ListItems.Add(, fld.Name, fld.Name, "Attrib", "Attrib")
                lst.SubItems(1) = IIf(IsNull(fld.Value), "", fld.Value)
           End If
            '�����ֶε���ʾЧ��
            Select Case fld.Name
                Case "��Ҫ��ϱ���", "��Ҫ�����", "��Ҫ��ϳ�Ժ���"
                    lst.ForeColor = RGB(0, 0, 213)
                    lst.ListSubItems(1).ForeColor = RGB(0, 0, 213)
            End Select
        Next
    End If
    'ȷ��ҽ����
    gstrSQL = "Select ��Ϣֵ From ������ҳ�ӱ� where ��Ϣ��='ҽ����' and ����ID=[1] and ��ҳID=[2]"
    Set rs��ҳ = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(tabMain.Tag), Val(Mid(tabMain.SelectedItem.Key, 2)))
    If Not rs��ҳ.EOF Then
        strҽ���� = zlCommFun.NVL(rs��ҳ!��Ϣֵ)
    Else
        strҽ���� = ""
    End If
    Set lst = lvw��ҳ.ListItems.Add(, "ҽ����", "ҽ����", "Attrib", "Attrib")
    lst.SubItems(1) = strҽ����
    rs��ҳ.Close
    
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
        strHeadInfo = "������ͷ��Ϣ"
    Else
        strHeadInfo = "������ͷ��Ϣ"
    End If
    zl_VsGrid_SaveToPara vsGrid, Me.Caption, mlngMoudal, strHeadInfo, True, True
End Sub

Private Sub RestoreHead(ByVal vsGrid As VSFlexGrid, intListOrDetail As Integer)
    Dim strHeadInfo As String
    If intListOrDetail = 1 Then
        strHeadInfo = "������ͷ��Ϣ"
    Else
        strHeadInfo = "������ͷ��Ϣ"
    End If
    zl_VsGrid_FromParaRestore vsGrid, Me.Caption, mlngMoudal, strHeadInfo, True, True
End Sub

Private Sub tabMain_Click()
    Call FillDetail
End Sub


Public Sub SetMenu()
    Dim i As Long
    Dim int��Ժ As Long, int��� As Long
    Dim blnCount  As Boolean, blnEnble As Boolean
   
    '����24813
    Dim lngCount As Long
    
    blnEnble = True
    If vsfPatient.Rows = 2 Then
        If vsfPatient.TextMatrix(1, vsfPatient.ColIndex("ѡ��")) = "" Then
            stbThis.Panels(2) = ""
            Exit Sub
        End If
    End If
    
    stbThis.Panels(2) = "��ǰ��ʾ��" & vsfPatient.Rows - 1 & "�ݲ�������ѯ�����ǡ�" & mstr��ʾ & "��"
    With vsfPatient
        lngCount = .Rows - 1
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("ѡ��")) = True Then
                If .TextMatrix(i, .ColIndex("�����洢״̬")) = "��Ժ" Then
                    int��Ժ = int��Ժ + 1
                Else
                    int��� = int��� + 1
                End If
                If int��Ժ > 0 And int��� > 0 Then Exit For
            End If
        Next
        If lngCount > 0 Then
            blnCount = .TextMatrix(.Row, .ColIndex("�����洢״̬")) = "��Ժ"
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
    mstr���� = ""
    strTemp = ""
    strFilterX = ""
    If lvwCombine.ListItems.count = 0 Then
        If mbln����ϵͳ Then
            mstrFilter = " A.��Ŀ���� is not null and  A.���ʱ�� is  null"
        Else
             mstrFilter = " A.���ʱ�� is null "
        End If
        mstr���� = "���в���"
    Else
        strTemp = IIf(opt(0).Value = True, " AND ", " OR ")
        For lngTemp = 1 To lvwCombine.ListItems.count
            If InStr(lvwCombine.ListItems(lngTemp).Tag, "X.") = 0 Then
                mstrFilter = mstrFilter & "(" & lvwCombine.ListItems(lngTemp).Tag & ")" & strTemp
            Else
                strFilterX = strFilterX & "(" & lvwCombine.ListItems(lngTemp).Tag & ")" & strTemp
            End If
            mstr���� = mstr���� & "(" & lvwCombine.ListItems(lngTemp) & " " & lvwCombine.ListItems(lngTemp).SubItems(1) & ")" & IIf(opt(0).Value = True, " ���� ", " ���� ")
        Next
        If Len(mstrFilter) > 4 Then mstrFilter = Left(mstrFilter, Len(mstrFilter) - 4)
        If Len(strFilterX) > 4 Then strFilterX = Left(strFilterX, Len(strFilterX) - 4)
        mstr���� = Left(mstr����, Len(mstr����) - 4)
        '�õ���֮��Ĺ�ϵ
        strTemp = ""
        'Ҫ��ѯ�ı������ñ�֮�����ϵ
        If InStr(mstrFilter, "��J.��") > 0 Then strTemp = strTemp & "��C.������ID=��J.��ID AND "
        If InStr(mstrFilter, "��I.��") > 0 Then strTemp = strTemp & "��B.����������ID=��I.��ID AND "
        If InStr(mstrFilter, "��G.��") > 0 Then strTemp = strTemp & "��A.����Ժ����ID=��G.��ID AND "
        If InStr(mstrFilter, "��E.��") > 0 Then strTemp = strTemp & "��A.����Ժ����ID=��E.��ID AND "
        '�����벡����ҳ��֮�����ϵ
        If InStr(mstrFilter & strTemp, "��B.��") > 0 Then strTemp = strTemp & "��A.������ID=��B.������ID AND ��A.����ҳID=��B.����ҳID AND "
        If InStr(mstrFilter & strTemp, "��C.��") > 0 Then strTemp = strTemp & "��A.������ID=��C.������ID AND ��A.����ҳID=��C.����ҳID AND "
        If InStr(mstrFilter & strTemp, "��O.��") > 0 Then strTemp = strTemp & "��A.������ID=��O.������ID AND ��A.����ҳID=��O.����ҳID AND "
        If InStr(mstrFilter & strTemp, "��L1.��") > 0 Then strTemp = strTemp & "��A.������ID=��L1.������ID AND ��A.����ҳID=��L1.����ҳID AND "
        If InStr(mstrFilter & strTemp, "��L2.��") > 0 Then strTemp = strTemp & "��A.������ID=��L2.������ID AND ��A.����ҳID=��L2.����ҳID AND "
        If InStr(mstrFilter & strTemp, "��L3.��") > 0 Then strTemp = strTemp & "��A.������ID=��L3.������ID AND ��A.����ҳID=��L3.����ҳID AND "
        If InStr(mstrFilter & strTemp, "��L4.��") > 0 Then strTemp = strTemp & "��A.������ID=��L4.������ID AND ��A.����ҳID=��L4.����ҳID AND "
        If InStr(mstrFilter & strTemp, "��L5.��") > 0 Then strTemp = strTemp & "��A.������ID=��L5.������ID AND ��A.����ҳID=��L5.����ҳID AND "
        If InStr(mstrFilter & strTemp, "��L6.��") > 0 Then strTemp = strTemp & "��A.������ID=��L6.������ID AND ��A.����ҳID=��L6.����ҳID AND "
        If InStr(mstrFilter & strTemp, "��L7.��") > 0 Then strTemp = strTemp & "��A.������ID=��L7.������ID AND ��A.����ҳID=��L7.����ҳID AND "
        If InStr(mstrFilter & strTemp, "��L8.��") > 0 Then strTemp = strTemp & "��A.������ID=��L8.������ID AND ��A.����ҳID=��L8.����ҳID AND "
        For lngCount = mlngIndex + 1 To lstFields.ListCount - 1
            If InStr(mstrFilter & strTemp, "��L" & lngCount & ".��") > 0 Then
                strTemp = strTemp & "��A.������ID=��L" & lngCount & ".������ID AND ��A.����ҳID=��L" & lngCount & ".����ҳID AND "
            End If
        Next lngCount
        If InStr(mstrFilter & strTemp, "��M.��") > 0 Then strTemp = strTemp & "��A.������ID=��M.������ID AND ��A.����ҳID=��M.����ҳID  AND "
        If InStr(mstrFilter & strTemp, "��N.��") > 0 Then strTemp = strTemp & "��A.������ID=��N.������ID AND "
        If InStr(mstrFilter & strTemp, "��P.��") > 0 Then strTemp = strTemp & "��A.������ID=��P.������ID AND "
        If InStr(mstrFilter & strTemp, "��P1.��") > 0 Then strTemp = strTemp & "��A.������ID=��P1.������ID AND "
        mstrFilter = strTemp & "(" & mstrFilter & ")"
        
        
        '�ó�Ҫ���õı�
        strTemp = ""
        If InStr(mstrFilter, "��P.��") > 0 Then strTemp = ",סԺ������¼ P" & strTemp
        If InStr(mstrFilter, "��P1.��") > 0 Then strTemp = ",������Ϣ P1" & strTemp
        If InStr(mstrFilter, "��N.��") > 0 Then strTemp = ",���˹���ҩ�� N" & strTemp
        If InStr(mstrFilter, "��M.��") > 0 Then strTemp = ",��Ϸ������ M" & strTemp
        If InStr(mstrFilter, "��L1.��") > 0 Then strTemp = ",������ҳ�ӱ� L1" & strTemp
        If InStr(mstrFilter, "��L2.��") > 0 Then strTemp = ",������ҳ�ӱ� L2" & strTemp
        If InStr(mstrFilter, "��L3.��") > 0 Then strTemp = ",������ҳ�ӱ� L3" & strTemp
        If InStr(mstrFilter, "��L4.��") > 0 Then strTemp = ",������ҳ�ӱ� L4" & strTemp
        If InStr(mstrFilter, "��L5.��") > 0 Then strTemp = ",������ҳ�ӱ� L5" & strTemp
        If InStr(mstrFilter, "��L6.��") > 0 Then strTemp = ",������ҳ�ӱ� L6" & strTemp
        If InStr(mstrFilter, "��L7.��") > 0 Then strTemp = ",������ҳ�ӱ� L7" & strTemp
        If InStr(mstrFilter, "��L8.��") > 0 Then strTemp = ",������ҳ�ӱ� L8" & strTemp
        For lngCount = mlngIndex + 1 To lstFields.ListCount - 1
            If InStr(mstrFilter & strTemp, "��L" & lngCount & ".��") > 0 Then
                 strTemp = ",������ҳ�ӱ� L" & lngCount & strTemp
            End If
        Next lngCount
        If InStr(mstrFilter, "��J.��") > 0 Then strTemp = ",��������Ŀ¼ J" & strTemp
        If InStr(mstrFilter, "��I.��") > 0 Then strTemp = ",��������Ŀ¼ I" & strTemp
        If InStr(mstrFilter, "��G.��") > 0 Then strTemp = ",���ű� G" & strTemp
        If InStr(mstrFilter, "��E.��") > 0 Then strTemp = ",���ű� E" & strTemp
        If InStr(mstrFilter, "��C.��") > 0 Then strTemp = ",������ϼ�¼ C" & strTemp
        If InStr(mstrFilter, "��B.��") > 0 Then strTemp = ",���������¼ B" & strTemp
        If InStr(mstrFilter, "��O.��") > 0 Then strTemp = ",���ļ�¼ O" & strTemp
    End If
    strTemp = "������ҳ A " & strTemp
    
    '����31176 by lesfeng 2010-07-19 IIf(InStr(1, strTemp, "���������¼") > 0, " And B.��¼��Դ = 4 ", "")
    mstrReturn = ",(select A.����ID,A.��ҳID from " & strTemp & " where A.��Ŀ���� is not null " & IIf(InStr(1, strTemp, "���ļ�¼") > 0, " And O.�黹ʱ�� Is Null ", "") & IIf(InStr(1, strTemp, "���������¼") > 0, " And B.��¼��Դ = 4 ", "") & IIf(InStr(1, strTemp, "������ϼ�¼") > 0, " And C.��¼��Դ=4 ", "") & IIf(mstrFilter = "()", "", " and " & mstrFilter) & ") Y where " & IIf(Len(strFilterX) = 0, "", "(" & strFilterX & ") and ")
    mstrSQL = "select X.סԺ�� as סԺ��, A.����ID,X.����,A.����,A.����״��,A.���� from ������Ϣ X," & strTemp & " where X.����ID=A.����ID and A.��Ŀ���� is not null " & IIf(InStr(1, strTemp, "���ļ�¼") > 0, " And O.�黹ʱ�� Is Null ", "") & IIf(InStr(1, strTemp, "������ϼ�¼") > 0, " And C.��¼��Դ=4 ", "") & IIf(mstrFilter = "()", "", " and " & mstrFilter) & IIf(Len(strFilterX) = 0, "", " and (" & strFilterX & ")  ")
    mstrExecel = "select A.����ID,A.��ҳID from ������Ϣ X," & strTemp & " where X.����ID=A.����ID and A.��Ŀ���� is not null " & IIf(InStr(1, strTemp, "���ļ�¼") > 0, " And O.�黹ʱ�� Is Null ", "") & IIf(InStr(1, strTemp, "������ϼ�¼") > 0, " And C.��¼��Դ=4 ", "") & IIf(mstrFilter = "()", "", " and " & mstrFilter) & IIf(Len(strFilterX) = 0, "", " and (" & strFilterX & ")  ")
    
    mstrReturn = Replace(mstrReturn, "��", "")
    mstrReturn = Replace(mstrReturn, "��", "")
    mstrSQL = Replace(mstrSQL, "��", "")
    mstrSQL = Replace(mstrSQL, "��", "")
    
    mstrExecel = Replace(mstrExecel, "��", "")
    mstrExecel = Replace(mstrExecel, "��", "")
    
'    mblnChange = False
    CreateSQL = True
End Function

Private Function ValidSQL() As Boolean
    Dim bln������� As Boolean, bln�������� As Boolean
    Dim lst As ListItem
    
    For Each lst In lvwCombine.ListItems
        If lst.Text = "�������" Then
            bln������� = True
        Else
            If lst.Text = "��������" Then
                bln�������� = True
            End If
        End If
    Next
    
    If bln�������� Xor bln������� Then
        MsgBox "��������������������Ҫͬʱ���ڡ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    ValidSQL = True
End Function

Private Sub cmdAdd_Click()
    Dim bln���� As Boolean
    Dim lst As ListItem
    
    Dim str������ As String
    Dim strField As String
    Dim strTemp As String
    
    '�ó������ݿ��б�ʾ��ֵ
    If ValidateDate = False Then Exit Sub

    str������ = lstFields.List(lstFields.ListIndex)
    
    '�����ǰ�Ƿ��Ѿ��й���������
    For Each lst In lvwCombine.ListItems
        If lst.Text = str������ Then
            bln���� = True
            Exit For
        End If
    Next
    
    strTemp = LeftB(cmbOperate.Text, 10) & IIf(cmbList.Visible = False, IIf(Trim(cmbExample.Text) = "", "��", cmbExample.Text), cmbList.Text)
    If bln���� = False Then
        '����һ��������
        Set lst = lvwCombine.ListItems.Add(, , str������)
        lst.SubItems(1) = strTemp
    Else
        If lst.ListSubItems(1).Tag = "" Then
            'Ҳֻ��������һ��
            If cmbLogical.Text = "����" Then
                lst.SubItems(1) = "(" & lst.SubItems(1) & ") ���� (" & strTemp & ")"
            Else
                lst.SubItems(1) = "(" & lst.SubItems(1) & ") ���� (" & strTemp & ")"
            End If
            lst.ListSubItems(1).Tag = "���"
        Else
            If cmbLogical.Text = "����" Then
                lst.SubItems(1) = lst.SubItems(1) & " ���� (" & strTemp & ")"
            Else
                lst.SubItems(1) = lst.SubItems(1) & " ���� (" & strTemp & ")"
            End If
        End If
    End If
    
    'Ϊ�˰ѱ�������ͨ�ַ������ֿ���ʹ�����±�ʾ���ڲ���SQL���ʱȥ��
    Select Case str������
        Case "����", "�Ա�", "���֤��", "��������"
            strField = "��X.��" & str������
        Case "��������"
            strField = "zlspellcode(��X.������)"
        Case "סԺ����"
            strField = "��A.����ҳID"
        Case "סԺ��"
            strField = "��A.��סԺ��"
        Case "������"
            strField = "��P.��������"
        Case "������"
            strField = "��P.��������"
        Case "��Ժ����"
            strField = "��E.������"
        Case "��Ժ����"
            strField = "��G.������"
        Case "������������", "�����п�", "��������", "������������"
            strField = "��B.��" & Mid(str������, 3)
        Case "��������", "����ҽʦ", "����ҽʦ"
            strField = "��B.��" & str������
        Case "��������"
            strField = "��I.������"
        Case "��������"
            If mint���뷽ʽ = 0 Then
                strField = "��I.������"
            Else
                strField = "��I.�������"
            End If
        Case "��ϱ���"
            strField = "��J.������"
        Case "��ϼ���"
            If mint���뷽ʽ = 0 Then
                strField = "��J.������"
            Else
                strField = "��J.�������"
            End If
        Case "���������Ϣ", "��ϳ�Ժ���", "��ϱ������"
            If str������ = "���������Ϣ" Then
                strField = "��C.���������"
            Else
                strField = "��C.��" & Mid(str������, 3)
            End If
            
        Case "��ϴ���"
            strField = "��C.����ϴ���"
        Case "�������"
            strField = "��C.���������"
        Case "δ��"
            strField = "��C.���Ƿ�δ��"
        Case "����"
            strField = "��C.���Ƿ�����"
        Case "��������"
            strField = "��M.����������"
        Case "�������"
            strField = "��M.���������"
        Case "��������"
            strField = "��L1.����Ϣ��='��������' AND ��L1.����Ϣֵ"
        Case "������"
            strField = "��L2.����Ϣ��='������' AND ��L2.����Ϣֵ"
        Case "����ҽʦ"
            strField = "��L3.����Ϣ��='����ҽʦ' AND ��L3.����Ϣֵ"
        Case "����ҽʦ"
            strField = "��L4.����Ϣ��='����ҽʦ' AND ��L4.����Ϣֵ"
        Case "ҽ����"
            strField = "��L3.����Ϣ��='ҽ����' AND ��L3.����Ϣֵ"
        Case "���ϸ��"
            strField = "��L5.����Ϣ��='���ϸ��' AND ��L5.����Ϣֵ"
        Case "��ѪС��"
            strField = "��L6.����Ϣ��='��ѪС��' AND ��L6.����Ϣֵ"
        Case "��Ѫ��"
            strField = "��L7.����Ϣ��='��Ѫ��' AND ��L7.����Ϣֵ"
        Case "��ȫѪ"
             strField = "��L8.����Ϣ��='��ȫѪ' AND ��L8.����Ϣֵ"
        Case "����ҩ��"
            strField = "��N.������ҩ��"
        Case "����"
            strField = "trunc((sysdate-��X.����������)/365)"
        Case "������", "�����", "��׼��", "��������", "���Ĳ���", "������;"
            If str������ = "��������" Then
                strField = "��O.������ʱ��"
            Else
                strField = "��O.��" & str������
            End If
        Case Else
            If mlngIndex < lstFields.ListIndex Then
                strField = "��L" & lstFields.ListIndex & ".����Ϣ��='" & str������ & "' AND ��L" & lstFields.ListIndex & ".����Ϣֵ"
            Else
                strField = "��A.��" & str������
            End If
    End Select
    If cmbList.Visible = True Then
        '�̶�ȡֵ�б�
        If cmbList.Text = "" Then
            strTemp = IIf(Left(cmbOperate.Text, 2) = "����", " IS ", " IS NOT ") & " NULL "
        Else
            strTemp = Right(cmbOperate.Text, 5) & cmbList.ItemData(cmbList.ListIndex)
            '���˺�:2007/11/05����
            If str������ = "�������" Then
                If cmbList.Text = "��Ժ��Ҫ���" Then
                    strTemp = strTemp & " And ��C.����ϴ��� " & Right(cmbOperate.Text, 5) & "1"
                ElseIf cmbList.Text = "��Ժ��Ҫ���" Then
                    strTemp = strTemp & IIf(Left(cmbOperate.Text, 2) = "����", " And ��C.����ϴ���> 1", " And ��C.����ϴ���<=1")
                End If
            End If
        End If
    Else
        
        If Trim(cmbExample.Text) = "" Then
            strTemp = IIf(Left(cmbOperate.Text, 2) = "����", " IS ", " IS NOT ") & " NULL "
        Else
            Select Case lstFields.ItemData(lstFields.ListIndex)
                Case 2 'Number
                    strTemp = Right(cmbOperate.Text, 5) & cmbExample.Text & " "
                Case 3 'VarChar
                    strTemp = Replace(cmbExample.Text, "'", "''")
                    Select Case str������
                        Case "���֤��", "��������", "��������", "��ϱ���", "��������", "��ϼ���"
                            strTemp = UCase(strTemp)
                    End Select
                    If Left(cmbOperate, 2) = "����" Then
                        strTemp = " LIKE '%" & strTemp & "%'"
                    ElseIf Left(cmbOperate, 2) = "��ͷ" Then
                        strTemp = " LIKE '" & strTemp & "%'"
                    Else
                        strTemp = Right(cmbOperate.Text, 5) & "'" & strTemp & "'"
                    End If
                Case 4 'Date
                    strTemp = Right(cmbOperate.Text, 5) & "TO_DATE('" & Format(CDate(cmbExample.Text), "yyyy-MM-dd") & "','YYYY-MM-DD')"
            End Select
        End If
    End If
    If bln���� Then
        '��ǰһֵ���ۼ�
        If cmbLogical.Text = "����" Then
            If lstFields.ItemData(lstFields.ListIndex) = 4 Then
                    lst.Tag = lst.Tag & " AND (trunc(" & strField & ") " & strTemp & ")"
            ElseIf str������ = "������" Then
                    lst.Tag = lst.Tag & " AND (" & strField & strTemp & ") AND (��A.��������" & strTemp & ")"
            Else
                    lst.Tag = lst.Tag & " AND (" & strField & strTemp & ")"
            End If
        Else
            If lstFields.ItemData(lstFields.ListIndex) = 4 Then
                    lst.Tag = lst.Tag & " OR (trunc(" & strField & ") " & strTemp & ")"
            ElseIf str������ = "������" Then
                lst.Tag = lst.Tag & " OR (" & strField & strTemp & ") or (��A.��������" & strTemp & ")"
            Else
                lst.Tag = lst.Tag & " OR (" & strField & strTemp & ")"
            End If
        End If
    Else
        If lstFields.ItemData(lstFields.ListIndex) = 4 Then
            lst.Tag = "(trunc(" & strField & ") " & strTemp & ")"
        ElseIf str������ = "������" Then
            lst.Tag = lst.Tag & "(" & strField & strTemp & " or ��A.��������" & strTemp & ")"
        Else
            lst.Tag = "(" & strField & strTemp & ")"
        End If
    End If
    cmd(2).Enabled = True
End Sub

Private Function ValidateDate() As Boolean
    
    If cmbExample.Visible = True Then
        If (Left(cmbOperate.Text, 2) <> "����" And Left(cmbOperate.Text, 2) <> "����") And Trim(cmbExample.Text = "") Then
            MsgBox "��ǰ�����²������ֵ��" & vbCrLf & "�����Ҫʹ�ÿ�ֵ��������ֻ���ǡ����ڡ��򡰲����ڡ���", vbInformation, gstrSysName
            cmbExample.SetFocus
            Exit Function
        End If
        If zlCommFun.ActualLen(cmbExample.Text) > 100 Then
            MsgBox "�������ݹ�����", vbInformation, gstrSysName
            cmbExample.SetFocus
            Exit Function
        End If
        Select Case lstFields.ItemData(lstFields.ListIndex)
            Case 2 'Number
                If Not IsNumeric(cmbExample.Text) And Trim(cmbExample.Text) <> "" Then
                    MsgBox "������һ���Ϸ������֡�", vbInformation, gstrSysName
                    cmbExample.SetFocus
                    Exit Function
                End If
            Case 3 'VarChar
                If InStr(cmbExample.Text, "'") > 0 Then
                    MsgBox "�����˷Ƿ��ַ���", vbInformation, gstrSysName
                    cmbExample.SetFocus
                    Exit Function
                End If
            Case 4 'Date
                If Not IsDate(cmbExample.Text) And Trim(cmbExample.Text) <> "" Then
                    MsgBox "������һ���Ϸ������ڡ�" & vbCrLf & "���磺1997-07-01��", vbInformation, gstrSysName
                    cmbExample.SetFocus
                    Exit Function
                End If
        End Select
    End If
    ValidateDate = True
End Function