VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.12#0"; "zlIDKind.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmManageDeposit 
   AutoRedraw      =   -1  'True
   Caption         =   "����Ԥ������"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8910
   Icon            =   "frmManageDeposit.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   5700
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picSearch 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   5940
      ScaleHeight     =   375
      ScaleWidth      =   2910
      TabIndex        =   5
      Top             =   195
      Width           =   2910
      Begin VB.TextBox txtValue 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   615
         TabIndex        =   6
         ToolTipText     =   "��λF3"
         Top             =   40
         Width           =   2235
      End
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   315
         Left            =   15
         TabIndex        =   7
         Top             =   48
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         Appearance      =   2
         IDKindStr       =   $"frmManageDeposit.frx":08CA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   9
         FontName        =   "����"
         IDKind          =   -1
         ShowPropertySet =   -1  'True
         DefaultCardType =   "0"
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         BackColor       =   -2147483633
      End
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   5490
      Top             =   750
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":0951
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":0B6B
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":0D85
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":0F9F
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":11B9
            Key             =   "View"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":1933
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":1B4D
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":1D67
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":1F81
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":219B
            Key             =   "RollingCurtain"
            Object.Tag             =   "RollingCurtain"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   4605
      Top             =   765
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":BB32
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":BD4C
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":BF66
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":C180
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":C39A
            Key             =   "View"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":CB14
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":CD2E
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":CF48
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":D162
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":D37C
            Key             =   "RollingCurtain"
            Object.Tag             =   "RollingCurtain"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.TabStrip tbPage 
      Height          =   300
      Left            =   75
      TabIndex        =   4
      Top             =   825
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   529
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid mshList 
      Height          =   3225
      Left            =   180
      TabIndex        =   3
      Top             =   1515
      Width           =   8670
      _cx             =   15293
      _cy             =   5689
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
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
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   5340
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmManageDeposit.frx":DA76
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6906
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3704
            MinWidth        =   3704
            Picture         =   "frmManageDeposit.frx":E30A
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
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   8910
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinWidth1       =   1995
      MinHeight1      =   720
      Width1          =   2010
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   8790
         _ExtentX        =   15505
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   14
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Preview"
               Description     =   "Ԥ��"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Description     =   "��ӡ"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�տ�"
               Key             =   "Deposit"
               Description     =   "�տ�"
               Object.ToolTipText     =   "�����տ��"
               Object.Tag             =   "�տ�"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˿�"
               Key             =   "Del"
               Description     =   "�˿�"
               Object.ToolTipText     =   "�Ե�ǰѡ�е����˿�"
               Object.Tag             =   "�˿�"
               ImageKey        =   "Del"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "View"
               Description     =   "����"
               Object.ToolTipText     =   "���ĵ�ǰ���ݵ�����"
               Object.Tag             =   "����"
               ImageKey        =   "View"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Filter"
               Description     =   "����"
               Object.ToolTipText     =   "�����������¶�ȡ�б�"
               Object.Tag             =   "����"
               ImageKey        =   "Filter"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��λ"
               Key             =   "Go"
               Description     =   "��λ"
               Object.ToolTipText     =   "��λ�ڵ�ǰ�б������������ļ�¼��"
               Object.Tag             =   "��λ"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "�շ�����"
               Object.Tag             =   "����"
               ImageKey        =   "RollingCurtain"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "SplitRollingCurtain"
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "����"
               Object.ToolTipText     =   "��ǰ��������"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Quit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFile_PrintSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFile_PreView 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFile_Print 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFile_Excel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMoneyEnum 
         Caption         =   "�ֽ�㳮(&E)"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRollingCurtain 
         Caption         =   "�շ�����(&M)"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuFileRollingCurtainSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileLocalSet 
         Caption         =   "��������(&R)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFileLocalSet_ 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuFile_quit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEdit_Deposit 
         Caption         =   "�տ�(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEdit_Del 
         Caption         =   "�˿�(&D)"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu mnuEidtBalanceDel 
         Caption         =   "����˿�(&Y)"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuEdit_View 
         Caption         =   "����(&V)"
      End
      Begin VB.Menu mnuEditSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditMzTozy 
         Caption         =   "����תסԺ(&M)"
      End
      Begin VB.Menu mnuEditZyToMz 
         Caption         =   "סԺת����(&Z)"
      End
      Begin VB.Menu mnuEdit_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Print 
         Caption         =   "�ش�Ʊ��(&R)"
      End
      Begin VB.Menu mnuEdit_Print_Supplemental 
         Caption         =   "����Ʊ��(&B)"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "����(&R)"
      Visible         =   0   'False
      Begin VB.Menu mnuReportItem 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "��׼��ť(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuView_Tlb_1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuView_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFilter 
         Caption         =   "����(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewGo 
         Caption         =   "��λ(&G)"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewreFlash 
         Caption         =   "ˢ��(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "��������(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&WEB�ϵ�����"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "������̳(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&K)..."
         End
      End
      Begin VB.Menu mnuHelp_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frmManageDeposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private mrsList As ADODB.Recordset  '�����б�
Private mstrFilter As String
Private mblnCancel As Boolean
Private mblnGo As Boolean, mlngGo As Long
Private mlngCurRow As Long, mlngTopRow As Long
Private mstrPrivs As String
Private mlngModul As Long
Private mblnNOMoved As Boolean '����ϸʱ��¼��ǰѡ��ĵ����Ƿ����������ݱ���,����������ʱ�������ж�
Private mcllFilterA As Collection
Private mblnNotClick As Boolean
Private mblnUnLoad As Boolean
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mstrPrivs_RollingCurtain As String  '�շ����ʹ���Ȩ��

Private Sub InitFilter()
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '���:
    '����:
    '����:
    '����:lesfeng
    '����:2010-01-11 16:10:40
    '-----------------------------------------------------------------------------------------------------------
    Set mcllFilterA = New Collection
    mcllFilterA.Add Array("1901-01-01 00:00:00", "1901-01-01 00:00:00"), "�տ�ʱ��"
    mcllFilterA.Add Array("", ""), "���ݺ�"
    mcllFilterA.Add Array("", ""), "Ʊ�ݺ�"
    mcllFilterA.Add "", "�����"
    mcllFilterA.Add "", "סԺ��"
    mcllFilterA.Add "", "����"
    mcllFilterA.Add "", "�������"
    mcllFilterA.Add "", "�տ���"
    mstrFilter = ""
End Sub
Private Sub InitPrepayType()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��Ԥ������
    '����:���˺�
    '����:2011-07-14 18:50:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    tbPage.Tabs.Clear
    mblnNotClick = True
    If InStr(1, mstrPrivs, ";����Ԥ��;") > 0 _
        And InStr(1, mstrPrivs, ";סԺԤ��;") > 0 Then
        tbPage.Tabs.Add , "ALL", "����Ԥ��"
        tbPage.Tabs("ALL").Selected = True
    End If
    If InStr(1, mstrPrivs, ";����Ԥ��;") > 0 Then
        tbPage.Tabs.Add , "K1", "����Ԥ��"
    End If
    If InStr(1, mstrPrivs, ";סԺԤ��;") > 0 Then
        tbPage.Tabs.Add , "K2", "סԺԤ��"
    End If
    If tbPage.SelectedItem Is Nothing And tbPage.Tabs.Count <> 0 Then
        tbPage.Tabs(0).Selected = True
    End If
    If tbPage.Tabs.Count = 0 Then
        MsgBox "�㲻�߱�����Ԥ����סԺԤ��Ȩ��,����ϵͳ����Ա��ϵ!", vbOKOnly + vbInformation, gstrSysName
        mblnUnLoad = True
    End If
    mblnNotClick = False
End Sub

 
Private Sub cboType_Click()
    If mblnNotClick Then Exit Sub
    ShowBills mstrFilter
End Sub

Private Sub cbr_Resize()
     Call Form_Resize
End Sub

Private Sub Form_Activate()
    If mblnUnLoad Then Unload Me: Exit Sub
    
    Call InitLocPar(mlngModul)
End Sub

Public Sub ActiveIDKindKey()
    IDKind.ActiveFastKey
End Sub

Private Sub mnuEditMzTozy_Click()
    '����תסԺ
      If frmDeposit.zlShowEdit(Me, 0, 4, mstrPrivs, mlngModul, 1) = True Then
        If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then mnuViewReFlash_Click
    End If
End Sub

Private Sub mnuEditZyToMz_Click()
    '����תסԺ
      If frmDeposit.zlShowEdit(Me, 0, 4, mstrPrivs, mlngModul, 2) = True Then
        If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then mnuViewReFlash_Click
    End If
End Sub

Private Sub mnuEidtBalanceDel_Click()
    '����˿�
    If frmDeposit.zlShowEdit(Me, 0, 3, mstrPrivs, mlngModul) = True Then
        If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then mnuViewReFlash_Click
    End If
End Sub

Private Sub mnuFileLocalSet_Click()
    Call frmLocalSet.zlSetPara(Me, mstrPrivs, mlngModul)
    
'    If glngԤ��ID > 0 Then
'        If Not ExistBill(glngԤ��ID, 2) Then
'            zldatabase.SetPara "����Ԥ��Ʊ������", 0, glngSys, mlngModul
'            glngԤ��ID = 0
'        End If
'    End If
End Sub

Private Sub mnuFileMoneyEnum_Click()
    Call frmMoneyEnum.ShowMe(Me)
End Sub

 

Private Sub mnuFileRollingCurtain_Click()
    Call zlExecuteChargeRollingCurtain(Me)
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim strNo As String, str����ID As String, strסԺ�� As String
    With mshList
        strNo = mshList.TextMatrix(mshList.Row, .ColIndex("���ݺ�"))
        str����ID = Trim(.TextMatrix(.Row, .ColIndex("����ID")))
        strסԺ�� = Trim(.TextMatrix(.Row, .ColIndex("סԺ��")))
    End With
    If strNo <> "" Then
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
                "NO=" & strNo, "����ID=" & str����ID, "סԺ��=" & strסԺ��)
    Else
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me)
    End If
End Sub

Private Sub mnuViewFilter_Click()
    frmDepositFilter.Show 1, Me
    If gblnOK Then
        mstrFilter = frmDepositFilter.mstrFilter
        Set mcllFilterA = frmDepositFilter.mcllFilter
        mblnCancel = (frmDepositFilter.chkCancel.Value = Checked And frmDepositFilter.chk�տ�.Value = 0)
        mnuViewReFlash_Click
    End If
End Sub
Private Sub mshList_DblClick()
    If mshList.MouseRow = 0 Then Exit Sub
    If mnuEdit_View.Enabled Then mnuEdit_View_Click
End Sub

Private Sub mshList_EnterCell()
    Dim strNo As String, lng��¼״̬ As Long
    Dim strƱ�ݺ� As String
    With mshList
        If .Row = 0 Or .TextMatrix(.Row, .ColIndex("���ݺ�")) = "" Then Exit Sub
        strNo = .TextMatrix(.Row, .ColIndex("���ݺ�"))
        lng��¼״̬ = Val(.TextMatrix(.Row, .ColIndex("��¼״̬")))
        strƱ�ݺ� = .TextMatrix(.Row, .ColIndex("Ʊ�ݺ�"))
        mlngGo = .Row: mlngCurRow = .Row: mlngTopRow = .TopRow
    End With
    If frmDepositFilter.mblnDateMoved Then
        mblnNOMoved = zlDatabase.NOMoved("����Ԥ����¼", strNo, , "1", Me.Caption)
    Else
        mblnNOMoved = False
    End If
    
    'mshList.TextMatrix(mshList.Row, mshList.Cols - 1)
    Select Case lng��¼״̬
        Case 1
            SetMenu (True)
            mnuEdit_Print_Supplemental.Enabled = strƱ�ݺ� = ""
        Case 2
            SetMenu (False)
            mnuEdit_View.Enabled = True
            tbr.Buttons.Item("View").Enabled = True
        Case 3
            SetMenu (False)
            mnuEdit_View.Enabled = True
            tbr.Buttons.Item("View").Enabled = True
    End Select
    
End Sub
Private Sub mshList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And mnuEdit_Del.Enabled And mnuEdit_Del.Visible Then Call mnuEdit_Del_Click
End Sub

Private Sub mshList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuEdit, 2
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            'ʼ�մӵ�ǰ�п�ʼ
            If mnuViewGo.Enabled Then Call SeekBill(False)
'        Case vbKeyReturn
'            If mnuEdit_View.Enabled Then mnuEdit_View_Click
        Case vbKeyEscape
            mblnGo = False
        Case Else
            IDKind.ActiveFastKey
    End Select
End Sub

Private Sub mnuEdit_Del_Click()
    Dim strNo As String, str����Ա As String
    Dim bytԤ������ As Byte
    With mshList
        strNo = .TextMatrix(.Row, .ColIndex("���ݺ�"))
        str����Ա = .TextMatrix(.Row, .ColIndex("����Ա"))
        bytԤ������ = Val(.TextMatrix(.Row, .ColIndex("Ԥ�����ID")))
    End With
    If strNo = "" Then
        MsgBox "��ǰû�м�¼�����˿", vbExclamation, gstrSysName
        Exit Sub
    End If
        
    '����Ȩ��
    If Not BillOperCheck(6, str����Ա, _
        CDate(mshList.TextMatrix(mshList.Row, mshList.ColIndex("����ʱ��"))), "�˿�") Then Exit Sub
    
    If Val(mshList.TextMatrix(mshList.Row, mshList.ColIndex("���"))) < 0 Then
        MsgBox "�ýɿ��¼���Ϊ��,��ʾ�˿�,����ִ�иò�����", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    '�Ƿ���ת������ݱ���
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNo, 6, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '��ʱ��ת���������ݱ�
    End If
    
    If is���տ�(strNo) Then
         If InStr(mstrPrivs, "���տ��˿�") = 0 Then
            MsgBox "��û��Ȩ�޽��д��տ��˿������", vbInformation, gstrSysName
            Exit Sub
        End If
    ElseIf InStr(mstrPrivs, "Ԥ���˿�") = 0 Then
        MsgBox "��û��Ȩ�޽���Ԥ���˿������", vbInformation, gstrSysName
        Exit Sub
    Else
        If HaveSpare(strNo) = 0 And InStr(mstrPrivs, "Ԥ�������˿�") = 0 Then
            MsgBox "�ò�����û��Ԥ�����,��û��Ȩ���������ŵ��ݣ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If HaveBalance(strNo) <> 0 Then
            MsgBox "�ñ�Ԥ���Ѿ��������ڽ���ʱʹ��,�㲻���������ŵ��ݣ�", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    On Error Resume Next
    Err.Clear
    If frmDeposit.zlShowEdit(Me, 0, 2, mstrPrivs, mlngModul, bytԤ������, strNo) = True Then
        If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then mnuViewReFlash_Click
    End If
End Sub

Private Sub mnuHelpTitle_Click()
ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub mnuEdit_Deposit_Click()
    Dim byt���� As Byte
    If Not tbPage.SelectedItem Is Nothing Then
        byt���� = Val(Mid(tbPage.SelectedItem.Key, 2))
    End If
    If frmDeposit.zlShowEdit(Me, 0, 0, mstrPrivs, mlngModul, byt����) = True Then
        If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then mnuViewReFlash_Click
    End If
End Sub

Private Sub mnuEdit_View_Click()
    Dim blnViewCancel  As Boolean
    Dim strNo As String, str����Ա As String, bytԤ������ As Byte
    Dim int��¼״̬ As Integer, blnNOMoved As Boolean
    With mshList
        strNo = .TextMatrix(.Row, .ColIndex("���ݺ�"))
        str����Ա = .TextMatrix(.Row, .ColIndex("����Ա"))
        bytԤ������ = Val(.TextMatrix(.Row, .ColIndex("Ԥ�����ID")))
        int��¼״̬ = Val(.TextMatrix(.Row, .ColIndex("��¼״̬")))
        blnViewCancel = int��¼״̬ = 2
    End With
    '�Ƿ���ת������ݱ���
    If mblnNOMoved Then
        blnNOMoved = zlDatabase.NOMoved("����Ԥ����¼", strNo, , "1")
    End If
    
    If strNo = "" Then MsgBox "��ǰû�м�¼���Բ��ģ�", vbExclamation, gstrSysName: Exit Sub
    '��ʾ��������
    Call frmDeposit.zlShowEdit(Me, 0, 1, mstrPrivs, mlngModul, bytԤ������, strNo, blnViewCancel, blnNOMoved)

End Sub

Private Sub mnuFile_Quit_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuViewReFlash_Click()
    ShowBills mstrFilter
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = Not stbThis.Visible
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    cbr.Visible = Not cbr.Visible
    mnuViewToolText.Enabled = Not mnuViewToolText.Enabled
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim i As Integer
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For i = 1 To tbr.Buttons.Count
        tbr.Buttons(i).Caption = IIf(mnuViewToolText.Checked, tbr.Buttons(i).Tag, "")
    Next
    cbr.Bands(1).MinHeight = tbr.ButtonHeight
    Form_Resize
End Sub

Private Sub tbPage_Click()
        If mblnNotClick Then Exit Sub
        ShowBills mstrFilter
       If mshList.Enabled And mshList.Visible Then mshList.SetFocus
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuFile_Quit_Click
        Case "Go" '��λ
            mnuViewGo_Click
        Case "Filter" '����
            mnuViewFilter_Click
        Case "View"
            mnuEdit_View_Click
        Case "Deposit"
            mnuEdit_Deposit_Click
        Case "Del"
            mnuEdit_Del_Click
        Case "Print"
            mnuFile_Print_Click
        Case "Preview"
            mnuFile_PreView_Click
        Case "����"
            mnuFileRollingCurtain_Click
        Case "Help"
            mnuHelpTitle_Click
    End Select
End Sub

Private Sub tbr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
End Sub

Private Sub mnuFile_Excel_Click()
    Call OutputList(3)
End Sub

Private Sub mnuFile_PreView_Click()
    Call OutputList(2)
End Sub

Private Sub mnuFile_Print_Click()
    Call OutputList(1)
End Sub

Private Sub mnuFile_PrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub OutputList(bytStyle As Byte)
'���ܣ�������б�
'������bytStyle=1-��ӡ,2-Ԥ��,3-�����Excel
    Dim objOut As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim bytR As Byte, intRow As Integer
    
    intRow = mshList.Row
    
    '��ͷ
    objOut.Title.Text = "Ԥ�����տ��嵥"
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '����
    With frmDepositFilter
        If IsNull(.dtpEnd.Value) Then
            objRow.Add "ʱ�䣺" & Format(.dtpBegin.Value, "yyyy-MM-dd")
        Else
            objRow.Add "ʱ�䣺" & Format(.dtpBegin.Value, "yyyy-MM-dd") & " �� " & Format(.dtpEnd.Value, "yyyy-MM-dd")
        End If
        objRow.Add "���ʣ�" & IIf(.chkCancel.Value = 1, "�˿��¼", "�տ��¼")
        objOut.UnderAppRows.Add objRow
    End With
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate(), "yyyy��MM��dd��")
    objOut.BelowAppRows.Add objRow
    
    '����
    mshList.Redraw = False
    Set objOut.Body = mshList
    
    '���
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    mshList.Row = intRow
    mshList.Col = 0: mshList.ColSel = mshList.Cols - 1
    mshList.Redraw = True
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage hWnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo hWnd
End Sub

Private Sub SetHeader()
    Dim strHead As String
    Dim i As Integer
    Dim lngRow As Long, lngCol As Long
    
 
    strHead = "���ݺ�,4,850|Ʊ�ݺ�,4,1050|����Ա,1,850|����ʱ��,4,1850|����ID,1,750|�����,1,750|סԺ��,1,750|����,1,800|�Ա�,4,500|����,4,500|����,1,850|���,7,850|���㷽ʽ,1,850|�������,1,1500|ժҪ,1,1500|��¼״̬,1,0|ҽ�Ƹ��ʽ,1,1500|Ԥ�����,4,800|Ԥ�����ID,1,0"
    With mshList
        .Redraw = False
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .FixedAlignment(i) = 4
            .ColKey(i) = UCase(Trim(.TextMatrix(0, i)))
        Next
        If Not Visible Then Call RestoreFlexState(mshList, App.ProductName & "\" & Me.Name)
        .ColHidden(.ColIndex("Ԥ�����ID")) = True
        .RowHeight(0) = 320
        '�ָ��ϴ���
        If mlngCurRow = 0 Then mlngCurRow = 1
        If mlngTopRow = 0 Then mlngTopRow = 1
        If mlngCurRow <= .Rows - 1 Then
            .Row = mlngCurRow
        Else
            .Row = .Rows - 1
        End If
        If mlngTopRow <= .Rows - 1 Then
            .TopRow = mlngTopRow
        Else
            .TopRow = .Row
        End If
        .Col = 0: .ColSel = .Cols - 1
        Call mshList_EnterCell
        For lngRow = 1 To .Rows - 1
            .Row = lngRow
            For lngCol = 1 To .Cols - 1
                .Col = lngCol
                If .TextMatrix(lngRow, .ColIndex("��¼״̬")) = "2" Then
                    .CellForeColor = &HFF&
                ElseIf .TextMatrix(lngRow, .ColIndex("��¼״̬")) = "3" Then
                    .CellForeColor = &HFF0000
                End If
                
                If .TextMatrix(0, lngCol) = "���" And IsNumeric(.TextMatrix(lngRow, lngCol)) Then
                   .TextMatrix(lngRow, lngCol) = Format(.TextMatrix(lngRow, lngCol), "0.00")
                End If
            Next lngCol
        Next lngRow
        .Redraw = True
    End With
End Sub

Private Sub ShowBills(Optional strIF As String, Optional blnSort As Boolean, Optional bytMode As Byte = 0, Optional objCard As Card)
'����:��������ȡ�����б�(���˹���)
'����:strIF=��"AND"��ʼ��������
'     blnSort=�����¶�ȡ����,��������ʾ�����������
    Dim dbl���  As Double, strKind As String, strFind As String
    Dim intԤ����� As Integer, strWhere As String, lng�����ID As Long
    Dim lng����ID As Long, strPassWord As String, strErrMsg As String
    Dim strTable As String
    On Error GoTo errH
    
    If bytMode <> 0 Then
        If txtValue.Text = "" Then Exit Sub
        strKind = objCard.����
        If (Left(txtValue.Text, 1) = "-" And IsNumeric(Mid(txtValue.Text, 2))) Then
            lng����ID = Val(Mid(txtValue.Text, 2))
            strIF = " And B.����ID=[1]"
            strFind = lng����ID
        '89607: ���ϴ�,2015/10/20,����סԺ�Ų���,����ֵȡ��Ч����
        ElseIf (Left(txtValue.Text, 1) = "*" And IsNumeric(Mid(txtValue.Text, 2))) Then
            strFind = Val(Mid(txtValue.Text, 2))
            strIF = " And B.�����=[1]"
        ElseIf (Left(txtValue.Text, 1) = "+" And IsNumeric(Mid(txtValue.Text, 2))) Then
            strFind = Val(Mid(txtValue.Text, 2))
            strIF = " And B.����ID=(Select Nvl(Max(����ID),0) as ����ID From ������ҳ Where סԺ��=[1])"
        Else
            Select Case strKind
            Case "����"
                lng�����ID = IDKind.GetDefaultCardTypeID
                If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, txtValue.Text, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                If lng����ID <= 0 Then
                    strFind = txtValue.Text & "%"
                    strIF = " And B.���� Like [1]"
                Else
                    strFind = lng����ID
                    strIF = " And B.����ID=[1]"
                End If
            Case "���֤��"
                strFind = txtValue.Text
                strIF = " And B.���֤��=[1]"
            Case "ҽ����"
                strFind = txtValue.Text
                strIF = " And B.ҽ����=[1]"
            Case "IC����"
                If gobjSquare.objSquareCard.zlGetPatiID("IC����", txtValue.Text, False, lng����ID, _
                    strPassWord, strErrMsg) = False Then lng����ID = 0
                strFind = lng����ID
                strIF = " And B.����ID=[1]"
            Case "�����"
                strFind = txtValue.Text
                strIF = " And B.�����=[1]"
            Case "סԺ��"
                strFind = txtValue.Text
                strIF = " And B.����ID=(Select Nvl(Max(����ID),0) as ����ID From ������ҳ Where סԺ��=[1])"
            Case Else
                '��������,��ȡ��صĲ���ID
                '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|
                '�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������)
                '��7λ��,��ֻ��������,��Ȼȡ������
                lng�����ID = objCard.�ӿ����
                If lng�����ID <> 0 Then
                    If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, txtValue.Text, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(strKind, txtValue.Text, False, lng����ID, _
                        strPassWord, strErrMsg) = False Then lng����ID = 0
                End If
                strFind = lng����ID
                strIF = " And B.����ID=[1]"
            End Select
        End If
    End If
    
    If Not blnSort Then
        Call zlCommFun.ShowFlash("���ڶ�ȡ�����б�,���Ժ� ...", Me)
        DoEvents
        Me.Refresh
        strWhere = ""
        If tbPage.SelectedItem Is Nothing Then Exit Sub
        
        If gbln��վ����ʾ Then
             strWhere = strWhere & _
            " ��And Exists (Select 1 From ��Ա�� E, ������Ա F, ���ű� G " & _
            " ��Where A.����Ա����=e.����  And e.Id = f.��Աid And f.����id = g.Id And (g.վ�� ='" & gstrNodeNo & "' Or g.վ�� Is Null))"
        End If
        
        '115601:���ϴ�,2017/10/23,ָ����ѯ��
        If frmDepositFilter.mblnDateMoved Then
            strTable = "(Select NO,ʵ��Ʊ��,����Ա����,�տ�ʱ��,����ID,���,���㷽ʽ,�������,ժҪ,��¼״̬,��ҳID,Ԥ�����,����ID,��¼���� From ����Ԥ����¼  " & _
                        "UNION ALL " & _
                        "Select NO,ʵ��Ʊ��,����Ա����,�տ�ʱ��,����ID,���,���㷽ʽ,�������,ժҪ,��¼״̬,��ҳID,Ԥ�����,����ID,��¼���� From H����Ԥ����¼)"
        Else
            strTable = "����Ԥ����¼"
        End If
            
        If Left(tbPage.SelectedItem.Key, 1) = "K" Then
            intԤ����� = Val(Mid(tbPage.SelectedItem.Key, 2))
            strWhere = " And  Nvl(A.Ԥ�����, 0) = " & IIf(bytMode = 0, "[12]", "[2]")
        End If
          
         gstrSQL = _
        "   Select A.NO as ���ݺ�,A.ʵ��Ʊ�� as Ʊ�ݺ�,A.����Ա���� as ����Ա," & _
        "           To_Char(A.�տ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as ����ʱ��," & _
        "           A.����ID,B.�����,B.סԺ��,B.����,B.�Ա�,B.����,D.���� as ����," & _
        "           To_Char(Sum(A.���),'9999999990.00') as ���," & _
        "           A.���㷽ʽ,A.�������,A.ժҪ,A.��¼״̬," & _
        "           Decode(Nvl(A.��ҳID,0),0,B.ҽ�Ƹ��ʽ,C.ҽ�Ƹ��ʽ) ҽ�Ƹ��ʽ, " & _
        "           Decode(nvl(A.Ԥ�����,2),1,'����Ԥ��', 'סԺԤ��') as Ԥ�����, nvl(A.Ԥ�����,0) as Ԥ�����ID" & _
        " From " & strTable & " A,������Ϣ B,������ҳ C,���ű� D " & _
        " Where A.����ID=B.����ID AND A.����ID=C.����ID(+) AND NVL(A.��ҳID,0)=C.��ҳID(+) And A.����ID=D.ID(+) And A.��¼����=1 " & strIF & strWhere & _
        " Group by A.NO,A.��¼״̬,A.ʵ��Ʊ�� ,Nvl(A.Ԥ�����, 0),Decode(nvl(A.Ԥ�����,2),1,'����Ԥ��', 'סԺԤ��'),A.����Ա����," & _
        "           To_Char(A.�տ�ʱ��,'YYYY-MM-DD HH24:MI:SS'),A.����ID,B.�����,B.סԺ��,B.����,B.����," & _
        "           B.�Ա� , D.����, A.���㷽ʽ,A.�������, A.ժҪ,Decode(Nvl(A.��ҳid, 0), 0, B.ҽ�Ƹ��ʽ, C.ҽ�Ƹ��ʽ)" & _
        " Order by ����ʱ�� Desc,���ݺ� Desc"
        
        Set mrsList = New ADODB.Recordset
            
        If bytMode = 0 Then
            Set mrsList = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CDate(mcllFilterA("�տ�ʱ��")(0)), CDate(mcllFilterA("�տ�ʱ��")(1)), _
            CStr(mcllFilterA("���ݺ�")(0)), CStr(mcllFilterA("���ݺ�")(1)), _
            CStr(mcllFilterA("Ʊ�ݺ�")(0)), CStr(mcllFilterA("Ʊ�ݺ�")(1)), CLng(Val(mcllFilterA("סԺ��"))), _
            CStr(mcllFilterA("����")), CStr(mcllFilterA("�������")), CStr(mcllFilterA("�տ���")), CLng(Val(mcllFilterA("�����"))), intԤ�����)
        Else
            Set mrsList = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strFind, intԤ�����)
        End If
    End If
    
    mshList.Clear
    mshList.Rows = 2
    If mrsList.EOF Then
        Call SetHeader
        stbThis.Panels(2).Text = "��ǰ����û�й��˳��κε���"
        Call SetMenu(False)
    Else
        If (Left(txtValue.Text, 1) = "-" And IsNumeric(Mid(txtValue.Text, 2))) Then
            txtValue.Text = NVL(mrsList!����)
            IDKind.IDKind = 1
        ElseIf (Left(txtValue.Text, 1) = "*" And IsNumeric(Mid(txtValue.Text, 2))) Then
            txtValue.Text = NVL(mrsList!����)
            IDKind.IDKind = 1
        ElseIf (Left(txtValue.Text, 1) = "+" And IsNumeric(Mid(txtValue.Text, 2))) Then
            txtValue.Text = NVL(mrsList!����)
            IDKind.IDKind = 1
        End If
        
        Set mshList.DataSource = mrsList: Call SetHeader
        mrsList.MoveFirst: dbl��� = 0
        Do While Not mrsList.EOF
            dbl��� = dbl��� + mrsList!���
            mrsList.MoveNext
        Loop
        mrsList.MoveFirst
        stbThis.Panels(2) = "�� " & mrsList.RecordCount & " �ŵ���,�ϼ�:" & Format(dbl���, "0.00")
        Call SetMenu(True)
    End If
    mnuEdit_Del.Enabled = Not mblnCancel And Not mrsList.EOF
    mnuEdit_Print.Enabled = Not mblnCancel And Not mrsList.EOF
    mnuEdit_Print_Supplemental.Enabled = Not mblnCancel And Not mrsList.EOF
    tbr.Buttons("Del").Enabled = Not mblnCancel And Not mrsList.EOF
    If Not blnSort Then Call zlCommFun.StopFlash
    Me.Refresh
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub SetMenu(blnUsed As Boolean)
'���ܣ��������޼�¼���ò˵�����״̬
    mnuFile_Print.Enabled = blnUsed
    mnuFile_PreView.Enabled = blnUsed
    mnuFile_Excel.Enabled = blnUsed
    tbr.Buttons("Print").Enabled = blnUsed
    tbr.Buttons("Preview").Enabled = blnUsed
    
    mnuEdit_Del.Enabled = blnUsed
    mnuEdit_View.Enabled = blnUsed
    mnuEdit_Print.Enabled = blnUsed
    mnuEdit_Print_Supplemental.Enabled = blnUsed
    tbr.Buttons("Del").Enabled = blnUsed
    tbr.Buttons("View").Enabled = blnUsed
    
    mnuViewGo.Enabled = blnUsed
    tbr.Buttons("Go").Enabled = blnUsed
End Sub

Private Sub Form_Load()
    Dim Curdate As Date
    Dim blnHavePrivs As Boolean
    
    mstrPrivs_RollingCurtain = ";" & GetPrivFunc(glngSys, 1506) & ";"
    mstrPrivs = gstrPrivs: mlngModul = glngModul
    mblnUnLoad = False
    Call InitFilter: Call InitPrepayType

    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs, "ZL" & glngSys \ 100 & "_INSIDE_1103_1")  '����Ԥ����ɿ���
    Call RestoreWinState(Me, App.ProductName)
     
    
    'Ȩ������
    If InStr(mstrPrivs, "Ԥ���տ�") = 0 And InStr(mstrPrivs, "���տ���ȡ") = 0 Then
        mnuEdit_Deposit.Visible = False
        tbr.Buttons("Deposit").Visible = False
        mnuEdit_Print.Visible = False
        mnuEdit_2.Visible = False
    End If
    '52328
    mnuEdit_Print_Supplemental.Visible = _
        (InStr(mstrPrivs, ";���տ���ȡ;") > 0 Or InStr(mstrPrivs, ";Ԥ���տ�;") > 0) _
        And InStr(mstrPrivs, ";����Ʊ��;") > 0
        
    If InStr(mstrPrivs, "Ԥ���˿�") = 0 And InStr(mstrPrivs, "���տ��˿�") = 0 Then
        mnuEdit_Del.Visible = False
        tbr.Buttons("Del").Visible = False
    End If
    mnuEidtBalanceDel.Visible = InStr(1, mstrPrivs, ";Ԥ���˿�;") > 0
    mnuEditMzTozy.Visible = InStr(1, mstrPrivs, ";����Ԥ��תסԺ;") > 0
    mnuEditZyToMz.Visible = InStr(1, mstrPrivs, ";סԺԤ��ת����;") > 0
    mnuEditSplit.Visible = InStr(1, mstrPrivs, ";����Ԥ��תסԺ;") > 0 Or InStr(1, mstrPrivs, ";סԺԤ��ת����;") > 0
    '�շ����ʹ���
    blnHavePrivs = InStr(mstrPrivs_RollingCurtain, ";����;") > 0
    mnuFileRollingCurtain.Visible = blnHavePrivs
    mnuFileRollingCurtainSplit.Visible = blnHavePrivs
    tbr.Buttons("����").Visible = blnHavePrivs
    tbr.Buttons("SplitRollingCurtain").Visible = blnHavePrivs

    If InStr(";" & mstrPrivs & ";", ";�ش�Ʊ��;") = 0 Then
        mnuEdit_Print.Visible = False
    End If
    
    'ȱʡ��������
    Curdate = zlDatabase.Currentdate
    'by lesfeng 2010-03-06 �����Ż�
    mstrFilter = ""
    mstrFilter = mstrFilter & " And (�տ�ʱ��  Between [1] And [2]) "
    mstrFilter = mstrFilter & " And ��¼״̬=1"
    mstrFilter = mstrFilter & " And ����Ա����=[10]"
    
    mcllFilterA.Remove "�տ�ʱ��"
    mcllFilterA.Add Array(Format(Curdate, "yyyy-mm-dd") & " 00:00:00", Format(Curdate, "yyyy-mm-dd") & " 23:59:59"), "�տ�ʱ��"
    mcllFilterA.Remove "�տ���"
    mcllFilterA.Add Trim(UserInfo.����), "�տ���"
    mblnCancel = False
    
    Call SetHeader
    Call SetMenu(False)
    Set mobjIDCard = New clsIDCard
    Set mobjICCard = New clsICCard
    Call mobjIDCard.SetParent(Me.hWnd)
    Call mobjICCard.SetParent(Me.hWnd)
    Call InitIDKind
    stbThis.Panels(2).Text = "��ˢ���嵥���������ù�������"
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If txtValue.Visible Then txtValue.SetFocus
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    txtValue.Text = objPatiInfor.����
    Call ShowBills("", , 1, objCard)
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNo As String)
    If txtValue.Locked Or txtValue.Text <> "" Or Not Me.ActiveControl Is txtValue Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("IC��", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtValue.Text = strCardNo
    Call ShowBills("", , 1, objCard)
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    If txtValue.Locked Or txtValue.Text <> "" Or Not Me.ActiveControl Is txtValue Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("���֤", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtValue.Text = strID
    Call ShowBills("", , 1, objCard)
End Sub

Private Sub txtValue_Change()
    If Me.ActiveControl Is txtValue Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtValue.Text = "")
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtValue.Text = "")
        IDKind.SetAutoReadCard txtValue.Text = ""
    End If
End Sub

Private Sub txtValue_GotFocus()
    Call zlControl.TxtSelAll(txtValue)
    Call zlCommFun.OpenIme(True)
    If txtValue.Text = "" And ActiveControl Is txtValue Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtValue.Text = "")
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtValue.Text = "")
        IDKind.SetAutoReadCard txtValue.Text = ""
    End If
End Sub

Private Sub txtValue_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        '0-�����;1-����;2-�Һŵ�;3-���￨��;4-ҽ����
        Call ShowBills("", , 1, IDKind.GetCurCard)
        zlControl.TxtSelAll txtValue
    End If
End Sub

Private Sub txtValue_KeyPress(KeyAscii As Integer)
    '0-�����,1-����,2-�Һŵ�,3-���￨��,4-ҽ����
    Dim blnCard As Boolean
    Dim strKind As String, intLen As Integer
    strKind = IDKind.GetCurCard.����
    txtValue.PasswordChar = IIf(IDKind.GetCurCard.�������Ĺ��� <> "" And IDKind.GetCurCard.�������Ĺ��� <> "0", "*", "")
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtValue.IMEMode = 0
    
    'ȡȱʡ��ˢ����ʽ
            '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|
            '�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������)
            '��7λ��,��ֻ��������,��Ȼȡ������
    Select Case strKind
    Case "����"
        blnCard = zlCommFun.InputIsCard(txtValue, KeyAscii, gobjSquare.blnȱʡ��������)
        intLen = gobjSquare.intȱʡ���ų���
    Case "�����"
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    Case "�Һŵ�"
    Case "ҽ����"
    Case Else
            If IDKind.GetCurCard.�ӿ���� <> 0 Then
                blnCard = zlCommFun.InputIsCard(txtValue, KeyAscii, IDKind.GetCurCard.�������Ĺ��� <> "" And IDKind.GetCurCard.�������Ĺ��� <> "0")
                intLen = IDKind.GetCurCard.���ų���
            End If
    End Select
    
    'ˢ����ϻ���������س�
    If blnCard And Len(txtValue.Text) = intLen - 1 And KeyAscii <> 8 Then
        If KeyAscii <> 13 Then
            txtValue.Text = txtValue.Text & Chr(KeyAscii)
            txtValue.SelStart = Len(txtValue.Text)
        End If
        KeyAscii = 0
        Call ShowBills("", , 1, IDKind.GetCurCard)
        zlControl.TxtSelAll txtValue
   End If
End Sub

Private Sub txtValue_LostFocus()
    Call zlCommFun.OpenIme
    IDKind.SetAutoReadCard False
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
End Sub

Private Sub txtValue_Validate(Cancel As Boolean)
    txtValue.Text = Trim(txtValue.Text)
End Sub

'��ʼ��IDKIND
Private Function InitIDKind() As Boolean
    Dim objCard As Card
    Dim lngCardID As Long
    Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "��|����|0|0|0|0|0|0;ҽ|ҽ����|0|0|0|0|0|0;��|���֤��|0|0|0|0|0|0;IC|IC����|1|0|0|0|0|0;��|�����|0|0|0|0|0|0;ס|סԺ��|0|0|0|0|0|0", txtValue)
    Set objCard = IDKind.GetfaultCard
    If IDKind.Cards.��ȱʡ������ And Not objCard Is Nothing Then
        gobjSquare.blnȱʡ�������� = objCard.�������Ĺ��� <> ""
        gobjSquare.intȱʡ���ų��� = objCard.���ų���
        Set gobjSquare.objDefaultCard = objCard
    Else
        gobjSquare.blnȱʡ�������� = IDKind.Cards.������ʾ
        gobjSquare.intȱʡ���ų��� = 100
    End If
End Function

Private Sub Form_Resize()
    Dim cbrH As Long '������ռ�ø߶�
    Dim staH As Long '״̬��ռ�ø߶�
    
    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    
    mshList.MousePointer = 0
    
    '����ؼ���Ⱥ͸߶�
    cbrH = IIf(cbr.Visible, cbr.Height, 0)
    staH = IIf(stbThis.Visible, stbThis.Height, 0)
    With mshList
        .Left = Me.ScaleLeft
        tbPage.Top = Me.ScaleTop + cbrH + 20
        .Top = tbPage.Top + tbPage.Height + 10
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - .Top - staH
        tbPage.Width = Me.ScaleWidth
        tbPage.Left = ScaleLeft
    End With
    picSearch.Left = Me.Width - 3500
    If picSearch.Left < 5500 Then picSearch.Left = 5500
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrFilter = ""
    Unload frmDepositFilter
    Unload frmDepositFind
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub mnuViewGo_Click()
    If Not mblnCancel Then
        frmDepositFind.lbl����Ա.Caption = "�տ���"
    Else
        frmDepositFind.lbl����Ա.Caption = "�˿���"
    End If
    frmDepositFind.Show 1, Me
    If gblnOK Then Call SeekBill(frmDepositFind.optHead)
End Sub

Private Sub SeekBill(blnHead As Boolean)
    Dim i As Long
    Dim blnFill As Boolean
    
    Screen.MousePointer = 11
    mblnGo = True
    stbThis.Panels(2).Text = "���ڶ�λ���������ĵ���,��ESC��ֹ ..."
    Me.Refresh
    
    For i = IIf(blnHead, 1, mlngGo) To mshList.Rows - 1
        DoEvents
        
        '�Ƚ�����
        blnFill = True
        With frmDepositFind
            If .txtNO.Text <> "" Then
                blnFill = blnFill And mshList.TextMatrix(i, mshList.ColIndex("���ݺ�")) = .txtNO.Text
            End If
            If .txtFact.Text <> "" Then
                blnFill = blnFill And mshList.TextMatrix(i, mshList.ColIndex("Ʊ�ݺ�")) = .txtFact.Text
            End If
            If .cbo����Ա.ListIndex > 0 Then
                If Not mblnCancel Then
                    blnFill = blnFill And mshList.TextMatrix(i, mshList.ColIndex("�տ���")) = zlCommFun.GetNeedName(.cbo����Ա.Text)
                Else
                    blnFill = blnFill And mshList.TextMatrix(i, mshList.ColIndex("�˿���")) = zlCommFun.GetNeedName(.cbo����Ա.Text)
                End If
            End If
            If .txt����.Text <> "" Then
                blnFill = blnFill And UCase(mshList.TextMatrix(i, mshList.ColIndex("����"))) Like "*" & UCase(.txt����.Text) & "*"
            End If
            If IsNumeric(.txtסԺ��.Text) Then
                blnFill = blnFill And Val(mshList.TextMatrix(i, mshList.ColIndex("סԺ��"))) = Val(.txtסԺ��.Text)
            End If
        End With
        
        '�������˳�
        If blnFill Then
            mlngGo = i + 1
            mshList.Row = i: mshList.TopRow = i
            mshList.Col = 0: mshList.ColSel = mshList.Cols - 1
            stbThis.Panels(2).Text = "�ҵ�һ����¼"
            Screen.MousePointer = 0: Exit Sub
        End If
        
        '��ESCȡ��
        If mblnGo = False Then
            stbThis.Panels(2).Text = "�û�ȡ����λ����"
            Screen.MousePointer = 0: Exit Sub
        End If
    Next
    mlngGo = 1
    stbThis.Panels(2).Text = "�Ѷ�λ���嵥β��"
    Screen.MousePointer = 0
End Sub
 

Private Sub mnuEdit_Print_Click()
    Call PrintBill(0)
End Sub

Private Sub mnuEdit_Print_Supplemental_Click()
    Call PrintBill(1)
End Sub

Private Sub PrintBill(bytMode As Byte)
    '���ܣ���ǰ�տ��¼�ش�򱻴�һ��Ʊ��
    'bytMode=0-�ش�,1-����
    Dim strSQL As String, strInvoice As String, strNo As String
    Dim lng����ID As Long, blnValid As Boolean, blnInput As Boolean, lng����ID As Long
    Dim bytԤ������ As Byte, str����Ա As String, str����ʱ�� As String, str��Ʊ�� As String
    Dim factProperty As Ty_FactProperty
    Dim strNos As String, intInvoiceFormat As Integer, blnTurnMZToZY As Boolean
                
    On Error GoTo errHandle
    With mshList
        strNo = .TextMatrix(.Row, .ColIndex("���ݺ�"))
        str����Ա = .TextMatrix(.Row, .ColIndex("����Ա"))
        bytԤ������ = Val(.TextMatrix(.Row, .ColIndex("Ԥ�����ID")))
        str����ʱ�� = .TextMatrix(.Row, .ColIndex("����ʱ��"))
        str��Ʊ�� = Trim(.TextMatrix(.Row, .ColIndex("Ʊ�ݺ�")))
        lng����ID = Val(.TextMatrix(.Row, .ColIndex("����ID")))
    End With
    If strNo = "" Then
        MsgBox "��ǰû�м�¼�����ش�Ʊ�ݣ�", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    '�Ƿ���ת������ݱ���
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNo, 6, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '��ʱ��ת���������ݱ�
    End If
    factProperty = zl_GetInvoicePreperty(mlngModul, 2, CStr(bytԤ������))
    
    strNos = GetTurnMZToZYMultiNOs(bytMode = 0, strNo, blnTurnMZToZY, mblnNOMoved)
    If blnTurnMZToZY Then
        If strNos = "" Then
            MsgBox "��ǰû�м�¼�����ش�Ʊ�ݣ�", vbExclamation, gstrSysName
            Exit Sub
        End If
        intInvoiceFormat = Val(zlDatabase.GetPara(284, glngSys, , "0"))
    Else
        strNos = strNo
        intInvoiceFormat = factProperty.intInvoiceFormat
    End If
    
    '����Ȩ��
    If bytMode = 0 Then
        If Not BillOperCheck(6, str����Ա, CDate(str����ʱ��), "�ش�") Then Exit Sub
    Else
        If str��Ʊ�� <> "" Then
            MsgBox "��ǰ�����Ѵ�ӡ��Ʊ��,���ܽ��в���", vbExclamation, gstrSysName
            Exit Sub
        End If
    End If
    
    '����ϸ����Ʊ��ʹ��
    If gblnBillԤ�� Then
        lng����ID = CheckUsedBill(2, IIf(lng����ID > 0, lng����ID, factProperty.lngShareUseID), , CStr(bytԤ������))
        Select Case lng����ID
            Case -1
                MsgBox "��û�����ú͹��õ�Ԥ��Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            Case -2
                MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
        End Select
        If lng����ID <= 0 Then Exit Sub
    End If

    If ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me) Then
        If Not gblnBillԤ�� Then
            '�п����ǵ�һ��ʹ��
            Do
                blnInput = False
                '���ϸ����ʱֱ�Ӵӱ��ض�ȡ
                strInvoice = UCase(zlDatabase.GetPara("��ǰԤ��Ʊ�ݺ�", glngSys, mlngModul, ""))
                If strInvoice = "" Then
                    strInvoice = UCase(InputBox("û���ҵ����õ����Ʊ�ݺ��룬�޷�ȷ����Ҫʹ�õĿ�ʼƱ�ݺš�" & _
                                    vbCrLf & "�����뽫Ҫʹ�õĿ�ʼƱ�ݺ��룺", gstrSysName, _
                                    "", Me.Left + 1500, Me.Top + 1500))
                    blnInput = True
                Else
                    strInvoice = zlCommFun.IncStr(strInvoice)
                    strInvoice = UCase(InputBox("��ȷ���ش�ʹ�õĿ�ʼƱ�ݺ��룺", gstrSysName, _
                                    strInvoice, Me.Left + 1500, Me.Top + 1500))
                    blnInput = True
                End If
                    
                '�û�ȡ������,�����ӡ
                If strInvoice = "" Then
                    If MsgBox("��ȷ��������Ʊ�ݺż�����ӡ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                    blnValid = True
                Else
                    '���������Ч��
                    If blnInput Then
                        If zlCommFun.ActualLen(strInvoice) <> gbytԤ�� Then
                            MsgBox "�����Ʊ�ݺ��볤��Ӧ��Ϊ " & gbytԤ�� & " λ��", vbInformation, gstrSysName
                        Else
                            blnValid = True
                        End If
                    Else
                        blnValid = True
                    End If
                End If
            Loop While Not blnValid
        Else
            Do
                '����Ʊ�����ö�ȡ
                blnInput = False
                strInvoice = GetNextBill(lng����ID)
                If strInvoice = "" Then
                    '�����;���ÿ���ĺ���,�������δ����,����һ�����ѳ�����Χ
                    strInvoice = UCase(InputBox("�޷�����Ʊ�����������ȡ��Ҫʹ�õĿ�ʼƱ�ݺţ�" & _
                                    vbCrLf & "�������뽫Ҫʹ�õĿ�ʼƱ�ݺ��룺", gstrSysName, _
                                    "", Me.Left + 1500, Me.Top + 1500))
                    blnInput = True
                Else
                    strInvoice = UCase(InputBox("��ȷ���ش�ʹ�õĿ�ʼƱ�ݺ��룺", gstrSysName, _
                                    strInvoice, Me.Left + 1500, Me.Top + 1500))
                    blnInput = True
                End If
                
                '�û�ȡ������,����ӡ
                If strInvoice = "" Then Exit Sub
                
                '���������Ч��
                If blnInput Then
                    If GetInvoiceGroupID(2, 1, lng����ID, factProperty.lngShareUseID, strInvoice, CStr(bytԤ������)) = -3 Then
                        MsgBox "�������Ʊ�ݺ��벻�ڵ�ǰ�������ε���Ч���÷�Χ��,���������룡", vbInformation, gstrSysName
                    Else
                        blnValid = True
                    End If
                Else
                    blnValid = True
                End If
            Loop While Not blnValid
        End If
        
        'ִ�����ݴ���
        'Zl_����Ԥ����¼_Reprint
        strSQL = "Zl_����Ԥ����¼_Reprint("
        '  ���ݺ�_In Varchar2,
        strSQL = strSQL & "'" & strNos & "',"
        '  Ʊ�ݺ�_In Ʊ��ʹ����ϸ.����%Type,
        strSQL = strSQL & "'" & strInvoice & "',"
        '  ����id_In Ʊ��ʹ����ϸ.����id%Type,
        strSQL = strSQL & "" & IIf(lng����ID = 0, "NULL", lng����ID) & ","
        '  ʹ����_In Ʊ��ʹ����ϸ.ʹ����%Type
        strSQL = strSQL & "'" & UserInfo.���� & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        '���Ʊ��
        '78751:���ϴ�,2014/10/20,����Ԥ��Ʊ�ݴ�ӡ��ʽ
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me, _
            "NO=" & strNos, "�տ�ʱ��=" & Format(str����ʱ��, "yyyy-mm-dd HH:MM:SS"), _
            "����ID=" & lng����ID, IIf(intInvoiceFormat = 0, "", "ReportFormat=" & intInvoiceFormat), 2)
        
        '���±���Ʊ��
        If Not gblnBillԤ�� Then
            zlDatabase.SetPara "��ǰԤ��Ʊ�ݺ�", strInvoice, glngSys, mlngModul
        End If
                
        If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then mnuViewReFlash_Click
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function GetTurnMZToZYMultiNOs(ByVal blnRePrint As Boolean, _
    ByVal strNo As String, ByRef blnTurnMZToZY As Boolean, Optional ByVal blnNOMoved As Boolean) As String
    '���ܣ���ȡ����תסԺ������Ԥ�����ݣ�������ͬʱתԤ��/һ�δ�ӡ�Ķ��ŵ��ݺ�
    '���:strNo-��Ҫ�ش�NO
    '     blnRePrint-�Ƿ��ش�Ʊ��
    '     blnNOMoved-�Ƿ�ת����ʷ��ռ�
    '����:
    '     blnTurnMZToZY-�Ƿ�����תסԺ������
    '����:һ�δ�ӡ�Ķ��ŵ��ݺţ���ʽ��A001,A002,A003,...
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strNos As String
    
    On Error GoTo errHandle
    blnTurnMZToZY = False
    strSQL = _
        "Select a.No, Max(a.��¼״̬) As ��¼״̬" & vbNewLine & _
        "From ����Ԥ����¼ A, ����Ԥ����¼ B" & vbNewLine & _
        "Where a.�տ�ʱ�� = b.�տ�ʱ�� And a.��¼���� = 1 And a.ժҪ = '����תסԺԤ��'" & vbNewLine & _
        "      And b.��¼���� = 1 And b.ժҪ = '����תסԺԤ��' And b.No = [1]" & vbNewLine & _
        "Group By a.NO" & vbNewLine & _
        "Order By NO"
    If blnNOMoved Then
        strSQL = Replace(strSQL, "����Ԥ����¼", "H����Ԥ����¼")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "", strNo)
    If rsTemp.EOF Then Exit Function
    
    blnTurnMZToZY = True
    If blnRePrint = False Then '����
        If Val(zlDatabase.GetPara(283, glngSys, , "0")) = 1 Then
            With rsTemp
                Do While Not .EOF
                    If Val(NVL(!��¼״̬)) = 1 Then '�����������ϵ���
                        strNos = strNos & "," & NVL(rsTemp!NO)
                    End If
                    .MoveNext
                Loop
            End With
            If strNos <> "" Then strNos = Mid(strNos, 2)
        Else
            '������Ƕ൥��һ�δ�ӡ����ֻ����ǰ����
            strNos = strNo
        End If
        GetTurnMZToZYMultiNOs = strNos
        Exit Function
    End If
    
    '�ش�
    'Ӧ�������һ�δ�ӡ���������
    strSQL = _
        "Select a.NO" & vbNewLine & _
        "From Ʊ�ݴ�ӡ���� A" & vbNewLine & _
        "Where a.�������� = 2" & vbNewLine & _
        "      And a.ID In (Select ID" & vbNewLine & _
        "                From (Select b.Id" & vbNewLine & _
        "                      From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B" & vbNewLine & _
        "                      Where a.��ӡid = b.Id And a.���� = 1 And a.ԭ�� In (1, 3) And b.�������� = 2 And b.No = [1]" & vbNewLine & _
        "                      Order By a.ʹ��ʱ�� Desc)" & vbNewLine & _
        "                Where Rownum < 2)" & vbNewLine & _
        "      And Not Exists(Select 1 From ����Ԥ����¼ Where ��¼���� = 1 And ��¼״̬ = 2 And No = a.No)" & vbNewLine & _
        "Order By No"
    If blnNOMoved Then
        strSQL = Replace(strSQL, "Ʊ�ݴ�ӡ����", "HƱ�ݴ�ӡ����")
        strSQL = Replace(strSQL, "Ʊ��ʹ����ϸ", "HƱ��ʹ����ϸ")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "", strNo)
    If rsTemp.EOF Then Exit Function
    
    With rsTemp
        Do While Not .EOF
            strNos = strNos & "," & NVL(rsTemp!NO)
            .MoveNext
        Loop
    End With
    If strNos <> "" Then strNos = Mid(strNos, 2)
    GetTurnMZToZYMultiNOs = strNos
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub mshList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mshList.MouseRow = 0 Then
        mshList.MousePointer = 99
    Else
        mshList.MousePointer = 0
    End If
End Sub

Private Sub mshList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long
    
    lngCol = mshList.MouseCol
    
    If Button = 1 And mshList.MousePointer = 99 Then
        If mshList.TextMatrix(0, lngCol) = "" Then Exit Sub
        If mshList.TextMatrix(1, mshList.ColIndex("���ݺ�")) = "" Then Exit Sub
        
        Set mshList.DataSource = Nothing

        mrsList.Sort = mshList.TextMatrix(0, lngCol) & IIf(mshList.ColData(lngCol) = 0, "", " DESC")
        mshList.ColData(lngCol) = (mshList.ColData(lngCol) + 1) Mod 2
        
        Call ShowBills(, True)
    End If
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

