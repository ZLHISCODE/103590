VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTechnoBilling 
   AutoRedraw      =   -1  'True
   Caption         =   "ҽ�����Ҽ���"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9885
   Icon            =   "frmTechnoBilling.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   Picture         =   "frmTechnoBilling.frx":08CA
   ScaleHeight     =   6195
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picHsc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   45
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   9780
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3885
      Width           =   9780
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   5835
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmTechnoBilling.frx":0A58
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8599
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
            Object.Width           =   3722
            MinWidth        =   3722
            Picture         =   "frmTechnoBilling.frx":0DCC
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
      TabIndex        =   4
      Top             =   0
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   9885
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinWidth1       =   6300
      MinHeight1      =   720
      Width1          =   4500
      NewRow1         =   0   'False
      Caption2        =   "ҽ������"
      Child2          =   "cboUnit"
      MinWidth2       =   2100
      MinHeight2      =   300
      Width2          =   1800
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   165
         TabIndex        =   5
         Top             =   30
         Width           =   6525
         _ExtentX        =   11509
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
               Caption         =   "����"
               Key             =   "Billing"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "Billing"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "Modi"
               Description     =   "�޸�"
               Object.ToolTipText     =   "�޸�"
               Object.Tag             =   "�޸�"
               ImageKey        =   "Modi"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Del"
               Description     =   "����"
               Object.ToolTipText     =   "�Ե�ǰѡ�е�������"
               Object.Tag             =   "����"
               ImageKey        =   "Del"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Del_"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "View"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "View"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Filter"
               Description     =   "����"
               Object.ToolTipText     =   "��������������ɸѡ��¼"
               Object.Tag             =   "����"
               ImageKey        =   "Filter"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��λ"
               Key             =   "Go"
               Description     =   "��λ"
               Object.ToolTipText     =   "��λ�����������ļ�¼��"
               Object.Tag             =   "��λ"
               ImageKey        =   "Go"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
      Begin VB.ComboBox cboUnit 
         Height          =   300
         Left            =   7695
         TabIndex        =   2
         Text            =   "cboUnit"
         Top             =   240
         Width           =   2100
      End
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   5205
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":0F6A
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":1184
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":139E
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":15B8
            Key             =   "View"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":1D32
            Key             =   "Go"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":1F4C
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":2166
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":2380
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":259A
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":27B4
            Key             =   "Billing"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":2EAE
            Key             =   "Exe"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":35A8
            Key             =   "Cancel"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   4620
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":3CA2
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":3EBC
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":40D6
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":42F0
            Key             =   "View"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":4A6A
            Key             =   "Go"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":4C84
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":4E9E
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":50B8
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":52D2
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":54EC
            Key             =   "Billing"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":5BE6
            Key             =   "Exe"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":62E0
            Key             =   "Cancel"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img32 
      Left            =   3105
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":69DA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   3690
      Top             =   90
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
            Picture         =   "frmTechnoBilling.frx":72B4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
      Height          =   3135
      Left            =   15
      TabIndex        =   0
      Top             =   735
      Width           =   9870
      _ExtentX        =   17410
      _ExtentY        =   5530
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      MouseIcon       =   "frmTechnoBilling.frx":784E
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   1875
      Left            =   0
      TabIndex        =   1
      Top             =   3960
      Width           =   9870
      _ExtentX        =   17410
      _ExtentY        =   3307
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      MouseIcon       =   "frmTechnoBilling.frx":7B68
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
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
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileLocalSet 
         Caption         =   "��������(&R)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFileLocalSet_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_quit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditBilling 
         Caption         =   "���ʵ�(&B)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditSimple 
         Caption         =   "�򵥼���(&S)"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuEditCust 
         Caption         =   "�Զ������(&U)"
         Begin VB.Menu mnuEditCustBill 
            Caption         =   "(��)"
            Index           =   1
         End
      End
      Begin VB.Menu mnuEditBilling_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditModi 
         Caption         =   "�޸ĵ���(&M)"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuEditAdjust 
         Caption         =   "����ʱ��(&J)"
         Shortcut        =   ^J
      End
      Begin VB.Menu mnuEditAdjust_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDel 
         Caption         =   "��������(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditDelApply 
         Caption         =   "��������(&Q)"
      End
      Begin VB.Menu mnuEditDelAudit 
         Caption         =   "�������(&H)"
      End
      Begin VB.Menu mnuEditDel_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditView 
         Caption         =   "���ĵ���(&V)"
      End
      Begin VB.Menu mnuEditPrint 
         Caption         =   "��ӡ����(&P)"
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
         Begin VB.Menu mnuViewToolUnit 
            Caption         =   "ҽ������(&U)"
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
         Caption         =   "����(&T)"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuViewGo 
         Caption         =   "��λ(&G)"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuView_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefeshOption 
         Caption         =   "ˢ�·�ʽ(&O)"
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "������Ҫˢ������(&1)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "��������ʾ�Ƿ�ˢ��(&2)"
            Index           =   1
         End
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "�������Զ�ˢ������(&3)"
            Index           =   2
         End
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
Attribute VB_Name = "frmTechnoBilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private mrsList As ADODB.Recordset  '�����б�
Private mrsTotal As ADODB.Recordset
Private mrsDetail As ADODB.Recordset  '�����б�
Private mstrFilter As String
Private mbln���� As Boolean, mbln���� As Boolean

Private Type Type_SQLCondition
    Default As Boolean          '�Ƿ���ȱʡ���룬��ʱû������ֵ,ȱʡֵ��mstrFilter��
    DateB As Date
    DateE As Date
    NOB As String
    NOE As String
    InPatientID As Double
    Patient As String
    Operator As String
End Type
Private SQLCondition As Type_SQLCondition

Private mblnGo As Boolean, mlngGo As Long
Private mlngCurRow As Long, mlngTopRow As Long

Private mblnMax As Boolean
Private mlngDeptID As Long, mlngUnitID As Long

Private mstrPrivs As String     '���浱ǰģ�����Ȩ����
Private mstrPrivsOpt As String '���ʲ���1150ģ�����Ȩ����
Private mlngModul As Long
Private mblnNOMoved As Boolean '��¼��ǰѡ��ĵ����Ƿ����ں����ݱ���
Private mrsDept As ADODB.Recordset

Private Sub cboUnit_Click()
    If cboUnit.ItemData(cboUnit.ListIndex) = mlngDeptID Then Exit Sub
    
    mlngDeptID = cboUnit.ItemData(cboUnit.ListIndex)
    mlngUnitID = Get����ID(mlngDeptID)
    
    If Visible Then Call ShowBills(mstrFilter)
End Sub

Private Sub cboUnit_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long, lngҽ��ID As Long
    If KeyAscii <> 13 Then Exit Sub
    
    If cboUnit.ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    
    If mrsDept Is Nothing Then Call InitUnits
    
    
    
    If zlSelectDept(Me, mlngModul, cboUnit, mrsDept, cboUnit.Text, True, "") = False Then
        KeyAscii = 0: Exit Sub
    End If
    Exit Sub


End Sub

Private Sub cboUnit_Validate(Cancel As Boolean)
    If cboUnit.ListIndex >= 0 Then Exit Sub
    zlcontrol.CboLocate cboUnit, mlngDeptID, True
    If cboUnit.ListIndex < 0 And cboUnit.ListCount <> 0 Then cboUnit.ListIndex = 0

End Sub

Private Sub cbr_Resize()
    Form_Resize
End Sub

Private Sub Form_Activate()
    Call InitLocPar(mlngModul)
    Call mshList_GotFocus
End Sub

Private Sub mnuEditAdjust_Click()
    Dim strNO As String
    
    strNO = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    If strNO = "" Then
        MsgBox "��ǰû�е��ݿ��Ե�����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '�Ƿ���ת������ݱ���
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNO, 2, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '��ʱ��ת���������ݱ�
    End If
    
    'δȫ����˻�����˵Ĳ������޸�
    If Not BillIdentical(strNO) Then
        MsgBox "�����а�������δ��˻�ֶ����˵����ݣ��������޸ġ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '�Ѿ�������(����)�ĵ��ݲ��������
    If BillExistDelete(strNO, 2) Then
        MsgBox "�õ��ݰ�������������,�����������", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '�Ƿ��Ѿ�����
    If HaveBilling(2, strNO) <> 0 Then
        Select Case gbytBillOpt
            Case 0
            Case 1
                If MsgBox("�ü��ʵ��ݰ����Ѿ����ʵ�����,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Case 2
                MsgBox "�ü��ʵ��ݰ����Ѿ����ʵ�����,���ܵ�����", vbExclamation, gstrSysName: Exit Sub
        End Select
    End If
    
    On Error Resume Next
    Err.Clear
    
    If BillisSimple(strNO) Then '�򵥼���
        frmSimpleBilling.mstrPrivs = mstrPrivs
        frmSimpleBilling.mbytInState = 2
        frmSimpleBilling.mstrInNO = strNO
        frmSimpleBilling.mbytUseType = 2
        frmSimpleBilling.mlngDeptID = mlngDeptID
        frmSimpleBilling.mlngUnitID = mlngUnitID
        frmSimpleBilling.mlngModule = mlngModul
        frmSimpleBilling.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    Else '���ʵ�
        Dim lng����ID As Long
        Dim varTemp As Variant
        
        lng����ID = mshList.TextMatrix(mshList.Row, GetColNum("���ʵ�ID"))
        
        If lng����ID = 0 Or gobjCustBill Is Nothing Then
            frmCharge.mstrPrivs = mstrPrivs
            frmCharge.mbytInState = 2
            frmCharge.mstrInNO = strNO
            frmCharge.mbytUseType = 2
            frmCharge.mlngDeptID = mlngDeptID
            frmCharge.mlngUnitID = mlngUnitID
            frmCharge.mlngModule = mlngModul
            frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
        Else
            '����ID��bytUseType��bytInState��strInNO��lngUnitID��lngDeptID��lng����ID��mstrPrivs
            varTemp = Array(lng����ID, 2, 2, strNO, mlngUnitID, mlngDeptID, 0, mstrPrivs)
            gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
            
            gblnOK = varTemp
        End If
    End If
End Sub

Private Sub mnuEditBilling_Click()
    Err.Clear
    On Error Resume Next
    
    gbytBilling = 0
    frmCharge.mstrPrivs = mstrPrivs
    frmCharge.mbytInState = 0
    frmCharge.mbytUseType = 2
    frmCharge.mlngDeptID = mlngDeptID
    frmCharge.mlngUnitID = mlngUnitID
    frmCharge.mlngModule = mlngModul
    frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ShowBills(mstrFilter)
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            Call ShowBills(mstrFilter)
        End If
    End If
End Sub

Private Sub mnuEditCustBill_Click(Index As Integer)
    '�Զ������
    Dim varTemp As Variant
            
    '�������������ǣ�
    '����ID��bytUseType��bytInState��strInNO��lngUnitID��lngDeptID��lng����ID��mstrPrivs��blnViewCancel
    varTemp = Array(mnuEditCustBill(Index).Tag, 2, 0, "", mlngUnitID, mlngDeptID, 0, mstrPrivs)
    gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
    
    gblnOK = varTemp '����ֵ
    
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ShowBills(mstrFilter)
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            Call ShowBills(mstrFilter)
        End If
    End If
End Sub

Private Sub mnuEditDelApply_Click()
    If mlngDeptID = 0 Then
        MsgBox "����ѡ��ǰ����!", vbInformation, gstrSysName
        cboUnit.SetFocus
        Exit Sub
    End If
    With frmReCharge
        .mlngDeptID = mlngDeptID
        .mbytUseType = 1
        .mbytFun = 0
        .mstrPrivs = mstrPrivs
        .Show IIf(gfrmMain Is Nothing, 0, 1), Me
    End With
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewReFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewReFlash_Click
        End If
    End If
End Sub

Private Sub mnuEditDelAudit_Click()
    If mlngDeptID = 0 Then
        MsgBox "����ѡ��ǰ����!", vbInformation, gstrSysName
        cboUnit.SetFocus
        Exit Sub
    End If
    With frmReCharge
        .mlngDeptID = mlngDeptID
        .mbytUseType = 1
        .mbytFun = 1
        .mstrPrivs = mstrPrivs
        .Show IIf(gfrmMain Is Nothing, 0, 1), Me
    End With
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewReFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewReFlash_Click
        End If
    End If
End Sub

Private Sub mnuEditPrint_Click()
    Dim strNO As String, strTime As String
    
    strNO = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    
    If strNO = "" Then
        MsgBox "��ǰû�е��ݿ��Դ�ӡ��", vbInformation, gstrSysName
        Exit Sub
    End If

    If mshList.TextMatrix(mshList.Row, GetColNum("����")) <> 1 Then
        MsgBox "�õ���Ϊ���ʵ��ݻ��ѱ����ʣ������ٴ�ӡ��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '�Ƿ���ת������ݱ���
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNO, 2, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '��ʱ��ת���������ݱ�
    End If
    
    strTime = mshList.TextMatrix(mshList.Row, GetColNum("�Ǽ�ʱ��"))
    
    If ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1135", Me) Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1135", Me, "NO=" & strNO, "�Ǽ�ʱ��=" & strTime, "ҩƷ��λ=" & IIf(gblnסԺ��λ, 1, 0), "PrintEmpty=0", "�ش�=1", 2)
    End If
End Sub

Private Sub mnuEditSimple_Click()
    Err.Clear
    On Error Resume Next
    
    gbytBilling = 0
    frmSimpleBilling.mstrPrivs = mstrPrivs
    frmSimpleBilling.mbytInState = 0
    frmSimpleBilling.mbytUseType = 2
    frmSimpleBilling.mlngDeptID = mlngDeptID
    frmSimpleBilling.mlngUnitID = mlngUnitID
    frmSimpleBilling.mlngModule = mlngModul
    frmSimpleBilling.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ShowBills(mstrFilter)
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            Call ShowBills(mstrFilter)
        End If
    End If
End Sub

Private Sub mnuEditModi_Click()
    Dim strNO As String, intInsure As Integer
    Dim strInfo As String, strUnitIDs As String
    
    strNO = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    If strNO = "" Then
        MsgBox "��ǰû�е��ݿ����޸ģ�", vbInformation, gstrSysName
        Exit Sub
    End If

    '�Ƿ���ת������ݱ���
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNO, 2, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '��ʱ��ת���������ݱ�
    End If
    
    'δȫ����˻�����˵Ĳ������޸�
    If Not BillIdentical(strNO) Then
        MsgBox "�����а�������δ��˻�ֶ����˵����ݣ��������޸ġ�", vbInformation, gstrSysName
        Exit Sub
    End If

    'Ȩ���ж�
    If Not BillOperCheck(5, mshList.TextMatrix(mshList.Row, GetColNum("����Ա")), _
        CDate(mshList.TextMatrix(mshList.Row, GetColNum("�Ǽ�ʱ��"))), "�޸�", strNO) Then Exit Sub
    
    'ȫԺ����
    If InStr(mstrPrivsOpt, ";ȫԺ����;") = 0 Then
        If strUnitIDs = "" Then strUnitIDs = GetUserUnits(True)
        
        If InStr("," & strUnitIDs & ",", "," & Val(mshList.TextMatrix(mshList.Row, GetColNum("��������ID"))) & ",") = 0 Then
            MsgBox "��û��Ȩ�޶��������ҵĵ�������,�������޸ĸõ��ݣ�", vbInformation, gstrSysName
            Exit Sub
        End If
    End If

    '���۲���Ȩ��
    strInfo = Check���۲���(strNO, mstrPrivsOpt)
    If strInfo <> "" Then
        MsgBox "�����а���" & strInfo & ",��û��Ȩ�޶Ըõ��ݽ��в�����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '�Ƿ��г�Ժ�޸�(����)Ȩ��
    If Not BillCanBeOperate(strNO, mstrPrivsOpt, "�޸�") Then Exit Sub
    
    'ȥ����ҽ������ƥ����
    
    '����������ʱ�۵�ҩƷ�������޸�
    If Not BillCanModi(strNO, 2) Then
        MsgBox "���ŵ����а���������ʱ��ҩƷ,�������޸ģ�", vbInformation, gstrSysName
        Exit Sub
    End If

    '�Ѿ�������(����)�ĵ��ݲ������޸�
    If BillExistDelete(strNO, 2) Then
        MsgBox "�õ��ݰ��������ʷ���,�������޸ģ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '�����������ִ�л�ȫ��ִ�е���Ŀ,��һ������ȫ������,�������޸�
    If HaveExecute(2, strNO, 2) Then
        MsgBox "�õ����а�����ȫִ�л򲿷�ִ�е���Ŀ,�������޸ģ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '�Ƿ��Ѿ�����
    If HaveBilling(2, strNO) <> 0 Then
        intInsure = BillExistInsure(strNO)
        If intInsure <> 0 Then
            If Not gclsInsure.GetCapability(support��������ѽ��ʵļ��ʵ���, , intInsure) Then
                'ҽ�����˵ĵ��ݹ̶�Ϊ�ѽ��ʾͽ�ֹ�޸�
                MsgBox "��ҽ�����ʵ��ݰ����Ѿ����ʵ�����,�����޸ģ�", vbExclamation, gstrSysName: Exit Sub
            End If
        Else
            Select Case gbytBillOpt
                Case 0
                Case 1
                    If MsgBox("�ü��ʵ��ݰ����Ѿ����ʵ�����,Ҫ�޸���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                Case 2
                    MsgBox "�ü��ʵ��ݰ����Ѿ����ʵ�����,�����޸ģ�", vbExclamation, gstrSysName: Exit Sub
            End Select
        End If
    End If
    
    gstrModiNO = ""
    
    On Error Resume Next
    Err.Clear
    
    gbytBilling = 0 '�����޸�
    If BillisSimple(strNO) Then '�򵥼���
        frmSimpleBilling.mstrPrivs = mstrPrivs
        frmSimpleBilling.mbytInState = 0
        frmSimpleBilling.mstrInNO = strNO
        frmSimpleBilling.mbytUseType = 2
        frmSimpleBilling.mlngDeptID = mlngDeptID
        frmSimpleBilling.mlngUnitID = mlngUnitID
        frmSimpleBilling.mlngModule = mlngModul
        frmSimpleBilling.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    Else '���ʵ�
        Dim lng����ID As Long
        Dim varTemp As Variant
        
        lng����ID = mshList.TextMatrix(mshList.Row, GetColNum("���ʵ�ID"))
        
        If lng����ID = 0 Or gobjCustBill Is Nothing Then
            frmCharge.mstrPrivs = mstrPrivs
            frmCharge.mbytInState = 0
            frmCharge.mstrInNO = strNO
            frmCharge.mbytUseType = 2
            frmCharge.mlngDeptID = mlngDeptID
            frmCharge.mlngUnitID = mlngUnitID
            frmCharge.mlngModule = mlngModul
            frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
        Else
            '����ID��bytUseType��bytInState��strInNO��lngUnitID��lngDeptID��lng����ID��mstrPrivs
            varTemp = Array(lng����ID, 2, 0, strNO, mlngUnitID, mlngDeptID, 0, mstrPrivs)
            gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
            
            gblnOK = varTemp
        End If
    End If

    If gblnOK Then
        If gstrModiNO <> "" Then
            If mnuViewRefeshOptionItem(1).Checked Then
                If MsgBox("��ǰ�����Ѹ��ĵ����嵥����,�޸ĺ�ĵ��ݺ�Ϊ:[" & gstrModiNO & "],Ҫˢ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    Call ShowBills(mstrFilter)
                End If
            ElseIf mnuViewRefeshOptionItem(2).Checked Then
                Call ShowBills(mstrFilter)
            End If
        Else
            If mnuViewRefeshOptionItem(1).Checked Then
                If MsgBox("��ǰ�����Ѹ��ĵ����嵥����,Ҫˢ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    Call ShowBills(mstrFilter)
                End If
            ElseIf mnuViewRefeshOptionItem(2).Checked Then
                Call ShowBills(mstrFilter)
            End If
        End If
    End If
End Sub

Private Sub mnuFileLocalSet_Click()
    Dim blnסԺ��λ As Boolean
    
    blnסԺ��λ = gblnסԺ��λ
    
    frmSetExpence.mlngModul = mlngModul
    frmSetExpence.mstrPrivs = mstrPrivs
    frmSetExpence.mbytInFun = 0
    frmSetExpence.mbytUseType = 2
    frmSetExpence.Show 1, Me
    If gblnOK Then
        
        If blnסԺ��λ <> gblnסԺ��λ Then
            If Not (mshList.Rows = 2 And mshList.TextMatrix(1, GetColNum("���ݺ�")) = "") Then
                Call mnuViewReFlash_Click
            End If
        End If
    End If
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim strNO As String
    
    strNO = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    If strNO = "" Then
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
            "����=" & mlngUnitID, "���˿���=" & mlngDeptID)
    Else
        With mshList
            Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
                "����=" & mlngUnitID, "���˿���=" & mlngDeptID, "NO=" & strNO, _
                "סԺ��=" & .TextMatrix(.Row, GetColNum("סԺ��")), _
                "������=" & .TextMatrix(.Row, GetColNum("������")))
        End With
    End If
End Sub

Private Sub mnuViewFilter_Click()
    
    If frmTechnoFilter.mlngDept <> mlngDeptID Then
        frmTechnoFilter.mlngDept = mlngDeptID
        frmTechnoFilter.LoadOper
    End If
    
    frmTechnoFilter.Show 1, Me
    If gblnOK Then
        With frmTechnoFilter
            mstrFilter = .mstrFilter
            mbln���� = .chk����.Value = 1
            mbln���� = .chk����.Value = 1
            
            SQLCondition.Default = False
            SQLCondition.DateB = .dtpBegin.Value
            SQLCondition.DateE = .dtpEnd.Value
            SQLCondition.NOB = .txtNOBegin.Text
            SQLCondition.NOE = .txtNoEnd.Text
            SQLCondition.InPatientID = Val(.txtסԺ��.Text)
            SQLCondition.Patient = gstrLike & UCase(.txt����.Text) & "%"
            SQLCondition.Operator = zlStr.NeedName(.cbo����Ա.Text)
        End With
        
        mnuViewReFlash_Click
    End If
End Sub

Private Sub mshDetail_EnterCell()
    mshDetail.ForeColorSel = mshDetail.CellForeColor
End Sub

Private Sub mshDetail_GotFocus()
    Call SetActiveList(mshDetail)
End Sub

Private Sub mshDetail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mshDetail.MouseRow = 0 Then
        mshDetail.MousePointer = 99
    Else
        mshDetail.MousePointer = 0
    End If
End Sub

Private Sub mshDetail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long, strTime As String, blnDel As Boolean
    
    lngCol = mshDetail.MouseCol
    
    If Button = 1 And mshDetail.MousePointer = 99 Then
        If mshDetail.TextMatrix(0, lngCol) = "" Then Exit Sub
        If mshDetail.TextMatrix(1, 0) = "" Then Exit Sub
        If mrsDetail Is Nothing Then Exit Sub
        
        strTime = mshList.TextMatrix(mshList.Row, GetColNum("�Ǽ�ʱ��"))
        blnDel = Val(mshList.TextMatrix(mshList.Row, GetColNum("����"))) = 2

        Set mshDetail.DataSource = Nothing

        mrsDetail.Sort = mshDetail.TextMatrix(0, lngCol) & IIf(mshDetail.ColData(lngCol) = 0, "", " DESC")
        mshDetail.ColData(lngCol) = (mshDetail.ColData(lngCol) + 1) Mod 2
        
        Call ShowDetail(, strTime, blnDel, True)
    End If
End Sub

Private Sub mshList_DblClick()
    If mshList.MouseRow = 0 Then Exit Sub
    If mnuEditView.Enabled Then mnuEditView_Click
End Sub

Private Sub mshList_EnterCell()
    Dim strNO As String, strTime As String, blnDel As Boolean
        
    strNO = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    
    If mshList.Row = 0 Or strNO = "" Then Exit Sub
    
    stbThis.Panels(2).Text = "�� " & Nvl(mrsTotal!����, 0) & " �ŵ���,�ϼ�:" & Format(Nvl(mrsTotal!���, 0), gstrDec)
    
    mlngGo = mshList.Row
    mlngCurRow = mshList.Row: mlngTopRow = mshList.TopRow
    
    strTime = mshList.TextMatrix(mshList.Row, GetColNum("�Ǽ�ʱ��"))
    blnDel = Val(mshList.TextMatrix(mshList.Row, GetColNum("����"))) = 2
    
    mnuEditAdjust.Enabled = Not blnDel
    mnuEditModi.Enabled = Not blnDel And Val(mshList.TextMatrix(mshList.Row, GetColNum("ҽ�����"))) = 0 _
                        And Val(mshList.TextMatrix(mshList.Row, GetColNum("��¼����"))) <> 3
    mnuEditDel.Enabled = Not blnDel
    tbr.Buttons("Modi").Enabled = mnuEditModi.Enabled
    tbr.Buttons("Del").Enabled = mnuEditDel.Enabled
        
    mshList.ForeColorSel = mshList.CellForeColor
    
    Call ShowDetail(strNO, strTime, blnDel)
End Sub

Private Sub mshList_GotFocus()
    Call SetActiveList(mshList)
End Sub

Private Sub mshList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And mnuEditDel.Enabled And mnuEditDel.Visible Then Call mnuEditDel_Click
End Sub

Private Sub mshList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuEdit, 2
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            'ʼ�մӵ�ǰ�п�ʼ
            If mnuViewGo.Enabled Then Call SeekBill(False)
        Case vbKeyReturn
            If Me.ActiveControl Is cboUnit Then
            Else
                If mnuEditView.Enabled Then mnuEditView_Click
            End If
        Case vbKeyEscape
            mblnGo = False
    End Select
End Sub

Private Sub mnuEditDel_Click()
    Dim intInsure As Integer, strInfo As String
    Dim strNO As String, strTime As String
    Dim intTmp As Integer, i As Long, blnFlagPrint As Boolean
    
    strNO = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    If strNO = "" Then
        MsgBox "��ǰû�е��ݿ������ʣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    strTime = mshList.TextMatrix(mshList.Row, GetColNum("�Ǽ�ʱ��"))
    
    'Ȩ���ж�
    If Not BillOperCheck(5, mshList.TextMatrix(mshList.Row, GetColNum("����Ա")), CDate(strTime), "����", strNO) Then Exit Sub

    'δȫ����˻�����˵Ĳ������޸�
    If Not BillIdentical(strNO) Then
        MsgBox "�����а�������δ��˻�ֶ����˵����ݣ��������޸ġ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '�Ƿ���ת������ݱ���
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNO, 2, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '��ʱ��ת���������ݱ�
    End If
    
    '���۲���Ȩ��
    strInfo = Check���۲���(strNO, mstrPrivsOpt, strTime)
    If strInfo <> "" Then
        MsgBox "�����а���" & strInfo & ",��û��Ȩ�޶Ըõ��ݽ��в�����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '�Ƿ���ִ��
    i = BillCanDelete(strNO, 2, , strTime, mstrPrivsOpt, blnFlagPrint)
    If i <> 0 Then
        Select Case i
            Case 1 '�õ��ݲ�����
                MsgBox "ָ�������е����ݲ����ڣ�", vbInformation, gstrSysName
            Case 2 '�Ѿ�ȫ����ȫִ��
                MsgBox "ָ�������е������Ѿ�ȫ����ȫִ�У�", vbInformation, gstrSysName
            Case 3 'δ��ȫִ�в���ʣ������Ϊ0
                MsgBox "ָ�������е�����δִ�в�����Ŀʣ������Ϊ��,û�п������ʵķ��ã�", vbInformation, gstrSysName
        End Select
        Exit Sub
    End If
    If blnFlagPrint Then
        If MsgBox("ע��:����ҽ���������Ѵ�ӡ���Ƿ������", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    
    '�Ƿ��г�Ժ�޸�(����)Ȩ��
    If Not BillCanBeOperate(strNO, mstrPrivsOpt, "����", strTime) Then Exit Sub
    
    '�Ƿ��Ѿ�����
    intInsure = BillExistInsure(strNO)
    intTmp = HaveBilling(2, strNO, False, strTime)
    If intTmp <> 0 Then
        If intInsure <> 0 Then
            If Not gclsInsure.GetCapability(support��������ѽ��ʵļ��ʵ���, , intInsure) Then
                'ҽ�����˵ĵ���,�̶�Ϊ�ѽ��ʵĽ�ֹ����
                If intTmp = 1 Then
                    MsgBox "��ҽ�����ʵ���δ���ʲ����Ѿ�����,�������ʣ�", vbExclamation, gstrSysName
                    Exit Sub
                Else
                    MsgBox "��ҽ�����ʵ��ݰ����Ѿ����ʵ�����,ֻ�ܶ�δ���ʲ��ֽ������ʣ�", vbExclamation, gstrSysName
                End If
            End If
        Else
            Select Case gbytBillOpt
                Case 0
                Case 1
                    If MsgBox("�ü��ʵ��ݰ����Ѿ����ʵ�����,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                Case 2
                    If intTmp = 1 Then
                        MsgBox "�ü��ʵ���δ���ʲ����Ѿ�����,�������ʣ�", vbExclamation, gstrSysName
                        Exit Sub
                    Else
                        MsgBox "�ü��ʵ��ݰ����Ѿ����ʵ�����,ֻ�ܶ�δ���ʲ��ֽ������ʣ�", vbExclamation, gstrSysName: Exit Sub
                    End If
            End Select
        End If
    End If
    
    'ҽ�����ʲ�����Ը�����¼��������
    If intInsure <> 0 Then
        If CheckNONegative(strNO) Then
            MsgBox "�õ��ݴ��ڸ������ʼ�¼,���������ҽ�����ʲ�����", vbInformation, gstrSysName
             Exit Sub
        End If
    End If
        
    '�Ƿ������������¼
    If CheckRecalcRecord(strNO) Then
        MsgBox "���ָü��ʵ��ݴ��ڰ��ѱ�����Ĵ��۳����¼!" & vbCrLf & _
            "����ǰ�밴�ѱ�������ã������˽����������ʵ��ݵĴ����Żݽ�", vbInformation, Me.Caption
    End If
    
    On Error Resume Next
    Err.Clear
    
    If BillisSimple(strNO) Then '�򵥼���
        frmSimpleBilling.mbytUseType = 2
        frmSimpleBilling.mstrPrivs = mstrPrivs
        frmSimpleBilling.mbytInState = 3
        frmSimpleBilling.mstrInNO = strNO
        frmSimpleBilling.mstrTime = strTime
        frmSimpleBilling.mlngDeptID = mlngDeptID
        frmSimpleBilling.mlngUnitID = mlngUnitID
        frmSimpleBilling.mlngModule = mlngModul
        frmSimpleBilling.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    Else '���ʵ�
        Dim lng����ID As Long, varTemp As Variant
        
        lng����ID = mshList.TextMatrix(mshList.Row, GetColNum("���ʵ�ID"))
        
        If lng����ID = 0 Or gobjCustBill Is Nothing Then
            frmCharge.mbytUseType = 2
            frmCharge.mstrPrivs = mstrPrivs
            frmCharge.mbytInState = 3
            frmCharge.mstrInNO = strNO
            frmCharge.mstrTime = strTime
            frmCharge.mlngDeptID = mlngDeptID
            frmCharge.mlngUnitID = mlngUnitID
            frmCharge.mlngModule = mlngModul
            frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
        Else
            '����ID��bytUseType��bytInState��strInNO��lngUnitID��lngDeptID��lng����ID��mstrPrivs
            varTemp = Array(lng����ID, 2, 3, strNO, 0, 0, 0, mstrPrivs)
            gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
            
            gblnOK = varTemp
        End If
    End If

    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����Ѹ��ĵ����嵥����,Ҫˢ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ShowBills(mstrFilter)
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            Call ShowBills(mstrFilter)
        End If
    End If
End Sub

Private Sub mnuHelpTitle_Click()
ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub mnuEditView_Click()
    Dim strNO As String, strTime As String, blnDel As Boolean
    
    strNO = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    If strNO = "" Then
        MsgBox "��ǰû�е��ݿ��Բ��ģ�", vbInformation, gstrSysName
        Exit Sub
    End If

    strTime = mshList.TextMatrix(mshList.Row, GetColNum("�Ǽ�ʱ��"))
    blnDel = Val(mshList.TextMatrix(mshList.Row, GetColNum("����"))) = 2

    On Error Resume Next
    Err.Clear
    
    If BillisSimple(strNO) Then '�򵥼���
        frmSimpleBilling.mbytUseType = 2
        frmSimpleBilling.mlngDeptID = mlngDeptID
        frmSimpleBilling.mstrPrivs = mstrPrivs
        frmSimpleBilling.mbytInState = 1
        frmSimpleBilling.mstrInNO = strNO
        frmSimpleBilling.mblnNOMoved = mblnNOMoved
        frmSimpleBilling.mstrTime = strTime
        frmSimpleBilling.mblnDelete = blnDel
        frmSimpleBilling.mlngModule = mlngModul
        frmSimpleBilling.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    Else '���ʵ�
        Dim lng����ID As Long
        Dim varTemp As Variant
        
        lng����ID = mshList.TextMatrix(mshList.Row, GetColNum("���ʵ�ID"))
        
        If lng����ID = 0 Or gobjCustBill Is Nothing Then
            frmCharge.mlngDeptID = mlngDeptID
            frmCharge.mbytUseType = 2
            frmCharge.mstrPrivs = mstrPrivs
            frmCharge.mbytInState = 1
            frmCharge.mstrInNO = strNO
            frmCharge.mblnNOMoved = mblnNOMoved
            frmCharge.mstrTime = strTime
            frmCharge.mblnDelete = blnDel
            frmCharge.mlngModule = mlngModul
            frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
        Else
            '����ID��bytUseType��bytInState��strInNO��lngUnitID��lngDeptID��lng����ID��mstrPrivs
            varTemp = Array(lng����ID, 2, 1, strNO, 0, mlngDeptID, 0, mstrPrivs, blnDel)
            gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
            
            gblnOK = varTemp
        End If
    End If
End Sub

Private Sub mnuFile_quit_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuViewReFlash_Click()
    Call ShowBills(mstrFilter)
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = Not stbThis.Visible
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim i As Long
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For i = 1 To tbr.Buttons.Count
        tbr.Buttons(i).Caption = IIf(mnuViewToolText.Checked, tbr.Buttons(i).Tag, "")
    Next
    cbr.Bands(1).minHeight = tbr.ButtonHeight
    Form_Resize
End Sub

Private Sub mnuViewToolUnit_Click()
    mnuViewToolUnit.Checked = Not mnuViewToolUnit.Checked
    If mnuViewToolButton.Checked Then cbr.Bands(1).Visible = False
    cbr.Bands(2).Visible = Not cbr.Bands(2).Visible
    If mnuViewToolButton.Checked Then cbr.Bands(1).Visible = True
    cbr.Visible = cbr.Bands(2).Visible Or cbr.Bands(1).Visible
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    cbr.Bands(1).Visible = Not cbr.Bands(1).Visible
    mnuViewToolText.Enabled = Not mnuViewToolText.Enabled
    cbr.Visible = cbr.Bands(2).Visible Or cbr.Bands(1).Visible
    Form_Resize
End Sub

Private Sub picHsc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If mshList.Height + Y < 1000 Or mshDetail.Height - Y < 1000 Then Exit Sub
        picHsc.Top = picHsc.Top + Y
        mshList.Height = mshList.Height + Y
        mshDetail.Top = mshDetail.Top + Y
        mshDetail.Height = mshDetail.Height - Y
        Me.Refresh
    End If
End Sub

Private Sub picHsc_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then mshList.SetFocus
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuFile_quit_Click
        Case "Go" '��λ
            mnuViewGo_Click
        Case "Filter" '����
            mnuViewFilter_Click
        Case "View"
            mnuEditView_Click
        Case "Billing"
            mnuEditBilling_Click
        Case "Modi"
            mnuEditModi_Click
        Case "Del"
            mnuEditDel_Click
        Case "Print"
            mnuFile_Print_Click
        Case "Preview"
            mnuFile_PreView_Click
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
    objOut.Title.Text = "סԺ���ʵ����嵥"
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '����
    With frmTechnoFilter
        objRow.Add "ʱ�䣺" & Format(.dtpBegin.Value, .dtpBegin.CustomFormat) & " �� " & Format(.dtpEnd.Value, .dtpEnd.CustomFormat)
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

Private Sub SetMenu(blnUsed As Boolean)
'���ܣ��������޼�¼���ò˵�����״̬
    mnuFile_Print.Enabled = blnUsed
    mnuFile_PreView.Enabled = blnUsed
    mnuFile_Excel.Enabled = blnUsed
    tbr.Buttons("Print").Enabled = blnUsed
    tbr.Buttons("Preview").Enabled = blnUsed
    
    mnuEditAdjust.Enabled = blnUsed
    mnuEditModi.Enabled = blnUsed
    tbr.Buttons("Modi").Enabled = blnUsed
    
    mnuEditDel.Enabled = blnUsed
    mnuEditView.Enabled = blnUsed
    mnuEditPrint.Enabled = blnUsed
    tbr.Buttons("Del").Enabled = blnUsed
    tbr.Buttons("View").Enabled = blnUsed
    
    mnuViewGo.Enabled = blnUsed
    tbr.Buttons("Go").Enabled = blnUsed
End Sub

Private Sub SetCustBill()
'�������Զ�����ʵ���ص�����
    Dim rsTmp As New ADODB.Recordset
    Dim lngCount As Long, lngSum As Long
    On Error Resume Next
    
    If gobjCustBill Is Nothing Then
        Set gobjCustBill = CreateObject("zl9CustAcc.clsCustAcc")
    End If
    If InStr(mstrPrivsOpt, ";ר�����;") = 0 Then
        mnuEditCust.Visible = False
        Exit Sub
    End If
    
    On Error GoTo errHandle
    
    '��������ɹ����ٶ�����Ӧ�Ĳ˵�
    If Not gobjCustBill Is Nothing Then
        gstrSQL = "Select ID,���� From �շѼ��ʵ� Where substr(���÷�Χ,4,1)='1' Order by ���"
        Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
        lngSum = rsTmp.RecordCount
    End If
    
    If lngSum > 0 Then
        For lngCount = 1 To lngSum
            '���ӵ����˵���
            If lngCount > 1 Then
                Load mnuEditCustBill(lngCount)
            End If
            mnuEditCustBill(lngCount).Caption = rsTmp("����") & "(&" & lngCount & ")"
            mnuEditCustBill(lngCount).Tag = rsTmp("ID")
            
            rsTmp.MoveNext
        Next
    Else
        mnuEditCustBill(1).Enabled = False
    End If
    

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Load()
    Dim i As Long
    
    mstrPrivs = gstrPrivs
    mstrPrivsOpt = GetInsidePrivs(Enum_Inside_Program.p���ʲ���)
    mlngModul = glngModul
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    
    Call SetCustBill
    Call RestoreWinState(Me, App.ProductName)
    Set stbThis.Panels(5).Picture = Me.Picture
    
    'ˢ�·�ʽ
    For i = 0 To mnuViewRefeshOptionItem.UBound
        If i = Val(zlDatabase.GetPara("ˢ�·�ʽ", glngSys, mlngModul, 2)) Then
            mnuViewRefeshOptionItem(i).Checked = True
        Else
            mnuViewRefeshOptionItem(i).Checked = False
        End If
    Next
    
    mlngCurRow = 1: mlngTopRow = 1
    
    'Ȩ������
    If InStr(mstrPrivsOpt, ";סԺ����;") = 0 Then
        mnuEditBilling.Visible = False
        mnuEditSimple.Visible = False
        mnuEditBilling_.Visible = False
        tbr.Buttons("Billing").Visible = False
        
        mnuEditCust.Visible = False
    End If
    '55380
    If InStr(mstrPrivsOpt, ";ҩƷ����;") = 0 _
        And InStr(mstrPrivsOpt, ";��������;") = 0 _
        And InStr(mstrPrivsOpt, ";��������;") = 0 Then
        mnuEditDel.Visible = False
        If InStr(mstrPrivsOpt, ";ҩƷ��������;") = 0 _
            And InStr(mstrPrivsOpt, ";������������;") = 0 _
            And InStr(mstrPrivsOpt, ";������������;") = 0 _
            And InStr(mstrPrivsOpt, ";�������;") = 0 Then
            mnuEditDel_.Visible = False
        End If
        tbr.Buttons("Del").Visible = False
    End If
    '55380
    If InStr(mstrPrivsOpt, ";ҩƷ��������;") = 0 _
        Or InStr(mstrPrivsOpt, ";������������;") = 0 _
        Or InStr(mstrPrivsOpt, ";������������;") = 0 _
        Or InStr(1, mstrPrivsOpt, ";��������;") = 0 Then
        mnuEditDelApply.Visible = False
    End If
    
    If InStr(mstrPrivsOpt, ";�������;") = 0 Then
        mnuEditDelAudit.Visible = False
    End If
    
    If InStr(mstrPrivsOpt, ";��¼�޸�;") = 0 Then
        mnuEditModi.Visible = False
        tbr.Buttons("Modi").Visible = False
    End If
    If InStr(mstrPrivsOpt, ";��¼����;") = 0 Then
        mnuEditAdjust.Visible = False
    End If
    If InStr(mstrPrivsOpt, ";��¼�޸�;") = 0 _
        And InStr(mstrPrivsOpt, ";��¼����;") = 0 Then
        mnuEditAdjust_.Visible = False
    End If
    '55380
    If InStr(mstrPrivsOpt, ";סԺ����;") = 0 _
        And InStr(mstrPrivsOpt, ";��¼�޸�;") = 0 _
        And (InStr(mstrPrivsOpt, ";ҩƷ����;") = 0 _
        And InStr(mstrPrivsOpt, ";��������;") = 0 _
        And InStr(mstrPrivsOpt, ";��������;") = 0) Then
        tbr.Buttons("Del_").Visible = False
    End If
        
    '����
    If Not InitUnits Then Unload Me: Exit Sub
    If cboUnit.ListIndex = -1 Then
        MsgBox "û�з�������������,���㲻�������п���Ȩ��,����ʹ��ҽ�����Ҽ��ʣ�", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    
    mbln���� = True
    mbln���� = False
    
    Call SetHeader
    Call SetDetail
    Call SetMenu(False)
    
    stbThis.Panels(2).Text = "��ˢ���嵥���������ù�������"
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long, staH As Long, sngVsc As Single

    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    
    mshList.MousePointer = 0
    
    '����ؼ���Ⱥ͸߶�
    cbrH = IIf(cbr.Visible, cbr.Height, 0)
    staH = IIf(stbThis.Visible, stbThis.Height, 0)
    sngVsc = mshDetail.Height / (mshDetail.Height + mshList.Height)
    
    If mblnMax Then
        sngVsc = 0.3: mblnMax = False
    End If
    If Me.WindowState = 2 Then mblnMax = True
    
    mshList.Left = 0
    mshList.Top = cbrH
    mshList.Width = Me.ScaleWidth
    mshList.Height = (Me.ScaleHeight - cbrH - staH - picHsc.Height) * (1 - sngVsc)
    
    picHsc.Left = Me.ScaleLeft
    picHsc.Top = mshList.Top + mshList.Height
    picHsc.Width = Me.ScaleWidth
    
    mshDetail.Left = Me.ScaleLeft
    mshDetail.Top = picHsc.Top + picHsc.Height
    mshDetail.Width = Me.ScaleWidth
    mshDetail.Height = Me.ScaleHeight - staH - cbrH - picHsc.Height - mshList.Height
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    
    mstrFilter = ""
    mlngDeptID = 0
    mlngUnitID = 0
    
    Unload frmTechnoFilter
    Unload frmTechnoGo
    Call SaveWinState(Me, App.ProductName)
    
    'ˢ�·�ʽ
    For i = 0 To mnuViewRefeshOptionItem.UBound
        If mnuViewRefeshOptionItem(i).Checked Then
            zlDatabase.SetPara "ˢ�·�ʽ", i, glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
            Exit For
        End If
    Next
    
End Sub

Private Sub mnuViewGo_Click()
    frmTechnoGo.Show 1, Me
    If gblnOK Then Call SeekBill(frmTechnoGo.optHead)
End Sub

Private Sub SeekBill(blnHead As Boolean)
    Dim i As Long, bln As Boolean, intRows As Integer
    Dim blnFill As Boolean, j As Long
    
    Screen.MousePointer = 11
    mblnGo = True
    stbThis.Panels(2).Text = "���ڶ�λ���������ĵ���,��ESC��ֹ ..."
    Me.Refresh
    
    For i = IIf(blnHead, 1, mlngGo) To mshList.Rows - 1
        DoEvents

        '�Ƚ�����
        blnFill = True
        With frmTechnoGo
            If .txtNO.Text <> "" Then
                blnFill = blnFill And mshList.TextMatrix(i, GetColNum("���ݺ�")) = .txtNO.Text
            End If
            If .txtסԺ��.Text <> "" Then
                blnFill = blnFill And mshList.TextMatrix(i, GetColNum("סԺ��")) = .txtסԺ��.Text
            End If
            If .txt����ID.Text <> "" Then
                blnFill = blnFill And mshList.TextMatrix(i, GetColNum("����ID")) = .txt����ID.Text
            End If
            If .txt����.Text <> "" Then
                blnFill = blnFill And UCase(mshList.TextMatrix(i, GetColNum("����"))) Like "*" & UCase(.txt����.Text) & "*"
            End If
        End With
        
        '�������˳�
        If blnFill Then
            mshList.Row = i: mshList.TopRow = i
            mshList.Col = 0: mshList.ColSel = mshList.Cols - 1
            
            Call mshList_EnterCell
            mlngGo = i + 1
            
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

Private Function GetColNum(strHead As String) As Integer
    Dim i As Long
    For i = 0 To mshList.Cols - 1
        If mshList.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
    Next
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
        If mshList.TextMatrix(1, GetColNum("���ݺ�")) = "" Then Exit Sub
        If mrsList Is Nothing Then Exit Sub
        
        Set mshList.DataSource = Nothing

        mrsList.Sort = mshList.TextMatrix(0, lngCol) & IIf(mshList.ColData(lngCol) = 0, "", " DESC")
        mshList.ColData(lngCol) = (mshList.ColData(lngCol) + 1) Mod 2
        
        Call ShowBills(, True)
    End If
End Sub

Private Sub SetHeader()
    Dim strHead As String
    Dim i As Long
    
    strHead = "���ݺ�,1,850|������,1,800|���˿���,1,850|סԺ��,1,750|����,1,500|����,1,700|�ѱ�,1,900|Ӧ�ս��,7,850|ʵ�ս��,7,850|" & _
            "����Ա,1,800|�Ǽ�ʱ��,1,1850|˵��,1,850|����,1,0|��¼����,1,0|�ಡ�˵�,1,0|���ʵ�ID,1,0|����ID,1,0|��ҳID,1,0|ҽ�����,1,0|��������ID,1,0"
    With mshList
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(mshList, App.ProductName & "\" & Me.Name)
        .RowHeight(0) = 320
        
        i = GetColNum("����"): mshList.ColWidth(i) = 0
        i = GetColNum("��¼����"): mshList.ColWidth(i) = 0
        i = GetColNum("�ಡ�˵�"): mshList.ColWidth(i) = 0
        i = GetColNum("���ʵ�ID"): mshList.ColWidth(i) = 0
        
        '�鿴ҽ����Ȩ��
        i = GetColNum("������")
        If InStr(mstrPrivsOpt, ";ҽ����ѯ;") = 0 Then
            mshList.ColWidth(i) = 0
        ElseIf mshList.ColWidth(i) = 0 Then
            mshList.ColWidth(i) = 800
        End If
        
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
    End With
End Sub

Private Sub ShowBills(Optional ByVal strIF As String, Optional blnSort As Boolean)
'����:��������ȡ�����б�(���˹���)
'����:strIF=��"AND"��ʼ��������
'     blnSort=�����¶�ȡ����,��������ʾ�����������
    Dim i As Long, j As Long, k As Long
    Dim Curdate As Date, strSql As String
    
    On Error GoTo errH
        
    If Not blnSort Then
        Call zlCommFun.ShowFlash("���ڶ�ȡ�����б�,���Ժ� ...", Me)
        DoEvents
        Me.Refresh
        
        'ȡȱʡ����(���ռ���)
        SQLCondition.Default = (strIF = "")
        If strIF = "" Then
            strIF = " And �Ǽ�ʱ�� Between trunc(sysdate) And trunc(sysdate+1)-1/24/60/60 And ��¼״̬ IN(1,3)"
            strIF = strIF & " And ����Ա����||''=[7]"
        End If
        strIF = strIF & " And ��������ID+0=[8]"
        strIF = " Where ��¼����=2 And �����־=2 And ��¼״̬<>0 And ����Ա���� is Not NULL And Nvl(�ಡ�˵�,0)=0  " & strIF
        
        'ɸѡʱ��ʱ�������һ��ת��֮ǰ
        If frmTechnoFilter.mblnDateMoved Then
            strIF = zlGetFullFieldsTable("סԺ���ü�¼", 2, strIF, False)
        Else
            strIF = zlGetFullFieldsTable("סԺ���ü�¼", 0, strIF, False)
        End If
                
        '���ݺ�,������,���˿���,סԺ��,����,����,�ѱ�,Ӧ�ս��,ʵ�ս��,����Ա,�Ǽ�ʱ��,˵��,����,��¼����,�ಡ�˵�,���ʵ�ID
        strSql = _
        " Select A.NO as ���ݺ�,A.������," & _
        "        B.���� as ���˿���,A.��ʶ�� as סԺ��,C.��Ժ���� as ����,A.����,A.�ѱ�," & _
        "        To_Char(Sum(Decode(A.��¼״̬,2,-1,1)*A.Ӧ�ս��),'9999999" & gstrDec & "') as Ӧ�ս��," & _
        "        To_Char(Sum(Decode(A.��¼״̬,2,-1,1)*A.ʵ�ս��),'9999999" & gstrDec & "') as ʵ�ս��," & _
        "        A.����Ա���� as ����Ա,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as �Ǽ�ʱ��," & _
        "        Decode(A.��¼����,3,Decode(Max(A.��¼״̬),2,'�Զ�����','�Զ�����'),Decode(Max(A.��¼״̬),2,'���ʼ�¼','���ʼ�¼')) as ˵��," & _
        "        Max(A.��¼״̬) as ����,A.��¼����,A.�ಡ�˵�,A.���ʵ�ID,A.����ID,A.��ҳID,Nvl(A.ҽ�����,0) ҽ�����,A.��������ID" & _
        " From (" & strIF & ") A,���ű� B,������ҳ C" & _
        " Where A.���˿���ID=B.ID(+) And A.����ID=C.����ID(+) And A.��ҳID=C.��ҳID (+)" & _
        " Group by A.NO,A.������,B.����,A.��ʶ��,C.��Ժ����,A.����,A.�ѱ�,A.����Ա����," & _
        "          A.�Ǽ�ʱ��,A.��¼����,A.�ಡ�˵�,A.���ʵ�ID,A.����ID,A.��ҳID,Nvl(A.ҽ�����,0),A.��������ID" & _
        " Order by A.�Ǽ�ʱ�� Desc,A.NO Desc"
        
        With SQLCondition
            If .Default Then .Operator = UserInfo.����
            Set mrsList = zlDatabase.OpenSQLRecord(strSql, Me.Caption, .DateB, .DateE, .NOB, .NOE, .InPatientID, .Patient, .Operator, mlngDeptID)
        End With
    End If
    
    mshList.Clear
    mshList.Rows = 2
    
    mshDetail.Clear
    mshDetail.Rows = 2
    
    If mrsList.EOF Then
        stbThis.Panels(2).Text = "��ǰ����û�й��˳��κε���"
        Call SetMenu(False)
    Else
        '��ʵ�պϼƽ��
        If Not blnSort Then
            strSql = "Select Sum(ʵ�ս��) as ���,Count(Distinct NO) as ���� From (" & Replace(strIF, "��¼״̬ IN(1,3)", "��¼״̬ IN(1,2,3)") & ")"
            With SQLCondition
                Set mrsTotal = zlDatabase.OpenSQLRecord(strSql, Me.Caption, .DateB, .DateE, .NOB, .NOE, .InPatientID, .Patient, .Operator, mlngDeptID)
            End With
        End If
    
        Set mshList.DataSource = mrsList
        stbThis.Panels(2).Text = "�� " & Nvl(mrsTotal!����, 0) & " �ŵ���,�ϼ�:" & Format(Nvl(mrsTotal!���, 0), gstrDec)
        Call SetMenu(True)
    End If

    mshList.Redraw = False
    '������ɫ
    If mbln���� And Not mbln���� Then
        mshList.ForeColor = &HC0
    Else
        mshList.ForeColor = ForeColor
        k = GetColNum("����")
        For i = 1 To mshList.Rows - 1
            If Val(mshList.TextMatrix(i, k)) = 2 Then
                '���ʼ�¼�ú�ɫ
                mshList.Row = i
                For j = 0 To mshList.Cols - 1
                    mshList.Col = j
                    mshList.CellForeColor = &HC0
                Next
            ElseIf Val(mshList.TextMatrix(i, k)) = 3 Then
                '�������ʵ�����ɫ
                mshList.Row = i
                For j = 0 To mshList.Cols - 1
                    mshList.Col = j
                    mshList.CellForeColor = &HC00000
                Next
            End If
        Next
    End If
    
    Call SetHeader
    If mshList.Row = 0 Or mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�")) = "" Then Call SetDetail
        
    mshList.Redraw = True
    
    If Not blnSort Then Call zlCommFun.StopFlash
    
    Me.Refresh
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function InitUnits() As Boolean
'���ܣ���ʼ��ҽ������
    Dim i As Long, strSql As String
    
    On Error GoTo errH
        
    '��������/סԺҽ������
    If InStr(mstrPrivs, ";���п���;") > 0 Then
        strSql = _
            " Select Distinct A.ID,A.����,A.����,A.����" & _
            " From ���ű� A,��������˵�� B " & _
            " Where B.����ID = A.ID " & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
            " And B.������� IN(1,2,3) And B.�������� IN('���','����','����','����','Ӫ��')" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
            " Order by A.����"
    Else
        strSql = _
            " Select Distinct A.ID,A.����,A.����,A.����,Nvl(C.ȱʡ,0) as ȱʡ" & _
            " From ���ű� A,��������˵�� B,������Ա C " & _
            " Where B.����ID = A.ID And A.ID=C.����ID And C.��ԱID=[1]" & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
            " And B.������� IN(1,2,3) And B.�������� IN('���','����','����','����','Ӫ��')" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
            " Order by A.����"
    End If
    Set mrsDept = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID)
    
    If Not mrsDept.EOF Then
        For i = 1 To mrsDept.RecordCount
            cboUnit.AddItem mrsDept!���� & "-" & mrsDept!����
            cboUnit.ItemData(cboUnit.NewIndex) = mrsDept!ID
            If cboUnit.ListIndex = -1 Then
                If InStr(mstrPrivs, ";���п���;") > 0 Then
                    If UserInfo.����ID = mrsDept!ID Then cboUnit.ListIndex = cboUnit.NewIndex
                Else
                    If mrsDept!ȱʡ = 1 Then cboUnit.ListIndex = cboUnit.NewIndex
                End If
            End If
            mrsDept.MoveNext
        Next
        If cboUnit.ListIndex = -1 And cboUnit.ListCount > 0 Then cboUnit.ListIndex = 0
    ElseIf InStr(mstrPrivs, ";���п���;") > 0 Then
        MsgBox "û�п��õ�ҽ������,���ȵ����Ź��������ã�", vbInformation, gstrSysName
        Exit Function
    End If
    InitUnits = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetActiveList(obj As Object)
    If obj Is mshList Then
        mshList.BackColorSel = &HC0C0C0
        mshDetail.BackColorSel = &HE0E0E0
    ElseIf obj Is mshDetail Then
        mshList.BackColorSel = &HE0E0E0
        mshDetail.BackColorSel = &HC0C0C0
    End If
End Sub

Private Sub SetDetail()
    Dim strHead As String
    Dim i As Long
    
    strHead = "���,1,650|����,1,1600" & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "|��Ʒ��,1,1600", "") & "|���,1,1000|��λ,4,500|����,7,850|����,7,850|Ӧ�ս��,7,850|ʵ�ս��,7,850|ͳ����,7,850|ִ�п���,1,850|����,1,850|˵��,1,1000|��¼״̬,1,0"
    
    With mshDetail
        .Redraw = False
        
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(mshDetail, App.ProductName & "\" & Me.Name)
        '���˺�:27990 2010-02-22 17:34:32
        For i = 0 To .Cols - 1
            If .TextMatrix(0, i) = "��Ʒ��" Then
                If gTy_System_Para.bytҩƷ������ʾ = 2 Then
                    If .ColWidth(i) <= 0 Then .ColWidth(i) = 1600
                Else
                    .ColWidth(i) = 0
                End If
            End If
        Next
                
        .RowHeight(0) = 320
        .ColWidth(.Cols - 1) = 0
        
        .Row = 1: .Col = 0: .ColSel = .Cols - 1
        
        Call mshDetail_EnterCell
        
        .Redraw = True
    End With
End Sub

Private Sub ShowDetail(Optional ByVal strNO As String, Optional ByVal strTime As String, _
    Optional ByVal blnDel As Boolean, Optional ByVal blnSort As Boolean)
    
    Dim strSql As String, i As Long, j As Long
    
    On Error GoTo errH
        
    If Not blnSort Then
        
        If frmTechnoFilter.mblnDateMoved Then
            mblnNOMoved = zlDatabase.NOMoved("סԺ���ü�¼", strNO, , 2, Me.Caption)
        Else
            mblnNOMoved = False   '����Ҫ����һ��
        End If
        
        strSql = _
        " Select C.���� as ���,Nvl(E.����,B.����) as ����," & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "E1.���� as ��Ʒ��,", "") & "B.���," & _
                IIf(gblnסԺ��λ, "Decode(X.ҩƷID,NULL,A.���㵥λ,X.סԺ��λ)", "A.���㵥λ") & " as ��λ," & _
        "       To_Char(Avg(Nvl(A.����,1)*" & IIf(blnDel, "-1*", "") & "A.����)" & _
                IIf(gblnסԺ��λ, "/Nvl(X.סԺ��װ,1)", "") & ",'9999990.00000') as ����, " & _
        "       To_Char(Sum(A.��׼����)" & IIf(gblnסԺ��λ, "*Nvl(X.סԺ��װ,1)", "") & ",'99999" & gstrFeePrecisionFmt & "') as ����, " & _
        "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.Ӧ�ս��),'9999999" & gstrDec & "') as Ӧ�ս��, " & _
        "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.ʵ�ս��),'9999999" & gstrDec & "') as ʵ�ս��, " & _
        "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.ͳ����),'9999999" & gstrDec & "') as ͳ����, " & _
        "       D.���� as ִ�п���,Nvl(A.��������,B.��������) as ����," & _
        "       Decode(Nvl(A.ִ��״̬,0),0,'δִ��',1,'��ִ��',2,'����ִ��','��'||ABS(A.ִ��״̬)||'���˷�') as ˵�� , A.��¼״̬" & _
        " From " & IIf(mblnNOMoved, zlGetFullFieldsTable("סԺ���ü�¼"), "סԺ���ü�¼ A") & " ," & _
        "       �շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,�շ���Ŀ���� E,ҩƷ��� X" & _
                  IIf(gTy_System_Para.bytҩƷ������ʾ = 2, ",�շ���Ŀ���� E1", "") & _
        " Where A.�շ�ϸĿID=B.ID And A.�շ����=C.���� And A.ִ�в���ID=D.ID(+)" & _
        "       And A.NO=[1] And A.��¼����=2 And A.�����־=2 And Nvl(A.�ಡ�˵�,0)=0" & _
        "       And A.�շ�ϸĿID=X.ҩƷID(+) And A.��¼״̬" & IIf(blnDel, "=2", " IN(1,3)") & IIf(strTime <> "", " And A.�Ǽ�ʱ��=[2]", "") & _
        "       And A.�շ�ϸĿID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
                IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3", "") & _
        " Group by Nvl(A.�۸񸸺�,A.���),C.����,Nvl(E.����,B.����)," & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "E1.���� ,", "") & " B.���,A.���㵥λ," & _
        "       D.����,Nvl(A.��������,B.��������),A.ִ��״̬,A.��¼״̬,X.ҩƷID,X.סԺ��λ,Nvl(X.סԺ��װ,1)" & _
        " Order by Nvl(A.�۸񸸺�,A.���)"
        If strTime <> "" Then
            Set mrsDetail = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNO, CDate(strTime))
        Else
            Set mrsDetail = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNO)
        End If
    End If
        
    mshDetail.Redraw = False
    
    mshDetail.Clear
    mshDetail.Rows = 2
    
    mshDetail.ForeColor = IIf(blnDel, &HC0, ForeColor)

    If Not mrsDetail.EOF Then Set mshDetail.DataSource = mrsDetail
    
    '������ɫ
    If blnDel Then
        '�˷�ֱ��Ϊ��ɫ
        mshDetail.ForeColor = &HC0
    Else
        'ԭʼ�����˹���Ϊ��ɫ
        mshDetail.ForeColor = ForeColor
        For i = 1 To mshDetail.Rows - 1
            If Val(mshDetail.TextMatrix(i, mshDetail.Cols - 1)) = 3 Then
                mshDetail.Row = i
                For j = 0 To mshDetail.Cols - 1
                    mshDetail.Col = j
                    mshDetail.CellForeColor = &HC00000
                Next
            End If
        Next
    End If

    Call SetDetail
        
    mshDetail.Redraw = True
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuViewRefeshOptionItem_Click(Index As Integer)
    Dim i As Long
    For i = 0 To mnuViewRefeshOptionItem.UBound
        mnuViewRefeshOptionItem(i).Checked = i = Index
    Next
End Sub


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

