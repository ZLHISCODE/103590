VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManageBilling 
   AutoRedraw      =   -1  'True
   Caption         =   "������ʹ���"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9675
   Icon            =   "frmManageBilling.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   Picture         =   "frmManageBilling.frx":08CA
   ScaleHeight     =   6210
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   5850
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmManageBilling.frx":0A58
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8229
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
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
            Picture         =   "frmManageBilling.frx":12EC
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
      TabIndex        =   5
      Top             =   0
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   9675
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinHeight1      =   720
      Width1          =   810
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   30
         TabIndex        =   6
         Top             =   30
         Width           =   9555
         _ExtentX        =   16854
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
            NumButtons      =   17
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
               Object.ToolTipText     =   "������ʴ���"
               Object.Tag             =   "����"
               ImageKey        =   "Billing"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Price"
               Description     =   "����"
               Object.ToolTipText     =   "���ʻ���"
               Object.Tag             =   "����"
               ImageKey        =   "Price"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "���"
               Key             =   "Auditing"
               Description     =   "���"
               Object.ToolTipText     =   "�������"
               Object.Tag             =   "���"
               ImageKey        =   "Auditing"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Billing_"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "Modi"
               Description     =   "�޸�"
               Object.ToolTipText     =   "�޸�"
               Object.Tag             =   "�޸�"
               ImageKey        =   "Modi"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Del"
               Description     =   "����"
               Object.ToolTipText     =   "�Ե�ǰѡ�е�������"
               Object.Tag             =   "����"
               ImageKey        =   "Del"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Del_"
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "View"
               Description     =   "����"
               Object.ToolTipText     =   "���ĵ�ǰ���ݵ�����"
               Object.Tag             =   "����"
               ImageKey        =   "View"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Filter"
               Description     =   "����"
               Object.ToolTipText     =   "��������������ɸѡ��¼"
               Object.Tag             =   "����"
               ImageKey        =   "Filter"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��λ"
               Key             =   "Go"
               Description     =   "��λ"
               Object.ToolTipText     =   "��λ�����������ļ�¼��"
               Object.Tag             =   "��λ"
               ImageKey        =   "Go"
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "����"
               Object.ToolTipText     =   "��ǰ��������"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
      Height          =   2805
      Left            =   15
      TabIndex        =   0
      Top             =   1065
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   4948
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
      MouseIcon       =   "frmManageBilling.frx":148A
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox picHsc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   45
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   9570
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3900
      Width           =   9570
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   1875
      Left            =   0
      TabIndex        =   1
      Top             =   3975
      Width           =   9660
      _ExtentX        =   17039
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
      MouseIcon       =   "frmManageBilling.frx":17A4
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   675
      Top             =   30
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
            Picture         =   "frmManageBilling.frx":1ABE
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":1CD8
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":1EF2
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":210C
            Key             =   "View"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":2886
            Key             =   "Go"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":2AA0
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":2CBA
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":2ED4
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":30EE
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":3308
            Key             =   "Billing"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":3A02
            Key             =   "Price"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":40FC
            Key             =   "Auditing"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   90
      Top             =   30
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
            Picture         =   "frmManageBilling.frx":47F6
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":4A10
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":4C2A
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":4E44
            Key             =   "View"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":55BE
            Key             =   "Go"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":57D8
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":59F2
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":5C0C
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":5E26
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":6040
            Key             =   "Billing"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":673A
            Key             =   "Price"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":6E34
            Key             =   "Auditing"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip tbs 
      Height          =   420
      Left            =   15
      TabIndex        =   2
      Top             =   735
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   741
      TabWidthStyle   =   2
      TabFixedWidth   =   2293
      TabFixedHeight  =   526
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "���ʵ���(&1)"
            Key             =   "Auditing"
            Object.ToolTipText     =   "��ʾֱ�Ӽ��ʻ򻮼ۺ�����˵ļ��ʵ���"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "���۵���(&2)"
            Key             =   "Price"
            Object.ToolTipText     =   "��ʾ���ۺ�δ��˵ļ��ʵ���"
            ImageVarType    =   2
         EndProperty
      EndProperty
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
      Begin VB.Menu mnuEdit_Billing 
         Caption         =   "�������(&B)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditCust 
         Caption         =   "�Զ�����ʵ�(&U)"
         Begin VB.Menu mnuEditCustBill 
            Caption         =   "(��)"
            Index           =   1
         End
      End
      Begin VB.Menu mnuEditPrice 
         Caption         =   "���ʻ���(&P)"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuEditAuditing 
         Caption         =   "�������(&A)"
         Shortcut        =   +{F2}
      End
      Begin VB.Menu mnuEditAuditingPati 
         Caption         =   "���������(&N)"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuEditBilling_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Modi 
         Caption         =   "�޸ĵ���(&M)"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuEdit_Adjust 
         Caption         =   "����ʱ��(&J)"
         Shortcut        =   ^J
      End
      Begin VB.Menu mnuEdit_Adjust_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Del 
         Caption         =   "��������(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEdit_Del_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_View 
         Caption         =   "���ĵ���(&V)"
      End
      Begin VB.Menu mnuEdit_Print 
         Caption         =   "��ӡ����(&I)"
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
      Begin VB.Menu mnuView_5 
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
Attribute VB_Name = "frmManageBilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private mrsList As ADODB.Recordset  '�����б�
Private mrsTotal As ADODB.Recordset
Private mrsDetail As ADODB.Recordset
Private Type Type_SQLCondition
    Default As Boolean          '�Ƿ���ȱʡ���룬��ʱû������ֵ,ȱʡֵ��mstrFilter��
    DateB As Date
    DateE As Date
    NOB As String
    NOE As String
    Patient As String
    PatientIdentity As String
    DeptID As Long
    Operator As String
    PatientNo As String '�����:38539
    PatientID As Long '�����:38539
End Type
Private SQLCondition As Type_SQLCondition
Private mstrFilter As String
Private mbln���� As Boolean, mbln���� As Boolean
Private mstr����Ա As String
Private mstrPage As String, mblnMax As Boolean
Private mblnGo As Boolean, mlngGo As Long
Private mlngCurRow As Long, mlngTopRow As Long
Private mstrPrivs As String   '���浱ǰ����Ȩ����
Private mlngModul As Long
Private mblnNOMoved As Boolean '����ϸʱ��¼��ǰѡ��ĵ����Ƿ����������ݱ���,����������ʱ�������ж�
'��Ϣ��ض������
Private WithEvents mobjMsgModule As clsMipModule
Attribute mobjMsgModule.VB_VarHelpID = -1

Private Sub cbr_Resize()
    Form_Resize
End Sub

Private Sub Form_Activate()
    Call InitLocPar(mlngModul)
    Call mshList_GotFocus
End Sub
Private Sub mnuEdit_Adjust_Click()
    Dim strNo As String
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    If strNo = "" Then
        MsgBox "��ǰû�е��ݿ��Ե�����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '�Ƿ���ת������ݱ���
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNo, 2, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '��ʱ��ת���������ݱ�
    End If
    
    '�Ѿ�������(����)�ĵ��ݲ��������
    If BillExistDelete(strNo, 2) Then
        MsgBox "�õ��ݰ�������������,�����������", vbInformation, gstrSysName
        Exit Sub
    End If

    '�ѽ��ʵ��ݸ��ݲ�������
    If HaveBilling(1, strNo) Then
        Select Case gbytBillOpt
            Case 0
            Case 1
                If MsgBox("�ü��ʵ����Ѿ�����,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Case 2
                MsgBox "�ü��ʵ����Ѿ�����,���ܵ�����", vbExclamation, gstrSysName: Exit Sub
        End Select
    End If

    On Error Resume Next
    Err.Clear

    '��ʾ��������
    Dim lng����ID As Long
    Dim varTemp As Variant
    
    lng����ID = mshList.TextMatrix(mshList.Row, GetColNum("���ʵ�ID"))
    
    If lng����ID = 0 Or gobjCustBill Is Nothing Then
        frmCharge.mlngModul = mlngModul
        frmCharge.mstrPrivs = mstrPrivs
        frmCharge.mbytInFun = 2
        frmCharge.mbytInState = 2
        frmCharge.mstrInNO = strNo
        Set frmCharge.mobjMsgModule = mobjMsgModule
        frmCharge.Show 1, Me
    Else
        '����ID��bytUseType��bytInState��strInNO��lngUnitID��lngDeptID��lng����ID��mstrPrivs
        varTemp = Array(lng����ID, 3, 2, strNo, 0, 0, 0, mstrPrivs)
        gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
        
        gblnOK = varTemp
    End If
End Sub

Private Sub mnuEdit_Modi_Click()
    Dim strNo As String
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    
    If strNo = "" Then
        MsgBox "��ǰû�е��ݿ����޸ģ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '�Ƿ���ת������ݱ���
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNo, 2, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '��ʱ��ת���������ݱ�
    End If
    gstrModiNO = ""
    
    On Error Resume Next
    Err.Clear
    
    Dim lng����ID As Long
    Dim varTemp As Variant
    
    lng����ID = mshList.TextMatrix(mshList.Row, GetColNum("���ʵ�ID"))
    
    If lng����ID = 0 Or gobjCustBill Is Nothing Then
        frmCharge.mlngModul = mlngModul
        frmCharge.mstrPrivs = mstrPrivs
        frmCharge.mbytInFun = 2
        frmCharge.mstrInNO = strNo
        frmCharge.mbytInState = 0
        frmCharge.mbytBilling = IIf(tbs.SelectedItem.Key = "Auditing", 0, 1)
        Set frmCharge.mobjMsgModule = mobjMsgModule
        frmCharge.Show 1, Me
    Else
        '����ID��bytUseType��bytInState��strInNO��lngUnitID��lngDeptID��lng����ID��mstrPrivs
        varTemp = Array(lng����ID, 3, 0, strNo, 0, 0, 0, mstrPrivs)
        gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
        
        gblnOK = varTemp
    End If
    
    If gblnOK Then
        If gstrModiNO <> "" Then
            If mnuViewRefeshOptionItem(1).Checked Then
                If MsgBox("��ǰ�����Ѹ��ĵ����嵥����,�޸ĺ�ĵ��ݺ�Ϊ:[" & gstrModiNO & "],Ҫˢ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    mnuViewReFlash_Click
                End If
            ElseIf mnuViewRefeshOptionItem(2).Checked Then
                mnuViewReFlash_Click
            End If
        Else
            If mnuViewRefeshOptionItem(1).Checked Then
                If MsgBox("��ǰ�����Ѹ��ĵ����嵥����,Ҫˢ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    mnuViewReFlash_Click
                End If
            ElseIf mnuViewRefeshOptionItem(2).Checked Then
                mnuViewReFlash_Click
            End If
        End If
    End If
End Sub

Private Sub mnuEdit_Print_Click()
    Dim strNo As String, strTime As String
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    If strNo = "" Then
        MsgBox "��ǰû�е��ݿ��Դ�ӡ��", vbInformation, gstrSysName
        Exit Sub
    End If
    If InStr(",0,1,", Val(mshList.TextMatrix(mshList.Row, GetColNum("����")))) = 0 Then
        MsgBox "�õ���Ϊ���ʵ��ݻ��ѱ����ʣ������ٴ�ӡ��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '�Ƿ���ת������ݱ���
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNo, 2, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '��ʱ��ת���������ݱ�
    End If
    
    If ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1122", Me) Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1122", Me, "NO=" & strNo, "ҩƷ��λ=" & IIf(gblnҩ����λ, 1, 0), 2)
    End If
End Sub

Private Sub mnuEditAuditing_Click()
    On Error Resume Next
    frmCharge.mlngModul = mlngModul
    frmCharge.mstrPrivs = mstrPrivs
    frmCharge.mbytInFun = 2
    frmCharge.mbytInState = 0
    frmCharge.mbytBilling = 2
    Set frmCharge.mobjMsgModule = mobjMsgModule
    frmCharge.Show 1, Me
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
Private Sub mnuEditAuditingPati_Click()
    On Error Resume Next
    If Not frmBillingAuditing.zlShowCard(Me, mlngModul, mstrPrivs) Then Exit Sub
    If mnuViewRefeshOptionItem(1).Checked Then
        If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            mnuViewReFlash_Click
        End If
    ElseIf mnuViewRefeshOptionItem(2).Checked Then
        mnuViewReFlash_Click
    End If
End Sub

Private Sub mnuEditPrice_Click()
    On Error Resume Next
    Err.Clear
    frmCharge.mlngModul = mlngModul
    frmCharge.mstrPrivs = mstrPrivs
    frmCharge.mbytInFun = 2
    frmCharge.mbytInState = 0
    frmCharge.mbytBilling = 1
    Set frmCharge.mobjMsgModule = mobjMsgModule
    frmCharge.Show 1, Me
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

Private Sub mnuFileLocalSet_Click()
    Dim blnPre As Boolean
    
    blnPre = gblnҩ����λ
    
    With frmSetExpence
        .mlngModul = mlngModul
        .mstrPrivs = mstrPrivs
        .mbytInFun = 2
        .mblnSetDrugStore = False
        .Show 1, Me
    End With
        
    '������ҩƷ��λ����,����ˢ��
    If gblnҩ����λ <> blnPre Then
        ShowBills mstrFilter
    End If
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim strNo As String
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    If strNo <> "" Then
        With mshList
            Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
                    "NO=" & .TextMatrix(.Row, GetColNum("���ݺ�")), "����ID=" & .TextMatrix(.Row, GetColNum("����ID")), _
                    "������=" & .TextMatrix(.Row, GetColNum("������")))
        End With
    Else
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me)
    End If
End Sub

Private Sub mnuViewFilter_Click()
    With frmBillingFilter
        .mstrPrivs = mstrPrivs
        .chk����.Enabled = tbs.SelectedItem.Key = "Auditing"
        .chk����.Enabled = tbs.SelectedItem.Key = "Auditing"
        .Show 1, Me
        If gblnOK Then
            mstrFilter = .mstrFilter
            mbln���� = .chk����.Value = 1
            mbln���� = .chk����.Value = 1
            If .cbo����Ա.Text <> "���в���Ա" Then
                mstr����Ա = zlStr.NeedName(.cbo����Ա.Text)
            Else
                mstr����Ա = ""
            End If
            SQLCondition.DateB = .dtpBegin.Value
            SQLCondition.DateE = .dtpEnd.Value
            SQLCondition.NOB = .txtNOBegin.Text
            SQLCondition.NOE = .txtNoEnd.Text
            SQLCondition.Patient = gstrLike & UCase(.txt����.Text) & "%"
            SQLCondition.PatientIdentity = .txt����ID.Text
            SQLCondition.DeptID = .cbo����.ItemData(.cbo����.ListIndex)
            SQLCondition.Operator = mstr����Ա
            '�����:38539
            SQLCondition.PatientNo = .txtPatientNo
            SQLCondition.PatientID = .mlngPrePatient
            
            mnuViewReFlash_Click
        End If
    End With
End Sub

Private Sub mnuViewRefeshOptionItem_Click(Index As Integer)
    Dim i As Long
    For i = 0 To mnuViewRefeshOptionItem.UBound
        mnuViewRefeshOptionItem(i).Checked = i = Index
    Next
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
    If mnuEdit_View.Enabled Then mnuEdit_View_Click
End Sub

Private Sub mshList_EnterCell()
    Dim strNo As String, strTime As String, blnDel As Boolean
        
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    
    If mshList.Row = 0 Or strNo = "" Then Exit Sub
    
    stbThis.Panels(2).Text = "�� " & NVL(mrsTotal!����, 0) & " �ŵ���,�ϼ�:" & Format(NVL(mrsTotal!���, 0), gstrDec)
    
    mlngGo = mshList.Row
    mlngCurRow = mshList.Row: mlngTopRow = mshList.TopRow
    
    strTime = mshList.TextMatrix(mshList.Row, GetColNum("�Ǽ�ʱ��"))
    blnDel = Val(mshList.TextMatrix(mshList.Row, GetColNum("����"))) = 2
    
    mnuEdit_Adjust.Enabled = Not blnDel
    mnuEdit_Modi.Enabled = Not blnDel
    mnuEdit_Del.Enabled = Not blnDel
    tbr.Buttons("Modi").Enabled = Not blnDel
    tbr.Buttons("Del").Enabled = Not blnDel
        
    mshList.ForeColorSel = mshList.CellForeColor
    
    Call ShowDetail(strNo, strTime, blnDel)
End Sub

Private Sub mshList_GotFocus()
    Call SetActiveList(mshList)
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
        Case vbKeyReturn
            If mnuEdit_View.Enabled Then mnuEdit_View_Click
        Case vbKeyEscape
            mblnGo = False
    End Select
End Sub

Private Sub mnuEdit_Del_Click()
    Dim strNo As String, strTime As String
    Dim strSQL As String, i As Long, blnFlagPrint As Boolean
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    
    If strNo = "" Then
        MsgBox "��ǰû�е��ݿ����˷ѣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    strTime = mshList.TextMatrix(mshList.Row, GetColNum("�Ǽ�ʱ��"))
    
    '���������еļ�����,�پ����Ƿ��뵽���߱�,�����м��ʱҪ��mblnNOMoved��������ж�,Ϊ��,�ݶ�Ϊ�ȼ���Ƿ���ת��
    '�Ƿ���ת������ݱ���
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNo, 2, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '��ʱ��ת���������ݱ�
    End If
    
    '����Ȩ�޼��
    If Not BillOperCheck(4, mshList.TextMatrix(mshList.Row, GetColNum("����Ա")), CDate(strTime), "����", strNo, , 2) Then Exit Sub
    
    '�Ƿ���ִ��
    i = BillCanDelete(strNo, 2, , strTime, blnFlagPrint)
    If i <> 0 Then
        Select Case i
            Case 1 '�õ��ݲ�����
                MsgBox "ָ���ĵ��ݲ����ڣ�", vbInformation, gstrSysName
            Case 2 '�Ѿ�ȫ����ȫִ��
                MsgBox "�õ����е���Ŀ�Ѿ�ȫ����ȫִ�У�", vbInformation, gstrSysName
            Case 3 'δ��ȫִ�в���ʣ������Ϊ0
                MsgBox "�õ�����δ��ȫִ�в�����Ŀʣ������Ϊ��,û�п������ʵķ��ã�", vbInformation, gstrSysName
        End Select
        Exit Sub
    End If
    If blnFlagPrint Then
        If MsgBox("ע��:����ҽ���������Ѵ�ӡ���Ƿ������", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    
    '�ѽ��ʵ��ݸ��ݲ�������
    If HaveBilling(1, strNo, True, strTime) Then
        Select Case gbytBillOpt
            Case 1
                If MsgBox("�õ����Ѿ�����,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Case 2
                MsgBox "�õ����Ѿ�����,�������ʣ�", vbExclamation, gstrSysName: Exit Sub
        End Select
    End If
    
    On Error Resume Next
    Err.Clear
    
    '��ʾ��������
    Dim lng����ID As Long, varTemp As Variant
    
    lng����ID = mshList.TextMatrix(mshList.Row, GetColNum("���ʵ�ID"))
    
    If lng����ID = 0 Or gobjCustBill Is Nothing Then
        frmCharge.mlngModul = mlngModul
        frmCharge.mstrPrivs = mstrPrivs
        frmCharge.mbytInFun = 2
        frmCharge.mbytInState = 3
        frmCharge.mstrInNO = strNo
        frmCharge.mstrTime = strTime
        frmCharge.mbytBilling = IIf(tbs.SelectedItem.Key = "Auditing", 0, 1)
        Set frmCharge.mobjMsgModule = mobjMsgModule
        frmCharge.Show 1, Me
    Else
        '����ID��bytUseType��bytInState��strInNO��lngUnitID��lngDeptID��lng����ID��mstrPrivs
        varTemp = Array(lng����ID, 3, 3, strNo, 0, 0, 0, mstrPrivs)
        gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
        gblnOK = varTemp
    End If
    
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

Private Sub mnuHelpTitle_Click()
ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub mnuEdit_Billing_Click()
    On Error Resume Next
    Err.Clear
    frmCharge.mlngModul = mlngModul
    frmCharge.mstrPrivs = mstrPrivs
    frmCharge.mbytInFun = 2
    frmCharge.mbytInState = 0
    frmCharge.mbytBilling = 0
    Set frmCharge.mobjMsgModule = mobjMsgModule
    frmCharge.Show 1, Me
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

Private Sub mnuEditCustBill_Click(Index As Integer)
    '�Զ������
    Dim varTemp As Variant
    '�������������ǣ�
    '����ID��bytUseType��bytInState��strInNO��lngUnitID��lngDeptID��lng����ID��mstrPrivs��blnViewCancel
    
    varTemp = Array(mnuEditCustBill(Index).Tag, 3, 0, "", 0, 0, 0, mstrPrivs)
    gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
    
    gblnOK = varTemp '����ֵ
    
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

Private Sub mnuEdit_View_Click()
    Dim strNo As String, strTime As String, blnDel As Boolean
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    If strNo = "" Then
        MsgBox "��ǰû�е��ݿ��Բ��ģ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    strTime = mshList.TextMatrix(mshList.Row, GetColNum("�Ǽ�ʱ��"))
    blnDel = Val(mshList.TextMatrix(mshList.Row, GetColNum("����"))) = 2
    
    On Error Resume Next
    Err.Clear
    
    '��ʾ��������
    Dim lng����ID As Long
    Dim varTemp As Variant
    lng����ID = mshList.TextMatrix(mshList.Row, GetColNum("���ʵ�ID"))
    
    If lng����ID = 0 Or gobjCustBill Is Nothing Then
        frmCharge.mlngModul = mlngModul
        frmCharge.mstrPrivs = mstrPrivs
        frmCharge.mbytInFun = 2
        frmCharge.mbytInState = 1
        frmCharge.mstrTime = strTime
        frmCharge.mblnDelete = blnDel
        frmCharge.mstrInNO = strNo
        frmCharge.mblnNOMoved = mblnNOMoved
        frmCharge.mbytBilling = IIf(tbs.SelectedItem.Key = "Auditing", 0, 1)
        Set frmCharge.mobjMsgModule = mobjMsgModule
        frmCharge.Show 1, Me
    Else
        '����ID��bytUseType��bytInState��strInNO��lngUnitID��lngDeptID��lng����ID��mstrPrivs
        varTemp = Array(lng����ID, 3, 1, strNo, 0, 0, 0, mstrPrivs, blnDel)
        gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
        
        gblnOK = varTemp
    End If
End Sub

Private Sub mnuFile_quit_Click()
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
    Dim i As Long
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For i = 1 To tbr.Buttons.Count
        tbr.Buttons(i).Caption = IIf(mnuViewToolText.Checked, tbr.Buttons(i).Tag, "")
    Next
    cbr.Bands(1).MinHeight = tbr.ButtonHeight
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
            mnuEdit_View_Click
        Case "Billing"
            mnuEdit_Billing_Click
        Case "Price"
            mnuEditPrice_Click
        Case "Auditing"
            mnuEditAuditing_Click
        Case "Modi"
            mnuEdit_Modi_Click
        Case "Del"
            mnuEdit_Del_Click
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
    objOut.Title.Text = "������ʵ����嵥"
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '����
    With frmBillingFilter
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
    mshList.Col = 0: mshList.ColSel = mshList.COLS - 1
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
    
    mnuEdit_Adjust.Enabled = blnUsed
    mnuEdit_Modi.Enabled = blnUsed
    tbr.Buttons("Modi").Enabled = blnUsed
    
    mnuEdit_Del.Enabled = blnUsed
    mnuEdit_View.Enabled = blnUsed
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
    
    mstrPrivs = mstrPrivs
    
    If gobjCustBill Is Nothing Then
        Set gobjCustBill = CreateObject("zl9CustAcc.clsCustAcc")
    End If
    If InStr(mstrPrivs, "ר�����") = 0 Then
        mnuEditCust.Visible = False
        Exit Sub
    End If

    On Error GoTo errHandle
    

    
    '��������ɹ����ٶ�����Ӧ�Ĳ˵�
    If Not gobjCustBill Is Nothing Then
        gstrSQL = "Select ID,���� From �շѼ��ʵ� Where substr(���÷�Χ,1,1)='1' Order by ���"
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
    
    i = IIf(zlDatabase.GetPara("ҳ��", glngSys, mlngModul, "1") = "1", 1, 2)
    tbs.Tabs(i).Selected = True
    
    mlngCurRow = 1: mlngTopRow = 1
    mstrPage = ""
    
    'Ȩ������
    If InStr(mstrPrivs, "�������") = 0 Then
        mnuEdit_Billing.Visible = False
        tbr.Buttons("Billing").Visible = False
        
        mnuEditCust.Visible = False
    End If
    If InStr(mstrPrivs, "���ʻ���") = 0 Then
        mnuEditPrice.Visible = False
        tbr.Buttons("Price").Visible = False
    End If
    If InStr(mstrPrivs, "�������") = 0 Then
        mnuEditAuditing.Visible = False
        mnuEditAuditingPati.Visible = False
        tbr.Buttons("Auditing").Visible = False
    End If
    If InStr(mstrPrivs, "�������") = 0 _
        And InStr(mstrPrivs, "���ʻ���") = 0 _
        And InStr(mstrPrivs, "�������") = 0 Then
        mnuEditBilling_.Visible = False
        tbr.Buttons("Billing_").Visible = False
    End If
    
    If InStr(mstrPrivs, "��¼�޸�") = 0 Then
        mnuEdit_Modi.Visible = False
        tbr.Buttons("Modi").Visible = False
    End If
    If InStr(mstrPrivs, "��¼����") = 0 Then
        mnuEdit_Adjust.Visible = False
    End If
    If InStr(mstrPrivs, "��¼�޸�") = 0 And InStr(mstrPrivs, "��¼����") = 0 Then
        mnuEdit_Adjust_.Visible = False
    End If
    
    If InStr(mstrPrivs, "��������") = 0 Then
        mnuEdit_Del.Visible = False
        mnuEdit_Del_.Visible = False
        tbr.Buttons("Del").Visible = False
        tbr.Buttons("Del_").Visible = False
    End If


    mbln���� = True
    mbln���� = False
    mstr����Ա = UserInfo.����
    
    Call SetHeader
    Call SetDetail
    Call SetMenu(False)
    
    stbThis.Panels(2).Text = "��ˢ���嵥���������ù�������"
    
    '��ʼ����Ϣ�������ģ��
    Call zlMsgModuleInit
    
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
    
    tbs.Left = Me.ScaleLeft
    tbs.Top = Me.ScaleTop + cbrH + 15
    
    mshList.Left = 0
    mshList.Top = tbs.Top + tbs.TabFixedHeight + 30
    mshList.Width = Me.ScaleWidth
    mshList.Height = (Me.ScaleHeight - cbrH - staH - (tbs.TabFixedHeight + 45) - picHsc.Height) * (1 - sngVsc)
    
    picHsc.Top = mshList.Top + mshList.Height
    picHsc.Left = Me.ScaleLeft
    picHsc.Width = Me.ScaleWidth
    
    mshDetail.Left = Me.ScaleLeft
    mshDetail.Top = picHsc.Top + picHsc.Height
    mshDetail.Width = Me.ScaleWidth
    mshDetail.Height = Me.ScaleHeight - cbrH - staH - (tbs.TabFixedHeight + 45) - picHsc.Height - mshList.Height
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    
    mstrFilter = ""
    Unload frmBillingFilter
    Unload frmBillingGo
    
    Call SaveWinState(Me, App.ProductName)
    zlDatabase.SetPara "ҳ��", tbs.SelectedItem.Index, glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
    'ˢ�·�ʽ
    For i = 0 To mnuViewRefeshOptionItem.UBound
        If mnuViewRefeshOptionItem(i).Checked Then
            zlDatabase.SetPara "ˢ�·�ʽ", i, glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
            Exit For
        End If
    Next
    '��ж��Ϣ����
    Call zlMsgModuleUnload
End Sub

Private Sub mnuViewGo_Click()
    frmBillingGo.Show 1, Me
    If gblnOK Then Call SeekBill(frmBillingGo.optHead)
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
        With frmBillingGo
            If .txtNO.Text <> "" Then
                blnFill = blnFill And mshList.TextMatrix(i, GetColNum("���ݺ�")) = .txtNO.Text
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
            mshList.Col = 0: mshList.ColSel = mshList.COLS - 1
            
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
    For i = 0 To mshList.COLS - 1
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
    
    strHead = "���ݺ�,1,850|��������,1,850|������,1,800|����ID,1,750|�����,1,900|����,1,700|�ѱ�,1,900|Ӧ�ս��,7,850|ʵ�ս��,7,850|����Ա,1,800|�Ǽ�ʱ��,1,1850|˵��,1,850|����,1,0|���ʵ�ID,1,0"
    With mshList
        .COLS = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(mshList, App.ProductName & "\" & Me.Name)
                
         'ҽ����
        i = GetColNum("������")
        If InStr(mstrPrivs, "ҽ����ѯ") = 0 Then
            .ColWidth(i) = 0
        ElseIf mshList.ColWidth(i) = 0 Then
            .ColWidth(i) = 800
        End If
                
        i = GetColNum("����"): mshList.ColWidth(i) = 0
        i = GetColNum("���ʵ�ID"): mshList.ColWidth(i) = 0
        
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
        
        .Col = 0: .ColSel = .COLS - 1
                
        Call mshList_EnterCell
    End With
End Sub

Private Sub SetDetail()
    Dim strHead As String
    Dim i As Long
    
    strHead = "���,1,750|����,1,1800" & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "|��Ʒ��,1,2000", "") & "|���,1,1000|��λ,4,500|����,7,850|����,7,850|Ӧ�ս��,7,850|ʵ�ս��,7,850|ִ�п���,1,850|����,1,850|˵��,1,1000|��¼״̬,1,0"
    
    With mshDetail
        .Redraw = False
        
        .COLS = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(mshDetail, App.ProductName & "\" & Me.Name)
        For i = 0 To .COLS - 1
            If .TextMatrix(0, i) = "��Ʒ��" Then
                If gTy_System_Para.bytҩƷ������ʾ = 2 Then
                    If .ColWidth(i) <= 0 Then .ColWidth(i) = 2000
                Else
                    .ColWidth(i) = 0
                End If
            End If
        Next
        
        .RowHeight(0) = 320
        .ColWidth(.COLS - 1) = 0
        
        .Row = 1: .Col = 0: .ColSel = .COLS - 1
        
        Call mshDetail_EnterCell
        
        .Redraw = True
    End With
End Sub

Private Sub ShowBills(Optional ByVal strFilter As String, Optional blnSort As Boolean)
'����:��������ȡ�����б�(���˹���)
'����:strFilter=��"AND"��ʼ��������
'     blnSort=�����¶�ȡ����,��������ʾ�����������
    Dim i As Long, j As Long, k As Long
    Dim strSQL As String
    
    On Error GoTo errH
    
    If Not blnSort Then
        Call zlCommFun.ShowFlash("���ڶ�ȡ�����б�,���Ժ� ...", Me)
        DoEvents
        Me.Refresh
        
        SQLCondition.Default = (strFilter = "")
        If strFilter = "" Then
            'ȱʡ��������(һ����)
            strFilter = " And �Ǽ�ʱ�� Between Trunc(Sysdate) And Trunc(Sysdate+1)-1/24/60/60"
        End If
            
        '����Ա��������
        If mstr����Ա <> "" Then
            If tbs.SelectedItem.Key = "Auditing" Then
                strFilter = strFilter & " And ����Ա����||''=[10]"
            Else
                strFilter = strFilter & " And ������||''=[10]"
            End If
        End If
        
        strFilter = " Where ��¼����=2 And �����־ in(1,4) " & strFilter
        
        If frmBillingFilter.mblnDateMoved And tbs.SelectedItem.Key = "Auditing" Then  'ɸѡʱ��ʱ�������һ��ת��֮ǰ,�ҵ�ǰ�б��ǻ��۵�
            strFilter = zlGetFullFieldsTable("������ü�¼", 2, strFilter, False)
        Else
            strFilter = zlGetFullFieldsTable("������ü�¼", 0, strFilter, False)
        End If
        
        '���ݺ�,��������,������,����ID,�����,����,�ѱ�,Ӧ�ս��,ʵ�ս��,����Ա,�Ǽ�ʱ��,˵��,����,���ʵ�ID
        If tbs.SelectedItem.Key = "Auditing" Then
            '���ʵ�״̬
            If mbln���� And mbln���� Then
                strFilter = strFilter & " And ��¼״̬ IN([11],[12],[13])"
            ElseIf mbln���� Then
                strFilter = strFilter & " And ��¼״̬ IN([11],[13])"
            ElseIf mbln���� Then
                strFilter = strFilter & " And ��¼״̬=[12]"
            End If
            
            strFilter = strFilter & " And ��¼״̬<>0 And ����Ա���� IS NOT NULL"
            
            strSQL = _
                "Select A.NO as ���ݺ�,B.���� as ��������,A.������,A.����ID,A.��ʶ�� as �����,A.����,A.�ѱ�," & _
                " To_Char(Sum(Decode(A.��¼״̬,2,-1,1)*A.Ӧ�ս��),'9999999" & gstrDec & "') as Ӧ�ս��," & _
                " To_Char(Sum(Decode(A.��¼״̬,2,-1,1)*A.ʵ�ս��),'9999999" & gstrDec & "') as ʵ�ս��," & _
                " A.����Ա���� as ����Ա,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as �Ǽ�ʱ��," & _
                " Decode(Max(A.��¼״̬),2,'���ʼ�¼','���ʼ�¼') as ˵��,Max(A.��¼״̬) as ����,A.���ʵ�ID" & _
                " From (" & strFilter & ") A,���ű� B" & _
                " Where A.��������ID = B.ID" & _
                " And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & vbNewLine & _
                " Group by A.NO,B.����,A.������,A.����ID,A.��ʶ��,A.����,A.�ѱ�,A.����Ա����,A.�Ǽ�ʱ��,A.���ʵ�ID" & _
                " Order by A.�Ǽ�ʱ�� Desc,A.NO Desc"
        Else
            '���ʻ��۵�״̬,��δ���״̬
            strFilter = strFilter & " And ��¼״̬=[14] And ����Ա���� IS NULL And ������ is Not NULL"
            
            strSQL = _
                "Select A.NO as ���ݺ�,B.���� as ��������,A.������,A.����ID,A.��ʶ�� as �����,A.����,A.�ѱ�," & _
                " To_Char(Sum(Decode(A.��¼״̬,2,-1,1)*A.Ӧ�ս��),'9999999" & gstrDec & "') as Ӧ�ս��," & _
                " To_Char(Sum(Decode(A.��¼״̬,2,-1,1)*A.ʵ�ս��),'9999999" & gstrDec & "') as ʵ�ս��," & _
                " A.������ as ����Ա,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as �Ǽ�ʱ��," & _
                " Decode(Max(A.��¼״̬),2,'���ʼ�¼','���ʼ�¼') as ˵��,Max(A.��¼״̬) as ����,A.���ʵ�ID" & _
                " From (" & strFilter & ") A,���ű� B" & _
                " Where A.��������ID = B.ID" & _
                " And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & vbNewLine & _
                " Group by A.NO,B.����,A.������,A.����ID,A.��ʶ��,A.����,A.�ѱ�,A.������,A.�Ǽ�ʱ��,A.���ʵ�ID" & _
                " Order by A.�Ǽ�ʱ�� Desc,A.NO Desc"
        End If
        With SQLCondition
            If .Default Then
                Set mrsList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "", "", "", "", "", "", "", "", "", mstr����Ա, 1, 2, 3, 0)
            Else
                Set mrsList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .DateB, .DateE, .NOB, .NOE, .Patient, .PatientIdentity, .DeptID, .PatientNo, .PatientID, mstr����Ա, 1, 2, 3, 0)
            End If
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
            strSQL = "Select Sum(ʵ�ս��) as ���,Count(Distinct NO) as ���� From (" & _
                Replace(strFilter, "��¼״̬ IN([11],[13]", "��¼״̬ IN([11],[12],[13]") & ") A,���ű� B Where A.��������ID = B.ID" & _
                " And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)"

            With SQLCondition
                If .Default Then
                    Set mrsTotal = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "", "", "", "", "", "", "", "", "", mstr����Ա, 1, 2, 3, 0)
                Else
                    Set mrsTotal = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .DateB, .DateE, .NOB, .NOE, .Patient, Val(.PatientIdentity), .DeptID, .PatientNo, .PatientID, mstr����Ա, 1, 2, 3, 0)
                End If
            End With
        End If

        Set mshList.DataSource = mrsList
        stbThis.Panels(2).Text = "�� " & NVL(mrsTotal!����, 0) & " �ŵ���,�ϼ�:" & Format(NVL(mrsTotal!���, 0), gstrDec)
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
                For j = 0 To mshList.COLS - 1
                    mshList.Col = j
                    mshList.CellForeColor = &HC0
                Next
            ElseIf Val(mshList.TextMatrix(i, k)) = 3 Then
                '�������ʵ�����ɫ
                mshList.Row = i
                For j = 0 To mshList.COLS - 1
                    mshList.Col = j
                    mshList.CellForeColor = &HC00000
                Next
            End If
        Next
    End If
    
    Call SetHeader
    Call SetDetail
        
    mshList.Redraw = True
    
    If Not blnSort Then Call zlCommFun.StopFlash
    
    Me.Refresh
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub tbs_Click()
    If Not Visible Then Exit Sub
    If tbs.SelectedItem.Key = mstrPage Then Exit Sub
    mstrPage = tbs.SelectedItem.Key
     
    ShowBills mstrFilter
    On Error Resume Next
    mshList.SetFocus
End Sub

Private Sub SetActiveList(obj As Object)
    If obj Is mshList Then
        mshList.BackColorSel = &HC0C0C0
        mshDetail.BackColorSel = &HE0E0E0
    ElseIf obj Is mshDetail Then
        mshList.BackColorSel = &HE0E0E0
        mshDetail.BackColorSel = &HC0C0C0
    End If
End Sub

Private Sub ShowDetail(Optional ByVal strNo As String, Optional ByVal strTime As String, _
    Optional ByVal blnDel As Boolean, Optional ByVal blnSort As Boolean)

    Dim strSQL As String, i As Long, j As Long
    
    On Error GoTo errH
    
    If Not blnSort Then
        If frmBillingFilter.mblnDateMoved And tbs.SelectedItem.Key = "Auditing" Then
            '���ʻ��۵�������Ƿ��ں󱸱���,��Ϊ����ת�����󱸱�
            mblnNOMoved = zlDatabase.NOMoved("������ü�¼", strNo, , "2")
        Else
            mblnNOMoved = False   '����Ҫ����һ��
        End If
        strSQL = _
        " Select C.���� as ���,Nvl(E.����,B.����) as ����," & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "E1.���� as ��Ʒ��,", "") & "B.���," & _
                IIf(gblnҩ����λ, "Decode(X.ҩƷID,NULL,A.���㵥λ,X." & gstrҩ����λ & ")", "A.���㵥λ") & " as ��λ," & _
        "       To_Char(Avg(Nvl(A.����,1)*" & IIf(blnDel, "-1*", "") & "A.����)" & _
                IIf(gblnҩ����λ, "/Nvl(X." & gstrҩ����װ & ",1)", "") & ",'9999990.00000') as ����, " & _
        "       To_Char(Sum(A.��׼����)" & IIf(gblnҩ����λ, "*Nvl(X." & gstrҩ����װ & ",1)", "") & ",'999999" & gstrFeePrecisionFmt & "') as ����, " & _
        "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.Ӧ�ս��),'9999999" & gstrDec & "') as Ӧ�ս��, " & _
        "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.ʵ�ս��),'9999999" & gstrDec & "') as ʵ�ս��, " & _
        "       D.���� as ִ�п���,Nvl(A.��������,B.��������) as ����," & _
        "       Decode(Nvl(A.ִ��״̬,0),0,'δִ��',1,'��ȫִ��',2,'����ִ��','��'||ABS(A.ִ��״̬)||'���˷�') as ˵��," & _
        "       A.��¼״̬" & _
        " From " & IIf(mblnNOMoved, zlGetFullFieldsTable("������ü�¼"), "������ü�¼ A") & "," & _
        "       �շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,�շ���Ŀ���� E," & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "�շ���Ŀ���� E1,", "") & "ҩƷ��� X" & _
        " Where A.�շ�ϸĿID=B.ID and A.�շ����=C.���� And A.ִ�в���ID=D.ID(+) And A.�շ�ϸĿID=X.ҩƷID(+)" & _
        "       And A.��¼����=2 And A.NO=[1] And A.�����־ in(1,4) And A.��¼״̬" & IIf(blnDel, "=2", " IN(0,1,3)") & _
                IIf(strTime <> "", " And A.�Ǽ�ʱ��=[2]", "") & _
        "       And A.�շ�ϸĿID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
                IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3", "") & _
        " Group by Nvl(A.�۸񸸺�,A.���),C.����,Nvl(E.����,B.����)," & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "E1.���� ,", "") & " B.���," & _
        " A.���㵥λ,D.����,Nvl(A.��������,B.��������),A.ִ��״̬,A.��¼״̬,X.ҩƷID,X." & gstrҩ����λ & ",Nvl(X." & gstrҩ����װ & ",1)" & _
        " Order by Nvl(A.�۸񸸺�,A.���)"
        If strTime <> "" Then
            Set mrsDetail = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo, CDate(strTime))
        Else
            Set mrsDetail = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo)
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
            If Val(mshDetail.TextMatrix(i, mshDetail.COLS - 1)) = 3 Then
                mshDetail.Row = i
                For j = 0 To mshDetail.COLS - 1
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

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub


Private Function zlMsgModuleInit() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����Ϣģ��
    '���:lngModule -ģ���
    '     strPivs-Ȩ�޴�
    '����:objMsgModule-������Ϣ����
    '����:��ʼ���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-03-11 11:46:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    Set mobjMsgModule = New clsMipModule
    Call mobjMsgModule.InitMessage(glngSys, mlngModul, mstrPrivs)
    Call AddMipModule(mobjMsgModule)
    zlMsgModuleInit = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Public Function zlMsgModuleUnload() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ж��Ϣģ��
    '���:objMsgModule-��Ϣ����
    '����:���˺�
    '����:2014-03-11 11:46:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    
    If mobjMsgModule Is Nothing Then Exit Function
    Call mobjMsgModule.CloseMessage
    Call DelMipModule(mobjMsgModule)
    Set mobjMsgModule = Nothing
    zlMsgModuleUnload = False
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function


