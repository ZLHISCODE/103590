VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManageExecute 
   AutoRedraw      =   -1  'True
   Caption         =   "ִ�еǼǹ���"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11490
   Icon            =   "frmManageExecute.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   6195
   ScaleWidth      =   11490
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   5835
      Width           =   11490
      _ExtentX        =   20267
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmManageExecute.frx":08CA
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15187
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
      TabIndex        =   3
      Top             =   0
      Width           =   11490
      _ExtentX        =   20267
      _ExtentY        =   1376
      _CBWidth        =   11490
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinHeight1      =   720
      Width1          =   5910
      NewRow1         =   0   'False
      Child2          =   "picCondition"
      MinWidth2       =   3105
      MinHeight2      =   495
      Width2          =   3105
      NewRow2         =   0   'False
      Caption3        =   "����"
      Child3          =   "cboUnit"
      MinWidth3       =   1605
      MinHeight3      =   300
      Width3          =   1605
      NewRow3         =   0   'False
      Begin VB.PictureBox picCondition 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   6045
         ScaleHeight     =   495
         ScaleWidth      =   3105
         TabIndex        =   5
         Top             =   135
         Width           =   3105
         Begin VB.CheckBox chkAuto 
            Caption         =   "������һ��ѡ"
            Height          =   375
            Left            =   0
            TabIndex        =   7
            Top             =   48
            Width           =   855
         End
         Begin VB.TextBox txtValue 
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1680
            TabIndex        =   6
            ToolTipText     =   "��λF4"
            Top             =   55
            Width           =   1425
         End
         Begin VB.Label lblKind 
            Caption         =   "�����ݺ�"
            Height          =   225
            Left            =   885
            TabIndex        =   8
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.ComboBox cboUnit 
         Height          =   300
         Left            =   9795
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1605
      End
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   165
         TabIndex        =   4
         Top             =   30
         Width           =   5655
         _ExtentX        =   9975
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
            NumButtons      =   16
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
               Caption         =   "�Ǽ�"
               Key             =   "Log"
               Description     =   "�Ǽ�"
               Object.ToolTipText     =   "ִ�еǼ�"
               Object.Tag             =   "�Ǽ�"
               ImageKey        =   "Log"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ȡ��"
               Key             =   "Cancel"
               Description     =   "ȡ��"
               Object.ToolTipText     =   "ȡ���Ǽ�"
               Object.Tag             =   "ȡ��"
               ImageKey        =   "Cancel"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Edit_"
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�鿴"
               Key             =   "View"
               Description     =   "�鿴"
               Object.ToolTipText     =   "�鿴�Ǽ�"
               Object.Tag             =   "�鿴"
               ImageKey        =   "View"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "View_"
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ȫѡ"
               Key             =   "SelAll"
               Description     =   "ȫѡ"
               Object.ToolTipText     =   "ȫ��ѡ��"
               Object.Tag             =   "ȫѡ"
               ImageKey        =   "SelAll"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ȫ��"
               Key             =   "Clear"
               Description     =   "ȫ��"
               Object.ToolTipText     =   "ȫ�����"
               Object.Tag             =   "ȫ��"
               ImageKey        =   "Clear"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Clear_"
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Filter"
               Description     =   "����"
               Object.ToolTipText     =   "��������������ɸѡ��¼"
               Object.Tag             =   "����"
               ImageKey        =   "Filter"
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��λ"
               Key             =   "Go"
               Description     =   "��λ"
               Object.ToolTipText     =   "��λ�����������ļ�¼��"
               Object.Tag             =   "��λ"
               ImageKey        =   "Go"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "����"
               Object.ToolTipText     =   "��ǰ��������"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
      Height          =   4995
      Left            =   75
      TabIndex        =   0
      Top             =   795
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   8811
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      MergeCells      =   1
      AllowUserResizing=   1
      MouseIcon       =   "frmManageExecute.frx":0C3E
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
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
            Picture         =   "frmManageExecute.frx":0F58
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":1172
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":138C
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":15A6
            Key             =   "View"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":1D20
            Key             =   "Go"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":1F3A
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":2154
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":236E
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":2588
            Key             =   "Log"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":2C82
            Key             =   "Cancel"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":337C
            Key             =   "SelAll"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":3596
            Key             =   "Clear"
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
            Picture         =   "frmManageExecute.frx":37B0
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":39CA
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":3BE4
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":3DFE
            Key             =   "View"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":4578
            Key             =   "Go"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":4792
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":49AC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":4BC6
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":4DE0
            Key             =   "Log"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":54DA
            Key             =   "Cancel"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":5BD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":5DEE
            Key             =   ""
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
            Picture         =   "frmManageExecute.frx":6008
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
            Picture         =   "frmManageExecute.frx":68E2
            Key             =   ""
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
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSetup 
         Caption         =   "��������(&R)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_quit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditLog 
         Caption         =   "ִ�еǼ�(&A)"
      End
      Begin VB.Menu mnuEditCancel 
         Caption         =   "ȡ��ִ��(&C)"
      End
      Begin VB.Menu mnuEditView 
         Caption         =   "�鿴�Ǽ�(&V)"
      End
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelAll 
         Caption         =   "ȫ��ѡ��(&S)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditClear 
         Caption         =   "ȫ�����(&R)"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuEdit_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditPrint 
         Caption         =   "��ӡƱ��"
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
            Caption         =   "ִ�п���(&U)"
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
      Begin VB.Menu mnuViewShowHead 
         Caption         =   "����ͷ(&H)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuView_4 
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
   Begin VB.Menu mnuIDKind 
      Caption         =   "������"
      Visible         =   0   'False
      Begin VB.Menu mnuIDKinds 
         Caption         =   "���ݺ�"
         Index           =   0
      End
      Begin VB.Menu mnuIDKinds 
         Caption         =   "�����"
         Index           =   1
      End
      Begin VB.Menu mnuIDKinds 
         Caption         =   "סԺ��"
         Index           =   2
      End
      Begin VB.Menu mnuIDKinds 
         Caption         =   "����"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmManageExecute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private mrsList As ADODB.Recordset  '�����б�
Private mstrFilter As String, mstrPreNO As String
Private mblnGo As Boolean, mlngGo As Long
Private mlngCurRow As Long, mlngTopRow As Long
Private mlngDeptID As Long
Private Type Type_SQLCondition
    Default As Boolean          '�Ƿ���ȱʡ���룬��ʱû������ֵ,ȱʡֵ��mstrFilter��
    DateB As Date
    DateE As Date
    NOB As String
    NOE As String
    State As Byte
    Operator As String
    ID As Double
    Patient As String
End Type
Private SQLCondition As Type_SQLCondition

Private Const COL_���� = 14
Private Const COL_״̬ = 15

Private mstrPrivs As String     '���浱ǰģ�����Ȩ����
Private mlngModul As Long
Private mblnNOMoved As Boolean '��¼��ǰѡ��ĵ����Ƿ����ں����ݱ���

Private mrsWarn As ADODB.Recordset

Private Sub cboUnit_Click()
    
    If cboUnit.ItemData(cboUnit.ListIndex) = mlngDeptID Then Exit Sub
    mlngDeptID = cboUnit.ItemData(cboUnit.ListIndex)
        
    If Visible Then Call ShowBills(mstrFilter)
End Sub

Private Sub cbr_Resize()
    Form_Resize
End Sub

Private Sub chkAuto_Click()
    zlDatabase.SetPara "������Ŀͬʱѡ��", chkAuto.Value, glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
End Sub

Private Sub Form_Activate()
    Call InitLocPar(mlngModul)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub lblKind_Click()
    PopupMenu mnuIDKind, 2
End Sub

Private Sub mnuIDKinds_Click(Index As Integer)
    Dim i As Long
    
    For i = 0 To mnuIDKinds.UBound
        mnuIDKinds(i).Checked = i = Index
    Next
    
    lblKind.Caption = "��" & Choose(Index + 1, "���ݺ�", "�����", "סԺ��", "����")
End Sub


Private Sub txtvalue_KeyPress(KeyAscii As Integer)
    If mnuIDKinds(1).Checked Or mnuIDKinds(2).Checked Then
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub


Private Sub txtvalue_GotFocus()
    zlControl.TxtSelAll txtValue
End Sub

Private Sub txtValue_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtValue.Text <> "" Then
            If mnuIDKinds(0).Checked And IsNumeric(txtValue.Text) Then txtValue.Text = GetFullNO(txtValue.Text, 0)
        End If
        txtValue.Text = UCase(Trim(txtValue.Text))
        
        With frmExecuteFilter
            If mnuIDKinds(0).Checked Then
                .txtNOBegin.Text = txtValue.Text
                .txtNoEnd.Text = ""
                .txt��ʶ��.Text = ""
                .txt����.Text = ""
            ElseIf mnuIDKinds(1).Checked Or mnuIDKinds(2).Checked Then
                .txtNOBegin.Text = ""
                .txtNoEnd.Text = ""
                .txt��ʶ��.Text = txtValue.Text
                .txt����.Text = ""
            ElseIf mnuIDKinds(3).Checked Then
                .txtNOBegin.Text = ""
                .txtNoEnd.Text = ""
                .txt��ʶ��.Text = ""
                .txt����.Text = txtValue.Text
            End If
            .MakeFilter
        End With
        
        Call FindBills
        
        zlControl.TxtSelAll txtValue
    ElseIf KeyCode = vbKeyF4 Then
        Dim i As Integer
        
        For i = 0 To mnuIDKinds.Count - 1
            If mnuIDKinds(i).Checked = True Then Exit For
        Next
        If i >= mnuIDKinds.Count - 1 Then
            i = 0
        Else
            i = i + 1
        End If
        Call mnuIDKinds_Click(i)
    End If
End Sub
Private Function Is�������(ByVal lngRow As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ǰ��¼�Ƿ�Ϊ�������
    '��Σ�lngRow:ָ���е�����
    '���Σ�
    '���أ����������,����true,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-03-08 15:45:31
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim int�����־ As Integer, int��¼���� As Integer, int��ҳID As Integer
    int�����־ = GetColValue(lngRow, "�����־")
    int��¼���� = GetColValue(lngRow, "��¼����")
    int��ҳID = GetColValue(lngRow, "��ҳID")
    '1-����;2-סԺ;3-����(���￨�ȶ�����շ�);4-���
    If int�����־ = 1 Or int�����־ = 4 Or int��¼���� = 1 Then Is������� = True: Exit Function
    If int�����־ = 2 Then
        If Val(int��ҳID) = 0 Then
            Is������� = True: Exit Function
        End If
    End If
End Function
Private Sub mnuEditCancel_Click()
    Dim strNO As String, int���� As Integer, int��� As Integer
    Dim blnDo As Boolean, i As Long, blnTrans As Boolean
    Dim cllData As Collection, j As Long, blnFind As Boolean
    Dim varTemp As Variant, cllPro As Collection
    Dim strSQL As String
    If cboUnit.ListIndex = -1 Then Exit Sub
    'arrSQL = Array()
    Set cllData = New Collection
    For i = 1 To mshList.Rows - 1
        If Val(mshList.TextMatrix(i, 0)) <> -1 Then
            strNO = GetColValue(i, "���ݺ�")
            If strNO <> "" Then
                If mshList.TextMatrix(i, 1) <> "" And mshList.TextMatrix(i, GetColNum("ִ����")) <> "" Then
                    '�����������ʾ���ڴ򹴵���
                    int���� = GetColValue(i, "��¼����")
                    blnDo = True
                     '��ǰѡ��ĵ����б���ܲ�ֹһ��,���Բ���ȡ֮ǰȷ�����Ƿ��ں󱸱�ı��,��Ҫ���ж�
                    '�Ƿ���ת������ݱ���
                    If frmExecuteFilter.mblnDateMoved Then
                        If zlDatabase.NOMoved(IIf(Is�������(i), "������ü�¼", "סԺ���ü�¼"), strNO, , int����, Me.Caption) Then
                            If Not ReturnMovedExes(strNO, int����, Me.Caption) Then blnDo = False
                            'mblnNOMoved = False  '�˾䲻��Ҫ,����Ӱ�첻ѡ�����
                        End If
                    End If
                
                    If blnDo Then
                        If InStr(mstrPrivs, ";ȡ�����˵Ǽ�;") = 0 And mshList.TextMatrix(i, GetColNum("ִ����")) <> UserInfo.���� Then
                            mshList.TopRow = i
                            mshList_LeaveCell
                            mshList.Row = i
                            mshList_EnterCell
                            MsgBox strNO & " ����Ŀ """ & mshList.TextMatrix(i, GetColNum("��Ŀ")) & """ ��ִ����Ϊ�����ˣ���û��Ȩ��ȡ���Ǽǣ�", vbInformation, gstrSysName
                            blnDo = False
                        End If
                    
                        '����SQL
                        int��� = Val(mshList.TextMatrix(i, 0))
                        
                        blnFind = False
                        For j = 1 To cllData.Count
                            varTemp = cllData(j)
                            If varTemp(0) = strNO And Val(varTemp(1)) = int���� Then
                                cllData.Remove j
                                cllData.Add Array(strNO, int����, varTemp(2) & "," & int���, IIf(Is�������(i), 1, 2))
                                blnFind = True: Exit For
                            End If
                        Next
                        If blnFind = False Then
                             cllData.Add Array(strNO, int����, "" & int���, IIf(Is�������(i), 1, 2))
                        End If
                        
'                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
'                        arrSQL(UBound(arrSQL)) = "zl_���˷��ü�¼_UNExecute('" & strNO & "'," & int���� & "," & int��� & "," & IIf(Is�������(i), 1, 2) & ")"
                    End If
                    
                    blnDo = True  '��ʾ��ǰ�б��д��ڴ򹴵���
                End If
            End If
        End If
    Next
    Set cllPro = New Collection
    For j = 1 To cllData.Count
        'NO,����,���(���),�����־
        varTemp = cllData(j)
        strSQL = "zl_���˷��ü�¼_UNExecute('" & varTemp(0) & "'," & Val(varTemp(1)) & ",'" & varTemp(2) & "'," & Val(varTemp(3)) & ")"
        Call zlAddArray(cllPro, strSQL)
    Next
    
    If blnDo Then  '��ʾ���ڴ򹴵���
        If cllPro.Count = 0 Then Exit Sub '������ڴ򹴵�,����ȫ���ں󱸱������˳�
        If MsgBox("ȷʵҪ��ѡ��ļ�¼ȫ��ȡ���Ǽ���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Else
        'û��ѡ��,��ֻ����ǰ��
        If MsgBox("ȷʵҪ����ǰ��¼ȡ���Ǽ���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        strNO = GetColValue(mshList.Row, "���ݺ�")
        If mshList.Row = 0 Or strNO = "" Then Exit Sub
        
        int��� = Val(mshList.TextMatrix(mshList.Row, 0))
        If int��� = -1 Then Exit Sub
        int���� = GetColValue(mshList.Row, "��¼����")
        
        '���û�д�,ֻ�е�ǰ�е����,�Ƿ���ת������ݱ���
        If mblnNOMoved Then
            If Not ReturnMovedExes(strNO, int����, Me.Caption) Then Exit Sub
            mblnNOMoved = False  '��ʱ��ת���������ݱ�
        End If
        
        If InStr(mstrPrivs, ";ȡ�����˵Ǽ�;") = 0 And mshList.TextMatrix(mshList.Row, GetColNum("ִ����")) <> UserInfo.���� Then
            MsgBox "��ǰ��Ŀ """ & mshList.TextMatrix(mshList.Row, GetColNum("��Ŀ")) & """ ��ִ����Ϊ�����ˣ���û��Ȩ��ȡ���Ǽǣ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        strSQL = "zl_���˷��ü�¼_UNExecute('" & strNO & "'," & int���� & "," & int��� & "," & IIf(Is�������(mshList.Row), 1, 2) & ")"
        Call zlAddArray(cllPro, strSQL)
    End If
            
    Screen.MousePointer = 11
    On Error GoTo errH
    zlExecuteProcedureArrAy cllPro, Me.Caption
    On Error GoTo 0
    Screen.MousePointer = 0
    mnuViewReFlash_Click
    Exit Sub
errH:
    Screen.MousePointer = 0
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditClear_Click()
    Dim i As Long, j As Long
    j = GetColNum("���ݺ�")
    For i = 1 To mshList.Rows - 1
        If mshList.TextMatrix(i, j) <> "" And Val(mshList.TextMatrix(i, 0)) <> -1 Then
            mshList.TextMatrix(i, 1) = ""
        End If
    Next
End Sub

Private Sub mnuEditLog_Click()
    Dim strNO As String, bytFlag As Byte, int��� As Integer
    Dim strOper As String, strLog As String, lng����ID As Long, lng��ҳID As Long
    Dim arrSQL() As Variant, i As Long
    Dim arrPar() As Variant, blnPrint As Boolean
    Dim strInfo As String, blnDo As Boolean, blnTrans As Boolean, blnCheck As Boolean
    
    If cboUnit.ListIndex = -1 Then Exit Sub
    
    arrSQL() = Array()
    arrPar() = Array()
    
    'a.����ִ�д���
    For i = 1 To mshList.Rows - 1
        If Val(mshList.TextMatrix(i, 0)) <> -1 Then
            strNO = GetColValue(i, "���ݺ�")
            
            If strNO <> "" And mshList.TextMatrix(i, 1) <> "" Then
                '�����������ʾ���ڴ򹴵���
                
                bytFlag = Val(GetColValue(i, "��¼����"))
            
                blnDo = True
                '��ǰѡ��ĵ����б���ܲ�ֹһ��,���Բ���ȡ֮ǰȷ�����Ƿ��ں󱸱�ı��,��Ҫ���ж�
                '�Ƿ���ת������ݱ���
                'ɸѡʱ��ʱ�������һ��ת��֮ǰ
                If frmExecuteFilter.mblnDateMoved Then
                    If zlDatabase.NOMoved(IIf(Is�������(i), "������ü�¼", "סԺ���ü�¼"), strNO, , bytFlag, Me.Caption) Then
                        If Not ReturnMovedExes(strNO, bytFlag, Me.Caption) Then blnDo = False
                        'mblnNOMoved = False '�˾䲻��Ҫ,����Ӱ�첻ѡ�����
                    End If
                End If
                
                If blnDo Then
                    int��� = Val(mshList.TextMatrix(i, 0))
                    '������ִ�еģ���ȡ��һ��ִ���˺͵Ǽ�������Ϊ����ִ�вο�ֵ
                    If strOper = "" Then strOper = mshList.TextMatrix(mshList.Row, GetColNum("ִ����"))
                    
                    If strLog = "" Then strLog = GetItemLog(IIf(Is�������(i), 1, 2), strNO, bytFlag, int���) 'ǰ���Ѵ���Ϊ���߱�,�˴����ش�mblnNOMoved
                    
                    If gblnִ�к���� And GetColValue(i, "��¼����") = 2 And GetColValue(i, "��¼״̬") = 0 Then
                        If AuditingWarn(mstrPrivs, mrsWarn, strNO, int���) Then
                            If lng����ID <> Val(GetColValue(i, "����ID")) Then
                                lng����ID = Val(GetColValue(i, "����ID"))
                                lng��ҳID = Val(GetColValue(i, "��ҳID"))
                                blnCheck = PatiCanBilling(lng����ID, lng��ҳID, mstrPrivs)
                            End If
                        Else
                            blnCheck = False
                        End If
                    Else
                        blnCheck = True
                    End If
                            
                    If blnCheck Then
                        '1.Ʊ�ݴ�ӡ����
                        ReDim Preserve arrPar(UBound(arrPar) + 1)
                        arrPar(UBound(arrPar)) = strNO & "," & bytFlag & "," & int���
                                        
                        '2.����SQL
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "zl_���˷��ü�¼_Execute('" & strNO & "'," & bytFlag & "," & int��� & "," & IIf(Is�������(i), 1, 2) & ","
                    End If
                End If
                
                blnDo = True  '��ʾ��ǰ�б��д��ڴ򹴵���
            End If
        End If
    Next
    
    If blnDo Then  '��ʾ���ڴ򹴵���
        If UBound(arrSQL) < 0 And blnDo Then Exit Sub   '������ڴ򹴵�,����ȫ���ں󱸱������˳�
    
        If MsgBox("ȷʵҪ��ѡ��ļ�¼ȫ�����еǼ���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    '�����ڴ򹴵��У���ֻ����ǰ��
    Else
        strNO = GetColValue(mshList.Row, "���ݺ�")
        If mshList.Row = 0 Or strNO = "" Then Exit Sub
        
        int��� = Val(mshList.TextMatrix(mshList.Row, 0))
        If int��� = -1 Then Exit Sub
        bytFlag = GetColValue(mshList.Row, "��¼����")
        strOper = mshList.TextMatrix(mshList.Row, GetColNum("ִ����"))
        strLog = GetItemLog(IIf(Is�������(mshList.Row), 1, 2), strNO, bytFlag, int���)
        
        '���û�д�,ֻ�е�ǰ�е����,�Ƿ���ת������ݱ���
        If mblnNOMoved Then
            If Not ReturnMovedExes(strNO, bytFlag, Me.Caption) Then Exit Sub
            mblnNOMoved = False  '��ʱ��ת���������ݱ�
        End If
        
        blnCheck = True
        If gblnִ�к���� And GetColValue(mshList.Row, "��¼����") = 2 And GetColValue(mshList.Row, "��¼״̬") = 0 Then
            If AuditingWarn(mstrPrivs, mrsWarn, strNO, int���) Then
                lng����ID = Val(GetColValue(mshList.Row, "����ID"))
                lng��ҳID = Val(GetColValue(mshList.Row, "��ҳID"))
                blnCheck = PatiCanBilling(lng����ID, lng��ҳID, mstrPrivs)
            Else
                blnCheck = False
            End If
        End If
        
        If blnCheck Then
            'Ʊ�ݴ�ӡ����
            ReDim arrPar(0)
            arrPar(0) = strNO & "," & bytFlag & "," & int���
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "zl_���˷��ü�¼_Execute('" & strNO & "'," & bytFlag & "," & int��� & "," & IIf(Is�������(mshList.Row), 1, 2) & ","
        Else
            Exit Sub
        End If
    End If
    
    
    On Error Resume Next
    frmExeEdit.mlngDeptID = cboUnit.ItemData(cboUnit.ListIndex)
    frmExeEdit.mstrOper = strOper
    frmExeEdit.mstrLog = strLog
    frmExeEdit.Show 1, Me
    If gblnOK Then
        For i = 0 To UBound(arrSQL)
            With frmExeEdit
                arrSQL(i) = arrSQL(i) & "'" & .mstrLog & "','" & .mstrOper & "',To_Date('" & Format(.mvDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
            End With
        Next
                
        Screen.MousePointer = 11
        On Error GoTo errH
        gcnOracle.BeginTrans: blnTrans = True
            For i = 0 To UBound(arrSQL)
                Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
            Next
        gcnOracle.CommitTrans: blnTrans = False
        On Error GoTo 0
        Call mshList_EnterCell
        Screen.MousePointer = 0
        
        '��ӡƱ��
        blnPrint = False
        If gbytExe��ӡ��ʽ = 1 Then
            blnPrint = True
        ElseIf gbytExe��ӡ��ʽ = 2 Then
            If UBound(arrPar) > 0 Then
                strInfo = "ִ�еǼ����,Ҫ��ӡ�ղ�ѡ�������ִ�еǼǵ���"
            Else
                strInfo = "ִ�еǼ����,Ҫ��ӡִ�еǼǵ���"
            End If
            blnPrint = MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
        End If
        
        If blnPrint Then
            Screen.MousePointer = 11
            For i = 0 To UBound(arrPar)
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1142", Me, _
                    "NO=" & Split(arrPar(i), ",")(0), _
                    "��¼����=" & Split(arrPar(i), ",")(1), _
                    "���=" & Split(arrPar(i), ",")(2), 2)
            Next
            Screen.MousePointer = 0
        End If
        
        'ˢ��
        mnuViewReFlash_Click
    End If
    Exit Sub
errH:
    Screen.MousePointer = 0
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditPrint_Click()
    Dim strNO As String, int���� As Integer, int��� As Integer
    Dim arrPar() As Variant, i As Long
    Dim blnDo As Boolean
    
    If cboUnit.ListIndex = -1 Then Exit Sub
    
    arrPar() = Array()
    For i = 1 To mshList.Rows - 1
        If Val(mshList.TextMatrix(i, 0)) <> -1 Then
            strNO = GetColValue(i, "���ݺ�")
            If strNO <> "" And mshList.TextMatrix(i, 1) <> "" Then
                int��� = Val(mshList.TextMatrix(i, 0))
                int���� = GetColValue(i, "��¼����")
                
                blnDo = True
            
                '��ǰѡ��ĵ����б���ܲ�ֹһ��,���Բ���ȡ֮ǰȷ�����Ƿ��ں󱸱�ı��,��Ҫ���ж�
                '�Ƿ���ת������ݱ���
                If frmExecuteFilter.mblnDateMoved Then
                    If zlDatabase.NOMoved(IIf(Is�������(i), "������ü�¼", "סԺ���ü�¼"), strNO, , int����, Me.Caption) Then
                        If Not ReturnMovedExes(strNO, int����, Me.Caption) Then blnDo = False
                        mblnNOMoved = False
                    End If
                End If
                
                If blnDo Then
                    'Ʊ�ݴ�ӡ����
                    ReDim Preserve arrPar(UBound(arrPar) + 1)
                    arrPar(UBound(arrPar)) = strNO & "," & int���� & "," & int���
                End If
            End If
        End If
    Next
    
    If UBound(arrPar) >= 0 Then
        If MsgBox("ȷʵҪ��ѡ��ļ�¼ȫ�����д�ӡ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Else
        strNO = GetColValue(mshList.Row, "���ݺ�")
        If mshList.Row = 0 Or strNO = "" Then Exit Sub
        int��� = Val(mshList.TextMatrix(mshList.Row, 0))
        If int��� = -1 Then Exit Sub
        int���� = GetColValue(mshList.Row, "��¼����")
        
        'Ʊ�ݴ�ӡ����
        ReDim arrPar(0)
        arrPar(0) = strNO & "," & int���� & "," & int���
    End If
    
    If Not ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1142", Me) Then Exit Sub
    
    Screen.MousePointer = 11
    For i = 0 To UBound(arrPar)
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1142", Me, _
            "NO=" & Split(arrPar(i), ",")(0), _
            "��¼����=" & Split(arrPar(i), ",")(1), _
            "���=" & Split(arrPar(i), ",")(2), 2)
    Next
    Screen.MousePointer = 0
    
    mnuViewReFlash_Click
    'Call mshList_EnterCell 'ˢ������ִ��
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditSelAll_Click()
    Dim i As Long, j As Long
    j = GetColNum("���ݺ�")
    For i = 1 To mshList.Rows - 1
        If mshList.TextMatrix(i, j) <> "" And Val(mshList.TextMatrix(i, 0)) <> -1 Then
            mshList.TextMatrix(i, 1) = "��"
        End If
    Next
End Sub

Private Sub mnuEditView_Click()
    Dim str���ݺ� As String, bytFlag As Byte, int��� As Integer
    Dim strOper As String, strLog As String, strDate As String
    
    If cboUnit.ListIndex = -1 Then Exit Sub
    
    str���ݺ� = GetColValue(mshList.Row, "���ݺ�")
    If mshList.Row = 0 Or str���ݺ� = "" Then Exit Sub
    
    int��� = Val(mshList.TextMatrix(mshList.Row, 0))
    If int��� = -1 Then Exit Sub
    
    bytFlag = GetColValue(mshList.Row, "��¼����")
    strDate = GetColValue(mshList.Row, "ִ��ʱ��")
    
    strOper = mshList.TextMatrix(mshList.Row, GetColNum("ִ����"))
    strLog = GetItemLog(IIf(Is�������(mshList.Row), 1, 2), str���ݺ�, bytFlag, int���, mblnNOMoved)
    
    frmExeEdit.mblnView = True
    frmExeEdit.mlngDeptID = cboUnit.ItemData(cboUnit.ListIndex)
    frmExeEdit.mstrOper = strOper
    frmExeEdit.mstrLog = strLog
    frmExeEdit.mstrDate = strDate
    
    frmExeEdit.Show 1, Me
End Sub

Private Sub mnuFileSetup_Click()
    frmExecuteSet.mlngModul = mlngModul
    frmExecuteSet.mstrPrivs = mstrPrivs
    frmExecuteSet.Show 1, Me
    If frmExecuteSet.mblnOK Then
        mnuViewReFlash_Click
    End If
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim strNO As String
    
    strNO = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    If strNO = "" Then
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
            "���˿���=" & mlngDeptID)
    Else
        With mshList
            Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
                "���˿���=" & mlngDeptID, "NO=" & strNO, _
                "����ID=" & .TextMatrix(.Row, GetColNum("����ID")), _
                "��ҳID=" & .TextMatrix(.Row, GetColNum("��ҳID")), _
                "סԺ��=" & .TextMatrix(.Row, GetColNum("סԺ��")), _
                "������=" & .TextMatrix(.Row, GetColNum("������")))
        End With
    End If
End Sub

Private Sub mnuViewFilter_Click()
    
    If frmExecuteFilter.mlngDept <> mlngDeptID Then
        frmExecuteFilter.mlngDept = mlngDeptID
        frmExecuteFilter.LoadOper
    End If
    
    frmExecuteFilter.Show 1, Me
    If gblnOK Then Call FindBills
End Sub

Private Sub FindBills()
    With frmExecuteFilter
        mstrFilter = .mstrFilter
        
        SQLCondition.Default = False
        SQLCondition.DateB = .dtpBegin.Value
        SQLCondition.DateE = .dtpEnd.Value
        SQLCondition.NOB = .txtNOBegin.Text
        SQLCondition.NOE = .txtNoEnd.Text
        If .cbo״̬.Text <> "����״̬" Then SQLCondition.State = Val(.cbo״̬.Text) - 1
        SQLCondition.Operator = zlStr.NeedName(.cboִ����.Text)
        SQLCondition.ID = Val(.txt��ʶ��.Text)
        SQLCondition.Patient = gstrLike & UCase(.txt����.Text) & "%"
    End With
    
    mnuViewReFlash_Click
End Sub

Private Sub mnuViewShowHead_Click()
    mnuViewShowHead.Checked = Not mnuViewShowHead.Checked
    Call SetBillHead(False)
    Call SetHeader
End Sub

Private Sub mshList_DblClick()
Dim i As Integer
Dim bln���� As Boolean
Dim strNO As String

'�㷨:�������ǰ����ѡ���ѡͬһ���ݵ���ϸ��,ͬһ������ͬ�������ϸ��

    If mshList.MouseRow = 0 Then Exit Sub
    If mshList.MouseCol = GetColNum("ѡ��") Then
        If Val(mshList.TextMatrix(mshList.Row, 0)) <> -1 Then                  '�����е�ֵΪ-1
            '1.�������ϸ��˫��
            strNO = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))  '�Ըõ��ݵ�һ����ϸ�ĵ��ݺ�Ϊ׼
            If strNO <> "" Then  '������ʱ�Ż��
                '���������,��ѡ���ѡ�������ж���,����Ǵ���,��ѡ���ѡ���ĸ��׼������ֵ�
                '������
                If chkAuto.Value = 1 Then
                    If mshList.Row <> mshList.Rows - 1 Then
                        For i = mshList.Row + 1 To mshList.Rows - 1       '��ǰ���ڵ���ʱ��ѡ��
                            If mshList.TextMatrix(i, GetColNum("���ݺ�")) <> strNO Then Exit For
                            If mshList.TextMatrix(i, GetColNum("����")) = mshList.TextMatrix(mshList.Row, GetColNum("����")) Then
                                If mshList.TextMatrix(i, 1) = "" Then
                                    mshList.TextMatrix(i, 1) = "��"
                                Else
                                    mshList.TextMatrix(i, 1) = ""
                                End If
                            End If
                        Next
                    End If
                    '�ٵ���
                    For i = mshList.Row To 0 Step -1
                        If mshList.TextMatrix(i, GetColNum("���ݺ�")) <> strNO Then Exit For
                        If mshList.TextMatrix(i, GetColNum("����")) = mshList.TextMatrix(mshList.Row, GetColNum("����")) Then
                            If mshList.TextMatrix(i, 1) = "" Then
                                mshList.TextMatrix(i, 1) = "��"
                            Else
                                mshList.TextMatrix(i, 1) = ""
                            End If
                        End If
                    Next
                Else
                    If mshList.TextMatrix(mshList.Row, 1) = "" Then
                        mshList.TextMatrix(mshList.Row, 1) = "��"
                    Else
                        mshList.TextMatrix(mshList.Row, 1) = ""
                    End If
                End If
            End If
        Else
            '2.������ڵ�����˫��,��ѡ���ѡ�õ��ݵ�������ϸ��
            '�ȼ�������,�������,�����������һ��,һ���ǵ���
            strNO = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
            strNO = Mid(strNO, InStr(1, strNO, ")") + 1, InStr(1, strNO, " ") - InStr(1, strNO, ")") - 1)
            
            If mshList.Row <> mshList.Rows - 1 Then
                For i = mshList.Row + 1 To mshList.Rows - 1
                    If mshList.TextMatrix(i, GetColNum("���ݺ�")) <> strNO Then Exit For
                    If mshList.TextMatrix(i, 1) = "" Then
                        mshList.TextMatrix(i, 1) = "��"
                    Else
                        mshList.TextMatrix(i, 1) = ""
                    End If
                    bln���� = True
                Next
            End If
            If Not bln���� Then
                For i = mshList.Row - 1 To 0 Step -1 '����ǵ�һ��,��ǰ��϶���ִ��
                    If mshList.TextMatrix(i, GetColNum("���ݺ�")) <> strNO Then Exit For
                    If mshList.TextMatrix(i, 1) = "" Then
                        mshList.TextMatrix(i, 1) = "��"
                    Else
                        mshList.TextMatrix(i, 1) = ""
                    End If
                Next
            End If
        End If
    ElseIf mnuEditView.Enabled Then
        Call mnuEditView_Click
    ElseIf mnuEditLog.Visible And mnuEditLog.Enabled Then
        Call mnuEditLog_Click
    End If
End Sub

Private Sub mshList_EnterCell()
    Dim strNO As String, int��� As Integer, i As Long
    Dim intRows As Integer, bln As Boolean
    Dim lng����ID As Long, lngִ��ID As Long
    Dim blnִ�� As Boolean, lng����ID As Long
    Dim bytFlag As Byte
    
    strNO = GetColValue(mshList.Row, "���ݺ�")
    If mshList.Row = 0 Or strNO = "" Then Exit Sub
    
    mlngGo = mshList.Row
    mlngCurRow = mshList.Row: mlngTopRow = mshList.TopRow
    
    
    bln = mshList.Redraw
    mshList.Redraw = False
    '���ñ���ɫ
    For i = 0 To mshList.Cols - 1
        mshList.Col = i
        mshList.CellBackColor = mshList.BackColorSel
        mshList.CellForeColor = mshList.ForeColorSel
    Next
    mshList.Col = 0
    
    '���ö���
    intRows = (mshList.Height - mshList.RowHeight(0) - 60) \ 250
    If mshList.TopRow > mshList.Row Then
        mshList.TopRow = mshList.Row
    ElseIf mshList.Row - mshList.TopRow >= intRows Then
        mshList.TopRow = mshList.Row - intRows + 1
    End If
    
    mshList.Redraw = bln
    
    int��� = Val(mshList.TextMatrix(mshList.Row, 0))
    blnִ�� = (GetColValue(mshList.Row, "״̬") = "��ִ��")
    
    
    mnuEditLog.Enabled = int��� <> -1 And Not blnִ��
    tbr.Buttons("Log").Enabled = mnuEditLog.Enabled
    
    mnuEditCancel.Enabled = int��� <> -1 And blnִ��
    tbr.Buttons("Cancel").Enabled = mnuEditCancel.Enabled
    
    mnuEditView.Enabled = mnuEditCancel.Enabled
    tbr.Buttons("View").Enabled = mnuEditCancel.Enabled
    
    mnuEditPrint.Enabled = mnuEditLog.Enabled
    
    If int��� = -1 Then
        stbThis.Panels(2) = stbThis.Tag
        mblnNOMoved = False
    Else
        bytFlag = GetColValue(mshList.Row, "��¼����")
        If frmExecuteFilter.mblnDateMoved Then
            mblnNOMoved = zlDatabase.NOMoved(IIf(Is�������(mshList.Row), "������ü�¼", "סԺ���ü�¼"), strNO, , bytFlag, Me.Caption)
        Else
            mblnNOMoved = False
        End If
        
        stbThis.Panels(2) = "ִ�����:" & GetItemLog(IIf(Is�������(mshList.Row), 1, 2), strNO, bytFlag, int���, mblnNOMoved)
    End If
End Sub

Private Sub mshList_GotFocus()
    mshList.BackColorSel = &H8000000D
    Call mshList_EnterCell
End Sub

Private Sub mshList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If mnuEditView.Enabled Then
            mnuEditView_Click
        ElseIf mnuEditLog.Enabled And mnuEditLog.Visible Then
            mnuEditLog_Click
        End If
    ElseIf KeyAscii = Asc(" ") Then
        If Val(mshList.TextMatrix(mshList.Row, 0)) <> -1 Then
            If mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�")) <> "" Then
                If mshList.TextMatrix(mshList.Row, 1) = "" Then
                    mshList.TextMatrix(mshList.Row, 1) = "��"
                Else
                    mshList.TextMatrix(mshList.Row, 1) = ""
                End If
            End If
        End If
    End If
End Sub

Private Sub mshList_LeaveCell()
    Dim i As Long
    Dim bln As Boolean
    
    '���ñ���ɫ
    bln = mshList.Redraw
    mshList.Redraw = False
    For i = 0 To mshList.Cols - 1
        mshList.Col = i
        If Val(mshList.TextMatrix(mshList.Row, 0)) = -1 Then
            mshList.CellBackColor = &HEBFFFF '&HE6FFFF '&HE0E0E0
        Else
            mshList.CellBackColor = mshList.BackColor
        End If
        mshList.CellForeColor = mshList.ForeColor
    Next
    mshList.Redraw = bln
End Sub

Private Sub mshList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If Button = 2 Then PopupMenu mnuEdit, 2
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF4
            If Not ActiveControl Is txtValue Then Call txtValue.SetFocus
        Case vbKeyF3
            'ʼ�մӵ�ǰ�п�ʼ
            If mnuViewGo.Enabled Then Call SeekBill(False)
        Case vbKeyEscape
            mblnGo = False
    End Select
End Sub

Private Sub mnuHelpTitle_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
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

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuFile_quit_Click
        Case "Go" '��λ
            mnuViewGo_Click
        Case "Filter" '����
            mnuViewFilter_Click
        Case "Log"
            mnuEditLog_Click
        Case "Cancel"
            mnuEditCancel_Click
        Case "View"
            mnuEditView_Click
        Case "SelAll"
            mnuEditSelAll_Click
        Case "Clear"
            mnuEditClear_Click
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
    objOut.Title.Text = "ҽ�������嵥"
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '����
    With frmExecuteFilter
        objRow.Add "ʱ�䣺" & Format(.dtpBegin.Value, .dtpBegin.CustomFormat) & " �� " & Format(.dtpEnd.Value, .dtpEnd.CustomFormat)
        objOut.UnderAppRows.Add objRow
    End With
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate(), "yyyy��MM��dd��")
    objOut.BelowAppRows.Add objRow
    
    '����
    mshList.Redraw = False
    mshList_LeaveCell
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
    mshList_EnterCell
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
    
    mnuEditLog.Enabled = blnUsed
    tbr.Buttons("Log").Enabled = blnUsed
    
    mnuEditCancel.Enabled = blnUsed
    tbr.Buttons("Cancel").Enabled = blnUsed
        
    mnuEditView.Enabled = blnUsed
    tbr.Buttons("View").Enabled = blnUsed
        
    mnuEditPrint.Enabled = blnUsed
    
    mnuViewGo.Enabled = blnUsed
    tbr.Buttons("Go").Enabled = blnUsed
End Sub

Private Sub Form_Load()
    mstrPrivs = gstrPrivs
    mlngModul = glngModul
    
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    Call RestoreWinState(Me, App.ProductName)
    mnuViewShowHead.Checked = zlDatabase.GetPara("��ʾ����ͷ", glngSys, mlngModul, "1") = "1"
    
    mlngCurRow = 1: mlngTopRow = 1
    
    'Ȩ������
    If InStr(mstrPrivs, ";ִ�еǼ�;") = 0 And InStr(mstrPrivs, ";ȡ���Ǽ�;") = 0 Then
        mnuEditLog.Visible = False
        mnuEditCancel.Visible = False
        
        tbr.Buttons("Log").Visible = False
        tbr.Buttons("Cancel").Visible = False
        tbr.Buttons("Edit_").Visible = False
    ElseIf InStr(mstrPrivs, ";ִ�еǼ�;") = 0 Then
        mnuEditLog.Visible = False
        tbr.Buttons("Log").Visible = False
    ElseIf InStr(mstrPrivs, ";ȡ���Ǽ�;") = 0 Then
        mnuEditCancel.Visible = False
        tbr.Buttons("Cancel").Visible = False
    End If
    
    '������Ŀͬʱѡ��
    chkAuto.Value = IIf(zlDatabase.GetPara("������Ŀͬʱѡ��", glngSys, mlngModul, "0") = "1", 1, 0)
    mnuIDKinds(0).Checked = True '���ݺ�
    
    If gblnִ�к���� Then Set mrsWarn = GetUnitWarn
    
    '����
    If Not InitUnits Then Unload Me: Exit Sub
    If cboUnit.ListIndex = -1 Then
        MsgBox "û�з�������������,���㲻�������п���Ȩ��,����ʹ��ҽ�����Ҽ��ʣ�", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    
    Call SetHeader
    Call SetMenu(False)
    stbThis.Panels(2).Text = "��ˢ���嵥���������ù�������"
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long, staH As Long

    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    
    mshList.MousePointer = 0
    
    '����ؼ���Ⱥ͸߶�
    cbrH = IIf(cbr.Visible, cbr.Height, 0)
    staH = IIf(stbThis.Visible, stbThis.Height, 0)
    
    mshList.Left = 0
    mshList.Top = cbrH
    mshList.Width = Me.ScaleWidth
    mshList.Height = Me.ScaleHeight - cbrH - staH
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrFilter = ""
    mlngDeptID = 0
    mstrPreNO = ""
    
    Unload frmExecuteFilter
    Unload frmExecuteGo
    Call SaveWinState(Me, App.ProductName)
    zlDatabase.SetPara "��ʾ����ͷ", IIf(mnuViewShowHead.Checked, 1, 0), glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
End Sub

Private Sub mnuViewGo_Click()
    frmExecuteGo.Show 1, Me
    mstrPreNO = ""
    If gblnOK Then Call SeekBill(frmExecuteGo.optHead)
End Sub

Private Sub SeekBill(blnHead As Boolean)
    Dim i As Long, bln As Boolean, intRows As Integer
    Dim blnFill As Boolean, j As Long
    Dim strCurNO As String
    
    Screen.MousePointer = 11
    mblnGo = True
    stbThis.Panels(2).Text = "���ڶ�λ���������ĵ���,��ESC��ֹ ..."
    Me.Refresh
    
    For i = IIf(blnHead, 1, mlngGo) To mshList.Rows - 1
        DoEvents

        strCurNO = GetColValue(i, "���ݺ�")
        
        If Val(mshList.TextMatrix(i, 0)) <> -1 Then
            '�Ƚ�����
            blnFill = True
            With frmExecuteGo
                If .txtNO.Text <> "" Then
                    blnFill = blnFill And strCurNO = .txtNO.Text
                End If
                If .txt��ʶ��.Text <> "" Then
                    blnFill = blnFill And GetColValue(i, frmExecuteGo.lbl��ʶ��.Caption) = .txt��ʶ��.Text
                End If
                If .txt����ID.Text <> "" Then
                    blnFill = blnFill And GetColValue(i, "����ID") = .txt����ID.Text
                End If
                If .txt����.Text <> "" Then
                    blnFill = blnFill And UCase(GetColValue(i, "����")) Like "*" & UCase(.txt����.Text) & "*"
                End If
            End With
            blnFill = blnFill And (strCurNO <> mstrPreNO)
            
            '�������˳�
            If blnFill Then
                mstrPreNO = strCurNO
                
                mlngGo = i + 1
    
                'LeaveCell����
                bln = mshList.Redraw
                mshList.Redraw = False
                For j = 0 To mshList.Cols - 1
                    mshList.Col = j
                    If Val(mshList.TextMatrix(mshList.Row, 0)) = -1 Then
                        mshList.CellBackColor = &HEBFFFF '&HE6FFFF '&HE0E0E0
                    Else
                        mshList.CellBackColor = mshList.BackColor
                    End If
                    mshList.CellForeColor = mshList.ForeColor
                Next
                '''''''''''''''''''''
    
                mshList.Row = i
                mshList.Col = 0: mshList.ColSel = mshList.Cols - 1
    
                'EnterCell����
                For j = 0 To mshList.Cols - 1
                    mshList.Col = j
                    mshList.CellBackColor = mshList.BackColorSel
                    mshList.CellForeColor = mshList.ForeColorSel
                Next
                intRows = (mshList.Height - mshList.RowHeight(0) - 60) \ 250
                If mshList.TopRow > mshList.Row Then
                    mshList.TopRow = mshList.Row
                ElseIf mshList.Row - mshList.TopRow >= intRows Then
                    mshList.TopRow = mshList.Row - intRows + 2
                End If
                mshList.Redraw = bln
                ''''''''''''''''''''
                
                stbThis.Panels(2).Text = "�ҵ�һ�ŵ���"
                Screen.MousePointer = 0: Exit Sub
            End If
        End If
        
        '��ESCȡ��
        If mblnGo = False Then
            stbThis.Panels(2).Text = "�û�ȡ����λ����"
            Screen.MousePointer = 0: Exit Sub
        End If
    Next
    mlngGo = 1: mstrPreNO = ""
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
    If lngCol = GetColNum("ѡ��") Then Exit Sub
    
    If Button = 1 And mshList.MousePointer = 99 Then
                 
        If mshList.TextMatrix(0, lngCol) = "" Then Exit Sub
        If GetColValue(mshList.Row, "���ݺ�") = "" Then Exit Sub
        If mrsList Is Nothing Then Exit Sub
        
        Set mshList.DataSource = Nothing
        
        If mshList.TextMatrix(0, lngCol) = "סԺ��" Or mshList.TextMatrix(0, lngCol) = "�����" Then
            mrsList.Sort = "���ݺ� Desc,��ʶ��" & IIf(mshList.ColData(lngCol) = 0, "", " DESC")
        Else
            mrsList.Sort = "���ݺ� Desc," & mshList.TextMatrix(0, lngCol) & IIf(mshList.ColData(lngCol) = 0, "", " DESC")
        End If
        mshList.ColData(lngCol) = (mshList.ColData(lngCol) + 1) Mod 2
        
        Call ShowBills(, True)
    End If
End Sub

Private Sub SetHeader()
    Dim strHead As String
    Dim i As Long
    
    strHead = "���,7,0|ѡ��,4,450|���ݺ�,1,0|����,1,0|����ID,1,0|��ʶ��,1,0|����,1,0|" & _
        "����,1,1000|������,1,0|���,4,0|��Ŀ,1,3000" & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "|��Ʒ��,1,1600", "") & "|���,1,1000|����,1,850|����,7,850|Ӧ�ս��,7,850|ʵ�ս��,7,850|" & _
        "״̬,1,650|ִ����,1,700|ִ��ʱ��,4,2000|����Ա,1,700|�Ǽ�ʱ��,4,2000|��¼����,1,0|�����־,1,0|����,1,0|סԺ��,1,0|��ҳID,1,0|��¼״̬,1,0"
    With mshList
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(mshList, App.ProductName & "\" & Me.Name)
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
        
        mshList.ColWidth(0) = 0 '���
        mshList.ColWidth(2) = 0 'NO
        mshList.ColWidth(3) = 0 '����
        mshList.ColWidth(4) = 0 '����ID
        mshList.ColWidth(5) = 0 '��ʶ��
        mshList.ColWidth(6) = 0 '����
        mshList.ColWidth(8) = 0 '������
        mshList.ColWidth(9) = 0 '���
        mshList.ColWidth(mshList.Cols - 3) = 0 '��¼����
        mshList.ColWidth(mshList.Cols - 2) = 0 '�����־
        mshList.ColWidth(mshList.Cols - 1) = 0 '����
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
    End With
End Sub

Private Sub ShowBills(Optional ByVal strIF As String, Optional blnSort As Boolean)
'����:��������ȡ�����б�(���˹���)
'����:strIF=��"AND"��ʼ��������
'     blnSort=�����¶�ȡ����,��������ʾ�����������
    Dim i As Long, Curdate As Date
    Dim strDept As String, strSQL As String, strType As String
    Dim bytType As Byte '0-����,1-סԺ,2-�����סԺ
    On Error GoTo errH
        
    If Not blnSort Then
        Call zlCommFun.ShowFlash("���ڶ�ȡ�����б�,���Ժ� ...", Me)
        DoEvents
        Me.Refresh
        LockWindowUpdate Me.hWnd
        
        'ȡȱʡ����
        If strIF = "" Then
            strIF = " And �Ǽ�ʱ�� Between Trunc(sysdate-3) And Trunc(sysdate+1)-1/24/60/60"
            strIF = strIF & " And Nvl(ִ��״̬,0)=0"
        Else
            strIF = strIF & " And Nvl(ִ��״̬,0)<>9"   '����:44510
        End If
        
        strIF = strIF & " And ִ�в���ID+0=[9]"
        If gstrExe��� = "" Then
            strIF = strIF & " And �շ���� Not IN('1','5','6','7','J')"
        Else
            strIF = strIF & " And �շ���� IN(" & gstrExe��� & ")"
        End If
        
        '000:����;סԺ;���
        If gstrExe��Դ <> "000" Then '��"111"��Ϊ000,��Ϊȫѡʱ,�������Ŀ��һ��ѡ�� :30493
            If Mid(gstrExe��Դ, 1, 1) = 1 Then strType = " �����־=1 And ��¼״̬ in(" & IIf(gbytExe���ﵥ������ = 2, "0,1", gbytExe���ﵥ������) & ")"
            If Mid(gstrExe��Դ, 2, 1) = 1 Then strType = IIf(strType <> "", strType & " Or", "") & " �����־=2 And ��¼״̬ in(" & IIf(gbytExeסԺ�������� = 2, "0,1", gbytExeסԺ��������) & ")"
            If Mid(gstrExe��Դ, 3, 1) = 1 Then strType = IIf(strType <> "", strType & " Or", "") & " �����־=4 And ��¼״̬ in(" & IIf(gbytExe��쵥������ = 2, "0,1", gbytExe��쵥������) & ")"
            If strType <> "" Then strIF = strIF & " And (" & strType & ")"
            If Mid(gstrExe��Դ, 2, 1) = 0 Then  '��סԺ,�϶������������
                bytType = 0
            ElseIf Mid(gstrExe��Դ, 1, 1) = 1 Or Mid(gstrExe��Դ, 3, 1) = 1 Then
                '�϶����������סԺ
                bytType = 2
            Else    'ֻ����סԺ���
                bytType = 1
            End If
           bytType = IIf(Mid(gstrExe��Դ, 2, 1) = 0, 0, 2)
        Else
            strIF = strIF & " And ��¼״̬ in(0,1)"
            bytType = 2
        End If
        
        
        strIF = strIF & IIf(Not gblnExeҽ��, "  And ҽ����� is NULL", "")
        
        If (gstrExe��� = "" Or InStr(gstrExe���, "'4'") > 0) And gblnִ�к��� = False Then
            strIF = strIF & " And (�շ���� <> '4' or �շ���� = '4' And Not Exists(Select 1 From �������� C Where A.�շ�ϸĿid = C.����id And C.�������� = 1))"
        End If
        '77838,Ƚ����,2014-9-16,ҽ�����˲����˷Ѻ����������ʣ��δִ�в��ֵļ�¼
        strIF = " Where ��¼����>0 And (��¼����<10 Or ��¼����=11) And ��¼����<>3 " & strIF
        
        Dim strTable As String
 
        strTable = "" & _
        "   Select  A.�۸񸸺�, A.���, A.��������, A.NO, A.����, A.����id,A.��ʶ��, A.�����־, A.����, A.��������id, A.������, A.�շ����, A.�շ�ϸĿid, " & _
        "           A.����, A.���㵥λ, A.��׼����, A.Ӧ�ս��, A.ʵ�ս��, A.ִ��״̬, A.ִ����, A.ִ��ʱ��, A.����Ա����, A.������, A.�Ǽ�ʱ��, A.��¼����, " & _
        "           A.�ಡ�˵�, A.��ҳid, ��¼״̬ " & _
        "  From סԺ���ü�¼ A " & _
           strIF
        If frmExecuteFilter.mblnDateMoved Then     'ɸѡʱ��ʱ�������һ��ת��֮ǰ
            strTable = strTable & " Union ALL " & Replace(strTable, "סԺ���ü�¼", "HסԺ���ü�¼")
        End If
        Select Case bytType
        Case 0  '����
            strTable = Replace(Replace(Replace(Replace(strTable, "סԺ���ü�¼", "������ü�¼"), "A.����", "'' as ����"), "A.�ಡ�˵�", " 0 as �ಡ�˵�"), "A.��ҳid", "0 as ��ҳid")
        Case 1
        Case Else
            strTable = strTable & " Union ALL " & Replace(Replace(Replace(Replace(strTable, "סԺ���ü�¼", "������ü�¼"), "A.����", "'' as ����"), "A.�ಡ�˵�", " 0 as �ಡ�˵�"), "A.��ҳid", "0 as ��ҳid")
        End Select
          
        strSQL = _
        " Select A.���,NULL as ѡ��,A.���ݺ�,A.����,A.����ID,A.��ʶ��,A.����,D.���� as ����,A.������," & _
        "       B.���� as ���,Nvl(E.����,C.����) as ��Ŀ," & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "E1.���� as ��Ʒ��,", "") & "C.���,A.����,A.����,A.Ӧ�ս��,A.ʵ�ս��,A.״̬,A.ִ����,A.ִ��ʱ��," & _
        "       A.����Ա,A.�Ǽ�ʱ��,A.��¼����,A.�����־,A.����,A.סԺ��,A.��ҳID,A.��¼״̬ " & _
        " From ( Select Nvl(A.�۸񸸺�,A.���) as ���,Nvl(A.��������,A.���) ����,A.NO as ���ݺ�,A.����,A.����ID,A.��ʶ��,Decode(A.�����־,2,A.����,NULL) as ����," & _
        "               A.��������ID,A.������,A.�շ����,A.�շ�ϸĿID,Avg(A.����)||A.���㵥λ as ����,To_Char(Sum(A.��׼����),'9999990.000') as ����," & _
        "               To_Char(Sum(A.Ӧ�ս��),'9999999" & gstrDec & "') as Ӧ�ս��,To_Char(Sum(A.ʵ�ս��),'9999999" & gstrDec & "') as ʵ�ս��," & _
        "               Decode(A.ִ��״̬,1,'��ִ��','δִ��') as ״̬,A.ִ����,To_Char(A.ִ��ʱ��,'YYYY-MM-DD HH24:MI:SS') as ִ��ʱ��," & _
        "               Nvl(A.����Ա����,A.������) as ����Ա,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as �Ǽ�ʱ��,A.��¼����,A.�����־,To_Number(Decode(Nvl(A.�ಡ�˵�,0),1,NULL,A.��ʶ��)) as סԺ��,A.��ҳID,A.��¼״̬" & _
        "        From (" & strTable & ") A " & _
        "        Group by Nvl(A.�۸񸸺�,A.���),A.NO,A.��ʶ��,A.����ID,A.����,Decode(A.�����־,2,A.����,NULL),A.��������ID,A.������," & _
        "                 A.�շ����,A.�շ�ϸĿID,A.���㵥λ,A.ִ��״̬,A.ִ����,A.ִ��ʱ��,Nvl(A.����Ա����,A.������),A.�Ǽ�ʱ��,A.��¼����,A.�����־,To_Number(Decode(Nvl(A.�ಡ�˵�,0),1,NULL,A.��ʶ��)),A.��ҳID,Nvl(A.��������,A.���),A.��¼״̬" & _
        "       ) A,�շ���Ŀ��� B,�շ���ĿĿ¼ C,���ű� D,�շ���Ŀ���� E" & _
                IIf(gTy_System_Para.bytҩƷ������ʾ = 2, ",�շ���Ŀ���� E1", "") & _
        " Where A.�շ���� = B.���� And A.�շ�ϸĿID = C.ID And A.��������ID=D.ID" & _
        "       And A.�շ�ϸĿID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
                IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3", "") & _
        " Union All" & _
        " Select -1 as ���,NULL as ѡ��,A.NO as ���ݺ�,Decode(A.�����־,1,'(����)',4,'(���)','(סԺ)')||A.NO||" & _
        " '  ������'||A.����||'  ��ʶ�ţ�'||A.��ʶ��||Decode(A.�����־,2,'  ���ţ�'||A.����,NULL)||'  ��'||LTrim(To_Char(Sum(A.ʵ�ս��),'9999999" & gstrDec & "')) as ����," & _
        " -NULL as ����ID,-NULL as ��ʶ��,NULL as ����,NULL as ����,NULL as ������,NULL as ���,NULL as ��Ŀ," & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "NULL as ��Ʒ��,", "") & _
        " NULL as ���,NULL as ����,NULL as ����,NULL as Ӧ�ս��,NULL as ʵ�ս��,NULL as ״̬,NULL,NULL," & _
        " NULL,NULL,A.��¼����,A.�����־,-Null as ����,-Null as סԺ��,-Null as ��ҳID,-Null as ��¼״̬ From (" & strTable & ") A" & _
        " Group by A.NO,A.��¼����,A.����,A.��ʶ��,A.����,A.�����־" & _
        " Order by ���ݺ� Desc,�����־,��¼����,���"
        With SQLCondition
            Set mrsList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .DateB, .DateE, .NOB, .NOE, .State, .Operator, .ID, .Patient, cboUnit.ItemData(cboUnit.ListIndex))
        End With
    End If
    
    mshList.Redraw = False
    mshList.ClearStructure
    mshList.Clear
    mshList.Rows = 2
    
    If mrsList.EOF Then
        stbThis.Panels(2).Text = "��ǰ����û�й��˳��κε���"
        Call SetMenu(False)
    Else
        Set mshList.DataSource = mrsList
        Call SetMenu(True)
    End If
    Call SetBillHead
    Call SetHeader
    
    'Call mshList_EnterCell   'SetHeader����ִ��
    mshList.Redraw = True
    LockWindowUpdate 0
    If Not blnSort Then Call zlCommFun.StopFlash
    
    Me.Refresh
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    LockWindowUpdate 0
End Sub

Private Sub SetBillHead(Optional blnMerge As Boolean = True)
    Dim i As Long, j As Long
    Dim lngRow As Long, lngRows As Long

    lngRow = mshList.Row
    
    Screen.MousePointer = 11
    Me.Refresh
    mshList.Redraw = False
    
    For i = 1 To mshList.Rows - 1
        If Val(mshList.TextMatrix(i, 0)) = -1 Then
            lngRows = lngRows + 1
            
            If blnMerge Then
                mshList.Row = i
                For j = 1 To mshList.Cols - 1
                    mshList.Col = j
                    mshList.CellBackColor = &HEBFFFF
                    mshList.CellAlignment = 1
                    If j <> 3 Then
                        mshList.TextMatrix(i, j) = mshList.TextMatrix(i, 3)
                    End If
                Next
                mshList.MergeRow(i) = True
            End If
            
            mshList.RowHeight(i) = IIf(mnuViewShowHead.Checked, 250, 0)
        ElseIf blnMerge Then
            mshList.MergeRow(i) = False
        End If
    Next
    
    If mshList.RowHeight(lngRow) = 0 Then
        For i = lngRow To mshList.Rows - 1
            If mshList.RowHeight(i) > 0 Then
                Call mshList_LeaveCell
                lngRow = i: Exit For
            End If
        Next
    End If
    mshList.Row = lngRow
    
    'Call mshList_EnterCell   '�����ܻ���setheader��ִ��
    
    mshList.Redraw = True
    Screen.MousePointer = 0
    
    stbThis.Panels(2) = "�� " & lngRows & " �ŵ���"
    stbThis.Tag = stbThis.Panels(2)
End Sub

Private Function GetColValue(ByVal intRow As Integer, strItem As String) As String
'���ܣ���ȡָ���е�ֵ,��ΪĳЩ���Ǻϲ���ʾ,����Ҫ��������
    Dim i As Long, strTmp As String
    If Val(mshList.TextMatrix(intRow, 0)) = -1 Then
        GetColValue = mshList.TextMatrix(IIf(mshList.Row < intRow, intRow + 1, intRow - 1), GetColNum(strItem))
    Else
        GetColValue = mshList.TextMatrix(intRow, GetColNum(strItem))
    End If
End Function

Private Function InitUnits() As Boolean
'���ܣ���ʼ��סԺ�ٴ�����
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo errH
    
    '��������/סԺҽ������
    If InStr(mstrPrivs, ";���п���;") > 0 Then
        gstrSQL = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B " & _
            " Where B.����ID = A.ID " & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
            " And B.������� IN(1,2,3) And B.�������� IN('���','����','����','����','Ӫ��')" & _
            " Order by A.����"
    Else
        gstrSQL = _
            " Select Distinct A.ID,A.����,A.����,Nvl(C.ȱʡ,0) as ȱʡ" & _
            " From ���ű� A,��������˵�� B,������Ա C " & _
            " Where B.����ID = A.ID And A.ID=C.����ID And C.��ԱID=[1]" & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
            " And B.������� IN(1,2,3) And B.�������� IN('���','����','����','����','Ӫ��')" & _
            " Order by A.����"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UserInfo.ID)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboUnit.AddItem rsTmp!���� & "-" & rsTmp!����
            cboUnit.ItemData(cboUnit.NewIndex) = rsTmp!ID
            If cboUnit.ListIndex = -1 Then
                If InStr(mstrPrivs, ";���п���;") > 0 Then
                    If UserInfo.����ID = rsTmp!ID Then cboUnit.ListIndex = cboUnit.NewIndex
                Else
                    If rsTmp!ȱʡ = 1 Then cboUnit.ListIndex = cboUnit.NewIndex
                End If
            End If
            rsTmp.MoveNext
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

Private Function GetFirstRow(ByVal lngRow As Long, lngCol As Long, strValue As String) As Long
'���ܣ��ڵ�ǰ�����У��Ե�ǰ��Ϊ׼����ȡָ������ֵΪstrValue���к�
    Dim lngRowB As Long, lngRowE As Long
    Dim i As Long
    
    If mshList.TextMatrix(lngRow, lngCol) = strValue Then GetFirstRow = lngRow
    
    If Val(mshList.TextMatrix(lngRow, 0)) = -1 Then
        lngRowB = lngRow + 1
    Else
        lngRowB = 2
        For i = lngRow To 1 Step -1
            If Val(mshList.TextMatrix(i, 0)) = -1 Then
                lngRowB = i + 1: Exit For
            End If
        Next
    End If
    
    lngRowE = mshList.Rows - 1
    For i = IIf(Val(mshList.TextMatrix(lngRow, 0)) = -1, lngRow + 1, lngRow) To mshList.Rows - 1
        If Val(mshList.TextMatrix(i, 0)) = -1 Then
            lngRowE = i - 1: Exit For
        End If
    Next
    
    For i = lngRowB To lngRowE
        If mshList.TextMatrix(i, lngCol) = strValue Then
            GetFirstRow = i: Exit For
        End If
    Next
End Function

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub


