VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmҽ���ʻ� 
   Caption         =   "ҽ���ʻ�����"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9975
   Icon            =   "frmҽ���ʻ�.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picSplitV 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2850
      Left            =   7050
      MousePointer    =   9  'Size W E
      ScaleHeight     =   2850
      ScaleWidth      =   45
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1890
      Width           =   45
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   8685
      Top             =   5820
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
            Picture         =   "frmҽ���ʻ�.frx":06EA
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ʻ�.frx":0904
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ʻ�.frx":0B1E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ʻ�.frx":0D38
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ʻ�.frx":0F52
            Key             =   "View"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ʻ�.frx":164C
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ʻ�.frx":1866
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ʻ�.frx":1A80
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   9285
      Top             =   5820
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
            Picture         =   "frmҽ���ʻ�.frx":1C9A
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ʻ�.frx":1EB4
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ʻ�.frx":20CE
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ʻ�.frx":22E8
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ʻ�.frx":2502
            Key             =   "View"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ʻ�.frx":2BFC
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ʻ�.frx":2E16
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ʻ�.frx":3030
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   6390
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmҽ���ʻ�.frx":324A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12515
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
      TabIndex        =   0
      Top             =   0
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   1376
      BandCount       =   2
      ForeColor       =   -2147483635
      FixedOrder      =   -1  'True
      _CBWidth        =   9975
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      BandForeColor1  =   -2147483635
      Child1          =   "tbrThis"
      MinWidth1       =   4995
      MinHeight1      =   720
      Width1          =   810
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Caption2        =   "�������"
      Child2          =   "cmb����"
      MinWidth2       =   1995
      MinHeight2      =   300
      Width2          =   1935
      NewRow2         =   0   'False
      AllowVertical2  =   0   'False
      Begin VB.ComboBox cmb���� 
         Height          =   300
         Left            =   7890
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   1995
      End
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   720
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   6855
         _ExtentX        =   12091
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
            NumButtons      =   11
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
               Caption         =   "��֤"
               Key             =   "Modify"
               Object.ToolTipText     =   "�����֤"
               Object.Tag             =   "��֤"
               ImageKey        =   "Modify"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "SplitModify"
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "���"
               Key             =   "Custom"
               Description     =   "Custom"
               Object.ToolTipText     =   "�Զ��岡�����"
               Object.Tag             =   "���"
               ImageKey        =   "View"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Find"
               Description     =   "����"
               Object.ToolTipText     =   "����ҽ���ʻ�"
               Object.Tag             =   "����"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "����"
               Object.ToolTipText     =   "��ǰ��������"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh�ʻ�_S 
      Height          =   5655
      Left            =   15
      TabIndex        =   3
      Top             =   735
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   9975
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
      Appearance      =   0
      MouseIcon       =   "frmҽ���ʻ�.frx":3ADC
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox picOther 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   5625
      Left            =   7200
      ScaleHeight     =   5595
      ScaleWidth      =   2715
      TabIndex        =   5
      Top             =   750
      Width           =   2745
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh�����Ϣ 
         Height          =   1365
         Left            =   -30
         TabIndex        =   9
         Top             =   4260
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   2408
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
         Appearance      =   0
         MouseIcon       =   "frmҽ���ʻ�.frx":3DF6
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh��� 
         Height          =   3405
         Left            =   -30
         TabIndex        =   10
         Top             =   450
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   6006
         _Version        =   393216
         FixedCols       =   0
         RowHeightMin    =   250
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483630
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         MouseIcon       =   "frmҽ���ʻ�.frx":4110
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Lbl���������Ϣ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H009B6737&
         Caption         =   "���������Ϣ��"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   3960
         Width           =   1260
      End
      Begin VB.Label Lbl������ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H009B6737&
         Caption         =   "��������"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   540
         TabIndex        =   7
         Top             =   135
         Width           =   900
      End
      Begin VB.Label lbl��� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2002"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   135
         TabIndex        =   6
         Top             =   120
         Width           =   390
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFilePreview 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFileSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileCard 
         Caption         =   "��Ƭ��ӡ(&A)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSplit2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileQuit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditModify 
         Caption         =   "�����֤(&I)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "ɾ���ʻ�(&D)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditSplit0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditPassword 
         Caption         =   "�޸�����(&M)"
      End
      Begin VB.Menu mnuEditSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDisease 
         Caption         =   "����ѡ��(&D)"
      End
      Begin VB.Menu mnuEditBalance 
         Caption         =   "���ý��㷽ʽ(&J)"
      End
      Begin VB.Menu mnuEditReckoning 
         Caption         =   "�������㷽ʽ(&C)"
      End
      Begin VB.Menu mnuEditICD 
         Caption         =   "�ϴ�ICD-10��������(&I)"
      End
      Begin VB.Menu mnuEditQuery 
         Caption         =   "��ѯ��λǷ��(&Q)"
      End
      Begin VB.Menu mnuEditBed 
         Caption         =   "ת�����ͥ����(&B)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSub 
         Caption         =   "����Ǽ�(&S)"
      End
      Begin VB.Menu mnuEditRollIncome 
         Caption         =   "������Ժ�Ǽ�(&R)"
      End
      Begin VB.Menu mnuEditRollAdmit 
         Caption         =   "��������Ǽ�(&R)"
      End
      Begin VB.Menu mnuEditOut 
         Caption         =   "�����Ժ�Ǽ�(&O)"
      End
      Begin VB.Menu mnuEditOutDel 
         Caption         =   "������Ժ�Ǽ�(&C)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditSplit3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditVerify_UpDetail 
         Caption         =   "�����ϴ�������ϸ(&E)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditVerify_Account 
         Caption         =   "�˶��ʻ�֧����Ϣ(&A)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditVerify_Hospital 
         Caption         =   "�˶����Ժ��Ϣ(&H)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditVerify_ZYPrice 
         Caption         =   "�˶�סԺ������Ϣ(&Y)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditVerify_Price 
         Caption         =   "�˶Է��ý�����Ϣ(&T)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditVerify_Detail 
         Caption         =   "�˶Է�����ϸ��Ϣ(&D)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditSp 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditSingleDisease 
         Caption         =   "��������ٱ༭(&B)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditXE 
         Caption         =   "�޶�༭(&B)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditSplit4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditMend 
         Caption         =   "����(&E)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditLoss 
         Caption         =   "��ֹ����(&L)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditSplit5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditVerify 
         Caption         =   "������Ŀ����(&V)"
         Visible         =   0   'False
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
         Begin VB.Menu mnuViewTool_1 
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
      Begin VB.Menu mnuViewCustom 
         Caption         =   "�Զ��������Ϣ(&A)"
      End
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "����(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuView_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
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
Attribute VB_Name = "frmҽ���ʻ�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'���볣�����ܶ���ɹ����ģ�������ʹ�õ��ĵط��������壬�ڱ���ʱͳһ�޸�
#Const gverControl = 99  ' 0-��֧�ֶ�̬ҽ��(9.19��ǰ),1-֧�ֶ�̬ҽ���޸��Ӳ���(9.22��ǰ) , _
    2-����������������ʽ��������һ��;����������ԭʼ��������һ��;�����շ�����������;99-���н������Ӹ��Ӳ���(���°�)

Private Enum ��Enum
    rowסԺ���� = 1
    row�ʻ���� = 2
    row�ʻ����� = 3
    row�ʻ�֧�� = 4
    row�������� = 5
    row�����ۼ� = 6
    rowͳ���޶� = 7
    row����ͳ�� = 8
    rowͳ�ﱨ�� = 9
    row����޶� = 10
    row����ۼ� = 11
    row������Ϣ = 12
End Enum

Private Enum ��Enum
    col���� = 0
    col���� = 1
    colҽ���� = 2
    col����ID = 3
    col���� = 4
    col�Ա� = 5
    col�������� = 6
    col���֤�� = 7
    col��Ա��� = 8
    col��ݱ��� = 9
    col��λ���� = 10
    col����֤�� = 11
    col���� = 12
    col״̬ = 13
    col�ʻ���� = 14
    col����ʱ�� = 15
End Enum

Private mblnLoad As Boolean                     '��һ������
Private mstr�����ֶ� As String                  '�û����õ��ֶ�
Private mstrFind As String                      '��������

Private mrs�ʻ� As New ADODB.Recordset
Private mint���� As Integer
Private mcol���� As New Collection              '����ҽ����������������
Private mcol���� As New Collection              '�����ҽ���Ƿ���Գ�ʼ��
Private msngStartX As Single
Private strServer As String, strUser As String, strPass As String
Private mcnYB As New ADODB.Connection   'ҽ��ǰ�÷���������
Private mrs���� As New ADODB.Recordset

Private Sub cmb����_Click()
    Dim blnCanUse As Boolean
    
    With cmb����
        If mint���� = .ItemData(.ListIndex) Then Exit Sub
        mint���� = .ItemData(.ListIndex)
    End With
    
    Call SetMenuState
    
    mnuEditPassword.Enabled = True
    mnuEditModify.Enabled = True
    
    mnuEditXE.Visible = False ' mint���� = TYPE_������ Or mint���� = TYPE_����������
    mnuEditSp.Visible = mnuEditXE.Visible
    tbrThis.Buttons("Modify").Enabled = mnuEditModify.Enabled
    
    blnCanUse = GetInsureInit(mint����)
    
    '������(2005-08-19):�ɶ���ҽ���޴˹��ܣ����Ρ�
    If mint���� = 20 Then
       mnuEditRollIncome.Visible = False
       mnuEditOut.Visible = False
    Else
       mnuEditRollIncome.Visible = True
       mnuEditOut.Visible = True
    End If
    
    mnuEditSub.Enabled = blnCanUse
    mnuEditDisease.Enabled = blnCanUse
    mnuEditRollIncome.Enabled = blnCanUse
    mnuEditRollAdmit.Enabled = blnCanUse
    mnuEditQuery.Enabled = blnCanUse
    mnuEditOutDel.Visible = (mint���� = TYPE_��ͨ)
    
    Call InitTable
    Call FillList
End Sub

Private Sub cbr_HeightChanged(ByVal NewHeight As Single)
    Call ResizeForm(NewHeight)
End Sub

Private Sub Form_Activate()
    mint���� = cmb����.ItemData(cmb����.ListIndex)
    If mblnLoad = True Then
        lbl���.Caption = Format(zlDatabase.Currentdate, "yyyy")
        mstrFind = " and A.����ʱ��>=sysdate-3"
        If mint���� = TYPE_�Ĵ�üɽ Then mstrFind = " And Nvl(A.�Ҷȼ�,0)<>9"
        
        '��ʾ�ʻ�
        Call FillList
        Call GetAccountInfo
    End If
    mblnLoad = False
End Sub

Private Sub Form_Load()
    On Error Resume Next
    mblnLoad = True
    'ȡע���
    mstr�����ֶ� = Replace(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "�����ֶ�", ""), "'", "")
    
    zlControl.CboSetHeight cmb����, 3600
    Call InitTable
    RestoreWinState Me, App.ProductName
    Call Ȩ�޿���
End Sub

Private Sub Form_Resize()
    Call ResizeForm(cbr.Height)
    Call GetAccountInfo
End Sub

Private Sub ResizeForm(ByVal cbrHeight As Single)
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    
    picSplitV.Left = Me.ScaleWidth - 3000
    With msh�ʻ�_S
        .Top = IIf(cbr.Visible, cbrHeight, 0)
        .Width = picSplitV.Left - 25
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
    With picSplitV
        .Top = msh�ʻ�_S.Top
        .Height = msh�ʻ�_S.Height
    End With
    With picOther
        .Top = msh�ʻ�_S.Top
        .Left = picSplitV.Left + picSplitV.Width
        .Height = msh�ʻ�_S.Height
        .Width = Me.ScaleWidth - .Left
    End With
End Sub

Private Sub mnuEditBalance_Click()
    '���ý��㷽ʽ
    Dim lng����ID  As Long, rsTemp As New ADODB.Recordset
    If mint���� <> TYPE_������ Then Exit Sub
    
    With msh�ʻ�_S
        '����ֱ�Ӵ��б���ȡ��
        lng����ID = Val(.TextMatrix(.Row, col����ID))
        
        If lng����ID <= 0 Then
            MsgBox "��ѡ��һλҽ�����ˡ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '�жϸò����Ƿ���סԺ��¼
        gstrSQL = "select A.��λ���� " & _
                  "  from �����ʻ� A " & _
                  "  Where A.����ID =[1] And A.���� = [2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID, mint����)
        If rsTemp.EOF = True Then
            '�޷��Ӽ�¼����ȡ�ò�������
            MsgBox "���� " & msh�ʻ�_S.TextMatrix(.Row, col����) & " �޷��ҵ���Ч�ĵǼ���Ϣ��", vbInformation, gstrSysName
            Exit Sub
        End If
        
        Call ���ý��㷽ʽ_����(lng����ID, Me, True)
    End With
End Sub

Private Sub mnuEditBed_Click()
    Dim lng����ID As Long, lng��ҳID As Long
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    lng����ID = Val(msh�ʻ�_S.TextMatrix(msh�ʻ�_S.Row, col����ID))
    If lng����ID <= 0 Then
        MsgBox "��ѡ��һλҽ�����ˡ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    'ȡ��ҳID
    gstrSQL = " Select סԺ���� AS ��ҳID From ������Ϣ Where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��ҳID", lng����ID)
    lng��ҳID = rsTemp!��ҳID
    
    Call ת�����ͥ����(lng����ID, lng��ҳID)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub mnuEditDelete_Click()
    Dim lng����ID As Long
    Dim blnDelete As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    'ɾ���ʻ���Ϣ(�Ҷȼ�:0-����;1-��ʧ;2-��ֹ����;3-��ֹͳ��;9-�ʻ��ѳ���)
    lng����ID = Val(msh�ʻ�_S.TextMatrix(msh�ʻ�_S.Row, col����ID))
    If lng����ID <= 0 Then
        MsgBox "��ѡ��һλҽ�����ˡ�", vbInformation, gstrSysName
        Exit Sub
    End If
    If MsgBox("�����Ҫɾ���ñ����ʻ���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    On Error Resume Next
    Err = 0
    gcnOracle.BeginTrans
    
'    gstrSQL = "Select count(*) Records From ���ս����¼ Where ����ID=" & lng����ID
'    Call OpenRecordset(rsTemp, Me.Caption)
'    blnDelete = (rsTemp.RecordCount = 0)
'
'    If blnDelete Then
'        gstrSQL = "Delete �����ʻ� Where ����=" & mint���� & " And ����ID=" & lng����ID
'        gcnOracle.Execute gstrSQL
'    Else
        gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_�Ĵ�üɽ & ",'�Ҷȼ�','9')"
        gcnOracle.Execute gstrSQL
'    End If
    gcnOracle.CommitTrans
    
    Call FillList
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
    gcnOracle.RollbackTrans
End Sub

Private Sub mnuEditDisease_Click()
    Dim lng����ID As Long, lng��ҳID As Long, rs���� As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim lng���� As Long, str���� As String
    Dim rsSelected As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    With msh�ʻ�_S
        '����ֱ�Ӵ��б���ȡ��
        lng����ID = Val(.TextMatrix(.Row, col����ID))
        If lng����ID <= 0 Then
            MsgBox "��ѡ��һλҽ�����ˡ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '�жϸò����Ƿ���סԺ��¼
        gstrSQL = "select A.����,B.��ҳID,B.��Ժ����,B.��Ժ���� " & _
                  "  from ������Ϣ A,������ҳ B " & _
                  "  Where A.����ID = [1] And A.����ID = B.����ID And B.���� = [2]" & _
                  "  Order by B.��Ժ���� Desc"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID, mint����)
        If rsTemp.EOF = True Then
            '�޷��Ӽ�¼����ȡ�ò�������
            MsgBox "���� " & msh�ʻ�_S.TextMatrix(.Row, col����) & " �޷��ҵ���Ч��סԺ��¼������δסԺ��δ��ҽ�������Ժ��", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If IsNull(rsTemp("��Ժ����")) = False Then
            If MsgBox("���� " & rsTemp("����") & " ����" & Format(rsTemp("��Ժ����"), "yyyy-MM-dd") & "��Ժ���Ƿ���Ҫ���¼�����Ϣ��", vbQuestion Or vbYesNo Or vbDefaultButton2) = vbNo Then
                Exit Sub
            End If
        End If
        
        lng��ҳID = rsTemp!��ҳID
        If mint���� = TYPE_������ Then
            Call ���³�Ժ����_����(lng����ID, lng��ҳID)
        ElseIf mint���� = TYPE_������ Then
            Call ����ѡ��_����(lng����ID)
        ElseIf mint���� = TYPE_���������� Then
            Call ���¼���_����������(Me, lng����ID, lng��ҳID)
        ElseIf mint���� = TYPE_ɽ�� Then
            Call ���²���_ɽ��(lng����ID, lng��ҳID)
        'Beging 2005-11-16 ���ջ�
        ElseIf mint���� = type_ͭ����ҽ Then
             If ����ѡ��_ͭ����ҽ(lng����ID, mint����) Then
                MsgBox "���²��ֳɹ���", vbInformation, gstrSysName
            End If
        'End 2005-11-16 ���ջ�
        'Beging 20051024 �¶�
        ElseIf mint���� = TYPE_�ɶ��ڽ� Then
            Call ����֢ѡ��_�ɶ��ڽ�(lng����ID, lng��ҳID)
        'End 20051024 �¶�
        ElseIf mint���� = TYPE_���Ͻ�ˮ Then
            'סԺҪѡ���֣���ȷ��һЩ�����շ���Ŀ
            gstrSQL = " Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ֲ�','��ͨ��') as ��� " & _
                    " From ���ղ��� A where A.����=[1]"
            Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, "�����֤", mint����)
            If frmListSel.ShowSelect(mint����, rs����, "ID", "����ѡ��", "��ѡ����") Then
                gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & mint���� & ",'����ID','" & rs����!ID & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "���²���")
                MsgBox "�Ѹ���Ϊ��ѡ��Ĳ��֣�", vbInformation, gstrSysName
            Else
                gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & mint���� & ",'����ID','" & 0 & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "���²���")
                MsgBox "�����������Ϣ��", vbInformation, gstrSysName
            End If
        ElseIf mint���� = TYPE_�Թ��� Then
            '��ȡ��ѡ��Ĳ���
            gstrSQL = " Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ֲ�','��ͨ��') as ��� " & _
                    " From ���ղ��� A,zlyb.������Ϣ B where A.����=[1] And B.����ID=[2] And A.ID=B.����ID And A.����=B.����"
            Set rsSelected = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ϴ���ѡ��Ĳ���", mint����, lng����ID)
            
            'סԺҪѡ���֣���ȷ��һЩ�����շ���Ŀ
            gstrSQL = " Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ֲ�','��ͨ��') as ��� " & _
                    " From ���ղ��� A where A.����=[1]"
            Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, "�����֤", mint����)
            If rs����.RecordCount > 0 Then
                If frm�ಡ��ѡ��.ShowSelect(rs����, "ID", "ҽ������ѡ��", "��ѡ��ҽ�����֣�", rsSelected, False) = True Then
                    lng���� = 0
                    str���� = ""
                    With rs����
                        If .RecordCount <> 0 Then
                            .MoveFirst
                            lng���� = rs����("ID")
                        End If
                        Do While Not .EOF
                            str���� = str���� & "|" & rs����!ID
                            .MoveNext
                        Loop
                        If str���� <> "" Then str���� = Mid(str����, 2)
                        
                        gstrSQL = "zlyb.zl_������Ϣ_INSERT(" & mint���� & "," & lng����ID & ",'" & str���� & "')"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, "���²���")
                    End With
                    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & mint���� & ",'����ID','" & lng���� & "')"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "���²���")
                End If
            End If
        ElseIf mint���� = TYPE_������ Then
            '��ȡ��ѡ��Ĳ���
            Dim int�ϴ� As Integer      '�жϲ�����ҳ�Ƿ��ϴ�������ϴ����޸ĵ�סԺ��������Ϊ��Ժ����
            
            '��ȡ������ҳ
            gstrSQL = "Select Nvl(�Ƿ��ϴ�,0) AS �Ƿ��ϴ� From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������ҳ", lng����ID, lng��ҳID)
            int�ϴ� = rsTemp!�Ƿ��ϴ�
            
            '��ȡ����ѡ��Ĳ���
            gstrSQL = " Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ֲ�','��ͨ��') as ��� " & _
                    " From ���ղ��� A,zlyb.������Ϣ B where A.����=[1] And B.����ID=[2] And B.��ҳID=[3] ANd B.״̬=[4] And A.ID=B.����ID And A.����=B.����"
            Set rsSelected = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ϴ���ѡ��Ĳ���", mint����, lng����ID, lng��ҳID, int�ϴ�)
            
            'סԺҪѡ���֣���ȷ��һЩ�����շ���Ŀ
            gstrSQL = " Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ֲ�','��ͨ��') as ��� " & _
                    " From ���ղ��� A where A.����=[1]"
            Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, "�����֤", mint����)
            If rs����.RecordCount > 0 Then
                If frm�ಡ��ѡ��.ShowSelect(rs����, "ID", "ҽ������ѡ��", "��ѡ��ҽ�����֣�", rsSelected, False) = True Then
                    lng���� = 0
                    str���� = ""
                    With rs����
                        If .RecordCount <> 0 Then
                            .MoveFirst
                            lng���� = rs����("ID")
                        End If
                        Do While Not .EOF
                            str���� = str���� & "|" & rs����!ID
                            .MoveNext
                        Loop
                        If str���� <> "" Then str���� = Mid(str����, 2)
                        
                        If int�ϴ� = 1 Then
                            '�����Ժ�����¼���������ϴ����������޸�
                            gstrSQL = "Select 1 from zlyb.���ս����¼ " & _
                                " Where ����ID=[1] And ��ҳID=[2]" & _
                                " And Nvl(�Ƿ��ϴ�,0)=1 And nvl(��;����,0)=0"
                            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ���ڳ�Ժ�����¼", lng����ID, lng��ҳID)
                            If rsTemp.RecordCount <> 0 Then
                                MsgBox "�ò��˵ĳ�Ժ�����¼���ϴ����������޸Ĳ��֣�", vbInformation, gstrSysName
                                Exit Sub
                            End If
                        End If
                        
                        gstrSQL = "zlyb.zl_������Ϣ_INSERT(" & mint���� & "," & lng����ID & "," & lng��ҳID & "," & int�ϴ� & ",'" & str���� & "')"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, "���²���")
                        If int�ϴ� = 0 Then
                            'ͬ�����³�Ժ����
                            gstrSQL = "zlyb.zl_������Ϣ_INSERT(" & mint���� & "," & lng����ID & "," & lng��ҳID & ",1,'" & str���� & "')"
                            Call zlDatabase.ExecuteProcedure(gstrSQL, "���²���")
                        End If
                    End With
                    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & mint���� & ",'����ID','" & lng���� & "')"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "���²���")
                End If
            End If
        ElseIf mint���� = TYPE_������ Then
            '�����޸Ĳ��֣�������Ժ�Ǽ�ʱ����ѡ�������������
            If Not ҽ����ʼ��_������ Then Exit Sub
            Call ���²���_������(lng����ID, lng��ҳID)
        ElseIf mint���� = TYPE_�Ͻ� Then
            If Not ҽ����ʼ��_�Ͻ� Then Exit Sub
            Call ���²���_�Ͻ�(lng����ID, lng��ҳID)
        End If
        Call FillList
    End With

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub mnuEditICD_Click()
    Dim lng����ID  As Long, rsTemp As New ADODB.Recordset
    If mint���� <> TYPE_������ Then Exit Sub
    
    With msh�ʻ�_S
        '����ֱ�Ӵ��б���ȡ��
        lng����ID = Val(.TextMatrix(.Row, col����ID))
        
        If lng����ID <= 0 Then
            MsgBox "��ѡ��һλҽ�����ˡ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If ����ICD����_����(lng����ID) Then
            MsgBox "ICD�����ϴ��ɹ���", vbInformation, gstrSysName
        End If
    End With
End Sub

Private Sub mnuEditLoss_Click()
    Dim int״̬ As Integer
    Dim lng����ID As Long
    Dim strMsg As String
    Dim rsAccount As New ADODB.Recordset
    
    On Error GoTo errHand
    '�������ʻ�(�Ҷȼ�:0-����;1-��ֹ�����ʻ�;9-�ʻ��ѳ���)
    lng����ID = Val(msh�ʻ�_S.TextMatrix(msh�ʻ�_S.Row, col����ID))
    If lng����ID <= 0 Then
        MsgBox "��ѡ��һλҽ�����ˡ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '��鿨��״̬������Ѿ����������ʾҪ���������򽫷����ÿ�
    gstrSQL = "Select Nvl(�Ҷȼ�,0) ״̬ From �����ʻ� Where ����ID=[1] And ����=[2]"
    Set rsAccount = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID, TYPE_�Ĵ�üɽ)
    If rsAccount!״̬ = 0 Then
        int״̬ = 1
        strMsg = "�����ÿ��𣿣������󽫲���ʹ�ã�"
    Else
        int״̬ = 0
        strMsg = "�ָ��ÿ���״̬Ϊ������"
    End If
    If MsgBox("��ȷ��Ҫ" & strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    '���������˵����
    strMsg = ""
    If int״̬ = 1 Then
        Do While True
            strMsg = InputBox("���������������Ϣ��", "����ҽ������ʹ��")
            If Trim(strMsg) <> "" Then Exit Do
        Loop
    End If
    
    '�����ʻ���Ϣ
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_�Ĵ�üɽ & ",'�Ҷȼ�','" & int״̬ & " ')"
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    If int״̬ = 1 Then
        gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_�Ĵ�üɽ & ",'��ע','''" & strMsg & " ''')"
        gcnOracle.Execute gstrSQL, , adCmdStoredProc
    Else
        gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_�Ĵ�üɽ & ",'��ע','NULL')"
        gcnOracle.Execute gstrSQL, , adCmdStoredProc
    End If
    
    Call FillList
    Call msh�ʻ�_S_EnterCell
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub mnuEditMend_Click()
    '����
    Dim strIdentify As String
    Dim bytType As Byte
    Dim lng����ID As Long
    
    On Error GoTo errHand
    lng����ID = Val(msh�ʻ�_S.TextMatrix(msh�ʻ�_S.Row, col����ID))
    If lng����ID = 0 Then
        MsgBox "��ѡ��һλҽ�����ˣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If mint���� = TYPE_�����山 Then
        '����ȡ����ǰ���˵�������¼
        Dim rsTemp As New ADODB.Recordset
        '�����ʻ�Ŀǰ���ֵ
        '--����id, ����, ����, ���ţ�ҽ������), ҽ����(���˱��), ����(֧����� ), ��Ա���(�α���Ա���ڵ��籣�����������), ��λ����(��λ����(��λ����)), ˳���(��),
        '--����֤��(ҽ����Ա���|ҽ���չ����|ҽ�Ʋ������|�ۼƽɷ�����), �ʻ����(�ʻ����), ��ǰ״̬, ����id������ID), ��ְ(1), �����(����), �Ҷȼ�, ����ʱ��
        Dim strTemp As String
        Dim strArr
        
        '�����ϵͳ����Ա,���ܸ���.
        gstrSQL = "select * from zlsystems  where upper(������)='" & UCase(gstrDbUser) & "'"
        
        zlDatabase.OpenRecordset rsTemp, gstrSQL, "ȷ���Ƿ�Ϊϵͳ����Ա�û�"
        If rsTemp.RecordCount = 0 Then
            ShowMsgbox "��ǰ�û�����ϵͳ����Ա,����ȡ��������Ϣ!"
            Exit Sub
        End If
        gstrSQL = "select a.����,a.ҽ����,a.����,a.��Ա���,a.��λ����,a.˳���,a.����֤��,a.�ʻ����,a.��ǰ״̬,a.����id,a.��ְ,a.�����,a.�Ҷȼ�,a.����ʱ��," & _
                 "        b.����,decode( b.�Ա�,'��','1','Ů','2','3') as �Ա�, b.����, b.��������, b.���֤��,A.������,A.������,A.֧����� " & _
                 " from �����ʻ� a,������Ϣ b " & _
                 " WHERE a.����id=" & lng����ID & " AND a.����id=b.����id and a.����=" & TYPE_�����山
 
        zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ������Ϣ"
    
            With g�������_�����山
                .���� = Nvl(rsTemp!����)
                .���˱�� = Nvl(rsTemp!ҽ����)
                .���� = Nvl(rsTemp!����)
                .�Ա� = Nvl(rsTemp!�Ա�)
                .���� = Nvl(rsTemp!�����, 0)
                .�������� = Format(rsTemp!��������, "yyyy-mm-dd")
                strTemp = Nvl(rsTemp!��λ����)
                If InStr(1, strTemp, "(") <> 0 Then
                    .��λ���� = Mid(strTemp, InStr(1, strTemp, "(") + 1)
                    .��λ���� = Val(Mid(.��λ����, 1, Len(.��λ����) - 1))
                    .��λ���� = Mid(strTemp, 1, InStr(1, strTemp, "(" - 1))
                Else
                    .��λ���� = strTemp
                    .��λ���� = 0
                End If
                .���� = Nvl(rsTemp!����)
                .֧����� = Nvl(rsTemp!֧�����)
                .�籣���칹������ = Nvl(rsTemp!��Ա���)
                
                strTemp = Nvl(rsTemp!����֤��, "|||")
                strTemp = IIf(strTemp = "", "|||", strTemp)
                strArr = Split(strTemp, "|")
                
                .ҽ����Ա��� = strArr(0)
                .ҽ���չ���� = strArr(1)
                .ҽ�Ʋ������ = strArr(2)
                .�ۼƽɷ����� = Val(strArr(3))
                .�ʻ���� = Nvl(rsTemp!�ʻ����, 0)
                
                .���֤�� = Nvl(rsTemp!���֤��)
                .����ID = Nvl(rsTemp!����ID, 0)
                .������ = Nvl(rsTemp!������)
                .������ = Nvl(rsTemp!������)
                .���ֱ��� = "000000"
                .������� = False
            End With
            If ������¼����_�����山 = True Then
                ShowMsgbox "ȡ���ɹ�"
            End If
            Exit Sub
    End If
    
    bytType = 4
'$IF HIS9.19
#If gverControl = 0 Then
    strIdentify = gclsInsure.Identify(bytType, lng����ID)
'ELSE
#Else
    strIdentify = gclsInsure.Identify(bytType, lng����ID, mint����)
#End If
'$END IF
    If strIdentify <> "" Then
        Call FillList
    End If
    
    Call msh�ʻ�_S_EnterCell
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub mnuEditOut_Click()
    Dim lng����ID As Long, lng��ҳID As Long
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    lng����ID = Val(msh�ʻ�_S.TextMatrix(msh�ʻ�_S.Row, col����ID))
    
    'Modified by ���� 20031218 ����������
    If Not (mint���� = TYPE_�������� Or mint���� = TYPE_����ʡ Or mint���� = TYPE_������ Or _
    mint���� = TYPE_��ƽ�� Or mint���� = TYPE_������ Or mint���� = TYPE_����ʡ Or _
    mint���� = TYPE_������ Or mint���� = TYPE_���������� Or mint���� = TYPE_���� Or mint���� = TYPE_������) Then Exit Sub
    If lng����ID = 0 Then
        MsgBox "��ѡ��һλҽ�����ˣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    If MsgBox("��ȷ��ҪΪ�ò��˲����Ժ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    'ȡ����ҳID
    gstrSQL = "Select Nvl(סԺ����,0) ��ҳID From ������Ϣ Where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��ҳID", lng����ID)
    lng��ҳID = rsTemp!��ҳID
    '����δ�����ʱ������������Ժ����
    If ����δ�����(lng����ID, lng��ҳID) Then
        MsgBox "��ҽ�����˻�����δ����ã�����������Ժ������", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '����ҽ����Ժ�ӿ�
    Select Case mint����
    Case TYPE_��������, TYPE_����ʡ, TYPE_������, TYPE_��ƽ��
        If Not frm�ȴ���Ӧ.ShowME(mint����, ������ʽ.��Ժ, ����Ŀ��.ˢ��) Then Exit Sub
        If lng����ID <> ��ȡ����ID(mint����) Then
            MsgBox "������Ϣ������", vbInformation, gstrSysName
            Exit Sub
        End If
        If Not frm�ȴ���Ӧ.ShowME(mint����, ������ʽ.��Ժ, ����Ŀ��.����, lng����ID) Then Exit Sub
    
        '��Ժ�Ǽ�
        gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & mint���� & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "��������")
        MsgBox "��ҽ�����˳ɹ������Ժ������", vbInformation, gstrSysName
    Case TYPE_������, TYPE_����ʡ
        gstrSQL = "Select A.��Ժ����,A.��Ժ����,Decode(A.��Ժ��ʽ,'����',0,'����',1,'תԺ',2,9) as ��Ժ��ʽ,B.����,D.סԺ��,Sysdate as ����ʱ��," & _
                " C.����,C.ҽ����,C.����,C.˳��� " & _
                " From ������ҳ A,���ű� B,�����ʻ� C,������Ϣ D " & _
                " Where A.����ID=D.����ID And A.����ID=[1] And A.��ҳID=[2]" & _
                " And A.��Ժ����ID=B.ID And A.����ID=C.����ID And C.����=[3]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ˳���", lng����ID, lng��ҳID, mint����)
    
        If rsTemp.EOF Then
            MsgBox "û�д˲��˻�˲��˲���ҽ�����ˣ��޷������Ժ������", vbExclamation, gstrSysName
            Exit Sub
        End If
        If IsNull(rsTemp!˳���) Then
            MsgBox "δ���ָò��˵�סԺ����˳���,����ִ�н��ף�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If Not ��Ժ�Ǽ�_����(lng����ID, lng��ҳID, rsTemp!˳���, mint����, True, True) Then Exit Sub
        MsgBox "��ҽ�����˳ɹ������Ժ������", vbInformation, gstrSysName
    Case TYPE_������
        Call ��Ժ�Ǽ�_������(lng����ID, lng��ҳID)
    Case TYPE_������
        Call ��Ժ�Ǽ�_����(lng����ID, lng��ҳID, True)
    Case TYPE_����������
        Call ��Ժ�Ǽ�_����������(lng����ID, lng��ҳID)
    Case TYPE_����
        Call ��Ժ�Ǽ�_����(lng����ID, lng��ҳID)
    End Select
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub mnuEditOutDel_Click()
    Dim strסԺ�� As String
    Dim lng����ID As Long, lng��ҳID  As Long
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    With msh�ʻ�_S
        '����ֱ�Ӵ��б���ȡ��
        lng����ID = Val(.TextMatrix(.Row, col����ID))
        If lng����ID <= 0 Then
            MsgBox "��ѡ��һλҽ�����ˡ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If MsgBox("���Ƿ�Ҫ�����ˡ�" & .TextMatrix(.Row, col����) & "���ָ���ҽ����Ժ״̬��", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
        
        '�����ҳID
        gstrSQL = "select MAX(��ҳID) AS ��ҳID from ������ҳ where ����ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", lng����ID)
        If IsNull(rsTemp("��ҳID")) = True Then
            MsgBox "ֻ��סԺ���˲ſ��Դ�ҽ������תΪ��ͨ���ˡ�", vbInformation, gstrSysName
            Exit Sub
        End If
        lng��ҳID = rsTemp("��ҳID")
        
        '�жϸò����Ƿ���סԺ��¼
        gstrSQL = "select A.˳��� " & _
                  "  from �����ʻ� A " & _
                  "  Where A.����ID = " & lng����ID & " And A.���� = " & mint����
        Call OpenRecordset(rsTemp, Me.Caption)
        If rsTemp.EOF = True Then
            '�޷��Ӽ�¼����ȡ�ò�������
            MsgBox "���� " & msh�ʻ�_S.TextMatrix(.Row, col����) & " �޷��ҵ���Ч�ĵǼ���Ϣ��", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If mint���� = TYPE_��ͨ Then
            strסԺ�� = Nvl(rsTemp!˳���)
            If frmConn��ͨ.Execute("I345", 0, strסԺ��, "����ȡ����Ժ......") = False Then Exit Sub
            MsgBox "���ˡ�" & .TextMatrix(.Row, col����) & "���ѻָ���ҽ����Ժ״̬��", vbInformation, gstrSysName
        End If
    End With
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub mnuEditPassword_Click()
    Dim str���� As String, strҽ���� As String, str������ As String, str���� As String
    Dim lng����ID As Long, str����ID As String, lng����ID As Long
    
    Select Case mint����
        Case TYPE_�Թ���
            Call frmIdentify����.GetPatient(False, True)
        Case TYPE_������
            Call frmIdentify����.GetPatient(2, True, lng����ID, str����ID)
        Case type_�ɶ�����
            Call frmIdentify�ɶ�����.GetIdentify(type_�ɶ�����, str����, strҽ����, str������, str����, True, True)
        Case TYPE_�¶�
            Call frmIdentify�ɶ�����.GetIdentify(TYPE_�¶�, str����, strҽ����, str������, str����, True, True)
        Case TYPE_��������
            lng����ID = Val(msh�ʻ�_S.TextMatrix(msh�ʻ�_S.Row, col����ID))
            If lng����ID <= 0 Then
                MsgBox "��ѡ��һλҽ�����ˡ�", vbInformation, gstrSysName
                Exit Sub
            End If
           If frmIdentify��ľ����.GetPatient(99, lng����ID) <> "" Then
                mnuViewRefresh_Click
           End If
    End Select
End Sub

Private Sub mnuEditQuery_Click()
    Dim lng����ID  As Long, rsTemp As New ADODB.Recordset
    
    If mint���� <> TYPE_������ Then Exit Sub
    
    With msh�ʻ�_S
        '����ֱ�Ӵ��б���ȡ��
        lng����ID = Val(.TextMatrix(.Row, col����ID))
        If lng����ID <= 0 Then
            MsgBox "��ѡ��һλҽ�����ˡ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '�жϸò����Ƿ���סԺ��¼
        gstrSQL = "select A.��λ����,A.����,A.ҽ����,A.����֤��,A.�������" & _
                  "  from �����ʻ� A " & _
                  "  Where A.����ID = " & lng����ID & " And A.���� = " & mint����
        Call OpenRecordset(rsTemp, Me.Caption)
        If rsTemp.EOF = True Then
            '�޷��Ӽ�¼����ȡ�ò�������
            MsgBox "���� " & msh�ʻ�_S.TextMatrix(.Row, col����) & " �޷��ҵ���Ч�ĵǼ���Ϣ��", vbInformation, gstrSysName
            Exit Sub
        End If
        If mint���� = TYPE_�ɶ��ڽ� Then
            Dim strͳ�� As String
            strͳ�� = Split(Nvl(rsTemp!����֤��) & "|||||", "|")(0)
            Call ��ѯ��λǷ��_�ɶ��ڽ�(Nvl(rsTemp!ҽ����), Nvl(rsTemp!����), strͳ��)
        Else
            Call ��ѯǷ�ѵ�λ_����(Nvl(rsTemp("��λ����"), ""), Nvl(rsTemp!�������))
        End If
    End With
End Sub
Private Sub ��ѯ��λǷ��_�ɶ��ڽ�(ByVal str���˱�� As String, str�籣������ As String, strͳ����� As String)
    Dim StrInput As String, strOutput As String
    '    ���˱��    String (8)  IN
    '    �籣������  String (10) IN
    '    ͳ���������    String (1)  IN
    '   ��λǷ�����    String(1)   OUT
    StrInput = Rpad(str���˱��, 8)
    StrInput = StrInput & vbTab & Rpad(str�籣������, 8)
    StrInput = StrInput & vbTab & Rpad(strͳ�����, 1)
    If gobj�ɶ��ڽ� Is Nothing Then
        If ҽ����ʼ��_�ɶ��ڽ� = False Then Exit Sub
    End If
    If ҵ������_�ɶ��ڽ�(��ȡ��λǷ�����_�ڽ�, StrInput, strOutput) = False Then
        Exit Sub
    End If
    If Val(strOutput) = 1 Then
        MsgBox "�ò������ڵ�λ�Ѿ�Ƿ��!", vbInformation + vbDefaultButton1
        Exit Sub
    Else
        MsgBox "�ò������ڵ�λ��Ƿ��!", vbInformation + vbDefaultButton1
        Exit Sub
    End If
End Sub

'Modified By ���� 2004-05-25 ԭ��ҽ���ӿڱ䶯
'------------------------------------------------
Private Sub mnuEditReckoning_Click()
    Dim lng����ID  As Long, rsTemp As New ADODB.Recordset
    If mint���� <> TYPE_������ Then Exit Sub
    
    With msh�ʻ�_S
        '����ֱ�Ӵ��б���ȡ��
        lng����ID = Val(.TextMatrix(.Row, col����ID))
        
        If lng����ID <= 0 Then
            MsgBox "��ѡ��һλҽ�����ˡ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '�жϸò����Ƿ���סԺ��¼
        gstrSQL = "select A.��λ���� " & _
                  "  from �����ʻ� A " & _
                  "  Where A.����ID = " & lng����ID & " And A.���� = " & mint����
        Call OpenRecordset(rsTemp, Me.Caption)
        If rsTemp.EOF = True Then
            '�޷��Ӽ�¼����ȡ�ò�������
            MsgBox "���� " & msh�ʻ�_S.TextMatrix(.Row, col����) & " �޷��ҵ���Ч�ĵǼ���Ϣ��", vbInformation, gstrSysName
            Exit Sub
        End If
        
        Call �������㷽ʽ_����(lng����ID, Me, True)
    End With
End Sub
'------------------------------------------------

Private Sub mnuEditRollAdmit_Click()
    Dim lng����ID As Long
    Dim rsTemp As New ADODB.Recordset
    
    With msh�ʻ�_S
        '����ֱ�Ӵ��б���ȡ��
        lng����ID = Val(.TextMatrix(.Row, col����ID))
        If lng����ID <= 0 Then
            MsgBox "��ѡ��һλҽ�����ˡ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If MsgBox("���Ƿ�Ҫ�����ˡ�" & .TextMatrix(.Row, col����) & "���ļ���Ǽǳ�����", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
        
        '�жϸò����Ƿ���סԺ��¼
        gstrSQL = "select A.˳���,B.���� " & _
                  "  from �����ʻ� A,���ղ��� B " & _
                  "  Where A.����ID = " & lng����ID & " And A.����ID = B.ID And B.���� = " & mint����
        Call OpenRecordset(rsTemp, Me.Caption)
        If rsTemp.EOF = True Then
            '�޷��Ӽ�¼����ȡ�ò�������
            MsgBox "���� " & msh�ʻ�_S.TextMatrix(.Row, col����) & " �޷��ҵ���Ч�ĵǼ���Ϣ������δ�Լ��ﲡ�˵Ǽǡ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If IsNull(rsTemp("˳���")) = True Then
            MsgBox "���� " & msh�ʻ�_S.TextMatrix(.Row, col����) & " �Ǽ���Ϣ�����������ܼ����������Ǽǡ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If rsTemp("����") = "0090" Then
            If ��������Ǽ�_����(rsTemp("˳���"), mint����) = True Then
                MsgBox "�����ɹ��������ٴν��м���Ǽǡ�", vbInformation, gstrSysName
            End If
        End If
    End With

End Sub

Private Sub mnuEditRollIncome_Click()
    Dim lng����ID As Long, lng��ҳID  As Long
    Dim rsTemp As New ADODB.Recordset
    
    With msh�ʻ�_S
        '����ֱ�Ӵ��б���ȡ��
        lng����ID = Val(.TextMatrix(.Row, col����ID))
        If lng����ID <= 0 Then
            MsgBox "��ѡ��һλҽ�����ˡ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If MsgBox("���Ƿ�Ҫ�����ˡ�" & .TextMatrix(.Row, col����) & "����ҽ������תΪ��ͨ���ˣ�", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
        
        '�����ҳID
        gstrSQL = "select MAX(��ҳID) AS ��ҳID from ������ҳ where ����ID=" & lng����ID
        Call OpenRecordset(rsTemp, "�������")
        If IsNull(rsTemp("��ҳID")) = True Then
            MsgBox "ֻ��סԺ���˲ſ��Դ�ҽ������תΪ��ͨ���ˡ�", vbInformation, gstrSysName
            Exit Sub
        End If
        lng��ҳID = rsTemp("��ҳID")
        
        '�жϸò����Ƿ���סԺ��¼
        gstrSQL = "select A.˳��� " & _
                  "  from �����ʻ� A " & _
                  "  Where A.����ID = " & lng����ID & " And A.���� = " & mint����
        Call OpenRecordset(rsTemp, Me.Caption)
        If rsTemp.EOF = True Then
            '�޷��Ӽ�¼����ȡ�ò�������
            MsgBox "���� " & msh�ʻ�_S.TextMatrix(.Row, col����) & " �޷��ҵ���Ч�ĵǼ���Ϣ������δ�Լ��ﲡ�˵Ǽǡ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If mint���� = TYPE_������ Then
            If ����ҽ����Ժ_����(lng����ID, lng��ҳID, rsTemp("˳���")) = True Then
                MsgBox "�����ɹ����ò����Ѿ���ҽ������תΪ��ͨ���ˡ�", vbInformation, gstrSysName
            End If
        ElseIf mint���� = TYPE_�Ͻ� Then
            If Not ��Ժ�Ǽǳ���_�Ͻ�(lng����ID, lng��ҳID, True) Then Exit Sub
            gstrSQL = "ZL_������ҳ_����ҽ����Ժ(" & lng����ID & "," & lng��ҳID & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ����Ժ")
            MsgBox "�����ɹ����ò����Ѿ���ҽ������תΪ��ͨ���ˡ�", vbInformation, gstrSysName
        ElseIf mint���� = TYPE_����ʡ Or mint���� = TYPE_������ Then
            gstrSQL = "Select A.��Ժ����,A.��Ժ����,Decode(A.��Ժ��ʽ,'����',0,'����',1,'תԺ',2,9) as ��Ժ��ʽ,B.����,D.סԺ��,Sysdate as ����ʱ��," & _
                    " C.����,C.ҽ����,C.����,C.˳��� " & _
                    " From ������ҳ A,���ű� B,�����ʻ� C,������Ϣ D " & _
                    " Where A.����ID=D.����ID And A.����ID=" & lng����ID & " And A.��ҳID=" & lng��ҳID & _
                    " And A.��Ժ����ID=B.ID And A.����ID=C.����ID And C.����=" & mint����
            Call OpenRecordset(rsTemp, "ȡ˳���")
        
            If rsTemp.EOF Then
                MsgBox "û�д˲��˻�˲��˲���ҽ�����ˣ��޷������Ժ������", vbExclamation, gstrSysName
                Exit Sub
            End If
            If IsNull(rsTemp!˳���) Then
                MsgBox "δ���ָò��˵�סԺ����˳���,����ִ�н��ף�", vbInformation, gstrSysName
                Exit Sub
            End If
            
            If Not ��Ժ�Ǽ�_����(lng����ID, lng��ҳID, rsTemp!˳���, mint����, True, True) Then Exit Sub
            gstrSQL = "ZL_������ҳ_����ҽ����Ժ(" & lng����ID & "," & lng��ҳID & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ����Ժ")
            MsgBox "�����ɹ����ò����Ѿ���ҽ������תΪ��ͨ���ˡ�", vbInformation, gstrSysName
        ElseIf mint���� = TYPE_�����ɽ Then
            If ��Ժ�Ǽ�_��ɽ(lng����ID, lng��ҳID, True) Then
                gstrSQL = "ZL_������ҳ_����ҽ����Ժ(" & lng����ID & "," & lng��ҳID & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ����Ժ")
                MsgBox "�����ɹ����ò����Ѿ���ҽ������תΪ��ͨ���ˡ�", vbInformation, gstrSysName
            End If
        ElseIf mint���� = TYPE_�������� Then
            If ��Ժ�Ǽ�_��������(lng����ID, lng��ҳID) Then
                gstrSQL = "ZL_������ҳ_����ҽ����Ժ(" & lng����ID & "," & lng��ҳID & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ����Ժ")
                MsgBox "�����ɹ����ò����Ѿ���ҽ������תΪ��ͨ���ˡ�", vbInformation, gstrSysName
            End If
        ElseIf mint���� = TYPE_���������� Then
            If תΪ��ͨ����_����(lng����ID) Then
                MsgBox "�����ɹ����ò����Ѿ���ҽ������תΪ��ͨ���ˡ�", vbInformation, gstrSysName
            End If
        ElseIf mint���� = TYPE_�ɶ����� Then
            If ����ҽ����Ժ_�ɶ�����(lng����ID, lng��ҳID) = True Then
                MsgBox "�����ɹ����ò����Ѿ���ҽ������תΪ��ͨ���ˡ�", vbInformation, gstrSysName
            End If
        '������(2005-12-10):���ҽ����֧��ҽ��ת��ͨ����
        ElseIf mint���� = TYPE_��Ԫ���� Then
            If ����ҽ����Ժ_������(lng����ID, lng��ҳID, mint����) = True Then
                MsgBox "�����ɹ����ò����Ѿ���ҽ������תΪ��ͨ���ˡ�", vbInformation, gstrSysName
            End If
        '������(2005-12-10):���ҽ����֧��ҽ��ת��ͨ����
        ElseIf mint���� = TYPE_�ϳ����� Then
            If ����ҽ����Ժ_����(lng����ID, lng��ҳID, mint����) = True Then
                MsgBox "�����ɹ����ò����Ѿ���ҽ������תΪ��ͨ���ˡ�", vbInformation, gstrSysName
            End If
        '����(2007-10-19))
        ElseIf mint���� = TYPE_�Ͼ��� Then
            If ����ҽ����Ժ_�Ͼ���(lng����ID, lng��ҳID, rsTemp("˳���")) = True Then
                MsgBox "�����ɹ����ò����Ѿ���ҽ������תΪ��ͨ���ˡ�", vbInformation, gstrSysName
            End If
        'MODIFIED BY ZYB ����ҽ���ӿڿ���
        ElseIf mint���� = TYPE_���� Then
            'תΪ�ԷѲ���
            Call ҽ��ת��ͨ����_����(lng����ID, lng��ҳID)
        'Beging 2005-11-16-���ջ�
        ElseIf mint���� = type_ͭ����ҽ Then
            If ��Ժ�Ǽǳ���_ͭ����ҽ(lng����ID, lng��ҳID, mint����) Then
                MsgBox "�����ɹ����ò����Ѿ���ҽ������תΪ��ͨ���ˡ�", vbInformation, gstrSysName
            End If
        'End 2005-11-16-���ջ�
        ElseIf mint���� = TYPE_ɽ�� Then
            Call ������Ժ�Ǽ�_ɽ��(lng����ID, lng��ҳID)
        ElseIf mint���� = TYPE_����ũ�� Or mint���� = TYPE_ͭ�� Then
            gstrSQL = "ZL_������ҳ_����ҽ����Ժ(" & lng����ID & "," & lng��ҳID & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ����Ժ")
            MsgBox "�����ɹ����ò����Ѿ���ҽ������תΪ��ͨ���ˡ�", vbInformation, gstrSysName
        ElseIf mint���� = TYPE_���� Then
            Dim str������ As String
            Dim blnReturn As Boolean
            
            If Not ҽ����ʼ��_����() Then Exit Sub
            
            gstrSQL = "Select nvl(˳���,0) as ˳��� From �����ʻ� Where ����ID=" & lng����ID & " And ����=" & TYPE_����
            Call OpenRecordset(rsTemp, gstrSysName)
            str������ = rsTemp!˳���
            Call initType
            blnReturn = fl_dall(gstrҽ����������, gstrҽԺ����, str������, gstrOutPara)
            If blnReturn = False Then Exit Sub
            
            gstrSQL = "ZL_������ҳ_����ҽ����Ժ(" & lng����ID & "," & lng��ҳID & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ����Ժ")
            MsgBox "�����ɹ����ò����Ѿ���ҽ������תΪ��ͨ���ˡ�", vbInformation, gstrSysName
        ElseIf mint���� = TYPE_�������� Or mint���� = TYPE_����ʡ Or mint���� = TYPE_������ Or mint���� = TYPE_��ƽ�� Then
            If Not frm�ȴ���Ӧ.ShowME(mint����, ������ʽ.��Ժ, ����Ŀ��.����, lng����ID) Then Exit Sub
            gstrSQL = "ZL_������ҳ_����ҽ����Ժ(" & lng����ID & "," & lng��ҳID & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ����Ժ")
        Else
            If mint���� = TYPE_������ Or mint���� = TYPE_���������� Then
                If ��Ժ�Ǽǳ���_����(lng����ID, lng��ҳID, mint����) = True Then
                    gstrSQL = "ZL_������ҳ_����ҽ����Ժ(" & lng����ID & "," & lng��ҳID & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ����Ժ")
                    MsgBox "�����ɹ����ò����Ѿ���ҽ������תΪ��ͨ���ˡ�", vbInformation, gstrSysName
                End If
            End If
        End If
    End With

End Sub

Private Sub mnuEditSingleDisease_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim bln������ As Boolean
    Dim lng����ID As Long
    '�Ե����ֲ��˽����������
    With msh�ʻ�_S
        lng����ID = Val(.TextMatrix(.Row, col����ID))
        If lng����ID <= 0 Then
            MsgBox "��ѡ��һλҽ�����ˡ�", vbInformation, gstrSysName
            Exit Sub
        End If
    End With
    
    '��鵱ǰ����ѡ��Ĳ��֣��Ƿ����ھ��񲡣��������˳�
    gstrSQL = " Select Nvl(���,0) AS ��� From ���ղ��� Where ID=" & _
              "     (Select ����ID From �����ʻ� Where ����=" & TYPE_���������� & " And ����ID=" & lng����ID & ")" & _
              " And ����=" & TYPE_����������
    Call OpenRecordset(rsTemp, "��ȡ���ֵ�����")
    If rsTemp.RecordCount <> 0 Then
        bln������ = (rsTemp!��� = 4)
    End If
    If Not bln������ Then
        MsgBox "��ǰ���˵Ĳ��ֲ����ھ��񲡣������������ٵǼǣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '����Ƿ���Ժ
    gstrSQL = "Select Nvl(��ǰ״̬,0) AS ״̬ From �����ʻ� Where ����ID=" & lng����ID & " ANd ����=" & TYPE_����������
    Call OpenRecordset(rsTemp, "����Ƿ���Ժ")
    If rsTemp!״̬ = 0 Then
        MsgBox "ֻ�ܶ���Ժ���˽�����ٵǼǱ༭��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Call frm��ٱ༭.ShowEditor(lng����ID)
End Sub

Private Sub mnuEditVerify_Account_Click()
    Dim lng����ID As Long
    With msh�ʻ�_S
        lng����ID = Val(.TextMatrix(.Row, col����ID))
        If lng����ID <= 0 Then
            MsgBox "��ѡ��һλҽ�����ˡ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If mint���� = type_�ɶ����� Then
            Call �˶��ʻ�֧��_�ɶ�Ч��(lng����ID)
        ElseIf mint���� = TYPE_�¶� Then
            Call �˶��ʻ�֧��_�¶�(lng����ID)
        ElseIf mint���� = type_���� Then
            Call �˶��ʻ�֧��_����
        ElseIf mint���� = TYPE_�����山 Then
            Call �˶Բ����ʻ�֧����Ϣ_�����山(lng����ID)
        End If
    End With
End Sub

Private Sub mnuEditVerify_Click()
    '�����ж���ҽ�����еĹ��ܣ�����������Ŀ�����������շ���Ŀ��ѪҺ�׵��ף�
    Call frm������Ŀ����.ShowME(mint����)
End Sub

Private Sub mnuEditVerify_Detail_Click()
    Dim lng����ID As Long
    With msh�ʻ�_S
        lng����ID = Val(.TextMatrix(.Row, col����ID))
        If lng����ID <= 0 Then
            MsgBox "��ѡ��һλҽ�����ˡ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If mint���� = TYPE_�����山 Then
            Call �˶Դ�����ϸ��Ϣ_�����山(lng����ID)
        End If
    End With
End Sub

Private Sub mnuEditVerify_Hospital_Click()
    Dim lng����ID As Long
    With msh�ʻ�_S
        lng����ID = Val(.TextMatrix(.Row, col����ID))
        If lng����ID <= 0 Then
            MsgBox "��ѡ��һλҽ�����ˡ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If mint���� = type_�ɶ����� Then
            Call �˶����Ժ_�ɶ�Ч��(lng����ID)
        ElseIf mint���� = TYPE_�¶� Then
            Call �˶����Ժ_�¶�(lng����ID)
        ElseIf mint���� = TYPE_�����山 Then
            Call �˶Բ��˾�����Ϣ_�����山(lng����ID)
            
        End If
    End With
End Sub

Private Sub mnuEditVerify_Price_Click()
    Dim lng����ID As Long
    With msh�ʻ�_S
        If mint���� <> TYPE_���������� Then
            lng����ID = Val(.TextMatrix(.Row, col����ID))
            If lng����ID <= 0 Then
                MsgBox "��ѡ��һλҽ�����ˡ�", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        '��������ҽ�������� 204-03-31
        If mint���� = type_�ɶ����� Then
            Call �˶Է��ý���_�ɶ�Ч��(lng����ID)
        ElseIf mint���� = TYPE_�¶� Then
            Call �˶Է��ý���_�¶�(lng����ID)
        ElseIf mint���� = TYPE_���������� Then
            Call �˶Է��ý���_����������
        ElseIf mint���� = TYPE_�����山 Then
            Call �˶Բ��˷��ý��������Ϣ_�����山(lng����ID)
        End If
    End With
End Sub

Private Sub mnuEditVerify_UpDetail_Click()
    If mint���� = type_���� Then
        Call �����ϴ�������ϸ
    End If
End Sub

Private Sub mnuEditVerify_ZYPrice_Click()
    Dim lng����ID As Long
    
    If mint���� = type_���� Then
        Call �˶�סԺ����_����
    ElseIf mint���� = TYPE_�����山 Then
        lng����ID = Val(msh�ʻ�_S.TextMatrix(msh�ʻ�_S.Row, col����ID))
        If lng����ID = 0 Then Exit Sub
        Call �˶Է��ý�����_�����山(lng����ID)
    End If
End Sub

Private Sub mnuEditXE_Click()
    '��Ҫ¼�����������޶�
    Dim lng����ID As Long
    Dim strIdentify As String
    Dim bytType As Byte
    
    lng����ID = Val(msh�ʻ�_S.TextMatrix(msh�ʻ�_S.Row, col����ID))
    If lng����ID = 0 Then Exit Sub
    bytType = 9
'$IF HIS9.16
#If gverControl = 0 Then
    strIdentify = gclsInsure.Identify(bytType, lng����ID)
'$ELSE
#Else
    strIdentify = gclsInsure.Identify(bytType, lng����ID, mint����)
#End If
'$END IF
    If strIdentify <> "" Then
        Call FillList
    End If
    
End Sub

Private Sub mnuFileCard_Click()
    Dim strҽ���� As String
    '��ӡ��Ƭ
    strҽ���� = Trim(msh�ʻ�_S.TextMatrix(msh�ʻ�_S.Row, colҽ����))
    If strҽ���� = "" Then Exit Sub
    Call zl9Report.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1604", Me, "����=" & mint����, "ҽ����=" & strҽ����, 2)
End Sub

Private Sub msh���_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    msh���.ToolTipText = msh���.TextMatrix(msh���.MouseRow, msh���.MouseCol)
End Sub

Private Sub msh�ʻ�_S_Scroll()
    Call GetAccountInfo
End Sub

Private Sub picOther_Resize()
    msh���.Left = 0
    msh���.Width = picOther.ScaleWidth
    
    msh�����Ϣ.Left = 0
    msh�����Ϣ.Width = picOther.ScaleWidth
    If picOther.ScaleHeight - msh�����Ϣ.Top > 0 Then
        msh�����Ϣ.Height = picOther.ScaleHeight - msh�����Ϣ.Top
    End If
End Sub

Private Sub picSplitV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        msngStartX = x
    End If
End Sub

Private Sub picSplitV_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sngTemp As Single
    If Button = 1 Then
        sngTemp = picSplitV.Left + x - msngStartX
        If sngTemp > msh�ʻ�_S.Left + 2000 And ScaleWidth - (sngTemp + picSplitV.Width) > 1000 Then
            picSplitV.Left = sngTemp
            msh�ʻ�_S.Width = picSplitV.Left - msh�ʻ�_S.Left
            picOther.Left = sngTemp + picSplitV.Width
            picOther.Width = ScaleWidth - (sngTemp + picSplitV.Width)
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub lbl��ֵ_Click(Index As Integer)
End Sub

Private Sub mnuEditModify_Click()
'�����֤
    Dim strIdentify As String
    Dim bytType As Byte
    
    bytType = 2
'$IF HIS9.19
#If gverControl = 0 Then
    strIdentify = gclsInsure.Identify(bytType, 0)
'$ELSE
#Else
    strIdentify = gclsInsure.Identify(bytType, 0, mint����)
#End If
'EnD IF
    If strIdentify <> "" Then
        Call FillList
    End If
End Sub

Private Sub mnuEditSub_Click()
    With frmҽ���ʻ�����Ժ
        .mint���� = mint����
        .Show vbModal, Me
    End With
End Sub

Private Sub mnuHelpTitle_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub mnuFileQuit_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuViewFind_Click()
    If frmҽ���ʻ�����.GetFind(mstrFind) = False Then
        Exit Sub
    End If
    
    Call FillList
End Sub

Private Sub mnuViewRefresh_Click()
    Call FillList
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
    For i = 1 To tbrThis.Buttons.Count
        tbrThis.Buttons(i).Caption = IIf(mnuViewToolText.Checked, tbrThis.Buttons(i).Tag, "")
    Next
    cbr.Bands(1).MinHeight = tbrThis.ButtonHeight
    Form_Resize
End Sub

Private Sub mnuViewCustom_Click()
    If frmҽ���ʻ���Ϣ����.SelectFields() = True Then
        'ȡע���
        mstr�����ֶ� = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "�����ֶ�", "")
        mstr�����ֶ� = Replace(mstr�����ֶ�, "'", "")
        
        Call Fill�ʻ������Ϣ
    End If
End Sub

Private Sub msh�ʻ�_S_EnterCell()
    Dim lng����ID As Long
    Dim rsAccount As New ADODB.Recordset
    'ѡ��ĳ���ʻ�,����ȡ�����Ϣ
    Call Fill�ʻ������Ϣ
    If mint���� = TYPE_�Ĵ�üɽ Then
        'ɾ���ʻ���Ϣ(�Ҷȼ�:0-����;1-��ֹ�����ʻ�;9-�ʻ��ѳ���)
        lng����ID = Val(msh�ʻ�_S.TextMatrix(msh�ʻ�_S.Row, col����ID))
        If lng����ID = 0 Then Exit Sub
        
        '��鿨��״̬������Ѿ����������ʾҪ���������򽫷����ÿ�
        gstrSQL = "Select Nvl(�Ҷȼ�,0) ״̬ From �����ʻ� Where ����ID=" & lng����ID & " And ����=" & TYPE_�Ĵ�üɽ
        Call OpenRecordset(rsAccount, Me.Caption)
        If rsAccount!״̬ = 0 Then
            mnuEditLoss.Caption = "����ҽ����(&L)"
        Else
            mnuEditLoss.Caption = "�������(&L)"
        End If
    End If
End Sub

Private Sub msh�ʻ�_S_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim strSort As String
    
    If Button = 1 Then
        '����ͷ����
        If msh�ʻ�_S.MouseRow = 0 Then
            strSort = msh�ʻ�_S.TextMatrix(0, msh�ʻ�_S.MouseCol)
            
            If strSort = "" Then Exit Sub
            If mrs�ʻ�.Sort = strSort Then
                mrs�ʻ�.Sort = strSort & " DESC"
            Else
                mrs�ʻ�.Sort = strSort
            End If
            Call ������
        End If
    
    End If
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuFileQuit_Click
        Case "Modify"
            mnuEditModify_Click
        Case "Find"
            mnuViewFind_Click
        Case "Custom"
            mnuViewCustom_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Preview"
            mnuFilePreview_Click
        Case "Help"
            mnuHelpTitle_Click
    End Select
End Sub

Private Sub tbrThis_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
End Sub

Private Sub mnuFileExcel_Click()
    Call subPrint(3)
End Sub

Private Sub mnuFilePreview_Click()
    Call subPrint(2)
End Sub

Private Sub mnuFilePrint_Click()
    Call subPrint(1)
End Sub

Private Sub mnuFilePrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub subPrint(bytStyle As Byte)
'���ܣ�������б�
'������bytStyle=1-��ӡ,2-Ԥ��,3-�����Excel
    Dim objOut As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim bytR As Byte, intRow As Integer
    
    If gstrUserName = "" Then GetUserInfo
    intRow = msh�ʻ�_S.Row
    
    '��ͷ
    objOut.Title.Text = "ҽ���ʻ��嵥"
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '����
    objRow.Add "ҽ�����" & cmb����.Text
    objOut.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate, "yyyy��MM��DD��")
    objOut.BelowAppRows.Add objRow
    
    '����
    Set objOut.Body = msh�ʻ�_S
    
    '���
    msh�ʻ�_S.Redraw = False
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    msh�ʻ�_S.Redraw = True
    
    msh�ʻ�_S.Row = intRow
    msh�ʻ�_S.COL = 0: msh�ʻ�_S.ColSel = msh�ʻ�_S.Cols - 1
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage Me.hwnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo Me.hwnd
End Sub

Private Function FillList()
    '��ȡ�����ʻ�(���Ȩ������,����������ֶ�)������
    If mrs�ʻ�.State = adStateOpen Then mrs�ʻ�.Close
    Dim str����֤�� As String
    Dim str��λ���� As String
    On Error GoTo errHand
    
    str��λ���� = "A.��λ����"
    Select Case mint����
        Case TYPE_������
            '����ҽ������֤�������ֱ��뱣��
            str����֤�� = "A.����֤�� AS ���ֱ���"
        Case TYPE_������, TYPE_����������
            str����֤�� = "A.����֤�� as ���������ʻ����"
            str��λ���� = "decode(A.��λ����,'0','��','1','�±�','����') as �α����"
        Case Else
            str����֤�� = "A.����֤��"
    End Select
    
    If mcol����("K" & mint����) = "1" Then
        '����ҽ������
        gstrSQL = " Select C.���� as ����,A.����,A.ҽ����,P.����ID,P.����,P.�Ա�,To_Char(P.��������,'yyyy-MM-dd') as  ��������,P.���֤�� " & _
                  "        ,E.���� ��Ա���,A.��Ա��� as ��ݱ���," & str��λ���� & "," & str����֤�� & ",D.���� as ����,Decode(A.��ǰ״̬,0,'��ͨ','��Ժ') as ״̬,A.�ʻ����,to_char(A.����ʱ��,'yyyy-MM-dd') as ����ʱ��  " & _
                  " " & IIf(mint���� = TYPE_�Ĵ�üɽ, ",��ע ������Ϣ", "") & IIf(mint���� = TYPE_������, ",A.����֢", "") & _
                  " From �����ʻ� A,������Ϣ P,��������Ŀ¼ C,���ղ��� D,������Ⱥ E " & _
                  " Where A.����ID = P.����ID and A.����=C.���� and A.����=C.��� " & IIf(mint���� = TYPE_�Ĵ�üɽ, " And Nvl(A.�Ҷȼ�,0)<>9", "") & _
                  "       And A.����=E.���� and A.��ְ=E.��� And A.����ID=D.ID(+) And A.����=" & mint���� & _
                  mstrFind & " Order by C.����,A.����"
    Else
        Select Case mint����
        Case TYPE_������, TYPE_����������
            gstrSQL = " Select '' as ����,A.����,A.ҽ����,P.����ID,P.����,P.�Ա�,To_Char(P.��������,'yyyy-MM-dd') as  ��������,P.���֤�� " & _
                      "        ,E.���� ��Ա���,A.��Ա��� as ��ݱ���," & str��λ���� & "," & str����֤�� & ",D.���� as ����," & _
                      "        decode(�α����1, 0,'�����ܸ߶�',1,'���ܸ߶�','ҽ�Ʊ��ղ�����') as �α����1," & _
                      "        decode(�α����2, 0,'������',1,'��ҵ','����Ա') as �α����2," & _
                      "        decode(�α����3, 0,'��','�±�') as �α����3," & _
                      "        decode(�α����4, 0,'����������',1,'��������','����������') as �α����4," & _
                      "        decode(�α����5, 0,'���˲�����',1,'���˿���','���˲�����') as �α����5," & _
                      "        to_char(����޶�,'90009000900099.99') as ����޶�," & _
                      "         Decode(A.��ǰ״̬,0,'��ͨ','��Ժ') as ״̬,A.�ʻ����,to_char(A.����ʱ��,'yyyy-MM-dd') as ����ʱ��  " & _
                      " " & IIf(mint���� = TYPE_�Ĵ�üɽ, ",��ע ������Ϣ", "") & IIf(mint���� = TYPE_������, ",A.����֢", "") & _
                      " From �����ʻ� A,������Ϣ P,���ղ��� D,������Ⱥ E " & _
                      " Where A.����ID = P.����ID And A.����=E.���� and A.��ְ=E.��� " & IIf(mint���� = TYPE_�Ĵ�üɽ, " And Nvl(A.�Ҷȼ�,0)<>9", "") & _
                      "       And A.����ID=D.ID(+) And A.����=" & mint���� & mstrFind & " Order by A.����"
        Case Else
            gstrSQL = " Select '' as ����,A.����,A.ҽ����,P.����ID,P.����,P.�Ա�,To_Char(P.��������,'yyyy-MM-dd') as  ��������,P.���֤�� " & _
                      "        ,E.���� ��Ա���,A.��Ա��� as ��ݱ���," & str��λ���� & "," & str����֤�� & ",D.���� as ����,Decode(A.��ǰ״̬,0,'��ͨ','��Ժ') as ״̬,A.�ʻ����,to_char(A.����ʱ��,'yyyy-MM-dd') as ����ʱ��  " & _
                      " " & IIf(mint���� = TYPE_�Ĵ�üɽ, ",��ע ������Ϣ", "") & IIf(mint���� = TYPE_������, ",A.����֢", "") & _
                      " From �����ʻ� A,������Ϣ P,���ղ��� D,������Ⱥ E " & _
                      " Where A.����ID = P.����ID And A.����=E.���� and A.��ְ=E.��� " & IIf(mint���� = TYPE_�Ĵ�üɽ, " And Nvl(A.�Ҷȼ�,0)<>9", "") & _
                      "       And A.����ID=D.ID(+) And A.����=" & mint���� & mstrFind & " Order by A.����"
        End Select
    End If
    Call OpenRecordset(mrs�ʻ�, Me.Caption)
    
    Call ������
    Call Fill�ʻ������Ϣ
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub ������()
    Dim lngCol As Long
    
    '���ʻ�����װ��FLEXGRID��������
    With msh�ʻ�_S
        If mrs�ʻ�.RecordCount <> 0 Then
            Set .DataSource = mrs�ʻ�
            
            DoEvents
            .COL = 0
            .Row = .FixedRows - 1
            .ColSel = .Cols - 1
            .RowSel = .Row
            .FillStyle = flexFillRepeat
            .CellAlignment = 4
            .FillStyle = flexFillSingle
            .AllowBigSelection = False
            .Row = .FixedRows: .COL = 0
            .ColSel = .Cols - 1: .RowSel = .Row
            
        Else
            Set .DataSource = Nothing
            .Rows = 2
            For lngCol = 0 To .Cols - 1
                .TextMatrix(1, lngCol) = ""
            Next
        End If
        
        '���ض������
        If mcol����("K" & mint����) = "0" Then
            .ColWidth(col����) = 0
        Else
            If .ColWidth(col����) = 0 Then
                .ColWidth(col����) = 1000
            End If
        End If
        
        .ColWidth(col����ID) = 0
        .ColWidth(col�ʻ����) = 0
        If mint���� = TYPE_�Ĵ�üɽ Then .ColWidth(.Cols - 1) = 1200
    End With
    Call SetMenu
End Sub

Private Sub InitTable()
    Dim lngCol As Integer
    '���ø�ʽ
    With msh�ʻ�_S
        .Rows = 2
        .Cols = 16
        If mint���� = TYPE_���������� Or mint���� = TYPE_������ Then
            .Cols = .Cols + 6
        End If
        For lngCol = 0 To .Cols - 1
            .TextMatrix(1, lngCol) = ""
        Next
        
        If mblnLoad Then
            .TextMatrix(0, col����) = "����"
            .TextMatrix(0, col����) = "����"
            .TextMatrix(0, colҽ����) = "ҽ����"
            .TextMatrix(0, col����ID) = "����ID"
            .TextMatrix(0, col����) = "����"
            .TextMatrix(0, col�Ա�) = "�Ա�"
            .TextMatrix(0, col��������) = "��������"
            .TextMatrix(0, col���֤��) = "���֤��"
            .TextMatrix(0, col��Ա���) = "��Ա���"
            .TextMatrix(0, col��ݱ���) = "��ݱ���"
            .TextMatrix(0, col��λ����) = "��λ����"
            .TextMatrix(0, col����֤��) = "����֤��"
            .TextMatrix(0, col����) = "����"
            lngCol = 0
            If mint���� = TYPE_���������� Or mint���� = TYPE_������ Then
                .TextMatrix(0, col���� + 1) = "�α����1"
                lngCol = lngCol + 1
                .TextMatrix(0, col���� + 2) = "�α����2"
                lngCol = lngCol + 1
                .TextMatrix(0, col���� + 3) = "�α����3"
                lngCol = lngCol + 1
                .TextMatrix(0, col���� + 4) = "�α����4"
                lngCol = lngCol + 1
                .TextMatrix(0, col���� + 5) = "�α����5"
                lngCol = lngCol + 1
                .TextMatrix(0, col���� + 6) = "����޶�"
                lngCol = lngCol + 1
            End If
            .TextMatrix(0, col״̬ + lngCol) = "״̬"
            .TextMatrix(0, col�ʻ���� + lngCol) = "�ʻ����"
            .TextMatrix(0, col����ʱ�� + lngCol) = "����ʱ��"
            .ColWidth(col����) = 0
            .ColWidth(col����) = 900
            .ColWidth(colҽ����) = 900
            .ColWidth(col����ID) = 0
            .ColWidth(col����) = 800
            .ColWidth(col�Ա�) = 400
            .ColWidth(col��������) = 1200
            .ColWidth(col���֤��) = 1400
            .ColWidth(col��Ա���) = 800
            .ColWidth(col��ݱ���) = 600
            .ColWidth(col��λ����) = 600
            .ColWidth(col����֤��) = 900
            .ColWidth(col����) = 800
            If mint���� = TYPE_���������� Or mint���� = TYPE_������ Then
                .ColWidth(col���� + 1) = 800
                .ColWidth(col���� + 2) = 800
                .ColWidth(col���� + 3) = 800
                .ColWidth(col���� + 4) = 800
                .ColWidth(col���� + 5) = 800
                .ColWidth(col���� + 6) = 800
                .ColWidth(col״̬ + 6) = 800
                .ColAlignment(col״̬ + 6) = 7
                .ColWidth(col�ʻ���� + 6) = 0
                .ColWidth(col����ʱ�� + 6) = 1400
            Else
                .ColWidth(col״̬) = 800
                .ColWidth(col�ʻ����) = 0
                .ColWidth(col����ʱ��) = 1400
                
            End If
            .ColWidth(col״̬) = 800
            .ColWidth(col�ʻ����) = 0
            .ColWidth(col����ʱ��) = 1400
        End If
        
        For lngCol = 0 To .Cols - 1
            .ColAlignmentFixed(lngCol) = 4
        Next
        
        .COL = 0
        .ColSel = .Cols - 1
    End With
    
    With msh���
        .Rows = 13: .Cols = 2
        .ColWidth(0) = 1600: .ColAlignment(0) = 1
        .ColWidth(1) = 1000: .ColAlignment(1) = 7
        
        .TextMatrix(0, 0) = "�����Ϣ": .TextMatrix(0, 1) = "ֵ"
        
        .TextMatrix(rowסԺ����, 0) = "סԺ����"
        .TextMatrix(row�ʻ����, 0) = "�ʻ����"
        .TextMatrix(row�ʻ�����, 0) = "�ʻ������ۼ�"
        .TextMatrix(row�ʻ�֧��, 0) = "�ʻ�֧���ۼ�"
        .TextMatrix(row��������, 0) = "��������"
        .TextMatrix(row�����ۼ�, 0) = "֧�������ۼ�"
        .TextMatrix(rowͳ���޶�, 0) = "����ͳ��֧���޶�"
        .TextMatrix(row����ͳ��, 0) = "�������ͳ���ۼ�"
        .TextMatrix(rowͳ�ﱨ��, 0) = "֧������ͳ���ۼ�"
        .TextMatrix(row����޶�, 0) = "���ͳ��֧���޶�"
        .TextMatrix(row����ۼ�, 0) = "���ͳ��֧���ۼ�"
        .TextMatrix(row������Ϣ, 0) = "������Ϣ"
    End With
End Sub

Private Function Fill�ʻ������Ϣ()
    Dim lngCount As Long, lng����ID As Long
    Dim arrayCol, strColumn As String, intColumn As Integer
    Dim rsTemp As New ADODB.Recordset
    
    '��������Ϣ
    Call ClearOther
    
    lng����ID = Val(msh�ʻ�_S.TextMatrix(msh�ʻ�_S.Row, col����ID))
    If lng����ID = 0 Then
        Exit Function
    End If
    
    '��ȡָ���ʻ��������Ϣ
    strColumn = ""
    arrayCol = Split(mstr�����ֶ�, ",")
    For intColumn = 0 To UBound(arrayCol)
        strColumn = strColumn & ",P." & arrayCol(intColumn)
    Next
    
    'If InStr(1, strColumn, "P.����") <> 0 Then strColumn = Replace(strColumn, "P.����", "trunc(Months_between(to_Date(to_Char(sysdate,'yyyy')||'-01'||'-01','yyyy-MM-dd'),Decode(P.��������,NULL,P.�Ǽ�ʱ��,P.��������))/12) ����")
    gstrSQL = " Select P.��������,P.������λ,P.����״��" & strColumn & _
              " From �����ʻ� A,������Ϣ P " & _
              " Where A.����ID = P.����ID " & _
              "       And A.����=" & mint���� & " And A.����ID=" & lng����ID
    Call OpenRecordset(rsTemp, Me.Caption)
    If rsTemp.RecordCount > 0 Then
        With msh�����Ϣ
            For lngCount = 1 To .Rows - 1
                .TextMatrix(lngCount, 1) = IIf(IsNull(rsTemp.Fields(.TextMatrix(lngCount, 0)).Value), "", rsTemp.Fields(.TextMatrix(lngCount, 0)).Value)
            Next
        End With
    End If
    
    If rsTemp.State = adStateOpen Then rsTemp.Close
    gstrSQL = " Select * " & _
              " From �ʻ������Ϣ Y" & _
              " Where Y.����=" & mint���� & " And Y.���=" & lbl���.Caption & " And Y.����ID=" & lng����ID
    Call OpenRecordset(rsTemp, Me.Caption)
    
    If rsTemp.RecordCount > 0 Then
        'װ��ָ���ʻ����������
        With msh���
            .TextMatrix(rowסԺ����, 1) = Format(rsTemp("סԺ�����ۼ�"), "#####;-#####; ;")
            .TextMatrix(row�ʻ����, 1) = Format(Val(msh�ʻ�_S.TextMatrix(msh�ʻ�_S.Row, col�ʻ����)), "#####0.00;-#####0.00; ;")
            .TextMatrix(row�ʻ�����, 1) = Format(rsTemp("�ʻ������ۼ�"), "#####0.00;-#####0.00; ;")
            .TextMatrix(row�ʻ�֧��, 1) = Format(rsTemp("�ʻ�֧���ۼ�"), "#####0.00;-#####0.00; ;")
            .TextMatrix(row��������, 1) = Format(rsTemp("��������"), "#####0.00;-#####0.00; ;")
            .TextMatrix(row�����ۼ�, 1) = Format(rsTemp("�����ۼ�"), "#####0.00;-#####0.00; ;")
            .TextMatrix(rowͳ���޶�, 1) = Format(rsTemp("����ͳ���޶�"), "#####0.00;-#####0.00; ;")
            .TextMatrix(row����ͳ��, 1) = Format(rsTemp("����ͳ���ۼ�"), "#####0.00;-#####0.00; ;")
            .TextMatrix(rowͳ�ﱨ��, 1) = Format(rsTemp("ͳ�ﱨ���ۼ�"), "#####0.00;-#####0.00; ;")
            .TextMatrix(row����޶�, 1) = Format(rsTemp("���ͳ���޶�"), "#####0.00;-#####0.00; ;")
            .TextMatrix(row����ۼ�, 1) = Format(rsTemp("���ͳ���ۼ�"), "#####0.00;-#####0.00; ;")
            .TextMatrix(row������Ϣ, 1) = IIf(IsNull(rsTemp("������Ϣ")), "", rsTemp("������Ϣ"))
        End With
    End If
End Function

Private Sub ClearOther()
    Dim lngCount As Long
    
    '��������Ϣ
    With msh���
        For lngCount = 1 To .Rows - 1
            .TextMatrix(lngCount, 1) = ""
        Next
    End With
        
    With msh�����Ϣ
        .ColWidth(0) = 1170
        .ColWidth(1) = 1380
        .TextMatrix(0, 0) = "����"
        .TextMatrix(0, 1) = "����"
        .ColAlignment(1) = 1
        
        .Rows = 5 + UBound(Split(mstr�����ֶ�, ",")) '�б���,��ʼ����,���û���Ҫ
        .TextMatrix(1, 0) = "��������"
        .TextMatrix(2, 0) = "������λ"
        .TextMatrix(3, 0) = "����״��"
        .TextMatrix(1, 1) = ""
        .TextMatrix(2, 1) = ""
        .TextMatrix(3, 1) = ""
        
        For lngCount = 4 To .Rows - 1
            .TextMatrix(lngCount, 0) = Split(mstr�����ֶ�, ",")(lngCount - 4)
            .TextMatrix(lngCount, 1) = ""
        Next
    End With
End Sub

Private Sub Ȩ�޿���()
    If InStr(gstrPrivs, "��ɾ��") = 0 Then
        mnuEdit.Visible = False
        mnuEditModify.Enabled = False
        mnuEditPassword.Enabled = False
        mnuEditSub.Enabled = False
        
        mnuEditPassword.Visible = False
        mnuEditSub.Visible = False
        tbrThis.Buttons("Modify").Visible = False
        tbrThis.Buttons("SplitModify").Visible = False
    End If
    
    mnuEditXE.Visible = False ' mint���� = TYPE_������ Or mint���� = TYPE_����������
    mnuEditSp.Visible = mnuEditXE.Visible
    
       
    mnuEditSub.Visible = InStr(1, ";" & gstrPrivs & ";", ";������Ժ;") <> 0
    mnuEditRollIncome.Visible = InStr(1, ";" & gstrPrivs & ";", ";������Ժ;") <> 0
    mnuEditOut.Visible = InStr(1, ";" & gstrPrivs & ";", ";�����Ժ;") <> 0
    mnuEditVerify.Visible = InStr(1, ";" & gstrPrivs & ";", ";������Ŀ����;") <> 0
    mnuEditSplit5.Visible = mnuEditVerify.Visible
End Sub

Private Sub SetMenu()
    Dim blnData As Boolean
        
    blnData = (mrs�ʻ�.RecordCount > 0)
    stbThis.Panels(2).Text = "��ǰ����" & mrs�ʻ�.RecordCount & "��ҽ���ʻ�"
    
    tbrThis.Buttons("Print").Enabled = blnData
    tbrThis.Buttons("Preview").Enabled = blnData
    
    mnuFilePreview.Enabled = blnData
    mnuFilePrint.Enabled = blnData
    mnuFileExcel.Enabled = blnData
    
    mnuEditDelete.Enabled = blnData
    mnuEditMend.Enabled = blnData
    mnuEditLoss.Enabled = blnData
End Sub

Public Sub ShowForm(frmParent As Form)
    Dim rsTemp As New ADODB.Recordset
    Dim blnCanUse As Boolean
    
    gstrSQL = "select ���,����,nvl(��������,0) as �������� from ������� where nvl(�Ƿ��ֹ,0)<>1 And ҽ������ Is NULL order by ���"
    Call OpenRecordset(rsTemp, "�����ʻ�")
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "û�п��ñ�����𣬲���ʹ�ñ����ܡ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If frmҽ���ʻ�.Visible = True Then
        frmҽ���ʻ�.Show
        Exit Sub
    End If
    
    Set mcol���� = New Collection
    
    With cmb����
        .Clear
        Do Until rsTemp.EOF
            .AddItem rsTemp("����")
            .ItemData(.NewIndex) = rsTemp("���")
            mcol����.Add Val(rsTemp("��������")), "K" & rsTemp("���")
            If rsTemp("���") = mint���� Then
                '��ǰҽ����
                'ʹ��API�����Բ�����Click�¼�
                zlControl.CboSetIndex .hwnd, .NewIndex
            End If
            
            rsTemp.MoveNext
        Loop
        If .ListIndex < 0 Then
            'ʹ��API�����Բ�����Click�¼�
            zlControl.CboSetIndex .hwnd, 0
        End If
        
        mint���� = .ItemData(.ListIndex)
    End With
        
    Call SetMenuState
        
    tbrThis.Buttons("Modify").Enabled = mnuEditModify.Enabled
    frmҽ���ʻ�.Show , frmParent
End Sub

Public Function CheckForm() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim blnCanUse As Boolean
    
    gstrSQL = "select ���,����,nvl(��������,0) as �������� from ������� where nvl(�Ƿ��ֹ,0)<>1 And ҽ������ Is NULL order by ���"
    Call OpenRecordset(rsTemp, "�����ʻ�")
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "û�п��ñ�����𣬲���ʹ�ñ����ܡ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    If frmҽ���ʻ�.Visible = True Then
        CheckForm = True
        Exit Function
    End If
    
    Set mcol���� = New Collection
    
    With cmb����
        .Clear
        Do Until rsTemp.EOF
            .AddItem rsTemp("����")
            .ItemData(.NewIndex) = rsTemp("���")
            mcol����.Add Val(rsTemp("��������")), "K" & rsTemp("���")
            If rsTemp("���") = mint���� Then
                '��ǰҽ����
                'ʹ��API�����Բ�����Click�¼�
                zlControl.CboSetIndex .hwnd, .NewIndex
            End If
            
            rsTemp.MoveNext
        Loop
        If .ListIndex < 0 Then
            'ʹ��API�����Բ�����Click�¼�
            zlControl.CboSetIndex .hwnd, 0
        End If
        
        mint���� = .ItemData(.ListIndex)
    End With
        
    Call SetMenuState
        
    tbrThis.Buttons("Modify").Enabled = mnuEditModify.Enabled
    CheckForm = True
End Function


Private Sub SetMenuState()
    Dim blnCanUse As Boolean
    blnCanUse = GetInsureInit(mint����)
    mnuEditSub.Enabled = blnCanUse
    mnuEditDisease.Enabled = blnCanUse
    mnuEditRollIncome.Enabled = blnCanUse
    mnuEditRollAdmit.Enabled = blnCanUse
    mnuEditQuery.Enabled = blnCanUse
    
    mnuEditPassword.Visible = (mint���� = TYPE_�Թ��� Or mint���� = TYPE_������ Or mint���� = type_�ɶ����� Or mint���� = TYPE_��������)
    mnuEditQuery.Visible = (mint���� = TYPE_������) Or (mint���� = TYPE_�ɶ��ڽ�)
    mnuEditSingleDisease.Visible = (mint���� = TYPE_����������)
    mnuEditBed.Visible = (mint���� = TYPE_������)
    If mint���� = TYPE_�������� Then
        mnuEditPassword.Caption = "�޸Ŀ���(&M)"
    End If
    
    'Modified By ���� 2004-05-25 ԭ��ҽ���ӿڱ䶯
    '------------------------------------------------
    mnuEditReckoning.Visible = (mint���� = TYPE_������)
    '------------------------------------------------
    mnuEditDisease.Visible = (mint���� = TYPE_������ Or mint���� = TYPE_���������� Or _
                               mint���� = TYPE_�Թ��� Or mint���� = TYPE_������ Or _
                               mint���� = TYPE_������ Or mint���� = TYPE_�Ͻ� Or _
                               mint���� = TYPE_������ Or mint���� = TYPE_ɽ�� Or _
                               mint���� = TYPE_���Ͻ�ˮ Or mint���� = TYPE_�ɶ��ڽ� Or _
                               mint���� = type_ͭ����ҽ)
    mnuEditICD.Visible = (mint���� = TYPE_������)
    'Modified by ���� 20031218 ����������
    'Modified by л�� 20100810 ��������ɽ
    mnuEditRollIncome.Enabled = (mint���� = TYPE_������ Or mint���� = TYPE_���������� _
                             Or mint���� = TYPE_�������� Or mint���� = TYPE_��ƽ�� _
                             Or mint���� = TYPE_����ʡ Or mint���� = TYPE_������ _
                             Or mint���� = TYPE_������ Or mint���� = TYPE_���������� _
                             Or mint���� = TYPE_���� Or mint���� = TYPE_���� _
                             Or mint���� = TYPE_ɽ�� Or mint���� = TYPE_�ɶ����� Or mint���� = TYPE_��Ԫ���� Or mint���� = TYPE_�Ͼ��� _
                             Or mint���� = type_ͭ����ҽ Or mint���� = TYPE_ͭ�� Or mint���� = TYPE_�Ͻ� Or mint���� = TYPE_��ɽ)
    'Modified by ���� 20031218 ����������
    mnuEditOut.Enabled = (mint���� = TYPE_�������� Or mint���� = TYPE_��ƽ�� Or mint���� = TYPE_����ʡ Or _
    mint���� = TYPE_������ Or mint���� = TYPE_������ Or mint���� = TYPE_����ʡ Or _
    mint���� = TYPE_������ Or mint���� = TYPE_���������� Or mint���� = TYPE_���� Or mint���� = TYPE_��Ԫ���� Or mint���� = TYPE_������)
    mnuEditSplit0.Visible = mnuEditPassword.Visible
    mnuEditSplit1.Visible = mnuEditDisease.Visible
    mnuEditRollAdmit.Visible = (mint���� = TYPE_����ʡ Or mint���� = TYPE_������) '���ڲ���ʹɾ��δ����ã����Բ�֧�� TYPE_���Ͻ�ˮ
    
    mnuEditXE.Visible = False ' mint���� = TYPE_������ Or mint���� = TYPE_����������
    mnuEditSp.Visible = mnuEditXE.Visible
    
    If TYPE_�Ĵ�üɽ = mint���� Then
        mnuEditSplit4.Visible = True
        mnuEditDelete.Visible = True
        mnuEditMend.Visible = True
        mnuEditLoss.Visible = True
        mnuFileCard.Visible = True
        mnuFileSplit2.Visible = True
    End If
    
    mnuEditModify.Enabled = True
    mnuEditPassword.Enabled = True
'
'    blnCanUse = GetInsureInit(mint����)
'    mnuEditSub.Enabled = blnCanUse
'    mnuEditDisease.Enabled = blnCanUse
'    mnuEditRollIncome.Enabled = blnCanUse
'    mnuEditRollAdmit.Enabled = blnCanUse
'    mnuEditQuery.Enabled = blnCanUse
    
    '����ҽ��֧�ֺ˶�
    mnuEditSplit3.Visible = (mint���� = type_�ɶ����� Or mint���� = TYPE_�¶� Or mint���� = type_�ɶ�����)
    mnuEditVerify_Account.Visible = (mint���� = type_�ɶ����� Or mint���� = TYPE_�¶� Or mint���� = type_�ɶ�����)
    mnuEditVerify_Hospital.Visible = (mint���� = type_�ɶ����� Or mint���� = TYPE_�¶� Or mint���� = type_�ɶ�����)
    mnuEditVerify_Price.Visible = (mint���� = type_�ɶ����� Or mint���� = TYPE_�¶� Or mint���� = type_�ɶ�����)
'
'    '������(2005-08-19):�ɶ���ҽ���޳�����Ժ�Ͳ����Ժ�Ǽǹ��ܣ����Ρ�
'    If mint���� = 20 Then
'       mnuEditRollIncome.Visible = False
'       mnuEditOut.Visible = False
'    End If
'    If mint���� = type_�ɶ����� Then
'        mnuEditSplit3.Visible = True
'        mnuEditVerify_Account.Visible = True
''        mnuEditVerify_Detail.Visible = True
'        mnuEditVerify_Hospital.Visible = True
'        mnuEditVerify_Price.Visible = True
'    ElseIf mint���� = type_���� Then
    '����ҽ��֧�ֺ˶�
    If mint���� = type_���� Then
        mnuEditSplit3.Visible = True
        mnuEditVerify_Account.Visible = True
        mnuEditVerify_ZYPrice.Visible = True
        mnuEditVerify_UpDetail.Visible = True
    '��������ҽ�������� 204-03-31
    ElseIf mint���� = TYPE_���������� Then
        mnuEditVerify_Price.Visible = True
    ElseIf mint���� = TYPE_�����山 Then
        mnuEditSplit3.Visible = True
        mnuEditVerify_Account.Visible = True
        mnuEditVerify_Hospital.Visible = True
        mnuEditVerify_Price.Caption = "�˶Խ��������Ϣ(&J)"
        mnuEditVerify_Price.Visible = True
        mnuEditVerify_ZYPrice.Caption = "�˶Խ�������Ϣ(&H)"
        mnuEditVerify_ZYPrice.Visible = True
        mnuEditVerify_Detail.Visible = True
        mnuEditMend.Visible = True
         
        mnuEditMend.Caption = "ȡ����ǰ������¼(&Q)"
    End If
End Sub

Private Function GetInsureInit(ByVal intinsure As Integer) As Boolean
'���ܣ���ȡ�������Ƿ����ҽ����ʼ��
    Dim blnCanUse As Boolean
    Dim varCanUse As Variant
    
    On Error Resume Next
    varCanUse = mcol����("K" & intinsure)
    
    If Err <> 0 Then
        '��δ������ҽ���Ƿ����
        blnCanUse = gclsInsure.InitInsure(gcnOracle, intinsure)
        '������뼯����
        mcol����.Add blnCanUse, "K" & intinsure
        GetInsureInit = blnCanUse
        Exit Function
    End If
    
    GetInsureInit = varCanUse
End Function

Private Sub GetAccountInfo()
    Dim lngRow As Long
    Dim strTemp As String
    '�Ա����ʻ����ж���Ĵ���
    '�������ҽ����������ԭ��������֤���б�����Ǽ����ı��룬���û���Ҫ�������������ƣ�����������Ϣ����ǰ�÷������ϣ�������ʱ���ܷ����仯��
    
    Select Case mint����
    Case TYPE_������
        '���ȶ���������������
        If mcnYB.State <> 1 Then
            gstrSQL = "Select ������,����ֵ From ���ղ��� Where ����=" & TYPE_������
            Call OpenRecordset(mrs����, Me.Caption)
            Do Until mrs����.EOF
                strTemp = IIf(IsNull(mrs����("����ֵ")), "", mrs����("����ֵ"))
                Select Case mrs����("������")
                    Case "ҽ��������"
                        strServer = strTemp
                    Case "ҽ���û���"
                        strUser = strTemp
                    Case "ҽ���û�����"
                        strPass = strTemp
                End Select
                mrs����.MoveNext
            Loop
            If OraDataOpen(mcnYB, strServer, strUser, strPass) = False Then
                MousePointer = vbDefault
                Exit Sub
            End If
            
            If mrs����.State = adStateOpen Then mrs����.Close
            mrs����.Open "select BZBM ����,BZMC ����,ZJM ����  from BZML Order by BZBM", mcnYB, adOpenStatic, adLockReadOnly
            If mrs����.EOF = True Then
                MousePointer = vbDefault
                MsgBox "δ��ҽ��ǰ�÷������ж�����ز��֡�", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        '�޸ĵ�ǰ��ʾ�����б����ʻ��Ĳ�����ʾ
        With msh�ʻ�_S
            lngRow = .TopRow
            Do While .RowIsVisible(lngRow)
                If Trim(.TextMatrix(lngRow, col����֤��)) <> "" And Trim(.TextMatrix(lngRow, col����)) = "" Then
                    mrs����.MoveFirst
                    mrs����.Find "����='" & .TextMatrix(lngRow, col����֤��) & "'"
                    If Not mrs����.EOF Then
                        .TextMatrix(lngRow, col����) = mrs����!����
                    End If
                End If
                lngRow = lngRow + 1
                If lngRow > .Rows - 1 Then Exit Do
            Loop
        End With
    Case Else
    End Select
End Sub
