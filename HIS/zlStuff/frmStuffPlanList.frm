VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmStuffPlanList 
   Caption         =   "���ļƻ�����"
   ClientHeight    =   5895
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11970
   Icon            =   "frmStuffPlanList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5895
   ScaleWidth      =   11970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin TabDlg.SSTab TabShow 
      Height          =   345
      Left            =   1680
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   609
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "�깺����(&0)"
      TabPicture(0)   =   "frmStuffPlanList.frx":014A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "���깺����(&1)"
      TabPicture(1)   =   "frmStuffPlanList.frx":0166
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
   End
   Begin VB.PictureBox picSeparate_s 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   30
      MousePointer    =   7  'Size N S
      ScaleHeight     =   300
      ScaleWidth      =   4815
      TabIndex        =   6
      Top             =   2790
      Width           =   4815
   End
   Begin VB.CommandButton Cmd���� 
      Caption         =   "����(&V)"
      Height          =   350
      Left            =   5160
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1100
   End
   Begin ComCtl3.CoolBar cbrTool 
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   11775
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tlbTool"
      MinHeight1      =   720
      Width1          =   6210
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Caption2        =   "�ⷿ"
      Child2          =   "cboStock"
      MinHeight2      =   300
      Width2          =   4095
      NewRow2         =   0   'False
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   6915
         TabIndex        =   4
         Text            =   "cboStock"
         Top             =   240
         Width           =   4770
      End
      Begin MSComctlLib.Toolbar tlbTool 
         Height          =   720
         Left            =   165
         TabIndex        =   1
         Top             =   30
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsCold"
         HotImageList    =   "ilsHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   16
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "PrintView"
               Description     =   "Ԥ��"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Description     =   "��ӡ"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "PrintSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Add"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "Modify"
               Description     =   "�޸�"
               Object.ToolTipText     =   "�޸�"
               Object.Tag             =   "�޸�"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "Delete"
               Description     =   "ɾ��"
               Object.ToolTipText     =   "ɾ��"
               Object.Tag             =   "ɾ��"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "EditSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "���"
               Key             =   "Verify"
               Description     =   "���"
               Object.ToolTipText     =   "���"
               Object.Tag             =   "���"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ȡ��"
               Key             =   "Cancel"
               Object.ToolTipText     =   "ȡ�����"
               Object.Tag             =   "ȡ��"
               ImageKey        =   "Cancel"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "���"
               Key             =   "Clear"
               Description     =   "���"
               Object.ToolTipText     =   "���"
               Object.Tag             =   "���"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "VerifySeparate"
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Search"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ˢ��"
               Key             =   "Refresh"
               Description     =   "ˢ��"
               Object.ToolTipText     =   "ˢ��"
               Object.Tag             =   "ˢ��"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FindSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "��������"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Exit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   11
            EndProperty
         EndProperty
         MouseIcon       =   "frmStuffPlanList.frx":0182
         Begin VB.Timer LimitTime 
            Enabled         =   0   'False
            Interval        =   8000
            Left            =   6660
            Top             =   180
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   5532
      Width           =   11964
      _ExtentX        =   21114
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmStuffPlanList.frx":049C
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16034
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
   Begin MSComctlLib.ImageList ilsCold 
      Left            =   0
      Top             =   600
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
            Picture         =   "frmStuffPlanList.frx":0D30
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":0F50
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":1170
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":138C
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":15AC
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":17CC
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":19E8
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":1C04
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":1E1E
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":1F78
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":2194
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":23B4
            Key             =   "Cancel"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHot 
      Left            =   600
      Top             =   600
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
            Picture         =   "frmStuffPlanList.frx":250E
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":272E
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":294E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":2B6A
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":2D8A
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":2FAA
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":31C6
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":33E2
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":35FC
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":3756
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":3976
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":3B96
            Key             =   "Cancle"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
      Height          =   1815
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   3201
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   1965
      Left            =   0
      TabIndex        =   7
      Top             =   3360
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   3466
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   0
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFilePreView 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileBillPrint 
         Caption         =   "���ݴ�ӡ(&B)"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuFileBillPreview 
         Caption         =   "����Ԥ��(&L)"
      End
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFileLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileParameter 
         Caption         =   "��������(&R)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFileLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditAdd 
         Caption         =   "����(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "�޸�(&M)"
      End
      Begin VB.Menu mnuEditDel 
         Caption         =   "ɾ��(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditVerify 
         Caption         =   "���(&C)"
      End
      Begin VB.Menu mnuEditCancel 
         Caption         =   "ȡ��(&Q)"
      End
      Begin VB.Menu mnuEditClear 
         Caption         =   "���(&S)"
      End
      Begin VB.Menu mnuEditLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDisplay 
         Caption         =   "�鿴����(&W)"
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
         Begin VB.Menu mnuViewToolText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewSearch 
         Caption         =   "����(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewLine4 
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
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "������̳(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&M)..."
         End
      End
      Begin VB.Menu mnuHelpLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frmStuffPlanList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngMode As Long
Private mstrFind As String
Private mintPreCol As Integer           'ǰһ�ε���ͷ��������
Private mintsort As Integer             'ǰһ�ε���ͷ������
Private mblnBootUp As Boolean
Private mlastRow As Long                '�ϴε������

Private mdtStartDate As Date
Private mdtEndDate As Date
Private mdtVerifyStart As Date
Private mdtVerifyEnd As Date
Private mstrPrivs As String
Private mintUnit  As Integer                '��ʾ��λ:0-ɢװ��λ,1-��װ��λ
Private mintOldY  As Integer
Private mstrOthers() As String  '0-��¼״̬(�ƻ�����),1-��ʼ����,2-��������,3-����id,4-�Է�����id(��������id����Ʒ���(�ƻ���)),5-������,6-�����,7-��Ӧ��ID,8-������,9-��ʼ��������,10-������������,11-��ʼ��Ʊ��,12-������Ʊ��
Private mblnCostView As Boolean             '�鿴�ɱ��������Ϣ true-����鿴 false-������鿴
Private mblnProvider As Boolean             '�鿴�ϴι�Ӧ�������Ϣ true-����鿴 false-������鿴
Private mstrCaption As String           '�������
Private mintFindDay As Integer          '��ѯ������Χ
 
'---------------------------------------------------------------------------------------------------------
'������صĹ�������:2008-08-22 16:35:52
'���˺�:
Private mblnNoClick As Boolean
Private mstr�������� As String
Private mbln����Ա���� As Boolean

Private mstrTitle As String '����

'----------------------------------------------------------------------------------------------------------
'���˺�:����С��λ���ĸ�ʽ��
'�޸�:2007/03/06
Private mOraFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Public Sub ShowList(ByVal lngMode As Long, ByVal strTitle As String, ByVal frmMain As Variant)
    '--------------------------------------------------------------------------------------------------------------------------
    '����:��ʾָ��ģ������
    '����:lngMode-ģ���
    '     strTitle-����
    '     frmMain-������
    '����:
    '����:���˺�
    '����:2007/12/26
    '����:11282
    '--------------------------------------------------------------------------------------------------------------------------
    Dim strOthers(0 To 12) As String
    Dim i As Integer
    Dim intCol As Integer
    
    mstrCaption = strTitle
    mblnBootUp = False
    mlngMode = lngMode
    mstrTitle = strTitle
    mstrPrivs = gstrPrivs
    
    mintFindDay = Val(zlDatabase.GetPara("��ѯ����", glngSys, mlngMode, 1))
    mdtStartDate = Format(DateAdd("d", -mintFindDay, sys.Currentdate), "yyyy-MM-dd")
    mdtEndDate = Format(sys.Currentdate, "yyyy-MM-dd")
    
    mdtVerifyStart = "1901-01-01"
    mdtVerifyEnd = "1901-01-01"
    
    mstrFind = " AND A.������� is Null And A.�������� Between [2] And [3]"
    
    If Not CheckDepend Then Exit Sub            '���������Բ���
    mblnCostView = zlStr.IsHavePrivs(mstrPrivs, "�鿴�ɱ���")
    mblnProvider = zlStr.IsHavePrivs(mstrPrivs, "�鿴��Ӧ��")
    
    Me.Caption = strTitle
    SetVisable  '����Ȩ�����ò�ͬ����ʾ��Ŀ
        
    For i = 0 To 12
        strOthers(i) = ""
    Next
    '������������
    strOthers(9) = "1901-01-01"
    strOthers(10) = "1901-01-01"
    mintUnit = Val(zlDatabase.GetPara("���ĵ�λ", glngSys, mlngMode, "0"))
  
    '���˺�:����С����ʽ����
    With mOraFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���, True)
        .FM_��� = GetFmtString(mintUnit, g_���, True)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�, True)
        .FM_���� = GetFmtString(mintUnit, g_����, True)
    End With
    mstrOthers = strOthers
    mstrPrivs = gstrPrivs
    mlastRow = 0
    
    If mlngMode <> 1725 Then
        TabShow.Visible = False
    Else
        TabShow.Visible = True
    End If
    
    GetList (mstrFind)   '�г�����ͷ
    RestoreWinState Me, App.ProductName, mstrTitle
    '�ָ����Ի��������ú󣬻���Ҫ��Ȩ�޿��Ƶ��н�һ������
    With mshDetail
        For intCol = 1 To .Cols - 1
            If mlngMode = 1725 Or mlngMode = 1724 Then
                If InStr(1, .TextMatrix(0, intCol), "�ɱ���") > 0 Or InStr(1, .TextMatrix(0, intCol), "�ɱ����") > 0 Then
                    .ColWidth(intCol) = IIf(mblnCostView = True, 900, 0)
                End If
            End If
            If mlngMode = 1725 Then
                If InStr(1, .TextMatrix(0, intCol), "�ϴι�Ӧ��") > 0 Then
                    .ColWidth(intCol) = IIf(mblnProvider = True, 1000, 0)
                End If
            End If
        Next
    End With
    
    '2006-04-25:���˺�,ͳһ���ӱ�������ģ��Ĺ���
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
    If mlngMode = 1725 Then
        cbrTool.Bands(2).Caption = "����"
    End If
    mblnBootUp = True
    
    If IsObject(frmMain) Then
        Me.Show , frmMain
    Else
        OS.ShowChildWindow Me.hwnd, frmMain
    End If
    Me.ZOrder 0
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub


'�������������
Private Function CheckDepend() As Boolean
    
    Dim rsTemp As New Recordset
    Dim strStock As String
    
    On Error GoTo ErrHandle
    CheckDepend = False
    strStock = " And b.���� In('V','K','12','W')"
    
    If mlngMode = 1725 Then
        gstrSQL = "" & _
            "   SELECT DISTINCT a.id, a.����||'-'||a.���� as ����" & _
            "   FROM ���ű� a  " & _
            "   where (a.����ʱ�� is null or TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01') " & _
            "       And (a.վ��=[2] or a.վ�� is null) " & _
            IIf(InStr(1, mstrPrivs, ";���в���;") > 0, "", " and  id in (Select ����id from ������Ա where ��Աid =[1])") & _
            "   Order by ����"
            mstr�������� = ""
            mbln����Ա���� = Not zlStr.IsHavePrivs(mstrPrivs, "���в���")
    Else
        gstrSQL = "" & _
            "   SELECT DISTINCT a.id , a.����||'-'||a.���� as ���� " & _
            "   FROM ��������˵�� c, �������ʷ��� b, ���ű� a " & _
            "   Where c.�������� = b.���� And (a.վ��=[2] or a.վ�� is null) " & _
            "             " & strStock & _
            "           AND a.id = c.����id " & _
            "           AND (a.����ʱ�� is null or TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01')" & _
            IIf(InStr(1, mstrPrivs, ";���пⷿ;") > 0, "", " and a.id in (Select ����id from ������Ա where ��Աid =[1])") & _
            "   Order by ����"
            mstr�������� = "V,K,12,W"
            mbln����Ա���� = Not zlStr.IsHavePrivs(mstrPrivs, "���пⷿ")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrTitle, UserInfo.Id, gstrNodeNo)
    
    If rsTemp.EOF Then
        If mlngMode = 1725 Then
            MsgBox "û�л��ֲ�����ϵ���㲻�߱���ص�Ȩ��,��鿴���Ź������ϵͳ����Ա��Ȩ��", vbInformation, gstrSysName
        Else
            MsgBox "û�л������Ŀ����ʵĲ��Ż򲻾߱���ص�Ȩ��,��鿴���Ź������ϵͳ����Ա��Ȩ��", vbInformation, gstrSysName
        End If
        rsTemp.Close
        Exit Function
    End If
            
    With cboStock
        .Clear
        If mlngMode <> 1725 Then
            If InStr(1, mstrPrivs, ";���пⷿ;") > 0 Then
                .AddItem "ȫԺ"
                .ItemData(.NewIndex) = 0
            End If
        End If
        
        If InStr(1, mstrPrivs, ";���в���;") > 0 Then
            .AddItem "���в���"
            .ItemData(.NewIndex) = 0
        End If
        Do While Not rsTemp.EOF
            .AddItem rsTemp!����
            .ItemData(.NewIndex) = rsTemp!Id
            If rsTemp!Id = UserInfo.����ID Then
                .ListIndex = .NewIndex
            End If
            rsTemp.MoveNext
        Loop
        rsTemp.Close
        If .ListIndex < 0 Then .ListIndex = 0
    End With
    
    CheckDepend = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Initialize()
    Call InitCommonControls
End Sub

Private Sub cboStock_Click()
    If mblnNoClick Then Exit Sub
    If cboStock.ListIndex >= 0 Then cboStock.Tag = cboStock.ItemData(cboStock.ListIndex)
    If mblnBootUp Then mnuViewRefresh_Click
End Sub
Private Sub cboStock_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cboStock.ListCount = 0 Then Call zlControl.ControlSetFocus(mshList): Exit Sub
    
    If cboStock.ListIndex >= 0 Then
        If Val(cboStock.Tag) = cboStock.ItemData(cboStock.ListIndex) Then
            Call zlControl.ControlSetFocus(mshList, True)
            Exit Sub
        End If
    End If
    
    If Select����ѡ����(Me, cboStock, Trim(cboStock.Text), mstr��������, mbln����Ա����) = False Then
        Exit Sub
    End If
    If cboStock.ListIndex >= 0 Then
        cboStock.Tag = cboStock.ItemData(cboStock.ListIndex)
    End If
End Sub

Private Sub cboStock_LostFocus()
    Dim i As Long
    If cboStock.ListCount = 0 Then Exit Sub
    If cboStock.ListIndex < 0 Then
        For i = 0 To cboStock.ListCount - 1
            If Val(cboStock.Tag) = cboStock.ItemData(i) Then
                mblnNoClick = True
                cboStock.ListIndex = i: Exit For
            End If
        Next
    End If
    mblnNoClick = False
End Sub

Private Sub cbrTool_Resize()
    Form_Resize
End Sub

Private Sub GetList(ByVal strFind As String)
    Dim rsTemp As New Recordset
    
    On Error GoTo ErrHandle
    Call FS.ShowFlash("�����������ϼƻ���¼,���Ժ� ...", Me)
    DoEvents
    Screen.MousePointer = vbHourglass
    mshList.Redraw = False
    If mlngMode = 1725 Then
        If TabShow.Tab = 0 Then
        If cboStock.ItemData(cboStock.ListIndex) <> 0 Then 'ѡ�������пⷿʱ�Ͳ���Ҫ�ⷿID��
            strFind = strFind & " and nvl(a.����id,0) =[1] "
        End If
        
        gstrSQL = "" & _
            "   SELECT a.no,a.id,b.���� as ����, decode(a.�ƻ�����,1,'�¶ȼƻ�',2,'���ȼƻ�',3,'��ȼƻ�',4,'�ܶȼƻ�') as �ƻ�����," & _
            "           a.�ڼ�,a.������,to_char(a.��������,'yyyy-mm-dd HH24:MI:SS') as ��������, a.�����," & _
            "           to_char(a.�������,'yyyy-mm-dd HH24:MI:SS') as �������,a.����˵�� " & _
            "   From ���ϲɹ��ƻ� a,���ű� b  " & _
            "  Where a.����=1 and a.����id=b.id " & strFind & _
            " ORDER BY a.no desc "
        Else
            If cboStock.ItemData(cboStock.ListIndex) <> 0 Then 'ѡ�������пⷿʱ�Ͳ���Ҫ�ⷿID��
                strFind = strFind & " and nvl(a.�ⷿid,0) =[1] "
            End If
            
            gstrSQL = "" & _
            "   SELECT a.no,a.id,b.���� as ����, decode(a.�ƻ�����,1,'�¶ȼƻ�',2,'���ȼƻ�',3,'��ȼƻ�',4,'�ܶȼƻ�') as �ƻ�����," & _
            "           a.�ڼ�,a.������,to_char(a.��������,'yyyy-mm-dd HH24:MI:SS') as ��������, a.�����," & _
            "           to_char(a.�������,'yyyy-mm-dd HH24:MI:SS') as �������,a.����˵�� " & _
            "   From ���ϲɹ��ƻ� a,���ű� b  " & _
            "  Where a.����=1 and a.�ⷿid=b.id " & strFind & _
            " ORDER BY a.no desc "
        End If
    Else
        gstrSQL = "" & _
            "   SELECT no,id, decode(�ƻ�����,1,'�¶ȼƻ�',2,'���ȼƻ�',3,'��ȼƻ�',4,'�ܶȼƻ�') as �ƻ����� ," & _
            "           �ڼ�,decode(���Ʒ���,1,'����ͬ�����β��շ�',2,'�ٽ��ڼ�ƽ�����շ�',3,'���ϴ���������շ�',4, '���������������շ�', '�����깺���շ�') as ���Ʒ��� ," & _
            "           ������,to_char(��������,'yyyy-mm-dd HH24:MI:SS') as ��������, �����," & _
            "           to_char(�������,'yyyy-mm-dd HH24:MI:SS') as �������,����˵�� " & _
            "   From ���ϲɹ��ƻ� a " & _
            "  Where nvl(�ⷿid,0) =[1] and ����=0 " & strFind & _
            " ORDER BY a.no desc "
    End If
    
    'mstrOthers(0 To 12) As String ' 0-��¼״̬(�ƻ�����),1-��ʼ����,2-��������,3-����id,4-�Է�����id(��������id����Ʒ���(�ƻ���)),5-������,6-�����,7-��Ӧ��ID,8-������,9-��ʼ��������,10-������������,11-��ʼ��Ʊ��,12-������Ʊ��
    '������Χ:[1]-�ⷿid,[2]:��ʼ��������,[3]������������,[4]��ʼ�������,[5] �����������,[6]-��¼״̬,[7]��ʼ���ݺ�,[8]�������ݺ�,[9]����id,[10]�Է�����id,[11]������,[12]�����[13]-��Ӧ��ID,[14]-������,[15]-��ʼ��������,[16]-������������,[17]-��ʼ��Ʊ��,[18]-������Ʊ��
    
    '��ʼ��������
    mstrOthers(9) = IIf(Trim(mstrOthers(9)) = "", "1901-01-01", mstrOthers(9))
    mstrOthers(10) = IIf(Trim(mstrOthers(10)) = "", "1901-01-01", mstrOthers(10))
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, cboStock.ItemData(cboStock.ListIndex), _
        CDate(Format(mdtStartDate, "yyyy-mm-dd") & " 00:00:00"), CDate(Format(mdtEndDate, "yyyy-mm-dd") & " 23:59:59"), _
        CDate(Format(mdtVerifyStart, "yyyy-mm-dd") & " 00:00:00"), CDate(Format(mdtVerifyEnd, "yyyy-mm-dd") & " 23:59:59"), _
        Val(mstrOthers(0)), mstrOthers(1), mstrOthers(2), Val(mstrOthers(3)), _
        Val(mstrOthers(4)), mstrOthers(5), mstrOthers(6), _
        Val(mstrOthers(7)), mstrOthers(8), CDate(mstrOthers(9) & " 00:00:00"), CDate(mstrOthers(10) & " 23:59:59"), _
         mstrOthers(11), mstrOthers(12))
          
          
    Set mshList.Recordset = rsTemp
    With mshList
        If .Rows = 1 Then
            .Rows = .Rows + 100
            .Row = 1
            .Redraw = True
            .TopRow = 1
            .Rows = .Rows - 99
        End If
        .Row = 1
        .Col = 0
        .ColSel = .Cols - 1
    End With
    SetListColWidth
    mshList.Redraw = True
    Call FS.StopFlash
    Screen.MousePointer = vbDefault
    SetEnable
    stbThis.Panels(2).Text = "��ǰ����" & rsTemp.RecordCount & "�ŵ���"
    rsTemp.Close
    Call mshlist_EnterCell
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'��ͷ�п��ʼ
Private Sub SetListColWidth()
    Dim intCol As Integer
    
    With mshList
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        
        If mblnBootUp = False Then
            For intCol = 0 To .Cols - 1
                .ColWidth(intCol) = 1500
                .ColAlignmentFixed(intCol) = 4
            Next
        Else
            For intCol = 0 To .Cols - 1
                .ColAlignmentFixed(intCol) = 4
            Next
        End If
        
        .ColWidth(1) = 0
    End With
End Sub

'����Ȩ�����ò�ͬ����ʾ��Ŀ
Private Sub SetVisable()
    '�⹺�������Ȩ�ޣ��������á����������пⷿ���Ǽǡ��޸ġ�ɾ�������ա����

    If InStr(1, mstrPrivs, ";����;") = 0 Then
        mnuEditAdd.Visible = False
        tlbTool.Buttons("Add").Visible = False
    End If
    
    If InStr(1, mstrPrivs, ";�޸�;") = 0 Then
        mnuEditModify.Visible = False
        tlbTool.Buttons("Modify").Visible = False
    End If
    
    
    If InStr(1, mstrPrivs, ";ɾ��;") = 0 Then
        mnuEditDel.Visible = False
        tlbTool.Buttons("Delete").Visible = False
         '��û�����б༭Ȩ��ʱ���Ѳ˵��͹������ϵ���Ӧ�ķָ������Ρ�
        If mnuEditAdd.Visible = False And mnuEditModify.Visible = False Then
            mnuEditLine1.Visible = False
            tlbTool.Buttons("EditSeparate").Visible = False
        End If
    End If
    
    If InStr(1, mstrPrivs, ";���;") = 0 Then
        mnuEditVerify.Visible = False
        tlbTool.Buttons("Verify").Visible = False
    End If
    If InStr(1, mstrPrivs, ";ȡ��;") = 0 Then
        mnuEditCancel.Visible = False
        tlbTool.Buttons("Cancel").Visible = False
    End If
    
    If InStr(1, mstrPrivs, ";���ݴ�ӡ;") = 0 Then
        mnuFileBillPrint.Visible = False
        mnuFileBillPreview.Visible = False
    End If
    
    If InStr(1, mstrPrivs, ";���;") = 0 Then
        mnuEditClear.Visible = False
        tlbTool.Buttons("Clear").Visible = False
        If mnuEditVerify.Visible = False And mnuEditCancel.Visible = False Then
            mnuEditLine2.Visible = False
            tlbTool.Buttons("VerifySeparate").Visible = False
        End If
    End If
    
End Sub
Private Sub Cmd����_Click()
    Call mnuEditDisplay_Click
End Sub

Private Sub Form_Load()
    PrintRange "��ѯ��Χ:" & Format(mdtStartDate, "yyyy��MM��dd��") & "��" & Format(mdtEndDate, "yyyy��MM��dd��")
End Sub

Private Sub Form_Resize()
    '����λ������
    
    On Error Resume Next
    
    If Me.WindowState = 1 Then Exit Sub
    
    cbrTool.Bands(1).MinHeight = tlbTool.Height
    With cbrTool
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth - .Left
    End With
    
    With picSeparate_s
        .Height = 300
        .Left = 0
        .Width = cbrTool.Width
    End With
    
    If mlngMode = 1725 Then
        With TabShow
            .Top = IIf(cbrTool.Visible, cbrTool.Height, 0)
            .Left = 0
        End With
        
        With mshList
            .Top = TabShow.Top + TabShow.Height
            .Left = 0
            .Width = cbrTool.Width
            .Height = picSeparate_s.Top - .Top
        End With
    Else
        With mshList
            .Top = IIf(cbrTool.Visible, cbrTool.Height, 0)
            .Left = 0
            .Width = cbrTool.Width
            .Height = picSeparate_s.Top - .Top
        End With
    End If
    
    With Cmd����
        .Left = Me.ScaleWidth - .Width - 100
        .Top = mshList.Top + mshList.Height + 30
        .ZOrder
    End With
    
    With mshDetail
        .Top = picSeparate_s.Top + picSeparate_s.Height + 100
        .Left = 0
        .Height = ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
        .Width = cbrTool.Width
    End With
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, mstrTitle
End Sub


 
Private Sub mnuEditAdd_Click()
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    strNo = ""
    '����
    Select Case mlngMode
    Case 1725
        If cboStock.ItemData(cboStock.ListIndex) <> 0 Then
            frmStuffRequestPlanCard.ShowCard Me, strNo, 1, mstrPrivs, blnSuccess
        End If
    Case 1724
        frmStuffPlanCard.ShowCard Me, strNo, 1, blnSuccess
    End Select
    If blnSuccess = True Then
        mnuViewRefresh_Click
    End If
End Sub


Private Sub mnuEditCancel_Click()
    '����
    
    Dim strNo As String
    Dim blnSuccess As Boolean
    Dim lngBillId  As Long
    
    With mshList
        strNo = Trim(.TextMatrix(.Row, 0))
        lngBillId = .TextMatrix(.Row, 1)
        If strNo = "" Then Exit Sub
        If MsgBox("��ȷʵҪȡ�����ݺ�Ϊ��" & strNo & "����" & IIf(mlngMode = 1725, "�빺��", "�ɹ��ƻ���") & "��������Ϣ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Sub
    End With
    err = 0: On Error GoTo ErrHand:
    'zl_���ϼƻ�����_Cancel(ID)
    gstrSQL = "zl_���ϼƻ�����_Cancel(" & lngBillId & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    mnuViewRefresh_Click
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub mnuEditClear_Click()
    '���
    Dim lngBillId As Long
    Dim intRow As Integer
    Dim strTitle As String
    Dim intReturn As Integer
    Dim intRecord As Integer
     
    With mshList
        
        On Error GoTo ErrHandle
        intRow = .Row
        lngBillId = .TextMatrix(intRow, 1)
        intReturn = MsgBox("��ȷʵҪ������ݺ�Ϊ��" & .TextMatrix(.Row, 0) & "����" & IIf(mlngMode = 1725, "�빺��", "�ɹ��ƻ���") & "��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
        intRecord = .Rows - 1
        If intReturn = vbYes Then
            gstrSQL = "zl_���ϼƻ�����_DELETE('" & lngBillId & "')"
            If gstrSQL = "" Then Exit Sub
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            intRecord = intRecord - 1
            If .Rows > 2 Then
                .RemoveItem intRow
            ElseIf .Rows = 2 Then
                .Rows = 3
                .RemoveItem intRow
                SetEnable
            End If
            If intRow < .Rows - 1 Then
                .Row = intRow
            Else
                If .Rows = 2 Then
                    .Row = 1
                Else
                    .Row = intRow - 1
                End If
            End If
            .Col = 0
            .ColSel = .Cols - 1
        End If
    End With
    
    mlastRow = 0
    Call mshlist_EnterCell
    stbThis.Panels(2).Text = "��ǰ����" & intRecord & "�ŵ���"
    Exit Sub

ErrHandle:
    Exit Sub
End Sub

Private Sub mnuEditVerify_Click()
    '����
    
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    With mshList
        strNo = .TextMatrix(.Row, 0)
        Select Case mlngMode
        Case 1725
            frmStuffRequestPlanCard.ShowCard Me, strNo, 3, mstrPrivs, blnSuccess
        Case 1724
            frmStuffPlanCard.ShowCard Me, strNo, 3, blnSuccess
        End Select

    
    End With
    If blnSuccess = True Then
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditDel_Click()
    'ɾ��
    Dim lngBillId As Long
    Dim intRow As Integer
    Dim strTitle As String
    Dim intReturn As Integer
    Dim intRecord As Integer
     
    With mshList
        
        On Error GoTo ErrHandle
        intRow = .Row
        lngBillId = .TextMatrix(intRow, 1)
        
        intReturn = MsgBox("��ȷʵҪɾ�����ݺ�Ϊ��" & .TextMatrix(.Row, 0) & "����" & IIf(mlngMode = 1725, "�빺��", "�ɹ��ƻ���") & "��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
        intRecord = .Rows - 1
        If intReturn = vbYes Then
            gstrSQL = "zl_���ϼƻ�����_DELETE('" & lngBillId & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            intRecord = intRecord - 1
            If .Rows > 2 Then
                .RemoveItem intRow
            ElseIf .Rows = 2 Then
                .Rows = 3
                .RemoveItem intRow
                
                SetEnable
            End If
            If intRow < .Rows - 1 Then
                .Row = intRow
            Else
                If .Rows = 2 Then
                    .Row = 1
                Else
                    .Row = intRow - 1
                End If
            End If
            .Col = 0
            .ColSel = .Cols - 1
        End If
    End With
    mlastRow = 0
    Call mshlist_EnterCell
    stbThis.Panels(2).Text = "��ǰ����" & intRecord & "�ŵ���"
    Exit Sub
ErrHandle:
    Exit Sub
End Sub

Private Sub mnuEditDisplay_Click()
    '�鿴����
    
    Dim strNo As String
    With mshList
        strNo = .TextMatrix(.Row, 0)
        Select Case mlngMode
        Case 1725
            frmStuffRequestPlanCard.ShowCard Me, strNo, 4, mstrPrivs
        Case 1724
            frmStuffPlanCard.ShowCard Me, strNo, 4
        End Select
    End With
End Sub

Private Sub mnuEditModify_Click()
    '�޸�
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    blnSuccess = False
    With mshList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        strNo = .TextMatrix(.Row, 0)
        Select Case mlngMode
        Case 1725
            frmStuffRequestPlanCard.ShowCard Me, strNo, 2, mstrPrivs, blnSuccess
        Case 1724
            frmStuffPlanCard.ShowCard Me, strNo, 2, blnSuccess
        End Select
        If blnSuccess = True Then
            mnuViewRefresh_Click
        End If
    End With
End Sub

Private Sub mnuFileBillPreview_Click()
    With mshList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        If mstrCaption = "���ļƻ�����" Then
            ReportOpen gcnOracle, glngSys, "zl1_bill_1724", Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��λ=" & mintUnit, 1
        Else
            ReportOpen gcnOracle, glngSys, "zl1_bill_1725", Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��λ=" & mintUnit, 1
        End If
    End With
End Sub

Private Sub mnuFileBillPrint_Click()
    With mshList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        If mstrCaption = "���ļƻ�����" Then
            ReportOpen gcnOracle, glngSys, "zl1_bill_1724", Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��λ=" & mintUnit, 2
        Else
            ReportOpen gcnOracle, glngSys, "zl1_bill_1725", Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��λ=" & mintUnit, 2
        End If
    End With
End Sub

Private Sub mnuFileExcel_Click()
    '�����Excel
    If Me.ActiveControl Is mshList Then
        mshList.Redraw = flexRDNone
        subPrint 3
        mshList.Redraw = flexRDDirect
        mshList.Col = 0
        mshList.ColSel = mshList.Cols - 1
    ElseIf Me.ActiveControl Is mshDetail Then
        mshDetail.Redraw = flexRDNone
        subPrint 3
        mshDetail.Redraw = flexRDDirect
        mshDetail.Col = 0
        mshDetail.ColSel = mshDetail.Cols - 1
    End If
End Sub

Private Sub mnufileexit_Click()
    '�˳�
    Unload Me
End Sub

Private Sub mnuFileParameter_Click()
'��������
    Dim strReg As String
    frmParaset.���ò��� mlngMode, mstrPrivs, Me, mstrCaption
    strReg = Val(zlDatabase.GetPara("���ĵ�λ", glngSys, mlngMode, "0"))
    mintUnit = Val(strReg)
    mlastRow = 0
  
    '���˺�:����С����ʽ����
    With mOraFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���, True)
        .FM_��� = GetFmtString(mintUnit, g_���, True)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�, True)
        .FM_���� = GetFmtString(mintUnit, g_����, True)
    End With
    mintFindDay = Val(zlDatabase.GetPara("��ѯ����", glngSys, mlngMode, 1))
    mdtStartDate = Format(DateAdd("d", -mintFindDay, sys.Currentdate), "yyyy-MM-dd")
    mdtEndDate = Format(sys.Currentdate, "yyyy-MM-dd")
    
    GetList (mstrFind)  '�г�����ͷ
'    Call mshlist_EnterCell
End Sub

Private Sub mnuFilePreView_Click()
    '��ӡԤ��
    mshList.Redraw = False
    subPrint 2
    mshList.Redraw = True
    mshList.Col = 0
    mshList.ColSel = mshList.Cols - 1
    
End Sub

Private Sub mnuFilePrint_Click()
    '��ӡ
    mshList.Redraw = False
    subPrint 1
    mshList.Redraw = True
    mshList.Col = 0
    mshList.ColSel = mshList.Cols - 1
End Sub

Private Sub mnuFilePrintSet_Click()
    '��ӡ����
    zlPrintSet
End Sub

Private Sub mnuHelpAbout_Click()
    '����
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
    '��������
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name)
End Sub

Private Sub mnuHelpWebHome_Click()
    '������ҳ
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    '���ͷ���
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub mnuViewRefresh_Click()
    'ˢ��
    mlastRow = 0
    GetList mstrFind
End Sub

Private Sub mnuViewSearch_Click()
    '����
    Dim strFind As String
    Dim strOthers() As String
    strFind = FrmStuffPlanSearch.GetSearch(Me, mdtStartDate, mdtEndDate, mdtVerifyStart, mdtVerifyEnd, strOthers)
    
    If strFind <> "" Then
        mstrFind = strFind
        mstrOthers = strOthers
        mlastRow = 0
        GetList mstrFind
        If Format(mdtStartDate, "yyyy-mm-dd") = "1901-01-01" And Format(mdtVerifyStart, "yyyy-mm-dd") = "1901-01-01" Then
        ElseIf Format(mdtStartDate, "yyyy-mm-dd") <> "1901-01-01" And Format(mdtVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
            PrintRange "��ѯ��Χ:�������� " & Format(mdtStartDate, "yyyy��MM��dd��") & "��" & Format(mdtEndDate, "yyyy��MM��dd��") & "  ������� " & Format(mdtVerifyStart, "yyyy��MM��dd��") & "��" & Format(mdtVerifyEnd, "yyyy��MM��dd��")
        ElseIf Format(mdtStartDate, "yyyy-mm-dd") <> "1901-01-01" Then
            PrintRange "��ѯ��Χ:�������� " & Format(mdtStartDate, "yyyy��MM��dd��") & "��" & Format(mdtEndDate, "yyyy��MM��dd��")
        ElseIf Format(mdtVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
            PrintRange "��ѯ��Χ:������� " & Format(mdtVerifyStart, "yyyy��MM��dd��") & "��" & Format(mdtVerifyEnd, "yyyy��MM��dd��")
        End If
     End If
End Sub

Private Sub mnuViewStatus_Click()
    With mnuViewStatus
        .Checked = Not .Checked
        stbThis.Visible = .Checked
    End With
    
    Form_Resize
End Sub
Private Sub mnuReportItem_Click(Index As Integer)
    Dim strNo As String
    Dim lng�ⷿID As Long
    
    With mshList
        strNo = Trim(.TextMatrix(.Row, 0))
    End With
    
    If cboStock.ListIndex < 0 Then
        lng�ⷿID = 0
    Else
        lng�ⷿID = cboStock.ItemData(cboStock.ListIndex)
    End If
    
    If mlngMode = 1725 Then
        '2006-04-25:���˺�:�����Զ��屨������ģ��Ĺ���
        If Format(mdtStartDate, "yyyy-mm-dd") = "1990-01-01" Then
            Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, "NO=" & strNo, "����id=" & lng�ⷿID)
        Else
            Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, "NO=" & strNo, "����id=" & lng�ⷿID, "��ʼʱ��=" & Format(mdtStartDate, "yyyy-mm-dd"), "����ʱ��=" & Format(mdtEndDate, "yyyy-mm-dd"))
        End If
    Else
        '2006-04-25:���˺�:�����Զ��屨������ģ��Ĺ���
        If Format(mdtStartDate, "yyyy-mm-dd") = "1990-01-01" Then
            Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, "NO=" & strNo, "�ⷿ=" & lng�ⷿID)
        Else
            Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, "NO=" & strNo, "�ⷿ=" & lng�ⷿID, "��ʼʱ��=" & Format(mdtStartDate, "yyyy-mm-dd"), "����ʱ��=" & Format(mdtEndDate, "yyyy-mm-dd"))
        End If
    End If
End Sub
Private Sub mnuViewToolButton_Click()
    With mnuViewToolButton
        .Checked = Not .Checked
        cbrTool.Bands(1).Visible = .Checked
        mnuViewToolText.Enabled = .Checked
    End With
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim intCount As Integer      '����������
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    With tlbTool.Buttons
        If mnuViewToolText.Checked = False Then
            'ȡ�����е��ı���ǩ��ʾ
            For intCount = 1 To .count
                .Item(intCount).Caption = ""
            Next
        Else
            '�����е��ı���ǩ��ʾ��˵����Tag�зŵ��ı���ǩ
            For intCount = 1 To .count
                .Item(intCount).Caption = .Item(intCount).Tag
            Next
        End If
    End With
    
    cbrTool.Bands(1).MinHeight = tlbTool.Height
    
    Form_Resize
End Sub


Private Sub mshList_Click()
    With mshList
         If .Row < 1 Then Exit Sub
         If .MouseRow = 0 Then
            ListSort
            Exit Sub
         End If
    End With
End Sub

Private Sub mshlist_DblClick()
    If mnuEditModify.Visible = False Then Exit Sub
    If mnuEditModify.Enabled = False Then Exit Sub
    If mshList.MouseRow = 0 Then Exit Sub
    mnuEditModify_Click
End Sub

Private Sub mshlist_EnterCell()
    Dim rsTemp As New Recordset
    Dim IntBill As Integer                      '��������  �磺1���⹺��⣻2��
    Dim strUnit As String                       '��λ����:�����ﵥλ��סԺ��λ��
    Dim str��װϵ�� As String
    
    mlastRow = mshList.Row
    
    On Error GoTo ErrHandle
    If mshList.Row >= 1 And LTrim(mshList.TextMatrix(mshList.Row, 0)) <> "" Then
        mshList.Col = 0
        mshList.ColSel = mshList.Cols - 1

        mshDetail.Redraw = False
        Select Case mintUnit
            Case 0
                str��װϵ�� = "1"
            Case Else
                str��װϵ�� = "D.����ϵ��"
        End Select
        If mlngMode = 1725 Then
            gstrSQL = "" & _
                "   SELECT M.����,M.���� as ͨ������, M.���," & IIf(mintUnit = 0, "M.���㵥λ", "D.��װ��λ") & " as  ��λ," & _
                "           trim(to_char(b.�빺���� /" & str��װϵ�� & "," & mOraFMT.FM_���� & ")) �빺����," & _
                "           trim(to_char(b.�������� /" & str��װϵ�� & "," & mOraFMT.FM_���� & ")) ��������," & _
                "           trim(to_char(b.�������� /" & str��װϵ�� & "," & mOraFMT.FM_���� & ")) ��������," & _
                "           trim(to_char(b.�ƻ����� /" & str��װϵ�� & "," & mOraFMT.FM_���� & ")) ��������," & _
                "           trim(to_char(b.���� *" & str��װϵ�� & "," & mOraFMT.FM_�ɱ��� & ")) �ɱ���," & _
                "           trim(to_char(b.���," & mOraFMT.FM_��� & ")) �ɱ����, b.�ϴι�Ӧ��,b.�ϴ������� " & _
                "   FROM ���ϲɹ��ƻ� a, ���ϼƻ����� b,���ű� c,�������� D,�շ���ĿĿ¼ M" & _
                "   Where a.id = b.�ƻ�id " & _
                "           and nvl(a.�ⷿid,0)=c.id(+) " & _
                "           and b.����id=d.����id and b.����id=M.id " & _
                "           AND b.�ƻ�ID =[1]" & _
                "   Order by ���"
        Else
            gstrSQL = "" & _
                "   SELECT M.����,M.���� as ͨ������, M.���," & IIf(mintUnit = 0, "M.���㵥λ", "D.��װ��λ") & " as  ��λ," & _
                "           trim(to_char(b.ǰ������ /" & str��װϵ�� & "," & mOraFMT.FM_���� & ")) ǰ������," & _
                "           trim(to_char(b.�������� /" & str��װϵ�� & "," & mOraFMT.FM_���� & ")) ��������," & _
                "           trim(to_char(b.������� /" & str��װϵ�� & "," & mOraFMT.FM_���� & ")) �������," & _
                "           trim(to_char(b.�������� /" & str��װϵ�� & "," & mOraFMT.FM_���� & ")) ��������," & _
                "           trim(to_char(b.�������� /" & str��װϵ�� & "," & mOraFMT.FM_���� & ")) ��������," & _
                "           trim(to_char(b.�ƻ����� /" & str��װϵ�� & "," & mOraFMT.FM_���� & ")) �ƻ�����," & _
                "           trim(to_char(b.���� *" & str��װϵ�� & "," & mOraFMT.FM_�ɱ��� & ")) �ɱ���," & _
                "           trim(to_char(b.���," & mOraFMT.FM_��� & ")) �ɱ����, b.�ϴι�Ӧ�� ��Ӧ��,b.�ϴ������� " & _
                "   FROM ���ϲɹ��ƻ� a, ���ϼƻ����� b,���ű� c,�������� D,�շ���ĿĿ¼ M" & _
                "   Where a.id = b.�ƻ�id " & _
                "           and nvl(a.�ⷿid,0)=c.id(+) " & _
                "           and b.����id=d.����id and b.����id=M.id " & _
                "           AND b.�ƻ�ID =[1] " & _
                "   Order by ���"
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mshList.TextMatrix(mshList.Row, 1)))
        Set mshDetail.Recordset = rsTemp
        
        With mshDetail
            If .Rows = 1 Then
                .Rows = .Rows + 100
                .Row = 1
                .Redraw = True

                .TopRow = 1
                .Rows = .Rows - 99
            End If
            .Row = 1
            .Col = 0
            .ColSel = .Cols - 1
        End With

        mshDetail.Redraw = True
    Else
        If mlngMode = 1725 Then
            With mshDetail
                .Cols = 12
                .Rows = 2
                .Clear
                .TextMatrix(0, 0) = "����"
                .TextMatrix(0, 1) = "����"
                .TextMatrix(0, 2) = "���"
                .TextMatrix(0, 3) = "��λ"
                .TextMatrix(0, 4) = "��������"
                .TextMatrix(0, 5) = "��������"
                .TextMatrix(0, 6) = "�빺����"
                .TextMatrix(0, 7) = "��������"
                .TextMatrix(0, 8) = "�ɱ���"
                .TextMatrix(0, 9) = "�ɱ����"
                .TextMatrix(0, 10) = "�ϴι�Ӧ��"
                .TextMatrix(0, 11) = "�ϴ�������"
                .Row = 1
                .Col = 0
                .ColSel = .Cols - 1
            End With
        Else
            With mshDetail
                .Cols = 14
                .Rows = 2
                .Clear
                .TextMatrix(0, 0) = "����"
                .TextMatrix(0, 1) = "����"
                .TextMatrix(0, 2) = "���"
                .TextMatrix(0, 3) = "��λ"
                .TextMatrix(0, 4) = "ǰ������"
                .TextMatrix(0, 5) = "��������"
                .TextMatrix(0, 6) = "�������"
                .TextMatrix(0, 7) = "��������"
                .TextMatrix(0, 8) = "��������"
                .TextMatrix(0, 9) = "�ƻ�����"
                .TextMatrix(0, 10) = "�ɱ���"
                .TextMatrix(0, 11) = "�ɱ����"
                .TextMatrix(0, 12) = "��Ӧ��"
                .TextMatrix(0, 13) = "�ϴ�������"
    
                .Row = 1
                .Col = 0
                .ColSel = .Cols - 1
            End With
        End If
    End If
    SetDetailColWidth
    SetEnable
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetDetailColWidth()
    Dim intCol As Integer
    
    With mshDetail
        If mblnBootUp = False Then
            .ColWidth(0) = 0
            .ColWidth(1) = 2500
            For intCol = 2 To .Cols - 1
                .ColWidth(intCol) = 1000
            Next
        End If
        
        Select Case mlngMode
            Case 1725
                For intCol = 0 To .Cols - 1
                    .ColAlignment(intCol) = 1
                    If InStr(1, .TextMatrix(0, intCol), "����") > 0 Or InStr(1, .TextMatrix(0, intCol), "����") > 0 Then
                        .ColAlignment(intCol) = flexAlignRightCenter     '����������
                    End If
                    If InStr(1, .TextMatrix(0, intCol), "�ɱ���") > 0 Then
                        .ColAlignment(intCol) = flexAlignRightCenter     '����
                        .ColWidth(intCol) = IIf(mblnCostView = False, 0, 1000)
                    End If
                    If InStr(1, .TextMatrix(0, intCol), "�ɱ����") > 0 Then
                        .ColAlignment(intCol) = flexAlignRightCenter     '���
                        .ColWidth(intCol) = IIf(mblnCostView = False, 0, 1000)
                    End If
                    If InStr(1, .TextMatrix(0, intCol), "�ϴι�Ӧ��") > 0 Then
                        .ColAlignment(intCol) = flexAlignLeftCenter     '�ϴι�Ӧ��
                        .ColWidth(intCol) = IIf(mblnProvider = False, 0, 1000)
                    End If
                    .ColAlignmentFixed(intCol) = 4
                Next
            Case 1724
                For intCol = 0 To .Cols - 1
                    .ColAlignment(intCol) = 1
                    If InStr(1, .TextMatrix(0, intCol), "����") > 0 Or InStr(1, .TextMatrix(0, intCol), "����") > 0 Then
                        .ColAlignment(intCol) = flexAlignRightCenter     '����������
                    End If
                    If InStr(1, .TextMatrix(0, intCol), "�ɱ���") > 0 Then
                        .ColAlignment(intCol) = flexAlignRightCenter     '����
                        .ColWidth(intCol) = IIf(mblnCostView = False, 0, 1000)
                    End If
                    If InStr(1, .TextMatrix(0, intCol), "�ɱ����") > 0 Then
                        .ColAlignment(intCol) = flexAlignRightCenter     '���
                        .ColWidth(intCol) = IIf(mblnCostView = False, 0, 1000)
                    End If
                    .ColAlignmentFixed(intCol) = 4
                Next
        End Select
    End With
End Sub

Private Sub mshlist_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If mnuEditModify.Visible = False Then Exit Sub
        If mnuEditModify.Enabled = False Then Exit Sub
        mnuEditModify_Click
    End If
        
End Sub

Private Sub mshlist_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    If mnuEdit.Visible = False Then Exit Sub
    PopupMenu mnuEdit, 2
End Sub

Private Sub picSeparate_s_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
        mintOldY = Y
End Sub

Private Sub picSeparate_s_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '�ָ�������
    
    If Button <> 1 Then Exit Sub
    
    With picSeparate_s
        If .Top + Y < 2000 Then Exit Sub
        If .Top + Y > ScaleHeight - 2000 Then Exit Sub
        .Move .Left, .Top + Y - mintOldY
    End With
    
    With mshList
        .Height = picSeparate_s.Top - .Top
    End With
    
    With Cmd����
        .Top = mshList.Top + mshList.Height + 30
    End With
    
    With mshDetail
        .Top = picSeparate_s.Top + picSeparate_s.Height + 100
        .Height = ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
End Sub

Private Sub picSeparate_s_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
        mintOldY = 0
End Sub

Private Sub tabShow_Click(PreviousTab As Integer)
    Call GetList(mstrFind)    '�г�����ͷ
End Sub

Private Sub tlbTool_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "PrintView"
            mnuFilePreView_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Add"
            mnuEditAdd_Click
        Case "Modify"
            mnuEditModify_Click
        Case "Delete"
            mnuEditDel_Click
        Case "Verify"
            mnuEditVerify_Click
        Case "Cancel"
            mnuEditCancel_Click
        Case "Clear"
            mnuEditClear_Click
        Case "Search"
            mnuViewSearch_Click
        Case "Refresh"
            mnuViewRefresh_Click
        Case "Help"
            mnuHelpTitle_Click
        Case "Exit"
            mnufileexit_Click
    End Select
End Sub

'���ò˵��͹��߰�ť�Ŀ�������
Private Sub SetEnable()
    With mshList
        .ToolTipText = ""
        If .TextMatrix(.Row, 0) = "" Or .Row = 0 Then         'û�е�
            mnuFilePreView.Enabled = False
            mnuFilePrint.Enabled = False
            mnuFileBillPreview.Enabled = False
            mnuFileBillPrint.Enabled = False
            mnuFileExcel.Enabled = False
            tlbTool.Buttons("Print").Enabled = False
            tlbTool.Buttons("PrintView").Enabled = False
            
            If mnuEditModify.Visible = True Then
                mnuEditModify.Enabled = False
                tlbTool.Buttons("Modify").Enabled = False
            End If
            If mnuEditDel.Visible = True Then
                mnuEditDel.Enabled = False
                tlbTool.Buttons("Delete").Enabled = False
            End If
            If mnuEditVerify.Visible = True Then
                mnuEditVerify.Enabled = False
                tlbTool.Buttons("Verify").Enabled = False
            End If
            
            If mnuEditCancel.Visible = True Then
                mnuEditCancel.Enabled = False
                tlbTool.Buttons("Cancel").Enabled = False
            End If
            
            If mnuEditClear.Visible = True Then
                mnuEditClear.Enabled = False
                tlbTool.Buttons("Clear").Enabled = False
            End If
            
            If mnuEditDisplay.Visible = True Then
                mnuEditDisplay.Enabled = False
            End If
            Cmd����.Enabled = False
        Else
            Cmd����.Enabled = True
            mnuFilePreView.Enabled = True
            mnuFilePrint.Enabled = True
            mnuFileBillPreview.Enabled = True
            mnuFileBillPrint.Enabled = True
            mnuFileExcel.Enabled = True
            tlbTool.Buttons("Print").Enabled = True
            tlbTool.Buttons("PrintView").Enabled = True
            
            If .TextMatrix(.Row, .Cols - 3) = "" Then    'δ��˵�
                If mnuEditModify.Visible = True Then
                    mnuEditModify.Enabled = True
                    tlbTool.Buttons("Modify").Enabled = True
                End If
                If mnuEditDel.Visible = True Then
                    mnuEditDel.Enabled = True
                    tlbTool.Buttons("Delete").Enabled = True
                End If
                If mnuEditVerify.Visible = True Then
                    mnuEditVerify.Enabled = True
                    tlbTool.Buttons("Verify").Enabled = True
                End If
                    
                If mnuEditCancel.Visible = True Then
                    mnuEditCancel.Enabled = False
                    tlbTool.Buttons("Cancel").Enabled = False
                End If
                
                If mnuEditClear.Visible = True Then
                    mnuEditClear.Enabled = False
                    tlbTool.Buttons("Clear").Enabled = False
                End If
                
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
                End If
            Else    '��˵�
                If mnuEditModify.Visible = True Then
                    mnuEditModify.Enabled = False
                    tlbTool.Buttons("Modify").Enabled = False
                End If
                If mnuEditDel.Visible = True Then
                    mnuEditDel.Enabled = False
                    tlbTool.Buttons("Delete").Enabled = False
                End If
                If mnuEditVerify.Visible = True Then
                    mnuEditVerify.Enabled = False
                    tlbTool.Buttons("Verify").Enabled = False
                End If
                If mnuEditCancel.Visible = True Then
                    mnuEditCancel.Enabled = True
                    tlbTool.Buttons("Cancel").Enabled = True
                End If
                
                If mnuEditClear.Visible = True Then
                    mnuEditClear.Enabled = True
                    tlbTool.Buttons("Clear").Enabled = True
                End If
                
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
                End If
            End If
        End If
    End With
End Sub

Private Sub subPrint(bytMode As Byte)
    '����:���д�ӡ,Ԥ���������EXCEL
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Dim objPrint As Object
    Dim objRow As New zlTabAppRow
    Dim strRange As String
    
    
    If Format(mdtStartDate, "yyyy-mm-dd") = "1901-01-01" And Format(mdtVerifyStart, "yyyy-mm-dd") = "1901-01-01" Then
        strRange = "������� " & Format(mdtVerifyStart, "yyyy��MM��dd��") & "��" & Format(mdtVerifyEnd, "yyyy��MM��dd��")
    ElseIf Format(mdtVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
        strRange = "�������� " & Format(mdtStartDate, "yyyy��MM��dd��") & "��" & Format(mdtEndDate, "yyyy��MM��dd��") & "  ������� " & Format(mdtVerifyStart, "yyyy��MM��dd��") & "��" & Format(mdtVerifyEnd, "yyyy��MM��dd��")
    Else
        strRange = "�������� " & Format(mdtStartDate, "yyyy��MM��dd��") & "��" & Format(mdtEndDate, "yyyy��MM��dd��")
    End If
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    objPrint.Title.Text = Me.Caption
        
    objRow.Add "ʱ�䣺" & strRange
    objRow.Add "���ţ�" & cboStock.Text
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow

    objRow.Add "��ӡ��:" & UserInfo.�û���
    objRow.Add "��ӡ����:" & Format(sys.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    
    If Me.ActiveControl Is mshList Then
        Set objPrint.Body = mshList
    Else
        Set objPrint.Body = mshDetail
    End If
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub tlbTool_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
        Case "Payment"
        Case "Imprest"
    End Select
End Sub

Private Sub tlbTool_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool
    End If
End Sub

'�Ե���ͷ������
Private Sub ListSort()
    Dim intCol As Integer
    Dim intRow As Integer
    Dim intTemp As String

    With mshList
        If .Rows > 1 Then
'            .Redraw = False
'            intCol = .MouseCol
'            .Col = intCol
'            .ColSel = intCol
'            intTemp = .TextMatrix(.Row, 0)
'            If intCol = mintPreCol And mintsort = flexSortStringNoCaseDescending Then
'               .Sort = flexSortStringNoCaseAscending
'               mintsort = flexSortStringNoCaseAscending
'            Else
'               .Sort = flexSortStringNoCaseDescending
'               mintsort = flexSortStringNoCaseDescending
'            End If
'            mintPreCol = intCol
'            .Row = Grid.MshGrdFindRow(mshList, intTemp, 0)
'            If .RowPos(.Row) + .RowHeight(.Row) > .Height Then
'                .TopRow = .Row
'            Else
'                .TopRow = 1
'            End If
            .Col = 0
            .ColSel = .Cols - 1
            .Redraw = True
            .SetFocus
        Else
            .ColSel = 0
        End If
    End With
End Sub

Private Sub PrintRange(ByVal strRange As String)
    '����:��ӡʱ�䷶Χ
    picSeparate_s.Cls
    picSeparate_s.CurrentX = 50
    picSeparate_s.CurrentY = 100
    picSeparate_s.Print strRange
End Sub
Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

