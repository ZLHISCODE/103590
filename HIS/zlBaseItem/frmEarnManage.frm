VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frmEarnManage 
   Caption         =   "������Ŀ����"
   ClientHeight    =   4980
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   6645
   Icon            =   "frmEarnManage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4980
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin MSComctlLib.ImageList ils32 
      Left            =   2610
      Top             =   1140
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEarnManage.frx":0442
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEarnManage.frx":089A
            Key             =   "ItemNo"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   2730
      Top             =   1800
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
            Picture         =   "frmEarnManage.frx":0CEE
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEarnManage.frx":1146
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEarnManage.frx":159E
            Key             =   "ItemNo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEarnManage.frx":19F2
            Key             =   "Write"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3225
      Left            =   3120
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3225
      ScaleWidth      =   45
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1590
      Width           =   45
   End
   Begin MSComctlLib.ListView lvwMain 
      Height          =   2235
      Left            =   3120
      TabIndex        =   1
      Top             =   1380
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   3942
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils32"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView tvwMain_S 
      Height          =   3345
      Left            =   240
      TabIndex        =   0
      Top             =   990
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   5900
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList Ilsrw 
      Left            =   5160
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEarnManage.frx":1E4A
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEarnManage.frx":206A
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEarnManage.frx":228A
            Key             =   "Parent"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEarnManage.frx":24A6
            Key             =   "Child"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEarnManage.frx":26C2
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEarnManage.frx":28E2
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEarnManage.frx":2B02
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEarnManage.frx":2D22
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEarnManage.frx":2F42
            Key             =   "View"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEarnManage.frx":3162
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEarnManage.frx":3382
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Ilscolor 
      Left            =   4320
      Top             =   300
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEarnManage.frx":35A2
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEarnManage.frx":37C2
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEarnManage.frx":39E2
            Key             =   "Parent"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEarnManage.frx":3BFE
            Key             =   "Child"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEarnManage.frx":3E1A
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEarnManage.frx":403A
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEarnManage.frx":425A
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEarnManage.frx":447A
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEarnManage.frx":469A
            Key             =   "View"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEarnManage.frx":48BA
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEarnManage.frx":4ADA
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   6645
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "Toolbar1"
      MinHeight1      =   720
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   720
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "Ilsrw"
         HotImageList    =   "Ilscolor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   15
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Preview"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Parent"
               Object.ToolTipText     =   "���ӷ���"
               Object.Tag             =   "����"
               ImageKey        =   "Parent"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��Ŀ"
               Key             =   "Child"
               Object.ToolTipText     =   "������Ŀ"
               Object.Tag             =   "��Ŀ"
               ImageKey        =   "Child"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "Modify"
               Object.ToolTipText     =   "�޸�"
               Object.Tag             =   "�޸�"
               ImageKey        =   "Modify"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "Delete"
               Object.ToolTipText     =   "ɾ��"
               Object.Tag             =   "ɾ��"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split1"
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Start"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "Start"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ͣ��"
               Key             =   "Stop"
               Object.ToolTipText     =   "ͣ��"
               Object.Tag             =   "ͣ��"
               ImageKey        =   "Stop"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "sdf"
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�鿴"
               Key             =   "View"
               Object.ToolTipText     =   "�鿴��ʽ"
               Object.Tag             =   "�鿴"
               ImageKey        =   "View"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "  ��ͼ��"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "  Сͼ��"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "  �б�"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "  ��ϸ����"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Quit"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   4620
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   635
      SimpleText      =   $"frmEarnManage.frx":4CFA
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmEarnManage.frx":4D41
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6641
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
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFileset 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFilepre 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnusplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileexit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditAddParent 
         Caption         =   "���ӷ���(&P)"
      End
      Begin VB.Menu mnuEditAddChild 
         Caption         =   "������Ŀ(&C)"
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "�޸�(&M)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "ɾ��(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditStart 
         Caption         =   "����(&S)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEditStop 
         Caption         =   "ͣ��(&T)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEditSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditExpand 
         Caption         =   "�ӳ��¼�����(&X)"
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
         Begin VB.Menu mnuviewspilt1 
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
      Begin VB.Menu mnuviewsplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "��ͼ��(&G)"
         Index           =   0
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "Сͼ��(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "�б�(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "��ϸ����(&D)"
         Checked         =   -1  'True
         Index           =   3
      End
      Begin VB.Menu mnuViewSplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewSelect 
         Caption         =   "ѡ����(&C)"
      End
      Begin VB.Menu mnuViewSplit4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewShowAll 
         Caption         =   "��ʾ�����¼�(&H)"
      End
      Begin VB.Menu mnuViewShowStop 
         Caption         =   "��ʾͣ����Ŀ(&P)"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "ˢ��(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTopic 
         Caption         =   "��������(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "Web�ϵ�����"
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
      Begin VB.Menu mnuHelpSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
   Begin VB.Menu mnuShort1 
      Caption         =   "��ݲ˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuShortMenu1 
         Caption         =   "���ӷ���(&P)"
         Index           =   1
      End
      Begin VB.Menu mnuShortMenu1 
         Caption         =   "�޸�(&M)"
         Index           =   2
      End
      Begin VB.Menu mnuShortMenu1 
         Caption         =   "ɾ��(&D)"
         Index           =   3
      End
   End
   Begin VB.Menu mnuShort2 
      Caption         =   "��ݲ˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "������Ŀ(&C)"
         Index           =   1
      End
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "�޸�(&M)"
         Index           =   2
      End
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "ɾ��(&D)"
         Index           =   3
      End
      Begin VB.Menu mnuShortsplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "��ͼ��(&G)"
         Index           =   0
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "Сͼ��(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "�б�(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "��ϸ����(&D)"
         Checked         =   -1  'True
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmEarnManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sngStartX As Single    '�ƶ�ǰ����λ��
Dim mblnItem As Boolean   'Ϊ���ʾ������ListViewĳһ����
Dim mintColumn As Integer
Dim mblnLoad As Boolean
Dim mstrKey As String
Private mstrLvw As String  '��ͷ���
Private mlngMode As Long
Private mstrPrivs As String                              'Ȩ�޴�
Private Sub Form_Activate()
    If mblnLoad = True Then
        Call Ȩ�޿���
        Call Form_Resize 'Ϊ��ʹCoolBar����Ӧ�߶�
        FillTree
    End If
    mblnLoad = False
End Sub

Private Sub Form_Load()
    Dim blnҩ�� As Boolean
    
    mblnLoad = True
    mlngMode = glngModul
    mstrPrivs = gstrPrivs
    
    '���������ɾ����ListView�������
    
    blnҩ�� = (glngSys \ 100 = 8)
    If blnҩ�� = True Then
        'ȥ��������Ŀ
        mstrLvw = "����,2000.126,0,1;����,1200,0,2;����,800,0,0;����,500,0,0;�վݷ�Ŀ,1440,0,0;����ʱ��,1100,0,0;����ʱ��,1100,0,0;��������,2000,0,0"
    Else
        mstrLvw = "����,2000.126,0,1;����,1200,0,2;����,800,0,0;����,500,0,0;�վݷ�Ŀ,1440,0,0;������Ŀ,1440,0,0;����ʱ��,1100,0,0;����ʱ��,1100,0,0;��������,2000,0,0"
    End If
    
    lvwMain.Tag = "�ɱ仯��"
    '���ListView�Ļ�δ�����ã������һ��ʹ�ã��Ǿ͵���ȱʡ�ĳ�ʼ��
    If lvwMain.ColumnHeaders.Count = 0 Then
        zlControl.LvwSelectColumns lvwMain, mstrLvw, True
    End If
    RestoreWinState Me, App.ProductName
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    
    mnuViewShowAll.Checked = (Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "��ʾ����", 0)) = 1)
    mnuViewShowStop.Checked = (Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "��ʾͣ��", 0)) = 1)
    '����LvwMain��ʾ���ö�Ӧ�˵�
     mnuViewIcon_Click lvwMain.View
End Sub

Private Sub Form_Resize()
    Dim sngTop As Single, sngBottom As Single
    sngTop = IIF(CoolBar1.Visible, CoolBar1.Top + CoolBar1.Height, 0)
    sngBottom = Me.ScaleHeight - IIF(stbThis.Visible, stbThis.Height, 0)
    
    tvwMain_S.Top = sngTop
    tvwMain_S.Height = IIF(sngBottom - tvwMain_S.Top > 0, sngBottom - tvwMain_S.Top, 0)
    tvwMain_S.Left = 0
    
    picSplit.Top = sngTop
    picSplit.Height = IIF(sngBottom - picSplit.Top > 0, sngBottom - picSplit.Top, 0)
    picSplit.Left = tvwMain_S.Left + tvwMain_S.Width
    
    lvwMain.Left = picSplit.Left + picSplit.Width
    lvwMain.Top = sngTop
    lvwMain.Height = IIF(sngBottom - lvwMain.Top > 0, sngBottom - lvwMain.Top, 0)
    If Me.ScaleWidth - lvwMain.Left > 0 Then lvwMain.Width = Me.ScaleWidth - lvwMain.Left
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrKey = ""
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "��ʾ����", IIF(mnuViewShowAll.Checked, 1, 0)
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "��ʾͣ��", IIF(mnuViewShowStop.Checked, 1, 0)
    SaveWinState Me, App.ProductName
End Sub

Private Sub lvwMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then '���Ǹղ�����
        lvwMain.SortOrder = IIF(lvwMain.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwMain.SortKey = mintColumn
        lvwMain.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwMain_DblClick()
    If mblnItem = True And mnuEditModify.Enabled And mnuEditModify.Visible Then mnuEditModify_Click
End Sub

Private Sub lvwMain_GotFocus()
    With lvwMain
        stbThis.Panels(2).Text = "��Ŀ�б��й���ʾ��" & .ListItems.Count & "��������Ŀ��"
    End With
    Call SetMenu
End Sub

Private Sub lvwMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mblnItem = True
    SetMenu
    stbThis.Panels(2).Text = "��Ŀ�б��й���ʾ��" & lvwMain.ListItems.Count & "��������Ŀ��"
End Sub

Private Sub lvwMain_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If mnuEditModify.Enabled And mnuEditModify.Visible Then mnuEditModify_Click
    End If
End Sub

Private Sub lvwMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnItem = False
End Sub

Private Sub lvwMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    If Button = 2 Then
        mnuShortMenu2(2).Enabled = mnuEditModify.Enabled
        mnuShortMenu2(3).Enabled = mnuEditDelete.Enabled
        For i = 0 To 3
            mnuShortIcon(i).Checked = mnuViewIcon(i).Checked
        Next
        PopupMenu mnuShort2, vbPopupMenuRightButton
    End If
End Sub

Private Sub mnuEditAddChild_Click()
    Dim str���� As String
    Dim str���� As String
    Dim i As Integer
    Dim blnReturn As Boolean
    
    With tvwMain_S.SelectedItem
        If .Key = "Root" Then
           blnReturn = frmEarnSet.�༭��Ŀ("��", "", "", , True)
        Else
            i = InStr(.Text, "��")
            str���� = Mid(.Text, 2, i - 2)
            str���� = Mid(.Text, i + 1)
            blnReturn = frmEarnSet.�༭��Ŀ(str����, Mid(.Key, 2), str����, , True)
        End If
        If blnReturn = True Then tvwMain_S_NodeClick tvwMain_S.SelectedItem
    End With
End Sub

Private Sub mnuEditAddParent_Click()
    Dim str���� As String
    Dim str���� As String
    Dim i As Integer
    Dim strKey As String
    Dim blnReturn As Boolean
    
    With tvwMain_S.SelectedItem
        strKey = .Key
        If .Key = "Root" Then
           blnReturn = frmEarnSet.�༭��Ŀ("��", "", "", , False)
        Else
            i = InStr(.Text, "��")
            str���� = Mid(.Text, 2, i - 2)
            str���� = Mid(.Text, i + 1)
           blnReturn = frmEarnSet.�༭��Ŀ(str����, Mid(.Key, 2), str����, , False)
        End If
    End With
    If blnReturn = True Then
        FillTree
    End If
End Sub

Private Sub mnuEditExpand_Click()
    Dim strTemp As String
    Dim str������ As String
    Dim str���� As String
    Dim intNew As Integer 'Ŀǰ���
    Dim intChild As Integer
    
    On Error GoTo ErrHandle
    With tvwMain_S.SelectedItem
        If .Key = "Root" Then
            str������ = ""
            intNew = GetDownCodeLength("", "������Ŀ")
            intChild = GetLocalCodeLength("", "������Ŀ")
        Else
            str������ = Mid(.Text, 2, InStr(.Text, "��") - 2)
            intNew = GetDownCodeLength(Mid(.Key, 2), "������Ŀ")
            intChild = GetLocalCodeLength(Mid(.Key, 2), "������Ŀ")
        End If
        If intNew = 0 Or intChild = 0 Then Exit Sub
        If intNew = 8 Then
            MsgBox "�����ټӳ����룬ĳһ���¼��Ѿ������˳��ȡ�", vbExclamation, gstrSysName
            Exit Sub
        End If
        
        intNew = frmCodingL.GetLength(intChild, 8 - (intNew - intChild), "", tvwMain_S.SelectedItem.Text)
        If intNew = 0 Then Exit Sub
        strTemp = str������ & String(intNew - intChild, "0")
        
        If .Key = "Root" Then
            gstrSQL = "zl_������Ŀ_EXPAND('" & strTemp & "'," & Len(str������) + 1 & ",0)"
        Else
            gstrSQL = "zl_������Ŀ_EXPAND('" & strTemp & "'," & Len(str������) + 1 & "," & Mid(.Key, 2) & ")"
        End If
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        FillTree
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditModify_Click()
    Dim str���� As String
    Dim str���� As String
    Dim i As Integer
    Dim strKey As String
    Dim blnReturn As Boolean
    
    With tvwMain_S.SelectedItem
        strKey = .Key
        If ActiveControl Is tvwMain_S Then
            If .Key = "Root" Then Exit Sub
            blnReturn = frmEarnSet.�༭��Ŀ("", "", "", Mid(.Key, 2))
        Else
            blnReturn = frmEarnSet.�༭��Ŀ("", "", "", Mid(lvwMain.SelectedItem.Key, 2))
        End If
    End With
    If blnReturn = True Then
        FillTree
    End If
End Sub

Private Sub mnuEditDelete_Click()
    On Error GoTo ErrHandle
    Dim strKey As String
    Dim intIndex As Long
    
    If ActiveControl Is tvwMain_S Then
        If MsgBox("ɾ������ͬʱҲ��ɾ����������Ŀ��" & vbCrLf & "��ȷ��Ҫɾ������Ϊ��" & tvwMain_S.SelectedItem.Text & "���ķ�����Ŀ��", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        gstrSQL = "zl_������Ŀ_delete(" & Mid(tvwMain_S.SelectedItem.Key, 2) & ")"
    Else
        If MsgBox("��ȷ��Ҫɾ������Ϊ��" & lvwMain.SelectedItem.Text & "����������Ŀ��", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        gstrSQL = "zl_������Ŀ_delete(" & Mid(lvwMain.SelectedItem.Key, 2) & ")"
    End If
    Me.MousePointer = 11
    
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    If ActiveControl Is tvwMain_S Then
        strKey = tvwMain_S.SelectedItem.Key
        If Not tvwMain_S.SelectedItem.Next Is Nothing Then
            tvwMain_S.SelectedItem.Next.Selected = True
            tvwMain_S_NodeClick tvwMain_S.SelectedItem
        Else
            tvwMain_S.SelectedItem.Parent.Selected = True
            tvwMain_S_NodeClick tvwMain_S.SelectedItem
        End If
        tvwMain_S.Nodes.Remove strKey
    Else
        With lvwMain
            intIndex = .SelectedItem.Index
            .ListItems.Remove .SelectedItem.Key
            If .ListItems.Count > 0 Then
                intIndex = IIF(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
                .ListItems(intIndex).Selected = True
                .ListItems(intIndex).EnsureVisible
            End If
            stbThis.Panels(2).Text = "��Ŀ�б��й���ʾ��" & lvwMain.ListItems.Count & "��������Ŀ��"
        End With
    End If
    SetMenu
    Me.MousePointer = 0
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Me.MousePointer = 0
End Sub

Private Sub mnuEditStart_Click()
    On Error GoTo ErrHandle
    
    gstrSQL = "zl_������Ŀ_reuse(" & Mid(lvwMain.SelectedItem.Key, 2) & ")"
    'ִ�����ù���
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '�ı�ͼ�����ɫ
    With lvwMain.SelectedItem
        .Icon = "Item"
        .SmallIcon = "Item"
        .ForeColor = RGB(0, 0, 0)
        
        Dim i As Integer
        For i = 1 To lvwMain.ColumnHeaders.Count
            If i < lvwMain.ColumnHeaders.Count Then
                .ListSubItems(i).ForeColor = RGB(0, 0, 0)
            End If
            '���³���ʱ��
            If lvwMain.ColumnHeaders(i).Text = "����ʱ��" Then
                .SubItems(i - 1) = "3000-01-01"
            End If
        Next
    End With
    '�ı�״̬���Ͳ˵�
    SetMenu
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditStop_Click()
    Dim strKey As String
    Dim rsTmp As ADODB.Recordset
    Dim strMsg As String
    Dim n As Integer
    
    On Error GoTo ErrHandle
    
    strKey = Mid(lvwMain.SelectedItem.Key, 2)
    
    '����Ƿ����շ���Ŀ��ʹ�õ�ǰ������Ŀ
    gstrSQL = "Select '[' || ���� || ']' || ���� As ��Ŀ���� From �շ���ĿĿ¼ " & _
          " Where ID In (Select �շ�ϸĿid From �շѼ�Ŀ " & _
          " Where ������Ŀid = [1] And ִ������ <= Sysdate And (��ֹ���� > Sysdate Or ��ֹ���� Is Null)) And " & _
          " (����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd') Or ����ʱ�� Is Null) Order By ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "����շ���Ŀ", Val(strKey))
    
    With rsTmp
        If Not .EOF Then
            strMsg = "��������Ŀ���������շ���Ŀ����ʹ�ã�"
            For n = 1 To .RecordCount
                strMsg = strMsg & vbCrLf & Space(4) & !��Ŀ����
                If n > 2 Then
                    strMsg = strMsg & vbCrLf & Space(4) & "��������" & .RecordCount - 3 & "����Ŀ������"
                    Exit For
                End If
                .MoveNext
            Next
        End If
    End With
    
    If strMsg <> "" Then
        If MsgBox(strMsg & vbCrLf & Space(4) & "�Ƿ�ͣ�ã�", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub
    End If
    
    gstrSQL = "zl_������Ŀ_stop(" & Val(strKey) & ")"
    'ִ�����ù���
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '�ı�ͼ�����ɫ
    If mnuViewShowStop.Checked = True Then 'Ҫ��ʾͣ�ò���
        With lvwMain.SelectedItem
            .Icon = "ItemNo"
            .SmallIcon = "ItemNo"
            .ForeColor = RGB(255, 0, 0)
            
            Dim i As Integer
            For i = 1 To lvwMain.ColumnHeaders.Count
                If i < lvwMain.ColumnHeaders.Count Then
                    .ListSubItems(i).ForeColor = RGB(255, 0, 0)
                End If
                '���³���ʱ��
                If lvwMain.ColumnHeaders(i).Text = "����ʱ��" Then
                    .SubItems(i - 1) = Format(Date, "yyyy-MM-dd")
                End If
            Next
        End With
        SetMenu
    Else '����ʾͣ�ò���
        With lvwMain
            .ListItems.Remove .SelectedItem.Key
            If .ListItems.Count > 0 Then
                .ListItems(1).Selected = True
                .ListItems(1).EnsureVisible
                lvwMain_ItemClick .SelectedItem
            Else
                Call lvwMain_GotFocus
            End If
        End With
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    'Ĭ�ϲ���������=����ID����Ŀ=��ĿID
    Dim lng����id As Long
    Dim lng��Ŀid As Long
    
    If Not tvwMain_S.SelectedItem Is Nothing Then
        If tvwMain_S.SelectedItem.Key <> "Root" Then
            lng����id = Mid(tvwMain_S.SelectedItem.Key, 2)
        End If
    End If
    
    If Not lvwMain.SelectedItem Is Nothing Then
        lng��Ŀid = Mid(lvwMain.SelectedItem.Key, 2)
    End If
    
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
        "����=" & IIF(lng����id = 0, "", lng����id), _
        "��Ŀ=" & IIF(lng��Ŀid = 0, "", lng��Ŀid))
End Sub

Private Sub mnuShortMenu1_Click(Index As Integer)
    Select Case Index
        Case 1
            mnuEditAddParent_Click
        Case 2
            mnuEditModify_Click
        Case 3
            mnuEditDelete_Click
    End Select
        
End Sub

Private Sub mnuShortMenu2_Click(Index As Integer)
    Select Case Index
        Case 1
            mnuEditAddChild_Click
        Case 2
            mnuEditModify_Click
        Case 3
            mnuEditDelete_Click
    End Select
        
End Sub

Private Sub mnuShortIcon_Click(Index As Integer)
    mnuViewIcon_Click Index
End Sub

Private Sub mnuViewIcon_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
        Toolbar1.Buttons("View").ButtonMenus(i + 1).Text = Replace(Toolbar1.Buttons("View").ButtonMenus(i + 1).Text, "��", "  ")
    Next
    mnuViewIcon(Index).Checked = True
    Toolbar1.Buttons("View").ButtonMenus(Index + 1).Text = Replace(Toolbar1.Buttons("View").ButtonMenus(Index + 1).Text, "  ", "��")
    lvwMain.View = Index
End Sub

Private Sub mnuViewRefresh_Click()
    Call FillTree
End Sub

Private Sub mnuViewSelect_Click()
    If zlControl.LvwSelectColumns(lvwMain, mstrLvw) = True Then
        '���б仯��Ҫ����ˢ��
        FillList tvwMain_S.SelectedItem.Key
    End If
End Sub

Private Sub mnuViewShowAll_Click()
    mnuViewShowAll.Checked = Not mnuViewShowAll.Checked
    FillList tvwMain_S.SelectedItem.Key
End Sub

Private Sub mnuViewShowStop_Click()
    mnuViewShowStop.Checked = Not mnuViewShowStop.Checked
    FillList tvwMain_S.SelectedItem.Key
End Sub

Private Sub picsplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        sngStartX = X
    End If
End Sub

Private Sub picsplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngTemp As Single
    If Button = 1 Then
        sngTemp = picSplit.Left + X - sngStartX
        If sngTemp > 0 And Me.ScaleWidth - (sngTemp + picSplit.Width) > 0 Then
            picSplit.Left = sngTemp
            tvwMain_S.Width = picSplit.Left - tvwMain_S.Left
            lvwMain.Left = picSplit.Left + picSplit.Width
            lvwMain.Width = Me.ScaleWidth - lvwMain.Left
        End If
    End If
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnufilepre_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub mnufileset_Click()
    zlPrintSet
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Parent"
            mnuEditAddParent_Click
        Case "Child"
            mnuEditAddChild_Click
        Case "Modify"
            mnuEditModify_Click
        Case "Delete"
            mnuEditDelete_Click
        Case "Stop"
            mnuEditStop_Click
        Case "Start"
            mnuEditStart_Click
        Case "Quit"
            mnuFileExit_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Preview"
            mnufilepre_Click
        Case "Help"
            mnuhelptopic_Click
        Case "View"
            mnuViewIcon(lvwMain.View).Checked = False
            If lvwMain.View = 3 Then
                mnuViewIcon(0).Checked = True
                lvwMain.View = 0
            Else
                mnuViewIcon(lvwMain.View + 1).Checked = True
                lvwMain.View = lvwMain.View + 1
            End If
    End Select
    
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    CoolBar1.Visible = mnuViewToolButton.Checked
    CoolBar1.Bands("only").MinHeight = Toolbar1.Height
    Form_Resize
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim buttTemp As Button
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For Each buttTemp In Toolbar1.Buttons
        If mnuViewToolText.Checked Then
            buttTemp.Caption = buttTemp.Tag
        Else
            buttTemp.Caption = ""
        End If
    Next
    CoolBar1.Bands("only").MinHeight = Toolbar1.Height
    Form_Resize
End Sub

Private Sub mnuhelptopic_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hwnd)
End Sub


Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim i As Integer
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
    Next
    mnuViewIcon(ButtonMenu.Index - 1).Checked = True
    lvwMain.View = ButtonMenu.Index - 1
End Sub

Private Sub Toolbar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool
    End If
End Sub

Private Sub tvwMain_S_GotFocus()
    stbThis.Panels(2).Text = "�����������" & lvwMain.ListItems.Count & "���¼���Ŀ"
    SetMenu
End Sub

Private Sub tvwMain_S_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If mnuShortMenu1(1).Visible = False Then Exit Sub
        mnuShortMenu1(2).Enabled = mnuEditModify.Enabled
        mnuShortMenu1(3).Enabled = mnuEditDelete.Enabled
        PopupMenu mnuShort1, vbPopupMenuRightButton
    End If
End Sub

Private Sub tvwMain_S_NodeClick(ByVal Node As MSComctlLib.Node)
    If mstrKey = Node.Key Then Exit Sub
    mstrKey = Node.Key
    
    FillList Node.Key
    mnuEditExpand.Enabled = lvwMain.ListItems.Count <> 0 Or Node.Children <> 0
    tvwMain_S_GotFocus
End Sub

Private Sub subPrint(bytMode As Byte)
'����:���д�ӡ,Ԥ���������EXCEL
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As New zlPrintLvw
    objPrint.Title.Text = "������Ŀ"
    Set objPrint.Body.objData = lvwMain
    objPrint.BelowAppItems.Add "��ӡ�ˣ�" & gstrUserName
    objPrint.BelowAppItems.Add "��ӡʱ�䣺" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrViewLvw objPrint, 1
          Case 2
              zlPrintOrViewLvw objPrint, 2
          Case 3
              zlPrintOrViewLvw objPrint, 3
      End Select
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
End Sub

Private Sub FillTree()
'����:װ��������Ŀ�����з��ൽtvwMain_S
    Dim strTemp As String
    Dim strKey As String
    Dim rs������Ŀ As New ADODB.Recordset
    
    mstrKey = ""
    
    On Error GoTo ErrHandle
    rs������Ŀ.CursorLocation = adUseClient
    rs������Ŀ.CursorType = adOpenKeyset
    rs������Ŀ.LockType = adLockReadOnly
    If Not tvwMain_S.SelectedItem Is Nothing Then
        strKey = tvwMain_S.SelectedItem.Key
    End If
    
    gstrSQL = "select ID,�ϼ�ID,����,���� from ������Ŀ  " & _
        "where ĩ�� <> 1 start with �ϼ�ID is null connect by prior ID =�ϼ�ID"
    Set rs������Ŀ = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)

    tvwMain_S.Nodes.Clear
    tvwMain_S.Nodes.Add , , "Root", "����������Ŀ", "Root", "Root"
    tvwMain_S.Nodes("Root").Sorted = True
    Do Until rs������Ŀ.EOF
        
        If IsNull(rs������Ŀ("�ϼ�id")) Then
            tvwMain_S.Nodes.Add "Root", tvwChild, "C" & rs������Ŀ("id"), "��" & rs������Ŀ("����") & "��" & rs������Ŀ("����"), "Write", "Write"
        Else
            tvwMain_S.Nodes.Add "C" & rs������Ŀ("�ϼ�id"), tvwChild, "C" & rs������Ŀ("id"), "��" & rs������Ŀ("����") & "��" & rs������Ŀ("����"), "Write", "Write"
        End If
        tvwMain_S.Nodes("C" & rs������Ŀ("ID")).Sorted = True
        rs������Ŀ.MoveNext
    Loop
    
    Dim nod As Node
    On Error Resume Next
    Set nod = tvwMain_S.Nodes(strKey)
    If Err <> 0 Then
        Set nod = tvwMain_S.Nodes("Root")
        nod.Selected = True
        nod.Expanded = True
        tvwMain_S_NodeClick nod
    Else
        Err.Clear
        nod.Selected = True
        nod.Expanded = True
        nod.EnsureVisible
        tvwMain_S_NodeClick nod
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub FillList(ByVal str������ĿID As String)
'����:װ���Ӧ������ӷ������Ŀ��lvwMain
'����:str������ĿID ����ı�ʶ
    Dim rs������Ŀ As New ADODB.Recordset
    Dim lst As ListItem
    Dim strKey As String
    Dim strͣ�� As String
    
    If Not lvwMain.SelectedItem Is Nothing Then
        '����ԭ�м�ֵ
        strKey = lvwMain.SelectedItem.Key
    End If
    
    rs������Ŀ.CursorLocation = adUseClient
    rs������Ŀ.CursorType = adOpenKeyset
    rs������Ŀ.LockType = adLockReadOnly
    
    On Error GoTo ErrHandle
    If mnuViewShowStop.Checked = False Then
        strͣ�� = " (����ʱ�� is null or ����ʱ�� = to_date('3000-01-01','YYYY-MM-DD')) and  "
    End If
    If mnuViewShowAll.Checked = True Then
        gstrSQL = "select A.*,B.���� as �������� from " & _
            "(select ID,�ϼ�ID,����,����,����,����,�վݷ�Ŀ,������Ŀ,to_char(����ʱ��,'YYYY-MM-DD') as ����ʱ��,to_char(����ʱ��,'YYYY-MM-DD') as ����ʱ�� from ������Ŀ where " & _
            IIF(strͣ�� = "", "", strͣ��) & " ĩ��=1  connect by prior id=�ϼ�id start with  " & _
            IIF(str������ĿID = "Root", "�ϼ�ID is null ", "�ϼ�ID = [1]") & ") A,������Ŀ B where A.�ϼ�ID=B.ID(+)"
    Else
        gstrSQL = "select ID,�ϼ�ID,����,����,����,����,�վݷ�Ŀ,������Ŀ,to_char(����ʱ��,'YYYY-MM-DD') as ����ʱ��,to_char(����ʱ��,'YYYY-MM-DD') as ����ʱ��,'' as �������� from ������Ŀ where " & _
            IIF(strͣ�� = "", "", strͣ��) & " ĩ��=1 and " & IIF(str������ĿID = "Root", "�ϼ�ID is null ", "�ϼ�ID = [1]")
    End If
    Set rs������Ŀ = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Mid(str������ĿID, 2)))
        
    zlControl.FormLock lvwMain.hwnd
    lvwMain.ListItems.Clear
    Do Until rs������Ŀ.EOF
        If CDate(IIF(IsNull(rs������Ŀ("����ʱ��")), "3000-01-01", rs������Ŀ("����ʱ��"))) = CDate("3000-01-01") Then
            Set lst = lvwMain.ListItems.Add(, "C" & rs������Ŀ("ID"), rs������Ŀ("����"), "Item", "Item")
        Else
            Set lst = lvwMain.ListItems.Add(, "C" & rs������Ŀ("ID"), rs������Ŀ("����"), "ItemNo", "ItemNo")
            lst.ForeColor = RGB(255, 0, 0)
        End If
        
        Dim lngCol  As Long
        Dim varValue As Variant
        '����ListView�����������ݿ�ȡ��
        For lngCol = 2 To lvwMain.ColumnHeaders.Count
            varValue = rs������Ŀ(lvwMain.ColumnHeaders(lngCol).Text).Value
            If lvwMain.ColumnHeaders(lngCol).Text <> "����" Then
                lst.SubItems(lngCol - 1) = IIF(IsNull(varValue), "", varValue)
            Else
                lst.SubItems(lngCol - 1) = IIF(varValue = 1, "��", "")
            End If
            If lst.Icon = "ItemNo" Then lst.ListSubItems(lngCol - 1).ForeColor = RGB(255, 0, 0)
        Next
        rs������Ŀ.MoveNext
    Loop
    zlControl.FormLock 0
    
    If lvwMain.ListItems.Count > 0 Then
        Dim Item As ListItem
        On Error Resume Next
        Set Item = lvwMain.ListItems(strKey)
        If Err <> 0 Then
            Set Item = lvwMain.ListItems(1)
            Item.Selected = True
            Item.EnsureVisible
            lvwMain_ItemClick Item
        Else
            Err.Clear
            Item.Selected = True
            Item.EnsureVisible
            lvwMain_ItemClick Item
        End If
    Else
        Call SetMenu
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetMenu()
'����:�����޸ĺ�ɾ����ť����Чֵ
'����:blnEnabled ��Чֵ
'    Dim blnNew As Boolean
    Dim blnModify As Boolean
    Dim blnStart As Boolean
    Dim blnStop As Boolean
    
    If ActiveControl Is tvwMain_S Then
        blnStart = False
        blnStop = False
        blnModify = tvwMain_S.SelectedItem.Key <> "Root"
    Else
        If lvwMain.SelectedItem Is Nothing Or lvwMain.ListItems.Count = 0 Then
            blnStart = False
            blnStop = False
            blnModify = False
        Else
            blnStart = (lvwMain.SelectedItem.Icon = "ItemNo")
            blnStop = (lvwMain.SelectedItem.Icon = "Item")
            blnModify = (lvwMain.SelectedItem.Icon = "Item")
        End If
    End If
    '���帳ֵ
'    Toolbar1.Buttons("Parent").Enabled = blnNew
'    Toolbar1.Buttons("Child").Enabled = blnNew
'    mnuEditAddParent.Enabled = blnNew
'    mnuEditAddChild.Enabled = blnNew
    
    Toolbar1.Buttons("Modify").Enabled = blnModify
    Toolbar1.Buttons("Delete").Enabled = blnModify
    mnuEditDelete.Enabled = blnModify
    mnuEditModify.Enabled = blnModify
    
    Toolbar1.Buttons("Start").Enabled = blnStart
    Toolbar1.Buttons("Stop").Enabled = blnStop
    mnuEditStart.Enabled = blnStart
    mnuEditStop.Enabled = blnStop
    If lvwMain.ListItems.Count > 0 Then
        mnuEditExpand.Enabled = True
    Else
        mnuEditExpand.Enabled = False
    End If

    EnablePrint (lvwMain.ListItems.Count > 0)
End Sub

Private Sub Ȩ�޿���()
'����:�����е��û�Ȩ�޲���,��ʹһЩ�˵����ť���ɼ�
    If InStr(mstrPrivs, "��ɾ��") = 0 Then
        mnuEdit.Visible = False
        mnuEditModify.Visible = False
        mnuShortMenu1(1).Visible = False
        mnuShortMenu2(1).Visible = False
        mnuShortMenu2(2).Visible = False
        mnuShortMenu2(3).Visible = False
        mnuShortsplit1.Visible = -False
        Toolbar1.Buttons("Split").Visible = False
        Toolbar1.Buttons("Parent").Visible = False
        Toolbar1.Buttons("Child").Visible = False
        Toolbar1.Buttons("Modify").Visible = False
        Toolbar1.Buttons("Delete").Visible = False
        Toolbar1.Buttons("Split1").Visible = False
        Toolbar1.Buttons("Start").Visible = False
        Toolbar1.Buttons("Stop").Visible = False
    End If
End Sub

Private Sub EnablePrint(ByVal blnEnabled As Boolean)
'����:���ô�ӡ��Ԥ����ť����Чֵ
'����:blnEnabled ��Чֵ
    Toolbar1.Buttons("Print").Enabled = blnEnabled
    Toolbar1.Buttons("Preview").Enabled = blnEnabled
    mnuFilepre.Enabled = blnEnabled
    mnuFilePrint.Enabled = blnEnabled
    mnuFileExcel.Enabled = blnEnabled
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

