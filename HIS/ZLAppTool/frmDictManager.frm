VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmDictManager 
   Caption         =   "�ֵ������"
   ClientHeight    =   5700
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   10680
   Icon            =   "frmDictManager.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5700
   ScaleWidth      =   10680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picSplit2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3225
      Left            =   6345
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3225
      ScaleWidth      =   45
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1155
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.PictureBox picTable 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   3630
      ScaleHeight     =   345
      ScaleWidth      =   2595
      TabIndex        =   7
      Top             =   780
      Width           =   2595
      Begin VB.Label lblTable 
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Left            =   90
         TabIndex        =   8
         Top             =   45
         Width           =   2220
      End
   End
   Begin zl9AppTool.zlOutLook outTable_S 
      Height          =   2265
      Left            =   480
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   3995
   End
   Begin ComCtl3.CoolBar clbOnly 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10680
      _ExtentX        =   18838
      _ExtentY        =   1376
      _CBWidth        =   10680
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tlbMain"
      MinWidth1       =   4995
      MinHeight1      =   720
      Width1          =   4995
      FixedBackground1=   0   'False
      Key1            =   "Comm"
      NewRow1         =   0   'False
      Caption2        =   "ϵͳ"
      Child2          =   "cmbSys"
      MinWidth2       =   915
      MinHeight2      =   300
      Width2          =   1455
      NewRow2         =   0   'False
      BandStyle2      =   1
      Caption3        =   "����"
      Child3          =   "txtSeek"
      MinWidth3       =   1200
      MinHeight3      =   375
      Width3          =   1200
      NewRow3         =   0   'False
      Begin VB.TextBox txtSeek 
         Height          =   375
         Left            =   7905
         TabIndex        =   11
         ToolTipText     =   "֧��˫��ƥ�䣬��ͬ�ı��ٴΰ��س��������һ��"
         Top             =   195
         Width           =   1200
      End
      Begin VB.ComboBox cmbSys 
         Height          =   300
         Left            =   9675
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin MSComctlLib.Toolbar tlbMain 
         Height          =   720
         Left            =   165
         TabIndex        =   3
         Top             =   30
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgToolsStard"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   13
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Preview"
               Object.ToolTipText     =   "��ӡԤ��"
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
               Key             =   "NewGroup"
               Object.ToolTipText     =   "������Ŀ����"
               Object.Tag             =   "����"
               ImageKey        =   "NewGroup"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "SplitGroup"
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "New"
               Object.ToolTipText     =   "������Ŀ��ϸ"
               Object.Tag             =   "����"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "Modify"
               Object.ToolTipText     =   "�޸���Ŀ��ϸ"
               Object.Tag             =   "�޸�"
               ImageKey        =   "Modify"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "Delete"
               Object.ToolTipText     =   "ɾ����Ŀ��ϸ"
               Object.Tag             =   "ɾ��"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split1"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�鿴"
               Key             =   "View"
               Object.ToolTipText     =   "�鿴��ʽ"
               Object.Tag             =   "�鿴"
               ImageKey        =   "View"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "��ͼ��"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Сͼ��"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "�б�"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "��ϸ����"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Object.ToolTipText     =   "��������"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Exit"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imgToolsHot 
      Left            =   1050
      Top             =   840
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
            Picture         =   "frmDictManager.frx":0442
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDictManager.frx":065C
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDictManager.frx":0876
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDictManager.frx":0CC8
            Key             =   "New"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDictManager.frx":0EE2
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDictManager.frx":10FC
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDictManager.frx":1316
            Key             =   "View"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDictManager.frx":1536
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDictManager.frx":1750
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgToolsStard 
      Left            =   360
      Top             =   840
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
            Picture         =   "frmDictManager.frx":196A
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDictManager.frx":1B84
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDictManager.frx":1D9E
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDictManager.frx":1FB8
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDictManager.frx":21D2
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDictManager.frx":23EC
            Key             =   "View"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDictManager.frx":260C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDictManager.frx":2826
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDictManager.frx":2A40
            Key             =   "NewGroup"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDictManager.frx":2E92
            Key             =   ""
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
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1590
      Width           =   45
   End
   Begin MSComctlLib.ListView lvwMain 
      Height          =   2235
      Left            =   3600
      TabIndex        =   0
      Top             =   1170
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
   Begin MSComctlLib.ImageList ils32 
      Left            =   2730
      Top             =   1695
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDictManager.frx":96F4
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDictManager.frx":9B4C
            Key             =   "Write"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDictManager.frx":9F9E
            Key             =   "Read"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   2820
      Top             =   2370
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDictManager.frx":A2B8
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDictManager.frx":A710
            Key             =   "Write"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDictManager.frx":AB62
            Key             =   "Read"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDictManager.frx":AE7C
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDictManager.frx":116DE
            Key             =   "Group"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDictManager.frx":17F40
            Key             =   "GroupOpen"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   5340
      Width           =   10680
      _ExtentX        =   18838
      _ExtentY        =   635
      SimpleText      =   $"frmDictManager.frx":1E7A2
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDictManager.frx":1E7E9
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13758
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
   Begin MSComctlLib.TreeView tvwMain 
      Height          =   1620
      Left            =   3315
      TabIndex        =   9
      Top             =   1935
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   2858
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFileset 
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
      Begin VB.Menu mnusplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileexit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditNew 
         Caption         =   "����(&A)"
         Shortcut        =   ^A
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
      Begin VB.Menu mnuEditNewGroup 
         Caption         =   "���ӷ���(&I)"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuEditModifyGroup 
         Caption         =   "�޸ķ���(&U)"
      End
      Begin VB.Menu mnuEditDeleteGroup 
         Caption         =   "ɾ������(&E)"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuEditSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDefault 
         Caption         =   "��Ϊȱʡ��(&F)"
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
         Begin VB.Menu mnuViewToolspilt1 
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
      Begin VB.Menu mnuViewSplit4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewReflash 
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
   Begin VB.Menu mnuShort 
      Caption         =   "��ݲ˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuShortMenu 
         Caption         =   "����(&A)"
         Index           =   1
      End
      Begin VB.Menu mnuShortMenu 
         Caption         =   "�޸�(&M)"
         Index           =   2
      End
      Begin VB.Menu mnuShortMenu 
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
   Begin VB.Menu mnuGroup 
      Caption         =   "����˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuGroupMenu 
         Caption         =   "���ӷ���(&I)"
         Index           =   1
      End
      Begin VB.Menu mnuGroupMenu 
         Caption         =   "�޸ķ���(&E)"
         Index           =   2
      End
      Begin VB.Menu mnuGroupMenu 
         Caption         =   "ɾ������(&E)"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmDictManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim mblnItem As Boolean  'Ϊ���ʾ������ListViewĳһ����
Dim mintListIndex As Integer 'cmbSys��ǰһ���б�����
Dim mintColumn As Integer    'ǰһ��ListView��ͷ���

Dim mblnFail As Boolean
Dim mcolSys As New Collection  '����������ϵͳ��������
Dim mstrOwner As String        '��ǰѡ��ϵͳ��������
Dim mstrTables As String
Dim mlngSys As Long              '�ӿڴ���ָ����ϵͳ��
 
Dim mLastNode As Node
Private mlngLastPos As Long
Public gstrSTOwner As String
Public gblnHaveRIS As Boolean
Public gblnMustRIS As Boolean '����Ӱ����Ϣϵͳ�ӿ�
Public gobjRIS As Object

Public Sub �ֵ����(Optional strTables As String, Optional ByVal lngSys As Long)
    'strTables ָ�����ֵ������ñ���ֵΪ������ʾ�����ֵ���������ʾָ�����ֵ��
    Dim rsSys As New ADODB.Recordset
    mstrTables = strTables
    mlngSys = lngSys

    If mcolSys.Count > 0 Then
        '�Ѿ�����˳�ʼ���������ǵڶ�����ʾ
        frmDictManager.Show , gfrmMain
        Exit Sub
    End If
    
    Load frmDictManager
    
    frmDictManager.clbOnly.Bands(2).Visible = strTables = ""
    frmDictManager.clbOnly.Bands(3).Visible = strTables = ""
    '��ɳ�ʼ��
    gstrSQL = "select A.���,A.����,A.������ " & _
               " from zlSystems A,zlBasecode B,all_tables C " & _
               " Where A.��� = B.ϵͳ And upper(B.����) = C.table_name  and A.������=C.OWNER " & _
               " group by A.���,A.����,A.������ " & _
               " Having Count(A.���) > 0" & _
               " Order by ��� "
    Call zlDatabase.OpenRecordset(rsSys, gstrSQL, Me.Caption)
    
    mblnFail = False
    If rsSys.EOF Then
        MsgBox "��û�п��Թ���������ֵ䡣", vbInformation, gstrSysName
        Unload frmDictManager
        Exit Sub
    End If
    Do While Not rsSys.EOF
        If rsSys("���") = 100 Then
            gstrSTOwner = rsSys("������") & ""
        End If
        cmbSys.AddItem rsSys("����") & "��" & rsSys("���") & "��"
        cmbSys.ItemData(Me.cmbSys.NewIndex) = rsSys("���")
        mcolSys.Add CStr(rsSys("������")), "C" & rsSys("���")
        rsSys.MoveNext
    Loop
    If gstrSTOwner <> "" Then
        gblnMustRIS = Val(zlDatabase.GetPara(255, 100, 0, "0")) = 1
        On Error Resume Next
        Set gobjRIS = CreateObject("zl9XWInterface.clsHISInner")
        Err.Clear: On Error GoTo 0
        gblnHaveRIS = Not gobjRIS Is Nothing
    Else
        gblnMustRIS = False
        gblnHaveRIS = False
    End If
    mintListIndex = -1
    If cmbSys.ListCount > 0 Then cmbSys.ListIndex = 0
    If cmbSys.ListCount = 1 Then cmbSys.Enabled = False
    
    If mblnFail = True Then
        Unload frmDictManager
        Exit Sub
    End If
    
    frmDictManager.Show , gfrmMain
End Sub

Private Sub clbOnly_HeightChanged(ByVal NewHeight As Single)
    Form_Resize
End Sub


Private Sub cmbSys_Click()
    If mintListIndex = cmbSys.ListIndex Then Exit Sub
    
    mstrOwner = mcolSys("C" & cmbSys.ItemData(cmbSys.ListIndex))
    If FillTable = False And mintListIndex >= 0 Then
        cmbSys.ListIndex = mintListIndex
        mstrOwner = mcolSys("C" & cmbSys.ItemData(cmbSys.ListIndex))
        Exit Sub
    End If
    
    mintListIndex = cmbSys.ListIndex
End Sub

Private Sub Form_Activate()
    Call SetMenu
End Sub

Private Sub Form_Load()
    Dim intView As Integer
    
    intView = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name, "OutlookView", 1)
    If intView <> 0 And intView <> 1 Then
        intView = 1
    End If
    outTable_S.View = intView
    RestoreWinState Me, App.ProductName
    
    Set outTable_S.ImageList = ils32
    Set outTable_S.SmallImageList = ils16
    
End Sub

Private Sub Form_Resize()
    Dim sngTop As Single, sngBottom As Single
    
    On Error Resume Next
    sngTop = IIf(clbOnly.Visible, clbOnly.Top + clbOnly.Height, 0)
    sngBottom = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
    
    outTable_S.Top = sngTop
    outTable_S.Height = IIf(sngBottom - outTable_S.Top > 0, sngBottom - outTable_S.Top, 0)
    outTable_S.Left = 0
    
    picSplit.Top = sngTop
    picSplit.Height = IIf(sngBottom - picSplit.Top > 0, sngBottom - picSplit.Top, 0)
    picSplit.Left = outTable_S.Left + outTable_S.Width
    
    picTable.Top = sngTop + 45
    picTable.Left = picSplit.Left + picSplit.Width
    If Me.ScaleWidth - picTable.Left > 0 Then picTable.Width = ScaleWidth - picTable.Left
    lblTable.Width = picTable.Width
    
    '-- 10152�޸�
    If tvwMain.Visible Then
        tvwMain.Left = picTable.Left
        tvwMain.Top = picTable.Top + picTable.Height + 45
        tvwMain.Height = IIf(sngBottom - tvwMain.Top > 0, sngBottom - tvwMain.Top, 0)
        
        picSplit2.Left = tvwMain.Left + tvwMain.Width
        picSplit2.Top = tvwMain.Top
        picSplit2.Height = tvwMain.Height
        
        lvwMain.Left = picSplit2.Left + picSplit2.Width
        lvwMain.Top = tvwMain.Top
        lvwMain.Width = picTable.Width - tvwMain.Width - picSplit2.Width - 45
        lvwMain.Height = tvwMain.Height
    Else
        lvwMain.Left = picTable.Left
        lvwMain.Top = picTable.Top + picTable.Height + 45
        lvwMain.Width = picTable.Width
        lvwMain.Height = IIf(sngBottom - lvwMain.Top > 0, sngBottom - lvwMain.Top, 0)
    End If
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mcolSys = Nothing
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name, "OutlookView", outTable_S.View)
    SaveWinState Me, App.ProductName
End Sub

Private Sub lvwMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then '���Ǹղ�����
        lvwMain.SortOrder = IIf(lvwMain.SortOrder = lvwAscending, lvwDescending, lvwAscending)
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
    Dim i As Integer
    With lvwMain
        For i = 0 To 3
            mnuViewIcon(i).Checked = False
        Next
        mnuViewIcon(.View).Checked = True
    End With

End Sub

Private Sub lvwMain_ItemClick(ByVal item As MSComctlLib.ListItem)
    mblnItem = True
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
        mnuShortMenu(1).Enabled = mnuEditNew.Enabled
        mnuShortMenu(2).Enabled = mnuEditModify.Enabled
        mnuShortMenu(3).Enabled = mnuEditDelete.Enabled
        
        For i = 0 To 3
            mnuShortIcon(i).Checked = mnuViewIcon(i).Checked
        Next
        PopupMenu mnuShort, vbPopupMenuRightButton
    End If
End Sub

Private Sub mnuEditDefault_Click()
    
    On Error GoTo errHandle
    gstrSQL = "Update " & mstrOwner & "." & Mid(lblTable.Tag, 2) & _
        " Set ȱʡ��־=Decode(����,'" & Mid(lvwMain.SelectedItem.Key, 2) & "',1,0)"
    gstrSQL = "ZL_�ֵ����_execute('" & Replace(gstrSQL, "'", "''") & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    Call FillList
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub mnuEditDelete_Click()
    On Error GoTo errHandle
    Dim lngSystem As Long
    Dim intIndex As Integer
    Dim strTable  As String, str���� As String
    Dim blnRISChange As Boolean, blnTrans As Boolean
    strTable = Mid(lblTable.Tag, 2)
    str���� = Mid(lvwMain.SelectedItem.Key, 2)
    If mstrOwner = gstrSTOwner Then
        '֪ͨRIS������䶯
        '�ѱ����ʱû��ͨ���ù��߹����Ա������״��Ϊ�̶���
        If InStr(",�ѱ�,ҽ�Ƹ��ʽ,����,����״��,ְҵ,�Ա�,", "," & strTable & ",") > 0 Then
            blnRISChange = True
            If gblnMustRIS And Not gblnHaveRIS Then
                MsgBox "RIS�ӿڴ���ʧ�ܣ����ܼ������ֵ��" & strTable & "���е����������ǽӿ��ļ���װ��ע�᲻����������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If
    If MsgBox("��ȷ��Ҫɾ����" & Mid(lblTable.Tag, 2) & "��������Ϊ��" & lvwMain.SelectedItem.Text & "������Ŀ��", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
        gstrSQL = "delete from " & mstrOwner & "." & strTable & _
            " where ����='" & str���� & "'"
        '�ù��̽��з�װ
        lngSystem = cmbSys.ItemData(cmbSys.ListIndex) \ 100
        gcnOracle.BeginTrans: blnTrans = True
        gstrSQL = "ZL_�ֵ����_execute('" & Replace(gstrSQL, "'", "''") & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        If blnRISChange And gblnHaveRIS Then
            If gobjRIS.HISBasicDictTable(Decode(strTable, "�ѱ�", 4, "ҽ�Ƹ��ʽ", 5, "����", 6, "����״��", 7, "ְҵ", 8, "�Ա�", 9), 3, str����) <> 1 And gblnMustRIS Then
                gcnOracle.RollbackTrans
                MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ� ������Ӱ����Ϣϵͳ�ӿ�(HISBasicDictTable)δ���óɹ������ܽ��е�ǰ������", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        gcnOracle.CommitTrans: blnTrans = False
        With lvwMain
            '���浱ǰ��Ŀ������
            intIndex = .SelectedItem.Index
            .ListItems.Remove .SelectedItem.Key
            If .ListItems.Count > 0 Then
                intIndex = IIf(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
                .ListItems(intIndex).Selected = True
                .ListItems(intIndex).EnsureVisible
            End If
            Call SetMenu
        End With
    End If
    Exit Sub
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnuEditDeleteGroup_Click()
    On Error GoTo errHandle
    
    If Not tvwMain.SelectedItem Is Nothing Then
        Set mLastNode = tvwMain.SelectedItem
    End If
    If Not mLastNode Is Nothing Then
        If MsgBox("��ȷ��Ҫɾ����" & Mid(lblTable.Tag, 2) & "��������Ϊ��" & mLastNode.Text & "���ķ����Լ����������¼���Ŀ��", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
            '�ù��̽��з�װ
            gstrSQL = "Delete " & mstrOwner & "." & Mid(lblTable.Tag, 2) & _
                    " Where ���� In (Select ���� From " & mstrOwner & "." & Mid(lblTable.Tag, 2) & _
                    " Start With Nvl(�ϼ�, '0') = '" & Mid(mLastNode.Key, 2) & "'" & _
                    " Connect By Prior ���� = �ϼ�)"

            gstrSQL = "ZL_�ֵ����_execute('" & Replace(gstrSQL, "'", "''") & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            
            gstrSQL = "delete from " & mstrOwner & "." & Mid(lblTable.Tag, 2) & _
                " where ����='" & Mid(mLastNode.Key, 2) & "'"
            '�ù��̽��з�װ
            gstrSQL = "ZL_�ֵ����_execute('" & Replace(gstrSQL, "'", "''") & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            'չ��TreeView��Nodes
            Call frmRefresh
            TreeViewExpand tvwMain, True
            Call SetMenu
            
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnuEditModify_Click()
    frmDictEdit.�༭���� mcolSys("C" & cmbSys.ItemData(cmbSys.ListIndex)), Mid(lblTable.Tag, 2), Mid(lvwMain.SelectedItem.Key, 2), 1
End Sub

Private Sub mnuEditModifyGroup_Click()
    If Not mLastNode Is Nothing Then
        frmDictEdit.�༭���� mcolSys("C" & cmbSys.ItemData(cmbSys.ListIndex)), Mid(lblTable.Tag, 2), Mid(mLastNode.Key, 2), 0
    End If
End Sub

Private Sub mnuEditNew_Click()
    If tvwMain.Visible Then
        If Not tvwMain.SelectedItem Is Nothing Then
            Set mLastNode = tvwMain.SelectedItem
        End If
        tvwMain.SetFocus
        If Not mLastNode Is Nothing Then
            frmDictEdit.�༭���� mcolSys("C" & cmbSys.ItemData(cmbSys.ListIndex)), Mid(lblTable.Tag, 2), , 1, Mid(mLastNode.Key, 2)
        Else
            frmDictEdit.�༭���� mcolSys("C" & cmbSys.ItemData(cmbSys.ListIndex)), Mid(lblTable.Tag, 2), , 1
        End If
    Else
        frmDictEdit.�༭���� mcolSys("C" & cmbSys.ItemData(cmbSys.ListIndex)), Mid(lblTable.Tag, 2), , 1
    End If
End Sub

Private Sub mnuEditNewGroup_Click()
    
    If Not tvwMain.SelectedItem Is Nothing Then
        Set mLastNode = tvwMain.SelectedItem
    End If
    If Not mLastNode Is Nothing Then
        frmDictEdit.�༭���� mcolSys("C" & cmbSys.ItemData(cmbSys.ListIndex)), Mid(lblTable.Tag, 2), , 0, Mid(mLastNode.Key, 2)
    Else
        frmDictEdit.�༭���� mcolSys("C" & cmbSys.ItemData(cmbSys.ListIndex)), Mid(lblTable.Tag, 2), , 0
    End If
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub mnuViewReflash_Click()
    Call FillList
End Sub

Private Sub outTable_S_GotFocus()
    lvwMain.SetFocus
End Sub

Private Sub outTable_S_ItemClick(item As OutItem)
    If lblTable.Tag = item.Key Then Exit Sub
    Set mLastNode = Nothing
    lblTable.Tag = item.Key
    FillList
End Sub


Private Sub picsplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
    If Button = 1 Then
        If outTable_S.Width + X < 300 Then Exit Sub
        If tvwMain.Visible Then
            If tvwMain.Width - X < 220 Then Exit Sub
        End If
        
        picSplit.Left = picSplit.Left + X
        outTable_S.Width = outTable_S.Width + X
            
        picTable.Left = picTable.Left + X
        picTable.Width = picTable.Width - X
        lblTable.Width = picTable.Width
        '-- 10152����
        If tvwMain.Visible Then
            tvwMain.Left = picTable.Left
            tvwMain.Width = tvwMain.Width - X
        Else
            lvwMain.Left = picTable.Left
            lvwMain.Width = picTable.Width
        End If
            
    End If
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnuFilePreview_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub mnufileset_Click()
    zlPrintSet
End Sub

Private Sub picSplit2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 1 Then
        If tvwMain.Width + X < 220 Or lvwMain.Width - X < 200 Then
            Exit Sub
        End If
        tvwMain.Width = tvwMain.Width + X
        picSplit2.Left = picSplit2.Left + X
        lvwMain.Left = lvwMain.Left + X
        lvwMain.Width = lvwMain.Width - X
    End If
End Sub

Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim lvwTemp As ListView
    Select Case Button.Key
        Case "New"
            mnuEditNew_Click
        Case "Modify"
            mnuEditModify_Click
        Case "Delete"
            mnuEditDelete_Click
        Case "Exit"
            mnuFileExit_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Preview"
            mnuFilePreview_Click
        Case "Help"
            mnuhelptopic_Click
        Case "View"
            Set lvwTemp = IIf(Me.ActiveControl Is outTable_S, outTable_S, lvwMain)
            
            mnuViewIcon(lvwTemp.View).Checked = False
            If lvwTemp.View = 3 Then
                mnuViewIcon(0).Checked = True
                lvwTemp.View = 0
            Else
                mnuViewIcon(lvwTemp.View + 1).Checked = True
                lvwTemp.View = lvwTemp.View + 1
            End If
        Case "NewGroup"
            mnuEditNewGroup_Click
    End Select
    
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    clbOnly.Visible = mnuViewToolButton.Checked
    clbOnly.Bands("Comm").MinHeight = tlbMain.Height
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
    For Each buttTemp In tlbMain.Buttons
        If mnuViewToolText.Checked Then
            buttTemp.Caption = buttTemp.Tag
        Else
            buttTemp.Caption = ""
        End If
    Next
    clbOnly.Bands("Comm").MinHeight = tlbMain.Height
    Form_Resize
End Sub

Private Sub mnuShortMenu_Click(Index As Integer)
    Select Case Index
        Case 1
            mnuEditNew_Click
        Case 2
            mnuEditModify_Click
        Case 3
            mnuEditDelete_Click
    End Select
End Sub

Private Sub mnuGroupMenu_Click(Index As Integer)
    Select Case Index
        Case 1
            mnuEditNewGroup_Click
        Case 2
            mnuEditModifyGroup_Click
        Case 3
            mnuEditDeleteGroup_Click
    End Select
End Sub

Private Sub mnuShortIcon_Click(Index As Integer)
    mnuViewIcon_Click Index
End Sub

Private Sub mnuViewIcon_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
    Next
    mnuViewIcon(Index).Checked = True
    If Me.ActiveControl Is outTable_S Then
        outTable_S.View = Index
    Else
        lvwMain.View = Index
    End If
End Sub

Private Sub mnuhelptopic_Click()
   Call ShowHelp(App.ProductName, Me.hWnd, "ZL9AppTool\" & Me.Name, 0)
End Sub

Private Sub tlbMain_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim i As Integer
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
    Next
    mnuViewIcon(ButtonMenu.Index - 1).Checked = True
    If Me.ActiveControl Is outTable_S Then
        outTable_S.View = ButtonMenu.Index - 1
    Else
        lvwMain.View = ButtonMenu.Index - 1
    End If
End Sub

Private Sub subPrint(bytMode As Byte)
'����:���д�ӡ,Ԥ���������EXCEL
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Dim objPrint As New zlPrintLvw
    objPrint.Title.Text = Mid(lblTable.Tag, 2)
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

Private Function FillTable() As Boolean
'����:װ�����б༭��outTable_S
    Dim rsTemp As New ADODB.Recordset
    Dim item As OutItem
    Dim strOwner As String, strGroup As String
    Dim varTmp() As String, strFilter As String
    Dim i As Integer
    
    On Error GoTo ErrH
    If cmbSys.ListIndex = -1 Then Exit Function
    
    strOwner = UCase(mcolSys("C" & cmbSys.ItemData(Me.cmbSys.ListIndex)))
    cmbSys.Tag = strOwner
    
    gstrSQL = "select A.����,A.�̶�,A.˵��,A.����,B.privilege Ȩ�� " & _
            " from zlBasecode A," & _
            "    (select table_name,privilege from all_tab_privs where TABLE_SCHEMA=[1] and privilege in('SELECT','INSERT','DELETE','UPDATE')" & _
            "     union select table_name,'SELECT' from all_tables where owner=[1] and (owner=user or exists(select 1 from session_roles where ROLE='DBA') OR exists(select 1 from USER_SYS_PRIVS where PRIVILEGE='SELECT ANY TABLE'))" & _
            "     union select table_name,'INSERT' from all_tables where owner=[1] and (owner=user or exists(select 1 from session_roles where ROLE='DBA'))" & _
            "     union select table_name,'DELETE' from all_tables where owner=[1] and (owner=user or exists(select 1 from session_roles where ROLE='DBA'))" & _
            "     union select table_name,'UPDATE' from all_tables where owner=[1] and (owner=user or exists(select 1 from session_roles where ROLE='DBA'))) B " & _
            " Where a.���� = b.table_name and A.ϵͳ=[2] " & IIf(mstrTables <> "", " and instr([3],','||b.table_name||',')>0  ", "") & " order by A.����"
    
    rsTemp.CursorLocation = adUseClient
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strOwner, IIf(mlngSys = 0, Val(cmbSys.ItemData(cmbSys.ListIndex)), mlngSys), "," & mstrTables & ",")
    rsTemp.Filter = "Ȩ��='SELECT'"
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "��û�������õı����,�������б�����", vbExclamation, gstrSysName
        mblnFail = True
        Exit Function
    End If
    
    outTable_S.AutoRedraw = False
    outTable_S.Visible = False
    outTable_S.Groups.Clear
    strGroup = ""
    Do Until rsTemp.EOF
        If rsTemp("����") <> strGroup Then
            strGroup = rsTemp("����")
            outTable_S.Groups.Add , strGroup
        End If
        
        If IIf(IsNull(rsTemp("�̶�")), 0, rsTemp("�̶�")) = 0 Then
            outTable_S.Items.Add "K" & rsTemp("����"), rsTemp("����"), "Write", strGroup
        Else
            outTable_S.Items.Add "K" & rsTemp("����"), rsTemp("����"), "Read", strGroup
        End If
        
        rsTemp.MoveNext
    Loop
    For Each item In outTable_S.Items
        rsTemp.Filter = "����='" & item.Caption & "'"
        item.Tag = IIf(IsNull(rsTemp("˵��")), "", rsTemp("˵��"))
        Do Until rsTemp.EOF
            item.Tag = item.Tag & "'" & rsTemp("Ȩ��")
            rsTemp.MoveNext
        Loop
    Next
    outTable_S.Visible = True
    outTable_S.AutoRedraw = True
    
    lblTable.Tag = ""
    outTable_S_ItemClick outTable_S.Items(1)
    
    FillTable = True
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub FillList()
'����:װ���Ӧ��������Ŀ��lvwMain
    Dim strTable As String
    Dim rsTemp As New ADODB.Recordset
    Dim strKey As String
    Dim fld As Field
    Dim lst As ListItem
    
    On Error GoTo ErrH
    strTable = Mid(lblTable.Tag, 2)
    
    If Not lvwMain.SelectedItem Is Nothing Then
        '����ԭ�м�ֵ
        strKey = lvwMain.SelectedItem.Key
    End If
    
    mnuEditSplit.Visible = False
    mnuEditDefault.Visible = False
        
    If strTable = "" Then
        lvwMain.ListItems.Clear
        lvwMain.ColumnHeaders.Clear
        lvwMain.ColumnHeaders.Add , , "��ѡ�������ֵ�", 2000
        tvwMain.Nodes.Clear
        Call SetMenu
        Exit Sub
    End If
    
    '-- 10152�޸� ����Ƿ���ĩ��,�����ϼ���ʾ��TreeList��
    gstrSQL = "Select table_name From all_col_comments Where owner = '" & mstrOwner & "' And table_name='" & UCase(strTable) & "' And column_name='�ϼ�'"
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    If Not rsTemp.EOF Then
        tvwMain.Visible = True
        tvwMain.Tag = strTable
        picSplit2.Visible = True
        Call FillTree(mstrOwner & "." & strTable)
    Else
        tvwMain.Tag = ""
        tvwMain.Visible = False
        picSplit2.Visible = False
    End If
    Call Form_Resize
    
    If Not mLastNode Is Nothing And tvwMain.Tag <> "" Then
        Call ShowList(strTable, Mid(mLastNode.Key, 2))
    Else
        Call ShowList(strTable)
    End If
    ' strTable
    '---------
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub ShowList(ByVal strTable As String, Optional ByVal strTreeNodeKey As String)
    Dim rsTemp As New ADODB.Recordset
    Dim strKey As String
    Dim fld As Field
    Dim lst As ListItem
    Dim strWhere As String, strTmp As String
    Dim i As Long
    Dim blnFirst As Boolean
    Dim varPathProp As Variant
    Dim bytDisplayType() As Byte
    Dim blnHide As Boolean
    
    On Error GoTo ErrH
    '�Ƿ�·���������
    strTmp = IsPathProperty(mstrOwner, strTable)
    If Len(strTmp) > 0 Then
        varPathProp = Split(strTmp, ";")
    End If
    
    rsTemp.CursorLocation = adUseClient

    If tvwMain.Tag <> "" Then
        strWhere = " Where ĩ��=1"
        If strTreeNodeKey = "" Or strTreeNodeKey = "oot" Then
            strWhere = strWhere & " And Nvl(�ϼ�,0)=0"
        Else
            strWhere = strWhere & " And �ϼ�='" & strTreeNodeKey & "'"
        End If
    End If
    
    If varPathProp(0) <> "" Then
        If strTable = "����" Then
            gstrSQL = "select a.*,'['||a." & varPathProp(0) & "||']'||b.���� as PathProp from " & mstrOwner & "." & strTable & " a, " & varPathProp(2) & " b " _
                    & IIf(Len(strWhere) <= 0, " where a.", strWhere & " and a.") & varPathProp(0) & "=b." & varPathProp(1) & "(+)"
        ElseIf strTable <> varPathProp(2) Then
            gstrSQL = "select a.*,b.���� as PathProp from " & mstrOwner & "." & strTable & " a, " & varPathProp(2) & " b " _
                    & IIf(Len(strWhere) <= 0, " where a.", strWhere & " and a.") & varPathProp(0) & "=b." & varPathProp(1) & "(+)"
        Else
            gstrSQL = "select * from " & mstrOwner & "." & strTable & strWhere
        End If
    Else
        gstrSQL = "select * from " & mstrOwner & "." & strTable & strWhere
    End If
    
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    
    mintColumn = 0
    ReDim bytDisplayType(rsTemp.Fields.Count)
    
    zlControl.FormLock lvwMain.hWnd
    lvwMain.ColumnHeaders.Clear
    lvwMain.ColumnHeaders.Add , "����", "����"
    For Each fld In rsTemp.Fields
        If InStr(",����,�ϼ�,ĩ��," & varPathProp(0) & ",", "," & fld.Name & ",") <= 0 Then
            If UCase(fld.Name) = "��ԴID" And UBound(varPathProp) >= 2 Then
                '����ʾ�������ԴID���ؼ���ResourceInfo��������ֶ�
                If UCase(varPathProp(2)) = "RESOURCEINFO" Then
                    blnHide = True
                Else
                    lvwMain.ColumnHeaders.Add , IIf(fld.Name = "PATHPROP", varPathProp(0), fld.Name), IIf(fld.Name = "PATHPROP", varPathProp(0), fld.Name)
                End If
            Else
                lvwMain.ColumnHeaders.Add , IIf(fld.Name = "PATHPROP", varPathProp(0), fld.Name), IIf(fld.Name = "PATHPROP", varPathProp(0), fld.Name)
            End If
        End If

        If fld.Name = "ȱʡ��־" Then
            '�ɼ�
            mnuEditSplit.Visible = True
            mnuEditDefault.Visible = True
        End If
    Next
    lvwMain.ListItems.Clear
    Do Until rsTemp.EOF
        If tvwMain.Tag <> "" Then
            Dim strIcon As String
            strIcon = IIf(zlCommFun.NVL(rsTemp("ĩ��"), 0) = 1, "Item", "Group")
            Set lst = lvwMain.ListItems.Add(, "C" & rsTemp("����"), IIf(IsNull(rsTemp("����")), "", rsTemp("����")), strIcon, strIcon)
        Else
            Set lst = lvwMain.ListItems.Add(, "C" & rsTemp("����"), IIf(IsNull(rsTemp("����")), "", rsTemp("����")), "Item", "Item")
        End If

'        For Each fld In rsTemp.Fields
'            '-- 10152�޸� ����ĩ���Ĵ���
'            If fld.Name = "ȱʡ��־" Or fld.Name Like "�Ƿ�*" Then
'                lst.SubItems(lvwMain.ColumnHeaders(fld.Name).Index - 1) = IIf(fld.Value = 1, "��", "")
'            Else
'                If InStr(",����,�ϼ�,ĩ��,", "," & fld.Name & ",") <= 0 Then
'                    lst.SubItems(lvwMain.ColumnHeaders(fld.Name).Index - 1) = IIf(IsNull(fld.Value), "", fld.Value)
'                End If
'            End If
'        Next
        
        If blnFirst Then
            For i = 0 To rsTemp.Fields.Count - 1
                Select Case bytDisplayType(i)
                Case 1, 3
                    lst.SubItems(lvwMain.ColumnHeaders(rsTemp.Fields(i).Name).Index - 1) = IIf(rsTemp.Fields(i).Value = 1, "��", "")
                Case 2
                    'ת����ͷ��
                    If rsTemp.Fields(i).Name = "PATHPROP" Then
                        lst.SubItems(lvwMain.ColumnHeaders(varPathProp(0)).Index - 1) = IIf(IsNull(rsTemp.Fields(i).Value), "", rsTemp.Fields(i).Value)
                    ElseIf rsTemp.Fields(i).Name = "����" And strTable = "���쳣��ԭ��" Then
                        If rsTemp.Fields(i).Value = 0 Then lst.SubItems(lvwMain.ColumnHeaders(rsTemp.Fields(i).Name).Index - 1) = "δ�����ԭ��"
                        If rsTemp.Fields(i).Value = 1 Then lst.SubItems(lvwMain.ColumnHeaders(rsTemp.Fields(i).Name).Index - 1) = "���������ԭ��"
                        If rsTemp.Fields(i).Value = 2 Then lst.SubItems(lvwMain.ColumnHeaders(rsTemp.Fields(i).Name).Index - 1) = "�����˳���ԭ��"
                    ElseIf blnHide And UCase(rsTemp.Fields(i).Name) = "��ԴID" Then
                        '����ʾ
                    Else
                        lst.SubItems(lvwMain.ColumnHeaders(rsTemp.Fields(i).Name).Index - 1) = IIf(IsNull(rsTemp.Fields(i).Value), "", rsTemp.Fields(i).Value)
                    End If
                End Select
            Next
        Else
            '��һ��
            For i = 0 To rsTemp.Fields.Count - 1
                If rsTemp.Fields(i).Name = "ȱʡ��־" Or rsTemp.Fields(i).Name Like "�Ƿ�*" Then
                    lst.SubItems(lvwMain.ColumnHeaders(rsTemp.Fields(i).Name).Index - 1) = IIf(rsTemp.Fields(i).Value = 1, "��", "")
                    bytDisplayType(i) = 1
                ElseIf rsTemp.Fields(i).Name = "����" And strTable = "���쳣��ԭ��" Then
                    If rsTemp.Fields(i).Value = 0 Then lst.SubItems(lvwMain.ColumnHeaders(rsTemp.Fields(i).Name).Index - 1) = "δ�����ԭ��"
                    If rsTemp.Fields(i).Value = 1 Then lst.SubItems(lvwMain.ColumnHeaders(rsTemp.Fields(i).Name).Index - 1) = "���������ԭ��"
                    If rsTemp.Fields(i).Value = 2 Then lst.SubItems(lvwMain.ColumnHeaders(rsTemp.Fields(i).Name).Index - 1) = "�����˳���ԭ��"
                    bytDisplayType(i) = 2
                ElseIf rsTemp.Fields(i).Type = adNumeric And rsTemp.Fields(i).Precision = 1 And InStr(",����,�ϼ�,ĩ��,", "," & rsTemp.Fields(i).Name & ",") <= 0 Then
                    If IsCheckConstraint(mstrOwner, strTable, rsTemp.Fields(i).Name, 1) Then
                        lst.SubItems(lvwMain.ColumnHeaders(rsTemp.Fields(i).Name).Index - 1) = IIf(rsTemp.Fields(i).Value = 1, "��", "")
                        bytDisplayType(i) = 3
                    Else
                        bytDisplayType(i) = 99
                    End If
                Else
                    If InStr(",����,�ϼ�,ĩ��," & varPathProp(0) & ",", "," & rsTemp.Fields(i).Name & ",") <= 0 Then
                        'ת����ͷ��
                        If rsTemp.Fields(i).Name = "PATHPROP" Then
                            lst.SubItems(lvwMain.ColumnHeaders(varPathProp(0)).Index - 1) = IIf(IsNull(rsTemp.Fields(i).Value), "", rsTemp.Fields(i).Value)
                        ElseIf blnHide And UCase(rsTemp.Fields(i).Name) = "��ԴID" Then
                            '������
                        Else
                            lst.SubItems(lvwMain.ColumnHeaders(rsTemp.Fields(i).Name).Index - 1) = IIf(IsNull(rsTemp.Fields(i).Value), "", rsTemp.Fields(i).Value)
                        End If
                        bytDisplayType(i) = 2
                    Else
                        bytDisplayType(i) = 99
                    End If
                End If
            Next
            blnFirst = True
        End If
        
        For Each fld In rsTemp.Fields
            If fld.Name = "��ɫ" Then
                lst.ForeColor = IIf(IsNull(rsTemp!��ɫ), 0, rsTemp!��ɫ)
                For i = 1 To lst.ListSubItems.Count
                    lst.ListSubItems.item(i).ForeColor = IIf(IsNull(rsTemp!��ɫ), 0, rsTemp!��ɫ)
                Next
            End If
        Next

        rsTemp.MoveNext
    Loop
    'ʹ�п�����Ӧ
    For i = 0 To lvwMain.ColumnHeaders.Count - 1
        SendMessage lvwMain.hWnd, LVM_SETCOLUMNWIDTH, i, LVSCW_AUTOSIZE_USEHEADER
        If lvwMain.ColumnHeaders(i + 1).Width < 600 Then lvwMain.ColumnHeaders(i + 1).Width = 600
    Next
    zlControl.FormLock 0
    
    If lvwMain.ListItems.Count > 0 Then
        Dim item As ListItem
        On Error Resume Next
        Set item = lvwMain.ListItems(strKey)
        If Err <> 0 Then
            Set item = lvwMain.ListItems(1)
            item.Selected = True
            item.EnsureVisible
        Else
            Err.Clear
            item.Selected = True
            item.EnsureVisible
        End If
    End If
    Call SetMenu
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub SetMenu()
    Dim item As OutItem
    Dim i As Integer
    Dim blnPrint As Boolean
    
    Set item = outTable_S.Items(lblTable.Tag)
    If item.Icon = "Write" Then
        If lvwMain.ListItems.Count = 0 Then
            mnuEditNew.Enabled = (InStr(item.Tag, "'INSERT") > 0)
            mnuEditDelete.Enabled = False
            mnuEditModify.Enabled = False
        Else
            mnuEditNew.Enabled = (InStr(item.Tag, "'INSERT") > 0)
            mnuEditDelete.Enabled = (InStr(item.Tag, "'DELETE") > 0)
            mnuEditModify.Enabled = (InStr(item.Tag, "'UPDATE") > 0)
        End If
        
        If tvwMain.Visible Then
            If tvwMain.Nodes.Count <= 1 Then
                mnuEditNewGroup.Enabled = (InStr(item.Tag, "'INSERT") > 0)
                mnuEditDeleteGroup.Enabled = False
                mnuEditModifyGroup.Enabled = False
            Else
                mnuEditNewGroup.Enabled = (InStr(item.Tag, "'INSERT") > 0)
                mnuEditDeleteGroup.Enabled = (InStr(item.Tag, "'DELETE") > 0)
                mnuEditModifyGroup.Enabled = (InStr(item.Tag, "'UPDATE") > 0)
                
                If Not tvwMain.SelectedItem Is Nothing Then
                    If tvwMain.SelectedItem.Key = "Root" Then
                        mnuEditDeleteGroup.Enabled = False
                        mnuEditModifyGroup.Enabled = False
                    End If
                End If
            End If
        Else
            mnuEditNewGroup.Enabled = False
            mnuEditDeleteGroup.Enabled = False
            mnuEditModifyGroup.Enabled = False
        End If
    Else
        mnuEditNew.Enabled = False
        mnuEditDelete.Enabled = False
        mnuEditModify.Enabled = False
        
        mnuEditNewGroup.Enabled = False
        mnuEditDeleteGroup.Enabled = False
        mnuEditModifyGroup.Enabled = False
    End If
    tlbMain.Buttons("New").Enabled = mnuEditNew.Enabled
    tlbMain.Buttons("Modify").Enabled = mnuEditModify.Enabled
    tlbMain.Buttons("Delete").Enabled = mnuEditDelete.Enabled
    tlbMain.Buttons("NewGroup").Enabled = mnuEditNewGroup.Enabled
    
    mnuEditDelete.Enabled = mnuEditModify.Enabled And Not lvwMain.SelectedItem Is Nothing
    
    blnPrint = lvwMain.ListItems.Count > 0
    tlbMain.Buttons("Preview").Enabled = blnPrint
    tlbMain.Buttons("Print").Enabled = blnPrint
    mnuFilePreview.Enabled = blnPrint
    mnuFilePrint.Enabled = blnPrint
    mnuFileExcel.Enabled = blnPrint
    
    i = InStr(item.Tag, "'")
    lblTable.Caption = " " & item.Caption & IIf(i > 0, "����" & Mid(item.Tag, 1, i - 1), "")
    stbThis.Panels(2) = "���ֵ乲��" & lvwMain.ListItems.Count & "�����롣"
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

Private Sub FillTree(ByVal strTable As String)
    
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim nodTmp As Node
    
    On Error GoTo ErrH
    strSQL = " Select * From " & strTable & " Where nvl(ĩ��,0)=0 Start with Nvl(�ϼ�,0)=0 connect by prior ���� =�ϼ�"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With tvwMain
        .Nodes.Clear
        .Nodes.Add , , "Root", "ȫ��", "Root", "Root"
        Do Until rsTmp.EOF
            If IsNull(rsTmp!�ϼ�) Then
                tvwMain.Nodes.Add "Root", tvwChild, "B" & rsTmp!����, "[" & rsTmp!���� & "]" & rsTmp!����, "Group", "GroupOpen"
            Else
                If nodTmp Is Nothing Then
                    Set nodTmp = tvwMain.Nodes.Add("B" & rsTmp!�ϼ�, tvwChild, "B" & rsTmp!����, "[" & rsTmp!���� & "]" & rsTmp!����, "Group", "Group")
                Else
                    tvwMain.Nodes.Add "B" & rsTmp!�ϼ�, tvwChild, "B" & rsTmp!����, "[" & rsTmp!���� & "]" & rsTmp!����, "Group", "GroupOpen"
                End If
            End If
            rsTmp.MoveNext
        Loop
        .Nodes.item("Root").Expanded = True
        If mLastNode Is Nothing Then
            .Nodes.item("Root").Selected = True
            Call tvwMain_NodeClick(.Nodes.item("Root"))
        Else
            .Nodes.item(mLastNode.Key).Selected = True
            Call tvwMain_NodeClick(.Nodes.item(mLastNode.Key))
        End If
    End With
    
    '����Icon��ʾ״̬
'    Dim i As Integer
'    For i = 2 To tvwMain.Nodes.Count
'        tvwMain.Nodes(i).Image = 5
'        If Not tvwMain.Nodes(i).Child Is Nothing Then
'            tvwMain.Nodes(i).ExpandedImage = 6
'        End If
'    Next
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub tvwMain_DblClick()
    If Not tvwMain.SelectedItem Is Nothing And mnuEditModifyGroup.Enabled And mnuEditModifyGroup.Visible Then mnuEditModifyGroup_Click
End Sub

Private Sub tvwMain_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Not tvwMain.SelectedItem Is Nothing And mnuEditModifyGroup.Enabled And mnuEditModifyGroup.Visible Then mnuEditModifyGroup_Click
    End If
End Sub

Private Sub tvwMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        mnuGroupMenu(1).Enabled = mnuEditNewGroup.Enabled
        mnuGroupMenu(2).Enabled = mnuEditModifyGroup.Enabled
        mnuGroupMenu(3).Enabled = mnuEditDeleteGroup.Enabled
        
        PopupMenu mnuGroup, vbPopupMenuRightButton
    End If
End Sub

Private Sub tvwMain_NodeClick(ByVal Node As MSComctlLib.Node)
    If Not mLastNode Is Nothing Then
        If mLastNode = Node Then Exit Sub
    End If
    Node.Expanded = True
    If tvwMain.Tag <> "" Then
        Call ShowList(tvwMain.Tag, Mid(Node.Key, 2))
    End If
    Set mLastNode = Node
End Sub

Public Sub frmRefresh()
    Set mLastNode = Nothing
    Call FillList
End Sub

'����/չ��TreeView�ؼ�Nodes
Public Sub TreeViewExpand(ByVal objTV As TreeView, Optional blgExpand As Boolean = False)
    Dim i As Integer
    
    For i = 1 To objTV.Nodes.Count
        objTV.Nodes(i).Expanded = blgExpand
    Next
End Sub

Private Sub txtSeek_GotFocus()
    txtSeek.SelStart = 0
    txtSeek.SelLength = Len(txtSeek.Text)
End Sub

Private Sub txtSeek_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Dim i As Long, strtxt As String, lngStart As Long
        strtxt = Trim(txtSeek.Text)
        If strtxt = "" Then Exit Sub
        
        If strtxt = txtSeek.Tag And Not (mlngLastPos = 0 Or mlngLastPos = outTable_S.Items.Count) Then
            lngStart = mlngLastPos + 1
        Else
            lngStart = 1
        End If
        With outTable_S
            For i = lngStart To .Items.Count
                If .Items(i).Caption Like "*" & strtxt & "*" Then
                    .SelectGroup = .Items(i).GroupName
                    .SelectItem = .Items(i).Caption
                    Call outTable_S_ItemClick(.Items(i))
                    mlngLastPos = i
                    Exit For
                End If
            Next
            If i > .Items.Count Then mlngLastPos = 1  'û���ҵ�
        End With
        
        txtSeek.Tag = strtxt
    'ElseIf KeyAscii = Asc("'") Then
    ElseIf InStr("~!@#$%^&*()_+-=`{}[]:|;'\<>?,./", Chr(KeyAscii)) > 1 Then
        KeyAscii = 0
    End If
End Sub

