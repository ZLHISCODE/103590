VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmIllManage 
   BackColor       =   &H8000000A&
   Caption         =   "�����������"
   ClientHeight    =   6750
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   8730
   Icon            =   "frmIllManage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cdgFile 
      Left            =   6000
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2625
      Left            =   4980
      MousePointer    =   9  'Size W E
      ScaleHeight     =   2625
      ScaleWidth      =   45
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2070
      Width           =   45
   End
   Begin MSComctlLib.ListView lvwMain 
      Height          =   2475
      Left            =   4050
      TabIndex        =   2
      Top             =   2070
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   4366
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
      BackColor       =   16777215
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView tvwMain_S 
      Height          =   4575
      Left            =   150
      TabIndex        =   1
      Top             =   1170
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   8070
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   4380
      Top             =   5490
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
            Picture         =   "frmIllManage.frx":030A
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":075E
            Key             =   "Stop"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   3810
      Top             =   5430
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":0BB0
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":1004
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":1458
            Key             =   "Stop"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar coolbar1 
      Align           =   1  'Align Top
      Height          =   1125
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   1984
      BandCount       =   2
      _CBWidth        =   8730
      _CBHeight       =   1125
      _Version        =   "6.7.9782"
      Child1          =   "Toolbar1"
      MinHeight1      =   720
      Width1          =   8535
      Key1            =   "only"
      NewRow1         =   0   'False
      Caption2        =   "�������"
      Child2          =   "cmbType"
      MinWidth2       =   3495
      MinHeight2      =   300
      Width2          =   1590
      FixedBackground2=   0   'False
      NewRow2         =   -1  'True
      Begin VB.ComboBox cmbType 
         Height          =   300
         ItemData        =   "frmIllManage.frx":18AA
         Left            =   945
         List            =   "frmIllManage.frx":18C3
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   780
         Width           =   7695
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   720
         Left            =   165
         TabIndex        =   5
         Top             =   30
         Width           =   8475
         _ExtentX        =   14949
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "Ilsrw"
         HotImageList    =   "Ilscolor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   17
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
               Key             =   "Split0"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Class"
               Object.ToolTipText     =   "���ӷ���"
               Object.Tag             =   "����"
               ImageKey        =   "Class"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Disease"
               Object.ToolTipText     =   "���Ӽ���"
               Object.Tag             =   "����"
               ImageKey        =   "Disease"
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
               Key             =   "����"
               Object.ToolTipText     =   "���ü�������"
               Object.Tag             =   "����"
               ImageKey        =   "Start"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ͣ��"
               Key             =   "ͣ��"
               Object.ToolTipText     =   "ͣ�ü�������"
               Object.Tag             =   "ͣ��"
               ImageKey        =   "Stop"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Splits"
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Find"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split3"
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split4"
               Style           =   3
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Quit"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList Ilsrw 
      Left            =   6900
      Top             =   1020
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
            Picture         =   "frmIllManage.frx":1925
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":1B3F
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":1D59
            Key             =   "Class"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":1F75
            Key             =   "Disease"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":2191
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":23B1
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":25D1
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":27EB
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":2A0B
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":2C2B
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":2E4B
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":3065
            Key             =   "Stop"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Ilscolor 
      Left            =   5880
      Top             =   1020
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
            Picture         =   "frmIllManage.frx":327F
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":349F
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":36BF
            Key             =   "Class"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":38DB
            Key             =   "Disease"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":3AF7
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":3D17
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":3F37
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":4151
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":4371
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":4591
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":47B1
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":49CB
            Key             =   "Stop"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   6390
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   635
      SimpleText      =   "CoolBar1"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmIllManage.frx":4BE5
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10319
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
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFilePrintView 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileEXCEL 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFileImport 
         Caption         =   "���뼲������"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuFileExport 
         Caption         =   "������������"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditClass 
         Caption         =   "���ӷ���(&C)"
      End
      Begin VB.Menu mnuEditDisease 
         Caption         =   "���Ӽ���(&D)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "�޸�(&M)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "ɾ��(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditStart 
         Caption         =   "����(&S)"
      End
      Begin VB.Menu mnuEditStop 
         Caption         =   "ͣ��(&T)"
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
            Caption         =   "��׼�ı�(&S)"
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
      Begin VB.Menu mnuViewLine1 
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
      Begin VB.Menu mnuViewLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewOrder 
         Caption         =   "����������(&O)"
      End
      Begin VB.Menu mnuViewSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewStop 
         Caption         =   "��ʾͣ����Ŀ(&P)"
      End
      Begin VB.Menu mnuViewLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewColumn 
         Caption         =   "ѡ����(&C)"
      End
      Begin VB.Menu mnuViewAll 
         Caption         =   "ȫ����ʾ(&A)"
      End
      Begin VB.Menu mnuViewLine4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "����(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "ˢ��(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "��������(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "Web�ϵ�����"
         Begin VB.Menu mnuHelpWebHomePage 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "������̳(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&K)..."
         End
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
   Begin VB.Menu mnuShort1 
      Caption         =   "��ݲ˵�"
      Visible         =   0   'False
      Begin VB.Menu mnushortMenu1 
         Caption         =   "���ӷ���(&C)"
         Index           =   1
      End
      Begin VB.Menu mnushortMenu1 
         Caption         =   "�޸ķ���(&M)"
         Index           =   2
      End
      Begin VB.Menu mnushortMenu1 
         Caption         =   "ɾ������(&D)"
         Index           =   3
      End
   End
   Begin VB.Menu mnuShort2 
      Caption         =   "��ݲ˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "���Ӽ���(&D)"
         Index           =   1
      End
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "�޸ļ���(&M)"
         Index           =   2
      End
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "ɾ������(&D)"
         Index           =   3
      End
      Begin VB.Menu mnuShortLine 
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
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmIllManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mintColumn As Integer
Dim mstr������� As String
Dim mblnLoad As Boolean

Dim msngStartX As Single    '�ƶ�ǰ����λ��
Dim mlng��� As Long
Dim mstrNodeKey As String, mstrTypeText As String
'ÿ�����Ķεĺ������������ơ���ȡ����롢��ѡ��
Private Const mstr���� As String = "����,1200,0,1;����,1200,0,0;����,2440,0,2;ƴ����,1000,0,0;�����,1000,0,0;�Ա�����,800,0,0;��������,800,0,0;ͳ����,800,0,0;������Ч,800,0,0;������Ϣ,800,0,0;˵��,3000,0,0;����ʱ��,1400,0,0;����ʱ��,1400,0,0"

Private mlngMode As Long
Private mstrPrivs As String 'Ȩ�޴�
Private mconnExcel As New ADODB.Connection
Private mintProgress As Integer

Private Sub Form_Load()
    Dim i As Long
    
    mlngMode = glngModul
    mstrPrivs = gstrPrivs
    
    Call Ȩ�޿���
    '���������ɾ����ListView�������
    lvwMain.Tag = "�ɱ仯��"
    RestoreWinState Me, App.ProductName
    Call zldatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    
    For i = 0 To 3
        Me.mnuViewIcon(i).Checked = False
    Next
    Me.mnuViewIcon(Me.lvwMain.View).Checked = True
    
    mnuViewAll.Checked = (Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "��ʾ����", 0)) = 1)
    
    mblnLoad = True
    '���ListView�Ļ�δ�����ã������һ��ʹ�ã��Ǿ͵���ȱʡ�ĳ�ʼ��
'    If lvwMain.ColumnHeaders.Count = 0 Then
        zlControl.LvwSelectColumns lvwMain, mstr����, True
'    End If
End Sub

Private Sub Form_Activate()
    On Error GoTo errHandle
    
    If mblnLoad = True Then
        mblnLoad = False '���ϰ����Ĺ���
                
        '20031112byZT��ͨ��Ȩ���ж��Ƿ�ʹ����ҽ
        gblnʹ����ҽ = InStr(mstrPrivs, "��ҽ") > 0
        
        '��ʼ����������б�������
        Call Fill���
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Unload Me
End Sub

Private Sub Form_Resize()
    If WindowState = 1 Then Exit Sub
    
    Dim sngTop As Single, sngBottom As Single
    On Error Resume Next
    sngTop = IIF(CoolBar1.Visible, CoolBar1.Top + CoolBar1.Height, 0)
    sngBottom = Me.ScaleHeight - IIF(stbThis.Visible, stbThis.Height, 0)
    
    tvwMain_S.Left = ScaleLeft
    tvwMain_S.Top = sngTop
    tvwMain_S.Height = sngBottom - sngTop
    
    With picSplit
        .Left = tvwMain_S.Left + tvwMain_S.Width
        .Top = tvwMain_S.Top
        .Height = tvwMain_S.Height
    End With
    
    lvwMain.Top = tvwMain_S.Top
    lvwMain.Height = tvwMain_S.Height
    
    If tvwMain_S.Visible = True Then
        lvwMain.Left = picSplit.Left + picSplit.Width
    Else
        lvwMain.Left = ScaleLeft
    End If
    lvwMain.Width = ScaleWidth - lvwMain.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mconnExcel.State = 1 Then mconnExcel.Close
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "��ʾ����", IIF(mnuViewAll.Checked, 1, 0)
    SaveWinState Me, App.ProductName
End Sub

Private Sub cmbType_Click()
    If cmbType.Text = mstrTypeText Then Exit Sub
    
    If cmbType.ItemData(cmbType.ListIndex) = 1 Then
        tvwMain_S.Visible = True
        picSplit.Visible = True
    Else
        tvwMain_S.Visible = False
        picSplit.Visible = False
    End If
    Call Form_Resize
    
    Call FillTree
End Sub

Private Sub coolbar1_HeightChanged(ByVal NewHeight As Single)
    Form_Resize
End Sub

Private Sub lvwMain_DblClick()
    
    If mnuEditModify.Visible And mnuEditModify.Enabled Then
        '�Ե�ǰ��Ŀ���б༭
        Call mnuEditModify_Click
    End If

End Sub

Private Sub lvwMain_GotFocus()
    Call SetMenu
End Sub


Private Sub lvwMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
     Call SetMenu
End Sub

Private Sub lvwMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    
    If Button = 2 Then
        mnuShortMenu2(1).Enabled = mnuEditDisease.Enabled
        mnuShortMenu2(2).Enabled = mnuEditModify.Enabled
        mnuShortMenu2(3).Enabled = mnuEditDelete.Enabled
        For i = 0 To 3
            mnuShortIcon(i).Checked = mnuViewIcon(i).Checked
        Next
        PopupMenu mnuShort2, vbPopupMenuRightButton
    End If
End Sub

Private Sub lvwMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then '���Ǹղ�����
        lvwMain.SortOrder = IIF(lvwMain.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwMain.SortKey = mintColumn
        lvwMain.SortOrder = lvwAscending
    End If
    If Not lvwMain.SelectedItem Is Nothing Then
        lvwMain.SelectedItem.EnsureVisible
    End If
End Sub


Private Sub mnuEditStart_Click()
    Call StopAndResume(False)
End Sub

Private Sub mnuEditDelete_Click()
'ɾ��
    Dim strKey As String
    Dim intIndex As Long
    
    If AllowContinue = False Then Exit Sub
    
    On Error GoTo errHandle
    If ActiveControl Is tvwMain_S Then
        If MsgBox("��ȷ��Ҫɾ������Ϊ��" & tvwMain_S.SelectedItem.Text & "���ķ�����", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
            Me.MousePointer = 11
            
            gstrSQL = "ZL_�����������_delete(" & Mid(tvwMain_S.SelectedItem.Key, 2) & ")"
            Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            
            Me.MousePointer = 0
            
            strKey = tvwMain_S.SelectedItem.Key
            If Not tvwMain_S.SelectedItem.Next Is Nothing Then
                tvwMain_S.SelectedItem.Next.Selected = True
            Else
                If Not tvwMain_S.SelectedItem.Parent Is Nothing Then
                    tvwMain_S.SelectedItem.Parent.Selected = True
                End If
            End If
            Call FillList
            tvwMain_S.Nodes.Remove strKey
        End If
    Else
        If MsgBox("��ȷ��Ҫɾ������Ϊ��" & lvwMain.SelectedItem.Text & "���ļ�����", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
            Me.MousePointer = 11
            
            On Error Resume Next
            gstrSQL = "ZL_��������Ŀ¼_delete(" & Mid(lvwMain.SelectedItem.Key, 2) & ")"
            Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            If Err.Number <> 0 Then
                If InStr(Err.Description, "ORA-20005") > 0 Then
                    MsgBox "��Ŀ�Ѿ�ʹ�ò���ɾ����ֻ��ͣ��", vbInformation, gstrSysName
   
                Else
                    MsgBox Err.Description, vbInformation, gstrSysName
                End If
                
                Me.MousePointer = 0
                Exit Sub
            End If
            
            Me.MousePointer = 0
            
            With lvwMain
                intIndex = .SelectedItem.Index
                .ListItems.Remove .SelectedItem.Key
                If .ListItems.Count > 0 Then
                    intIndex = IIF(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
                    .ListItems(intIndex).Selected = True
                    .ListItems(intIndex).EnsureVisible
                End If
                Call SetMenu
            End With
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Me.MousePointer = 0
End Sub



Private Sub mnuEditModify_Click()
    Dim nodTemp As Node
    Dim str���� As String
    
    If AllowContinue = False Then Exit Sub
    
    Set nodTemp = tvwMain_S.SelectedItem
    If ActiveControl Is tvwMain_S And tvwMain_S.Visible = True Then
        '�޸ķ���
        If nodTemp Is Nothing Then
            Exit Sub
        Else
            
            With tvwMain_S.SelectedItem
                frmIllSortEdit.�����༭ "", "", mstr�������, Mid(nodTemp.Key, 2)
            End With
        End If
    Else
        '�޸ļ���
        If lvwMain.SelectedItem Is Nothing Then Exit Sub
        If nodTemp Is Nothing Then
            If tvwMain_S.Visible = True Then Exit Sub '��������ǲ�����
            
            frmIllItemEdit.�����༭ tvwMain_S.Visible, "��", "", mstr�������, Mid(lvwMain.SelectedItem.Key, 2)
        Else
            
            frmIllItemEdit.�����༭ tvwMain_S.Visible, nodTemp.Text, Mid(nodTemp.Key, 2), mstr�������, Mid(lvwMain.SelectedItem.Key, 2)
        End If
    End If
End Sub

Private Sub mnuEditClass_Click()
    Dim nodTemp As Node
    Dim str���� As String
    
    If AllowContinue = False Then Exit Sub
    
    Set nodTemp = tvwMain_S.SelectedItem
    '���ӷ���
    If nodTemp Is Nothing Then
        frmIllSortEdit.�����༭ "��", "", mstr�������
    Else
        frmIllSortEdit.�����༭ nodTemp.Text, Mid(nodTemp.Key, 2), mstr�������
    End If
End Sub

Private Sub mnuEditDisease_Click()
    Dim nodTemp As Node
    Dim str���� As String
    
    If AllowContinue = False Then Exit Sub
    
    Set nodTemp = tvwMain_S.SelectedItem
    '���Ӽ���
    If nodTemp Is Nothing Then
        If tvwMain_S.Visible = True Then Exit Sub '��������ǲ�����
        
        frmIllItemEdit.�����༭ tvwMain_S.Visible, "��", "", mstr�������
    Else
        If Mid(nodTemp.Key, 2) = 1 Then
            If MsgBox("�ڸ���Ŀ�����Ӽ���������ϵͳ�Դ�����ļ�������Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        
        frmIllItemEdit.�����༭ tvwMain_S.Visible, nodTemp.Text, Mid(nodTemp.Key, 2), mstr�������
    End If
End Sub

Private Function AllowContinue() As Boolean
'����Ƿ���������༭����
    If MsgBox("���ʼ���������Ҫͳһ�ı�׼������һ���ǳ�������£�" & vbCrLf & _
        "��������ڵ�������ͳ�Ƶ�Ȩ������ָ������ɱ�������" & vbCrLf & vbCrLf & _
        "�Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    
    AllowContinue = True
End Function

Private Sub mnuEditStop_Click()
    Call StopAndResume(True)
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnuFileExport_Click()
    Dim strPath As String
    strPath = zlcommfun.OpenDir(Me.hwnd, "����Ŀ¼", App.Path)
    If strPath <> "" Then
        If Not Right(strPath, 1) = "\" Then
            strPath = strPath & "\"
        End If
        Call FuncCreateSQL(strPath)
    End If
End Sub

Private Sub mnuFileImport_Click()
    Dim objfrm As Form
    Set objfrm = New frmIllImport
    Call objfrm.ShowMe(Me)
    Call FillTree
End Sub

Private Sub mnuFilePrintView_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub mnuFilePrintset_Click()
    zlPrintSet
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub mnuHelpWebHomePage_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    'Ĭ�ϲ��������=�����룬����=����id����Ŀ=Ŀ¼id
    Dim str������ As String
    Dim lng����id As Long
    Dim lng��Ŀid As Long
    
    If cmbType.ListIndex <> -1 Then
        str������ = Mid(cmbType.List(cmbType.ListIndex), 1, 1)
    End If
    
    If Not tvwMain_S.SelectedItem Is Nothing Then
        lng����id = Mid(tvwMain_S.SelectedItem.Key, 2)
    End If
    
    If Not lvwMain.SelectedItem Is Nothing Then
        lng��Ŀid = Mid(lvwMain.SelectedItem.Key, 2)
    End If
    
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
        "���=" & str������, _
        "����=" & IIF(lng����id = 0, "", lng����id), _
        "��Ŀ=" & IIF(lng��Ŀid = 0, "", lng��Ŀid))
End Sub

Private Sub mnuShortIcon_Click(Index As Integer)
    mnuViewIcon_Click Index
End Sub

Private Sub mnuShortMenu1_Click(Index As Integer)
    Select Case Index
        Case 1
            Call mnuEditClass_Click
        Case 2
            Call mnuEditModify_Click
        Case 3
            Call mnuEditDelete_Click
    End Select
End Sub

Private Sub mnuShortMenu2_Click(Index As Integer)
    Select Case Index
        Case 1
            Call mnuEditDisease_Click
        Case 2
            Call mnuEditModify_Click
        Case 3
            Call mnuEditDelete_Click
    End Select
End Sub

Private Sub mnuViewAll_Click()
    mnuViewAll.Checked = Not mnuViewAll.Checked
    Call FillList
End Sub

Private Sub mnuViewColumn_Click()
    If zlControl.LvwSelectColumns(lvwMain, mstr����) = True Then
        '���б仯��Ҫ����ˢ��
        Call FillList
    End If
End Sub

Private Sub mnuViewFind_Click()
    frmIllFind.ShowFind mstr�������, mnuViewStop.Checked
End Sub

Private Sub mnuViewOrder_Click()
'�������Ƿ����¼������������еģ���Ҫ�Ƿ�ֹ���Ż�˳��������
    Dim nodTemp As Node
    
    mlng��� = 1
    
    If tvwMain_S.SelectedItem Is Nothing Then Exit Sub
    
    MousePointer = vbHourglass
    Set nodTemp = CheckOrder(tvwMain_S.SelectedItem.Root)
    If nodTemp Is Nothing Then
        MsgBox "�����ϣ������ȷ���С�", vbInformation, gstrSysName
    Else
        nodTemp.Selected = True
        nodTemp.EnsureVisible
        Call FillList
        MsgBox "�÷������ȷ���Ӧ����" & mlng��� & "�����޸ġ�", vbExclamation, gstrSysName
    End If
    MousePointer = 0
    
End Sub

Private Function CheckOrder(ByVal nod As Node) As Node
    Dim lngTemp As Long
    Dim nodTemp As Node
    
    '���ڵ㱾��
    lngTemp = Mid(nod.Text, 2, InStr(nod.Text, "��") - 2)
    If lngTemp <> mlng��� Then
        Set CheckOrder = nod
        Exit Function
    End If
    
    mlng��� = mlng��� + 1
    '�ݹ������ӽڵ�
    Set nod = nod.Child
    Do Until nod Is Nothing
        Set nodTemp = CheckOrder(nod)
        
        '����з���ֵ���Ǳ�ʾ�Ѿ�������
        If Not nodTemp Is Nothing Then
            Set CheckOrder = nodTemp
            Exit Function
        End If
        Set nod = nod.Next
    Loop
    
End Function

Private Sub mnuViewRefresh_Click()
    Call FillTree
End Sub

Private Sub mnuViewIcon_Click(Index As Integer)
    Dim i As Integer
    
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
    Next
    mnuViewIcon(Index).Checked = True
    lvwMain.View = Index
End Sub

Private Sub mnuViewStop_Click()
    mnuViewStop.Checked = Not mnuViewStop.Checked
    Call mnuViewRefresh_Click
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    CoolBar1.Visible = mnuViewToolButton.Checked
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

Private Sub picsplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        msngStartX = X
        picSplit.Tag = "���ƶ�"
    End If
End Sub

Private Sub picsplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngTemp As Single
    
    If picSplit.Tag = "���ƶ�" Then
        sngTemp = picSplit.Left + X - msngStartX
        
        If sngTemp > 1500 And ScaleWidth - sngTemp > 1500 Then
            tvwMain_S.Width = sngTemp - ScaleLeft
            picSplit.Left = sngTemp
            lvwMain.Left = sngTemp + picSplit.Width
            lvwMain.Width = ScaleWidth - lvwMain.Left
        End If
    End If
End Sub

Private Sub picSplit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picSplit.Tag = "" '�ı��־
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Preview"
            mnuFilePrintView_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Class"
            mnuEditClass_Click
        Case "Disease"
            mnuEditDisease_Click
        Case "Modify"
            mnuEditModify_Click
        Case "Delete"
            mnuEditDelete_Click
        Case "ͣ��"
            mnuEditStop_Click
        Case "����"
            mnuEditStart_Click
        Case "View"
            mnuViewIcon(lvwMain.View).Checked = False
            If lvwMain.View = 3 Then
                mnuViewIcon(0).Checked = True
                lvwMain.View = 0
            Else
                mnuViewIcon(lvwMain.View + 1).Checked = True
                lvwMain.View = lvwMain.View + 1
            End If
        Case "Find"
            Call mnuViewFind_Click
        Case "Help"
            Call mnuHelpHelp_Click
        Case "Quit"
            Call mnuFileExit_Click
    End Select
End Sub

Private Sub Toolbar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool
    End If
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Form_Resize
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim i As Integer
    
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
    Next
    mnuViewIcon(ButtonMenu.Index - 1).Checked = True
    lvwMain.View = ButtonMenu.Index - 1
End Sub

Private Sub tvwMain_S_GotFocus()
    Call SetMenu
End Sub

Private Sub tvwMain_S_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If mnuEdit.Visible = False Then Exit Sub
        mnuShortMenu1(1).Enabled = mnuEditClass.Enabled
        mnuShortMenu1(2).Enabled = mnuEditModify.Enabled
        mnuShortMenu1(3).Enabled = mnuEditDelete.Enabled
        PopupMenu mnuShort1
    End If
End Sub

Private Sub tvwMain_S_NodeClick(ByVal Node As MSComctlLib.Node)
    If mstrNodeKey = Node.Key Then Exit Sub
    
    FillList
End Sub

Private Sub subPrint(bytMode As Byte)
'����:���д�ӡ,Ԥ���������EXCEL
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As Object
            
    Set objPrint = New zlPrintLvw
    objPrint.Title.Text = "�����"
    Set objPrint.Body.objData = lvwMain
    objPrint.BelowAppItems.Add "��ӡ�ˣ�" & gstrUserName
    objPrint.BelowAppItems.Add "��ӡʱ�䣺" & Format(zldatabase.Currentdate, "yyyy��MM��dd��")
    If bytMode = 1 Then
        Select Case zlPrintAsk(objPrint)
        Case 1
            zlPrintOrViewLvw objPrint, 1
        Case 2
            zlPrintOrViewLvw objPrint, 2
        Case 3
            zlPrintOrViewLvw objPrint, 3
        Case Else
        End Select
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If

End Sub

Private Sub Fill���()
'���ܣ�װ�뼲���������
    Dim rsTemp As New ADODB.Recordset
    
    mstrTypeText = ""
    
    On Error GoTo errHandle
    gstrSQL = ""
    If gbln������ҽ = False Or gblnʹ����ҽ = False Then
        gstrSQL = " where ����<>'B' and ����<>'Z'"
    End If
    gstrSQL = "select ����,���,�Ƿ���� from ����������� " & gstrSQL & " order by ���ȼ�"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption)

    cmbType.Clear
    Do Until rsTemp.EOF
        cmbType.AddItem rsTemp("����") & ". " & rsTemp("���")
        cmbType.ItemData(cmbType.NewIndex) = rsTemp("�Ƿ����")
        rsTemp.MoveNext
    Loop
    
    cmbType.ListIndex = 0
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub FillTree()
'�����ǰ�����������ô�ѷ���װ�뵽����
    Dim rsTemp As New ADODB.Recordset
    Dim nodTemp As Node
    Dim strKey As String
    Dim strTemp As String
    
    If Not tvwMain_S.SelectedItem Is Nothing Then
        strKey = tvwMain_S.SelectedItem.Key
    End If
    
    mstrTypeText = cmbType.Text
    mstr������� = Left(mstrTypeText, 1)
    
    rsTemp.CursorLocation = adUseClient

    On Error GoTo errHandle
    tvwMain_S.Nodes.Clear
    If tvwMain_S.Visible = True Then
        'ֻ���������ı���
        gstrSQL = "select ID,�ϼ�ID,���,����, ����ʱ�� from ����������� where ���=[1] " & vbNewLine & _
            IIF(mnuViewStop.Checked, "", " And (����ʱ�� is null or ����ʱ��>=to_date('3000-01-01','yyyy-mm-dd'))") & _
            " Start With �ϼ�ID is null connect by prior id=�ϼ�ID order by level,���"

        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstr�������)
                
        Do Until rsTemp.EOF
            strTemp = IIF(Format(rsTemp!����ʱ�� & "", "YYYY-MM-DD") = "3000-01-01", "", Nvl(rsTemp!����ʱ��))
            
            If IsNull(rsTemp("�ϼ�ID")) Then
                Set nodTemp = tvwMain_S.Nodes.Add(, , "K" & rsTemp("ID"), "��" & rsTemp("���") & "��" & Trim(rsTemp("����")), "Root", "Root")
            Else
                Set nodTemp = tvwMain_S.Nodes.Add("K" & rsTemp("�ϼ�ID"), tvwChild, "K" & rsTemp("ID"), "��" & rsTemp("���") & "��" & Trim(rsTemp("����")), "Root", "Root")
            End If
            If strTemp <> "" Then
                nodTemp.ForeColor = vbRed
                nodTemp.Tag = strTemp
            End If
            rsTemp.MoveNext
        Loop
    End If
    On Error Resume Next
    Set nodTemp = tvwMain_S.Nodes(strKey)
    If Err <> 0 Then
        Set nodTemp = tvwMain_S.Nodes(1)
        nodTemp.Selected = True
        nodTemp.Expanded = True
    Else
        Err.Clear
        nodTemp.Selected = True
        nodTemp.Expanded = True
        nodTemp.EnsureVisible
    End If
    Call FillList
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub FillList()
'����:����ListView�е�����
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim i As Integer
    Dim lst As ListItem
    Dim strKey As String
    
    On Error GoTo errHandleList
    
    If tvwMain_S.SelectedItem Is Nothing And tvwMain_S.Visible = True Then
        lvwMain.ListItems.Clear
        Call SetMenu
        Exit Sub
    End If
    If Not lvwMain.SelectedItem Is Nothing Then
        '����ԭ�м�ֵ
        strKey = lvwMain.SelectedItem.Key
    End If
    
    rsTemp.CursorLocation = adUseClient
    
    If tvwMain_S.Visible = False Then
        'û�����¼��Ĺ�ϵ
        gstrSQL = " A.���=[1] "
    Else
        mstrNodeKey = tvwMain_S.SelectedItem.Key '��¼��ǰ���ʵĽڵ�
        If mnuViewAll.Checked = True Then
            gstrSQL = " A.����ID in  " & _
                "(select ID from ����������� start with id=[1] " & _
                " connect by prior id=�ϼ�ID)"
        Else
            gstrSQL = " A.����ID=[1] "
        End If
    End If
        
    gstrSQL = "" & _
    "   Select A.ID,A.����,����,A.����,A.���� as ƴ����,A.�����,A.˵��,A.�Ա�����,A.��Ч���� as ������Ч,A.��������," & _
    "          A.ͳ����,decode(A.����,1,'¼��') ������Ϣ,to_char(A.����ʱ��,'yyyy-mm-dd') as  ����ʱ��, " & _
    "          to_char(A.����ʱ��,'yyyy-mm-dd') as ����ʱ��" & _
    "   From ��������Ŀ¼ A  " & _
    "   Where " & gstrSQL & IIF(mnuViewStop.Checked, "", " and (a.����ʱ�� is null or a.����ʱ��>=to_date('3000-01-01','yyyy-mm-dd'))")
    
    If tvwMain_S.Visible = False Then
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstr�������)
    Else
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Mid(tvwMain_S.SelectedItem.Key, 2)))
    End If
    Dim strIco As String
    
    With lvwMain.ListItems
        .Clear
        Do Until rsTemp.EOF
            '�ó���ȷ��ͼ��
            '��ӽڵ�
            strTemp = IIF(Nvl(rsTemp!����ʱ��) = "3000-01-01", "", Nvl(rsTemp!����ʱ��))
            strIco = IIF(strTemp <> "", "Stop", "Item")
            
       
            Set lst = .Add(, "K" & rsTemp("id"), rsTemp("����"), strIco, strIco)
            If strTemp <> "" Then
                lst.ForeColor = vbRed
            Else
                lst.ForeColor = lvwMain.ForeColor
            End If
            Dim varValue As Variant
            '����ListView�����������ݿ�ȡ��
            For i = 2 To lvwMain.ColumnHeaders.Count
                varValue = rsTemp(lvwMain.ColumnHeaders(i).Text).value
                If lvwMain.ColumnHeaders(i).Text = "����ʱ��" Then
                    If Nvl(varValue) = "3000-01-01" Then
                        lst.SubItems(i - 1) = ""
                    Else
                        lst.SubItems(i - 1) = IIF(IsNull(varValue), "", varValue)
                        
                    End If
                Else
                    lst.SubItems(i - 1) = IIF(IsNull(varValue), "", varValue)
                End If
                If strTemp <> "" Then lst.ListSubItems(i - 1).ForeColor = vbRed
            Next
            rsTemp.MoveNext
        Loop
    End With
    
    If lvwMain.ListItems.Count > 0 Then
        On Error Resume Next
        Set lst = lvwMain.ListItems(strKey)
        If Err <> 0 Then
            Set lst = lvwMain.ListItems(1)
            lst.Selected = True
            lst.EnsureVisible
        Else
            Err.Clear
            lst.Selected = True
            lst.EnsureVisible
        End If
    End If
    Call SetMenu
    
    Exit Sub
errHandleList:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Sub

Private Sub Ȩ�޿���()
'����:�����е��û�Ȩ�޲���,��ʹһЩ�˵����ť���ɼ�
    gbln������ҽ = InStr(mstrPrivs, "��ҽ") > 0
    If InStr(mstrPrivs, "��ɾ��") = 0 Then
        mnuEdit.Visible = False
        mnuEditClass.Visible = False
        mnuEditDisease.Visible = False
        mnuEditModify.Visible = False
        mnuFileExport.Visible = False
        mnuFileImport.Visible = False
        mnuShortMenu2(1).Visible = False
        mnuShortMenu2(2).Visible = False
        mnuShortMenu2(3).Visible = False
        mnuShortLine.Visible = False
        
        Toolbar1.Buttons("Split0").Visible = False
        Toolbar1.Buttons("Class").Visible = False
        Toolbar1.Buttons("Disease").Visible = False
        Toolbar1.Buttons("Modify").Visible = False
        Toolbar1.Buttons("Delete").Visible = False
        Toolbar1.Buttons("Splits").Visible = False
        Toolbar1.Buttons("ͣ��").Visible = False
        Toolbar1.Buttons("����").Visible = False
        
    End If
      
End Sub

Public Sub SetMenu()
'����:�����޸ĺ�ɾ����ť����Чֵ
'����:blnEnabled ��Чֵ
'    Dim blnNew As Boolean
    Dim blnModify As Boolean
    Dim blnStop As Boolean
    
    blnStop = False
    blnModify = True
    If ActiveControl Is tvwMain_S Then
        If tvwMain_S.SelectedItem Is Nothing Then
            blnModify = False
            stbThis.Panels(2).Text = "�������÷��ࡣ"
        Else
            stbThis.Panels(2).Text = "��ǰ���๲��" & tvwMain_S.SelectedItem.Children & "�����࣬" & lvwMain.ListItems.Count & "���������롣"
        End If
        If Not tvwMain_S.SelectedItem Is Nothing Then
            blnStop = tvwMain_S.SelectedItem.ForeColor = vbRed
        End If
    Else
        If lvwMain.SelectedItem Is Nothing Or lvwMain.ListItems.Count = 0 Then
            blnModify = False
        End If
        stbThis.Panels(2).Text = "��ǰ���๲��" & lvwMain.ListItems.Count & "���������롣"
        If Not lvwMain.SelectedItem Is Nothing Then
            blnStop = lvwMain.SelectedItem.SubItems(lvwMain.ColumnHeaders("_����ʱ��").Index - 1) <> ""
        End If
    End If
    
    Toolbar1.Buttons("ͣ��").Enabled = Not blnStop
    Toolbar1.Buttons("����").Enabled = blnStop
    
    mnuEditStart.Enabled = blnStop
    mnuEditStop.Enabled = Not blnStop
    
    'ֻ�������б�ɼ�ʱ���ſ����ӷ���
    Toolbar1.Buttons("Class").Enabled = tvwMain_S.Visible
    mnuEditClass.Enabled = tvwMain_S.Visible
    mnuViewAll.Enabled = tvwMain_S.Visible
    mnuViewOrder.Enabled = tvwMain_S.Visible
    
    mnuEditDisease.Enabled = (Not tvwMain_S.Visible) Or (Not tvwMain_S.SelectedItem Is Nothing)
    Toolbar1.Buttons("Disease").Enabled = mnuEditDisease.Enabled
    
    Toolbar1.Buttons("Modify").Enabled = blnModify
    Toolbar1.Buttons("Delete").Enabled = blnModify
    mnuEditDelete.Enabled = blnModify
    mnuEditModify.Enabled = blnModify

    EnablePrint lvwMain.ListItems.Count > 0
End Sub

Private Sub EnablePrint(ByVal blnEnabled As Boolean)
'����:���ô�ӡ��Ԥ����ť����Чֵ
'����:blnEnabled ��Чֵ
    Toolbar1.Buttons("Print").Enabled = blnEnabled
    Toolbar1.Buttons("Preview").Enabled = blnEnabled
    mnuFilePrintView.Enabled = blnEnabled
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

Private Sub StopAndResume(ByVal blnStop As Boolean)
    '--------------------------------------------------------------------------------------
    '����:ͣ�û���������
    '����:blnStop-�Ƿ�ͣ��
    '����:
    '����:���˺�
    '����:11689
    '�޸�:2007/12/28
    '--------------------------------------------------------------------------------------
    
    Dim lng����ID As Long, lng����id As Long
    Dim strSQL As String, intIndex As Integer
    Dim i As Integer
    Dim ReMoveRow As Long
    Dim nodTemp As Node
    Dim str���� As String
    Dim strDate As String
    
    
    If ActiveControl Is tvwMain_S And tvwMain_S.Visible = True Then
        '�޸ķ���
        Set nodTemp = tvwMain_S.SelectedItem
        If nodTemp Is Nothing Then
            Exit Sub
        Else
            With tvwMain_S.SelectedItem
                If AllowContinue = False Then Exit Sub
                If MsgBox("���Ƿ����Ҫ" & IIF(blnStop, "ͣ��", "����") & "��" & tvwMain_S.SelectedItem.Text & "�������������з�����Ŀ�Լ�������Ŀ�����еļ���������", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
                lng����id = Val(Mid(tvwMain_S.SelectedItem.Key, 2))
                If lng����id <= 0 Then Exit Sub
                Err = 0: On Error GoTo ErrHand:
                If blnStop Then
                    strSQL = "Zl_�����������_STOP(" & lng����id & ")"
                Else
                    If nodTemp.Tag <> "" Then
                        strDate = "To_Date('" & nodTemp.Tag & "','YYYY-MM-DD HH24:MI:SS')"
                        strSQL = "Zl_�����������_REUSE(" & lng����id & "," & strDate & ")"
                    Else
                        strSQL = "Zl_�����������_REUSE(" & lng����id & ")"
                    End If
                    
                End If
                zldatabase.ExecuteProcedure strSQL, Me.Caption
            End With
            'ˢ��
            Call FillTree
        End If
    Else
        '�޸ļ���
        If lvwMain.SelectedItem Is Nothing Then Exit Sub
        If MsgBox("���Ƿ����Ҫ" & IIF(blnStop, "ͣ��", "����") & "��" & lvwMain.SelectedItem.Text & "���ļ�����", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
        
        lng����ID = Val(Mid(lvwMain.SelectedItem.Key, 2))
        If lng����ID <= 0 Then Exit Sub
        
        Err = 0: On Error GoTo ErrHand:
        If blnStop Then
            strSQL = "Zl_��������Ŀ¼_STOP(" & lng����ID & ")"
        Else
            strSQL = "Zl_��������Ŀ¼_REUSE(" & lng����ID & ")"
        End If
        zldatabase.ExecuteProcedure strSQL, Me.Caption
        
        With lvwMain.SelectedItem
            .Icon = IIF(blnStop, "Stop", "Item")
            .SmallIcon = IIF(blnStop, "Stop", "Item")
            .ForeColor = IIF(blnStop, vbRed, &H80000008)
        End With
        For i = 2 To lvwMain.ColumnHeaders.Count
             lvwMain.SelectedItem.ListSubItems(i - 1).ForeColor = IIF(blnStop, vbRed, &H80000008)
        Next
        If mnuViewStop.Checked Then
            If Not blnStop Then
                lvwMain.SelectedItem.SubItems(lvwMain.ColumnHeaders("_����ʱ��").Index - 1) = ""
            Else
                lvwMain.SelectedItem.SubItems(lvwMain.ColumnHeaders("_����ʱ��").Index - 1) = Format(zldatabase.Currentdate, "yyyy-mm-dd")
            End If
            Call SetMenu
            Exit Sub
        End If
        
        If blnStop = False Then
            lvwMain.SelectedItem.SubItems(lvwMain.ColumnHeaders("_����ʱ��").Index - 1) = ""
            Call SetMenu
            Exit Sub
        End If
        Me.MousePointer = 0
        With lvwMain
            intIndex = .SelectedItem.Index
            .ListItems.Remove .SelectedItem.Key
            If .ListItems.Count > 0 Then
                intIndex = IIF(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
                .ListItems(intIndex).Selected = True
                .ListItems(intIndex).EnsureVisible
            End If
            Call SetMenu
        End With
    End If
    

    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncCreateSQL(ByVal strFilePath As String)
'����:���ɰ�װ�Ų�
'����:strFilePath-�����ļ���λ��
    Dim rsType As New ADODB.Recordset
    Dim rsContent As New ADODB.Recordset
    Dim strTitle As String, strTemp As String, strValue As String
    Dim strName As String, strType As String
    
    Dim lngId As Long, i As Long
    Dim colTemp As Collection
    Dim colSQL As Collection
    Dim blnOver As Boolean
    Dim objFile As New FileSystemObject
    Dim strFileName As String
    
    On Error GoTo errH
    'ֻ���������������Ϊ D-ICD-10;Y-�����ж�;M-������̬ѧ;S-��������
    gstrSQL = "Select a.Id, a.�ϼ�id, a.���, a.����,a.����, a.���, a.���뷶Χ, a.�Ƿ���,NULL AS ID_TEMP, NULL AS �ϼ�ID_TEMP " & vbNewLine & _
                "From ����������� A, ����������� B" & vbNewLine & _
                "Where a.��� = b.���� And (a.����ʱ�� Is Null Or Trunc(a.����ʱ��) = To_Date('3000-01-01', 'YYYY-MM-DD')) And A.��� In ('D','Y','M','S')" & vbNewLine & _
                "Order By b.���ȼ�,a.���"
    Call zldatabase.OpenRecordset(rsType, gstrSQL, Me.Caption, adOpenStatic, adLockOptimistic)
    
    gstrSQL = "Select a.Id, NULL AS ID_TEMP, a.����, a.���, a.����, a.ͳ����, a.����, a.����, a.�����, a.˵��, a.�Ա�����, a.��Ч����, a.��������, a.����, a.����id, NULL As ����ID_TEMP, a.���÷�Χ, a.���" & vbNewLine & _
        "From ��������Ŀ¼ A, ����������� B" & vbNewLine & _
        "Where a.��� = b.���� And (a.����ʱ�� Is Null Or Trunc(a.����ʱ��) = To_Date('3000-01-01', 'YYYY-MM-DD')) And A.��� In ('D','Y','M','S') " & vbNewLine & _
        "Order By b.���ȼ�, a.����, a.��� "
    Call zldatabase.OpenRecordset(rsContent, gstrSQL, Me.Caption, adOpenStatic, adLockOptimistic)
    
    'ID����
    lngId = 1: Set colTemp = New Collection
    For i = 1 To rsType.RecordCount
        colTemp.Add lngId, "_" & rsType!ID
        lngId = lngId + 1
        rsType.MoveNext
    Next
    
    rsType.Filter = ""
    For i = 1 To rsType.RecordCount
        rsType!ID_TEMP = colTemp("_" & rsType!ID)
        If Nvl(rsType!�ϼ�id, 0) <> 0 Then rsType!�ϼ�ID_TEMP = colTemp("_" & rsType!�ϼ�id)
        rsType.MoveNext
    Next
    
    rsContent.Filter = ""
    lngId = 1
    For i = 1 To rsContent.RecordCount
        rsContent!ID_TEMP = lngId
        rsContent!����ID_TEMP = colTemp("_" & rsContent!����id)
        lngId = lngId + 1
        rsContent.MoveNext
    Next
    
    Set colSQL = New Collection
    Set colTemp = New Collection
    
    With rsType
        rsType.Filter = ""
        strType = ""
        strTitle = "Insert Into �����������(ID, �ϼ�id, ���, ���, ����, ����, ���뷶Χ, �Ƿ���) " & vbCrLf
        For i = 1 To .RecordCount
            strName = FuncGetStr(!����)
            strTemp = "Select " & !ID_TEMP & "," & IIF(Val(!�ϼ�id & "") = 0, "Null", !�ϼ�ID_TEMP) & ",'" & Trim(!���) & "'," & !��� & ",'" & strName & "','" & !���� & "','" & FuncGetStr(!���뷶Χ & "") & "'," & Val(!�Ƿ��� & "") & " From Dual UNION ALL" & vbCrLf
            If Len(strTitle & strValue & strTemp) > 100000 Or (!��� & "" <> strType And strType <> "") Then
                colSQL.Add "--���=" & strType  '���һ�п���
                strValue = Mid(strValue, 1, InStrRev(strValue, "UNION ALL") - 1) & ";"
                colSQL.Add strTitle & strValue
                strValue = strTemp
                blnOver = True
            Else
                blnOver = False
                strValue = strValue & strTemp
            End If
            strType = !��� & ""
            .MoveNext
            If .EOF Then
                If Not blnOver Then
                    colSQL.Add "--���=" & strType  '���һ�п���
                    strValue = Mid(strValue, 1, InStrRev(strValue, "UNION ALL") - 1) & ";"
                    colSQL.Add strTitle & strValue
                    Exit For
                End If
            End If
        Next
    End With
    strFileName = strFilePath & "��������.SQL"
    If objFile.FileExists(strFileName) Then objFile.DeleteFile strFileName, True
    SaveLog strFileName, "--�����������", "-1"
    For i = 1 To colSQL.Count
        SaveLog strFileName, colSQL(i), "-1"
    Next
    
    Set colSQL = New Collection
    With rsContent
        .Filter = "": strValue = "": strType = ""
        strTitle = "Insert Into ��������Ŀ¼ (ID, ����id, ���, ����, ���, ����, ����, ����, �����, ˵��, �Ա�����, ��Ч����, ��������, ����, ���÷�Χ)" & vbCrLf
        For i = 1 To .RecordCount
            strName = FuncGetStr(!����)
            strTemp = "Select " & !ID_TEMP & "," & !����ID_TEMP & ",'" & !��� & "','" & !���� & "'," & !��� & ",'" & !���� & "','" & strName & "','" & !���� & "','" & !����� & "','" & !˵�� & "','" & !�Ա����� & "','" & _
                    !��Ч���� & "','" & !�������� & "','" & !���� & "','" & !���÷�Χ & "' From Dual UNION ALL" & vbCrLf
      
            If Len(strTitle & strValue & strTemp) > 100000 Or (!��� & "" <> strType And strType <> "") Then
                colSQL.Add "--���=" & strType  '���һ�п���
                strValue = Mid(strValue, 1, InStrRev(strValue, "UNION ALL") - 1) & ";"
                colSQL.Add strTitle & strValue
                strValue = strTemp
                blnOver = True
            Else
                blnOver = False
                strValue = strValue & strTemp
            End If
            strType = !��� & ""
            .MoveNext
            If .EOF Then
                If Not blnOver Then
                    colSQL.Add "--���=" & strType  '���һ�п���
                    strValue = Mid(strValue, 1, InStrRev(strValue, "UNION ALL") - 1) & ";"
                    colSQL.Add strTitle & strValue
                    Exit For
                End If
            End If
        Next
    End With
    SaveLog strFileName, "--��������Ŀ¼", "-1"
    For i = 1 To colSQL.Count
        SaveLog strFileName, colSQL(i), "-1"
    Next
    MsgBox "�����ɹ�,�ļ�λ��:" & vbCrLf & strFileName, vbInformation + vbOKOnly, Me.Caption
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



