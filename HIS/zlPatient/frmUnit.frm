VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frmUnit 
   Caption         =   "��Լ��λ����"
   ClientHeight    =   4980
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   6645
   Icon            =   "frmUnit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin MSComctlLib.ImageList ils32 
      Left            =   2580
      Top             =   1020
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
            Picture         =   "frmUnit.frx":0442
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":089A
            Key             =   "ItemNo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":0CEE
            Key             =   "Write"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   2700
      Top             =   1680
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
            Picture         =   "frmUnit.frx":1146
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":159E
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":19F6
            Key             =   "ItemNo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":1E4A
            Key             =   "No"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":22A2
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
      ScaleWidth      =   30
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1590
      Width           =   30
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "����"
         Text            =   "����"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "����"
         Text            =   "����"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "����"
         Text            =   "����"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "��ַ"
         Text            =   "��ַ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "�绰"
         Text            =   "�绰"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "��������"
         Text            =   "��������"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "�ʺ�"
         Text            =   "�ʺ�"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Key             =   "��ϵ��"
         Text            =   "��ϵ��"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Key             =   "վ��"
         Text            =   "Ժ��"
         Object.Width           =   1411
      EndProperty
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
      Indentation     =   18
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList Ilsrw 
      Left            =   4470
      Top             =   60
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
            Picture         =   "frmUnit.frx":26FA
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":291A
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":2B3A
            Key             =   "Parent"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":2D56
            Key             =   "Child"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":2F72
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":3192
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":33B2
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":35D2
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":37F2
            Key             =   "View"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":3A12
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":3C32
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Ilscolor 
      Left            =   3270
      Top             =   210
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
            Picture         =   "frmUnit.frx":3E52
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":4072
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":4292
            Key             =   "Parent"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":44AE
            Key             =   "Child"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":46CA
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":48EA
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":4B0A
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":4D2A
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":4F4A
            Key             =   "View"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":516A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":538A
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
      _Version        =   "6.7.8988"
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
      SimpleText      =   $"frmUnit.frx":55AA
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmUnit.frx":55F1
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
      End
      Begin VB.Menu mnuEditStop 
         Caption         =   "ͣ��(&T)"
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
         Caption         =   "��ʾͣ�õ�λ(&P)"
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
Attribute VB_Name = "frmUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Dim msngStartX As Single         '�ƶ�ǰ����λ��
Dim mblnItem As Boolean         'Ϊ���ʾ������ListViewĳһ����
Dim mintColumn As Integer
Dim mblnLoad As Boolean
Dim mstrKey As String
Private mstrPrivs As String
Private mlngModul As Long
Private Const mstrLvw As String = "����,1300,0,1;����,800,0,2;����,900,0,0;��ַ,1440,0,0;�绰,1440,0,0;��������,1440,0,0;�ʺ�,1440,0,0;��ϵ��,1440,0,0;����ʱ��,1100,0,0;����ʱ��,1100,0,0;��������,2000,0,0;Ժ��,800,0,0"

Private Sub Form_Activate()
    If mblnLoad = True Then
        Call Ȩ�޿���
        Call Form_Resize 'Ϊ��ʹCoolBar����Ӧ�߶�
        FillTree
    End If
    mblnLoad = False
End Sub

Private Sub Form_Load()
    mlngModul = glngModul
    mstrPrivs = gstrPrivs
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    mblnLoad = True
    '���������ɾ����ListView�������
    lvwMain.Tag = "�ɱ仯��"
    RestoreWinState Me, App.ProductName
    If lvwMain.ColumnHeaders(9).Text = "վ��" Then
        lvwMain.ColumnHeaders(9).Text = "Ժ��"
    End If
    mnuViewShowAll.Checked = zlDatabase.GetPara("��ʾ�����¼�", glngSys, mlngModul, 0) <> "0"
    mnuViewShowStop.Checked = zlDatabase.GetPara("��ʾͣ�õ�λ", glngSys, mlngModul, 0) <> "0"
    '����LvwMain��ʾ���ö�Ӧ�˵�
     mnuViewIcon_Click lvwMain.View
End Sub

Private Sub Form_Resize()
    Dim sngTop As Single, sngBottom As Single
    sngTop = IIf(CoolBar1.Visible, CoolBar1.Top + CoolBar1.Height, 0)
    sngBottom = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
    
    tvwMain_S.Top = sngTop
    tvwMain_S.Height = IIf(sngBottom - tvwMain_S.Top > 0, sngBottom - tvwMain_S.Top, 0)
    tvwMain_S.Left = 0
    
    picSplit.Top = sngTop
    picSplit.Height = IIf(sngBottom - picSplit.Top > 0, sngBottom - picSplit.Top, 0)
    picSplit.Left = tvwMain_S.Left + tvwMain_S.Width
    
    lvwMain.Left = picSplit.Left + picSplit.Width
    lvwMain.Top = sngTop
    lvwMain.Height = IIf(sngBottom - lvwMain.Top > 0, sngBottom - lvwMain.Top, 0)
    If Me.ScaleWidth - lvwMain.Left > 0 Then lvwMain.Width = Me.ScaleWidth - lvwMain.Left
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrKey = ""
    zlDatabase.SetPara "��ʾ�����¼�", IIf(mnuViewShowAll.Checked, 1, 0), glngSys, mlngModul
    zlDatabase.SetPara "��ʾͣ�õ�λ", IIf(mnuViewShowStop.Checked, 1, 0), glngSys, mlngModul
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
    If mblnItem = True And mnuEditModify.Visible Then mnuEditModify_Click
End Sub

Private Sub lvwMain_GotFocus()
    With lvwMain
        stbThis.Panels(2).Text = "��λ�б��й���ʾ��" & .ListItems.Count & "����Լ��λ��"
    End With
    Call SetMenu
End Sub

Private Sub lvwMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mblnItem = True
    SetMenu
    stbThis.Panels(2).Text = "��λ�б��й���ʾ��" & lvwMain.ListItems.Count & "����Լ��λ��"
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
           blnReturn = frmUnitEdit.�༭��λ("��", "", "", , True)
        Else
            i = InStr(.Text, "��")
            str���� = Mid(.Text, 2, i - 2)
            str���� = Mid(.Text, i + 1)
            blnReturn = frmUnitEdit.�༭��λ(str����, Mid(.Key, 2), str����, , True)
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
           blnReturn = frmUnitEdit.�༭��λ("��", "", "", , False)
        Else
            i = InStr(.Text, "��")
            str���� = Mid(.Text, 2, i - 2)
            str���� = Mid(.Text, i + 1)
           blnReturn = frmUnitEdit.�༭��λ(str����, Mid(.Key, 2), str����, , False)
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
    
    On Error GoTo errHandle
    With tvwMain_S.SelectedItem
        If .Key = "Root" Then
            str������ = ""
            intNew = GetDownCodeLength("", "��Լ��λ")
            intChild = GetLocalCodeLength("", "��Լ��λ")
        Else
            str������ = Mid(.Text, 2, InStr(.Text, "��") - 2)
            intNew = GetDownCodeLength(Mid(.Key, 2), "��Լ��λ")
            intChild = GetLocalCodeLength(Mid(.Key, 2), "��Լ��λ")
        End If
        If intNew = 0 Or intChild = 0 Then Exit Sub
        If intNew = 10 Then
            MsgBox "�����ټӳ����룬ĳһ���¼��Ѿ������˳��ȡ�", vbExclamation, gstrSysName
            Exit Sub
        End If
        
        intNew = frmCount.GetLength(intChild, 10 - (intNew - intChild))
        If intNew = 0 Then Exit Sub
        strTemp = str������ & String(intNew - intChild, "0")
        
        If .Key = "Root" Then
            gstrSQL = "zl_��Լ��λ_EXPAND('" & strTemp & "'," & Len(str������) + 1 & ",0)"
        Else
            gstrSQL = "zl_��Լ��λ_EXPAND('" & strTemp & "'," & Len(str������) + 1 & "," & Mid(.Key, 2) & ")"
        End If
'        Call SQLTest(App.ProductName, Me.Caption, gstrSQL)
'        gcnOracle.Execute gstrSQL, , adCmdStoredProc
'        Call SQLTest
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
        FillTree
    End With
    Exit Sub
errHandle:
    If errCenter() = 1 Then Resume
    Resume
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
            If .Parent.Key = "Root" Then
               blnReturn = frmUnitEdit.�༭��λ("��", "", "", Mid(.Key, 2))
            Else
                i = InStr(.Parent.Text, "��")
                str���� = Mid(.Parent.Text, 2, i - 2)
                str���� = Mid(.Parent.Text, i + 1)
                blnReturn = frmUnitEdit.�༭��λ(str����, Mid(.Parent.Key, 2), str����, Mid(.Key, 2))
            End If
        Else
            If .Key = "Root" Then
                blnReturn = frmUnitEdit.�༭��λ("��", "", "", Mid(lvwMain.SelectedItem.Key, 2))
            Else
                i = InStr(.Text, "��")
                str���� = Mid(.Text, 2, i - 2)
                str���� = Mid(.Text, i + 1)
                blnReturn = frmUnitEdit.�༭��λ(str����, Mid(.Key, 2), str����, Mid(lvwMain.SelectedItem.Key, 2))
            End If
        End If
    End With
    If blnReturn = True Then
        FillTree
    End If
End Sub

Private Sub mnuEditDelete_Click()
    On Error GoTo errHandle
    
    If ActiveControl Is tvwMain_S Then
        If MsgBox("ɾ������ͬʱҲ��ɾ����������Ŀ��" & vbCrLf & "��ȷ��Ҫɾ������Ϊ��" & tvwMain_S.SelectedItem.Text & "���ķ�����Ŀ��", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        gstrSQL = "zl_��Լ��λ_delete(" & Mid(tvwMain_S.SelectedItem.Key, 2) & ")"
    Else
        If MsgBox("��ȷ��Ҫɾ������Ϊ��" & lvwMain.SelectedItem.Text & "���ĺ�Լ��λ��", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        gstrSQL = "zl_��Լ��λ_delete(" & Mid(lvwMain.SelectedItem.Key, 2) & ")"
    End If
    Me.MousePointer = 11
    
'    Call SQLTest(App.ProductName, Me.Caption, gstrSQL)
'    gcnOracle.Execute gstrSQL, , adCmdStoredProc
'    Call SQLTest
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    
    If ActiveControl Is tvwMain_S Then
        FillTree
        Call tvwMain_S_GotFocus
    Else
        FillList tvwMain_S.SelectedItem.Key
        Call lvwMain_GotFocus
    End If
    Me.MousePointer = 0
    Exit Sub
errHandle:
    If errCenter() = 1 Then Resume
    Call SaveErrLog
    Me.MousePointer = 0
End Sub

Private Sub mnuEditStart_Click()
    On Error GoTo errHandle
    Dim strKey As String

    strKey = lvwMain.SelectedItem.Key
    gstrSQL = "zl_��Լ��λ_reuse(" & Mid(lvwMain.SelectedItem.Key, 2) & ")"
    'ִ�����ù���
'    Call SQLTest(App.ProductName, Me.Caption, gstrSQL)
'    gcnOracle.Execute gstrSQL, , adCmdStoredProc
'    Call SQLTest
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    
    '�ı�ͼ�����ɫ
    With lvwMain.SelectedItem
        .Icon = "Item"
        .SmallIcon = "Item"
        .ForeColor = RGB(0, 0, 0)
        
        Dim i As Integer
        For i = 1 To lvwMain.ColumnHeaders.Count - 1
            .ListSubItems(i).ForeColor = RGB(0, 0, 0)
        Next
    End With
    '�ı�״̬���Ͳ˵�
    SetMenu
    Exit Sub
errHandle:
    If errCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditStop_Click()
    On Error GoTo errHandle
    Dim strKey As String

    strKey = lvwMain.SelectedItem.Key
    gstrSQL = "zl_��Լ��λ_stop(" & Mid(lvwMain.SelectedItem.Key, 2) & ")"
    'ִ�����ù���
'    Call SQLTest(App.ProductName, Me.Caption, gstrSQL)
'    gcnOracle.Execute gstrSQL, , adCmdStoredProc
'    Call SQLTest
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    
    '�ı�ͼ�����ɫ
    If mnuViewShowStop.Checked = True Then 'Ҫ��ʾͣ�ò���
        With lvwMain.SelectedItem
            .Icon = "ItemNo"
            .SmallIcon = "ItemNo"
            .ForeColor = RGB(255, 0, 0)
            
            Dim i As Integer
            For i = 1 To lvwMain.ColumnHeaders.Count - 1
                .ListSubItems(i).ForeColor = RGB(255, 0, 0)
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
errHandle:
    If errCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me)
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
    Next
    mnuViewIcon(Index).Checked = True
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
        msngStartX = X
    End If
End Sub

Private Sub picsplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngTemp As Single
    If Button = 1 Then
        sngTemp = picSplit.Left + X - msngStartX
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
ShowHelp App.ProductName, Me.hwnd, Me.Name
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
    stbThis.Panels(2).Text = "����λ������" & lvwMain.ListItems.Count & "���¼���Ŀ"
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
    objPrint.Title.Text = "��Լ��λ"
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
'����:װ���Լ��λ�����з��ൽtvwMain_S
    Dim strTemp As String
    Dim strKey As String
    Dim rs��Լ��λ As New ADODB.Recordset
    
    mstrKey = ""
    
    rs��Լ��λ.CursorLocation = adUseClient
    rs��Լ��λ.CursorType = adOpenKeyset
    rs��Լ��λ.LockType = adLockReadOnly
    If Not tvwMain_S.SelectedItem Is Nothing Then
        strKey = tvwMain_S.SelectedItem.Key
    End If
    
    On Error GoTo errHandle
    gstrSQL = "select ID,�ϼ�ID,����,���� from ��Լ��λ  " & _
        "where ĩ�� <> 1 start with �ϼ�ID is null connect by prior ID =�ϼ�ID"
    Set rs��Լ��λ = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    tvwMain_S.Nodes.Clear
    tvwMain_S.Nodes.Add , , "Root", "���к�Լ��λ", "Root", "Root"
    tvwMain_S.Nodes("Root").Sorted = True
    Do Until rs��Լ��λ.EOF
        
        If IsNull(rs��Լ��λ("�ϼ�id")) Then
            tvwMain_S.Nodes.Add "Root", tvwChild, "C" & rs��Լ��λ("id"), "��" & rs��Լ��λ("����") & "��" & rs��Լ��λ("����"), "Write", "Write"
        Else
            tvwMain_S.Nodes.Add "C" & rs��Լ��λ("�ϼ�id"), tvwChild, "C" & rs��Լ��λ("id"), "��" & rs��Լ��λ("����") & "��" & rs��Լ��λ("����"), "Write", "Write"
        End If
        tvwMain_S.Nodes("C" & rs��Լ��λ("ID")).Sorted = True
        rs��Լ��λ.MoveNext
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
errHandle:
    If errCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub FillList(ByVal str��Լ��λID As String)
'����:װ���Ӧ������ӷ������Ŀ��lvwMain
'����:str��Լ��λID ����ı�ʶ
    Dim rs��Լ��λ As New ADODB.Recordset
    Dim fld As Field
    Dim lst As ListItem
    Dim strKey As String
    Dim strͣ�� As String
    
    If Not lvwMain.SelectedItem Is Nothing Then
        '����ԭ�м�ֵ
        strKey = lvwMain.SelectedItem.Key
    End If
    
    rs��Լ��λ.CursorLocation = adUseClient
    rs��Լ��λ.CursorType = adOpenKeyset
    rs��Լ��λ.LockType = adLockReadOnly
    
    If mnuViewShowStop.Checked = False Then
        strͣ�� = " (����ʱ�� is null or ����ʱ�� = to_date('3000-01-01','YYYY-MM-DD')) and "
    End If
    
    On Error GoTo errHandle
    'by lesfeng 2010-03-08 �����Ż�
    If mnuViewShowAll.Checked = True Then
        gstrSQL = "select A.ID,A.�ϼ�ID,A.����,A.����,A.����,A.��ַ,A.�绰,A.��������,A.�ʺ�,A.��ϵ��,A.����ʱ��,A.����ʱ��,A.վ�� Ժ��,B.���� as �������� from " & _
            "(select ID,�ϼ�ID,����,����,����,��ַ,�绰,��������,�ʺ�,��ϵ��,to_char(����ʱ��,'YYYY-MM-DD') as ����ʱ��,to_char(����ʱ��,'YYYY-MM-DD') as ����ʱ��,վ�� Ժ�� from ��Լ��λ where " & _
            IIf(strͣ�� = "", "", strͣ��) & " ĩ��=1 connect by prior id=�ϼ�id start with  " & _
            IIf(str��Լ��λID = "Root", "�ϼ�ID is null ", "�ϼ�ID = [1] ") & ") A,��Լ��λ B where A.�ϼ�ID=B.ID(+)"
    Else
        gstrSQL = "select ID,�ϼ�ID,����,����,����,��ַ,�绰,��������,�ʺ�,��ϵ��,to_char(����ʱ��,'YYYY-MM-DD') as ����ʱ��,to_char(����ʱ��,'YYYY-MM-DD') as ����ʱ��,'' as ��������,վ�� Ժ�� from ��Լ��λ where " & _
            IIf(strͣ�� = "", "", strͣ��) & " ĩ��=1 and " & IIf(str��Լ��λID = "Root", "�ϼ�ID is null ", "�ϼ�ID = [1]")
    End If
    
    Set rs��Լ��λ = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Mid(str��Լ��λID, 2))
    
    LockWindowUpdate lvwMain.hwnd
    lvwMain.ListItems.Clear
    Do Until rs��Լ��λ.EOF
        If CDate(IIf(IsNull(rs��Լ��λ("����ʱ��")), "3000-01-01", rs��Լ��λ("����ʱ��"))) = CDate("3000-01-01") Then
            Set lst = lvwMain.ListItems.Add(, "C" & rs��Լ��λ("ID"), rs��Լ��λ("����"), "Item", "Item")
        Else
            Set lst = lvwMain.ListItems.Add(, "C" & rs��Լ��λ("ID"), rs��Լ��λ("����"), "ItemNo", "ItemNo")
            lst.ForeColor = RGB(255, 0, 0)
        End If
        
        Dim lngCol  As Long
        Dim varValue As Variant
        '����ListView�����������ݿ�ȡ��
        For lngCol = 2 To lvwMain.ColumnHeaders.Count
            varValue = rs��Լ��λ(lvwMain.ColumnHeaders(lngCol).Text).Value
            lst.SubItems(lngCol - 1) = IIf(IsNull(varValue), "", varValue)
            If lst.Icon = "ItemNo" Then lst.ListSubItems(lngCol - 1).ForeColor = RGB(255, 0, 0)
        Next
        rs��Լ��λ.MoveNext
    Loop
    LockWindowUpdate 0
    
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
errHandle:
    If errCenter() = 1 Then Resume
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

