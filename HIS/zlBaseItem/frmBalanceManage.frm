VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmBalanceManage 
   Caption         =   "���㷽ʽ����"
   ClientHeight    =   4980
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   6645
   Icon            =   "frmBalanceManage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin MSComctlLib.ListView lvw_S 
      Height          =   1125
      Index           =   3
      Left            =   3630
      TabIndex        =   9
      Top             =   2490
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   1984
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   15658994
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "_ҽ�Ƹ��ʽ"
         Object.Tag             =   "ҽ�Ƹ��ʽ"
         Text            =   "ҽ�Ƹ��ʽ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "_ȱʡ���㷽ʽ"
         Object.Tag             =   "ȱʡ���㷽ʽ"
         Text            =   "ȱʡ���㷽ʽ"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.PictureBox picSplitH 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   3000
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   3000
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2760
      Width           =   3000
   End
   Begin MSComctlLib.ListView lvw_S 
      Height          =   855
      Index           =   2
      Left            =   3660
      TabIndex        =   1
      Top             =   1515
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   1508
      View            =   3
      Arrange         =   2
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
      BackColor       =   15658994
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "���㷽ʽ"
         Text            =   "���㷽ʽ"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "ȱʡ��"
         Text            =   "ȱʡ��"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Key             =   "Ӧ�տ�"
         Text            =   "Ӧ�տ�"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Key             =   "Ӧ����"
         Object.Tag             =   "Ӧ����"
         Text            =   "Ӧ����"
         Object.Width           =   1499
      EndProperty
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3225
      Left            =   3210
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3225
      ScaleWidth      =   45
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1560
      Width           =   45
   End
   Begin MSComctlLib.ListView lvw_S 
      Height          =   2235
      Index           =   1
      Left            =   360
      TabIndex        =   0
      Top             =   1500
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   3942
      Arrange         =   2
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
      BackColor       =   15658994
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "����"
         Text            =   "����"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "����"
         Text            =   "����"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "����"
         Text            =   "����"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Key             =   "Ӧ�տ�"
         Text            =   "Ӧ�տ�"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Key             =   "Ӧ����"
         Object.Tag             =   "Ӧ����"
         Text            =   "Ӧ����"
         Object.Width           =   1499
      EndProperty
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   2985
      Top             =   1845
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalanceManage.frx":0442
            Key             =   "Item1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalanceManage.frx":089A
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalanceManage.frx":0AB4
            Key             =   "No"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalanceManage.frx":0F0C
            Key             =   "Item3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalanceManage.frx":1360
            Key             =   "Item31"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalanceManage.frx":17B4
            Key             =   "Item4"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalanceManage.frx":3F66
            Key             =   "Item5"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalanceManage.frx":4500
            Key             =   "Item9"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   3000
      Top             =   2490
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalanceManage.frx":53DA
            Key             =   "Item1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalanceManage.frx":5832
            Key             =   "No"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalanceManage.frx":5C8A
            Key             =   "Item3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalanceManage.frx":60DE
            Key             =   "Item31"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalanceManage.frx":6532
            Key             =   "Item4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalanceManage.frx":8CE4
            Key             =   "Item5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalanceManage.frx":927E
            Key             =   "Item9"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   4
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
      Width1          =   615
      FixedBackground1=   0   'False
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   720
         Left            =   30
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
         ImageList       =   "imgToolsStard"
         HotImageList    =   "imgToolsHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   11
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
               Key             =   "New"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "Modify"
               Object.ToolTipText     =   "�޸�"
               Object.Tag             =   "�޸�"
               ImageKey        =   "Modify"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "Delete"
               Object.ToolTipText     =   "ɾ��"
               Object.Tag             =   "ɾ��"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split1"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Object.ToolTipText     =   "��������"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
      Left            =   2715
      Top             =   390
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
            Picture         =   "frmBalanceManage.frx":A158
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalanceManage.frx":A372
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalanceManage.frx":A58C
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalanceManage.frx":A7A6
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalanceManage.frx":A9C0
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalanceManage.frx":ABDA
            Key             =   "View"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalanceManage.frx":ADFA
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalanceManage.frx":B014
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgToolsStard 
      Left            =   2130
      Top             =   390
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
            Picture         =   "frmBalanceManage.frx":B22E
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalanceManage.frx":B448
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalanceManage.frx":B662
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalanceManage.frx":B87C
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalanceManage.frx":BA96
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalanceManage.frx":BCB0
            Key             =   "View"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalanceManage.frx":BED0
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalanceManage.frx":C0EA
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   4620
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   635
      SimpleText      =   $"frmBalanceManage.frx":C304
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmBalanceManage.frx":C34B
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
   Begin MSComctlLib.TabStrip tabMain 
      Height          =   795
      Left            =   3780
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1402
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.ShortcutCaption lbl 
      Height          =   300
      Index           =   2
      Left            =   3000
      TabIndex        =   8
      Top             =   960
      Width           =   3420
      _Version        =   589884
      _ExtentX        =   6032
      _ExtentY        =   529
      _StockProps     =   6
      Caption         =   "Ӧ���ڲ�ͬ���㳡���µĽ��㷽ʽ(&N)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
   Begin XtremeSuiteControls.ShortcutCaption lbl 
      Height          =   300
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   2100
      _Version        =   589884
      _ExtentX        =   3704
      _ExtentY        =   529
      _StockProps     =   6
      Caption         =   "���㷽ʽ(&M)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
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
      Begin VB.Menu mnusplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileexit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditAddNew 
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
      Begin VB.Menu mnuEditSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDefault 
         Caption         =   "����Ϊȱʡ��(&F)"
      End
      Begin VB.Menu mnuEditSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSetDefault 
         Caption         =   "����ȱʡ���㷽ʽ(&S)"
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
      Begin VB.Menu mnuViewSplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewSelect 
         Caption         =   "ѡ����(&C)"
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
      Begin VB.Menu mnuShortDefault 
         Caption         =   "����Ϊȱʡ��(&F)"
      End
      Begin VB.Menu mnuShortsplit2 
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
Attribute VB_Name = "frmBalanceManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintIndex As Integer '��ǰListView�����
Dim msngStartX As Single, msngStartY As Single    '�ƶ�ǰ����λ��
Dim mblnItem As Boolean  'Ϊ���ʾ������ListViewĳһ����
Dim mintColumn(1 To 3) As Integer
Dim mblnLoad As Boolean
Private Const mstrLvw1 As String = "����,1400,0,1;����,800,0,2;����,1440,0,0;Ӧ�տ�,840,0,0;Ӧ����,840,0,0"
Private Const mstrLvw2 As String = "���㷽ʽ,1400,0,1;ȱʡ��,600,0,0;Ӧ�տ�,840,0,0;Ӧ����,840,0,0"
Private Const mstrLvw3 As String = "ҽ�Ƹ��ʽ,1400,0,1;ȱʡ���㷽ʽ,1400,0,1"
Private mlngMode As Long
Private mstrPrivs As String                              'Ȩ�޴�
'Private mintProperty As Integer

Private Sub Form_Activate()
    If mblnLoad = True Then
        mblnLoad = False
        mnuViewReflash_Click
    End If
    mblnLoad = False
    If lvw_S(mintIndex).Visible And lvw_S(mintIndex).Enabled Then lvw_S(mintIndex).SetFocus
End Sub

Private Sub Form_Load()
    mblnLoad = True
    mlngMode = glngModul
    mstrPrivs = gstrPrivs
    
    Call Ȩ�޿���
    '���������ɾ����ListView�������
    lvw_S(1).Tag = "�ɱ仯��"
    lvw_S(2).Tag = "�ɱ仯��"
    '-----------
    RestoreWinState Me, App.ProductName
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    
    mintIndex = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name & "\ListView", "mintIndex", 1)
    If mintIndex > 2 Or mintIndex < 1 Then mintIndex = 2 '���ĳ���
    
    '���ListView�Ļ�δ�����ã������һ��ʹ�ã��Ǿ͵���ȱʡ�ĳ�ʼ��
    If lvw_S(1).ColumnHeaders.Count = 0 Then
        zlControl.LvwSelectColumns lvw_S(1), mstrLvw1, True
    End If
    If lvw_S(2).ColumnHeaders.Count = 0 Then
        zlControl.LvwSelectColumns lvw_S(2), mstrLvw2, True
    End If
    If lvw_S(3).ColumnHeaders.Count = 0 Then
        zlControl.LvwSelectColumns lvw_S(3), mstrLvw3, True
    End If
    '����LvwMain��ʾ���ö�Ӧ�˵�
     mnuViewIcon_Click lvw_S(mintIndex).View
End Sub

Private Sub Form_Resize()
    Dim sngTop As Single, sngBottom As Single
    On Error Resume Next
    sngTop = IIF(CoolBar1.Visible, CoolBar1.Top + CoolBar1.Height, 0)
    sngBottom = Me.ScaleHeight - IIF(stbThis.Visible, stbThis.Height, 0)
    
    lbl(1).Top = sngTop
    lbl(1).Width = lvw_S(1).Width
    lvw_S(1).Top = sngTop + lbl(1).Height
    lvw_S(1).Height = IIF(sngBottom - lvw_S(1).Top > 0, sngBottom - lvw_S(1).Top, 0)
    lvw_S(1).Left = 0
    lbl(1).Left = 0
    
    picSplit.Top = sngTop
    picSplit.Height = IIF(sngBottom - picSplit.Top > 0, sngBottom - picSplit.Top, 0)
    picSplit.Left = lvw_S(1).Left + lvw_S(1).Width
    
    lbl(2).Left = picSplit.Left + picSplit.Width
    lbl(2).Top = sngTop
    tabMain.Top = sngTop + lbl(2).Height
    tabMain.Height = sngBottom - tabMain.Top
    tabMain.Left = lbl(2).Left
    If Me.ScaleWidth - tabMain.Left > 0 Then tabMain.Width = Me.ScaleWidth - tabMain.Left
    lbl(2).Width = tabMain.Width
    
    If lvw_S(3).Visible = False Then
        lvw_S(2).Left = tabMain.ClientLeft
        lvw_S(2).Top = tabMain.ClientTop
        lvw_S(2).Width = tabMain.ClientWidth
        lvw_S(2).Height = tabMain.ClientHeight
    Else
        lvw_S(2).Left = tabMain.ClientLeft
        lvw_S(2).Top = tabMain.ClientTop
        lvw_S(2).Width = tabMain.ClientWidth
        lvw_S(2).Height = tabMain.ClientHeight - picSplitH.Height - lvw_S(3).Height
        
        picSplitH.Left = tabMain.ClientLeft
        picSplitH.Top = lvw_S(2).Top + lvw_S(2).Height
        picSplitH.Width = tabMain.ClientWidth
        
        lvw_S(3).Left = tabMain.ClientLeft
        lvw_S(3).Top = picSplitH.Top + picSplitH.Height
        lvw_S(3).Width = tabMain.ClientWidth
    End If
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name & "\ListView", "mintIndex", mintIndex)
    SaveWinState Me, App.ProductName
End Sub

Private Sub lvw_S_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn(Index) = ColumnHeader.Index - 1 Then '���Ǹղ�����
        lvw_S(Index).SortOrder = IIF(lvw_S(Index).SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn(Index) = ColumnHeader.Index - 1
        lvw_S(Index).SortKey = mintColumn(Index)
        lvw_S(Index).SortOrder = lvwAscending
    End If
End Sub

Private Sub lvw_S_DblClick(Index As Integer)
    If mblnItem = True And mnuEditModify.Enabled And mnuEditModify.Visible Then
        mnuEditModify_Click
    End If
End Sub

Private Sub lvw_S_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    On Error GoTo errDrop
    
    If Source.DragIcon.Handle = ils32.ListImages("No").Picture.Handle Then Exit Sub
    
    gstrSQL = ""
    If Index = 2 Then
        '�ڽ��㳡��������һ�ֽ��㷽ʽ
        If Source Is lvw_S(1) Then
            gstrSQL = "zl_���㷽ʽӦ��_insert('" & tabMain.SelectedItem.Caption & "','" & lvw_S(1).SelectedItem.Text & "')"
        End If
    Else
        '�ӽ��㳡����ɾ��һ�ֽ��㷽ʽ
        If Source Is lvw_S(2) Then
            gstrSQL = "zl_���㷽ʽӦ��_delete('" & tabMain.SelectedItem.Caption & "','" & lvw_S(2).SelectedItem.Text & "')"
        End If
    End If
    If gstrSQL <> "" Then
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Fill����Ӧ��
    End If
    Exit Sub
errDrop:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvw_S_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
    Dim strSource As String, i As Long
    
    If Index = 2 Then
        With lvw_S(2)
            If State = 1 Or Source Is lvw_S(2) Then
                Source.DragIcon = ils32.ListImages("No").Picture
            Else
                '75134:���ϴ�,2014/7/14,����Ϊ9�ķ�ʽ����������Ӧ�ó���
                If Val(Split(lvw_S(1).SelectedItem.Tag & ",", ",")(0)) = 9 Then
                    Source.DragIcon = ils32.ListImages("No").Picture
                    Exit Sub
                End If
                
                '82990:���ϴ�,2015/3/9,ҽ���������ڲ�����
                '�����㡢���ѿ���ֻ��ʹ������Ϊ1,2,8��
                If InStr(",1,2,8,", "," & Val(Split(lvw_S(1).SelectedItem.Tag & ",", ",")(0)) & ",") = 0 _
                    And InStr(",������,���ѿ�,", "," & tabMain.SelectedItem.Caption & ",") > 0 Then
                    Source.DragIcon = ils32.ListImages("No").Picture
                    Exit Sub
                End If
                
                '���տ�ֻ��Ӧ����Ԥ����
                If Val(Split(lvw_S(1).SelectedItem.Tag & ",", ",")(0)) = 5 And tabMain.SelectedItem.Caption <> "Ԥ����" Then
                    Source.DragIcon = ils32.ListImages("No").Picture
                    Exit Sub
                End If
                
                '�ж��Ƿ��Ѿ�����
                strSource = lvw_S(1).SelectedItem.Text
                For i = 1 To lvw_S(2).ListItems.Count
                    If strSource = lvw_S(2).ListItems(i).Text Then
                        Source.DragIcon = ils32.ListImages("No").Picture
                        Exit Sub
                    End If
                Next
                Source.DragIcon = lvw_S(1).SelectedItem.CreateDragImage
            End If
        End With
    ElseIf Index = 1 Then
        With lvw_S(1)
            If State = 1 Or Source Is lvw_S(1) Then
                Source.DragIcon = ils32.ListImages("No").Picture
            Else
                Source.DragIcon = ils32.ListImages("Delete").Picture
            End If
        End With
    Else
        Source.DragIcon = ils32.ListImages("No").Picture
    End If
End Sub

Private Sub lvw_S_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    Dim varData As Variant
    
    If Index = 2 Then
        mblnItem = True
        varData = Split(Item.Tag & ",", ",")
        'Ϊ�ұߵ�Ӧ�ó���ʱ
        If Not Item Is Nothing Then
            '����Ϊ1,2,7,8���ҷ�Ӧ����Ľ��㷽ʽ�ɱ�����Ϊ�ó����µ�ȱʡ��
            If InStr("1,2,7,8", Val(varData(0))) > 0 And Val(varData(1)) = 0 Then
                mnuEditDefault.Enabled = True
            Else
                mnuEditDefault.Enabled = False
            End If
        End If
     '75134:���ϴ�,2014/7/14,����������Ϊ9ʱ���������޸ĺ�ɾ��
    ElseIf Index = 1 Then
        mblnItem = True
        varData = Split(Item.Tag & ",", ",")
        If Not Item Is Nothing Then
            mnuEditDefault.Enabled = False
            If Val(varData(0)) = 9 Then
                Me.mnuEditModify.Enabled = False
                Me.mnuEditDelete.Enabled = False
                Toolbar1.Buttons("Modify").Enabled = False
                Toolbar1.Buttons("Delete").Enabled = False
            ElseIf Val(varData(2)) = 1 Then
                Me.mnuEditModify.Enabled = True
                Me.mnuEditDelete.Enabled = False
                Toolbar1.Buttons("Modify").Enabled = True
                Toolbar1.Buttons("Delete").Enabled = False
            Else
                Me.mnuEditModify.Enabled = True
                Me.mnuEditDelete.Enabled = True
                Toolbar1.Buttons("Modify").Enabled = True
                Toolbar1.Buttons("Delete").Enabled = True
            End If
        End If
    End If
    stbThis.Panels(2).Text = ""
End Sub

Private Sub lvw_S_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim Item As ListItem
    If mnuEditModify.Visible = False Or Index = 3 Then Exit Sub
    Set Item = lvw_S(Index).HitTest(X, Y)
    If Button = 1 And Not Item Is Nothing And (Abs(X - msngStartX) > 100 Or Abs(Y - msngStartY) > 100) Then
        lvw_S(Index).DragIcon = ils32.ListImages("No").Picture
        lvw_S(Index).Drag 1
    End If
End Sub

Private Sub lvw_S_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    If Index = 3 Then Exit Sub
    If Button = 2 Then
        mnuShortMenu(1).Enabled = mnuEditAddNew.Enabled
        mnuShortMenu(2).Enabled = mnuEditModify.Enabled
        mnuShortMenu(3).Enabled = mnuEditDelete.Enabled
        For i = 0 To 3
            mnuShortIcon(i).Checked = mnuViewIcon(i).Checked
        Next
        PopupMenu mnuShort, vbPopupMenuRightButton
    End If
End Sub
Private Sub lvw_S_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And mnuEditModify.Enabled And mnuEditModify.Visible Then mnuEditModify_Click
End Sub

Private Sub lvw_S_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnItem = False
    msngStartX = X: msngStartY = Y
End Sub

Private Sub lvw_S_GotFocus(Index As Integer)
    Dim i As Integer
    Dim Item As ListItem
    mintIndex = Index
    With lvw_S(Index)
        For i = 1 To 3
            lvw_S(i).BackColor = &HEEEFF2
        Next
        .BackColor = vbHighlightText
        For i = 0 To 3
            mnuViewIcon(i).Checked = False
        Next
        mnuViewIcon(.View).Checked = True
    End With
    Call SetMenu
    
    Set Item = lvw_S(1).SelectedItem
    '75134:���ϴ�,2014/7/14,���²˵���������״̬
    If mintIndex = 1 And Not Item Is Nothing Then Call lvw_S_ItemClick(1, Item)
End Sub

Private Sub mnuEditAddNew_Click()
    If mintIndex = 1 Then
        If frmBalanceEdit.�༭���㷽ʽ("") = True Then
            Fill���㷽ʽ
            Fill����Ӧ��
        End If
    Else
        If frmBalanceUse.�༭����(tabMain.SelectedItem.Caption) = True Then
            Fill����Ӧ��
        End If
    End If
End Sub

Private Sub mnuEditModify_Click()
    If mintIndex = 1 Then
        If frmBalanceEdit.�༭���㷽ʽ(Mid(lvw_S(1).SelectedItem.Key, 2)) = True Then
            Fill���㷽ʽ
            Fill����Ӧ��
        End If
    Else
        If frmBalanceUse.�༭����(tabMain.SelectedItem.Caption) = True Then
            Fill����Ӧ��
        End If
    End If
End Sub

Private Sub mnuEditDefault_Click()
    On Error GoTo ErrHandle
    
    If lvw_S(2).SelectedItem Is Nothing Then Exit Sub
    gstrSQL = "zl_���㷽ʽӦ��_default('" & tabMain.SelectedItem.Caption & "','" & lvw_S(2).SelectedItem.Text & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Fill����Ӧ��
    stbThis.Panels(2).Text = "���óɹ���"
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditDelete_Click()
    On Error GoTo errDelete
    Dim intIndex As Integer
    
    '75134:���ϴ�,2014/7/14,����Ϊ9��̶�Ϊ1�Ľ��㷽ʽ������ɾ��
    If mintIndex = 1 And (Split(lvw_S(1).SelectedItem.Tag & ",", ",")(0) = 9 Or Split(lvw_S(1).SelectedItem.Tag & ",", ",")(2) = 1) Then Exit Sub
        
    Select Case mintIndex
        Case 1
            If MsgBox("��ȷ��Ҫɾ������Ϊ��" & lvw_S(1).SelectedItem.Text & "���Ľ��㷽ʽ��", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            gstrSQL = "zl_���㷽ʽ_delete('" & Mid(lvw_S(1).SelectedItem.Key, 2) & "')"
        Case 2
            If MsgBox("��ȷ��Ҫɾ������Ϊ��" & lvw_S(2).SelectedItem.Text & "���Ľ��㷽ʽӦ����", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            gstrSQL = "zl_���㷽ʽӦ��_delete('" & tabMain.SelectedItem.Caption & "','" & lvw_S(2).SelectedItem.Text & "')"
    End Select
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    Select Case mintIndex
        Case 1
            With lvw_S(1)
                intIndex = .SelectedItem.Index
                .ListItems.Remove .SelectedItem.Key
                If .ListItems.Count > 0 Then
                    intIndex = IIF(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
                    .ListItems(intIndex).Selected = True
                    .ListItems(intIndex).EnsureVisible
                End If
            End With
            Fill����Ӧ��
            '75134:���ϴ�,2014/7/14,���²˵���������״̬
            If Not lvw_S(1).SelectedItem Is Nothing Then Call lvw_S_ItemClick(mintIndex, lvw_S(1).SelectedItem)
        Case 2
            With lvw_S(2)
                intIndex = .SelectedItem.Index
                .ListItems.Remove .SelectedItem.Key
                If .ListItems.Count > 0 Then
                    intIndex = IIF(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
                    .ListItems(intIndex).Selected = True
                    .ListItems(intIndex).EnsureVisible
                End If
            End With
            Call SetMenu
    End Select
    
    Exit Sub
errDelete:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditSetDefault_Click()
    If lvw_S(2).ListItems.Count = 0 Then Exit Sub
    If frmBalanceDefaultSet.ShowMe(Me, tabMain.SelectedItem.Caption) Then
        Fill����Ӧ��
    End If
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    'Ĭ�ϲ���������=���㷽ʽ����
    Dim str���� As String
    
    If Not lvw_S(1).SelectedItem Is Nothing Then
        str���� = Mid(lvw_S(1).SelectedItem.Key, 2)
    End If
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
        "����=" & str����)
End Sub

Private Sub mnuShortDefault_Click()
    mnuEditDefault_Click
End Sub

Private Sub mnuShortMenu_Click(Index As Integer)
    Select Case Index
        Case 1
            mnuEditAddNew_Click
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
    If Index > 3 Then Index = 0
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
        Toolbar1.Buttons("View").ButtonMenus(i + 1).Text = Replace(Toolbar1.Buttons("View").ButtonMenus(i + 1).Text, "��", "  ")
    Next
    mnuViewIcon(Index).Checked = True
    Toolbar1.Buttons("View").ButtonMenus(Index + 1).Text = Replace(Toolbar1.Buttons("View").ButtonMenus(Index + 1).Text, "  ", "��")
    lvw_S(mintIndex).View = Index
End Sub

Private Sub mnuViewReflash_Click()
    Fill���㷽ʽ
    Fill���㳡��
    Fill����Ӧ��
End Sub

Private Sub mnuViewSelect_Click()
    If mintIndex = 1 Then
        If zlControl.LvwSelectColumns(lvw_S(1), mstrLvw1) = True Then
            '���б仯��Ҫ����ˢ��
            Fill���㷽ʽ
        End If
    ElseIf mintIndex = 2 Then
        If zlControl.LvwSelectColumns(lvw_S(2), mstrLvw2) = True Then
            '���б仯��Ҫ����ˢ��
            Fill����Ӧ��
        End If
    End If
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
        If sngTemp > 500 And Me.ScaleWidth - (sngTemp + picSplit.Width) > 1000 Then
            picSplit.Left = sngTemp
            lvw_S(1).Width = picSplit.Left - lvw_S(1).Left
            tabMain.Left = picSplit.Left + picSplit.Width
            tabMain.Width = Me.ScaleWidth - tabMain.Left
            lbl(1).Width = lvw_S(1).Width
            
            lbl(2).Left = tabMain.Left
            lbl(2).Width = tabMain.Width
            lvw_S(2).Left = tabMain.ClientLeft

            lvw_S(2).Width = tabMain.ClientWidth
            picSplitH.Left = tabMain.ClientLeft
            picSplitH.Width = tabMain.ClientWidth
            lvw_S(3).Left = tabMain.ClientLeft
            lvw_S(3).Width = tabMain.ClientWidth
        End If
    End If
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnufilePreview_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub mnufileset_Click()
    zlPrintSet
End Sub

Private Sub picSplitH_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        msngStartY = Y
    End If
End Sub

Private Sub picSplitH_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngTemp As Single
    
    On Error Resume Next
    If Button = 1 Then
        sngTemp = picSplitH.Top + Y - msngStartY
        If sngTemp > tabMain.Top + 1000 And Me.ScaleHeight - (sngTemp + picSplitH.Height) > 1000 Then
            picSplitH.Top = sngTemp
            lvw_S(2).Height = picSplitH.Top - lvw_S(2).Top
            lvw_S(3).Move lvw_S(3).Left, picSplitH.Top + picSplitH.Height, lvw_S(3).Width, _
                tabMain.Top + tabMain.Height - (picSplitH.Top + picSplitH.Height) - 50
        End If
    End If
End Sub

Private Sub tabMain_Click()
    Fill����Ӧ��
    lvw_S(2).SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "New"
            mnuEditAddNew_Click
        Case "Modify"
            mnuEditModify_Click
        Case "Delete"
            mnuEditDelete_Click
        Case "Exit"
            mnuFileExit_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Preview"
            mnufilePreview_Click
        Case "Help"
            mnuhelptopic_Click
        Case "View"
            mnuViewIcon_Click lvw_S(mintIndex).View + 1
    End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    mnuViewIcon_Click ButtonMenu.Index - 1
End Sub

Private Sub Toolbar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool
    End If
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


Private Sub subPrint(bytMode As Byte)
'����:���д�ӡ,Ԥ���������EXCEL
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    If gstrUserName = "" Then Call GetUserInfo
    On Error Resume Next
    Dim objPrint As New zlPrintLvw
    Set objPrint.Body.objData = lvw_S(mintIndex)
    Select Case mintIndex
        Case 1
            objPrint.Title.Text = "���㷽ʽ"
        Case 2
            objPrint.Title.Text = "���㷽ʽ��" & tabMain.SelectedItem.Caption & "�����µ�Ӧ��"
        Case 3
            objPrint.Title.Text = tabMain.SelectedItem.Caption & "����ȱʡ���㷽ʽ"
    End Select
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

Public Sub Fill���㷽ʽ()
'����:�ѽ��㷽ʽװ�뵽lvw_S(1)��
    Dim rs���㷽ʽ As New ADODB.Recordset
    Dim lst As ListItem
    Dim strKey As String
    Dim i As Integer
    Dim varValue As Variant
    
    On Error GoTo ErrHandle
    If Not lvw_S(1).SelectedItem Is Nothing Then
        '����ԭ�м�ֵ
        strKey = lvw_S(1).SelectedItem.Key
    End If
    rs���㷽ʽ.CursorLocation = adUseClient
    rs���㷽ʽ.CursorType = adOpenKeyset
    rs���㷽ʽ.LockType = adLockReadOnly
    gstrSQL = "Select ����,����,����,����,�Ƿ�̶�,ȱʡ��־,Decode(Nvl(Ӧ�տ�,0),1,'��',' ') Ӧ�տ� ,Decode(Nvl(Ӧ����,0),1,'��',' ') Ӧ���� From ���㷽ʽ"
    Set rs���㷽ʽ = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)

    lvw_S(1).ListItems.Clear
    Do Until rs���㷽ʽ.EOF
        Set lst = lvw_S(1).ListItems.Add(, "C" & rs���㷽ʽ!����, rs���㷽ʽ!����)
        
        Select Case Nvl(rs���㷽ʽ!����, 1)
            Case 1, 2, 6, 7
                lst.Icon = "Item1": lst.SmallIcon = "Item1"
            Case 3, 4
                lst.Icon = "Item4": lst.SmallIcon = "Item4"
            Case 5
                lst.Icon = "Item5": lst.SmallIcon = "Item5"
            Case 8
                lst.Icon = "Item1": lst.SmallIcon = "Item1"
            Case 9
                lst.Icon = "Item9": lst.SmallIcon = "Item9"
        End Select
        '75134:���ϴ�,2014/7/14,tag������ʽ��Ϊ"����,Ӧ����,�Ƿ�̶�"
        lst.Tag = Nvl(rs���㷽ʽ!����, 1) & "," & IIF(Nvl(rs���㷽ʽ!Ӧ����) = "", 0, 1) & "," & Nvl(rs���㷽ʽ!�Ƿ�̶�, 0)

        For i = 2 To lvw_S(1).ColumnHeaders.Count
            varValue = rs���㷽ʽ(lvw_S(1).ColumnHeaders(i).Text).value
            lst.SubItems(i - 1) = Nvl(varValue)
        Next
        rs���㷽ʽ.MoveNext
    Loop
    
    Dim Item As ListItem
    On Error Resume Next
    Set Item = lvw_S(1).ListItems(strKey)
    If err <> 0 Then
        Set Item = lvw_S(1).ListItems(1)
        Item.Selected = True
        Item.EnsureVisible
    Else
        err.Clear
        Item.Selected = True
        Item.EnsureVisible
    End If
    Call lvw_S_GotFocus(mintIndex)
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub Fill���㳡��()
'����:�ѽ��㷽ʽװ�뵽tabMain��
    Dim rs���㳡�� As New ADODB.Recordset
    Dim lst As ListItem
    Dim strKey As String
    
    On Error GoTo ErrHandle
    If Not tabMain.SelectedItem Is Nothing Then
        '����ԭ�м�ֵ
        strKey = tabMain.SelectedItem.Key
    End If
    rs���㳡��.CursorLocation = adUseClient
    gstrSQL = "select ����,����,���� from ���㳡�� Order By ����"
    Set rs���㳡�� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)

    If rs���㳡��.EOF = True Then
        MsgBox "��ϵͳ����Ա���䡰���㳡�ϡ��ĳ�ʼ�����ݡ�", vbExclamation, gstrSysName
        Unload Me
        Exit Sub
    End If
    tabMain.Tabs.Clear
    Do Until rs���㳡��.EOF
        tabMain.Tabs.Add , "C" & rs���㳡��("����"), rs���㳡��("����")
        rs���㳡��.MoveNext
    Loop
    
    Dim Item As MSComctlLib.Tab
    On Error Resume Next
    Set Item = tabMain.Tabs(strKey)
    If err <> 0 Then
        Set Item = tabMain.Tabs(1)
        Item.Selected = True
    Else
        err.Clear
        Item.Selected = True
    End If
    Call Fill����Ӧ��
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub Fill����Ӧ��()
'����:�ѽ��㷽ʽװ�뵽lvw_S(2)��
    Dim rs����Ӧ�� As New ADODB.Recordset
    Dim lst As ListItem
    Dim strIco As String
    Dim strKey As String
    Dim i As Integer, int���� As Integer
    Dim str���㳡��  As String
    
    On Error GoTo ErrHandle
    If tabMain.SelectedItem Is Nothing Then
        lvw_S(2).ListItems.Clear
        Call SetMenu
        lvw_S(3).Visible = False
        picSplitH.Visible = False
        Exit Sub
    End If
    str���㳡�� = tabMain.SelectedItem.Caption
    If str���㳡�� = "" Then
        lvw_S(2).ListItems.Clear
        Call SetMenu
        lvw_S(3).Visible = False
        picSplitH.Visible = False
        Exit Sub
    End If
    
    If Not lvw_S(2).SelectedItem Is Nothing Then
        '����ԭ�м�ֵ
        strKey = lvw_S(2).SelectedItem.Key
    End If
    
    rs����Ӧ��.CursorLocation = adUseClient
    gstrSQL = _
        " Select A.���㷽ʽ,Nvl(B.����,1) as ����,A.ȱʡ��־,Nvl(B.Ӧ�տ�,0) Ӧ�տ�,Nvl(B.Ӧ����,0) Ӧ����" & _
        " From ���㷽ʽӦ�� A,���㷽ʽ B" & _
        " Where A.���㷽ʽ=B.����(+) And A.Ӧ�ó���=[1] And A.���ʽ Is Null"
    Set rs����Ӧ�� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str���㳡��)
        
    lvw_S(2).ListItems.Clear
    Do Until rs����Ӧ��.EOF
        If Nvl(rs����Ӧ��!ȱʡ��־, 0) = 1 Then
            Set lst = lvw_S(2).ListItems.Add(, "C" & rs����Ӧ��.AbsolutePosition, rs����Ӧ��!���㷽ʽ, "Item31", "Item31")
            '�����б�ɾ��
            If lvw_S(2).ColumnHeaders.Count > 1 Then lst.SubItems(1) = "��"
        Else
            Set lst = lvw_S(2).ListItems.Add(, "C" & rs����Ӧ��.AbsolutePosition, rs����Ӧ��!���㷽ʽ, "Item3", "Item3")
        End If
        
        If lvw_S(2).ColumnHeaders.Count > 2 Then lst.SubItems(2) = IIF(rs����Ӧ��!Ӧ�տ� = 1, "��", " ")
        If lvw_S(2).ColumnHeaders.Count > 3 Then lst.SubItems(3) = IIF(rs����Ӧ��!Ӧ���� = 1, "��", " ")
        int���� = Val(Nvl(rs����Ӧ��!����))
        lst.Tag = int���� & "," & Val(Nvl(rs����Ӧ��!Ӧ����))
        
        If int���� = 3 Or int���� = 4 Then
            lst.Icon = "Item4": lst.SmallIcon = "Item4"
        ElseIf int���� = 5 Then
            lst.Icon = "Item5": lst.SmallIcon = "Item5"
        End If
                
        rs����Ӧ��.MoveNext
    Loop
    
    If str���㳡�� = "�շ�" Then
        gstrSQL = _
            " Select A.���ʽ,A.���㷽ʽ" & _
            " From ���㷽ʽӦ�� A" & _
            " Where A.Ӧ�ó���=[1] And A.���ʽ Is Not Null"
        Set rs����Ӧ�� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str���㳡��)
            
        lvw_S(3).ListItems.Clear
        Do Until rs����Ӧ��.EOF
            Set lst = lvw_S(3).ListItems.Add(, "C" & rs����Ӧ��.AbsolutePosition, Nvl(rs����Ӧ��!���ʽ))
            lst.SubItems(1) = Nvl(rs����Ӧ��!���㷽ʽ)
            rs����Ӧ��.MoveNext
        Loop
    End If
    
    Dim Item As ListItem
    On Error Resume Next
    Set Item = lvw_S(2).ListItems(strKey)
    If err <> 0 Then
        Set Item = lvw_S(2).ListItems(1)
        Item.Selected = True
        Item.EnsureVisible
    Else
        err.Clear
        Item.Selected = True
        Item.EnsureVisible
    End If
    Call SetMenu
    '���շѡ�Ӧ�ó��ϴ���ҽ�Ƹ��ʽȱʡ���㷽ʽʱ����ʾ�б�
    lvw_S(3).Visible = (str���㳡�� = "�շ�") And lvw_S(3).ListItems.Count > 0
    picSplitH.Visible = lvw_S(3).Visible
    Call Form_Resize
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub Ȩ�޿���()
'����:�����е��û�Ȩ�޲���,��ʹһЩ�˵����ť���ɼ�
    If InStr(mstrPrivs, "��ɾ��") = 0 Then
        mnuEdit.Visible = False
        mnuEditModify.Visible = False
        mnuShortMenu(1).Visible = False
        mnuShortMenu(2).Visible = False
        mnuShortMenu(3).Visible = False
        mnuShortsplit1.Visible = False
        Toolbar1.Buttons("Split").Visible = False
        Toolbar1.Buttons("New").Visible = False
        Toolbar1.Buttons("Modify").Visible = False
        Toolbar1.Buttons("Delete").Visible = False
        
        mnuEditSplit.Visible = False
        mnuEditDefault.Visible = False
        mnuEditSplit1.Visible = False
        mnuEditSetDefault.Visible = False
    End If
End Sub

Private Sub EnablePrint(ByVal blnEnabled As Boolean)
'����:���ô�ӡ��Ԥ����ť����Чֵ
'����:blnEnabled ��Чֵ
    Toolbar1.Buttons("Print").Enabled = blnEnabled
    Toolbar1.Buttons("Preview").Enabled = blnEnabled
    mnuFilePreview.Enabled = blnEnabled
    mnuFilePrint.Enabled = blnEnabled
    mnuFileExcel.Enabled = blnEnabled
End Sub

Private Sub SetMenu()
    Dim varData As Variant, i As Integer
    If mintIndex = 3 Then
        For i = mnuViewIcon.LBound To mnuViewIcon.UBound
            mnuViewIcon(i).Enabled = False
        Next
        Toolbar1.Buttons("View").Enabled = False
        mnuViewSelect.Enabled = False
        
        With lvw_S(3)
            mnuEditModify.Enabled = False
            mnuEditDelete.Enabled = False
            Toolbar1.Buttons("Modify").Enabled = mnuEditModify.Enabled
            Toolbar1.Buttons("Delete").Enabled = mnuEditDelete.Enabled
            EnablePrint lvw_S(3).ListItems.Count > 0
            mnuEditDefault.Enabled = False
            mnuEditSetDefault.Enabled = lvw_S(2).ListItems.Count > 0 And tabMain.SelectedItem.Caption = "�շ�"
        End With
    Else
        For i = mnuViewIcon.LBound To mnuViewIcon.UBound
            mnuViewIcon(i).Enabled = True
        Next
        Toolbar1.Buttons("View").Enabled = True
        mnuViewSelect.Enabled = True
        
        With lvw_S(mintIndex)
            If .ListItems.Count = 0 Then
                mnuEditModify.Enabled = False
                mnuEditDelete.Enabled = False
                mnuEditDefault.Enabled = False
                mnuEditSetDefault.Enabled = False
                Toolbar1.Buttons("Modify").Enabled = False
                Toolbar1.Buttons("Delete").Enabled = False
                EnablePrint False
            Else
                mnuEditModify.Enabled = True
                mnuEditDelete.Enabled = True
                mnuEditDefault.Enabled = True
                mnuEditSetDefault.Enabled = mintIndex = 2 And tabMain.SelectedItem.Caption = "�շ�"
                Toolbar1.Buttons("Modify").Enabled = True
                Toolbar1.Buttons("Delete").Enabled = True
                EnablePrint True
            End If
        End With
        If lvw_S(2).SelectedItem Is Nothing Then
            mnuEditDefault.Enabled = False
        Else
            varData = Split(lvw_S(2).SelectedItem.Tag & ",", ",")
            '����Ϊ1,2,7,8���ҷ�Ӧ����Ľ��㷽ʽ�ɱ�����Ϊ�ó����µ�ȱʡ��
            If InStr("1,2,7,8", Val(varData(0))) > 0 And Val(varData(1)) = 0 Then
                mnuEditDefault.Enabled = False
            Else
                mnuEditDefault.Enabled = True
            End If
        End If
        If mintIndex = 1 Then
            mnuEditAddNew.Enabled = True
            Toolbar1.Buttons("New").Enabled = True
        Else
            mnuEditAddNew.Enabled = lvw_S(1).ListItems.Count > 0
            Toolbar1.Buttons("New").Enabled = lvw_S(1).ListItems.Count > 0
        End If
    End If
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

