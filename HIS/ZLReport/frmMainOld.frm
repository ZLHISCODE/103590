VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "�������"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11475
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7005
   ScaleWidth      =   11475
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImgGroup32 
      Left            =   4290
      Top             =   2190
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":014A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0464
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":077E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0A98
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0DB2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgGroup16 
      Left            =   4290
      Top             =   1620
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
            Picture         =   "frmMain.frx":10CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1226
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1380
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1634
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img32 
      Left            =   4845
      Top             =   2190
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":178E
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1AA8
            Key             =   "Publish"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1DC2
            Key             =   "Fixed"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":20DC
            Key             =   "PubFixed"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23F6
            Key             =   "Bill"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2710
            Key             =   "BillPublish"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   4860
      Top             =   1620
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
            Picture         =   "frmMain.frx":2A2A
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B84
            Key             =   "Publish"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2CDE
            Key             =   "Fixed"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E38
            Key             =   "PubFixed"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F92
            Key             =   "Bill"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":30EC
            Key             =   "BillPublish"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwReport 
      Height          =   5025
      Left            =   2355
      TabIndex        =   2
      Top             =   1590
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   8864
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDropMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img32"
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      OLEDropMode     =   1
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   4762
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "���"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "˵��"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "�޸�ʱ��"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "����ʱ��"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "���ִ��ʱ��"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "���ִ����"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "����"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "����"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "������������Դ"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "����"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "������������"
         Object.Width           =   6223
      EndProperty
   End
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   11475
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinHeight1      =   720
      Width1          =   810
      Key1            =   "cbr_Funcs"
      NewRow1         =   0   'False
      BandForeColor2  =   8388608
      Child2          =   "picSysFind"
      MinWidth2       =   2310
      MinHeight2      =   495
      Width2          =   2370
      UseCoolbarColors2=   0   'False
      Key2            =   "cbr_SysFind"
      NewRow2         =   0   'False
      BandStyle2      =   1
      Begin VB.PictureBox picSysFind 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   9075
         ScaleHeight     =   495
         ScaleWidth      =   2370
         TabIndex        =   8
         Top             =   135
         Width           =   2370
         Begin VB.CommandButton cmdNext 
            Caption         =   "��һ��(F3)"
            Height          =   350
            Left            =   5450
            TabIndex        =   13
            Top             =   75
            Width           =   1065
         End
         Begin VB.ComboBox cboSys 
            Height          =   300
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   100
            Width           =   2100
         End
         Begin VB.TextBox txtFind 
            Height          =   300
            Left            =   3900
            MaxLength       =   50
            TabIndex        =   12
            Tag             =   "����"
            ToolTipText     =   "֧�ְ����ơ���š�ƴ���������"
            Top             =   100
            Width           =   1515
         End
         Begin VB.Label lblSys 
            AutoSize        =   -1  'True
            Caption         =   "ϵͳ(&S)"
            ForeColor       =   &H008B0000&
            Height          =   180
            Left            =   60
            TabIndex        =   9
            Top             =   165
            Width           =   630
         End
         Begin VB.Label lblFind 
            AutoSize        =   -1  'True
            Caption         =   "���ұ���(&L)"
            ForeColor       =   &H008B0000&
            Height          =   180
            Left            =   2880
            TabIndex        =   11
            Top             =   165
            Width           =   990
         End
      End
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   165
         TabIndex        =   1
         Top             =   30
         Width           =   8820
         _ExtentX        =   15558
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   23
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ִ��"
               Key             =   "Report"
               Description     =   "ִ��"
               Object.ToolTipText     =   "ִ�б���"
               Object.Tag             =   "ִ��"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "GroupAdd"
               Description     =   "����"
               Object.ToolTipText     =   "���ӱ�����"
               Object.Tag             =   "����"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "GroupModify"
               Description     =   "�޸�"
               Object.ToolTipText     =   "�޸ı���"
               Object.Tag             =   "�޸�"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "GroupDel"
               Description     =   "ɾ��"
               Object.ToolTipText     =   "ɾ��������"
               Object.Tag             =   "ɾ��"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Group_"
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Add"
               Description     =   "����"
               Object.ToolTipText     =   "�����Զ��屨��"
               Object.Tag             =   "����"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "Modi"
               Description     =   "�޸�"
               Object.ToolTipText     =   "�޸ı�������"
               Object.Tag             =   "�޸�"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "Del"
               Description     =   "ɾ��"
               Object.ToolTipText     =   "ɾ����ǰ����"
               Object.Tag             =   "ɾ��"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Report_"
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "���"
               Key             =   "Design"
               Description     =   "���"
               Object.ToolTipText     =   "��Ʊ���"
               Object.Tag             =   "���"
               ImageKey        =   "Design"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Design_"
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��"
               Key             =   "Guide"
               Description     =   "��"
               Object.ToolTipText     =   "������"
               Object.Tag             =   "��"
               ImageKey        =   "Guide"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Guide_"
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Publish"
               Description     =   "����"
               Object.ToolTipText     =   "��������"
               Object.Tag             =   "����"
               ImageIndex      =   9
               Style           =   5
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ȡ��"
               Key             =   "unPub"
               Description     =   "ȡ��"
               Object.ToolTipText     =   "ȡ������"
               Object.Tag             =   "ȡ��"
               ImageIndex      =   10
               Style           =   5
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Pub_"
               Style           =   3
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�鿴"
               Key             =   "View"
               Description     =   "�鿴"
               Object.ToolTipText     =   "�б�鿴��ʽ"
               Object.Tag             =   "�鿴"
               ImageIndex      =   11
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   5
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Icon"
                     Object.Tag             =   "��ͼ��(&I)"
                     Text            =   "��ͼ��(&I)"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Small"
                     Object.Tag             =   "Сͼ��(&S)"
                     Text            =   "Сͼ��(&S)"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "List"
                     Object.Tag             =   "�б�(&L)"
                     Text            =   "�б�(&L)"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Detail"
                     Object.Tag             =   "��ϸ����(&D)"
                     Text            =   "��ϸ����(&D)"
                  EndProperty
                  BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "RunLog"
                     Object.Tag             =   "������־"
                     Text            =   "������־"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "����"
               Object.ToolTipText     =   "��ǰ��������"
               Object.Tag             =   "����"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Quit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   13
            EndProperty
         EndProperty
         Begin MSComctlLib.Toolbar tbrCheck 
            Height          =   720
            Left            =   2610
            TabIndex        =   7
            Top             =   0
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   1270
            ButtonWidth     =   1455
            ButtonHeight    =   1270
            AllowCustomize  =   0   'False
            Wrappable       =   0   'False
            Style           =   1
            ImageList       =   "imgGray"
            HotImageList    =   "imgColor"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   1
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "���ܼ��"
                  Key             =   "Check"
                  Description     =   "���ܼ��"
                  Object.ToolTipText     =   "���ܼ��"
                  Object.Tag             =   "���ܼ��"
                  ImageIndex      =   15
               EndProperty
            EndProperty
         End
      End
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   75
      Top             =   690
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3246
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3460
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":367A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3894
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3AAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3CC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3EE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":40FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4316
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4530
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":474A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4964
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4B7E
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4D98
            Key             =   "Guide"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4FB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":51CC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   705
      Top             =   690
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":53E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5600
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":581A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5A34
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5C4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5E68
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6082
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":629C
            Key             =   "Design"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":64B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":66D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":68EA
            Key             =   "View"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6B04
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6D1E
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6F38
            Key             =   "Guide"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7152
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":736C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdg 
      Left            =   2265
      Top             =   1245
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvwGroup 
      Height          =   5085
      Left            =   30
      TabIndex        =   4
      Top             =   1530
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   8969
      View            =   2
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImgGroup32"
      SmallIcons      =   "ImgGroup16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "���"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "˵��"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "����ʱ��"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "����"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   6645
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMain.frx":7586
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15161
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
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
   Begin VB.Label LblReport 
      Alignment       =   2  'Center
      BackColor       =   &H009B6737&
      Caption         =   "����"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2400
      TabIndex        =   6
      Top             =   1290
      Width           =   2055
   End
   Begin VB.Label lblGroup 
      Alignment       =   2  'Center
      BackColor       =   &H009B6737&
      Caption         =   "������"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   0
      TabIndex        =   5
      Top             =   1290
      Width           =   2055
   End
   Begin VB.Image ImgSplit_S 
      Height          =   4980
      Left            =   2280
      MousePointer    =   9  'Size W E
      Top             =   1065
      Width           =   45
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFile_Report 
         Caption         =   "ִ�б���(&E)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Exp 
         Caption         =   "��������(&E)"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuFile_Imp 
         Caption         =   "���뱨��(&I)"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuFile_ExpAll 
         Caption         =   "ȫ������"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuFile_ImpAll 
         Caption         =   "ȫ������"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuFile_PARA_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Para 
         Caption         =   "��������"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnuFile_IO_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_quit 
         Caption         =   "�˳�(&X)"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEdit_Add 
         Caption         =   "��������(&W)"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuEdit_Modi 
         Caption         =   "�޸ı���(&M)"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuEdit_Del 
         Caption         =   "ɾ������(&R)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEdit_Report_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Group_Add 
         Caption         =   "����������(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuEdit_Group_Modify 
         Caption         =   "�޸ı�����(&T)"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuEdit_Group_Delete 
         Caption         =   "ɾ��������(&O)"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu mnuEdit_Group_Setup 
         Caption         =   "�����ӱ���(&S)"
      End
      Begin VB.Menu mnuEdit_Group_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Design 
         Caption         =   "��Ʊ���(&D)"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuEdit_Clear 
         Caption         =   "�����ʷ����Դ(&C)"
      End
      Begin VB.Menu mnuEdit_Design_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Guide 
         Caption         =   "������(&G)"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuEdit_Guide_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Group_Publish 
         Caption         =   "�����鷢��(&P)"
      End
      Begin VB.Menu mnuEdit_Group_unPub 
         Caption         =   "ȡ������������(&T)"
      End
      Begin VB.Menu mnuEdit_Publish 
         Caption         =   "������(&B)"
         Begin VB.Menu mnuEdit_Publish_Main 
            Caption         =   "������̨�˵�(&1)"
         End
         Begin VB.Menu mnuEdit_Publish_Module 
            Caption         =   "��ģ���ڲ˵�(&2)"
         End
      End
      Begin VB.Menu mnuEdit_unPub 
         Caption         =   "ȡ������(&U)"
         Begin VB.Menu mnuEdit_unPub_Main 
            Caption         =   "�ӵ���̨�˵�(&1)"
         End
         Begin VB.Menu mnuEdit_unPub_Module 
            Caption         =   "��ģ���ڲ˵�(&2)"
         End
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "��׼��ť(&B)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuView_Tlb_1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "�ı���ǩ(&L)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuView_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView_View 
         Caption         =   "��ͼ��(&G)"
         Index           =   0
      End
      Begin VB.Menu mnuView_View 
         Caption         =   "Сͼ��(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuView_View 
         Caption         =   "�б�(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuView_View 
         Caption         =   "��ϸ����(&D)"
         Checked         =   -1  'True
         Index           =   3
      End
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewOnly 
         Caption         =   "����ʾ������(&O)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuView_3 
         Caption         =   "-"
      End
      Begin VB.Menu RunLog 
         Caption         =   "������־"
         Index           =   5
      End
      Begin VB.Menu mnuView_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFindNext 
         Caption         =   "������һ��"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuView_reFlash 
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
   Begin VB.Menu mnuPopPublish 
      Caption         =   "����"
      Visible         =   0   'False
      Begin VB.Menu mnuPopPublish_Group 
         Caption         =   "�����鵽����̨�˵�"
      End
      Begin VB.Menu mnuPopPublish_ReportMain 
         Caption         =   "��������̨�˵�"
      End
      Begin VB.Menu mnuPopPublish_ReportModule 
         Caption         =   "����ģ���ڲ˵�"
      End
   End
   Begin VB.Menu mnuPopUnpub 
      Caption         =   "ȡ��"
      Visible         =   0   'False
      Begin VB.Menu mnuPopUnpub_Group 
         Caption         =   "������ӵ���̨�˵�"
      End
      Begin VB.Menu mnuPopUnpub_ReportMain 
         Caption         =   "����ӵ���̨�˵�"
      End
      Begin VB.Menu mnuPopUnpub_ReportModule 
         Caption         =   "�����ģ���ڲ˵�"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnItem As Boolean
Private mobjSelItem As ListItem
Private mblnMouseDown As Boolean
Private mblnGrant As Boolean '�Ƿ�ע���˱�����ɾ����
Private mblnModule As Boolean '�Ƿ���������ģ��
Private mstrRepName As String
Private mstrPreGroup As String
Private mstrFindValue As String     '��¼��ѯ�ı����ֵ
Private mrsFind As New ADODB.Recordset
Private Enum CurSelect
    CS_������ = 0
    CS_���� = 1
End Enum
Private mcsActive As CurSelect '��ؼ�,0-�������б�1-�����б�
Private mfrmReportPara As frmReportPara
'SubItems����
Private Enum ReportCol
    RC_���� = 0 '��������
    RC_��� = 1
    RC_˵�� = 2
    RC_�޸�ʱ�� = 3
    RC_����ʱ�� = 4
    RC_���ִ��ʱ�� = 5
    RC_���ִ���� = 6
    RC_���� = 7
    RC_���� = 8
    RC_������������Դ = 9
    RC_���� = 10
    RC_������������ = 11
End Enum
'SubItems����
Private Enum GroupCol
    GC_���� = 0 '��������
    GC_��� = 1
    GC_˵�� = 2
    GC_����ʱ�� = 3
    GC_���� = 4
End Enum

Private Sub cboSys_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    mblnModule = False
    If cboSys.ListIndex <> -1 Then
        If InStr(",0,2,5,7,", cboSys.ItemData(cboSys.ListIndex) \ 100) > 0 Then
            '�������£��ɱ�����������������ģ��
            mblnModule = True
        Else
            '����ϵͳ��10�汾����
            strSQL = "Select �汾�� From zlSystems Where ���=[1]"
            Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, cboSys.ItemData(cboSys.ListIndex))
            If Not rsTmp.EOF Then
                mblnModule = Val(Split(rsTmp!�汾��, ".")(0)) >= 10
            End If
        End If
    End If
    On Error GoTo 0
    
    Call mnuView_reFlash_Click
    
    On Error Resume Next
    If lvwReport.Visible Then lvwReport.SetFocus
    On Error GoTo 0
    
    Call SetFuncEnabled(True)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbr_HeightChanged(ByVal NewHeight As Single)
    Call Form_Resize
End Sub

Private Sub GotoLVW(ByVal intCurPosition As Integer, lvwCur As ListView)
    Dim objItem As ListItem
    Dim i As Integer
    
    On Error Resume Next
    
    Set objItem = lvwCur.FindItem(mstrRepName, 0, intCurPosition, 1)
    If Not objItem Is Nothing Then
        Set lvwCur.SelectedItem = objItem
        lvwCur.ListItems(lvwCur.SelectedItem.Index).Selected = True
        lvwCur.SelectedItem.EnsureVisible
        If lvwCur.name = "lvwGroup" Then
            mstrPreGroup = ""
            Call LvwGroup_ItemClick(lvwGroup.SelectedItem)
        End If
    End If
End Sub

Private Sub cmdNext_Click()
    Call txtFind_KeyPress(vbKeyReturn)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF And Shift = vbCtrlMask Then
        txtFind.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Static strPass As String
    Static vTime As Date
    Dim blnDo As Boolean
    If Me.ActiveControl Is txtFind Then
        strPass = ""
    Else
        blnDo = cboSys.ItemData(cboSys.ListIndex) > 0
        If blnDo Then blnDo = lvwGroup.SelectedItem.Key = "_-1"
        If blnDo Then blnDo = Not lvwReport.SelectedItem Is Nothing
        If blnDo Then blnDo = InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ 0123456789", UCase(Chr(KeyAscii))) > 0
        If blnDo Then
            If DateDiff("s", vTime, Now) <= 2 Then
                strPass = strPass & Chr(KeyAscii)
            Else
                strPass = Chr(KeyAscii)
            End If
            KeyAscii = 0
            vTime = Now
            If UCase(strPass) = UCase("Publish Report") Then
                Call mnuEdit_Publish_Module_Click
                strPass = ""
            ElseIf UCase(strPass) = UCase("unPublish Report") Then
                Call mnuEdit_unPub_Module_Click
                strPass = ""
            End If
        Else
            strPass = ""
        End If
    End If
End Sub

Private Sub Form_Load()
    '��ȡ�Զ��屨����Ȩ����
    mblnModule = True
    mblnGrant = (zlRegTool() And 2) = 2
    If Not mblnGrant Then
        mnuEdit_Add.Visible = False
        mnuEdit_Del.Visible = False
        
        mnuEdit_Group_Add.Visible = False
        mnuEdit_Group_Delete.Visible = False
        
        mnuEdit_Guide.Visible = False
        mnuEdit_Guide_.Visible = False
        
        mnuEdit_Publish.Visible = False
        mnuEdit_unPub.Visible = False
                
        mnuEdit_Design_.Visible = False
        
        tbr.Buttons("Add").Visible = False
        tbr.Buttons("Del").Visible = False
        tbr.Buttons("GroupAdd").Visible = False
        tbr.Buttons("GroupDel").Visible = False
        tbr.Buttons("Guide").Visible = False
        tbr.Buttons("Guide_").Visible = False
        tbr.Buttons("Publish").Visible = False
        tbr.Buttons("unPub").Visible = False
        tbr.Buttons("Pub_").Visible = False
    End If
    
    lvwReport.ColumnHeaders(RC_��� + 1).Position = 1
    mblnMouseDown = False
    RestoreWinState Me, App.ProductName
    tbrCheck.ZOrder
    mnuViewOnly.Checked = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name, "��ʾ��", 1)

    Call ReadSystem
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long '������ռ�ø߶�s
    Dim staH As Long '״̬��ռ�ø߶�
    
    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    
    If cbr.Bands(2).MinWidth < 6615 Then cbr.Bands(2).MinWidth = 6615
    If Width < 8000 Then Width = 8000
    If Height < 5000 Then Height = 5000
    
    '����ؼ���Ⱥ͸߶�
    cbrH = IIF(cbr.Visible, cbr.Height, 0)
    staH = IIF(sta.Visible, sta.Height, 0)
    With ImgSplit_S
        .Top = ScaleTop + cbrH
        .Height = ScaleHeight - cbrH - staH
    End With
    
    With lblGroup
        .Left = 0
        .Top = ImgSplit_S.Top + 30
        .Width = ImgSplit_S.Left
    End With
    
    With lvwGroup
        .Left = 0
        .Top = lblGroup.Top + lblGroup.Height + 30
        .Width = ImgSplit_S.Left
        .Height = ImgSplit_S.Top + ImgSplit_S.Height - .Top
    End With
    
    With LblReport
        .Left = ImgSplit_S.Left + ImgSplit_S.Width
        .Top = ImgSplit_S.Top + 30
        .Width = ScaleWidth - .Left
    End With
    
    With lvwReport
        .Left = ImgSplit_S.Left + ImgSplit_S.Width
        .Top = LblReport.Top + LblReport.Height + 30
        .Width = ScaleWidth - .Left
        .Height = ImgSplit_S.Top + ImgSplit_S.Height - .Top
    End With
End Sub

Private Sub Sub�鿴�˵�(ByVal mnuLable As String)
    Dim i As Integer
    
    Select Case mnuLable
        Case "��׼��ť(&B)"
            mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
            mnuViewToolText.Enabled = Not mnuViewToolText.Enabled
            cbr.Visible = Not cbr.Visible
            Form_Resize
        Case "�ı���ǩ(&L)"
            mnuViewToolText.Checked = Not mnuViewToolText.Checked
            For i = 1 To tbr.Buttons.count
                If mnuViewToolText.Checked Then
                    tbr.Buttons(i).Caption = tbr.Buttons(i).Tag
                Else
                    tbr.Buttons(i).Caption = ""
                End If
            Next
            cbr.Bands(1).MinHeight = tbr.ButtonHeight
            Form_Resize
        Case "״̬��(&S)"
            mnuViewStatus.Checked = Not mnuViewStatus.Checked
            sta.Visible = Not sta.Visible
            Form_Resize
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub ImgSplit_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        With ImgSplit_S
            If .Left + X < 1500 Or Me.ScaleWidth - .Left - X < 2000 Then Exit Sub
            .Move .Left + X
        End With
        Form_Resize
    End If
End Sub


Private Sub lvwGroup_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuEdit, 2
End Sub

Private Sub lvwReport_DblClick()
    mcsActive = CS_����
    lblFind.Caption = "���ұ���(&F)"
    If mblnItem Then mnuEdit_Design_Click
End Sub

Private Sub lvwReport_GotFocus()
    mcsActive = CS_����
    lblFind.Caption = "���ұ���(&F)"
    Call SetFuncEnabled(True)
End Sub

Private Sub lvwReport_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mcsActive = CS_����
    lblFind.Caption = "���ұ���(&F)"
    Item.Selected = True '���Զ�ѡʱ�����д��仰��SelectedItem ΪNothing
    Call SetFuncEnabled(True)
    
    If Item.SubItems(RC_����ʱ��) <> "" Then
        sta.Panels(2) = Item.Text & "λ��:" & GetMenuPath(Val(Mid(Item.Key, 2)))
    Else
        sta.Panels(2) = "[" & Item.SubItems(RC_���) & "]" & Item.Text & IIF(Item.SubItems(RC_˵��) = "", "", ":" & Item.SubItems(RC_˵��))
    End If
    
End Sub

Private Sub lvwReport_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    If KeyCode = vbKeyDelete Then
        mnuEdit_Del_Click
    ElseIf KeyCode = vbKeyA And Shift = vbCtrlMask Then
        For i = 1 To lvwReport.ListItems.count
            lvwReport.ListItems(i).Selected = True
        Next
    End If
End Sub

Private Sub lvwReport_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Not lvwReport.SelectedItem Is Nothing Then
        mnuEdit_Modi_Click
    End If
End Sub

Private Sub lvwReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnMouseDown = False
    If lvwReport.HitTest(X, Y) Is Nothing Then
        mblnItem = False
        If Button = 1 Then sta.Panels(2) = "�� " & lvwReport.ListItems.count & " �ű���"
    Else
        mblnItem = True
        mblnMouseDown = (Button = 1) And (cboSys.Text = "����ϵͳ����")
    End If
End Sub

Private Sub lvwReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnMouseDown = False
    Set lvwReport.DragIcon = Nothing
    lvwReport.Drag 0
    
    If Button = 2 Then
        If Not mblnItem Then
            PopupMenu mnuView, 2
        Else
            PopupMenu mnuEdit, 2
        End If
    End If
End Sub

Private Sub LvwGroup_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnDesc As Boolean
    Static intIdx As Integer
    mcsActive = CS_������
    lblFind.Caption = "���ҷ���(&F)"
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvwGroup.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvwGroup.SortOrder = lvwDescending
    Else
        lvwGroup.SortOrder = lvwAscending
    End If
    lvwGroup.Sorted = True
    intIdx = ColumnHeader.Index
    
    If Not lvwGroup.SelectedItem Is Nothing Then lvwGroup.SelectedItem.EnsureVisible
End Sub

Private Sub LvwGroup_DblClick()
    mcsActive = CS_������
    lblFind.Caption = "���ҷ���(&F)"
    mnuEdit_Group_Modify_Click
End Sub

Private Sub LvwGroup_DragDrop(Source As Control, X As Single, Y As Single)
    Dim rsInsert As New ADODB.Recordset
    Dim rsGetGroups As New ADODB.Recordset
    Dim objLastSel As ListItem, objCurSel As ListItem
    Dim lngReportID As Long, intRptCount As Integer
    Dim blnInsert As Boolean, strSQL As String
    
    With lvwGroup
        If .SelectedItem.Key = "_-1" Then Exit Sub
                        
        strSQL = "Select 1 From zlRPTSubs A,zlReports B Where B.����=[1] And A.����ID=B.ID And A.��ID=[2]"
        Set rsGetGroups = OpenSQLRecord(strSQL, Me.Caption, Source.SelectedItem.Text, Val(Mid(lvwGroup.SelectedItem.Key, 2)))
        If Not rsGetGroups.EOF Then
            MsgBox "�ñ��������Ѿ�������ͬ���Ƶı���", vbInformation, App.Title
            lvwGroup.ListItems("_-1").Selected = True: Exit Sub
        End If
        
        intRptCount = 1
        strSQL = "Select Count(*) Records From zlRPTSubs Where ��ID=[1]"
        Set rsGetGroups = OpenSQLRecord(strSQL, Me.Caption, Val(Mid(lvwGroup.SelectedItem.Key, 2)))
        If Not rsGetGroups.EOF Then
            intRptCount = Nvl(rsGetGroups!Records, 0) + 1
        End If
    End With
    
    Set objLastSel = mobjSelItem
    Set objCurSel = lvwGroup.SelectedItem
    'ɾ����ǰ�����϶����ӱ�,�����뵽����
    If objLastSel.Key <> "_-1" Then
        '�޸����
        gcnOracle.Execute "Update zlRPTSubs Set ���=���-1 Where ���>(Select ��� From zlRPTSubs Where ��ID=" & Mid(objLastSel.Key, 2) & " And ����ID=" & Mid(Source.SelectedItem.Key, 2) & ") And ��ID=" & Mid(objLastSel.Key, 2)
        gcnOracle.Execute "Delete zlRPTSubs Where ��ID=" & Mid(objLastSel.Key, 2) & " And ����ID=" & Mid(Source.SelectedItem.Key, 2)
    End If
    
    blnInsert = True
    strSQL = "Select Count(*) Records From zlRPTSubs Where ��ID=[1] And ����ID=[2]"
    Set rsInsert = OpenSQLRecord(strSQL, Me.Caption, Val(Mid(objCurSel.Key, 2)), Val(Mid(Source.SelectedItem.Key, 2)))
    If Not rsInsert.EOF Then
        blnInsert = Nvl(rsInsert!Records, 0) = 0
    End If

    If blnInsert Then gcnOracle.Execute "Insert Into zlRPTSubs(��ID,����ID,���,����) Values(" & Mid(objCurSel.Key, 2) & "," & Mid(Source.SelectedItem.Key, 2) & "," & intRptCount & ",'" & Source.SelectedItem.Text & "')"
    If Not mobjSelItem Is Nothing Then Set lvwGroup.SelectedItem = mobjSelItem
    mstrPreGroup = ""
    Call LvwGroup_ItemClick(lvwGroup.SelectedItem)
    
    '�����ѷ����������Ȩ��
    If Val(objCurSel.Tag) <> 0 Then Call ReportGrantToNavigatorAgain(objCurSel)
    If objCurSel.Key <> objLastSel.Key Then
        If Val(objLastSel.Tag) <> 0 Then
            strSQL = "Select Count(*) Records from zlRPTSubs Where ��ID=[1]"
            Set rsGetGroups = OpenSQLRecord(strSQL, Me.Caption, Val(Mid(objLastSel.Key, 2)))
            If Not rsGetGroups.EOF Then
                If rsGetGroups!Records = 0 Then
                    'ȡ���ñ�����ķ���
                    Call ReportRevokeFromNavigator(True)
                Else
                    Call ReportGrantToNavigatorAgain(objLastSel)
                End If
            Else
                Call ReportRevokeFromNavigator(True)
            End If
        End If
    End If
End Sub

Private Sub LvwGroup_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Dim objTest As ListItem
    
    Set objTest = lvwGroup.HitTest(X, Y)
    If Not objTest Is Nothing Then
        If objTest.Key = "_-1" Then Exit Sub
        objTest.Selected = True
    End If
End Sub

Private Sub LvwGroup_GotFocus()
    Call SetFuncEnabled(False)
End Sub

Private Sub LvwGroup_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mcsActive = CS_������
    lblFind.Caption = "���ҷ���(&F)"
    Call SetFuncEnabled(False)
    
    Set mobjSelItem = Item
    
    If Item.Key <> mstrPreGroup Then
        If Not ReadReports(Mid(Item.Key, 2)) Then
            MsgBox "�����ȡʧ�ܣ�", vbInformation, App.Title
            Exit Sub
        End If
        mstrPreGroup = Item.Key
    End If
    sta.Panels(2) = "�� " & lvwReport.ListItems.count & " �ű���"
    If Item.SubItems(GC_����ʱ��) <> "" Then
        sta.Panels(2) = sta.Panels(2) & "," & Item.Text & "λ��:" & GetMenuPath(Val(Mid(Item.Key, 2)), True)
    End If
End Sub

Private Sub LvwGroup_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then LvwGroup_DblClick
End Sub

Private Sub LvwGroup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strPath As String
    Static objItem As Object
    
    If Not objItem Is Nothing And Not lvwGroup.HitTest(X, Y) Is Nothing Then
        If objItem.Key = lvwGroup.HitTest(X, Y).Key Then Exit Sub
    End If
    
    Set objItem = lvwGroup.HitTest(X, Y)
    lvwGroup.ToolTipText = ""
End Sub

Private Sub mnuEdit_Add_Click()
    Dim strName As String, lngID As Long, str���� As String, str˵�� As String
    
    If cboSys.ItemData(cboSys.ListIndex) <> 0 Then cboSys.ListIndex = 0
    If frmReportEdit.ShowMe(Me, cboSys.ItemData(cboSys.ListIndex), False, Val(lvwGroup.SelectedItem.Tag), IIF(lvwGroup.SelectedItem.Key = "_-1", 0, Mid(lvwGroup.SelectedItem.Key, 2)), _
                                                lngID, strName, str����, str˵��) Then
        mstrPreGroup = ""
        Call AfterItemEdit(True, False, lngID, strName, str����, str˵��)
        Call LvwGroup_ItemClick(lvwGroup.SelectedItem)
    Else
        Call CustomToolBarRefresh
    End If
End Sub

Private Sub mnuEdit_Clear_Click()
    frmClearHistory.Show 1, Me
End Sub

Private Sub mnuEdit_Del_Click()
    Dim rsCheck As New ADODB.Recordset
    Dim rsGetGroups As New ADODB.Recordset
    Dim intIdx As Integer, strSQL As String
    
    If lvwReport.SelectedItem Is Nothing Then
        MsgBox "��ǰû�б������ɾ����", vbInformation, App.Title: Exit Sub
    End If
    If cboSys.ItemData(cboSys.ListIndex) > 0 Then Exit Sub
    
    intIdx = lvwReport.SelectedItem.Index
    If lvwGroup.SelectedItem.Key = "_-1" Then
        '����Ƿ����ڱ����飬�Ƿ�����ɾ��
        strSQL = "Select ID ��ID,���� From zlRPTGroups Where ID=(Select ��ID From zlRPTSubs Where ����ID=[1])"
        Set rsCheck = OpenSQLRecord(strSQL, Me.Caption, Val(Mid(lvwReport.SelectedItem.Key, 2)))
        If Not rsCheck.EOF Then
            If Not IsNull(rsCheck!��ID) Then
                MsgBox "���Ȱѱ���[" & lvwReport.SelectedItem & "]�ӱ�����[" & rsCheck!���� & "]���Ƴ�����ɾ����", vbInformation, App.Title
                Exit Sub
            End If
        End If
        '����Ƿ��ѷ���
        If Val(lvwReport.SelectedItem.Tag) <> 0 Then
            MsgBox "�ñ����Ѿ�����,����ȡ����������ɾ����", vbInformation, App.Title: Exit Sub
        End If
        strSQL = "Select ����ID From zlRPTPuts Where ����ID=[1]"
        Set rsCheck = OpenSQLRecord(strSQL, Me.Caption, Val(Mid(lvwReport.SelectedItem.Key, 2)))
        If Not rsCheck.EOF Then
            MsgBox "�ñ����Ѿ�����,����ȡ����������ɾ����", vbInformation, App.Title: Exit Sub
        End If
        
        If MsgBox("ȷʵҪɾ������[" & lvwReport.SelectedItem & "]��", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
        On Error GoTo errH
        gcnOracle.BeginTrans
        gcnOracle.Execute "Delete From zlReports Where ID=" & Mid(lvwReport.SelectedItem.Key, 2)
        gcnOracle.CommitTrans
        On Error GoTo 0
    Else
        If MsgBox("��ȷ��Ҫ�ӱ�����[" & lvwGroup.SelectedItem & "]���Ƴ�����[" & lvwReport.SelectedItem & "]��", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
        On Error GoTo errH
        gcnOracle.BeginTrans
        gcnOracle.Execute "Update zlRPTSubs Set ���=���-1 Where ���>(Select ��� From zlRPTSubs Where ����ID=" & Mid(lvwReport.SelectedItem.Key, 2) & " And ��ID=" & Mid(lvwGroup.SelectedItem.Key, 2) & ") And ��ID=" & Mid(lvwGroup.SelectedItem.Key, 2)
        gcnOracle.Execute "Delete From zlRPTSubs Where ����ID=" & Mid(lvwReport.SelectedItem.Key, 2) & " And ��ID=" & Mid(lvwGroup.SelectedItem.Key, 2)
        gcnOracle.CommitTrans
        On Error GoTo 0
    End If
    
    lvwReport.ListItems.Remove intIdx
    
    If lvwReport.ListItems.count <> 0 Then
        If intIdx <= lvwReport.ListItems.count Then
            lvwReport.ListItems(intIdx).Selected = True
        Else
            lvwReport.ListItems(lvwReport.ListItems.count).Selected = True
        End If
        lvwReport.SelectedItem.EnsureVisible
        Call lvwReport_ItemClick(lvwReport.SelectedItem)
    Else
        sta.Panels(2) = "�� 0 �ű���"
    End If
    
    '�����ѷ����������Ȩ��
    If Val(lvwGroup.SelectedItem.Tag) <> 0 Then
        strSQL = "Select Count(*) Records from zlRPTSubs Where ��ID=[1]"
        Set rsGetGroups = OpenSQLRecord(strSQL, Me.Caption, Val(Mid(lvwGroup.SelectedItem.Key, 2)))
        If Not rsGetGroups.EOF Then
            If rsGetGroups!Records = 0 Then
                'ȡ���ñ�����ķ���
                Call ReportRevokeFromNavigator(True)
            Else
                Call ReportGrantToNavigatorAgain(lvwGroup.SelectedItem)
            End If
        Else
            Call ReportRevokeFromNavigator(True)
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub

Private Sub mnuEdit_Design_Click()
    Dim lngIndex As Long, i As Long
    
    If lvwReport.SelectedItem Is Nothing Then
        MsgBox "��ǰû�б��������ƣ�", vbInformation, App.Title: Exit Sub
    End If
    If CheckPass(CLng(Mid(lvwReport.SelectedItem.Key, 2))) = False Then
        MsgBox "�������ݴ��󣬲�����Ƹñ���", vbInformation, App.Title: Exit Sub
    End If
    If Not CheckReportPriv(CLng(Mid(lvwReport.SelectedItem.Key, 2))) Then
        MsgBox "��û��Ȩ�޲�ѯ�ñ���ĳЩ����Դ�еĶ���,������ƻ�����������", vbInformation, App.Title
    End If
    
    glngSys = cboSys.ItemData(cboSys.ListIndex)
    frmDesign.lngRPTID = CLng(Mid(lvwReport.SelectedItem.Key, 2))
    
    On Error Resume Next
    frmDesign.Show 1, Me
    On Error GoTo 0
    
    If gblnModi Then
        '����ѡ���ѡ�������
        lngIndex = lvwReport.SelectedItem.Index
        Call ReadGroups
        If lngIndex > lvwReport.ListItems.count Then
            lngIndex = lvwReport.ListItems.count
        End If
        For i = 1 To lvwReport.ListItems.count
            If lvwReport.ListItems(i).Selected Then
                lvwReport.ListItems(i).Selected = False
            End If
        Next
        lvwReport.ListItems(lngIndex).Selected = True
    End If
End Sub

Private Sub mnuEdit_Group_Add_Click()
    Dim strName As String, lngID As Long, str���� As String, str˵�� As String
    If frmReportEdit.ShowMe(Me, cboSys.ItemData(cboSys.ListIndex), True, 0, lngID, , strName, str����, str˵��) Then
        Call AfterItemEdit(True, True, lngID, strName, str����, str˵��)
        Call mnuView_reFlash_Click
    End If
End Sub

Private Sub mnuEdit_Group_Delete_Click()
    If lvwGroup.SelectedItem Is Nothing Then
        MsgBox "��ǰû�б��������ɾ����", vbInformation, App.Title: Exit Sub
    End If
    If lvwGroup.SelectedItem.Key = "_-1" Then
        MsgBox "��ǰû�б��������ɾ����", vbInformation, App.Title: Exit Sub
    End If
    If lvwGroup.SelectedItem.Icon = 3 Then
        MsgBox "ϵͳ���еı����鲻��ɾ����", vbInformation, App.Title: Exit Sub
    End If
    If Val(lvwGroup.SelectedItem.Tag) <> 0 Then
        MsgBox "��ȡ���ñ�����ķ��������ԣ�", vbInformation, App.Title: Exit Sub
    End If
    If MsgBox("��ȷ��Ҫɾ��������[" & lvwGroup.SelectedItem & "]��", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
    
    On Error GoTo ErrHand
    gcnOracle.BeginTrans
    gcnOracle.Execute "Delete zlRPTSubs Where ��ID=" & Mid(lvwGroup.SelectedItem.Key, 2)
    gcnOracle.Execute "Delete zlRPTGroups Where ID=" & Mid(lvwGroup.SelectedItem.Key, 2)
    gcnOracle.CommitTrans
    
    mnuView_reFlash_Click
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    gcnOracle.RollbackTrans
End Sub

Private Sub mnuEdit_Group_Modify_Click()
    Dim lngSys As Long
    Dim strName As String, lngID As Long, str���� As String, str˵�� As String
    
    If lvwGroup.SelectedItem Is Nothing Then
        MsgBox "��ǰû�б���������޸ģ�", vbInformation, App.Title: Exit Sub
    End If
    If lvwGroup.SelectedItem.Key = "_-1" Then Exit Sub
    lngSys = cboSys.ItemData(cboSys.ListIndex)
    lngID = CLng(Mid(lvwGroup.SelectedItem.Key, 2))
    str���� = lvwGroup.SelectedItem.SubItems(GC_���)
    strName = lvwGroup.SelectedItem.Text
    str˵�� = lvwGroup.SelectedItem.SubItems(GC_˵��)
    If frmReportEdit.ShowMe(Me, lngSys, True, Val(lvwGroup.SelectedItem.Tag), lngID, , strName, str����, str˵��) Then
        Call AfterItemEdit(False, True, lngID, strName, str����, str˵��)
        mnuView_reFlash_Click
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    gcnOracle.RollbackTrans
    Call SaveErrLog
    Unload frmReportEdit
End Sub

Private Sub mnuEdit_Group_Publish_Click()
    Call ReportGrantToNavigator
End Sub

Private Sub mnuEdit_Group_Setup_Click()
    Dim rsGetGroups As New ADODB.Recordset
    Dim strSQL As String
    
    '������Щ�������ڸñ�����
    If lvwGroup.SelectedItem Is Nothing Then Exit Sub
    If lvwGroup.SelectedItem.Key = "_-1" Then Exit Sub
    
    With frmSetGroup
        .LngGroupID = Mid(lvwGroup.SelectedItem.Key, 2)
        .strCaption = "���ñ�����[" & lvwGroup.SelectedItem & "]�Ĵ�������"
        .Show 1, Me
    End With
    
    '�����ѷ����������Ȩ��
    If Val(lvwGroup.SelectedItem.Tag) <> 0 Then
        strSQL = "Select Count(*) Records From zlRPTSubs Where ��ID=[1]"
        Set rsGetGroups = OpenSQLRecord(strSQL, Me.Caption, Val(Mid(lvwGroup.SelectedItem.Key, 2)))
        If Not rsGetGroups.EOF Then
            If rsGetGroups!Records = 0 Then
                'ȡ���ñ�����ķ���
                Call ReportRevokeFromNavigator(True)
            Else
                Call ReportGrantToNavigatorAgain(lvwGroup.SelectedItem)
            End If
        Else
            Call ReportRevokeFromNavigator(True)
        End If
    End If
    
    mstrPreGroup = ""
    Call LvwGroup_ItemClick(lvwGroup.SelectedItem)
End Sub

Private Sub mnuEdit_Group_unPub_Click()
    Call ReportRevokeFromNavigator
End Sub

Private Sub mnuEdit_Modi_Click()
    Dim lngSys As Long, lngRPTID As Long
    Dim strName As String, lngID As Long, str���� As String, str˵�� As String
    
    If lvwReport.SelectedItem Is Nothing Then
        MsgBox "��ǰû�б�������޸ģ�", vbInformation, App.Title: Exit Sub
    End If
    lngSys = cboSys.ItemData(cboSys.ListIndex)
    lngRPTID = CLng(Mid(lvwReport.SelectedItem.Key, 2))
    lngID = IIF(lvwGroup.SelectedItem.Key = "_-1", 0, Mid(lvwGroup.SelectedItem.Key, 2))
    str���� = lvwReport.SelectedItem.SubItems(RC_���)
    strName = lvwReport.SelectedItem.Text
    str˵�� = lvwReport.SelectedItem.SubItems(RC_˵��)
    If frmReportEdit.ShowMe(Me, lngSys, False, Val(lvwReport.SelectedItem.Tag), lngID, _
                                                lngRPTID, strName, str����, str˵��) Then
        Call AfterItemEdit(False, False, lngRPTID, strName, str����, str˵��)
        mstrPreGroup = ""
        Call LvwGroup_ItemClick(lvwGroup.SelectedItem)
    Else
        Call CustomToolBarRefresh
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    gcnOracle.RollbackTrans
    Call SaveErrLog
    Unload frmReportEdit
End Sub

Private Sub mnuEdit_Publish_Module_Click()
    Call ReportGrantToModule
End Sub

Private Sub mnuEdit_unPub_Module_Click()
    Call ReportRevokeFromModule
End Sub

Private Sub mnuFile_ExpAll_Click()
    Dim rsReportInfo As New ADODB.Recordset, strSQL As String
    Dim strPath As String, strFile As String, strPathTmp As String
    Dim i As Long, j As Long, lngCount As Long, lngExp As Long
    Dim objFile As New FileSystemObject
    
    strPath = BrowseForFolder(Me.hwnd, "ѡ�񱨱���Ŀ¼", strPath)
    If strPath <> "" Then
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\Path", "Export", strPath
        strSQL = "Select A.Id, A.���, A.����, C.Id ��id, C.��� ����, C.���� ����" & vbNewLine & _
                    "From zlReports A, zlRPTSubs B, zlRPTGroups C" & vbNewLine & _
                    "Where A.Id = B.����id(+) And B.��id = C.Id(+)  And " & IIF(cboSys.ItemData(cboSys.ListIndex) = 0, " A.ϵͳ Is Null ", " A.ϵͳ=[1] ") & vbNewLine & _
                    "Order By A.���"
        Set rsReportInfo = OpenSQLRecord(strSQL, Me.Caption, cboSys.ItemData(cboSys.ListIndex))
        lngCount = rsReportInfo.RecordCount
        If MsgBox("���ι����� " & cboSys.List(cboSys.ListIndex) & lngCount & " �ű��� " & strPath & "��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
        lngExp = 0
        For i = 1 To lvwGroup.ListItems.count
            '���б���
            If Val(Mid(lvwGroup.ListItems(i).Key, 2)) = -1 Then
                rsReportInfo.Filter = "��id=Null"
                strPathTmp = strPath
            Else
                rsReportInfo.Filter = "��id=" & Val(Mid(lvwGroup.ListItems(i).Key, 2))
                strPathTmp = strPath & "\[" & lvwGroup.ListItems(i).SubItems(GC_���) & "]" & lvwGroup.ListItems(i).Text
                If Not objFile.FolderExists(strPathTmp) Then
                    Call objFile.CreateFolder(strPathTmp)
                End If
            End If
            For j = 1 To rsReportInfo.RecordCount
                lngExp = lngExp + 1
                Call ShowFlash("���ڵ���:" & rsReportInfo!���� & ".ZLR", lngExp / lngCount, Me, True)
                strFile = "[" & rsReportInfo!��� & "]" & rsReportInfo!���� & ".ZLR"
                If Not ExportReport(Val(rsReportInfo!ID & ""), strPathTmp & "\" & strFile) Then
                    Call ShowFlash
                    If MsgBox("��������ʱ���ִ���Ҫ����������һ�ű�����", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
                End If
                rsReportInfo.MoveNext
            Next
        Next
        Call ShowFlash
    End If
End Sub

Private Sub mnuFile_ImpAll_Click()
    Dim strPath As String, objFSO As New FileSystemObject, objFile As File, objFolder As Folder
    Dim lngSys As Long, lngCurGroup As Long
    Dim rsFiles As ADODB.Recordset
    Dim arrTmp As Variant, strFile As String, i As Long
    Dim rsGroups As ADODB.Recordset, strName As String, strCode As String, strSQL As String
    Dim LngGroupID As Long, lngReportID As Long
    
    On Error GoTo errH
    strPath = BrowseForFolder(Me.hwnd, "ѡ����Ҫ���뱨������Ŀ¼", strPath)
    If strPath <> "" Then
        If MsgBox("�Ƿ���""" & strPath & """�ļ��м����ļ����µ����б���", vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then Exit Sub
        lngSys = cboSys.ItemData(cboSys.ListIndex)
        lngCurGroup = IIF(lvwGroup.SelectedItem.Key = "_-1", 0, Val(Mid(lvwGroup.SelectedItem.Key, 2)))
        If Not lvwReport.SelectedItem Is Nothing Then lngReportID = CLng(Mid(lvwReport.SelectedItem.Key, 2))
        'FilePath=����ȫ·����FileName=�����ļ�������ID=����Ҫ����ı�����ID
        'ͬ��ID=�뽫Ҫ����ı���ͬ���ı���ı���ID���̶�����ͨ������ƥ�䣬�ǹ̶�ͨ������ƥ��
        '��������=0-�����룬1-��������,2-���ǵ���;��������=0-���帲�ǣ�1-������Դ����
        'ErrType=0-�޴���,1-�����ͬ����һ��������2-�����ͬ����һ�𸲸ǣ�3-ϵͳ����ֻ�ܸ��ǣ�������ͬ������
        '                            4-���ݴ�������,5-�汾��������,6-���Ʊ�Ŵ�������
        'ImportResult=-1-�Ѿ��ɹ����뵫�Ǳ��������δͨ����0-������,1-����ɹ�,2-����ʧ��
        'ImportInfo=����ɹ�����󷵻صı�����Ϣ
        Set rsFiles = CopyNewRec(Nothing, , True, _
                                    Array("FilePath", adVarChar, 1000, Empty, "FileName", adVarChar, 200, Empty, "��ID", adBigInt, Empty, Empty, _
                                             "ͬ��ID", adBigInt, Empty, Empty, "��������", adInteger, Empty, Empty, "��������", adInteger, Empty, Empty, _
                                             "ErrType", adInteger, Empty, Empty, "ImportResult", adInteger, Empty, Empty, "ImportInfo", adVarChar, 200, Empty))
        
        
        With rsFiles
            '�Ѽ����뵽���б����еĵı���,����ǰ�ļ����µı���
            For Each objFile In objFSO.GetFolder(strPath).Files
                If UCase(objFile.name) Like "*.ZLR" Then
                    rsFiles.AddNew Array("FilePath", "FileName", "��ID", "ͬ��ID", "��������", "��������", "ErrType", "ImportResult", "ImportInfo"), _
                                            Array(objFile.Path, objFile.name, 0, 0, 0, 0, 0, 0, "")
                End If
            Next
            '����Ҫ�����Զ��屨��ķ���
            '�̶��������ڱ���Ψһ�ԣ��Ѿ�ȷ������
            If lngSys = 0 Then
                strSQL = "Select ID,���,���� From zlRPTGroups  Where  ϵͳ Is Null"
                Set rsGroups = CopyNewRec(OpenSQLRecord(strSQL, Me.Caption))
            End If
            '�Ѽ���ǰ�ļ��µ��Ӽ��ļ���
            For Each objFolder In objFSO.GetFolder(strPath).SubFolders
                strFile = ""
                For Each objFile In objFolder.Files
                    If UCase(objFile.name) Like "*.ZLR" Then
                        strFile = strFile & "|" & objFile.name
                    End If
                Next
                If strFile <> "" Then
                    arrTmp = Split(Mid(strFile, 2), "|")
                    LngGroupID = 0
                    '���Զ�������Ҫ���ҷ��飬�̶��������ϵͳ�ű���ȷ������
                    If lngSys = 0 Then
                        Call SplitNameCode(objFolder.name, strName, strCode)
                        rsGroups.Filter = "���='" & strCode & "'" '���Ψһ��
                        If rsGroups.EOF Then rsGroups.Filter = "����='" & strName & "'"  '�����ӷ���û�б���
                        If Not rsGroups.EOF Then
                            LngGroupID = Val(rsGroups!ID & "")
                        Else '���ɾ����Եı�����
                            '���������ƹ淶�����������µı�������
                            LngGroupID = GetNextID("zlRPTGroups")
                            If TLen(strName) > 30 Then strName = ConvertSBC(MidB(strName, 1, 30))
                            If strCode <> "" Then
                                If TLen(strCode) > 20 Then strCode = ConvertSBC(MidB(strCode, 1, 20))
                                If CheckExist("zlRPTGroups", "���", strCode) Then
                                    strCode = GetNextNO(True)
                                End If
                            Else
                                strCode = GetNextNO(True)
                            End If
                            strSQL = "Insert Into zlRPTGroups(ID,���,����,˵��) Values(" & LngGroupID & ",'" & strCode & "','" & strName & "',Null)"
                            On Error Resume Next
                            gcnOracle.Execute strSQL
                            If Err.Number <> 0 Then
                                LngGroupID = 0 '���ɱ�����ʧ�ܣ����Զ����÷����µı����뵽��������
                            Else '���ɷ���ɹ������뵽����Ϣ������
                                rsGroups.AddNew Array("ID", "���", "����"), Array(LngGroupID, strCode, strName)
                            End If
                            On Error GoTo errH
                        End If
                    End If
                    For i = LBound(arrTmp) To UBound(arrTmp)
                        rsFiles.AddNew Array("FilePath", "FileName", "��ID", "ͬ��ID", "��������", "��������", "ErrType", "ImportResult", "ImportInfo"), _
                                                Array(objFolder.Path & "\" & arrTmp(i), arrTmp(i), LngGroupID, 0, 0, 0, 0, 0, "")
                    
                    Next
                End If
            Next
            .Filter = "": .Sort = "��ID"
            If .RecordCount = 0 Then
                MsgBox "��ǰ·����δ�ҵ��κοɵ���ı���", vbInformation, App.Title
                Exit Sub
            End If
            Call ImportReportBeach(lngSys, lngCurGroup, lngReportID, rsFiles, True)
        End With
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SplitNameCode(ByVal strInput As String, ByRef strName As String, ByRef strCode As String)
'����:�ָ��������
'������strInput=������ַ����������ʽΪ[����]����,���Զ��ָ����Ĭ��Ϊֻ��ȡ������
'���أ�strName=����
'           strCode=����
    Dim arrTmp As Variant
    Dim strTmp As Variant
    If InStr(strInput, "\") > 0 Then
        strTmp = strReverse(strInput)
        strInput = strReverse(Mid(strTmp, 1, InStr(strTmp, "\") - 1))
    End If
    
    If strInput Like "[[]?*[]]?*" Then '���Ϲ淶���ļ���
        arrTmp = Split(strInput, "]")
        strName = arrTmp(1)
        strCode = Mid(arrTmp(0), 2)
    Else
        strName = strInput
        strCode = ""
    End If
End Sub

Private Sub mnuFile_Para_Click()
    '�򿪲�������
    If mfrmReportPara Is Nothing Then
        Set mfrmReportPara = New frmReportPara
    End If
    If mfrmReportPara.ShowMe(Me) Then
        '���²���
        Call InitPar
    End If
End Sub

Private Sub mnuFile_Report_Click()
    Dim objCheck As ListItem
    
    If Me.ActiveControl.name = "lvwReport" Or lvwGroup.SelectedItem.Key = "_-1" Then
        If lvwReport.SelectedItem Is Nothing Then MsgBox "��ǰû�п�ִ�еı���", vbInformation, App.Title: Exit Sub
        If Not CheckReportPriv(CLng(Mid(lvwReport.SelectedItem.Key, 2))) Then
            MsgBox "��û��Ȩ�޲�ѯ�ñ���ĳЩ����Դ�еĶ���", vbInformation, App.Title: Exit Sub
        End If
    Else
        For Each objCheck In lvwReport.ListItems
            If Not CheckReportPriv(CLng(Mid(objCheck.Key, 2))) Then
                MsgBox "��û��Ȩ�޲�ѯ����[" & objCheck.Text & "]��ĳЩ����Դ�Ķ���", vbInformation, App.Title: Exit Sub
            End If
        Next
    End If
    
    If Not (Me.ActiveControl.name = "lvwReport" Or lvwGroup.SelectedItem.Key = "_-1") Then
        'ִ�б�����
        Set gobjReport = Nothing
        glngGroup = CLng(Mid(lvwGroup.SelectedItem.Key, 2))
    Else
        'ִ�б���
        If CheckPass(CLng(Mid(lvwReport.SelectedItem.Key, 2))) = False Then
            MsgBox "�������ݴ��󣬲���ִ�иñ���", vbInformation, App.Title: Exit Sub
        End If
        
        glngGroup = 0
        Set gobjReport = Nothing
        Set gobjReport = ReadReport(CLng(Mid(lvwReport.SelectedItem.Key, 2)))
    End If
    
    glngSys = cboSys.ItemData(cboSys.ListIndex)
    garrPars = Array() 'ʹ��ȱʡ����
    If Not ShowReport(Me) Then MsgBox "�����ʧ�ܣ�", vbInformation, App.Title
End Sub

Private Sub mnuFile_Quit_Click()
    Unload Me
End Sub

Private Sub mnuFindNext_Click()
    Call txtFind_KeyPress(vbKeyReturn)
End Sub

Private Sub mnuHelpAbout_Click()
    Call ShowAbout(Me)
End Sub

Private Sub mnuHelpTitle_Click()
    Call ShowHelpRpt(Me.hwnd, "main", 0)
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage hwnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo hwnd
End Sub

Private Sub mnuPopPublish_Group_Click()
    Call mnuEdit_Group_Publish_Click
End Sub

Private Sub mnuPopPublish_ReportMain_Click()
    Call mnuEdit_Publish_Main_Click
End Sub

Private Sub mnuPopPublish_ReportModule_Click()
    Call mnuEdit_Publish_Module_Click
End Sub

Private Sub mnuPopUnpub_Group_Click()
    Call mnuEdit_Group_unPub_Click
End Sub

Private Sub mnuPopUnpub_ReportMain_Click()
    Call mnuEdit_unPub_Main_Click
End Sub

Private Sub mnuPopUnpub_ReportModule_Click()
    Call mnuEdit_unPub_Module_Click
End Sub

Private Sub mnuView_reFlash_Click()
    Call ReadGroups
End Sub

Private Sub mnuViewOnly_Click()
    mnuViewOnly.Checked = mnuViewOnly.Checked Xor True
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name, "��ʾ��", mnuViewOnly.Checked
    If lvwGroup.SelectedItem.Key = "_-1" Then
        mstrPreGroup = ""
        Call LvwGroup_ItemClick(lvwGroup.ListItems("_-1"))
    End If
End Sub

Private Sub mnuViewStatus_Click()
    Sub�鿴�˵� mnuViewStatus.Caption
End Sub

Private Sub mnuViewToolButton_Click()
    Sub�鿴�˵� mnuViewToolButton.Caption
End Sub

Private Sub mnuViewToolText_Click()
    Sub�鿴�˵� mnuViewToolText.Caption
End Sub

Private Sub picSysFind_Resize()
    txtFind.Top = (picSysFind.Height - txtFind.Height) / 2
    cboSys.Top = txtFind.Top
    lblSys.Top = (picSysFind.Height - lblSys.Height) / 2
    lblFind.Top = lblSys.Top
End Sub

Private Sub RunLog_Click(Index As Integer)
    If Not lvwReport.SelectedItem Is Nothing Then
        Call ShowRunLog
    End If
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuFile_Quit_Click
        Case "View"
            If Me.ActiveControl.name = "lvwGroup" Then
                Call SetView((lvwGroup.View + 1) Mod 4)
            Else
                Call SetView((lvwReport.View + 1) Mod 4)
            End If
        Case "Add"
            mnuEdit_Add_Click
        Case "Modi"
            mnuEdit_Modi_Click
        Case "Del"
            mnuEdit_Del_Click
        Case "GroupAdd"
            mnuEdit_Group_Add_Click
        Case "GroupModify"
            mnuEdit_Group_Modify_Click
        Case "GroupDel"
            mnuEdit_Group_Delete_Click
        Case "Design"
            mnuEdit_Design_Click
        Case "Report"
            mnuFile_Report_Click
        Case "Publish"
            PopupButtonMenu tbr, Button, mnuPopPublish
        Case "unPub"
            PopupButtonMenu tbr, Button, mnuPopUnpub
        Case "Guide"
            mnuEdit_Guide_Click
        Case "Help"
            mnuHelpTitle_Click
    End Select
End Sub

Private Sub tbr_ButtonDropDown(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Publish"
            PopupButtonMenu tbr, Button, mnuPopPublish
        Case "unPub"
            PopupButtonMenu tbr, Button, mnuPopUnpub
    End Select
End Sub

Private Sub mnuView_View_Click(Index As Integer)
    Call SetView(CByte(Index))
End Sub

Private Sub SetView(bytStyle As Byte)
'���ܣ�������λ�б���ʾ��ʽ
'������bytstyle=0-��ͼ��,1-Сͼ��,2-�б�,3-��ϸ����
    mnuView_View(0).Checked = False
    mnuView_View(1).Checked = False
    mnuView_View(2).Checked = False
    mnuView_View(3).Checked = False
    mnuView_View(bytStyle).Checked = True
    
    On Error Resume Next
    If Me.ActiveControl.name = "lvwGroup" Then
        lvwGroup.View = bytStyle
    Else
        lvwReport.View = bytStyle
    End If
End Sub

Private Sub ShowRunLog()
    Dim lngReportKey As Long
    Dim strReportName As String
    lngReportKey = Val(Mid(lvwReport.SelectedItem.Key, 2))
    '�鿴����������־��¼
    If lngReportKey > 0 Then
        Call frmReportRunLog.ShowMe(Me, lngReportKey, "����[" & lvwReport.SelectedItem.Text & "]��������־")
    End If
End Sub

Private Sub tbr_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
        Case "Icon"
            Call SetView(0)
        Case "Small"
            Call SetView(1)
        Case "List"
            Call SetView(2)
        Case "Detail"
            Call SetView(3)
        Case "RunLog"
            If Not lvwReport.SelectedItem Is Nothing Then
                Call ShowRunLog
            End If
    End Select
End Sub

Private Function ReadReports(ByVal lngKey As Long) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim objItem As Object, strKey As String
    Dim strSQL As String, i As Integer
    
    On Error GoTo errH
    
    If Not lvwReport.SelectedItem Is Nothing Then
        strKey = lvwReport.SelectedItem.Key
    End If
    lvwReport.ListItems.Clear
    
    LockWindowUpdate lvwReport.hwnd
    
    If lngKey = -1 Then     '���б���
        If mnuViewOnly.Checked Then '����ʾ������
            strSQL = _
                "Select Distinct A.ID,A.���,A.����,A.˵��,A.����ID,A.�޸�ʱ��,A.����ʱ��,A.ϵͳ,Nvl(A.Ʊ��,0) Ʊ��,A.���ִ��ʱ��, " & vbCr & _
                "     A.ִ����Ա ���ִ����, zlSpellCode(A.����) ����, b.������������ " & vbCr & _
                "From zlReports A, " & vbCr & _
                "     (Select B1.����id, f_list2str(Cast(Collect(B2.����) As t_Strlist)) ������������ " & vbCr & _
                "      From zlRPTDatas B1, zlConnections B2 " & vbCr & _
                "      Where b1.�������ӱ�� = b2.��� And Not Exists(Select 1 From zlRPTSubs Where ����id = b1.����id) " & vbCr & _
                "      Group By b1.����id) B " & vbCr & _
                "Where a.id = b.����id(+) " & vbCr & _
                IIF(cboSys.ItemData(cboSys.ListIndex) = 0, " and A.ϵͳ Is Null ", " and A.ϵͳ=[1] ") & vbCr & _
                "     And Not Exists(Select 1 From zlRPTSubs Where ����id = a.Id) " & vbCr & _
                "Order by A.���"
        Else
            strSQL = _
                "Select Distinct A.ID,A.���,A.����,A.˵��,A.����ID,A.�޸�ʱ��,A.����ʱ��,A.ϵͳ,Nvl(A.Ʊ��,0) Ʊ��,A.���ִ��ʱ��, " & vbCr & _
                "     A.ִ����Ա ���ִ����, zlSpellCode(A.����) ����, b.������������  " & vbCr & _
                "From zlReports A, " & vbCr & _
                "     (Select B1.����id, f_list2str(Cast(Collect(B2.����) As t_Strlist)) ������������ " & vbCr & _
                "      From zlRPTDatas B1, zlConnections B2 " & vbCr & _
                "      Where b1.�������ӱ�� = b2.��� " & vbCr & _
                "      Group By b1.����id) B " & vbCr & _
                "Where a.id = b.����id(+) " & vbCr & _
                IIF(cboSys.ItemData(cboSys.ListIndex) = 0, " and A.ϵͳ Is Null ", " and A.ϵͳ=[1] ") & vbCr & _
                "Order by A.���"
        End If
    Else
        strSQL = _
            "Select Distinct A.ID,A.���,A.����,A.˵��,A.����ID,A.�޸�ʱ��,A.����ʱ��,A.ϵͳ,Nvl(A.Ʊ��,0) Ʊ��,A.���ִ��ʱ��, " & vbCr & _
            "     A.ִ����Ա ���ִ����, zlSpellCode(A.����) ����, b.������������ " & vbCr & _
            "From zlReports A, " & vbCr & _
            "     (Select B1.����id, f_list2str(Cast(Collect(B2.����) As t_Strlist)) ������������ " & vbCr & _
            "      From zlRPTDatas B1, zlConnections B2 " & vbCr & _
            "      Where b1.�������ӱ�� = b2.��� And Exists(Select 1 From zlRPTSubs Where ����id = b1.����id And ��ID=[2]) " & vbCr & _
            "      Group By b1.����id) B, zlRPTSubs C " & vbCr & _
            "Where a.id = b.����id(+) And a.Id = c.����Id And c.��ID=[2] " & vbCr & _
            "Order by A.���"
    End If
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, cboSys.ItemData(cboSys.ListIndex), lngKey)
    For i = 1 To rsTmp.RecordCount
        If Not IsNull(rsTmp!ϵͳ) Then '�̶���װ����
            If IsNull(rsTmp!����ʱ��) Then
                Set objItem = lvwReport.ListItems.Add(, "_" & rsTmp!ID, rsTmp!����, "Fixed", "Fixed")
            Else
                Set objItem = lvwReport.ListItems.Add(, "_" & rsTmp!ID, rsTmp!����, "PubFixed", "PubFixed")
            End If
            objItem.Tag = Val(Nvl(rsTmp!����ID, 0))
        Else
            If Not IsNull(rsTmp!����ʱ��) Then '�ѷ���
                If Nvl(rsTmp!Ʊ��, 0) = 1 Then
                    Set objItem = lvwReport.ListItems.Add(, "_" & rsTmp!ID, rsTmp!����, "BillPublish", "BillPublish")
                Else
                    Set objItem = lvwReport.ListItems.Add(, "_" & rsTmp!ID, rsTmp!����, "Publish", "Publish")
                End If
            Else
                If Nvl(rsTmp!Ʊ��, 0) = 1 Then
                    Set objItem = lvwReport.ListItems.Add(, "_" & rsTmp!ID, rsTmp!����, "Bill", "Bill")
                Else
                    Set objItem = lvwReport.ListItems.Add(, "_" & rsTmp!ID, rsTmp!����, "Report", "Report")
                End If
            End If
            objItem.Tag = Val(Nvl(rsTmp!����ID, 0))
        End If
        objItem.SubItems(RC_���) = rsTmp!���
        objItem.SubItems(RC_˵��) = Nvl(rsTmp!˵��)
        objItem.SubItems(RC_�޸�ʱ��) = Format(rsTmp!�޸�ʱ��, "yyyy-MM-dd")
        objItem.SubItems(RC_����ʱ��) = Format(Nvl(rsTmp!����ʱ��), "yyyy-MM-dd")
        objItem.SubItems(RC_���ִ��ʱ��) = Format(Nvl(rsTmp!���ִ��ʱ��), "yyyy-MM-dd hh:mm")
        objItem.SubItems(RC_���ִ����) = Nvl(rsTmp!���ִ����)
        objItem.SubItems(RC_����) = IIF(Nvl(rsTmp!Ʊ��, 0) = 1, "Ʊ��", "����")
        objItem.SubItems(RC_����) = IIF(IsNull(rsTmp!ϵͳ), "����", "ϵͳ")
        objItem.SubItems(RC_����) = rsTmp!���� & ""
        objItem.SubItems(RC_������������) = mdlPublic.Nvl(rsTmp!������������)
        If objItem.Key = strKey Then objItem.Selected = True
        rsTmp.MoveNext
    Next
    
    If Not lvwReport.SelectedItem Is Nothing Then
        lvwReport.SelectedItem.EnsureVisible
    End If
    
    'If rsTmp.RecordCount > 0 Then Call AutoSizeCol(lvw)
    LockWindowUpdate 0
    
    ReadReports = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub lvwReport_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnDesc As Boolean
    Static intIdx As Integer
    mcsActive = CS_����
    lblFind.Caption = "���ұ���(&F)"
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvwReport.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvwReport.SortOrder = lvwDescending
    Else
        lvwReport.SortOrder = lvwAscending
    End If
    lvwReport.Sorted = True
    intIdx = ColumnHeader.Index
    
    If Not lvwReport.SelectedItem Is Nothing Then lvwReport.SelectedItem.EnsureVisible
End Sub

Private Sub tbr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuView, 2
End Sub

Private Function GetNewProgID() As Long
'���ܣ���ȡ��һ�����õ��Զ��屨������,���ڷ���
'˵��������Ŵ�100000��ʼ,���Զ���ȱ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Decode(Sign(Max(���)-99999),1,Max(���),99999) as ID From zlPrograms"
    Call OpenRecord(rsTmp, strSQL, Me.Caption)
    GetNewProgID = IIF(IsNull(rsTmp!ID), 100000, rsTmp!ID + 1)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetMainTreeMenu(Optional ByVal lngProgID As Long) As ADODB.Recordset
'���ܣ���ȡ����������̨�������β˵���ϵ
'������lngProgID=�Ƿ�ֻ��ʾָ������ID�ı���
'˵�����˵���ϵ�а����Զ��屨�����Ĳ˵���(�����),��־Ϊ"FLAG=999"
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lngSys As Long
    
    On Error GoTo errH
    
    lngSys = cboSys.ItemData(cboSys.ListIndex)
    If lngSys = 0 Then
        'ֻ��ʾ�û��������ݱ���
        strSQL = _
            "Select Distinct * From (" & _
            " Select ��� as SCOL,0 as Flag,-��� as ID,-NULL as �ϼ�ID,'['||���||']'||���� as ����,-NULL as ģ�� From zlSystems Union ALL" & _
            " Select 99999 as SCOL,Level as FLAG,ID,Nvl(�ϼ�ID,-ϵͳ) as �ϼ�ID,����,ģ�� From zlMenus Where ���='ȱʡ' And ģ�� is NULL" & _
            " Start With �ϼ�ID is NULL And ���='ȱʡ' Connect by Prior ID=�ϼ�ID And ���='ȱʡ'" & _
            " Union ALL" & _
            " Select 99999 as SCOL,999 as FLAG,A.ID,A.�ϼ�ID,A.����,A.ģ��" & _
            " From zlMenus A,zlPrograms B,zlRPTGroups C" & _
            " Where A.ģ��=B.��� And A.���='ȱʡ' And C.����ID=A.ģ�� " & _
            " And Upper(B.����)='ZL9REPORT'" & IIF(lngProgID = 0, "", " And B.���=[1]") & _
            " And A.ϵͳ is NULL And B.ϵͳ is Null And C.ϵͳ is Null" & _
            " Union ALL" & _
            " Select 99999 as SCOL,888 as FLAG,A.ID,A.�ϼ�ID,A.����,A.ģ��" & _
            " From zlMenus A,zlPrograms B,zlReports C" & _
            " Where A.ģ��=B.��� And A.���='ȱʡ' And C.����ID=A.ģ�� " & _
            " And Upper(B.����)='ZL9REPORT'" & IIF(lngProgID = 0, "", " And B.���=[1]") & _
            " And A.ϵͳ is NULL And B.ϵͳ is Null And C.ϵͳ is Null" & _
            " ) Order by SCOL,FLAG,ID"
    Else
        'ֻ��ʾ�̶����ݱ���(����Ȩ����)
        strSQL = _
            "Select Distinct * From (" & _
            " Select ��� as SCOL,0 as Flag,-��� as ID,-NULL as �ϼ�ID,'['||���||']'||���� as ����,-NULL as ģ�� From zlSystems Union ALL" & _
            " Select 99999 as SCOL,Level as FLAG,ID,Nvl(�ϼ�ID,-ϵͳ) as �ϼ�ID,����,ģ�� From zlMenus Where ���='ȱʡ' And ģ�� is NULL" & _
            " Start With �ϼ�ID is NULL And ���='ȱʡ' Connect by Prior ID=�ϼ�ID And ���='ȱʡ'" & _
            " Union ALL" & _
            " Select 99999 as SCOL,999 as FLAG,A.ID,A.�ϼ�ID,A.����,A.ģ��" & _
            " From zlMenus A,zlPrograms B,zlRPTGroups C,(Select ϵͳ,��� From zlRegFunc Group By ϵͳ,���) D" & _
            " Where A.ģ��=B.��� And A.���='ȱʡ' And C.����ID=A.ģ�� " & _
            " And Upper(B.����)='ZL9REPORT'" & IIF(lngProgID = 0, "", " And B.���=[1]") & _
            " And A.ϵͳ=B.ϵͳ And A.ϵͳ=C.ϵͳ And Trunc(B.ϵͳ/100)=D.ϵͳ And B.���=D.���" & _
            " Union ALL" & _
            " Select 99999 as SCOL,888 as FLAG,A.ID,A.�ϼ�ID,A.����,A.ģ��" & _
            " From zlMenus A,zlPrograms B,zlReports C,(Select ϵͳ,��� From zlRegFunc Group By ϵͳ,���) D" & _
            " Where A.ģ��=B.��� And A.���='ȱʡ' And C.����ID=A.ģ�� " & _
            " And Upper(B.����)='ZL9REPORT'" & IIF(lngProgID = 0, "", " And B.���=[1]") & _
            " And A.ϵͳ=B.ϵͳ And A.ϵͳ=C.ϵͳ And Trunc(B.ϵͳ/100)=D.ϵͳ And B.���=D.���" & _
            " ) Order by SCOL,FLAG,ID"
    End If
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngProgID)
    Set GetMainTreeMenu = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetModuleTreeMenu(ByVal lngRPTID As Long) As ADODB.Recordset
'���ܣ���ȡ������ģ��ı������β˵���ϵ
'������lngRPTID=Ҫ������ȡ�������ı���ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '���˵���ʾģ��ķ�ʽ
    '-------------------------------------------------------------------------------------------------------------
    'ϵͳ + �м�˵� + ģ��˵�(��Ȩģ��) + ��������(��������Ȩģ����)
    'ע��ͬһģ������ظ�λ�ڲ�ͬ�˵�,����ʾ(��������ı���)
    '�ſ��������Զ��屨��ģ��(����='zl9Report')
    
    '�ſ�����Чģ��Ĳ˵�����(�Ƚ���)
    strSQL = _
        " Select Distinct Id From zlMenus Where ���='ȱʡ'" & _
        " Start With (ϵͳ,ģ��) In(Select ϵͳ,��� From zlPrograms Where Upper(����)<>Upper('zl9Report'))" & _
        " Connect By Prior �ϼ�ID=Id"
    
    strSQL = _
        " Select '1' as Sort1,To_Char(���) as Sort2," & _
        "   'S'||��� as ID,Null as �ϼ�ID,��� as ϵͳ,-Null as ����ID,Null as ����,'['||���||']'||���� as ����" & _
        " From zlSystems" & _
        " Union ALL " & _
        " Select '2' as Sort1,To_Char(Level) as Sort2," & _
        "   'T'||ID as ID,Decode(�ϼ�ID,NULL,'S'||ϵͳ,'T'||�ϼ�ID) as �ϼ�ID,ϵͳ,-Null as ����ID,Null as ����,����" & _
        " From zlMenus Where ���='ȱʡ' And ģ�� is Null" & _
        " Start With �ϼ�ID is NULL And ���='ȱʡ' Connect by Prior ID=�ϼ�ID And ���='ȱʡ'" & _
        " Union ALL " & _
        " Select '3' as Sort1,To_Char(B.���) as Sort2," & _
        "   'M'||B.���||'_'||A.ID as ID,'T'||A.�ϼ�ID as �ϼ�ID,B.ϵͳ,B.��� as ����ID,Null as ����,B.����" & _
        " From zlMenus A,zlPrograms B,(Select ϵͳ,��� From zlRegFunc Group By ϵͳ,���) C" & _
        " Where A.���='ȱʡ' And A.ϵͳ=B.ϵͳ And A.ģ��=B.��� And Upper(B.����)<>Upper('zl9Report')" & _
        " And Trunc(B.ϵͳ/100)=C.ϵͳ And B.���=C.���" & _
        " Union All " & _
        " Select '4' as Sort1,C.��� as Sort2," & _
        "   'R'||Rownum as ID,'M'||B.����ID||'_'||X.ID as �ϼ�ID,B.ϵͳ,B.����ID,B.����,'['||C.���||']'||C.���� as ����" & _
        " From zlMenus X,zlPrograms A,zlRPTPuts B,zlReports C,(Select ϵͳ,��� From zlRegFunc Group By ϵͳ,���) D" & _
        " Where X.���='ȱʡ' And X.ϵͳ=A.ϵͳ And X.ģ��=A.���" & _
        "   And A.ϵͳ=B.ϵͳ And A.���=B.����ID And Upper(A.����)<>Upper('zl9Report')" & _
        "   And Trunc(A.ϵͳ/100)=D.ϵͳ And A.���=D.���" & _
        "   And B.����ID=C.ID And C.ID=[1]" & _
        " Order by Sort1,Sort2"
    
    'ֻ��ʾģ��ķ�ʽ
    '-------------------------------------------------------------------------------------------------------------
    strSQL = _
        " Select '1' as Sort1,To_Char(���) as Sort2," & _
        "   'S'||��� as ID,Null as �ϼ�ID,��� as ϵͳ,-Null as ����ID,Null as ����,'['||���||']'||���� as ����" & _
        " From zlSystems" & _
        " Union ALL " & _
        " Select '3' as Sort1,To_Char(B.���) as Sort2," & _
        "   'M'||B.���||'_'||B.ϵͳ as ID,'S'||B.ϵͳ as �ϼ�ID,B.ϵͳ,B.��� as ����ID,Null as ����,'['||B.���||']'||B.����" & _
        " From zlPrograms B,(Select ϵͳ,��� From zlRegFunc Group By ϵͳ,���) C" & _
        " Where Upper(B.����)<>Upper('zl9Report') And Trunc(B.ϵͳ/100)=C.ϵͳ And B.���=C.���" & _
        " Union All " & _
        " Select '4' as Sort1,C.��� as Sort2," & _
        "   'R'||Rownum as ID,'M'||B.����ID||'_'||B.ϵͳ as �ϼ�ID,B.ϵͳ,B.����ID,B.����,'['||C.���||']'||C.���� as ����" & _
        " From zlPrograms A,zlRPTPuts B,zlReports C,(Select ϵͳ,��� From zlRegFunc Group By ϵͳ,���) D" & _
        " Where A.ϵͳ=B.ϵͳ And A.���=B.����ID And Upper(A.����)<>Upper('zl9Report')" & _
        "   And Trunc(A.ϵͳ/100)=D.ϵͳ And A.���=D.���" & _
        "   And B.����ID=C.ID And C.ID=[1]" & _
        " Order by Sort1,Sort2"
    
    '�̶������������¡��ɱ�������ϵͳ��ģ�飬����ϵͳ��10�汾����
    strSQL = "Select A.* From (" & strSQL & ") A,zlSystems B" & _
        " Where A.ϵͳ=B.��� And (To_Number(Substr(B.�汾��,1,Instr(B.�汾��,'.')-1))>=10 Or Trunc(���/100) IN(2,5,7))" & _
        " Order by Sort1,Sort2"
    
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngRPTID)
    Set GetModuleTreeMenu = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub mnuEdit_Publish_Main_Click()
    Call ReportGrantToNavigator
End Sub

Private Sub mnuEdit_unPub_Main_Click()
    Call ReportRevokeFromNavigator
End Sub

Private Sub mnuEdit_Guide_Click()
    Dim objReport As Report, objItem As Object
    Dim lngNext As Long, lngSys As Long, strSQL As String
    Dim i As Integer
    
    Set objReport = New Report
    With objReport
        .��ֽ = 15 'ȱʡΪ�Զ�ѡ��
        'ȱʡʹ�õ�ǰ��ӡ��
        If Printers.count > 0 Then .��ӡ�� = Printer.DeviceName
        'ȱʡΪA4����,Ϊ����
        .Fmts.Add 1, "��ʽ1", INIT_WIDTH, INIT_HEIGHT, 9, 1, False, 0, "_1"
    End With
    
    frmGuide.blnNew = True
    Set frmGuide.objReport = objReport
    Set frmGuide.mobjFmt = objReport.Fmts(1)
    frmGuide.Show 1, Me
    
    If gblnOK Then
        If cboSys.ListIndex <> 0 Then cboSys.ListIndex = 0
        Me.Refresh
        With frmGuide
            Set objReport.Items = .objGuide.Items
            Set objReport.Datas = .objGuide.Datas
            Set objReport.Fmts = .objGuide.Fmts
            
            '���ӱ���
            'lngSys = Split(GetSysNO, ",")(0)
            lngNext = GetNextID("zlReports")
            strSQL = "Insert Into zlReports(ID,���,����,˵��,ϵͳ,����) Values(" & _
                lngNext & ",'" & .txtNO.Text & "','" & .txtTitle.Text & "','" & _
                .txtNote.Text & "'," & IIF(lngSys = 0, "NULL", lngSys) & "," & AdjustStr(GetPass(.txtNO, .txtTitle)) & ")"
        
            On Error GoTo errH
            gcnOracle.BeginTrans
            gcnOracle.Execute strSQL
            gcnOracle.CommitTrans
            On Error GoTo 0
            
            '��������
            If Not SaveReport(lngNext, objReport, sta.Panels(2)) Then
                On Error GoTo errH
                gcnOracle.BeginTrans
                gcnOracle.Execute "Delete From zlReports Where ID=" & lngNext
                gcnOracle.CommitTrans
                On Error GoTo 0
                MsgBox "�����ɱ���ʱ�����������,�����Ըò�����", vbInformation, App.Title
                Unload frmGuide: Exit Sub
            End If
        
            '��������
            Set objItem = lvwReport.ListItems.Add(, "_" & lngNext, .txtTitle.Text, "Report", "Report")
            objItem.Tag = 0
            objItem.SubItems(RC_���) = .txtNO.Text
            objItem.SubItems(RC_˵��) = .txtNote.Text
            objItem.SubItems(RC_�޸�ʱ��) = Format(Currentdate, "yyyy-MM-dd")
            
            '����ѡ��
            For i = 1 To lvwReport.ListItems.count
                lvwReport.ListItems(i).Selected = (i >= lvwReport.ListItems.count)
            Next
            
            '����״̬
            lvwReport.SelectedItem.EnsureVisible
            lvwReport_ItemClick lvwReport.SelectedItem
        End With
        Unload frmGuide
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    gcnOracle.RollbackTrans
    Call SaveErrLog
    Unload frmGuide
End Sub

Private Sub lvwReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strPath As String
    Static objItem As Object
    
    If mblnMouseDown And Button = 1 Then
        lvwReport.DragIcon = lvwReport.SelectedItem.CreateDragImage
        lvwReport.Drag 1
    Else
        Set lvwReport.DragIcon = Nothing
        lvwReport.Drag 0
        mblnMouseDown = False
    End If
    If Not objItem Is Nothing And Not lvwReport.HitTest(X, Y) Is Nothing Then
        If objItem.Key = lvwReport.HitTest(X, Y).Key Then Exit Sub
    End If
    
    Set objItem = lvwReport.HitTest(X, Y)
    If Not objItem Is Nothing Then
        lvwReport.ToolTipText = objItem.SubItems(RC_˵��)
    End If
End Sub

Private Sub mnuFile_Exp_Click()
    Dim strMsg As String
    Dim strPath As String
    Dim strFile As String
    Dim i As Long, lngSelCount As Long
    
    If lvwReport.SelectedItem Is Nothing Then
        MsgBox "��ǰû�б�����Ե�����", vbInformation, App.Title: Exit Sub
    Else
        'SelectedItemֻ�����һ��ѡ�е��У������ѭ���������鿴�Ƿ��ѡ
        lngSelCount = 0
        For i = 1 To lvwReport.ListItems.count
            If lvwReport.ListItems(i).Selected Then lngSelCount = lngSelCount + 1
        Next
        If lngSelCount = 1 Then
            strMsg = frmMsgBox.ShowMsgBox(App.Title, "��ѡ�񱨱�����ʽ��^������ǰ�嵥�е����б���ʱ���ļ��Զ���""[���]����""������^�������Ŀ¼�д�����ͬ���Ƶı����ļ����ļ����ݽ������ǡ�", "���б���(&Y),!��ǰ����(&N),?ȡ��(&C)", Me)
             If strMsg = "" Then Exit Sub
        End If
    End If
    
    strPath = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\Path", "Export", GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\Path", "Import", App.Path))
    If strMsg = "��ǰ����" Then
        cdg.DialogTitle = "���������ļ�"
        cdg.Filter = "�Զ��屨���ļ�|*.ZLR"
        cdg.Flags = &H200000 Or &H4 Or &H2 Or &H800 Or &H4000
        cdg.InitDir = strPath
        
        strFile = "[" & lvwReport.SelectedItem.SubItems(RC_���) & "]" & lvwReport.SelectedItem.Text & ".ZLR"  'ȱʡ�Ա����������ļ���
        strFile = Replace(strFile, "\", "��")
        strFile = Replace(strFile, "/", "�M")
        strFile = Replace(strFile, ":", "��")
        strFile = Replace(strFile, "*", "�~")
        strFile = Replace(strFile, "?", "��")
        strFile = Replace(strFile, """", "")
        strFile = Replace(strFile, "<", "��")
        strFile = Replace(strFile, ">", "��")
        strFile = Replace(strFile, "|", "�O")
        cdg.FileName = strFile
        cdg.CancelError = True
        
        On Error Resume Next
        
        cdg.ShowSave
        If Err.Number = 0 Then
            Err.Clear
            On Error GoTo 0
            Me.Refresh
            SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\Path", "Export", Left(cdg.FileName, Len(cdg.FileName) - Len(cdg.FileTitle))
            Call ExportReport(CLng(Mid(lvwReport.SelectedItem.Key, 2)), cdg.FileName)
            VBA.Beep
        End If
    ElseIf strMsg = "���б���" Or lngSelCount > 1 Then
        strFile = BrowseForFolder(Me.hwnd, "ѡ�񱨱���Ŀ¼", strPath)
        If strFile <> "" Then
            strPath = strFile
            SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\Path", "Export", strPath
            lngSelCount = IIF(strMsg = "", lngSelCount, lvwReport.ListItems.count)
            If MsgBox("���ι����� " & lngSelCount & " �ű��� " & strPath & "��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
            
            For i = 1 To lvwReport.ListItems.count
                If lvwReport.ListItems(i).Selected Or strMsg <> "" Then
                    Call ShowFlash("���ڵ���:" & lvwReport.ListItems(i).Text & ".ZLR", i / lngSelCount, Me, True)
                    
                    strFile = "[" & lvwReport.ListItems(i).SubItems(RC_���) & "]" & lvwReport.ListItems(i).Text & ".ZLR"
                    If Not ExportReport(CLng(Mid(lvwReport.ListItems(i).Key, 2)), strPath & "\" & strFile) Then
                        Call ShowFlash
                        If MsgBox("��������ʱ���ִ���Ҫ����������һ�ű�����", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
                    End If
                End If
            Next
            Call ShowFlash
        End If
    End If
End Sub

Private Sub mnuFile_Imp_Click()
    Dim arrFile As Variant, strFile As String, i As Long
    Dim lngSys As Long, LngGroupID As Long, lngReportID As Long
    Dim rsFiles As ADODB.Recordset
    
    On Error GoTo errH
    cdg.DialogTitle = "ѡ���뱨��"
    cdg.Filter = "�Զ��屨���ļ�|*.ZLR"
    cdg.Flags = &H200 Or &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
    cdg.InitDir = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\Path", "Import", GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\Path", "Export", App.Path))
    cdg.FileName = ""
    cdg.MaxFileSize = 32767
    cdg.CancelError = True
    On Error Resume Next
    cdg.ShowOpen
    If Err.Number = 0 Then
        On Error GoTo errH
        Me.Refresh
        If cdg.FileTitle = "" Then
            'ѡ�����ļ�����
            SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\Path", "Import", Left(cdg.FileName, InStr(cdg.FileName, Chr(0)) - 1)
            arrFile = Split(cdg.FileName, Chr(0))
            For i = 1 To UBound(arrFile)
                strFile = strFile & "|" & arrFile(0) & "\" & arrFile(i)
            Next
            strFile = Mid(strFile, 2)
        Else
            'ѡ�񵥸��ļ�����
            SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\Path", "Import", Left(cdg.FileName, Len(cdg.FileName) - Len(cdg.FileTitle))
            strFile = cdg.FileName
        End If
        If strFile = "" Then Exit Sub
        arrFile = Split(strFile, "|")
        lngSys = cboSys.ItemData(cboSys.ListIndex)
        LngGroupID = IIF(lvwGroup.SelectedItem.Key = "_-1", 0, Val(Mid(lvwGroup.SelectedItem.Key, 2)))
        If Not lvwReport.SelectedItem Is Nothing Then lngReportID = CLng(Mid(lvwReport.SelectedItem.Key, 2))
        'FilePath=����ȫ·����FileName=�����ļ�������ID=����Ҫ����ı�����ID
        'ͬ��ID=�뽫Ҫ����ı���ͬ���ı���ı���ID���̶�����ͨ������ƥ�䣬�ǹ̶�ͨ������ƥ��
        '��������=0-�����룬1-��������,2-���ǵ���;��������=0-���帲�ǣ�1-������Դ����
        'ErrType=0-�޴���,1-�����ͬ����һ��������2-�����ͬ����һ�𸲸ǣ�3-ϵͳ����ֻ�ܸ��ǣ�������ͬ������
        '                            4-���ݴ�������,5-�汾��������,6-���Ʊ�Ŵ�������
        'ImportResult=-1-�Ѿ��ɹ����뵫�Ǳ��������δͨ����0-������,1-����ɹ�,2-����ʧ��
        'ImportInfo=����ɹ�����󷵻صı�����Ϣ
        Set rsFiles = CopyNewRec(Nothing, , True, _
                                    Array("FilePath", adVarChar, 1000, Empty, "FileName", adVarChar, 200, Empty, "��ID", adBigInt, Empty, Empty, _
                                             "ͬ��ID", adBigInt, Empty, Empty, "��������", adInteger, Empty, Empty, "��������", adInteger, Empty, Empty, _
                                             "ErrType", adInteger, Empty, Empty, "ImportResult", adInteger, Empty, Empty, "ImportInfo", adVarChar, 200, Empty))
        For i = LBound(arrFile) To UBound(arrFile)
            rsFiles.AddNew Array("FilePath", "FileName", "��ID", "ͬ��ID", "��������", "��������", "ErrType", "ImportResult", "ImportInfo"), _
                                    Array(arrFile(i), gobjFile.GetFileName(arrFile(i)), 0, 0, 0, 0, 0, 0, "")
        
        Next
        Call ImportReportBeach(lngSys, LngGroupID, lngReportID, rsFiles)
    Else
        Err.Clear
    End If
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function ImportReportBeach(ByVal lngSys As Long, ByVal lngGroup As Long, ByVal lngCurPRTID As Long, ByVal rsFiles As ADODB.Recordset, Optional ByVal blnALLImp As Boolean) As Boolean
'���ܣ��������뱨�����Ե���1�������
'������
'          lngSys=��ǰѡ���ϵͳ
'          lngGroup=��ǰѡ��ļ�¼��
'          rsFiles=��Ҫ����ı����ļ�
'          lngCurPRTID=��ǰѡ��ı���ID
'          blnALLImp=�Ƿ���ȫ�����룬�ǹ̶�����ȫ������ʱ��Ҳ��Ҫ��ȡ���б���
'���أ��Ƿ�ɹ�����

    Dim rsReports As New ADODB.Recordset, strSQL As String
    Dim arrTmp As Variant, strInfo As String
    Dim strFilter As String
    Dim intErrType As Integer, intImpType As Integer, lngImpGroup As Long, lngRPTID As Long
    Dim strMsg As String, strOption As String, strReturn As String
    Dim i As Long, lngCount As Long
    Dim blnSingle  As Boolean, strFileName As String
    Dim strCurRPT As String, strSameRPT As String
    
    On Error GoTo errH
    '�̶������Լ�����ʾ�������µķǹ̶���������б������ʱ����Ҫ��ȡ���б���
    If lngSys <> 0 Or Not mnuViewOnly.Checked And lngGroup = 0 And lngSys = 0 Or blnALLImp Then
        '��ѯ���еı���
        strSQL = "Select A.ID,A.���,A.����,A.˵��,Nvl(B.��id,0) ��id" & vbNewLine & _
                        "From zlReports A,zlRPTSubs B" & vbNewLine & _
                        "Where " & IIF(lngSys = 0, " A.ϵͳ Is Null", "A.ϵͳ=[1]") & vbNewLine & _
                        "And  A. ID=B.����ID(+)" & vbNewLine & _
                        "Order by A.���"
    Else '�ǹ̶������ȡ
        If lngGroup <> 0 Then
            strSQL = "Select Id, ���, ����,[2] ��id" & vbNewLine & _
                            "From Zlreports" & vbNewLine & _
                            "Where Id In (Select ����id From Zlrptsubs Where ��id = [2])" & vbNewLine & _
                            "Order By ���"
        Else
            strSQL = "Select ID,���,����,0 ��id" & vbNewLine & _
                            "From zlReports" & vbNewLine & _
                            "Where " & IIF(lngSys = 0, " ϵͳ Is Null", "ϵͳ=[1]") & vbNewLine & _
                            "And ID Not In (Select ����ID From zlRPTSubs)" & vbNewLine & _
                            "Order by ���"
        End If
    End If
    Set rsReports = CopyNewRec(OpenSQLRecord(strSQL, Me.Caption, lngSys, lngGroup))
    If lngCurPRTID <> 0 Then
        rsReports.Filter = "ID=" & lngCurPRTID
        If rsReports.EOF Then
            MsgBox "��ǰѡ�б����Ѿ���ɾ������ˢ�º������", vbInformation, App.Title
            Exit Function
        Else
            strCurRPT = "[" & rsReports!��� & "]" & rsReports!����
        End If
    End If
    With rsFiles
        '��ͬ���ļ����뵽ͬһ����ʱ��ͬ���ļ����
        '����������£�[GROUP_001]סԺ��������ASD��סԺ��������[GROUP_001]סԺ��������
        '                       ���������ļ��ı�����Ե��뵽[GROUP_001]סԺ������������
        '��ͬ�ļ����ı���������ͬһ������
        '��鵼���ļ����Լ�ȷ���������ͣ���������Լ����ǵı���ID��
        .Filter = "": .Sort = "FilePath Desc"
        blnSingle = rsFiles.RecordCount = 1 '�Ƿ񵥸�������
        If blnSingle Then strFileName = rsFiles!FileName
        Do While Not .EOF
            intErrType = 0: intImpType = 0: lngImpGroup = 0: lngRPTID = 0
            arrTmp = Split(GetReportInfo(!FilePath & ""), ";") '��ȡ�ļ���Ϣ
            If UBound(arrTmp) <> 2 Then
                intErrType = 4 '�ļ����
            ElseIf Val(arrTmp(2)) <> 9 Then
                intErrType = 5  '�汾���
                If blnSingle Then strFileName = strFileName & "(ԭʼ���ƣ�[" & arrTmp(0) & "]" & arrTmp(1) & ")"
            Else
                If blnSingle Then strFileName = strFileName & "(ԭʼ���ƣ�[" & arrTmp(0) & "]" & arrTmp(1) & ")"
                If lngSys = 0 Then '��ϵͳ����Ҫ�����ı����в��ܴ�����ͬ����
                    '�ǹ̶�����ȫ�������Ѿ�ȷ������Ҫ����ķ���
                    rsReports.Filter = "����='" & arrTmp(1) & "' And ���='" & arrTmp(0) & "' And ID>0 " & IIF(blnALLImp, " And ��ID=" & !��ID, "")
                    If rsReports.EOF Then rsReports.Filter = "����='" & arrTmp(1) & "'  And ID>0 " & IIF(blnALLImp, " And ��ID=" & !��ID, "")
                Else 'ϵͳ����ͨ�����ֱ�Ӳ���
                    rsReports.Filter = "����='" & arrTmp(1) & "' And ���='" & arrTmp(0) & "' And ID>0"
                    If rsReports.EOF Then rsReports.Filter = "���='" & arrTmp(0) & "' And ID>0"
                End If
                'ȷ��������ķ��飬������ڵ�ͬ���ģ����Ȳ���û�з���ı���
                rsReports.Sort = "ID Desc,��ID"
                If Not rsReports.EOF Then
                    lngRPTID = rsReports!ID: lngImpGroup = rsReports!��ID
                    If lngRPTID = 0 Then
                        intErrType = 1 '�ñ����Ѿ����������
                    ElseIf lngRPTID < 0 Then
                        intErrType = 2 '�ñ����Ѿ�����Ǹ���
                    Else
                        intImpType = 2
                        '������Ʋ�ƥ��
                        If (CStr(arrTmp(0)) <> rsReports!��� & "" Or CStr(arrTmp(1)) <> rsReports!����) Then intErrType = 6
                        rsReports.Update "Id", lngRPTID * -1 '����Ѿ�����
                        If blnSingle Then strSameRPT = "[" & rsReports!��� & "]" & rsReports!����
                    End If
                Else
                    If lngSys <> 0 Then
                        intErrType = 3  'ϵͳ�̶�������븲��ͬ������
                    Else
                        intImpType = 1  '��ϵͳ����û��ͬ��������������
                        If lngSys = 0 And blnALLImp Then lngImpGroup = !��ID '�ǹ̶�������ȡԭ���ķ���
                        '�ñ�����������������뻺�棬��ֹ�������
                        rsReports.AddNew Array("Id", "���", "����", "��iD"), Array(lngRPTID, arrTmp(0), arrTmp(1), !��ID)
                    End If
                End If
            End If
            If lngSys = 0 And blnALLImp Then lngImpGroup = !��ID '�ǹ̶�������ȡԭ���ķ���
            .Update Array("��ID", "ͬ��ID", "��������", "ErrType"), Array(lngImpGroup, lngRPTID, intImpType, intErrType)
            .MoveNext
        Loop
        If blnSingle Then
            .Filter = ""
            Select Case !ErrType
                Case 4
                    MsgBox "����""" & strFileName & """�������ݴ���������޷����룡", vbInformation, App.Title
                    Exit Function
                Case 5
                    MsgBox "����""" & strFileName & """���ڰ汾���Զ��޷����룡", vbInformation, App.Title
                    Exit Function
                Case 3
                    If lngCurPRTID <> 0 Then '����״̬��Ĭ�ϸ��ǵ�ǰ�ı���
                        .Update Array("��ID", "ͬ��ID", "��������", "ErrType"), Array(lngGroup, lngCurPRTID, 2, 6)
                    Else
                        MsgBox "��ѡ����Ҫ���ǵı���������", vbInformation, App.Title
                        Exit Function
                    End If
            End Select
            Select Case !��������
                Case 1
                    strReturn = frmMsgBox.ShowMsgBox(App.Title, "�Ƿ��������뱨��""" & strFileName & """��", "��������(&N),!?ȡ��(&C)", Me)
                Case 2
                    If lngSys = 0 And lngGroup = 0 Then '����ϵͳ�����Ϊ����ı���,��ʱ���Դ�����������ѡ��
                        If lngCurPRTID = !ͬ��ID Then
                            strMsg = IIF(!ErrType = 6, "����""" & strFileName & """��Ż�����" & vbNewLine & "��Ҫ���ǵĵ�ǰѡ�񱨱�""" & strCurRPT & """���������ѡ��ȷ�ϣ�", _
                                        "����""" & strFileName & """��ź�����" & vbNewLine & "�뵱ǰѡ�񱨱�""" & strCurRPT & """���������ѡ��ȷ�ϣ�") & vbNewLine & "^^ע�⣺���Ҫ���Ǳ������ȶ�Ҫ���Ǳ�����б��ݡ�"
                            strReturn = frmMsgBox.ShowMsgBox(App.Title, strMsg, "���ǵ�ǰ(&S),��������(&N),!?ȡ��(&C)", Me)
                        ElseIf lngCurPRTID = 0 Then
                            strMsg = IIF(!ErrType = 6, "����""" & strFileName & """���ڲ���ƥ��ı���""" & strSameRPT & """," & vbNewLine & "���Ƕ��߱�Ż����Ʋ��������ѡ��ȷ�ϣ�", _
                                        "����""" & strFileName & """���ڱ��������ƾ�����ı���""" & strSameRPT & """����ѡ��ȷ�ϣ�") & vbNewLine & "^^ע�⣺���Ҫ���Ǳ������ȶ�Ҫ���Ǳ�����б��ݡ�"
                            strReturn = frmMsgBox.ShowMsgBox(App.Title, strMsg, "����ƥ��(&O),��������(&N),!?ȡ��(&C)", Me)
                        Else
                            strMsg = IIF(!ErrType = 6, "����""" & strFileName & """�ı�Ż�����" & vbNewLine & "�벿��ƥ�䱨��""" & strSameRPT & """" & vbNewLine & "�Լ���ǰѡ�񱨱�""" & strCurRPT & """�����������ѡ��ȷ�ϣ�", _
                                        "����""" & strFileName & """��Ż�����" & vbNewLine & "�뵱ǰѡ��""" & strCurRPT & """�������" & vbNewLine & "���Ǵ��ڱ��������ƾ�����ı���""" & strSameRPT & """����ѡ��ȷ�ϣ�") & vbNewLine & "^^ע�⣺���Ҫ���Ǳ������ȶ�Ҫ���Ǳ�����б��ݡ�"
                            strReturn = frmMsgBox.ShowMsgBox(App.Title, strMsg, "���ǵ�ǰ(&S),����ƥ��(&O),��������(&N),!?ȡ��(&C)", Me)
                        End If
                    Else
                       If lngCurPRTID = !ͬ��ID Then
                            strMsg = IIF(!ErrType = 6, "����""" & strFileName & """��Ż�����" & vbNewLine & "��Ҫ���ǵĵ�ǰѡ�񱨱�""" & strCurRPT & """���������ѡ��ȷ�ϣ�", _
                                        "����""" & strFileName & """��ź�����" & vbNewLine & "�뵱ǰѡ�񱨱�""" & strCurRPT & """���������ѡ��ȷ�ϣ�") & vbNewLine & "^^ע�⣺���Ҫ���Ǳ������ȶ�Ҫ���Ǳ�����б��ݡ�"
                            strReturn = frmMsgBox.ShowMsgBox(App.Title, strMsg, "���ǵ�ǰ(&S),!?ȡ��(&C)", Me)
                        ElseIf lngCurPRTID = 0 Then
                            strMsg = IIF(!ErrType = 6, "����""" & strFileName & """���ڲ���ƥ��ı���""" & strSameRPT & """," & vbNewLine & "���Ƕ��߱�Ż����Ʋ��������ѡ��ȷ�ϣ�", _
                                        "����""" & strFileName & """����" & vbNewLine & "���������ƾ�����ı���""" & strSameRPT & """����ѡ��ȷ�ϣ�") & vbNewLine & "^^ע�⣺���Ҫ���Ǳ������ȶ�Ҫ���Ǳ�����б��ݡ�"
                            strReturn = frmMsgBox.ShowMsgBox(App.Title, strMsg, "����ƥ��(&O),!?ȡ��(&C)", Me)
                        Else
                            strMsg = IIF(!ErrType = 6, "����""" & strFileName & """�ı�Ż�����" & vbNewLine & "�벿��ƥ�䱨��""" & strSameRPT & """" & vbNewLine & " �Լ���ǰѡ�񱨱�""" & strCurRPT & """�����������ѡ��ȷ�ϣ�", _
                                        "����""" & strFileName & """��Ż�����" & vbNewLine & "�뵱ǰѡ��""" & strCurRPT & """�������" & vbNewLine & "���Ǵ��ڱ��������ƾ�����ı���""" & strSameRPT & """����ѡ��ȷ�ϣ�") & vbNewLine & "^^ע�⣺���Ҫ���Ǳ������ȶ�Ҫ���Ǳ�����б��ݡ�"
                            strReturn = frmMsgBox.ShowMsgBox(App.Title, strMsg, "���ǵ�ǰ(&S),����ƥ��(&O),!?ȡ��(&C)", Me)
                        End If
                    End If
            End Select
            If strReturn = "" Then
                Exit Function
            ElseIf strReturn = "��������" Then
                .Update Array("��ID", "ͬ��ID", "��������", "ErrType"), Array(lngGroup, 0, 1, 0)
            Else
                If strReturn = "���ǵ�ǰ" Then
                    .Update Array("��ID", "ͬ��ID", "��������", "ErrType"), Array(lngGroup, lngCurPRTID, 2, 0)
                Else
                    .Update Array("��������", "ErrType"), Array(2, 0)
                End If
                strMsg = frmMsgBox.ShowMsgBox(App.Title, "�Ƿ�ֻ��������Դ��" & vbNewLine & "ֻ��������Դ���Ա������б���ĸ�ʽ������ϸ���������ѯϵͳ����Ա��", "������Դ(&D),!?���嵼��(&F)", Me)
                If strMsg = "������Դ" Then
                    .Update "��������", 1
                End If
            End If
        Else
            If MsgBox("��ǰ������ű���ϵͳ���Զ�Ѱ�ұ��������ƥ��ı�����и��ǡ���ȷ���Ƿ������", vbInformation + vbYesNo, App.Title) = vbNo Then
                Exit Function
            End If
            '���ܵ����������Ϣ����
            .Filter = "ErrType>0 And ErrType<6": .Sort = "ErrType": intImpType = 0
            Do While Not .EOF
                If intImpType <> Val(!ErrType & "") Then
                    If intImpType <> 0 Then
                        strMsg = strMsg & vbNewLine
                    End If
                    intImpType = Val(!ErrType & ""): lngCount = 0
                    Select Case intImpType
                        Case 1
                            strMsg = strMsg & vbNewLine & "���±������ڴ�����ͬ���ݵı�����޷��������룺"
                        Case 2
                            strMsg = strMsg & vbNewLine & "���±������ڴ�����ͬ���ݵı�����޷����ǵ��룺"
                        Case 3
                            strMsg = strMsg & vbNewLine & "���±�������û�п��Ը��ǵı�����޷����룺"
                        Case 4
                            strMsg = strMsg & vbNewLine & "���±����������ݴ���������޷����룺"
                        Case 5
                            strMsg = strMsg & vbNewLine & "���±������ڰ汾���Զ��޷����룺"
                    End Select
                End If
                If lngCount < 4 Then
                    strMsg = strMsg & vbNewLine & !FileName
                ElseIf lngCount = 4 Then
                    strMsg = strMsg & vbNewLine & "... ..."
                End If
                lngCount = lngCount + 1: .MoveNext
                If .EOF Then strMsg = strMsg & vbNewLine
            Loop
            .Filter = "��������<>0"
            If .RecordCount = 0 Then 'û�е��뱨��
                MsgBox "û�п��Ե���ı���" & Mid(strMsg, 1, Len(strMsg) - 2) & "��", vbInformation, App.Title
                Exit Function
            End If
            '�ļ����Լ����벻ƥ����ʾ
            .Filter = "ErrType=6"
            If Not .EOF Then
                lngCount = 0: strMsg = strMsg & vbNewLine & "��Ż������븲�ǵı����������ѡ��ȷ�ϣ�"
                Do While Not .EOF
                    If lngCount < 4 Then
                        strMsg = strMsg & vbNewLine & !FileName
                    ElseIf lngCount = 4 Then
                        strMsg = strMsg & vbNewLine & "... ..."
                    End If
                    lngCount = lngCount + 1: .MoveNext
                    If .EOF Then strMsg = strMsg & vbNewLine
                Loop
                .Filter = "ErrType=0" '�����ڿ���ֱ�ӵ���ģ�����ʾ�Ƿ����
                If .RecordCount = 0 Then
                    strReturn = frmMsgBox.ShowMsgBox(App.Title, Mid(strMsg, 1, Len(strMsg) - Len(vbNewLine)), "���帲��(&A),����Դ����(&D),!?ȡ��(&C)", Me)
                    If strReturn = "" Then Exit Function
                End If
            End If
            .Filter = "��������=2 And ErrType=0": .Sort = "ErrType" '���ڸ��Ǳ�������ʾѡ�����帲�ǣ���������Դ����
            If Not .EOF Then
                strMsg = strMsg & vbNewLine & "���±����Ḳ��ԭ�б�����ѡ��ȷ�ϣ�"
                strOption = "���帲��(&A),����Դ����(&D),!?ȡ��(&C)"
                lngCount = 0
            End If

            Do While Not .EOF
                If lngCount < 4 Then
                    strMsg = strMsg & vbNewLine & !FileName
                ElseIf lngCount = 4 Then
                    strMsg = strMsg & vbNewLine & "... ..."
                End If
                lngCount = lngCount + 1: .MoveNext
                If .EOF Then strMsg = strMsg & vbNewLine
            Loop
            .Filter = "��������=1" '��������
            If .RecordCount <> 0 And strReturn = "" And strOption = "" Then '���б�������
                strReturn = frmMsgBox.ShowMsgBox(App.Title, Mid(strMsg, Len(vbNewLine) + 1) & "��ȷ���Ƿ��룿", "����(&N),!?ȡ��(&C)", Me)
                If strReturn = "" Then Exit Function
            End If
            'ѡ�񸲸�����
            If strReturn = "" And strOption <> "" Then '���ڸ���,�Ҳ�����ErrType=6������
                strReturn = frmMsgBox.ShowMsgBox(App.Title, Mid(strMsg, Len(vbNewLine) + 1, Len(strMsg) - Len(vbNewLine) * 2), strOption, Me)
                If strReturn = "" Then Exit Function
            End If
        End If
        If strReturn = "����Դ����" Then
            .Filter = "��������=2"
            Do While Not .EOF
                .Update "��������", 1
                .MoveNext
            Loop
        End If
        Screen.MousePointer = 11
        .Filter = "��������<>0": .Sort = "��������"
        lngCount = .RecordCount
        Do While Not .EOF
            If Not blnSingle Then
                Call ShowFlash("���ڵ���:" & !FileName, i / lngCount, Me, True)
            Else
                Call ShowFlash("���ڵ���:" & !FileName, , Me, True)
            End If
            Me.Refresh
            DoEvents
            strInfo = ImportReport(!FilePath & "", Val(!ͬ��ID & ""), Val(!�������� & "") = 1, Val(!��ID & ""))
            .Update Array("ImportResult", "ImportInfo"), Array(IIF(strInfo <> "", 1, 2), strInfo)
            '�������Ȩ�޼��
            If strInfo <> "" Then
                arrTmp = Split(strInfo, "|")
                If Not CheckReportPriv(CLng(arrTmp(0))) Then
                    .Update Array("ImportResult", "ͬ��ID"), Array(-1, Val(arrTmp(0)))
                Else
                    .Update "ͬ��ID", Val(arrTmp(0))
                End If
            End If
            i = i + 1
            .MoveNext
        Loop
        Call ShowFlash
        If Not blnSingle Then
            lngGroup = Val(Mid(mstrPreGroup, 2))
        Else
            .Filter = ""
            lngGroup = Val(!��ID & "")
        End If
        'ˢ�½��棬���¼�������
        On Error Resume Next
        mstrPreGroup = ""
        '�ǹ̶�����ȫ��������Ҫˢ�·���
        If lngSys = 0 And blnALLImp Then
            '��¼������ID
            Call ReadGroups
        End If
        '���¶�λ��ǰ����
        For i = 1 To lvwGroup.ListItems.count
            If lvwGroup.ListItems(i).Key = "_" & IIF(lngGroup = 0, -1, lngGroup) Then
                lvwGroup.ListItems(i).Selected = True
            Else
                lvwGroup.ListItems(i).Selected = False
            End If
        Next
        lvwGroup.SelectedItem.EnsureVisible: lvwGroup.Refresh
        Call LvwGroup_ItemClick(lvwGroup.SelectedItem)
        '���ѡ��ı���
        For i = 1 To lvwReport.ListItems.count
            lvwReport.ListItems(i).Selected = False
        Next
        '���뱨��ѡ��
        .Filter = "��ID= " & lngGroup
        .Sort = "ͬ��ID"
        Do While Not .EOF
            lvwReport.ListItems("_" & !ͬ��ID).Selected = True
            .MoveNext
        Loop
        lvwReport.SelectedItem.EnsureVisible: lvwReport.Refresh
        Call cbr.Refresh
        Err.Clear: On Error GoTo errH
        '���������ʾ
        strMsg = ""
        If Not blnSingle Then
            .Filter = "ImportResult=1 Or ImportResult=-1"
            If .RecordCount = 0 Then
                strMsg = "���б����Ϊ����ɹ���"
            Else
                strMsg = "�ɹ������� " & .RecordCount & " �ű���"
            End If
            .Filter = "ImportResult=2"
            If .RecordCount <> 0 Then
                lngCount = 0: strMsg = strMsg & vbNewLine & "���±���ı����ļ����ݿ����ѱ��Ƿ��޸ģ�"
                Do While Not .EOF
                    If lngCount < 4 Then
                        strMsg = strMsg & vbNewLine & !FileName
                    ElseIf lngCount = 4 Then
                        strMsg = strMsg & vbNewLine & "... ..."
                    End If
                    lngCount = lngCount + 1: .MoveNext
                    If .EOF Then strMsg = strMsg & vbNewLine
                Loop
            End If
            .Filter = "ImportResult=-1 And ��������=1"
            If .RecordCount <> 0 Then
                lngCount = 0: strMsg = strMsg & vbNewLine & "��û��Ȩ�޲�ѯ���µ��뱨����ȫ���򲿷����ݶ���"
                Do While Not .EOF
                    If lngCount < 4 Then
                        strMsg = strMsg & vbNewLine & !FileName
                    ElseIf lngCount = 4 Then
                        strMsg = strMsg & vbNewLine & "... ..."
                    End If
                    lngCount = lngCount + 1: .MoveNext
                    If .EOF Then strMsg = strMsg & vbNewLine
                Loop
            End If
            .Filter = "ImportResult=-1 And ��������=2"
            If .RecordCount <> 0 Then
                lngCount = 0: strMsg = strMsg & vbNewLine & "��û��Ȩ�޲�ѯ���µ��뱨����ȫ���򲿷����ݶ���,��ʹ�øñ���֮ǰ,���ֹ��Ա������ݽ��е�����"
                Do While Not .EOF
                    If lngCount < 4 Then
                        strMsg = strMsg & vbNewLine & !FileName
                    ElseIf lngCount = 4 Then
                        strMsg = strMsg & vbNewLine & "... ..."
                    End If
                    lngCount = lngCount + 1: .MoveNext
                    If .EOF Then strMsg = strMsg & vbNewLine
                Loop
            End If
            .Filter = "ImportResult=1 And ��������=2"
            If .RecordCount <> 0 And lngSys <> 0 Then
                lngCount = 0: strMsg = strMsg & vbNewLine & "���±���ɹ�������Ӧ����,�������Ҫ������Ȩ��������ʹ����Щ����"
                Do While Not .EOF
                    If lngCount < 4 Then
                        strMsg = strMsg & vbNewLine & !FileName
                    ElseIf lngCount = 4 Then
                        strMsg = strMsg & vbNewLine & "... ..."
                    End If
                    lngCount = lngCount + 1: .MoveNext
                    If .EOF Then strMsg = strMsg & vbNewLine
                Loop
            End If
            .Filter = "ImportResult=2"
            If .RecordCount <> 0 Then
                lngCount = 0: strMsg = strMsg & vbNewLine & "���±�����ʧ�ܣ�"
                Do While Not .EOF
                    If lngCount < 4 Then
                        strMsg = strMsg & vbNewLine & !FileName
                    ElseIf lngCount = 4 Then
                        strMsg = strMsg & vbNewLine & "... ..."
                    End If
                    lngCount = lngCount + 1: .MoveNext
                    If .EOF Then strMsg = strMsg & vbNewLine
                Loop
            End If
        Else
            .Filter = ""
            Select Case !ImportResult
                Case -1
                    strMsg = "��û��Ȩ�޲�ѯ����""" & strFileName & """��ȫ���򲿷����ݶ���" & IIF(!�������� = 2, "���������Ҫ�ֹ��Ա������ݽ��е�����������Ȩ��������ʹ�øñ���", "��")
                Case 1
                    strMsg = "����""" & strFileName & """����ɹ�" & IIF(!�������� = 2, "���������Ҫ������Ȩ��������ʹ�øñ���", "��")
                Case 2
                    strMsg = "����""" & strFileName & """" & IIF(!�������� = 2, "����ʧ�ܡ������ļ����ݿ����ѱ��Ƿ��޸ģ�", "��������ʧ�ܣ�")
            End Select
        End If
        MsgBox strMsg, vbInformation, App.Title
        Screen.MousePointer = 0
    End With
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    Call ShowFlash
    Call SaveErrLog
End Function

Private Sub ReadSystem()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    
    On Error GoTo errH
    
    cboSys.Clear
    cboSys.AddItem "����ϵͳ����"
    
    strSQL = "Select ���,���� From zlSystems Order by ���"
    Call OpenRecord(rsTmp, strSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        cboSys.AddItem Lpad(rsTmp!���, 4) & "-" & rsTmp!����
        cboSys.ItemData(cboSys.NewIndex) = rsTmp!���
        rsTmp.MoveNext
    Next
    cboSys.ListIndex = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function ReadGroups() As Boolean
    Dim rsReportGroup As New ADODB.Recordset
    Dim strSQL As String, ItemThis As ListItem
    'װ�����б�����
    
    strSQL = "Select ID,���,����,˵��,ϵͳ,����ID,����ʱ�� , zlSpellCode(����) ���� From zlRPTGroups " & _
        " Where " & IIF(cboSys.ItemData(cboSys.ListIndex) = 0, " ϵͳ Is Null", " ϵͳ=[1]")
    Set rsReportGroup = OpenSQLRecord(strSQL, Me.Caption, cboSys.ItemData(cboSys.ListIndex))
    
    LockWindowUpdate lvwGroup.hwnd
    lvwGroup.ListItems.Clear
    lvwGroup.ListItems.Add , "_-1", "���б���", 5, 5
    
    With rsReportGroup
        Do While Not .EOF
            If Not IsNull(!ϵͳ) Then
                If Not IsNull(!����ʱ��) Then     '�̶�����(����������ȡ������)
                    Set ItemThis = lvwGroup.ListItems.Add(, "_" & !ID, !����, 4, 4)
                Else
                    Set ItemThis = lvwGroup.ListItems.Add(, "_" & !ID, !����, 3, 3)
                End If
            Else
                If Not IsNull(!����ʱ��) Then     '�ǹ̶�����(��������ȡ������)
                    Set ItemThis = lvwGroup.ListItems.Add(, "_" & !ID, !����, 2, 2)
                Else
                    Set ItemThis = lvwGroup.ListItems.Add(, "_" & !ID, !����, 1, 1)
                End If
            End If
            ItemThis.SubItems(GC_���) = !���
            ItemThis.SubItems(GC_˵��) = Nvl(!˵��)
            ItemThis.SubItems(GC_����ʱ��) = Format(Nvl(!����ʱ��), "yyyy-MM-dd")
            ItemThis.SubItems(GC_����) = !����
            ItemThis.Tag = Val(Nvl(!����ID, 0))
            .MoveNext
        Loop
    End With
    
    lvwGroup.ListItems("_-1").Selected = True
    lvwGroup.SelectedItem.Selected = True
    
    'Call AutoSizeCol(lvwGroup)
    LockWindowUpdate 0
    
    mstrPreGroup = ""
    Call LvwGroup_ItemClick(lvwGroup.SelectedItem)
    ReadGroups = True
End Function

Private Sub ReportGrantToNavigator()
'���ܣ�������ǰ����(��)������̨,���ܲ��ǵ�һ��
    Dim rsTmp As ADODB.Recordset
    Dim rsCheck As New ADODB.Recordset
    Dim rsSubRPT As New ADODB.Recordset
    Dim objNode As Object, i As Integer, j As Integer, k As Integer
    Dim strObject As String, strOwner As String, strName As String
    Dim lngRPTID As Long, lngProgID As Long, lngMenu As Long
    Dim strTmp As String, lngNewMenu As Long, lngSys As Long
    Dim strSQL As String
    
    If lvwGroup.SelectedItem.Key = "_-1" Or mcsActive = CS_���� Then
        'ѡ�����б���ʱ
        If lvwReport.SelectedItem Is Nothing Then MsgBox "��ǰû�б�����Է�����", vbInformation, App.Title: Exit Sub
        If lvwReport.SelectedItem.Tag <> 0 Then
            If lvwReport.SelectedItem.Icon = "Fixed" Or lvwReport.SelectedItem.Icon = "PubFixed" Then
                MsgBox "�ñ���Ϊϵͳ���еı���,�������ܼ�����", vbInformation, App.Title: Exit Sub
            End If
        End If
        If CheckPass(CLng(Mid(lvwReport.SelectedItem.Key, 2))) = False Then
            MsgBox "�������ݴ��󣬲��ܷ����ñ���", vbInformation, App.Title: Exit Sub
        End If
        If Not CheckReportPriv(CLng(Mid(lvwReport.SelectedItem.Key, 2))) Then
            MsgBox "��û��Ȩ�޲�ѯ�ñ���ĳЩ����Դ�еĶ���,�������ܼ�����", vbInformation, App.Title
            Exit Sub
        End If
        lngRPTID = CLng(Mid(lvwReport.SelectedItem.Key, 2))
    Else
        '����ĳ��������
        If lvwGroup.SelectedItem Is Nothing Then MsgBox "��ǰû�б�������Է�����", vbInformation, App.Title: Exit Sub
        If lvwGroup.SelectedItem.Tag <> 0 Then
            If lvwGroup.SelectedItem.Icon = 3 Or lvwGroup.SelectedItem.Icon = 4 Then
                MsgBox "�ñ�����Ϊϵͳ���еı���,�������ܼ�����", vbInformation, App.Title: Exit Sub
            End If
        End If
        
        If Me.lvwReport.ListItems.count = 0 Then
            MsgBox "�ñ������в������κα������ܷ�����", vbInformation, App.Title
            Exit Sub
        Else
            For i = 1 To lvwReport.ListItems.count - 1
                For j = i + 1 To lvwReport.ListItems.count
                    If lvwReport.ListItems(i).Text = lvwReport.ListItems(j).Text Then
                        MsgBox "�ñ������а�����ͬ���Ƶı���""" & lvwReport.ListItems(i).Text & """�����ܷ�����", vbInformation, App.Title
                        Exit Sub
                    End If
                Next
            Next
        End If
        
        strSQL = "Select ID,���� From zlReports Where ID in (Select ����ID From zlRPTSubs Where ��ID=[1])"
        Set rsCheck = OpenSQLRecord(strSQL, Me.Caption, Val(Mid(lvwGroup.SelectedItem.Key, 2)))
        Do While Not rsCheck.EOF
            If Not CheckReportPriv(rsCheck!ID) Then
                MsgBox "��û��Ȩ�޲�ѯ����[" & rsCheck!���� & "]��ĳЩ����Դ�Ķ���", vbInformation, App.Title: Exit Sub
            End If
            rsCheck.MoveNext
        Loop
        lngRPTID = CLng(Mid(lvwGroup.SelectedItem.Key, 2))
    End If
    
    '1.ѡ��һ���˵�λ��
    Set rsTmp = GetMainTreeMenu
    If rsTmp Is Nothing Then MsgBox "��ȡ�˵���ϵʱ�����������,�������жϣ�", vbInformation, App.Title: Exit Sub
    
    Load frmSelTree
    If lvwGroup.SelectedItem.Key = "_-1" Or mcsActive = CS_���� Then
        frmSelTree.Caption = "������������̨ - �˵�λ��ѡ��"
    Else
        frmSelTree.Caption = "���������鵽����̨ - �˵�λ��ѡ��"
    End If
    With frmSelTree.tvw
        .Nodes.Clear
        For i = 1 To rsTmp.RecordCount
            If rsTmp!Flag = 0 Then
                Set objNode = .Nodes.Add(, , "_" & rsTmp!ID, rsTmp!����, "Root")
                objNode.Tag = "��ѡ��ϵͳ��һ������Ĳ˵�λ�ã�"
            Else
                If InStr(1, "888,999", rsTmp!Flag) = 0 Then
                    Set objNode = .Nodes.Add("_" & rsTmp!�ϼ�ID, 4, "_" & rsTmp!ID, rsTmp!����, "Path")
                Else
                    Set objNode = .Nodes.Add("_" & rsTmp!�ϼ�ID, 4, "_" & rsTmp!ID, rsTmp!����, IIF(rsTmp!Flag = 999, "GroupNode", "ReportNode"))
                    objNode.ForeColor = vbBlue
                    objNode.Tag = "�����ѷ����ı���,ѡ��һ���˵�λ�ã�"
                    
                    '���ܷ�������ͬλ��
                    If lvwGroup.SelectedItem.Key = "_-1" Or mcsActive = CS_���� Then
                        If objNode.Text = lvwReport.SelectedItem.Text Then
                            objNode.Parent.Tag = "ͬһ��������鲻�ܷ�������ͬ��λ��,��ѡ�������˵�λ�ã�"
                        End If
                    Else
                        If objNode.Text = lvwGroup.SelectedItem.Text Then
                            objNode.Parent.Tag = "ͬһ��������鲻�ܷ�������ͬ��λ��,��ѡ�������˵�λ�ã�"
                        End If
                    End If
                End If
            End If
            objNode.Expanded = True
            rsTmp.MoveNext
        Next
        If .Nodes.count > 0 Then .Nodes(1).Selected = True
    End With
    frmSelTree.Show 1, Me
    If Not gblnOK Then Exit Sub
    lngMenu = CLng(Mid(frmSelTree.tvw.SelectedItem.Key, 2)) 'Ҫ�Ӳ˵����ϼ�ID
    Unload frmSelTree
    
    lngNewMenu = GetNextID("zlMenus")
    
    '2.��д����Ȩ��
    If lvwGroup.SelectedItem.Key = "_-1" Or mcsActive = CS_���� Then
        If lvwReport.SelectedItem.Tag <> 0 Then
            '�����ٴ���
            lngProgID = lvwReport.SelectedItem.Tag
        Else
            lngProgID = GetNewProgID
            
            '�����ñ��������Դ���ʶ���
            strSQL = "Select ���� From zlRPTDatas Where ���� is Not NULL And ����ID=[1]"
            Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngRPTID)
            If Not rsTmp.EOF Then
                For i = 1 To rsTmp.RecordCount
                    For j = 0 To UBound(Split(rsTmp!����, ","))
                        If InStr(strObject & ",", "," & Split(rsTmp!����, ",")(j) & ",") = 0 Then
                            strObject = strObject & "," & Split(rsTmp!����, ",")(j)
                        End If
                    Next
                    rsTmp.MoveNext
                Next
            End If
            
            '�����ñ���Ĳ�������Դ���ʶ���
            strSQL = "Select B.���� From zlRPTDatas A,zlRPTPars B Where A.ID=B.ԴID And B.���� is Not NULL And A.����ID=[1]"
            Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngRPTID)
            If Not rsTmp.EOF Then
                For i = 1 To rsTmp.RecordCount
                    For j = 0 To UBound(Split(rsTmp!����, "|"))
                        strTmp = Split(rsTmp!����, "|")(j)
                        For k = 0 To UBound(Split(strTmp, ","))
                            If InStr(strObject & ",", "," & Split(strTmp, ",")(k) & ",") = 0 Then
                                strObject = strObject & "," & Split(strTmp, ",")(k)
                            End If
                        Next
                    Next
                    rsTmp.MoveNext
                Next
            End If
            
            If strObject <> "" Then strObject = Mid(strObject, 2)
        End If
    Else
        If lvwGroup.SelectedItem.Tag <> 0 Then
            '�����ٴ���
            lngProgID = lvwGroup.SelectedItem.Tag
        Else
            lngProgID = GetNewProgID
        End If
    End If
    
    lngSys = cboSys.ItemData(cboSys.ListIndex)
    
    On Error GoTo errH
    
    gcnOracle.BeginTrans
    
    If lvwGroup.SelectedItem.Key = "_-1" Or mcsActive = CS_���� Then
        If lvwReport.SelectedItem.Tag = 0 Then
            gcnOracle.Execute "Update zlReports Set ����='����',����ID=" & lngProgID & ",����ʱ��=Sysdate Where ID=" & lngRPTID
            gcnOracle.Execute "Insert Into zlPrograms(���,����,˵��,ϵͳ,����)" & _
                " Values(" & lngProgID & ",'" & lvwReport.SelectedItem.Text & "','" & lvwReport.SelectedItem.SubItems(RC_˵��) & "'," & _
                IIF(lngSys = 0, "NULL", lngSys) & ",'zl9Report')"
            gcnOracle.Execute "Insert Into zlProgFuncs(ϵͳ,���,����) Values(" & IIF(lngSys = 0, "NULL", lngSys) & "," & lngProgID & ",'����')"
            If strObject <> "" Then '�ñ���п��ܲ��������ݿ�
                For i = 0 To UBound(Split(strObject, ","))
                    strOwner = Left(Split(strObject, ",")(i), InStr(Split(strObject, ",")(i), ".") - 1)
                    If strOwner <> "SYS" And strOwner <> "ZLTOOLS" And strOwner <> "SYSTEM" Then
                        strName = Mid(Split(strObject, ",")(i), InStr(Split(strObject, ",")(i), ".") + 1)
                        gcnOracle.Execute "Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(" & _
                        IIF(lngSys = 0, "NULL", lngSys) & "," & lngProgID & ",'����','" & strName & "','" & strOwner & "','SELECT')"
                    End If
                Next
            End If
        Else
            gcnOracle.Execute "Update zlReports Set ����ʱ��=Sysdate Where ID=" & lngRPTID
        End If
    Else
        If lvwGroup.SelectedItem.Tag = 0 Then
            '���±������и��ӱ���Ĺ���Ϊ���ӱ��������
            gcnOracle.Execute "Update zlRPTSubs A Set ����=(Select ���� From zlReports Where ID=A.����ID) Where ��ID=" & lngRPTID
            gcnOracle.Execute "Update zlRPTGroups Set ����ID=" & lngProgID & ",����ʱ��=Sysdate Where ID=" & lngRPTID
            gcnOracle.Execute "Insert Into zlPrograms(���,����,˵��,ϵͳ,����) Values(" & lngProgID & "," & _
                " '" & lvwGroup.SelectedItem.Text & "','" & lvwGroup.SelectedItem.SubItems(GC_˵��) & "'," & _
                IIF(lngSys = 0, "NULL", lngSys) & ",'zl9Report')"
            gcnOracle.Execute "Insert Into zlProgFuncs(ϵͳ,���,����,˵��)" & _
                " Select " & IIF(lngSys = 0, "NULL", lngSys) & "," & lngProgID & ",����,˵�� From zlReports" & _
                " Where ID In (Select ����ID From zlRPTSubs Where ��ID=" & Mid(lvwGroup.SelectedItem.Key, 2) & ")"
            
            strSQL = "Select A.����ID,B.���� From zlRPTSubs A,zlReports B Where A.��ID=[1] And A.����ID=B.ID"
            Set rsSubRPT = OpenSQLRecord(strSQL, Me.Caption, lngRPTID)
            'ѭ����ȡ���ӱ����Ȩ��
            Do While Not rsSubRPT.EOF
                '�������ӱ��������Դ���ʶ���
                strObject = ""
                strSQL = "Select ���� From zlRPTDatas Where ���� is Not NULL And ����ID=[1]"
                Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, Val(rsSubRPT!����id))
                If Not rsTmp.EOF Then
                    For i = 1 To rsTmp.RecordCount
                        For j = 0 To UBound(Split(rsTmp!����, ","))
                            If InStr(strObject & ",", "," & Split(rsTmp!����, ",")(j) & ",") = 0 Then
                                strObject = strObject & "," & Split(rsTmp!����, ",")(j)
                            End If
                        Next
                        rsTmp.MoveNext
                    Next
                End If
                
                '�������ӱ���Ĳ�������Դ���ʶ���
                strSQL = "Select B.���� From zlRPTDatas A,zlRPTPars B Where A.ID=B.ԴID And B.���� is Not NULL And A.����ID=[1]"
                Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, Val(rsSubRPT!����id))
                If Not rsTmp.EOF Then
                    For i = 1 To rsTmp.RecordCount
                        For j = 0 To UBound(Split(rsTmp!����, "|"))
                            strTmp = Split(rsTmp!����, "|")(j)
                            For k = 0 To UBound(Split(strTmp, ","))
                                If InStr(strObject & ",", "," & Split(strTmp, ",")(k) & ",") = 0 Then
                                    strObject = strObject & "," & Split(strTmp, ",")(k)
                                End If
                            Next
                        Next
                        rsTmp.MoveNext
                    Next
                End If
                
                If strObject <> "" Then '�ñ���п��ܲ��������ݿ�
                    strObject = Mid(strObject, 2)
                    For i = 0 To UBound(Split(strObject, ","))
                        strOwner = Left(Split(strObject, ",")(i), InStr(Split(strObject, ",")(i), ".") - 1)
                        If strOwner <> "SYS" And strOwner <> "ZLTOOLS" And strOwner <> "SYSTEM" Then
                            strName = Mid(Split(strObject, ",")(i), InStr(Split(strObject, ",")(i), ".") + 1)
                            gcnOracle.Execute "Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(" & _
                            IIF(lngSys = 0, "NULL", lngSys) & "," & lngProgID & ",'" & rsSubRPT!���� & "','" & strName & "','" & strOwner & "','SELECT')"
                        End If
                    Next
                End If
                
                rsSubRPT.MoveNext
            Loop
        Else
            gcnOracle.Execute "Update zlRPTGroups Set ����ʱ��=Sysdate Where ID=" & lngRPTID
        End If
    End If
    
    '3.��д�˵�
    If lvwGroup.SelectedItem.Key = "_-1" Or mcsActive = CS_���� Then
        gcnOracle.Execute "Insert Into zlMenus(���,ID,�ϼ�ID,����,���,˵��,ϵͳ,ģ��,�̱���,ͼ��)" & _
            " Values('ȱʡ'," & lngNewMenu & "," & lngMenu & ",'" & lvwReport.SelectedItem.Text & "',NULL," & _
            "'" & lvwReport.SelectedItem.SubItems(RC_˵��) & "'," & IIF(lngSys = 0, "NULL", lngSys) & "," & _
            lngProgID & ",'" & lvwReport.SelectedItem.Text & "',105)"
    Else
        gcnOracle.Execute "Insert Into zlMenus(���,ID,�ϼ�ID,����,���,˵��,ϵͳ,ģ��,�̱���,ͼ��)" & _
            " Values('ȱʡ'," & lngNewMenu & "," & lngMenu & ",'" & lvwGroup.SelectedItem.Text & "',NULL," & _
            " '" & lvwGroup.SelectedItem.SubItems(GC_˵��) & "'," & IIF(lngSys = 0, "NULL", lngSys) & "," & _
            lngProgID & ",'" & lvwGroup.SelectedItem.Text & "',105)"
    End If
    
    gcnOracle.CommitTrans
    
    Set grsReport = Nothing '�������
    
    '4.���½���
    If lvwGroup.SelectedItem.Key = "_-1" Or mcsActive = CS_���� Then
        If lngSys = 0 Then
            If lvwReport.SelectedItem.SubItems(RC_����) = "Ʊ��" Then
                lvwReport.SelectedItem.Icon = "BillPublish"
                lvwReport.SelectedItem.SmallIcon = "BillPublish"
            Else
                lvwReport.SelectedItem.Icon = "Publish"
                lvwReport.SelectedItem.SmallIcon = "Publish"
            End If
        Else
            lvwReport.SelectedItem.Icon = "PubFixed"
            lvwReport.SelectedItem.SmallIcon = "PubFixed"
        End If
        lvwReport.SelectedItem.Tag = lngProgID
        lvwReport.SelectedItem.SubItems(RC_����ʱ��) = Format(Currentdate, "yyyy-MM-dd")
        Call lvwReport_ItemClick(lvwReport.SelectedItem)
    Else
        lvwGroup.SelectedItem.Icon = IIF(lngSys = 0, 2, 4)
        lvwGroup.SelectedItem.SmallIcon = IIF(lngSys = 0, 2, 4)
        lvwGroup.SelectedItem.Tag = lngProgID
        lvwGroup.SelectedItem.SubItems(GC_����ʱ��) = Format(Currentdate, "yyyy-MM-dd")
        Call LvwGroup_ItemClick(lvwGroup.SelectedItem)
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub

Private Sub ReportRevokeFromNavigator(Optional ByVal blnRevokeByProgram As Boolean = False)
'���ܣ�ȡ����ǰ����(��)�ڵ���̨�ϵ�һ������
'1:�������λ�ô���1,����ʹ����ѡ��ȡ��������һ��λ��,ɾ��zlMenus��Ӧλ������,���
'2:���ֻ��һ������λ��,��zlReport�еĳ���ID=NULL,ɾ��zlPrograms�еķ���ģ��,���
    Dim rsTmp As ADODB.Recordset
    Dim objNode As Node, lngSys As Long
    Dim lngProgID As Long, lngMenu As Long, i As Integer
    
    If lvwGroup.SelectedItem.Key = "_-1" Or mcsActive = CS_���� Then
        If lvwReport.SelectedItem Is Nothing Then MsgBox "��ǰû�б������ȡ��������", vbInformation, App.Title: Exit Sub
        If lvwReport.SelectedItem.Tag = 0 Then MsgBox "��ǰ����û�з���������̨�˵���", vbInformation, App.Title: Exit Sub
        If lvwReport.SelectedItem.Icon = "Fixed" Or lvwReport.SelectedItem.Icon = "PubFixed" Then
            MsgBox "�ñ���Ϊϵͳ���еı���,�������ܼ�����", vbInformation, App.Title: Exit Sub
        End If
    
        lngProgID = CLng(lvwReport.SelectedItem.Tag)
    Else
        If lvwGroup.SelectedItem Is Nothing Then MsgBox "��ǰû�б��������ȡ��������", vbInformation, App.Title: Exit Sub
        If lvwGroup.SelectedItem.Tag = 0 Then MsgBox "��ǰ������û�з�����", vbInformation, App.Title: Exit Sub
        If lvwGroup.SelectedItem.Icon = 3 Or lvwGroup.SelectedItem.Icon = 4 Then
            MsgBox "�ñ���Ϊϵͳ���еı�����,�������ܼ�����", vbInformation, App.Title: Exit Sub
        End If
        
        lngProgID = CLng(lvwGroup.SelectedItem.Tag)
    End If
    
    '1.������ǰ����λ��
    Set rsTmp = GetMainTreeMenu(lngProgID)
    If rsTmp Is Nothing Then MsgBox "��ȡ�˵���ϵʱ�����������,ȡ�������жϣ�", vbInformation, App.Title: Exit Sub
    
    rsTmp.Filter = "ģ��=" & lngProgID
    lngSys = cboSys.ItemData(cboSys.ListIndex)
    
    If rsTmp.EOF Then
        MsgBox "��ǰ����ķ������ڲ�����״̬,����������ݲ���ȷ����ģ�", vbInformation, App.Title
        On Error GoTo errH
        
        gcnOracle.BeginTrans
        If lvwGroup.SelectedItem.Key = "_-1" Or mcsActive = CS_���� Then
            gcnOracle.Execute "Update zlReports Set ����=NULL,����ID=NULL,����ʱ��=NULL Where ID=" & Mid(lvwReport.SelectedItem.Key, 2)
        Else
            gcnOracle.Execute "Update zlRPTGroups Set ����ID=NULL,����ʱ��=NULL Where ID=" & lngProgID
            gcnOracle.Execute " Update zlRPTSubs A Set ����=Null Where ��ID=" & Mid(lvwGroup.SelectedItem.Key, 2)
        End If
        gcnOracle.Execute "Delete From zlMenus Where ģ��=" & lngProgID & " And Nvl(ϵͳ,0)=" & lngSys
        gcnOracle.Execute "Delete From zlProgPrivs Where ���=" & lngProgID & " And Nvl(ϵͳ,0)=" & lngSys
        gcnOracle.Execute "Delete From zlProgFuncs Where ���=" & lngProgID & " And Nvl(ϵͳ,0)=" & lngSys
        gcnOracle.Execute "Delete From zlPrograms Where ���=" & lngProgID & " And Nvl(ϵͳ,0)=" & lngSys
        gcnOracle.Execute "Delete From zlRoleGrant Where ���=" & lngProgID & " And Nvl(ϵͳ,0)=" & lngSys
        
        gcnOracle.CommitTrans
        
        Set grsReport = Nothing '�������
        
        If lvwGroup.SelectedItem.Key = "_-1" Or mcsActive = CS_���� Then
            If lvwReport.SelectedItem.SubItems(RC_����) = "Ʊ��" Then
                lvwReport.SelectedItem.Icon = "Bill"
                lvwReport.SelectedItem.SmallIcon = "Bill"
            Else
                lvwReport.SelectedItem.Icon = "Report"
                lvwReport.SelectedItem.SmallIcon = "Report"
            End If
            lvwReport.SelectedItem.Tag = 0
            lvwReport.SelectedItem.SubItems(RC_����ʱ��) = ""
        Else
            lvwGroup.SelectedItem.Icon = 1
            lvwGroup.SelectedItem.SmallIcon = 1
            lvwGroup.SelectedItem.Tag = 0
            lvwGroup.SelectedItem.SubItems(GC_����ʱ��) = ""
        End If
    ElseIf rsTmp.RecordCount = 1 Then
        'ֻʣһ������λ��
        If Not blnRevokeByProgram Then
            If MsgBox("����Ѹñ���ӵ���̨�˵���ȡ�������������û�������ʹ�øñ���Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
        End If
        On Error GoTo errH
        
        gcnOracle.BeginTrans
        
        If lvwGroup.SelectedItem.Key = "_-1" Or mcsActive = CS_���� Then
            gcnOracle.Execute "Update zlReports Set ����=NULL,����ID=NULL,����ʱ��=NULL Where ID=" & Mid(lvwReport.SelectedItem.Key, 2)
        Else
            gcnOracle.Execute "Update zlRPTGroups Set ����ID=NULL,����ʱ��=NULL Where ID=" & Mid(lvwGroup.SelectedItem.Key, 2)
            gcnOracle.Execute " Update zlRPTSubs A Set ����=Null Where ��ID=" & Mid(lvwGroup.SelectedItem.Key, 2)
        End If
        gcnOracle.Execute "Delete From zlMenus Where ģ��=" & lngProgID & " And Nvl(ϵͳ,0)=" & lngSys
        gcnOracle.Execute "Delete From zlProgPrivs Where ���=" & lngProgID & " And Nvl(ϵͳ,0)=" & lngSys
        gcnOracle.Execute "Delete From zlProgFuncs Where ���=" & lngProgID & " And Nvl(ϵͳ,0)=" & lngSys
        gcnOracle.Execute "Delete From zlPrograms Where ���=" & lngProgID & " And Nvl(ϵͳ,0)=" & lngSys
        gcnOracle.Execute "Delete From zlRoleGrant Where ���=" & lngProgID & " And Nvl(ϵͳ,0)=" & lngSys
        
        gcnOracle.CommitTrans
        
        Set grsReport = Nothing '�������
        
        If lvwGroup.SelectedItem.Key = "_-1" Or mcsActive = CS_���� Then
            If lvwReport.SelectedItem.SubItems(RC_����) = "Ʊ��" Then
                lvwReport.SelectedItem.Icon = "Bill"
                lvwReport.SelectedItem.SmallIcon = "Bill"
            Else
                lvwReport.SelectedItem.Icon = "Report"
                lvwReport.SelectedItem.SmallIcon = "Report"
            End If
            lvwReport.SelectedItem.Tag = 0
            lvwReport.SelectedItem.SubItems(RC_����ʱ��) = ""
        Else
            lvwGroup.SelectedItem.Icon = 1
            lvwGroup.SelectedItem.SmallIcon = 1
            lvwGroup.SelectedItem.Tag = 0
            lvwGroup.SelectedItem.SubItems(GC_����ʱ��) = ""
        End If
    Else
        '���ж������λ��,ѡ����ȡ��
        rsTmp.Filter = 0
        
        Load frmSelTree
        frmSelTree.Caption = "ȡ������ - ����̨�˵�λ��"
        With frmSelTree.tvw
            .Nodes.Clear
            For i = 1 To rsTmp.RecordCount
                If rsTmp!Flag = 0 Then
                    Set objNode = .Nodes.Add(, , "_" & rsTmp!ID, rsTmp!����, "Root")
                    objNode.Tag = "���ڱ�ϵͳ��ѡ��һ��Ҫȡ�������ı�����飡"
                Else
                    If rsTmp!Flag <> 999 And rsTmp!Flag <> 888 Then
                        Set objNode = .Nodes.Add("_" & rsTmp!�ϼ�ID, 4, "_" & rsTmp!ID, rsTmp!����, "Path")
                        objNode.Tag = "���ڲ˵���ѡ��һ��Ҫȡ�������ı�����飡"
                    Else
                        Set objNode = .Nodes.Add("_" & rsTmp!�ϼ�ID, 4, "_" & rsTmp!ID, rsTmp!����, IIF(rsTmp!Flag = 999, "GroupNode", "ReportNode"))
                        objNode.ForeColor = vbBlue
                        If .SelectedItem Is Nothing Then
                            objNode.Selected = True
                        ElseIf .SelectedItem.Index = 1 Then
                            objNode.Selected = True
                        End If
                    End If
                End If
                objNode.Expanded = True
                
                '����б���(��)��·��
                If rsTmp!Flag = 999 Or rsTmp!Flag = 888 Then
                    objNode.SelectedImage = objNode.Image
                    Do While Not objNode.Parent Is Nothing
                        Set objNode = objNode.Parent
                        objNode.SelectedImage = objNode.Image
                    Loop
                End If
                
                rsTmp.MoveNext
            Next
            
            'ɾ��û�б���(��)��·��
            For i = .Nodes.count To 1 Step -1
                If .Nodes(i).SelectedImage = "" Then
                    .Nodes.Remove i
                End If
            Next
        End With
        frmSelTree.Show 1, Me
        If Not gblnOK Then Exit Sub
        lngMenu = CLng(Mid(frmSelTree.tvw.SelectedItem.Key, 2)) '����˵�ID
        Unload frmSelTree
        
        On Error GoTo errH
        
        gcnOracle.BeginTrans
        If lvwGroup.SelectedItem.Key = "_-1" Or mcsActive = CS_���� Then
            gcnOracle.Execute "Update zlReports Set ����ʱ��=Sysdate Where ID=" & Mid(lvwReport.SelectedItem.Key, 2)
        Else
            gcnOracle.Execute "Update zlRPTGroups Set ����ʱ��=Sysdate Where ID=" & Mid(lvwGroup.SelectedItem.Key, 2)
            gcnOracle.Execute "Update zlRPTSubs A Set ����=Null Where ��ID=" & Mid(lvwGroup.SelectedItem.Key, 2)
        End If
        'ֻ��ɾ���˵�����
        gcnOracle.Execute "Delete From zlMenus Where ID=" & lngMenu & " And Nvl(ϵͳ,0)=" & lngSys
        gcnOracle.CommitTrans
        
        Set grsReport = Nothing '�������
        
        If lvwGroup.SelectedItem.Key = "_-1" Or mcsActive = CS_���� Then
            lvwReport.SelectedItem.SubItems(RC_����ʱ��) = Format(Currentdate, "yyyy-MM-dd")
        Else
            lvwGroup.SelectedItem.SubItems(GC_����ʱ��) = Format(Currentdate, "yyyy-MM-dd")
        End If
    End If
    
    If lvwGroup.SelectedItem.Key = "_-1" Or mcsActive = CS_���� Then
        Call lvwReport_ItemClick(lvwReport.SelectedItem)
    Else
        Call LvwGroup_ItemClick(lvwGroup.SelectedItem)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub

Public Sub ReportGrantToNavigatorAgain(ByVal objItem As ListItem)
'���ܣ����ݱ��������ӱ����ɾ��������¸���ָ���������顱�ķ�����Ȩ���
    Dim rsTmp As New ADODB.Recordset
    Dim rsSubRPT As New ADODB.Recordset
    Dim strTmp As String, strOwner As String, strName As String
    Dim strSQL As String, i As Integer, j As Integer, k As Integer
    Dim strObject As String, lngSys As Long, lngProgID As Long
    
    lngProgID = Val(objItem.Tag)
    If lngProgID = 0 Then Exit Sub
    lngSys = cboSys.ItemData(cboSys.ListIndex)
    
    On Error GoTo errH
    
    gcnOracle.BeginTrans
    gcnOracle.Execute "Update zlRPTGroups Set ����ID=" & lngProgID & ",����ʱ��=Sysdate Where ID=" & Mid(objItem.Key, 2)
    gcnOracle.Execute "Delete zlProgFuncs Where ���=" & lngProgID & " And Nvl(ϵͳ,0)=" & lngSys
    gcnOracle.Execute "Delete zlProgPrivs Where ���=" & lngProgID & " And Nvl(ϵͳ,0)=" & lngSys
    gcnOracle.Execute "Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Select " & IIF(lngSys = 0, "NULL", lngSys) & "," & _
        lngProgID & ",����,˵�� From zlReports Where ID In (Select ����ID From zlRPTSubs Where ��ID=" & Mid(objItem.Key, 2) & ")"
            
    strSQL = "Select A.����ID,B.���� From zlRPTSubs A,zlReports B Where A.��ID=[1] And A.����ID=B.ID"
    Set rsSubRPT = OpenSQLRecord(strSQL, Me.Caption, Val(Mid(objItem.Key, 2)))
    'ѭ����ȡ���ӱ����Ȩ��
    Do While Not rsSubRPT.EOF
        '�������ӱ��������Դ���ʶ���
        strObject = ""
        strSQL = "Select ���� From zlRPTDatas Where ���� is Not NULL And ����ID=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, Val(rsSubRPT!����id))
        If Not rsTmp.EOF Then
            For i = 1 To rsTmp.RecordCount
                For j = 0 To UBound(Split(rsTmp!����, ","))
                    If InStr(strObject & ",", "," & Split(rsTmp!����, ",")(j) & ",") = 0 Then
                        strObject = strObject & "," & Split(rsTmp!����, ",")(j)
                    End If
                Next
                rsTmp.MoveNext
            Next
        End If
        
        '�������ӱ���Ĳ�������Դ���ʶ���
        strSQL = "Select B.���� From zlRPTDatas A,zlRPTPars B Where A.ID=B.ԴID And B.���� is Not NULL And A.����ID=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, Val(rsSubRPT!����id))
        If Not rsTmp.EOF Then
            For i = 1 To rsTmp.RecordCount
                For j = 0 To UBound(Split(rsTmp!����, "|"))
                    strTmp = Split(rsTmp!����, "|")(j)
                    For k = 0 To UBound(Split(strTmp, ","))
                        If InStr(strObject & ",", "," & Split(strTmp, ",")(k) & ",") = 0 Then
                            strObject = strObject & "," & Split(strTmp, ",")(k)
                        End If
                    Next
                Next
                rsTmp.MoveNext
            Next
        End If
        
        If strObject <> "" Then '�ñ���п��ܲ��������ݿ�
            strObject = Mid(strObject, 2)
            For i = 0 To UBound(Split(strObject, ","))
                strOwner = Left(Split(strObject, ",")(i), InStr(Split(strObject, ",")(i), ".") - 1)
                If strOwner <> "SYS" And strOwner <> "ZLTOOLS" And strOwner <> "SYSTEM" Then
                    strName = Mid(Split(strObject, ",")(i), InStr(Split(strObject, ",")(i), ".") + 1)
                    gcnOracle.Execute "Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(" & _
                        IIF(lngSys = 0, "NULL", lngSys) & "," & lngProgID & ",'" & rsSubRPT!���� & "'," & _
                        "'" & strName & "','" & strOwner & "','SELECT')"
                End If
            Next
        End If
        rsSubRPT.MoveNext
    Loop
    gcnOracle.CommitTrans
    
    Set grsReport = Nothing '�������
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub

Private Sub ReportGrantToModule()
'���ܣ�������ǰ����ģ��,���ܲ��ǵ�һ��
    Dim rsTmp As ADODB.Recordset
    Dim objNode As Node, strSQL As String
    Dim strObject As String, strOwner As String, strName As String
    Dim lngRPTID As Long, lngSys As Long, lngProgID As Long
    Dim i As Integer, j As Integer, k As Integer
    Dim strFunc As String, blnTran As Boolean
    
    '��ǰ�о��屨����ѡ��ʱ����֧�ֱ����鷢����ģ��
    If Val(Mid(lvwGroup.SelectedItem.Key, 2)) <> -1 And mcsActive = CS_������ Then Exit Sub
    
    If lvwReport.SelectedItem Is Nothing Then
        MsgBox "��ǰû�б�����Է�����", vbInformation, App.Title: Exit Sub
    End If
    If CheckPass(CLng(Mid(lvwReport.SelectedItem.Key, 2))) = False Then
        MsgBox "�������ݴ��󣬲��ܷ����ñ���", vbInformation, App.Title: Exit Sub
    End If
    If Not CheckReportPriv(Val(Mid(lvwReport.SelectedItem.Key, 2))) Then
        MsgBox "��û��Ȩ�޲�ѯ�ñ���ĳЩ����Դ�еĶ���,�������ܼ�����", vbInformation, App.Title
        Exit Sub
    End If
    lngRPTID = CLng(Mid(lvwReport.SelectedItem.Key, 2))
    
    On Error GoTo errH
    
    '1.ѡ��һ���˵�ģ��λ��
    '----------------------------------------------------
    Set rsTmp = GetModuleTreeMenu(lngRPTID)
    If rsTmp Is Nothing Then
        MsgBox "��ȡģ��˵���ϵʱ�����������,�������жϣ�", vbInformation, App.Title: Exit Sub
    End If
    Load frmSelTree
    frmSelTree.Caption = "��������ģ�� - ģ��λ��ѡ��"
    With frmSelTree.tvw
        .Nodes.Clear
        For i = 1 To rsTmp.RecordCount
            If IsNull(rsTmp!�ϼ�ID) Then
                Set objNode = .Nodes.Add(, , "_" & rsTmp!ID, rsTmp!����)
            Else
                Set objNode = .Nodes.Add("_" & rsTmp!�ϼ�ID, 4, "_" & rsTmp!ID, rsTmp!����)
            End If
            If Left(rsTmp!ID, 1) = "S" Then 'System
                objNode.Image = "Root"
                objNode.Tag = "��ѡ��ϵͳ�в˵��µ�ģ��λ�á�"
            ElseIf Left(rsTmp!ID, 1) = "T" Then 'MenuTree
                objNode.Image = "Path"
                objNode.Tag = "��ѡ��ϵͳ�в˵��µ�ģ��λ�á�"
            ElseIf Left(rsTmp!ID, 1) = "M" Then 'Module
                objNode.Image = "App"
            ElseIf Left(rsTmp!ID, 1) = "R" Then 'Report
                objNode.Image = "ReportNode"
                objNode.ForeColor = vbBlue
                objNode.Tag = "�����ѷ����ı���,ѡ�������˵��µ�ģ��λ�á�"
                objNode.Parent.Tag = "�������ظ�������ͬһ��ģ��,��ѡ������ģ�顣"
            End If
            objNode.Expanded = True
            
            '������¼�ģ��Ĳ˵�(��SQL����)
            If Left(rsTmp!ID, 1) = "M" Then
                If objNode.Parent.SelectedImage = "" Then
                    Do While Not objNode.Parent Is Nothing
                        Set objNode = objNode.Parent
                        objNode.SelectedImage = objNode.Image
                    Loop
                End If
            End If
            
            rsTmp.MoveNext
        Next
        
        'ɾ�����¼�ģ��Ŀղ˵�
        For i = .Nodes.count To 1 Step -1
            If .Nodes(i).SelectedImage = "" And Mid(.Nodes(i).Key, 2, 1) = "T" Then
                .Nodes.Remove i
            End If
        Next
        
        If .Nodes.count > 0 Then .Nodes(1).Selected = True
    End With
    frmSelTree.Show 1, Me
    If Not gblnOK Then Exit Sub
    rsTmp.Filter = "ID='" & Mid(frmSelTree.tvw.SelectedItem.Key, 2) & "'"
    If rsTmp.EOF Then Exit Sub
    lngSys = rsTmp!ϵͳ: lngProgID = rsTmp!����ID
    strFunc = lvwReport.SelectedItem.Text
    Unload frmSelTree
        
    '�����ظ����
    strSQL = _
        " Select ���� From zlRPTPuts Where ����ID=[1] And ϵͳ=[2] And ����ID=[3]" & _
        " Union ALL " & _
        " Select ���� From zlProgFuncs Where ϵͳ=[2] And ���=[3] And ����=[4]"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngRPTID, lngSys, lngProgID, strFunc)
    If Not rsTmp.EOF Then
        MsgBox "������λ�û򷢲������ظ������ݿ��е����ݿ��ܲ���ȷ��", vbInformation, App.Title
        Exit Sub
    End If
    
    '2.��ȨȨ�޷���
    '----------------------------------------------------
    '�����ñ��������Դ���ʶ���
    strSQL = "Select ���� From zlRPTDatas Where ���� is Not NULL And ����ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngRPTID)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            For j = 0 To UBound(Split(rsTmp!����, ","))
                If InStr(strObject & ",", "," & Split(rsTmp!����, ",")(j) & ",") = 0 Then
                    strObject = strObject & "," & Split(rsTmp!����, ",")(j)
                End If
            Next
            rsTmp.MoveNext
        Next
    End If
    
    '�����ñ���Ĳ�������Դ���ʶ���
    strSQL = "Select B.���� From zlRPTDatas A,zlRPTPars B Where A.ID=B.ԴID And B.���� is Not NULL And A.����ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngRPTID)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            For j = 0 To UBound(Split(rsTmp!����, "|"))
                strName = Split(rsTmp!����, "|")(j)
                For k = 0 To UBound(Split(strName, ","))
                    If InStr(strObject & ",", "," & Split(strName, ",")(k) & ",") = 0 Then
                        strObject = strObject & "," & Split(strName, ",")(k)
                    End If
                Next
            Next
            rsTmp.MoveNext
        Next
    End If
    If strObject <> "" Then strObject = Mid(strObject, 2)
        
    '3.��д����Ȩ��
    '----------------------------------------------------
    gcnOracle.BeginTrans: blnTran = True
    
    gcnOracle.Execute "Update zlReports Set ����ʱ��=Sysdate Where ID=" & lngRPTID
    gcnOracle.Execute "Insert Into zlRPTPuts(����ID,ϵͳ,����ID,����) Values(" & _
        lngRPTID & "," & lngSys & "," & lngProgID & ",'" & strFunc & "')"
    gcnOracle.Execute "Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(" & _
        lngSys & "," & lngProgID & ",'" & strFunc & "','" & lvwReport.SelectedItem.SubItems(RC_˵��) & "')"
    If strObject <> "" Then '�ñ���п��ܲ��������ݿ�
        For i = 0 To UBound(Split(strObject, ","))
            strOwner = Left(Split(strObject, ",")(i), InStr(Split(strObject, ",")(i), ".") - 1)
            If strOwner <> "SYS" And strOwner <> "ZLTOOLS" And strOwner <> "SYSTEM" Then
                strName = Mid(Split(strObject, ",")(i), InStr(Split(strObject, ",")(i), ".") + 1)
                gcnOracle.Execute "Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(" & _
                lngSys & "," & lngProgID & ",'" & strFunc & "','" & strName & "','" & strOwner & "','SELECT')"
            End If
        Next
    End If
    
    gcnOracle.CommitTrans: blnTran = False
    
    Set grsReport = Nothing '�������
    
    '4.���½���
    If cboSys.ItemData(cboSys.ListIndex) = 0 Then
        If lvwReport.SelectedItem.SubItems(RC_����) = "Ʊ��" Then
            lvwReport.SelectedItem.Icon = "BillPublish"
            lvwReport.SelectedItem.SmallIcon = "BillPublish"
        Else
            lvwReport.SelectedItem.Icon = "Publish"
            lvwReport.SelectedItem.SmallIcon = "Publish"
        End If
    Else
        lvwReport.SelectedItem.Icon = "PubFixed"
        lvwReport.SelectedItem.SmallIcon = "PubFixed"
    End If
    lvwReport.SelectedItem.SubItems(RC_����ʱ��) = Format(Currentdate, "yyyy-MM-dd")
    Call lvwReport_ItemClick(lvwReport.SelectedItem)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    If blnTran Then gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub

Private Sub ReportRevokeFromModule()
'���ܣ�ȡ����ǰ������ģ���ϵ�һ������
'1:�������λ�ô���1,����ʹ����ѡ��ȡ��������һ��λ��
'2:���ֻ��һ������λ��,��ֱ����ʾ����
    Dim rsTmp As ADODB.Recordset, strFunc As String
    Dim objNode As Node, blnTran As Boolean
    Dim lngRPTID As Long, lngSys As Long, lngProgID As Long
    Dim strSQL As String, i As Integer
    
    If lvwReport.SelectedItem Is Nothing Then
        MsgBox "��ǰû�б������ȡ��������", vbInformation, App.Title: Exit Sub
    End If
    lngRPTID = CLng(Mid(lvwReport.SelectedItem.Key, 2))
        
    On Error GoTo errH
    
    '1.������ǰ����λ��
    strSQL = "Select ϵͳ,����ID,���� From zlRPTPuts Where ����ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngRPTID)
    If rsTmp.EOF Then
        MsgBox "��ǰ����û�з�����ģ���С�", vbInformation, App.Title: Exit Sub
    ElseIf rsTmp.RecordCount = 1 Then
        'ֻʣһ������λ��
        If MsgBox("����ѱ���Ӹ�ģ��˵���ȡ�������������û�������ʹ�øñ���Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
        
        lngSys = rsTmp!ϵͳ: lngProgID = rsTmp!����ID: strFunc = rsTmp!����
        
        gcnOracle.BeginTrans: blnTran = True

        gcnOracle.Execute "Update zlReports Set ����ʱ��=NULL Where ����ID Is Null And ID=" & lngRPTID
        gcnOracle.Execute "Delete From zlRPTPuts Where ����ID=" & lngRPTID & " And ϵͳ=" & lngSys & " And ����ID=" & lngProgID
        gcnOracle.Execute "Delete From zlProgPrivs Where ϵͳ=" & lngSys & " And ���=" & lngProgID & " And ����='" & strFunc & "'"
        gcnOracle.Execute "Delete From zlProgFuncs Where ϵͳ=" & lngSys & " And ���=" & lngProgID & " And ����='" & strFunc & "'"
        gcnOracle.Execute "Delete From zlRoleGrant Where ϵͳ=" & lngSys & " And ���=" & lngProgID & " And ����='" & strFunc & "'"
        
        gcnOracle.CommitTrans: blnTran = False
        
        Set grsReport = Nothing '�������
        
        If Val(lvwReport.SelectedItem.Tag) = 0 Then
            If cboSys.ItemData(cboSys.ListIndex) = 0 Then
                If lvwReport.SelectedItem.SubItems(RC_����) = "Ʊ��" Then
                    lvwReport.SelectedItem.Icon = "Bill"
                    lvwReport.SelectedItem.SmallIcon = "Bill"
                Else
                    lvwReport.SelectedItem.Icon = "Report"
                    lvwReport.SelectedItem.SmallIcon = "Report"
                End If
            Else
                lvwReport.SelectedItem.Icon = "Fixed"
                lvwReport.SelectedItem.SmallIcon = "Fixed"
            End If
            lvwReport.SelectedItem.SubItems(RC_����ʱ��) = ""
        End If
    Else
        '���ж������λ��,ѡ����ȡ��
        Set rsTmp = GetModuleTreeMenu(lngRPTID)
        If rsTmp Is Nothing Then
            MsgBox "��ȡģ��˵���ϵʱ�����������,�������жϣ�", vbInformation, App.Title: Exit Sub
        End If
        Load frmSelTree
        frmSelTree.Caption = "ȡ������ - ģ��˵�λ��"
        With frmSelTree.tvw
            .Nodes.Clear
            For i = 1 To rsTmp.RecordCount
                If IsNull(rsTmp!�ϼ�ID) Then
                    Set objNode = .Nodes.Add(, , "_" & rsTmp!ID, rsTmp!����)
                Else
                    Set objNode = .Nodes.Add("_" & rsTmp!�ϼ�ID, 4, "_" & rsTmp!ID, rsTmp!����)
                End If
                If Left(rsTmp!ID, 1) = "S" Then 'System
                    objNode.Image = "Root"
                    objNode.Tag = "��ѡ��Ҫȡ�������ı���"
                ElseIf Left(rsTmp!ID, 1) = "T" Then 'MenuTree
                    objNode.Image = "Path"
                    objNode.Tag = "��ѡ��Ҫȡ�������ı���"
                ElseIf Left(rsTmp!ID, 1) = "M" Then 'Module
                    objNode.Image = "App"
                    objNode.Tag = "��ѡ��Ҫȡ�������ı���"
                ElseIf Left(rsTmp!ID, 1) = "R" Then 'Report
                    objNode.Image = "ReportNode"
                    objNode.ForeColor = vbBlue
                End If
                objNode.Expanded = True
                
                '����з���������ϼ�
                If Left(rsTmp!ID, 1) = "R" Then
                    objNode.SelectedImage = objNode.Image
                    If objNode.Parent.SelectedImage = "" Then
                        Do While Not objNode.Parent Is Nothing
                            Set objNode = objNode.Parent
                            objNode.SelectedImage = objNode.Image
                        Loop
                    End If
                End If
                
                rsTmp.MoveNext
            Next
            
            'ɾ���޷��������·��
            For i = .Nodes.count To 1 Step -1
                If .Nodes(i).SelectedImage = "" Then
                    .Nodes.Remove i
                End If
            Next
            
            If .Nodes.count > 0 Then .Nodes(1).Selected = True
        End With
        frmSelTree.Show 1, Me
        If Not gblnOK Then Exit Sub
        rsTmp.Filter = "ID='" & Mid(frmSelTree.tvw.SelectedItem.Key, 2) & "'"
        If rsTmp.EOF Then Exit Sub
        lngSys = rsTmp!ϵͳ: lngProgID = rsTmp!����ID: strFunc = rsTmp!����
        Unload frmSelTree
        
        gcnOracle.BeginTrans: blnTran = True

        gcnOracle.Execute "Delete From zlRPTPuts Where ����ID=" & lngRPTID & " And ϵͳ=" & lngSys & " And ����ID=" & lngProgID
        gcnOracle.Execute "Delete From zlProgPrivs Where ϵͳ=" & lngSys & " And ���=" & lngProgID & " And ����='" & strFunc & "'"
        gcnOracle.Execute "Delete From zlProgFuncs Where ϵͳ=" & lngSys & " And ���=" & lngProgID & " And ����='" & strFunc & "'"
        gcnOracle.Execute "Delete From zlRoleGrant Where ϵͳ=" & lngSys & " And ���=" & lngProgID & " And ����='" & strFunc & "'"
        
        gcnOracle.CommitTrans: blnTran = False
        
        Set grsReport = Nothing '�������
    End If
    
    Call lvwReport_ItemClick(lvwReport.SelectedItem)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    If blnTran Then gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub

Private Sub SetFuncEnabled(ByVal blnReport As Boolean)
'���ܣ����ò˵�����ť����״̬
'������blnReport=��ǰ��б��Ƿ񱨱��б�
    Dim blnFree As Boolean
    
    On Error Resume Next
    
    'ָ��ϵͳ,ֻ��ִ���޸����ơ�ִ�м���ƵĹ���
    blnFree = cboSys.ItemData(cboSys.ListIndex) = 0
        
    '���ù��ܿɼ���
    '------------------------------------------------------------
    mnuEdit_Group_Add.Visible = Not blnReport And mblnGrant
    mnuEdit_Group_Delete.Visible = Not blnReport And mblnGrant
    mnuEdit_Group_Modify.Visible = Not blnReport
    mnuEdit_Group_Setup.Visible = Not blnReport
    mnuEdit_Group_.Visible = Not blnReport
    tbr.Buttons("GroupAdd").Visible = Not blnReport And mblnGrant
    tbr.Buttons("GroupDel").Visible = Not blnReport And mblnGrant
    tbr.Buttons("GroupModify").Visible = Not blnReport
    tbr.Buttons("Group_").Visible = Not blnReport
    
    mnuEdit_Add.Visible = blnReport And mblnGrant
    mnuEdit_Del.Visible = blnReport And mblnGrant
    mnuEdit_Modi.Visible = blnReport
    mnuEdit_Report_.Visible = blnReport
    tbr.Buttons("Add").Visible = blnReport And mblnGrant
    tbr.Buttons("Del").Visible = blnReport And mblnGrant
    tbr.Buttons("Modi").Visible = blnReport
    tbr.Buttons("Report_").Visible = blnReport
    
    '���ù��ܿ�����
    '------------------------------------------------------------
    If Me.ActiveControl Is lvwReport Or lvwGroup.SelectedItem.Key = "_-1" Then
        mnuFile_Report.Enabled = Not lvwReport.SelectedItem Is Nothing
    Else
        mnuFile_Report.Enabled = True
    End If
    tbr.Buttons("Report").Enabled = mnuFile_Report.Enabled
    
    If Not blnFree Then
        mnuFile_Imp.Enabled = True
        mnuFile_ImpAll.Enabled = True
    Else
        mnuFile_Imp.Enabled = mblnGrant
        mnuFile_ImpAll.Enabled = mblnGrant
    End If
    mnuFile_Exp.Enabled = Not lvwReport.SelectedItem Is Nothing
    mnuEdit_Add.Enabled = blnFree
    mnuEdit_Modi.Enabled = Not lvwReport.SelectedItem Is Nothing
    mnuEdit_Del.Enabled = blnFree And Not lvwReport.SelectedItem Is Nothing
    tbr.Buttons("Add").Enabled = mnuEdit_Add.Enabled
    tbr.Buttons("Modi").Enabled = mnuEdit_Modi.Enabled
    tbr.Buttons("Del").Enabled = mnuEdit_Del.Enabled
    
    mnuEdit_Group_Add.Enabled = blnFree
    mnuEdit_Group_Modify.Enabled = lvwGroup.SelectedItem.Key <> "_-1"
    mnuEdit_Group_Delete.Enabled = blnFree And lvwGroup.SelectedItem.Key <> "_-1"
    tbr.Buttons("GroupAdd").Enabled = mnuEdit_Group_Add.Enabled
    tbr.Buttons("GroupModify").Enabled = mnuEdit_Group_Modify.Enabled
    tbr.Buttons("GroupDel").Enabled = mnuEdit_Group_Delete.Enabled
    
    mnuEdit_Group_Setup.Enabled = blnFree And lvwGroup.SelectedItem.Key <> "_-1"
    
    mnuEdit_Design.Enabled = Not lvwReport.SelectedItem Is Nothing
    tbr.Buttons("Design").Enabled = mnuEdit_Design.Enabled
    
    mnuEdit_Guide.Enabled = blnFree
    tbr.Buttons("Guide").Enabled = mnuEdit_Guide.Enabled
        
    '����������Է���,�����鲻�ܷ�����ģ��
    mnuEdit_Publish.Enabled = blnFree
    mnuEdit_unPub.Enabled = blnFree
    '�жϷ���������ͷ�������ѡ�����ʾ״̬
    If mcsActive = CS_���� Then
        If lvwGroup.SelectedItem.Key = "_-1" Then
            mnuEdit_Group_Publish.Visible = False
            mnuEdit_Group_unPub.Visible = False
        End If
        mnuEdit_Publish.Visible = True
        mnuEdit_unPub.Visible = True
    Else
        mnuEdit_Group_Publish.Visible = True And blnFree
        mnuEdit_Group_unPub.Visible = True And blnFree
        mnuEdit_Publish.Visible = False
        mnuEdit_unPub.Visible = False
    End If
    mnuPopPublish_Group.Visible = mnuEdit_Group_Publish.Visible
    mnuPopUnpub_Group.Visible = mnuEdit_Group_unPub.Visible
    mnuPopUnpub_ReportMain.Visible = mnuEdit_unPub_Module.Enabled
    mnuPopUnpub_ReportModule.Visible = mnuEdit_unPub_Module.Enabled
    mnuPopPublish_ReportMain.Visible = mnuEdit_Publish_Module.Enabled
    mnuPopPublish_ReportModule.Visible = mnuEdit_Publish_Module.Enabled
    If lvwGroup.SelectedItem.Key <> "_-1" Or mcsActive = CS_���� Then
        '�������屨���鵽����̨
        mnuEdit_Publish_Main.Enabled = blnFree
        mnuEdit_unPub_Main.Enabled = blnFree
        '�������屨���鵽ģ��
        mnuEdit_Publish_Module.Enabled = IIF(mcsActive = CS_������, False, True)
        mnuEdit_unPub_Module = IIF(mcsActive = CS_������, False, True)
    ElseIf lvwGroup.SelectedItem.Key = "_-1" And mcsActive = CS_������ Then
        mnuEdit_Group_Publish.Visible = False
        mnuEdit_Group_unPub.Visible = False
        mnuEdit_Publish.Visible = False
        mnuEdit_unPub.Visible = False
    Else
        '�������屨������̨
        mnuEdit_Publish_Main.Enabled = blnFree And Not lvwReport.SelectedItem Is Nothing
        mnuEdit_unPub_Main.Enabled = blnFree And Not lvwReport.SelectedItem Is Nothing
        
        '�������屨���鵽ģ��
        mnuEdit_Publish_Module.Enabled = mblnModule And blnFree And Not lvwReport.SelectedItem Is Nothing
        mnuEdit_unPub_Module = mblnModule And blnFree And Not lvwReport.SelectedItem Is Nothing
    End If
    If Not mnuEdit_Publish_Main.Enabled And Not mnuEdit_Publish_Module.Enabled Then
        mnuEdit_Publish.Enabled = False
    End If
    If Not mnuEdit_unPub_Main.Enabled And Not mnuEdit_unPub_Module.Enabled Then
        mnuEdit_unPub.Enabled = False
    End If
    tbr.Buttons("Publish").Enabled = (mnuEdit_Group_Publish.Visible Or mnuEdit_Publish.Visible) And blnFree
    tbr.Buttons("unPub").Enabled = (mnuEdit_Group_unPub.Visible Or mnuEdit_Publish.Visible) And blnFree
    '���⼰��ͼ
    '------------------------------------------------------------
    If blnReport Or lvwGroup.SelectedItem.Key = "_-1" Then
        mnuFile_Report.Caption = "ִ�б���(&E)"
        tbr.Buttons("Report").ToolTipText = "ִ�б���"
        tbr.Buttons("Report").Image = 1
    Else
        mnuFile_Report.Caption = "ִ�б�����(&E)"
        tbr.Buttons("Report").ToolTipText = "ִ�б�����"
        tbr.Buttons("Report").Image = 15
    End If
    
    mnuEdit_Del.Caption = IIF(lvwGroup.SelectedItem.Key <> "_-1", "�Ƴ�", "ɾ��")
    Me.tbr.Buttons("Del").Caption = mnuEdit_Del.Caption
    Me.tbr.Buttons("Del").Tag = mnuEdit_Del.Caption
    Me.tbr.Buttons("Del").ToolTipText = mnuEdit_Del.Caption & "��ǰ����"
    mnuEdit_Del.Caption = mnuEdit_Del.Caption & "����(&D)"
    
    mnuView_View(0).Checked = False
    mnuView_View(1).Checked = False
    mnuView_View(2).Checked = False
    mnuView_View(3).Checked = False
    If blnReport Then
        mnuView_View(lvwReport.View).Checked = True
    Else
        mnuView_View(lvwGroup.View).Checked = True
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

Private Sub tbrCheck_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Key = "Check" Then
        Call CheckSQLPlanEx
    End If
End Sub

Private Sub CheckSQLPlanEx()
'���ܣ���鵱ǰ�б��еı���ִ�мƻ��Ƿ������������
    Dim i As Long, objReport As Report, objData As RPTData
    Dim strSQLCheck As String, strErr As String, strFields As String
    Dim strMsg As String, objPar As RPTPar, strSQL As String
    Dim lngCount As Long
    
    If MsgBox("��ǰĿ¼һ��" & lvwReport.ListItems.count & "�ű�����������Щ����(������)����Դ�е�SQL����ִ�мƻ���" & _
         "Ȼ����ִ�мƻ��Ƿ�������������" & vbCrLf & _
         "    1.�������ͱ��ȫ��ɨ��;" & vbCrLf & _
         "    2.�������ͱ������ȫɨ�����Ծʽ����ɨ��;" & vbCrLf & _
         "    3.��������û������Ǵ���������������������ҽ����¼_IX_������ĿID��;" & vbCrLf & _
         "    ���д����ָzlBakTables ZlBigTables�ж���ı�;" & vbCrLf & _
         "    ���ͱ���ָ�ռ�ͳ����Ϣ���¼����ȱʡ��3ǧ��1����֮��ı� (����ƽ����ִ�мƻ��鿴�п����¶���);" & vbCrLf & vbCrLf & _
         "�˹��̿��ܻỨ�Ѽ����ӵ�ʱ�䣬��ȷ��Ҫ������" _
         , vbQuestion + vbOKCancel + vbDefaultButton1, "���ܼ��") = vbCancel Then Exit Sub
    
    If lvwReport.ColumnHeaders(RC_������������Դ + 1).Width = 0 Then lvwReport.ColumnHeaders(RC_������������Դ + 1).Width = 3440

    For i = 1 To lvwReport.ListItems.count
        Set objReport = ReadReport(Val(Mid(lvwReport.ListItems(i).Key, 2)), , True)
        strMsg = ""
        For Each objData In objReport.Datas
            With objData
                '�ȼ������Դ��SQL
                strSQLCheck = ""
                strFields = ""
                strSQL = RemoveNote(.SQL)
                strSQL = TrimChar(strSQL)
                strSQL = Replace(strSQL, "[ϵͳ]", cboSys.ItemData(cboSys.ListIndex))
                If GetParCount(strSQL) = 0 Then
                    strFields = CheckSQL(strSQL, strErr, , strSQLCheck, , objReport.Datas, .�������ӱ��)
                Else
                    strFields = CheckSQL(strSQL, strErr, ReplaceParSysNo(.Pars, cboSys.ItemData(cboSys.ListIndex)) _
                        , strSQLCheck, , objReport.Datas, .�������ӱ��)
                End If
                If strFields <> "" Then
                    If strSQLCheck <> "" Then
                        If CheckSQLPlan(strSQLCheck, , .�������ӱ��) = True Then
                            strMsg = strMsg & "," & .����
                        End If
                    End If
                End If
                '�ټ�������ϸ�ͷ���SQL
                For Each objPar In .Pars
                    '�ų��Ѿ�������
                    If objPar.����SQL <> "" And InStr(strMsg, "(" & objPar.���� & ")[����]") = 0 Then
                        strSQLCheck = ""
                        strFields = ""
                        strSQL = RemoveNote(objPar.����SQL)
                        strSQL = TrimChar(strSQL)
                        strSQL = Replace(strSQL, "[ϵͳ]", cboSys.ItemData(cboSys.ListIndex))
                        Call CheckParsRela(strSQL, objReport.Datas, objPar.����, True)
                        strFields = CheckSQL(strSQL, strErr, , strSQLCheck, , objReport.Datas, .�������ӱ��)
                        If strFields <> "" Then
                            If strSQLCheck <> "" Then
                                If CheckSQLPlan(strSQLCheck, , .�������ӱ��) = True Then
                                    strMsg = strMsg & "," & .���� & "(" & objPar.���� & ")[����]"
                                End If
                            End If
                        End If
                    End If
                    
                    If objPar.��ϸSQL <> "" And InStr(strMsg, "(" & objPar.���� & ")[��ϸ]") = 0 Then
                        strSQLCheck = ""
                        strFields = ""
                        strSQL = RemoveNote(objPar.��ϸSQL)
                        strSQL = TrimChar(strSQL)
                        strSQL = Replace(strSQL, "[ϵͳ]", cboSys.ItemData(cboSys.ListIndex))
                        Call CheckParsRela(strSQL, objReport.Datas, objPar.����, True)
                        strFields = CheckSQL(strSQL, strErr, , strSQLCheck, , , .�������ӱ��)
                        If strFields <> "" Then
                            If strSQLCheck <> "" Then
                                If CheckSQLPlan(strSQLCheck, , .�������ӱ��) = True Then
                                    strMsg = strMsg & "," & .���� & "(" & objPar.���� & ")[��ϸ]"
                                End If
                            End If
                        End If
                    End If
                Next
            End With
        Next
        strMsg = Mid(strMsg, 2)
        lvwReport.ListItems(i).SubItems(RC_������������Դ) = strMsg
        If strMsg <> "" Then lngCount = lngCount + 1
        ShowFlash "���ڼ�鱨������ԴSQL���ڵ���������,���Ժ� ...", i / lvwReport.ListItems.count
    Next
    ShowFlash
    If lngCount > 0 Then
        MsgBox "һ����" & lvwReport.ListItems.count & "�ű�����������ܼ�飬����" & lngCount & "�ű���(������)������Դ���ܴ����������⣬���""������������Դ""�е���Ϣ��" & vbCrLf & vbCrLf & _
            "���ڱ�����ƽ���鿴��ϸ��ִ�мƻ���������SQL�����Ż���", vbInformation, "���ܼ����"
    End If
End Sub

Private Sub txtFind_GotFocus()
    txtFind.SelStart = 0: txtFind.SelLength = Len(txtFind.Text)
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim lngKey As Long, intActive As Integer
    Dim strSQL As String
    If KeyAscii = vbKeyReturn Then
        If txtFind.Text = "" Then Exit Sub
        If mstrFindValue <> txtFind.Text And txtFind.Text <> "" Then
            mstrFindValue = txtFind.Text
            Call LocateItem(mstrFindValue, True)
        Else
            Call LocateItem(mstrFindValue, False)
        End If
    End If
End Sub

Private Sub LocateItem(ByVal strInput As String, Optional ByVal blnClearSel As Boolean)
'���ܣ���λƥ����Ŀ
'������strInput=��������
'           blnClearSel=�����ǰѡ�е���Ŀ
    Dim i As Long, lngStart As Long
    Dim lng���� As Long, lng���� As Long
    Dim lvwTmp As ListView
    Dim strTmp As String
    Dim blnFind As Long
    Dim lngOldSel As Long
    
    strInput = UCase(strInput)
     If mcsActive = CS_���� Then
        lblFind.Caption = "���ұ���(&F)"
        Set lvwTmp = lvwReport
        lng���� = RC_���: lng���� = RC_����
    Else
        lblFind.Caption = "���ҷ���(&F)"
        Set lvwTmp = lvwGroup
        lng���� = GC_���: lng���� = GC_����
    End If
    With lvwTmp
        If Not .SelectedItem Is Nothing And Not blnClearSel Then lngStart = .SelectedItem.Index + 1
        lngOldSel = .SelectedItem.Index
        For i = 1 To .ListItems.count
            .ListItems(i).Selected = False
        Next
        Set .SelectedItem = Nothing
        .SetFocus
        For i = IIF(lngStart = 0, 1, lngStart) To lvwTmp.ListItems.count
            
            strTmp = UCase(lvwTmp.ListItems(i).Text & "|" & .ListItems(i).SubItems(lng����) & "|" & .ListItems(i).SubItems(lng����))
            If strTmp Like "*" & strInput & "*" Then
                Set .SelectedItem = .ListItems(i)
                .ListItems(.SelectedItem.Index).Selected = True
                .SelectedItem.EnsureVisible
                If mcsActive = CS_������ Then
                    mstrPreGroup = ""
                    Call LvwGroup_ItemClick(.SelectedItem)
                End If
                blnFind = True: Exit For
            End If
        Next
        If blnFind Then
            Exit Sub
        ElseIf i >= .ListItems.count Then
            '�ָ�ԭʼ����򱨱�����ʼ��ѡ��
            Set .SelectedItem = .ListItems(lngOldSel)
            .ListItems(.SelectedItem.Index).Selected = True
            .SelectedItem.EnsureVisible
            If mcsActive = CS_������ Then
                mstrPreGroup = ""
                Call LvwGroup_ItemClick(.SelectedItem)
            End If
            If lngStart <> 0 Then
                If MsgBox(" �Ѿ���λ�������ҵ�����Ϣ���Ƿ����²��ң�", vbInformation + vbYesNo, App.Title) = vbYes Then
                    Call LocateItem(strInput, True)
                Else
                    txtFind.SetFocus
                End If
                Exit Sub
            Else
                MsgBox " û���ҵ�������������Ϣ��", vbInformation, App.Title
                txtFind.SetFocus
                Exit Sub
            End If
        End If
    End With
End Sub

Private Sub AfterItemEdit(ByVal blnAdd As Boolean, ByVal blnGroup As Boolean, ByVal lngID As Long, _
                                            ByVal str���� As String, ByVal str���� As String, ByVal str˵�� As String)
'���ܣ��޸ģ���������򱨱������洦��
'����:blnAdd=True-������False-�޸�
'       blnGroup=True-������б༭��False-�Ա�����б༭
'       lngID=��ID�򱨱�ID
'       str����=��򱨱�����
'       str����=��򱨱����
'       str˵��=��򱨱�˵��
    Dim objItem As ListItem
    If blnGroup Then
        If blnAdd Then
            Set objItem = lvwGroup.ListItems.Add(, "_" & lngID, str����, 2, 2)
            objItem.Tag = 0
        Else
            Set objItem = lvwGroup.SelectedItem
            objItem.Text = str����
        End If
        objItem.SubItems(GC_���) = str����
        objItem.SubItems(GC_˵��) = str˵��
        objItem.SubItems(GC_����ʱ��) = Format(Currentdate, "yyyy-MM-dd")
        objItem.Selected = True
        lvwGroup.SelectedItem.EnsureVisible
    Else
        If blnAdd Then
            Set objItem = lvwReport.ListItems.Add(, "_" & lngID, str����, "Report", "Report")
            objItem.Tag = 0
        Else
            Set objItem = lvwReport.SelectedItem
            objItem.Text = str����
        End If
        objItem.SubItems(RC_���) = str����
        objItem.SubItems(RC_˵��) = str˵��
        objItem.SubItems(RC_����ʱ��) = Format(Currentdate, "yyyy-MM-dd")
        objItem.Selected = True
        lvwReport.SelectedItem.EnsureVisible
    End If
End Sub

Private Sub CustomToolBarRefresh()
    Dim i As Integer
    
    For i = 1 To cbr.Bands.count
        cbr.Bands(i).Visible = False
        cbr.Bands(i).Visible = True
    Next
End Sub
