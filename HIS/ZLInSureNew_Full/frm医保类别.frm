VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmҽ����� 
   Caption         =   "�����������"
   ClientHeight    =   6105
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9765
   Icon            =   "frmҽ�����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cmb���� 
      Height          =   300
      Left            =   6960
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   900
      Width           =   1815
   End
   Begin VB.PictureBox picSplitV 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5490
      Left            =   5370
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5490
      ScaleWidth      =   45
      TabIndex        =   1
      Top             =   690
      Width           =   45
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   3660
      Top             =   5340
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ�����.frx":0E42
            Key             =   "Fix"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ�����.frx":115C
            Key             =   "FixD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ�����.frx":1476
            Key             =   "Common"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ�����.frx":1790
            Key             =   "CommonD"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   2850
      Top             =   5310
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
            Picture         =   "frmҽ�����.frx":1AAA
            Key             =   "Fix"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ�����.frx":1DC4
            Key             =   "FixD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ�����.frx":20DE
            Key             =   "Common"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ�����.frx":23F8
            Key             =   "CommonD"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   7770
      Top             =   480
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
            Picture         =   "frmҽ�����.frx":2712
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ�����.frx":292C
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ�����.frx":2B46
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ�����.frx":2D60
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ�����.frx":2F7A
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ�����.frx":3194
            Key             =   "Select"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ�����.frx":388E
            Key             =   "Parameter"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ�����.frx":3F88
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ�����.frx":41A2
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ�����.frx":43BC
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   7110
      Top             =   510
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
            Picture         =   "frmҽ�����.frx":45D6
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ�����.frx":47F0
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ�����.frx":4A0A
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ�����.frx":4C24
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ�����.frx":4E3E
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ�����.frx":5058
            Key             =   "Select"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ�����.frx":5752
            Key             =   "Parameter"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ�����.frx":5E4C
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ�����.frx":6066
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ�����.frx":6280
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplitH 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   5550
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   3000
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2880
      Width           =   3000
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9765
      _ExtentX        =   17224
      _ExtentY        =   1244
      BandCount       =   1
      _CBWidth        =   9765
      _CBHeight       =   705
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   645
      Width1          =   615
      FixedBackground1=   0   'False
      Key1            =   "Comm"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   645
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   9645
         _ExtentX        =   17013
         _ExtentY        =   1138
         ButtonWidth     =   820
         ButtonHeight    =   1138
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   15
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
               Object.ToolTipText     =   "���ӱ������"
               Object.Tag             =   "����"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "Modify"
               Object.ToolTipText     =   "�޸ı������"
               Object.Tag             =   "�޸�"
               ImageKey        =   "Modify"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "Delete"
               Object.ToolTipText     =   "ɾ���������"
               Object.Tag             =   "ɾ��"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split1"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ѡ��"
               Key             =   "Select"
               Object.ToolTipText     =   "��Ϊ��ǰʹ��ҽ��"
               Object.Tag             =   "ѡ��"
               ImageKey        =   "Select"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split4"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Parameter"
               Description     =   "����"
               Object.ToolTipText     =   "��������"
               Object.Tag             =   "����"
               ImageKey        =   "Parameter"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split2"
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�鿴"
               Key             =   "View"
               Description     =   "View"
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
               Key             =   "Split3"
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Object.ToolTipText     =   "��������"
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
   Begin MSComctlLib.ListView lvwKind_S 
      Height          =   4755
      Left            =   30
      TabIndex        =   0
      Top             =   810
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   8387
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   5745
      Width           =   9765
      _ExtentX        =   17224
      _ExtentY        =   635
      SimpleText      =   $"frmҽ�����.frx":649A
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmҽ�����.frx":64E1
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12144
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh�ֶ� 
      Height          =   1800
      Left            =   5220
      TabIndex        =   7
      Top             =   3960
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   3175
      _Version        =   393216
      Rows            =   5
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   -2147483638
      BackColorBkg    =   -2147483643
      GridColor       =   4210752
      GridColorFixed  =   4210752
      GridLinesFixed  =   1
      ScrollBars      =   2
      MergeCells      =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh���� 
      Height          =   1560
      Left            =   5160
      TabIndex        =   6
      Top             =   2040
      Width           =   3270
      _ExtentX        =   5768
      _ExtentY        =   2752
      _Version        =   393216
      BackColor       =   16777215
      Rows            =   6
      FixedCols       =   0
      BackColorFixed  =   13684944
      BackColorBkg    =   -2147483643
      GridColor       =   4210752
      GridColorFixed  =   4210752
      GridLinesFixed  =   1
      ScrollBars      =   2
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      Caption         =   "ҽ������(&N)"
      Height          =   180
      Left            =   5580
      TabIndex        =   9
      Top             =   960
      Width           =   990
   End
   Begin VB.Label lbl���� 
      Alignment       =   2  'Center
      BackColor       =   &H00E6F5FD&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "������в���"
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   5160
      TabIndex        =   5
      Top             =   1755
      Width           =   3360
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
      Begin VB.Menu mnuFileSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "���(&E)"
      Begin VB.Menu mnuEditAdd 
         Caption         =   "����(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "�޸�(&M)"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "ɾ��(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditSplit0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelect 
         Caption         =   "��Ϊ��ǰʹ��ҽ��(&S)"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuEditDeselect 
         Caption         =   "ȡ��ѡ��(&E)"
      End
   End
   Begin VB.Menu mnuCenter 
      Caption         =   "����(&C)"
      Begin VB.Menu mnuCenterAdd 
         Caption         =   "����(&A)"
      End
      Begin VB.Menu mnuCenterModify 
         Caption         =   "�޸�(&M)"
      End
      Begin VB.Menu mnuCenterDelete 
         Caption         =   "ɾ��(&D)"
      End
      Begin VB.Menu mnuCenterSplitPara 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCenterParameter 
         Caption         =   "���в�������(&P)"
      End
      Begin VB.Menu mnuCenterSplitYear 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCenterYear 
         Caption         =   "�����(&I)"
         Index           =   0
      End
      Begin VB.Menu mnuCenterSplitSect 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCenterSect 
         Caption         =   "֧�����õ�(&E)"
      End
      Begin VB.Menu mnuCenterSplitSpec 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCenterHome 
         Caption         =   "�����ͥ����(&H)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCenterSpec 
         Caption         =   "�����������ⲡ(&S)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCenterEspecial 
         Caption         =   "���������ؼ�(&T)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCenterOut 
         Caption         =   "����תԺ(&O)"
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
      Begin VB.Menu mnuViewSplit0 
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
         Index           =   3
      End
      Begin VB.Menu mnuViewSplit1 
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
         Caption         =   "Web�ϵ�����"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&K)..."
         End
      End
      Begin VB.Menu mnuHelpSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frmҽ�����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const mstrLvw As String = "����,2000,0,1;���,800,0,2;ҽԺ����,1440,0,0;˵��,2000,0,0"

Dim msngStartX As Single, msngStartY As Single    '�ƶ�ǰ����λ��
Dim mblnLoad As Boolean  '���ڻ�δ��ʱΪ��
Dim mintColumn As Integer '
Dim mstrKey As String       '��ǰѡ���ListItem��Keyֵ
Dim mbln����� As Boolean   '�Ƿ���б༭����ε�Ȩ��

Private Sub Form_Activate()
    If mblnLoad = True Then
        Call Form_Resize 'Ϊ��ʹCoolBar����Ӧ�߶�
        FillList
    End If
    mblnLoad = False
End Sub

Private Sub Form_Load()
    mblnLoad = True
    
    Call Ȩ�޿���
    '���������ɾ����ListView�������
    lvwKind_S.Tag = "�ɱ仯��"
    '-----------
    RestoreWinState Me, App.ProductName
    '���ListView�Ļ�δ�����ã������һ��ʹ�ã��Ǿ͵���ȱʡ�ĳ�ʼ��
    If lvwKind_S.ColumnHeaders.Count = 0 Then
        zlControl.LvwSelectColumns lvwKind_S, mstrLvw, True
    End If
    '����lvwKind_S��ʾ���ö�Ӧ�˵�
    mnuViewIcon_Click lvwKind_S.View
    
    lvwKind_S.Sorted = True
    lvwKind_S.SortKey = 1
    
    zlControl.CboSetHeight cmb����, 3600
    Call InitTable
End Sub

Private Sub InitTable()
'���ܣ���ʼ�����
    With msh����
        .Rows = 2: .Cols = 2
        .TextMatrix(0, 0) = "������"
        .TextMatrix(0, 1) = "����ֵ"
        .ColWidth(0) = 1900
        .ColWidth(1) = 3200
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        
        .COL = 0
        .Row = 0
        .ColSel = .Cols - 1
        .RowSel = 0
        .FillStyle = flexFillRepeat
        .CellAlignment = 4
        .FillStyle = flexFillSingle
        .AllowBigSelection = False
        .Row = 1
    End With
    
    With msh�ֶ�
        .Cols = 4
        .ColWidth(0) = 1500
        .ColWidth(1) = 1000
        .ColWidth(2) = 1000
        .ColWidth(3) = 2000
        .ColAlignment(0) = 1
        .ColAlignment(1) = 7
        .ColAlignment(2) = 7
        .ColAlignment(3) = 1
    End With
End Sub

Private Sub Form_Resize()
    Dim sngTop As Single, sngBottom As Single
    On Error Resume Next
    
    sngTop = IIf(cbrThis.Visible, cbrThis.Top + cbrThis.Height, 0)
    sngBottom = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
    
    lvwKind_S.Top = sngTop
    lvwKind_S.Height = IIf(sngBottom - lvwKind_S.Top > 0, sngBottom - lvwKind_S.Top, 0)
    lvwKind_S.Left = ScaleLeft
    
    picSplitV.Top = sngTop
    picSplitV.Height = IIf(sngBottom - picSplitV.Top > 0, sngBottom - picSplitV.Top, 0)
    picSplitV.Left = lvwKind_S.Left + lvwKind_S.Width
    
    With cmb����
        '���ÿؼ�����߾�����
        lbl����.Left = picSplitV.Left + picSplitV.Width
        .Left = lbl����.Left + lbl����.Width + 30
        .Width = IIf(ScaleWidth - cmb����.Left > 0, ScaleWidth - cmb����.Left, 0)
    
        lbl����.Left = lbl����.Left
        lbl����.Width = IIf(ScaleWidth - lbl����.Left > 0, ScaleWidth - lbl����.Left, 0)
    End With
    With lbl����
        msh����.Left = .Left
        msh����.Width = .Width
        picSplitH.Left = .Left
        picSplitH.Width = .Width
        msh�ֶ�.Left = .Left
        msh�ֶ�.Width = .Width
    End With
    
    If cmb����.Visible = True Then
        cmb����.Top = sngTop
        lbl����.Top = sngTop + 60
        lbl����.Top = cmb����.Top + cmb����.Height + 120
    Else
        lbl����.Top = sngTop + 90
    End If
    
    msh����.Top = lbl����.Top + lbl����.Height
    picSplitH.Top = msh����.Top + msh����.Height
    msh�ֶ�.Top = picSplitH.Top + picSplitH.Height
    msh�ֶ�.Height = IIf(sngBottom - msh�ֶ�.Top > 0, sngBottom - msh�ֶ�.Top, 0)
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub lvwKind_S_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then '���Ǹղ�����
        lvwKind_S.SortOrder = IIf(lvwKind_S.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwKind_S.SortKey = mintColumn
        lvwKind_S.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwKind_S_DblClick()
    If mnuEditModify.Visible = True And mnuEditModify.Enabled = True Then
        Call mnuEditModify_Click
    End If
End Sub

Private Sub lvwKind_S_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If mstrKey = Item.Key Then Exit Sub
    mstrKey = Item.Key
    
    Dim rsTemp As New ADODB.Recordset
    Dim lngCount As Long, lngIndex As Long
    
    On Error GoTo errHandle
    If mbln����� = True Then
        '�����������������
        mnuCenterYear(0).Visible = False
        mnuCenterSplitYear.Visible = False
        For lngCount = 1 To mnuCenterYear.UBound
            Unload mnuCenterYear(lngCount)
        Next
        
        'Ȼ���ٰ��µ���Ⱥ�����
        gstrSQL = "select * from ������Ⱥ where ����=[1] order by ���"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CInt(Mid(Item.Key, 2)))
        
        Do Until rsTemp.EOF
            lngIndex = rsTemp("���") - 1
            If lngIndex = 0 Then
                mnuCenterYear(0).Visible = True
                mnuCenterSplitYear.Visible = True
            Else
                Load mnuCenterYear(lngIndex)
            End If
            
            mnuCenterYear(lngIndex).Caption = rsTemp("����") & "(&" & rsTemp("���") & ")"
            rsTemp.MoveNext
        Loop
        rsTemp.Close
    End If
        
    cmb����.Clear
    cmb����.Visible = (Item.Tag = "1")
    lbl����.Visible = cmb����.Visible
    Call Form_Resize
    
    If cmb����.Visible = False Then
        '��ҽ��ֻ����һ������
        cmb����.AddItem "1." & Item.Text
        cmb����.ListIndex = 0
        Exit Sub
    End If
    
    On Error GoTo errHandle
    
    gstrSQL = "select ���,����,���� from ��������Ŀ¼ where ����=[1] order by ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CInt(Mid(Item.Key, 2)))
    Do Until rsTemp.EOF
        cmb����.AddItem rsTemp("����") & "." & rsTemp("����")
        cmb����.ItemData(cmb����.NewIndex) = rsTemp("���")
        rsTemp.MoveNext
    Loop
    
    If cmb����.ListCount > 0 Then
        cmb����.ListIndex = 0
    Else
        Call FillItem
    End If
    
    '���������ҽ���������������ҵ�������
    If Mid(Item.Key, 2) = TYPE_������ Then
        mnuCenterSplitSpec.Visible = True
        mnuCenterSpec.Visible = True
        mnuCenterEspecial.Visible = True
        mnuCenterHome.Visible = True
        mnuCenterOut.Visible = True
    Else
        mnuCenterSplitSpec.Visible = False
        mnuCenterSpec.Visible = False
        mnuCenterEspecial.Visible = False
        mnuCenterHome.Visible = False
        mnuCenterOut.Visible = False
    End If
    Exit Sub

errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmb����_Click()
    '����ˢ�²�����ֵ���ʱ�䲢��������Ϊ�˱�֤һ����ListIndex����ˢ��
    '����û���ٱ�����һ�ε�ListIndexֵ
    Call FillItem
End Sub


Private Sub mnuCenterEspecial_Click()
    '
End Sub

Private Sub mnuCenterHome_Click()
    '
End Sub

Private Sub mnuCenterOut_Click()
    frm����ҵ������.Show 1, Me
End Sub

Private Sub mnuCenterSpec_Click()
    '
End Sub

Private Sub mnuEditAdd_Click()
    If frmҽ�����༭.�༭ҽ�����("") = True Then
        lvwKind_S_ItemClick lvwKind_S.SelectedItem
    End If
End Sub

Private Sub mnuEditModify_Click()
    If lvwKind_S.SelectedItem Is Nothing Then Exit Sub
    
    If frmҽ�����༭.�༭ҽ�����(Mid(lvwKind_S.SelectedItem.Key, 2)) = True Then
        lvwKind_S_ItemClick lvwKind_S.SelectedItem
    End If
End Sub

Private Sub mnuEditDelete_Click()
    Dim intIndex As Integer
    
    If lvwKind_S.SelectedItem Is Nothing Then Exit Sub
    If MsgBox("��ȷ��Ҫɾ����" & lvwKind_S.SelectedItem.Text & "��ҽ�������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
        On Error GoTo errHandle
        
        gstrSQL = "zl_�������_delete(" & Mid(lvwKind_S.SelectedItem.Key, 2) & ")"
        
        MousePointer = vbHourglass
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        With lvwKind_S
            mstrKey = ""
            intIndex = .SelectedItem.Index
            .ListItems.Remove .SelectedItem.Key
            If .ListItems.Count > 0 Then
                intIndex = IIf(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
                .ListItems(intIndex).Selected = True
                .ListItems(intIndex).EnsureVisible
                lvwKind_S_ItemClick .SelectedItem
            Else
                cmb����.Clear
                cmb����.Visible = False
                Call Form_Resize
                Call FillItem
            End If
        End With
        MousePointer = vbDefault
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    MousePointer = vbDefault
End Sub

Private Sub mnuEditDeselect_Click()
    Dim lst As ListItem
    Dim strIcon As String
    
    SaveSetting "ZLSOFT", "����ȫ��", "�Ƿ�֧��ҽ��", "No"
    gintInsure = 0
    SaveSetting "ZLSOFT", "����ȫ��", "ҽ�����", 0
    For Each lst In lvwKind_S.ListItems
        strIcon = IIf(Left(lst.Icon, 3) = "Fix", "Fix", "Common")
        
        lst.Icon = strIcon
        lst.SmallIcon = strIcon
    Next
    Call SetMenu
End Sub

Private Sub mnuEditSelect_Click()
    Dim lst As ListItem
    Dim strIcon As String
    
    SaveSetting "ZLSOFT", "����ȫ��", "�Ƿ�֧��ҽ��", "Yes"
    For Each lst In lvwKind_S.ListItems
        If lst Is lvwKind_S.SelectedItem Then
            '��Ϊ��ǰҽ��
            gintInsure = Mid(lst.Key, 2)
            SaveSetting "ZLSOFT", "����ȫ��", "ҽ�����", gintInsure
            strIcon = IIf(Left(lst.Icon, 3) = "Fix", "FixD", "CommonD")
        Else
            strIcon = IIf(Left(lst.Icon, 3) = "Fix", "Fix", "Common")
        End If
        
        lst.Icon = strIcon
        lst.SmallIcon = strIcon
    Next
    Call SetMenu
End Sub

Private Sub mnuCenterAdd_Click()
    If lvwKind_S.SelectedItem Is Nothing Then Exit Sub
    
    Call frmҽ���������.�༭��������(Mid(lvwKind_S.SelectedItem.Key, 2), "")
End Sub

Private Sub mnuCenterModify_Click()
    If lvwKind_S.SelectedItem Is Nothing Then Exit Sub
    
    Call frmҽ���������.�༭��������(Mid(lvwKind_S.SelectedItem.Key, 2), cmb����.ItemData(cmb����.ListIndex))
End Sub

Private Sub mnuCenterDelete_Click()
    If cmb����.ListIndex < 0 Then Exit Sub
    If MsgBox("��ȷ��Ҫɾ����" & cmb����.Text & "��ҽ��������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
        On Error GoTo errHandle
        
        gstrSQL = "zl_��������Ŀ¼_delete(" & Mid(lvwKind_S.SelectedItem.Key, 2) & "," & cmb����.ItemData(cmb����.ListIndex) & ")"
        
        MousePointer = vbHourglass
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        With cmb����
            .RemoveItem .ListIndex
            If .ListCount > 0 Then
                .ListIndex = 0
            Else
                Call FillItem
            End If
        End With
        MousePointer = vbDefault
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    MousePointer = vbDefault
End Sub

Private Sub mnuCenterParameter_Click()
'���ܣ��޸�ҽ���������в���
'ע�⣺��ͬҽ��������������ɲ�ͬ�ĳ���ʵ�ֵ�
    Dim blnReturn As Boolean
    Dim lng���� As Long
    
    If lvwKind_S.SelectedItem Is Nothing Then Exit Sub
    
    On Error GoTo errHandle
    lng���� = Val(Mid(lvwKind_S.SelectedItem.Key, 2))
    Select Case lng����
        Case TYPE_�Ͼ���
            blnReturn = frmSet�Ͼ���.��������(lng����)
        Case TYPE_��ͨ
            blnReturn = frmSet��ͨ.��������
        Case TYPE_������
            blnReturn = frmSet������.��������()
        Case TYPE_����ũ��
            blnReturn = frmset����ũ��.��������()
        Case TYPE_����
            blnReturn = frmset����.��������()
        Case TYPE_�ɶ���ũҽ
            blnReturn = frmSet�ɶ���ũҽ.��������
        Case TYPE_��Ҧ
            blnReturn = ҽ������_��Ҧ()
        Case TYPE_��������
            blnReturn = ҽ������_��������()
        Case TYPE_�㽭
            blnReturn = frmSet�㽭.��������()
        Case TYPE_�¶�
            blnReturn = ҽ������_�¶�()
        Case TYPE_������
            '���ù�����ȫ�ɴ������
            blnReturn = frmSet����.��������()
        Case TYPE_����
            blnReturn = ҽ������_����()
        Case TYPE_��Ԫ
            blnReturn = ҽ������_��Ԫ()
        Case TYPE_����
            blnReturn = ҽ������_����()
        Case TYPE_�����ɽ
            blnReturn = frmset��ɽ.��������()
        Case TYPE_��������
            blnReturn = frmSet����.��������(Mid(lvwKind_S.SelectedItem.Key, 2), cmb����.ItemData(cmb����.ListIndex))
        Case TYPE_��������ɽ
            blnReturn = frmSet����ɽ.��������()
        Case TYPE_����ʡ, TYPE_������, TYPE_���Ͻ�ˮ
            Dim msgReturn As VbMsgBoxResult
            
            msgReturn = MsgBox("���ʱ�ҽ���Ƿ�֧�ֻ����Բ������ֲ���ҽ�����ˣ�", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName)
            
            gstrSQL = "zl_���ղ���_Delete(" & lng���� & ",0)"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            
            '������������
            gstrSQL = "zl_���ղ���_Insert(" & lng���� & ",0,'֧�����Բ������ֲ�','" & IIf(msgReturn = vbYes, "1", "0") & "',1)"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            
            blnReturn = True
        Case TYPE_������
            blnReturn = frmSet����.��������(Mid(lvwKind_S.SelectedItem.Key, 2))
        Case TYPE_�Թ���
            '���ù�����ȫ�ɴ������
            blnReturn = frmSet����.��������(Mid(lvwKind_S.SelectedItem.Key, 2), cmb����.ItemData(cmb����.ListIndex))
        Case TYPE_�Ĵ��Թ�
            blnReturn = frmSet�Թ�.��������(Mid(lvwKind_S.SelectedItem.Key, 2), cmb����.ItemData(cmb����.ListIndex))
        Case TYPE_������
            '���ù�����ȫ�ɴ������
            blnReturn = frmSet����.��������(Mid(lvwKind_S.SelectedItem.Key, 2), cmb����.ItemData(cmb����.ListIndex))
        Case TYPE_ͭ��
            '���ù�����ȫ�ɴ������
            blnReturn = frmSetͭ��.��������(Mid(lvwKind_S.SelectedItem.Key, 2), cmb����.ItemData(cmb����.ListIndex))
        Case TYPE_������
            '���ù�����ȫ�ɴ������
            blnReturn = frmSet����.��������()
        Case TYPE_�ɶ���
            blnReturn = ҽ������_�ɶ�
        Case TYPE_�ɶ�����
            blnReturn = ҽ������_����
        Case TYPE_����
            blnReturn = ҽ������_����
        Case type_�ɶ�����
            blnReturn = ҽ������_�ɶ�����
        Case TYPE_�ɶ��ϳ�
            blnReturn = ҽ������_�ɶ��ϳ�
        Case TYPE_��������, TYPE_����ʡ, TYPE_������, TYPE_��ƽ��
            blnReturn = ҽ������_��������(lng����)
        Case type_����
            blnReturn = ҽ������_����
        Case TYPE_�Ĵ�üɽ
            blnReturn = ҽ������_üɽ
        Case TYPE_������
            blnReturn = ҽ������_����
        Case TYPE_��ɽ
            blnReturn = ҽ������_��ɽ
        Case TYPE_������, TYPE_����������
            '200311
            If cmb����.ListIndex < 0 Then Exit Sub
            blnReturn = ҽ������_����(Val(Mid(lvwKind_S.SelectedItem.Key, 2)), cmb����.ItemData(cmb����.ListIndex))
        Case TYPE_�ش�У԰��
            '���˺�(200403)
            blnReturn = ҽ������_�ش�У԰��(Val(Mid(lvwKind_S.SelectedItem.Key, 2)), 0)
        Case TYPE_����������
            blnReturn = ҽ������_����������()
        Case TYPE_�����山
            '20040715
            blnReturn = ҽ������_�����山()
        Case TYPE_ǭ��
            '200410
            blnReturn = ҽ������_ǭ��()
        Case TYPE_�ɶ�����
            '200411
            blnReturn = ҽ������_�ɶ�����()
        Case TYPE_�ɶ��ڽ�
            '200411
            blnReturn = ҽ������_�ɶ��ڽ�()
        Case TYPE_�˰�
            '20050125
            blnReturn = ҽ������_�˰�()
        Case TYPE_����
            blnReturn = ҽ������_����()
        Case TYPE_�ٲ׷���
            blnReturn = ҽ������_����()
        Case TYPE_����
            blnReturn = ҽ������_����
        Case TYPE_�Ͻ�
            blnReturn = ҽ������_�Ͻ�
        Case TYPE_����
            blnReturn = ҽ������_����
        Case TYPE_��Ϫũҽ
            blnReturn = ҽ������_��Ϫũҽ
        Case TYPE_��Ԫ����
            blnReturn = ҽ������_��Ԫ����(Mid(lvwKind_S.SelectedItem.Key, 2), cmb����.ItemData(cmb����.ListIndex))
        Case TYPE_�ϳ�����
            blnReturn = ҽ������_�ϳ�����(Mid(lvwKind_S.SelectedItem.Key, 2), cmb����.ItemData(cmb����.ListIndex))
        Case TYPE_�山ũҽ
            blnReturn = ҽ������_�山ũҽ
        Case TYPE_�˳ɺ˹�ҵ
            blnReturn = ҽ������_�˳�(Mid(lvwKind_S.SelectedItem.Key, 2), cmb����.ItemData(cmb����.ListIndex))
        Case TYPE_��������
            blnReturn = ҽ������_��ľ����(Mid(lvwKind_S.SelectedItem.Key, 2), cmb����.ItemData(cmb����.ListIndex))
        Case TYPE_ɽ��
            '�¶���20050304
            blnReturn = frmSetɽ��.��������()
        Case TYPE_ͭɽ��
            blnReturn = frmSetͭɽ��.��������()
        Case Is > 900, TYPE_������Ժ            '����ҽ��
            blnReturn = frmSet����.��������(Mid(lvwKind_S.SelectedItem.Key, 2), cmb����.ItemData(cmb����.ListIndex))
    End Select
    
    If blnReturn = True Then
        '���óɹ���ˢ����ʾ
        Call Fill����
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub mnuCenterSect_Click()
    Dim blnReturn As Boolean
    
    If lvwKind_S.SelectedItem Is Nothing Then Exit Sub
    
    blnReturn = frmҽ����𵵴�.��������(Mid(lvwKind_S.SelectedItem.Key, 2), cmb����.ItemData(cmb����.ListIndex))
    If blnReturn = True Then
        '���óɹ���ˢ����ʾ
        Call Fill�ֶ�
    End If
End Sub

Private Sub mnuCenterYear_Click(Index As Integer)
    Dim blnReturn As Boolean
    Dim STRNAME As String
    
    If lvwKind_S.SelectedItem Is Nothing Then Exit Sub
    
    STRNAME = Left(mnuCenterYear(Index).Caption, InStr(mnuCenterYear(Index).Caption, "(") - 1)
    blnReturn = frmҽ�������.��������(Mid(lvwKind_S.SelectedItem.Key, 2), cmb����.ItemData(cmb����.ListIndex), Index + 1, STRNAME)
    If blnReturn = True Then
        '���óɹ���ˢ����ʾ
        Call Fill�ֶ�
    End If
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFilePreview_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub mnuFilePrintSet_Click()
    zlPrintSet
End Sub

Private Sub subPrint(ByVal bytMode As Byte)
    '����:���д�ӡ,Ԥ���������EXCEL
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    
    Dim objPrint As New zlPrintGrds
    Dim objRow As New zlTabAppRow
    
    Set objPrint.Grds = New Collection
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    objPrint.Title.Text = "�����������"
        
    objRow.Add "�������" & lvwKind_S.SelectedItem.Text
    If cmb����.Visible = True Then
        objRow.Add "ҽ�����ģ�" & cmb����.Text
    End If
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add " "
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & UserInfo.����    '& "   ��ӡ����:" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ����:" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    
    objPrint.Grds.Add msh����
    objPrint.Grds.Add msh�ֶ�
    
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrViewGrds objPrint, 1
          Case 2
              zlPrintOrViewGrds objPrint, 2
          Case 3
              zlPrintOrViewGrds objPrint, 3
      End Select
    Else
        zlPrintOrViewGrds objPrint, bytMode
    End If
End Sub


Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(hwnd)
End Sub


Private Sub mnuViewRefresh_Click()
    Call FillList
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim intCOUNT As Integer
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For intCOUNT = 1 To tbrThis.Buttons.Count
        tbrThis.Buttons(intCOUNT).Caption = IIf(mnuViewToolText.Checked, tbrThis.Buttons(intCOUNT).Tag, "")
    Next
    
    cbrThis.Bands(1).MinHeight = tbrThis.Height
    Call Form_Resize
End Sub

Private Sub mnuViewIcon_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
    Next
    mnuViewIcon(Index).Checked = True
    lvwKind_S.View = Index
End Sub

Private Sub msh����_DblClick()
    If mnuCenterParameter.Visible = True And mnuCenterParameter.Enabled = True Then
        Call mnuCenterParameter_Click
    End If
End Sub

Private Sub msh�ֶ�_DblClick()
    If msh�ֶ�.RowData(msh�ֶ�.Row) <> 0 Then
        '��������γ���
        If mbln����� = True And mnuCenterYear(0).Enabled = True Then
            Call mnuCenterYear_Click(msh�ֶ�.RowData(msh�ֶ�.Row) - 1)
        End If
    Else
        '���÷��õ��γ���
        If mnuCenterSect.Visible And mnuCenterSect.Enabled Then
            Call mnuCenterSect_Click
        End If
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
        If sngTemp > 2000 And Me.ScaleWidth - (sngTemp + picSplitV.Width) > 1000 Then
            picSplitV.Left = sngTemp
            lvwKind_S.Width = picSplitV.Left - lvwKind_S.Left
            
            Call Form_Resize
        End If
        lvwKind_S.SetFocus
    End If
End Sub

Private Sub picSplitH_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        msngStartY = y
    End If
End Sub

Private Sub picSplitH_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sngTemp As Single
    
    If Button = 1 Then
        sngTemp = picSplitH.Top + y - msngStartY
        If sngTemp - msh����.Top > 500 And (msh�ֶ�.Top + msh�ֶ�.Height) - (sngTemp + picSplitV.Width) > 1000 Then
            picSplitH.Top = sngTemp
            msh����.Height = picSplitH.Top - msh����.Top
            
            Call Form_Resize
        End If
        msh����.SetFocus
    End If
End Sub

Private Sub tbrThis_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim i As Integer
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
    Next
    mnuViewIcon(ButtonMenu.Index - 1).Checked = True
    lvwKind_S.View = ButtonMenu.Index - 1
End Sub

Private Sub tbrThis_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Preview"
            mnuFilePreview_Click
        Case "Print"
            mnuFilePrint_Click
        Case "New"
            mnuEditAdd_Click
        Case "Modify"
            mnuEditModify_Click
        Case "Delete"
            mnuEditDelete_Click
        Case "Select"
            mnuEditSelect_Click
        Case "Parameter"
            mnuCenterParameter_Click
        Case "Help"
            mnuHelpTitle_Click
        Case "Quit"
            mnuFileExit_Click
        Case "View"
            mnuViewIcon(lvwKind_S.View).Checked = False
            If lvwKind_S.View = 3 Then
                mnuViewIcon(0).Checked = True
                lvwKind_S.View = 0
            Else
                mnuViewIcon(lvwKind_S.View + 1).Checked = True
                lvwKind_S.View = lvwKind_S.View + 1
            End If
    End Select
End Sub

Private Sub FillList()
'���ܣ���ʾ����ҽ������б�
    Dim rsTemp As New ADODB.Recordset
    Dim strIcon As String
    Dim lst As ListItem, strKey As String
    Dim lngCol  As Long, varValue As Variant

    If Not lvwKind_S.SelectedItem Is Nothing Then
        strKey = lvwKind_S.SelectedItem.Key
    End If
    
    lvwKind_S.ListItems.Clear
    mstrKey = ""
    
    gstrSQL = "select ���,����,˵��,ҽԺ����,�Ƿ�̶�,��������,�Ƿ��ֹ from ������� Where ҽ������ Is NULL"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    Do Until rsTemp.EOF
        strIcon = IIf(rsTemp("�Ƿ�̶�") = 1, "Fix", "Common")
        If rsTemp("���") = gintInsure Then strIcon = strIcon & "D"
        
        Set lst = lvwKind_S.ListItems.Add(, "K" & rsTemp("���"), rsTemp("����"), strIcon, strIcon)
        If rsTemp("���") = gintInsure Then
            lst.Selected = True
        End If
        '����ListView�����������ݿ�ȡ��
        For lngCol = 2 To lvwKind_S.ColumnHeaders.Count
            varValue = rsTemp(lvwKind_S.ColumnHeaders(lngCol).Text).Value
            lst.SubItems(lngCol - 1) = IIf(IsNull(varValue), "", varValue)
        Next
        lst.Tag = IIf(rsTemp("��������") = 1, 1, 0)
        If rsTemp("�Ƿ��ֹ") = 1 Then
            lst.Ghosted = True
        End If
        rsTemp.MoveNext
    Loop
    
    If lvwKind_S.ListItems.Count > 0 Then
        On Error Resume Next
        Set lst = lvwKind_S.ListItems(strKey)
        If Err <> 0 Then
            Err.Clear
            If lvwKind_S.SelectedItem Is Nothing Then
                Set lst = lvwKind_S.ListItems(1)
                lst.Selected = True
            Else
                Set lst = lvwKind_S.SelectedItem
            End If
        Else
            lst.Selected = True
        End If
        lst.EnsureVisible
        lvwKind_S_ItemClick lst
    Else
        cmb����.Clear
        cmb����.Visible = False
        Call Form_Resize
        Call FillItem
    End If
End Sub

Private Sub FillItem()
'���ܣ�����ҽ�����ĵ������ʾ�������ֶ�����
    
    Call Fill����
    Call Fill�ֶ�
    Call SetMenu
End Sub

Private Sub Fill����()
'���ܣ�����ҽ�����ĵ������ʾ����
    Dim rsTemp As New ADODB.Recordset
    Dim lst As ListItem, lngRow As Long
    Dim strTemp As String
    
    With msh����
        .Rows = 2
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 1) = ""
    End With
    '���û�����ģ��϶���������ʾ��
    If lvwKind_S.SelectedItem Is Nothing Or cmb����.ListIndex < 0 Then Exit Sub
    
    With msh����
        Select Case Val(Mid(lvwKind_S.SelectedItem.Key, 2))
            Case TYPE_�ɶ���
                .Rows = 3
                .TextMatrix(1, 0) = "���Ӵ�"
                .TextMatrix(1, 1) = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("ConnectionStrINg"), "dsn=cnnSyb;uID=face;pwd=facepass")
                .TextMatrix(2, 0) = "���ų���"
                .TextMatrix(2, 1) = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("CardNOLength"), 20)
            Case TYPE_�ɶ�����
                .Rows = 3
                .TextMatrix(1, 0) = "���Ӵ�"
                .TextMatrix(1, 1) = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("LHConnectionStrINg"), "dsn=lhyb;uid=sa;pwd=;")
                .TextMatrix(2, 0) = "ҽ������"
                .TextMatrix(2, 1) = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("intercode"), 713)
            Case TYPE_��������, TYPE_����ʡ, TYPE_������, TYPE_��ƽ��
                gstrSQL = "select ������,����ֵ from ���ղ��� where ����=[1] and (����=[2] or ���� is null) order by ���"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CInt(Mid(lvwKind_S.SelectedItem.Key, 2)), CInt(cmb����.ItemData(cmb����.ListIndex)))
                
                If rsTemp.RecordCount = 0 Then Exit Sub
                .Rows = rsTemp.RecordCount + 1
                lngRow = 1
                Do Until rsTemp.EOF
                    .TextMatrix(lngRow, 0) = rsTemp("������")
                    .TextMatrix(lngRow, 1) = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
                    
                    lngRow = lngRow + 1
                    rsTemp.MoveNext
                Loop
            Case Else
                '�����Թ�ҽ��������ҽ��
                '�������ΪNull����ʾ�ò���������������Ч
                
                '�̶����������޸���鿴��ֻ����
                gstrSQL = "select ������,����ֵ from ���ղ��� where ����=[1] and (����=[2] or ���� is null) and ������ not like '%����%' and ������ not in('����֤��') and (�Ƿ�̶�<>1 Or �Ƿ�̶� Is null ) order by ���"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CInt(Mid(lvwKind_S.SelectedItem.Key, 2)), CInt(cmb����.ItemData(cmb����.ListIndex)))
                
                If rsTemp.RecordCount = 0 Then Exit Sub
                
                .Rows = rsTemp.RecordCount + 1
                lngRow = 1
                Do Until rsTemp.EOF
                    .TextMatrix(lngRow, 0) = rsTemp("������")
                    Select Case rsTemp("������")
                        Case "������"  '����ҽ��
                            .TextMatrix(lngRow, 1) = IIf(rsTemp("����ֵ") = "1", "������", "��ȡ��")
                        Case "������֤"
                            .TextMatrix(lngRow, 1) = IIf(rsTemp("����ֵ") = "1", "��Ҫ", "����Ҫ")
                        Case "�շ�ʹ��ҽ������"
                            .TextMatrix(lngRow, 1) = IIf(rsTemp("����ֵ") = "1", "����", "������")
                        Case "֧�����Բ������ֲ�", "����ҽ�ƻ���", "��������", "���������շ�", "֧����������", "��Ժʱѡ��α�ǰ��Ժ"
                            .TextMatrix(lngRow, 1) = IIf(rsTemp("����ֵ") = "1", "��", "��")
                        Case "�շѸ����ʻ�ʹ�÷�Χ", "��������ʻ�ʹ�÷�Χ"
                            strTemp = IIf(IsNull(rsTemp("����ֵ")), "00", rsTemp("����ֵ"))
                            
                            .TextMatrix(lngRow, 1) = IIf(Left(strTemp, 1) = "1", "ȫ�ԷѲ��֡�", "") & _
                                                     IIf(Mid(strTemp, 2, 1) = "1", "�����Ը����֡�", "") & _
                                                     IIf(Mid(strTemp, 3, 1) = "1", "���޲��֡�", "")
                            If .TextMatrix(lngRow, 1) <> "" Then
                                .TextMatrix(lngRow, 1) = Mid(.TextMatrix(lngRow, 1), 1, Len(.TextMatrix(lngRow, 1)) - 1)
                            End If
                        Case "�ȿ�����"
                            strTemp = IIf(IsNull(rsTemp("����ֵ")), "0", rsTemp("����ֵ"))
                            .TextMatrix(lngRow, 1) = IIf(strTemp = "1", "��", "��")
                        Case Else
                            If rsTemp!������ Like "*����*" Then
                                .TextMatrix(lngRow, 1) = "********"
                            Else
                                .TextMatrix(lngRow, 1) = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
                            End If
                    End Select
                    
                    lngRow = lngRow + 1
                    rsTemp.MoveNext
                Loop
                
                Select Case Val(Mid(lvwKind_S.SelectedItem.Key, 2))
                    Case TYPE_�����ɽ
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, 0) = "��ǰʹ�õĴ���"
                        If IsNumeric(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "��ǰʹ�õĴ���", "")) = True Then
                            .TextMatrix(.Rows - 1, 1) = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "��ǰʹ�õĴ���", "") + 1
                        End If
                End Select
        End Select
    End With
End Sub

Private Sub Fill�ֶ�()
'���ܣ�����ҽ�����ĵ������ʾ����ֶ�����õ���
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long, lng���� As Long, lng���� As Long
    Dim str��ע As String
    
    '���û�����ģ��϶���������ʾ��
    rsTemp.CursorLocation = adUseClient
    With msh�ֶ�
        .Clear
        .BackColor = &HFFFFFF
        If lvwKind_S.SelectedItem Is Nothing Or cmb����.ListIndex < 0 Then
            'û�����ݣ���ʾһ�ſձ�
            .Rows = 9
            Set��ͷ 0, "������Ⱥ�����", True, 1
            Set��ͷ 1, "����,����,����,��ע", False, 1
            Set��ͷ 3, "֧�����õ�", True, 0
            Set��ͷ 4, "����,����,����,��ע", False, 0
            Exit Sub
        End If
        
        lng���� = Mid(lvwKind_S.SelectedItem.Key, 2)
        lng���� = cmb����.ItemData(cmb����.ListIndex)
        
        gstrSQL = "select A.��ְ,A.����,A.����,A.����,B.���� as ��Ⱥ����,B.��� " & _
                "   ,nvl(ȫ��ͳ��,0) as ȫ��ͳ��,nvl(������,0) as ������,nvl(�޷ⶥ��,0) as �޷ⶥ��" & _
                " from ��������� A ,������Ⱥ B" & _
                " where A.����(+)=B.���� and A.��ְ(+)=B.��� and B.����=[1] and A.����(+)=[2]" & _
                " Order by A.��ְ,A.�����"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����, lng����)
        
        If rsTemp.RecordCount = 0 Then
            .Rows = 3 '����������
        Else
            .Rows = rsTemp.RecordCount + 2
        End If
        Set��ͷ 0, "������Ⱥ�����", True, 1
        Set��ͷ 1, "����,����,����,��ע", False, 1
        lngRow = 2
        Do Until rsTemp.EOF
            .MergeRow(lngRow) = False
            .TextMatrix(lngRow, 0) = IIf(IsNull(rsTemp("����")), rsTemp("��Ⱥ����"), rsTemp("����"))
            .TextMatrix(lngRow, 1) = Format(rsTemp("����"), "###;-###; ; ")
            .TextMatrix(lngRow, 2) = Format(rsTemp("����"), "###;-###; ; ")
            
            If IsNull(rsTemp("����")) = True Then
                str��ע = "��δ����"
            Else
                str��ע = ""
                 If rsTemp("ȫ��ͳ��") = 1 Then str��ע = ",ȫ��ͳ��"
                 If rsTemp("������") = 1 Then str��ע = str��ע & ",������"
                 If rsTemp("�޷ⶥ��") = 1 Then str��ע = str��ע & ",�޷ⶥ��"
                 
                 str��ע = Mid(str��ע, 2)
            End If
            .TextMatrix(lngRow, 3) = str��ע
            
            .RowData(lngRow) = rsTemp("���")
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        
        gstrSQL = "select ����,����,���� from ���շ��õ� " & _
            "where ����=[1] and ����=[2]"
        If lng���� = TYPE_�Ĵ�üɽ Then gstrSQL = gstrSQL & " And ����<>0 "
        gstrSQL = gstrSQL & " Order by ����"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����, lng����)
        lngRow = .Rows
        If rsTemp.RecordCount = 0 Then
            .Rows = .Rows + 3 '����������
        Else
            .Rows = .Rows + rsTemp.RecordCount + 2
        End If
        Set��ͷ lngRow, "����֧�����õ�", True, 0
        Set��ͷ lngRow + 1, "����,����,����,��ע", False, 0
        lngRow = lngRow + 2
        Do Until rsTemp.EOF
            .MergeRow(lngRow) = False
            .TextMatrix(lngRow, 0) = rsTemp("����")
            .TextMatrix(lngRow, 1) = Format(rsTemp("����"), "########0.00;-########0.00; ; ")
            .TextMatrix(lngRow, 2) = Format(rsTemp("����"), "########0.00;-########0.00; ; ")
            
            .RowData(lngRow) = 0
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
End Sub

Private Sub SetMenu()
'���ܣ����ݵ�ǰ����ʾ�������ò˵�������
    Dim blnItem As Boolean
    Dim lngIndex As Long
    Dim blnEnable As Boolean
    
    If lvwKind_S.SelectedItem Is Nothing Then
        '��ǰû�п����õ�
        stbThis.Panels(2).Text = "����ҽ�����" & lvwKind_S.ListItems.Count & "��"
        mnuEditModify.Enabled = False
        mnuEditDelete.Enabled = False
        mnuEditSelect.Enabled = False
        mnuEditDeselect.Enabled = False
        
        mnuCenterAdd.Enabled = False
        mnuCenterModify.Enabled = False
        mnuCenterDelete.Enabled = False
        mnuCenterParameter.Enabled = False
        mnuCenterSect.Enabled = False
    Else
        blnEnable = Val(Mid(lvwKind_S.SelectedItem.Key, 2)) <> TYPE_�Թ��� And Val(Mid(lvwKind_S.SelectedItem.Key, 2)) <> TYPE_ͭ�� And Val(Mid(lvwKind_S.SelectedItem.Key, 2)) <> TYPE_��������ɽ
        
        stbThis.Panels(2).Text = "����ҽ�����" & lvwKind_S.ListItems.Count & "������ѡ��Ϊ" & lvwKind_S.SelectedItem.Text
        mnuEditModify.Enabled = True
        mnuEditDelete.Enabled = (Left(lvwKind_S.SelectedItem.Icon, 6) = "Common")
        mnuEditDeselect.Enabled = (Right(lvwKind_S.SelectedItem.Icon, 1) = "D")
        mnuEditSelect.Enabled = Not mnuEditDeselect.Enabled
        
        mnuCenterAdd.Enabled = cmb����.Visible
        mnuCenterModify.Enabled = cmb����.Visible And cmb����.ListIndex > -1
        mnuCenterDelete.Enabled = mnuCenterModify.Enabled
        
        mnuCenterParameter.Enabled = cmb����.ListIndex > -1
        mnuCenterSect.Enabled = mnuCenterParameter.Enabled And blnEnable
    End If
    
    For lngIndex = mnuCenterYear.LBound To mnuCenterYear.UBound
        mnuCenterYear(lngIndex).Enabled = mnuCenterSect.Enabled
    Next
    
    tbrThis.Buttons("Modify").Enabled = mnuEditModify.Enabled
    tbrThis.Buttons("Delete").Enabled = mnuEditDelete.Enabled
    tbrThis.Buttons("Select").Enabled = mnuEditSelect.Enabled
    tbrThis.Buttons("Parameter").Enabled = mnuCenterParameter.Enabled
End Sub

Private Sub Ȩ�޿���()
    If InStr(gstrPrivs, "��ɾ��") = 0 Then
        tbrThis.Buttons("New").Visible = False
        tbrThis.Buttons("Modify").Visible = False
        tbrThis.Buttons("Delete").Visible = False
        tbrThis.Buttons("Split1").Visible = False
        
        mnuEditAdd.Visible = False
        mnuEditModify.Visible = False
        mnuEditDelete.Visible = False
        mnuEditSplit0.Visible = False
        
        mnuCenterAdd.Visible = False
        mnuCenterModify.Visible = False
        mnuCenterDelete.Visible = False
        mnuCenterSplitPara.Visible = False
    End If
    
    If InStr(gstrPrivs, "�����") = 0 Then
        mbln����� = False
        mnuCenterYear(0).Visible = False
        mnuCenterSplitYear.Visible = False
    Else
        mbln����� = True
    End If
    
    If InStr(gstrPrivs, "���շ��õ�") = 0 Then
        mnuCenterSect.Visible = False
        mnuCenterSplitSect.Visible = False
    End If
    
    If InStr(gstrPrivs, "���в�������") = 0 Then
        tbrThis.Buttons("Parameter").Visible = False
        tbrThis.Buttons("Split2").Visible = False
        
        If gstrPrivs = "����" Then
            '��ȫ����
            mnuCenter.Visible = False
            mnuCenterSplitPara.Visible = True
            mnuCenterParameter.Visible = False
        Else
            'ֻ���������Ŀ
            mnuCenterParameter.Visible = False
            mnuCenterSplitPara.Visible = False
        End If
    End If
End Sub


Private Sub Set��ͷ(ByVal lngRow As Long, strCaptions As String, ByVal blnMerge As Boolean, ByVal lngIndex As Long)
'���ܣ�Ϊ���õ��ļ����ӱ����ñ�ͷ
    Dim lngCol As Long
    Dim varCaptions As Variant
    
    With msh�ֶ�
        .MergeRow(lngRow) = blnMerge
        .RowData(lngRow) = lngIndex
        
        varCaptions = Split(strCaptions, ",")
        For lngCol = 0 To .Cols - 1
            If blnMerge = False Then
                .TextMatrix(lngRow, lngCol) = varCaptions(lngCol)
            Else
                .TextMatrix(lngRow, lngCol) = strCaptions
            End If
        Next
        
        .AllowBigSelection = True
        .Row = lngRow
        .COL = 0
        .RowSel = lngRow
        .ColSel = .Cols - 1
        .FillStyle = flexFillRepeat
        .CellAlignment = 4
        .CellBackColor = IIf(blnMerge, &HE6F5FD, &HD0D0D0)
        .CellForeColor = IIf(blnMerge, &H800000, 0)
        .FillStyle = flexFillSingle
        .AllowBigSelection = False
    End With
End Sub


