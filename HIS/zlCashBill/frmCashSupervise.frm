VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCashSupervise 
   Caption         =   "�շѲ�����"
   ClientHeight    =   6795
   ClientLeft      =   -135
   ClientTop       =   240
   ClientWidth     =   10800
   Icon            =   "frmCashSupervise.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   10800
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picGroup 
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   30
      ScaleHeight     =   420
      ScaleWidth      =   3015
      TabIndex        =   12
      Top             =   780
      Width           =   3015
      Begin VB.ComboBox cbo��Ա�� 
         Height          =   300
         Left            =   810
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   105
         Width           =   2100
      End
      Begin VB.Label lblGroup 
         AutoSize        =   -1  'True
         Caption         =   "��Ա����"
         Height          =   180
         Left            =   60
         TabIndex        =   14
         Top             =   165
         Width           =   720
      End
   End
   Begin MSComctlLib.ImageList ilssmall 
      Left            =   4350
      Top             =   4140
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
            Picture         =   "frmCashSupervise.frx":0442
            Key             =   "man"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   3150
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5175
      ScaleMode       =   0  'User
      ScaleWidth      =   38.572
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   840
      Width           =   45
   End
   Begin MSComctlLib.ImageList Ilsrw 
      Left            =   5760
      Top             =   120
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
            Picture         =   "frmCashSupervise.frx":0766
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashSupervise.frx":0980
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashSupervise.frx":0BA0
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashSupervise.frx":0DC0
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashSupervise.frx":0FE0
            Key             =   "View"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashSupervise.frx":1200
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashSupervise.frx":1420
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashSupervise.frx":1640
            Key             =   "Filter"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Ilscolor 
      Left            =   5160
      Top             =   120
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
            Picture         =   "frmCashSupervise.frx":185A
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashSupervise.frx":1A7A
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashSupervise.frx":1C9A
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashSupervise.frx":1EBA
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashSupervise.frx":20DA
            Key             =   "View"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashSupervise.frx":22FA
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashSupervise.frx":251A
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashSupervise.frx":273A
            Key             =   "Filter"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   10800
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "Toolbar1"
      MinHeight1      =   720
      Width1          =   8070
      Key1            =   "only"
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Caption2        =   "��Ա����"
      Child2          =   "cboKind"
      MinWidth2       =   1110
      MinHeight2      =   300
      Width2          =   2010
      NewRow2         =   0   'False
      BandStyle2      =   1
      AllowVertical2  =   0   'False
      Begin VB.ComboBox cboKind 
         Height          =   300
         Left            =   9600
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   240
         Width           =   1170
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   720
         Left            =   165
         TabIndex        =   1
         Top             =   30
         Width           =   8505
         _ExtentX        =   15002
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "Ilsrw"
         HotImageList    =   "Ilscolor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   11
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
               Key             =   "Split1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "New"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "New"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "PayFree"
                     Text            =   "�ֹ��ɿ�(&A)"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "PayAll"
                     Text            =   "ȫ��ɿ�(&B)"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "PayDay"
                     Text            =   "���սɿ�(&C)"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "Delete"
               Object.ToolTipText     =   "ɾ��"
               Object.Tag             =   "ɾ��"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split2"
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Filter"
               Object.ToolTipText     =   "������������"
               Object.Tag             =   "����"
               ImageKey        =   "Filter"
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
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Quit"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ilsbig 
      Left            =   4560
      Top             =   2790
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
            Picture         =   "frmCashSupervise.frx":2954
            Key             =   "man"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picContainer 
      BackColor       =   &H00FFFFFF&
      Height          =   5550
      Left            =   3210
      ScaleHeight     =   5490
      ScaleWidth      =   6195
      TabIndex        =   3
      Top             =   810
      Width           =   6255
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshRecord 
         Height          =   2370
         Left            =   120
         TabIndex        =   6
         Top             =   2730
         Width           =   5940
         _ExtentX        =   10478
         _ExtentY        =   4180
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         BackColorFixed  =   -2147483648
         BackColorBkg    =   16777215
         GridColor       =   8421504
         GridColorFixed  =   8421504
         GridLinesFixed  =   1
         GridLinesUnpopulated=   1
         SelectionMode   =   1
         MergeCells      =   2
         AllowUserResizing=   1
         Appearance      =   0
         MouseIcon       =   "frmCashSupervise.frx":2C78
         _NumberOfBands  =   1
         _Band(0).Cols   =   8
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshTotal 
         Height          =   1455
         Left            =   180
         TabIndex        =   7
         Top             =   690
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   2566
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorFixed  =   -2147483648
         BackColorBkg    =   -2147483643
         BackColorUnpopulated=   -2147483644
         GridColor       =   8421504
         GridColorFixed  =   8421504
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         GridLinesFixed  =   1
         AllowUserResizing=   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
      End
      Begin VB.Label lblSplit 
         BackStyle       =   0  'Transparent
         Height          =   60
         Left            =   600
         MousePointer    =   7  'Size N S
         TabIndex        =   9
         Top             =   2370
         Width           =   1065
      End
      Begin VB.Label lblCaption1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ݴ������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   1920
         TabIndex        =   8
         Top             =   150
         Width           =   2175
      End
      Begin VB.Label lblCaption2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ɿ��¼"
         BeginProperty Font 
            Name            =   "����"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   2310
         TabIndex        =   4
         Top             =   2250
         Width           =   1455
      End
   End
   Begin MSComctlLib.ListView lvwMain_S 
      Height          =   5040
      Left            =   45
      TabIndex        =   2
      Top             =   1275
      Width           =   3030
      _ExtentX        =   5345
      _ExtentY        =   8890
      Arrange         =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDropMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ilsbig"
      SmallIcons      =   "ilssmall"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDropMode     =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Name"
         Text            =   "�տ�Ա"
         Object.Width           =   3175
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   10
      Top             =   6435
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   635
      SimpleText      =   $"frmCashSupervise.frx":2F92
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCashSupervise.frx":2FD9
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13970
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
      Begin VB.Menu mnuFileSet 
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
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuPay 
      Caption         =   "�ɿ�(&J)"
      Begin VB.Menu mnuPayNewFree 
         Caption         =   "�����ֹ��ɿ�(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuPayNewAll 
         Caption         =   "����ȫ��ɿ�(&B)"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuPayNewDay 
         Caption         =   "�������սɿ�(C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuDelSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPayDelete 
         Caption         =   "ɾ���ɿ��¼(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuPayPrint 
         Caption         =   "�ش�ɿ(&P)"
      End
      Begin VB.Menu mnuEditSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditPersonGroup 
         Caption         =   "��Ա����(&F)"
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
      Begin VB.Menu mnuViewAll 
         Caption         =   "�����տ�Ա(&A)"
      End
      Begin VB.Menu mnuViewHave 
         Caption         =   "���ݴ����(&H)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuviewspilt2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "��ͼ��(&G)"
         Checked         =   -1  'True
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
      Begin VB.Menu mnuViewSplit4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFilter 
         Caption         =   "����(&I)"
      End
      Begin VB.Menu mnuViewFlash 
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
   Begin VB.Menu mnuAdd 
      Caption         =   "��ݲ˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuAddAll 
         Caption         =   "��ʾ�����տ�Ա(&A)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuAddHave 
         Caption         =   "��ʾ���ݴ�����տ�Ա(&H)"
      End
      Begin VB.Menu mnuaddsplit 
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
Attribute VB_Name = "frmCashSupervise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private msngStart As Single    '�ƶ�ǰ����λ��
Private mdatBegin As Date, mdatEnd As Date
Private mblnLoad As Boolean  '���ڻ�δ��ʱΪ��
Private mstrKey As String
Private mstrOperator As String, mstrPrivs As String, mlngModul As Long
Private mrsHandin As Recordset '�ɿ��¼
Private mblnDateMoved As Boolean '��ǰʱ�䷶Χ�Ƿ���ת��֮ǰ
Private mblnGroups As Boolean '�Ƿ���ڷ���
Private mblnNotClick As Boolean
Private Sub cboKind_Click()
    If cboKind.Text <> cboKind.Tag And Me.Visible Then
        Call FillTree
        cboKind.Tag = cboKind.Text
    End If
End Sub
Private Function LoadGroups() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������Ϣ
    '����:���˺�
    '����:�ɹ�,����true,���򷵻�False
    '����:2010-11-29 10:32:02
    '����:33633
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim lngPreID As Long, strWhere As String
    
    On Error GoTo errHandle
    gstrSQL = "" & _
    "   Select A.Id, A.������,A.����, A.˵��, A.������id, A.ɾ������,B.���� as ������  " & _
    "   From ����ɿ���� A,��Ա�� B " & _
    "   Where A.������ID=B.Id(+) And A.ɾ������>Sysdate " & _
    "   Order by ID"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If cbo��Ա��.ListIndex >= 0 Then lngPreID = cbo��Ա��.ItemData(cbo��Ա��.ListIndex)
    mblnNotClick = True
    With rsTemp
        cbo��Ա��.Clear
        mblnGroups = .RecordCount <> 0
        If Not zlStr.IsHavePrivs(mstrPrivs, "������Ա��") Then
            rsTemp.Filter = "  ������ID=" & UserInfo.ID
        End If
        Do While Not .EOF
            cbo��Ա��.AddItem Nvl(rsTemp!������)
            cbo��Ա��.ItemData(cbo��Ա��.NewIndex) = Val(Nvl(rsTemp!ID))
            If Val(Nvl(rsTemp!ID)) = lngPreID Then cbo��Ա��.ListIndex = cbo��Ա��.NewIndex
            rsTemp.MoveNext
        Loop
        If cbo��Ա��.ListCount > 0 And cbo��Ա��.ListIndex < 0 Then cbo��Ա��.ListIndex = 0
        If mblnGroups = True And cbo��Ա��.ListCount = 0 Then
            ShowMsgbox "��û���κ���Ĳ���Ȩ��,����ϵͳ����Ա��ϵ����Ȩ(������Ա��������鸺����)!"
            picGroup.Visible = False
            Call Form_Resize
            Call picGroup_Resize
            Exit Function
        End If
    End With
    picGroup.Visible = mblnGroups
    Call Form_Resize
    Call picGroup_Resize
    Call FillTree
    mblnNotClick = False
    
    LoadGroups = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
 
 
Private Sub cbo��Ա��_Click()
    If mblnNotClick = True Then Exit Sub
    '������Ա����Ϣ
    Call FillTree
End Sub
Private Sub Form_Activate()
    If mblnLoad = True Then
        Call Form_Resize 'Ϊ��ʹCoolBar����Ӧ�߶�
        If LoadGroups = False Then mblnLoad = False:   Unload Me: Exit Sub
        'If FillTree() = False Then mblnLoad = False:   Unload Me: Exit Sub
    End If
    mblnLoad = False
End Sub

Private Sub Form_Load()
    mblnLoad = True
    mstrPrivs = gstrPrivs
    mlngModul = glngModul
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    
    Call Ȩ�޿���
    Call InitFace
    '-----------
    RestoreWinState Me, App.ProductName
    
    mnuViewAll.Checked = zlDatabase.GetPara("��ʾ�����տ�Ա", glngSys, mlngModul, "0") = "1"
    mnuViewHave.Checked = Not mnuViewAll.Checked
    'Call FillTree
    '����LvwMain��ʾ���ö�Ӧ�˵�
     mnuViewIcon_Click lvwMain_S.View
End Sub

Private Sub InitFace()
    '��ʼ�����
    Dim arrTemp1 As Variant, arrTemp2 As Variant, arrTemp3 As Variant
    Dim intColumn As Integer, i As Long
    
    '��ʼ������
    mdatEnd = TruncateDate(zlDatabase.Currentdate)
    mdatBegin = TruncateDate(DateAdd("m", -1, mdatEnd))
    mblnDateMoved = zlDatabase.DateMoved(Format(mdatBegin, "yyyy-MM-dd hh:mm:ss"), , , Me.Caption)
    
    arrTemp1 = Array("����", "���㷽ʽ", "������", "�����", "��ֹʱ��", "�Ǽ���", "ժҪ", "�ɿ��")
    arrTemp2 = Array(" 1999��10��31�� ", "  ���㷽ʽ ", "-########0.00", Space(15), " yyyy-MM-dd HH:mm:ss ", " �������� ", Space(30), Space(15))
    arrTemp3 = Array(1, 1, 7, 1, 1, 1, 1, 1)
    mshRecord.Row = 0
    For intColumn = 0 To mshRecord.Cols - 1
        mshRecord.Col = intColumn
        mshRecord.Text = arrTemp1(intColumn)
        mshRecord.ColWidth(intColumn) = TextWidth(arrTemp2(intColumn))
        mshRecord.ColAlignment(intColumn) = arrTemp3(intColumn)
        mshRecord.CellAlignment = 4
    Next                              '��ʼ���ɿ��¼��
    mshRecord.ColAlignment(2) = 7
    mshRecord.MergeCol(0) = True
    
    mshTotal.ColAlignment(0) = 2
    arrTemp1 = Array("���㷽ʽ", "�ڳ��ݴ�", "�ɿ�ϼ�", "��ĩ�ݴ�")
    arrTemp2 = Array("1234567890", "123456789.123", "123456789.123", "123456789.123")
    arrTemp3 = Array(1, 7, 7, 7)
    mshTotal.Row = 0
    For intColumn = 0 To mshTotal.Cols - 1
        mshTotal.Col = intColumn
        mshTotal.Text = arrTemp1(intColumn)
        mshTotal.ColWidth(intColumn) = TextWidth(arrTemp2(intColumn))
        mshTotal.ColAlignment(intColumn) = arrTemp3(intColumn)
        mshTotal.CellAlignment = 4
    Next '��ʼ���ɿ��¼��
    
    arrTemp1 = Array("ȫ��", "����Һ�Ա", "�����շ�Ա", "Ԥ���տ�Ա", "סԺ����Ա", "��Ժ�Ǽ�Ա", "�����Ǽ���")
    cboKind.Clear
    For i = 0 To UBound(arrTemp1)
        cboKind.AddItem arrTemp1(i)
    Next
    cboKind.ListIndex = 0   '����click�¼�
    
End Sub
 
Private Sub Form_Resize()
    Dim sngTop As Single, sngBottom As Single
    On Error Resume Next
    sngTop = IIf(CoolBar1.Visible, CoolBar1.Top + CoolBar1.Height, 0)
    sngBottom = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
    
    picGroup.Top = IIf(picGroup.Visible, sngTop, 0)
    lvwMain_S.Top = IIf(picGroup.Visible, picGroup.Height + 50, 0) + sngTop
    lvwMain_S.Height = IIf(sngBottom - lvwMain_S.Top > 0, sngBottom - lvwMain_S.Top, 0)
    lvwMain_S.Left = 0
    picGroup.Width = lvwMain_S.Width
    picGroup.Left = 0
    
    picSplit.Top = sngTop
    picSplit.Height = IIf(sngBottom - picSplit.Top > 0, sngBottom - picSplit.Top, 0)
    picSplit.Left = lvwMain_S.Left + lvwMain_S.Width
    
    picContainer.Left = picSplit.Left + picSplit.Width
    picContainer.Top = sngTop
    picContainer.Width = ScaleWidth - picContainer.Left
    picContainer.Height = sngBottom - picContainer.Top
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrKey = ""
    Set mrsHandin = Nothing
    zlDatabase.SetPara "��ʾ�����տ�Ա", IIf(mnuViewAll.Checked, 1, 0), glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
    
    SaveWinState Me, App.ProductName
End Sub

Private Sub lvwMain_S_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwMain_S.SortOrder = IIf(lvwMain_S.SortOrder = lvwAscending, lvwDescending, lvwAscending)
End Sub

Private Sub lvwMain_S_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If mstrKey = Item.Key Then Exit Sub
    mstrKey = Item.Key
    
    FillList Item.Text
End Sub

Private Sub lvwMain_S_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    If Button = 2 Then
        mnuAddAll.Checked = mnuViewAll.Checked
        mnuAddHave.Checked = mnuViewHave.Checked
        For i = 0 To 3
            mnuShortIcon(i).Checked = mnuViewIcon(i).Checked
        Next
        PopupMenu mnuAdd, 2
    End If
End Sub

Private Sub mnuEditPersonGroup_Click()
    If frmGroupAndPesons.ShowGroups(Me, mlngModul, mstrPrivs) = False Then
        Exit Sub
    End If
    '���¼�������
    Call LoadGroups
End Sub

Private Sub mnuHelpTopic_Click()
   Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub mnuPayPrint_Click()
   Dim lng����ID As Long
   
   lng����ID = mshRecord.RowData(mshRecord.Row)
   If lng����ID = 0 Then Exit Sub
   
   If MsgBox("��ȷ��Ҫ�ش�" & mshRecord.TextMatrix(mshRecord.Row, 0) & "�Ľɿ��?", vbYesNo + vbQuestion, gstrSysName) = vbYes Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1500", Me, "����ID=" & lng����ID, 2)
   End If
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    If lvwMain_S.SelectedItem Is Nothing Then Exit Sub
    
    If Val(mshRecord.RowData(1)) = 0 Then
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
            "�շ�Ա=" & lvwMain_S.SelectedItem.Text)
    Else
        With mshRecord
            Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
                "�շ�Ա=" & lvwMain_S.SelectedItem.Text, "����ID=" & .RowData(.Row), "��ֹʱ��=" & .TextMatrix(.Row, MshGetColNum(mshRecord, "��ֹʱ��")), _
                "�Ǽ���=" & .TextMatrix(.Row, MshGetColNum(mshRecord, "�Ǽ���")))
        End With
    End If
End Sub

Private Sub mnuViewFlash_Click()
    'ˢ��,�ȼ���
     Call LoadGroups
    
End Sub

Private Sub mnuViewIcon_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
    Next
    mnuViewIcon(Index).Checked = True
    lvwMain_S.View = Index
End Sub

Private Sub mnuShortIcon_Click(Index As Integer)
    mnuViewIcon_Click Index
End Sub

Private Sub mnuAddAll_Click()
    mnuViewAll_Click
End Sub

Private Sub mnuAddHave_Click()
    mnuviewHave_Click
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub mnuFilePreview_Click()
    subPrint 2
End Sub

Private Sub mnuFileSet_Click()
        zlPrintSet
End Sub

Private Sub mnuPayDelete_Click()
    On Error GoTo errH
    Dim rsTmp As New Recordset
    Dim datSys As Date, i As Long, strTmp As String
    On Error GoTo errH:
    
    With mshRecord
        If .RowData(.Row) = 0 Then Exit Sub
        
        If MsgBox("��ȷʵҪɾ������Ϊ" & Trim(.TextMatrix(.Row, 0)) & "�Ľɿ�Ǽǿ���", _
            vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
            
            gstrSQL = zlGetFullFieldsTable("��Ա�ɿ��¼", 1, "Where id=[1]", False, "")
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(.RowData(.Row)))
            
            If Not rsTmp Is Nothing Then
                If Not rsTmp.EOF Then
                     MsgBox "��ǰѡ��Ľɿ��¼�ں����ݱ���!" & vbCrLf _
                         & "����ϵͳ����Ա��ϵ,ת�뵽�������ݱ��ٲ���!", vbInformation, gstrSysName
                     Exit Sub
                End If
            End If
            
            datSys = zlDatabase.Currentdate
            i = datSys - CDate(.TextMatrix(.Row, 0))
            
            If i > 1 Then strTmp = "����:����ɾ���Ľɿ��¼��������ǰ��!" & vbCrLf & vbCrLf
            strTmp = strTmp & "Ϊ������ɾ��,���ٴ�ȷ�ϲ��ܼ���." & vbCrLf & vbCrLf & "������OK"
            
            If UCase(InputBox(strTmp, "����ȷ��")) <> "OK" Then
                MsgBox "�����ȷ�Ϲؼ��ֲ���OK!��ȡ�����β���!", vbInformation + vbOKOnly, Me.Caption
                Exit Sub
            End If
            
            gstrSQL = "zl_��Ա�ɿ��¼_delete(" & .RowData(.Row) & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            If mnuViewHave.Checked = True Then
                FillTree
            Else
                FillList lvwMain_S.SelectedItem.Text
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuViewFilter_Click()
    Dim List1 As ListItem, strPersonelKind As String
    
    If cboKind.ListIndex > 0 Then strPersonelKind = cboKind.Text
    If Not lvwMain_S.SelectedItem Is Nothing Then mstrOperator = lvwMain_S.SelectedItem.Text
    If frmTimeSet.ShowMe(Me, 0, 0, mlngModul, mstrPrivs, mdatBegin, mdatEnd, mstrOperator, mblnDateMoved, strPersonelKind, mnuViewHave.Checked) = True Then
        
        If mstrOperator <> "" Then
            For Each List1 In lvwMain_S.ListItems
                If List1.Text = mstrOperator Then
                    List1.Selected = True
                    Call List1.EnsureVisible
                    Exit For
                End If
            Next
        End If
        
        If Not lvwMain_S.SelectedItem Is Nothing Then
            FillList lvwMain_S.SelectedItem.Text
        End If
    End If
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    CoolBar1.Visible = mnuViewToolButton.Checked
    CoolBar1.Bands("only").MinHeight = Toolbar1.Height
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

Private Sub mshRecord_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngCol As Long, i As Long, lngID As Long
    Dim bln�ո� As Boolean, strColName As String
    
    lngCol = mshRecord.MouseCol
    
    If Button = 1 And mshRecord.MousePointer = 99 Then
        strColName = mshRecord.TextMatrix(0, lngCol)
        If strColName = "" Then Exit Sub
        If mrsHandin Is Nothing Then Exit Sub
        
        mshRecord.ColData(lngCol) = (mshRecord.ColData(lngCol) + 1) Mod 2
        strColName = Switch(strColName = "����", "�Ǽ�ʱ��", strColName = "������", "���", strColName <> "", strColName)
        mrsHandin.Sort = strColName & IIf(mshRecord.ColData(lngCol) = 0, "", " DESC")
                
        i = 1
        Do Until mrsHandin.EOF
            If mrsHandin("����ID") <> lngID Then
                lngID = mrsHandin("����ID")
                bln�ո� = Not bln�ո�
            End If
        
            mshRecord.TextMatrix(i, 0) = Format(mrsHandin("�Ǽ�ʱ��"), "yyyy��MM��dd��") & IIf(bln�ո�, " ", "")
            mshRecord.TextMatrix(i, 1) = mrsHandin("���㷽ʽ")
            mshRecord.TextMatrix(i, 2) = Format(mrsHandin("���"), "##########0.00;-##########0.00;;")
            mshRecord.TextMatrix(i, 3) = IIf(IsNull(mrsHandin("�����")), "", mrsHandin("�����"))
            mshRecord.TextMatrix(i, 4) = Format(Nvl(mrsHandin!��ֹʱ��), "yyyy-MM-dd HH:mm:ss")
            mshRecord.TextMatrix(i, 5) = IIf(IsNull(mrsHandin("�Ǽ���")), " ", mrsHandin("�Ǽ���"))
            mshRecord.TextMatrix(i, 6) = IIf(IsNull(mrsHandin("ժҪ")), " ", mrsHandin("ժҪ"))
            mshRecord.TextMatrix(i, 7) = IIf(IsNull(mrsHandin("�ɿ��")), " ", mrsHandin("�ɿ��"))
            mshRecord.RowData(i) = mrsHandin("����ID")
            i = i + 1
            mrsHandin.MoveNext
        Loop
        mshRecord.Row = 1
    End If
End Sub

Private Sub mshRecord_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mshRecord.MouseRow = 0 Then
        mshRecord.MousePointer = 99
    Else
        mshRecord.MousePointer = 0
    End If
End Sub

Private Sub mshRecord_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And mnuPay.Visible Then PopupMenu mnuPay, 2
End Sub

Private Sub mnuPayNewFree_Click()
'���ܣ������ֹ��ɿ��¼
    If frmCashPay.�༭�ɿ��¼(lvwMain_S.SelectedItem.Text, Mid(lvwMain_S.SelectedItem.Key, 2)) = True Then
        If mnuViewHave.Checked = True Then
            FillTree
        Else
            FillList lvwMain_S.SelectedItem.Text
        End If
    End If
End Sub
Private Sub mnuPayNewDay_Click()
    If frmCashPayAll.ShowMe(lvwMain_S.SelectedItem.Text, Mid(lvwMain_S.SelectedItem.Key, 2), Me, PM_���սɿ�) Then
        If mnuViewHave.Checked Then
            Call FillTree
        Else
            Call FillList(lvwMain_S.SelectedItem.Text)
        End If
    End If
End Sub
Private Sub mnuPayNewAll_Click()
'���ܣ�����ȫ��ɿ��¼
    If frmCashPayAll.ShowMe(lvwMain_S.SelectedItem.Text, Mid(lvwMain_S.SelectedItem.Key, 2), Me, PM_ȫ��ɿ�) Then
        If mnuViewHave.Checked Then
            Call FillTree
        Else
            Call FillList(lvwMain_S.SelectedItem.Text)
        End If
    End If
End Sub

Private Sub mnuViewAll_Click()
    mnuViewAll.Checked = Not mnuViewAll.Checked
    mnuViewHave.Checked = Not mnuViewAll.Checked
    If Me.Visible Then FillTree
End Sub

Private Sub mnuviewHave_Click()
    mnuViewHave.Checked = Not mnuViewHave.Checked
    mnuViewAll.Checked = Not mnuViewHave.Checked
    
    If Me.Visible Then FillTree
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Form_Resize
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub picSplit_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then msngStart = x
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sngTemp As Single
    If Button = 1 Then
        sngTemp = picSplit.Left + x - msngStart
        If sngTemp - lvwMain_S.Left > 1500 And Me.ScaleWidth - sngTemp > 3000 Then
            picSplit.Left = sngTemp
            lvwMain_S.Width = picSplit.Left - lvwMain_S.Left
            picGroup.Width = lvwMain_S.Width
            
            picContainer.Left = picSplit.Left + picSplit.Width
            picContainer.Width = Me.ScaleWidth - picContainer.Left
            
        End If
    End If
End Sub

Private Sub picContainer_Resize()
    On Error Resume Next
    lblCaption1.Left = (picContainer.ScaleWidth - lblCaption1.Width) / 2
    mshTotal.Top = lblCaption1.Top + lblCaption1.Height + 300
    mshTotal.Left = -15
    mshTotal.Width = picContainer.ScaleWidth + 30
    mshTotal.Height = lblSplit.Top - mshTotal.Top
    
    lblSplit.Width = picContainer.ScaleWidth
    
    lblCaption2.Top = lblSplit.Top + lblSplit.Height + 100
    lblCaption2.Left = (picContainer.ScaleWidth - lblCaption2.Width) / 2
    mshRecord.Top = lblCaption2.Top + lblCaption2.Height + 300
    mshRecord.Left = -15
    mshRecord.Width = picContainer.ScaleWidth + 30
    mshRecord.Height = picContainer.ScaleHeight - mshRecord.Top
    
End Sub

Private Sub lblSplit_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then msngStart = y
End Sub

Private Sub lblSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sngTemp As Single
    If Button = 1 Then
        sngTemp = lblSplit.Top + y - msngStart
        If sngTemp - mshTotal.Top > 1000 And picContainer.ScaleHeight - sngTemp > 1500 Then
            lblSplit.Top = sngTemp
            mshTotal.Height = sngTemp - mshTotal.Top
            lblCaption2.Top = lblSplit.Top + lblSplit.Height + 100
            mshRecord.Top = lblCaption2.Top + lblCaption2.Height + 300
            mshRecord.Height = picContainer.ScaleHeight - mshRecord.Top
        End If
    End If
End Sub

 

Private Sub picGroup_Resize()
    Err = 0: On Error Resume Next
    With picGroup
        '33633
        cbo��Ա��.Width = .ScaleWidth - cbo��Ա��.Left - 50
    End With
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "New"
            If Button.ButtonMenus("PayFree").Visible Then
                mnuPayNewFree_Click
            ElseIf Button.ButtonMenus("PayAll").Visible Then
                mnuPayNewAll_Click
            ElseIf Button.ButtonMenus("PayDay").Visible Then
                mnuPayNewDay_Click
            End If
        Case "Delete"
            mnuPayDelete_Click
        Case "Filter"
            mnuViewFilter_Click
        Case "Quit"
            mnuFileExit_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Preview"
            mnuFilePreview_Click
        Case "Help"
            mnuHelpTopic_Click
        Case "View"
            mnuViewIcon(lvwMain_S.View).Checked = False
            If lvwMain_S.View = 3 Then
                mnuViewIcon(0).Checked = True
                lvwMain_S.View = 0
            Else
                mnuViewIcon(lvwMain_S.View + 1).Checked = True
                lvwMain_S.View = lvwMain_S.View + 1
            End If
    End Select
End Sub

Private Sub ClearTable()
    Dim i As Integer

    mshTotal.Rows = 2
    mshRecord.Rows = 2
    For i = 0 To mshRecord.Cols - 1
        mshRecord.TextMatrix(1, i) = ""
    Next
    For i = 0 To mshTotal.Cols - 1
        mshTotal.TextMatrix(1, i) = ""
    Next
    mshRecord.RowData(1) = 0
    Call SetMenu
End Sub

Private Function FillTree() As Boolean
'����:װ�������շ�Ա��lvwMain_S
    Dim strKey As String, strKind As String
    Dim rs�շ�Ա As New ADODB.Recordset
    Dim lng��ID As Long
    On Error GoTo errH
    
    '�õ��տ�Ա����
    mstrKey = ""
    gstrSQL = ""
    If cboKind.ListIndex > 0 Then
        strKind = cboKind.Text
        gstrSQL = " And C.��Ա����=[1]"
    ElseIf mnuViewHave.Checked = False Then
        gstrSQL = " And C.��Ա���� in ('����Һ�Ա','�����շ�Ա','Ԥ���տ�Ա','סԺ����Ա','��Ժ�Ǽ�Ա','�����Ǽ���')"
    End If
    
    If mnuViewHave.Checked = True Then
        '��ָ�㶨�ڼ������ݴ��Ĳ���Ա
        gstrSQL = "" & _
        "   Select Distinct A.�տ�Ա,B.ID " & _
        "    From ��Ա�ɿ���� A,��Ա�� B,��Ա����˵�� C" & IIf(mblnGroups, ",�ɿ��Ա��� M", "") & vbNewLine & _
        "   Where A.�տ�Ա=B.���� And ���<>0 And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null) " & _
        "           And B.id=C.��ԱID" & gstrSQL & vbNewLine & _
                IIf(mblnGroups, " And B.ID=M.��ԱID And M.��ID=[2] ", "") & _
        "   Order by �տ�Ա"
    Else
        '�����ڼ��ڲ���Ա
        gstrSQL = "" & _
        "   Select Distinct A.���� as �տ�Ա,A.ID  " & _
        "   From ��Ա�� A,��Ա����˵�� C" & IIf(mblnGroups, ",�ɿ��Ա��� M", "") & vbNewLine & _
        "   Where A.ID=C.��ԱID And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
        "           And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & gstrSQL & vbNewLine & _
                IIf(mblnGroups, " And A.ID=M.��ԱID And M.��ID=[2] ", "") & _
        "   Order by �տ�Ա"
    End If
    If cbo��Ա��.ListIndex < 0 Then
        lng��ID = 0
    Else
        lng��ID = cbo��Ա��.ItemData(cbo��Ա��.ListIndex)
    End If
    DoEvents
    Me.Refresh
    Set rs�շ�Ա = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strKind, lng��ID)
    'If rs�շ�Ա.EOF Then MsgBox "��ǰû���շ�Ա���������б�����", vbExclamation, gstrSysName: Exit Function
    If Not lvwMain_S.SelectedItem Is Nothing Then
        strKey = lvwMain_S.SelectedItem.Key
    End If
    
    With lvwMain_S.ListItems
        .Clear
        Do Until rs�շ�Ա.EOF
            If Not IsNull(rs�շ�Ա("�տ�Ա")) Then
                .Add , "C" & rs�շ�Ա("ID"), rs�շ�Ա("�տ�Ա"), "man", "man"
            End If
            rs�շ�Ա.MoveNext
        Loop
        If .Count > 0 Then
            Dim Item As ListItem
            On Error Resume Next
            Set Item = lvwMain_S.ListItems(strKey)
            If Err <> 0 Then
                Set Item = lvwMain_S.ListItems(1)
                Item.Selected = True
                Item.EnsureVisible
            Else
                Err.Clear
                Item.Selected = True
                Item.EnsureVisible
            End If
            FillList lvwMain_S.SelectedItem.Text
        Else
            FillList ""
        End If
    End With
    FillTree = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub FillList(ByVal str�շ�Ա As String)
'����:��ʾָ���շ�Ա���տ���ܱ�ͽɿ��¼
'����:str�շ�Ա �շ�Ա������
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    Dim lngID As Long
    Dim bln�ո� As Boolean, strDate As String
    
    On Error GoTo errH
    
    If str�շ�Ա = "" Then
        Call ClearTable
        Exit Sub
    End If
    
    '��ʾͳ�Ʊ�
    strDate = Format(mdatEnd, "yyyyMMdd")
    gstrSQL = _
        "Select ���㷽ʽ,sum(���+�ɿ�ϼ�) as �ڳ�,sum(�ɿ�ϼ�-��ĩ�ɿ�) as ����,sum(���+��ĩ�ɿ�) as ��ĩ from( " & _
        "Select ���㷽ʽ,��� as �ɿ�ϼ�, " & _
        "Decode(Sign(To_Char(�Ǽ�ʱ��,'YYYYMMDD')-[3]),1,���,0) as ��ĩ�ɿ�,0 as ��� " & _
        "From " & IIf(mblnDateMoved, zlGetFullFieldsTable("��Ա�ɿ��¼"), "��Ա�ɿ��¼ ") & _
        "Where �տ�Ա = [1] and �Ǽ�ʱ��>=[2] " & _
        "Union All " & _
        "Select ���㷽ʽ,0 as �ɿ�ϼ�,0 as ��ĩ�ɿ�,��� " & _
        "From ��Ա�ɿ���� " & _
        "Where ����=1 and ���<>0 and �տ�Ա =[1]) " & _
        " group by ���㷽ʽ " & _
        " having sum(���+�ɿ�ϼ�)<>0 or sum(�ɿ�ϼ�-��ĩ�ɿ�)<>0 or sum(���+��ĩ�ɿ�)<>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str�շ�Ա, mdatBegin, strDate)
    
    If rsTmp.EOF Then
        mshTotal.Rows = 2
        For i = 0 To mshTotal.Cols - 1
            mshTotal.TextMatrix(1, i) = ""
        Next
    Else
        mshTotal.Rows = rsTmp.RecordCount + 1
        i = 1
        Do Until rsTmp.EOF
            mshTotal.TextMatrix(i, 0) = rsTmp("���㷽ʽ")
            mshTotal.TextMatrix(i, 1) = Format(rsTmp("�ڳ�"), "##########0.00;-##########0.00; ;")
            mshTotal.TextMatrix(i, 2) = Format(rsTmp("����"), "##########0.00;-##########0.00; ;")
            mshTotal.TextMatrix(i, 3) = Format(rsTmp("��ĩ"), "##########0.00;-##########0.00; ;")
            i = i + 1
            rsTmp.MoveNext
        Loop
    End If
    rsTmp.Close
    
    
    '��ʾ�ɿ��¼
    gstrSQL = _
        "Select ����ID,�Ǽ�ʱ��,���㷽ʽ,���,�����,��ֹʱ��,�Ǽ���,ժҪ,B.���� �ɿ��" & _
        " From " & IIf(mblnDateMoved, zlGetFullFieldsTable("��Ա�ɿ��¼"), "��Ա�ɿ��¼") & " A,���ű� B Where A.�տ��ID=B.ID(+) And �տ�Ա=[1]" & _
        " And �Ǽ�ʱ�� Between [2] And [3] order by �Ǽ�ʱ��"
    Set mrsHandin = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str�շ�Ա, mdatBegin, DateAdd("s", -1, DateAdd("d", 1, mdatEnd)))
    
    If mrsHandin.EOF Then
        mshRecord.Rows = 2
        For i = 0 To mshRecord.Cols - 1
            mshRecord.TextMatrix(1, i) = ""
        Next
        mshRecord.RowData(1) = 0
        Call SetMenu
    Else
        mshRecord.Rows = mrsHandin.RecordCount + 1
        i = 1
        Do Until mrsHandin.EOF
            If mrsHandin("����ID") <> lngID Then
                lngID = mrsHandin("����ID")
                bln�ո� = Not bln�ո�
            End If
        
            mshRecord.TextMatrix(i, 0) = Format(mrsHandin("�Ǽ�ʱ��"), "yyyy��MM��dd��") & IIf(bln�ո�, " ", "")
            mshRecord.TextMatrix(i, 1) = mrsHandin("���㷽ʽ")
            mshRecord.TextMatrix(i, 2) = Format(mrsHandin("���"), "##########0.00;-##########0.00;;")
            mshRecord.TextMatrix(i, 3) = IIf(IsNull(mrsHandin("�����")), "", mrsHandin("�����"))
            mshRecord.TextMatrix(i, 4) = Format(Nvl(mrsHandin!��ֹʱ��), "yyyy-MM-dd HH:mm:ss")
            mshRecord.TextMatrix(i, 5) = IIf(IsNull(mrsHandin("�Ǽ���")), " ", mrsHandin("�Ǽ���"))
            mshRecord.TextMatrix(i, 6) = IIf(IsNull(mrsHandin("ժҪ")), " ", mrsHandin("ժҪ"))
            mshRecord.TextMatrix(i, 7) = IIf(IsNull(mrsHandin("�ɿ��")), " ", mrsHandin("�ɿ��"))
            mshRecord.RowData(i) = mrsHandin("����ID")
            i = i + 1
            mrsHandin.MoveNext
        Loop
        mshRecord.Row = 1
        Call SetMenu
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub subPrint(bytMode As Byte)
'����:���д�ӡ,Ԥ���������EXCEL
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Dim objPrint1Grd As New zlPrint1Grd
    Dim objAppRow As New zlTabAppRow
    
    If lvwMain_S.ListItems.Count = 0 Then Exit Sub
    objPrint1Grd.Title.Text = "�ɿ��¼��"
    objPrint1Grd.Title.Color = RGB(255, 0, 0)
    objPrint1Grd.Title.Font.Name = lblCaption2.Font.Name
    objPrint1Grd.Title.Font.Size = lblCaption2.Font.Size
    
    objAppRow.Add "�տ�Ա��" & lvwMain_S.SelectedItem.Text
    objAppRow.Add "ʱ�䷶Χ��" & Format(mdatBegin, "YYYY��MM��DD��") & "��" & Format(mdatEnd, "YYYY��MM��DD��")
    objPrint1Grd.UnderAppRows.Add objAppRow
    Set objPrint1Grd.Body = mshRecord
    
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint1Grd)
          Case 1
               zlPrintOrView1Grd objPrint1Grd, 1
          Case 2
              zlPrintOrView1Grd objPrint1Grd, 2
          Case 3
              zlPrintOrView1Grd objPrint1Grd, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint1Grd, bytMode
    End If
    
    Set objPrint1Grd = Nothing
    Set objAppRow = Nothing
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim i As Integer
    
    Select Case ButtonMenu.Key
        Case "PayFree"
            Call mnuPayNewFree_Click
        Case "PayAll"
            Call mnuPayNewAll_Click
        Case "PayDay"
            Call mnuPayNewDay_Click
        Case Else
            For i = 0 To 3
                mnuViewIcon(i).Checked = False
            Next
            mnuViewIcon(ButtonMenu.Index - 1).Checked = True
            lvwMain_S.View = ButtonMenu.Index - 1
    End Select
End Sub

Private Sub Toolbar1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
End Sub

Private Sub Ȩ�޿���()
'����:�����е��û�Ȩ�޲���,��ʹһЩ�˵����ť���ɼ�
    If InStr(mstrPrivs, "ɾ���ɿ�") = 0 And InStr(mstrPrivs, "�ֹ��ɿ�") = 0 _
        And InStr(mstrPrivs, "ȫ��ɿ�") = 0 _
        And InStr(mstrPrivs, "���սɿ�") = 0 And zlStr.IsHavePrivs(mstrPrivs, "��Ա����") = False Then
        mnuPay.Visible = False
        Toolbar1.Buttons("New").Visible = False
        Toolbar1.Buttons("Delete").Visible = False
        Toolbar1.Buttons("Split2").Visible = False
    Else
        If InStr(mstrPrivs, "�ֹ��ɿ�") = 0 And InStr(mstrPrivs, "ȫ��ɿ�") = 0 And InStr(mstrPrivs, "���սɿ�") = 0 Then
            Toolbar1.Buttons("New").Visible = False
        End If
        If InStr(mstrPrivs, "�ֹ��ɿ�") = 0 Then
            mnuPayNewFree.Visible = False
            Toolbar1.Buttons("New").ButtonMenus("PayFree").Visible = False
        End If
        If InStr(mstrPrivs, "ȫ��ɿ�") = 0 Then
            mnuPayNewAll.Visible = False
            Toolbar1.Buttons("New").ButtonMenus("PayAll").Visible = False
        End If
        If InStr(mstrPrivs, "���սɿ�") = 0 Then
            mnuPayNewDay.Visible = False
            Toolbar1.Buttons("New").ButtonMenus("PayDay").Visible = False
        End If
        If InStr(mstrPrivs, "ɾ���ɿ�") = 0 Then
            mnuPayDelete.Visible = False
            Toolbar1.Buttons("Delete").Visible = False
        End If
        If InStr(mstrPrivs, "�ش�ɿ") = 0 Then
            mnuPayPrint.Visible = False
        End If
        mnuEditPersonGroup.Visible = zlStr.IsHavePrivs(mstrPrivs, "��Ա����")
        mnuEditSplit.Visible = mnuEditPersonGroup.Visible
    End If
End Sub

Private Sub SetMenu()
    Dim blnNew As Boolean
    Dim blnDelete As Boolean
    Dim lngCount As Long, lngID As Long
    Dim i As Integer
    
    blnNew = Not (lvwMain_S.SelectedItem Is Nothing)
    blnDelete = mshRecord.RowData(mshRecord.Row) <> 0
    
    mnuPayNewFree.Enabled = blnNew
    mnuPayNewAll.Enabled = blnNew
    mnuPayNewDay.Enabled = blnNew
    Toolbar1.Buttons("New").Enabled = blnNew
    
    mnuPayDelete.Enabled = blnDelete
    Toolbar1.Buttons("Delete").Enabled = blnDelete
    mnuPayPrint.Enabled = blnDelete
    
    blnDelete = mshRecord.RowData(1) <> 0
    mnuFilePreview.Enabled = blnDelete
    mnuFilePrint.Enabled = blnDelete
    mnuFileExcel.Enabled = blnDelete
    Toolbar1.Buttons("Preview").Enabled = blnDelete
    Toolbar1.Buttons("Print").Enabled = blnDelete
    
    
    For i = 1 To mshRecord.Rows - 1
        If lngID <> mshRecord.RowData(i) Then
            lngID = mshRecord.RowData(i)
            lngCount = lngCount + 1
        End If
    Next
    If lvwMain_S.SelectedItem Is Nothing Then
        stbThis.Panels(2).Text = ""
    Else
        stbThis.Panels(2).Text = lvwMain_S.SelectedItem.Text & "��" & _
            Format(mdatBegin, "yyyy��MM��dd��") & "����" & _
            Format(mdatEnd, "yyyy��MM��dd��") & "֮�乲��" & lngCount & "���ɿ��¼��"
    End If
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

