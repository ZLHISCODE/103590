VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBillSupervise 
   Caption         =   "Ʊ��ʹ�ü��"
   ClientHeight    =   6510
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   9195
   Icon            =   "frmBillSupervise.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   6150
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   635
      SimpleText      =   $"frmBillSupervise.frx":0442
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmBillSupervise.frx":0489
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11139
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
   Begin VB.PictureBox picH 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   6480
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleMode       =   0  'User
      ScaleWidth      =   1530.013
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2820
      Width           =   1785
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2715
      Left            =   3180
      ScaleHeight     =   2715
      ScaleMode       =   0  'User
      ScaleWidth      =   38.572
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3120
      Width           =   45
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh���� 
      Height          =   2145
      Left            =   4170
      TabIndex        =   5
      Top             =   3480
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   3784
      _Version        =   393216
      Rows            =   3
      Cols            =   4
      FixedRows       =   2
      FixedCols       =   0
      BackColorFixed  =   -2147483648
      BackColorBkg    =   -2147483643
      BackColorUnpopulated=   -2147483644
      GridColor       =   8421504
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   0
      GridLinesFixed  =   1
      MergeCells      =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin MSComctlLib.ImageList Ilscolor 
      Left            =   7965
      Top             =   930
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
            Picture         =   "frmBillSupervise.frx":0D1D
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":0F3D
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":115D
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":137D
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":159D
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":17BD
            Key             =   "Cancel"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":19DD
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":1BFD
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":1E17
            Key             =   "View"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":2033
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":2253
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":2473
            Key             =   "Check"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Ilsrw 
      Left            =   8670
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
            Picture         =   "frmBillSupervise.frx":268D
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":28A7
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":2AC7
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":2CE7
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":2F07
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":3127
            Key             =   "Cancel"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":3347
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":3567
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":3781
            Key             =   "View"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":399D
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":3BBD
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":3DDD
            Key             =   "Check"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   9195
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "Toolbar1"
      MinHeight1      =   720
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Caption2        =   "ʹ�����"
      Child2          =   "cbo���"
      MinWidth2       =   1995
      MinHeight2      =   300
      Width2          =   1695
      NewRow2         =   0   'False
      Begin VB.ComboBox cbo��� 
         Height          =   300
         Left            =   7110
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   1995
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   720
         Left            =   165
         TabIndex        =   7
         Top             =   30
         Width           =   5940
         _ExtentX        =   10478
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
               Key             =   "New"
               Object.ToolTipText     =   "����Ʊ��"
               Object.Tag             =   "����"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "Modify"
               Object.ToolTipText     =   "�޸ļ�¼"
               Object.Tag             =   "�޸�"
               ImageKey        =   "Modify"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "Delete"
               Object.ToolTipText     =   "ɾ����¼"
               Object.Tag             =   "ɾ��"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split1"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Cancel"
               Object.ToolTipText     =   "Ʊ�ݱ���"
               Object.Tag             =   "����"
               ImageKey        =   "Cancel"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˶�"
               Key             =   "Check"
               Object.ToolTipText     =   "�˶�Ʊ����ϸ"
               Object.Tag             =   "�˶�"
               ImageKey        =   "Check"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split2"
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Filter"
               Description     =   "��������"
               Object.ToolTipText     =   "��������"
               Object.Tag             =   "����"
               ImageKey        =   "Filter"
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
   Begin MSComctlLib.ImageList ils32 
      Left            =   2880
      Top             =   1350
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
            Picture         =   "frmBillSupervise.frx":3FF7
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":4449
            Key             =   "C2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":4763
            Key             =   "C3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":4A7D
            Key             =   "C5"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":4D97
            Key             =   "C1"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":50B1
            Key             =   "C4"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":53CB
            Key             =   "C7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":56E5
            Key             =   "C6"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   2970
      Top             =   2220
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
            Picture         =   "frmBillSupervise.frx":59FF
            Key             =   "Item"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvw����_S 
      Height          =   1185
      Left            =   3960
      TabIndex        =   8
      Top             =   900
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   2090
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
   Begin MSComctlLib.ListView lvwMain 
      Height          =   3345
      Left            =   330
      TabIndex        =   9
      Top             =   1890
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   5900
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ils32"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Ʊ������"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblDown 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ʹ�����"
      Height          =   240
      Left            =   7170
      TabIndex        =   4
      Top             =   3030
      Width           =   1095
   End
   Begin VB.Label lblUp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "���ü�¼"
      Height          =   240
      Left            =   7020
      TabIndex        =   3
      Top             =   1110
      Width           =   1095
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
      Begin VB.Menu mnuFileSpit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuBill 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuBillGet 
         Caption         =   "����Ʊ��(&N)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuBillModify 
         Caption         =   "�޸ļ�¼(&M)"
      End
      Begin VB.Menu mnuBillDelete 
         Caption         =   "ɾ����¼(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuBillSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBillCancel 
         Caption         =   "Ʊ�ݱ���(&R)"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuBillCheck 
         Caption         =   "�˶����õ�(&B)"
         Index           =   0
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuBillCheck 
         Caption         =   "�˶�Ʊ����ϸ(&C)"
         Index           =   1
         Shortcut        =   ^H
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
         Caption         =   "��ʾ�������ü�¼(&A)"
      End
      Begin VB.Menu mnuViewHave 
         Caption         =   "����ʾδ����(&P)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewCheck 
         Caption         =   "��ʾ�˶���Ϣ(&H)"
      End
      Begin VB.Menu mnuviewsplit2 
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
      Begin VB.Menu mnuViewDetail 
         Caption         =   "��ϸ�嵥(&E)"
      End
      Begin VB.Menu mnuViewSelect 
         Caption         =   "ѡ����(&C)"
      End
      Begin VB.Menu mnuViewFilter 
         Caption         =   "����(&I)"
      End
      Begin VB.Menu mnuViewSplit45 
         Caption         =   "-"
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
         Caption         =   "��ʾ�������ü�¼(&A)"
      End
      Begin VB.Menu mnuAddHave 
         Caption         =   "����ʾδ����(&P)"
      End
      Begin VB.Menu mnuAddSplit 
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
   Begin VB.Menu mnuAdd2 
      Caption         =   "��ݲ˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuAddDetail 
         Caption         =   "��ϸ�嵥(&D)"
      End
   End
End
Attribute VB_Name = "frmBillSupervise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private msngStart As Single    '�ƶ�ǰ����λ��
Private mdatBegin As Date, mdatEnd As Date
Private mstrOperator As String  '���˵�Ʊ��������
Private mlngModul As Long

Private mblnUnload As Boolean
Private mblnLoad As Boolean  '���ڻ�δ��ʱΪ��
Private mblnItem As Boolean  'Ϊ���ʾ������ListViewĳһ����
Private mintColumn As Integer '
Private mstrƱ�� As String   '��һ�ε�Ʊ������
Private mstrƱ�ݳ���  As String  'Ʊ�ݳ���
Private mstrKey As String    '��һ�εļ�¼
Private mstrPrivs As String
Private Const mstrLvw As String = "��ʼ����,1000,0,1;��ֹ����,1000,0,2;ʹ�����,1000,0,1;" & _
    "������,800,0,2;��ǰ����,1000,0,0;ʣ������,600,0,0;" & _
    "����,1000,0,0;ʹ�÷�ʽ,600,0,0;�Ǽ�ʱ��,1200,0,0;�Ǽ���,800,0,2;�˶���,800,0,2;" & _
    "ǰ׺�ı�,0,0,2;ǩ����,800,0,2;ǩ��ʱ��,1200,0,2"
Private mblnNotClick As Boolean
Private mblnNOMoved As Boolean '��ǰƱ���ǲ����ں����ݱ���
Private mblnDateMoved As Boolean '��ǰʱ�䷶Χ�Ƿ���ת��֮ǰ
Private mblnҩ��  As Boolean

Private Sub cbo���_Click()
    Call SetDefaultUserType
    If mblnNotClick = True Then Exit Sub
    mstrKey = ""
    Call Fill��¼
End Sub

Private Sub Form_Activate()
    If mblnUnload Then Unload Me: Exit Sub
    Call LoadCombox
    
    If mblnLoad = True Then
        Call Form_Resize 'Ϊ��ʹCoolBar����Ӧ�߶�
        'Call Fill��¼
         If lvwMain.Enabled And lvwMain.View Then lvwMain.SetFocus
        Call lvwMain_ItemClick(lvwMain.SelectedItem)
    End If
End Sub

Private Sub Form_Load()
    mblnLoad = True
    mblnUnload = False
    mstrPrivs = gstrPrivs
    mlngModul = glngModul
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    
    Call PrivilegeCTRL
    If Not InitFace Then
        mblnUnload = True
        Exit Sub
    End If
    lvw����_S.Tag = "�ɱ仯��"
    '-----------
    RestoreWinState Me, App.ProductName
    
    mnuViewAll.Checked = zlDatabase.GetPara("��ʾ�������ü�¼", glngSys, mlngModul, "0") = "1"
    mnuViewHave.Checked = Not mnuViewAll.Checked
    mnuViewCheck.Checked = zlDatabase.GetPara("�鿴�˶���Ϣ", glngSys, mlngModul, "0") = "1"
    '���ListView�Ļ�δ�����ã������һ��ʹ�ã��Ǿ͵���ȱʡ�ĳ�ʼ��
    If lvw����_S.ColumnHeaders.Count <> UBound(Split(mstrLvw, ";")) + 1 Then
        zlControl.LvwSelectColumns lvw����_S, mstrLvw, True
    End If
    '����lvw����_S��ʾ���ö�Ӧ�˵�
     mnuViewIcon_Click lvw����_S.View
     
     
    '����������Ʊ�ݴ�ӡ����
    On Error Resume Next
    gblnBillPrint = False
    Set gobjBillPrint = CreateObject("zlBillPrint.clsBillPrint")
    If Not gobjBillPrint Is Nothing Then
        gblnBillPrint = gobjBillPrint.zlInitialize(gcnOracle, glngSys, glngModul, UserInfo.���, UserInfo.����)
    End If
    Err.Clear: On Error GoTo 0
End Sub

Private Function InitFace() As Boolean
    Dim arrTemp1 As Variant, arrTemp2 As Variant, arrTemp3 As Variant
    Dim i As Integer, strTmp As String
    Dim objListItem As ListItem, strKeyValue As String
    
    '��ʼ������
    mdatEnd = TruncateDate(zlDatabase.Currentdate)
    mdatBegin = TruncateDate(DateAdd("m", -1, mdatEnd))
    mblnDateMoved = zlDatabase.DateMoved(Format(mdatBegin, "yyyy-MM-dd hh:mm:ss"), , , Me.Caption)
    
    If InStr(mstrPrivs, ";���в���Ա;") = 0 Then
        mstrOperator = UserInfo.����
    Else
        mstrOperator = ""
    End If
    
    mstrƱ�ݳ��� = zlDatabase.GetPara(20, glngSys, , "7|7|7|7|7")
    If mblnҩ�� = False Then
        If zlStr.IsHavePrivs(mstrPrivs, "�շ��վ�") Then
            strTmp = strTmp & "|" & "�շ��վ�,1"
        End If
        If zlStr.IsHavePrivs(mstrPrivs, "Ԥ���վ�") And _
            (zlStr.IsHavePrivs(mstrPrivs, "Ԥ������Ʊ��") _
                Or zlStr.IsHavePrivs(mstrPrivs, "Ԥ��סԺƱ��")) Then
            strTmp = strTmp & "|" & "Ԥ���վ�,2"
        End If
        If zlStr.IsHavePrivs(mstrPrivs, "�����վ�") Then
            strTmp = strTmp & "|" & "�����վ�,3"
        End If
        If zlStr.IsHavePrivs(mstrPrivs, "�Һ��վ�") Then
            strTmp = strTmp & "|" & "�Һ��վ�,4"
        End If
        If zlStr.IsHavePrivs(mstrPrivs, "ҽ�ƿ�") Then
            strTmp = strTmp & "|" & "ҽ�ƿ�,5"
        End If
        If zlStr.IsHavePrivs(mstrPrivs, "���ѿ�") Then
            strTmp = strTmp & "|" & "���ѿ�,6"
        End If
        If strTmp = "" Then
            MsgBox "��û�в����κ�Ʊ�ݵ�Ȩ��!", vbInformation, App.ProductName
            Exit Function
        Else
            strTmp = Mid(strTmp, 2)
        End If
        
        arrTemp1 = Split(strTmp, "|")
        For i = 0 To UBound(arrTemp1)
            arrTemp2 = Split(arrTemp1(i), ",")
            Set objListItem = lvwMain.ListItems.Add(, "C" & arrTemp2(1), arrTemp2(0), "C" & arrTemp2(1))
            
            GetRegInFor g˽��ģ��, Me.Name, "C" & arrTemp2(1), strKeyValue
            objListItem.Tag = strKeyValue
        Next
    Else
        lvwMain.ListItems.Add , "C1", "�շ��վ�", "C1"
        lvwMain.ListItems.Add , "C5", "��Ա��", "C7"
    End If
    lvwMain.ListItems(1).Selected = True
    
    '��ʼ�����
    arrTemp1 = Array("ʹ����", "ʹ��", "ʹ��", "ʹ��", "ʹ��", "�ջ�", "�ջ�", "�ջ�")
    arrTemp2 = Array("ʹ����", "����", "�ش�", "����", "����", "����", "�ش�", "����")
    arrTemp3 = Array(1000, 800, 800, 800, 800, 800, 800, 800)
    With msh����
        .Cols = 8
        .MergeCol(0) = True
        .MergeRow(0) = True
        .ColAlignment(0) = 1
        For i = 0 To .Cols - 1
            .TextMatrix(0, i) = arrTemp1(i)
            .TextMatrix(1, i) = arrTemp2(i)
            .ColWidth(i) = arrTemp3(i)
        Next                              '��ʼ���ɿ��¼��
        .AllowBigSelection = True
        .FillStyle = flexFillRepeat
        .Row = 0: .Col = 0: .RowSel = 1: .ColSel = .Cols - 1
        .CellAlignment = 4
        .FillStyle = flexFillSingle
        .AllowBigSelection = False
    End With
    
    InitFace = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Resize()
    Dim sngTop As Single, sngBottom As Single
    
    On Error Resume Next
    sngTop = IIf(cbrThis.Visible, cbrThis.Top + cbrThis.Height, 0)
    sngBottom = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
    
    lvwMain.Top = sngTop
    lvwMain.Height = IIf(sngBottom - lvwMain.Top > 0, sngBottom - lvwMain.Top, 0)
    lvwMain.Left = 0
    
    picSplit.Top = sngTop
    picSplit.Height = IIf(sngBottom - picSplit.Top > 0, sngBottom - picSplit.Top, 0)
    picSplit.Left = lvwMain.Left + lvwMain.Width
    
    lblUp.Top = sngTop
    lblUp.Left = picSplit.Left + picSplit.Width
    If Me.ScaleWidth - lblUp.Left > 0 Then lblUp.Width = Me.ScaleWidth - lblUp.Left
    
    lvw����_S.Left = lblUp.Left
    lvw����_S.Top = lblUp.Top + lblUp.Height
    lvw����_S.Width = lblUp.Width
    
    picH.Left = lblUp.Left
    picH.Top = lvw����_S.Top + lvw����_S.Height
    picH.Width = lblUp.Width
    
    lblDown.Left = lblUp.Left
    lblDown.Top = picH.Top + picH.Height
    lblDown.Width = lblUp.Width
    
    msh����.Left = lblUp.Left
    msh����.Top = lblDown.Top + lblDown.Height
    msh����.Width = lblUp.Width
    msh����.Height = sngBottom - msh����.Top
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    
    mstrKey = ""
    mstrƱ�� = ""
    mblnItem = False
    zlDatabase.SetPara "��ʾ�������ü�¼", IIf(mnuViewAll.Checked, 1, 0), glngSys, mlngModul, zlStr.IsHavePrivs(mstrPrivs, "��������")
    zlDatabase.SetPara "�鿴�˶���Ϣ", IIf(mnuViewCheck.Checked, 1, 0), glngSys, mlngModul, zlStr.IsHavePrivs(mstrPrivs, "��������")
    
    SaveWinState Me, App.ProductName
    For i = 1 To lvwMain.ListItems.Count
        SaveRegInFor g˽��ģ��, Me.Name, lvwMain.ListItems(i).Key, lvwMain.ListItems(i).Tag
    Next
    
    If Not gobjBillPrint Is Nothing Then
        Call gobjBillPrint.zlTerminate
        Set gobjBillPrint = Nothing
    End If
End Sub

Private Sub lvwMain_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    lvwMain.Drag 0
End Sub

Private Sub lvwMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If mstrƱ�� = Item.Key Then Exit Sub

    Call LoadCombox
    mstrƱ�� = Item.Key
    
    '�����б�����ʾ����
    If CurrentIsBill(Val(Mid(Item.Key, 2))) Then
        lvw����_S.ColumnHeaders(1).Text = "��ʼ����"
        lvw����_S.ColumnHeaders(2).Text = "��ֹ����"
        lvw����_S.ColumnHeaders(5).Text = "��ǰ����"
    Else
        lvw����_S.ColumnHeaders(1).Text = "��ʼ����"
        lvw����_S.ColumnHeaders(2).Text = "��ֹ����"
        lvw����_S.ColumnHeaders(5).Text = "��ǰ����"
    End If
    
    Call Fill��¼
End Sub

Private Sub lvwMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If lvwMain.HitTest(x, y) Is Nothing Then Exit Sub
        
        lvwMain.Drag 1
    End If
End Sub

Private Sub mnuAddDetail_Click()
    Call mnuViewDetail_Click
End Sub

Private Sub mnuBillCheck_Click(Index As Integer)
    Dim lng����ID As Long, strǰ׺ As String, blnChecked As Boolean
    Dim lngƱ�� As gBillType, strSQL As String
    
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    
    lngƱ�� = Val(Mid(lvwMain.SelectedItem.Key, 2))
    If lvw����_S.SelectedItem Is Nothing And Index = 1 Then
        Call frmBillUses.ShowMe(Me, mstrPrivs, 1, True, mblnNOMoved, lngƱ��, 0, "")
        Exit Sub
    End If
    
    If lvw����_S.SelectedItem Is Nothing Then Exit Sub
    lng����ID = Val(Mid(lvw����_S.SelectedItem.Key, 2))
    If Index = 0 Then
        blnChecked = (lvw����_S.SelectedItem.SubItems(GetItemCOL("�˶���")) <> "")
        If blnChecked Then
            If MsgBox("��ȷ��Ҫȡ�������õ��ĺ˶Լ�¼��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub
            
            On Error GoTo errHandle
            If lngƱ�� = gBillType.���ѿ� Then
                ' Zl_���ѿ����ü�¼_Check
                strSQL = " Zl_���ѿ����ü�¼_Check("
                '  Id_In       ���ѿ����ü�¼.Id%Type,
                strSQL = strSQL & "" & lng����ID & ","
                '  �˶Խ��_In ���ѿ����ü�¼.�˶Խ��%Type,
                strSQL = strSQL & "" & "NULL" & ","
                '  �˶���_In   ���ѿ����ü�¼.�˶���%Type,
                strSQL = strSQL & "" & "NULL" & ","
                '  ��ע_In     ���ѿ����ü�¼.��ע%Type,
                strSQL = strSQL & "" & "NULL" & ","
                '  �˶�ģʽ_In ���ѿ����ü�¼.�˶�ģʽ%Type
                strSQL = strSQL & "" & "NULL" & ")"
            Else
                'Zl_Ʊ�����ü�¼_Check
                strSQL = "Zl_Ʊ�����ü�¼_Check("
                '  Id_In       In Ʊ�����ü�¼.Id%Type,
                strSQL = strSQL & "" & lng����ID & ","
                '  �˶Խ��_In In Ʊ�����ü�¼.�˶Խ��%Type,
                strSQL = strSQL & "" & "NULL" & ","
                '  �˶���_In   In Ʊ�����ü�¼.�˶���%Type,
                strSQL = strSQL & "" & "NULL" & ","
                '  ��ע_In     In Ʊ�����ü�¼.��ע%Type,
                strSQL = strSQL & "" & "NULL" & ","
                '  �˶�ģʽ_In In Ʊ�����ü�¼.�˶�ģʽ%Type
                strSQL = strSQL & "" & "NULL" & ")"
            End If
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        Else
            If frmBillEdit.ShowMe(Me, 1, mlngModul, mstrPrivs, lng����ID, , lngƱ��) = False Then Exit Sub
        End If
        Call Fill��¼
        Call SetMenu
    Else
        strǰ׺ = lvw����_S.SelectedItem.SubItems(GetItemCOL("ǰ׺�ı�"))
        Call frmBillUses.ShowMe(Me, mstrPrivs, 1, True, mblnNOMoved, lngƱ��, lng����ID, strǰ׺)
            
        Call Fill��¼
        If mnuViewCheck.Checked Then Fill����
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub

Private Sub mnuBillDelete_Click()
    On Error GoTo errHandle
    Dim intIndex As Long
    Dim lngƱ�� As gBillType
    
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    
    lngƱ�� = Val(Mid(lvwMain.SelectedItem.Key, 2))
    If MsgBox("��ȷ��Ҫɾ����ʼ" & _
        IIf(lngƱ�� = gBillType.���￨ Or lngƱ�� = gBillType.���ѿ�, "����", "����") & _
        "Ϊ��" & lvw����_S.SelectedItem.Text & "����" & lvwMain.SelectedItem.Text & "���ü�¼��", _
        vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    If mblnNOMoved Then
        MsgBox "��ǰѡ������ü�¼�ں����ݱ���!" & vbCrLf _
            & "����ϵͳ����Ա��ϵ,ת�뵽�������ݱ��ٲ���!", vbInformation, gstrSysName
        Exit Sub
    End If
    If zlIsModify(Val(Mid(lvw����_S.SelectedItem.Key, 2))) = False Then Exit Sub
    
    Me.MousePointer = 11
    If lngƱ�� = gBillType.���ѿ� Then
        gstrSQL = "Zl_���ѿ����ü�¼_Delete(" & Mid(lvw����_S.SelectedItem.Key, 2) & ")"
    Else
        gstrSQL = "zl_Ʊ�����ü�¼_delete(" & Mid(lvw����_S.SelectedItem.Key, 2) & ")"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)

    Me.MousePointer = 0
    
    With lvw����_S
        intIndex = .SelectedItem.Index
        .ListItems.Remove .SelectedItem.Key
        If .ListItems.Count > 0 Then
            intIndex = IIf(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
            .ListItems(intIndex).Selected = True
            .ListItems(intIndex).EnsureVisible
        End If
    End With
    Call Fill����
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Me.MousePointer = 0
End Sub

Private Sub mnuBillGet_Click()
    Dim intƱ�� As gBillType, str��� As String
    
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    
    intƱ�� = Val(Mid(lvwMain.SelectedItem.Key, 2))
    Select Case intƱ��
    Case gBillType.�շ��վ�, gBillType.�����վ�
        str��� = Trim(cbo���.Text)
    Case gBillType.Ԥ���վ�, gBillType.���￨, gBillType.���ѿ�
        If cbo���.ListIndex < 0 Then Exit Sub
        str��� = cbo���.ItemData(cbo���.ListIndex)
    End Select
    If frmBillEdit.ShowMe(Me, 0, mlngModul, mstrPrivs, 0, str���, intƱ��) = False Then Exit Sub
    
    Call Fill��¼
    Call Fill����
End Sub

Private Sub mnuBillModify_Click()
    Dim lngLen As Long
    Dim intƱ�� As gBillType

    If lvw����_S.SelectedItem Is Nothing Or mnuBill.Visible = False Then Exit Sub
    
    intƱ�� = Val(Mid(lvwMain.SelectedItem.Key, 2))
    '102181:���ϴ�,2016/11/10,ҽ�ƿ�Ʊ�ݳ���
    If CurrentIsBill(intƱ��) = True Then
        lngLen = Val(Split(mstrƱ�ݳ���, "|")(Mid(lvwMain.SelectedItem.Key, 2) - 1))
        If msh����.Rows > 3 And Len(lvw����_S.SelectedItem.Text) <> lngLen Then
            MsgBox lvwMain.SelectedItem.Text & "�ĺ���涨����Ӧ����" & lngLen & "λ��" & _
                vbCrLf & "������¼�ĺ��볤�Ȳ����涨����������ʹ�ã��ʲ����޸ġ�", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    If mblnNOMoved Then
        MsgBox "��ǰѡ������ü�¼�ں����ݱ���!" & vbCrLf _
            & "����ϵͳ����Ա��ϵ,ת�뵽�������ݱ��ٲ���!", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If zlIsModify(Val(Mid(lvw����_S.SelectedItem.Key, 2))) = False Then Exit Sub
    
    
    If frmBillEdit.ShowMe(Me, 0, mlngModul, mstrPrivs, Mid(lvw����_S.SelectedItem.Key, 2), , intƱ��) = False Then Exit Sub
    
    Call Fill��¼
    Call Fill����
End Sub

Public Function zlIsModify(ByVal lngID As Long, Optional blnMsg As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ƿ������޸����˵�Ʊ��
    '���:lngID-����ID
    '     blnMsg-�Ƿ���ʾ��Ϣ
    '����:
    '����:�����޸�,����true,���򷵻�False
    '����:���˺�
    '����:2010-02-01 10:49:54
    '����:27372
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    '��鵱ǰ�Ƿ������޸����ǵĵ���
    If zlStr.IsHavePrivs(mstrPrivs, "����������˵Ǽ�Ʊ��") Then
       zlIsModify = True: Exit Function
    End If
    '��Ϊ�ڶ�ȡʱ���Ѿ��м��,�ֲ������ж�.��δ����Ժ���ܴ��ڸĶ������Ա���
    zlIsModify = True: Exit Function
    '����Ƿ�Ϊ������
    gstrSQL = "Select ID From Ʊ�����ü�¼ where id=[1] and �Ǽ���=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngID, UserInfo.����)
    If rsTemp.EOF And blnMsg Then
        ShowMsgbox "ע��:" & vbCrLf & "    �㲻�ܲ��������˵Ǽǵ�Ʊ��!"
    End If
    zlIsModify = rsTemp.EOF = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Sub mnuBillCancel_Click()
    If lvw����_S.SelectedItem Is Nothing Then Exit Sub
    If mblnNOMoved Then
        MsgBox "��ǰѡ������ü�¼�ں����ݱ���!" & vbCrLf _
            & "����ϵͳ����Ա��ϵ,ת�뵽�������ݱ��ٲ���!", vbInformation, gstrSysName
        Exit Sub
    End If
    If zlIsModify(Val(Mid(lvw����_S.SelectedItem.Key, 2))) = False Then Exit Sub
    If frmBillDiscard.�༭Ʊ�ݱ���(Me, mstrPrivs, _
        Val(Mid(lvwMain.SelectedItem.Key, 2)), Val(Mid(lvw����_S.SelectedItem.Key, 2))) Then
        Call Fill��¼
    End If
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

Private Sub mnuReportItem_Click(Index As Integer)
    Dim str������ As String, str����ID As String
    
    If Not lvw����_S.SelectedItem Is Nothing Then
        str����ID = Mid(lvw����_S.SelectedItem.Key, 2)
        str������ = lvw����_S.SelectedItem.SubItems(GetItemCOL("������"))
    End If
    Call ReportOpen(gcnOracle, _
        Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
        "Ʊ��=" & Val(Mid(lvwMain.SelectedItem.Key, 2)), "������=" & str������, "����ID=" & str����ID)
End Sub

Private Sub mnuViewCheck_Click()
    mnuViewCheck.Checked = Not mnuViewCheck.Checked
    Call Fill����
End Sub

Private Sub mnuViewDetail_Click()
    Dim lng����ID As Long, lngԭ�� As Long, lng���� As Long
    Dim strCondition As String, str��ʾ As String, strʹ���� As String, strǰ׺ As String
    Dim blnOne As Boolean, lngƱ�� As gBillType
    
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    
    With msh����
        If .Rows = 3 Or .TextMatrix(.Row, .Col) = " " Then
            Exit Sub
        End If
        
        blnOne = .Row < 2 Or .Row = .Rows - 1
        Select Case .Col
            Case 0
                str��ʾ = "ȫ����ϸ�嵥"
                strCondition = ""
            Case 1
                str��ʾ = "����ʹ����ϸ�嵥"
                strCondition = " and ԭ��=[2]": lngԭ�� = 1
            Case 2
                str��ʾ = "�ش�ʹ����ϸ�嵥"
                strCondition = " and ԭ��=[2]": lngԭ�� = 3
            Case 3
                str��ʾ = "����ʹ����ϸ�嵥"
                strCondition = " and ԭ��=[2]": lngԭ�� = 5
            Case 4
                str��ʾ = "ȫ��ʹ����ϸ�嵥"
                strCondition = " and ����=[3]": lng���� = 1
            Case 5
                str��ʾ = "�����ջ���ϸ�嵥"
                strCondition = " and ԭ��=[2]": lngԭ�� = 2
            Case 6
                str��ʾ = "�ش��ջ���ϸ�嵥"
                strCondition = " and ԭ��=[2]": lngԭ�� = 4
            Case 7
                str��ʾ = "ȫ���ջ���ϸ�嵥"
                strCondition = " and ����=[3]": lng���� = 2
        End Select
        If blnOne = False Then
            strʹ���� = .TextMatrix(.Row, 0)
            str��ʾ = strʹ���� & "��" & str��ʾ
            strCondition = strCondition & " and ʹ����||''=[4]"
        Else
            str��ʾ = "�����˵�" & str��ʾ
        End If
    End With
    lngƱ�� = Val(Mid(lvwMain.SelectedItem.Key, 2))
    lng����ID = Val(Mid(lvw����_S.SelectedItem.Key, 2))
    strǰ׺ = lvw����_S.SelectedItem.SubItems(GetItemCOL("ǰ׺�ı�"))
    
    Call frmBillUses.ShowMe(Me, mstrPrivs, 0, mnuViewCheck.Checked, mblnNOMoved, _
        lngƱ��, lng����ID, strǰ׺, strCondition, lngԭ��, lng����, strʹ����, str��ʾ)
End Sub

Private Function GetItemCOL(strColName As String)
'����:�������Ʒ���listsubitems�б���к�
    Dim lngCol As Long
    
    For lngCol = 2 To lvw����_S.ColumnHeaders.Count
        'ColumnHeaders�ĵ�һ����listitem�Ŀ�ʼ����,listsubitems�ĵ�һ���Ǵ�ColumnHeaders�ĵ�2�п�ʼ��
        If lvw����_S.ColumnHeaders(lngCol).Text = strColName Then
            GetItemCOL = lngCol - 1
            Exit For
        End If
    Next
End Function

Private Sub mnuViewFlash_Click()
    Call Fill��¼
End Sub

Private Sub mnuAddAll_Click()
    mnuViewAll_Click
End Sub

Private Sub mnuAddHave_Click()
    mnuviewHave_Click
End Sub

Private Sub lvw����_S_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then '���Ǹղ�����
        lvw����_S.SortOrder = IIf(lvw����_S.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvw����_S.SortKey = mintColumn
        lvw����_S.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvw����_S_DblClick()
    If mblnItem = True And mnuBillModify.Enabled And mnuBillModify.Visible Then
        Call mnuBillModify_Click
    End If
End Sub

Public Sub lvw����_S_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim rsTmp As Recordset
    Dim intƱ�� As gBillType
    
    mblnItem = True
    
    On Error GoTo errHandle
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    If mstrKey = Item.Key Then Exit Sub
    mstrKey = Item.Key
    
    intƱ�� = Val(Mid(lvwMain.SelectedItem.Key, 2))
    '��ǰ���ü�¼�Ƿ��ں󱸱���
    mblnNOMoved = False
    If mblnDateMoved Then
        If intƱ�� = gBillType.���ѿ� Then
            gstrSQL = "Select id From H���ѿ����ü�¼ Where id=[1]"
        Else
            gstrSQL = "Select id From HƱ�����ü�¼ Where id=[1]"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Mid(lvw����_S.SelectedItem.Key, 2))
        If rsTmp.RecordCount > 0 Then mblnNOMoved = True
    End If
    
    Call Fill����
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub lvw����_S_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If mnuBillModify.Enabled And mnuBillModify.Visible Then
            Call mnuBillModify_Click
        End If
    End If
End Sub
 
 Sub lvw����_S_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    If Button = 2 Then
        mnuAddAll.Checked = mnuViewAll.Checked
        mnuAddHave.Checked = mnuViewHave.Checked
        For i = 0 To 3
            mnuShortIcon(i).Checked = mnuViewIcon(i).Checked
        Next
        PopupMenu mnuAdd, vbPopupMenuRightButton
    End If
End Sub

Private Sub mnuViewFilter_Click()
    Dim lngKind As Long
    
    lngKind = Val(Mid(lvwMain.SelectedItem.Key, 2))
    If frmTimeSet.ShowMe(Me, 1, lngKind, mlngModul, mstrPrivs, _
        mdatBegin, mdatEnd, mstrOperator, mblnDateMoved) Then
        Call Fill��¼
    End If
End Sub

Private Sub mnuViewSelect_Click()
    If zlControl.LvwSelectColumns(lvw����_S, mstrLvw) = True Then
        '���б仯��Ҫ����ˢ��
        Fill��¼
    End If
End Sub

Private Sub msh����_DblClick()
    If mnuViewDetail.Enabled = True And mnuViewDetail.Visible = True Then
        Call mnuViewDetail_Click
    End If
End Sub

Private Sub msh����_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If msh����.Rows = 3 Then Exit Sub
    msh����.SetFocus
    If Button = 2 Then PopupMenu mnuAdd2, vbPopupMenuRightButton
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim i As Integer
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
    Next
    mnuViewIcon(ButtonMenu.Index - 1).Checked = True
    lvw����_S.View = ButtonMenu.Index - 1
End Sub

Private Sub mnuViewIcon_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
    Next
    mnuViewIcon(Index).Checked = True
    lvw����_S.View = Index
End Sub

Private Sub mnuShortIcon_Click(Index As Integer)
    mnuViewIcon_Click Index
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

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
    cbrThis.Bands("only").minHeight = Toolbar1.Height
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
    cbrThis.Bands("only").minHeight = Toolbar1.Height
    Form_Resize
End Sub

Private Sub mnuViewAll_Click()
    mnuViewAll.Checked = Not mnuViewAll.Checked
    mnuViewHave.Checked = Not mnuViewAll.Checked
    Fill��¼
End Sub

Private Sub mnuviewHave_Click()
    mnuViewHave.Checked = Not mnuViewHave.Checked
    mnuViewAll.Checked = Not mnuViewHave.Checked
    Fill��¼
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

Private Sub picH_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then msngStart = y
End Sub

Private Sub picH_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sngTemp As Single
    Dim sngBottom As Single
    
    If Button = 1 Then
        sngTemp = picH.Top + y - msngStart
        sngBottom = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
        If sngTemp - lvw����_S.Top > 1500 And sngBottom - sngTemp > 2000 Then
            picH.Top = sngTemp
            lvw����_S.Height = sngTemp - lvw����_S.Top
            
            lblDown.Top = sngTemp + 45
            msh����.Top = lblDown.Top + lblDown.Height
            msh����.Height = sngBottom - msh����.Top
        End If
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "New"
            mnuBillGet_Click
        Case "Modify"
            mnuBillModify_Click
        Case "Delete"
            mnuBillDelete_Click
        Case "Cancel"
            mnuBillCancel_Click
        Case "Check"
            Call mnuBillCheck_Click(1)
        Case "Quit"
            mnuFileExit_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Preview"
            mnuFilePreview_Click
        Case "Filter"
            mnuViewFilter_Click
        Case "Help"
            mnuHelpTopic_Click
        Case "View"
            mnuViewIcon(lvw����_S.View).Checked = False
            If lvw����_S.View = 3 Then
                mnuViewIcon(0).Checked = True
                lvw����_S.View = 0
            Else
                mnuViewIcon(lvw����_S.View + 1).Checked = True
                lvw����_S.View = lvw����_S.View + 1
            End If
    End Select
End Sub

Private Sub Toolbar1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool, vbPopupMenuRightButton
    End If
End Sub

Private Sub Fill��¼()
'����:װ�������շ�Ա��lvw����_S_S
    Dim rsTmp As ADODB.Recordset
    Dim lst As ListItem, intƱ�� As gBillType, strWhere As String
    Dim strKey As String, str��� As String, strʹ����� As String
    Dim lngCol  As Long, strColName As String
    Dim varValue As Variant
        
    If Not lvw����_S.SelectedItem Is Nothing Then
        strKey = lvw����_S.SelectedItem.Key '����ԭ�м�ֵ
    End If
    'mstrOperator:û�����в���ԱȨ��ʱ,ֻ����ʾ��������Ʊ�ݻ���Ʊ��
    '����:35834
    intƱ�� = Val(Mid(lvwMain.SelectedItem.Key, 2))
    strWhere = ""
    Select Case intƱ��
    Case gBillType.�շ��վ�, gBillType.�����վ�
        str��� = cbo���.Text
        If str��� = " " Then
            strWhere = strWhere & " And nvl(ʹ�����,'LXH')=[6]"
            str��� = "LXH"
        ElseIf str��� <> "�������" Then
            strWhere = strWhere & " And nvl(ʹ�����,'LXH')=[6]"
        End If
    Case gBillType.Ԥ���վ�
        If cbo���.ListIndex < 0 Then Exit Sub
        str��� = cbo���.ItemData(cbo���.ListIndex)
        '58071
        If Val(str���) <> -1 Then
            str��� = Val(str���)
            strWhere = strWhere & " And nvl(ʹ�����,'0')=[6]"
        End If
    Case gBillType.���￨
        If cbo���.ListIndex < 0 Then Exit Sub
        str��� = cbo���.ItemData(cbo���.ListIndex)
        If Val(str���) <> 0 Then
            str��� = Val(str���)
            strWhere = strWhere & " And nvl(ʹ�����,'0')=[6]"
        End If
    Case Else
    End Select

    strʹ����� = "A.ʹ�����,"
    If intƱ�� = gBillType.Ԥ���վ� Then
        '58071
        strʹ����� = "decode(nvl(A.ʹ�����,'0'),'0','','1','����','סԺ') as ʹ�����,"
    ElseIf intƱ�� = gBillType.���￨ Then
        strʹ����� = "nvl(M.����,'���￨') As ʹ�����,"
    End If
    
    If intƱ�� = gBillType.���ѿ� Then
        If cbo���.ListIndex < 0 Then Exit Sub
        str��� = cbo���.ItemData(cbo���.ListIndex)
        
        gstrSQL = _
            "Select A.ID, nvl(M.����,'���ѿ�') As ʹ�����,A.������,A.ǰ׺�ı�," & vbNewLine & _
            "       A.��ʼ���� As ��ʼ����,A.��ֹ���� As ��ֹ����," & vbNewLine & _
            "       Decode(A.ʹ�÷�ʽ,1,'����','����') as ʹ�÷�ʽ," & vbNewLine & _
            "       to_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD') as �Ǽ�ʱ��," & vbNewLine & _
            "       A.�Ǽ���,A.��ǰ���� As ��ǰ����,A.ʣ������,A.����,A.�˶���," & vbNewLine & _
            "       A.ǩ����,to_Char(A.ǩ��ʱ��,'YYYY-MM-DD HH24:mi:ss') as ǩ��ʱ��" & vbNewLine & _
            "From " & IIf(mblnDateMoved, zlGetFullFieldsTable("���ѿ����ü�¼"), "���ѿ����ü�¼ A") & _
                " ,��Ա�� B,���ѿ����Ŀ¼ M" & vbNewLine & _
            "Where a.�ӿڱ�� = m.���(+) And a.�ӿڱ��=[6]" & vbNewLine & _
                    IIf(mnuViewHave.Checked, " And A.ʣ������<>0", "") & vbNewLine & _
            "      And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & vbNewLine & _
            "      And A.������=B.���� And A.�Ǽ�ʱ�� Between [2] And [3]" & vbNewLine & _
                    IIf(mstrOperator = "", "", " And (A.������=[4] Or nvl(A.ʹ�÷�ʽ,0)=2)")
    Else
        gstrSQL = _
            "Select A.ID," & strʹ����� & _
            "       A.������,A.ǰ׺�ı�,A.��ʼ����,A.��ֹ����," & vbNewLine & _
            "       Decode(A.ʹ�÷�ʽ,1,'����','����') as ʹ�÷�ʽ," & vbNewLine & _
            "       to_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD') as �Ǽ�ʱ��," & vbNewLine & _
            "       A.�Ǽ���,A.��ǰ����,A.ʣ������,A.����,A.�˶���," & vbNewLine & _
            "       A.ǩ����,to_Char(A.ǩ��ʱ��,'YYYY-MM-DD HH24:mi:ss') as ǩ��ʱ��" & vbNewLine & _
            "From " & IIf(mblnDateMoved, zlGetFullFieldsTable("Ʊ�����ü�¼"), "Ʊ�����ü�¼ A") & _
                " ,��Ա�� B" & IIf(intƱ�� = gBillType.���￨, ",ҽ�ƿ���� M", "") & vbNewLine & _
            "Where A.Ʊ��=[1] " & IIf(mnuViewHave.Checked, " And A.ʣ������<>0", "") & vbNewLine & _
            "      And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & strWhere & vbNewLine & _
            "      And A.������=B.���� And A.�Ǽ�ʱ�� Between [2] And [3]" & vbNewLine & _
                    IIf(mstrOperator = "", "", " And (A.������=[4] Or nvl(A.ʹ�÷�ʽ,0)=2)") & vbNewLine & _
                    IIf(intƱ�� = gBillType.���￨, " And to_number(nvl(A.ʹ�����,'0'))=M.ID(+)", "")
    End If

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
        intƱ��, mdatBegin, DateAdd("s", -1, DateAdd("d", 1, mdatEnd)), _
        mstrOperator, UserInfo.����, str���)
    
    LockWindowUpdate lvw����_S.hWnd
    With lvw����_S
        .ListItems.Clear
        Do Until rsTmp.EOF
            Set lst = .ListItems.Add(, "C" & rsTmp("ID"), rsTmp("��ʼ����"), "Item", "Item")
            
            '����ListView�����������ݿ�ȡ��
            For lngCol = 2 To lvw����_S.ColumnHeaders.Count
                strColName = lvw����_S.ColumnHeaders(lngCol).Text
                If strColName = "��ʼ����" Then strColName = "��ʼ����"
                If strColName = "��ֹ����" Then strColName = "��ֹ����"
                If strColName = "��ǰ����" Then strColName = "��ǰ����"
                varValue = rsTmp(strColName).Value
                lst.SubItems(lngCol - 1) = IIf(IsNull(varValue), "", varValue)
            Next
            rsTmp.MoveNext
        Loop
        If .ListItems.Count > 0 Then
            Dim Item As ListItem
            On Error Resume Next
            Set Item = .ListItems(strKey)
            If Err <> 0 Then
                Set Item = .ListItems(1)
                Item.Selected = True
                Item.EnsureVisible
                lvw����_S_ItemClick Item
            Else
                Err.Clear
                Item.Selected = True
                Item.EnsureVisible
                mstrKey = "" '���״̬����,ˢ�»����б�
                lvw����_S_ItemClick Item
            End If
        Else
            Call Fill����
        End If
    End With
    LockWindowUpdate 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    LockWindowUpdate 0
End Sub

Public Sub Fill����()
'����:��Ʊ��ʹ���������
    Dim rsTmp As ADODB.Recordset
    Dim lngCol As Long
    Dim lngSum(1 To 5) As Long
    Dim lngCSum(1 To 5) As Long, intƱ�� As Integer
    On Error GoTo errH
    
    If lvw����_S.SelectedItem Is Nothing Then
        msh����.Rows = 3
        For lngCol = 0 To msh����.Cols - 1
            msh����.TextMatrix(2, lngCol) = ""
        Next
        msh����.Row = 2
        Call SetMenu
        Exit Sub
    End If
    
    intƱ�� = Val(Mid(lvwMain.SelectedItem.Key, 2))
    If intƱ�� = gBillType.���ѿ� Then
        If mnuViewCheck.Checked Then
            gstrSQL = _
                "Select ʹ����, Sum(Decode(ԭ��, 1, 1, 6, 1, 0)) As ����," & vbNewLine & _
                "       Sum(Decode(ԭ��, 2, 1, 0)) As ����," & vbNewLine & _
                "       Sum(Decode(ԭ��, 3, 1, 0)) As �ش�," & vbNewLine & _
                "       Sum(Decode(ԭ��, 4, 1, 0)) As �ش��ջ�," & vbNewLine & _
                "       Sum(Decode(ԭ��, 5, 1, 0)) As ����," & vbNewLine & _
                "       Sum(Decode(�˶Խ��, 1, 1, 0)) As C����," & vbNewLine & _
                "       Sum(Decode(�˶Խ��, 2, 1, 0)) As C����," & vbNewLine & _
                "       Sum(Decode(�˶Խ��, 3, 1, 0)) As C�ش�," & vbNewLine & _
                "       Sum(Decode(�˶Խ��, 4, 1, 0)) As C�ش��ջ�," & vbNewLine & _
                "       Sum(Decode(�˶Խ��, 5, 1, 0)) As C����" & vbNewLine & _
                "From " & IIf(mblnDateMoved, zlGetFullFieldsTable("���ѿ�ʹ�ü�¼"), "���ѿ�ʹ�ü�¼") & vbNewLine & _
                "Where ����id = [1]" & vbNewLine & _
                "Group By ʹ����"
        Else
            gstrSQL = _
                "Select ʹ����,Sum(Decode(ԭ��, 1, 1, 6, 1, 0)) As ����," & vbNewLine & _
                "       Sum(Decode(ԭ��, 2, 1, 0)) As ����, " & vbNewLine & _
                "       Sum(decode(ԭ��,3,1,0)) As �ش�," & vbNewLine & _
                "       Sum(decode(ԭ��,4,1,0)) As �ش��ջ�," & vbNewLine & _
                "       Sum(decode(ԭ��,5,1,0)) As ���� " & vbNewLine & _
                "From " & IIf(mblnDateMoved, zlGetFullFieldsTable("���ѿ�ʹ�ü�¼"), "���ѿ�ʹ�ü�¼") & vbNewLine & _
                "Where ����ID = [1]" & vbNewLine & _
                "Group By ʹ����"
        End If
    Else
        If mnuViewCheck.Checked Then
            gstrSQL = _
                "Select ʹ����, Sum(Decode(ԭ��, 1, 1, 6, 1, 0)) As ����," & vbNewLine & _
                "       Sum(Decode(ԭ��, 2, 1, 0)) As ����," & vbNewLine & _
                "       Sum(Decode(ԭ��, 3, 1, 0)) As �ش�," & vbNewLine & _
                "       Sum(Decode(ԭ��, 4, 1, 0)) As �ش��ջ�," & vbNewLine & _
                "       Sum(Decode(ԭ��, 5, 1, 0)) As ����," & vbNewLine & _
                "       Sum(Decode(�˶Խ��, 1, 1, 0)) As C����," & vbNewLine & _
                "       Sum(Decode(�˶Խ��, 2, 1, 0)) As C����," & vbNewLine & _
                "       Sum(Decode(�˶Խ��, 3, 1, 0)) As C�ش�," & vbNewLine & _
                "       Sum(Decode(�˶Խ��, 4, 1, 0)) As C�ش��ջ�," & vbNewLine & _
                "       Sum(Decode(�˶Խ��, 5, 1, 0)) As C����" & vbNewLine & _
                "From " & IIf(mblnDateMoved, zlGetFullFieldsTable("Ʊ��ʹ����ϸ"), "Ʊ��ʹ����ϸ") & vbNewLine & _
                "Where ����id = [1]" & vbNewLine & _
                "Group By ʹ����"
        Else
            gstrSQL = _
                "Select ʹ����,Sum(Decode(ԭ��, 1, 1, 6, 1, 0)) As ����," & vbNewLine & _
                "       Sum(Decode(ԭ��, 2, 1, 0)) As ����, " & vbNewLine & _
                "       Sum(decode(ԭ��,3,1,0)) As �ش�," & vbNewLine & _
                "       Sum(decode(ԭ��,4,1,0)) As �ش��ջ�," & vbNewLine & _
                "       Sum(decode(ԭ��,5,1,0)) As ���� " & vbNewLine & _
                "From " & IIf(mblnDateMoved, zlGetFullFieldsTable("Ʊ��ʹ����ϸ"), "Ʊ��ʹ����ϸ") & vbNewLine & _
                "Where ����ID = [1]" & vbNewLine & _
                "Group By ʹ����"
        End If
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Mid(lvw����_S.SelectedItem.Key, 2))
    With msh����
        .Redraw = False
        If rsTmp.EOF Then
            .Rows = 3
            For lngCol = 0 To .Cols - 1
                .TextMatrix(2, lngCol) = ""
            Next
        Else
            .Rows = rsTmp.RecordCount + 3 '�����Ǳ�ͷ����һ���Ǻϼ�
            lngCol = 2
            Do Until rsTmp.EOF
                .TextMatrix(lngCol, 0) = rsTmp!ʹ����
                If mnuViewCheck.Checked Then
                    .TextMatrix(lngCol, 1) = Format(rsTmp!����, "#########;-#########;" & IIf(rsTmp!���� = 0 And rsTmp!C���� <> 0, "0", "") & "; ") & Format(rsTmp!C����, "\/#########;-#########; ; "): lngSum(1) = lngSum(1) + rsTmp!����: lngCSum(1) = lngCSum(1) + rsTmp!C����
                    .TextMatrix(lngCol, 2) = Format(rsTmp!�ش�, "#########;-#########;" & IIf(rsTmp!�ش� = 0 And rsTmp!C�ش� <> 0, "0", "") & "; ") & Format(rsTmp!C�ش�, "\/#########;-#########; ; "): lngSum(2) = lngSum(2) + rsTmp!�ش�: lngCSum(2) = lngCSum(2) + rsTmp!C�ش�
                    .TextMatrix(lngCol, 3) = Format(rsTmp!����, "#########;-#########;" & IIf(rsTmp!���� = 0 And rsTmp!C���� <> 0, "0", "") & "; ") & Format(rsTmp!C����, "\/#########;-#########; ; "): lngSum(3) = lngSum(3) + rsTmp!����: lngCSum(3) = lngCSum(3) + rsTmp!C����
                    .TextMatrix(lngCol, 4) = Format(rsTmp!���� + rsTmp!�ش� + rsTmp!����, "#########;-#########;" & IIf((rsTmp!���� + rsTmp!�ش� + rsTmp!����) = 0 And (rsTmp!C���� + rsTmp!C�ش� + rsTmp!C����) <> 0, "0", "") & "; ") & Format(rsTmp!C���� + rsTmp!C�ش� + rsTmp!C����, "\/#########;-#########; ; ")
                    .TextMatrix(lngCol, 5) = Format(rsTmp!����, "#########;-#########;" & IIf(rsTmp!���� = 0 And rsTmp!C���� <> 0, "0", "") & "; ") & Format(rsTmp!C����, "\/#########;-#########; ; "): lngSum(4) = lngSum(4) + rsTmp!����: lngCSum(4) = lngCSum(4) + rsTmp!C����
                    .TextMatrix(lngCol, 6) = Format(rsTmp!�ش��ջ�, "#########;-#########;" & IIf(rsTmp!�ش��ջ� = 0 And rsTmp!C�ش��ջ� <> 0, "0", "") & "; ") & Format(rsTmp!C�ش��ջ�, "\/#########;-#########; ; "): lngSum(5) = lngSum(5) + rsTmp!�ش��ջ�: lngCSum(5) = lngCSum(5) + rsTmp!C�ش��ջ�
                    .TextMatrix(lngCol, 7) = Format(rsTmp!���� + rsTmp!�ش��ջ�, "#########;-#########;" & IIf((rsTmp!���� + rsTmp!�ش��ջ�) = 0 And (rsTmp!C���� + rsTmp!C�ش��ջ�) <> 0, "0", "") & "; ") & Format(rsTmp!C���� + rsTmp!C�ش��ջ�, "\/#########;-#########; ; ")

                Else
                    .TextMatrix(lngCol, 1) = Format(rsTmp!����, "#########;-#########; ; "): lngSum(1) = lngSum(1) + rsTmp!����
                    .TextMatrix(lngCol, 2) = Format(rsTmp!�ش�, "#########;-#########; ; "): lngSum(2) = lngSum(2) + rsTmp!�ش�
                    .TextMatrix(lngCol, 3) = Format(rsTmp!����, "#########;-#########; ; "): lngSum(3) = lngSum(3) + rsTmp!����
                    .TextMatrix(lngCol, 4) = Format(rsTmp!���� + rsTmp!�ش� + rsTmp!����, "#########;-#########; ; ")
                    .TextMatrix(lngCol, 5) = Format(rsTmp!����, "#########;-#########; ; "): lngSum(4) = lngSum(4) + rsTmp!����
                    .TextMatrix(lngCol, 6) = Format(rsTmp!�ش��ջ�, "#########;-#########; ; "): lngSum(5) = lngSum(5) + rsTmp!�ش��ջ�
                    .TextMatrix(lngCol, 7) = Format(rsTmp!���� + rsTmp!�ش��ջ�, "#########;-#########; ; ")
                End If
                
                lngCol = lngCol + 1
                rsTmp.MoveNext
            Loop
            lngCol = .Rows - 1
            .TextMatrix(lngCol, 0) = "   �ϼ�"
            If mnuViewCheck.Checked Then
                .TextMatrix(lngCol, 1) = Format(lngSum(1), "#########;-#########;" & IIf(lngSum(1) = 0 And lngCSum(1) <> 0, "0", "") & "; ") & Format(lngCSum(1), "\/#########;-#########; ; ")
                .TextMatrix(lngCol, 2) = Format(lngSum(2), "#########;-#########;" & IIf(lngSum(2) = 0 And lngCSum(2) <> 0, "0", "") & "; ") & Format(lngCSum(2), "\/#########;-#########; ; ")
                .TextMatrix(lngCol, 3) = Format(lngSum(3), "#########;-#########;" & IIf(lngSum(3) = 0 And lngCSum(3) <> 0, "0", "") & "; ") & Format(lngCSum(3), "\/#########;-#########; ; ")
                .TextMatrix(lngCol, 4) = Format(lngSum(1) + lngSum(2) + lngSum(3), "#########;-#########;" & IIf((lngSum(1) + lngSum(2) + lngSum(3)) = 0 And (lngCSum(1) + lngCSum(2) + lngCSum(3)) <> 0, "0", "") & "; ") & Format(lngCSum(1) + lngCSum(2) + lngCSum(3), "\/#########;-#########; ; ")
                .TextMatrix(lngCol, 5) = Format(lngSum(4), "#########;-#########;" & IIf(lngSum(4) = 0 And lngCSum(4) <> 0, "0", "") & "; ") & Format(lngCSum(4), "\/#########;-#########; ; ")
                .TextMatrix(lngCol, 6) = Format(lngSum(5), "#########;-#########;" & IIf(lngSum(5) = 0 And lngCSum(5) <> 0, "0", "") & "; ") & Format(lngCSum(5), "\/#########;-#########; ; ")
                .TextMatrix(lngCol, 7) = Format(lngSum(5) + lngSum(4), "#########;-#########;" & IIf((lngSum(5) + lngSum(4)) = 0 And (lngCSum(5) + lngCSum(4)) <> 0, "0", "") & "; ") & Format(lngCSum(5) + lngCSum(4), "\/#########;-#########; ; ")
            
            Else
                .TextMatrix(lngCol, 1) = Format(lngSum(1), "#########;-#########; ; ")
                .TextMatrix(lngCol, 2) = Format(lngSum(2), "#########;-#########; ; ")
                .TextMatrix(lngCol, 3) = Format(lngSum(3), "#########;-#########; ; ")
                .TextMatrix(lngCol, 4) = Format(lngSum(1) + lngSum(2) + lngSum(3), "#########;-#########; ; ")
                .TextMatrix(lngCol, 5) = Format(lngSum(4), "#########;-#########; ; ")
                .TextMatrix(lngCol, 6) = Format(lngSum(5), "#########;-#########; ; ")
                .TextMatrix(lngCol, 7) = Format(lngSum(5) + lngSum(4), "#########;-#########; ; ")
            End If
        End If
        .Redraw = True
        .Row = 2
    End With
    Call SetMenu
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
    Dim objPrint As Object
    Dim objRow As New zlTabAppRow
    
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    
    If ActiveControl Is msh���� Then
        If msh����.Rows = 3 Then Exit Sub
        
        Set objPrint = New zlPrint1Grd
        objPrint.Title.Text = lvwMain.SelectedItem.Text & "ʹ�����"
        Set objPrint.Body = msh����
        objRow.Add "�����ˣ�" & lvw����_S.SelectedItem.SubItems(GetItemCOL("������"))
        If CurrentIsBill(Val(Mid(lvwMain.SelectedItem.Key, 2))) Then
            objRow.Add "���룺" & lvw����_S.SelectedItem.Text & _
                "����" & lvw����_S.SelectedItem.SubItems(GetItemCOL("��ֹ����"))
        Else
            objRow.Add "���ţ�" & lvw����_S.SelectedItem.Text & _
                "����" & lvw����_S.SelectedItem.SubItems(GetItemCOL("��ֹ����"))
        End If
        objPrint.UnderAppRows.Add objRow
        
        Set objRow = New zlTabAppRow
        objRow.Add "��ӡ�ˣ�" & UserInfo.����
        objRow.Add "��ӡʱ�䣺" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
        objPrint.BelowAppRows.Add objRow
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
    
    Else
        Set objPrint = New zlPrintLvw
        objPrint.Title.Text = lvwMain.SelectedItem.Text & "���ü�¼"
        Set objPrint.Body.objData = lvw����_S
        objPrint.UnderAppItems.Add "����ʱ�䣺" & Format(mdatBegin, "yyyy��MM��dd��") & _
            "����" & Format(mdatEnd, "yyyy��MM��dd��")
        objPrint.BelowAppItems.Add "��ӡ�ˣ�" & UserInfo.����
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
    End If
End Sub

Private Sub PrivilegeCTRL()
'����:�����е��û�Ȩ�޲���,��ʹһЩ�˵����ť���ɼ�
    If InStr(mstrPrivs, "��ɾ��") = 0 _
        And InStr(mstrPrivs, "Ʊ�ݱ���") = 0 _
        And InStr(mstrPrivs, "Ʊ�ݺ˶�") = 0 Then
        mnuBill.Visible = False
        Toolbar1.Buttons("New").Visible = False
        Toolbar1.Buttons("Modify").Visible = False
        Toolbar1.Buttons("Delete").Visible = False
        Toolbar1.Buttons("Split1").Visible = False
        Toolbar1.Buttons("Cancel").Visible = False
        Toolbar1.Buttons("Check").Visible = False
        Toolbar1.Buttons("Split2").Visible = False
    ElseIf InStr(mstrPrivs, "��ɾ��") = 0 Then
        mnuBillGet.Visible = False
        mnuBillModify.Visible = False
        mnuBillDelete.Visible = False
        mnuBillSplit.Visible = False
        Toolbar1.Buttons("New").Visible = False
        Toolbar1.Buttons("Modify").Visible = False
        Toolbar1.Buttons("Delete").Visible = False
        Toolbar1.Buttons("Split1").Visible = False
    ElseIf InStr(mstrPrivs, "Ʊ�ݱ���") = 0 _
        And InStr(mstrPrivs, "Ʊ�ݺ˶�") = 0 Then
        mnuBillSplit.Visible = False
        mnuBillCancel.Visible = False
        mnuBillCheck(0).Visible = False
        mnuBillCheck(1).Visible = False
        Toolbar1.Buttons("Split2").Visible = False
        Toolbar1.Buttons("Cancel").Visible = False
        Toolbar1.Buttons("Check").Visible = False
    ElseIf InStr(mstrPrivs, "Ʊ�ݱ���") = 0 Then
        mnuBillCancel.Visible = False
        Toolbar1.Buttons("Cancel").Visible = False
    ElseIf InStr(mstrPrivs, "Ʊ�ݺ˶�") = 0 Then
        mnuBillCheck(0).Visible = False
        mnuBillCheck(1).Visible = False
        Toolbar1.Buttons("Check").Visible = False
    End If
    mblnҩ�� = (glngSys \ 100 = 8)
End Sub

Private Sub SetMenu()
    Dim blnDetail As Boolean, blnModify As Boolean, blnChecked As Boolean
    Dim blnHavePrivs As Boolean  '�Ƿ��в���Ȩ��
    
    blnModify = Not (lvw����_S.SelectedItem Is Nothing)
    
    blnHavePrivs = zlStr.IsHavePrivs(mstrPrivs, "����������˵Ǽ�Ʊ��")
    If blnHavePrivs = False And blnModify Then
        '����жϵǼ���:
       blnHavePrivs = lvw����_S.SelectedItem.SubItems(GetItemCOL("�Ǽ���")) = UserInfo.����
    End If
    
    blnDetail = (msh����.Rows > 3)
    
    If Not (lvw����_S.SelectedItem Is Nothing) Then
        blnChecked = (lvw����_S.SelectedItem.SubItems(GetItemCOL("�˶���")) <> "")
    End If

    mnuBillModify.Enabled = blnModify And blnHavePrivs
    mnuBillCancel.Enabled = blnModify
    
    mnuBillCheck(0).Enabled = blnModify And Not blnDetail    '�˶����õ�
    If blnChecked Then
        mnuBillCheck(0).Caption = "ȡ���˶����õ�(&B)"
    Else
        mnuBillCheck(0).Caption = "�˶����õ�(&B)"
    End If
    Toolbar1.Buttons("Modify").Enabled = blnModify And blnHavePrivs
    Toolbar1.Buttons("Cancel").Enabled = blnModify
    Toolbar1.Buttons("Check").Enabled = blnModify And blnDetail
    
    
    mnuBillDelete.Enabled = blnModify And Not blnDetail And blnHavePrivs
    Toolbar1.Buttons("Delete").Enabled = blnModify And Not blnDetail And blnHavePrivs
    mnuViewDetail.Enabled = blnDetail
    

    mnuFilePreview.Enabled = blnModify
    mnuFilePrint.Enabled = blnModify
    mnuFileExcel.Enabled = blnModify
    Toolbar1.Buttons("Preview").Enabled = blnModify
    Toolbar1.Buttons("Print").Enabled = blnModify
    
    stbThis.Panels(2).Text = "��" & lvwMain.SelectedItem.Text & "����" & _
        Format(mdatBegin, "yyyy��MM��dd��") & "����" & _
        Format(mdatEnd, "yyyy��MM��dd��") & "֮�乲��" & lvw����_S.ListItems.Count & "�����ü�¼��"
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

Private Function LoadCombox() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Combox����
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-04-27 10:22:29
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intƱ�� As gBillType, str��� As String
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    If lvwMain.SelectedItem Is Nothing Then Exit Function
    
    intƱ�� = Val(Mid(lvwMain.SelectedItem.Key, 2))
    str��� = lvwMain.SelectedItem.Tag
    
    Select Case intƱ��
    Case gBillType.�շ��վ�, gBillType.�����վ�
        strSQL = "Select ����,����,����,ȱʡ��־ From Ʊ��ʹ����� "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        mblnNotClick = True
        With cbo���
            .Clear
            
            .AddItem "�������"
            If str��� = "�������" Then .ListIndex = .NewIndex
            
            Do While Not rsTemp.EOF
                .AddItem Nvl(rsTemp!����)
                .ItemData(.NewIndex) = 1
                If Val(Nvl(rsTemp!ȱʡ��־)) = 1 And .ListIndex < 0 Then .ListIndex = .NewIndex
                If str��� = Nvl(rsTemp!����) Then .ListIndex = .NewIndex
                rsTemp.MoveNext
            Loop
            
            .AddItem " "
            .ItemData(.NewIndex) = -1
            If str��� = " " Then .ListIndex = .NewIndex
            
            If .ListIndex < 0 And .ListCount > 0 Then .ListIndex = 0
        End With
        cbrThis.Bands(2).Visible = True
        cbrThis.Bands(2).Caption = "ʹ�����"
        mblnNotClick = False
    Case gBillType.Ԥ���վ�
        mblnNotClick = True
        With cbo���
            .Clear
            If zlStr.IsHavePrivs(mstrPrivs, "Ԥ������Ʊ��") _
                And zlStr.IsHavePrivs(mstrPrivs, "Ԥ��סԺƱ��") Then
                .AddItem "����Ԥ��"
                .ItemData(.NewIndex) = -1
                If Val(str���) = -1 Then .ListIndex = .NewIndex
            End If
            If zlStr.IsHavePrivs(mstrPrivs, "Ԥ������Ʊ��") Then
                .AddItem "����Ԥ��"
                .ItemData(.NewIndex) = 1
                If Val(str���) = 1 Then .ListIndex = .NewIndex
            End If
            If zlStr.IsHavePrivs(mstrPrivs, "Ԥ��סԺƱ��") Then
                .AddItem "סԺԤ��"
                .ItemData(.NewIndex) = 2
                If Val(str���) = 2 Then .ListIndex = .NewIndex
            End If
            '58071
            If zlStr.IsHavePrivs(mstrPrivs, "Ԥ������Ʊ��") _
                And zlStr.IsHavePrivs(mstrPrivs, "Ԥ��סԺƱ��") Then
                .AddItem ""
                .ItemData(.NewIndex) = 0
                If Val(str���) = 0 Then .ListIndex = .NewIndex
            End If
            If .ListIndex < 0 And .ListCount > 0 Then .ListIndex = 0
        End With
        cbrThis.Bands(2).Visible = True
        cbrThis.Bands(2).Caption = "ʹ�����"
        mblnNotClick = False
    Case gBillType.���￨
        strSQL = _
            "Select ID, ����, ����, ȱʡ��־" & vbNewLine & _
            "From ҽ�ƿ����" & vbNewLine & _
            "Where Nvl(�Ƿ�����, 0) >= 1" & vbNewLine & _
            "Order By ����"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        mblnNotClick = True
        With cbo���
            .Clear
            Do While Not rsTemp.EOF
                .AddItem Nvl(rsTemp!����) & "-" & Nvl(rsTemp!����)
                .ItemData(.NewIndex) = Val(Nvl(rsTemp!ID))
                If Val(Nvl(rsTemp!ȱʡ��־)) = 1 And .ListIndex < 0 Then .ListIndex = .NewIndex
                If Val(str���) = Val(Nvl(rsTemp!ID)) Then .ListIndex = .NewIndex
                rsTemp.MoveNext
            Loop
            If .ListIndex < 0 And .ListCount > 0 Then .ListIndex = 0
        End With
        cbrThis.Bands(2).Visible = True
        cbrThis.Bands(2).Caption = "�����"
        mblnNotClick = False
    Case gBillType.���ѿ�
        strSQL = "Select ���, ���� From ���ѿ����Ŀ¼ Where Nvl(����, 0) >= 1 Order By ���"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        mblnNotClick = True
        With cbo���
            .Clear
            Do While Not rsTemp.EOF
                .AddItem Nvl(rsTemp!���) & "-" & Nvl(rsTemp!����)
                .ItemData(.NewIndex) = Val(Nvl(rsTemp!���))
                If Val(str���) = Val(Nvl(rsTemp!���)) Then .ListIndex = .NewIndex
                rsTemp.MoveNext
            Loop
            If .ListIndex < 0 And .ListCount > 0 Then .ListIndex = 0
        End With
        cbrThis.Bands(2).Visible = True
        cbrThis.Bands(2).Caption = "�����"
        mblnNotClick = False
    Case Else
        cbrThis.Bands(2).Visible = False
    End Select
    LoadCombox = True
     
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetDefaultUserType()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ȱʡ��ʹ�����
    '����:���˺�
    '����:2011-04-27 14:23:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intƱ�� As gBillType
    
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    intƱ�� = Val(Mid(lvwMain.SelectedItem.Key, 2))
    Select Case intƱ��
    Case gBillType.�շ��վ�, gBillType.�����վ�
        lvwMain.SelectedItem.Tag = cbo���.Text
    Case gBillType.Ԥ���վ�, gBillType.���￨, gBillType.���ѿ�
        If cbo���.ListIndex >= 0 Then
            lvwMain.SelectedItem.Tag = cbo���.ItemData(cbo���.ListIndex)
        Else
            lvwMain.SelectedItem.Tag = ""
        End If
    Case Else
    End Select
End Sub
