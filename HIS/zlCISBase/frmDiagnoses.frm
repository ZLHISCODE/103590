VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form frmDiagnoses 
   BackColor       =   &H8000000C&
   Caption         =   "������ϲο�"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9795
   Icon            =   "frmDiagnoses.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   9795
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picHBar 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   30
      Left            =   2685
      MousePointer    =   7  'Size N S
      ScaleHeight     =   30
      ScaleWidth      =   6075
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2955
      Width           =   6075
   End
   Begin VB.PictureBox picVBar 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5235
      Left            =   2580
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5235
      ScaleWidth      =   30
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   795
      Width           =   30
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   1905
      Top             =   6765
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
            Picture         =   "frmDiagnoses.frx":0442
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoses.frx":09DC
            Key             =   "expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoses.frx":0F76
            Key             =   "item"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoses.frx":13C8
            Key             =   "itemstop"
         EndProperty
      EndProperty
   End
   Begin SysInfoLib.SysInfo SysInfo 
      Left            =   8940
      Top             =   6870
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvwList 
      Height          =   1710
      Left            =   2835
      TabIndex        =   9
      Top             =   930
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   3016
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   7590
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDiagnoses.frx":1B42
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12197
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
   Begin ComCtl3.CoolBar clbThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   9795
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tlbThis"
      MinWidth1       =   24000
      MinHeight1      =   720
      Width1          =   8730
      FixedBackground1=   0   'False
      Key1            =   "Comm"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tlbThis 
         Height          =   720
         Left            =   30
         TabIndex        =   12
         Top             =   30
         Width           =   24000
         _ExtentX        =   42333
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
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
               Description     =   "Ԥ��"
               Object.ToolTipText     =   "Ԥ����ǰ��"
               Object.Tag             =   "Ԥ��"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Description     =   "��ӡ"
               Object.ToolTipText     =   "��ӡ��ǰ��"
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
               Object.ToolTipText     =   "�·���"
               Object.Tag             =   "����"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Add"
               Description     =   "����"
               Object.ToolTipText     =   "�¼������"
               Object.Tag             =   "����"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "Mod"
               Description     =   "�޸�"
               Object.ToolTipText     =   "�޸ļ������"
               Object.Tag             =   "�޸�"
               ImageKey        =   "Modify"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "Del"
               Description     =   "ɾ��"
               Object.ToolTipText     =   "ɾ���������"
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
               Object.Tag             =   "����"
               ImageKey        =   "Start"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ͣ��"
               Key             =   "Stop"
               Object.Tag             =   "ͣ��"
               ImageKey        =   "Stop"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Find"
               Description     =   "����"
               Object.ToolTipText     =   "���������Ŀ"
               Object.Tag             =   "����"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "����"
               Object.ToolTipText     =   "��ǰ��������"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Exit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   7680
      Top             =   525
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
            Picture         =   "frmDiagnoses.frx":23D4
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoses.frx":25EE
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoses.frx":2808
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoses.frx":2A22
            Key             =   "New"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoses.frx":2C3C
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoses.frx":2E5C
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoses.frx":307C
            Key             =   "View"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoses.frx":329C
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoses.frx":34B6
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoses.frx":36D6
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoses.frx":38F6
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoses.frx":3B10
            Key             =   "Stop"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   6915
      Top             =   525
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
            Picture         =   "frmDiagnoses.frx":3D2A
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoses.frx":3F4A
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoses.frx":416A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoses.frx":4384
            Key             =   "New"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoses.frx":459E
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoses.frx":47BE
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoses.frx":49DE
            Key             =   "View"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoses.frx":4BFE
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoses.frx":4E18
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoses.frx":5038
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoses.frx":5258
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoses.frx":5472
            Key             =   "Stop"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picClass 
      Height          =   6270
      Left            =   0
      ScaleHeight     =   6210
      ScaleWidth      =   2340
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   720
      Width           =   2400
      Begin VB.CommandButton cmdKind 
         Caption         =   "��ҽ��ϲο�(&2)"
         Height          =   300
         Index           =   1
         Left            =   0
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   300
         Width           =   2295
      End
      Begin VB.CommandButton cmdKind 
         Caption         =   "��ҽ��ϲο�(&1)"
         Height          =   300
         Index           =   0
         Left            =   0
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   15
         Width           =   2295
      End
      Begin MSComctlLib.TreeView tvwClass 
         Height          =   4800
         Left            =   45
         TabIndex        =   8
         Tag             =   "1000"
         Top             =   645
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   8467
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   353
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "imgList"
         Appearance      =   0
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgdCodex 
      Height          =   2880
      Left            =   2925
      TabIndex        =   11
      Top             =   4170
      Visible         =   0   'False
      Width           =   5850
      _ExtentX        =   10319
      _ExtentY        =   5080
      _Version        =   393216
      BackColor       =   -2147483628
      Rows            =   5
      Cols            =   4
      FixedRows       =   2
      BackColorBkg    =   -2147483628
      GridColor       =   -2147483628
      GridColorFixed  =   16777215
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      GridLines       =   0
      GridLinesFixed  =   0
      ScrollBars      =   2
      MergeCells      =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgdRefer 
      Height          =   2880
      Left            =   2790
      TabIndex        =   10
      Top             =   3420
      Width           =   5850
      _ExtentX        =   10319
      _ExtentY        =   5080
      _Version        =   393216
      BackColor       =   -2147483628
      Rows            =   5
      Cols            =   4
      FixedRows       =   2
      BackColorBkg    =   -2147483628
      GridColor       =   -2147483628
      GridColorFixed  =   16777215
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      GridLines       =   0
      GridLinesFixed  =   0
      ScrollBars      =   2
      MergeCells      =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin MSComctlLib.TabStrip tabContent 
      Height          =   3390
      Left            =   2715
      TabIndex        =   6
      Top             =   3045
      Width           =   6060
      _ExtentX        =   10689
      _ExtentY        =   5980
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�������Ʋο�(&R)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "������������(&C)"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblScale 
      AutoSize        =   -1  'True
      Caption         =   "�����ߴ�"
      Height          =   180
      Left            =   3015
      TabIndex        =   13
      Top             =   7080
      Visible         =   0   'False
      Width           =   1185
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFilePrintset 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFilePreview 
         Caption         =   "Ԥ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFileSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOut 
         Caption         =   "��ӡ�ο�(&O)"
      End
      Begin VB.Menu mnuFileSpt2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditNew 
         Caption         =   "�����(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuEditAdd 
         Caption         =   "�¼���(&A)"
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
      Begin VB.Menu mnuEditSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditStart 
         Caption         =   "����(&S)"
      End
      Begin VB.Menu mnuEditStop 
         Caption         =   "ͣ��(&T)"
      End
      Begin VB.Menu mnuEditSpt2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditRefer 
         Caption         =   "���Ʋο�(&R)"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuEditCodex 
         Caption         =   "��������(&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuEditDiagAndSymptom 
         Caption         =   "��ϲ���(&B)"
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
      Begin VB.Menu mnuToolBar 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuViewToolbarStand 
            Caption         =   "��׼��ť(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolbarText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStates 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "����(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewSpt2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewStop 
         Caption         =   "��ʾͣ����Ŀ(&P)"
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
         WindowList      =   -1  'True
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
      Begin VB.Menu mnuHelpSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frmDiagnoses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngMode As Long
Public mstrPrivs As String       '�û����б�����ľ���Ȩ��

Dim rsTemp As New ADODB.Recordset
Dim objNode As Node
Dim objItem As ListItem
Dim intCount As Integer
Dim strTemp As String

Private Sub cmdKind_Click(Index As Integer)
    Dim intCount As Integer
    For intCount = Me.cmdKind.LBound To Me.cmdKind.UBound
        If intCount <= Index Then
            Me.cmdKind(intCount).Tag = 0
        Else
            Me.cmdKind(intCount).Tag = 1
        End If
    Next
    
    'Ȩ�޿���
    If Index = 0 Then
        Me.mnuEdit.Visible = (InStr(1, mstrPrivs, "��ҽ����") > 0)
        Me.tlbThis.Buttons("Split").Visible = (InStr(1, mstrPrivs, "��ҽ����") > 0)
        Me.tlbThis.Buttons("Split1").Visible = (InStr(1, mstrPrivs, "��ҽ����") > 0)
        Me.tlbThis.Buttons("New").Visible = (InStr(1, mstrPrivs, "��ҽ����") > 0)
        Me.tlbThis.Buttons("Add").Visible = (InStr(1, mstrPrivs, "��ҽ����") > 0)
        Me.tlbThis.Buttons("Mod").Visible = (InStr(1, mstrPrivs, "��ҽ����") > 0)
        Me.tlbThis.Buttons("Del").Visible = (InStr(1, mstrPrivs, "��ҽ����") > 0)
        Me.tlbThis.Buttons("Start").Visible = (InStr(1, mstrPrivs, "��ҽ����") > 0)
        Me.tlbThis.Buttons("Stop").Visible = (InStr(1, mstrPrivs, "��ҽ����") > 0)
    Else
        Me.mnuEdit.Visible = (InStr(1, mstrPrivs, "��ҽ����") > 0)
        Me.tlbThis.Buttons("Split").Visible = (InStr(1, mstrPrivs, "��ҽ����") > 0)
        Me.tlbThis.Buttons("Split1").Visible = (InStr(1, mstrPrivs, "��ҽ����") > 0)
        Me.tlbThis.Buttons("New").Visible = (InStr(1, mstrPrivs, "��ҽ����") > 0)
        Me.tlbThis.Buttons("Add").Visible = (InStr(1, mstrPrivs, "��ҽ����") > 0)
        Me.tlbThis.Buttons("Mod").Visible = (InStr(1, mstrPrivs, "��ҽ����") > 0)
        Me.tlbThis.Buttons("Del").Visible = (InStr(1, mstrPrivs, "��ҽ����") > 0)
        Me.tlbThis.Buttons("Start").Visible = (InStr(1, mstrPrivs, "��ҽ����") > 0)
        Me.tlbThis.Buttons("Stop").Visible = (InStr(1, mstrPrivs, "��ҽ����") > 0)
    End If
    
    'װ���ݲ���������
    If Me.lvwList.Visible Then
        Call picClass_Resize
        Me.tvwClass.SetFocus
    End If
    If Val(tvwClass.Tag) <> Index Then
        Me.tvwClass.Tag = Index
        Call zlRefClasses
    End If
End Sub

Private Sub clbThis_Resize()
    Me.clbThis.Bands(1).MinHeight = Me.tlbThis.Height
    Me.clbThis.Refresh
    Call Form_Resize
End Sub

Private Sub Form_Activate()
    Me.lvwList.Visible = True
End Sub

Private Sub Form_Load()
    mlngMode = glngModul
    mstrPrivs = gstrPrivs
    
    '---------�ؼ�������̬---------
    With Me.lvwList.ColumnHeaders
        .Clear
        .Add , "_����", "����", 2500
        .Add , "_����", "����", 1100
        .Add , "_˵��", "˵��", 4000
        .Add , "_����", "����", 900
        .Add , "_����ʱ��", "����ʱ��", 1400
        .Add , "_����ʱ��", "����ʱ��", 1400
    End With
    With Me.lvwList
        .ColumnHeaders("_����").Position = 1
        .SortKey = .ColumnHeaders("_����").Index - 1
        .SortOrder = lvwAscending
    End With
    
    Call RestoreWinState(Me, App.ProductName)
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    
    If GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser, "ʹ�ø��Ի����", "1") = "1" Then
        strTemp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\�ָ�", "����", "0")
        If strTemp <> "0" Then
            Me.picVBar.Left = CLng(strTemp)
        End If
        strTemp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\�ָ�", "����", "0")
        If strTemp <> "0" Then
            Me.picHBar.Top = CLng(strTemp)
        End If
    End If
    
    '---------Ȩ�޿���-------------
    If InStr(1, mstrPrivs, "�ο��༭") = 0 Then
        Me.mnuEditRefer.Enabled = False
    End If
    If InStr(1, mstrPrivs, "��Ϲ���") = 0 Then
        Me.mnuEditCodex.Enabled = False
    End If
    
    With Me.hgdRefer
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
    End With
    
    With Me.hgdCodex
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
    End With
    
    mnuViewStop.Checked = (Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��ʾͣ��", 0)) = 1)
    
    Call cmdKind_Click(0)
    
End Sub

Private Sub Form_Resize()
    Dim lngTools As Single, lngStatus As Single
    
    If WindowState = 1 Then Exit Sub
    lngTools = IIf(Me.clbThis.Visible, Me.clbThis.Height, 0)
    lngStatus = IIf(Me.stbThis.Visible, Me.stbThis.Height, 0)
    
    Err = 0: On Error Resume Next
    
    With Me.picVBar
        .Top = lngTools
        .Height = Me.ScaleHeight - picClass.Top - lngStatus
        If .Left < 2000 Then .Left = 2000
        If .Left > Me.ScaleWidth - 4000 Then .Left = Me.ScaleWidth - 4000
    End With
    With Me.picHBar
        .Left = Me.picVBar.Left + Me.picVBar.Width
        .Width = Me.ScaleWidth - .Left
        If .Top < 2000 Then .Top = 2000
        If .Top > Me.ScaleHeight - lngStatus - 3000 Then .Top = Me.ScaleHeight - lngStatus - 3000
    End With
    With Me.picClass
        .Left = Me.ScaleLeft
        .Top = lngTools
        .Height = Me.ScaleHeight - picClass.Top - lngStatus
        .Width = Me.picVBar.Left - Me.picClass.Left
    End With
    
    With Me.lvwList
        .Left = Me.picVBar.Left + Me.picVBar.Width
        .Top = lngTools
        .Height = Me.picHBar.Top - .Top
        .Width = Me.ScaleWidth - .Left
    End With
    
    With Me.tabContent
        .Left = Me.picVBar.Left + Me.picVBar.Width
        .Top = Me.picHBar.Top + Me.picHBar.Height
        .Height = Me.ScaleHeight - lngStatus - .Top + 15
        .Width = Me.ScaleWidth - .Left + 15
    End With
    
    With Me.hgdRefer
        .Redraw = False
        .Left = Me.tabContent.Left + 90
        .Top = Me.tabContent.Top + 350
        .Width = Me.tabContent.Width - 90 * 2
        .Height = Me.tabContent.Height - 350 - 90
        .ColWidth(0) = 0
        .ColWidth(1) = Me.TextWidth("�ո�")
        .ColWidth(2) = .Width - .ColWidth(1) - Me.SysInfo.ScrollBarSize - 15
        .ColWidth(3) = 600
        Call zlGrdRowHeight
        .Redraw = True
    End With
   
    With Me.hgdCodex
        .Redraw = False
        .Left = Me.tabContent.Left + 90
        .Top = Me.tabContent.Top + 350
        .Width = Me.tabContent.Width - 90 * 2
        .Height = Me.tabContent.Height - 350 - 90
        .ColWidth(0) = 0
        .ColWidth(1) = Me.TextWidth("�ո�")
        .ColWidth(3) = 800
        .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(3)
        .Redraw = True
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\�ָ�", "����", Me.picVBar.Left)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\�ָ�", "����", Me.picHBar.Top)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��ʾͣ��", IIf(mnuViewStop.Checked, 1, 0))
End Sub

Private Sub hgdCodex_DblClick()
    If Me.mnuEdit.Visible And Me.mnuEditCodex.Enabled Then
        Call mnuEditCodex_Click
    End If
End Sub

Private Sub hgdCodex_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeySpace Then Exit Sub
    Call hgdCodex_DblClick
End Sub

Private Sub hgdRefer_DblClick()
    If Me.mnuEdit.Visible And Me.mnuEditRefer.Enabled Then
        Call mnuEditRefer_Click
    End If
End Sub

Private Sub hgdRefer_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeySpace Then Exit Sub
    Call hgdRefer_DblClick
End Sub

Private Sub lvwList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwList.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwList.SortOrder = IIf(Me.lvwList.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwList.SortKey = ColumnHeader.Index - 1
        Me.lvwList.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwList_DblClick()
    If Me.lvwList.SelectedItem Is Nothing Then Exit Sub
    If Me.mnuEdit.Visible = False Then Exit Sub
    Call mnuEditModify_Click
    Call lvwList_ItemClick(Me.lvwList.SelectedItem)
End Sub

Private Sub lvwList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim strTemp As String
    
    Err = 0: On Error GoTo ErrHand
    '----------��д�ο���������----------------------
    Me.hgdRefer.Redraw = False
    Me.hgdRefer.Clear
    
    '�������������ȡ
    strTemp = ""
    
    gstrSql = "select distinct ����,����||decode(����,1,'',2,'(Ӣ����)','(����)') as ����" & _
            " from ������ϱ���" & _
            " where ���ID=[1] " & _
            " order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(Item.Key, 2)))
    
    With rsTemp
        Do While Not .EOF
            strTemp = strTemp & "  " & !����
            .MoveNext
        Loop
    End With
    If strTemp = "" Then
        strTemp = "������ƣ�"
    Else
        strTemp = "������ƣ�" & Mid(strTemp, 3)
    End If
    Me.hgdRefer.TextMatrix(0, 1) = strTemp
    Me.hgdRefer.TextMatrix(0, 2) = strTemp
    Me.hgdRefer.TextMatrix(0, 3) = strTemp
    Me.hgdRefer.MergeRow(0) = True
    
    '��׼�����������ȡ
    strTemp = ""
    
    gstrSql = "select L.����||'('||K.���||')' as ����" & _
            " from ������϶��� R,��������Ŀ¼ L,����������� K" & _
            " where R.����ID=L.ID and L.���=K.���� and R.���ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(Item.Key, 2)))
    
    With rsTemp
        Do While Not .EOF
            strTemp = strTemp & "  " & .Fields(0).Value
            .MoveNext
        Loop
    End With
    If strTemp = "" Then
        strTemp = "��׼���룺"
    Else
        strTemp = "��׼���룺" & Mid(strTemp, 3)
    End If
    Me.hgdRefer.TextMatrix(1, 1) = strTemp
    Me.hgdRefer.TextMatrix(1, 2) = strTemp
    Me.hgdRefer.TextMatrix(1, 3) = strTemp
    Me.hgdRefer.MergeRow(1) = True
    
    '�ο����ݵ���ȡ��ʾ
    gstrSql = "select ��Ŀ���,��Ŀ���,nvl(֤�����,0) as ֤�����,0 as �����к�,decode(nvl(֤������,''),'',�ο���Ŀ,֤������) as ����" & _
            " from ������ϲο� " & _
            " where ���id=[1] " & _
            " union" & _
            " select ��Ŀ���,��Ŀ���,nvl(֤�����,0) as ֤�����,�����к�,decode(nvl(֤������,''),'',�����ı�,�ο���Ŀ||'��'||�����ı�) as ����" & _
            " from ������ϲο�" & _
            " where ���id=[1] and length(ltrim(nvl(�����ı�,'')))<>0" & _
            " order by ��Ŀ���,֤�����,�����к�"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(Item.Key, 2)))
    
    With rsTemp
        If .EOF Or .BOF Then
            Me.hgdRefer.Rows = Me.hgdRefer.FixedRows + 1
        Else
            Me.hgdRefer.Rows = Me.hgdRefer.FixedRows + .RecordCount
        End If
        Do While Not .EOF
            If !�����к� = 0 Then
                If !��Ŀ��� = 1 Then
                    Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 1) = "��" & !���� & "��"
                    Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 2) = "��" & !���� & "��"
                    Me.hgdRefer.MergeRow(.AbsolutePosition + Me.hgdRefer.FixedRows - 1) = True
                Else
                    Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 1) = ""
                    If !֤����� = 0 Then
                        Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 2) = "��" & !���� & "��"
                    Else
                        Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 2) = !֤����� & "." & !����
                    End If
                    Me.hgdRefer.MergeRow(.AbsolutePosition + Me.hgdRefer.FixedRows - 1) = False
                End If
            Else
                If !��Ŀ��� = 1 Then
                    Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 1) = Space(4) & !����
                    Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 2) = Space(4) & !����
                    Me.hgdRefer.MergeRow(.AbsolutePosition + Me.hgdRefer.FixedRows - 1) = True
                Else
                    Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 1) = ""
                    Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 2) = Space(4) & !����
                    Me.hgdRefer.MergeRow(.AbsolutePosition + Me.hgdRefer.FixedRows - 1) = False
                End If
            End If
            .MoveNext
        Loop
    End With
    Call zlGrdRowHeight
    Me.hgdRefer.Redraw = True

    '----------��д�ο���������----------------------
    Me.hgdCodex.Redraw = False
    Me.hgdCodex.Clear
    
    '����ԭ����д��
    strTemp = ""
    If Val(Split(Item.Tag, ",")(0)) <> 0 Then
        strTemp = strTemp & "�����ɶȴ�" & Split(Item.Tag, ",")(0) & "ʱ��ʾΪ���Ʋ���"
    End If
    If Val(Split(Item.Tag, ",")(1)) <> 0 Then
        strTemp = strTemp & "�����ɶȴ�" & Split(Item.Tag, ",")(1) & "ʱ��ʾΪ�ٴ�����"
    End If
    strTemp = "����������" & Mid(strTemp, 2)
    With Me.hgdCodex
        .TextMatrix(0, 1) = strTemp
        .TextMatrix(0, 2) = strTemp
        .TextMatrix(0, 3) = strTemp
        .MergeRow(0) = True
        .TextMatrix(1, 1) = "����ϸ��"
        .TextMatrix(1, 2) = "����ϸ��"
        .MergeRow(1) = True
        .TextMatrix(1, 3) = "���ɶ�"
    End With
    
    '�ο����ݵ���ȡ��ʾ
    gstrSql = "select �����,0 as ������,������ as ����,0 as ���ɶ�" & _
            " from ������Ϲ���" & _
            " where ���id=[1] and ������ is not null" & _
            " union" & _
            " select E.�����,E.������,I.������||' '||E.��ϵʽ||' '||E.����ֵ as ����,E.���ɶ�" & _
            " from ������Ϲ��� E,����������Ŀ I" & _
            " where E.��ĿID=I.ID and E.���id=[1] " & _
            " order by �����,������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(Item.Key, 2)))
    
    With rsTemp
        If .EOF Or .BOF Then
            Me.hgdCodex.Rows = Me.hgdCodex.FixedRows + 1
        Else
            Me.hgdCodex.Rows = Me.hgdCodex.FixedRows + .RecordCount
        End If
        Do While Not .EOF
            If !������ = 0 Then
                Me.hgdCodex.TextMatrix(.AbsolutePosition + Me.hgdCodex.FixedRows - 1, 1) = !����� & "��" & !���� & "��"
                Me.hgdCodex.TextMatrix(.AbsolutePosition + Me.hgdCodex.FixedRows - 1, 2) = !����� & "��" & !���� & "��"
                Me.hgdCodex.TextMatrix(.AbsolutePosition + Me.hgdCodex.FixedRows - 1, 3) = !����� & "��" & !���� & "��"
                Me.hgdCodex.MergeRow(.AbsolutePosition + Me.hgdCodex.FixedRows - 1) = True
            Else
                Me.hgdCodex.TextMatrix(.AbsolutePosition + Me.hgdCodex.FixedRows - 1, 1) = ""
                Me.hgdCodex.TextMatrix(.AbsolutePosition + Me.hgdCodex.FixedRows - 1, 2) = !����
                Me.hgdCodex.TextMatrix(.AbsolutePosition + Me.hgdCodex.FixedRows - 1, 3) = !���ɶ�
                Me.hgdCodex.MergeRow(.AbsolutePosition + Me.hgdCodex.FixedRows - 1) = False
            End If
            .MoveNext
        Loop
    End With
    Me.hgdCodex.Redraw = True
    
    Call SetMenu
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub lvwList_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Me.lvwList.SelectedItem Is Nothing Then Exit Sub
    Call mnuEditModify_Click
    Call lvwList_ItemClick(Me.lvwList.SelectedItem)
End Sub

Private Sub lvwList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    If Me.mnuEdit.Visible = False Then Exit Sub
    Me.mnuEditNew.Tag = Me.mnuEditNew.Visible
    
    On Error GoTo RESHOW
    Me.mnuEditNew.Visible = False
    PopupMenu Me.mnuEdit, 2
RESHOW:
    Me.mnuEditNew.Visible = Me.mnuEditNew.Tag
End Sub

Private Sub mnuEditAdd_Click()
    If Me.tvwClass.SelectedItem Is Nothing Then MsgBox "��δ���÷���,������ɾ������", vbExclamation, gstrSysName: Exit Sub
    With frmDiagItem
        .lblnote(0).Tag = IIf(Val(Me.tvwClass.Tag) = 0, "��ҽ", "��ҽ")
        .hgdClass.RowData(0) = Mid(Me.tvwClass.SelectedItem.Key, 2)
        .hgdClass.TextMatrix(0, 1) = "1."
        .hgdClass.TextMatrix(0, 2) = Me.tvwClass.SelectedItem.Text
        .Tag = "����"
        .Show 1, Me
    End With
    Call zlRefRecords
End Sub

Private Sub mnuEditCodex_Click()
    If Me.lvwList.SelectedItem Is Nothing Then Exit Sub
    With frmDiagCodex
        .mlngBarSize = Me.SysInfo.ScrollBarSize
        .hgdCodex.Tag = Mid(Me.lvwList.SelectedItem.Key, 2)
        .Tag = IIf(Val(Me.tvwClass.Tag) = 0, "��ҽ", "��ҽ")
        .Show 1, Me
    End With
    Call lvwList_ItemClick(Me.lvwList.SelectedItem)
End Sub

Private Sub mnuEditDelete_Click()
    Err = 0: On Error GoTo ErrHand
    If Me.ActiveControl.Name = Me.tvwClass.Name Then
        If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
        If MsgBox("���ɾ���÷�����", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        gstrSql = "zl_������Ϸ���_delete(" & Mid(Me.tvwClass.SelectedItem.Key, 2) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
        
        Dim strParentKey As String
        If Me.tvwClass.SelectedItem.Next Is Nothing Then
            If Me.tvwClass.SelectedItem.Parent Is Nothing Then
                Call zlRefClasses
            Else
                strParentKey = Me.tvwClass.SelectedItem.Parent.Key
                Call Me.tvwClass.Nodes.Remove(Me.tvwClass.SelectedItem.Key)
                If Me.tvwClass.Nodes(strParentKey).Children = 0 Then
                    Call zlRefClasses(Mid(Me.tvwClass.Nodes(strParentKey).Key, 2))
                Else
                    Call zlRefClasses(Mid(Me.tvwClass.Nodes(strParentKey).Child.Key, 2))
                End If
            End If
        Else
            Call zlRefClasses(Mid(Me.tvwClass.SelectedItem.Next.Key, 2))
        End If
    Else
        If Me.lvwList.SelectedItem Is Nothing Then Exit Sub
        If MsgBox("���ɾ���òο���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        gstrSql = "zl_�������Ŀ¼_delete(" & Mid(Me.lvwList.SelectedItem.Key, 2) & ")"
        
        On Error Resume Next
        Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
        
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
        
        Call zlRefRecords
    End If
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditDiagAndSymptom_Click()
    '��ϲ��ֶ�Ӧ����
    '���˺�:2007/08/17
    
    Dim lng���ID As Long, lng����ID As Long, bln��ҽ As Boolean
    Dim blnEdit As Boolean
    
    If Me.lvwList.SelectedItem Is Nothing Then
        lng���ID = 0
    Else
        lng���ID = Val(Mid(lvwList.SelectedItem.Key, 2))
    End If
    If tvwClass.SelectedItem Is Nothing Then
        lng����ID = 0
    Else
        lng����ID = Mid(Me.tvwClass.SelectedItem.Key, 2)
    End If
    bln��ҽ = IIf(Val(Me.tvwClass.Tag) = 0, False, True)
    blnEdit = InStr(1, mstrPrivs, ";��ϲ��ֶ�Ӧ;") > 0
    Call frmDiagAndSymptom.ShowEdit(Me, lng���ID, lng����ID, blnEdit, bln��ҽ)
End Sub

Private Sub mnuEditModify_Click()
    If Me.ActiveControl.Name = Me.tvwClass.Name Then
        If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
        With frmDiagClass
            .lblKind.Caption = IIf(Val(Me.tvwClass.Tag) = 0, "��ҽ", "��ҽ")
            If Me.tvwClass.SelectedItem.Parent Is Nothing Then
                .txtParent.Tag = 0
                .txtParent.Text = "(��)"
                .txtUpCode.Text = ""
                .txtCode.Text = Mid(Split(Me.tvwClass.SelectedItem.Text, "]")(0), 2)
                .txtCode.MaxLength = Len(.txtCode.Text)
                .txtCode.Tag = .txtCode.MaxLength
            Else
                .txtParent.Tag = Mid(Me.tvwClass.SelectedItem.Parent.Key, 2)
                .txtParent.Text = Me.tvwClass.SelectedItem.Parent.Text
                .txtUpCode.Text = Mid(Split(Me.tvwClass.SelectedItem.Parent.Text, "]")(0), 2)
                .txtCode.Text = Mid(Split(Me.tvwClass.SelectedItem.Text, "]")(0), Len(.txtUpCode.Text) + 2)
                .txtCode.MaxLength = Len(.txtCode.Text)
                .txtCode.Tag = .txtCode.MaxLength
            End If
            .txtName = Split(Me.tvwClass.SelectedItem.Text, "]")(1)
            .txtSymbol = Me.tvwClass.SelectedItem.Tag
            .Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
            .Show 1, Me
        End With
        Call zlRefClasses(Mid(Me.tvwClass.SelectedItem.Key, 2))
    Else
        If Me.lvwList.SelectedItem Is Nothing Then Exit Sub
        With frmDiagItem
            .lblnote(0).Tag = IIf(Val(Me.tvwClass.Tag) = 0, "��ҽ", "��ҽ")     '��ǰ���
            .txtItem(0).Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)              '��ǰ����ID
            .Tag = Mid(lvwList.SelectedItem.Key, 2)                             '��ǰ��ĿID
            .Show 1, Me
        End With
        Call zlRefRecords(Mid(lvwList.SelectedItem.Key, 2))
    End If
End Sub

Private Sub mnuEditNew_Click()
    With frmDiagClass
        .lblKind.Caption = IIf(Val(Me.tvwClass.Tag) = 0, "��ҽ", "��ҽ")
        If Me.tvwClass.SelectedItem Is Nothing Then
            .txtParent.Tag = 0
        Else
            .txtParent.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
        End If
        .Tag = "����"
        .Show 1, Me
    End With
    If Me.tvwClass.SelectedItem Is Nothing Then
        Call zlRefClasses
    Else
        Call zlRefClasses(Mid(Me.tvwClass.SelectedItem.Key, 2))
    End If
End Sub

Private Sub mnuEditRefer_Click()
    Dim frmRefer As New frmDiagRefer
    If Me.lvwList.SelectedItem Is Nothing Then Exit Sub
    With frmRefer
        .mlngBarSize = Me.SysInfo.ScrollBarSize
        .hgdRefer.Tag = Mid(Me.lvwList.SelectedItem.Key, 2)
        .Tag = IIf(Val(Me.tvwClass.Tag) = 0, "��ҽ", "��ҽ")
        .Show , Me
    End With
End Sub

Private Sub mnuEditStart_Click()
    Call StopAndReuse(False)
End Sub

Private Sub StopAndReuse(ByVal blnStop As Boolean)
    '--------------------------------------------------------------------------------------
    '����:ͣ�û�����
    '����:blnStop-�Ƿ�ͣ��,true-ͣ��,false-����
    '--------------------------------------------------------------------------------------
    
    Dim lng����ID As Long
    Dim strSQL As String, intIndex As Integer
    Dim i As Integer
    Dim ReMoveRow As Long
    
    If lvwList.SelectedItem Is Nothing Then Exit Sub
    If MsgBox("���Ƿ����Ҫ" & IIf(blnStop, "ͣ��", "����") & "��" & lvwList.SelectedItem.Text & "���ļ�����", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    
    lng����ID = Val(Mid(lvwList.SelectedItem.Key, 2))
    If lng����ID <= 0 Then Exit Sub
    
'    Err = 0: On Error GoTo ErrHand:
    
    If blnStop Then
        strSQL = "Zl_�������Ŀ¼_Stop(" & lng����ID & ")"
    Else
        strSQL = "Zl_�������Ŀ¼_Reuse(" & lng����ID & ")"
    End If
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    '��ͣ����Ŀ��ͣ��ͼ���ʶ
    With lvwList.SelectedItem
        .Icon = IIf(blnStop, "itemstop", "item")
        .SmallIcon = IIf(blnStop, "itemstop", "item")
        .ForeColor = IIf(blnStop, vbRed, &H80000008)
    End With
    
    '��ͣ���ú�ɫǰ��ɫ��ʶ
    For i = 2 To lvwList.ColumnHeaders.Count
         lvwList.SelectedItem.ListSubItems(i - 1).ForeColor = IIf(blnStop, vbRed, &H80000008)
    Next
    
    If mnuViewStop.Checked Then
        '��ʾͣ����Ŀʱ
        If blnStop = False Then
            '����ʱ
            lvwList.SelectedItem.SubItems(lvwList.ColumnHeaders("_����ʱ��").Index - 1) = "3000-01-01"
        Else
            'ͣ��ʱ
            lvwList.SelectedItem.SubItems(lvwList.ColumnHeaders("_����ʱ��").Index - 1) = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
        End If
    Else
        '����ʾͣ����Ŀʱ
        If blnStop = False Then
            '����ʱ
            lvwList.SelectedItem.SubItems(lvwList.ColumnHeaders("_����ʱ��").Index - 1) = "3000-01-01"
        Else
            'ͣ��ʱ�Ƴ���ͣ����Ŀ
            With lvwList
                intIndex = .SelectedItem.Index
                .ListItems.Remove .SelectedItem.Key
                If .ListItems.Count > 0 Then
                    intIndex = IIf(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
                    .ListItems(intIndex).Selected = True
                    .ListItems(intIndex).EnsureVisible
                End If
            End With
        End If
    End If
    
    '����˵��Ͱ�ť״̬
    Call SetMenu
   
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetMenu()
    '����:���ò˵��Ͱ�ť����Ч״̬
    Dim blnStop As Boolean
    
    If mnuEdit.Visible = False Then Exit Sub
    
    If Not lvwList.SelectedItem Is Nothing Then
        blnStop = lvwList.SelectedItem.SubItems(lvwList.ColumnHeaders("_����ʱ��").Index - 1) <> "3000-01-01" And lvwList.SelectedItem.SubItems(lvwList.ColumnHeaders("_����ʱ��").Index - 1) <> ""
    End If
    
    tlbThis.Buttons("Stop").Enabled = Not blnStop
    tlbThis.Buttons("Start").Enabled = blnStop

    mnuEditStart.Enabled = blnStop
    mnuEditStop.Enabled = Not blnStop
        
    tlbThis.Buttons("Mod").Enabled = Not blnStop
    tlbThis.Buttons("Del").Enabled = Not blnStop
    mnuEditDelete.Enabled = Not blnStop
    mnuEditModify.Enabled = Not blnStop
End Sub
Private Sub mnuEditStop_Click()
    Call StopAndReuse(True)
End Sub

Private Sub mnuFileExcel_Click()
    Call zlRptPrint(3)
End Sub

Private Sub mnuFileOut_Click()
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim bytMode As Byte
    If Me.lvwList.SelectedItem Is Nothing Then Exit Sub
    If Me.hgdRefer.TextMatrix(Me.hgdRefer.FixedRows, 2) = "" Then Exit Sub
    'On Error Resume Next
    Set objPrint.Body = Me.hgdRefer
    With objPrint.Title
        .Text = Me.lvwList.SelectedItem.Text & "���Ʋο��淶"
        .Font.Size = 11
    End With
    If Me.lvwList.SelectedItem.SubItems(Me.lvwList.ColumnHeaders("_����").Index - 1) <> "" Then
        objRow.Add ""
        objRow.Add "(" & Me.lvwList.SelectedItem.SubItems(Me.lvwList.ColumnHeaders("_����").Index - 1) & ")"
        objPrint.BelowAppRows.Add objRow
    End If
    bytMode = zlPrintAsk(objPrint)
    If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
End Sub

Private Sub mnuFilePreview_Click()
    Call zlRptPrint(0)
End Sub

Private Sub mnuFilePrint_Click()
    Call zlRptPrint(1)
End Sub

Private Sub mnuFilePrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub mnuhelpAbout_Click()
    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
End Sub

Private Sub mnuHelpHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    'Ĭ�ϲ���������=����id����Ŀ=�����Ŀid
    Dim lng����ID As Long
    Dim lng��Ŀid As Long
    
    If Not Me.tvwClass.SelectedItem Is Nothing Then
        lng����ID = Mid(Me.tvwClass.SelectedItem.Key, 2)
    End If
    
    If Not Me.lvwList.SelectedItem Is Nothing Then
        lng��Ŀid = Mid(lvwList.SelectedItem.Key, 2)
    End If
    
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
        "����=" & IIf(lng����ID = 0, "", lng����ID), _
        "��Ŀ=" & IIf(lng��Ŀid = 0, "", lng��Ŀid))
End Sub

Private Sub mnuViewFind_Click()
    Call frmDiagnoseFind.ShowFind(mnuViewStop.Checked)
End Sub

Private Sub mnuViewRefresh_Click()
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    Call zlRefRecords
End Sub

Private Sub mnuViewStates_Click()
    Me.mnuViewStates.Checked = Not Me.mnuViewStates.Checked
    Me.stbThis.Visible = Me.mnuViewStates.Checked
    Form_Resize
End Sub

Private Sub mnuViewStop_Click()
    mnuViewStop.Checked = Not mnuViewStop.Checked
    Call mnuViewRefresh_Click
End Sub

Private Sub mnuViewToolbarStand_Click()
    Me.mnuViewToolbarStand.Checked = Not Me.mnuViewToolbarStand.Checked
    Me.clbThis.Visible = Me.mnuViewToolbarStand.Checked
    Form_Resize
End Sub

Private Sub mnuViewToolBarText_Click()
    Dim i As Integer
    Me.mnuViewToolbarText.Checked = Not Me.mnuViewToolbarText.Checked
    If Me.mnuViewToolbarText.Checked Then
        For i = 1 To Me.tlbThis.Buttons.Count
            Me.tlbThis.Buttons(i).Caption = Me.tlbThis.Buttons(i).Tag
        Next
    Else
        For i = 1 To Me.tlbThis.Buttons.Count
            Me.tlbThis.Buttons(i).Caption = ""
        Next
    End If
    Me.clbThis.Bands(1).MinHeight = Me.tlbThis.Height
    Me.clbThis.Refresh
    Form_Resize
End Sub

Private Sub picClass_Resize()
    Dim intCount As Integer
    Err = 0: On Error Resume Next
    For intCount = Me.cmdKind.LBound To Me.cmdKind.UBound
        Me.cmdKind(intCount).Left = Me.picClass.ScaleLeft + 15
        Me.cmdKind(intCount).Width = Me.picClass.ScaleWidth
        Me.cmdKind(intCount).Height = 300
        If Val(Me.cmdKind(intCount).Tag) = 0 Then
            Me.cmdKind(intCount).Top = Me.picClass.ScaleTop + 285 * intCount
            Me.tvwClass.Top = Me.picClass.ScaleTop + 285 * (intCount + 1)
        Else
            Me.cmdKind(intCount).Top = Me.picClass.ScaleHeight - 285 * (Me.cmdKind.UBound - intCount + 1)
        End If
    Next
    Me.tvwClass.Left = Me.picClass.ScaleLeft + 15
    Me.tvwClass.Width = Me.picClass.ScaleWidth
    Me.tvwClass.Height = Me.picClass.ScaleHeight - 285 * (Me.cmdKind.UBound + 1) - 15
End Sub

Private Sub picHBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Me.picHBar.Top = Me.picHBar.Top + Y
    End If
End Sub

Private Sub picHBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Call Form_Resize
    End If
End Sub

Private Sub picVBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Me.picVBar.Left = Me.picVBar.Left + X
    End If
End Sub

Private Sub picVBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Call Form_Resize
    End If
End Sub

Private Sub tabContent_Click()
    If Me.tabContent.Tabs(1).Selected Then
        Me.hgdRefer.Visible = True
        Me.hgdCodex.Visible = False
    Else
        Me.hgdRefer.Visible = False
        Me.hgdCodex.Visible = True
    End If
End Sub

Private Sub tlbThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Preview"
        Call mnuFilePreview_Click
    Case "Print"
        Call mnuFilePrint_Click
    Case "New"
        Call mnuEditNew_Click
    Case "Add"
        Call mnuEditAdd_Click
    Case "Mod"
        Call mnuEditModify_Click
    Case "Del"
        Call mnuEditDelete_Click
    Case "Find"
        Call mnuViewFind_Click
    Case "Help"
        Call mnuHelpHelp_Click
    Case "Exit"
        Call mnuFileExit_Click
    Case "Start"
        Call mnuEditStart_Click
    Case "Stop"
        Call mnuEditStop_Click
    End Select
End Sub

Private Sub tlbThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    PopupMenu Me.mnuToolBar, 2
End Sub

Private Sub tvwClass_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    If Me.mnuEdit.Visible = False Then Exit Sub
    Me.mnuEditAdd.Tag = Me.mnuEditAdd.Visible
    Me.mnuEditSpt1.Tag = Me.mnuEditSpt1.Visible
    Me.mnuEditRefer.Tag = Me.mnuEditRefer.Visible
    Me.mnuEditCodex.Tag = Me.mnuEditRefer.Visible
    
    On Error GoTo RESHOW
    Me.mnuEditAdd.Visible = False
    Me.mnuEditSpt1.Visible = False
    Me.mnuEditRefer.Visible = False
    Me.mnuEditCodex.Visible = False
    PopupMenu Me.mnuEdit, 2
RESHOW:
    Me.mnuEditAdd.Visible = Me.mnuEditAdd.Tag
    Me.mnuEditSpt1.Visible = Me.mnuEditSpt1.Tag
    Me.mnuEditRefer.Visible = Me.mnuEditRefer.Tag
    Me.mnuEditCodex.Visible = Me.mnuEditCodex.Tag
End Sub

Private Sub tvwClass_NodeClick(ByVal Node As MSComctlLib.Node)
    If Me.lvwList.Tag = Node.Key Then Exit Sub
    Me.lvwList.Tag = Node.Key
    Call zlRefRecords
End Sub

Private Sub zlRefClasses(Optional lngNode As Long)
    '---------------------------------------------
    '��д������Ϸ���
    '---------------------------------------------
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select ID,�ϼ�ID,����,����,����" & _
            " From ������Ϸ���" & _
            " Where ��� = [1] " & _
            " start with �ϼ�ID is null" & _
            " connect by prior ID=�ϼ�ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, 1 + Val(Me.tvwClass.Tag))
    
    With rsTemp
        Me.tvwClass.Visible = False
        Me.tvwClass.Nodes.Clear
        Do While Not .EOF
            If IsNull(!�ϼ�ID) Then
                Set objNode = Me.tvwClass.Nodes.Add(, , "_" & !ID, "[" & !���� & "]" & !����, "close")
            Else
                Set objNode = Me.tvwClass.Nodes.Add("_" & !�ϼ�ID, tvwChild, "_" & !ID, "[" & !���� & "]" & !����, "close")
            End If
            objNode.Sorted = True
            objNode.Tag = IIf(IsNull(!����), "", !����)
            objNode.ExpandedImage = "expend"
            .MoveNext
        Loop
        Me.tvwClass.Visible = True
    End With
    If Me.tvwClass.Nodes.Count > 0 Then
        If lngNode <> 0 Then
            Me.tvwClass.Nodes("_" & lngNode).Selected = True
        Else
            Me.tvwClass.Nodes(1).Selected = True
        End If
        Call zlRefRecords
    Else
        Me.lvwList.ListItems.Clear
        Call zlGrdClear
    End If
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub zlRefRecords(Optional lngItem As Long)
    '---------------------------------------------
    '��д�����ο��б�
    '---------------------------------------------
    Dim strTemp As String
    Dim strIco As String
    Dim i As Integer
    
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select L.ID,L.����,L.����,L.˵��,L.����,nvl(L.����,0)||','||nvl(L.�ٴ�,0) as ���ɶ�,L.����ʱ��,L.����ʱ��" & _
            " from ����������� C, �������Ŀ¼ L" & _
            " where C.���ID=L.ID and C.����ID=[1] " & _
            IIf(mnuViewStop.Checked, "", " and (L.����ʱ�� is null or L.����ʱ��>=to_date('3000-01-01','yyyy-mm-dd'))") & _
            " order by L.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(Me.tvwClass.SelectedItem.Key, 2)))
    
    With rsTemp
        Me.lvwList.ListItems.Clear
        Me.lvwList.ForeColor = &H80000008
        Do While Not .EOF
            '�ó���ȷ��ͼ��
            strTemp = IIf(NVL(rsTemp!����ʱ��) = "", "3000-01-01", Format(!����ʱ��, "YYYY-MM-DD"))
            strIco = IIf(strTemp <> "3000-01-01", "itemstop", "item")
            
            Set objItem = Me.lvwList.ListItems.Add(, "_" & !ID, !����, strIco, strIco)
            objItem.SubItems(Me.lvwList.ColumnHeaders("_����").Index - 1) = !����
            objItem.SubItems(Me.lvwList.ColumnHeaders("_˵��").Index - 1) = IIf(IsNull(!˵��), "", !˵��)
            objItem.SubItems(Me.lvwList.ColumnHeaders("_����").Index - 1) = IIf(IsNull(!����), "", !����)
            objItem.SubItems(Me.lvwList.ColumnHeaders("_����ʱ��").Index - 1) = IIf(IsNull(!����ʱ��), "", Format(!����ʱ��, "YYYY-MM-DD"))
            objItem.SubItems(Me.lvwList.ColumnHeaders("_����ʱ��").Index - 1) = IIf(IsNull(!����ʱ��), "", Format(!����ʱ��, "YYYY-MM-DD"))
            objItem.Tag = !���ɶ�
            If !ID = lngItem Then
                objItem.Selected = True
            End If
   
            If strTemp <> "3000-01-01" Then
                objItem.ForeColor = vbRed
                For i = 2 To lvwList.ColumnHeaders.Count
                    objItem.ListSubItems(i - 1).ForeColor = vbRed
                Next
            End If
            
            .MoveNext
        Loop
    End With
    If Me.lvwList.ListItems.Count > 0 Then
        If Me.lvwList.SelectedItem Is Nothing Then Me.lvwList.ListItems(1).Selected = True
        Call lvwList_ItemClick(Me.lvwList.SelectedItem)
        Err = 0: On Error Resume Next
        DoEvents: Me.lvwList.SelectedItem.EnsureVisible
        Me.stbThis.Panels(2).Text = "�÷��๲��" & Me.lvwList.ListItems.Count & "����ϲο�"
    Else
        Call zlGrdClear
        Me.stbThis.Panels(2).Text = ""
    End If
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub zlGrdRowHeight()
    '---------------------------------------------
    '���ݵ������ݵ�������������и߶ȣ��Ա�֤���ݵ�������ʾ
    '---------------------------------------------
    Dim intRow As Integer, lngColWidth As Long
    
    On Error Resume Next
    With Me.hgdRefer
        For intRow = .FixedRows To .Rows - 1
            If .TextMatrix(intRow, 1) = "" Then
                lngColWidth = .ColWidth(2)
            Else
                lngColWidth = .ColWidth(1) + .ColWidth(2)
            End If
            Me.lblScale.Width = lngColWidth - 90
            Me.lblScale.Caption = .TextMatrix(intRow, 2)
            .RowHeight(intRow) = Me.lblScale.Height + 75
        Next
    End With
End Sub

Private Sub zlGrdClear()
    '---------------------------------------------
    '����������ʾ����
    '---------------------------------------------
    With Me.hgdRefer
        .Clear
        .TextMatrix(0, 1) = "������ƣ�"
        .TextMatrix(0, 2) = "������ƣ�"
        .TextMatrix(0, 3) = "������ƣ�"
        .MergeRow(0) = True
        .TextMatrix(1, 1) = "��׼���룺"
        .TextMatrix(1, 2) = "��׼���룺"
        .TextMatrix(1, 3) = "��׼���룺"
        .MergeRow(1) = True
    End With
    With Me.hgdCodex
        .Clear
        .TextMatrix(0, 1) = "����������"
        .TextMatrix(0, 2) = "����������"
        .TextMatrix(0, 3) = "����������"
        .MergeRow(0) = True
        .TextMatrix(1, 1) = "����ϸ��"
        .TextMatrix(1, 2) = "����ϸ��"
        .MergeRow(1) = True
        .TextMatrix(1, 3) = "���ɶ�"
    End With
End Sub


Private Sub zlRptPrint(ByVal bytMode As Byte)
    '-------------------------------------------------
    '����:��¼���ӡ
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '-------------------------------------------------
    Dim objPrint As New zlPrintLvw
    Err = 0: On Error Resume Next
    Set objPrint.Body.objData = Me.lvwList
    objPrint.Title.Text = "������ϲο�Ŀ¼"
    objPrint.UnderAppItems.Add "���ࣺ" & Me.tvwClass.SelectedItem.Text
    objPrint.BelowAppItems.Add "��ӡʱ�䣺" & Now
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrViewLvw objPrint, bytMode
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
End Sub

Public Sub zlLocateItem(lng����ID As Long, lng���ID As Long)
    '---------------------------------------------
    '��λ��ָ������ϲο���Ŀ���ڲ���ʱʹ��
    '---------------------------------------------
    Set Me.tvwClass.SelectedItem = Me.tvwClass.Nodes("_" & lng����ID)
    Me.tvwClass.Nodes("_" & lng����ID).Selected = True
    Me.tvwClass.SelectedItem.EnsureVisible
    Call zlRefRecords
    Set Me.lvwList.SelectedItem = Me.lvwList.ListItems("_" & lng���ID)
    Me.lvwList.SelectedItem.EnsureVisible
    Call lvwList_ItemClick(Me.lvwList.SelectedItem)
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub
