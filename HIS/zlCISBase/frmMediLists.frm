VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMediLists 
   BackColor       =   &H8000000C&
   Caption         =   "ҩƷĿ¼����"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   11205
   Icon            =   "frmMediLists.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8310
   ScaleWidth      =   11205
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picHBar 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
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
      Left            =   2805
      MousePointer    =   7  'Size N S
      ScaleHeight     =   30
      ScaleWidth      =   6075
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3870
      Width           =   6075
   End
   Begin VB.PictureBox picVBar 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
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
      Height          =   6660
      Left            =   2580
      MousePointer    =   9  'Size W E
      ScaleHeight     =   6660
      ScaleWidth      =   30
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   795
      Width           =   30
   End
   Begin VB.PictureBox picClass 
      Height          =   6735
      Left            =   0
      ScaleHeight     =   6675
      ScaleWidth      =   2340
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   720
      Width           =   2400
      Begin VB.CommandButton cmdKind 
         Caption         =   "���˽��"
         Height          =   300
         Index           =   4
         Left            =   0
         TabIndex        =   24
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   1155
         Width           =   2295
      End
      Begin VB.CommandButton cmdKind 
         Caption         =   "����ҩƷ"
         Height          =   300
         Index           =   3
         Left            =   0
         TabIndex        =   23
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   870
         Width           =   2295
      End
      Begin VB.CommandButton cmdKind 
         Caption         =   "�в�ҩ(&7)"
         Height          =   300
         Index           =   2
         Left            =   0
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   585
         Width           =   2295
      End
      Begin VB.CommandButton cmdKind 
         Caption         =   "�г�ҩ(&6)"
         Height          =   300
         Index           =   1
         Left            =   0
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   300
         Width           =   2295
      End
      Begin VB.CommandButton cmdKind 
         Caption         =   "����ҩ(&5)"
         Height          =   300
         Index           =   0
         Left            =   0
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   15
         Width           =   2295
      End
      Begin MSComctlLib.TreeView tvwClass 
         Height          =   4800
         Left            =   0
         TabIndex        =   5
         Tag             =   "1000"
         Top             =   1440
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   8467
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   353
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "imgList"
         Appearance      =   0
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   1785
      Top             =   6750
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":030A
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":08A4
            Key             =   "expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":0E3E
            Key             =   "��ҩU"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":13D8
            Key             =   "��ҩS"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":1972
            Key             =   "���U"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":1F0C
            Key             =   "���S"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":24A6
            Key             =   "��ҩU"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":2A40
            Key             =   "��ҩS"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":2FDA
            Key             =   "Packer"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":3574
            Key             =   "NoPacker"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":3B0E
            Key             =   "�ݹ�S"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":43E8
            Key             =   "�ݹ�U"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   2895
      Left            =   2835
      TabIndex        =   1
      Top             =   855
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   5106
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
      Top             =   7935
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMediLists.frx":4CC2
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14684
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
      TabIndex        =   3
      Top             =   0
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   11205
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
         TabIndex        =   9
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
            NumButtons      =   17
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Preview"
               Description     =   "Ԥ��"
               Object.ToolTipText     =   "Ԥ����ǰ��"
               Object.Tag             =   "Ԥ��"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Description     =   "��ӡ"
               Object.ToolTipText     =   "��ӡ��ǰ��"
               Object.Tag             =   "��ӡ"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Class"
               Description     =   "����"
               Object.ToolTipText     =   "����ҩƷ����"
               Object.Tag             =   "����"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ʒ��"
               Key             =   "Item"
               Description     =   "Ʒ��"
               Object.ToolTipText     =   "����ҩƷƷ��"
               Object.Tag             =   "Ʒ��"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "���"
               Key             =   "Spec"
               Description     =   "���"
               Object.ToolTipText     =   "����ͬ��ҩƷ�Ĺ��"
               Object.Tag             =   "���"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "sp2"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Start"
               Description     =   "����"
               Object.ToolTipText     =   "����ָ����ͣ��ҩƷ"
               Object.Tag             =   "����"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ͣ��"
               Key             =   "Stop"
               Description     =   "ͣ��"
               Object.ToolTipText     =   "ͣ��ָ��������ҩƷ"
               Object.Tag             =   "ͣ��"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Limit"
               Description     =   "����"
               Object.ToolTipText     =   "���������޶�"
               Object.Tag             =   "����"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Find"
               Description     =   "����"
               Object.ToolTipText     =   "���������Ŀ"
               Object.Tag             =   "����"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Filter"
               Description     =   "����"
               Object.ToolTipText     =   "����ҩƷĿ¼"
               Object.Tag             =   "����"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "����"
               Object.ToolTipText     =   "��ǰ��������"
               Object.Tag             =   "����"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Exit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   11
            EndProperty
         EndProperty
         Begin VB.PictureBox picFind 
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   9000
            ScaleHeight     =   285.714
            ScaleMode       =   0  'User
            ScaleWidth      =   495
            TabIndex        =   26
            Top             =   210
            Width           =   495
            Begin VB.Label lbl���� 
               Caption         =   "����"
               Height          =   255
               Left            =   120
               TabIndex        =   27
               Top             =   75
               Width           =   495
            End
         End
         Begin VB.TextBox txtFind 
            Height          =   300
            Left            =   9600
            MaxLength       =   10
            TabIndex        =   25
            Tag             =   "����"
            Top             =   240
            Width           =   1425
         End
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
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":5554
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":576E
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":5988
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":5BA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":5DBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":64B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":66D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":68EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":6FE4
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":71FE
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":741E
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":763E
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":7958
            Key             =   "Limit"
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
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":8052
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":8272
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":8492
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":86AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":88C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":8FC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":91DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":93F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":9AEE
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":9D08
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":9F28
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":A148
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":A462
            Key             =   "Limit"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab tabContent 
      Height          =   2805
      HelpContextID   =   1
      Left            =   2760
      TabIndex        =   10
      Top             =   4635
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   4948
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      WordWrap        =   0   'False
      TabCaption(0)   =   "ҩƷ���(&S)"
      TabPicture(0)   =   "frmMediLists.frx":AB5C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraComment(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lvwSpecs"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "�ۼۼ�¼(&L)"
      TabPicture(1)   =   "frmMediLists.frx":AB78
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "hgdPrice"
      Tab(1).Control(1)=   "fraComment(1)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "�ɱ��ۼ�¼(&C)"
      TabPicture(2)   =   "frmMediLists.frx":AB94
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "hgdCost"
      Tab(2).Control(1)=   "chkStock"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "�ѱ�ȼ�(&F)"
      TabPicture(3)   =   "frmMediLists.frx":ABB0
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "hgdCharge"
      Tab(3).ControlCount=   1
      Begin VB.CheckBox chkStock 
         Caption         =   "ֻ��ʾ�п��۸��¼"
         Height          =   180
         Left            =   -69840
         TabIndex        =   28
         Top             =   80
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin MSComctlLib.ListView lvwSpecs 
         Height          =   1395
         Left            =   105
         TabIndex        =   12
         Top             =   405
         Width           =   7410
         _ExtentX        =   13070
         _ExtentY        =   2461
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imgList"
         SmallIcons      =   "imgList"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgdPrice 
         Height          =   1665
         Left            =   -74880
         TabIndex        =   11
         Top             =   360
         Width           =   7440
         _ExtentX        =   13123
         _ExtentY        =   2937
         _Version        =   393216
         Rows            =   4
         Cols            =   13
         FixedCols       =   0
         AllowBigSelection=   0   'False
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   13
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgdCost 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   21
         Top             =   360
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   3836
         _Version        =   393216
         Rows            =   4
         Cols            =   9
         FixedCols       =   0
         AllowBigSelection=   0   'False
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   9
      End
      Begin VB.Frame fraComment 
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   0
         Left            =   90
         TabIndex        =   13
         Top             =   1860
         Width           =   7410
         Begin VB.Label lblComment 
            AutoSize        =   -1  'True
            Caption         =   "3�������Ч�ÿ˱���3��"
            Height          =   180
            Index           =   2
            Left            =   0
            TabIndex        =   16
            Top             =   510
            Width           =   1980
         End
         Begin VB.Label lblComment 
            AutoSize        =   -1  'True
            Caption         =   "2�������з�����Ч�ڸ��ٹ���"
            Height          =   180
            Index           =   1
            Left            =   0
            TabIndex        =   15
            Top             =   270
            Width           =   2430
         End
         Begin VB.Label lblComment 
            AutoSize        =   -1  'True
            Caption         =   "1���ۼ۵�λ��Ƭ�����ﵥλ��ƿ"
            Height          =   180
            Index           =   0
            Left            =   0
            TabIndex        =   14
            Top             =   30
            Width           =   2610
         End
      End
      Begin VB.Frame fraComment 
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   1
         Left            =   -74910
         TabIndex        =   17
         Top             =   2085
         Width           =   7410
         Begin VB.Label lblComment 
            AutoSize        =   -1  'True
            Caption         =   "1��ʱ��ҩ��ָ��������6Ԫ/ƿ������"
            Height          =   180
            Index           =   3
            Left            =   0
            TabIndex        =   19
            Top             =   30
            Width           =   2970
         End
         Begin VB.Label lblComment 
            AutoSize        =   -1  'True
            Caption         =   "2������ۼ�198.25Ԫ/ƿ�����ݲ�����ݷѱ�����Żݻ�Ӽۡ�"
            Height          =   180
            Index           =   4
            Left            =   0
            TabIndex        =   18
            Top             =   270
            Width           =   5040
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgdCharge 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   22
         Top             =   360
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   3836
         _Version        =   393216
         Rows            =   4
         Cols            =   9
         FixedCols       =   0
         AllowBigSelection=   0   'False
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   9
      End
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
      Begin VB.Menu mnuFileSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePara 
         Caption         =   "��������(&A)"
         Shortcut        =   {F12}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSpt2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuClass 
      Caption         =   "����(&K)"
      Begin VB.Menu mnuClassAdd 
         Caption         =   "����(&I)"
         Shortcut        =   +{INSERT}
      End
      Begin VB.Menu mnuClassMod 
         Caption         =   "�޸�(&U)"
      End
      Begin VB.Menu mnuClassDel 
         Caption         =   "ɾ��(&E)"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu mnuClassSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClassStar 
         Caption         =   "���÷���(&R)"
      End
      Begin VB.Menu mnuClassStop 
         Caption         =   "ͣ�÷���(&S)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "ҩƷ(&E)"
      Begin VB.Menu mnuEditItemAdd 
         Caption         =   "����Ʒ��(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditItemMod 
         Caption         =   "�޸�Ʒ��(&M)"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuEditItemDel 
         Caption         =   "ɾ��Ʒ��(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditSpt6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditItemTabu 
         Caption         =   "�������(&T)..."
      End
      Begin VB.Menu mnuEditItemUsage 
         Caption         =   "�÷�����(&U)"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuEditItemBill 
         Caption         =   "��Ӧ����(&B)"
      End
      Begin VB.Menu mnuEditSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSpecAdd 
         Caption         =   "�������(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuEditSpecMod 
         Caption         =   "�޸Ĺ��(&G)"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuEditSpecDel 
         Caption         =   "ɾ�����(&Y)"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuEditSpecExp 
         Caption         =   "�����չ��Ϣ����(&E)"
      End
      Begin VB.Menu mnuEditSpt7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditItemPart 
         Caption         =   "�洢�ⷿ(&R)..."
      End
      Begin VB.Menu mnuEditSpecLimit 
         Caption         =   "��������(&L)..."
      End
      Begin VB.Menu mnuEditSpecProtocol 
         Caption         =   "Э��ҩƷ(&P)..."
      End
      Begin VB.Menu mnuEditSpecSelf 
         Caption         =   "����ҩƷ(&H)..."
      End
      Begin VB.Menu mnuEditSpecUnit 
         Caption         =   "�б굥λ(&V)..."
      End
      Begin VB.Menu mnuEditManFac 
         Caption         =   "������׼�ĺ�(&C)"
      End
      Begin VB.Menu mnuEditSendType 
         Caption         =   "��ҩ����(&S)..."
      End
      Begin VB.Menu mnuEditSpt5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditRate 
         Caption         =   "�ֶμӳ���(&L)"
      End
      Begin VB.Menu mnuEditSpt2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditVariBatch 
         Caption         =   "Ʒ�������޸�(&W)"
      End
      Begin VB.Menu mnuEditSpecBatch 
         Caption         =   "��������޸�(&Y)"
      End
      Begin VB.Menu mnuEditExcel 
         Caption         =   "������Ŀ"
      End
      Begin VB.Menu mnuEditContrast 
         Caption         =   "�����ɹ�����(&G)"
      End
      Begin VB.Menu mnuEditSpt3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditStart 
         Caption         =   "����(&R)"
      End
      Begin VB.Menu mnuEditStop 
         Caption         =   "ͣ��(&S)"
      End
      Begin VB.Menu mnuEditSpt4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPriceChargeSet1 
         Caption         =   "�ѱ�����(&C)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditSptPacker 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUploadDrugInfo 
         Caption         =   "�����ϴ�����ƽ̨"
      End
   End
   Begin VB.Menu mnuPrice 
      Caption         =   "�۸�(&P)"
      Begin VB.Menu mnuPriceChargeSet 
         Caption         =   "�ѱ�����(&C)"
      End
      Begin VB.Menu mnuPriceSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPriceTable 
         Caption         =   "���ۼ�¼��(&S)"
      End
      Begin VB.Menu mnuPriceLists 
         Caption         =   "ҩƷ��Ŀ��(&L)..."
         Shortcut        =   ^L
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
      Begin VB.Menu mnuViewToolbar 
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
      Begin VB.Menu mnuViewShowAll 
         Caption         =   "��ʾ�����¼�(&L)"
      End
      Begin VB.Menu mnuViewList 
         Caption         =   "��ʾͣ��Ŀ¼(&M)"
      End
      Begin VB.Menu mnuViewStoped 
         Caption         =   "��ʾͣ��ҩƷ(&C)"
      End
      Begin VB.Menu mnuViewPrices 
         Caption         =   "��ʾ��ʷ�۸�(&H)"
      End
      Begin VB.Menu mnuViewSpt2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "����(&F)..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewFindNext 
         Caption         =   "������һ��(&N)"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "����(&T)"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuViewRefer 
         Caption         =   "�ο�(&R)..."
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuViewSpt3 
         Caption         =   "-"
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
         Caption         =   "Web�ϵ�����(&W)"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "������̳(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&E)..."
         End
      End
      Begin VB.Menu mnuHelp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frmMediLists"
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
Dim intCount As Integer, intRow As Integer, intCol As Integer
Dim strTemp As String
Dim mintҩ�ⵥλ  As Integer
Dim mstrType As String              '���˴��ڷ��ص�ҩƷ�������ʹ�  '5-����ҩ 6-�г�ҩ 7-�в�ҩ
Dim mstrDrugId As String            '���˴��ڷ��ص�ҩ��ID��
Dim mlngCurrDrug As Long
Private mstrDBNodeClick As String
Private mstrNodeClick����ҩ As String     '��¼�ϴ�ѡ�еķ���
Private mstrNodeClick�г�ҩ As String
Private mstrNodeClick�в�ҩ As String
Private mstrNodeSelect����ҩ As String
Private mstrNodeSelect�г�ҩ As String
Private mstrNodeSelect�в�ҩ As String
Private mstrItemClick����ҩ As String     '��¼�ϴ�ѡ�е�ҩƷ
Private mstrItemClick�г�ҩ As String     '��¼�ϴ�ѡ�е�ҩƷ
Private mstrItemClick�в�ҩ As String     '��¼�ϴ�ѡ�е�ҩƷ
Private mstrKey As String           '��¼����ѡ�еķ���
Private Declare Function SetParent Lib "user32 " (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private mstrFindValue As String     '�����ַ���
Private mrsFind As ADODB.Recordset  '��¼��ѯ�����ݼ�
Private mstr���� As String          '��¼��ѡ���ʲô���͵�ҩƷ "0"-����ҩ,"1"-�г�ҩ,"2"-�в�ҩ
Private mintPage As Integer         '��¼��ǰ��ѡ�е�SSTABҳ��
Private mStrItem As String          'ѡ�е�Ʒ�ֽڵ�

Private Const colPriceҩ������ As Integer = 0
Private Const colPriceָ������ As Integer = 1
Private Const colPrice���� As Integer = 2
Private Const colPriceָ���ۼ� As Integer = 3
Private Const colPriceָ������ As Integer = 4
Private Const colPrice���ηѱ� As Integer = 5
Private Const colPriceִ����� As Integer = 6
Private Const colPrice����NO As Integer = 7
Private Const colPriceҩƷ As Integer = 8
Private Const colPrice��λ As Integer = 9
Private Const colPrice�ۼ� As Integer = 10
Private Const colPrice������Ŀ As Integer = 11
Private Const colPriceִ������ As Integer = 12
Private Const colPrice˵�� As Integer = 13
Private Const colPriceҩƷID As Integer = 14

Private Const colCostҩƷid As Integer = 0
Private Const colCostNO As Integer = 1
Private Const colCostҩƷ As Integer = 2
Private Const colCost�ⷿ As Integer = 3

Private Const colCost���� As Integer = 4
Private Const colCostЧ�� As Integer = 5
Private Const colCost���� As Integer = 6
Private Const colCost��λ As Integer = 7
Private Const colCostԭ�ɱ��� As Integer = 8
Private Const colCost�ɱ��� As Integer = 9
Private Const colCostִ������ As Integer = 10
Private Const colCost˵�� As Integer = 11

Private Const colҩƷ As Integer = 0
Private Const col�ⷿ As Integer = 1
Private Const col���� As Integer = 2
Private Const col���� As Integer = 3
Private Const col���� As Integer = 4
Private Const col���� As Integer = 5
Private Const col���� As Integer = 6
Private Const col���� As Integer = 7
Private Const col��λ As Integer = 8

Private mstrCostFormat As String
Private mstrPriceFormat As String
Private mintCostDigit As Integer        '�ɱ���С��λ��
Private mintPriceDigit As Integer       '�ۼ�С��λ��
Private mbln�Թ�ҩ As Boolean           '������¼�Ƿ���ͨ���Թ�ҩ���÷�ʽ�򿪵Ĵ���
Private mintIndex As Integer    '������¼������ķ���

Private Const mconColor_Stop As Long = &HFF&

Public Sub ShowMe(ByVal frmPar As Form, ByVal bln�Թ�ҩ As Boolean)
    '��ʾ����
    mbln�Թ�ҩ = bln�Թ�ҩ
    Me.Show , frmPar
End Sub

Private Sub GetCostAdjust(ByVal lngDrug As Long)
    Dim strSqlCon As String
    
    '----------��д�ɱ���-----------------
    On Error GoTo ErrHandle
    With Me.hgdCost
        .Rows = .FixedRows + 1: .RowData(.FixedRows) = 0
        For intCol = 0 To .Cols - 1
            .TextMatrix(.FixedRows, intCol) = ""
        Next
    End With
    
    If Me.mnuViewStoped.Checked = False Then
        strSqlCon = " and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))"
    End If
    
    gstrSql = " Select B.NO, I.ID As ҩƷid, '[' || I.���� || ']' || I.���� || ' ' || I.��� || ' ' || I.���� As ҩƷ, P.���� As �ⷿ,A.����,A.Ч��,A.����, " & _
            " I.���㵥λ As ��λ, S.ҩ�ⵥλ, Nvl(S.ҩ���װ, 1) ҩ���װ, A.ԭ�� As ԭ�ɱ���,A.�ּ� As �ɱ���, A.ִ������, A.����˵�� " & _
            " From ҩƷ�շ���¼ B, �շ���ĿĿ¼ I, ҩƷ��� S, ���ű� P, ҩƷ�۸��¼ A " & _
            " Where A.�۸�����=2 And A.�շ�id = B.ID(+) And A.ҩƷid = I.ID And " & _
            " I.ID = S.ҩƷid And A.�ⷿid = P.ID(+) And S.ҩ��id = [1] " & strSqlCon
    
    If chkStock.Value = 1 Then
        gstrSql = gstrSql & " And Exists (Select 1 From ҩƷ��� K Where ���� = 1 And k.�ⷿid = a.�ⷿid And k.ҩƷid = a.ҩƷid And k.���� = a.����) "
    End If
    
    gstrSql = gstrSql & " Order By ҩƷ, ִ������ Desc, NO Desc "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngDrug)
    
    With rsTemp
        Me.hgdCost.Redraw = False
        If .BOF Or .EOF Then
            With Me.hgdCost
                .Rows = .FixedRows + 1: .RowData(.FixedRows) = 0
                For intCol = 0 To .Cols - 1
                    .TextMatrix(.FixedRows, intCol) = ""
                Next
            End With
        Else
            Me.hgdCost.Rows = Me.hgdCost.FixedRows + .RecordCount
        End If
        Do While Not .EOF
            Me.hgdCost.RowData(.AbsolutePosition) = !ҩƷid
            Me.hgdCost.TextMatrix(.AbsolutePosition, colCostNO) = IIf(IsNull(!No), "", !No)
            Me.hgdCost.TextMatrix(.AbsolutePosition, colCostҩƷ) = !ҩƷ
            Me.hgdCost.TextMatrix(.AbsolutePosition, colCost�ⷿ) = IIf(IsNull(!�ⷿ), "", !�ⷿ)
            Me.hgdCost.TextMatrix(.AbsolutePosition, colCost����) = IIf(IsNull(!����), "", !����)
            Me.hgdCost.TextMatrix(.AbsolutePosition, colCostЧ��) = IIf(IsNull(!Ч��), "", Format(!Ч��, "yyyy-mm-dd"))
            Me.hgdCost.TextMatrix(.AbsolutePosition, colCost����) = IIf(IsNull(!����), "", !����)
            Me.hgdCost.TextMatrix(.AbsolutePosition, colCost��λ) = IIf(mintҩ�ⵥλ = 0, !��λ, !ҩ�ⵥλ)
            Me.hgdCost.TextMatrix(.AbsolutePosition, colCostԭ�ɱ���) = Format(!ԭ�ɱ��� * IIf(mintҩ�ⵥλ = 0, 1, !ҩ���װ), mstrCostFormat)
            Me.hgdCost.TextMatrix(.AbsolutePosition, colCost�ɱ���) = Format(!�ɱ��� * IIf(mintҩ�ⵥλ = 0, 1, !ҩ���װ), mstrCostFormat)
            Me.hgdCost.TextMatrix(.AbsolutePosition, colCostִ������) = IIf(IsNull(!ִ������), "", Format(!ִ������, "yyyy-mm-dd hh:mm:ss"))
            Me.hgdCost.TextMatrix(.AbsolutePosition, colCost˵��) = IIf(IsNull(!����˵��), "", !����˵��)
            Me.hgdCost.TextMatrix(.AbsolutePosition, colCostҩƷid) = !ҩƷid
            
            Me.hgdCost.Row = .AbsolutePosition
            For intCol = 0 To Me.hgdCost.Cols - 1
                Me.hgdCost.Col = intCol
                If IIf(IsNull(!ִ������), "", !ִ������) = "" Then
                    Me.hgdCost.CellBackColor = RGB(225, 255, 255)
                Else
                    Me.hgdCost.CellBackColor = RGB(240, 240, 240)
                End If
            Next
            .MoveNext
        Loop
        Me.hgdCost.Row = Me.hgdCost.FixedRows

        Me.hgdCost.Redraw = True
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetChargeSet(ByVal lngDrug As Long)
    Dim strSqlCon As String
    
    '----------��д�ɱ���-----------------
    On Error GoTo ErrHandle
    With Me.hgdCharge
        .Rows = .FixedRows + 1: .RowData(.FixedRows) = 0
        For intCol = 0 To .Cols - 1
            .TextMatrix(.FixedRows, intCol) = ""
        Next
    End With
    
    gstrSql = "Select B.ID, '[' || B.���� || ']' || B.���� As ����, A.�ѱ�, A.�κ�, Ӧ�ն���ֵ, Ӧ�ն�βֵ, ʵ�ձ���, Decode(���㷽��, 1, '1-�ɱ��ۼ��ձ�������', '0-�ֶα�������') As ���㷽�� " & _
        " From �ѱ���ϸ A, �շ���ĿĿ¼ B, ҩƷ��� C " & _
        " Where A.�շ�ϸĿid = B.ID And B.ID = C.ҩƷid And C.ҩ��id = [1] " & _
        " Order By ����, A.�ѱ�, A.�κ�, A.Ӧ�ն���ֵ"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngDrug)
    
    With rsTemp
        Me.hgdCharge.Redraw = False
        If .BOF Or .EOF Then
            With Me.hgdCharge
                .Rows = .FixedRows + 1: .RowData(.FixedRows) = 0
                For intCol = 0 To .Cols - 1
                    .TextMatrix(.FixedRows, intCol) = ""
                Next
            End With
        Else
            Me.hgdCharge.Rows = Me.hgdCharge.FixedRows + .RecordCount
        End If
        Do While Not .EOF
            Me.hgdCharge.RowData(.AbsolutePosition) = !ID
            Me.hgdCharge.TextMatrix(.AbsolutePosition, 0) = .Fields("����").Value
            Me.hgdCharge.TextMatrix(.AbsolutePosition, 1) = .Fields("�ѱ�").Value
            Me.hgdCharge.TextMatrix(.AbsolutePosition, 2) = Format(.Fields("Ӧ�ն���ֵ").Value, "##########0.00;-#########0.00;0.00;0.00") & _
                " �� " & Format(.Fields("Ӧ�ն�βֵ").Value, "##########0.00;-#########0.00;0.00;0.00")
            Me.hgdCharge.TextMatrix(.AbsolutePosition, 3) = Format(.Fields("ʵ�ձ���").Value, "###0.00;-##0.00;0.00;0.00")
            Me.hgdCharge.TextMatrix(.AbsolutePosition, 4) = .Fields("���㷽��").Value
            
            .MoveNext
        Loop
        
        Me.hgdCharge.ColAlignment(2) = 1
        Me.hgdCharge.ColAlignment(3) = 1
        Me.hgdCharge.MergeCells = flexMergeRestrictColumns
        Me.hgdCharge.MergeCol(0) = True
        Me.hgdCharge.MergeCol(1) = True
        
        Me.hgdCharge.Redraw = True
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Public Sub zlGetFilter(ByVal strType As String, ByVal strDrugId As String)
    mstrType = strType
    mstrDrugId = strDrugId
    Call cmdKind_Click(4)
End Sub

Private Sub zlPopupClassMenu()
    With tvwClass
        Select Case .Tag
        Case 0
            If InStr(1, mstrPrivs, ";��������ҩ;") = 0 Then
                mnuClassAdd.Visible = False
                mnuClassMod.Visible = False
                mnuClassDel.Visible = False
            Else
                mnuClassAdd.Visible = True
                mnuClassMod.Visible = True
                mnuClassDel.Visible = True
            End If
        Case 1
            If InStr(1, mstrPrivs, ";�����г�ҩ;") = 0 Then
                mnuClassAdd.Visible = False
                mnuClassMod.Visible = False
                mnuClassDel.Visible = False
            Else
                mnuClassAdd.Visible = True
                mnuClassMod.Visible = True
                mnuClassDel.Visible = True
            End If
        Case 2
            If InStr(1, mstrPrivs, ";�����в�ҩ;") = 0 Then
                mnuClassAdd.Visible = False
                mnuClassMod.Visible = False
                mnuClassDel.Visible = False
            Else
                mnuClassAdd.Visible = True
                mnuClassMod.Visible = True
                mnuClassDel.Visible = True
            End If
        End Select
    End With
    If InStr(1, mstrPrivs, ";ҩƷ����;") = 0 Then
        mnuClassStar.Visible = False
    Else
        mnuClassStar.Visible = True
    End If
    If InStr(1, mstrPrivs, ";ҩƷͣ��;") = 0 Then
        mnuClassStop.Visible = False
    Else
        mnuClassStop.Visible = True
    End If
                
    If mnuClassAdd.Visible = False And mnuClassMod.Visible = False And mnuClassDel.Visible = False And mnuClassStar.Visible = False And mnuClassStop.Visible = False Then
        Exit Sub
    End If
    Set objNode = Me.tvwClass.SelectedItem
    
    If objNode Is Nothing Then
        Me.mnuClassMod.Enabled = False
        Me.mnuClassDel.Enabled = False
        Me.mnuClassStar.Enabled = False
        Me.mnuClassStop.Enabled = False
    Else
        If Val(objNode.Tag) <= 2 Then
            If objNode.ForeColor = mconColor_Stop Then
                Me.mnuClassAdd.Enabled = False
                Me.mnuClassMod.Enabled = False
                Me.mnuClassStar.Enabled = True
                Me.mnuClassStop.Enabled = False
            Else
                Me.mnuClassAdd.Enabled = True
                Me.mnuClassMod.Enabled = True
                Me.mnuClassStar.Enabled = False
                Me.mnuClassStop.Enabled = True
            End If
        Else
            Me.mnuClassStar.Enabled = False
            Me.mnuClassStop.Enabled = False
        End If
    End If
    
    Call setMenu�Թ�ҩ
    Call PopupMenu(Me.mnuClass, 2)
End Sub

Private Sub chkStock_Click()
    Call GetCostAdjust(mlngCurrDrug)
End Sub

Private Sub cmdKind_Click(Index As Integer)
    Dim intCount As Integer
    Dim objNode As Node
    Dim strTemp As String
    Dim strItem As String
    
    mintIndex = Index
    mstrKey = ""
    mstrFindValue = ""
    Call SaveListViewState(Me.lvwItems, Me.Name & Val(Me.tvwClass.Tag), Me.lvwItems.View)
    For intCount = Me.cmdKind.LBound To Me.cmdKind.UBound
        If intCount <= Index Then
            Me.cmdKind(intCount).Tag = 0
        Else
            Me.cmdKind(intCount).Tag = 1
        End If
    Next
    
    mstr���� = Index & ""
    'װ���ݲ���������
    If Me.lvwItems.Visible Then
        Call picClass_Resize
        Me.tvwClass.SetFocus
    End If
    If Index < 3 Then
        If Val(tvwClass.Tag) <> Index Then
            Me.tvwClass.Tag = Index
            Call zlRefClasses
        End If
        Me.mnuViewFind.Enabled = True
        Me.mnuViewFindNext.Enabled = True
        Me.tlbThis.Buttons("Find").Enabled = True
        Me.mnuEditExcel.Enabled = True
    Else
        Me.tvwClass.Tag = Index
        Call zlRefClasses
        Me.mnuViewFind.Enabled = False
        Me.mnuViewFindNext.Enabled = False
        Me.tlbThis.Buttons("Find").Enabled = False
        Me.mnuEditExcel.Enabled = False
        frmMediFind.Hide
    End If
    If Val(tvwClass.Tag) >= 3 Then
        txtFind.Enabled = False
        txtFind.BackColor = &H8000000F  '��ɫ������
    Else
        txtFind.Enabled = True
        txtFind.BackColor = vbWhite '��ɫ�����޸�
    End If
    
    If mstr���� = "0" Then
        strTemp = mstrNodeSelect����ҩ
    ElseIf mstr���� = "1" Then
        strTemp = mstrNodeSelect�г�ҩ
    ElseIf mstr���� = "2" Then
        strTemp = mstrNodeSelect�в�ҩ
    End If
    
    For Each objNode In tvwClass.Nodes
        If objNode.Key = strTemp Then
            objNode.Selected = True
            Call tvwClass_NodeClick(objNode)
            Exit For
        End If
    Next
End Sub

Private Sub clbThis_Resize()
    Me.clbThis.Bands(1).MinHeight = Me.tlbThis.Height
    Me.clbThis.Refresh
    Call Form_Resize
End Sub

Private Sub Form_Activate()
    Me.lvwItems.Visible = True
    Call Form_Resize
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    Dim i As Long
    Dim lngCount As Long
    Dim rs������Ŀ As ADODB.Recordset
    Dim bln������Ŀ As Boolean
    
    mlngMode = glngModul
    mstrPrivs = gstrPrivs
    
    gblnIncomeItem = False
    '����ָ�
    Call RestoreWinState(Me, App.ProductName)
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    If mbln�Թ�ҩ = True Then '�Թ�ҩ������ģ����ȡ����ƥ���ֵ���Թ�ҩ��������
        gstrMatch = IIf(Val(zlDatabase.GetPara("����ƥ��", , , True)) = 0, "%", "")
    End If
    
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
    
    mnuViewShowAll.Checked = (Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��ʾ����", 0)) = 1)
    mnuViewList.Checked = (Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��ʾͣ��Ŀ¼", 0)) = 1)
    mnuViewStoped.Checked = (Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��ʾͣ����Ŀ", 0)) = 1)
    mnuViewPrices.Checked = (Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��ʾ��ʷ�۸�", 0)) = 1)
    
    '����Ƿ���ҩ�ⵥλ��ʾ�۸�
    mintҩ�ⵥλ = Val(zlDatabase.GetPara(29, glngSys))
    
    mintCostDigit = GetDigit(1, 1, IIf(mintҩ�ⵥλ = 0, 1, 4))
    mintPriceDigit = GetDigit(1, 2, IIf(mintҩ�ⵥλ = 0, 1, 4))
    
    mstrCostFormat = "0." & String(mintCostDigit, "0") & ";-0." & String(mintCostDigit, "0") & ";0"
    mstrPriceFormat = "0." & String(mintPriceDigit, "0") & ";-0." & String(mintPriceDigit, "0") & ";0"
    
    
    '��ֱ��ͨ���˵����е�Ȩ�޿���
'    If InStr(1, mstrPrivs, "��������") = 0 Then Me.mnuFilePara.Visible = False: Me.mnuFileSpt2.Visible = False
    If InStr(1, mstrPrivs, "Э��ҩƷ����") = 0 Then Me.mnuEditSpecProtocol.Visible = False:
    
    If InStr(1, mstrPrivs, "ҩƷͣ��") = 0 Then Me.mnuEditStop.Visible = False:
    If InStr(1, mstrPrivs, "ҩƷ����") = 0 Then Me.mnuEditStart.Visible = False:
    If InStr(1, mstrPrivs, "�������÷�ҩ����") = 0 Then Me.mnuEditSendType.Visible = False
    Me.mnuEditSpt2.Visible = (Me.mnuEditStop.Visible Or Me.mnuEditStart.Visible)
    
    tlbThis.Buttons("Stop").Visible = Me.mnuEditStop.Visible
    tlbThis.Buttons("Start").Visible = Me.mnuEditStart.Visible
    tlbThis.Buttons(7).Visible = Me.mnuEditSpt2.Visible
    
    '�ۼ۱������
    With Me.hgdPrice
        .Redraw = False
        .Rows = .FixedRows + 1: .Cols = 15
        
        .TextMatrix(0, colPriceҩƷ) = "ҩƷ": .TextMatrix(0, colPrice��λ) = "��λ": .TextMatrix(0, colPrice�ۼ�) = "�ۼ�"
        .TextMatrix(0, colPrice������Ŀ) = "������Ŀ": .TextMatrix(0, colPrice˵��) = "˵��": .TextMatrix(0, colPriceִ������) = "ִ������"
        .TextMatrix(0, colPriceҩƷID) = "ҩƷID"
        .TextMatrix(0, colPrice����NO) = "���۵��ݺ�"
        
        .ColWidth(colPriceҩ������) = 0: .ColWidth(colPriceָ������) = 0: .ColWidth(colPrice����) = 0: .ColWidth(colPriceָ���ۼ�) = 0
        .ColWidth(colPriceָ������) = 0: .ColWidth(colPrice���ηѱ�) = 0: .ColWidth(colPriceִ�����) = 0
        .ColWidth(colPriceҩƷ) = 3500: .ColWidth(colPrice��λ) = 550: .ColWidth(colPrice�ۼ�) = 1000
        .ColWidth(colPrice������Ŀ) = 850: .ColWidth(colPrice˵��) = 2500: .ColWidth(colPriceִ������) = 1800
        .ColWidth(colPriceҩƷID) = 0
        .ColWidth(colPrice����NO) = 1000
        
        For intCol = 0 To .Cols - 1
            .ColAlignmentFixed(intCol) = 4
        Next
        .ColAlignment(colPriceҩƷ) = 1: .ColAlignment(colPrice��λ) = 4: .ColAlignment(colPrice�ۼ�) = 7
        .ColAlignment(colPrice������Ŀ) = 1: .ColAlignment(colPrice˵��) = 1: .ColAlignment(colPriceִ������) = 1
        .Redraw = True
    End With
    
    '�ɱ��۱������
    With Me.hgdCost
        .Redraw = False
        .Rows = .FixedRows + 1
        .Cols = 12
        
        .TextMatrix(0, colCostҩƷid) = "ҩƷid"
        .TextMatrix(0, colCostNO) = "����NO"
        .TextMatrix(0, colCostҩƷ) = "ҩƷ"
        .TextMatrix(0, colCost�ⷿ) = "�ⷿ"
        .TextMatrix(0, colCost����) = "����"
        .TextMatrix(0, colCostЧ��) = "Ч��"
        .TextMatrix(0, colCost����) = "����"
        .TextMatrix(0, colCost��λ) = "��λ"
        .TextMatrix(0, colCostԭ�ɱ���) = "ԭ�ɱ���"
        .TextMatrix(0, colCost�ɱ���) = "�³ɱ���"
        .TextMatrix(0, colCostִ������) = "ִ������"
        .TextMatrix(0, colCost˵��) = "˵��"
        
        .ColWidth(colCostҩƷid) = 0
        .ColWidth(colCostNO) = 1000
        .ColWidth(colCostҩƷ) = 3500
        .ColWidth(colCost�ⷿ) = 1500
        .ColWidth(colCost����) = 1000
        .ColWidth(colCostЧ��) = 1000
        .ColWidth(colCost����) = 1000
        .ColWidth(colCost��λ) = 550
        .ColWidth(colCostԭ�ɱ���) = 1000
        .ColWidth(colCost�ɱ���) = 1000
        .ColWidth(colCostִ������) = 1800
        .ColWidth(colCost˵��) = 2500
        
        For intCol = 0 To .Cols - 1
            .ColAlignmentFixed(intCol) = 4
        Next
        
        .ColAlignment(colCostNO) = 1
        .ColAlignment(colCostҩƷ) = 1
        .ColAlignment(colCost�ⷿ) = 1
        .ColAlignment(colCost����) = 1
        .ColAlignment(colCostЧ��) = 1
        .ColAlignment(colCost����) = 1
        .ColAlignment(colCost��λ) = 4
        .ColAlignment(colCostԭ�ɱ���) = 7
        .ColAlignment(colCost�ɱ���) = 7
        .ColAlignment(colCostִ������) = 1
        .ColAlignment(colCost˵��) = 1
        
        .Redraw = True
    End With
    
    '�ѱ�ȼ��б�����
    With hgdCharge
        .Cols = 5
        .ColWidth(0) = 4000
        .ColWidth(1) = 1500
        .ColWidth(2) = 3000
        .ColWidth(3) = 1050
        .ColWidth(4) = 2000
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .TextMatrix(0, 0) = "ҩƷ"
        .TextMatrix(0, 1) = "�ѱ�"
        .TextMatrix(0, 2) = "Ӧ�ս��(Ԫ)"
        .TextMatrix(0, 3) = "ʵ�ձ���(%)"
        .TextMatrix(0, 4) = "���㷽��"
        
        .MergeCol(0) = True
'        .MergeCol(1) = True
    End With
    
    Me.picHBar.Top = Me.ScaleHeight - IIf(Me.stbThis.Visible, Me.stbThis.Height, 0) - 2500
    Call cmdKind_Click(0)
    
    If mbln�Թ�ҩ = False Then
        '����ƽ̨�ӿ�
        On Error Resume Next
        LogisticPlatformInterface
    End If
    
    '�Ƿ����ÿ���ҩ���ϸ����
    gblnKSSStrict = CheckKSSPrivilege
    
    If gblnKSSStrict = False Then
        lngCount = 0
        For i = 0 To Me.mnuReportItem.UBound
            If Trim(Me.mnuReportItem(i).Tag) <> "" Then
                If Split(Me.mnuReportItem(i).Tag, ",")(1) = "ZL1_INSIDE_1261_2" Or Split(Me.mnuReportItem(i).Tag, ",")(1) = "ZL1_INSIDE_1261_3" Then
                    lngCount = lngCount + 1
                    If lngCount = Me.mnuReportItem.Count Then
                        Me.mnuReport.Visible = False
                    Else
                        Me.mnuReportItem(i).Visible = False
                    End If
                End If
            End If
        Next
    End If
    
    If mstrPrivs Like "*����ҩ*" Then
        bln������Ŀ = IIf(zlDatabase.GetPara("����ҩ������Ŀ", 100, 1023) = 0, True, False)
    End If
    If mstrPrivs Like "*�г�ҩ*" And bln������Ŀ = False Then
        bln������Ŀ = IIf(zlDatabase.GetPara("�г�ҩ������Ŀ", 100, 1023) = 0, True, False)
    End If
    If mstrPrivs Like "*�г�ҩ*" And bln������Ŀ = False Then
        bln������Ŀ = IIf(zlDatabase.GetPara("�в�ҩ������Ŀ", 100, 1023) = 0, True, False)
    End If
    If bln������Ŀ = True Then
        'ģ�鹫�������Ѿ�������ҩƷ��������ģ�飬Ŀǰû��˽�л򱾻���������ʱ���β������ý���
        MsgBox "�뵽ҩƷ��������ģ�����ø����ʶ�Ӧ��������Ŀ��", vbInformation, gstrSysName
'        frmMediPara.ShowMe mstrPrivs, Me
        If gblnIncomeItem = False Then
            Unload Me
        End If
    End If
    
    If mbln�Թ�ҩ = True Then
        tabContent.TabVisible(1) = False
        tabContent.TabVisible(2) = False
        tabContent.TabVisible(3) = False
    End If
End Sub



Private Sub LogisticPlatformInterface()
'����ƽ̨�ӿ�
    
    If gobjLogisticPlatform Is Nothing Then
        On Error Resume Next
        Set gobjLogisticPlatform = CreateObject("zlDrugPurchase.clsDrugPurchase")
        If err <> 0 Then
            mnuUploadDrugInfo.Visible = False
            err.Clear: On Error GoTo 0
            Exit Sub
        End If
        
    End If
    
    If mnuEditSptPacker.Visible = False Then mnuEditSptPacker.Visible = True
    mnuUploadDrugInfo.Visible = True
       
End Sub

Private Sub Form_Resize()
    Dim lngTools As Single, lngStatus As Single
    
    If WindowState = 1 Then Exit Sub
    lngTools = IIf(Me.clbThis.Visible, Me.clbThis.Height, 0)
    lngStatus = IIf(Me.stbThis.Visible, Me.stbThis.Height, 0)
    
    err = 0: On Error Resume Next
    
    With Me.picVBar
        .Top = lngTools
        .Height = Me.ScaleHeight - lngTools - lngStatus
        If .Left < 2000 Then .Left = 2000
        If .Left > Me.ScaleWidth - 4000 Then .Left = Me.ScaleWidth - 4000
    End With
    With Me.picHBar
        .Left = Me.picVBar.Left + Me.picVBar.Width
        .Width = Me.ScaleWidth - .Left
        If .Top < 2000 Then .Top = 2000
        If .Top > Me.ScaleHeight - lngStatus - 2500 Then .Top = Me.ScaleHeight - lngStatus - 2500
    End With
    With Me.picClass
        .Left = Me.ScaleLeft
        .Top = lngTools
        .Height = Me.ScaleHeight - picClass.Top - lngStatus
        .Width = Me.picVBar.Left - Me.picClass.Left
    End With
    
    With Me.lvwItems
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
    
    With Me.fraComment(0)
        .Left = 90
        .Width = Me.tabContent.Width - .Left * 2
        .Top = Me.tabContent.Height - .Height - 50 '- 90
    End With
    With Me.fraComment(1)
        .Left = 90
        .Width = Me.tabContent.Width - .Left * 2
        .Top = Me.tabContent.Height - .Height - 60
    End With
    
    With Me.lvwSpecs
        .Left = 90
        .Top = 395
        .Width = Me.tabContent.Width - .Left * 2
        If lblComment(0).Caption = "" Then
            lvwSpecs.Height = tabContent.Height - lvwSpecs.Top - 50
        Else
            lvwSpecs.Height = tabContent.Height - lvwSpecs.Top - 50 - fraComment(0).Height
        End If
    End With
    With Me.hgdPrice
        .Left = 90
        .Top = 395
        .Width = Me.tabContent.Width - .Left * 2
        If lblComment(3).Caption = "" Then
            hgdPrice.Height = tabContent.Height - hgdPrice.Top - 50
        Else
            hgdPrice.Height = tabContent.Height - hgdPrice.Top - 350 - lblComment(3).Height
        End If
    End With
    
    With Me.hgdCost
        .Left = 90
        .Top = 395
        .Width = Me.tabContent.Width - .Left * 2
        .Height = Me.tabContent.Height - .Top - 50
    End With
    
    With Me.hgdCharge
        .Left = 90
        .Top = 395
        .Width = Me.tabContent.Width - .Left * 2
        .Height = Me.tabContent.Height - .Top - 50
    End With
    
    SetParent txtFind.hwnd, tlbThis.hwnd
    SetParent picFind.hwnd, tlbThis.hwnd
    txtFind.Left = Me.ScaleWidth - txtFind.Width - 200
    picFind.Left = txtFind.Left - 100 - picFind.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    Call SaveListViewState(Me.lvwItems, Me.Name & Val(Me.tvwClass.Tag), Me.lvwItems.View)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\�ָ�", "����", Me.picVBar.Left)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\�ָ�", "����", Me.picHBar.Top)
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��ʾ����", IIf(mnuViewShowAll.Checked, 1, 0)
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��ʾͣ��Ŀ¼", IIf(mnuViewList.Checked, 1, 0)
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��ʾͣ����Ŀ", IIf(mnuViewStoped.Checked, 1, 0)
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��ʾ��ʷ�۸�", IIf(mnuViewPrices.Checked, 1, 0)
    mstrNodeClick����ҩ = ""
    mstrNodeClick�г�ҩ = ""
    mstrNodeClick�в�ҩ = ""
    mstrItemClick����ҩ = ""
    mstrItemClick�г�ҩ = ""
    mstrItemClick�в�ҩ = ""
    mstrNodeSelect����ҩ = ""
    mstrNodeSelect�г�ҩ = ""
    mstrNodeSelect�в�ҩ = ""
    mstrKey = ""
    mstrFindValue = ""
    Set mrsFind = Nothing
End Sub

Private Sub hgdCharge_DblClick()
    Dim strCharge As String
    On Error GoTo ErrHandle
    
    If InStr(mstrPrivs, "�ѱ�����") = 0 Then Exit Sub
    
    If Me.hgdCharge.Rows > 1 Then
        If Me.hgdCharge.TextMatrix(Me.hgdCharge.Rows - 1, 1) <> "" Then
            strCharge = Me.hgdCharge.TextMatrix(Me.hgdCharge.Row, 1)
        End If
    End If
    If mnuPriceChargeSet.Enabled = True Then
'        frmChargeSortItemEdit.ShowMe Me, 3, strCharge, Val(hgdCharge.RowData(hgdCharge.Row)), hgdCharge.TextMatrix(Me.hgdCharge.Row, 0)
        frmSetExpense.ShowMe Me, Val(hgdCharge.RowData(hgdCharge.Row)), hgdCharge.TextMatrix(Me.hgdCharge.Row, 0)
        
        If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
        Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub hgdCost_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    Call PopupMenu(Me.mnuPrice, 2)
End Sub


Private Sub hgdPrice_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    
    Call PopupMenu(Me.mnuPrice, 2)
End Sub

Private Sub hgdPrice_RowColChange()
    Dim bln�б� As Boolean
    Dim rsCheck As New ADODB.Recordset
    If Val(hgdPrice.RowData(hgdPrice.Row)) = 0 Then Exit Sub
    
    On Error GoTo ErrHandle
    gstrSql = "Select Nvl(�б�ҩƷ,0) �б�ҩƷ From ҩƷ��� Where ҩƷID=(Select �շ�ϸĿID From �շѼ�Ŀ Where ID=[1]" & _
            GetPriceClassString("") & ")"
    
    Set rsCheck = zlDatabase.OpenSQLRecord(gstrSql, "�жϵ�ǰҩƷ�Ƿ�Ϊ�б�ҩƷ", Val(hgdPrice.RowData(hgdPrice.Row)))
    
    bln�б� = (rsCheck!�б�ҩƷ = 1)
    
    With Me.hgdPrice
        Me.lblComment(3).Caption = "1��" & _
            .TextMatrix(.Row, colPriceҩ������) & "ҩƷ��" & _
            IIf(bln�б�, "�б�۸�", "ָ������") & .TextMatrix(.Row, colPriceָ������) & "Ԫ/" & .TextMatrix(.Row, colPrice��λ) & "��" & _
            "�ɹ�����" & .TextMatrix(.Row, colPrice����) & "%��"
        Me.lblComment(4).Caption = "2��" & _
            "ָ���ۼ�" & .TextMatrix(.Row, colPriceָ���ۼ�) & "Ԫ/" & .TextMatrix(.Row, colPrice��λ) & "��" & _
            "ָ������" & .TextMatrix(.Row, colPriceָ������) & "%��" & _
            IIf(Val(.TextMatrix(.Row, colPrice���ηѱ�)) = 0, "���ݲ�����ݷѱ�����Żݻ�Ӽۡ�", "���ܲ�����ݷѱ�Ӱ�졣")
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwItems.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwItems.SortOrder = IIf(Me.lvwItems.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwItems.SortKey = ColumnHeader.Index - 1
        Me.lvwItems.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwItems_DblClick()
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    If Val(Me.tvwClass.Tag) < 2 Then
        With frmMediItem
            .Tag = IIf(Me.tvwClass.Tag = 0, 1, 2)
            .cmdCancel.Tag = "����"
            .lng����id = Mid(Me.tvwClass.SelectedItem.Key, 2)
            .lngҩ��id = Mid(Me.lvwItems.SelectedItem.Key, 2)
            .Show 1, Me
        End With
    ElseIf Val(Me.tvwClass.Tag) = 2 Then
        With frmMediHerbalItem
            .Tag = 3
            .cmdCancel.Tag = "����"
            .lng����id = Mid(Me.tvwClass.SelectedItem.Key, 2)
            .lngҩ��id = Mid(Me.lvwItems.SelectedItem.Key, 2)
            .strPrivs = Me.mstrPrivs
            .Show 1, Me
        End With
    ElseIf Val(Me.tvwClass.Tag) = 4 Then
        Set objItem = Me.lvwItems.SelectedItem
        If objItem Is Nothing Then
            Exit Sub
        End If
        If mstrType <> "7" Then
             With frmMediItem
                .Tag = objItem.SubItems(Me.lvwItems.ColumnHeaders("_����").Index - 1)
                .cmdCancel.Tag = "����"
                .lng����id = objItem.SubItems(Me.lvwItems.ColumnHeaders("_����id").Index - 1)
                .lngҩ��id = Mid(Me.lvwItems.SelectedItem.Key, 2)
                .strPrivs = Me.mstrPrivs
                .Show 1, Me
            End With
        Else
            With frmMediHerbalItem
                .Tag = objItem.SubItems(Me.lvwItems.ColumnHeaders("_����").Index - 1)
                .cmdCancel.Tag = "����"
                .lng����id = objItem.SubItems(Me.lvwItems.ColumnHeaders("_����id").Index - 1)
                .lngҩ��id = Mid(Me.lvwItems.SelectedItem.Key, 2)
                .strPrivs = Me.mstrPrivs
                .Show 1, Me
            End With
        End If
    End If
End Sub

Private Sub lvwItems_GotFocus()
    Set objItem = Me.lvwItems.SelectedItem
    If objItem Is Nothing Then
        Exit Sub
    End If
        
    If lvwItems.ListItems.Count = 0 Then Exit Sub
    If Val(Me.tvwClass.Tag) = 0 And InStr(1, mstrPrivs, "��������ҩ") = 0 Then Exit Sub
    If Val(Me.tvwClass.Tag) = 1 And InStr(1, mstrPrivs, "�����г�ҩ") = 0 Then Exit Sub
    If Val(Me.tvwClass.Tag) = 2 And InStr(1, mstrPrivs, "�����в�ҩ") = 0 Then Exit Sub
    If Val(Me.tvwClass.Tag) = 4 And InStr(1, mstrPrivs, "��������ҩ") = 0 _
        And InStr(1, mstrPrivs, "�����г�ҩ") = 0 _
        And InStr(1, mstrPrivs, "�����в�ҩ") = 0 Then
        Exit Sub
    End If
    
    '����ҩƷ��Ƭ�����á�ͣ�ñ�־
    If lvwItems.SelectedItem.Icon = "��ҩS" Or lvwItems.SelectedItem.Icon = "��ҩS" Then
        If mnuEditStart.Visible = True Then mnuEditStart.Enabled = True
        mnuEditStop.Enabled = False
        tlbThis.Buttons("Start").Enabled = mnuEditStart.Enabled
        tlbThis.Buttons("Stop").Enabled = False
    Else
        mnuEditStart.Enabled = False
        If mnuEditStop.Visible = True Then mnuEditStop.Enabled = True
        tlbThis.Buttons("Start").Enabled = False
        tlbThis.Buttons("Stop").Enabled = mnuEditStop.Enabled
    End If
    
    If Val(Me.tvwClass.Tag) = 0 Then
        If InStr(1, mstrPrivs, "��������ҩƷ��") = 0 Then
            Me.mnuEditItemAdd.Enabled = False
            Me.mnuEditItemMod.Enabled = False
            Me.mnuEditItemDel.Enabled = False
            mnuEditStart.Enabled = False
            mnuEditStop.Enabled = False
            tlbThis.Buttons("Start").Enabled = False
            tlbThis.Buttons("Stop").Enabled = False
        End If
    ElseIf Val(Me.tvwClass.Tag) = 1 Then
        If InStr(1, mstrPrivs, "�����г�ҩƷ��") = 0 Then
            Me.mnuEditItemAdd.Enabled = False
            Me.mnuEditItemMod.Enabled = False
            Me.mnuEditItemDel.Enabled = False
            mnuEditStart.Enabled = False
            mnuEditStop.Enabled = False
            tlbThis.Buttons("Start").Enabled = False
            tlbThis.Buttons("Stop").Enabled = False
        End If
    ElseIf Val(Me.tvwClass.Tag) = 2 Then
        If InStr(1, mstrPrivs, "�����в�ҩƷ��") = 0 Then
            Me.mnuEditItemAdd.Enabled = False
            Me.mnuEditItemMod.Enabled = False
            Me.mnuEditItemDel.Enabled = False
            mnuEditStart.Enabled = False
            mnuEditStop.Enabled = False
            tlbThis.Buttons("Start").Enabled = False
            tlbThis.Buttons("Stop").Enabled = False
        End If
    ElseIf Val(Me.tvwClass.Tag) = 4 Then
        If objItem.SubItems(Me.lvwItems.ColumnHeaders("_����").Index - 1) = "1" Then
            If InStr(1, mstrPrivs, "��������ҩƷ��") = 0 Then
                Me.mnuEditItemAdd.Enabled = False
                Me.mnuEditItemMod.Enabled = False
                Me.mnuEditItemDel.Enabled = False
                mnuEditStart.Enabled = False
                mnuEditStop.Enabled = False
                tlbThis.Buttons("Start").Enabled = False
                tlbThis.Buttons("Stop").Enabled = False
            End If
        ElseIf objItem.SubItems(Me.lvwItems.ColumnHeaders("_����").Index - 1) = "2" Then
            If InStr(1, mstrPrivs, "�����г�ҩƷ��") = 0 Then
                Me.mnuEditItemAdd.Enabled = False
                Me.mnuEditItemMod.Enabled = False
                Me.mnuEditItemDel.Enabled = False
                mnuEditStart.Enabled = False
                mnuEditStop.Enabled = False
                tlbThis.Buttons("Start").Enabled = False
                tlbThis.Buttons("Stop").Enabled = False
            End If
        ElseIf objItem.SubItems(Me.lvwItems.ColumnHeaders("_����").Index - 1) = "3" Then
            If InStr(1, mstrPrivs, "�����в�ҩƷ��") = 0 Then
                Me.mnuEditItemAdd.Enabled = False
                Me.mnuEditItemMod.Enabled = False
                Me.mnuEditItemDel.Enabled = False
                mnuEditStart.Enabled = False
                mnuEditStop.Enabled = False
                tlbThis.Buttons("Start").Enabled = False
                tlbThis.Buttons("Stop").Enabled = False
            End If
        End If
    End If
    
End Sub

Private Sub lvwItems_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim strSqlCon As String
    Dim strCaption As String
    Dim str�ۼۼ�¼ As String
    
    err = 0: On Error GoTo ErrHand
    strCaption = lblComment(0).Caption
    str�ۼۼ�¼ = lblComment(3).Caption
    If Item.Index <> 1 Then '��һ����¼Ĭ����ѡ�еĲ����ٴ�ѡ����
        If mstr���� = "0" Then
            mstrItemClick����ҩ = Item.Key
            mstrNodeClick����ҩ = tvwClass.SelectedItem.Key
        ElseIf mstr���� = "1" Then
            mstrItemClick�г�ҩ = Item.Key
            mstrNodeClick�г�ҩ = tvwClass.SelectedItem.Key
        ElseIf mstr���� = "2" Then
            mstrItemClick�в�ҩ = Item.Key
            mstrNodeClick�в�ҩ = tvwClass.SelectedItem.Key
        End If
    End If
    
    '----------��д���-----------------
   
    gstrSql = "select Distinct I.ID,I.����,I.���,I.���� as ����,S.ԭ����,N.���� as ��Ʒ��,I.�������� as ҽ������,S.ҩƷ��Դ," & _
            "        decode(I.�������,1,'����',2,'סԺ',3,'�����סԺ','��ֱ��Ӧ���ڲ���') as �������," & _
            "        decode(S.����ҩƷ,1,'��',' ') as ����,decode(S.Э��ҩƷ,1,'��',' ') Э��," & _
            "        decode(S.�б�ҩƷ,1,'��',' ') �б�,decode(S.��ҩ��̬,1,'��ҩ��Ƭ',2,'����','ɢװ') ��ҩ��̬," & _
            "        S.��׼�ĺ�,Nvl(I.�Ƿ���,0) �Ƿ���,Nvl(S.�б�ҩƷ,0) �б�ҩƷ," & _
            "        nvl(I.����ʱ��,to_date('3000-01-01','YYYY-MM-DD')) as ����ʱ��,G.���� ��ͬ��λ,I.˵��,I.��ѡ��,I.վ�� " & _
            " from �շ���ĿĿ¼ I,ҩƷ��� S,�շ���Ŀ���� N,(Select Id,���� From ��Ӧ�� Where ĩ�� = 1 And substr(����,1,1) = '1' And " & _
            " (����ʱ�� Is Null Or ����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) ) G " & _
            " where I.ID=S.ҩƷID and G.id(+)=S.��ͬ��λid and I.ID=N.�շ�ϸĿID(+) And N.����(+) = 3 and S.ҩ��ID=[1] "
'            " where (I.վ�� = '" & gstrNodeNo & "' Or I.վ�� is Null) And I.ID=S.ҩƷID and G.id(+)=S.��ͬ��λid and I.ID=N.�շ�ϸĿID(+) And N.����(+) = 3 and S.ҩ��ID=[1] "
    If Me.mnuViewStoped.Checked = False Then
        gstrSql = gstrSql & " and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(Item.Key, 2)))
    
    With rsTemp
        Me.lvwSpecs.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwSpecs.ListItems.Add(, "_" & !ID, IIf(IsNull(!���), "", !���))
            
            objItem.SubItems(Me.lvwSpecs.ColumnHeaders("����").Index - 1) = !����
            objItem.SubItems(Me.lvwSpecs.ColumnHeaders("����").Index - 1) = IIf(IsNull(!����), "", !����)
            If Val(Me.tvwClass.Tag) = 2 Or Val(Me.tvwClass.Tag) = 4 And mstrType = "7" Then
            objItem.SubItems(Me.lvwSpecs.ColumnHeaders("ԭ����").Index - 1) = IIf(IsNull(!ԭ����), "", !ԭ����)
            End If
            If Not (Val(Me.tvwClass.Tag) = 2 Or Val(Me.tvwClass.Tag) = 4 And mstrType = "7") Then
                objItem.SubItems(Me.lvwSpecs.ColumnHeaders("��Ʒ��").Index - 1) = IIf(IsNull(!��Ʒ��), "", !��Ʒ��)
            End If
            objItem.SubItems(Me.lvwSpecs.ColumnHeaders("ҽ������").Index - 1) = IIf(IsNull(!ҽ������), "", !ҽ������)
            objItem.SubItems(Me.lvwSpecs.ColumnHeaders("ҩƷ��Դ").Index - 1) = IIf(IsNull(!ҩƷ��Դ), "", !ҩƷ��Դ)
            objItem.SubItems(Me.lvwSpecs.ColumnHeaders("����").Index - 1) = IIf(IsNull(!����), "", !����)
            objItem.SubItems(Me.lvwSpecs.ColumnHeaders("Э��").Index - 1) = IIf(IsNull(!Э��), "", !Э��)
            objItem.SubItems(Me.lvwSpecs.ColumnHeaders("�б�").Index - 1) = IIf(IsNull(!�б�), "", !�б�)
            If Val(Me.tvwClass.Tag) = 2 Or Val(Me.tvwClass.Tag) = 4 And mstrType = "7" Then
                objItem.SubItems(Me.lvwSpecs.ColumnHeaders("��ҩ��̬").Index - 1) = IIf(IsNull(!��ҩ��̬), "", !��ҩ��̬)
            End If
            objItem.SubItems(Me.lvwSpecs.ColumnHeaders("��׼�ĺ�").Index - 1) = IIf(IsNull(!��׼�ĺ�), "", !��׼�ĺ�)
            objItem.SubItems(Me.lvwSpecs.ColumnHeaders("��ͬ��λ").Index - 1) = IIf(IsNull(!��ͬ��λ), "", !��ͬ��λ)
            objItem.SubItems(Me.lvwSpecs.ColumnHeaders("˵��").Index - 1) = IIf(IsNull(!˵��), "", !˵��)
            objItem.SubItems(Me.lvwSpecs.ColumnHeaders("��ѡ��").Index - 1) = IIf(IsNull(!��ѡ��), "", !��ѡ��)
            objItem.SubItems(Me.lvwSpecs.ColumnHeaders("վ��").Index - 1) = IIf(IsNull(!վ��), "", !վ��)
            'If Val(Me.tvwClass.Tag) < 2 Or (Val(Me.tvwClass.Tag) = 3 And mstrType <> "7") Then
                objItem.SubItems(Me.lvwSpecs.ColumnHeaders("�������").Index - 1) = IIf(IsNull(!�������), "", !�������)
            'End If
            
            If Format(!����ʱ��, "YYYY-MM-DD") = "3000-01-01" Then
                If Val(Me.tvwClass.Tag) = 2 Or Val(Me.tvwClass.Tag) = 4 And mstrType = "7" Then
                    objItem.Icon = "�ݹ�U": objItem.SmallIcon = "�ݹ�U"
                Else
                    objItem.Icon = "���U": objItem.SmallIcon = "���U"
                End If
            Else
                If Val(Me.tvwClass.Tag) = 2 Or Val(Me.tvwClass.Tag) = 4 And mstrType = "7" Then
                    objItem.Icon = "�ݹ�S": objItem.SmallIcon = "�ݹ�S"
                Else
                    objItem.Icon = "���S": objItem.SmallIcon = "���S"
                End If
                
                objItem.ForeColor = mconColor_Stop
                For intCount = 1 To Me.lvwSpecs.ColumnHeaders.Count - 1
                    objItem.ListSubItems(intCount).ForeColor = mconColor_Stop
                Next
            End If

            '������б�ҩƷ������ɫ�����Ƿ���ʱ�ۻ��Ƕ���ҩƷ
            If !�б�ҩƷ = 1 Then
                objItem.ListSubItems(1).ForeColor = IIf(!�Ƿ��� = 0, &H800000, &H800080)
            Else
                objItem.ListSubItems(1).ForeColor = IIf(!�Ƿ��� = 0, &H0, &H40&)
            End If
            .MoveNext
        Loop
    End With
    If Me.lvwSpecs.ListItems.Count > 0 Then
        If Me.lvwSpecs.SelectedItem Is Nothing Then Me.lvwSpecs.ListItems(1).Selected = True
        Call lvwSpecs_ItemClick(Me.lvwSpecs.SelectedItem)
        mnuEditSpecUnit.Enabled = True
    Else
        mnuEditSpecUnit.Enabled = False
        For intCount = Me.lblComment.LBound To Me.lblComment.UBound
            Me.lblComment(intCount).Caption = ""
        Next
    End If
    
    '----------��д�ۼ�-----------------
    With Me.hgdPrice
        .Rows = .FixedRows + 1: .RowData(.FixedRows) = 0
        For intCol = 0 To .Cols - 1
            .TextMatrix(.FixedRows, intCol) = ""
        Next
    End With
    
    gstrSql = "select P.ID,decode(I.�Ƿ���,1,'ʱ��','����') as ҩ������,nvl(S.ָ��������,0) as ָ������,nvl(S.����,0) as ����," & _
            "        nvl(S.ָ�����ۼ�,0) as ָ���ۼ�,nvl(S.ָ�������,0) as ָ������,nvl(I.���ηѱ�,0)  as ���ηѱ�," & _
            "        decode(sign(P.ִ������-sysdate),1,1,decode(sign(P.��ֹ����-sysdate),-1,-1,0)) as ִ�����," & _
            "        '['||I.����||']'||I.����||' '||I.���||' '||I.���� as ҩƷ,I.���㵥λ as ��λ,S.ҩ�ⵥλ,Nvl(S.ҩ���װ,1) ҩ���װ," & _
            "        P.�ּ� as �ۼ�,U.���� as ������Ŀ,P.����˵��," & _
            "        to_char(P.ִ������,'YYYY-MM-DD HH24:MI:SS') as ִ������,I.ID ҩƷID,P.No ����No " & _
            " from �շѼ�Ŀ P,������Ŀ U,�շ���ĿĿ¼ I,ҩƷ��� S" & _
            " where P.�շ�ϸĿID=I.ID and P.������ĿID=U.ID and I.ID=S.ҩƷID" & _
            "       And S.ҩ��ID=[1] " & GetPriceClassString("P")
    If Me.mnuViewPrices.Checked = False Then
        gstrSql = gstrSql & " and (P.��ֹ���� is null or P.��ֹ����>=sysdate)"
    End If
    If Me.mnuViewStoped.Checked = False Then
        gstrSql = gstrSql & " and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))"
    End If
    gstrSql = gstrSql & " order by I.����,P.ִ������ desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(Item.Key, 2)))
    
    With rsTemp
        Me.hgdPrice.Redraw = False
        If .BOF Or .EOF Then
            With Me.hgdPrice
                .Rows = .FixedRows + 1: .RowData(.FixedRows) = 0
                For intCol = 0 To .Cols - 1
                    .TextMatrix(.FixedRows, intCol) = ""
                Next
            End With
        Else
            Me.hgdPrice.Rows = Me.hgdPrice.FixedRows + .RecordCount
        End If
        Do While Not .EOF
            Me.hgdPrice.RowData(.AbsolutePosition) = !ID
            Me.hgdPrice.TextMatrix(.AbsolutePosition, colPriceҩ������) = !ҩ������
            Me.hgdPrice.TextMatrix(.AbsolutePosition, colPriceָ������) = Format(!ָ������ * IIf(mintҩ�ⵥλ = 0, 1, !ҩ���װ), mstrCostFormat)
            Me.hgdPrice.TextMatrix(.AbsolutePosition, colPrice����) = Format(!����, "0.00000;-0.00000;0")
            Me.hgdPrice.TextMatrix(.AbsolutePosition, colPriceָ���ۼ�) = Format(!ָ���ۼ� * IIf(mintҩ�ⵥλ = 0, 1, !ҩ���װ), mstrPriceFormat)
            Me.hgdPrice.TextMatrix(.AbsolutePosition, colPriceָ������) = Format(!ָ������, "0.00000;-0.00000;0")
            Me.hgdPrice.TextMatrix(.AbsolutePosition, colPrice���ηѱ�) = !���ηѱ�
            Me.hgdPrice.TextMatrix(.AbsolutePosition, colPriceִ�����) = !ִ�����
            Me.hgdPrice.TextMatrix(.AbsolutePosition, colPrice����NO) = IIf(IsNull(!����No), "", !����No)
            Me.hgdPrice.TextMatrix(.AbsolutePosition, colPriceҩƷ) = !ҩƷ
            Me.hgdPrice.TextMatrix(.AbsolutePosition, colPrice��λ) = IIf(mintҩ�ⵥλ = 0, !��λ, !ҩ�ⵥλ)
            Me.hgdPrice.TextMatrix(.AbsolutePosition, colPrice�ۼ�) = Format(!�ۼ� * IIf(mintҩ�ⵥλ = 0, 1, !ҩ���װ), mstrPriceFormat)
            Me.hgdPrice.TextMatrix(.AbsolutePosition, colPrice������Ŀ) = !������Ŀ
            Me.hgdPrice.TextMatrix(.AbsolutePosition, colPrice˵��) = IIf(IsNull(!����˵��), "", !����˵��)
            Me.hgdPrice.TextMatrix(.AbsolutePosition, colPriceִ������) = !ִ������
            Me.hgdPrice.TextMatrix(.AbsolutePosition, colPriceҩƷID) = !ҩƷid
            Me.hgdPrice.Row = .AbsolutePosition
            For intCol = 0 To Me.hgdPrice.Cols - 1
                Me.hgdPrice.Col = intCol
                Select Case !ִ�����
                Case -1
                    Me.hgdPrice.CellBackColor = RGB(240, 240, 240)
                Case 0
                    Me.hgdPrice.CellBackColor = RGB(255, 255, 255)
                Case 1
                    Me.hgdPrice.CellBackColor = RGB(225, 255, 255)
                End Select
            Next
            .MoveNext
        Loop
        Me.hgdPrice.Row = Me.hgdPrice.FixedRows
        If Val(Me.tvwClass.Tag) < 2 Or (Val(Me.tvwClass.Tag) = 4 And mstrType <> "7") Then
            If Me.hgdPrice.ColWidth(colPriceҩƷ) = 0 Or Me.hgdPrice.ColWidth(colPrice��λ) = 0 Then
                Me.hgdPrice.ColWidth(colPriceҩƷ) = 3500
                Me.hgdPrice.ColWidth(colPrice��λ) = 550
            End If
        ElseIf Val(Me.tvwClass.Tag) = 2 Or (Val(Me.tvwClass.Tag) = 4 And mstrType = "7") Then
            Me.hgdPrice.ColWidth(colPriceҩƷ) = 0
            Me.hgdPrice.ColWidth(colPrice��λ) = 0
        End If
        Me.hgdPrice.Redraw = True
    End With
    
    Call hgdPrice_RowColChange
    
    'ȡ�ɱ��۵��ۼ�¼
    mlngCurrDrug = Val(Mid(Item.Key, 2))
    If Me.tabContent.Tab = 2 Then
        Call GetCostAdjust(mlngCurrDrug)
    End If
    
    'ȡ�ѱ�ȼ�����
    Call GetChargeSet(mlngCurrDrug)
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Call lvwItems_GotFocus
    If lvwSpecs.SelectedItem Is Nothing Then
        mnuEditSpecMod.Enabled = False
        mnuEditSpecDel.Enabled = False
    ElseIf lvwSpecs.SelectedItem.ForeColor = vbRed Then
        mnuEditSpecMod.Enabled = False
    Else
        mnuEditSpecMod.Enabled = True
        mnuEditSpecDel.Enabled = True
    End If
    If lvwItems.SelectedItem.ForeColor = vbRed Then
        mnuEditSpecAdd.Enabled = False
        mnuEditItemMod.Enabled = False
    Else
        mnuEditSpecAdd.Enabled = True
    End If
    
    If lblComment(0).Caption = "" Then
        lvwSpecs.Height = tabContent.Height - lvwSpecs.Top - 50
    Else
        lvwSpecs.Height = tabContent.Height - lvwSpecs.Top - 50 - fraComment(0).Height
    End If

    If lblComment(3).Caption = "" Then
        hgdPrice.Height = tabContent.Height - hgdPrice.Top - 50
    Else
        hgdPrice.Height = tabContent.Height - hgdPrice.Top - 350 - lblComment(3).Height
    End If
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub lvwItems_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call lvwItems_DblClick
End Sub

Private Sub lvwItems_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    Call zlPopupEditMenu(1, True)
End Sub

Private Sub lvwSpecs_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwSpecs.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwSpecs.SortOrder = IIf(Me.lvwSpecs.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwSpecs.SortKey = ColumnHeader.Index - 1
        Me.lvwSpecs.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwSpecs_DblClick()
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    If Me.lvwSpecs.SelectedItem Is Nothing Then Exit Sub
    If Val(Me.tvwClass.Tag) = 2 Or mstrType = "7" Then
        With frmMediHerbalSpec
            .stbSpec.Tag = "����"
            .lngҩ��id = Mid(Me.lvwItems.SelectedItem.Key, 2)
            .lngҩƷID = Mid(Me.lvwSpecs.SelectedItem.Key, 2)
            .strPrivs = Me.mstrPrivs
            .Show 1, Me
        End With
    Else
        With frmMediSpec
            .stbSpec.Tag = "����"
            .lngҩ��id = Mid(Me.lvwItems.SelectedItem.Key, 2)
            .lngҩƷID = Mid(Me.lvwSpecs.SelectedItem.Key, 2)
            .strPrivs = Me.mstrPrivs
            .Show 1, Me
        End With
    End If
End Sub

Private Sub lvwSpecs_GotFocus()
    Set objItem = Me.lvwItems.SelectedItem
    If objItem Is Nothing Then
        Exit Sub
    End If
    
    If lvwSpecs.ListItems.Count = 0 Then Exit Sub
    If Val(Me.tvwClass.Tag) = 0 And InStr(1, mstrPrivs, "��������ҩ") = 0 Then Exit Sub
    If Val(Me.tvwClass.Tag) = 1 And InStr(1, mstrPrivs, "�����г�ҩ") = 0 Then Exit Sub
    If Val(Me.tvwClass.Tag) = 2 And InStr(1, mstrPrivs, "�����в�ҩ") = 0 Then Exit Sub
    If Val(Me.tvwClass.Tag) = 4 And InStr(1, mstrPrivs, "��������ҩ") = 0 _
        And InStr(1, mstrPrivs, "�����г�ҩ") = 0 _
        And InStr(1, mstrPrivs, "�����в�ҩ") = 0 Then
        Exit Sub
    End If
    
    '����ҩƷ��Ƭ�����á�ͣ�ñ�־
    If lvwSpecs.SelectedItem.Icon = "���S" Or lvwSpecs.SelectedItem.Icon = "�ݹ�S" Then
        If mnuEditStart.Visible = True Then mnuEditStart.Enabled = True
        mnuEditStop.Enabled = False
        tlbThis.Buttons("Start").Enabled = mnuEditStart.Enabled
        tlbThis.Buttons("Stop").Enabled = False
    Else
        mnuEditStart.Enabled = False
        If mnuEditStop.Visible = True Then mnuEditStop.Enabled = True
        tlbThis.Buttons("Start").Enabled = False
        tlbThis.Buttons("Stop").Enabled = mnuEditStop.Enabled
    End If
    
    If Val(Me.tvwClass.Tag) = 0 Then
        If InStr(1, mstrPrivs, "��������ҩ���") = 0 Then
            Me.mnuEditSpecAdd.Enabled = False
            Me.mnuEditSpecMod.Enabled = False
            Me.mnuEditSpecDel.Enabled = False
            mnuEditStart.Enabled = False
            mnuEditStop.Enabled = False
            tlbThis.Buttons("Start").Enabled = False
            tlbThis.Buttons("Stop").Enabled = False
        End If
    ElseIf Val(Me.tvwClass.Tag) = 1 Then
        If InStr(1, mstrPrivs, "�����г�ҩ���") = 0 Then
            Me.mnuEditSpecAdd.Enabled = False
            Me.mnuEditSpecMod.Enabled = False
            Me.mnuEditSpecDel.Enabled = False
            mnuEditStart.Enabled = False
            mnuEditStop.Enabled = False
            tlbThis.Buttons("Start").Enabled = False
            tlbThis.Buttons("Stop").Enabled = False
        End If
    ElseIf Val(Me.tvwClass.Tag) = 2 Then
        If InStr(1, mstrPrivs, "�����в�ҩ���") = 0 Then
            Me.mnuEditSpecAdd.Enabled = False
            Me.mnuEditSpecMod.Enabled = False
            Me.mnuEditSpecDel.Enabled = False
            mnuEditStart.Enabled = False
            mnuEditStop.Enabled = False
            tlbThis.Buttons("Start").Enabled = False
            tlbThis.Buttons("Stop").Enabled = False
        End If
    ElseIf Val(Me.tvwClass.Tag) = 4 Then
        If objItem.SubItems(Me.lvwItems.ColumnHeaders("_����").Index - 1) = "1" Then
            If InStr(1, mstrPrivs, "��������ҩ���") = 0 Then
                Me.mnuEditSpecAdd.Enabled = False
                Me.mnuEditSpecMod.Enabled = False
                Me.mnuEditSpecDel.Enabled = False
                mnuEditStart.Enabled = False
                mnuEditStop.Enabled = False
                tlbThis.Buttons("Start").Enabled = False
                tlbThis.Buttons("Stop").Enabled = False
            End If
        ElseIf objItem.SubItems(Me.lvwItems.ColumnHeaders("_����").Index - 1) = "2" Then
            If InStr(1, mstrPrivs, "�����г�ҩ���") = 0 Then
                Me.mnuEditSpecAdd.Enabled = False
                Me.mnuEditSpecMod.Enabled = False
                Me.mnuEditSpecDel.Enabled = False
                mnuEditStart.Enabled = False
                mnuEditStop.Enabled = False
                tlbThis.Buttons("Start").Enabled = False
                tlbThis.Buttons("Stop").Enabled = False
            End If
        ElseIf objItem.SubItems(Me.lvwItems.ColumnHeaders("_����").Index - 1) = "3" Then
            If InStr(1, mstrPrivs, "�����в�ҩ���") = 0 Then
                Me.mnuEditSpecAdd.Enabled = False
                Me.mnuEditSpecMod.Enabled = False
                Me.mnuEditSpecDel.Enabled = False
                mnuEditStart.Enabled = False
                mnuEditStop.Enabled = False
                tlbThis.Buttons("Start").Enabled = False
                tlbThis.Buttons("Stop").Enabled = False
            End If
        End If
    End If
End Sub

Private Sub lvwSpecs_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim rsData As ADODB.Recordset
    
    err = 0: On Error GoTo ErrHand
    
    '����ѵ�ִ�����ڶ��۸�δִ�У�ִ�м������
'    gstrSql = " Select ID From �շѼ�Ŀ Where �շ�ϸĿID=[1] And �䶯ԭ��=0" & GetPriceClassString("")
'
'    Set rsData = zlDatabase.OpenSQLRecord(gstrSql, "���δִ�еļ۸�", Val(Mid(Item.Key, 2)))
'
'    With rsData
'        If Not .EOF Then
'            If Not IsNull(!ID) Then
'                gstrSql = "zl_ҩƷ�շ���¼_Adjust(" & Val(!ID) & ")"
'                Call zlDatabase.ExecuteProcedure(gstrSql, "����ҩƷ�۸������¼")
'            End If
'        End If
'    End With
    
    gstrSql = "zl_ҩƷ�շ���¼_Adjust(" & Val(Mid(Item.Key, 2)) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSql, "����ҩƷ�۸������¼")
        
    gstrSql = "select I.���㵥λ||decode(I.���㵥λ,O.���㵥λ,'','(='||decode(sign(S.����ϵ��-1),-1,'0','')||to_char(S.����ϵ��)||O.���㵥λ||')') as �ۼ۵�λ," & _
            "        S.���ﵥλ||decode(S.���ﵥλ,I.���㵥λ,'','(='||decode(sign(S.�����װ-1),-1,'0','')||to_char(S.�����װ)||I.���㵥λ||')') as ���ﵥλ," & _
            "        S.סԺ��λ||decode(S.סԺ��λ,I.���㵥λ,'','(='||decode(sign(S.סԺ��װ-1),-1,'0','')||to_char(S.סԺ��װ)||I.���㵥λ||')') as סԺ��λ," & _
            "        S.ҩ�ⵥλ||decode(S.ҩ�ⵥλ,I.���㵥λ,'','(='||decode(sign(S.ҩ���װ-1),-1,'0','')||to_char(S.ҩ���װ)||I.���㵥λ||')') as ҩ�ⵥλ," & _
            "        nvl(S.ҩ�����,0) as ҩ�����,nvl(S.ҩ������,0) as ҩ������,nvl(S.���Ч��,0) as ���Ч��,nvl(S.סԺ�ɷ����,0) as �ɷ����,Nvl(To_Char(I.����ʱ��,'yyyy-MM-dd'),'3000-01-01') As ����ʱ��" & _
            " from ҩƷ��� S,�շ���ĿĿ¼ I,������ĿĿ¼ O" & _
            " where S.ҩƷID=I.ID and S.ҩ��ID=O.ID" & _
            "       and S.ҩƷid=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(Item.Key, 2)))
        
    With rsTemp
        If Val(Me.tvwClass.Tag) = 2 Or mstrType = "7" Then
            Me.lblComment(0).Caption = "1���ۼ۵�λ��" & !�ۼ۵�λ & "�� ҩ����λ��" & !���ﵥλ & "�� ҩ�ⵥλ��" & !ҩ�ⵥλ & "��"
        Else
            Me.lblComment(0).Caption = "1���ۼ۵�λ��" & !�ۼ۵�λ & "�� ���ﵥλ��" & !���ﵥλ & "�� סԺ��λ��" & !סԺ��λ & "�� ҩ�ⵥλ��" & !ҩ�ⵥλ & "��"
        End If
        
        If !ҩ����� = 0 Then
            Me.lblComment(1).Caption = "2����ҩƷ�����з�������"
        Else
            If !ҩ������ = 0 Then
                Me.lblComment(1).Caption = "2����ҩƷ��ҩ���з�������"
            Else
                Me.lblComment(1).Caption = "2����ҩƷ��ҩ��ҩ������Ҫ��������"
            End If
            If !���Ч�� = 0 Then
                Me.lblComment(1).Caption = Me.lblComment(1).Caption & "��������Ч�ڸ��١�"
            Else
                Me.lblComment(1).Caption = Me.lblComment(1).Caption & "�������" & !���Ч�� & "�¡�"
            End If
        End If
        Select Case !�ɷ����
        Case 0
            Me.lblComment(2).Caption = "3����ҩƷ�������Ӧ�á�"
        Case 1
            Me.lblComment(2).Caption = "3����ҩƷ���������Ӧ�á�"
        Case 2
            Me.lblComment(2).Caption = "3����ҩƷΪһ����ҩƷ��"
        Case Is < 0
            Me.lblComment(2).Caption = "3����ҩƷ�����" & Abs(!�ɷ����) & "����ʹ����Ч��"
        Case Else
        End Select
    End With
    Call lvwSpecs_GotFocus
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub lvwSpecs_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Me.lvwSpecs.SelectedItem Is Nothing Then Exit Sub
    Call lvwSpecs_DblClick
End Sub

Private Sub lvwSpecs_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    
    Call zlPopupEditMenu(2, True)
End Sub

Private Sub mnuClassAdd_Click()
    Dim intTab As Integer
    
    intTab = tabContent.Tab
    If Val(Me.tvwClass.Tag) = 4 Then            '=3����ʾ���˽����״̬��������༭���
        Exit Sub
    End If
    With frmClinicClass
        .lblKind.Tag = Val(Me.tvwClass.Tag) + 1
        If Me.tvwClass.SelectedItem Is Nothing Then
            .txtParent.Tag = 0
        Else
            .txtParent.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
            If tvwClass.SelectedItem.Text <> "" Then
                .txtParent.Text = tvwClass.SelectedItem.Text
            End If
        End If
        .Tag = "����"
        .Show 1, Me
    End With
    If gblnCancel = False Then
        If Me.tvwClass.SelectedItem Is Nothing Then
            Call zlRefClasses
        Else
            Call zlRefClasses(Mid(Me.tvwClass.SelectedItem.Key, 2))
        End If
    End If
    tabContent.Tab = intTab
End Sub

Private Sub mnuClassDel_Click()
    If Val(Me.tvwClass.Tag) = 4 Then           '=3����ʾ���˽����״̬��������༭���
        Exit Sub
    End If
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    If MsgBox("���ɾ���÷��ࡰ" & Me.tvwClass.SelectedItem.Text & "����", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    err = 0: On Error GoTo ErrHand
    gstrSql = "zl_���Ʒ���Ŀ¼_delete(" & Mid(Me.tvwClass.SelectedItem.Key, 2) & ")"
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
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuClassMod_Click()
    Dim intTab As Integer   '��¼��ǰ��ѡ���ҳ��
    
    intTab = tabContent.Tab
    If Val(Me.tvwClass.Tag) = 4 Then            '=3����ʾ���˽����״̬��������༭���
        Exit Sub
    End If
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    With frmClinicClass
        .lblKind.Tag = Val(Me.tvwClass.Tag) + 1
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
    If gblnCancel = False Then
        Call zlRefClasses(Mid(Me.tvwClass.SelectedItem.Key, 2))
    End If
    tabContent.Tab = intTab
End Sub

Private Sub mnuClassStar_Click()
    'ͣ�÷��ࡢ�ӷ��ࡢ������Ʒ�ּ����
    
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    
    frmMediClassReuse.ShowForm Val(Mid(tvwClass.SelectedItem.Key, 2)), Me.tvwClass.Tag
    
    Call zlRefClasses
End Sub

Private Sub mnuClassStop_Click()
    'ͣ�÷��ࡢ�ӷ��ࡢ������Ʒ�ּ����
    
    On Error GoTo ErrHand
    
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    
    If MsgBox("�Ƿ�ͣ�ø÷��༰�÷���������ҩƷ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    gstrSql = "Zl_���Ʒ���Ŀ¼_ҩƷ����ͣ��(" & Val(Mid(tvwClass.SelectedItem.Key, 2)) & "," & Val(Me.tvwClass.Tag) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    
    Call zlRefClasses
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub mnuEditContrast_Click()
    frmMediContrast.ShowMe Me
End Sub

Private Sub mnuEditExcel_Click()
'    frmItemImport.ShowMe 2, Me
    frmImportFile.ShowMe 2, Me
'    Call zlRefClasses
    Call zlRefRecords
    If Me.tvwClass.SelectedItem Is Nothing Then
        Call zlRefClasses
    Else
        Call zlRefClasses(Mid(Me.tvwClass.SelectedItem.Key, 2))
    End If
End Sub

Private Sub mnuEditManFac_Click()
    Dim strType As String
    Dim str���� As String
    Dim lngҩƷID As String
    
    On Error Resume Next
    '���Һ���׼�ĺ�����
    With frmSetManfac
        
        If Me.tvwClass.Tag = 4 Or Me.tvwClass.Tag = 3 Then '����������˽��
            If Me.lvwItems.SelectedItem Is Nothing Then     '���û�м�¼���˳�����Ϊ�޷��ж�ҩƷ����
                Exit Sub
            End If
            strType = Me.lvwItems.SelectedItem.SubItems(Me.lvwItems.ColumnHeaders("_����").Index - 1)
            strType = Switch(strType = "1", "5", strType = "2", "6", strType = "3", "7")
            str���� = strType
        Else
            str���� = Switch(Me.tvwClass.Tag = "0", "5", Me.tvwClass.Tag = "1", "6", Me.tvwClass.Tag = "2", "7")
        End If
        If Me.lvwSpecs.SelectedItem Is Nothing Then
            lngҩƷID = 0
        Else
            lngҩƷID = Mid(Me.lvwSpecs.SelectedItem.Key, 2)
        End If
        .ShowMe str����, Me.mstrPrivs, lngҩƷID
    End With
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
End Sub

Private Sub mnuEditRate_Click()
    frm�ӳ�������.ShowMe Me
End Sub

Private Sub mnuEditSendType_Click()
    frmMediSendType.Show vbModal, Me
End Sub

Private Sub mnuEditSpecBatch_Click()
    '����޸�
    frmBatchUpdate.ShowMe 2, mstrPrivs, mbln�Թ�ҩ
End Sub

Private Sub mnuEditSpecExp_Click()
    frmMediSpecExp.Show
End Sub

Private Sub mnuEditVariBatch_Click()
    '1��Ʒ���޸�
    frmBatchUpdate.ShowMe 1, mstrPrivs, mbln�Թ�ҩ
End Sub

Private Sub mnuPriceChargeSet_Click()
    If Me.lvwSpecs.SelectedItem Is Nothing Then Exit Sub
'    frmChargeSortItemEdit.ShowMe Me, 3, "", Val(Mid(Me.lvwSpecs.SelectedItem.Key, 2)), Me.lvwSpecs.SelectedItem.Text
    
    frmSetExpense.ShowMe Me, Val(Mid(Me.lvwSpecs.SelectedItem.Key, 2)), Me.lvwSpecs.SelectedItem.Text
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
End Sub

Private Sub mnuEditItemAdd_Click()
    Dim lng����id As Long
    Dim lngҩ��id As Long
    Dim int���� As Integer
    
    If Me.tvwClass.SelectedItem Is Nothing Then MsgBox "��δ���÷���,������ɾƷ�֣�", vbExclamation, gstrSysName: Exit Sub
    If Val(Me.tvwClass.Tag) < 2 Then
        With frmMediItem
            .Tag = IIf(Me.tvwClass.Tag = 0, 1, 2)
            .cmdCancel.Tag = "����"
            .lng����id = Mid(Me.tvwClass.SelectedItem.Key, 2)
            If Me.lvwItems.SelectedItem Is Nothing Then
                .lngҩ��id = 0
            Else
                .lngҩ��id = Mid(Me.lvwItems.SelectedItem.Key, 2)
            End If
            .strPrivs = Me.mstrPrivs
            .lng������ = 0
            .ShowMe mbln�Թ�ҩ, Me
        End With
    ElseIf Val(Me.tvwClass.Tag) = 2 Then
        '��ҩƷ��
        With frmMediHerbalItem
            .Tag = 3
            .cmdCancel.Tag = "����"
            .lng����id = Mid(Me.tvwClass.SelectedItem.Key, 2)
            If Me.lvwItems.SelectedItem Is Nothing Then
                .lngҩ��id = 0
            Else
                .lngҩ��id = Mid(Me.lvwItems.SelectedItem.Key, 2)
            End If
            .strPrivs = Me.mstrPrivs
            .ShowMe mbln�Թ�ҩ, Me
        End With
    ElseIf Val(Me.tvwClass.Tag) = 3 Then
        If Not lvwItems.SelectedItem Is Nothing Then
            lng����id = objItem.SubItems(Me.lvwItems.ColumnHeaders("_����id").Index - 1)
            lngҩ��id = Mid(Me.lvwItems.SelectedItem.Key, 2)
            int���� = objItem.SubItems(Me.lvwItems.ColumnHeaders("_����").Index - 1)
        Else
            Exit Sub
        End If
        If mstrType <> "7" Then
             With frmMediItem
                .lng������ = 0
                .chkԭ��ҩ.Value = 0
                .chkר��ҩ.Value = 0
                .chk��������.Value = 0
                .Tag = int����
                .cmdCancel.Tag = "����"
                .lng����id = lng����id
                .lngҩ��id = lngҩ��id
                If Val(Me.tvwClass.Tag) = 3 Then
                    If Not Me.tvwClass.SelectedItem.Parent Is Nothing Then
                        If IsNumeric(Mid(Me.tvwClass.SelectedItem.Key, 2)) Then
                            If Me.tvwClass.SelectedItem.Parent.Key Like "_L*" Then
                                .lng������ = Mid(Me.tvwClass.SelectedItem.Key, 2, 1)
                            ElseIf Me.tvwClass.SelectedItem.Parent.Key Like "_ԭ��ҩ" Then
                                .chkԭ��ҩ.Value = 1
                            ElseIf Me.tvwClass.SelectedItem.Parent.Key Like "_ר��ҩ" Then
                                .chkר��ҩ.Value = 1
                            ElseIf Me.tvwClass.SelectedItem.Parent.Key Like "_��������" Then
                                .chk��������.Value = 1
                            End If
                        Else
                            If Me.tvwClass.SelectedItem.Parent.Key Like "_����ҩ" Then
                                .lng������ = Mid(Me.tvwClass.SelectedItem.Key, 7, 1)
                            End If
                        End If
                    Else
                        If Me.tvwClass.SelectedItem.Key Like "_����ҩ" Then
                            .lng������ = 1
                        ElseIf Me.tvwClass.SelectedItem.Key Like "_ԭ��ҩ" Then
                            .chkԭ��ҩ.Value = 1
                        ElseIf Me.tvwClass.SelectedItem.Key Like "_ר��ҩ" Then
                            .chkר��ҩ.Value = 1
                        ElseIf Me.tvwClass.SelectedItem.Key Like "_��������" Then
                            .chk��������.Value = 1
                        End If
                    End If
                End If
                
                .strPrivs = Me.mstrPrivs
                .ShowMe mbln�Թ�ҩ, Me
            End With
        Else
            With frmMediHerbalItem
                .Tag = 3
                .cmdCancel.Tag = "����"
                .lng����id = lng����id
                .lngҩ��id = lngҩ��id
                .strPrivs = Me.mstrPrivs
                .ShowMe mbln�Թ�ҩ, Me
            End With
        End If
    ElseIf Val(Me.tvwClass.Tag) = 4 Then        '����������˽��
        If (tvwClass.SelectedItem.Key Like "_L*" Or tvwClass.SelectedItem.Key Like "_A*") And lvwItems.SelectedItem Is Nothing Then
            Exit Sub
        ElseIf (tvwClass.SelectedItem.Key Like "_L*" Or tvwClass.SelectedItem.Key Like "_A*") And Not lvwItems.SelectedItem Is Nothing Then
            lng����id = objItem.SubItems(Me.lvwItems.ColumnHeaders("_����id").Index - 1)
            lngҩ��id = Mid(Me.lvwItems.SelectedItem.Key, 2)
            int���� = objItem.SubItems(Me.lvwItems.ColumnHeaders("_����").Index - 1)
        ElseIf (Not tvwClass.SelectedItem.Key Like "_L*" Or tvwClass.SelectedItem.Key Like "_A*") And lvwItems.SelectedItem Is Nothing Then
            lng����id = Mid(Me.tvwClass.SelectedItem.Key, IIf(Val(Me.tvwClass.Tag) = 3, 3, 2))
            lngҩ��id = 0
            int���� = 1
        ElseIf (Not tvwClass.SelectedItem.Key Like "_L*" Or tvwClass.SelectedItem.Key Like "_A*") And Not lvwItems.SelectedItem Is Nothing Then
            lng����id = Mid(Me.tvwClass.SelectedItem.Key, IIf(Val(Me.tvwClass.Tag) = 3, 3, 2))
            lngҩ��id = Mid(Me.lvwItems.SelectedItem.Key, 2)
            int���� = objItem.SubItems(Me.lvwItems.ColumnHeaders("_����").Index - 1)
        End If
        If mstrType <> "7" Then
             With frmMediItem
                .Tag = int����
                .cmdCancel.Tag = "����"
                .lng����id = lng����id
                .lngҩ��id = lngҩ��id
                .strPrivs = Me.mstrPrivs
                .ShowMe mbln�Թ�ҩ, Me
            End With
        Else
            With frmMediHerbalItem
                .Tag = 3
                .cmdCancel.Tag = "����"
                .lng����id = lng����id
                .lngҩ��id = lngҩ��id
                .strPrivs = Me.mstrPrivs
                .ShowMe mbln�Թ�ҩ, Me
            End With
        End If
    End If
    If gblnCancel = False Then
        Call zlRefRecords
    End If
End Sub

Private Sub mnuEditItemBill_Click()
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call frmClinicBill.ShowMe(Me, Mid(Me.lvwItems.SelectedItem.Key, 2))
    Call zlRefRecords(Mid(Me.lvwItems.SelectedItem.Key, 2))
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
End Sub

Private Sub mnuEditItemDel_Click()
    Dim lngItem As Long
    Dim intCol As Integer
    Dim blnTrans As Boolean
    Dim rsSpec As New ADODB.Recordset
    
    On Error GoTo ErrHand
    
    With Me.lvwItems
        If .SelectedItem Is Nothing Then Exit Sub
        If MsgBox("���ɾ����" & .SelectedItem.Text & "����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        lngItem = Mid(.SelectedItem.Key, 2)
        gstrSql = "Select ҩƷID From ҩƷ��� Where ҩ��ID=[1]"
        Set rsSpec = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItem)
        
        gcnOracle.BeginTrans
        blnTrans = True
        'ɾ��������ĿĿ¼
'        If Val(Me.tvwClass.Tag) < 2 Or (Val(Me.tvwClass.Tag) = 3 And mstrType <> "7") Then
            gstrSql = "zl_��ҩƷ��_DELETE(" & lngItem & ")"
'        ElseIf Val(Me.tvwClass.Tag) = 2 Or (Val(Me.tvwClass.Tag) = 3 And mstrType = "7") Then
'            gstrSql = "zl_��ҩҩƷ_DELETE(" & lngItem & ")"
'        End If
        Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
        
        'ɾ����Ӧ���շ���ĿĿ¼
        Do While Not rsSpec.EOF
            gstrSql = "zl_��ҩ���_DELETE(" & rsSpec!ҩƷid & ")"
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            rsSpec.MoveNext
        Loop
        gcnOracle.CommitTrans
        blnTrans = False
        
        'ͬ��ɾ������ƽ̨ҩƷ��Ϣ
        If Not gobjLogisticPlatform Is Nothing And rsSpec.RecordCount > 0 Then
            rsSpec.MoveFirst
            Do While Not rsSpec.EOF
                gobjLogisticPlatform.ClearDrugInfo rsSpec!ҩƷid, 0
                rsSpec.MoveNext
            Loop
        End If
        
        Call .ListItems.Remove(.SelectedItem.Key)
        If .SelectedItem Is Nothing Then
            lvwSpecs.ListItems.Clear
            With Me.hgdPrice
                .Rows = .FixedRows + 1: .RowData(.FixedRows) = 0
                For intCol = 0 To .Cols - 1
                    .TextMatrix(.FixedRows, intCol) = ""
                Next
            End With
        Else
            Call lvwItems_ItemClick(.SelectedItem)
        End If
        
        '������˽�����ڷ���ҩ��ID����ȥ���Ѿ�ɾ����ҩ��id
        Dim i As Integer
        Dim strAryDrugId() As String
        Dim strTmp As String
        
        If Val(Me.tvwClass.Tag) = 4 Then
            mstrDrugId = mstrDrugId & ","
            strAryDrugId = Split(mstrDrugId, ",")
            For i = 0 To UBound(strAryDrugId) - 1
                If strAryDrugId(i) <> CStr(lngItem) Then
                    strTmp = strTmp & strAryDrugId(i) & ","
                End If
            Next
            If Len(strTmp) > 1 Then
                strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
            Else
                strTmp = ""
            End If
            mstrDrugId = strTmp
        End If
        
    End With
    Exit Sub
ErrHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditItemMod_Click()
    If Me.tvwClass.SelectedItem Is Nothing Then MsgBox "��δ���÷���,������ɾƷ�֣�", vbExclamation, gstrSysName: Exit Sub
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    If Me.lvwItems.SelectedItem.Icon = "��ҩS" Then
        MsgBox "���ܶ�ͣ��ҩƷҩƷ�����޸ģ�", vbExclamation, gstrSysName
        Exit Sub
    End If
    If Val(Me.tvwClass.Tag) < 2 Then
        With frmMediItem
            .Tag = IIf(Me.tvwClass.Tag = 0, 1, 2)
            .cmdCancel.Tag = "�޸�"
            .lng����id = Mid(Me.tvwClass.SelectedItem.Key, 2)
            .lngҩ��id = Mid(Me.lvwItems.SelectedItem.Key, 2)
            .strPrivs = Me.mstrPrivs
            .lng������ = 0
            .Show 1, Me
        End With
    ElseIf Val(Me.tvwClass.Tag) = 2 Then
        With frmMediHerbalItem
            .Tag = 3
            .cmdCancel.Tag = "�޸�"
            .lng����id = Mid(Me.tvwClass.SelectedItem.Key, 2)
            .lngҩ��id = Mid(Me.lvwItems.SelectedItem.Key, 2)
            .strPrivs = Me.mstrPrivs
            .Show 1, Me
        End With
    ElseIf Val(Me.tvwClass.Tag) = 3 Or Val(Me.tvwClass.Tag) = 4 Then        '����������˽��
        Set objItem = Me.lvwItems.SelectedItem
        If objItem Is Nothing Then
            Exit Sub
        End If
        If mstrType <> "7" Then
             With frmMediItem
                .Tag = objItem.SubItems(Me.lvwItems.ColumnHeaders("_����").Index - 1)
                .cmdCancel.Tag = "�޸�"
                .lng����id = objItem.SubItems(Me.lvwItems.ColumnHeaders("_����id").Index - 1)
                .lngҩ��id = Mid(Me.lvwItems.SelectedItem.Key, 2)
                .strPrivs = Me.mstrPrivs
                If Val(Me.tvwClass.Tag) = 3 Then
                    If Not Me.tvwClass.SelectedItem.Parent Is Nothing Then
                        If IsNumeric(Mid(Me.tvwClass.SelectedItem.Key, 2)) Then
                            If Me.tvwClass.SelectedItem.Parent.Key Like "_L*" Then
                                .lng������ = Mid(Me.tvwClass.SelectedItem.Key, 2, 1)
                            End If
                        Else
                            If Me.tvwClass.SelectedItem.Parent.Key Like "_����ҩ" Then
                                .lng������ = Mid(Me.tvwClass.SelectedItem.Key, 7, 1)
                            End If
                        End If
                    End If
                End If
                
                .Show 1, Me
            End With
        Else
            With frmMediHerbalItem
                .Tag = 3
                .cmdCancel.Tag = "�޸�"
                .lng����id = objItem.SubItems(Me.lvwItems.ColumnHeaders("_����id").Index - 1)
                .lngҩ��id = Mid(Me.lvwItems.SelectedItem.Key, 2)
                .strPrivs = Me.mstrPrivs
                .Show 1, Me
            End With
        End If
    End If
    If gblnCancel = False Then
        If Not (Me.lvwItems.SelectedItem Is Nothing) Then Call zlRefRecords(Mid(Me.lvwItems.SelectedItem.Key, 2))
    End If
End Sub

Private Sub mnuEditItemPart_Click()
    Dim int��;���� As Integer, lngҩƷID As Long, bln�༭ As Boolean
    Dim strStationNo As String
    With frmServiceSectOffice
        Dim strType As String
        If Me.tvwClass.Tag = 4 Or Me.tvwClass.Tag = 3 Then '����������˽��
            If Me.lvwItems.SelectedItem Is Nothing Then     '���û�м�¼���˳�����Ϊ�޷��ж�ҩƷ����
                Exit Sub
            End If
            int��;���� = CInt(Me.lvwItems.SelectedItem.SubItems(Me.lvwItems.ColumnHeaders("_����").Index - 1))
            int��;���� = Switch(int��;���� = 1, 5, int��;���� = 2, 6, int��;���� = 3, 7)
        Else
            int��;���� = Switch(Me.tvwClass.Tag = "0", 5, Me.tvwClass.Tag = "1", 6, Me.tvwClass.Tag = "2", 7)
        End If
        If Me.lvwSpecs.SelectedItem Is Nothing Then
            lngҩƷID = 0
        Else
            lngҩƷID = Mid(Me.lvwSpecs.SelectedItem.Key, 2)
            If gstrNodeNo <> "-" Then
                strStationNo = Me.lvwSpecs.SelectedItem.SubItems(Me.lvwSpecs.ColumnHeaders("վ��").Index - 1)
            End If
        End If
        bln�༭ = (InStr(1, mstrPrivs, "�洢�ⷿ") <> 0)
        Call .ShowMe(Me, lngҩƷID, int��;����, bln�༭, strStationNo)
    End With
End Sub

Private Sub mnuEditItemTabu_Click()
    With frmMediTabu
        Dim strType As String
        If Me.tvwClass.Tag = 4 Or Me.tvwClass.Tag = 3 Then '������˽�����Ϳ���ҩ��
            If Me.lvwItems.SelectedItem Is Nothing Then     '���û�м�¼���˳�����Ϊ�޷��ж�ҩƷ����
                Exit Sub
            End If
            strType = Me.lvwItems.SelectedItem.SubItems(Me.lvwItems.ColumnHeaders("_����").Index - 1)
            strType = Switch(strType = "1", "5", strType = "2", "6", strType = "3", "7")
            .Tag = strType
        Else
            .Tag = Switch(Me.tvwClass.Tag = "0", "5", Me.tvwClass.Tag = "1", "6", Me.tvwClass.Tag = "2", "7")
        End If
        If InStr(1, mstrPrivs, "������ɹ�ϵ") = 0 Then
            .cmdClose.Tag = "����"
        Else
            .cmdClose.Tag = "�޸�"
        End If
        If Me.lvwItems.SelectedItem Is Nothing Then
            .lblMedi.Tag = 0
        Else
            .lblMedi.Tag = Mid(Me.lvwItems.SelectedItem.Key, 2)
        End If
        .Show 1, Me
    End With
End Sub

Private Sub mnuEditItemUsage_Click()
    If Me.ActiveControl Is lvwItems Then
        If Me.lvwItems.SelectedItem Is Nothing Then
            If InStr(1, mstrPrivs, "�÷�����") = 0 Then Exit Sub
            Call frmMediUsage.ShowMe(Me, True)
        Else
            If InStr(1, mstrPrivs, "�÷�����") = 0 Then
                Call frmMediUsage.ShowMe(Me, False, Mid(Me.lvwItems.SelectedItem.Key, 2))
            Else
                If Right(Me.lvwItems.SelectedItem.Icon, 1) = "S" Then MsgBox "ͣ��ҩƷ�����������÷�������", vbExclamation, gstrSysName: Exit Sub
                Call frmMediUsage.ShowMe(Me, True, Mid(Me.lvwItems.SelectedItem.Key, 2))
                Call zlRefRecords(Mid(Me.lvwItems.SelectedItem.Key, 2))
            End If
        End If
    ElseIf Me.ActiveControl Is lvwSpecs Then
        If Me.lvwSpecs.SelectedItem.Icon = "���S" Then MsgBox "ͣ�ù�񣬲��������÷�������", vbExclamation, gstrSysName: Exit Sub
        Call frmMediUsage.ShowMe(Me, True, Mid(Me.lvwItems.SelectedItem.Key, 2), Mid(Me.lvwSpecs.SelectedItem.Key, 2))
    End If
End Sub

Private Sub mnuEditSpecAdd_Click()
    
    If Me.lvwItems.SelectedItem Is Nothing Then MsgBox "��δ����Ʒ��,�������ӹ��", vbExclamation, gstrSysName: Exit Sub
    mStrItem = lvwItems.SelectedItem.Key
    
    If Val(Me.tvwClass.Tag) = 2 Or mstrType = "7" Then
        With frmMediHerbalSpec
            .stbSpec.Tag = "����"
            .mlng����id = Val(Mid(tvwClass.SelectedItem.Key, 2))
            .lngҩ��id = Mid(Me.lvwItems.SelectedItem.Key, 2)
            If Me.lvwSpecs.SelectedItem Is Nothing Then
                .lngҩƷID = 0
            Else
                .lngҩƷID = Mid(Me.lvwSpecs.SelectedItem.Key, 2)
            End If
            .strPrivs = Me.mstrPrivs
            .Show 1, Me
        End With
    Else
        With frmMediSpec
            .stbSpec.Tag = "����"
            If Me.tvwClass.Tag < 3 Then
                .mlng����id = Val(Mid(tvwClass.SelectedItem.Key, 2))
                .lngҩ��id = Mid(Me.lvwItems.SelectedItem.Key, 2)
            Else
                If lvwSpecs.ListItems.Count = 0 Then '��ʾû�й��
                    .mlng����id = Get����id(Mid(lvwItems.SelectedItem.Key, 2), True)
                    .lngҩ��id = Mid(Me.lvwItems.SelectedItem.Key, 2)
                Else '��ʾ�й��
                    .mlng����id = Get����id(Mid(lvwSpecs.SelectedItem.Key, 2))
                    .lngҩ��id = Mid(Me.lvwItems.SelectedItem.Key, 2)
                End If
            End If
            If Me.lvwSpecs.SelectedItem Is Nothing Then
                .lngҩƷID = 0
            Else
                .lngҩƷID = Mid(Me.lvwSpecs.SelectedItem.Key, 2)
            End If
            .strPrivs = Me.mstrPrivs
            .Show 1, Me
        End With
    End If
    Call zlRefRecords
'    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
End Sub

Private Function Get����id(ByVal ID As Long, Optional ByVal blnƷ�� As Boolean) As Long
    '����:��ȡҩƷ����Ӧ�ķ���
    '����:blnƷ�ֱ�ʾ�Ƿ��Ǵ����Ʒ��id
    '����:����id
    Dim rsTemp As ADODB.Recordset
    On Error GoTo ErrHandle:
    If blnƷ�� = False Then
        gstrSql = "select c.����id from �շ���ĿĿ¼ a,ҩƷ��� b,������ĿĿ¼ c where a.id=b.ҩƷid and b.ҩ��id=c.id and a.id=[1]"
    Else
        gstrSql = "select ����id from ������ĿĿ¼ where id=[1]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "��ѯ����id", ID)
    
    If rsTemp.RecordCount > 0 Then
        Get����id = rsTemp!����id
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub mnuEditSpecDel_Click()
    Dim intCol As Integer
    With Me.lvwSpecs
        If .SelectedItem Is Nothing Then Exit Sub
        strTemp = Me.lvwItems.SelectedItem.Text & " " & .SelectedItem.Text & " " & .SelectedItem.SubItems(.ColumnHeaders("����").Index - 1)
        If MsgBox("���ɾ����" & strTemp & "����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        gstrSql = "zl_��ҩ���_DELETE(" & Mid(.SelectedItem.Key, 2) & ")"
        err = 0: On Error GoTo ErrHand
        Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
        
        'ͬ��ɾ������ƽ̨ҩƷ��Ϣ
        If Not gobjLogisticPlatform Is Nothing Then
            gobjLogisticPlatform.ClearDrugInfo Mid(.SelectedItem.Key, 2), 0
        End If
        
        Call .ListItems.Remove(.SelectedItem.Key)
    End With
    
    With Me.hgdPrice
        .Rows = .FixedRows + 1: .RowData(.FixedRows) = 0
        For intCol = 0 To .Cols - 1
            .TextMatrix(.FixedRows, intCol) = ""
        Next
    End With
    Call lvwItems_ItemClick(lvwItems.SelectedItem)
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditSpecLimit_Click()
    With frmMediLimit
        Dim strType As String
        If Me.tvwClass.Tag = 4 Or Me.tvwClass.Tag = 3 Then '����������˽��
            If Me.lvwItems.SelectedItem Is Nothing Then     '���û�м�¼���˳�����Ϊ�޷��ж�ҩƷ����
                Exit Sub
            End If
            strType = Me.lvwItems.SelectedItem.SubItems(Me.lvwItems.ColumnHeaders("_����").Index - 1)
            strType = Switch(strType = "1", "5", strType = "2", "6", strType = "3", "7")
            'strType = Switch(strType = "1", "5", strType = "2", "6", strType = "3", "6")
            .Tag = strType
        Else
            .Tag = Switch(Me.tvwClass.Tag = "0", "5", Me.tvwClass.Tag = "1", "6", Me.tvwClass.Tag = "2", "7")
            '.Tag = Switch(Me.tvwClass.Tag = "0", "5", Me.tvwClass.Tag = "1", "6", Me.tvwClass.Tag = "2", "6")
        End If
        .strPrivs = Me.mstrPrivs
        .Show 1, Me
    End With
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
End Sub

Private Sub mnuEditSpecMod_Click()
    Dim lngҩƷID As Long
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    If Me.lvwSpecs.SelectedItem Is Nothing Then Exit Sub
    If Me.lvwSpecs.SelectedItem.Icon = "���S" Or Me.lvwSpecs.SelectedItem.Icon = "�ݹ�S" Then
        MsgBox "���ܶ�ͣ��ҩƷҩƷ�����޸ģ�", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If Val(Me.tvwClass.Tag) = 2 Or (Val(Me.tvwClass.Tag) = 4 And mstrType = "7") Then
        With frmMediHerbalSpec
            .stbSpec.Tag = "�޸�"
            .lngҩ��id = Mid(Me.lvwItems.SelectedItem.Key, 2)
            .lngҩƷID = Mid(Me.lvwSpecs.SelectedItem.Key, 2)
            .strPrivs = Me.mstrPrivs
            lngҩƷID = .lngҩƷID
            .Show 1, Me
        End With
    Else
        With frmMediSpec
            .stbSpec.Tag = "�޸�"
            .lngҩ��id = Mid(Me.lvwItems.SelectedItem.Key, 2)
            .lngҩƷID = Mid(Me.lvwSpecs.SelectedItem.Key, 2)
            .strPrivs = Me.mstrPrivs
            lngҩƷID = .lngҩƷID
            .Show 1, Me
        End With
    End If
    
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
    '��λ�޸ĵ�ҩƷ
    On Error Resume Next
    err = 0
    Set lvwSpecs.SelectedItem = lvwSpecs.ListItems("_" & lngҩƷID)
    If err <> 0 Then Set lvwSpecs.SelectedItem = lvwSpecs.ListItems(1)
End Sub

Private Sub mnuEditSpecProtocol_Click()
    With frmMediMember
        Dim strType As String
        If Me.tvwClass.Tag = 4 Or Me.tvwClass.Tag = 3 Then '����������˽��
            If Me.lvwItems.SelectedItem Is Nothing Then     '���û�м�¼���˳�����Ϊ�޷��ж�ҩƷ����
                Exit Sub
            End If
            strType = Me.lvwItems.SelectedItem.SubItems(Me.lvwItems.ColumnHeaders("_����").Index - 1)
            strType = Switch(strType = "1", "5", strType = "2", "6", strType = "3", "7")
            .Tag = strType
        Else
            .Tag = Switch(Me.tvwClass.Tag = "0", "5", Me.tvwClass.Tag = "1", "6", Me.tvwClass.Tag = "2", "7")
        End If
        If InStr(1, mstrPrivs, "Э��ҩƷ����") = 0 Then
            .cmdClose.Tag = "����"
        Else
            .cmdClose.Tag = "�޸�"
        End If
        If Me.lvwSpecs.SelectedItem Is Nothing Then
            .lblMedi.Tag = 0
        Else
            .lblMedi.Tag = Mid(Me.lvwSpecs.SelectedItem.Key, 2)
        End If
        .msfMember.Tag = "Э��"
        .Show 1, Me
    End With
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
End Sub

Private Sub mnuEditSpecSelf_Click()
    With frmMediMember
        Dim strType As String
        If Me.tvwClass.Tag = 4 Or Me.tvwClass.Tag = 3 Then '����������˽��
            If Me.lvwItems.SelectedItem Is Nothing Then     '���û�м�¼���˳�����Ϊ�޷��ж�ҩƷ����
                Exit Sub
            End If
            strType = Me.lvwItems.SelectedItem.SubItems(Me.lvwItems.ColumnHeaders("_����").Index - 1)
            strType = Switch(strType = "1", "5", strType = "2", "6", strType = "3", "7")
            .Tag = strType
        Else
            .Tag = Switch(Me.tvwClass.Tag = "0", "5", Me.tvwClass.Tag = "1", "6", Me.tvwClass.Tag = "2", "7")
        End If
        If InStr(1, mstrPrivs, "����ҩƷ����") = 0 Then
            .cmdClose.Tag = "����"
        Else
            .cmdClose.Tag = "�޸�"
        End If
        If Me.lvwSpecs.SelectedItem Is Nothing Then
            .lblMedi.Tag = 0
        Else
            .lblMedi.Tag = Mid(Me.lvwSpecs.SelectedItem.Key, 2)
        End If
        .msfMember.Tag = "����"
        .Show 1, Me
    End With
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
End Sub

Private Sub mnuEditSpecUnit_Click()
    On Error Resume Next
    '�б�ҩƷ�б굥λ����
    With frmMediUnit
        Dim strType As String
        If Me.tvwClass.Tag = 4 Or Me.tvwClass.Tag = 3 Then '����������˽��
            If Me.lvwItems.SelectedItem Is Nothing Then     '���û�м�¼���˳�����Ϊ�޷��ж�ҩƷ����
                Exit Sub
            End If
            strType = Me.lvwItems.SelectedItem.SubItems(Me.lvwItems.ColumnHeaders("_����").Index - 1)
            strType = Switch(strType = "1", "5", strType = "2", "6", strType = "3", "7")
            .frmTag = strType
        Else
            .frmTag = Switch(Me.tvwClass.Tag = "0", "5", Me.tvwClass.Tag = "1", "6", Me.tvwClass.Tag = "2", "7")
        End If
        If Me.lvwSpecs.SelectedItem Is Nothing Then
            .lblTag = 0
        Else
            .lblTag = Mid(Me.lvwSpecs.SelectedItem.Key, 2)
        End If
        .strPrivs = Me.mstrPrivs
        .Show 1, Me
    End With
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
End Sub

Private Sub mnuEditStart_Click()
    If Me.ActiveControl.Name = Me.lvwItems.Name Then
        With Me.lvwItems
            If .SelectedItem Is Nothing Then Exit Sub
            If .SelectedItem.Icon = "��ҩU" Or .SelectedItem.Icon = "��ҩU" Then Exit Sub
            
            If MsgBox("����������á�" & .SelectedItem.Text & "����", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
'            If Val(Me.tvwClass.Tag) < 2 Or (Val(Me.tvwClass.Tag) = 3 And mstrType <> "7") Then
                gstrSql = "zl_��ҩƷ��_REUSE(" & Mid(.SelectedItem.Key, 2) & ")"
'            ElseIf Val(Me.tvwClass.Tag) = 2 Or (Val(Me.tvwClass.Tag) = 3 And mstrType = "7") Then
'                gstrSql = "zl_��ҩҩƷ_REUSE(" & Mid(.SelectedItem.Key, 2) & ")"
'            End If
            err = 0: On Error GoTo ErrHand
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            If Val(Me.tvwClass.Tag) < 2 Or Val(Me.tvwClass.Tag) = 4 And mstrType <> "7" Then
                .SelectedItem.Icon = "��ҩU": .SelectedItem.SmallIcon = "��ҩU"
            ElseIf Val(Me.tvwClass.Tag) = 2 Or Val(Me.tvwClass.Tag) = 4 And mstrType = "7" Then
                .SelectedItem.Icon = "��ҩU": .SelectedItem.SmallIcon = "��ҩU"
            End If
            '�ָ�������Ŀ��ʾ��ɫ
            .SelectedItem.ForeColor = .ForeColor
            For intCount = 1 To .ColumnHeaders.Count - 1
                .SelectedItem.ListSubItems(intCount).ForeColor = .ForeColor
            Next
        End With
    Else
        With Me.lvwSpecs
            If .Visible = False Then Exit Sub
            If .SelectedItem Is Nothing Then Exit Sub
            If .SelectedItem.Icon = "���U" Then Exit Sub
            
            strTemp = Me.lvwItems.SelectedItem.Text & " " & .SelectedItem.Text & " " & .SelectedItem.SubItems(.ColumnHeaders("����").Index - 1)
            If MsgBox("����������á�" & strTemp & "����", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
            gstrSql = "zl_��ҩ���_REUSE(" & Mid(.SelectedItem.Key, 2) & ")"
            err = 0: On Error GoTo ErrHand
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            If Val(Me.tvwClass.Tag) < 2 Or Val(Me.tvwClass.Tag) = 4 And mstrType <> "7" Then
                .SelectedItem.Icon = "���U": .SelectedItem.SmallIcon = "���U"
            Else
                .SelectedItem.Icon = "�ݹ�U": .SelectedItem.SmallIcon = "�ݹ�U"
            End If
            '�ָ�������Ŀ��ʾ��ɫ
            .SelectedItem.ForeColor = .ForeColor
            For intCount = 1 To .ColumnHeaders.Count - 1
                .SelectedItem.ListSubItems(intCount).ForeColor = .ForeColor
            Next
        End With
    End If
    
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditStop_Click()
    Dim lngҩƷID As Long
    Dim rsTemp As ADODB.Recordset
    Dim blnStop As Boolean
    
    If Me.ActiveControl.Name = Me.lvwItems.Name Then
        With Me.lvwItems
            If .SelectedItem Is Nothing Then Exit Sub
            If .SelectedItem.Icon = "��ҩS" Or .SelectedItem.Icon = "��ҩS" Then Exit Sub
            
            gstrSql = "select b.ʵ������ from ҩƷ��� a,ҩƷ��� b where a.ҩƷid=b.ҩƷid and a.ҩ��id=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "�����", Mid(.SelectedItem.Key, 2))
            
            If rsTemp.RecordCount > 0 Then
                If IIf(IsNull(rsTemp!ʵ������), "0", rsTemp!ʵ������) > 0 Then
                    If MsgBox("��ҩƷ�п��������ȷ��ͣ�ã�", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Sub
                    End If
                    blnStop = True
                End If
            End If
            
            If blnStop = False Then
                If MsgBox("���Ҫͣ�á�" & .SelectedItem.Text & "����", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
            End If
            
'            If Val(Me.tvwClass.Tag) < 2 Or (Val(Me.tvwClass.Tag) = 3 And mstrType <> "7") Then
                gstrSql = "zl_��ҩƷ��_STOP(" & Mid(.SelectedItem.Key, 2) & ")"
'            ElseIf Val(Me.tvwClass.Tag) = 2 Or (Val(Me.tvwClass.Tag) = 3 And mstrType = "7") Then
'                gstrSql = "zl_��ҩҩƷ_STOP(" & Mid(.SelectedItem.Key, 2) & ")"
'            End If
            err = 0: On Error GoTo ErrHand
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            If Me.mnuViewStoped.Checked = True Then
                If Val(Me.tvwClass.Tag) < 2 Or (Val(Me.tvwClass.Tag) = 4 And mstrType <> "7") Then
                    .SelectedItem.Icon = "��ҩS": .SelectedItem.SmallIcon = "��ҩS"
                ElseIf Val(Me.tvwClass.Tag) = 2 Or (Val(Me.tvwClass.Tag) = 4 And mstrType = "7") Then
                    .SelectedItem.Icon = "��ҩS": .SelectedItem.SmallIcon = "��ҩS"
                End If
                '��ͣ����Ŀ��ʾΪ��ɫ
                .SelectedItem.ForeColor = mconColor_Stop
                For intCount = 1 To .ColumnHeaders.Count - 1
                    .SelectedItem.ListSubItems(intCount).ForeColor = mconColor_Stop
                Next
            Else
                Call .ListItems.Remove(.SelectedItem.Key)
            End If
        End With
    Else
        With Me.lvwSpecs
            If .Visible = False Then Exit Sub
            If .SelectedItem Is Nothing Then Exit Sub
            If .SelectedItem.Icon = "���S" Then Exit Sub
            
            gstrSql = "select ʵ������ from ҩƷ��� where ҩƷid=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "�����", Mid(.SelectedItem.Key, 2))
            
            If rsTemp.RecordCount > 0 Then
                If IIf(IsNull(rsTemp!ʵ������), "0", rsTemp!ʵ������) > 0 Then
                    If MsgBox("��ҩƷ�п��������ȷ��ͣ�ã�", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Sub
                    End If
                    blnStop = True
                End If
            End If
            
            strTemp = Me.lvwItems.SelectedItem.Text & " " & .SelectedItem.Text & " " & .SelectedItem.SubItems(.ColumnHeaders("����").Index - 1)
            
            If blnStop = False Then
                If MsgBox("���Ҫͣ�á�" & .SelectedItem.Text & "����", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
            End If
            
            gstrSql = "zl_��ҩ���_STOP(" & Mid(.SelectedItem.Key, 2) & ")"
            err = 0: On Error GoTo ErrHand
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            If Me.mnuViewStoped.Checked = True Then
                If Val(Me.tvwClass.Tag) < 2 Or (Val(Me.tvwClass.Tag) = 4 And mstrType <> "7") Then
                    .SelectedItem.Icon = "���S": .SelectedItem.SmallIcon = "���S"
                Else
                    .SelectedItem.Icon = "�ݹ�S": .SelectedItem.SmallIcon = "�ݹ�S"
                End If
                '��ͣ����Ŀ��ʾΪ��ɫ
                .SelectedItem.ForeColor = mconColor_Stop
                For intCount = 1 To .ColumnHeaders.Count - 1
                    .SelectedItem.ListSubItems(intCount).ForeColor = mconColor_Stop
                Next
            Else
                Call .ListItems.Remove(.SelectedItem.Key)
            End If
        End With
    End If
    
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuFileExcel_Click()
    Call zlRptPrint(3)
End Sub

Private Sub mnuFilePara_Click()
    'ģ�鹫�������Ѿ�������ҩƷ��������ģ�飬Ŀǰû��˽�л򱾻���������ʱ���β������ý���
'    frmMediPara.ShowMe mstrPrivs, Me
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

Private Sub mnuFilter_Click()
    With frmMediFilter
        Call .ShowMe(Me, mnuViewStoped.Checked, mbln�Թ�ҩ)
     End With
End Sub

Private Sub mnuhelpAbout_Click()
    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
End Sub

Private Sub mnuHelpHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuPriceChargeSet1_Click()
    Call mnuPriceChargeSet_Click
End Sub

Private Sub mnuPriceLists_Click()
    Dim str��� As String
    Dim lng����id As Long
    
    Select Case Val(Me.tvwClass.Tag)
    Case 0
        str��� = "5"
    Case 1
        str��� = "6"
    Case 2
        str��� = "7"
    End Select
    
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    lng����id = Val(Mid(Me.tvwClass.SelectedItem.Key, 2))
    
    Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1023_2", "ZL8_BILL_1023_2"), Me, "���=" & str���, "����=" & lng����id)
End Sub


Private Sub mnuPriceTable_Click()
    If Me.tvwClass.SelectedItem Is Nothing Then
        Exit Sub
    End If
    Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "Zl1_BILL_1023_1", "ZL8_BILL_1023_1"), Me)
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    'Ĭ�ϲ���������=����id��Ʒ��=ҩ��id�����=ҩƷid
    Dim lng����id As Long
    Dim lngҩ��id As Long
    Dim lng���id As Long
    
    If Me.tvwClass.Tag <> 3 Then
        If Not Me.tvwClass.SelectedItem Is Nothing Then
            lng����id = Val(Mid(Me.tvwClass.SelectedItem.Key, 2))
        End If
    End If
    
    If Not Me.lvwItems.SelectedItem Is Nothing Then
        lngҩ��id = Mid(lvwItems.SelectedItem.Key, 2)
    End If
    
    If Not Me.lvwSpecs.SelectedItem Is Nothing Then
        lng���id = Mid(lvwSpecs.SelectedItem.Key, 2)
    End If
    
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
        "����=" & IIf(lng����id = 0, "", lng����id), _
        "Ʒ��=" & IIf(lngҩ��id = 0, "", lngҩ��id), _
        "���=" & IIf(lng���id = 0, "", lng���id))
End Sub

Private Sub mnuUploadDrugInfo_Click()
    '�����ϴ�ҩƷ��Ϣ
    If Not gobjLogisticPlatform Is Nothing Then
        gobjLogisticPlatform.UploadDrugInfo Me, gcnOracle, 0
    End If
End Sub

Private Sub mnuViewFind_Click()
    With frmMediFind
        Call .ShowMe(Me, mnuViewStoped.Checked, mbln�Թ�ҩ)
    End With
End Sub

Private Sub mnuViewFindNext_Click()
    On Error Resume Next
    
    Select Case Val(tvwClass.Tag)
    Case 0
        frmMediFind.Tag = 5: Me.Caption = "����ҩ����..."
    Case 1
        frmMediFind.Tag = 6: Me.Caption = "�г�ҩ����..."
    Case 2
        frmMediFind.Tag = 7: Me.Caption = "�в�ҩ����..."
    End Select
    Call frmMediFind.FindNext
End Sub

Private Sub mnuViewList_Click()
    mstrFindValue = ""
    Set mrsFind = Nothing
    If Me.tvwClass.SelectedItem Is Nothing Then
        Exit Sub
    End If
    Me.mnuViewList.Checked = Not Me.mnuViewList.Checked
    Call zlRefClasses
End Sub
Private Sub mnuViewPrices_Click()
    Me.mnuViewPrices.Checked = Not Me.mnuViewPrices.Checked
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
End Sub

Private Sub mnuViewRefer_Click()
    Call gobjKernel.InitCISKernel(gcnOracle, Me, glngSys, mstrPrivs)
    If Me.lvwItems.SelectedItem Is Nothing Then
        Call gobjKernel.ShowClincHelp(0, Me)
    Else
        Call gobjKernel.ShowClincHelp(0, Me, Val(Mid(Me.lvwItems.SelectedItem.Key, 2)))
    End If
End Sub

Private Sub mnuViewRefresh_Click()
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    Call zlRefRecords
End Sub

Private Sub mnuViewShowAll_Click()
    On Error GoTo ErrHandle
    mnuViewShowAll.Checked = Not mnuViewShowAll.Checked
    If tvwClass.SelectedItem Is Nothing Then
        If tvwClass.Nodes.Count > 0 Then
            MsgBox "��ѡ��һ�·��࣡", vbInformation, gstrSysName
        Else
            MsgBox "���κη������ʾ��", vbInformation, gstrSysName
        End If
        Exit Sub
    End If
    Call zlRefRecords
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuViewStates_Click()
    Me.mnuViewStates.Checked = Not Me.mnuViewStates.Checked
    Me.stbThis.Visible = Me.mnuViewStates.Checked
    Form_Resize
End Sub

Private Sub mnuViewStoped_Click()
    mstrFindValue = ""
    Set mrsFind = Nothing
    If Me.tvwClass.SelectedItem Is Nothing Then
        Exit Sub
    End If
    Me.mnuViewStoped.Checked = Not Me.mnuViewStoped.Checked
    Call zlRefRecords
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
    err = 0: On Error Resume Next
    For intCount = Me.cmdKind.LBound To Me.cmdKind.UBound
        Me.cmdKind(intCount).Left = Me.picClass.ScaleLeft + 15
        Me.cmdKind(intCount).Width = Me.picClass.ScaleWidth
        Me.cmdKind(intCount).Height = 300
        
        If intCount = 2 And mbln�Թ�ҩ = True Then '�Թ�ҩ����ʾ�в�ҩ ��������
            cmdKind(intCount).Visible = False
        End If
        If intCount <= mintIndex Then
            If mbln�Թ�ҩ = True And intCount > 2 Then
                Me.cmdKind(intCount).Top = Me.picClass.ScaleTop + 285 * (intCount - 1)
                Me.tvwClass.Top = Me.picClass.ScaleTop + 285 * intCount
            Else
                Me.cmdKind(intCount).Top = Me.picClass.ScaleTop + 285 * intCount
                Me.tvwClass.Top = Me.picClass.ScaleTop + 285 * (intCount + 1)
            End If
        Else
            If mbln�Թ�ҩ = True And intCount > 2 Then
                Me.cmdKind(intCount).Top = Me.picClass.ScaleHeight - 285 * (Me.cmdKind.UBound - intCount + 2)
            Else
                Me.cmdKind(intCount).Top = Me.picClass.ScaleHeight - 285 * (Me.cmdKind.UBound - intCount + 1)
            End If
        End If
    Next
    Me.tvwClass.Left = Me.picClass.ScaleLeft + 15
    Me.tvwClass.Width = Me.picClass.ScaleWidth
    Me.tvwClass.Height = Me.picClass.ScaleHeight - 285 * (Me.cmdKind.UBound + 1) - 15
End Sub

Private Sub picHBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Me.picHBar.Top = Me.picHBar.Top + y
    End If
End Sub

Private Sub picHBar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Call Form_Resize
    End If
End Sub

Private Sub picVBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Me.picVBar.Left = Me.picVBar.Left + x
    End If
End Sub

Private Sub picVBar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Call Form_Resize
    End If
End Sub

Private Sub tabContent_Click(PreviousTab As Integer)
    mintPage = Me.tabContent.Tab
    Select Case Me.tabContent.Tab
    Case 0
        Me.lvwSpecs.Visible = True
        Me.fraComment(0).Visible = True
        Me.hgdPrice.Visible = False
        Me.fraComment(1).Visible = False
        Me.hgdCost.Visible = False
    Case 1
        Me.lvwSpecs.Visible = False
        Me.fraComment(0).Visible = False
        Me.hgdPrice.Visible = True
        Me.fraComment(1).Visible = True
        Me.hgdCost.Visible = False
    Case 2
        Me.lvwSpecs.Visible = False
        Me.fraComment(0).Visible = False
        Me.hgdPrice.Visible = False
        Me.fraComment(1).Visible = False
        Me.hgdCost.Visible = True
        hgdCharge.Visible = False
        Call GetCostAdjust(mlngCurrDrug)
    Case 3
        Me.lvwSpecs.Visible = False
        Me.fraComment(0).Visible = False
        Me.hgdPrice.Visible = False
        Me.fraComment(1).Visible = False
        hgdCost.Visible = False
        Me.hgdCharge.Visible = True
        Call GetChargeSet(mlngCurrDrug)
    End Select
End Sub

Private Sub tlbThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Preview"
        Call mnuFilePreview_Click
    Case "Print"
        Call mnuFilePrint_Click
    Case "Class"
        Call PopupMenu(Me.mnuClass, 2)
    Case "Item"
        Call zlPopupEditMenu(1, False)
    Case "Spec"
        Call zlPopupEditMenu(2, False)
    Case "Start"
        If Me.ActiveControl Is tvwClass Then
            Call mnuClassStar_Click
        Else
            Call mnuEditStart_Click
        End If
    Case "Stop"
        If Me.ActiveControl Is tvwClass Then
            Call mnuClassStop_Click
        Else
            Call mnuEditStop_Click
        End If
    Case "Limit"
        Call mnuEditSpecLimit_Click
    Case "Find"
        Call mnuViewFind_Click
    Case "Help"
        Call mnuHelpHelp_Click
    Case "Exit"
        Call mnuFileExit_Click
    Case "Filter"
        Call mnuFilter_Click
    End Select
End Sub

Private Sub tlbThis_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    PopupMenu Me.mnuViewToolbar, 2
End Sub

Private Sub tvwClass_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    If Val(Me.tvwClass.Tag) >= 3 Then           '=3����ʾ���˽����״̬��������༭���
        Exit Sub
    End If
    If mbln�Թ�ҩ = False Then
        Call zlPopupClassMenu
    End If
End Sub

Private Sub tvwClass_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim bln���� As Boolean
    Dim blnͣ�� As Boolean
    Dim objItem As ListItem
    Dim strTemp As String
    Dim strNode As String
    
    If mstrKey <> Node.Key Then
        mstrKey = Node.Key
    Else
        Exit Sub
    End If
    
    If InStr(1, mstrPrivs, ";ҩƷ����;") = 0 Then
        mnuClassStar.Visible = False
        tlbThis.Buttons("Start").Visible = False
    Else
        mnuClassStar.Visible = True
        tlbThis.Buttons("Start").Visible = True
    End If
    If InStr(1, mstrPrivs, ";ҩƷͣ��;") = 0 Then
        mnuClassStop.Visible = False
        tlbThis.Buttons("Stop").Visible = False
    Else
        mnuClassStop.Visible = True
        tlbThis.Buttons("Stop").Visible = True
    End If
    If tlbThis.Buttons("Start").Visible = False Or tlbThis.Buttons("Stop").Visible = False Then
        tlbThis.Buttons("sp2").Visible = False
    End If
    
    Call zlRefRecords
    
    bln���� = mnuClassStar.Visible
    blnͣ�� = mnuClassStop.Visible
    
    If mstr���� = "0" Then
        mstrNodeSelect����ҩ = tvwClass.SelectedItem.Key
        strTemp = mstrItemClick����ҩ
        strNode = mstrNodeClick����ҩ
    ElseIf mstr���� = "1" Then
        mstrNodeSelect�г�ҩ = tvwClass.SelectedItem.Key
        strTemp = mstrItemClick�г�ҩ
        strNode = mstrNodeClick�г�ҩ
    ElseIf mstr���� = "2" Then
        mstrNodeSelect�в�ҩ = tvwClass.SelectedItem.Key
        strTemp = mstrItemClick�в�ҩ
        strNode = mstrNodeClick�в�ҩ
    End If
    
    If strNode = Node.Key Then    '��λҩƷ
        For Each objItem In lvwItems.ListItems
            If objItem.Key = strTemp Then
                lvwItems.ListItems(objItem.Key).Selected = True
                Exit For
            End If
        Next
    End If
        
    With tvwClass
        If .Nodes.Count = 0 Then Exit Sub
        If lvwItems.SelectedItem Is Nothing Then
            mnuEditItemMod.Enabled = False
            mnuEditItemDel.Enabled = False
            mnuEditItemBill.Enabled = False
            mnuEditStart.Enabled = False
            mnuEditStop.Enabled = False
            tlbThis.Buttons("Start").Enabled = False
            tlbThis.Buttons("Stop").Enabled = False
            mnuEditSpecAdd.Enabled = False
            mnuEditSpecMod.Enabled = False
            mnuEditSpecDel = False
        Else
            mnuEditItemMod.Enabled = True
            mnuEditItemDel.Enabled = True
            mnuEditItemBill.Enabled = True
            mnuEditStart.Enabled = bln����
            mnuEditStop.Enabled = blnͣ��
            tlbThis.Buttons("Start").Enabled = bln����
            tlbThis.Buttons("Stop").Enabled = blnͣ��
            If lvwItems.SelectedItem.ForeColor = vbRed Then
                mnuEditSpecAdd.Enabled = False
                mnuEditSpecMod.Enabled = False
                mnuEditSpecDel = True
            Else
                mnuEditSpecAdd.Enabled = True
                mnuEditSpecMod.Enabled = True
                mnuEditSpecDel = True
            End If
            If lvwSpecs.SelectedItem Is Nothing Then
                mnuEditSpecMod.Enabled = False
                mnuEditSpecDel = False
            End If
        End If
    End With
End Sub

Private Sub zlRefPurview()
    '---------------------------------------------
    '��д���Ʒ�����Ŀ(�˴�ΪҩƷ����)�����ղ�ͬ���͵�������
    '---------------------------------------------
    '����Ȩ�޿���
    Me.mnuEdit.Enabled = True
    Me.mnuPrice.Enabled = True
    If Val(Me.tvwClass.Tag) = 0 Then
        If InStr(1, mstrPrivs, "��������ҩ") = 0 Then
            mnuClassAdd.Visible = False
            mnuClassDel.Visible = False
            mnuClassMod.Visible = False
        Else
            mnuClass.Visible = True
            tlbThis.Buttons("Class").Visible = True
            mnuClassAdd.Visible = True
            mnuClassDel.Visible = True
            mnuClassMod.Visible = True
        End If
        If InStr(1, mstrPrivs, "��������ҩƷ��") = 0 Then
            mnuEditItemAdd.Visible = False
            mnuEditItemDel.Visible = False
            mnuEditItemMod.Visible = False
        Else
            mnuEditItemAdd.Visible = True
            mnuEditItemDel.Visible = True
            mnuEditItemMod.Visible = True
        End If
        If InStr(1, mstrPrivs, "��������ҩ���") = 0 Then
            mnuEditSpecAdd.Visible = False
            mnuEditSpecDel.Visible = False
            mnuEditSpecMod.Visible = False
        Else
            mnuEditSpecAdd.Visible = True
            mnuEditSpecDel.Visible = True
            mnuEditSpecMod.Visible = True
        End If
    End If
    If Val(Me.tvwClass.Tag) = 1 Then
        If InStr(1, mstrPrivs, "�����г�ҩ") = 0 Then
            mnuClassAdd.Visible = False
            mnuClassDel.Visible = False
            mnuClassMod.Visible = False
        Else
            mnuClass.Visible = True
            tlbThis.Buttons("Class").Visible = True
            mnuClassAdd.Visible = True
            mnuClassDel.Visible = True
            mnuClassMod.Visible = True
        End If
        If InStr(1, mstrPrivs, "�����г�ҩƷ��") = 0 Then
            mnuEditItemAdd.Visible = False
            mnuEditItemDel.Visible = False
            mnuEditItemMod.Visible = False
        Else
            mnuEditItemAdd.Visible = True
            mnuEditItemDel.Visible = True
            mnuEditItemMod.Visible = True
        End If
        If InStr(1, mstrPrivs, "�����г�ҩ���") = 0 Then
            mnuEditSpecAdd.Visible = False
            mnuEditSpecDel.Visible = False
            mnuEditSpecMod.Visible = False
        Else
            mnuEditSpecAdd.Visible = True
            mnuEditSpecDel.Visible = True
            mnuEditSpecMod.Visible = True
        End If
    End If
    If Val(Me.tvwClass.Tag) = 2 Then
        If InStr(1, mstrPrivs, "�����в�ҩ") = 0 Then
            mnuClassAdd.Visible = False
            mnuClassDel.Visible = False
            mnuClassMod.Visible = False
        Else
            mnuClass.Visible = True
            tlbThis.Buttons("Class").Visible = True
            mnuClassAdd.Visible = True
            mnuClassDel.Visible = True
            mnuClassMod.Visible = True
        End If
        If InStr(1, mstrPrivs, "�����в�ҩƷ��") = 0 Then
            mnuEditItemAdd.Visible = False
            mnuEditItemDel.Visible = False
            mnuEditItemMod.Visible = False
        Else
            mnuEditItemAdd.Visible = True
            mnuEditItemDel.Visible = True
            mnuEditItemMod.Visible = True
        End If
        If InStr(1, mstrPrivs, "�����в�ҩ���") = 0 Then
            mnuEditSpecAdd.Visible = False
            mnuEditSpecDel.Visible = False
            mnuEditSpecMod.Visible = False
        Else
            mnuEditSpecAdd.Visible = True
            mnuEditSpecDel.Visible = True
            mnuEditSpecMod.Visible = True
        End If
    End If
    If Val(Me.tvwClass.Tag) = 4 And InStr(1, mstrPrivs, "��������ҩ") = 0 _
        And InStr(1, mstrPrivs, "�����г�ҩ") = 0 _
        And InStr(1, mstrPrivs, "�����в�ҩ") = 0 Then
        Me.mnuEdit.Enabled = False: Me.mnuPrice.Enabled = False
    End If
    
    If InStr(1, mstrPrivs, ";ҩƷ����;") = 0 Then
        mnuClassStar.Visible = False
        mnuEditStart.Visible = False
        tlbThis.Buttons("Start").Visible = False
    Else
        mnuClassStar.Visible = True
        mnuEditStart.Visible = True
        tlbThis.Buttons("Start").Visible = True
    End If
    If InStr(1, mstrPrivs, ";ҩƷͣ��;") = 0 Then
        mnuClassStop.Visible = False
        mnuEditStop.Visible = False
        tlbThis.Buttons("Stop").Visible = False
    Else
        mnuClassStop.Visible = True
        mnuEditStop.Visible = True
        tlbThis.Buttons("Stop").Visible = True
    End If
    If tlbThis.Buttons("Start").Visible = False Or tlbThis.Buttons("Stop").Visible = False Then
        tlbThis.Buttons("sp2").Visible = False
    End If
    
    If Val(Me.tvwClass.Tag) = 2 Then
        Me.mnuEditItemUsage.Visible = False
    Else
        Me.mnuEditItemUsage.Visible = True
    End If
    
    If InStr(1, mstrPrivs, "��Ӧ����") = 0 Then
        Me.mnuEditItemBill.Enabled = False
    Else
        Me.mnuEditItemBill.Enabled = Me.mnuEdit.Enabled
    End If
    
    If InStr(1, mstrPrivs, "�ۼ۹���") = 0 And InStr(1, mstrPrivs, "�ɱ��۹���") = 0 Then
    Else
        If InStr(mstrPrivs, "�ѱ�����") = 0 Then
            mnuPriceChargeSet.Visible = False
        Else
            mnuPriceChargeSet.Visible = True
            mnuPriceChargeSet.Enabled = Me.mnuPrice.Enabled
        End If
    End If
    Me.tlbThis.Buttons("Limit").Enabled = Me.mnuEditSpecLimit.Enabled
    
    If InStr(1, mstrPrivs, "���ۼ�¼��ѯ") = 0 Then
        Me.mnuPriceLists.Visible = False
        Me.mnuPriceTable.Visible = False
    Else
        Me.mnuPriceSpt1.Visible = Me.mnuPrice.Enabled
        Me.mnuPriceLists.Visible = Me.mnuPrice.Enabled
        Me.mnuPriceTable.Visible = Me.mnuPrice.Enabled
    End If
    
    '������ʾ����
    If Val(Me.tvwClass.Tag) > 2 Then '����2��Ϊ����ҩ��͹��˽��
        Me.mnuClass.Visible = False
        Me.tlbThis.Buttons("Class").Visible = False
    End If
End Sub

Private Sub zlRefClasses(Optional lngNode As Long)
    Dim intCol As Integer
    Dim intType As Integer
    Dim i As Integer
    
    On Error GoTo ErrHand
    
    Call zlRefPurview 'Ȩ����֤
    
    If Val(Me.tvwClass.Tag) < 3 Then
        '����ҩ���г�ҩ����ҩ
'        Me.mnuEditSpecAdd.Visible = True
'        Me.mnuEditSpecMod.Visible = True
'        Me.mnuEditSpecDel.Visible = True
'        Me.tlbThis.Buttons("Spec").Visible = True
        Me.lvwItems.ListItems.Clear
        With Me.lvwItems.ColumnHeaders
            .Clear
            .Add , "_����", "����", 2500
            .Add , "_����", "����", 1000
            .Add , "_������λ", "������λ", 900
            .Add , "_����", "����", 1800
            '.Add , "_�������", "�������", 1100
            .Add , "_��������", "��������", 1000
            .Add , "_��������", "��������", 900
            .Add , "_��������", "��������", 900
            .Add , "_����", "����", 750
            .Add , "_��Դ", "��Դ", 600
            .Add , "_��ֵ", "��ֵ", 600
            .Add , "_�ݴ�", "�ݴ�", 600
            .Add , "_ԭ��ҩ", "ԭ��ҩ", 750
            If Val(Me.tvwClass.Tag) = 2 Then
                .Add , "_��ζʹ��", "��ζʹ��", 900
            Else
                .Add , "_����ҩ", "����ҩ", 750
                .Add , "_��ҩ", "��ҩ", 600
                .Add , "_ԭ��ҩ", "ԭ��ҩ", 800
                .Add , "_ר��ҩ", "ר��ҩ", 800
                .Add , "_��������", "��������", 900
                .Add , "_����ҩ��", "����ҩ��", 1500
            End If
            .Add , "_��ҩƷ�³���ҽ��", "��ҩƷ�³���ҽ��", 1500
            .Add , "_�����Ա�", "�����Ա�", 1200
            .Add , "_������ҩ", "������ҩ", 1200
        End With
        With Me.lvwItems
            .ColumnHeaders("_����").Position = 1
            .SortKey = .ColumnHeaders("_����").Index - 1
            .SortOrder = lvwAscending
        End With
        
        '���������
        Me.lvwSpecs.ListItems.Clear
        With Me.lvwSpecs.ColumnHeaders
            .Clear
            .Add , "���", "���", 1500: .Add , "����", "����", 1100
            If Not (Val(Me.tvwClass.Tag) = 2 Or Val(Me.tvwClass.Tag) = 4 And mstrType = "7") Then
                .Add , "����", "������", 2000
                .Add , "��Ʒ��", "��Ʒ��", 2000
            Else
                .Add , "����", "������", 2000
                .Add , "ԭ����", "ԭ����", 2000
            End If
            .Add , "�������", "�������", 1200: .Add , "ҽ������", "ҽ������", 900: .Add , "ҩƷ��Դ", "ҩƷ��Դ", 900: .Add , "����", "����", 600
            .Add , "Э��", "Э��", 600
            If Val(Me.tvwClass.Tag) = 2 Or Val(Me.tvwClass.Tag) = 4 And mstrType = "7" Then .Add , "��ҩ��̬", "��ҩ��̬", 900
            .Add , "�б�", "�б�", 600: .Add , "��׼�ĺ�", "��׼�ĺ�", 1600
            .Add , "��ͬ��λ", "��ͬ��λ", 3000: .Add , "˵��", "˵��", 2000
            .Add , "��ѡ��", "��ѡ��", 1000: .Add , "վ��", "Ժ��", IIf(gstrNodeNo = "-", 0, 1000)
        End With
        With Me.lvwSpecs
            .ColumnHeaders("����").Position = 1
            .SortKey = .ColumnHeaders("����").Index - 1
            .SortOrder = lvwAscending
        End With
        
        Me.tabContent.TabVisible(0) = True
        Me.tabContent.Tab = mintPage
        Call tabContent_Click(mintPage)
    End If
    
   
    With Me.hgdPrice
        .Rows = .FixedRows + 1: .RowData(.FixedRows) = 0
        For intCol = 0 To .Cols - 1
            .TextMatrix(.FixedRows, intCol) = ""
        Next
    End With
    Call RestoreListViewState(Me.lvwItems, Me.Name & Val(Me.tvwClass.Tag), Me.lvwItems.View)
    
    
    ''''''''''''''''''''''''''''''����ǹ��˽�����򵥶�����
    If Me.tvwClass.Tag >= 3 Then
        Me.tvwClass.Nodes.Clear
        Me.lvwItems.ListItems.Clear
        Me.lvwSpecs.ListItems.Clear
       
        '������������ʾ��
        Me.mnuEditItemAdd.Visible = True
        Me.mnuEditSpecAdd.Visible = True
        Me.mnuEditSpecMod.Visible = True
        Me.mnuEditSpecDel.Visible = True
        Me.mnuEditItemUsage.Visible = True
        Me.tlbThis.Buttons("Spec").Visible = True
        Me.lvwItems.ListItems.Clear
        With Me.lvwItems.ColumnHeaders
            .Clear
            .Add , "_����", "����", 2500
            .Add , "_����", "����", 1000
            .Add , "_������λ", "������λ", 900
            .Add , "_����", "����", 600
            '.Add , "_�������", "�������", 1100
            .Add , "_��������", "��������", 1000
            .Add , "_��������", "��������", 900
            .Add , "_��������", "��������", 900
            .Add , "_����", "����", 750
            .Add , "_��Դ", "��Դ", 600
            .Add , "_��ֵ", "��ֵ", 600
            .Add , "_�ݴ�", "�ݴ�", 600
            .Add , "_ԭ��ҩ", "ԭ��ҩ", 750
            If mstrType = "7" Then
                .Add , "_��ζʹ��", "��ζʹ��", 900
            Else
                .Add , "_����ҩ", "����ҩ", 750
                .Add , "_��ҩ", "��ҩ", 600
                .Add , "_ԭ��ҩ", "ԭ��ҩ", 800
                .Add , "_ר��ҩ", "ר��ҩ", 800
                .Add , "_��������", "��������", 900
                .Add , "_����ҩ��", "����ҩ��", 1500
            End If
            .Add , "_��ҩƷ�³���ҽ��", "��ҩƷ�³���ҽ��", 1500
            .Add , "_�����Ա�", "�����Ա�", 1200
            .Add , "_����", "����", 0
            .Add , "_����id", "����id", 0
            .Add , "_������ҩ", "������ҩ", 1200
        End With
        With Me.lvwItems
            .ColumnHeaders("_����").Position = 1
            .SortKey = .ColumnHeaders("_����").Index - 1
            .SortOrder = lvwAscending
        End With
        
        '���������
        Me.lvwSpecs.ListItems.Clear
        With Me.lvwSpecs.ColumnHeaders
            .Clear
            .Add , "���", "���", 1500: .Add , "����", "����", 1100
            If Not (Val(Me.tvwClass.Tag) = 2 Or Val(Me.tvwClass.Tag) = 4 And mstrType = "7") Then
                .Add , "����", "������", 2000
                .Add , "��Ʒ��", "��Ʒ��", 2000
            Else
                .Add , "����", "������", 2000
                .Add , "ԭ����", "ԭ����", 2000
            End If
            .Add , "�������", "�������", 1500: .Add , "ҽ������", "ҽ������", 900: .Add , "ҩƷ��Դ", "ҩƷ��Դ", 900: .Add , "����", "����", 600
            .Add , "Э��", "Э��", 600
            If Val(Me.tvwClass.Tag) = 2 Or Val(Me.tvwClass.Tag) = 4 And mstrType = "7" Then .Add , "��ҩ��̬", "��ҩ��̬", 900
            .Add , "�б�", "�б�", 600: .Add , "��׼�ĺ�", "��׼�ĺ�", 1600: .Add , "��ͬ��λ", "��ͬ��λ", 3000: .Add , "˵��", "˵��", 2000
            .Add , "��ѡ��", "��ѡ��", 1000: .Add , "վ��", "Ժ��", IIf(gstrNodeNo = "-", 0, 1000)
        End With
        With Me.lvwSpecs
            .ColumnHeaders("����").Position = 1
            .SortKey = .ColumnHeaders("����").Index - 1
            .SortOrder = lvwAscending
        End With
        
        Me.tabContent.TabVisible(0) = True
        Me.tabContent.Tab = mintPage
        Call tabContent_Click(mintPage)
        Call setMenu�Թ�ҩ
       If Val(Me.tvwClass.Tag) = 4 Then
       
            If mstrType = "" Then
                Exit Sub
            End If
            '���ù��˽����������
            Me.tvwClass.Visible = False
            Set objNode = Me.tvwClass.Nodes.Add(, , "_ALL", "���й��˽��", "close")
            If mstrType = "7" Then
                Set objNode = Me.tvwClass.Nodes.Add("_ALL", tvwChild, "_ALL7", "�в�ҩ", "close")
            ElseIf mstrType = "5,6" Then
                Set objNode = Me.tvwClass.Nodes.Add("_ALL", tvwChild, "_ALL5", "����ҩ", "close")
                Set objNode = Me.tvwClass.Nodes.Add("_ALL", tvwChild, "_ALL6", "�г�ҩ", "close")
            ElseIf mstrType = "5" Then
                Set objNode = Me.tvwClass.Nodes.Add("_ALL", tvwChild, "_ALL5", "����ҩ", "close")
            ElseIf mstrType = "6" Then
                Set objNode = Me.tvwClass.Nodes.Add("_ALL", tvwChild, "_ALL6", "����ҩ", "close")
            End If
            
            gstrSql = "select Distinct A.ID,A.�ϼ�ID,A.����,A.����,A.����,B.���,Nvl(To_Char(A.����ʱ��, 'YYYY-MM-DD'), '3000-01-01') ����ʱ�� " & _
                " From ���Ʒ���Ŀ¼ A,������ĿĿ¼ B,Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) C,Table(Cast(f_Num2List([2]) As zlTools.t_NumList)) D " & _
                " Where A.id=B.����id And B.���=C.Column_Value And B.id=D.Column_Value " & IIf(mnuViewList.Checked = False, " And Nvl(To_Char(A.����ʱ��, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01' ", "")
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mstrType, mstrDrugId)
            
            With rsTemp
                Do While Not .EOF
                    Set objNode = Me.tvwClass.Nodes.Add("_ALL" & !���, tvwChild, "_" & !ID, "[" & !���� & "]" & !����, "close")
                    objNode.Sorted = True
                    objNode.Tag = IIf(IsNull(!����), "", !����)
                    objNode.ExpandedImage = "expend"
                    If Format(!����ʱ��, "YYYY-MM-DD") <> "3000-01-01" Then
                        objNode.ForeColor = mconColor_Stop
                    End If
                    .MoveNext
                Loop
                Me.tvwClass.Visible = True
            End With
            Call setMenu�Թ�ҩ
        Else
            Me.tvwClass.Visible = False
            
            Set objNode = Me.tvwClass.Nodes.Add(, , "_����ҩ", "1-����ҩ", "close")
            objNode.Expanded = True
            Set objNode = Me.tvwClass.Nodes.Add(, , "_ԭ��ҩ", "2-ԭ��ҩ", "close")
            Set objNode = Me.tvwClass.Nodes.Add(, , "_ר��ҩ", "3-ר��ҩ", "close")
            Set objNode = Me.tvwClass.Nodes.Add(, , "_��������", "4-��������", "close")
            
            gstrSql = "Select Distinct a.Id, a.�ϼ�id, a.����, a.����, a.����, b.���, Nvl(e.������, 0) ������, Nvl(e.�Ƿ�ԭ��ҩ, 0) �Ƿ�ԭ��ҩ, Nvl(e.�Ƿ�ר��ҩ, 0) �Ƿ�ר��ҩ," & vbNewLine & _
                      " Nvl(e.�Ƿ񵥶�����, 0) �Ƿ񵥶�����, Nvl(To_Char(a.����ʱ��, 'YYYY-MM-DD'), '3000-01-01') ����ʱ��" & vbNewLine & _
                      " From ���Ʒ���Ŀ¼ A, ������ĿĿ¼ B, ҩƷ���� E" & vbNewLine & _
                      " where a.Id = b.����id And e.ҩ��id = b.Id And Nvl(To_Char(a.����ʱ��, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01'"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
            
            '����ҩ
            Set objNode = Me.tvwClass.Nodes.Add("_����ҩ", tvwChild, "_Limit1", "1-������ʹ��", "close")
            Set objNode = Me.tvwClass.Nodes.Add("_����ҩ", tvwChild, "_Limit2", "2-����ʹ��", "close")
            Set objNode = Me.tvwClass.Nodes.Add("_����ҩ", tvwChild, "_Limit3", "3-����ʹ��", "close")
            rsTemp.Filter = ""
            rsTemp.Filter = "������<>0"
            With rsTemp
                Do While Not .EOF
                    For i = 1 To Me.tvwClass.Nodes.Count
                        If Me.tvwClass.Nodes(i).Key = "A" & !������ & !ID Then
                            .MoveNext
                            i = 1
                            If .EOF Then Exit For
                        End If
                    Next
                    If Not .EOF Then
                        Set objNode = Me.tvwClass.Nodes.Add("_Limit" & !������, tvwChild, "A" & !������ & !ID, "[" & !���� & "]" & !����, "close")
                        objNode.Sorted = True
                        objNode.Tag = IIf(IsNull(!����), "", !����)
                        objNode.ExpandedImage = "expend"
                        If Format(!����ʱ��, "YYYY-MM-DD") <> "3000-01-01" Then
                            objNode.ForeColor = mconColor_Stop
                        End If
                        .MoveNext
                    End If
                Loop
            End With
            'ԭ��ҩ
            rsTemp.Filter = ""
            rsTemp.Filter = "�Ƿ�ԭ��ҩ=1"
            With rsTemp
                Do While Not .EOF
                    For i = 1 To Me.tvwClass.Nodes.Count
                        If Me.tvwClass.Nodes(i).Key = "B" & !ID Then
                            .MoveNext
                            i = 1
                            If .EOF Then Exit For
                        End If
                    Next
                    If Not .EOF Then
                        Set objNode = Me.tvwClass.Nodes.Add("_ԭ��ҩ", tvwChild, "B" & !ID, "[" & !���� & "]" & !����, "close")
                        objNode.Sorted = True
                        objNode.Tag = IIf(IsNull(!����), "", !����)
                        objNode.ExpandedImage = "expend"
                        If Format(!����ʱ��, "YYYY-MM-DD") <> "3000-01-01" Then
                            objNode.ForeColor = mconColor_Stop
                        End If
                        .MoveNext
                    End If
                Loop
            End With
            'ר��ҩ
            rsTemp.Filter = ""
            rsTemp.Filter = "�Ƿ�ר��ҩ=1"
            With rsTemp
                Do While Not .EOF
                    For i = 1 To Me.tvwClass.Nodes.Count
                        If Me.tvwClass.Nodes(i).Key = "C" & !ID Then
                            .MoveNext
                            i = 1
                            If .EOF Then Exit For
                        End If
                    Next
                    If Not .EOF Then
                        Set objNode = Me.tvwClass.Nodes.Add("_ר��ҩ", tvwChild, "C" & !ID, "[" & !���� & "]" & !����, "close")
                        objNode.Sorted = True
                        objNode.Tag = IIf(IsNull(!����), "", !����)
                        objNode.ExpandedImage = "expend"
                        If Format(!����ʱ��, "YYYY-MM-DD") <> "3000-01-01" Then
                            objNode.ForeColor = mconColor_Stop
                        End If
                        .MoveNext
                    End If
                Loop
            End With
            '��������
            rsTemp.Filter = ""
            rsTemp.Filter = "�Ƿ񵥶�����=1"
            With rsTemp
                Do While Not .EOF
                    For i = 1 To Me.tvwClass.Nodes.Count
                        If Me.tvwClass.Nodes(i).Key = "D" & !ID Then
                            .MoveNext
                            i = 1
                            If .EOF Then Exit For
                        End If
                    Next
                    If Not .EOF Then
                        Set objNode = Me.tvwClass.Nodes.Add("_��������", tvwChild, "D" & !ID, "[" & !���� & "]" & !����, "close")
                        objNode.Sorted = True
                        objNode.Tag = IIf(IsNull(!����), "", !����)
                        objNode.ExpandedImage = "expend"
                        If Format(!����ʱ��, "YYYY-MM-DD") <> "3000-01-01" Then
                            objNode.ForeColor = mconColor_Stop
                        End If
                        .MoveNext
                    End If
                Loop
            End With
            
            Me.tvwClass.Visible = True
            
            Call setMenu�Թ�ҩ
        End If
        Me.stbThis.Panels(2).Text = ""
        If Me.tvwClass.Nodes.Count > 0 Then
            If Val(Me.tvwClass.Tag) = 4 Then
                Me.tvwClass.Nodes("_ALL").Selected = True
            Else
                Me.tvwClass.Nodes("_Limit1").Selected = True
            End If
            Call zlRefRecords
        End If
        Exit Sub
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '��д����

    gstrSql = "select ID,�ϼ�ID,����,����,����,Nvl(To_Char(����ʱ��, 'YYYY-MM-DD'), '3000-01-01') ����ʱ�� " & _
            " From ���Ʒ���Ŀ¼" & _
            " Where ���� = [1] " & IIf(mnuViewList.Checked = False, " And Nvl(To_Char(����ʱ��, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01' ", "") & _
            " start with �ϼ�ID is null" & _
            " connect by prior ID=�ϼ�ID Order By Level, ���� "
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
            If Format(!����ʱ��, "YYYY-MM-DD") <> "3000-01-01" Then
                objNode.ForeColor = mconColor_Stop
            End If
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
    End If
    Call setMenu�Թ�ҩ
    Exit Sub
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub zlRefRecords(Optional lngItem As Long)
    Dim objListitem As ListItem
    
    '---------------------------------------------
    '��дҩƷ�б�
    '---------------------------------------------
    err = 0: On Error GoTo ErrHand
   
    If Val(Me.tvwClass.Tag) <= 2 Then
        gstrSql = "select I.ID,I.����,I.����,I.���㵥λ,T.ҩƷ����," & _
                "        decode(I.�������,1,'����',2,'סԺ',3,'�����סԺ','��ֱ��Ӧ���ڲ���') as �������," & _
                "        decode(T.ҩƷ����,1,'����ҩ',2,'����Ǵ���ҩ',3,'����Ǵ���ҩ',4,'�Ǵ���ҩ',5,'����ҩƷ',' ') as ҩƷ����," & _
                "        to_char(nvl(T.��������,0)) as ��������,decode(T.�Ƿ�Ƥ��,1,'��Ҫ',' ') as �Ƿ�Ƥ��," & _
                "        T.�������,T.��Դ���,T.��ֵ����,T.��ҩ�ݴ�," & _
                "        decode(T.�Ƿ�ԭ��,1,'��',' ') as �Ƿ�ԭ��," & _
                "        decode(T.����ҩ��,1,'��',' ') as ����ҩ��," & _
                "        decode(I.����Ӧ��,1,'��',' ') as ����Ӧ��," & _
                "        decode(T.�Ƿ���ҩ,1,'��',' ') as �Ƿ���ҩ," & _
                "        decode(T.Ʒ��ҽ��,1,'��',' ') as Ʒ��ҽ��," & _
                "        decode(T.�Ƿ�ԭ��ҩ,1,'��',' ') as �Ƿ�ԭ��ҩ," & _
                "        decode(T.�Ƿ�ר��ҩ,1,'��',' ') as �Ƿ�ר��ҩ," & _
                "        decode(T.�Ƿ񵥶�����,1,'��',' ') as �Ƿ񵥶�����," & _
                "        decode(nvl(t.������,0),0,'',1,'������ʹ��',2,'����ʹ��','����ʹ��') as ����ҩ��," & _
                "        decode(T.�Ƿ�����ҩ,1,'��',' ') as �Ƿ�����ҩ," & _
                "        nvl(I.����ʱ��,to_date('3000-01-01','YYYY-MM-DD')) as ����ʱ��,Nvl(I.�����Ա�,0) As �����Ա� " & _
                " from ������ĿĿ¼ I,ҩƷ���� T" & _
                " where I.ID=T.ҩ��ID and "
        If mnuViewShowAll.Checked = False Then
            gstrSql = gstrSql & " I.����ID=[1] "
        Else
            gstrSql = gstrSql & " I.����ID IN " & _
                " (Select ID From ���Ʒ���Ŀ¼ Where ���� In (1,2,3) " & _
                "  Start With ID=[1] Connect By Prior ID=�ϼ�ID)"
        End If
        If Me.mnuViewStoped.Checked = False Then
            gstrSql = gstrSql & " and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))"
        End If
        If mbln�Թ�ҩ = True Then
            gstrSql = gstrSql & " and t.�ٴ��Թ�ҩ=1"
        Else
            gstrSql = gstrSql & " and t.�ٴ��Թ�ҩ is null"
        End If
        gstrSql = gstrSql & " order by I.����"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(Me.tvwClass.SelectedItem.Key, 2)))
        
        With rsTemp
            Me.lvwItems.ListItems.Clear
            Do While Not .EOF
                Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !����)
                If Format(!����ʱ��, "YYYY-MM-DD") = "3000-01-01" Then
                    If Val(Me.tvwClass.Tag) = 2 Then
                        objItem.Icon = "��ҩU": objItem.SmallIcon = "��ҩU"
                    Else
                        objItem.Icon = "��ҩU": objItem.SmallIcon = "��ҩU"
                    End If
                Else
                    If Val(Me.tvwClass.Tag) = 2 Then
                        objItem.Icon = "��ҩS": objItem.SmallIcon = "��ҩS"
                    Else
                        objItem.Icon = "��ҩS": objItem.SmallIcon = "��ҩS"
                    End If
                End If
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_����").Index - 1) = !����
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_������λ").Index - 1) = IIf(IsNull(!���㵥λ), "", !���㵥λ)
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_����").Index - 1) = IIf(IsNull(!ҩƷ����), "", !ҩƷ����)
                'objItem.SubItems(Me.lvwItems.ColumnHeaders("_�������").Index - 1) = !�������
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_��������").Index - 1) = !ҩƷ����
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_��������").Index - 1) = !��������
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_��������").Index - 1) = !�Ƿ�Ƥ��
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_����").Index - 1) = !�������
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_��Դ").Index - 1) = !��Դ���
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_��ֵ").Index - 1) = !��ֵ����
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_�ݴ�").Index - 1) = !��ҩ�ݴ�
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_ԭ��ҩ").Index - 1) = !�Ƿ�ԭ��
                If Val(Me.tvwClass.Tag) = 2 Then
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_��ζʹ��").Index - 1) = !����Ӧ��
                Else
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_����ҩ").Index - 1) = !����ҩ��
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_��ҩ").Index - 1) = !�Ƿ���ҩ
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_ԭ��ҩ").Index - 1) = !�Ƿ�ԭ��ҩ
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_ר��ҩ").Index - 1) = !�Ƿ�ר��ҩ
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_��������").Index - 1) = !�Ƿ񵥶�����
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_����ҩ��").Index - 1) = zlStr.Nvl(!����ҩ��, "")
                End If
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_��ҩƷ�³���ҽ��").Index - 1) = !Ʒ��ҽ��
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_������ҩ").Index - 1) = !�Ƿ�����ҩ
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_�����Ա�").Index - 1) = IIf(!�����Ա� = 1, "����", IIf(!�����Ա� = 2, "Ů��", "���Ա�����"))
                If !ID = lngItem Then
                    objItem.Selected = True
                End If
                If Format(!����ʱ��, "YYYY-MM-DD") <> "3000-01-01" Then
                    objItem.ForeColor = mconColor_Stop
                    For intCount = 1 To Me.lvwItems.ColumnHeaders.Count - 1
                        objItem.ListSubItems(intCount).ForeColor = mconColor_Stop
                    Next
                End If
                .MoveNext
            Loop
        End With
    Else        '����������˽��
        Dim strType As String
        
        strType = Mid(Me.tvwClass.SelectedItem.Key, 2)      '������ѡ�ڵ������ɲ�ѯ����
        
        If Not IsNumeric(strType) Then                      '����������־��ǻ�������
            If strType = "ALL" Then
                strType = " And I.��� in('5','6','7') "
            ElseIf strType = "ALL5" Then
                 strType = " And I.���='5' "
            ElseIf strType = "ALL6" Then
                strType = " And I.���='6' "
            ElseIf strType = "ALL7" Then
                strType = " And I.���='7' "
            Else
                Select Case strType
                    Case "����ҩ"
                        strType = " And nvl(T.������,0)<>0 "
                    Case "ԭ��ҩ"
                        strType = " And nvl(T.�Ƿ�ԭ��ҩ,0)<>0 "
                    Case "ר��ҩ"
                        strType = " And nvl(T.�Ƿ�ר��ҩ,0)<>0 "
                    Case "��������"
                        strType = " And nvl(T.�Ƿ񵥶�����,0)<>0 "
                    Case "Limit1", "Limit2", "Limit3"
                        strType = " And nvl(T.������,0)=" & Mid(strType, 6) & " "
                End Select
            End If
        Else                                                '�����־��Ƿ���ID����
            If Val(Me.tvwClass.Tag) = 4 Then
                strType = " And I.����id=" & strType & " "
            Else
                Select Case Mid(Me.tvwClass.SelectedItem.Key, 1, 1)
                    Case "A"
                        strType = " And nvl(T.������,0)=" & Mid(strType, 1, 1) & " And I.����id=" & Mid(strType, 2) & " "
                    Case "B"
                        strType = " And nvl(T.�Ƿ�ԭ��ҩ,0)<>0 And I.����id=" & Mid(strType, 1) & " "
                    Case "C"
                        strType = " And nvl(T.�Ƿ�ר��ҩ,0)<>0 And I.����id=" & Mid(strType, 1) & " "
                    Case "D"
                        strType = " And nvl(T.�Ƿ񵥶�����,0)<>0 And I.����id=" & Mid(strType, 1) & " "
                End Select
            End If
        End If
        
        Me.lvwItems.Visible = False
        
        gstrSql = "select I.ID,I.����,I.����,I.���㵥λ,T.ҩƷ����," & _
                "        decode(I.�������,1,'����',2,'סԺ',3,'�����סԺ','��ֱ��Ӧ���ڲ���') as �������," & _
                "        decode(T.ҩƷ����,1,'����ҩ',2,'����Ǵ���ҩ',3,'����Ǵ���ҩ',4,'�Ǵ���ҩ',5,'����ҩƷ',' ') as ҩƷ����," & _
                "        to_char(nvl(T.��������,0)) as ��������,decode(T.�Ƿ�Ƥ��,1,'��Ҫ',' ') as �Ƿ�Ƥ��," & _
                "        T.�������,T.��Դ���,T.��ֵ����,T.��ҩ�ݴ�," & _
                "        decode(T.�Ƿ�ԭ��,1,'��',' ') as �Ƿ�ԭ��," & _
                "        decode(T.����ҩ��,1,'��',' ') as ����ҩ��," & _
                "        decode(I.����Ӧ��,1,'��',' ') as ����Ӧ��," & _
                "        decode(T.�Ƿ���ҩ,1,'��',' ') as �Ƿ���ҩ," & _
                "        decode(T.Ʒ��ҽ��,1,'��',' ') as Ʒ��ҽ��," & _
                "        decode(T.�Ƿ�ԭ��ҩ,1,'��',' ') as �Ƿ�ԭ��ҩ," & _
                "        decode(T.�Ƿ�ר��ҩ,1,'��',' ') as �Ƿ�ר��ҩ," & _
                "        decode(T.�Ƿ񵥶�����,1,'��',' ') as �Ƿ񵥶�����," & _
                "        decode(I.���,'5','1','6','2','3') as ����," & _
                "        I.����id," & _
                "        decode(T.�Ƿ�����ҩ,1,'��',' ') as �Ƿ�����ҩ," & _
                "        nvl(I.����ʱ��,to_date('3000-01-01','YYYY-MM-DD')) as ����ʱ��,Nvl(I.�����Ա�,0) As �����Ա�, " & _
                "        decode(nvl(t.������,0),0,'',1,'������ʹ��',2,'����ʹ��','����ʹ��') as ����ҩ�� " & _
                " from ������ĿĿ¼ I,ҩƷ���� T" & _
                " where I.ID=T.ҩ��ID " & IIf(Val(Me.tvwClass.Tag) = 4, IIf(mstrDrugId <> "", "And I.ID IN ( " & mstrDrugId & ") ", ""), "") & strType
        
        If mbln�Թ�ҩ = True Then
            gstrSql = gstrSql & " and t.�ٴ��Թ�ҩ=1"
        Else
            gstrSql = gstrSql & " and t.�ٴ��Թ�ҩ is null"
        End If
    
        If Me.mnuViewStoped.Checked = False Then
            gstrSql = gstrSql & " and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))"
        End If
        gstrSql = gstrSql & " order by I.����"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "zlRefRecords")

        Me.lvwItems.ListItems.Clear
        With rsTemp
            Do While Not .EOF
                Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !����)
                If Format(!����ʱ��, "YYYY-MM-DD") = "3000-01-01" Then
                    If mstrType = "7" Then
                        objItem.Icon = "��ҩU": objItem.SmallIcon = "��ҩU"
                    Else
                        objItem.Icon = "��ҩU": objItem.SmallIcon = "��ҩU"
                    End If
                Else
                    If mstrType = "7" Then
                        objItem.Icon = "��ҩS": objItem.SmallIcon = "��ҩS"
                    Else
                        objItem.Icon = "��ҩS": objItem.SmallIcon = "��ҩS"
                    End If
                End If
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_����").Index - 1) = !����
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_������λ").Index - 1) = IIf(IsNull(!���㵥λ), "", !���㵥λ)
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_����").Index - 1) = IIf(IsNull(!ҩƷ����), "", !ҩƷ����)
                'objItem.SubItems(Me.lvwItems.ColumnHeaders("_�������").Index - 1) = !�������
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_��������").Index - 1) = !ҩƷ����
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_��������").Index - 1) = !��������
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_��������").Index - 1) = !�Ƿ�Ƥ��
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_����").Index - 1) = !�������
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_��Դ").Index - 1) = !��Դ���
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_��ֵ").Index - 1) = !��ֵ����
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_�ݴ�").Index - 1) = !��ҩ�ݴ�
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_ԭ��ҩ").Index - 1) = !�Ƿ�ԭ��
                If mstrType = "7" Then
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_��ζʹ��").Index - 1) = !����Ӧ��
                Else
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_����ҩ").Index - 1) = !����ҩ��
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_��ҩ").Index - 1) = !�Ƿ���ҩ
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_ԭ��ҩ").Index - 1) = !�Ƿ�ԭ��ҩ
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_ר��ҩ").Index - 1) = !�Ƿ�ר��ҩ
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_��������").Index - 1) = !�Ƿ񵥶�����
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_����ҩ��").Index - 1) = zlStr.Nvl(!����ҩ��, "")
                End If
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_��ҩƷ�³���ҽ��").Index - 1) = !Ʒ��ҽ��
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_�����Ա�").Index - 1) = IIf(!�����Ա� = 1, "����", IIf(!�����Ա� = 2, "Ů��", "���Ա�����"))
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_����").Index - 1) = !����
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_����id").Index - 1) = !����id
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_������ҩ").Index - 1) = !�Ƿ�����ҩ
                If !ID = lngItem Then
                    objItem.Selected = True
                End If
                If Format(!����ʱ��, "YYYY-MM-DD") <> "3000-01-01" Then
                    objItem.ForeColor = mconColor_Stop
                    For intCount = 1 To Me.lvwItems.ColumnHeaders.Count - 1
                        objItem.ListSubItems(intCount).ForeColor = mconColor_Stop
                    Next
                End If
                .MoveNext
            Loop
        End With
        Me.lvwItems.Visible = True
    End If
    
    For Each objItem In lvwItems.ListItems
        If objItem.Key = mStrItem Then
            lvwItems.ListItems(objItem.Key).Selected = True
            Exit For
        End If
    Next
    
    If Me.lvwItems.ListItems.Count > 0 Then
        If Me.lvwItems.SelectedItem Is Nothing Then Me.lvwItems.ListItems(1).Selected = True
        Call lvwItems_ItemClick(lvwItems.SelectedItem)
        err = 0: On Error Resume Next
        DoEvents: Me.lvwItems.SelectedItem.EnsureVisible
        Me.stbThis.Panels(2).Text = "�÷��๲��" & Me.lvwItems.ListItems.Count & "��ҩƷ"
    Else
        Me.lvwSpecs.ListItems.Clear
        With Me.hgdPrice
            .Redraw = False
            .Rows = .FixedRows + 1: .RowData(.FixedRows) = 0
            For intCol = 0 To .Cols - 1
                .TextMatrix(.FixedRows, intCol) = ""
            Next
            .Redraw = True
        End With
        For intCount = Me.lblComment.LBound To Me.lblComment.UBound
            Me.lblComment(intCount).Caption = ""
        Next
        
        If fraComment(0).Caption = "" Then
            lvwSpecs.Height = tabContent.Height - 450
        End If
        Me.stbThis.Panels(2).Text = ""
    End If
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub zlPopupEditMenu(bytEditKind As Byte, blnStopUse As Boolean)
    '-------------------------------------------------
    '����:�����༭�˵�
    '���:  bytEditKind:1-Ʒ�ֱ༭ 2-���༭
    '       blnStopUse:�Ƿ����ͣ�ú����ù���
    '-------------------------------------------------
    Dim objItem As ListItem
    Dim StrClass As String
    
    On Error GoTo RESHOW
    Me.mnuEditItemAdd.Tag = Me.mnuEditItemAdd.Visible
    Me.mnuEditItemMod.Tag = Me.mnuEditItemMod.Visible
    Me.mnuEditItemDel.Tag = Me.mnuEditItemDel.Visible
    Me.mnuEditItemTabu.Tag = Me.mnuEditItemTabu.Visible
    Me.mnuEditItemPart.Tag = Me.mnuEditItemPart.Visible
    Me.mnuPriceChargeSet.Tag = Me.mnuPriceChargeSet.Visible
    Me.mnuEditSpt1.Tag = Me.mnuEditSpt1.Visible
    Me.mnuEditSpecAdd.Tag = Me.mnuEditSpecAdd.Visible
    Me.mnuEditSpecMod.Tag = Me.mnuEditSpecMod.Visible
    Me.mnuEditSpecDel.Tag = Me.mnuEditSpecDel.Visible
    Me.mnuEditSpecLimit.Tag = Me.mnuEditSpecLimit.Visible
    Me.mnuEditSendType.Tag = Me.mnuEditSendType.Visible
    Me.mnuEditSpecProtocol.Tag = Me.mnuEditSpecProtocol.Visible
    Me.mnuEditSpecSelf.Tag = Me.mnuEditSpecSelf.Visible
    Me.mnuEditSpt2.Tag = Me.mnuEditSpt2.Visible
    Me.mnuEditStart.Tag = Me.mnuEditStart.Visible
    Me.mnuEditStop.Tag = Me.mnuEditStop.Visible
    Me.mnuEditSptPacker.Tag = Me.mnuEditSptPacker.Visible
    Me.mnuUploadDrugInfo.Tag = Me.mnuUploadDrugInfo.Visible
    
    Me.mnuPriceChargeSet1.Visible = Me.mnuPriceChargeSet.Visible
    Me.mnuEditSpt3.Visible = Me.mnuPriceChargeSet.Visible
    
    Me.mnuPriceChargeSet1.Enabled = Me.mnuPriceChargeSet.Enabled
    
    Select Case bytEditKind
    Case 1  'Ʒ��
        If InStr(1, mstrPrivs, ";��Ӧ����;") = 0 Then
            mnuEditItemBill.Visible = False
        Else
            mnuEditItemBill.Visible = True
        End If
        If InStr(1, mstrPrivs, ";������ɹ�ϵ;") = 0 Then
            mnuEditItemTabu.Visible = False
        Else
            mnuEditItemTabu.Visible = True
        End If
        With lvwItems
            If tvwClass.Tag >= 3 Then
                Set objItem = .SelectedItem
                If objItem.SubItems(.ColumnHeaders("_����").Index - 1) = "1" Then
                    StrClass = "5"
                ElseIf objItem.SubItems(.ColumnHeaders("_����").Index - 1) = "2" Then
                    StrClass = "6"
                ElseIf objItem.SubItems(.ColumnHeaders("_����").Index - 1) = "3" Then
                    StrClass = "7"
                End If
            Else
                If tvwClass.Tag = 0 Then
                    StrClass = "5"
                ElseIf tvwClass.Tag = 1 Then
                    StrClass = "6"
                ElseIf tvwClass.Tag = 2 Then
                    StrClass = "7"
                End If
            End If
            
            Select Case StrClass
            Case "5" '����ҩ
                If InStr(1, mstrPrivs, ";��������ҩƷ��;") = 0 Then
                    mnuEditItemAdd.Visible = False
                    mnuEditItemMod.Visible = False
                    mnuEditItemDel.Visible = False
                    mnuEditStart.Visible = False
                    mnuEditStop.Visible = False
                Else
                    mnuEditItemAdd.Visible = True
                    mnuEditItemMod.Visible = True
                    mnuEditItemDel.Visible = True
                    
                    If InStr(1, mstrPrivs, ";ҩƷ����;") = 0 Then
                        mnuEditStart.Visible = False
                    Else
                        mnuEditStart.Visible = True
                    End If
                    If InStr(1, mstrPrivs, ";ҩƷͣ��;") = 0 Then
                        mnuEditStop.Visible = False
                    Else
                        mnuEditStop.Visible = True
                    End If
                End If
                If InStr(1, mstrPrivs, ";�÷�����;") = 0 Then
                    mnuEditItemUsage.Visible = False
                Else
                    mnuEditItemUsage.Visible = True
                End If
                mnuEditItemTabu.Enabled = True
                mnuEditItemUsage.Enabled = True
                mnuEditItemBill.Enabled = True
            Case "6" '�г�ҩ
                If InStr(1, mstrPrivs, ";�����г�ҩƷ��;") = 0 Then
                    mnuEditItemAdd.Visible = False
                    mnuEditItemMod.Visible = False
                    mnuEditItemDel.Visible = False
                    mnuEditStart.Visible = False
                    mnuEditStop.Visible = False
                Else
                    mnuEditItemAdd.Visible = True
                    mnuEditItemMod.Visible = True
                    mnuEditItemDel.Visible = True
                    
                    If InStr(1, mstrPrivs, ";ҩƷ����;") = 0 Then
                        mnuEditStart.Visible = False
                    Else
                        mnuEditStart.Visible = True
                    End If
                    If InStr(1, mstrPrivs, ";ҩƷͣ��;") = 0 Then
                        mnuEditStop.Visible = False
                    Else
                        mnuEditStop.Visible = True
                    End If
                End If
                If InStr(1, mstrPrivs, ";�÷�����;") = 0 Then
                    mnuEditItemUsage.Visible = False
                Else
                    mnuEditItemUsage.Visible = True
                End If
                mnuEditItemTabu.Enabled = True
                mnuEditItemUsage.Enabled = True
                mnuEditItemBill.Enabled = True
            Case "7"   '�в�ҩ
                If InStr(1, mstrPrivs, ";�����в�ҩƷ��;") = 0 Then
                    mnuEditItemAdd.Visible = False
                    mnuEditItemMod.Visible = False
                    mnuEditItemDel.Visible = False
                    mnuEditStart.Visible = False
                    mnuEditStop.Visible = False
                Else
                    mnuEditItemAdd.Visible = True
                    mnuEditItemMod.Visible = True
                    mnuEditItemDel.Visible = True
                    
                    If InStr(1, mstrPrivs, ";ҩƷ����;") = 0 Then
                        mnuEditStart.Visible = False
                    Else
                        mnuEditStart.Visible = True
                    End If
                    If InStr(1, mstrPrivs, ";ҩƷͣ��;") = 0 Then
                        mnuEditStop.Visible = False
                    Else
                        mnuEditStop.Visible = True
                    End If
                End If
                mnuEditItemUsage.Visible = False '�÷�����
                mnuEditItemTabu.Enabled = True
                mnuEditItemBill.Enabled = True
            End Select
        End With
        
        If tvwClass.Nodes.Count > 0 Then    '���з���ʱ
            mnuEditItemAdd.Enabled = True   '����Ʒ��
        End If
        mnuEditItemDel.Enabled = True   'ɾ��Ʒ��
        If lvwItems.ListItems.Count > 0 Then '��Ʒ��ʱ
            If lvwItems.SelectedItem.Icon Like "*U" = True Then  'δͣ��Ʒ��
                mnuEditItemMod.Enabled = True   '�޸�Ʒ��
                mnuEditStart.Enabled = False     '����
                mnuEditStop.Enabled = True     'ͣ��
            ElseIf lvwItems.SelectedItem.Icon Like "*S" = True Then   '��ͣ��
                mnuEditItemMod.Enabled = False   '�޸�Ʒ��
                mnuEditStart.Enabled = True     '����
                mnuEditStop.Enabled = False     'ͣ��
            End If
        Else
            mnuEditItemAdd.Enabled = True
            mnuEditItemMod.Enabled = False
            mnuEditItemDel.Enabled = False
        End If
        mnuEditSpecAdd.Visible = False
        mnuEditSpecMod.Visible = False
        mnuEditSpecDel.Visible = False
        mnuEditSpt1.Visible = False
        mnuEditSpt7.Visible = False
        mnuEditItemPart.Visible = False
        mnuEditSpecLimit.Visible = False
        mnuEditSpecProtocol.Visible = False
        mnuEditSpecSelf.Visible = False
        mnuEditSpecUnit.Visible = False
        mnuEditManFac.Visible = False
        mnuEditSendType.Visible = False
        mnuPriceChargeSet1.Visible = False   '�ѱ�����
        mnuEditSpt4.Visible = False
        mnuEditSpt3.Visible = mnuEditStop.Visible
    Case 2  '���
        With lvwItems
            If tvwClass.Tag >= 3 Then
                Set objItem = .SelectedItem
                If objItem.SubItems(.ColumnHeaders("_����").Index - 1) = "1" Then
                    StrClass = "5"
                ElseIf objItem.SubItems(.ColumnHeaders("_����").Index - 1) = "2" Then
                    StrClass = "6"
                ElseIf objItem.SubItems(.ColumnHeaders("_����").Index - 1) = "3" Then
                    StrClass = "7"
                End If
            Else
                If tvwClass.Tag = 0 Then
                    StrClass = "5"
                ElseIf tvwClass.Tag = 1 Then
                    StrClass = "6"
                ElseIf tvwClass.Tag = 2 Then
                    StrClass = "7"
                End If
            End If
            
            Select Case StrClass
            Case "5" '����ҩ
                If InStr(1, mstrPrivs, ";��������ҩ���;") = 0 Then
                    mnuEditSpecAdd.Visible = False
                    mnuEditSpecMod.Visible = False
                    mnuEditSpecDel.Visible = False
                    mnuEditStart.Visible = False
                    mnuEditStop.Visible = False
                Else
                    mnuEditSpecAdd.Visible = True
                    mnuEditSpecMod.Visible = True
                    mnuEditSpecDel.Visible = True
                    
                    If InStr(1, mstrPrivs, ";ҩƷ����;") = 0 Then
                        mnuEditStart.Visible = False
                    Else
                        mnuEditStart.Visible = True
                    End If
                    If InStr(1, mstrPrivs, ";ҩƷͣ��;") = 0 Then
                        mnuEditStop.Visible = False
                    Else
                        mnuEditStop.Visible = True
                    End If
                    
                    If InStr(1, mstrPrivs, ";�÷�����;") = 0 Then
                        mnuEditItemUsage.Visible = False
                    Else
                        mnuEditItemUsage.Visible = True
                        mnuEditItemUsage.Enabled = True
                    End If
                End If
            Case "6" '�г�ҩ
                If InStr(1, mstrPrivs, ";�����г�ҩ���;") = 0 Then
                    mnuEditSpecAdd.Visible = False
                    mnuEditSpecMod.Visible = False
                    mnuEditSpecDel.Visible = False
                    mnuEditStart.Visible = False
                    mnuEditStop.Visible = False
                Else
                    mnuEditSpecAdd.Visible = True
                    mnuEditSpecMod.Visible = True
                    mnuEditSpecDel.Visible = True
                    
                    If InStr(1, mstrPrivs, ";ҩƷ����;") = 0 Then
                        mnuEditStart.Visible = False
                    Else
                        mnuEditStart.Visible = True
                    End If
                    If InStr(1, mstrPrivs, ";ҩƷͣ��;") = 0 Then
                        mnuEditStop.Visible = False
                    Else
                        mnuEditStop.Visible = True
                    End If
                    
                    If InStr(1, mstrPrivs, ";�÷�����;") = 0 Then
                        mnuEditItemUsage.Visible = False
                    Else
                        mnuEditItemUsage.Visible = True
                        mnuEditItemUsage.Enabled = True
                    End If
                End If
            Case "7"   '�в�ҩ
                If InStr(1, mstrPrivs, ";�����в�ҩ���;") = 0 Then
                    mnuEditSpecAdd.Visible = False
                    mnuEditSpecMod.Visible = False
                    mnuEditSpecDel.Visible = False
                    mnuEditStart.Visible = False
                    mnuEditStop.Visible = False
                Else
                    mnuEditSpecAdd.Visible = True
                    mnuEditSpecMod.Visible = True
                    mnuEditSpecDel.Visible = True
                    
                    If InStr(1, mstrPrivs, ";ҩƷ����;") = 0 Then
                        mnuEditStart.Visible = False
                    Else
                        mnuEditStart.Visible = True
                    End If
                    If InStr(1, mstrPrivs, ";ҩƷͣ��;") = 0 Then
                        mnuEditStop.Visible = False
                    Else
                        mnuEditStop.Visible = True
                    End If
                End If
            End Select
            If InStr(1, mstrPrivs, ";�ѱ�����;") = 0 Then
                mnuPriceChargeSet1.Visible = False
                mnuEditSpt4.Visible = False
            Else
                mnuPriceChargeSet1.Visible = True
                mnuEditSpt4.Visible = True
            End If
        End With
        
        If lvwItems.ListItems.Count > 0 And lvwItems.SelectedItem.Icon Like "*U" = True Then '��Ʒ��,δͣ��
            mnuEditSpecAdd.Enabled = True
        End If
        mnuEditSpecDel.Enabled = True   'ɾ�����
        If lvwSpecs.ListItems.Count > 0 Then '�й��ʱ
            If lvwSpecs.SelectedItem.Icon Like "*U" = True Then 'δͣ��
                mnuEditSpecMod.Enabled = True   '�޸Ĺ��
                mnuEditStart.Enabled = False     '����
                mnuEditStop.Enabled = True     'ͣ��
            ElseIf lvwSpecs.SelectedItem.Icon Like "*S" = True Then '��ͣ��
                mnuEditSpecMod.Enabled = False   '�޸Ĺ��
                mnuEditStart.Enabled = True     '����
                mnuEditStop.Enabled = False     'ͣ��
            End If
        Else
            mnuEditSpecDel.Enabled = False
            mnuEditSpecMod.Enabled = False
            mnuEditStart.Enabled = False     '����
            mnuEditStop.Enabled = False     'ͣ��
        End If
        mnuEditItemAdd.Visible = False
        mnuEditItemMod.Visible = False
        mnuEditItemDel.Visible = False
        mnuEditItemTabu.Visible = False '�������
'        mnuEditItemUsage.Visible = False '�÷�����
        mnuEditItemBill.Visible = False '��Ӧ����
        mnuEditSpt1.Visible = False
        mnuEditSpt7.Visible = mnuEditSpecAdd.Visible
        mnuEditItemPart.Visible = True
        mnuEditSpecLimit.Visible = True
        mnuEditSpecProtocol.Visible = True
        mnuEditSpecSelf.Visible = True
        mnuEditSpecUnit.Visible = True
        mnuEditManFac.Visible = True
        mnuEditSendType.Visible = True
        mnuEditSpt3.Visible = mnuEditStop.Visible
    End Select
    
    Call setMenu�Թ�ҩ
    Call PopupMenu(Me.mnuEdit, 2)
    
RESHOW:
    Me.mnuEditItemAdd.Visible = Me.mnuEditItemAdd.Tag
    Me.mnuEditItemMod.Visible = Me.mnuEditItemMod.Tag
    Me.mnuEditItemDel.Visible = Me.mnuEditItemDel.Tag
    Me.mnuEditItemTabu.Visible = Me.mnuEditItemTabu.Tag
    Me.mnuEditItemPart.Visible = Me.mnuEditItemPart.Tag
    Me.mnuPriceChargeSet.Visible = Me.mnuPriceChargeSet.Tag
    Me.mnuEditSpt1.Visible = Me.mnuEditSpt1.Tag
    Me.mnuEditSpecAdd.Visible = Me.mnuEditSpecAdd.Tag
    Me.mnuEditSpecMod.Visible = Me.mnuEditSpecMod.Tag
    Me.mnuEditSpecDel.Visible = Me.mnuEditSpecDel.Tag
    Me.mnuEditSpecLimit.Visible = Me.mnuEditSpecLimit.Tag
    Me.mnuEditSendType.Visible = Me.mnuEditSendType.Tag
    Me.mnuEditSpecProtocol.Visible = Me.mnuEditSpecProtocol.Tag
    Me.mnuEditSpecSelf.Visible = Me.mnuEditSpecSelf.Tag
    Me.mnuEditSpt2.Visible = Me.mnuEditSpt2.Tag
    Me.mnuEditStart.Visible = Me.mnuEditStart.Tag
    Me.mnuEditStop.Visible = Me.mnuEditStop.Tag
    Me.mnuEditSptPacker.Visible = Me.mnuEditSptPacker.Tag
    Me.mnuUploadDrugInfo.Visible = Me.mnuUploadDrugInfo.Tag
    Me.mnuEditSpecUnit.Visible = True

    Call setMenu�Թ�ҩ
End Sub

Private Sub setMenu�Թ�ҩ()
    '���ܣ��Թ�ҩ�˵�����
    If mbln�Թ�ҩ = True Then   '�Թ�ҩ�˵�����
        mnuFilePara.Visible = False
        mnuFileSpt2.Visible = False
        mnuClass.Visible = False
        mnuEditItemPart.Visible = False
        mnuEditSpecLimit.Visible = False
        mnuEditSpecProtocol.Visible = False
        mnuEditSpecSelf.Visible = False
        mnuEditSpecUnit.Visible = False
        mnuEditManFac.Visible = False
        mnuEditSendType.Visible = False
        mnuEditSpt5.Visible = False
        mnuEditRate.Visible = False
        mnuEditSpt2.Visible = False
        mnuEditVariBatch.Visible = False
        mnuEditSpecBatch.Visible = False
        mnuEditExcel.Visible = False
        mnuEditSpt4.Visible = False
        mnuPriceChargeSet1.Visible = False
        mnuPriceSpt1.Visible = False
        mnuEditSptPacker.Visible = False
        mnuUploadDrugInfo.Visible = False
        mnuPrice.Visible = False
        mnuEditSpt3.Visible = False
        mnuViewPrices.Visible = False
        tlbThis.Buttons("Limit").Visible = False
        tlbThis.Buttons("Class").Visible = False
        tlbThis.Buttons(10).Visible = False
        mnuEditSpt6.Visible = True
    End If
End Sub
Private Sub zlRptPrint(ByVal bytMode As Byte)
    '-------------------------------------------------
    '����:��¼���ӡ
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '-------------------------------------------------
    Dim objPrint As New zlPrintLvw
    err = 0: On Error Resume Next
    Set objPrint.Body.objData = Me.lvwItems
    Select Case Val(Me.tvwClass.Tag)
    Case 0
        objPrint.Title.Text = "����ҩƷ���嵥"
    Case 1
        objPrint.Title.Text = "�г�ҩƷ���嵥"
    Case 2
        objPrint.Title.Text = "�в�ҩ�嵥"
    End Select
    objPrint.UnderAppItems.Add "���ࣺ" & Me.tvwClass.SelectedItem.Text
    objPrint.BelowAppItems.Add "��ӡʱ�䣺" & Now
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrViewLvw objPrint, bytMode
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
End Sub

Public Sub zlLocateItem(lng����id As Long, lngҩ��id As Long, lngҩƷID As Long)
    Dim lstItem As ListItem, lstSpec As ListItem, tvwNode As Node
    '---------------------------------------------
    '��λ��ָ������ϲο���Ŀ���ڲ���ʱʹ��
    '---------------------------------------------
    On Error GoTo ErrHand
    Set tvwNode = tvwClass.SelectedItem
    Set lstItem = lvwItems.SelectedItem
    Set lstSpec = lvwSpecs.SelectedItem
    
'    If lstItem Is Nothing Then
'        Exit Sub
'    End If
    'ѡ�����
    Set Me.tvwClass.SelectedItem = Me.tvwClass.Nodes("_" & lng����id)
    Me.tvwClass.Nodes("_" & lng����id).Selected = True
    Me.tvwClass.SelectedItem.EnsureVisible
    Call zlRefRecords
    If lvwItems.ListItems.Count <> 0 Then '�����������û��Ʒ��ʱ��û�б�Ҫ�ڶ�λ
        'ѡ��Ʒ��
        Set Me.lvwItems.SelectedItem = Me.lvwItems.ListItems("_" & lngҩ��id)
        Me.lvwItems.SelectedItem.EnsureVisible
        Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
        'ѡ��ҩƷ
        If lngҩƷID <> 0 Then
            Set Me.lvwSpecs.SelectedItem = Me.lvwSpecs.ListItems("_" & lngҩƷID)
            If err <> 0 Then
                Set Me.lvwSpecs.SelectedItem = Me.lvwSpecs.ListItems(1)
            End If
            Me.lvwSpecs.SelectedItem.EnsureVisible
        End If
    End If
    Exit Sub
ErrHand:
    Set tvwClass.SelectedItem = tvwNode
    Call zlRefRecords
    Set Me.lvwItems.SelectedItem = Me.lvwItems.ListItems(lstItem.Key)
    Me.lvwItems.SelectedItem.EnsureVisible
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
    Set Me.lvwSpecs.SelectedItem = Me.lvwSpecs.ListItems(lstSpec.Key)
    Me.lvwSpecs.SelectedItem.EnsureVisible
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub


Public Sub ZlRefBut(ByVal intType As Integer)
    If intType = 3 Then
        cmdKind_Click (3)
    End If
End Sub

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim strTemp As String
    Dim strTag As String
    
    On Error GoTo ErrHandle
    
    If KeyAscii = vbKeyReturn Then
        zlControl.TxtSelAll txtFind
        If txtFind.Text = "" Then Exit Sub
        If mstrFindValue <> txtFind.Text And txtFind.Text <> "" Then
            mstrFindValue = txtFind.Text
            Set mrsFind = Nothing
            
            strTemp = " And (I.����ʱ�� Is NULL Or to_Char(I.����ʱ��,'yyyy-MM-dd')='3000-01-01')"
            
            If mbln�Թ�ҩ = False Then
                gstrSql = "SELECT DISTINCT I.����ID,I.ID AS ҩ��ID,0 AS ҩƷID" & _
                    " FROM ������ĿĿ¼ I,������Ŀ���� N" & _
                    " WHERE I.ID=N.������ĿID " & _
                    " AND I.���=[1] " & _
                    " AND (I.���� LIKE [2] " & _
                    "     OR N.���� LIKE [2] " & _
                    "     OR N.���� LIKE [2])"
            Else
                gstrSql = "Select Distinct i.����id, i.Id As ҩ��id, 0 As ҩƷid" & vbNewLine & _
                    "From ������ĿĿ¼ I, ������Ŀ���� N, ҩƷ���� A" & vbNewLine & _
                    "Where i.Id = n.������Ŀid And i.Id = a.ҩ��id And a.�ٴ��Թ�ҩ = 1 And i.��� = [1] And" & vbNewLine & _
                    "      (i.���� Like [2] Or n.���� Like [2] Or n.���� Like [2])"
            End If
            
            If tvwClass.Tag = "0" Then
                strTag = "5"
            ElseIf tvwClass.Tag = "1" Then
                strTag = "6"
            ElseIf tvwClass.Tag = "2" Then
                strTag = "7"
            End If
            If mnuViewStoped.Checked = False Then
                gstrSql = gstrSql & strTemp
            End If
            Set mrsFind = zlDatabase.OpenSQLRecord(gstrSql, "ҩƷ��ѯ", strTag, gstrMatch & UCase(txtFind.Text) & "%")
            If mrsFind.RecordCount > 0 Then
                Call zlLocateItem(mrsFind!����id, mrsFind!ҩ��ID, mrsFind!ҩƷid)
            End If
        Else
            If Not mrsFind.EOF Then
                mrsFind.MoveNext
                If Not mrsFind.EOF Then
                    Call zlLocateItem(mrsFind!����id, mrsFind!ҩ��ID, mrsFind!ҩƷid)
                Else
                    MsgBox "�Ѳ�ѯ�����һ����¼��", vbInformation, gstrSysName
                    mrsFind.MoveFirst
                    Call zlLocateItem(mrsFind!����id, mrsFind!ҩ��ID, mrsFind!ҩƷid)
                End If
            ElseIf mrsFind.RecordCount <> 0 And mrsFind.EOF Then
                mrsFind.MoveFirst
                Call zlLocateItem(mrsFind!����id, mrsFind!ҩ��ID, mrsFind!ҩƷid)
            End If
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub




